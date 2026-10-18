// db.js - IndexedDB wrapper for Teams Chat Backup

const DB_NAME = 'TeamsBackupDB';
const DB_VERSION = 2; // Upgraded for trash feature

// ==================== IndexedDB Error Handling (P1-3) ====================

// Classify IndexedDB errors for better handling
function classifyIndexedDBError(error) {
    if (!error) return { type: 'unknown', recoverable: true };

    const message = error.message || error.name || '';

    if (message.includes('QuotaExceeded') || message.includes('quota')) {
        return { type: 'quota', recoverable: false, userMessage: 'Storage quota exceeded. Please delete some old backups.' };
    }
    if (message.includes('UnknownError') || message.includes('disk')) {
        return { type: 'disk', recoverable: false, userMessage: 'Disk space error. Please free up space and try again.' };
    }
    if (message.includes('ConstraintError') || message.includes('duplicate')) {
        return { type: 'constraint', recoverable: true, userMessage: 'Data conflict occurred.' };
    }
    if (message.includes('TransactionInactive') || message.includes('abort')) {
        return { type: 'transaction', recoverable: true, userMessage: 'Transaction was interrupted.' };
    }
    if (message.includes('InvalidStateError')) {
        return { type: 'state', recoverable: false, userMessage: 'Database state error. Please reload the extension.' };
    }
    if (message.includes('VersionError')) {
        return { type: 'version', recoverable: false, userMessage: 'Database version mismatch. Please clear browser data.' };
    }

    return { type: 'unknown', recoverable: true, userMessage: `Database error: ${message}` };
}

// Wrap IndexedDB operations with error handling
function wrapIDBOperation(promise, operationName) {
    return promise.catch(error => {
        const classification = classifyIndexedDBError(error);
        console.error(`IndexedDB ${operationName} failed:`, error);
        console.error('Error classification:', classification);

        const enhancedError = new Error(classification.userMessage);
        enhancedError.originalError = error;
        enhancedError.type = classification.type;
        enhancedError.recoverable = classification.recoverable;
        enhancedError.operation = operationName;

        throw enhancedError;
    });
}

// Check available storage before save operations
async function checkStorageAvailability() {
    if ('storage' in navigator && 'estimate' in navigator.storage) {
        try {
            const estimate = await navigator.storage.estimate();
            const usage = estimate.usage || 0;
            const quota = estimate.quota || Number.MAX_SAFE_INTEGER;
            const usagePercent = (usage / quota) * 100;

            console.log(`Storage usage: ${(usage / 1024 / 1024).toFixed(2)}MB / ${(quota / 1024 / 1024).toFixed(2)}MB (${usagePercent.toFixed(1)}%)`);

            if (usagePercent > 95) {
                console.warn('Storage usage over 95%, may encounter quota errors');
                return { available: false, usagePercent, message: 'Storage almost full' };
            }
            if (usagePercent > 80) {
                console.warn('Storage usage over 80%, approaching quota limit');
                return { available: true, usagePercent, message: 'Storage moderately used' };
            }
            return { available: true, usagePercent, message: 'Storage available' };
        } catch (error) {
            console.warn('Could not estimate storage availability:', error);
            return { available: true, usagePercent: null, message: 'Unknown storage status' };
        }
    }
    return { available: true, usagePercent: null, message: 'Storage API not available' };
}

// Recycle bin retention period (60 days in milliseconds)
const TRASH_RETENTION_DAYS = 60;
const TRASH_RETENTION_MS = TRASH_RETENTION_DAYS * 24 * 60 * 60 * 1000;
// Free plan: one chat per browser profile. The allowance stays local, including
// after a backup is deleted. Chat identities are never written to Chrome sync.
const FREE_CHAT_LIMIT = 1;
const FREE_USAGE_STORAGE_KEY = 'freeBackupUsage';
const FREE_USAGE_SEEDED_KEY = 'freeBackupUsageSeeded';
const FREE_USAGE_MAX_STORED_CHATS = 25;

let dbInstance = null;

function storageAreaGet(area, keys, { strict = false } = {}) {
    const storageArea = typeof chrome !== 'undefined' ? chrome.storage?.[area] : null;
    if (!storageArea) return Promise.resolve({});

    return new Promise((resolve, reject) => {
        storageArea.get(keys, (result) => {
            if (chrome.runtime && chrome.runtime.lastError) {
                if (strict) {
                    reject(new Error(chrome.runtime.lastError.message || 'Could not read storage'));
                    return;
                }
                resolve({});
                return;
            }
            resolve(result || {});
        });
    });
}

function storageAreaRemove(area, keys) {
    return new Promise((resolve, reject) => {
        chrome.storage[area].remove(keys, () => {
            if (chrome.runtime.lastError) reject(new Error(chrome.runtime.lastError.message));
            else resolve();
        });
    });
}

function storageAreaSet(area, items) {
    const storageArea = typeof chrome !== 'undefined' ? chrome.storage?.[area] : null;
    if (!storageArea) return Promise.resolve(false);

    return new Promise((resolve) => {
        storageArea.set(items, () => {
            resolve(!(chrome.runtime && chrome.runtime.lastError));
        });
    });
}

async function getStoredProStatus() {
    const result = await storageAreaGet('local', ['isPro', 'purchaseVerified']);
    return result?.isPro === true || result?.purchaseVerified === true;
}

function openDB() {
    return new Promise((resolve, reject) => {
        if (dbInstance) {
            resolve(dbInstance);
            return;
        }

        const request = indexedDB.open(DB_NAME, DB_VERSION);

        request.onerror = () => {
            console.error('Failed to open IndexedDB:', request.error);
            reject(request.error);
        };

        request.onsuccess = () => {
            dbInstance = request.result;
            resolve(dbInstance);
        };

        request.onupgradeneeded = (event) => {
            const db = event.target.result;
            const oldVersion = event.oldVersion;

            // Create chats object store
            if (!db.objectStoreNames.contains('chats')) {
                const chatsStore = db.createObjectStore('chats', { keyPath: 'id' });
                chatsStore.createIndex('createdAt', 'createdAt', { unique: false });
                chatsStore.createIndex('name', 'name', { unique: false });
                chatsStore.createIndex('trashedAt', 'trashedAt', { unique: false });
            } else if (oldVersion < 2) {
                // Upgrade from version 1: add trashedAt index
                const transaction = event.target.transaction;
                const chatsStore = transaction.objectStore('chats');
                if (!chatsStore.indexNames.contains('trashedAt')) {
                    chatsStore.createIndex('trashedAt', 'trashedAt', { unique: false });
                }
            }

            // Create messages object store
            if (!db.objectStoreNames.contains('messages')) {
                const messagesStore = db.createObjectStore('messages', { keyPath: 'id' });
                messagesStore.createIndex('chatId', 'chatId', { unique: false });
                messagesStore.createIndex('order', 'order', { unique: false });
            }
        };
    });
}

function generateId() {
    return 'chat_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
}

function generateMessageId() {
    return 'msg_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
}

// Generate a consistent color from a string (for avatars)
function stringToColor(str) {
    let hash = 0;
    for (let i = 0; i < str.length; i++) {
        hash = str.charCodeAt(i) + ((hash << 5) - hash);
    }
    const colors = [
        '#6264A7', // Teams purple
        '#464775', // Teams dark purple
        '#0078D4', // Microsoft blue
        '#107C10', // Green
        '#D83B01', // Orange
        '#8764B8', // Light purple
        '#038387', // Teal
        '#CA5010', // Burnt orange
        '#4F6BED', // Blue
        '#C239B3'  // Pink
    ];
    return colors[Math.abs(hash) % colors.length];
}

// Get initials from author name
function getInitials(name) {
    if (!name) return '?';
    const parts = name.trim().split(/\s+/);
    if (parts.length >= 2) {
        return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
    }
    return name.substring(0, 2).toUpperCase();
}

const CHAT_STATUS_SUFFIX_PATTERN = /\s*(?:[-|]\s*)?(offline|online|away|busy|do not disturb|be right back|appear offline|out of office)\s*$/i;

function normalizeChatName(name) {
    return (name || '')
        .toString()
        .replace(/\u00a0/g, ' ')
        .trim()
        .replace(/\s*\(\d+\)\s*$/, '') // Remove participant counts like "(5)"
        .replace(CHAT_STATUS_SUFFIX_PATTERN, '')
        .replace(/\s+/g, ' ')
        .toLowerCase();
}

function getNormalizedNameTokens(normalizedName) {
    return normalizedName
        .split(/[\s,;|/()[\]{}'"`~!@#$%^&*+=<>?:.-]+/)
        .map(token => token.trim())
        .filter(token => token.length >= 2);
}

// Chats (not channels) must match on whole words and cover most of the longer
// name, so "John Smith" never matches "John Smith, Jane Doe" or "John Smithson",
// and "Team A" never matches "Team B". A group chat that gained a member
// ("Alice, Bob, Carol" -> "Alice, Bob, Carol, Dave") or a reordered name
// ("Smith, John" -> "John Smith") still matches.
const MIN_CHAT_NAME_COVERAGE = 0.6;

function scoreStrictChatNameMatch(existing, incoming) {
    const toWords = (name) => name
        .split(/[\s,;|/()[\]{}'"`~!@#$%^&*+=<>?:.-]+/)
        .filter(Boolean);
    const existingWords = toWords(existing);
    const incomingWords = toWords(incoming);
    if (existingWords.length === 0 || incomingWords.length === 0) return 0;

    const existingJoined = existingWords.join(' ');
    const incomingJoined = incomingWords.join(' ');
    if (existingJoined === incomingJoined) return 950;

    const [shorter, longer] = existingJoined.length <= incomingJoined.length
        ? [existingJoined, incomingJoined]
        : [incomingJoined, existingJoined];
    if (shorter.length / longer.length < MIN_CHAT_NAME_COVERAGE) return 0;

    if (` ${longer} `.includes(` ${shorter} `)) {
        return 800 - (longer.length - shorter.length);
    }

    const [smallerWords, largerWords] = existingWords.length <= incomingWords.length
        ? [existingWords, incomingWords]
        : [incomingWords, existingWords];
    const largerSet = new Set(largerWords);
    const matches = smallerWords.filter(word => largerSet.has(word)).length;
    if (matches === smallerWords.length) {
        return 650 + matches * 12 - Math.abs(existingWords.length - incomingWords.length) * 3;
    }

    return 0;
}

function scoreChatNameMatch(existingName, incomingName, { strict = false } = {}) {
    const existing = normalizeChatName(existingName);
    const incoming = normalizeChatName(incomingName);

    if (!existing || !incoming) return 0;
    if (existing === incoming) return 1000;
    if (strict) return scoreStrictChatNameMatch(existing, incoming);

    if (existing.includes(incoming) || incoming.includes(existing)) {
        return 800 - Math.abs(existing.length - incoming.length);
    }

    const existingTokens = getNormalizedNameTokens(existing);
    const incomingTokens = getNormalizedNameTokens(incoming);
    if (existingTokens.length === 0 || incomingTokens.length === 0) return 0;

    const smaller = existingTokens.length <= incomingTokens.length ? existingTokens : incomingTokens;
    const largerSet = new Set(existingTokens.length <= incomingTokens.length ? incomingTokens : existingTokens);
    const matches = smaller.filter(token => largerSet.has(token)).length;
    const ratio = matches / smaller.length;

    if (ratio >= 1) {
        return 650 + matches * 12 - Math.abs(existingTokens.length - incomingTokens.length) * 3;
    }

    if (ratio >= 0.75 && matches >= 2) {
        return 350 + matches * 10;
    }

    return 0;
}

function isChatNameMatch(existingName, incomingName) {
    return scoreChatNameMatch(existingName, incomingName) > 0;
}

function getChatRecency(chat) {
    const updated = Date.parse(chat?.updatedAt || '');
    if (Number.isFinite(updated)) return updated;
    const created = Date.parse(chat?.createdAt || '');
    if (Number.isFinite(created)) return created;
    return 0;
}

function canonicalizeSourceKind(sourceKind, isChannel) {
    const normalizedKind = (sourceKind || '').toString().trim().toLowerCase();

    if (normalizedKind === 'channel_chat' || normalizedKind === 'channel_posts' || normalizedKind === 'channel') {
        return normalizedKind;
    }
    if (isChannel === true) {
        return 'channel';
    }
    if (isChannel === false) {
        return 'chat';
    }
    return normalizedKind || '';
}

function buildChatLookupTarget(chatInput) {
    if (chatInput && typeof chatInput === 'object') {
        const isChannel = typeof chatInput.isChannel === 'boolean' ? chatInput.isChannel : null;
        return {
            name: chatInput.name || '',
            rawName: chatInput.rawName || chatInput.name || '',
            teamName: chatInput.teamName || '',
            isChannel,
            sourceKind: canonicalizeSourceKind(chatInput.sourceKind, isChannel === true)
        };
    }

    return {
        name: chatInput || '',
        rawName: chatInput || '',
        teamName: '',
        isChannel: null,
        sourceKind: ''
    };
}

function scoreSourceKindMatch(existingKind, incomingKind, isChannelContext) {
    if (!isChannelContext) return 0;
    if (!existingKind || !incomingKind) return 25;
    if (existingKind === incomingKind) return 140;
    if (existingKind === 'channel' || incomingKind === 'channel') return 85;
    return -1;
}

function scoreChatIdentityMatch(existingChat, incomingChat) {
    const incomingIsChannelKnown = typeof incomingChat?.isChannel === 'boolean';
    const existingIsChannel = existingChat?.isChannel === true;

    if (incomingIsChannelKnown && existingIsChannel !== incomingChat.isChannel) {
        return 0;
    }

    const existingNames = Array.from(new Set([
        existingChat?.rawName,
        existingChat?.name
    ].map(normalizeChatName).filter(Boolean)));
    const incomingNames = Array.from(new Set([
        incomingChat?.rawName,
        incomingChat?.name
    ].map(normalizeChatName).filter(Boolean)));

    if (existingNames.length === 0 || incomingNames.length === 0) {
        return 0;
    }

    const isChannelContext = existingIsChannel ||
        incomingChat?.isChannel === true ||
        canonicalizeSourceKind(existingChat?.sourceKind, existingIsChannel).startsWith('channel') ||
        canonicalizeSourceKind(incomingChat?.sourceKind, incomingChat?.isChannel === true).startsWith('channel');
    // Channels keep loose name matching; the team-name check below separates them.
    const nameMatchOptions = { strict: !isChannelContext };

    let bestNameScore = 0;
    existingNames.forEach((existingName) => {
        incomingNames.forEach((incomingName) => {
            const score = scoreChatNameMatch(existingName, incomingName, nameMatchOptions);
            if (score > bestNameScore) {
                bestNameScore = score;
            }
        });
    });

    if (bestNameScore <= 0) {
        return 0;
    }

    let score = bestNameScore;
    const existingDisplayName = normalizeChatName(existingChat?.name || '');
    const incomingDisplayName = normalizeChatName(incomingChat?.name || '');
    const existingRawName = normalizeChatName(existingChat?.rawName || '');
    const incomingRawName = normalizeChatName(incomingChat?.rawName || '');

    if (existingDisplayName && incomingDisplayName && existingDisplayName === incomingDisplayName) {
        score += 180;
    }
    if (existingRawName && incomingRawName && existingRawName === incomingRawName) {
        score += 220;
    }

    if (isChannelContext) {
        const existingTeam = normalizeChatName(existingChat?.teamName || '');
        const incomingTeam = normalizeChatName(incomingChat?.teamName || '');

        if (existingTeam && incomingTeam) {
            if (existingTeam !== incomingTeam) {
                return 0;
            }
            score += 240;
        } else if (existingTeam || incomingTeam) {
            score += 70;
        }

        const sourceKindScore = scoreSourceKindMatch(
            canonicalizeSourceKind(existingChat?.sourceKind, existingIsChannel),
            canonicalizeSourceKind(incomingChat?.sourceKind, incomingChat?.isChannel === true),
            true
        );

        if (sourceKindScore < 0) {
            return 0;
        }

        score += sourceKindScore;
    }

    return score;
}

async function getAllChatsRaw(db) {
    return new Promise((resolve, reject) => {
        const transaction = db.transaction(['chats'], 'readonly');
        const store = transaction.objectStore('chats');
        const request = store.getAll();

        request.onerror = () => reject(request.error);
        request.onsuccess = () => resolve(request.result || []);
    });
}

async function getActiveChatsRaw(db) {
    const chats = await getAllChatsRaw(db);
    return chats.filter(chat => !chat.trashedAt);
}

// ==================== Free plan usage ====================

function toFreeUsageIdentity(chatInput) {
    const target = buildChatLookupTarget(chatInput);
    const clip = (value) => String(value || '').slice(0, 120);
    return {
        name: clip(target.name),
        rawName: clip(target.rawName),
        teamName: clip(target.teamName),
        isChannel: target.isChannel === true,
        sourceKind: target.sourceKind || ''
    };
}

function normalizeFreeUsage(value) {
    const chats = Array.isArray(value?.chats)
        ? value.chats.filter(chat => chat && (chat.name || chat.rawName))
        : [];
    const count = Number.isFinite(value?.count) ? Math.max(value.count, chats.length) : chats.length;
    return { count, chats };
}

function freeUsageIncludes(usage, chatInput) {
    if (!chatInput) return false;
    const target = buildChatLookupTarget(chatInput);
    if (!normalizeChatName(target.name) && !normalizeChatName(target.rawName)) return false;
    return usage.chats.some(identity => scoreChatIdentityMatch(identity, target) > 0);
}

function addChatToFreeUsage(usage, chatInput) {
    if (freeUsageIncludes(usage, chatInput)) return false;
    const identity = toFreeUsageIdentity(chatInput);
    if (!normalizeChatName(identity.name) && !normalizeChatName(identity.rawName)) return false;
    usage.chats = [...usage.chats, identity].slice(-FREE_USAGE_MAX_STORED_CHATS);
    usage.count += 1;
    return true;
}

// Preserve the v6.6 allowance locally before removing its chat metadata from
// Chrome sync. Retry on a later call if reading, writing, or removal fails.
let freeUsageMigration = null;
function migrateFreeBackupUsageToLocal() {
    if (freeUsageMigration) return freeUsageMigration;
    freeUsageMigration = migrateLegacyFreeUsage().finally(() => { freeUsageMigration = null; });
    return freeUsageMigration;
}

async function migrateLegacyFreeUsage() {
    const [syncData, localData] = await Promise.all([
        storageAreaGet('sync', [FREE_USAGE_STORAGE_KEY], { strict: true }),
        storageAreaGet('local', [FREE_USAGE_STORAGE_KEY], { strict: true })
    ]);
    if (!syncData[FREE_USAGE_STORAGE_KEY]) return;
    const usage = normalizeFreeUsage(localData[FREE_USAGE_STORAGE_KEY]);
    const legacy = normalizeFreeUsage(syncData[FREE_USAGE_STORAGE_KEY]);
    legacy.chats.forEach(chat => addChatToFreeUsage(usage, chat));
    usage.count = Math.max(usage.count, legacy.count);
    if (!await storageAreaSet('local', { [FREE_USAGE_STORAGE_KEY]: usage })) {
        throw new Error('Could not save the free-plan allowance locally');
    }
    await storageAreaRemove('sync', [FREE_USAGE_STORAGE_KEY]);
}

async function getFreeBackupUsage() {
    await migrateFreeBackupUsageToLocal();
    const localData = await storageAreaGet('local', [FREE_USAGE_STORAGE_KEY, FREE_USAGE_SEEDED_KEY], { strict: true });
    const usage = normalizeFreeUsage(localData[FREE_USAGE_STORAGE_KEY]);

    if (!localData[FREE_USAGE_SEEDED_KEY] || !localData[FREE_USAGE_STORAGE_KEY]) {
        const db = await openDB();
        const existingChats = await getAllChatsRaw(db);
        existingChats
            .sort((a, b) => getChatRecency(a) - getChatRecency(b))
            .forEach(chat => addChatToFreeUsage(usage, chat));
        if (!await storageAreaSet('local', { [FREE_USAGE_STORAGE_KEY]: usage, [FREE_USAGE_SEEDED_KEY]: true })) {
            throw new Error('Could not save the free-plan allowance locally');
        }
    }

    return usage;
}

async function recordFreeBackupUsage(chatInput) {
    const usage = await getFreeBackupUsage();
    if (addChatToFreeUsage(usage, chatInput)) {
        if (!await storageAreaSet('local', { [FREE_USAGE_STORAGE_KEY]: usage })) {
            throw new Error('Could not save the free-plan allowance locally');
        }
    }
    return usage;
}

// Whether a free user may create a backup record for this chat. chatInput may be
// null when the open chat could not be identified; it then counts as a new chat.
async function getFreeBackupAllowance(chatInput) {
    const [usage, activeCount] = await Promise.all([getFreeBackupUsage(), getChatCount()]);
    const usedChatNames = usage.chats.map(chat => chat.name || chat.rawName).filter(Boolean);
    const base = { limit: FREE_CHAT_LIMIT, activeCount, usedCount: usage.count, usedChatNames };

    if (activeCount >= FREE_CHAT_LIMIT) {
        return { ...base, allowed: false, reason: 'active_limit' };
    }
    if (usage.count >= FREE_CHAT_LIMIT && !freeUsageIncludes(usage, chatInput)) {
        return { ...base, allowed: false, reason: 'lifetime_limit' };
    }
    return { ...base, allowed: true, reason: null };
}

function describeFreeLimit(allowance) {
    const usedName = allowance?.usedChatNames?.[allowance.usedChatNames.length - 1];
    const plan = `The free plan includes ${FREE_CHAT_LIMIT} chat backup${FREE_CHAT_LIMIT === 1 ? '' : 's'}`;
    if (allowance?.reason === 'active_limit' && allowance.activeCount > FREE_CHAT_LIMIT) {
        return `${plan}. You already have ${allowance.activeCount} backups saved. Upgrade to Pro to back up more chats.`;
    }
    return usedName
        ? `${plan}, already used for "${usedName}". Upgrade to Pro to back up more chats.`
        : `${plan}, and it has been used. Upgrade to Pro to back up more chats.`;
}

// Find existing chat by identity (excludes trashed chats).
// Accepts either a plain name string or a richer chat metadata object.
async function findChatByName(chatInput) {
    const db = await openDB();
    const lookupTarget = buildChatLookupTarget(chatInput);
    const hasLookupName = normalizeChatName(lookupTarget.name) || normalizeChatName(lookupTarget.rawName);
    if (!hasLookupName) return null;

    const activeChats = await getActiveChatsRaw(db);

    const candidates = activeChats
        .map(chat => ({
            chat,
            matchScore: scoreChatIdentityMatch(chat, lookupTarget)
        }))
        .filter(candidate => candidate.matchScore > 0)
        .sort((a, b) => {
            if (b.matchScore !== a.matchScore) return b.matchScore - a.matchScore;
            return getChatRecency(b.chat) - getChatRecency(a.chat);
        });

    if (candidates.length === 0) return null;

    const matched = candidates[0].chat;
    const normalizedIncomingName = normalizeChatName(lookupTarget.rawName || lookupTarget.name);
    const normalizedMatchedName = normalizeChatName(matched.rawName || matched.name);
    if (normalizedMatchedName !== normalizedIncomingName) {
        console.log(`findChatByName: identity matched "${lookupTarget.rawName || lookupTarget.name}" -> "${matched.name}"`);
    }
    return matched;
}

function parseClockTimeOnDate(timeText, baseDate) {
    const match = (timeText || '')
        .trim()
        .match(/^(\d{1,2})(?::(\d{2}))?(?::(\d{2}))?\s*(am|pm)?$/i);

    if (!match) return NaN;

    let hours = Number(match[1]);
    const minutes = Number(match[2] || 0);
    const seconds = Number(match[3] || 0);
    const meridiem = (match[4] || '').toLowerCase();

    if (minutes > 59 || seconds > 59 || hours > 24) return NaN;

    if (meridiem) {
        if (hours === 12) {
            hours = meridiem === 'am' ? 0 : 12;
        } else if (meridiem === 'pm') {
            hours += 12;
        }
    }

    if (hours > 23) return NaN;

    const resolved = new Date(baseDate);
    resolved.setHours(hours, minutes, seconds, 0);
    return resolved.getTime();
}

// Date.parse fills a missing year with 2001 ("9/22 4:57 PM", "September 28").
// Teams leaves the year out for recent dates, so use the latest such date instead.
function inferMissingYear(parsed, rawTimestamp) {
    if (new Date(parsed).getFullYear() !== 2001 || rawTimestamp.includes('2001')) return parsed;

    const now = new Date();
    const inferred = new Date(parsed);
    inferred.setFullYear(now.getFullYear());
    if (inferred.getTime() > now.getTime() + 24 * 60 * 60 * 1000) {
        inferred.setFullYear(now.getFullYear() - 1);
    }
    return inferred.getTime();
}

function parseFlexibleTimestamp(rawTimestamp) {
    const normalizedRaw = (rawTimestamp || '').replace(/\u00a0/g, ' ').trim();
    if (!normalizedRaw) return NaN;

    const direct = Date.parse(normalizedRaw);
    if (Number.isFinite(direct)) return inferMissingYear(direct, normalizedRaw);

    const now = new Date();
    const lower = normalizedRaw.toLowerCase();

    const todayMatch = lower.match(/^today(?:\s+at)?\s+(.+)$/i);
    if (todayMatch) {
        return parseClockTimeOnDate(todayMatch[1], now);
    }

    const yesterdayMatch = lower.match(/^yesterday(?:\s+at)?\s+(.+)$/i);
    if (yesterdayMatch) {
        const yesterday = new Date(now);
        yesterday.setDate(now.getDate() - 1);
        return parseClockTimeOnDate(yesterdayMatch[1], yesterday);
    }

    // Handles formats where removing "at" makes Date.parse work.
    const normalizedWithoutAt = normalizedRaw.replace(/\bat\b/ig, ' ');
    const withoutAt = Date.parse(normalizedWithoutAt);
    if (Number.isFinite(withoutAt)) return inferMissingYear(withoutAt, normalizedRaw);

    return NaN;
}

function normalizeTimestampForKey(timestamp) {
    const rawTimestamp = (timestamp || '').trim();
    if (!rawTimestamp) return '';

    const parsedTimestamp = parseFlexibleTimestamp(rawTimestamp);
    return Number.isFinite(parsedTimestamp)
        ? new Date(parsedTimestamp).toISOString()
        : rawTimestamp.toLowerCase().replace(/\s+/g, ' ');
}

function buildMessageMediaSignature(message) {
    const imageSignature = (message?.images || [])
        .map(img => [
            normalizeAssetIdentityPart(img?.alt),
            Number.isFinite(Number(img?.width)) ? String(Number(img.width)) : '',
            Number.isFinite(Number(img?.height)) ? String(Number(img.height)) : ''
        ].filter(Boolean).join(':'))
        .filter(Boolean)
        .sort();

    const attachmentSignature = (message?.attachments || [])
        .map(att => [
            normalizeAssetIdentityPart(att?.filename),
            normalizeAssetIdentityPart(att?.sharePointUrl),
            normalizeAssetIdentityPart(att?.arrivalTime)
        ].filter(Boolean).join(':'))
        .filter(Boolean)
        .sort();

    return [...imageSignature, ...attachmentSignature].join('|').substring(0, 180);
}

// Create a unique key for a message to detect duplicates
function createMessageKey(msg) {
    const type = (msg.type || 'message').trim().toLowerCase();
    const author = (msg.author || '').trim().toLowerCase();
    const timestamp = normalizeTimestampForKey(msg.timestamp);
    const content = (msg.content || '').trim().replace(/\s+/g, ' ').substring(0, 120).toLowerCase();
    const mediaSignature = buildMessageMediaSignature(msg);
    return `${type}|${author}|${timestamp}|${content}|${mediaSignature}`;
}

function normalizeContentForLooseMatch(content) {
    return (content || '').trim().replace(/\s+/g, ' ').toLowerCase();
}

function isRelaxedMeetingControlContent(content) {
    const normalized = normalizeContentForLooseMatch(content);
    if (!normalized) return false;

    return normalized.includes('meeting ended:') ||
        normalized.endsWith('was invited to the meeting.') ||
        normalized.endsWith('was added to the meeting.') ||
        normalized.endsWith('was removed from the meeting.');
}

function isRelaxedMeetingControlMessage(msg) {
    if (!msg) return false;

    const type = normalizeAssetIdentityPart(msg.type);
    if (type === 'system') {
        return isRelaxedMeetingControlContent(msg.content);
    }

    const author = normalizeAssetIdentityPart(msg.author);
    return (!author || author === 'unknown' || author === 'unknown author') &&
        isRelaxedMeetingControlContent(msg.content);
}

function timestampsLooselyMatch(a, b) {
    const normalizedA = normalizeTimestampForKey(a);
    const normalizedB = normalizeTimestampForKey(b);

    if (!normalizedA || !normalizedB) {
        return true;
    }

    return normalizedA === normalizedB;
}

function mergeMessageMetadata(existing, incoming) {
    let changed = false;
    const merged = { ...existing };

    if ((!existing.timestamp || !normalizeTimestampForKey(existing.timestamp)) && incoming.timestamp) {
        merged.timestamp = incoming.timestamp;
        changed = true;
    }

    if ((!existing.content || existing.content.length < (incoming.content || '').length) && incoming.content) {
        merged.content = incoming.content;
        changed = true;
    }

    if ((existing.type || 'message') !== (incoming.type || 'message') && incoming.type === 'system') {
        merged.type = incoming.type;
        changed = true;
    }

    const existingAuthor = normalizeAssetIdentityPart(existing.author);
    const incomingAuthor = normalizeAssetIdentityPart(incoming.author);
    if ((!existingAuthor || existingAuthor === 'unknown' || existingAuthor === 'unknown author') &&
        incomingAuthor && incomingAuthor !== 'unknown' && incomingAuthor !== 'unknown author') {
        merged.author = incoming.author;
        merged.avatarColor = stringToColor(incoming.author || 'Unknown');
        merged.initials = getInitials(incoming.author);
        changed = true;
    }

    if (hasNewAttachments(existing, incoming)) {
        merged.attachments = mergeAttachments(existing.attachments, incoming.attachments);
        merged.images = mergeImages(existing.images, incoming.images);
        changed = true;
    }

    return changed ? merged : null;
}

function getMessageRichnessScore(message) {
    if (!message) return 0;

    let score = 0;
    if (normalizeTimestampForKey(message.timestamp)) score += 8;

    const author = normalizeAssetIdentityPart(message.author);
    if (author && author !== 'unknown' && author !== 'unknown author') score += 4;

    const contentLength = (message.content || '').trim().length;
    score += Math.min(contentLength, 120) / 20;

    if ((message.type || 'message') === 'system') score += 2;
    score += (message.attachments?.length || 0) * 6;
    score += (message.images?.length || 0) * 6;
    if (message.isEdited) score += 2;
    if (message.isDeleted) score += 2;
    if (message.isStarred) score += 2;

    return score;
}

function mergeEditHistoryEntries(primary = [], secondary = []) {
    const merged = [...(primary || []), ...(secondary || [])];
    const seen = new Set();

    return merged.filter((entry) => {
        const key = JSON.stringify(entry || {});
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
    });
}

function mergeDuplicateMessageState(left, right) {
    const preferred = getMessageRichnessScore(right) > getMessageRichnessScore(left) ? right : left;
    const secondary = preferred === left ? right : left;

    let merged = { ...preferred };
    const metadataUpdate = mergeMessageMetadata(merged, secondary);
    if (metadataUpdate) {
        merged = { ...merged, ...metadataUpdate };
    }

    const mergedAttachments = mergeAttachments(preferred.attachments, secondary.attachments);
    if (mergedAttachments) {
        merged.attachments = mergedAttachments;
    }

    const mergedImages = mergeImages(preferred.images, secondary.images);
    if (mergedImages) {
        merged.images = mergedImages;
    }

    merged.isEdited = preferred.isEdited || secondary.isEdited;
    merged.editHistory = mergeEditHistoryEntries(preferred.editHistory, secondary.editHistory);
    merged.isDeleted = preferred.isDeleted || secondary.isDeleted;
    merged.deletedContent = preferred.deletedContent || secondary.deletedContent || null;
    merged.deletedAt = preferred.deletedAt || secondary.deletedAt || null;
    merged.isStarred = preferred.isStarred || secondary.isStarred;
    merged.starredAt = preferred.starredAt || secondary.starredAt || null;

    if ((!merged.avatarColor || !merged.initials) && (merged.author || secondary.author)) {
        const bestAuthor = merged.author || secondary.author || 'Unknown';
        merged.avatarColor = stringToColor(bestAuthor);
        merged.initials = getInitials(bestAuthor);
    }

    return merged;
}

function collapseRelaxedMeetingControlDuplicates(messages) {
    const deduped = [];
    // Only meeting-control messages can be duplicates, so search just those.
    const controlIndices = [];

    for (const message of messages) {
        if (!isRelaxedMeetingControlMessage(message)) {
            deduped.push(message);
            continue;
        }

        const normalizedContent = normalizeContentForLooseMatch(message.content);
        const duplicateIndex = controlIndices.find((index) =>
            normalizeContentForLooseMatch(deduped[index].content) === normalizedContent &&
            timestampsLooselyMatch(deduped[index].timestamp, message.timestamp)
        ) ?? -1;

        if (duplicateIndex < 0) {
            controlIndices.push(deduped.length);
            deduped.push(message);
            continue;
        }

        const merged = mergeDuplicateMessageState(deduped[duplicateIndex], message);
        merged._mergeSequence = Math.min(
            Number.isFinite(deduped[duplicateIndex]._mergeSequence) ? deduped[duplicateIndex]._mergeSequence : duplicateIndex,
            Number.isFinite(message._mergeSequence) ? message._mergeSequence : duplicateIndex
        );
        deduped[duplicateIndex] = merged;
    }

    return deduped;
}

// Save a complete chat backup
// Options:
//   - isPro: if false, enforces the free-tier chat limit and blocks updates to existing backups
//   - forceNew: if true, always creates a new chat entry
//   - existingChatId: chat saved by an earlier checkpoint of the same backup run;
//     skips the identity scan over every saved chat
async function saveChat(chatData) {
    const db = await openDB();
    const chatName = chatData.name || 'Unnamed Chat';
    const hasProAccess = typeof chatData.isPro === 'boolean'
        ? chatData.isPro
        : await getStoredProStatus();
    const allowIncompleteMerge = chatData.allowIncompleteMerge !== false;

    // Check if this chat already exists by stable identity
    let existingChat = null;
    if (chatData.existingChatId && !chatData.forceNew) {
        const hintedChat = await getChat(chatData.existingChatId);
        if (hintedChat && !hintedChat.trashedAt) {
            existingChat = hintedChat;
        }
    }
    if (!existingChat) {
        existingChat = await findChatByName(chatData);
    }

    if (existingChat && !chatData.forceNew) {
        const canResumeIncompleteChat = allowIncompleteMerge &&
            (existingChat.isPreview === true || existingChat.inProgress === true);

        if (!hasProAccess && !canResumeIncompleteChat) {
            return {
                id: existingChat.id,
                error: 'update_locked',
                message: 'This chat is already backed up. The free plan keeps the saved copy but cannot refresh it with newer messages. Upgrade to Pro to keep it up to date.'
            };
        }

        console.log('Found existing chat, merging messages:', existingChat.id);
        const mergeResult = await mergeChatMessages(existingChat.id, chatData);
        return { id: mergeResult.id, merged: true, newMessagesCount: mergeResult.newMessagesCount };
    }

    if (!hasProAccess && !chatData.forceNew) {
        const allowance = await getFreeBackupAllowance(chatData);
        if (!allowance.allowed) {
            return {
                id: null,
                error: 'chat_limit',
                reason: allowance.reason,
                count: allowance.activeCount,
                usedCount: allowance.usedCount,
                usedChatNames: allowance.usedChatNames,
                message: describeFreeLimit(allowance)
            };
        }
    }

    // Create new chat
    const chatId = generateId();

    const chat = {
        id: chatId,
        name: chatName,
        rawName: chatData.rawName || chatName,
        isChannel: chatData.isChannel === true,
        teamName: chatData.teamName || null,
        sourceKind: chatData.sourceKind || (chatData.isChannel ? 'channel' : 'chat'),
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
        messageCount: chatData.messages ? chatData.messages.length : 0,
        participants: chatData.participants || [],
        pdfBlob: chatData.pdfBlob || null,
        textContent: chatData.textContent || '',
        thumbnail: chatData.thumbnail || null,
        isPreview: chatData.isPreview || false, // Track if this is an incomplete/preview backup
        inProgress: chatData.inProgress === true,
        latestMessageAt: getLatestTimestampIso(chatData.messages || []),
        createdByPro: !!hasProAccess
    };

    // Check storage availability before save (P1-3)
    const storageCheck = await checkStorageAvailability();
    if (!storageCheck.available) {
        const error = new Error(storageCheck.message);
        error.type = 'quota';
        error.recoverable = false;
        throw error;
    }

    const saveResult = await new Promise((resolve, reject) => {
        const transaction = db.transaction(['chats', 'messages'], 'readwrite');

        transaction.onerror = () => {
            const classification = classifyIndexedDBError(transaction.error);
            console.error('Transaction error:', transaction.error);
            console.error('Error classification:', classification);

            const enhancedError = new Error(classification.userMessage);
            enhancedError.originalError = transaction.error;
            enhancedError.type = classification.type;
            enhancedError.recoverable = classification.recoverable;

            reject(enhancedError);
        };

        transaction.onabort = () => {
            const classification = classifyIndexedDBError(transaction.error);
            console.warn('Transaction aborted:', transaction.error);

            const enhancedError = new Error(classification.userMessage);
            enhancedError.originalError = transaction.error;
            enhancedError.type = classification.type;
            enhancedError.recoverable = classification.recoverable;

            reject(enhancedError);
        };

        transaction.oncomplete = () => {
            console.log('Chat saved successfully:', chatId);
            resolve({ id: chatId, merged: false, newMessagesCount: chatData.messages ? chatData.messages.length : 0 });
        };

        try {
            // Save chat metadata
            const chatsStore = transaction.objectStore('chats');
            chatsStore.add(chat);

            // Save individual messages
            if (chatData.messages && chatData.messages.length > 0) {
                const messagesStore = transaction.objectStore('messages');
                chatData.messages.forEach((msg, index) => {
                    const message = {
                        id: generateMessageId(),
                        chatId: chatId,
                        author: msg.author || 'Unknown',
                        timestamp: msg.timestamp || '',
                        content: msg.content || '',
                        type: msg.type || 'message',
                        avatarColor: stringToColor(msg.author || 'Unknown'),
                        initials: getInitials(msg.author),
                        order: index,
                        images: msg.images || undefined, // Store images array if present
                        attachments: msg.attachments || undefined, // Store file attachments array if present
                        // Preserve edit/delete/star state from imports
                        isEdited: msg.isEdited || false,
                        editHistory: msg.editHistory || [],
                        isDeleted: msg.isDeleted || false,
                        deletedContent: msg.deletedContent || null,
                        deletedAt: msg.deletedAt || null,
                        isStarred: msg.isStarred || false,
                        starredAt: msg.starredAt || null
                    };
                    messagesStore.add(message);
                });
            }
        } catch (error) {
            // Catch synchronous errors (e.g., constraint errors)
            transaction.abort();
            const classification = classifyIndexedDBError(error);
            const enhancedError = new Error(classification.userMessage);
            enhancedError.originalError = error;
            enhancedError.type = classification.type;
            enhancedError.recoverable = classification.recoverable;
            reject(enhancedError);
        }
    });

    if (!hasProAccess) {
        await recordFreeBackupUsage(chat);
    }

    return saveResult;
}

function getMessageSortTimestamp(message) {
    return parseFlexibleTimestamp(message?.timestamp || '');
}

function sortMessagesChronologically(messages) {
    // Parse each timestamp once instead of on every comparison.
    const entries = messages.map((message) => {
        const time = getMessageSortTimestamp(message);
        return {
            message,
            time,
            hasTime: Number.isFinite(time),
            seq: Number.isFinite(message._mergeSequence) ? message._mergeSequence : Number.MAX_SAFE_INTEGER,
            key: null
        };
    });
    const getKey = (entry) => {
        if (entry.key === null) entry.key = createMessageKey(entry.message);
        return entry.key;
    };

    entries.sort((a, b) => {
        if (a.hasTime && b.hasTime && a.time !== b.time) {
            return a.time - b.time; // Oldest first.
        }

        // Prefer deterministic timestamps when only one side can be parsed.
        if (a.hasTime !== b.hasTime) {
            return a.hasTime ? -1 : 1;
        }

        if (a.seq !== b.seq) {
            return a.seq - b.seq;
        }

        return getKey(a).localeCompare(getKey(b));
    });

    return entries.map(entry => entry.message);
}

function normalizeAssetIdentityPart(value) {
    return String(value || '').trim().toLowerCase();
}

function normalizeImageSourceForIdentity(src) {
    const normalized = normalizeAssetIdentityPart(src);
    if (!normalized || normalized.startsWith('data:') || normalized.startsWith('blob:')) {
        return '';
    }
    return normalized;
}

function getAttachmentIdentity(att) {
    if (!att) return '';
    return normalizeAssetIdentityPart(att.filename) ||
        normalizeAssetIdentityPart(att.sharePointUrl) ||
        normalizeAssetIdentityPart(att.arrivalTime);
}

function getImageIdentity(img) {
    if (!img) return '';

    return [
        normalizeAssetIdentityPart(img.filename),
        normalizeAssetIdentityPart(img.alt),
        Number.isFinite(Number(img.width)) ? String(Number(img.width)) : '',
        Number.isFinite(Number(img.height)) ? String(Number(img.height)) : '',
        normalizeImageSourceForIdentity(img.src)
    ].filter(Boolean).join('|');
}

function mergeAttachmentRecord(existing = {}, incoming = {}) {
    const hasData = !!(incoming.dataUrl || existing.dataUrl);
    return {
        ...existing,
        ...incoming,
        filename: incoming.filename || existing.filename || 'Unknown file',
        sharePointUrl: incoming.sharePointUrl || existing.sharePointUrl || '',
        fileDomain: incoming.fileDomain || existing.fileDomain || '',
        arrivalTime: incoming.arrivalTime || existing.arrivalTime || '',
        extension: incoming.extension || existing.extension || '',
        fileSizeText: incoming.fileSizeText || existing.fileSizeText || '',
        fileSizeMB: Number.isFinite(incoming.fileSizeMB)
            ? incoming.fileSizeMB
            : (Number.isFinite(existing.fileSizeMB) ? existing.fileSizeMB : 0),
        tooLarge: incoming.tooLarge === true || existing.tooLarge === true,
        dataUrl: incoming.dataUrl || existing.dataUrl || null,
        mimeType: incoming.mimeType || existing.mimeType || null,
        downloadError: incoming.downloadError || existing.downloadError || null,
        existsInDb: incoming.existsInDb === true || existing.existsInDb === true,
        checkpointPlaceholder: !hasData && (
            incoming.checkpointPlaceholder === true ||
            existing.checkpointPlaceholder === true
        )
    };
}

function mergeImageRecord(existing = {}, incoming = {}) {
    const hasData = !!(incoming.dataUrl || existing.dataUrl);
    return {
        ...existing,
        ...incoming,
        alt: incoming.alt || existing.alt || 'Image',
        width: Number.isFinite(incoming.width)
            ? incoming.width
            : (Number.isFinite(existing.width) ? existing.width : undefined),
        height: Number.isFinite(incoming.height)
            ? incoming.height
            : (Number.isFinite(existing.height) ? existing.height : undefined),
        src: incoming.src || existing.src || undefined,
        dataUrl: incoming.dataUrl || existing.dataUrl || null,
        checkpointPlaceholder: !hasData && (
            incoming.checkpointPlaceholder === true ||
            existing.checkpointPlaceholder === true
        )
    };
}

function normalizeAttachmentForComparison(att) {
    return [
        getAttachmentIdentity(att),
        att?.sharePointUrl || '',
        att?.fileSizeText || '',
        att?.tooLarge === true ? '1' : '0',
        att?.dataUrl ? '1' : '0',
        att?.downloadError || '',
        att?.checkpointPlaceholder === true ? '1' : '0'
    ].join('|');
}

function normalizeImageForComparison(img) {
    return [
        getImageIdentity(img),
        img?.dataUrl ? '1' : '0',
        img?.alt || '',
        Number.isFinite(img?.width) ? img.width : '',
        Number.isFinite(img?.height) ? img.height : '',
        img?.checkpointPlaceholder === true ? '1' : '0'
    ].join('|');
}

function didMediaCollectionChange(existing, merged, normalizeItem) {
    const existingItems = Array.isArray(existing) ? existing : [];
    const mergedItems = Array.isArray(merged) ? merged : [];

    if (existingItems.length !== mergedItems.length) {
        return true;
    }

    for (let i = 0; i < existingItems.length; i++) {
        if (normalizeItem(existingItems[i]) !== normalizeItem(mergedItems[i])) {
            return true;
        }
    }

    return false;
}

// Merge new messages into an existing chat (preserves local modifications)
async function mergeChatMessages(existingChatId, newChatData) {
    const db = await openDB();

    const existingMessages = await getMessages(existingChatId);

    const existingMessagesByKey = new Map();
    const relaxedMeetingControlMessages = [];
    existingMessages.forEach(m => {
        existingMessagesByKey.set(createMessageKey(m), m);
        if (isRelaxedMeetingControlMessage(m)) {
            relaxedMeetingControlMessages.push(m);
        }
    });

    const updatedExistingMessagesById = new Map();
    const newMessages = [];

    (newChatData.messages || []).forEach((msg, incomingIndex) => {
        const key = createMessageKey(msg);
        let existing = existingMessagesByKey.get(key);

        if (!existing && isRelaxedMeetingControlMessage(msg)) {
            existing = relaxedMeetingControlMessages.find((candidate) =>
                normalizeContentForLooseMatch(candidate.content) === normalizeContentForLooseMatch(msg.content) &&
                timestampsLooselyMatch(candidate.timestamp, msg.timestamp)
            ) || null;
        }

        if (!existing) {
            newMessages.push({
                id: generateMessageId(),
                chatId: existingChatId,
                author: msg.author || 'Unknown',
                timestamp: msg.timestamp || '',
                content: msg.content || '',
                type: msg.type || 'message',
                avatarColor: stringToColor(msg.author || 'Unknown'),
                initials: getInitials(msg.author),
                order: existingMessages.length + incomingIndex, // Temporary sequence, re-assigned after timeline sort.
                images: msg.images || undefined,
                attachments: msg.attachments || undefined,
                isEdited: msg.isEdited || false,
                editHistory: msg.editHistory || [],
                isDeleted: msg.isDeleted || false,
                deletedContent: msg.deletedContent || null,
                deletedAt: msg.deletedAt || null,
                isStarred: msg.isStarred || false,
                starredAt: msg.starredAt || null
            });
            return;
        }

        const metadataUpdate = mergeMessageMetadata(existing, msg);
        if (metadataUpdate) {
            updatedExistingMessagesById.set(existing.id, metadataUpdate);
        }
    });

    // Stored messages that this merge leaves untouched. They are rewritten only if
    // their position in the timeline changes, so a checkpoint that appends new
    // messages writes just those instead of the whole chat.
    const unchangedExistingMessages = new Set();
    const mergedExistingMessages = existingMessages.map((message, index) => {
        const updatedMessage = updatedExistingMessagesById.get(message.id);
        const entry = {
            ...(updatedMessage || message),
            _mergeSequence: Number.isFinite(message.order) ? message.order : index
        };
        if (!updatedMessage) {
            unchangedExistingMessages.add(entry);
        }
        return entry;
    });

    const mergedNewMessages = newMessages.map((message, index) => ({
        ...message,
        _mergeSequence: existingMessages.length + index
    }));

    const sortedTimeline = sortMessagesChronologically([
        ...mergedExistingMessages,
        ...mergedNewMessages
    ]);
    const collapsedTimeline = collapseRelaxedMeetingControlDuplicates(sortedTimeline);
    // Collapsing only removes entries; re-sort only when it merged something.
    const mergedTimeline = collapsedTimeline.length === sortedTimeline.length
        ? collapsedTimeline
        : sortMessagesChronologically(collapsedTimeline);
    const mergedMessageIds = new Set(mergedTimeline.map(message => message.id));
    const deletedExistingMessageIds = existingMessages
        .filter((message) => !mergedMessageIds.has(message.id))
        .map((message) => message.id);

    console.log(`Merging: ${existingMessages.length} existing, ${newChatData.messages?.length || 0} incoming, ${newMessages.length} new, ${updatedExistingMessagesById.size} to update, ${deletedExistingMessageIds.length} stale duplicates removed`);

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(['chats', 'messages'], 'readwrite');

        transaction.onerror = () => {
            const classification = classifyIndexedDBError(transaction.error);
            console.error('Transaction error:', transaction.error);
            console.error('Error classification:', classification);

            const enhancedError = new Error(classification.userMessage);
            enhancedError.originalError = transaction.error;
            enhancedError.type = classification.type;
            enhancedError.recoverable = classification.recoverable;

            reject(enhancedError);
        };

        transaction.onabort = () => {
            const classification = classifyIndexedDBError(transaction.error);
            console.warn('Transaction aborted:', transaction.error);

            const enhancedError = new Error(classification.userMessage);
            enhancedError.originalError = transaction.error;
            enhancedError.type = classification.type;
            enhancedError.recoverable = classification.recoverable;

            reject(enhancedError);
        };

        transaction.oncomplete = () => {
            console.log(`Merged ${newMessages.length} new messages, updated ${updatedExistingMessagesById.size} messages, removed ${deletedExistingMessageIds.length} stale duplicates`);
            resolve({ id: existingChatId, newMessagesCount: newMessages.length, updatedCount: updatedExistingMessagesById.size, removedCount: deletedExistingMessageIds.length });
        };

        const chatsStore = transaction.objectStore('chats');
        const getRequest = chatsStore.get(existingChatId);

        getRequest.onsuccess = () => {
            const chat = getRequest.result;
            if (chat) {
                chat.updatedAt = new Date().toISOString();
                chat.messageCount = mergedTimeline.length;
                chat.latestMessageAt = getLatestTimestampIso(mergedTimeline, { sortedChronologically: true });
                if (newChatData.pdfBlob) chat.pdfBlob = newChatData.pdfBlob;
                if (newChatData.textContent) chat.textContent = newChatData.textContent;
                if (newChatData.rawName) chat.rawName = newChatData.rawName;
                if (typeof newChatData.isChannel === 'boolean') chat.isChannel = newChatData.isChannel;
                if (Object.prototype.hasOwnProperty.call(newChatData, 'teamName')) {
                    chat.teamName = newChatData.teamName || null;
                }
                if (newChatData.sourceKind) chat.sourceKind = newChatData.sourceKind;
                if (newChatData.participants) {
                    const allParticipants = new Set([...(chat.participants || []), ...newChatData.participants]);
                    chat.participants = Array.from(allParticipants);
                }
                if (!newChatData.isPreview) {
                    chat.isPreview = false;
                } else {
                    chat.isPreview = true;
                }
                chat.inProgress = newChatData.inProgress === true;
                chatsStore.put(chat);
            }
        };

        const messagesStore = transaction.objectStore('messages');
        deletedExistingMessageIds.forEach((messageId) => {
            messagesStore.delete(messageId);
        });
        mergedTimeline.forEach((message, index) => {
            if (unchangedExistingMessages.has(message) && message.order === index) {
                return;
            }
            const { _mergeSequence, ...messageToSave } = message;
            messageToSave.order = index;
            messagesStore.put(messageToSave);
        });
    });
}

// Check if incoming message has attachments not in existing message
function hasNewAttachments(existing, incoming) {
    const mergedAttachments = mergeAttachments(existing.attachments, incoming.attachments);
    if (didMediaCollectionChange(existing.attachments, mergedAttachments, normalizeAttachmentForComparison)) {
        return true;
    }

    const mergedImages = mergeImages(existing.images, incoming.images);
    return didMediaCollectionChange(existing.images, mergedImages, normalizeImageForComparison);
}

// Merge attachments arrays, preferring existing data but adding new ones
function mergeAttachments(existing, incoming) {
    if (!incoming || incoming.length === 0) return existing;
    if (!existing || existing.length === 0) return incoming;

    const merged = [...existing];
    const indexByIdentity = new Map();
    merged.forEach((att, index) => {
        const identity = getAttachmentIdentity(att);
        if (identity) {
            indexByIdentity.set(identity, index);
        }
    });

    incoming.forEach(att => {
        const identity = getAttachmentIdentity(att);
        const existingIdx = identity ? indexByIdentity.get(identity) : -1;

        if (existingIdx === undefined || existingIdx < 0) {
            merged.push(att);
            if (identity) {
                indexByIdentity.set(identity, merged.length - 1);
            }
            return;
        }

        merged[existingIdx] = mergeAttachmentRecord(merged[existingIdx], att);
    });

    return merged;
}

// Merge images arrays
function mergeImages(existing, incoming) {
    if (!incoming || incoming.length === 0) return existing;
    if (!existing || existing.length === 0) return incoming;

    const merged = [...existing];
    const indexByIdentity = new Map();
    merged.forEach((img, index) => {
        const identity = getImageIdentity(img);
        if (identity) {
            indexByIdentity.set(identity, index);
        }
    });

    incoming.forEach(img => {
        const identity = getImageIdentity(img);
        const existingIdx = identity ? indexByIdentity.get(identity) : -1;

        if (existingIdx === undefined || existingIdx < 0) {
            merged.push(img);
            if (identity) {
                indexByIdentity.set(identity, merged.length - 1);
            }
            return;
        }

        merged[existingIdx] = mergeImageRecord(merged[existingIdx], img);
    });

    return merged;
}

// Get existing attachment filenames for a chat (used to skip re-downloading)
// Returns an array of filenames that already have data in the database
async function getExistingAttachmentFilenames(chatIdentity) {
    const existingChat = await findChatByName(chatIdentity);
    if (!existingChat) return [];

    const messages = await getMessages(existingChat.id);
    const filenames = new Set();

    messages.forEach(msg => {
        (msg.attachments || []).forEach(att => {
            // Only include files that already have data downloaded
            if (att.filename && att.dataUrl) {
                filenames.add(att.filename);
            }
        });
    });

    return Array.from(filenames);
}

async function getExistingChatId(chatIdentity) {
    const existingChat = await findChatByName(chatIdentity);
    return existingChat ? existingChat.id : null;
}

// sortedChronologically: messages come from sortMessagesChronologically, so the
// last parseable timestamp is the newest one.
function getLatestTimestampIso(messages, { sortedChronologically = false } = {}) {
    let latestTimestamp = null;

    for (let i = messages.length - 1; i >= 0; i--) {
        const parsed = parseFlexibleTimestamp(messages[i]?.timestamp || '');
        if (!Number.isFinite(parsed)) continue;
        if (!Number.isFinite(latestTimestamp) || parsed > latestTimestamp) {
            latestTimestamp = parsed;
        }
        if (sortedChronologically) break;
    }

    return Number.isFinite(latestTimestamp) ? new Date(latestTimestamp).toISOString() : null;
}

async function getLatestMessageTimestamp(chatIdentity) {
    const existingChat = await findChatByName(chatIdentity);
    if (!existingChat) return null;
    // An incomplete backup is missing older history, so it must be rescanned from
    // the top instead of stopping at its newest saved message.
    if (existingChat.isPreview === true || existingChat.inProgress === true) return null;

    // Saves since v6.6 keep this on the chat record; older backups need one scan.
    if (Object.prototype.hasOwnProperty.call(existingChat, 'latestMessageAt')) {
        return existingChat.latestMessageAt || null;
    }

    const messages = await getMessages(existingChat.id);
    return getLatestTimestampIso(messages);
}

// Get all active chats (excludes trashed, metadata only, sorted by date)
async function getChats() {
    const db = await openDB();

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(['chats'], 'readonly');
        const store = transaction.objectStore('chats');
        const index = store.index('createdAt');
        const request = index.openCursor(null, 'prev'); // Newest first

        const chats = [];

        request.onerror = () => reject(request.error);

        request.onsuccess = (event) => {
            const cursor = event.target.result;
            if (cursor) {
                const chat = { ...cursor.value };
                // Skip trashed chats
                if (!chat.trashedAt) {
                    // Don't include full blobs in list view
                    delete chat.pdfBlob;
                    delete chat.thumbnail;
                    chats.push(chat);
                }
                cursor.continue();
            } else {
                resolve(chats);
            }
        };
    });
}

// Get a single chat with all data
async function getChat(chatId) {
    const db = await openDB();

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(['chats'], 'readonly');
        const store = transaction.objectStore('chats');
        const request = store.get(chatId);

        request.onerror = () => reject(request.error);
        request.onsuccess = () => resolve(request.result);
    });
}

// Get all messages for a chat
async function getMessages(chatId) {
    const db = await openDB();

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(['messages'], 'readonly');
        const store = transaction.objectStore('messages');
        const index = store.index('chatId');
        const request = index.getAll(chatId);

        request.onerror = () => reject(request.error);
        request.onsuccess = () => {
            // Sort by order
            const messages = request.result.sort((a, b) => a.order - b.order);
            resolve(messages);
        };
    });
}

// Move chat to trash (soft delete with 60-day retention)
async function trashChat(chatId) {
    const db = await openDB();

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(['chats'], 'readwrite');
        const store = transaction.objectStore('chats');
        const getRequest = store.get(chatId);

        getRequest.onerror = () => reject(getRequest.error);

        getRequest.onsuccess = () => {
            const chat = getRequest.result;
            if (!chat) {
                reject(new Error('Chat not found'));
                return;
            }

            // Mark as trashed
            chat.trashedAt = new Date().toISOString();
            chat.expiresAt = new Date(Date.now() + TRASH_RETENTION_MS).toISOString();

            const putRequest = store.put(chat);
            putRequest.onerror = () => reject(putRequest.error);
            putRequest.onsuccess = () => resolve(chat);
        };
    });
}

// Restore chat from trash. Free users can restore only into an empty slot.
async function restoreChat(chatId) {
    const hasProAccess = await getStoredProStatus();
    if (!hasProAccess) {
        const activeCount = await getChatCount();
        if (activeCount >= FREE_CHAT_LIMIT) {
            const error = new Error(`The free plan keeps ${FREE_CHAT_LIMIT} backup at a time. Delete your current backup to restore this one, or upgrade to Pro.`);
            error.code = 'chat_limit';
            throw error;
        }
    }

    const db = await openDB();

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(['chats'], 'readwrite');
        const store = transaction.objectStore('chats');
        const getRequest = store.get(chatId);

        getRequest.onerror = () => reject(getRequest.error);

        getRequest.onsuccess = () => {
            const chat = getRequest.result;
            if (!chat) {
                reject(new Error('Chat not found'));
                return;
            }

            // Remove trash markers
            delete chat.trashedAt;
            delete chat.expiresAt;

            const putRequest = store.put(chat);
            putRequest.onerror = () => reject(putRequest.error);
            putRequest.onsuccess = () => resolve(chat);
        };
    });
}

// Get all trashed chats
async function getTrashedChats() {
    const db = await openDB();

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(['chats'], 'readonly');
        const store = transaction.objectStore('chats');
        // Use trashedAt index to only fetch trashed chats (where trashedAt is not null)
        // IDBKeyRange.lowerBound with open=false includes all valid date strings
        const index = store.index('trashedAt');
        const range = IDBKeyRange.lowerBound('', true); // Excludes empty string, gets all dates
        const request = index.getAll(range);

        request.onerror = () => reject(request.error);
        request.onsuccess = () => {
            // Process trashed chats, sorted by trashedAt (newest first)
            const trashedChats = (request.result || [])
                .filter(chat => chat.trashedAt) // Extra safety check
                .map(chat => {
                    // Calculate days remaining
                    const expiresAt = new Date(chat.expiresAt);
                    const now = new Date();
                    const daysRemaining = Math.max(0, Math.ceil((expiresAt - now) / (24 * 60 * 60 * 1000)));
                    return { ...chat, daysRemaining, pdfBlob: undefined, thumbnail: undefined };
                })
                .sort((a, b) => new Date(b.trashedAt) - new Date(a.trashedAt));
            resolve(trashedChats);
        };
    });
}

// Permanently delete a chat and its messages (no recovery)
async function permanentlyDeleteChat(chatId) {
    const db = await openDB();

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(['chats', 'messages'], 'readwrite');

        transaction.onerror = () => reject(transaction.error);
        transaction.oncomplete = () => resolve(true);

        // Delete chat
        const chatsStore = transaction.objectStore('chats');
        chatsStore.delete(chatId);

        // Delete associated messages
        const messagesStore = transaction.objectStore('messages');
        const index = messagesStore.index('chatId');
        const request = index.openCursor(IDBKeyRange.only(chatId));

        request.onsuccess = (event) => {
            const cursor = event.target.result;
            if (cursor) {
                cursor.delete();
                cursor.continue();
            }
        };
    });
}

// Purge expired trash (chats older than 60 days)
async function purgeExpiredTrash() {
    const db = await openDB();
    const now = new Date();

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(['chats', 'messages'], 'readwrite');
        const chatsStore = transaction.objectStore('chats');
        const messagesStore = transaction.objectStore('messages');
        const request = chatsStore.getAll();
        let purgedCount = 0;

        request.onerror = () => reject(request.error);
        request.onsuccess = () => {
            const expiredChats = (request.result || []).filter(chat => {
                if (!chat.trashedAt) return false;
                const expiresAt = new Date(chat.expiresAt);
                return now >= expiresAt;
            });

            expiredChats.forEach(chat => {
                // Delete chat
                chatsStore.delete(chat.id);
                purgedCount++;

                // Delete associated messages
                const msgIndex = messagesStore.index('chatId');
                const msgRequest = msgIndex.openCursor(IDBKeyRange.only(chat.id));
                msgRequest.onsuccess = (event) => {
                    const cursor = event.target.result;
                    if (cursor) {
                        cursor.delete();
                        cursor.continue();
                    }
                };
            });
        };

        transaction.oncomplete = () => {
            console.log(`Purged ${purgedCount} expired chats from trash`);
            resolve(purgedCount);
        };
    });
}

// Legacy deleteChat - now moves to trash (for backward compatibility)
async function deleteChat(chatId) {
    return trashChat(chatId);
}

// Update a single message (for edit/delete functionality)
async function updateMessage(messageId, updates) {
    const db = await openDB();

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(['messages'], 'readwrite');
        const store = transaction.objectStore('messages');
        const getRequest = store.get(messageId);

        getRequest.onerror = () => reject(getRequest.error);

        getRequest.onsuccess = () => {
            const message = getRequest.result;
            if (!message) {
                reject(new Error('Message not found'));
                return;
            }

            // Apply updates
            Object.assign(message, updates);
            message.updatedAt = new Date().toISOString();

            const putRequest = store.put(message);
            putRequest.onerror = () => reject(putRequest.error);
            putRequest.onsuccess = () => resolve(message);
        };
    });
}

// Edit a message content (stores history)
async function editMessage(messageId, newContent) {
    const db = await openDB();

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(['messages'], 'readwrite');
        const store = transaction.objectStore('messages');
        const getRequest = store.get(messageId);

        getRequest.onerror = () => reject(getRequest.error);

        getRequest.onsuccess = () => {
            const message = getRequest.result;
            if (!message) {
                reject(new Error('Message not found'));
                return;
            }

            // Initialize edit history if not exists
            if (!message.editHistory) {
                message.editHistory = [];
            }

            // Store the previous version
            message.editHistory.push({
                content: message.content,
                editedAt: new Date().toISOString()
            });

            // Update to new content
            message.content = newContent;
            message.isEdited = true;
            message.lastEditedAt = new Date().toISOString();

            const putRequest = store.put(message);
            putRequest.onerror = () => reject(putRequest.error);
            putRequest.onsuccess = () => resolve(message);
        };
    });
}

// Soft delete a message (keeps history)
async function softDeleteMessage(messageId) {
    const db = await openDB();

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(['messages'], 'readwrite');
        const store = transaction.objectStore('messages');
        const getRequest = store.get(messageId);

        getRequest.onerror = () => reject(getRequest.error);

        getRequest.onsuccess = () => {
            const message = getRequest.result;
            if (!message) {
                reject(new Error('Message not found'));
                return;
            }

            // Store original content before deleting
            message.originalContent = message.content;
            message.originalAuthor = message.author;
            message.isDeleted = true;
            message.deletedAt = new Date().toISOString();

            const putRequest = store.put(message);
            putRequest.onerror = () => reject(putRequest.error);
            putRequest.onsuccess = () => resolve(message);
        };
    });
}

// Restore a deleted message
async function restoreMessage(messageId) {
    const db = await openDB();

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(['messages'], 'readwrite');
        const store = transaction.objectStore('messages');
        const getRequest = store.get(messageId);

        getRequest.onerror = () => reject(getRequest.error);

        getRequest.onsuccess = () => {
            const message = getRequest.result;
            if (!message) {
                reject(new Error('Message not found'));
                return;
            }

            // Restore content
            if (message.originalContent) {
                message.content = message.originalContent;
            }
            message.isDeleted = false;
            delete message.deletedAt;

            const putRequest = store.put(message);
            putRequest.onerror = () => reject(putRequest.error);
            putRequest.onsuccess = () => resolve(message);
        };
    });
}

// Toggle star status on a message
async function toggleStarMessage(messageId) {
    const db = await openDB();

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(['messages'], 'readwrite');
        const store = transaction.objectStore('messages');
        const getRequest = store.get(messageId);

        getRequest.onerror = () => reject(getRequest.error);

        getRequest.onsuccess = () => {
            const message = getRequest.result;
            if (!message) {
                reject(new Error('Message not found'));
                return;
            }

            // Toggle star status
            message.isStarred = !message.isStarred;
            message.starredAt = message.isStarred ? new Date().toISOString() : null;

            const putRequest = store.put(message);
            putRequest.onerror = () => reject(putRequest.error);
            putRequest.onsuccess = () => resolve(message);
        };
    });
}

// Get all starred messages for a chat
async function getStarredMessages(chatId) {
    const messages = await getMessages(chatId);
    return messages.filter(msg => msg.isStarred);
}

// Get storage usage estimate
async function getStorageUsage() {
    if (navigator.storage && navigator.storage.estimate) {
        const estimate = await navigator.storage.estimate();
        return {
            used: estimate.usage || 0,
            quota: estimate.quota || 0,
            percentUsed: estimate.quota ? ((estimate.usage / estimate.quota) * 100).toFixed(2) : 0
        };
    }
    return { used: 0, quota: 0, percentUsed: 0 };
}

// Count active chats only (trashed chats do not count toward free-tier limits or the UI badge).
// Counts keys instead of loading every chat record.
async function getChatCount() {
    const db = await openDB();

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(['chats'], 'readonly');
        const store = transaction.objectStore('chats');
        const totalRequest = store.count();
        // Only records with a trashedAt value appear in this index.
        const trashedRequest = store.index('trashedAt').count();

        transaction.onerror = () => reject(transaction.error);
        transaction.oncomplete = () => {
            resolve(Math.max(0, totalRequest.result - trashedRequest.result));
        };
    });
}

async function markInProgressBackupsInterrupted() {
    const db = await openDB();

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(['chats'], 'readwrite');
        const store = transaction.objectStore('chats');
        const request = store.openCursor();
        let updatedCount = 0;

        transaction.onerror = () => reject(transaction.error);
        transaction.onabort = () => reject(transaction.error);
        transaction.oncomplete = () => resolve(updatedCount);

        request.onerror = () => reject(request.error);
        request.onsuccess = (event) => {
            const cursor = event.target.result;
            if (!cursor) return;

            const chat = cursor.value;
            if (chat && chat.inProgress === true) {
                chat.inProgress = false;
                chat.isPreview = true;
                chat.updatedAt = new Date().toISOString();
                cursor.update(chat);
                updatedCount += 1;
            }

            cursor.continue();
        };
    });
}

// Export for use in extension pages and the background service worker.
if (typeof globalThis !== 'undefined') {
    globalThis.TeamsBackupDB = {
        openDB,
        saveChat,
        findChatByName,
        mergeChatMessages,
        getExistingAttachmentFilenames,
        getExistingChatId,
        getLatestMessageTimestamp,
        getChats,
        getChat,
        getMessages,
        markInProgressBackupsInterrupted,
        deleteChat, // Legacy - now moves to trash
        trashChat,
        restoreChat,
        getTrashedChats,
        permanentlyDeleteChat,
        purgeExpiredTrash,
        updateMessage,
        editMessage,
        softDeleteMessage,
        restoreMessage,
        toggleStarMessage,
        getStarredMessages,
        getStorageUsage,
        getChatCount,
        getFreeBackupUsage,
        migrateFreeBackupUsageToLocal,
        getFreeBackupAllowance,
        describeFreeLimit,
        stringToColor,
        getInitials,
        TRASH_RETENTION_DAYS,
        FREE_CHAT_LIMIT
    };
}
