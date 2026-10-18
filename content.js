// content.js

let isCancelled = false;
let isScrolling = false;
let availableChatsScan = null; // in-flight sidebar scan shared by overlapping requests
// The full sidebar scan toggles Teams filters and views, so reuse its result for
// a few minutes instead of repeating it every time the popup opens.
const AVAILABLE_CHATS_CACHE_TTL_MS = 3 * 60 * 1000;
let availableChatsCache = null; // { chats, diagnostics, at }

// Operation tracking system for reliability (prevents concurrent backup corruption)
let currentOperationId = null;
let operationQueue = [];
let currentOperationStartedAt = 0;
let currentOperationLastActivityAt = 0;
const OPERATION_TIMEOUT_MS = 4 * 60 * 60 * 1000; // Long chats can take well over 30 minutes on Teams.

// Generate a unique operation ID for tracking backup requests
function generateOperationId() {
    return 'op_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
}

function parseOperationTimestamp(opId) {
    if (!opId) return NaN;
    const timestamp = parseInt(opId.split('_')[1], 10);
    return Number.isFinite(timestamp) ? timestamp : NaN;
}

function clearCurrentOperationState() {
    currentOperationId = null;
    currentOperationStartedAt = 0;
    currentOperationLastActivityAt = 0;
}

function markOperationActivity(operationId = currentOperationId) {
    if (!operationId || currentOperationId !== operationId) return;
    if (!currentOperationStartedAt) {
        const parsedTimestamp = parseOperationTimestamp(operationId);
        currentOperationStartedAt = Number.isFinite(parsedTimestamp) ? parsedTimestamp : Date.now();
    }
    currentOperationLastActivityAt = Date.now();
}

// Check if an operation is still active (within timeout)
function isOperationActive(opId) {
    if (!opId) return false;
    const timestamp = opId === currentOperationId && currentOperationLastActivityAt
        ? currentOperationLastActivityAt
        : parseOperationTimestamp(opId);
    return (Date.now() - timestamp) < OPERATION_TIMEOUT_MS;
}

// Clean up stale operations
function cleanupStaleOperations() {
    const now = Date.now();
    operationQueue = operationQueue.filter(op => {
        const timestamp = parseOperationTimestamp(op.operationId);
        return (now - timestamp) < OPERATION_TIMEOUT_MS;
    });
}

// Check if a backup is already in progress
function isBackupInProgress() {
    if (currentOperationId) {
        if (isOperationActive(currentOperationId)) {
            return true;
        }
        // Operation expired, clean up
        clearCurrentOperationState();
    }
    return false;
}

function getOperationAgeMs(opId = currentOperationId) {
    if (!opId) return null;
    const timestamp = opId === currentOperationId && currentOperationStartedAt
        ? currentOperationStartedAt
        : parseOperationTimestamp(opId);
    if (!Number.isFinite(timestamp)) return null;
    return Math.max(0, Date.now() - timestamp);
}

function isOperationCancelled(operationId = null) {
    return isCancelled || (!!operationId && currentOperationId !== operationId);
}

function getBackupStatusPayload() {
    cleanupStaleOperations();
    const inProgress = isBackupInProgress();
    return {
        status: inProgress ? "busy" : "idle",
        inProgress,
        operationId: inProgress ? currentOperationId : null,
        ageMs: inProgress ? getOperationAgeMs(currentOperationId) : null,
        lastActivityAgeMs: inProgress && currentOperationLastActivityAt
            ? Math.max(0, Date.now() - currentOperationLastActivityAt)
            : null,
        timeoutMs: OPERATION_TIMEOUT_MS
    };
}

const CHECKPOINT_BATCH_MESSAGE_LIMIT = 200;
const CHECKPOINT_BATCH_SIZE_LIMIT_BYTES = 1.5 * 1024 * 1024;
const CHECKPOINT_BATCH_MAX_AGE_MS = 15000;

function sendRuntimeMessageWithResponse(message) {
    if (message?.operationId) {
        markOperationActivity(message.operationId);
    }
    return new Promise((resolve, reject) => {
        chrome.runtime.sendMessage(message, (response) => {
            if (chrome.runtime.lastError) {
                reject(new Error(chrome.runtime.lastError.message));
                return;
            }
            if (message?.operationId) {
                markOperationActivity(message.operationId);
            }
            resolve(response);
        });
    });
}

// Track Blob URLs for cleanup to prevent memory leaks
const createdBlobUrls = new Set();

// Create a Blob URL and track it for later cleanup
function createTrackedBlobUrl(blob) {
    const url = URL.createObjectURL(blob);
    createdBlobUrls.add(url);
    return url;
}

// Revoke a specific Blob URL and remove from tracking
function revokeTrackedBlobUrl(url) {
    if (url && createdBlobUrls.has(url)) {
        URL.revokeObjectURL(url);
        createdBlobUrls.delete(url);
    }
}

// Revoke all tracked Blob URLs (call on backup complete/cancel/error)
function revokeAllBlobUrls() {
    console.log(`Cleaning up ${createdBlobUrls.size} Blob URLs`);
    createdBlobUrls.forEach(url => {
        URL.revokeObjectURL(url);
    });
    createdBlobUrls.clear();
}

// Capture a lightweight DOM snapshot for diagnostics
function captureDomDiagnostics() {
    try {
        const mainArea = document.querySelector('[role="main"]') || document.querySelector('[data-tid="app-layout-area--main"]');
        // Capture data-tid tree inside the main content area for debugging channel variants
        const mainAreaTids = mainArea
            ? [...new Set([...mainArea.querySelectorAll('[data-tid]')].map(e => e.getAttribute('data-tid')))].sort().slice(0, 40)
            : [];
        const mainAreaRoles = mainArea
            ? [...new Set([...mainArea.querySelectorAll('[role]')].map(e => e.getAttribute('role')))].sort()
            : [];
        // Capture the first 15 "interesting" elements inside the selected container (or layout)
        const container = cachedChatContainer || document.querySelector('[data-tid="message-pane-layout"]');
        const containerTree = container
            ? [...container.querySelectorAll('*')].slice(0, 200).map(el => {
                const tid = el.getAttribute('data-tid');
                const role = el.getAttribute('role');
                const cls = typeof el.className === 'string' ? el.className : '';
                const hasTime = el.tagName === 'TIME';
                const isFui = cls.includes('fui-') && !cls.includes('fui-Primitive') && !cls.includes('fui-Flex');
                if (!tid && !role && !isFui && !hasTime) return null;
                return {
                    tag: el.tagName,
                    tid: tid || undefined,
                    role: role || undefined,
                    cls: isFui ? cls.substring(0, 80) : undefined,
                    time: hasTime ? (el.getAttribute('datetime') || el.textContent?.substring(0, 30)) : undefined
                };
              }).filter(Boolean).slice(0, 30)
            : [];
        return {
            url: window.location.href,
            title: document.title,
            hasRoleMain: !!document.querySelector('[role="main"]'),
            hasAppLayoutMain: !!document.querySelector('[data-tid="app-layout-area--main"]'),
            hasMessagePaneViewport: !!document.querySelector('[data-tid="message-pane-list-viewport"]'),
            hasMessagePaneLayout: !!document.querySelector('[data-tid="message-pane-layout"]'),
            hasChannelPaneViewport: !!document.querySelector('[data-tid="channel-pane-viewport"]'),
            hasChannelEmptyState: !!document.querySelector('[data-tid="channel-empty-state"]'),
            chatPaneItemCount: document.querySelectorAll('[data-tid="chat-pane-item"]').length,
            postItemCount: document.querySelectorAll('[data-tid*="post-item"]').length,
            replyItemCount: document.querySelectorAll('[data-tid*="reply-item"]').length,
            threadItemCount: document.querySelectorAll('[data-tid*="thread-item"]').length,
            fuiUnstableChatItemCount: document.querySelectorAll('[class*="fui-unstable-ChatItem"]').length,
            fuiChatMessageCount: document.querySelectorAll('[class*="fui-ChatMessage"]').length,
            fuiChatMyMessageCount: document.querySelectorAll('[class*="fui-ChatMyMessage"]').length,
            roleLog: document.querySelectorAll('[role="log"]').length,
            roleFeed: document.querySelectorAll('[role="feed"]').length,
            roleList: document.querySelectorAll('[role="list"]').length,
            roleArticle: document.querySelectorAll('[role="article"]').length,
            timeElementCount: document.querySelectorAll('time').length,
            mainAreaTids,
            mainAreaRoles,
            containerTree,
            scrollableContainers: [...document.querySelectorAll('div')].filter(d => {
                const s = window.getComputedStyle(d);
                return (s.overflowY === 'auto' || s.overflowY === 'scroll') && d.scrollHeight > d.clientHeight + 20;
            }).slice(0, 5).map(d => ({
                dataTid: d.getAttribute('data-tid') || '',
                role: d.getAttribute('role') || '',
                className: (typeof d.className === 'string' ? d.className : '').substring(0, 60),
                dims: `${d.scrollHeight}/${d.clientHeight}`
            }))
        };
    } catch (e) {
        return { error: e.message };
    }
}

// DOM query caching for performance - getChatContainer runs 20+ selectors
let cachedChatContainer = null;
let containerCacheTimestamp = 0;
const CONTAINER_CACHE_TTL = 5000; // 5 seconds - container rarely changes during backup
let lastContainerSelectionDiagnostics = null;

// Mutation observer for chat container changes (P1-5)
let containerMutationObserver = null;
let observedContainer = null;

// Set up mutation observer to detect DOM changes that might invalidate our container cache
function setupContainerObserver(container) {
    // Clean up existing observer
    if (containerMutationObserver && observedContainer) {
        containerMutationObserver.disconnect();
        containerMutationObserver = null;
        observedContainer = null;
    }

    if (!container) return;

    try {
        containerMutationObserver = new MutationObserver((mutations) => {
            // Check if significant structural changes occurred
            const hasSignificantChange = mutations.some(mutation => {
                // Type of changes that might invalidate our container
                return mutation.type === 'childList' &&
                       (mutation.addedNodes.length > 0 || mutation.removedNodes.length > 0);
            });

            if (hasSignificantChange) {
                console.log('Container structure changed, invalidating cache');
                invalidateContainerCache();
            }
        });

        // Observe the container for structural changes
        containerMutationObserver.observe(container, {
            childList: true,    // Watch for added/removed children
            subtree: false,     // Only watch direct children (performance)
            attributes: false,  // Don't watch attributes
            characterData: false // Don't watch text content
        });

        observedContainer = container;
        console.log('Mutation observer set up for container');
    } catch (error) {
        console.warn('Failed to set up mutation observer:', error);
    }
}

// Clean up mutation observer
function cleanupContainerObserver() {
    if (containerMutationObserver) {
        containerMutationObserver.disconnect();
        containerMutationObserver = null;
        observedContainer = null;
    }
}

// Invalidate the container cache (call when switching chats or starting new backup)
function invalidateContainerCache() {
    cachedChatContainer = null;
    containerCacheTimestamp = 0;
    lastContainerSelectionDiagnostics = null;
    // Also clean up observer since container is being invalidated
    cleanupContainerObserver();
}

// Utility function to sleep for a given number of milliseconds
function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
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

    const withoutAt = Date.parse(normalizedRaw.replace(/\bat\b/ig, ' '));
    if (Number.isFinite(withoutAt)) return inferMissingYear(withoutAt, normalizedRaw);

    return NaN;
}

// Process async tasks in parallel with concurrency limit
// Returns results in same order as input items
async function processWithConcurrency(items, asyncFn, concurrencyLimit = 3, onProgress = null) {
    const results = new Array(items.length);
    let currentIndex = 0;
    let completedCount = 0;

    async function processNext() {
        while (currentIndex < items.length) {
            const index = currentIndex++;
            const item = items[index];
            try {
                results[index] = await asyncFn(item, index);
            } catch (error) {
                results[index] = { error: error.message || 'Unknown error' };
            }
            completedCount++;
            if (onProgress) {
                onProgress(completedCount, items.length);
            }
        }
    }

    // Start concurrent workers up to the limit
    const workers = [];
    for (let i = 0; i < Math.min(concurrencyLimit, items.length); i++) {
        workers.push(processNext());
    }

    await Promise.all(workers);
    return results;
}

// Function to get all visible chats from Teams sidebar
// Supports: Personal Teams, Business Teams (Chats), Business Teams (Teams/Channels)
function normalizeChatIdentityPart(value) {
    return (value || '')
        .toString()
        .trim()
        .toLowerCase()
        .replace(/\s+/g, ' ');
}

function buildChatIdentityKey({ rawName, name, teamName, isChannel }) {
    const normalizedName = normalizeChatIdentityPart(rawName || name);
    const normalizedTeam = normalizeChatIdentityPart(teamName);
    const type = isChannel ? 'channel' : 'chat';
    return `${type}|${normalizedName}|${normalizedTeam}`;
}

function looksLikeSeeAllNavigation(text) {
    const normalizedText = normalizeChatIdentityPart(text);
    return normalizedText.includes('see all') ||
        normalizedText.includes('view all') ||
        normalizedText.includes('ver todo');
}

function looksLikeNewChatNavigation(text, dataTid = '') {
    const normalizedText = normalizeChatIdentityPart(text);
    const normalizedDataTid = normalizeChatIdentityPart(dataTid);
    return normalizedText === 'new message' ||
        normalizedText === 'new chat' ||
        normalizedText === 'start a new chat' ||
        normalizedText === 'nuevo mensaje' ||
        normalizedText === 'nuevo chat' ||
        normalizedDataTid.includes('new-message') ||
        normalizedDataTid.includes('new-chat') ||
        normalizedDataTid.includes('compose-chat');
}

function isChatSectionHeaderText(text, dataTid = '') {
    const normalizedText = normalizeChatIdentityPart(text);
    const normalizedDataTid = normalizeChatIdentityPart(dataTid);
    return normalizedText.startsWith('favorites') ||
        normalizedText.startsWith('favourites') ||
        normalizedText.startsWith('favoritos') ||
        normalizedText.startsWith('chats') ||
        normalizedText.startsWith('chat') ||
        normalizedText.startsWith('recent') ||
        normalizedText.startsWith('recientes') ||
        normalizedText.startsWith('pinned') ||
        normalizedText.startsWith('fijado') ||
        normalizedText.startsWith('messages') ||
        normalizedText.startsWith('mensajes') ||
        normalizedDataTid.includes('chat') ||
        normalizedDataTid.includes('pinned') ||
        normalizedDataTid.includes('recent') ||
        normalizedDataTid.includes('favorite');
}

function isTeamsSectionHeaderText(text, dataTid = '') {
    const normalizedText = normalizeChatIdentityPart(text);
    const normalizedDataTid = normalizeChatIdentityPart(dataTid);
    return normalizedText.startsWith('teams and channels') ||
        normalizedText.startsWith('teams & channels') ||
        normalizedText.startsWith('equipos y canales') ||
        normalizedText.startsWith('channels') ||
        normalizedText.startsWith('canales') ||
        normalizedDataTid.includes('teams') ||
        normalizedDataTid.includes('channels');
}

function getConversationFilterToggle(filterKind) {
    const acceptedLabels = filterKind === 'channel'
        ? ['channel', 'channels', 'canal', 'canales']
        : filterKind === 'unread'
            ? ['unread', 'no leido', 'no leídos', 'no leidas', 'no leídas']
            : ['chat', 'chats'];

    // The filter toolbar is a sibling of the sidebar tree in current Teams,
    // so search the document and constrain candidates to the conversation pane.
    const candidates = document.querySelectorAll('[aria-pressed], button, [role="button"]');
    for (const candidate of candidates) {
        const label = normalizeChatIdentityPart(
            candidate.getAttribute('aria-label') ||
            candidate.getAttribute('title') ||
            candidate.textContent || ''
        );
        const isMatchingLabel = acceptedLabels.some(accepted =>
            label === accepted || label.startsWith(`${accepted} (`));
        if (!isMatchingLabel) continue;

        const rect = candidate.getBoundingClientRect();
        if (rect.width <= 0 || rect.height <= 0) continue;
        if (rect.left < 45) continue;
        if (rect.left > Math.min(window.innerWidth * 0.4, 420)) continue;

        // Quick filters are real toggles. Requiring aria-pressed keeps section
        // headers and individual conversations with similar names out.
        if (candidate.getAttribute('aria-pressed') !== null) {
            return candidate;
        }
    }

    return null;
}

function getChatFilterToggle() {
    return getConversationFilterToggle('chat');
}

function getChannelFilterToggle() {
    return getConversationFilterToggle('channel');
}

function getConversationFilterState() {
    return {
        chat: isTogglePressed(getChatFilterToggle()),
        channel: isTogglePressed(getChannelFilterToggle()),
        unread: isTogglePressed(getConversationFilterToggle('unread'))
    };
}

async function applyConversationFilterState(nextState) {
    for (const kind of ['chat', 'channel', 'unread']) {
        const toggle = getConversationFilterToggle(kind);
        if (!toggle) continue;
        const shouldBePressed = !!nextState?.[kind];
        if (isTogglePressed(toggle) === shouldBePressed) continue;
        if (clickElementBestEffort(toggle)) {
            await sleep(220);
        }
    }
}

function isTogglePressed(toggle) {
    return toggle?.getAttribute?.('aria-pressed') === 'true';
}

function mergeDetectedChats(...chatLists) {
    const merged = [];
    const seen = new Set();

    chatLists.flat().forEach(chat => {
        if (!chat) return;
        const key = chat.key || buildChatIdentityKey(chat);
        if (seen.has(key)) return;
        seen.add(key);
        merged.push(chat);
    });

    return merged.map((chat, index) => ({ ...chat, id: index }));
}

function getChatSidebarScrollContainer() {
    const sidebarRoot = getChatSidebarRoot();
    if (!sidebarRoot) return null;

    const candidates = [sidebarRoot, ...sidebarRoot.querySelectorAll('*')];
    let best = null;
    let bestScore = 0;

    for (const candidate of candidates) {
        const overflow = candidate.scrollHeight - candidate.clientHeight;
        if (overflow <= 8 || candidate.clientHeight < 80) continue;

        const rect = candidate.getBoundingClientRect();
        if (rect.width <= 0 || rect.height <= 0) continue;
        if (rect.left > Math.min(window.innerWidth * 0.4, 420)) continue;

        const overflowY = getComputedStyle(candidate).overflowY;
        const score = overflow + (overflowY === 'auto' || overflowY === 'scroll' ? 500 : 0);
        if (score > bestScore) {
            best = candidate;
            bestScore = score;
        }
    }

    return best;
}

function elementLooksSelected(element) {
    if (!element) return false;
    const selectionHosts = [element.closest('[aria-selected], [aria-current], [aria-pressed]')];
    let ancestor = element;
    for (let depth = 0; ancestor && depth < 5; depth++) {
        selectionHosts.push(ancestor);
        ancestor = ancestor.parentElement;
    }

    return selectionHosts.filter(Boolean).some(host => {
        const className = typeof host.className === 'string' ? host.className.toLowerCase() : '';
        const hasSelectedClass = className.split(/\s+/).some(token =>
            /(?:^|[-_])(selected|active|current)(?:$|[-_])/.test(token));
        const ariaCurrent = (host.getAttribute('aria-current') || '').toLowerCase();
        return host.getAttribute('aria-selected') === 'true' ||
            host.getAttribute('aria-pressed') === 'true' ||
            host.getAttribute('data-is-selected') === 'true' ||
            host.getAttribute('data-active') === 'true' ||
            ariaCurrent === 'page' ||
            ariaCurrent === 'true' ||
            hasSelectedClass;
    });
}

function clickElementBestEffort(element) {
    if (!element || typeof element.click !== 'function') return false;
    try {
        element.click();
        return true;
    } catch (error) {
        console.warn('Best-effort click failed:', error?.message || error);
        return false;
    }
}

function conversationRailLabelMatches(value, appKind) {
    const normalized = normalizeChatIdentityPart(value);
    if (!normalized) return false;

    const accepted = appKind === 'teams'
        ? ['teams', 'teams app', 'teams and channels', 'teams & channels']
        : ['chat', 'chats', 'chat app', 'chats app'];

    return accepted.includes(normalized) || accepted.some(label =>
        normalized.startsWith(`${label} (`) ||
        normalized.startsWith(`${label}, `) ||
        normalized.startsWith(`${label} - `));
}

function findConversationRailButton(appKind) {
    const candidates = document.querySelectorAll('[role="tab"], [role="button"], button, a');
    for (const candidate of candidates) {
        const text = normalizeChatIdentityPart(candidate.textContent || '');
        const ariaLabel = normalizeChatIdentityPart(candidate.getAttribute('aria-label') || '');
        const title = normalizeChatIdentityPart(candidate.getAttribute('title') || '');
        const dataTid = normalizeChatIdentityPart(candidate.getAttribute('data-tid') || '');
        const combined = `${text} ${ariaLabel} ${title} ${dataTid}`.trim();
        if (!combined) continue;
        if (candidate.closest('[role="tree"]')) continue;

        const labelMatch = [text, ariaLabel, title].some(value =>
            conversationRailLabelMatches(value, appKind));
        const dataTidMatch = (dataTid.includes('app-bar') || dataTid.includes('appbar')) &&
            (appKind === 'teams'
                ? dataTid.includes('team')
                : dataTid.includes('chat'));
        const describedMatch = appKind === 'teams'
            ? combined.includes('app bar teams')
            : combined.includes('app bar chat');
        if (!labelMatch && !dataTidMatch && !describedMatch) continue;

        const rect = candidate.getBoundingClientRect();
        if (rect.width <= 0 || rect.height <= 0) continue;
        if (rect.left > 140) continue;
        if (rect.top > window.innerHeight - 40) continue;

        return candidate;
    }
    return null;
}

function findChatRailButton() {
    return findConversationRailButton('chat');
}

function findTeamsRailButton() {
    return findConversationRailButton('teams');
}

function findCollapsedSidebarExpandControl() {
    const directControl = document.querySelector(
        'button[data-testid="resize-handler-divider-button-end-mid-nav"], button[data-maximizing-button="true"]'
    );
    if (directControl) {
        const rect = directControl.getBoundingClientRect();
        if (rect.width > 0 && rect.height > 0 && rect.left <= Math.min(window.innerWidth * 0.25, 280)) {
            return directControl;
        }
    }

    const candidates = document.querySelectorAll('button, [role="button"]');
    for (const candidate of candidates) {
        const text = normalizeChatIdentityPart(candidate.textContent || '');
        const ariaLabel = normalizeChatIdentityPart(candidate.getAttribute('aria-label') || '');
        const title = normalizeChatIdentityPart(candidate.getAttribute('title') || '');
        const dataTid = normalizeChatIdentityPart(candidate.getAttribute('data-tid') || '');
        const combined = `${text} ${ariaLabel} ${title} ${dataTid}`.trim();
        const isSidebarExpandControl = combined.includes('select or drag to expand') ||
            combined.includes('resize larger') ||
            ((combined.includes('expand') || combined.includes('show')) &&
                (combined.includes('left panel') || combined.includes('side panel') || combined.includes('navigation pane')));
        if (!isSidebarExpandControl) continue;

        const rect = candidate.getBoundingClientRect();
        if (rect.width <= 0 || rect.height <= 0) continue;
        if (rect.left > Math.min(window.innerWidth * 0.25, 280)) continue;

        return candidate;
    }
    return null;
}

function getChatSidebarRoot() {
    const candidates = [
        document.querySelector('[role="tree"]'),
        document.querySelector('[data-tid="chat-pane"]'),
        document.querySelector('[data-tid*="chat-list"]'),
        document.querySelector('[data-tid*="conversation-list"]'),
        document.querySelector('[role="navigation"]')
    ].filter(Boolean);

    let best = null;
    let bestScore = -Infinity;

    for (const candidate of candidates) {
        const rect = candidate.getBoundingClientRect();
        if (rect.width <= 0 || rect.height <= 0) continue;
        if (rect.right <= 0 || rect.bottom <= 0) continue;
        if (rect.left > Math.min(window.innerWidth * 0.4, 420)) continue;
        // The app bar can be taller than a short conversation list. Only use
        // generic navigation when no conversation-specific root is available.
        if (candidate.matches('[role="navigation"]') && best) continue;

        const score = Math.min(rect.height, window.innerHeight) +
            Math.max(0, 420 - rect.left) +
            candidate.querySelectorAll('[role="treeitem"], [aria-expanded]').length * 20;

        if (score > bestScore) {
            best = candidate;
            bestScore = score;
        }
    }

    return best;
}

function getSidebarSectionToggleRecords() {
    const toggles = [];
    const sidebarRoot = getChatSidebarRoot() || document;
    const candidates = sidebarRoot.querySelectorAll(
        '[aria-expanded], [role="treeitem"], [role="button"], button, [data-tid], [class*="header"], [class*="section"], [class*="group"]'
    );
    const seen = new Set();
    const separateTeamsViewActive = elementLooksSelected(findTeamsRailButton()) && !!findChatRailButton();

    for (const candidate of candidates) {
        const text = candidate.textContent || '';
        const dataTid = candidate.getAttribute('data-tid') || '';
        const rect = candidate.getBoundingClientRect();
        if (rect.width <= 0 || rect.height <= 0) continue;
        if (rect.left > Math.min(window.innerWidth * 0.4, 420)) continue;

        const isChatSection = isChatSectionHeaderText(text, dataTid);
        const isTeamsSection = isTeamsSectionHeaderText(text, dataTid);
        const isSeparateTeamsTreeToggle = separateTeamsViewActive &&
            candidate.matches('[role="treeitem"][aria-expanded]') &&
            !isChatSection;
        if (!isChatSection && !isTeamsSection && !isSeparateTeamsTreeToggle) continue;

        const clickTarget = candidate.matches('[aria-expanded], [role="treeitem"], [role="button"], button')
            ? candidate
            : candidate.querySelector('[aria-expanded], [role="treeitem"], [role="button"], button') || candidate;
        const expandedSource = clickTarget?.hasAttribute?.('aria-expanded')
            ? clickTarget
            : candidate.hasAttribute('aria-expanded')
                ? candidate
                : candidate.querySelector('[aria-expanded]');
        const expandedState = expandedSource?.getAttribute?.('aria-expanded') ?? null;

        const recordKey = `${normalizeChatIdentityPart(text)}|${normalizeChatIdentityPart(dataTid)}`;
        if (seen.has(recordKey)) continue;
        seen.add(recordKey);

        toggles.push({
            element: candidate,
            clickTarget,
            text,
            dataTid,
            expandedState,
            isTeamsSection: isTeamsSection || isSeparateTeamsTreeToggle
        });
    }

    return toggles;
}

function getChatSidebarDiagnostics() {
    const treeItems = Array.from(document.querySelectorAll('[role="tree"] [role="treeitem"]'))
        .filter(item => {
            const rect = item.getBoundingClientRect();
            return rect.width > 0 && rect.height > 0;
        });
    const railButton = findChatRailButton();
    const teamsRailButton = findTeamsRailButton();
    const sidebarExpandControl = findCollapsedSidebarExpandControl();
    const collapsedToggles = getSidebarSectionToggleRecords().filter(toggle => toggle.expandedState === 'false');
    const currentConversationVisible = !!document.querySelector(
        '[data-tid="message-pane-list-viewport"], [data-tid="message-pane-layout"], [data-tid="channel-pane-viewport"], [data-tid="channel-pane-list-viewport"], [data-tid="compose-start-post"], [data-tid="chat-pane-compose-message-footer"], [data-tid="post-message-subheader"], [role="main"] [role="feed"], [role="main"] [role="log"]'
    );

    return {
        treeItemCount: treeItems.length,
        hasTreeItems: treeItems.length > 0,
        hasChatRailButton: !!railButton,
        chatRailSelected: elementLooksSelected(railButton),
        hasTeamsRailButton: !!teamsRailButton,
        teamsRailSelected: elementLooksSelected(teamsRailButton),
        separateChatAndTeamsView: !!railButton && !!teamsRailButton,
        hasCollapsedSidebarExpandControl: !!sidebarExpandControl,
        currentConversationVisible,
        chatSectionCollapsed: collapsedToggles.some(toggle =>
            !toggle.isTeamsSection && isChatSectionHeaderText(toggle.text || '', toggle.dataTid || '')),
        teamsSectionCollapsed: collapsedToggles.some(toggle => toggle.isTeamsSection)
    };
}

function tryRevealChatSidebar(existingDiagnostics = null, options = {}) {
    const diagnostics = existingDiagnostics || getChatSidebarDiagnostics();
    const actions = [];
    const allowChatRailSwitch = options.allowChatRailSwitch !== false;

    const sidebarExpandControl = findCollapsedSidebarExpandControl();
    const expandedHiddenSidebar = !diagnostics.hasTreeItems &&
        sidebarExpandControl &&
        clickElementBestEffort(sidebarExpandControl);
    if (expandedHiddenSidebar) {
        actions.push('expand_hidden_sidebar');
    }

    const railButton = findChatRailButton();
    if (allowChatRailSwitch &&
        !expandedHiddenSidebar &&
        !diagnostics.teamsRailSelected &&
        (!diagnostics.hasTreeItems || !diagnostics.chatRailSelected) &&
        railButton &&
        !elementLooksSelected(railButton)) {
        if (clickElementBestEffort(railButton)) {
            actions.push('open_chat_app');
        }
    }

    getSidebarSectionToggleRecords().forEach(toggle => {
        if (toggle.expandedState !== 'false') return;
        const text = toggle.text || '';
        const dataTid = toggle.dataTid || '';
        if (!isChatSectionHeaderText(text, dataTid) && !toggle.isTeamsSection) return;
        if (clickElementBestEffort(toggle.clickTarget || toggle.element)) {
            actions.push(toggle.isTeamsSection ? 'expand_teams_and_channels' : 'expand_chat_section');
        }
    });

    return {
        attempted: actions.length > 0,
        actions
    };
}

// Current Teams no longer marks the open chat in the sidebar (no aria-selected
// or active class), but the page title always names it:
// "Chat | <name> | Microsoft Teams" or "<channel> | <team> | Microsoft Teams".
function createPageTitleMatcher() {
    const titleParts = (document.title || '').split('|')
        .map(part => sanitizeDetectedChatName(part).toLowerCase())
        .filter(Boolean);
    return (value) => !!value &&
        titleParts.includes(sanitizeDetectedChatName(value).toLowerCase());
}

// Cached scan entries keep their names; only "which chat is open" can change.
function getCachedAvailableChats() {
    if (!availableChatsCache || Date.now() - availableChatsCache.at > AVAILABLE_CHATS_CACHE_TTL_MS) {
        return null;
    }
    const isNamedInPageTitle = createPageTitleMatcher();
    const cachedChats = availableChatsCache.chats.map(chat => ({
        ...chat,
        isActive: isNamedInPageTitle(chat.rawName || chat.name) &&
            (!chat.isChannel || !chat.teamName || isNamedInPageTitle(chat.teamName))
    }));
    return {
        chats: mergeDetectedChats(getAvailableChats(), cachedChats),
        diagnostics: { ...availableChatsCache.diagnostics, fromCache: true }
    };
}

function getAvailableChats() {
    const chats = [];

    // Use shared helper function to get filtered chat items (now returns objects with metadata)
    const chatItems = getFilteredChatElements();
    console.log(`getAvailableChats: Found ${chatItems.length} chat/channel items`);

    const isNamedInPageTitle = createPageTitleMatcher();

    chatItems.forEach((item, index) => {
        try {
            const el = item.element;
            const isChannel = item.isChannel;
            const teamName = item.teamName;

            // Try multiple approaches to get the chat name
            let name = '';

            // Name selectors - comprehensive list for different Teams versions
            const nameSelectors = [
                // Current Teams separates the title from preview and timestamp.
                '[id^="title-chat-list-item"]',
                '[id^="title-channel-list-item"]',
                '[id^="title-team-list-item"]',
                '.fui-TreeItemLayout__main span[role="text"]',
                // Common title patterns
                'span[class*="title"]',
                'div[class*="title"]',
                'span[data-tid*="name"]',
                'span[data-tid*="title"]',

                // Fluent UI patterns
                '[class*="fui-Text"]',
                '[class*="displayName"]',
                '[class*="conversationName"]',
                '[class*="chatName"]',

                // Channel/Team names
                '[data-tid="channel-name"]',
                '[data-tid="team-name"]',

                // Generic text holders
                'span[class*="name"]',
                'div[class*="name"]',

                // Fallback - first significant text
                'span:not(:empty)',
                'div > span'
            ];

            for (const nameSelector of nameSelectors) {
                const nameEl = el.querySelector(nameSelector);
                if (nameEl && nameEl.textContent.trim()) {
                    name = nameEl.textContent.trim();
                    // Filter out very short or likely non-name content
                    if (name.length > 1 && name.length < 100 && !name.match(/^\d+$/)) {
                        break;
                    }
                }
            }

            // If still no name, try getting first text content
            if (!name) {
                const textContent = el.textContent.trim();
                if (textContent) {
                    // Get first line or first 50 chars
                    name = textContent.split('\n')[0].substring(0, 50).trim();
                }
            }

            if (!name) {
                name = isChannel ? `Channel ${index + 1}` : `Chat ${index + 1}`;
            }

            // For channels, append team name as indicator
            let displayName = name;
            if (isChannel && teamName) {
                displayName = `${name} (${teamName})`;
            }

            // Try to get last message preview
            let preview = '';
            const previewSelectors = [
                'span[class*="preview"]',
                'div[class*="preview"]',
                'span[class*="lastMessage"]',
                'span[class*="secondary"]',
                '[class*="messagePreview"]',
                '[data-tid="message-preview"]'
            ];

            for (const previewSelector of previewSelectors) {
                const previewEl = el.querySelector(previewSelector);
                if (previewEl && previewEl.textContent.trim()) {
                    preview = previewEl.textContent.trim().substring(0, 50);
                    break;
                }
            }

            // Check if this is the currently active/selected chat
            const isActive = el.classList.contains('active') ||
                           el.classList.contains('selected') ||
                           el.getAttribute('aria-selected') === 'true' ||
                           el.getAttribute('aria-current') === 'true' ||
                           el.querySelector('[aria-selected="true"]') !== null ||
                           el.matches('[class*="active"]') ||
                           el.matches('[class*="selected"]') ||
                           (isNamedInPageTitle(name) && (!isChannel || !teamName || isNamedInPageTitle(teamName)));

            // Avoid duplicates by checking if we already have this display name
            const isDuplicate = chats.some(c => c.name === displayName);
            if (!isDuplicate) {
                const rawName = name;
                const isSupported = !isChannel || !isActive || !isUnsupportedThreadsOrPostsLayout();
                chats.push({
                    id: index,
                    name: displayName,
                    rawName: rawName,
                    preview: preview,
                    isActive: isActive,
                    isChannel: isChannel,
                    teamName: teamName,
                    key: buildChatIdentityKey({ rawName, teamName, isChannel }),
                    isSupported,
                    unsupportedReason: isSupported ? null : 'channel_posts_layout_not_supported',
                    element: null // Don't send DOM element
                });
            }
        } catch (e) {
            console.error('Error parsing chat item:', e);
        }
    });

    console.log(`Total chats/channels found: ${chats.length}`);
    return chats;
}

async function collectRenderedSidebarChats() {
    let chats = getAvailableChats();
    const initialReveal = tryRevealChatSidebar(getChatSidebarDiagnostics(), { allowChatRailSwitch: false });
    if (initialReveal.attempted) {
        await sleep(225);
        chats = mergeDetectedChats(chats, getAvailableChats());
    }
    const scrollContainer = getChatSidebarScrollContainer();
    if (!scrollContainer) {
        return { chats, scrolled: false, pagesScanned: 1 };
    }

    const originalScrollTop = scrollContainer.scrollTop;
    let pagesScanned = 0;

    try {
        scrollContainer.scrollTop = 0;
        scrollContainer.dispatchEvent(new Event('scroll', { bubbles: true }));
        await sleep(120);

        for (let page = 0; page < 12; page++) {
            pagesScanned++;
            const reveal = tryRevealChatSidebar(getChatSidebarDiagnostics(), { allowChatRailSwitch: false });
            if (reveal.attempted) {
                await sleep(180);
            }
            chats = mergeDetectedChats(chats, getAvailableChats());

            const maxScrollTop = Math.max(0, scrollContainer.scrollHeight - scrollContainer.clientHeight);
            if (scrollContainer.scrollTop >= maxScrollTop - 4) break;

            const nextScrollTop = Math.min(
                maxScrollTop,
                scrollContainer.scrollTop + Math.max(120, scrollContainer.clientHeight * 0.8)
            );
            if (nextScrollTop <= scrollContainer.scrollTop + 1) break;

            scrollContainer.scrollTop = nextScrollTop;
            scrollContainer.dispatchEvent(new Event('scroll', { bubbles: true }));
            await sleep(120);
        }
    } finally {
        scrollContainer.scrollTop = Math.min(
            originalScrollTop,
            Math.max(0, scrollContainer.scrollHeight - scrollContainer.clientHeight)
        );
        scrollContainer.dispatchEvent(new Event('scroll', { bubbles: true }));
        await sleep(80);
    }

    return { chats, scrolled: pagesScanned > 1, pagesScanned };
}

async function scanUnfilteredSidebar() {
    const originalFilterState = getConversationFilterState();
    const filtersChanged = Object.values(originalFilterState).some(Boolean);

    try {
        if (filtersChanged) {
            await applyConversationFilterState({ chat: false, channel: false, unread: false });
            await sleep(225);
        }

        const scan = await collectRenderedSidebarChats();
        return {
            ...scan,
            attempted: true,
            restored: true
        };
    } finally {
        if (filtersChanged) {
            await applyConversationFilterState(originalFilterState);
            await sleep(180);
        }
    }
}

async function scanChatsFilter() {
    const toggle = getChatFilterToggle();
    if (!toggle) {
        return { chats: [], attempted: false, restored: true, pagesScanned: 0 };
    }

    const originalFilterState = getConversationFilterState();

    try {
        await applyConversationFilterState({ chat: true, channel: false, unread: false });
        await sleep(225);

        const scan = await collectRenderedSidebarChats();
        return {
            chats: scan.chats.filter(chat => !chat.isChannel),
            attempted: true,
            restored: true,
            pagesScanned: scan.pagesScanned
        };
    } finally {
        await applyConversationFilterState(originalFilterState);
        await sleep(180);
    }
}

function getSelectedConversationRailKind(fallbackChats = []) {
    if (elementLooksSelected(findTeamsRailButton())) return 'teams';
    if (elementLooksSelected(findChatRailButton())) return 'chat';
    if (fallbackChats.some(chat => chat?.isChannel)) return 'teams';
    if (fallbackChats.some(chat => chat && !chat.isChannel)) return 'chat';
    return null;
}

async function selectConversationRail(appKind) {
    let railButton = appKind === 'teams' ? findTeamsRailButton() : findChatRailButton();
    if (!railButton) {
        return { available: false, changed: false, selected: false };
    }

    if (elementLooksSelected(railButton)) {
        return { available: true, changed: false, selected: true };
    }

    if (!clickElementBestEffort(railButton)) {
        return { available: true, changed: false, selected: false };
    }

    for (let attempt = 0; attempt < 12; attempt++) {
        await sleep(attempt === 0 ? 425 : 150);
        railButton = appKind === 'teams' ? findTeamsRailButton() : findChatRailButton();
        if (elementLooksSelected(railButton)) {
            return { available: true, changed: true, selected: true };
        }
    }

    // Some Teams rings do not expose selected state on app-bar buttons. The
    // click still changes the mounted conversation tree, so continue scanning.
    return { available: true, changed: true, selected: false };
}

async function scanConversationRail(appKind) {
    const selection = await selectConversationRail(appKind);
    if (!selection.available) {
        return { chats: [], attempted: false, pagesScanned: 0, actions: [] };
    }

    const actions = [];
    for (let attempt = 0; attempt < 5; attempt++) {
        const diagnostics = getChatSidebarDiagnostics();
        const reveal = tryRevealChatSidebar(diagnostics, { allowChatRailSwitch: false });
        actions.push(...reveal.actions);
        if (diagnostics.hasTreeItems && !reveal.attempted) break;
        await sleep(attempt === 0 ? 425 : 225);
    }

    const scan = await scanUnfilteredSidebar();
    return {
        chats: scan.chats
            .filter(chat => appKind === 'teams' ? chat.isChannel : !chat.isChannel)
            .map(chat => ({
                ...chat,
                kind: chat.kind || (chat.isChannel ? getChatKindForUi(chat) : 'chat')
            })),
        attempted: true,
        pagesScanned: scan.pagesScanned,
        actions,
        selectionChanged: selection.changed
    };
}

async function scanSeparateConversationRails(initialChats = []) {
    const chatRailButton = findChatRailButton();
    const teamsRailButton = findTeamsRailButton();
    if (!chatRailButton || !teamsRailButton) {
        return {
            chats: [],
            attempted: false,
            restored: true,
            initialRail: null,
            scans: {}
        };
    }

    const initialRail = getSelectedConversationRailKind(initialChats);
    const scanOrder = initialRail === 'teams' ? ['teams', 'chat'] : ['chat', 'teams'];
    const scans = {};
    let chats = [];
    let restored = !initialRail;

    try {
        for (const appKind of scanOrder) {
            scans[appKind] = await scanConversationRail(appKind);
            chats = mergeDetectedChats(chats, scans[appKind].chats);
        }
    } finally {
        if (initialRail) {
            const restore = await selectConversationRail(initialRail);
            restored = restore.available && (restore.selected || restore.changed);
            await sleep(225);
        }
    }

    return {
        chats,
        attempted: true,
        restored,
        initialRail,
        scans
    };
}

async function getAvailableChatsWithRecovery() {
    let chats = getAvailableChats();
    let diagnostics = getChatSidebarDiagnostics();
    const allowChatRailSwitch = chats.length === 0 &&
        !diagnostics.currentConversationVisible &&
        !diagnostics.separateChatAndTeamsView &&
        !diagnostics.teamsRailSelected;

    const shouldAttemptRecovery = chats.length === 0 ||
        !diagnostics.hasTreeItems ||
        diagnostics.chatSectionCollapsed ||
        diagnostics.teamsSectionCollapsed ||
        (allowChatRailSwitch && diagnostics.hasChatRailButton && !diagnostics.chatRailSelected);

    const recovery = shouldAttemptRecovery
        ? tryRevealChatSidebar(diagnostics, { allowChatRailSwitch })
        : { attempted: false, actions: [] };

    if (recovery.attempted) {
        for (let attempt = 0; attempt < 6; attempt++) {
            await sleep(attempt === 0 ? 425 : 275);
            chats = mergeDetectedChats(chats, getAvailableChats());
            diagnostics = getChatSidebarDiagnostics();

            if (chats.length > 0 && diagnostics.hasTreeItems &&
                !diagnostics.chatSectionCollapsed && !diagnostics.teamsSectionCollapsed) {
                break;
            }

            if (attempt === 1 || attempt === 3) {
                tryRevealChatSidebar(diagnostics, { allowChatRailSwitch });
            }
        }
    }

    diagnostics = getChatSidebarDiagnostics();
    const shouldScanCombinedSidebar = !diagnostics.separateChatAndTeamsView;
    let unfilteredScan = { chats: [], attempted: false, restored: true, pagesScanned: 0 };
    let chatFilterScan = { chats: [], attempted: false, restored: true, pagesScanned: 0 };

    // Combined view keeps chats and channels in one tree. Separate view is
    // handled below by visiting the Chat and Teams rails once each, so avoid
    // rescanning the currently mounted tree several times first.
    if (shouldScanCombinedSidebar) {
        try {
            unfilteredScan = await scanUnfilteredSidebar();
            chats = mergeDetectedChats(chats, unfilteredScan.chats);
        } catch (error) {
            console.warn('Unfiltered conversation scan failed:', error?.message || error);
            unfilteredScan = {
                chats: [],
                attempted: true,
                restored: true,
                pagesScanned: 0,
                error: error?.message || 'unknown'
            };
        }

        try {
            chatFilterScan = await scanChatsFilter();
            chats = mergeDetectedChats(chats, chatFilterScan.chats);
        } catch (error) {
            console.warn('Chats-only sidebar scan failed:', error?.message || error);
            chatFilterScan = {
                chats: [],
                attempted: true,
                restored: true,
                pagesScanned: 0,
                error: error?.message || 'unknown'
            };
        }
    }

    let separateRailScan = {
        chats: [],
        attempted: false,
        restored: true,
        initialRail: null,
        scans: {}
    };
    try {
        separateRailScan = await scanSeparateConversationRails(chats);
        chats = mergeDetectedChats(chats, separateRailScan.chats);
    } catch (error) {
        console.warn('Separate Chat/Teams rail scan failed:', error?.message || error);
        separateRailScan = {
            chats: [],
            attempted: true,
            restored: false,
            initialRail: null,
            scans: {},
            error: error?.message || 'unknown'
        };
    }

    diagnostics = getChatSidebarDiagnostics();

    return {
        chats,
        diagnostics: {
            ...diagnostics,
            preservedCurrentConversation: diagnostics.currentConversationVisible && !recovery.actions.includes('open_chat_app'),
            recoveryAttempted: recovery.attempted,
            recoverySucceeded: recovery.attempted && chats.length > 0 && diagnostics.hasTreeItems &&
                !diagnostics.chatSectionCollapsed && !diagnostics.teamsSectionCollapsed,
            recoveryActions: recovery.actions,
            unfilteredScanAttempted: unfilteredScan.attempted,
            unfilteredScanRestored: unfilteredScan.restored,
            unfilteredScanPagesScanned: unfilteredScan.pagesScanned,
            unfilteredScanError: unfilteredScan.error || null,
            chatsFilterScanAttempted: chatFilterScan.attempted,
            chatsFilterRestored: chatFilterScan.restored,
            chatsFilterPagesScanned: chatFilterScan.pagesScanned,
            chatsFoundViaFilter: chatFilterScan.chats.length,
            chatsFilterError: chatFilterScan.error || null,
            separateRailScanAttempted: separateRailScan.attempted,
            separateRailScanRestored: separateRailScan.restored,
            separateRailInitialView: separateRailScan.initialRail,
            chatsFoundViaChatRail: separateRailScan.scans.chat?.chats?.length || 0,
            channelsFoundViaTeamsRail: separateRailScan.scans.teams?.chats?.length || 0,
            separateRailScanError: separateRailScan.error || null
        }
    };
}

function getChatKindForUi(chat) {
    if (!chat) return 'unknown';
    if (!chat.isChannel) return 'chat';
    if (chat.isActive) {
        const activeTabText = normalizeChatIdentityPart(document.querySelector('[role="tab"][aria-selected="true"]')?.textContent || '');
        const hasPostsComposer = !!document.querySelector('[data-tid="compose-start-post"], [data-tid="quick-reply-placeholder-reply"], [data-tid="subject-line"], [data-tid="post-message-subheader"]');
        const hasChatPane = !!document.querySelector('[data-tid="message-pane-list-viewport"], [data-tid="message-pane-layout"], [data-tid="chat-pane-compose-message-footer"]');
        if ((activeTabText && (activeTabText.startsWith('posts') || activeTabText.startsWith('publicaciones'))) ||
            (hasPostsComposer && !hasChatPane)) {
            return 'channel_posts';
        }
        return 'channel_chat';
    }
    return 'channel';
}

function findMatchingChatIndex(liveChats, chatTarget) {
    if (!chatTarget || typeof chatTarget !== 'object') return -1;

    const targetKey = chatTarget.key || buildChatIdentityKey(chatTarget);
    const matchedChat = liveChats.find(chat => chat.key === targetKey);
    // Discovery deduplicates rows; id retains the index in the live DOM list.
    if (matchedChat) return matchedChat.id;

    const targetName = normalizeChatIdentityPart(chatTarget.rawName || chatTarget.name);
    const targetTeam = normalizeChatIdentityPart(chatTarget.teamName);
    const fallbackChat = liveChats.find(chat =>
        normalizeChatIdentityPart(chat.rawName || chat.name) === targetName &&
        normalizeChatIdentityPart(chat.teamName) === targetTeam &&
        !!chat.isChannel === !!chatTarget.isChannel
    );
    return fallbackChat ? fallbackChat.id : -1;
}

async function revealChatTargetForBackup(chatTarget) {
    if (!chatTarget) return null;

    const initialRail = getSelectedConversationRailKind(getAvailableChats());
    const desiredRail = chatTarget.isChannel ? 'teams' : 'chat';
    const desiredRailButton = desiredRail === 'teams' ? findTeamsRailButton() : findChatRailButton();
    let railChanged = false;

    if (desiredRailButton && !elementLooksSelected(desiredRailButton)) {
        const selection = await selectConversationRail(desiredRail);
        railChanged = selection.changed;
    }

    for (let attempt = 0; attempt < 5; attempt++) {
        const diagnostics = getChatSidebarDiagnostics();
        const reveal = tryRevealChatSidebar(diagnostics, { allowChatRailSwitch: false });
        if (diagnostics.hasTreeItems && !reveal.attempted) break;
        await sleep(attempt === 0 ? 425 : 225);
    }

    const originalFilterState = getConversationFilterState();
    const targetFilterState = chatTarget.isChannel
        ? { chat: false, channel: false, unread: false }
        : { chat: true, channel: false, unread: false };
    await applyConversationFilterState(targetFilterState);
    await sleep(225);

    const scrollContainer = getChatSidebarScrollContainer();
    const originalScrollTop = scrollContainer?.scrollTop || 0;
    if (scrollContainer) {
        scrollContainer.scrollTop = 0;
        scrollContainer.dispatchEvent(new Event('scroll', { bubbles: true }));
        await sleep(120);
    }

    const discoveryState = { originalFilterState, scrollContainer, originalScrollTop };
    const maxPages = scrollContainer ? 12 : 1;

    for (let page = 0; page < maxPages; page++) {
        const liveChats = getAvailableChats();
        const matchedIndex = findMatchingChatIndex(liveChats, chatTarget);
        if (matchedIndex !== -1) {
            return { matchedIndex, discoveryState };
        }

        if (!scrollContainer) break;
        const maxScrollTop = Math.max(0, scrollContainer.scrollHeight - scrollContainer.clientHeight);
        if (scrollContainer.scrollTop >= maxScrollTop - 4) break;

        const nextScrollTop = Math.min(
            maxScrollTop,
            scrollContainer.scrollTop + Math.max(120, scrollContainer.clientHeight * 0.8)
        );
        if (nextScrollTop <= scrollContainer.scrollTop + 1) break;

        scrollContainer.scrollTop = nextScrollTop;
        scrollContainer.dispatchEvent(new Event('scroll', { bubbles: true }));
        await sleep(120);
    }

    await restoreChatDiscoveryState(discoveryState);
    if (railChanged && initialRail && initialRail !== desiredRail) {
        await selectConversationRail(initialRail);
    }
    return null;
}

async function restoreChatDiscoveryState(discoveryState) {
    if (!discoveryState) return;

    const { originalFilterState, scrollContainer, originalScrollTop } = discoveryState;
    if (scrollContainer?.isConnected) {
        scrollContainer.scrollTop = Math.min(
            originalScrollTop,
            Math.max(0, scrollContainer.scrollHeight - scrollContainer.clientHeight)
        );
        scrollContainer.dispatchEvent(new Event('scroll', { bubbles: true }));
    }

    if (originalFilterState) {
        await applyConversationFilterState(originalFilterState);
        await sleep(180);
    }
}

// Function to click on a specific chat to open it
async function openChat(chatIndex, chatTarget = null) {
    let resolvedIndex = chatIndex;
    let discoveryState = null;

    // Resolve by stable key/name/team first to avoid index drift in Teams sidebar.
    if (chatTarget && typeof chatTarget === 'object') {
        const liveChats = getAvailableChats();
        let matchedIndex = findMatchingChatIndex(liveChats, chatTarget);

        if (matchedIndex === -1) {
            const revealedTarget = await revealChatTargetForBackup(chatTarget);
            if (revealedTarget) {
                matchedIndex = revealedTarget.matchedIndex;
                discoveryState = revealedTarget.discoveryState;
            }
        }

        if (matchedIndex !== -1) {
            resolvedIndex = matchedIndex;
            console.log(`openChat: Resolved index ${chatIndex} -> ${resolvedIndex} using stable chat identity`);
        } else {
            console.warn('openChat: Could not resolve chat by stable identity', chatTarget);
            return false;
        }
    }

    // Get filtered chat items using the same logic as getAvailableChats
    const chatItems = getFilteredChatElements();

    console.log(`openChat: Found ${chatItems.length} chats/channels, trying to open index ${resolvedIndex}`);

    if (chatItems.length > resolvedIndex && resolvedIndex >= 0) {
        const chatItem = chatItems[resolvedIndex];
        const chatElement = chatItem.element;
        const isChannel = chatItem.isChannel;
        console.log(`openChat: Clicking on ${isChannel ? 'channel' : 'chat'}: ${chatElement.textContent?.substring(0, 50)}...`);

        // Click the chat/channel element
        if (!clickElementBestEffort(chatElement)) {
            await restoreChatDiscoveryState(discoveryState);
            return false;
        }

        // A hidden chat may have required Teams' Chats-only filter and a
        // virtual-list scroll. Restore the user's sidebar immediately after
        // opening the target conversation.
        await restoreChatDiscoveryState(discoveryState);

        // Invalidate container cache since we switched chats
        invalidateContainerCache();

        // Wait for content to actually load instead of fixed delay
        const maxWaitMs = isChannel ? 8000 : 6000;
        const pollInterval = 500;
        const startTime = Date.now();
        let contentLoaded = false;

        while (Date.now() - startTime < maxWaitMs) {
            await sleep(pollInterval);

            // Check for content indicators
            const hasContainer = !!document.querySelector(
                '[data-tid="message-pane-list-viewport"], [data-tid="message-pane-layout"], [data-tid="channel-pane-viewport"], [data-tid="channel-pane-list-viewport"], [data-tid*="post-list"], [role="feed"], [role="log"]'
            );
            const hasMessages = document.querySelectorAll(
                '[data-tid="chat-pane-item"], [data-tid*="message-pane-item"], [data-tid*="post-item"], [data-tid*="reply-item"], [data-tid*="thread-item"], [class*="fui-unstable-ChatItem"], [class*="fui-ChatMessage"], [class*="fui-ChatMyMessage"], [role="article"]'
            ).length > 0;
            const hasEmptyState = !!document.querySelector('[data-tid="channel-empty-state"]');
            const hasCompose = !!document.querySelector(
                '[data-tid="chat-pane-compose-message-footer"], [data-tid="compose-start-post"]'
            );

            if (hasContainer && (hasMessages || hasEmptyState || hasCompose)) {
                contentLoaded = true;
                console.log(`openChat: Content loaded after ${Date.now() - startTime}ms (messages: ${hasMessages}, empty: ${hasEmptyState})`);
                break;
            }
        }

        if (!contentLoaded) {
            // Fallback: wait a bit more and proceed anyway
            console.warn(`openChat: Content not detected after ${maxWaitMs}ms, proceeding anyway`);
            await sleep(2000);
        }

        // Re-invalidate cache after content load
        invalidateContainerCache();

        // Verify the chat opened by checking if it's now selected
        const isSelected = chatElement.getAttribute('aria-selected') === 'true' ||
                          chatElement.classList.contains('selected') ||
                          chatElement.hasAttribute('data-is-focusable');

        console.log(`openChat: ${isChannel ? 'Channel' : 'Chat'} click completed, selected: ${isSelected}`);
        return true;
    }

    await restoreChatDiscoveryState(discoveryState);
    console.log(`openChat: Could not find chat/channel at index ${resolvedIndex}`);
    return false;
}

function getTreeItemOwnName(item) {
    if (!item) return null;
    const spans = item.querySelectorAll('span');
    for (const span of spans) {
        if (span.closest('[role="treeitem"]') !== item) continue;
        if (typeof span.className === 'string' && span.className.includes('Avatar')) continue;
        const spanText = span.textContent.trim();
        if (spanText && spanText.length > 1) return spanText;
    }

    const fallback = (item.textContent || '').trim().split('\n')[0].trim();
    return fallback ? fallback.substring(0, 60) : null;
}

function looksLikeGenericTeamsSection(text, dataTid = '') {
    const normalizedText = normalizeChatIdentityPart(text);
    const normalizedDataTid = normalizeChatIdentityPart(dataTid);
    return normalizedText === 'teams' ||
        normalizedText.startsWith('teams and channels') ||
        normalizedText.startsWith('teams & channels') ||
        normalizedText.startsWith('channels') ||
        normalizedText.startsWith('equipos y canales') ||
        normalizedText.startsWith('canales') ||
        normalizedDataTid.includes('group-item') ||
        normalizedDataTid.includes('section-header');
}

function isNonConversationTreeItem(item) {
    const testId = item.getAttribute('data-testid') || '';
    const labelledBy = item.getAttribute('aria-labelledby') || '';
    return testId.startsWith('slice-list-item-') ||
        labelledBy.includes('title-slice-list-item_') ||
        labelledBy.includes('engage-community-type-name') ||
        labelledBy.includes('title-engage-community-list-item_');
}

// Helper function to get filtered chat elements (shared between getAvailableChats and openChat)
// Returns array of objects: {element, isChannel, teamName}
function getFilteredChatElements() {
    let chatItems = []; // Array of {element, isChannel, teamName}

    // Method 1: Sequential scan through all treeitems in the tree
    // Personal Teams (teams.live.com) has a flat structure where level 1 headers and level 2 items are siblings
    // Business Teams (teams.microsoft.com) may have nested structure
    const allTreeitems = document.querySelectorAll('[role="tree"] [role="treeitem"]');
    const hasStructuredTreeItems = allTreeitems.length > 0;
    const separateTeamsViewActive = elementLooksSelected(findTeamsRailButton()) && !!findChatRailButton();

    if (allTreeitems.length > 0) {
        let inChatSection = false;
        let inTeamsSection = false;
        let currentTeamName = null;

        for (let itemIndex = 0; itemIndex < allTreeitems.length; itemIndex++) {
            const item = allTreeitems[itemIndex];
            const level = item.getAttribute('aria-level');
            const text = (item.textContent || '').toLowerCase();
            const dataTid = (item.getAttribute('data-tid') || '').toLowerCase();

            if (level === '1') {
                // Quick views and Communities are navigation groups, not teams.
                // Inspect their row markers so this also works in translated UI
                // and in flat trees where rows are siblings of the header.
                let nonConversationSection = false;
                for (let childIndex = itemIndex + 1; childIndex < allTreeitems.length; childIndex++) {
                    const child = allTreeitems[childIndex];
                    if (child.getAttribute('aria-level') === '1') break;
                    if (isNonConversationTreeItem(child)) nonConversationSection = true;
                }
                const isChatSection = isChatSectionHeaderText(text, dataTid);
                const isTeamsSection = !nonConversationSection && (isTeamsSectionHeaderText(text, dataTid) ||
                    (separateTeamsViewActive &&
                        item.getAttribute('aria-expanded') !== null &&
                        !isChatSection));

                inChatSection = !nonConversationSection && isChatSection && !isTeamsSection;
                inTeamsSection = isTeamsSection;
                currentTeamName = null; // Reset team name when entering new section
                if (inTeamsSection &&
                    item.getAttribute('aria-expanded') !== null &&
                    !looksLikeGenericTeamsSection(text, dataTid)) {
                    // Some separate-view tenants use the team itself as the
                    // level-1 node and render its channels at level 2.
                    currentTeamName = getTreeItemOwnName(item);
                }
            } else if (level === '2') {
                if (inChatSection) {
                    // Regular chat
                    chatItems.push({element: item, isChannel: false, teamName: null});
                } else if (inTeamsSection) {
                    // This is a team - check if it has children (channels)
                    const hasChildren = item.querySelectorAll('[role="treeitem"]').length > 0;
                    const isExpandable = item.getAttribute('aria-expanded') !== null;
                    if (hasChildren || isExpandable) {
                        // Extract team name - find first non-empty span that's not inside a child treeitem
                        // and not an avatar (avatars usually have fui-Avatar class or are empty)
                        const allSpans = item.querySelectorAll('span');
                        let teamName = null;
                        for (const span of allSpans) {
                            // Skip if this span is inside a child treeitem (channel)
                            if (span.closest('[role="treeitem"]') !== item) continue;
                            // Skip avatar spans
                            if (span.className.includes('Avatar')) continue;
                            const spanText = span.textContent.trim();
                            // Skip empty spans or very short content
                            if (spanText && spanText.length > 1) {
                                teamName = spanText;
                                break;
                            }
                        }
                        currentTeamName = teamName || item.textContent.trim().split('\n')[0].substring(0, 30);
                    } else {
                        // Teams' "Channels only" and some separate-view tenants
                        // render channels directly at level 2, without a team
                        // wrapper. Treat non-navigation leaves as channels.
                        if (!looksLikeSeeAllNavigation(text) &&
                            !looksLikeNewChatNavigation(text, dataTid)) {
                            chatItems.push({element: item, isChannel: true, teamName: currentTeamName});
                        }
                    }
                }
            } else if (level === '3' && inTeamsSection && currentTeamName) {
                // This is a channel under a team
                const channelText = (item.textContent || '').toLowerCase();
                // Skip "See all channels" and similar navigation items
                if (!looksLikeSeeAllNavigation(channelText)) {
                    chatItems.push({element: item, isChannel: true, teamName: currentTeamName});
                }
            }
        }
    }

    // Method 2: Try nested structure (for Business Teams where level 2 might be inside level 1)
    if (chatItems.length === 0) {
        const allLevel1Sections = document.querySelectorAll('[role="treeitem"][aria-level="1"]');

        for (const section of allLevel1Sections) {
            const sectionText = (section.textContent || '').toLowerCase();
            const dataTid = (section.getAttribute('data-tid') || '').toLowerCase();

            const isChatSection = isChatSectionHeaderText(sectionText, dataTid);
            const isTeamsSection = isTeamsSectionHeaderText(sectionText, dataTid);

            if (isChatSection && !isTeamsSection) {
                const sectionChats = section.querySelectorAll('[role="treeitem"][aria-level="2"]');
                sectionChats.forEach(el => chatItems.push({element: el, isChannel: false, teamName: null}));
            } else if (isTeamsSection) {
                // Also check for channels in Teams section (nested structure)
                const teams = section.querySelectorAll('[role="treeitem"][aria-level="2"]');
                teams.forEach(team => {
                    const hasChildren = team.querySelectorAll('[role="treeitem"]').length > 0;
                    const isExpandable = team.getAttribute('aria-expanded') !== null;
                    if (hasChildren || isExpandable) {
                        // Extract team name - find first non-empty span not inside a child treeitem
                        const allSpans = team.querySelectorAll('span');
                        let teamName = null;
                        for (const span of allSpans) {
                            if (span.closest('[role="treeitem"]') !== team) continue;
                            if (span.className.includes('Avatar')) continue;
                            const spanText = span.textContent.trim();
                            if (spanText && spanText.length > 1) {
                                teamName = spanText;
                                break;
                            }
                        }
                        teamName = teamName || team.textContent.trim().split('\n')[0].substring(0, 30);
                        const channels = team.querySelectorAll('[role="treeitem"][aria-level="3"]');
                        channels.forEach(channel => {
                            const channelText = (channel.textContent || '').toLowerCase();
                            if (!looksLikeSeeAllNavigation(channelText)) {
                                chatItems.push({element: channel, isChannel: true, teamName: teamName});
                            }
                        });
                    } else {
                        const channelText = (team.textContent || '').toLowerCase();
                        const channelTid = (team.getAttribute('data-tid') || '').toLowerCase();
                        if (!looksLikeSeeAllNavigation(channelText) &&
                            !looksLikeNewChatNavigation(channelText, channelTid)) {
                            chatItems.push({element: team, isChannel: true, teamName: null});
                        }
                    }
                });
            }
        }
    }

    // Method 3: Fallback selectors for various Teams versions
    if (chatItems.length === 0 && !hasStructuredTreeItems) {
        const fallbackSelectors = [
            // Modern Teams (2024+)
            '#chat-pane-list [role="treeitem"][aria-level="2"]',
            'div[role="treeitem"][data-tid^="chat-list-item"]',
            'div[data-tid="chat-list-group-item-PINNED"] [role="treeitem"]',
            'div[data-tid="chat-list-group-item-RECENTS"] [role="treeitem"]',
            'div[data-tid="chat-list-item"]',
            // Legacy Teams
            '[data-tid="chat-pane"] [role="listitem"]',
            '[data-tid="convList"] [role="listitem"]',
            '[data-tid="conversation-list"] [role="listitem"]',
            // Classic Teams
            'div[class*="chatListItem"]',
            'div[class*="ChatListItem"]',
            'div[class*="conversationListItem"]',
            // Fluent UI based
            'div[class*="fui-ListItem"]',
            '[class*="fui-TreeItem"]',
            // List-based layouts
            '[role="list"] [role="listitem"]',
            '[role="listbox"] [role="option"]',
            // Grid-based layouts (some older versions)
            '[role="grid"] [role="row"]',
            // Generic clickable chat items
            '[data-tid*="chat"][data-tid*="item"]',
            '[data-testid*="chat"][data-testid*="item"]'
        ];

        for (const selector of fallbackSelectors) {
            try {
                const elements = document.querySelectorAll(selector);
                if (elements.length > 0) {
                    chatItems = Array.from(elements).map(el => ({element: el, isChannel: false, teamName: null}));
                    console.log(`getFilteredChatElements: Found ${elements.length} elements via fallback selector: ${selector}`);
                    break;
                }
            } catch (e) {
                // Invalid selector, skip
            }
        }
    }

    // Method 4: Last resort - get all level 2 items with smart filtering
    if (chatItems.length === 0 && !hasStructuredTreeItems) {
        const allLevel2 = document.querySelectorAll('[role="treeitem"][aria-level="2"]');
        const filteredElements = Array.from(allLevel2).filter(el => {
            const text = (el.textContent || '').toLowerCase();
            const dataTid = (el.getAttribute('data-tid') || '').toLowerCase();
            // Exclude teams/channels/sections
            if (text.includes('see all') ||
                dataTid.includes('team') ||
                dataTid.includes('channel') ||
                dataTid.includes('group-item')) {
                return false;
            }
            return true;
        });
        chatItems = filteredElements.map(el => ({element: el, isChannel: false, teamName: null}));
    }

    // Method 5: Ultimate fallback - any interactive list items that look like chats
    if (chatItems.length === 0 && !hasStructuredTreeItems) {
        const interactiveItems = document.querySelectorAll('[role="listitem"], [role="option"], [role="row"]');
        const filteredItems = Array.from(interactiveItems).filter(el => {
            const parent = el.closest('[data-tid*="chat"], [class*="chat"], [id*="chat"]');
            return parent !== null;
        });
        chatItems = filteredItems.map(el => ({element: el, isChannel: false, teamName: null}));
        if (chatItems.length > 0) {
            console.log(`getFilteredChatElements: Found ${chatItems.length} elements via ultimate fallback`);
        }
    }

    // Filter out non-chat items
    chatItems = chatItems.filter(item => {
        const el = item.element;
        const text = (el.textContent || '').toLowerCase();
        const dataTid = (el.getAttribute('data-tid') || '').toLowerCase();
        // Both Business and Personal Teams use aria-describedby for real chats
        const hasContextMenu = el.getAttribute('aria-describedby') || el.getAttribute('aria-description');

        if (isNonConversationTreeItem(el)) return false;
        // Exclude group headers
        if (dataTid.includes('group-item')) return false;
        // Exclude level 1 items
        if (el.getAttribute('aria-level') === '1') return false;
        // Exclude "See all your teams" and similar
        if (looksLikeSeeAllNavigation(text)) return false;
        // Teams renders "New message" inside the Chats section as a navigation
        // action. It is not a conversation and must never appear as a backup target.
        if (!item.isChannel && !hasContextMenu && looksLikeNewChatNavigation(text, dataTid)) return false;
        // Exclude suggestions - they don't have a context menu
        // Business Teams: aria-describedby="simple-collab-item-context-menu"
        // Personal Teams: aria-describedby="chat-list-context-menu-instruction-message"
        // Note: Some older versions might not have this, so only filter if we have enough items
        // But don't filter channels this way as they may not have context menu
        if (!hasContextMenu && chatItems.length > 5 && !item.isChannel) return false;
        // Exclude teams (expandable level 2 items with children are teams) - but not if already marked as channel
        if (!item.isChannel && el.getAttribute('aria-expanded') !== null && el.querySelectorAll('[role="treeitem"]').length > 0) return false;

        return true;
    });

    console.log(`getFilteredChatElements: Found ${chatItems.length} actual chats/channels`);
    return chatItems;
}

function sanitizeDetectedChatName(name) {
    let value = (name || '')
        .replace(/\u00a0/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();

    // Group headers can include participant counts like "Team Chat (5)".
    value = value.replace(/\s*\(\d+\)\s*$/, '');
    // Strip transient presence status suffixes.
    value = value.replace(/\s*(?:[-|]\s*)?(Offline|Online|Away|Busy|Do Not Disturb|Be Right Back|Appear Offline|Out of Office)\s*$/i, '');

    return value.trim();
}

// Function to extract chat name from the Teams header
function getChatName() {
    // First, try to get from page title - most reliable in Teams 2024+
    // Format: "Chat | Person Name | Microsoft Teams" or "ChannelName | TeamName | Microsoft Teams"
    const pageTitle = document.title;
    if (pageTitle && pageTitle.includes('|')) {
        const parts = pageTitle.split('|').map(p => p.trim());
        // The chat/person/channel name is typically in the first or second part
        // For channels: "ChannelName | TeamName | Microsoft Teams"
        // For chats: "Chat | Person Name | Microsoft Teams"
        if (parts.length >= 2) {
            // If first part is not "Chat" or "Teams", use it (likely channel name)
            if (parts[0] && parts[0] !== 'Chat' && parts[0] !== 'Teams' && parts[0] !== 'Microsoft Teams') {
                // For channels, combine channel name and team name
                if (parts[1] && parts[1] !== 'Microsoft Teams') {
                    return sanitizeDetectedChatName(`${parts[0]} (${parts[1]})`);
                }
                return sanitizeDetectedChatName(parts[0]);
            }
            // Otherwise use second part (person/group name for chats)
            if (parts[1] && parts[1] !== 'Microsoft Teams') {
                return sanitizeDetectedChatName(parts[1]);
            }
        }
    }

    // Try button with haspopup="dialog" which contains chat name in Teams 2024+
    const chatHeaderBtn = document.querySelector('button[haspopup="dialog"]');
    if (chatHeaderBtn) {
        let text = chatHeaderBtn.textContent.trim();
        // Remove status indicators like "Offline", "Online", "Away", "Busy", etc.
        text = text.replace(/\s*(Offline|Online|Away|Busy|Do Not Disturb|Be Right Back|Appear Offline|Out of Office)\s*$/i, '').trim();
        if (text && text !== 'Chat') {
            return sanitizeDetectedChatName(text);
        }
    }

    // Channel-specific selectors
    const channelSelectors = [
        '[data-tid="channel-header-title"]',
        '[data-tid="channel-name"]',
        'h1[data-tid="channel-header"]',
        '[class*="channelName"]',
        '[class*="channel-header"] h1',
        '[class*="channel-header"] span[class*="title"]'
    ];

    for (let selector of channelSelectors) {
        const element = document.querySelector(selector);
        if (element && element.textContent.trim()) {
            return sanitizeDetectedChatName(element.textContent.trim());
        }
    }

    // Legacy chat selectors for older Teams versions
    const selectors = [
        'h1[data-tid="messagesHeader-Title"]',
        'span[data-tid="chat-header-title"]',
        'div[class*="chatHeader"] h1',
        'div[class*="threadHeader"] span'
    ];

    for (let selector of selectors) {
        const element = document.querySelector(selector);
        if (element && element.textContent.trim() && element.textContent.trim() !== 'Chat') {
            return sanitizeDetectedChatName(element.textContent.trim());
        }
    }

    return 'Teams Chat Backup';
}

function isUnsupportedThreadsOrPostsLayout() {
    const threadLayoutHints = [
        '[data-tid*="thread-layout"]',
        '[data-tid*="threads-layout"]',
        '[class*="threadLayout"]',
        '[class*="threads-layout"]'
    ];
    const postsComposerHints = [
        '[data-tid*="start-post"]',
        '[data-tid*="new-post"]',
        '[data-tid*="post-compose"]',
        '[data-tid*="reply-compose"]'
    ];
    const chatComposerHints = [
        '[data-tid*="message-compose"]',
        '[data-tid*="message-composer"]',
        '[data-tid*="compose-box"]',
        '[data-tid*="ckeditor"]',
        '[data-tid*="rich-text-editor"]'
    ];

    let hasStrongThreadLayoutHint = false;
    for (const selector of threadLayoutHints) {
        try {
            if (document.querySelector(selector)) {
                hasStrongThreadLayoutHint = true;
                break;
            }
        } catch {
            // Ignore invalid selectors/hints
        }
    }

    const activeTabText = (document.querySelector('[role="tab"][aria-selected="true"]')?.textContent || '').toLowerCase();
    const mainText = (document.querySelector('[role="main"]')?.textContent || '').toLowerCase();
    const composerCandidates = document.querySelectorAll('textarea, input[type="text"], [role="textbox"], [contenteditable="true"], [data-tid*="compose"], [data-tid*="editor"]');
    const composerHintsText = Array.from(composerCandidates)
        .map(el => [
            el.getAttribute('placeholder') || '',
            el.getAttribute('aria-label') || '',
            el.getAttribute('data-placeholder') || '',
            el.textContent || ''
        ].join(' '))
        .map(text => text.trim().toLowerCase())
        .filter(Boolean)
        .join(' | ');

    const hasPostComposerHint = postsComposerHints.some(selector => {
        try {
            return !!document.querySelector(selector);
        } catch {
            return false;
        }
    }) || /(start a post|new post|reply to this post|reply to post|write a reply)/i.test(composerHintsText);

    const hasChatComposerHint = chatComposerHints.some(selector => {
        try {
            return !!document.querySelector(selector);
        } catch {
            return false;
        }
    }) || /(message in|type a message|write a message|send a message)/i.test(composerHintsText);

    const mayBePostsOrThreadsLayout = activeTabText.includes('posts') ||
        hasStrongThreadLayoutHint ||
        hasPostComposerHint ||
        mainText.includes('threads layout') ||
        mainText.includes('threads work');
    const hasExtractableFeed = mayBePostsOrThreadsLayout ? hasExtractableVisibleFeed() : false;

    if (hasExtractableFeed) {
        console.log('Allowing thread/post layout because a readable message feed is visible.');
        return false;
    }

    // A visible message-style composer is a stronger signal than the informational banner.
    if (hasChatComposerHint && !hasPostComposerHint) {
        return false;
    }

    if (activeTabText.includes('posts')) {
        return true;
    }

    if (hasStrongThreadLayoutHint && !hasChatComposerHint) {
        return true;
    }

    if ((mainText.includes('threads layout') || mainText.includes('threads work')) && hasPostComposerHint) {
        return true;
    }

    return false;
}

// Function to extract unique participants from messages
function extractParticipants(messages) {
    const participants = new Set();
    messages.forEach(msg => {
        if (msg.author && msg.author !== 'Unknown Author') {
            participants.add(msg.author);
        }
    });
    return Array.from(participants);
}

// Function to get the chat container with multiple selectors for robustness
// Supports: Chats, Channels, Personal Teams, Business Teams, different versions
// useCache: set to false to bypass cache (e.g., for validation checks)
function getChatContainer(useCache = true) {
    // Check cache first (valid for 5 seconds)
    const now = Date.now();
    if (useCache && cachedChatContainer && (now - containerCacheTimestamp) < CONTAINER_CACHE_TTL) {
        // Verify cached element is still in DOM
        if (document.contains(cachedChatContainer)) {
            return cachedChatContainer;
        }
        // Cache is stale, element was removed
        invalidateContainerCache();
    }

    const selectors = [
        // Primary chat message pane selectors
        'div[data-tid="message-pane-list-viewport"]',
        'div[data-tid="messageListContainer"]',
        'div[data-tid="chat-pane-list"]',
        'div[data-tid*="message-pane-list"]',
        'div[data-tid*="message-list"]',

        // Channel-specific selectors
        'div[data-tid="channel-pane-list-viewport"]',
        'div[data-tid="channel-pane-viewport"]',
        'div[data-tid="channel-message-list"]',
        'div[data-tid="posts-list"]',
        '[data-tid="post-list-viewport"]',
        '[data-tid*="channel-pane-list"]',
        '[data-tid*="channel-pane"]',
        '[data-tid*="post-list"]',

        // Fluent UI message containers
        'div[class*="fui-MessageList"]',
        'div[class*="ts-message-list"]',
        'div[class*="message-list"]',

        // Generic scrollable containers in main area
        'div[role="main"] div.scrollable',
        'div[role="main"] [class*="scrollable"]',
        'div[role="main"] [class*="scroll-container"]',

        // Teams-specific content areas
        '[class*="ts-messages-container"]',
        '[class*="chat-body"]',
        '[class*="conversation-body"]',
        '[class*="channel-body"]',

        // Virtualized list containers (Teams uses virtualization)
        '[class*="virtualized"]',
        '[class*="virtual-list"]',

        // Role-based selectors
        'div[role="log"]',
        'div[role="feed"]',
        'div[role="list"]',

        // Broader fallbacks - look for scrollable containers with messages
        'div[role="main"] div[style*="overflow"]',
    ];

    const mainArea = document.querySelector('div[role="main"]') ||
                     document.querySelector('[role="main"]') ||
                     document.querySelector('[data-tid="app-layout-area--main"]');
    const seen = new Set();
    const candidates = [];
    const strongerContainerSelectors = [
        '[data-tid="message-pane-list-viewport"]',
        '[data-tid*="message-pane-list"]',
        '[data-tid="channel-pane-viewport"]',
        '[data-tid*="channel-pane-list"]',
        '[data-tid*="post-list"]',
        '[class*="virtualized"]',
        '[class*="virtual-list"]',
        '[role="feed"]',
        '[role="log"]'
    ];
    const strongerContainerSelector = strongerContainerSelectors.join(', ');
    const itemSelectors = [
        '[data-tid*="chat-pane-item"]',
        '[data-tid*="message-pane-item"]',
        '[data-tid*="post-item"]',
        '[data-tid*="reply-item"]',
        '[data-tid*="thread-item"]',
        '[role="article"]',
        '[role="listitem"]'
    ];
    const itemSelector = itemSelectors.join(', ');

    const countDirectMatches = (element, selectors) => {
        try {
            return selectors.reduce((total, selector) => {
                return total + element.querySelectorAll(`:scope > ${selector}`).length;
            }, 0);
        } catch {
            return 0;
        }
    };

    const scoreContainer = (element, source, selectorHint = null) => {
        if (!element || !document.contains(element)) return -Infinity;

        const rect = element.getBoundingClientRect();
        if (rect.width < 120 || rect.height < 120) return -Infinity;

        const style = window.getComputedStyle(element);
        if (style.display === 'none' || style.visibility === 'hidden' || style.opacity === '0') {
            return -Infinity;
        }

        const scrollRoom = Math.max(0, element.scrollHeight - element.clientHeight);
        const role = (element.getAttribute('role') || '').toLowerCase();
        const dataTid = (element.getAttribute('data-tid') || '').toLowerCase();
        const className = typeof element.className === 'string' ? element.className.toLowerCase() : '';
        const inMain = !!element.closest('[role="main"], [data-tid="app-layout-area--main"]');
        const nestedStrongerContainer = element.matches(strongerContainerSelector)
            ? null
            : element.querySelector(strongerContainerSelector);
        const totalItemCount = element.querySelectorAll(itemSelector).length;
        const directItemCount = countDirectMatches(element, itemSelectors);
        let depth = 0;
        let parent = element.parentElement;
        while (parent && parent !== mainArea && depth < 12) {
            depth++;
            parent = parent.parentElement;
        }

        // Fast prefilter: ignore tiny, non-scrollable containers without any message hints.
        const hasScrollSignals = scrollRoom > 20 || style.overflowY === 'auto' || style.overflowY === 'scroll';
        const hasMessageHints = !!element.querySelector(
            '[data-tid*="chat-pane-item"], [data-tid*="message-pane-item"], [data-tid*="post-item"], [data-tid*="reply-item"], [data-tid*="thread-item"], [data-tid="message-author-name"], time, [role="heading"], [class*="fui-ChatMessage"], [class*="fui-ChatMyMessage"], [class*="fui-ChatControlMessageItem"]'
        );
        if (!hasScrollSignals && !hasMessageHints) return -Infinity;
        if (!hasScrollSignals && nestedStrongerContainer) return -Infinity;

        let score = 0;
        score += Math.min(scrollRoom, 5000);
        if (inMain) score += 2000;
        if (mainArea && (element.closest('[role="main"]') === mainArea || element.closest('[data-tid="app-layout-area--main"]') === mainArea)) score += 1200;
        if (source === 'selector') score += 700;
        if (source === 'main_descendant') score += 500;
        if (source === 'cache_neighbor') score += 900;
        if (role === 'feed' || role === 'log' || role === 'list') score += 800;
        if (dataTid.includes('message') || dataTid.includes('chat-pane') || dataTid.includes('post')) score += 1200;
        if (className.includes('virtual')) score += 600;
        if (selectorHint && (selectorHint.includes('viewport') || selectorHint.includes('list'))) score += 900;
        if (rect.top < window.innerHeight && rect.bottom > 0) score += 350;
        if (rect.left < window.innerWidth && rect.right > 0) score += 250;
        if (element.clientHeight > 260) score += 300;
        if (style.overflowY === 'auto' || style.overflowY === 'scroll') score += 350;
        score += Math.min(totalItemCount, 30) * 80;
        score += Math.min(directItemCount, 8) * 220;
        score += Math.min(depth, 6) * 100;

        if (hasMessageHints) {
            score += 2800;
        }

        if (nestedStrongerContainer) {
            score -= 3200;
        }
        if (!hasScrollSignals) {
            score -= 900;
        }
        if (dataTid.includes('layout')) {
            score -= 3600;
        }

        // Penalize huge wrappers that mostly host the app shell.
        if (role === 'main' || element === document.body || element === document.documentElement) {
            score -= 1800;
        }

        return score;
    };

    const addCandidate = (element, source, selectorHint = null) => {
        if (!element || seen.has(element)) return;
        seen.add(element);
        const score = scoreContainer(element, source, selectorHint);
        if (!Number.isFinite(score) || score <= 0) return;
        candidates.push({ element, score, source, selectorHint });
    };

    for (const selector of selectors) {
        let matches = [];
        try {
            matches = document.querySelectorAll(selector);
        } catch (error) {
            continue;
        }
        for (const match of matches) {
            addCandidate(match, 'selector', selector);
        }
    }

    if (mainArea) {
        const descendants = mainArea.querySelectorAll('div, [role="feed"], [role="log"], [role="list"]');
        for (const el of descendants) {
            if ((el.clientHeight || 0) < 140 || (el.clientWidth || 0) < 140) continue;
            const style = window.getComputedStyle(el);
            const role = (el.getAttribute('role') || '').toLowerCase();
            const dataTid = (el.getAttribute('data-tid') || '').toLowerCase();
            const className = typeof el.className === 'string' ? el.className.toLowerCase() : '';
            const looksRelevant = role === 'feed' ||
                                  role === 'log' ||
                                  role === 'list' ||
                                  style.overflowY === 'auto' ||
                                  style.overflowY === 'scroll' ||
                                  dataTid.includes('message') ||
                                  dataTid.includes('channel') ||
                                  dataTid.includes('post') ||
                                  dataTid.includes('pane') ||
                                  className.includes('message') ||
                                  className.includes('virtual') ||
                                  className.includes('scroll');
            if (!looksRelevant) continue;
            addCandidate(el, 'main_descendant');
        }
    }

    if (cachedChatContainer && document.contains(cachedChatContainer)) {
        addCandidate(cachedChatContainer, 'cache_neighbor', 'cached');
        if (cachedChatContainer.parentElement) {
            addCandidate(cachedChatContainer.parentElement, 'cache_neighbor', 'cached_parent');
        }
    }

    if (candidates.length > 0) {
        candidates.sort((a, b) => b.score - a.score);
        const winner = candidates[0];
        const summary = candidates.slice(0, 3).map((c) => ({
            score: Math.round(c.score),
            source: c.source,
            selector: c.selectorHint,
            role: c.element.getAttribute('role') || '',
            dataTid: c.element.getAttribute('data-tid') || '',
            className: typeof c.element.className === 'string' ? c.element.className.substring(0, 80) : ''
        }));
        console.log('getChatContainer: Selected best candidate', summary);
        lastContainerSelectionDiagnostics = {
            mode: 'scored',
            selected: summary[0] || null,
            topCandidates: summary
        };

        cachedChatContainer = winner.element;
        containerCacheTimestamp = now;
        setupContainerObserver(winner.element);
        return winner.element;
    }

    // Ultimate fallback: Find first scrollable element in main area with message hints.
    if (mainArea) {
        const fallbackDivs = mainArea.querySelectorAll('div');
        for (const div of fallbackDivs) {
            const style = window.getComputedStyle(div);
            const isScrollable = style.overflowY === 'auto' || style.overflowY === 'scroll';
            const hasHeight = div.scrollHeight > div.clientHeight;
            const containsMessages = div.querySelector('[class*="message"], [class*="Message"], [data-tid*="message"]');

            if (isScrollable && hasHeight && containsMessages) {
                console.log('getChatContainer: Found container via ultimate fallback');
                lastContainerSelectionDiagnostics = {
                    mode: 'ultimate_fallback',
                    selected: {
                        role: div.getAttribute('role') || '',
                        dataTid: div.getAttribute('data-tid') || '',
                        className: typeof div.className === 'string' ? div.className.substring(0, 80) : ''
                    }
                };
                cachedChatContainer = div;
                containerCacheTimestamp = now;
                setupContainerObserver(div);
                return div;
            }
        }
    }

    lastContainerSelectionDiagnostics = {
        mode: 'not_found',
        reason: 'no_matching_container'
    };
    console.error('getChatContainer: No container found with any selector');
    return null;
}

async function waitForChatContainer(options = {}) {
    const timeoutMs = options.timeoutMs ?? 15000;
    const intervalMs = options.intervalMs ?? 350;
    const operationId = options.operationId || null;
    const startedAt = Date.now();
    let attempts = 0;
    let lastDiagnostics = null;

    while ((Date.now() - startedAt) <= timeoutMs) {
        if (isOperationCancelled(operationId)) {
            return {
                container: null,
                cancelled: true,
                attempts,
                waitedMs: Date.now() - startedAt,
                domDiagnostics: lastDiagnostics || captureDomDiagnostics()
            };
        }

        const container = getChatContainer(false);
        if (container) {
            return {
                container,
                cancelled: false,
                attempts,
                waitedMs: Date.now() - startedAt,
                domDiagnostics: captureDomDiagnostics()
            };
        }

        attempts++;
        lastDiagnostics = captureDomDiagnostics();
        await sleep(intervalMs);
    }

    return {
        container: null,
        cancelled: false,
        attempts,
        waitedMs: Date.now() - startedAt,
        domDiagnostics: lastDiagnostics || captureDomDiagnostics()
    };
}

const VIEWPORT_PROBE_MESSAGE_SELECTORS = [
    '[data-tid="chat-pane-item"]',
    '[data-tid="message-pane-item"]',
    '[data-tid*="chat-pane-item"]',
    '[data-tid*="message-pane-item"]',
    '[data-tid*="message-item"]',
    '[data-tid*="post-item"]',
    '[data-tid*="reply-item"]',
    '[data-tid*="thread-item"]',
    '[role="listitem"]',
    '[role="article"]'
];
const VIEWPORT_PROBE_MESSAGE_SELECTOR = VIEWPORT_PROBE_MESSAGE_SELECTORS.join(', ');
const VIEWPORT_PROBE_TEXT_SELECTOR = [
    '[class*="__body"]',
    '[class*="messageBody"]',
    '[class*="MessageBody"]',
    '[class*="messageContent"]',
    '[class*="MessageContent"]',
    '[class*="messageText"]',
    '[class*="MessageText"]',
    '[class*="postBody"]',
    '[class*="PostBody"]',
    '[class*="postContent"]',
    '[class*="PostContent"]',
    '[class*="replyBody"]',
    '[class*="ReplyBody"]',
    '[class*="replyContent"]',
    '[class*="ReplyContent"]',
    '[data-tid="messageContent"]',
    '[data-tid="message-content"]',
    '[data-tid*="messageBody"]',
    '[data-tid*="message-body"]',
    '[data-tid*="messageContent"]',
    '[data-tid*="message-content"]',
    '[data-tid*="messageText"]',
    '[data-tid*="message-text"]',
    '[data-tid*="postBody"]',
    '[data-tid*="post-body"]',
    '[data-tid*="postContent"]',
    '[data-tid*="post-content"]',
    '[data-tid*="replyBody"]',
    '[data-tid*="reply-body"]',
    '[data-tid*="replyContent"]',
    '[data-tid*="reply-content"]',
    '[data-tid*="channel-message"]',
    '[role="heading"]',
    'h1',
    'h2',
    'h3',
    'h4',
    'div[dir="auto"]',
    'p',
    'blockquote',
    'li',
].join(', ');
const VIEWPORT_PROBE_AUTHOR_SELECTOR = [
    '[data-tid="message-author-name"]',
    '[class*="__author"]',
    '[class*="personaName"]',
    '[class*="displayName"]'
].join(', ');

function normalizeProbeText(value) {
    return (value || '')
        .replace(/\u00a0/g, ' ')
        .replace(/[\u200b-\u200d\ufeff]/g, '')
        .replace(/\s+/g, ' ')
        .trim();
}

function sortElementsInDomOrder(elements) {
    return elements.sort((a, b) => {
        if (a === b) return 0;
        const relation = a.compareDocumentPosition(b);
        return (relation & Node.DOCUMENT_POSITION_FOLLOWING) ? -1 : 1;
    });
}

function isLikelyViewportProbeChannelCard(element) {
    if (!element || typeof element.querySelector !== 'function') return false;

    const role = (element.getAttribute('role') || '').toLowerCase();
    const dataTid = (element.getAttribute('data-tid') || '').toLowerCase();
    const text = normalizeProbeText(element.textContent || '');
    if (!text || text.length < 8) return false;

    if (dataTid.includes('post-item') || dataTid.includes('reply-item') || dataTid.includes('thread-item')) {
        return true;
    }

    if (role !== 'group') return false;

    const hasTime = element.querySelector('time') !== null;
    const hasSubject = element.querySelector('[data-tid="subject-line"]') !== null;
    const hasSubheader = element.querySelector('[data-tid="post-message-subheader"]') !== null;
    const hasBody = element.querySelector('div[dir="auto"], p, blockquote, li') !== null;

    return hasTime && (hasSubject || hasSubheader || hasBody);
}

function scoreViewportProbeMessageRoot(element) {
    if (!element) return -Infinity;

    const text = normalizeProbeText(element.textContent || '');
    if (!text) return -Infinity;

    const role = (element.getAttribute('role') || '').toLowerCase();
    const dataTid = (element.getAttribute('data-tid') || '').toLowerCase();
    const hasTime = element.querySelector('time') !== null;
    const hasBody = element.querySelector('div[dir="auto"], p, blockquote, li') !== null;
    const hasSubject = element.querySelector('[data-tid="subject-line"]') !== null;
    const hasSubheader = element.querySelector('[data-tid="post-message-subheader"]') !== null;
    const nestedRoots = element.querySelectorAll(VIEWPORT_PROBE_MESSAGE_SELECTOR).length;

    let score = Math.min(text.length, 260);
    if (dataTid.includes('chat-pane-item') || dataTid.includes('message-pane-item') || dataTid.includes('message-item')) score += 500;
    if (dataTid.includes('post-item') || dataTid.includes('reply-item') || dataTid.includes('thread-item')) score += 520;
    if (role === 'listitem' || role === 'article') score += 120;
    if (isLikelyViewportProbeChannelCard(element)) score += 380;
    if (hasTime) score += 100;
    if (hasBody) score += 140;
    if (hasSubject) score += 180;
    if (hasSubheader) score += 90;
    if (nestedRoots > 1) score -= Math.min(nestedRoots * 40, 180);

    return score;
}

function normalizeViewportProbeMessageElements(elements, container) {
    const promoted = [];
    const seen = new Set();

    for (const candidate of elements) {
        let best = candidate;
        let bestScore = scoreViewportProbeMessageRoot(candidate);
        let current = candidate.parentElement;
        let depth = 0;

        while (current && current !== container && depth < 6) {
            const currentScore = scoreViewportProbeMessageRoot(current);
            if (currentScore >= bestScore + 40) {
                best = current;
                bestScore = currentScore;
            }
            current = current.parentElement;
            depth++;
        }

        if (!seen.has(best)) {
            seen.add(best);
            promoted.push(best);
        }
    }

    const filtered = promoted.filter((element) => {
        const score = scoreViewportProbeMessageRoot(element);
        return !promoted.some((other) =>
            other !== element &&
            other.contains(element) &&
            scoreViewportProbeMessageRoot(other) >= score
        );
    });

    return sortElementsInDomOrder(filtered);
}

function getViewportProbeMessageElements(container) {
    if (!container) return [];

    const seen = new Set();
    const matches = [];

    try {
        container.querySelectorAll(VIEWPORT_PROBE_MESSAGE_SELECTOR).forEach((element) => {
            if (!element || seen.has(element) || !document.contains(element)) return;
            seen.add(element);
            matches.push(element);
        });
    } catch (error) {
        console.warn('Viewport probe selector query failed:', error?.message || error);
    }

    try {
        container.querySelectorAll('[role="group"]').forEach((element) => {
            if (!element || seen.has(element) || !document.contains(element)) return;
            if (!isLikelyViewportProbeChannelCard(element)) return;
            seen.add(element);
            matches.push(element);
        });
    } catch (error) {
        console.warn('Viewport probe role=group query failed:', error?.message || error);
    }

    if (matches.length > 0) {
        return normalizeViewportProbeMessageElements(matches, container);
    }

    const fallback = [];
    container.querySelectorAll('time').forEach((timeEl) => {
        let candidate = timeEl.parentElement;
        let bestCandidate = null;

        for (let depth = 0; candidate && candidate !== container && depth < 8; depth++) {
            const text = normalizeProbeText(candidate.textContent || '');
            if (text.length > 10 && text.length < 6000) {
                bestCandidate = candidate;
            }
            candidate = candidate.parentElement;
        }

        if (bestCandidate && !seen.has(bestCandidate) && document.contains(bestCandidate)) {
            seen.add(bestCandidate);
            fallback.push(bestCandidate);
        }
    });

    return normalizeViewportProbeMessageElements(fallback, container);
}

function hasExtractableVisibleFeed() {
    const container = getChatContainer(false);
    if (!container) return false;

    const probeMessages = getViewportProbeMessageElements(container).slice(0, 24);
    if (probeMessages.length === 0) return false;

    const collectedTextsFor = (element) => {
        const fragments = [];
        const seen = new Set();

        const addFragment = (value) => {
            const text = normalizeProbeText(value);
            if (!text) return;
            if (/^(edited|\(edited\)|message edited)$/i.test(text)) return;
            if (seen.has(text)) return;

            for (const existing of seen) {
                if (existing.includes(text) || text.includes(existing)) {
                    return;
                }
            }

            seen.add(text);
            fragments.push(text);
        };

        addFragment(element.textContent || '');
        try {
            if (element.matches(VIEWPORT_PROBE_TEXT_SELECTOR)) {
                addFragment(element.textContent || '');
            }
        } catch {
            // Ignore transient selector failures.
        }

        try {
            element.querySelectorAll(VIEWPORT_PROBE_TEXT_SELECTOR).forEach((node) => {
                addFragment(node.textContent || '');
            });
        } catch {
            // Ignore transient selector failures.
        }

        return fragments;
    };

    let viableCount = 0;

    for (const element of probeMessages) {
        const fragments = collectedTextsFor(element);
        const bestText = fragments.find(text => text.length >= 8) || fragments[0] || '';
        const author = normalizeProbeText(element.querySelector(VIEWPORT_PROBE_AUTHOR_SELECTOR)?.textContent || '');
        const hasTime = !!element.querySelector('time');
        const dataTid = (element.getAttribute('data-tid') || '').toLowerCase();
        const rawText = normalizeProbeText(element.textContent || '');

        if (bestText.length >= 8) {
            viableCount++;
        } else if (bestText && ((author && author.toLowerCase() !== bestText.toLowerCase()) || hasTime)) {
            viableCount++;
        } else if ((dataTid.includes('post-item') || dataTid.includes('reply-item') || dataTid.includes('thread-item')) && rawText.length >= 20) {
            viableCount++;
        } else if (findReplyExpansionCandidates(element).length > 0) {
            // A collapsed "N replies" thread is extractable once expanded.
            viableCount++;
        }

        if (viableCount >= Math.min(2, probeMessages.length)) {
            return true;
        }
    }

    return false;
}

// Tracks reply-expansion buttons already clicked so virtualized re-renders
// don't trigger duplicate clicks on the same live element.
let expandedReplyButtonSet = new WeakSet();

const REPLY_EXPANSION_LABEL_PATTERNS = [
    /\b\d+\s+repl(?:y|ies)\b/i,                                              // "3 replies", "See 12 replies from..."
    /\b(?:see|view|show|load)\s+(?:\d+\s+)?(?:more|all|earlier|previous)\s+repl(?:y|ies)\b/i,
    /\bmore\s+repl(?:y|ies)\b/i,
    /\b\d+\s+(?:respuestas?|antworten|r[ée]ponses?|antwoorden|risposte|respostas?|svar|odpowiedzi|답글|返信|回复)\b/i
];

function findReplyExpansionCandidates(root) {
    const scope = root || document.querySelector('[role="main"]') || document;
    let candidates;
    try {
        candidates = scope.querySelectorAll('button, [role="button"], [data-tid*="replies" i], [data-tid*="expand-thread" i]');
    } catch {
        candidates = scope.querySelectorAll('button, [role="button"]');
    }

    const buttons = [];
    for (const el of candidates) {
        if (expandedReplyButtonSet.has(el)) continue;
        // Never touch compose areas — a bare "Reply" button opens an editor.
        if (typeof el.closest === 'function' && el.closest('[contenteditable="true"], [data-tid*="compose" i], [data-tid*="editor" i]')) continue;

        const label = normalizeProbeText([
            el.getAttribute('aria-label') || '',
            el.textContent || ''
        ].join(' '));
        if (!label || label.length > 160) continue;
        if (/^(reply|reply in thread|respond)$/i.test(label)) continue;

        if (REPLY_EXPANSION_LABEL_PATTERNS.some(pattern => pattern.test(label))) {
            buttons.push(el);
        }
    }
    return buttons;
}

// Expands collapsed reply threads ("3 replies", "See more replies") currently
// rendered by the virtualized list. Teams channel posts keep replies out of the
// DOM until these buttons are clicked, so skipping this silently drops replies.
async function expandCollapsedRepliesInViewport(container, operationId, maxClicks = 8) {
    let clicked = 0;
    try {
        const buttons = findReplyExpansionCandidates(container);
        for (const button of buttons) {
            if (clicked >= maxClicks) break;
            if (isOperationCancelled(operationId)) break;
            expandedReplyButtonSet.add(button);
            try {
                button.click();
                clicked++;
                await sleep(350);
            } catch (error) {
                console.warn('expandCollapsedRepliesInViewport: click failed', error);
            }
        }
        if (clicked > 0) {
            // Replies are fetched lazily after the click; give Teams time to render them.
            await sleep(600);
            console.log(`expandCollapsedRepliesInViewport: expanded ${clicked} reply thread(s)`);
        }
    } catch (error) {
        console.warn('expandCollapsedRepliesInViewport: scan failed', error);
    }
    return clicked;
}

// Best-effort gate for the new Posts/Threads channel layout: instead of failing
// immediately, expand visible reply threads and wait for a readable feed.
async function waitForExtractableThreadsFeed(operationId, timeoutMs = 8000) {
    const startedAt = Date.now();
    while (Date.now() - startedAt < timeoutMs) {
        if (isOperationCancelled(operationId)) return false;
        const container = getChatContainer(false);
        await expandCollapsedRepliesInViewport(container, operationId, 4);
        if (hasExtractableVisibleFeed()) return true;
        await sleep(500);
    }
    return hasExtractableVisibleFeed();
}

function buildViewportMessageSignature(element) {
    if (!element) return '';

    const authorElement = element.querySelector(VIEWPORT_PROBE_AUTHOR_SELECTOR);
    const timeElement = element.querySelector('time');
    const textElement = element.querySelector(VIEWPORT_PROBE_TEXT_SELECTOR);

    const author = normalizeProbeText(authorElement?.textContent || '');
    const timestamp = normalizeProbeText(
        timeElement?.getAttribute('datetime') ||
        timeElement?.getAttribute('title') ||
        timeElement?.textContent ||
        ''
    );
    const text = normalizeProbeText((textElement || element).textContent || '').slice(0, 140);

    return `${author}|${timestamp}|${text}`;
}

function captureViewportMessageProbe(container) {
    if (!container) {
        return {
            scrollTop: 0,
            scrollHeight: 0,
            clientHeight: 0,
            messageCount: 0,
            visibleCount: 0,
            earliestVisibleTimestamp: null,
            isChannelLike: false,
            firstSignature: '',
            lastSignature: ''
        };
    }

    const messageElements = getViewportProbeMessageElements(container);
    const containerRect = container.getBoundingClientRect();
    const visibleElements = messageElements.filter((element) => {
        const rect = element.getBoundingClientRect();
        return rect.bottom > containerRect.top + 4 && rect.top < containerRect.bottom - 4;
    });

    const probeElements = visibleElements.length > 0 ? visibleElements : messageElements;
    const firstElement = probeElements[0] || null;
    const lastElement = probeElements[probeElements.length - 1] || null;
    const containerTid = (container.getAttribute('data-tid') || '').toLowerCase();
    const isChannelLike = containerTid.includes('channel') ||
        containerTid.includes('post') ||
        probeElements.some((element) => isLikelyViewportProbeChannelCard(element));

    return {
        scrollTop: Math.round(container.scrollTop || 0),
        scrollHeight: Math.round(container.scrollHeight || 0),
        clientHeight: Math.round(container.clientHeight || 0),
        messageCount: messageElements.length,
        visibleCount: visibleElements.length,
        earliestVisibleTimestamp: getEarliestVisibleTimestamp(container),
        isChannelLike,
        firstSignature: buildViewportMessageSignature(firstElement),
        lastSignature: buildViewportMessageSignature(lastElement)
    };
}

function normalizeTimestampCandidateText(rawText) {
    return (rawText || '').trim().replace(/\s+/g, ' ');
}

function getParseableTimestampCandidate(rawText, excludedTexts = []) {
    const normalized = normalizeTimestampCandidateText(rawText);
    if (!normalized) return '';

    const normalizedExclusions = excludedTexts
        .map(text => normalizeTimestampCandidateText(text))
        .filter(Boolean);
    if (normalizedExclusions.includes(normalized)) {
        return '';
    }

    return Number.isFinite(parseFlexibleTimestamp(normalized)) ? normalized : '';
}

function collectParseableTimestampCandidates(node, excludedTexts = []) {
    if (!node) return [];

    const candidates = [];
    const seen = new Set();
    const addCandidate = (rawText) => {
        const candidate = getParseableTimestampCandidate(rawText, excludedTexts);
        if (!candidate || seen.has(candidate)) return;
        seen.add(candidate);
        candidates.push(candidate);
    };

    addCandidate(node.getAttribute?.('datetime'));
    addCandidate(node.getAttribute?.('title'));

    if (typeof node.matches === 'function' && node.matches('time, [role="heading"]')) {
        addCandidate(node.textContent || '');
    }

    if (typeof node.querySelectorAll === 'function') {
        node.querySelectorAll('time, [role="heading"]').forEach((candidateNode) => {
            addCandidate(candidateNode.getAttribute?.('datetime'));
            addCandidate(candidateNode.getAttribute?.('title'));
            addCandidate(candidateNode.textContent || '');
        });
    }

    return candidates;
}

function inferTimestampContextForElement(element, currentDateHeader = '') {
    if (!element) {
        return currentDateHeader || '';
    }

    const elementContent = normalizeTimestampCandidateText(element.textContent || '');
    const directCandidates = collectParseableTimestampCandidates(element, [elementContent]);
    if (directCandidates.length > 0) {
        return directCandidates[0];
    }

    let sibling = element.previousElementSibling;
    let siblingSteps = 0;
    while (sibling && siblingSteps < 8) {
        const siblingCandidates = collectParseableTimestampCandidates(sibling, [elementContent]);
        if (siblingCandidates.length > 0) {
            return siblingCandidates[0];
        }
        sibling = sibling.previousElementSibling;
        siblingSteps++;
    }

    let ancestor = element.parentElement;
    let ancestorDepth = 0;
    while (ancestor && ancestorDepth < 3) {
        let ancestorSibling = ancestor.previousElementSibling;
        let ancestorSteps = 0;
        while (ancestorSibling && ancestorSteps < 6) {
            const ancestorCandidates = collectParseableTimestampCandidates(ancestorSibling, [elementContent]);
            if (ancestorCandidates.length > 0) {
                return ancestorCandidates[0];
            }
            ancestorSibling = ancestorSibling.previousElementSibling;
            ancestorSteps++;
        }
        ancestor = ancestor.parentElement;
        ancestorDepth++;
    }

    return currentDateHeader || '';
}

function getEarliestVisibleTimestamp(container) {
    if (!container) return null;

    const containerRect = container.getBoundingClientRect();
    const timestampElements = container.querySelectorAll('time, [role="heading"]');
    let earliestTimestamp = null;

    for (const timestampElement of timestampElements) {
        const rect = timestampElement.getBoundingClientRect();
        const isVisible = rect.bottom > containerRect.top + 4 && rect.top < containerRect.bottom - 4;
        if (!isVisible) continue;

        const candidates = collectParseableTimestampCandidates(timestampElement);
        for (const candidate of candidates) {
            const parsedTimestamp = parseFlexibleTimestamp(candidate);
            if (!Number.isFinite(parsedTimestamp)) continue;

            if (!Number.isFinite(earliestTimestamp) || parsedTimestamp < earliestTimestamp) {
                earliestTimestamp = parsedTimestamp;
            }
        }
    }

    return Number.isFinite(earliestTimestamp) ? earliestTimestamp : null;
}

function didViewportProbeChange(previousProbe, nextProbe, options = {}) {
    if (!previousProbe || !nextProbe) return true;

    const focus = options.focus === 'top' ? 'top' : 'full';
    const scrollChanged = Math.abs((nextProbe.scrollTop || 0) - (previousProbe.scrollTop || 0)) > 1;
    const heightChanged = Math.abs((nextProbe.scrollHeight || 0) - (previousProbe.scrollHeight || 0)) > 1;
    const earliestTimestampChanged = (nextProbe.earliestVisibleTimestamp || null) !== (previousProbe.earliestVisibleTimestamp || null);
    const firstSignatureChanged = nextProbe.firstSignature !== previousProbe.firstSignature;

    if (focus === 'top') {
        if (scrollChanged || heightChanged || earliestTimestampChanged || firstSignatureChanged) {
            return true;
        }

        const missingTopAnchors =
            !previousProbe.firstSignature &&
            !nextProbe.firstSignature &&
            !Number.isFinite(previousProbe.earliestVisibleTimestamp) &&
            !Number.isFinite(nextProbe.earliestVisibleTimestamp);

        if (missingTopAnchors) {
            return nextProbe.messageCount !== previousProbe.messageCount ||
                nextProbe.visibleCount !== previousProbe.visibleCount ||
                nextProbe.lastSignature !== previousProbe.lastSignature;
        }

        return false;
    }

    const channelLike = !!previousProbe.isChannelLike || !!nextProbe.isChannelLike;
    if (channelLike) {
        return scrollChanged ||
            heightChanged ||
            nextProbe.messageCount !== previousProbe.messageCount ||
            earliestTimestampChanged ||
            firstSignatureChanged;
    }

    return scrollChanged ||
        heightChanged ||
        nextProbe.messageCount !== previousProbe.messageCount ||
        nextProbe.visibleCount !== previousProbe.visibleCount ||
        firstSignatureChanged ||
        nextProbe.lastSignature !== previousProbe.lastSignature;
}

function sendTopLoadNudges(container) {
    if (!container) return;

    let temporaryTabIndex = false;
    try {
        const maxScrollTop = Math.max(0, (container.scrollHeight || 0) - (container.clientHeight || 0));
        if ((container.scrollTop || 0) <= 1 && maxScrollTop > 0) {
            // Crossing the top boundary helps virtualized lists request the next older page.
            const bounceDistance = Math.min(
                maxScrollTop,
                Math.max(180, Math.min(Math.round((container.clientHeight || 0) * 0.35), 640))
            );
            container.scrollTop = bounceDistance;
            if (typeof container.scrollTo === 'function') {
                container.scrollTo({ top: bounceDistance, behavior: 'auto' });
            }
        }

        if (!container.hasAttribute('tabindex')) {
            container.setAttribute('tabindex', '-1');
            temporaryTabIndex = true;
        }

        if (typeof container.focus === 'function') {
            container.focus({ preventScroll: true });
        }

        container.dispatchEvent(new KeyboardEvent('keydown', {
            key: 'Home',
            code: 'Home',
            keyCode: 36,
            which: 36,
            bubbles: true,
            cancelable: true
        }));
        container.dispatchEvent(new KeyboardEvent('keydown', {
            key: 'PageUp',
            code: 'PageUp',
            keyCode: 33,
            which: 33,
            bubbles: true,
            cancelable: true
        }));
        container.dispatchEvent(new WheelEvent('wheel', {
            deltaY: -1400,
            bubbles: true,
            cancelable: true
        }));
    } catch (error) {
        console.log('Top-load nudge failed:', error?.message || error);
    } finally {
        container.scrollTop = 0;
        if (typeof container.scrollTo === 'function') {
            container.scrollTo({ top: 0, behavior: 'auto' });
        }
        if (temporaryTabIndex) {
            container.removeAttribute('tabindex');
        }
    }
}

// Waits up to timeoutMs, returning early (after a short burst debounce) once the
// container's DOM changes. Resolves { mutated, elapsedMs }.
function waitForContainerMutation(container, timeoutMs, debounceMs = 120) {
    const startedAt = Date.now();
    if (!container || typeof MutationObserver !== 'function') {
        return sleep(timeoutMs).then(() => ({ mutated: false, elapsedMs: Date.now() - startedAt }));
    }

    return new Promise((resolve) => {
        let debounceTimer = null;
        let finished = false;
        const finish = (mutated) => {
            if (finished) return;
            finished = true;
            observer.disconnect();
            clearTimeout(timeoutTimer);
            clearTimeout(debounceTimer);
            resolve({ mutated, elapsedMs: Date.now() - startedAt });
        };
        const observer = new MutationObserver(() => {
            clearTimeout(debounceTimer);
            const remainingMs = timeoutMs - (Date.now() - startedAt);
            debounceTimer = setTimeout(() => finish(true), Math.max(0, Math.min(debounceMs, remainingMs)));
        });
        const timeoutTimer = setTimeout(() => finish(false), timeoutMs);
        observer.observe(container, { childList: true, subtree: true, characterData: true });
    });
}

async function waitForViewportToSettle(container, options = {}) {
    const {
        intervalMs = 500,
        maxRounds = 6,
        stableRoundsRequired = 2,
        requireTop = false,
        nudge = null,
        operationId = null,
        // A full interval with no DOM change at all counts as settled. Only for
        // already-loaded content: at the top a slow history fetch is quiet too.
        settleOnQuietInterval = false
    } = options;

    let previousProbe = captureViewportMessageProbe(container);
    let stableRounds = 0;
    let changeRounds = 0;
    let topDrifted = requireTop && previousProbe.scrollTop > 1;

    for (let round = 0; round < maxRounds; round++) {
        if (isOperationCancelled(operationId)) {
            return {
                settled: false,
                cancelled: true,
                probe: previousProbe,
                iterations: round,
                stableRounds,
                changeRounds,
                topDrifted
            };
        }

        if (typeof nudge === 'function') {
            nudge(container, round, previousProbe);
        }

        // Wake early when Teams renders something, so loading rounds take as long
        // as the content needs rather than a fixed interval.
        const wait = await waitForContainerMutation(container, intervalMs);

        if (isOperationCancelled(operationId)) {
            return {
                settled: false,
                cancelled: true,
                probe: previousProbe,
                iterations: round + 1,
                stableRounds,
                changeRounds,
                topDrifted
            };
        }

        let nextProbe = captureViewportMessageProbe(container);
        let changed = didViewportProbeChange(previousProbe, nextProbe, {
            focus: requireTop ? 'top' : 'full'
        });
        if (!changed && wait.mutated && wait.elapsedMs < intervalMs) {
            // Woke early for a change the probe does not track; a stable round
            // still needs the whole interval, as before.
            await sleep(intervalMs - wait.elapsedMs);
            nextProbe = captureViewportMessageProbe(container);
            changed = didViewportProbeChange(previousProbe, nextProbe, {
                focus: requireTop ? 'top' : 'full'
            });
        }
        if (!changed && !wait.mutated && settleOnQuietInterval && (!requireTop || nextProbe.scrollTop <= 1)) {
            return {
                settled: true,
                probe: nextProbe,
                iterations: round + 1,
                stableRounds: stableRoundsRequired,
                changeRounds,
                topDrifted
            };
        }
        const atRequiredPosition = !requireTop || nextProbe.scrollTop <= 1;

        if (changed) {
            changeRounds++;
        }
        if (!atRequiredPosition) {
            topDrifted = true;
        }

        if (!changed && atRequiredPosition) {
            stableRounds++;
        } else {
            stableRounds = 0;
        }

        previousProbe = nextProbe;

        if (stableRounds >= stableRoundsRequired) {
            return {
                settled: true,
                probe: nextProbe,
                iterations: round + 1,
                stableRounds,
                changeRounds,
                topDrifted
            };
        }
    }

    return {
        settled: false,
        probe: previousProbe,
        iterations: maxRounds,
        stableRounds,
        changeRounds,
        topDrifted
    };
}

// Enhanced function to scroll to the top of the chat.
// longChatMode: when true, scroll fast but wait much longer at the top to let content load.
// stopTimestamp: optional lower bound timestamp for incremental backups; once reached we stop early.
async function scrollToTop(longChatMode = false, stopTimestamp = null, operationId = null) {
    const initialWait = await waitForChatContainer({
        timeoutMs: longChatMode ? 25000 : 15000,
        intervalMs: 350,
        operationId
    });
    const initialContainer = initialWait.container;
    if (!initialContainer) {
        console.error("Chat container not found.");
        return {
            success: false,
            stoppedByUser: false,
            reason: initialWait.cancelled ? 'cancelled' : 'container_not_found',
            strategy: null,
            currentPosition: 0,
            finalScrollTop: 0,
            scrollHeight: 0,
            clientHeight: 0,
            attemptedGlobalRetries: 0,
            maxGlobalRetries: 0,
            strategies: [],
            containerWaitAttempts: initialWait.attempts,
            containerWaitMs: initialWait.waitedMs,
            domDiagnostics: initialWait.domDiagnostics || captureDomDiagnostics()
        };
    }

    isScrolling = true;
    const hasStopTimestamp = Number.isFinite(stopTimestamp);
    const maxGlobalRetries = longChatMode ? 45 : 5;
    const scrollWait = 500;
    const globalPause = longChatMode ? 15000 : 2000;
    const topSettleIntervalMs = longChatMode ? 2500 : 1000;
    const topSettleMaxRounds = longChatMode ? 22 : 8;
    const topStableRoundsRequired = longChatMode ? 4 : 2;
    const topLoadBudgetMs = longChatMode ? 30 * 60 * 1000 : 5 * 60 * 1000;
    const strategyLogs = [];
    const topLoadDeadline = Date.now() + topLoadBudgetMs;

    const describeContainer = (container) => ({
        tag: container?.tagName?.toLowerCase() || 'unknown',
        role: container?.getAttribute?.('role') || '',
        dataTid: container?.getAttribute?.('data-tid') || '',
        className: typeof container?.className === 'string' ? container.className.substring(0, 120) : ''
    });

    const getMetrics = (container) => ({
        currentPosition: Math.round(container?.scrollTop || 0),
        finalScrollTop: Math.round(container?.scrollTop || 0),
        scrollHeight: Math.round(container?.scrollHeight || 0),
        clientHeight: Math.round(container?.clientHeight || 0)
    });

    const checkStopDateReached = (container) => {
        if (!hasStopTimestamp) return null;

        const earliestVisibleTimestamp = getEarliestVisibleTimestamp(container);
        if (!Number.isFinite(earliestVisibleTimestamp)) return null;

        if (earliestVisibleTimestamp <= stopTimestamp) {
            return {
                reached: true,
                earliestVisibleTimestamp
            };
        }

        return {
            reached: false,
            earliestVisibleTimestamp
        };
    };

    const isClearlyShortConversation = (container) => {
        if (!longChatMode || !container) return false;

        const scrollRoom = Math.max(0, container.scrollHeight - container.clientHeight);
        if (scrollRoom > 32) return false;

        const renderedMessages = container.querySelectorAll(
            '[data-tid="chat-pane-item"], [data-tid*="message-pane-item"], [class*="fui-unstable-ChatItem"], [class*="fui-ChatMessage"], [class*="fui-ChatMyMessage"], [role="article"]'
        ).length;
        const hasLoadingIndicator = !!container.querySelector(
            '[aria-busy="true"], [role="progressbar"], [data-tid*="loading"], [data-tid*="spinner"]'
        );

        return renderedMessages > 0 && renderedMessages <= 30 && !hasLoadingIndicator;
    };

    const verifyTopHistoryLoaded = async (container) => {
        const quickShortCheck = isClearlyShortConversation(container);
        return waitForViewportToSettle(container, {
            intervalMs: quickShortCheck ? 700 : topSettleIntervalMs,
            maxRounds: quickShortCheck ? 6 : topSettleMaxRounds,
            stableRoundsRequired: quickShortCheck ? 2 : topStableRoundsRequired,
            requireTop: true,
            nudge: sendTopLoadNudges,
            operationId
        });
    };

    const verifyTopHistoryExhausted = async (container) => {
        const quickShortCheck = isClearlyShortConversation(container);
        const maxCycles = quickShortCheck ? 3 : (longChatMode ? 14 : 4);
        const repeatedStableCyclesRequired = quickShortCheck ? 1 : (longChatMode ? 4 : 1);
        let previousStableProbe = null;
        let repeatedStableCycles = 0;
        const cycles = [];

        for (let cycle = 0; cycle < maxCycles; cycle++) {
            if (longChatMode) {
                chrome.runtime.sendMessage({
                    action: 'updatePhase',
                    phase: quickShortCheck
                        ? 'Confirming the full chat is loaded...'
                        : `Checking for older messages (${cycle + 1}/${maxCycles})...`,
                    progress: Math.min(14, 6 + cycle),
                    operationId
                });
            }

            if (isOperationCancelled(operationId)) {
                return {
                    settled: false,
                    cancelled: true,
                    probe: previousStableProbe,
                    cycles,
                    repeatedStableCycles,
                    changeRounds: cycles.reduce((sum, entry) => sum + (entry.changeRounds || 0), 0),
                    topDrifted: cycles.some(entry => entry.topDrifted)
                };
            }

            if (Date.now() >= topLoadDeadline) {
                return {
                    settled: false,
                    timedOut: true,
                    probe: previousStableProbe,
                    cycles,
                    repeatedStableCycles,
                    changeRounds: cycles.reduce((sum, entry) => sum + (entry.changeRounds || 0), 0),
                    topDrifted: cycles.some(entry => entry.topDrifted)
                };
            }

            const stopCheck = checkStopDateReached(container);
            if (stopCheck?.reached) {
                return {
                    settled: true,
                    stoppedByDateRange: true,
                    earliestVisibleTimestamp: stopCheck.earliestVisibleTimestamp,
                    probe: previousStableProbe,
                    cycles,
                    repeatedStableCycles,
                    changeRounds: cycles.reduce((sum, entry) => sum + (entry.changeRounds || 0), 0),
                    topDrifted: cycles.some(entry => entry.topDrifted)
                };
            }

            const settle = await verifyTopHistoryLoaded(container);
            cycles.push({
                cycle: cycle + 1,
                settled: settle.settled,
                iterations: settle.iterations,
                changeRounds: settle.changeRounds,
                topDrifted: settle.topDrifted,
                probe: settle.probe
            });

            if (!settle.settled) {
                return {
                    settled: false,
                    probe: settle.probe,
                    cycles,
                    repeatedStableCycles,
                    changeRounds: cycles.reduce((sum, entry) => sum + (entry.changeRounds || 0), 0),
                    topDrifted: cycles.some(entry => entry.topDrifted)
                };
            }

            if (previousStableProbe && !didViewportProbeChange(previousStableProbe, settle.probe, { focus: 'top' })) {
                repeatedStableCycles++;
            } else {
                repeatedStableCycles = 0;
            }

            previousStableProbe = settle.probe;

            if (repeatedStableCycles >= repeatedStableCyclesRequired) {
                return {
                    settled: true,
                    probe: settle.probe,
                    cycles,
                    repeatedStableCycles,
                    changeRounds: cycles.reduce((sum, entry) => sum + (entry.changeRounds || 0), 0),
                    topDrifted: cycles.some(entry => entry.topDrifted)
                };
            }

            if (cycle < maxCycles - 1) {
                sendTopLoadNudges(container);
                await sleep(quickShortCheck ? 500 : (longChatMode ? 2200 : 400));
            }
        }

        return {
            settled: false,
            probe: previousStableProbe,
            cycles,
            repeatedStableCycles,
            changeRounds: cycles.reduce((sum, entry) => sum + (entry.changeRounds || 0), 0),
            topDrifted: cycles.some(entry => entry.topDrifted)
        };
    };

    const adoptContainer = (container) => {
        if (!container) return;
        cachedChatContainer = container;
        containerCacheTimestamp = Date.now();
        setupContainerObserver(container);
    };

    const finishResult = (base) => ({
        success: !!base.success,
        stoppedByUser: !!base.stoppedByUser,
        reason: base.reason || null,
        strategy: base.strategy || null,
        attemptedGlobalRetries: base.attemptedGlobalRetries || 0,
        maxGlobalRetries: maxGlobalRetries,
        strategies: strategyLogs,
        ...getMetrics(base.container || initialContainer)
    });

    const runDirectTopStrategy = async (container) => {
        let globalRetryCount = 0;
        let attempts = 0;
        let minTop = container.scrollTop;
        let maxTop = container.scrollTop;
        const startedAt = Date.now();

        while (globalRetryCount < maxGlobalRetries) {
            let lastScrollTop = container.scrollTop;
            let localRetryCount = 0;
            const maxLocalRetries = 5;
            let settleMadeProgress = false;

            while (localRetryCount < maxLocalRetries) {
                const stopCheck = checkStopDateReached(container);
                if (stopCheck?.reached) {
                    strategyLogs.push({
                        name: 'direct_top_date_range',
                        success: true,
                        attempts,
                        earliestVisibleTimestamp: stopCheck.earliestVisibleTimestamp,
                        elapsedMs: Date.now() - startedAt,
                        container: describeContainer(container)
                    });
                    return { success: true, stoppedByUser: false, reason: 'date_range_reached', strategy: 'direct_top', container, attemptedGlobalRetries: globalRetryCount };
                }

                if (!isScrolling || isOperationCancelled(operationId)) {
                    const log = {
                        name: 'direct_top',
                        success: false,
                        stoppedByUser: true,
                        attempts,
                        minTop: Math.round(minTop),
                        maxTop: Math.round(maxTop),
                        elapsedMs: Date.now() - startedAt,
                        container: describeContainer(container)
                    };
                    strategyLogs.push(log);
                    return { success: true, stoppedByUser: true, reason: 'stopped_by_user', strategy: 'direct_top', container, attemptedGlobalRetries: globalRetryCount };
                }

                attempts++;
                container.scrollTop = 0;
                if (typeof container.scrollTo === 'function') {
                    container.scrollTo({ top: 0, behavior: 'auto' });
                }
                await sleep(scrollWait);

                const currentScrollTop = container.scrollTop;
                minTop = Math.min(minTop, currentScrollTop);
                maxTop = Math.max(maxTop, currentScrollTop);

                if (currentScrollTop === 0) {
                    const probeBeforeSettle = captureViewportMessageProbe(container);
                    const topSettle = await verifyTopHistoryExhausted(container);
                    if (topSettle.settled) {
                        const settledByDateRange = topSettle.stoppedByDateRange === true;
                        const log = {
                            name: 'direct_top',
                            success: true,
                            attempts,
                            attemptedGlobalRetries: globalRetryCount,
                            minTop: Math.round(minTop),
                            maxTop: Math.round(maxTop),
                            settleCycles: topSettle.cycles?.length || 0,
                            settleChangeRounds: topSettle.changeRounds,
                            topDrifted: topSettle.topDrifted,
                            repeatedStableCycles: topSettle.repeatedStableCycles,
                            stoppedByDateRange: settledByDateRange,
                            earliestVisibleTimestamp: topSettle.earliestVisibleTimestamp || null,
                            elapsedMs: Date.now() - startedAt,
                            container: describeContainer(container)
                        };
                        strategyLogs.push(log);
                        return {
                            success: true,
                            stoppedByUser: false,
                            reason: settledByDateRange ? 'date_range_reached' : null,
                            strategy: 'direct_top',
                            container,
                            attemptedGlobalRetries: globalRetryCount
                        };
                    }

                    strategyLogs.push({
                        name: 'direct_top_settle',
                        success: false,
                        attempts,
                        attemptedGlobalRetries: globalRetryCount,
                        settleCycles: topSettle.cycles?.length || 0,
                        settleChangeRounds: topSettle.changeRounds,
                        topDrifted: topSettle.topDrifted,
                    repeatedStableCycles: topSettle.repeatedStableCycles,
                    timedOut: !!topSettle.timedOut,
                    finalProbe: topSettle.probe,
                    elapsedMs: Date.now() - startedAt,
                    container: describeContainer(container)
                });

                    if (topSettle.timedOut) {
                        return { success: false, stoppedByUser: false, reason: 'top_load_timeout', strategy: 'direct_top', container, attemptedGlobalRetries: globalRetryCount };
                    }

                    // Progress means more history rendered, not just scroll drift.
                    const probeAfterSettle = topSettle.probe || captureViewportMessageProbe(container);
                    settleMadeProgress =
                        probeAfterSettle.scrollHeight > probeBeforeSettle.scrollHeight + 1 ||
                        probeAfterSettle.messageCount > probeBeforeSettle.messageCount ||
                        (Number.isFinite(probeAfterSettle.earliestVisibleTimestamp) &&
                            Number.isFinite(probeBeforeSettle.earliestVisibleTimestamp) &&
                            probeAfterSettle.earliestVisibleTimestamp < probeBeforeSettle.earliestVisibleTimestamp);
                    localRetryCount = 0;
                    lastScrollTop = container.scrollTop;
                    break;
                }

                if (currentScrollTop === lastScrollTop) {
                    localRetryCount++;
                    await sleep(scrollWait);
                } else {
                    localRetryCount = 0;
                    lastScrollTop = currentScrollTop;
                }
            }

            // Older history is still arriving: keep loading right away instead of
            // pausing and spending a retry. The top-load budget still bounds this.
            if (settleMadeProgress && Date.now() < topLoadDeadline) {
                continue;
            }

            globalRetryCount++;
            await sleep(globalPause);
        }

        const log = {
            name: 'direct_top',
            success: false,
            attempts,
            attemptedGlobalRetries: globalRetryCount,
            minTop: Math.round(minTop),
            maxTop: Math.round(maxTop),
            elapsedMs: Date.now() - startedAt,
            container: describeContainer(container)
        };
        strategyLogs.push(log);
        return { success: false, stoppedByUser: false, reason: 'direct_top_failed', strategy: 'direct_top', container, attemptedGlobalRetries: globalRetryCount };
    };

    const runPagedScrollStrategy = async (container) => {
        const maxAttempts = longChatMode ? 160 : 35;
        const step = longChatMode
            ? Math.max(Math.round(container.clientHeight * 0.8), 260)
            : Math.max(Math.round(container.clientHeight * 1.25), 500);
        const stagnationLimit = longChatMode ? 24 : 12;
        const baseStagnationWait = longChatMode ? 900 : 350;
        let attempts = 0;
        let stagnationCount = 0;
        let minTop = container.scrollTop;
        let maxTop = container.scrollTop;
        let lastObservedScrollHeight = container.scrollHeight;
        const startedAt = Date.now();

        while (attempts < maxAttempts) {
            const stopCheck = checkStopDateReached(container);
            if (stopCheck?.reached) {
                strategyLogs.push({
                    name: 'paged_scroll_up_date_range',
                    success: true,
                    attempts,
                    minTop: Math.round(minTop),
                    maxTop: Math.round(maxTop),
                    step,
                    earliestVisibleTimestamp: stopCheck.earliestVisibleTimestamp,
                    elapsedMs: Date.now() - startedAt,
                    container: describeContainer(container)
                });
                return { success: true, stoppedByUser: false, reason: 'date_range_reached', strategy: 'paged_scroll_up', container };
            }

            if (Date.now() >= topLoadDeadline) {
                strategyLogs.push({
                    name: 'paged_scroll_up',
                    success: false,
                    timedOut: true,
                    attempts,
                    minTop: Math.round(minTop),
                    maxTop: Math.round(maxTop),
                    step,
                    elapsedMs: Date.now() - startedAt,
                    container: describeContainer(container)
                });
                return { success: false, stoppedByUser: false, reason: 'top_load_timeout', strategy: 'paged_scroll_up', container };
            }

            if (!isScrolling || isOperationCancelled(operationId)) {
                strategyLogs.push({
                    name: 'paged_scroll_up',
                    success: false,
                    stoppedByUser: true,
                    attempts,
                    minTop: Math.round(minTop),
                    maxTop: Math.round(maxTop),
                    step,
                    elapsedMs: Date.now() - startedAt,
                    container: describeContainer(container)
                });
                return { success: true, stoppedByUser: true, reason: 'stopped_by_user', strategy: 'paged_scroll_up', container };
            }

            attempts++;
            const before = container.scrollTop;
            if (typeof container.scrollBy === 'function') {
                container.scrollBy({ top: -step, behavior: 'auto' });
            }
            container.scrollTop = Math.max(0, container.scrollTop - step);
            await sleep(longChatMode ? 650 : 220);
            const after = container.scrollTop;

            minTop = Math.min(minTop, after);
            maxTop = Math.max(maxTop, after);

            if (after === 0) {
                const topSettle = await verifyTopHistoryExhausted(container);
                if (topSettle.settled) {
                    const settledByDateRange = topSettle.stoppedByDateRange === true;
                    strategyLogs.push({
                        name: 'paged_scroll_up',
                        success: true,
                        attempts,
                        minTop: Math.round(minTop),
                        maxTop: Math.round(maxTop),
                        step,
                        settleCycles: topSettle.cycles?.length || 0,
                        settleChangeRounds: topSettle.changeRounds,
                        topDrifted: topSettle.topDrifted,
                        repeatedStableCycles: topSettle.repeatedStableCycles,
                        stoppedByDateRange: settledByDateRange,
                        earliestVisibleTimestamp: topSettle.earliestVisibleTimestamp || null,
                        elapsedMs: Date.now() - startedAt,
                        container: describeContainer(container)
                    });
                    return {
                        success: true,
                        stoppedByUser: false,
                        reason: settledByDateRange ? 'date_range_reached' : null,
                        strategy: 'paged_scroll_up',
                        container
                    };
                }

                strategyLogs.push({
                    name: 'paged_scroll_up_settle',
                    success: false,
                    attempts,
                    minTop: Math.round(minTop),
                    maxTop: Math.round(maxTop),
                    step,
                    settleCycles: topSettle.cycles?.length || 0,
                    settleChangeRounds: topSettle.changeRounds,
                    topDrifted: topSettle.topDrifted,
                    repeatedStableCycles: topSettle.repeatedStableCycles,
                    timedOut: !!topSettle.timedOut,
                    finalProbe: topSettle.probe,
                    elapsedMs: Date.now() - startedAt,
                    container: describeContainer(container)
                });

                    if (topSettle.timedOut) {
                        return { success: false, stoppedByUser: false, reason: 'top_load_timeout', strategy: 'paged_scroll_up', container };
                    }

                    stagnationCount = 0;
                    continue;
                }

            if (after >= before) {
                const currentHeight = container.scrollHeight;
                const heightChanged = Math.abs(currentHeight - lastObservedScrollHeight) > 1;
                lastObservedScrollHeight = currentHeight;

                if (heightChanged) {
                    stagnationCount = 0;
                } else {
                    stagnationCount++;
                }

                if (stagnationCount >= stagnationLimit) {
                    break;
                }

                const adaptiveWaitMs = Math.min(baseStagnationWait + (stagnationCount * 220), longChatMode ? 5000 : 1800);
                await sleep(adaptiveWaitMs);
            } else {
                stagnationCount = 0;
                lastObservedScrollHeight = container.scrollHeight;
            }
        }

        strategyLogs.push({
            name: 'paged_scroll_up',
            success: false,
            attempts,
            minTop: Math.round(minTop),
            maxTop: Math.round(maxTop),
            step,
            elapsedMs: Date.now() - startedAt,
            container: describeContainer(container)
        });
        return { success: false, stoppedByUser: false, reason: 'paged_scroll_failed', strategy: 'paged_scroll_up', container };
    };

    const collectAlternateContainers = (baseContainer) => {
        const unique = new Set();
        const candidates = [];

        const addCandidate = (el, source) => {
            if (!el || unique.has(el)) return;
            unique.add(el);
            const scrollRoom = el.scrollHeight - el.clientHeight;
            if (scrollRoom <= 20) return;

            let score = scrollRoom;
            if (source === 'descendant') score += 1500;
            if (source === 'ancestor') score += 1000;
            if (source === 'main') score += 600;
            if ((el.getAttribute('role') || '').toLowerCase() === 'feed') score += 400;
            candidates.push({ el, source, score });
        };

        // Descendants inside main area often contain the actual virtualized scroller.
        const mainArea = document.querySelector('div[role="main"]') || document.body;
        const descendants = mainArea.querySelectorAll('div, [role="feed"], [role="log"], [role="list"]');
        for (const el of descendants) {
            addCandidate(el, 'descendant');
        }

        // Ancestors of current container can own the real scroll state in some Teams variants.
        let parent = baseContainer.parentElement;
        while (parent) {
            addCandidate(parent, 'ancestor');
            parent = parent.parentElement;
        }

        addCandidate(mainArea, 'main');

        return candidates
            .sort((a, b) => b.score - a.score)
            .map(c => c.el)
            .filter(el => el !== baseContainer)
            .slice(0, 8);
    };

    const runAlternateContainerStrategy = async (baseContainer) => {
        const candidates = collectAlternateContainers(baseContainer);
        const startedAt = Date.now();
        const maxNudgesPerContainer = longChatMode ? 24 : 8;
        const details = [];

        for (const candidate of candidates) {
            const preCandidateStopCheck = checkStopDateReached(candidate);
            if (preCandidateStopCheck?.reached) {
                strategyLogs.push({
                    name: 'alternate_container_nudges_date_range',
                    success: true,
                    candidatesTried: details.length,
                    earliestVisibleTimestamp: preCandidateStopCheck.earliestVisibleTimestamp,
                    elapsedMs: Date.now() - startedAt,
                    details
                });
                return { success: true, stoppedByUser: false, reason: 'date_range_reached', strategy: 'alternate_container_nudges', container: candidate };
            }

            if (Date.now() >= topLoadDeadline) {
                strategyLogs.push({
                    name: 'alternate_container_nudges',
                    success: false,
                    timedOut: true,
                    candidatesTried: details.length,
                    elapsedMs: Date.now() - startedAt,
                    details
                });
                return { success: false, stoppedByUser: false, reason: 'top_load_timeout', strategy: 'alternate_container_nudges', container: baseContainer };
            }

            if (!isScrolling || isOperationCancelled(operationId)) {
                strategyLogs.push({
                    name: 'alternate_container_nudges',
                    success: false,
                    stoppedByUser: true,
                    candidatesTried: details.length,
                    elapsedMs: Date.now() - startedAt,
                    details
                });
                return { success: true, stoppedByUser: true, reason: 'stopped_by_user', strategy: 'alternate_container_nudges', container: baseContainer };
            }

            let temporaryTabIndex = false;
            if (!candidate.hasAttribute('tabindex')) {
                candidate.setAttribute('tabindex', '-1');
                temporaryTabIndex = true;
            }
            if (typeof candidate.focus === 'function') {
                candidate.focus({ preventScroll: true });
            }

            let successForCandidate = false;
            let minTop = candidate.scrollTop;
            let maxTop = candidate.scrollTop;
            let settleSnapshot = null;

            for (let i = 0; i < maxNudgesPerContainer; i++) {
                if (!isScrolling || isOperationCancelled(operationId)) break;
                sendTopLoadNudges(candidate);
                await sleep(longChatMode ? 800 : 280);
                minTop = Math.min(minTop, candidate.scrollTop);
                maxTop = Math.max(maxTop, candidate.scrollTop);

                const stopCheck = checkStopDateReached(candidate);
                if (stopCheck?.reached) {
                    settleSnapshot = {
                        stoppedByDateRange: true,
                        earliestVisibleTimestamp: stopCheck.earliestVisibleTimestamp,
                        cycles: []
                    };
                    successForCandidate = true;
                    break;
                }

                if (candidate.scrollTop === 0) {
                    settleSnapshot = await verifyTopHistoryExhausted(candidate);
                    if (settleSnapshot.settled) {
                        successForCandidate = true;
                        break;
                    }
                    continue;
                }
            }

            details.push({
                container: describeContainer(candidate),
                minTop: Math.round(minTop),
                maxTop: Math.round(maxTop),
                finalTop: Math.round(candidate.scrollTop),
                settleCycles: settleSnapshot?.cycles?.length || 0,
                settleChangeRounds: settleSnapshot?.changeRounds || 0,
                topDrifted: !!settleSnapshot?.topDrifted,
                repeatedStableCycles: settleSnapshot?.repeatedStableCycles || 0,
                success: successForCandidate
            });

            if (temporaryTabIndex) {
                candidate.removeAttribute('tabindex');
            }

            if (successForCandidate) {
                const settledByDateRange = settleSnapshot?.stoppedByDateRange === true;
                strategyLogs.push({
                    name: 'alternate_container_nudges',
                    success: true,
                    candidatesTried: details.length,
                    stoppedByDateRange: settledByDateRange,
                    earliestVisibleTimestamp: settleSnapshot?.earliestVisibleTimestamp || null,
                    elapsedMs: Date.now() - startedAt,
                    details
                });
                return {
                    success: true,
                    stoppedByUser: false,
                    reason: settledByDateRange ? 'date_range_reached' : null,
                    strategy: 'alternate_container_nudges',
                    container: candidate
                };
            }
        }

        strategyLogs.push({
            name: 'alternate_container_nudges',
            success: false,
            candidatesTried: details.length,
            elapsedMs: Date.now() - startedAt,
            details
        });
        return { success: false, stoppedByUser: false, reason: 'alternate_container_failed', strategy: 'alternate_container_nudges', container: baseContainer };
    };

    console.log(`scrollToTop: longChatMode=${longChatMode}, stopTimestamp=${hasStopTimestamp ? stopTimestamp : 'none'}, maxGlobalRetries=${maxGlobalRetries}, topSettleIntervalMs=${topSettleIntervalMs}, topSettleMaxRounds=${topSettleMaxRounds}, strategies=3`);

    const directOutcome = await runDirectTopStrategy(initialContainer);
    if (directOutcome.success) {
        adoptContainer(directOutcome.container);
        await sleep(1000);
        return finishResult(directOutcome);
    }

    const pagedOutcome = await runPagedScrollStrategy(initialContainer);
    if (pagedOutcome.success) {
        adoptContainer(pagedOutcome.container);
        await sleep(800);
        return finishResult(pagedOutcome);
    }

    const alternateOutcome = await runAlternateContainerStrategy(initialContainer);
    if (alternateOutcome.success) {
        adoptContainer(alternateOutcome.container);
        await sleep(800);
        return finishResult(alternateOutcome);
    }

    console.warn('Could not confirm actual top after trying all scroll strategies.', { strategyLogs });
    return finishResult({
        success: false,
        stoppedByUser: false,
        reason: 'top_not_reached',
        strategy: 'all_failed',
        container: initialContainer,
        attemptedGlobalRetries: directOutcome.attemptedGlobalRetries || 0
    });
}

// Enhanced function to scroll to the bottom of the chat
async function scrollToBottom() {
    const chatContainer = getChatContainer();
    if (!chatContainer) {
        console.error("Chat container not found.");
        return false;
    }

    let lastScrollHeight = chatContainer.scrollHeight;
    let retryCount = 0;
    const maxRetries = 10; // Increase retries to ensure complete scrolling

    while (retryCount < maxRetries) {
        if (isCancelled) {
            console.log("Backup process was cancelled during scrolling down.");
            return false;
        }

        // Scroll down to the bottom
        chatContainer.scrollTop = chatContainer.scrollHeight;

        await sleep(500); // Wait for half a second to allow content to load

        const currentScrollHeight = chatContainer.scrollHeight;
        console.log(`Scrolled to position: ${chatContainer.scrollTop} / Scroll Height: ${currentScrollHeight}`);

        if (currentScrollHeight === lastScrollHeight) {
            retryCount++;
            console.log(`Retry ${retryCount}/${maxRetries}: No change in scroll height.`);
            await sleep(1000); // Wait before retrying
        } else {
            retryCount = 0; // Reset retry count if scroll height changed
            lastScrollHeight = currentScrollHeight;
        }

        // Check if reached the bottom
        if (chatContainer.scrollTop >= chatContainer.scrollHeight - chatContainer.clientHeight) {
            console.log("Reached the bottom of the chat.");
            break;
        }
    }

    if (chatContainer.scrollTop < chatContainer.scrollHeight - chatContainer.clientHeight) {
        console.warn("Could not scroll to the bottom after maximum retries.");
        return false;
    }

    return true;
}

// Function to capture all screenshots by scrolling from top to bottom
async function captureAllScreenshots(captureInterval, operationId = null) {
    const chatContainer = getChatContainer();
    if (!chatContainer) {
        console.error("Chat container not found.");
        return {
            screenshots: [],
            diagnostics: {
                reason: 'container_not_found',
                attempts: 0,
                failures: 0
            }
        };
    }

    const visibleHeight = chatContainer.offsetHeight;
    let scrollStep = visibleHeight * 0.9; // Overlap by 10% to ensure no content is missed
    let lastScrollTop = chatContainer.scrollTop;
    let screenshots = [];
    let index = 0;
    let reachedBottom = false;
    let attempts = 0;
    let failures = 0;
    let fallbackCaptures = 0;
    let firstError = null;

    async function tryCaptureElement(element, label) {
        if (!element) return null;
        let canvas;
        try {
            canvas = await html2canvas(element, {
                scale: 1,
                useCORS: true, // Enable CORS for images
                allowTaint: true, // Allow cross-origin images
                logging: false, // Disable logging for better performance
                imageTimeout: 2000, // Image loading timeout
                onclone: (clonedDoc) => {
                    // Force images to be visible in the cloned document
                    const images = clonedDoc.getElementsByTagName('img');
                    for (let img of images) {
                        img.style.visibility = 'visible';
                        img.style.opacity = '1';
                    }
                }
            });
        } catch (screenshotError) {
            failures++;
            if (!firstError) firstError = screenshotError?.message || String(screenshotError);
            console.error(`Screenshot capture failed (${label}):`, screenshotError);
            return null;
        }

        if (!canvas) {
            failures++;
            if (!firstError) firstError = 'html2canvas returned null';
            console.warn(`html2canvas returned null (${label})`);
            return null;
        }

        try {
            return canvas.toDataURL("image/jpeg", 0.8); // Use JPEG for better compression
        } catch (encodeError) {
            failures++;
            if (!firstError) firstError = encodeError?.message || String(encodeError);
            console.error(`Failed to encode screenshot (${label}):`, encodeError);
            return null;
        }
    }

    // Wait for images to load before starting capture
    await waitForImages(chatContainer, operationId);

    // Additional wait for content to fully render (especially for first capture)
    await sleep(1500);

    while (!reachedBottom) {
        if (isOperationCancelled(operationId)) {
            console.log("Backup process was cancelled during capturing screenshots.");
            break;
        }

        attempts++;
        const currentScrollTop = chatContainer.scrollTop;

        // Wait for content and images to load before each capture
        // Use captureInterval to control speed (higher = slower/safer for slow PCs)
        const waitTime = Math.round(500 * captureInterval);
        await sleep(waitTime);
        // Expand collapsed reply threads so PDF screenshots include channel replies.
        await expandCollapsedRepliesInViewport(chatContainer, operationId);
        await waitForImages(chatContainer, operationId);

        let dataURL = await tryCaptureElement(chatContainer, 'chatContainer');
        if (!dataURL) {
            // Fallback capture from broader main area when container capture fails.
            const mainArea = document.querySelector('div[role="main"]') || document.querySelector('[role="main"]');
            dataURL = await tryCaptureElement(mainArea, 'mainArea');
            if (dataURL) fallbackCaptures++;
        }

        if (dataURL) {
            screenshots.push(dataURL);
            index++;
            console.log(`Captured screenshot ${index}`);
        }

        // Calculate and send progress
        const safeScrollHeight = Math.max(chatContainer.scrollHeight, 1);
        const progress = Math.min((currentScrollTop / safeScrollHeight) * 100, 100);
        chrome.runtime.sendMessage({ action: "updateProgress", progress: progress });

        // Check if we've reached the bottom BEFORE scrolling
        // This ensures the last visible content is captured
        if (chatContainer.scrollTop >= chatContainer.scrollHeight - chatContainer.clientHeight - 5) {
            console.log("Reached the bottom of the chat after capturing final screenshot.");
            reachedBottom = true;
            break;
        }

        // Scroll down - speed controlled by captureInterval
        chatContainer.scrollTop += scrollStep;
        const scrollWait = Math.round(100 * captureInterval);
        await sleep(scrollWait);

        // Check if scroll position has changed
        const newScrollTop = chatContainer.scrollTop;
        if (newScrollTop === lastScrollTop) {
            // Try one more time with a small delay
            await sleep(500);
            chatContainer.scrollTop += scrollStep;
            if (chatContainer.scrollTop === newScrollTop) {
                console.log("No more content to load, stopping.");
                reachedBottom = true;
                break;
            }
        }

        lastScrollTop = newScrollTop;
    }

    if (screenshots.length === 0) {
        // Last-resort single capture to avoid hard PDF failure on some Teams layouts.
        const fallbackCandidates = [
            getChatContainer(false),
            document.querySelector('div[role="main"]'),
            document.querySelector('[role="main"]'),
            document.body
        ].filter(Boolean);

        const seen = new Set();
        for (const candidate of fallbackCandidates) {
            if (seen.has(candidate)) continue;
            seen.add(candidate);
            if ((candidate.clientWidth || 0) < 50 || (candidate.clientHeight || 0) < 50) continue;
            const fallbackShot = await tryCaptureElement(candidate, 'finalFallback');
            if (fallbackShot) {
                screenshots.push(fallbackShot);
                fallbackCaptures++;
                console.log('Captured final fallback screenshot');
                break;
            }
        }
    }

    return {
        screenshots,
        diagnostics: {
            reason: screenshots.length > 0 ? null : 'no_screenshots',
            attempts,
            failures,
            fallbackCaptures,
            firstError,
            finalScrollTop: chatContainer.scrollTop,
            scrollHeight: chatContainer.scrollHeight,
            clientHeight: chatContainer.clientHeight
        }
    };
}

// Message images that are still loading; avatars are not saved, so skip them.
function getPendingContentImages(element) {
    return Array.from(element.querySelectorAll('img')).filter(img =>
        !img.complete && !img.closest('[class*="avatar"], [class*="Avatar"]')
    );
}

// Wait for images to load. Accepts a container or a list of <img> elements.
async function waitForImages(target, operationId = null) {
    if (isOperationCancelled(operationId)) {
        return;
    }

    const images = Array.isArray(target)
        ? target
        : Array.from(target.getElementsByTagName('img'));
    const pendingImages = images.filter(img => !img.complete);
    if (pendingImages.length === 0) {
        return;
    }

    // addEventListener leaves Teams' own onload/onerror handlers in place.
    const imagePromises = pendingImages.map(img => new Promise(resolve => {
        img.addEventListener('load', resolve, { once: true });
        img.addEventListener('error', resolve, { once: true }); // Continue even if an image fails
    }));

    // Wait for all images with a timeout
    return Promise.race([
        Promise.all(imagePromises),
        new Promise(resolve => setTimeout(resolve, 3000)) // 3 second timeout
    ]);
}

// Update the compressImage function for better image quality
async function compressImage(dataURL, quality) {
    return new Promise((resolve) => {
        const img = new Image();
        img.onload = () => {
            const canvas = document.createElement('canvas');
            canvas.width = img.width;
            canvas.height = img.height;
            const ctx = canvas.getContext('2d');

            // Use better image rendering
            ctx.imageSmoothingEnabled = true;
            ctx.imageSmoothingQuality = 'high';

            ctx.drawImage(img, 0, 0);
            const compressedDataURL = canvas.toDataURL('image/jpeg', quality);
            resolve(compressedDataURL);
        };
        img.src = dataURL;
    });
}

// Function to generate a continuous PDF by stitching screenshots together
async function generateContinuousPDF(screenshots, isPreview = false, includeCover = true, operationId = null) {
    if (screenshots.length === 0) {
        console.error("No screenshots to process.");
        return null;
    }

    // Initialize jsPDF
    const jsPDF = window.jspdf.jsPDF;  // Updated to use correct namespace from local library
    const pdf = new jsPDF('p', 'pt', 'a4');
    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();
    const margin = 0; // Remove margins completely

    // Microsoft Teams colors
    const teamsBlue = [0, 120, 212]; // Primary Teams blue
    const teamsDarkBlue = [25, 65, 165]; // Secondary darker blue
    const teamsLightGray = [243, 242, 241]; // Light background color

    let currentPage = 1;
    let yPosition = 0;

    // Add cover header (compact, integrated with content) if enabled
    if (includeCover) {
        const headerHeight = 120;

        // Add cover header with Microsoft Teams styling
        pdf.setFillColor(teamsBlue[0], teamsBlue[1], teamsBlue[2]); // Teams blue color
        pdf.rect(0, 0, pageWidth, headerHeight, 'F');
        pdf.setTextColor(255, 255, 255);
        pdf.setFontSize(20);
        const title = isPreview ? 'Microsoft Teams Chat Backup - PREVIEW' : 'Microsoft Teams Chat Backup';
        pdf.text(title, pageWidth / 2, 35, { align: 'center' });
        pdf.setFontSize(10);
        pdf.text(`Generated on ${new Date().toLocaleString()}`, pageWidth / 2, 60, { align: 'center' });

        // Add description
        pdf.setFontSize(9);
        if (isPreview) {
            pdf.text('Free plan backup. Upgrade to Pro to back up every chat without the watermark.',
                 pageWidth / 2, 85, { align: 'center', maxWidth: pageWidth - 40 });
        } else {
            pdf.text('This backup includes all messages and content from your Microsoft Teams chat.',
                 pageWidth / 2, 85, { align: 'center', maxWidth: pageWidth - 40 });
        }

        pdf.setFontSize(8);
        pdf.text('Powered by Teams Chat Backup Extension', pageWidth / 2, 105, { align: 'center' });

        yPosition = headerHeight; // Start content after header
    }

    for (let i = 0; i < screenshots.length; i++) {
        if (isOperationCancelled(operationId)) {
            console.log("Backup process was cancelled during PDF generation.");
            return null;
        }

        const img = new Image();
        img.src = screenshots[i];
        await new Promise((resolve) => {
            img.onload = () => {
                const imgWidth = img.width;
                const imgHeight = img.height;
                const ratio = pageWidth / imgWidth;
                const width = pageWidth;
                const height = imgHeight * ratio;

                // If this image will exceed the page height
                if (yPosition + height > pageHeight) {
                    // Calculate remaining space on current page
                    const remainingSpace = pageHeight - yPosition;

                    // Draw first part of the image on current page
                    pdf.addImage(
                        screenshots[i],
                        'JPEG',
                        0, // x
                        yPosition, // y
                        width, // width
                        height, // height
                        `img${i}`, // unique id
                        'FAST', // compression
                        0 // rotation
                    );

                    // Add new page
                    pdf.addPage();
                    currentPage++;

                    // Calculate how much of the image was drawn
                    const drawnHeight = remainingSpace;
                    const remainingHeight = height - drawnHeight;

                    // Add small negative offset to improve alignment (reduce tiny overlap)
                    // Use different offset for first page break vs subsequent ones
                    const alignmentOffset = (i === 0) ? 6  : 9 ; // 4 points for first break, 6 for rest

                    // Draw remaining part of image on new page with alignment adjustment
                    pdf.addImage(
                        screenshots[i],
                        'JPEG',
                        0, // x
                        -(drawnHeight - alignmentOffset), // slight adjustment to reduce overlap
                        width,
                        height,
                        `img${i}_continued`, // unique id
                        'FAST', // compression
                        0 // rotation
                    );

                    yPosition = remainingHeight - alignmentOffset;
                } else {
                    // Normal case - draw entire image
                    pdf.addImage(
                        screenshots[i],
                        'JPEG',
                        0,
                        yPosition,
                        width,
                        height,
                        `img${i}`,
                        'FAST',
                        0
                    );
                    yPosition += height;
                }
                resolve();
            };
            img.onerror = () => {
                console.error(`Failed to load image ${i}`);
                resolve();
            };
        });
    }

    // Add a compact footer at the bottom of the last page if cover is enabled
    if (includeCover) {
        const footerHeight = 80;

        // Check if we need a new page or can fit footer on current page
        if (yPosition > pageHeight - footerHeight - 20) {
            pdf.addPage();
            currentPage++;
            yPosition = 0;
        }

        // Add footer
        const footerY = pageHeight - footerHeight;
        pdf.setFillColor(teamsBlue[0], teamsBlue[1], teamsBlue[2]); // Teams blue color
        pdf.rect(0, footerY, pageWidth, footerHeight, 'F');
        pdf.setTextColor(255, 255, 255);
        pdf.setFontSize(14);
        pdf.text('End of Teams Chat Backup', pageWidth / 2, footerY + 30, { align: 'center' });
        pdf.setFontSize(9);
        pdf.text(`Total pages: ${currentPage}`, pageWidth / 2, footerY + 50, { align: 'center' });
        pdf.text(`Generated by Microsoft Teams Chat Backup Extension`, pageWidth / 2, footerY + 65, { align: 'center' });
    }

    // Set PDF properties for better viewing
    pdf.setProperties({
        title: `Microsoft Teams Chat Backup ${new Date().toISOString()}`,
        subject: 'Microsoft Teams Chat History',
        creator: 'Teams Chat Backup Extension',
        author: 'Teams Chat Backup',
        keywords: 'microsoft, teams, chat, backup',
        producer: 'Teams Chat Backup Extension'
    });

    const pdfBlob = pdf.output('blob');
    const pdfUrl = createTrackedBlobUrl(pdfBlob);

    return pdfUrl;
}

// Add this function to extract text content from messages - extract all messages
// If startPosition is provided (number), start from that position; otherwise scroll to top first
// turboMode: when true, use minimal waits for faster text-only extraction (no PDF/screenshots)
// turboMode & longChatMode: both use fast scrolling (text-only extraction)
// longChatMode differs only in scrollToTop() which waits much longer at position 0
// fileAttachments: when true, extract file attachment metadata from messages
// maxFileSize: maximum file size in MB to include (files larger than this will have placeholder)
// existingFilenames: array of filenames that already exist in the database (skip re-downloading)
// checkpointConfig: { enabled, chatIdentity, isPro } for streaming Viewer-mode saves to the background worker
async function extractChatText(startPosition = null, turboMode = false, longChatMode = false, fileAttachments = false, maxFileSize = 50, existingFilenames = [], incrementalFromTimestamp = null, checkpointConfig = null, operationId = null) {
    const containerWait = await waitForChatContainer({
        timeoutMs: longChatMode ? 25000 : 15000,
        intervalMs: 350,
        operationId
    });
    let chatContainer = containerWait.container;
    if (!chatContainer) {
        console.error("Chat container not found for text extraction.");
        return {
            messages: [],
            errors: {
                messageErrors: 0,
                imageErrors: 0,
                fileErrors: 0,
                pdfError: null,
                screenshotDiagnostics: null,
                containerDiagnostics: lastContainerSelectionDiagnostics,
                containerLost: true,
                selectorMismatch: false,
                selectorUsed: null,
                topNotReached: false,
                scrollDiagnostics: {
                    reason: containerWait.cancelled ? 'cancelled' : 'container_not_found',
                    containerWaitAttempts: containerWait.attempts,
                    containerWaitMs: containerWait.waitedMs
                },
                domDiagnostics: containerWait.domDiagnostics || captureDomDiagnostics()
            }
        };
    }

    let chatText = [];
    const seenMessageKeys = new Set(); // O(1) duplicate detection instead of O(n)
    let extractedMessageCount = 0;
    const participantSet = new Set();
    let checkpointingEnabled = !!checkpointConfig?.enabled;
    let checkpointChatId = null;
    let checkpointPendingMessages = [];
    let checkpointPendingBytes = 0;
    let checkpointFlushCount = 0;
    let checkpointNewMessagesCount = 0;
    let checkpointStartedAsMerge = false;
    let checkpointFirstSaveCompleted = false;
    let lastCheckpointFlushAt = Date.now();

    if (checkpointingEnabled) {
        const identity = checkpointConfig?.chatIdentity || {};
        checkpointingEnabled = Boolean(identity.rawName || identity.name);
    }

    function addParticipant(author) {
        const normalized = (author || '').trim();
        if (!normalized || normalized === 'Unknown Author') return;
        participantSet.add(normalized);
    }
    const isTimestampInRange = (timestamp) => {
        const parsedTimestamp = parseFlexibleTimestamp(timestamp || '');

        if (Number.isFinite(incrementalFromTimestamp)) {
            if (!Number.isFinite(parsedTimestamp)) return false;
            // Strictly newer than the latest stored timestamp.
            if (parsedTimestamp <= incrementalFromTimestamp) return false;
        }

        return true;
    };

    // Error tracking for user notification
    const extractionErrors = {
        messageErrors: 0,
        imageErrors: 0,
        fileErrors: 0,
        pdfError: null,
        screenshotDiagnostics: null,
        containerDiagnostics: lastContainerSelectionDiagnostics,
        containerLost: false,
        selectorMismatch: false,
        selectorUsed: null,
        topNotReached: false,
        scrollDiagnostics: null,
        domDiagnostics: null
    };
    let detectedMessageElements = 0;

    if (isOperationCancelled(operationId)) {
        return {
            messages: chatText,
            errors: extractionErrors
        };
    }

    // Helper to create a unique key for a message
    function buildMessageMediaSignature(images = [], attachments = []) {
        const imageSignature = (images || [])
            .map(img => [
                normalizeWhitespace(img?.alt || ''),
                Number.isFinite(Number(img?.width)) ? String(Number(img.width)) : '',
                Number.isFinite(Number(img?.height)) ? String(Number(img.height)) : ''
            ].filter(Boolean).join(':'))
            .filter(Boolean)
            .sort();

        const attachmentSignature = (attachments || [])
            .map(att => [
                normalizeWhitespace(att?.filename || ''),
                normalizeWhitespace(att?.sharePointUrl || ''),
                normalizeWhitespace(att?.arrivalTime || '')
            ].filter(Boolean).join(':'))
            .filter(Boolean)
            .sort();

        return [...imageSignature, ...attachmentSignature].join('|').substring(0, 180);
    }

    function isRelaxedMeetingControlContent(content) {
        const normalized = normalizeWhitespace(content).toLowerCase();
        if (!normalized) return false;

        return normalized.includes('meeting ended:') ||
            normalized.endsWith('was invited to the meeting.') ||
            normalized.endsWith('was added to the meeting.') ||
            normalized.endsWith('was removed from the meeting.');
    }

    function createMessageKey(type, author, timestamp, content, mediaSignature = '') {
        const normalizedType = (type || 'message').trim().toLowerCase();
        const normalizedAuthor = (author || '').trim().toLowerCase();
        const rawTimestamp = (timestamp || '').trim();
        const parsedTimestamp = rawTimestamp ? parseFlexibleTimestamp(rawTimestamp) : NaN;
        const normalizedTimestamp = Number.isFinite(parsedTimestamp)
            ? new Date(parsedTimestamp).toISOString()
            : rawTimestamp.toLowerCase();
        const normalizedContent = (content || '').trim().replace(/\s+/g, ' ').substring(0, 120).toLowerCase();
        const normalizedMediaSignature = (mediaSignature || '').trim().toLowerCase();
        return `${normalizedType}|${normalizedAuthor}|${normalizedTimestamp}|${normalizedContent}|${normalizedMediaSignature}`;
    }

    function estimateMessageSize(message) {
        try {
            return JSON.stringify(message).length;
        } catch (error) {
            return (message?.content || '').length;
        }
    }

    function buildMessageImageReference(img) {
        if (!img) return null;

        const width = Number.isFinite(img.naturalWidth) && img.naturalWidth > 0
            ? img.naturalWidth
            : undefined;
        const height = Number.isFinite(img.naturalHeight) && img.naturalHeight > 0
            ? img.naturalHeight
            : undefined;
        const src = typeof img.src === 'string' && img.src && !img.src.startsWith('data:')
            ? img.src
            : undefined;

        return {
            alt: normalizeWhitespace(img.alt || '') || 'Image',
            width,
            height,
            src
        };
    }

    // Teams animated emojis/emoticons are vertical sprite strips (square frames
    // stacked top-to-bottom, animated via CSS). Detect them so they are not
    // captured as one tall unrolled image.
    function getAnimatedEmojiSpriteInfo(img) {
        if (!img || !img.complete) return null;

        const frameWidth = img.naturalWidth;
        const totalHeight = img.naturalHeight;
        if (!frameWidth || !totalHeight || totalHeight < frameWidth * 2) return null;

        const frameCount = Math.round(totalHeight / frameWidth);
        if (frameCount < 2 || frameCount > 400) return null;
        // Sprite strips are an exact stack of square frames; regular tall images are not.
        if (Math.abs(totalHeight - frameCount * frameWidth) > frameCount) return null;

        let matchesEmojiSelector = false;
        try {
            matchesEmojiSelector = img.matches(EMOJI_IMG_SELECTOR);
        } catch {
            matchesEmojiSelector = false;
        }

        if (!matchesEmojiSelector) {
            // Unhinted (e.g. blob: URL) sprites: emoticon frames are small and the
            // rendered box shows a single square frame, not the full strip.
            if (frameWidth > 128) return null;
            const rect = img.getBoundingClientRect();
            const showsSingleFrame = rect.width > 0 && rect.height > 0 && rect.height <= rect.width * 1.5;
            if (!showsSingleFrame) return null;
        }

        return { frameWidth, frameHeight: frameWidth, frameCount };
    }

    function extractAttachmentMetadata(message) {
        const attachmentElements = Array.from(message.querySelectorAll(ATTACHMENT_SELECTOR));
        const attachmentMetadata = [];
        const seenAttachmentKeys = new Set();

        for (const attachmentEl of attachmentElements) {
            try {
                const title = attachmentEl.getAttribute('title') || attachmentEl.getAttribute('aria-label') || '';
                const httpsIdx = title.indexOf('https');
                let filename = httpsIdx > 0 ? title.substring(0, httpsIdx).trim() : title.trim();
                let sharePointUrl = httpsIdx > 0 ? title.substring(httpsIdx).trim() : '';

                if (!sharePointUrl) {
                    const linkEl = attachmentEl.querySelector('a[href], [data-href]');
                    sharePointUrl = linkEl?.href || linkEl?.getAttribute('data-href') || '';
                }

                if (!filename) {
                    const nameEl = attachmentEl.querySelector('[class*="name"], [class*="title"], [data-tid*="file-name"], a[href]');
                    filename = getNodeText(nameEl);
                }

                if (!filename && sharePointUrl) {
                    try {
                        const url = new URL(sharePointUrl, window.location.href);
                        filename = decodeURIComponent(url.pathname.split('/').pop() || '').trim();
                    } catch (error) {
                        // Ignore malformed URLs and continue with empty filename.
                    }
                }

                const fileDomain = attachmentEl.getAttribute('filedomain') || '';
                const arrivalTime = attachmentEl.getAttribute('originalarrivaltime') || '';

                const sizeEl = attachmentEl.querySelector('[class*="size"], [class*="fileSize"]');
                const fileSizeText = sizeEl ? sizeEl.textContent.trim() : '';
                let fileSizeMB = 0;

                if (fileSizeText) {
                    const sizeMatch = fileSizeText.match(/([\d.]+)\s*(KB|MB|GB)/i);
                    if (sizeMatch) {
                        const size = parseFloat(sizeMatch[1]);
                        const unit = sizeMatch[2].toUpperCase();
                        if (unit === 'KB') fileSizeMB = size / 1024;
                        else if (unit === 'MB') fileSizeMB = size;
                        else if (unit === 'GB') fileSizeMB = size * 1024;
                    }
                }

                const extMatch = filename.match(/\.([^.]+)$/);
                const extension = extMatch ? extMatch[1].toLowerCase() : '';
                const attachmentKey = [
                    filename || '',
                    sharePointUrl || '',
                    fileDomain || '',
                    arrivalTime || ''
                ].join('|');

                if (seenAttachmentKeys.has(attachmentKey)) {
                    continue;
                }
                seenAttachmentKeys.add(attachmentKey);

                if (filename || sharePointUrl || fileDomain || arrivalTime) {
                    attachmentMetadata.push({
                        filename,
                        sharePointUrl,
                        fileDomain,
                        arrivalTime,
                        extension,
                        fileSizeText,
                        fileSizeMB
                    });
                }
            } catch (error) {
                console.error('Error collecting file metadata:', error);
            }
        }

        return attachmentMetadata;
    }

    function buildAttachmentFallbackContent(author, images = [], attachments = []) {
        const normalizedAuthor = normalizeWhitespace(author || '') || 'Unknown author';
        const imageCount = Array.isArray(images) ? images.length : 0;
        const attachmentNames = (attachments || [])
            .map(att => normalizeWhitespace(att?.filename || ''))
            .filter(Boolean);

        if (attachmentNames.length === 1 && imageCount === 0) {
            return `Message by ${normalizedAuthor}, attachment: ${attachmentNames[0]}.`;
        }

        if (attachmentNames.length > 1 && imageCount === 0) {
            const previewNames = attachmentNames.slice(0, 2).join(', ');
            const remainder = attachmentNames.length > 2 ? ` +${attachmentNames.length - 2} more` : '';
            return `Message by ${normalizedAuthor}, attachments: ${previewNames}${remainder}.`;
        }

        if (imageCount > 0 && attachmentNames.length === 0) {
            const firstImage = images[0] || {};
            const dimensionLabel = firstImage.width && firstImage.height
                ? ` (${firstImage.width}x${firstImage.height})`
                : '';
            if (imageCount === 1) {
                return `Message by ${normalizedAuthor}, image attachment${dimensionLabel}.`;
            }
            return `Message by ${normalizedAuthor}, ${imageCount} image attachments.`;
        }

        const totalAttachments = imageCount + attachmentNames.length;
        return `Message by ${normalizedAuthor}, has ${totalAttachments} attachment${totalAttachments === 1 ? '' : 's'}.`;
    }

    function createCheckpointSafeMessage(message) {
        const sanitizedImages = Array.isArray(message?.images)
            ? message.images.map(img => {
                if (!img) return null;
                return {
                    alt: normalizeWhitespace(img.alt || '') || 'Image',
                    width: Number.isFinite(img.width) ? img.width : undefined,
                    height: Number.isFinite(img.height) ? img.height : undefined,
                    src: typeof img.src === 'string' && img.src && !img.src.startsWith('data:')
                        ? img.src
                        : undefined,
                    checkpointPlaceholder: true
                };
            }).filter(Boolean)
            : [];

        const sanitizedAttachments = Array.isArray(message?.attachments)
            ? message.attachments.map(att => {
                if (!att) return null;
                return {
                    filename: att.filename || 'Attachment',
                    sharePointUrl: att.sharePointUrl || '',
                    fileDomain: att.fileDomain || '',
                    arrivalTime: att.arrivalTime || '',
                    extension: att.extension || '',
                    fileSizeText: att.fileSizeText || '',
                    fileSizeMB: Number.isFinite(att.fileSizeMB) ? att.fileSizeMB : 0,
                    tooLarge: att.tooLarge === true,
                    downloadError: att.downloadError || null,
                    existsInDb: att.existsInDb === true,
                    checkpointPlaceholder: true
                };
            }).filter(Boolean)
            : [];

        const rawContent = typeof message?.content === 'string' ? message.content : '';
        let content = rawContent;
        if (!normalizeWhitespace(rawContent) && (sanitizedImages.length > 0 || sanitizedAttachments.length > 0)) {
            content = buildAttachmentFallbackContent(message?.author, sanitizedImages, sanitizedAttachments);
        }

        return {
            ...message,
            content,
            images: sanitizedImages.length > 0 ? sanitizedImages : undefined,
            attachments: sanitizedAttachments.length > 0 ? sanitizedAttachments : undefined
        };
    }

    async function flushCheckpointMessages(reason = 'threshold') {
        if (!checkpointingEnabled || checkpointPendingMessages.length === 0) {
            return;
        }

        const pendingMessages = checkpointPendingMessages;
        checkpointPendingMessages = [];
        checkpointPendingBytes = 0;

        try {
            const response = await sendRuntimeMessageWithResponse({
                action: 'checkpointBackupChunk',
                operationId,
                reason,
                chatIdentity: checkpointConfig.chatIdentity,
                // After the first save, point straight at the saved chat record.
                checkpointChatId,
                messages: pendingMessages,
                participants: Array.from(participantSet),
                isPro: checkpointConfig?.isPro === true
            });

            if (!response || response.status !== 'success') {
                const saveError = new Error(response?.message || response?.error || 'Checkpoint save failed');
                // Keep db.js codes such as 'chat_limit' / 'update_locked' so the popup
                // can show the free-tier upgrade message instead of a generic error.
                saveError.code = response?.error || null;
                saveError.count = response?.count;
                throw saveError;
            }

            if (!checkpointFirstSaveCompleted) {
                checkpointFirstSaveCompleted = true;
                checkpointStartedAsMerge = response.merged === true;
            }

            checkpointChatId = response.id || checkpointChatId;
            checkpointFlushCount += 1;
            checkpointNewMessagesCount += Number(response.newMessagesCount || 0);
            lastCheckpointFlushAt = Date.now();

            chrome.runtime.sendMessage({
                action: "updatePhase",
                phase: `Saved checkpoint (${extractedMessageCount} messages)...`,
                progress: 35,
                operationId
            });
        } catch (error) {
            checkpointPendingMessages = pendingMessages.concat(checkpointPendingMessages);
            checkpointPendingBytes = checkpointPendingMessages.reduce((sum, message) => sum + estimateMessageSize(message), 0);
            const checkpointError = error instanceof Error
                ? error
                : new Error(String(error || 'Checkpoint save failed'));
            checkpointError.isCheckpointFailure = true;
            throw checkpointError;
        }
    }

    async function queueExtractedMessage(message) {
        markOperationActivity(operationId);
        extractedMessageCount += 1;
        addParticipant(message.author);

        const messageToStore = checkpointingEnabled
            ? createCheckpointSafeMessage(message)
            : message;

        if (checkpointingEnabled) {
            checkpointPendingMessages.push(messageToStore);
            checkpointPendingBytes += estimateMessageSize(messageToStore);

            const batchTooLarge = checkpointPendingMessages.length >= CHECKPOINT_BATCH_MESSAGE_LIMIT ||
                checkpointPendingBytes >= CHECKPOINT_BATCH_SIZE_LIMIT_BYTES;
            const batchTooOld = (Date.now() - lastCheckpointFlushAt) >= CHECKPOINT_BATCH_MAX_AGE_MS;

            if (batchTooLarge || batchTooOld) {
                await flushCheckpointMessages(batchTooLarge ? 'batch_limit' : 'interval');
            }
            return;
        }

        chatText.push(messageToStore);
    }

    const MESSAGE_TEXT_SELECTOR = [
        'div[class*="fui-ChatMyMessage__body"]',
        'div[class*="fui-ChatMessage__body"]',
        '[class*="__body"]',
        '[class*="messageBody"]',
        '[class*="MessageBody"]',
        '[class*="messageContent"]',
        '[class*="MessageContent"]',
        '[class*="messageText"]',
        '[class*="MessageText"]',
        '[class*="postBody"]',
        '[class*="PostBody"]',
        '[class*="postContent"]',
        '[class*="PostContent"]',
        '[class*="replyBody"]',
        '[class*="ReplyBody"]',
        '[class*="replyContent"]',
        '[class*="ReplyContent"]',
        '[data-tid="messageContent"]',
        '[data-tid="message-content"]',
        '[data-tid*="messageBody"]',
        '[data-tid*="message-body"]',
        '[data-tid*="messageContent"]',
        '[data-tid*="message-content"]',
        '[data-tid*="messageText"]',
        '[data-tid*="message-text"]',
        '[data-tid*="postBody"]',
        '[data-tid*="post-body"]',
        '[data-tid*="postContent"]',
        '[data-tid*="post-content"]',
        '[data-tid*="replyBody"]',
        '[data-tid*="reply-body"]',
        '[data-tid*="replyContent"]',
        '[data-tid*="reply-content"]',
        '[data-tid*="channel-message"]',
        '[role="heading"]',
        'h1',
        'h2',
        'h3',
        'h4',
        'div[dir="auto"]',
        'p',
        'blockquote',
        'li'
    ].join(', ');
    const MESSAGE_META_SELECTOR = [
        '[data-tid="message-author-name"]',
        '[class*="__author"]',
        '[class*="personaName"]',
        '[class*="displayName"]',
        '[class*="Avatar"]',
        '[class*="avatar"]',
        '[class*="__timestamp"]',
        'time'
    ].join(', ');
    const ATTACHMENT_SELECTOR = [
        '.ui-attachment',
        '[data-tid*="attachment"]',
        '[data-tid*="file-attachment"]',
        '[data-tid*="file-card"]',
        '[data-tid*="message-file"]',
        '[class*="FileCard"]',
        '[class*="filecard"]'
    ].join(', ');
    const MESSAGE_BODY_HINT_TOKENS = [
        '__body',
        'messagebody',
        'message-body',
        'messagecontent',
        'message-content',
        'messagetext',
        'message-text',
        'postbody',
        'post-body',
        'postcontent',
        'post-content',
        'replybody',
        'reply-body',
        'replycontent',
        'reply-content',
        'channel-message'
    ];
    const MESSAGE_CONTROL_HINT_TOKENS = [
        'toolbar',
        'reaction',
        'actions',
        'menu',
        'footer',
        'status',
        'edited'
    ];

    function isLikelyDateDividerText(value) {
        const text = normalizeWhitespace(value);
        if (!text || text.length > 80) return false;

        // Avoid treating message headings such as "file by Jane Doe" as date dividers.
        if (/\b(by|has an attachment|could not download|open in sharepoint|reply|message|chat)\b/i.test(text)) {
            return false;
        }

        const weekday = '(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)';
        const month = '(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)';
        const weekdayMonthDay = new RegExp(`^${weekday},?\\s+${month}\\.?\\s+\\d{1,2}(?:,?\\s+\\d{4})?$`, 'i');
        const monthDay = new RegExp(`^${month}\\.?\\s+\\d{1,2}(?:,?\\s+\\d{4})?$`, 'i');

        return /^(today|yesterday)$/i.test(text) ||
            weekdayMonthDay.test(text) ||
            monthDay.test(text) ||
            /^\d{1,2}[\/.-]\d{1,2}[\/.-]\d{2,4}$/.test(text);
    }

    function isLikelyChannelPostCard(element) {
        if (!element || typeof element.querySelector !== 'function') return false;

        const role = (element.getAttribute('role') || '').toLowerCase();
        const dataTid = (element.getAttribute('data-tid') || '').toLowerCase();

        if (dataTid.includes('post-item') || dataTid.includes('thread-item') || dataTid.includes('reply-item')) {
            return true;
        }

        const hasPostHeader = element.querySelector('[data-tid="post-message-subheader"]') !== null;
        const hasSubjectLine = element.querySelector('[data-tid="subject-line"]') !== null;
        const hasTimestamp = element.querySelector('time[data-tid="timestamp"], time') !== null;
        const hasReactionSummary = element.querySelector('[data-tid="channel-message-reaction-summary"]') !== null;

        return role === 'group' && hasTimestamp && (hasPostHeader || hasSubjectLine || hasReactionSummary);
    }

    function normalizeWhitespace(value) {
        return (value || '')
            .replace(/\u00a0/g, ' ')
            .replace(/[\u200b-\u200d\ufeff]/g, '')
            .replace(/\s+/g, ' ')
            .trim();
    }

    // Teams renders emojis/emoticons as <img> elements, which contribute nothing to
    // innerText/textContent. Match them so their alt text can be inlined into the
    // extracted message text (and so sprite strips are not captured as images).
    const EMOJI_IMG_SELECTOR = [
        'img[itemtype*="Emoji" i]',
        'img[itemtype*="emoticon" i]',
        'img[class*="emoji" i]',
        'img[class*="emoticon" i]',
        'img[src*="emoticon" i]',
        'img[src*="personal-expressions" i]',
        'img[data-emoticon-name]'
    ].join(', ');

    function getEmojiAltText(img) {
        const alt = normalizeWhitespace(
            img.getAttribute('alt') ||
            img.getAttribute('data-emoticon-name') ||
            img.getAttribute('aria-label') ||
            ''
        );
        if (!alt || alt.length > 60 || /^image$/i.test(alt)) return '';
        return alt;
    }

    function getNodeText(node) {
        if (!node) return '';

        let emojiImgs = null;
        try {
            emojiImgs = typeof node.querySelectorAll === 'function'
                ? node.querySelectorAll(EMOJI_IMG_SELECTOR)
                : null;
        } catch {
            emojiImgs = null;
        }

        if (emojiImgs && emojiImgs.length > 0) {
            // Clone so emoji alt text can be spliced in without mutating the live Teams DOM.
            const clone = node.cloneNode(true);
            try {
                clone.querySelectorAll(EMOJI_IMG_SELECTOR).forEach(img => {
                    const alt = getEmojiAltText(img);
                    img.replaceWith((clone.ownerDocument || document).createTextNode(alt ? ` ${alt} ` : ''));
                });
                return normalizeWhitespace(clone.innerText || clone.textContent || '');
            } catch {
                // Fall through to the plain extraction below.
            }
        }

        return normalizeWhitespace(node.innerText || node.textContent || '');
    }

    function elementLooksLikeMessageBody(element) {
        if (!element || typeof element.matches !== 'function') return false;

        try {
            if (element.matches(MESSAGE_TEXT_SELECTOR)) {
                return true;
            }
        } catch {
            // Ignore selector matching failures for detached/transient nodes.
        }

        const className = typeof element.className === 'string' ? element.className.toLowerCase() : '';
        const dataTid = (element.getAttribute('data-tid') || '').toLowerCase();
        return MESSAGE_BODY_HINT_TOKENS.some(token => className.includes(token) || dataTid.includes(token));
    }

    function elementHasMessageBodyHint(element) {
        if (!element) return false;
        if (elementLooksLikeMessageBody(element)) return true;

        try {
            return element.querySelector(MESSAGE_TEXT_SELECTOR) !== null;
        } catch {
            return false;
        }
    }

    function isShortStatusOnlyText(text) {
        const normalized = normalizeWhitespace(text).toLowerCase();
        return normalized === 'edited' ||
            normalized === '(edited)' ||
            normalized === 'message edited' ||
            normalized === 'see more' ||
            normalized === 'view more' ||
            normalized === 'like reaction' ||
            normalized === '1 like reaction' ||
            normalized === 'reply in thread' ||
            normalized === 'post in channel' ||
            normalized === 'responder en la conversacion' ||
            normalized === 'responder en la conversación' ||
            normalized === 'publicar en el canal';
    }

    function looksLikeMessageControlElement(element) {
        if (!element) return false;

        const tagName = (element.tagName || '').toLowerCase();
        const role = (element.getAttribute('role') || '').toLowerCase();
        const className = typeof element.className === 'string' ? element.className.toLowerCase() : '';
        const dataTid = (element.getAttribute('data-tid') || '').toLowerCase();
        const ariaLabel = (element.getAttribute('aria-label') || '').toLowerCase();
        const text = normalizeWhitespace(element.textContent || '').toLowerCase();

        if (tagName === 'button') return true;
        if (role === 'button' || role === 'toolbar' || role === 'menu' || role === 'menuitem') return true;
        if (/^(reply in thread|post in channel|responder en la conversaci(?:o|ó)n|publicar en el canal)$/.test(text)) return true;

        return MESSAGE_CONTROL_HINT_TOKENS.some(token =>
            className.includes(token) ||
            dataTid.includes(token) ||
            ariaLabel.includes(token)
        );
    }

    function stripMessageNoise(root) {
        if (!root || typeof root.querySelectorAll !== 'function') return;

        root.querySelectorAll('*').forEach(node => {
            if (isMetadataLikeElement(node) || looksLikeMessageControlElement(node)) {
                node.remove();
                return;
            }

            try {
                if (node.matches(ATTACHMENT_SELECTOR)) {
                    node.remove();
                }
            } catch {
                // Ignore transient selector/match failures while cleaning clones.
            }
        });
    }

    function stripKnownMessageMetadata(text, author, timestamp) {
        let value = normalizeWhitespace(text);
        if (!value) return '';

        const edgeTokens = [...new Set(
            [author, timestamp]
                .map(token => normalizeWhitespace(token))
                .filter(Boolean)
        )];

        for (const token of edgeTokens) {
            const normalizedToken = token.toLowerCase();
            if (value.toLowerCase().startsWith(normalizedToken)) {
                value = normalizeWhitespace(value.slice(token.length));
            }
            if (value.toLowerCase().endsWith(normalizedToken)) {
                value = normalizeWhitespace(value.slice(0, value.length - token.length));
            }
        }

        value = value
            .replace(/^(today|yesterday)\s+/i, '')
            .replace(/^\d{1,2}:\d{2}(?::\d{2})?\s*(AM|PM)?\s*/i, '')
            .replace(/^\d{1,2}[-\/.]\d{1,2}[-\/.]\d{2,4}\s+\d{1,2}:\d{2}(?::\d{2})?\s*(AM|PM)?\s*/i, '')
            .replace(/^[A-Za-z]{3,9},\s+[A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4}\s+\d{1,2}:\d{2}(?::\d{2})?\s*(AM|PM)?\s*/i, '')
            .replace(/\b(?:reply in thread|post in channel|view more replies|see more replies|responder en la conversaci(?:o|ó)n|publicar en el canal)\b\s*$/i, '')
            .replace(/\bsee more\b\s*$/i, '')
            .replace(/\bview more\b\s*$/i, '')
            .replace(/\b\d+\s+like reactions?\b\.?\s*$/i, '')
            .replace(/\blike reactions?\b\.?\s*$/i, '');

        return normalizeWhitespace(value);
    }

    function normalizeChannelPostFragment(value, author, timestamp) {
        return stripKnownMessageMetadata(value, author, timestamp)
            .replace(/\bsee more\b(?=\s*(?:\d+\s+like reactions?\b|like reactions?\b|$))/ig, '')
            .replace(/\bview more\b(?=\s*(?:\d+\s+like reactions?\b|like reactions?\b|$))/ig, '')
            .replace(/\b\d+\s+like reactions?\b/ig, '')
            .replace(/\blike reactions?\b/ig, '')
            .replace(/\s+/g, ' ')
            .trim();
    }

    function isMetadataLikeElement(element) {
        if (!element) return false;
        if (isLikelyChannelPostCard(element)) return false;
        const tagName = (element.tagName || '').toLowerCase();
        const className = typeof element.className === 'string' ? element.className.toLowerCase() : '';
        const dataTid = (element.getAttribute('data-tid') || '').toLowerCase();
        const text = getNodeText(element);
        const hasBodyDescendants = !!element.querySelector('div[dir="auto"], p, blockquote, li, [role="heading"], h1, h2, h3, h4');
        const hasTimeDescendant = tagName === 'time' || element.querySelector('time') !== null;
        const isHeaderLikeOnly = (
            dataTid.includes('header') ||
            className.includes('__header') ||
            className.includes('messageheader') ||
            className.includes('postheader')
        ) && !hasBodyDescendants && text.length < 180;

        return tagName === 'time' ||
            dataTid.includes('author') ||
            dataTid.includes('timestamp') ||
            dataTid.includes('persona') ||
            className.includes('author') ||
            className.includes('timestamp') ||
            className.includes('persona') ||
            className.includes('avatar') ||
            (isHeaderLikeOnly && hasTimeDescendant);
    }

    function scoreMessageRoot(element) {
        const text = getNodeText(element);
        if (!text) return -Infinity;

        const className = typeof element.className === 'string' ? element.className : '';
        const lowerClass = className.toLowerCase();
        const dataTid = (element.getAttribute('data-tid') || '').toLowerCase();
        const role = (element.getAttribute('role') || '').toLowerCase();

        let score = Math.min(text.length, 260);

        if (dataTid.includes('chat-pane-item') ||
            dataTid.includes('message-pane-item') ||
            dataTid.includes('message-item') ||
            dataTid.includes('post-item') ||
            dataTid.includes('reply-item') ||
            dataTid.includes('thread-item')) {
            score += 800;
        }
        if (lowerClass.includes('fui-chatmymessage') || lowerClass.includes('fui-chatmessage')) score += 750;
        if (lowerClass.includes('fui-chatcontrolmessageitem')) score += 700;
        if (lowerClass.includes('fui-unstable-chatitem')) score += 260;
        if (role === 'listitem' || role === 'article') score += 80;
        if (isLikelyChannelPostCard(element)) score += 620;
        if (element.querySelector('time')) score += 140;
        if (element.querySelector(MESSAGE_META_SELECTOR)) score += 120;
        if (elementHasMessageBodyHint(element)) score += 260;
        if (element.querySelector('[data-tid="subject-line"]')) score += 220;
        if (element.querySelector('[data-tid="post-message-subheader"]')) score += 120;
        if (element.querySelector('img')) score += 40;
        if (element.querySelector(ATTACHMENT_SELECTOR)) score += 60;
        if (element.querySelector('[role="listitem"], [role="article"]')) score -= 120;
        if (isMetadataLikeElement(element)) score -= 500;
        if (element.childElementCount <= 1 && !elementHasMessageBodyHint(element)) score -= 140;

        return score;
    }

    function sortElementsByDomOrder(elements) {
        return elements.sort((a, b) => {
            if (a === b) return 0;
            const relation = a.compareDocumentPosition(b);
            return (relation & Node.DOCUMENT_POSITION_FOLLOWING) ? -1 : 1;
        });
    }

    function normalizeMessageCandidates(candidates, container) {
        const promoted = [];
        const seen = new Set();

        for (const candidate of candidates) {
            let best = candidate;
            let bestScore = scoreMessageRoot(candidate);
            let current = candidate.parentElement;
            let depth = 0;

            while (current && current !== container && depth < 6) {
                if (isLikelyMessageElement(current)) {
                    const currentScore = scoreMessageRoot(current);
                    if (currentScore >= bestScore + 40) {
                        best = current;
                        bestScore = currentScore;
                    }
                }
                current = current.parentElement;
                depth++;
            }

            if (!seen.has(best)) {
                seen.add(best);
                promoted.push(best);
            }
        }

        const filtered = promoted.filter((element) => {
            const score = scoreMessageRoot(element);
            return !promoted.some((other) =>
                other !== element &&
                other.contains(element) &&
                scoreMessageRoot(other) >= score
            );
        });

        return sortElementsByDomOrder(filtered);
    }

    function scoreMessageTextCandidate(element, message, author, timestamp) {
        if (!element) return -Infinity;
        if (element.closest(ATTACHMENT_SELECTOR)) return -Infinity;
        if (isMetadataLikeElement(element)) return -Infinity;

        const text = getNodeText(element);
        if (!text) return -Infinity;

        const normalizedAuthor = normalizeWhitespace(author).toLowerCase();
        const normalizedTimestamp = normalizeWhitespace(timestamp).toLowerCase();
        const normalizedText = text.toLowerCase();

        if (normalizedAuthor && normalizedText === normalizedAuthor) return -Infinity;
        if (normalizedTimestamp && normalizedText === normalizedTimestamp) return -Infinity;

        const className = typeof element.className === 'string' ? element.className.toLowerCase() : '';
        const dataTid = (element.getAttribute('data-tid') || '').toLowerCase();
        const tagName = (element.tagName || '').toLowerCase();
        const isDirectBodyCandidate = elementLooksLikeMessageBody(element);
        const isControlCandidate = looksLikeMessageControlElement(element);
        const cleanedText = stripKnownMessageMetadata(text, author, timestamp);
        if (!cleanedText) return -Infinity;
        if (isControlCandidate && !isDirectBodyCandidate) return -Infinity;
        if (isShortStatusOnlyText(cleanedText) && !isDirectBodyCandidate) return -Infinity;

        let score = Math.min(cleanedText.length, 280);

        if (tagName === 'p' || tagName === 'li' || tagName === 'blockquote') score += 280;
        if (tagName === 'h1' || tagName === 'h2' || tagName === 'h3' || tagName === 'h4') score += 320;
        if (isDirectBodyCandidate) {
            score += 360;
        }
        if (element.matches('div[dir="auto"]')) score += 160;
        if (dataTid.includes('subject-line')) score += 260;
        if (isLikelyChannelPostCard(message) && element.parentElement === message) score += 80;
        if (element.querySelector(MESSAGE_META_SELECTOR)) score -= 240;
        if (normalizedAuthor && normalizedText.includes(normalizedAuthor)) score -= 120;
        if (normalizedTimestamp && normalizedText.includes(normalizedTimestamp)) score -= 100;
        if (element === message) score -= 180;
        if (element.childElementCount <= 4) score += 60;
        if (element.childElementCount > 8) score -= Math.min(element.childElementCount * 10, 100);
        if (isControlCandidate) score -= 260;

        // Text living inside a nested message root (a quoted/forwarded message
        // embedded in this one) is the quote, not this message's own body.
        const NESTED_ROOT_SELECTOR = 'div[class*="fui-ChatMessage"]:not([class*="__"]), div[class*="fui-ChatMyMessage"]:not([class*="__"]), [data-tid*="quoted"], [class*="quoted"]';
        const nearestRoot = element.closest(NESTED_ROOT_SELECTOR);
        if (nearestRoot && nearestRoot !== message && message.contains(nearestRoot)) {
            const outerRoot = nearestRoot.parentElement && nearestRoot.parentElement.closest(NESTED_ROOT_SELECTOR);
            const isQuoteMarker = /quoted/i.test(nearestRoot.getAttribute('data-tid') || '') ||
                /quoted/i.test(typeof nearestRoot.className === 'string' ? nearestRoot.className : '');
            if (isQuoteMarker || (outerRoot && outerRoot !== nearestRoot && message.contains(outerRoot))) {
                score -= 400;
            }
        }

        return score;
    }

    function extractLeafTextFragments(message, author, timestamp) {
        const fragments = [];
        const candidates = message.querySelectorAll('p, blockquote, li, div[dir="auto"], span, div');

        for (const element of candidates) {
            if (element.closest(ATTACHMENT_SELECTOR)) continue;
            if (element.closest('[data-tid="post-message-subheader"], [data-tid="channel-message-reaction-summary"], [data-tid="status-bar"]')) continue;
            if (isMetadataLikeElement(element)) continue;
            if (looksLikeMessageControlElement(element) && !elementLooksLikeMessageBody(element)) continue;
            if (element.querySelector(MESSAGE_META_SELECTOR)) continue;
            if (elementHasMessageBodyHint(element) && !elementLooksLikeMessageBody(element)) continue;

            const cleanedText = stripKnownMessageMetadata(getNodeText(element), author, timestamp);
            if (!cleanedText) continue;
            if (isShortStatusOnlyText(cleanedText) && !elementLooksLikeMessageBody(element)) continue;
            if (fragments.some(existing => existing === cleanedText || existing.includes(cleanedText) || cleanedText.includes(existing))) {
                continue;
            }
            fragments.push(cleanedText);
        }

        return fragments.join('\n').trim();
    }

    function shouldUseChannelFragmentExtraction(message) {
        if (isLikelyChannelPostCard(message)) return true;
        // Split thread units and post cards wrapped in a plain chat-pane-item:
        // the winner-takes-all text scorer would keep only the title OR the body,
        // so anything carrying channel-post markers uses fragment extraction.
        if (message.__isChannelThreadUnit === true) return true;
        try {
            return message.querySelector('[data-tid="subject-line"], [data-tid="post-message-subheader"]') !== null;
        } catch {
            return false;
        }
    }

    function extractMessageContent(message, author, timestamp) {
        if (shouldUseChannelFragmentExtraction(message)) {
            const fragments = [];
            const normalizedFragments = [];
            const replaceFragmentAt = (index, value, normalized) => {
                fragments[index] = value;
                normalizedFragments[index] = normalized;
            };
            const addFragment = (value) => {
                const cleaned = normalizeChannelPostFragment(value, author, timestamp);
                if (!cleaned || isShortStatusOnlyText(cleaned)) return;
                const normalized = cleaned.toLowerCase();

                for (let i = 0; i < normalizedFragments.length; i++) {
                    const existing = normalizedFragments[i];
                    if (!existing) continue;
                    if (existing === normalized || existing.includes(normalized)) {
                        return;
                    }
                    if (normalized.includes(existing)) {
                        replaceFragmentAt(i, cleaned, normalized);
                        return;
                    }
                }

                fragments.push(cleaned);
                normalizedFragments.push(normalized);
            };

            addFragment(message.querySelector('[data-tid="subject-line"]')?.textContent || '');

            const scopedNodes = message.querySelectorAll(
                '[data-tid*="message-body"], [data-tid*="message-content"], [data-tid*="post-body"], [data-tid*="post-content"], div[dir="auto"], p, blockquote, li, [role="heading"], h1, h2, h3, h4'
            );
            scopedNodes.forEach((node) => {
                if (!node || node === message) return;
                if (node.closest('[data-tid="post-message-subheader"], [data-tid="channel-message-reaction-summary"], [data-tid="status-bar"]')) return;
                if (looksLikeMessageControlElement(node) && !elementLooksLikeMessageBody(node)) return;
                if (node.querySelector('[data-tid="channel-message-reaction-summary"], [data-tid="status-bar"]') &&
                    !node.matches('p, blockquote, li, div[dir="auto"], [role="heading"], h1, h2, h3, h4')) {
                    return;
                }
                if (node.querySelector('p, blockquote, li, div[dir="auto"]') &&
                    !node.matches('p, blockquote, li, div[dir="auto"], [role="heading"], h1, h2, h3, h4')) {
                    return;
                }
                addFragment(getNodeText(node));
            });

            const combined = fragments.join('\n').trim();
            if (combined) return combined;

            const leafCombined = extractLeafTextFragments(message, author, timestamp)
                .split('\n')
                .map(fragment => normalizeChannelPostFragment(fragment, author, timestamp))
                .filter(Boolean)
                .join('\n')
                .trim();
            if (leafCombined) return leafCombined;
        }

        const candidates = [];
        const seen = new Set();

        const addCandidate = (element) => {
            if (!element || seen.has(element) || !message.contains(element)) return;
            seen.add(element);
            candidates.push(element);
        };

        addCandidate(message);
        message.querySelectorAll(MESSAGE_TEXT_SELECTOR).forEach(addCandidate);
        message.querySelectorAll('div, p, span, blockquote, li').forEach(addCandidate);

        let bestText = '';
        let bestScore = -Infinity;
        let bestElement = null;

        for (const candidate of candidates) {
            const score = scoreMessageTextCandidate(candidate, message, author, timestamp);
            if (!Number.isFinite(score)) continue;

            const cleanedText = stripKnownMessageMetadata(getNodeText(candidate), author, timestamp);
            if (!cleanedText) continue;

            if (score > bestScore) {
                bestScore = score;
                bestText = cleanedText;
                bestElement = candidate;
            }
        }

        // A message written as several paragraphs or list items must keep all of
        // them, not only the highest-scoring block, so widen to the body container.
        if (bestElement && bestElement.matches('p, li, blockquote, h1, h2, h3, h4')) {
            const blockContainer = bestElement.parentElement?.closest('div[dir="auto"], [class*="__body"], [id^="content-"]');
            if (blockContainer && blockContainer !== message && message.contains(blockContainer) &&
                !blockContainer.querySelector(MESSAGE_META_SELECTOR) &&
                !blockContainer.querySelector('[data-tid*="quoted"], [class*="quoted"]')) {
                const containerText = stripKnownMessageMetadata(getNodeText(blockContainer), author, timestamp);
                if (containerText.length > bestText.length && containerText.includes(bestText)) {
                    bestText = containerText;
                }
            }
        }

        if (bestText) return bestText;

        const messageClone = message.cloneNode(true);
        stripMessageNoise(messageClone);
        const clonedFallback = stripKnownMessageMetadata(getNodeText(messageClone), author, timestamp);
        if (clonedFallback && !isShortStatusOnlyText(clonedFallback)) return clonedFallback;

        return extractLeafTextFragments(message, author, timestamp);
    }

    // File download progress tracking
    let totalFilesToDownload = 0;
    let filesDownloaded = 0;
    let downloadStartTime = null;

    // Helper to send file download progress update
    function sendFileDownloadProgress() {
        if (totalFilesToDownload > 0) {
            const elapsed = downloadStartTime ? (Date.now() - downloadStartTime) / 1000 : 0;
            const avgTimePerFile = filesDownloaded > 0 ? elapsed / filesDownloaded : 2; // estimate 2 seconds per file
            const remaining = totalFilesToDownload - filesDownloaded;
            const estimatedRemaining = Math.ceil(remaining * avgTimePerFile);

            chrome.runtime.sendMessage({
                action: "fileDownloadProgress",
                downloaded: filesDownloaded,
                total: totalFilesToDownload,
                estimatedSeconds: estimatedRemaining
            });
        }
    }

    // Scroll through entire chat to extract all text
    console.log(`Extracting all chat text by scrolling... (turboMode: ${turboMode}, longChatMode: ${longChatMode}, incrementalFrom: ${Number.isFinite(incrementalFromTimestamp) ? new Date(incrementalFromTimestamp).toISOString() : 'off'})`);

    // Save original scroll position for restoration later
    const originalScrollTop = chatContainer.scrollTop;

    // Determine wait times based on mode
    // Long chat mode uses TURBO speed - fast scrolling, but scrollToTop already handled the long waits
    let shortWait;
    if (longChatMode) {
        shortWait = 500;
    } else if (turboMode) {
        // Text-only backups still need enough time for Teams' virtualized list to render.
        shortWait = 250;
    } else {
        shortWait = 500;
    }

    const extractionSettleOptions = longChatMode
        ? { intervalMs: 850, maxRounds: 18, stableRoundsRequired: 4 }
        : turboMode
            ? { intervalMs: 320, maxRounds: 9, stableRoundsRequired: 2 }
            : { intervalMs: 600, maxRounds: 9, stableRoundsRequired: 2 };

    if (startPosition === null) {
        // Do not bounce to the bottom here. On very long chats that can unload the
        // older chunk we just loaded and start extraction from a much newer slice.
        const settleResult = await waitForViewportToSettle(chatContainer, {
            ...extractionSettleOptions,
            requireTop: true,
            nudge: longChatMode ? sendTopLoadNudges : null,
            operationId
        });
        if (settleResult?.cancelled || isOperationCancelled(operationId)) {
            return {
                messages: chatText,
                errors: extractionErrors
            };
        }
    } else {
        // User stopped scrolling early - start from the specified position
        console.log(`Starting extraction from position ${startPosition} (user stopped scrolling early)`);
        chatContainer.scrollTop = startPosition;
        const settleResult = await waitForViewportToSettle(chatContainer, {
            ...extractionSettleOptions,
            operationId
        });
        if (settleResult?.cancelled || isOperationCancelled(operationId)) {
            return {
                messages: chatText,
                errors: extractionErrors
            };
        }
    }

    let lastMessageCount = 0;
    let retryCount = 0;
    const maxRetries = longChatMode ? 120 : 16;

    // Helper function to convert image to base64 data URL
    // Handles blob URLs by fetching the blob directly
    async function imageToDataUrl(img, maxWidth = 400, maxHeight = 300, spriteInfo = null) {
        try {
            // Skip if image not loaded or broken
            if (!img.complete || img.naturalWidth === 0 || !img.src) {
                return null;
            }

            const imgSrc = img.src;

            // Animated GIFs must keep their original bytes; re-encoding through a
            // canvas would freeze them to a single frame.
            const MAX_ANIMATED_BYTES = 5 * 1024 * 1024;

            if (imgSrc.startsWith('data:image/gif')) {
                return imgSrc;
            }

            const blobToDataUrl = (blob) => new Promise((resolve) => {
                const reader = new FileReader();
                reader.onloadend = () => resolve(reader.result);
                reader.onerror = () => {
                    console.error('FileReader error while converting image blob');
                    resolve(null);
                };
                reader.readAsDataURL(blob);
            });

            // For blob URLs, fetch the blob and convert to base64
            if (imgSrc.startsWith('blob:')) {
                try {
                    const response = await fetch(imgSrc);
                    const blob = await response.blob();
                    return await blobToDataUrl(blob);
                } catch (e) {
                    console.error('Error fetching blob URL:', e);
                    return null;
                }
            }

            // http(s) GIFs (e.g. Giphy in Teams): fetch the raw bytes so the
            // animation survives. Falls back to the static canvas path below.
            const urlPath = imgSrc.split(/[?#]/)[0].toLowerCase();
            if (urlPath.endsWith('.gif') || imgSrc.toLowerCase().includes('giphy')) {
                try {
                    const response = await fetch(imgSrc);
                    if (response.ok) {
                        const blob = await response.blob();
                        if (blob.type.includes('gif') && blob.size <= MAX_ANIMATED_BYTES) {
                            const dataUrl = await blobToDataUrl(blob);
                            if (dataUrl) return dataUrl;
                        }
                    }
                } catch (e) {
                    console.warn('Could not fetch GIF bytes, capturing static frame instead:', e);
                }
            }

            // For regular URLs, try canvas approach with crossOrigin
            // Create a new image with crossOrigin to avoid tainting
            return new Promise((resolve) => {
                const tempImg = new Image();
                tempImg.crossOrigin = 'anonymous';

                tempImg.onload = () => {
                    try {
                        // Calculate scaled dimensions (preserve aspect ratio)
                        let width = tempImg.naturalWidth;
                        let height = tempImg.naturalHeight;

                        // Emoji sprite strips keep their natural size (frames are small)
                        // and stay PNG so transparency and the frame grid survive.
                        if (spriteInfo) {
                            const canvas = document.createElement('canvas');
                            canvas.width = width;
                            canvas.height = height;
                            const ctx = canvas.getContext('2d');
                            ctx.drawImage(tempImg, 0, 0);
                            resolve(canvas.toDataURL('image/png'));
                            return;
                        }

                        if (width > maxWidth) {
                            height = (height * maxWidth) / width;
                            width = maxWidth;
                        }
                        if (height > maxHeight) {
                            width = (width * maxHeight) / height;
                            height = maxHeight;
                        }

                        const canvas = document.createElement('canvas');
                        canvas.width = width;
                        canvas.height = height;
                        const ctx = canvas.getContext('2d');
                        ctx.drawImage(tempImg, 0, 0, width, height);

                        resolve(canvas.toDataURL('image/jpeg', 0.7));
                    } catch (e) {
                        console.error('Canvas error:', e);
                        resolve(null);
                    }
                };

                tempImg.onerror = () => {
                    console.error('Failed to load image with crossOrigin');
                    resolve(null);
                };

                // Set timeout for image loading
                setTimeout(() => resolve(null), 5000);

                tempImg.src = imgSrc;
            });
        } catch (e) {
            console.error('Error converting image to data URL:', e);
            return null;
        }
    }

    // Helper function to download file from SharePoint URL and convert to base64
    // Returns { dataUrl, mimeType, size } or { error } if failed
    // Includes 30-second timeout to prevent hanging on slow/failed downloads
    async function fileToDataUrl(sharePointUrl, maxSizeMB = 5) {
        if (!sharePointUrl) {
            return { error: 'No URL provided' };
        }

        // Create AbortController for timeout
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 30000); // 30 second timeout

        try {
            console.log('Attempting to download file from:', sharePointUrl);

            // Try to fetch the file with credentials (user should be logged into SharePoint)
            const response = await fetch(sharePointUrl, {
                method: 'GET',
                credentials: 'include', // Include cookies for authentication
                redirect: 'follow',
                signal: controller.signal // Add abort signal
            });

            clearTimeout(timeoutId); // Clear timeout if fetch succeeded

            if (!response.ok) {
                console.error('File fetch failed:', response.status, response.statusText);
                return { error: `HTTP ${response.status}: ${response.statusText}` };
            }

            // Get content length if available
            const contentLength = response.headers.get('content-length');
            const sizeBytes = contentLength ? parseInt(contentLength) : 0;
            const sizeMB = sizeBytes / (1024 * 1024);

            // Check if file exceeds max size
            if (sizeBytes > 0 && sizeMB > maxSizeMB) {
                return {
                    error: 'File too large',
                    tooLarge: true,
                    sizeMB: sizeMB.toFixed(2)
                };
            }

            // Get the blob
            const blob = await response.blob();
            const actualSizeMB = blob.size / (1024 * 1024);

            // Double check size after getting blob
            if (actualSizeMB > maxSizeMB) {
                return {
                    error: 'File too large',
                    tooLarge: true,
                    sizeMB: actualSizeMB.toFixed(2)
                };
            }

            // Convert to base64
            return new Promise((resolve) => {
                const reader = new FileReader();
                reader.onloadend = () => {
                    resolve({
                        dataUrl: reader.result,
                        mimeType: blob.type || 'application/octet-stream',
                        sizeMB: actualSizeMB.toFixed(2)
                    });
                };
                reader.onerror = () => {
                    console.error('FileReader error for file');
                    resolve({ error: 'Failed to read file data' });
                };
                reader.readAsDataURL(blob);
            });
        } catch (e) {
            clearTimeout(timeoutId); // Clear timeout on error
            if (e.name === 'AbortError') {
                console.error('File download timed out after 30 seconds');
                return { error: 'Download timed out', timeout: true };
            }
            console.error('Error downloading file:', e);
            return { error: e.message || 'Download failed' };
        }
    }

    const messageSelectors = [
        '[data-tid="chat-pane-item"]',
        '[data-tid="message-pane-item"]',
        '[data-tid*="chat-pane-item"]',
        '[data-tid*="message-pane-item"]',
        '[data-tid*="message-item"]',
        '[data-tid*="post-item"]',
        '[data-tid*="reply-item"]',
        '[data-tid*="thread-item"]',
        'div[class*="fui-ChatMyMessage"]:not([class*="__"])',
        'div[class*="fui-ChatMessage"]:not([class*="__"])',
        'div[class*="fui-ChatControlMessageItem"]',
        'div[class*="fui-unstable-ChatItem"]',
        '[role="listitem"]',
        '[role="article"]'
    ];
    const nestedMessageContainerSelectors = [
        '[data-tid="message-pane-list-viewport"]',
        '[data-tid*="message-pane-list"]',
        '[data-tid="channel-pane-viewport"]',
        '[data-tid*="channel-pane-list"]',
        '[data-tid*="post-list"]',
        '[role="feed"]',
        '[role="log"]',
        '[class*="virtualized"]',
        '[class*="virtual-list"]'
    ];

    const isLikelyMessageElement = (element) => {
        const className = typeof element.className === 'string' ? element.className : '';
        const dataTid = (element.getAttribute('data-tid') || '').toLowerCase();
        const role = (element.getAttribute('role') || '').toLowerCase();
        const text = (element.textContent || '').trim();
        const hasHeadingElement = element.querySelector('[role="heading"]') !== null;
        const hasAvatarElement = element.querySelector('[class*="Avatar"], [class*="avatar"], [data-tid*="persona"]') !== null;
        const hasRichTextBlocks = element.querySelector('p, blockquote, ul, ol, li') !== null;
        const hasTimestampElement = element.querySelector('time') !== null;
        const hasAuthorLikeElement = element.querySelector('[data-tid="message-author-name"], [class*="__author"]') !== null ||
                                    element.querySelector('[class*="Avatar"], [class*="avatar"], [data-tid*="persona"]') !== null;

        if (!hasTimestampElement && !hasAuthorLikeElement && isLikelyDateDividerText(text)) {
            return false;
        }

        const hasKnownRootClass = className.includes('fui-ChatMyMessage') ||
                                 className.includes('fui-ChatMessage') ||
                                 className.includes('fui-ChatControlMessageItem') ||
                                 className.includes('fui-unstable-ChatItem');

        const hasMessageIndicators = hasTimestampElement ||
                                   element.querySelector('[data-tid="message-author-name"]') !== null ||
                                   element.querySelector('[class*="__author"]') !== null ||
                                   elementHasMessageBodyHint(element) ||
                                   element.querySelector('div[dir="auto"]') !== null;

        const dataTidLooksRight = dataTid.includes('chat-pane-item') ||
                                  dataTid.includes('message-pane-item') ||
                                  dataTid.includes('message-item') ||
                                  dataTid.includes('post-item') ||
                                  dataTid.includes('reply-item') ||
                                  dataTid.includes('thread-item');

        const roleLooksRight = role === 'listitem' || role === 'article';
        const isPostCard = isLikelyChannelPostCard(element);

        if (isMetadataLikeElement(element)) return false;

        return hasKnownRootClass ||
               dataTidLooksRight ||
               isPostCard ||
               (roleLooksRight && text.length > 0 && (hasMessageIndicators || hasHeadingElement || hasAvatarElement || hasRichTextBlocks));
    };

    // A channel thread card holds the root post plus every reply. Extracting the
    // card as one element merges the whole thread into a single message and the
    // text scorer then keeps only one fragment of it. Split cards back into
    // their per-message units (each unit carries its own <time>).
    const THREAD_UNIT_SELECTOR = [
        '[data-tid*="reply-item"]',
        '[data-tid*="thread-item"]',
        '[data-tid*="post-item"]',
        'div[class*="fui-ChatMyMessage"]:not([class*="__"])',
        'div[class*="fui-ChatMessage"]:not([class*="__"])',
        '[role="article"]',
        '[role="listitem"]'
    ].join(', ');

    function findThreadMessageUnits(card) {
        let units;
        try {
            units = Array.from(card.querySelectorAll(THREAD_UNIT_SELECTOR));
        } catch {
            return [];
        }

        units = units.filter(unit =>
            unit !== card &&
            unit.querySelector('time') !== null &&
            !isMetadataLikeElement(unit));

        // The thread card (or an inner wrapper) can itself match the unit selector
        // and contains every real unit — drop container-like matches that wrap two
        // or more other units.
        const containers = new Set(units.filter(unit =>
            units.filter(other => other !== unit && unit.contains(other)).length >= 2));
        const messageUnits = units.filter(unit => !containers.has(unit));

        // Keep only outermost units — anything nested inside another unit is that
        // message's own markup (e.g. a quote), not a sibling message.
        const outer = messageUnits.filter(unit =>
            !messageUnits.some(other => other !== unit && other.contains(unit)));

        return outer.length >= 2 ? outer : [];
    }

    const looksLikeThreadCard = (element) => {
        if (isLikelyChannelPostCard(element)) return true;
        try {
            // Thread cards are often wrapped in a plain chat-pane-item; look for
            // channel-post markers anywhere inside before splitting.
            return element.querySelector(
                '[data-tid="post-message-subheader"], [data-tid="subject-line"], ' +
                '[data-tid*="reply-item"], [data-tid*="thread-item"], [data-tid*="post-item"]'
            ) !== null;
        } catch {
            return false;
        }
    };

    const expandThreadCards = (elements) => {
        const expanded = [];
        for (const element of elements) {
            const units = looksLikeThreadCard(element) ? findThreadMessageUnits(element) : [];
            if (units.length >= 2) {
                // Post titles (subject-line) can sit on the card outside the root
                // post unit; remember the card so the title is not lost.
                const subjectEl = element.querySelector('[data-tid="subject-line"]');
                if (subjectEl && !units.some(unit => unit.contains(subjectEl))) {
                    units[0].__threadSubjectText = getNodeText(subjectEl);
                }
                units.forEach(unit => { unit.__isChannelThreadUnit = true; });
                expanded.push(...units);
            } else {
                expanded.push(element);
            }
        }
        return expanded;
    };

    // ---- Timestamp-anchored adaptive fallbacks -------------------------------
    // Teams builds ship different data-tids/classes per ring and tenant, so
    // selector lists alone cannot cover every variant. The one stable invariant
    // is that every rendered message carries a <time> element. These helpers
    // only refine or augment the selector-based result when it clearly misses
    // or merges timestamps, so known-good DOMs behave exactly as before.
    const QUOTE_SCOPE_SELECTOR = 'blockquote, [class*="quot" i], [data-tid*="quot" i], [class*="reference" i]';

    function getMessageTimeElements(scope) {
        let times;
        try {
            times = Array.from(scope.querySelectorAll('time'));
        } catch {
            return [];
        }
        return times.filter(timeEl =>
            !timeEl.closest(QUOTE_SCOPE_SELECTOR) &&
            !timeEl.closest('[contenteditable="true"], [data-tid*="compose" i], [data-tid*="editor" i]'));
    }

    // Climb from a timestamp to the largest ancestor that still contains only
    // this one timestamp: that ancestor is the message's own root.
    function deriveUnitFromTime(timeEl, boundary) {
        let node = timeEl;
        while (node.parentElement && node.parentElement !== boundary) {
            if (getMessageTimeElements(node.parentElement).length > 1) break;
            node = node.parentElement;
        }
        return node === timeEl ? timeEl.parentElement : node;
    }

    // A candidate holding 3+ own timestamps is a merged multi-message block
    // (unknown-variant thread card). Split it along its timestamps. The >=3
    // threshold keeps quoted messages and meeting cards (2 times) intact.
    const splitMultiTimestampCandidates = (candidates, container) => {
        const result = [];
        for (const element of candidates) {
            const times = getMessageTimeElements(element);
            if (times.length >= 3) {
                const units = [];
                const seenUnits = new Set();
                for (const timeEl of times) {
                    const unit = deriveUnitFromTime(timeEl, element);
                    if (!unit || unit === element || seenUnits.has(unit) || isMetadataLikeElement(unit)) continue;
                    seenUnits.add(unit);
                    units.push(unit);
                }
                if (units.length >= 2) {
                    const subjectEl = element.querySelector('[data-tid="subject-line"]');
                    const sortedUnits = sortElementsByDomOrder(units);
                    if (subjectEl && !sortedUnits.some(unit => unit.contains(subjectEl))) {
                        sortedUnits[0].__threadSubjectText = getNodeText(subjectEl);
                    }
                    sortedUnits.forEach(unit => { unit.__isChannelThreadUnit = true; });
                    result.push(...sortedUnits);
                    console.log(`splitMultiTimestampCandidates: split a ${times.length}-timestamp block into ${sortedUnits.length} messages`);
                    continue;
                }
            }
            result.push(element);
        }
        return result;
    };

    // Timestamps not covered by any candidate mean the selector family missed
    // part (or all) of the feed — derive message units directly for those.
    const addUncoveredTimeUnits = (candidates, container) => {
        const containerTimes = getMessageTimeElements(container);
        if (containerTimes.length === 0) return candidates;

        const uncovered = containerTimes.filter(timeEl => !candidates.some(c => c.contains(timeEl)));
        if (uncovered.length === 0) return candidates;
        // With a working selector result, ignore the odd stray timestamp
        // (status bars, headers); only augment on a meaningful miss.
        if (candidates.length > 0 && uncovered.length < Math.max(2, containerTimes.length * 0.25)) {
            return candidates;
        }

        const seenUnits = new Set(candidates);
        const added = [];
        for (const timeEl of uncovered) {
            const unit = deriveUnitFromTime(timeEl, container);
            if (!unit || unit === container || seenUnits.has(unit) || isMetadataLikeElement(unit)) continue;
            if (candidates.some(c => c.contains(unit) || unit.contains(c))) continue;
            unit.__isChannelThreadUnit = true;
            seenUnits.add(unit);
            added.push(unit);
        }
        if (added.length === 0) return candidates;
        console.log(`addUncoveredTimeUnits: recovered ${added.length} message(s) missed by selectors`);
        return sortElementsByDomOrder([...candidates, ...added]);
    };

    const refineCandidates = (elements, container) => {
        // Only refine non-empty selector results; when a selector finds nothing,
        // let the loop try the next (possibly correct) selector family. Pure
        // time-anchored discovery runs only after every selector has failed.
        if (elements.length === 0) return elements;
        let refined = expandThreadCards(elements);
        refined = splitMultiTimestampCandidates(refined, container);
        refined = addUncoveredTimeUnits(refined, container);
        return refined;
    };

    const findMessageElements = (container) => {
        for (const selector of messageSelectors) {
            const rawCandidates = Array.from(container.querySelectorAll(selector)).filter(isLikelyMessageElement);
            const candidates = refineCandidates(normalizeMessageCandidates(rawCandidates, container), container);
            if (candidates.length > 0) {
                return { elements: candidates, selector };
            }
        }

        const postCardCandidates = Array.from(container.querySelectorAll('[role="group"]')).filter(isLikelyChannelPostCard);
        const normalizedPostCards = refineCandidates(normalizeMessageCandidates(postCardCandidates, container), container);
        if (normalizedPostCards.length > 0) {
            return { elements: normalizedPostCards, selector: '[role="group"][channel-post-card]' };
        }

        // Last resort for fully unknown markup: build message units purely from
        // timestamps.
        const timeAnchored = addUncoveredTimeUnits([], container);
        if (timeAnchored.length > 0) {
            return { elements: timeAnchored, selector: '_time_anchored' };
        }

        return { elements: [], selector: null };
    };

    const findNestedMessageContainer = (container) => {
        for (const selector of nestedMessageContainerSelectors) {
            const matches = Array.from(container.querySelectorAll(selector));
            for (const match of matches) {
                if (match === container) continue;
                const lookup = findMessageElements(match);
                if (lookup.elements.length > 0) {
                    return { container: match, ...lookup, nestedSelector: selector };
                }
            }
        }
        return null;
    };

    // Helper function to extract messages from current view (async for image handling)
    async function extractVisibleMessages() {
        if (isOperationCancelled(operationId)) {
            return;
        }

        // Teams 2024+ structure (discovered via Chrome DevTools MCP):
        // - All chat items have data-tid="chat-pane-item"
        // - System messages (meeting ended, left chat, invited): class contains "fui-ChatControlMessageItem"
        // - Regular messages: class contains "fui-ChatMyMessage" or "fui-ChatMessage"
        // - Date dividers: class contains "fui-unstable-ChatItem" (we skip these)
        // IMPORTANT: Query all items together to maintain DOM order (chronological order)

        let lookup = findMessageElements(chatContainer);
        let allElements = lookup.elements;

        if (allElements.length === 0) {
            const nestedLookup = findNestedMessageContainer(chatContainer);
            if (nestedLookup) {
                console.warn(`extractVisibleMessages: drilling into nested container via "${nestedLookup.nestedSelector}"`);
                chatContainer = nestedLookup.container;
                extractionErrors.containerDiagnostics = {
                    ...(lastContainerSelectionDiagnostics || {}),
                    nestedRebindSelector: nestedLookup.nestedSelector
                };
                lookup = { elements: nestedLookup.elements, selector: nestedLookup.selector };
                allElements = lookup.elements;
            }
        }

        if (allElements.length === 0) {
            const reboundContainer = getChatContainer(false);
            if (reboundContainer && reboundContainer !== chatContainer) {
                console.warn('extractVisibleMessages: rebinding chat container after empty selector result');
                chatContainer = reboundContainer;
                extractionErrors.containerDiagnostics = lastContainerSelectionDiagnostics;
                lookup = findMessageElements(chatContainer);
                allElements = lookup.elements;
            }
        }

        // STRUCTURAL FALLBACK: When all named selectors fail, walk the DOM tree
        // looking for message-like structures (elements containing <time> and text content).
        // This handles unknown Teams channel variants where our selectors don't match.
        if (allElements.length === 0) {
            console.warn('extractVisibleMessages: all selectors failed, trying structural deep-walk fallback');
            const structuralCandidates = [];
            const timeElements = chatContainer.querySelectorAll('time');
            for (const timeEl of timeElements) {
                // Walk up from each <time> to find the message-level ancestor.
                // Prefer the smallest ancestor that still contains the visible post body/title.
                let candidate = timeEl.parentElement;
                let bestCandidate = null;
                let bestScore = -Infinity;
                for (let depth = 0; candidate && candidate !== chatContainer && depth < 8; depth++) {
                    const text = (candidate.textContent || '').trim();
                    if (text.length > 10 && text.length < 5000) {
                        const headingCount = candidate.querySelectorAll('[role="heading"], h1, h2, h3, h4').length;
                        const bodyHints = candidate.querySelectorAll('div[dir="auto"], p, blockquote, li').length;
                        const timeCount = candidate.querySelectorAll('time').length;
                        const buttonCount = candidate.querySelectorAll('button, [role="button"]').length;
                        const tabCount = candidate.querySelectorAll('[role="tab"]').length;
                        const strippedText = stripKnownMessageMetadata(text, '', timeEl.getAttribute('datetime') || timeEl.textContent || '');

                        let score = Math.min(text.length, 240);
                        score += headingCount * 260;
                        score += bodyHints * 120;
                        score += strippedText.length >= 8 ? 220 : 0;
                        score += timeCount === 1 ? 100 : -Math.max(0, timeCount - 1) * 70;
                        score -= buttonCount * 45;
                        score -= tabCount * 220;
                        score -= depth * 18;

                        if (score > bestScore) {
                            bestScore = score;
                            bestCandidate = candidate;
                        }
                    }
                    // Stop if we hit an element that looks like a message wrapper
                    const tid = (candidate.getAttribute('data-tid') || '').toLowerCase();
                    const cls = typeof candidate.className === 'string' ? candidate.className : '';
                    if (tid.includes('item') || tid.includes('message') || tid.includes('post') || tid.includes('thread') ||
                        cls.includes('fui-unstable-ChatItem') || cls.includes('fui-ChatMessage') || cls.includes('fui-ChatMyMessage')) {
                        bestCandidate = candidate;
                        break;
                    }
                    candidate = candidate.parentElement;
                }
                if (bestCandidate && !structuralCandidates.includes(bestCandidate)) {
                    structuralCandidates.push(bestCandidate);
                }
            }
            if (structuralCandidates.length > 0) {
                console.log(`extractVisibleMessages: structural fallback found ${structuralCandidates.length} message-like elements`);
                allElements = normalizeMessageCandidates(structuralCandidates, chatContainer);
                lookup = { elements: allElements, selector: '_structural_time_walk' };
            }
        }

        if (allElements.length === 0) {
            extractionErrors.selectorMismatch = true;
            extractionErrors.domDiagnostics = captureDomDiagnostics();
            return;
        }

        extractionErrors.selectorMismatch = false;
        extractionErrors.selectorUsed = lookup.selector;
        if (lookup.selector !== '[data-tid="chat-pane-item"]') {
            console.log(`extractVisibleMessages: using fallback selector "${lookup.selector}" (${allElements.length} elements)`);
        }
        detectedMessageElements += allElements.length;

        // Let this screen's message images finish loading together (one 3 s cap)
        // instead of waiting up to 3 s per message. Avatars are skipped.
        await waitForImages(Array.from(allElements).flatMap(element => getPendingContentImages(element)), operationId);

        // Track current date header for system messages that don't have timestamps
        let currentDateHeader = '';

        for (const element of allElements) {
            if (isOperationCancelled(operationId)) {
                return;
            }

            const className = element.className || '';
            const textContent = element.textContent || '';
            const normalizedTextContent = normalizeWhitespace(textContent);
            const controlMessageAncestor = typeof element.closest === 'function'
                ? element.closest('div[class*="fui-ChatControlMessageItem"]')
                : null;
            if (controlMessageAncestor && controlMessageAncestor !== element) {
                continue;
            }

            // Check if this is a date header (e.g., "Friday, August 30, 2024 1:15 AM")
            // LANGUAGE-AGNOSTIC detection: Date headers are dividers with NO author/body elements
            // They typically have fui-unstable-ChatItem class and short text with date/time pattern
            const hasNoAuthor = !element.querySelector('[data-tid="message-author-name"], [class*="__author"], [class*="Avatar"], [class*="avatar"]');
            const isDateOnlyHeaderText = isLikelyDateDividerText(normalizedTextContent);
            const hasNoBody = isDateOnlyHeaderText || !elementHasMessageBodyHint(element);
            const isShortText = textContent.length < 100;  // Date headers are usually short
            const hasDatePattern = isDateOnlyHeaderText ||
                                   /\d{1,4}[-\/\.]\d{1,2}[-\/\.]\d{1,4}|\d{1,2}:\d{2}/.test(textContent) || // Date or time pattern
                                   textContent.match(/\d{4}/);  // Year like 2024

            // Date headers: short text with date pattern and no author/body.
            // For structural fallback (unknown variants), also treat any short date-only node as a header.
            const isDateHeader = (
                isDateOnlyHeaderText &&
                hasNoAuthor &&
                !element.querySelector('time')
            ) || (
                className.includes('fui-unstable-ChatItem') &&
                !className.includes('f1locze1') &&
                !className.includes('fui-ChatControlMessageItem') &&
                hasNoAuthor && hasNoBody && isShortText && hasDatePattern
            ) || (
                // Structural fallback: element has no <time> tag, no author, only date-like text
                hasNoAuthor && hasNoBody && isShortText && hasDatePattern &&
                !element.querySelector('time') && textContent.length < 60
            );

            if (isDateHeader) {
                // Store this date header for subsequent messages
                currentDateHeader = normalizedTextContent;
                continue; // Skip date headers, don't extract them as messages
            }

            // Check if it's a system message (fui-ChatControlMessageItem)
            const isSystemMessage = className.includes('fui-ChatControlMessageItem');

            // Check if it's a regular message - multiple detection methods
            // LANGUAGE-AGNOSTIC: Use structural elements, not text patterns
            const isMyMessage = className.match(/\bfui-ChatMyMessage\b/) && !className.match(/fui-ChatMyMessage__/);
            const isOtherMessage = className.match(/\bfui-ChatMessage\b/) && !className.match(/fui-ChatMessage__/) && !className.match(/fui-ChatMyMessage/);
            const isRegularMessage = isMyMessage || isOtherMessage;

            // Structural detection: check for message-like elements
            // Works for ALL languages - no text pattern dependencies
            const hasTimeElement = element.querySelector('time') !== null;
            const hasAuthorElement = element.querySelector('[data-tid="message-author-name"]') !== null ||
                                    element.querySelector('[class*="__author"]') !== null ||
                                    element.querySelector('[class*="Avatar"], [class*="avatar"]') !== null;
            const hasBodyElement = elementHasMessageBodyHint(element) ||
                                  element.querySelector('div[dir="auto"]') !== null;

            // Check for fui-unstable-ChatItem with message indicators (language-agnostic)
            const isUnstableChatItemMessage = className.includes('fui-unstable-ChatItem') &&
                                               !isSystemMessage &&
                                               !isDateHeader &&
                                               (hasTimeElement || hasAuthorElement || hasBodyElement);

            if (isSystemMessage) {
                // Extract system message
                try {
                    // Get the heading text (the main system message content)
                    const headingEl = element.querySelector('[role="heading"]');
                    const content = headingEl ? headingEl.textContent.trim() : textContent.trim();

                    // Try to get timestamp from multiple sources:
                    let timestamp = '';

                    // 1. Try <time> element first
                    const timeEl = element.querySelector('time');
                    if (timeEl) {
                        timestamp = timeEl.getAttribute('datetime') || timeEl.getAttribute('title') || timeEl.textContent.trim();
                    }

                    // 2. Try to extract date from the message text (e.g., "8/30/2024 2:53 AM")
                    if (!timestamp) {
                        const dateMatch = textContent.match(/(\d{1,2}\/\d{1,2}\/\d{4}\s+\d{1,2}:\d{2}\s*(AM|PM)?)/i);
                        if (dateMatch) {
                            timestamp = dateMatch[1];
                        }
                    }

                    // 2b. Use nearby heading/date context for control messages that do not carry their own <time>.
                    if (!timestamp) {
                        timestamp = inferTimestampContextForElement(element, currentDateHeader);
                    }

                    // 3. Fall back to the current date header
                    if (!timestamp && currentDateHeader) {
                        timestamp = currentDateHeader;
                    }

                    // Only add if has meaningful content and not a duplicate
                    if (content && content.length > 0) {
                        if (!isTimestampInRange(timestamp)) {
                            continue;
                        }

                        if (!timestamp && !currentDateHeader && isRelaxedMeetingControlContent(content)) {
                            console.log('Skipping duplicate-prone meeting control message without timestamp context:', content.substring(0, 80));
                            continue;
                        }

                        const messageKey = createMessageKey('system', '', timestamp, content);
                        if (!seenMessageKeys.has(messageKey)) {
                            seenMessageKeys.add(messageKey);
                            await queueExtractedMessage({
                                type: 'system',
                                content: content,
                                timestamp: timestamp
                            });
                            console.log('Extracted system message:', content.substring(0, 50), 'timestamp:', timestamp);
                        }
                    }
                } catch (error) {
                    console.error('Error extracting system message:', error);
                }
                continue; // Skip to next element
            }

            // Ultimate fallback: any chat-pane-item with structural message indicators
            // LANGUAGE-AGNOSTIC: Works for ALL Teams versions regardless of language
            const isLikelyMessage = !isSystemMessage && !isDateHeader &&
                                    (hasTimeElement || hasAuthorElement || hasBodyElement) &&
                                    textContent.length > 10;

            // Skip if not a regular message (e.g., date dividers)
            if (!isRegularMessage && !isUnstableChatItemMessage && !isLikelyMessage) {
                continue;
            }

            // Regular message extraction
            const message = element;
            try {
                // Try multiple author selectors (ordered by specificity)
                let authorElement = message.querySelector('[data-tid="message-author-name"]');
                if (!authorElement) authorElement = message.querySelector('div[class*="fui-ChatMyMessage__author"]');
                if (!authorElement) authorElement = message.querySelector('div[class*="fui-ChatMessage__author"]');
                if (!authorElement) authorElement = message.querySelector('[class*="__author"]');
                // Fallback: look for span/div near avatar or at start of message
                if (!authorElement) authorElement = message.querySelector('[class*="Avatar"] + span, [class*="avatar"] + span');
                if (!authorElement) authorElement = message.querySelector('[class*="personaName"], [class*="displayName"]');
                // Korean Teams fallback: first bold or emphasized text near time
                if (!authorElement) {
                    const timeEl = message.querySelector('time');
                    if (timeEl && timeEl.parentElement) {
                        authorElement = timeEl.parentElement.querySelector('span:first-child, div:first-child');
                    }
                }
                const author = authorElement ? authorElement.textContent.trim() : 'Unknown Author';

                // Try multiple timestamp selectors
                let timestampElement = message.querySelector('time');
                if (!timestampElement) timestampElement = message.querySelector('[class*="__timestamp"]');
                // Prefer datetime attribute (ISO format like "2025-12-11T07:57:10.436Z") over title ("Today at 4:57 PM")
                const timestamp = timestampElement ?
                    (timestampElement.getAttribute('datetime') ||
                     timestampElement.getAttribute('title') ||
                     timestampElement.textContent.trim()) : '';

                if (!isTimestampInRange(timestamp)) {
                    continue;
                }

                let content = extractMessageContent(message, author, timestamp);

                // Root post of a split thread card: re-attach the post title when it
                // lives on the card outside this unit.
                const threadSubject = typeof message.__threadSubjectText === 'string'
                    ? message.__threadSubjectText.trim()
                    : '';
                if (threadSubject) {
                    delete message.__threadSubjectText;
                    if (!content) {
                        content = threadSubject;
                    } else if (!content.toLowerCase().includes(threadSubject.toLowerCase())) {
                        content = `${threadSubject}\n${content}`;
                    }
                }
                const attachmentMetadata = extractAttachmentMetadata(message);

                // Extract images from the message (parallel for better performance)
                const imgElements = Array.from(message.querySelectorAll('img'));
                const imagePromises = [];
                const imageReferences = [];
                for (const img of imgElements) {
                    // Skip avatars, emojis, and UI icons (small images)
                    const isAvatar = img.closest('[class*="avatar"]') || img.closest('[class*="Avatar"]');
                    const spriteInfo = getAnimatedEmojiSpriteInfo(img);
                    // Animated emoticon sprites rendered small are inline emojis: their
                    // alt text is already inlined into the message text by getNodeText.
                    // Larger standalone ones are captured with sprite metadata so the
                    // viewer can play them instead of showing the unrolled strip.
                    const renderedWidth = img.getBoundingClientRect().width || img.width || 0;
                    const isInlineEmojiSprite = spriteInfo && renderedWidth > 0 && renderedWidth < 40;
                    const isEmoji = !spriteInfo && (
                                   img.classList.contains('emoji') ||
                                   img.src?.includes('emoji') ||
                                   img.src?.includes('emoticon') ||
                                   img.src?.includes('sticker') ||
                                   (img.naturalWidth < 50 && img.naturalHeight < 50));
                    const isIcon = !spriteInfo && (img.naturalWidth < 30 || img.naturalHeight < 30);

                    if (!isAvatar && !isEmoji && !isIcon && !isInlineEmojiSprite && img.src) {
                        const imgRef = buildMessageImageReference(img);
                        if (!imgRef) continue;
                        if (spriteInfo) {
                            imgRef.isAnimatedEmoji = true;
                            imgRef.frameWidth = spriteInfo.frameWidth;
                            imgRef.frameHeight = spriteInfo.frameHeight;
                            imgRef.frameCount = spriteInfo.frameCount;
                        }
                        imageReferences.push(imgRef);
                        // Queue up parallel image conversions
                        imagePromises.push(
                            imageToDataUrl(img, 400, 300, spriteInfo).then(dataUrl => {
                                if (dataUrl) {
                                    return { ...imgRef, dataUrl };
                                }
                                return checkpointingEnabled ? { ...imgRef, checkpointPlaceholder: true } : null;
                            })
                        );
                    }
                }
                // Wait for all image conversions in parallel
                const imageResults = await Promise.all(imagePromises);
                const images = imageResults.filter(img => img !== null);
                if (images.length > 0) {
                    console.log(`Successfully extracted ${images.length} images in parallel`);
                }

                // Extract file attachments from the message if enabled
                const attachments = [];
                if (fileAttachments && attachmentMetadata.length > 0) {
                    const downloadableAttachmentMetadata = attachmentMetadata.map((meta) => {
                        const isTooLarge = meta.fileSizeMB > 0 && meta.fileSizeMB > maxFileSize;
                        const alreadyExists = !!meta.filename && existingFilenames.includes(meta.filename);
                        const needsDownload = !isTooLarge && !!meta.sharePointUrl && !alreadyExists;

                        if (alreadyExists && meta.filename) {
                            console.log('Skipping already-downloaded file:', meta.filename);
                        }

                        if (needsDownload) {
                            totalFilesToDownload++;
                        }

                        return {
                            ...meta,
                            isTooLarge,
                            needsDownload,
                            alreadyExists
                        };
                    });

                    // Second pass: download files in parallel with concurrency limit (3 at a time)
                    // Separate files that need downloading
                    const filesToDownload = downloadableAttachmentMetadata.filter(m => m.needsDownload);

                    // Start timer if we have files to download
                    if (filesToDownload.length > 0) {
                        downloadStartTime = Date.now();
                        sendFileDownloadProgress();

                        // Download files in parallel (max 3 concurrent)
                        const downloadResults = await processWithConcurrency(
                            filesToDownload,
                            async (meta) => {
                                console.log('Downloading file:', meta.filename);
                                return await fileToDataUrl(meta.sharePointUrl, maxFileSize);
                            },
                            3, // concurrency limit
                            (completed, total) => {
                                filesDownloaded = completed;
                                sendFileDownloadProgress();
                            }
                        );

                        // Create a map of filename -> download result for quick lookup
                        const downloadMap = new Map();
                        filesToDownload.forEach((meta, idx) => {
                            downloadMap.set(meta.filename || meta.sharePointUrl || String(idx), downloadResults[idx]);
                        });

                        // Build attachments array with download results
                        for (const meta of downloadableAttachmentMetadata) {
                            let dataUrl = null;
                            let downloadError = null;
                            let mimeType = null;
                            let isTooLarge = meta.isTooLarge;

                            if (meta.needsDownload) {
                                const downloadResult = downloadMap.get(meta.filename || meta.sharePointUrl || '');
                                if (downloadResult && downloadResult.dataUrl) {
                                    dataUrl = downloadResult.dataUrl;
                                    mimeType = downloadResult.mimeType;
                                    console.log('Successfully downloaded file:', meta.filename, downloadResult.sizeMB + ' MB');
                                } else if (downloadResult && downloadResult.tooLarge) {
                                    isTooLarge = true;
                                    downloadError = `File too large (${downloadResult.sizeMB} MB)`;
                                } else {
                                    downloadError = (downloadResult && downloadResult.error) || 'Download failed';
                                    console.log('Could not download file:', meta.filename, downloadError);
                                }
                            } else if (meta.isTooLarge) {
                                downloadError = `File exceeds ${maxFileSize} MB limit`;
                            } else if (meta.alreadyExists) {
                                // File already exists in database - dataUrl will be preserved during merge
                                downloadError = null;
                            }

                            const attachment = {
                                filename: meta.filename,
                                sharePointUrl: meta.sharePointUrl,
                                fileDomain: meta.fileDomain,
                                arrivalTime: meta.arrivalTime,
                                extension: meta.extension,
                                fileSizeText: meta.fileSizeText,
                                fileSizeMB: meta.fileSizeMB,
                                tooLarge: isTooLarge,
                                dataUrl: dataUrl,
                                mimeType: mimeType,
                                downloadError: downloadError,
                                existsInDb: meta.alreadyExists || false
                            };
                            attachments.push(attachment);
                            const status = dataUrl ? '(downloaded)' : (meta.alreadyExists ? '(exists in db)' : '(not downloaded)');
                            console.log('Extracted file attachment:', meta.filename, meta.fileSizeText || 'size unknown', status);
                        }
                    } else {
                        // No files to download - just build attachments from metadata
                        for (const meta of downloadableAttachmentMetadata) {
                            const attachment = {
                                filename: meta.filename,
                                sharePointUrl: meta.sharePointUrl,
                                fileDomain: meta.fileDomain,
                                arrivalTime: meta.arrivalTime,
                                extension: meta.extension,
                                fileSizeText: meta.fileSizeText,
                                fileSizeMB: meta.fileSizeMB,
                                tooLarge: meta.isTooLarge,
                                dataUrl: null,
                                mimeType: null,
                                downloadError: meta.isTooLarge ? `File exceeds ${maxFileSize} MB limit` : null,
                                existsInDb: meta.alreadyExists || false
                            };
                            attachments.push(attachment);
                            const status = meta.alreadyExists ? '(exists in db)' : '(not downloaded)';
                            console.log('Extracted file attachment:', meta.filename, meta.fileSizeText || 'size unknown', status);
                        }
                    }
                }

                const attachmentSignals = attachmentMetadata.length > 0 || imageReferences.length > 0;
                const effectiveContent = content || (attachmentSignals
                    ? buildAttachmentFallbackContent(author, imageReferences, attachmentMetadata)
                    : '');

                // Only add message if it has content, images, or attachments
                if (effectiveContent || images.length > 0 || attachments.length > 0 || attachmentSignals) {
                    // O(1) duplicate check using Set
                    const messageKey = createMessageKey(
                        'message',
                        author,
                        timestamp,
                        effectiveContent,
                        buildMessageMediaSignature(images.length > 0 ? images : imageReferences, attachments.length > 0 ? attachments : attachmentMetadata)
                    );
                    if (!seenMessageKeys.has(messageKey)) {
                        seenMessageKeys.add(messageKey);
                        await queueExtractedMessage({
                            author,
                            timestamp,
                            content: effectiveContent,
                            images: images.length > 0 ? images : undefined,
                            attachments: attachments.length > 0 ? attachments : undefined
                        });
                    }
                }
            } catch (error) {
                if (error?.isCheckpointFailure) {
                    extractionErrors.checkpointSaveError = error.message || 'Checkpoint save failed';
                    throw error;
                }
                console.error('Error extracting message:', error);
                extractionErrors.messageErrors++;
            }
        }
    }

    // Scroll down gradually and collect all messages
    // In turbo mode, scroll faster with shorter waits
    // Scroll through chat and extract messages
    // Long chat mode and turbo mode both use fast scrolling
    let containerValidationCounter = 0;
    while (retryCount < maxRetries) {
        if (isOperationCancelled(operationId)) {
            break;
        }

        // Periodically validate chat container still exists (every 10 iterations)
        containerValidationCounter++;
        if (containerValidationCounter % 10 === 0) {
            // Bypass cache for validation - need fresh DOM check
            const currentContainer = getChatContainer(false);
            if (!currentContainer) {
                console.error('Chat container was removed during extraction');
                extractionErrors.containerLost = true;
                break; // Exit gracefully with whatever messages we have
            }

            if (currentContainer !== chatContainer) {
                const oldStillConnected = document.contains(chatContainer);
                const oldStillScrollable = (chatContainer.scrollHeight - chatContainer.clientHeight) > 20;
                if (!oldStillConnected || !oldStillScrollable) {
                    console.warn('Chat container changed during extraction, rebinding to new container');
                    chatContainer = currentContainer;
                    extractionErrors.containerDiagnostics = lastContainerSelectionDiagnostics;
                } else {
                    // Keep existing container when still valid to avoid needless context switches.
                    console.log('Detected alternate chat container but keeping current active container');
                }
            }
        }

        // Expand collapsed reply threads before extracting, otherwise channel
        // replies never enter the DOM and are silently dropped from the backup.
        await expandCollapsedRepliesInViewport(chatContainer, operationId);

        await extractVisibleMessages();

        if (isOperationCancelled(operationId)) {
            break;
        }

        const currentMessageCount = extractedMessageCount;

        // Check if at bottom BEFORE scrolling - extract messages first
        if (chatContainer.scrollTop >= chatContainer.scrollHeight - chatContainer.clientHeight - 5) {
            // Final extraction at the bottom to ensure we get the last messages
            await sleep(shortWait);
            await extractVisibleMessages();
            if (isOperationCancelled(operationId)) {
                break;
            }
            console.log("Reached bottom of chat, extracted final messages.");
            break;
        }

        // Scroll down - turbo and long chat mode use larger steps for speed
        const beforeProbe = captureViewportMessageProbe(chatContainer);
        const beforeScrollTop = beforeProbe.scrollTop;
        const beforeScrollHeight = beforeProbe.scrollHeight;
        const scrollRatio = longChatMode ? 0.5 : 0.85;
        const minScrollAmount = longChatMode ? 140 : 260;
        const scrollAmount = Math.max(Math.round(chatContainer.clientHeight * scrollRatio), minScrollAmount);
        const maxScrollTop = Math.max(0, chatContainer.scrollHeight - chatContainer.clientHeight);
        chatContainer.scrollTop = Math.min(maxScrollTop, beforeScrollTop + scrollAmount);
        const settleResult = await waitForViewportToSettle(chatContainer, {
            ...extractionSettleOptions,
            // Long chat mode keeps the slower two-round check for lazily fetched pages;
            // the no-progress retries below still catch a page that arrives late.
            settleOnQuietInterval: !longChatMode,
            operationId
        });

        if (settleResult?.cancelled || isOperationCancelled(operationId)) {
            break;
        }

        const afterProbe = settleResult?.probe || captureViewportMessageProbe(chatContainer);
        const afterScrollTop = afterProbe.scrollTop;
        const afterScrollHeight = afterProbe.scrollHeight;
        const messageProgress = currentMessageCount > lastMessageCount;
        const scrollProgress = afterScrollTop > beforeScrollTop + 1;
        const heightProgress = Math.abs(afterScrollHeight - beforeScrollHeight) > 1;
        const viewportProgress = didViewportProbeChange(beforeProbe, afterProbe);

        if (messageProgress || scrollProgress || heightProgress || viewportProgress) {
            retryCount = 0;
        } else {
            retryCount++;
            console.log(`extractChatText: no progress detected (${retryCount}/${maxRetries}) at scrollTop=${Math.round(afterScrollTop)}`);
            if (longChatMode) {
                await sleep(Math.min(800 + retryCount * 250, 4000));
            }
        }

        lastMessageCount = Math.max(lastMessageCount, currentMessageCount);
    }

    if (retryCount >= maxRetries && chatContainer.scrollTop < chatContainer.scrollHeight - chatContainer.clientHeight - 5) {
        console.warn('extractChatText stopped before confirmed bottom due to repeated no-progress passes', {
            retryCount,
            maxRetries,
            scrollTop: Math.round(chatContainer.scrollTop),
            scrollHeight: Math.round(chatContainer.scrollHeight),
            clientHeight: Math.round(chatContainer.clientHeight)
        });
        extractionErrors.bottomNotReached = true;
        extractionErrors.scrollDiagnostics = {
            ...(extractionErrors.scrollDiagnostics || {}),
            bottomReached: false,
            retryCount,
            maxRetries,
            scrollTop: Math.round(chatContainer.scrollTop),
            scrollHeight: Math.round(chatContainer.scrollHeight),
            clientHeight: Math.round(chatContainer.clientHeight)
        };
    }

    if (detectedMessageElements === 0) {
        extractionErrors.selectorMismatch = true;
    }

    if (checkpointingEnabled && checkpointPendingMessages.length > 0 && !isOperationCancelled(operationId)) {
        await flushCheckpointMessages('final');
    }

    console.log(`Extracted ${extractedMessageCount} messages total`);

    // Log any extraction errors
    if (extractionErrors.messageErrors > 0 || extractionErrors.imageErrors > 0 ||
        extractionErrors.fileErrors > 0 || extractionErrors.containerLost) {
        console.warn('Extraction completed with errors:', extractionErrors);
    }

    // Restore scroll position to where we started extracting from
    // If a specific startPosition was provided, restore to that; otherwise restore to original
    const restorePosition = startPosition !== null ? startPosition : originalScrollTop;
    chatContainer.scrollTop = restorePosition;
    console.log(`Restored scroll position to ${restorePosition}`);

    return {
        messages: checkpointingEnabled ? [] : chatText,
        messageCount: extractedMessageCount,
        participants: Array.from(participantSet),
        checkpoint: checkpointingEnabled ? {
            enabled: true,
            chatId: checkpointChatId,
            flushCount: checkpointFlushCount,
            newMessagesCount: checkpointNewMessagesCount,
            mergedExisting: checkpointStartedAsMerge
        } : null,
        errors: extractionErrors
    };
}

// Function to format chat text for file export
function formatChatText(chatData) {
    let formattedText = 'Teams Chat Backup\n';
    formattedText += '=================\n\n';

    for (const message of chatData) {
        formattedText += `Author: ${message.author}\n`;
        formattedText += `Time: ${message.timestamp}\n`;
        formattedText += `Message:\n${message.content}\n`;
        formattedText += '----------------------------------------\n\n';
    }

    return formattedText;
}

// Update the message listener to include text extraction
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (request.action === "ping") {
        sendResponse({
            status: "ready",
            url: location.href,
            title: document.title,
            chatName: getChatName()
        });
        return false;
    }

    if (request.action === "getBackupStatus") {
        sendResponse(getBackupStatusPayload());
        return false;
    }

    if (request.action === "backupChat") {
        // Prevent concurrent backups - reliability fix (P1-1)
        cleanupStaleOperations();
        if (isBackupInProgress()) {
            console.warn('Backup already in progress, rejecting new request');
            sendResponse({
                status: "failed",
                message: "A backup is already in progress. Please wait for it to complete.",
                retryable: true,
                operationId: currentOperationId
            });
            return false;
        }

        // Generate operation ID for this backup
        const operationId = generateOperationId();
        currentOperationId = operationId;
        currentOperationStartedAt = Date.now();
        currentOperationLastActivityAt = currentOperationStartedAt;

        const captureInterval = request.captureInterval || 1.0;
        const isPro = request.isPro || false;
        const generatePDF = request.generatePDF !== false; // default true
        const generateTXT = request.generateTXT !== false; // default true
        const includeCover = request.includeCover !== false; // default true
        const longChatMode = request.longChatMode || false; // default false
        const requestedIncrementalFrom = typeof request.incrementalFrom === 'string' ? request.incrementalFrom.trim() : '';
        const parsedIncrementalFrom = parseFlexibleTimestamp(requestedIncrementalFrom);
        const incrementalFromTimestamp = Number.isFinite(parsedIncrementalFrom)
            ? parsedIncrementalFrom
            : null;
        const effectiveStopTimestamp = incrementalFromTimestamp;
        const effectiveLongChatMode = longChatMode;
        const fileAttachments = request.fileAttachments || false; // default false
        const maxFileSize = request.maxFileSize || 50; // default 50 MB
        const existingFilenames = request.existingFilenames || []; // filenames already in DB
        const requestCheckpointing = request.checkpointing || null;

        isScrolling = true;
        isCancelled = false;
        invalidateContainerCache(); // Fresh lookup for new backup

        console.log(`backupChat [${operationId}] started: generatePDF=${generatePDF}, longChatMode=${effectiveLongChatMode}, incrementalFrom=${incrementalFromTimestamp ? new Date(incrementalFromTimestamp).toISOString() : 'off'}, fileAttachments=${fileAttachments}, maxFileSize=${maxFileSize}MB, existingFilenames=${existingFilenames.length}`);

        (async () => {
            try {
                // Verify operation is still active (wasn't cancelled/stopped)
                if (currentOperationId !== operationId || isCancelled) {
                    console.log(`backupChat [${operationId}] cancelled before start`);
                    revokeAllBlobUrls();
                    sendResponse({ status: "cancelled", operationId });
                    return;
                }

                let textUrl = null;
                let pdfUrl = null;
                let chatText = [];
                let formattedText = '';

                // Get chat name early
                const chatName = getChatName();
                const effectiveCheckpointConfig = {
                    enabled: !!requestCheckpointing?.enabled && !generatePDF,
                    isPro,
                    chatIdentity: requestCheckpointing?.chatIdentity || {
                        name: chatName,
                        rawName: chatName,
                        isChannel: false,
                        teamName: null,
                        sourceKind: 'chat'
                    }
                };

                if (isUnsupportedThreadsOrPostsLayout()) {
                    chrome.runtime.sendMessage({
                        action: "updatePhase",
                        phase: "Posts layout detected. Preparing thread backup...",
                        progress: 3,
                        operationId
                    });
                    const feedReady = await waitForExtractableThreadsFeed(operationId, 8000);
                    if (currentOperationId !== operationId || isCancelled) {
                        revokeAllBlobUrls();
                        sendResponse({ status: "cancelled", operationId });
                        return;
                    }
                    if (!feedReady) {
                        sendResponse({
                            status: "failed",
                            message: "This channel uses the Posts/Threads layout and no readable messages were found. Open a conversation in this channel so messages are visible, then try again.",
                            unsupportedLayout: "threads_posts",
                            operationId
                        });
                        return;
                    }
                    console.log('Threads/posts layout detected but feed is readable; continuing with best-effort backup.');
                }

                // Use turbo mode (faster scrolling) when not generating PDF
                // Note: longChatMode takes precedence over turboMode for wait times
                const turboMode = !generatePDF;
                let scrollResult = { success: true, stoppedByUser: false };
                let partialBackup = false;
                let partialReason = null;

                // Phase 1: Scroll to top - ALWAYS scroll to ensure full history capture
                // This prevents partial extractions (e.g., only capturing 45 messages when chat has thousands)
                // Prioritize reliability over speed - even turbo mode needs full scroll
                chrome.runtime.sendMessage({
                    action: "updatePhase",
                    phase: incrementalFromTimestamp ? "Locating last saved message..." : "Scrolling to top...",
                    progress: 5,
                    operationId
                });

                scrollResult = await scrollToTop(effectiveLongChatMode, effectiveStopTimestamp, operationId);

                // Check if operation was cancelled during scroll
                if (currentOperationId !== operationId || isCancelled) {
                    console.log(`backupChat [${operationId}] cancelled during scroll`);
                    revokeAllBlobUrls();
                    sendResponse({ status: "cancelled", operationId });
                    return;
                }

                if (!scrollResult.success && !scrollResult.stoppedByUser) {
                    partialBackup = true;
                    partialReason = scrollResult.reason || 'top_not_reached';
                    console.warn(`backupChat [${operationId}] could not reach top, continuing with visible-message backup`, scrollResult);
                    chrome.runtime.sendMessage({
                        action: "updatePhase",
                        phase: "Could not reach top. Backing up visible messages...",
                        progress: 15,
                        operationId
                    });
                }

                if (scrollResult.stoppedByUser) {
                    partialBackup = true;
                    partialReason = 'stopped_by_user';
                    console.warn(`backupChat [${operationId}] user stopped top scroll early; saving visible-message preview`, scrollResult);
                    chrome.runtime.sendMessage({
                        action: "updatePhase",
                        phase: "Stopped early. Backing up visible messages...",
                        progress: 15,
                        operationId
                    });
                }

                // Phase 2: Extract text content (always extract for storage)
                // If user stopped scrolling, start from that position instead of scrolling to top again
                chrome.runtime.sendMessage({
                    action: "updatePhase",
                    phase: "Extracting messages...",
                    progress: 20,
                    operationId
                });

                const currentContainerForExtraction = getChatContainer();
                const startPosition = scrollResult.stoppedByUser
                    ? (typeof scrollResult.currentPosition === 'number' ? scrollResult.currentPosition : null)
                    : (partialBackup
                        ? (typeof currentContainerForExtraction?.scrollTop === 'number' ? currentContainerForExtraction.scrollTop : null)
                        : null);
                let extractResult = await extractChatText(
                    startPosition,
                    turboMode,
                    effectiveLongChatMode,
                    fileAttachments,
                    maxFileSize,
                    existingFilenames,
                    incrementalFromTimestamp,
                    effectiveCheckpointConfig,
                    operationId
                );

                if (Number(extractResult.messageCount || 0) === 0 && Number.isFinite(incrementalFromTimestamp)) {
                    const fallbackErrors = extractResult.errors || {};
                    const visibleMessageCandidates = currentContainerForExtraction
                        ? currentContainerForExtraction.querySelectorAll('[data-tid="chat-pane-item"], [data-tid*="message-pane-item"], [data-tid*="post-item"], [data-tid*="reply-item"], [data-tid*="thread-item"], div[class*="fui-ChatMyMessage"], div[class*="fui-ChatMessage"], [role="article"], [role="listitem"]').length
                        : 0;
                    const viewportProbeCandidates = currentContainerForExtraction
                        ? getViewportProbeMessageElements(currentContainerForExtraction).length
                        : 0;
                    const visibleTimeElements = currentContainerForExtraction
                        ? currentContainerForExtraction.querySelectorAll('time').length
                        : 0;
                    const shouldRetryFullExtraction = !fallbackErrors.selectorMismatch &&
                        !fallbackErrors.containerLost &&
                        (visibleMessageCandidates > 0 || viewportProbeCandidates > 0 || visibleTimeElements > 0);

                    if (shouldRetryFullExtraction) {
                        console.warn(`backupChat [${operationId}] incremental extraction returned 0 messages despite visible message signals; retrying full extraction`, {
                            visibleMessageCandidates,
                            viewportProbeCandidates,
                            visibleTimeElements
                        });
                        chrome.runtime.sendMessage({
                            action: "updatePhase",
                            phase: "Retrying full extraction...",
                            progress: 22,
                            operationId
                        });
                        extractResult = await extractChatText(
                            startPosition,
                            turboMode,
                            effectiveLongChatMode,
                            fileAttachments,
                            maxFileSize,
                            existingFilenames,
                            null,
                            effectiveCheckpointConfig,
                            operationId
                        );
                        extractResult.errors = {
                            ...(extractResult.errors || {}),
                            incrementalFallbackUsed: true,
                            incrementalFallbackVisibleCandidates: visibleMessageCandidates,
                            incrementalFallbackProbeCandidates: viewportProbeCandidates,
                            incrementalFallbackTimeElements: visibleTimeElements
                        };
                    }
                }

                // Check if operation was cancelled during extraction
                if (currentOperationId !== operationId || isCancelled) {
                    console.log(`backupChat [${operationId}] cancelled during extraction`);
                    revokeAllBlobUrls();
                    sendResponse({ status: "cancelled", operationId });
                    return;
                }

                chatText = extractResult.messages;
                const extractedMessageCount = Number(extractResult.messageCount || (chatText || []).length);
                const checkpointSummary = extractResult.checkpoint || null;
                const extractionErrors = extractResult.errors;
                if (extractionErrors.bottomNotReached) {
                    partialBackup = true;
                    if (!partialReason) {
                        partialReason = 'bottom_not_reached';
                    }
                }
                // Extraction stopped early when the chat pane disappeared, so the
                // backup must stay resumable instead of being marked complete.
                if (extractionErrors.containerLost) {
                    partialBackup = true;
                    if (!partialReason) {
                        partialReason = 'container_lost';
                    }
                }
                if (partialBackup) {
                    if (partialReason === 'top_not_reached' || partialReason === 'stopped_by_user') {
                        extractionErrors.topNotReached = true;
                    }
                    extractionErrors.scrollDiagnostics = {
                        ...(extractionErrors.scrollDiagnostics || {}),
                        reason: partialReason,
                        strategy: scrollResult.strategy || null,
                        finalScrollTop: scrollResult.finalScrollTop,
                        scrollHeight: scrollResult.scrollHeight,
                        clientHeight: scrollResult.clientHeight,
                        currentPosition: scrollResult.currentPosition,
                        attemptedGlobalRetries: scrollResult.attemptedGlobalRetries,
                        maxGlobalRetries: scrollResult.maxGlobalRetries,
                        strategies: scrollResult.strategies || [],
                        domDiagnostics: scrollResult.domDiagnostics || null
                    };
                }
                if (!checkpointSummary?.enabled) {
                    formattedText = formatChatText(chatText);
                }

                if (generateTXT && !checkpointSummary?.enabled) {
                    // Add UTF-8 BOM for proper encoding of Korean/international characters
                    const BOM = '\uFEFF';
                    const textBlob = new Blob([BOM + formattedText], { type: 'text/plain;charset=utf-8' });
                    textUrl = createTrackedBlobUrl(textBlob);
                }

                // Phase 3: Capture screenshots and generate PDF if needed
                if (generatePDF) {
                    chrome.runtime.sendMessage({
                        action: "updatePhase",
                        phase: "Capturing screenshots...",
                        progress: 40,
                        operationId
                    });

                    const captureResult = await captureAllScreenshots(captureInterval, operationId);
                    const screenshots = captureResult?.screenshots || [];
                    const screenshotDiagnostics = captureResult?.diagnostics || null;
                    extractionErrors.screenshotDiagnostics = screenshotDiagnostics;

                    if (currentOperationId !== operationId || isCancelled) {
                        console.log(`backupChat [${operationId}] cancelled during screenshot capture`);
                        revokeAllBlobUrls(); // Clean up on cancel
                        sendResponse({ status: "cancelled", operationId });
                        return;
                    }

                    // Phase 4: Generate PDF
                    chrome.runtime.sendMessage({
                        action: "updatePhase",
                        phase: "Generating PDF...",
                        progress: 80,
                        operationId
                    });

                    if (screenshots.length === 0) {
                        extractionErrors.pdfError = 'no_screenshots_captured';
                        console.warn(`backupChat [${operationId}] PDF skipped: no screenshots captured`, screenshotDiagnostics);
                    } else {
                        try {
                            pdfUrl = await generateContinuousPDF(screenshots, !isPro, includeCover, operationId);
                            if (!pdfUrl) {
                                extractionErrors.pdfError = 'pdf_generation_returned_null';
                                console.warn(`backupChat [${operationId}] PDF generation returned null`);
                            }
                        } catch (pdfError) {
                            extractionErrors.pdfError = `pdf_generation_exception:${pdfError?.message || 'unknown'}`;
                            console.error(`backupChat [${operationId}] PDF generation threw:`, pdfError);
                        }
                    }
                }

                // Extract participants from messages
                chrome.runtime.sendMessage({
                    action: "updatePhase",
                    phase: "Finalizing...",
                    progress: 95,
                    operationId
                });

                const participants = Array.isArray(extractResult.participants) && extractResult.participants.length > 0
                    ? extractResult.participants
                    : extractParticipants(chatText);
                let finalizedCheckpointChatId = checkpointSummary?.chatId || null;

                if (checkpointSummary?.enabled && extractedMessageCount > 0) {
                    try {
                        const finalizeResponse = await sendRuntimeMessageWithResponse({
                            action: 'finalizeCheckpointBackup',
                            operationId,
                            chatIdentity: effectiveCheckpointConfig.chatIdentity,
                            checkpointChatId: checkpointSummary?.chatId || null,
                            participants,
                            partialBackup: !!partialBackup,
                            textContent: '',
                            isPro
                        });

                        if (!finalizeResponse || finalizeResponse.status !== 'success') {
                            extractionErrors.checkpointFinalizeError = finalizeResponse?.message || 'Checkpoint finalization failed';
                        } else {
                            finalizedCheckpointChatId = finalizeResponse.id || finalizedCheckpointChatId;
                        }
                    } catch (error) {
                        extractionErrors.checkpointFinalizeError = error?.message || 'Checkpoint finalization failed';
                    }
                }

                // Validate extraction produced meaningful results
                if (extractedMessageCount === 0) {
                    // No messages extracted - capture DOM diagnostics for debugging
                    if (!extractionErrors.domDiagnostics) {
                        extractionErrors.domDiagnostics = captureDomDiagnostics();
                    }

                    const containerExists = getChatContainer(false); // bypass cache
                    if (!containerExists) {
                        sendResponse({
                            status: "failed",
                            message: "Chat container disappeared during extraction. Please try again.",
                            errors: extractionErrors,
                            operationId
                        });
                        return;
                    }

                    if (partialBackup) {
                        sendResponse({
                            status: "failed",
                            message: "Could not scroll to older messages and no visible messages were captured. Please open a message-rich area and try again.",
                            errors: extractionErrors,
                            operationId,
                            retryable: true
                        });
                        return;
                    }

                    if (extractionErrors.selectorMismatch) {
                        sendResponse({
                            status: "failed",
                            message: "Could not read messages from this Teams layout. Please refresh Teams and try again.",
                            errors: extractionErrors,
                            operationId,
                            retryable: true
                        });
                        return;
                    }

                    if (isUnsupportedThreadsOrPostsLayout()) {
                        sendResponse({
                            status: "failed",
                            message: "This Teams channel layout is not supported in the current view yet. Please open its chat-style Conversation view and try again.",
                            errors: extractionErrors,
                            operationId,
                            unsupportedLayout: "threads_posts"
                        });
                        return;
                    }

                    // Check if this is an empty channel
                    const isEmptyChannel = !!document.querySelector('[data-tid="channel-empty-state"]');

                    // Container exists but no messages - might be empty chat/channel or selector mismatch
                    sendResponse({
                        status: "failed",
                        message: isEmptyChannel
                            ? "This channel has no messages yet. Skipping empty channel."
                            : "No messages were found in this chat. If the chat has messages, please report this issue.",
                        errors: extractionErrors,
                        operationId,
                        emptyChannel: isEmptyChannel
                    });
                    return;
                }

                // Check if we got suspiciously few messages (might be incomplete extraction)
                if (extractedMessageCount < 10 && !turboMode) {
                    console.warn(`Only ${extractedMessageCount} messages extracted - this might be incomplete`);
                }

                chrome.runtime.sendMessage({ action: "backupComplete", operationId });
                sendResponse({
                    status: "completed",
                    operationId,
                    pdfUrl: pdfUrl,
                    textUrl: textUrl,
                    chatName: chatName,
                    messages: checkpointSummary?.enabled ? [] : chatText,
                    messageCount: extractedMessageCount,
                    participants: participants,
                    textContent: formattedText,
                    incrementalFrom: incrementalFromTimestamp ? new Date(incrementalFromTimestamp).toISOString() : null,
                    extractionErrors: extractionErrors,
                    checkpointed: checkpointSummary?.enabled === true,
                    checkpointChatId: finalizedCheckpointChatId,
                    checkpointMerged: checkpointSummary?.mergedExisting === true,
                    checkpointNewMessagesCount: checkpointSummary?.newMessagesCount || 0,
                    checkpointFlushCount: checkpointSummary?.flushCount || 0,
                    partialBackup: partialBackup,
                    partialReason: partialReason
                });
            } catch (error) {
                console.error(`backupChat [${operationId}] failed with error:`, error);
                revokeAllBlobUrls(); // Clean up on error
                sendResponse({
                    status: "failed",
                    message: error.message || "An unexpected error occurred during backup.",
                    error: error?.code || undefined,
                    count: error?.count,
                    operationId
                });
            } finally {
                // Always clear operation ID when done
                if (currentOperationId === operationId) {
                    clearCurrentOperationState();
                    isScrolling = false;
                }
            }
        })();
        return true;
    }

    if (request.action === "stopScrolling") {
        isScrolling = false;
        sendResponse({ status: "stopped" });
    }

    // Handle cancellation
    if (request.action === "cancelBackup") {
        // Optional: Only cancel if operation ID matches (for future enhancement)
        isCancelled = true;
        isScrolling = false;
        revokeAllBlobUrls(); // Clean up Blob URLs on cancel
        console.log(`backupChat cancelled, current operation: ${currentOperationId}`);
        sendResponse({ status: "cancelled", operationId: currentOperationId });
    }

    // Clean up Blob URLs (called by popup after processing results)
    if (request.action === "cleanupBlobUrls") {
        revokeAllBlobUrls();
        sendResponse({ status: "cleaned" });
    }

    // Get available chats
    if (request.action === "getAvailableChats") {
        // Scanning the sidebar can switch Teams views and filters, which would
        // pull the running backup away from its chat. Popup reopens mid-backup.
        if (isBackupInProgress()) {
            sendResponse({ status: "success", chats: [], diagnostics: { backupInProgress: true } });
            return false;
        }
        (async () => {
            try {
                const cached = request.forceRefresh ? null : getCachedAvailableChats();
                // The popup's preload and its Backup button can ask at the same time.
                // Overlapping scans save and "restore" each other's filter state, which
                // leaves Teams' Chats filter switched on, so share one scan.
                if (!cached && !availableChatsScan) {
                    availableChatsScan = getAvailableChatsWithRecovery()
                        .then((scanResult) => {
                            if ((scanResult.chats || []).length > 0) {
                                availableChatsCache = { chats: scanResult.chats, diagnostics: scanResult.diagnostics, at: Date.now() };
                            }
                            return scanResult;
                        })
                        .finally(() => {
                            availableChatsScan = null;
                        });
                }
                const result = cached || await availableChatsScan;
                const chats = (result.chats || []).map(chat => ({
                    ...chat,
                    kind: chat.kind || getChatKindForUi(chat)
                }));
                sendResponse({ status: "success", chats, diagnostics: result.diagnostics });
            } catch (error) {
                console.error('getAvailableChats failed:', error);
                sendResponse({
                    status: "failed",
                    error: error?.message || 'getAvailableChats_failed',
                    chats: [],
                    diagnostics: {
                        ...(getChatSidebarDiagnostics() || {}),
                        requestFailed: true,
                        requestError: error?.message || 'unknown',
                        domDiagnostics: captureDomDiagnostics()
                    }
                });
            }
        })();
        return true;
    }

    // Open a specific chat
    if (request.action === "openChat") {
        (async () => {
            const success = await openChat(request.chatIndex, request.chatTarget);
            sendResponse({ status: success ? "success" : "failed" });
        })();
        return true; // async response
    }
});
