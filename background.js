// background.js
importScripts('db.js');

// Public RevenueCat Web Billing key. Store builds inject the real value; source builds run as the free tier.
const REVENUECAT_API_KEY = 'rcb_YOUR_PUBLIC_WEB_BILLING_KEY';
const REVENUECAT_ENTITLEMENT_ID = 'Teamsbackuppro';
const REVENUECAT_PRODUCT_ID = 'Teamsbackuppro';
const REVENUECAT_API_BASE = 'https://api.revenuecat.com/v1';
// Store installs get an update_url in their manifest; unpacked test builds do not.
// Only test builds accept RevenueCat sandbox purchases.
const IS_DEVELOPMENT_BUILD = !('update_url' in chrome.runtime.getManifest());
// After checkout opens, check for the payment every 30 s for 30 minutes, even
// though the popup closed when the checkout tab took focus.
const PURCHASE_WATCH_ALARM = 'purchaseWatch';
const PURCHASE_WATCH_WINDOW_MS = 30 * 60 * 1000;
const PENDING_CHECKOUT_KEY = 'pendingCheckout';
// Retained only to remove the alarm created by v6.6.
const ENTITLEMENT_REVERIFY_ALARM = 'entitlementReverify';
const V6_PERMISSION_MIGRATION_KEY = 'v6LeastPrivilegeMigrationComplete';
const LEGACY_HOST_ORIGINS = [
    'https://teams.microsoft.com/*',
    'https://teams.live.com/*',
    'https://teams.cloud.microsoft/*',
    'https://*.mcas.ms/*',
    'https://*.office.net/*',
    'https://*.office.com/*',
    'https://api.revenuecat.com/*',
    'https://pay.rev.cat/*',
    'https://bullfighterstudio.com/*',
    'https://www.bullfighterstudio.com/*'
];

// ==================== Fetch Timeout Handling (P1-4) ====================
const FETCH_TIMEOUT_MS = 15000; // 15 seconds per fetch request
const FETCH_RETRY_ATTEMPTS = 2;
const FETCH_RETRY_DELAY_MS = 1000;

function getAllPermissions() {
    return new Promise((resolve) => {
        chrome.permissions.getAll((permissions) => {
            resolve(permissions || { permissions: [], origins: [] });
        });
    });
}

function removePermissions(permissions) {
    return new Promise((resolve) => {
        chrome.permissions.remove(permissions, (removed) => {
            if (chrome.runtime.lastError) {
                console.warn('Background: Could not remove legacy permission:', chrome.runtime.lastError.message);
                resolve(false);
                return;
            }
            resolve(removed === true);
        });
    });
}

async function migrateLegacyPermissions() {
    const migrationState = await new Promise((resolve) => {
        chrome.storage.local.get([V6_PERMISSION_MIGRATION_KEY], resolve);
    });
    if (migrationState[V6_PERMISSION_MIGRATION_KEY]) {
        return;
    }

    const granted = await getAllPermissions();
    const grantedOrigins = new Set(granted.origins || []);
    const originsToRemove = LEGACY_HOST_ORIGINS.filter((origin) => grantedOrigins.has(origin));
    const namedPermissionsToRemove = (granted.permissions || []).includes('downloads')
        ? ['downloads']
        : [];

    const results = [];
    if (originsToRemove.length > 0) {
        results.push(await removePermissions({ origins: originsToRemove }));
    }
    if (namedPermissionsToRemove.length > 0) {
        results.push(await removePermissions({ permissions: namedPermissionsToRemove }));
    }

    if (results.every(Boolean)) {
        chrome.storage.local.set({ [V6_PERMISSION_MIGRATION_KEY]: true });
    }
}

migrateLegacyPermissions().catch((error) => {
    console.warn('Background: Legacy permission migration failed:', error);
});

// Fetch with timeout and retry logic
async function fetchWithTimeout(url, options = {}, timeoutMs = FETCH_TIMEOUT_MS, retryCount = FETCH_RETRY_ATTEMPTS) {
    let lastError = null;

    for (let attempt = 0; attempt <= retryCount; attempt++) {
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), timeoutMs);

        try {
            console.log(`Fetch attempt ${attempt + 1}/${retryCount + 1} for ${url}`);

            const response = await fetch(url, {
                ...options,
                signal: controller.signal
            });

            clearTimeout(timeoutId);

            // Check for timeout specifically
            if (response.status === 408 || response.headers.get('x-status') === 'timeout') {
                throw new Error(`Request timeout after ${timeoutMs}ms`);
            }

            return response;
        } catch (error) {
            clearTimeout(timeoutId);
            lastError = error;

            // If aborted, it was a timeout
            if (error.name === 'AbortError') {
                console.warn(`Fetch timeout on attempt ${attempt + 1}/${retryCount + 1}`);
                lastError = new Error(`Request timeout after ${timeoutMs}ms`);
                lastError.name = 'TimeoutError';
            } else {
                console.warn(`Fetch failed on attempt ${attempt + 1}/${retryCount + 1}:`, error.message);
            }

            // Don't retry on certain error codes
            if (error.status === 401 || error.status === 403 || error.status === 404) {
                console.log('Not retrying on authentication/not-found error');
                throw error;
            }

            // Wait before retry (exponential backoff)
            if (attempt < retryCount) {
                const delay = FETCH_RETRY_DELAY_MS * Math.pow(2, attempt);
                console.log(`Waiting ${delay}ms before retry...`);
                await new Promise(resolve => setTimeout(resolve, delay));
            }
        }
    }

    // All retries exhausted
    console.error(`Fetch failed after ${retryCount + 1} attempts:`, lastError);
    throw lastError;
}

function notifyPurchaseVerified() {
    chrome.runtime.sendMessage({ action: 'purchaseVerified' }, () => {
        if (chrome.runtime.lastError) {
            // No active listeners - safe to ignore.
        }
    });
}

async function fetchSubscriberData(userId, environment) {
    const envQuery = environment ? `?environment=${environment}` : '';
    const encodedUserId = encodeURIComponent(userId);
    const url = `${REVENUECAT_API_BASE}/subscribers/${encodedUserId}${envQuery}`;

    try {
        console.log(`Background: Fetching RevenueCat subscriber ${userId} (${environment || 'production'})`);

        const response = await fetchWithTimeout(url, {
            headers: {
                Authorization: `Bearer ${REVENUECAT_API_KEY}`,
                'Content-Type': 'application/json',
                'X-Platform': 'web',
                'X-Client-Version': '1.0.0'
            }
        }, FETCH_TIMEOUT_MS, FETCH_RETRY_ATTEMPTS);

        if (response.status === 404) {
            console.warn(`Background: RevenueCat returned 404 for ${userId} (${environment || 'production'})`);
            return null;
        }

        if (!response.ok) {
            const errorText = await response.text().catch(() => '');
            const error = new Error(`RevenueCat responded with HTTP ${response.status}`);
            error.status = response.status;
            error.body = errorText;
            error.environment = environment || 'production';
            throw error;
        }

        const data = await response.json();
        const entitlements = data?.subscriber?.entitlements || {};
        const nonSubscriptions = data?.subscriber?.non_subscriptions || {};
        console.log(`Background: RevenueCat returned data for ${userId} (${environment || 'production'})`, data);
        console.log('Background: Entitlements snapshot', JSON.stringify(entitlements));
        console.log('Background: Non-subscriptions snapshot', JSON.stringify(nonSubscriptions));
        return data;
    } catch (error) {
        // Distinguish between timeout and other errors
        if (error.name === 'TimeoutError') {
            const timeoutError = new Error(`RevenueCat API timeout after ${FETCH_TIMEOUT_MS}ms`);
            timeoutError.cause = error;
            timeoutError.isTimeout = true;
            throw timeoutError;
        }
        if (error instanceof TypeError) {
            const networkError = new Error('Network error while contacting RevenueCat');
            networkError.cause = error;
            throw networkError;
        }
        console.error('Background: Fetch subscriber failed', error);
        throw error;
    }
}

function extractEntitlement(data) {
    const entitlements = data?.subscriber?.entitlements;
    if (!entitlements) {
        return null;
    }
    return entitlements[REVENUECAT_ENTITLEMENT_ID] || null;
}

// RevenueCat can return test-mode purchases next to real ones; store builds ignore them.
function isSandboxOnlyEntitlement(data, entitlement) {
    const productId = entitlement?.product_identifier;
    if (!productId) return false;

    const purchases = data?.subscriber?.non_subscriptions?.[productId];
    if (Array.isArray(purchases) && purchases.length > 0) {
        return purchases.every(purchase => purchase?.is_sandbox === true);
    }

    const subscription = data?.subscriber?.subscriptions?.[productId];
    return subscription?.is_sandbox === true;
}

function extractLatestNonSubscription(data) {
    const nonSubs = data?.subscriber?.non_subscriptions;
    if (!nonSubs) {
        return null;
    }

    const entries = (nonSubs[REVENUECAT_PRODUCT_ID] || [])
        .filter(entry => IS_DEVELOPMENT_BUILD || entry?.is_sandbox !== true);
    if (entries.length === 0) {
        return null;
    }

    return entries.reduce((latest, entry) => {
        if (!latest) {
            return entry;
        }

        const latestTime = Date.parse(latest.purchase_date || latest.original_purchase_date || 0);
        const entryTime = Date.parse(entry.purchase_date || entry.original_purchase_date || 0);

        if (!Number.isFinite(latestTime)) {
            return entry;
        }

        if (!Number.isFinite(entryTime)) {
            return latest;
        }

        return entryTime >= latestTime ? entry : latest;
    }, null);
}

function isEntitlementActive(entitlement) {
    if (!entitlement) {
        return false;
    }

    if (!entitlement.expires_date) {
        return true; // lifetime purchase
    }

    const expiresAt = Date.parse(entitlement.expires_date);
    if (!Number.isFinite(expiresAt)) {
        return false;
    }

    if (expiresAt > Date.now()) {
        return true;
    }

    if (entitlement.grace_period_expires_date) {
        const graceExpires = Date.parse(entitlement.grace_period_expires_date);
        if (Number.isFinite(graceExpires) && graceExpires > Date.now()) {
            return true;
        }
    }

    return false;
}

async function resolveRevenueCatEntitlementStatus(userId) {
    if (!userId) {
        throw new Error('Missing user ID');
    }

    // Production first; sandbox purchases count only in unpacked test builds.
    const environments = IS_DEVELOPMENT_BUILD ? [null, 'sandbox'] : [null];
    let lastError = null;
    // True once production RevenueCat answered with the customer's record, so a
    // "not Pro" result is RevenueCat's answer rather than a missing response.
    let productionAnswered = false;

    for (const environment of environments) {
        try {
            const subscriberData = await fetchSubscriberData(userId, environment);
            if (!subscriberData) {
                continue;
            }
            if (!environment) {
                productionAnswered = true;
            }

            const entitlement = extractEntitlement(subscriberData);
            const sandboxOnly = !IS_DEVELOPMENT_BUILD && isSandboxOnlyEntitlement(subscriberData, entitlement);
            if (isEntitlementActive(entitlement) && !sandboxOnly) {
                return {
                    isPro: true,
                    environment: environment || 'production',
                    entitlement,
                    source: 'entitlement',
                    provider: 'revenuecat'
                };
            }

            const latestNonSubscription = extractLatestNonSubscription(subscriberData);
            if (latestNonSubscription) {
                const hasPurchase = Boolean(latestNonSubscription.purchase_date || latestNonSubscription.original_purchase_date);
                if (hasPurchase) {
                    return {
                        isPro: true,
                        environment: latestNonSubscription.is_sandbox ? 'sandbox' : (environment || 'production'),
                        entitlement: latestNonSubscription,
                        source: 'non_subscription',
                        provider: 'revenuecat'
                    };
                }
            }
        } catch (error) {
            if (error.status === 404) {
                continue;
            }
            lastError = error;
        }
    }

    if (lastError) {
        throw lastError;
    }

    return {
        isPro: false,
        definitive: productionAnswered,
        environment: null,
        entitlement: null,
        source: null,
        provider: 'revenuecat'
    };
}

async function resolveEntitlementStatus(userId) {
    return resolveRevenueCatEntitlementStatus(userId);
}

async function getCachedEntitlement() {
    const stored = await chrome.storage.local.get(['isPro', 'purchaseVerified', 'paymentProvider', 'entitlementUserId', 'userId']);
    return {
        isPro: stored.isPro === true && stored.purchaseVerified === true,
        provider: stored.paymentProvider || null,
        userId: stored.entitlementUserId || stored.userId,
        cached: true
    };
}

const entitlementChecks = new Map();

// Called only for a pending checkout or an explicit Restore. Successful
// activation is permanent on this browser profile; it never needs a recheck.
async function determineAndPersistEntitlement(userId) {
    const cached = await getCachedEntitlement();
    if (cached.isPro) return cached;
    if (entitlementChecks.has(userId)) return entitlementChecks.get(userId);

    const check = (async () => {
        const result = await resolveEntitlementStatus(userId);
        // Another Restore may have completed while this request was in flight.
        const activated = await getCachedEntitlement();
        if (activated.isPro) return activated;
        if (result.isPro) {
            await chrome.storage.local.set({
                isPro: true,
                purchaseVerified: true,
                purchaseVerifiedTime: Date.now(),
                paymentProvider: result.provider || null,
                entitlementUserId: userId,
                userId
            });
            await stopPurchaseWatch();
        }
        return { ...result, userId };
    })();
    entitlementChecks.set(userId, check);
    try {
        return await check;
    } finally {
        entitlementChecks.delete(userId);
    }
}

// ==================== Purchase watch ====================

async function startPurchaseWatch(userId) {
    await backgroundReady;
    const cached = await getCachedEntitlement();
    if (cached.isPro) return { status: 'already_pro', isPro: true };
    if (!userId) throw new Error('Missing License ID');
    await chrome.storage.local.set({
        [PENDING_CHECKOUT_KEY]: { userId, startedAt: Date.now() }
    });
    await chrome.alarms.create(PURCHASE_WATCH_ALARM, { delayInMinutes: 0.5, periodInMinutes: 0.5 });
    return { status: 'watching' };
}

async function stopPurchaseWatch() {
    await chrome.alarms.clear(PURCHASE_WATCH_ALARM);
    await chrome.storage.local.remove([PENDING_CHECKOUT_KEY, 'checkEntitlementsOnOpen', 'checkEntitlementsTimestamp']);
}

async function showProActivatedPage() {
    const url = `${chrome.runtime.getURL('viewer.html')}?pro=activated`;
    const { viewerTabId } = await chrome.storage.local.get(['viewerTabId']);

    if (viewerTabId) {
        try {
            const tab = await chrome.tabs.update(viewerTabId, { url, active: true });
            if (tab?.windowId) {
                await chrome.windows.update(tab.windowId, { focused: true });
            }
            return;
        } catch (error) {
            // The stored viewer tab was closed; open a new one.
        }
    }

    await chrome.tabs.create({ url });
}

async function handlePurchaseVerified() {
    await stopPurchaseWatch();
    await chrome.storage.local.remove(['checkEntitlementsOnOpen', 'checkEntitlementsTimestamp', 'purchaseJustCompleted', 'purchaseCompletedTime']);
    notifyPurchaseVerified();
    await showProActivatedPage().catch((error) => {
        console.warn('Background: Could not open the Pro confirmation page:', error);
    });
}

function isPendingCheckoutActive(pending) {
    const age = Date.now() - pending?.startedAt;
    return !!pending?.userId && Number.isFinite(age) && age >= 0 && age < PURCHASE_WATCH_WINDOW_MS;
}

async function checkPendingCheckout() {
    await backgroundReady;
    const stored = await chrome.storage.local.get([PENDING_CHECKOUT_KEY, 'isPro', 'purchaseVerified']);
    const pending = stored[PENDING_CHECKOUT_KEY];

    if (!isPendingCheckoutActive(pending) || (stored.isPro && stored.purchaseVerified)) {
        await stopPurchaseWatch();
        return { isPro: stored.isPro === true && stored.purchaseVerified === true };
    }

    try {
        const result = await determineAndPersistEntitlement(pending.userId);
        if (result.isPro) {
            await handlePurchaseVerified();
            return { isPro: true };
        }
    } catch (error) {
        console.warn('Background: Purchase watch check failed:', error);
    }

    if (!isPendingCheckoutActive(pending)) {
        await stopPurchaseWatch();
    }
    return { isPro: false };
}

async function ensureBackgroundAlarms() {
    // Remove v6.6's recurring checks, including when this worker starts after an update.
    await chrome.alarms.clear(ENTITLEMENT_REVERIFY_ALARM);
    await chrome.storage.local.remove(['lastEntitlementCheckAt', 'entitlementMissingSince', 'checkEntitlementsOnOpen', 'checkEntitlementsTimestamp']);
    const stored = await chrome.storage.local.get([PENDING_CHECKOUT_KEY, 'isPro', 'purchaseVerified']);
    if ((stored.isPro && stored.purchaseVerified) || !isPendingCheckoutActive(stored[PENDING_CHECKOUT_KEY])) {
        await stopPurchaseWatch();
    } else if (!await chrome.alarms.get(PURCHASE_WATCH_ALARM)) {
        // Alarms may not survive a browser restart. Keep the original expiry.
        await chrome.alarms.create(PURCHASE_WATCH_ALARM, { delayInMinutes: 0.5, periodInMinutes: 0.5 });
    }
}

chrome.alarms.onAlarm.addListener((alarm) => {
    if (alarm.name === PURCHASE_WATCH_ALARM) {
        checkPendingCheckout().catch((error) => console.warn('Background: Purchase watch failed:', error));
    } else if (alarm.name === ENTITLEMENT_REVERIFY_ALARM) {
        chrome.alarms.clear(ENTITLEMENT_REVERIFY_ALARM);
    }
});

const backgroundReady = ensureBackgroundAlarms();
backgroundReady.catch((error) => console.warn('Background: Purchase watch setup failed:', error));
chrome.runtime.onInstalled.addListener(() => backgroundReady);
chrome.runtime.onStartup.addListener(() => backgroundReady);

// Run for all users, including existing Pro users who never open the Free UI.
TeamsBackupDB.migrateFreeBackupUsageToLocal().catch((error) => {
    console.warn('Background: Could not move legacy free-plan usage to local storage:', error);
});
chrome.storage.onChanged.addListener((changes, area) => {
    if (area === 'sync' && changes.freeBackupUsage?.newValue) {
        TeamsBackupDB.migrateFreeBackupUsageToLocal().catch((error) => {
            console.warn('Background: Could not remove legacy synced chat metadata:', error);
        });
    }
});

async function findUserByAppUserId(userId) {
    const normalizedUserId = (userId || '').trim();
    if (!normalizedUserId) {
        return { found: false, userId: null, entitlementResult: null };
    }

    const entitlementResult = await determineAndPersistEntitlement(normalizedUserId);
    if (entitlementResult.isPro) {
        return {
            found: true,
            userId: entitlementResult.userId || normalizedUserId,
            entitlementResult
        };
    }

    return { found: false, userId: null, entitlementResult };
}

function getTeamsBackupDB() {
    const db = globalThis.TeamsBackupDB;
    if (!db || typeof db.saveChat !== 'function') {
        throw new Error('Teams backup database is not available');
    }
    return db;
}

function buildCheckpointChatPayload(request, { finalize = false } = {}) {
    const chatIdentity = request.chatIdentity || {};
    const sourceKind = chatIdentity.sourceKind || (chatIdentity.isChannel ? 'channel' : 'chat');

    return {
        name: chatIdentity.name || request.chatName || 'Teams Chat Backup',
        rawName: chatIdentity.rawName || chatIdentity.name || request.chatName || 'Teams Chat Backup',
        isChannel: chatIdentity.isChannel === true,
        teamName: chatIdentity.teamName || null,
        sourceKind,
        messages: Array.isArray(request.messages) ? request.messages : [],
        participants: Array.isArray(request.participants) ? request.participants : [],
        textContent: finalize ? (request.textContent || '') : '',
        isPreview: finalize ? !!request.partialBackup : true,
        inProgress: !finalize,
        isPro: request.isPro === true,
        allowIncompleteMerge: true,
        existingChatId: request.checkpointChatId || null
    };
}

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (request.action === 'checkpointBackupChunk') {
        (async () => {
            try {
                const messages = Array.isArray(request.messages) ? request.messages : [];
                if (messages.length === 0) {
                    sendResponse({ status: 'error', message: 'Checkpoint batch contained no messages' });
                    return;
                }

                const result = await getTeamsBackupDB().saveChat(
                    buildCheckpointChatPayload(request, { finalize: false })
                );

                if (!result.error && result.id) {
                    chrome.storage.local.set({
                        backupCheckpointUpdated: {
                            chatId: result.id,
                            updatedAt: Date.now(),
                            reason: request.reason || 'checkpoint',
                            batchMessageCount: messages.length
                        }
                    });
                }

                sendResponse({
                    status: result.error ? 'error' : 'success',
                    ...result
                });
            } catch (error) {
                console.error('Background: checkpointBackupChunk failed', error);
                sendResponse({
                    status: 'error',
                    message: error?.message || 'Checkpoint save failed'
                });
            }
        })();

        return true;
    }

    if (request.action === 'finalizeCheckpointBackup') {
        (async () => {
            try {
                const result = await getTeamsBackupDB().saveChat(
                    buildCheckpointChatPayload(request, { finalize: true })
                );

                chrome.storage.local.set({
                    backupInProgress: false,
                    backupCheckpointUpdated: {
                        chatId: result.id || null,
                        updatedAt: Date.now(),
                        reason: 'finalize',
                        finalized: !result.error
                    }
                });
                chrome.storage.local.remove(['backupState']);

                sendResponse({
                    status: result.error ? 'error' : 'success',
                    ...result
                });
            } catch (error) {
                console.error('Background: finalizeCheckpointBackup failed', error);
                sendResponse({
                    status: 'error',
                    message: error?.message || 'Checkpoint finalization failed'
                });
            }
        })();

        return true;
    }

    if (request.action === 'restorePurchaseLookup' || request.action === 'restoreByEmail') {
        const rawIdentifier = request.identifier || request.email || '';
        const identifier = rawIdentifier.trim();

        (async () => {
            if (!identifier) {
                sendResponse({ status: 'error', message: 'No restore identifier provided' });
                return;
            }

            try {
                const result = await findUserByAppUserId(identifier);

                if (result.found) {
                    await stopPurchaseWatch();
                    notifyPurchaseVerified();
                    sendResponse({
                        status: 'success',
                        isPro: true,
                        userId: result.userId,
                        lookupType: 'app_user_id'
                    });
                    return;
                }

                sendResponse({
                    status: 'not_found',
                    lookupType: 'app_user_id',
                    message: result.message || 'No purchase was found for that License ID.'
                });
            } catch (error) {
                console.error('Background: Error restoring purchase:', error);
                sendResponse({ status: 'error', message: error.message });
            }
        })();

        return true;
    }

    if (request.action === 'checkoutOpened') {
        startPurchaseWatch(request.userId)
            .then(sendResponse)
            .catch((error) => sendResponse({ status: 'error', message: error?.message || 'Could not start purchase watch' }));
        return true;
    }

    if (request.action === 'checkPurchaseNow') {
        checkPendingCheckout()
            .then((result) => sendResponse({ status: 'success', isPro: result.isPro === true }))
            .catch((error) => sendResponse({ status: 'error', message: error?.message || 'Purchase check failed' }));
        return true;
    }

    if (request.action === 'checkEntitlements') {
        (async () => {
            try {
                // Compatibility with old extension pages: opening a popup must
                // never cause a remote check, even if it sends forceRemote.
                const result = await getCachedEntitlement();
                sendResponse({
                    status: 'success',
                    isPro: result.isPro,
                    environment: result.environment,
                    provider: result.provider || null
                });
            } catch (error) {
                console.error('Background: Failed to check entitlements:', error);
                sendResponse({
                    status: 'error',
                    message: error.message || 'Unknown entitlement error'
                });
            }
        })();

        return true; // Keep the message channel alive for async response
    }
});
