// license.js - License ID and checkout helpers shared by the popup and the viewer.

// Store builds inject the real checkout link.
const CHECKOUT_BASE_URL = 'https://pay.rev.cat/YOUR_CHECKOUT_ID/';
const SUPPORT_EMAIL = 'support@teamschatbackup.com';

function licenseStorageGet(area, keys) {
    return new Promise((resolve) => {
        chrome.storage[area].get(keys, (items) => {
            resolve(chrome.runtime.lastError ? {} : (items || {}));
        });
    });
}

function licenseStorageSet(area, items) {
    return new Promise((resolve) => {
        chrome.storage[area].set(items, () => resolve(!chrome.runtime.lastError));
    });
}

function buildCheckoutUrl(userId) {
    return `${CHECKOUT_BASE_URL.replace(/\/$/, '')}/${encodeURIComponent(userId)}`;
}

// Returns { userId, fromSync }. fromSync is true when this browser had no License ID
// yet and Chrome sync supplied one, i.e. a new device or a reinstall.
async function getOrCreateLicenseId() {
    const [syncData, localData] = await Promise.all([
        licenseStorageGet('sync', ['userId']),
        licenseStorageGet('local', ['userId', 'isPro', 'purchaseVerified', 'entitlementUserId'])
    ]);

    // Keep the verified identity stable if Chrome sync later brings another ID.
    const activatedId = localData.isPro && localData.purchaseVerified
        ? (localData.entitlementUserId || localData.userId)
        : null;
    if (activatedId) {
        if (localData.userId !== activatedId) await licenseStorageSet('local', { userId: activatedId });
        return { userId: activatedId, fromSync: false };
    }

    if (syncData.userId) {
        if (localData.userId !== syncData.userId) {
            await licenseStorageSet('local', { userId: syncData.userId });
        }
        return { userId: syncData.userId, fromSync: !localData.userId };
    }

    if (localData.userId) {
        await licenseStorageSet('sync', { userId: localData.userId });
        return { userId: localData.userId, fromSync: false };
    }

    const userId = `user_${Date.now()}_${Math.random().toString(36).slice(2, 11)}`;
    await Promise.all([
        licenseStorageSet('local', { userId }),
        licenseStorageSet('sync', { userId })
    ]);
    return { userId, fromSync: false };
}

function sendLicenseMessage(message) {
    return new Promise((resolve) => {
        chrome.runtime.sendMessage(message, (response) => {
            resolve(chrome.runtime.lastError ? null : response);
        });
    });
}

// Opens the hosted checkout and asks the background worker to watch for the
// payment, so Pro unlocks even though the popup closes when the tab opens.
async function openCheckout() {
    const { userId } = await getOrCreateLicenseId();
    const response = await sendLicenseMessage({ action: 'checkoutOpened', userId });
    if (response?.status === 'already_pro') return { userId, alreadyPro: true };
    if (response?.status !== 'watching') {
        throw new Error(response?.message || 'Could not start purchase confirmation');
    }
    await chrome.tabs.create({ url: buildCheckoutUrl(userId) });
    return { userId, alreadyPro: false };
}

function buildSupportMailto(userId, topic = 'Restore my Pro purchase') {
    const body = [
        'Hi,',
        '',
        `${topic}.`,
        userId ? `License ID on this device: ${userId}` : '',
        'Purchase email or receipt: '
    ].filter(line => line !== '').join('\n');
    return `mailto:${SUPPORT_EMAIL}?subject=${encodeURIComponent(`Teams Chat Backup - ${topic}`)}&body=${encodeURIComponent(body)}`;
}

globalThis.TeamsBackupLicense = {
    getOrCreateLicenseId,
    openCheckout,
    buildCheckoutUrl,
    buildSupportMailto,
    SUPPORT_EMAIL
};
