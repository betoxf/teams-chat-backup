import fs from 'node:fs';
import vm from 'node:vm';
import assert from 'node:assert/strict';
import test from 'node:test';

const read = name => fs.readFileSync(new URL('../' + name, import.meta.url), 'utf8');
const active = { subscriber: {
    entitlements: { Teamsbackuppro: { expires_date: null, product_identifier: 'Teamsbackuppro' } },
    non_subscriptions: { Teamsbackuppro: [{ purchase_date: '2026-10-01T00:00:00Z', is_sandbox: false }] }
} };
const paid = { isPro: true, purchaseVerified: true, paymentProvider: 'revenuecat', entitlementUserId: 'paid-license', userId: 'paid-license' };

function env({ local = {}, sync = {}, development = false, alarms = [] } = {}) {
    local = { v6LeastPrivilegeMigrationComplete: true, ...structuredClone(local) };
    sync = structuredClone(sync);
    const listeners = {}, tabs = [], writes = [], alarmMap = new Map(alarms);
    const event = name => ({ addListener(fn) { (listeners[name] ||= []).push(fn); } });
    let context;
    const failures = {};
    const area = (name, data) => ({
        get(keys, cb) {
            const value = Object.fromEntries(keys.filter(k => k in data).map(k => [k, structuredClone(data[k])]));
            cb?.(value); return Promise.resolve(value);
        },
        set(value, cb) {
            if (failures[name + 'Set']) {
                context.chrome.runtime.lastError = { message: 'Storage write failed' };
                cb?.(); delete context.chrome.runtime.lastError; return Promise.resolve();
            }
            Object.assign(data, structuredClone(value)); writes.push({ area: name, value: structuredClone(value) });
            cb?.(); return Promise.resolve();
        },
        remove(keys, cb) {
            if (failures[name + 'Remove']) {
                context.chrome.runtime.lastError = { message: 'Storage removal failed' };
                cb?.(); delete context.chrome.runtime.lastError; return Promise.resolve();
            }
            for (const key of [].concat(keys)) delete data[key];
            cb?.(); return Promise.resolve();
        }
    });
    context = vm.createContext({
        console: { log() {}, warn() {}, error() {} }, importScripts() {}, Date, URL,
        setTimeout, clearTimeout, AbortController,
        TeamsBackupDB: { migrateFreeBackupUsageToLocal: async () => {} },
        chrome: {
            storage: { local: area('local', local), sync: area('sync', sync), onChanged: event('storage') },
            alarms: {
                create: async (name, value) => { alarmMap.set(name, value); },
                clear: async name => alarmMap.delete(name),
                get: async name => alarmMap.get(name), onAlarm: event('alarm')
            },
            runtime: {
                getManifest: () => development ? {} : { update_url: 'https://clients2.google.com/service/update2/crx' },
                getURL: path => 'chrome-extension://test/' + path,
                sendMessage(_message, cb) { cb?.(); },
                onInstalled: event('installed'), onStartup: event('startup'), onMessage: event('message')
            },
            tabs: { create: async value => { tabs.push(value); return { id: 1 }; }, update: async () => { throw new Error('Closed'); } },
            windows: { update: async () => {} }
        }
    });
    const run = expression => vm.runInContext(expression, context);
    return { context, local, sync, tabs, writes, failures, listeners, alarms: alarmMap, run,
        message: request => new Promise(resolve => listeners.message.at(-1)(request, {}, resolve)) };
}

async function background(options) {
    const e = env(options);
    e.run(read('background.js'));
    await e.run('backgroundReady');
    e.context.fixture = { subscriber: { entitlements: {}, non_subscriptions: {} } };
    e.context.requests = 0;
    e.run('fetchSubscriberData = async () => { requests++; return fixture; }');
    return e;
}

test('checkout confirmation activates once and stops all purchase checks', async () => {
    const e = await background();
    e.context.fixture = active;
    assert.equal((await e.message({ action: 'checkoutOpened', userId: 'new-license' })).status, 'watching');
    assert.equal(e.alarms.get('purchaseWatch').periodInMinutes, 0.5);
    await e.run('checkPendingCheckout()');
    assert.equal(e.context.requests, 1);
    assert.equal(e.local.isPro, true);
    assert.equal(e.local.purchaseVerified, true);
    assert.equal(e.local.entitlementUserId, 'new-license');
    assert.equal(e.local.pendingCheckout, undefined);
    assert.equal(e.alarms.size, 0);
    assert.match(e.tabs[0].url, /viewer.html\?pro=activated$/);
    e.run('fetchSubscriberData = async () => { requests++; throw new Error("offline"); }');
    await e.run('checkPendingCheckout()');
    await e.message({ action: 'checkEntitlements', userId: 'new-license', forceRemote: true });
    await e.message({ action: 'restorePurchaseLookup', identifier: 'new-license' });
    assert.equal(e.context.requests, 1);
    assert.equal(e.local.isPro, true);
});

test('startup preserves paid users and removes v6.6 alarms and revocation state', async () => {
    for (const provider of ['revenuecat', 'legacy']) {
        const e = await background({
            local: { ...paid, paymentProvider: provider, pendingCheckout: { userId: 'paid-license', startedAt: Date.now() },
                entitlementMissingSince: 1, lastEntitlementCheckAt: 1, checkEntitlementsOnOpen: true },
            alarms: [['entitlementReverify', {}], ['purchaseWatch', {}]]
        });
        await Promise.all(e.listeners.startup.map(fn => fn()));
        await e.listeners.alarm[0]({ name: 'entitlementReverify' });
        assert.equal(e.local.isPro, true);
        assert.equal(e.local.purchaseVerified, true);
        assert.equal(e.local.entitlementMissingSince, undefined);
        assert.equal(e.local.checkEntitlementsOnOpen, undefined);
        assert.equal(e.alarms.size, 0);
        assert.equal(e.context.requests, 0);
    }
});

test('expired or invalid watches make zero requests, including legacy check messages', async () => {
    for (const startedAt of [Date.now() - 1800001, Date.now() + 86400000, undefined]) {
        const e = await background({ local: { pendingCheckout: { userId: 'unpaid', startedAt } } });
        await e.message({ action: 'checkPurchaseNow', userId: 'unpaid' });
        assert.equal(e.context.requests, 0);
        assert.equal(e.local.pendingCheckout, undefined);
        assert.equal(e.alarms.size, 0);
    }
    const e = await background();
    await e.message({ action: 'checkoutOpened', userId: 'unpaid' });
    e.local.pendingCheckout.startedAt = Date.now() - 1800001;
    await e.run('checkPendingCheckout()');
    assert.equal(e.context.requests, 0);
    assert.equal(e.alarms.size, 0);
});

test('browser startup resumes only the remainder of an active checkout', async () => {
    const startedAt = Date.now() - 600000;
    const e = await background({ local: { pendingCheckout: { userId: 'pending', startedAt } } });
    assert.equal(e.alarms.get('purchaseWatch').periodInMinutes, 0.5);
    assert.equal(e.local.pendingCheckout.startedAt, startedAt);
    assert.equal(e.context.requests, 0);
    await e.run('checkPendingCheckout()');
    assert.equal(e.context.requests, 1);
    assert.notEqual(e.local.isPro, true);
});

test('unpaid checkout failures keep the watch bounded and grant no access', async () => {
    const e = await background();
    await e.run("startPurchaseWatch('unpaid')");
    e.run('fetchSubscriberData = async () => { requests++; throw new Error("offline"); }');
    await e.run('checkPendingCheckout()');
    assert.notEqual(e.local.isPro, true);
    assert.ok(e.local.pendingCheckout);
    e.local.pendingCheckout.startedAt = Date.now() - 1800001;
    await e.run('checkPendingCheckout()');
    assert.equal(e.local.pendingCheckout, undefined);
});

test('explicit Restore activates a new profile once; opening or syncing does not', async () => {
    const e = await background({ local: { userId: 'synced-license' }, sync: { userId: 'synced-license' } });
    e.context.fixture = active;
    assert.equal((await e.message({ action: 'checkEntitlements', userId: 'synced-license', forceRemote: true })).isPro, false);
    assert.equal(e.context.requests, 0);
    const response = await e.message({ action: 'restorePurchaseLookup', identifier: 'synced-license' });
    assert.equal(response.isPro, true);
    assert.equal(e.context.requests, 1);
    assert.equal(e.alarms.size, 0);
    const cached = await e.message({ action: 'restorePurchaseLookup', identifier: 'different-unverified-id' });
    assert.equal(cached.userId, 'synced-license');
    assert.equal(e.local.entitlementUserId, 'synced-license');
    assert.equal(e.context.requests, 1);
});

test('invalid Restore grants no access and starts no background polling', async () => {
    const e = await background();
    const response = await e.message({ action: 'restorePurchaseLookup', identifier: 'invalid' });
    assert.equal(response.status, 'not_found');
    assert.notEqual(e.local.purchaseVerified, true);
    assert.equal(e.alarms.size, 0);
});

test('concurrent checks for the same checkout share one request', async () => {
    const e = await background();
    e.context.fixture = active;
    await Promise.all([e.run("determineAndPersistEntitlement('same')"), e.run("determineAndPersistEntitlement('same')")]);
    assert.equal(e.context.requests, 1);
});

test('already activated users cannot open another checkout', async () => {
    const e = await background({ local: paid, sync: { userId: 'another-synced-id' } });
    e.context.chrome.runtime.sendMessage = (request, cb) => { e.message(request).then(cb); };
    e.run(read('license.js'));
    assert.equal((await e.run('TeamsBackupLicense.openCheckout()')).alreadyPro, true);
    assert.equal((await e.run('TeamsBackupLicense.getOrCreateLicenseId()')).userId, paid.userId);
    assert.equal(e.local.userId, paid.userId);
    assert.equal(e.tabs.length, 0);
    assert.equal(e.context.requests, 0);
    assert.equal(e.alarms.size, 0);
});

test('checkout opens with the License ID only when purchase confirmation starts', async () => {
    const e = await background({ local: { userId: 'new-license' } });
    e.context.chrome.runtime.sendMessage = (request, cb) => { e.message(request).then(cb); };
    e.run(read('license.js'));
    assert.equal((await e.run('TeamsBackupLicense.openCheckout()')).alreadyPro, false);
    assert.equal(new URL(e.tabs[0].url).hostname, 'pay.rev.cat');
    assert.ok(e.tabs[0].url.endsWith('/new-license'));
    assert.equal(e.local.pendingCheckout.userId, 'new-license');
});

test('store builds reject sandbox purchases; unpacked QA builds accept them', async () => {
    for (const development of [false, true]) {
        const e = await background({ development });
        e.context.fixture = structuredClone(active);
        e.context.fixture.subscriber.non_subscriptions.Teamsbackuppro[0].is_sandbox = true;
        assert.equal((await e.run("determineAndPersistEntitlement('sandbox')")).isPro, development);
    }
});

function database(options = {}) {
    const e = env(options);
    e.run(read('db.js'));
    e.context.existingChats = [];
    e.run('openDB = async () => ({}); getAllChatsRaw = async () => existingChats; getChatCount = async () => 0;');
    return e;
}

test('legacy synced chat names move locally and are removed from sync', async () => {
    const e = database({
        local: { freeBackupUsageSeeded: true, freeBackupUsage: { count: 1, chats: [{ name: 'Local chat' }] } },
        sync: { userId: 'license', freeBackupUsage: { count: 3, chats: [{ name: 'Synced chat', teamName: 'Private team', isChannel: true }] } }
    });
    await e.run('migrateFreeBackupUsageToLocal()');
    assert.equal(e.sync.freeBackupUsage, undefined);
    assert.equal(e.sync.userId, 'license');
    assert.equal(e.local.freeBackupUsage.count, 3);
    assert.deepEqual(e.local.freeBackupUsage.chats.map(c => c.name), ['Local chat', 'Synced chat']);
    await e.run('getFreeBackupUsage()');
    assert.ok(e.writes.every(w => w.area === 'local'));
});

test('migration does not delete the old allowance until the local save succeeds', async () => {
    const e = database({ sync: { freeBackupUsage: { count: 1, chats: [{ name: 'Existing chat' }] } } });
    e.failures.localSet = true;
    await assert.rejects(e.run('migrateFreeBackupUsageToLocal()'), /locally/);
    assert.equal(e.sync.freeBackupUsage.count, 1);
    e.failures.localSet = false;
    e.failures.syncRemove = true;
    await assert.rejects(e.run('migrateFreeBackupUsageToLocal()'), /removal/);
    assert.equal(e.local.freeBackupUsage.count, 1);
    e.failures.syncRemove = false;
    await e.run('migrateFreeBackupUsageToLocal()');
    assert.equal(e.sync.freeBackupUsage, undefined);
    assert.equal(e.local.freeBackupUsage.count, 1);
});

test('existing and trashed chats seed local usage without any sync writes', async () => {
    const e = database({ local: { freeBackupUsageSeeded: true } });
    e.context.existingChats = [{ name: 'Saved chat', createdAt: 10 }, { name: 'Trashed chat', trashedAt: 20 }];
    const usage = await e.run('getFreeBackupUsage()');
    assert.equal(usage.count, 2);
    assert.ok(e.writes.every(w => w.area === 'local'));
    assert.equal((await e.run("getFreeBackupAllowance({name:'Different chat'})")).allowed, false);
});

test('deleting a backup keeps the local allowance; same-chat duplicates do not consume another slot', async () => {
    const e = database();
    await e.run("recordFreeBackupUsage({name:'Chosen chat'})");
    await e.run("recordFreeBackupUsage({name:'Chosen chat'})");
    assert.equal(e.local.freeBackupUsage.count, 1);
    assert.equal((await e.run("getFreeBackupAllowance({name:'Chosen chat'})")).allowed, true);
    assert.equal((await e.run("getFreeBackupAllowance({name:'Other chat'})")).allowed, false);
    assert.ok(e.writes.every(w => w.area === 'local'));
});

test('popup has no automatic or periodic remote-verification entry points', () => {
    const popup = read('popup.js');
    assert.doesNotMatch(popup, /forceRemote|restoreProFromSyncedLicense|checkPurchaseNow|PENDING_PURCHASE_MAX_AGE_MS/);
    assert.doesNotMatch(popup, /action:\s*['"]checkEntitlements['"]/);
});
