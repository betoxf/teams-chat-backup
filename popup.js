// popup.js

const FORCE_VIEWER_EXPORT_ONLY = true; // Always text-only backup in popup; PDF/TXT export happens in viewer
// The PDF libraries are only needed for popup PDF capture, which is off while
// exports happen in the viewer; skipping them keeps injection into Teams light.
const TEAMS_CONTENT_SCRIPT_FILES = FORCE_VIEWER_EXPORT_ONLY
    ? ['content.js']
    : [
        'libraries/pdfobject.min.js',
        'libraries/html2canvas.min.js',
        'libraries/jspdf.min.js',
        'content.js'
    ];
const POPUP_FREE_CHAT_LIMIT = window.TeamsBackupDB?.FREE_CHAT_LIMIT ?? 1;
const LOAD_AVAILABLE_CHATS_TIMEOUT_MS = 25000; // Separate Chat/Teams views take ~20 s to scan the first time
const BACKUP_RUNTIME_RECONNECT_DELAY_MS = 1200;
const BACKUP_RUNTIME_RECONNECT_TIMEOUT_MS = 12000;
const BACKUP_RUNTIME_RECONNECT_POLL_MS = 500;
const BACKUP_STALE_RECOVERY_TIMEOUT_MS = 15000;
const BACKUP_STALE_RECOVERY_POLL_MS = 500;
const ENABLE_TEST_MODE = false; // Set to true to test Pro features without purchasing
const PURCHASE_WATCH_WINDOW_MS = 30 * 60 * 1000; // Background checks a new checkout for 30 minutes
const ISSUE_REPORT_EMAIL = 'support@teamschatbackup.com';
const ISSUE_REPORT_FALLBACK_URL = 'https://bullfighterstudio.com/teams-backup#feedback';
const CLOSED_ASYNC_RESPONSE_ERROR_RE = /A listener indicated an asynchronous response by returning true, but the message channel closed before a response was received/i;
const CLOSED_MESSAGE_PORT_ERROR_RE = /message (?:port|channel) closed before a response was received/i;
const NO_RECEIVER_RUNTIME_ERROR_RE = /Receiving end does not exist|Could not establish connection/i;
const VIEWER_TAB_STORAGE_KEY = 'viewerTabId';

document.addEventListener('DOMContentLoaded', async () => {
    // Display version from manifest
    const versionBadge = document.getElementById('versionBadge');
    if (versionBadge) {
        const manifest = chrome.runtime.getManifest();
        versionBadge.textContent = `v${manifest.version}`;
    }

    const backupButton = document.getElementById('backupButton');
    const stopButton = document.getElementById('stopButton');
    const statusDiv = document.getElementById('status');
    const loadingMessageDiv = document.getElementById('loadingMessage');
    const progressContainer = document.getElementById('progressContainer');
    const progressBar = document.getElementById('progressBar');
    const downloadContainer = document.getElementById('downloadContainer');
    const backupHint = document.getElementById('backupHint');
    const downloadButton = document.getElementById('downloadButton');
    const viewBackupButton = document.getElementById('viewBackupButton');
    const upgradeSection = document.getElementById('upgradeSection');
    const purchaseButton = document.getElementById('purchaseButton');
    const includeCoverCheckbox = document.getElementById('includeCover');
    const coverToggleContainer = document.getElementById('coverToggleContainer');
    const purchasePollingMessage = document.getElementById('purchasePollingMessage');
    const proBadge = document.getElementById('proBadge');
    const freeBadge = document.getElementById('freeBadge');
    const restoreButton = document.getElementById('restoreButton');
    const emailModal = document.getElementById('emailModal');
    const emailInput = document.getElementById('emailInput');
    const emailError = document.getElementById('emailError');
    const emailCancelBtn = document.getElementById('emailCancelBtn');
    const emailSubmitBtn = document.getElementById('emailSubmitBtn');
    const appUserIdValue = document.getElementById('appUserIdValue');
    const copyUserIdButton = document.getElementById('copyUserIdButton');
    const restoreSupportLink = document.getElementById('restoreSupportLink');
    const settingsLicenseSection = document.getElementById('settingsLicenseSection');
    const settingsLicenseIdValue = document.getElementById('settingsLicenseIdValue');
    const settingsCopyLicenseButton = document.getElementById('settingsCopyLicenseButton');
    const settingsSupportLink = document.getElementById('settingsSupportLink');

    // Settings elements
    const settingsBtn = document.getElementById('settingsBtn');
    const settingsPanel = document.getElementById('settingsPanel');
    const closeSettingsBtn = document.getElementById('closeSettingsBtn');
    const captureSpeedSlider = document.getElementById('captureSpeedSlider');
    const captureSpeedValue = document.getElementById('captureSpeedValue');
    const turboModeCheckbox = document.getElementById('turboModeCheckbox');
    const turboModeSection = document.getElementById('turboModeSection');
    const longChatSection = document.getElementById('longChatSection');
    const captureSpeedSection = document.getElementById('captureSpeedSection');
    const turboOverrideHint = document.getElementById('turboOverrideHint');
    const viewerOnlyModeNotice = document.getElementById('viewerOnlyModeNotice');
    const settingsNewDot = document.getElementById('settingsNewDot');
    const settingsIncludeCover = document.getElementById('settingsIncludeCover');
    const settingsCoverToggle = document.getElementById('settingsCoverToggle');

    // View backed files button
    const viewBackedFilesButton = document.getElementById('viewBackedFilesButton');
    const backedFilesCount = document.getElementById('backedFilesCount');

    // Chat selection elements
    const chatSelectionContainer = document.getElementById('chatSelectionContainer');
    const chatList = document.getElementById('chatList');
    const freeUserNotice = document.getElementById('freeUserNotice');
    const startBackupBtn = document.getElementById('startBackupBtn');
    const backupButtonText = document.querySelector('.backup-button__text');

    const PRO_CELEBRATION_KEY = 'proCelebrationAcknowledged';
    const PRO_INCLUDE_COVER_KEY = 'proIncludeCover';
    const PRO_SUCCESS_MESSAGE = 'Purchase successful! You now have Pro access!';

    // Settings storage keys
    const SETTINGS_CAPTURE_SPEED_KEY = 'settingsCaptureSpeed';
    const SETTINGS_TURBO_MODE_KEY = 'settingsTurboMode';
    const SETTINGS_LONG_CHAT_MODE_KEY = 'settingsLongChatMode';
    const SETTINGS_FILE_ATTACHMENTS_KEY = 'settingsFileAttachments';
    const SETTINGS_FILE_ATTACHMENTS_DEFAULT_OFF_MIGRATION_KEY = 'settingsFileAttachmentsDefaultOffMigrationApplied';
    const SETTINGS_MAX_FILE_SIZE_KEY = 'settingsMaxFileSize';
    const SETTINGS_OPENED_KEY = 'settingsOpened';

    // Backup state storage keys (for persistence across popup close/reopen)
    const BACKUP_STATE_KEY = 'backupState'; // { chatName, progress, phase, startTime, lastUpdate }

    // Speed labels for the slider
    const SPEED_LABELS = ['Very Slow', 'Slow', 'Normal', 'Fast', 'Very Fast'];
    // Capture intervals in seconds (higher = slower/safer)
    const SPEED_INTERVALS = [3.0, 2.0, 1.0, 0.5, 0.25];

    let downloadedPdfUrl = null;
    let downloadedTextUrl = null;
    let isPro = false;
    // autoStopTimer removed — free users now get full backups
    let backupInProgress = false;
    let cachedUserId = null;
    let storedProIncludeCover = null;
    let proCelebrationAcknowledged = false;
    let lastSavedChatId = null; // Track the last saved chat for "View in Viewer" button
    let availableChats = []; // Store available chats from Teams
    let chatLoadDiagnostics = null; // Sidebar/load diagnostics from content script
    let selectedChatIndices = []; // Track selected chat indices for backup
    let currentBackupChatName = null; // Track current backup's chat name
    let currentBackupStartTime = null; // Track when current backup started
    let monitoringDetachedBackup = false; // Popup reopened while content script backup is still running

    // Settings state
    let storedCaptureSpeed = 3; // Default: Normal (index 2, but slider is 1-5)
    let storedTurboMode = false; // Default: off (generate PDF with screenshots)
    let storedLongChatMode = false; // Default: off (normal scrolling waits)
    let storedFileAttachments = false; // Default: off (faster backups)
    let storedMaxFileSize = 5; // Default: 5 MB max file size
    let settingsOpened = false;

    // Long Chat Mode checkbox element
    const longChatModeCheckbox = document.getElementById('longChatModeCheckbox');

    // File Attachments settings elements
    const fileAttachmentsCheckbox = document.getElementById('fileAttachmentsCheckbox');
    const fileAttachmentsSection = document.getElementById('fileAttachmentsSection');
    const fileSettingsDetails = document.getElementById('fileSettingsDetails');
    const maxFileSizeSlider = document.getElementById('maxFileSizeSlider');
    const maxFileSizeHint = document.getElementById('maxFileSizeHint');

    function addSafeListener(element, eventName, handler, contextLabel) {
        if (!element) {
            console.warn(`Skipping ${eventName} listener${contextLabel ? ` for ${contextLabel}` : ''}: element not found`);
            return;
        }
        element.addEventListener(eventName, handler);
    }

    // Enhanced storage functions with retry logic and error handling (P1-2)
    const MAX_STORAGE_RETRIES = 3;
    const STORAGE_RETRY_DELAY_MS = 100;

    async function readStorage(area, keys) {
        let lastError = null;
        for (let attempt = 0; attempt < MAX_STORAGE_RETRIES; attempt++) {
            try {
                return await new Promise((resolve, reject) => {
                    const storageArea = chrome.storage?.[area];
                    if (!storageArea) {
                        reject(new Error(`Storage area '${area}' not available`));
                        return;
                    }
                    storageArea.get(keys, (items) => {
                        if (chrome.runtime.lastError) {
                            reject(new Error(chrome.runtime.lastError.message));
                            return;
                        }
                        resolve(items || {});
                    });
                });
            } catch (error) {
                lastError = error;
                console.warn(`Storage (${area}) read attempt ${attempt + 1}/${MAX_STORAGE_RETRIES} failed:`, error.message);

                // Check if it's a quota error (no retry for quota)
                if (error.message.includes('QUOTA') || error.message.includes('quota')) {
                    console.error(`Storage quota exceeded for ${area}`);
                    // Try fallback to sessionStorage for local area
                    if (area === 'local') {
                        try {
                            const result = {};
                            for (const key of Array.isArray(keys) ? keys : Object.keys(keys)) {
                                const value = sessionStorage.getItem(key);
                                if (value !== null) {
                                    try {
                                        result[key] = JSON.parse(value);
                                    } catch {
                                        result[key] = value;
                                    }
                                }
                            }
                            console.log('Fallback to sessionStorage successful for read');
                            return result;
                        } catch (sessionError) {
                            console.error('SessionStorage fallback also failed:', sessionError);
                        }
                    }
                    break;
                }

                // Exponential backoff for retries
                if (attempt < MAX_STORAGE_RETRIES - 1) {
                    await new Promise(r => setTimeout(r, STORAGE_RETRY_DELAY_MS * Math.pow(2, attempt)));
                }
            }
        }

        // All retries failed
        console.error(`Storage (${area}) read failed after ${MAX_STORAGE_RETRIES} attempts:`, lastError);
        return {}; // Return empty object as fallback
    }

    async function writeStorage(area, data) {
        let lastError = null;
        for (let attempt = 0; attempt < MAX_STORAGE_RETRIES; attempt++) {
            try {
                return await new Promise((resolve, reject) => {
                    const storageArea = chrome.storage?.[area];
                    if (!storageArea) {
                        reject(new Error(`Storage area '${area}' not available`));
                        return;
                    }
                    storageArea.set(data, () => {
                        if (chrome.runtime.lastError) {
                            reject(new Error(chrome.runtime.lastError.message));
                            return;
                        }
                        resolve(true);
                    });
                });
            } catch (error) {
                lastError = error;
                console.warn(`Storage (${area}) write attempt ${attempt + 1}/${MAX_STORAGE_RETRIES} failed:`, error.message);

                // Check if it's a quota error
                if (error.message.includes('QUOTA') || error.message.includes('quota')) {
                    console.error(`Storage quota exceeded for ${area}`);
                    // Try fallback to sessionStorage for local area
                    if (area === 'local') {
                        try {
                            for (const [key, value] of Object.entries(data)) {
                                sessionStorage.setItem(key, JSON.stringify(value));
                            }
                            console.log('Fallback to sessionStorage successful for write');
                            return true;
                        } catch (sessionError) {
                            console.error('SessionStorage fallback also failed:', sessionError);
                        }
                    }
                    // Show user-friendly message for quota errors
                    if (attempt === 0) { // Only show once
                        showToast('Storage almost full. Some data may not be saved.', 5000, 'warning');
                    }
                    break;
                }

                // Exponential backoff for retries
                if (attempt < MAX_STORAGE_RETRIES - 1) {
                    await new Promise(r => setTimeout(r, STORAGE_RETRY_DELAY_MS * Math.pow(2, attempt)));
                }
            }
        }

        // All retries failed - log but don't block UI
        console.error(`Storage (${area}) write failed after ${MAX_STORAGE_RETRIES} attempts:`, lastError);
        return false;
    }

    function applyProState({ celebrate = false } = {}) {
        isPro = true;
        upgradeSection.style.display = 'none';
        if (settingsLicenseSection) settingsLicenseSection.style.display = 'block';
        coverToggleContainer.style.display = FORCE_VIEWER_EXPORT_ONLY ? 'none' : 'flex';
        purchasePollingMessage.style.display = 'none';
        if (proBadge) proBadge.style.display = 'inline-flex';
        if (freeBadge) freeBadge.style.display = 'none';
        document.body.classList.add('pro-theme');

        const coverPreference = typeof storedProIncludeCover === 'boolean' ? storedProIncludeCover : false;
        includeCoverCheckbox.checked = coverPreference;
        // Also sync settings panel cover checkbox
        if (settingsIncludeCover) settingsIncludeCover.checked = coverPreference;
        // Show cover toggle in settings for Pro users (re-enable if was disabled for free)
        if (settingsCoverToggle && !FORCE_VIEWER_EXPORT_ONLY) {
            settingsCoverToggle.style.display = 'block';
            const coverCheckbox = settingsCoverToggle.querySelector('input[type="checkbox"]');
            if (coverCheckbox) coverCheckbox.disabled = false;
            const coverHint = settingsCoverToggle.querySelector('.settings-hint');
            if (coverHint) coverHint.textContent = 'Pro feature: Add branded pages to PDF';
        }
        // Re-enable file attachments for Pro users
        if (fileAttachmentsSection) {
            const fileCheckbox = fileAttachmentsSection.querySelector('input[type="checkbox"]');
            if (fileCheckbox) fileCheckbox.disabled = false;
            const fileHint = fileAttachmentsSection.querySelector('.settings-hint');
            if (fileHint) fileHint.textContent = 'Download shared files (PDF, Word, Excel, etc.) and include them in backups.';
        }
        if (typeof storedProIncludeCover !== 'boolean') {
            storedProIncludeCover = coverPreference;
            writeStorage('local', { [PRO_INCLUDE_COVER_KEY]: coverPreference });
            writeStorage('sync', { [PRO_INCLUDE_COVER_KEY]: coverPreference });
        }

        enforceViewerExportOnlyMode();

        if (celebrate) {
            statusDiv.textContent = PRO_SUCCESS_MESSAGE;
            statusDiv.style.color = '#f1c40f';
            statusDiv.style.display = 'block';
            proCelebrationAcknowledged = true;
            writeStorage('local', { [PRO_CELEBRATION_KEY]: true });
        } else if (statusDiv.textContent === PRO_SUCCESS_MESSAGE) {
            statusDiv.style.display = 'none';
        }
    }

    function applyFreeState() {
        isPro = false;
        upgradeSection.style.display = 'block';
        if (settingsLicenseSection) settingsLicenseSection.style.display = 'none';
        coverToggleContainer.style.display = 'none';
        if (proBadge) proBadge.style.display = 'none';
        if (freeBadge) freeBadge.style.display = 'inline-flex';
        document.body.classList.remove('pro-theme');
        includeCoverCheckbox.checked = true;
        // Show cover toggle in settings but disabled with PRO badge for free users
        if (settingsCoverToggle && !FORCE_VIEWER_EXPORT_ONLY) {
            settingsCoverToggle.style.display = 'block';
            const coverCheckbox = settingsCoverToggle.querySelector('input[type="checkbox"]');
            if (coverCheckbox) {
                coverCheckbox.disabled = true;
                coverCheckbox.checked = true;
            }
            const coverHint = settingsCoverToggle.querySelector('.settings-hint');
            if (coverHint) coverHint.textContent = 'PRO — Upgrade to toggle cover & footer pages';
        }
        // Show file attachments but disabled with PRO badge for free users
        if (fileAttachmentsSection) {
            const fileCheckbox = fileAttachmentsSection.querySelector('input[type="checkbox"]');
            if (fileCheckbox) {
                fileCheckbox.disabled = true;
                fileCheckbox.checked = false;
            }
            const fileHint = fileAttachmentsSection.querySelector('.settings-hint');
            if (fileHint) fileHint.textContent = 'PRO — Upgrade to include file attachments in backups';
        }

        enforceViewerExportOnlyMode();
    }

    function enforceViewerExportOnlyMode() {
        if (!FORCE_VIEWER_EXPORT_ONLY) return;

        // Backups from popup are always text-first; PDF is generated on-demand in viewer.
        storedTurboMode = true;
        includeCoverCheckbox.checked = false;

        if (viewerOnlyModeNotice) {
            viewerOnlyModeNotice.style.display = 'block';
        }

        if (turboModeCheckbox) {
            turboModeCheckbox.checked = true;
            turboModeCheckbox.disabled = true;
        }

        if (turboModeSection) {
            turboModeSection.style.display = 'none';
            const turboHint = turboModeSection.querySelector('.settings-hint');
            if (turboHint) {
                turboHint.textContent = 'Popup backups run in text-only mode. Export PDF/TXT from Viewer.';
            }
        }

        if (captureSpeedSlider) captureSpeedSlider.disabled = true;
        if (captureSpeedSection) {
            captureSpeedSection.classList.add('disabled');
            captureSpeedSection.style.display = 'none';
        }
        if (turboOverrideHint) {
            turboOverrideHint.textContent = 'PDF capture is disabled in popup. Use Viewer download options.';
            turboOverrideHint.style.display = 'block';
        }

        if (settingsCoverToggle) {
            settingsCoverToggle.style.display = 'none';
        }
    }

    // ==================== Backup State Persistence ====================

    // Save backup state to storage (progress, phase, chat name) - with error handling (P1-2)
    async function saveBackupState(state) {
        const backupState = {
            chatName: state.chatName || 'Unknown Chat',
            progress: state.progress || 0,
            phase: state.phase || 'Starting',
            startTime: state.startTime || Date.now(),
            lastUpdate: Date.now(),
            isPro: isPro
        };
        const success = await writeStorage('local', { [BACKUP_STATE_KEY]: backupState });
        if (!success) {
            console.warn('Failed to save backup state to chrome.storage, state may not persist after popup close');
            // Try sessionStorage as fallback
            try {
                sessionStorage.setItem(BACKUP_STATE_KEY, JSON.stringify(backupState));
            } catch (e) {
                console.error('SessionStorage fallback also failed:', e);
            }
        }
    }

    // Clear backup state from storage
    function clearBackupState() {
        chrome.storage.local.remove([BACKUP_STATE_KEY, 'backupInProgress']);
    }

    // Format time elapsed for display
    function formatTimeElapsed(ms) {
        const seconds = Math.floor(ms / 1000);
        if (seconds < 60) return `${seconds}s`;
        const minutes = Math.floor(seconds / 60);
        const remainingSeconds = seconds % 60;
        return `${minutes}m ${remainingSeconds}s`;
    }

    // ==================== Settings Functions ====================

    function openSettings() {
        if (!settingsPanel) return;
        settingsPanel.style.display = 'flex';
        // Show cover toggle for Pro users
        if (isPro && settingsCoverToggle && !FORCE_VIEWER_EXPORT_ONLY) {
            settingsCoverToggle.style.display = 'block';
        }
        // Hide red dot when settings is opened for the first time
        if (!settingsOpened) {
            settingsNewDot.style.display = 'none';
            settingsOpened = true;
            writeStorage('local', { [SETTINGS_OPENED_KEY]: true });
        }
        enforceViewerExportOnlyMode();
    }

    function closeSettings() {
        if (!settingsPanel) return;
        settingsPanel.style.display = 'none';
    }

    function updateSpeedLabel(value) {
        const index = parseInt(value) - 1; // slider is 1-5, array is 0-4
        if (captureSpeedValue) {
            captureSpeedValue.textContent = SPEED_LABELS[index] || 'Normal';
        }
    }

    function getCaptureInterval() {
        const index = storedCaptureSpeed - 1; // slider is 1-5, array is 0-4
        return SPEED_INTERVALS[index] || 1.0;
    }

    async function loadSettings() {
        const hasSettingsUi = captureSpeedSlider && turboModeCheckbox && longChatModeCheckbox &&
            fileAttachmentsCheckbox && fileSettingsDetails && maxFileSizeSlider &&
            maxFileSizeHint && settingsNewDot;
        if (!hasSettingsUi) {
            console.warn('Settings UI not found. Skipping settings initialization.');
            return;
        }

        const data = await readStorage('local', [
            SETTINGS_CAPTURE_SPEED_KEY,
            SETTINGS_TURBO_MODE_KEY,
            SETTINGS_LONG_CHAT_MODE_KEY,
            SETTINGS_FILE_ATTACHMENTS_KEY,
            SETTINGS_FILE_ATTACHMENTS_DEFAULT_OFF_MIGRATION_KEY,
            SETTINGS_MAX_FILE_SIZE_KEY,
            SETTINGS_OPENED_KEY
        ]);

        // Load capture speed
        if (typeof data[SETTINGS_CAPTURE_SPEED_KEY] === 'number') {
            storedCaptureSpeed = data[SETTINGS_CAPTURE_SPEED_KEY];
        } else {
            storedCaptureSpeed = 3; // Default: Normal
        }
        captureSpeedSlider.value = storedCaptureSpeed;
        updateSpeedLabel(storedCaptureSpeed);

        // Load turbo mode setting
        if (FORCE_VIEWER_EXPORT_ONLY) {
            storedTurboMode = true;
        } else if (typeof data[SETTINGS_TURBO_MODE_KEY] === 'boolean') {
            storedTurboMode = data[SETTINGS_TURBO_MODE_KEY];
        } else {
            storedTurboMode = true; // Default: on (fast mode)
        }
        turboModeCheckbox.checked = storedTurboMode;
        updateTurboModeUI(storedTurboMode);

        // Load long chat mode setting
        if (typeof data[SETTINGS_LONG_CHAT_MODE_KEY] === 'boolean') {
            storedLongChatMode = data[SETTINGS_LONG_CHAT_MODE_KEY];
        } else {
            storedLongChatMode = false; // Default: off
        }
        longChatModeCheckbox.checked = storedLongChatMode;
        updateLongChatModeUI(storedLongChatMode);

        // Load file attachments setting
        const hasAppliedDefaultOffMigration = data[SETTINGS_FILE_ATTACHMENTS_DEFAULT_OFF_MIGRATION_KEY] === true;
        if (!hasAppliedDefaultOffMigration) {
            // One-time migration: disable file downloads by default for existing users too.
            storedFileAttachments = false;
            await writeStorage('local', {
                [SETTINGS_FILE_ATTACHMENTS_KEY]: false,
                [SETTINGS_FILE_ATTACHMENTS_DEFAULT_OFF_MIGRATION_KEY]: true
            });
        } else if (typeof data[SETTINGS_FILE_ATTACHMENTS_KEY] === 'boolean') {
            storedFileAttachments = data[SETTINGS_FILE_ATTACHMENTS_KEY];
        } else {
            storedFileAttachments = false; // Default: off
        }
        fileAttachmentsCheckbox.checked = storedFileAttachments;
        updateFileAttachmentsUI(storedFileAttachments);

        // Load max file size setting
        if (typeof data[SETTINGS_MAX_FILE_SIZE_KEY] === 'number') {
            storedMaxFileSize = data[SETTINGS_MAX_FILE_SIZE_KEY];
        } else {
            storedMaxFileSize = 5; // Default: 5 MB
        }
        maxFileSizeSlider.value = storedMaxFileSize;
        updateMaxFileSizeHint(storedMaxFileSize);

        // Load settings opened state and show red dot if first time
        if (typeof data[SETTINGS_OPENED_KEY] === 'boolean') {
            settingsOpened = data[SETTINGS_OPENED_KEY];
        }
        // Show red dot if settings hasn't been opened yet
        if (!settingsOpened) {
            settingsNewDot.style.display = 'block';
        }

        enforceViewerExportOnlyMode();
    }

    // Update UI when turbo mode changes
    function updateTurboModeUI(isTurbo) {
        if (FORCE_VIEWER_EXPORT_ONLY) {
            if (captureSpeedSlider) captureSpeedSlider.disabled = true;
            if (captureSpeedSection) captureSpeedSection.classList.add('disabled');
            if (turboOverrideHint) {
                turboOverrideHint.textContent = 'PDF capture is disabled in popup. Use Viewer download options.';
                turboOverrideHint.style.display = 'block';
            }
            return;
        }

        // Don't update if long chat mode is on (it takes precedence)
        if (storedLongChatMode) return;

        if (isTurbo) {
            // Disable capture speed slider when turbo mode is on
            captureSpeedSlider.disabled = true;
            captureSpeedSection.classList.add('disabled');
            turboOverrideHint.style.display = 'block';
        } else {
            // Enable capture speed slider when turbo mode is off
            captureSpeedSlider.disabled = false;
            captureSpeedSection.classList.remove('disabled');
            turboOverrideHint.style.display = 'none';
        }
    }

    // Update UI when long chat mode changes - it disables all other settings
    function updateLongChatModeUI(isLongChat) {
        const turboSection = turboModeCheckbox.closest('.settings-section');
        const coverSection = settingsCoverToggle;

        if (FORCE_VIEWER_EXPORT_ONLY) {
            turboModeCheckbox.disabled = true;
            if (turboSection) turboSection.classList.add('disabled');
            if (captureSpeedSlider) captureSpeedSlider.disabled = true;
            if (captureSpeedSection) captureSpeedSection.classList.add('disabled');
            if (turboOverrideHint) {
                turboOverrideHint.textContent = 'PDF capture is disabled in popup. Use Viewer download options.';
                turboOverrideHint.style.display = 'block';
            }
            if (coverSection) coverSection.style.display = 'none';
            return;
        }

        if (isLongChat) {
            // Disable turbo mode section
            turboModeCheckbox.disabled = true;
            if (turboSection) turboSection.classList.add('disabled');

            // Disable capture speed section
            captureSpeedSlider.disabled = true;
            captureSpeedSection.classList.add('disabled');
            turboOverrideHint.textContent = 'Long Chat Mode overrides this setting';
            turboOverrideHint.style.display = 'block';

            // Disable cover page toggle
            if (settingsIncludeCover) settingsIncludeCover.disabled = true;
            if (coverSection) coverSection.classList.add('disabled');
        } else {
            // Re-enable turbo mode section
            turboModeCheckbox.disabled = false;
            if (turboSection) turboSection.classList.remove('disabled');

            // Re-enable cover page toggle
            if (settingsIncludeCover) settingsIncludeCover.disabled = false;
            if (coverSection) coverSection.classList.remove('disabled');

            // Restore turbo mode UI state
            turboOverrideHint.textContent = 'Turbo Mode overrides this setting';
            updateTurboModeUI(storedTurboMode);
        }
    }

    // Update UI when file attachments setting changes
    function updateFileAttachmentsUI(isEnabled) {
        if (!fileSettingsDetails) return;
        if (isEnabled) {
            fileSettingsDetails.style.display = 'block';
        } else {
            fileSettingsDetails.style.display = 'none';
        }
    }

    // Update max file size hint text
    function updateMaxFileSizeHint(sizeMB) {
        if (maxFileSizeHint) {
            maxFileSizeHint.textContent = `Up to ${sizeMB} MB per file`;
        }
    }

    function saveSettings() {
        writeStorage('local', {
            [SETTINGS_CAPTURE_SPEED_KEY]: storedCaptureSpeed,
            [SETTINGS_TURBO_MODE_KEY]: FORCE_VIEWER_EXPORT_ONLY ? true : storedTurboMode,
            [SETTINGS_LONG_CHAT_MODE_KEY]: storedLongChatMode,
            [SETTINGS_FILE_ATTACHMENTS_KEY]: storedFileAttachments,
            [SETTINGS_MAX_FILE_SIZE_KEY]: storedMaxFileSize
        });
    }

    // Settings event listeners
    addSafeListener(settingsBtn, 'click', openSettings, 'settings button');
    addSafeListener(closeSettingsBtn, 'click', closeSettings, 'close settings button');

    addSafeListener(captureSpeedSlider, 'input', (e) => {
        if (FORCE_VIEWER_EXPORT_ONLY) {
            captureSpeedSlider.value = storedCaptureSpeed;
            return;
        }
        storedCaptureSpeed = parseInt(e.target.value);
        updateSpeedLabel(storedCaptureSpeed);
        saveSettings();
    }, 'capture speed slider');

    addSafeListener(turboModeCheckbox, 'change', (e) => {
        if (FORCE_VIEWER_EXPORT_ONLY) {
            storedTurboMode = true;
            turboModeCheckbox.checked = true;
            return;
        }
        storedTurboMode = e.target.checked;
        updateTurboModeUI(storedTurboMode);
        saveSettings();
    }, 'turbo mode checkbox');

    addSafeListener(longChatModeCheckbox, 'change', (e) => {
        storedLongChatMode = e.target.checked;
        updateLongChatModeUI(storedLongChatMode);
        saveSettings();
    }, 'long chat mode checkbox');

    addSafeListener(fileAttachmentsCheckbox, 'change', (e) => {
        storedFileAttachments = e.target.checked;
        updateFileAttachmentsUI(storedFileAttachments);
        saveSettings();
    }, 'file attachments checkbox');

    addSafeListener(maxFileSizeSlider, 'input', (e) => {
        storedMaxFileSize = parseInt(e.target.value);
        updateMaxFileSizeHint(storedMaxFileSize);
        saveSettings();
    }, 'max file size slider');

    // Sync settings cover checkbox with the hidden original checkbox
    addSafeListener(settingsIncludeCover, 'change', (e) => {
        if (FORCE_VIEWER_EXPORT_ONLY) {
            includeCoverCheckbox.checked = false;
            settingsIncludeCover.checked = false;
            return;
        }
        includeCoverCheckbox.checked = e.target.checked;
        storedProIncludeCover = e.target.checked;
        writeStorage('local', { [PRO_INCLUDE_COVER_KEY]: storedProIncludeCover });
        writeStorage('sync', { [PRO_INCLUDE_COVER_KEY]: storedProIncludeCover });
    }, 'settings include cover checkbox');

    // Load settings on startup
    loadSettings();

    // ==================== End Settings Functions ====================

    function sendRuntimeMessage(message) {
        return new Promise((resolve, reject) => {
            chrome.runtime.sendMessage(message, (response) => {
                if (chrome.runtime.lastError) {
                    reject(new Error(chrome.runtime.lastError.message));
                    return;
                }
                resolve(response);
            });
        });
    }

    function requestOptionalAccess(request) {
        return new Promise((resolve, reject) => {
            chrome.permissions.request(request, (granted) => {
                if (chrome.runtime.lastError) {
                    reject(new Error(chrome.runtime.lastError.message));
                    return;
                }
                resolve(granted === true);
            });
        });
    }

    function injectTeamsContentScript(tabId) {
        return new Promise((resolve, reject) => {
            chrome.scripting.executeScript({
                target: { tabId },
                files: TEAMS_CONTENT_SCRIPT_FILES
            }, () => {
                if (chrome.runtime.lastError) {
                    reject(new Error(chrome.runtime.lastError.message));
                    return;
                }
                resolve();
            });
        });
    }

    function delay(ms) {
        return new Promise((resolve) => setTimeout(resolve, ms));
    }

    function getViewerUrl(chatId = null) {
        const baseUrl = chrome.runtime.getURL('viewer.html');
        if (!chatId) return baseUrl;
        const url = new URL(baseUrl);
        url.searchParams.set('chat', chatId);
        return url.toString();
    }

    function sendTabMessage(tabId, message) {
        return new Promise((resolve, reject) => {
            chrome.tabs.sendMessage(tabId, message, (response) => {
                if (chrome.runtime.lastError) {
                    reject(new Error(chrome.runtime.lastError.message));
                    return;
                }
                resolve(response);
            });
        });
    }

    function queryTabs(queryInfo) {
        return new Promise((resolve, reject) => {
            chrome.tabs.query(queryInfo, (tabs) => {
                if (chrome.runtime.lastError) {
                    reject(new Error(chrome.runtime.lastError.message));
                    return;
                }
                resolve(tabs || []);
            });
        });
    }

    function getTab(tabId) {
        return new Promise((resolve, reject) => {
            chrome.tabs.get(tabId, (tab) => {
                if (chrome.runtime.lastError) {
                    reject(new Error(chrome.runtime.lastError.message));
                    return;
                }
                resolve(tab || null);
            });
        });
    }

    function updateTab(tabId, updateProperties) {
        return new Promise((resolve, reject) => {
            chrome.tabs.update(tabId, updateProperties, (tab) => {
                if (chrome.runtime.lastError) {
                    reject(new Error(chrome.runtime.lastError.message));
                    return;
                }
                resolve(tab || null);
            });
        });
    }

    function focusWindow(windowId) {
        return new Promise((resolve, reject) => {
            if (typeof windowId !== 'number') {
                resolve();
                return;
            }
            chrome.windows.update(windowId, { focused: true }, () => {
                if (chrome.runtime.lastError) {
                    reject(new Error(chrome.runtime.lastError.message));
                    return;
                }
                resolve();
            });
        });
    }

    function reloadTab(tabId) {
        return new Promise((resolve, reject) => {
            chrome.tabs.reload(tabId, {}, () => {
                if (chrome.runtime.lastError) {
                    reject(new Error(chrome.runtime.lastError.message));
                    return;
                }
                resolve();
            });
        });
    }

    function createTab(createProperties) {
        return new Promise((resolve, reject) => {
            chrome.tabs.create(createProperties, (tab) => {
                if (chrome.runtime.lastError) {
                    reject(new Error(chrome.runtime.lastError.message));
                    return;
                }
                resolve(tab || null);
            });
        });
    }

    function removeTabs(tabIds) {
        const ids = (Array.isArray(tabIds) ? tabIds : [tabIds])
            .filter((id) => typeof id === 'number');

        if (ids.length === 0) {
            return Promise.resolve();
        }

        return new Promise((resolve, reject) => {
            chrome.tabs.remove(ids, () => {
                if (chrome.runtime.lastError) {
                    reject(new Error(chrome.runtime.lastError.message));
                    return;
                }
                resolve();
            });
        });
    }

    function getLocalStorageValue(key) {
        return new Promise((resolve) => {
            chrome.storage.local.get([key], (result) => {
                if (chrome.runtime.lastError) {
                    console.warn(`Failed to read ${key} from storage:`, chrome.runtime.lastError.message);
                    resolve(null);
                    return;
                }
                resolve(result?.[key] ?? null);
            });
        });
    }

    function isTeamsUrl(url) {
        try {
            const hostname = new URL(url).hostname.toLowerCase();
            const matchesHost = (host) => hostname === host || hostname.endsWith(`.${host}`);
            return matchesHost('teams.microsoft.com') ||
                matchesHost('teams.live.com') ||
                matchesHost('teams.cloud.microsoft') ||
                hostname.endsWith('.mcas.ms');
        } catch (error) {
            return false;
        }
    }

    function isRecoverableContentScriptError(message = '') {
        return CLOSED_ASYNC_RESPONSE_ERROR_RE.test(message) ||
            CLOSED_MESSAGE_PORT_ERROR_RE.test(message) ||
            NO_RECEIVER_RUNTIME_ERROR_RE.test(message);
    }

    function isViewerTab(tab) {
        return !!(tab && typeof tab.url === 'string' && tab.url.startsWith(chrome.runtime.getURL('viewer.html')));
    }

    async function getAllViewerTabs() {
        const allTabs = await queryTabs({}).catch(() => []);
        return allTabs.filter(isViewerTab);
    }

    async function findExistingViewerTab() {
        const storedViewerTabId = await getLocalStorageValue(VIEWER_TAB_STORAGE_KEY);
        const viewerTabs = await getAllViewerTabs();

        if (viewerTabs.length === 0) {
            return { primary: null, duplicates: [] };
        }

        const viewerTabsById = new Map(viewerTabs.map((tab) => [tab.id, tab]));

        if (typeof storedViewerTabId === 'number') {
            const storedTab = viewerTabsById.get(storedViewerTabId) || await getTab(storedViewerTabId).catch(() => null);
            if (isViewerTab(storedTab)) {
                return {
                    primary: storedTab,
                    duplicates: viewerTabs.filter((tab) => tab.id !== storedTab.id)
                };
            }
        }

        viewerTabs.sort((a, b) => {
            if ((b.active ? 1 : 0) !== (a.active ? 1 : 0)) {
                return (b.active ? 1 : 0) - (a.active ? 1 : 0);
            }
            return (b.id || 0) - (a.id || 0);
        });

        return {
            primary: viewerTabs[0],
            duplicates: viewerTabs.slice(1)
        };
    }

    async function openOrFocusViewer({ chatId = null, forceRoot = false, refresh = true } = {}) {
        const targetUrl = chatId ? getViewerUrl(chatId) : getViewerUrl();
        const viewerState = await findExistingViewerTab().catch((error) => {
            console.warn('Could not locate existing viewer tab:', error);
            return { primary: null, duplicates: [] };
        });
        const existingViewerTab = viewerState?.primary || null;
        const duplicateViewerTabs = viewerState?.duplicates || [];

        if (!existingViewerTab) {
            const createdTab = await createTab({ url: targetUrl });
            if (createdTab?.id) {
                writeStorage('local', { [VIEWER_TAB_STORAGE_KEY]: createdTab.id });
            }
            return;
        }

        if (duplicateViewerTabs.length > 0) {
            removeTabs(duplicateViewerTabs.map((tab) => tab.id)).catch((error) => {
                console.warn('Could not remove duplicate viewer tabs:', error);
            });
        }

        await focusWindow(existingViewerTab.windowId).catch((error) => {
            console.warn('Could not focus viewer window:', error);
        });

        const currentUrl = existingViewerTab.url || '';
        const nextUrl = forceRoot
            ? targetUrl
            : (chatId ? targetUrl : currentUrl || targetUrl);
        const shouldNavigate = !!nextUrl && currentUrl !== nextUrl;

        const updatedTab = await updateTab(existingViewerTab.id, shouldNavigate ? { active: true, url: nextUrl } : { active: true });
        if (updatedTab?.id) {
            writeStorage('local', { [VIEWER_TAB_STORAGE_KEY]: updatedTab.id });
        }

        if (!shouldNavigate && refresh) {
            await reloadTab(existingViewerTab.id).catch((error) => {
                console.warn('Could not refresh viewer tab:', error);
            });
        }
    }

    async function waitForTeamsContentScript(tabId, timeoutMs = BACKUP_RUNTIME_RECONNECT_TIMEOUT_MS) {
        const startTime = Date.now();
        let lastError = null;
        let injectionAttempted = false;

        while (Date.now() - startTime < timeoutMs) {
            const tab = await getTab(tabId).catch((error) => {
                lastError = error;
                return null;
            });

            if (!tab) {
                throw new Error(lastError?.message || 'Teams tab is no longer available.');
            }

            if (!isTeamsUrl(tab.url || '')) {
                throw new Error('Teams tab navigated away during backup.');
            }

            if (tab.status === 'complete') {
                try {
                    const pingResponse = await sendTabMessage(tabId, { action: "ping" });
                    if (pingResponse && pingResponse.status === "ready") {
                        return pingResponse;
                    }
                } catch (error) {
                    lastError = error;
                    if (!isRecoverableContentScriptError(error?.message || '')) {
                        throw error;
                    }
                    // The content script is injected on demand, so a reloaded Teams page
                    // has none until it is added back.
                    if (!injectionAttempted && NO_RECEIVER_RUNTIME_ERROR_RE.test(error?.message || '')) {
                        injectionAttempted = true;
                        await injectTeamsContentScript(tabId).catch((injectionError) => {
                            lastError = injectionError;
                        });
                    }
                }
            }

            await delay(BACKUP_RUNTIME_RECONNECT_POLL_MS);
        }

        throw new Error(lastError?.message || 'Timed out waiting for Teams to reconnect.');
    }

    async function recoverBackupMessageChannel(tabId, { statusPrefix = '', chatIndex, chatTarget, totalChats = 1, currentIndex = 0 } = {}) {
        statusDiv.textContent = statusPrefix + 'Teams reloaded. Reconnecting...';
        await delay(BACKUP_RUNTIME_RECONNECT_DELAY_MS);
        await waitForTeamsContentScript(tabId);

        if (chatIndex !== undefined) {
            statusDiv.textContent = totalChats > 1
                ? `Reopening chat ${currentIndex + 1} of ${totalChats}...`
                : 'Reopening chat...';

            const openResponse = await sendTabMessage(tabId, {
                action: "openChat",
                chatIndex,
                chatTarget: chatTarget || null
            });

            if (!openResponse || openResponse.status !== "success") {
                throw new Error('Teams reconnected, but the selected chat could not be reopened.');
            }

            await delay(750);
        } else {
            statusDiv.textContent = statusPrefix + 'Reconnected. Retrying backup...';
            await delay(500);
        }
    }

    async function sendBackupChatMessageWithRecovery(tabId, message, recoveryContext) {
        try {
            return await sendTabMessage(tabId, message);
        } catch (error) {
            if (!isRecoverableContentScriptError(error?.message || '')) {
                throw error;
            }

            try {
                await recoverBackupMessageChannel(tabId, recoveryContext);
            } catch (recoveryError) {
                throw Object.assign(error, {
                    recoveryAttempted: true,
                    recoveryError: recoveryError?.message || String(recoveryError),
                    retryable: true
                });
            }

            try {
                return await sendTabMessage(tabId, message);
            } catch (retryError) {
                throw Object.assign(retryError, {
                    recoveryAttempted: true,
                    retryable: isRecoverableContentScriptError(retryError?.message || '')
                });
            }
        }
    }

    function isBackupAlreadyInProgressResponse(response) {
        return response?.status === 'failed' &&
            /already in progress/i.test(response?.message || '');
    }

    async function getContentBackupStatus(tabId) {
        try {
            return await sendTabMessage(tabId, { action: "getBackupStatus" });
        } catch (error) {
            if (!isRecoverableContentScriptError(error?.message || '')) {
                console.warn('Could not read backup status from content script:', error?.message || error);
            }
            return null;
        }
    }

    async function waitForContentBackupToStop(tabId, timeoutMs = BACKUP_STALE_RECOVERY_TIMEOUT_MS) {
        const startTime = Date.now();

        while (Date.now() - startTime < timeoutMs) {
            const status = await getContentBackupStatus(tabId);
            if (status && (!status.inProgress || status.status === 'idle')) {
                return true;
            }
            await delay(BACKUP_STALE_RECOVERY_POLL_MS);
        }

        return false;
    }

    async function recoverBusyBackupLock(tabId, backupRequest, { statusPrefix = '', chatIndex, chatTarget, totalChats = 1, currentIndex = 0 } = {}) {
        statusDiv.textContent = statusPrefix + 'Clearing previous backup...';

        try {
            await sendTabMessage(tabId, { action: "cancelBackup" });
        } catch (error) {
            if (!isRecoverableContentScriptError(error?.message || '')) {
                console.warn('Cancel backup request failed:', error?.message || error);
            }
        }

        const lockReleased = await waitForContentBackupToStop(tabId);
        if (!lockReleased) {
            return null;
        }

        if (chatIndex !== undefined) {
            statusDiv.textContent = totalChats > 1
                ? `Reopening chat ${currentIndex + 1} of ${totalChats}...`
                : 'Reopening chat...';

            const openResponse = await sendTabMessage(tabId, {
                action: "openChat",
                chatIndex,
                chatTarget: chatTarget || null
            });

            if (!openResponse || openResponse.status !== "success") {
                throw new Error('Previous backup stopped, but the selected chat could not be reopened.');
            }

            await delay(750);
        }

        statusDiv.textContent = statusPrefix + 'Retrying backup...';
        return sendBackupChatMessageWithRecovery(tabId, backupRequest, {
            statusPrefix,
            chatIndex,
            chatTarget,
            totalChats,
            currentIndex
        });
    }

    function renderAppUserId(userId) {
        [appUserIdValue, settingsLicenseIdValue].forEach((element) => {
            if (!element) return;
            element.textContent = userId || 'Unavailable';
            element.title = userId || '';
        });
        if (restoreSupportLink) {
            restoreSupportLink.href = window.TeamsBackupLicense.buildSupportMailto(userId);
        }
        if (settingsSupportLink) {
            settingsSupportLink.href = window.TeamsBackupLicense.buildSupportMailto(userId, 'Help with my Pro purchase');
        }
    }

    // Plan state is applied once stored Pro status is read, so Pro users never
    // see the upgrade section flash on open.
    checkEntitlements({ celebrateOverride: false }).catch((error) => {
        console.error('Payment entitlement check error:', error);
        applyFreeState();
    });

    // Get or create unique user ID
    async function getOrCreateUserId() {
        if (cachedUserId) {
            renderAppUserId(cachedUserId);
            return cachedUserId;
        }

        const { userId } = await window.TeamsBackupLicense.getOrCreateLicenseId();
        cachedUserId = userId;
        renderAppUserId(userId);

        // A synced License ID is available for an explicit Restore Purchase.
        return userId;
    }

    // Check user entitlements
    async function checkEntitlements({ celebrateOverride = null } = {}) {
        try {
            // Test mode override
            if (ENABLE_TEST_MODE) {
                const celebrateTest = celebrateOverride !== null ? celebrateOverride : !proCelebrationAcknowledged;
                applyProState({ celebrate: celebrateTest });
                console.log('TEST MODE: Pro features enabled');
                return true;
            }

            // Check storage first for quick verification
            const storageData = await readStorage('local', ['isPro', 'purchaseVerified', PRO_CELEBRATION_KEY, PRO_INCLUDE_COVER_KEY]);

            if (typeof storageData[PRO_CELEBRATION_KEY] === 'boolean') {
                proCelebrationAcknowledged = storageData[PRO_CELEBRATION_KEY];
            }

            if (typeof storageData[PRO_INCLUDE_COVER_KEY] === 'boolean') {
                storedProIncludeCover = storageData[PRO_INCLUDE_COVER_KEY];
            }

            if (storageData.isPro && storageData.purchaseVerified) {
                const shouldCelebrate = celebrateOverride !== null ? celebrateOverride : !proCelebrationAcknowledged;
                applyProState({ celebrate: shouldCelebrate });
                console.log('Pro status verified from storage');
                if (!shouldCelebrate) {
                    statusDiv.style.display = 'none';
                }
                return true;
            }

            applyFreeState();
            return false;
        } catch (error) {
            console.error('Error checking entitlements:', error);
            applyFreeState();
            return false;
        }

        return isPro;
    }

    // The background worker watches the checkout and stores Pro when payment
    // lands; reflect it here if the popup is open at that moment.
    chrome.storage.onChanged.addListener((changes, areaName) => {
        if (areaName !== 'local') return;
        if (changes.isPro?.newValue === true && !isPro) {
            purchasePollingMessage.style.display = 'none';
            applyProState({ celebrate: !proCelebrationAcknowledged });
        }
    });

    async function openCheckout() {
        const result = await window.TeamsBackupLicense.openCheckout();
        if (result.alreadyPro) {
            applyProState({ celebrate: false });
            return;
        }
        // Only reached if the popup stays open while the checkout tab loads.
        statusDiv.textContent = 'Secure checkout opened in a new tab. Pro unlocks automatically after payment.';
        statusDiv.style.color = '#2196f3';
        statusDiv.style.display = 'block';
        purchasePollingMessage.textContent = '⏳ Waiting for payment confirmation...';
        purchasePollingMessage.style.display = 'block';
    }

    function createUpgradeButton(label = 'Upgrade to Pro') {
        const button = document.createElement('button');
        button.type = 'button';
        button.className = 'inline-upgrade-button';
        button.textContent = label;
        button.addEventListener('click', () => {
            openCheckout().catch((error) => {
                console.error('Purchase error:', error);
                showToast('Could not open checkout. Please try again.', 4000, 'warning');
            });
        });
        return button;
    }

    // Handle purchase. The RevenueCat API allows CORS from this extension, so no
    // extra permission prompt is needed before opening the hosted checkout.
    addSafeListener(purchaseButton, 'click', async () => {
        try {
            purchaseButton.disabled = true;
            purchaseButton.textContent = 'Opening checkout...';
            await openCheckout();
        } catch (error) {
            console.error('Purchase error:', error);
            statusDiv.textContent = 'Unable to open checkout. Please try again.';
            statusDiv.style.color = '#dc3545';
            statusDiv.style.display = 'block';
        } finally {
            purchaseButton.disabled = false;
            purchaseButton.textContent = 'Unlock Full Version';
        }
    }, 'purchase button');

    addSafeListener(restoreButton, 'click', async () => {
        // First try to restore with current user ID (Chrome sync)
        try {
            restoreButton.disabled = true;
            restoreButton.textContent = 'Checking...';

            statusDiv.textContent = 'Checking for existing purchases...';
            statusDiv.style.color = '#2196f3';
            statusDiv.style.display = 'block';

            const currentUserId = await getOrCreateUserId();
            const response = await sendRuntimeMessage({
                action: 'restorePurchaseLookup',
                identifier: currentUserId
            });
            if (response?.status === 'error') {
                throw new Error(response.message || 'Could not check RevenueCat');
            }

            const restored = response?.status === 'success' && response.isPro === true;

            if (restored) {
                applyProState({ celebrate: true });
                statusDiv.textContent = 'Pro access restored on this device.';
                statusDiv.style.color = '#4caf50';
                statusDiv.style.display = 'block';
                chrome.storage.local.remove(['checkEntitlementsOnOpen', 'checkEntitlementsTimestamp']);
                restoreButton.disabled = false;
                restoreButton.textContent = 'Restore Purchase';
                return;
            }

            // No purchase found with current ID - show email modal
            restoreButton.disabled = false;
            restoreButton.textContent = 'Restore Purchase';
            statusDiv.style.display = 'none';
            showEmailModal();

        } catch (error) {
            console.error('Restore purchase error:', error);
            // On error, also show email modal as fallback
            restoreButton.disabled = false;
            restoreButton.textContent = 'Restore Purchase';
            if (statusDiv) {
                statusDiv.textContent = `Restore check failed: ${error.message || 'Unknown error'}`;
                statusDiv.style.color = '#dc3545';
                statusDiv.style.display = 'block';
            }
            showEmailModal();
        }
    }, 'restore button');

    // Backup confirmation modal elements
    const backupConfirmModal = document.getElementById('backupConfirmModal');
    const confirmContinueBtn = document.getElementById('confirmContinueBtn');
    const confirmCancelBtn = document.getElementById('confirmCancelBtn');

    // Track cancellation reason for user feedback
    let cancellationReason = null;
    const CANCEL_REASONS = {
        USER_STOP: 'user_stop',
        POPUP_CLOSED: 'popup_closed'
    };

    // Track if user unfocused during backup
    let unfocusedDuringBackup = false;
    let unfocusWarningShown = false;

    // Visibility change detection - warn user when they navigate away during backup
    function handleVisibilityChange() {
        if (backupInProgress && document.hidden) {
            unfocusedDuringBackup = true;
        } else if (backupInProgress && !document.hidden && unfocusedDuringBackup && !unfocusWarningShown) {
            // User came back while backup is still running - show warning
            showUnfocusWarning();
            unfocusWarningShown = true;
        }
    }

    // Window blur detection (when popup loses focus but stays open)
    function handleWindowBlur() {
        if (backupInProgress) {
            unfocusedDuringBackup = true;
        }
    }

    function handleWindowFocus() {
        if (backupInProgress && unfocusedDuringBackup && !unfocusWarningShown) {
            showUnfocusWarning();
            unfocusWarningShown = true;
        }
    }

    // Show unfocus warning toast
    function showUnfocusWarning() {
        showToast("Keep the Teams tab open. Saved checkpoints will appear in the Viewer.", 5000, 'warning');
    }

    // Update the backup hint text based on Pro status during backup
    function updateBackupHintForBackup() {
        if (backupHint) {
            if (isPro) {
                backupHint.textContent = "Keep Teams open. Checkpoints are saved locally as messages are found.";
                backupHint.style.color = "#2196f3";
            } else {
                backupHint.textContent = "Keep Teams open. Checkpoints are saved locally as messages are found.";
                backupHint.style.color = "#ff9800";
            }
            backupHint.style.fontWeight = "bold";
        }
    }

    // Add event listeners for visibility/focus
    document.addEventListener('visibilitychange', handleVisibilityChange);
    window.addEventListener('blur', handleWindowBlur);
    window.addEventListener('focus', handleWindowFocus);

    // Cleanup handler when popup closes
    window.addEventListener('pagehide', () => {
        // If backup was in progress, tell content script to clean up Blob URLs
        if (backupInProgress) {
            chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
                if (tabs.length > 0) {
                    chrome.tabs.sendMessage(tabs[0].id, { action: "cleanupBlobUrls" });
                }
            });
        }
    });

    // Email modal functions
    function showEmailModal() {
        if (!emailModal || !emailInput || !emailError) return;
        emailModal.style.display = 'flex';
        emailInput.value = '';
        emailError.style.display = 'none';
        emailInput.focus();
    }

    function hideEmailModal() {
        if (!emailModal || !emailInput || !emailError) return;
        emailModal.style.display = 'none';
        emailInput.value = '';
        emailError.style.display = 'none';
    }

    // Backup confirmation modal functions
    function showBackupConfirmModal() {
        if (backupConfirmModal) {
            backupConfirmModal.style.display = 'flex';
        }
    }

    function hideBackupConfirmModal() {
        if (backupConfirmModal) {
            backupConfirmModal.style.display = 'none';
        }
    }

    // Handle "Continue Backup" button
    addSafeListener(confirmContinueBtn, 'click', () => {
        hideBackupConfirmModal();
    }, 'confirm continue backup button');

    // Handle "Open Viewer" button
    addSafeListener(confirmCancelBtn, 'click', () => {
        hideBackupConfirmModal();
        openViewerDuringBackup();
    }, 'confirm cancel backup button');

    // Open the viewer while the Teams content script keeps checkpointing in the background.
    function openViewerDuringBackup() {
        openOrFocusViewer({ forceRoot: true, refresh: true }).catch((error) => {
            console.error('Failed to open viewer:', error);
        });

        statusDiv.textContent = "Viewer opened. Keep the Teams tab open so backup checkpoints can continue saving.";
        statusDiv.style.color = "#2196f3";
        statusDiv.style.display = "block";
    }

    // Handle clicking backdrop to close modal
    const backupModalBackdrop = backupConfirmModal ? backupConfirmModal.querySelector('.confirm-modal__backdrop') : null;
    addSafeListener(backupModalBackdrop, 'click', hideBackupConfirmModal, 'backup confirm modal backdrop');

    addSafeListener(emailCancelBtn, 'click', hideEmailModal, 'email cancel button');

    const emailModalBackdrop = emailModal ? emailModal.querySelector('.email-modal__backdrop') : null;
    addSafeListener(emailModalBackdrop, 'click', hideEmailModal, 'email modal backdrop');

    async function copyLicenseId() {
        try {
            const userId = await getOrCreateUserId();
            if (!userId) {
                throw new Error('No License ID available');
            }

            await navigator.clipboard.writeText(userId);
            if (statusDiv) {
                statusDiv.textContent = 'License ID copied.';
                statusDiv.style.color = '#4caf50';
                statusDiv.style.display = 'block';
            }
        } catch (error) {
            console.error('Copy License ID failed:', error);
            if (statusDiv) {
                statusDiv.textContent = 'Could not copy License ID.';
                statusDiv.style.color = '#dc3545';
                statusDiv.style.display = 'block';
            }
        }
    }

    addSafeListener(copyUserIdButton, 'click', copyLicenseId, 'copy app user id button');
    addSafeListener(settingsCopyLicenseButton, 'click', async () => {
        await copyLicenseId();
        showToast('License ID copied.', 2500, 'success');
    }, 'settings copy license button');

    addSafeListener(emailSubmitBtn, 'click', async () => {
        const restoreIdentifier = emailInput.value.trim();

        if (!restoreIdentifier) {
            emailError.textContent = 'Enter a License ID.';
            emailError.style.display = 'block';
            return;
        }

        if (restoreIdentifier.includes('@')) {
            emailError.textContent = 'Enter the License ID from the original device (it starts with "user_"). Lost it? Use "Email support" below with your purchase receipt.';
            emailError.style.display = 'block';
            return;
        }

        emailSubmitBtn.disabled = true;
        emailSubmitBtn.textContent = 'Checking...';
        emailError.style.display = 'none';

        try {
            const currentUserId = await getOrCreateUserId();
            const response = await sendRuntimeMessage({
                action: 'restorePurchaseLookup',
                identifier: restoreIdentifier,
                currentUserId
            });

            if (response && response.status === 'success' && response.isPro) {
                // Found purchase. Use the verified License ID returned by the payment server.
                if (response.userId) {
                    cachedUserId = response.userId;
                    await Promise.all([
                        writeStorage('local', { userId: response.userId }),
                        writeStorage('sync', { userId: response.userId })
                    ]);
                    renderAppUserId(response.userId);
                }

                hideEmailModal();
                applyProState({ celebrate: true });
                chrome.storage.local.set({ isPro: true, purchaseVerified: true });
                statusDiv.textContent = 'Pro access restored successfully!';
                statusDiv.style.color = '#4caf50';
                statusDiv.style.display = 'block';
            } else if (response && response.status === 'not_found') {
                emailError.textContent = response.message || 'No matching purchase was found.';
                emailError.style.display = 'block';
            } else {
                emailError.textContent = response?.message || 'Unable to verify. Please try again.';
                emailError.style.display = 'block';
            }
        } catch (error) {
            console.error('Email restore error:', error);
            emailError.textContent = `Error checking purchase: ${error.message || 'Unknown error'}`;
            emailError.style.display = 'block';
        } finally {
            emailSubmitBtn.disabled = false;
            emailSubmitBtn.textContent = 'Restore';
        }
    }, 'email submit button');

    // Allow Enter key to submit email
    addSafeListener(emailInput, 'keypress', (e) => {
        if (e.key === 'Enter') {
            emailSubmitBtn.click();
        }
    }, 'email input');

    // Make the current License ID visible for manual cross-device restore.
    getOrCreateUserId().catch((error) => {
        console.error('Failed to initialize License ID:', error);
        renderAppUserId(null);
    });

    // Add stop button visibility control with proper styling
    function toggleButtons(isBackupInProgress) {
        backupButton.style.display = isBackupInProgress ? 'none' : 'block';
        stopButton.style.display = isBackupInProgress ? 'block' : 'none';
    }

    // Function to check if current tab is Teams (Business or Personal/Consumer)
    async function checkIfTeamsTab() {
        return new Promise((resolve) => {
            chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
                if (tabs.length === 0) {
                    resolve(false);
                    return;
                }
                const url = tabs[0].url;
                resolve(isTeamsUrl(url));
            });
        });
    }

    function getAvailableChatsWithTimeout(tabId) {
        return Promise.race([
            sendTabMessage(tabId, { action: "getAvailableChats" }),
            delay(LOAD_AVAILABLE_CHATS_TIMEOUT_MS).then(() => {
                const error = new Error('getAvailableChats timed out');
                error.code = 'timeout';
                throw error;
            })
        ]);
    }

    // Inject only after the user clicks the extension. This keeps Teams access
    // temporary and avoids persistent website permissions at install/update time.
    async function loadAvailableChats({ allowInjection = true } = {}) {
        const tabs = await queryTabs({ active: true, currentWindow: true }).catch(() => []);
        const activeTab = tabs[0] || null;

        if (!activeTab?.id || !isTeamsUrl(activeTab.url || '')) {
            return { chats: [], error: null };
        }

        try {
            const response = await getAvailableChatsWithTimeout(activeTab.id);
            if (!response || response.status !== "success") {
                return {
                    chats: response?.chats || [],
                    error: response?.error || 'request_failed',
                    diagnostics: response?.diagnostics || null
                };
            }
            return {
                chats: response.chats || [],
                error: null,
                diagnostics: response.diagnostics || null
            };
        } catch (error) {
            if (allowInjection && NO_RECEIVER_RUNTIME_ERROR_RE.test(error.message || '')) {
                try {
                    await injectTeamsContentScript(activeTab.id);
                    await delay(100);
                    return loadAvailableChats({ allowInjection: false });
                } catch (injectionError) {
                    console.error('Could not enable Teams access for this tab:', injectionError);
                    return {
                        chats: [],
                        error: 'no_connection',
                        diagnostics: { injectionError: injectionError.message || String(injectionError) }
                    };
                }
            }

            if (error.code === 'timeout') {
                console.warn(error.message);
                return { chats: [], error: 'timeout', diagnostics: { requestTimedOut: true } };
            }

            console.log('Content script not responding:', error.message || error);
            return { chats: [], error: 'no_connection' };
        }
    }

    function isChatSupportedForBackup(chat) {
        if (!chat) return false;
        if (typeof chat.isSupported === 'boolean') return chat.isSupported;
        // Fallback for older content script payloads.
        return chat.isChannel !== true;
    }

    function getNoChatsDetectedHint() {
        if (chatLoadDiagnostics?.preservedCurrentConversation && chatLoadDiagnostics?.hasChatRailButton && !chatLoadDiagnostics?.chatRailSelected) {
            return 'Staying on the current Teams/channel view. Use Backup Open Chat below, or open Chat in Teams if you want to browse the sidebar list.';
        }
        if (chatLoadDiagnostics?.requestTimedOut) {
            return 'Teams took too long to return the sidebar. Try clicking Backup Chats & Channels again, or use Backup Open Chat below.';
        }
        if (chatLoadDiagnostics?.requestFailed) {
            return 'The Teams sidebar could not be read. Try reopening Chat or use Backup Open Chat below.';
        }
        if (chatLoadDiagnostics?.chatSectionCollapsed || chatLoadDiagnostics?.teamsSectionCollapsed) {
            return 'Expand the Chats or Teams and channels section in Teams, then open the backup list again.';
        }
        if (chatLoadDiagnostics?.hasChatRailButton && !chatLoadDiagnostics?.chatRailSelected) {
            return 'Open the Chats or Teams app in Teams, then open the backup list again.';
        }
        if (chatLoadDiagnostics?.recoveryAttempted && !chatLoadDiagnostics?.recoverySucceeded) {
            return 'Teams looks collapsed. Open Chat and expand the sidebar sections, or use Backup Open Chat below.';
        }
        return 'Open Chats or Teams in Microsoft Teams, or uncollapse the sidebar, then click below to backup the open conversation.';
    }

    function getChatKindLabel(chat) {
        const kind = chat?.kind || (chat?.isChannel ? 'channel' : 'chat');
        switch (kind) {
            case 'channel_chat':
                return 'Channel chat';
            case 'channel_posts':
                return 'Channel posts';
            case 'channel':
                return 'Channel';
            case 'chat':
            default:
                return 'Chat';
        }
    }

    function getChatKindClass(chat) {
        const kind = chat?.kind || (chat?.isChannel ? 'channel' : 'chat');
        return `chat-selection__kind--${kind}`;
    }

    function getChatItemHint(chat) {
        if (!chat) return '';
        if (chat.kind === 'channel_posts') {
            return isChatSupportedForBackup(chat)
                ? 'Post-style channel view.'
                : 'Posts layout detected. Replies are expanded automatically during backup.';
        }
        if (chat.kind === 'channel_chat') {
            return 'Conversation-style channel view.';
        }
        if (chat.kind === 'channel') {
            return 'Channel item';
        }
        return '';
    }

    // One plan-limit message with a direct upgrade button.
    function renderPlanLimitMessage(message) {
        statusDiv.textContent = '';
        statusDiv.appendChild(document.createTextNode(`${message} `));
        statusDiv.appendChild(createUpgradeButton());
        statusDiv.style.color = '#ff9800';
        statusDiv.style.display = 'block';
    }

    function renderFreeTierLimitMessage(limitResult = {}) {
        renderPlanLimitMessage(limitResult.message ||
            window.TeamsBackupDB.describeFreeLimit({
                reason: limitResult.reason || 'lifetime_limit',
                activeCount: limitResult.count,
                usedChatNames: limitResult.usedChatNames || []
            }));
    }

    const UPDATE_LOCKED_MESSAGE = 'This chat is already backed up. The free plan keeps the saved copy but cannot refresh it with newer messages. Upgrade to Pro to keep it up to date.';

    async function getOpenChatIdentity() {
        const tabs = await queryTabs({ active: true, currentWindow: true }).catch(() => []);
        if (!tabs[0]?.id) return null;
        const ping = await sendTabMessage(tabs[0].id, { action: 'ping' }).catch(() => null);
        return ping?.chatName ? { name: ping.chatName, rawName: ping.chatName } : null;
    }

    async function passesFreeLimitPrecheck() {
        if (isPro) return true;
        const db = window.TeamsBackupDB;
        if (!db || typeof db.getFreeBackupAllowance !== 'function') return true;

        for (const index of selectedChatIndices) {
            // "Backup Open Chat" has no sidebar entry; ask Teams which chat is open.
            const chat = index === -1 ? await getOpenChatIdentity() : availableChats[index];
            const existing = chat?.name ? await db.findChatByName(chat) : null;

            // Free users can finish an incomplete backup but cannot refresh a complete
            // one. Say so before scrolling instead of failing at the first save.
            if (existing && existing.isPreview !== true && existing.inProgress !== true) {
                renderPlanLimitMessage(UPDATE_LOCKED_MESSAGE);
                return false;
            }

            if (!existing) {
                const allowance = await db.getFreeBackupAllowance(chat);
                if (!allowance.allowed) {
                    renderPlanLimitMessage(db.describeFreeLimit(allowance));
                    return false;
                }
            }
        }

        return true;
    }

    // Function to render chat selection list
    async function renderChatList() {
        // Drop stale indices but keep best-effort (posts layout) channels selectable.
        selectedChatIndices = selectedChatIndices.filter(idx => idx === -1 || !!availableChats[idx]);

        if (availableChats.length === 0) {
            // No chats detected - show informative message with fallback option
            chatList.innerHTML = `
                <div class="chat-selection__empty-state">
                    <div class="chat-selection__empty-icon">💬</div>
                    <div class="chat-selection__empty-text">No chats or channels detected</div>
                    <div class="chat-selection__empty-hint">${escapeHtml(getNoChatsDetectedHint())}</div>
                </div>
                <div class="chat-selection__item selected" data-index="-1">
                    <input type="checkbox" class="chat-selection__checkbox" checked>
                    <div class="chat-selection__info">
                        <div class="chat-selection__name">Backup Open Chat</div>
                        <div class="chat-selection__preview">Backup whatever chat is currently visible</div>
                    </div>
                </div>
            `;
            selectedChatIndices = [-1]; // -1 means "current chat"

            // Add click handler for the fallback item
            const fallbackItem = chatList.querySelector('.chat-selection__item');
            if (fallbackItem) {
                fallbackItem.addEventListener('click', () => {
                    selectedChatIndices = [-1];
                    updateBackupButtonState();
                });
            }

            updateBackupButtonState();
            return;
        }

        // Show ALL chats - free users can see all but only select one at a time
        const chatItemsHtml = availableChats.map((chat, index) => `
            <div class="chat-selection__item ${selectedChatIndices.includes(index) ? 'selected' : ''} ${!isChatSupportedForBackup(chat) ? 'chat-selection__item--unsupported' : ''}" data-index="${index}">
                <input type="checkbox" class="chat-selection__checkbox" ${selectedChatIndices.includes(index) ? 'checked' : ''}>
                <div class="chat-selection__info">
                    <div class="chat-selection__name-row">
                        <div class="chat-selection__name">${escapeHtml(chat.name)}</div>
                        <span class="chat-selection__kind ${getChatKindClass(chat)}">${escapeHtml(getChatKindLabel(chat))}</span>
                    </div>
                    ${chat.preview ? `<div class="chat-selection__preview">${escapeHtml(chat.preview)}</div>` : ''}
                    ${getChatItemHint(chat) ? `<div class="chat-selection__preview chat-selection__preview--hint">${escapeHtml(getChatItemHint(chat))}</div>` : ''}
                    ${!isChatSupportedForBackup(chat) ? '<div class="chat-selection__preview chat-selection__preview--unsupported">This channel uses the Posts/Threads layout. Backup runs in best-effort mode.</div>' : ''}
                </div>
                ${!isChatSupportedForBackup(chat) ? '<span class="chat-selection__active chat-selection__active--unsupported">LIMITED</span>' : ''}
                ${chat.isActive ? '<span class="chat-selection__active">CURRENT</span>' : ''}
            </div>
        `).join('');

        chatList.innerHTML = chatItemsHtml;

        // Add click handlers
        chatList.querySelectorAll('.chat-selection__item').forEach(item => {
            item.addEventListener('click', (e) => {
                const index = parseInt(item.dataset.index);
                const chat = availableChats[index];

                if (!isChatSupportedForBackup(chat) && !selectedChatIndices.includes(index)) {
                    showToast('This channel uses the Posts/Threads layout. Backup will run in best-effort mode and may miss some replies.', 4500, 'warning');
                }

                if (selectedChatIndices.includes(index)) {
                    // Deselect this chat
                    selectedChatIndices = selectedChatIndices.filter(i => i !== index);
                } else {
                    // Select this chat
                    if (!isPro) {
                        // Free user can only select one - replace any existing selection
                        selectedChatIndices = [index];
                    } else {
                        // Pro user can select multiple
                        selectedChatIndices.push(index);
                    }
                }

                // Re-render to update selection states
                renderChatList();
                updateBackupButtonState();
            });
        });

        // Show notice - different text for free vs pro
        if (!isPro) {
            const planText = `Free plan: ${POPUP_FREE_CHAT_LIMIT} chat backup${POPUP_FREE_CHAT_LIMIT === 1 ? '' : 's'}`;
            try {
                const usage = await window.TeamsBackupDB.getFreeBackupUsage();
                const usedCount = Math.min(usage.count, POPUP_FREE_CHAT_LIMIT);
                freeUserNotice.textContent = `${planText} (${usedCount}/${POPUP_FREE_CHAT_LIMIT} used). Pro backs up every chat.`;
            } catch (e) {
                freeUserNotice.textContent = `${planText}. Pro backs up every chat.`;
            }
            freeUserNotice.style.display = 'block';
        } else {
            freeUserNotice.textContent = 'Pro: Select multiple chats or channels for full backups.';
            freeUserNotice.style.display = 'block';
        }
    }

    // Update backup button and start button state based on selection
    function updateBackupButtonState() {
        selectedChatIndices = selectedChatIndices.filter(idx => idx === -1 || !!availableChats[idx]);
        const count = selectedChatIndices.length;

        // Update the main button appearance
        if (count === 0) {
            backupButton.classList.add('no-selection');
            backupButton.classList.remove('has-selection');
            startBackupBtn.disabled = true;
            startBackupBtn.textContent = 'Select a chat or channel';
        } else if (count === 1) {
            backupButton.classList.remove('no-selection');
            backupButton.classList.add('has-selection');
            startBackupBtn.disabled = false;
            startBackupBtn.textContent = 'Start Backup';
        } else {
            backupButton.classList.remove('no-selection');
            backupButton.classList.add('has-selection');
            startBackupBtn.disabled = false;
            startBackupBtn.textContent = `Backup ${count} Conversations`;
        }
    }

    // Escape HTML to prevent XSS
    function escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    function getTeamsHostForReport(url) {
        try {
            const hostname = new URL(url || '').hostname;
            return ['teams.microsoft.com', 'teams.live.com', 'teams.cloud.microsoft'].includes(hostname)
                ? hostname
                : 'other';
        } catch {
            return 'unknown';
        }
    }

    function getInlineErrorDetails(result) {
        if (!result || !result.errorDetails) return '';
        const details = result.errorDetails;
        const extraction = details.errors || {};
        const parts = [];

        if (details.type) parts.push(`reason=${details.type}`);
        if (extraction.selectorUsed) {
            parts.push(`selector=${extraction.selectorUsed}`);
        } else if (details.selectorUsed) {
            parts.push(`selector=${details.selectorUsed}`);
        }
        if (extraction.selectorMismatch || details.selectorMismatch) {
            parts.push('selectorMismatch=true');
        }
        if (extraction.topNotReached || details.topNotReached) {
            parts.push('topNotReached=true');
        }
        const scrollDiagnostics = extraction.scrollDiagnostics || details.scrollDiagnostics;
        if (scrollDiagnostics) {
            if (scrollDiagnostics.strategy) {
                parts.push(`strategy=${scrollDiagnostics.strategy}`);
            }
            if (typeof scrollDiagnostics.finalScrollTop === 'number') {
                parts.push(`scrollTop=${scrollDiagnostics.finalScrollTop}`);
            }
            if (typeof scrollDiagnostics.clientHeight === 'number' && typeof scrollDiagnostics.scrollHeight === 'number') {
                parts.push(`viewport=${scrollDiagnostics.clientHeight}/${scrollDiagnostics.scrollHeight}`);
            }
            if (Array.isArray(scrollDiagnostics.strategies)) {
                const failedStrategies = scrollDiagnostics.strategies
                    .filter(s => s && s.success === false)
                    .map(s => s.name)
                    .slice(0, 3);
                if (failedStrategies.length > 0) {
                    parts.push(`failedStrategies=${failedStrategies.join('|')}`);
                }
            }
        }
        if (typeof extraction.messageErrors === 'number' && extraction.messageErrors > 0) {
            parts.push(`messageErrors=${extraction.messageErrors}`);
        }
        const containerDiagnostics = extraction.containerDiagnostics || details.containerDiagnostics;
        if (containerDiagnostics && containerDiagnostics.mode) {
            parts.push(`containerMode=${containerDiagnostics.mode}`);
            if (containerDiagnostics.selected && containerDiagnostics.selected.dataTid) {
                parts.push(`containerTid=${containerDiagnostics.selected.dataTid}`);
            }
        }
        if (extraction.pdfError) {
            parts.push(`pdfError=${extraction.pdfError}`);
        }
        if (extraction.screenshotDiagnostics && typeof extraction.screenshotDiagnostics.attempts === 'number') {
            parts.push(`shotAttempts=${extraction.screenshotDiagnostics.attempts}`);
        }
        if (extraction.domDiagnostics || details.domDiagnostics) {
            parts.push('domSnapshot=true');
        }
        if (details.operationId) {
            parts.push(`op=${details.operationId}`);
        }

        return parts.length > 0 ? ` (${parts.join(', ')})` : '';
    }

    function buildIssueReportLink(_errorMessage, result, tabUrl) {
        const details = result?.errorDetails || {};
        const extraction = details.errors || {};
        const manifest = chrome.runtime.getManifest();
        const safeCount = value => Number.isSafeInteger(value) && value >= 0 ? value : 'n/a';
        const safeStatus = ['failed', 'error', 'cancelled', 'success'].includes(details.status)
            ? details.status
            : 'unknown';
        // Reports are user-sent drafts. Never include chat content or raw diagnostics.
        const reportLines = [
            'Teams Chat Backup - Failure Report',
            `Extension Version: ${manifest.version}`,
            `Teams Host: ${getTeamsHostForReport(tabUrl)}`,
            `Response Status: ${safeStatus}`,
            `Retryable: ${details.retryable === true ? 'yes' : 'no'}`,
            `Selector Mismatch: ${extraction.selectorMismatch === true || details.selectorMismatch === true ? 'yes' : 'no'}`,
            `Top Not Reached: ${extraction.topNotReached === true || details.topNotReached === true ? 'yes' : 'no'}`,
            `Message Errors: ${safeCount(extraction.messageErrors)}`,
            `Image Errors: ${safeCount(extraction.imageErrors)}`,
            `File Errors: ${safeCount(extraction.fileErrors)}`,
            `Container Lost: ${extraction.containerLost === true ? 'yes' : 'no'}`,
            '',
            'Description (avoid including private chat information):',
            ''
        ];
        const subject = encodeURIComponent(`Teams Chat Backup failure (${manifest.version})`);
        const body = encodeURIComponent(reportLines.join('\n'));
        if (ISSUE_REPORT_EMAIL && ISSUE_REPORT_EMAIL.includes('@')) {
            return `mailto:${ISSUE_REPORT_EMAIL}?subject=${subject}&body=${body}`;
        }
        return ISSUE_REPORT_FALLBACK_URL;
    }

    // Track if dropdown is open
    let dropdownOpen = false;

    // Track if we need a page refresh (content script not injected)
    let needsPageRefresh = false;

    // Function to refresh the Teams tab and close popup
    function refreshTeamsTab() {
        chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
            if (tabs.length > 0) {
                chrome.tabs.reload(tabs[0].id);
                // Close popup after a brief delay so user sees the action
                setTimeout(() => {
                    window.close();
                }, 300);
            }
        });
    }

    // Backup button - toggles the chat selection dropdown
    addSafeListener(backupButton, 'click', async () => {
        // Prevent starting new backup if one is already in progress
        if (backupInProgress) {
            showToast('A backup is already in progress. Please wait.', 3000, 'warning');
            return;
        }

        // Check if on Teams first
        const isTeamsTab = await checkIfTeamsTab();
        if (!isTeamsTab) {
            statusDiv.innerHTML = 'Please open <a href="https://teams.microsoft.com/" target="_blank">Microsoft Teams</a>, select a chat you want to backup, then try again.';
            statusDiv.style.color = "#dc3545";
            statusDiv.style.display = "block";
            return;
        }

        dropdownOpen = !dropdownOpen;

        if (dropdownOpen) {
            // Show the dropdown and load chats
            chatSelectionContainer.style.display = 'block';
            backupButton.classList.add('dropdown-open');

            // Load chats if not already loaded
            if (availableChats.length === 0 && !needsPageRefresh) {
                chatList.innerHTML = '<div class="chat-selection__empty">Loading chats and channels...</div>';
                const result = await loadAvailableChats();

                if (result.error === 'no_connection') {
                    // Content script not injected - show refresh message
                    needsPageRefresh = true;
                    chatSelectionContainer.style.display = 'none';
                    backupButton.classList.remove('dropdown-open');
                    dropdownOpen = false;

                    statusDiv.innerHTML = 'Please <a href="#" id="refreshPageLink">refresh the page</a> to enable chat backup.';
                    statusDiv.style.color = "#dc3545";
                    statusDiv.style.display = "block";

                    // Add click handler for refresh link
                    const refreshLink = document.getElementById('refreshPageLink');
                    if (refreshLink) {
                        refreshLink.addEventListener('click', (e) => {
                            e.preventDefault();
                            refreshTeamsTab();
                        });
                    }
                    return;
                }

                if (result.error === 'timeout' || result.error === 'request_failed') {
                    console.warn('Loading chats fallback:', result.error, result.diagnostics || null);
                }

                availableChats = result.chats;
                chatLoadDiagnostics = result.diagnostics || null;
            }

            // Auto-select the current/active chat if found. This also covers the
            // list preloaded when the popup opened.
            const activeIndex = availableChats.findIndex(c => c.isActive);
            if (activeIndex >= 0 && selectedChatIndices.length === 0 && isChatSupportedForBackup(availableChats[activeIndex])) {
                selectedChatIndices = [activeIndex];
            } else if (selectedChatIndices.length === 0) {
                const firstSupportedIndex = availableChats.findIndex(c => isChatSupportedForBackup(c));
                if (firstSupportedIndex >= 0) {
                    selectedChatIndices = [firstSupportedIndex];
                }
            }

            renderChatList();
            updateBackupButtonState();
        } else {
            // Hide the dropdown
            chatSelectionContainer.style.display = 'none';
            backupButton.classList.remove('dropdown-open');
        }
    }, 'backup button');

    // Start Backup button - actually starts the backup process
    addSafeListener(startBackupBtn, 'click', async () => {
        // Prevent concurrent backups
        if (backupInProgress) {
            showToast('A backup is already in progress. Please wait.', 3000, 'warning');
            return;
        }

        if (selectedChatIndices.length === 0) return;

        const withinLimit = await passesFreeLimitPrecheck();
        if (!withinLimit) {
            return;
        }

        // Close dropdown
        dropdownOpen = false;
        chatSelectionContainer.style.display = 'none';
        backupButton.classList.remove('dropdown-open');

        // Start the backup
        startBackup();
    }, 'start backup button');

    // Toast notification function
    // type: 'info' (default), 'warning', 'success'
    function showToast(message, duration = 4000, type = 'info') {
        // Remove existing toast if any
        const existingToast = document.querySelector('.toast-notification');
        if (existingToast) {
            existingToast.remove();
        }

        const toast = document.createElement('div');
        toast.className = `toast-notification toast-notification--${type}`;

        // Different icons based on type
        let icon = '↻';
        if (type === 'warning') {
            icon = '!';
        } else if (type === 'success') {
            icon = '✓';
        }

        const iconSpan = document.createElement('span');
        iconSpan.className = 'toast-icon';
        iconSpan.textContent = icon;

        const messageSpan = document.createElement('span');
        messageSpan.className = 'toast-message';
        messageSpan.textContent = message;

        toast.appendChild(iconSpan);
        toast.appendChild(messageSpan);
        document.body.appendChild(toast);

        // Trigger animation
        setTimeout(() => toast.classList.add('show'), 10);

        // Auto-remove after duration
        setTimeout(() => {
            toast.classList.remove('show');
            setTimeout(() => toast.remove(), 300);
        }, duration);
    }

    // Function to download files
    async function downloadFiles() {
        let downloadAccessGranted = false;
        try {
            downloadAccessGranted = await requestOptionalAccess({ permissions: ['downloads'] });
        } catch (error) {
            console.error('Could not request download access:', error);
        }

        if (!downloadAccessGranted) {
            statusDiv.textContent = 'Download access was not granted. Your backup remains saved in the Viewer.';
            statusDiv.style.color = '#ff9800';
            statusDiv.style.display = 'block';
            return;
        }

        const timestamp = Date.now();
        let downloadCount = 0;
        let expectedDownloads = 0;

        if (downloadedPdfUrl) expectedDownloads++;
        if (downloadedTextUrl) expectedDownloads++;

        function checkComplete() {
            downloadCount++;
            if (downloadCount === expectedDownloads) {
                statusDiv.textContent = "Downloads initiated.";
                statusDiv.style.color = "#4caf50";
                statusDiv.style.display = "block";
                downloadContainer.style.display = "none";

                // Clean up URLs
                if (downloadedPdfUrl) URL.revokeObjectURL(downloadedPdfUrl);
                if (downloadedTextUrl) URL.revokeObjectURL(downloadedTextUrl);
                downloadedPdfUrl = null;
                downloadedTextUrl = null;
            }
        }

        // Download PDF
        if (downloadedPdfUrl) {
            chrome.downloads.download({
                url: downloadedPdfUrl,
                filename: `teams_chat_backup_${timestamp}.pdf`,
                saveAs: true
            }, () => checkComplete());
        }

        // Download Text
        if (downloadedTextUrl) {
            chrome.downloads.download({
                url: downloadedTextUrl,
                filename: `teams_chat_backup_${timestamp}.txt`,
                saveAs: true
            }, () => checkComplete());
        }
    }

    // Check IndexedDB compatibility (detects incognito mode issues)
    async function checkIndexedDBCompatibility() {
        try {
            // Try to open IndexedDB
            const testDB = await new Promise((resolve, reject) => {
                const request = indexedDB.open('_test_db', 1);
                request.onsuccess = () => {
                    request.result.close();
                    indexedDB.deleteDatabase('_test_db');
                    resolve(true);
                };
                request.onerror = () => reject(request.error);
                request.onblocked = () => reject(new Error('IndexedDB blocked'));
            });
            return { compatible: true };
        } catch (error) {
            console.error('IndexedDB compatibility check failed:', error);
            return {
                compatible: false,
                error: 'Storage unavailable. If using incognito mode, please switch to a normal window.'
            };
        }
    }

    // Check if running in incognito mode
    async function isIncognitoMode() {
        return new Promise((resolve) => {
            chrome.storage.local.set({ _test: 1 }, () => {
                chrome.storage.local.get('_test', (result) => {
                    resolve(!result._test);
                    chrome.storage.local.remove('_test');
                });
            });
        });
    }

    // Function to backup a single chat
    async function backupSingleChat(tabId, chatIndex, totalChats, currentIndex) {
        return new Promise(async (resolve) => {
            // Popup backups always run in viewer-export mode (no PDF generation here).
            const generatePDF = FORCE_VIEWER_EXPORT_ONLY ? false : !storedTurboMode;
            const generateTXT = true; // Always generate text
            const includeCover = FORCE_VIEWER_EXPORT_ONLY ? false : (isPro ? includeCoverCheckbox.checked : true);
            // Get capture interval from settings (ignored in turbo mode)
            const captureInterval = getCaptureInterval();

            // If we need to switch to a different chat, do it first
            if (chatIndex !== undefined && availableChats.length > 0) {
                const currentActive = availableChats.findIndex(c => c.isActive);
                if (currentActive !== chatIndex) {
                    statusDiv.textContent = `Opening chat ${currentIndex + 1} of ${totalChats}...`;
                    const chatTarget = availableChats[chatIndex] || null;
                    try {
                        const openResponse = await sendTabMessage(tabId, {
                            action: "openChat",
                            chatIndex,
                            chatTarget
                        });
                        if (!openResponse || openResponse.status !== "success") {
                            resolve({
                                success: false,
                                error: 'Could not open the selected chat. Please try again.',
                                errorDetails: {
                                    type: 'open_chat_failed',
                                    tabId,
                                    chatIndex
                                }
                            });
                            return;
                        }
                        // Keep the cached list in sync with Teams so a later backup of the
                        // originally open chat switches back to it instead of saving this one.
                        availableChats.forEach((chat, index) => {
                            chat.isActive = index === chatIndex;
                        });
                        await delay(500);
                    } catch (error) {
                        resolve({
                            success: false,
                            error: 'Connection error',
                            errorDetails: {
                                type: 'runtime_last_error',
                                runtimeError: error?.message || 'Unknown runtime error',
                                tabId,
                                chatIndex,
                                during: 'open_chat',
                                retryable: isRecoverableContentScriptError(error?.message || '')
                            }
                        });
                        return;
                    }
                }
            }

            // Only continue immediately if we're already on the correct chat
            continueBackup();

            async function continueBackup() {
                const statusPrefix = totalChats > 1 ? `[${currentIndex + 1}/${totalChats}] ` : '';
                statusDiv.textContent = statusPrefix + "Starting backup...";

                const selectedChatName = chatIndex !== undefined && availableChats.length > 0
                    ? availableChats[chatIndex].name
                    : null;
                const selectedChatMeta = chatIndex !== undefined && availableChats.length > 0
                    ? (availableChats[chatIndex] || null)
                    : (availableChats.find(chat => chat.isActive) || null);
                const selectedChatLookup = selectedChatMeta || selectedChatName;

                // Get chat name for looking up existing attachments
                let existingFilenames = [];
                let incrementalFrom = null;
                try {
                    if (selectedChatLookup && storedFileAttachments) {
                        existingFilenames = await window.TeamsBackupDB.getExistingAttachmentFilenames(selectedChatLookup);
                        console.log(`Found ${existingFilenames.length} existing attachment files for "${selectedChatName}"`);
                    }

                    // Auto-incremental mode:
                    // stop once we reach the latest timestamp we already have.
                    // Long-chat mode intentionally skips this shortcut so older incomplete
                    // backups can heal themselves through a full checkpointed rescan.
                    if (!storedLongChatMode &&
                        selectedChatLookup &&
                        typeof window.TeamsBackupDB.getLatestMessageTimestamp === 'function') {
                        incrementalFrom = await window.TeamsBackupDB.getLatestMessageTimestamp(selectedChatLookup);
                        if (incrementalFrom) {
                            console.log(`Incremental backup anchor for "${selectedChatName}": ${incrementalFrom}`);
                        }
                    } else if (storedLongChatMode) {
                        console.log(`Long chat mode enabled for "${selectedChatName}". Forcing full checkpointed rescan.`);
                    }
                } catch (e) {
                    console.log('Could not get existing attachment filenames:', e);
                }

                const backupRequest = {
                    action: "backupChat",
                    captureInterval: captureInterval,
                    isPro: isPro,
                    generatePDF: generatePDF,
                    generateTXT: generateTXT,
                    includeCover: includeCover,
                    longChatMode: storedLongChatMode,
                    incrementalFrom: incrementalFrom || null,
                    fileAttachments: storedFileAttachments,
                    maxFileSize: storedMaxFileSize,
                    existingFilenames: existingFilenames,
                    checkpointing: {
                        enabled: !generatePDF,
                        // Without a detected sidebar entry, let the content script name the
                        // chat from the page instead of saving every chat as one fixed name.
                        chatIdentity: (selectedChatMeta?.name || selectedChatName) ? {
                            name: selectedChatMeta?.name || selectedChatName,
                            rawName: selectedChatMeta?.rawName || selectedChatMeta?.name || selectedChatName,
                            isChannel: selectedChatMeta?.isChannel === true,
                            teamName: selectedChatMeta?.teamName || null,
                            sourceKind: selectedChatMeta?.kind || (selectedChatMeta?.isChannel ? 'channel' : 'chat')
                        } : null
                    }
                };

                let response;
                try {
                    response = await sendBackupChatMessageWithRecovery(tabId, backupRequest, {
                        statusPrefix,
                        chatIndex,
                        chatTarget: selectedChatMeta || null,
                        totalChats,
                        currentIndex
                    });
                } catch (error) {
                    const runtimeError = error?.message || 'Unknown runtime error';
                    const retryable = !!error?.retryable || isRecoverableContentScriptError(runtimeError);
                    resolve({
                        success: false,
                        error: retryable ? 'Teams reloaded during backup' : 'Connection error',
                        errorDetails: {
                            type: 'runtime_last_error',
                            runtimeError,
                            tabId,
                            retryable,
                            recoveryAttempted: !!error?.recoveryAttempted,
                            recoveryError: error?.recoveryError || null
                        }
                    });
                    return;
                }

                // Critical: Check if response exists before accessing properties
                if (!response) {
                    resolve({
                        success: false,
                        error: 'No response from content script. Please refresh the page and try again.',
                        errorDetails: {
                            type: 'no_response',
                            tabId
                        }
                    });
                    return;
                }

                if (isBackupAlreadyInProgressResponse(response)) {
                    try {
                        response = await recoverBusyBackupLock(tabId, backupRequest, {
                            statusPrefix,
                            chatIndex,
                            chatTarget: selectedChatMeta || null,
                            totalChats,
                            currentIndex
                        });
                    } catch (error) {
                        const runtimeError = error?.message || 'Unknown runtime error';
                        const retryable = !!error?.retryable || isRecoverableContentScriptError(runtimeError);
                        resolve({
                            success: false,
                            error: retryable ? 'Teams reloaded during backup' : 'Connection error',
                            errorDetails: {
                                type: 'runtime_last_error',
                                runtimeError,
                                tabId,
                                retryable,
                                recoveryAttempted: true,
                                recoveryError: null
                            }
                        });
                        return;
                    }

                    if (!response) {
                        resolve({
                            success: false,
                            error: 'Previous backup is still stopping. Please wait a few seconds and try again.',
                            errorDetails: {
                                type: 'content_script_busy_timeout',
                                tabId,
                                retryable: true
                            }
                        });
                        return;
                    }
                }

                if (response.status === "completed") {
                        const isCheckpointed = response.checkpointed === true;
                        const capturedMessageCount = Number.isFinite(response.messageCount)
                            ? response.messageCount
                            : (response.messages ? response.messages.length : 0);
                        // Save to IndexedDB - even without PDF, save messages and text
                        // Improved validation: check actual message count, not just truthy text
                        const hasActualMessages = capturedMessageCount > 0;
                        const hasContent = response.pdfUrl || hasActualMessages;

                        if (!hasActualMessages && response.textContent) {
                            // Text content exists but no messages - likely header-only
                            console.warn('Backup has text content but no messages extracted');
                        }

                        // Validate that we actually captured content
                        if (!hasContent) {
                            const errorDetails = {
                                messageCount: response.messages ? response.messages.length : 0,
                                hasPdf: !!response.pdfUrl,
                                hasText: !!response.textContent,
                                errors: response.errors || response.extractionErrors || {}
                            };
                            console.error('Backup failed - no content captured:', errorDetails);

                            let errorMessage = 'No messages were captured. ';
                            const errors = response.errors || response.extractionErrors;
                            if (errors) {
                                const { messageErrors, imageErrors, containerLost } = errors;
                                if (containerLost) {
                                    errorMessage += 'Chat container was lost during extraction. ';
                                }
                                if (messageErrors > 0) {
                                    errorMessage += `${messageErrors} messages could not be read. `;
                                }
                            }
                            errorMessage += 'Please try again or contact support if issue persists.';

                            resolve({
                                success: false,
                                error: errorMessage,
                                errorDetails: {
                                    type: 'no_captured_content',
                                    status: response.status,
                                    operationId: response.operationId || null,
                                    errors: response.errors || response.extractionErrors || {}
                                }
                            });
                            return;
                        }

                        if (hasContent) {
                            if (isCheckpointed) {
                                lastSavedChatId = response.checkpointChatId || null;
                                updateBackupCount();
                                resolve({
                                    success: true,
                                    pdfUrl: response.pdfUrl,
                                    textUrl: response.textUrl,
                                    chatName: response.chatName,
                                    merged: response.checkpointMerged === true,
                                    newMessagesCount: Number(response.checkpointNewMessagesCount || 0),
                                    extractionErrors: response.extractionErrors,
                                    partialBackup: !!response.partialBackup,
                                    partialReason: response.partialReason || null,
                                    checkpointed: true
                                });
                                return;
                            }

                            try {
                                let pdfBlob = null;
                                if (response.pdfUrl) {
                                    try {
                                        const pdfResponse = await fetch(response.pdfUrl);
                                        if (!pdfResponse.ok) {
                                            throw new Error(`Blob fetch failed: ${pdfResponse.status}`);
                                        }
                                        pdfBlob = await pdfResponse.blob();
                                    } catch (blobError) {
                                        console.error('Failed to fetch PDF blob:', blobError);
                                        // Don't fail entire backup - just skip PDF
                                        pdfBlob = null;
                                        showToast('Warning: PDF generation failed, but text backup was saved', 5000, 'warning');
                                    }
                                }

                                const saveResult = await window.TeamsBackupDB.saveChat({
                                    name: selectedChatMeta?.name || response.chatName || 'Teams Chat Backup',
                                    rawName: selectedChatMeta?.rawName || response.chatName || 'Teams Chat Backup',
                                    isChannel: selectedChatMeta?.isChannel === true,
                                    teamName: selectedChatMeta?.teamName || null,
                                    sourceKind: selectedChatMeta?.kind || (selectedChatMeta?.isChannel ? 'channel' : 'chat'),
                                    messages: response.messages || [],
                                    participants: response.participants || [],
                                    pdfBlob: pdfBlob,
                                    textContent: response.textContent || '',
                                    isPreview: !!response.partialBackup,
                                    isPro: isPro // Pass Pro status to control merge behavior
                                });

                                // Handle free user chat limit error
                                if (saveResult.error === 'chat_limit') {
                                    resolve({
                                        success: false,
                                        error: 'chat_limit',
                                        reason: saveResult.reason,
                                        count: saveResult.count,
                                        usedChatNames: saveResult.usedChatNames,
                                        message: saveResult.message
                                    });
                                    return;
                                }

                                if (saveResult.error === 'update_locked') {
                                    resolve({
                                        success: false,
                                        error: 'update_locked',
                                        message: saveResult.message || 'Free tier backups cannot be updated.',
                                        existingChatId: saveResult.id || null
                                    });
                                    return;
                                }

                                lastSavedChatId = saveResult.id;
                                console.log('Chat saved to IndexedDB:', saveResult);
                                updateBackupCount();

                                // Clean up Blob URLs in content script after saving
                                chrome.tabs.sendMessage(tabId, { action: "cleanupBlobUrls" });

                                resolve({
                                    success: true,
                                    pdfUrl: response.pdfUrl,
                                    textUrl: response.textUrl,
                                    chatName: response.chatName,
                                    merged: saveResult.merged,
                                    newMessagesCount: saveResult.newMessagesCount,
                                    extractionErrors: response.extractionErrors,
                                    partialBackup: !!response.partialBackup,
                                    partialReason: response.partialReason || null
                                });
                            } catch (error) {
                                console.error('Error saving to IndexedDB:', error);
                                resolve({
                                    success: false,
                                    error: 'Save error',
                                    errorDetails: {
                                        type: 'indexeddb_save_error',
                                        message: error?.message || 'Unknown save error'
                                    }
                                });
                            }
                        } else {
                            resolve({
                                success: false,
                                error: 'No content captured',
                                errorDetails: {
                                    type: 'no_content',
                                    status: response.status,
                                    operationId: response.operationId || null
                                }
                            });
                        }
                } else if (response.status === "cancelled") {
                    resolve({ success: false, cancelled: true });
                } else if (response.error === 'chat_limit' || response.error === 'update_locked') {
                    // Free-tier limits reported by the checkpoint save path.
                    resolve({
                        success: false,
                        error: response.error,
                        count: response.count,
                        message: response.message
                    });
                } else {
                    resolve({
                        success: false,
                        error: response.message || 'Backup failed',
                        emptyChannel: !!response.emptyChannel,
                        errorDetails: {
                            type: 'content_script_failed',
                            status: response.status || 'unknown',
                            operationId: response.operationId || null,
                            retryable: !!response.retryable,
                            unsupportedLayout: response.unsupportedLayout || null,
                            errors: response.errors || response.extractionErrors || {}
                        }
                    });
                }
            }
        });
    }

    // Function to start the backup process
    async function startBackup() {
        // Check storage compatibility first (detects incognito mode)
        const storageCheck = await checkIndexedDBCompatibility();
        if (!storageCheck.compatible) {
            statusDiv.textContent = storageCheck.error;
            statusDiv.style.color = '#dc3545';
            statusDiv.style.display = 'block';
            return;
        }

        toggleButtons(true);
        backupInProgress = true;
        if (ratingPrompt) ratingPrompt.hidden = true;
        monitoringDetachedBackup = false;
        cancellationReason = null; // Reset cancellation reason for new backup
        unfocusedDuringBackup = false; // Reset unfocus tracking
        unfocusWarningShown = false;
        currentBackupStartTime = Date.now(); // Track when backup started
        currentBackupChatName = null; // Will be set when we know which chat

        statusDiv.textContent = "Starting backup...";
        statusDiv.style.color = "#2c3e50";
        statusDiv.style.display = "block";
        loadingMessageDiv.textContent = "Loading Content...";
        loadingMessageDiv.style.display = "block";
        progressContainer.style.display = "block";
        progressBar.style.width = "0%";
        downloadContainer.style.display = "none";
        if (backupHint) {
            backupHint.style.display = "block";
        }
        updateBackupHintForBackup(); // Show appropriate warning based on Pro status

        // Save initial backup state
        writeStorage('local', { backupInProgress: true });
        saveBackupState({
            chatName: 'Starting...',
            progress: 0,
            phase: 'Initializing',
            startTime: currentBackupStartTime
        });

        function formatBackupError(error) {
            if (!error) return "Backup failed or was cancelled.";
            if (error === 'Connection error') {
                return 'Could not communicate with Teams. Please refresh Teams and try again.';
            }
            if (error === 'Teams reloaded during backup') {
                return 'Teams refreshed or reloaded during backup. Reopen the chat and try again.';
            }
            if (error === 'Save error') {
                return 'Backup was captured but could not be saved locally. Please free up storage and try again.';
            }
            if (error === 'No content captured') {
                return 'No messages were captured from this chat. Please open the chat and try again.';
            }
            if (error === 'update_locked') {
                return UPDATE_LOCKED_MESSAGE;
            }
            return error;
        }

        chrome.tabs.query({ active: true, currentWindow: true }, async (tabs) => {
            if (tabs.length === 0) {
                statusDiv.textContent = "No active tab found.";
                statusDiv.style.color = "#dc3545";
                statusDiv.style.display = "block";
                loadingMessageDiv.style.display = "none";
                if (backupHint) {
                    backupHint.style.display = "none";
                }
                toggleButtons(false);
                return;
            }

            const tabId = tabs[0].id;
            const activeTabUrl = tabs[0].url || '';

            // Determine which chats to backup
            // undefined means "current chat" (used when no chats detected)
            let chatsToBackup = selectedChatIndices.length > 0
                ? selectedChatIndices.map(idx => idx === -1 ? undefined : idx)
                : [undefined]; // Fallback to current chat if nothing selected
            const totalChats = chatsToBackup.length;
            let successCount = 0;
            let mergedCount = 0;
            let partialCount = 0;
            let failedCount = 0;
            let lastResult = null;

            for (let i = 0; i < chatsToBackup.length; i++) {
                if (!backupInProgress) break; // User cancelled

                // Update current chat name for progress tracking
                const chatIndex = chatsToBackup[i];
                if (chatIndex !== undefined && availableChats[chatIndex]) {
                    currentBackupChatName = availableChats[chatIndex].name;
                } else {
                    currentBackupChatName = 'Current Chat';
                }

                loadingMessageDiv.textContent = totalChats > 1 ? `Backing up chat ${i + 1} of ${totalChats}...` : "Loading Content...";

                // Save state with current chat info
                saveBackupState({
                    chatName: currentBackupChatName,
                    progress: 0,
                    phase: totalChats > 1 ? `Chat ${i + 1} of ${totalChats}` : 'Scrolling to top',
                    startTime: currentBackupStartTime
                });

                const result = await backupSingleChat(tabId, chatsToBackup[i], totalChats, i);

                if (result.success) {
                    successCount++;
                    if (result.merged) {
                        mergedCount++;
                    }
                    if (result.partialBackup) {
                        partialCount++;
                    }
                    lastResult = result;
                } else if (result.cancelled) {
                    break;
                } else if (result.emptyChannel) {
                    // Empty channel - skip and continue to next
                    console.log(`Skipping empty channel at index ${chatsToBackup[i]}`);
                    if (totalChats > 1) {
                        loadingMessageDiv.textContent = `Skipping empty channel (${i + 1}/${totalChats})...`;
                    }
                    continue;
                } else if (result.error === 'chat_limit') {
                    lastResult = result;
                    break;
                } else if (result.error === 'update_locked') {
                    lastResult = result;
                    break;
                } else if (result.error === 'chat_exists') {
                    // Free user tried to backup a conversation that already exists
                    lastResult = result;
                    break; // Stop processing
                } else {
                    failedCount++;
                    lastResult = result;
                }

                // Small delay between chats
                if (i < chatsToBackup.length - 1) {
                    await new Promise(r => setTimeout(r, 300));
                }
            }

            backupInProgress = false;
            currentBackupChatName = null;
            currentBackupStartTime = null;
            clearBackupState(); // Clear persisted state

            if (successCount > 0) {
                // Determine the appropriate success message
                let statusMessage = '';
                if (isPro) {
                    if (successCount === 1) {
                        if (lastResult.merged) {
                            const newMsgs = lastResult.newMessagesCount || 0;
                            statusMessage = newMsgs > 0
                                ? `Backup updated! ${newMsgs} new message${newMsgs !== 1 ? 's' : ''} added.`
                                : "Backup updated! Already up to date.";
                        } else {
                            statusMessage = "Backup completed and saved!";
                        }
                    } else {
                        statusMessage = mergedCount > 0
                            ? `${successCount} chats backed up! (${mergedCount} updated)`
                            : `${successCount} chats backed up and saved!`;
                    }
                } else {
                    if (lastResult && lastResult.merged) {
                        statusMessage = "Backup updated and saved!";
                    } else {
                        statusMessage = "Backup completed and saved!";
                    }
                }
                statusDiv.textContent = statusMessage;
                statusDiv.style.color = "#4caf50";

                if (partialCount > 0) {
                    if (successCount === 1 && partialCount === 1) {
                        statusDiv.textContent = "Backup completed (visible messages only).";
                    } else {
                        statusDiv.textContent = `${statusMessage} (${partialCount} partial)`;
                    }
                }

                // Multi-chat runs continue past a failed chat; don't hide it behind the success message.
                if (failedCount > 0) {
                    statusDiv.textContent += ` ${failedCount} chat${failedCount !== 1 ? 's' : ''} could not be backed up. Try ${failedCount !== 1 ? 'them' : 'it'} again.`;
                    statusDiv.style.color = "#ff9800";
                }

                // Show toast for merged backups
                if (lastResult && lastResult.merged) {
                    showToast(lastResult.newMessagesCount > 0
                        ? `Updated existing backup with ${lastResult.newMessagesCount} new message${lastResult.newMessagesCount !== 1 ? 's' : ''}`
                        : 'Backup already exists - updated with latest content');
                }

                let runHadIssues = partialCount > 0 || failedCount > 0;

                // Show warning if there were extraction errors
                if (lastResult && lastResult.extractionErrors) {
                    const errors = lastResult.extractionErrors;
                    const hasErrors = errors.messageErrors > 0 || errors.imageErrors > 0 ||
                                      errors.fileErrors > 0 || errors.containerLost || errors.topNotReached ||
                                      !!errors.checkpointFinalizeError ||
                                      !!errors.pdfError;
                    if (hasErrors) {
                        let warningParts = [];
                        if (errors.messageErrors > 0) {
                            warningParts.push(`${errors.messageErrors} message${errors.messageErrors > 1 ? 's' : ''} couldn't be read`);
                        }
                        if (errors.imageErrors > 0) {
                            warningParts.push(`${errors.imageErrors} image${errors.imageErrors > 1 ? 's' : ''} failed to load`);
                        }
                        if (errors.fileErrors > 0) {
                            warningParts.push(`${errors.fileErrors} file${errors.fileErrors > 1 ? 's' : ''} couldn't be downloaded`);
                        }
                        if (errors.containerLost) {
                            warningParts.push('chat switched during extraction');
                        }
                        if (errors.topNotReached) {
                            warningParts.push('could not reach chat top (visible messages only)');
                        }
                        if (errors.pdfError) {
                            warningParts.push('PDF could not be generated (text backup saved)');
                        }
                        if (errors.checkpointFinalizeError) {
                            warningParts.push('final checkpoint metadata could not be finalized');
                        }
                        const warningMsg = `Backup completed with issues: ${warningParts.join(', ')}`;
                        showToast(warningMsg, 5000, 'warning');
                        runHadIssues = true;
                    }
                }

                // Don't show download container - user can access backups via View Backed Files
                // downloadContainer.style.display = "block";
                if (lastResult) {
                    downloadedPdfUrl = lastResult.pdfUrl;
                    downloadedTextUrl = lastResult.textUrl;
                }
                // Update the backup count to show new backup
                updateBackupCount();

                // Trigger pop animation on View Backed Files button
                triggerBackupCompleteAnimation();

                if (!runHadIssues) maybeShowRatingPrompt();
            } else if (lastResult && lastResult.error === 'chat_limit') {
                // Free user reached the chat backup limit
                renderFreeTierLimitMessage(lastResult);
            } else if (lastResult && lastResult.error === 'update_locked') {
                renderPlanLimitMessage(UPDATE_LOCKED_MESSAGE);
            } else {
                // Show appropriate message based on cancellation reason
                let cancelMessage = "Backup failed or was cancelled.";
                let messageColor = "#ff9800";
                let useHtmlStatus = false;
                if (cancellationReason === CANCEL_REASONS.USER_STOP) {
                    cancelMessage = "Backup stopped by user.";
                } else if (cancellationReason === CANCEL_REASONS.POPUP_CLOSED) {
                    cancelMessage = "Backup interrupted - popup was closed.";
                } else if (lastResult?.errorDetails?.unsupportedLayout === 'threads_posts') {
                    cancelMessage = "This channel uses the Posts/Threads layout and no readable messages were found. Open a conversation in the channel so messages are visible, then try again.";
                    messageColor = "#ff9800";
                } else if (lastResult && lastResult.error) {
                    cancelMessage = `${formatBackupError(lastResult.error)}${getInlineErrorDetails(lastResult)}`;
                    messageColor = "#dc3545";
                    const reportLink = buildIssueReportLink(cancelMessage, lastResult, activeTabUrl);
                    statusDiv.innerHTML = `${escapeHtml(cancelMessage)} <a href="${reportLink}" target="_blank" rel="noopener noreferrer">Report this issue</a>`;
                    useHtmlStatus = true;
                }
                if (!useHtmlStatus) {
                    statusDiv.textContent = cancelMessage;
                }
                statusDiv.style.color = messageColor;

                // Reset cancellation reason
                cancellationReason = null;
            }

            statusDiv.style.display = "block";
            loadingMessageDiv.style.display = "none";
            progressContainer.style.display = "none";
            if (backupHint) {
                backupHint.style.display = "none";
            }
            toggleButtons(false);
        });
    }

    addSafeListener(stopButton, 'click', () => {
        cancellationReason = CANCEL_REASONS.USER_STOP;

        chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
            chrome.tabs.sendMessage(tabs[0].id, { action: "stopScrolling" }, (response) => {
                if (response && response.status === "stopped") {
                    statusDiv.textContent = "Stopped by user. Starting capture from current position...";
                    statusDiv.style.color = "#ff9800";
                    statusDiv.style.display = "block";
                    toggleButtons(false);
                }
            });
        });
    }, 'stop button');

    addSafeListener(includeCoverCheckbox, 'change', () => {
        if (FORCE_VIEWER_EXPORT_ONLY) {
            includeCoverCheckbox.checked = false;
            if (settingsIncludeCover) settingsIncludeCover.checked = false;
            return;
        }

        if (isPro) {
            storedProIncludeCover = includeCoverCheckbox.checked;
            writeStorage('local', { [PRO_INCLUDE_COVER_KEY]: storedProIncludeCover });
            writeStorage('sync', { [PRO_INCLUDE_COVER_KEY]: storedProIncludeCover });
        }
    }, 'include cover checkbox');

    // Update download button to handle both files
    addSafeListener(downloadButton, 'click', downloadFiles, 'download button');

    // View backup button - open viewer with the last saved chat
    addSafeListener(viewBackupButton, 'click', () => {
        openOrFocusViewer({ chatId: lastSavedChatId || null, forceRoot: !lastSavedChatId }).catch((error) => {
            console.error('Failed to open viewer:', error);
        });
    }, 'view backup button');

    // View Backed Files button - opens the full viewer
    addSafeListener(viewBackedFilesButton, 'click', () => {
        // Check if backup is in progress
        if (backupInProgress) {
            showBackupConfirmModal();
            return;
        }
        openOrFocusViewer({ refresh: true }).catch((error) => {
            console.error('Failed to open viewer:', error);
        });
    }, 'view backed files button');

    // Update backup count on the green button
    async function updateBackupCount() {
        try {
            const count = await window.TeamsBackupDB.getChatCount();
            if (count > 0) {
                backedFilesCount.textContent = count;
            } else {
                backedFilesCount.textContent = '';
            }
        } catch (error) {
            console.error('Error getting backup count:', error);
        }
    }

    // Trigger pop animation on View Backed Files button when backup completes
    // Ask for a rating after a clean backup: first time, then once more 3 clean backups after
    // "Not now". Clicking Rate, or a second "Not now", stops it for good. No incentive is offered.
    const RATING_PROMPT_KEY = 'ratingPrompt';
    const ratingPrompt = document.getElementById('ratingPrompt');

    async function maybeShowRatingPrompt() {
        if (!ratingPrompt) return;
        const data = await readStorage('local', [RATING_PROMPT_KEY]).catch(() => ({}));
        const state = { cleanBackups: 0, dismissals: 0, showAt: 1, done: false, ...(data?.[RATING_PROMPT_KEY] || {}) };
        if (state.done) return;
        state.cleanBackups += 1;
        await writeStorage('local', { [RATING_PROMPT_KEY]: state }).catch(() => {});
        if (state.cleanBackups >= state.showAt) ratingPrompt.hidden = false;
    }

    async function closeRatingPrompt(rated) {
        ratingPrompt.hidden = true;
        const data = await readStorage('local', [RATING_PROMPT_KEY]).catch(() => ({}));
        const state = { cleanBackups: 0, dismissals: 0, showAt: 1, done: false, ...(data?.[RATING_PROMPT_KEY] || {}) };
        if (rated) {
            state.done = true;
        } else {
            state.dismissals += 1;
            state.done = state.dismissals >= 2;
            state.showAt = state.cleanBackups + 3;
        }
        await writeStorage('local', { [RATING_PROMPT_KEY]: state }).catch(() => {});
    }

    document.getElementById('ratingPromptRate')?.addEventListener('click', () => closeRatingPrompt(true));
    document.getElementById('ratingPromptLater')?.addEventListener('click', () => closeRatingPrompt(false));

    function triggerBackupCompleteAnimation() {
        // Remove any existing animation classes
        viewBackedFilesButton.classList.remove('pop-animation', 'highlight');

        // Force reflow to reset animation
        void viewBackedFilesButton.offsetWidth;

        // Add pop animation
        viewBackedFilesButton.classList.add('pop-animation');

        // After pop animation, add highlight glow
        setTimeout(() => {
            viewBackedFilesButton.classList.add('highlight');
        }, 600);

        // Remove animation classes after animations complete
        setTimeout(() => {
            viewBackedFilesButton.classList.remove('pop-animation', 'highlight');
        }, 3600); // 600ms pop + 3000ms highlight (3 pulses)
    }

    // Initialize backup count
    updateBackupCount();

    // Listen for progress updates from content script
    chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
        if (request.action === "updateProgress") {
            const progress = request.progress;
            progressContainer.style.display = "block";
            progressBar.style.width = `${progress}%`;
            loadingMessageDiv.textContent = `Processing: ${Math.round(progress)}%`;

            // Save progress to storage for persistence across popup close/reopen
            saveBackupState({
                chatName: currentBackupChatName,
                progress: progress,
                phase: request.phase || 'Processing',
                startTime: currentBackupStartTime
            });
        } else if (request.action === "updatePhase") {
            // Content script is reporting which phase it's in
            loadingMessageDiv.textContent = request.phase || 'Processing...';
            saveBackupState({
                chatName: currentBackupChatName,
                progress: request.progress || 0,
                phase: request.phase || 'Processing',
                startTime: currentBackupStartTime
            });
        } else if (request.action === "fileDownloadProgress") {
            // Show file download progress
            const { downloaded, total, estimatedSeconds } = request;
            let message = `Downloading files: ${downloaded} of ${total}`;

            if (estimatedSeconds > 0 && downloaded < total) {
                if (estimatedSeconds < 60) {
                    message += ` (~${estimatedSeconds}s remaining)`;
                } else {
                    const minutes = Math.ceil(estimatedSeconds / 60);
                    message += ` (~${minutes}m remaining)`;
                }
            } else if (downloaded === total) {
                message = `Downloaded ${total} file${total !== 1 ? 's' : ''}`;
            }

            loadingMessageDiv.textContent = message;
            loadingMessageDiv.style.display = "block";
        } else if (request.action === "backupComplete") {
            if (monitoringDetachedBackup) {
                monitoringDetachedBackup = false;
                backupInProgress = false;
                currentBackupChatName = null;
                currentBackupStartTime = null;
                clearBackupState();
                toggleButtons(false);
                loadingMessageDiv.style.display = "none";
                progressContainer.style.display = "none";
                if (backupHint) {
                    backupHint.style.display = "none";
                }
                statusDiv.textContent = "Background backup reached the end. Saved checkpoints are available in View Backed Files.";
                statusDiv.style.color = "#4caf50";
                statusDiv.style.display = "block";
                updateBackupCount();
                triggerBackupCompleteAnimation();
                return;
            }

            // Each chat reports completion; keep Stop visible until a multi-chat run
            // finishes (startBackup restores the buttons at the end).
            if (!backupInProgress) {
                toggleButtons(false);
            }
        }
    });

    // Check if Teams is open on popup load - always show button
    checkIfTeamsTab().then(async (isTeamsTab) => {
        // Always show the backup button
        backupButton.style.display = 'flex';

        if (!isTeamsTab) {
            // Informational, not an error: the popup is also opened from the Viewer's upgrade button.
            statusDiv.innerHTML = 'Please open <a href="https://teams.microsoft.com/" target="_blank">Microsoft Teams</a>, select a chat you want to backup, then try again.';
            statusDiv.style.color = "#2196f3";
            statusDiv.style.display = 'block';
        } else {
            // Teams is open - check if content script is responding
            const result = await loadAvailableChats();
            if (result.error === 'no_connection') {
                // Content script not injected - show refresh message immediately
                needsPageRefresh = true;
                statusDiv.innerHTML = 'Please <a href="#" id="refreshPageLink">refresh the page</a> to enable chat backup.';
                statusDiv.style.color = "#dc3545";
                statusDiv.style.display = "block";

                // Add click handler for refresh link
                const refreshLink = document.getElementById('refreshPageLink');
                if (refreshLink) {
                    refreshLink.addEventListener('click', (e) => {
                        e.preventDefault();
                        refreshTeamsTab();
                    });
                }
            } else {
                // Content script is working - pre-load available chats
                availableChats = result.chats;
            }
        }
    });

    // Read local activation and pending-checkout state without contacting RevenueCat.
    chrome.storage.local.get(['pendingCheckout', 'backupInProgress', BACKUP_STATE_KEY, 'purchaseVerified', 'isPro', PRO_CELEBRATION_KEY, PRO_INCLUDE_COVER_KEY], async (result) => {
        if (typeof result[PRO_CELEBRATION_KEY] === 'boolean') {
            proCelebrationAcknowledged = result[PRO_CELEBRATION_KEY];
        }

        if (typeof result[PRO_INCLUDE_COVER_KEY] === 'boolean') {
            storedProIncludeCover = result[PRO_INCLUDE_COVER_KEY];
        }

        // Check if purchase was verified by background script. Keep going afterwards
        // so an interrupted or still-running backup is also surfaced for Pro users.
        const hasVerifiedPro = result.purchaseVerified && result.isPro;
        if (hasVerifiedPro) {
            const celebrateNow = !proCelebrationAcknowledged;
            applyProState({ celebrate: celebrateNow });
            if (!celebrateNow) {
                statusDiv.style.display = 'none';
            }
        }

        // Only the background watch checks an explicitly opened checkout.
        // Opening this popup never triggers a license request.
        const checkoutAge = Date.now() - result.pendingCheckout?.startedAt;
        if (!hasVerifiedPro && result.pendingCheckout?.userId && checkoutAge >= 0 && checkoutAge < PURCHASE_WATCH_WINDOW_MS) {
            purchasePollingMessage.textContent = '⏳ Waiting for payment confirmation. If activation takes longer than 30 minutes, choose Restore Purchase.';
            purchasePollingMessage.style.display = 'block';
        }

        // Restore backup state if the popup was reopened mid-backup.
        const backupState = result[BACKUP_STATE_KEY];

        if (result.backupInProgress) {
            const tabs = await new Promise((resolve) => {
                chrome.tabs.query({ active: true, currentWindow: true }, resolve);
            });
            const activeTab = tabs[0] || null;
            const isActiveTeamsTab = !!activeTab && isTeamsUrl(activeTab.url || '');
            const activeBackupStatus = isActiveTeamsTab
                ? await getContentBackupStatus(activeTab.id)
                : null;

            if (activeBackupStatus?.inProgress) {
                monitoringDetachedBackup = true;
                backupInProgress = true;
                currentBackupChatName = backupState?.chatName || 'Unknown Chat';
                currentBackupStartTime = backupState?.startTime || Date.now();

                statusDiv.textContent = `Backup of "${currentBackupChatName}" is still running.`;
                statusDiv.style.color = "#2196f3";
                statusDiv.style.display = "block";
                loadingMessageDiv.textContent = backupState?.phase || 'Processing...';
                loadingMessageDiv.style.display = "block";
                progressContainer.style.display = "block";
                progressBar.style.width = `${Math.round(backupState?.progress || 0)}%`;
                if (backupHint) {
                    backupHint.style.display = "block";
                    updateBackupHintForBackup();
                }
                toggleButtons(true);
            } else {
                if (window.TeamsBackupDB?.markInProgressBackupsInterrupted) {
                    await window.TeamsBackupDB.markInProgressBackupsInterrupted().catch((error) => {
                        console.warn('Could not mark interrupted checkpoint backups:', error);
                    });
                }
                clearBackupState();

                let interruptedMessage = isPro
                    ? "Your previous backup was interrupted."
                    : "Backup was interrupted (popup closed).";

                if (backupState) {
                    const chatName = backupState.chatName || 'Unknown Chat';
                    const progress = Math.round(backupState.progress || 0);
                    const phase = backupState.phase || 'Unknown';
                    const elapsed = backupState.startTime ? formatTimeElapsed(Date.now() - backupState.startTime) : '';

                    interruptedMessage = `Backup of "${chatName}" was interrupted at ${progress}% (${phase}).`;
                    if (elapsed) {
                        interruptedMessage += ` It ran for ${elapsed}.`;
                    }
                }

                interruptedMessage += " Saved checkpoints can be opened in View Backed Files and continued by starting the backup again.";

                statusDiv.textContent = interruptedMessage;
                statusDiv.style.color = "#ff9800";
                statusDiv.style.display = "block";
            }
        }
    });
});
