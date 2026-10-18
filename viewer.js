// viewer.js - Chat Viewer Logic

document.addEventListener('DOMContentLoaded', async () => {
    const VIEWER_TAB_STORAGE_KEY = 'viewerTabId';

    // DOM Elements
    const chatList = document.getElementById('chatList');
    const emptyState = document.getElementById('emptyState');
    const welcomeState = document.getElementById('welcomeState');
    const chatView = document.getElementById('chatView');
    const chatAvatar = document.getElementById('chatAvatar');
    const chatName = document.getElementById('chatName');
    const chatTypeBadge = document.getElementById('chatTypeBadge');
    const chatMeta = document.getElementById('chatMeta');
    const messagesContainer = document.getElementById('messagesContainer');
    const searchInput = document.getElementById('searchChats');
    const downloadBtn = document.getElementById('downloadBtn');
    const downloadDropdown = document.getElementById('downloadDropdown');
    const downloadMenu = document.getElementById('downloadMenu');
    const downloadPdfOption = document.getElementById('downloadPdfOption');
    const downloadHtmlOption = document.getElementById('downloadHtmlOption');
    const downloadTxtOption = document.getElementById('downloadTxtOption');
    const deleteChatBtn = document.getElementById('deleteChatBtn');
    const deleteModal = document.getElementById('deleteModal');
    const deleteModalTitle = document.getElementById('deleteModalTitle');
    const deleteModalMessage = document.getElementById('deleteModalMessage');
    const cancelDeleteBtn = document.getElementById('cancelDeleteBtn');
    const confirmDeleteBtn = document.getElementById('confirmDeleteBtn');

    // Pro upgrade elements
    const proUpgradeBanner = document.getElementById('proUpgradeBanner');
    const upgradeBannerBtn = document.getElementById('upgradeBannerBtn');
    const closeBannerBtn = document.getElementById('closeBannerBtn');
    const continueBackupNotice = document.getElementById('continueBackupNotice');
    const continueBackupBtn = document.getElementById('continueBackupBtn');

    // Export All button
    const exportAllBtn = document.getElementById('exportAllBtn');
    const exportAllBtnLabel = document.getElementById('exportAllBtnLabel');

    // Import button
    const importBtn = document.getElementById('importBtn');
    const importBtnLabel = document.getElementById('importBtnLabel');
    const importFileInput = document.getElementById('importFileInput');
    const proOnlyActionsNote = document.getElementById('proOnlyActionsNote');
    const freeTierLimitNote = document.getElementById('freeTierLimitNote');

    // Search toolbar elements
    const searchToolbar = document.getElementById('searchToolbar');
    const messageSearchInput = document.getElementById('messageSearch');
    const messagesTabLabel = document.getElementById('messagesTabLabel');
    const attachmentsTabLabel = document.getElementById('attachmentsTabLabel');
    const clearSearchBtn = document.getElementById('clearSearch');
    const searchResultsDiv = document.getElementById('searchResults');
    const searchResultCount = document.getElementById('searchResultCount');
    const prevResultBtn = document.getElementById('prevResult');
    const nextResultBtn = document.getElementById('nextResult');
    const dateFilter = document.getElementById('dateFilter');

    // Edit/Delete modal elements
    const editMessageModal = document.getElementById('editMessageModal');
    const editOriginalContent = document.getElementById('editOriginalContent');
    const editMessageInput = document.getElementById('editMessageInput');
    const cancelEditBtn = document.getElementById('cancelEditBtn');
    const confirmEditBtn = document.getElementById('confirmEditBtn');
    const deleteMessageModal = document.getElementById('deleteMessageModal');
    const deleteMessagePreview = document.getElementById('deleteMessagePreview');
    const cancelDeleteMsgBtn = document.getElementById('cancelDeleteMsgBtn');
    const confirmDeleteMsgBtn = document.getElementById('confirmDeleteMsgBtn');

    // Viewer settings elements
    const viewerSettings = document.getElementById('viewerSettings');
    const viewerSettingsBtn = document.getElementById('viewerSettingsBtn');
    const closeSettingsBtn = document.getElementById('closeSettingsBtn');
    const hideDeletedCheckbox = document.getElementById('hideDeletedMessages');
    const showEditHistoryCheckbox = document.getElementById('showEditHistory');

    let allChats = [];
    let currentChatId = null;
    let currentChat = null;
    let isPro = false;
    const FREE_CHAT_LIMIT = window.TeamsBackupDB.FREE_CHAT_LIMIT;
    const FREE_PLAN_TEXT = `Free plan: ${FREE_CHAT_LIMIT} chat backup${FREE_CHAT_LIMIT === 1 ? '' : 's'}`;

    // Upgrade modal (plan limits)
    const upgradeModal = document.getElementById('upgradeModal');
    const upgradeModalTitle = document.getElementById('upgradeModalTitle');
    const upgradeModalMessage = document.getElementById('upgradeModalMessage');
    const cancelUpgradeBtn = document.getElementById('cancelUpgradeBtn');
    const confirmUpgradeBtn = document.getElementById('confirmUpgradeBtn');

    // Search state
    let currentMessages = [];
    let searchMatches = [];
    let currentMatchIndex = -1;
    let searchDebounceTimer = null;

    // Edit/Delete state
    let editingMessageId = null;
    let deletingMessageId = null;
    let viewerSettingsState = {
        hideDeleted: false,
        showEditHistory: true
    };

    // Starred filter elements and state
    const starredFilterBtn = document.getElementById('starredFilterBtn');
    const starredCount = document.getElementById('starredCount');
    let showingStarredOnly = false;

    // Tab navigation elements
    const chatTabs = document.querySelectorAll('.chat-tab');
    const attachmentsContainer = document.getElementById('attachmentsContainer');
    const attachmentsGrid = document.getElementById('attachmentsGrid');
    const noAttachments = document.getElementById('noAttachments');
    const attachmentsCountBadge = document.getElementById('attachmentsCount');
    let currentTab = 'messages';

    // Trash/Recycle bin elements
    const trashToggleBtn = document.getElementById('trashToggleBtn');
    const trashCount = document.getElementById('trashCount');
    const trashPanel = document.getElementById('trashPanel');
    const trashSection = document.querySelector('.sidebar__trash-section');
    const closeTrashBtn = document.getElementById('closeTrashBtn');
    const trashList = document.getElementById('trashList');
    const trashEmpty = document.getElementById('trashEmpty');
    const trashActionModal = document.getElementById('trashActionModal');
    const trashActionTitle = document.getElementById('trashActionTitle');
    const trashActionMessage = document.getElementById('trashActionMessage');
    const cancelTrashActionBtn = document.getElementById('cancelTrashActionBtn');
    const confirmTrashActionBtn = document.getElementById('confirmTrashActionBtn');
    let trashActionChatId = null;
    let trashActionType = null; // 'restore' or 'delete'

    function registerViewerTab() {
        if (!chrome?.tabs?.getCurrent || !chrome?.storage?.local) return;
        chrome.tabs.getCurrent((tab) => {
            if (chrome.runtime.lastError || !tab?.id) {
                return;
            }
            chrome.storage.local.set({ [VIEWER_TAB_STORAGE_KEY]: tab.id });
        });
    }

    registerViewerTab();

    // Check Pro status
    async function checkProStatus() {
        return new Promise((resolve) => {
            chrome.storage.local.get(['isPro', 'purchaseVerified'], (result) => {
                isPro = result.isPro === true || result.purchaseVerified === true;
                resolve(isPro);
            });
        });
    }

    function getAccessibleChatIds() {
        if (isPro) {
            return new Set(allChats.map(chat => chat.id));
        }
        return new Set(allChats.slice(0, FREE_CHAT_LIMIT).map(chat => chat.id));
    }

    function getLockedChatCount() {
        if (isPro) return 0;
        return Math.max(0, allChats.length - FREE_CHAT_LIMIT);
    }

    function getBackupKind(chat) {
        if (!chat) return 'chat';
        if (chat.sourceKind) return chat.sourceKind;
        if (chat.isChannel) return 'channel';
        return 'chat';
    }

    function getBackupKindLabel(chat) {
        switch (getBackupKind(chat)) {
            case 'channel_chat':
                return 'Channel chat';
            case 'channel_posts':
                return 'Channel thread';
            case 'channel':
                return 'Channel';
            case 'chat':
            default:
                return 'Chat';
        }
    }

    function getBackupKindClass(chat) {
        return `chat-kind-badge--${getBackupKind(chat)}`;
    }

    function getMessageContentLines(content) {
        return (content || '')
            .replace(/\r\n/g, '\n')
            .split('\n')
            .map(part => part.trim())
            .filter(Boolean);
    }

    function looksLikeChannelThreadMessage(msg) {
        if (!msg || msg.type === 'system') return false;

        const lines = getMessageContentLines(msg.content);
        if (lines.length < 2) return false;

        const title = lines[0];
        const body = lines.slice(1).join(' ');
        if (title.length < 8 || title.length > 140) return false;
        if (body.length < 2) return false;

        const normalizedTitle = title.toLowerCase();
        if (/^(reply in thread|post in channel|see more|view more)$/i.test(normalizedTitle)) return false;

        return normalizedTitle !== body.toLowerCase();
    }

    function resolveBackupKindFromMessages(chat, messages = []) {
        const storedKind = getBackupKind(chat);
        if (!chat?.isChannel) return storedKind;
        if (storedKind === 'channel_posts' || storedKind === 'channel_chat') return storedKind;

        const contentMessages = messages.filter(msg => msg && msg.type !== 'system');
        if (contentMessages.length === 0) return storedKind;

        const threadLikeCount = contentMessages.filter(looksLikeChannelThreadMessage).length;
        const threshold = contentMessages.length === 1 ? 1 : Math.max(2, Math.ceil(contentMessages.length * 0.6));

        if (threadLikeCount >= threshold) {
            return 'channel_posts';
        }

        return storedKind;
    }

    function isChannelPostsBackup(chat = currentChat) {
        return getBackupKind(chat) === 'channel_posts';
    }

    function isChannelChatBackup(chat = currentChat) {
        return getBackupKind(chat) === 'channel_chat';
    }

    function updateChatViewMode(chat) {
        const kind = getBackupKind(chat);
        chatView.classList.remove('chat-view--kind-chat', 'chat-view--kind-channel-chat', 'chat-view--kind-channel-posts');
        messagesContainer.classList.remove('messages-container--kind-chat', 'messages-container--kind-channel-chat', 'messages-container--kind-channel-posts');

        if (kind === 'channel_posts') {
            chatView.classList.add('chat-view--kind-channel-posts');
            messagesContainer.classList.add('messages-container--kind-channel-posts');
        } else if (kind === 'channel_chat') {
            chatView.classList.add('chat-view--kind-channel-chat');
            messagesContainer.classList.add('messages-container--kind-channel-chat');
        } else {
            chatView.classList.add('chat-view--kind-chat');
            messagesContainer.classList.add('messages-container--kind-chat');
        }

        if (messagesTabLabel) messagesTabLabel.textContent = 'Messages';
        if (attachmentsTabLabel) attachmentsTabLabel.textContent = 'Attachments';
        if (messageSearchInput) messageSearchInput.placeholder = 'Search messages...';
    }

    function isChatLockedForFree(chatId) {
        if (isPro || !chatId) return false;
        return !getAccessibleChatIds().has(chatId);
    }

    function updateFreeTierLimitNote() {
        if (!freeTierLimitNote) return;

        if (isPro) {
            freeTierLimitNote.style.display = 'none';
            return;
        }

        const lockedCount = getLockedChatCount();
        freeTierLimitNote.textContent = lockedCount > 0
            ? `${FREE_PLAN_TEXT}. ${lockedCount} older backup${lockedCount === 1 ? ' is' : 's are'} locked until you upgrade to Pro.`
            : `${FREE_PLAN_TEXT}. Downloads include a FREE watermark.`;
        freeTierLimitNote.style.display = 'block';
    }

    function enforceVisibleChatAccessLimit() {
        if (!isChatLockedForFree(currentChatId)) return;

        currentChatId = null;
        currentChat = null;
        chatView.style.display = 'none';
        welcomeState.style.display = 'flex';
    }

    function updateProOnlyActionButtons() {
        // Free users can click these to see what Pro unlocks.
        exportAllBtn.disabled = false;
        importBtn.disabled = false;
        exportAllBtn.title = isPro ? 'Export all backups as ZIP' : 'Pro users only: export all backups as ZIP';
        importBtn.title = isPro ? 'Import backups from file' : 'Pro users only: import backups from file';

        if (exportAllBtnLabel) exportAllBtnLabel.textContent = isPro ? 'Export All' : 'Export All (Pro)';
        if (importBtnLabel) importBtnLabel.textContent = isPro ? 'Import' : 'Import (Pro)';
        if (proOnlyActionsNote) {
            proOnlyActionsNote.textContent = 'Bulk import/export all is Pro. Select an unlocked backup to download PDF, HTML, or TXT.';
            proOnlyActionsNote.style.display = isPro ? 'none' : 'block';
        }
        updateFreeTierLimitNote();
    }

    function updateDeleteModalForPlan() {
        if (!deleteModalTitle || !deleteModalMessage || !confirmDeleteBtn) return;

        if (isPro) {
            deleteModalTitle.textContent = 'Move to Recycle Bin?';
            deleteModalMessage.textContent = 'This backup will be moved to the recycle bin and automatically deleted after 60 days. You can restore it anytime before then.';
            confirmDeleteBtn.textContent = 'Move to Bin';
        } else {
            deleteModalTitle.textContent = 'Move to Recycle Bin?';
            deleteModalMessage.textContent = `This backup moves to the Recycle Bin for 60 days. Deleting it does not free your free-plan slot for a different chat, but you can restore this backup later.`;
            confirmDeleteBtn.textContent = 'Move to Bin';
        }
    }

    // Show/hide Pro upgrade banner and update Pro-related UI
    async function updateProBanner() {
        await checkProStatus();
        updateProOnlyActionButtons();
        updateDeleteModalForPlan();
        if (!isPro) {
            // Check if banner was dismissed in this session
            const dismissed = sessionStorage.getItem('proBannerDismissed');
            if (!dismissed) {
                proUpgradeBanner.style.display = 'block';
            }
        } else {
            proUpgradeBanner.style.display = 'none';
            // Update Export All button tooltip for Pro users
            exportAllBtn.title = 'Export all backups as ZIP';
        }
    }

    // Opens the hosted checkout directly; the background worker unlocks Pro
    // (and this page) as soon as the payment is confirmed.
    async function startUpgrade() {
        try {
            const result = await window.TeamsBackupLicense.openCheckout();
            if (result.alreadyPro) await updateProBanner();
        } catch (error) {
            console.error('Unable to open checkout:', error);
            alert('Could not open checkout. Open Teams Chat Backup from the Chrome toolbar and choose Unlock Full Version.');
        }
    }

    function showUpgradeModal(message, title = 'Upgrade to Pro') {
        if (!upgradeModal) return;
        upgradeModalTitle.textContent = title;
        upgradeModalMessage.textContent = message;
        upgradeModal.style.display = 'flex';
    }

    function hideUpgradeModal() {
        if (upgradeModal) upgradeModal.style.display = 'none';
    }

    const LOCKED_CHAT_MESSAGE = `${FREE_PLAN_TEXT}. Upgrade to Pro to open your other backups.`;

    cancelUpgradeBtn?.addEventListener('click', hideUpgradeModal);
    upgradeModal?.querySelector('.modal__backdrop')?.addEventListener('click', hideUpgradeModal);
    confirmUpgradeBtn?.addEventListener('click', () => {
        hideUpgradeModal();
        startUpgrade();
    });

    // Banner button handlers
    upgradeBannerBtn.addEventListener('click', startUpgrade);

    closeBannerBtn.addEventListener('click', () => {
        proUpgradeBanner.style.display = 'none';
        sessionStorage.setItem('proBannerDismissed', 'true');
    });

    continueBackupBtn.addEventListener('click', () => {
        window.open('https://teams.microsoft.com/', '_blank');
    });

    // Initialize
    await updateProBanner();
    await reconcileInterruptedBackups();
    await loadChats();

    chrome.storage.onChanged.addListener((changes, areaName) => {
        if (areaName === 'local' && (changes.isPro || changes.purchaseVerified)) {
            // Re-render the list too, so chats locked on the free plan unlock right away.
            updateProBanner().then(() => loadChats()).catch(error => {
                console.error('Error refreshing Pro status:', error);
            });
        }

        if (areaName === 'local' && changes.backupCheckpointUpdated) {
            scheduleCheckpointRefresh(changes.backupCheckpointUpdated.newValue);
        }
    });

    // Checkpoints land every few seconds during a backup. Re-render at most every
    // 10 seconds (immediately when the backup finishes) instead of on each one.
    const CHECKPOINT_REFRESH_INTERVAL_MS = 10000;
    let checkpointRefreshTimer = null;
    let pendingCheckpointUpdate = null;
    let lastCheckpointRefreshAt = 0;

    function scheduleCheckpointRefresh(update = {}) {
        pendingCheckpointUpdate = update;
        const isFinal = update?.reason === 'finalize';
        if (checkpointRefreshTimer && !isFinal) return;
        clearTimeout(checkpointRefreshTimer);

        const delayMs = isFinal ? 0 : Math.max(0, lastCheckpointRefreshAt + CHECKPOINT_REFRESH_INTERVAL_MS - Date.now());
        checkpointRefreshTimer = setTimeout(() => {
            checkpointRefreshTimer = null;
            lastCheckpointRefreshAt = Date.now();
            const nextUpdate = pendingCheckpointUpdate;
            pendingCheckpointUpdate = null;
            refreshCheckpointView(nextUpdate).catch(error => {
                console.error('Error refreshing checkpoint view:', error);
            });
        }, delayMs);
    }

    // Shown after the background worker confirms a purchase.
    if (new URLSearchParams(window.location.search).get('pro') === 'activated') {
        const notice = document.createElement('div');
        notice.className = 'pro-activated-notice';
        notice.setAttribute('role', 'status');
        notice.textContent = 'Pro is active. Thank you for your purchase! All backups are unlocked.';
        document.body.appendChild(notice);
        setTimeout(() => notice.remove(), 8000);
    }

    // Initialize trash: purge expired trash and update count
    await initializeTrash();

    // Check if a specific chat was requested via URL
    const urlParams = new URLSearchParams(window.location.search);
    const requestedChatId = urlParams.get('chat');
    if (requestedChatId) {
        await openChat(requestedChatId);
    }

    // Repair stale checkpoint metadata from interrupted runs before rendering.
    async function reconcileInterruptedBackups() {
        if (!window.TeamsBackupDB?.markInProgressBackupsInterrupted) return;

        const state = await new Promise((resolve) => {
            chrome.storage.local.get(['backupInProgress'], resolve);
        });

        if (!state?.backupInProgress) {
            await window.TeamsBackupDB.markInProgressBackupsInterrupted();
        }
    }

    // Load all chats into sidebar
    async function loadChats() {
        try {
            allChats = await window.TeamsBackupDB.getChats();
            enforceVisibleChatAccessLimit();
            updateFreeTierLimitNote();

            if (allChats.length === 0) {
                chatList.style.display = 'none';
                emptyState.style.display = 'flex';
                return;
            }

            chatList.style.display = 'block';
            emptyState.style.display = 'none';
            renderChatList(allChats);

        } catch (error) {
            console.error('Error loading chats:', error);
            chatList.innerHTML = '<div style="color: #ff6b6b; padding: 20px; text-align: center;">Error loading backups</div>';
        }
    }

    async function refreshCheckpointView(update = {}) {
        const updatedChatId = update?.chatId || null;
        await loadChats();

        if (currentChatId && (!updatedChatId || currentChatId === updatedChatId)) {
            await openChat(currentChatId);
        }
    }

    function getChatListStatusLabel(chat) {
        if (chat?.inProgress) {
            return ' · Saving...';
        }
        if (chat?.isPreview) {
            return ' · Incomplete';
        }
        return '';
    }

    // Render chat list
    function renderChatList(chats) {
        const accessibleChatIds = getAccessibleChatIds();
        chatList.innerHTML = chats.map(chat => `
            <div class="chat-item ${chat.id === currentChatId ? 'active' : ''} ${!isPro && !accessibleChatIds.has(chat.id) ? 'chat-item--locked' : ''}" data-id="${chat.id}" data-locked="${!isPro && !accessibleChatIds.has(chat.id) ? 'true' : 'false'}">
                <div class="chat-item__avatar" style="background: ${window.TeamsBackupDB.stringToColor(chat.name)}">
                    ${window.TeamsBackupDB.getInitials(chat.name)}
                </div>
                <div class="chat-item__info">
                    <div class="chat-item__name-row">
                        <div class="chat-item__name">${escapeHtml(chat.name)}</div>
                        <span class="chat-kind-badge ${getBackupKindClass(chat)}">${escapeHtml(getBackupKindLabel(chat))}</span>
                    </div>
                    <div class="chat-item__meta">${formatDate(chat.createdAt)} · ${chat.messageCount} msgs${getChatListStatusLabel(chat)}</div>
                </div>
                ${!isPro && !accessibleChatIds.has(chat.id) ? '<div class="chat-item__lock-badge">Locked · Pro</div>' : ''}
            </div>
        `).join('');

        // Add click handlers
        chatList.querySelectorAll('.chat-item').forEach(item => {
            item.addEventListener('click', () => {
                if (item.dataset.locked === 'true') {
                    showUpgradeModal(LOCKED_CHAT_MESSAGE);
                    return;
                }
                const chatId = item.dataset.id;
                openChat(chatId);
            });
        });
    }

    // Open a chat
    async function openChat(chatId) {
        if (isChatLockedForFree(chatId)) {
            showUpgradeModal(LOCKED_CHAT_MESSAGE);
            return;
        }

        try {
            currentChatId = chatId;
            currentChat = await window.TeamsBackupDB.getChat(chatId);

            if (!currentChat) {
                console.error('Chat not found:', chatId);
                return;
            }

            // Update sidebar active state
            chatList.querySelectorAll('.chat-item').forEach(item => {
                item.classList.toggle('active', item.dataset.id === chatId);
            });

            // Show chat view
            welcomeState.style.display = 'none';
            chatView.style.display = 'flex';

            // Update header
            chatAvatar.textContent = window.TeamsBackupDB.getInitials(currentChat.name);
            chatAvatar.style.background = window.TeamsBackupDB.stringToColor(currentChat.name);
            chatName.textContent = currentChat.name;
            if (chatTypeBadge) {
                chatTypeBadge.textContent = getBackupKindLabel(currentChat);
                chatTypeBadge.className = `chat-kind-badge ${getBackupKindClass(currentChat)}`;
            }
            updateChatViewMode(currentChat);

            // Show checkpoint/incomplete state clearly.
            if (currentChat.inProgress) {
                chatMeta.textContent = `${currentChat.messageCount} messages · Saving checkpoint`;
                continueBackupNotice.style.display = 'block';
                const noticeTitle = continueBackupNotice.querySelector('strong');
                const noticeBody = continueBackupNotice.querySelector('span');
                if (noticeTitle) noticeTitle.textContent = 'Backup still running';
                if (noticeBody) noticeBody.textContent = 'Saved checkpoints will continue appearing here while Teams remains open.';
                if (continueBackupBtn) continueBackupBtn.style.display = 'none';
            } else if (currentChat.isPreview) {
                chatMeta.textContent = `${currentChat.messageCount} messages · Incomplete backup`;
                continueBackupNotice.style.display = 'block';
                const noticeTitle = continueBackupNotice.querySelector('strong');
                const noticeBody = continueBackupNotice.querySelector('span');
                if (noticeTitle) noticeTitle.textContent = 'This backup is incomplete';
                if (noticeBody) noticeBody.textContent = 'Open Teams and run this chat again to continue from saved checkpoints.';
                if (continueBackupBtn) continueBackupBtn.style.display = '';
            } else {
                chatMeta.textContent = `${currentChat.messageCount} messages · Backed up ${formatDate(currentChat.createdAt)}`;
                continueBackupNotice.style.display = 'none';
                if (continueBackupBtn) continueBackupBtn.style.display = '';
            }

            // Reset starred filter state when switching chats
            showingStarredOnly = false;
            if (starredFilterBtn) {
                starredFilterBtn.classList.remove('active');
            }

            // Reset to messages tab when switching chats
            currentTab = 'messages';
            chatTabs.forEach(tab => {
                tab.classList.toggle('chat-tab--active', tab.dataset.tab === 'messages');
            });
            messagesContainer.style.display = 'block';
            attachmentsContainer.style.display = 'none';
            searchToolbar.style.display = 'flex';

            // Load messages
            await loadMessages(chatId);

        } catch (error) {
            console.error('Error opening chat:', error);
        }
    }

    function escapeHtmlWithBreaks(text) {
        return escapeHtml(text).replace(/\n/g, '<br>');
    }

    function splitPostMessageContent(content) {
        const normalized = (content || '')
            .replace(/\r\n/g, '\n')
            .split('\n')
            .map(part => part.trim())
            .filter(Boolean);

        if (normalized.length >= 2) {
            return {
                title: normalized[0],
                body: normalized.slice(1).join('\n')
            };
        }

        if (normalized.length === 1) {
            return {
                title: '',
                body: normalized[0]
            };
        }

        return { title: '', body: '' };
    }

    // Render incomplete backup notice for inside chat history
    function renderIncompleteNotice(position) {
        // Don't show notices for Pro users or if chat is not marked as preview
        if (isPro) {
            return '';
        }
        const isTop = position === 'top';
        return `
            <div class="incomplete-backup-notice" data-position="${position}">
                <div class="incomplete-backup-notice__icon">◇</div>
                <div class="incomplete-backup-notice__text">
                    <strong>${isTop ? 'Start of Saved Checkpoint' : 'End of Saved Checkpoint'}</strong>
                    ${isTop ? 'This backup was interrupted or stopped early. Open Teams and run it again to continue from the saved messages.' : 'Individual PDF, HTML, and TXT downloads are available from the Download button above.'}
                </div>
                <button class="incomplete-backup-notice__btn upgrade-from-chat">Open Teams</button>
            </div>
        `;
    }

    // Load and render messages
    async function loadMessages(chatId) {
        try {
            const messages = await window.TeamsBackupDB.getMessages(chatId);

            if (currentChat) {
                const resolvedKind = resolveBackupKindFromMessages(currentChat, messages);
                if (resolvedKind !== getBackupKind(currentChat)) {
                    currentChat = {
                        ...currentChat,
                        sourceKind: resolvedKind
                    };

                    const chatIndex = allChats.findIndex(chat => chat.id === chatId);
                    if (chatIndex !== -1) {
                        allChats[chatIndex] = {
                            ...allChats[chatIndex],
                            sourceKind: resolvedKind
                        };
                        renderChatList(allChats);
                    }

                    if (chatTypeBadge) {
                        chatTypeBadge.textContent = getBackupKindLabel(currentChat);
                        chatTypeBadge.className = `chat-kind-badge ${getBackupKindClass(currentChat)}`;
                    }

                    updateChatViewMode(currentChat);
                }
            }

            // Store messages for search functionality
            currentMessages = messages;

            // Clear any existing search
            clearSearch();

            if (messages.length === 0) {
                messagesContainer.innerHTML = `
                    <div style="text-align: center; padding: 40px; color: #888;">
                        <p>No messages found in this backup</p>
                        <p style="font-size: 13px; margin-top: 8px;">The text content may not have been extracted during backup</p>
                    </div>
                `;
                populateDateFilter([]);
                return;
            }

            // Group messages by date
            const groupedMessages = groupMessagesByDate(messages);
            const messageIndexes = new Map(messages.map((message, index) => [message, index]));
            let html = '';

            // Add incomplete notice at the beginning if this is a preview backup
            if (currentChat && currentChat.isPreview) {
                html += renderIncompleteNotice('top');
            }

            for (const [, dateInfo] of Object.entries(groupedMessages)) {
                // Add data-date attribute for date filtering
                html += `<div class="date-divider" data-date="${dateInfo.dateKey}">${dateInfo.label}</div>`;

                for (const msg of dateInfo.messages) {
                    const originalIndex = messageIndexes.get(msg) ?? -1;
                    if (msg.type === 'system') {
                        html += renderSystemMessage(msg, originalIndex);
                    } else {
                        html += renderMessage(msg, originalIndex);
                    }
                }
            }

            // Add incomplete notice at the end if this is a preview backup
            if (currentChat && currentChat.isPreview) {
                html += renderIncompleteNotice('bottom');
            }

            messagesContainer.innerHTML = html;

            // Populate date filter dropdown
            populateDateFilter(messages);

            // Add click handlers for incomplete-checkpoint action buttons inside chat
            messagesContainer.querySelectorAll('.upgrade-from-chat').forEach(btn => {
                btn.addEventListener('click', () => {
                    window.open('https://teams.microsoft.com/', '_blank');
                });
            });

            // Scroll to bottom
            messagesContainer.scrollTop = messagesContainer.scrollHeight;

            // Update starred count badge
            updateStarredCount();

            // Update attachments count badge
            updateAttachmentsCount();

        } catch (error) {
            console.error('Error loading messages:', error);
            messagesContainer.textContent = '';
            const errorDiv = document.createElement('div');
            errorDiv.style.cssText = 'color: #ff6b6b; padding: 20px; text-align: center;';
            errorDiv.textContent = 'Error loading messages';
            messagesContainer.appendChild(errorDiv);
        }
    }

    // Group messages by date
    function groupMessagesByDate(messages) {
        const groups = {};
        const dateKeyToTimestamp = {}; // Map local date key to actual timestamp for sorting

        messages.forEach(msg => {
            const date = parseTimestamp(msg.timestamp);
            const displayLabel = date ? formatDateHeader(date) : 'Unknown Date';
            const localDateKey = date ? getLocalDateKey(date) : 'unknown';

            if (!groups[localDateKey]) {
                groups[localDateKey] = {
                    label: displayLabel,
                    messages: [],
                    dateKey: localDateKey
                };
                // Store the timestamp for sorting - use 0 for Unknown Date so it appears first
                dateKeyToTimestamp[localDateKey] = date ? date.getTime() : 0;
            }
            groups[localDateKey].messages.push(msg);
        });

        // Sort groups by date (oldest first) using stable local-day keys.
        const sortedEntries = Object.entries(groups).sort((a, b) => {
            const timeA = dateKeyToTimestamp[a[0]] || 0;
            const timeB = dateKeyToTimestamp[b[0]] || 0;
            return timeA - timeB; // Oldest first
        });

        // Convert back to object with sorted insertion order
        const sortedGroups = {};
        sortedEntries.forEach(([key, value]) => {
            sortedGroups[key] = value;
        });

        return sortedGroups;
    }

    // Parse timestamp string to Date
    function parseTimestamp(timestamp) {
        if (!timestamp) return null;

        // Try direct parsing (works for ISO format like "2025-12-11T07:57:10.436Z")
        const date = new Date(timestamp);
        if (!isNaN(date.getTime())) {
            return date;
        }

        // Handle "Today at X:XX AM/PM" format
        const todayMatch = timestamp.match(/Today\s+at\s+(\d{1,2}):(\d{2})\s*(AM|PM)?/i);
        if (todayMatch) {
            const today = new Date();
            let hours = parseInt(todayMatch[1], 10);
            const minutes = parseInt(todayMatch[2], 10);
            const period = todayMatch[3];
            if (period) {
                if (period.toUpperCase() === 'PM' && hours < 12) hours += 12;
                if (period.toUpperCase() === 'AM' && hours === 12) hours = 0;
            }
            today.setHours(hours, minutes, 0, 0);
            return today;
        }

        // Handle "Yesterday at X:XX AM/PM" format
        const yesterdayMatch = timestamp.match(/Yesterday\s+at\s+(\d{1,2}):(\d{2})\s*(AM|PM)?/i);
        if (yesterdayMatch) {
            const yesterday = new Date();
            yesterday.setDate(yesterday.getDate() - 1);
            let hours = parseInt(yesterdayMatch[1], 10);
            const minutes = parseInt(yesterdayMatch[2], 10);
            const period = yesterdayMatch[3];
            if (period) {
                if (period.toUpperCase() === 'PM' && hours < 12) hours += 12;
                if (period.toUpperCase() === 'AM' && hours === 12) hours = 0;
            }
            yesterday.setHours(hours, minutes, 0, 0);
            return yesterday;
        }

        // Try common formats
        const patterns = [
            /(\d{1,2})\/(\d{1,2})\/(\d{4}),?\s*(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?/i,
            /(\d{4})-(\d{2})-(\d{2})T?(\d{2}):(\d{2})(?::(\d{2}))?/
        ];

        for (const pattern of patterns) {
            const match = timestamp.match(pattern);
            if (match) {
                return new Date(timestamp);
            }
        }

        return null;
    }

    // Format date for header
    function formatDateHeader(date) {
        const today = new Date();
        const yesterday = new Date(today);
        yesterday.setDate(yesterday.getDate() - 1);

        if (date.toDateString() === today.toDateString()) {
            return 'Today';
        } else if (date.toDateString() === yesterday.toDateString()) {
            return 'Yesterday';
        } else {
            return date.toLocaleDateString(undefined, {
                weekday: 'long',
                year: 'numeric',
                month: 'long',
                day: 'numeric'
            });
        }
    }

    function getLocalDateKey(date) {
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const day = String(date.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
    }

    function isEmbeddedImageDataUrl(value) {
        return typeof value === 'string' &&
            /^data:image\/(?:png|jpe?g|gif|webp|bmp|avif|x-icon|vnd\.microsoft\.icon);base64,[a-z0-9+/]+={0,2}$/i.test(value);
    }

    // Animated Teams emojis are stored as vertical sprite strips (square frames
    // stacked top-to-bottom). Render them as a CSS steps() animation so they
    // play instead of showing the unrolled strip.
    function isAnimatedEmojiImage(img) {
        return !!(img && img.isAnimatedEmoji && img.frameCount > 1 && isEmbeddedImageDataUrl(img.dataUrl));
    }

    function renderAnimatedEmojiSprite(img) {
        if (!isEmbeddedImageDataUrl(img?.dataUrl)) return '';
        const frameWidth = img.frameWidth || img.width || 50;
        const frameHeight = img.frameHeight || frameWidth;
        const frameCount = img.frameCount;
        const displayW = Math.min(frameWidth, 64);
        const displayH = Math.round(frameHeight * (displayW / frameWidth));
        const travel = displayH * frameCount;
        const duration = Math.max(1, frameCount / 12).toFixed(2);
        const label = escapeHtml(img.alt || 'Animated emoji');
        return `<span class="emoji-sprite" role="img" aria-label="${label}" title="${label}" ` +
            `style="width:${displayW}px;height:${displayH}px;background-image:url('${img.dataUrl}');` +
            `background-size:${displayW}px auto;--emoji-sprite-travel:-${travel}px;` +
            `animation:emoji-sprite-play ${duration}s steps(${frameCount}) infinite;"></span>`;
    }

    // Render images in a message
    function renderMessageImages(images) {
        if (!images || images.length === 0) return '';

        const imagesHtml = images.map(img => {
            if (isAnimatedEmojiImage(img)) {
                return `
            <div class="message__image-wrapper message__image-wrapper--emoji">
                ${renderAnimatedEmojiSprite(img)}
            </div>
        `;
            }
            return `
            <div class="message__image-wrapper">
                ${isEmbeddedImageDataUrl(img.dataUrl) ? `
                <img
                    src="${img.dataUrl}"
                    alt="${escapeHtml(img.alt || 'Image')}"
                    class="message__image"
                    loading="lazy"
                    onclick="window.open('${img.dataUrl}', '_blank')"
                />
                ` : `
                <div class="message__image message__image--placeholder">
                    <span class="message__image-label">${escapeHtml(img.alt || 'Image attachment')}</span>
                    <span class="message__image-meta">Preview unavailable in checkpoint save</span>
                </div>
                `}
            </div>
        `;
        }).join('');

        return `<div class="message__images">${imagesHtml}</div>`;
    }

    // Render file attachments for a message
    // Get file type category for styling (also used in attachment cards)
    function getFileTypeCategoryInline(extension) {
        const ext = (extension || '').toLowerCase();
        const typeMap = {
            excel: ['xlsx', 'xls', 'xlsm', 'xlsb', 'csv'],
            word: ['docx', 'doc', 'docm', 'rtf'],
            powerpoint: ['pptx', 'ppt', 'pptm'],
            pdf: ['pdf'],
            image: ['jpg', 'jpeg', 'png', 'gif', 'webp', 'bmp', 'svg', 'ico', 'tiff'],
            archive: ['zip', 'rar', '7z', 'tar', 'gz', 'bz2'],
            '3d': ['dae', 'glb', 'gltf', 'obj', 'fbx', 'stl', 'blend', '3ds'],
            code: ['js', 'ts', 'jsx', 'tsx', 'py', 'java', 'cpp', 'c', 'h', 'cs', 'php', 'rb', 'go', 'rs', 'swift', 'kt', 'html', 'css', 'scss', 'json', 'xml', 'yaml', 'yml', 'sql', 'sh', 'bash', 'md'],
            audio: ['mp3', 'wav', 'ogg', 'flac', 'aac', 'm4a', 'wma'],
            video: ['mp4', 'mov', 'avi', 'mkv', 'webm', 'wmv', 'flv', 'm4v']
        };
        for (const [category, extensions] of Object.entries(typeMap)) {
            if (extensions.includes(ext)) return category;
        }
        return 'default';
    }

    function renderMessageAttachments(attachments) {
        if (!attachments || attachments.length === 0) return '';

        const attachmentsHtml = attachments.map(att => {
            const filename = escapeHtml(att.filename || 'Unknown file');
            const sizeText = att.fileSizeText ? escapeHtml(att.fileSizeText) : '';
            const ext = (att.extension || '').toLowerCase();
            const extDisplay = ext.toUpperCase() || 'FILE';
            const fileType = getFileTypeCategoryInline(ext);

            // Determine if file has data or is just a reference
            const hasData = att.dataUrl && !att.downloadError;
            const isTooLarge = att.tooLarge;
            const hasError = att.downloadError;

            let statusClass = '';
            let statusText = '';

            if (hasData) {
                statusClass = 'attachment--available';
            } else if (isTooLarge) {
                statusClass = 'attachment--too-large';
                statusText = `<span class="attachment__status">File exceeds size limit</span>`;
            } else if (hasError) {
                statusClass = 'attachment--error';
                statusText = `<span class="attachment__status">Could not download</span>`;
            } else {
                statusClass = 'attachment--reference';
            }

            // Create action buttons
            let actionButtons = '';
            if (hasData) {
                actionButtons = `
                    <a href="${att.dataUrl}" download="${filename}" class="attachment__download" title="Download ${filename}">
                        <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                            <polyline points="7 10 12 15 17 10"/>
                            <line x1="12" y1="15" x2="12" y2="3"/>
                        </svg>
                        Download
                    </a>
                `;
                if (att.sharePointUrl) {
                    actionButtons += `
                        <a href="${escapeHtml(att.sharePointUrl)}" target="_blank" rel="noopener noreferrer" class="attachment__link" title="Open in SharePoint">
                            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                                <path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/>
                                <polyline points="15 3 21 3 21 9"/>
                                <line x1="10" y1="14" x2="21" y2="3"/>
                            </svg>
                            Open
                        </a>
                    `;
                }
            } else if (att.sharePointUrl) {
                actionButtons = `
                    <a href="${escapeHtml(att.sharePointUrl)}" target="_blank" rel="noopener noreferrer" class="attachment__link" title="Open in SharePoint">
                        <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                            <path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/>
                            <polyline points="15 3 21 3 21 9"/>
                            <line x1="10" y1="14" x2="21" y2="3"/>
                        </svg>
                        Open in SharePoint
                    </a>
                `;
            }

            return `
                <div class="attachment ${statusClass} attachment--${fileType}">
                    <div class="attachment__info">
                        <span class="attachment__ext attachment__ext--${fileType}">${extDisplay}</span>
                        <span class="attachment__name">${filename}</span>
                        ${sizeText ? `<span class="attachment__size">${sizeText}</span>` : ''}
                    </div>
                    ${statusText}
                    <div class="attachment__actions">
                        ${actionButtons}
                    </div>
                </div>
            `;
        }).join('');

        return `<div class="message__attachments">${attachmentsHtml}</div>`;
    }

    // Render a message
    function renderMessage(msg, messageIndex = -1) {
        if (isChannelPostsBackup()) {
            return renderChannelPostMessage(msg, messageIndex);
        }

        const time = extractTime(msg.timestamp);
        const imagesHtml = renderMessageImages(msg.images);
        const attachmentsHtml = renderMessageAttachments(msg.attachments);
        const hasText = msg.content && msg.content.trim();
        const dataAttr = messageIndex >= 0 ? `data-message-index="${messageIndex}"` : '';
        const msgIdAttr = msg.id ? `data-message-id="${msg.id}"` : '';

        // Handle deleted messages
        if (msg.isDeleted) {
            return renderDeletedMessage(msg, messageIndex);
        }

        // Check if message was edited
        const isEdited = msg.isEdited || (msg.editHistory && msg.editHistory.length > 0);
        const editedClass = isEdited ? 'message--edited' : '';
        const editIndicator = isEdited ? `<span class="message__edit-indicator">edited</span>` : '';

        // Check if message is starred
        const isStarred = msg.isStarred;
        const starredClass = isStarred ? 'message--starred' : '';
        const starIndicator = isStarred ? `
            <span class="message__star-indicator">
                <svg viewBox="0 0 24 24" fill="currentColor" stroke="currentColor" stroke-width="2">
                    <polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"></polygon>
                </svg>
            </span>
        ` : '';

        // Generate diff popup for edited messages
        const diffPopup = isEdited ? generateEditDiffPopup(msg) : '';

        // Action buttons (star/edit/delete)
        const actionButtons = `
            <div class="message__actions">
                <button class="message__action-btn message__action-btn--star ${isStarred ? 'starred' : ''}" title="${isStarred ? 'Unstar message' : 'Star message'}" data-action="star">
                    <svg viewBox="0 0 24 24" fill="${isStarred ? 'currentColor' : 'none'}" stroke="currentColor" stroke-width="2">
                        <polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"></polygon>
                    </svg>
                </button>
                <button class="message__action-btn message__action-btn--edit" title="Edit message" data-action="edit">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"></path>
                        <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"></path>
                    </svg>
                </button>
                <button class="message__action-btn message__action-btn--delete" title="Delete message" data-action="delete">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <polyline points="3 6 5 6 21 6"></polyline>
                        <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
                    </svg>
                </button>
            </div>
        `;

        const channelChatClass = isChannelChatBackup() ? 'message--channel-chat' : '';

        return `
            <div class="message ${channelChatClass} ${editedClass} ${starredClass}" ${dataAttr} ${msgIdAttr}>
                <div class="message__avatar" style="background: ${msg.avatarColor || window.TeamsBackupDB.stringToColor(msg.author)}">
                    ${msg.initials || window.TeamsBackupDB.getInitials(msg.author)}
                </div>
                <div class="message__content">
                    <div class="message__header">
                        <span class="message__author">${escapeHtml(msg.author)}</span>
                        <span class="message__time">${time}</span>
                        ${starIndicator}
                        ${editIndicator}
                    </div>
                    ${hasText ? `<div class="message__text">${escapeHtml(msg.content)}</div>` : ''}
                    ${imagesHtml}
                    ${attachmentsHtml}
                    ${diffPopup}
                </div>
                ${actionButtons}
            </div>
        `;
    }

    function renderChannelPostMessage(msg, messageIndex = -1) {
        const time = extractTime(msg.timestamp);
        const imagesHtml = renderMessageImages(msg.images);
        const attachmentsHtml = renderMessageAttachments(msg.attachments);
        const dataAttr = messageIndex >= 0 ? `data-message-index="${messageIndex}"` : '';
        const msgIdAttr = msg.id ? `data-message-id="${msg.id}"` : '';
        const hasText = msg.content && msg.content.trim();
        const postContent = splitPostMessageContent(msg.content || '');
        const titleHtml = postContent.title ? `<div class="message__post-title">${escapeHtml(postContent.title)}</div>` : '';
        const bodyHtml = postContent.body
            ? `<div class="message__text message__text--channel-post">${escapeHtmlWithBreaks(postContent.body)}</div>`
            : (!postContent.title && hasText ? `<div class="message__text message__text--channel-post">${escapeHtmlWithBreaks(msg.content)}</div>` : '');

        if (msg.isDeleted) {
            return renderDeletedMessage(msg, messageIndex);
        }

        const isEdited = msg.isEdited || (msg.editHistory && msg.editHistory.length > 0);
        const editedClass = isEdited ? 'message--edited' : '';
        const editIndicator = isEdited ? `<span class="message__edit-indicator">edited</span>` : '';
        const isStarred = msg.isStarred;
        const starredClass = isStarred ? 'message--starred' : '';
        const starIndicator = isStarred ? `
            <span class="message__star-indicator">
                <svg viewBox="0 0 24 24" fill="currentColor" stroke="currentColor" stroke-width="2">
                    <polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"></polygon>
                </svg>
            </span>
        ` : '';
        const diffPopup = isEdited ? generateEditDiffPopup(msg) : '';
        const actionButtons = `
            <div class="message__actions">
                <button class="message__action-btn message__action-btn--star ${isStarred ? 'starred' : ''}" title="${isStarred ? 'Unstar message' : 'Star message'}" data-action="star">
                    <svg viewBox="0 0 24 24" fill="${isStarred ? 'currentColor' : 'none'}" stroke="currentColor" stroke-width="2">
                        <polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"></polygon>
                    </svg>
                </button>
                <button class="message__action-btn message__action-btn--edit" title="Edit message" data-action="edit">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"></path>
                        <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"></path>
                    </svg>
                </button>
                <button class="message__action-btn message__action-btn--delete" title="Delete message" data-action="delete">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <polyline points="3 6 5 6 21 6"></polyline>
                        <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
                    </svg>
                </button>
            </div>
        `;

        return `
            <div class="message message--channel-post ${editedClass} ${starredClass}" ${dataAttr} ${msgIdAttr}>
                <div class="message__post-card">
                    <div class="message__post-accent"></div>
                    <div class="message__post-main">
                        <div class="message__post-header">
                            <div class="message__avatar message__avatar--post" style="background: ${msg.avatarColor || window.TeamsBackupDB.stringToColor(msg.author)}">
                                ${msg.initials || window.TeamsBackupDB.getInitials(msg.author)}
                            </div>
                            <div class="message__post-meta">
                                <div class="message__post-meta-row">
                                    <span class="message__author">${escapeHtml(msg.author)}</span>
                                    <span class="message__time">${time}</span>
                                    ${starIndicator}
                                    ${editIndicator}
                                </div>
                            </div>
                        </div>
                        <div class="message__post-body">
                            ${titleHtml}
                            ${bodyHtml}
                            ${imagesHtml}
                            ${attachmentsHtml}
                            ${diffPopup}
                        </div>
                        <div class="message__post-footer">
                            <span class="message__post-footer-icon">↩</span>
                            <span class="message__post-footer-text">Reply in thread</span>
                        </div>
                    </div>
                    ${actionButtons}
                </div>
            </div>
        `;
    }

    // Render a deleted message
    function renderDeletedMessage(msg, messageIndex = -1) {
        const time = extractTime(msg.timestamp);
        const dataAttr = messageIndex >= 0 ? `data-message-index="${messageIndex}"` : '';
        const msgIdAttr = msg.id ? `data-message-id="${msg.id}"` : '';

        // Generate diff popup showing original content
        const diffPopup = generateDeletedDiffPopup(msg);

        return `
            <div class="message message--deleted" ${dataAttr} ${msgIdAttr}>
                <div class="message__avatar" style="background: ${msg.avatarColor || window.TeamsBackupDB.stringToColor(msg.originalAuthor || msg.author)}">
                    ${msg.initials || window.TeamsBackupDB.getInitials(msg.originalAuthor || msg.author)}
                </div>
                <div class="message__content">
                    <div class="message__header">
                        <span class="message__author">${escapeHtml(msg.originalAuthor || msg.author)}</span>
                        <span class="message__time">${time}</span>
                    </div>
                    <div class="message__deleted-label">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                            <polyline points="3 6 5 6 21 6"></polyline>
                            <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
                        </svg>
                        <span>This message was deleted</span>
                        <button class="message__restore-btn" data-action="restore">Restore</button>
                    </div>
                    ${diffPopup}
                </div>
            </div>
        `;
    }

    // Generate diff popup for edited messages
    function generateEditDiffPopup(msg) {
        if (!msg.editHistory || msg.editHistory.length === 0) return '';

        let historyHtml = '';

        // Show edit history (oldest first)
        msg.editHistory.forEach((edit, index) => {
            const editTime = new Date(edit.editedAt).toLocaleString();
            historyHtml += `
                <div class="diff-popup__item diff-popup__item--removed">
                    <div class="diff-popup__label">Version ${index + 1}</div>
                    <div class="diff-popup__text">${escapeHtml(edit.content)}</div>
                    <div class="diff-popup__time">${editTime}</div>
                </div>
            `;
        });

        // Current version
        historyHtml += `
            <div class="diff-popup__item diff-popup__item--current">
                <div class="diff-popup__label">Current</div>
                <div class="diff-popup__text">${escapeHtml(msg.content)}</div>
            </div>
        `;

        return `
            <div class="diff-popup">
                <div class="diff-popup__header">Edit History</div>
                <div class="diff-popup__content">
                    ${historyHtml}
                </div>
            </div>
        `;
    }

    // Generate diff popup for deleted messages
    function generateDeletedDiffPopup(msg) {
        if (!msg.originalContent) return '';

        const deletedTime = msg.deletedAt ? new Date(msg.deletedAt).toLocaleString() : 'Unknown';

        return `
            <div class="diff-popup">
                <div class="diff-popup__header">Deleted Message</div>
                <div class="diff-popup__content">
                    <div class="diff-popup__item diff-popup__item--removed">
                        <div class="diff-popup__label">Original Content</div>
                        <div class="diff-popup__text">${escapeHtml(msg.originalContent)}</div>
                        <div class="diff-popup__time">Deleted: ${deletedTime}</div>
                    </div>
                </div>
            </div>
        `;
    }

    // Render system message with appropriate icon
    function renderSystemMessage(msg, messageIndex = -1) {
        const content = msg.content || '';
        const time = extractTime(msg.timestamp);
        const dataAttr = messageIndex >= 0 ? `data-message-index="${messageIndex}"` : '';

        // Determine icon based on message content
        let iconSvg = '';
        const lowerContent = content.toLowerCase();

        if (lowerContent.includes('meeting ended') || lowerContent.includes('call ended')) {
            // Video/call ended icon
            iconSvg = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <rect x="2" y="6" width="13" height="12" rx="2"></rect>
                <path d="M20 8l-5 3v2l5 3V8z"></path>
            </svg>`;
        } else if (lowerContent.includes('left the chat') || lowerContent.includes('left the meeting')) {
            // User left icon
            iconSvg = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <path d="M17 16l4-4m0 0l-4-4m4 4H7"></path>
                <circle cx="9" cy="12" r="5"></circle>
            </svg>`;
        } else if (lowerContent.includes('joined') || lowerContent.includes('added')) {
            // User joined icon
            iconSvg = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <path d="M16 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path>
                <circle cx="8.5" cy="7" r="4"></circle>
                <line x1="20" y1="8" x2="20" y2="14"></line>
                <line x1="17" y1="11" x2="23" y2="11"></line>
            </svg>`;
        } else if (lowerContent.includes('invited')) {
            // Invitation icon
            iconSvg = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path>
                <polyline points="22 4 12 14.01 9 11.01"></polyline>
            </svg>`;
        } else if (lowerContent.includes('meeting') || lowerContent.includes('call')) {
            // Generic meeting/call icon
            iconSvg = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <rect x="2" y="6" width="13" height="12" rx="2"></rect>
                <path d="M20 8l-5 3v2l5 3V8z"></path>
            </svg>`;
        } else {
            // Generic info icon
            iconSvg = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <circle cx="12" cy="12" r="10"></circle>
                <line x1="12" y1="16" x2="12" y2="12"></line>
                <line x1="12" y1="8" x2="12.01" y2="8"></line>
            </svg>`;
        }

        return `
            <div class="message message--system" ${dataAttr}>
                <div class="message__icon">${iconSvg}</div>
                <div class="message__content">
                    <div class="message__text">${escapeHtml(content)}</div>
                </div>
                ${time ? `<span class="message__time">${time}</span>` : ''}
            </div>
        `;
    }

    // Extract time from timestamp
    function extractTime(timestamp) {
        if (!timestamp) return '';

        const date = parseTimestamp(timestamp);
        if (date) {
            return date.toLocaleTimeString(undefined, {
                hour: '2-digit',
                minute: '2-digit'
            });
        }

        // Fallback: try to extract time pattern
        const timeMatch = timestamp.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?/i);
        if (timeMatch) {
            return timeMatch[0];
        }

        return timestamp;
    }

    // Search functionality
    searchInput.addEventListener('input', (e) => {
        const query = e.target.value.toLowerCase().trim();

        if (!query) {
            renderChatList(allChats);
            return;
        }

        const filtered = allChats.filter(chat =>
            chat.name.toLowerCase().includes(query)
        );
        renderChatList(filtered);
    });

    // Download dropdown toggle
    function toggleDownloadDropdown(forceClose) {
        const isOpen = downloadDropdown.classList.contains('open');
        if (forceClose || isOpen) {
            downloadDropdown.classList.remove('open');
            downloadMenu.style.display = 'none';
        } else {
            downloadDropdown.classList.add('open');
            downloadMenu.style.display = 'block';
        }
    }

    downloadBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        toggleDownloadDropdown();
    });

    document.addEventListener('click', (e) => {
        if (!downloadDropdown.contains(e.target)) {
            toggleDownloadDropdown(true);
        }
    });

    downloadMenu.addEventListener('click', (e) => {
        if (e.target === downloadMenu) {
            toggleDownloadDropdown(true);
        }
    });

    // Helper to trigger file download from a Blob
    function downloadBlob(blob, filename) {
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }

    // Download PDF option
    downloadPdfOption.addEventListener('click', async () => {
        toggleDownloadDropdown(true);
        if (!currentChat || !currentMessages || currentMessages.length === 0) {
            alert('No messages available to generate PDF');
            return;
        }
        try {
            const pdfBlob = await generateStyledPdf(currentChat, currentMessages);
            downloadBlob(pdfBlob, `${sanitizeFilename(currentChat.name)}_backup.pdf`);
        } catch (error) {
            console.error('Error generating PDF:', error);
            alert('Error generating PDF');
        }
    });

    // Download HTML option
    downloadHtmlOption.addEventListener('click', () => {
        toggleDownloadDropdown(true);
        if (!currentChat || !currentMessages || currentMessages.length === 0) {
            alert('No messages available to generate HTML');
            return;
        }
        try {
            const htmlContent = generateHtmlExport(currentChat, currentMessages);
            const blob = new Blob([htmlContent], { type: 'text/html;charset=utf-8' });
            downloadBlob(blob, `${sanitizeFilename(currentChat.name)}_backup.html`);
        } catch (error) {
            console.error('Error generating HTML:', error);
            alert('Error generating HTML file');
        }
    });

    // Download TXT option
    downloadTxtOption.addEventListener('click', () => {
        toggleDownloadDropdown(true);
        if (!currentChat || (!currentChat.textContent && (!currentMessages || currentMessages.length === 0))) {
            alert('No text content available for this backup');
            return;
        }
        try {
            const BOM = '\uFEFF';
            const txtContent = generateTxtExport(currentChat, currentMessages);
            const blob = new Blob([BOM + txtContent], { type: 'text/plain;charset=utf-8' });
            downloadBlob(blob, `${sanitizeFilename(currentChat.name)}_backup.txt`);
        } catch (error) {
            console.error('Error downloading TXT:', error);
            alert('Error downloading text file');
        }
    });

    // Delete chat
    deleteChatBtn.addEventListener('click', () => {
        deleteModal.style.display = 'flex';
    });

    cancelDeleteBtn.addEventListener('click', () => {
        deleteModal.style.display = 'none';
    });

    confirmDeleteBtn.addEventListener('click', async () => {
        if (!currentChatId) return;

        try {
            await window.TeamsBackupDB.deleteChat(currentChatId);

            // Reset state
            currentChatId = null;
            currentChat = null;
            chatView.style.display = 'none';
            welcomeState.style.display = 'flex';
            deleteModal.style.display = 'none';

            // Reload chat list and update trash count
            await loadChats();
            await updateTrashCount();

        } catch (error) {
            console.error('Error deleting chat:', error);
            alert('Error deleting backup');
        }
    });

    // Close modal on backdrop click
    deleteModal.querySelector('.modal__backdrop').addEventListener('click', () => {
        deleteModal.style.display = 'none';
    });

    // Export All functionality
    exportAllBtn.addEventListener('click', async () => {
        if (!isPro) {
            showUpgradeModal('Export All saves every backup in one ZIP file. It is part of Pro.');
            return;
        }

        if (allChats.length === 0) {
            alert('No backups to export.');
            return;
        }

        // Show loading state
        const originalText = exportAllBtn.innerHTML;
        exportAllBtn.innerHTML = `
            <svg class="export-all-btn__icon spinning" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <circle cx="12" cy="12" r="10"></circle>
                <path d="M12 6v6l4 2"></path>
            </svg>
            <span id="exportAllProgress">Exporting...</span>
        `;
        exportAllBtn.disabled = true;
        const exportAllProgress = document.getElementById('exportAllProgress');

        try {
            const zip = new JSZip();
            const mainFolder = zip.folder('Teams_Chat_Backup');
            const usedFolderNames = new Set();

            function getUniqueFolderName(chatName) {
                const baseName = sanitizeFilename(chatName) || 'chat';
                let candidate = baseName;
                let suffixIndex = 2;

                while (usedFolderNames.has(candidate)) {
                    const suffix = ` (${suffixIndex})`;
                    const maxBaseLength = Math.max(1, 100 - suffix.length);
                    candidate = `${baseName.substring(0, maxBaseLength).trim()}${suffix}`;
                    suffixIndex++;
                }

                usedFolderNames.add(candidate);
                return candidate;
            }

            // Create importable backup data
            const importableBackup = {
                version: '1.0',
                exportDate: new Date().toISOString(),
                chats: []
            };

            for (const [chatIndex, chatMeta] of allChats.entries()) {
                if (exportAllProgress) {
                    exportAllProgress.textContent = `Exporting ${chatIndex + 1} of ${allChats.length}...`;
                }
                // Get full chat data including messages
                const [chat, messages] = await Promise.all([
                    window.TeamsBackupDB.getChat(chatMeta.id),
                    window.TeamsBackupDB.getMessages(chatMeta.id)
                ]);
                if (!chat) continue;

                // Create folder for this chat
                const folderName = getUniqueFolderName(chat.name);
                const chatFolder = mainFolder.folder(folderName);

                // Add PDF export. Free exports are regenerated so the watermark is applied.
                if (chat.pdfBlob || (messages && messages.length > 0)) {
                    const pdfBlob = (isPro && chat.pdfBlob)
                        ? chat.pdfBlob
                        : (messages && messages.length > 0 ? await generateStyledPdf(chat, messages) : null);
                    if (pdfBlob) {
                        chatFolder.file(`${folderName}_backup.pdf`, pdfBlob);
                    }
                }

                // Add HTML if messages available
                if (messages && messages.length > 0) {
                    const htmlContent = generateHtmlExport(chat, messages);
                    chatFolder.file(`${folderName}_backup.html`, htmlContent);
                }

                // Add TXT if available (with UTF-8 BOM for Korean/international characters)
                if (chat.textContent || (messages && messages.length > 0)) {
                    const BOM = '\uFEFF';
                    chatFolder.file(`${folderName}_backup.txt`, BOM + generateTxtExport(chat, messages));
                }

                // Add metadata file
                    const metadata = {
                        name: chat.name,
                        rawName: chat.rawName || chat.name,
                        isChannel: chat.isChannel === true,
                        teamName: chat.teamName || null,
                        sourceKind: chat.sourceKind || (chat.isChannel ? 'channel' : 'chat'),
                        messageCount: chat.messageCount,
                        participants: chat.participants,
                        createdAt: chat.createdAt,
                        updatedAt: chat.updatedAt || chat.createdAt,
                        isPreview: chat.isPreview
                };
                chatFolder.file('metadata.json', JSON.stringify(metadata, null, 2));

                // Add to importable backup (excluding blob for JSON compatibility)
                importableBackup.chats.push({
                    name: chat.name,
                    rawName: chat.rawName || chat.name,
                    isChannel: chat.isChannel === true,
                    teamName: chat.teamName || null,
                    sourceKind: chat.sourceKind || (chat.isChannel ? 'channel' : 'chat'),
                    messageCount: chat.messageCount,
                    participants: chat.participants || [],
                    createdAt: chat.createdAt,
                    updatedAt: chat.updatedAt || chat.createdAt,
                    isPreview: chat.isPreview || false,
                    textContent: chat.textContent || '',
                    messages: messages || []
                });
            }

            // Add importable backup file (.tcb = Teams Chat Backup)
            mainFolder.file('backup_data.tcb', JSON.stringify(importableBackup, null, 2));

            // Generate ZIP file
            const zipBlob = await zip.generateAsync({
                type: 'blob',
                compression: 'DEFLATE',
                compressionOptions: { level: 6 }
            });

            // Download the ZIP
            const timestamp = new Date().toISOString().split('T')[0];
            const url = URL.createObjectURL(zipBlob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `Teams_Chat_Backup_${timestamp}.zip`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);

        } catch (error) {
            console.error('Error exporting backups:', error);
            alert('Error exporting backups. Please try again.');
        } finally {
            // Restore button state
            exportAllBtn.innerHTML = originalText;
            exportAllBtn.disabled = false;
        }
    });

    // Import functionality
    importBtn.addEventListener('click', () => {
        if (!isPro) {
            showUpgradeModal('Importing backups from a file is part of Pro.');
            return;
        }
        importFileInput.click();
    });

    importFileInput.addEventListener('change', async (e) => {
        if (!isPro) {
            importFileInput.value = '';
            alert('Import is for Pro users only.');
            return;
        }

        const file = e.target.files[0];
        if (!file) return;

        // Show loading state
        const originalText = importBtn.innerHTML;
        importBtn.innerHTML = `
            <svg class="import-btn__icon spinning" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <circle cx="12" cy="12" r="10"></circle>
                <path d="M12 6v6l4 2"></path>
            </svg>
            <span>Importing...</span>
        `;
        importBtn.disabled = true;

        try {
            let backupData;

            if (file.name.endsWith('.tcb') || file.name.endsWith('.json')) {
                // Direct JSON/TCB file
                const text = await file.text();
                backupData = JSON.parse(text);
            } else if (file.name.endsWith('.zip')) {
                // ZIP file - look for backup_data.tcb inside
                const zip = await JSZip.loadAsync(file);

                // Find the backup_data.tcb file
                let tcbFile = null;
                zip.forEach((relativePath, zipEntry) => {
                    if (relativePath.endsWith('backup_data.tcb')) {
                        tcbFile = zipEntry;
                    }
                });

                if (!tcbFile) {
                    throw new Error('No backup_data.tcb found in ZIP file. This ZIP may not be a Teams Chat Backup export.');
                }

                const text = await tcbFile.async('string');
                backupData = JSON.parse(text);
            } else {
                throw new Error('Unsupported file type. Please use a .tcb, .json, or .zip file.');
            }

            // Validate backup data
            if (!backupData.chats || !Array.isArray(backupData.chats)) {
                throw new Error('Invalid backup file format.');
            }

            // Import chats
            let importedCount = 0;
            let skippedCount = 0;

            for (const chatData of backupData.chats) {
                try {
                    // Check if chat already exists by name
                    const existingChats = await window.TeamsBackupDB.getChats();
                    const existingChat = existingChats.find(c => c.name === chatData.name);

                    if (existingChat) {
                        // Chat exists - merge messages
                        const result = await window.TeamsBackupDB.saveChat({
                            name: chatData.name,
                            rawName: chatData.rawName || chatData.name,
                            isChannel: chatData.isChannel === true,
                            teamName: chatData.teamName || null,
                            sourceKind: chatData.sourceKind || (chatData.isChannel ? 'channel' : 'chat'),
                            messages: chatData.messages || [],
                            participants: chatData.participants || [],
                            pdfBlob: null, // PDFs are not included in importable backup
                            textContent: chatData.textContent || '',
                            isPreview: chatData.isPreview || false,
                            isPro: isPro
                        });

                        if (result.merged) {
                            importedCount++;
                        } else {
                            skippedCount++;
                        }
                    } else {
                        // New chat - save it
                        await window.TeamsBackupDB.saveChat({
                            name: chatData.name,
                            rawName: chatData.rawName || chatData.name,
                            isChannel: chatData.isChannel === true,
                            teamName: chatData.teamName || null,
                            sourceKind: chatData.sourceKind || (chatData.isChannel ? 'channel' : 'chat'),
                            messages: chatData.messages || [],
                            participants: chatData.participants || [],
                            pdfBlob: null,
                            textContent: chatData.textContent || '',
                            isPreview: chatData.isPreview || false,
                            isPro: isPro
                        });
                        importedCount++;
                    }
                } catch (chatError) {
                    console.error('Error importing chat:', chatData.name, chatError);
                    skippedCount++;
                }
            }

            // Reload chat list
            await loadChats();

            // Show success message
            let message = `Import complete! ${importedCount} chat(s) imported.`;
            if (skippedCount > 0) {
                message += ` ${skippedCount} chat(s) skipped (already exist or error).`;
            }
            alert(message);

        } catch (error) {
            console.error('Error importing backups:', error);
            alert(`Error importing backups: ${error.message}`);
        } finally {
            // Restore button state
            importBtn.innerHTML = originalText;
            importBtn.disabled = false;
            // Reset file input
            importFileInput.value = '';
        }
    });

    // ==========================================
    // SEARCH FUNCTIONALITY
    // ==========================================

    // Perform search with debouncing
    function performSearch(query) {
        // Clear previous highlights
        clearSearchHighlights();
        searchMatches = [];
        currentMatchIndex = -1;

        if (!query || query.length < 2) {
            searchResultsDiv.style.display = 'none';
            clearSearchBtn.style.display = 'none';
            return;
        }

        clearSearchBtn.style.display = 'block';
        const lowerQuery = query.toLowerCase();

        // Find all matching messages
        const messageElements = messagesContainer.querySelectorAll('.message[data-message-index]');

        messageElements.forEach((msgEl, index) => {
            const messageIndex = parseInt(msgEl.dataset.messageIndex);
            const msg = currentMessages[messageIndex];
            if (!msg) return;

            const authorMatch = msg.author && msg.author.toLowerCase().includes(lowerQuery);
            const contentMatch = msg.content && msg.content.toLowerCase().includes(lowerQuery);

            if (authorMatch || contentMatch) {
                searchMatches.push({
                    element: msgEl,
                    messageIndex: messageIndex,
                    message: msg
                });
                msgEl.classList.add('message--search-match');

                // Highlight matching text
                if (contentMatch) {
                    const titleEl = msgEl.querySelector('.message__post-title');
                    if (titleEl) {
                        const postContent = splitPostMessageContent(msg.content || '');
                        titleEl.innerHTML = highlightText(postContent.title, query);
                    }
                    const textEl = msgEl.querySelector('.message__text');
                    if (textEl) {
                        const postContent = splitPostMessageContent(msg.content || '');
                        const textToHighlight = titleEl ? (postContent.body || '') : msg.content;
                        textEl.innerHTML = highlightText(textToHighlight, query);
                    }
                }
                if (authorMatch) {
                    const authorEl = msgEl.querySelector('.message__author');
                    if (authorEl) {
                        authorEl.innerHTML = highlightText(msg.author, query);
                    }
                }
            }
        });

        // Update UI
        updateSearchResultsUI();

        // Jump to first result if any
        if (searchMatches.length > 0) {
            navigateToMatch(0);
        }
    }

    // Highlight text with search term
    function highlightText(text, query) {
        if (!text || !query) return escapeHtmlWithBreaks(text);

        const escaped = escapeHtml(text);
        const lowerText = escaped.toLowerCase();
        const lowerQuery = query.toLowerCase();

        let result = '';
        let lastIndex = 0;
        let index = lowerText.indexOf(lowerQuery);

        while (index !== -1) {
            result += escaped.slice(lastIndex, index);
            result += `<mark class="search-highlight">${escaped.slice(index, index + query.length)}</mark>`;
            lastIndex = index + query.length;
            index = lowerText.indexOf(lowerQuery, lastIndex);
        }

        result += escaped.slice(lastIndex);
        return result.replace(/\n/g, '<br>');
    }

    // Clear all search highlights
    function clearSearchHighlights() {
        // Remove match classes
        messagesContainer.querySelectorAll('.message--search-match').forEach(el => {
            el.classList.remove('message--search-match', 'message--search-current');
        });

        // Restore original text (re-render without highlights)
        messagesContainer.querySelectorAll('.search-highlight').forEach(mark => {
            const parent = mark.parentNode;
            parent.replaceChild(document.createTextNode(mark.textContent), mark);
            parent.normalize();
        });
    }

    // Update search results UI
    function updateSearchResultsUI() {
        if (searchMatches.length === 0) {
            searchResultCount.textContent = 'No results';
            prevResultBtn.disabled = true;
            nextResultBtn.disabled = true;
        } else {
            searchResultCount.textContent = `${currentMatchIndex + 1} of ${searchMatches.length}`;
            prevResultBtn.disabled = searchMatches.length <= 1;
            nextResultBtn.disabled = searchMatches.length <= 1;
        }
        searchResultsDiv.style.display = 'flex';
    }

    // Navigate to a specific match
    function navigateToMatch(index) {
        if (searchMatches.length === 0) return;

        // Remove current highlight
        if (currentMatchIndex >= 0 && searchMatches[currentMatchIndex]) {
            searchMatches[currentMatchIndex].element.classList.remove('message--search-current');
        }

        // Update index with wraparound
        currentMatchIndex = index;
        if (currentMatchIndex < 0) currentMatchIndex = searchMatches.length - 1;
        if (currentMatchIndex >= searchMatches.length) currentMatchIndex = 0;

        // Add current highlight and scroll
        const match = searchMatches[currentMatchIndex];
        match.element.classList.add('message--search-current');
        match.element.scrollIntoView({ behavior: 'smooth', block: 'center' });

        updateSearchResultsUI();
    }

    // Clear search
    function clearSearch() {
        messageSearchInput.value = '';
        clearSearchHighlights();
        searchMatches = [];
        currentMatchIndex = -1;
        searchResultsDiv.style.display = 'none';
        clearSearchBtn.style.display = 'none';
    }

    // Search event handlers
    messageSearchInput.addEventListener('input', (e) => {
        const query = e.target.value.trim();

        // Debounce search for performance
        clearTimeout(searchDebounceTimer);
        searchDebounceTimer = setTimeout(() => {
            performSearch(query);
        }, 150);
    });

    // Enter key navigates to next result
    messageSearchInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            e.preventDefault();
            if (e.shiftKey) {
                navigateToMatch(currentMatchIndex - 1);
            } else {
                navigateToMatch(currentMatchIndex + 1);
            }
        } else if (e.key === 'Escape') {
            clearSearch();
            messageSearchInput.blur();
        }
    });

    clearSearchBtn.addEventListener('click', clearSearch);
    prevResultBtn.addEventListener('click', () => navigateToMatch(currentMatchIndex - 1));
    nextResultBtn.addEventListener('click', () => navigateToMatch(currentMatchIndex + 1));

    // Keyboard shortcut: Ctrl/Cmd + F to focus search
    document.addEventListener('keydown', (e) => {
        if ((e.ctrlKey || e.metaKey) && e.key === 'f') {
            // Only intercept if chat view is visible
            if (chatView.style.display !== 'none') {
                e.preventDefault();
                messageSearchInput.focus();
                messageSearchInput.select();
            }
        }
    });

    // ==========================================
    // DATE FILTER FUNCTIONALITY
    // ==========================================

    // Populate date filter dropdown
    function populateDateFilter(messages) {
        // Clear existing options except the first one
        dateFilter.innerHTML = '<option value="">Jump to date...</option>';

        if (!messages || messages.length === 0) return;

        // Get unique dates
        const dateMap = new Map();

        messages.forEach((msg, index) => {
            const date = parseTimestamp(msg.timestamp);
            if (date) {
                const dateKey = getLocalDateKey(date);
                if (!dateMap.has(dateKey)) {
                    dateMap.set(dateKey, {
                        date: date,
                        label: formatDateHeader(date),
                        messageCount: 1
                    });
                } else {
                    dateMap.get(dateKey).messageCount++;
                }
            }
        });

        // Sort dates (oldest first to match chat order)
        const sortedDates = Array.from(dateMap.entries()).sort((a, b) =>
            new Date(a[0]) - new Date(b[0])
        );

        // Add options
        sortedDates.forEach(([dateKey, info]) => {
            const option = document.createElement('option');
            option.value = dateKey;
            option.textContent = `${info.label} (${info.messageCount} msgs)`;
            dateFilter.appendChild(option);
        });
    }

    // Jump to date
    dateFilter.addEventListener('change', (e) => {
        const dateKey = e.target.value;
        if (!dateKey) return;

        // Find the date divider with matching date
        const dateDividers = messagesContainer.querySelectorAll('.date-divider[data-date]');

        for (const divider of dateDividers) {
            if (divider.dataset.date === dateKey) {
                divider.scrollIntoView({ behavior: 'smooth', block: 'start' });

                // Flash effect to show where we jumped
                divider.style.background = 'rgba(98, 100, 167, 0.2)';
                setTimeout(() => {
                    divider.style.background = '';
                }, 1500);
                break;
            }
        }

        // Reset the dropdown to placeholder
        setTimeout(() => {
            dateFilter.value = '';
        }, 100);
    });

    // ==========================================
    // END SEARCH & DATE FILTER
    // ==========================================

    // ==========================================
    // EDIT/DELETE FUNCTIONALITY
    // ==========================================

    // Load viewer settings from storage
    async function loadViewerSettings() {
        return new Promise((resolve) => {
            chrome.storage.local.get(['viewerSettings'], (result) => {
                if (result.viewerSettings) {
                    viewerSettingsState = { ...viewerSettingsState, ...result.viewerSettings };
                }

                // Apply settings to UI
                hideDeletedCheckbox.checked = viewerSettingsState.hideDeleted;
                showEditHistoryCheckbox.checked = viewerSettingsState.showEditHistory;

                // Apply hide-deleted class to messages container
                if (viewerSettingsState.hideDeleted) {
                    messagesContainer.classList.add('hide-deleted');
                }

                resolve(viewerSettingsState);
            });
        });
    }

    // Save viewer settings to storage
    function saveViewerSettings() {
        chrome.storage.local.set({ viewerSettings: viewerSettingsState });
    }

    // Initialize settings on load
    loadViewerSettings();

    // Settings panel handlers
    viewerSettingsBtn.addEventListener('click', () => {
        viewerSettings.style.display = 'block';
        // Create backdrop
        const backdrop = document.createElement('div');
        backdrop.className = 'viewer-settings-backdrop';
        backdrop.id = 'viewerSettingsBackdrop';
        backdrop.addEventListener('click', closeSettings);
        document.body.appendChild(backdrop);
    });

    function closeSettings() {
        viewerSettings.style.display = 'none';
        const backdrop = document.getElementById('viewerSettingsBackdrop');
        if (backdrop) backdrop.remove();
    }

    closeSettingsBtn.addEventListener('click', closeSettings);

    // Settings checkbox handlers
    hideDeletedCheckbox.addEventListener('change', (e) => {
        viewerSettingsState.hideDeleted = e.target.checked;
        saveViewerSettings();

        if (e.target.checked) {
            messagesContainer.classList.add('hide-deleted');
        } else {
            messagesContainer.classList.remove('hide-deleted');
        }
    });

    showEditHistoryCheckbox.addEventListener('change', (e) => {
        viewerSettingsState.showEditHistory = e.target.checked;
        saveViewerSettings();

        // Re-render messages to apply setting
        if (currentChatId) {
            loadMessages(currentChatId);
        }
    });

    // Handle click events on message actions (event delegation)
    messagesContainer.addEventListener('click', async (e) => {
        const actionBtn = e.target.closest('[data-action]');
        if (!actionBtn) return;

        const action = actionBtn.dataset.action;
        const messageEl = actionBtn.closest('.message');
        const messageId = messageEl?.dataset.messageId;
        const messageIndex = parseInt(messageEl?.dataset.messageIndex);

        if (!messageId && messageIndex === undefined) return;

        const msg = currentMessages[messageIndex];
        if (!msg) return;

        switch (action) {
            case 'star':
                await toggleStar(msg);
                break;
            case 'edit':
                openEditModal(msg);
                break;
            case 'delete':
                openDeleteModal(msg);
                break;
            case 'restore':
                await restoreMessage(msg);
                break;
        }
    });

    // Toggle star on a message
    async function toggleStar(msg) {
        if (!msg.id) return;

        try {
            await window.TeamsBackupDB.toggleStarMessage(msg.id);
            // Reload messages to show updated state
            if (showingStarredOnly) {
                await loadStarredMessages();
            } else {
                await loadMessages(currentChatId);
            }
        } catch (error) {
            console.error('Error toggling star:', error);
        }
    }

    // Update starred count badge
    function updateStarredCount() {
        const count = currentMessages.filter(msg => msg.isStarred).length;
        if (count > 0) {
            starredCount.textContent = count;
            starredCount.style.display = 'inline-block';
        } else {
            starredCount.style.display = 'none';
        }
    }

    // Starred filter button handler
    starredFilterBtn.addEventListener('click', async () => {
        showingStarredOnly = !showingStarredOnly;
        starredFilterBtn.classList.toggle('active', showingStarredOnly);

        if (showingStarredOnly) {
            await loadStarredMessages();
        } else {
            await loadMessages(currentChatId);
        }
    });

    // Load only starred messages
    async function loadStarredMessages() {
        if (!currentChatId) return;

        try {
            const allMessages = await window.TeamsBackupDB.getMessages(currentChatId);
            const starredMessages = allMessages.filter(msg => msg.isStarred);

            // Store all messages for reference but show filtered
            currentMessages = allMessages;

            if (starredMessages.length === 0) {
                messagesContainer.innerHTML = `
                    <div class="starred-empty-state">
                        <svg class="starred-empty-state__icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5">
                            <polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"></polygon>
                        </svg>
                        <h3>No starred messages</h3>
                        <p>Star important messages to find them quickly later. Click the star icon on any message to star it.</p>
                    </div>
                `;
                return;
            }

            // Build HTML with starred header
            let html = `
                <div class="starred-filter-header">
                    <div class="starred-filter-header__text">
                        <svg viewBox="0 0 24 24" stroke="currentColor" stroke-width="2">
                            <polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"></polygon>
                        </svg>
                        Showing ${starredMessages.length} starred message${starredMessages.length !== 1 ? 's' : ''}
                    </div>
                    <button class="starred-filter-header__close" id="closeStarredFilter">&times;</button>
                </div>
            `;

            // Group starred messages by date
            const groupedMessages = groupMessagesByDate(starredMessages);

            for (const [date, dateInfo] of Object.entries(groupedMessages)) {
                html += `<div class="date-divider" data-date="${dateInfo.dateKey}">${date}</div>`;

                for (const msg of dateInfo.messages) {
                    const originalIndex = allMessages.findIndex(m => m.id === msg.id);
                    if (msg.type === 'system') {
                        html += renderSystemMessage(msg, originalIndex);
                    } else {
                        html += renderMessage(msg, originalIndex);
                    }
                }
            }

            messagesContainer.innerHTML = html;

            // Add click handler for close button
            const closeBtn = document.getElementById('closeStarredFilter');
            if (closeBtn) {
                closeBtn.addEventListener('click', async () => {
                    showingStarredOnly = false;
                    starredFilterBtn.classList.remove('active');
                    await loadMessages(currentChatId);
                });
            }

            // Update starred count
            updateStarredCount();

        } catch (error) {
            console.error('Error loading starred messages:', error);
        }
    }

    // Open edit modal
    function openEditModal(msg) {
        editingMessageId = msg.id;
        editOriginalContent.textContent = msg.content || '';
        editMessageInput.value = msg.content || '';
        editMessageModal.style.display = 'flex';
        editMessageInput.focus();
    }

    // Close edit modal
    function closeEditModal() {
        editMessageModal.style.display = 'none';
        editingMessageId = null;
        editMessageInput.value = '';
    }

    cancelEditBtn.addEventListener('click', closeEditModal);
    editMessageModal.querySelector('.modal__backdrop').addEventListener('click', closeEditModal);

    // Confirm edit
    confirmEditBtn.addEventListener('click', async () => {
        if (!editingMessageId) return;

        const newContent = editMessageInput.value.trim();
        if (!newContent) {
            alert('Message content cannot be empty');
            return;
        }

        try {
            await window.TeamsBackupDB.editMessage(editingMessageId, newContent);
            closeEditModal();
            // Reload messages to show updated content
            await loadMessages(currentChatId);
        } catch (error) {
            console.error('Error editing message:', error);
            alert('Failed to edit message');
        }
    });

    // Open delete modal
    function openDeleteModal(msg) {
        deletingMessageId = msg.id;
        deleteMessagePreview.textContent = msg.content || '(no content)';
        deleteMessageModal.style.display = 'flex';
    }

    // Close delete modal
    function closeDeleteModal() {
        deleteMessageModal.style.display = 'none';
        deletingMessageId = null;
    }

    cancelDeleteMsgBtn.addEventListener('click', closeDeleteModal);
    deleteMessageModal.querySelector('.modal__backdrop').addEventListener('click', closeDeleteModal);

    // Confirm delete
    confirmDeleteMsgBtn.addEventListener('click', async () => {
        if (!deletingMessageId) return;

        try {
            await window.TeamsBackupDB.softDeleteMessage(deletingMessageId);
            closeDeleteModal();
            // Reload messages to show deleted state
            await loadMessages(currentChatId);
        } catch (error) {
            console.error('Error deleting message:', error);
            alert('Failed to delete message');
        }
    });

    // Restore a deleted message
    async function restoreMessage(msg) {
        if (!msg.id) return;

        try {
            await window.TeamsBackupDB.restoreMessage(msg.id);
            // Reload messages to show restored content
            await loadMessages(currentChatId);
        } catch (error) {
            console.error('Error restoring message:', error);
            alert('Failed to restore message');
        }
    }

    // ==========================================
    // END EDIT/DELETE FUNCTIONALITY
    // ==========================================

    // ==========================================
    // TAB NAVIGATION & ATTACHMENTS
    // ==========================================

    // Tab click handlers
    chatTabs.forEach(tab => {
        tab.addEventListener('click', () => {
            const tabName = tab.dataset.tab;
            switchTab(tabName);
        });
    });

    // Switch between tabs
    function switchTab(tabName) {
        currentTab = tabName;

        // Update tab active state
        chatTabs.forEach(tab => {
            tab.classList.toggle('chat-tab--active', tab.dataset.tab === tabName);
        });

        // Show/hide containers
        if (tabName === 'messages') {
            messagesContainer.style.display = 'block';
            attachmentsContainer.style.display = 'none';
            searchToolbar.style.display = 'flex';
        } else if (tabName === 'attachments') {
            messagesContainer.style.display = 'none';
            attachmentsContainer.style.display = 'block';
            searchToolbar.style.display = 'none';
            populateAttachmentsGrid();
        }
    }

    // Collect all attachments (images + files) from messages
    function collectAllAttachments() {
        const allAttachments = [];

        currentMessages.forEach((msg, index) => {
            // Skip deleted messages
            if (msg.isDeleted) return;

            // Collect images
            if (msg.images && msg.images.length > 0) {
                msg.images.forEach(img => {
                    // Animated emoji sprites are part of the message, not media attachments.
                    if (img.isAnimatedEmoji) return;
                    allAttachments.push({
                        type: 'image',
                        dataUrl: img.dataUrl,
                        src: img.src,
                        filename: img.alt || 'Image',
                        checkpointPlaceholder: img.checkpointPlaceholder === true,
                        author: msg.author,
                        timestamp: msg.timestamp,
                        messageIndex: index
                    });
                });
            }

            // Collect file attachments
            if (msg.attachments && msg.attachments.length > 0) {
                msg.attachments.forEach(att => {
                    allAttachments.push({
                        type: 'file',
                        dataUrl: att.dataUrl,
                        filename: att.filename || 'Unknown file',
                        extension: att.extension || '',
                        fileSizeText: att.fileSizeText || '',
                        sharePointUrl: att.sharePointUrl,
                        downloadError: att.downloadError,
                        tooLarge: att.tooLarge,
                        author: msg.author,
                        timestamp: msg.timestamp,
                        messageIndex: index
                    });
                });
            }
        });

        return allAttachments;
    }

    // Get file type category based on extension
    function getFileTypeCategory(extension) {
        const ext = (extension || '').toLowerCase();
        const typeMap = {
            excel: ['xlsx', 'xls', 'xlsm', 'xlsb', 'csv'],
            word: ['docx', 'doc', 'docm', 'rtf'],
            powerpoint: ['pptx', 'ppt', 'pptm'],
            pdf: ['pdf'],
            image: ['jpg', 'jpeg', 'png', 'gif', 'webp', 'bmp', 'svg', 'ico', 'tiff'],
            archive: ['zip', 'rar', '7z', 'tar', 'gz', 'bz2'],
            '3d': ['dae', 'glb', 'gltf', 'obj', 'fbx', 'stl', 'blend', '3ds'],
            code: ['js', 'ts', 'jsx', 'tsx', 'py', 'java', 'cpp', 'c', 'h', 'cs', 'php', 'rb', 'go', 'rs', 'swift', 'kt', 'html', 'css', 'scss', 'json', 'xml', 'yaml', 'yml', 'sql', 'sh', 'bash', 'md'],
            audio: ['mp3', 'wav', 'ogg', 'flac', 'aac', 'm4a', 'wma'],
            video: ['mp4', 'mov', 'avi', 'mkv', 'webm', 'wmv', 'flv', 'm4v']
        };

        for (const [category, extensions] of Object.entries(typeMap)) {
            if (extensions.includes(ext)) return category;
        }
        return 'default';
    }

    // Render a single attachment card (Teams-style with file type icons and actions)
    function createAttachmentCard(att, idx) {
        const card = document.createElement('div');
        card.className = 'attachment-card';
        card.dataset.index = idx;

        const time = extractTime(att.timestamp);
        const author = att.author || 'Unknown';
        const ext = (att.extension || '').toLowerCase();
        const extDisplay = ext.toUpperCase() || 'FILE';
        const fileType = getFileTypeCategory(ext);
        const hasData = (att.type === 'image' ? isEmbeddedImageDataUrl(att.dataUrl) : !!att.dataUrl) && !att.downloadError;
        const isTooLarge = att.tooLarge;
        const hasError = att.downloadError;
        const isImageReference = att.type === 'image' && !hasData;

        // Add state classes
        if (att.type === 'image') {
            card.classList.add('attachment-card--image');
            if (hasData) {
                card.classList.add('attachment-card--available');
            } else {
                card.classList.add('attachment-card--reference');
            }
        } else {
            card.classList.add('attachment-card--file');
            if (hasData) {
                card.classList.add('attachment-card--available');
            } else if (isTooLarge) {
                card.classList.add('attachment-card--too-large');
            } else if (hasError) {
                card.classList.add('attachment-card--error');
            } else {
                card.classList.add('attachment-card--reference');
            }
        }

        // Icon/Preview area
        const iconArea = document.createElement('div');
        iconArea.className = 'attachment-card__icon';

        if (att.type === 'image' && hasData) {
            // Image thumbnail
            const img = document.createElement('img');
            img.src = att.dataUrl;
            img.alt = att.filename;
            img.loading = 'lazy';
            iconArea.appendChild(img);
        } else if (att.type === 'image') {
            const iconInner = document.createElement('div');
            iconInner.className = 'attachment-card__icon-inner attachment-card__icon--image';
            iconInner.textContent = 'IMG';
            iconArea.appendChild(iconInner);
        } else {
            // File type icon
            const iconInner = document.createElement('div');
            iconInner.className = `attachment-card__icon-inner attachment-card__icon--${fileType}`;
            iconInner.textContent = extDisplay.substring(0, 4);
            iconArea.appendChild(iconInner);
        }

        // Body area (info + actions)
        const body = document.createElement('div');
        body.className = 'attachment-card__body';

        // File info
        const info = document.createElement('div');
        info.className = 'attachment-card__info';

        const nameSpan = document.createElement('span');
        nameSpan.className = 'attachment-card__name';
        nameSpan.textContent = att.filename;
        nameSpan.title = att.filename;

        const meta = document.createElement('div');
        meta.className = 'attachment-card__meta';

        const authorSpan = document.createElement('span');
        authorSpan.textContent = author;
        meta.appendChild(authorSpan);

        const divider1 = document.createElement('span');
        divider1.className = 'attachment-card__meta-divider';
        meta.appendChild(divider1);

        const timeSpan = document.createElement('span');
        timeSpan.textContent = time;
        meta.appendChild(timeSpan);

        if (att.fileSizeText) {
            const divider2 = document.createElement('span');
            divider2.className = 'attachment-card__meta-divider';
            meta.appendChild(divider2);

            const sizeSpan = document.createElement('span');
            sizeSpan.textContent = att.fileSizeText;
            meta.appendChild(sizeSpan);
        }

        // Status badge
        if (isTooLarge) {
            const status = document.createElement('span');
            status.className = 'attachment-card__status attachment-card__status--large';
            status.textContent = 'Too large';
            meta.appendChild(status);
        } else if (hasError) {
            const status = document.createElement('span');
            status.className = 'attachment-card__status attachment-card__status--error';
            status.textContent = 'Failed';
            meta.appendChild(status);
        } else if (hasData) {
            const status = document.createElement('span');
            status.className = 'attachment-card__status attachment-card__status--available';
            status.textContent = 'Available';
            meta.appendChild(status);
        } else if (isImageReference) {
            const status = document.createElement('span');
            status.className = 'attachment-card__status attachment-card__status--reference';
            status.textContent = 'Preview unavailable';
            meta.appendChild(status);
        }

        info.appendChild(nameSpan);
        info.appendChild(meta);

        // Action buttons
        const actions = document.createElement('div');
        actions.className = 'attachment-card__actions';

        if (att.type === 'image' && hasData) {
            // View image button
            const viewBtn = document.createElement('button');
            viewBtn.className = 'attachment-card__btn attachment-card__btn--primary';
            viewBtn.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg> View';
            viewBtn.addEventListener('click', (e) => {
                e.stopPropagation();
                window.open(att.dataUrl, '_blank');
            });
            actions.appendChild(viewBtn);

            // Download button
            const downloadBtn = document.createElement('button');
            downloadBtn.className = 'attachment-card__btn attachment-card__btn--secondary';
            downloadBtn.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg> Save';
            downloadBtn.addEventListener('click', (e) => {
                e.stopPropagation();
                const a = document.createElement('a');
                a.href = att.dataUrl;
                a.download = att.filename;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
            });
            actions.appendChild(downloadBtn);
        } else if (hasData) {
            // Download button for files with data
            const downloadBtn = document.createElement('button');
            downloadBtn.className = 'attachment-card__btn attachment-card__btn--primary';
            downloadBtn.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg> Download';
            downloadBtn.addEventListener('click', (e) => {
                e.stopPropagation();
                const a = document.createElement('a');
                a.href = att.dataUrl;
                a.download = att.filename;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
            });
            actions.appendChild(downloadBtn);

            // Open in SharePoint if URL available
            if (att.sharePointUrl) {
                const openBtn = document.createElement('button');
                openBtn.className = 'attachment-card__btn attachment-card__btn--secondary';
                openBtn.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg> Open';
                openBtn.addEventListener('click', (e) => {
                    e.stopPropagation();
                    window.open(att.sharePointUrl, '_blank');
                });
                actions.appendChild(openBtn);
            }
        } else if (att.sharePointUrl) {
            // Only SharePoint link available (no local data)
            const openBtn = document.createElement('button');
            openBtn.className = 'attachment-card__btn attachment-card__btn--primary';
            openBtn.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg> Open in SharePoint';
            openBtn.addEventListener('click', (e) => {
                e.stopPropagation();
                window.open(att.sharePointUrl, '_blank');
            });
            actions.appendChild(openBtn);
        } else {
            // No data available
            const unavailableSpan = document.createElement('span');
            unavailableSpan.style.cssText = 'font-size: 12px; color: var(--teams-text-light); font-style: italic;';
            unavailableSpan.textContent = 'File not available';
            actions.appendChild(unavailableSpan);
        }

        body.appendChild(info);
        body.appendChild(actions);

        card.appendChild(iconArea);
        card.appendChild(body);

        return card;
    }

    // Populate the attachments grid
    function populateAttachmentsGrid() {
        const allAttachments = collectAllAttachments();

        // Update count badge
        if (allAttachments.length > 0) {
            attachmentsCountBadge.textContent = allAttachments.length;
            attachmentsCountBadge.style.display = 'inline-block';
        } else {
            attachmentsCountBadge.style.display = 'none';
        }

        // Show empty state if no attachments
        if (allAttachments.length === 0) {
            attachmentsGrid.style.display = 'none';
            noAttachments.style.display = 'flex';
            return;
        }

        attachmentsGrid.style.display = 'grid';
        noAttachments.style.display = 'none';

        // Clear existing content
        attachmentsGrid.textContent = '';

        // Render attachment cards using DOM methods
        // Action buttons have their own click handlers built-in
        allAttachments.forEach((att, idx) => {
            const card = createAttachmentCard(att, idx);
            attachmentsGrid.appendChild(card);
        });
    }

    // Update attachments count when switching chats (called from openChat and loadMessages)
    function updateAttachmentsCount() {
        const allAttachments = collectAllAttachments();
        if (allAttachments.length > 0) {
            attachmentsCountBadge.textContent = allAttachments.length;
            attachmentsCountBadge.style.display = 'inline-block';
        } else {
            attachmentsCountBadge.style.display = 'none';
        }
    }

    // ==========================================
    // END TAB NAVIGATION & ATTACHMENTS
    // ==========================================

    // ==========================================
    // TRASH / RECYCLE BIN
    // ==========================================

    // Initialize trash: purge expired items and update count
    async function initializeTrash() {
        try {
            // Purge expired trash items (older than 60 days)
            const purgedCount = await window.TeamsBackupDB.purgeExpiredTrash();
            if (purgedCount > 0) {
                console.log(`Purged ${purgedCount} expired chats from trash`);
            }
            // Update trash count badge
            await updateTrashCount();
        } catch (error) {
            console.error('Error initializing trash:', error);
        }
    }

    // Update trash count badge in sidebar
    async function updateTrashCount() {
        try {
            const trashedChats = await window.TeamsBackupDB.getTrashedChats();
            if (trashedChats.length > 0) {
                trashCount.textContent = trashedChats.length;
                trashCount.style.display = 'inline-block';
            } else {
                trashCount.style.display = 'none';
            }
        } catch (error) {
            console.error('Error updating trash count:', error);
        }
    }

    // Open trash panel
    async function openTrashPanel() {
        try {
            const trashedChats = await window.TeamsBackupDB.getTrashedChats();

            if (trashedChats.length === 0) {
                trashList.style.display = 'none';
                trashEmpty.style.display = 'flex';
            } else {
                trashEmpty.style.display = 'none';
                trashList.style.display = 'block';
                renderTrashList(trashedChats);
            }

            trashPanel.style.display = 'flex';
        } catch (error) {
            console.error('Error opening trash panel:', error);
        }
    }

    // Render trash list
    function renderTrashList(trashedChats) {
        trashList.innerHTML = trashedChats.map(chat => {
            const daysText = chat.daysRemaining === 1 ? '1 day left' : `${chat.daysRemaining} days left`;
            const trashedDate = new Date(chat.trashedAt).toLocaleDateString();

            return `
                <div class="trash-item" data-id="${chat.id}">
                    <div class="trash-item__avatar" style="background: ${window.TeamsBackupDB.stringToColor(chat.name)}">
                        ${window.TeamsBackupDB.getInitials(chat.name)}
                    </div>
                    <div class="trash-item__info">
                        <div class="trash-item__name">${escapeHtml(chat.name)}</div>
                        <div class="trash-item__meta">
                            <span>Deleted ${trashedDate}</span>
                            <span class="trash-item__expires">
                                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                                    <circle cx="12" cy="12" r="10"></circle>
                                    <polyline points="12 6 12 12 16 14"></polyline>
                                </svg>
                                ${daysText}
                            </span>
                        </div>
                    </div>
                    <div class="trash-item__actions">
                        <button class="trash-item__btn trash-item__btn--restore" data-action="restore" data-id="${chat.id}" title="Restore to chat list">
                            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                                <polyline points="1 4 1 10 7 10"></polyline>
                                <path d="M3.51 15a9 9 0 1 0 2.13-9.36L1 10"></path>
                            </svg>
                            Restore
                        </button>
                        <button class="trash-item__btn trash-item__btn--delete" data-action="delete" data-id="${chat.id}" title="Delete permanently">
                            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                                <polyline points="3 6 5 6 21 6"></polyline>
                                <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
                                <line x1="10" y1="11" x2="10" y2="17"></line>
                                <line x1="14" y1="11" x2="14" y2="17"></line>
                            </svg>
                        </button>
                    </div>
                </div>
            `;
        }).join('');

        // Add click handlers for trash item actions
        trashList.querySelectorAll('.trash-item__btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                const action = btn.dataset.action;
                const chatId = btn.dataset.id;
                showTrashActionModal(chatId, action);
            });
        });
    }

    // Show trash action confirmation modal
    function showTrashActionModal(chatId, action) {
        trashActionChatId = chatId;
        trashActionType = action;

        if (action === 'restore') {
            trashActionTitle.textContent = 'Restore Chat?';
            trashActionMessage.textContent = 'This chat will be restored to your chat list.';
            confirmTrashActionBtn.textContent = 'Restore';
            confirmTrashActionBtn.className = 'modal__btn modal__btn--primary';
        } else {
            trashActionTitle.textContent = 'Delete Permanently?';
            trashActionMessage.textContent = 'This action cannot be undone. The backup will be permanently deleted.';
            confirmTrashActionBtn.textContent = 'Delete Forever';
            confirmTrashActionBtn.className = 'modal__btn modal__btn--danger';
        }

        trashActionModal.style.display = 'flex';
    }

    // Close trash panel
    function closeTrashPanel() {
        trashPanel.style.display = 'none';
    }

    // Trash panel event handlers
    if (trashToggleBtn) {
        trashToggleBtn.addEventListener('click', openTrashPanel);
    }

    if (closeTrashBtn) {
        closeTrashBtn.addEventListener('click', closeTrashPanel);
    }

    // Close trash panel on backdrop click
    if (trashPanel) {
        trashPanel.querySelector('.trash-panel__backdrop').addEventListener('click', closeTrashPanel);
    }

    // Trash action modal handlers
    if (cancelTrashActionBtn) {
        cancelTrashActionBtn.addEventListener('click', () => {
            trashActionModal.style.display = 'none';
            trashActionChatId = null;
            trashActionType = null;
        });
    }

    if (confirmTrashActionBtn) {
        confirmTrashActionBtn.addEventListener('click', async () => {
            if (!trashActionChatId || !trashActionType) return;

            try {
                if (trashActionType === 'restore') {
                    await window.TeamsBackupDB.restoreChat(trashActionChatId);
                    console.log('Chat restored:', trashActionChatId);
                } else {
                    await window.TeamsBackupDB.permanentlyDeleteChat(trashActionChatId);
                    console.log('Chat permanently deleted:', trashActionChatId);
                }

                // Refresh both chat list and trash panel
                await loadChats();
                await updateTrashCount();
                await openTrashPanel(); // Refresh trash list

                trashActionModal.style.display = 'none';
                trashActionChatId = null;
                trashActionType = null;

            } catch (error) {
                console.error('Error performing trash action:', error);
                if (error?.code === 'chat_limit') {
                    trashActionModal.style.display = 'none';
                    trashActionChatId = null;
                    trashActionType = null;
                    showUpgradeModal(error.message, 'Free plan: one backup at a time');
                    return;
                }
                alert('Error: ' + error.message);
            }
        });
    }

    // Close trash action modal on backdrop click
    if (trashActionModal) {
        trashActionModal.querySelector('.modal__backdrop').addEventListener('click', () => {
            trashActionModal.style.display = 'none';
            trashActionChatId = null;
            trashActionType = null;
        });
    }

    // ==========================================
    // END TRASH / RECYCLE BIN
    // ==========================================

    // ==========================================
    // EXPORT FUNCTIONS
    // ==========================================

    function shouldApplyFreeExportBranding() {
        return !isPro;
    }

    function getFreeExportPromoText() {
        return 'Free version export. Update to Pro version for unlimited backups and watermark-free downloads.';
    }

    function buildPlainTextExport(chat, messages) {
        const lines = [
            'Teams Chat Backup',
            `Chat: ${(chat && chat.name) || 'Unknown Chat'}`,
            `Messages: ${(messages || []).length}`,
            `Exported: ${new Date().toLocaleString()}`,
            ''
        ];

        const groups = groupMessagesByDate(messages || []);
        for (const [dateLabel, group] of Object.entries(groups)) {
            lines.push(`=== ${dateLabel} ===`);
            for (const msg of group.messages) {
                if (msg.type === 'system' || msg.isSystemMessage) {
                    lines.push(`[System] ${msg.content || msg.text || ''}`);
                    lines.push('');
                    continue;
                }

                const author = msg.author || 'Unknown';
                const time = extractTime(msg.timestamp);
                const isDeleted = msg.isDeleted || msg.deletedAt;
                const isEdited = msg.isEdited || (msg.editHistory && msg.editHistory.length > 0);
                const attachments = Array.isArray(msg.attachments)
                    ? msg.attachments
                    : (Array.isArray(msg.files) ? msg.files : []);
                let content = isDeleted ? '[This message was deleted]' : (msg.content || msg.text || '');
                if (!content) {
                    content = '(no text)';
                }
                if (isEdited && !isDeleted) {
                    content += ' (edited)';
                }

                lines.push(`${author} [${time}]`);
                lines.push(content);

                if (msg.isStarred) {
                    lines.push('[Starred]');
                }

                if (!isDeleted && attachments.length > 0) {
                    for (const file of attachments) {
                        lines.push(`[File] ${file.filename || file.name || 'File attachment'}`);
                    }
                }

                if (!isDeleted && msg.images && msg.images.length > 0) {
                    lines.push(`[Images] ${msg.images.length} attachment${msg.images.length === 1 ? '' : 's'}`);
                }

                lines.push('');
            }
        }

        return lines.join('\n').trim();
    }

    function generateTxtExport(chat, messages) {
        const hasMessages = Array.isArray(messages) && messages.length > 0;
        const baseText = hasMessages
            ? buildPlainTextExport(chat, messages)
            : ((chat && typeof chat.textContent === 'string' && chat.textContent.trim().length > 0)
                ? chat.textContent.trimEnd()
                : buildPlainTextExport(chat, messages));

        if (!shouldApplyFreeExportBranding()) {
            return baseText;
        }

        const promoLine = `[Teams Chat Backup] ${getFreeExportPromoText()}`;
        const sourceLines = baseText ? baseText.split(/\r?\n/) : [];
        const output = [promoLine, ''];
        let nonEmptyLineCount = 0;

        for (const line of sourceLines) {
            output.push(line);
            if (line.trim()) {
                nonEmptyLineCount++;
            }

            if (nonEmptyLineCount > 0 && nonEmptyLineCount % 80 === 0) {
                output.push('');
                output.push(promoLine);
                output.push('');
            }
        }

        if (output[output.length - 1] !== '') {
            output.push('');
        }
        output.push(promoLine);
        return output.join('\n');
    }

    // Generate self-contained HTML export
    function generateHtmlExport(chat, messages) {
        const groups = groupMessagesByDate(messages);
        const exportDate = new Date().toLocaleString();
        const avatarColor = window.TeamsBackupDB.stringToColor(chat.name);
        const initials = window.TeamsBackupDB.getInitials(chat.name);
        const isFreeExport = shouldApplyFreeExportBranding();
        const freePromoText = getFreeExportPromoText();

        let messagesHtml = '';
        for (const [dateLabel, group] of Object.entries(groups)) {
            messagesHtml += `<div class="date-divider"><span>${escapeHtml(dateLabel)}</span></div>`;
            for (const msg of group.messages) {
                if (msg.type === 'system' || msg.isSystemMessage) {
                    messagesHtml += `<div class="message message--system"><em>${escapeHtml(msg.content || msg.text || '')}</em></div>`;
                    continue;
                }
                const author = msg.author || 'Unknown';
                const time = extractTime(msg.timestamp);
                const msgColor = window.TeamsBackupDB.stringToColor(author);
                const msgInitials = window.TeamsBackupDB.getInitials(author);
                const isDeleted = msg.isDeleted || msg.deletedAt;
                const isEdited = msg.isEdited || (msg.editHistory && msg.editHistory.length > 0);
                const isStarred = msg.isStarred;

                let contentHtml = '';
                if (isDeleted) {
                    contentHtml = '<em class="deleted-text">This message was deleted</em>';
                } else {
                    contentHtml = escapeHtml(msg.content || msg.text || '');
                    contentHtml = contentHtml.replace(/\n/g, '<br>');
                    if (isEdited) {
                        contentHtml += ' <span class="edited-tag">(edited)</span>';
                    }
                }

                // Images
                let imagesHtml = '';
                if (!isDeleted && msg.images && msg.images.length > 0) {
                    imagesHtml = '<div class="message-images">';
                    for (const img of msg.images) {
                        if (isAnimatedEmojiImage(img)) {
                            imagesHtml += renderAnimatedEmojiSprite(img);
                            continue;
                        }
                        const src = img.dataUrl || img.src || img;
                        if (isEmbeddedImageDataUrl(src)) {
                            imagesHtml += `<img src="${src}" alt="Image" class="message-image" />`;
                        }
                    }
                    imagesHtml += '</div>';
                }

                // File attachments (backups store them as msg.attachments; older data used msg.files)
                let filesHtml = '';
                const fileAttachments = Array.isArray(msg.attachments)
                    ? msg.attachments
                    : (Array.isArray(msg.files) ? msg.files : []);
                if (!isDeleted && fileAttachments.length > 0) {
                    filesHtml = '<div class="message-files">';
                    for (const file of fileAttachments) {
                        const fileName = file.filename || file.name;
                        if (file.dataUrl) {
                            filesHtml += `<a href="${escapeHtml(file.dataUrl)}" download="${escapeHtml(fileName || 'file')}" class="file-link">${escapeHtml(fileName || 'Download file')}</a>`;
                        } else {
                            filesHtml += `<span class="file-link">${escapeHtml(fileName || 'File attachment')}</span>`;
                        }
                    }
                    filesHtml += '</div>';
                }

                const starClass = isStarred ? ' message--starred' : '';
                const deletedClass = isDeleted ? ' message--deleted' : '';

                messagesHtml += `
                <div class="message${starClass}${deletedClass}">
                    <div class="avatar" style="background:${msgColor}">${escapeHtml(msgInitials)}</div>
                    <div class="message-body">
                        <div class="message-header">
                            <strong class="author">${escapeHtml(author)}</strong>
                            <span class="time">${escapeHtml(time)}</span>
                            ${isStarred ? '<span class="star-icon" title="Starred">&#9733;</span>' : ''}
                        </div>
                        <div class="message-content">${contentHtml}</div>
                        ${imagesHtml}
                        ${filesHtml}
                    </div>
                </div>`;
            }
        }

        return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>${escapeHtml(chat.name)} - Chat Backup</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;background:#f5f5f5;color:#252423;line-height:1.5;position:relative}
body.free-export::before{content:'FREE VERSION';position:fixed;top:48%;left:50%;transform:translate(-50%,-50%) rotate(-24deg);font-size:72px;font-weight:800;letter-spacing:.32em;color:rgba(70,71,117,.08);pointer-events:none;z-index:0;white-space:nowrap}
.header,.container,.footer,.free-export-banner{position:relative;z-index:1}
.header{background:linear-gradient(135deg,#6264A7,#464775);color:#fff;padding:32px 24px;display:flex;align-items:center;gap:16px}
.header-avatar{width:56px;height:56px;border-radius:50%;background:${avatarColor};display:flex;align-items:center;justify-content:center;font-size:22px;font-weight:700;flex-shrink:0}
.header-info h1{font-size:22px;font-weight:600}
.header-info p{opacity:.8;font-size:14px}
.free-export-banner{max-width:800px;margin:16px auto 0;padding:12px 16px;border:1px solid rgba(240,180,41,.45);border-radius:12px;background:rgba(240,180,41,.14);color:#7a4b00;font-size:13px;font-weight:600}
.container{max-width:800px;margin:0 auto;padding:20px 16px}
.date-divider{text-align:center;margin:24px 0 16px;position:relative}
.date-divider::before{content:'';position:absolute;top:50%;left:0;right:0;height:1px;background:#e1dfdd}
.date-divider span{position:relative;background:#f5f5f5;padding:0 12px;font-size:12px;color:#605e5c;font-weight:600}
.message{display:flex;gap:12px;margin-bottom:16px;padding:8px 12px;border-radius:8px}
.message:hover{background:#fff}
.message--starred{border-left:3px solid #f0b429;padding-left:9px}
.message--deleted{opacity:.6}
.message--deleted .message-content{font-style:italic;border:1px dashed #a19f9d;padding:6px 10px;border-radius:6px}
.message--system{justify-content:center;color:#605e5c;font-size:13px;padding:8px 0}
.avatar{width:36px;height:36px;border-radius:50%;display:flex;align-items:center;justify-content:center;color:#fff;font-weight:600;font-size:13px;flex-shrink:0}
.message-body{flex:1;min-width:0}
.message-header{display:flex;align-items:baseline;gap:8px;margin-bottom:2px}
.author{font-size:14px;color:#252423}
.time{font-size:12px;color:#a19f9d}
.star-icon{color:#f0b429;font-size:14px}
.message-content{font-size:14px;word-break:break-word}
.edited-tag{color:#a19f9d;font-size:12px;font-style:italic}
.deleted-text{color:#a19f9d}
.message-images{margin-top:8px}
.message-image{max-width:400px;max-height:300px;border-radius:8px;margin:4px 4px 4px 0;cursor:pointer}
.emoji-sprite{display:inline-block;background-repeat:no-repeat;background-position:0 0;vertical-align:middle}
@keyframes emoji-sprite-play{from{background-position:0 0}to{background-position:0 var(--emoji-sprite-travel,0px)}}
.message-files{margin-top:6px;display:flex;flex-wrap:wrap;gap:6px}
.file-link{display:inline-block;padding:4px 10px;background:#e5e5f7;border-radius:6px;font-size:13px;color:#6264A7;text-decoration:none}
.file-link:hover{background:#d5d5ef}
.footer{text-align:center;padding:24px 16px;color:#a19f9d;font-size:12px;border-top:1px solid #e1dfdd;margin-top:24px}
@media print{body{background:#fff}.message:hover{background:transparent}.header{-webkit-print-color-adjust:exact;print-color-adjust:exact}}
@media(prefers-color-scheme:dark){
body{background:#1a1a1a;color:#e0e0e0}
body.free-export::before{color:rgba(255,255,255,.06)}
.date-divider::before{background:#444}
.date-divider span{background:#1a1a1a;color:#aaa}
.message:hover{background:#252525}
.message--deleted .message-content{border-color:#555}
.author{color:#e0e0e0}
.free-export-banner{background:rgba(240,180,41,.16);border-color:rgba(240,180,41,.35);color:#ffd87a}
.file-link{background:#333;color:#8B8CC7}
.file-link:hover{background:#444}
.footer{border-top-color:#333;color:#666}
}
</style>
</head>
<body class="${isFreeExport ? 'free-export' : ''}">
${isFreeExport ? `<div class="free-export-banner">${escapeHtml(freePromoText)}</div>` : ''}
<div class="header">
    <div class="header-avatar" style="background:${avatarColor}">${escapeHtml(initials)}</div>
    <div class="header-info">
        <h1>${escapeHtml(chat.name)}</h1>
        <p>${messages.length} messages &middot; Exported ${escapeHtml(exportDate)}</p>
    </div>
</div>
<div class="container">
${messagesHtml}
</div>
<div class="footer">${escapeHtml(isFreeExport ? 'Generated by Teams Chat Backup Extension · Free Version' : 'Generated by Teams Chat Backup Extension')}</div>
</body>
</html>`;
    }

    // jsPDF's built-in Helvetica only covers WinAnsi (Latin-1 plus a few typographic marks).
    // Anything else prints as garbage, so emoji become readable [names] and the rest becomes '?'.
    const PDF_EMOJI_NAMES = {
        '😀': 'grinning', '😃': 'smiley', '😄': 'smile', '😁': 'grin', '😆': 'laughing', '😅': 'sweat smile',
        '😂': 'joy', '🤣': 'rofl', '🙂': 'slight smile', '🙃': 'upside down', '😉': 'wink', '😊': 'blush',
        '😇': 'innocent', '🥰': 'smiling hearts', '😍': 'heart eyes', '😘': 'kiss', '😋': 'yum', '😛': 'tongue',
        '😜': 'winking tongue', '🤪': 'zany', '🤔': 'thinking', '🤗': 'hug', '🤭': 'hand over mouth', '🤫': 'shush',
        '😐': 'neutral', '😑': 'expressionless', '😶': 'no mouth', '😏': 'smirk', '😒': 'unamused', '🙄': 'eye roll',
        '😬': 'grimace', '😌': 'relieved', '😔': 'pensive', '😴': 'sleeping', '😷': 'mask', '🤒': 'sick',
        '🤯': 'mind blown', '🥳': 'party face', '😎': 'sunglasses', '🤓': 'nerd', '😕': 'confused', '😟': 'worried',
        '🙁': 'frown', '😮': 'open mouth', '😲': 'astonished', '😳': 'flushed', '🥺': 'pleading', '😢': 'cry',
        '😭': 'sob', '😱': 'scream', '😖': 'confounded', '😞': 'disappointed', '😓': 'sweat', '😩': 'weary',
        '😫': 'tired', '😤': 'triumph', '😡': 'angry', '😠': 'mad', '🤬': 'cursing', '💀': 'skull', '💩': 'poop',
        '🤡': 'clown', '👻': 'ghost', '👽': 'alien', '🤖': 'robot', '😺': 'cat smile', '🙈': 'see no evil',
        '👋': 'wave', '🤚': 'raised back of hand', '✋': 'raised hand', '🖐': 'hand', '👌': 'ok hand', '🤌': 'pinched fingers',
        '✌': 'victory', '🤞': 'fingers crossed', '🤟': 'love you', '🤘': 'rock on', '🤙': 'call me', '👈': 'point left',
        '👉': 'point right', '👆': 'point up', '👇': 'point down', '☝': 'point up', '👍': 'thumbs up',
        '👎': 'thumbs down', '✊': 'fist', '👊': 'punch', '👏': 'clap', '🙌': 'raised hands', '👐': 'open hands',
        '🤝': 'handshake', '🙏': 'pray', '✍': 'writing', '💪': 'muscle', '👀': 'eyes', '🧠': 'brain',
        '❤': 'heart', '🧡': 'orange heart', '💛': 'yellow heart', '💚': 'green heart', '💙': 'blue heart',
        '💜': 'purple heart', '🖤': 'black heart', '🤍': 'white heart', '💔': 'broken heart', '💕': 'two hearts',
        '💖': 'sparkling heart', '💯': '100', '💥': 'boom', '💫': 'dizzy', '💬': 'speech bubble', '💤': 'zzz',
        '🔥': 'fire', '✨': 'sparkles', '⭐': 'star', '🌟': 'glowing star', '⚡': 'zap', '☀': 'sun', '🌈': 'rainbow',
        '☕': 'coffee', '🍕': 'pizza', '🍔': 'burger', '🍰': 'cake', '🎂': 'birthday cake', '🍺': 'beer', '🍻': 'cheers',
        '🥂': 'clink', '🍾': 'champagne', '🎉': 'tada', '🎊': 'confetti', '🎈': 'balloon', '🎁': 'gift', '🏆': 'trophy',
        '🥇': 'gold medal', '🎯': 'bullseye', '🚀': 'rocket', '✈': 'airplane', '🚗': 'car', '🏠': 'house',
        '💻': 'laptop', '🖥': 'desktop', '📱': 'phone', '☎': 'telephone', '📞': 'phone call', '📧': 'email', '✉': 'envelope',
        '📅': 'calendar', '📆': 'calendar', '📌': 'pin', '📎': 'paperclip', '📝': 'memo', '📄': 'page', '📁': 'folder',
        '📊': 'bar chart', '📈': 'chart up', '📉': 'chart down', '🔒': 'lock', '🔑': 'key', '🔔': 'bell', '💡': 'idea',
        '⏰': 'alarm clock', '⏳': 'hourglass', '⌛': 'hourglass', '💰': 'money bag', '💵': 'dollar', '🛠': 'tools',
        '✅': 'check', '✔': 'check', '☑': 'checked box', '❌': 'x', '✖': 'x', '❎': 'x', '❗': 'exclamation',
        '❓': 'question', '‼': 'double exclamation', '⁉': 'interrobang', '⚠': 'warning', '🚫': 'no entry', '⛔': 'no entry',
        '🔴': 'red circle', '🟢': 'green circle', '🟡': 'yellow circle', '🔵': 'blue circle', '➡': 'right arrow',
        '⬅': 'left arrow', '⬆': 'up arrow', '⬇': 'down arrow', '🆗': 'ok', '🆕': 'new', '🙋': 'raising hand',
        '🤷': 'shrug', '🤦': 'facepalm', '🙆': 'ok gesture', '🙅': 'no gesture', '👨‍💻': 'technologist',
        '👩‍💻': 'technologist', '🧑‍💻': 'technologist', '👨': 'man', '👩': 'woman', '🧑': 'person', '👪': 'family',
        '👨‍👩‍👧': 'family', '👨‍👩‍👦': 'family', '👨‍👩‍👧‍👦': 'family'
    };

    function toPdfText(value, dropEmoji) {
        if (value === null || value === undefined) return '';
        var emojiPattern = /\p{RI}\p{RI}|[#*0-9]️?⃣|\p{Extended_Pictographic}[️\p{EMod}]*(?:‍\p{Extended_Pictographic}[️\p{EMod}]*)*/gu;
        var text = String(value).normalize('NFC').replace(emojiPattern, function(match) {
            // (c), (R) and (TM) are pictographic but Helvetica has them.
            if (/^[©®™]$/u.test(match)) return match;
            if (dropEmoji) return '';
            if (/^\p{RI}\p{RI}$/u.test(match)) {
                return '[flag ' + Array.from(match).map(function(ch) {
                    return String.fromCharCode(ch.codePointAt(0) - 0x1F1E6 + 65);
                }).join('') + ']';
            }
            if (/⃣$/.test(match)) return '[' + match[0] + ']';
            var plain = match.replace(/[️\p{EMod}]/gu, '');
            var name = PDF_EMOJI_NAMES[plain] || PDF_EMOJI_NAMES[Array.from(plain)[0]];
            return '[' + (name || 'emoji') + ']';
        });
        text = text
            .replace(/[★☆]/g, '*')
            .replace(/[→➡]/g, '->')
            .replace(/←/g, '<-')
            .replace(/[​-‏⁠﻿︎️]/g, '')
            .replace(/[ -   　]/g, ' ')
            .replace(/[‐‑‒−]/g, '-');
        // WinAnsi = printable ASCII, Latin-1, and the typographic marks jsPDF maps into 0x80-0x9F.
        return text.replace(/[^\t\n\r\x20-\x7E -ÿŒœŠšŸŽžƒˆ˜–—‘’‚“”„†‡•…‰‹›€™]/gu, '?');
    }

    // Generate styled PDF from message data using jsPDF
    async function generateStyledPdf(chat, messages) {
        const { jsPDF } = window.jspdf;
        const pdf = new jsPDF('p', 'pt', 'a4');
        const pageWidth = pdf.internal.pageSize.getWidth();
        const pageHeight = pdf.internal.pageSize.getHeight();
        const margin = 40;
        const contentWidth = pageWidth - margin * 2;
        let y = margin;
        let pageNum = 1;

        function addPageBranding(pageNumber) {
            pdf.setPage(pageNumber);

            if (shouldApplyFreeExportBranding()) {
                pdf.setFont('helvetica', 'bold');
                pdf.setFontSize(42);
                pdf.setTextColor(236, 236, 236);
                pdf.text('FREE VERSION', pageWidth / 2, pageHeight * 0.52, { align: 'center' });

                pdf.setFont('helvetica', 'normal');
                pdf.setFontSize(8);
                pdf.setTextColor(210, 210, 210);
                pdf.text('Update to Pro version for unlimited backups and watermark-free downloads.', pageWidth / 2, pageHeight - 34, { align: 'center' });
            }

            pdf.setFont('helvetica', 'normal');
            pdf.setFontSize(9);
            pdf.setTextColor(160, 160, 160);
            pdf.text('Page ' + pageNumber, pageWidth / 2, pageHeight - 20, { align: 'center' });
            pdf.text(
                shouldApplyFreeExportBranding() ? 'Teams Chat Backup · Free Version' : 'Teams Chat Backup',
                pageWidth - margin,
                pageHeight - 20,
                { align: 'right' }
            );
        }

        function checkNewPage(needed) {
            if (y + needed > pageHeight - 50) {
                pdf.addPage();
                pageNum++;
                y = margin;
            }
            return y;
        }

        function getImageSize(dataUrl) {
            return new Promise(function(resolve) {
                var img = new Image();
                img.onload = function() { resolve({ width: img.naturalWidth, height: img.naturalHeight }); };
                img.onerror = function() { resolve(null); };
                img.src = dataUrl;
            });
        }

        function hexToRgb(hex) {
            var result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
            return result ? {
                r: parseInt(result[1], 16),
                g: parseInt(result[2], 16),
                b: parseInt(result[3], 16)
            } : { r: 100, g: 100, b: 100 };
        }

        // Cover header
        pdf.setFillColor(98, 100, 167);
        pdf.rect(0, 0, pageWidth, 100, 'F');

        var avatarColorRgb = hexToRgb(window.TeamsBackupDB.stringToColor(chat.name));
        pdf.setFillColor(avatarColorRgb.r, avatarColorRgb.g, avatarColorRgb.b);
        pdf.circle(margin + 22, 50, 22, 'F');
        pdf.setTextColor(255, 255, 255);
        pdf.setFontSize(16);
        pdf.setFont('helvetica', 'bold');
        var coverInitials = toPdfText(window.TeamsBackupDB.getInitials(toPdfText(chat.name, true)));
        pdf.text(coverInitials, margin + 22, 55, { align: 'center' });

        pdf.setFontSize(20);
        pdf.setFont('helvetica', 'bold');
        pdf.text(toPdfText(chat.name), margin + 56, 42, { maxWidth: contentWidth - 60 });

        pdf.setFontSize(11);
        pdf.setFont('helvetica', 'normal');
        pdf.setTextColor(220, 220, 240);
        var subtitle = messages.length + ' messages \u2022 ' + new Date().toLocaleDateString();
        pdf.text(subtitle, margin + 56, 62);

        y = 120;

        var groups = groupMessagesByDate(messages);

        for (var dateLabel in groups) {
            if (!groups.hasOwnProperty(dateLabel)) continue;
            var group = groups[dateLabel];

            // Date divider
            checkNewPage(30);
            pdf.setDrawColor(200, 200, 200);
            pdf.setLineWidth(0.5);
            var dividerY = y + 8;
            pdf.line(margin, dividerY, margin + contentWidth, dividerY);
            pdf.setFontSize(10);
            pdf.setFont('helvetica', 'bold');
            pdf.setTextColor(130, 130, 130);
            var dateLabelWidth = pdf.getTextWidth(dateLabel);
            var dateLabelX = pageWidth / 2;
            pdf.setFillColor(255, 255, 255);
            pdf.rect(dateLabelX - dateLabelWidth / 2 - 6, dividerY - 6, dateLabelWidth + 12, 12, 'F');
            pdf.text(dateLabel, dateLabelX, dividerY + 3, { align: 'center' });
            y = dividerY + 20;

            for (var mi = 0; mi < group.messages.length; mi++) {
                var msg = group.messages[mi];

                // System message
                if (msg.type === 'system' || msg.isSystemMessage) {
                    checkNewPage(20);
                    pdf.setFontSize(10);
                    pdf.setFont('helvetica', 'italic');
                    pdf.setTextColor(140, 140, 140);
                    var sysText = toPdfText(msg.content || msg.text || '');
                    var sysLines = pdf.splitTextToSize(sysText, contentWidth - 40);
                    pdf.text(sysLines, pageWidth / 2, y, { align: 'center' });
                    y += sysLines.length * 13 + 10;
                    continue;
                }

                var author = toPdfText(msg.author || 'Unknown');
                var time = extractTime(msg.timestamp);
                var isDeleted = msg.isDeleted || msg.deletedAt;
                var isEdited = msg.isEdited || (msg.editHistory && msg.editHistory.length > 0);
                var isStarred = msg.isStarred;
                var msgContent = isDeleted ? '[This message was deleted]' : toPdfText(msg.content || msg.text || '');

                var wrappedText = pdf.splitTextToSize(msgContent, contentWidth - 56);
                var neededHeight = 20 + wrappedText.length * 14 + 10;

                var imageDataUrls = [];
                if (!isDeleted && msg.images && msg.images.length > 0) {
                    for (var ii = 0; ii < msg.images.length; ii++) {
                        var imgItem = msg.images[ii];
                        var src = imgItem.dataUrl || imgItem.src || imgItem;
                        if (isEmbeddedImageDataUrl(src)) {
                            imageDataUrls.push(src);
                        }
                    }
                    neededHeight += imageDataUrls.length * 120;
                }

                checkNewPage(Math.min(neededHeight, pageHeight - 100));

                // Avatar circle
                var authorColorRgb = hexToRgb(window.TeamsBackupDB.stringToColor(author));
                pdf.setFillColor(authorColorRgb.r, authorColorRgb.g, authorColorRgb.b);
                pdf.circle(margin + 14, y + 10, 14, 'F');
                pdf.setTextColor(255, 255, 255);
                pdf.setFontSize(10);
                pdf.setFont('helvetica', 'bold');
                var msgInitials = toPdfText(window.TeamsBackupDB.getInitials(toPdfText(msg.author || 'Unknown', true)));
                pdf.text(msgInitials, margin + 14, y + 13, { align: 'center' });

                // Starred indicator
                if (isStarred) {
                    // Drawn, because Helvetica has no star glyph.
                    var starCx = margin + contentWidth - 7;
                    var starCy = y + 8;
                    var starPoints = [];
                    for (var sp = 0; sp < 10; sp++) {
                        var starRadius = sp % 2 === 0 ? 7 : 2.9;
                        var starAngle = -Math.PI / 2 + sp * Math.PI / 5;
                        starPoints.push([starCx + starRadius * Math.cos(starAngle), starCy + starRadius * Math.sin(starAngle)]);
                    }
                    var starDeltas = [];
                    for (var sd = 1; sd < starPoints.length; sd++) {
                        starDeltas.push([starPoints[sd][0] - starPoints[sd - 1][0], starPoints[sd][1] - starPoints[sd - 1][1]]);
                    }
                    pdf.setFillColor(240, 180, 41);
                    pdf.lines(starDeltas, starPoints[0][0], starPoints[0][1], [1, 1], 'F', true);
                }

                // Author + time
                var textX = margin + 38;
                pdf.setFontSize(12);
                pdf.setFont('helvetica', 'bold');
                pdf.setTextColor(37, 36, 35);
                pdf.text(author, textX, y + 10);
                var authorWidth = pdf.getTextWidth(author);

                pdf.setFontSize(10);
                pdf.setFont('helvetica', 'normal');
                pdf.setTextColor(160, 160, 160);
                pdf.text(time, textX + authorWidth + 8, y + 10);

                y += 20;

                // Message content
                pdf.setFontSize(11);
                pdf.setTextColor(37, 36, 35);
                if (isDeleted) {
                    pdf.setFont('helvetica', 'italic');
                    pdf.setTextColor(160, 160, 160);
                } else {
                    pdf.setFont('helvetica', 'normal');
                }
                for (var li = 0; li < wrappedText.length; li++) {
                    checkNewPage(16);
                    pdf.text(wrappedText[li], textX, y);
                    y += 14;
                }

                // Edited indicator
                if (isEdited && !isDeleted) {
                    pdf.setFontSize(9);
                    pdf.setFont('helvetica', 'italic');
                    pdf.setTextColor(160, 160, 160);
                    pdf.text('(edited)', textX, y);
                    y += 12;
                }

                // Images
                for (var ij = 0; ij < imageDataUrls.length; ij++) {
                    try {
                        var size = await getImageSize(imageDataUrls[ij]);
                        if (size) {
                            var maxW = 350;
                            var maxH = 250;
                            var w = size.width;
                            var h = size.height;
                            if (w > maxW) { h = h * (maxW / w); w = maxW; }
                            if (h > maxH) { w = w * (maxH / h); h = maxH; }

                            checkNewPage(h + 10);
                            var mimeMatch = imageDataUrls[ij].match(/^data:image\/(\w+);/);
                            var mimeType = mimeMatch ? mimeMatch[1].toUpperCase() : 'JPEG';
                            var supportedFormats = ['PNG', 'JPEG', 'WEBP', 'GIF'];
                            var format = supportedFormats.indexOf(mimeType) !== -1 ? mimeType : 'JPEG';
                            pdf.addImage(imageDataUrls[ij], format, textX, y, w, h);
                            y += h + 8;
                        }
                    } catch (e) {
                        // Skip failed images
                    }
                }

                // File attachments (backups store them as msg.attachments; older data used msg.files)
                var pdfFileAttachments = Array.isArray(msg.attachments)
                    ? msg.attachments
                    : (Array.isArray(msg.files) ? msg.files : []);
                if (!isDeleted && pdfFileAttachments.length > 0) {
                    for (var fi = 0; fi < pdfFileAttachments.length; fi++) {
                        checkNewPage(18);
                        pdf.setFontSize(10);
                        pdf.setFont('helvetica', 'normal');
                        pdf.setTextColor(98, 100, 167);
                        pdf.text('[File] ' + toPdfText(pdfFileAttachments[fi].filename || pdfFileAttachments[fi].name || 'File attachment'), textX, y);
                        y += 16;
                    }
                }

                y += 8;
            }
        }

        for (var pageIndex = 1; pageIndex <= pageNum; pageIndex++) {
            addPageBranding(pageIndex);
        }
        pdf.setPage(pageNum);
        return pdf.output('blob');
    }

    // ==========================================
    // END EXPORT FUNCTIONS
    // ==========================================

    // Utility functions
    function escapeHtml(text) {
        if (!text) return '';
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    function formatDate(dateString) {
        const date = new Date(dateString);
        const now = new Date();
        const diffMs = now - date;
        const diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));

        if (diffDays === 0) {
            return 'Today ' + date.toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit' });
        } else if (diffDays === 1) {
            return 'Yesterday';
        } else if (diffDays < 7) {
            return diffDays + ' days ago';
        } else {
            return date.toLocaleDateString();
        }
    }

    function sanitizeFilename(name) {
        // Only remove characters that are unsafe for filenames across operating systems
        // Preserves Unicode characters (Korean, Chinese, Japanese, etc.)
        // Unsafe: \ / : * ? " < > | and control characters
        return name
            .replace(/[\\/:*?"<>|]/g, '_')  // Replace unsafe characters with underscore
            .replace(/[\x00-\x1f]/g, '')     // Remove control characters
            .trim()
            .substring(0, 100);              // Allow longer names for international text
    }
});
