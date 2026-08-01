import React, { useState, useEffect, useMemo } from 'react';
import { createRoot } from 'react-dom/client';
import { Box } from '@mui/material';
import { DynamicsAction, ExtensionDisplayMode } from '#types/global';
import { messageService } from '#services/MessageService';
import { ExtensionConfigService, ExtensionConfig } from '#services/ExtensionConfigService';
import { ThemeProvider } from '#contexts/ThemeContext';
import { checkDynamicsViaXrm } from '#utils/dynamicsDetection';
import { formActions, navigationActions, debuggingActions } from '#config/actions';
import ThemeSwitchButtons from '#components/ThemeSwitchButtons';
import ExtendedDisplayModeSelector from '#components/ExtendedDisplayModeSelector';
import ActionSection from '#components/ActionSection';
import Favorites from '#components/Favorites';
import Impersonation from '#components/Impersonation';
import InputDialog from '#components/InputDialog';
import LoadingOverlay from '#components/LoadingOverlay';
import Toast from '#components/Toast';
import InlineAlert from '#components/InlineAlert';
import MyCommands from '#components/MyCommands';
import RecentlyUsed from '#components/RecentlyUsed';
import SortableSection from '#components/SortableSection';
import { useSectionOrder, SectionId } from './hooks/useSectionOrder';
import {
  getActionDisabledReason,
  probeFormContextInPage,
} from './utils/actionAvailability';

const App: React.FC = () => {
  const [loadingOpen, setLoadingOpen] = useState(false);
  const [toastOpen, setToastOpen] = useState(false);
  const [toastMessage, setToastMessage] = useState('');
  const [toastSeverity, setToastSeverity] = useState<'success' | 'info' | 'warning' | 'error'>(
    'success'
  );
  const [inlineAlertOpen, setInlineAlertOpen] = useState(false);
  const [inlineAlertMessage, setInlineAlertMessage] = useState('');
  const [inlineAlertSeverity, setInlineAlertSeverity] = useState<
    'success' | 'info' | 'warning' | 'error'
  >('warning');
  const [inputDialogOpen, setInputDialogOpen] = useState(false);
  const [dialogType, setDialogType] = useState<'open-by-id' | 'new-record' | 'open-list'>(
    'open-by-id'
  );
  const [entityName, setEntityName] = useState('');
  const [recordId, setRecordId] = useState('');
  const [isConnected, setIsConnected] = useState(false);
  const [extensionConfig, setExtensionConfig] = useState<ExtensionConfig>(
    ExtensionConfigService.getConfig()
  );
  const [favoriteIds, setFavoriteIds] = useState<DynamicsAction[]>([]);
  /** null = still probing; false = not on a form; true = form record open */
  const [isOnForm, setIsOnForm] = useState<boolean | null>(null);
  const { order: sectionOrder, moveSectionRelative } = useSectionOrder();

  const disabledReasons = useMemo(() => {
    const map: Partial<Record<DynamicsAction, string>> = {};
    const opts = { connected: isConnected, isOnForm };
    for (const action of [...formActions, ...navigationActions, ...debuggingActions]) {
      const reason = getActionDisabledReason(action.id, opts);
      if (reason) map[action.id] = reason;
    }
    return map;
  }, [isConnected, isOnForm]);

  const resolveDisabledReason = (id: DynamicsAction) => disabledReasons[id] || null;

  // Function to show inline alert for critical messages
  const showInlineAlert = (
    message: string,
    severity: 'success' | 'info' | 'warning' | 'error' = 'warning'
  ) => {
    setInlineAlertMessage(message);
    setInlineAlertSeverity(severity);
    setInlineAlertOpen(true);
    // Auto-hide after 10 seconds for warnings/errors, 6 seconds for others
    setTimeout(
      () => {
        setInlineAlertOpen(false);
      },
      severity === 'warning' || severity === 'error' ? 10000 : 6000
    );
  };

  // Memoize action arrays for better performance
  const memoizedFormActions = useMemo(() => formActions, []);
  const memoizedNavigationActions = useMemo(() => navigationActions, []);
  const memoizedDebuggingActions = useMemo(() => debuggingActions, []);

  useEffect(() => {
    // Use shared helpers for detection and environment URL extraction

    const probeFormContext = async (tabId: number): Promise<boolean> => {
      try {
        // Xrm lives in the page's MAIN world — isolated-world probes always miss it.
        const results = await chrome.scripting.executeScript({
          target: { tabId },
          world: 'MAIN',
          func: probeFormContextInPage,
        });
        const result = results?.[0]?.result as { isOnForm?: boolean; pageType?: string } | undefined;
        return result?.isOnForm === true;
      } catch {
        // Fallback: inject a one-shot MAIN-world probe via script tag through content script bridge
        try {
          const results = await chrome.scripting.executeScript({
            target: { tabId },
            world: 'MAIN',
            func: () => {
              try {
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const win = window as any;
                if (win.Xrm?.Page?.data?.entity) return true;
                const pageType = win.Xrm?.Utility?.getPageContext?.()?.input?.pageType;
                return String(pageType || '').toLowerCase() === 'entityrecord';
              } catch {
                return false;
              }
            },
          });
          return results?.[0]?.result === true;
        } catch {
          return false;
        }
      }
    };

    const updateConnectionState = async (tab: chrome.tabs.Tab) => {
      const connected = await checkDynamicsViaXrm();
      setIsConnected(connected);

      if (connected && tab.id != null) {
        setIsOnForm(await probeFormContext(tab.id));
      } else {
        setIsOnForm(false);
      }
    };

    const getTargetTab = async (): Promise<chrome.tabs.Tab | undefined> => {
      const looksLikeDynamics = (url?: string) => {
        if (!url) return false;
        try {
          const host = new URL(url).hostname.toLowerCase();
          return (
            host.endsWith('.dynamics.com') ||
            host.endsWith('.powerapps.com') ||
            host.endsWith('.powerplatform.com') ||
            host.endsWith('.powerplatformusercontent.com')
          );
        } catch {
          return false;
        }
      };

      // Prefer an active Dynamics/Power Apps tab (popout iframe can steal focus).
      const [lastFocused] = await chrome.tabs.query({ active: true, lastFocusedWindow: true });
      if (lastFocused && looksLikeDynamics(lastFocused.url)) return lastFocused;
      const [current] = await chrome.tabs.query({ active: true, currentWindow: true });
      if (current && looksLikeDynamics(current.url)) return current;
      const all = await chrome.tabs.query({ active: true });
      const dyn = all.find(t => looksLikeDynamics(t.url));
      if (dyn) return dyn;
      return lastFocused || current;
    };

    const checkDynamicsConnection = async () => {
      try {
        const tab = await getTargetTab();
        if (tab) {
          await updateConnectionState(tab);
        } else {
          setIsConnected(false);
          setIsOnForm(false);
        }
      } catch (err) {
        console.error('Error checking connection:', err);
        setIsConnected(false);
        setIsOnForm(false);
      }
    };

    checkDynamicsConnection();

    const handleTabUpdate = async (
      tabId: number,
      changeInfo: chrome.tabs.TabChangeInfo,
      tab: chrome.tabs.Tab
    ) => {
      if (tab.active && (changeInfo.url || changeInfo.status === 'complete')) {
        await updateConnectionState(tab);
      }
    };

    const handleTabActivated = async (activeInfo: chrome.tabs.TabActiveInfo) => {
      try {
        const tab = await chrome.tabs.get(activeInfo.tabId);
        if (tab) {
          await updateConnectionState(tab);
        }
      } catch (error) {
        console.error('Error handling tab activation:', error);
        setIsConnected(false);
        setIsOnForm(false);
      }
    };

    chrome.tabs.onUpdated.addListener(handleTabUpdate);
    chrome.tabs.onActivated.addListener(handleTabActivated);

    // Re-probe while the side panel stays open (in-app navigation often won't retarget tabs).
    const pollId = window.setInterval(() => {
      void checkDynamicsConnection();
    }, 2000);

    return () => {
      window.clearInterval(pollId);
      chrome.tabs.onUpdated.removeListener(handleTabUpdate);
      chrome.tabs.onActivated.removeListener(handleTabActivated);
    };
  }, []);

  // Load favorites
  useEffect(() => {
    const saved = localStorage.getItem('levelup-favorites');
    if (saved) {
      try {
        setFavoriteIds(JSON.parse(saved));
      } catch {}
    }
  }, []);

  // Listen for unsolicited toast notifications from MessageService
  useEffect(() => {
    const handleToastEvent = (
      event: CustomEvent<{ message: string; severity: 'success' | 'info' | 'warning' | 'error' }>
    ) => {
      setToastMessage(event.detail.message);
      setToastSeverity(event.detail.severity);
      setToastOpen(true);
    };

    window.addEventListener('levelup-toast', handleToastEvent as EventListener);

    return () => {
      window.removeEventListener('levelup-toast', handleToastEvent as EventListener);
    };
  }, []);

  // Initialize extension configuration
  useEffect(() => {
    const initializeConfig = async () => {
      await ExtensionConfigService.initialize();
      setExtensionConfig(ExtensionConfigService.getConfig());
    };

    initializeConfig();

    // Subscribe to config changes
    const unsubscribe = ExtensionConfigService.subscribe(setExtensionConfig);

    return unsubscribe;
  }, []);

  const handleDisplayModeChange = async (mode: ExtensionDisplayMode) => {
    await ExtensionConfigService.setDisplayMode(mode);
  };

  const handleFavoriteToggle = (buttonId: DynamicsAction) => {
    const newFavorites = favoriteIds.includes(buttonId)
      ? favoriteIds.filter(id => id !== buttonId)
      : [...favoriteIds, buttonId];
    setFavoriteIds(newFavorites);
    localStorage.setItem('levelup-favorites', JSON.stringify(newFavorites));
  };

  const favoriteButtons = useMemo(() => {
    const all = [...memoizedFormActions, ...memoizedNavigationActions, ...memoizedDebuggingActions];
    return all.filter(a => favoriteIds.includes(a.id));
  }, [memoizedFormActions, memoizedNavigationActions, memoizedDebuggingActions, favoriteIds]);

  // All actions for recently used component
  const allActions = useMemo(() => {
    return [...memoizedFormActions, ...memoizedNavigationActions, ...memoizedDebuggingActions];
  }, [memoizedFormActions, memoizedNavigationActions, memoizedDebuggingActions]);

  // Keep actions in their sections even when favorited. (Previously favorites were
  // removed from sections — if Favorites was hidden by display mode, the panel went blank.)
  const filteredFormActions = memoizedFormActions;
  const filteredNavigationActions = memoizedNavigationActions;
  const filteredDebuggingActions = memoizedDebuggingActions;

  const handleActionClick = async (id: DynamicsAction) => {
    if (!isConnected) {
      setToastMessage('Please navigate to a Dynamics 365/Power Apps page to use this feature');
      setToastSeverity('warning');
      setToastOpen(true);
      return;
    }

    const disabledReason = resolveDisabledReason(id);
    if (disabledReason) {
      showInlineAlert(disabledReason, 'warning');
      return;
    }

    // Dispatch custom event for recently used tracking
    window.dispatchEvent(
      new CustomEvent('levelup-action-used', {
        detail: { actionId: id },
      })
    );

    if (id === 'navigation:open-record-by-id') {
      setDialogType('open-by-id');
      setInputDialogOpen(true);
      return;
    }
    if (id === 'navigation:new-record') {
      setDialogType('new-record');
      setInputDialogOpen(true);
      return;
    }
    if (id === 'navigation:open-list') {
      setDialogType('open-list');
      setInputDialogOpen(true);
      return;
    }

    setLoadingOpen(true);
    try {
      const response = await messageService.sendMessage(id);
      if (!response.success) {
        setToastMessage(`Action failed: ${response.error || 'Unknown error'}`);
        setToastSeverity('error');
        setToastOpen(true);
      }
    } catch (error) {
      setToastMessage(`Error: ${error instanceof Error ? error.message : 'Unknown error'}`);
      setToastSeverity('error');
      setToastOpen(true);
    } finally {
      setLoadingOpen(false);
    }
  };

  const handleDialogSubmit = async () => {
    if (dialogType === 'open-by-id') {
      if (!entityName.trim() || !recordId.trim()) {
        setToastMessage('Entity name and record ID are required');
        setToastSeverity('error');
        setToastOpen(true);
        return;
      }
    } else if (dialogType === 'new-record' || dialogType === 'open-list') {
      if (!entityName.trim()) {
        setToastMessage('Entity name is required');
        setToastSeverity('error');
        setToastOpen(true);
        return;
      }
    }

    setInputDialogOpen(false);
    setLoadingOpen(true);
    try {
      let action = '';
      let data: Record<string, string> = {};
      switch (dialogType) {
        case 'open-by-id':
          action = 'navigation:open-record-by-id';
          data = { entityName: entityName.trim(), recordId: recordId.trim() };
          break;
        case 'new-record':
          action = 'navigation:new-record';
          data = { entityName: entityName.trim() };
          break;
        case 'open-list':
          action = 'navigation:open-list';
          data = { entityName: entityName.trim() };
          break;
      }

      const response = await messageService.sendMessage(action as DynamicsAction, data);
      if (response.success) {
        let successMessage = '';
        switch (dialogType) {
          case 'open-by-id':
            successMessage = `Record opened: ${entityName} (${recordId})`;
            break;
          case 'new-record':
            successMessage = `New ${entityName} record created`;
            break;
          case 'open-list':
            successMessage = `${entityName} list opened`;
            break;
        }
        setToastMessage(successMessage);
        setToastSeverity('success');
        setToastOpen(true);
        setEntityName('');
        setRecordId('');
      } else {
        setToastMessage(`Action failed: ${response.error || 'Unknown error'}`);
        setToastSeverity('error');
        setToastOpen(true);
      }
    } catch (error) {
      setToastMessage(`Error: ${error instanceof Error ? error.message : 'Unknown error'}`);
      setToastSeverity('error');
      setToastOpen(true);
    } finally {
      setLoadingOpen(false);
    }
  };

  const handleDialogCancel = () => {
    setInputDialogOpen(false);
    setEntityName('');
    setRecordId('');
  };

  // Hosted inside the SPO popout iframe — chrome lives in the outer purple header.
  // frameElement is null across origins (Dynamics parent ↔ extension iframe), so use ?host=popout.
  const hostedInPopout =
    typeof window !== 'undefined' &&
    new URLSearchParams(window.location.search).get('host') === 'popout';

  return (
    <Box
      sx={{
        height: '100vh',
        display: 'flex',
        flexDirection: 'column',
        boxSizing: 'border-box',
        bgcolor: 'background.default',
        color: 'text.primary',
      }}
    >
      {!hostedInPopout && (
        <Box
          className='sidebar-chrome-bar'
          sx={{
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'flex-end',
            gap: 1,
            px: 1.25,
            py: 0.75,
            flexShrink: 0,
            bgcolor: 'background.paper',
            borderBottom: 1,
            borderColor: 'divider',
            zIndex: 20,
          }}
        >
          <ExtendedDisplayModeSelector
            currentMode={extensionConfig.displayMode}
            onModeChange={handleDisplayModeChange}
          />
          <ThemeSwitchButtons />
        </Box>
      )}

      <Box sx={{ flex: 1, overflowY: 'auto', padding: '6px' }}>
        {isConnected ? (
          <>
            {sectionOrder.map(sectionId => {
              let content: React.ReactNode = null;

              switch (sectionId as SectionId) {
                case 'impersonation':
                  if (extensionConfig.showImpersonation) content = <Impersonation />;
                  break;
                case 'recentlyUsed':
                  if (extensionConfig.showRecentlyUsed) {
                    content = (
                      <RecentlyUsed
                        allActions={allActions}
                        onActionClick={handleActionClick}
                        onActionUsed={() => {}}
                        getDisabledReason={resolveDisabledReason}
                      />
                    );
                  }
                  break;
                case 'favorites':
                  if (favoriteButtons.length > 0) {
                    content = (
                      <Favorites
                        favoriteButtons={favoriteButtons}
                        onActionClick={handleActionClick}
                        onFavoriteToggle={handleFavoriteToggle}
                        getDisabledReason={resolveDisabledReason}
                      />
                    );
                  }
                  break;
                case 'commands':
                  if (extensionConfig.showCustomCommands) {
                    content = (
                      <MyCommands
                        onToast={(message, severity) => {
                          setToastMessage(message);
                          setToastSeverity(severity);
                          setToastOpen(true);
                        }}
                      />
                    );
                  }
                  break;
                case 'form':
                  if (
                    extensionConfig.showActionSections &&
                    extensionConfig.showFormSection !== false &&
                    filteredFormActions.length > 0
                  ) {
                    content = (
                      <ActionSection
                        title='Form'
                        buttons={filteredFormActions}
                        onActionClick={handleActionClick}
                        onFavoriteToggle={handleFavoriteToggle}
                        favoriteIds={favoriteIds}
                        getDisabledReason={resolveDisabledReason}
                      />
                    );
                  }
                  break;
                case 'navigation':
                  if (
                    extensionConfig.showActionSections &&
                    extensionConfig.showNavigationSection !== false &&
                    filteredNavigationActions.length > 0
                  ) {
                    content = (
                      <ActionSection
                        title='Navigation'
                        buttons={filteredNavigationActions}
                        onActionClick={handleActionClick}
                        onFavoriteToggle={handleFavoriteToggle}
                        favoriteIds={favoriteIds}
                        getDisabledReason={resolveDisabledReason}
                      />
                    );
                  }
                  break;
                case 'debugging':
                  if (
                    extensionConfig.showActionSections &&
                    extensionConfig.showDebuggingSection !== false &&
                    filteredDebuggingActions.length > 0
                  ) {
                    content = (
                      <ActionSection
                        title='Debugging'
                        buttons={filteredDebuggingActions}
                        onActionClick={handleActionClick}
                        onFavoriteToggle={handleFavoriteToggle}
                        favoriteIds={favoriteIds}
                        getDisabledReason={resolveDisabledReason}
                      />
                    );
                  }
                  break;
                default:
                  break;
              }

              if (!content) return null;

              return (
                <SortableSection
                  key={sectionId}
                  id={sectionId}
                  onReorder={moveSectionRelative}
                >
                  {content}
                </SortableSection>
              );
            })}
          </>
        ) : (
          <Box
            sx={{
              padding: '24px 16px',
              textAlign: 'center',
              color: 'text.secondary',
              display: 'flex',
              flexDirection: 'column',
              alignItems: 'center',
              justifyContent: 'center',
              gap: 1.5,
              height: '100%',
            }}
          >
            <Box component='p' sx={{ margin: 0, fontWeight: 500 }}>
              These tools are available only on Microsoft Dynamics 365/Power Apps pages.
            </Box>
            <Box component='p' sx={{ fontSize: '0.875rem', margin: 0 }}>
              Open or switch to a Dynamics 365/Power Apps tab to access actions, tools, and
              commands.
            </Box>
          </Box>
        )}
      </Box>

      <InputDialog
        open={inputDialogOpen}
        onClose={handleDialogCancel}
        onSubmit={handleDialogSubmit}
        title={
          dialogType === 'open-by-id'
            ? 'Open Record by ID'
            : dialogType === 'new-record'
              ? 'Create New Record'
              : dialogType === 'open-list'
                ? 'Open Entity List'
                : ''
        }
        type={dialogType}
        entityName={entityName}
        recordId={recordId}
        onEntityNameChange={setEntityName}
        onRecordIdChange={setRecordId}
      />
      <LoadingOverlay open={loadingOpen} />
      <Toast
        open={toastOpen}
        onClose={() => setToastOpen(false)}
        message={toastMessage}
        severity={toastSeverity}
      />

      {/* Floating overlay alerts - no layout shift */}
      <InlineAlert
        open={inlineAlertOpen}
        onClose={() => setInlineAlertOpen(false)}
        message={inlineAlertMessage}
        severity={inlineAlertSeverity}
      />
    </Box>
  );
};

const container = document.getElementById('root');
if (container) {
  const root = createRoot(container);
  root.render(
    <ThemeProvider>
      <App />
    </ThemeProvider>
  );
}

export default App;

