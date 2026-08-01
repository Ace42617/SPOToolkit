import { DynamicsAction } from '#types/global';

/**
 * Result of probing the active Dynamics tab for form vs non-form context.
 * `isOnForm === null` means still checking.
 */
export type FormContextState = {
  isOnForm: boolean | null;
  pageType: string | null;
};

/**
 * Returns a short reason when an action cannot be used on the current page.
 * Navigation / debugging stay available when connected (they navigate or set URL flags).
 */
export function getActionDisabledReason(
  actionId: DynamicsAction,
  opts: { connected: boolean; isOnForm: boolean | null }
): string | null {
  if (!opts.connected) {
    return 'Open a Dynamics 365 / Power Apps page to use this action';
  }

  if (actionId.startsWith('form:') && opts.isOnForm !== true) {
    return opts.isOnForm === null
      ? 'Checking whether a form is open…'
      : 'Requires an open form (record) page';
  }

  return null;
}

/**
 * Injected into the Dynamics tab.
 * Mirrors DynamicsUtils.isFormContext() — no URL heuristics
 * (etn/id appear on many non-form pages and caused false positives).
 */
export function probeFormContextInPage(): { isOnForm: boolean; pageType: string } {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const win = window as any;

  try {
    if (win.Xrm?.Page?.data?.entity) {
      return { isOnForm: true, pageType: 'entityrecord' };
    }
  } catch {
    /* continue */
  }

  try {
    const input = win.Xrm?.Utility?.getPageContext?.()?.input;
    if (input?.pageType) {
      const pageType = String(input.pageType);
      return {
        isOnForm: pageType.toLowerCase() === 'entityrecord',
        pageType,
      };
    }
  } catch {
    /* continue */
  }

  // Classic (non-UCI): check first visible iframe for form entity
  try {
    const isUci = win.Xrm?.Internal?.isUci?.();
    if (!isUci) {
      const visibleIframes = Array.from(document.querySelectorAll('iframe')).filter(
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (iframe: any) => iframe.style?.visibility !== 'hidden'
      ) as HTMLIFrameElement[];
      if (visibleIframes.length > 0) {
        const frameWindow = visibleIframes[0].contentWindow as any;
        if (frameWindow?.Xrm?.Page?.data?.entity) {
          return { isOnForm: true, pageType: 'entityrecord' };
        }
      }
    }
  } catch {
    /* cross-origin or unavailable */
  }

  return { isOnForm: false, pageType: 'unknown' };
}
