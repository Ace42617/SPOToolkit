import { useCallback, useEffect, useState } from 'react';

export const SECTION_IDS = [
  'impersonation',
  'recentlyUsed',
  'favorites',
  'commands',
  'form',
  'navigation',
  'debugging',
] as const;

export type SectionId = (typeof SECTION_IDS)[number];

const STORAGE_KEY = 'levelup-section-order';
export const DEFAULT_SECTION_ORDER: SectionId[] = [...SECTION_IDS];

function normalizeOrder(raw: unknown): SectionId[] {
  const known = new Set<string>(SECTION_IDS);
  const seen = new Set<SectionId>();
  const next: SectionId[] = [];

  if (Array.isArray(raw)) {
    for (const id of raw) {
      if (typeof id === 'string' && known.has(id) && !seen.has(id as SectionId)) {
        const sid = id as SectionId;
        seen.add(sid);
        next.push(sid);
      }
    }
  }

  for (const id of DEFAULT_SECTION_ORDER) {
    if (!seen.has(id)) next.push(id);
  }

  return next;
}

function loadOrder(): SectionId[] {
  try {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (!saved) return [...DEFAULT_SECTION_ORDER];
    return normalizeOrder(JSON.parse(saved));
  } catch {
    return [...DEFAULT_SECTION_ORDER];
  }
}

function sameOrder(a: SectionId[], b: SectionId[]): boolean {
  return a.length === b.length && a.every((id, i) => id === b[i]);
}

export function useSectionOrder() {
  const [order, setOrder] = useState<SectionId[]>(() => loadOrder());

  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(order));
    } catch {
      // ignore quota / private mode
    }
  }, [order]);

  const moveSectionRelative = useCallback(
    (fromId: SectionId, toId: SectionId, place: 'before' | 'after') => {
      if (fromId === toId) return;
      setOrder(prev => {
        const without = prev.filter(id => id !== fromId);
        const toIndex = without.indexOf(toId);
        if (toIndex < 0) return prev;
        const insertAt = place === 'before' ? toIndex : toIndex + 1;
        const next = [...without];
        next.splice(insertAt, 0, fromId);
        return sameOrder(prev, next) ? prev : next;
      });
    },
    []
  );

  return { order, moveSectionRelative };
}
