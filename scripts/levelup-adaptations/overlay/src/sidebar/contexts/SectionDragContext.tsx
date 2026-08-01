import React from 'react';

export interface SectionDragHandleProps {
  draggable: boolean;
  onDragStart: (event: React.DragEvent) => void;
  onDragEnd: (event: React.DragEvent) => void;
}

export const SectionDragContext = React.createContext<SectionDragHandleProps | null>(null);

export function useSectionDragHandle(): SectionDragHandleProps | null {
  return React.useContext(SectionDragContext);
}
