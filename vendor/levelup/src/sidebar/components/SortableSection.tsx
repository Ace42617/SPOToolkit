import React, { useCallback, useState } from 'react';
import { Box } from '@mui/material';
import { SectionDragContext } from '../contexts/SectionDragContext';
import type { SectionId } from '../hooks/useSectionOrder';

interface SortableSectionProps {
  id: SectionId;
  onReorder: (fromId: SectionId, toId: SectionId, place: 'before' | 'after') => void;
  children: React.ReactNode;
}

/** Shared during drag — dataTransfer.getData is unavailable in dragOver in most browsers. */
let activeDragId: SectionId | null = null;

const SortableSection: React.FC<SortableSectionProps> = ({ id, onReorder, children }) => {
  const [dragging, setDragging] = useState(false);

  const onDragStart = useCallback(
    (event: React.DragEvent) => {
      activeDragId = id;
      event.dataTransfer.effectAllowed = 'move';
      event.dataTransfer.setData('text/plain', id);
      setDragging(true);
    },
    [id]
  );

  const onDragEnd = useCallback(() => {
    activeDragId = null;
    setDragging(false);
  }, []);

  const handleDragOver = useCallback(
    (event: React.DragEvent) => {
      if (!activeDragId || activeDragId === id) return;
      event.preventDefault();
      event.dataTransfer.dropEffect = 'move';
      const rect = (event.currentTarget as HTMLElement).getBoundingClientRect();
      const place: 'before' | 'after' =
        event.clientY < rect.top + rect.height / 2 ? 'before' : 'after';
      onReorder(activeDragId, id, place);
    },
    [id, onReorder]
  );

  const handleDrop = useCallback((event: React.DragEvent) => {
    event.preventDefault();
    activeDragId = null;
  }, []);

  return (
    <SectionDragContext.Provider
      value={{
        draggable: true,
        onDragStart,
        onDragEnd,
      }}
    >
      <Box
        onDragOver={handleDragOver}
        onDrop={handleDrop}
        sx={{
          mb: 0.75,
          opacity: dragging ? 0.4 : 1,
          transition: 'opacity 0.12s ease',
          borderRadius: 2,
          // Section components may return null (e.g. empty Recently Used)
          '&:empty': { display: 'none', mb: 0 },
        }}
      >
        {children}
      </Box>
    </SectionDragContext.Provider>
  );
};

export default SortableSection;
