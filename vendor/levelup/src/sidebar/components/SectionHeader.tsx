import React from 'react';
import { Box, Typography, IconButton, Chip, Tooltip } from '@mui/material';
import {
  ExpandMore as ExpandMoreIcon,
  DragIndicator as DragIndicatorIcon,
} from '@mui/icons-material';
import { useSectionDragHandle } from '../contexts/SectionDragContext';

/** Shared collapsed header height so every section aligns. */
export const SECTION_HEADER_HEIGHT = 40;

export const sectionActionIconSx = {
  width: 28,
  height: 28,
  p: 0.5,
} as const;

interface SectionHeaderProps {
  title: string;
  expanded: boolean;
  onToggle: () => void;
  count?: number | string;
  icon?: React.ReactNode;
  /** Extra chips / labels after the count chip */
  badges?: React.ReactNode;
  /** Compact action buttons before the chevron */
  actions?: React.ReactNode;
}

const SectionHeader: React.FC<SectionHeaderProps> = ({
  title,
  expanded,
  onToggle,
  count,
  icon,
  badges,
  actions,
}) => {
  const dragHandle = useSectionDragHandle();

  return (
    <Box
      onClick={onToggle}
      sx={{
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'space-between',
        gap: 0.5,
        height: SECTION_HEADER_HEIGHT,
        minHeight: SECTION_HEADER_HEIGHT,
        px: 1.25,
        cursor: 'pointer',
        backgroundColor: 'action.hover',
        borderRadius: '8px 8px 0 0',
        transition: 'background-color 0.2s ease-in-out',
        '&:hover': {
          backgroundColor: 'action.selected',
        },
      }}
    >
      <Box
        sx={{
          display: 'flex',
          alignItems: 'center',
          gap: 0.75,
          minWidth: 0,
          flex: 1,
        }}
      >
        {dragHandle && (
          <Tooltip title='Drag to reorder' placement='top'>
            <Box
              component='span'
              draggable={dragHandle.draggable}
              onDragStart={dragHandle.onDragStart}
              onDragEnd={dragHandle.onDragEnd}
              onClick={e => e.stopPropagation()}
              sx={{
                display: 'inline-flex',
                alignItems: 'center',
                justifyContent: 'center',
                width: 22,
                height: 28,
                color: 'text.disabled',
                cursor: 'grab',
                flexShrink: 0,
                borderRadius: 0.75,
                '&:hover': { color: 'text.secondary', backgroundColor: 'action.hover' },
                '&:active': { cursor: 'grabbing' },
              }}
            >
              <DragIndicatorIcon sx={{ fontSize: '1.05rem' }} />
            </Box>
          </Tooltip>
        )}
        {icon}
        <Typography
          variant='h6'
          component='h2'
          sx={{
            fontSize: '0.95rem',
            fontWeight: 600,
            color: 'text.primary',
            userSelect: 'none',
            whiteSpace: 'nowrap',
            overflow: 'hidden',
            textOverflow: 'ellipsis',
          }}
        >
          {title}
        </Typography>
        {count !== undefined && (
          <Chip
            label={count}
            size='small'
            sx={{
              height: 20,
              fontSize: '0.7rem',
              fontWeight: 600,
              backgroundColor: 'primary.main',
              color: 'primary.contrastText',
              flexShrink: 0,
              '& .MuiChip-label': { px: 1 },
            }}
          />
        )}
        {badges}
      </Box>

      <Box sx={{ display: 'flex', alignItems: 'center', gap: 0.25, flexShrink: 0 }}>
        {actions}
        <IconButton
          size='small'
          tabIndex={-1}
          sx={{
            ...sectionActionIconSx,
            transform: expanded ? 'rotate(180deg)' : 'rotate(0deg)',
            transition: 'transform 0.2s ease-in-out',
            color: 'text.secondary',
          }}
        >
          <ExpandMoreIcon fontSize='small' />
        </IconButton>
      </Box>
    </Box>
  );
};

export default SectionHeader;
