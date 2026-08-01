import React, { useState, useEffect, useMemo } from 'react';
import { Card, CardContent, Box, IconButton, Collapse, Tooltip } from '@mui/material';
import { History as HistoryIcon, Clear as ClearIcon } from '@mui/icons-material';
import { DynamicsAction } from '#types/global';
import StandardActionButton, { StandardActionGrid } from '#components/StandardActionButton';
import SectionHeader, { sectionActionIconSx } from '#components/SectionHeader';

interface ActionButton {
  id: DynamicsAction;
  label: string;
  icon: React.ComponentType<any>;
  tooltip?: string;
}

interface RecentlyUsedProps {
  allActions: ActionButton[];
  onActionClick: (id: DynamicsAction) => void;
  onActionUsed: (id: DynamicsAction) => void;
  getDisabledReason?: (id: DynamicsAction) => string | null;
}

interface RecentAction {
  id: DynamicsAction;
  timestamp: number;
  count: number;
}

const RECENT_ACTIONS_KEY = 'levelup-recent-actions';
const MAX_RECENT_ACTIONS = 5;

const RecentlyUsedComponent: React.FC<RecentlyUsedProps> = ({
  allActions,
  onActionClick,
  onActionUsed,
  getDisabledReason,
}) => {
  const [expanded, setExpanded] = useState(true);
  const [recentActions, setRecentActions] = useState<RecentAction[]>([]);

  useEffect(() => {
    const saved = localStorage.getItem(RECENT_ACTIONS_KEY);
    if (saved) {
      try {
        const parsed = JSON.parse(saved) as RecentAction[];
        const weekAgo = Date.now() - 7 * 24 * 60 * 60 * 1000;
        setRecentActions(parsed.filter(action => action.timestamp > weekAgo));
      } catch (error) {
        console.warn('Failed to parse recent actions from localStorage:', error);
      }
    }
  }, []);

  useEffect(() => {
    if (recentActions.length > 0) {
      localStorage.setItem(RECENT_ACTIONS_KEY, JSON.stringify(recentActions));
    }
  }, [recentActions]);

  const addToRecent = (actionId: DynamicsAction) => {
    setRecentActions(prev => {
      const now = Date.now();
      const existing = prev.find(action => action.id === actionId);

      if (existing) {
        return prev
          .map(action =>
            action.id === actionId
              ? { ...action, timestamp: now, count: action.count + 1 }
              : action
          )
          .sort((a, b) => b.timestamp - a.timestamp);
      }

      return [{ id: actionId, timestamp: now, count: 1 }, ...prev].slice(0, MAX_RECENT_ACTIONS);
    });
  };

  useEffect(() => {
    const handleActionUsed = (event: CustomEvent<{ actionId: DynamicsAction }>) => {
      addToRecent(event.detail.actionId);
    };

    window.addEventListener('levelup-action-used', handleActionUsed as EventListener);
    return () => {
      window.removeEventListener('levelup-action-used', handleActionUsed as EventListener);
    };
  }, []);

  const recentButtons = useMemo(() => {
    return recentActions
      .map(recentAction => {
        const actionConfig = allActions.find(action => action.id === recentAction.id);
        if (!actionConfig) return null;
        return { ...actionConfig, recentData: recentAction };
      })
      .filter(Boolean) as (ActionButton & { recentData: RecentAction })[];
  }, [recentActions, allActions]);

  const handleActionClick = (id: DynamicsAction) => {
    addToRecent(id);
    onActionClick(id);
    onActionUsed(id);
  };

  const clearRecentActions = () => {
    setRecentActions([]);
    localStorage.removeItem(RECENT_ACTIONS_KEY);
  };

  if (recentButtons.length === 0) {
    return null;
  }

  return (
    <Card
      sx={{
        mb: 0,
        borderRadius: 2,
        boxShadow: 'none',
        border: 1,
        borderColor: 'divider',
      }}
    >
      <CardContent sx={{ py: 0, px: 0, '&:last-child': { pb: 0 } }}>
        <SectionHeader
          title='Recently Used'
          expanded={expanded}
          onToggle={() => setExpanded(!expanded)}
          count={recentButtons.length}
          icon={<HistoryIcon sx={{ fontSize: '1.05rem', color: 'primary.main' }} />}
          actions={
            <Tooltip title='Clear recent actions' placement='top'>
              <IconButton
                size='small'
                onClick={e => {
                  e.stopPropagation();
                  clearRecentActions();
                }}
                sx={{
                  ...sectionActionIconSx,
                  color: 'text.secondary',
                  '&:hover': { color: 'error.main', backgroundColor: 'error.light' },
                }}
              >
                <ClearIcon fontSize='small' />
              </IconButton>
            </Tooltip>
          }
        />

        <Collapse in={expanded} timeout='auto' unmountOnExit>
          <Box sx={{ p: 2, pt: 1.5 }}>
            <StandardActionGrid>
              {recentButtons.map(button => {
                const timeSince = Date.now() - button.recentData.timestamp;
                const minutesAgo = Math.floor(timeSince / (1000 * 60));
                const hoursAgo = Math.floor(timeSince / (1000 * 60 * 60));
                const daysAgo = Math.floor(hoursAgo / 24);

                let timeLabel;
                if (daysAgo > 0) timeLabel = `${daysAgo}d ago`;
                else if (hoursAgo > 0) timeLabel = `${hoursAgo}h ago`;
                else if (minutesAgo > 0) timeLabel = `${minutesAgo}m ago`;
                else timeLabel = 'Just now';

                const disabledReason = getDisabledReason?.(button.id) || null;
                return (
                  <StandardActionButton
                    key={button.id}
                    id={button.id}
                    label={button.label}
                    icon={button.icon}
                    tooltip={button.tooltip || button.label}
                    onClick={() => handleActionClick(button.id)}
                    additionalInfo={timeLabel}
                    showLabel={true}
                    disabled={!!disabledReason}
                    disabledReason={disabledReason || undefined}
                  />
                );
              })}
            </StandardActionGrid>
          </Box>
        </Collapse>
      </CardContent>
    </Card>
  );
};

const RecentlyUsed = React.memo(RecentlyUsedComponent);

export default RecentlyUsed;
