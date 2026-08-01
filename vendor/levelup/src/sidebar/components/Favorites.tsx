import React from 'react';
import { Card, CardContent, Box, Collapse } from '@mui/material';
import { Favorite as HeartIcon, Close as CloseIcon } from '@mui/icons-material';
import { DynamicsAction } from '#types/global';
import StandardActionButton, { StandardActionGrid } from '#components/StandardActionButton';
import SectionHeader from '#components/SectionHeader';

interface ActionButton {
  id: DynamicsAction;
  label: string;
  icon: React.ComponentType<any>;
  tooltip?: string;
}

interface FavoritesProps {
  favoriteButtons: ActionButton[];
  onActionClick: (id: DynamicsAction) => void;
  onFavoriteToggle: (id: DynamicsAction) => void;
  getDisabledReason?: (id: DynamicsAction) => string | null;
}

const FavoritesComponent: React.FC<FavoritesProps> = ({
  favoriteButtons,
  onActionClick,
  onFavoriteToggle,
  getDisabledReason,
}) => {
  const [expanded, setExpanded] = React.useState(true);

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
          title='Favorites'
          expanded={expanded}
          onToggle={() => setExpanded(!expanded)}
          count={favoriteButtons.length}
          icon={<HeartIcon sx={{ fontSize: '1.05rem', color: 'error.main' }} />}
        />

        <Collapse in={expanded} timeout='auto' unmountOnExit>
          <Box sx={{ p: 2, pt: 1.5 }}>
            <StandardActionGrid minColumnWidth={80}>
              {favoriteButtons.map(button => {
                const disabledReason = getDisabledReason?.(button.id) || null;
                return (
                  <StandardActionButton
                    key={button.id}
                    id={button.id}
                    label={button.label}
                    icon={button.icon}
                    tooltip={button.tooltip}
                    onClick={() => onActionClick(button.id)}
                    isFavorite={true}
                    onFavoriteToggle={() => onFavoriteToggle(button.id)}
                    showFavorite={true}
                    showLabel={true}
                    favoriteVariant='subtle'
                    favoriteIcon={CloseIcon}
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

export default React.memo(FavoritesComponent);
