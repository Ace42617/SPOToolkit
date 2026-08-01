import React, { useState } from 'react';
import {
  IconButton,
  Menu,
  MenuItem,
  ListItemText,
  ListItemIcon,
  Tooltip,
  Typography,
  Divider,
} from '@mui/material';
import {
  Settings as SettingsIcon,
  ViewModule as DefaultIcon,
  ViewList as SimpleIcon,
  Check as CheckIcon,
} from '@mui/icons-material';
import { ExtensionDisplayMode } from '#types/global';

interface ExtendedDisplayModeSelectorProps {
  currentMode: ExtensionDisplayMode;
  onModeChange: (mode: ExtensionDisplayMode) => void;
}

const DISPLAY_MODE_CONFIG = {
  default: {
    label: 'Default',
    description: 'All features visible',
    icon: DefaultIcon,
  },
  simple: {
    label: 'Simple',
    description: 'Essential features only',
    icon: SimpleIcon,
  },
} as const;

const ExtendedDisplayModeSelector: React.FC<ExtendedDisplayModeSelectorProps> = ({
  currentMode,
  onModeChange,
}) => {
  const [anchorEl, setAnchorEl] = useState<null | HTMLElement>(null);
  const open = Boolean(anchorEl);

  const handleClick = (event: React.MouseEvent<HTMLButtonElement>) => {
    setAnchorEl(event.currentTarget);
  };

  const handleClose = () => {
    setAnchorEl(null);
  };

  const handleModeSelect = (mode: ExtensionDisplayMode) => {
    onModeChange(mode);
    handleClose();
  };

  return (
    <>
      <Tooltip title='Display mode'>
        <IconButton
          onClick={handleClick}
          size='small'
          aria-label='Display mode'
          className='spo-chrome-icon-btn'
          sx={{
            width: 28,
            height: 28,
            color: 'text.secondary',
            borderRadius: '10px',
            '&:hover': {
              backgroundColor: 'action.hover',
              color: 'primary.main',
            },
          }}
        >
          <SettingsIcon sx={{ fontSize: '1.05rem' }} />
        </IconButton>
      </Tooltip>

      <Menu
        anchorEl={anchorEl}
        open={open}
        onClose={handleClose}
        slotProps={{
          paper: {
            sx: {
              minWidth: 220,
              maxWidth: 260,
              borderRadius: '9px',
              '& .MuiMenuItem-root': {
                py: 1,
                borderRadius: '7px',
                mx: 0.5,
              },
            },
          },
        }}
        anchorOrigin={{
          vertical: 'bottom',
          horizontal: 'right',
        }}
        transformOrigin={{
          vertical: 'top',
          horizontal: 'right',
        }}
      >
        <Typography
          variant='overline'
          sx={{
            px: 2,
            py: 1,
            fontWeight: 600,
            color: 'text.secondary',
            fontSize: '0.7rem',
          }}
        >
          Display Mode
        </Typography>
        <Divider />

        {(
          Object.entries(DISPLAY_MODE_CONFIG) as [
            ExtensionDisplayMode,
            (typeof DISPLAY_MODE_CONFIG)[ExtensionDisplayMode],
          ][]
        ).map(([mode, config]) => {
          const IconComponent = config.icon;
          const isSelected = currentMode === mode;

          return (
            <MenuItem
              key={mode}
              onClick={() => handleModeSelect(mode)}
              selected={isSelected}
              sx={{
                '&.Mui-selected': {
                  backgroundColor: 'primary.main',
                  color: 'primary.contrastText',
                  '&:hover': {
                    backgroundColor: 'primary.dark',
                  },
                  // Ensure text is always readable
                  '& .MuiListItemText-primary': {
                    color: 'primary.contrastText',
                    fontWeight: 600,
                  },
                  '& .MuiListItemText-secondary': {
                    color: 'primary.contrastText',
                    opacity: 0.9,
                  },
                },
              }}
            >
              <ListItemIcon
                sx={{
                  color: isSelected ? 'primary.contrastText' : 'text.secondary',
                  minWidth: 36,
                }}
              >
                <IconComponent fontSize='small' />
              </ListItemIcon>
              <ListItemText
                primary={config.label}
                secondary={config.description}
                sx={{
                  '& .MuiListItemText-primary': {
                    color: isSelected ? 'inherit' : 'text.primary',
                    fontWeight: isSelected ? 600 : 400,
                  },
                  '& .MuiListItemText-secondary': {
                    color: isSelected ? 'inherit' : 'text.secondary',
                    opacity: isSelected ? 0.9 : 0.7,
                    fontSize: '0.75rem',
                  },
                }}
              />
              {isSelected && (
                <CheckIcon
                  fontSize='small'
                  sx={{
                    color: 'primary.contrastText',
                    ml: 1,
                  }}
                />
              )}
            </MenuItem>
          );
        })}
      </Menu>
    </>
  );
};

export default ExtendedDisplayModeSelector;
