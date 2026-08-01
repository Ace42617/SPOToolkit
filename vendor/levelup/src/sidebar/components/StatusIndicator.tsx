import React from 'react';
import { Box, Typography, CircularProgress } from '@mui/material';
import { Circle as CircleIcon } from '@mui/icons-material';

export interface ConnectionInfo {
  appName?: string;
  orgLabel?: string;
  host?: string;
  clientUrl?: string;
}

interface StatusIndicatorProps {
  loading: boolean;
  connected: boolean;
  /** @deprecated Prefer connectionInfo.host / clientUrl */
  environmentUrl?: string;
  connectionInfo?: ConnectionInfo;
  onOpenAdmin?: (environmentUrl?: string) => void;
}

const StatusIndicatorComponent: React.FC<StatusIndicatorProps> = ({
  loading,
  connected,
  environmentUrl,
  connectionInfo,
  onOpenAdmin,
}) => {
  if (loading) {
    return (
      <Box sx={{ display: 'flex', alignItems: 'center', gap: 1 }}>
        <CircularProgress size={12} />
        <Typography variant='body2' sx={{ fontSize: '0.75rem' }}>
          Checking...
        </Typography>
      </Box>
    );
  }

  const appName = (connectionInfo?.appName || '').trim();
  const orgLabel = (connectionInfo?.orgLabel || '').trim();
  const host = (connectionInfo?.host || environmentUrl || '').trim();
  const clientUrl = (connectionInfo?.clientUrl || '').trim();

  const primary = connected
    ? appName || host || 'Connected'
    : 'Not Connected';

  const secondaryParts: string[] = [];
  if (connected) {
    if (orgLabel) secondaryParts.push(orgLabel);
    if (host && host !== primary) secondaryParts.push(host);
    else if (clientUrl && !host) secondaryParts.push(clientUrl);
  }
  const secondary = secondaryParts.join(' · ');

  const handleClick = () => {
    if (onOpenAdmin) onOpenAdmin(clientUrl || host || environmentUrl);
  };

  const tip = connected
    ? [appName && `App: ${appName}`, orgLabel && `Org: ${orgLabel}`, clientUrl || host]
        .filter(Boolean)
        .join('\n') + (clientUrl || host ? '\nClick to open Power Platform Admin Center' : '')
    : undefined;

  return (
    <Box sx={{ display: 'flex', alignItems: 'center', gap: 1.25, minWidth: 0, flex: 1 }}>
      <CircleIcon
        sx={{
          fontSize: 14,
          color: connected ? 'success.main' : 'error.main',
          flexShrink: 0,
        }}
      />
      <Box
        component='button'
        onClick={handleClick}
        title={tip}
        sx={{
          appearance: 'none',
          border: 'none',
          background: 'transparent',
          padding: 0,
          margin: 0,
          cursor: connected && (clientUrl || host || environmentUrl) ? 'pointer' : 'default',
          minWidth: 0,
          flex: 1,
          textAlign: 'left',
          display: 'flex',
          flexDirection: 'column',
          gap: 0.15,
        }}
      >
        <Typography
          variant='body2'
          sx={{
            color: 'text.primary',
            overflow: 'hidden',
            textOverflow: 'ellipsis',
            whiteSpace: 'nowrap',
            fontSize: '0.8rem',
            fontWeight: 600,
            lineHeight: 1.2,
            minWidth: 0,
          }}
        >
          {primary}
        </Typography>
        {secondary ? (
          <Typography
            variant='caption'
            sx={{
              color: 'text.secondary',
              overflow: 'hidden',
              textOverflow: 'ellipsis',
              whiteSpace: 'nowrap',
              fontSize: '0.68rem',
              lineHeight: 1.2,
              minWidth: 0,
            }}
          >
            {secondary}
          </Typography>
        ) : null}
      </Box>
    </Box>
  );
};

const StatusIndicator = React.memo(StatusIndicatorComponent);

export default StatusIndicator;
