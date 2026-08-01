import React, { useState } from 'react';
import {
  Card,
  CardContent,
  Box,
  Chip,
  CircularProgress,
  Alert,
  Collapse,
} from '@mui/material';
import { useImpersonation, UserToImpersonate } from '#hooks/useImpersonation';
import PrivilegeWarning from './PrivilegeWarning';
import ImpersonationStatusBanner from './ImpersonationStatusBanner';
import UserSearchInput from './UserSearchInput';
import SectionHeader from '#components/SectionHeader';

const Impersonation = () => {
  const [expanded, setExpanded] = useState(true);
  const {
    hasImpersonationPrivilege,
    isCheckingPrivilege,
    isImpersonating,
    impersonatedUser,
    isCheckingStatus,
    searchResults,
    searchMessage,
    isSearching,
    hasMoreResults,
    favoriteUsers,
    addToFavorites,
    removeFromFavorites,
    isFavorite,
    error,
    startImpersonation,
    stopImpersonation,
    searchUsers,
    clearError,
    retryPrivilegeCheck,
  } = useImpersonation();

  const handleUserSelect = (_user: UserToImpersonate | null) => {
    // Handled within UserSearchInput
  };

  const statusBadge = (
    <>
      {(hasImpersonationPrivilege === null || isCheckingPrivilege) && (
        <CircularProgress size={16} sx={{ mx: 0.5 }} />
      )}
      {hasImpersonationPrivilege === false && (
        <Chip label='No Access' color='error' size='small' sx={{ height: 20, fontSize: '0.7rem' }} />
      )}
      {hasImpersonationPrivilege === true && (
        <Box sx={{ display: 'flex', alignItems: 'center', gap: 0.5 }}>
          {isCheckingStatus && <CircularProgress size={14} />}
          <Chip
            label={isImpersonating ? 'Active' : 'Inactive'}
            color={isImpersonating ? 'success' : 'default'}
            size='small'
            sx={{ height: 20, fontSize: '0.7rem' }}
          />
        </Box>
      )}
    </>
  );

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
          title='Impersonation'
          expanded={expanded}
          onToggle={() => setExpanded(!expanded)}
          badges={statusBadge}
        />

        <Collapse in={expanded} timeout='auto' unmountOnExit>
          <Box sx={{ p: 2, pt: 1.5 }}>
            <PrivilegeWarning
              show={hasImpersonationPrivilege === false}
              onRetry={retryPrivilegeCheck}
              isRetrying={isCheckingPrivilege}
            />

            {(hasImpersonationPrivilege === true || isImpersonating) && (
              <>
                <ImpersonationStatusBanner
                  isImpersonating={isImpersonating}
                  impersonatedUser={impersonatedUser}
                  onStopImpersonation={stopImpersonation}
                />

                {error && (
                  <Alert severity='error' sx={{ mb: 2 }} onClose={clearError}>
                    {error}
                  </Alert>
                )}

                {!isImpersonating && (
                  <UserSearchInput
                    searchResults={searchResults}
                    isSearching={isSearching}
                    searchMessage={searchMessage}
                    hasMoreResults={hasMoreResults}
                    favoriteUsers={favoriteUsers}
                    onSearchChange={searchUsers}
                    onUserSelect={handleUserSelect}
                    onStartImpersonation={startImpersonation}
                    onAddToFavorites={addToFavorites}
                    onRemoveFromFavorites={removeFromFavorites}
                    isFavorite={isFavorite}
                  />
                )}
              </>
            )}
          </Box>
        </Collapse>
      </CardContent>
    </Card>
  );
};

export default Impersonation;
