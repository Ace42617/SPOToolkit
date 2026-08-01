import React, { createContext, useContext, useState, useEffect, ReactNode } from 'react';
import { createTheme, Theme, ThemeProvider as MuiThemeProvider } from '@mui/material/styles';
import { CssBaseline } from '@mui/material';

type ThemeMode = 'light' | 'dark';

interface ThemeContextType {
  mode: ThemeMode;
  toggleTheme: () => void;
  isDarkMode: boolean;
}

const ThemeContext = createContext<ThemeContextType | undefined>(undefined);

const STORAGE_KEY = 'levelup-theme-mode';
/** Shared with SPO Toolkit popup (`popup.js` darkMode). */
const SPO_DARK_MODE_KEY = 'darkMode';

const FONT_FAMILY = '"DM Sans", system-ui, "Segoe UI", sans-serif';

// Light theme — SPO shell + Power Apps purple
const lightTheme = createTheme({
  typography: {
    fontFamily: FONT_FAMILY,
    h6: {
      fontWeight: 600,
      fontSize: '0.95rem',
    },
    body2: { fontSize: '0.8125rem' },
    caption: { fontSize: '0.7rem' },
  },
  shape: { borderRadius: 9 },
  palette: {
    mode: 'light',
    primary: {
      main: '#742774',
      light: '#8e4a8e',
      dark: '#5a1d5a',
      contrastText: '#ffffff',
    },
    secondary: {
      main: '#5a5f8a',
      light: '#9ea3c8',
      dark: '#1c1f4a',
      contrastText: '#ffffff',
    },
    success: {
      main: '#22c55e',
      light: '#86efac',
      dark: '#16a34a',
    },
    warning: {
      main: '#f59e0b',
      light: '#fcd34d',
      dark: '#d97706',
    },
    error: {
      main: '#f04e65',
      light: '#fda4af',
      dark: '#e11d48',
    },
    background: {
      default: '#f3f4fb',
      paper: '#ffffff',
    },
    text: {
      primary: '#1c1f4a',
      secondary: '#5a5f8a',
    },
    divider: '#e8dff0',
    action: {
      hover: 'rgba(116, 39, 116, 0.06)',
      selected: 'rgba(116, 39, 116, 0.1)',
    },
  },
  components: {
    MuiCssBaseline: {
      styleOverrides: {
        body: {
          fontFamily: FONT_FAMILY,
          backgroundColor: '#f3f4fb',
          color: '#1c1f4a',
        },
      },
    },
    MuiCard: {
      styleOverrides: {
        root: {
          backgroundColor: '#ffffff',
          border: '1px solid #e8dff0',
          borderRadius: '9px',
          boxShadow: 'none',
          transition: 'border-color 0.15s ease',
          '&:hover': {
            boxShadow: 'none',
            borderColor: '#d4c2de',
          },
        },
      },
    },
    MuiIconButton: {
      styleOverrides: {
        root: {
          color: '#5a5f8a',
          borderRadius: '10px',
          transition: 'background-color 0.15s ease, color 0.15s ease',
          '&:hover': {
            backgroundColor: 'rgba(116, 39, 116, 0.08)',
            color: '#742774',
          },
          '&.Mui-disabled': {
            color: '#9ea3c8',
          },
        },
      },
    },
    MuiButton: {
      styleOverrides: {
        root: {
          borderRadius: '9px',
          textTransform: 'none',
          fontWeight: 600,
          boxShadow: 'none',
          '&:hover': {
            boxShadow: 'none',
          },
        },
        contained: {
          backgroundColor: '#742774',
          color: '#ffffff',
          '&:hover': {
            backgroundColor: '#5a1d5a',
          },
        },
        outlined: {
          borderWidth: '1.5px',
          borderColor: '#e8dff0',
        },
      },
    },
    MuiChip: {
      styleOverrides: {
        root: {
          borderRadius: '999px',
          fontWeight: 600,
          fontSize: '0.7rem',
        },
        filled: {
          backgroundColor: '#f3e8f8',
          color: '#5a1d5a',
          '&:hover': {
            backgroundColor: '#e8dff0',
          },
        },
      },
    },
    MuiOutlinedInput: {
      styleOverrides: {
        root: {
          borderRadius: '9px',
          fontSize: '0.8125rem',
        },
      },
    },
  },
});

// Dark theme — SPO Night Owl + Power Apps purple
const darkTheme = createTheme({
  typography: {
    fontFamily: FONT_FAMILY,
    h6: {
      fontWeight: 600,
      fontSize: '0.95rem',
    },
    body2: { fontSize: '0.8125rem' },
    caption: { fontSize: '0.7rem' },
  },
  shape: { borderRadius: 9 },
  palette: {
    mode: 'dark',
    primary: {
      main: '#b794f6',
      light: '#dbb2ff',
      dark: '#8a4fbc',
      contrastText: '#0f1117',
    },
    secondary: {
      main: '#9ea3c8',
      light: '#c8cce0',
      dark: '#5a5f8a',
      contrastText: '#0f1117',
    },
    success: {
      main: '#4ade80',
      light: '#86efac',
      dark: '#22c55e',
    },
    warning: {
      main: '#fbbf24',
      light: '#fcd34d',
      dark: '#d97706',
    },
    error: {
      main: '#fb7185',
      light: '#fda4af',
      dark: '#f04e65',
    },
    background: {
      default: '#0f1117',
      paper: '#181b23',
    },
    text: {
      primary: '#e8eaef',
      secondary: '#9ea3c8',
    },
    divider: '#4a3a58',
    action: {
      hover: 'rgba(183, 148, 246, 0.1)',
      selected: 'rgba(183, 148, 246, 0.16)',
      disabled: 'rgba(255, 255, 255, 0.3)',
    },
  },
  components: {
    MuiCssBaseline: {
      styleOverrides: {
        body: {
          fontFamily: FONT_FAMILY,
          backgroundColor: '#0f1117',
          color: '#e8eaef',
        },
      },
    },
    MuiCard: {
      styleOverrides: {
        root: {
          backgroundColor: '#181b23',
          border: '1px solid #4a3a58',
          borderRadius: '9px',
          boxShadow: 'none',
          transition: 'border-color 0.15s ease',
          '&:hover': {
            boxShadow: 'none',
            borderColor: '#6b5578',
          },
          '& .MuiCardContent-root': {
            color: '#e8eaef',
          },
        },
      },
    },
    MuiIconButton: {
      styleOverrides: {
        root: {
          color: '#9ea3c8',
          backgroundColor: 'transparent',
          border: 'none',
          borderRadius: '10px',
          transition: 'background-color 0.15s ease, color 0.15s ease',
          '&:hover': {
            backgroundColor: 'rgba(183, 148, 246, 0.12)',
            color: '#b794f6',
          },
          '&.Mui-disabled': {
            color: '#5a5f8a',
            backgroundColor: 'transparent',
          },
          '& .MuiSvgIcon-root': {
            color: 'inherit',
          },
        },
      },
    },
    MuiButton: {
      styleOverrides: {
        root: {
          borderRadius: '9px',
          textTransform: 'none',
          fontWeight: 600,
          boxShadow: 'none',
          '&:hover': {
            boxShadow: 'none',
          },
        },
        contained: {
          backgroundColor: '#b794f6',
          color: '#0f1117',
          '&:hover': {
            backgroundColor: '#dbb2ff',
          },
        },
        outlined: {
          borderWidth: '1.5px',
          borderColor: '#4a3a58',
        },
      },
    },
    MuiTypography: {
      styleOverrides: {
        root: {
          color: '#e8eaef',
        },
        caption: {
          color: '#9ea3c8',
        },
        h6: {
          color: '#e8eaef',
          fontWeight: 600,
        },
      },
    },
    MuiAlert: {
      styleOverrides: {
        root: {
          backgroundColor: '#181b23',
          border: '1px solid #4a3a58',
          color: '#e8eaef',
          borderRadius: '9px',
        },
      },
    },
    MuiTextField: {
      styleOverrides: {
        root: {
          '& .MuiInputLabel-root': {
            color: '#9ea3c8',
          },
          '& .MuiInputBase-input': {
            color: '#e8eaef',
          },
          '& .MuiOutlinedInput-root': {
            borderRadius: '9px',
            '& fieldset': {
              borderColor: '#4a3a58',
            },
            '&:hover fieldset': {
              borderColor: '#6b5578',
            },
            '&.Mui-focused fieldset': {
              borderColor: '#b794f6',
            },
          },
        },
      },
    },
    MuiChip: {
      styleOverrides: {
        root: {
          borderRadius: '999px',
          fontWeight: 600,
          fontSize: '0.7rem',
          backgroundColor: '#2a2035',
          color: '#e8eaef',
          border: '1px solid #4a3a58',
        },
        filled: {
          backgroundColor: '#2a2035',
          color: '#e8eaef',
        },
      },
    },
    MuiAutocomplete: {
      styleOverrides: {
        paper: {
          backgroundColor: '#181b23',
          border: '1px solid #4a3a58',
          borderRadius: '9px',
        },
        option: {
          color: '#e8eaef',
          '&:hover': {
            backgroundColor: 'rgba(183, 148, 246, 0.1)',
          },
          '&[aria-selected="true"]': {
            backgroundColor: 'rgba(183, 148, 246, 0.16)',
          },
        },
        listbox: {
          backgroundColor: '#181b23',
        },
        groupLabel: {
          backgroundColor: '#2a2035 !important',
          color: '#e8eaef !important',
          fontWeight: 600,
        },
      },
    },
    MuiPaper: {
      styleOverrides: {
        root: {
          backgroundColor: '#181b23',
          color: '#e8eaef',
          borderRadius: '9px',
        },
      },
    },
    MuiListSubheader: {
      styleOverrides: {
        root: {
          backgroundColor: '#2a2035 !important',
          color: '#e8eaef !important',
          fontWeight: 600,
        },
      },
    },
    MuiMenu: {
      styleOverrides: {
        paper: {
          backgroundColor: '#181b23 !important',
          border: '1px solid #4a3a58',
          borderRadius: '9px',
          boxShadow: '0 8px 24px rgba(0, 0, 0, 0.45)',
          '& .MuiMenuItem-root': {
            color: '#e8eaef',
            fontSize: '0.8125rem',
            '&:hover': {
              backgroundColor: 'rgba(183, 148, 246, 0.1)',
            },
          },
        },
      },
    },
    MuiSelect: {
      styleOverrides: {
        root: {
          '& .MuiSelect-select': {
            color: '#e8eaef !important',
          },
          '& .MuiSelect-icon': {
            color: '#9ea3c8 !important',
          },
        },
      },
    },
    MuiInputBase: {
      styleOverrides: {
        root: {
          '& .MuiInputBase-input': {
            color: '#e8eaef !important',
          },
        },
      },
    },
  },
});

interface ThemeProviderProps {
  children: ReactNode;
}

function applyDocumentTheme(next: ThemeMode) {
  try {
    document.documentElement.classList.toggle('spo-theme-dark', next === 'dark');
    document.documentElement.dataset.spoTheme = next;
  } catch {
    // ignore
  }
}

export const ThemeProvider: React.FC<ThemeProviderProps> = ({ children }) => {
  const [mode, setMode] = useState<ThemeMode>('light');

  const saveTheme = async (newMode: ThemeMode) => {
    try {
      if (typeof chrome !== 'undefined' && chrome.storage?.local) {
        await chrome.storage.local.set({
          [STORAGE_KEY]: newMode,
          [SPO_DARK_MODE_KEY]: newMode === 'dark',
        });
      }
      localStorage.setItem(STORAGE_KEY, newMode);
    } catch (error) {
      console.error('Error saving theme preference:', error);
    }
  };

  useEffect(() => {
    const loadTheme = async () => {
      try {
        if (typeof chrome !== 'undefined' && chrome.storage?.local) {
          const result = await chrome.storage.local.get([STORAGE_KEY, SPO_DARK_MODE_KEY]);
          // Prefer shared SPO darkMode so popup + Power Apps stay in sync
          if (typeof result[SPO_DARK_MODE_KEY] === 'boolean') {
            const next: ThemeMode = result[SPO_DARK_MODE_KEY] ? 'dark' : 'light';
            setMode(next);
            applyDocumentTheme(next);
            return;
          }
          if (result[STORAGE_KEY] === 'light' || result[STORAGE_KEY] === 'dark') {
            setMode(result[STORAGE_KEY] as ThemeMode);
            applyDocumentTheme(result[STORAGE_KEY] as ThemeMode);
            return;
          }
        }

        const saved = localStorage.getItem(STORAGE_KEY);
        if (saved === 'light' || saved === 'dark') {
          setMode(saved);
          applyDocumentTheme(saved);
          return;
        }

        if (typeof window !== 'undefined' && window.matchMedia) {
          const systemTheme: ThemeMode = window.matchMedia('(prefers-color-scheme: dark)').matches
            ? 'dark'
            : 'light';
          setMode(systemTheme);
          applyDocumentTheme(systemTheme);
          saveTheme(systemTheme);
        }
      } catch (error) {
        console.error('Error loading theme preference:', error);
      }
    };

    loadTheme();
  }, []);

  // Follow SPO popup theme changes while the sidebar is open
  useEffect(() => {
    if (typeof chrome === 'undefined' || !chrome.storage?.onChanged) return;
    const onChange = (
      changes: { [key: string]: chrome.storage.StorageChange },
      area: string
    ) => {
      if (area !== 'local' || !changes[SPO_DARK_MODE_KEY]) return;
      const isDark = !!changes[SPO_DARK_MODE_KEY].newValue;
      const next: ThemeMode = isDark ? 'dark' : 'light';
      setMode(next);
      applyDocumentTheme(next);
    };
    chrome.storage.onChanged.addListener(onChange);
    return () => chrome.storage.onChanged.removeListener(onChange);
  }, []);

  useEffect(() => {
    applyDocumentTheme(mode);
  }, [mode]);

  const toggleTheme = () => {
    const newMode = mode === 'light' ? 'dark' : 'light';
    setMode(newMode);
    applyDocumentTheme(newMode);
    saveTheme(newMode);
  };

  const theme = mode === 'light' ? lightTheme : darkTheme;
  const isDarkMode = mode === 'dark';

  const contextValue: ThemeContextType = {
    mode,
    toggleTheme,
    isDarkMode,
  };

  return (
    <ThemeContext.Provider value={contextValue}>
      <MuiThemeProvider theme={theme}>
        <CssBaseline />
        {children}
      </MuiThemeProvider>
    </ThemeContext.Provider>
  );
};

export const useTheme = (): ThemeContextType => {
  const context = useContext(ThemeContext);
  if (context === undefined) {
    throw new Error('useTheme must be used within a ThemeProvider');
  }
  return context;
};
