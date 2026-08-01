import React from 'react';
import { useTheme } from '#contexts/ThemeContext';

/**
 * SPO-style Early Riser / Night Owl pill (matches popup.html .dark-toggle-btn).
 */
const ThemeSwitchButtons: React.FC = () => {
  const { mode, toggleTheme } = useTheme();
  const isDark = mode === 'dark';

  return (
    <button
      type='button'
      className={'spo-dark-toggle-btn' + (isDark ? ' spo-dark-toggle-btn--night' : '')}
      onClick={toggleTheme}
      title={isDark ? 'Night Owl' : 'Early Riser'}
      aria-label={
        isDark
          ? 'Theme: Night Owl. Click to switch to Early Riser.'
          : 'Theme: Early Riser. Click to switch to Night Owl.'
      }
    >
      <span className='spo-dark-toggle-knob' aria-hidden='true' />
    </button>
  );
};

export default ThemeSwitchButtons;
