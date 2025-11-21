import { useState, useEffect, useCallback, useRef } from 'react';

export type KeyTipState = 'idle' | 'awaitingTab' | 'awaitingCommand' | 'executed';

export interface KeyTipMapping {
  key: string;
  label: string;
  handler: () => void;
  tab?: string;
}

export interface KeyTipSettings {
  enabled: boolean;
  macModifier: 'ctrlOption' | 'alt' | 'cmd';
}

const DEFAULT_SETTINGS: KeyTipSettings = {
  enabled: true,
  macModifier: 'ctrlOption',
};

export function useKeyTips() {
  const [state, setState] = useState<KeyTipState>('idle');
  const [currentSequence, setCurrentSequence] = useState<string[]>([]);
  const [activeTab, setActiveTab] = useState<string | null>(null);
  const [settings, setSettings] = useState<KeyTipSettings>(() => {
    const stored = localStorage.getItem('keyTipSettings');
    return stored ? JSON.parse(stored) : DEFAULT_SETTINGS;
  });
  
  const timeoutRef = useRef<NodeJS.Timeout | null>(null);
  const mappingsRef = useRef<Map<string, KeyTipMapping>>(new Map());

  const isMac = /Mac|iPhone|iPad|iPod/.test(navigator.userAgent);

  const saveSettings = useCallback((newSettings: KeyTipSettings) => {
    setSettings(newSettings);
    localStorage.setItem('keyTipSettings', JSON.stringify(newSettings));
  }, []);

  const registerKeyTip = useCallback((id: string, mapping: KeyTipMapping) => {
    mappingsRef.current.set(id, mapping);
  }, []);

  const unregisterKeyTip = useCallback((id: string) => {
    mappingsRef.current.delete(id);
  }, []);

  const clearKeyTips = useCallback(() => {
    mappingsRef.current.clear();
  }, []);

  const resetState = useCallback(() => {
    setState('idle');
    setCurrentSequence([]);
    setActiveTab(null);
    if (timeoutRef.current) {
      clearTimeout(timeoutRef.current);
      timeoutRef.current = null;
    }
  }, []);

  const startTimeout = useCallback(() => {
    if (timeoutRef.current) {
      clearTimeout(timeoutRef.current);
    }
    timeoutRef.current = setTimeout(() => {
      resetState();
    }, 2500);
  }, [resetState]);

  const handleKeyPress = useCallback((key: string) => {
    if (!settings.enabled) return;

    const upperKey = key.toUpperCase();

    if (state === 'idle') {
      return;
    }

    if (state === 'awaitingTab') {
      // Look for tab mappings
      const tabMapping = Array.from(mappingsRef.current.values()).find(
        (m) => !m.tab && m.key === upperKey
      );

      if (tabMapping) {
        setActiveTab(upperKey);
        setCurrentSequence([...currentSequence, upperKey]);
        setState('awaitingCommand');
        startTimeout();
        tabMapping.handler();
      } else {
        resetState();
      }
    } else if (state === 'awaitingCommand' && activeTab) {
      // Look for command mappings within the active tab
      const commandMapping = Array.from(mappingsRef.current.values()).find(
        (m) => m.tab === activeTab && m.key === upperKey
      );

      if (commandMapping) {
        setCurrentSequence([...currentSequence, upperKey]);
        setState('executed');
        commandMapping.handler();
        setTimeout(() => resetState(), 300);
      } else {
        resetState();
      }
    }
  }, [state, activeTab, currentSequence, settings.enabled, startTimeout, resetState]);

  useEffect(() => {
    if (!settings.enabled) return;

    const handleKeyDown = (e: KeyboardEvent) => {
      // Check if Alt key is pressed (or Ctrl+Option on Mac)
      const isActivationKey = isMac
        ? settings.macModifier === 'ctrlOption'
          ? e.ctrlKey && e.altKey && !e.metaKey
          : settings.macModifier === 'alt'
          ? e.altKey && !e.ctrlKey && !e.metaKey
          : e.metaKey && !e.ctrlKey && !e.altKey
        : e.altKey && !e.ctrlKey && !e.metaKey;

      if (isActivationKey && state === 'idle' && !e.repeat) {
        e.preventDefault();
        setState('awaitingTab');
        setCurrentSequence(['Alt']);
        startTimeout();
        return;
      }

      // Handle Escape to cancel
      if (e.key === 'Escape' && state !== 'idle') {
        e.preventDefault();
        resetState();
        return;
      }

      // Handle letter keys when in KeyTip mode
      if (state !== 'idle' && /^[a-zA-Z]$/.test(e.key)) {
        e.preventDefault();
        handleKeyPress(e.key);
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [state, settings, isMac, handleKeyPress, startTimeout, resetState]);

  return {
    state,
    currentSequence,
    activeTab,
    settings,
    isActive: state !== 'idle',
    registerKeyTip,
    unregisterKeyTip,
    clearKeyTips,
    resetState,
    saveSettings,
  };
}
