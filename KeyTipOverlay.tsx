interface KeyTipOverlayProps {
  sequence: string[];
  visible: boolean;
}

export function KeyTipOverlay({ sequence, visible }: KeyTipOverlayProps) {
  if (!visible || sequence.length === 0) return null;

  return (
    <div
      className="fixed top-4 left-1/2 -translate-x-1/2 z-50 bg-card border border-border rounded-lg shadow-lg px-4 py-2"
      role="status"
      aria-live="polite"
      aria-atomic="true"
    >
      <div className="flex items-center gap-2 text-sm font-medium">
        {sequence.map((key, index) => (
          <span key={index} className="flex items-center gap-2">
            <kbd className="px-2 py-1 bg-muted rounded text-xs font-mono border border-border">
              {key}
            </kbd>
            {index < sequence.length - 1 && (
              <span className="text-muted-foreground">→</span>
            )}
          </span>
        ))}
      </div>
    </div>
  );
}
