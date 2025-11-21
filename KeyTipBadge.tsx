import { useEffect, useRef } from 'react';

interface KeyTipBadgeProps {
  keyTip: string;
  visible: boolean;
  targetRef: React.RefObject<HTMLElement | null>;
}

export function KeyTipBadge({ keyTip, visible, targetRef }: KeyTipBadgeProps) {
  const badgeRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!visible || !targetRef.current || !badgeRef.current) return;

    const target = targetRef.current;
    const badge = badgeRef.current;
    const rect = target.getBoundingClientRect();

    badge.style.left = `${rect.left + rect.width / 2}px`;
    badge.style.top = `${rect.top - 8}px`;
  }, [visible, targetRef]);

  if (!visible) return null;

  return (
    <div
      ref={badgeRef}
      className="keytip-badge"
      role="tooltip"
      aria-label={`Press ${keyTip} to activate`}
    >
      {keyTip}
    </div>
  );
}
