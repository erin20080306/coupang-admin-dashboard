import { useLayoutEffect, useMemo, useRef } from 'react';
import gsap from 'gsap';
import type { AttendanceSummary } from '../lib/mockApi';

type Item = { id: string; name: string; summary: AttendanceSummary };

function labelFor(status: AttendanceSummary['status']): string {
  if (status === 'normal') return '正常';
  if (status === 'low') return '偏低';
  return '異常';
}

export default function AttendanceCards({
  items,
  onItemClick,
}: {
  items: Item[];
  onItemClick?: (item: Item) => void;
}) {
  const wrapRef = useRef<HTMLDivElement | null>(null);

  const normalized = useMemo(() => items.map((it) => ({
    ...it,
    pct: Math.round(it.summary.rate * 100),
  })), [items]);

  useLayoutEffect(() => {
    if (!wrapRef.current) return;
    const bars = Array.from(wrapRef.current.querySelectorAll<HTMLElement>('[data-bar]'));
    gsap.fromTo(
      bars,
      { width: '0%' },
      { width: (i) => `${normalized[i]?.pct ?? 0}%`, duration: 0.9, ease: 'power2.out', stagger: 0.05 }
    );
  }, [normalized]);

  return (
    <div className="attList" ref={wrapRef}>
      {normalized.map((it) => (
        <div
          className="attCard"
          key={it.id}
          role={onItemClick ? 'button' : undefined}
          tabIndex={onItemClick ? 0 : undefined}
          onClick={onItemClick ? () => onItemClick(it) : undefined}
          onKeyDown={
            onItemClick
              ? (e) => {
                  if (e.key === 'Enter' || e.key === ' ') onItemClick(it);
                }
              : undefined
          }
          style={onItemClick ? { cursor: 'pointer' } : undefined}
        >
          <div className="attTop">
            <div className="attName">{it.name}</div>
            <div className={`badge badge-${it.summary.status}`}>{labelFor(it.summary.status)}</div>
          </div>
          <div className="attMeta">
            <div className="attPct">{it.pct}%</div>
            <div className="attCounts">
              {it.summary.attended}/{it.summary.expected}
            </div>
          </div>
          <div className="bar">
            <div className="barFill" data-bar style={{ width: `${it.pct}%` }} />
          </div>
        </div>
      ))}
    </div>
  );
}
