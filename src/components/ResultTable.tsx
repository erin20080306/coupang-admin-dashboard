import { useLayoutEffect, useMemo, useRef, useState } from 'react';

type SortDir = 'asc' | 'desc';

type KeyOf<T> = Extract<keyof T, string>;

export type ColumnDef<T> = {
  key: KeyOf<T>;
  header: string;
  sortable?: boolean;
  render?: (row: T) => React.ReactNode;
};

function cmp(a: unknown, b: unknown): number {
  const ax = typeof a === 'string' ? a : String(a ?? '');
  const bx = typeof b === 'string' ? b : String(b ?? '');
  return ax.localeCompare(bx, 'zh-Hant', { numeric: true, sensitivity: 'base' });
}

export default function ResultTable<T extends Record<string, unknown>>({
  columns,
  rows,
  tbodyRef,
  frozenLeft,
  cellStyle,
}: {
  columns: ColumnDef<T>[];
  rows: T[];
  tbodyRef?: React.RefObject<HTMLTableSectionElement | null>;
  frozenLeft?: number;
  cellStyle?: (args: { row: T; col: ColumnDef<T>; colIndex: number; rowIndex: number }) => React.CSSProperties | undefined;
}) {
  const [sortKey, setSortKey] = useState<string>(columns.find((c) => c.sortable)?.key ?? columns[0]?.key ?? '');
  const [sortDir, setSortDir] = useState<SortDir>('asc');
  const [filters, setFilters] = useState<Record<string, string>>({});

  const headCellRefs = useRef<Array<HTMLTableCellElement | null>>([]);
  const [leftOffsets, setLeftOffsets] = useState<number[]>([]);

  const filtered = useMemo(() => {
    const active = Object.entries(filters)
      .map(([k, v]) => [k, String(v ?? '').trim().toLowerCase()] as const)
      .filter((x) => x[1]);
    if (!active.length) return rows;
    return rows.filter((r) => {
      for (const [k, v] of active) {
        const raw = String((r as any)[k] ?? '').toLowerCase();
        if (!raw.includes(v)) return false;
      }
      return true;
    });
  }, [rows, filters]);

  const sorted = useMemo(() => {
    if (!sortKey) return filtered;
    const out = [...filtered];
    out.sort((ra, rb) => {
      const v = cmp(ra[sortKey], rb[sortKey]);
      return sortDir === 'asc' ? v : -v;
    });
    return out;
  }, [filtered, sortKey, sortDir]);

  function onSort(col: ColumnDef<T>) {
    if (!col.sortable) return;
    if (sortKey === col.key) {
      setSortDir((d) => (d === 'asc' ? 'desc' : 'asc'));
    } else {
      setSortKey(col.key);
      setSortDir('asc');
    }
  }

  useLayoutEffect(() => {
    const n = columns.length;
    const k = Math.max(0, Math.min(frozenLeft ?? 0, n));
    if (!k) {
      setLeftOffsets([]);
      return;
    }

    const widths: number[] = [];
    for (let i = 0; i < k; i++) {
      const el = headCellRefs.current[i];
      widths[i] = el ? el.getBoundingClientRect().width : 0;
    }

    const offs: number[] = [];
    let acc = 0;
    for (let i = 0; i < k; i++) {
      offs[i] = acc;
      acc += widths[i] || 0;
    }
    setLeftOffsets(offs);
  }, [columns, frozenLeft, rows.length]);

  return (
    <div className="tableWrap">
      <table className="table">
        <thead>
          <tr>
            {columns.map((c) => (
              (() => {
                const colIndex = columns.findIndex((x) => x.key === c.key);
                const isFrozen = Boolean(frozenLeft && colIndex < frozenLeft);
                const left = isFrozen ? leftOffsets[colIndex] ?? colIndex * 0 : undefined;
                return (
              <th
                key={c.key}
                className={c.sortable ? 'thSortable' : undefined}
                onClick={() => onSort(c)}
                ref={(el) => {
                  headCellRefs.current[colIndex] = el;
                }}
                style={
                  isFrozen
                    ? { position: 'sticky', left: left != null ? `${left}px` : undefined, zIndex: 3 }
                    : undefined
                }
              >
                <span>{c.header}</span>
                {c.sortable && sortKey === c.key ? (
                  <span className="sort">{sortDir === 'asc' ? '↑' : '↓'}</span>
                ) : null}
              </th>
                );
              })()
            ))}
          </tr>
          <tr className="filterRow">
            {columns.map((c) => (
              (() => {
                const colIndex = columns.findIndex((x) => x.key === c.key);
                const isFrozen = Boolean(frozenLeft && colIndex < frozenLeft);
                const left = isFrozen ? leftOffsets[colIndex] ?? colIndex * 0 : undefined;
                return (
                  <th
                    key={c.key}
                    style={
                      isFrozen
                        ? { position: 'sticky', left: left != null ? `${left}px` : undefined, zIndex: 3, background: 'white' }
                        : undefined
                    }
                  >
                    <input
                      className="thFilterInput"
                      value={filters[c.key] ?? ''}
                      onChange={(e) => setFilters((s) => ({ ...s, [c.key]: e.target.value }))}
                      onClick={(e) => e.stopPropagation()}
                      placeholder="篩選"
                    />
                  </th>
                );
              })()
            ))}
          </tr>
        </thead>
        <tbody ref={tbodyRef ?? undefined}>
          {sorted.map((r, idx) => (
            <tr key={(r as any).id ?? idx}>
              {columns.map((c) => (
                (() => {
                  const colIndex = columns.findIndex((x) => x.key === c.key);
                  const isFrozen = Boolean(frozenLeft && colIndex < frozenLeft);
                  const left = isFrozen ? leftOffsets[colIndex] ?? colIndex * 0 : undefined;
                  return (
                <td
                  key={c.key}
                  style={{
                    ...(cellStyle ? (cellStyle({ row: r, col: c, colIndex, rowIndex: idx }) ?? {}) : {}),
                    ...(isFrozen
                      ? { position: 'sticky', left: left != null ? `${left}px` : undefined, zIndex: 2, background: 'white' }
                      : {}),
                  }}
                >
                  {c.render ? c.render(r) : String(r[c.key] ?? '')}
                </td>
                  );
                })()
              ))}
            </tr>
          ))}
        </tbody>
      </table>

      <div className="cardList">
        {sorted.map((r, idx) => (
          <div className="rowCard" key={(r as any).id ?? idx}>
            {columns.map((c) => (
              <div className="rowCardLine" key={c.key}>
                <div className="rowCardLabel">{c.header}</div>
                <div className="rowCardValue">{c.render ? c.render(r) : String(r[c.key] ?? '')}</div>
              </div>
            ))}
          </div>
        ))}
      </div>
    </div>
  );
}
