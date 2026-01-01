import { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import gsap from 'gsap';
import { useNavigate } from 'react-router-dom';
import { getUser, logout } from '../lib/auth';
import { mockPages, mockWarehouses, queryPeople, type AttendanceSummary, type PersonRow, type QueryParams, type QueryResult } from '../lib/mockApi';
import { gasGetSheets, gasGetWarehouseId, gasIsConfigured, gasQuerySheet } from '../lib/gasApi';
import { gasPayloadToRows, type GasRecordRow } from '../lib/transformGas';
import AttendanceCards from '../components/AttendanceCards';
import ResultTable, { type ColumnDef } from '../components/ResultTable';
import SkeletonTable from '../components/SkeletonTable';
import EmptyState from '../components/EmptyState';
import { exportCsv, exportElementPng, exportExcelHtml } from '../lib/export';

type DisplayRow = PersonRow | GasRecordRow;

const PAGES_CACHE = new Map<string, { ts: number; pages: string[] }>();

const PAGES_CACHE_MS = 20_000;

function getAttendanceFromRow(r: DisplayRow): AttendanceSummary | undefined {
  if ((r as any)._attendance) return (r as any)._attendance as AttendanceSummary;
  if ((r as any).attendance) return (r as any).attendance as AttendanceSummary;
  return undefined;
}

function DeptKpi({
  rows,
  headers,
}: {
  rows: DisplayRow[];
  headers: string[];
}) {
  const deptKey = useMemo(() => findDeptKey(headers), [headers]);

  const stat = useMemo(() => {
    if (!deptKey) return null;
    const byDept = new Map<string, Set<string>>();
    for (const r of rows) {
      const dept = String((r as any)[deptKey] ?? '').trim() || '未填部門';
      const name = getNameFromRow(r);
      if (!byDept.has(dept)) byDept.set(dept, new Set());
      byDept.get(dept)!.add(name);
    }
    const deptList = Array.from(byDept.entries())
      .map(([dept, set]) => ({ dept, count: set.size }))
      .sort((a, b) => b.count - a.count);

    const totalPeople = deptList.reduce((acc, x) => acc + x.count, 0);
    return { deptList, totalPeople };
  }, [rows, deptKey]);

  if (!stat) {
    return (
      <>
        <div className="kpi">
          <div className="kpiLabel">部門</div>
          <div className="kpiValue">—</div>
        </div>
        <div className="kpi">
          <div className="kpiLabel">人數</div>
          <div className="kpiValue">—</div>
        </div>
        <div className="kpi">
          <div className="kpiLabel">部門別</div>
          <div className="kpiValue">—</div>
        </div>
      </>
    );
  }

  const topNames = stat.deptList.slice(0, 4).map((x) => x.dept).join('、');

  return (
    <>
      <div className="kpi">
        <div className="kpiLabel">部門數</div>
        <div className="kpiValue">{stat.deptList.length}</div>
      </div>
      <div className="kpi">
        <div className="kpiLabel">分頁人數</div>
        <div className="kpiValue">{stat.totalPeople}</div>
      </div>
      <div className="kpi">
        <div className="kpiLabel">部門別</div>
        <div className="kpiValue">{topNames || '—'}</div>
      </div>
    </>
  );
}

function LeaveStatPanel({
  headers,
  headersISO,
  dateCols,
  rowsAll,
  rowsSingle,
  isAdmin,
  userName,
  sheetName,
}: {
  headers: string[];
  headersISO: string[];
  dateCols: number[];
  rowsAll: GasRecordRow[];
  rowsSingle: GasRecordRow[];
  isAdmin: boolean;
  userName: string;
  sheetName: string;
}) {
  const isSchedule = sheetName.includes('班表');
  const isRecord = sheetName.includes('出勤記錄');

  const [open, setOpen] = useState(false);

  const [mode, setMode] = useState<'single' | 'all'>('single');
  const [singleName, setSingleName] = useState('');
  const [leaveFilter, setLeaveFilter] = useState('');

  const effectiveMode = isAdmin ? mode : 'single';

  const effectiveSingleName = isAdmin ? singleName.trim() : (userName || '').trim();

  const rows = useMemo(() => {
    if (effectiveMode === 'all') return rowsAll;
    if (effectiveSingleName) {
      return rowsAll.filter((r) => getNameFromRow(r) === effectiveSingleName);
    }
    return rowsSingle;
  }, [effectiveMode, rowsAll, rowsSingle, effectiveSingleName]);

  const data = useMemo(() => {
    if (!rows?.length) return [] as Array<{ tag: string; dates: string; count: number }>;

    const tagDatesMap = new Map<string, Set<string>>();

    if (isSchedule && Array.isArray(dateCols) && dateCols.length) {
      for (const ci of dateCols) {
        const iso = headersISO?.[ci] || '';
        const md = iso ? fmtMDFromISO(iso) : String(headers?.[ci] || '');
        for (const row of rows) {
          const toks = tokenizeCell((row as any)[headers[ci]]);
          for (const tag of toks) {
            if (!tagDatesMap.has(tag)) tagDatesMap.set(tag, new Set());
            tagDatesMap.get(tag)!.add(md);
          }
        }
      }
    } else if (isRecord) {
      const dateKey = findDateKey(headers);
      const leaveKey = findLeaveKey(headers);
      if (!dateKey || !leaveKey) return [];
      for (const row of rows) {
        const ds = String((row as any)[dateKey] ?? '').trim();
        const iso = guessISOFromText(ds);
        const md = iso ? fmtMDFromISO(iso) : ds;
        const tags = tokenizeCell((row as any)[leaveKey]);
        for (const tag of tags) {
          if (!tagDatesMap.has(tag)) tagDatesMap.set(tag, new Set());
          tagDatesMap.get(tag)!.add(md);
        }
      }
    } else {
      return [];
    }

    const out = Array.from(tagDatesMap.entries())
      .map(([tag, set]) => {
        const list = Array.from(set);
        list.sort((a, b) => a.localeCompare(b, 'zh-Hant', { numeric: true, sensitivity: 'base' }));
        return { tag, dates: list.join('、'), count: list.length };
      })
      .sort((a, b) => a.tag.localeCompare(b.tag, 'zh-Hant', { numeric: true, sensitivity: 'base' }));
    return out;
  }, [rows, headers, headersISO, dateCols, isSchedule, isRecord]);

  const leaveOptions = useMemo(() => {
    const set = new Set<string>();
    for (const r of data) set.add(r.tag);
    return Array.from(set).sort((a, b) => a.localeCompare(b, 'zh-Hant', { numeric: true, sensitivity: 'base' }));
  }, [data]);

  const dataShown = useMemo(() => {
    const f = leaveFilter.trim();
    if (!f) return data;
    return data.filter((x) => x.tag === f);
  }, [data, leaveFilter]);

  const wrapRef = useRef<HTMLDivElement | null>(null);

  const exportBase = useMemo(() => {
    const modeLabel = effectiveMode === 'all' ? '全員假別統計' : '單人假別統計';
    const nameKey = effectiveMode === 'all' ? '' : (effectiveSingleName || '');
    const deptKey = findDeptKey(headers);
    const shiftKey = findShiftKey(headers);

    let deptLabel = '';
    let shiftLabel = '';
    if (deptKey && nameKey) {
      const set = new Set<string>();
      rowsAll.forEach((r) => {
        if (getNameFromRow(r) !== nameKey) return;
        const v = String((r as any)[deptKey] ?? '').trim();
        if (v) set.add(v);
      });
      deptLabel = Array.from(set).join('、');
    }
    if (shiftKey && nameKey) {
      const set = new Set<string>();
      rowsAll.forEach((r) => {
        if (getNameFromRow(r) !== nameKey) return;
        const v = String((r as any)[shiftKey] ?? '').trim();
        if (v) set.add(v);
      });
      shiftLabel = Array.from(set).join('、');
    }

    const parts = [String(sheetName || '').trim(), deptLabel, shiftLabel, nameKey, modeLabel].filter(Boolean);
    return parts.join('_');
  }, [effectiveMode, effectiveSingleName, headers, rowsAll, sheetName]);

  const exportRows = useMemo(() => {
    const modeLabel = effectiveMode === 'all' ? '全員' : '單人';
    const nameKey = effectiveMode === 'all' ? '' : (effectiveSingleName || '');
    const deptKey = findDeptKey(headers);
    const shiftKey = findShiftKey(headers);

    let deptLabel = '';
    let shiftLabel = '';
    if (deptKey && nameKey) {
      const set = new Set<string>();
      rowsAll.forEach((r) => {
        if (getNameFromRow(r) !== nameKey) return;
        const v = String((r as any)[deptKey] ?? '').trim();
        if (v) set.add(v);
      });
      deptLabel = Array.from(set).join('、');
    }
    if (shiftKey && nameKey) {
      const set = new Set<string>();
      rowsAll.forEach((r) => {
        if (getNameFromRow(r) !== nameKey) return;
        const v = String((r as any)[shiftKey] ?? '').trim();
        if (v) set.add(v);
      });
      shiftLabel = Array.from(set).join('、');
    }

    return dataShown.map((x) => [deptLabel, shiftLabel, nameKey, x.tag, x.dates, x.count, modeLabel]);
  }, [dataShown, effectiveMode, effectiveSingleName, headers, rowsAll]);

  if (!isSchedule && !isRecord) return null;

  return (
    <section className="panel">
      <div className="panelTitle" style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 10 }}>
        <span>假別統計</span>
        <div style={{ display: 'flex', gap: 10, alignItems: 'center' }}>
          {isAdmin ? (
            <>
              <button
                className="btnGhost"
                onClick={() => setMode('single')}
                aria-pressed={effectiveMode === 'single'}
              >
                單人
              </button>
              <button
                className="btnGhost"
                onClick={() => setMode('all')}
                aria-pressed={effectiveMode === 'all'}
              >
                全員
              </button>
            </>
          ) : null}
          <button className="btnGhost" onClick={() => setOpen((v) => !v)}>
            {open ? '收合' : '點開'}
          </button>
        </div>
      </div>

      {!open ? (
        <div style={{ color: 'var(--muted)', fontSize: 13 }}>點開即可查看假別統計</div>
      ) : (
        <>
          {effectiveMode === 'single' ? (
            <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap', marginBottom: 10 }}>
              {isAdmin ? (
                <label className="filter" style={{ minWidth: 220 }}>
                  <span>姓名</span>
                  <input
                    value={singleName}
                    onChange={(e) => setSingleName(e.target.value)}
                    placeholder="輸入姓名（留空=依搜尋結果）"
                  />
                </label>
              ) : null}

              <label className="filter" style={{ minWidth: 220 }}>
                <span>假別</span>
                <select value={leaveFilter} onChange={(e) => setLeaveFilter(e.target.value)}>
                  <option value="">全部</option>
                  {leaveOptions.map((x) => (
                    <option key={x} value={x}>{x}</option>
                  ))}
                </select>
              </label>
            </div>
          ) : (
            <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap', marginBottom: 10 }}>
              <label className="filter" style={{ minWidth: 220 }}>
                <span>假別</span>
                <select value={leaveFilter} onChange={(e) => setLeaveFilter(e.target.value)}>
                  <option value="">全部</option>
                  {leaveOptions.map((x) => (
                    <option key={x} value={x}>{x}</option>
                  ))}
                </select>
              </label>
            </div>
          )}

          <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', marginBottom: 10 }}>
            <button
              className="btnGhost"
              onClick={() => exportExcelHtml(exportBase, ['部門', '班別', '姓名', '假別', '日期', '天數'], exportRows.map((r) => [r[0], r[1], r[2], r[3], r[4], r[5]]))}
              disabled={!dataShown.length}
            >
              Excel
            </button>
            <button
              className="btnGhost"
              onClick={() => exportCsv(exportBase, ['部門', '班別', '姓名', '假別', '日期', '天數'], exportRows.map((r) => [r[0], r[1], r[2], r[3], r[4], r[5]]))}
              disabled={!dataShown.length}
            >
              CSV
            </button>
            <button
              className="btnGhost"
              onClick={() => {
                if (!wrapRef.current) return;
                exportElementPng(wrapRef.current, exportBase);
              }}
              disabled={!dataShown.length}
            >
              PNG
            </button>
          </div>

          {dataShown.length ? (
            <div className="tableWrap" ref={wrapRef}>
              <table className="table">
                <thead>
                  <tr>
                    <th>部門</th>
                    <th>班別</th>
                    <th>姓名</th>
                    <th>假別</th>
                    <th>日期</th>
                    <th>天數</th>
                  </tr>
                </thead>
                <tbody>
                  {exportRows.map((r) => (
                    <tr key={`${r[2]}_${r[3]}`}>
                      <td>{r[0]}</td>
                      <td>{r[1]}</td>
                      <td>{r[2]}</td>
                      <td>{r[3]}</td>
                      <td style={{ maxWidth: 520, overflow: 'hidden', textOverflow: 'ellipsis' }}>{r[4]}</td>
                      <td>{r[5]}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : (
            <div style={{ color: 'var(--muted)', fontSize: 13 }}>本分頁沒有可統計的假別資料。</div>
          )}
        </>
      )}
    </section>
  );
}

function getNameFromRow(r: DisplayRow): string {
  const v = (r as any).name ?? (r as any)['姓名'] ?? '';
  return String(v || '').trim() || '（未命名）';
}

function rowHasLeaveToken(r: DisplayRow): boolean {
  const v = Object.values(r as any).map((x) => String(x ?? '')).join(' ');
  return v.includes('離');
}

function tokenizeCell(raw: unknown): string[] {
  const s = (raw == null ? '' : String(raw)).trim();
  if (!s) return [];
  return s
    .split(/[、，,;／\/\n]+/)
    .map((x) => x.trim())
    .filter(Boolean);
}

/** 排除分母（應到）的項目 - 對應舊版 EXCLUDE_FROM_DENOM */
function buildExcludeForAttRateSet(): Set<string> {
  return new Set<string>([
    '例',
    '例假',
    '例假日',
    '例休',
    '休',
    '休假',
    '休假日',
    '國',
    '離',
    '調倉',
    '調任',
    '轉正',
  ]);
}

/** 排除缺勤（自動算實到）的項目 - 對應舊版 EXCLUDE_FROM_ABS */
function buildExcludeFromAbsSet(): Set<string> {
  return new Set<string>([
    '例',
    '例假',
    '例假日',
    '例休',
    '休',
    '休假',
    '休假日',
    '國',
    '離',
    '調倉',
    '調任',
    '轉正',
    '未',
    '特',
  ]);
}

/** 純英數字（無中文）判斷 */
const ALNUM_RE = /^[A-Za-z0-9]+$/;
const HAS_CJK_RE = /[\u4e00-\u9fff]/;

/** 判斷單元格是否「自動算實到」（命中 EXCLUDE_FROM_ABS 或純英數字） */
function isAutoAttendCell(raw: unknown, excludeAbs: Set<string>): boolean {
  const s = (raw == null ? '' : String(raw)).trim();
  if (!s) return false;
  const tokens = tokenizeCell(s);
  // 命中排除缺勤項目
  if (tokens.some((t) => excludeAbs.has(t))) return true;
  // 純英數字（無中文）
  if (tokens.length === 1 && ALNUM_RE.test(tokens[0]) && !HAS_CJK_RE.test(tokens[0])) return true;
  return false;
}

function statusFromRate(rate: number): AttendanceSummary['status'] {
  if (rate >= 0.9) return 'normal';
  if (rate >= 0.75) return 'low';
  return 'abnormal';
}

function normalizeWarehouseKey(wh: string): string {
  return String(wh || '').trim().toUpperCase();
}

function pickWorstSourceSheet(warehouse: string, pages: string[]): string {
  const wh = normalizeWarehouseKey(warehouse);
  const list = Array.isArray(pages) ? pages : [];

  function pickFirst(pred: (p: string) => boolean): string {
    return list.find((p) => pred(String(p || '').trim())) || '';
  }

  if (wh === 'TAO1' || wh === 'TA01') {
    return (
      pickFirst((p) => p.includes('班表')) ||
      pickFirst((p) => p.includes('出勤紀律')) ||
      pickFirst((p) => p.includes('出勤記錄')) ||
      pickFirst((p) => p.includes('出勤')) ||
      list[0] ||
      ''
    );
  }

  return (
    pickFirst((p) => p.includes('班表')) ||
    pickFirst((p) => p.includes('出勤記錄')) ||
    pickFirst((p) => p.includes('出勤')) ||
    list[0] ||
    ''
  );
}

function isShouldAttendCell(raw: unknown, exclude: Set<string>): boolean {
  const s = (raw == null ? '' : String(raw)).trim();
  if (!s) return true;
  const toks = tokenizeCell(s);
  return !toks.some((t) => exclude.has(t));
}

/**
 * 判斷某一格是否「實到」：
 * 1. 先檢查是否「自動算實到」（命中 EXCLUDE_FROM_ABS 或純英數字）→ 直接算實到
 * 2. 再檢查 _att 陣列（GAS 後端從出勤時數/時間/備份分頁比對姓名+日期建立）
 */
function isActualAttendCell(
  colIndex: number,
  row: GasRecordRow,
  cellValue: unknown,
  excludeAbs: Set<string>
): boolean {
  // 1. 命中 EXCLUDE_FROM_ABS 或純英數字 → 自動算實到
  if (isAutoAttendCell(cellValue, excludeAbs)) {
    return true;
  }
  // 2. 檢查 _att 陣列（從出勤時數/時間/備份分頁建立）
  const att = (row as any)._att as number[] | undefined;
  if (Array.isArray(att) && att.length > 0) {
    return Boolean(att[colIndex]);
  }
  return false;
}

function findShiftKey(headers: string[]): string | null {
  for (const h of headers) {
    const s = String(h || '').trim();
    if (!s) continue;
    if (s.includes('班別') || s.includes('班') || /shift/i.test(s)) return s;
  }
  return null;
}

function guessISOFromText(s: string): string {
  const t = String(s || '').trim();
  if (!t) return '';
  let m = t.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (m) {
    const y = m[1];
    const mo = (`0${m[2]}`).slice(-2);
    const d = (`0${m[3]}`).slice(-2);
    return `${y}-${mo}-${d}`;
  }
  m = t.match(/^(\d{1,2})[\/\-](\d{1,2})$/);
  if (m) {
    const y = String(new Date().getFullYear());
    const mo = (`0${m[1]}`).slice(-2);
    const d = (`0${m[2]}`).slice(-2);
    return `${y}-${mo}-${d}`;
  }
  m = t.match(/^(?:(\d{4})年)?\s*(\d{1,2})月\s*(\d{1,2})日$/);
  if (m) {
    const y = m[1] || String(new Date().getFullYear());
    const mo = (`0${m[2]}`).slice(-2);
    const d = (`0${m[3]}`).slice(-2);
    return `${y}-${mo}-${d}`;
  }
  return '';
}

function fmtMDFromISO(iso: string): string {
  const m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return String(iso || '');
  return `${Number(m[2])}/${Number(m[3])}`;
}

function findDeptKey(headers: string[]): string | null {
  const exact = ['部門', '組別', '組', '部門別', 'Department', 'Dept', 'dept', 'group'];
  for (const h of headers) {
    const s = String(h || '').trim();
    if (!s) continue;
    if (exact.includes(s)) return s;
  }
  for (const h of headers) {
    const s = String(h || '').trim();
    if (!s) continue;
    if (s.includes('部門') || s.includes('組別') || /dept|department|group/i.test(s)) return s;
  }
  return null;
}

function findLeaveKey(headers: string[]): string | null {
  for (const h of headers) {
    const s = String(h || '').trim();
    if (!s) continue;
    if (s.includes('假別') || s.includes('狀態')) return s;
  }
  return null;
}

function findDateKey(headers: string[]): string | null {
  for (const h of headers) {
    const s = String(h || '').trim();
    if (!s) continue;
    if (s.includes('日期')) return s;
  }
  return null;
}

function buildMockColumns(): ColumnDef<PersonRow>[] {
  return [
    { key: 'warehouse', header: '倉別', sortable: true },
    { key: 'page', header: '分頁' },
    { key: 'name', header: '姓名', sortable: true },
    { key: 'birthdayOrPhone', header: '生日/電話' },
    {
      key: 'attendance',
      header: '出勤率',
      render: (r: PersonRow) => `${Math.round(r.attendance.rate * 100)}% (${r.attendance.attended}/${r.attendance.expected})`,
    },
    { key: 'late', header: '遲到', sortable: true },
    { key: 'absent', header: '缺勤', sortable: true },
  ];
}

function buildGasColumns(headers: string[]): ColumnDef<GasRecordRow>[] {
  const cols: ColumnDef<GasRecordRow>[] = headers
    .filter((h) => String(h || '').trim())
    .map((h) => ({
      key: h as any,
      header: h,
      sortable: true,
    }));

  cols.push({
    key: '_attendanceRate' as any,
    header: '出勤率',
    sortable: true,
    render: (r: GasRecordRow) => {
      const a = (r as any)._attendance as AttendanceSummary | undefined;
      if (!a) return '—';
      return `${Math.round(a.rate * 100)}% (${a.attended}/${a.expected})`;
    },
  });

  return cols;
}

function buildGasColumnsNoAttendance(headers: string[]): ColumnDef<GasRecordRow>[] {
  const cols: ColumnDef<GasRecordRow>[] = headers
    .filter((h) => String(h || '').trim())
    .map((h) => ({
      key: h as any,
      header: h,
      sortable: true,
    }));

  cols.push({
    key: '_attendanceRate' as any,
    header: '出勤率',
    sortable: true,
    render: (r: GasRecordRow) => {
      const a = (r as any)._attendance as AttendanceSummary | undefined;
      if (!a) return '—';
      return `${Math.round(a.rate * 100)}% (${a.attended}/${a.expected})`;
    },
  });

  return cols;
}

export default function DashboardPage() {
  const navigate = useNavigate();
  const tbodyRef = useRef<HTMLTableSectionElement | null>(null);
  const mainTableWrapRef = useRef<HTMLDivElement | null>(null);
  const attSingleWrapRef = useRef<HTMLDivElement | null>(null);
  const attAllWrapRef = useRef<HTMLDivElement | null>(null);
  const useGas = gasIsConfigured();
  const user = getUser();
  const isAdmin = Boolean(user?.isAdmin);

  const [availablePages, setAvailablePages] = useState<string[]>(mockPages);
  const [query, setQuery] = useState<QueryParams>({
    warehouse: (!isAdmin && useGas && user?.warehouseKey) ? (user.warehouseKey || mockWarehouses[0]) : mockWarehouses[0],
    page: mockPages[0],
    name: '',
    birthdayOrPhone: '',
  });

  const [status, setStatus] = useState<'idle' | 'loading' | 'success' | 'error' | 'empty'>('idle');
  const [result, setResult] = useState<QueryResult | null>(null);
  const [error, setError] = useState<string>('');

  const isHoursPage = useMemo(() => query.page.trim() === '出勤時數', [query.page]);
  const isAttPage = useMemo(
    () => query.page.includes('班表') || query.page.includes('出勤記錄') || query.page.includes('出勤紀律'),
    [query.page]
  );

  const [search, setSearch] = useState('');
  const [attWorstOpen, setAttWorstOpen] = useState(false);
  const [attBestOpen, setAttBestOpen] = useState(false);

  const [openFreezePanel, setOpenFreezePanel] = useState(false);
  const [openAttStatPanel, setOpenAttStatPanel] = useState(false);
  const [openAttAllPanel, setOpenAttAllPanel] = useState(false);
  const [openLeaveFilterPanel, setOpenLeaveFilterPanel] = useState(false);

  const rows: DisplayRow[] = (result?.rows as unknown as DisplayRow[]) ?? [];
  const stats = result?.stats;

  const filteredRows = useMemo(() => {
    const q = String(search || '').trim().toLowerCase();
    if (!q) return rows;
    return rows.filter((r) => {
      const v = Object.values(r as any)
        .map((x) => String(x ?? ''))
        .join(' ')
        .toLowerCase();
      return v.includes(q);
    });
  }, [rows, search]);

  const [gasHeaders, setGasHeaders] = useState<string[]>([]);
  const [gasHeadersISO, setGasHeadersISO] = useState<string[]>([]);
  const [gasDateCols, setGasDateCols] = useState<number[]>([]);
  const [gasFrozenLeft, setGasFrozenLeft] = useState<number>(0);

  const [freezeStart, setFreezeStart] = useState<number>(0);
  const [freezeEnd, setFreezeEnd] = useState<number>(0);
  const [manualFrozenLeft, setManualFrozenLeft] = useState<number | null>(0);

  const [attEndDate, setAttEndDate] = useState<string>('');
  const [attSingleDate, setAttSingleDate] = useState<string>('');
  const [attName, setAttName] = useState<string>('');
  const [attSingleBuilt, setAttSingleBuilt] = useState<Array<[string, string, string, string, number, number, number, string]>>([]);
  const [attAllBuilt, setAttAllBuilt] = useState<Array<[string, string, string, number, number, string]>>([]);

  const [leaveTag, setLeaveTag] = useState<string>('');
  const [leaveName, setLeaveName] = useState<string>('');

  const [attWorstSourcePage, setAttWorstSourcePage] = useState<string>('');
  const [attAll, setAttAll] = useState<Array<{ id: string; name: string; summary: AttendanceSummary }>>([]);

  const attendanceAllAgg = useMemo(() => {
    let expected = 0;
    let attended = 0;
    let n = 0;
    for (const r of rows) {
      const a = getAttendanceFromRow(r);
      if (!a) continue;
      expected += a.expected;
      attended += a.attended;
      n += 1;
    }
    const rate = expected ? attended / expected : 0;
    return { expected, attended, rate, n };
  }, [rows]);

  const pageStats = useMemo(() => {
    const set = new Set<string>();
    const leaveSet = new Set<string>();
    for (const r of rows as DisplayRow[]) {
      const nm = getNameFromRow(r);
      if (!nm) continue;
      if (rowHasLeaveToken(r)) leaveSet.add(nm);
      set.add(nm);
    }
    return {
      total: set.size,
      leave: leaveSet.size,
      active: Math.max(0, set.size - leaveSet.size),
    };
  }, [rows]);

  const attWorst5 = useMemo(() => attAll.slice(0, 5), [attAll]);

  const attBest5 = useMemo(() => {
    if (!attAll.length) return [] as Array<{ id: string; name: string; summary: AttendanceSummary }>;
    const start = Math.max(0, attAll.length - 5);
    return attAll.slice(start).reverse();
  }, [attAll]);

  const hasAttendanceStats = useMemo(() => attAll.length > 0, [attAll.length]);

  function openAttendanceFullList() {
    try {
      sessionStorage.setItem(
        'coupang_att_full',
        JSON.stringify({
          warehouse: query.warehouse,
          page: attWorstSourcePage || query.page,
          items: attAll,
        })
      );
    } catch {
      // ignore
    }
    navigate('/attendance', { replace: false });
  }

  async function refreshWorstAttendance() {
    if (!useGas) return;
    if (!user) return;
    const source = pickWorstSourceSheet(query.warehouse, availablePages);
    if (!source) {
      setAttWorstSourcePage('');
      setAttAll([]);
      return;
    }

    const apiName = isAdmin ? '' : (user?.name || '');

    try {
      const payload = await gasQuerySheet(query.warehouse, source, apiName);
      const { headers, rows: gasRows } = gasPayloadToRows(payload, { disableAttendance: true, sheetName: source });
      const dateCols = Array.isArray(payload.dateCols) ? payload.dateCols : [];
      const headersISO = (payload.headersISO ?? []).map((h) => String(h ?? ''));

      // 如果 dateCols 是空的，從 headersISO 推測日期欄
      let effectiveDateCols = dateCols;
      if (!effectiveDateCols.length && headersISO.length) {
        effectiveDateCols = headersISO
          .map((iso, i) => (iso && iso.match(/^\d{4}-\d{2}-\d{2}$/) ? i : -1))
          .filter((i) => i >= 0);
      }

      if (!gasRows.length || !effectiveDateCols.length) {
        setAttWorstSourcePage(source);
        setAttAll([]);
        return;
      }

      const exclude = buildExcludeForAttRateSet();

      const byName = new Map<string, GasRecordRow[]>();
      gasRows.forEach((r) => {
        const nm = getNameFromRow(r);
        if (!nm) return;
        if (!byName.has(nm)) byName.set(nm, []);
        byName.get(nm)!.push(r);
      });

      const list: Array<{ id: string; name: string; summary: AttendanceSummary }> = [];
      byName.forEach((personRows, nm) => {
        // ✅ 同一人多列時，以「日期」去重，避免 expected 灌水
        const expectedSet = new Set<string>();
        const attendedSet = new Set<string>();

        const excludeAbs = buildExcludeFromAbsSet();
        for (const row of personRows) {
          for (const ci of effectiveDateCols) {
            const hk = headers[ci];
            if (!hk || !String(hk).trim()) continue;
            const cellValue = (row as any)[hk];
            if (!isShouldAttendCell(cellValue, exclude)) continue;

            const iso = String(headersISO?.[ci] || '').trim() || String(hk || '').trim() || String(ci);
            expectedSet.add(iso);
            if (isActualAttendCell(ci, row, cellValue, excludeAbs)) attendedSet.add(iso);
          }
        }

        const expected = expectedSet.size;
        const attended = attendedSet.size;
        if (!expected) return;
        const rate = attended / expected;
        list.push({
          id: `att_${nm}`,
          name: nm,
          summary: { rate, attended, expected, status: statusFromRate(rate) },
        });
      });

      list.sort((a, b) => a.summary.rate - b.summary.rate);
      setAttWorstSourcePage(source);
      setAttAll(list);
    } catch {
      setAttWorstSourcePage(source);
      setAttAll([]);
    }
  }

  useEffect(() => {
    if (!useGas) {
      setAvailablePages(mockPages);
      return;
    }
    const wk = user?.warehouseKey;
    if (!isAdmin && wk) {
      setQuery((s) => (s.warehouse === wk ? s : { ...s, warehouse: wk }));
    }
    refreshPages(query.warehouse);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [useGas, isAdmin, user?.warehouseKey]);

  useEffect(() => {
    if (!useGas) return;
    if (!user) return;
    if (!availablePages.length) return;
    if (isHoursPage) return;
    if (!isAttPage) return;

    const t = window.setTimeout(() => {
      void refreshWorstAttendance();
    }, 450);
    return () => window.clearTimeout(t);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [useGas, user?.name, isAdmin, query.warehouse, availablePages.join('|'), isHoursPage, isAttPage]);

  useEffect(() => {
    if (!useGas) return;
    if (!query.page.includes('班表')) return;
    if (!attAll.length) return;

    setResult((prev) => {
      if (!prev?.rows || !(prev.rows as any[]).length) return prev;

      const map = new Map<string, AttendanceSummary>();
      attAll.forEach((x) => {
        const nm = String(x?.name || '').trim();
        if (!nm) return;
        if (x?.summary) map.set(nm, x.summary);
      });
      if (!map.size) return prev;

      let changed = false;
      const newRows = (prev.rows as any[]).map((r) => {
        if ((r as any)?._attendance) return r;
        const nm = getNameFromRow(r as any);
        const s = map.get(nm);
        if (!s) return r;
        changed = true;
        return { ...r, _attendance: s, _attendanceRate: s.rate };
      });

      if (!changed) return prev;

      const stat = newRows.reduce(
        (acc, r) => {
          acc.total += 1;
          const a = (r as any)._attendance as AttendanceSummary | undefined;
          if (a) {
            acc.attended += a.attended;
            acc.absent += Math.max(0, a.expected - a.attended);
          }
          return acc;
        },
        { total: 0, attended: 0, late: 0, absent: 0 }
      );

      return { ...prev, rows: newRows as any, stats: stat };
    });
  }, [useGas, query.page, attAll]);

  useEffect(() => {
    if (!useGas) return;
    if (!user) return;
    if (status === 'loading') return;
    if (!query.page) return;
    doQuery();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [useGas, query.warehouse, query.page]);

  const columns = useMemo(() => {
    if (useGas && gasHeaders.length) {
      if (isHoursPage) {
        return buildGasColumnsNoAttendance(gasHeaders) as unknown as ColumnDef<DisplayRow>[];
      }
      return buildGasColumns(gasHeaders) as unknown as ColumnDef<DisplayRow>[];
    }
    return buildMockColumns() as unknown as ColumnDef<DisplayRow>[];
  }, [useGas, gasHeaders, isHoursPage]);

  async function refreshPages(warehouse: string) {
    if (!useGas) {
      setAvailablePages(mockPages);
      return;
    }

    const hit = PAGES_CACHE.get(warehouse);
    if (hit && Date.now() - hit.ts < PAGES_CACHE_MS) {
      const pages = hit.pages;
      setAvailablePages(pages.length ? pages : mockPages);
      setQuery((s) => ({ ...s, page: pages.includes(s.page) ? s.page : (pages[0] || s.page) }));
      return;
    }
    try {
      const pages = await gasGetSheets(warehouse);
      PAGES_CACHE.set(warehouse, { ts: Date.now(), pages });
      setAvailablePages(pages.length ? pages : mockPages);
      setQuery((s) => ({ ...s, page: pages[0] || s.page }));
    } catch {
      setAvailablePages(mockPages);
    }
  }

  async function doQuery() {
    setError('');
    setStatus('loading');
    setResult(null);

    try {
      if (useGas) {
        const apiName = isAdmin ? '' : (user?.name || '');
        const payload = await gasQuerySheet(query.warehouse, query.page, apiName);
        const headersISO = (payload.headersISO ?? []).map((h) => String(h ?? ''));
        setGasHeadersISO(headersISO);
        let dateCols = Array.isArray(payload.dateCols) ? payload.dateCols : [];
        
        // DEBUG
        console.log('[doQuery] payload.dateCols:', payload.dateCols, 'headersISO sample:', headersISO.slice(0, 10));
        
        // 如果 dateCols 是空的，從 headersISO 推測日期欄
        if (!dateCols.length && headersISO.length) {
          dateCols = headersISO
            .map((iso, i) => (iso && iso.match(/^\d{4}-\d{2}-\d{2}$/) ? i : -1))
            .filter((i) => i >= 0);
          console.log('[doQuery] inferred dateCols:', dateCols);
        }
        setGasDateCols(dateCols);
        setGasFrozenLeft(Number((payload as any).frozenLeft ?? 0) || 0);
        const { headers, rows: gasRows } = gasPayloadToRows(payload, {
          disableAttendance: isHoursPage,
          sheetName: query.page,
        });
        setGasHeaders(headers);

        const isSchedulePage = query.page.includes('班表');
        
        // DEBUG
        console.log('[doQuery] isHoursPage:', isHoursPage, 'isSchedulePage:', isSchedulePage, 'dateCols.length:', dateCols.length, 'gasRows.length:', gasRows.length);

        // 為所有列設定 _attendance（確保出勤率欄位永遠顯示）
        const exclude = buildExcludeForAttRateSet();
        const excludeAbs = buildExcludeFromAbsSet();

        gasRows.forEach((row) => {
          // 如果已經有 _attendance，跳過
          if ((row as any)._attendance) return;

          let expected = 0;
          let attended = 0;

          // 只有班表分頁才計算實際出勤率
          if (!isHoursPage && isSchedulePage && dateCols.length) {
            for (const ci of dateCols) {
              const hk = headers[ci];
              if (!hk || !String(hk).trim()) continue;
              const cellValue = (row as any)[hk];
              if (!isShouldAttendCell(cellValue, exclude)) continue;
              expected += 1;

              if (isActualAttendCell(ci, row, cellValue, excludeAbs)) attended += 1;
            }
          }

          // 無條件設定 _attendance，讓出勤率欄位能顯示
          const rate = expected > 0 ? attended / expected : 0;
          (row as any)._attendance = { rate, attended, expected, status: statusFromRate(rate) };
          (row as any)._attendanceRate = rate;
        });
        
        // DEBUG: 檢查第一列的 _attendance
        if (gasRows.length) {
          const first = gasRows[0] as any;
          console.log('[doQuery] first row _attendance:', first._attendance, 'name:', getNameFromRow(first as any));
        }

        if (!gasRows.length) {
          setStatus('empty');
          setResult({ rows: [] as any, stats: { total: 0, attended: 0, late: 0, absent: 0 } });
          return;
        }

        const stat = gasRows.reduce(
          (acc, r) => {
            acc.total += 1;
            const a = (r as any)._attendance as AttendanceSummary | undefined;
            if (a) {
              acc.attended += a.attended;
              acc.absent += Math.max(0, a.expected - a.attended);
            }
            return acc;
          },
          { total: 0, attended: 0, late: 0, absent: 0 }
        );

        setResult({ rows: gasRows as any, stats: stat });
        setStatus('success');
        setAttWorstOpen(false);
        setAttBestOpen(false);
        setManualFrozenLeft(0);
        setFreezeStart(0);
        setFreezeEnd(0);
        setAttEndDate('');
        setAttSingleDate('');
        setAttName('');
        setAttSingleBuilt([]);
        setAttAllBuilt([]);
        setLeaveTag('');
        return;
      }

      const res = await queryPeople(query);
      if (!res.rows.length) {
        setStatus('empty');
        setResult(res);
        return;
      }
      setResult(res);
      setStatus('success');
      setAttWorstOpen(false);
      setAttBestOpen(false);
    } catch (e) {
      setStatus('error');
      setError(e instanceof Error ? e.message : '查詢失敗');
    }
  }

  function clear() {
    const baseWh = (!isAdmin && useGas && user?.warehouseKey) ? user.warehouseKey : mockWarehouses[0];
    setQuery({ warehouse: baseWh, page: mockPages[0], name: '', birthdayOrPhone: '' });
    setSearch('');
    setStatus('idle');
    setResult(null);
    setError('');
    setGasHeaders([]);
    setManualFrozenLeft(0);
    setFreezeStart(0);
    setFreezeEnd(0);
    setAttEndDate('');
    setAttSingleDate('');
    setAttName('');
    setAttSingleBuilt([]);
    setAttAllBuilt([]);
    setLeaveTag('');
  }

  const effectiveFrozenLeft = useMemo(() => {
    const base = manualFrozenLeft == null ? gasFrozenLeft : manualFrozenLeft;
    if (!base) return 0;
    return Math.max(0, Math.min(base, gasHeaders.length ? gasHeaders.length + (isHoursPage ? 1 : 1) : base));
  }, [manualFrozenLeft, gasFrozenLeft, gasHeaders.length, isHoursPage]);

  const dateList = useMemo(() => {
    const out: Array<{ ci: number; iso: string; label: string }> = [];
    if (!Array.isArray(gasDateCols) || !gasDateCols.length) return out;
    gasDateCols.forEach((ci) => {
      const iso = String(gasHeadersISO?.[ci] || '').trim();
      const label = iso ? fmtMDFromISO(iso) : String(gasHeaders?.[ci] || '');
      out.push({ ci, iso, label });
    });
    out.sort((a, b) => {
      if (a.iso && b.iso && a.iso !== b.iso) return a.iso.localeCompare(b.iso);
      return a.label.localeCompare(b.label, 'zh-Hant', { numeric: true, sensitivity: 'base' });
    });
    return out;
  }, [gasDateCols, gasHeadersISO, gasHeaders]);

  const dateIndicesUpToEnd = useMemo(() => {
    if (!dateList.length) return gasDateCols.slice();
    if (!attEndDate) return dateList.map((d) => d.ci);
    const idx = dateList.findIndex((d) => (d.iso || `idx_${d.ci}`) === attEndDate);
    if (idx < 0) return dateList.map((d) => d.ci);
    return dateList.slice(0, idx + 1).map((d) => d.ci);
  }, [dateList, attEndDate, gasDateCols]);

  const dateIndexForSingle = useMemo(() => {
    if (!attSingleDate) return null;
    const found = dateList.find((d) => (d.iso || `idx_${d.ci}`) === attSingleDate);
    return found ? found.ci : null;
  }, [attSingleDate, dateList]);

  const leaveOptions = useMemo(() => {
    if (!useGas) return [] as string[];
    const set = new Set<string>();
    const effectiveLeaveName = isAdmin ? leaveName.trim() : (user?.name || '').trim();
    const baseRows = (filteredRows as any as GasRecordRow[]) || [];
    const rowsForOpt = effectiveLeaveName
      ? baseRows.filter((r) => getNameFromRow(r) === effectiveLeaveName)
      : baseRows;
    if (!rowsForOpt.length) return [];

    const isMatrix = Array.isArray(gasDateCols) && gasDateCols.length > 0;
    if (isMatrix) {
      gasDateCols.forEach((ci) => {
        const hk = gasHeaders[ci];
        rowsForOpt.forEach((r) => {
          tokenizeCell((r as any)[hk]).forEach((t) => set.add(t));
        });
      });
    } else {
      const leaveKey = findLeaveKey(gasHeaders);
      if (!leaveKey) return [];
      rowsForOpt.forEach((r) => tokenizeCell((r as any)[leaveKey]).forEach((t) => set.add(t)));
    }

    return Array.from(set).sort((a, b) => a.localeCompare(b, 'zh-Hant', { numeric: true, sensitivity: 'base' }));
  }, [useGas, filteredRows, gasDateCols, gasHeaders, isAdmin, leaveName, user?.name]);

  const leaveFilterMode = useMemo<'matrix' | 'record'>(() => {
    return Array.isArray(gasDateCols) && gasDateCols.length > 0 ? 'matrix' : 'record';
  }, [gasDateCols]);

  const leaveFilteredRows = useMemo(() => {
    if (!useGas) return filteredRows;

    const effectiveLeaveName = isAdmin ? leaveName.trim() : (user?.name || '').trim();
    let base = filteredRows as any as GasRecordRow[];
    if (effectiveLeaveName) base = base.filter((r) => getNameFromRow(r) === effectiveLeaveName);

    if (!leaveTag) return base as any;
    if (leaveFilterMode === 'matrix') return base as any;
    const leaveKey = findLeaveKey(gasHeaders);
    if (!leaveKey) return base as any;
    const re = new RegExp(`(^|[\\s\\n、，,;／/])${leaveTag.replace(/[.*+?^${}()|[\\]\\]/g, '\\$&')}($|[\\s\\n、，,;／/])`);
    return base.filter((r) => {
      const raw = String((r as any)[leaveKey] ?? '');
      return re.test(raw);
    }) as any;
  }, [useGas, filteredRows, leaveTag, leaveFilterMode, gasHeaders, isAdmin, leaveName, user?.name]);

  const leaveHiddenDateCols = useMemo(() => {
    if (!useGas) return new Set<number>();
    if (!leaveTag) return new Set<number>();
    if (leaveFilterMode !== 'matrix') return new Set<number>();
    const rowsForCheck = leaveFilteredRows as any as GasRecordRow[];
    const hide = new Set<number>();
    if (!rowsForCheck.length) return hide;
    gasDateCols.forEach((ci) => {
      const hk = gasHeaders[ci];
      let has = false;
      for (let i = 0; i < rowsForCheck.length; i++) {
        const toks = tokenizeCell((rowsForCheck[i] as any)[hk]);
        if (toks.includes(leaveTag)) {
          has = true;
          break;
        }
      }
      if (!has) hide.add(ci);
    });
    return hide;
  }, [useGas, leaveTag, leaveFilterMode, gasDateCols, gasHeaders, leaveFilteredRows]);

  const columnsShown = useMemo(() => {
    if (!useGas || !gasHeaders.length) return columns;
    if (!leaveHiddenDateCols.size) return columns;
    const baseHeaders = gasHeaders.filter((h) => String(h || '').trim());
    const kept = baseHeaders.filter((_, idx) => !leaveHiddenDateCols.has(idx));
    const cols: ColumnDef<GasRecordRow>[] = isHoursPage ? buildGasColumnsNoAttendance(kept) : buildGasColumns(kept);
    return cols as unknown as ColumnDef<DisplayRow>[];
  }, [useGas, gasHeaders, leaveHiddenDateCols, isHoursPage, columns]);

  const rowsShown = useMemo(() => {
    return leaveFilteredRows;
  }, [leaveFilteredRows]);

  function applyFreezeRange() {
    if (!gasHeaders.length) return;
    const max = Math.max(freezeStart, freezeEnd);
    setManualFrozenLeft(Math.min(gasHeaders.length, max + 1));
  }

  function resetFreezeRange() {
    setManualFrozenLeft(0);
    setFreezeStart(0);
    setFreezeEnd(0);
  }

  function restoreSheetFreeze() {
    setManualFrozenLeft(null);
    setFreezeStart(0);
    setFreezeEnd(0);
  }

  async function buildSingleAttStat() {
    if (!useGas) return;
    if (!query.page.includes('班表')) return;
    const all = rows as any as GasRecordRow[];
    if (!all.length || !gasDateCols.length || !dateList.length) return;

    const targetName = isAdmin ? attName.trim() : (user?.name || '').trim();
    if (isAdmin && !targetName) return;

    const nameKey = '姓名';
    const personRows = all.filter((r) => String((r as any)[nameKey] ?? '').trim() === targetName);
    if (!personRows.length) {
      setAttSingleBuilt([]);
      return;
    }

    const exclude = buildExcludeForAttRateSet();

    const deptKey = findDeptKey(gasHeaders);
    const shiftKey = findShiftKey(gasHeaders);
    const deptLabel = deptKey
      ? Array.from(
          new Set(
            all
              .filter((r) => String((r as any)[nameKey] ?? '').trim() === targetName)
              .map((r) => String((r as any)[deptKey] ?? '').trim())
              .filter(Boolean)
          )
        ).join('、') || '未填部門'
      : '未填部門';

    const shiftLabel = shiftKey
      ? Array.from(new Set(personRows.map((r) => String((r as any)[shiftKey] ?? '').trim()).filter(Boolean))).join('、')
      : '';

    const rangeLabel = !attEndDate
      ? '全部日期'
      : (() => {
          const d = dateList.find((x) => (x.iso || `idx_${x.ci}`) === attEndDate);
          return d ? `起始日至 ${d.label}` : '起始日至選擇日';
        })();

    function computeForCols(colIndices: number[]) {
      const excludeAbs = buildExcludeFromAbsSet();
      let shouldDays = 0;
      let actDays = 0;
      personRows.forEach((row) => {
        colIndices.forEach((ci) => {
          const hk = gasHeaders[ci];
          if (!hk || !String(hk).trim()) return;
          const cellValue = (row as any)[hk];
          if (!isShouldAttendCell(cellValue, exclude)) return;
          shouldDays += 1;
          if (isActualAttendCell(ci, row, cellValue, excludeAbs)) actDays += 1;
        });
      });
      const absent = shouldDays - actDays;
      const rate = shouldDays > 0 ? (actDays / shouldDays) * 100 : null;
      return { shouldDays, actDays, absent, rate };
    }

    const dataRows: Array<[string, string, string, string, number, number, number, string]> = [];
    const rangeStat = computeForCols(dateIndicesUpToEnd);
    dataRows.push([
      targetName,
      deptLabel,
      shiftLabel,
      rangeLabel,
      rangeStat.shouldDays,
      rangeStat.actDays,
      rangeStat.absent,
      rangeStat.rate != null ? `${rangeStat.rate.toFixed(1)}%` : '—',
    ]);

    if (dateIndexForSingle != null) {
      const d = dateList.find((x) => x.ci === dateIndexForSingle);
      const singleLabel = `單日 ${d ? d.label : '指定日期'}`;
      const singleStat = computeForCols([dateIndexForSingle]);
      dataRows.push([
        targetName,
        deptLabel,
        shiftLabel,
        singleLabel,
        singleStat.shouldDays,
        singleStat.actDays,
        singleStat.absent,
        singleStat.rate != null ? `${singleStat.rate.toFixed(1)}%` : '—',
      ]);
    }

    setAttSingleBuilt(dataRows);
    try {
      sessionStorage.setItem(
        'coupang_att_single_meta',
        JSON.stringify({ name: targetName, dept: deptLabel, shift: shiftLabel, sheet: query.page })
      );
    } catch {
      // ignore
    }
  }

  function clearSingleAttStat() {
    setAttEndDate('');
    setAttSingleDate('');
    setAttSingleBuilt([]);
  }

  async function buildAllAttStat() {
    if (!useGas) return;
    if (!isAdmin) return;
    if (!query.page.includes('班表')) return;
    const all = rows as any as GasRecordRow[];
    if (!all.length || !gasDateCols.length || !dateList.length) return;

    const exclude = buildExcludeForAttRateSet();
    const deptKey = findDeptKey(gasHeaders);
    const shiftKey = findShiftKey(gasHeaders);
    const nameKey = '姓名';

    const byName = new Map<string, GasRecordRow[]>();
    all.forEach((r) => {
      const nm = String((r as any)[nameKey] ?? '').trim();
      if (!nm) return;
      if (!byName.has(nm)) byName.set(nm, []);
      byName.get(nm)!.push(r);
    });

    function computeForCols(personRows: GasRecordRow[], colIndices: number[]) {
      const excludeAbs = buildExcludeFromAbsSet();
      let shouldDays = 0;
      let actDays = 0;
      personRows.forEach((row) => {
        colIndices.forEach((ci) => {
          const hk = gasHeaders[ci];
          if (!hk || !String(hk).trim()) return;
          const cellValue = (row as any)[hk];
          if (!isShouldAttendCell(cellValue, exclude)) return;
          shouldDays += 1;
          if (isActualAttendCell(ci, row, cellValue, excludeAbs)) actDays += 1;
        });
      });
      const rate = shouldDays > 0 ? (actDays / shouldDays) * 100 : null;
      return { shouldDays, actDays, rate };
    }

    const dataRows: Array<[string, string, string, number, number, string]> = [];
    byName.forEach((personRows, nm) => {
      const dept = deptKey
        ? Array.from(new Set(personRows.map((r) => String((r as any)[deptKey] ?? '').trim()).filter(Boolean))).join('、') || '未填部門'
        : '未填部門';
      const shift = shiftKey
        ? Array.from(new Set(personRows.map((r) => String((r as any)[shiftKey] ?? '').trim()).filter(Boolean))).join('、')
        : '';
      const st = computeForCols(personRows, dateIndicesUpToEnd);
      dataRows.push([dept, shift, nm, st.shouldDays, st.actDays, st.rate != null ? `${st.rate.toFixed(1)}%` : '—']);
    });

    dataRows.sort((a, b) => {
      const cmpDept = String(a[0] || '').localeCompare(String(b[0] || ''), 'zh-Hant', { numeric: true, sensitivity: 'base' });
      if (cmpDept !== 0) return cmpDept;
      return String(a[2] || '').localeCompare(String(b[2] || ''), 'zh-Hant', { numeric: true, sensitivity: 'base' });
    });

    setAttAllBuilt(dataRows);
  }

  function clearAllAttStat() {
    setAttAllBuilt([]);
  }

  async function openWarehouseSheet() {
    if (!isAdmin) return;
    if (!useGas) return;
    const sid = await gasGetWarehouseId(query.warehouse);
    window.open(`https://docs.google.com/spreadsheets/d/${sid}/edit`, '_blank', 'noopener,noreferrer');
  }

  useLayoutEffect(() => {
    if (status !== 'success') return;
    if (!tbodyRef.current) return;
    const rowsEl = Array.from(tbodyRef.current.querySelectorAll('tr'));
    gsap.fromTo(
      rowsEl,
      { opacity: 0, y: 8 },
      { opacity: 1, y: 0, duration: 0.35, stagger: 0.04, ease: 'power2.out' }
    );
  }, [status, rows.length]);

  return (
    <div className="page">
      <header className="topbar">
        <div className="topbarLeft">
          <div className="brandMark">CP</div>
          <div>
            <div className="brandTitle">企業後台查詢</div>
            <div className="brandSub">表格結果 + 出勤率卡片（示範）</div>
          </div>
        </div>
        <div className="topbarRight">
          {isAdmin && useGas ? (
            <button className="btnGhost" onClick={() => void openWarehouseSheet()}>
              開啟試算表
            </button>
          ) : null}
          <button
            className="btnGhost"
            onClick={() => {
              logout();
              navigate('/login', { replace: true });
            }}
          >
            登出
          </button>
        </div>
      </header>

      <main className="content">
        <section className="panel">
          <div className="panelTitle">搜尋</div>
          <div className="filters">
            <label className="filter">
              <span>倉別</span>
              <select
                value={query.warehouse}
                onChange={(e) => {
                  if (!isAdmin && useGas) return;
                  const w = e.target.value;
                  // ✅ 避免切倉時先用舊分頁觸發一次 doQuery（會造成切倉很慢）
                  setQuery((s) => ({ ...s, warehouse: w, page: '' }));
                  refreshPages(w);
                }}
                disabled={useGas && !isAdmin}
              >
                {mockWarehouses.map((w) => (
                  <option key={w} value={w}>{w}</option>
                ))}
              </select>
            </label>

            <label className="filter">
              <span>分頁</span>
              <select value={query.page} onChange={(e) => setQuery((s) => ({ ...s, page: e.target.value }))}>
                {availablePages.map((p) => (
                  <option key={p} value={p}>{p}</option>
                ))}
              </select>
            </label>

            <label className="filter" style={{ gridColumn: 'span 6' }}>
              <span>搜尋</span>
              <input
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                placeholder={isAdmin ? '輸入姓名/部門/任意關鍵字（全員資料已載入）' : '輸入關鍵字'}
              />
            </label>

            <div className="filterActions">
              <button className="btnPrimary" onClick={doQuery} disabled={status === 'loading'}>
                {status === 'loading' ? '查詢中…' : '查詢'}
              </button>
              <button className="btnSecondary" onClick={() => setSearch('')} disabled={status === 'loading'}>
                清除搜尋
              </button>
            </div>
          </div>
        </section>

        {status === 'loading' ? (
          <section className="panel">
            <div className="panelTitle">查詢結果</div>
            <SkeletonTable />
          </section>
        ) : null}

        {status === 'error' ? (
          <section className="panel">
            <div className="panelTitle">查詢結果</div>
            <EmptyState title="查詢失敗" description={error} actionLabel="重試" onAction={doQuery} />
          </section>
        ) : null}

        {status === 'empty' ? (
          <section className="panel">
            <div className="panelTitle">查詢結果</div>
            <EmptyState title="沒有資料" description="請調整查詢條件後再試一次。" actionLabel="返回" onAction={clear} />
          </section>
        ) : null}

        {status === 'success' && stats ? (
          <section className="grid2 grid2Summary">
            <div className="panel">
              <div className="panelTitle">統計摘要</div>
              <div className="kpiRow">
                <div className="kpi">
                  <div className="kpiLabel">筆數</div>
                  <div className="kpiValue">{stats.total}</div>
                </div>
                <DeptKpi rows={rows} headers={gasHeaders} />
                <div className="kpi">
                  <div className="kpiLabel">分頁人數</div>
                  <div className="kpiValue">{pageStats.total}</div>
                </div>
                <div className="kpi">
                  <div className="kpiLabel">離職人數</div>
                  <div className="kpiValue">{pageStats.leave}</div>
                </div>
                {!isHoursPage && isAttPage && attendanceAllAgg.n ? (
                  <>
                    <div className="kpi">
                      <div className="kpiLabel">全員出勤率</div>
                      <div className="kpiValue">{Math.round(attendanceAllAgg.rate * 100)}%</div>
                    </div>
                    <div className="kpi">
                      <div className="kpiLabel">全員出勤（合計）</div>
                      <div className="kpiValue">{attendanceAllAgg.attended}/{attendanceAllAgg.expected}</div>
                    </div>
                  </>
                ) : null}
              </div>
            </div>

            {!isHoursPage && isAttPage ? (
              <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
                <div className="panel">
                  <div
                    className="panelTitle"
                    style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 10 }}
                  >
                    <span>出勤最差前五位</span>
                    {hasAttendanceStats ? (
                      <div style={{ display: 'flex', gap: 10 }}>
                        <button className="btnGhost" onClick={() => setAttWorstOpen((v) => !v)}>
                          {attWorstOpen ? '收合' : '點開'}
                        </button>
                        <button className="btnGhost" onClick={openAttendanceFullList}>
                          查看全部
                        </button>
                      </div>
                    ) : null}
                  </div>

                  {hasAttendanceStats ? (
                    attWorstOpen ? (
                      <AttendanceCards items={attWorst5} onItemClick={openAttendanceFullList} />
                    ) : (
                      <div style={{ color: 'var(--muted)', fontSize: 13 }}>點開即可查看最差前五位</div>
                    )
                  ) : (
                    <div style={{ color: 'var(--muted)', fontSize: 13 }}>本分頁不統計</div>
                  )}
                </div>

                <div className="panel">
                  <div
                    className="panelTitle"
                    style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 10 }}
                  >
                    <span>出勤最好前五位</span>
                    {hasAttendanceStats ? (
                      <div style={{ display: 'flex', gap: 10 }}>
                        <button className="btnGhost" onClick={() => setAttBestOpen((v) => !v)}>
                          {attBestOpen ? '收合' : '點開'}
                        </button>
                        <button className="btnGhost" onClick={openAttendanceFullList}>
                          查看全部
                        </button>
                      </div>
                    ) : null}
                  </div>

                  {hasAttendanceStats ? (
                    attBestOpen ? (
                      <AttendanceCards items={attBest5} onItemClick={openAttendanceFullList} />
                    ) : (
                      <div style={{ color: 'var(--muted)', fontSize: 13 }}>點開即可查看最好前五位</div>
                    )
                  ) : (
                    <div style={{ color: 'var(--muted)', fontSize: 13 }}>本分頁不統計</div>
                  )}
                </div>
              </div>
            ) : null}
          </section>
        ) : null}

        {status === 'success' ? (
          <section className="panel">
            <div className="panelTitle" style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 10 }}>
              <span>結果表格</span>
              <div style={{ display: 'flex', gap: 10 }}>
                <button
                  className="btnGhost"
                  onClick={() => {
                    const headers = (columnsShown as any[]).map((c) => String(c.header ?? ''));
                    const rows2 = (rowsShown as any[]).map((r) => (columnsShown as any[]).map((c) => {
                      const k = c.key;
                      if (k === '_attendanceRate') {
                        const a = (r as any)._attendance;
                        if (!a) return '';
                        return `${Math.round(a.rate * 100)}% (${a.attended}/${a.expected})`;
                      }
                      return String((r as any)[k] ?? '');
                    }));
                    exportExcelHtml('分頁資料', headers, rows2);
                  }}
                  disabled={!rowsShown || !(rowsShown as any[]).length}
                >
                  Excel
                </button>
                <button
                  className="btnGhost"
                  onClick={() => {
                    const headers = (columnsShown as any[]).map((c) => String(c.header ?? ''));
                    const rows2 = (rowsShown as any[]).map((r) => (columnsShown as any[]).map((c) => {
                      const k = c.key;
                      if (k === '_attendanceRate') {
                        const a = (r as any)._attendance;
                        if (!a) return '';
                        return `${Math.round(a.rate * 100)}% (${a.attended}/${a.expected})`;
                      }
                      return String((r as any)[k] ?? '');
                    }));
                    exportCsv('分頁資料', headers, rows2);
                  }}
                  disabled={!rowsShown || !(rowsShown as any[]).length}
                >
                  CSV
                </button>
                <button
                  className="btnGhost"
                  onClick={() => {
                    if (!mainTableWrapRef.current) return;
                    exportElementPng(mainTableWrapRef.current, '分頁資料');
                  }}
                  disabled={!rowsShown || !(rowsShown as any[]).length}
                >
                  PNG
                </button>
              </div>
            </div>
            <div ref={mainTableWrapRef}>
              <ResultTable
                columns={columnsShown as any}
                rows={rowsShown as any}
                tbodyRef={tbodyRef}
                frozenLeft={useGas ? effectiveFrozenLeft : 0}
                cellStyle={({ row, colIndex }: { row: any; colIndex: number }) => {
                  if (!useGas) return undefined;
                  const r = row as any;
                  const bgArr = r?._bg as string[] | undefined;
                  const fcArr = r?._fc as string[] | undefined;

                  const wh = query.warehouse.trim().toLowerCase();
                  const isTA01Record = (wh === 'ta01' || wh === 'tao1') && query.page.includes('出勤記錄');
                  const allowColor = query.page.includes('班表') || isTA01Record;
                  if (!allowColor) return undefined;

                  const bg = bgArr?.[colIndex];
                  const fc = fcArr?.[colIndex];
                  if (!bg && !fc) return undefined;
                  return {
                    background: bg || undefined,
                    color: fc || undefined,
                  };
                }}
              />
            </div>
          </section>
        ) : null}

        {status === 'success' && useGas ? (
          <section className="panel">
            <div className="panelTitle" style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 10 }}>
              <span>凍結欄位</span>
              <button className="btnGhost" onClick={() => setOpenFreezePanel((v) => !v)}>
                {openFreezePanel ? '收合' : '點開'}
              </button>
            </div>
            {!openFreezePanel ? (
              <div style={{ color: 'var(--muted)', fontSize: 13 }}>點開即可設定凍結欄位</div>
            ) : (
              <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap' }}>
                <label className="filter" style={{ minWidth: 240 }}>
                  <span>起始欄</span>
                  <select value={freezeStart} onChange={(e) => setFreezeStart(Number(e.target.value))}>
                    {gasHeaders.map((t, i) => (
                      <option key={i} value={i}>{i + 1}. {t || ''}</option>
                    ))}
                  </select>
                </label>
                <label className="filter" style={{ minWidth: 240 }}>
                  <span>訖欄</span>
                  <select value={freezeEnd} onChange={(e) => setFreezeEnd(Number(e.target.value))}>
                    {gasHeaders.map((t, i) => (
                      <option key={i} value={i}>{i + 1}. {t || ''}</option>
                    ))}
                  </select>
                </label>
                <button className="btnPrimary" onClick={applyFreezeRange}>套用凍結</button>
                <button className="btnSecondary" onClick={resetFreezeRange}>解除凍結</button>
                <button className="btnGhost" onClick={restoreSheetFreeze}>依試算表</button>
                <div style={{ color: 'var(--muted)', fontSize: 13, alignSelf: 'center' }}>目前凍結：{effectiveFrozenLeft} 欄</div>
              </div>
            )}
          </section>
        ) : null}

        {status === 'success' && useGas ? (
          <LeaveStatPanel
            headers={gasHeaders}
            headersISO={gasHeadersISO}
            dateCols={gasDateCols}
            rowsAll={rows as any}
            rowsSingle={filteredRows as any}
            isAdmin={isAdmin}
            userName={user?.name || ''}
            sheetName={query.page}
          />
        ) : null}

        {status === 'success' && useGas && query.page.includes('班表') ? (
          <section className="panel">
            <div className="panelTitle" style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 10 }}>
              <span>出勤率統計</span>
              <button className="btnGhost" onClick={() => setOpenAttStatPanel((v) => !v)}>
                {openAttStatPanel ? '收合' : '點開'}
              </button>
            </div>
            {!openAttStatPanel ? (
              <div style={{ color: 'var(--muted)', fontSize: 13 }}>點開即可查看/匯出出勤率統計</div>
            ) : (
              <>
                <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', marginBottom: 10 }}>
                  <button
                    className="btnGhost"
                    onClick={() => {
                      let meta: any = null;
                      try { meta = JSON.parse(sessionStorage.getItem('coupang_att_single_meta') || 'null'); } catch {}
                      const base = [meta?.dept, meta?.shift, meta?.name, meta?.sheet, '單人出勤率統計'].filter(Boolean).join('_') || '單人出勤率統計';
                      exportExcelHtml(base, ['姓名', '部門', '班別', '統計範圍', '應到天數', '實到天數', '未到天數', '出勤率'], attSingleBuilt);
                    }}
                    disabled={!attSingleBuilt.length}
                  >
                    Excel
                  </button>
              <button
                className="btnGhost"
                onClick={() => {
                  let meta: any = null;
                  try { meta = JSON.parse(sessionStorage.getItem('coupang_att_single_meta') || 'null'); } catch {}
                  const base = [meta?.dept, meta?.shift, meta?.name, meta?.sheet, '單人出勤率統計'].filter(Boolean).join('_') || '單人出勤率統計';
                  exportCsv(base, ['姓名', '部門', '班別', '統計範圍', '應到天數', '實到天數', '未到天數', '出勤率'], attSingleBuilt);
                }}
                disabled={!attSingleBuilt.length}
              >
                CSV
              </button>
              <button
                className="btnGhost"
                onClick={() => {
                  if (!attSingleWrapRef.current) return;
                  let meta: any = null;
                  try { meta = JSON.parse(sessionStorage.getItem('coupang_att_single_meta') || 'null'); } catch {}
                  const base = [meta?.dept, meta?.shift, meta?.name, meta?.sheet, '單人出勤率統計'].filter(Boolean).join('_') || '單人出勤率統計';
                  exportElementPng(attSingleWrapRef.current, base);
                }}
                disabled={!attSingleBuilt.length}
              >
                PNG
              </button>
                </div>
                <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap', marginBottom: 10 }}>
              {isAdmin ? (
                <label className="filter" style={{ minWidth: 220 }}>
                  <span>姓名</span>
                  <input value={attName} onChange={(e) => setAttName(e.target.value)} placeholder="輸入要統計的姓名" />
                </label>
              ) : null}

              <label className="filter" style={{ minWidth: 220 }}>
                <span>統計至</span>
                <select value={attEndDate} onChange={(e) => setAttEndDate(e.target.value)}>
                  <option value="">全部日期</option>
                  {dateList.map((d) => (
                    <option key={d.iso || `idx_${d.ci}`} value={d.iso || `idx_${d.ci}`}>{d.label}</option>
                  ))}
                </select>
              </label>

              <label className="filter" style={{ minWidth: 220 }}>
                <span>單日</span>
                <select value={attSingleDate} onChange={(e) => setAttSingleDate(e.target.value)}>
                  <option value="">不指定</option>
                  {dateList.map((d) => (
                    <option key={d.iso || `idx_${d.ci}`} value={d.iso || `idx_${d.ci}`}>{d.label}</option>
                  ))}
                </select>
              </label>

              <div className="filterActions" style={{ marginLeft: 'auto' }}>
                <button className="btnPrimary" onClick={buildSingleAttStat}>
                  單人出勤率統計
                </button>
                <button className="btnSecondary" onClick={clearSingleAttStat}>
                  清除
                </button>
              </div>
            </div>

            {attSingleBuilt.length ? (
              <div className="tableWrap" ref={attSingleWrapRef}>
                <table className="table">
                  <thead>
                    <tr>
                      <th>姓名</th>
                      <th>部門</th>
                      <th>班別</th>
                      <th>統計範圍</th>
                      <th>應到天數</th>
                      <th>實到天數</th>
                      <th>未到天數</th>
                      <th>出勤率</th>
                    </tr>
                  </thead>
                  <tbody>
                    {attSingleBuilt.map((r, i) => (
                      <tr key={i}>
                        <td>{r[0]}</td>
                        <td>{r[1]}</td>
                        <td>{r[2]}</td>
                        <td>{r[3]}</td>
                        <td>{r[4]}</td>
                        <td>{r[5]}</td>
                        <td>{r[6]}</td>
                        <td>{r[7]}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : (
              <div style={{ color: 'var(--muted)', fontSize: 13 }}>—</div>
            )}
              </>
            )}
          </section>
        ) : null}

        {status === 'success' && useGas && isAdmin && query.page.includes('班表') ? (
          <section className="panel">
            <div className="panelTitle" style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 10 }}>
              <span>全員出勤率統計</span>
              <div style={{ display: 'flex', gap: 10, alignItems: 'center' }}>
                <button className="btnGhost" onClick={() => setOpenAttAllPanel((v) => !v)}>
                  {openAttAllPanel ? '收合' : '點開'}
                </button>
                <button className="btnPrimary" onClick={buildAllAttStat} disabled={!openAttAllPanel}>建立</button>
                <button className="btnSecondary" onClick={clearAllAttStat} disabled={!openAttAllPanel}>清除</button>
              </div>
            </div>

            {!openAttAllPanel ? (
              <div style={{ color: 'var(--muted)', fontSize: 13 }}>點開即可建立/匯出全員出勤率統計</div>
            ) : (
              <>
                <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', marginBottom: 10 }}>
                  <button
                    className="btnGhost"
                    onClick={() => exportExcelHtml('全員出勤率統計', ['部門', '班別', '姓名', '應到天數', '實到天數', '出勤率'], attAllBuilt)}
                    disabled={!attAllBuilt.length}
                  >
                    Excel
                  </button>
                  <button
                    className="btnGhost"
                    onClick={() => exportCsv('全員出勤率統計', ['部門', '班別', '姓名', '應到天數', '實到天數', '出勤率'], attAllBuilt)}
                    disabled={!attAllBuilt.length}
                  >
                    CSV
                  </button>
                </div>

                {attAllBuilt.length ? (
                  <div className="tableWrap" ref={attAllWrapRef}>
                    <table className="table">
                      <thead>
                        <tr>
                          <th>部門</th>
                          <th>班別</th>
                          <th>姓名</th>
                          <th>應到天數</th>
                          <th>實到天數</th>
                          <th>出勤率</th>
                        </tr>
                      </thead>
                      <tbody>
                        {attAllBuilt.map((r, i) => (
                          <tr key={i}>
                            <td>{r[0]}</td>
                            <td>{r[1]}</td>
                            <td>{r[2]}</td>
                            <td>{r[3]}</td>
                            <td>{r[4]}</td>
                            <td>{r[5]}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                ) : (
                  <div style={{ color: 'var(--muted)', fontSize: 13 }}>—</div>
                )}
              </>
            )}
          </section>
        ) : null}

        {status === 'success' && useGas && (query.page.includes('班表') || query.page.includes('出勤記錄')) ? (
          <section className="panel">
            <div className="panelTitle" style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 10 }}>
              <span>單人假別欄位篩選</span>
              <button className="btnGhost" onClick={() => setOpenLeaveFilterPanel((v) => !v)}>
                {openLeaveFilterPanel ? '收合' : '點開'}
              </button>
            </div>

            {!openLeaveFilterPanel ? (
              <div style={{ color: 'var(--muted)', fontSize: 13 }}>點開即可依假別過濾/隱藏日期欄</div>
            ) : (
              <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap' }}>
                {isAdmin ? (
                  <label className="filter" style={{ minWidth: 220 }}>
                    <span>姓名</span>
                    <input
                      value={leaveName}
                      onChange={(e) => setLeaveName(e.target.value)}
                      placeholder="輸入姓名（留空=依搜尋結果）"
                    />
                  </label>
                ) : null}
                <label className="filter" style={{ minWidth: 240 }}>
                  <span>假別</span>
                  <select value={leaveTag} onChange={(e) => setLeaveTag(e.target.value)}>
                    <option value="">全部日期</option>
                    {leaveOptions.map((t) => (
                      <option key={t} value={t}>{t}</option>
                    ))}
                  </select>
                </label>
                <button className="btnSecondary" onClick={() => setLeaveTag('')}>清除</button>
                <div style={{ color: 'var(--muted)', fontSize: 13, alignSelf: 'center' }}>
                  {leaveFilterMode === 'matrix' ? '班表：依假別顯示/隱藏日期欄' : '出勤記錄：依假別列式過濾'}
                </div>
              </div>
            )}
          </section>
        ) : null}

      </main>
    </div>
  );
}
