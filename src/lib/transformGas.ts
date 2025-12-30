import type { AttendanceSummary } from './mockApi';
import type { GasPayload, GasRow } from './gasApi';

export type GasRecordRow = {
  id: string;
  _attendance?: AttendanceSummary;
  _attendanceRate?: number;
  _bg?: string[];
  _fc?: string[];
  _att?: number[];
  [key: string]: unknown;
};

type GasTransformOptions = {
  disableAttendance?: boolean;
  sheetName?: string;
  excludeForAttRate?: string[];
};

function tokenizeCell(raw: unknown): string[] {
  const s = (raw == null ? '' : String(raw)).trim();
  if (!s) return [];
  return s
    .split(/[、，,;／\/\n]+/)
    .map((x) => x.trim())
    .filter(Boolean);
}

function buildExcludeSet(list: string[] | undefined): Set<string> {
  const base = new Set<string>([
    '休',
    '休假',
    '公休',
    '特休',
    '病假',
    '事假',
    '喪假',
    '婚假',
    '產假',
    '育嬰',
    'OFF',
    'off',
    '離',
  ]);
  (list || []).forEach((t) => {
    const v = String(t || '').trim();
    if (v) base.add(v);
  });
  return base;
}

function isShouldAttendCell(raw: unknown, exclude: Set<string>): boolean {
  const s = (raw == null ? '' : String(raw)).trim();
  if (!s) return true;
  const toks = tokenizeCell(s);
  return !toks.some((t) => exclude.has(t));
}

function findNameIndex(headers: string[]): number {
  const exact = ['姓名', 'Name', 'name', '員工姓名', '中文姓名'];
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] ?? '').trim();
    if (!h) continue;
    if (exact.includes(h)) return i;
  }
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] ?? '').trim();
    if (!h) continue;
    if (h.includes('姓名') || /^name$/i.test(h)) return i;
  }
  return -1;
}

function statusFromRate(rate: number): AttendanceSummary['status'] {
  if (rate >= 0.9) return 'normal';
  if (rate >= 0.75) return 'low';
  return 'abnormal';
}

function calcAttendance(
  headers: string[] | undefined,
  dateCols: number[] | undefined,
  row: GasRow,
  options?: GasTransformOptions
): AttendanceSummary | undefined {
  const sheetName = String(options?.sheetName || '').trim();
  if (!sheetName.includes('班表')) return undefined;
  if (!dateCols?.length) return undefined;
  if (!headers?.length) return undefined;

  const att = row.att;
  if (!att || !att.length) return undefined;

  const exclude = buildExcludeSet(options?.excludeForAttRate);

  let shouldDays = 0;
  let actDays = 0;
  for (const ci of dateCols) {
    const hk = headers[ci];
    if (!hk || !String(hk).trim()) continue;
    if (!row.v || ci < 0 || ci >= row.v.length) continue;

    if (isShouldAttendCell(row.v?.[ci], exclude)) {
      shouldDays += 1;
      if (att[ci]) actDays += 1;
    }
  }

  const expected = shouldDays;
  const attended = actDays;
  const rate = expected ? attended / expected : 0;
  return { rate, attended, expected, status: statusFromRate(rate) };
}

export function gasPayloadToRows(
  payload: GasPayload,
  options?: GasTransformOptions
): { headers: string[]; rows: GasRecordRow[] } {
  const headers = (payload.headers ?? []).map((h) => String(h ?? ''));
  const nameIdx = findNameIndex(headers);

  function cleanForBlankCheck(v: unknown): string {
    return String(v ?? '')
      .replace(/[\u00A0\u200B-\u200F\u202A-\u202E\u2060\u2066-\u2069\uFEFF]/g, '')
      .replace(/\p{Cf}/gu, '')
      .replace(/\s+/g, '')
      .trim();
  }

  function isBlankRow(r: GasRow): boolean {
    const v = r?.v || [];
    for (let i = 0; i < headers.length; i++) {
      const cell = cleanForBlankCheck(v[i]);
      if (cell) return false;
    }
    return true;
  }

  const rows: GasRecordRow[] = (payload.rows ?? [])
    .filter((r) => !isBlankRow(r))
    .map((r, idx) => {
      const obj: GasRecordRow = { id: `gas_${idx}` };
      for (let i = 0; i < headers.length; i++) {
        const key = headers[i] || `col_${i + 1}`;
        obj[key] = r.v?.[i] ?? '';
      }

      if (Array.isArray((r as any).bg)) obj._bg = (r as any).bg as string[];
      if (Array.isArray((r as any).fc)) obj._fc = (r as any).fc as string[];
      if (Array.isArray((r as any).att)) obj._att = (r as any).att as number[];

      if (!options?.disableAttendance) {
        const att = calcAttendance(headers, payload.dateCols, r, options);
        if (att) {
          obj._attendance = att;
          obj._attendanceRate = att.rate;
        }
      }

      if (nameIdx >= 0) {
        obj['姓名'] = r.v?.[nameIdx] ?? obj['姓名'];
      }

      return obj;
    });

  return { headers, rows };
}
