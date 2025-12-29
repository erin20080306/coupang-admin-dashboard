export type AttendanceStatus = 'normal' | 'low' | 'abnormal';

export type AttendanceSummary = {
  rate: number;
  attended: number;
  expected: number;
  status: AttendanceStatus;
};

export type PersonRow = {
  id: string;
  warehouse: string;
  page: string;
  name: string;
  birthdayOrPhone: string;
  attendance: AttendanceSummary;
  late: number;
  absent: number;
};

export type QueryParams = {
  warehouse: string;
  page: string;
  name: string;
  birthdayOrPhone: string;
};

export type QueryResult = {
  rows: PersonRow[];
  stats: { total: number; attended: number; late: number; absent: number };
};

const WAREHOUSES = ['TAO1', 'TAO3', 'TAO4', 'TAO5', 'TAO6', 'TAO7', 'TAO10'];
const PAGES = ['班表', '出勤記錄', '出勤時數'];

function mulberry32(seed: number) {
  return function () {
    let t = (seed += 0x6d2b79f5);
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function pick<T>(rng: () => number, arr: T[]): T {
  return arr[Math.floor(rng() * arr.length)];
}

function statusFromRate(rate: number): AttendanceStatus {
  if (rate >= 0.9) return 'normal';
  if (rate >= 0.75) return 'low';
  return 'abnormal';
}

function makeRow(rng: () => number, i: number, p: QueryParams): PersonRow {
  const expected = 26;
  const attended = Math.max(0, Math.min(expected, Math.round(rng() * expected)));
  const rate = attended / expected;
  const late = Math.round(rng() * 4);
  const absent = Math.max(0, expected - attended);

  return {
    id: `row_${i}`,
    warehouse: p.warehouse || pick(rng, WAREHOUSES),
    page: p.page || pick(rng, PAGES),
    name: p.name || `員工${String(i + 1).padStart(3, '0')}`,
    birthdayOrPhone: p.birthdayOrPhone || `09${Math.floor(rng() * 100000000).toString().padStart(8, '0')}`,
    attendance: {
      rate,
      attended,
      expected,
      status: statusFromRate(rate),
    },
    late,
    absent,
  };
}

export async function queryPeople(p: QueryParams): Promise<QueryResult> {
  await new Promise((r) => setTimeout(r, 900));

  const seed = (p.warehouse + '|' + p.page + '|' + p.name + '|' + p.birthdayOrPhone)
    .split('')
    .reduce((acc, c) => acc + c.charCodeAt(0), 0);
  const rng = mulberry32(seed || 1);

  if (p.name.trim().toLowerCase() === 'fail') {
    throw new Error('查詢失敗（示範）');
  }

  const count = p.name.trim() ? 1 : 18;
  const rows = Array.from({ length: count }, (_, i) => makeRow(rng, i, p));

  const stats = rows.reduce(
    (acc, r) => {
      acc.total += 1;
      acc.attended += r.attendance.attended;
      acc.late += r.late;
      acc.absent += r.absent;
      return acc;
    },
    { total: 0, attended: 0, late: 0, absent: 0 }
  );

  return { rows, stats };
}

export const mockWarehouses = WAREHOUSES;
export const mockPages = PAGES;
