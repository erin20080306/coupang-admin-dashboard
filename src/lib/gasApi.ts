export type GasRow = {
  v: string[];
  bg?: string[];
  fc?: string[];
  att?: number[];
};

export type GasPayload = {
  headers: string[];
  headersISO?: string[];
  headersTop?: string[];
  rows: GasRow[];
  dateCols?: number[];
  frozenLeft?: number;
  engine?: string;
  counts?: { rows: number; cols: number };
  error?: string;
};

export type GasLoginResult = {
  ok: boolean;
  msg?: string;
  isAdmin?: boolean;
  name?: string;
  warehouseKey?: string;
  warehouse?: string;
  whKey?: string;
};

export type GasWarehouseIdResult =
  | { ok: true; warehouse: string; spreadsheetId: string }
  | { ok: false; error: string };

type CacheEntry<T> = { ts: number; value: T };

const SHEETS_CACHE = new Map<string, CacheEntry<string[]>>();
const SHEETS_CACHE_MS = 5 * 60_000;

const QUERY_CACHE = new Map<string, CacheEntry<GasPayload>>();
const QUERY_CACHE_MS = 2 * 60_000;

const WAREHOUSE_ID_CACHE = new Map<string, CacheEntry<string>>();
const WAREHOUSE_ID_CACHE_MS = 10 * 60_000;

const FIND_WAREHOUSE_BY_NAME_CACHE = new Map<string, CacheEntry<string>>();
const FIND_WAREHOUSE_BY_NAME_CACHE_MS = 10 * 60_000;

const SHEETS_INFLIGHT = new Map<string, Promise<string[]>>();
const QUERY_INFLIGHT = new Map<string, Promise<GasPayload>>();
const WAREHOUSE_ID_INFLIGHT = new Map<string, Promise<string>>();
const VERIFY_LOGIN_INFLIGHT = new Map<string, Promise<GasLoginResult>>();
const FIND_WAREHOUSE_BY_NAME_INFLIGHT = new Map<string, Promise<string>>();

function getBaseUrl(): string | null {
  const v = (import.meta as any).env?.VITE_GAS_URL as string | undefined;
  const s = (v || '').replace(/\s+/g, '').trim();
  return s ? s : null;
}

async function fetchJsonNoStore<T>(url: string): Promise<T> {
  try {
    const res = await fetch(url, { cache: 'no-store' });
    if (!res.ok) throw new Error(`API 錯誤 ${res.status}`);
    return (await res.json()) as T;
  } catch (e) {
    const hint =
      '無法連線到 GAS。請確認：1) VITE_GAS_URL 是 Web App 的 /exec 連結 2) GAS 部署權限為「任何人」3) 修改 .env 後已重啟 npm run dev';
    const msg = e instanceof Error ? e.message : String(e);
    throw new Error(`${hint}\nURL: ${url}\n原因: ${msg}`);
  }
}

export async function gasGetWarehouseId(warehouse: string): Promise<string> {
  const base = getBaseUrl();
  if (!base) throw new Error('尚未設定 VITE_GAS_URL');
  const ck = `${base}|${warehouse}`;
  const hit = WAREHOUSE_ID_CACHE.get(ck);
  if (hit && Date.now() - hit.ts < WAREHOUSE_ID_CACHE_MS) return hit.value;

  const inflight = WAREHOUSE_ID_INFLIGHT.get(ck);
  if (inflight) return inflight;

  const p = (async () => {
    const url = toUrl(base, { mode: 'getWarehouseId', wh: warehouse, t: String(Date.now()) });
    const json = await fetchJsonNoStore<GasWarehouseIdResult>(url);
    if (!json || json.ok === false) throw new Error((json as any)?.error || '取得試算表 ID 失敗');
    WAREHOUSE_ID_CACHE.set(ck, { ts: Date.now(), value: json.spreadsheetId });
    return json.spreadsheetId;
  })();

  WAREHOUSE_ID_INFLIGHT.set(ck, p);
  try {
    return await p;
  } finally {
    WAREHOUSE_ID_INFLIGHT.delete(ck);
  }
}

export async function gasVerifyLogin(name: string, birthday: string): Promise<GasLoginResult> {
  const base = getBaseUrl();
  if (!base) throw new Error('尚未設定 VITE_GAS_URL');
  const ck = `${base}|${String(name || '').trim()}|${String(birthday || '').trim()}`;
  const inflight = VERIFY_LOGIN_INFLIGHT.get(ck);
  if (inflight) return inflight;

  const p = (async () => {
    const url = toUrl(base, { mode: 'verifyLogin', name, birthday, t: String(Date.now()) });
    return await fetchJsonNoStore<GasLoginResult>(url);
  })();

  VERIFY_LOGIN_INFLIGHT.set(ck, p);
  try {
    return await p;
  } finally {
    VERIFY_LOGIN_INFLIGHT.delete(ck);
  }
}

export async function gasFindWarehouseByName(name: string): Promise<string> {
  const base = getBaseUrl();
  if (!base) throw new Error('尚未設定 VITE_GAS_URL');
  const nm = String(name || '').trim();
  if (!nm) throw new Error('請輸入姓名');

  const ck = `${base}|${nm}`;
  const hit = FIND_WAREHOUSE_BY_NAME_CACHE.get(ck);
  if (hit && Date.now() - hit.ts < FIND_WAREHOUSE_BY_NAME_CACHE_MS) return hit.value;

  const inflight = FIND_WAREHOUSE_BY_NAME_INFLIGHT.get(ck);
  if (inflight) return inflight;

  const p = (async () => {
    const url = toUrl(base, { mode: 'findWarehouseByName', name: nm, t: String(Date.now()) });
    const json = await fetchJsonNoStore<any>(url);
    if (!json) throw new Error('查詢失敗');
    if (json.ok === false) throw new Error(String(json.error || json.msg || '查詢失敗'));

    const wh = String(json.warehouseKey ?? json.warehouse ?? json.whKey ?? json.key ?? '').trim().toUpperCase();
    if (!wh) throw new Error(String(json.error || json.msg || '查詢失敗'));

    FIND_WAREHOUSE_BY_NAME_CACHE.set(ck, { ts: Date.now(), value: wh });
    return wh;
  })();

  FIND_WAREHOUSE_BY_NAME_INFLIGHT.set(ck, p);
  try {
    return await p;
  } finally {
    FIND_WAREHOUSE_BY_NAME_INFLIGHT.delete(ck);
  }
}

function toUrl(base: string, params: Record<string, string | undefined>): string {
  const u = new URL(base);
  Object.entries(params).forEach(([k, v]) => {
    if (v == null) return;
    const vv = String(v).trim();
    if (!vv) return;
    u.searchParams.set(k, vv);
  });
  return u.toString();
}

export async function gasGetSheets(warehouse: string): Promise<string[]> {
  const base = getBaseUrl();
  if (!base) throw new Error('尚未設定 VITE_GAS_URL');
  const ck = `${base}|${warehouse}`;
  const hit = SHEETS_CACHE.get(ck);
  if (hit && Date.now() - hit.ts < SHEETS_CACHE_MS) return hit.value;

  const inflight = SHEETS_INFLIGHT.get(ck);
  if (inflight) return inflight;

  const p = (async () => {
    const url = toUrl(base, { mode: 'getSheets', wh: warehouse, t: String(Date.now()) });
    const json = await fetchJsonNoStore<{ sheetNames?: string[]; error?: string }>(url);
    if (json.error) throw new Error(json.error);
    const out = Array.isArray(json.sheetNames) ? json.sheetNames : [];
    SHEETS_CACHE.set(ck, { ts: Date.now(), value: out });
    return out;
  })();

  SHEETS_INFLIGHT.set(ck, p);
  try {
    return await p;
  } finally {
    SHEETS_INFLIGHT.delete(ck);
  }
}

export async function gasQuerySheet(
  warehouse: string,
  sheet: string,
  name?: string
): Promise<GasPayload> {
  const base = getBaseUrl();
  if (!base) throw new Error('尚未設定 VITE_GAS_URL');
  const ck = `${base}|${warehouse}|${sheet}|${(name || '').trim()}`;
  const hit = QUERY_CACHE.get(ck);
  if (hit && Date.now() - hit.ts < QUERY_CACHE_MS) return hit.value;

  const inflight = QUERY_INFLIGHT.get(ck);
  if (inflight) return inflight;

  const p = (async () => {
    const url = toUrl(base, {
      mode: 'api',
      wh: warehouse,
      sheet,
      name: (name || '').trim(),
      t: String(Date.now()),
    });
    const json = await fetchJsonNoStore<GasPayload>(url);
    if ((json as any).error) throw new Error(String((json as any).error));
    QUERY_CACHE.set(ck, { ts: Date.now(), value: json });
    return json;
  })();

  QUERY_INFLIGHT.set(ck, p);
  try {
    return await p;
  } finally {
    QUERY_INFLIGHT.delete(ck);
  }
}

export function gasIsConfigured(): boolean {
  return Boolean(getBaseUrl());
}
