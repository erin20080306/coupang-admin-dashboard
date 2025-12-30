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
  const url = toUrl(base, { mode: 'getWarehouseId', wh: warehouse, t: String(Date.now()) });
  const json = await fetchJsonNoStore<GasWarehouseIdResult>(url);
  if (!json || json.ok === false) throw new Error((json as any)?.error || '取得試算表 ID 失敗');
  return json.spreadsheetId;
}

export async function gasVerifyLogin(name: string, birthday: string): Promise<GasLoginResult> {
  const base = getBaseUrl();
  if (!base) throw new Error('尚未設定 VITE_GAS_URL');
  const url = toUrl(base, { mode: 'verifyLogin', name, birthday, t: String(Date.now()) });
  return await fetchJsonNoStore<GasLoginResult>(url);
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
  const url = toUrl(base, { mode: 'getSheets', wh: warehouse, t: String(Date.now()) });
  const json = await fetchJsonNoStore<{ sheetNames?: string[]; error?: string }>(url);
  if (json.error) throw new Error(json.error);
  return Array.isArray(json.sheetNames) ? json.sheetNames : [];
}

export async function gasQuerySheet(
  warehouse: string,
  sheet: string,
  name?: string
): Promise<GasPayload> {
  const base = getBaseUrl();
  if (!base) throw new Error('尚未設定 VITE_GAS_URL');
  const url = toUrl(base, {
    mode: 'api',
    wh: warehouse,
    sheet,
    name: (name || '').trim(),
    t: String(Date.now()),
  });
  const json = await fetchJsonNoStore<GasPayload>(url);
  if ((json as any).error) throw new Error(String((json as any).error));
  return json;
}

export function gasIsConfigured(): boolean {
  return Boolean(getBaseUrl());
}
