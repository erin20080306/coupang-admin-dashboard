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
};

export type GasWarehouseIdResult =
  | { ok: true; warehouse: string; spreadsheetId: string }
  | { ok: false; error: string };

function getBaseUrl(): string | null {
  const v = (import.meta as any).env?.VITE_GAS_URL as string | undefined;
  const s = (v || '').trim();
  return s ? s : null;
}

export async function gasGetWarehouseId(warehouse: string): Promise<string> {
  const base = getBaseUrl();
  if (!base) throw new Error('尚未設定 VITE_GAS_URL');
  const url = toUrl(base, { mode: 'getWarehouseId', wh: warehouse, t: String(Date.now()) });
  const res = await fetch(url, { cache: 'no-store' });
  if (!res.ok) throw new Error(`API 錯誤 ${res.status}`);
  const json = (await res.json()) as GasWarehouseIdResult;
  if (!json || json.ok === false) throw new Error((json as any)?.error || '取得試算表 ID 失敗');
  return json.spreadsheetId;
}

export async function gasVerifyLogin(name: string, birthday: string): Promise<GasLoginResult> {
  const base = getBaseUrl();
  if (!base) throw new Error('尚未設定 VITE_GAS_URL');
  const url = toUrl(base, { mode: 'verifyLogin', name, birthday, t: String(Date.now()) });
  const res = await fetch(url, { cache: 'no-store' });
  if (!res.ok) throw new Error(`API 錯誤 ${res.status}`);
  const json = (await res.json()) as GasLoginResult;
  return json;
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
  const res = await fetch(url, { cache: 'no-store' });
  if (!res.ok) throw new Error(`API 錯誤 ${res.status}`);
  const json = (await res.json()) as { sheetNames?: string[]; error?: string };
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
  const res = await fetch(url, { cache: 'no-store' });
  if (!res.ok) throw new Error(`API 錯誤 ${res.status}`);
  const json = (await res.json()) as GasPayload;
  if ((json as any).error) throw new Error(String((json as any).error));
  return json;
}

export function gasIsConfigured(): boolean {
  return Boolean(getBaseUrl());
}
