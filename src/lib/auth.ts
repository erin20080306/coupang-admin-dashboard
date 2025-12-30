export type AuthUser = {
  name: string;
  isAdmin: boolean;
  warehouseKey?: string;
};

const TOKEN_KEY = 'coupang_demo_token';
const USER_KEY = 'coupang_demo_user';

export function getToken(): string | null {
  return localStorage.getItem(TOKEN_KEY);
}

export function getUser(): AuthUser | null {
  const raw = localStorage.getItem(USER_KEY);
  if (!raw) return null;
  try {
    return JSON.parse(raw) as AuthUser;
  } catch {
    return null;
  }
}

export function isAuthed(): boolean {
  return Boolean(getToken());
}

export function setSession(user: AuthUser): void {
  const token = `gas-token:${user.isAdmin ? 'admin' : 'user'}`;
  localStorage.setItem(TOKEN_KEY, token);
  localStorage.setItem(USER_KEY, JSON.stringify(user));
}

export function logout(): void {
  localStorage.removeItem(TOKEN_KEY);
  localStorage.removeItem(USER_KEY);
}
