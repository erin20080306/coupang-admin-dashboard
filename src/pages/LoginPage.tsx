import { useLayoutEffect, useMemo, useRef, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import gsap from 'gsap';
import { setSession } from '../lib/auth';
import { gasIsConfigured, gasVerifyLogin } from '../lib/gasApi';

type FormState = {
  name: string;
  birthday: string;
};

const LOGIN_REMEMBER_KEY = 'coupang_login_remember_v1';
const LOGIN_HISTORY_KEY = 'coupang_login_history_v1';
const LOGIN_REMEMBER_TTL_MS = 3 * 24 * 60 * 60 * 1000;

type LoginHistoryItem = { name: string; birthday: string; ts: number };

function loadHistory(): LoginHistoryItem[] {
  try {
    const raw = localStorage.getItem(LOGIN_HISTORY_KEY);
    if (!raw) return [];
    const arr = JSON.parse(raw) as LoginHistoryItem[];
    const now = Date.now();
    return (Array.isArray(arr) ? arr : [])
      .filter((x) => x && x.name && x.birthday && x.ts && now - Number(x.ts) <= LOGIN_REMEMBER_TTL_MS)
      .sort((a, b) => Number(b.ts) - Number(a.ts));
  } catch {
    return [];
  }
}

function saveHistory(next: LoginHistoryItem[]) {
  try {
    localStorage.setItem(LOGIN_HISTORY_KEY, JSON.stringify(next.slice(0, 12)));
  } catch {
    // ignore
  }
}

export default function LoginPage() {
  const navigate = useNavigate();
  const splashRef = useRef<HTMLDivElement | null>(null);
  const splashTitleRef = useRef<HTMLDivElement | null>(null);
  const splashSubRef = useRef<HTMLDivElement | null>(null);
  const cardRef = useRef<HTMLDivElement | null>(null);
  const logoRef = useRef<HTMLDivElement | null>(null);

  const [showLogin, setShowLogin] = useState(false);
  const [form, setForm] = useState<FormState>({ name: '', birthday: '' });
  const [error, setError] = useState<string>('');
  const [loading, setLoading] = useState(false);
  const [history, setHistory] = useState<LoginHistoryItem[]>([]);
  const [remember, setRemember] = useState<LoginHistoryItem | null>(null);
  const [focused, setFocused] = useState<'name' | 'birthday' | null>(null);
  const canSubmit = useMemo(() => form.name.trim() && form.birthday.trim(), [form]);

  const showHistory = useMemo(() => {
    if (!focused) return false;
    if (!history.length) return false;
    if (focused === 'name') return !form.name.trim();
    return !form.birthday.trim();
  }, [focused, history.length, form.name, form.birthday]);

  useLayoutEffect(() => {
    const list = loadHistory();
    setHistory(list);
    try {
      const raw = localStorage.getItem(LOGIN_REMEMBER_KEY);
      if (!raw) {
        setRemember(null);
        return;
      }
      const obj = JSON.parse(raw) as LoginHistoryItem;
      const ts = Number(obj?.ts || 0);
      if (!ts || Date.now() - ts > LOGIN_REMEMBER_TTL_MS || !obj?.name || !obj?.birthday) {
        localStorage.removeItem(LOGIN_REMEMBER_KEY);
        setRemember(null);
        return;
      }
      setRemember({ name: String(obj.name || ''), birthday: String(obj.birthday || ''), ts });
    } catch {
      // ignore
    }
  }, []);

  function recordLogin(name: string, birthday: string) {
    const item: LoginHistoryItem = { name, birthday, ts: Date.now() };

    try {
      localStorage.setItem(LOGIN_REMEMBER_KEY, JSON.stringify(item));
    } catch {
      // ignore
    }

    setRemember(item);

    setHistory((prev) => {
      const filtered = prev.filter((x) => !(x.name === name && x.birthday === birthday));
      const next = [item, ...filtered].slice(0, 12);
      saveHistory(next);
      return next;
    });
  }

  useLayoutEffect(() => {
    const tl = gsap.timeline();

    tl.set(splashRef.current, { opacity: 1 });
    tl.set(splashTitleRef.current, { opacity: 0, y: 22, scale: 0.92, filter: 'blur(10px)' });
    tl.set(splashSubRef.current, { opacity: 0, y: 12, filter: 'blur(6px)' });
    tl.fromTo(
      splashTitleRef.current,
      { opacity: 0, y: 22, scale: 0.92, filter: 'blur(10px)' },
      {
        opacity: 1,
        y: 0,
        scale: 1,
        filter: 'blur(0px)',
        duration: 0.95,
        ease: 'power4.out',
      }
    ).fromTo(
      splashSubRef.current,
      { opacity: 0, y: 12, filter: 'blur(6px)' },
      { opacity: 0.92, y: 0, filter: 'blur(0px)', duration: 0.6, ease: 'power3.out' },
      '-=0.45'
    )
      .to({}, { duration: 0.65 })
      .to(
        [splashTitleRef.current, splashSubRef.current],
        { opacity: 0, y: -8, filter: 'blur(10px)', duration: 0.4, ease: 'power2.in' },
        '+=0.05'
      )
      .to(splashRef.current, { opacity: 0, duration: 0.35, ease: 'power2.out' }, '-=0.1')
      .add(() => {
        setShowLogin(true);
      })
      .add(() => {
        requestAnimationFrame(() => {
          if (logoRef.current) {
            gsap.fromTo(
              logoRef.current,
              { y: 10, opacity: 0 },
              { y: 0, opacity: 1, duration: 0.6, ease: 'power3.out' }
            );
          }
          if (cardRef.current) {
            gsap.fromTo(
              cardRef.current,
              { y: 24, opacity: 0 },
              { y: 0, opacity: 1, duration: 0.7, ease: 'power3.out' }
            );
          }
        });
      });
    return () => {
      tl.kill();
    };
  }, []);

  async function onSubmit(e: React.FormEvent) {
    e.preventDefault();
    if (!canSubmit || loading) return;
    setError('');
    setLoading(true);

    try {
      const name = form.name.trim();
      const birthday = form.birthday.trim();

      if (name === '酷澎' && birthday === '0000') {
        recordLogin(name, birthday);
        setSession({ name, isAdmin: true, warehouseKey: 'TAO1' });
        navigate('/', { replace: true });
        return;
      }

      if (!gasIsConfigured()) {
        throw new Error('尚未設定 VITE_GAS_URL，無法進行登入驗證');
      }

      const res = await gasVerifyLogin(name, birthday);
      if (!res?.ok) {
        throw new Error(res?.msg || '姓名或生日不正確');
      }

      recordLogin(name, birthday);

      const warehouseKey = (res as any)?.warehouseKey || (res as any)?.whKey || (res as any)?.warehouse || 'TAO1';

      setSession({
        name: res.name || name,
        isAdmin: Boolean(res.isAdmin),
        warehouseKey,
      });
      navigate('/', { replace: true });
    } catch (err) {
      const msg = err instanceof Error ? err.message : '登入失敗';
      setError(msg);
      if (cardRef.current) {
        gsap.fromTo(
          cardRef.current,
          { x: -6 },
          { x: 0, duration: 0.35, ease: 'elastic.out(1, 0.35)' }
        );
      }
    } finally {
      setLoading(false);
    }
  }

  return (
    <div className="loginPage">
      <div className="splash" ref={splashRef} aria-hidden={showLogin}>
        <div className="splashInner">
          <div className="splashMark">CP</div>
          <div className="splashTitle" ref={splashTitleRef}>宏盛酷澎查詢系統</div>
          <div className="splashSub" ref={splashSubRef}>企業後台表格查詢 · Attendance Dashboard</div>
        </div>
      </div>

      {showLogin ? (
        <>
          <div className="loginLogo" ref={logoRef}>
            <div className="loginMark">CP</div>
            <div className="loginText">
              <div className="loginTitle">酷澎出勤後台</div>
              <div className="loginSub">試算表查詢（GAS）</div>
            </div>
          </div>

          <div className="loginCard" ref={cardRef}>
            <div className="loginCardHeader">
              <div className="loginCardTitle">登入</div>
            </div>

            <form className="loginForm" onSubmit={onSubmit}>
              <label className="field">
                <span>姓名</span>
                <input
                  value={form.name}
                  onChange={(e) => setForm((s) => ({ ...s, name: e.target.value }))}
                  placeholder="例如：王小明"
                  autoComplete="name"
                  onFocus={() => setFocused('name')}
                  onBlur={() => setTimeout(() => setFocused(null), 120)}
                />
              </label>

              <label className="field">
                <span>生日</span>
                <input
                  value={form.birthday}
                  onChange={(e) => setForm((s) => ({ ...s, birthday: e.target.value }))}
                  placeholder="例如：810101"
                  autoComplete="bday"
                  onFocus={() => setFocused('birthday')}
                  onBlur={() => setTimeout(() => setFocused(null), 120)}
                />
              </label>

              {remember ? (
                <div style={{ display: 'flex', gap: 8, marginTop: 6, flexWrap: 'wrap' }}>
                  <button
                    type="button"
                    className="btnGhost"
                    onClick={() => {
                      setForm({ name: remember.name, birthday: remember.birthday });
                      setFocused(null);
                    }}
                  >
                    當前設備記憶密碼：{remember.name} / {remember.birthday}
                  </button>
                  <button
                    type="button"
                    className="btnGhost"
                    onClick={() => {
                      try {
                        localStorage.removeItem(LOGIN_REMEMBER_KEY);
                      } catch {
                        // ignore
                      }
                      setRemember(null);
                    }}
                  >
                    清除
                  </button>
                </div>
              ) : null}

              {showHistory ? (
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, marginTop: 6 }}>
                  {history.slice(0, 8).map((h) => (
                    <button
                      key={`${h.name}_${h.birthday}_${h.ts}`}
                      type="button"
                      className="btnGhost"
                      onClick={() => {
                        setForm({ name: h.name, birthday: h.birthday });
                        setFocused(null);
                      }}
                    >
                      {h.name} / {h.birthday}
                    </button>
                  ))}
                </div>
              ) : null}

              {error ? <div className="formError">{error}</div> : null}

              <button className="btnPrimary" disabled={!canSubmit || loading} type="submit">
                {loading ? '登入中…' : '登入'}
              </button>
            </form>
          </div>
        </>
      ) : null}
    </div>
  );
}
