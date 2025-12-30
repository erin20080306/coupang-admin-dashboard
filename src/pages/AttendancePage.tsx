import { useEffect, useMemo, useRef, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import type { AttendanceSummary } from '../lib/mockApi';
import { exportCsv, exportElementPng, exportExcelHtml } from '../lib/export';

type Item = { id: string; name: string; summary: AttendanceSummary };

type SortDir = 'asc' | 'desc';
type SortKey = 'status' | 'name' | 'rate' | 'counts';

type Stored = {
  warehouse: string;
  page: string;
  items: Item[];
};

function labelFor(status: AttendanceSummary['status']): string {
  if (status === 'normal') return '正常';
  if (status === 'low') return '偏低';
  return '異常';
}

export default function AttendancePage() {
  const navigate = useNavigate();
  const [data, setData] = useState<Stored | null>(null);
  const wrapRef = useRef<HTMLDivElement | null>(null);

  const [sortKey, setSortKey] = useState<SortKey>('rate');
  const [sortDir, setSortDir] = useState<SortDir>('asc');

  useEffect(() => {
    try {
      const raw = sessionStorage.getItem('coupang_att_full');
      if (!raw) return;
      setData(JSON.parse(raw) as Stored);
    } catch {
      setData(null);
    }
  }, []);

  const sorted = useMemo(() => {
    const items = data?.items ? [...data.items] : [];

    function statusRank(s: AttendanceSummary['status']): number {
      if (s === 'abnormal') return 0;
      if (s === 'low') return 1;
      return 2;
    }

    function cmp(a: Item, b: Item): number {
      if (sortKey === 'status') {
        const v = statusRank(a.summary.status) - statusRank(b.summary.status);
        if (v !== 0) return v;
        return a.summary.rate - b.summary.rate;
      }
      if (sortKey === 'name') {
        return a.name.localeCompare(b.name, 'zh-Hant', { numeric: true, sensitivity: 'base' });
      }
      if (sortKey === 'counts') {
        const ar = a.summary.expected ? a.summary.attended / a.summary.expected : 0;
        const br = b.summary.expected ? b.summary.attended / b.summary.expected : 0;
        const v = ar - br;
        if (v !== 0) return v;
        if (a.summary.attended !== b.summary.attended) return a.summary.attended - b.summary.attended;
        return a.summary.expected - b.summary.expected;
      }
      return a.summary.rate - b.summary.rate;
    }

    items.sort((a, b) => {
      const v = cmp(a, b);
      return sortDir === 'asc' ? v : -v;
    });
    return items;
  }, [data, sortDir, sortKey]);

  function onSort(key: SortKey) {
    if (sortKey === key) {
      setSortDir((d) => (d === 'asc' ? 'desc' : 'asc'));
    } else {
      setSortKey(key);
      setSortDir('asc');
    }
  }

  const exportRows = useMemo(() => {
    return sorted.map((it) => [
      labelFor(it.summary.status),
      it.name,
      `${Math.round(it.summary.rate * 100)}%`,
      `${it.summary.attended}/${it.summary.expected}`,
    ]);
  }, [sorted]);

  return (
    <div className="page">
      <header className="topbar">
        <div className="topbarLeft">
          <div className="brandMark">CP</div>
          <div>
            <div className="brandTitle">出勤率清單</div>
            <div className="brandSub">
              {data ? `${data.warehouse} · ${data.page}` : '—'}
            </div>
          </div>
        </div>
        <div className="topbarRight">
          <button className="btnGhost" onClick={() => navigate(-1)}>
            返回
          </button>
        </div>
      </header>

      <main className="content">
        <section className="panel">
          <div className="panelTitle" style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 10 }}>
            <span>全部（由低到高）</span>
            <div style={{ display: 'flex', gap: 10 }}>
              <button
                className="btnGhost"
                onClick={() => exportExcelHtml('出勤率清單', ['狀態', '姓名', '出勤率', '實到/應到'], exportRows)}
                disabled={!sorted.length}
              >
                Excel
              </button>
              <button
                className="btnGhost"
                onClick={() => exportCsv('出勤率清單', ['狀態', '姓名', '出勤率', '實到/應到'], exportRows)}
                disabled={!sorted.length}
              >
                CSV
              </button>
              <button
                className="btnGhost"
                onClick={() => {
                  if (!wrapRef.current) return;
                  exportElementPng(wrapRef.current, '出勤率清單');
                }}
                disabled={!sorted.length}
              >
                PNG
              </button>
            </div>
          </div>
          {sorted.length ? (
            <div className="tableWrap" ref={wrapRef}>
              <table className="table">
                <thead>
                  <tr>
                    <th className="thSortable" onClick={() => onSort('status')}>
                      <div className="thInner">
                        <span>狀態</span>
                        <span className="sortPair" aria-hidden="true">
                          <span className={`sortTri up${sortKey === 'status' && sortDir === 'asc' ? ' on' : ''}`} />
                          <span className={`sortTri down${sortKey === 'status' && sortDir === 'desc' ? ' on' : ''}`} />
                        </span>
                      </div>
                    </th>
                    <th className="thSortable" onClick={() => onSort('name')}>
                      <div className="thInner">
                        <span>姓名</span>
                        <span className="sortPair" aria-hidden="true">
                          <span className={`sortTri up${sortKey === 'name' && sortDir === 'asc' ? ' on' : ''}`} />
                          <span className={`sortTri down${sortKey === 'name' && sortDir === 'desc' ? ' on' : ''}`} />
                        </span>
                      </div>
                    </th>
                    <th className="thSortable" onClick={() => onSort('rate')}>
                      <div className="thInner">
                        <span>出勤率</span>
                        <span className="sortPair" aria-hidden="true">
                          <span className={`sortTri up${sortKey === 'rate' && sortDir === 'asc' ? ' on' : ''}`} />
                          <span className={`sortTri down${sortKey === 'rate' && sortDir === 'desc' ? ' on' : ''}`} />
                        </span>
                      </div>
                    </th>
                    <th className="thSortable" onClick={() => onSort('counts')}>
                      <div className="thInner">
                        <span>實到/應到</span>
                        <span className="sortPair" aria-hidden="true">
                          <span className={`sortTri up${sortKey === 'counts' && sortDir === 'asc' ? ' on' : ''}`} />
                          <span className={`sortTri down${sortKey === 'counts' && sortDir === 'desc' ? ' on' : ''}`} />
                        </span>
                      </div>
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {sorted.map((it) => (
                    <tr key={it.id}>
                      <td>
                        <span className={`badge badge-${it.summary.status}`}>{labelFor(it.summary.status)}</span>
                      </td>
                      <td>{it.name}</td>
                      <td>{Math.round(it.summary.rate * 100)}%</td>
                      <td>
                        {it.summary.attended}/{it.summary.expected}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : (
            <div style={{ color: 'var(--muted)', fontSize: 13 }}>沒有可顯示的出勤資料。</div>
          )}
        </section>
      </main>
    </div>
  );
}
