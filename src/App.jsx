import { useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import {
  Bar, BarChart, CartesianGrid, Cell, Line, LineChart,
  ResponsiveContainer, Tooltip, XAxis, YAxis,
} from 'recharts'

const C = { coral: '#cc785c', teal: '#5db8a6', ink: '#141413', muted: '#6c6a64', hair: '#e6dfd8', error: '#c64545' }
const QUARTERS = ['Q1', 'Q2', 'Q3', 'Q4']
const fmt = (n) => Math.round(n).toLocaleString('ko-KR')
const toNum = (v) => (typeof v === 'number' ? v : Number(String(v ?? '').replace(/,/g, '')) || 0)

// 첫 번째 시트에서 헤더 행("법인명" 포함)을 찾아 객체 배열로 변환
function parseSheet(wb) {
  const ws = wb.Sheets[wb.SheetNames[0]]
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' })
  let h = rows.findIndex((r) => r.some((c) => String(c).includes('법인명')))
  if (h < 0) h = rows.findIndex((r) => r.filter((c) => c !== '').length >= 5)
  if (h < 0) throw new Error('헤더 행을 찾을 수 없습니다.')
  const headers = rows[h].map((c) => String(c).trim())
  const data = rows.slice(h + 1)
    .filter((r) => String(r[0]).trim() !== '')
    .map((r) => Object.fromEntries(headers.map((k, i) => [k, r[i]])))
  return { sheetName: wb.SheetNames[0], headers, data }
}

function Card({ label, value, unit, tone }) {
  return (
    <div className={`kpi ${tone || ''}`}>
      <div className="kpi-label">{label}</div>
      <div className="kpi-value">{value}<span>{unit}</span></div>
    </div>
  )
}

function ChartCard({ title, sub, children }) {
  return (
    <section className="card">
      <h3>{title}</h3>
      <p className="sub">{sub}</p>
      {children}
    </section>
  )
}

const tip = {
  contentStyle: { background: '#faf9f5', border: `1px solid ${C.hair}`, borderRadius: 8, fontSize: 13 },
  formatter: (v) => fmt(v),
}
const axisTick = { fill: C.muted, fontSize: 12 }

export default function App() {
  const [file, setFile] = useState(null)
  const [error, setError] = useState('')
  const [sheet, setSheet] = useState(null)
  const [q, setQ] = useState('')
  const [region, setRegion] = useState('전체')
  const [segment, setSegment] = useState('전체')
  const [product, setProduct] = useState('전체')
  const [sort, setSort] = useState({ key: null, dir: 1 })
  const [drag, setDrag] = useState(false)

  const load = async (f) => {
    if (!f) return
    setError('')
    if (!/\.xlsx$/i.test(f.name)) return setError('.xlsx 파일만 업로드할 수 있습니다.')
    try {
      const wb = XLSX.read(await f.arrayBuffer(), { type: 'array' })
      const parsed = parseSheet(wb)
      if (!parsed.data.length) throw new Error('데이터 행이 없습니다.')
      setSheet(parsed); setFile(f.name); setQ(''); setRegion('전체'); setSegment('전체'); setProduct('전체'); setSort({ key: null, dir: 1 })
    } catch (e) {
      setError(`파일을 읽을 수 없습니다: ${e.message}`)
    }
  }

  const distinct = (k) => (sheet && sheet.headers.includes(k) ? [...new Set(sheet.data.map((r) => r[k]))].filter(Boolean) : [])
  const regions = useMemo(() => distinct('지역'), [sheet])
  const segments = useMemo(() => distinct('사업부문'), [sheet])
  const products = useMemo(() => distinct('주요제품'), [sheet])

  const rows = useMemo(() => {
    if (!sheet) return []
    const kw = q.trim().toLowerCase()
    let out = sheet.data.filter((r) =>
      (region === '전체' || r['지역'] === region) &&
      (segment === '전체' || r['사업부문'] === segment) &&
      (product === '전체' || r['주요제품'] === product) &&
      (!kw || Object.values(r).some((v) => String(v).toLowerCase().includes(kw))))
    if (sort.key) {
      out = [...out].sort((a, b) => {
        const x = a[sort.key], y = b[sort.key]
        return (typeof x === 'number' && typeof y === 'number' ? x - y : String(x).localeCompare(String(y), 'ko')) * sort.dir
      })
    }
    return out
  }, [sheet, q, region, segment, product, sort])

  const stats = useMemo(() => {
    const sum = (k) => rows.reduce((s, r) => s + toNum(r[k]), 0)
    const rev = sum('연매출'), op = sum('연영업이익')
    return {
      n: rows.length, rev, op, margin: rev ? (op / rev) * 100 : 0,
      loss: rows.filter((r) => toNum(r['연영업이익']) < 0).length,
      emp: sum('임직원수'),
      product: (() => {
        const m = {}
        rows.forEach((r) => { const p = String(r['주요제품'] ?? '').trim(); if (p) m[p] = (m[p] || 0) + toNum(r['연매출']) })
        return Object.entries(m).sort((a, b) => b[1] - a[1])[0]?.[0] || '-'
      })(),
    }
  }, [rows])

  const byRegion = useMemo(() => {
    const m = {}
    rows.forEach((r) => { const k = r['지역'] || '기타'; m[k] = (m[k] || 0) + toNum(r['연매출']) })
    return Object.entries(m).map(([name, 매출]) => ({ name, 매출 })).sort((a, b) => b.매출 - a.매출)
  }, [rows])

  const byQuarter = useMemo(() => QUARTERS.map((qt) => ({
    name: qt,
    매출: rows.reduce((s, r) => s + toNum(r[`${qt}매출`]), 0),
  })), [rows])

  const byEntity = useMemo(() => rows
    .map((r) => ({ name: r['법인명'], 영업이익: toNum(r['연영업이익']) }))
    .sort((a, b) => b.영업이익 - a.영업이익)
    .slice(0, 15), [rows])

  const isNumCol = (k) => sheet.data.some((r) => typeof r[k] === 'number')
  const toggleSort = (k) => setSort((s) => (s.key === k ? { key: k, dir: -s.dir } : { key: k, dir: 1 }))
  const reset = () => { setQ(''); setRegion('전체'); setSegment('전체'); setProduct('전체') }

  return (
    <div className="app">
      <header className="nav">
        <div className="wrap nav-in">
          <div className="brand"><Spike />Global Dashboard</div>
          {file && <span className="badge">{file}</span>}
        </div>
      </header>

      <main className="wrap">
        <section className="hero">
          <h1>2026년 해외법인 분기별 실적,<br />한눈에 살펴보세요.</h1>
          <p className="lead">엑셀 파일을 업로드하면 표와 차트로 바로 보여드립니다. 파일은 브라우저에서만 처리되며 서버로 전송되지 않습니다.</p>
          <label
            className={`drop ${drag ? 'on' : ''}`}
            onDragOver={(e) => { e.preventDefault(); setDrag(true) }}
            onDragLeave={() => setDrag(false)}
            onDrop={(e) => { e.preventDefault(); setDrag(false); load(e.dataTransfer.files[0]) }}
          >
            <input type="file" accept=".xlsx" hidden onChange={(e) => { load(e.target.files[0]); e.target.value = '' }} />
            <span className="btn">{sheet ? '다른 파일 선택' : '.xlsx 파일 선택'}</span>
            <span className="hint">또는 여기로 파일을 끌어다 놓으세요</span>
          </label>
          {error && <p className="error">{error}</p>}
        </section>

        {sheet && (
          <>
            <section className="filters">
              <input className="input search" placeholder="법인명, 국가, 제품 등 검색" value={q} onChange={(e) => setQ(e.target.value)} />
              {regions.length > 0 && (
                <select className="input" value={region} onChange={(e) => setRegion(e.target.value)}>
                  <option>전체</option>{regions.map((r) => <option key={r}>{r}</option>)}
                </select>
              )}
              {segments.length > 0 && (
                <select className="input" value={segment} onChange={(e) => setSegment(e.target.value)}>
                  <option>전체</option>{segments.map((r) => <option key={r}>{r}</option>)}
                </select>
              )}
              {products.length > 0 && (
                <select className="input" value={product} onChange={(e) => setProduct(e.target.value)}>
                  <option value="전체">주요제품: 전체</option>{products.map((r) => <option key={r}>{r}</option>)}
                </select>
              )}
              <button className="btn secondary" onClick={reset}>초기화</button>
              <span className="count">{rows.length} / {sheet.data.length}개 법인</span>
            </section>

            {rows.length === 0 ? (
              <p className="empty">조건에 맞는 데이터가 없습니다.</p>
            ) : (
              <>
                <section className="kpis">
                  <Card label="법인 수" value={stats.n} unit="개" />
                  <Card label="연매출 합계" value={(stats.rev / 1e6).toFixed(2)} unit=" 조원" tone="dark" />
                  <Card label="연영업이익 합계" value={(stats.op / 1000).toLocaleString('ko-KR', { maximumFractionDigits: 0 })} unit=" 십억원" />
                  <Card label="영업이익률" value={stats.margin.toFixed(1)} unit="%" />
                  <Card label="임직원수" value={fmt(stats.emp)} unit="명" />
                  <Card label="주요제품 (매출 1위)" value={stats.product} unit="" tone="text" />
                  <Card label="적자 법인" value={stats.loss} unit="개" tone={stats.loss ? 'coral' : ''} />
                </section>

                <div className="grid">
                  <ChartCard title="지역별 매출" sub="연매출 합계 · 백만원">
                    <ResponsiveContainer width="100%" height={280}>
                      <BarChart data={byRegion}>
                        <CartesianGrid stroke={C.hair} vertical={false} />
                        <XAxis dataKey="name" tick={axisTick} axisLine={false} tickLine={false} />
                        <YAxis tick={axisTick} axisLine={false} tickLine={false} tickFormatter={(v) => `${fmt(v / 1000)}k`} />
                        <Tooltip {...tip} cursor={{ fill: '#efe9de' }} />
                        <Bar dataKey="매출" fill={C.coral} radius={[4, 4, 0, 0]} />
                      </BarChart>
                    </ResponsiveContainer>
                  </ChartCard>

                  <ChartCard title="분기별 매출 추이" sub="Q1–Q4 합계 · 백만원">
                    <ResponsiveContainer width="100%" height={280}>
                      <LineChart data={byQuarter}>
                        <CartesianGrid stroke={C.hair} vertical={false} />
                        <XAxis dataKey="name" tick={axisTick} axisLine={false} tickLine={false} />
                        <YAxis tick={axisTick} axisLine={false} tickLine={false} domain={['auto', 'auto']} tickFormatter={(v) => `${fmt(v / 1000)}k`} />
                        <Tooltip {...tip} />
                        <Line type="monotone" dataKey="매출" stroke={C.ink} strokeWidth={2.5} dot={{ r: 4, fill: C.coral, stroke: C.coral }} />
                      </LineChart>
                    </ResponsiveContainer>
                  </ChartCard>
                </div>

                <ChartCard
                  title={rows.length > 15 ? '법인별 영업이익 (상위 15개)' : '법인별 영업이익'}
                  sub="연영업이익 · 백만원 (적자는 빨간색)"
                >
                  <ResponsiveContainer width="100%" height={Math.max(240, byEntity.length * 30)}>
                    <BarChart data={byEntity} layout="vertical" margin={{ left: 10 }}>
                      <CartesianGrid stroke={C.hair} horizontal={false} />
                      <XAxis type="number" tick={axisTick} axisLine={false} tickLine={false} tickFormatter={fmt} />
                      <YAxis type="category" dataKey="name" width={200} tick={{ fill: C.ink, fontSize: 12 }} axisLine={false} tickLine={false} />
                      <Tooltip {...tip} cursor={{ fill: '#efe9de' }} />
                      <Bar dataKey="영업이익" radius={4}>
                        {byEntity.map((d, i) => <Cell key={i} fill={d.영업이익 < 0 ? C.error : C.teal} />)}
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </ChartCard>

                <section className="card">
                  <h3>원본 데이터</h3>
                  <p className="sub">{sheet.sheetName} · 헤더를 클릭하면 정렬됩니다</p>
                  <div className="table-wrap">
                    <table>
                      <thead>
                        <tr>
                          {sheet.headers.map((h) => (
                            <th key={h} onClick={() => toggleSort(h)} className={isNumCol(h) ? 'num' : ''}>
                              {h}{sort.key === h ? (sort.dir > 0 ? ' ▲' : ' ▼') : ''}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {rows.map((r, i) => (
                          <tr key={i}>
                            {sheet.headers.map((h) => (
                              <td key={h} className={typeof r[h] === 'number' ? `num ${r[h] < 0 ? 'neg' : ''}` : ''}>
                                {typeof r[h] === 'number' ? fmt(r[h]) : String(r[h] ?? '')}
                              </td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </section>
              </>
            )}
          </>
        )}
      </main>

      <footer className="footer"><div className="wrap"><Spike light />Global Dashboard · 파일은 브라우저에서만 처리됩니다</div></footer>
    </div>
  )
}

function Spike({ light }) {
  return (
    <svg width="18" height="18" viewBox="0 0 24 24" aria-hidden="true" style={{ marginRight: 8, verticalAlign: '-3px' }}>
      <g stroke={light ? '#faf9f5' : '#141413'} strokeWidth="2.4" strokeLinecap="round" fill="none">
        <path d="M12 2v20M2 12h20M5 5l14 14M19 5L5 19" />
      </g>
    </svg>
  )
}
