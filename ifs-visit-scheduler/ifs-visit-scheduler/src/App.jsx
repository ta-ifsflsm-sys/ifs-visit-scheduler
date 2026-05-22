import { useState, useEffect, useCallback, useRef } from 'react'
import { supabase } from './lib/supabase'
import { downloadICS } from './utils/ics'

// ── constants ────────────────────────────────────────────────────────────────
const SH = 9, EH = 18, HPX = 80, TH = (EH - SH) * HPX

const VCOLS = [
  { bg: '#FAECE7', col: '#993C1D', bd: '#D85A30' },
  { bg: '#E6F1FB', col: '#185FA5', bd: '#378ADD' },
  { bg: '#FAEEDA', col: '#854F0B', bd: '#BA7517' },
  { bg: '#E1F5EE', col: '#0F6E56', bd: '#1D9E75' },
  { bg: '#EEEDFE', col: '#534AB7', bd: '#7F77DD' },
  { bg: '#FBEAF0', col: '#993556', bd: '#D4537E' },
]

const OWC = {
  TY: { bg: '#EEEDFE', col: '#3C3489' },
  SK: { bg: '#E1F5EE', col: '#0F6E56' },
  TK: { bg: '#FAC775', col: '#633806' },
}

const STC = {
  confirmed: { bg: '#EAF3DE', bd: '#639922', tx: '#27500A' },
  tentative:  { bg: '#FAEEDA', bd: '#BA7517', tx: '#633806' },
  travel:     { bg: '#E6F1FB', bd: '#185FA5', tx: '#0C447C' },
}

const OWNERS = ['TY', 'SK', 'TK']

const T = {
  ja: {
    days: ['月', '火', '水', '木', '金'],
    wk: '週の開始（月）:', ow: '担当者:', st: '状態:', vis: '来訪者:',
    all: '全員', allst: '全て', allvis: '全来訪者',
    conf: '確定', tent: '仮予約', trvl: '移動', add: '新規ミーティング',
    exec: '担当来訪者', time: '時間', atts: '参加者', notes: 'メモ',
    brf: 'ブリーフィング', ics: 'カレンダーに追加 (.ics)',
    save: '保存', cancel: 'キャンセル', del: '削除',
    addT: 'ミーティングを追加', editT: 'ミーティングを編集',
    fttl: 'タイトル (JP)', fattl: 'タイトル (EN)',
    fday: '曜日', fst: '開始', fen: '終了',
    fstat: '状態', fow: '担当者', fexec: '来訪者',
    fatts: '参加者（カンマ区切り）', fnotes: 'メモ', fbrf: 'ブリーフィング（確定のみ）',
    ownN: { TY: 'ティアナ', SK: '佐藤', TK: '田中' },
    visSetup: '来訪者設定', visDesc: 'この週の来日エグゼクティブ',
    visAdd: '来訪者を追加', visName: '氏名 (JP)', visNameEn: '氏名 (EN)',
    visRole: '役職', visColor: 'カラー', visNone: '来訪者未登録',
    noExec: '── なし ──', allExec: '全来訪者', otherExec: 'その他...',
    hint: 'ブロックをクリックで詳細 ／ 空白クリックで追加',
    loading: '読み込み中...', saving: '保存中...',
    del_confirm: 'このミーティングを削除しますか？',
    vis_del_confirm: 'この来訪者を削除しますか？（関連ミーティングの紐付けは解除されます）',
  },
  en: {
    days: ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'],
    wk: 'Week start (Mon):', ow: 'Owner:', st: 'Status:', vis: 'Visitors:',
    all: 'All', allst: 'All', allvis: 'All Visitors',
    conf: 'Confirmed', tent: 'Tentative', trvl: 'Travel', add: 'Add Meeting',
    exec: 'Executive', time: 'Time', atts: 'Attendees', notes: 'Notes',
    brf: 'Briefing', ics: 'Add to Calendar (.ics)',
    save: 'Save', cancel: 'Cancel', del: 'Delete',
    addT: 'Add Meeting', editT: 'Edit Meeting',
    fttl: 'Title (JP)', fattl: 'Title (EN)',
    fday: 'Day', fst: 'Start', fen: 'End',
    fstat: 'Status', fow: 'Owner', fexec: 'Visitor',
    fatts: 'Attendees (comma-separated)', fnotes: 'Notes', fbrf: 'Briefing (confirmed only)',
    ownN: { TY: 'Tiana', SK: 'Sato', TK: 'Tanaka' },
    visSetup: 'Visitor Setup', visDesc: 'Executives visiting this week',
    visAdd: 'Add Visitor', visName: 'Name (JP)', visNameEn: 'Name (EN)',
    visRole: 'Role/Title', visColor: 'Color', visNone: 'No visitors registered',
    noExec: '── None ──', allExec: 'All Visitors', otherExec: 'Other...',
    hint: 'Click block for details  •  Click empty space to add',
    loading: 'Loading...', saving: 'Saving...',
    del_confirm: 'Delete this meeting?',
    vis_del_confirm: 'Remove this visitor? (linked meetings will be unlinked)',
  },
}

// ── helpers ──────────────────────────────────────────────────────────────────
const h2y   = h => (h - SH) * HPX
const ft    = h => { const hr = Math.floor(h), mn = Math.round((h % 1) * 60); return `${hr}:${String(mn).padStart(2,'0')}` }
const pad2  = n => String(n).padStart(2, '0')
const fmtDate = d => `${d.getFullYear()}-${pad2(d.getMonth()+1)}-${pad2(d.getDate())}`

function getMondayOf(d = new Date()) {
  const day = d.getDay()
  const diff = (day === 0 ? -6 : 1 - day)
  const mon = new Date(d)
  mon.setDate(d.getDate() + diff)
  mon.setHours(0,0,0,0)
  return mon
}

function getWeekDates(monday) {
  return Array.from({ length: 5 }, (_, i) => {
    const d = new Date(monday)
    d.setDate(monday.getDate() + i)
    return d
  })
}

function vcol(ci) { return VCOLS[ci % VCOLS.length] }

// ── App ───────────────────────────────────────────────────────────────────────
export default function App() {
  // ── state ──
  const [lang,     setLangState]  = useState('ja')
  const [meetings, setMeetings]   = useState([])
  const [visitors, setVisitors]   = useState([])
  const [monday,   setMonday]     = useState(() => {
    const p = new URLSearchParams(window.location.search).get('week')
    if (p) { const d = new Date(p); if (!isNaN(d)) return d }
    return getMondayOf()
  })
  const [foOw,  setFoOw]   = useState('all')
  const [foSt,  setFoSt]   = useState('all')
  const [foVid, setFoVid]  = useState('all')
  const [panel, setPanel]  = useState(null)   // { type, id?, dayIdx?, startH? }
  const [saving, setSaving] = useState(false)
  const [loading, setLoading] = useState(true)
  const tl = T[lang]

  // ── URL sync ──
  useEffect(() => {
    const url = new URL(window.location)
    url.searchParams.set('week', fmtDate(monday))
    window.history.replaceState({}, '', url)
  }, [monday])

  // ── data fetch ──
  const fetchData = useCallback(async () => {
    const ws = fmtDate(monday)
    const [{ data: m }, { data: v }] = await Promise.all([
      supabase.from('meetings').select('*').eq('week_start', ws).order('day_index').order('start_time'),
      supabase.from('visitors').select('*').eq('week_start', ws).order('created_at'),
    ])
    setMeetings(m || [])
    setVisitors(v || [])
    setLoading(false)
  }, [monday])

  useEffect(() => { setLoading(true); fetchData() }, [fetchData])

  // ── realtime ──
  useEffect(() => {
    const ws = fmtDate(monday)
    const ch = supabase.channel(`scheduler-${ws}`)
      .on('postgres_changes', { event: '*', schema: 'public', table: 'meetings',  filter: `week_start=eq.${ws}` }, fetchData)
      .on('postgres_changes', { event: '*', schema: 'public', table: 'visitors',  filter: `week_start=eq.${ws}` }, fetchData)
      .subscribe()
    return () => supabase.removeChannel(ch)
  }, [monday, fetchData])

  // ── CRUD ──
  async function addMeeting(data) {
    setSaving(true)
    await supabase.from('meetings').insert({ ...data, week_start: fmtDate(monday) })
    setSaving(false)
    setPanel(null)
  }

  async function updateMeeting(id, data) {
    setSaving(true)
    await supabase.from('meetings').update(data).eq('id', id)
    setSaving(false)
    setPanel(null)
  }

  async function deleteMeeting(id) {
    if (!window.confirm(tl.del_confirm)) return
    setSaving(true)
    await supabase.from('meetings').delete().eq('id', id)
    setSaving(false)
    setPanel(null)
  }

  async function addVisitor(data) {
    setSaving(true)
    await supabase.from('visitors').insert({ ...data, week_start: fmtDate(monday) })
    setSaving(false)
  }

  async function deleteVisitor(id) {
    if (!window.confirm(tl.vis_del_confirm)) return
    setSaving(true)
    await supabase.from('visitors').delete().eq('id', id)
    if (foVid === id) setFoVid('all')
    setSaving(false)
  }

  // ── week change ──
  function changeWeek(val) {
    const [y, m, d] = val.split('-').map(Number)
    setMonday(new Date(y, m - 1, d))
    setFoVid('all')
  }

  // ── lang ──
  function setLang(l) {
    setLangState(l)
    if (panel?.type === 'detail') setPanel({ ...panel })
  }

  // ── filter helpers ──
  function visibleMeetings() {
    return meetings.filter(m => {
      const owOk  = foOw  === 'all' || m.owner === foOw
      const stOk  = foSt  === 'all' || m.status === foSt
      const vidOk = foVid === 'all' || m.visitor_id === foVid || m.visitor_scope === 'all'
      return owOk && stOk && vidOk
    })
  }

  // ── render helpers ──
  const weekDates = getWeekDates(monday)
  const today = new Date()

  function vById(id) { return visitors.find(v => v.id === id) }
  function vName(v)  { return v ? (lang === 'en' ? (v.name_en || v.name) : v.name) : '' }

  function execDisplay(m) {
    if (m.visitor_scope === 'all') return lang === 'ja' ? '全来訪者' : 'All Visitors'
    if (m.visitor_id) { const v = vById(m.visitor_id); if (v) return `${vName(v)} (${v.role})` }
    return ''
  }

  function stLabel(s) { return tl[{ confirmed:'conf', tentative:'tent', travel:'trvl' }[s]] }

  // ── ICS download ──
  function handleICS(m) {
    downloadICS(m, weekDates, visitors)
  }

  // ── render ─────────────────────────────────────────────────────────────────
  return (
    <div style={{ display:'flex', flexDirection:'column', height:'100vh', overflow:'hidden' }}>

      {/* ── topbar ── */}
      <div style={styles.topbar}>
        <div style={{ display:'flex', alignItems:'center', gap:8 }}>
          <span style={styles.logo}>IFS</span>
          <span style={styles.appTitle}>Japan Visit Scheduler</span>
          {saving && <span style={{ fontSize:10, color:'#888' }}>{tl.saving}</span>}
        </div>
        <div style={{ display:'flex', alignItems:'center', gap:6 }}>
          <button className="btn btn-p" onClick={() => setPanel({ type:'add', dayIdx:0, startH:10 })}>
            + {tl.add}
          </button>
          <div className="lang-sw">
            <button className={lang==='ja'?'active':''} onClick={() => setLang('ja')}>JA</button>
            <button className={lang==='en'?'active':''} onClick={() => setLang('en')}>EN</button>
          </div>
        </div>
      </div>

      {/* ── filter bar: owner + status ── */}
      <div style={styles.bar}>
        <div style={styles.barGroup}>
          <span style={styles.barLabel}>{tl.wk}</span>
          <input type="date" value={fmtDate(monday)} onChange={e => changeWeek(e.target.value)}
            style={{ fontSize:11, padding:'4px 7px', borderRadius:6, border:'0.5px solid var(--border-2)', background:'var(--bg)', color:'var(--text)' }} />
        </div>
        <div style={styles.barGroup}>
          <span style={styles.barLabel}>{tl.ow}</span>
          <span className={`chip${foOw==='all'?' active':''}`} onClick={() => setFoOw('all')}>{tl.all}</span>
          {OWNERS.map(k => (
            <span key={k} className={`chip${foOw===k?' active':''}`} onClick={() => setFoOw(k)}>
              <span className="dot" style={{ background: OWC[k].bg, color: OWC[k].col }}>{k}</span>
              {tl.ownN[k]}
            </span>
          ))}
        </div>
        <div style={styles.barGroup}>
          <span style={styles.barLabel}>{tl.st}</span>
          {[['all','allst'],['confirmed','conf'],['tentative','tent'],['travel','trvl']].map(([v,lk]) => (
            <span key={v} className={`chip${foSt===v?' active':''}`} onClick={() => setFoSt(v)}>{tl[lk]}</span>
          ))}
        </div>
      </div>

      {/* ── visitor bar ── */}
      <div style={{ ...styles.bar, background:'var(--bg-2)' }}>
        <span style={styles.barLabel}>{tl.vis}</span>
        <span className={`chip${foVid==='all'?' active':''}`} onClick={() => setFoVid('all')}>{tl.allvis}</span>
        {visitors.map(v => {
          const c = vcol(v.color_idx)
          const nm = vName(v)
          const isOn = foVid === v.id
          return (
            <span key={v.id} className="chip" onClick={() => setFoVid(v.id)}
              style={isOn ? { background:c.bg, color:c.col, borderColor:c.bd } : {}}>
              <span className="vdot" style={{ background:c.bg, color:c.col, borderColor:c.bd }}>{v.name.charAt(0)}</span>
              {nm}
              <span style={{ fontSize:9, opacity:.6 }}>{v.role}</span>
            </span>
          )
        })}
        <button className="btn" onClick={() => setPanel({ type:'visitors' })}
          style={{ marginLeft:'auto', fontSize:10, padding:'3px 9px' }}>
          ⚙ {tl.visSetup}
        </button>
      </div>

      {/* ── main ── */}
      <div style={{ flex:1, display:'flex', minHeight:0 }}>

        {/* ── calendar ── */}
        <div style={{ flex:1, display:'flex', flexDirection:'column', minWidth:0, background:'var(--bg)', overflow:'hidden' }}>

          {/* day headers */}
          <div style={{ display:'flex', borderBottom:'0.5px solid var(--border)', flexShrink:0 }}>
            <div style={{ width:48, flexShrink:0, borderRight:'0.5px solid var(--border)' }} />
            {weekDates.map((d, i) => {
              const isToday = d.toDateString() === today.toDateString()
              return (
                <div key={i} style={{ flex:1, padding:'5px 4px', textAlign:'center',
                  borderRight: i<4 ? '0.5px solid var(--border)' : 'none',
                  background:'var(--bg)' }}>
                  <div style={{ fontSize:10, color:'var(--text-2)' }}>{tl.days[i]}</div>
                  <div>
                    {isToday
                      ? <span style={{ display:'inline-flex', alignItems:'center', justifyContent:'center',
                          width:22, height:22, borderRadius:'50%', background:'var(--purple)', color:'#fff',
                          fontSize:12, fontWeight:500 }}>{d.getDate()}</span>
                      : <span style={{ fontSize:14, fontWeight:500 }}>{d.getDate()}</span>
                    }
                  </div>
                </div>
              )
            })}
          </div>

          {/* time + grid */}
          <div style={{ flex:1, overflow:'auto', display:'flex' }}>
            <TimeColumn />
            <div style={{ flex:1, display:'flex', height:TH }}>
              {[0,1,2,3,4].map(i => (
                <DayColumn key={i} dayIdx={i}
                  meetings={visibleMeetings().filter(m => m.day_index === i)}
                  visitors={visitors} lang={lang} tl={tl}
                  stLabel={stLabel} execDisplay={execDisplay} vById={vById}
                  onClickMeeting={id => setPanel({ type:'detail', id })}
                  onClickEmpty={(di, sh) => setPanel({ type:'add', dayIdx:di, startH:sh })} />
              ))}
            </div>
          </div>
        </div>

        {/* ── right panel ── */}
        {panel && (
          <div style={styles.rp}>
            <div style={styles.rpi}>
              {panel.type === 'detail' && (
                <DetailPanel
                  meeting={meetings.find(m => m.id === panel.id)}
                  visitors={visitors} lang={lang} tl={tl}
                  stLabel={stLabel} execDisplay={execDisplay} vById={vById} vName={vName}
                  weekDates={weekDates}
                  onClose={() => setPanel(null)}
                  onEdit={() => setPanel({ type:'edit', id:panel.id })}
                  onICS={handleICS}
                />
              )}
              {(panel.type === 'add' || panel.type === 'edit') && (
                <MeetingForm
                  meeting={panel.type === 'edit' ? meetings.find(m => m.id === panel.id) : null}
                  visitors={visitors} lang={lang} tl={tl}
                  initialDay={panel.dayIdx ?? 0}
                  initialStart={panel.startH ?? 10}
                  onSave={panel.type === 'edit'
                    ? (data) => updateMeeting(panel.id, data)
                    : addMeeting}
                  onDelete={panel.type === 'edit' ? () => deleteMeeting(panel.id) : null}
                  onClose={() => setPanel(null)}
                  vName={vName}
                />
              )}
              {panel.type === 'visitors' && (
                <VisitorPanel
                  visitors={visitors} lang={lang} tl={tl}
                  meetings={meetings}
                  onAdd={addVisitor}
                  onDelete={deleteVisitor}
                  onClose={() => setPanel(null)}
                  vName={vName}
                />
              )}
            </div>
          </div>
        )}
      </div>

      {/* ── legend ── */}
      <div style={styles.legend}>
        {[['confirmed','conf','#EAF3DE','#639922'],['tentative','tent','#FAEEDA','#BA7517'],['travel','trvl','#E6F1FB','#185FA5']].map(([s,lk,bg,bd]) => (
          <span key={s} style={{ display:'flex', alignItems:'center', gap:4, fontSize:10, color:'var(--text-2)' }}>
            <span style={{ width:10, height:10, borderRadius:2, background:bg, borderLeft:`3px solid ${bd}` }} />
            {tl[lk]}
          </span>
        ))}
        <span style={{ marginLeft:'auto', fontSize:10, color:'var(--text-3)' }}>{tl.hint}</span>
      </div>

      {loading && <div className="loading">{tl.loading}</div>}
    </div>
  )
}

// ── TimeColumn ────────────────────────────────────────────────────────────────
function TimeColumn() {
  return (
    <div style={{ width:48, flexShrink:0, position:'relative', height:TH,
      borderRight:'0.5px solid var(--border)' }}>
      {Array.from({ length: EH - SH + 1 }, (_, i) => SH + i).map(h => (
        <div key={h} style={{ position:'absolute', right:5, top:h2y(h),
          transform:'translateY(-50%)', fontSize:9, color:'var(--text-3)', whiteSpace:'nowrap', lineHeight:1 }}>
          {h}:00
        </div>
      ))}
    </div>
  )
}

// ── DayColumn ─────────────────────────────────────────────────────────────────
function DayColumn({ dayIdx, meetings, visitors, lang, tl, stLabel, execDisplay, vById, onClickMeeting, onClickEmpty }) {
  const calBdRef = useRef(null)

  function handleClick(e) {
    if (e.target !== e.currentTarget) return
    const parent = document.getElementById('cal-scroll')
    const rect = e.currentTarget.getBoundingClientRect()
    const scrollTop = parent ? parent.scrollTop : 0
    const relY = e.clientY - rect.top
    const rh = Math.round((SH + (relY + (parent?.scrollTop ?? 0) - (parent?.getBoundingClientRect().top ?? 0) + rect.top - (parent?.getBoundingClientRect().top ?? 0)) / HPX * 2)) / 2
    // simpler: use offsetY
    const oh = e.nativeEvent.offsetY
    const h = Math.max(SH, Math.min(EH - 0.5, Math.round((SH + oh / HPX) * 2) / 2))
    onClickEmpty(dayIdx, h)
  }

  return (
    <div onClick={handleClick}
      style={{ flex:1, position:'relative', height:TH,
        borderRight: dayIdx < 4 ? '0.5px solid var(--border)' : 'none',
        cursor:'crosshair',
        backgroundImage:'repeating-linear-gradient(to bottom,transparent 0,transparent 19px,rgba(128,128,128,.08) 19px,rgba(128,128,128,.08) 20px,transparent 20px,transparent 39px,rgba(128,128,128,.2) 39px,rgba(128,128,128,.2) 40px)',
      }}>
      {meetings.map(m => (
        <MeetingBlock key={m.id} meeting={m} lang={lang} tl={tl}
          stLabel={stLabel} execDisplay={execDisplay} vById={vById}
          onClick={() => onClickMeeting(m.id)} />
      ))}
    </div>
  )
}

// ── MeetingBlock ──────────────────────────────────────────────────────────────
function MeetingBlock({ meeting: m, lang, tl, stLabel, execDisplay, vById, onClick }) {
  const c   = STC[m.status]
  const ow  = OWC[m.owner] || OWC['TY']
  const top = h2y(m.start_time)
  const hpx = Math.max(h2y(m.end_time) - top, 18)
  const ttl = lang === 'en' && m.title_en ? m.title_en : m.title
  const ex  = execDisplay(m)
  const v   = m.visitor_id ? vById(m.visitor_id) : null
  const vc  = v ? vcol(v.color_idx) : null

  return (
    <div onClick={e => { e.stopPropagation(); onClick() }}
      style={{ position:'absolute', left:3, right:3, top, height:hpx,
        borderRadius:4, padding:'3px 5px', cursor:'pointer', overflow:'hidden', zIndex:2,
        background:c.bg, borderLeft:`3px solid ${c.bd}`, color:c.tx,
        transition:'filter .12s' }}
      onMouseEnter={e => e.currentTarget.style.filter='brightness(.9)'}
      onMouseLeave={e => e.currentTarget.style.filter=''}>
      <div style={{ fontSize:10, fontWeight:500, lineHeight:1.3, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{ttl}</div>
      {hpx > 32 && ex && <div style={{ fontSize:9, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap', marginTop:1, opacity:.75 }}>{ex}</div>}
      {hpx > 54 && (
        <div style={{ display:'flex', gap:3, marginTop:2, alignItems:'center' }}>
          <span style={{ fontSize:8, padding:'1px 5px', borderRadius:6, fontWeight:500, background:'rgba(0,0,0,.1)', color:c.tx }}>{stLabel(m.status)}</span>
          <span className="dot" style={{ background:ow.bg, color:ow.col, width:13, height:13, fontSize:7 }}>{m.owner}</span>
          {vc && <span className="vdot" style={{ background:vc.bg, color:vc.col, borderColor:vc.bd, width:13, height:13, fontSize:8 }}>{v.name.charAt(0)}</span>}
        </div>
      )}
    </div>
  )
}

// ── DetailPanel ───────────────────────────────────────────────────────────────
function DetailPanel({ meeting: m, visitors, lang, tl, stLabel, execDisplay, vById, vName, weekDates, onClose, onEdit, onICS }) {
  if (!m) return null
  const c  = STC[m.status]
  const ow = OWC[m.owner] || OWC['TY']
  const v  = m.visitor_id ? vById(m.visitor_id) : null
  const vc = v ? vcol(v.color_idx) : null
  const ttl = lang === 'en' && m.title_en ? m.title_en : m.title
  const ex  = execDisplay(m)
  const d   = weekDates[m.day_index]
  const dateStr = d ? `${d.getMonth()+1}/${d.getDate()} ` : ''

  return <>
    <PanelHeader title={ttl} onClose={onClose} />
    <div style={{ display:'flex', alignItems:'center', gap:6, flexWrap:'wrap' }}>
      <span style={{ display:'inline-flex', padding:'2px 9px', borderRadius:10, fontSize:10, fontWeight:500,
        background:c.bg, color:c.tx, border:`0.5px solid ${c.bd}` }}>{stLabel(m.status)}</span>
      <span className="dot" style={{ background:ow.bg, color:ow.col }}>{m.owner}</span>
      <span style={{ fontSize:10, color:'var(--text-2)' }}>{tl.ownN[m.owner]}</span>
      {vc && <><span className="vdot" style={{ background:vc.bg, color:vc.col, borderColor:vc.bd }}>{v.name.charAt(0)}</span>
        <span style={{ fontSize:10, color:'var(--text-2)' }}>{vName(v)} ({v.role})</span></>}
      {m.visitor_scope === 'all' && <span style={{ fontSize:10, color:'var(--text-2)' }}>{tl.allExec}</span>}
    </div>
    <FieldRow label={tl.time} value={`${dateStr}${ft(m.start_time)} – ${ft(m.end_time)}`} />
    {(m.attendees||[]).length > 0 && (
      <div><div style={styles.label}>{tl.atts}</div>
        <div style={{ display:'flex', flexWrap:'wrap', gap:3, marginTop:2 }}>
          {(m.attendees||[]).map(a => <span key={a} style={styles.chip}>{a}</span>)}
        </div>
      </div>
    )}
    {m.notes && <FieldRow label={tl.notes} value={m.notes} muted />}
    {m.status === 'confirmed' && m.briefing && <>
      <hr style={{ border:'none', borderTop:'0.5px solid var(--border)', margin:'2px 0' }} />
      <div>
        <div style={styles.label}>📋 {tl.brf}</div>
        <div style={{ background:'var(--bg-2)', border:'0.5px solid var(--border)', borderRadius:6,
          padding:9, fontSize:11, lineHeight:1.65, marginTop:4 }}>{m.briefing}</div>
      </div>
    </>}
    <div style={{ display:'flex', flexDirection:'column', gap:5, marginTop:'auto', paddingTop:4 }}>
      <button className="btn btn-teams" onClick={() => onICS(m)} style={{ justifyContent:'center' }}>
        📅 {tl.ics}
      </button>
      <button className="btn" onClick={onEdit} style={{ justifyContent:'center', fontSize:10 }}>
        ✏️ {lang==='ja'?'編集':'Edit'}
      </button>
    </div>
  </>
}

// ── MeetingForm ───────────────────────────────────────────────────────────────
function MeetingForm({ meeting, visitors, lang, tl, initialDay, initialStart, onSave, onDelete, onClose, vName }) {
  const isEdit = !!meeting
  const fmtT = h => { const hr=Math.floor(h), mn=Math.round((h%1)*60); return `${pad2(hr)}:${pad2(mn)}` }
  const [form, setForm] = useState({
    title:        meeting?.title || '',
    title_en:     meeting?.title_en || '',
    day_index:    meeting?.day_index ?? initialDay,
    start_time:   meeting?.start_time ?? initialStart,
    end_time:     meeting?.end_time ?? Math.min(EH, (meeting?.start_time ?? initialStart) + 1),
    status:       meeting?.status || 'tentative',
    owner:        meeting?.owner || 'TY',
    visitor_id:   meeting?.visitor_id || '',
    visitor_scope: meeting?.visitor_scope || '',
    custom_exec:  (!meeting?.visitor_id && meeting?.visitor_scope !== 'all') ? '' : '',
    attendees:    (meeting?.attendees || []).join(', '),
    notes:        meeting?.notes || '',
    briefing:     meeting?.briefing || '',
  })

  const set = (k, v) => setForm(f => ({ ...f, [k]: v }))

  function parseTime(s) { const [h,m] = s.split(':').map(Number); return h + m/60 }

  async function handleSave() {
    const vid = form.visitor_id === 'all' || form.visitor_id === 'custom' ? null : form.visitor_id
    const vsc = form.visitor_id === 'all' ? 'all' : null
    const atts = form.attendees.split(',').map(s=>s.trim()).filter(Boolean)
    await onSave({
      title: form.title || (lang==='ja'?'(無題)':'(Untitled)'),
      title_en: form.title_en,
      day_index: Number(form.day_index),
      start_time: form.start_time,
      end_time:   form.end_time,
      status:     form.status,
      owner:      form.owner,
      visitor_id: vid || null,
      visitor_scope: vsc,
      attendees:  atts,
      notes:      form.notes,
      briefing:   form.briefing,
    })
  }

  return <>
    <PanelHeader title={isEdit ? tl.editT : tl.addT}
      onClose={onClose} backLabel={isEdit ? '←' : null} />
    <div className="fld"><label>{tl.fttl}</label>
      <input value={form.title} onChange={e=>set('title',e.target.value)} placeholder={lang==='ja'?'例：KHI商談':'e.g. KHI Meeting'} /></div>
    <div className="fld"><label>{tl.fattl}</label>
      <input value={form.title_en} onChange={e=>set('title_en',e.target.value)} /></div>
    <div className="fld"><label>{tl.fday}</label>
      <select value={form.day_index} onChange={e=>set('day_index',e.target.value)}>
        {tl.days.map((d,i) => <option key={i} value={i}>{d}</option>)}
      </select></div>
    <div className="f2">
      <div className="fld"><label>{tl.fst}</label>
        <input type="time" value={fmtT(form.start_time)} onChange={e=>set('start_time',parseTime(e.target.value))} /></div>
      <div className="fld"><label>{tl.fen}</label>
        <input type="time" value={fmtT(form.end_time)} onChange={e=>set('end_time',parseTime(e.target.value))} /></div>
    </div>
    <div className="fld"><label>{tl.fstat}</label>
      <select value={form.status} onChange={e=>set('status',e.target.value)}>
        <option value="confirmed">{tl.conf}</option>
        <option value="tentative">{tl.tent}</option>
        <option value="travel">{tl.trvl}</option>
      </select></div>
    <div className="fld"><label>{tl.fow}</label>
      <select value={form.owner} onChange={e=>set('owner',e.target.value)}>
        {Object.entries(tl.ownN).map(([k,v]) => <option key={k} value={k}>{k} – {v}</option>)}
      </select></div>
    <div className="fld"><label>{tl.fexec}</label>
      <select value={form.visitor_id} onChange={e=>set('visitor_id',e.target.value)}>
        <option value="">{tl.noExec}</option>
        {visitors.map(v => <option key={v.id} value={v.id}>{vName(v)} ({v.role})</option>)}
        <option value="all">{tl.allExec}</option>
      </select></div>
    <div className="fld"><label>{tl.fatts}</label>
      <input value={form.attendees} onChange={e=>set('attendees',e.target.value)} placeholder="IFS.TY, Client.Tanaka" /></div>
    <div className="fld"><label>{tl.fnotes}</label>
      <textarea value={form.notes} onChange={e=>set('notes',e.target.value)} style={{ height:44 }} /></div>
    <div className="fld"><label>{tl.fbrf}</label>
      <textarea value={form.briefing} onChange={e=>set('briefing',e.target.value)} style={{ height:60 }} /></div>
    <div className="f2" style={{ marginTop:4 }}>
      {onDelete
        ? <button className="btn btn-danger" onClick={onDelete}>{tl.del}</button>
        : <button className="btn" onClick={onClose}>{tl.cancel}</button>}
      <button className="btn btn-p" onClick={handleSave}>{tl.save}</button>
    </div>
  </>
}

// ── VisitorPanel ──────────────────────────────────────────────────────────────
function VisitorPanel({ visitors, lang, tl, meetings, onAdd, onDelete, onClose, vName }) {
  const [selColor, setSelColor] = useState(0)
  const [nm, setNm]   = useState('')
  const [nmEn, setNmEn] = useState('')
  const [role, setRole] = useState('')

  async function handleAdd() {
    if (!nm.trim() && !nmEn.trim()) return
    await onAdd({ name: nm.trim() || nmEn.trim(), name_en: nmEn.trim() || nm.trim(), role: role.trim(), color_idx: selColor })
    setNm(''); setNmEn(''); setRole('')
  }

  return <>
    <PanelHeader title={tl.visSetup} onClose={onClose}
      subtitle={tl.visDesc} />
    {visitors.length === 0
      ? <div style={{ textAlign:'center', padding:'16px 0', fontSize:11, color:'var(--text-3)' }}>{tl.visNone}</div>
      : visitors.map(v => {
          const c = vcol(v.color_idx)
          const cnt = meetings.filter(m => m.visitor_id === v.id).length
          return (
            <div key={v.id} style={{ display:'flex', alignItems:'center', gap:8, padding:'7px 9px',
              borderRadius:6, border:'0.5px solid var(--border)', background:'var(--bg-2)' }}>
              <span className="vdot" style={{ background:c.bg, color:c.col, borderColor:c.bd, width:24, height:24, fontSize:11 }}>{v.name.charAt(0)}</span>
              <div style={{ flex:1, minWidth:0 }}>
                <div style={{ fontSize:12, fontWeight:500, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{vName(v)}</div>
                <div style={{ fontSize:9, color:'var(--text-2)' }}>{v.role} · {cnt}{lang==='ja'?'件':' mtgs'}</div>
              </div>
              <button className="btn btn-icon" onClick={() => onDelete(v.id)} style={{ fontSize:11 }}>×</button>
            </div>
          )
        })
    }
    <hr style={{ border:'none', borderTop:'0.5px solid var(--border)', margin:'2px 0' }} />
    <div style={{ fontSize:11, fontWeight:500, marginBottom:2 }}>{tl.visAdd}</div>
    <div className="fld"><label>{tl.visName}</label>
      <input value={nm} onChange={e=>setNm(e.target.value)} placeholder={lang==='ja'?'例：山田太郎':'e.g. Yamada'} /></div>
    <div className="fld"><label>{tl.visNameEn}</label>
      <input value={nmEn} onChange={e=>setNmEn(e.target.value)} placeholder="e.g. Alex Smith" /></div>
    <div className="fld"><label>{tl.visRole}</label>
      <input value={role} onChange={e=>setRole(e.target.value)} placeholder={lang==='ja'?'例：CCO':'e.g. CCO'} /></div>
    <div className="fld">
      <label>{tl.visColor}</label>
      <div style={{ display:'flex', gap:6, padding:'4px 0' }}>
        {VCOLS.map((c,i) => (
          <span key={i} onClick={() => setSelColor(i)}
            style={{ width:22, height:22, borderRadius:'50%', cursor:'pointer',
              background:c.bg, border:`2px solid ${c.bd}`,
              transform: selColor===i ? 'scale(1.3)' : 'scale(1)',
              boxShadow: selColor===i ? `0 0 0 2px #fff, 0 0 0 3.5px ${c.bd}` : 'none',
              transition:'transform .12s, box-shadow .12s' }} />
        ))}
      </div>
    </div>
    <button className="btn btn-p" onClick={handleAdd} style={{ width:'100%', justifyContent:'center' }}>
      + {tl.visAddBtn ?? tl.visAdd}
    </button>
  </>
}

// ── small shared components ───────────────────────────────────────────────────
function PanelHeader({ title, subtitle, onClose }) {
  return (
    <div style={{ display:'flex', alignItems:'flex-start', justifyContent:'space-between', gap:6 }}>
      <div>
        <div style={{ fontSize:13, fontWeight:500, lineHeight:1.4 }}>{title}</div>
        {subtitle && <div style={{ fontSize:10, color:'var(--text-2)', marginTop:2 }}>{subtitle}</div>}
      </div>
      <button className="btn btn-icon" onClick={onClose} style={{ flexShrink:0 }}>×</button>
    </div>
  )
}

function FieldRow({ label, value, muted }) {
  return (
    <div>
      <div style={styles.label}>{label}</div>
      <div style={{ fontSize:11, color: muted ? 'var(--text-2)' : 'var(--text)', lineHeight:1.5, marginTop:2 }}>{value}</div>
    </div>
  )
}

// ── inline styles ─────────────────────────────────────────────────────────────
const styles = {
  topbar: { background:'var(--bg)', borderBottom:'0.5px solid var(--border)', padding:'8px 12px',
    display:'flex', alignItems:'center', justifyContent:'space-between', gap:8, flexShrink:0 },
  logo: { background:'#3B1A6E', color:'#fff', fontSize:11, fontWeight:500,
    padding:'3px 9px', borderRadius:6, letterSpacing:.3 },
  appTitle: { fontSize:13, fontWeight:500 },
  bar: { background:'var(--bg)', borderBottom:'0.5px solid var(--border)', padding:'5px 12px',
    display:'flex', alignItems:'center', gap:7, flexWrap:'wrap', minHeight:34, flexShrink:0 },
  barGroup: { display:'flex', alignItems:'center', gap:5 },
  barLabel: { fontSize:11, color:'var(--text-2)', whiteSpace:'nowrap' },
  rp: { width:268, flexShrink:0, borderLeft:'0.5px solid var(--border)',
    background:'var(--bg)', overflowY:'auto', display:'flex', flexDirection:'column' },
  rpi: { padding:13, display:'flex', flexDirection:'column', gap:9, minHeight:'100%' },
  legend: { display:'flex', gap:10, alignItems:'center', padding:'5px 12px',
    borderTop:'0.5px solid var(--border)', background:'var(--bg)', flexWrap:'wrap', flexShrink:0 },
  label: { fontSize:9, color:'var(--text-2)', textTransform:'uppercase', letterSpacing:.5 },
  chip: { fontSize:9, padding:'2px 6px', borderRadius:8, background:'var(--bg-2)',
    color:'var(--text-2)', border:'0.5px solid var(--border)' },
}
