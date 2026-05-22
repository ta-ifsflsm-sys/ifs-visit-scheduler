import { useState, useEffect, useCallback, useRef } from 'react'
import { supabase } from './lib/supabase'
import { downloadICS } from './utils/ics'
import * as XLSX from 'xlsx'

// ── constants ─────────────────────────────────────────────────────────────────
const SH = 8, EH = 20, HPX = 80, TH = (EH - SH) * HPX
const SL = HPX / 2

const VCOLS = [
  { bg: '#FAECE7', col: '#993C1D', bd: '#D85A30' },
  { bg: '#E6F1FB', col: '#185FA5', bd: '#378ADD' },
  { bg: '#FAEEDA', col: '#854F0B', bd: '#BA7517' },
  { bg: '#E1F5EE', col: '#0F6E56', bd: '#1D9E75' },
  { bg: '#EEEDFE', col: '#534AB7', bd: '#7F77DD' },
  { bg: '#FBEAF0', col: '#993556', bd: '#D4537E' },
]

const STC = {
  confirmed: { bg: '#EAF3DE', bd: '#639922', tx: '#27500A' },
  tentative:  { bg: '#FAEEDA', bd: '#BA7517', tx: '#633806' },
  travel:     { bg: '#E6F1FB', bd: '#185FA5', tx: '#0C447C' },
}

const T = {
  ja: {
    days: ['月','火','水','木','金'],
    wk:'週の開始（月）:', ow:'担当者:', st:'状態:', vis:'来訪者:',
    all:'全員', allst:'全て', allvis:'全来訪者',
    conf:'確定', tent:'仮予約', trvl:'移動', add:'新規ミーティング',
    exec:'担当来訪者', time:'時間', atts:'参加者', notes:'メモ',
    brf:'ブリーフィング', ics:'カレンダーに追加 (.ics)',
    save:'保存', cancel:'キャンセル', del:'削除',
    addT:'ミーティングを追加', editT:'ミーティングを編集',
    fttl:'タイトル (JP)', fattl:'タイトル (EN)',
    fday:'曜日', fst:'開始', fen:'終了',
    fstat:'状態', fow:'担当者', fexec:'来訪者',
    fatts:'参加者（カンマ区切り）', fnotes:'メモ', fbrf:'ブリーフィング（確定のみ）',
    visSetup:'来訪者設定', visDesc:'この週の来日エグゼクティブ',
    visAdd:'来訪者を追加', visName:'氏名 (JP)', visNameEn:'氏名 (EN)',
    visRole:'役職', visColor:'カラー', visNone:'来訪者未登録',
    noExec:'── なし ──', allExec:'全来訪者',
    owSetup:'担当者設定', owDesc:'チームメンバーを管理',
    owAdd:'担当者を追加', owInit:'イニシャル（最大3文字）',
    owName:'氏名 (JP)', owNameEn:'氏名 (EN)', owColor:'カラー',
    owNone:'担当者未登録',
    owInitHint:'例: TY、SK（英大文字3文字以内）',
    owDupErr:'このイニシャルは既に使用されています',
    owDelConfirm:'この担当者を削除しますか？（関連ミーティングの担当者欄は空になります）',
    hint:'ブロックをクリックで詳細 ／ 空白クリックで追加',
    loading:'読み込み中...', saving:'保存中...',
    del_confirm:'このミーティングを削除しますか？',
    vis_del_confirm:'この来訪者を削除しますか？（関連ミーティングの紐付けは解除されます）',
    xlsImport:'Excelからインポート', xlsTemplate:'テンプレートをダウンロード',
    xlsPreview:'インポートプレビュー', xlsConfirm:'インポート実行',
    xlsCancel:'キャンセル', xlsError:'ファイルの読み込みに失敗しました',
    xlsVisFormat:'列順: 氏名(JP) | 氏名(EN) | 役職 | カラー番号(0〜5)',
    xlsOwFormat:'列順: イニシャル | 氏名(JP) | 氏名(EN) | カラー番号(0〜5)',
  },
  en: {
    days: ['Mon','Tue','Wed','Thu','Fri'],
    wk:'Week start (Mon):', ow:'Owner:', st:'Status:', vis:'Visitors:',
    all:'All', allst:'All', allvis:'All Visitors',
    conf:'Confirmed', tent:'Tentative', trvl:'Travel', add:'Add Meeting',
    exec:'Executive', time:'Time', atts:'Attendees', notes:'Notes',
    brf:'Briefing', ics:'Add to Calendar (.ics)',
    save:'Save', cancel:'Cancel', del:'Delete',
    addT:'Add Meeting', editT:'Edit Meeting',
    fttl:'Title (JP)', fattl:'Title (EN)',
    fday:'Day', fst:'Start', fen:'End',
    fstat:'Status', fow:'Owner', fexec:'Visitor',
    fatts:'Attendees (comma-separated)', fnotes:'Notes', fbrf:'Briefing (confirmed only)',
    visSetup:'Visitor Setup', visDesc:'Executives visiting this week',
    visAdd:'Add Visitor', visName:'Name (JP)', visNameEn:'Name (EN)',
    visRole:'Role/Title', visColor:'Color', visNone:'No visitors registered',
    noExec:'── None ──', allExec:'All Visitors',
    owSetup:'Owner Setup', owDesc:'Manage team members',
    owAdd:'Add Owner', owInit:'Initials (max 3 chars)',
    owName:'Name (JP)', owNameEn:'Name (EN)', owColor:'Color',
    owNone:'No owners registered',
    owInitHint:'e.g. TY, SK (uppercase, max 3 chars)',
    owDupErr:'This initial is already in use',
    owDelConfirm:'Delete this owner? (linked meetings will have owner cleared)',
    hint:'Click block for details  •  Click empty space to add',
    loading:'Loading...', saving:'Saving...',
    del_confirm:'Delete this meeting?',
    vis_del_confirm:'Remove this visitor? (linked meetings will be unlinked)',
    xlsImport:'Import from Excel', xlsTemplate:'Download Template',
    xlsPreview:'Import Preview', xlsConfirm:'Confirm Import',
    xlsCancel:'Cancel', xlsError:'Failed to read file',
    xlsVisFormat:'Columns: Name(JP) | Name(EN) | Role | Color(0–5)',
    xlsOwFormat:'Columns: Initials | Name(JP) | Name(EN) | Color(0–5)',
  },
}

// ── helpers ───────────────────────────────────────────────────────────────────
const h2y     = h => (h - SH) * HPX
const ft      = h => { const hr=Math.floor(h),mn=Math.round((h%1)*60); return `${hr}:${String(mn).padStart(2,'0')}` }
const pad2    = n => String(n).padStart(2,'0')
const fmtDate = d => `${d.getFullYear()}-${pad2(d.getMonth()+1)}-${pad2(d.getDate())}`
const vcol    = ci => VCOLS[ci % VCOLS.length]

function getMondayOf(d = new Date()) {
  const day = d.getDay(), diff = day===0 ? -6 : 1-day
  const mon = new Date(d); mon.setDate(d.getDate()+diff); mon.setHours(0,0,0,0); return mon
}
function getWeekDates(monday) {
  return Array.from({length:5}, (_,i) => { const d=new Date(monday); d.setDate(monday.getDate()+i); return d })
}

// ── Excel helpers ─────────────────────────────────────────────────────────────
function parseExcel(file) {
  return new Promise((res,rej) => {
    const r = new FileReader()
    r.onload = e => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), {type:'array'})
        const ws = wb.Sheets[wb.SheetNames[0]]
        res(XLSX.utils.sheet_to_json(ws, {header:1, defval:''}))
      } catch(e) { rej(e) }
    }
    r.onerror = rej; r.readAsArrayBuffer(file)
  })
}
function dlTemplate(data, filename) {
  const wb = XLSX.utils.book_new()
  const ws = XLSX.utils.aoa_to_sheet(data)
  ws['!cols'] = data[0].map(() => ({wch:18}))
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1')
  XLSX.writeFile(wb, filename)
}

// ── App ───────────────────────────────────────────────────────────────────────
export default function App() {
  const [lang,     setLangState]  = useState('ja')
  const [meetings, setMeetings]   = useState([])
  const [visitors, setVisitors]   = useState([])
  const [owners,   setOwners]     = useState([])
  const [monday,   setMonday]     = useState(() => {
    const p = new URLSearchParams(window.location.search).get('week')
    if (p) { const d=new Date(p); if (!isNaN(d)) return d }
    return getMondayOf()
  })
  const [foOw,  setFoOw]   = useState('all')
  const [foSt,  setFoSt]   = useState('all')
  const [foVid, setFoVid]  = useState('all')
  const [owBarOpen,  setOwBarOpen]  = useState(true)
  const [visBarOpen, setVisBarOpen] = useState(false)
  const [panel, setPanel]  = useState(null)
  const [saving, setSaving] = useState(false)
  const [loading, setLoading] = useState(true)
  const tl = T[lang]

  // URL sync
  useEffect(() => {
    const url = new URL(window.location)
    url.searchParams.set('week', fmtDate(monday))
    window.history.replaceState({}, '', url)
  }, [monday])

  // Fetch
  const fetchData = useCallback(async () => {
    const ws = fmtDate(monday)
    const [{ data:m }, { data:v }, { data:o }] = await Promise.all([
      supabase.from('meetings').select('*').eq('week_start',ws).order('day_index').order('start_time'),
      supabase.from('visitors').select('*').eq('week_start',ws).order('created_at'),
      supabase.from('owners').select('*').order('created_at'),
    ])
    setMeetings(m||[]); setVisitors(v||[]); setOwners(o||[])
    setLoading(false)
  }, [monday])

  useEffect(() => { setLoading(true); fetchData() }, [fetchData])

  // Realtime
  useEffect(() => {
    const ws = fmtDate(monday)
    const ch = supabase.channel(`scheduler-${ws}`)
      .on('postgres_changes',{event:'*',schema:'public',table:'meetings', filter:`week_start=eq.${ws}`}, fetchData)
      .on('postgres_changes',{event:'*',schema:'public',table:'visitors', filter:`week_start=eq.${ws}`}, fetchData)
      .on('postgres_changes',{event:'*',schema:'public',table:'owners'},  fetchData)
      .subscribe()
    return () => supabase.removeChannel(ch)
  }, [monday, fetchData])

  // ── CRUD: meetings ──
  async function addMeeting(data)        { setSaving(true); await supabase.from('meetings').insert({...data, week_start:fmtDate(monday)}); setSaving(false); setPanel(null) }
  async function updateMeeting(id, data) { setSaving(true); await supabase.from('meetings').update(data).eq('id',id); setSaving(false); setPanel(null) }
  async function deleteMeeting(id)       { if (!window.confirm(tl.del_confirm)) return; setSaving(true); await supabase.from('meetings').delete().eq('id',id); setSaving(false); setPanel(null) }

  // ── CRUD: visitors ──
  async function addVisitor(data)      { setSaving(true); await supabase.from('visitors').insert({...data, week_start:fmtDate(monday)}); setSaving(false) }
  async function bulkAddVisitors(list) { setSaving(true); await supabase.from('visitors').insert(list.map(v=>({...v, week_start:fmtDate(monday)}))); setSaving(false) }
  async function deleteVisitor(id)     { if (!window.confirm(tl.vis_del_confirm)) return; setSaving(true); await supabase.from('visitors').delete().eq('id',id); if (foVid===id) setFoVid('all'); setSaving(false) }

  // ── CRUD: owners ──
  async function addOwner(data) {
    const dup = owners.find(o => o.id.toUpperCase() === data.id.toUpperCase())
    if (dup) return tl.owDupErr
    setSaving(true)
    await supabase.from('owners').insert({...data, id: data.id.toUpperCase()})
    setSaving(false)
    return null
  }
  async function deleteOwner(id) {
    if (!window.confirm(tl.owDelConfirm)) return
    setSaving(true)
    await supabase.from('owners').delete().eq('id', id)
    if (foOw === id) setFoOw('all')
    setSaving(false)
  }
  async function bulkAddOwners(list) {
    setSaving(true)
    const existing = new Set(owners.map(o => o.id.toUpperCase()))
    const toInsert = list
      .filter(o => !existing.has(o.id.toUpperCase()))
      .map(o => ({ ...o, id: o.id.toUpperCase() }))
    if (toInsert.length > 0) await supabase.from('owners').insert(toInsert)
    setSaving(false)
  }

  function changeWeek(val) {
    const [y,m,d] = val.split('-').map(Number)
    setMonday(new Date(y,m-1,d)); setFoVid('all')
  }
  function setLang(l) { setLangState(l) }

  // ── filter ──
  function visibleMeetings() {
    return meetings.filter(m => {
      const owOk  = foOw  === 'all' || m.owner === foOw
      const stOk  = foSt  === 'all' || m.status === foSt
      const vidOk = foVid === 'all' || m.visitor_id === foVid || m.visitor_scope === 'all'
      return owOk && stOk && vidOk
    })
  }

  // ── helpers ──
  const weekDates  = getWeekDates(monday)
  const today      = new Date()
  const vById      = id => visitors.find(v => v.id === id)
  const oById      = id => owners.find(o => o.id === id)
  const oName      = (id) => { const o=oById(id); if (!o) return id; return lang==='en' ? (o.name_en||o.name) : o.name }
  const oColor     = (id) => { const o=oById(id); return o ? vcol(o.color_idx) : {bg:'#F1EFE8',col:'#5F5E5A',bd:'#B4B2A9'} }
  const vName      = v => v ? (lang==='en' ? (v.name_en||v.name) : v.name) : ''
  const stLabel    = s => tl[{confirmed:'conf',tentative:'tent',travel:'trvl'}[s]]

  function execDisplay(m) {
    if (m.visitor_scope==='all') return lang==='ja' ? '全来訪者' : 'All Visitors'
    if (m.visitor_id) { const v=vById(m.visitor_id); if (v) return `${vName(v)} (${v.role})` }
    return ''
  }

  return (
    <div style={{display:'flex',flexDirection:'column',height:'100vh',overflow:'hidden',width:'100%',maxWidth:'100vw'}}>

      {/* topbar */}
      <div style={S.topbar}>
        <div style={{display:'flex',alignItems:'center',gap:8}}>
          <span style={S.logo}>IFS</span>
          <span style={{fontSize:13,fontWeight:500}}>Japan Visit Scheduler</span>
          {saving && <span style={{fontSize:10,color:'#888'}}>{tl.saving}</span>}
        </div>
        <div style={{display:'flex',alignItems:'center',gap:6}}>
          <button className="btn btn-p" onClick={() => setPanel({type:'add',dayIdx:0,startH:9})}>+ {tl.add}</button>
          <div className="lang-sw">
            <button className={lang==='ja'?'active':''} onClick={() => setLang('ja')}>JA</button>
            <button className={lang==='en'?'active':''} onClick={() => setLang('en')}>EN</button>
          </div>
        </div>
      </div>

      {/* owner + status filter bar */}
      <div style={S.bar}>
        <div style={S.bg}>
          <span style={S.bl}>{tl.wk}</span>
          <input type="date" value={fmtDate(monday)} onChange={e=>changeWeek(e.target.value)}
            style={{fontSize:11,padding:'4px 7px',borderRadius:6,border:'0.5px solid var(--border-2)',background:'var(--bg)',color:'var(--text)'}} />
        </div>
        <div style={S.bg}>
          {/* collapsible owner chips */}
          <button onClick={()=>setOwBarOpen(v=>!v)}
            style={{fontSize:10,padding:'2px 5px',border:'none',background:'transparent',cursor:'pointer',color:'var(--text-2)'}}>
            {owBarOpen ? '▼' : '▶'}
          </button>
          <span style={S.bl}>{tl.ow}</span>
          {owBarOpen ? <>
            <span className={`chip${foOw==='all'?' active':''}`} onClick={() => setFoOw('all')}>{tl.all}</span>
            {owners.map(o => {
              const c = vcol(o.color_idx); const isOn = foOw===o.id
              return (
                <span key={o.id} className="chip" onClick={() => setFoOw(o.id)}
                  title={lang==='en'?(o.name_en||o.name):o.name}
                  style={{...(isOn?{background:c.bg,color:c.col,borderColor:c.bd}:{}),maxWidth:120,overflow:'hidden'}}>
                  <span className="dot" style={{background:c.bg,color:c.col,flexShrink:0}}>{o.id}</span>
                  <span style={{overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{(lang==='en'?(o.name_en||o.name):o.name).split(' ')[0]}</span>
                </span>
              )
            })}
          </> : foOw!=='all' && (() => {
            const o=owners.find(x=>x.id===foOw); const c=o?vcol(o.color_idx):null
            return o ? <span className="chip" style={{background:c.bg,color:c.col,borderColor:c.bd}}
              onClick={()=>setOwBarOpen(true)}>
              <span className="dot" style={{background:c.bg,color:c.col}}>{o.id}</span>
              {lang==='en'?(o.name_en||o.name):o.name} ×
            </span> : null
          })()}
          <button className="btn" onClick={() => setPanel({type:'owners'})}
            style={{fontSize:10,padding:'3px 9px'}}>⚙ {tl.owSetup}</button>
        </div>
        <div style={S.bg}>
          <span style={S.bl}>{tl.st}</span>
          {[['all','allst'],['confirmed','conf'],['tentative','tent'],['travel','trvl']].map(([v,lk]) => (
            <span key={v} className={`chip${foSt===v?' active':''}`} onClick={() => setFoSt(v)}>{tl[lk]}</span>
          ))}
        </div>
      </div>

      {/* visitor bar — collapsible */}
      <div style={{...S.bar,background:'var(--bg-2)'}}>
        <button onClick={()=>setVisBarOpen(v=>!v)}
          style={{fontSize:10,padding:'2px 5px',border:'none',background:'transparent',cursor:'pointer',color:'var(--text-2)'}}>
          {visBarOpen ? '▼' : '▶'}
        </button>
        <span style={S.bl}>{tl.vis}</span>
        {visBarOpen ? <>
          <span className={`chip${foVid==='all'?' active':''}`} onClick={() => setFoVid('all')}>{tl.allvis}</span>
          {visitors.map(v => {
            const c=vcol(v.color_idx); const isOn=foVid===v.id
            return (
              <span key={v.id} className="chip" onClick={() => setFoVid(v.id)}
                title={`${vName(v)} (${v.role})`}
                style={{...(isOn?{background:c.bg,color:c.col,borderColor:c.bd}:{}),maxWidth:130,overflow:'hidden'}}>
                <span className="vdot" style={{background:c.bg,color:c.col,borderColor:c.bd,flexShrink:0}}>{v.name.charAt(0)}</span>
                <span style={{overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{vName(v).split(' ')[0]}</span>
              </span>
            )
          })}
        </> : foVid!=='all' && (() => {
          const v=visitors.find(x=>x.id===foVid); const c=v?vcol(v.color_idx):null
          return v ? <span className="chip" style={{background:c.bg,color:c.col,borderColor:c.bd}}
            onClick={()=>setVisBarOpen(true)}>
            <span className="vdot" style={{background:c.bg,color:c.col,borderColor:c.bd}}>{v.name.charAt(0)}</span>
            {vName(v)} ×
          </span> : null
        })()}
        {!visBarOpen && <span style={{fontSize:10,color:'var(--text-3)',marginLeft:4}}>
          ({visitors.length}{lang==='ja'?'名':' visitors'})
        </span>}
        <button className="btn" onClick={() => setPanel({type:'visitors'})}
          style={{marginLeft:'auto',fontSize:10,padding:'3px 9px'}}>⚙ {tl.visSetup}</button>
      </div>

      {/* main */}
      <div style={{flex:1,display:'flex',minHeight:0}}>

        {/* CSS Grid calendar */}
        <div id="cal-scroll" style={{
          flex:1,minWidth:0,overflowY:'auto',
          display:'grid', gridTemplateColumns:'48px repeat(5,1fr)',
          alignContent:'start', background:'var(--bg)',
        }}>
          {/* sticky header */}
          <div style={{position:'sticky',top:0,zIndex:20,height:44,
            background:'var(--bg)',borderBottom:'0.5px solid var(--border)'}} />
          {weekDates.map((d,i) => {
            const isToday = d.toDateString()===today.toDateString()
            return (
              <div key={i} style={{position:'sticky',top:0,zIndex:20,height:44,
                background:'var(--bg)',
                borderLeft:'0.5px solid var(--border)',
                borderBottom:'0.5px solid var(--border)',
                display:'flex',flexDirection:'column',alignItems:'center',justifyContent:'center'}}>
                <div style={{fontSize:10,color:'var(--text-2)'}}>{tl.days[i]}</div>
                {isToday
                  ? <span style={{display:'inline-flex',alignItems:'center',justifyContent:'center',
                      width:22,height:22,borderRadius:'50%',background:'#3B1A6E',color:'#fff',
                      fontSize:12,fontWeight:500}}>{d.getDate()}</span>
                  : <span style={{fontSize:14,fontWeight:500}}>{d.getDate()}</span>
                }
              </div>
            )
          })}

          {/* time column */}
          <div style={{position:'relative',height:TH}}>
            {Array.from({length:EH-SH+1},(_,i) => (
              <div key={i} style={{position:'absolute',right:5,top:i*HPX-8,
                fontSize:9,color:'var(--text-3)',lineHeight:1,
                visibility:i===0?'hidden':'visible'}}>
                {SH+i}:00
              </div>
            ))}
          </div>

          {/* day columns */}
          {[0,1,2,3,4].map(i => (
            <DayColumn key={i} dayIdx={i}
              meetings={visibleMeetings().filter(m=>m.day_index===i)}
              oColor={oColor} lang={lang} tl={tl}
              stLabel={stLabel} execDisplay={execDisplay} vById={vById}
              onClickMeeting={id => setPanel({type:'detail',id})}
              onClickEmpty={(di,sh) => setPanel({type:'add',dayIdx:di,startH:sh})} />
          ))}
        </div>

        {/* right panel */}
        {panel && (
          <div style={S.rp}>
            <div style={S.rpi}>
              {panel.type==='detail' && (
                <DetailPanel
                  meeting={meetings.find(m=>m.id===panel.id)}
                  visitors={visitors} lang={lang} tl={tl}
                  stLabel={stLabel} execDisplay={execDisplay} vById={vById} vName={vName}
                  oName={oName} oColor={oColor} weekDates={weekDates}
                  onClose={() => setPanel(null)}
                  onEdit={() => setPanel({type:'edit',id:panel.id})}
                  onICS={m => downloadICS(m,weekDates,visitors)} />
              )}
              {(panel.type==='add'||panel.type==='edit') && (
                <MeetingForm
                  meeting={panel.type==='edit' ? meetings.find(m=>m.id===panel.id) : null}
                  visitors={visitors} owners={owners} lang={lang} tl={tl}
                  initialDay={panel.dayIdx??0} initialStart={panel.startH??9}
                  onSave={panel.type==='edit' ? data=>updateMeeting(panel.id,data) : addMeeting}
                  onDelete={panel.type==='edit' ? ()=>deleteMeeting(panel.id) : null}
                  onClose={() => setPanel(null)}
                  vName={vName} oName={oName} />
              )}
              {panel.type==='visitors' && (
                <VisitorPanel
                  visitors={visitors} lang={lang} tl={tl} meetings={meetings}
                  onAdd={addVisitor} onBulkAdd={bulkAddVisitors} onDelete={deleteVisitor}
                  onClose={() => setPanel(null)} vName={vName} />
              )}
              {panel.type==='owners' && (
                <OwnerPanel
                  owners={owners} lang={lang} tl={tl} meetings={meetings}
                  onAdd={addOwner} onBulkAdd={bulkAddOwners} onDelete={deleteOwner}
                  onClose={() => setPanel(null)} oName={oName} />
              )}
            </div>
          </div>
        )}
      </div>

      {/* legend */}
      <div style={S.legend}>
        {[['confirmed','conf','#EAF3DE','#639922'],['tentative','tent','#FAEEDA','#BA7517'],['travel','trvl','#E6F1FB','#185FA5']].map(([s,lk,bg,bd]) => (
          <span key={s} style={{display:'flex',alignItems:'center',gap:4,fontSize:10,color:'var(--text-2)'}}>
            <span style={{width:10,height:10,borderRadius:2,background:bg,borderLeft:`3px solid ${bd}`}} />{tl[lk]}
          </span>
        ))}
        <span style={{marginLeft:'auto',fontSize:10,color:'var(--text-3)'}}>{tl.hint}</span>
      </div>

      {loading && <div className="loading">{tl.loading}</div>}
    </div>
  )
}

// ── GridLines ─────────────────────────────────────────────────────────────────
function GridLines() {
  const hours = EH - SH
  return <>
    {Array.from({length:hours*4+1},(_,i) => {
      if (i%2===0) return null
      return <div key={`q${i}`} style={{position:'absolute',left:0,right:0,top:i*20,
        borderTop:'0.5px solid rgba(128,128,128,.06)',pointerEvents:'none'}} />
    })}
    {Array.from({length:hours*2+1},(_,i) => {
      if (i%2===0) return null
      return <div key={`m${i}`} style={{position:'absolute',left:0,right:0,top:i*SL,
        borderTop:'0.5px solid rgba(128,128,128,.14)',pointerEvents:'none'}} />
    })}
    {Array.from({length:hours+1},(_,i) => (
      <div key={`h${i}`} style={{position:'absolute',left:0,right:0,top:i*HPX,
        borderTop:i===0?'none':'0.5px solid rgba(128,128,128,.28)',pointerEvents:'none'}} />
    ))}
  </>
}

// ── DayColumn ─────────────────────────────────────────────────────────────────
function DayColumn({ dayIdx, meetings, oColor, lang, tl, stLabel, execDisplay, vById, onClickMeeting, onClickEmpty }) {
  function handleClick(e) {
    if (e.target!==e.currentTarget) return
    const h = Math.max(SH, Math.min(EH-0.5, Math.round((SH+e.nativeEvent.offsetY/HPX)*2)/2))
    onClickEmpty(dayIdx, h)
  }
  return (
    <div onClick={handleClick} style={{position:'relative',height:TH,
      borderLeft:'0.5px solid var(--border)',cursor:'crosshair'}}>
      <GridLines />
      {meetings.map(m => (
        <MeetingBlock key={m.id} meeting={m} lang={lang} tl={tl}
          stLabel={stLabel} execDisplay={execDisplay} vById={vById} oColor={oColor}
          onClick={() => onClickMeeting(m.id)} />
      ))}
    </div>
  )
}

// ── MeetingBlock ──────────────────────────────────────────────────────────────
function MeetingBlock({ meeting:m, lang, tl, stLabel, execDisplay, vById, oColor, onClick }) {
  const c   = STC[m.status]
  const ow  = oColor(m.owner)
  const top = h2y(m.start_time)
  const hpx = Math.max(h2y(m.end_time)-top, 18)
  const ttl = lang==='en'&&m.title_en ? m.title_en : m.title
  const ex  = execDisplay(m)
  const v   = m.visitor_id ? vById(m.visitor_id) : null
  const vc  = v ? vcol(v.color_idx) : null
  return (
    <div onClick={e=>{e.stopPropagation();onClick()}}
      style={{position:'absolute',left:3,right:3,top,height:hpx,zIndex:2,
        borderRadius:4,padding:'3px 5px',cursor:'pointer',overflow:'hidden',
        background:c.bg,borderLeft:`3px solid ${c.bd}`,color:c.tx,transition:'filter .12s'}}
      onMouseEnter={e=>e.currentTarget.style.filter='brightness(.9)'}
      onMouseLeave={e=>e.currentTarget.style.filter=''}>
      <div style={{fontSize:10,fontWeight:500,lineHeight:1.3,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{ttl}</div>
      {hpx>32&&ex && <div style={{fontSize:9,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',marginTop:1,opacity:.75}}>{ex}</div>}
      {hpx>54 && (
        <div style={{display:'flex',gap:3,marginTop:2,alignItems:'center'}}>
          <span style={{fontSize:8,padding:'1px 5px',borderRadius:6,fontWeight:500,background:'rgba(0,0,0,.1)',color:c.tx}}>{stLabel(m.status)}</span>
          <span className="dot" style={{background:ow.bg,color:ow.col,width:13,height:13,fontSize:7}}>{m.owner}</span>
          {vc && <span className="vdot" style={{background:vc.bg,color:vc.col,borderColor:vc.bd,width:13,height:13,fontSize:8}}>{v.name.charAt(0)}</span>}
        </div>
      )}
    </div>
  )
}

// ── DetailPanel ───────────────────────────────────────────────────────────────
function DetailPanel({ meeting:m, lang, tl, stLabel, execDisplay, vById, vName, oName, oColor, weekDates, onClose, onEdit, onICS }) {
  if (!m) return null
  const c  = STC[m.status]
  const ow = oColor(m.owner)
  const v  = m.visitor_id ? vById(m.visitor_id) : null
  const vc = v ? vcol(v.color_idx) : null
  const ttl = lang==='en'&&m.title_en ? m.title_en : m.title
  const d   = weekDates[m.day_index]
  const ds  = d ? `${d.getMonth()+1}/${d.getDate()} ` : ''
  return <>
    <PanelHeader title={ttl} onClose={onClose} />
    <div style={{display:'flex',alignItems:'center',gap:6,flexWrap:'wrap'}}>
      <span style={{display:'inline-flex',padding:'2px 9px',borderRadius:10,fontSize:10,fontWeight:500,
        background:c.bg,color:c.tx,border:`0.5px solid ${c.bd}`}}>{stLabel(m.status)}</span>
      <span className="dot" style={{background:ow.bg,color:ow.col}}>{m.owner}</span>
      <span style={{fontSize:10,color:'var(--text-2)'}}>{oName(m.owner)}</span>
      {vc && <><span className="vdot" style={{background:vc.bg,color:vc.col,borderColor:vc.bd}}>{v.name.charAt(0)}</span>
        <span style={{fontSize:10,color:'var(--text-2)'}}>{vName(v)} ({v.role})</span></>}
      {m.visitor_scope==='all' && <span style={{fontSize:10,color:'var(--text-2)'}}>{tl.allExec}</span>}
    </div>
    <FieldRow label={tl.time} value={`${ds}${ft(m.start_time)} – ${ft(m.end_time)}`} />
    {(m.attendees||[]).length>0 && (
      <div><div style={S.label}>{tl.atts}</div>
        <div style={{display:'flex',flexWrap:'wrap',gap:3,marginTop:2}}>
          {(m.attendees||[]).map(a=><span key={a} style={S.achip}>{a}</span>)}
        </div>
      </div>
    )}
    {m.notes && <FieldRow label={tl.notes} value={m.notes} muted />}
    {m.status==='confirmed'&&m.briefing && <>
      <hr style={{border:'none',borderTop:'0.5px solid var(--border)',margin:'2px 0'}} />
      <div><div style={S.label}>📋 {tl.brf}</div>
        <div style={{background:'var(--bg-2)',border:'0.5px solid var(--border)',borderRadius:6,
          padding:9,fontSize:11,lineHeight:1.65,marginTop:4}}>{m.briefing}</div>
      </div>
    </>}
    <div style={{display:'flex',flexDirection:'column',gap:5,marginTop:'auto',paddingTop:4}}>
      <button className="btn btn-teams" onClick={() => onICS(m)} style={{justifyContent:'center'}}>📅 {tl.ics}</button>
      <button className="btn" onClick={onEdit} style={{justifyContent:'center',fontSize:10}}>✏️ {lang==='ja'?'編集':'Edit'}</button>
    </div>
  </>
}

// ── MeetingForm ───────────────────────────────────────────────────────────────
function MeetingForm({ meeting, visitors, owners, lang, tl, initialDay, initialStart, onSave, onDelete, onClose, vName, oName }) {
  const isEdit = !!meeting
  const fmtT = h => { const hr=Math.floor(h),mn=Math.round((h%1)*60); return `${pad2(hr)}:${pad2(mn)}` }
  const defaultOwner = owners.length>0 ? owners[0].id : ''
  const [form, setForm] = useState({
    title:         meeting?.title||'',
    title_en:      meeting?.title_en||'',
    day_index:     meeting?.day_index??initialDay,
    start_time:    meeting?.start_time??initialStart,
    end_time:      meeting?.end_time??Math.min(EH,(meeting?.start_time??initialStart)+1),
    status:        meeting?.status||'tentative',
    owner:         meeting?.owner||defaultOwner,
    visitor_id:    meeting?.visitor_id||'',
    visitor_scope: meeting?.visitor_scope||'',
    attendees:     (meeting?.attendees||[]).join(', '),
    notes:         meeting?.notes||'',
    briefing:      meeting?.briefing||'',
  })
  const set = (k,v) => setForm(f=>({...f,[k]:v}))
  const parseTime = s => { const [h,m]=s.split(':').map(Number); return h+m/60 }

  async function handleSave() {
    const vid = (form.visitor_id==='all'||form.visitor_id==='') ? null : form.visitor_id
    const vsc = form.visitor_id==='all' ? 'all' : null
    await onSave({
      title:form.title||(lang==='ja'?'(無題)':'(Untitled)'), title_en:form.title_en,
      day_index:Number(form.day_index), start_time:form.start_time, end_time:form.end_time,
      status:form.status, owner:form.owner,
      visitor_id:vid||null, visitor_scope:vsc,
      attendees:form.attendees.split(',').map(s=>s.trim()).filter(Boolean),
      notes:form.notes, briefing:form.briefing,
    })
  }

  return <>
    <PanelHeader title={isEdit?tl.editT:tl.addT} onClose={onClose} />
    <div className="fld"><label>{tl.fttl}</label><input value={form.title} onChange={e=>set('title',e.target.value)} /></div>
    <div className="fld"><label>{tl.fattl}</label><input value={form.title_en} onChange={e=>set('title_en',e.target.value)} /></div>
    <div className="fld"><label>{tl.fday}</label>
      <select value={form.day_index} onChange={e=>set('day_index',e.target.value)}>
        {tl.days.map((d,i)=><option key={i} value={i}>{d}</option>)}
      </select></div>
    <div className="f2">
      <div className="fld"><label>{tl.fst}</label><input type="time" value={fmtT(form.start_time)} onChange={e=>set('start_time',parseTime(e.target.value))} /></div>
      <div className="fld"><label>{tl.fen}</label><input type="time" value={fmtT(form.end_time)} onChange={e=>set('end_time',parseTime(e.target.value))} /></div>
    </div>
    <div className="fld"><label>{tl.fstat}</label>
      <select value={form.status} onChange={e=>set('status',e.target.value)}>
        <option value="confirmed">{tl.conf}</option>
        <option value="tentative">{tl.tent}</option>
        <option value="travel">{tl.trvl}</option>
      </select></div>
    <div className="fld"><label>{tl.fow}</label>
      <select value={form.owner} onChange={e=>set('owner',e.target.value)}>
        {owners.map(o=><option key={o.id} value={o.id}>{o.id} – {lang==='en'?(o.name_en||o.name):o.name}</option>)}
      </select></div>
    <div className="fld"><label>{tl.fexec}</label>
      <select value={form.visitor_id} onChange={e=>set('visitor_id',e.target.value)}>
        <option value="">{tl.noExec}</option>
        {visitors.map(v=><option key={v.id} value={v.id}>{vName(v)} ({v.role})</option>)}
        <option value="all">{tl.allExec}</option>
      </select></div>
    <div className="fld"><label>{tl.fatts}</label><input value={form.attendees} onChange={e=>set('attendees',e.target.value)} placeholder="IFS.TY, Client.Tanaka" /></div>
    <div className="fld"><label>{tl.fnotes}</label><textarea value={form.notes} onChange={e=>set('notes',e.target.value)} style={{height:44}} /></div>
    <div className="fld"><label>{tl.fbrf}</label><textarea value={form.briefing} onChange={e=>set('briefing',e.target.value)} style={{height:60}} /></div>
    <div className="f2" style={{marginTop:4}}>
      {onDelete
        ? <button className="btn btn-danger" onClick={onDelete}>{tl.del}</button>
        : <button className="btn" onClick={onClose}>{tl.cancel}</button>}
      <button className="btn btn-p" onClick={handleSave}>{tl.save}</button>
    </div>
  </>
}

// ── OwnerPanel ────────────────────────────────────────────────────────────────
function OwnerPanel({ owners, lang, tl, meetings, onAdd, onBulkAdd, onDelete, onClose, oName }) {
  const [selColor, setSelColor] = useState(0)
  const [initials, setInitials] = useState('')
  const [nm, setNm]   = useState('')
  const [nmEn, setNmEn] = useState('')
  const [err, setErr]  = useState('')
  const [xlsPreview, setXlsPreview] = useState(null)
  const [xlsError, setXlsError]     = useState('')
  const fileRef = useRef()

  async function handleAdd() {
    if (!initials.trim()) return
    const id = initials.trim().toUpperCase().replace(/[^A-Z0-9]/g,'').slice(0,3)
    if (!id) return
    const e = await onAdd({ id, name: nm.trim()||id, name_en: nmEn.trim()||nm.trim()||id, color_idx: selColor })
    if (e) { setErr(e); return }
    setInitials(''); setNm(''); setNmEn(''); setErr('')
  }

  async function handleFile(e) {
    const file = e.target.files[0]; if (!file) return
    setXlsError('')
    try {
      const rows = await parseExcel(file)
      const parsed = rows.slice(1).filter(r=>String(r[0]||'').trim()).map((r,i)=>({
        id:   String(r[0]||'').trim().toUpperCase().slice(0,3),
        name: String(r[1]||r[0]||'').trim(),
        name_en: String(r[2]||r[1]||r[0]||'').trim(),
        color_idx: Math.min(5, Math.max(0, parseInt(r[3])||i%6)),
      })).filter(o=>o.id)
      setXlsPreview(parsed)
    } catch { setXlsError(tl.xlsError) }
    e.target.value=''
  }

  function dlTemplate() {
    const rows = lang==='ja'
      ? [['イニシャル','氏名(JP)','氏名(EN)','カラー番号(0〜5)'],['TY','ティアナ','Tiana','0'],['SK','佐藤','Sato','1']]
      : [['Initials','Name(JP)','Name(EN)','Color(0–5)'],['TY','Tiana','Tiana','0'],['SK','Sato','Sato','1']]
    dlTemplate2(rows, 'owners_template.xlsx')
  }
  function dlTemplate2(data, fname) {
    const wb=XLSX.utils.book_new(); const ws=XLSX.utils.aoa_to_sheet(data)
    ws['!cols']=data[0].map(()=>({wch:16})); XLSX.utils.book_append_sheet(wb,ws,'Sheet1'); XLSX.writeFile(wb,fname)
  }

  return <>
    <PanelHeader title={tl.owSetup} subtitle={tl.owDesc} onClose={onClose} />

    {/* Excel import */}
    <div style={{background:'var(--bg-2)',border:'0.5px solid var(--border)',borderRadius:6,padding:9}}>
      <div style={{fontSize:10,fontWeight:500,marginBottom:4}}>📊 {tl.xlsImport}</div>
      <div style={{fontSize:9,color:'var(--text-3)',marginBottom:7}}>{tl.xlsOwFormat}</div>
      <div style={{display:'flex',gap:6}}>
        <button className="btn" style={{fontSize:10,flex:1}} onClick={dlTemplate}>⬇ {tl.xlsTemplate}</button>
        <button className="btn btn-p" style={{fontSize:10,flex:1}} onClick={()=>fileRef.current?.click()}>⬆ {tl.xlsImport}</button>
      </div>
      <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{display:'none'}} onChange={handleFile} />
      {xlsError && <div style={{fontSize:10,color:'#a32d2d',marginTop:5}}>{xlsError}</div>}
    </div>

    {/* preview */}
    {xlsPreview && (
      <div style={{background:'var(--bg-2)',border:'0.5px solid #639922',borderRadius:6,padding:9}}>
        <div style={{fontSize:10,fontWeight:500,marginBottom:6}}>✅ {tl.xlsPreview} ({xlsPreview.length}{lang==='ja'?'件':''})</div>
        {xlsPreview.slice(0,5).map((o,i)=>{
          const c=vcol(o.color_idx)
          return <div key={i} style={{display:'flex',alignItems:'center',gap:6,marginBottom:4}}>
            <span className="dot" style={{background:c.bg,color:c.col,width:20,height:20,fontSize:9}}>{o.id}</span>
            <span style={{fontSize:11}}>{o.name}</span>
            <span style={{fontSize:9,color:'var(--text-2)'}}>{o.name_en}</span>
          </div>
        })}
        {xlsPreview.length>5 && <div style={{fontSize:9,color:'var(--text-3)'}}>... +{xlsPreview.length-5}</div>}
        <div className="f2" style={{marginTop:7}}>
          <button className="btn" style={{justifyContent:'center',fontSize:10}} onClick={()=>setXlsPreview(null)}>{tl.xlsCancel}</button>
          <button className="btn btn-p" style={{justifyContent:'center',fontSize:10}} onClick={async()=>{await onBulkAdd(xlsPreview);setXlsPreview(null)}}>{tl.xlsConfirm}</button>
        </div>
      </div>
    )}

    <hr style={{border:'none',borderTop:'0.5px solid var(--border)',margin:'2px 0'}} />

    {/* current owners */}
    {owners.length===0
      ? <div style={{textAlign:'center',padding:'12px 0',fontSize:11,color:'var(--text-3)'}}>{tl.owNone}</div>
      : owners.map(o=>{
          const c=vcol(o.color_idx); const cnt=meetings.filter(m=>m.owner===o.id).length
          return (
            <div key={o.id} style={{display:'flex',alignItems:'center',gap:8,padding:'7px 9px',
              borderRadius:6,border:'0.5px solid var(--border)',background:'var(--bg-2)'}}>
              <span className="dot" style={{background:c.bg,color:c.col,width:26,height:26,fontSize:10,fontWeight:600,border:`1.5px solid ${c.bd}`}}>{o.id}</span>
              <div style={{flex:1,minWidth:0}}>
                <div style={{fontSize:12,fontWeight:500,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>
                  {lang==='en'?(o.name_en||o.name):o.name}
                </div>
                <div style={{fontSize:9,color:'var(--text-2)'}}>{o.id} · {cnt}{lang==='ja'?'件':' mtgs'}</div>
              </div>
              <button className="btn btn-icon" onClick={()=>onDelete(o.id)} style={{fontSize:11}}>×</button>
            </div>
          )
        })
    }

    <hr style={{border:'none',borderTop:'0.5px solid var(--border)',margin:'2px 0'}} />
    <div style={{fontSize:11,fontWeight:500,marginBottom:2}}>{tl.owAdd}</div>
    <div className="fld"><label>{tl.owInit}</label>
      <input value={initials} onChange={e=>setInitials(e.target.value.toUpperCase().slice(0,3))}
        placeholder={tl.owInitHint} style={{textTransform:'uppercase'}} />
      {err && <span style={{fontSize:10,color:'#a32d2d'}}>{err}</span>}
    </div>
    <div className="fld"><label>{tl.owName}</label><input value={nm} onChange={e=>setNm(e.target.value)} /></div>
    <div className="fld"><label>{tl.owNameEn}</label><input value={nmEn} onChange={e=>setNmEn(e.target.value)} /></div>
    <div className="fld"><label>{tl.owColor}</label>
      <div style={{display:'flex',gap:6,padding:'4px 0'}}>
        {VCOLS.map((c,i)=>(
          <span key={i} onClick={()=>setSelColor(i)}
            style={{width:22,height:22,borderRadius:'50%',cursor:'pointer',
              background:c.bg,border:`2px solid ${c.bd}`,
              transform:selColor===i?'scale(1.3)':'scale(1)',
              boxShadow:selColor===i?`0 0 0 2px var(--bg),0 0 0 3.5px ${c.bd}`:'none',
              transition:'all .12s'}} />
        ))}
      </div>
    </div>
    <button className="btn btn-p" onClick={handleAdd} style={{width:'100%',justifyContent:'center'}}>+ {tl.owAdd}</button>
  </>
}

// ── VisitorPanel ──────────────────────────────────────────────────────────────
function VisitorPanel({ visitors, lang, tl, meetings, onAdd, onBulkAdd, onDelete, onClose, vName }) {
  const [selColor, setSelColor] = useState(0)
  const [nm, setNm]   = useState('')
  const [nmEn, setNmEn] = useState('')
  const [role, setRole] = useState('')
  const [xlsPreview, setXlsPreview] = useState(null)
  const [xlsError, setXlsError]     = useState('')
  const fileRef = useRef()

  async function handleAdd() {
    if (!nm.trim()&&!nmEn.trim()) return
    await onAdd({name:nm.trim()||nmEn.trim(),name_en:nmEn.trim()||nm.trim(),role:role.trim(),color_idx:selColor})
    setNm(''); setNmEn(''); setRole('')
  }
  async function handleFile(e) {
    const file=e.target.files[0]; if (!file) return; setXlsError('')
    try {
      const rows = await parseExcel(file)
      const parsed = rows.slice(1).filter(r=>String(r[0]||r[1]||'').trim()).map((r,i)=>({
        name:    String(r[0]||r[1]||'').trim(),
        name_en: String(r[1]||r[0]||'').trim(),
        role:    String(r[2]||'').trim(),
        color_idx: Math.min(5,Math.max(0,parseInt(r[3])||i%6)),
      }))
      setXlsPreview(parsed)
    } catch { setXlsError(tl.xlsError) }
    e.target.value=''
  }
  function dlTemplate() {
    const wb=XLSX.utils.book_new()
    const data = lang==='ja'
      ? [['氏名(JP)','氏名(EN)','役職','カラー番号(0〜5)'],['Mark Moffat','Mark Moffat','CCO','0']]
      : [['Name(JP)','Name(EN)','Role','Color(0–5)'],['Mark Moffat','Mark Moffat','CCO','0']]
    const ws=XLSX.utils.aoa_to_sheet(data); ws['!cols']=data[0].map(()=>({wch:18}))
    XLSX.utils.book_append_sheet(wb,ws,'Sheet1'); XLSX.writeFile(wb,'visitors_template.xlsx')
  }

  return <>
    <PanelHeader title={tl.visSetup} subtitle={tl.visDesc} onClose={onClose} />
    <div style={{background:'var(--bg-2)',border:'0.5px solid var(--border)',borderRadius:6,padding:9}}>
      <div style={{fontSize:10,fontWeight:500,marginBottom:4}}>📊 {tl.xlsImport}</div>
      <div style={{fontSize:9,color:'var(--text-3)',marginBottom:7}}>{tl.xlsVisFormat}</div>
      <div style={{display:'flex',gap:6}}>
        <button className="btn" style={{fontSize:10,flex:1}} onClick={dlTemplate}>⬇ {tl.xlsTemplate}</button>
        <button className="btn btn-p" style={{fontSize:10,flex:1}} onClick={()=>fileRef.current?.click()}>⬆ {tl.xlsImport}</button>
      </div>
      <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{display:'none'}} onChange={handleFile} />
      {xlsError && <div style={{fontSize:10,color:'#a32d2d',marginTop:5}}>{xlsError}</div>}
    </div>
    {xlsPreview && (
      <div style={{background:'var(--bg-2)',border:'0.5px solid #639922',borderRadius:6,padding:9}}>
        <div style={{fontSize:10,fontWeight:500,marginBottom:6}}>✅ {tl.xlsPreview} ({xlsPreview.length}{lang==='ja'?'件':''})</div>
        {xlsPreview.slice(0,5).map((v,i)=>{const c=vcol(v.color_idx); return (
          <div key={i} style={{display:'flex',alignItems:'center',gap:6,marginBottom:4}}>
            <span className="vdot" style={{background:c.bg,color:c.col,borderColor:c.bd}}>{v.name.charAt(0)}</span>
            <span style={{fontSize:11}}>{v.name}</span>
            <span style={{fontSize:9,color:'var(--text-2)'}}>{v.role}</span>
          </div>
        )})}
        {xlsPreview.length>5 && <div style={{fontSize:9,color:'var(--text-3)'}}>... +{xlsPreview.length-5}</div>}
        <div className="f2" style={{marginTop:7}}>
          <button className="btn" style={{justifyContent:'center',fontSize:10}} onClick={()=>setXlsPreview(null)}>{tl.xlsCancel}</button>
          <button className="btn btn-p" style={{justifyContent:'center',fontSize:10}} onClick={async()=>{await onBulkAdd(xlsPreview);setXlsPreview(null)}}>{tl.xlsConfirm}</button>
        </div>
      </div>
    )}
    <hr style={{border:'none',borderTop:'0.5px solid var(--border)',margin:'2px 0'}} />
    {visitors.length===0
      ? <div style={{textAlign:'center',padding:'12px 0',fontSize:11,color:'var(--text-3)'}}>{tl.visNone}</div>
      : visitors.map(v=>{
          const c=vcol(v.color_idx); const cnt=meetings.filter(m=>m.visitor_id===v.id).length
          return (
            <div key={v.id} style={{display:'flex',alignItems:'center',gap:8,padding:'7px 9px',
              borderRadius:6,border:'0.5px solid var(--border)',background:'var(--bg-2)'}}>
              <span className="vdot" style={{background:c.bg,color:c.col,borderColor:c.bd,width:24,height:24,fontSize:11}}>{v.name.charAt(0)}</span>
              <div style={{flex:1,minWidth:0}}>
                <div style={{fontSize:12,fontWeight:500,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{vName(v)}</div>
                <div style={{fontSize:9,color:'var(--text-2)'}}>{v.role} · {cnt}{lang==='ja'?'件':' mtgs'}</div>
              </div>
              <button className="btn btn-icon" onClick={()=>onDelete(v.id)} style={{fontSize:11}}>×</button>
            </div>
          )
        })
    }
    <hr style={{border:'none',borderTop:'0.5px solid var(--border)',margin:'2px 0'}} />
    <div style={{fontSize:11,fontWeight:500,marginBottom:2}}>{tl.visAdd}</div>
    <div className="fld"><label>{tl.visName}</label><input value={nm} onChange={e=>setNm(e.target.value)} /></div>
    <div className="fld"><label>{tl.visNameEn}</label><input value={nmEn} onChange={e=>setNmEn(e.target.value)} /></div>
    <div className="fld"><label>{tl.visRole}</label><input value={role} onChange={e=>setRole(e.target.value)} /></div>
    <div className="fld"><label>{tl.visColor}</label>
      <div style={{display:'flex',gap:6,padding:'4px 0'}}>
        {VCOLS.map((c,i)=>(
          <span key={i} onClick={()=>setSelColor(i)}
            style={{width:22,height:22,borderRadius:'50%',cursor:'pointer',
              background:c.bg,border:`2px solid ${c.bd}`,
              transform:selColor===i?'scale(1.3)':'scale(1)',
              boxShadow:selColor===i?`0 0 0 2px var(--bg),0 0 0 3.5px ${c.bd}`:'none',
              transition:'all .12s'}} />
        ))}
      </div>
    </div>
    <button className="btn btn-p" onClick={handleAdd} style={{width:'100%',justifyContent:'center'}}>+ {tl.visAdd}</button>
  </>
}

// ── shared ────────────────────────────────────────────────────────────────────
function PanelHeader({ title, subtitle, onClose }) {
  return (
    <div style={{display:'flex',alignItems:'flex-start',justifyContent:'space-between',gap:6}}>
      <div>
        <div style={{fontSize:13,fontWeight:500,lineHeight:1.4}}>{title}</div>
        {subtitle && <div style={{fontSize:10,color:'var(--text-2)',marginTop:2}}>{subtitle}</div>}
      </div>
      <button className="btn btn-icon" onClick={onClose} style={{flexShrink:0}}>×</button>
    </div>
  )
}
function FieldRow({ label, value, muted }) {
  return (
    <div>
      <div style={S.label}>{label}</div>
      <div style={{fontSize:11,color:muted?'var(--text-2)':'var(--text)',lineHeight:1.5,marginTop:2}}>{value}</div>
    </div>
  )
}

const S = {
  topbar: {background:'var(--bg)',borderBottom:'0.5px solid var(--border)',padding:'8px 12px',
    display:'flex',alignItems:'center',justifyContent:'space-between',gap:8,flexShrink:0},
  logo:   {background:'#3B1A6E',color:'#fff',fontSize:11,fontWeight:500,padding:'3px 9px',borderRadius:6},
  bar:    {background:'var(--bg)',borderBottom:'0.5px solid var(--border)',padding:'5px 12px',
    display:'flex',alignItems:'center',gap:7,flexWrap:'wrap',minHeight:34,flexShrink:0,overflow:'hidden',maxWidth:'100%'},
  bg:     {display:'flex',alignItems:'center',gap:5},
  bl:     {fontSize:11,color:'var(--text-2)',whiteSpace:'nowrap'},
  rp:     {width:268,flexShrink:0,borderLeft:'0.5px solid var(--border)',background:'var(--bg)',overflowY:'auto'},
  rpi:    {padding:13,display:'flex',flexDirection:'column',gap:9,minHeight:'100%'},
  legend: {display:'flex',gap:10,alignItems:'center',padding:'5px 12px',
    borderTop:'0.5px solid var(--border)',background:'var(--bg)',flexWrap:'wrap',flexShrink:0},
  label:  {fontSize:9,color:'var(--text-2)',textTransform:'uppercase',letterSpacing:.5},
  achip:  {fontSize:9,padding:'2px 6px',borderRadius:8,background:'var(--bg-2)',
    color:'var(--text-2)',border:'0.5px solid var(--border)'},
}
