/**
 * .ics ファイルを生成してダウンロードする
 * Microsoftアカウント認証なしでOutlook/Teamsカレンダーに取り込める
 */
export function downloadICS(meeting, weekDates, visitors) {
  const d = weekDates[meeting.day_index]
  if (!d) return

  const pad = n => String(n).padStart(2, '0')
  const yr  = d.getFullYear()
  const mo  = pad(d.getMonth() + 1)
  const dy  = pad(d.getDate())

  const toHHMM = t => {
    const hr  = Math.floor(t)
    const min = Math.round((t % 1) * 60)
    return `${pad(hr)}${pad(min)}00`
  }

  const dtStart = `${yr}${mo}${dy}T${toHHMM(meeting.start_time)}`
  const dtEnd   = `${yr}${mo}${dy}T${toHHMM(meeting.end_time)}`
  const uid     = `${meeting.id}@ifs-visit-scheduler`

  const v = visitors.find(x => x.id === meeting.visitor_id)
  const execStr = v
    ? `${v.name} (${v.role})`
    : meeting.visitor_scope === 'all' ? 'All Visitors' : ''

  const description = [
    execStr && `Executive: ${execStr}`,
    meeting.notes,
    meeting.briefing,
  ].filter(Boolean).join('\\n\\n')

  const foldLine = s =>
    s.match(/.{1,75}/g)?.join('\r\n ') ?? s

  const attendeeLines = (meeting.attendees || [])
    .map(a => {
      const email = a.includes('@') ? a : `${a.replace(/\s/g, '.')}@ifs.com`
      return foldLine(`ATTENDEE;CUTYPE=INDIVIDUAL;ROLE=REQ-PARTICIPANT;CN=${a}:mailto:${email}`)
    })
    .join('\r\n')

  const lines = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//IFS Japan Visit Scheduler//EN',
    'CALSCALE:GREGORIAN',
    'METHOD:REQUEST',
    'BEGIN:VEVENT',
    `UID:${uid}`,
    `DTSTART;TZID=Asia/Tokyo:${dtStart}`,
    `DTEND;TZID=Asia/Tokyo:${dtEnd}`,
    foldLine(`SUMMARY:${meeting.title}`),
    foldLine(`DESCRIPTION:${description}`),
    `STATUS:${meeting.status === 'confirmed' ? 'CONFIRMED' : 'TENTATIVE'}`,
    attendeeLines,
    'END:VEVENT',
    'END:VCALENDAR',
  ].filter(Boolean).join('\r\n')

  const blob = new Blob([lines], { type: 'text/calendar;charset=utf-8' })
  const url  = URL.createObjectURL(blob)
  const a    = document.createElement('a')
  a.href     = url
  a.download = `${(meeting.title || 'meeting').replace(/\s+/g, '_')}.ics`
  document.body.appendChild(a)
  a.click()
  document.body.removeChild(a)
  URL.revokeObjectURL(url)
}
