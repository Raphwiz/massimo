import ExcelJS from 'exceljs'
import pg from 'pg'

const [workbookFile, mode = '--dry-run'] = process.argv.slice(2)
if (!workbookFile) throw new Error('Usage: node sync_classes_from_workbook.mjs <workbook.xlsx> [--apply]')
if (!process.env.DATABASE_URL) throw new Error('DATABASE_URL is required.')
if (!['--dry-run', '--apply'].includes(mode)) throw new Error('Use --dry-run or --apply.')

const gradeMap = new Map([
  ['BABY CLASS', 'Baby Class'], ['PP1', 'PP1'], ['PP2', 'PP2'],
  ['GRADE ONE', 'Grade 1'], ['GRADE TWO', 'Grade 2'], ['GRADE THREE', 'Grade 3'], ['GRADE FOUR', 'Grade 4'], ['GRADE FIVE', 'Grade 5'], ['GRADE SIX', 'Grade 6'], ['SEVEN', 'Grade 7'], ['GRADE SEVEN', 'Grade 7'], ['GRADE EIGHT', 'Grade 8'], ['GRADE NINE', 'Grade 9']
])
const normaliseNumber = value => String(value ?? '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '')
const normaliseName = value => String(value ?? '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '')
const scalar = value => value instanceof Date ? '' : String(value ?? '').trim()

const workbook = new ExcelJS.Workbook()
await workbook.xlsx.readFile(workbookFile)
const sheet = workbook.worksheets[0]
let currentGrade = ''
const sourceRows = []
for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
  const values = sheet.getRow(rowNumber).values.slice(1).map(scalar)
  const heading = values[2]?.toUpperCase()
  if (gradeMap.has(heading)) { currentGrade = gradeMap.get(heading); continue }
  const admissionNumber = normaliseNumber(values[0])
  const learnerName = values[2]
  if (!currentGrade || !learnerName || learnerName.toUpperCase() === 'NAME' || !values[1]) continue
  sourceRows.push({ rowNumber, grade: currentGrade, admissionNumber, learnerName, nameKey: normaliseName(learnerName) })
}
if (!sourceRows.length) throw new Error('No learner rows were found under a recognised class heading.')

const { Pool } = pg
const database = new Pool({ connectionString: process.env.DATABASE_URL, ssl: process.env.DATABASE_SSL === 'true' ? { rejectUnauthorized: false } : undefined })
const client = await database.connect()
try {
  await client.query('BEGIN')
  const result = await client.query('SELECT payload FROM school_portal_state WHERE id = 1 FOR UPDATE')
  const data = result.rows[0]?.payload
  if (!data) throw new Error('Live school portal state was not found.')
  const admissions = data.admissions || []
  const byExact = new Map(), bySource = new Map(), byName = new Map()
  for (const learner of admissions) {
    const add = (map, key) => { if (key) map.set(key, [...(map.get(key) || []), learner]) }
    add(byExact, normaliseNumber(learner.admissionNumber))
    add(bySource, normaliseNumber(learner.sourceAdmissionNumber))
    add(byName, normaliseName(learner.learnerName))
  }
  const academicYear = (data.schoolClasses || []).map(item => String(item.academicYear || '')).filter(Boolean).sort().at(-1)
  const classes = (data.schoolClasses || []).filter(item => item.status === 'Active' && item.academicYear === academicYear)
  const changes = [], unmatched = [], ambiguous = [], noClass = []
  for (const source of sourceRows) {
    const candidates = [...new Set([...(byExact.get(source.admissionNumber) || []), ...(bySource.get(source.admissionNumber) || [])])]
    const named = candidates.filter(item => normaliseName(item.learnerName) === source.nameKey)
    const nameCandidates = byName.get(source.nameKey) || []
    const matches = named.length ? named : candidates.length === 1 ? candidates : nameCandidates.length === 1 ? nameCandidates : nameCandidates.filter(item => !source.admissionNumber || normaliseNumber(item.admissionNumber).replace(/R$/, '') === source.admissionNumber.replace(/R$/, ''))
    if (!matches.length) { unmatched.push(source); continue }
    if (matches.length !== 1) { ambiguous.push({ source, matches: matches.map(item => `${item.learnerName} (${item.admissionNumber})`) }); continue }
    const learner = matches[0]
    const oldClass = learner.placement?.className || learner.grade || ''
    const stream = oldClass.replace(/^(Baby Class|PP1|PP2|Grade \d+)\s*/i, '').trim()
    const targetClass = classes.find(item => item.grade === source.grade && item.stream === stream) || classes.find(item => item.grade === source.grade)
    if (!targetClass) { noClass.push({ source, learner }); continue }
    if (learner.grade === source.grade && learner.placement?.className === targetClass.className && learner.placement?.academicYear === academicYear) continue
    changes.push({ learner, source, targetClass, oldClass })
  }
  const summary = { mode, workbookRows: sourceRows.length, matchedChanges: changes.length, unchanged: sourceRows.length - changes.length - unmatched.length - ambiguous.length - noClass.length, unmatched: unmatched.length, ambiguous: ambiguous.length, missingTargetClass: noClass.length, academicYear, sampleChanges: changes.slice(0, 12).map(item => ({ learner: item.learner.learnerName, admissionNumber: item.learner.admissionNumber, from: item.oldClass, to: item.targetClass.className })) }
  if (mode === '--apply') {
    const timestamp = new Date().toISOString()
    const changeById = new Map(changes.map(item => [item.learner.id, item]))
    data.admissions = admissions.map(learner => {
      const change = changeById.get(learner.id)
      if (!change) return learner
      const placement = { ...(learner.placement || {}), admissionId: learner.id, admissionNumber: learner.admissionNumber, learnerName: learner.learnerName, requestedGrade: change.source.grade, className: change.targetClass.className, academicYear, term: learner.placement?.term || 'Term 1', decision: 'Confirmed', notes: 'Class corrected from supplied register.', updatedAt: timestamp }
      return { ...learner, grade: change.source.grade, placement, status: `Placement confirmed · ${change.targetClass.className}`, updatedAt: timestamp }
    })
    data.classPlacements = (data.classPlacements || []).map(placement => {
      const change = changeById.get(placement.admissionId)
      return change ? { ...placement, requestedGrade: change.source.grade, className: change.targetClass.className, academicYear, updatedAt: timestamp, notes: 'Class corrected from supplied register.' } : placement
    })
    const placementIds = new Set(data.classPlacements.map(placement => placement.admissionId))
    const restoredPlacements = data.admissions.filter(learner => learner.placement && !placementIds.has(learner.id)).map((learner, index) => ({
      ...learner.placement,
      id: `placement-${Date.now()}-class-sync-${index + 1}`,
      admissionId: learner.id,
      admissionNumber: learner.admissionNumber,
      learnerName: learner.learnerName,
      requestedGrade: learner.grade,
      decision: 'Confirmed',
      notes: 'Placement register restored from verified workbook class sync.',
      assignedBy: 'Class sync',
      assignedAt: learner.placement.assignedAt || timestamp,
      updatedAt: timestamp
    }))
    data.classPlacements = [...restoredPlacements, ...data.classPlacements]
    await client.query('UPDATE school_portal_state SET payload = $1::jsonb, updated_at = NOW() WHERE id = 1', [JSON.stringify(data)])
  }
  await client.query(mode === '--apply' ? 'COMMIT' : 'ROLLBACK')
  console.log(JSON.stringify(summary, null, 2))
  if (unmatched.length || ambiguous.length || noClass.length) console.log(JSON.stringify({ unmatched: unmatched.slice(0, 12), ambiguous: ambiguous.slice(0, 8), missingTargetClass: noClass.slice(0, 8).map(item => ({ learner: item.learner.learnerName, targetGrade: item.source.grade })) }, null, 2))
} catch (error) { await client.query('ROLLBACK').catch(() => {}); throw error } finally { client.release(); await database.end() }
