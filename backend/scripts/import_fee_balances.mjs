import ExcelJS from 'exceljs'
import pg from 'pg'

const workbookPath = process.argv[2]
const applyChanges = process.argv.includes('--apply')
if (!workbookPath) throw new Error('Usage: node backend/scripts/import_fee_balances.mjs <workbook.xlsx> [--apply]')
if (!process.env.DATABASE_URL) throw new Error('DATABASE_URL is required.')

const gradeLabels = new Map([
  ['BABY CLASS', 'Baby Class'], ['PP1', 'PP1'], ['PP2', 'PP2'],
  ['GRADE ONE', 'Grade 1'], ['GRADE TWO', 'Grade 2'], ['GRADE THREE', 'Grade 3'],
  ['GRADE FOUR', 'Grade 4'], ['GRADE FIVE', 'Grade 5'], ['GRADE SIX', 'Grade 6'],
  ['SEVEN', 'Grade 7'], ['GRADE SEVEN', 'Grade 7'], ['GRADE EIGHT', 'Grade 8'], ['GRADE NINE', 'Grade 9']
])
const sourceName = 'Book1 gikundi.xlsx'
const normText = value => String(value ?? '').toUpperCase().replace(/[^A-Z0-9]/g, '')
const normAdmission = value => String(value ?? '').trim().toUpperCase().replace(/\s+/g, '')
const baseAdmission = value => normAdmission(value).replace(/R\d*$/, '')
const normPhone = value => String(value ?? '').replace(/\D/g, '').replace(/^0(?=7\d{8}$)/, '254')
const cellText = cell => String(cell?.text ?? cell?.value ?? '').trim()
const numeric = cell => {
  const value = cell?.value
  if (typeof value === 'number' && Number.isFinite(value)) return value
  const parsed = Number(String(value ?? '').replace(/[^0-9.-]/g, ''))
  return Number.isFinite(parsed) ? parsed : 0
}

async function parseRows() {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(workbookPath)
  const rows = []
  for (const sheet of workbook.worksheets) {
    let grade = ''
    let isDataSection = false
    for (const row of sheet.getRows(1, sheet.rowCount) || []) {
      const values = row.values.slice(1)
      const detected = values.map(value => gradeLabels.get(String(value ?? '').trim().toUpperCase())).find(Boolean)
      if (detected) { grade = detected; isDataSection = false; continue }
      const isHeader = cellText(row.getCell(2)).toUpperCase().replace(/\./g, '') === 'NO' && cellText(row.getCell(3)).toUpperCase() === 'NAME'
      if (isHeader) { isDataSection = Boolean(grade); continue }
      if (!isDataSection || !grade) continue
      const serial = numeric(row.getCell(2))
      const learnerName = cellText(row.getCell(3))
      if (!serial || !learnerName || learnerName.toUpperCase() === 'NAME') continue
      const admissionNumber = cellText(row.getCell(1))
      const feePlanAmount = Math.max(0, Math.round(numeric(row.getCell(9))))
      const balance = [9, 10, 11].reduce((sum, column) => sum + numeric(row.getCell(column)), 0)
      rows.push({ sheet: sheet.name, row: row.number, grade, admissionNumber, learnerName, fatherPhone: normPhone(cellText(row.getCell(6))), motherPhone: normPhone(cellText(row.getCell(8))), feePlanAmount, balance: Math.max(0, Math.round(balance)) })
    }
  }
  return rows
}

function matchRow(row, learners) {
  const name = normText(row.learnerName)
  const phones = new Set([row.fatherPhone, row.motherPhone].filter(Boolean))
  const exactAdmission = row.admissionNumber ? learners.filter(item => normAdmission(item.admissionNumber) === normAdmission(row.admissionNumber)) : []
  const choose = candidates => candidates.length === 1 ? candidates[0] : null
  const sameName = candidates => candidates.filter(item => normText(item.learnerName) === name)
  const sameGrade = candidates => candidates.filter(item => item.grade === row.grade)
  const samePhone = candidates => candidates.filter(item => [item.fatherPhone, item.motherPhone, item.guardianPhone, ...(item.guardians || []).map(guardian => guardian.phone)].map(normPhone).some(phone => phones.has(phone)))
  let learner = choose(sameName(exactAdmission))
  if (learner) return { learner, method: 'admission number' }
  learner = choose(samePhone(sameName(exactAdmission)))
  if (learner) return { learner, method: 'admission number and parent phone' }
  const baseMatches = row.admissionNumber ? learners.filter(item => baseAdmission(item.admissionNumber) === baseAdmission(row.admissionNumber)) : []
  learner = choose(sameName(baseMatches))
  if (learner) return { learner, method: 'repeated admission number and name' }
  learner = choose(samePhone(sameName(baseMatches))) || choose(samePhone(baseMatches))
  if (learner) return { learner, method: 'repeated admission number and parent phone' }
  if (row.admissionNumber && (exactAdmission.length || baseMatches.length)) return { learner: null, reason: 'admission number belongs to a different learner', candidates: [...new Set([...exactAdmission, ...baseMatches].map(item => item.id))] }
  const nameMatches = learners.filter(item => normText(item.learnerName) === name)
  learner = choose(sameGrade(nameMatches)) || choose(nameMatches)
  if (learner) return { learner, method: 'learner name' }
  learner = choose(samePhone(sameGrade(nameMatches))) || choose(samePhone(nameMatches))
  if (learner) return { learner, method: 'learner name and parent phone' }
  return { learner: null, reason: nameMatches.length || exactAdmission.length || baseMatches.length ? 'ambiguous' : 'not found', candidates: [...new Set([...nameMatches, ...exactAdmission, ...baseMatches].map(item => item.id))] }
}

const importedRows = await parseRows()
const { Pool } = pg
const database = new Pool({ connectionString: process.env.DATABASE_URL, ssl: process.env.DATABASE_SSL === 'true' ? { rejectUnauthorized: false } : undefined })
try {
  await database.query('BEGIN')
  const result = await database.query('SELECT payload FROM school_portal_state WHERE id = 1 FOR UPDATE')
  if (!result.rowCount) throw new Error('The school portal state was not found.')
  const data = result.rows[0].payload
  const matched = [], unmatched = [], ambiguous = []
  const usedLearnerIds = new Set()
  for (const row of importedRows) {
    const result = matchRow(row, data.admissions || [])
    if (!result.learner) (result.reason === 'ambiguous' ? ambiguous : unmatched).push({ ...row, candidates: result.candidates || [] })
    else if (usedLearnerIds.has(result.learner.id)) ambiguous.push({ ...row, candidates: [result.learner.id], reason: 'multiple spreadsheet rows matched the same learner' })
    else { usedLearnerIds.add(result.learner.id); matched.push({ ...row, learner: result.learner, method: result.method }) }
  }
  const boarders = matched.filter(item => item.feePlanAmount >= 20800)
  const report = {
    spreadsheetRows: importedRows.length,
    matched: matched.length,
    unmatched: unmatched.length,
    ambiguous: ambiguous.length,
    boarders: boarders.length,
    totalImportedBalance: matched.reduce((sum, item) => sum + item.balance, 0),
    unmatchedExamples: unmatched.slice(0, 25).map(item => `${item.grade}: ${item.learnerName} (${item.admissionNumber || 'no admission number'})`),
    ambiguousExamples: ambiguous.slice(0, 25).map(item => `${item.grade}: ${item.learnerName} (${item.admissionNumber || 'no admission number'})`),
    ambiguousCandidates: ambiguous.slice(0, 25).map(item => ({ row: `${item.grade}: ${item.learnerName} (${item.admissionNumber || 'no admission number'})`, candidates: item.candidates.map(id => { const learner = data.admissions.find(record => record.id === id); return learner ? `${learner.learnerName} (${learner.admissionNumber || 'no admission number'}, ${learner.grade})` : id }) }))
  }
  if (!applyChanges) {
    console.log(JSON.stringify(report, null, 2))
    await database.query('ROLLBACK')
    process.exitCode = report.unmatched || report.ambiguous ? 2 : 0
  } else {
    if (unmatched.length || ambiguous.length) throw new Error(`Import stopped: ${unmatched.length} unmatched and ${ambiguous.length} ambiguous rows. Resolve these before applying.`)
    const importedAt = new Date().toISOString()
    data.feeBalanceImports = (data.feeBalanceImports || []).filter(item => item.source !== sourceName)
    for (const item of matched) data.feeBalanceImports.push({ id: `fee-balance-import-${item.learner.id}`, admissionId: item.learner.id, learnerName: item.learner.learnerName, admissionNumber: item.learner.admissionNumber, balance: item.balance, source: sourceName, importedAt })
    const boarderIds = new Set(boarders.map(item => item.learner.id))
    data.admissions = data.admissions.map(learner => boarderIds.has(learner.id) ? { ...learner, residency: 'Boarder', residencyUpdatedAt: importedAt, residencyUpdatedBy: 'Fee balance import', updatedAt: importedAt } : learner)
    data.financeAudit ??= []
    data.financeAudit.unshift({ id: `finance-audit-${Date.now()}`, action: 'Imported', recordType: 'Fee balances', recordId: sourceName, detail: `${matched.length} learner balances imported. ${boarders.length} learners with a KES 20,800 or higher fee plan marked Boarder.`, by: 'System data import', role: 'system-admin', createdAt: importedAt })
    await database.query('UPDATE school_portal_state SET payload = $1::jsonb, updated_at = NOW() WHERE id = 1', [JSON.stringify(data)])
    await database.query('COMMIT')
    console.log(JSON.stringify({ ...report, applied: true }, null, 2))
  }
} catch (error) {
  await database.query('ROLLBACK').catch(() => {})
  throw error
} finally {
  await database.end()
}
