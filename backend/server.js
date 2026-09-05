import { createServer } from 'node:http'
import 'dotenv/config'
import { mkdir, readFile, writeFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { randomBytes, randomInt, scryptSync, timingSafeEqual } from 'node:crypto'
import pg from 'pg'
import ExcelJS from 'exceljs'
import PDFDocument from 'pdfkit'
import { ZipArchive } from 'archiver'

const backendRoot = dirname(fileURLToPath(import.meta.url))
const root = dirname(backendRoot)
const dataFile = join(backendRoot, 'data', 'school-data.json')
const configuredDatabaseUrl = String(process.env.DATABASE_URL || '').trim()
const databaseUrl = configuredDatabaseUrl && !['database-host', 'replace-with-a-long-unique-password'].some(placeholder => configuredDatabaseUrl.includes(placeholder)) ? configuredDatabaseUrl : ''
const { Pool } = pg
const database = databaseUrl ? new Pool({ connectionString: databaseUrl, ssl: process.env.DATABASE_SSL === 'true' ? { rejectUnauthorized: false } : undefined }) : null
let databaseReady = false
const documentUploadDirectory = join(backendRoot, 'uploads', 'learner-documents')
const reportWorkbookFile = join(root, 'outputs', '01a03f7c-d568-7f21-a80f-279d6f99324f', 'st_dorothy_financial_report.xlsx')
const crestPath = join(root, 'frontend', 'public', 'assets', 'st-dorothy-crest-refined.png')
const port = Number(process.env.PORT || 8787)
const isProduction = process.env.NODE_ENV === 'production'
const allowedOrigins = new Set((process.env.ALLOWED_ORIGINS || 'http://127.0.0.1:4173,http://localhost:4173').split(',').map(item => item.trim()).filter(Boolean))
if (!isProduction) {
  allowedOrigins.add('http://127.0.0.1:4173')
  allowedOrigins.add('http://localhost:4173')
}
const sessionCookiePrefix = 'st_dorothy_portal_'
const sessionDurationMs = 8 * 60 * 60 * 1000
const sessions = new Map()
const rateBuckets = new Map()
const initialPortalPassword = process.env.INITIAL_PORTAL_PASSWORD || 'Dorothy@2026'
const systemAdminEmail = String(process.env.SYSTEM_ADMIN_EMAIL || 'system-admin@stdorothyschool.ac.ke').trim().toLowerCase()
const systemAdminInitialPassword = process.env.SYSTEM_ADMIN_INITIAL_PASSWORD || initialPortalPassword
const activationWindowMs = 24 * 60 * 60 * 1000
if (isProduction && !process.env.INITIAL_PORTAL_PASSWORD) throw new Error('INITIAL_PORTAL_PASSWORD must be set in production.')
if (isProduction && (!process.env.SYSTEM_ADMIN_EMAIL || !process.env.SYSTEM_ADMIN_INITIAL_PASSWORD)) throw new Error('SYSTEM_ADMIN_EMAIL and SYSTEM_ADMIN_INITIAL_PASSWORD must be set in production.')
if (isProduction && !databaseUrl) throw new Error('DATABASE_URL must be set in production. PostgreSQL is required for the live school portal.')
const seedPortalAccounts = [
  { id: 'account-system-admin', identifier: systemAdminEmail, role: 'system-admin', name: 'System Administrator', title: 'Technical administration' },
  { id: 'account-headteacher', identifier: 'headteacher@stdorothyschool.ac.ke', role: 'headteacher', name: 'Sr. Mary', title: 'Headteacher' },
  { id: 'account-secretary', identifier: 'secretary@stdorothyschool.ac.ke', role: 'secretary', name: 'Jane Mwangi', title: 'School Secretary' },
  { id: 'account-bursar', identifier: 'bursar@stdorothyschool.ac.ke', role: 'bursar', name: 'Anne Njeri', title: 'Bursar' },
  { id: 'account-teacher', identifier: 'teacher@stdorothyschool.ac.ke', role: 'teacher', name: 'John Otieno', title: 'Grade 7 teacher' },
  { id: 'account-parent-betty', identifier: '254717089754', role: 'parent', name: 'Betty Ntinyari', title: 'Parent / Guardian', parentPhone: '254717089754' }
]
const defaultNotices = [
  { id: 'notice-parents-meeting', title: 'Parents’ Meeting — Grade 7', message: 'Parents and guardians of Grade 7 learners are invited to a meeting in the school hall at 2:00 PM.', audience: 'All parents and guardians', publishDate: '2026-05-30', status: 'Published', showOnWebsite: false, createdAt: '2026-05-20T08:00:00.000Z', updatedAt: '2026-05-20T08:00:00.000Z' },
  { id: 'notice-midterm-break', title: 'Mid-term Break', message: 'The school will close for the mid-term break at 3:30 PM. Learners should report back on the communicated reopening date.', audience: 'All portal users', publishDate: '2026-06-06', status: 'Published', showOnWebsite: true, createdAt: '2026-05-21T08:00:00.000Z', updatedAt: '2026-05-21T08:00:00.000Z' },
  { id: 'notice-admissions', title: 'New admission interview dates', message: 'Parents with pending admission applications should contact the school office to confirm their interview or assessment appointment.', audience: 'All parents and guardians', publishDate: '2026-06-02', status: 'Draft', showOnWebsite: false, createdAt: '2026-05-22T08:00:00.000Z', updatedAt: '2026-05-22T08:00:00.000Z' }
]

const seedData = {
  calendarEvents: [
    { id: 'meeting', date: '2026-05-30', title: 'Parents’ Meeting — Grade 7', detail: '2:00 PM · School hall' },
    { id: 'break', date: '2026-06-06', title: 'Mid-term Break begins', detail: 'School closes at 3:30 PM' },
    { id: 'mass', date: '2026-06-23', title: 'Whole School Mass', detail: '8:00 AM · Parish church' },
    { id: 'briefing', date: '2026-06-27', title: 'Grade 9 assessment briefing', detail: '10:00 AM · Multipurpose room' }
  ],
  smsMessages: [],
  admissions: [],
  schoolClasses: [],
  classPlacements: [],
  staff: [],
  subjects: [],
  teacherAssignments: [],
  timetableEntries: [],
  lessonPlans: [],
  attendanceRecords: [],
  assessments: [],
  assessmentMarks: [],
  feeStructures: [],
  leaveRequests: [],
  staffAttendance: [],
  promotions: [],
  receipts: [],
  mpesaTransactions: [],
  documentRecords: [],
  notices: defaultNotices,
  internalMessages: [],
  parentMessages: [],
  accounts: [],
  accountAudit: [],
  admissionEnquiries: []
}

async function loadData() {
  if (database) {
    await ensureDatabase()
    const result = await database.query('SELECT payload FROM school_portal_state WHERE id = 1')
    return result.rows[0].payload
  }
  if (isProduction) throw new Error('PostgreSQL is required in production.')
  try { return JSON.parse(await readFile(dataFile, 'utf8')) }
  catch (error) {
    if (error.code !== 'ENOENT') throw error
    await saveData(seedData)
    return structuredClone(seedData)
  }
}

async function saveData(data) {
  if (database) {
    await ensureDatabase()
    await database.query('UPDATE school_portal_state SET payload = $1::jsonb, updated_at = NOW() WHERE id = 1', [JSON.stringify(data)])
    return
  }
  if (isProduction) throw new Error('PostgreSQL is required in production.')
  await mkdir(dirname(dataFile), { recursive: true })
  await writeFile(dataFile, `${JSON.stringify(data, null, 2)}\n`, 'utf8')
}

async function legacyOrSeedData() {
  try { return JSON.parse(await readFile(dataFile, 'utf8')) }
  catch (error) { if (error.code === 'ENOENT') return structuredClone(seedData); throw error }
}

async function ensureDatabase() {
  if (databaseReady) return
  await database.query(`CREATE TABLE IF NOT EXISTS school_portal_state (
    id SMALLINT PRIMARY KEY CHECK (id = 1),
    payload JSONB NOT NULL,
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
  )`)
  const result = await database.query('SELECT id FROM school_portal_state WHERE id = 1')
  if (!result.rowCount) {
    const initialData = await legacyOrSeedData()
    await database.query('INSERT INTO school_portal_state (id, payload) VALUES (1, $1::jsonb)', [JSON.stringify(initialData)])
    console.log('PostgreSQL initialised with the existing local school data.')
  }
  databaseReady = true
}

function send(response, status, payload) {
  response.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8'
  })
  response.end(JSON.stringify(payload))
}

function applySecurityHeaders(request, response) {
  const origin = request.headers.origin
  response.setHeader('X-Content-Type-Options', 'nosniff')
  response.setHeader('X-Frame-Options', 'DENY')
  response.setHeader('Referrer-Policy', 'no-referrer')
  response.setHeader('Permissions-Policy', 'camera=(), microphone=(), geolocation=(), payment=()')
  response.setHeader('Content-Security-Policy', "default-src 'none'; base-uri 'none'; form-action 'none'; frame-ancestors 'none'")
  response.setHeader('Cache-Control', 'no-store, private')
  response.setHeader('Vary', 'Origin')
  if (isProduction) response.setHeader('Strict-Transport-Security', 'max-age=31536000; includeSubDomains')
  if (origin && allowedOrigins.has(origin)) {
    response.setHeader('Access-Control-Allow-Origin', origin)
    response.setHeader('Access-Control-Allow-Credentials', 'true')
    response.setHeader('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS')
    response.setHeader('Access-Control-Allow-Headers', 'Content-Type, X-Portal-Role')
  }
}

function parseCookies(request) {
  return Object.fromEntries(String(request.headers.cookie || '').split(';').map(item => item.trim()).filter(Boolean).map(item => { const index = item.indexOf('='); return index < 0 ? [item, ''] : [item.slice(0, index), decodeURIComponent(item.slice(index + 1))] }))
}

function requestIp(request) {
  // The production Nginx proxy passes the visitor address here. Without this,
  // every visitor appears to come from the frontend container and shares one
  // sign-in limit.
  const forwarded = String(request.headers['x-forwarded-for'] || '').split(',')[0].trim()
  return forwarded || request.socket.remoteAddress || 'unknown'
}

function rateLimited(request, key, limit, windowMs) {
  const now = Date.now()
  const bucketKey = `${requestIp(request)}:${key}`
  const attempts = (rateBuckets.get(bucketKey) || []).filter(timestamp => now - timestamp < windowMs)
  if (attempts.length >= limit) { rateBuckets.set(bucketKey, attempts); return true }
  attempts.push(now)
  rateBuckets.set(bucketKey, attempts)
  return false
}

function signInRateLimited(request, key, limit, windowMs) {
  const now = Date.now()
  const bucketKey = `${requestIp(request)}:${key}`
  const attempts = (rateBuckets.get(bucketKey) || []).filter(timestamp => now - timestamp < windowMs)
  rateBuckets.set(bucketKey, attempts)
  return attempts.length >= limit
}

function recordFailedAttempt(request, key, windowMs) {
  const now = Date.now()
  const bucketKey = `${requestIp(request)}:${key}`
  const attempts = (rateBuckets.get(bucketKey) || []).filter(timestamp => now - timestamp < windowMs)
  attempts.push(now)
  rateBuckets.set(bucketKey, attempts)
}

function clearAttempts(request, key) {
  rateBuckets.delete(`${requestIp(request)}:${key}`)
}

function sessionCookieNameForRole(role) { return `${sessionCookiePrefix}${String(role || '').replace(/[^a-z-]/g, '')}_session` }

function requestedSessionRole(request) {
  const role = String(request.headers['x-portal-role'] || '').trim()
  return ['system-admin', 'headteacher', 'secretary', 'teacher', 'bursar', 'parent'].includes(role) ? role : ''
}

function createSession(response, account) {
  const token = randomBytes(32).toString('base64url')
  sessions.set(token, { accountId: account.id, expiresAt: Date.now() + sessionDurationMs })
  const cookie = [`${sessionCookieNameForRole(account.role)}=${token}`, 'HttpOnly', 'SameSite=Lax', 'Path=/', `Max-Age=${Math.floor(sessionDurationMs / 1000)}`]
  if (isProduction) cookie.push('Secure')
  response.setHeader('Set-Cookie', cookie.join('; '))
}

function clearSession(request, response) {
  const role = requestedSessionRole(request)
  response.setHeader('Set-Cookie', `${sessionCookieNameForRole(role)}=; HttpOnly; SameSite=Lax; Path=/; Max-Age=0${isProduction ? '; Secure' : ''}`)
}

function sessionAccount(request, data) {
  const cookies = parseCookies(request)
  const role = requestedSessionRole(request)
  const namedTokens = role ? [cookies[sessionCookieNameForRole(role)]] : Object.entries(cookies).filter(([name]) => name.startsWith(sessionCookiePrefix)).map(([, token]) => token)
  const token = namedTokens.filter(Boolean)[0]
  const session = token && sessions.get(token)
  if (!session || session.expiresAt <= Date.now()) { if (token) sessions.delete(token); return null }
  const account = data.accounts.find(item => item.id === session.accountId && item.active !== false)
  if (!account || (role && account.role !== role)) { sessions.delete(token); return null }
  return account
}

function routeRoles(method, path) {
  if ((method === 'GET' && ['/api/health', '/api/public/notices', '/api/calendar-events'].includes(path)) || (method === 'POST' && (path === '/api/auth/login' || path === '/api/auth/activate' || path === '/api/admission-enquiries'))) return null
  if (method === 'POST' && path === '/api/auth/logout') return ['system-admin', 'headteacher', 'secretary', 'teacher', 'bursar', 'parent']
  if (method === 'POST' && path === '/api/auth/change-password') return ['system-admin', 'headteacher', 'secretary', 'teacher', 'bursar', 'parent']
  if (method === 'GET' && path === '/api/auth/session') return ['system-admin', 'headteacher', 'secretary', 'teacher', 'bursar', 'parent']
  if (path === '/api/system-admin/overview') return ['system-admin']
  if (path === '/api/accounts' || path.startsWith('/api/accounts/')) return ['system-admin', 'headteacher']
  if (path.startsWith('/api/finance/')) return ['bursar']
  if (path === '/api/parent/learners' || path === '/api/parent/payment-reports') return ['parent']
  if (path === '/api/parent/messages') return method === 'POST' ? ['parent'] : ['parent', 'teacher', 'secretary']
  if (path === '/api/internal-messages') return ['teacher', 'headteacher']
  if (path === '/api/notices') return method === 'GET' ? ['headteacher', 'secretary', 'teacher', 'bursar', 'parent'] : ['secretary']
  if (path.startsWith('/api/documents')) return ['secretary']
  if (path === '/api/sms') return ['secretary']
  if (path === '/api/admissions') return ['secretary']
  if (path.startsWith('/api/admissions/')) return ['secretary']
  if (path === '/api/calendar-events') return ['secretary']
  if (path === '/api/school-classes') return method === 'GET' ? ['headteacher', 'secretary', 'teacher'] : ['headteacher']
  if (path.startsWith('/api/school-classes/')) return ['headteacher']
  if (path === '/api/class-placements') return ['headteacher', 'secretary']
  if (path === '/api/class-lists/workbook') return ['secretary', 'teacher']
  if (path === '/api/staff-list/workbook') return ['secretary']
  if (path === '/api/staff' || path === '/api/staff-attendance' || path === '/api/promotions') return ['headteacher']
  if (path === '/api/academics') return ['headteacher', 'teacher']
  if (['/api/subjects', '/api/teacher-assignments', '/api/timetable'].includes(path)) return ['headteacher']
  if (path === '/api/lesson-plans') return ['headteacher', 'teacher']
  if (path.startsWith('/api/lesson-plans/')) return ['headteacher', 'teacher']
  if (path === '/api/leave-requests') return ['headteacher', 'teacher']
  if (path === '/api/attendance' || path === '/api/attendance/workbook' || path === '/api/assessments' || path === '/api/assessment-marks') return ['headteacher', 'teacher']
  if (path === '/api/fee-structures') return ['bursar']
  return []
}

function requireAccess(request, response, data, url) {
  const roles = routeRoles(request.method, url.pathname)
  if (roles === null) return true
  if (!roles.length) { send(response, 404, { error: 'Endpoint not found.' }); return false }
  const account = sessionAccount(request, data)
  if (!account) { send(response, 401, { error: 'Please sign in to continue.' }); return false }
  if (!roles.includes(account.role)) { send(response, 403, { error: 'You do not have permission to access this resource.' }); return false }
  if (account.mustChangePassword && url.pathname !== '/api/auth/change-password' && url.pathname !== '/api/auth/logout') { send(response, 403, { error: 'Change your temporary password before accessing the portal.' }); return false }
  request.account = account
  return true
}

async function body(request) {
  let raw = ''
  let size = 0
  for await (const chunk of request) { size += chunk.length; if (size > 8 * 1024 * 1024) { const error = new Error('Request is too large.'); error.statusCode = 413; throw error } raw += chunk }
  try { return raw ? JSON.parse(raw) : {} }
  catch { const error = new Error('Request body is not valid JSON.'); error.statusCode = 400; throw error }
}

function text(value, maxLength) {
  return typeof value === 'string' ? value.trim().slice(0, maxLength) : ''
}

function normalisePhone(value) {
  const digits = String(value || '').replace(/\D/g, '')
  return digits.replace(/^0(?=7\d{8}$)/, '254')
}

function createPasswordHash(password, salt = randomBytes(16).toString('hex')) {
  return { passwordSalt: salt, passwordHash: scryptSync(password, salt, 64).toString('hex') }
}

function verifyHash(value, hash, salt) {
  if (!hash || !salt) return false
  const supplied = Buffer.from(scryptSync(value, salt, 64).toString('hex'), 'hex')
  const stored = Buffer.from(hash, 'hex')
  return supplied.length === stored.length && timingSafeEqual(supplied, stored)
}

function verifyPassword(password, account) { return verifyHash(password, account?.passwordHash, account?.passwordSalt) }

function validPassword(password) {
  return password.length >= 12 && /[a-z]/.test(password) && /[A-Z]/.test(password) && /\d/.test(password) && /[^A-Za-z0-9]/.test(password)
}

function issueActivation(account) {
  const code = String(randomInt(10_000_000, 100_000_000))
  const { passwordHash: activationCodeHash, passwordSalt: activationCodeSalt } = createPasswordHash(code)
  const activationExpiresAt = new Date(Date.now() + activationWindowMs).toISOString()
  return {
    code,
    account: { ...account, activationCodeHash, activationCodeSalt, activationExpiresAt, activationIssuedAt: new Date().toISOString(), mustChangePassword: true }
  }
}

function verifyActivationCode(code, account) {
  return Boolean(account?.activationExpiresAt && Date.parse(account.activationExpiresAt) > Date.now() && verifyHash(code, account.activationCodeHash, account.activationCodeSalt))
}

function portalAccountForClient(account) {
  return {
    id: account.id, identifier: account.identifier, role: account.role, name: account.name, title: account.title || '', parentPhone: account.parentPhone || '', active: account.active !== false,
    createdAt: account.createdAt || '', activatedAt: account.activatedAt || '', passwordChangedAt: account.passwordChangedAt || '',
    activationPending: Boolean(account.activationExpiresAt && Date.parse(account.activationExpiresAt) > Date.now()), activationExpiresAt: account.activationExpiresAt || ''
  }
}

function accountIdentifier(input, role) {
  return role === 'parent' ? normalisePhone(input) : text(input, 120).toLowerCase()
}

function activationMessage(account, code, expiresAt) {
  return `St. Dorothy School portal activation for ${account.name}: use code ${code}. It expires ${new Date(expiresAt).toLocaleString('en-KE', { timeZone: 'Africa/Nairobi' })}. Do not share this code.`
}

function createSeedAccounts() {
  return seedPortalAccounts.map(account => ({ ...account, active: true, ...createPasswordHash(account.role === 'system-admin' ? systemAdminInitialPassword : initialPortalPassword), createdAt: new Date().toISOString(), mustChangePassword: true }))
}

function ensureSystemAdminAccount(data) {
  if (data.accounts.some(account => account.id === 'account-system-admin' || account.role === 'system-admin')) return false
  const source = seedPortalAccounts.find(account => account.id === 'account-system-admin')
  data.accounts.push({ ...source, active: true, ...createPasswordHash(systemAdminInitialPassword), createdAt: new Date().toISOString(), mustChangePassword: true })
  return true
}

const allowedDocumentMimeTypes = new Set(['application/pdf', 'image/jpeg', 'image/png'])
const documentTypes = new Set(['Birth certificate', 'Transfer letter', 'Medical form', 'Passport photograph', 'Admission form', 'Other supporting document'])
const noticeAudiences = new Set(['All parents and guardians', 'Specific grade parents and guardians', 'All staff', 'All portal users'])
const noticeStatuses = new Set(['Draft', 'Published'])
const noticeTargetGrades = new Set(['Pre-Primary', ...Array.from({ length: 9 }, (_, index) => `Grade ${index + 1}`)])

function safeFileName(value) {
  const name = text(value, 120).replace(/[^a-zA-Z0-9.() _-]/g, '_')
  return name || 'learner-document'
}

function documentExtension(mimeType) {
  return mimeType === 'application/pdf' ? '.pdf' : mimeType === 'image/png' ? '.png' : '.jpg'
}

function hasExpectedDocumentSignature(mimeType, content) {
  if (mimeType === 'application/pdf') return content.subarray(0, 5).toString('ascii') === '%PDF-'
  if (mimeType === 'image/jpeg') return content.length > 3 && content[0] === 0xff && content[1] === 0xd8 && content[2] === 0xff
  if (mimeType === 'image/png') return content.length > 8 && content.subarray(0, 8).equals(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]))
  return false
}

function decodeDocumentContent(value) {
  if (typeof value !== 'string' || value.length > 7_100_000) throw new Error('Upload a PDF, JPG, or PNG file no larger than 5 MB.')
  const match = value.match(/^data:([^;]+);base64,([A-Za-z0-9+/=]+)$/)
  if (!match || !allowedDocumentMimeTypes.has(match[1])) throw new Error('Only PDF, JPG, and PNG documents can be uploaded.')
  const content = Buffer.from(match[2], 'base64')
  if (!content.length || content.length > 5 * 1024 * 1024) throw new Error('The uploaded document must be no larger than 5 MB.')
  if (!hasExpectedDocumentSignature(match[1], content)) throw new Error('The uploaded file content does not match its declared document type.')
  return { mimeType: match[1], content }
}

function documentForClient(data, record, request) {
  const learner = data.admissions.find(item => item.id === record.admissionId)
  return {
    id: record.id,
    admissionId: record.admissionId,
    learnerName: learner?.learnerName || 'Learner record unavailable',
    admissionNumber: learner?.admissionNumber || '',
    documentType: record.documentType,
    fileName: record.fileName,
    receivedDate: record.receivedDate,
    status: record.status,
    uploadedAt: record.uploadedAt,
    downloadUrl: `${String(request.headers['x-forwarded-proto'] || (isProduction ? 'https' : 'http')).split(',')[0]}://${request.headers.host || `127.0.0.1:${port}`}/api/documents/${encodeURIComponent(record.id)}/download`
  }
}

function ensureNoticeFields(notices) {
  let changed = false
  for (const notice of notices) {
    if (typeof notice.showOnWebsite === 'boolean') continue
    notice.showOnWebsite = notice.id === 'notice-midterm-break'
    changed = true
  }
  return changed
}

function nextAdmissionNumber(createdAt, admissions) {
  const year = new Date(createdAt).getFullYear() || new Date().getFullYear()
  const prefix = `SD/ADM/${year}/`
  const highest = admissions.reduce((maximum, application) => {
    if (!application.admissionNumber?.startsWith(prefix)) return maximum
    const value = Number(application.admissionNumber.slice(prefix.length))
    return Number.isInteger(value) ? Math.max(maximum, value) : maximum
  }, 0)
  return `${prefix}${String(highest + 1).padStart(4, '0')}`
}

function ensureAdmissionNumbers(admissions) {
  let changed = false
  for (const application of admissions) {
    if (application.admissionNumber) continue
    application.admissionNumber = nextAdmissionNumber(application.createdAt, admissions)
    changed = true
  }
  return changed
}

function nextReceiptNumber(issuedAt, receipts) {
  const year = new Date(issuedAt).getFullYear() || new Date().getFullYear()
  const prefix = `SD/RCT/${year}/`
  const highest = receipts.reduce((maximum, receipt) => {
    if (!receipt.receiptNumber?.startsWith(prefix)) return maximum
    const value = Number(receipt.receiptNumber.slice(prefix.length))
    return Number.isInteger(value) ? Math.max(maximum, value) : maximum
  }, 0)
  return `${prefix}${String(highest + 1).padStart(4, '0')}`
}

const defaultClassCapacity = 35
const schoolGrades = new Set(['Pre-Primary', ...Array.from({ length: 9 }, (_, index) => `Grade ${index + 1}`)])
const classStatuses = new Set(['Active', 'Inactive'])
const placementDecisions = new Set(['Pending headteacher approval', 'Confirmed', 'Waitlisted'])
const attendanceStatuses = new Set(['Present', 'Absent', 'Late', 'Excused'])
const staffRoles = new Set(['Headteacher', 'Deputy Headteacher', 'Teacher', 'Class Teacher', 'Secretary', 'Bursar', 'Support Staff'])

function currentFeeSummary(data, application) {
  const placement = application.placement
  if (!placement?.academicYear || !placement?.term) return { feeDue: null, totalPaid: 0, balance: null, term: null }
  const structure = data.feeStructures.find(item => item.grade === application.grade && item.academicYear === placement.academicYear && item.term === placement.term)
  const totalPaid = data.receipts.filter(receipt => receipt.admissionId === application.id).reduce((total, receipt) => total + Number(receipt.amount || 0), 0)
  const feeDue = structure ? Number(structure.amount) : null
  return { feeDue, totalPaid, balance: feeDue === null ? null : Math.max(0, feeDue - totalPaid), term: structure ? `${structure.term} ${structure.academicYear}` : null }
}

const money = value => `KES ${Number(value || 0).toLocaleString('en-KE')}`
const safePdfName = value => String(value || 'learner').replace(/[^a-z0-9_-]+/gi, '_').replace(/^_+|_+$/g, '')
async function createFeeStatementPdf(data, learner) {
  const finance = currentFeeSummary(data, learner); const receipts = data.receipts.filter(item => item.admissionId === learner.id).slice().sort((a, b) => String(a.issuedAt).localeCompare(String(b.issuedAt)))
  const statementNumber = `SD/STM/${new Date().getFullYear()}/${learner.admissionNumber || learner.id}`; const generatedAt = new Date(); const payBy = new Date(generatedAt); payBy.setDate(payBy.getDate() + 14)
  const chunks = []; const document = new PDFDocument({ size: 'A4', margin: 42 }); document.on('data', chunk => chunks.push(chunk))
  const finished = new Promise((resolve, reject) => { document.on('end', () => resolve(Buffer.concat(chunks))); document.on('error', reject) })
  try { document.image(await readFile(crestPath), 42, 36, { fit: [54, 54] }) } catch { /* Statement remains valid without its crest. */ }
  document.fillColor('#173662').fontSize(17).font('Helvetica-Bold').text("ST. DOROTHY'S SCHOOL MASSIMO", 108, 42); document.fontSize(8.5).font('Helvetica').text('P.O. Box 36-60607, Mikinduri, Meru, Kenya | +254 722 841 569 | stdorothysschoolmassimo@gmail.com', 108, 66); document.moveTo(42, 100).lineTo(553, 100).stroke('#1D5EAA')
  document.fillColor('#173662').fontSize(15).font('Helvetica-Bold').text('FEE STATEMENT', 42, 116); document.fontSize(8.5).font('Helvetica').fillColor('#536b84').text(`Statement No: ${statementNumber}`, 42, 138).text(`Generated: ${generatedAt.toLocaleString('en-KE')} | Statement period: ${finance.term || 'Current placement period'}`, 42, 151)
  document.fillColor('#173662').fontSize(10).font('Helvetica-Bold').text('Learner and guardian', 42, 180); document.font('Helvetica').fontSize(9).text(`Learner: ${learner.learnerName}`, 42, 198).text(`Admission number: ${learner.admissionNumber || 'Pending'}    Grade/Class: ${learner.placement?.className || learner.grade}`, 42, 212).text(`Guardian: ${learner.guardianName || 'Not recorded'}    Phone: ${learner.guardianPhone || 'Not recorded'}`, 42, 226)
  document.fillColor('#173662').font('Helvetica-Bold').fontSize(10).text('Account summary', 42, 258); document.font('Helvetica').fontSize(9).text(`Fee charged: ${finance.feeDue === null ? 'Fee plan pending' : money(finance.feeDue)}`, 42, 277).text(`Discounts / bursaries: ${money(0)} (none recorded)`, 42, 291).text(`Recorded payments: ${money(finance.totalPaid)}`, 42, 305); document.fillColor('#933B3B').font('Helvetica-Bold').fontSize(13).text(`Outstanding balance: ${finance.balance === null ? 'Fee plan pending' : money(finance.balance)}`, 42, 328)
  document.fillColor('#173662').fontSize(10).text('Recorded payments and receipt references', 42, 360); document.font('Helvetica-Bold').fontSize(8).text('DATE', 42, 380).text('RECEIPT / REFERENCE', 125, 380).text('METHOD', 315, 380).text('AMOUNT', 440, 380); document.moveTo(42, 392).lineTo(553, 392).stroke('#B8C9DA'); document.font('Helvetica').fontSize(8)
  let y = 403; if (!receipts.length) document.text('No confirmed payments recorded for this learner.', 42, y)
  receipts.slice(0, 12).forEach(receipt => { document.text(String(receipt.issuedAt || '').slice(0, 10), 42, y).text(`${receipt.receiptNumber || ''}${receipt.paymentReference ? ` / ${receipt.paymentReference}` : ''}`, 125, y, { width: 180 }).text(receipt.paymentMethod || '', 315, y).text(money(receipt.amount), 440, y); y += 16 })
  const guidanceY = Math.max(y + 24, 455); document.fillColor('#173662').font('Helvetica-Bold').fontSize(10).text(`Payment guidance - Pay by ${payBy.toLocaleDateString('en-KE')}`, 42, guidanceY); document.font('Helvetica').fontSize(8.5).text('Use the learner name and grade as your account reference. Retain the payment confirmation for reconciliation.', 42, guidanceY + 16, { width: 500 }); document.font('Helvetica-Bold').fontSize(9).text('Official payment channels', 42, guidanceY + 42); document.font('Helvetica').fontSize(8.5).text('Equity Paybill 247247 - Account 803434#Learner Name/Grade', 42, guidanceY + 57).text('DTB Paybill 516600 - Account 808616#Learner Name/Grade', 42, guidanceY + 71).text('Trans Nation Paybill 507900 - Account 9000482#Learner Name/Grade', 42, guidanceY + 85).text('Bank: Equity Bank Meru A/C 0140296813439 | DTB Bank Meru A/C 0259364001', 42, guidanceY + 99, { width: 500 }).text('For help, contact the Bursar on +254 722 841 569.', 42, guidanceY + 113)
  document.fillColor('#536b84').fontSize(7.5).text('This is a computer-generated private fee statement. It is intended only for the learner’s parent or guardian.', 42, 760, { align: 'center', width: 510 }); document.end(); return finished
}

function learnerAttendanceSummary(data, admissionId) {
  const records = data.attendanceRecords.filter(item => item.admissionId === admissionId)
  const present = records.filter(item => item.status === 'Present' || item.status === 'Late').length
  return { total: records.length, present, percentage: records.length ? Math.round(present / records.length * 100) : null, latest: records.sort((a, b) => b.date.localeCompare(a.date))[0] || null }
}

function learnerResultSummary(data, admissionId) {
  const marks = data.assessmentMarks.filter(item => item.admissionId === admissionId)
  const total = marks.reduce((sum, item) => sum + Number(item.score || 0), 0)
  const maximum = marks.reduce((sum, item) => sum + Number(item.outOf || 0), 0)
  return { count: marks.length, average: maximum ? Math.round(total / maximum * 100) : null, latest: marks.sort((a, b) => b.updatedAt.localeCompare(a.updatedAt))[0] || null }
}

function ensureSchoolClasses(data) {
  let changed = false
  for (const placement of data.classPlacements) {
    if (!placement.className || !placement.academicYear || data.schoolClasses.some(item => item.className === placement.className && item.academicYear === placement.academicYear)) continue
    const [grade, ...streamParts] = placement.className.startsWith('Pre-Primary') ? ['Pre-Primary', placement.className.replace('Pre-Primary', '').trim()] : placement.className.split(' ')
    const stream = placement.className.startsWith('Pre-Primary') ? streamParts[0] : streamParts.slice(1).join(' ')
    data.schoolClasses.push({
      id: `class-migrated-${placement.id}`,
      className: placement.className,
      grade: grade === 'Grade' ? `${grade} ${streamParts[0]}` : grade,
      stream: grade === 'Grade' ? streamParts.slice(1).join(' ') : stream,
      academicYear: placement.academicYear,
      capacity: defaultClassCapacity,
      classTeacher: '',
      status: 'Active',
      createdAt: placement.assignedAt || new Date().toISOString(),
      createdBy: 'System migration'
    })
    changed = true
  }
  return changed
}

const server = createServer(async (request, response) => {
  applySecurityHeaders(request, response)
  if (request.method === 'OPTIONS') {
    if (request.headers.origin && !allowedOrigins.has(request.headers.origin)) return send(response, 403, { error: 'This origin is not allowed.' })
    return response.writeHead(204).end()
  }
  try {
    const url = new URL(request.url, `http://${request.headers.host}`)
    if (rateLimited(request, 'api', 300, 60_000)) return send(response, 429, { error: 'Too many requests. Please wait a moment and try again.' })
    if (['POST', 'PUT'].includes(request.method) && !String(request.headers['content-type'] || '').toLowerCase().startsWith('application/json')) return send(response, 415, { error: 'Requests must use application/json.' })
    const data = await loadData()
    data.admissions ??= []
    data.schoolClasses ??= []
    data.classPlacements ??= []
    data.staff ??= []
    data.subjects ??= []
    data.teacherAssignments ??= []
    data.timetableEntries ??= []
    data.lessonPlans ??= []
    data.attendanceRecords ??= []
    data.assessments ??= []
    data.assessmentMarks ??= []
    data.feeStructures ??= []
    data.leaveRequests ??= []
    data.staffAttendance ??= []
    data.promotions ??= []
    data.receipts ??= []
    data.mpesaTransactions ??= []
    data.documentRecords ??= []
    data.internalMessages ??= []
    data.parentMessages ??= []
    const accountsAdded = !Array.isArray(data.accounts) || !data.accounts.length
    data.accounts = accountsAdded ? createSeedAccounts() : data.accounts
    const systemAdminAdded = ensureSystemAdminAccount(data)
    data.accountAudit ??= []
    const noticesAdded = !Array.isArray(data.notices)
    data.notices = noticesAdded ? structuredClone(defaultNotices) : data.notices
    const noticeFieldsAdded = ensureNoticeFields(data.notices)
    const admissionNumbersAdded = ensureAdmissionNumbers(data.admissions)
    const schoolClassesAdded = ensureSchoolClasses(data)

    if (accountsAdded || systemAdminAdded) await saveData(data)
    if (!requireAccess(request, response, data, url)) return

    if (request.method === 'GET' && url.pathname === '/api/health') return send(response, 200, { status: 'ok' })
    if (request.method === 'POST' && url.pathname === '/api/auth/login') {
      if (signInRateLimited(request, 'login', 8, 15 * 60_000)) return send(response, 429, { error: 'Too many sign-in attempts. Please wait 15 minutes before trying again.' })
      const input = await body(request)
      const portal = text(input.portal, 12)
      const identifier = portal === 'parent' ? normalisePhone(input.identifier) : text(input.identifier, 120).toLowerCase()
      const password = typeof input.password === 'string' ? input.password : ''
      const account = data.accounts.find(item => item.identifier === identifier && item.active !== false)
      if (!account || !password || !verifyPassword(password, account) || (portal === 'parent' && account.role !== 'parent') || (portal === 'staff' && (account.role === 'parent' || account.role === 'system-admin')) || (portal === 'system-admin' && account.role !== 'system-admin')) {
        recordFailedAttempt(request, 'login', 15 * 60_000)
        return send(response, 401, { error: 'The sign-in details are not recognised. Please try again or contact the school office.' })
      }
      clearAttempts(request, 'login')
      createSession(response, account)
      return send(response, 200, { id: account.id, role: account.role, name: account.name, title: account.title, parentPhone: account.parentPhone || '', mustChangePassword: Boolean(account.mustChangePassword) })
    }
    if (request.method === 'GET' && url.pathname === '/api/auth/session') return send(response, 200, { id: request.account.id, role: request.account.role, name: request.account.name, title: request.account.title || '', parentPhone: request.account.parentPhone || '', mustChangePassword: Boolean(request.account.mustChangePassword) })
    if (request.method === 'POST' && url.pathname === '/api/auth/activate') {
      if (signInRateLimited(request, 'activation', 8, 15 * 60_000)) return send(response, 429, { error: 'Too many activation attempts. Please wait 15 minutes before trying again.' })
      const input = await body(request)
      const rawIdentifier = text(input.identifier, 120)
      const identifier = rawIdentifier.includes('@') ? rawIdentifier.toLowerCase() : normalisePhone(rawIdentifier)
      const code = text(input.code, 16)
      const newPassword = typeof input.newPassword === 'string' ? input.newPassword : ''
      const account = data.accounts.find(item => item.identifier === identifier && item.active !== false)
      if (!account || !verifyActivationCode(code, account)) {
        recordFailedAttempt(request, 'activation', 15 * 60_000)
        return send(response, 400, { error: 'The activation details are not recognised or the code has expired.' })
      }
      clearAttempts(request, 'activation')
      if (!validPassword(newPassword)) return send(response, 400, { error: 'Use at least 12 characters with upper- and lower-case letters, a number, and a symbol.' })
      const updatedAccount = { ...account, ...createPasswordHash(newPassword), mustChangePassword: false, activatedAt: new Date().toISOString(), passwordChangedAt: new Date().toISOString() }
      delete updatedAccount.activationCodeHash
      delete updatedAccount.activationCodeSalt
      delete updatedAccount.activationExpiresAt
      delete updatedAccount.activationIssuedAt
      data.accounts = data.accounts.map(item => item.id === account.id ? updatedAccount : item)
      await saveData(data)
      createSession(response, updatedAccount)
      return send(response, 200, { role: updatedAccount.role, name: updatedAccount.name, title: updatedAccount.title, parentPhone: updatedAccount.parentPhone || '', mustChangePassword: false })
    }
    if (request.method === 'POST' && url.pathname === '/api/auth/logout') {
      const token = parseCookies(request)[sessionCookieNameForRole(requestedSessionRole(request))]
      if (token) sessions.delete(token)
      clearSession(request, response)
      return send(response, 200, { status: 'signed out' })
    }
    if (request.method === 'POST' && url.pathname === '/api/auth/change-password') {
      const input = await body(request)
      const currentPassword = typeof input.currentPassword === 'string' ? input.currentPassword : ''
      const newPassword = typeof input.newPassword === 'string' ? input.newPassword : ''
      if (!verifyPassword(currentPassword, request.account)) return send(response, 400, { error: 'Your current password is incorrect.' })
      if (!validPassword(newPassword)) return send(response, 400, { error: 'Use at least 12 characters with upper- and lower-case letters, a number, and a symbol.' })
      if (currentPassword === newPassword) return send(response, 400, { error: 'Choose a new password that differs from the current password.' })
      const updatedAccount = { ...request.account, ...createPasswordHash(newPassword), mustChangePassword: false, passwordChangedAt: new Date().toISOString() }
      data.accounts = data.accounts.map(account => account.id === updatedAccount.id ? updatedAccount : account)
      for (const [token, session] of sessions) if (session.accountId === updatedAccount.id && token !== parseCookies(request)[sessionCookieNameForRole(requestedSessionRole(request))]) sessions.delete(token)
      await saveData(data)
      return send(response, 200, { status: 'password updated' })
    }
    if (request.method === 'GET' && url.pathname === '/api/system-admin/overview') return send(response, 200, { accounts: { total: data.accounts.length, active: data.accounts.filter(account => account.active !== false).length, pendingActivation: data.accounts.filter(account => account.activationExpiresAt && Date.parse(account.activationExpiresAt) > Date.now()).length }, integrations: [{ name: 'M-Pesa Daraja', configured: Boolean(process.env.MPESA_CONSUMER_KEY && process.env.MPESA_CONSUMER_SECRET && process.env.MPESA_PASSKEY) }, { name: 'SMS provider', configured: Boolean(process.env.SMS_API_KEY) }, { name: 'School email', configured: Boolean(process.env.SMTP_HOST && process.env.SMTP_USER) }], security: { secureCookies: isProduction, allowedOrigins: allowedOrigins.size, twoFactorAuthentication: false } })
    if (request.method === 'GET' && url.pathname === '/api/accounts') return send(response, 200, data.accounts.filter(account => request.account.role === 'system-admin' || account.role !== 'system-admin').map(portalAccountForClient).sort((a, b) => a.name.localeCompare(b.name)))
    if (request.method === 'POST' && url.pathname === '/api/accounts') {
      const input = await body(request)
      const role = text(input.role, 20)
      const allowedRoles = new Set(['headteacher', 'secretary', 'bursar', 'teacher', 'parent'])
      const name = text(input.name, 100)
      const identifier = accountIdentifier(input.identifier, role)
      const title = text(input.title, 100)
      if (!allowedRoles.has(role) || (request.account.role !== 'system-admin' && role === 'headteacher') || !name || !identifier || (role === 'parent' && !/^2547\d{8}$/.test(identifier)) || (role !== 'parent' && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(identifier))) return send(response, 400, { error: request.account.role !== 'system-admin' && role === 'headteacher' ? 'Only the system administrator can create headteacher accounts.' : 'Enter a name and a valid work email or Kenyan parent phone number.' })
      if (data.accounts.some(account => account.identifier === identifier)) return send(response, 400, { error: 'An account already exists with those sign-in details.' })
      if (role === 'parent' && !data.admissions.some(admission => normalisePhone(admission.guardianPhone) === identifier)) return send(response, 400, { error: 'Create or update the learner admission first so this parent phone is linked to a learner.' })
      const baseAccount = { id: `account-${Date.now()}-${randomBytes(4).toString('hex')}`, identifier, role, name, title: title || (role === 'parent' ? 'Parent / Guardian' : role), parentPhone: role === 'parent' ? identifier : '', active: true, createdAt: new Date().toISOString() }
      const activation = issueActivation(baseAccount)
      data.accounts.push(activation.account)
      await saveData(data)
      return send(response, 201, { account: portalAccountForClient(activation.account), activation: { code: activation.code, expiresAt: activation.account.activationExpiresAt, message: activationMessage(activation.account, activation.code, activation.account.activationExpiresAt) } })
    }
    if (request.method === 'POST' && /^\/api\/accounts\/[^/]+\/reset$/.test(url.pathname)) {
      const id = decodeURIComponent(url.pathname.split('/')[3])
      const account = data.accounts.find(item => item.id === id)
      if (!account) return send(response, 404, { error: 'Portal account not found.' })
      const activation = issueActivation(account)
      delete activation.account.passwordHash
      delete activation.account.passwordSalt
      data.accounts = data.accounts.map(item => item.id === id ? activation.account : item)
      for (const [token, session] of sessions) if (session.accountId === id) sessions.delete(token)
      await saveData(data)
      return send(response, 200, { account: portalAccountForClient(activation.account), activation: { code: activation.code, expiresAt: activation.account.activationExpiresAt, message: activationMessage(activation.account, activation.code, activation.account.activationExpiresAt) } })
    }
    if (request.method === 'PUT' && /^\/api\/accounts\/[^/]+\/status$/.test(url.pathname)) {
      const id = decodeURIComponent(url.pathname.split('/')[3])
      const input = await body(request)
      const account = data.accounts.find(item => item.id === id)
      if (!account) return send(response, 404, { error: 'Portal account not found.' })
      if (id === request.account.id && input.active !== true) return send(response, 400, { error: 'You cannot deactivate the account currently in use.' })
      const updatedAccount = { ...account, active: input.active === true }
      data.accounts = data.accounts.map(item => item.id === id ? updatedAccount : item)
      if (!updatedAccount.active) for (const [token, session] of sessions) if (session.accountId === id) sessions.delete(token)
      await saveData(data)
      return send(response, 200, { account: portalAccountForClient(updatedAccount) })
    }
    if (request.method === 'DELETE' && /^\/api\/accounts\/[^/]+$/.test(url.pathname)) {
      if (request.account.role !== 'system-admin') return send(response, 403, { error: 'Only the System Administrator can remove a portal profile.' })
      const id = decodeURIComponent(url.pathname.split('/')[3])
      const account = data.accounts.find(item => item.id === id)
      if (!account) return send(response, 404, { error: 'Portal account not found.' })
      if (id === request.account.id) return send(response, 400, { error: 'You cannot remove the account currently in use.' })
      if (account.activatedAt) return send(response, 400, { error: 'This profile has been used. Deactivate it instead so the school retains an accurate access history.' })
      data.accounts = data.accounts.filter(item => item.id !== id)
      data.accountAudit.push({ id: `audit-${Date.now()}`, action: 'account-removed-before-activation', accountId: id, accountName: account.name, performedBy: request.account.id, createdAt: new Date().toISOString() })
      await saveData(data)
      return send(response, 200, { status: 'profile removed' })
    }
    if (request.method === 'GET' && url.pathname === '/api/internal-messages') return send(response, 200, data.internalMessages.slice().sort((a, b) => b.createdAt.localeCompare(a.createdAt)))
    if (request.method === 'GET' && url.pathname === '/api/parent/messages') {
      const phone = normalisePhone(url.searchParams.get('phone'))
      if (request.account.role === 'parent' && (!phone || phone !== request.account.parentPhone)) return send(response, 403, { error: 'You can only view messages linked to your own parent account.' })
      const messages = data.parentMessages.filter(message => request.account.role === 'parent' ? message.parentPhone === request.account.parentPhone : !phone || message.parentPhone === phone).slice().sort((a, b) => b.createdAt.localeCompare(a.createdAt))
      return send(response, 200, messages)
    }
    if (request.method === 'GET' && url.pathname === '/api/notices') {
      if (noticesAdded || noticeFieldsAdded) await saveData(data)
      const notices = ['headteacher', 'secretary'].includes(request.account.role) ? data.notices : data.notices.filter(item => item.status === 'Published')
      return send(response, 200, notices.slice().sort((a, b) => b.publishDate.localeCompare(a.publishDate) || b.updatedAt.localeCompare(a.updatedAt)))
    }
    if (request.method === 'GET' && url.pathname === '/api/public/notices') {
      if (noticesAdded || noticeFieldsAdded) await saveData(data)
      return send(response, 200, data.notices.filter(item => item.status === 'Published' && item.showOnWebsite).sort((a, b) => b.publishDate.localeCompare(a.publishDate)).slice(0, 4))
    }
    if (request.method === 'GET' && url.pathname === '/api/documents') return send(response, 200, data.documentRecords.map(record => documentForClient(data, record, request)).sort((a, b) => b.uploadedAt.localeCompare(a.uploadedAt)))
    if (request.method === 'GET' && /^\/api\/documents\/[^/]+\/download$/.test(url.pathname)) {
      const id = decodeURIComponent(url.pathname.slice('/api/documents/'.length, -'/download'.length))
      const record = data.documentRecords.find(item => item.id === id)
      if (!record) return send(response, 404, { error: 'Learner document not found.' })
      const content = await readFile(join(documentUploadDirectory, record.storedFileName))
      response.writeHead(200, { 'Content-Type': record.mimeType, 'Content-Disposition': `attachment; filename="${safeFileName(record.fileName)}"` })
      return response.end(content)
    }
    if (request.method === 'GET' && url.pathname === '/api/calendar-events') return send(response, 200, data.calendarEvents)
    if (request.method === 'GET' && url.pathname === '/api/sms') return send(response, 200, data.smsMessages)
    if (request.method === 'GET' && url.pathname === '/api/admissions') {
      if (admissionNumbersAdded || schoolClassesAdded) await saveData(data)
      return send(response, 200, data.admissions)
    }
    if (request.method === 'GET' && url.pathname === '/api/school-classes') {
      if (schoolClassesAdded) await saveData(data)
      return send(response, 200, data.schoolClasses.sort((a, b) => a.academicYear.localeCompare(b.academicYear) || a.className.localeCompare(b.className)))
    }
    if (request.method === 'GET' && url.pathname === '/api/class-placements') {
      if (schoolClassesAdded) await saveData(data)
      return send(response, 200, { defaultClassCapacity, classes: data.schoolClasses, placements: data.classPlacements })
    }
    if (request.method === 'GET' && url.pathname === '/api/staff') return send(response, 200, data.staff)
    if (request.method === 'GET' && url.pathname === '/api/academics') return send(response, 200, { subjects: data.subjects, assignments: data.teacherAssignments, timetable: data.timetableEntries, assessments: data.assessments })
    if (request.method === 'GET' && url.pathname === '/api/lesson-plans') {
      const plans = request.account.role === 'teacher' ? data.lessonPlans.filter(item => item.preparedBy === request.account.name) : data.lessonPlans
      return send(response, 200, plans.slice().sort((a, b) => `${b.lessonDate}${b.updatedAt}`.localeCompare(`${a.lessonDate}${a.updatedAt}`)))
    }
    if (request.method === 'GET' && url.pathname === '/api/leave-requests') return send(response, 200, data.leaveRequests)
    if (request.method === 'GET' && url.pathname === '/api/staff-attendance') return send(response, 200, data.staffAttendance)
    if (request.method === 'GET' && url.pathname === '/api/promotions') return send(response, 200, data.promotions)
    if (request.method === 'GET' && url.pathname === '/api/fee-structures') return send(response, 200, data.feeStructures)
    if (request.method === 'GET' && url.pathname === '/api/attendance') {
      const className = text(url.searchParams.get('className'), 60)
      const date = text(url.searchParams.get('date'), 10)
      return send(response, 200, data.attendanceRecords.filter(item => (!className || item.className === className) && (!date || item.date === date)))
    }
    if (request.method === 'GET' && url.pathname === '/api/attendance/workbook') {
      const className = text(url.searchParams.get('className'), 60)
      const date = text(url.searchParams.get('date'), 10)
      const hasRecords = data.attendanceRecords.some(item => item.className === className && item.date === date)
      if (!className || !/^\d{4}-\d{2}-\d{2}$/.test(date) || !hasRecords) return send(response, 400, { error: 'Save attendance for the selected class and date before downloading the Excel register.' })
      const safeClass = className.replace(/[^a-zA-Z0-9]+/g, '_').replace(/^_|_$/g, '')
      const fileName = `st_dorothy_attendance_${safeClass}_${date}.xlsx`
      const records = data.attendanceRecords.filter(item => item.className === className && item.date === date).sort((a, b) => a.learnerName.localeCompare(b.learnerName))
      const admissionById = new Map(data.admissions.map(item => [item.id, item]))
      const workbook = new ExcelJS.Workbook()
      workbook.creator = 'St. Dorothy School Massimo'
      workbook.created = new Date()
      const sheet = workbook.addWorksheet('Attendance Register', { views: [{ state: 'frozen', ySplit: 5 }] })
      sheet.columns = [{ width: 8 }, { width: 22 }, { width: 30 }, { width: 20 }, { width: 38 }, { width: 23 }]
      sheet.mergeCells('A1:F1')
      sheet.getCell('A1').value = "St. Dorothy's School Massimo — Attendance Register"
      sheet.getCell('A1').font = { name: 'Aptos Display', size: 16, bold: true, color: { argb: 'FFFFFFFF' } }
      sheet.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF173B6B' } }
      sheet.getRow(1).height = 30
      sheet.mergeCells('A2:F2')
      sheet.getCell('A2').value = `Class: ${className}  |  Date: ${date}  |  Exported ${new Date().toLocaleString('en-KE')}`
      sheet.getCell('A2').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEDF5FF' } }
      sheet.getCell('A2').font = { italic: true, color: { argb: 'FF173B6B' } }
      sheet.addRow([])
      const summary = sheet.addRow(['Registered learners', 'Present / Late', 'Absent', 'Excused', 'Attendance rate', 'Exported at'])
      const present = records.filter(item => ['Present', 'Late'].includes(item.status)).length
      const absent = records.filter(item => item.status === 'Absent').length
      const excused = records.filter(item => item.status === 'Excused').length
      sheet.addRow([records.length, present, absent, excused, records.length ? present / records.length : 0, new Date()])
      sheet.getCell('E5').numFmt = '0%'
      sheet.getCell('F5').numFmt = 'yyyy-mm-dd hh:mm'
      sheet.addRow([])
      const header = sheet.addRow(['No.', 'Admission number', 'Learner name', 'Attendance status', 'Reason / note', 'Marked at'])
      for (const row of [summary, header]) {
        row.font = { bold: true, color: { argb: 'FFFFFFFF' } }
        row.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
        row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: row === header ? 'FF2B72C8' : 'FF173B6B' } }
      }
      for (const [index, record] of records.entries()) {
        const row = sheet.addRow([index + 1, admissionById.get(record.admissionId)?.admissionNumber || '', record.learnerName, record.status, record.reason || '', new Date(record.markedAt)])
        row.getCell(6).numFmt = 'yyyy-mm-dd hh:mm'
        row.alignment = { vertical: 'center', wrapText: true }
      }
      sheet.eachRow(row => row.eachCell(cell => { cell.border = { top: { style: 'thin', color: { argb: 'FFDCE5EF' } }, left: { style: 'thin', color: { argb: 'FFDCE5EF' } }, bottom: { style: 'thin', color: { argb: 'FFDCE5EF' } }, right: { style: 'thin', color: { argb: 'FFDCE5EF' } } } }))
      const workbookBytes = await workbook.xlsx.writeBuffer()
      response.writeHead(200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="${fileName}"`
      })
      return response.end(Buffer.from(workbookBytes))
    }
    if (request.method === 'GET' && url.pathname === '/api/class-lists/workbook') {
      const classId = text(url.searchParams.get('classId'), 100)
      const schoolClass = data.schoolClasses.find(item => item.id === classId)
      if (!schoolClass) return send(response, 404, { error: 'Choose a valid class before downloading its list.' })
      if (request.account.role === 'teacher' && !data.teacherAssignments.some(item => item.classId === schoolClass.id && item.staffName === request.account.name)) return send(response, 403, { error: 'You can download class lists only for your assigned classes.' })
      const learners = data.admissions.filter(item => item.placement?.className === schoolClass.className && item.placement?.academicYear === schoolClass.academicYear).slice().sort((a, b) => a.learnerName.localeCompare(b.learnerName))
      const classTeachers = [...new Set(data.teacherAssignments.filter(item => item.classId === schoolClass.id).map(item => item.staffName).filter(Boolean))]
      const teacherName = classTeachers.length ? classTeachers.join(', ') : 'Not assigned'
      const safeClass = schoolClass.className.replace(/[^a-zA-Z0-9]+/g, '_').replace(/^_|_$/g, '')
      const workbook = new ExcelJS.Workbook()
      workbook.creator = 'St. Dorothy School Massimo'
      workbook.created = new Date()
      const sheet = workbook.addWorksheet('Class List', { views: [{ state: 'frozen', ySplit: 7 }] })
      sheet.columns = [{ width: 7 }, { width: 19 }, { width: 30 }, { width: 28 }, { width: 18 }, { width: 18 }]
      const crestPath = join(root, 'frontend', 'public', 'assets', 'st-dorothy-crest-refined.png')
      try { const crest = workbook.addImage({ buffer: await readFile(crestPath), extension: 'png' }); sheet.addImage(crest, { tl: { col: 0, row: 0 }, ext: { width: 58, height: 58 } }) } catch { /* A valid workbook is still created if the logo asset is unavailable. */ }
      sheet.mergeCells('B1:F1'); sheet.getCell('B1').value = "ST. DOROTHY'S SCHOOL MASSIMO"
      sheet.getCell('B1').font = { name: 'Aptos Display', size: 16, bold: true, color: { argb: 'FF123664' } }
      sheet.getCell('B1').alignment = { vertical: 'middle' }
      sheet.mergeCells('B2:F2'); sheet.getCell('B2').value = 'P.O. Box 36-60607, Mikinduri, Meru, Kenya · +254 722 841 569 · stdorothysschoolmassimo@gmail.com'
      sheet.getCell('B2').font = { size: 9, color: { argb: 'FF536B84' } }; sheet.getCell('B2').alignment = { wrapText: true }
      sheet.mergeCells('A4:F4'); sheet.getCell('A4').value = `CLASS LIST — ${schoolClass.className.toUpperCase()} · ${schoolClass.academicYear}`
      sheet.getCell('A4').font = { size: 13, bold: true, color: { argb: 'FFFFFFFF' } }; sheet.getCell('A4').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF123664' } }; sheet.getCell('A4').alignment = { horizontal: 'center' }
      sheet.mergeCells('A5:C5'); sheet.getCell('A5').value = `Class teacher: ${teacherName}`
      sheet.mergeCells('D5:F5'); sheet.getCell('D5').value = `Generated: ${new Intl.DateTimeFormat('en-KE', { dateStyle: 'medium', timeStyle: 'short', timeZone: 'Africa/Nairobi' }).format(new Date())}`
      for (const cell of ['A5', 'D5']) { sheet.getCell(cell).font = { bold: true, color: { argb: 'FF294A6E' } }; sheet.getCell(cell).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F6FB' } } }
      const header = sheet.addRow(['No.', 'Admission No.', 'Learner name', 'Parent / guardian', 'Parent phone', 'Grade'])
      header.font = { bold: true, color: { argb: 'FFFFFFFF' } }; header.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1D5EAA' } }; header.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
      learners.forEach((learner, index) => { const row = sheet.addRow([index + 1, learner.admissionNumber || 'Pending', learner.learnerName, learner.guardianName || '—', learner.guardianPhone || '—', learner.grade || schoolClass.grade]); row.alignment = { vertical: 'middle', wrapText: true }; if (index % 2 === 1) row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF6F9FC' } } })
      if (!learners.length) { const row = sheet.addRow(['', '', 'No learners placed in this stream yet.', '', '', '']); row.font = { italic: true, color: { argb: 'FF60758C' } } }
      sheet.eachRow(row => row.eachCell(cell => { cell.border = { top: { style: 'thin', color: { argb: 'FFD9E4EF' } }, left: { style: 'thin', color: { argb: 'FFD9E4EF' } }, bottom: { style: 'thin', color: { argb: 'FFD9E4EF' } }, right: { style: 'thin', color: { argb: 'FFD9E4EF' } } } }))
      const bytes = await workbook.xlsx.writeBuffer()
      response.writeHead(200, { 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'Content-Disposition': `attachment; filename="st_dorothy_${safeClass}_class_list_${schoolClass.academicYear}.xlsx"` })
      return response.end(Buffer.from(bytes))
    }
    if (request.method === 'GET' && url.pathname === '/api/staff-list/workbook') {
      const staffByName = new Map(data.staff.filter(item => /teacher/i.test(item.role || '')).map(item => [String(item.fullName || '').toLowerCase(), item]))
      data.accounts.filter(item => item.role === 'teacher').forEach(account => { const key = String(account.name || '').toLowerCase(); if (!staffByName.has(key)) staffByName.set(key, { fullName: account.name, role: account.title || 'Teacher', phone: '', email: account.identifier, employmentStatus: account.active === false ? 'Inactive' : 'Active' }) })
      const teachers = [...staffByName.values()].filter(item => item.fullName).sort((a, b) => String(a.fullName).localeCompare(String(b.fullName)))
      const workbook = new ExcelJS.Workbook(); workbook.creator = 'St. Dorothy School Massimo'; workbook.created = new Date()
      const sheet = workbook.addWorksheet('Teaching Staff', { views: [{ state: 'frozen', ySplit: 7 }] }); sheet.columns = [{ width: 7 }, { width: 30 }, { width: 25 }, { width: 29 }, { width: 18 }, { width: 30 }, { width: 15 }]
      const crestPath = join(root, 'frontend', 'public', 'assets', 'st-dorothy-crest-refined.png')
      try { const crest = workbook.addImage({ buffer: await readFile(crestPath), extension: 'png' }); sheet.addImage(crest, { tl: { col: 0, row: 0 }, ext: { width: 58, height: 58 } }) } catch { /* The report can still be generated if the crest file is unavailable. */ }
      sheet.mergeCells('B1:G1'); sheet.getCell('B1').value = "ST. DOROTHY'S SCHOOL MASSIMO"; sheet.getCell('B1').font = { name: 'Aptos Display', size: 16, bold: true, color: { argb: 'FF123664' } }; sheet.getCell('B1').alignment = { vertical: 'middle' }
      sheet.mergeCells('B2:G2'); sheet.getCell('B2').value = 'P.O. Box 36-60607, Mikinduri, Meru, Kenya · +254 722 841 569 · stdorothysschoolmassimo@gmail.com'; sheet.getCell('B2').font = { size: 9, color: { argb: 'FF536B84' } }; sheet.getCell('B2').alignment = { wrapText: true }
      sheet.mergeCells('A4:G4'); sheet.getCell('A4').value = 'TEACHING STAFF LIST'; sheet.getCell('A4').font = { size: 13, bold: true, color: { argb: 'FFFFFFFF' } }; sheet.getCell('A4').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF123664' } }; sheet.getCell('A4').alignment = { horizontal: 'center' }
      sheet.mergeCells('A5:D5'); sheet.getCell('A5').value = `Total teaching staff: ${teachers.length}`; sheet.mergeCells('E5:G5'); sheet.getCell('E5').value = `Generated: ${new Intl.DateTimeFormat('en-KE', { dateStyle: 'medium', timeStyle: 'short', timeZone: 'Africa/Nairobi' }).format(new Date())}`
      for (const cell of ['A5', 'E5']) { sheet.getCell(cell).font = { bold: true, color: { argb: 'FF294A6E' } }; sheet.getCell(cell).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F6FB' } } }
      const header = sheet.addRow(['No.', 'Teacher name', 'Assigned role', 'Class responsibility', 'Phone', 'Email', 'Status']); header.font = { bold: true, color: { argb: 'FFFFFFFF' } }; header.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1D5EAA' } }; header.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
      teachers.forEach((teacher, index) => { const assignments = [...new Set(data.teacherAssignments.filter(item => item.staffName === teacher.fullName).map(item => item.className).filter(Boolean))]; const row = sheet.addRow([index + 1, teacher.fullName, teacher.role || 'Teacher', assignments.length ? assignments.join(', ') : 'Not assigned', teacher.phone || '—', teacher.email || '—', teacher.employmentStatus || 'Active']); row.alignment = { vertical: 'middle', wrapText: true }; if (index % 2 === 1) row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF6F9FC' } } })
      if (!teachers.length) { const row = sheet.addRow(['', 'No teaching staff profiles recorded yet.', '', '', '', '', '']); row.font = { italic: true, color: { argb: 'FF60758C' } } }
      sheet.eachRow(row => row.eachCell(cell => { cell.border = { top: { style: 'thin', color: { argb: 'FFD9E4EF' } }, left: { style: 'thin', color: { argb: 'FFD9E4EF' } }, bottom: { style: 'thin', color: { argb: 'FFD9E4EF' } }, right: { style: 'thin', color: { argb: 'FFD9E4EF' } } } }))
      const bytes = await workbook.xlsx.writeBuffer(); response.writeHead(200, { 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'Content-Disposition': 'attachment; filename="st_dorothy_teaching_staff_list.xlsx"' }); return response.end(Buffer.from(bytes))
    }
    if (request.method === 'GET' && url.pathname === '/api/assessment-marks') {
      const assessmentId = text(url.searchParams.get('assessmentId'), 80)
      return send(response, 200, data.assessmentMarks.filter(item => !assessmentId || item.assessmentId === assessmentId))
    }
    if (request.method === 'GET' && url.pathname === '/api/parent/learners') {
      const phone = normalisePhone(url.searchParams.get('phone'))
      if (!phone) return send(response, 400, { error: 'Enter the guardian mobile number linked to the learner record.' })
      if (phone !== request.account.parentPhone) return send(response, 403, { error: 'You can only view learner records linked to your own parent account.' })
      const learners = data.admissions.filter(application => normalisePhone(application.guardianPhone) === phone).map(application => ({
        id: application.id,
        learnerName: application.learnerName,
        admissionNumber: application.admissionNumber,
        grade: application.grade,
        status: application.status,
        createdAt: application.createdAt,
        placement: application.placement ? {
          className: application.placement.className,
          academicYear: application.placement.academicYear,
          term: application.placement.term,
          decision: application.placement.decision,
          updatedAt: application.placement.updatedAt
        } : null,
        finance: currentFeeSummary(data, application),
        attendance: learnerAttendanceSummary(data, application.id),
        results: learnerResultSummary(data, application.id),
        payments: data.receipts.filter(receipt => receipt.admissionId === application.id).slice(0, 8).map(receipt => ({ receiptNumber: receipt.receiptNumber, amount: receipt.amount, method: receipt.paymentMethod, issuedAt: receipt.issuedAt })),
        pendingPayments: data.mpesaTransactions.filter(transaction => transaction.admissionId === application.id && transaction.status !== 'Matched').slice(0, 8).map(transaction => ({ transactionCode: transaction.transactionCode, amount: transaction.amount, senderPhone: transaction.senderPhone, reportedAt: transaction.reportedAt || transaction.receivedAt, status: transaction.status }))
      }))
      return send(response, 200, learners)
    }
    if (request.method === 'GET' && url.pathname === '/api/finance/learners') {
      const learners = data.admissions.map(application => ({
        id: application.id,
        admissionNumber: application.admissionNumber,
        learnerName: application.learnerName,
        grade: application.grade,
        guardianName: application.guardianName,
        guardianPhone: application.guardianPhone,
        className: application.placement?.className || null,
        placementStatus: application.placement?.decision || 'Awaiting placement',
        finance: currentFeeSummary(data, application)
      }))
      return send(response, 200, learners)
    }
    if (request.method === 'GET' && url.pathname === '/api/finance/statements-workbook') {
      const outstandingOnly = url.searchParams.get('outstanding') !== 'false'
      const rows = data.admissions.map(learner => ({ learner, finance: currentFeeSummary(data, learner) })).filter(({ finance }) => !outstandingOnly || Number(finance.balance || 0) > 0)
      const workbook = new ExcelJS.Workbook(); workbook.creator = 'St. Dorothy School Massimo'; workbook.created = new Date()
      const sheet = workbook.addWorksheet('Fee Statements Register', { views: [{ state: 'frozen', ySplit: 7 }] })
      sheet.columns = [{ width: 7 }, { width: 28 }, { width: 20 }, { width: 18 }, { width: 27 }, { width: 18 }, { width: 16 }, { width: 16 }, { width: 16 }]
      try { const crest = workbook.addImage({ buffer: await readFile(crestPath), extension: 'png' }); sheet.addImage(crest, { tl: { col: 0, row: 0 }, ext: { width: 58, height: 58 } }) } catch { /* The workbook remains usable without the crest. */ }
      sheet.mergeCells('B1:I1'); sheet.getCell('B1').value = "ST. DOROTHY'S SCHOOL MASSIMO"; sheet.getCell('B1').font = { bold: true, size: 16, color: { argb: '173662' } }
      sheet.mergeCells('B2:I2'); sheet.getCell('B2').value = 'P.O. Box 36-60607, Mikinduri, Meru, Kenya · +254 722 841 569 · stdorothysschoolmassimo@gmail.com'
      sheet.mergeCells('B3:I3'); sheet.getCell('B3').value = outstandingOnly ? 'Outstanding Fee Balances — Take-home Statement Register' : 'Fee Balances Register'
      sheet.mergeCells('B4:I4'); sheet.getCell('B4').value = `Generated: ${new Date().toLocaleString('en-KE')}`
      const headerRow = sheet.getRow(6); headerRow.values = ['No.', 'Learner name', 'Admission number', 'Class / Grade', 'Parent / Guardian', 'Parent phone', 'Fee due (KES)', 'Paid (KES)', 'Balance (KES)']; headerRow.font = { bold: true, color: { argb: 'FFFFFF' } }; headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1D5EAA' } }
      rows.forEach(({ learner, finance }, index) => sheet.addRow([index + 1, learner.learnerName, learner.admissionNumber, learner.placement?.className || learner.grade, learner.guardianName, learner.guardianPhone, finance.feeDue ?? '', finance.totalPaid ?? 0, finance.balance ?? '']))
      sheet.eachRow((row, number) => { if (number >= 6) row.alignment = { vertical: 'middle', wrapText: true } }); sheet.getColumn(7).numFmt = '#,##0'; sheet.getColumn(8).numFmt = '#,##0'; sheet.getColumn(9).numFmt = '#,##0'
      const bytes = await workbook.xlsx.writeBuffer(); response.writeHead(200, { 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'Content-Disposition': `attachment; filename="st_dorothy_${outstandingOnly ? 'outstanding_fee_statements' : 'fee_balances'}_${new Date().toISOString().slice(0, 10)}.xlsx"` }); return response.end(Buffer.from(bytes))
    }
    if (request.method === 'GET' && url.pathname === '/api/finance/statements-bulk') {
      const className = text(url.searchParams.get('className'), 80)
      const learners = data.admissions.filter(learner => { const currentClass = learner.placement?.className || learner.grade; const summary = currentFeeSummary(data, learner); return (!className || currentClass === className) && Number(summary.balance || 0) > 0 })
      if (!learners.length) return send(response, 404, { error: 'No learners with an outstanding balance were found for that grade or class.' })
      const workbook = new ExcelJS.Workbook(); const sheet = workbook.addWorksheet('Delivery Register'); sheet.columns = [{ width: 8 }, { width: 28 }, { width: 20 }, { width: 18 }, { width: 26 }, { width: 18 }, { width: 18 }, { width: 20 }]
      try { const crest = workbook.addImage({ buffer: await readFile(crestPath), extension: 'png' }); sheet.addImage(crest, { tl: { col: 0, row: 0 }, ext: { width: 58, height: 58 } }) } catch { /* The delivery register remains usable without the crest. */ }
      sheet.mergeCells('B1:H1'); sheet.getCell('B1').value = "ST. DOROTHY'S SCHOOL MASSIMO"; sheet.getCell('B1').font = { bold: true, size: 16, color: { argb: '173662' } }; sheet.mergeCells('B2:H2'); sheet.getCell('B2').value = 'Fee Statement Delivery Register'; sheet.mergeCells('B3:H3'); sheet.getCell('B3').value = `Generated: ${new Date().toLocaleString('en-KE')}`
      sheet.getRow(5).values = ['No.', 'Learner name', 'Admission number', 'Class / Grade', 'Guardian', 'Phone', 'Balance (KES)', 'Issued / Signature']; const header = sheet.getRow(5); header.font = { bold: true, color: { argb: 'FFFFFF' } }; header.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1D5EAA' } }
      learners.forEach((learner, index) => { const finance = currentFeeSummary(data, learner); sheet.addRow([index + 1, learner.learnerName, learner.admissionNumber, learner.placement?.className || learner.grade, learner.guardianName, learner.guardianPhone, finance.balance, '']) }); sheet.getColumn(7).numFmt = '#,##0'; sheet.views = [{ state: 'frozen', ySplit: 5 }]
      const deliveryRegister = await workbook.xlsx.writeBuffer(); const zipName = `st_dorothy_fee_statements_${safePdfName(className || 'all_classes')}_${new Date().toISOString().slice(0, 10)}.zip`
      response.writeHead(200, { 'Content-Type': 'application/zip', 'Content-Disposition': `attachment; filename="${zipName}"` }); const archive = new ZipArchive({ zlib: { level: 9 } }); archive.on('error', error => response.destroy(error)); archive.pipe(response)
      for (const learner of learners) archive.append(await createFeeStatementPdf(data, learner), { name: `fee-statements/${safePdfName(learner.admissionNumber || learner.id)}_${safePdfName(learner.learnerName)}.pdf` })
      archive.append(Buffer.from(deliveryRegister), { name: 'bursar-delivery-register.xlsx' }); await archive.finalize(); return
    }
    if (request.method === 'GET' && /^\/api\/finance\/statements\/[^/]+\/pdf$/.test(url.pathname)) {
      const admissionId = decodeURIComponent(url.pathname.split('/')[4]); const learner = data.admissions.find(item => item.id === admissionId)
      if (!learner) return send(response, 404, { error: 'Learner account not found.' })
      const pdf = await createFeeStatementPdf(data, learner); response.writeHead(200, { 'Content-Type': 'application/pdf', 'Content-Disposition': `inline; filename="st_dorothy_fee_statement_${safePdfName(learner.admissionNumber || learner.id)}.pdf"` }); return response.end(pdf)
    }
    if (request.method === 'GET' && /^\/api\/finance\/statements\/[^/]+$/.test(url.pathname)) {
      const admissionId = decodeURIComponent(url.pathname.split('/')[4])
      const learner = data.admissions.find(item => item.id === admissionId)
      if (!learner) return send(response, 404, { error: 'Learner account not found.' })
      const finance = currentFeeSummary(data, learner)
      const receipts = data.receipts.filter(item => item.admissionId === learner.id).slice().sort((a, b) => String(a.issuedAt).localeCompare(String(b.issuedAt)))
      return send(response, 200, { statementNumber: `SD/STM/${new Date().getFullYear()}/${learner.admissionNumber || learner.id}`, generatedAt: new Date().toISOString(), learner: { learnerName: learner.learnerName, admissionNumber: learner.admissionNumber, grade: learner.grade, className: learner.placement?.className || learner.grade, guardianName: learner.guardianName, guardianPhone: learner.guardianPhone }, finance, receipts })
    }
    if (request.method === 'GET' && url.pathname === '/api/finance/receipts') return send(response, 200, data.receipts)
    if (request.method === 'GET' && url.pathname === '/api/finance/mpesa-transactions') return send(response, 200, data.mpesaTransactions)
    if (request.method === 'GET' && url.pathname === '/api/finance/reports') {
      const start = text(url.searchParams.get('start'), 10)
      const end = text(url.searchParams.get('end'), 10)
      const isInPeriod = receipt => {
        const date = receipt.issuedAt.slice(0, 10)
        return (!start || date >= start) && (!end || date <= end)
      }
      const receipts = data.receipts.filter(isInPeriod)
      const totalCollected = receipts.reduce((total, receipt) => total + Number(receipt.amount || 0), 0)
      const byMethod = ['M-Pesa', 'Bank', 'Cash'].map(method => {
        const methodReceipts = receipts.filter(receipt => receipt.paymentMethod === method)
        return { method, count: methodReceipts.length, total: methodReceipts.reduce((total, receipt) => total + Number(receipt.amount || 0), 0) }
      })
      const classMap = new Map()
      for (const receipt of receipts) {
        const current = classMap.get(receipt.className) || { className: receipt.className, count: 0, total: 0 }
        current.count += 1
        current.total += Number(receipt.amount || 0)
        classMap.set(receipt.className, current)
      }
      const unmatchedMpesa = data.mpesaTransactions.filter(transaction => transaction.status !== 'Matched' && (!start || transaction.receivedAt.slice(0, 10) >= start) && (!end || transaction.receivedAt.slice(0, 10) <= end))
      return send(response, 200, {
        start: start || null,
        end: end || null,
        totalCollected,
        receiptCount: receipts.length,
        byMethod,
        byClass: [...classMap.values()].sort((a, b) => b.total - a.total),
        unmatchedMpesaCount: unmatchedMpesa.length,
        unmatchedMpesaTotal: unmatchedMpesa.reduce((total, transaction) => total + Number(transaction.amount || 0), 0),
        recentReceipts: receipts.slice(0, 8)
      })
    }
    if (request.method === 'GET' && url.pathname === '/api/finance/report-workbook') {
      const workbook = await readFile(reportWorkbookFile)
      response.writeHead(200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': 'attachment; filename="st_dorothy_financial_report.xlsx"'
      })
      return response.end(workbook)
    }

    if (request.method === 'POST' && url.pathname === '/api/school-classes') {
      const input = await body(request)
      const grade = text(input.grade, 20)
      const stream = text(input.stream, 30)
      const academicYear = text(input.academicYear, 4)
      const capacity = Number(input.capacity)
      const classTeacher = text(input.classTeacher, 100)
      const status = text(input.status, 20) || 'Active'
      if (!schoolGrades.has(grade) || !stream || !/^\d{4}$/.test(academicYear)) return send(response, 400, { error: 'Choose a valid grade, stream, and four-digit academic year.' })
      if (!Number.isInteger(capacity) || capacity < 1 || capacity > 70) return send(response, 400, { error: 'Class capacity must be between 1 and 70 learners.' })
      if (!classStatuses.has(status)) return send(response, 400, { error: 'Choose an active or inactive class status.' })
      const className = `${grade} ${stream}`
      if (data.schoolClasses.some(item => item.className.toLowerCase() === className.toLowerCase() && item.academicYear === academicYear)) return send(response, 400, { error: `${className} already exists for ${academicYear}.` })
      const schoolClass = { id: `class-${Date.now()}`, className, grade, stream, academicYear, capacity, classTeacher, status, createdAt: new Date().toISOString(), createdBy: 'Administrator' }
      data.schoolClasses.push(schoolClass)
      await saveData(data)
      return send(response, 201, schoolClass)
    }

    if (request.method === 'PUT' && url.pathname.startsWith('/api/school-classes/')) {
      const id = decodeURIComponent(url.pathname.slice('/api/school-classes/'.length))
      const existing = data.schoolClasses.find(item => item.id === id)
      if (!existing) return send(response, 404, { error: 'Class entry not found.' })
      const input = await body(request)
      const grade = text(input.grade, 20)
      const stream = text(input.stream, 30)
      const academicYear = text(input.academicYear, 4)
      const capacity = Number(input.capacity)
      const classTeacher = text(input.classTeacher, 100)
      const status = text(input.status, 20) || 'Active'
      const className = `${grade} ${stream}`
      if (!schoolGrades.has(grade) || !stream || !/^\d{4}$/.test(academicYear)) return send(response, 400, { error: 'Choose a valid grade, stream, and four-digit academic year.' })
      if (!Number.isInteger(capacity) || capacity < 1 || capacity > 70) return send(response, 400, { error: 'Class capacity must be between 1 and 70 learners.' })
      if (!classStatuses.has(status)) return send(response, 400, { error: 'Choose an active or inactive class status.' })
      if (data.schoolClasses.some(item => item.id !== id && item.className.toLowerCase() === className.toLowerCase() && item.academicYear === academicYear)) return send(response, 400, { error: `${className} already exists for ${academicYear}.` })
      const hasPlacements = data.classPlacements.some(item => item.className === existing.className && item.academicYear === existing.academicYear)
      const occupied = data.classPlacements.filter(item => item.className === existing.className && item.academicYear === existing.academicYear && item.decision !== 'Waitlisted').length
      if (hasPlacements && (grade !== existing.grade || academicYear !== existing.academicYear)) return send(response, 400, { error: 'This class has learner placements. Keep its grade and academic year unchanged, or create a new class entry.' })
      if (capacity < occupied) return send(response, 400, { error: `Capacity cannot be lower than the ${occupied} learner placement(s) already assigned.` })
      const updated = { ...existing, className, grade, stream, academicYear, capacity, classTeacher, status, updatedAt: new Date().toISOString(), updatedBy: 'Headteacher' }
      data.schoolClasses = data.schoolClasses.map(item => item.id === id ? updated : item)
      if (className !== existing.className) {
        data.classPlacements = data.classPlacements.map(item => item.className === existing.className && item.academicYear === existing.academicYear ? { ...item, className, updatedAt: new Date().toISOString() } : item)
        data.admissions = data.admissions.map(application => application.placement?.className === existing.className && application.placement?.academicYear === existing.academicYear ? { ...application, placement: { ...application.placement, className, updatedAt: new Date().toISOString() }, status: application.status.replace(existing.className, className), updatedAt: new Date().toISOString() } : application)
        data.teacherAssignments = data.teacherAssignments.map(item => item.classId === existing.id ? { ...item, className } : item)
        data.attendanceRecords = data.attendanceRecords.map(item => item.className === existing.className ? { ...item, className } : item)
        data.assessments = data.assessments.map(item => item.className === existing.className ? { ...item, className } : item)
      }
      await saveData(data)
      return send(response, 200, updated)
    }

    if (request.method === 'DELETE' && url.pathname.startsWith('/api/school-classes/')) {
      const id = decodeURIComponent(url.pathname.slice('/api/school-classes/'.length))
      const existing = data.schoolClasses.find(item => item.id === id)
      if (!existing) return send(response, 404, { error: 'Class entry not found.' })
      const hasPlacements = data.classPlacements.some(item => item.className === existing.className && item.academicYear === existing.academicYear)
      const hasAssignments = data.teacherAssignments.some(item => item.classId === existing.id)
      if (hasPlacements || hasAssignments) return send(response, 400, { error: 'This class cannot be deleted because it has learner placements or teacher assignments. Mark it Inactive instead.' })
      data.schoolClasses = data.schoolClasses.filter(item => item.id !== id)
      await saveData(data)
      return send(response, 200, { id })
    }

    if (request.method === 'POST' && url.pathname === '/api/documents') {
      const input = await body(request)
      const learner = data.admissions.find(item => item.id === text(input.admissionId, 80))
      const documentType = text(input.documentType, 80)
      const receivedDate = text(input.receivedDate, 10)
      if (!learner) return send(response, 400, { error: 'Select a valid learner admission record.' })
      if (!documentTypes.has(documentType)) return send(response, 400, { error: 'Select a valid learner document type.' })
      if (!/^\d{4}-\d{2}-\d{2}$/.test(receivedDate)) return send(response, 400, { error: 'Enter the date the document was received.' })
      const upload = decodeDocumentContent(input.content)
      const id = `document-${Date.now()}`
      const storedFileName = `${id}${documentExtension(upload.mimeType)}`
      await mkdir(documentUploadDirectory, { recursive: true })
      await writeFile(join(documentUploadDirectory, storedFileName), upload.content)
      const record = { id, admissionId: learner.id, documentType, fileName: safeFileName(input.fileName), receivedDate, status: 'Received', mimeType: upload.mimeType, storedFileName, uploadedAt: new Date().toISOString() }
      data.documentRecords.push(record)
      await saveData(data)
      return send(response, 201, documentForClient(data, record))
    }

    if (request.method === 'POST' && url.pathname === '/api/internal-messages') {
      const input = await body(request)
      const senderRole = request.account.role === 'teacher' ? 'Teacher' : 'Headteacher'
      const senderName = request.account.name
      const message = { id: `message-${Date.now()}`, senderName, senderRole, recipient: text(input.recipient, 80), subject: text(input.subject, 120), body: text(input.body, 1000), replyToId: text(input.replyToId, 80), status: 'Sent', createdAt: new Date().toISOString() }
      if (!message.recipient || !message.subject || !message.body) return send(response, 400, { error: 'Recipient, subject, and message are required.' })
      if (request.account.role === 'teacher' && message.recipient !== 'Headteacher') return send(response, 403, { error: 'Teachers can send internal messages only to the Headteacher.' })
      data.internalMessages.unshift(message)
      await saveData(data)
      return send(response, 201, message)
    }

    if (request.method === 'POST' && url.pathname === '/api/parent/messages') {
      const input = await body(request)
      const parentPhone = normalisePhone(input.parentPhone)
      const admissionId = text(input.admissionId, 80)
      const learner = data.admissions.find(application => application.id === admissionId)
      const recipient = text(input.recipient, 40)
      const subject = text(input.subject, 120)
      const bodyText = text(input.body, 1000)
      if (parentPhone !== request.account.parentPhone) return send(response, 403, { error: 'You can only send messages from your own parent account.' })
      if (!learner || normalisePhone(learner.guardianPhone) !== parentPhone) return send(response, 403, { error: 'Choose one of the children linked to your parent account.' })
      if (!['School office', 'Class teacher'].includes(recipient) || !subject || !bodyText) return send(response, 400, { error: 'Choose a recipient and complete both the subject and message.' })
      const message = { id: `parent-message-${Date.now()}`, parentPhone, guardianName: learner.guardianName, admissionId: learner.id, admissionNumber: learner.admissionNumber, learnerName: learner.learnerName, className: learner.placement?.className || learner.grade, recipient, subject, body: bodyText, status: 'New', createdAt: new Date().toISOString() }
      data.parentMessages.unshift(message)
      await saveData(data)
      return send(response, 201, message)
    }

    if (request.method === 'POST' && url.pathname === '/api/notices') {
      const input = await body(request)
      const notice = { id: `notice-${Date.now()}`, title: text(input.title, 120), message: text(input.message, 500), audience: text(input.audience, 60), targetGrade: text(input.targetGrade, 20), publishDate: text(input.publishDate, 10), status: text(input.status, 20), showOnWebsite: Boolean(input.showOnWebsite), createdAt: new Date().toISOString(), updatedAt: new Date().toISOString() }
      if (!notice.title || !notice.message || !noticeAudiences.has(notice.audience) || !/^\d{4}-\d{2}-\d{2}$/.test(notice.publishDate) || !noticeStatuses.has(notice.status)) return send(response, 400, { error: 'Title, message, audience, date, and notice status are required.' })
      if (notice.audience === 'Specific grade parents and guardians' && !noticeTargetGrades.has(notice.targetGrade)) return send(response, 400, { error: 'Choose the grade that should receive this notice.' })
      if (notice.audience !== 'Specific grade parents and guardians') notice.targetGrade = ''
      if (notice.showOnWebsite && (notice.audience === 'All staff' || notice.audience === 'Specific grade parents and guardians')) return send(response, 400, { error: 'Staff-only or grade-specific notices cannot be shown on the public school website.' })
      data.notices.unshift(notice)
      await saveData(data)
      return send(response, 201, notice)
    }

    if (request.method === 'POST' && url.pathname === '/api/staff') {
      const input = await body(request)
      const fullName = text(input.fullName, 100)
      const role = text(input.role, 40)
      const phone = normalisePhone(input.phone)
      if (!fullName || !staffRoles.has(role)) return send(response, 400, { error: 'Staff name and a valid role are required.' })
      const staffMember = { id: `staff-${Date.now()}`, fullName, role, phone, email: text(input.email, 100), employmentStatus: text(input.employmentStatus, 30) || 'Active', createdAt: new Date().toISOString() }
      data.staff.push(staffMember)
      await saveData(data)
      return send(response, 201, staffMember)
    }

    if (request.method === 'POST' && url.pathname === '/api/subjects') {
      const input = await body(request)
      const name = text(input.name, 70)
      const grade = text(input.grade, 20)
      if (!name || !schoolGrades.has(grade)) return send(response, 400, { error: 'Subject name and grade are required.' })
      if (data.subjects.some(item => item.name.toLowerCase() === name.toLowerCase() && item.grade === grade)) return send(response, 400, { error: 'That subject has already been added for this grade.' })
      const subject = { id: `subject-${Date.now()}`, name, grade, createdAt: new Date().toISOString() }
      data.subjects.push(subject)
      await saveData(data)
      return send(response, 201, subject)
    }

    if (request.method === 'POST' && url.pathname === '/api/teacher-assignments') {
      const input = await body(request)
      const staff = data.staff.find(item => item.id === text(input.staffId, 80))
      const schoolClass = data.schoolClasses.find(item => item.id === text(input.classId, 80))
      const subject = data.subjects.find(item => item.id === text(input.subjectId, 80))
      if (!staff || !schoolClass || !subject) return send(response, 400, { error: 'Choose a staff member, class, and subject.' })
      const assignment = { id: `assignment-${Date.now()}`, staffId: staff.id, staffName: staff.fullName, classId: schoolClass.id, className: schoolClass.className, subjectId: subject.id, subjectName: subject.name, academicYear: schoolClass.academicYear, createdAt: new Date().toISOString() }
      data.teacherAssignments.push(assignment)
      await saveData(data)
      return send(response, 201, assignment)
    }

    if (request.method === 'POST' && url.pathname === '/api/timetable') {
      const input = await body(request)
      const assignment = data.teacherAssignments.find(item => item.id === text(input.assignmentId, 80))
      const day = text(input.day, 15)
      const time = text(input.time, 20)
      if (!assignment || !day || !time) return send(response, 400, { error: 'Choose a teacher assignment, day, and lesson time.' })
      const entry = { id: `timetable-${Date.now()}`, assignmentId: assignment.id, className: assignment.className, subjectName: assignment.subjectName, staffName: assignment.staffName, day, time, room: text(input.room, 50), createdAt: new Date().toISOString() }
      data.timetableEntries.push(entry)
      await saveData(data)
      return send(response, 201, entry)
    }

    if (request.method === 'POST' && url.pathname === '/api/assessments') {
      const input = await body(request)
      const className = text(input.className, 60)
      const subjectName = text(input.subjectName, 70)
      const title = text(input.title, 100)
      const outOf = Number(input.outOf)
      if (!className || !subjectName || !title || !Number.isFinite(outOf) || outOf <= 0) return send(response, 400, { error: 'Class, subject, assessment title, and maximum score are required.' })
      const assessment = { id: `assessment-${Date.now()}`, className, subjectName, title, outOf, term: text(input.term, 20) || 'Term 3', academicYear: text(input.academicYear, 10) || '2026', createdAt: new Date().toISOString() }
      data.assessments.push(assessment)
      await saveData(data)
      return send(response, 201, assessment)
    }

    if (request.method === 'POST' && url.pathname === '/api/assessment-marks') {
      const input = await body(request)
      const assessment = data.assessments.find(item => item.id === text(input.assessmentId, 80))
      const marks = Array.isArray(input.marks) ? input.marks : []
      if (!assessment || !marks.length) return send(response, 400, { error: 'Choose an assessment and enter at least one mark.' })
      const timestamp = new Date().toISOString()
      const updated = marks.map(item => {
        const learner = data.admissions.find(application => application.id === text(item.admissionId, 80))
        const score = Number(item.score)
        if (!learner || !Number.isFinite(score) || score < 0 || score > assessment.outOf) return null
        return { id: `mark-${assessment.id}-${learner.id}`, assessmentId: assessment.id, admissionId: learner.id, learnerName: learner.learnerName, className: assessment.className, subjectName: assessment.subjectName, score, outOf: assessment.outOf, updatedAt: timestamp }
      }).filter(Boolean)
      if (!updated.length) return send(response, 400, { error: 'Marks must be between zero and the assessment maximum.' })
      data.assessmentMarks = [...data.assessmentMarks.filter(item => !updated.some(mark => mark.id === item.id)), ...updated]
      await saveData(data)
      return send(response, 200, updated)
    }

    if (request.method === 'POST' && url.pathname === '/api/attendance') {
      const input = await body(request)
      const className = text(input.className, 60)
      const date = text(input.date, 10)
      const entries = Array.isArray(input.entries) ? input.entries : []
      if (!className || !/^\d{4}-\d{2}-\d{2}$/.test(date) || !entries.length) return send(response, 400, { error: 'Class, date, and learner attendance are required.' })
      const timestamp = new Date().toISOString()
      const records = entries.map(entry => {
        const learner = data.admissions.find(application => application.id === text(entry.admissionId, 80) && (application.placement?.className || application.grade) === className)
        const status = text(entry.status, 20)
        if (!learner || !attendanceStatuses.has(status)) return null
        return { id: `attendance-${date}-${learner.id}`, admissionId: learner.id, learnerName: learner.learnerName, className, date, status, reason: text(entry.reason, 120), markedAt: timestamp, markedBy: 'Teacher' }
      }).filter(Boolean)
      if (!records.length) return send(response, 400, { error: 'No valid learners were supplied for this class.' })
      data.attendanceRecords = [...data.attendanceRecords.filter(item => !records.some(record => record.id === item.id)), ...records]
      await saveData(data)
      return send(response, 200, records)
    }

    if (request.method === 'POST' && url.pathname === '/api/lesson-plans') {
      const input = await body(request)
      const className = text(input.className, 60)
      const subjectName = text(input.subjectName, 80)
      const lessonDate = text(input.lessonDate, 10)
      const topic = text(input.topic, 160)
      const learningOutcomes = text(input.learningOutcomes, 1200)
      if (!className || !subjectName || !/^\d{4}-\d{2}-\d{2}$/.test(lessonDate) || !topic || !learningOutcomes) return send(response, 400, { error: 'Class, subject, lesson date, topic, and learning outcomes are required.' })
      const now = new Date().toISOString()
      const lessonPlan = {
        id: `lesson-plan-${Date.now()}`,
        className,
        subjectName,
        week: text(input.week, 30) || 'Week not set',
        lessonDate,
        lessonTime: text(input.lessonTime, 40),
        topic,
        strand: text(input.strand, 160),
        subStrand: text(input.subStrand, 160),
        learningOutcomes,
        learningActivities: text(input.learningActivities, 1600),
        resources: text(input.resources, 600),
        assessment: text(input.assessment, 900),
        reflection: text(input.reflection, 900),
        status: ['Draft', 'Ready for review', 'Published'].includes(text(input.status, 30)) ? text(input.status, 30) : 'Draft',
        preparedBy: request.account.name,
        createdAt: now,
        updatedAt: now
      }
      data.lessonPlans.unshift(lessonPlan)
      await saveData(data)
      return send(response, 201, lessonPlan)
    }

    if (request.method === 'PUT' && url.pathname.startsWith('/api/lesson-plans/')) {
      const id = decodeURIComponent(url.pathname.slice('/api/lesson-plans/'.length))
      const existing = data.lessonPlans.find(item => item.id === id)
      if (!existing) return send(response, 404, { error: 'Lesson plan not found.' })
      if (request.account.role === 'teacher' && existing.preparedBy !== request.account.name) return send(response, 403, { error: 'You can only edit your own lesson plans.' })
      const input = await body(request)
      const className = text(input.className, 60)
      const subjectName = text(input.subjectName, 80)
      const lessonDate = text(input.lessonDate, 10)
      const topic = text(input.topic, 160)
      const learningOutcomes = text(input.learningOutcomes, 1200)
      if (!className || !subjectName || !/^\d{4}-\d{2}-\d{2}$/.test(lessonDate) || !topic || !learningOutcomes) return send(response, 400, { error: 'Class, subject, lesson date, topic, and learning outcomes are required.' })
      const updated = {
        ...existing,
        className,
        subjectName,
        week: text(input.week, 30) || 'Week not set',
        lessonDate,
        lessonTime: text(input.lessonTime, 40),
        topic,
        strand: text(input.strand, 160),
        subStrand: text(input.subStrand, 160),
        learningOutcomes,
        learningActivities: text(input.learningActivities, 1600),
        resources: text(input.resources, 600),
        assessment: text(input.assessment, 900),
        reflection: text(input.reflection, 900),
        status: ['Draft', 'Ready for review', 'Published'].includes(text(input.status, 30)) ? text(input.status, 30) : 'Draft',
        updatedAt: new Date().toISOString()
      }
      data.lessonPlans = data.lessonPlans.map(item => item.id === id ? updated : item)
      await saveData(data)
      return send(response, 200, updated)
    }

    if (request.method === 'POST' && url.pathname === '/api/fee-structures') {
      const input = await body(request)
      const grade = text(input.grade, 20)
      const academicYear = text(input.academicYear, 10)
      const term = text(input.term, 20)
      const amount = Number(input.amount)
      if (!schoolGrades.has(grade) || !/^\d{4}$/.test(academicYear) || !term || !Number.isFinite(amount) || amount < 0) return send(response, 400, { error: 'Choose a grade, academic year, term, and valid fee amount.' })
      const structure = { id: `fees-${Date.now()}`, grade, academicYear, term, amount, updatedAt: new Date().toISOString(), setBy: 'Bursar' }
      data.feeStructures = [...data.feeStructures.filter(item => !(item.grade === grade && item.term === term && item.academicYear === academicYear)), structure]
      await saveData(data)
      return send(response, 201, structure)
    }

    if (request.method === 'POST' && url.pathname === '/api/leave-requests') {
      const input = await body(request)
      const staff = data.staff.find(item => item.id === text(input.staffId, 80))
      const startDate = text(input.startDate, 10)
      const endDate = text(input.endDate, 10)
      if (!staff || !startDate || !endDate || endDate < startDate) return send(response, 400, { error: 'Choose a staff member and a valid leave period.' })
      const requestItem = { id: `leave-${Date.now()}`, staffId: staff.id, staffName: staff.fullName, leaveType: text(input.leaveType, 30) || 'Annual leave', startDate, endDate, reason: text(input.reason, 200), status: 'Pending', createdAt: new Date().toISOString() }
      data.leaveRequests.unshift(requestItem)
      await saveData(data)
      return send(response, 201, requestItem)
    }

    if (request.method === 'POST' && url.pathname === '/api/staff-attendance') {
      const input = await body(request)
      const staff = data.staff.find(item => item.id === text(input.staffId, 80))
      const date = text(input.date, 10)
      const status = text(input.status, 20)
      if (!staff || !/^\d{4}-\d{2}-\d{2}$/.test(date) || !attendanceStatuses.has(status)) return send(response, 400, { error: 'Choose a staff member, date, and valid attendance status.' })
      const record = { id: `staff-attendance-${date}-${staff.id}`, staffId: staff.id, staffName: staff.fullName, date, status, markedAt: new Date().toISOString(), markedBy: 'Administrator' }
      data.staffAttendance = [...data.staffAttendance.filter(item => item.id !== record.id), record]
      await saveData(data)
      return send(response, 200, record)
    }

    if (request.method === 'POST' && url.pathname === '/api/promotions') {
      const input = await body(request)
      const learner = data.admissions.find(item => item.id === text(input.admissionId, 80))
      const nextClass = data.schoolClasses.find(item => item.id === text(input.nextClassId, 80) && item.status === 'Active')
      if (!learner || !nextClass) return send(response, 400, { error: 'Choose a learner and an active receiving class.' })
      const promotion = { id: `promotion-${Date.now()}`, admissionId: learner.id, learnerName: learner.learnerName, admissionNumber: learner.admissionNumber, fromClass: learner.placement?.className || learner.grade, toClass: nextClass.className, academicYear: nextClass.academicYear, decision: text(input.decision, 30) || 'Approved', notes: text(input.notes, 300), promotedAt: new Date().toISOString(), approvedBy: 'Headteacher' }
      data.promotions = [...data.promotions.filter(item => item.admissionId !== learner.id || item.academicYear !== promotion.academicYear), promotion]
      await saveData(data)
      return send(response, 201, promotion)
    }

    if (request.method === 'POST' && url.pathname === '/api/calendar-events') {
      const input = await body(request)
      const event = { id: `event-${Date.now()}`, title: text(input.title, 100), date: text(input.date, 10), detail: text(input.detail, 120) }
      if (!event.title || !/^\d{4}-\d{2}-\d{2}$/.test(event.date) || !event.detail) return send(response, 400, { error: 'Title, valid date, and time/location are required.' })
      data.calendarEvents.push(event)
      data.calendarEvents.sort((a, b) => a.date.localeCompare(b.date))
      await saveData(data)
      return send(response, 201, event)
    }

    if (request.method === 'POST' && url.pathname === '/api/sms') {
      const input = await body(request)
      const message = { id: `sms-${Date.now()}`, audience: text(input.audience, 80), count: Number(input.count) || 0, text: text(input.text, 160), status: 'Prepared', createdAt: new Date().toISOString() }
      if (!message.audience || !message.text || message.count < 1) return send(response, 400, { error: 'Audience, recipient count, and message are required.' })
      data.smsMessages.unshift(message)
      await saveData(data)
      return send(response, 201, message)
    }

    if (request.method === 'POST' && url.pathname === '/api/admissions') {
      const input = await body(request)
      const createdAt = new Date().toISOString()
      const application = {
        id: `admission-${Date.now()}`,
        admissionNumber: nextAdmissionNumber(createdAt, data.admissions),
        learnerName: text(input.learnerName, 100),
        grade: text(input.grade, 20),
        guardianName: text(input.guardianName, 100),
        guardianPhone: text(input.guardianPhone, 20),
        relationship: text(input.relationship, 40),
        formerSchool: text(input.formerSchool, 100),
        medicalNotes: text(input.medicalNotes, 500),
        birthCertificateProvided: Boolean(input.birthCertificateProvided),
        status: 'Application received',
        createdAt
      }
      if (!application.learnerName || !application.grade || !application.guardianName || !application.guardianPhone) return send(response, 400, { error: 'Learner name, applying grade, guardian name, and guardian phone are required.' })
      data.admissions.unshift(application)
      await saveData(data)
      return send(response, 201, application)
    }

    if (request.method === 'POST' && url.pathname === '/api/finance/receipts') {
      const input = await body(request)
      const admissionId = text(input.admissionId, 80)
      const learner = data.admissions.find(application => application.id === admissionId)
      const amount = Number(input.amount)
      const paymentMethod = text(input.paymentMethod, 30)
      const paymentReference = text(input.paymentReference, 80)
      if (!learner) return send(response, 404, { error: 'Select a valid learner account before issuing the receipt.' })
      if (!Number.isFinite(amount) || amount <= 0) return send(response, 400, { error: 'Enter a valid payment amount.' })
      if (!paymentMethod || !paymentReference) return send(response, 400, { error: 'Payment method and reference are required.' })
      const issuedAt = new Date().toISOString()
      const receipt = {
        id: `receipt-${Date.now()}`,
        receiptNumber: nextReceiptNumber(issuedAt, data.receipts),
        admissionId: learner.id,
        admissionNumber: learner.admissionNumber,
        learnerName: learner.learnerName,
        className: learner.placement?.className || learner.grade,
        guardianName: learner.guardianName,
        guardianPhone: learner.guardianPhone,
        amount,
        paymentMethod,
        paymentReference,
        issuedAt
      }
      data.receipts.unshift(receipt)
      await saveData(data)
      return send(response, 201, receipt)
    }

    if (request.method === 'POST' && url.pathname === '/api/finance/mpesa-transactions') {
      const input = await body(request)
      const transactionCode = text(input.transactionCode, 30).toUpperCase()
      const senderPhone = normalisePhone(input.senderPhone)
      const amount = Number(input.amount)
      if (!transactionCode || !senderPhone || !Number.isFinite(amount) || amount <= 0) return send(response, 400, { error: 'M-Pesa code, sender phone, and valid amount are required.' })
      if (data.mpesaTransactions.some(transaction => transaction.transactionCode === transactionCode)) return send(response, 400, { error: 'That M-Pesa transaction code has already been recorded.' })
      const transaction = { id: `mpesa-${Date.now()}`, transactionCode, senderPhone, amount, receivedAt: new Date().toISOString(), status: 'Awaiting match' }
      data.mpesaTransactions.unshift(transaction)
      await saveData(data)
      return send(response, 201, transaction)
    }

    if (request.method === 'POST' && url.pathname === '/api/parent/payment-reports') {
      const input = await body(request)
      const admissionId = text(input.admissionId, 80)
      const parentPhone = normalisePhone(input.parentPhone)
      const learner = data.admissions.find(application => application.id === admissionId)
      const transactionCode = text(input.transactionCode, 30).toUpperCase()
      const senderPhone = normalisePhone(input.senderPhone)
      const amount = Number(input.amount)
      if (parentPhone !== request.account.parentPhone) return send(response, 403, { error: 'You can only report payments from your own parent account.' })
      if (!learner || normalisePhone(learner.guardianPhone) !== parentPhone) return send(response, 403, { error: 'This learner is not linked to the signed-in guardian account.' })
      if (!transactionCode || !senderPhone || !Number.isFinite(amount) || amount <= 0) return send(response, 400, { error: 'Enter the M-Pesa code, sender number, and amount paid.' })
      if (data.mpesaTransactions.some(transaction => transaction.transactionCode === transactionCode)) return send(response, 400, { error: 'That M-Pesa transaction code has already been reported or reconciled.' })
      const reportedAt = new Date().toISOString()
      const transaction = { id: `mpesa-${Date.now()}`, transactionCode, senderPhone, amount, receivedAt: reportedAt, reportedAt, source: 'Parent report', status: 'Parent reported', admissionId: learner.id, admissionNumber: learner.admissionNumber, learnerName: learner.learnerName, className: learner.placement?.className || learner.grade, guardianName: learner.guardianName }
      data.mpesaTransactions.unshift(transaction)
      await saveData(data)
      return send(response, 201, transaction)
    }

    if (request.method === 'PUT' && url.pathname.startsWith('/api/finance/mpesa-transactions/')) {
      const id = decodeURIComponent(url.pathname.slice('/api/finance/mpesa-transactions/'.length))
      const transaction = data.mpesaTransactions.find(item => item.id === id)
      if (!transaction) return send(response, 404, { error: 'M-Pesa transaction not found.' })
      if (transaction.status === 'Matched') return send(response, 400, { error: 'This M-Pesa transaction has already been reconciled.' })
      const input = await body(request)
      const learner = data.admissions.find(application => application.id === text(input.admissionId, 80))
      if (!learner) return send(response, 404, { error: 'Select a valid learner account to reconcile this payment.' })
      const reconciledAt = new Date().toISOString()
      const updatedTransaction = { ...transaction, status: 'Matched', admissionId: learner.id, admissionNumber: learner.admissionNumber, learnerName: learner.learnerName, className: learner.placement?.className || learner.grade, guardianName: learner.guardianName, reconciledAt, reconciledBy: 'Bursar' }
      const receipt = {
        id: `receipt-${Date.now()}`,
        receiptNumber: nextReceiptNumber(reconciledAt, data.receipts),
        mpesaTransactionId: transaction.id,
        admissionId: learner.id,
        admissionNumber: learner.admissionNumber,
        learnerName: learner.learnerName,
        className: learner.placement?.className || learner.grade,
        guardianName: learner.guardianName,
        guardianPhone: learner.guardianPhone,
        amount: transaction.amount,
        paymentMethod: 'M-Pesa',
        paymentReference: transaction.transactionCode,
        issuedAt: reconciledAt
      }
      data.mpesaTransactions = data.mpesaTransactions.map(item => item.id === id ? updatedTransaction : item)
      data.receipts.unshift(receipt)
      await saveData(data)
      return send(response, 200, { transaction: updatedTransaction, receipt })
    }

    if (request.method === 'POST' && url.pathname === '/api/class-placements') {
      const input = await body(request)
      const admissionId = text(input.admissionId, 80)
      const classId = text(input.classId, 80)
      const decision = text(input.decision, 40)
      const admission = data.admissions.find(application => application.id === admissionId)
      if (!admission) return send(response, 404, { error: 'The selected admission application was not found.' })
      const academicYear = text(input.academicYear, 10)
      const schoolClass = data.schoolClasses.find(item => item.id === classId && item.academicYear === academicYear)
      if (!schoolClass || schoolClass.status !== 'Active') return send(response, 400, { error: 'Select an active class created for the chosen academic year.' })
      if (schoolClass.grade !== admission.grade) return send(response, 400, { error: `Choose a class set up for ${admission.grade}.` })
      if (!placementDecisions.has(decision)) return send(response, 400, { error: 'Select a valid placement decision.' })
      const className = schoolClass.className
      const existing = data.classPlacements.find(placement => placement.admissionId === admissionId)
      const occupied = data.classPlacements.filter(placement => placement.className === className && placement.admissionId !== admissionId && placement.decision !== 'Waitlisted').length
      if (decision !== 'Waitlisted' && occupied >= schoolClass.capacity) return send(response, 400, { error: `${className} is full. Choose another stream or waitlist the learner.` })
      const timestamp = new Date().toISOString()
      const placement = {
        id: existing?.id || `placement-${Date.now()}`,
        admissionId,
        admissionNumber: admission.admissionNumber,
        learnerName: admission.learnerName,
        requestedGrade: admission.grade,
        className,
        academicYear,
        term: text(input.term, 20),
        decision,
        assessmentResult: text(input.assessmentResult, 120),
        notes: text(input.notes, 500),
        assignedBy: 'School office',
        assignedAt: existing?.assignedAt || timestamp,
        updatedAt: timestamp
      }
      if (!placement.academicYear || !placement.term) return send(response, 400, { error: 'Academic year and term are required.' })
      data.classPlacements = existing ? data.classPlacements.map(item => item.admissionId === admissionId ? placement : item) : [placement, ...data.classPlacements]
      const updatedAdmission = {
        ...admission,
        placement,
        status: decision === 'Confirmed' ? `Placement confirmed · ${className}` : `${decision} · ${className}`,
        updatedAt: timestamp
      }
      data.admissions = data.admissions.map(item => item.id === admissionId ? updatedAdmission : item)
      await saveData(data)
      return send(response, existing ? 200 : 201, { placement, admission: updatedAdmission, classCapacity: schoolClass.capacity, classes: data.schoolClasses })
    }

    if (request.method === 'PUT' && url.pathname.startsWith('/api/admissions/')) {
      const id = decodeURIComponent(url.pathname.slice('/api/admissions/'.length))
      const existing = data.admissions.find(application => application.id === id)
      if (!existing) return send(response, 404, { error: 'Admission application not found.' })
      const input = await body(request)
      const updated = {
        ...existing,
        learnerName: text(input.learnerName, 100),
        grade: text(input.grade, 20),
        guardianName: text(input.guardianName, 100),
        guardianPhone: text(input.guardianPhone, 20),
        relationship: text(input.relationship, 40),
        formerSchool: text(input.formerSchool, 100),
        medicalNotes: text(input.medicalNotes, 500),
        birthCertificateProvided: Boolean(input.birthCertificateProvided),
        updatedAt: new Date().toISOString()
      }
      if (!updated.learnerName || !updated.grade || !updated.guardianName || !updated.guardianPhone) return send(response, 400, { error: 'Learner name, applying grade, guardian name, and guardian phone are required.' })
      data.admissions = data.admissions.map(application => application.id === id ? updated : application)
      await saveData(data)
      return send(response, 200, updated)
    }

    if (request.method === 'PUT' && url.pathname.startsWith('/api/notices/')) {
      const id = decodeURIComponent(url.pathname.slice('/api/notices/'.length))
      const existing = data.notices.find(item => item.id === id)
      if (!existing) return send(response, 404, { error: 'School notice not found.' })
      const input = await body(request)
      const updated = { ...existing, title: text(input.title, 120), message: text(input.message, 500), audience: text(input.audience, 60), targetGrade: text(input.targetGrade, 20), publishDate: text(input.publishDate, 10), status: text(input.status, 20), showOnWebsite: Boolean(input.showOnWebsite), updatedAt: new Date().toISOString() }
      if (!updated.title || !updated.message || !noticeAudiences.has(updated.audience) || !/^\d{4}-\d{2}-\d{2}$/.test(updated.publishDate) || !noticeStatuses.has(updated.status)) return send(response, 400, { error: 'Title, message, audience, date, and notice status are required.' })
      if (updated.audience === 'Specific grade parents and guardians' && !noticeTargetGrades.has(updated.targetGrade)) return send(response, 400, { error: 'Choose the grade that should receive this notice.' })
      if (updated.audience !== 'Specific grade parents and guardians') updated.targetGrade = ''
      if (updated.showOnWebsite && (updated.audience === 'All staff' || updated.audience === 'Specific grade parents and guardians')) return send(response, 400, { error: 'Staff-only or grade-specific notices cannot be shown on the public school website.' })
      data.notices = data.notices.map(item => item.id === id ? updated : item)
      await saveData(data)
      return send(response, 200, updated)
    }

    if (request.method === 'POST' && url.pathname === '/api/admission-enquiries') {
      if (rateLimited(request, 'admission-enquiry', 10, 60_000)) return send(response, 429, { error: 'Too many admission enquiries. Please wait a moment and try again.' })
      const input = await body(request)
      const enquiry = { id: `admission-${Date.now()}`, guardianName: text(input.guardianName, 100), phone: text(input.phone, 20), grade: text(input.grade, 20), createdAt: new Date().toISOString() }
      if (!enquiry.guardianName || !enquiry.phone || !enquiry.grade) return send(response, 400, { error: 'Guardian name, phone number, and applying grade are required.' })
      data.admissionEnquiries.unshift(enquiry)
      await saveData(data)
      return send(response, 201, enquiry)
    }

    return send(response, 404, { error: 'Endpoint not found.' })
  } catch (error) {
    console.error(error)
    return send(response, 500, { error: 'The school portal server could not process this request.' })
  }
})

const host = process.env.HOST || '127.0.0.1'
server.listen(port, host, () => console.log(`St. Dorothy API running at http://${host}:${port}`))
