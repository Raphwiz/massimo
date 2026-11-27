import { createServer } from 'node:http'
import 'dotenv/config'
import { mkdir, readFile, stat, writeFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { randomBytes, randomInt, scryptSync, timingSafeEqual } from 'node:crypto'
import pg from 'pg'
import ExcelJS from 'exceljs'
import PDFDocument from 'pdfkit'
import { ZipArchive } from 'archiver'
import unzipper from 'unzipper'

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
const sessionIdleTimeoutMs = 10 * 60 * 1000
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
  smsWallet: { balance: 0, pricePerCreditKes: 0.8, senderId: '' },
  smsTopups: [],
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
  feeAdjustments: [],
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
  financeAudit: [],
  financeApprovals: [],
  websiteSettings: { socialLinks: { facebook: '', instagram: '', youtube: '', tiktok: '' }, gallery: [] },
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
  if (isProduction) return { ...structuredClone(seedData), calendarEvents: [], notices: [], accounts: [] }
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
function accountRoles(account) { return Array.from(new Set(Array.isArray(account?.roles) && account.roles.length ? account.roles : [account?.role].filter(Boolean))) }
function accountHasRole(account, role) { return accountRoles(account).includes(role) }

function requestedSessionRole(request) {
  const role = String(request.headers['x-portal-role'] || '').trim()
  return ['system-admin', 'headteacher', 'secretary', 'teacher', 'bursar', 'parent'].includes(role) ? role : ''
}

function createSession(response, account) {
  const token = randomBytes(32).toString('base64url')
  const now = Date.now()
  sessions.set(token, { accountId: account.id, expiresAt: now + sessionDurationMs, lastActiveAt: now })
  const cookies = accountRoles(account).map(role => { const cookie = [`${sessionCookieNameForRole(role)}=${token}`, 'HttpOnly', 'SameSite=Lax', 'Path=/', `Max-Age=${Math.floor(sessionDurationMs / 1000)}`]; if (isProduction) cookie.push('Secure'); return cookie.join('; ') })
  response.setHeader('Set-Cookie', cookies)
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
  const now = Date.now()
  if (!session || session.expiresAt <= now || now - session.lastActiveAt >= sessionIdleTimeoutMs) { if (token) sessions.delete(token); return null }
  const account = data.accounts.find(item => item.id === session.accountId && item.active !== false)
  if (!account || (role && !accountHasRole(account, role))) { sessions.delete(token); return null }
  session.lastActiveAt = now
  return account
}

function routeRoles(method, path) {
  if ((method === 'GET' && ['/api/health', '/api/public/notices', '/api/calendar-events'].includes(path)) || (method === 'POST' && (path === '/api/auth/login' || path === '/api/auth/activate' || path === '/api/admission-enquiries'))) return null
  if (method === 'POST' && path === '/api/auth/logout') return ['system-admin', 'headteacher', 'secretary', 'teacher', 'bursar', 'parent']
  if (method === 'POST' && path === '/api/auth/activity') return ['system-admin', 'headteacher', 'secretary', 'teacher', 'bursar', 'parent']
  if (method === 'POST' && path === '/api/auth/change-password') return ['system-admin', 'headteacher', 'secretary', 'teacher', 'bursar', 'parent']
  if (method === 'GET' && path === '/api/auth/session') return ['system-admin', 'headteacher', 'secretary', 'teacher', 'bursar', 'parent']
  if (path === '/api/system-admin/overview') return ['system-admin']
  if (path === '/api/system-admin/logs') return ['system-admin']
  if (path === '/api/system-admin/backup/status' || path === '/api/system-admin/backup/download' || path === '/api/system-admin/backup/restore') return ['system-admin']
  if (path === '/api/headteacher/dashboard') return ['headteacher']
  if (path === '/api/headteacher/academics') return ['headteacher']
  if (path === '/api/headteacher/learners') return ['headteacher']
  if (path === '/api/headteacher/attendance') return ['headteacher']
  if (path === '/api/teacher/dashboard') return ['teacher']
  if (path === '/api/system-admin/go-live') return ['system-admin']
  if (path === '/api/system-admin/website-settings') return ['system-admin']
  if (path === '/api/public/website-settings') return null
  if (path === '/api/accounts' || path.startsWith('/api/accounts/')) return ['system-admin']
  if (path === '/api/parent-portal') return ['system-admin']
  if (path === '/api/finance/approval-requests') return method === 'POST' ? ['bursar'] : ['bursar', 'headteacher']
  if (path.startsWith('/api/finance/approval-requests/')) return ['headteacher']
  // The secretary may confirm a payment only after receiving the school's payment notification.
  // Other finance operations remain restricted to the bursar.
  if (path === '/api/finance/learners' || path === '/api/finance/receipts' || /^\/api\/finance\/receipts\/[^/]+\/pdf$/.test(path)) return ['bursar', 'secretary']
  if (path === '/api/secretary/fee-lookup') return ['secretary']
  if (path === '/api/secretary/boarders/pdf') return ['secretary']
  if (path === '/api/secretary/learner-records' || path === '/api/secretary/learner-records/residency' || /^\/api\/secretary\/learner-records\/[^/]+\/pdf$/.test(path)) return ['secretary']
  if (path === '/api/finance/audit') return ['system-admin']
  if (path.startsWith('/api/finance/')) return ['bursar']
  if (path === '/api/parent/learners' || path === '/api/parent/payment-reports') return ['parent']
  if (path === '/api/parent/messages') return method === 'POST' ? ['parent'] : ['parent', 'teacher', 'secretary']
  if (path === '/api/internal-messages') return ['teacher', 'headteacher']
  if (path === '/api/notices') return method === 'GET' ? ['headteacher', 'secretary', 'teacher', 'bursar', 'parent'] : ['secretary']
  if (path.startsWith('/api/documents')) return ['secretary']
  if (path === '/api/sms' || path === '/api/sms/wallet' || path === '/api/sms/topups') return ['secretary']
  if (path === '/api/admissions') return method === 'GET' ? ['secretary', 'headteacher', 'teacher'] : ['secretary']
  if (path.startsWith('/api/admissions/')) return ['secretary']
  if (path.startsWith('/api/learner-import/') || path.startsWith('/api/teacher-import/')) return ['system-admin']
  if (path === '/api/calendar-events') return ['secretary']
  if (path === '/api/school-classes') return method === 'GET' ? ['headteacher', 'secretary', 'teacher'] : ['headteacher', 'secretary']
  if (path.startsWith('/api/school-classes/')) return method === 'PUT' ? ['headteacher', 'secretary'] : ['headteacher']
  if (path === '/api/class-placements' || path === '/api/class-placements/suggestions' || path === '/api/class-placements/auto-place') return ['headteacher', 'secretary']
  if (path === '/api/class-lists/workbook' || path === '/api/class-lists/pdf') return ['secretary', 'teacher']
  if (path === '/api/staff-list/workbook') return ['secretary']
  if (path === '/api/staff') return ['headteacher', 'secretary']
  if (path === '/api/staff-attendance' || path === '/api/promotions') return ['headteacher']
  if (path === '/api/academics') return ['headteacher', 'teacher']
  if (path === '/api/subjects' || path.startsWith('/api/subjects/') || ['/api/teacher-assignments', '/api/timetable'].includes(path)) return ['headteacher']
  if (path === '/api/lesson-plans') return ['headteacher', 'teacher']
  if (path.startsWith('/api/lesson-plans/')) return ['headteacher', 'teacher']
  if (path === '/api/leave-requests') return ['headteacher', 'teacher']
  if (path.startsWith('/api/leave-requests/')) return ['headteacher']
  if (path === '/api/attendance' || path === '/api/attendance/workbook' || path === '/api/assessments' || path === '/api/assessment-marks') return ['headteacher', 'teacher']
  if (path === '/api/fee-structures') return ['bursar']
  return []
}

function requireAccess(request, response, data, url) {
  const roles = routeRoles(request.method, url.pathname)
  if (roles === null) return true
  if (!roles.length) { send(response, 404, { error: 'Endpoint not found.' }); return false }
  const account = sessionAccount(request, data)
  if (!account) { clearSession(request, response); send(response, 401, { error: 'Please sign in to continue.' }); return false }
  if (!roles.some(role => accountHasRole(account, role))) { send(response, 403, { error: 'You do not have permission to access this resource.' }); return false }
  if (account.mustChangePassword && url.pathname !== '/api/auth/change-password' && url.pathname !== '/api/auth/logout') { send(response, 403, { error: 'Change your temporary password before accessing the portal.' }); return false }
  request.account = account
  return true
}

async function body(request) {
  let raw = ''
  let size = 0
  for await (const chunk of request) { size += chunk.length; if (size > 24 * 1024 * 1024) { const error = new Error('Request is too large.'); error.statusCode = 413; throw error } raw += chunk }
  try { return raw ? JSON.parse(raw) : {} }
  catch { const error = new Error('Request body is not valid JSON.'); error.statusCode = 400; throw error }
}

function text(value, maxLength) {
  return typeof value === 'string' ? value.trim().slice(0, maxLength) : ''
}

function normalisePhone(value) {
  const digits = String(value || '').replace(/\D/g, '')
  return digits.replace(/^0(?=[17]\d{8}$)/, '254')
}

function smsSegments(message) {
  const basic = "@£$¥èéùìòÇ\nØø\rÅåΔ_ΦΓΛΩΠΨΣΘΞÆæßÉ !\"#¤%&'()*+,-./0123456789:;<=>?¡ABCDEFGHIJKLMNOPQRSTUVWXYZÄÖÑÜ§¿abcdefghijklmnopqrstuvwxyzäöñüà"
  const extended = '^{}\\[~]|€'
  const isGsm = [...message].every(character => basic.includes(character) || extended.includes(character))
  const length = isGsm ? [...message].reduce((total, character) => total + (extended.includes(character) ? 2 : 1), 0) : Buffer.from(message, 'utf16le').length / 2
  const single = isGsm ? 160 : 70
  const multipart = isGsm ? 153 : 67
  return Math.max(1, Math.ceil(length / (length <= single ? single : multipart)))
}

function smsRecipients(data, input) {
  const audience = text(input.audience, 12)
  const grade = text(input.grade, 24)
  const admissionId = text(input.admissionId, 100)
  const learners = audience === 'learner' ? data.admissions.filter(item => item.id === admissionId) : audience === 'grade' ? data.admissions.filter(item => item.grade === grade) : data.admissions
  const contacts = new Map()
  for (const learner of learners) {
    const guardians = Array.isArray(learner.guardians) && learner.guardians.length ? learner.guardians : [{ name: learner.guardianName, phone: learner.guardianPhone, relationship: learner.relationship }]
    for (const guardian of guardians) {
      const phone = normalisePhone(guardian.phone)
      if (/^2547\d{8}$/.test(phone)) contacts.set(phone, { phone, learnerName: learner.learnerName, guardianName: text(guardian.name, 100), relationship: text(guardian.relationship, 40) })
    }
  }
  const audienceLabel = audience === 'learner' ? `${learners[0]?.learnerName || 'Selected learner'} — both recorded parents/guardians` : audience === 'grade' ? `${grade} parents and guardians` : 'All parents and guardians'
  return { audience: audienceLabel, recipients: [...contacts.values()] }
}

async function sendBlessedText(phone, message) {
  const apiKey = text(process.env.BLESSEDTEXTS_API_KEY, 300)
  const senderId = text(process.env.SMS_SENDER_ID, 11)
  if (!apiKey || !senderId) throw new Error('The BlessedTexts gateway or approved Sender ID is not configured.')
  const apiBase = text(process.env.BLESSEDTEXTS_API_BASE, 200) || 'https://api.blessedtexts.com/api'
  const providerResponse = await fetch(`${apiBase.replace(/\/$/, '')}/sms/v1/sendsms`, { method: 'POST', headers: { 'Content-Type': 'application/json', Accept: 'application/json' }, body: JSON.stringify({ api_key: apiKey, sender_id: senderId, message, phone }) })
  const result = await providerResponse.json().catch(() => null)
  if (!providerResponse.ok || !Array.isArray(result) || !result[0] || result[0].status === false) throw new Error('BlessedTexts did not accept this message.')
  return result[0]
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
    id: account.id, identifier: account.identifier, role: account.role, roles: accountRoles(account), name: account.name, title: account.title || '', parentPhone: account.parentPhone || '', active: account.active !== false,
    createdAt: account.createdAt || '', activatedAt: account.activatedAt || '', passwordChangedAt: account.passwordChangedAt || '',
    activationPending: Boolean(account.activationExpiresAt && Date.parse(account.activationExpiresAt) > Date.now()), activationExpiresAt: account.activationExpiresAt || ''
  }
}

function accountIdentifier(input, role) {
  const raw = text(input, 120)
  return (role === 'parent' || (role === 'teacher' && !raw.includes('@'))) ? normalisePhone(raw) : raw.toLowerCase()
}

function activationMessage(account, code, expiresAt) {
  return `St. Dorothy School portal activation for ${account.name}: use code ${code}. It expires ${new Date(expiresAt).toLocaleString('en-KE', { timeZone: 'Africa/Nairobi' })}. Do not share this code.`
}

function createSeedAccounts() {
  const accounts = isProduction ? seedPortalAccounts.filter(account => account.role === 'system-admin') : seedPortalAccounts
  return accounts.map(account => ({ ...account, active: true, ...createPasswordHash(account.role === 'system-admin' ? systemAdminInitialPassword : initialPortalPassword), createdAt: new Date().toISOString(), mustChangePassword: true }))
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
  const numbers = admissions.map(application => String(application.admissionNumber || '').trim()).filter(value => /^\d+$/.test(value)).map(Number).filter(value => Number.isSafeInteger(value) && value > 0)
  if (!numbers.length) return '1'
  return String(Math.max(...numbers) + 1)
}

function repeatedAdmissionNumber(number, admissions) {
  const base = String(number || '').trim()
  let suffix = 1
  let candidate = `${base}R`
  while (admissions.some(application => application.admissionNumber === candidate)) {
    suffix += 1
    candidate = `${base}R${suffix}`
  }
  return candidate
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
const schoolGrades = new Set(['Pre-Primary', 'Baby Class', 'PP1', 'PP2', ...Array.from({ length: 9 }, (_, index) => `Grade ${index + 1}`)])
const cbcSubjectCatalogVersion = 'KICD-2026-1'
const cbcSubjectCatalog = [
  ...['PP1', 'PP2'].flatMap(grade => ['Language Activities', 'Mathematical Activities', 'Environmental Activities', 'Creative Arts', 'Christian Religious Education Activities'].map(name => ({ grade, name }))),
  ...['Grade 1', 'Grade 2', 'Grade 3'].flatMap(grade => ['English', 'Kiswahili', 'Mathematical Activities', 'Environmental Activities', 'Creative Activities', 'Christian Religious Education', 'Indigenous Language'].map(name => ({ grade, name, optional: name === 'Indigenous Language' }))),
  ...['Grade 4', 'Grade 5', 'Grade 6'].flatMap(grade => ['English', 'Kiswahili', 'Mathematics', 'Science and Technology', 'Agriculture and Nutrition', 'Social Studies', 'Creative Activities', 'Christian Religious Education', 'Indigenous Language', 'Foreign Language'].map(name => ({ grade, name, optional: ['Indigenous Language', 'Foreign Language'].includes(name) }))),
  ...['Grade 7', 'Grade 8', 'Grade 9'].flatMap(grade => ['English', 'Kiswahili', 'Mathematics', 'Integrated Science', 'Social Studies', 'Agriculture', 'Pre-Technical Studies', 'Creative Arts and Sports', 'Christian Religious Education', 'Indigenous Language', 'Foreign Language'].map(name => ({ grade, name, optional: ['Indigenous Language', 'Foreign Language'].includes(name) })))
]
const learnerResidencies = new Set(['Day scholar', 'Boarder'])
const classStatuses = new Set(['Active', 'Inactive'])
const placementDecisions = new Set(['Pending headteacher approval', 'Confirmed', 'Waitlisted'])
const attendanceStatuses = new Set(['Present', 'Absent', 'Late', 'Excused'])
const staffRoles = new Set(['Headteacher', 'Deputy Headteacher', 'Teacher', 'Class Teacher', 'Secretary', 'Bursar', 'Support Staff'])
const standardFeePlans = [
  ...['Baby Class','PP1','PP2'].flatMap(grade => ['Term 1','Term 2','Term 3'].map(term => ({grade,residency:'Day scholar',term,amount:8150}))),
  ...['Grade 1','Grade 2','Grade 3'].flatMap(grade => [['Term 1',10150,16000],['Term 2',10150,15500],['Term 3',10150,15000]].flatMap(([term,day,boarder]) => [{grade,residency:'Day scholar',term,amount:day},{grade,residency:'Boarder',term,amount:boarder}])) ,
  ...['Grade 4','Grade 5','Grade 6','Grade 7','Grade 8','Grade 9'].flatMap(grade => ['Term 1','Term 2','Term 3'].flatMap(term => [{grade,residency:'Day scholar',term,amount:13500},{grade,residency:'Boarder',term,amount:20800}]))
]

function placementSuggestions(data, academicYear) {
  const activeClasses = data.schoolClasses.filter(item => item.status === 'Active' && item.academicYear === academicYear)
  const occupancy = new Map(activeClasses.map(item => [item.id, 0]))
  for (const placement of data.classPlacements) {
    if (placement.academicYear !== academicYear || placement.decision === 'Waitlisted') continue
    const schoolClass = activeClasses.find(item => item.className === placement.className)
    if (schoolClass) occupancy.set(schoolClass.id, (occupancy.get(schoolClass.id) || 0) + 1)
  }
  const byPhone = new Map()
  for (const learner of data.admissions) {
    const phone = normalisePhone(learner.guardianPhone)
    if (phone) byPhone.set(phone, [...(byPhone.get(phone) || []), learner])
  }
  const unplaced = data.admissions
    .filter(learner => !learner.placement && learner.completionStatus !== 'Completed / Alumni')
    .sort((a, b) => `${a.grade}|${a.learnerName}`.localeCompare(`${b.grade}|${b.learnerName}`))
  return unplaced.map(learner => {
    const candidates = activeClasses.filter(item => item.grade === learner.grade && (occupancy.get(item.id) || 0) < item.capacity)
    const siblings = byPhone.get(normalisePhone(learner.guardianPhone)) || []
    const siblingStreams = siblings.filter(item => item.id !== learner.id && item.placement?.academicYear === academicYear).map(item => activeClasses.find(schoolClass => schoolClass.className === item.placement.className)?.stream).filter(Boolean)
    const siblingCandidate = candidates.find(item => siblingStreams.includes(item.stream))
    const suggested = siblingCandidate || candidates.slice().sort((a, b) => (occupancy.get(a.id) || 0) - (occupancy.get(b.id) || 0) || a.className.localeCompare(b.className))[0]
    if (suggested) occupancy.set(suggested.id, (occupancy.get(suggested.id) || 0) + 1)
    return { admissionId: learner.id, learnerName: learner.learnerName, admissionNumber: learner.admissionNumber, grade: learner.grade, guardianPhone: learner.guardianPhone || '', classId: suggested?.id || '', className: suggested?.className || '', stream: suggested?.stream || '', occupancy: suggested ? occupancy.get(suggested.id) : 0, capacity: suggested?.capacity || 0, reason: suggested ? (siblingCandidate ? `Sibling stream match: ${suggested.stream}` : 'Lowest current class size') : 'No active class with available space' }
  })
}

function currentFeeSummary(data, application) {
  const balanceImport = (data.feeBalanceImports || []).filter(item => item.admissionId === application.id && Number.isFinite(Number(item.balance))).sort((a, b) => String(b.importedAt).localeCompare(String(a.importedAt)))[0]
  if (balanceImport) {
    const importedAt = String(balanceImport.importedAt || '')
    const totalPaid = data.receipts.filter(receipt => receipt.admissionId === application.id && (!importedAt || String(receipt.issuedAt || '') >= importedAt)).reduce((total, receipt) => total + Number(receipt.amount || 0), 0)
    const feeDue = Math.max(0, Number(balanceImport.balance || 0))
    return { feeDue, grossFeeDue: feeDue, bursaryAmount: 0, totalPaid, balance: Math.max(0, feeDue - totalPaid), term: 'Imported school balance', importedBalance: true, importedAt }
  }
  const placement = application.placement
  if (!placement?.academicYear || !placement?.term) return { feeDue: null, totalPaid: 0, balance: null, term: null }
  const residency = application.residency || 'Day scholar'
  const structure = data.feeStructures.find(item => item.grade === application.grade && (item.residency || 'Day scholar') === residency && item.academicYear === placement.academicYear && item.term === placement.term)
  const totalPaid = data.receipts.filter(receipt => receipt.admissionId === application.id).reduce((total, receipt) => total + Number(receipt.amount || 0), 0)
  const grossFeeDue = structure ? Number(structure.amount) : null
  const bursaryAmount = grossFeeDue === null ? 0 : (data.feeAdjustments || []).filter(item => item.admissionId === application.id && item.academicYear === placement.academicYear && item.term === placement.term).reduce((total, item) => total + Number(item.amount || 0), 0)
  const feeDue = grossFeeDue === null ? null : Math.max(0, grossFeeDue - bursaryAmount)
  return { feeDue, grossFeeDue, bursaryAmount, totalPaid, balance: feeDue === null ? null : Math.max(0, feeDue - totalPaid), term: structure ? `${structure.term} ${structure.academicYear}` : null }
}

function addFinanceAudit(data, request, action, recordType, recordId, detail = '') { data.financeAudit ??= []; data.financeAudit.unshift({ id:`finance-audit-${Date.now()}-${Math.random().toString(36).slice(2,7)}`, action, recordType, recordId, detail, by:request.account.name, role:request.account.role, createdAt:new Date().toISOString() }) }
function addAccountAudit(data, action, account, detail = '') { data.accountAudit ??= []; data.accountAudit.unshift({ id:`audit-${Date.now()}-${Math.random().toString(36).slice(2,7)}`, action, accountId:account?.id || '', accountName:account?.name || '', by:account?.name || 'System', detail, createdAt:new Date().toISOString() }) }

const money = value => `KES ${Number(value || 0).toLocaleString('en-KE')}`
const receiptVoteHeadList = ['Tuition / School fees','Boarding fees','Meals / Lunch programme','Transport','Examinations / Assessments','Activity, Sports & Trips','Development / PTA projects','Books, Stationery & Learning materials','Medical / Insurance','Uniform','Damages / Lost items']
const receiptVoteHeads = new Set(receiptVoteHeadList)
const safePdfName = value => String(value || 'learner').replace(/[^a-z0-9_-]+/gi, '_').replace(/^_+|_+$/g, '')
async function createClassListPdf(data, schoolClass) {
  const learners = data.admissions
    .filter(item => item.placement?.className === schoolClass.className && item.placement?.academicYear === schoolClass.academicYear)
    .slice().sort((a, b) => a.learnerName.localeCompare(b.learnerName))
  const classTeachers = [...new Set(data.teacherAssignments.filter(item => item.classId === schoolClass.id).map(item => item.staffName).filter(Boolean))]
  const teacherName = classTeachers.length ? classTeachers.join(', ') : 'Not assigned'
  let crest = null
  try { crest = await readFile(crestPath) } catch { /* The register remains usable if the crest asset is unavailable. */ }
  const chunks = []; const document = new PDFDocument({ size: 'A4', margin: 42, bufferPages: true, autoFirstPage: false })
  document.on('data', chunk => chunks.push(chunk))
  const finished = new Promise((resolve, reject) => { document.on('end', () => resolve(Buffer.concat(chunks))); document.on('error', reject) })
  const generated = new Intl.DateTimeFormat('en-KE', { dateStyle: 'medium', timeStyle: 'short', timeZone: 'Africa/Nairobi' }).format(new Date())
  let y = 0
  const startPage = () => {
    document.addPage()
    if (crest) document.image(crest, 281, 18, { fit: [33, 33] })
    document.fillColor('#173662').font('Helvetica-Bold').fontSize(14).text("ST. DOROTHY'S SCHOOL MASSIMO", 42, 56, { width: 511, align: 'center' })
    document.font('Helvetica').fontSize(7.5).fillColor('#536B84').text('P.O. Box 36-60607, Mikinduri, Meru, Kenya | +254 722 841 569 | stdorothysschoolmassimo@gmail.com', 42, 74, { width: 511, align: 'center' }).text('We lead by example', 42, 86, { width: 511, align: 'center' })
    document.moveTo(42, 102).lineTo(553, 102).stroke('#D8A20C')
    document.fillColor('#1D5EAA').font('Helvetica-Bold').fontSize(12).text(`CLASS LIST - ${schoolClass.className.toUpperCase()}`, 42, 113, { width: 511, align: 'center' })
    document.font('Helvetica').fontSize(7.5).fillColor('#536B84').text(`Academic year: ${schoolClass.academicYear}   |   Class teacher: ${teacherName}   |   Generated: ${generated}`, 42, 129, { width: 511, align: 'center' })
    const headerY = 145
    document.rect(42, headerY, 511, 18).fill('#1D5EAA')
    document.fillColor('#FFFFFF').font('Helvetica-Bold').fontSize(7).text('NO.', 48, headerY + 5).text('ADMISSION NO.', 80, headerY + 5).text('LEARNER NAME', 180, headerY + 5).text('CATEGORY', 442, headerY + 5)
    y = headerY + 18
  }
  startPage()
  if (!learners.length) {
    document.fillColor('#536B84').font('Helvetica').fontSize(10).text('No learners are currently placed in this class stream.', 42, y + 24, { width: 511, align: 'center' })
  }
  learners.forEach((learner, index) => {
    const rowHeight = 15
    if (y + rowHeight > 779) startPage()
    if (index % 2 === 0) document.rect(42, y, 511, rowHeight).fill('#EDF5FF')
    document.fillColor('#173662').font('Helvetica').fontSize(7.4).text(String(index + 1), 48, y + 4, { width: 24, align: 'center' })
    document.text(learner.admissionNumber || 'Pending', 78, y + 4, { width: 92 })
    document.font('Helvetica-Bold').text(learner.learnerName, 180, y + 4, { width: 252, height: 10, ellipsis: true })
    document.font('Helvetica').text(learner.residency || 'Day scholar', 442, y + 4, { width: 103 })
    document.strokeColor('#DCE5EF').lineWidth(0.5).moveTo(42, y + rowHeight).lineTo(553, y + rowHeight).stroke()
    y += rowHeight
  })
  const range = document.bufferedPageRange()
  for (let index = 0; index < range.count; index += 1) {
    document.switchToPage(range.start + index)
    document.fillColor('#536B84').font('Helvetica').fontSize(7.5).text(`Official class register | ${schoolClass.className} | Page ${index + 1} of ${range.count}`, 42, 787, { width: 511, align: 'center' })
  }
  document.end()
  return finished
}
async function createBoarderListPdf(data, grade = '') {
  const learners = data.admissions.filter(item => (item.residency || 'Day scholar') === 'Boarder' && (!grade || item.grade === grade)).slice().sort((a, b) => `${a.placement?.className || a.grade}|${a.learnerName}`.localeCompare(`${b.placement?.className || b.grade}|${b.learnerName}`))
  let crest = null
  try { crest = await readFile(crestPath) } catch { /* The report remains usable if the crest asset is unavailable. */ }
  const chunks = []; const document = new PDFDocument({ size: 'A4', margin: 42, bufferPages: true, autoFirstPage: false })
  document.on('data', chunk => chunks.push(chunk))
  const finished = new Promise((resolve, reject) => { document.on('end', () => resolve(Buffer.concat(chunks))); document.on('error', reject) })
  const generated = new Intl.DateTimeFormat('en-KE', { dateStyle: 'medium', timeStyle: 'short', timeZone: 'Africa/Nairobi' }).format(new Date())
  const reportGrade = grade || 'ALL GRADES'
  let y = 0
  const startPage = () => {
    document.addPage()
    if (crest) document.image(crest, 281, 18, { fit: [33, 33] })
    document.fillColor('#173662').font('Helvetica-Bold').fontSize(14).text("ST. DOROTHY'S SCHOOL MASSIMO", 42, 56, { width: 511, align: 'center' })
    document.font('Helvetica').fontSize(7.5).fillColor('#536B84').text('P.O. Box 36-60607, Mikinduri, Meru, Kenya | +254 722 841 569 | stdorothysschoolmassimo@gmail.com', 42, 74, { width: 511, align: 'center' }).text('We lead by example', 42, 86, { width: 511, align: 'center' })
    document.moveTo(42, 102).lineTo(553, 102).stroke('#D8A20C')
    document.fillColor('#1D5EAA').font('Helvetica-Bold').fontSize(12).text(`BOARDING REGISTER - ${reportGrade.toUpperCase()}`, 42, 113, { width: 511, align: 'center' })
    document.font('Helvetica').fontSize(7.5).fillColor('#536B84').text(`Confirmed boarders only | Learners: ${learners.length} | Generated: ${generated}`, 42, 129, { width: 511, align: 'center' })
    const headerY = 145
    document.rect(42, headerY, 511, 18).fill('#1D5EAA')
    document.fillColor('#FFFFFF').font('Helvetica-Bold').fontSize(7).text('NO.', 48, headerY + 5).text('ADMISSION NO.', 80, headerY + 5).text('LEARNER NAME', 180, headerY + 5).text('CLASS', 390, headerY + 5).text('CATEGORY', 472, headerY + 5)
    y = headerY + 18
  }
  startPage()
  if (!learners.length) document.fillColor('#536B84').font('Helvetica').fontSize(10).text('No confirmed boarders were found for this grade.', 42, y + 24, { width: 511, align: 'center' })
  learners.forEach((learner, index) => {
    const rowHeight = 15
    if (y + rowHeight > 779) startPage()
    if (index % 2 === 0) document.rect(42, y, 511, rowHeight).fill('#EDF5FF')
    document.fillColor('#173662').font('Helvetica').fontSize(7.2).text(String(index + 1), 48, y + 4, { width: 24, align: 'center' })
    document.text(learner.admissionNumber || 'Pending', 78, y + 4, { width: 92 })
    document.font('Helvetica-Bold').text(learner.learnerName, 180, y + 4, { width: 202, height: 10, ellipsis: true })
    document.font('Helvetica').text(learner.placement?.className || learner.grade || '—', 390, y + 4, { width: 76, height: 10, ellipsis: true })
    document.text('Boarder', 472, y + 4, { width: 70 })
    document.strokeColor('#DCE5EF').lineWidth(0.5).moveTo(42, y + rowHeight).lineTo(553, y + rowHeight).stroke()
    y += rowHeight
  })
  const range = document.bufferedPageRange()
  for (let index = 0; index < range.count; index += 1) { document.switchToPage(range.start + index); document.fillColor('#536B84').font('Helvetica').fontSize(7.5).text(`Official boarding register | ${reportGrade} | Page ${index + 1} of ${range.count}`, 42, 787, { width: 511, align: 'center' }) }
  document.end()
  return finished
}
async function createDefaulterListPdf(data, className = '') {
  const rows = data.admissions
    .map(learner => ({ learner, finance: currentFeeSummary(data, learner) }))
    .filter(({ learner, finance }) => (!className || (learner.placement?.className || learner.grade) === className) && Number(finance.balance || 0) > 0)
    .sort((a, b) => a.learner.learnerName.localeCompare(b.learner.learnerName))
  let crest = null
  try { crest = await readFile(crestPath) } catch { /* The report remains usable if the crest asset is unavailable. */ }
  const chunks = []; const document = new PDFDocument({ size: 'A4', margin: 42, bufferPages: true, autoFirstPage: false })
  document.on('data', chunk => chunks.push(chunk))
  const finished = new Promise((resolve, reject) => { document.on('end', () => resolve(Buffer.concat(chunks))); document.on('error', reject) })
  const generated = new Intl.DateTimeFormat('en-KE', { dateStyle: 'medium', timeStyle: 'short', timeZone: 'Africa/Nairobi' }).format(new Date())
  const total = rows.reduce((sum, item) => sum + Number(item.finance.balance || 0), 0)
  const reportClass = className || 'ALL CLASSES'
  let y = 0
  const startPage = () => {
    document.addPage()
    if (crest) document.image(crest, 281, 18, { fit: [33, 33] })
    document.fillColor('#173662').font('Helvetica-Bold').fontSize(14).text("ST. DOROTHY'S SCHOOL MASSIMO", 42, 56, { width: 511, align: 'center' })
    document.font('Helvetica').fontSize(7.5).fillColor('#536B84').text('P.O. Box 36-60607, Mikinduri, Meru, Kenya | +254 722 841 569 | stdorothysschoolmassimo@gmail.com', 42, 74, { width: 511, align: 'center' }).text('We lead by example', 42, 86, { width: 511, align: 'center' })
    document.moveTo(42, 102).lineTo(553, 102).stroke('#D8A20C')
    document.fillColor('#1D5EAA').font('Helvetica-Bold').fontSize(12).text(`OUTSTANDING FEE BALANCES - ${reportClass.toUpperCase()}`, 42, 113, { width: 511, align: 'center' })
    document.font('Helvetica').fontSize(7.5).fillColor('#536B84').text(`Generated: ${generated}   |   Learners: ${rows.length}   |   Total outstanding: ${money(total)}`, 42, 129, { width: 511, align: 'center' })
    const headerY = 145
    document.rect(42, headerY, 511, 18).fill('#1D5EAA')
    document.fillColor('#FFFFFF').font('Helvetica-Bold').fontSize(7).text('NO.', 48, headerY + 5).text('ADMISSION NO.', 80, headerY + 5).text('LEARNER NAME', 180, headerY + 5).text('CLASS', 340, headerY + 5).text('FEE DUE', 405, headerY + 5, { width: 55, align: 'right' }).text('PAID', 465, headerY + 5, { width: 40, align: 'right' }).text('BALANCE', 510, headerY + 5, { width: 37, align: 'right' })
    y = headerY + 18
  }
  startPage()
  if (!rows.length) document.fillColor('#536B84').font('Helvetica').fontSize(10).text('No learners with outstanding balances were found.', 42, y + 24, { width: 511, align: 'center' })
  rows.forEach(({ learner, finance }, index) => {
    const rowHeight = 15
    if (y + rowHeight > 779) startPage()
    if (index % 2 === 0) document.rect(42, y, 511, rowHeight).fill('#EDF5FF')
    document.fillColor('#173662').font('Helvetica').fontSize(7.2).text(String(index + 1), 48, y + 4, { width: 24, align: 'center' })
    document.text(learner.admissionNumber || 'Pending', 78, y + 4, { width: 92 })
    document.font('Helvetica-Bold').text(learner.learnerName, 180, y + 4, { width: 150, height: 10, ellipsis: true })
    document.font('Helvetica').text(learner.placement?.className || learner.grade || '—', 340, y + 4, { width: 60, height: 10, ellipsis: true })
    document.text(Number(finance.feeDue || 0).toLocaleString('en-KE'), 405, y + 4, { width: 55, align: 'right' })
    document.text(Number(finance.totalPaid || 0).toLocaleString('en-KE'), 465, y + 4, { width: 40, align: 'right' })
    document.font('Helvetica-Bold').fillColor('#933B3B').text(Number(finance.balance || 0).toLocaleString('en-KE'), 510, y + 4, { width: 37, align: 'right' })
    document.strokeColor('#DCE5EF').lineWidth(0.5).moveTo(42, y + rowHeight).lineTo(553, y + rowHeight).stroke()
    y += rowHeight
  })
  const range = document.bufferedPageRange()
  for (let index = 0; index < range.count; index += 1) {
    document.switchToPage(range.start + index)
    document.fillColor('#536B84').font('Helvetica').fontSize(7.5).text(`Official defaulter register | ${reportClass} | Page ${index + 1} of ${range.count}`, 42, 787, { width: 511, align: 'center' })
  }
  document.end()
  return finished
}
async function createExpensesPdf(entries) {
  let crest = null
  try { crest = await readFile(crestPath) } catch { /* The report remains usable if the crest asset is unavailable. */ }
  const chunks = []; const document = new PDFDocument({ size: 'A4', margin: 42, bufferPages: true, autoFirstPage: false })
  document.on('data', chunk => chunks.push(chunk))
  const finished = new Promise((resolve, reject) => { document.on('end', () => resolve(Buffer.concat(chunks))); document.on('error', reject) })
  const generated = new Intl.DateTimeFormat('en-KE', { dateStyle: 'medium', timeStyle: 'short', timeZone: 'Africa/Nairobi' }).format(new Date())
  const total = entries.reduce((sum, entry) => sum + Number(entry.amount || 0), 0)
  let y = 0
  const startPage = () => {
    document.addPage()
    if (crest) document.image(crest, 281, 18, { fit: [33, 33] })
    document.fillColor('#173662').font('Helvetica-Bold').fontSize(14).text("ST. DOROTHY'S SCHOOL MASSIMO", 42, 56, { width: 511, align: 'center' })
    document.font('Helvetica').fontSize(7.5).fillColor('#536B84').text('P.O. Box 36-60607, Mikinduri, Meru, Kenya | +254 722 841 569 | stdorothysschoolmassimo@gmail.com', 42, 74, { width: 511, align: 'center' }).text('We lead by example', 42, 86, { width: 511, align: 'center' })
    document.moveTo(42, 102).lineTo(553, 102).stroke('#D8A20C')
    document.fillColor('#1D5EAA').font('Helvetica-Bold').fontSize(12).text('EXPENDITURE REGISTER', 42, 113, { width: 511, align: 'center' })
    document.font('Helvetica').fontSize(7.5).fillColor('#536B84').text(`Generated: ${generated}   |   Entries: ${entries.length}   |   Total expenditure: ${money(total)}`, 42, 129, { width: 511, align: 'center' })
    const headerY = 145
    document.rect(42, headerY, 511, 18).fill('#1D5EAA')
    document.fillColor('#FFFFFF').font('Helvetica-Bold').fontSize(6.5).text('DATE', 47, headerY + 5).text('PAYEE / DESCRIPTION', 91, headerY + 5).text('CATEGORY', 238, headerY + 5).text('METHOD', 316, headerY + 5).text('REFERENCE / BANK', 372, headerY + 5).text('AMOUNT', 492, headerY + 5, { width: 55, align: 'right' })
    y = headerY + 18
  }
  startPage()
  if (!entries.length) document.fillColor('#536B84').font('Helvetica').fontSize(10).text('No expenses have been recorded yet.', 42, y + 24, { width: 511, align: 'center' })
  entries.forEach((entry, index) => {
    const rowHeight = 20
    if (y + rowHeight > 779) startPage()
    if (index % 2 === 0) document.rect(42, y, 511, rowHeight).fill('#EDF5FF')
    const method = entry.paymentMethod || (entry.type === 'Cash out' ? 'Cash' : entry.type || 'Cash')
    const reference = method === 'M-Pesa' ? entry.mpesaReference || entry.reference || '—' : method === 'Cheque' ? `${entry.bank || 'Bank not recorded'} · ${entry.chequeNumber || entry.reference || '—'}` : entry.reference || '—'
    document.fillColor('#173662').font('Helvetica').fontSize(6.8).text(entry.date || '—', 47, y + 6, { width: 40 })
    document.font('Helvetica-Bold').text(entry.payee || entry.description || '—', 91, y + 4, { width: 138, height: 8, ellipsis: true }).font('Helvetica').fillColor('#536B84').text(entry.description || '', 91, y + 12, { width: 138, height: 7, ellipsis: true })
    document.fillColor('#173662').text(entry.category || 'Other', 238, y + 6, { width: 70, height: 8, ellipsis: true })
    document.text(method, 316, y + 6, { width: 51, height: 8, ellipsis: true })
    document.text(reference, 372, y + 6, { width: 112, height: 8, ellipsis: true })
    document.font('Helvetica-Bold').text(Number(entry.amount || 0).toLocaleString('en-KE'), 492, y + 6, { width: 55, align: 'right' })
    document.strokeColor('#DCE5EF').lineWidth(0.5).moveTo(42, y + rowHeight).lineTo(553, y + rowHeight).stroke()
    y += rowHeight
  })
  const range = document.bufferedPageRange()
  for (let index = 0; index < range.count; index += 1) { document.switchToPage(range.start + index); document.fillColor('#536B84').font('Helvetica').fontSize(7.5).text(`Official expenditure register | Page ${index + 1} of ${range.count}`, 42, 787, { width: 511, align: 'center' }) }
  document.end()
  return finished
}
async function createLearnerRecordPdf(data, learner) {
  const finance = currentFeeSummary(data, learner)
  let crest = null
  try { crest = await readFile(crestPath) } catch { /* The record remains usable if the crest asset is unavailable. */ }
  const chunks = []; const document = new PDFDocument({ size: 'A4', margin: 42 })
  document.on('data', chunk => chunks.push(chunk))
  const finished = new Promise((resolve, reject) => { document.on('end', () => resolve(Buffer.concat(chunks))); document.on('error', reject) })
  if (crest) document.image(crest, 274, 28, { fit: [47, 47] })
  document.fillColor('#173662').font('Helvetica-Bold').fontSize(16).text("ST. DOROTHY'S SCHOOL MASSIMO", 42, 82, { width: 511, align: 'center' })
  document.font('Helvetica').fontSize(8.5).fillColor('#536B84').text('P.O. Box 36-60607, Mikinduri, Meru, Kenya | +254 722 841 569', 42, 104, { width: 511, align: 'center' }).text('stdorothysschoolmassimo@gmail.com | We lead by example', 42, 117, { width: 511, align: 'center' })
  document.moveTo(42, 137).lineTo(553, 137).stroke('#D8A20C')
  document.fillColor('#1D5EAA').font('Helvetica-Bold').fontSize(14).text('LEARNER RECORD', 42, 151, { width: 511, align: 'center' })
  document.font('Helvetica').fontSize(8.5).fillColor('#536B84').text(`Printed ${new Intl.DateTimeFormat('en-KE', { dateStyle: 'medium', timeStyle: 'short', timeZone: 'Africa/Nairobi' }).format(new Date())}`, 42, 172, { width: 511, align: 'center' })
  const box = (title, rows, top) => {
    const height = 30 + Math.ceil(rows.length / 2) * 39
    document.roundedRect(42, top, 511, height, 7).fillAndStroke('#F7FAFD', '#DCE5EF')
    document.fillColor('#173662').font('Helvetica-Bold').fontSize(10).text(title, 56, top + 13)
    rows.forEach(([label, value], index) => {
      const column = index % 2; const row = Math.floor(index / 2); const x = column ? 306 : 56; const y = top + 34 + row * 39
      document.fillColor('#6B8099').font('Helvetica-Bold').fontSize(7.5).text(label.toUpperCase(), x, y)
      document.fillColor('#173662').font('Helvetica').fontSize(9).text(value || 'Not recorded', x, y + 12, { width: 220 })
    })
    return top + height + 14
  }
  const father = learner.fatherName || learner.guardians?.find(item => /father/i.test(item.relationship || ''))?.name || ''
  const fatherPhone = learner.fatherPhone || learner.guardians?.find(item => /father/i.test(item.relationship || ''))?.phone || ''
  const mother = learner.motherName || learner.guardians?.find(item => /mother|guardian/i.test(item.relationship || ''))?.name || learner.guardianName || ''
  const motherPhone = learner.motherPhone || learner.guardians?.find(item => /mother|guardian/i.test(item.relationship || ''))?.phone || learner.guardianPhone || ''
  let y = box('Learner details', [['Learner name', learner.learnerName], ['Admission number', learner.admissionNumber], ['Grade / class', learner.placement?.className || learner.grade], ['Category', learner.residency || 'Day scholar'], ['Status', learner.status || 'Active'], ['Date of birth', learner.dateOfBirth || 'Not recorded']], 195)
  y = box('Parent and guardian contacts', [["Father's name", father], ["Father's phone", fatherPhone], ['Mother / guardian', mother], ['Mother / guardian phone', motherPhone]], y)
  y = box('Fee summary', [['Fee plan', finance.feeDue === null ? 'Not yet set' : money(finance.feeDue)], ['Payments recorded', money(finance.totalPaid)], ['Outstanding balance', finance.balance === null ? 'Not yet set' : money(finance.balance)], ['Fee period', finance.term || 'Current placement period']], y)
  document.fillColor('#536B84').font('Helvetica-Oblique').fontSize(7.5).text('This computer-generated learner record is confidential and for authorised school use only.', 42, 790, { width: 511, align: 'center' })
  document.end()
  return finished
}
async function createFeeStatementPdf(data, learner, servedBy = 'School office') {
  const finance = currentFeeSummary(data, learner); const receipts = data.receipts.filter(item => item.admissionId === learner.id).slice().sort((a, b) => String(a.issuedAt).localeCompare(String(b.issuedAt)))
  const statementNumber = `SD/STM/${new Date().getFullYear()}/${learner.admissionNumber || learner.id}`; const generatedAt = new Date(); const payBy = new Date(generatedAt); payBy.setDate(payBy.getDate() + 14)
  const chunks = []; const document = new PDFDocument({ size: 'A4', margin: 42 }); document.on('data', chunk => chunks.push(chunk))
  const finished = new Promise((resolve, reject) => { document.on('end', () => resolve(Buffer.concat(chunks))); document.on('error', reject) })
  try { document.image(await readFile(crestPath), 42, 36, { fit: [54, 54] }) } catch { /* Statement remains valid without its crest. */ }
  document.fillColor('#173662').fontSize(17).font('Helvetica-Bold').text("ST. DOROTHY'S SCHOOL MASSIMO", 108, 42); document.fontSize(8.5).font('Helvetica').text('P.O. Box 36-60607, Mikinduri, Meru, Kenya | +254 722 841 569 | stdorothysschoolmassimo@gmail.com', 108, 66); document.moveTo(42, 100).lineTo(553, 100).stroke('#1D5EAA')
  document.fillColor('#173662').fontSize(15).font('Helvetica-Bold').text('FEE STATEMENT', 42, 116); document.fontSize(8.5).font('Helvetica').fillColor('#536b84').text(`Statement No: ${statementNumber}`, 42, 138).text(`Generated: ${generatedAt.toLocaleString('en-KE')} | Statement period: ${finance.term || 'Current placement period'}`, 42, 151)
  document.fillColor('#173662').fontSize(10).font('Helvetica-Bold').text('Learner and guardian', 42, 180); document.font('Helvetica').fontSize(9).text(`Learner: ${learner.learnerName}`, 42, 198).text(`Admission number: ${learner.admissionNumber || 'Pending'}    Grade/Class: ${learner.placement?.className || learner.grade}`, 42, 212).text(`Guardian: ${learner.guardianName || 'Not recorded'}    Phone: ${learner.guardianPhone || 'Not recorded'}`, 42, 226)
  document.fillColor('#173662').font('Helvetica-Bold').fontSize(10).text('Account summary', 42, 258); document.font('Helvetica').fontSize(9).text(`Gross fee plan: ${finance.grossFeeDue === null ? 'Fee plan pending' : money(finance.grossFeeDue)}`, 42, 277).text(`Approved discounts / bursaries: ${money(finance.bursaryAmount)}${finance.bursaryAmount ? '' : ' (none recorded)'}`, 42, 291).text(`Net fee due: ${finance.feeDue === null ? 'Fee plan pending' : money(finance.feeDue)}`, 42, 305).text(`Recorded payments: ${money(finance.totalPaid)}`, 42, 319); document.fillColor('#933B3B').font('Helvetica-Bold').fontSize(13).text(`Outstanding balance: ${finance.balance === null ? 'Fee plan pending' : money(finance.balance)}`, 42, 342)
  document.fillColor('#173662').fontSize(10).text('Recorded payments and receipt references', 42, 360); document.font('Helvetica-Bold').fontSize(8).text('DATE', 42, 380).text('RECEIPT / REFERENCE', 125, 380).text('METHOD', 315, 380).text('AMOUNT', 440, 380); document.moveTo(42, 392).lineTo(553, 392).stroke('#B8C9DA'); document.font('Helvetica').fontSize(8)
  let y = 403; if (!receipts.length) document.text('No confirmed payments recorded for this learner.', 42, y)
  receipts.slice(0, 12).forEach(receipt => { document.text(String(receipt.issuedAt || '').slice(0, 10), 42, y).text(`${receipt.receiptNumber || ''}${receipt.paymentReference ? ` / ${receipt.paymentReference}` : ''}`, 125, y, { width: 180 }).text(receipt.paymentMethod || '', 315, y).text(money(receipt.amount), 440, y); y += 16 })
  const guidanceY = Math.max(y + 24, 455); document.fillColor('#173662').font('Helvetica-Bold').fontSize(10).text(`Payment guidance - Pay by ${payBy.toLocaleDateString('en-KE')}`, 42, guidanceY); document.font('Helvetica').fontSize(8.5).text('Use the learner name and grade as your account reference. Retain the payment confirmation for reconciliation.', 42, guidanceY + 16, { width: 500 }); document.font('Helvetica-Bold').fontSize(9).text('Official payment channels', 42, guidanceY + 42); document.font('Helvetica').fontSize(8.5).text('Equity Paybill 247247 - Account 803434#Learner Name/Grade', 42, guidanceY + 57).text('DTB Paybill 516600 - Account 808616#Learner Name/Grade', 42, guidanceY + 71).text('Trans Nation Paybill 507900 - Account 9000482#Learner Name/Grade', 42, guidanceY + 85).text('Bank: Equity Bank Meru A/C 0140296813439 | DTB Bank Meru A/C 0259364001', 42, guidanceY + 99, { width: 500 }).text('For help, contact the Bursar on +254 722 841 569.', 42, guidanceY + 113)
  document.fillColor('#536b84').fontSize(7.5).text(`Prepared for you by ${servedBy}. Thank you.`, 42, 746, { align: 'center', width: 510 }); document.text('This is a computer-generated private fee statement. It is intended only for the learner’s parent or guardian.', 42, 760, { align: 'center', width: 510 }); document.end(); return finished
}

async function createReceiptPdf(receipt) {
  const chunks = []; const document = new PDFDocument({ size: 'A5', margin: 36 }); document.on('data', chunk => chunks.push(chunk))
  const finished = new Promise((resolve, reject) => { document.on('end', () => resolve(Buffer.concat(chunks))); document.on('error', reject) })
  try { document.image(await readFile(crestPath), 36, 22, { fit: [40, 40] }) } catch { /* The receipt remains valid if the crest is unavailable. */ }
  document.fillColor('#173662').font('Helvetica-Bold').fontSize(12.5).text("ST. DOROTHY'S SCHOOL MASSIMO", 84, 27)
  document.font('Helvetica').fontSize(6.6).fillColor('#536B84').text('P.O. Box 36-60607, Mikinduri, Meru, Kenya | +254 722 841 569', 84, 44).text('stdorothysschoolmassimo@gmail.com | We lead by example', 84, 54)
  document.moveTo(36, 70).lineTo(384, 70).stroke('#1D5EAA')
  document.fillColor('#173662').font('Helvetica-Bold').fontSize(13).text('PAYMENT RECEIPT', 36, 82)
  document.font('Helvetica').fontSize(6.9).fillColor('#536B84').text(`Receipt No: ${receipt.receiptNumber || 'Pending'} | Generated: ${new Intl.DateTimeFormat('en-KE', { dateStyle: 'medium', timeStyle: 'short', timeZone: 'Africa/Nairobi' }).format(new Date(receipt.issuedAt || Date.now()))}`, 36, 99, { width: 348 })
  const heading = (label, y) => document.fillColor('#173662').font('Helvetica-Bold').fontSize(9).text(label, 36, y)
  const detail = (label, value, y, x = 36) => { document.fillColor('#536B84').font('Helvetica-Bold').fontSize(7.2).text(label, x, y); document.fillColor('#173662').font('Helvetica').fontSize(8.2).text(value || 'Not recorded', x, y + 10, { width: 154 }) }
  heading('Learner and payment', 122)
  detail('LEARNER', receipt.learnerName, 137)
  detail('ADMISSION NUMBER', receipt.admissionNumber, 137, 210)
  detail('GRADE / CLASS', receipt.className, 161)
  detail('PAYMENT DATE', receipt.paymentDate || String(receipt.issuedAt || '').slice(0, 10), 161, 210)
  heading('Account summary', 193)
  const balanceText = value => value === null || value === undefined ? 'Not available' : money(value)
  document.rect(36, 208, 348, 18).fill('#F3F6F9')
  document.fillColor('#173662').font('Helvetica-Bold').fontSize(7.8).text('Balance before payment', 44, 213).text(balanceText(receipt.balanceBeforePayment), 254, 213, { width: 122, align: 'right' })
  document.rect(36, 228, 348, 18).fill('#EEF6FF')
  document.fillColor('#173662').font('Helvetica-Bold').fontSize(7.8).text('Amount received', 44, 233).text(money(receipt.amount), 254, 233, { width: 122, align: 'right' })
  document.rect(36, 248, 348, 18).fill('#E6F4E9')
  document.fillColor('#1F7044').font('Helvetica-Bold').fontSize(7.8).text('Remaining balance after payment', 44, 253).text(balanceText(receipt.remainingBalance), 254, 253, { width: 122, align: 'right' })
  heading('Payment allocation', 278)
  document.fillColor('#536B84').font('Helvetica-Bold').fontSize(6.8).text('VOTE HEAD / PARTICULAR', 36, 293).text('AMOUNT', 286, 293, { width: 98, align: 'right' })
  document.moveTo(36, 303).lineTo(384, 303).stroke('#B8C9DA')
  let allocationY = 308
  const savedAllocations = Array.isArray(receipt.allocations) ? receipt.allocations : []
  const amountFor = voteHead => savedAllocations.filter(item => item.voteHead === voteHead || (voteHead === 'Tuition / School fees' && ['School fees','Other approved charge'].includes(item.voteHead)) || (voteHead === 'Books, Stationery & Learning materials' && item.voteHead === 'ICT / Computer')).reduce((total, item) => total + Number(item.amount || 0), 0)
  const allocationRows = ['Arrears', ...receiptVoteHeadList].map(voteHead => ({ voteHead, amount: amountFor(voteHead) }))
  allocationRows.forEach((item, index) => {
    if (index % 2 === 0) document.rect(36, allocationY - 2, 348, 10).fill('#F8FAFC')
    document.fillColor(item.amount > 0 ? '#173662' : '#536B84').font(item.amount > 0 ? 'Helvetica-Bold' : 'Helvetica').fontSize(6.7).text(item.voteHead, 40, allocationY, { width: 238 }).text(money(item.amount), 286, allocationY, { width: 94, align:'right' })
    allocationY += 10
  })
  document.moveTo(36, allocationY + 1).lineTo(384, allocationY + 1).stroke('#D8A20C')
  document.font('Helvetica-Bold').fontSize(8).text('TOTAL', 36, allocationY + 8).text(money(receipt.amount), 286, allocationY + 8, { width:98, align:'right' })
  const paymentY = allocationY + 34
  heading('Payment record', paymentY)
  document.fillColor('#536B84').font('Helvetica-Bold').fontSize(7).text('CHANNEL', 36, paymentY + 18).text('REFERENCE', 151, paymentY + 18).text('SERVED BY', 272, paymentY + 18)
  document.moveTo(36, paymentY + 29).lineTo(384, paymentY + 29).stroke('#B8C9DA')
  document.fillColor('#173662').font('Helvetica').fontSize(8).text(receipt.paymentMethod || 'Not recorded', 36, paymentY + 37, { width: 105 }).text(receipt.paymentReference || 'Not recorded', 151, paymentY + 37, { width: 111 }).text(receipt.servedBy || 'School office', 272, paymentY + 37, { width: 112 })
  document.moveTo(36, paymentY + 60).lineTo(384, paymentY + 60).stroke('#D8A20C')
  document.fillColor('#173662').font('Helvetica-Bold').fontSize(8.8).text(`Served by ${receipt.servedBy || 'School office'}. Thank you.`, 36, paymentY + 72, { width: 348, align: 'center' })
  document.fillColor('#536B84').font('Helvetica').fontSize(6.8).text('This is a computer-generated school-fee receipt. Keep it for account verification.', 36, paymentY + 88, { width: 348, align: 'center' }); document.end(); return finished
}

function receiptForPrint(data, receipt, servedBy) {
  const learner = data.admissions.find(item => item.id === receipt.admissionId)
  const hasSavedBalance = Number.isFinite(Number(receipt.remainingBalance))
  let remainingBalance = hasSavedBalance ? Number(receipt.remainingBalance) : null
  if (!hasSavedBalance && learner) {
    const paymentsAfterThisReceipt = data.receipts
      .filter(item => item.admissionId === learner.id && String(item.issuedAt || '') > String(receipt.issuedAt || ''))
      .reduce((total, item) => total + Number(item.amount || 0), 0)
    remainingBalance = currentFeeSummary(data, learner).balance + paymentsAfterThisReceipt
  }
  const balanceBeforePayment = remainingBalance === null ? null : remainingBalance + Number(receipt.amount || 0)
  let allocations = Array.isArray(receipt.allocations) && receipt.allocations.length ? receipt.allocations : []
  if (!allocations.length) {
    const historicalBalance = learner ? currentFeeSummary(data, learner).importedBalance : false
    const arrearsAmount = historicalBalance && balanceBeforePayment !== null ? Math.min(Number(receipt.amount || 0), Math.max(0, balanceBeforePayment)) : 0
    const remainder = Math.max(0, Number(receipt.amount || 0) - arrearsAmount)
    allocations = [...(arrearsAmount > 0 ? [{ voteHead:'Arrears', amount:arrearsAmount, automatic:true }] : []), ...(remainder > 0 ? [{ voteHead:'Tuition / School fees', amount:remainder }] : [])]
  }
  return { ...receipt, allocations, remainingBalance, balanceBeforePayment, servedBy: servedBy || receipt.servedBy || 'School office' }
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

function importPhone(value) { const phone=normalisePhone(value); return /^2547\d{8}$/.test(phone)?phone:'' }
function importDate(value) { if (!value) return ''; if (value instanceof Date && !Number.isNaN(value)) return value.toISOString().slice(0,10); const textValue=String(value).trim(); const match=textValue.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); if (match) { const date=new Date(Date.UTC(Number(match[3]),Number(match[2])-1,Number(match[1]))); return date.getUTCFullYear()===Number(match[3])&&date.getUTCMonth()===Number(match[2])-1&&date.getUTCDate()===Number(match[1])?date.toISOString().slice(0,10):'' } return /^\d{4}-\d{2}-\d{2}/.test(textValue)?textValue.slice(0,10):'' }
function importGrade(value) { const raw=String(value||'').trim().toUpperCase(); if (raw==='BABY CLASS') return 'Baby Class'; if (raw==='PP1'||raw==='PP2') return raw; const names={ONE:'Grade 1',TWO:'Grade 2',THREE:'Grade 3',FOUR:'Grade 4',FIVE:'Grade 5',SIX:'Grade 6',SEVEN:'Grade 7',EIGHT:'Grade 8',NINE:'Grade 9'}; return names[raw.replace(/^GRADE\s+/,'')]||'' }
async function parseLearnerWorkbook(encoded,data) { const bytes=Buffer.from(String(encoded||''),'base64'); if (!bytes.length||bytes.length>6*1024*1024) throw new Error('Choose an Excel workbook smaller than 6 MB.')
  const workbook=new ExcelJS.Workbook(); await workbook.xlsx.load(bytes); const sheet=workbook.worksheets[0]; if (!sheet) throw new Error('The workbook has no first worksheet.')
  let section=''; const rows=[]; sheet.eachRow((row,rowNumber)=>{ const cells=row.values.slice(1); const [admission,serial,name,dob,father,fatherPhone,mother,motherPhone]=cells; const nameText=String(name||'').trim(); const header=String(admission||'').trim().toUpperCase()==='ADM. NO'&&nameText.toUpperCase()==='NAME'; const sectionHeader=/^(PP1|PP2|GRADE\s+(ONE|TWO|THREE|FOUR|FIVE|SIX|SEVEN|EIGHT|NINE))$/i.test(nameText)&&(!admission&&!serial||String(dob||'').trim().toUpperCase()==='DOB'); if (sectionHeader) { section=importGrade(nameText); return } if (header||!nameText||!section) return; const guardians=[{name:String(father||'').trim(),phone:importPhone(fatherPhone),relationship:'Father'},{name:String(mother||'').trim(),phone:importPhone(motherPhone),relationship:'Mother / guardian'}].filter(item=>item.name||item.phone); const primary=guardians.find(item=>item.phone)||guardians[0]||{name:'Guardian contact pending',phone:'',relationship:'Guardian'}; rows.push({row:rowNumber,learnerName:nameText,grade:section,admissionNumber:String(admission||'').trim(),dateOfBirth:importDate(dob),guardians,guardianName:primary.name||'Guardian contact pending',guardianPhone:primary.phone,relationship:primary.relationship,issues:[]}) })
  const seen=new Map(); rows.forEach(item=>{if(item.admissionNumber) seen.set(item.admissionNumber,(seen.get(item.admissionNumber)||0)+1)}); const existing=new Set(data.admissions.map(item=>String(item.admissionNumber||''))); rows.forEach(item=>{if(!item.admissionNumber)item.issues.push('Admission number missing — system will issue one'); if(item.admissionNumber&&seen.get(item.admissionNumber)>1)item.issues.push('Duplicate admission number — repeated record will receive R'); if(item.admissionNumber&&existing.has(item.admissionNumber))item.issues.push('Admission number already exists — repeated record will receive R'); if(!item.dateOfBirth&&String(item.dateOfBirth||'')==='')item.issues.push('Date of birth missing or needs review'); if(!item.guardianPhone)item.issues.push('No valid guardian phone')}); return rows }

async function parseTeacherWorkbook(encoded, data) {
  const bytes = Buffer.from(String(encoded || ''), 'base64')
  if (!bytes.length || bytes.length > 4 * 1024 * 1024) throw new Error('Choose an Excel workbook smaller than 4 MB.')
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(bytes)
  const sheet = workbook.worksheets[0]
  if (!sheet) throw new Error('The workbook has no first worksheet.')
  const rows = []
  const seenPhones = new Set()
  const existingIdentifiers = new Set(data.accounts.map(item => String(item.identifier || '').toLowerCase()))
  const existingStaffPhones = new Set(data.staff.map(item => normalisePhone(item.phone)).filter(Boolean))
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return
    const cells = row.values.slice(1)
    const name = text(cells[1], 100)
    const phone = normalisePhone(cells[2])
    if (!name && !phone) return
    const issues = []
    const validPhone = /^254[17]\d{8}$/.test(phone)
    if (!name) issues.push('Teacher name is missing')
    if (!validPhone) issues.push('Phone number is not a valid Kenyan mobile number')
    if (validPhone && seenPhones.has(phone)) issues.push('Duplicate phone number in this workbook')
    if (validPhone && existingIdentifiers.has(phone.toLowerCase())) issues.push('Portal account already exists')
    if (validPhone && existingStaffPhones.has(phone)) issues.push('Staff profile already exists')
    if (validPhone) seenPhones.add(phone)
    rows.push({ row: rowNumber, name, phone, role: 'teacher', title: 'Teacher', issues, ready: Boolean(name && validPhone && !issues.some(issue => issue.includes('Duplicate') || issue.includes('already exists'))) })
  })
  return rows
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
    const cbcSubjectsAdded = data.subjectCatalogVersion !== cbcSubjectCatalogVersion
    if (cbcSubjectsAdded) {
      const existingSubjects = new Set(data.subjects.map(item => `${String(item.grade || '').toLowerCase()}|${String(item.name || '').toLowerCase()}`))
      const seededAt = new Date().toISOString()
      cbcSubjectCatalog.forEach((subject, index) => {
        const key = `${subject.grade.toLowerCase()}|${subject.name.toLowerCase()}`
        if (!existingSubjects.has(key)) data.subjects.push({ id: `cbc-subject-${String(index + 1).padStart(3, '0')}`, name: subject.name, grade: subject.grade, status: subject.optional ? 'Inactive' : 'Active', optional: Boolean(subject.optional), source: 'KICD rationalised curriculum', createdAt: seededAt, updatedAt: seededAt })
      })
      data.subjectCatalogVersion = cbcSubjectCatalogVersion
    }
    data.teacherAssignments ??= []
    data.timetableEntries ??= []
    data.lessonPlans ??= []
    data.attendanceRecords ??= []
    data.assessments ??= []
    data.assessmentMarks ??= []
    data.feeStructures ??= []
    const feePlansAdded = !data.feeStructures.some(item => item.academicYear === '2027')
    if (feePlansAdded) data.feeStructures = [...data.feeStructures, ...standardFeePlans.map((plan,index) => ({ id:`standard-fees-2027-${index + 1}`, ...plan, academicYear:'2027', updatedAt:new Date().toISOString(), setBy:'School fee schedule' }))]
    data.feeAdjustments ??= []
    data.leaveRequests ??= []
    data.staffAttendance ??= []
    data.promotions ??= []
    data.receipts ??= []
    data.smsMessages ??= []
    data.smsWallet ??= { balance: 0, pricePerCreditKes: 0.8, senderId: '' }
    data.smsWallet.pricePerCreditKes = 0.8
    data.smsTopups ??= []
    data.mpesaTransactions ??= []
    data.cashbookEntries ??= []
    data.bankDeposits ??= []
    data.documentRecords ??= []
    data.internalMessages ??= []
    data.parentMessages ??= []
    const accountsAdded = !Array.isArray(data.accounts) || !data.accounts.length
    data.accounts = accountsAdded ? createSeedAccounts() : data.accounts
    const systemAdminAdded = ensureSystemAdminAccount(data)
    data.accountAudit ??= []
    data.financeAudit ??= []
    data.financeApprovals ??= []
    data.websiteSettings ??= { socialLinks: { facebook: '', instagram: '', youtube: '', tiktok: '' }, gallery: [] }
    const noticesAdded = !Array.isArray(data.notices)
    data.notices = noticesAdded ? structuredClone(defaultNotices) : data.notices
    const noticeFieldsAdded = ensureNoticeFields(data.notices)
    const admissionNumbersAdded = ensureAdmissionNumbers(data.admissions)
    const schoolClassesAdded = ensureSchoolClasses(data)

    if (accountsAdded || systemAdminAdded || feePlansAdded || cbcSubjectsAdded) await saveData(data)
    if (!requireAccess(request, response, data, url)) return

    if (request.method === 'GET' && url.pathname === '/api/health') return send(response, 200, { status: 'ok' })
    if (request.method === 'POST' && url.pathname === '/api/auth/login') {
      if (signInRateLimited(request, 'login', 8, 15 * 60_000)) return send(response, 429, { error: 'Too many sign-in attempts. Please wait 15 minutes before trying again.' })
      const input = await body(request)
      const portal = text(input.portal, 12)
      const identifier = portal === 'parent' ? normalisePhone(input.identifier) : text(input.identifier, 120).toLowerCase()
      const password = typeof input.password === 'string' ? input.password : ''
      const account = data.accounts.find(item => item.identifier === identifier && item.active !== false)
      if (!account || !password || !verifyPassword(password, account) || (portal === 'parent' && !accountHasRole(account, 'parent')) || (portal === 'staff' && (accountHasRole(account, 'parent') || accountHasRole(account, 'system-admin'))) || (portal === 'system-admin' && !accountHasRole(account, 'system-admin'))) {
        recordFailedAttempt(request, 'login', 15 * 60_000)
        return send(response, 401, { error: 'The sign-in details are not recognised. Please try again or contact the school office.' })
      }
      clearAttempts(request, 'login')
      addAccountAudit(data, 'Successful sign-in', account, `${portal || 'portal'} portal`)
      await saveData(data)
      createSession(response, account)
      return send(response, 200, { id: account.id, role: account.role, roles: accountRoles(account), name: account.name, title: account.title, parentPhone: account.parentPhone || '', mustChangePassword: Boolean(account.mustChangePassword) })
    }
    if (request.method === 'GET' && url.pathname === '/api/auth/session') return send(response, 200, { id: request.account.id, role: requestedSessionRole(request) || request.account.role, roles: accountRoles(request.account), name: request.account.name, title: request.account.title || '', parentPhone: request.account.parentPhone || '', mustChangePassword: Boolean(request.account.mustChangePassword) })
    if (request.method === 'POST' && url.pathname === '/api/auth/activity') return send(response, 200, { status: 'active' })
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
      addAccountAudit(data, 'Account activated', updatedAccount, 'Activation code accepted and password set')
      await saveData(data)
      createSession(response, updatedAccount)
      return send(response, 200, { role: updatedAccount.role, roles: accountRoles(updatedAccount), name: updatedAccount.name, title: updatedAccount.title, parentPhone: updatedAccount.parentPhone || '', mustChangePassword: false })
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
      addAccountAudit(data, 'Password changed', updatedAccount, 'Portal password updated')
      await saveData(data)
      return send(response, 200, { status: 'password updated' })
    }
    if (request.method === 'GET' && url.pathname === '/api/system-admin/overview') return send(response, 200, { accounts: { total: data.accounts.length, active: data.accounts.filter(account => account.active !== false).length, pendingActivation: data.accounts.filter(account => account.activationExpiresAt && Date.parse(account.activationExpiresAt) > Date.now()).length }, integrations: [{ name: 'M-Pesa Daraja', configured: Boolean(process.env.MPESA_CONSUMER_KEY && process.env.MPESA_CONSUMER_SECRET && process.env.MPESA_PASSKEY) }, { name: 'BlessedTexts SMS', configured: Boolean(process.env.BLESSEDTEXTS_API_KEY && process.env.SMS_SENDER_ID) }, { name: 'School email', configured: Boolean(process.env.SMTP_HOST && process.env.SMTP_USER) }], security: { secureCookies: isProduction, allowedOrigins: allowedOrigins.size, twoFactorAuthentication: false } })
    if (request.method === 'GET' && url.pathname === '/api/system-admin/logs') { const accountLogs=(data.accountAudit||[]).map(item=>({ id:`account-${item.id}`, area:'System access', action:item.action||'Account activity', detail:item.detail||item.accountName||'', by:item.by||item.performedBy||'System', createdAt:item.createdAt })); const financeLogs=(data.financeAudit||[]).map(item=>({ id:`finance-${item.id}`, area:'Finance', action:`${item.action||'Updated'} ${item.recordType||'record'}`, detail:item.detail||'', by:item.by||'System', createdAt:item.createdAt })); return send(response,200,[...accountLogs,...financeLogs].sort((a,b)=>String(b.createdAt||'').localeCompare(String(a.createdAt||''))).slice(0,300)) }
    if (request.method === 'GET' && url.pathname === '/api/system-admin/backup/status') { const documents=await Promise.all((data.documentRecords||[]).map(async item=>{try{return await stat(join(documentUploadDirectory,item.storedFileName))}catch{return null}})); const dataBytes=Buffer.byteLength(JSON.stringify(data)); const documentBytes=documents.reduce((total,item)=>total+Number(item?.size||0),0); return send(response,200,{generatedAt:new Date().toISOString(),dataBytes,documentBytes,totalBytes:dataBytes+documentBytes,documentsIncluded:documents.filter(Boolean).length,documentsMissing:documents.filter(item=>!item).length}) }
    if (request.method === 'GET' && url.pathname === '/api/system-admin/backup/download') { const generatedAt=new Date().toISOString(); const documentRecords=data.documentRecords||[]; const documents=await Promise.all(documentRecords.map(async item=>{try{return {record:item,content:await readFile(join(documentUploadDirectory,item.storedFileName))}}catch{return null}})); const availableDocuments=documents.filter(Boolean); const manifest={format:'st-dorothy-logical-backup-v1',generatedAt,documentsIncluded:availableDocuments.length,documentsMissing:documentRecords.length-availableDocuments.length,restoreNote:'This archive contains the complete school application state and uploaded learner documents. Keep it encrypted. Restore only with the portal technical administrator.'}; const zipName=`st-dorothy-school-backup_${generatedAt.slice(0,10)}.zip`; response.writeHead(200,{'Content-Type':'application/zip','Content-Disposition':`attachment; filename="${zipName}"`,'Cache-Control':'no-store'}); const archive=new ZipArchive({zlib:{level:9}}); archive.on('error',error=>response.destroy(error)); archive.pipe(response); archive.append(Buffer.from(JSON.stringify(data,null,2)),{name:'school-data.json'}); archive.append(Buffer.from(JSON.stringify(manifest,null,2)),{name:'backup-manifest.json'}); archive.append(Buffer.from('St. Dorothy School Massimo backup\n\nThis archive contains school-data.json and uploaded learner documents. Save it only on an encrypted SSD or trusted encrypted storage. Contact the portal technical administrator before restoring.\n'),{name:'README.txt'}); for(const item of availableDocuments) archive.append(item.content,{name:`learner-documents/${safeFileName(item.record.storedFileName)}`}); await archive.finalize(); return }
    if (request.method === 'POST' && url.pathname === '/api/system-admin/backup/restore') { const input=await body(request); if (text(input.confirmation,40)!=='RESTORE SCHOOL DATA') return send(response,400,{error:'Type RESTORE SCHOOL DATA to confirm this destructive action.'}); const archive=Buffer.from(String(input.archive||''),'base64'); if (!archive.length || archive.length>16*1024*1024) return send(response,400,{error:'Choose a valid St. Dorothy backup ZIP under 16 MB.'}); let directory; try{directory=await unzipper.Open.buffer(archive)}catch{return send(response,400,{error:'The selected file is not a valid ZIP archive.'})}; const find=name=>directory.files.find(file=>file.path===name); const manifestFile=find('backup-manifest.json'); const dataFileEntry=find('school-data.json'); if(!manifestFile||!dataFileEntry)return send(response,400,{error:'This is not a St. Dorothy backup ZIP.'}); let manifest;let restoredData;try{manifest=JSON.parse((await manifestFile.buffer()).toString('utf8'));restoredData=JSON.parse((await dataFileEntry.buffer()).toString('utf8'))}catch{return send(response,400,{error:'The backup ZIP is incomplete or damaged.'})}; if(manifest.format!=='st-dorothy-logical-backup-v1'||!restoredData||!Array.isArray(restoredData.admissions)||!Array.isArray(restoredData.accounts))return send(response,400,{error:'This backup is not in the supported St. Dorothy format.'}); await mkdir(documentUploadDirectory,{recursive:true}); let restoredDocuments=0; for(const file of directory.files.filter(item=>item.type==='File'&&item.path.startsWith('learner-documents/'))){const name=safeFileName(file.path.slice('learner-documents/'.length));if(!name||name==='learner-document')continue;await writeFile(join(documentUploadDirectory,name),await file.buffer());restoredDocuments++} addAccountAudit(restoredData,'System backup restored',request.account,`Restored backup created ${manifest.generatedAt||'on an unknown date'}; ${restoredDocuments} learner documents restored.`); await saveData(restoredData); sessions.clear(); return send(response,200,{status:'Backup restored. All users have been signed out. Sign in again to verify the restored records.',restoredDocuments,generatedAt:manifest.generatedAt||''}) }
    if (request.method === 'GET' && url.pathname === '/api/public/website-settings') return send(response,200,data.websiteSettings)
    if (request.method === 'GET' && url.pathname === '/api/system-admin/website-settings') return send(response,200,data.websiteSettings)
    if (request.method === 'PUT' && url.pathname === '/api/system-admin/website-settings') { const input=await body(request);const link=value=>/^https?:\/\//i.test(String(value||''))?text(value,300):'';data.websiteSettings={socialLinks:{facebook:link(input.socialLinks?.facebook),instagram:link(input.socialLinks?.instagram),youtube:link(input.socialLinks?.youtube),tiktok:link(input.socialLinks?.tiktok)},gallery:(Array.isArray(input.gallery)?input.gallery:[]).map(link).filter(Boolean).slice(0,6)};await saveData(data);return send(response,200,data.websiteSettings) }
    if (request.method === 'GET' && url.pathname === '/api/headteacher/dashboard') {
      const today = new Intl.DateTimeFormat('en-CA', { timeZone: 'Africa/Nairobi' }).format(new Date())
      const activeClasses = data.schoolClasses.filter(item => item.status === 'Active')
      const activeStaff = data.staff.filter(item => item.employmentStatus !== 'Inactive')
      const classAttendance = activeClasses.map(schoolClass => {
        const enrolled = data.admissions.filter(item => item.placement?.className === schoolClass.className).length
        const records = data.attendanceRecords.filter(item => item.className === schoolClass.className && item.date === today)
        const present = records.filter(item => item.status === 'Present' || item.status === 'Late').length
        return { className: schoolClass.className, enrolled, marked: records.length, present, percentage: enrolled ? Math.round(present / enrolled * 100) : null }
      })
      const staffRecords = data.staffAttendance.filter(item => item.date === today)
      const assessmentProgress = data.assessments.map(assessment => {
        const enrolled = data.admissions.filter(item => item.placement?.className === assessment.className).length
        const marks = data.assessmentMarks.filter(item => item.assessmentId === assessment.id)
        return { id: assessment.id, title: assessment.title, className: assessment.className, subjectName: assessment.subjectName, enrolled, entered: marks.length, missing: Math.max(0, enrolled - marks.length) }
      })
      return send(response, 200, { today, learners: data.admissions.length, staff: activeStaff.length, activeClasses: activeClasses.length, pendingAdmissions: data.admissions.filter(item => !item.placement || item.placement.decision !== 'Confirmed').length, unassignedClasses: activeClasses.filter(item => !item.classTeacher).length, classAttendance, incompleteAttendance: classAttendance.filter(item => item.enrolled > 0 && item.marked < item.enrolled).length, staffAttendance: { present: staffRecords.filter(item => item.status === 'Present' || item.status === 'Late').length, marked: staffRecords.length, total: activeStaff.length }, assessments: assessmentProgress, marksMissing: assessmentProgress.reduce((sum, item) => sum + item.missing, 0), lessonPlans: { review: data.lessonPlans.filter(item => item.status === 'Ready for review').length, draft: data.lessonPlans.filter(item => item.status === 'Draft').length, published: data.lessonPlans.filter(item => item.status === 'Published').length }, pendingLeave: data.leaveRequests.filter(item => item.status === 'Pending').length, recentNotices: data.notices.filter(item => item.status === 'Published').slice(0, 4) })
    }
    if (request.method === 'GET' && url.pathname === '/api/headteacher/academics') {
      const assessments = data.assessments.map(assessment => { const enrolled = data.admissions.filter(item => item.placement?.className === assessment.className).length; const marks = data.assessmentMarks.filter(item => item.assessmentId === assessment.id); const percentages = marks.map(mark => Number(assessment.outOf) > 0 ? Number(mark.score || 0) / Number(assessment.outOf) * 100 : 0); return { ...assessment, enrolled, entered: marks.length, missing: Math.max(0, enrolled - marks.length), average: percentages.length ? Math.round(percentages.reduce((sum, value) => sum + value, 0) / percentages.length) : null } })
      return send(response, 200, { assessments })
    }
    if (request.method === 'GET' && url.pathname === '/api/headteacher/learners') {
      const learners = data.admissions.map(learner => ({
        ...learner,
        finance: currentFeeSummary(data, learner)
      }))
      return send(response, 200, learners)
    }
    if (request.method === 'GET' && url.pathname === '/api/headteacher/attendance') { const date=text(url.searchParams.get('date'),10); const classes=data.schoolClasses.filter(item=>item.status==='Active').map(item=>item.className); const rows=classes.map(className=>{const enrolled=data.admissions.filter(item=>(item.placement?.className||item.grade)===className).length;const records=data.attendanceRecords.filter(item=>item.className===className&&(!date||item.date===date));const present=records.filter(item=>item.status==='Present').length;const late=records.filter(item=>item.status==='Late').length;const absent=records.filter(item=>item.status==='Absent').length;const excused=records.filter(item=>item.status==='Excused').length;return {className,enrolled,marked:records.length,present,late,absent,excused,percentage:enrolled?Math.round((present+late)/enrolled*100):null}});return send(response,200,{date,classes:rows}) }
    if (request.method === 'GET' && url.pathname === '/api/teacher/dashboard') {
      const teacherName = String(request.account.name || '').trim().toLowerCase()
      const assignments = data.teacherAssignments.filter(item => String(item.staffName || '').trim().toLowerCase() === teacherName)
      const assignedClassIds = new Set(assignments.map(item => item.classId))
      const assignedClassNames = new Set(assignments.map(item => item.className))
      data.schoolClasses.filter(item => String(item.classTeacher || '').trim().toLowerCase() === teacherName).forEach(item => { assignedClassIds.add(item.id); assignedClassNames.add(item.className) })
      const classes = data.schoolClasses.filter(item => item.status === 'Active' && (assignedClassIds.has(item.id) || assignedClassNames.has(item.className))).map(schoolClass => {
        const classAssignments = assignments.filter(item => item.classId === schoolClass.id || item.className === schoolClass.className)
        return { ...schoolClass, subjects: [...new Set(classAssignments.map(item => item.subjectName).filter(Boolean))] }
      })
      const learners = data.admissions.filter(item => assignedClassNames.has(item.placement?.className || item.grade)).map(item => ({ id: item.id, learnerName: item.learnerName, admissionNumber: item.admissionNumber, className: item.placement?.className || item.grade, residency: item.residency || 'Day scholar', medicalNotes: item.medicalNotes || item.medicalAllergyNotes || '', guardianPhone: item.guardianPhone || item.fatherPhone || item.motherPhone || '' }))
      const now = new Date()
      const today = now.toLocaleDateString('en-CA', { timeZone: 'Africa/Nairobi' })
      const day = new Intl.DateTimeFormat('en-US', { weekday: 'long', timeZone: 'Africa/Nairobi' }).format(now)
      const attendance = classes.map(schoolClass => {
        const enrolled = learners.filter(item => item.className === schoolClass.className).length
        const records = data.attendanceRecords.filter(item => item.className === schoolClass.className && item.date === today)
        const present = records.filter(item => item.status === 'Present' || item.status === 'Late').length
        return { classId: schoolClass.id, className: schoolClass.className, enrolled, marked: records.length, present, percentage: enrolled && records.length >= enrolled ? Math.round(present / enrolled * 100) : null, status: !records.length ? 'Not started' : records.length < enrolled ? 'In progress' : 'Submitted' }
      })
      const assessments = data.assessments.filter(item => assignments.some(assignment => assignment.className === item.className && (!assignment.subjectName || assignment.subjectName === item.subjectName))).map(item => {
        const enrolled = learners.filter(learner => learner.className === item.className).length
        const entered = new Set(data.assessmentMarks.filter(mark => mark.assessmentId === item.id).map(mark => mark.admissionId)).size
        return { ...item, enrolled, entered, pending: Math.max(enrolled - entered, 0) }
      })
      const timetable = data.timetableEntries.filter(item => String(item.staffName || '').trim().toLowerCase() === teacherName).sort((a, b) => String(a.time).localeCompare(String(b.time)))
      const notices = data.notices.filter(item => item.status === 'Published' && ['All staff', 'All teaching staff', 'All portal users'].includes(item.audience)).slice(0, 4)
      return send(response, 200, { today, day, classes, learners, assignments, attendance, assessments, timetable, notices })
    }
    if (request.method === 'GET' && url.pathname === '/api/system-admin/go-live') return send(response, 200, { admissions: data.admissions.length, parentAccounts: data.accounts.filter(account => account.role === 'parent').length, receipts: data.receipts.length, mpesaTransactions: data.mpesaTransactions.length, attendanceRecords: data.attendanceRecords.length, assessments: data.assessments.length, documents: data.documentRecords.length, messages: data.parentMessages.length + data.internalMessages.length })
    if (request.method === 'POST' && url.pathname === '/api/system-admin/go-live') {
      const input = await body(request)
      if (input.backupConfirmed !== true || input.confirmation !== 'GO LIVE RESET') return send(response, 400, { error: 'Confirm a verified backup and type GO LIVE RESET exactly.' })
      data.admissions=[]; data.classPlacements=[]; data.attendanceRecords=[]; data.assessments=[]; data.assessmentMarks=[]; data.receipts=[]; data.mpesaTransactions=[]; data.documentRecords=[]; data.parentMessages=[]; data.internalMessages=[]; data.smsMessages=[]; data.calendarEvents=[]; data.promotions=[]; data.admissionEnquiries=[]; data.accounts=data.accounts.filter(account=>account.role!=='parent'); data.accountAudit.unshift({ id:`audit-${Date.now()}`, action:'Go-live demo reset', by:request.account.name, createdAt:new Date().toISOString() })
      await saveData(data); return send(response, 200, { status: 'Demo records cleared. School configuration, staff accounts, classes, subjects, fee structures, and staff setup were retained.' })
    }
    if (request.method === 'GET' && url.pathname === '/api/accounts') return send(response, 200, data.accounts.filter(account => request.account.role === 'system-admin' || account.role !== 'system-admin').map(portalAccountForClient).sort((a, b) => a.name.localeCompare(b.name)))
    if (request.method === 'POST' && url.pathname === '/api/parent-portal') {
      const input = await body(request); const admissionId = text(input.admissionId, 100); const learner = data.admissions.find(item => item.id === admissionId)
      if (!learner || !normalisePhone(learner.guardianPhone)) return send(response, 400, { error: 'Select a learner with a valid guardian phone number first.' })
      const identifier = normalisePhone(learner.guardianPhone); const existing = data.accounts.find(account => account.role === 'parent' && account.identifier === identifier)
      if (existing) return send(response, 409, { error: 'A parent portal account already exists for this guardian phone number.' })
      const activation = issueActivation({ id:`account-${Date.now()}-${randomBytes(4).toString('hex')}`, identifier, role:'parent', name: learner.guardianName || 'Parent / Guardian', title:'Parent / Guardian', parentPhone:identifier, active:true, createdAt:new Date().toISOString() }); data.accounts.push(activation.account); await saveData(data)
      return send(response, 201, { account: portalAccountForClient(activation.account), activation: { code: activation.code, expiresAt: activation.account.activationExpiresAt, message: activationMessage(activation.account, activation.code, activation.account.activationExpiresAt) } })
    }
    if (request.method === 'POST' && url.pathname === '/api/accounts') {
      const input = await body(request)
      const role = text(input.role, 20)
      const allowedRoles = new Set(['headteacher', 'secretary', 'bursar', 'teacher', 'parent'])
      const name = text(input.name, 100)
      const identifier = accountIdentifier(input.identifier, role)
      const title = text(input.title, 100)
      const staffIdentifierValid = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(identifier) || (role === 'teacher' && /^254[17]\d{8}$/.test(identifier))
      if (!allowedRoles.has(role) || (request.account.role !== 'system-admin' && role === 'headteacher') || !name || !identifier || (role === 'parent' && !/^254[17]\d{8}$/.test(identifier)) || (role !== 'parent' && !staffIdentifierValid)) return send(response, 400, { error: request.account.role !== 'system-admin' && role === 'headteacher' ? 'Only the system administrator can create headteacher accounts.' : 'Enter a name and a valid work email or Kenyan mobile number.' })
      if (data.accounts.some(account => account.identifier === identifier)) return send(response, 400, { error: 'An account already exists with those sign-in details.' })
      if (role === 'parent' && !data.admissions.some(admission => normalisePhone(admission.guardianPhone) === identifier)) return send(response, 400, { error: 'Create or update the learner admission first so this parent phone is linked to a learner.' })
      const baseAccount = { id: `account-${Date.now()}-${randomBytes(4).toString('hex')}`, identifier, role, name, title: title || (role === 'parent' ? 'Parent / Guardian' : role), parentPhone: role === 'parent' ? identifier : '', active: true, createdAt: new Date().toISOString() }
      const activation = issueActivation(baseAccount)
      data.accounts.push(activation.account)
      if (role === 'teacher' && !data.staff.some(item => [String(item.email || '').trim().toLowerCase(), normalisePhone(item.phone)].includes(identifier.toLowerCase()))) data.staff.push({ id: `staff-${Date.now()}-${randomBytes(3).toString('hex')}`, fullName: name, role: 'Teacher', staffType: 'Teaching', phone: /^254[17]\d{8}$/.test(identifier) ? identifier : '', email: identifier.includes('@') ? identifier : '', employmentStatus: 'Active', createdAt: new Date().toISOString(), createdBy: request.account.name })
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
    if (request.method === 'PUT' && /^\/api\/accounts\/[^/]+\/roles$/.test(url.pathname)) {
      if (request.account.role !== 'system-admin') return send(response, 403, { error: 'Only the System Administrator can grant additional portal roles.' })
      const id = decodeURIComponent(url.pathname.split('/')[3]); const input = await body(request); const role = text(input.role, 20); const action = text(input.action, 12) || 'grant'; const account = data.accounts.find(item => item.id === id)
      if (!account) return send(response, 404, { error: 'Portal account not found.' })
      if (!['headteacher', 'secretary', 'bursar'].includes(account.role)) return send(response, 400, { error: 'Additional office roles can only be managed for existing Headteacher, Secretary, or Bursar profiles.' })
      if (!['headteacher', 'secretary', 'bursar'].includes(role) || !['grant', 'revoke'].includes(action)) return send(response, 400, { error: 'Choose a valid portal role action.' })
      if (action === 'revoke' && account.role === role) return send(response, 400, { error: 'The account’s primary role cannot be removed. Deactivate the account or assign a different primary profile instead.' })
      const currentRoles = accountRoles(account)
      const roles = action === 'grant' ? Array.from(new Set([...currentRoles, role])) : currentRoles.filter(item => item !== role)
      const updatedAccount = { ...account, roles }
      data.accounts = data.accounts.map(item => item.id === id ? updatedAccount : item)
      data.accountAudit.push({ id:`audit-${Date.now()}`, action:`portal-role-${action}`, role, accountId:id, accountName:account.name, performedBy:request.account.id, createdAt:new Date().toISOString() })
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
    if (request.method === 'GET' && url.pathname === '/api/sms/wallet') return send(response, 200, { balance: Number(data.smsWallet.balance || 0), pricePerCreditKes: Number(data.smsWallet.pricePerCreditKes || 0.8), senderId: text(process.env.SMS_SENDER_ID, 11) || '', gatewayConfigured: Boolean(process.env.BLESSEDTEXTS_API_KEY && process.env.SMS_SENDER_ID), topups: data.smsTopups.slice(0, 8) })
    if (request.method === 'GET' && url.pathname === '/api/admissions') {
      if (admissionNumbersAdded || schoolClassesAdded) await saveData(data)
      if (request.account.role === 'teacher') {
        const teacherName = String(request.account.name || '').trim().toLowerCase()
        const assignedClasses = new Set(data.teacherAssignments.filter(item => String(item.staffName || '').trim().toLowerCase() === teacherName).map(item => item.className))
        data.schoolClasses.filter(item => String(item.classTeacher || '').trim().toLowerCase() === teacherName).forEach(item => assignedClasses.add(item.className))
        return send(response, 200, data.admissions.filter(item => assignedClasses.has(item.placement?.className || item.grade)))
      }
      return send(response, 200, data.admissions)
    }
    if (request.method === 'GET' && url.pathname === '/api/secretary/learner-records') {
      return send(response, 200, data.admissions.map(learner => ({ ...learner, finance: currentFeeSummary(data, learner) })))
    }
    if (request.method === 'POST' && url.pathname === '/api/secretary/learner-records/residency') {
      const input = await body(request)
      const residency = learnerResidencies.has(text(input.residency, 20)) ? text(input.residency, 20) : ''
      const ids = Array.isArray(input.admissionIds) ? [...new Set(input.admissionIds.map(value => text(value, 100)).filter(Boolean))] : []
      if (!residency || !ids.length || ids.length > 1000) return send(response, 400, { error: 'Select one or more learners and a valid learner category.' })
      const selected = data.admissions.filter(learner => ids.includes(learner.id))
      if (selected.length !== ids.length) return send(response, 400, { error: 'One or more selected learners could not be found.' })
      const updatedAt = new Date().toISOString()
      const selectedIds = new Set(ids)
      data.admissions = data.admissions.map(learner => selectedIds.has(learner.id) ? { ...learner, residency, residencyUpdatedAt: updatedAt, residencyUpdatedBy: request.account.name, updatedAt } : learner)
      await saveData(data)
      return send(response, 200, { updated: ids.length, residency, updatedAt })
    }
    if (request.method === 'GET' && /^\/api\/secretary\/learner-records\/[^/]+\/pdf$/.test(url.pathname)) {
      const admissionId = decodeURIComponent(url.pathname.split('/')[4])
      const learner = data.admissions.find(item => item.id === admissionId)
      if (!learner) return send(response, 404, { error: 'Learner record not found.' })
      const pdf = await createLearnerRecordPdf(data, learner)
      response.writeHead(200, { 'Content-Type': 'application/pdf', 'Content-Disposition': `inline; filename="st_dorothy_learner_record_${safePdfName(learner.admissionNumber || learner.id)}.pdf"`, 'Cache-Control': 'no-store' })
      return response.end(pdf)
    }
    if (request.method === 'GET' && url.pathname === '/api/school-classes') {
      if (schoolClassesAdded) await saveData(data)
      const classes = request.account.role === 'teacher' ? data.schoolClasses.filter(item => data.teacherAssignments.some(assignment => assignment.classId === item.id && String(assignment.staffName || '').trim().toLowerCase() === String(request.account.name || '').trim().toLowerCase()) || String(item.classTeacher || '').trim().toLowerCase() === String(request.account.name || '').trim().toLowerCase()) : data.schoolClasses
      return send(response, 200, classes.sort((a, b) => a.academicYear.localeCompare(b.academicYear) || a.className.localeCompare(b.className)))
    }
    if (request.method === 'GET' && url.pathname === '/api/class-placements') {
      if (schoolClassesAdded) await saveData(data)
      return send(response, 200, { defaultClassCapacity, classes: data.schoolClasses, placements: data.classPlacements })
    }
    if (request.method === 'GET' && url.pathname === '/api/staff') {
      const teacherAccounts = data.accounts.filter(account => account.role === 'teacher' && account.active !== false)
      const staffEmails = new Set(data.staff.map(item => String(item.email || '').trim().toLowerCase()).filter(Boolean))
      const accountTeachers = teacherAccounts.filter(account => !staffEmails.has(String(account.identifier || '').trim().toLowerCase())).map(account => ({ id: `account-staff-${account.id}`, accountId: account.id, fullName: account.name, role: 'Teacher', staffType: 'Teaching', phone: '', email: account.identifier, employmentStatus: 'Active', source: 'Portal account' }))
      return send(response, 200, [...data.staff, ...accountTeachers])
    }
    if (request.method === 'GET' && url.pathname === '/api/academics') {
      if (request.account.role !== 'teacher') return send(response, 200, { subjects: data.subjects, assignments: data.teacherAssignments, timetable: data.timetableEntries, assessments: data.assessments })
      const teacherName = String(request.account.name || '').trim().toLowerCase()
      const assignments = data.teacherAssignments.filter(item => String(item.staffName || '').trim().toLowerCase() === teacherName)
      const assignmentKeys = new Set(assignments.map(item => `${item.className}|${item.subjectName}`))
      const assignedSubjectIds = new Set(assignments.map(item => item.subjectId))
      return send(response, 200, { subjects: data.subjects.filter(item => item.status !== 'Inactive' && assignedSubjectIds.has(item.id)), assignments, timetable: data.timetableEntries.filter(item => String(item.staffName || '').trim().toLowerCase() === teacherName), assessments: data.assessments.filter(item => assignmentKeys.has(`${item.className}|${item.subjectName}`)) })
    }
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
      if (request.account.role === 'teacher') {
        const teacherName = String(request.account.name || '').trim().toLowerCase()
        const allowed = data.teacherAssignments.some(item => item.className === className && String(item.staffName || '').trim().toLowerCase() === teacherName) || data.schoolClasses.some(item => item.className === className && String(item.classTeacher || '').trim().toLowerCase() === teacherName)
        if (!allowed) return send(response, 403, { error: 'You can view attendance only for your assigned classes.' })
      }
      return send(response, 200, data.attendanceRecords.filter(item => (!className || item.className === className) && (!date || item.date === date)))
    }
    if (request.method === 'GET' && url.pathname === '/api/attendance/workbook') {
      const className = text(url.searchParams.get('className'), 60)
      const date = text(url.searchParams.get('date'), 10)
      if (!className || !/^\d{4}-\d{2}-\d{2}$/.test(date)) return send(response, 400, { error: 'Choose a class and valid attendance date before downloading the Excel register.' })
      if (request.account.role === 'teacher') {
        const teacherName = String(request.account.name || '').trim().toLowerCase()
        const allowed = data.teacherAssignments.some(item => item.className === className && String(item.staffName || '').trim().toLowerCase() === teacherName) || data.schoolClasses.some(item => item.className === className && String(item.classTeacher || '').trim().toLowerCase() === teacherName)
        if (!allowed) return send(response, 403, { error: 'You can download attendance only for your assigned classes.' })
      }
      const safeClass = className.replace(/[^a-zA-Z0-9]+/g, '_').replace(/^_|_$/g, '')
      const fileName = `st_dorothy_attendance_${safeClass}_${date}.xlsx`
      const savedRecords = data.attendanceRecords.filter(item => item.className === className && item.date === date)
      const admissionById = new Map(data.admissions.map(item => [item.id, item]))
      const recordByAdmission = new Map(savedRecords.map(item => [item.admissionId, item]))
      const records = data.admissions.filter(item => (item.placement?.className || item.grade) === className).slice().sort((a,b) => a.learnerName.localeCompare(b.learnerName)).map(learner => ({ admissionId: learner.id, learnerName: learner.learnerName, status: recordByAdmission.get(learner.id)?.status || 'Not marked', reason: recordByAdmission.get(learner.id)?.reason || '', markedAt: recordByAdmission.get(learner.id)?.markedAt || null }))
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
        const row = sheet.addRow([index + 1, admissionById.get(record.admissionId)?.admissionNumber || '', record.learnerName, record.status, record.reason || '', record.markedAt ? new Date(record.markedAt) : 'Not marked'])
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
    if (request.method === 'GET' && url.pathname === '/api/class-lists/pdf') {
      const classId = text(url.searchParams.get('classId'), 100)
      const schoolClass = data.schoolClasses.find(item => item.id === classId)
      if (!schoolClass) return send(response, 404, { error: 'Choose a valid class before preparing its list.' })
      if (request.account.role === 'teacher' && !data.teacherAssignments.some(item => item.classId === schoolClass.id && item.staffName === request.account.name)) return send(response, 403, { error: 'You can print class lists only for your assigned classes.' })
      const pdf = await createClassListPdf(data, schoolClass)
      const fileName = `st_dorothy_${safePdfName(schoolClass.className)}_${safePdfName(schoolClass.academicYear)}_class_list.pdf`
      response.writeHead(200, { 'Content-Type': 'application/pdf', 'Content-Disposition': `inline; filename="${fileName}"`, 'Cache-Control': 'no-store' })
      return response.end(pdf)
    }
    if (request.method === 'GET' && url.pathname === '/api/secretary/boarders/pdf') {
      const grade = text(url.searchParams.get('grade'), 40)
      const pdf = await createBoarderListPdf(data, grade)
      const fileName = `st_dorothy_boarding_register_${safePdfName(grade || 'all_grades')}.pdf`
      response.writeHead(200, { 'Content-Type': 'application/pdf', 'Content-Disposition': `inline; filename="${fileName}"`, 'Cache-Control': 'no-store' })
      return response.end(pdf)
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
      sheet.columns = [{ width: 7 }, { width: 19 }, { width: 38 }, { width: 20 }]
      const crestPath = join(root, 'frontend', 'public', 'assets', 'st-dorothy-crest-refined.png')
      try { const crest = workbook.addImage({ buffer: await readFile(crestPath), extension: 'png' }); sheet.addImage(crest, { tl: { col: 0, row: 0 }, ext: { width: 58, height: 58 } }) } catch { /* A valid workbook is still created if the logo asset is unavailable. */ }
      sheet.mergeCells('B1:D1'); sheet.getCell('B1').value = "ST. DOROTHY'S SCHOOL MASSIMO"
      sheet.getCell('B1').font = { name: 'Aptos Display', size: 16, bold: true, color: { argb: 'FF123664' } }
      sheet.getCell('B1').alignment = { vertical: 'middle' }
      sheet.mergeCells('B2:D2'); sheet.getCell('B2').value = 'P.O. Box 36-60607, Mikinduri, Meru, Kenya · +254 722 841 569 · stdorothysschoolmassimo@gmail.com'
      sheet.getCell('B2').font = { size: 9, color: { argb: 'FF536B84' } }; sheet.getCell('B2').alignment = { wrapText: true }
      sheet.mergeCells('A4:D4'); sheet.getCell('A4').value = `CLASS LIST — ${schoolClass.className.toUpperCase()} · ${schoolClass.academicYear}`
      sheet.getCell('A4').font = { size: 13, bold: true, color: { argb: 'FFFFFFFF' } }; sheet.getCell('A4').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF123664' } }; sheet.getCell('A4').alignment = { horizontal: 'center' }
      sheet.mergeCells('A5:B5'); sheet.getCell('A5').value = `Class teacher: ${teacherName}`
      sheet.mergeCells('C5:D5'); sheet.getCell('C5').value = `Generated: ${new Intl.DateTimeFormat('en-KE', { dateStyle: 'medium', timeStyle: 'short', timeZone: 'Africa/Nairobi' }).format(new Date())}`
      for (const cell of ['A5', 'C5']) { sheet.getCell(cell).font = { bold: true, color: { argb: 'FF294A6E' } }; sheet.getCell(cell).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F6FB' } } }
      const header = sheet.addRow(['No.', 'Admission No.', 'Learner name', 'Category'])
      header.font = { bold: true, color: { argb: 'FFFFFFFF' } }; header.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1D5EAA' } }; header.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
      learners.forEach((learner, index) => { const row = sheet.addRow([index + 1, learner.admissionNumber || 'Pending', learner.learnerName, learner.residency || 'Day scholar']); row.alignment = { vertical: 'middle', wrapText: true }; if (index % 2 === 1) row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF6F9FC' } } })
      if (!learners.length) { const row = sheet.addRow(['', '', 'No learners placed in this stream yet.', '']); row.font = { italic: true, color: { argb: 'FF60758C' } } }
      sheet.eachRow(row => row.eachCell(cell => { cell.border = { top: { style: 'thin', color: { argb: 'FFD9E4EF' } }, left: { style: 'thin', color: { argb: 'FFD9E4EF' } }, bottom: { style: 'thin', color: { argb: 'FFD9E4EF' } }, right: { style: 'thin', color: { argb: 'FFD9E4EF' } } } }))
      const bytes = await workbook.xlsx.writeBuffer()
      response.writeHead(200, { 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'Content-Disposition': `attachment; filename="st_dorothy_${safeClass}_class_list_${schoolClass.academicYear}.xlsx"` })
      return response.end(Buffer.from(bytes))
    }
    if (request.method === 'GET' && url.pathname === '/api/staff-list/workbook') {
      const staffByName = new Map(data.staff.map(item => [String(item.fullName || '').toLowerCase(), item]))
      data.accounts.filter(item => item.role === 'teacher').forEach(account => { const key = String(account.name || '').toLowerCase(); if (!staffByName.has(key)) staffByName.set(key, { fullName: account.name, role: account.title || 'Teacher', phone: '', email: account.identifier, employmentStatus: account.active === false ? 'Inactive' : 'Active' }) })
      const staffMembers = [...staffByName.values()].filter(item => item.fullName).sort((a, b) => String(a.fullName).localeCompare(String(b.fullName)))
      const workbook = new ExcelJS.Workbook(); workbook.creator = 'St. Dorothy School Massimo'; workbook.created = new Date()
      const sheet = workbook.addWorksheet('School Staff', { views: [{ state: 'frozen', ySplit: 7 }] }); sheet.columns = [{ width: 7 }, { width: 30 }, { width: 16 }, { width: 25 }, { width: 29 }, { width: 18 }, { width: 30 }, { width: 15 }]
      const crestPath = join(root, 'frontend', 'public', 'assets', 'st-dorothy-crest-refined.png')
      try { const crest = workbook.addImage({ buffer: await readFile(crestPath), extension: 'png' }); sheet.addImage(crest, { tl: { col: 0, row: 0 }, ext: { width: 58, height: 58 } }) } catch { /* The report can still be generated if the crest file is unavailable. */ }
      sheet.mergeCells('B1:H1'); sheet.getCell('B1').value = "ST. DOROTHY'S SCHOOL MASSIMO"; sheet.getCell('B1').font = { name: 'Aptos Display', size: 16, bold: true, color: { argb: 'FF123664' } }; sheet.getCell('B1').alignment = { vertical: 'middle' }
      sheet.mergeCells('B2:H2'); sheet.getCell('B2').value = 'P.O. Box 36-60607, Mikinduri, Meru, Kenya · +254 722 841 569 · stdorothysschoolmassimo@gmail.com'; sheet.getCell('B2').font = { size: 9, color: { argb: 'FF536B84' } }; sheet.getCell('B2').alignment = { wrapText: true }
      sheet.mergeCells('A4:H4'); sheet.getCell('A4').value = 'SCHOOL STAFF REGISTER'; sheet.getCell('A4').font = { size: 13, bold: true, color: { argb: 'FFFFFFFF' } }; sheet.getCell('A4').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF123664' } }; sheet.getCell('A4').alignment = { horizontal: 'center' }
      sheet.mergeCells('A5:D5'); sheet.getCell('A5').value = `Total staff: ${staffMembers.length}`; sheet.mergeCells('E5:H5'); sheet.getCell('E5').value = `Generated: ${new Intl.DateTimeFormat('en-KE', { dateStyle: 'medium', timeStyle: 'short', timeZone: 'Africa/Nairobi' }).format(new Date())}`
      for (const cell of ['A5', 'E5']) { sheet.getCell(cell).font = { bold: true, color: { argb: 'FF294A6E' } }; sheet.getCell(cell).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F6FB' } } }
      const header = sheet.addRow(['No.', 'Staff name', 'Category', 'Assigned role', 'Class responsibility', 'Phone', 'Email', 'Status']); header.font = { bold: true, color: { argb: 'FFFFFFFF' } }; header.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1D5EAA' } }; header.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
      staffMembers.forEach((member, index) => { const assignments = [...new Set(data.teacherAssignments.filter(item => item.staffName === member.fullName).map(item => item.className).filter(Boolean))]; const category = member.staffType || (/teacher|headteacher/i.test(member.role || '') ? 'Teaching' : 'Support'); const row = sheet.addRow([index + 1, member.fullName, category, member.role || 'Staff', assignments.length ? assignments.join(', ') : '—', member.phone || '—', member.email || '—', member.employmentStatus || 'Active']); row.alignment = { vertical: 'middle', wrapText: true }; if (index % 2 === 1) row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF6F9FC' } } })
      if (!staffMembers.length) { const row = sheet.addRow(['', 'No school staff profiles recorded yet.', '', '', '', '', '', '']); row.font = { italic: true, color: { argb: 'FF60758C' } } }
      sheet.eachRow(row => row.eachCell(cell => { cell.border = { top: { style: 'thin', color: { argb: 'FFD9E4EF' } }, left: { style: 'thin', color: { argb: 'FFD9E4EF' } }, bottom: { style: 'thin', color: { argb: 'FFD9E4EF' } }, right: { style: 'thin', color: { argb: 'FFD9E4EF' } } } }))
      const bytes = await workbook.xlsx.writeBuffer(); response.writeHead(200, { 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'Content-Disposition': 'attachment; filename="st_dorothy_school_staff_register.xlsx"' }); return response.end(Buffer.from(bytes))
    }
    if (request.method === 'GET' && url.pathname === '/api/assessment-marks') {
      const assessmentId = text(url.searchParams.get('assessmentId'), 80)
      if (request.account.role === 'teacher') {
        const assessment = data.assessments.find(item => item.id === assessmentId)
        const teacherName = String(request.account.name || '').trim().toLowerCase()
        const allowed = assessment && data.teacherAssignments.some(item => item.className === assessment.className && item.subjectName === assessment.subjectName && String(item.staffName || '').trim().toLowerCase() === teacherName)
        if (!allowed) return send(response, 403, { error: 'You can view marks only for your assigned classes and subjects.' })
      }
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
    if (request.method === 'GET' && url.pathname === '/api/secretary/fee-lookup') { const needle=String(url.searchParams.get('query')||'').trim().toLowerCase(); if(needle.length<2)return send(response,400,{error:'Enter at least two characters to search learner balances.'}); const compact=needle.replace(/\s+/g,''); const learners=data.admissions.filter(item=>[item.learnerName,item.admissionNumber,item.guardianPhone].some(value=>String(value||'').toLowerCase().replace(/\s+/g,'').includes(compact))).slice(0,8).map(item=>({id:item.id,learnerName:item.learnerName,admissionNumber:item.admissionNumber,guardianPhone:item.guardianPhone,className:item.placement?.className||item.grade,finance:currentFeeSummary(data,item)}));return send(response,200,learners) }
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
    if (request.method === 'GET' && url.pathname === '/api/finance/defaulters') return send(response, 200, data.admissions.map(learner => ({ id:learner.id, learnerName:learner.learnerName, admissionNumber:learner.admissionNumber, className:learner.placement?.className || learner.grade, guardianName:learner.guardianName, guardianPhone:learner.guardianPhone, finance:currentFeeSummary(data, learner) })).filter(item => Number(item.finance.balance || 0) > 0).sort((a,b) => b.finance.balance - a.finance.balance))
    if (request.method === 'GET' && url.pathname === '/api/finance/budget-summary') { const rows=data.feeStructures.map(structure=>{const residency=structure.residency||'Day scholar';const learners=data.admissions.filter(learner=>learner.grade===structure.grade&&(learner.residency||'Day scholar')===residency&&learner.placement?.academicYear===structure.academicYear&&learner.placement?.term===structure.term);const summaries=learners.map(learner=>currentFeeSummary(data,learner));return {id:structure.id,academicYear:structure.academicYear,term:structure.term,grade:structure.grade,residency,learnerCount:learners.length,feePerLearner:Number(structure.amount||0),grossExpected:Number(structure.amount||0)*learners.length,bursaryImpact:summaries.reduce((sum,item)=>sum+Number(item.bursaryAmount||0),0),netExpected:summaries.reduce((sum,item)=>sum+Number(item.feeDue||0),0),collected:summaries.reduce((sum,item)=>sum+Number(item.totalPaid||0),0),arrears:summaries.reduce((sum,item)=>sum+Number(item.balance||0),0)} }).filter(row=>row.learnerCount>0);return send(response,200,rows) }
    if (request.method === 'GET' && url.pathname === '/api/finance/audit') return send(response,200,(data.financeAudit||[]).slice().sort((a,b)=>String(b.createdAt).localeCompare(String(a.createdAt))))
    if (request.method === 'GET' && url.pathname === '/api/finance/approval-requests') { const items=(data.financeApprovals||[]).filter(item=>request.account.role==='headteacher'||item.requestedBy===request.account.name); return send(response,200,items.slice().sort((a,b)=>String(b.createdAt).localeCompare(String(a.createdAt)))) }
    if (request.method === 'GET' && url.pathname === '/api/finance/adjustments') return send(response,200,(data.feeAdjustments||[]).slice().sort((a,b)=>String(b.createdAt).localeCompare(String(a.createdAt))))
    if (request.method === 'GET' && url.pathname === '/api/finance/cashbook') return send(response, 200, data.cashbookEntries.slice().sort((a,b) => String(b.date).localeCompare(String(a.date))))
    if (request.method === 'POST' && url.pathname === '/api/finance/cashbook') { const input=await body(request); const amount=Number(input.amount); const paymentMethod=text(input.paymentMethod,20); const payee=text(input.payee,120); const description=text(input.description,180); const category=text(input.category,60)||'Other'; const mpesaReference=text(input.mpesaReference,80); const bank=text(input.bank,80); const chequeNumber=text(input.chequeNumber,80); if(!['Cash','M-Pesa','Cheque'].includes(paymentMethod)||!Number.isFinite(amount)||amount<=0||!payee||!description) return send(response,400,{error:'Enter the payment method, payee, description, and a valid amount.'}); if(paymentMethod==='M-Pesa'&&!mpesaReference)return send(response,400,{error:'Enter the M-Pesa transaction reference.'}); if(paymentMethod==='Cheque'&&(!bank||!chequeNumber))return send(response,400,{error:'Enter the cheque bank and cheque number.'}); const reference=paymentMethod==='M-Pesa'?mpesaReference:paymentMethod==='Cheque'?chequeNumber:text(input.reference,80); const entry={id:`expense-${Date.now()}`,date:text(input.date,10)||new Date().toISOString().slice(0,10),type:'Cash out',paymentMethod,payee,category,description,reference,mpesaReference,bank,chequeNumber,amount,createdAt:new Date().toISOString(),createdBy:request.account.name}; data.cashbookEntries.unshift(entry);data.accountAudit.unshift({id:`audit-${Date.now()}`,action:'Expense recorded',by:request.account.name,createdAt:entry.createdAt});addFinanceAudit(data,request,'Created','Expense',entry.id,`${paymentMethod} · ${money(amount)} · ${payee} · ${reference||description}`);await saveData(data);return send(response,201,entry) }
    if (request.method === 'GET' && url.pathname === '/api/finance/expenses/pdf') { const entries=data.cashbookEntries.filter(item=>item.type==='Cash out').slice().sort((a,b)=>String(b.date).localeCompare(String(a.date))); const pdf=await createExpensesPdf(entries); response.writeHead(200,{'Content-Type':'application/pdf','Content-Disposition':'inline; filename="st_dorothy_expenditure_register.pdf"','Cache-Control':'no-store'}); return response.end(pdf) }
    if (request.method === 'GET' && url.pathname === '/api/finance/expenses/workbook') { const entries=data.cashbookEntries.filter(item=>item.type==='Cash out').slice().sort((a,b)=>String(b.date).localeCompare(String(a.date))); const workbook=new ExcelJS.Workbook();workbook.creator='St. Dorothy School Massimo';workbook.created=new Date();const sheet=workbook.addWorksheet('Expenses',{views:[{state:'frozen',ySplit:6}]});sheet.columns=[{width:14},{width:23},{width:20},{width:19},{width:18},{width:30},{width:18},{width:22}];try{const crest=workbook.addImage({buffer:await readFile(crestPath),extension:'png'});sheet.addImage(crest,{tl:{col:0,row:0},ext:{width:58,height:58}})}catch{}sheet.mergeCells('B1:H1');sheet.getCell('B1').value="ST. DOROTHY'S SCHOOL MASSIMO";sheet.getCell('B1').font={bold:true,size:16,color:{argb:'FF173662'}};sheet.mergeCells('B2:H2');sheet.getCell('B2').value='P.O. Box 36-60607, Mikinduri, Meru, Kenya · +254 722 841 569 · stdorothysschoolmassimo@gmail.com';sheet.mergeCells('A4:H4');sheet.getCell('A4').value='EXPENDITURE REGISTER';sheet.getCell('A4').font={bold:true,size:13,color:{argb:'FFFFFFFF'}};sheet.getCell('A4').fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF173662'}};sheet.getCell('A4').alignment={horizontal:'center'};sheet.mergeCells('A5:H5');sheet.getCell('A5').value=`Generated: ${new Intl.DateTimeFormat('en-KE',{dateStyle:'medium',timeStyle:'short',timeZone:'Africa/Nairobi'}).format(new Date())} · Total: ${money(entries.reduce((sum,item)=>sum+Number(item.amount||0),0))}`;const header=sheet.addRow(['Date','Payee','Category','Payment method','Reference','Bank / cheque no.','Description','Amount (KES)']);header.font={bold:true,color:{argb:'FFFFFFFF'}};header.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF1D5EAA'}};header.alignment={horizontal:'center',vertical:'middle',wrapText:true};entries.forEach((item,index)=>{const method=item.paymentMethod||(item.type==='Cash out'?'Cash':item.type);const row=sheet.addRow([item.date,item.payee||'',item.category||'Other',method,method==='M-Pesa'?item.mpesaReference||item.reference||'':item.reference||'',method==='Cheque'?`${item.bank||''} · ${item.chequeNumber||''}`.trim():item.bank||'',item.description||'',Number(item.amount||0)]);row.alignment={vertical:'middle',wrapText:true};if(index%2===1)row.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFF6F9FC'}}});sheet.getColumn(8).numFmt='#,##0';sheet.eachRow(row=>row.eachCell(cell=>{cell.border={top:{style:'thin',color:{argb:'FFD9E4EF'}},left:{style:'thin',color:{argb:'FFD9E4EF'}},bottom:{style:'thin',color:{argb:'FFD9E4EF'}},right:{style:'thin',color:{argb:'FFD9E4EF'}}}}));const bytes=await workbook.xlsx.writeBuffer();response.writeHead(200,{'Content-Type':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','Content-Disposition':'attachment; filename="st_dorothy_expenditure_register.xlsx"'});return response.end(Buffer.from(bytes)) }
    if (request.method === 'GET' && url.pathname === '/api/finance/bank-deposits') return send(response, 200, data.bankDeposits.slice().sort((a,b) => String(b.depositDate).localeCompare(String(a.depositDate))))
    if (request.method === 'POST' && url.pathname === '/api/finance/bank-deposits') { const input=await body(request);const amount=Number(input.amount);if(!Number.isFinite(amount)||amount<=0||!text(input.bank,80)||!text(input.slipNumber,80))return send(response,400,{error:'Enter the bank, deposit slip number, and a valid amount.'});const deposit={id:`deposit-${Date.now()}`,depositDate:text(input.depositDate,10)||new Date().toISOString().slice(0,10),bank:text(input.bank,80),slipNumber:text(input.slipNumber,80),amount,notes:text(input.notes,180),createdAt:new Date().toISOString(),createdBy:request.account.name};data.bankDeposits.unshift(deposit);data.accountAudit.unshift({id:`audit-${Date.now()}`,action:'Bank deposit recorded',by:request.account.name,createdAt:deposit.createdAt});addFinanceAudit(data,request,'Created','Bank deposit',deposit.id,`${deposit.bank} · ${deposit.slipNumber} · ${money(amount)}`);await saveData(data);return send(response,201,deposit) }
    if (request.method === 'POST' && url.pathname === '/api/finance/adjustments') { const input=await body(request);const learner=data.admissions.find(item=>item.id===text(input.admissionId,80));const amount=Number(input.amount);const placement=learner?.placement;if(!learner||!placement?.academicYear||!placement?.term||!Number.isFinite(amount)||amount<=0)return send(response,400,{error:'Choose a placed learner and valid bursary or discount amount.'});const requestItem={id:`approval-${Date.now()}`,type:'Bursary / discount',status:'Pending Headteacher approval',admissionId:learner.id,learnerName:learner.learnerName,admissionNumber:learner.admissionNumber,className:placement.className||learner.grade,academicYear:placement.academicYear,term:placement.term,amount,reason:text(input.reason,180)||'Bursary / approved discount',createdAt:new Date().toISOString(),requestedBy:request.account.name};data.financeApprovals.unshift(requestItem);addFinanceAudit(data,request,'Requested approval','Bursary / discount',requestItem.id,`${requestItem.learnerName} · ${money(amount)} · ${requestItem.term} ${requestItem.academicYear}`);await saveData(data);return send(response,201,requestItem) }
    if (request.method === 'PUT' && url.pathname.startsWith('/api/finance/approval-requests/')) { const id=decodeURIComponent(url.pathname.slice('/api/finance/approval-requests/'.length));const item=data.financeApprovals.find(entry=>entry.id===id);const input=await body(request);const decision=text(input.decision,12);if(!item)return send(response,404,{error:'Approval request not found.'});if(item.status!=='Pending Headteacher approval')return send(response,400,{error:'This request has already been decided.'});if(!['Approved','Rejected'].includes(decision))return send(response,400,{error:'Choose Approved or Rejected.'});item.status=decision;item.decidedAt=new Date().toISOString();item.decidedBy=request.account.name;item.decisionNote=text(input.decisionNote,180);if(decision==='Approved'&&item.type==='Bursary / discount'){const adjustment={id:`adjustment-${Date.now()}`,admissionId:item.admissionId,learnerName:item.learnerName,admissionNumber:item.admissionNumber,className:item.className,academicYear:item.academicYear,term:item.term,amount:item.amount,reason:item.reason,createdAt:item.decidedAt,createdBy:`Approved by ${request.account.name}`};data.feeAdjustments.unshift(adjustment);addFinanceAudit(data,request,'Approved','Bursary / discount',adjustment.id,`${item.learnerName} · ${money(item.amount)} · requested by ${item.requestedBy}`)}else addFinanceAudit(data,request,'Rejected','Bursary / discount',item.id,`${item.learnerName} · ${money(item.amount)} · ${item.decisionNote||'No reason recorded'}`);await saveData(data);return send(response,200,item) }
    if (request.method === 'GET' && url.pathname === '/api/finance/defaulters/pdf') { const className = text(url.searchParams.get('className'), 80); const pdf = await createDefaulterListPdf(data, className); const fileName = `st_dorothy_defaulters_${safePdfName(className || 'all_classes')}.pdf`; response.writeHead(200, { 'Content-Type': 'application/pdf', 'Content-Disposition': `inline; filename="${fileName}"`, 'Cache-Control': 'no-store' }); return response.end(pdf) }
    if (request.method === 'GET' && url.pathname === '/api/finance/operations-workbook') { const className=text(url.searchParams.get('className'),80); const report=text(url.searchParams.get('report'),20)||'all'; if(!['all','defaulters','deposits','cashbook'].includes(report)) return send(response,400,{error:'Invalid finance report requested.'}); const workbook=new ExcelJS.Workbook();workbook.creator='St. Dorothy School Massimo';const addTitle=async(sheet,title,headers)=>{sheet.columns=headers.map(()=>({width:22}));try{const crest=workbook.addImage({buffer:await readFile(crestPath),extension:'png'});sheet.addImage(crest,{tl:{col:0,row:0},ext:{width:54,height:54}})}catch{}sheet.mergeCells(`B1:${String.fromCharCode(64+headers.length)}1`);sheet.getCell('B1').value="ST. DOROTHY'S SCHOOL MASSIMO";sheet.getCell('B1').font={bold:true,size:16,color:{argb:'173662'}};sheet.mergeCells(`B2:${String.fromCharCode(64+headers.length)}2`);sheet.getCell('B2').value=`${title} · Generated ${new Date().toLocaleString('en-KE')}`;sheet.getRow(4).values=headers;sheet.getRow(4).font={bold:true,color:{argb:'FFFFFF'}};sheet.getRow(4).fill={type:'pattern',pattern:'solid',fgColor:{argb:'1D5EAA'}};sheet.views=[{state:'frozen',ySplit:4}]};if(report==='all'||report==='defaulters'){const def=workbook.addWorksheet('Defaulters');await addTitle(def,'Outstanding fee balances',['Learner','Admission number','Class','Guardian','Phone','Fee due','Paid','Balance']);data.admissions.map(l=>({l,f:currentFeeSummary(data,l)})).filter(({l,f})=>(!className||(l.placement?.className||l.grade)===className)&&Number(f.balance||0)>0).forEach(({l,f})=>def.addRow([l.learnerName,l.admissionNumber,l.placement?.className||l.grade,l.guardianName,l.guardianPhone,f.feeDue||0,f.totalPaid||0,f.balance||0]))}if(report==='all'||report==='deposits'){const dep=workbook.addWorksheet('Bank Deposits');await addTitle(dep,'Bank deposits',['Date','Bank','Deposit slip','Amount (KES)','Notes','Recorded by']);data.bankDeposits.forEach(i=>dep.addRow([i.depositDate,i.bank,i.slipNumber,i.amount,i.notes,i.createdBy]))}if(report==='all'||report==='cashbook'){const cash=workbook.addWorksheet('Cashbook Expenses');await addTitle(cash,'Daily cashbook expenses',['Date','Type','Description','Reference','Amount (KES)','Recorded by']);data.cashbookEntries.filter(i=>i.type==='Cash out').forEach(i=>cash.addRow([i.date,i.type,i.description,i.reference,i.amount,i.createdBy]))}const filenames={all:'st_dorothy_finance_operations.xlsx',defaulters:`st_dorothy_defaulters_${safePdfName(className||'all_classes')}.xlsx`,deposits:'st_dorothy_bank_deposits.xlsx',cashbook:'st_dorothy_cashbook_expenses.xlsx'};const bytes=await workbook.xlsx.writeBuffer();response.writeHead(200,{'Content-Type':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','Content-Disposition':`attachment; filename="${filenames[report]}"`});return response.end(Buffer.from(bytes)) }
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
      for (const learner of learners) archive.append(await createFeeStatementPdf(data, learner, request.account.name), { name: `fee-statements/${safePdfName(learner.admissionNumber || learner.id)}_${safePdfName(learner.learnerName)}.pdf` })
      archive.append(Buffer.from(deliveryRegister), { name: 'bursar-delivery-register.xlsx' }); await archive.finalize(); return
    }
    if (request.method === 'GET' && /^\/api\/finance\/statements\/[^/]+\/pdf$/.test(url.pathname)) {
      const admissionId = decodeURIComponent(url.pathname.split('/')[4]); const learner = data.admissions.find(item => item.id === admissionId)
      if (!learner) return send(response, 404, { error: 'Learner account not found.' })
      const pdf = await createFeeStatementPdf(data, learner, request.account.name); response.writeHead(200, { 'Content-Type': 'application/pdf', 'Content-Disposition': `inline; filename="st_dorothy_fee_statement_${safePdfName(learner.admissionNumber || learner.id)}.pdf"` }); return response.end(pdf)
    }
    if (request.method === 'GET' && /^\/api\/finance\/statements\/[^/]+$/.test(url.pathname)) {
      const admissionId = decodeURIComponent(url.pathname.split('/')[4])
      const learner = data.admissions.find(item => item.id === admissionId)
      if (!learner) return send(response, 404, { error: 'Learner account not found.' })
      const finance = currentFeeSummary(data, learner)
      const receipts = data.receipts.filter(item => item.admissionId === learner.id).slice().sort((a, b) => String(a.issuedAt).localeCompare(String(b.issuedAt)))
      return send(response, 200, { statementNumber: `SD/STM/${new Date().getFullYear()}/${learner.admissionNumber || learner.id}`, generatedAt: new Date().toISOString(), learner: { learnerName: learner.learnerName, admissionNumber: learner.admissionNumber, grade: learner.grade, className: learner.placement?.className || learner.grade, guardianName: learner.guardianName, guardianPhone: learner.guardianPhone }, finance, receipts })
    }
    if (request.method === 'GET' && /^\/api\/finance\/receipts\/[^/]+\/pdf$/.test(url.pathname)) {
      const receiptId = decodeURIComponent(url.pathname.split('/')[4]); const receipt = data.receipts.find(item => item.id === receiptId)
      if (!receipt) return send(response, 404, { error: 'Payment receipt not found.' })
      const printableReceipt = receiptForPrint(data, receipt, request.account.name)
      const pdf = await createReceiptPdf(printableReceipt); response.writeHead(200, { 'Content-Type': 'application/pdf', 'Content-Disposition': `${url.searchParams.has('download') ? 'attachment' : 'inline'}; filename="st_dorothy_receipt_${safePdfName(receipt.receiptNumber || receipt.id)}.pdf"` }); return response.end(pdf)
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
      const channels = ['M-Pesa', 'Equity Bank', 'DTB Bank', 'Trans Nation SACCO', 'Cash']
      const byMethod = [...channels, ...[...new Set(receipts.map(receipt => receipt.paymentMethod).filter(method => method && !channels.includes(method)))]] .map(method => {
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
      const schoolClass = { id: `class-${Date.now()}`, className, grade, stream, academicYear, capacity, classTeacher, status, createdAt: new Date().toISOString(), createdBy: request.account.name || request.account.role }
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
      const staffType = ['Teaching', 'Support'].includes(text(input.staffType, 20)) ? text(input.staffType, 20) : (/teacher|headteacher/i.test(role) ? 'Teaching' : 'Support')
      const staffMember = { id: `staff-${Date.now()}`, fullName, role, staffType, phone, email: text(input.email, 100), employmentStatus: text(input.employmentStatus, 30) || 'Active', createdAt: new Date().toISOString() }
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
      const subject = { id: `subject-${Date.now()}`, name, grade, status: ['Active', 'Inactive'].includes(text(input.status, 20)) ? text(input.status, 20) : 'Active', optional: false, source: 'School custom subject', createdAt: new Date().toISOString(), updatedAt: new Date().toISOString() }
      data.subjects.push(subject)
      await saveData(data)
      return send(response, 201, subject)
    }

    if (request.method === 'PUT' && url.pathname.startsWith('/api/subjects/')) {
      const subjectId = decodeURIComponent(url.pathname.split('/').pop())
      const subject = data.subjects.find(item => item.id === subjectId)
      if (!subject) return send(response, 404, { error: 'Subject not found.' })
      const input = await body(request)
      const name = text(input.name, 70)
      const grade = text(input.grade, 20)
      const status = text(input.status, 20)
      if (!name || !schoolGrades.has(grade) || !['Active', 'Inactive'].includes(status)) return send(response, 400, { error: 'Subject name, grade, and status are required.' })
      if (data.subjects.some(item => item.id !== subjectId && item.name.toLowerCase() === name.toLowerCase() && item.grade === grade)) return send(response, 400, { error: 'That subject has already been added for this grade.' })
      const previousName = subject.name
      Object.assign(subject, { name, grade, status, updatedAt: new Date().toISOString(), updatedBy: request.account.name })
      if (previousName !== name) {
        data.teacherAssignments.filter(item => item.subjectId === subjectId).forEach(item => { item.subjectName = name })
        data.assessments.filter(item => item.subjectName === previousName).forEach(item => { item.subjectName = name })
        data.lessonPlans.filter(item => item.subjectName === previousName).forEach(item => { item.subjectName = name })
      }
      await saveData(data)
      return send(response, 200, subject)
    }

    if (request.method === 'POST' && url.pathname === '/api/teacher-assignments') {
      const input = await body(request)
      const staffId = text(input.staffId, 100)
      const staff = data.staff.find(item => item.id === staffId) || (() => { const accountId = staffId.startsWith('account-staff-') ? staffId.slice('account-staff-'.length) : ''; const account = data.accounts.find(item => item.id === accountId && item.role === 'teacher' && item.active !== false); return account ? { id: staffId, fullName: account.name } : null })()
      const schoolClass = data.schoolClasses.find(item => item.id === text(input.classId, 80))
      const subject = data.subjects.find(item => item.id === text(input.subjectId, 80))
      if (!staff || !schoolClass || !subject) return send(response, 400, { error: 'Choose a staff member, class, and subject.' })
      if (subject.status === 'Inactive') return send(response, 400, { error: 'Activate this subject before assigning a teacher.' })
      const classGrade = String(schoolClass.className || '').match(/^(PP1|PP2|Grade [1-9])/i)?.[1] || schoolClass.grade
      if (String(subject.grade).toLowerCase() !== String(classGrade).toLowerCase()) return send(response, 400, { error: `Choose a ${classGrade} subject for ${schoolClass.className}.` })
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
      if (request.account.role === 'teacher' && !data.teacherAssignments.some(item => item.className === className && item.subjectName === subjectName && String(item.staffName || '').trim().toLowerCase() === String(request.account.name || '').trim().toLowerCase())) return send(response, 403, { error: 'You can create assessments only for your assigned classes and subjects.' })
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
      if (request.account.role === 'teacher' && !data.teacherAssignments.some(item => item.className === assessment.className && item.subjectName === assessment.subjectName && String(item.staffName || '').trim().toLowerCase() === String(request.account.name || '').trim().toLowerCase())) return send(response, 403, { error: 'You can enter marks only for your assigned classes and subjects.' })
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
      if (request.account.role === 'teacher') {
        const teacherName = String(request.account.name || '').trim().toLowerCase()
        const allowed = data.teacherAssignments.some(item => item.className === className && String(item.staffName || '').trim().toLowerCase() === teacherName) || data.schoolClasses.some(item => item.className === className && String(item.classTeacher || '').trim().toLowerCase() === teacherName)
        if (!allowed) return send(response, 403, { error: 'You can save attendance only for your assigned classes.' })
      }
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
        reviewNote: request.account.role === 'headteacher' ? text(input.reviewNote, 600) : existing.reviewNote,
        reviewedBy: request.account.role === 'headteacher' ? request.account.name : existing.reviewedBy,
        reviewedAt: request.account.role === 'headteacher' ? new Date().toISOString() : existing.reviewedAt,
        updatedAt: new Date().toISOString()
      }
      data.lessonPlans = data.lessonPlans.map(item => item.id === id ? updated : item)
      await saveData(data)
      return send(response, 200, updated)
    }

    if (request.method === 'POST' && url.pathname === '/api/fee-structures') {
      const input = await body(request)
      const grade = text(input.grade, 20)
      const residency = learnerResidencies.has(text(input.residency, 20)) ? text(input.residency, 20) : 'Day scholar'
      const academicYear = text(input.academicYear, 10)
      const term = text(input.term, 20)
      const amount = Number(input.amount)
      if (!schoolGrades.has(grade) || !/^\d{4}$/.test(academicYear) || !term || !Number.isFinite(amount) || amount < 0) return send(response, 400, { error: 'Choose a grade, academic year, term, and valid fee amount.' })
      const structure = { id: `fees-${Date.now()}`, grade, residency, academicYear, term, amount, updatedAt: new Date().toISOString(), setBy: 'Bursar' }
      data.feeStructures = [...data.feeStructures.filter(item => !(item.grade === grade && (item.residency || 'Day scholar') === residency && item.term === term && item.academicYear === academicYear)), structure]; addFinanceAudit(data,request,'Changed','Fee plan',structure.id,`${grade} · ${residency} · ${term} ${academicYear} · ${money(amount)}`)
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

    if (request.method === 'PUT' && url.pathname.startsWith('/api/leave-requests/')) {
      const id = decodeURIComponent(url.pathname.slice('/api/leave-requests/'.length))
      const existing = data.leaveRequests.find(item => item.id === id)
      if (!existing) return send(response, 404, { error: 'Leave request not found.' })
      const input = await body(request)
      const decision = text(input.decision, 20)
      if (!['Approved', 'Rejected'].includes(decision)) return send(response, 400, { error: 'Choose Approve or Reject.' })
      const updated = { ...existing, status: decision, decisionNote: text(input.note, 500), decidedBy: request.account.name, decidedAt: new Date().toISOString() }
      data.leaveRequests = data.leaveRequests.map(item => item.id === id ? updated : item)
      await saveData(data)
      return send(response, 200, updated)
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

    if (request.method === 'POST' && url.pathname === '/api/sms/topups') {
      const input = await body(request); const credits = Number(input.credits); const paymentReference = text(input.paymentReference, 80); const paymentChannel = text(input.paymentChannel, 40)
      if (![100, 500, 1000, 2000, 5000].includes(credits) || !paymentReference || !paymentChannel) return send(response, 400, { error: 'Choose an SMS credit pack and enter the verified payment channel and reference.' })
      if (data.smsTopups.some(item => item.paymentChannel.toLowerCase() === paymentChannel.toLowerCase() && item.paymentReference.toLowerCase() === paymentReference.toLowerCase())) return send(response, 409, { error: 'This SMS top-up payment reference has already been recorded.' })
      const amountKes = Math.round(credits * Number(data.smsWallet.pricePerCreditKes || 0.8) * 100) / 100
      const topup = { id: `sms-topup-${Date.now()}`, credits, amountKes, paymentChannel, paymentReference, status: 'Awaiting payment confirmation', createdAt: new Date().toISOString(), createdBy: request.account.name }
      data.smsTopups.unshift(topup); data.accountAudit.unshift({ id: `audit-${Date.now()}`, action: 'SMS top-up requested', by: request.account.name, createdAt: topup.createdAt }); await saveData(data)
      return send(response, 201, { topup, balance: data.smsWallet.balance })
    }
    if (request.method === 'POST' && url.pathname === '/api/sms') {
      const input = await body(request); const content = text(input.text, 480); const selection = smsRecipients(data, input); const segments = smsSegments(content); const credits = segments * selection.recipients.length
      if (!content || !selection.recipients.length) return send(response, 400, { error: 'Choose a learner, grade, or parent group with at least one valid parent phone number.' })
      const message = { id: `sms-${Date.now()}`, audience: selection.audience, count: selection.recipients.length, text: content, segments, credits, status: 'Pending gateway setup', createdAt: new Date().toISOString(), createdBy: request.account.name }
      const gatewayConfigured = Boolean(process.env.BLESSEDTEXTS_API_KEY && process.env.SMS_SENDER_ID)
      if (!gatewayConfigured) return send(response, 409, { error: 'The school SMS service is not ready. Ask System Administration to complete the BlessedTexts and Sender ID setup first.' })
      if (Number(data.smsWallet.balance || 0) < credits) return send(response, 409, { error: `Insufficient SMS credits. This message needs ${credits} credit${credits === 1 ? '' : 's'}.` })
      data.smsWallet.balance -= credits; message.status = 'Sending'; data.smsMessages.unshift(message); await saveData(data)
      const outcomes = await Promise.allSettled(selection.recipients.map(recipient => sendBlessedText(recipient.phone, content)))
      const accepted = outcomes.filter(item => item.status === 'fulfilled').length; const failed = outcomes.length - accepted
      if (failed) data.smsWallet.balance += failed * segments
      message.status = failed ? (accepted ? 'Partially accepted' : 'Not accepted') : 'Accepted by gateway'; message.accepted = accepted; message.failed = failed; message.senderId = text(process.env.SMS_SENDER_ID, 11); message.updatedAt = new Date().toISOString()
      data.accountAudit.unshift({ id: `audit-${Date.now()}`, action: 'SMS send requested', by: request.account.name, createdAt: message.updatedAt }); await saveData(data)
      return send(response, 201, message)
    }

    if (request.method === 'POST' && url.pathname === '/api/learner-import/preview') { const input=await body(request); const rows=await parseLearnerWorkbook(input.fileBase64,data); return send(response,200,{rows,summary:{total:rows.length,ready:rows.length,review:rows.filter(row=>row.issues.some(issue=>issue.includes('Duplicate')||issue.includes('already exists'))).length}}) }
    if (request.method === 'POST' && url.pathname === '/api/learner-import/commit') { const input=await body(request); const rows=Array.isArray(input.rows)?input.rows:[]; if(!rows.length||rows.length>1200)return send(response,400,{error:'No valid learner import rows were supplied.'}); const createdAt=new Date().toISOString(); const imported=[]; for(const row of rows){const name=text(row.learnerName,100),grade=text(row.grade,20);if(!name||!grade||data.admissions.some(item=>item.learnerName===name&&item.grade===grade))continue;let number=text(row.admissionNumber,40);if(!number)number=nextAdmissionNumber(createdAt,data.admissions);else if(data.admissions.some(item=>item.admissionNumber===number))number=repeatedAdmissionNumber(number,data.admissions);const guardians=Array.isArray(row.guardians)?row.guardians.slice(0,2).map(item=>({name:text(item.name,100),phone:normalisePhone(item.phone),relationship:text(item.relationship,40)})):[];const primary=guardians.find(item=>/^2547\d{8}$/.test(item.phone))||guardians[0]||{name:'Guardian contact pending',phone:'',relationship:'Guardian'};const application={id:`admission-${Date.now()}-${imported.length}`,admissionNumber:number,sourceAdmissionNumber:text(row.admissionNumber,40),learnerName:name,grade,guardianName:primary.name||'Guardian contact pending',guardianPhone:primary.phone||'',relationship:primary.relationship||'Guardian',guardians,dateOfBirth:text(row.dateOfBirth,10),medicalNotes:'',birthCertificateProvided:false,status:'Imported — placement pending',createdAt,importedAt:createdAt};data.admissions.push(application);imported.push(application)} await saveData(data); return send(response,201,{imported:imported.length,skipped:rows.length-imported.length}) }
    if (request.method === 'POST' && url.pathname === '/api/teacher-import/preview') { const input=await body(request); const rows=await parseTeacherWorkbook(input.fileBase64,data); return send(response,200,{rows,summary:{total:rows.length,ready:rows.filter(row=>row.ready).length,review:rows.filter(row=>!row.ready).length}}) }
    if (request.method === 'POST' && url.pathname === '/api/teacher-import/commit') {
      const input = await body(request)
      const rows = Array.isArray(input.rows) ? input.rows : []
      if (!rows.length || rows.length > 200) return send(response, 400, { error: 'No valid teacher import rows were supplied.' })
      const activations = []
      let skipped = 0
      for (const row of rows) {
        const name = text(row.name, 100)
        const identifier = normalisePhone(row.phone)
        if (!name || !/^254[17]\d{8}$/.test(identifier) || data.accounts.some(item => item.identifier === identifier) || data.staff.some(item => normalisePhone(item.phone) === identifier)) { skipped++; continue }
        const createdAt = new Date().toISOString()
        const baseAccount = { id:`account-${Date.now()}-${randomBytes(4).toString('hex')}`, identifier, role:'teacher', name, title:'Teacher', parentPhone:'', active:true, createdAt }
        const activation = issueActivation(baseAccount)
        data.accounts.push(activation.account)
        data.staff.push({ id:`staff-${Date.now()}-${randomBytes(3).toString('hex')}`, fullName:name, role:'Teacher', staffType:'Teaching', phone:identifier, email:'', employmentStatus:'Active', createdAt, createdBy:request.account.name })
        addAccountAudit(data, 'Teacher account imported', request.account, `${name} · ${identifier}`)
        activations.push({ name, identifier, code:activation.code, expiresAt:activation.account.activationExpiresAt })
      }
      await saveData(data)
      return send(response, 201, { imported: activations.length, skipped, activations })
    }
    if (request.method === 'POST' && url.pathname === '/api/admissions') {
      const input = await body(request)
      const createdAt = new Date().toISOString()
      const fatherName = text(input.fatherName, 100), fatherPhone = text(input.fatherPhone, 20), motherName = text(input.motherName, 100), motherPhone = text(input.motherPhone, 20), residency = learnerResidencies.has(text(input.residency, 20)) ? text(input.residency, 20) : 'Day scholar'
      const primaryName = motherName || fatherName, primaryPhone = motherPhone || fatherPhone, primaryRelationship = motherName ? 'Mother / guardian' : 'Father / guardian'
      const application = {
        id: `admission-${Date.now()}`,
        admissionNumber: nextAdmissionNumber(createdAt, data.admissions),
        learnerName: text(input.learnerName, 100),
        grade: text(input.grade, 20), residency, admissionFee: 500, admissionFeeStatus: 'Unpaid',
        fatherName, fatherPhone, motherName, motherPhone,
        guardianName: primaryName,
        guardianPhone: primaryPhone,
        relationship: primaryRelationship,
        guardians: [{ name: fatherName, phone: fatherPhone, relationship: 'Father' }, { name: motherName, phone: motherPhone, relationship: 'Mother / guardian' }].filter(item => item.name || item.phone),
        medicalNotes: text(input.medicalNotes, 500),
        birthCertificateProvided: Boolean(input.birthCertificateProvided),
        status: 'Application received',
        createdAt
      }
      if (!application.learnerName || !application.grade || !application.guardianName || !application.guardianPhone) return send(response, 400, { error: 'Learner name, applying grade, and at least one parent or guardian contact are required.' })
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
      const paymentDate = text(input.paymentDate, 10)
      if (!learner) return send(response, 404, { error: 'Select a valid learner account before issuing the receipt.' })
      if (!Number.isFinite(amount) || amount <= 0) return send(response, 400, { error: 'Enter a valid payment amount.' })
      if (!paymentMethod || !paymentReference) return send(response, 400, { error: 'Payment method and reference are required.' })
      if (data.receipts.some(receipt => String(receipt.paymentMethod || '').toLowerCase() === paymentMethod.toLowerCase() && String(receipt.paymentReference || '').toLowerCase() === paymentReference.toLowerCase())) return send(response, 409, { error: 'This payment reference has already been confirmed. Check the saved receipts before recording it again.' })
      const financeBefore = currentFeeSummary(data, learner)
      const arrearsDue = financeBefore.importedBalance ? Math.max(0, Number(financeBefore.balance || 0)) : 0
      const arrearsApplied = Math.min(amount, arrearsDue)
      const manualAllocations = (Array.isArray(input.allocations) ? input.allocations : []).map(item => ({ voteHead:text(item.voteHead,80), amount:Number(item.amount), note:text(item.note,100) })).filter(item => receiptVoteHeads.has(item.voteHead) && Number.isFinite(item.amount) && item.amount > 0)
      const manuallyAllocated = manualAllocations.reduce((sum,item)=>sum+item.amount,0)
      const allocationRemainder = Math.round((amount-arrearsApplied-manuallyAllocated)*100)/100
      if (Math.abs(allocationRemainder) > 0.001) return send(response,400,{error:`Allocate the remaining ${money(Math.max(0,amount-arrearsApplied))} to one or more vote heads before confirming this receipt.`})
      const allocations = [...(arrearsApplied > 0 ? [{ voteHead:'Arrears', amount:arrearsApplied, automatic:true }] : []), ...manualAllocations]
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
        allocations,
        arrearsBeforePayment: arrearsDue,
        arrearsApplied,
        paymentMethod,
        paymentReference,
        paymentDate: /^\d{4}-\d{2}-\d{2}$/.test(paymentDate) ? paymentDate : issuedAt.slice(0, 10),
        confirmationSource: 'Manual school notification',
        servedBy: request.account.name || 'School office',
        issuedAt
      }
      data.receipts.unshift(receipt)
      receipt.remainingBalance = currentFeeSummary(data, learner).balance
      addFinanceAudit(data,request,'Confirmed','Manual payment',receipt.id,`${receipt.receiptNumber} · ${receipt.learnerName} · ${money(amount)} · ${paymentMethod} · ${paymentReference}`)
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
      addFinanceAudit(data,request,'Created','M-Pesa confirmation',transaction.id,`${transactionCode} · ${money(amount)} · awaiting match`)
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
      addFinanceAudit(data,request,'Created','Parent M-Pesa report',transaction.id,`${transactionCode} · ${learner.learnerName} · ${money(amount)}`)
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
      addFinanceAudit(data,request,'Reconciled','M-Pesa confirmation',transaction.id,`${transaction.transactionCode} · ${learner.learnerName} · ${money(transaction.amount)} · receipt ${receipt.receiptNumber}`)
      await saveData(data)
      return send(response, 200, { transaction: updatedTransaction, receipt })
    }

    if (request.method === 'GET' && url.pathname === '/api/class-placements/suggestions') {
      const academicYear = text(url.searchParams.get('academicYear'), 10) || data.schoolClasses.map(item => item.academicYear).sort().at(-1) || ''
      return send(response, 200, { academicYear, suggestions: placementSuggestions(data, academicYear) })
    }

    if (request.method === 'POST' && url.pathname === '/api/class-placements/auto-place') {
      const input = await body(request)
      const academicYear = text(input.academicYear, 10)
      const term = text(input.term, 20) || 'Term 1'
      const requestedIds = new Set(Array.isArray(input.admissionIds) ? input.admissionIds.map(item => text(item, 80)).filter(Boolean) : [])
      if (!/^\d{4}$/.test(academicYear) || !['Term 1', 'Term 2', 'Term 3'].includes(term)) return send(response, 400, { error: 'Choose a valid academic year and term.' })
      const suggestions = placementSuggestions(data, academicYear).filter(item => (!requestedIds.size || requestedIds.has(item.admissionId)) && item.classId)
      if (!suggestions.length) return send(response, 400, { error: 'There are no eligible learners with space available for automatic placement.' })
      const timestamp = new Date().toISOString()
      const placements = suggestions.map((suggestion, index) => {
        const learner = data.admissions.find(item => item.id === suggestion.admissionId)
        return { id: `placement-${Date.now()}-${index + 1}`, admissionId: learner.id, admissionNumber: learner.admissionNumber, learnerName: learner.learnerName, requestedGrade: learner.grade, className: suggestion.className, academicYear, term, decision: 'Confirmed', assessmentResult: '', notes: `Placement assistant: ${suggestion.reason}`, assignedBy: request.account.name || 'School office', assignedAt: timestamp, updatedAt: timestamp }
      })
      const placedIds = new Set(placements.map(item => item.admissionId))
      data.classPlacements = [...placements, ...data.classPlacements]
      data.admissions = data.admissions.map(item => { const placement = placements.find(next => next.admissionId === item.id); return placement ? { ...item, placement, status: `Placement confirmed · ${placement.className}`, updatedAt: timestamp } : item })
      await saveData(data)
      return send(response, 201, { placed: placements.length, skipped: requestedIds.size ? requestedIds.size - placedIds.size : 0, placements })
    }

    if (request.method === 'POST' && url.pathname === '/api/class-placements') {
      const input = await body(request)
      const admissionId = text(input.admissionId, 80)
      const classId = text(input.classId, 80)
      const decision = 'Confirmed'
      const admission = data.admissions.find(application => application.id === admissionId)
      if (!admission) return send(response, 404, { error: 'The selected admission application was not found.' })
      const academicYear = text(input.academicYear, 10)
      const schoolClass = data.schoolClasses.find(item => item.id === classId && item.academicYear === academicYear)
      if (!schoolClass || schoolClass.status !== 'Active') return send(response, 400, { error: 'Select an active class created for the chosen academic year.' })
      if (schoolClass.grade !== admission.grade) return send(response, 400, { error: `Choose a class set up for ${admission.grade}.` })
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
      const requestedAdmissionNumber = text(input.admissionNumber, 40)
      if (requestedAdmissionNumber && requestedAdmissionNumber !== existing.admissionNumber && !/^\d+(?:R\d*)?$/.test(requestedAdmissionNumber)) return send(response, 400, { error: 'Use a numeric admission number, with R only for an approved repeated record (for example 2342R).' })
      if (requestedAdmissionNumber && data.admissions.some(application => application.id !== id && application.admissionNumber === requestedAdmissionNumber)) return send(response, 409, { error: 'That admission number is already assigned to another learner.' })
      const savedFather = existing.fatherName || existing.guardians?.find(item => /father/i.test(item.relationship || '')) || {}, savedMother = existing.motherName || existing.guardians?.find(item => /mother/i.test(item.relationship || '')) || {}
      const fatherName = text(Object.hasOwn(input, 'fatherName') ? input.fatherName : (typeof savedFather === 'string' ? savedFather : savedFather.name), 100), fatherPhone = text(Object.hasOwn(input, 'fatherPhone') ? input.fatherPhone : (typeof savedFather === 'string' ? '' : savedFather.phone), 20), motherName = text(Object.hasOwn(input, 'motherName') ? input.motherName : (typeof savedMother === 'string' ? savedMother : savedMother.name), 100), motherPhone = text(Object.hasOwn(input, 'motherPhone') ? input.motherPhone : (typeof savedMother === 'string' ? '' : savedMother.phone), 20)
      const primaryName = motherName || fatherName, primaryPhone = motherPhone || fatherPhone, primaryRelationship = motherName ? 'Mother / guardian' : 'Father / guardian'
      const updated = {
        ...existing,
        admissionNumber: requestedAdmissionNumber || existing.admissionNumber,
        admissionNumberVerified: requestedAdmissionNumber && requestedAdmissionNumber !== existing.admissionNumber ? true : Boolean(existing.admissionNumberVerified),
        learnerName: text(input.learnerName, 100),
        grade: text(input.grade, 20), residency: learnerResidencies.has(text(input.residency, 20)) ? text(input.residency, 20) : (existing.residency || 'Day scholar'),
        fatherName, fatherPhone, motherName, motherPhone,
        guardianName: primaryName,
        guardianPhone: primaryPhone,
        relationship: primaryRelationship,
        guardians: [{ name: fatherName, phone: fatherPhone, relationship: 'Father' }, { name: motherName, phone: motherPhone, relationship: 'Mother / guardian' }].filter(item => item.name || item.phone),
        medicalNotes: text(input.medicalNotes, 500),
        birthCertificateProvided: Boolean(input.birthCertificateProvided),
        updatedAt: new Date().toISOString()
      }
      if (!updated.learnerName || !updated.grade || !updated.guardianName || !updated.guardianPhone) return send(response, 400, { error: 'Learner name, applying grade, and at least one parent or guardian contact are required.' })
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
