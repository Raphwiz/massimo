async function send(path, options = {}) {
  const portalRole = typeof window === 'undefined' ? '' : window.sessionStorage.getItem('st_dorothy_portal_role') || ''
  const response = await fetch(path, {
    headers: { 'Content-Type': 'application/json', ...(portalRole ? { 'X-Portal-Role': portalRole } : {}), ...options.headers },
    credentials: 'include',
    ...options
  })
  const payload = await response.json().catch(() => ({}))
  if (response.status === 401 && typeof window !== 'undefined' && !String(path).endsWith('/api/auth/logout')) window.dispatchEvent(new Event('st-dorothy-session-expired'))
  if (!response.ok) throw new Error(payload.error || 'The school portal server is unavailable.')
  return payload
}

async function request(path, options = {}) {
  const isLocalBrowser = typeof window !== 'undefined' && ['localhost', '127.0.0.1'].includes(window.location.hostname)
  if (isLocalBrowser && path.startsWith('/api/')) return send(`http://127.0.0.1:8787${path}`, options)
  try {
    return await send(path, options)
  } catch (error) {
    if (isLocalBrowser && path.startsWith('/api/')) return send(`http://127.0.0.1:8787${path}`, options)
    throw error
  }
}

export const getCalendarEvents = () => request('/api/calendar-events')
export const loginPortal = (credentials) => request('/api/auth/login', { method: 'POST', body: JSON.stringify(credentials) })
export const getPortalSession = () => request('/api/auth/session')
export const logoutPortal = () => request('/api/auth/logout', { method: 'POST' })
export const changePortalPassword = (passwords) => request('/api/auth/change-password', { method: 'POST', body: JSON.stringify(passwords) })
export const activatePortalAccount = (details) => request('/api/auth/activate', { method: 'POST', body: JSON.stringify(details) })
export const getPortalAccounts = () => request('/api/accounts')
export const createPortalAccount = (details) => request('/api/accounts', { method: 'POST', body: JSON.stringify(details) })
export const createParentPortal = (admissionId) => request('/api/parent-portal', { method: 'POST', body: JSON.stringify({ admissionId }) })
export const resetPortalAccount = (id) => request(`/api/accounts/${encodeURIComponent(id)}/reset`, { method: 'POST', body: JSON.stringify({}) })
export const setPortalAccountStatus = (id, active) => request(`/api/accounts/${encodeURIComponent(id)}/status`, { method: 'PUT', body: JSON.stringify({ active }) })
export const updatePortalAccountRole = (id, role, action = 'grant') => request(`/api/accounts/${encodeURIComponent(id)}/roles`, { method: 'PUT', body: JSON.stringify({ role, action }) })
export const deletePortalAccount = (id) => request(`/api/accounts/${encodeURIComponent(id)}`, { method: 'DELETE' })
export const getSystemAdminOverview = () => request('/api/system-admin/overview')
export const getSystemLogs = () => request('/api/system-admin/logs')
export const getHeadteacherDashboard = () => request('/api/headteacher/dashboard')
export const getHeadteacherAttendance = (date) => request(`/api/headteacher/attendance?date=${encodeURIComponent(date || '')}`)
export const getGoLivePreview = () => request('/api/system-admin/go-live')
export const runGoLiveReset = (details) => request('/api/system-admin/go-live', { method: 'POST', body: JSON.stringify(details) })
export const createCalendarEvent = (event) => request('/api/calendar-events', { method: 'POST', body: JSON.stringify(event) })
export const getSmsMessages = () => request('/api/sms')
export const createSmsMessage = (message) => request('/api/sms', { method: 'POST', body: JSON.stringify(message) })
export const getSmsWallet = () => request('/api/sms/wallet')
export const createSmsTopup = (topup) => request('/api/sms/topups', { method: 'POST', body: JSON.stringify(topup) })
export const getAdmissions = () => request('/api/admissions')
export const previewLearnerImport = (fileBase64) => request('/api/learner-import/preview', { method: 'POST', body: JSON.stringify({ fileBase64 }) })
export const commitLearnerImport = (rows) => request('/api/learner-import/commit', { method: 'POST', body: JSON.stringify({ rows }) })
export const createAdmission = (application) => request('/api/admissions', { method: 'POST', body: JSON.stringify(application) })
export const updateAdmission = (id, application) => request(`/api/admissions/${encodeURIComponent(id)}`, { method: 'PUT', body: JSON.stringify(application) })
export const getDocuments = () => request('/api/documents')
export const createDocument = (document) => request('/api/documents', { method: 'POST', body: JSON.stringify(document) })
export const getNotices = () => request('/api/notices')
export const getWebsiteNotices = () => request('/api/public/notices')
export const getWebsiteSettings = () => request('/api/public/website-settings')
export const getAdminWebsiteSettings = () => request('/api/system-admin/website-settings')
export const saveAdminWebsiteSettings = (settings) => request('/api/system-admin/website-settings', { method:'PUT', body:JSON.stringify(settings) })
export const createNotice = (notice) => request('/api/notices', { method: 'POST', body: JSON.stringify(notice) })
export const updateNotice = (id, notice) => request(`/api/notices/${encodeURIComponent(id)}`, { method: 'PUT', body: JSON.stringify(notice) })
export const getInternalMessages = () => request('/api/internal-messages')
export const createInternalMessage = (message) => request('/api/internal-messages', { method: 'POST', body: JSON.stringify(message) })
export const getClassPlacements = () => request('/api/class-placements')
export const saveClassPlacement = (placement) => request('/api/class-placements', { method: 'POST', body: JSON.stringify(placement) })
export const getPlacementSuggestions = (academicYear) => request(`/api/class-placements/suggestions?academicYear=${encodeURIComponent(academicYear || '')}`)
export const autoPlaceLearners = (details) => request('/api/class-placements/auto-place', { method: 'POST', body: JSON.stringify(details) })
export const getSchoolClasses = () => request('/api/school-classes')
export const createSchoolClass = (schoolClass) => request('/api/school-classes', { method: 'POST', body: JSON.stringify(schoolClass) })
export const updateSchoolClass = (id, schoolClass) => request(`/api/school-classes/${encodeURIComponent(id)}`, { method: 'PUT', body: JSON.stringify(schoolClass) })
export const deleteSchoolClass = (id) => request(`/api/school-classes/${encodeURIComponent(id)}`, { method: 'DELETE' })
export const getStaff = () => request('/api/staff')
export const createStaff = (staff) => request('/api/staff', { method: 'POST', body: JSON.stringify(staff) })
export const getAcademics = () => request('/api/academics')
export const createSubject = (subject) => request('/api/subjects', { method: 'POST', body: JSON.stringify(subject) })
export const createTeacherAssignment = (assignment) => request('/api/teacher-assignments', { method: 'POST', body: JSON.stringify(assignment) })
export const createTimetableEntry = (entry) => request('/api/timetable', { method: 'POST', body: JSON.stringify(entry) })
export const getAttendance = ({ className, date } = {}) => request(`/api/attendance?className=${encodeURIComponent(className || '')}&date=${encodeURIComponent(date || '')}`)
export const saveAttendance = (attendance) => request('/api/attendance', { method: 'POST', body: JSON.stringify(attendance) })
export const attendanceWorkbookUrl = ({ className, date } = {}) => {
  const localApi = typeof window !== 'undefined' && ['localhost', '127.0.0.1'].includes(window.location.hostname) ? 'http://127.0.0.1:8787' : ''
  return `${localApi}/api/attendance/workbook?className=${encodeURIComponent(className || '')}&date=${encodeURIComponent(date || '')}`
}
export const classListWorkbookUrl = (classId) => {
  const localApi = typeof window !== 'undefined' && ['localhost', '127.0.0.1'].includes(window.location.hostname) ? 'http://127.0.0.1:8787' : ''
  return `${localApi}/api/class-lists/workbook?classId=${encodeURIComponent(classId || '')}`
}
export const classListPdfUrl = (classId) => {
  const localApi = typeof window !== 'undefined' && ['localhost', '127.0.0.1'].includes(window.location.hostname) ? 'http://127.0.0.1:8787' : ''
  return `${localApi}/api/class-lists/pdf?classId=${encodeURIComponent(classId || '')}`
}
export const boarderRegisterPdfUrl = (grade = '') => {
  const localApi = typeof window !== 'undefined' && ['localhost', '127.0.0.1'].includes(window.location.hostname) ? 'http://127.0.0.1:8787' : ''
  return `${localApi}/api/secretary/boarders/pdf?grade=${encodeURIComponent(grade)}`
}
export const staffListWorkbookUrl = () => {
  const localApi = typeof window !== 'undefined' && ['localhost', '127.0.0.1'].includes(window.location.hostname) ? 'http://127.0.0.1:8787' : ''
  return `${localApi}/api/staff-list/workbook`
}
export const getLessonPlans = () => request('/api/lesson-plans')
export const createLessonPlan = (lessonPlan) => request('/api/lesson-plans', { method: 'POST', body: JSON.stringify(lessonPlan) })
export const updateLessonPlan = (id, lessonPlan) => request(`/api/lesson-plans/${encodeURIComponent(id)}`, { method: 'PUT', body: JSON.stringify(lessonPlan) })
export const createAssessment = (assessment) => request('/api/assessments', { method: 'POST', body: JSON.stringify(assessment) })
export const getAssessmentMarks = (assessmentId) => request(`/api/assessment-marks?assessmentId=${encodeURIComponent(assessmentId || '')}`)
export const saveAssessmentMarks = (marks) => request('/api/assessment-marks', { method: 'POST', body: JSON.stringify(marks) })
export const getFeeStructures = () => request('/api/fee-structures')
export const createFeeStructure = (feeStructure) => request('/api/fee-structures', { method: 'POST', body: JSON.stringify(feeStructure) })
export const getLeaveRequests = () => request('/api/leave-requests')
export const createLeaveRequest = (leaveRequest) => request('/api/leave-requests', { method: 'POST', body: JSON.stringify(leaveRequest) })
export const getStaffAttendance = () => request('/api/staff-attendance')
export const saveStaffAttendance = (attendance) => request('/api/staff-attendance', { method: 'POST', body: JSON.stringify(attendance) })
export const getPromotions = () => request('/api/promotions')
export const createPromotion = (promotion) => request('/api/promotions', { method: 'POST', body: JSON.stringify(promotion) })
export const getParentLearners = (phone) => request(`/api/parent/learners?phone=${encodeURIComponent(phone)}`)
export const createParentPaymentReport = (report) => request('/api/parent/payment-reports', { method: 'POST', body: JSON.stringify(report) })
export const getParentMessages = (phone = '') => request(`/api/parent/messages?phone=${encodeURIComponent(phone)}`)
export const createParentMessage = (message) => request('/api/parent/messages', { method: 'POST', body: JSON.stringify(message) })
export const getFinanceLearners = () => request('/api/finance/learners')
export const getSecretaryFeeLookup = (query) => request(`/api/secretary/fee-lookup?query=${encodeURIComponent(query || '')}`)
export const getSecretaryLearnerRecords = () => request('/api/secretary/learner-records')
export const updateLearnerResidencyBatch = (details) => request('/api/secretary/learner-records/residency', { method: 'POST', body: JSON.stringify(details) })
export const learnerRecordPdfUrl = (admissionId) => {
  const localApi = typeof window !== 'undefined' && ['localhost', '127.0.0.1'].includes(window.location.hostname) ? 'http://127.0.0.1:8787' : ''
  return `${localApi}/api/secretary/learner-records/${encodeURIComponent(admissionId || '')}/pdf`
}
export const getDefaulters = () => request('/api/finance/defaulters')
export const getBudgetSummary = () => request('/api/finance/budget-summary')
export const getFinanceAudit = () => request('/api/finance/audit')
export const getFeeAdjustments = () => request('/api/finance/adjustments')
export const createFeeAdjustment = (adjustment) => request('/api/finance/adjustments', { method:'POST', body:JSON.stringify(adjustment) })
export const getFinanceApprovalRequests = () => request('/api/finance/approval-requests')
export const decideFinanceApproval = (id, decision, decisionNote = '') => request(`/api/finance/approval-requests/${encodeURIComponent(id)}`, { method:'PUT', body:JSON.stringify({ decision, decisionNote }) })
export const getCashbook = () => request('/api/finance/cashbook')
export const createCashbookEntry = (entry) => request('/api/finance/cashbook', { method:'POST', body:JSON.stringify(entry) })
export const getBankDeposits = () => request('/api/finance/bank-deposits')
export const createBankDeposit = (deposit) => request('/api/finance/bank-deposits', { method:'POST', body:JSON.stringify(deposit) })
export const financeOperationsWorkbookUrl = (className='', report='all') => `${typeof window!=='undefined'&&['localhost','127.0.0.1'].includes(window.location.hostname)?'http://127.0.0.1:8787':''}/api/finance/operations-workbook?className=${encodeURIComponent(className)}&report=${encodeURIComponent(report)}`
export const defaulterListPdfUrl = (className='') => `${typeof window!=='undefined'&&['localhost','127.0.0.1'].includes(window.location.hostname)?'http://127.0.0.1:8787':''}/api/finance/defaulters/pdf?className=${encodeURIComponent(className)}`
export const expensesPdfUrl = () => `${typeof window!=='undefined'&&['localhost','127.0.0.1'].includes(window.location.hostname)?'http://127.0.0.1:8787':''}/api/finance/expenses/pdf`
export const expensesWorkbookUrl = () => `${typeof window!=='undefined'&&['localhost','127.0.0.1'].includes(window.location.hostname)?'http://127.0.0.1:8787':''}/api/finance/expenses/workbook`
export const getFeeStatement = (admissionId) => request(`/api/finance/statements/${encodeURIComponent(admissionId)}`)
export const feeStatementsWorkbookUrl = (outstandingOnly = true) => {
  const localApi = typeof window !== 'undefined' && ['localhost', '127.0.0.1'].includes(window.location.hostname) ? 'http://127.0.0.1:8787' : ''
  return `${localApi}/api/finance/statements-workbook?outstanding=${outstandingOnly}`
}
const financeDownloadBase = () => typeof window !== 'undefined' && ['localhost', '127.0.0.1'].includes(window.location.hostname) ? 'http://127.0.0.1:8787' : ''
export const feeStatementPdfUrl = (admissionId) => `${financeDownloadBase()}/api/finance/statements/${encodeURIComponent(admissionId)}/pdf`
export const feeStatementsBulkUrl = (className = '') => `${financeDownloadBase()}/api/finance/statements-bulk?className=${encodeURIComponent(className)}`
export const getReceipts = () => request('/api/finance/receipts')
export const createReceipt = (receipt) => request('/api/finance/receipts', { method: 'POST', body: JSON.stringify(receipt) })
export const receiptPdfUrl = (receiptId) => `${financeDownloadBase()}/api/finance/receipts/${encodeURIComponent(receiptId || '')}/pdf`
export const getMpesaTransactions = () => request('/api/finance/mpesa-transactions')
export const createMpesaTransaction = (transaction) => request('/api/finance/mpesa-transactions', { method: 'POST', body: JSON.stringify(transaction) })
export const reconcileMpesaTransaction = (id, payment) => request(`/api/finance/mpesa-transactions/${encodeURIComponent(id)}`, { method: 'PUT', body: JSON.stringify(payment) })
export const getFinanceReport = ({ start, end } = {}) => request(`/api/finance/reports?start=${encodeURIComponent(start || '')}&end=${encodeURIComponent(end || '')}`)
export const createAdmissionEnquiry = (enquiry) => request('/api/admission-enquiries', { method: 'POST', body: JSON.stringify(enquiry) })
