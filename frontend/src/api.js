async function send(path, options = {}) {
  const response = await fetch(path, {
    headers: { 'Content-Type': 'application/json', ...options.headers },
    credentials: 'include',
    ...options
  })
  const payload = await response.json().catch(() => ({}))
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
export const logoutPortal = () => request('/api/auth/logout', { method: 'POST' })
export const changePortalPassword = (passwords) => request('/api/auth/change-password', { method: 'POST', body: JSON.stringify(passwords) })
export const activatePortalAccount = (details) => request('/api/auth/activate', { method: 'POST', body: JSON.stringify(details) })
export const getPortalAccounts = () => request('/api/accounts')
export const createPortalAccount = (details) => request('/api/accounts', { method: 'POST', body: JSON.stringify(details) })
export const resetPortalAccount = (id) => request(`/api/accounts/${encodeURIComponent(id)}/reset`, { method: 'POST', body: JSON.stringify({}) })
export const setPortalAccountStatus = (id, active) => request(`/api/accounts/${encodeURIComponent(id)}/status`, { method: 'PUT', body: JSON.stringify({ active }) })
export const getSystemAdminOverview = () => request('/api/system-admin/overview')
export const createCalendarEvent = (event) => request('/api/calendar-events', { method: 'POST', body: JSON.stringify(event) })
export const getSmsMessages = () => request('/api/sms')
export const createSmsMessage = (message) => request('/api/sms', { method: 'POST', body: JSON.stringify(message) })
export const getAdmissions = () => request('/api/admissions')
export const createAdmission = (application) => request('/api/admissions', { method: 'POST', body: JSON.stringify(application) })
export const updateAdmission = (id, application) => request(`/api/admissions/${encodeURIComponent(id)}`, { method: 'PUT', body: JSON.stringify(application) })
export const getDocuments = () => request('/api/documents')
export const createDocument = (document) => request('/api/documents', { method: 'POST', body: JSON.stringify(document) })
export const getNotices = () => request('/api/notices')
export const getWebsiteNotices = () => request('/api/public/notices')
export const createNotice = (notice) => request('/api/notices', { method: 'POST', body: JSON.stringify(notice) })
export const updateNotice = (id, notice) => request(`/api/notices/${encodeURIComponent(id)}`, { method: 'PUT', body: JSON.stringify(notice) })
export const getInternalMessages = () => request('/api/internal-messages')
export const createInternalMessage = (message) => request('/api/internal-messages', { method: 'POST', body: JSON.stringify(message) })
export const getClassPlacements = () => request('/api/class-placements')
export const saveClassPlacement = (placement) => request('/api/class-placements', { method: 'POST', body: JSON.stringify(placement) })
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
export const getReceipts = () => request('/api/finance/receipts')
export const createReceipt = (receipt) => request('/api/finance/receipts', { method: 'POST', body: JSON.stringify(receipt) })
export const getMpesaTransactions = () => request('/api/finance/mpesa-transactions')
export const createMpesaTransaction = (transaction) => request('/api/finance/mpesa-transactions', { method: 'POST', body: JSON.stringify(transaction) })
export const reconcileMpesaTransaction = (id, payment) => request(`/api/finance/mpesa-transactions/${encodeURIComponent(id)}`, { method: 'PUT', body: JSON.stringify(payment) })
export const getFinanceReport = ({ start, end } = {}) => request(`/api/finance/reports?start=${encodeURIComponent(start || '')}&end=${encodeURIComponent(end || '')}`)
export const createAdmissionEnquiry = (enquiry) => request('/api/admission-enquiries', { method: 'POST', body: JSON.stringify(enquiry) })
