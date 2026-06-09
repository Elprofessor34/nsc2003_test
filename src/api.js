import { supabase } from './supabase'

/* MAPPERS — DB snake_case ↔ JS camelCase */
const fromStudent = (r) => r ? ({
  id: r.id,
  fullName: r.full_name,
  studentClass: r.student_class,
  rollNumber: r.roll_number,
  section: r.section,
  gender: r.gender,
  dob: r.dob,
  enrollmentDate: r.enrollment_date,
  parentName: r.parent_name,
  parentPhone: r.parent_phone,
  motherName: r.mother_name,
  motherPhone: r.mother_phone,
  address: r.address,
  monthlyFee: Number(r.monthly_fee) || 0,
  sessionFee: Number(r.session_fee) || 0,
  notes: r.notes,
  hasPhoto: !!r.has_photo,
  photoUrl: r.photo_url,
  createdAt: r.created_at,
}) : null

const toStudent = (s) => ({
  id: s.id,
  full_name: s.fullName,
  student_class: s.studentClass,
  roll_number: s.rollNumber ? String(s.rollNumber) : null,
  section: s.section || null,
  gender: s.gender || null,
  dob: s.dob || null,
  enrollment_date: s.enrollmentDate || null,
  parent_name: s.parentName || null,
  parent_phone: s.parentPhone || null,
  mother_name: s.motherName || null,
  mother_phone: s.motherPhone || null,
  address: s.address || null,
  monthly_fee: Number(s.monthlyFee) || 0,
  session_fee: Number(s.sessionFee) || 0,
  notes: s.notes || null,
  has_photo: !!s.hasPhoto,
  photo_url: s.photoUrl || null,
})

const fromPayment = (r) => r ? ({
  id: r.id,
  studentId: r.student_id,
  paymentType: r.payment_type,
  month: r.month,
  year: r.year,
  amount: Number(r.amount) || 0,
  method: r.method,
  description: r.description,
  paidDate: r.paid_date,
  recordedBy: r.recorded_by,
  recordedAt: r.recorded_at,
  archived: !!r.archived,
  archivedClass: r.archived_class,
  archivedSection: r.archived_section,
  archivedRoll: r.archived_roll,
  archivedName: r.archived_name,
}) : null

const toPayment = (p) => ({
  id: p.id,
  student_id: p.studentId,
  payment_type: p.paymentType,
  month: p.month || null,
  year: p.year,
  amount: Number(p.amount) || 0,
  method: p.method || null,
  description: p.description || null,
  paid_date: p.paidDate || null,
  recorded_by: p.recordedBy || null,
})

/* STUDENTS */
export const listStudents = async () => {
  const { data, error } = await supabase.from('students').select('*').order('created_at', { ascending: true })
  if (error) throw error
  return (data || []).map(fromStudent)
}

export const addStudent = async (s) => {
  const { data, error } = await supabase.from('students').insert(toStudent(s)).select().single()
  if (error) throw error
  return fromStudent(data)
}

export const updateStudent = async (id, s) => {
  const { data, error } = await supabase.from('students').update(toStudent(s)).eq('id', id).select().single()
  if (error) throw error
  return fromStudent(data)
}

export const updateStudentFields = async (id, fields) => {
  const dbFields = {}
  if ('studentClass' in fields) dbFields.student_class = fields.studentClass
  if ('rollNumber' in fields) dbFields.roll_number = fields.rollNumber ? String(fields.rollNumber) : null
  if ('section' in fields) dbFields.section = fields.section || null
  if ('monthlyFee' in fields) dbFields.monthly_fee = Number(fields.monthlyFee) || 0
  if ('sessionFee' in fields) dbFields.session_fee = Number(fields.sessionFee) || 0
  const { error } = await supabase.from('students').update(dbFields).eq('id', id)
  if (error) throw error
}

export const deleteStudent = async (id) => {
  const { error } = await supabase.from('students').delete().eq('id', id)
  if (error) throw error
}

export const deleteAllStudents = async () => {
  const { error } = await supabase.from('students').delete().neq('id', '__never_match__')
  if (error) throw error
}

/* PAYMENTS */
// Active payments only (archived ones are excluded from the day-to-day views).
export const listPayments = async () => {
  const { data, error } = await supabase.from('payments').select('*').eq('archived', false).order('recorded_at', { ascending: false })
  if (error) throw error
  return (data || []).map(fromPayment)
}

// Archived payments (previous classes/years), used for the archived export section.
export const listArchivedPayments = async () => {
  const { data, error } = await supabase.from('payments').select('*').eq('archived', true).order('year', { ascending: false })
  if (error) throw error
  return (data || []).map(fromPayment)
}

// On promotion: flag a student's active payments as archived, snapshotting the
// class / section / roll / name they had at the time.
export const archiveStudentPayments = async (studentId, snap) => {
  const { error } = await supabase.from('payments').update({
    archived: true,
    archived_class: snap.archivedClass || null,
    archived_section: snap.archivedSection || null,
    archived_roll: snap.archivedRoll ? String(snap.archivedRoll) : null,
    archived_name: snap.archivedName || null,
  }).eq('student_id', studentId).eq('archived', false)
  if (error) throw error
}

export const addPayment = async (p) => {
  const { data, error } = await supabase.from('payments').insert(toPayment(p)).select().single()
  if (error) throw error
  return fromPayment(data)
}

export const deletePayment = async (id) => {
  const { error } = await supabase.from('payments').delete().eq('id', id)
  if (error) throw error
}

export const deleteAllPayments = async () => {
  const { error } = await supabase.from('payments').delete().neq('id', '__never_match__')
  if (error) throw error
}

/* SETTINGS */
export const getSettings = async () => {
  const { data, error } = await supabase.from('settings').select('*').eq('id', 1).maybeSingle()
  if (error || !data) return null
  return {
    schoolName: data.school_name,
    currency: data.currency,
    defaultMonthlyFee: Number(data.default_monthly_fee) || 0,
    defaultSessionFee: Number(data.default_session_fee) || 0,
    academicYear: data.academic_year || new Date().getFullYear() + '',
  }
}

export const updateSettings = async (s) => {
  const { error } = await supabase.from('settings').update({
    school_name: s.schoolName,
    currency: s.currency,
    default_monthly_fee: Number(s.defaultMonthlyFee) || 0,
    default_session_fee: Number(s.defaultSessionFee) || 0,
    academic_year: s.academicYear,
    updated_at: new Date().toISOString(),
  }).eq('id', 1)
  if (error) throw error
}

/* PHOTOS - Supabase Storage */
const PHOTO_BUCKET = 'photos'

export const uploadPhoto = async (studentId, dataUrl) => {
  const res = await fetch(dataUrl)
  const blob = await res.blob()
  const path = `student-${studentId}.jpg`
  const { error } = await supabase.storage.from(PHOTO_BUCKET).upload(path, blob, {
    upsert: true, contentType: 'image/jpeg', cacheControl: '3600',
  })
  if (error) throw error
  const { data } = supabase.storage.from(PHOTO_BUCKET).getPublicUrl(path)
  return `${data.publicUrl}?t=${Date.now()}`
}

export const removePhoto = async (studentId) => {
  await supabase.storage.from(PHOTO_BUCKET).remove([`student-${studentId}.jpg`])
}

/* AUTH */
export const getCurrentSession = async () => {
  const { data } = await supabase.auth.getSession()
  return data.session
}

export const signIn = async (email, password) => {
  const { error } = await supabase.auth.signInWithPassword({ email, password })
  return { ok: !error, msg: error?.message }
}

export const signUp = async (email, password) => {
  const { error } = await supabase.auth.signUp({ email, password })
  return { ok: !error, msg: error?.message }
}

export const signOut = async () => {
  await supabase.auth.signOut()
}

export const resetPassword = async (email) => {
  const { error } = await supabase.auth.resetPasswordForEmail(email, {
    redirectTo: typeof window !== 'undefined' ? window.location.origin : undefined,
  })
  return { ok: !error, msg: error?.message }
}

export const updatePassword = async (newPassword) => {
  const { error } = await supabase.auth.updateUser({ password: newPassword })
  return { ok: !error, msg: error?.message }
}

/* PROFILES & APPROVAL */
// Returns the current user's profile, or a default unapproved profile if none exists yet.
export const getMyProfile = async () => {
  const { data: { user } } = await supabase.auth.getUser()
  if (!user) return null
  const { data, error } = await supabase.from('profiles').select('*').eq('id', user.id).maybeSingle()
  if (error) { console.error('getMyProfile', error); return { id: user.id, email: user.email, approved: false, isAdmin: false } }
  if (!data) return { id: user.id, email: user.email, approved: false, isAdmin: false }
  return { id: data.id, email: data.email, approved: !!data.approved, isAdmin: !!data.is_admin, createdAt: data.created_at }
}

// All profiles (only readable by approved admins, enforced via RLS).
export const listProfiles = async () => {
  const { data, error } = await supabase.from('profiles').select('*').order('created_at', { ascending: true })
  if (error) { console.error('listProfiles', error); return [] }
  return (data || []).map(p => ({ id: p.id, email: p.email, approved: !!p.approved, isAdmin: !!p.is_admin, createdAt: p.created_at }))
}

// Approve or revoke a user's access.
export const setApproval = async (id, approved) => {
  const { error } = await supabase.from('profiles').update({ approved }).eq('id', id)
  if (error) throw error
}
