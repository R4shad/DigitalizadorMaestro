import axios from 'axios'
import type {
  ProcessGradesResponse,
  SheetAnalysis,
  ColumnMappingConfig,
  MatchedStudentGrade,
} from '../types/index.ts'

const API_BASE_URL = 'http://localhost:3001/api'

const apiClient = axios.create({
  baseURL: API_BASE_URL,
})

export const inspectExcel = async (
  file: File,
): Promise<{ sheets: SheetAnalysis[] }> => {
  const formData = new FormData()
  formData.append('file', file)
  const res = await apiClient.post<{ sheets: SheetAnalysis[] }>(
    '/inspect-excel',
    formData,
  )
  return res.data
}

export const processGrades = async (
  excel: File,
  images: File[],
  sheetName?: string,
): Promise<ProcessGradesResponse> => {
  const formData = new FormData()
  formData.append('excel', excel)
  images.forEach((img) => formData.append('images', img))
  if (sheetName) {
    formData.append('sheet_name', sheetName)
  }

  const res = await apiClient.post<ProcessGradesResponse>(
    '/process-grades',
    formData,
  )
  return res.data
}

export const injectDirectExcel = async (
  templateFile: File,
  sheetName: string,
  mappings: ColumnMappingConfig[],
  students: MatchedStudentGrade[],
): Promise<Blob> => {
  const formData = new FormData()
  formData.append('template', templateFile)
  formData.append('sheet_name', sheetName)

  const payloadMappings = mappings.map((m) => ({
    ocrColumn: m.ocrColumn,
    excelColumnLetter: m.excelColumn,
    scoreType: m.dimension,
  }))

  const payloadStudents = students.map((s) => ({
    rowNumber: s.row_number,
    scores: s.scores,
  }))

  formData.append('column_mappings', JSON.stringify(payloadMappings))
  formData.append('students', JSON.stringify(payloadStudents))

  const res = await apiClient.post('/inject-direct-excel', formData, {
    responseType: 'blob',
  })
  return res.data
}
