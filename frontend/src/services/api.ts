import axios from 'axios'
import type { ProcessGradesResponse, SheetAnalysis } from '../types/index.ts'

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

export const exportExcel = async (
  templateFile: File,
  sheetName: string,
  columnMappings: Record<string, string>,
  students: Array<{
    row_number: number
    scores: Record<string, number | null>
  }>,
): Promise<Blob> => {
  const formData = new FormData()
  formData.append('template', templateFile)
  formData.append('sheet_name', sheetName)
  formData.append('column_mappings', JSON.stringify(columnMappings))
  formData.append('students', JSON.stringify(students))

  const res = await apiClient.post('/export-excel', formData, {
    responseType: 'blob',
  })
  return res.data
}
