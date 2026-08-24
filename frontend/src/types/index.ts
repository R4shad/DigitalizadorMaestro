export interface ExcelRosterStudent {
  row_number: number
  student_id?: string
  full_name: string
  normalized_name: string
}

export interface SheetAnalysis {
  sheet_name: string
  header_row: number
  id_column?: string
  name_column: string
  score_columns: string[]
  roster: ExcelRosterStudent[]
}

export interface MatchedStudentGrade {
  row_number: number
  student_id?: string
  full_name: string
  scores: Record<string, number | null>
  confidence: number
  matched_ocr_text?: string
}

export interface ProcessGradesResponse {
  sheets: SheetAnalysis[]
  selected_sheet: string
  result: {
    columns_detected: string[]
    students: MatchedStudentGrade[]
    unmatched_ocr_rows: Array<{
      raw_text: string
      extracted_scores: number[]
    }>
  }
}
