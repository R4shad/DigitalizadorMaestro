import ExcelJS from 'exceljs'

export interface ExcelRosterStudent {
  list_number: number
  student_code: string
  full_name: string
  row_number: number
}

export interface SheetAnalysis {
  sheet_name: string
  available_columns: Array<{
    column_letter: string
    header_name: string
    dimension?: string
  }>
  roster: ExcelRosterStudent[]
}

export class ExcelService {
  async inspectWorkbook(buffer: Buffer): Promise<SheetAnalysis[]> {
    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(buffer as any)

    const results: SheetAnalysis[] = []

    workbook.eachSheet((worksheet) => {
      const roster: ExcelRosterStudent[] = []
      const available_columns: SheetAnalysis['available_columns'] = []

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber > 5) {
          const listNumberCell = row.getCell(1).value
          const codeCell = row.getCell(2).value
          const nameCell = row.getCell(3).value

          if (listNumberCell && nameCell) {
            const listNum = Number(listNumberCell)
            if (!isNaN(listNum)) {
              roster.push({
                list_number: listNum,
                student_code: String(codeCell ?? ''),
                full_name: String(nameCell),
                row_number: rowNumber,
              })
            }
          }
        }
      })

      const headerRow = worksheet.getRow(4)
      headerRow.eachCell((cell, colNumber) => {
        const headerText = String(cell.value ?? '').trim()
        if (headerText && colNumber > 3) {
          const colLetter = worksheet.getColumn(colNumber).letter
          let dimension: string | undefined

          const upper = headerText.toUpperCase()
          if (upper.includes('SER')) dimension = 'SER'
          else if (upper.includes('SABER')) dimension = 'SABER'
          else if (upper.includes('HACER')) dimension = 'HACER'
          else if (upper.includes('AUTO')) dimension = 'AUTOEVALUACION'

          available_columns.push({
            column_letter: colLetter,
            header_name: headerText,
            dimension,
          })
        }
      })

      results.push({
        sheet_name: worksheet.name,
        available_columns,
        roster,
      })
    })

    return results
  }

  async injectScores(
    templateBuffer: Buffer,
    sheetName: string,
    columnMappings: Record<string, string>,
    students: Array<{
      list_number: number
      scores: Record<string, { score: number | null }>
    }>,
  ): Promise<Buffer> {
    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(templateBuffer as any)

    const worksheet = workbook.getWorksheet(sheetName)
    if (!worksheet) {
      throw new Error(`Sheet ${sheetName} not found`)
    }

    const rowMap = new Map<number, number>()
    worksheet.eachRow((row, rowNumber) => {
      const listNum = Number(row.getCell(1).value)
      if (!isNaN(listNum)) {
        rowMap.set(listNum, rowNumber)
      }
    })

    for (const student of students) {
      const targetRowNumber = rowMap.get(student.list_number)
      if (!targetRowNumber) continue

      const targetRow = worksheet.getRow(targetRowNumber)

      for (const [colName, targetLetter] of Object.entries(columnMappings)) {
        if (!targetLetter) continue
        const scoreData = student.scores[colName]
        if (
          scoreData &&
          scoreData.score !== null &&
          scoreData.score !== undefined
        ) {
          targetRow.getCell(targetLetter).value = scoreData.score
        }
      }

      targetRow.commit()
    }

    const outputBuffer = await workbook.xlsx.writeBuffer()
    return Buffer.from(outputBuffer)
  }
}
