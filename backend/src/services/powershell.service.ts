import { exec } from 'node:child_process'
import fs from 'node:fs/promises'
import path from 'node:path'
import os from 'node:os'

export interface ColumnTargetMapping {
  ocrColumn: string
  excelColumnLetter: string
  scoreType: 'SER' | 'SABER' | 'HACER' | 'AUTOEVALUACION' | 'CUSTOM'
}

export interface StudentRowUpdate {
  rowNumber: number
  scores: Record<string, number | null>
}

export class PowerShellService {
  async injectViaCom(
    originalFilePath: string,
    sheetName: string,
    columnMappings: ColumnTargetMapping[],
    students: StudentRowUpdate[],
  ): Promise<void> {
    const tempDir = os.tmpdir()
    const scriptPath = path.join(tempDir, `inject_grades_${Date.now()}.ps1`)

    const updatesScript = students
      .map((student) => {
        const statements: string[] = []
        for (const mapping of columnMappings) {
          const val = student.scores[mapping.ocrColumn]
          if (val !== null && val !== undefined) {
            statements.push(
              `$worksheet.Range("${mapping.excelColumnLetter}${student.rowNumber}").Value2 = ${val}`,
            )
          }
        }
        return statements.join('\n')
      })
      .filter((s) => s.length > 0)
      .join('\n')

    const psContent = `
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $workbook = $excel.Workbooks.Open("${originalFilePath.replace(/\\/g, '\\\\')}")
    $worksheet = $workbook.Sheets.Item("${sheetName}")

    ${updatesScript}

    $workbook.Save()
    $workbook.Close()
}
finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
`

    await fs.writeFile(scriptPath, psContent, 'utf-8')

    return new Promise((resolve, reject) => {
      exec(
        `powershell.exe -ExecutionPolicy Bypass -File "${scriptPath}"`,
        async (error, stdout, stderr) => {
          await fs.unlink(scriptPath).catch(() => {})
          if (error) {
            reject(
              new Error(
                `PowerShell Execution Error: ${stderr || error.message}`,
              ),
            )
            return
          }
          resolve()
        },
      )
    })
  }
}
