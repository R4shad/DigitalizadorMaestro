import express, { Request, Response } from 'express'
import cors from 'cors'
import multer from 'multer'
import dotenv from 'dotenv'
import fs from 'node:fs/promises'
import path from 'node:path'
import os from 'node:os'
import { ExcelService, SheetAnalysis } from './services/excel.service.ts'
import { OcrService } from './services/ocr.service.ts'
import {
  PowerShellService,
  ColumnTargetMapping,
  StudentRowUpdate,
} from './services/powershell.service.ts'

dotenv.config()

const app = express()
const port = process.env.PORT || 3001

app.use(cors())
app.use(express.json())

const upload = multer({ storage: multer.memoryStorage() })
const excelService = new ExcelService()
const ocrService = new OcrService()
const powershellService = new PowerShellService()

app.post(
  '/api/inspect-excel',
  upload.single('file'),
  async (req: Request, res: Response): Promise<void> => {
    try {
      if (!req.file) {
        res.status(400).json({ error: 'Excel file is required' })
        return
      }
      const sheets: SheetAnalysis[] = await excelService.inspectWorkbook(
        req.file.buffer,
      )
      res.json({ sheets })
    } catch (err: unknown) {
      const message =
        err instanceof Error ? err.message : 'Internal Server Error'
      res.status(500).json({ error: message })
    }
  },
)

app.post(
  '/api/process-grades',
  upload.fields([
    { name: 'excel', maxCount: 1 },
    { name: 'images', maxCount: 10 },
  ]),
  async (req: Request, res: Response): Promise<void> => {
    try {
      const files = req.files as
        | { [fieldname: string]: Express.Multer.File[] }
        | undefined
      const excelFile = files?.excel?.[0]
      const imageFiles = files?.images || []

      if (!excelFile || imageFiles.length === 0) {
        res
          .status(400)
          .json({ error: 'Excel file and at least one image are required' })
        return
      }

      const sheets: SheetAnalysis[] = await excelService.inspectWorkbook(
        excelFile.buffer,
      )
      const selectedSheetName = req.body.sheet_name || sheets[0]?.sheet_name
      const currentSheet =
        sheets.find((s: SheetAnalysis) => s.sheet_name === selectedSheetName) ||
        sheets[0]

      if (!currentSheet) {
        res.status(400).json({ error: 'No valid sheet found in workbook' })
        return
      }

      const imageBuffers = imageFiles.map((f: Express.Multer.File) => f.buffer)
      const ocrResult = await ocrService.processGradeImages(
        imageBuffers,
        currentSheet.roster,
      )

      res.json({
        sheets,
        selected_sheet: currentSheet.sheet_name,
        result: ocrResult,
      })
    } catch (err: unknown) {
      const message =
        err instanceof Error ? err.message : 'Internal Server Error'
      res.status(500).json({ error: message })
    }
  },
)

app.post(
  '/api/inject-direct-excel',
  upload.single('template'),
  async (req: Request, res: Response): Promise<void> => {
    try {
      if (!req.file) {
        res.status(400).json({ error: 'Template file is required' })
        return
      }

      const { sheet_name, column_mappings, students } = req.body
      const parsedMappings: ColumnTargetMapping[] =
        typeof column_mappings === 'string'
          ? JSON.parse(column_mappings)
          : column_mappings
      const parsedStudents: StudentRowUpdate[] =
        typeof students === 'string' ? JSON.parse(students) : students

      const tempFilePath = path.join(
        os.tmpdir(),
        `temp_template_${Date.now()}.xlsx`,
      )
      await fs.writeFile(tempFilePath, req.file.buffer)

      await powershellService.injectViaCom(
        tempFilePath,
        sheet_name,
        parsedMappings,
        parsedStudents,
      )

      const updatedBuffer = await fs.readFile(tempFilePath)
      await fs.unlink(tempFilePath).catch(() => {})

      res.setHeader(
        'Content-Type',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      )
      res.setHeader(
        'Content-Disposition',
        `attachment; filename=calificaciones_actualizadas_${Date.now()}.xlsx`,
      )
      res.send(updatedBuffer)
    } catch (err: unknown) {
      const message =
        err instanceof Error ? err.message : 'Internal Server Error'
      res.status(500).json({ error: message })
    }
  },
)

app.listen(port, () => {
  console.log(`Backend server running on port ${port}`)
})
