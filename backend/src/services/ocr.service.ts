import sharp from 'sharp'
import FormData from 'form-data'
import axios from 'axios'
import { ExcelRosterStudent } from './excel.service.ts'

interface RawOcrItem {
  text: string
  confidence: number
  center_x: number
  center_y: number
}

interface ProcessResult {
  detected_period: string
  detected_grade: string
  detected_subject: string
  dimensions: {
    ser_columns: string[]
    saber_columns: string[]
    hacer_columns: string[]
    autoevaluacion_columns: string[]
  }
  criteria: Record<string, unknown>
  students: Array<{
    list_number: number
    raw_name: string
    matched_name: string
    student_code: string
    scores: Record<string, { score: number | null; is_uncertain: boolean }>
    needs_review: boolean
    review_reason: string
  }>
}

export class OcrService {
  private async cropRows(
    imageBuffer: Buffer,
    startRow: number,
    totalRows: number,
    rowsInCrop: number,
  ): Promise<Buffer> {
    const source = sharp(imageBuffer).rotate()
    const metadata = await source.metadata()
    if (!metadata.width || !metadata.height) {
      throw new Error('Image dimensions could not be determined')
    }

    const tableTop = Number(process.env.OCR_TABLE_TOP || 0.26)
    const tableBottom = Number(process.env.OCR_TABLE_BOTTOM || 0.93)
    const tableHeight = Math.floor(metadata.height * (tableBottom - tableTop))
    const rowHeight = tableHeight / totalRows
    const top = Math.max(
      0,
      Math.floor(metadata.height * tableTop + (startRow - 1) * rowHeight),
    )
    const bottom = Math.min(
      metadata.height,
      Math.ceil(
        metadata.height * tableTop + (startRow - 1 + rowsInCrop) * rowHeight,
      ),
    )

    return source
      .extract({ left: 0, top, width: metadata.width, height: bottom - top })
      .jpeg({ quality: 100 })
      .toBuffer()
  }

  private groupIntoRows(
    ocrData: RawOcrItem[],
    maxRowDiff: number = 30,
  ): RawOcrItem[][] {
    if (!ocrData || ocrData.length === 0) return []

    const sortedByY = [...ocrData].sort((a, b) => a.center_y - b.center_y)
    const rows: RawOcrItem[][] = []
    let currentRow: RawOcrItem[] = [sortedByY[0]]

    for (let i = 1; i < sortedByY.length; i++) {
      const item = sortedByY[i]
      const lastItem = currentRow[currentRow.length - 1]

      if (Math.abs(item.center_y - lastItem.center_y) <= maxRowDiff) {
        currentRow.push(item)
      } else {
        rows.push(currentRow)
        currentRow = [item]
      }
    }
    if (currentRow.length > 0) rows.push(currentRow)

    rows.forEach((row) => row.sort((a, b) => a.center_x - b.center_x))
    return rows
  }

  async processGradeImages(
    imageBuffers: Buffer[],
    officialRoster: ExcelRosterStudent[],
  ): Promise<ProcessResult> {
    const validRoster = officialRoster.filter(
      (s) =>
        s.full_name &&
        s.full_name !== '0' &&
        !s.full_name.includes('[object Object]'),
    )

    const detectedScoreMap = new Map<
      number,
      { raw_scores: (number | null)[] }
    >()

    for (let imageIndex = 0; imageIndex < imageBuffers.length; imageIndex++) {
      const rawBuf = imageBuffers[imageIndex]
      if (!rawBuf) continue

      const firstListNumber = imageIndex * 20 + 1
      const lastListNumber = imageIndex === 0 ? 20 : validRoster.length
      const rowRanges = Array.from(
        { length: Math.ceil((lastListNumber - firstListNumber + 1) / 5) },
        (_, rangeIndex) => [
          firstListNumber + rangeIndex * 5,
          Math.min(firstListNumber + rangeIndex * 5 + 4, lastListNumber),
        ],
      )

      for (const [rangeIndex, [rangeStart, rangeEnd]] of rowRanges.entries()) {
        try {
          const croppedImage = await this.cropRows(
            rawBuf,
            rangeIndex * 5 + 1,
            lastListNumber - firstListNumber + 1,
            rangeEnd - rangeStart + 1,
          )

          const formData = new FormData()
          formData.append('file', croppedImage, {
            filename: 'crop.jpg',
            contentType: 'image/jpeg',
          })

          const response = await axios.post<{ data: RawOcrItem[] }>(
            'http://127.0.0.1:8000/extract',
            formData,
            {
              headers: formData.getHeaders(),
            },
          )

          const rows = this.groupIntoRows(response.data.data)

          let currentListNum = rangeStart
          for (const row of rows) {
            if (currentListNum > rangeEnd) break

            const rawScores = row.map((item) => {
              const num = Number(item.text)
              return isNaN(num) ? null : num
            })

            detectedScoreMap.set(currentListNum, { raw_scores: rawScores })
            currentListNum++
          }
        } catch {
          // Keep loop running on crop/request failure
        }
      }
    }

    const finalStudents = validRoster.map((rosterStudent) => {
      const detected = detectedScoreMap.get(rosterStudent.list_number) || {
        raw_scores: [],
      }
      const scores: Record<
        string,
        { score: number | null; is_uncertain: boolean }
      > = {}

      detected.raw_scores.forEach((val, idx) => {
        scores[`__RAW__${idx}`] = { score: val, is_uncertain: false }
      })

      return {
        list_number: rosterStudent.list_number,
        raw_name: '',
        matched_name: rosterStudent.full_name,
        student_code: rosterStudent.student_code,
        scores,
        needs_review: true,
        review_reason: 'Pending parameterization',
      }
    })

    return {
      detected_period: 'Registro 2° Trimestre',
      detected_grade: 'CUARTO-A',
      detected_subject: 'CIENCIAS NATURALES BIOLOGÍA - GEOGRAFÍA',
      dimensions: {
        ser_columns: [],
        saber_columns: [],
        hacer_columns: [],
        autoevaluacion_columns: [],
      },
      criteria: {},
      students: finalStudents,
    }
  }
}
