import { useState } from 'react'
import {
  FileSpreadsheet,
  Image as ImageIcon,
  ArrowRight,
  Download,
  Loader2,
  AlertCircle,
  CheckCircle2,
  RefreshCw,
} from 'lucide-react'
import { processGrades, exportExcel } from './services/api.ts'
import type {
  ProcessGradesResponse,
  MatchedStudentGrade,
} from './types/index.ts'

export default function App() {
  const [excelFile, setExcelFile] = useState<File | null>(null)
  const [images, setImages] = useState<File[]>([])
  const [isLoading, setIsLoading] = useState(false)
  const [errorMessage, setErrorMessage] = useState<string | null>(null)
  const [processedData, setProcessedData] =
    useState<ProcessGradesResponse | null>(null)
  const [studentsData, setStudentsData] = useState<MatchedStudentGrade[]>([])
  const [selectedSheet, setSelectedSheet] = useState<string>('')

  const handleExcelUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setExcelFile(e.target.files[0])
      setProcessedData(null)
      setErrorMessage(null)
    }
  }

  const handleImagesUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      setImages(Array.from(e.target.files))
      setProcessedData(null)
      setErrorMessage(null)
    }
  }

  const handleProcess = async () => {
    if (!excelFile || images.length === 0) return
    setIsLoading(true)
    setErrorMessage(null)

    try {
      const response = await processGrades(
        excelFile,
        images,
        selectedSheet || undefined,
      )
      setProcessedData(response)
      setStudentsData(response.result.students)
      setSelectedSheet(response.selected_sheet)
    } catch (err: unknown) {
      const msg =
        err instanceof Error
          ? err.message
          : 'Error al procesar las calificaciones'
      setErrorMessage(msg)
    } finally {
      setIsLoading(false)
    }
  }

  const handleScoreChange = (index: number, column: string, value: string) => {
    const updated = [...studentsData]
    const numericValue = value === '' ? null : Number(value)
    updated[index].scores[column] = isNaN(numericValue as number)
      ? null
      : numericValue
    setStudentsData(updated)
  }

  const handleExport = async () => {
    if (!excelFile || !processedData) return
    setIsLoading(true)
    try {
      const currentSheetObj = processedData.sheets.find(
        (s) => s.sheet_name === selectedSheet,
      )
      const columnMappings: Record<string, string> = {}

      processedData.result.columns_detected.forEach((col, idx) => {
        if (currentSheetObj?.score_columns[idx]) {
          columnMappings[col] = currentSheetObj.score_columns[idx]
        } else {
          columnMappings[col] = col
        }
      })

      const blob = await exportExcel(
        excelFile,
        selectedSheet,
        columnMappings,
        studentsData.map((s) => ({
          row_number: s.row_number,
          scores: s.scores,
        })),
      )

      const url = window.URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = `calificaciones_digitalizadas_${Date.now()}.xlsx`
      document.body.appendChild(a)
      a.click()
      a.remove()
      window.URL.revokeObjectURL(url)
    } catch (err: unknown) {
      const msg =
        err instanceof Error ? err.message : 'Error al exportar archivo Excel'
      setErrorMessage(msg)
    } finally {
      setIsLoading(false)
    }
  }

  return (
    <div className="min-h-screen bg-slate-900 text-slate-100 p-6 md:p-10 font-sans">
      <div className="max-w-6xl mx-auto space-y-8">
        <header className="border-b border-slate-800 pb-4 flex justify-between items-center">
          <div>
            <h1 className="text-2xl font-bold tracking-tight text-white">
              Digitalizador Maestro
            </h1>
            <p className="text-slate-400 mt-1 text-xs">
              Extracción automatizada de calificaciones y mapeo a nóminas
              oficiales
            </p>
          </div>
          {processedData && (
            <button
              onClick={() => {
                setProcessedData(null)
                setExcelFile(null)
                setImages([])
              }}
              className="flex items-center space-x-1 text-xs text-slate-400 hover:text-white transition-colors"
            >
              <RefreshCw className="w-3.5 h-3.5" />
              <span>Nuevo Proceso</span>
            </button>
          )}
        </header>

        {errorMessage && (
          <div className="p-4 bg-rose-950/40 border border-rose-800/60 rounded-xl flex items-center space-x-3 text-rose-300 text-sm">
            <AlertCircle className="w-5 h-5 shrink-0" />
            <span>{errorMessage}</span>
          </div>
        )}

        {!processedData ? (
          <div className="space-y-6">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <div className="border border-slate-800 bg-slate-800/40 rounded-xl p-6 flex flex-col items-center justify-center text-center space-y-4">
                <FileSpreadsheet className="w-10 h-10 text-emerald-400" />
                <div>
                  <h2 className="font-medium text-slate-200">
                    Plantilla Excel Oficial
                  </h2>
                  <p className="text-xs text-slate-400 mt-0.5">
                    Formato del curso (.xlsx, .xls)
                  </p>
                </div>
                <label className="cursor-pointer bg-slate-700 hover:bg-slate-600 px-4 py-2 rounded-lg text-xs font-medium transition-colors text-white">
                  <span>
                    {excelFile ? excelFile.name : 'Seleccionar archivo'}
                  </span>
                  <input
                    type="file"
                    accept=".xlsx,.xls"
                    className="hidden"
                    onChange={handleExcelUpload}
                  />
                </label>
              </div>

              <div className="border border-slate-800 bg-slate-800/40 rounded-xl p-6 flex flex-col items-center justify-center text-center space-y-4">
                <ImageIcon className="w-10 h-10 text-blue-400" />
                <div>
                  <h2 className="font-medium text-slate-200">
                    Hojas de Calificaciones
                  </h2>
                  <p className="text-xs text-slate-400 mt-0.5">
                    Imágenes escaneadas o fotografías
                  </p>
                </div>
                <label className="cursor-pointer bg-slate-700 hover:bg-slate-600 px-4 py-2 rounded-lg text-xs font-medium transition-colors text-white">
                  <span>
                    {images.length > 0
                      ? `${images.length} imágenes listas`
                      : 'Seleccionar imágenes'}
                  </span>
                  <input
                    type="file"
                    accept="image/*"
                    multiple
                    className="hidden"
                    onChange={handleImagesUpload}
                  />
                </label>
              </div>
            </div>

            <div className="flex justify-end">
              <button
                onClick={handleProcess}
                disabled={!excelFile || images.length === 0 || isLoading}
                className="flex items-center space-x-2 bg-emerald-600 hover:bg-emerald-500 disabled:opacity-40 disabled:cursor-not-allowed px-6 py-2.5 rounded-lg text-sm font-medium transition-colors text-white shadow-sm"
              >
                {isLoading ? (
                  <>
                    <Loader2 className="w-4 h-4 animate-spin" />
                    <span>Procesando imágenes con OCR...</span>
                  </>
                ) : (
                  <>
                    <span>Procesar Calificaciones</span>
                    <ArrowRight className="w-4 h-4" />
                  </>
                )}
              </button>
            </div>
          </div>
        ) : (
          <div className="space-y-6">
            <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 bg-slate-800/30 p-4 border border-slate-800 rounded-xl">
              <div className="flex items-center space-x-3">
                <CheckCircle2 className="w-5 h-5 text-emerald-400 shrink-0" />
                <div>
                  <h2 className="text-sm font-medium text-white">
                    Resultados de Coincidencia OCR
                  </h2>
                  <p className="text-xs text-slate-400">
                    {studentsData.length} estudiantes procesados. Puedes revisar
                    y editar los puntajes antes de exportar.
                  </p>
                </div>
              </div>

              <div className="flex items-center space-x-3">
                {processedData.sheets.length > 1 && (
                  <select
                    value={selectedSheet}
                    onChange={(e) => setSelectedSheet(e.target.value)}
                    className="bg-slate-800 border border-slate-700 rounded-lg px-3 py-1.5 text-xs text-slate-200 outline-none focus:border-emerald-500"
                  >
                    {processedData.sheets.map((sheet) => (
                      <option key={sheet.sheet_name} value={sheet.sheet_name}>
                        Hoja: {sheet.sheet_name}
                      </option>
                    ))}
                  </select>
                )}

                <button
                  onClick={handleExport}
                  disabled={isLoading}
                  className="flex items-center space-x-2 bg-emerald-600 hover:bg-emerald-500 disabled:opacity-50 px-4 py-2 rounded-lg text-xs font-medium text-white transition-colors"
                >
                  <Download className="w-4 h-4" />
                  <span>Exportar Excel</span>
                </button>
              </div>
            </div>

            <div className="border border-slate-800 rounded-xl overflow-hidden bg-slate-800/20">
              <div className="overflow-x-auto max-h-125">
                <table className="w-full text-left text-xs border-collapse">
                  <thead className="bg-slate-800 text-slate-300 sticky top-0 uppercase font-semibold text-[11px] tracking-wider">
                    <tr>
                      <th className="p-3 border-b border-slate-700 w-16">
                        Fila
                      </th>
                      <th className="p-3 border-b border-slate-700">
                        Nombre del Estudiante
                      </th>
                      <th className="p-3 border-b border-slate-700 w-24">
                        Confianza
                      </th>
                      {processedData.result.columns_detected.map((col) => (
                        <th
                          key={col}
                          className="p-3 border-b border-slate-700 text-center w-28"
                        >
                          {col}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800 text-slate-300 font-mono">
                    {studentsData.map((student, idx) => (
                      <tr
                        key={student.row_number}
                        className="hover:bg-slate-800/40 transition-colors"
                      >
                        <td className="p-3 text-slate-500 font-sans">
                          {student.row_number}
                        </td>
                        <td className="p-3 font-sans font-medium text-slate-200">
                          {student.full_name}
                        </td>
                        <td className="p-3">
                          <span
                            className={`px-2 py-0.5 rounded text-[10px] font-sans font-medium ${
                              student.confidence >= 0.85
                                ? 'bg-emerald-950 text-emerald-400 border border-emerald-800'
                                : student.confidence >= 0.6
                                  ? 'bg-amber-950 text-amber-400 border border-amber-800'
                                  : 'bg-rose-950 text-rose-400 border border-rose-800'
                            }`}
                          >
                            {Math.round(student.confidence * 100)}%
                          </span>
                        </td>
                        {processedData.result.columns_detected.map((col) => (
                          <td key={col} className="p-2 text-center">
                            <input
                              type="number"
                              step="any"
                              value={
                                student.scores[col] !== null
                                  ? student.scores[col]!
                                  : ''
                              }
                              onChange={(e) =>
                                handleScoreChange(idx, col, e.target.value)
                              }
                              className="w-20 bg-slate-900 border border-slate-700 rounded px-2 py-1 text-center text-xs text-white focus:outline-none focus:border-emerald-500"
                            />
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  )
}
