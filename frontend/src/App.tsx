import { useState, useMemo } from 'react'
import {
  FileSpreadsheet,
  Image as ImageIcon,
  ArrowRight,
  ArrowLeft,
  Download,
  Loader2,
  AlertCircle,
  CheckCircle2,
  AlertTriangle,
  RefreshCw,
  Settings2,
  Search,
} from 'lucide-react'
import { processGrades, injectDirectExcel } from './services/api.ts'
import type {
  ProcessGradesResponse,
  MatchedStudentGrade,
  ColumnMappingConfig,
  DimensionType,
  DimensionRule,
} from './types/index.ts'

const DIMENSION_RULES: Record<DimensionType, DimensionRule> = {
  SER: { type: 'SER', min: 1, max: 10, label: 'Ser (1 - 10)' },
  SABER: { type: 'SABER', min: 1, max: 45, label: 'Saber (1 - 45)' },
  HACER: { type: 'HACER', min: 1, max: 50, label: 'Hacer (1 - 50)' },
  AUTOEVALUACION: {
    type: 'AUTOEVALUACION',
    min: 1,
    max: 5,
    label: 'Autoevaluación (1 - 5)',
  },
  CUSTOM: {
    type: 'CUSTOM',
    min: 0,
    max: 100,
    label: 'Personalizado (0 - 100)',
  },
}

export default function App() {
  const [step, setStep] = useState<1 | 2 | 3>(1)
  const [excelFile, setExcelFile] = useState<File | null>(null)
  const [images, setImages] = useState<File[]>([])
  const [isLoading, setIsLoading] = useState(false)
  const [errorMessage, setErrorMessage] = useState<string | null>(null)

  const [processedData, setProcessedData] =
    useState<ProcessGradesResponse | null>(null)
  const [selectedSheet, setSelectedSheet] = useState<string>('')
  const [mappings, setMappings] = useState<ColumnMappingConfig[]>([])
  const [studentsData, setStudentsData] = useState<MatchedStudentGrade[]>([])

  const [filterOnlyIssues, setFilterOnlyIssues] = useState(false)
  const [selectedColumnFilter, setSelectedColumnFilter] =
    useState<string>('ALL')

  const handleExcelUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setExcelFile(e.target.files[0])
      resetAll()
    }
  }

  const handleImagesUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      setImages(Array.from(e.target.files))
      resetAll()
    }
  }

  const resetAll = () => {
    setProcessedData(null)
    setErrorMessage(null)
    setStep(1)
    setMappings([])
    setStudentsData([])
  }

  const handleStartProcess = async () => {
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
      setSelectedSheet(response.selected_sheet)
      setStudentsData(response.result.students)

      const activeSheetObj = response.sheets.find(
        (s) => s.sheet_name === response.selected_sheet,
      )
      const initialMappings: ColumnMappingConfig[] =
        response.result.columns_detected.map((col, idx) => {
          let detectedDim: DimensionType = 'CUSTOM'
          const lower = col.toLowerCase()
          if (lower.includes('ser')) detectedDim = 'SER'
          else if (lower.includes('saber')) detectedDim = 'SABER'
          else if (lower.includes('hacer')) detectedDim = 'HACER'
          else if (lower.includes('auto') || lower.includes('decidir'))
            detectedDim = 'AUTOEVALUACION'

          return {
            ocrColumn: col,
            excelColumn: activeSheetObj?.score_columns[idx] || col,
            dimension: detectedDim,
          }
        })
      setMappings(initialMappings)
    } catch (err: unknown) {
      const msg =
        err instanceof Error ? err.message : 'Error al procesar los archivos'
      setErrorMessage(msg)
    } finally {
      setIsLoading(false)
    }
  }

  const checkGradeAnomaly = (
    val: number | null,
    dim: DimensionType,
  ): string | null => {
    if (val === null || val === undefined) return 'Sin nota'
    const rule = DIMENSION_RULES[dim]
    if (val < rule.min || val > rule.max) {
      return `Fuera de rango (${rule.min}-${rule.max})`
    }
    return null
  }

  const handleScoreChange = (
    rowNumber: number,
    column: string,
    value: string,
  ) => {
    setStudentsData((prev) =>
      prev.map((s) => {
        if (s.row_number === rowNumber) {
          const numericValue = value === '' ? null : Number(value)
          return {
            ...s,
            scores: {
              ...s.scores,
              [column]: isNaN(numericValue as number) ? null : numericValue,
            },
          }
        }
        return s
      }),
    )
  }

  const handleExecuteInjection = async () => {
    if (!excelFile || !processedData) return
    setIsLoading(true)
    try {
      const blob = await injectDirectExcel(
        excelFile,
        selectedSheet,
        mappings,
        studentsData,
      )
      const url = window.URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = `calificaciones_consolidadas_${Date.now()}.xlsx`
      document.body.appendChild(a)
      a.click()
      a.remove()
      window.URL.revokeObjectURL(url)
    } catch (err: unknown) {
      const msg =
        err instanceof Error
          ? err.message
          : 'Error al inyectar notas mediante PowerShell'
      setErrorMessage(msg)
    } finally {
      setIsLoading(false)
    }
  }

  const criticalCasesCount = useMemo(() => {
    let count = 0
    studentsData.forEach((student) => {
      if (student.confidence < 0.8) {
        count++
        return
      }
      for (const m of mappings) {
        if (checkGradeAnomaly(student.scores[m.ocrColumn], m.dimension)) {
          count++
          break
        }
      }
    })
    return count
  }, [studentsData, mappings])

  const filteredStudents = useMemo(() => {
    return studentsData.filter((student) => {
      const hasConfidenceIssue = student.confidence < 0.8
      let hasAnomaly = false

      for (const m of mappings) {
        if (
          selectedColumnFilter !== 'ALL' &&
          m.ocrColumn !== selectedColumnFilter
        )
          continue
        if (checkGradeAnomaly(student.scores[m.ocrColumn], m.dimension)) {
          hasAnomaly = true
          break
        }
      }

      if (filterOnlyIssues) {
        return hasConfidenceIssue || hasAnomaly
      }
      return true
    })
  }, [studentsData, mappings, filterOnlyIssues, selectedColumnFilter])

  return (
    <div className="min-h-screen bg-slate-950 text-slate-100 p-6 md:p-10 font-sans">
      <div className="max-w-6xl mx-auto space-y-6">
        <header className="border-b border-slate-800 pb-4 flex justify-between items-center">
          <div>
            <h1 className="text-2xl font-bold tracking-tight text-white">
              Digitalizador Maestro
            </h1>
            <p className="text-slate-400 mt-1 text-xs">
              Extracción OCR con fidelización y pegado nativo en Excel
            </p>
          </div>
          {processedData && (
            <button
              onClick={resetAll}
              className="flex items-center space-x-1 text-xs text-slate-400 hover:text-white transition-colors"
            >
              <RefreshCw className="w-3.5 h-3.5" />
              <span>Reiniciar</span>
            </button>
          )}
        </header>

        {errorMessage && (
          <div className="p-4 bg-rose-950/40 border border-rose-800/60 rounded-xl flex items-center space-x-3 text-rose-300 text-sm">
            <AlertCircle className="w-5 h-5 shrink-0" />
            <span>{errorMessage}</span>
          </div>
        )}

        {/* Wizard Progress Indicator */}
        <div className="grid grid-cols-3 gap-2 border border-slate-800 p-1.5 rounded-xl bg-slate-900/50 text-xs text-center font-medium">
          <div
            className={`py-2 rounded-lg transition-colors ${
              step === 1
                ? 'bg-emerald-600 text-white shadow-sm'
                : 'text-slate-400'
            }`}
          >
            1. Fidelización de Columnas
          </div>
          <div
            className={`py-2 rounded-lg transition-colors ${
              step === 2
                ? 'bg-emerald-600 text-white shadow-sm'
                : 'text-slate-400'
            }`}
          >
            2. Validación de Notas ({criticalCasesCount} alertas)
          </div>
          <div
            className={`py-2 rounded-lg transition-colors ${
              step === 3
                ? 'bg-emerald-600 text-white shadow-sm'
                : 'text-slate-400'
            }`}
          >
            3. Inyección y Exportación
          </div>
        </div>

        {/* PASO 1: Carga y Fidelización */}
        {step === 1 && (
          <div className="space-y-6">
            {!processedData ? (
              <div className="space-y-6">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  <div className="border border-slate-800 bg-slate-900/40 rounded-xl p-6 flex flex-col items-center justify-center text-center space-y-4">
                    <FileSpreadsheet className="w-10 h-10 text-emerald-400" />
                    <div>
                      <h2 className="font-medium text-slate-200">
                        Plantilla Excel Oficial
                      </h2>
                      <p className="text-xs text-slate-400 mt-0.5">
                        Archivo original de calificaciones (.xlsx)
                      </p>
                    </div>
                    <label className="cursor-pointer bg-slate-800 hover:bg-slate-700 px-4 py-2 rounded-lg text-xs font-medium transition-colors text-white border border-slate-700">
                      <span>
                        {excelFile ? excelFile.name : 'Seleccionar Excel'}
                      </span>
                      <input
                        type="file"
                        accept=".xlsx,.xls"
                        className="hidden"
                        onChange={handleExcelUpload}
                      />
                    </label>
                  </div>

                  <div className="border border-slate-800 bg-slate-900/40 rounded-xl p-6 flex flex-col items-center justify-center text-center space-y-4">
                    <ImageIcon className="w-10 h-10 text-blue-400" />
                    <div>
                      <h2 className="font-medium text-slate-200">
                        Fotos de Calificaciones
                      </h2>
                      <p className="text-xs text-slate-400 mt-0.5">
                        Hojas físicas con notas impresas o manuscritas
                      </p>
                    </div>
                    <label className="cursor-pointer bg-slate-800 hover:bg-slate-700 px-4 py-2 rounded-lg text-xs font-medium transition-colors text-white border border-slate-700">
                      <span>
                        {images.length > 0
                          ? `${images.length} imágenes listas`
                          : 'Seleccionar Imágenes'}
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
                    onClick={handleStartProcess}
                    disabled={!excelFile || images.length === 0 || isLoading}
                    className="flex items-center space-x-2 bg-emerald-600 hover:bg-emerald-500 disabled:opacity-40 disabled:cursor-not-allowed px-6 py-2.5 rounded-lg text-sm font-medium transition-colors text-white"
                  >
                    {isLoading ? (
                      <>
                        <Loader2 className="w-4 h-4 animate-spin" />
                        <span>Analizando imágenes y planilla...</span>
                      </>
                    ) : (
                      <>
                        <span>Iniciar Análisis y Mapeo</span>
                        <ArrowRight className="w-4 h-4" />
                      </>
                    )}
                  </button>
                </div>
              </div>
            ) : (
              <div className="space-y-6">
                <div className="border border-slate-800 rounded-xl p-5 bg-slate-900/40 space-y-4">
                  <div className="flex justify-between items-center">
                    <div className="flex items-center space-x-2">
                      <Settings2 className="w-5 h-5 text-emerald-400" />
                      <h2 className="text-sm font-semibold text-white">
                        Configuración de Hoja y Columnas Destino
                      </h2>
                    </div>
                    {processedData.sheets.length > 1 && (
                      <select
                        value={selectedSheet}
                        onChange={(e) => setSelectedSheet(e.target.value)}
                        className="bg-slate-800 border border-slate-700 rounded-lg px-3 py-1.5 text-xs text-slate-200 outline-none"
                      >
                        {processedData.sheets.map((sheet) => (
                          <option
                            key={sheet.sheet_name}
                            value={sheet.sheet_name}
                          >
                            Hoja Destino: {sheet.sheet_name}
                          </option>
                        ))}
                      </select>
                    )}
                  </div>

                  <div className="divide-y divide-slate-800 border border-slate-800 rounded-lg overflow-hidden">
                    {mappings.map((mapping, idx) => (
                      <div
                        key={mapping.ocrColumn}
                        className="p-3 bg-slate-900/80 flex flex-wrap items-center gap-4 text-xs"
                      >
                        <span className="font-mono font-medium text-slate-300 w-32 truncate">
                          {mapping.ocrColumn}
                        </span>
                        <div className="flex items-center space-x-2">
                          <span className="text-slate-400 text-[11px]">
                            Columna Excel:
                          </span>
                          <input
                            type="text"
                            value={mapping.excelColumn}
                            onChange={(e) => {
                              const updated = [...mappings]
                              updated[idx].excelColumn =
                                e.target.value.toUpperCase()
                              setMappings(updated)
                            }}
                            className="w-16 bg-slate-950 border border-slate-700 rounded px-2 py-1 text-center font-mono uppercase"
                          />
                        </div>
                        <div className="flex items-center space-x-2">
                          <span className="text-slate-400 text-[11px]">
                            Dimensión:
                          </span>
                          <select
                            value={mapping.dimension}
                            onChange={(e) => {
                              const updated = [...mappings]
                              updated[idx].dimension = e.target
                                .value as DimensionType
                              setMappings(updated)
                            }}
                            className="bg-slate-950 border border-slate-700 rounded px-2 py-1 text-slate-200 outline-none"
                          >
                            {Object.values(DIMENSION_RULES).map((dim) => (
                              <option key={dim.type} value={dim.type}>
                                {dim.label}
                              </option>
                            ))}
                          </select>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>

                <div className="flex justify-end">
                  <button
                    onClick={() => setStep(2)}
                    className="flex items-center space-x-2 bg-emerald-600 hover:bg-emerald-500 px-6 py-2 rounded-lg text-xs font-medium text-white transition-colors"
                  >
                    <span>Continuar a Validación de Notas</span>
                    <ArrowRight className="w-4 h-4" />
                  </button>
                </div>
              </div>
            )}
          </div>
        )}

        {/* PASO 2: Validación de Casos de Baja Confianza */}
        {step === 2 && (
          <div className="space-y-4">
            <div className="flex flex-wrap items-center justify-between gap-3 bg-slate-900/60 p-4 border border-slate-800 rounded-xl">
              <div className="flex items-center space-x-2">
                <Search className="w-4 h-4 text-slate-400" />
                <span className="text-xs text-slate-300">Filtrar columna:</span>
                <select
                  value={selectedColumnFilter}
                  onChange={(e) => setSelectedColumnFilter(e.target.value)}
                  className="bg-slate-800 border border-slate-700 rounded-lg px-2.5 py-1 text-xs text-slate-200 outline-none"
                >
                  <option value="ALL">Todas las columnas</option>
                  {mappings.map((m) => (
                    <option key={m.ocrColumn} value={m.ocrColumn}>
                      {m.ocrColumn} ({m.dimension})
                    </option>
                  ))}
                </select>
              </div>

              <label className="flex items-center space-x-2 cursor-pointer text-xs">
                <input
                  type="checkbox"
                  checked={filterOnlyIssues}
                  onChange={(e) => setFilterOnlyIssues(e.target.checked)}
                  className="rounded bg-slate-800 border-slate-700 text-emerald-500 focus:ring-0"
                />
                <span className="text-amber-400 font-medium">
                  Mostrar solo casos con alertas o dudas ({criticalCasesCount})
                </span>
              </label>
            </div>

            <div className="border border-slate-800 rounded-xl overflow-hidden bg-slate-900/30">
              <div className="overflow-x-auto max-h-125">
                <table className="w-full text-left text-xs border-collapse">
                  <thead className="bg-slate-900 text-slate-300 sticky top-0 uppercase font-semibold text-[11px] tracking-wider">
                    <tr>
                      <th className="p-3 border-b border-slate-800 w-16">
                        Fila
                      </th>
                      <th className="p-3 border-b border-slate-800">
                        Estudiante
                      </th>
                      <th className="p-3 border-b border-slate-800 w-24">
                        Confianza
                      </th>
                      {mappings.map((m) => (
                        <th
                          key={m.ocrColumn}
                          className="p-3 border-b border-slate-800 text-center w-32"
                        >
                          <div>{m.ocrColumn}</div>
                          <div className="text-[9px] text-slate-500 font-normal">
                            {DIMENSION_RULES[m.dimension].label}
                          </div>
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800 font-mono">
                    {filteredStudents.map((student) => {
                      const isLowConfidence = student.confidence < 0.8
                      return (
                        <tr
                          key={student.row_number}
                          className="hover:bg-slate-800/30 transition-colors"
                        >
                          <td className="p-3 text-slate-500 font-sans">
                            {student.row_number}
                          </td>
                          <td className="p-3 font-sans">
                            <div className="font-medium text-slate-200">
                              {student.full_name}
                            </div>
                            {student.matched_ocr_text && (
                              <div className="text-[10px] text-slate-500 font-mono">
                                OCR: {student.matched_ocr_text}
                              </div>
                            )}
                          </td>
                          <td className="p-3 font-sans">
                            <span
                              className={`px-2 py-0.5 rounded text-[10px] font-medium ${
                                !isLowConfidence
                                  ? 'bg-emerald-950 text-emerald-400 border border-emerald-800'
                                  : 'bg-rose-950 text-rose-400 border border-rose-800'
                              }`}
                            >
                              {Math.round(student.confidence * 100)}%
                            </span>
                          </td>
                          {mappings.map((m) => {
                            const val = student.scores[m.ocrColumn]
                            const anomaly = checkGradeAnomaly(val, m.dimension)
                            return (
                              <td key={m.ocrColumn} className="p-2 text-center">
                                <div className="flex flex-col items-center">
                                  <input
                                    type="number"
                                    step="any"
                                    value={
                                      val !== null && val !== undefined
                                        ? val
                                        : ''
                                    }
                                    onChange={(e) =>
                                      handleScoreChange(
                                        student.row_number,
                                        m.ocrColumn,
                                        e.target.value,
                                      )
                                    }
                                    className={`w-20 border rounded px-2 py-1 text-center text-xs text-white focus:outline-none ${
                                      anomaly
                                        ? 'bg-rose-950/50 border-rose-700 focus:border-rose-500'
                                        : 'bg-slate-950 border-slate-700 focus:border-emerald-500'
                                    }`}
                                  />
                                  {anomaly && (
                                    <span className="text-[9px] text-rose-400 font-sans mt-0.5 flex items-center">
                                      <AlertTriangle className="w-2.5 h-2.5 mr-0.5 shrink-0" />
                                      {anomaly}
                                    </span>
                                  )}
                                </div>
                              </td>
                            )
                          })}
                        </tr>
                      )
                    })}
                  </tbody>
                </table>
              </div>
            </div>

            <div className="flex justify-between items-center pt-2">
              <button
                onClick={() => setStep(1)}
                className="flex items-center space-x-1 bg-slate-800 hover:bg-slate-700 px-4 py-2 rounded-lg text-xs font-medium text-slate-300 transition-colors"
              >
                <ArrowLeft className="w-4 h-4" />
                <span>Volver a Mapeo</span>
              </button>

              <button
                onClick={() => setStep(3)}
                className="flex items-center space-x-2 bg-emerald-600 hover:bg-emerald-500 px-6 py-2 rounded-lg text-xs font-medium text-white transition-colors"
              >
                <span>Finalizar y Preparar Exportación</span>
                <ArrowRight className="w-4 h-4" />
              </button>
            </div>
          </div>
        )}

        {/* PASO 3: Inyección Segura */}
        {step === 3 && (
          <div className="space-y-6">
            <div className="border border-slate-800 rounded-xl p-6 bg-slate-900/40 space-y-4">
              <div className="flex items-center space-x-3 text-emerald-400">
                <CheckCircle2 className="w-6 h-6" />
                <h2 className="font-semibold text-lg text-white">
                  Todo listo para la inyección
                </h2>
              </div>
              <p className="text-xs text-slate-400 leading-relaxed">
                Se actualizarán las calificaciones validadas directamente en el
                archivo original mediante el motor nativo de Excel (PowerShell
                COM Object), conservando todas las fórmulas y el formato
                institucional.
              </p>

              <div className="grid grid-cols-2 md:grid-cols-4 gap-4 pt-2 text-xs">
                <div className="bg-slate-950 p-3 rounded-lg border border-slate-800">
                  <div className="text-slate-400">Hoja Destino</div>
                  <div className="font-semibold text-slate-200 mt-1">
                    {selectedSheet}
                  </div>
                </div>
                <div className="bg-slate-950 p-3 rounded-lg border border-slate-800">
                  <div className="text-slate-400">Estudiantes</div>
                  <div className="font-semibold text-slate-200 mt-1">
                    {studentsData.length}
                  </div>
                </div>
                <div className="bg-slate-950 p-3 rounded-lg border border-slate-800">
                  <div className="text-slate-400">Columnas Mapeadas</div>
                  <div className="font-semibold text-slate-200 mt-1">
                    {mappings.length}
                  </div>
                </div>
                <div className="bg-slate-950 p-3 rounded-lg border border-slate-800">
                  <div className="text-slate-400">Alertas Pendientes</div>
                  <div
                    className={`font-semibold mt-1 ${criticalCasesCount > 0 ? 'text-amber-400' : 'text-emerald-400'}`}
                  >
                    {criticalCasesCount}
                  </div>
                </div>
              </div>
            </div>

            <div className="flex justify-between items-center">
              <button
                onClick={() => setStep(2)}
                className="flex items-center space-x-1 bg-slate-800 hover:bg-slate-700 px-4 py-2 rounded-lg text-xs font-medium text-slate-300 transition-colors"
              >
                <ArrowLeft className="w-4 h-4" />
                <span>Revisar Notas</span>
              </button>

              <button
                onClick={handleExecuteInjection}
                disabled={isLoading}
                className="flex items-center space-x-2 bg-emerald-600 hover:bg-emerald-500 disabled:opacity-50 px-6 py-2.5 rounded-lg text-xs font-medium text-white transition-colors shadow-sm"
              >
                {isLoading ? (
                  <>
                    <Loader2 className="w-4 h-4 animate-spin" />
                    <span>Inyectando datos vía PowerShell...</span>
                  </>
                ) : (
                  <>
                    <Download className="w-4 h-4" />
                    <span>Inyectar y Descargar Excel Oficial</span>
                  </>
                )}
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  )
}
