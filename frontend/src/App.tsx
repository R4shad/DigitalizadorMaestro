import { useState } from 'react'
import { FileSpreadsheet, Image as ImageIcon, ArrowRight } from 'lucide-react'

export default function App() {
  const [excelFile, setExcelFile] = useState<File | null>(null)
  const [images, setImages] = useState<File[]>([])

  const handleExcelUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setExcelFile(e.target.files[0])
    }
  }

  const handleImagesUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      setImages(Array.from(e.target.files))
    }
  }

  return (
    <div className="min-h-screen bg-slate-900 text-slate-100 p-8">
      <div className="max-w-4xl mx-auto space-y-8">
        <header className="border-b border-slate-700 pb-4">
          <h1 className="text-3xl font-bold tracking-tight text-white">Digitalizador Maestro</h1>
          <p className="text-slate-400 mt-1 text-sm">
            Digitalización de calificaciones sin conexión y mapeo automático de notas
          </p>
        </header>

        <main className="grid grid-cols-1 md:grid-cols-2 gap-6">
          <div className="border border-slate-800 bg-slate-800/50 rounded-xl p-6 flex flex-col items-center justify-center text-center space-y-4">
            <FileSpreadsheet className="w-12 h-12 text-emerald-400" />
            <div>
              <h2 className="font-semibold text-lg">Plantilla Excel</h2>
              <p className="text-xs text-slate-400">Sube la nómina oficial del curso (.xlsx)</p>
            </div>
            <label className="cursor-pointer bg-slate-700 hover:bg-slate-600 px-4 py-2 rounded-lg text-sm font-medium transition-colors">
              <span>{excelFile ? excelFile.name : 'Seleccionar archivo'}</span>
              <input type="file" accept=".xlsx,.xls" className="hidden" onChange={handleExcelUpload} />
            </label>
          </div>

          <div className="border border-slate-800 bg-slate-800/50 rounded-xl p-6 flex flex-col items-center justify-center text-center space-y-4">
            <ImageIcon className="w-12 h-12 text-blue-400" />
            <div>
              <h2 className="font-semibold text-lg">Fotos de Calificaciones</h2>
              <p className="text-xs text-slate-400">Sube las fotos de las hojas impresas</p>
            </div>
            <label className="cursor-pointer bg-slate-700 hover:bg-slate-600 px-4 py-2 rounded-lg text-sm font-medium transition-colors">
              <span>{images.length > 0 ? `${images.length} imágenes seleccionadas` : 'Seleccionar imágenes'}</span>
              <input type="file" accept="image/*" multiple className="hidden" onChange={handleImagesUpload} />
            </label>
          </div>
        </main>

        <div className="flex justify-end">
          <button
            disabled={!excelFile || images.length === 0}
            className="flex items-center space-x-2 bg-emerald-600 hover:bg-emerald-500 disabled:opacity-40 disabled:cursor-not-allowed px-6 py-2.5 rounded-lg font-medium transition-colors text-white"
          >
            <span>Procesar Calificaciones</span>
            <ArrowRight className="w-4 h-4" />
          </button>
        </div>
      </div>
    </div>
  )
}