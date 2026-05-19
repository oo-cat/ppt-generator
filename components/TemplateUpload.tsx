'use client'

import { useRef, useState } from 'react'
import { TemplateTheme } from '@/types'

interface Props {
  onTemplate: (theme: TemplateTheme) => void
  onClear: () => void
  hasTemplate: boolean
}

export default function TemplateUpload({ onTemplate, onClear, hasTemplate }: Props) {
  const inputRef = useRef<HTMLInputElement>(null)
  const [loading, setLoading] = useState(false)
  const [fileName, setFileName] = useState('')

  async function handleFile(file: File) {
    setLoading(true)
    setFileName(file.name)
    try {
      const formData = new FormData()
      formData.append('file', file)
      const res = await fetch('/api/parse-template', { method: 'POST', body: formData })
      const data = await res.json()
      if (data.error) {
        alert(data.error)
        setFileName('')
      } else {
        onTemplate(data.theme)
      }
    } catch {
      alert('模板上传失败')
      setFileName('')
    } finally {
      setLoading(false)
    }
  }

  function handleDrop(e: React.DragEvent) {
    e.preventDefault()
    const file = e.dataTransfer.files[0]
    if (file) handleFile(file)
  }

  function handleClear(e: React.MouseEvent) {
    e.stopPropagation()
    setFileName('')
    onClear()
  }

  return (
    <div
      className="border-2 border-dashed rounded-xl p-4 text-center cursor-pointer transition-colors border-purple-300 hover:border-purple-500 hover:bg-purple-50"
      onDrop={handleDrop}
      onDragOver={(e) => e.preventDefault()}
      onClick={() => inputRef.current?.click()}
    >
      <input
        ref={inputRef}
        type="file"
        className="hidden"
        accept=".pptx"
        onChange={(e) => {
          const file = e.target.files?.[0]
          if (file) handleFile(file)
          e.target.value = ''
        }}
      />
      {loading ? (
        <p className="text-purple-500 text-sm">解析模板中...</p>
      ) : hasTemplate && fileName ? (
        <div className="flex items-center justify-center gap-2">
          <p className="text-purple-600 text-sm">✓ {fileName}</p>
          <button
            onClick={handleClear}
            className="text-xs text-gray-400 hover:text-red-400"
          >
            ✕ 移除
          </button>
        </div>
      ) : (
        <>
          <p className="text-gray-500 text-sm">拖拽或点击上传 PPT 模板</p>
          <p className="text-gray-400 text-xs mt-1">仅支持 .pptx，将提取配色和字体</p>
        </>
      )}
    </div>
  )
}
