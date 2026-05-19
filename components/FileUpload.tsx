'use client'

import { useRef, useState } from 'react'

interface Props {
  onParsed: (content: string, fileName: string) => void
  onOcrImage: (base64: string, mimeType: string, fileName: string) => void
  disabled?: boolean
}

export default function FileUpload({ onParsed, onOcrImage, disabled }: Props) {
  const inputRef = useRef<HTMLInputElement>(null)
  const [loading, setLoading] = useState(false)
  const [fileName, setFileName] = useState('')

  async function handleFile(file: File) {
    setLoading(true)
    setFileName(file.name)
    try {
      const formData = new FormData()
      formData.append('file', file)
      const res = await fetch('/api/parse-file', { method: 'POST', body: formData })
      const data = await res.json()

      if (data.type === 'text') {
        onParsed(data.content, file.name)
      } else if (data.type === 'image') {
        onOcrImage(data.base64, data.mimeType, file.name)
      } else {
        alert(data.error ?? '解析失败')
        setFileName('')
      }
    } catch {
      alert('文件上传失败')
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

  return (
    <div
      className={`border-2 border-dashed rounded-xl p-6 text-center cursor-pointer transition-colors
        ${disabled ? 'opacity-50 cursor-not-allowed border-gray-200' : 'border-blue-300 hover:border-blue-500 hover:bg-blue-50'}`}
      onDrop={disabled ? undefined : handleDrop}
      onDragOver={(e) => e.preventDefault()}
      onClick={() => !disabled && inputRef.current?.click()}
    >
      <input
        ref={inputRef}
        type="file"
        className="hidden"
        accept=".pdf,.txt,.md,image/*"
        disabled={disabled}
        onChange={(e) => {
          const file = e.target.files?.[0]
          if (file) handleFile(file)
          e.target.value = ''
        }}
      />
      {loading ? (
        <p className="text-blue-500 text-sm">解析中...</p>
      ) : fileName ? (
        <p className="text-green-600 text-sm">✓ {fileName}</p>
      ) : (
        <>
          <p className="text-gray-500 text-sm">拖拽或点击上传</p>
          <p className="text-gray-400 text-xs mt-1">支持 PDF、图片（思维导图截图）、TXT、MD</p>
        </>
      )}
    </div>
  )
}
