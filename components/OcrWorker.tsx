'use client'

import { useEffect, useRef } from 'react'

interface Props {
  imageBase64: string | null
  mimeType: string
  onResult: (text: string) => void
  onError: (msg: string) => void
}

export default function OcrWorker({ imageBase64, mimeType, onResult, onError }: Props) {
  const ranRef = useRef(false)

  useEffect(() => {
    if (!imageBase64 || ranRef.current) return
    ranRef.current = true

    async function run() {
      try {
        const { createWorker } = await import('tesseract.js')
        const worker = await createWorker(['chi_sim', 'eng'])
        const dataUrl = `data:${mimeType};base64,${imageBase64}`
        const { data } = await worker.recognize(dataUrl)
        await worker.terminate()
        onResult(data.text)
      } catch (e) {
        onError('OCR 失败: ' + String(e))
      }
    }

    run()
  }, [imageBase64, mimeType, onResult, onError])

  return null
}
