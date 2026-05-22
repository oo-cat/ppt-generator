import { NextRequest } from 'next/server'
import { generatePptx } from '@/lib/pptx'
import { generateRevealHtml } from '@/lib/reveal'
import { PresentationData } from '@/types'

export async function POST(request: NextRequest) {
  try {
    const body = await request.json()
    const { presentation, format }: { presentation: PresentationData; format: 'pptx' | 'html' } = body

    if (!presentation || !format) {
      return Response.json({ error: '缺少必要参数' }, { status: 400 })
    }

    if (format === 'html') {
      const html = generateRevealHtml(presentation)
      return new Response(html, {
        headers: {
          'Content-Type': 'text/html; charset=utf-8',
        },
      })
    }

    if (format === 'pptx') {
      const buffer = await generatePptx(presentation)
      const filename = `${presentation.title.replace(/[^\w一-龥]/g, '_')}.pptx`
      // Use a clean copy to avoid byteOffset issues
      const arrayBuffer = buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength) as ArrayBuffer
      return new Response(arrayBuffer, {
        headers: {
          'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
          'Content-Disposition': `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`,
        },
      })
    }

    return Response.json({ error: '不支持的格式' }, { status: 400 })
  } catch (error) {
    const msg = error instanceof Error ? error.message : String(error)
    console.error('Generate error:', msg)
    return Response.json({ error: `PPT 生成失败：${msg}` }, { status: 500 })
  }
}
