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
      return new Response(buffer as BodyInit, {
        headers: {
          'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
          'Content-Disposition': `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`,
        },
      })
    }

    return Response.json({ error: '不支持的格式' }, { status: 400 })
  } catch (error) {
    console.error('Generate error:', error)
    return Response.json({ error: 'PPT 生成失败' }, { status: 500 })
  }
}
