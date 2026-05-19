import { NextRequest } from 'next/server'

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData()
    const file = formData.get('file') as File | null

    if (!file) {
      return Response.json({ error: '未收到文件' }, { status: 400 })
    }

    const fileType = file.type
    const fileName = file.name.toLowerCase()
    const arrayBuffer = await file.arrayBuffer()
    const base64 = Buffer.from(arrayBuffer).toString('base64')

    if (fileType === 'application/pdf') {
      const text = await extractPdfText(arrayBuffer)
      return Response.json({ type: 'text', content: text })
    }

    if (fileType === DOCX_MIME || fileName.endsWith('.docx')) {
      const text = await extractDocxText(arrayBuffer)
      return Response.json({ type: 'text', content: text })
    }

    if (fileType.startsWith('image/')) {
      return Response.json({ type: 'image', base64, mimeType: fileType })
    }

    if (fileType === 'text/plain' || fileType === 'text/markdown' || fileName.endsWith('.md')) {
      const text = new TextDecoder().decode(arrayBuffer)
      return Response.json({ type: 'text', content: text })
    }

    return Response.json({ error: '不支持的文件类型，支持：PDF、Word(.docx)、图片、TXT、MD' }, { status: 400 })
  } catch (error) {
    console.error('Parse file error:', error)
    return Response.json({ error: '文件解析失败' }, { status: 500 })
  }
}

async function extractPdfText(buffer: ArrayBuffer): Promise<string> {
  // eslint-disable-next-line @typescript-eslint/no-require-imports
  const pdfjsLib = await import('pdfjs-dist/legacy/build/pdf.mjs')
  pdfjsLib.GlobalWorkerOptions.workerSrc = ''

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const pdf = await (pdfjsLib as any).getDocument({
    data: new Uint8Array(buffer),
    useWorkerFetch: false,
    isEvalSupported: false,
    useSystemFonts: true,
  }).promise
  const pages: string[] = []

  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i)
    const content = await page.getTextContent()
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const text = content.items.map((item: any) => item.str ?? '').join(' ')
    pages.push(text)
  }

  return pages.join('\n\n')
}

async function extractDocxText(buffer: ArrayBuffer): Promise<string> {
  const mammoth = await import('mammoth')
  const result = await mammoth.extractRawText({ buffer: Buffer.from(buffer) })
  return result.value
}
