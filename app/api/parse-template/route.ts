import { NextRequest } from 'next/server'
import JSZip from 'jszip'
import { TemplateTheme } from '@/types'

const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData()
    const file = formData.get('file') as File | null

    if (!file) {
      return Response.json({ error: '未收到文件' }, { status: 400 })
    }

    const fileName = file.name.toLowerCase()
    if (file.type !== PPTX_MIME && !fileName.endsWith('.pptx')) {
      return Response.json({ error: '请上传 .pptx 格式的模板文件' }, { status: 400 })
    }

    const arrayBuffer = await file.arrayBuffer()
    const theme = await extractTheme(arrayBuffer)
    return Response.json({ theme })
  } catch (error) {
    console.error('Parse template error:', error)
    return Response.json({ error: '模板解析失败' }, { status: 500 })
  }
}

async function extractTheme(buffer: ArrayBuffer): Promise<TemplateTheme> {
  const zip = await JSZip.loadAsync(buffer)

  // 默认值
  let bgColor = 'FFFFFF'
  let titleColor = '1F3864'
  let bodyColor = '333333'
  let accentColor = '4472C4'
  let fontName = '微软雅黑'

  // 读取主题文件
  const themeFile = zip.file(/ppt\/theme\/theme\d*\.xml/)[0]
  if (themeFile) {
    const xml = await themeFile.async('text')

    // 提取配色方案中的颜色（dk1=深色1/标题, dk2=深色2, accent1=强调色1, lt1=浅色1/背景）
    const lt1 = xml.match(/<a:lt1>[\s\S]*?<a:srgbClr val="([0-9A-Fa-f]{6})"/)
    const dk1 = xml.match(/<a:dk1>[\s\S]*?<a:srgbClr val="([0-9A-Fa-f]{6})"/)
    const dk2 = xml.match(/<a:dk2>[\s\S]*?<a:srgbClr val="([0-9A-Fa-f]{6})"/)
    const accent1 = xml.match(/<a:accent1>[\s\S]*?<a:srgbClr val="([0-9A-Fa-f]{6})"/)

    if (lt1?.[1]) bgColor = lt1[1]
    if (dk1?.[1]) titleColor = dk1[1]
    if (dk2?.[1]) bodyColor = dk2[1]
    if (accent1?.[1]) accentColor = accent1[1]

    // 提取字体
    const latinFont = xml.match(/<a:latin typeface="([^"]+)"/)
    if (latinFont?.[1] && latinFont[1] !== '+mj-lt' && latinFont[1] !== '+mn-lt') {
      fontName = latinFont[1]
    }
  }

  // 读取第一张幻灯片的背景（可能覆盖主题背景）
  const slide1 = zip.file('ppt/slides/slide1.xml')
  if (slide1) {
    const xml = await slide1.async('text')
    const solidBg = xml.match(/<p:bg>[\s\S]*?<a:srgbClr val="([0-9A-Fa-f]{6})"/)
    if (solidBg?.[1]) bgColor = solidBg[1]
  }

  return { bgColor, titleColor, bodyColor, accentColor, fontName }
}
