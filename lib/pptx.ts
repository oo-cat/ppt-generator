import pptxgen from 'pptxgenjs'
import JSZip from 'jszip'
import { PresentationData } from '@/types'

const SLIDE_W = 13.33
const SLIDE_H = 7.5

const THEMES = {
  academic: {
    bgColor: 'FFFFFF',
    titleColor: '1F3864',
    bulletColor: '2E4057',
    accentColor: '2E75B6',
    footerColor: '999999',
    dark: false,
  },
  business: {
    bgColor: '1A1A2E',
    titleColor: 'E2E2E2',
    bulletColor: 'C8D0E0',
    accentColor: 'E94560',
    footerColor: '777777',
    dark: true,
  },
  colorful: {
    bgColor: 'FAFAFA',
    titleColor: '6C3483',
    bulletColor: '2C3E50',
    accentColor: 'E74C3C',
    footerColor: 'AAAAAA',
    dark: false,
  },
}

type ThemeConfig = {
  bgColor: string
  titleColor: string
  bulletColor: string
  accentColor: string
  footerColor: string
  dark: boolean
  fontFace?: string
}

export async function generatePptx(data: PresentationData): Promise<Uint8Array> {
  const prs = new pptxgen()
  prs.layout = 'LAYOUT_WIDE'

  const base = THEMES[data.theme] ?? THEMES.academic
  const t = data.templateTheme
  const theme: ThemeConfig = t
    ? {
        bgColor: t.bgColor,
        titleColor: t.titleColor,
        bulletColor: t.bodyColor,
        accentColor: t.accentColor,
        footerColor: 'AAAAAA',
        dark: isDarkColor(t.bgColor),
        fontFace: t.fontName,
      }
    : { ...base }

  // ── 封面 ──────────────────────────────────────────────
  const cover = prs.addSlide()
  cover.background = { color: theme.bgColor }

  cover.addShape(prs.ShapeType.rect, {
    x: 0, y: 0, w: SLIDE_W, h: 0.25,
    fill: { color: theme.accentColor },
    line: { type: 'none' },
  })
  cover.addShape(prs.ShapeType.rect, {
    x: 0, y: SLIDE_H - 0.25, w: SLIDE_W, h: 0.25,
    fill: { color: theme.accentColor },
    line: { type: 'none' },
  })
  cover.addShape(prs.ShapeType.rect, {
    x: 0.8, y: 1.8, w: SLIDE_W - 1.6, h: 3.8,
    fill: { color: theme.accentColor },
    line: { type: 'none' },
  })
  // 封面标题放在色块上，始终用白色（色块上白字 WPS 可识别）
  cover.addText(data.title, {
    x: 1.0, y: 2.0, w: SLIDE_W - 2.0, h: 3.4,
    fontSize: 36,
    bold: true,
    color: 'FFFFFF',
    align: 'center',
    valign: 'middle',
    wrap: true,
    ...(theme.fontFace ? { fontFace: theme.fontFace } : {}),
  })

  // ── 内容页 ─────────────────────────────────────────────
  for (const slide of data.slides) {
    const s = prs.addSlide()
    s.background = { color: theme.bgColor }

    s.addShape(prs.ShapeType.rect, {
      x: 0, y: 0, w: SLIDE_W, h: 0.22,
      fill: { color: theme.accentColor },
      line: { type: 'none' },
    })

    s.addText(slide.title, {
      x: 0.5, y: 0.28, w: SLIDE_W - 1.0, h: 1.0,
      fontSize: 28,
      bold: true,
      color: theme.titleColor,
      valign: 'middle',
      wrap: true,
      ...(theme.fontFace ? { fontFace: theme.fontFace } : {}),
    })

    s.addShape(prs.ShapeType.line, {
      x: 0.5, y: 1.35, w: SLIDE_W - 1.0, h: 0,
      line: { color: theme.accentColor, width: 1.5 },
    })

    if (slide.bullets && slide.bullets.length > 0) {
      const bulletItems = slide.bullets.map((b) => ({
        text: `• ${b}`,
        options: {
          fontSize: 20,
          color: theme.bulletColor,
          paraSpaceAfter: 10,
          breakLine: true,
          ...(theme.fontFace ? { fontFace: theme.fontFace } : {}),
        },
      }))
      s.addText(bulletItems, {
        x: 0.7, y: 1.5, w: SLIDE_W - 1.4, h: SLIDE_H - 2.1,
        valign: 'top',
        wrap: true,
      })
    }

    s.addText(data.title, {
      x: 0.5, y: SLIDE_H - 0.35, w: SLIDE_W - 1.0, h: 0.28,
      fontSize: 9,
      color: theme.footerColor,
      align: 'right',
    })
  }

  const arraybuf = await (prs.write({ outputType: 'arraybuffer' }) as Promise<ArrayBuffer>)

  // 深色主题：修改 theme XML 的 dk1（默认文字色）为白色
  // 这样 WPS 读取主题后，默认文字颜色就是白色，不会再用深色覆盖
  if (theme.dark) {
    return patchThemeTextColor(arraybuf, 'FFFFFF')
  }

  return new Uint8Array(arraybuf)
}

async function patchThemeTextColor(buf: ArrayBuffer, textColor: string): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(buf)
  const themeFiles = zip.file(/ppt\/theme\/theme\d*\.xml/)

  for (const file of themeFiles) {
    let xml = await file.async('text')
    // 将 dk1 颜色节点替换为指定颜色
    // 原始格式: <a:dk1><a:sysClr ...> 或 <a:dk1><a:srgbClr val="...">
    xml = xml.replace(
      /<a:dk1>[\s\S]*?<\/a:dk1>/,
      `<a:dk1><a:srgbClr val="${textColor}"/></a:dk1>`
    )
    zip.file(file.name, xml)
  }

  const result = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
  return result
}

function isDarkColor(hex: string): boolean {
  const h = hex.replace('#', '')
  const r = parseInt(h.slice(0, 2), 16) / 255
  const g = parseInt(h.slice(2, 4), 16) / 255
  const b = parseInt(h.slice(4, 6), 16) / 255
  return 0.299 * r + 0.587 * g + 0.114 * b < 0.4
}
