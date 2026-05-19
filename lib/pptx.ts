import pptxgen from 'pptxgenjs'
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
  },
  business: {
    bgColor: 'F5F5F5',
    titleColor: '1A1A2E',
    bulletColor: '333344',
    accentColor: 'C0392B',
    footerColor: 'AAAAAA',
  },
  colorful: {
    bgColor: 'FAFAFA',
    titleColor: '6C3483',
    bulletColor: '2C3E50',
    accentColor: 'E74C3C',
    footerColor: 'AAAAAA',
  },
}

type ThemeConfig = typeof THEMES.academic & { fontFace?: string }

export async function generatePptx(data: PresentationData): Promise<Uint8Array> {
  const prs = new pptxgen()
  prs.layout = 'LAYOUT_WIDE'

  const base = THEMES[data.theme] ?? THEMES.academic
  const t = data.templateTheme
  const theme: ThemeConfig = t
    ? (() => {
        const isDark = isDarkColor(t.bgColor)
        return {
          // WPS 会忽略浅色文字颜色，强制改为浅色背景保证可读性
          bgColor: isDark ? 'F5F5F5' : t.bgColor,
          titleColor: isDark ? t.accentColor : t.titleColor,
          bulletColor: isDark ? '333333' : t.bodyColor,
          accentColor: t.accentColor,
          footerColor: 'AAAAAA',
          fontFace: t.fontName,
        }
      })()
    : { ...base }

  // ── 封面颜色块始终用强调色（深色），文字白色在色块上 ──

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
  // 封面中央色块
  cover.addShape(prs.ShapeType.rect, {
    x: 0.8, y: 1.8, w: SLIDE_W - 1.6, h: 3.8,
    fill: { color: theme.accentColor },
    line: { type: 'none' },
  })
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

    // 顶部色条
    s.addShape(prs.ShapeType.rect, {
      x: 0, y: 0, w: SLIDE_W, h: 0.22,
      fill: { color: theme.accentColor },
      line: { type: 'none' },
    })

    // 标题
    s.addText(slide.title, {
      x: 0.5, y: 0.28, w: SLIDE_W - 1.0, h: 1.0,
      fontSize: 28,
      bold: true,
      color: theme.titleColor,
      valign: 'middle',
      wrap: true,
      ...(theme.fontFace ? { fontFace: theme.fontFace } : {}),
    })

    // 标题下分割线
    s.addShape(prs.ShapeType.line, {
      x: 0.5, y: 1.35, w: SLIDE_W - 1.0, h: 0,
      line: { color: theme.accentColor, width: 1.5 },
    })

    // 要点
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

    // 页脚
    s.addText(data.title, {
      x: 0.5, y: SLIDE_H - 0.35, w: SLIDE_W - 1.0, h: 0.28,
      fontSize: 9,
      color: theme.footerColor,
      align: 'right',
    })
  }

  const result = await (prs.write({ outputType: 'arraybuffer' }) as Promise<ArrayBuffer>)
  return new Uint8Array(result)
}

// 判断颜色是否为深色（亮度 < 0.4）
function isDarkColor(hex: string): boolean {
  const h = hex.replace('#', '')
  const r = parseInt(h.slice(0, 2), 16) / 255
  const g = parseInt(h.slice(2, 4), 16) / 255
  const b = parseInt(h.slice(4, 6), 16) / 255
  const luminance = 0.299 * r + 0.587 * g + 0.114 * b
  return luminance < 0.4
}
