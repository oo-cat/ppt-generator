import pptxgen from 'pptxgenjs'
import { PresentationData } from '@/types'

// LAYOUT_WIDE = 13.33 x 7.5 inches
const SLIDE_W = 13.33
const SLIDE_H = 7.5

const THEMES = {
  academic: {
    bgColor: 'FFFFFF',
    titleColor: '1a1a2e',
    bulletColor: '16213e',
    accentColor: '0f3460',
    footerColor: '888888',
  },
  business: {
    bgColor: '1a1a2e',
    titleColor: 'e2e2e2',
    bulletColor: 'c8d0e0',
    accentColor: 'e94560',
    footerColor: '777777',
  },
  colorful: {
    bgColor: 'f8f9fa',
    titleColor: '6c3483',
    bulletColor: '2c3e50',
    accentColor: 'e74c3c',
    footerColor: '999999',
  },
}

type ThemeConfig = typeof THEMES.academic & { fontFace?: string }

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
        footerColor: 'aaaaaa',
        fontFace: t.fontName,
      }
    : { ...base }

  // ── 封面 ────────────────────────────────────────────────
  const cover = prs.addSlide()
  cover.background = { color: theme.bgColor }

  // 顶部装饰条
  cover.addShape(prs.ShapeType.rect, {
    x: 0, y: 0, w: SLIDE_W, h: 0.18,
    fill: { color: theme.accentColor },
    line: { type: 'none' },
  })
  // 底部装饰条
  cover.addShape(prs.ShapeType.rect, {
    x: 0, y: SLIDE_H - 0.18, w: SLIDE_W, h: 0.18,
    fill: { color: theme.accentColor },
    line: { type: 'none' },
  })

  cover.addText(data.title, {
    x: 1.2, y: 2.2, w: SLIDE_W - 2.4, h: 2.0,
    fontSize: 36,
    bold: true,
    color: theme.titleColor,
    align: 'center',
    valign: 'middle',
    wrap: true,
    ...(theme.fontFace ? { fontFace: theme.fontFace } : {}),
  })

  // ── 内容页 ───────────────────────────────────────────────
  for (const slide of data.slides) {
    const s = prs.addSlide()
    s.background = { color: theme.bgColor }

    // 顶部色条
    s.addShape(prs.ShapeType.rect, {
      x: 0, y: 0, w: SLIDE_W, h: 0.18,
      fill: { color: theme.accentColor },
      line: { type: 'none' },
    })

    // 标题区域（顶部色条下方）
    s.addText(slide.title, {
      x: 0.5, y: 0.25, w: SLIDE_W - 1.0, h: 1.1,
      fontSize: 28,
      bold: true,
      color: theme.titleColor,
      valign: 'middle',
      wrap: true,
      ...(theme.fontFace ? { fontFace: theme.fontFace } : {}),
    })

    // 标题下分割线
    s.addShape(prs.ShapeType.line, {
      x: 0.5, y: 1.4, w: SLIDE_W - 1.0, h: 0,
      line: { color: theme.accentColor, width: 1.5 },
    })

    // 要点内容区
    if (slide.bullets && slide.bullets.length > 0) {
      const bulletItems = slide.bullets.map((b) => ({
        text: b,
        options: {
          bullet: { type: 'bullet' as const },
          fontSize: 20,
          color: theme.bulletColor,
          paraSpaceAfter: 10,
          ...(theme.fontFace ? { fontFace: theme.fontFace } : {}),
        },
      }))
      s.addText(bulletItems, {
        x: 0.7, y: 1.55, w: SLIDE_W - 1.4, h: SLIDE_H - 2.1,
        valign: 'top',
        wrap: true,
      })
    }

    // 底部页脚
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
