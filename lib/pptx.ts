import pptxgen from 'pptxgenjs'
import { PresentationData } from '@/types'

const THEMES = {
  academic: {
    background: { color: 'FFFFFF' },
    titleColor: '1a1a2e',
    titleFontSize: 32,
    bulletColor: '16213e',
    bulletFontSize: 18,
    accentColor: '0f3460',
    footerColor: '888888',
  },
  business: {
    background: { color: '1a1a2e' },
    titleColor: 'e2e2e2',
    titleFontSize: 32,
    bulletColor: 'c0c0c0',
    bulletFontSize: 18,
    accentColor: 'e94560',
    footerColor: '666666',
  },
  colorful: {
    background: { color: 'f8f9fa' },
    titleColor: '6c3483',
    titleFontSize: 32,
    bulletColor: '2c3e50',
    bulletFontSize: 18,
    accentColor: 'e74c3c',
    footerColor: '999999',
  },
}

type ThemeConfig = {
  background: { color: string }
  titleColor: string
  titleFontSize: number
  bulletColor: string
  bulletFontSize: number
  accentColor: string
  footerColor: string
  fontFace?: string
}

export async function generatePptx(data: PresentationData): Promise<Uint8Array> {
  const prs = new pptxgen()

  const base = THEMES[data.theme] ?? THEMES.academic
  const t = data.templateTheme
  const theme: ThemeConfig = t
    ? {
        background: { color: t.bgColor },
        titleColor: t.titleColor,
        titleFontSize: 32,
        bulletColor: t.bodyColor,
        bulletFontSize: 18,
        accentColor: t.accentColor,
        footerColor: 'aaaaaa',
        fontFace: t.fontName,
      }
    : base

  prs.layout = 'LAYOUT_WIDE'

  // 封面
  const coverSlide = prs.addSlide()
  coverSlide.background = theme.background
  coverSlide.addText(data.title, {
    x: '10%',
    y: '35%',
    w: '80%',
    h: '20%',
    fontSize: 40,
    bold: true,
    color: theme.titleColor,
    align: 'center',
    valign: 'middle',
    ...(theme.fontFace ? { fontFace: theme.fontFace } : {}),
  })

  // 内容页
  for (const slide of data.slides) {
    const s = prs.addSlide()
    s.background = theme.background

    s.addShape(prs.ShapeType.rect, {
      x: 0,
      y: 0,
      w: '100%',
      h: 0.12,
      fill: { color: theme.accentColor },
      line: { type: 'none' },
    })

    s.addText(slide.title, {
      x: '5%',
      y: '8%',
      w: '90%',
      h: '15%',
      fontSize: theme.titleFontSize,
      bold: true,
      color: theme.titleColor,
      valign: 'middle',
      ...(theme.fontFace ? { fontFace: theme.fontFace } : {}),
    })

    if (slide.bullets && slide.bullets.length > 0) {
      const bulletItems = slide.bullets.map((b) => ({
        text: b,
        options: {
          bullet: { type: 'bullet' as const },
          fontSize: theme.bulletFontSize,
          color: theme.bulletColor,
          paraSpaceAfter: 8,
          ...(theme.fontFace ? { fontFace: theme.fontFace } : {}),
        },
      }))
      s.addText(bulletItems, {
        x: '5%',
        y: '28%',
        w: '90%',
        h: '65%',
        valign: 'top',
      })
    }

    s.addText(data.title, {
      x: '5%',
      y: '93%',
      w: '90%',
      h: '5%',
      fontSize: 10,
      color: theme.footerColor,
      align: 'right',
    })
  }

  const result = await (prs.write({ outputType: 'arraybuffer' }) as Promise<ArrayBuffer>)
  return new Uint8Array(result)
}
