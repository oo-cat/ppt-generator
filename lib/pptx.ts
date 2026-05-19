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
    fontFace: undefined as string | undefined,
  },
  business: {
    background: { color: '1a1a2e' },
    titleColor: 'e2e2e2',
    titleFontSize: 32,
    bulletColor: 'c0c0c0',
    bulletFontSize: 18,
    accentColor: 'e94560',
    footerColor: '666666',
    fontFace: undefined as string | undefined,
  },
  colorful: {
    background: { color: 'f8f9fa' },
    titleColor: '6c3483',
    titleFontSize: 32,
    bulletColor: '2c3e50',
    bulletFontSize: 18,
    accentColor: 'e74c3c',
    footerColor: '999999',
    fontFace: undefined as string | undefined,
  },
}

export async function generatePptx(data: PresentationData): Promise<Buffer> {
  const prs = new pptxgen()

  // 如果有模板主题，用模板配色覆盖默认主题
  const base = THEMES[data.theme] ?? THEMES.academic
  const t = data.templateTheme
  const theme = t ? {
    background: { color: t.bgColor },
    titleColor: t.titleColor,
    titleFontSize: 32,
    bulletColor: t.bodyColor,
    bulletFontSize: 18,
    accentColor: t.accentColor,
    footerColor: 'aaaaaa',
    fontFace: t.fontName !== '微软雅黑' ? t.fontName : undefined,
  } : base

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
    fontFace: theme.fontFace,
  })

  // 内容页
  for (const slide of data.slides) {
    const s = prs.addSlide()
    s.background = theme.background

    // 顶部色条
    s.addShape(prs.ShapeType.rect, {
      x: 0,
      y: 0,
      w: '100%',
      h: 0.12,
      fill: { color: theme.accentColor },
      line: { type: 'none' },
    })

    // 标题
    s.addText(slide.title, {
      x: '5%',
      y: '8%',
      w: '90%',
      h: '15%',
      fontSize: theme.titleFontSize,
      bold: true,
      color: theme.titleColor,
      valign: 'middle',
      fontFace: theme.fontFace,
    })

    // 要点
    if (slide.bullets && slide.bullets.length > 0) {
      const bulletItems = slide.bullets.map((b) => ({
        text: b,
        options: { bullet: { type: 'bullet' as const }, fontSize: theme.bulletFontSize, color: theme.bulletColor, paraSpaceAfter: 8 },
      }))
      s.addText(bulletItems, {
        x: '5%',
        y: '28%',
        w: '90%',
        h: '65%',
        valign: 'top',
      })
    }

    // 底部页脚
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

  return prs.write({ outputType: 'nodebuffer' }) as unknown as Promise<Buffer>
}
