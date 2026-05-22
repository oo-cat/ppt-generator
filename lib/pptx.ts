import pptxgen from 'pptxgenjs'
import JSZip from 'jszip'
import { PresentationData, ThemeId } from '@/types'

const SLIDE_W = 13.33
const SLIDE_H = 7.5

type ThemeConfig = {
  bgColor: string
  titleColor: string
  bulletColor: string
  accentColor: string
  accent2Color: string
  footerColor: string
  dark: boolean
  fontFace?: string
  layout: 'classic' | 'chinese-blue' | 'chinese-ink' | 'chinese-elegant' | 'minimal-white' | 'minimal-particles'
}

const THEMES: Record<ThemeId, ThemeConfig> = {
  academic: {
    bgColor: 'FFFFFF', titleColor: '1F3864', bulletColor: '2E4057',
    accentColor: '2E75B6', accent2Color: '1F3864', footerColor: '999999',
    dark: false, layout: 'classic',
  },
  business: {
    bgColor: '1A1A2E', titleColor: 'E2E2E2', bulletColor: 'C8D0E0',
    accentColor: 'E94560', accent2Color: 'C8D0E0', footerColor: '777777',
    dark: true, layout: 'classic',
  },
  colorful: {
    bgColor: 'FAFAFA', titleColor: '6C3483', bulletColor: '2C3E50',
    accentColor: 'E74C3C', accent2Color: '6C3483', footerColor: 'AAAAAA',
    dark: false, layout: 'classic',
  },
  'chinese-blue': {
    bgColor: 'FFFFFF', titleColor: '1B3F5E', bulletColor: '2E4057',
    accentColor: '4A80A1', accent2Color: '2583BD', footerColor: '999999',
    dark: false, fontFace: '微软雅黑', layout: 'chinese-blue',
  },
  'chinese-ink': {
    bgColor: 'FAFAF8', titleColor: '1A1A1A', bulletColor: '333333',
    accentColor: '4472C4', accent2Color: 'ED7D31', footerColor: 'AAAAAA',
    dark: false, fontFace: '微软雅黑', layout: 'chinese-ink',
  },
  'chinese-elegant': {
    bgColor: 'FFFFFF', titleColor: '2C3A4A', bulletColor: '3D4F62',
    accentColor: '7E97BA', accent2Color: 'B2C1D6', footerColor: 'AAAAAA',
    dark: false, fontFace: '微软雅黑', layout: 'chinese-elegant',
  },
  'minimal-white': {
    bgColor: 'FFFFFF', titleColor: '1A1A1A', bulletColor: '333333',
    accentColor: '5B9BD5', accent2Color: 'ED7D31', footerColor: 'BBBBBB',
    dark: false, fontFace: '微软雅黑', layout: 'minimal-white',
  },
  'minimal-particles': {
    bgColor: '0D1B2A', titleColor: 'E8F4FD', bulletColor: 'C5DCF0',
    accentColor: '5B9BD5', accent2Color: '3A7DC9', footerColor: '556677',
    dark: true, fontFace: '微软雅黑', layout: 'minimal-particles',
  },
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
        accent2Color: t.accentColor,
        footerColor: 'AAAAAA',
        dark: isDarkColor(t.bgColor),
        fontFace: t.fontName,
        layout: 'classic',
      }
    : { ...base }

  const ff = (size?: number) => ({
    ...(theme.fontFace ? { fontFace: theme.fontFace } : {}),
    ...(size ? { fontSize: size } : {}),
  })

  switch (theme.layout) {
    case 'chinese-blue':
      buildChineseBlue(prs, data, theme, ff)
      break
    case 'chinese-ink':
      buildChineseInk(prs, data, theme, ff)
      break
    case 'chinese-elegant':
      buildChineseElegant(prs, data, theme, ff)
      break
    case 'minimal-white':
      buildMinimalWhite(prs, data, theme, ff)
      break
    case 'minimal-particles':
      buildMinimalParticles(prs, data, theme, ff)
      break
    default:
      buildClassic(prs, data, theme, ff)
  }

  const arraybuf = await (prs.write({ outputType: 'arraybuffer' }) as Promise<ArrayBuffer>)

  if (theme.dark) {
    return patchThemeTextColor(arraybuf, 'FFFFFF')
  }
  return new Uint8Array(arraybuf)
}

// ── Classic (academic / business / colorful) ────────────────────────────────

function buildClassic(
  prs: pptxgen,
  data: PresentationData,
  theme: ThemeConfig,
  ff: (s?: number) => object,
) {
  const cover = prs.addSlide()
  cover.background = { color: theme.bgColor }
  cover.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: SLIDE_W, h: 0.25, fill: { color: theme.accentColor }, line: { type: 'none' } })
  cover.addShape(prs.ShapeType.rect, { x: 0, y: SLIDE_H - 0.25, w: SLIDE_W, h: 0.25, fill: { color: theme.accentColor }, line: { type: 'none' } })
  cover.addShape(prs.ShapeType.rect, { x: 0.8, y: 1.8, w: SLIDE_W - 1.6, h: 3.8, fill: { color: theme.accentColor }, line: { type: 'none' } })
  cover.addText(data.title, { x: 1.0, y: 2.0, w: SLIDE_W - 2.0, h: 3.4, ...ff(36), bold: true, color: 'FFFFFF', align: 'center', valign: 'middle', wrap: true })

  for (const slide of data.slides) {
    const s = prs.addSlide()
    s.background = { color: theme.bgColor }
    s.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: SLIDE_W, h: 0.22, fill: { color: theme.accentColor }, line: { type: 'none' } })
    s.addText(slide.title, { x: 0.5, y: 0.28, w: SLIDE_W - 1.0, h: 1.0, ...ff(28), bold: true, color: theme.titleColor, valign: 'middle', wrap: true })
    s.addShape(prs.ShapeType.line, { x: 0.5, y: 1.35, w: SLIDE_W - 1.0, h: 0, line: { color: theme.accentColor, width: 1.5 } })
    if (slide.bullets && slide.bullets.length > 0) {
      s.addText(
        slide.bullets.map((b) => ({ text: `• ${b}`, options: { ...ff(20), color: theme.bulletColor, paraSpaceAfter: 10, breakLine: true } })),
        { x: 0.7, y: 1.5, w: SLIDE_W - 1.4, h: SLIDE_H - 2.1, valign: 'top', wrap: true },
      )
    }
    s.addText(data.title, { x: 0.5, y: SLIDE_H - 0.35, w: SLIDE_W - 1.0, h: 0.28, ...ff(9), color: theme.footerColor, align: 'right' })
  }
}

// ── Chinese Blue (中国风商务汇报 style) ──────────────────────────────────────

function buildChineseBlue(
  prs: pptxgen,
  data: PresentationData,
  theme: ThemeConfig,
  ff: (s?: number) => object,
) {
  // Cover: full-width accent band in the middle, title in white, subtitle below
  const cover = prs.addSlide()
  cover.background = { color: theme.bgColor }

  // Top thin bar
  cover.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: SLIDE_W, h: 0.12, fill: { color: theme.accentColor }, line: { type: 'none' } })
  // Left vertical accent strip
  cover.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 0.6, h: SLIDE_H, fill: { color: theme.accentColor }, line: { type: 'none' } })
  // Secondary vertical strip (lighter)
  cover.addShape(prs.ShapeType.rect, { x: 0.6, y: 0, w: 0.15, h: SLIDE_H, fill: { color: theme.accent2Color }, line: { type: 'none' } })
  // Bottom bar
  cover.addShape(prs.ShapeType.rect, { x: 0, y: SLIDE_H - 0.5, w: SLIDE_W, h: 0.5, fill: { color: theme.accentColor }, line: { type: 'none' } })

  // Title area
  cover.addText(data.title, {
    x: 1.2, y: 2.0, w: SLIDE_W - 2.0, h: 2.2,
    ...ff(40), bold: true, color: theme.titleColor, align: 'left', valign: 'middle', wrap: true,
  })
  // Decorative line under title
  cover.addShape(prs.ShapeType.line, { x: 1.2, y: 4.3, w: 5, h: 0, line: { color: theme.accentColor, width: 3 } })
  // Subtitle placeholder
  cover.addText('演示文稿 · Presentation', {
    x: 1.2, y: 4.5, w: 8, h: 0.5,
    ...ff(14), color: theme.accentColor, align: 'left',
  })

  for (const slide of data.slides) {
    const s = prs.addSlide()
    s.background = { color: theme.bgColor }

    // Left accent strip
    s.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 0.6, h: SLIDE_H, fill: { color: theme.accentColor }, line: { type: 'none' } })
    s.addShape(prs.ShapeType.rect, { x: 0.6, y: 0, w: 0.15, h: SLIDE_H, fill: { color: theme.accent2Color }, line: { type: 'none' } })

    // Title bar
    s.addShape(prs.ShapeType.rect, { x: 0.9, y: 0.3, w: SLIDE_W - 1.2, h: 1.0, fill: { color: theme.accentColor }, line: { type: 'none' }, rectRadius: 0.05 })
    s.addText(slide.title, {
      x: 1.1, y: 0.3, w: SLIDE_W - 1.6, h: 1.0,
      ...ff(26), bold: true, color: 'FFFFFF', valign: 'middle', wrap: true,
    })

    // Content area
    if (slide.bullets && slide.bullets.length > 0) {
      s.addText(
        slide.bullets.map((b, i) => ([
          { text: `${i + 1}`, options: { bold: true, color: theme.accentColor, ...ff(18) } },
          { text: `  ${b}`, options: { color: theme.bulletColor, ...ff(18), breakLine: true, paraSpaceAfter: 12 } },
        ])).flat(),
        { x: 0.9, y: 1.55, w: SLIDE_W - 1.3, h: SLIDE_H - 2.1, valign: 'top', wrap: true },
      )
    }

    // Footer
    s.addText(data.title, {
      x: 0.9, y: SLIDE_H - 0.35, w: SLIDE_W - 1.2, h: 0.28,
      ...ff(9), color: theme.footerColor, align: 'right',
    })
  }
}

// ── Chinese Ink (极简中国风 style) ───────────────────────────────────────────

function buildChineseInk(
  prs: pptxgen,
  data: PresentationData,
  theme: ThemeConfig,
  ff: (s?: number) => object,
) {
  // Cover: clean white, large bold title, accent underline
  const cover = prs.addSlide()
  cover.background = { color: theme.bgColor }

  // Top accent bar (thin)
  cover.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: SLIDE_W, h: 0.08, fill: { color: theme.accentColor }, line: { type: 'none' } })
  // Bottom accent bar
  cover.addShape(prs.ShapeType.rect, { x: 0, y: SLIDE_H - 0.08, w: SLIDE_W, h: 0.08, fill: { color: theme.accentColor }, line: { type: 'none' } })

  // Centered title block
  cover.addText(data.title, {
    x: 1.5, y: 1.8, w: SLIDE_W - 3.0, h: 2.8,
    ...ff(44), bold: true, color: theme.titleColor, align: 'center', valign: 'middle', wrap: true,
  })
  // Decorative double line
  cover.addShape(prs.ShapeType.line, { x: 4.0, y: 4.75, w: 5.33, h: 0, line: { color: theme.accentColor, width: 2 } })
  cover.addShape(prs.ShapeType.line, { x: 4.5, y: 5.0, w: 4.33, h: 0, line: { color: theme.accent2Color, width: 1 } })

  for (const slide of data.slides) {
    const s = prs.addSlide()
    s.background = { color: theme.bgColor }

    // Top thin accent line
    s.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: SLIDE_W, h: 0.06, fill: { color: theme.accentColor }, line: { type: 'none' } })

    // Title with left accent block
    s.addShape(prs.ShapeType.rect, { x: 0.5, y: 0.45, w: 0.15, h: 0.7, fill: { color: theme.accentColor }, line: { type: 'none' } })
    s.addText(slide.title, {
      x: 0.8, y: 0.4, w: SLIDE_W - 1.3, h: 0.8,
      ...ff(30), bold: true, color: theme.titleColor, valign: 'middle', wrap: true,
    })

    // Separator line
    s.addShape(prs.ShapeType.line, { x: 0.5, y: 1.35, w: SLIDE_W - 1.0, h: 0, line: { color: theme.accentColor, width: 1.5 } })

    // Bullets with dot markers
    if (slide.bullets && slide.bullets.length > 0) {
      s.addText(
        slide.bullets.map((b) => ([
          { text: '◆', options: { color: theme.accentColor, ...ff(12) } },
          { text: `  ${b}`, options: { color: theme.bulletColor, ...ff(19), breakLine: true, paraSpaceAfter: 14 } },
        ])).flat(),
        { x: 0.7, y: 1.5, w: SLIDE_W - 1.4, h: SLIDE_H - 2.0, valign: 'top', wrap: true },
      )
    }

    // Bottom line
    s.addShape(prs.ShapeType.line, { x: 0.5, y: SLIDE_H - 0.4, w: SLIDE_W - 1.0, h: 0, line: { color: theme.accentColor, width: 0.5 } })
    s.addText(data.title, {
      x: 0.5, y: SLIDE_H - 0.4, w: SLIDE_W - 1.0, h: 0.32,
      ...ff(9), color: theme.footerColor, align: 'right',
    })
  }
}

// ── Chinese Elegant (雅韵中国风 style) ───────────────────────────────────────

function buildChineseElegant(
  prs: pptxgen,
  data: PresentationData,
  theme: ThemeConfig,
  ff: (s?: number) => object,
) {
  // Cover: light, airy, elegant center composition
  const cover = prs.addSlide()
  cover.background = { color: theme.bgColor }

  // Top decorative border
  cover.addShape(prs.ShapeType.rect, { x: 0.4, y: 0.3, w: SLIDE_W - 0.8, h: 0.04, fill: { color: theme.accentColor }, line: { type: 'none' } })
  cover.addShape(prs.ShapeType.rect, { x: 0.4, y: 0.45, w: SLIDE_W - 0.8, h: 0.02, fill: { color: theme.accent2Color }, line: { type: 'none' } })
  // Bottom decorative border
  cover.addShape(prs.ShapeType.rect, { x: 0.4, y: SLIDE_H - 0.5, w: SLIDE_W - 0.8, h: 0.04, fill: { color: theme.accentColor }, line: { type: 'none' } })
  cover.addShape(prs.ShapeType.rect, { x: 0.4, y: SLIDE_H - 0.6, w: SLIDE_W - 0.8, h: 0.02, fill: { color: theme.accent2Color }, line: { type: 'none' } })

  // Centered accent diamond
  cover.addShape(prs.ShapeType.rect, { x: SLIDE_W / 2 - 0.12, y: 1.6, w: 0.24, h: 0.24, fill: { color: theme.accentColor }, line: { type: 'none' }, rotate: 45 })

  // Title
  cover.addText(data.title, {
    x: 1.2, y: 2.0, w: SLIDE_W - 2.4, h: 2.5,
    ...ff(42), bold: true, color: theme.titleColor, align: 'center', valign: 'middle', wrap: true,
  })
  // Separator line pair
  cover.addShape(prs.ShapeType.line, { x: 3.0, y: 4.65, w: 7.33, h: 0, line: { color: theme.accentColor, width: 1 } })
  cover.addShape(prs.ShapeType.line, { x: 4.0, y: 4.85, w: 5.33, h: 0, line: { color: theme.accent2Color, width: 0.5 } })

  for (const slide of data.slides) {
    const s = prs.addSlide()
    s.background = { color: theme.bgColor }

    // Top double border
    s.addShape(prs.ShapeType.rect, { x: 0.4, y: 0.15, w: SLIDE_W - 0.8, h: 0.04, fill: { color: theme.accentColor }, line: { type: 'none' } })
    s.addShape(prs.ShapeType.rect, { x: 0.4, y: 0.25, w: SLIDE_W - 0.8, h: 0.02, fill: { color: theme.accent2Color }, line: { type: 'none' } })

    // Title with ornamental marks
    s.addText('◇', { x: 0.4, y: 0.5, w: 0.5, h: 0.7, ...ff(16), color: theme.accentColor, valign: 'middle', align: 'center' })
    s.addText(slide.title, { x: 0.9, y: 0.5, w: SLIDE_W - 1.6, h: 0.7, ...ff(28), bold: true, color: theme.titleColor, valign: 'middle', wrap: true })
    s.addText('◇', { x: SLIDE_W - 0.9, y: 0.5, w: 0.5, h: 0.7, ...ff(16), color: theme.accentColor, valign: 'middle', align: 'center' })

    // Title separator
    s.addShape(prs.ShapeType.line, { x: 0.4, y: 1.3, w: SLIDE_W - 0.8, h: 0, line: { color: theme.accentColor, width: 1 } })

    if (slide.bullets && slide.bullets.length > 0) {
      s.addText(
        slide.bullets.map((b) => ([
          { text: '◈', options: { color: theme.accentColor, ...ff(13) } },
          { text: `  ${b}`, options: { color: theme.bulletColor, ...ff(19), breakLine: true, paraSpaceAfter: 15 } },
        ])).flat(),
        { x: 0.6, y: 1.45, w: SLIDE_W - 1.2, h: SLIDE_H - 2.0, valign: 'top', wrap: true },
      )
    }

    // Bottom double border
    s.addShape(prs.ShapeType.rect, { x: 0.4, y: SLIDE_H - 0.28, w: SLIDE_W - 0.8, h: 0.04, fill: { color: theme.accentColor }, line: { type: 'none' } })
    s.addText(data.title, { x: 0.5, y: SLIDE_H - 0.55, w: SLIDE_W - 1.0, h: 0.28, ...ff(9), color: theme.footerColor, align: 'center' })
  }
}

// ── Minimal White (极简白色PPT style) ────────────────────────────────────────

function buildMinimalWhite(
  prs: pptxgen,
  data: PresentationData,
  theme: ThemeConfig,
  ff: (s?: number) => object,
) {
  // Cover: ultra-clean, large title, small accent square
  const cover = prs.addSlide()
  cover.background = { color: theme.bgColor }

  // Accent square top-left
  cover.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 1.2, h: 1.2, fill: { color: theme.accentColor }, line: { type: 'none' } })
  // Thin bottom bar
  cover.addShape(prs.ShapeType.rect, { x: 0, y: SLIDE_H - 0.12, w: SLIDE_W, h: 0.12, fill: { color: theme.accentColor }, line: { type: 'none' } })

  // Title
  cover.addText(data.title, {
    x: 1.0, y: 1.8, w: SLIDE_W - 2.0, h: 2.8,
    ...ff(44), bold: true, color: theme.titleColor, align: 'left', valign: 'middle', wrap: true,
  })
  // Horizontal rule
  cover.addShape(prs.ShapeType.line, { x: 1.0, y: 4.75, w: 6.0, h: 0, line: { color: theme.accentColor, width: 2.5 } })
  // Tagline area
  cover.addText('PRESENTATION', {
    x: 1.0, y: 5.0, w: 6.0, h: 0.5,
    ...ff(12), color: theme.accentColor, align: 'left', bold: true,
  })

  for (const slide of data.slides) {
    const s = prs.addSlide()
    s.background = { color: theme.bgColor }

    // Top accent bar (thin)
    s.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: SLIDE_W, h: 0.08, fill: { color: theme.accentColor }, line: { type: 'none' } })

    // Title
    s.addText(slide.title, {
      x: 0.7, y: 0.2, w: SLIDE_W - 1.4, h: 1.0,
      ...ff(30), bold: true, color: theme.titleColor, valign: 'middle', wrap: true,
    })

    // Title underline (left-aligned, matches title width approx)
    s.addShape(prs.ShapeType.rect, { x: 0.7, y: 1.25, w: 3.5, h: 0.06, fill: { color: theme.accentColor }, line: { type: 'none' } })

    // Bullet content
    if (slide.bullets && slide.bullets.length > 0) {
      s.addText(
        slide.bullets.map((b) => ([
          { text: '—', options: { color: theme.accentColor, ...ff(17), bold: true } },
          { text: `  ${b}`, options: { color: theme.bulletColor, ...ff(18), breakLine: true, paraSpaceAfter: 13 } },
        ])).flat(),
        { x: 0.7, y: 1.4, w: SLIDE_W - 1.4, h: SLIDE_H - 1.95, valign: 'top', wrap: true },
      )
    }

    // Footer
    s.addShape(prs.ShapeType.line, { x: 0.7, y: SLIDE_H - 0.38, w: SLIDE_W - 1.4, h: 0, line: { color: 'DDDDDD', width: 0.5 } })
    s.addText(data.title, { x: 0.7, y: SLIDE_H - 0.38, w: SLIDE_W - 1.4, h: 0.3, ...ff(9), color: theme.footerColor, align: 'right' })
  }
}

// ── Minimal Particles (极简点线粒子 style, dark) ──────────────────────────────

function buildMinimalParticles(
  prs: pptxgen,
  data: PresentationData,
  theme: ThemeConfig,
  ff: (s?: number) => object,
) {
  // Cover: dark navy, glowing accent elements, centered title
  const cover = prs.addSlide()
  cover.background = { color: theme.bgColor }

  // Horizontal glow lines
  cover.addShape(prs.ShapeType.rect, { x: 0, y: 2.4, w: SLIDE_W, h: 0.03, fill: { color: theme.accentColor }, line: { type: 'none' } })
  cover.addShape(prs.ShapeType.rect, { x: 0, y: 5.05, w: SLIDE_W, h: 0.03, fill: { color: theme.accentColor }, line: { type: 'none' } })

  // Corner accent squares
  cover.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 0.5, h: 0.5, fill: { color: theme.accentColor }, line: { type: 'none' } })
  cover.addShape(prs.ShapeType.rect, { x: SLIDE_W - 0.5, y: SLIDE_H - 0.5, w: 0.5, h: 0.5, fill: { color: theme.accentColor }, line: { type: 'none' } })
  // Small dot accents
  cover.addShape(prs.ShapeType.ellipse, { x: 0.8, y: 2.3, w: 0.16, h: 0.16, fill: { color: theme.accentColor }, line: { type: 'none' } })
  cover.addShape(prs.ShapeType.ellipse, { x: SLIDE_W - 1.0, y: 5.0, w: 0.16, h: 0.16, fill: { color: theme.accentColor }, line: { type: 'none' } })

  // Title
  cover.addText(data.title, {
    x: 1.0, y: 2.5, w: SLIDE_W - 2.0, h: 2.5,
    ...ff(42), bold: true, color: theme.titleColor, align: 'center', valign: 'middle', wrap: true,
  })
  // Sub accent line
  cover.addShape(prs.ShapeType.rect, { x: SLIDE_W / 2 - 1.5, y: 5.1, w: 3.0, h: 0.04, fill: { color: theme.accent2Color }, line: { type: 'none' } })

  for (const slide of data.slides) {
    const s = prs.addSlide()
    s.background = { color: theme.bgColor }

    // Top accent line
    s.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: SLIDE_W, h: 0.06, fill: { color: theme.accentColor }, line: { type: 'none' } })
    // Title left accent
    s.addShape(prs.ShapeType.rect, { x: 0.5, y: 0.3, w: 0.08, h: 0.75, fill: { color: theme.accentColor }, line: { type: 'none' } })
    s.addShape(prs.ShapeType.ellipse, { x: 0.46, y: 0.26, w: 0.16, h: 0.16, fill: { color: theme.accentColor }, line: { type: 'none' } })
    s.addShape(prs.ShapeType.ellipse, { x: 0.46, y: 0.89, w: 0.16, h: 0.16, fill: { color: theme.accentColor }, line: { type: 'none' } })

    // Title
    s.addText(slide.title, {
      x: 0.85, y: 0.3, w: SLIDE_W - 1.5, h: 0.85,
      ...ff(28), bold: true, color: theme.titleColor, valign: 'middle', wrap: true,
    })

    // Separator with dots
    s.addShape(prs.ShapeType.line, { x: 0.5, y: 1.25, w: SLIDE_W - 1.0, h: 0, line: { color: theme.accentColor, width: 0.75 } })
    s.addShape(prs.ShapeType.ellipse, { x: 0.46, y: 1.22, w: 0.08, h: 0.08, fill: { color: theme.accentColor }, line: { type: 'none' } })

    // Bullets
    if (slide.bullets && slide.bullets.length > 0) {
      s.addText(
        slide.bullets.map((b) => ([
          { text: '▸', options: { color: theme.accentColor, ...ff(14) } },
          { text: `  ${b}`, options: { color: theme.bulletColor, ...ff(18), breakLine: true, paraSpaceAfter: 13 } },
        ])).flat(),
        { x: 0.7, y: 1.4, w: SLIDE_W - 1.4, h: SLIDE_H - 1.95, valign: 'top', wrap: true },
      )
    }

    // Bottom corner accent
    s.addShape(prs.ShapeType.rect, { x: SLIDE_W - 0.4, y: SLIDE_H - 0.4, w: 0.4, h: 0.4, fill: { color: theme.accentColor }, line: { type: 'none' } })
    s.addText(data.title, { x: 0.5, y: SLIDE_H - 0.38, w: SLIDE_W - 1.3, h: 0.3, ...ff(9), color: theme.footerColor, align: 'right' })
  }
}

// ── Helpers ──────────────────────────────────────────────────────────────────

async function patchThemeTextColor(buf: ArrayBuffer, textColor: string): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(buf)
  const themeFiles = zip.file(/ppt\/theme\/theme\d*\.xml/)
  for (const file of themeFiles) {
    let xml = await file.async('text')
    xml = xml.replace(/<a:dk1>[\s\S]*?<\/a:dk1>/, `<a:dk1><a:srgbClr val="${textColor}"/></a:dk1>`)
    zip.file(file.name, xml)
  }
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

function isDarkColor(hex: string): boolean {
  const h = hex.replace('#', '')
  const r = parseInt(h.slice(0, 2), 16) / 255
  const g = parseInt(h.slice(2, 4), 16) / 255
  const b = parseInt(h.slice(4, 6), 16) / 255
  return 0.299 * r + 0.587 * g + 0.114 * b < 0.4
}
