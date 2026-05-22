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
  darkColor: string
  footerColor: string
  dark: boolean
  fontFace?: string
  layout: 'classic' | 'chinese-blue' | 'chinese-ink' | 'chinese-elegant' | 'minimal-white' | 'minimal-particles'
}

const THEMES: Record<ThemeId, ThemeConfig> = {
  academic: {
    bgColor: 'FFFFFF', titleColor: '1F3864', bulletColor: '2E4057',
    accentColor: '2E75B6', accent2Color: '1F3864', darkColor: '1F3864',
    footerColor: '999999', dark: false, layout: 'classic',
  },
  business: {
    bgColor: '1A1A2E', titleColor: 'E2E2E2', bulletColor: 'C8D0E0',
    accentColor: 'E94560', accent2Color: 'C8D0E0', darkColor: '0F0F1A',
    footerColor: '777777', dark: true, layout: 'classic',
  },
  colorful: {
    bgColor: 'FAFAFA', titleColor: '6C3483', bulletColor: '2C3E50',
    accentColor: 'E74C3C', accent2Color: '6C3483', darkColor: '4A235A',
    footerColor: 'AAAAAA', dark: false, layout: 'classic',
  },
  'chinese-blue': {
    bgColor: 'F4F8FC', titleColor: '1B3F5E', bulletColor: '2E4A60',
    accentColor: '4A80A1', accent2Color: '2583BD', darkColor: '2D5F7F',
    footerColor: 'AABBCC', dark: false, fontFace: '微软雅黑', layout: 'chinese-blue',
  },
  'chinese-ink': {
    bgColor: 'FFFFFF', titleColor: '1A1A1A', bulletColor: '333333',
    accentColor: '4472C4', accent2Color: 'ED7D31', darkColor: '2B4B8A',
    footerColor: 'BBBBBB', dark: false, fontFace: '微软雅黑', layout: 'chinese-ink',
  },
  'chinese-elegant': {
    bgColor: 'FDFCF8', titleColor: '2C3A4A', bulletColor: '3D4F62',
    accentColor: '7E97BA', accent2Color: 'B2C1D6', darkColor: '4A6280',
    footerColor: 'AAAAAA', dark: false, fontFace: '微软雅黑', layout: 'chinese-elegant',
  },
  'minimal-white': {
    bgColor: 'FFFFFF', titleColor: '1A1A1A', bulletColor: '333333',
    accentColor: '5B9BD5', accent2Color: 'ED7D31', darkColor: '2D6BA0',
    footerColor: 'BBBBBB', dark: false, fontFace: '微软雅黑', layout: 'minimal-white',
  },
  'minimal-particles': {
    bgColor: '0D1B2A', titleColor: 'E8F4FD', bulletColor: 'C5DCF0',
    accentColor: '5B9BD5', accent2Color: '3A7DC9', darkColor: '071018',
    footerColor: '445566', dark: true, fontFace: '微软雅黑', layout: 'minimal-particles',
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
        darkColor: t.accentColor,
        footerColor: 'AAAAAA',
        dark: isDarkColor(t.bgColor),
        fontFace: t.fontName,
        layout: 'classic',
      }
    : { ...base }

  // ff() returns font+size options for addText
  const ff = (size?: number): object => ({
    ...(theme.fontFace ? { fontFace: theme.fontFace } : {}),
    ...(size ? { fontSize: size } : {}),
  })

  switch (theme.layout) {
    case 'chinese-blue':      buildChineseBlue(prs, data, theme, ff); break
    case 'chinese-ink':       buildChineseInk(prs, data, theme, ff); break
    case 'chinese-elegant':   buildChineseElegant(prs, data, theme, ff); break
    case 'minimal-white':     buildMinimalWhite(prs, data, theme, ff); break
    case 'minimal-particles': buildMinimalParticles(prs, data, theme, ff); break
    default:                  buildClassic(prs, data, theme, ff)
  }

  const arraybuf = await (prs.write({ outputType: 'arraybuffer' }) as Promise<ArrayBuffer>)
  if (theme.dark) return patchThemeTextColor(arraybuf, 'FFFFFF')
  return new Uint8Array(arraybuf)
}

// ── Classic ──────────────────────────────────────────────────────────────────

function buildClassic(prs: pptxgen, data: PresentationData, theme: ThemeConfig, ff: (s?: number) => object) {
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
    if (slide.bullets?.length) {
      s.addText(slide.bullets.map((b) => ({ text: `• ${b}`, options: { ...ff(20), color: theme.bulletColor, paraSpaceAfter: 10, breakLine: true } })),
        { x: 0.7, y: 1.5, w: SLIDE_W - 1.4, h: SLIDE_H - 2.1, valign: 'top', wrap: true })
    }
    s.addText(data.title, { x: 0.5, y: SLIDE_H - 0.35, w: SLIDE_W - 1.0, h: 0.28, ...ff(9), color: theme.footerColor, align: 'right' })
  }
}

// ── Chinese Blue (中国风商务汇报 style) ──────────────────────────────────────
// 封面：左侧大色块面板 + 装饰性空心圆跨越边界 + 白色标题
// 内容：彩色全宽页眉 + 圆环装饰 + 编号圆圈要点

function buildChineseBlue(prs: pptxgen, data: PresentationData, theme: ThemeConfig, ff: (s?: number) => object) {
  const PANEL_W = 5.0

  // ── 封面 ──
  const cover = prs.addSlide()
  cover.background = { color: theme.bgColor }

  // 左侧主色面板
  cover.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: PANEL_W, h: SLIDE_H, fill: { color: theme.accentColor }, line: { type: 'none' } })
  // 左侧面板底部深色条
  cover.addShape(prs.ShapeType.rect, { x: 0, y: SLIDE_H - 0.55, w: PANEL_W, h: 0.55, fill: { color: theme.darkColor }, line: { type: 'none' } })

  // 大装饰圆（空心效果 = 大圆 + 内层白色小圆）
  const cirX = PANEL_W - 0.8
  const cirY = 1.0
  const cirR = 2.5
  cover.addShape(prs.ShapeType.ellipse, { x: cirX, y: cirY, w: cirR * 2, h: cirR * 2, fill: { color: '5C9FC0' }, line: { type: 'none' } })
  cover.addShape(prs.ShapeType.ellipse, { x: cirX + 0.3, y: cirY + 0.3, w: (cirR - 0.3) * 2, h: (cirR - 0.3) * 2, fill: { color: theme.bgColor }, line: { type: 'none' } })

  // 右侧小装饰块
  cover.addShape(prs.ShapeType.rect, { x: SLIDE_W - 0.65, y: 0, w: 0.65, h: 0.65, fill: { color: theme.accent2Color }, line: { type: 'none' } })
  cover.addShape(prs.ShapeType.rect, { x: SLIDE_W - 1.1, y: 0.7, w: 0.35, h: 0.35, fill: { color: '9BBDD5' }, line: { type: 'none' } })
  cover.addShape(prs.ShapeType.rect, { x: SLIDE_W - 0.4, y: SLIDE_H - 0.4, w: 0.4, h: 0.4, fill: { color: theme.darkColor }, line: { type: 'none' } })

  // 面板上的标题
  cover.addText(data.title, { x: 0.4, y: 1.4, w: PANEL_W - 0.6, h: 3.3, ...ff(34), bold: true, color: 'FFFFFF', align: 'left', valign: 'middle', wrap: true })
  // 装饰线 + 副标题
  cover.addShape(prs.ShapeType.rect, { x: 0.4, y: 5.0, w: 2.8, h: 0.06, fill: { color: 'FFFFFF' }, line: { type: 'none' } })
  cover.addShape(prs.ShapeType.rect, { x: 0.4, y: 5.15, w: 1.8, h: 0.04, fill: { color: 'A8CCE0' }, line: { type: 'none' } })
  cover.addText('BUSINESS  PRESENTATION', { x: 0.4, y: 5.38, w: 4.4, h: 0.42, ...ff(10), color: 'C5DEF0', align: 'left' })

  // ── 内容页 ──
  for (const slide of data.slides) {
    const s = prs.addSlide()
    s.background = { color: theme.bgColor }

    // 全宽页眉色块
    s.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: SLIDE_W, h: 1.45, fill: { color: theme.accentColor }, line: { type: 'none' } })
    // 顶部深色细条
    s.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: SLIDE_W, h: 0.1, fill: { color: theme.darkColor }, line: { type: 'none' } })
    // 页眉左侧装饰圆环
    s.addShape(prs.ShapeType.ellipse, { x: 0.2, y: 0.18, w: 1.1, h: 1.1, fill: { color: '5C9FC0' }, line: { type: 'none' } })
    s.addShape(prs.ShapeType.ellipse, { x: 0.45, y: 0.43, w: 0.6, h: 0.6, fill: { color: theme.accentColor }, line: { type: 'none' } })

    // 标题
    s.addText(slide.title, { x: 1.55, y: 0.1, w: SLIDE_W - 2.25, h: 1.25, ...ff(24), bold: true, color: 'FFFFFF', valign: 'middle', wrap: true })

    // 右侧深色细竖条
    s.addShape(prs.ShapeType.rect, { x: SLIDE_W - 0.28, y: 0, w: 0.28, h: SLIDE_H, fill: { color: theme.darkColor }, line: { type: 'none' } })

    // 底部装饰带
    s.addShape(prs.ShapeType.rect, { x: 0, y: SLIDE_H - 0.32, w: SLIDE_W, h: 0.32, fill: { color: 'EAF2F8' }, line: { type: 'none' } })
    s.addText(data.title, { x: 0.5, y: SLIDE_H - 0.32, w: SLIDE_W - 1.1, h: 0.28, ...ff(9), color: theme.footerColor, align: 'right' })

    // 编号圆圈要点
    if (slide.bullets?.length) {
      const bullets = slide.bullets.slice(0, 6)
      const startY = 1.6
      const availH = SLIDE_H - startY - 0.38
      const itemH = availH / bullets.length

      for (let i = 0; i < bullets.length; i++) {
        const iy = startY + i * itemH
        const circleY = iy + itemH / 2 - 0.3

        // 编号圆
        s.addShape(prs.ShapeType.ellipse, { x: 0.32, y: circleY, w: 0.58, h: 0.58, fill: { color: theme.accentColor }, line: { type: 'none' } })
        s.addText(String(i + 1).padStart(2, '0'), { x: 0.32, y: circleY, w: 0.58, h: 0.58, ...ff(11), bold: true, color: 'FFFFFF', align: 'center', valign: 'middle' })

        // 分割线（最后一项不加）
        if (i < bullets.length - 1) {
          s.addShape(prs.ShapeType.line, { x: 1.1, y: iy + itemH, w: SLIDE_W - 1.65, h: 0, line: { color: 'D5E8F5', width: 0.5 } })
        }

        // 文字
        s.addText(bullets[i], { x: 1.1, y: iy, w: SLIDE_W - 1.7, h: itemH, ...ff(18), color: theme.bulletColor, valign: 'middle', wrap: true })
      }
    }
  }
}

// ── Chinese Ink (极简中国风 style) ────────────────────────────────────────────
// 封面：顶部大色块（占40%高度）+ 标题在下方 + 左对齐布局 + 橙色点缀
// 内容：左侧竖线装饰 + 大章节序号 + 清爽要点列表

function buildChineseInk(prs: pptxgen, data: PresentationData, theme: ThemeConfig, ff: (s?: number) => object) {
  const TOP_H = 2.8

  // ── 封面 ──
  const cover = prs.addSlide()
  cover.background = { color: theme.bgColor }

  // 顶部大色块
  cover.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: SLIDE_W, h: TOP_H, fill: { color: theme.accentColor }, line: { type: 'none' } })
  // 顶部色块底部橙色细线
  cover.addShape(prs.ShapeType.rect, { x: 0, y: TOP_H, w: SLIDE_W, h: 0.1, fill: { color: theme.accent2Color }, line: { type: 'none' } })

  // 顶部色块内装饰：大空心圆
  cover.addShape(prs.ShapeType.ellipse, { x: 9.5, y: -1.0, w: 4.5, h: 4.5, fill: { color: '5580C5' }, line: { type: 'none' } })
  cover.addShape(prs.ShapeType.ellipse, { x: 9.9, y: -0.6, w: 3.7, h: 3.7, fill: { color: theme.accentColor }, line: { type: 'none' } })
  // 色块内小方块
  cover.addShape(prs.ShapeType.rect, { x: 0.5, y: 0.4, w: 0.08, h: 1.6, fill: { color: 'FFFFFF' }, line: { type: 'none' } })
  cover.addShape(prs.ShapeType.rect, { x: 0.72, y: 0.4, w: 0.04, h: 1.2, fill: { color: 'A0BEE0' }, line: { type: 'none' } })

  // 顶部色块内副标题（英文，小字）
  cover.addText('PRESENTATION', { x: 1.0, y: 0.5, w: 6.0, h: 0.5, ...ff(11), color: 'C8DCF8', align: 'left' })
  // 顶部色块内大标题（白色）
  cover.addText(data.title, { x: 1.0, y: 1.0, w: 8.5, h: 1.6, ...ff(36), bold: true, color: 'FFFFFF', align: 'left', valign: 'middle', wrap: true })

  // 下半部分：装饰 + 补充信息
  cover.addShape(prs.ShapeType.rect, { x: 0.5, y: TOP_H + 0.5, w: 0.06, h: 2.5, fill: { color: theme.accent2Color }, line: { type: 'none' } })
  cover.addText('在这里输入副标题或日期信息', { x: 0.9, y: TOP_H + 0.7, w: 9.0, h: 0.6, ...ff(15), color: '888888', align: 'left' })

  // 右下角装饰方块群
  cover.addShape(prs.ShapeType.rect, { x: SLIDE_W - 1.5, y: TOP_H + 0.3, w: 1.5, h: 0.5, fill: { color: 'EEF3FB' }, line: { type: 'none' } })
  cover.addShape(prs.ShapeType.rect, { x: SLIDE_W - 1.2, y: TOP_H + 1.0, w: 1.2, h: 0.5, fill: { color: 'DDE8F8' }, line: { type: 'none' } })
  cover.addShape(prs.ShapeType.rect, { x: SLIDE_W - 0.8, y: TOP_H + 1.7, w: 0.8, h: 0.5, fill: { color: theme.accentColor }, line: { type: 'none' } })
  cover.addShape(prs.ShapeType.rect, { x: SLIDE_W - 0.8, y: TOP_H + 1.7, w: 0.8, h: 0.5, fill: { color: theme.accentColor }, line: { type: 'none' } })

  // ── 内容页 ──
  for (let si = 0; si < data.slides.length; si++) {
    const slide = data.slides[si]
    const s = prs.addSlide()
    s.background = { color: theme.bgColor }

    // 顶部细蓝条
    s.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: SLIDE_W, h: 0.08, fill: { color: theme.accentColor }, line: { type: 'none' } })

    // 右上角大章节序号（装饰性）
    const numStr = String(si + 1).padStart(2, '0')
    s.addText(numStr, { x: SLIDE_W - 2.8, y: 0.15, w: 2.5, h: 1.8, ...ff(88), bold: true, color: 'EEF3FA', align: 'right', valign: 'top' })

    // 标题区域：左侧双竖线 + 标题文字
    s.addShape(prs.ShapeType.rect, { x: 0.5, y: 0.3, w: 0.1, h: 0.85, fill: { color: theme.accentColor }, line: { type: 'none' } })
    s.addShape(prs.ShapeType.rect, { x: 0.72, y: 0.3, w: 0.05, h: 0.65, fill: { color: theme.accent2Color }, line: { type: 'none' } })
    s.addText(slide.title, { x: 0.95, y: 0.25, w: SLIDE_W - 3.5, h: 0.95, ...ff(28), bold: true, color: theme.titleColor, valign: 'middle', wrap: true })

    // 橙色下划线
    s.addShape(prs.ShapeType.rect, { x: 0.5, y: 1.28, w: 4.5, h: 0.06, fill: { color: theme.accentColor }, line: { type: 'none' } })
    s.addShape(prs.ShapeType.rect, { x: 0.5, y: 1.39, w: 2.5, h: 0.04, fill: { color: theme.accent2Color }, line: { type: 'none' } })

    // 内容要点
    if (slide.bullets?.length) {
      const bullets = slide.bullets.slice(0, 7)
      const startY = 1.6
      const availH = SLIDE_H - startY - 0.5
      const itemH = Math.min(availH / bullets.length, 0.88)

      for (let i = 0; i < bullets.length; i++) {
        const iy = startY + i * itemH
        // 橙色方块标记
        s.addShape(prs.ShapeType.rect, { x: 0.5, y: iy + itemH / 2 - 0.1, w: 0.2, h: 0.2, fill: { color: theme.accentColor }, line: { type: 'none' } })
        // 蓝色竖线
        s.addShape(prs.ShapeType.rect, { x: 0.84, y: iy + 0.08, w: 0.04, h: itemH - 0.16, fill: { color: 'D8E5F5' }, line: { type: 'none' } })
        s.addText(bullets[i], { x: 1.05, y: iy, w: SLIDE_W - 1.6, h: itemH, ...ff(18), color: theme.bulletColor, valign: 'middle', wrap: true })
      }
    }

    // 底部
    s.addShape(prs.ShapeType.rect, { x: 0, y: SLIDE_H - 0.08, w: SLIDE_W, h: 0.08, fill: { color: theme.accentColor }, line: { type: 'none' } })
    s.addText(data.title, { x: 0.5, y: SLIDE_H - 0.38, w: SLIDE_W - 1.0, h: 0.3, ...ff(9), color: theme.footerColor, align: 'right' })
  }
}

// ── Chinese Elegant (雅韵中国风 style) ────────────────────────────────────────
// 封面：双边框镶边 + 菱形角饰 + 居中典雅布局
// 内容：四角装饰框 + 侧边花纹 + 精致分割线

function buildChineseElegant(prs: pptxgen, data: PresentationData, theme: ThemeConfig, ff: (s?: number) => object) {
  const BORDER = 0.35 // 边框距边缘距离

  function addCornerDiamond(s: pptxgen.Slide, x: number, y: number, size: number, color: string) {
    s.addShape(prs.ShapeType.rect, { x: x - size / 2, y: y - size / 2, w: size, h: size, fill: { color }, line: { type: 'none' }, rotate: 45 })
  }

  // ── 封面 ──
  const cover = prs.addSlide()
  cover.background = { color: theme.bgColor }

  // 外边框（深色粗线）
  cover.addShape(prs.ShapeType.rect, { x: BORDER, y: BORDER, w: SLIDE_W - BORDER * 2, h: 0.05, fill: { color: theme.accentColor }, line: { type: 'none' } })
  cover.addShape(prs.ShapeType.rect, { x: BORDER, y: SLIDE_H - BORDER - 0.05, w: SLIDE_W - BORDER * 2, h: 0.05, fill: { color: theme.accentColor }, line: { type: 'none' } })
  cover.addShape(prs.ShapeType.rect, { x: BORDER, y: BORDER, w: 0.05, h: SLIDE_H - BORDER * 2, fill: { color: theme.accentColor }, line: { type: 'none' } })
  cover.addShape(prs.ShapeType.rect, { x: SLIDE_W - BORDER - 0.05, y: BORDER, w: 0.05, h: SLIDE_H - BORDER * 2, fill: { color: theme.accentColor }, line: { type: 'none' } })

  // 内边框（浅色细线）
  const IB = BORDER + 0.18
  cover.addShape(prs.ShapeType.rect, { x: IB, y: IB, w: SLIDE_W - IB * 2, h: 0.03, fill: { color: theme.accent2Color }, line: { type: 'none' } })
  cover.addShape(prs.ShapeType.rect, { x: IB, y: SLIDE_H - IB - 0.03, w: SLIDE_W - IB * 2, h: 0.03, fill: { color: theme.accent2Color }, line: { type: 'none' } })
  cover.addShape(prs.ShapeType.rect, { x: IB, y: IB, w: 0.03, h: SLIDE_H - IB * 2, fill: { color: theme.accent2Color }, line: { type: 'none' } })
  cover.addShape(prs.ShapeType.rect, { x: SLIDE_W - IB - 0.03, y: IB, w: 0.03, h: SLIDE_H - IB * 2, fill: { color: theme.accent2Color }, line: { type: 'none' } })

  // 四角菱形装饰
  addCornerDiamond(cover, BORDER, BORDER, 0.28, theme.accentColor)
  addCornerDiamond(cover, SLIDE_W - BORDER, BORDER, 0.28, theme.accentColor)
  addCornerDiamond(cover, BORDER, SLIDE_H - BORDER, 0.28, theme.accentColor)
  addCornerDiamond(cover, SLIDE_W - BORDER, SLIDE_H - BORDER, 0.28, theme.accentColor)

  // 中央顶部小装饰
  addCornerDiamond(cover, SLIDE_W / 2, BORDER, 0.2, theme.accentColor)
  addCornerDiamond(cover, SLIDE_W / 2, SLIDE_H - BORDER, 0.2, theme.accentColor)

  // 顶部英文装饰文字
  cover.addText('· ELEGANT  PRESENTATION ·', { x: 1.5, y: 1.0, w: SLIDE_W - 3.0, h: 0.5, ...ff(12), color: theme.accentColor, align: 'center' })

  // 中央横线组
  cover.addShape(prs.ShapeType.line, { x: 1.8, y: 1.65, w: SLIDE_W - 3.6, h: 0, line: { color: theme.accentColor, width: 1.0 } })
  cover.addShape(prs.ShapeType.line, { x: 2.5, y: 1.78, w: SLIDE_W - 5.0, h: 0, line: { color: theme.accent2Color, width: 0.5 } })

  // 主标题
  cover.addText(data.title, { x: 1.5, y: 1.85, w: SLIDE_W - 3.0, h: 3.2, ...ff(40), bold: true, color: theme.titleColor, align: 'center', valign: 'middle', wrap: true })

  // 标题下装饰线组
  cover.addShape(prs.ShapeType.line, { x: 2.5, y: 5.2, w: SLIDE_W - 5.0, h: 0, line: { color: theme.accent2Color, width: 0.5 } })
  cover.addShape(prs.ShapeType.line, { x: 1.8, y: 5.35, w: SLIDE_W - 3.6, h: 0, line: { color: theme.accentColor, width: 1.0 } })

  // 底部文字
  cover.addText('典雅·专业·简约', { x: 1.5, y: 5.5, w: SLIDE_W - 3.0, h: 0.6, ...ff(13), color: theme.accentColor, align: 'center' })

  // ── 内容页 ──
  for (let si = 0; si < data.slides.length; si++) {
    const slide = data.slides[si]
    const s = prs.addSlide()
    s.background = { color: theme.bgColor }

    // 四边装饰框
    s.addShape(prs.ShapeType.rect, { x: BORDER, y: BORDER, w: SLIDE_W - BORDER * 2, h: 0.04, fill: { color: theme.accentColor }, line: { type: 'none' } })
    s.addShape(prs.ShapeType.rect, { x: BORDER, y: SLIDE_H - BORDER - 0.04, w: SLIDE_W - BORDER * 2, h: 0.04, fill: { color: theme.accentColor }, line: { type: 'none' } })
    s.addShape(prs.ShapeType.rect, { x: BORDER, y: BORDER, w: 0.04, h: SLIDE_H - BORDER * 2, fill: { color: theme.accentColor }, line: { type: 'none' } })
    s.addShape(prs.ShapeType.rect, { x: SLIDE_W - BORDER - 0.04, y: BORDER, w: 0.04, h: SLIDE_H - BORDER * 2, fill: { color: theme.accentColor }, line: { type: 'none' } })
    // 四角菱形
    addCornerDiamond(s, BORDER, BORDER, 0.22, theme.accentColor)
    addCornerDiamond(s, SLIDE_W - BORDER, BORDER, 0.22, theme.accentColor)
    addCornerDiamond(s, BORDER, SLIDE_H - BORDER, 0.22, theme.accentColor)
    addCornerDiamond(s, SLIDE_W - BORDER, SLIDE_H - BORDER, 0.22, theme.accentColor)

    // 右上页码装饰
    const numStr = String(si + 1).padStart(2, '0')
    s.addText(numStr, { x: SLIDE_W - 1.6, y: 0.55, w: 1.0, h: 0.9, ...ff(40), bold: true, color: 'EBF0F5', align: 'right' })
    s.addShape(prs.ShapeType.rect, { x: SLIDE_W - 0.85, y: 0.55, w: 0.5, h: 0.04, fill: { color: theme.accentColor }, line: { type: 'none' } })

    // 标题区：左侧菱形标记 + 标题
    addCornerDiamond(s, 0.85, 0.75, 0.22, theme.accentColor)
    s.addText(slide.title, { x: 1.15, y: 0.45, w: SLIDE_W - 3.0, h: 0.7, ...ff(26), bold: true, color: theme.titleColor, valign: 'middle', wrap: true })

    // 标题下双横线
    s.addShape(prs.ShapeType.line, { x: 0.7, y: 1.25, w: SLIDE_W - 1.4, h: 0, line: { color: theme.accentColor, width: 1.5 } })
    s.addShape(prs.ShapeType.line, { x: 1.2, y: 1.35, w: SLIDE_W - 2.4, h: 0, line: { color: theme.accent2Color, width: 0.5 } })

    // 内容要点
    if (slide.bullets?.length) {
      const bullets = slide.bullets.slice(0, 7)
      const startY = 1.5
      const availH = SLIDE_H - startY - 0.55
      const itemH = Math.min(availH / bullets.length, 0.88)

      for (let i = 0; i < bullets.length; i++) {
        const iy = startY + i * itemH
        // 菱形标记
        addCornerDiamond(s, 0.95, iy + itemH / 2, 0.16, theme.accentColor)
        // 浅色背景条（交替）
        if (i % 2 === 0) {
          s.addShape(prs.ShapeType.rect, { x: 0.7, y: iy + 0.04, w: SLIDE_W - 1.4, h: itemH - 0.08, fill: { color: 'F5F7FA' }, line: { type: 'none' } })
        }
        s.addText(bullets[i], { x: 1.25, y: iy, w: SLIDE_W - 1.9, h: itemH, ...ff(18), color: theme.bulletColor, valign: 'middle', wrap: true })
      }
    }
  }
}

// ── Minimal White (极简白色PPT style) ─────────────────────────────────────────
// 封面：右侧大色块 + 标题在左侧 + 大号英文装饰字 + 几何组合
// 内容：左侧色块页眉 + 蓝色序号方块 + 下划线

function buildMinimalWhite(prs: pptxgen, data: PresentationData, theme: ThemeConfig, ff: (s?: number) => object) {
  const PANEL_X = 7.5  // 右侧色块起始位置

  // ── 封面 ──
  const cover = prs.addSlide()
  cover.background = { color: theme.bgColor }

  // 大英文装饰字（浅灰，背景装饰）
  cover.addText('SLIDE', { x: 0.3, y: 4.5, w: 8.0, h: 2.5, ...ff(120), bold: true, color: 'F0F4F8', align: 'left', valign: 'bottom' })

  // 右侧大色块
  cover.addShape(prs.ShapeType.rect, { x: PANEL_X, y: 0, w: SLIDE_W - PANEL_X, h: SLIDE_H, fill: { color: theme.accentColor }, line: { type: 'none' } })
  // 右侧色块内装饰：小圆点群
  for (let row = 0; row < 3; row++) {
    for (let col = 0; col < 4; col++) {
      cover.addShape(prs.ShapeType.ellipse, {
        x: PANEL_X + 0.4 + col * 1.2, y: 0.5 + row * 1.1, w: 0.12, h: 0.12,
        fill: { color: 'FFFFFF' }, line: { type: 'none' },
      })
    }
  }
  // 右侧色块内大圆
  cover.addShape(prs.ShapeType.ellipse, { x: PANEL_X + 0.5, y: 3.5, w: 2.5, h: 2.5, fill: { color: theme.darkColor }, line: { type: 'none' } })
  cover.addShape(prs.ShapeType.ellipse, { x: PANEL_X + 0.9, y: 3.9, w: 1.7, h: 1.7, fill: { color: theme.accentColor }, line: { type: 'none' } })

  // 左侧标题区
  cover.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 0.12, h: SLIDE_H, fill: { color: theme.accentColor }, line: { type: 'none' } })
  cover.addShape(prs.ShapeType.rect, { x: 0.12, y: 0, w: 0.06, h: SLIDE_H, fill: { color: theme.accent2Color }, line: { type: 'none' } })

  cover.addText(data.title, { x: 0.5, y: 1.5, w: PANEL_X - 0.8, h: 3.0, ...ff(40), bold: true, color: theme.titleColor, align: 'left', valign: 'middle', wrap: true })

  // 标题下方装饰线
  cover.addShape(prs.ShapeType.rect, { x: 0.5, y: 4.65, w: 3.8, h: 0.08, fill: { color: theme.accentColor }, line: { type: 'none' } })
  cover.addShape(prs.ShapeType.rect, { x: 0.5, y: 4.82, w: 2.0, h: 0.05, fill: { color: theme.accent2Color }, line: { type: 'none' } })
  cover.addText('PRESENTATION', { x: 0.5, y: 5.1, w: 6.5, h: 0.45, ...ff(12), color: theme.accentColor, align: 'left', bold: true })

  // ── 内容页 ──
  for (let si = 0; si < data.slides.length; si++) {
    const slide = data.slides[si]
    const s = prs.addSlide()
    s.background = { color: theme.bgColor }

    // 顶部双色页眉
    s.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 4.5, h: 1.1, fill: { color: theme.accentColor }, line: { type: 'none' } })
    s.addShape(prs.ShapeType.rect, { x: 4.5, y: 0, w: SLIDE_W - 4.5, h: 1.1, fill: { color: 'EEF5FC' }, line: { type: 'none' } })
    // 页眉橙色序号方块
    const numStr = String(si + 1).padStart(2, '0')
    s.addShape(prs.ShapeType.rect, { x: SLIDE_W - 1.3, y: 0.15, w: 1.0, h: 0.8, fill: { color: theme.accent2Color }, line: { type: 'none' } })
    s.addText(numStr, { x: SLIDE_W - 1.3, y: 0.15, w: 1.0, h: 0.8, ...ff(24), bold: true, color: 'FFFFFF', align: 'center', valign: 'middle' })
    // 左侧竖线
    s.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 0.12, h: SLIDE_H, fill: { color: theme.accentColor }, line: { type: 'none' } })

    // 标题
    s.addText(slide.title, { x: 0.35, y: 0.1, w: SLIDE_W - 1.8, h: 0.9, ...ff(26), bold: true, color: 'FFFFFF', valign: 'middle', wrap: true })

    // 内容区下划线
    s.addShape(prs.ShapeType.rect, { x: 0.3, y: 1.18, w: 4.0, h: 0.07, fill: { color: theme.accentColor }, line: { type: 'none' } })
    s.addShape(prs.ShapeType.rect, { x: 0.3, y: 1.3, w: 2.5, h: 0.04, fill: { color: theme.accent2Color }, line: { type: 'none' } })

    // 要点列表
    if (slide.bullets?.length) {
      const bullets = slide.bullets.slice(0, 7)
      const startY = 1.45
      const availH = SLIDE_H - startY - 0.4
      const itemH = Math.min(availH / bullets.length, 0.88)

      for (let i = 0; i < bullets.length; i++) {
        const iy = startY + i * itemH
        // 蓝色方块标记
        s.addShape(prs.ShapeType.rect, { x: 0.32, y: iy + itemH / 2 - 0.12, w: 0.24, h: 0.24, fill: { color: theme.accentColor }, line: { type: 'none' } })
        // 橙色小方块
        s.addShape(prs.ShapeType.rect, { x: 0.56, y: iy + itemH / 2 + 0.0, w: 0.1, h: 0.1, fill: { color: theme.accent2Color }, line: { type: 'none' } })
        s.addText(bullets[i], { x: 0.85, y: iy, w: SLIDE_W - 1.4, h: itemH, ...ff(18), color: theme.bulletColor, valign: 'middle', wrap: true })
        if (i < bullets.length - 1) {
          s.addShape(prs.ShapeType.line, { x: 0.85, y: iy + itemH, w: SLIDE_W - 1.5, h: 0, line: { color: 'DDEAF8', width: 0.5 } })
        }
      }
    }

    // 底部
    s.addShape(prs.ShapeType.rect, { x: 0, y: SLIDE_H - 0.3, w: SLIDE_W, h: 0.3, fill: { color: 'F0F5FB' }, line: { type: 'none' } })
    s.addText(data.title, { x: 0.3, y: SLIDE_H - 0.3, w: SLIDE_W - 0.6, h: 0.28, ...ff(9), color: theme.footerColor, align: 'right' })
  }
}

// ── Minimal Particles (极简点线粒子 style, 深色科技) ────────────────────────
// 封面：深色背景 + 多条横向光线 + 圆点集群 + 居中标题
// 内容：深色 + 蓝色电光装饰 + 明亮要点

function buildMinimalParticles(prs: pptxgen, data: PresentationData, theme: ThemeConfig, ff: (s?: number) => object) {
  // ── 封面 ──
  const cover = prs.addSlide()
  cover.background = { color: theme.bgColor }

  // 背景装饰：多条粗细不一的横线（模拟光栅/粒子轨迹）
  const lines = [
    { y: 0.8, w: SLIDE_W * 0.7, x: 0, h: 0.04, color: theme.accentColor },
    { y: 1.1, w: SLIDE_W * 0.45, x: 0, h: 0.02, color: '3A5A7A' },
    { y: 1.3, w: SLIDE_W * 0.25, x: 0, h: 0.015, color: '2A4A6A' },
    { y: 5.8, w: SLIDE_W * 0.6, x: SLIDE_W * 0.4, h: 0.04, color: theme.accentColor },
    { y: 6.1, w: SLIDE_W * 0.4, x: SLIDE_W * 0.6, h: 0.02, color: '3A5A7A' },
    { y: 6.3, w: SLIDE_W * 0.2, x: SLIDE_W * 0.8, h: 0.015, color: '2A4A6A' },
  ]
  for (const l of lines) {
    cover.addShape(prs.ShapeType.rect, { x: l.x, y: l.y, w: l.w, h: l.h, fill: { color: l.color }, line: { type: 'none' } })
  }

  // 圆点集群（左上角）
  const dots = [[0.4, 2.2], [0.8, 2.0], [1.2, 2.4], [0.6, 2.6], [1.0, 2.8], [1.4, 2.1], [0.2, 2.8]]
  for (const [dx, dy] of dots) {
    cover.addShape(prs.ShapeType.ellipse, { x: dx, y: dy, w: 0.09, h: 0.09, fill: { color: theme.accentColor }, line: { type: 'none' } })
  }
  // 圆点集群（右下角）
  const dots2 = [[11.8, 5.2], [12.2, 5.0], [12.6, 5.4], [11.6, 5.5], [12.0, 5.7], [12.5, 5.1]]
  for (const [dx, dy] of dots2) {
    cover.addShape(prs.ShapeType.ellipse, { x: dx, y: dy, w: 0.09, h: 0.09, fill: { color: theme.accentColor }, line: { type: 'none' } })
  }

  // 角落装饰方块
  cover.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 0.5, h: 0.5, fill: { color: theme.accentColor }, line: { type: 'none' } })
  cover.addShape(prs.ShapeType.rect, { x: SLIDE_W - 0.5, y: SLIDE_H - 0.5, w: 0.5, h: 0.5, fill: { color: theme.accentColor }, line: { type: 'none' } })
  cover.addShape(prs.ShapeType.rect, { x: 0.6, y: 0, w: 0.25, h: 0.25, fill: { color: theme.accent2Color }, line: { type: 'none' } })

  // 中央横线组（标题上下）
  cover.addShape(prs.ShapeType.rect, { x: 1.5, y: 2.5, w: SLIDE_W - 3.0, h: 0.04, fill: { color: theme.accentColor }, line: { type: 'none' } })
  cover.addShape(prs.ShapeType.rect, { x: 3.0, y: 2.58, w: SLIDE_W - 6.0, h: 0.02, fill: { color: theme.accent2Color }, line: { type: 'none' } })

  // 主标题
  cover.addText(data.title, { x: 1.2, y: 2.65, w: SLIDE_W - 2.4, h: 2.4, ...ff(42), bold: true, color: theme.titleColor, align: 'center', valign: 'middle', wrap: true })

  // 标题下装饰线
  cover.addShape(prs.ShapeType.rect, { x: 3.0, y: 5.1, w: SLIDE_W - 6.0, h: 0.02, fill: { color: theme.accent2Color }, line: { type: 'none' } })
  cover.addShape(prs.ShapeType.rect, { x: 1.5, y: 5.18, w: SLIDE_W - 3.0, h: 0.04, fill: { color: theme.accentColor }, line: { type: 'none' } })

  // ── 内容页 ──
  for (let si = 0; si < data.slides.length; si++) {
    const slide = data.slides[si]
    const s = prs.addSlide()
    s.background = { color: theme.bgColor }

    // 顶部装饰线组
    s.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: SLIDE_W, h: 0.06, fill: { color: theme.accentColor }, line: { type: 'none' } })
    s.addShape(prs.ShapeType.rect, { x: 0, y: 0.06, w: SLIDE_W, h: 0.03, fill: { color: theme.accent2Color }, line: { type: 'none' } })

    // 左侧标题装饰：竖线+圆点
    s.addShape(prs.ShapeType.rect, { x: 0.48, y: 0.28, w: 0.07, h: 0.85, fill: { color: theme.accentColor }, line: { type: 'none' } })
    s.addShape(prs.ShapeType.ellipse, { x: 0.44, y: 0.22, w: 0.16, h: 0.16, fill: { color: theme.accentColor }, line: { type: 'none' } })
    s.addShape(prs.ShapeType.ellipse, { x: 0.44, y: 0.97, w: 0.16, h: 0.16, fill: { color: theme.accentColor }, line: { type: 'none' } })

    // 右上角序号
    const numStr = String(si + 1).padStart(2, '0')
    s.addText(numStr, { x: SLIDE_W - 2.0, y: 0.2, w: 1.7, h: 0.9, ...ff(40), bold: true, color: '1A3A5A', align: 'right', valign: 'middle' })

    // 标题
    s.addText(slide.title, { x: 0.85, y: 0.25, w: SLIDE_W - 2.9, h: 0.85, ...ff(26), bold: true, color: theme.titleColor, valign: 'middle', wrap: true })

    // 分割线+小圆点
    s.addShape(prs.ShapeType.rect, { x: 0.5, y: 1.2, w: SLIDE_W - 1.0, h: 0.03, fill: { color: theme.accentColor }, line: { type: 'none' } })
    s.addShape(prs.ShapeType.ellipse, { x: 0.44, y: 1.17, w: 0.1, h: 0.1, fill: { color: theme.accentColor }, line: { type: 'none' } })
    s.addShape(prs.ShapeType.ellipse, { x: SLIDE_W - 0.6, y: 1.17, w: 0.1, h: 0.1, fill: { color: theme.accentColor }, line: { type: 'none' } })

    // 要点列表
    if (slide.bullets?.length) {
      const bullets = slide.bullets.slice(0, 6)
      const startY = 1.38
      const availH = SLIDE_H - startY - 0.4
      const itemH = Math.min(availH / bullets.length, 0.95)

      for (let i = 0; i < bullets.length; i++) {
        const iy = startY + i * itemH
        // 三角箭头标记
        s.addShape(prs.ShapeType.rect, { x: 0.5, y: iy + itemH / 2 - 0.1, w: 0.18, h: 0.2, fill: { color: theme.accentColor }, line: { type: 'none' } })
        s.addText(bullets[i], { x: 0.9, y: iy, w: SLIDE_W - 1.5, h: itemH, ...ff(18), color: theme.bulletColor, valign: 'middle', wrap: true })
        if (i < bullets.length - 1) {
          s.addShape(prs.ShapeType.line, { x: 0.9, y: iy + itemH, w: SLIDE_W - 1.6, h: 0, line: { color: '1A3A5A', width: 0.75 } })
        }
      }
    }

    // 底部装饰
    s.addShape(prs.ShapeType.rect, { x: SLIDE_W - 0.5, y: SLIDE_H - 0.5, w: 0.5, h: 0.5, fill: { color: theme.accentColor }, line: { type: 'none' } })
    s.addText(data.title, { x: 0.5, y: SLIDE_H - 0.35, w: SLIDE_W - 1.3, h: 0.28, ...ff(9), color: theme.footerColor, align: 'right' })
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
