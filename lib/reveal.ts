import { PresentationData } from '@/types'

import { ThemeId } from '@/types'

const THEME_STYLES: Record<ThemeId, { bg: string; text: string; accent: string; revealTheme: string }> = {
  academic: { bg: '#ffffff', text: '#1a1a2e', accent: '#0f3460', revealTheme: 'white' },
  business: { bg: '#1a1a2e', text: '#e2e2e2', accent: '#e94560', revealTheme: 'black' },
  colorful: { bg: '#f8f9fa', text: '#2c3e50', accent: '#6c3483', revealTheme: 'sky' },
  'chinese-blue': { bg: '#ffffff', text: '#1B3F5E', accent: '#4A80A1', revealTheme: 'white' },
  'chinese-ink': { bg: '#FAFAF8', text: '#1A1A1A', accent: '#4472C4', revealTheme: 'white' },
  'chinese-elegant': { bg: '#ffffff', text: '#2C3A4A', accent: '#7E97BA', revealTheme: 'white' },
  'minimal-white': { bg: '#ffffff', text: '#1A1A1A', accent: '#5B9BD5', revealTheme: 'white' },
  'minimal-particles': { bg: '#0D1B2A', text: '#E8F4FD', accent: '#5B9BD5', revealTheme: 'black' },
}

export function generateRevealHtml(data: PresentationData): string {
  const style = THEME_STYLES[data.theme] ?? THEME_STYLES.academic

  const coverSlide = `
    <section>
      <h1 style="color:${style.accent}">${escapeHtml(data.title)}</h1>
    </section>`

  const contentSlides = data.slides
    .map((slide) => {
      const bullets =
        slide.bullets && slide.bullets.length > 0
          ? `<ul>${slide.bullets.map((b) => `<li>${escapeHtml(b)}</li>`).join('')}</ul>`
          : ''
      const notes = slide.notes ? `<aside class="notes">${escapeHtml(slide.notes)}</aside>` : ''
      return `
    <section>
      <h2>${escapeHtml(slide.title)}</h2>
      ${bullets}
      ${notes}
    </section>`
    })
    .join('')

  return `<!DOCTYPE html>
<html lang="zh">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${escapeHtml(data.title)}</title>
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/reveal.js@5.1.0/dist/reset.css">
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/reveal.js@5.1.0/dist/reveal.css">
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/reveal.js@5.1.0/dist/theme/${style.revealTheme}.css">
  <style>
    .reveal h1, .reveal h2 { color: ${style.accent}; }
    .reveal ul li { margin: 0.4em 0; font-size: 0.9em; }
  </style>
</head>
<body>
  <div class="reveal">
    <div class="slides">
      ${coverSlide}
      ${contentSlides}
    </div>
  </div>
  <script src="https://cdn.jsdelivr.net/npm/reveal.js@5.1.0/dist/reveal.js"></script>
  <script>
    Reveal.initialize({ hash: true, slideNumber: true, transition: 'slide' });
  </script>
</body>
</html>`
}

function escapeHtml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}
