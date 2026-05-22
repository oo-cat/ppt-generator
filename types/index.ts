export interface Message {
  role: 'user' | 'assistant'
  content: string
}

export interface SlideContent {
  title: string
  bullets?: string[]
  notes?: string
}

export interface TemplateTheme {
  bgColor: string
  titleColor: string
  bodyColor: string
  accentColor: string
  fontName: string
}

export type ThemeId =
  | 'academic'
  | 'business'
  | 'colorful'
  | 'chinese-blue'
  | 'chinese-ink'
  | 'chinese-elegant'
  | 'minimal-white'
  | 'minimal-particles'

export interface PresentationData {
  title: string
  theme: ThemeId
  slides: SlideContent[]
  templateTheme?: TemplateTheme
}

export interface ChatRequest {
  messages: Message[]
  fileContent?: string
  scene?: string
  stage: 'analyze' | 'outline' | 'generate'
}

export interface GenerateRequest {
  presentation: PresentationData
  format: 'pptx' | 'html'
}
