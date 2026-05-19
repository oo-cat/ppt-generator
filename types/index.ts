export interface Message {
  role: 'user' | 'assistant'
  content: string
}

export interface SlideContent {
  title: string
  bullets?: string[]
  notes?: string
}

export interface PresentationData {
  title: string
  theme: 'academic' | 'business' | 'colorful'
  slides: SlideContent[]
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
