'use client'

import { useState, useCallback } from 'react'
import FileUpload from '@/components/FileUpload'
import ChatPanel from '@/components/ChatPanel'
import OcrWorker from '@/components/OcrWorker'
import TemplateUpload from '@/components/TemplateUpload'
import ThemeSelector from '@/components/ThemeSelector'
import { Message, PresentationData, TemplateTheme, ThemeId } from '@/types'
import { extractPresentationJSON } from '@/lib/utils'

function isDark(hex: string): boolean {
  const h = hex.replace('#', '')
  const r = parseInt(h.slice(0, 2), 16) / 255
  const g = parseInt(h.slice(2, 4), 16) / 255
  const b = parseInt(h.slice(4, 6), 16) / 255
  return 0.299 * r + 0.587 * g + 0.114 * b < 0.4
}

type Stage = 'input' | 'chatting' | 'ready'

export default function Home() {
  const [messages, setMessages] = useState<Message[]>([])
  const [input, setInput] = useState('')
  const [loading, setLoading] = useState(false)
  const [stage, setStage] = useState<Stage>('input')
  const [presentation, setPresentation] = useState<PresentationData | null>(null)
  const [generating, setGenerating] = useState<'pptx' | 'html' | null>(null)

  const [fileContent, setFileContent] = useState('')
  const [scene, setScene] = useState('')
  const [duration, setDuration] = useState('')
  const [ocrPending, setOcrPending] = useState<{ base64: string; mimeType: string } | null>(null)
  const [ocrStatus, setOcrStatus] = useState<'idle' | 'running' | 'done'>('idle')
  const [templateTheme, setTemplateTheme] = useState<TemplateTheme | null>(null)
  const [selectedTheme, setSelectedTheme] = useState<ThemeId>('chinese-blue')

  const handleOcrImage = useCallback((_base64: string, mimeType: string, fileName: string) => {
    setOcrPending({ base64: _base64, mimeType })
    setOcrStatus('running')
    setFileContent(`[正在识别图片 ${fileName}...]`)
  }, [])

  const handleOcrResult = useCallback((text: string) => {
    setFileContent(text)
    setOcrStatus('done')
    setOcrPending(null)
  }, [])

  const handleOcrError = useCallback((msg: string) => {
    setFileContent('')
    setOcrStatus('idle')
    setOcrPending(null)
    alert(msg)
  }, [])

  async function startAnalysis() {
    if (!fileContent && !scene) return
    const themeNames: Record<ThemeId, string> = {
      'chinese-blue': '中国风蓝', 'chinese-ink': '极简墨', 'chinese-elegant': '雅韵',
      'minimal-white': '极简白', 'minimal-particles': '点线粒子深色',
      'academic': '学术白', 'business': '商务深色', 'colorful': '活泼彩色',
    }
    const parts = [
      fileContent && `【素材内容】\n${fileContent}`,
      scene && `【使用场景】${scene}`,
      duration && `【展示时长】${duration}分钟`,
      `【风格偏好】${themeNames[selectedTheme]}（主题标识：${selectedTheme}）`,
    ].filter(Boolean)

    const userMsg: Message = { role: 'user', content: parts.join('\n\n') }
    const newMessages = [userMsg]
    setMessages(newMessages)
    setStage('chatting')
    await sendToAI(newMessages)
  }

  async function sendMessage() {
    if (!input.trim() || loading) return
    const userMsg: Message = { role: 'user', content: input.trim() }
    const newMessages = [...messages, userMsg]
    setMessages(newMessages)
    setInput('')
    await sendToAI(newMessages)
  }

  async function sendToAI(msgs: Message[]) {
    setLoading(true)
    try {
      const res = await fetch('/api/chat', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ messages: msgs }),
      })
      const data = await res.json()
      if (data.error) throw new Error(data.error)

      const reply = data.reply as string
      const assistantMsg: Message = { role: 'assistant', content: reply }
      const updatedMessages = [...msgs, assistantMsg]
      setMessages(updatedMessages)

      const pptData = extractPresentationJSON(reply)
      if (pptData) {
        const ppt = pptData as PresentationData
        if (templateTheme) {
          ppt.templateTheme = templateTheme
        } else {
          // Ensure selected theme is preserved even if AI returns a different one
          ppt.theme = selectedTheme
        }
        setPresentation(ppt)
        setStage('ready')
      }
    } catch (e) {
      const errMsg: Message = { role: 'assistant', content: `出错了：${String(e)}` }
      setMessages((prev) => [...prev, errMsg])
    } finally {
      setLoading(false)
    }
  }

  async function generate(format: 'pptx' | 'html') {
    if (!presentation) return
    setGenerating(format)
    try {
      const res = await fetch('/api/generate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ presentation, format }),
      })

      if (!res.ok) {
        const err = await res.json()
        throw new Error(err.error)
      }

      if (format === 'pptx') {
        const blob = await res.blob()
        const url = URL.createObjectURL(blob)
        const a = document.createElement('a')
        a.href = url
        a.download = `${presentation.title}.pptx`
        a.click()
        URL.revokeObjectURL(url)
      } else {
        const html = await res.text()
        const blob = new Blob([html], { type: 'text/html' })
        const url = URL.createObjectURL(blob)
        window.open(url, '_blank')
      }
    } catch (e) {
      alert('生成失败：' + String(e))
    } finally {
      setGenerating(null)
    }
  }

  function reset() {
    setMessages([])
    setInput('')
    setStage('input')
    setPresentation(null)
    setFileContent('')
    setScene('')
    setOcrPending(null)
    setOcrStatus('idle')
    setTemplateTheme(null)
    setDuration('')
    setSelectedTheme('chinese-blue')
  }

  return (
    <div className="min-h-screen bg-gray-50 flex flex-col">
      <header className="bg-white border-b px-6 py-4 flex items-center justify-between">
        <h1 className="text-xl font-bold text-gray-800">PPT 生成助手</h1>
        {stage !== 'input' && (
          <button onClick={reset} className="text-sm text-gray-500 hover:text-gray-700">
            重新开始
          </button>
        )}
      </header>

      <div className="flex-1 flex overflow-hidden" style={{ height: 'calc(100vh - 65px)' }}>
        {stage === 'input' && (
          <div className="w-full max-w-2xl mx-auto p-6 flex flex-col gap-4 overflow-y-auto">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">上传素材（可选）</label>
              <FileUpload
                onParsed={(content) => setFileContent(content)}
                onOcrImage={handleOcrImage}
                disabled={ocrStatus === 'running'}
              />
              {ocrStatus === 'running' && (
                <p className="text-xs text-blue-500 mt-1">正在 OCR 识别图片文字...</p>
              )}
              {ocrStatus === 'done' && fileContent && (
                <details className="mt-2">
                  <summary className="text-xs text-gray-500 cursor-pointer">查看识别结果</summary>
                  <p className="text-xs text-gray-400 mt-1 max-h-32 overflow-y-auto whitespace-pre-wrap">{fileContent}</p>
                </details>
              )}
            </div>

            <div className="flex gap-3">
              <div className="flex-1">
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  使用场景 <span className="text-red-500">*</span>
                </label>
                <textarea
                  className="w-full border rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-300 resize-none"
                  placeholder={'例：组会汇报，5人研究小组\n例：课程期末汇报，面向老师和同学'}
                  rows={3}
                  value={scene}
                  onChange={(e) => setScene(e.target.value)}
                />
              </div>
              <div className="w-28 shrink-0">
                <label className="block text-sm font-medium text-gray-700 mb-2">展示时长</label>
                <div className="relative">
                  <input
                    type="number"
                    min={1}
                    max={120}
                    className="w-full border rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-300 pr-8"
                    placeholder="15"
                    value={duration}
                    onChange={(e) => setDuration(e.target.value)}
                  />
                  <span className="absolute right-3 top-1/2 -translate-y-1/2 text-xs text-gray-400">分钟</span>
                </div>
              </div>
            </div>

            <ThemeSelector value={selectedTheme} onChange={setSelectedTheme} />

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">上传 PPT 模板（可选，会覆盖上方风格配色）</label>
              <TemplateUpload
                onTemplate={(theme) => setTemplateTheme(theme)}
                onClear={() => setTemplateTheme(null)}
                hasTemplate={templateTheme !== null}
              />
              {templateTheme && (
                <div className="mt-2 space-y-1">
                  <div className="flex gap-2 items-center">
                    <span className="text-xs text-purple-600">已提取配色：</span>
                    {[templateTheme.bgColor, templateTheme.titleColor, templateTheme.accentColor, templateTheme.bodyColor].map((c, i) => (
                      <span key={i} className="w-4 h-4 rounded-full border border-gray-200 inline-block" style={{ backgroundColor: `#${c}` }} />
                    ))}
                    <span className="text-xs text-gray-400">{templateTheme.fontName}</span>
                  </div>
                  {isDark(templateTheme.bgColor) && (
                    <p className="text-xs text-gray-400">深色背景模板，已自动修复文字可读性</p>
                  )}
                </div>
              )}
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">或直接输入文字素材</label>
              <textarea
                className="w-full border rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-300 resize-none"
                placeholder="粘贴文章、笔记、提纲等文字内容..."
                rows={6}
                value={fileContent}
                onChange={(e) => setFileContent(e.target.value)}
              />
            </div>

            <button
              onClick={startAnalysis}
              disabled={(!fileContent && !scene) || ocrStatus === 'running' || loading}
              className="w-full py-3 bg-blue-500 text-white rounded-xl font-medium hover:bg-blue-600 disabled:opacity-40 disabled:cursor-not-allowed transition-colors"
            >
              开始分析
            </button>
          </div>
        )}

        {(stage === 'chatting' || stage === 'ready') && (
          <div className="flex flex-1 overflow-hidden">
            <div className="flex-1 flex flex-col border-r min-w-0">
              <ChatPanel
                messages={messages}
                loading={loading}
                input={input}
                onInputChange={setInput}
                onSend={sendMessage}
              />
            </div>

            <div className="w-64 p-4 flex flex-col gap-4 bg-white shrink-0">
              <div>
                <h3 className="text-sm font-medium text-gray-700 mb-2">当前素材</h3>
                <p className="text-xs text-gray-400 max-h-24 overflow-y-auto whitespace-pre-wrap">
                  {fileContent ? fileContent.slice(0, 200) + (fileContent.length > 200 ? '...' : '') : '无'}
                </p>
              </div>

              {stage === 'ready' && presentation && (
                <div className="mt-auto space-y-3">
                  <div className="p-3 bg-green-50 rounded-xl border border-green-200">
                    <p className="text-xs font-medium text-green-700">PPT 已就绪</p>
                    <p className="text-xs text-green-600 mt-1">{presentation.title}</p>
                    <p className="text-xs text-green-500">{presentation.slides.length} 页</p>
                  </div>
                  <button
                    onClick={() => generate('html')}
                    disabled={generating !== null}
                    className="w-full py-2 bg-indigo-500 text-white rounded-xl text-sm hover:bg-indigo-600 disabled:opacity-40 transition-colors"
                  >
                    {generating === 'html' ? '生成中...' : '在线预览'}
                  </button>
                  <button
                    onClick={() => generate('pptx')}
                    disabled={generating !== null}
                    className="w-full py-2 bg-blue-500 text-white rounded-xl text-sm hover:bg-blue-600 disabled:opacity-40 transition-colors"
                  >
                    {generating === 'pptx' ? '生成中...' : '下载 .pptx'}
                  </button>
                </div>
              )}

              {stage === 'chatting' && (
                <p className="text-xs text-gray-400 text-center mt-auto">
                  与助手对话完成后，
                  <br />
                  会自动出现生成按钮
                </p>
              )}
            </div>
          </div>
        )}
      </div>

      {ocrPending && (
        <OcrWorker
          imageBase64={ocrPending.base64}
          mimeType={ocrPending.mimeType}
          onResult={handleOcrResult}
          onError={handleOcrError}
        />
      )}
    </div>
  )
}
