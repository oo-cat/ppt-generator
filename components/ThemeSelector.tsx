'use client'

import { ThemeId } from '@/types'

interface ThemeOption {
  id: ThemeId
  name: string
  desc: string
  bg: string
  accent: string
  text: string
  tag?: string
}

const THEMES: ThemeOption[] = [
  { id: 'chinese-blue', name: '中国风蓝', desc: '竖栏布局，编号要点', bg: '#FFFFFF', accent: '#4A80A1', text: '#1B3F5E', tag: '中国风' },
  { id: 'chinese-ink', name: '极简墨', desc: '白底，双线装饰', bg: '#FAFAF8', accent: '#4472C4', text: '#1A1A1A', tag: '中国风' },
  { id: 'chinese-elegant', name: '雅韵', desc: '边框精致，菱形点缀', bg: '#FFFFFF', accent: '#7E97BA', text: '#2C3A4A', tag: '中国风' },
  { id: 'minimal-white', name: '极简白', desc: '方块+破折号，简洁现代', bg: '#FFFFFF', accent: '#5B9BD5', text: '#1A1A1A', tag: '极简' },
  { id: 'minimal-particles', name: '点线粒子', desc: '深蓝暗色，科技感', bg: '#0D1B2A', accent: '#5B9BD5', text: '#E8F4FD', tag: '极简' },
  { id: 'academic', name: '学术白', desc: '经典白底，蓝色调', bg: '#FFFFFF', accent: '#2E75B6', text: '#1F3864', tag: '通用' },
  { id: 'business', name: '商务深色', desc: '暗夜蓝，红色强调', bg: '#1A1A2E', accent: '#E94560', text: '#E2E2E2', tag: '通用' },
  { id: 'colorful', name: '活泼彩色', desc: '浅底，紫红色调', bg: '#FAFAFA', accent: '#E74C3C', text: '#6C3483', tag: '通用' },
]

interface Props {
  value: ThemeId
  onChange: (theme: ThemeId) => void
}

export default function ThemeSelector({ value, onChange }: Props) {
  return (
    <div>
      <label className="block text-sm font-medium text-gray-700 mb-2">选择风格</label>
      <div className="grid grid-cols-4 gap-2">
        {THEMES.map((t) => (
          <button
            key={t.id}
            onClick={() => onChange(t.id)}
            className={`relative rounded-xl border-2 p-2 text-left transition-all ${
              value === t.id
                ? 'border-blue-500 ring-2 ring-blue-200'
                : 'border-gray-200 hover:border-gray-300'
            }`}
          >
            {/* Color preview */}
            <div
              className="w-full rounded-md mb-1.5 flex flex-col overflow-hidden"
              style={{ height: 44, backgroundColor: t.bg, border: '1px solid #eee' }}
            >
              {/* Simulate slide layout preview */}
              <div style={{ backgroundColor: t.accent, height: 6, width: '100%' }} />
              <div className="flex-1 flex items-center px-1.5 gap-1">
                <div style={{ backgroundColor: t.accent, width: 3, height: 20, borderRadius: 1, flexShrink: 0 }} />
                <div className="flex flex-col gap-0.5 flex-1">
                  <div style={{ backgroundColor: t.text, height: 5, width: '80%', borderRadius: 1, opacity: 0.85 }} />
                  <div style={{ backgroundColor: t.text, height: 3, width: '95%', borderRadius: 1, opacity: 0.4 }} />
                  <div style={{ backgroundColor: t.text, height: 3, width: '70%', borderRadius: 1, opacity: 0.4 }} />
                </div>
              </div>
            </div>
            <p className="text-xs font-medium text-gray-800 leading-none">{t.name}</p>
            <p className="text-[10px] text-gray-400 mt-0.5 leading-tight">{t.desc}</p>
            {value === t.id && (
              <span className="absolute top-1 right-1 w-3.5 h-3.5 bg-blue-500 rounded-full flex items-center justify-center">
                <svg width="8" height="8" viewBox="0 0 8 8" fill="none"><path d="M1.5 4L3 5.5L6.5 2.5" stroke="white" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/></svg>
              </span>
            )}
          </button>
        ))}
      </div>
    </div>
  )
}
