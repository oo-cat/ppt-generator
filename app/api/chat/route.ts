import { NextRequest } from 'next/server'
import { chat } from '@/lib/deepseek'
import { Message } from '@/types'

export async function POST(request: NextRequest) {
  try {
    const body = await request.json()
    const messages: Message[] = body.messages

    if (!messages || !Array.isArray(messages)) {
      return Response.json({ error: '无效的消息格式' }, { status: 400 })
    }

    const reply = await chat(messages)
    return Response.json({ reply })
  } catch (error) {
    console.error('Chat error:', error)
    return Response.json({ error: '对话请求失败，请检查 API Key' }, { status: 500 })
  }
}
