import OpenAI from 'openai'
import { Message } from '@/types'

const client = new OpenAI({
  apiKey: process.env.DEEPSEEK_API_KEY,
  baseURL: 'https://api.deepseek.com',
})

export const SYSTEM_PROMPT = `你是一个专业的PPT制作助手。你的工作流程如下：

**第一阶段（分析）**：用户发送素材和场景后，你需要提问澄清以下信息（如果用户没有明确说明）：
- 目标受众是谁
- 汇报时长
- 期望幻灯片数量（建议范围）
- 重点强调的内容
- 风格偏好（学术/商务/活泼）

一次最多问3个最关键的问题，不要一次问太多。

**第二阶段（大纲）**：澄清完毕后，输出结构化大纲，格式如下：
\`\`\`outline
标题: [演示文稿标题]
主题: academic/business/colorful
幻灯片数: [数字]

1. [页面标题]
   - 要点1
   - 要点2

2. [页面标题]
   ...
\`\`\`
然后询问用户是否满意，是否需要调整。

**第三阶段（确认生成）**：用户确认后，输出 JSON，格式严格如下：
\`\`\`json
{
  "ready": true,
  "presentation": {
    "title": "演示文稿标题",
    "theme": "academic",
    "slides": [
      {
        "title": "页面标题",
        "bullets": ["要点1", "要点2", "要点3"],
        "notes": "演讲者备注（可选）"
      }
    ]
  }
}
\`\`\`

始终用中文回复。保持专业、简洁。`

export async function chat(messages: Message[]): Promise<string> {
  const response = await client.chat.completions.create({
    model: 'deepseek-chat',
    messages: [
      { role: 'system', content: SYSTEM_PROMPT },
      ...messages,
    ],
    temperature: 0.7,
    max_tokens: 4096,
  })

  return response.choices[0].message.content ?? ''
}

export function extractPresentationJSON(text: string): object | null {
  const match = text.match(/```json\s*([\s\S]*?)\s*```/)
  if (!match) return null
  try {
    const parsed = JSON.parse(match[1])
    if (parsed.ready && parsed.presentation) return parsed.presentation
    return null
  } catch {
    return null
  }
}
