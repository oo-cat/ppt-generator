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
