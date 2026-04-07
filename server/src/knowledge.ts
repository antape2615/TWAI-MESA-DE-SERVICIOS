import fs from 'node:fs/promises'
import path from 'node:path'
import { KNOWLEDGE_FILE } from './paths.js'
import { z } from 'zod'

const EntrySchema = z.object({
  id: z.string(),
  keywords: z.array(z.string()),
  title: z.string(),
  response: z.string(),
})

const KnowledgeSchema = z.object({
  entries: z.array(EntrySchema),
  slaHoursByPriority: z.record(z.string(), z.number()),
})

export type KnowledgeEntry = z.infer<typeof EntrySchema>
export type KnowledgeData = z.infer<typeof KnowledgeSchema>

const defaultPath = path.join(
  path.dirname(KNOWLEDGE_FILE),
  'knowledge.default.json',
)

async function ensureFile(): Promise<void> {
  try {
    await fs.access(KNOWLEDGE_FILE)
  } catch {
    const raw = await fs.readFile(defaultPath, 'utf-8')
    await fs.writeFile(KNOWLEDGE_FILE, raw, 'utf-8')
  }
}

export async function readKnowledge(): Promise<KnowledgeData> {
  await ensureFile()
  const raw = await fs.readFile(KNOWLEDGE_FILE, 'utf-8')
  return KnowledgeSchema.parse(JSON.parse(raw))
}

export async function writeKnowledge(data: KnowledgeData): Promise<void> {
  KnowledgeSchema.parse(data)
  await fs.mkdir(path.dirname(KNOWLEDGE_FILE), { recursive: true })
  await fs.writeFile(KNOWLEDGE_FILE, JSON.stringify(data, null, 2), 'utf-8')
}

export function knowledgeToPromptBlock(data: KnowledgeData): string {
  const lines = data.entries.map(
    (e) =>
      `- [${e.id}] ${e.title}\n  Palabras clave: ${e.keywords.join(', ')}\n  Respuesta sugerida: ${e.response}`,
  )
  const sla = Object.entries(data.slaHoursByPriority)
    .map(([k, v]) => `${k}: ${v} h`)
    .join('; ')
  return `Base de conocimiento (respuestas parametrizadas — úsalas cuando apliquen):\n${lines.join('\n\n')}\n\nANS referencia (horas de primera respuesta por prioridad): ${sla}`
}
