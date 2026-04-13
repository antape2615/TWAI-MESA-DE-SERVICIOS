import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import './loadEnv.js'
import express from 'express'
import cors from 'cors'
import { z } from 'zod'
import { runChat } from './chat.js'
import { ticketsFromPortalEnabled } from './features.js'
import { hasHelpdeskPowerAppsUrl } from './helpdeskTemplate.js'
import { createTicket, listTickets } from './tickets.js'
import { logEmailStartupHint, sendTicketCreatedEmail } from './email.js'
import { readKnowledge, writeKnowledge, type KnowledgeData } from './knowledge.js'
import type { ChatCompletionMessageParam } from 'openai/resources/chat/completions'

const app = express()
app.use(cors({ origin: true }))
app.use(express.json({ limit: '20mb' }))

app.get('/api/health', (_req, res) => {
  res.json({ ok: true })
})

app.get('/api/config', (_req, res) => {
  res.json({
    ticketsFromPortal: ticketsFromPortalEnabled(),
    helpdeskDeepLink: hasHelpdeskPowerAppsUrl(),
  })
})

app.get('/api/tickets', async (_req, res) => {
  try {
    const tickets = await listTickets()
    res.json({ tickets })
  } catch (e) {
    console.error(e)
    res.status(500).json({ error: 'No se pudieron cargar los tickets' })
  }
})

const CreateTicketBodySchema = z.object({
  title: z.string(),
  description: z.string(),
  category: z.string(),
  priority: z.enum(['baja', 'media', 'alta', 'critica']),
  possibleSolutions: z.array(z.string()),
  userEmail: z.union([z.string().email(), z.literal('')]).optional(),
})

app.post('/api/tickets', async (req, res) => {
  try {
    if (!ticketsFromPortalEnabled()) {
      res.status(403).json({
        error:
          'La creación de tickets desde esta aplicación está deshabilitada. Use HelpDesk y la plantilla que le indique el asistente.',
      })
      return
    }
    const parsed = CreateTicketBodySchema.safeParse(req.body)
    if (!parsed.success) {
      res.status(400).json({ error: 'Datos inválidos' })
      return
    }
    const knowledge = await readKnowledge()
    const { title, description, category, priority, possibleSolutions, userEmail } =
      parsed.data
    const ansHours =
      knowledge.slaHoursByPriority[priority] ??
      knowledge.slaHoursByPriority['media'] ??
      24
    const ticket = await createTicket({
      title,
      description,
      category,
      priority,
      ansHours,
      possibleSolutions,
      userEmail: userEmail || undefined,
    })
    const email = await sendTicketCreatedEmail(ticket)
    res.json({ ticket, email })
  } catch (e) {
    console.error(e)
    res.status(500).json({ error: 'No se pudo crear el ticket' })
  }
})

app.get('/api/knowledge', async (_req, res) => {
  try {
    const data = await readKnowledge()
    res.json(data)
  } catch (e) {
    console.error(e)
    res.status(500).json({ error: 'No se pudo leer la base de conocimiento' })
  }
})

const PutKnowledgeSchema = z.object({
  entries: z.array(
    z.object({
      id: z.string(),
      keywords: z.array(z.string()),
      title: z.string(),
      response: z.string(),
    }),
  ),
  slaHoursByPriority: z.record(z.string(), z.number()),
})

app.put('/api/knowledge', async (req, res) => {
  try {
    const parsed = PutKnowledgeSchema.safeParse(req.body)
    if (!parsed.success) {
      res.status(400).json({ error: 'Payload inválido', details: parsed.error.flatten() })
      return
    }
    await writeKnowledge(parsed.data as KnowledgeData)
    res.json({ ok: true })
  } catch (e) {
    console.error(e)
    res.status(500).json({ error: 'No se pudo guardar' })
  }
})

const allowedImageMime = new Set([
  'image/png',
  'image/jpeg',
  'image/jpg',
  'image/gif',
  'image/webp',
])

const ChatSchema = z.object({
  messages: z.array(
    z.object({
      role: z.enum(['user', 'assistant', 'system']),
      content: z.string(),
      image: z
        .object({
          mimeType: z.string(),
          dataBase64: z.string(),
        })
        .optional(),
    }),
  ),
  userEmail: z.union([z.string().email(), z.literal('')]).optional(),
})

function toOpenAIMessages(
  messages: z.infer<typeof ChatSchema>['messages'],
): ChatCompletionMessageParam[] {
  const maxB64 = 18_000_000
  return messages.map((m) => {
    if (m.role === 'user' && m.image) {
      if (!allowedImageMime.has(m.image.mimeType.toLowerCase())) {
        throw new Error(
          `Tipo de imagen no permitido: ${m.image.mimeType}. Use PNG, JPEG, GIF o WebP.`,
        )
      }
      if (m.image.dataBase64.length > maxB64) {
        throw new Error(
          'La imagen es demasiado grande. Pruebe con una captura más pequeña o mayor compresión.',
        )
      }
      const text =
        m.content.trim() ||
        'El usuario adjunta una captura de pantalla del error (analízala).'
      return {
        role: 'user' as const,
        content: [
          { type: 'text' as const, text },
          {
            type: 'image_url' as const,
            image_url: {
              url: `data:${m.image.mimeType};base64,${m.image.dataBase64}`,
            },
          },
        ],
      }
    }
    return { role: m.role, content: m.content }
  })
}

app.post('/api/chat', async (req, res) => {
  try {
    const parsed = ChatSchema.safeParse(req.body)
    if (!parsed.success) {
      res.status(400).json({ error: 'Mensajes inválidos' })
      return
    }
    const { messages, userEmail } = parsed.data
    let openaiMessages: ChatCompletionMessageParam[]
    try {
      openaiMessages = toOpenAIMessages(messages)
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'Petición inválida'
      res.status(400).json({ error: msg })
      return
    }
    const result = await runChat({
      messages: openaiMessages,
      userEmail: userEmail || undefined,
    })
    res.json(result)
  } catch (e) {
    console.error(e)
    const msg = e instanceof Error ? e.message : 'Error en el chat'
    res.status(500).json({ error: msg })
  }
})

/** En producción (Render, etc.): servir el build de Vite desde el mismo origen. */
const distPath = path.join(
  path.dirname(fileURLToPath(import.meta.url)),
  '../../frontend/dist',
)
if (fs.existsSync(distPath)) {
  app.use(express.static(distPath))
  app.get('*', (req, res) => {
    if (req.path.startsWith('/api')) {
      res.status(404).json({ error: 'No encontrado' })
      return
    }
    res.sendFile(path.join(distPath, 'index.html'))
  })
}

const port = Number(process.env.PORT ?? 8787)
app.listen(port, () => {
  console.log(`Mesa Servicios en http://localhost:${port}`)
  logEmailStartupHint()
})
