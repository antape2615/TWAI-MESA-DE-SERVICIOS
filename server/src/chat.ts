import OpenAI from 'openai'
import { readKnowledge, knowledgeToPromptBlock } from './knowledge.js'
import { createTicket } from './tickets.js'
import { sendTicketCreatedEmail, type EmailSendResult } from './email.js'
import type { ChatCompletionMessageParam } from 'openai/resources/chat/completions'

export type TicketDraftPayload = {
  title: string
  description: string
  category: string
  priority: 'baja' | 'media' | 'alta' | 'critica'
  possibleSolutions: string[]
}

const ticketParams = {
  type: 'object',
  properties: {
    title: { type: 'string', description: 'Título breve del incidente' },
    description: {
      type: 'string',
      description: 'Detalle del problema y pasos ya intentados',
    },
    category: {
      type: 'string',
      enum: ['hardware', 'software', 'red', 'acceso', 'otro'],
    },
    priority: {
      type: 'string',
      enum: ['baja', 'media', 'alta', 'critica'],
      description: 'Impacto en el negocio',
    },
    possible_solutions: {
      type: 'array',
      items: { type: 'string' },
      description: 'Pasos o soluciones que se podrían intentar mientras tanto',
    },
  },
  required: ['title', 'description', 'priority'],
} as const

const proposeTicketTool = {
  type: 'function' as const,
  function: {
    name: 'propose_ticket',
    description:
      'Úsala cuando convenga abrir un ticket (escalamiento, sin solución clara, o el usuario lo pide) y quieras que confirme con el botón en pantalla antes de crearlo. No crea el ticket; solo prepara el borrador para la interfaz.',
    parameters: ticketParams,
  },
}

const createTicketTool = {
  type: 'function' as const,
  function: {
    name: 'create_ticket',
    description:
      'Crea el ticket de inmediato solo si el usuario confirmó por texto que desea crearlo ya (p. ej. «sí, créalo», «genera el ticket») y no basta con el botón. Si hubo imagen, resume en la descripción lo visible.',
    parameters: ticketParams,
  },
}

function getClient(): OpenAI {
  const endpoint = (process.env.AZURE_OPENAI_ENDPOINT ?? '').replace(/\/$/, '')
  const deployment = process.env.AZURE_OPENAI_DEPLOYMENT_NAME
  const apiKey = process.env.AZURE_OPENAI_API_KEY
  const apiVersion =
    process.env.AZURE_OPENAI_API_VERSION ?? '2024-08-01-preview'
  if (!endpoint || !deployment || !apiKey) {
    throw new Error(
      'Faltan variables AZURE_OPENAI_ENDPOINT, AZURE_OPENAI_DEPLOYMENT_NAME o AZURE_OPENAI_API_KEY',
    )
  }
  return new OpenAI({
    apiKey,
    baseURL: `${endpoint}/openai/deployments/${deployment}`,
    defaultQuery: { 'api-version': apiVersion },
    defaultHeaders: { 'api-key': apiKey },
  })
}

function priorityFromString(
  p: string,
): 'baja' | 'media' | 'alta' | 'critica' {
  if (p === 'critica' || p === 'alta' || p === 'media' || p === 'baja') return p
  return 'media'
}

function parseTicketArgs(args: Record<string, unknown>): TicketDraftPayload {
  return {
    title: String(args.title ?? 'Sin título'),
    description: String(args.description ?? ''),
    category: String(args.category ?? 'otro'),
    priority: priorityFromString(String(args.priority ?? 'media')),
    possibleSolutions: Array.isArray(args.possible_solutions)
      ? (args.possible_solutions as unknown[]).map(String)
      : [],
  }
}

export async function runChat(params: {
  messages: ChatCompletionMessageParam[]
  userEmail?: string
}): Promise<{
  message: string
  ticketId?: string
  ticketDraft?: TicketDraftPayload
  email?: EmailSendResult
}> {
  const knowledge = await readKnowledge()
  const kbBlock = knowledgeToPromptBlock(knowledge)
  const system: ChatCompletionMessageParam = {
    role: 'system',
    content: `Eres el asistente de Mesa de Servicios de Periferia. Idioma: español.
Sé breve y empático. Primero intenta resolver con la base de conocimiento siguiente.
Si el usuario envía una imagen o captura, analízala (texto visible, códigos de error, ventanas) y propón solución concreta.
Cuando el problema parezca un error en pantalla, mensaje del sistema, código o interfaz y aún NO conste ninguna imagen en los mensajes del usuario sobre ese incidente, pide amablemente una captura de pantalla antes de proponer ticket; explica que puede usar «Captura del error» o pegar con Ctrl+V. Si el usuario indica que no puede adjuntar imagen, continúa con lo que tengas.
Cuando haga falta escalamiento o un ticket, usa por defecto la herramienta **propose_ticket**: el usuario confirmará con el botón «Sí, generar ticket» en la interfaz. Explica brevemente que puede pulsar el botón para crear el ticket.
Usa **create_ticket** solo si el usuario escribió de forma explícita que quieres crear el ticket ya en este mensaje (p. ej. confirma sin ambigüedad tras haber visto la propuesta).
Indica el ANS (horas de primera respuesta) según la prioridad usando la tabla de la base.
No inventes datos de sistemas internos; si no sabes, propón ticket o pide más detalle.

${kbBlock}`,
  }

  const client = getClient()
  const deployment = process.env.AZURE_OPENAI_DEPLOYMENT_NAME!

  let messages: ChatCompletionMessageParam[] = [system, ...params.messages]
  let lastTicketId: string | undefined
  let lastTicketDraft: TicketDraftPayload | undefined
  let lastEmailResult: EmailSendResult | undefined

  for (let round = 0; round < 8; round++) {
    const completion = await client.chat.completions.create({
      model: deployment,
      messages,
      tools: [proposeTicketTool, createTicketTool],
      tool_choice: 'auto',
      temperature: 0.4,
    })

    const choice = completion.choices[0]
    const msg = choice?.message
    if (!msg) {
      return { message: 'No hubo respuesta del modelo.' }
    }

    if (msg.tool_calls?.length) {
      messages.push({
        role: 'assistant',
        content: msg.content,
        tool_calls: msg.tool_calls,
      })

      for (const call of msg.tool_calls) {
        if (call.type !== 'function') {
          continue
        }
        const name = call.function.name
        let args: Record<string, unknown>
        try {
          args = JSON.parse(call.function.arguments || '{}') as Record<
            string,
            unknown
          >
        } catch {
          messages.push({
            role: 'tool',
            tool_call_id: call.id,
            content: JSON.stringify({ error: 'JSON inválido' }),
          })
          continue
        }

        if (name === 'propose_ticket') {
          const draft = parseTicketArgs(args)
          lastTicketDraft = draft
          messages.push({
            role: 'tool',
            tool_call_id: call.id,
            content: JSON.stringify({
              ok: true,
              paso: 'Se mostrarán botones de confirmación en la interfaz; el usuario puede crear el ticket desde ahí.',
            }),
          })
          continue
        }

        if (name === 'create_ticket') {
          const d = parseTicketArgs(args)
          const ansHours =
            knowledge.slaHoursByPriority[d.priority] ??
            knowledge.slaHoursByPriority['media'] ??
            24

          const ticket = await createTicket({
            title: d.title,
            description: d.description,
            category: d.category,
            priority: d.priority,
            ansHours,
            possibleSolutions: d.possibleSolutions,
            userEmail: params.userEmail,
          })
          lastTicketId = ticket.id
          lastTicketDraft = undefined
          lastEmailResult = await sendTicketCreatedEmail(ticket)

          messages.push({
            role: 'tool',
            tool_call_id: call.id,
            content: JSON.stringify({
              ok: true,
              ticket_id: ticket.id,
              ans_hours: ticket.ansHours,
              priority: ticket.priority,
            }),
          })
          continue
        }

        messages.push({
          role: 'tool',
          tool_call_id: call.id,
          content: JSON.stringify({ error: 'Herramienta no soportada' }),
        })
      }
      continue
    }

    const text = msg.content?.trim() || ''
    return {
      message: text,
      ticketId: lastTicketId,
      ticketDraft: lastTicketDraft,
      ...(lastEmailResult !== undefined ? { email: lastEmailResult } : {}),
    }
  }

  return {
    message:
      'Se alcanzó el límite de pasos en la conversación. Intente de nuevo o abra un ticket desde el portal.',
  }
}
