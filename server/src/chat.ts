import OpenAI from 'openai'
import { readKnowledge, knowledgeToPromptBlock } from './knowledge.js'
import {
  collectUserText,
  lastUserMessagePlainText,
  selectRelevantFaqs,
  faqsToPromptBlock,
} from './faqMatch.js'
import { createTicket } from './tickets.js'
import {
  emailNotificationsEnabled,
  sendTicketCreatedEmail,
  type EmailSendResult,
} from './email.js'
import { ticketsFromPortalEnabled } from './features.js'
import {
  getSharePointListUrl,
  hasSharePointRequesterIdentity,
  sharePointTicketsEnabled,
} from './sharepoint.js'
import { azureAuthEnabled } from './azureAuth.js'
import {
  buildFallbackDeepLink,
  buildPowerAppsDeepLink,
  hasHelpdeskPowerAppsUrl,
  parseHelpdeskLinkArgs,
} from './helpdeskTemplate.js'
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

const helpdeskLinkParams = {
  type: 'object',
  properties: {
    titulo: {
      type: 'string',
      description: 'Título o asunto del incidente (resumen breve para mesa de ayuda)',
    },
    descripcion: {
      type: 'string',
      description: 'Descripción detallada: síntomas, errores y pasos ya probados',
    },
    categoria: {
      type: 'string',
      description: 'Categoría (hardware, software, red, acceso, otro, etc.)',
    },
    solicitado_por: {
      type: 'string',
      description: 'Correo corporativo o nombre completo de quien solicita el soporte',
    },
    numero_contacto: {
      type: 'string',
      description: 'Teléfono u otro número de contacto',
    },
    pais: { type: 'string', description: 'País o ubicación geográfica relevante' },
    departamento: {
      type: 'string',
      description: 'Departamento, área o dirección dentro de la organización',
    },
    torre: {
      type: 'string',
      description: 'Torre, sede u oficina física si aplica',
    },
  },
} as const

const openHelpdeskLinkTool = {
  type: 'function' as const,
  function: {
    name: 'open_helpdesk_link',
    description:
      'Al escalar a HelpDesk, infiere titulo, descripcion, categoria, etc. y obtén la URL de Power Apps. El usuario verá un botón para abrir esa URL; no hace falta listar campos del formulario en el mensaje.',
    parameters: helpdeskLinkParams,
  },
}

function azureOpenAIDeployment(): string | undefined {
  return (
    process.env.AZURE_OPENAI_DEPLOYMENT_NAME?.trim() ||
    process.env.AZURE_OPENAI_DEPLOYMENT?.trim() ||
    undefined
  )
}

function getClient(): OpenAI {
  const endpoint = (process.env.AZURE_OPENAI_ENDPOINT ?? '').replace(/\/$/, '')
  const deployment = azureOpenAIDeployment()
  const apiKey = process.env.AZURE_OPENAI_API_KEY
  const apiVersion =
    process.env.AZURE_OPENAI_API_VERSION ?? '2024-08-01-preview'
  if (!endpoint || !deployment || !apiKey) {
    throw new Error(
      'Faltan variables AZURE_OPENAI_ENDPOINT, AZURE_OPENAI_DEPLOYMENT (o AZURE_OPENAI_DEPLOYMENT_NAME) o AZURE_OPENAI_API_KEY',
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

function resolveHelpdeskUrlForResponse(
  toolUrl: string | undefined,
  portalTickets: boolean,
  helpdeskConfigured: boolean,
  requestMessages: ChatCompletionMessageParam[],
  userEmail?: string,
  userName?: string,
): string | undefined {
  if (portalTickets || !helpdeskConfigured) return undefined
  if (toolUrl) return toolUrl
  return (
    buildFallbackDeepLink(
      lastUserMessagePlainText(requestMessages),
      userEmail,
      userName,
    ) ?? undefined
  )
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
  userName?: string
  jobTitle?: string
  department?: string
  officeLocation?: string
  phone?: string
  accessToken?: string
  sharePointAccessToken?: string
}): Promise<{
  message: string
  ticketId?: string
  ticketDraft?: TicketDraftPayload
  email?: EmailSendResult
  helpdeskUrl?: string
}> {
  const knowledge = await readKnowledge()
  const kbBlock = knowledgeToPromptBlock(knowledge)
  const userBlob = collectUserText(params.messages)
  const relevantFaqs = await selectRelevantFaqs(userBlob, 3)
  const faqBlock = faqsToPromptBlock(relevantFaqs)
  const noFaqMatch = relevantFaqs.length === 0
  const portalTickets = ticketsFromPortalEnabled()
  const sharePointTickets = sharePointTicketsEnabled()
  const helpdeskDeepLink = hasHelpdeskPowerAppsUrl()
  const sharePointListUrl = getSharePointListUrl()

  const ticketDestination = sharePointTickets
    ? 'la lista de SharePoint de Mesa de Servicios (RPA_SOLICITUD_TICKET_SOPORTE)'
    : 'el sistema de tickets del portal'

  const requesterKnown = hasSharePointRequesterIdentity({
    userEmail: params.userEmail,
    userName: params.userName,
  })
  const loginRequiredForTickets =
    sharePointTickets && azureAuthEnabled() && !requesterKnown

  const ticketPolicyPortal = `Cuando haga falta escalamiento o un ticket, usa por defecto la herramienta **propose_ticket**: el usuario confirmará con el botón «Sí, generar ticket» en la interfaz. El ticket se registrará en ${ticketDestination}. Explica brevemente que puede pulsar el botón para crear el ticket con los datos que la IA preparó.
Usa **create_ticket** solo si el usuario escribió de forma explícita que quiere crear el ticket ya en este mensaje (p. ej. confirma sin ambigüedad tras haber visto la propuesta).
${loginRequiredForTickets ? '**IMPORTANTE:** el usuario NO ha iniciado sesión con Microsoft. NO llames a propose_ticket ni create_ticket; pídele que pulse «Iniciar sesión con Microsoft» arriba. Sin sesión no se puede rellenar «Solicitado Por» en SharePoint.' : ''}
${requesterKnown ? `Usuario autenticado: ${params.userName ?? '—'} (${params.userEmail ?? 'sin correo'}). Al crear ticket, Solicitado Por debe ser este usuario.` : ''}`

  const noFaqTicketPolicy =
    noFaqMatch && portalTickets
      ? `La consulta del usuario **no encaja con ninguna FAQ oficial**. Revisa la base de conocimiento; si no hay solución clara, el problema requiere intervención humana o el usuario pide escalamiento, usa **propose_ticket** para que confirme la creación del ticket en ${ticketDestination}. No inventes pasos técnicos que no estén en la base.`
      : ''

  const helpdeskLinkInstruction = helpdeskDeepLink
    ? `
Si hace falta escalar a HelpDesk, llama a **open_helpdesk_link** con los datos que puedas inferir; el usuario tendrá un botón con el enlace. No repitas en el texto listados de campos del formulario ni plantillas para rellenar a mano.`
    : ''

  const ticketPolicyHelpdesk = `NO hay creación de tickets en este portal: está desactivada.
Cuando convenga escalamiento o registrar el caso en mesa de ayuda, orienta con brevedad hacia **HelpDesk Periferia** (Power Apps — «Nuevo Ticket»).${helpdeskLinkInstruction}
No incluyas en tu mensaje bloques tipo «plantilla para copiar», listados campo a campo ni texto pensado para pegar en formularios externos.
Puedes mencionar de forma breve el ANS referencial según la gravedad usando la tabla de la base; el registro oficial es en HelpDesk.`

  const systemParts = [
    `Eres el asistente de Mesa de Servicios de Periferia. Idioma: español.
Sé breve y empático. Si en tu contexto aparecen **FAQ oficiales** y la consulta del usuario encaja con ese tema, basa la respuesta principalmente en ellas (resumen o pasos en lista) sin contradecir esa guía.
Para el resto, usa la base de conocimiento parametrizada que viene después de las FAQ (si las hay).
Si el usuario envía una imagen o captura, analízala (texto visible, códigos de error, ventanas) y propón solución concreta.
Cuando el problema parezca un error en pantalla, mensaje del sistema, código o interfaz y aún NO conste ninguna imagen en los mensajes del usuario sobre ese incidente, pide amablemente una captura de pantalla antes de sugerir escalamiento; explica que puede usar «Captura del error» o pegar con Ctrl+V. Si el usuario indica que no puede adjuntar imagen, continúa con lo que tengas.
${noFaqTicketPolicy}
${portalTickets ? ticketPolicyPortal : ticketPolicyHelpdesk}
${portalTickets ? 'Indica el ANS (horas de primera respuesta) según la prioridad usando la tabla de la base.' : ''}
${sharePointListUrl ? `Tras crear un ticket, puedes mencionar que quedó registrado en SharePoint (lista de soporte).` : ''}
No inventes datos de sistemas internos; si no sabes, pide más detalle o orienta con la plantilla HelpDesk.`,
  ]
  if (faqBlock) systemParts.push(faqBlock)
  systemParts.push(kbBlock)

  const system: ChatCompletionMessageParam = {
    role: 'system',
    content: systemParts.join('\n\n'),
  }

  const client = getClient()
  const deployment = azureOpenAIDeployment()!

  let messages: ChatCompletionMessageParam[] = [system, ...params.messages]
  let lastTicketId: string | undefined
  let lastTicketDraft: TicketDraftPayload | undefined
  let lastEmailResult: EmailSendResult | undefined
  let lastHelpdeskUrl: string | undefined

  const toolsHelpdeskOnly = !portalTickets && helpdeskDeepLink ? [openHelpdeskLinkTool] : undefined

  for (let round = 0; round < 8; round++) {
    const completion = portalTickets
      ? await client.chat.completions.create({
          model: deployment,
          messages,
          tools: [proposeTicketTool, createTicketTool],
          tool_choice: 'auto',
          temperature: 0.4,
        })
      : toolsHelpdeskOnly
        ? await client.chat.completions.create({
            model: deployment,
            messages,
            tools: toolsHelpdeskOnly,
            tool_choice: 'auto',
            temperature: 0.4,
          })
        : await client.chat.completions.create({
            model: deployment,
            messages,
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
          if (
            sharePointTickets &&
            azureAuthEnabled() &&
            !hasSharePointRequesterIdentity({
              userEmail: params.userEmail,
              userName: params.userName,
            })
          ) {
            messages.push({
              role: 'tool',
              tool_call_id: call.id,
              content: JSON.stringify({
                ok: false,
                error:
                  'El usuario debe iniciar sesión con Microsoft antes de crear el ticket (Solicitado Por en SharePoint).',
              }),
            })
            continue
          }

          const d = parseTicketArgs(args)
          const ansHours =
            knowledge.slaHoursByPriority[d.priority] ??
            knowledge.slaHoursByPriority['media'] ??
            24

          let ticket
          try {
            ticket = await createTicket({
            title: d.title,
            description: d.description,
            category: d.category,
            priority: d.priority,
            ansHours,
            possibleSolutions: d.possibleSolutions,
            userEmail: params.userEmail,
            userName: params.userName,
            jobTitle: params.jobTitle,
            department: params.department,
            officeLocation: params.officeLocation,
            phone: params.phone,
            accessToken: params.accessToken,
            sharePointAccessToken: params.sharePointAccessToken,
          })
          } catch (err) {
            const msg =
              err instanceof Error ? err.message : 'No se pudo crear el ticket'
            messages.push({
              role: 'tool',
              tool_call_id: call.id,
              content: JSON.stringify({ ok: false, error: msg }),
            })
            continue
          }
          lastTicketId = ticket.id
          lastTicketDraft = undefined
          lastEmailResult = emailNotificationsEnabled()
            ? await sendTicketCreatedEmail(ticket)
            : undefined

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

        if (name === 'open_helpdesk_link') {
          const fields = parseHelpdeskLinkArgs(args)
          const url = buildPowerAppsDeepLink(fields)
          if (!url) {
            messages.push({
              role: 'tool',
              tool_call_id: call.id,
              content: JSON.stringify({
                ok: false,
                error: 'HELPDESK_POWERAPPS_URL no está configurada en el servidor',
              }),
            })
            continue
          }
          lastHelpdeskUrl = url
          messages.push({
            role: 'tool',
            tool_call_id: call.id,
            content: JSON.stringify({
              ok: true,
              url,
              campos_usados: fields,
              instruction:
                'Incluye la URL en tu respuesta breve. El usuario verá el botón para abrir HelpDesk.',
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

    const rawText = msg.content?.trim() ?? ''
    const helpdeskUrl = resolveHelpdeskUrlForResponse(
      lastHelpdeskUrl,
      portalTickets,
      helpdeskDeepLink,
      params.messages,
      params.userEmail,
      params.userName,
    )
    const text =
      rawText ||
      (helpdeskUrl
        ? 'Puede abrir HelpDesk con el botón inferior para continuar en Power Apps.'
        : '')
    return {
      message: text,
      ticketId: lastTicketId,
      ticketDraft: lastTicketDraft,
      ...(lastEmailResult !== undefined ? { email: lastEmailResult } : {}),
      ...(helpdeskUrl ? { helpdeskUrl } : {}),
    }
  }

  const helpdeskUrlFallback = resolveHelpdeskUrlForResponse(
    lastHelpdeskUrl,
    portalTickets,
    helpdeskDeepLink,
    params.messages,
    params.userEmail,
    params.userName,
  )
  return {
    message: portalTickets
      ? 'Se alcanzó el límite de pasos en la conversación. Intente de nuevo o use HelpDesk para registrar el caso.'
      : 'Se alcanzó el límite de pasos en la conversación. Intente de nuevo o registre el caso en HelpDesk si aplica.',
    ...(helpdeskUrlFallback ? { helpdeskUrl: helpdeskUrlFallback } : {}),
  }
}
