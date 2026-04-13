export type ChatImagePayload = { mimeType: string; dataBase64: string }

export type TicketDraft = {
  title: string
  description: string
  category: string
  priority: 'baja' | 'media' | 'alta' | 'critica'
  possibleSolutions: string[]
}

export type ChatMessage = {
  role: 'user' | 'assistant' | 'system'
  content: string
  image?: ChatImagePayload
  /** Borrador ofrecido por el modelo; el usuario confirma con el botón */
  ticketDraft?: TicketDraft
  /** Tras confirmar desde la UI */
  ticketCreatedId?: string
  /** Enlace Power Apps HelpDesk con parámetros (si el servidor lo generó) */
  helpdeskUrl?: string
}

export type KnowledgeEntry = {
  id: string
  keywords: string[]
  title: string
  response: string
}

export type KnowledgePayload = {
  entries: KnowledgeEntry[]
  slaHoursByPriority: Record<string, number>
}

export type Ticket = {
  id: string
  title: string
  description: string
  category: string
  priority: 'baja' | 'media' | 'alta' | 'critica'
  status: string
  ansHours: number
  possibleSolutions: string[]
  createdAt: string
  userEmail?: string
}

export type EmailSendResult = {
  sent: boolean
  error?: string
  method?: 'resend' | 'smtp' | 'none'
}

export type AppConfig = {
  ticketsFromPortal: boolean
  /** True si HELPDESK_POWERAPPS_URL está definida (enlace profundo disponible) */
  helpdeskDeepLink: boolean
  /** URL play de HelpDesk para abrir la pantalla desde el chat */
  helpdeskPowerAppsUrl?: string
}

async function parseJson<T>(res: Response): Promise<T> {
  const text = await res.text()
  if (!res.ok) {
    let msg = res.statusText
    try {
      const j = JSON.parse(text) as { error?: string }
      if (j.error) msg = j.error
    } catch {
      /* ignore */
    }
    throw new Error(msg)
  }
  return JSON.parse(text) as T
}

export async function fetchAppConfig(): Promise<AppConfig> {
  const res = await fetch('/api/config')
  return parseJson(res)
}

export type ChatApiResponse = {
  message: string
  ticketId?: string
  ticketDraft?: TicketDraft
  email?: EmailSendResult
  helpdeskUrl?: string
}

export async function postChat(
  messages: ChatMessage[],
  userEmail?: string,
): Promise<ChatApiResponse> {
  const res = await fetch('/api/chat', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ messages, userEmail: userEmail || '' }),
  })
  return parseJson(res)
}

export async function createTicketConfirm(
  draft: TicketDraft,
  userEmail?: string,
): Promise<{ ticket: Ticket; email: EmailSendResult }> {
  const res = await fetch('/api/tickets', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      title: draft.title,
      description: draft.description,
      category: draft.category,
      priority: draft.priority,
      possibleSolutions: draft.possibleSolutions,
      userEmail: userEmail || '',
    }),
  })
  return parseJson(res)
}

const MAX_IMAGE_BYTES = 8 * 1024 * 1024

export async function fileToImagePayload(file: File): Promise<ChatImagePayload> {
  if (file.size > MAX_IMAGE_BYTES) {
    throw new Error('La imagen supera 8 MB. Reduzca el tamaño o use otra captura.')
  }
  const mime = file.type || 'image/png'
  if (!/^image\/(png|jpeg|jpg|gif|webp)$/i.test(mime)) {
    throw new Error('Formato no admitido. Use PNG, JPEG, GIF o WebP.')
  }
  const dataBase64 = await new Promise<string>((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => {
      const r = reader.result
      if (typeof r !== 'string') {
        reject(new Error('No se pudo leer el archivo'))
        return
      }
      const i = r.indexOf(',')
      resolve(i >= 0 ? r.slice(i + 1) : r)
    }
    reader.onerror = () => reject(new Error('Error al leer el archivo'))
    reader.readAsDataURL(file)
  })
  return { mimeType: mime.toLowerCase().replace('image/jpg', 'image/jpeg'), dataBase64 }
}

export async function fetchTickets(): Promise<{ tickets: Ticket[] }> {
  const res = await fetch('/api/tickets')
  return parseJson(res)
}

export async function fetchKnowledge(): Promise<KnowledgePayload> {
  const res = await fetch('/api/knowledge')
  return parseJson(res)
}

export async function saveKnowledge(data: KnowledgePayload): Promise<void> {
  const res = await fetch('/api/knowledge', {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(data),
  })
  await parseJson(res)
}
