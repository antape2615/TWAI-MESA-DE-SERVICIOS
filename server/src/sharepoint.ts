import type { Ticket } from './tickets.js'

type GraphToken = { accessToken: string; expiresAt: number }

let cachedToken: GraphToken | null = null
let cachedSiteId: string | null = null
let cachedListId: string | null = null

/** Valores de la lista lookup «Prioridad» (id → título). */
const PRIORITY_LOOKUP_IDS: Record<string, number> = {
  critica: 1, // URGENTE
  alta: 2, // ALTO
  media: 3, // MEDIO
  baja: 4, // BAJO
}

/** Categoría del chat → lista lookup «Categoria». */
const CATEGORY_LOOKUP_IDS: Record<string, number> = {
  hardware: 13, // PROBLEMA CON EQUIPO FÍSICO
  software: 14, // PROBLEMA CON PROGRAMAS O APLICACIONES
  red: 10, // PROBLEMAS CON INTERNET
  acceso: 18, // ACCESO CUENTA
  otro: 1, // ADMINISTRACIÓN DE SERVICIOS
}

const DEFAULT_ESTADO_LOOKUP_ID = 1 // Abierto

function env(name: string): string | undefined {
  const v = process.env[name]?.trim()
  return v || undefined
}

function fieldName(key: string, fallback: string): string {
  return env(`SHAREPOINT_FIELD_${key}`) ?? fallback
}

export function sharePointTicketsEnabled(): boolean {
  return Boolean(
    env('AZURE_TENANT_ID') &&
      env('SHAREPOINT_CLIENT_ID') &&
      env('SHAREPOINT_CLIENT_SECRET') &&
      env('SHAREPOINT_SITE_URL') &&
      env('SHAREPOINT_LIST_NAME'),
  )
}

export function logSharePointStartupHint(): void {
  if (!sharePointTicketsEnabled()) {
    console.warn(
      '[sharepoint] Tickets en SharePoint desactivados: defina AZURE_TENANT_ID, SHAREPOINT_CLIENT_ID, SHAREPOINT_CLIENT_SECRET, SHAREPOINT_SITE_URL y SHAREPOINT_LIST_NAME (ver .env.example)',
    )
    return
  }
  const listUrl = getSharePointListUrl()
  console.log(
    `[sharepoint] Tickets se crearán en la lista ${env('SHAREPOINT_LIST_NAME')}${listUrl ? ` (${listUrl})` : ''}`,
  )
}

export function getSharePointListUrl(): string | null {
  const explicit = env('SHAREPOINT_LIST_URL')
  if (explicit?.startsWith('http')) return explicit
  const site = env('SHAREPOINT_SITE_URL')
  const list = env('SHAREPOINT_LIST_NAME')
  if (!site || !list) return null
  const base = site.replace(/\/$/, '')
  return `${base}/Lists/${encodeURIComponent(list)}/AllItems.aspx`
}

async function getGraphToken(): Promise<string> {
  const now = Date.now()
  if (cachedToken && cachedToken.expiresAt > now + 60_000) {
    return cachedToken.accessToken
  }

  const tenantId = env('AZURE_TENANT_ID')
  const clientId = env('SHAREPOINT_CLIENT_ID')
  const clientSecret = env('SHAREPOINT_CLIENT_SECRET')
  if (!tenantId || !clientId || !clientSecret) {
    throw new Error('Faltan credenciales de SharePoint (tenant, client id o secret)')
  }

  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials',
  })

  const res = await fetch(
    `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body,
    },
  )

  const data = (await res.json().catch(() => ({}))) as {
    access_token?: string
    expires_in?: number
    error_description?: string
    error?: string
  }

  if (!res.ok || !data.access_token) {
    const msg =
      data.error_description || data.error || res.statusText || 'Error de autenticación Graph'
    throw new Error(`No se pudo autenticar en Microsoft Graph: ${msg}`)
  }

  const expiresIn = typeof data.expires_in === 'number' ? data.expires_in : 3600
  cachedToken = {
    accessToken: data.access_token,
    expiresAt: now + expiresIn * 1000,
  }
  return cachedToken.accessToken
}

async function graphGet<T>(path: string): Promise<T> {
  const token = await getGraphToken()
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  })
  const data = (await res.json().catch(() => ({}))) as T & {
    error?: { message?: string }
  }
  if (!res.ok) {
    const msg = data.error?.message || res.statusText || 'Error Graph API'
    throw new Error(msg)
  }
  return data
}

async function graphPost<T>(path: string, body: unknown): Promise<T> {
  const token = await getGraphToken()
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(body),
  })
  const data = (await res.json().catch(() => ({}))) as T & {
    error?: { message?: string }
  }
  if (!res.ok) {
    const msg = data.error?.message || res.statusText || 'Error Graph API'
    throw new Error(msg)
  }
  return data
}

function sitePathFromUrl(siteUrl: string): string {
  try {
    const u = new URL(siteUrl)
    const host = u.hostname
    const path = u.pathname.replace(/\/$/, '') || '/'
    return `/sites/${host}:${path}`
  } catch {
    throw new Error(`SHAREPOINT_SITE_URL inválida: ${siteUrl}`)
  }
}

async function resolveSiteId(): Promise<string> {
  if (cachedSiteId) return cachedSiteId
  const siteUrl = env('SHAREPOINT_SITE_URL')
  if (!siteUrl) throw new Error('SHAREPOINT_SITE_URL no configurada')
  const site = await graphGet<{ id: string }>(sitePathFromUrl(siteUrl))
  cachedSiteId = site.id
  return site.id
}

async function resolveListId(siteId: string): Promise<string> {
  if (cachedListId) return cachedListId
  const listName = env('SHAREPOINT_LIST_NAME')
  if (!listName) throw new Error('SHAREPOINT_LIST_NAME no configurada')
  const list = await graphGet<{ id: string }>(
    `/sites/${siteId}/lists/${encodeURIComponent(listName)}`,
  )
  cachedListId = list.id
  return list.id
}

/** ID de usuario en el sitio SharePoint (columna SolicitadoPorLookupId). */
async function resolveSiteUserLookupId(
  siteId: string,
  email: string,
): Promise<string | undefined> {
  try {
    const user = await graphPost<{ id?: string }>(`/sites/${siteId}/users`, {
      email: email.trim(),
    })
    return user.id?.trim() || undefined
  } catch {
    return undefined
  }
}

function buildSharePointFields(input: {
  title: string
  description: string
  category: string
  priority: string
  ansHours: number
  possibleSolutions: string[]
  userEmail?: string
  userName?: string
  jobTitle?: string
  department?: string
  officeLocation?: string
  phone?: string
  solicitadoPorLookupId?: string
}): Record<string, string | number> {
  const fields: Record<string, string | number> = {
    Title: input.title.slice(0, 255),
  }

  const descriptionField = fieldName('DESCRIPTION', 'Descripcion')
  const requesterTextField = fieldName('REQUESTER_TEXT', 'SolicitadoPorTexto')
  const requesterLookupField = fieldName('REQUESTER_LOOKUP', 'SolicitadoPorLookupId')
  const categoryTextField = fieldName('CATEGORY_TEXT', 'CategoriaTexto')
  const departmentTextField = fieldName('DEPARTMENT_TEXT', 'DepartamentoTexto')
  const towerTextField = fieldName('TOWER_TEXT', 'TorreTexto')
  const contactPhoneField = fieldName('CONTACT_PHONE', 'Numero_Contacto')
  const priorityLookupField = fieldName('PRIORITY_LOOKUP', 'PrioridadLookupId')
  const categoryLookupField = fieldName('CATEGORY_LOOKUP', 'CategoriaLookupId')
  const estadoLookupField = fieldName('STATUS_LOOKUP', 'EstadoLookupId')

  const sourceLabel = env('SHAREPOINT_SOURCE_VALUE') ?? 'Chatbot IA — Mesa Servicios'
  const descriptionParts = [
    input.description.trim(),
    '',
    `Origen: ${sourceLabel}`,
    `ANS referencial (horas): ${input.ansHours}`,
  ]
  if (input.possibleSolutions.length) {
    descriptionParts.push(
      '',
      'Posibles soluciones sugeridas por IA:',
      ...input.possibleSolutions.map((s) => `• ${s}`),
    )
  }
  fields[descriptionField] = descriptionParts.join('\n').slice(0, 8000)

  const priorityId = PRIORITY_LOOKUP_IDS[input.priority] ?? PRIORITY_LOOKUP_IDS.media
  const categoryId =
    CATEGORY_LOOKUP_IDS[input.category] ?? CATEGORY_LOOKUP_IDS.otro
  fields[priorityLookupField] = priorityId
  fields[categoryLookupField] = categoryId
  fields[categoryTextField] = input.category
  fields[estadoLookupField] =
    Number(env('SHAREPOINT_ESTADO_LOOKUP_ID')) || DEFAULT_ESTADO_LOOKUP_ID

  const requesterLabel =
    input.userName?.trim() ||
    input.userEmail?.trim() ||
    ''
  if (requesterLabel) fields[requesterTextField] = requesterLabel.slice(0, 255)
  if (input.solicitadoPorLookupId) {
    fields[requesterLookupField] = input.solicitadoPorLookupId
  }
  if (input.department) fields[departmentTextField] = input.department.slice(0, 255)
  if (input.jobTitle) fields[towerTextField] = input.jobTitle.slice(0, 255)
  if (input.officeLocation) {
    descriptionParts.push(`Ubicación (oficina): ${input.officeLocation}`)
    fields[descriptionField] = descriptionParts.join('\n').slice(0, 8000)
  }
  if (input.phone) {
    const digits = input.phone.replace(/\D/g, '')
    if (digits) fields[contactPhoneField] = Number(digits.slice(0, 15))
  }

  return fields
}

export async function createSharePointTicket(input: {
  title: string
  description: string
  category: string
  priority: 'baja' | 'media' | 'alta' | 'critica'
  ansHours: number
  possibleSolutions: string[]
  userEmail?: string
  userName?: string
  jobTitle?: string
  department?: string
  officeLocation?: string
  phone?: string
}): Promise<Ticket> {
  const siteId = await resolveSiteId()
  const listId = await resolveListId(siteId)
  let solicitadoPorLookupId: string | undefined
  if (input.userEmail) {
    solicitadoPorLookupId = await resolveSiteUserLookupId(siteId, input.userEmail)
  }
  const fields = buildSharePointFields({ ...input, solicitadoPorLookupId })

  const created = await graphPost<{
    id: string
    webUrl?: string
    createdDateTime?: string
    fields?: Record<string, unknown>
  }>(`/sites/${siteId}/lists/${listId}/items`, { fields })

  const itemId = created.id
  const createdAt = created.createdDateTime ?? new Date().toISOString()

  return {
    id: `SP-${itemId}`,
    title: input.title,
    description: input.description,
    category: input.category,
    priority: input.priority,
    status: 'abierto',
    ansHours: input.ansHours,
    possibleSolutions: input.possibleSolutions,
    createdAt,
    userEmail: input.userEmail,
  }
}
