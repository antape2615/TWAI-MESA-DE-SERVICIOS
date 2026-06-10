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

export function hasSharePointRequesterIdentity(input: {
  userEmail?: string
  userName?: string
}): boolean {
  return Boolean(input.userEmail?.trim() || input.userName?.trim())
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

/** Origen del sitio SharePoint (p. ej. https://periferiaitgroup.sharepoint.com). */
export function getSharePointResourceOrigin(): string | null {
  const site = env('SHAREPOINT_SITE_URL')
  if (!site?.startsWith('http')) return null
  try {
    return new URL(site).origin
  } catch {
    return null
  }
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

type LookupCache = { at: number; byName: Map<string, number> }
let lookupCache: LookupCache | null = null
const LOOKUP_CACHE_TTL_MS = 10 * 60 * 1000

function normalizePersonName(value: string): string {
  return value
    .toLowerCase()
    .normalize('NFD')
    .replace(/\p{M}/gu, '')
    .replace(/[^a-z0-9\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
}

function personNameTokens(value: string): string[] {
  return normalizePersonName(value)
    .split(' ')
    .filter((t) => t.length > 1)
}

/** ID de usuario en el sitio SharePoint (columna SolicitadoPorLookupId). App-only. */
async function resolveSiteUserLookupIdApp(
  siteId: string,
  email: string,
): Promise<number | undefined> {
  try {
    const user = await graphPost<{ id?: string | number }>(`/sites/${siteId}/users`, {
      email: email.trim(),
    })
    const id = user.id
    if (id == null || id === '') return undefined
    return Number(id)
  } catch {
    return undefined
  }
}

function parseEnsureUserId(payload: unknown): number | undefined {
  if (!payload || typeof payload !== 'object') return undefined
  const data = payload as {
    Id?: number | string
    d?: { Id?: number | string; EnsureUser?: { Id?: number | string } }
  }
  const raw = data.Id ?? data.d?.Id ?? data.d?.EnsureUser?.Id
  if (raw == null || raw === '') return undefined
  const id = Number(raw)
  return Number.isNaN(id) ? undefined : id
}

function ensureUserLogonNames(email: string): string[] {
  const trimmed = email.trim()
  const lower = trimmed.toLowerCase()
  const variants = trimmed === lower ? [trimmed] : [trimmed, lower]
  const out = new Set<string>()
  for (const e of variants) {
    out.add(`i:0#.f|membership|${e}`)
    out.add(e)
  }
  return [...out]
}

/** Con token delegado del usuario (requiere scope Sites.ReadWrite.All en MSAL). */
async function resolveSiteUserLookupIdDelegated(
  siteId: string,
  email: string,
  accessToken: string,
): Promise<number | undefined> {
  try {
    const res = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/users`, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ email: email.trim() }),
    })
    if (!res.ok) {
      const err = await res.text().catch(() => '')
      console.warn(`[sharepoint] Graph /sites/users delegado ${res.status}: ${err.slice(0, 200)}`)
      return undefined
    }
    const user = (await res.json()) as { id?: string | number }
    if (user.id == null || user.id === '') return undefined
    return Number(user.id)
  } catch (e) {
    console.warn('[sharepoint] Graph /sites/users delegado:', e)
    return undefined
  }
}

/** SharePoint REST ensureuser — requiere token con audiencia SharePoint (no solo Graph). */
async function resolveViaSharePointEnsureUser(
  email: string,
  accessToken: string,
): Promise<number | undefined> {
  const siteUrl = env('SHAREPOINT_SITE_URL')
  if (!siteUrl) return undefined
  const endpoint = `${siteUrl.replace(/\/$/, '')}/_api/web/ensureuser`

  for (const logonName of ensureUserLogonNames(email)) {
    for (const accept of [
      'application/json;odata=nometadata',
      'application/json;odata=verbose',
    ] as const) {
      try {
        const res = await fetch(endpoint, {
          method: 'POST',
          headers: {
            Authorization: `Bearer ${accessToken}`,
            Accept: accept,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({ logonName }),
        })
        if (!res.ok) {
          const err = await res.text().catch(() => '')
          console.warn(
            `[sharepoint] ensureuser ${res.status} (${logonName.slice(0, 40)}…): ${err.slice(0, 160)}`,
          )
          continue
        }
        const data = await res.json().catch(() => null)
        const id = parseEnsureUserId(data)
        if (id != null) return id
      } catch (e) {
        console.warn('[sharepoint] ensureuser:', e)
      }
    }
  }

  return undefined
}

async function buildRequesterLookupCache(
  siteId: string,
  listId: string,
): Promise<Map<string, number>> {
  const now = Date.now()
  if (lookupCache && now - lookupCache.at < LOOKUP_CACHE_TTL_MS) {
    return lookupCache.byName
  }

  const byName = new Map<string, number>()
  let url: string | null =
    `/sites/${siteId}/lists/${listId}/items?` +
    new URLSearchParams({
      $expand: 'fields($select=SolicitadoPorLookupId,SolicitadoPorTexto)',
      $top: '500',
      $orderby: 'createdDateTime desc',
    })

  type ListItemsPage = {
    value?: Array<{ fields?: Record<string, unknown> }>
    '@odata.nextLink'?: string
  }

  for (let page = 0; page < 6 && url; page++) {
    const path = url.startsWith('http')
      ? url.replace('https://graph.microsoft.com/v1.0', '')
      : url
    const data: ListItemsPage = await graphGet<ListItemsPage>(path)

    for (const item of data.value ?? []) {
      const text = String(item.fields?.SolicitadoPorTexto ?? '').trim()
      const lookupRaw = item.fields?.SolicitadoPorLookupId
      if (!text || lookupRaw == null || lookupRaw === '') continue
      const lookupId = Number(lookupRaw)
      if (Number.isNaN(lookupId)) continue
      const key = normalizePersonName(text)
      if (!byName.has(key)) byName.set(key, lookupId)
    }

    url = data['@odata.nextLink']
      ? data['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0', '')
      : null
  }

  lookupCache = { at: now, byName }
  return byName
}

function matchLookupIdFromHistory(
  userName: string,
  cache: Map<string, number>,
): number | undefined {
  const norm = normalizePersonName(userName)
  const exact = cache.get(norm)
  if (exact != null) return exact

  const tokens = personNameTokens(userName)
  if (tokens.length < 2) return undefined

  const first = tokens[0]
  const last = tokens[tokens.length - 1]
  let best: { id: number; score: number } | undefined

  for (const [key, id] of cache) {
    const keyTokens = personNameTokens(key)
    if (keyTokens[0] !== first) continue
    let score = 1
    if (keyTokens.includes(last)) score = 10
    else if (key.includes(last)) score = 5
    if (!best || score > best.score) best = { id, score }
  }

  return best && best.score >= 5 ? best.id : undefined
}

async function resolveSolicitadoPorLookupId(
  siteId: string,
  listId: string,
  input: {
    email?: string
    name?: string
    accessToken?: string
    sharePointAccessToken?: string
  },
): Promise<number | undefined> {
  const email = input.email?.trim()
  const name = input.name?.trim()
  const accessToken = input.accessToken?.trim()
  const sharePointAccessToken = input.sharePointAccessToken?.trim()

  if (email && sharePointAccessToken) {
    const ensured = await resolveViaSharePointEnsureUser(email, sharePointAccessToken)
    if (ensured != null) return ensured
  }

  if (email && accessToken) {
    const delegated = await resolveSiteUserLookupIdDelegated(siteId, email, accessToken)
    if (delegated != null) return delegated
    const ensuredGraph = await resolveViaSharePointEnsureUser(email, accessToken)
    if (ensuredGraph != null) return ensuredGraph
  }

  if (email) {
    const appOnly = await resolveSiteUserLookupIdApp(siteId, email)
    if (appOnly != null) return appOnly
  }

  if (name) {
    const cache = await buildRequesterLookupCache(siteId, listId)
    const fromHistory = matchLookupIdFromHistory(name, cache)
    if (fromHistory != null) return fromHistory
  }

  return undefined
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
  solicitadoPorLookupId?: number
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
  if (input.solicitadoPorLookupId != null) {
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
  accessToken?: string
  sharePointAccessToken?: string
}): Promise<Ticket> {
  if (!hasSharePointRequesterIdentity(input)) {
    throw new Error(
      'Debe iniciar sesión con Microsoft para crear el ticket en SharePoint (campo Solicitado Por).',
    )
  }
  const siteId = await resolveSiteId()
  const listId = await resolveListId(siteId)
  const solicitadoPorLookupId = await resolveSolicitadoPorLookupId(siteId, listId, {
    email: input.userEmail,
    name: input.userName,
    accessToken: input.accessToken,
    sharePointAccessToken: input.sharePointAccessToken,
  })
  if (solicitadoPorLookupId == null) {
    console.warn(
      `[sharepoint] SolicitadoPor sin lookup (solo texto). email=${input.userEmail ?? '—'} name=${input.userName ?? '—'} tokenGraph=${input.accessToken ? 'sí' : 'no'} tokenSP=${input.sharePointAccessToken ? 'sí' : 'no'}`,
    )
  } else {
    console.log(
      `[sharepoint] SolicitadoPorLookupId=${solicitadoPorLookupId} (${input.userName ?? input.userEmail ?? 'usuario'})`,
    )
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
