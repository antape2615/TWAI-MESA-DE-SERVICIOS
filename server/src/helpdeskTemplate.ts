/**
 * Plantilla texto + enlace profundo a Power Apps.
 *
 * La app canvas HelpDesk debe leer los mismos nombres en Param(), por ejemplo:
 * Param("Titulo"), Param("Descripcion"), Param("Categoria"), Param("SolicitadoPor"),
 * Param("NumeroContacto"), Param("Pais"), Param("Departamento"), Param("Torre")
 *
 * Si los nombres en Power Apps difieren, ajuste HELPDESK_URL_PARAM_* en .env (opcional).
 */

/** Nombres por defecto en el query string (ASCII, estables en URLs). */
const DEFAULT_QUERY_NAMES: Record<string, string> = {
  titulo: 'Titulo',
  descripcion: 'Descripcion',
  categoria: 'Categoria',
  solicitado_por: 'SolicitadoPor',
  numero_contacto: 'NumeroContacto',
  pais: 'Pais',
  departamento: 'Departamento',
  torre: 'Torre',
}

export type HelpdeskLinkFields = {
  titulo?: string
  descripcion?: string
  categoria?: string
  solicitado_por?: string
  numero_contacto?: string
  pais?: string
  departamento?: string
  torre?: string
}

const MAX_DESCRIPCION = 1800

export function hasHelpdeskPowerAppsUrl(): boolean {
  const u = process.env.HELPDESK_POWERAPPS_URL?.trim()
  return Boolean(u && u.startsWith('http'))
}

/** URL base del reproductor HelpDesk (para botón «Ir a HelpDesk» en el chat). */
export function getHelpdeskPowerAppsUrl(): string | null {
  const u = process.env.HELPDESK_POWERAPPS_URL?.trim()
  if (!u || !u.startsWith('http')) return null
  return u
}

function paramName(internal: keyof HelpdeskLinkFields): string {
  const envKey = `HELPDESK_URL_PARAM_${String(internal).toUpperCase()}`
  const fromEnv = process.env[envKey]?.trim()
  if (fromEnv) return fromEnv
  return DEFAULT_QUERY_NAMES[internal] ?? String(internal)
}

/** Construye URL de play de Power Apps con query params para precarga. */
export function buildPowerAppsDeepLink(fields: HelpdeskLinkFields): string | null {
  const base = process.env.HELPDESK_POWERAPPS_URL?.trim()
  if (!base) return null

  let url: URL
  try {
    url = new URL(base)
  } catch {
    return null
  }

  url.searchParams.set('sourcetime', String(Date.now()))

  const entries: [keyof HelpdeskLinkFields, string | undefined][] = [
    ['titulo', fields.titulo],
    ['descripcion', fields.descripcion],
    ['categoria', fields.categoria],
    ['solicitado_por', fields.solicitado_por],
    ['numero_contacto', fields.numero_contacto],
    ['pais', fields.pais],
    ['departamento', fields.departamento],
    ['torre', fields.torre],
  ]

  for (const [key, raw] of entries) {
    if (!raw?.trim()) continue
    let v = raw.trim()
    if (key === 'descripcion' && v.length > MAX_DESCRIPCION) {
      v = `${v.slice(0, MAX_DESCRIPCION)}…`
    }
    url.searchParams.set(paramName(key), v)
  }

  return url.toString()
}

/** Normaliza argumentos JSON de la herramienta (snake_case). */
export function parseHelpdeskLinkArgs(
  args: Record<string, unknown>,
): HelpdeskLinkFields {
  const s = (k: string) => {
    const v = args[k]
    return v == null ? undefined : String(v).trim() || undefined
  }
  return {
    titulo: s('titulo'),
    descripcion: s('descripcion'),
    categoria: s('categoria'),
    solicitado_por: s('solicitado_por'),
    numero_contacto: s('numero_contacto'),
    pais: s('pais'),
    departamento: s('departamento'),
    torre: s('torre'),
  }
}

/** Campos alineados al formulario HelpDesk (Nuevo Ticket). */
export const HELPDESK_TICKET_TEMPLATE = `Título: [por completar]

Descripción: [por completar]

Categoría: [por completar]

Solicitado por: [correo o nombre]
(Nota: si un compañero tiene problemas para usar la herramienta, asigne la solicitud a su nombre.)

Número de contacto: [por completar]

País: [por completar]

Departamento: [por completar]

Torre: [por completar]`

export function helpdeskTemplateForPrompt(): string {
  return HELPDESK_TICKET_TEMPLATE
}
