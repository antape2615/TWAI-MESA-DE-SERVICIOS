/**
 * Plantilla texto + enlace al reproductor Power Apps (HelpDesk).
 *
 * Por defecto **no** se añaden parámetros de consulta desde el servidor (solo la URL base de
 * HELPDESK_POWERAPPS_URL). Para precargar Titulo, Descripcion, etc. en la URL, defina
 * HELPDESK_URL_QUERY_PARAMS_ENABLED=true (la app canvas debe leer Param(...) con los nombres
 * acordados; HELPDESK_URL_PARAM_* opcional si difieren).
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

/** URL base del reproductor HelpDesk (cabecera del chat). */
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

/** true = añadir Titulo, Descripcion, etc. al query string (precarga en Power Apps). */
export function helpdeskUrlQueryParamsEnabled(): boolean {
  return process.env.HELPDESK_URL_QUERY_PARAMS_ENABLED === 'true'
}

/** Construye URL de play: por defecto solo la base; con HELPDESK_URL_QUERY_PARAMS_ENABLED precarga campos. */
export function buildPowerAppsDeepLink(fields: HelpdeskLinkFields): string | null {
  const base = process.env.HELPDESK_POWERAPPS_URL?.trim()
  if (!base) return null

  let url: URL
  try {
    url = new URL(base)
  } catch {
    return null
  }

  if (!helpdeskUrlQueryParamsEnabled()) {
    return url.toString()
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

/**
 * Si el modelo no llamó a open_helpdesk_link, generamos igual la URL con título/descripción
 * desde el último mensaje del usuario y el correo opcional (Power Apps lee Param(...)).
 */
/** Campos inferidos del último mensaje y correo (misma lógica que el enlace fallback). */
export function buildFallbackHelpdeskFields(
  lastUserText: string,
  userEmail?: string,
  userName?: string,
): HelpdeskLinkFields {
  const raw = lastUserText.trim()
  const firstLine = raw.split(/[\n\r]+/)[0]?.trim() ?? ''
  const titulo =
    (firstLine || raw).slice(0, 120).trim() ||
    'Solicitud desde el asistente de Mesa de Servicios (Periferia)'
  const descripcion =
    raw.slice(0, MAX_DESCRIPCION) ||
    '(Sin texto en el chat; complete el caso en HelpDesk.)'
  return {
    titulo,
    descripcion,
    solicitado_por:
      userName?.trim() || userEmail?.trim() || undefined,
  }
}

export function buildFallbackDeepLink(
  lastUserText: string,
  userEmail?: string,
  userName?: string,
): string | null {
  if (!hasHelpdeskPowerAppsUrl()) return null
  return buildPowerAppsDeepLink(
    buildFallbackHelpdeskFields(lastUserText, userEmail, userName),
  )
}

/** Combina lo devuelto por la herramienta con el fallback del último mensaje. */
export function mergeHelpdeskFields(
  tool: HelpdeskLinkFields | undefined,
  fallback: HelpdeskLinkFields,
): HelpdeskLinkFields {
  const pick = (key: keyof HelpdeskLinkFields): string | undefined => {
    const tv = tool?.[key]?.trim()
    if (tv) return tv
    return fallback[key]?.trim()
  }
  return {
    titulo:
      pick('titulo') ??
      fallback.titulo ??
      'Solicitud desde el asistente de Mesa de Servicios (Periferia)',
    descripcion:
      pick('descripcion') ??
      fallback.descripcion ??
      '(Sin texto en el chat; complete el caso en HelpDesk.)',
    categoria: pick('categoria'),
    solicitado_por: pick('solicitado_por'),
    numero_contacto: pick('numero_contacto'),
    pais: pick('pais'),
    departamento: pick('departamento'),
    torre: pick('torre'),
  }
}

export type HelpdeskCopyField = {
  key: string
  label: string
  value: string
  /** Texto auxiliar bajo el valor (p. ej. nota para «Solicitado por»). */
  note?: string
}

const COPY_FIELD_ORDER: {
  key: keyof HelpdeskLinkFields
  label: string
  note?: string
}[] = [
  {
    key: 'titulo',
    label: 'Título o asunto del incidente (resumen para mesa de ayuda)',
  },
  {
    key: 'descripcion',
    label:
      'Descripción detallada: qué ocurre, mensajes de error y pasos que ya probó',
  },
  {
    key: 'categoria',
    label: 'Categoría del caso (p. ej. hardware, software, red, acceso, otro)',
  },
  {
    key: 'solicitado_por',
    label: 'Persona que solicita el soporte (correo corporativo o nombre completo)',
    note:
      'Si un compañero no puede usar la herramienta, indique quién queda como solicitante.',
  },
  {
    key: 'numero_contacto',
    label: 'Teléfono u otro número de contacto para localizarle con rapidez',
  },
  { key: 'pais', label: 'País o ubicación geográfica relevante para el soporte' },
  {
    key: 'departamento',
    label: 'Departamento, área o dirección dentro de la organización',
  },
  {
    key: 'torre',
    label: 'Torre, sede u oficina física (si aplica a su solicitud)',
  },
]

const PLACEHOLDER = '[por completar]'

/** Filas para la UI: valor vacío → marcador de plantilla. */
export function helpdeskFieldsToCopyRows(fields: HelpdeskLinkFields): HelpdeskCopyField[] {
  return COPY_FIELD_ORDER.map(({ key, label, note }) => {
    const raw = fields[key]?.trim()
    return {
      key,
      label,
      value: raw || PLACEHOLDER,
      ...(note ? { note } : {}),
    }
  })
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
export const HELPDESK_TICKET_TEMPLATE = `Título o asunto del incidente (resumen para mesa de ayuda): [por completar]

Descripción detallada (síntomas, errores en pantalla, pasos ya probados): [por completar]

Categoría del caso (hardware, software, red, acceso, otro): [por completar]

Persona que solicita el soporte (correo o nombre completo): [por completar]
(Nota: si un compañero no puede usar la herramienta, indique quién queda como solicitante.)

Teléfono u otro número de contacto: [por completar]

País o ubicación geográfica relevante: [por completar]

Departamento o área dentro de la organización: [por completar]

Torre, sede u oficina (si aplica): [por completar]`

export function helpdeskTemplateForPrompt(): string {
  return HELPDESK_TICKET_TEMPLATE
}
