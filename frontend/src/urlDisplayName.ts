/**
 * Nombre del usuario para el saludo cuando la app se abre embebida (p. ej. Power Apps)
 * vía URL. Se aceptan varias claves por si el equipo de Power Apps usa nombres distintos.
 *
 * En Power Apps (botón que abre el chatbot), por ejemplo:
 *   "https://tu-chatbot.com/?nombre=" & EncodeUrl(User().FullName) &
 *   "&email=" & EncodeUrl(User().Email)
 * Si además está activo el login Microsoft (MSAL), el correo de la URL acelera el SSO.
 */

const NAME_PARAM_KEYS = [
  'nombre',
  'name',
  'displayName',
  'displayname',
  'fullName',
  'fullname',
  'usuario',
  'user',
  'givenName',
  'givenname',
] as const

const EMAIL_PARAM_KEYS = [
  'email',
  'correo',
  'mail',
  'userEmail',
  'useremail',
  'upn',
] as const

function decodeParamValue(raw: string): string {
  let s = raw.replace(/\+/g, ' ').trim()
  for (let i = 0; i < 2; i += 1) {
    try {
      const next = decodeURIComponent(s)
      if (next === s) break
      s = next
    } catch {
      break
    }
  }
  return s.trim()
}

function firstMatchFromParams(params: URLSearchParams, keys: readonly string[]): string {
  for (const key of keys) {
    const raw = params.get(key) ?? params.get(key.toLowerCase())
    if (raw) {
      const decoded = decodeParamValue(raw)
      if (decoded) return decoded
    }
  }
  return ''
}

/** Query string opcional dentro del hash (#/ruta?nombre=...) */
function paramsFromHash(): URLSearchParams | null {
  const { hash } = window.location
  if (!hash?.includes('?')) return null
  const q = hash.slice(hash.indexOf('?') + 1)
  return new URLSearchParams(q)
}

function readParamSet(): URLSearchParams[] {
  const sets = [new URLSearchParams(window.location.search)]
  const fromHash = paramsFromHash()
  if (fromHash) sets.push(fromHash)
  return sets
}

export function readDisplayNameFromUrl(): string {
  for (const params of readParamSet()) {
    const name = firstMatchFromParams(params, NAME_PARAM_KEYS)
    if (name) return name
  }
  return ''
}

/** Correo pasado desde Power Apps (User().Email) u otra fuente embebida. */
export function readEmailFromUrl(): string {
  for (const params of readParamSet()) {
    const email = firstMatchFromParams(params, EMAIL_PARAM_KEYS)
    if (email && email.includes('@')) return email
  }
  return ''
}

export function readUserFromUrl(): { name: string; email: string } {
  return {
    name: readDisplayNameFromUrl(),
    email: readEmailFromUrl(),
  }
}
