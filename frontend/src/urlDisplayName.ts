/**
 * Nombre del usuario para el saludo cuando la app se abre embebida (p. ej. Power Apps)
 * vía URL. Se aceptan varias claves por si el equipo de Power Apps usa nombres distintos.
 *
 * En Power Apps (ejemplo): URL del control Web + "?nombre=" & EncodeUrl( User().FullName )
 */

const PARAM_KEYS = [
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

function firstNameFromParams(params: URLSearchParams): string {
  for (const key of PARAM_KEYS) {
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

export function readDisplayNameFromUrl(): string {
  const fromSearch = firstNameFromParams(new URLSearchParams(window.location.search))
  if (fromSearch) return fromSearch
  const fromHash = paramsFromHash()
  if (fromHash) {
    const fromHashName = firstNameFromParams(fromHash)
    if (fromHashName) return fromHashName
  }
  return ''
}
