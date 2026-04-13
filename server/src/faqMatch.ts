import fs from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import type { ChatCompletionMessageParam } from 'openai/resources/chat/completions'

export type FaqItem = {
  id: string
  title: string
  category: string
  content: string
}

let cache: FaqItem[] | null = null

async function loadFaqItems(): Promise<FaqItem[]> {
  if (cache) return cache
  const here = path.dirname(fileURLToPath(import.meta.url))
  const candidates = [
    path.join(process.cwd(), 'data', 'soporte-faq.json'),
    path.join(process.cwd(), '..', 'data', 'soporte-faq.json'),
    path.join(here, '..', '..', 'data', 'soporte-faq.json'),
  ]
  for (const p of candidates) {
    try {
      const raw = await fs.readFile(p, 'utf-8')
      const parsed = JSON.parse(raw) as FaqItem[]
      if (Array.isArray(parsed) && parsed.length > 0) {
        cache = parsed
        return cache
      }
    } catch {
      /* siguiente candidato */
    }
  }
  console.warn(
    '[faq] No se encontró data/soporte-faq.json; las FAQ no se inyectarán en el prompt.',
  )
  cache = []
  return cache
}

/** Palabras / frases extra para acercar la consulta del usuario a la FAQ correcta */
const ID_HINTS: Record<string, string[]> = {
  'maquina-password': [
    'maquina',
    'máquina',
    'computador',
    'computadora',
    'equipo',
    'windows',
    'sesion',
    'sesión',
    'inicio',
    'login',
    'clave',
    'contraseña',
    'password',
    'bloqueado',
    'ingresar',
  ],
  vpn: ['vpn', 'forticlient', 'forti', 'remoto', 'teletrabajo', 'tunnel'],
  'correo-login': ['correo', 'email', 'outlook', 'office', 'bloqueada', 'desbloqueo'],
  'correos-no-llegan': [
    'llegan',
    'recibir',
    'bandeja',
    'mensajes',
    'mail',
    'envian',
    'envían',
    'no llega',
  ],
  'pc-no-enciende': [
    'enciende',
    'prender',
    'encender',
    'luz',
    'cargador',
    'corriente',
    'apagado',
  ],
  'pc-sin-internet': [
    'internet',
    'wifi',
    'wi-fi',
    'red',
    'conexion',
    'conexión',
    'sin acceso',
    'navegar',
  ],
  'equipo-lento-windows': [
    'lento',
    'lentitud',
    'despacio',
    'actualizacion',
    'actualización',
    'windows update',
    'controladores',
    'drivers',
    'rendimiento',
  ],
  'equipo-lento-temp': [
    'temp',
    'temporal',
    'basura',
    'espacio',
    'disco',
    'limpiar',
    'desinstalar',
    '%temp%',
  ],
  'permisos-admin': [
    'admin',
    'administrador',
    'instalar',
    'permiso',
    'elevado',
    'lider',
    'líder',
    'torre',
    'aprobacion',
    'aprobación',
  ],
  'bloqueo-app': [
    'bloqueo',
    'bloqueada',
    'pegada',
    'congelada',
    'aplicacion',
    'aplicación',
    'programa',
    'tareas',
    'finalizar',
    'administrador de tareas',
  ],
}

function normalize(s: string): string {
  return s
    .toLowerCase()
    .normalize('NFD')
    .replace(/\p{M}/gu, '')
}

function tokenizeUser(s: string): string[] {
  const n = normalize(s).replace(/[^a-z0-9ñ\s]/g, ' ')
  return n.split(/\s+/).filter((w) => w.length >= 2)
}

function scoreItem(item: FaqItem, userTokens: Set<string>, rawUser: string): number {
  const nu = normalize(rawUser)
  let score = 0
  const haystack = normalize(`${item.title} ${item.category} ${item.content}`)

  for (const t of userTokens) {
    if (t.length < 2) continue
    if (haystack.includes(t)) score += 2
  }

  const titleN = normalize(item.title)
  for (const t of userTokens) {
    if (t.length >= 3 && titleN.includes(t)) score += 4
  }

  for (const h of ID_HINTS[item.id] ?? []) {
    const hn = normalize(h)
    if (hn.length >= 4 && nu.includes(hn)) score += 5
    else if (nu.includes(hn)) score += 3
    for (const t of userTokens) {
      if (t.length >= 3 && hn.includes(t)) score += 2
    }
  }

  if (nu.includes('vpn')) {
    if (item.id === 'vpn') score += 18
    else score = Math.floor(score * 0.2)
  }
  if (
    nu.includes('computador') ||
    nu.includes('computadora') ||
    nu.includes('maquina') ||
    nu.includes('máquina')
  ) {
    if (item.id === 'maquina-password') score += 12
  }

  const powerIssue =
    nu.includes('enciende') ||
    nu.includes('prender') ||
    nu.includes('encender') ||
    nu.includes('cargador') ||
    (nu.includes('luz') && (nu.includes('no') || nu.includes('led')))
  if (powerIssue) {
    if (item.id === 'pc-no-enciende') score += 16
    if (item.id === 'maquina-password') score = Math.floor(score * 0.3)
  }

  return score
}

/**
 * Selecciona FAQs cuyo texto coincide con la consulta del usuario (palabras clave + pistas).
 */
export async function selectRelevantFaqs(
  userText: string,
  max = 3,
): Promise<FaqItem[]> {
  const items = await loadFaqItems()
  if (!items.length || !userText.trim()) return []

  const tokens = tokenizeUser(userText)
  if (tokens.length === 0) return []
  const ut = new Set(tokens)

  const scored = items
    .map((item) => ({ item, score: scoreItem(item, ut, userText) }))
    .filter((x) => x.score > 0)
    .sort((a, b) => b.score - a.score)

  if (!scored.length) return []

  const best = scored[0].score
  if (best < 3) return []

  const minInclude = Math.max(3, Math.floor(best * 0.42))
  const out: FaqItem[] = []
  for (const { item, score } of scored) {
    if (score < minInclude) break
    out.push(item)
    if (out.length >= max) break
  }
  return out
}

export function collectUserText(messages: ChatCompletionMessageParam[]): string {
  const parts: string[] = []
  for (const m of messages) {
    if (m.role !== 'user') continue
    const c = m.content
    if (typeof c === 'string') parts.push(c)
    else if (Array.isArray(c)) {
      for (const part of c) {
        if (part.type === 'text' && 'text' in part) parts.push(part.text)
      }
    }
  }
  return parts.join('\n')
}

export function faqsToPromptBlock(items: FaqItem[]): string {
  if (!items.length) return ''
  const blocks = items.map(
    (i) => `### ${i.title} (${i.category})\n${i.content}`,
  )
  return [
    '--- FAQ oficiales (Consulta Soporte / Excel) ---',
    'Si la consulta del usuario está relacionada con alguno de estos temas, respóndele de forma conversacional y clara en español.',
    'Basa los pasos y recomendaciones en el texto siguiente; puedes resumir o usar listas; no contradigas esta guía.',
    '',
    ...blocks,
  ].join('\n')
}
