import { useCallback, useEffect, useMemo, useState, type ClipboardEvent, type ReactNode } from 'react'
import './App.css'
import {
  postChat,
  fileToImagePayload,
  createTicketConfirm,
  fetchKnowledge,
  saveKnowledge,
  fetchAppConfig,
  type AppConfig,
  type ChatMessage,
  type TicketDraft,
  type KnowledgePayload,
  type KnowledgeEntry,
} from './api'
import {
  microsoftSignInHint,
  readInitialUserSession,
  refreshTicketAuthTokens,
  resolveUserSession,
  shouldOfferMicrosoftSignIn,
  type UserSession,
} from './userContext'
import { readDisplayNameFromUrl } from './urlDisplayName'

/** Pestaña "Parámetros": oculta hasta definir acceso por usuario/rol */
const SHOW_KNOWLEDGE_TAB = false

type Tab = 'chat' | 'knowledge'

function priorityClass(p: string): string {
  if (p === 'critica') return 'badge badge--critica'
  if (p === 'alta') return 'badge badge--alta'
  if (p === 'media') return 'badge badge--media'
  if (p === 'baja') return 'badge badge--baja'
  return 'badge'
}

function shortenUrl(rawUrl: string): string {
  try {
    const parsed = new URL(rawUrl)
    const compactPath = parsed.pathname.length > 26 ? `${parsed.pathname.slice(0, 26)}...` : parsed.pathname
    return `${parsed.hostname}${compactPath}`
  } catch {
    return rawUrl.length > 40 ? `${rawUrl.slice(0, 40)}...` : rawUrl
  }
}

function renderAutoLinks(text: string, keyPrefix: string): ReactNode[] {
  const urlRegex = /https?:\/\/[^\s)]+/g
  const nodes: ReactNode[] = []
  let cursor = 0
  let match: RegExpExecArray | null
  while ((match = urlRegex.exec(text)) !== null) {
    if (match.index > cursor) nodes.push(text.slice(cursor, match.index))
    const url = match[0]
    nodes.push(
      <a
        key={`${keyPrefix}-url-${match.index}`}
        href={url}
        target="_blank"
        rel="noreferrer"
        className="bubble-link"
      >
        {shortenUrl(url)}
      </a>,
    )
    cursor = match.index + url.length
  }
  if (cursor < text.length) nodes.push(text.slice(cursor))
  return nodes
}

function renderMessageContent(content: string): ReactNode[] {
  const markdownLinkRegex = /\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)/g
  const nodes: ReactNode[] = []
  let cursor = 0
  let match: RegExpExecArray | null
  while ((match = markdownLinkRegex.exec(content)) !== null) {
    if (match.index > cursor) {
      nodes.push(...renderAutoLinks(content.slice(cursor, match.index), `txt-${match.index}`))
    }
    const [, label, url] = match
    nodes.push(
      <a
        key={`md-${match.index}`}
        href={url}
        target="_blank"
        rel="noreferrer"
        className="bubble-link"
      >
        {label}
      </a>,
    )
    cursor = match.index + match[0].length
  }
  if (cursor < content.length) {
    nodes.push(...renderAutoLinks(content.slice(cursor), `tail-${cursor}`))
  }
  return nodes
}

function MessageText({ text }: { text: string }) {
  return <p className="bubble-text">{renderMessageContent(text)}</p>
}

function ChatAvatarUser() {
  return (
    <span className="chat-avatar chat-avatar--user" aria-hidden>
      <svg viewBox="0 0 32 32" width="22" height="22" fill="none" xmlns="http://www.w3.org/2000/svg">
        <circle cx="16" cy="11" r="5" fill="rgba(255,255,255,0.95)" />
        <path
          d="M8 26c0-5 3.5-8 8-8s8 3 8 8"
          stroke="rgba(255,255,255,0.95)"
          strokeWidth="2.2"
          strokeLinecap="round"
        />
      </svg>
    </span>
  )
}

function ChatAvatarAssistant() {
  return (
    <span className="chat-avatar chat-avatar--assistant" aria-hidden>
      <img className="chat-avatar__img" src="/robot.png" alt="" />
    </span>
  )
}

function SendPlaneIcon() {
  return (
    <svg width="22" height="22" viewBox="0 0 24 24" fill="currentColor" aria-hidden xmlns="http://www.w3.org/2000/svg">
      <path d="M3.4 19.6 20.85 12 3.4 4.4l-.05 6.55L16 12 3.35 13.05z" />
    </svg>
  )
}

function MicrosoftIcon() {
  return (
    <svg width="18" height="18" viewBox="0 0 21 21" aria-hidden xmlns="http://www.w3.org/2000/svg">
      <rect x="1" y="1" width="9" height="9" fill="#f25022" />
      <rect x="11" y="1" width="9" height="9" fill="#7fba00" />
      <rect x="1" y="11" width="9" height="9" fill="#00a4ef" />
      <rect x="11" y="11" width="9" height="9" fill="#ffb900" />
    </svg>
  )
}

function ChatSection({
  appConfig,
  ticketsFromPortal,
  helpdeskDeepLink,
  sharePointListUrl,
  userSession,
  showMicrosoftSignIn,
  authHint,
  authLoading,
  onSignIn,
}: {
  appConfig: AppConfig | null
  ticketsFromPortal: boolean
  helpdeskDeepLink: boolean
  sharePointListUrl?: string
  userSession: UserSession
  showMicrosoftSignIn: boolean
  authHint: string | null
  authLoading: boolean
  onSignIn: () => void
}) {
  const [messages, setMessages] = useState<ChatMessage[]>([])
  const [input, setInput] = useState('')
  const [pendingFile, setPendingFile] = useState<File | null>(null)
  const [previewUrl, setPreviewUrl] = useState<string | null>(null)
  const [loading, setLoading] = useState(false)
  const [confirmingIndex, setConfirmingIndex] = useState<number | null>(null)
  const [error, setError] = useState<string | null>(null)

  const sessionPayload = useMemo(
    () => ({
      userEmail: userSession.email || undefined,
      userName: userSession.name || undefined,
      accessToken: userSession.accessToken,
      sharePointAccessToken: userSession.sharePointAccessToken,
      jobTitle: userSession.jobTitle,
      department: userSession.department,
      officeLocation: userSession.officeLocation,
      phone: userSession.phone,
    }),
    [userSession],
  )
  const mustSignInForTickets = Boolean(
    appConfig?.azureAuth?.enabled &&
      appConfig?.sharePointTickets &&
      !userSession.email?.trim() &&
      !authLoading,
  )

  const visitorName = useMemo(() => {
    const fromSession = userSession.name?.trim()
    if (fromSession) return fromSession
    const fromUrl = readDisplayNameFromUrl().trim()
    if (fromUrl) return fromUrl
    const email = userSession.email?.trim()
    if (email?.includes('@')) return email.split('@')[0] ?? ''
    return ''
  }, [userSession.name, userSession.email])

  const onPickImage = useCallback((file: File | null) => {
    setPreviewUrl((prev) => {
      if (prev) URL.revokeObjectURL(prev)
      return file ? URL.createObjectURL(file) : null
    })
    setPendingFile(file)
  }, [])

  const onPasteImage = useCallback((e: ClipboardEvent<HTMLTextAreaElement>) => {
    const cd = e.clipboardData
    if (!cd) return
    const items = cd.items
    if (items?.length) {
      for (let i = 0; i < items.length; i++) {
        const item = items[i]
        if (item.kind !== 'file' || !item.type.startsWith('image/')) continue
        const file = item.getAsFile()
        if (!file) continue
        e.preventDefault()
        onPickImage(file)
        setError(null)
        return
      }
    }
    const files = cd.files
    if (files?.length) {
      for (let i = 0; i < files.length; i++) {
        const file = files[i]
        if (!file.type.startsWith('image/')) continue
        e.preventDefault()
        onPickImage(file)
        setError(null)
        return
      }
    }
  }, [onPickImage])

  const buildApiSession = useCallback(async () => {
    if (!appConfig?.sharePointTickets || !userSession.email) return sessionPayload
    const refreshed = await refreshTicketAuthTokens(appConfig, userSession)
    return {
      ...sessionPayload,
      accessToken: refreshed.accessToken ?? sessionPayload.accessToken,
      sharePointAccessToken:
        refreshed.sharePointAccessToken ?? sessionPayload.sharePointAccessToken,
    }
  }, [appConfig, sessionPayload, userSession])

  const send = useCallback(async () => {
    const trimmed = input.trim()
    if ((!trimmed && !pendingFile) || loading) return
    if (mustSignInForTickets) {
      setError('Inicie sesión con Microsoft para usar el chat y crear tickets en SharePoint.')
      return
    }
    setError(null)
    let image: ChatMessage['image']
    try {
      if (pendingFile) image = await fileToImagePayload(pendingFile)
    } catch (e) {
      setError(e instanceof Error ? e.message : 'Error con la imagen')
      return
    }
    const text =
      trimmed || (image ? '(Captura del error adjunta — analizar imagen)' : '')
    const next: ChatMessage[] = [...messages, { role: 'user', content: text, image }]
    setMessages(next)
    setInput('')
    if (previewUrl) URL.revokeObjectURL(previewUrl)
    setPreviewUrl(null)
    setPendingFile(null)
    setLoading(true)
    try {
      const res = await postChat(next, await buildApiSession())
      setMessages([
        ...next,
        {
          role: 'assistant',
          content: res.message,
          ...(res.ticketDraft ? { ticketDraft: res.ticketDraft } : {}),
          ...(res.ticketId ? { ticketCreatedId: res.ticketId } : {}),
          ...(res.helpdeskUrl ? { helpdeskUrl: res.helpdeskUrl } : {}),
        },
      ])
    } catch (e) {
      setError(e instanceof Error ? e.message : 'Error al enviar')
    } finally {
      setLoading(false)
    }
  }, [
    buildApiSession,
    input,
    loading,
    messages,
    mustSignInForTickets,
    pendingFile,
    previewUrl,
  ])

  const onConfirmTicket = useCallback(
    async (messageIndex: number, draft: TicketDraft) => {
      if (confirmingIndex !== null) return
      setConfirmingIndex(messageIndex)
      setError(null)
      try {
        if (mustSignInForTickets) {
          setError('Inicie sesión con Microsoft para crear el ticket (Solicitado Por).')
          return
        }
        const payload = await buildApiSession()
        if (
          appConfig?.sharePointTickets &&
          !payload.userEmail &&
          !payload.userName
        ) {
          setError('Inicie sesión con Microsoft para crear el ticket (Solicitado Por).')
          return
        }
        const { ticket } = await createTicketConfirm(draft, payload)
        setMessages((prev) =>
          prev.map((msg, i) =>
            i === messageIndex
              ? {
                  ...msg,
                  ticketDraft: undefined,
                  ticketCreatedId: ticket.id,
                }
              : msg,
          ),
        )
      } catch (e) {
        setError(e instanceof Error ? e.message : 'No se pudo crear el ticket')
      } finally {
        setConfirmingIndex(null)
      }
    },
    [appConfig, buildApiSession, confirmingIndex, mustSignInForTickets],
  )

  const onDismissTicket = useCallback((messageIndex: number) => {
    setMessages((prev) =>
      prev.map((msg, i) =>
        i === messageIndex ? { ...msg, ticketDraft: undefined } : msg,
      ),
    )
  }, [])

  const greeting =
    visitorName.length > 0
      ? `Hola ${visitorName}, Soy Perxia y estoy aquí para solucionar tus preguntas`
      : 'Hola, Soy Perxia y estoy aquí para solucionar tus preguntas'

  return (
    <div className="chatbot-layout">
      <section className="chatbot-card" aria-labelledby="chat-heading">
        <h2 id="chat-heading" className="sr-only">
          Asistente HelpDesk Periferia IT
        </h2>
        <header className="chatbot-card__header">
          <div className="chatbot-card__header-top">
            <div className="chatbot-logo" aria-hidden>
              <div className="chatbot-logo__mark">
                <img className="chatbot-logo__img" src="/logohelpdesk.png" alt="" />
              </div>
              <div className="chatbot-logo__text">
                <span className="chatbot-logo__title">HelpDesk</span>
                <span className="chatbot-logo__subtitle">Periferia IT</span>
              </div>
            </div>
            <p className="chatbot-card__greeting">{greeting}</p>
          </div>
          {userSession.jobTitle || userSession.department ? (
            <p className="chatbot-card__role">
              {[userSession.jobTitle, userSession.department].filter(Boolean).join(' · ')}
            </p>
          ) : null}
          {userSession.graphProfile || userSession.accountProfile ? (
            <details className="chatbot-profile-details">
              <summary>Datos de sesión Microsoft</summary>
              <dl className="chatbot-profile-dl">
                <dt>Puesto / título</dt>
                <dd>{userSession.graphProfile?.jobTitle?.trim() || '—'}</dd>
                <dt>Departamento</dt>
                <dd>{userSession.graphProfile?.department?.trim() || '—'}</dd>
                <dt>Ubicación</dt>
                <dd>{userSession.graphProfile?.officeLocation?.trim() || '—'}</dd>
                <dt>Correo</dt>
                <dd>
                  {userSession.graphProfile?.mail?.trim() ||
                    userSession.graphProfile?.userPrincipalName?.trim() ||
                    userSession.email ||
                    '—'}
                </dd>
                <dt>Nombre</dt>
                <dd>{userSession.accountProfile?.displayName || userSession.name || '—'}</dd>
              </dl>
            </details>
          ) : null}
          {showMicrosoftSignIn && !authLoading ? (
            <button type="button" className="ms-signin-btn" onClick={onSignIn}>
              <MicrosoftIcon />
              Iniciar sesión con Microsoft
            </button>
          ) : null}
          {authHint ? <p className="chatbot-card__auth-hint">{authHint}</p> : null}
          {authLoading ? (
            <p className="chatbot-card__auth-hint">Verificando sesión…</p>
          ) : null}
        </header>
        <p className="chatbot-card__hint">
          {mustSignInForTickets ? (
            <>
              <strong>Inicie sesión con Microsoft</strong> para describir su problema y crear
              tickets con su usuario en «Solicitado Por».
            </>
          ) : (
            <>
              Describa el error o adjunte una captura (pegar con Ctrl+V / Cmd+V).{' '}
              {ticketsFromPortal
                ? 'Si aparece un borrador de ticket, confírmelo con los botones.'
                : helpdeskDeepLink
                  ? 'Para HelpDesk use el enlace que indique el asistente.'
                  : 'El registro formal del caso es en HelpDesk cuando el asistente lo indique.'}
            </>
          )}
        </p>
        {error ? <div className="error-banner chatbot-banner">{error}</div> : null}
        <div className="chat-messages">
          {messages.length === 0 ? (
            <p className="empty-state">Escriba su consulta para comenzar.</p>
          ) : (
            messages.map((m, i) =>
              m.role === 'user' ? (
                <div key={`${i}-${m.role}`} className="chat-row chat-row--user">
                  <div className="chat-row__inner">
                    <div className="bubble bubble--user">
                      <span className="bubble-label">Usted</span>
                      {m.image ? (
                        <img
                          className="bubble-image"
                          src={`data:${m.image.mimeType};base64,${m.image.dataBase64}`}
                          alt="Captura del error"
                        />
                      ) : null}
                      <MessageText text={m.content} />
                    </div>
                    <ChatAvatarUser />
                  </div>
                </div>
              ) : (
                <div key={`${i}-${m.role}`} className="chat-row chat-row--assistant">
                  <div className="chat-row__inner">
                    <ChatAvatarAssistant />
                    <div className="bubble bubble--assistant">
                      <span className="bubble-label">Asistente</span>
                      {m.image ? (
                        <img
                          className="bubble-image"
                          src={`data:${m.image.mimeType};base64,${m.image.dataBase64}`}
                          alt="Captura del error"
                        />
                      ) : null}
                      <MessageText text={m.content} />
                      {ticketsFromPortal && m.ticketCreatedId ? (
                        <p className="ticket-created-note">
                          Ticket <strong>{m.ticketCreatedId}</strong> creado
                          {sharePointListUrl ? (
                            <>
                              {' '}
                              en{' '}
                              <a href={sharePointListUrl} target="_blank" rel="noreferrer">
                                SharePoint
                              </a>
                            </>
                          ) : null}
                          .
                        </p>
                      ) : null}
                      {ticketsFromPortal && m.ticketDraft && !m.ticketCreatedId ? (
                        <div className="ticket-offer" role="group" aria-label="Confirmar ticket">
                          <p className="ticket-offer-title">¿Generar ticket con estos datos?</p>
                          <ul className="ticket-offer-meta">
                            <li>
                              <span className="ticket-offer-k">Título</span> {m.ticketDraft.title}
                            </li>
                            <li>
                              <span className="ticket-offer-k">Prioridad</span>{' '}
                              <span className={priorityClass(m.ticketDraft.priority)}>
                                {m.ticketDraft.priority}
                              </span>
                            </li>
                          </ul>
                          <div className="ticket-offer-actions">
                            <button
                              type="button"
                              className="btn-primary"
                              disabled={confirmingIndex !== null || loading}
                              onClick={() => void onConfirmTicket(i, m.ticketDraft!)}
                            >
                              {confirmingIndex === i ? 'Creando…' : 'Sí, generar ticket'}
                            </button>
                            <button
                              type="button"
                              className="btn-secondary"
                              disabled={confirmingIndex !== null}
                              onClick={() => onDismissTicket(i)}
                            >
                              No, gracias
                            </button>
                          </div>
                        </div>
                      ) : null}
                    </div>
                  </div>
                </div>
              ),
            )
          )}
          {loading ? (
            <div className="chat-row chat-row--assistant">
              <div className="chat-row__inner">
                <ChatAvatarAssistant />
                <div className="bubble bubble--assistant bubble--typing">
                  <span className="bubble-label">Asistente</span>
                  <p className="bubble-text bubble-text--plain">Pensando…</p>
                </div>
              </div>
            </div>
          ) : null}
        </div>
        <div className="chat-attach-row chat-attach-row--compact">
          <label className="file-attach">
            <input
              type="file"
              accept="image/png,image/jpeg,image/jpg,image/gif,image/webp"
              className="sr-only"
              onChange={(e) => onPickImage(e.target.files?.[0] ?? null)}
            />
            <span className="file-attach-label">Adjuntar captura</span>
          </label>
          {previewUrl ? (
            <div className="preview-wrap">
              <img src={previewUrl} alt="Vista previa" className="preview-thumb" />
              <button type="button" className="btn-secondary" onClick={() => onPickImage(null)}>
                Quitar
              </button>
            </div>
          ) : null}
        </div>
        <div className="chatbot-composer">
          <textarea
            className="chatbot-input"
            value={input}
            onChange={(e) => setInput(e.target.value)}
            placeholder={
              mustSignInForTickets
                ? 'Inicie sesión con Microsoft para continuar…'
                : 'Escribe tu mensaje… (Enter para enviar, Mayús+Enter salto)'
            }
            rows={2}
            disabled={mustSignInForTickets}
            onPaste={onPasteImage}
            onKeyDown={(e) => {
              if (mustSignInForTickets) return
              if (e.key === 'Enter' && !e.shiftKey) {
                e.preventDefault()
                void send()
              }
            }}
          />
          <button
            type="button"
            className="chatbot-send"
            disabled={loading || mustSignInForTickets}
            onClick={() => void send()}
            aria-label="Enviar mensaje"
          >
            <SendPlaneIcon />
          </button>
        </div>
      </section>
    </div>
  )
}

const emptyEntry = (): KnowledgeEntry => ({
  id: `kb-${Date.now()}`,
  keywords: [],
  title: '',
  response: '',
})

function KnowledgeSection() {
  const [data, setData] = useState<KnowledgePayload | null>(null)
  const [loading, setLoading] = useState(true)
  const [saving, setSaving] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [ok, setOk] = useState(false)

  useEffect(() => {
    let cancelled = false
    ;(async () => {
      setLoading(true)
      try {
        const k = await fetchKnowledge()
        if (!cancelled) setData(k)
      } catch (e) {
        if (!cancelled) setError(e instanceof Error ? e.message : 'Error')
      } finally {
        if (!cancelled) setLoading(false)
      }
    })()
    return () => {
      cancelled = true
    }
  }, [])

  const persist = async () => {
    if (!data) return
    setSaving(true)
    setError(null)
    setOk(false)
    try {
      await saveKnowledge(data)
      setOk(true)
      setTimeout(() => setOk(false), 2500)
    } catch (e) {
      setError(e instanceof Error ? e.message : 'No se pudo guardar')
    } finally {
      setSaving(false)
    }
  }

  if (loading || !data) {
    return (
      <section className="panel">
        <h2>Parámetros de respuestas</h2>
        <p className="empty-state">{error ?? 'Cargando…'}</p>
      </section>
    )
  }

  const sla = data.slaHoursByPriority

  return (
    <section className="panel" aria-labelledby="kb-heading">
      <h2 id="kb-heading">Parámetros de respuestas</h2>
      <p className="hint">
        Entradas de conocimiento y tiempos ANS por prioridad (horas hasta primera respuesta). El
        asistente las usa en el prompt del modelo.
      </p>
      {error ? <div className="error-banner">{error}</div> : null}
      {ok ? (
        <div className="hint" style={{ color: '#2e7d32', fontWeight: 600 }}>
          Guardado correctamente.
        </div>
      ) : null}

      <h3 style={{ margin: '0 0 0.5rem', fontSize: '1rem', color: '#15601d' }}>ANS por prioridad</h3>
      <div className="kb-sla">
        {(['critica', 'alta', 'media', 'baja'] as const).map((key) => (
          <label key={key}>
            {key}
            <input
              type="number"
              min={1}
              step={1}
              value={sla[key] ?? ''}
              onChange={(e) => {
                const n = parseInt(e.target.value, 10)
                setData({
                  ...data,
                  slaHoursByPriority: {
                    ...data.slaHoursByPriority,
                    [key]: Number.isNaN(n) ? 24 : n,
                  },
                })
              }}
            />
          </label>
        ))}
      </div>

      <h3 style={{ margin: '0.75rem 0 0.5rem', fontSize: '1rem', color: '#15601d' }}>
        Entradas (palabras clave y respuesta sugerida)
      </h3>
      {data.entries.map((entry, idx) => (
        <div key={entry.id} className="kb-entry">
          <header>
            <strong>{entry.title || `Entrada ${idx + 1}`}</strong>
          </header>
          <input
            placeholder="ID interno"
            value={entry.id}
            onChange={(e) => {
              const entries = [...data.entries]
              entries[idx] = { ...entry, id: e.target.value }
              setData({ ...data, entries })
            }}
          />
          <input
            placeholder="Título"
            value={entry.title}
            onChange={(e) => {
              const entries = [...data.entries]
              entries[idx] = { ...entry, title: e.target.value }
              setData({ ...data, entries })
            }}
          />
          <input
            placeholder="Palabras clave (separadas por coma)"
            value={entry.keywords.join(', ')}
            onChange={(e) => {
              const kw = e.target.value
                .split(',')
                .map((s) => s.trim())
                .filter(Boolean)
              const entries = [...data.entries]
              entries[idx] = { ...entry, keywords: kw }
              setData({ ...data, entries })
            }}
          />
          <textarea
            placeholder="Respuesta sugerida / pasos"
            value={entry.response}
            onChange={(e) => {
              const entries = [...data.entries]
              entries[idx] = { ...entry, response: e.target.value }
              setData({ ...data, entries })
            }}
          />
          <button
            type="button"
            className="btn-secondary"
            onClick={() => {
              const entries = data.entries.filter((_, i) => i !== idx)
              setData({ ...data, entries })
            }}
          >
            Eliminar entrada
          </button>
        </div>
      ))}

      <div className="kb-actions">
        <button
          type="button"
          className="btn-secondary"
          onClick={() => setData({ ...data, entries: [...data.entries, emptyEntry()] })}
        >
          Añadir entrada
        </button>
        <button type="button" className="btn-primary" disabled={saving} onClick={() => void persist()}>
          {saving ? 'Guardando…' : 'Guardar parámetros'}
        </button>
      </div>
    </section>
  )
}

function App() {
  const [tab, setTab] = useState<Tab>('chat')
  const [ticketsFromPortal, setTicketsFromPortal] = useState(false)
  const [helpdeskDeepLink, setHelpdeskDeepLink] = useState(false)
  const [sharePointListUrl, setSharePointListUrl] = useState<string | undefined>()
  const [appConfig, setAppConfig] = useState<Awaited<ReturnType<typeof fetchAppConfig>> | null>(
    null,
  )
  const [userSession, setUserSession] = useState<UserSession>(readInitialUserSession)
  const [authLoading, setAuthLoading] = useState(true)
  const [authHint, setAuthHint] = useState<string | null>(null)

  useEffect(() => {
    if (!SHOW_KNOWLEDGE_TAB && tab === 'knowledge') setTab('chat')
  }, [tab])

  const loadSession = useCallback(
    async (config: NonNullable<typeof appConfig>, interactive = false) => {
      setAuthLoading(true)
      setAuthHint(microsoftSignInHint(config))
      try {
        const session = await resolveUserSession(config, { interactive })
        setUserSession(session)
        if (interactive && !session.email) {
          setAuthHint(
            'No se pudo iniciar sesión. Verifique en Azure AD que la app tenga plataforma SPA y redirect URI ' +
              `${window.location.origin}/`,
          )
        }
      } finally {
        setAuthLoading(false)
      }
    },
    [],
  )

  useEffect(() => {
    void fetchAppConfig()
      .then(async (c) => {
        setAppConfig(c)
        setTicketsFromPortal(c.ticketsFromPortal)
        setHelpdeskDeepLink(c.helpdeskDeepLink)
        setSharePointListUrl(c.sharePointListUrl)
        await loadSession(c, false)
      })
      .catch(() => {
        setTicketsFromPortal(false)
        setHelpdeskDeepLink(false)
        setSharePointListUrl(undefined)
        setAuthLoading(false)
      })
  }, [loadSession])

  return (
    <div className="app-shell">
      <main className={tab === 'chat' ? 'main--chat' : undefined}>
        {tab === 'chat' ? (
          <ChatSection
            appConfig={appConfig}
            ticketsFromPortal={ticketsFromPortal}
            helpdeskDeepLink={helpdeskDeepLink}
            sharePointListUrl={sharePointListUrl}
            userSession={userSession}
            showMicrosoftSignIn={
              appConfig ? shouldOfferMicrosoftSignIn(appConfig, userSession) : false
            }
            authHint={authHint}
            authLoading={authLoading}
            onSignIn={() => {
              if (appConfig) void loadSession(appConfig, true)
            }}
          />
        ) : null}
        {SHOW_KNOWLEDGE_TAB && tab === 'knowledge' ? <KnowledgeSection /> : null}
      </main>
    </div>
  )
}

export default App
