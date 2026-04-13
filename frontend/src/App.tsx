import { useCallback, useEffect, useState, type ClipboardEvent } from 'react'
import './App.css'
import {
  postChat,
  fileToImagePayload,
  createTicketConfirm,
  fetchKnowledge,
  saveKnowledge,
  fetchAppConfig,
  type ChatMessage,
  type TicketDraft,
  type KnowledgePayload,
  type KnowledgeEntry,
} from './api'

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

function ChatSection({
  userEmail,
  onUserEmail,
  ticketsFromPortal,
  helpdeskDeepLink,
}: {
  userEmail: string
  onUserEmail: (v: string) => void
  ticketsFromPortal: boolean
  helpdeskDeepLink: boolean
}) {
  const [messages, setMessages] = useState<ChatMessage[]>([])
  const [input, setInput] = useState('')
  const [pendingFile, setPendingFile] = useState<File | null>(null)
  const [previewUrl, setPreviewUrl] = useState<string | null>(null)
  const [loading, setLoading] = useState(false)
  const [confirmingIndex, setConfirmingIndex] = useState<number | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [mailNotice, setMailNotice] = useState<string | null>(null)

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

  const send = useCallback(async () => {
    const trimmed = input.trim()
    if ((!trimmed && !pendingFile) || loading) return
    setError(null)
    setMailNotice(null)
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
      const res = await postChat(next, userEmail.trim() || undefined)
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
      if (res.ticketId && res.email && !res.email.sent) {
        setMailNotice(
          `Ticket ${res.ticketId} creado, pero el aviso por correo no se envió: ${res.email.error ?? 'configure RESEND_API_KEY o SMTP en el servidor'}.`,
        )
      }
    } catch (e) {
      setError(e instanceof Error ? e.message : 'Error al enviar')
    } finally {
      setLoading(false)
    }
  }, [input, loading, messages, userEmail, pendingFile, previewUrl])

  const onConfirmTicket = useCallback(
    async (messageIndex: number, draft: TicketDraft) => {
      if (confirmingIndex !== null) return
      setConfirmingIndex(messageIndex)
      setError(null)
      try {
        const { ticket, email } = await createTicketConfirm(
          draft,
          userEmail.trim() || undefined,
        )
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
        if (!email.sent) {
          setMailNotice(
            `Ticket ${ticket.id} creado, pero el correo no se envió: ${email.error ?? 'configure RESEND_API_KEY o SMTP'}.`,
          )
        } else {
          setMailNotice(null)
        }
      } catch (e) {
        setError(e instanceof Error ? e.message : 'No se pudo crear el ticket')
      } finally {
        setConfirmingIndex(null)
      }
    },
    [confirmingIndex, userEmail],
  )

  const onDismissTicket = useCallback((messageIndex: number) => {
    setMessages((prev) =>
      prev.map((msg, i) =>
        i === messageIndex ? { ...msg, ticketDraft: undefined } : msg,
      ),
    )
  }, [])

  return (
    <section className="panel" aria-labelledby="chat-heading">
      <h2 id="chat-heading">Asistente de soporte</h2>
      <p className="hint">
        Describa el error o adjunte una captura (botón o pegar con Ctrl+V / Cmd+V en el cuadro de texto).
        El asistente usa las guías oficiales de soporte cuando aplica.{' '}
        {ticketsFromPortal
          ? 'Si se ofrece un borrador de ticket, confirme con los botones.'
          : helpdeskDeepLink
            ? 'Para HelpDesk, el asistente puede generar un enlace a Power Apps con datos y una plantilla para copiar (no se crean tickets desde aquí).'
            : 'Para registrar un caso en mesa de ayuda, el asistente le dará una plantilla para copiar y pegar en HelpDesk (no se crean tickets desde aquí).'}
      </p>
      {error ? <div className="error-banner">{error}</div> : null}
      {mailNotice ? <div className="warning-banner">{mailNotice}</div> : null}
      <div className="field-row">
        <label htmlFor="user-email">
          {ticketsFromPortal ? 'Correo (opcional, para el ticket)' : 'Correo (opcional)'}
        </label>
        <input
          id="user-email"
          type="email"
          value={userEmail}
          onChange={(e) => onUserEmail(e.target.value)}
          placeholder="usuario@empresa.com"
          autoComplete="email"
        />
      </div>
      <div className="chat-messages">
        {messages.length === 0 ? (
          <p className="empty-state">Escriba su consulta para comenzar.</p>
        ) : (
          messages.map((m, i) => (
            <div
              key={`${i}-${m.role}`}
              className={`bubble ${m.role === 'user' ? 'bubble--user' : 'bubble--assistant'}`}
            >
              <span className="bubble-label">{m.role === 'user' ? 'Usted' : 'Asistente'}</span>
              {m.image ? (
                <img
                  className="bubble-image"
                  src={`data:${m.image.mimeType};base64,${m.image.dataBase64}`}
                  alt="Captura del error"
                />
              ) : null}
              <p>{m.content}</p>
              {!ticketsFromPortal && m.helpdeskUrl ? (
                <div className="helpdesk-open-wrap">
                  <a
                    className="btn-primary helpdesk-open-btn"
                    href={m.helpdeskUrl}
                    target="_blank"
                    rel="noopener noreferrer"
                  >
                    Abrir HelpDesk (Power Apps)
                  </a>
                </div>
              ) : null}
              {ticketsFromPortal && m.ticketCreatedId ? (
                <p className="ticket-created-note">
                  Ticket <strong>{m.ticketCreatedId}</strong> creado.
                </p>
              ) : null}
              {ticketsFromPortal &&
              m.role === 'assistant' &&
              m.ticketDraft &&
              !m.ticketCreatedId ? (
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
          ))
        )}
        {loading ? (
          <div className="bubble bubble--assistant">
            <span className="bubble-label">Asistente</span>
            <p>Pensando…</p>
          </div>
        ) : null}
      </div>
      <div className="chat-attach-row">
        <label className="file-attach">
          <input
            type="file"
            accept="image/png,image/jpeg,image/jpg,image/gif,image/webp"
            className="sr-only"
            onChange={(e) => onPickImage(e.target.files?.[0] ?? null)}
          />
          <span className="file-attach-label">Captura del error</span>
        </label>
        {previewUrl ? (
          <div className="preview-wrap">
            <img src={previewUrl} alt="Vista previa" className="preview-thumb" />
            <button type="button" className="btn-secondary" onClick={() => onPickImage(null)}>
              Quitar imagen
            </button>
          </div>
        ) : null}
      </div>
      <div className="chat-input-row">
        <textarea
          value={input}
          onChange={(e) => setInput(e.target.value)}
          placeholder="Ej.: No puedo entrar a la VPN… Pegue aquí una captura (Ctrl+V / Cmd+V)"
          rows={3}
          onPaste={onPasteImage}
          onKeyDown={(e) => {
            if (e.key === 'Enter' && !e.shiftKey) {
              e.preventDefault()
              void send()
            }
          }}
        />
        <button type="button" className="btn-primary" disabled={loading} onClick={() => void send()}>
          Enviar
        </button>
      </div>
    </section>
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
  const [userEmail, setUserEmail] = useState('')
  const [ticketsFromPortal, setTicketsFromPortal] = useState(false)
  const [helpdeskDeepLink, setHelpdeskDeepLink] = useState(false)

  useEffect(() => {
    if (!SHOW_KNOWLEDGE_TAB && tab === 'knowledge') setTab('chat')
  }, [tab])

  useEffect(() => {
    void fetchAppConfig()
      .then((c) => {
        setTicketsFromPortal(c.ticketsFromPortal)
        setHelpdeskDeepLink(c.helpdeskDeepLink)
      })
      .catch(() => {
        setTicketsFromPortal(false)
        setHelpdeskDeepLink(false)
      })
  }, [])

  return (
    <div className="app-shell">
      <header>
        <div className="header-top">
          <div>
            <p className="eyebrow">Periferia · Mesa de servicios</p>
            <h1>Soporte con asistente inteligente</h1>
            <p className="subtitle">
              Converse sobre errores e incidencias; el asistente aplicará las guías de Mesa de Servicios cuando
              correspondan.
              {ticketsFromPortal
                ? ' Podrá orientarle en tickets con ANS definido desde esta aplicación.'
                : helpdeskDeepLink
                  ? ' El registro en HelpDesk puede abrirse con enlace desde el chat y plantilla para copiar.'
                  : ' El registro formal del caso es en HelpDesk: el asistente le dará una plantilla para copiar y pegar.'}
            </p>
          </div>
          {SHOW_KNOWLEDGE_TAB ? (
            <nav className="tabs" aria-label="Secciones">
              <button
                type="button"
                className={tab === 'chat' ? 'active' : ''}
                onClick={() => setTab('chat')}
              >
                Chat
              </button>
              <button
                type="button"
                className={tab === 'knowledge' ? 'active' : ''}
                onClick={() => setTab('knowledge')}
              >
                Parámetros
              </button>
            </nav>
          ) : null}
        </div>
      </header>
      <main>
        {tab === 'chat' ? (
          <ChatSection
            userEmail={userEmail}
            onUserEmail={setUserEmail}
            ticketsFromPortal={ticketsFromPortal}
            helpdeskDeepLink={helpdeskDeepLink}
          />
        ) : null}
        {SHOW_KNOWLEDGE_TAB && tab === 'knowledge' ? <KnowledgeSection /> : null}
      </main>
    </div>
  )
}

export default App
