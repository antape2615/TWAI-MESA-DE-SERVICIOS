import nodemailer from 'nodemailer'
import type { Ticket } from './tickets.js'

export type EmailSendResult = {
  sent: boolean
  error?: string
  method?: 'resend' | 'smtp' | 'none'
}

function buildBody(ticket: Ticket): string {
  return [
    `Se ha creado un nuevo ticket en Mesa de Servicios Periferia.`,
    ``,
    `ID: ${ticket.id}`,
    `Título: ${ticket.title}`,
    `Prioridad: ${ticket.priority}`,
    `Categoría: ${ticket.category}`,
    `ANS (primera respuesta): ${ticket.ansHours} horas`,
    `Estado: ${ticket.status}`,
    ``,
    `Descripción:`,
    ticket.description,
    ``,
    `Posibles soluciones / pasos sugeridos:`,
    ...(ticket.possibleSolutions.length
      ? ticket.possibleSolutions.map((s) => `• ${s}`)
      : ['• (ninguna indicada)']),
    ``,
    `Creado: ${ticket.createdAt}`,
  ].join('\n')
}

function getSmtpTransport() {
  const host = process.env.SMTP_HOST?.trim()
  const port = process.env.SMTP_PORT ? parseInt(process.env.SMTP_PORT, 10) : 587
  const user = process.env.SMTP_USER?.trim()
  const pass = process.env.SMTP_PASS
  if (!host || !user || !pass) return null
  return nodemailer.createTransport({
    host,
    port,
    secure: port === 465,
    auth: { user, pass },
    requireTLS: port !== 465,
    tls: {
      minVersion: 'TLSv1.2' as const,
    },
  })
}

async function sendViaResend(
  ticket: Ticket,
  notifyTo: string,
  body: string,
): Promise<EmailSendResult> {
  const key = process.env.RESEND_API_KEY?.trim()
  if (!key) return { sent: false, method: 'none' }

  const from =
    process.env.RESEND_FROM?.trim() ||
    'Mesa Periferia <onboarding@resend.dev>'

  try {
    const res = await fetch('https://api.resend.com/emails', {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${key}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        from,
        to: [notifyTo],
        subject: `[Mesa Periferia] Nuevo ticket ${ticket.id} — ${ticket.title}`,
        text: body,
      }),
    })

    const data = (await res.json().catch(() => ({}))) as {
      message?: string
    }
    if (!res.ok) {
      const err = data.message || res.statusText || 'Error Resend'
      console.error('[email] Resend:', err)
      return { sent: false, error: err, method: 'resend' }
    }
    return { sent: true, method: 'resend' }
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e)
    console.error('[email] Resend:', msg)
    return { sent: false, error: msg, method: 'resend' }
  }
}

async function sendViaSmtp(
  subject: string,
  notifyTo: string,
  body: string,
): Promise<EmailSendResult> {
  const transport = getSmtpTransport()
  if (!transport) {
    return { sent: false, method: 'none' }
  }

  const from =
    process.env.SMTP_FROM?.trim() ||
    process.env.SMTP_USER?.trim() ||
    'no-reply@localhost'

  try {
    await transport.sendMail({
      from,
      to: notifyTo,
      subject,
      text: body,
    })
    return { sent: true, method: 'smtp' }
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e)
    console.error('[email] SMTP:', msg)
    return { sent: false, error: msg, method: 'smtp' }
  }
}

export function logEmailStartupHint(): void {
  if (process.env.RESEND_API_KEY?.trim()) {
    console.log('[email] Notificaciones: Resend (RESEND_API_KEY)')
    return
  }
  if (
    process.env.SMTP_HOST?.trim() &&
    process.env.SMTP_USER?.trim() &&
    process.env.SMTP_PASS
  ) {
    const p = process.env.SMTP_PORT || '587'
    console.log(
      `[email] Notificaciones: SMTP ${process.env.SMTP_HOST?.trim()}:${p}`,
    )
    return
  }
  console.warn(
    '[email] Sin envío de correo: defina RESEND_API_KEY o SMTP_HOST + SMTP_USER + SMTP_PASS (ver .env.example)',
  )
}

export async function sendTicketCreatedEmail(ticket: Ticket): Promise<EmailSendResult> {
  const notifyTo =
    process.env.TICKET_NOTIFY_EMAIL?.trim() ?? 'angiepena@cbit-online.com'
  const body = buildBody(ticket)

  if (process.env.RESEND_API_KEY?.trim()) {
    return sendViaResend(ticket, notifyTo, body)
  }

  if (getSmtpTransport()) {
    const subject = `[Mesa Periferia] Nuevo ticket ${ticket.id} — ${ticket.title}`
    return sendViaSmtp(subject, notifyTo, body)
  }

  console.warn(
    '[email] No hay RESEND_API_KEY ni SMTP completo. Correo no enviado. Destinatario previsto:',
    notifyTo,
  )
  console.warn('[email] Cuerpo que se habría enviado:\n', body)
  return {
    sent: false,
    error:
      'No hay método de envío configurado (RESEND_API_KEY o SMTP_HOST, SMTP_USER, SMTP_PASS)',
    method: 'none',
  }
}
