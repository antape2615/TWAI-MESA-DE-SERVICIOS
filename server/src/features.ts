import { sharePointTicketsEnabled } from './sharepoint.js'

/**
 * Tickets desde el chat/API cuando:
 * - `TICKETS_FROM_PORTAL_ENABLED=true`, o
 * - SharePoint está configurado (lista RPA_SOLICITUD_TICKET_SOPORTE u otra).
 * Sin ninguna de las dos: solo guía con plantilla HelpDesk.
 */
export function ticketsFromPortalEnabled(): boolean {
  return (
    process.env.TICKETS_FROM_PORTAL_ENABLED === 'true' || sharePointTicketsEnabled()
  )
}
