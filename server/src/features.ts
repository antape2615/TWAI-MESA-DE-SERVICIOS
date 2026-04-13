/**
 * Con `TICKETS_FROM_PORTAL_ENABLED=true` se habilitan herramientas de ticket en el chat y POST /api/tickets.
 * Por defecto (sin variable o distinto de `true`): solo guía con plantilla HelpDesk.
 */
export function ticketsFromPortalEnabled(): boolean {
  return process.env.TICKETS_FROM_PORTAL_ENABLED === 'true'
}
