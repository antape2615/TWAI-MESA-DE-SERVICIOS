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
