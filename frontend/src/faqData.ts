/**
 * Fuente única: `data/soporte-faq.json` (exportado desde Consulta Soporte.xlsx).
 * Para actualizar: edite el JSON o regenere desde Excel y mantenga alineado el servidor.
 */
import soporteFaq from '../../data/soporte-faq.json'

export type FaqItem = {
  id: string
  title: string
  category: string
  content: string
}

export const FAQ_ITEMS: FaqItem[] = soporteFaq as FaqItem[]

export function promptFromFaq(item: FaqItem): string {
  return `[Consulta desde FAQ: ${item.title}]\n\n${item.content}`
}
