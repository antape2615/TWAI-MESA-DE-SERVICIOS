import fs from 'node:fs/promises'
import path from 'node:path'
import { TICKETS_FILE } from './paths.js'
import { z } from 'zod'

const Priority = z.enum(['baja', 'media', 'alta', 'critica'])
const Status = z.enum(['abierto', 'en_progreso', 'resuelto', 'cerrado'])

export const TicketSchema = z.object({
  id: z.string(),
  title: z.string(),
  description: z.string(),
  category: z.string(),
  priority: Priority,
  status: Status,
  ansHours: z.number(),
  possibleSolutions: z.array(z.string()),
  createdAt: z.string(),
  userEmail: z.string().optional(),
})

export type Ticket = z.infer<typeof TicketSchema>

const StoreSchema = z.object({ tickets: z.array(TicketSchema) })

async function readStore(): Promise<{ tickets: Ticket[] }> {
  try {
    const raw = await fs.readFile(TICKETS_FILE, 'utf-8')
    return StoreSchema.parse(JSON.parse(raw))
  } catch {
    return { tickets: [] }
  }
}

async function writeStore(data: { tickets: Ticket[] }): Promise<void> {
  await fs.mkdir(path.dirname(TICKETS_FILE), { recursive: true })
  await fs.writeFile(TICKETS_FILE, JSON.stringify(data, null, 2), 'utf-8')
}

function nextId(existing: Ticket[]): string {
  const nums = existing
    .map((t) => {
      const m = /^TK-(\d+)$/.exec(t.id)
      return m ? parseInt(m[1], 10) : 0
    })
    .filter((n) => !Number.isNaN(n))
  const max = nums.length ? Math.max(...nums) : 0
  return `TK-${String(max + 1).padStart(5, '0')}`
}

export async function listTickets(): Promise<Ticket[]> {
  const { tickets } = await readStore()
  return [...tickets].sort(
    (a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime(),
  )
}

export async function createTicket(input: {
  title: string
  description: string
  category: string
  priority: z.infer<typeof Priority>
  ansHours: number
  possibleSolutions: string[]
  userEmail?: string
}): Promise<Ticket> {
  const { tickets } = await readStore()
  const ticket: Ticket = {
    id: nextId(tickets),
    title: input.title,
    description: input.description,
    category: input.category,
    priority: input.priority,
    status: 'abierto',
    ansHours: input.ansHours,
    possibleSolutions: input.possibleSolutions,
    createdAt: new Date().toISOString(),
    userEmail: input.userEmail,
  }
  tickets.push(ticket)
  await writeStore({ tickets })
  return ticket
}
