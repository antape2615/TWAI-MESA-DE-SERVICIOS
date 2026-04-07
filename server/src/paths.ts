import path from 'node:path'
import { fileURLToPath } from 'node:url'

const __dirname = path.dirname(fileURLToPath(import.meta.url))

export const DATA_DIR = path.join(__dirname, '..', 'data')
export const KNOWLEDGE_FILE = path.join(DATA_DIR, 'knowledge.json')
export const TICKETS_FILE = path.join(DATA_DIR, 'tickets.json')
