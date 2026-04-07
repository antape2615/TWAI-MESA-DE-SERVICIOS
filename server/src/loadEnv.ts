import path from 'node:path'
import { fileURLToPath } from 'node:url'
import dotenv from 'dotenv'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const serverDir = path.join(__dirname, '..')
const repoRoot = path.join(__dirname, '..', '..')
dotenv.config({ path: path.join(repoRoot, '.env') })
dotenv.config({ path: path.join(serverDir, '.env'), override: true })
