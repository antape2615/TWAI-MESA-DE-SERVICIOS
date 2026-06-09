import type { AccountInfo } from '@azure/msal-browser'
import type { AppConfig } from './api'
import {
  accountFromResult,
  acquireTicketAuthTokens,
  isEmbeddedFrame,
  signInWithRedirect,
  trySilentAuth,
} from './azureAuth'
import {
  buildAccountProfile,
  fetchGraphProfile,
  type AccountProfile,
  type GraphUserProfile,
} from './authProfile'
import { readUserFromUrl } from './urlDisplayName'

export type UserSession = {
  name: string
  email: string
  accessToken?: string
  sharePointAccessToken?: string
  source: 'azure' | 'url' | 'none'
  jobTitle?: string
  department?: string
  officeLocation?: string
  phone?: string
  accountProfile?: AccountProfile | null
  graphProfile?: GraphUserProfile | null
}

const emptySession = (): UserSession => ({
  name: '',
  email: '',
  source: 'none',
})

function sessionFromUrl(): UserSession | null {
  const fromUrl = readUserFromUrl()
  if (!fromUrl.name && !fromUrl.email) return null
  return { ...fromUrl, source: 'url' }
}

export function readInitialUserSession(): UserSession {
  return sessionFromUrl() ?? emptySession()
}

async function enrichWithGraph(
  base: UserSession,
  accessToken?: string,
  account?: AccountInfo | null,
): Promise<UserSession> {
  if (!accessToken) return base
  const graphProfile = await fetchGraphProfile(accessToken)
  const accountProfile = buildAccountProfile(account ?? null)

  const phone =
    graphProfile?.mobilePhone?.trim() ||
    graphProfile?.businessPhones?.map((p) => p.trim()).find(Boolean) ||
    undefined

  return {
    ...base,
    name:
      graphProfile?.displayName?.trim() ||
      base.name ||
      accountProfile?.displayName ||
      '',
    email:
      graphProfile?.mail?.trim() ||
      graphProfile?.userPrincipalName?.trim() ||
      base.email ||
      accountProfile?.username ||
      '',
    jobTitle: graphProfile?.jobTitle?.trim() || undefined,
    department: graphProfile?.department?.trim() || undefined,
    officeLocation: graphProfile?.officeLocation?.trim() || undefined,
    phone,
    accountProfile,
    graphProfile,
  }
}

export async function resolveUserSession(
  config: AppConfig,
  options: { interactive?: boolean } = {},
): Promise<UserSession> {
  const fromUrl = readUserFromUrl()

  if (config.azureAuth?.enabled) {
    const silent = await trySilentAuth(config, fromUrl.email || undefined)
    const fromAzure = accountFromResult(silent)
    if (fromAzure) {
      const base: UserSession = {
        name: fromAzure.name || fromUrl.name,
        email: fromAzure.email || fromUrl.email,
        accessToken: fromAzure.accessToken,
        source: 'azure',
      }
      const enriched = await enrichWithGraph(base, fromAzure.accessToken, silent?.account)
      if (config.sharePointTickets && config.sharePointResourceOrigin) {
        const tokens = await acquireTicketAuthTokens(config, enriched.email || fromUrl.email)
        if (tokens?.accessToken) enriched.accessToken = tokens.accessToken
        if (tokens?.sharePointAccessToken) {
          enriched.sharePointAccessToken = tokens.sharePointAccessToken
        }
      }
      return enriched
    }

    if (options.interactive) {
      const started = await signInWithRedirect(config, fromUrl.email || undefined)
      if (started === 'redirecting') return sessionFromUrl() ?? emptySession()
    }
  }

  const urlSession = sessionFromUrl()
  if (urlSession) return urlSession

  return emptySession()
}

/** Renueva tokens Graph + SharePoint antes de crear un ticket. */
export async function refreshTicketAuthTokens(
  config: AppConfig,
  session: UserSession,
): Promise<UserSession> {
  if (!config.azureAuth?.enabled || !config.sharePointTickets) return session
  const tokens = await acquireTicketAuthTokens(config, session.email || undefined)
  if (!tokens) return session
  return {
    ...session,
    ...(tokens.accessToken ? { accessToken: tokens.accessToken } : {}),
    ...(tokens.sharePointAccessToken
      ? { sharePointAccessToken: tokens.sharePointAccessToken }
      : {}),
  }
}

export function shouldOfferMicrosoftSignIn(
  config: AppConfig,
  session: UserSession,
): boolean {
  if (!config.azureAuth?.enabled) return false
  if (session.email) return false
  return true
}

export function microsoftSignInHint(config: AppConfig): string | null {
  if (!config.azureAuth?.enabled) return null
  if (isEmbeddedFrame()) {
    return 'Para iniciar sesión con Microsoft, abra el chatbot en una ventana completa (no embebido en Power Apps).'
  }
  return null
}
