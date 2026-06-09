import {
  PublicClientApplication,
  type AccountInfo,
  type AuthenticationResult,
} from '@azure/msal-browser'
import type { AppConfig } from './api'

let msalApp: PublicClientApplication | null = null
let msalReady: Promise<PublicClientApplication | null> | null = null
let redirectChecked = false
let redirectResult: AuthenticationResult | null = null

/** Graph: perfil + /sites/.../users. SharePoint: ensureuser para SolicitadoPor. */
const GRAPH_SCOPES = ['User.Read', 'Sites.ReadWrite.All']

function sharePointScopes(host: string): string[] {
  return [`${host.replace(/\/$/, '')}/AllSites.Write`]
}

function allTicketScopes(sharePointHost?: string): string[] {
  if (!sharePointHost) return GRAPH_SCOPES
  return [...GRAPH_SCOPES, ...sharePointScopes(sharePointHost)]
}

function isInteractionRequired(err: unknown): boolean {
  if (!err || typeof err !== 'object') return false
  const e = err as { errorCode?: string; name?: string }
  return (
    e.errorCode === 'interaction_required' ||
    e.errorCode === 'consent_required' ||
    e.errorCode === 'login_required' ||
    e.name === 'InteractionRequiredAuthError'
  )
}

export type TicketAuthTokens = {
  accessToken?: string
  sharePointAccessToken?: string
}

function redirectUri(): string {
  return `${window.location.origin}/`
}

function authConfigFromApp(config: AppConfig) {
  const auth = config.azureAuth
  if (!auth?.enabled || !auth.tenantId || !auth.clientId) return null
  return { tenantId: auth.tenantId, clientId: auth.clientId }
}

export function isEmbeddedFrame(): boolean {
  try {
    return window.self !== window.top
  } catch {
    return true
  }
}

async function getMsal(config: AppConfig): Promise<PublicClientApplication | null> {
  const auth = authConfigFromApp(config)
  if (!auth) return null
  if (msalApp) return msalApp
  if (msalReady) return msalReady

  msalReady = (async () => {
    const app = new PublicClientApplication({
      auth: {
        clientId: auth.clientId,
        authority: `https://login.microsoftonline.com/${auth.tenantId}`,
        redirectUri: redirectUri(),
        postLogoutRedirectUri: redirectUri(),
      },
      cache: { cacheLocation: 'localStorage' },
    })
    await app.initialize()
    msalApp = app
    return app
  })()

  return msalReady
}

/** Debe llamarse al cargar la app para capturar el retorno de Microsoft. */
export async function processRedirectReturn(
  config: AppConfig,
): Promise<AuthenticationResult | null> {
  if (redirectChecked) return redirectResult
  const app = await getMsal(config)
  if (!app) {
    redirectChecked = true
    return null
  }

  try {
    redirectResult = await app.handleRedirectPromise()
  } catch (e) {
    console.error('[auth] Error al procesar retorno de Microsoft:', e)
    redirectResult = null
  } finally {
    redirectChecked = true
  }

  if (redirectResult) {
    cleanAuthFromUrl()
  }

  return redirectResult
}

function cleanAuthFromUrl(): void {
  const url = new URL(window.location.href)
  const authKeys = ['code', 'state', 'session_state', 'client_info']
  let changed = false
  for (const key of authKeys) {
    if (url.searchParams.has(key)) {
      url.searchParams.delete(key)
      changed = true
    }
  }
  if (url.hash && /code=|state=|error=/.test(url.hash)) {
    url.hash = ''
    changed = true
  }
  if (changed) {
    window.history.replaceState({}, document.title, `${url.pathname}${url.search}${url.hash}`)
  }
}

async function acquireSilent(
  app: PublicClientApplication,
  account: AccountInfo,
  sharePointHost?: string,
): Promise<AuthenticationResult | null> {
  try {
    return await app.acquireTokenSilent({
      scopes: allTicketScopes(sharePointHost),
      account,
    })
  } catch {
    return null
  }
}

async function acquireSharePointSilent(
  app: PublicClientApplication,
  account: AccountInfo,
  sharePointHost: string,
): Promise<string | undefined> {
  try {
    const result = await app.acquireTokenSilent({
      scopes: sharePointScopes(sharePointHost),
      account,
    })
    return result.accessToken
  } catch {
    return undefined
  }
}

/** Tokens con permisos para rellenar Solicitado Por en SharePoint. */
export async function acquireTicketAuthTokens(
  config: AppConfig,
  loginHint?: string,
): Promise<TicketAuthTokens | null> {
  const redirect = await processRedirectReturn(config)
  const app = await getMsal(config)
  if (!app) return null

  const host = config.sharePointResourceOrigin
  let account = redirect?.account ?? app.getAllAccounts()[0] ?? null

  if (!account && loginHint?.includes('@')) {
    try {
      const sso = await app.ssoSilent({
        scopes: allTicketScopes(host),
        loginHint,
      })
      account = sso.account
    } catch {
      /* sin cuenta */
    }
  }

  if (!account) return null

  try {
    const graph = await app.acquireTokenSilent({
      scopes: allTicketScopes(host),
      account,
    })
    const sharePointAccessToken =
      host != null ? await acquireSharePointSilent(app, account, host) : undefined
    return {
      accessToken: graph.accessToken,
      sharePointAccessToken,
    }
  } catch (err) {
    if (!isInteractionRequired(err)) return null
    try {
      await app.acquireTokenRedirect({
        scopes: allTicketScopes(host),
        account,
        prompt: 'consent',
      })
    } catch (e) {
      console.error('[auth] No se pudo renovar consentimiento SharePoint:', e)
    }
    return null
  }
}

export async function trySilentAuth(
  config: AppConfig,
  loginHint?: string,
): Promise<AuthenticationResult | null> {
  const redirect = await processRedirectReturn(config)
  if (redirect?.account) return redirect

  const app = await getMsal(config)
  if (!app) return null

  const host = config.sharePointResourceOrigin
  const accounts = app.getAllAccounts()
  if (accounts.length > 0) {
    const token = await acquireSilent(app, accounts[0], host)
    if (token) return token
  }

  if (loginHint?.includes('@')) {
    try {
      return await app.ssoSilent({ scopes: allTicketScopes(host), loginHint })
    } catch {
      return null
    }
  }

  return null
}

/**
 * Login en la misma ventana (no popup). Si está embebido en iframe, abre la ventana superior.
 */
export async function signInWithRedirect(
  config: AppConfig,
  loginHint?: string,
): Promise<'redirecting' | null> {
  const app = await getMsal(config)
  if (!app) return null

  const request = {
    scopes: allTicketScopes(config.sharePointResourceOrigin),
    loginHint,
    prompt: 'consent' as const,
  }

  if (isEmbeddedFrame()) {
    try {
      window.top!.location.href = window.location.href
      return 'redirecting'
    } catch {
      /* sin acceso a top; continuar */
    }
  }

  try {
    await app.loginRedirect(request)
    return 'redirecting'
  } catch (e) {
    console.error('[auth] loginRedirect falló:', e)
    return null
  }
}

export function accountFromResult(
  result: AuthenticationResult | null,
): { name: string; email: string; accessToken?: string } | null {
  if (!result) return null
  const account = result.account
  const email =
    account?.username?.trim() ||
    (account?.idTokenClaims?.preferred_username as string | undefined)?.trim() ||
    (account?.idTokenClaims?.email as string | undefined)?.trim() ||
    ''
  const name =
    account?.name?.trim() ||
    (account?.idTokenClaims?.name as string | undefined)?.trim() ||
    email
  if (!email && !name) return null
  return {
    name,
    email,
    accessToken: result.accessToken,
  }
}
