import type { AccountInfo } from '@azure/msal-browser'

export type AccountProfile = {
  displayName: string
  username: string
  preferredUsername: string
  emailFromClaims: string
  localAccountId: string
  homeAccountId: string
  tenantId: string
  objectId: string
  environment: string
  applicationRolesDisplay: string
  directoryRoleIdsDisplay: string
}

export type GraphUserProfile = {
  displayName?: string
  jobTitle?: string
  department?: string
  officeLocation?: string
  mail?: string
  userPrincipalName?: string
  mobilePhone?: string
  businessPhones?: string[]
}

function formatApplicationRoles(claims: Record<string, unknown>): string {
  const raw = claims.roles
  if (Array.isArray(raw)) return raw.filter(Boolean).join(', ')
  if (typeof raw === 'string' && raw.trim()) return raw.trim()
  return ''
}

function formatDirectoryRoleIds(claims: Record<string, unknown>): string {
  const w = claims.wids
  if (Array.isArray(w) && w.length) return w.map(String).join(', ')
  return ''
}

export function buildAccountProfile(account: AccountInfo | null): AccountProfile | null {
  if (!account) return null
  const claims =
    account.idTokenClaims && typeof account.idTokenClaims === 'object'
      ? (account.idTokenClaims as Record<string, unknown>)
      : {}
  const preferred = String(claims.preferred_username || claims.upn || '')
  const claimEmail =
    String(claims.email || '') ||
    (Array.isArray(claims.emails) ? String(claims.emails[0] || '') : '') ||
    ''
  return {
    displayName: account.name || String(claims.name || ''),
    username: account.username || '',
    preferredUsername: preferred,
    emailFromClaims: claimEmail,
    localAccountId: account.localAccountId || '',
    homeAccountId: account.homeAccountId || '',
    tenantId: String(claims.tid || account.tenantId || ''),
    objectId: String(claims.oid || claims.sub || ''),
    environment: account.environment || '',
    applicationRolesDisplay: formatApplicationRoles(claims),
    directoryRoleIdsDisplay: formatDirectoryRoleIds(claims),
  }
}

export async function fetchGraphProfile(accessToken: string): Promise<GraphUserProfile | null> {
  try {
    const url =
      'https://graph.microsoft.com/v1.0/me?' +
      new URLSearchParams({
        $select:
          'displayName,jobTitle,department,officeLocation,mail,userPrincipalName,mobilePhone,businessPhones',
      })
    const res = await fetch(url, {
      headers: { Authorization: `Bearer ${accessToken}` },
    })
    if (!res.ok) return null
    return (await res.json()) as GraphUserProfile
  } catch {
    return null
  }
}
