export type ResolvedUser = {
  email?: string
  name?: string
  jobTitle?: string
  department?: string
  officeLocation?: string
  phone?: string
  objectId?: string
  tenantId?: string
}

export type GraphMeProfile = {
  displayName?: string
  jobTitle?: string
  department?: string
  officeLocation?: string
  mail?: string
  userPrincipalName?: string
  mobilePhone?: string
  businessPhones?: string[]
  id?: string
}

function env(name: string): string | undefined {
  const v = process.env[name]?.trim()
  return v || undefined
}

export function azureAuthClientId(): string | undefined {
  return env('AZURE_CLIENT_ID') ?? env('SHAREPOINT_CLIENT_ID')
}

export function azureAuthEnabled(): boolean {
  if (env('AZURE_AUTH_ENABLED') === 'false') return false
  return Boolean(env('AZURE_TENANT_ID') && azureAuthClientId())
}

export function getAzureAuthConfig(): {
  enabled: boolean
  tenantId?: string
  clientId?: string
} {
  const tenantId = env('AZURE_TENANT_ID')
  const clientId = azureAuthClientId()
  return {
    enabled: azureAuthEnabled(),
    ...(tenantId ? { tenantId } : {}),
    ...(clientId ? { clientId } : {}),
  }
}

async function fetchGraphMe(accessToken: string): Promise<GraphMeProfile | null> {
  try {
    const url =
      'https://graph.microsoft.com/v1.0/me?' +
      new URLSearchParams({
        $select:
          'displayName,jobTitle,department,officeLocation,mail,userPrincipalName,mobilePhone,businessPhones,id',
      })
    const res = await fetch(url, {
      headers: { Authorization: `Bearer ${accessToken}` },
    })
    if (!res.ok) return null
    return (await res.json()) as GraphMeProfile
  } catch (e) {
    console.warn('[auth] Graph /me:', e)
    return null
  }
}

function phoneFromGraph(me: GraphMeProfile): string | undefined {
  const mobile = me.mobilePhone?.trim()
  if (mobile) return mobile
  const biz = me.businessPhones?.map((p) => p.trim()).filter(Boolean)
  if (biz?.length) return biz[0]
  return undefined
}

function userFromGraph(
  me: GraphMeProfile,
  fallback: { userEmail?: string; userName?: string },
): ResolvedUser {
  const email = (me.mail || me.userPrincipalName || fallback.userEmail)?.trim()
  const name = (me.displayName || fallback.userName)?.trim()
  return {
    email,
    name,
    jobTitle: me.jobTitle?.trim() || undefined,
    department: me.department?.trim() || undefined,
    officeLocation: me.officeLocation?.trim() || undefined,
    phone: phoneFromGraph(me),
    objectId: me.id?.trim() || undefined,
  }
}

export async function resolveUserProfile(input: {
  accessToken?: string
  userEmail?: string
  userName?: string
  jobTitle?: string
  department?: string
  officeLocation?: string
  phone?: string
}): Promise<ResolvedUser> {
  const token = input.accessToken?.trim()
  if (token) {
    const me = await fetchGraphMe(token)
    if (me) return userFromGraph(me, input)
  }

  const email = input.userEmail?.trim() || undefined
  const name = input.userName?.trim() || undefined
  const jobTitle = input.jobTitle?.trim() || undefined
  const department = input.department?.trim() || undefined
  const officeLocation = input.officeLocation?.trim() || undefined
  const phone = input.phone?.trim() || undefined
  if (!email && !name && !jobTitle && !department) return {}
  return { email, name, jobTitle, department, officeLocation, phone }
}
