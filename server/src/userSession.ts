import { resolveUserProfile, type ResolvedUser } from './azureAuth.js'

export type UserSessionInput = {
  accessToken?: string
  userEmail?: string
  userName?: string
  jobTitle?: string
  department?: string
  officeLocation?: string
  phone?: string
}

function str(v: unknown): string | undefined {
  return typeof v === 'string' && v.trim() ? v.trim() : undefined
}

export function parseUserSessionBody(body: Record<string, unknown>): UserSessionInput {
  return {
    accessToken: str(body.accessToken),
    userEmail: str(body.userEmail),
    userName: str(body.userName),
    jobTitle: str(body.jobTitle),
    department: str(body.department),
    officeLocation: str(body.officeLocation),
    phone: str(body.phone),
  }
}

export async function resolveSessionUser(
  input: UserSessionInput,
): Promise<ResolvedUser> {
  return resolveUserProfile(input)
}
