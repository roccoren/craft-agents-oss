/**
 * Teams (Azure Bot) credentials: stored as JSON in the `messaging_bearer`
 * credential row, `name = 'teams'`.
 */
export interface TeamsCredentials {
  appId: string
  appPassword: string
  /** Present ⇒ single-tenant app; absent ⇒ multi-tenant. */
  tenantId?: string
}

export function parseTeamsCredentials(raw: string): TeamsCredentials {
  let parsed: unknown
  try {
    parsed = JSON.parse(raw)
  } catch {
    throw new Error('Teams credentials are not valid JSON')
  }
  const obj = parsed as { appId?: unknown; appPassword?: unknown; tenantId?: unknown }
  if (typeof obj !== 'object' || obj === null) {
    throw new Error('Teams credentials must be an object')
  }
  if (typeof obj.appId !== 'string' || obj.appId.length === 0) {
    throw new Error('Teams credentials must include a non-empty "appId"')
  }
  if (typeof obj.appPassword !== 'string' || obj.appPassword.length === 0) {
    throw new Error('Teams credentials must include a non-empty "appPassword"')
  }
  const creds: TeamsCredentials = { appId: obj.appId, appPassword: obj.appPassword }
  if (typeof obj.tenantId === 'string' && obj.tenantId.length > 0) creds.tenantId = obj.tenantId
  return creds
}

/**
 * Validate credentials by acquiring a Bot Connector app token via the OAuth2
 * client-credentials flow. Single-tenant uses the tenant authority; multi-tenant
 * uses the shared `botframework.com` authority. No inbound traffic required.
 */
export async function testTeamsCredentials(
  creds: TeamsCredentials,
  fetchImpl: typeof fetch = fetch,
): Promise<{ success: boolean; botName?: string; error?: string }> {
  if (!creds.appId || !creds.appPassword) {
    return { success: false, error: 'App ID and password are required' }
  }
  const authority = creds.tenantId ?? 'botframework.com'
  const url = `https://login.microsoftonline.com/${authority}/oauth2/v2.0/token`
  const form = new URLSearchParams({
    grant_type: 'client_credentials',
    client_id: creds.appId,
    client_secret: creds.appPassword,
    scope: 'https://api.botframework.com/.default',
  })
  try {
    const res = await fetchImpl(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: form.toString(),
    })
    if (res.status === 401 || res.status === 400) {
      return { success: false, error: 'Invalid App ID or password' }
    }
    if (!res.ok) {
      return { success: false, error: `Azure AD error (${res.status})` }
    }
    const body = (await res.json()) as { access_token?: string }
    if (!body.access_token) {
      return { success: false, error: 'No access token returned' }
    }
    return { success: true, botName: creds.appId }
  } catch (err) {
    return { success: false, error: err instanceof Error ? err.message : 'Network error' }
  }
}
