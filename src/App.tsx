import { useEffect, useRef, useState } from 'react'
import { PublicClientApplication } from '@azure/msal-browser'
import * as microsoftTeams from '@microsoft/teams-js'
import './App.css'

type AuthStatus = 'loading' | 'waitingConsent' | 'manual' | 'success' | 'error'
type HostEnvironment = 'unknown' | 'teams-desktop' | 'teams-web' | 'teams-mobile' | 'standalone'

interface UserInfo {
  displayName?: string
  email?: string
  userPrincipalName?: string
  id?: string
  tenantId?: string
  chineseSurname?: string
  chineseGivenName?: string
  givenName?: string
  surname?: string
  jobTitle?: string
  department?: string
  officeLocation?: string
  mobilePhone?: string
  businessPhones?: string[]
  preferredLanguage?: string
}

const CLIENT_ID = '33abd69a-d012-498a-bddb-8608cbf10c2d'
const TENANT_ID = 'cd4e36bd-ac9a-4236-9f91-a6718b6b5e45'
const GRAPH_SCOPES = ['User.Read']

const msalConfig = {
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    redirectUri: window.location.origin + window.location.pathname
  },
  cache: {
    cacheLocation: 'sessionStorage' as const,
    storeAuthStateInCookie: false
  }
}

function App() {
  const [status, setStatus] = useState<AuthStatus>('loading')
  const [userInfo, setUserInfo] = useState<UserInfo | null>(null)
  const [errorMessage, setErrorMessage] = useState<string>('')
  const [environment, setEnvironment] = useState<HostEnvironment>('unknown')
  const [manualAuthInProgress, setManualAuthInProgress] = useState(false)
  const msalInstanceRef = useRef<PublicClientApplication | null>(null)
  const teamsContextRef = useRef<microsoftTeams.app.Context | null>(null)

  useEffect(() => {
    const init = async () => {
      try {
        setStatus('loading')
        await microsoftTeams.app.initialize()
        const context = await microsoftTeams.app.getContext()
        teamsContextRef.current = context
        const clientType = context.app?.host?.clientType ?? 'unknown'
        const detectedEnv: HostEnvironment =
          clientType === 'desktop' ? 'teams-desktop'
            : clientType === 'web' ? 'teams-web'
              : clientType === 'android' || clientType === 'ios' ? 'teams-mobile'
                : 'standalone'

        setEnvironment(detectedEnv)

        const isTeamsHost =
          detectedEnv === 'teams-desktop' ||
          detectedEnv === 'teams-web' ||
          detectedEnv === 'teams-mobile'

        if (isTeamsHost) {
          await loginViaTeams(context)
        } else {
          await loginViaMsal(context.user?.userPrincipalName, context.user?.tenant?.id)
        }
      } catch (error) {
        console.info('Not running inside Teams,改採瀏覽器流程', error)
        setEnvironment('standalone')
        try {
          await loginViaMsal()
        } catch (msalError) {
          console.error('登入流程失敗', msalError)
          const message = normalizeErrorMessage(msalError)
          setErrorMessage(message)
          setStatus('error')
        }
      }
    }

    void init()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  const loginViaTeams = async (context: microsoftTeams.app.Context) => {
    try {
      setStatus('waitingConsent')
      const teamsUser = context.user as TeamsUser | undefined
      const loginHint = teamsUser?.userPrincipalName || teamsUser?.email

      let tokenResult = getCachedGraphToken()
      const requireManual =
        environment === 'teams-web' && isSafari() && !tokenResult

      if (requireManual) {
        setStatus('manual')
        return
      }

      if (!tokenResult) {
        tokenResult = await requestGraphToken(loginHint)
      }
      const graphProfile = await fetchGraphProfile(tokenResult.token)
      const user = deriveUserInfo({
        teamsUser,
        graphProfile,
        tenantId: context.user?.tenant?.id,
        tokenPayload: tokenResult.payload
      })

      setUserInfo(user)
      setStatus('success')
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error)
      if (
        message.includes('embedded browser') ||
        message.includes('NotAllowedInThisContext') ||
        message.includes('Authentication cancelled by user') ||
        message.includes('was unable to be opened')
      ) {
        console.warn('偵測到 Safari 需要手動授權:', message)
        setStatus('manual')
        return
      }

      console.error('Teams 登入流程失敗', error)
      const friendly = normalizeErrorMessage(error)
      setErrorMessage(friendly)
      setStatus('error')
    }
  }

  const handleManualAuth = async () => {
    if (manualAuthInProgress) {
      return
    }

    try {
      setManualAuthInProgress(true)
      setStatus('waitingConsent')

      const context = teamsContextRef.current
      if (!context) {
        throw new Error('尚未取得 Teams context，請重新整理頁面後重試。')
      }

      const teamsUser = context.user as TeamsUser | undefined
      const loginHint = teamsUser?.userPrincipalName || teamsUser?.email

      const tokenResult = await requestGraphToken(loginHint)
      const graphProfile = await fetchGraphProfile(tokenResult.token)
      const user = deriveUserInfo({
        teamsUser,
        graphProfile,
        tenantId: context.user?.tenant?.id,
        tokenPayload: tokenResult.payload
      })

      setUserInfo(user)
      setStatus('success')
    } catch (error) {
      console.error('手動授權失敗', error)
      const message = normalizeErrorMessage(error)
      setErrorMessage(message)
      setStatus('error')
    } finally {
      setManualAuthInProgress(false)
    }
  }

  const loginViaMsal = async (loginHint?: string, tenantId?: string) => {
    try {
      setStatus('waitingConsent')
        const tokenResult = await ensureMsalGraphToken(loginHint)
        if (!tokenResult) {
          return
        }

        const graphProfile = await fetchGraphProfile(tokenResult.token)
        const user = deriveUserInfo({
          graphProfile,
          tenantId: tenantId ?? tokenResult.payload?.tid,
          tokenPayload: tokenResult.payload
        })

        setUserInfo(user)
        setStatus('success')
    } catch (error) {
      console.error('MSAL 登入流程失敗', error)
      const message = normalizeErrorMessage(error)
      setErrorMessage(message)
      setStatus('error')
    }
  }

  const ensureMsalGraphToken = async (loginHint?: string): Promise<CachedGraphToken> => {
    const cached = getCachedGraphToken()
    if (cached) {
      return cached
    }

    if (!msalInstanceRef.current) {
      msalInstanceRef.current = new PublicClientApplication(msalConfig)
      await msalInstanceRef.current.initialize()
    }

    const msalInstance = msalInstanceRef.current
    const redirectResult = await msalInstance.handleRedirectPromise()
    if (redirectResult?.accessToken) {
      return cacheGraphToken(redirectResult.accessToken)
    }

    const accounts = msalInstance.getAllAccounts()
    if (accounts.length > 0) {
      try {
        const tokenResponse = await msalInstance.acquireTokenSilent({
          scopes: GRAPH_SCOPES,
          account: accounts[0]
        })
        return cacheGraphToken(tokenResponse.accessToken)
      } catch (error) {
        await msalInstance.acquireTokenRedirect({
          scopes: GRAPH_SCOPES,
          account: accounts[0],
          loginHint,
          redirectUri: window.location.href
        })
        throw new Error('redirect_pending')
      }
    }

    await msalInstance.loginRedirect({
      scopes: GRAPH_SCOPES,
      loginHint,
      redirectUri: window.location.href
    })
    throw new Error('redirect_pending')
  }

  const environmentInfo = getEnvironmentMessage(environment)

  return (
    <div className="app">
      <div className="container">
        {environmentInfo && (
          <div className={`environment-banner environment-${environment}`}>
            <strong>{environmentInfo.title}</strong>
            <span>{environmentInfo.message}</span>
          </div>
        )}

        {status === 'loading' && (
          <div className="status-card loading">
            <div className="spinner"></div>
            <h1>正在建立連線...</h1>
            <p>等待 Teams 提供目前使用者資訊</p>
          </div>
        )}

        {status === 'waitingConsent' && (
          <div className="status-card loading">
            <div className="spinner"></div>
            <h1>請授權存取 Microsoft 365</h1>
            <p>若未看到 Teams 對話框，請檢查是否被視窗擋住或被彈出視窗阻擋</p>
          </div>
        )}

        {status === 'manual' && (
          <div className="status-card manual">
            <div className="manual-icon">⚠️</div>
            <h1>Safari 需手動授權</h1>
            <p>
              Safari 網頁版 Teams 會阻擋自動授權。請點擊下方按鈕在新視窗完成授權，授權後視窗可關閉並回到此頁。
            </p>
            <button
              className="manual-button"
              onClick={handleManualAuth}
              disabled={manualAuthInProgress}
            >
              {manualAuthInProgress ? '授權進行中…' : '在新視窗授權'}
            </button>
            <small>授權完成後重新整理，此提示將不再出現。</small>
          </div>
        )}

        {status === 'success' && userInfo && (
          <div className="status-card success">
            <div className="success-icon">✓</div>
            <h1>登入成功！</h1>
            <div className="user-info">
              {renderUserInfoRows(userInfo).map(item => (
                <p key={item.label}>
                  <strong>{item.label}：</strong>{item.value}
                </p>
              ))}
            </div>
          </div>
        )}

        {status === 'error' && (
          <div className="status-card error">
            <div className="error-icon">✗</div>
            <h1>登入失敗</h1>
            <p className="error-message">{errorMessage || '發生未知錯誤'}</p>
            <p className="error-hint">可以重新整理此頁，或關閉後再從 Teams 個人應用程式重新開啟</p>
          </div>
        )}
      </div>
    </div>
  )
}

export default App

interface TeamsUser {
  displayName?: string
  email?: string
  userPrincipalName?: string
  id?: string
  aadObjectId?: string
}

interface GraphProfile {
  id?: string
  displayName?: string
  userPrincipalName?: string
  mail?: string
  givenName?: string
  surname?: string
  jobTitle?: string
  department?: string
  officeLocation?: string
  mobilePhone?: string
  businessPhones?: string[]
  preferredLanguage?: string
}

interface TokenPayload {
  tid?: string
  preferred_username?: string
  name?: string
  given_name?: string
  family_name?: string
  upn?: string
  exp?: number
}

interface DeriveParams {
  teamsUser?: TeamsUser
  graphProfile?: GraphProfile
  tenantId?: string
  tokenPayload?: TokenPayload
}

function deriveUserInfo(params: DeriveParams): UserInfo {
  const { teamsUser, graphProfile, tenantId, tokenPayload } = params

  const displayName = graphProfile?.displayName ?? teamsUser?.displayName
  const email = graphProfile?.mail ?? teamsUser?.email ?? tokenPayload?.preferred_username ?? teamsUser?.userPrincipalName
  const userPrincipalName = graphProfile?.userPrincipalName ?? teamsUser?.userPrincipalName ?? tokenPayload?.preferred_username
  const id = graphProfile?.id ?? teamsUser?.id ?? teamsUser?.aadObjectId

  const preferSurname = graphProfile?.surname ?? tokenPayload?.family_name
  const preferGivenName = graphProfile?.givenName ?? tokenPayload?.given_name

  const fallback = splitChineseName(displayName)

  return {
    displayName,
    email,
    userPrincipalName,
    id,
    tenantId: tenantId ?? tokenPayload?.tid,
    chineseSurname: preferSurname ?? fallback.surname,
    chineseGivenName: preferGivenName ?? fallback.givenName,
    givenName: graphProfile?.givenName ?? tokenPayload?.given_name,
    surname: graphProfile?.surname ?? tokenPayload?.family_name,
    jobTitle: graphProfile?.jobTitle,
    department: graphProfile?.department,
    officeLocation: graphProfile?.officeLocation,
    mobilePhone: graphProfile?.mobilePhone,
    businessPhones: graphProfile?.businessPhones,
    preferredLanguage: graphProfile?.preferredLanguage
  }
}

let graphAuthPromise: Promise<CachedGraphToken> | null = null

async function requestGraphToken(loginHint?: string): Promise<CachedGraphToken> {
  const cached = getCachedGraphToken()
  if (cached) {
    return cached
  }

  if (graphAuthPromise) {
    return graphAuthPromise
  }

  graphAuthPromise = new Promise((resolve, reject) => {
    const hintParam = loginHint ? `?loginHint=${encodeURIComponent(loginHint)}` : ''

    const authUrl = `${window.location.origin}/auth.html${hintParam}`
    const windowFeatures = 'noopener,noreferrer,width=600,height=535'

    const openAuthWindow = () => {
      const authWindow = window.open(authUrl, '_blank', windowFeatures)
      if (!authWindow) {
        throw new Error('無法開啟授權視窗，請允許彈出視窗或改用桌面版。')
      }

      const checkClosed = setInterval(() => {
        if (authWindow.closed) {
          clearInterval(checkClosed)
          if (graphAuthPromise) {
            graphAuthPromise = null
            reject(new Error('使用者已關閉授權視窗'))
          }
        }
      }, 500)
    }

    try {
      microsoftTeams.authentication.authenticate({
        url: authUrl,
        width: 600,
        height: 535,
        successCallback: (token: string) => {
          console.log('取得 Graph Token 成功')
          const cachedToken = cacheGraphToken(token)
          graphAuthPromise = null
          resolve(cachedToken)
        },
        failureCallback: (reason: string) => {
          console.warn('Teams authenticate 失敗，改用 window.open:', reason)
          try {
            openAuthWindow()
          } catch (error) {
            graphAuthPromise = null
            reject(error instanceof Error ? error : new Error(String(error)))
          }
        }
      })
    } catch (error) {
      console.warn('Teams authenticate 拋例外，改用 window.open:', error)
      try {
        openAuthWindow()
      } catch (fallbackError) {
        graphAuthPromise = null
        reject(fallbackError instanceof Error ? fallbackError : new Error(String(fallbackError)))
      }
    }
  })

  return graphAuthPromise
}

async function fetchGraphProfile(token: string): Promise<GraphProfile> {
  const response = await fetch('https://graph.microsoft.com/v1.0/me', {
    headers: {
      Authorization: `Bearer ${token}`
    }
  })

  if (!response.ok) {
    const text = await response.text()
    console.error('Graph API 錯誤:', response.status, text)
    if (response.status === 401 || response.status === 403) {
      clearGraphTokenCache()
    }
    throw new Error(`Graph API 錯誤：${response.status}`)
  }

  const data = (await response.json()) as GraphProfile
  console.log('Graph 使用者資訊:', data)
  return data
}

function normalizeErrorMessage(error: unknown): string {
  if (typeof error === 'string') {
    return translateAuthError(error)
  }

  if (error instanceof Error) {
    return translateAuthError(error.message)
  }

  return '未知錯誤，請重新整理並重試。'
}

function translateAuthError(message: string): string {
  if (!message) {
    return '登入流程中斷，請重新嘗試。'
  }

  if (message.includes('interaction_in_progress') || message.includes('redirect_pending')) {
    return '授權流程尚未完成，請稍後再試或關閉視窗後重新開啟應用程式。'
  }

  if (message.includes('CancelledByUser')) {
    return '您取消了授權。若要完成登入，請允許 Teams 對話框中的授權要求。'
  }

  if (message.includes('popup_window_error')) {
    return '瀏覽器阻擋了授權視窗。請允許 Teams / Microsoft 的彈出視窗，或改用 Teams 桌面版。'
  }

  if (message.includes('not_supported')) {
    return '目前的 Teams 用戶端不支援此登入流程，請更新 Teams 或改用網頁版。'
  }

  return message
}

const doubleSurnames = [
  '歐陽', '司馬', '端木', '上官', '夏侯', '諸葛', '尉遲', '皇甫',
  '澹台', '公孫', '仲孫', '軒轅', '令狐', '鍾離', '宇文', '長孫',
  '慕容', '司徒', '司空', '司寇', '申屠', '南宮', '東方', '西門'
]

const GRAPH_TOKEN_CACHE_KEY = 'teams-graph-token'
const GRAPH_TOKEN_EXPIRES_AT_KEY = 'teams-graph-token-exp'
const GRAPH_TOKEN_PAYLOAD_KEY = 'teams-graph-token-payload'

function splitChineseName(name?: string) {
  if (!name) {
    return { surname: undefined, givenName: undefined }
  }

  const trimmed = name.trim()
  if (!trimmed) {
    return { surname: undefined, givenName: undefined }
  }

  if (/\s+/.test(trimmed)) {
    const parts = trimmed.split(/\s+/)
    const surname = parts.shift()
    return {
      surname,
      givenName: parts.length > 0 ? parts.join(' ') : undefined
    }
  }

  if (/^[\x00-\x7F]+$/.test(trimmed)) {
    return { surname: trimmed, givenName: undefined }
  }

  const possibleDouble = trimmed.slice(0, 2)
  if (doubleSurnames.includes(possibleDouble) && trimmed.length > 2) {
    return { surname: possibleDouble, givenName: trimmed.slice(2) }
  }

  if (trimmed.length >= 2) {
    return { surname: trimmed.slice(0, 1), givenName: trimmed.slice(1) }
  }

  return { surname: trimmed, givenName: undefined }
}

function getCachedGraphToken(): CachedGraphToken | null {
  try {
    const token = sessionStorage.getItem(GRAPH_TOKEN_CACHE_KEY)
    const expiresAtRaw = sessionStorage.getItem(GRAPH_TOKEN_EXPIRES_AT_KEY)
    const payloadRaw = sessionStorage.getItem(GRAPH_TOKEN_PAYLOAD_KEY)
    if (!token || !expiresAtRaw) {
      return null
    }

    const expiresAt = Number(expiresAtRaw)
    if (Number.isNaN(expiresAt) || Date.now() >= expiresAt - 60_000) {
      clearGraphTokenCache()
      return null
    }

    const payload = payloadRaw ? (JSON.parse(payloadRaw) as TokenPayload) : undefined
    return { token, payload }
  } catch (error) {
    console.warn('讀取快取 Token 失敗', error)
    return null
  }
}

function cacheGraphToken(token: string): CachedGraphToken {
  try {
    const payload = decodeJwt(token)
    const expiresAt = payload?.exp ? payload.exp * 1000 : Date.now() + 50 * 60 * 1000
    sessionStorage.setItem(GRAPH_TOKEN_CACHE_KEY, token)
    sessionStorage.setItem(GRAPH_TOKEN_EXPIRES_AT_KEY, String(expiresAt))
    if (payload) {
      const { exp, ...rest } = payload
      sessionStorage.setItem(GRAPH_TOKEN_PAYLOAD_KEY, JSON.stringify(rest))
    }
    return { token, payload }
  } catch (error) {
    console.warn('快取 Graph Token 失敗', error)
    return { token }
  }
}

function clearGraphTokenCache() {
  sessionStorage.removeItem(GRAPH_TOKEN_CACHE_KEY)
  sessionStorage.removeItem(GRAPH_TOKEN_EXPIRES_AT_KEY)
  sessionStorage.removeItem(GRAPH_TOKEN_PAYLOAD_KEY)
}

function decodeJwt(token: string): TokenPayload {
  const parts = token.split('.')
  if (parts.length < 2) {
    return {}
  }

  const payload = parts[1]
  const normalized = payload.replace(/-/g, '+').replace(/_/g, '/')
  try {
    const decoded = atob(normalized)
    return JSON.parse(decoded)
  } catch (error) {
    console.warn('解析 Token 失敗', error)
    return {}
  }
}

function renderUserInfoRows(user: UserInfo) {
  const chineseName = user.chineseSurname || user.chineseGivenName
    ? `${user.chineseSurname ?? ''}${user.chineseGivenName ?? ''}`
    : undefined
  const englishName =
    user.givenName || user.surname ? `${user.givenName ?? ''} ${user.surname ?? ''}`.trim() : undefined

  const rows = [
    { label: '顯示名稱', value: user.displayName },
    { label: '中文姓名', value: chineseName },
    { label: '英文姓名', value: englishName },
    { label: '帳號 (UPN)', value: user.userPrincipalName },
    { label: 'Email', value: user.email },
    { label: '職稱', value: user.jobTitle },
    { label: '部門', value: user.department },
    { label: '辦公地點', value: user.officeLocation },
    { label: '手機', value: user.mobilePhone },
    {
      label: '公司電話',
      value: user.businessPhones && user.businessPhones.length > 0 ? user.businessPhones.join('、') : undefined
    },
    { label: '偏好語言', value: user.preferredLanguage },
    { label: '使用者 ID', value: user.id },
    { label: '租戶 ID', value: user.tenantId }
  ]

  return rows.filter(item => item.value)
}

function getEnvironmentMessage(env: HostEnvironment) {
  switch (env) {
    case 'teams-desktop':
      return {
        title: '偵測到 Teams 桌面版',
        message: '使用 Teams 內建授權視窗，過程中不需要輸入帳密。'
      }
    case 'teams-mobile':
      return {
        title: '偵測到 Teams 行動版',
        message: '使用 Teams 行動版內建授權視窗。'
      }
    case 'teams-web':
      return {
        title: '偵測到 Teams 網頁版',
        message: '授權流程將在同一頁面或新分頁進行。若是 Safari 可使用下方手動授權按鈕或改用桌面版。'
      }
    case 'standalone':
      return {
        title: '偵測到瀏覽器模式',
        message: '使用 MSAL loginRedirect 流程完成登入，可直接在同一視窗登入。'
      }
    default:
      return null
  }
}

function isSafari(): boolean {
  const ua = navigator.userAgent
  return ua.includes('Safari') && !ua.includes('Chrome') && !ua.includes('Chromium')
}

interface CachedGraphToken {
  token: string
  payload?: TokenPayload
}

