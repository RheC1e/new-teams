import { useEffect, useState } from 'react'
import * as microsoftTeams from '@microsoft/teams-js'
import './App.css'

type AuthStatus = 'loading' | 'waitingConsent' | 'success' | 'error'

interface UserInfo {
  displayName?: string
  email?: string
  userPrincipalName?: string
  id?: string
  tenantId?: string
  chineseSurname?: string
  chineseGivenName?: string
}

function App() {
  const [status, setStatus] = useState<AuthStatus>('loading')
  const [userInfo, setUserInfo] = useState<UserInfo | null>(null)
  const [errorMessage, setErrorMessage] = useState<string>('')

  useEffect(() => {
    const init = async () => {
      try {
        setStatus('loading')
        await microsoftTeams.app.initialize()
        console.log('Teams SDK 初始化成功')

        const context = await microsoftTeams.app.getContext()
        console.log('Teams context:', context)

        const teamsUser = context.user as TeamsUser | undefined
        const tenantId = context.user?.tenant?.id

        setStatus('waitingConsent')
        const loginHint = teamsUser?.userPrincipalName || teamsUser?.email
        const graphToken = await requestGraphToken(loginHint)
        const graphProfile = await fetchGraphProfile(graphToken)

        const user = deriveUserInfo({
          teamsUser,
          graphProfile,
          tenantId
        })

        setUserInfo(user)
        setStatus('success')
      } catch (error) {
        console.error('登入流程失敗', error)
        const message = normalizeErrorMessage(error)
        setErrorMessage(message)
        setStatus('error')
      }
    }

    void init()
  }, [])

  return (
    <div className="app">
      <div className="container">
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

        {status === 'success' && userInfo && (
          <div className="status-card success">
            <div className="success-icon">✓</div>
            <h1>登入成功！</h1>
            <div className="user-info">
              <p><strong>顯示名稱：</strong>{userInfo.displayName || '未提供'}</p>
              <p><strong>中文姓名：</strong>
                {userInfo.chineseSurname || userInfo.chineseGivenName
                  ? `${userInfo.chineseSurname ?? ''}${userInfo.chineseGivenName ?? ''}`
                  : '未提供'}
              </p>
              <p><strong>帳號：</strong>{userInfo.userPrincipalName || userInfo.email || '未提供'}</p>
              <p><strong>使用者 ID：</strong>{userInfo.id || '未提供'}</p>
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
}

interface DeriveParams {
  teamsUser?: TeamsUser
  graphProfile?: GraphProfile
  tenantId?: string
}

function deriveUserInfo(params: DeriveParams): UserInfo {
  const { teamsUser, graphProfile, tenantId } = params

  const displayName = graphProfile?.displayName ?? teamsUser?.displayName
  const email = graphProfile?.mail ?? teamsUser?.email ?? teamsUser?.userPrincipalName
  const userPrincipalName = graphProfile?.userPrincipalName ?? teamsUser?.userPrincipalName
  const id = graphProfile?.id ?? teamsUser?.id ?? teamsUser?.aadObjectId

  const preferSurname = graphProfile?.surname
  const preferGivenName = graphProfile?.givenName

  const fallback = splitChineseName(displayName)

  return {
    displayName,
    email,
    userPrincipalName,
    id,
    tenantId,
    chineseSurname: preferSurname ?? fallback.surname,
    chineseGivenName: preferGivenName ?? fallback.givenName
  }
}

let graphAuthPromise: Promise<string> | null = null

async function requestGraphToken(loginHint?: string): Promise<string> {
  const cached = getCachedGraphToken()
  if (cached) {
    return cached
  }

  if (graphAuthPromise) {
    return graphAuthPromise
  }

  graphAuthPromise = new Promise((resolve, reject) => {
    const hintParam = loginHint ? `?loginHint=${encodeURIComponent(loginHint)}` : ''
    microsoftTeams.authentication.authenticate({
      url: `${window.location.origin}/auth.html${hintParam}`,
      width: 600,
      height: 535,
      successCallback: (token: string) => {
        console.log('取得 Graph Token 成功')
        cacheGraphToken(token)
        graphAuthPromise = null
        resolve(token)
      },
      failureCallback: (reason: string) => {
        console.error('Graph 授權失敗:', reason)
        clearGraphTokenCache()
        graphAuthPromise = null
        reject(new Error(reason || 'Graph 授權失敗'))
      }
    })
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

  if (message.includes('interaction_in_progress')) {
    return '授權流程尚未完成，請稍後再試或關閉視窗後重新開啟應用程式。'
  }

  if (message.includes('CancelledByUser')) {
    return '您取消了授權。若要完成登入，請允許 Teams 對話框中的授權要求。'
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

function getCachedGraphToken(): string | null {
  try {
    const token = sessionStorage.getItem(GRAPH_TOKEN_CACHE_KEY)
    const expiresAtRaw = sessionStorage.getItem(GRAPH_TOKEN_EXPIRES_AT_KEY)
    if (!token || !expiresAtRaw) {
      return null
    }

    const expiresAt = Number(expiresAtRaw)
    if (Number.isNaN(expiresAt) || Date.now() >= expiresAt - 60_000) {
      clearGraphTokenCache()
      return null
    }

    return token
  } catch (error) {
    console.warn('讀取快取 Token 失敗', error)
    return null
  }
}

function cacheGraphToken(token: string) {
  try {
    const payload = decodeJwt(token)
    if (!payload?.exp) {
      return
    }
    const expiresAt = payload.exp * 1000
    sessionStorage.setItem(GRAPH_TOKEN_CACHE_KEY, token)
    sessionStorage.setItem(GRAPH_TOKEN_EXPIRES_AT_KEY, String(expiresAt))
  } catch (error) {
    console.warn('快取 Graph Token 失敗', error)
  }
}

function clearGraphTokenCache() {
  sessionStorage.removeItem(GRAPH_TOKEN_CACHE_KEY)
  sessionStorage.removeItem(GRAPH_TOKEN_EXPIRES_AT_KEY)
}

function decodeJwt(token: string): { exp?: number } {
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

