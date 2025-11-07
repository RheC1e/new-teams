import { useEffect, useState } from 'react'
import * as microsoftTeams from '@microsoft/teams-js'
import './App.css'

interface UserInfo {
  displayName?: string
  email?: string
  userPrincipalName?: string
  id?: string
}

function App() {
  const [status, setStatus] = useState<'loading' | 'success' | 'error'>('loading')
  const [userInfo, setUserInfo] = useState<UserInfo | null>(null)
  const [errorMessage, setErrorMessage] = useState<string>('')

  useEffect(() => {
    const init = async () => {
      try {
        await microsoftTeams.app.initialize()
        console.log('Teams SDK 初始化成功')

        const context = await microsoftTeams.app.getContext()
        console.log('Teams context:', context)

        const teamsUser = context.user as {
          displayName?: string
          email?: string
          userPrincipalName?: string
          id?: string
          aadObjectId?: string
        } | undefined

        const user = {
          displayName: teamsUser?.displayName,
          email: teamsUser?.email ?? teamsUser?.userPrincipalName,
          userPrincipalName: teamsUser?.userPrincipalName,
          id: teamsUser?.id ?? teamsUser?.aadObjectId
        }

        setUserInfo(user)
        setStatus('success')

        // 嘗試取得 SSO Token（不成功也不影響顯示）
        try {
          const token = await microsoftTeams.authentication.getAuthToken({
            silent: true
          })
          console.log('取得 Token 成功:', token.substring(0, 20) + '...')
        } catch (tokenError) {
          console.warn('取得 Token 失敗，但已取得 Teams Context', tokenError)
        }
      } catch (error) {
        console.error('自動登入流程失敗', error)
        const message = error instanceof Error ? error.message : '未知錯誤'
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
            <h1>正在登入...</h1>
            <p>正在使用您的 Microsoft 365 帳號自動登入</p>
          </div>
        )}

        {status === 'success' && (
          <div className="status-card success">
            <div className="success-icon">✓</div>
            <h1>登入成功！</h1>
            {userInfo && (
              <div className="user-info">
                <p><strong>顯示名稱：</strong>{userInfo.displayName || '未提供'}</p>
                <p><strong>帳號：</strong>{userInfo.userPrincipalName || '未提供'}</p>
                <p><strong>使用者 ID：</strong>{userInfo.id || '未提供'}</p>
              </div>
            )}
          </div>
        )}

        {status === 'error' && (
          <div className="status-card error">
            <div className="error-icon">✗</div>
            <h1>登入失敗</h1>
            <p className="error-message">{errorMessage || '發生未知錯誤'}</p>
            <p className="error-hint">請確認您已正確登入 Microsoft Teams</p>
          </div>
        )}
      </div>
    </div>
  )
}

export default App

