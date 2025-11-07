import { useEffect, useState } from 'react'
import * as microsoftTeams from '@microsoft/teams-js'
import './App.css'

interface UserInfo {
  displayName?: string
  userPrincipalName?: string
  id?: string
}

function App() {
  const [status, setStatus] = useState<'loading' | 'success' | 'error'>('loading')
  const [userInfo, setUserInfo] = useState<UserInfo | null>(null)
  const [errorMessage, setErrorMessage] = useState<string>('')

  useEffect(() => {
    // 初始化 Teams SDK
    microsoftTeams.app.initialize().then(() => {
      console.log('Teams SDK 初始化成功')
      
      // 取得使用者資訊（自動登入）
      microsoftTeams.authentication.getAuthToken({
        resources: [],
        silent: true
      }).then((token: string) => {
        console.log('取得 Token 成功:', token.substring(0, 20) + '...')
        
        // 使用 Token 取得使用者資訊
        fetch('https://graph.microsoft.com/v1.0/me', {
          headers: {
            'Authorization': `Bearer ${token}`
          }
        })
        .then(response => {
          if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`)
          }
          return response.json()
        })
        .then((data: UserInfo) => {
          console.log('使用者資訊:', data)
          setUserInfo(data)
          setStatus('success')
        })
        .catch((error: Error) => {
          console.error('取得使用者資訊失敗:', error)
          setErrorMessage(error.message)
          setStatus('error')
        })
      }).catch((error: Error) => {
        console.error('取得 Token 失敗:', error)
        setErrorMessage(error.message)
        setStatus('error')
      })
    }).catch((error: Error) => {
      console.error('Teams SDK 初始化失敗:', error)
      setErrorMessage(error.message)
      setStatus('error')
    })
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

