# Teams 自動登入測試應用程式

這是一個簡單的 Microsoft Teams 桌面版應用程式，用於測試 Microsoft 365 自動登入功能。

## 功能

- ✅ 自動觸發 Teams 認證流程並於必要時顯示授權對話框
- ✅ 透過 Microsoft Graph 取得真實使用者資訊（顯示名稱、帳號、ID、中文姓名）
- ✅ 顯示登入狀態（建立連線、等待授權、成功、失敗）

## 技術棧

- **React 18** - UI 框架
- **TypeScript** - 類型安全
- **Vite** - 建置工具
- **@microsoft/teams-js** - Teams SDK
- **Microsoft Graph API** - 取得使用者資訊

## 開發

### 安裝依賴

```bash
npm install
```

### 本地開發

```bash
npm run dev
```

### 建置

```bash
npm run build
```

## 部署

部署到 Vercel 後，需要更新以下設定：

1. **manifest.json** - 更新 `contentUrl`、`websiteUrl`、`validDomains` 和 `webApplicationInfo`
2. **Azure AD 應用程式註冊**
   - SPA Redirect URI 新增：`https://{your-domain}.vercel.app`
   - SPA Redirect URI 新增：`https://{your-domain}.vercel.app/auth.html`
   - Application ID URI 與公開 API 與 Teams manifest 保持一致

## 注意事項

- 此應用程式需要在 Teams 桌面版中運行
- 需要正確配置 Azure AD 應用程式註冊
- 需要設定正確的 API 權限

