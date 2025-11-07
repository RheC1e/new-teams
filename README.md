# Teams 自動登入測試應用程式

這是一個簡單的 Microsoft Teams 桌面版應用程式，用於測試 Microsoft 365 自動登入功能。

## 功能

- ✅ 自動使用當前使用者的 Microsoft 365 帳號登入
- ✅ 顯示登入狀態（載入中、成功、失敗）
- ✅ 顯示使用者資訊（名稱、帳號、ID）

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
2. **Azure AD 應用程式註冊** - 更新重新導向 URI

## 注意事項

- 此應用程式需要在 Teams 桌面版中運行
- 需要正確配置 Azure AD 應用程式註冊
- 需要設定正確的 API 權限

