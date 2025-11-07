## RHEMA Teams 自動登入專案技術總結

### 1. 以往專案失敗的原因（來源：`以往teams失敗自動登入/`）

**pwa-demo（請款 PWA）**
- 採純 MSAL.js `loginRedirect/loginPopup` 流程，適合一般瀏覽器，但在 Teams 桌面版的 WebView 中會被攔截或需要互動，無法做到自動登入。
- 沒有整合 `@microsoft/teams-js`，無法存取 Teams 提供的 SSO context；結果是登入流程完全脫離 Teams，導致需要額外授權視窗。

**rhema-pwa-demo（第二版請款 PWA）**
- 架構仍以單頁 PWA 為核心，登入邏輯同樣依賴 MSAL redirect，缺少 Teams manifest 與 `webApplicationInfo` 設定。
- 重新導向 URI 與 Teams 內嵌 iframe 不相容，造成 Teams 桌面版中無法完成登入或卡在 redirect loop。

**teams-sso-test（SSO 測試專案）**
- 已啟用 Teams SDK 與 manifest，但流程嘗試同時取得「應用程式 Token」與 Microsoft Graph Token。
- `getAuthToken()` 取得的是 **應用程式自己的 Token**（針對 `api://...` 資源），沒有 Graph 權限，直接呼叫 `https://graph.microsoft.com/v1.0/me` 會回傳 401。
- 為補救而引入 MSAL（`ssoSilent` / `acquireTokenSilent`），流程複雜且仍需互動授權；桌面版快取舊腳本時就容易再次回到 401 錯誤。

### 2. 本次成功專案的核心做法（`main` 分支現況）

**技術重點**
- 以 Vite + React + TypeScript 建置，使用 `@microsoft/teams-js@2.x`。
- 啟動後立即呼叫 Teams 認證流程，在桌面版中會跳出內嵌授權視窗。
- 授權頁 (`public/auth.html`) 先嘗試 `ssoSilent`，若需要互動則改走 `loginRedirect`（無彈窗、避免被阻擋）。
- 取得 Microsoft Graph Access Token 後，呼叫 `/me` 取得顯示名稱、中文姓/名、帳號與使用者 ID。

**程式流程（`src/App.tsx`）**
1. `await microsoftTeams.app.initialize()`。
2. `await microsoftTeams.app.getContext()` 取得 Teams 使用者資訊與租戶 ID。
3. 呼叫 `microsoftTeams.authentication.authenticate()` 開啟 `auth.html`，將 `loginHint` 帶入授權頁。
4. 授權頁先嘗試 `ssoSilent` 取得 Token；若失敗則重導至 Microsoft 登入頁完成授權，回到同一頁後透過 `handleRedirectPromise()` 提取 Token，最後 `notifySuccess` 回主頁。
5. 主頁使用 `fetch('https://graph.microsoft.com/v1.0/me')` 取得真實帳號資訊並顯示。

- **必要設定**
  - `manifest.json`
    - `id` 與 `webApplicationInfo.id`：`33abd69a-d012-498a-bddb-8608cbf10c2d`
    - `webApplicationInfo.resource`：`api://new-teams-potp.vercel.app/33abd69a-d012-498a-bddb-8608cbf10c2d`
    - `contentUrl` / `websiteUrl` / `validDomains`：`https://new-teams-potp.vercel.app`
  - Azure Entra ID
    - 應用程式註冊同上 Client ID。
    - SPA Redirect URI 新增：`https://new-teams-potp.vercel.app`
    - SPA Redirect URI 新增：`https://new-teams-potp.vercel.app/auth.html`
    - Application ID URI：`api://new-teams-potp.vercel.app/33abd69a-d012-498a-bddb-8608cbf10c2d`
    - 定義 scope：`access_as_user`（完整值同 Application ID URI + `/access_as_user`）
    - 授權客戶端：Teams 桌面 `1fec8e78-bce4-4aaf-ab1b-5451cc387264`

**部署與測試流程**
1. **推送程式碼**：`git push origin main`
2. **Vercel 自動部署**：確認 Deployments 最新狀態為 Ready，必要時 `Redeploy`。
3. **Teams 套件**：
   - 執行 `zip -r teams-autologin-package.zip manifest.json icon-color.png icon-outline.png`
   - 在 Teams 管理中心上傳 ZIP。
4. **Azure Entra 設定**：確認兩個 SPA Redirect URI 都存在（根網址與 `/auth.html`）。
5. **驗證**：Teams 桌面版按 `⌘ + R` 重新整理；應跳出授權對話框，授權後畫面顯示 Graph 回傳的名稱、帳號與 ID。

**快取刷新建議**
- 若 Teams 未抓到新版，可在 `contentUrl` 加查詢參數（例：`?build=c074614`）。
- Mac 清快取：刪除 `~/Library/Containers/com.microsoft.teams2.*` 與 `~/Library/Group Containers/UBF8T346G9.com.microsoft.teams` 後重開 Teams（僅需在重大更新時進行）。

### 3. 未來專案複製手冊

1. **複製此專案**（或以 `create-vite` 建立 React/TS，再套用 `src/App.tsx` 架構）。
2. **調整 Vercel 網域**：
   - 更新 `manifest.json` 的 `contentUrl / validDomains / webApplicationInfo.resource`。
   - 於 Azure Entra ID 同步更新 Redirect URI 與 Application ID URI。
3. **部署與上傳**：照上節部署流程操作。
4. **驗證邏輯**：沿用 `auth.html` + `loginRedirect` 流程即可；只要確保 Azure 的 Redirect URI 正確，Teams 桌面版便能自動帶入帳號並完成授權。
5. **擴充需求**：若要調用更多 Graph API，可在 Azure Entra ID 增加對應 scope，授權頁會一併完成同意；若需調用自家 API，可在後端驗證此次取得的 Graph Token 或實作 On-Behalf-Of 流程。

### 4. 快速指令摘要

```bash
# 開發
npm install
npm run dev

# 建置
npm run build

# 產生 Teams 套件 ZIP
zip -r teams-autologin-package.zip manifest.json icon-color.png icon-outline.png

# Git 推送
git add .
git commit -m "更新內容"
git push
```

---

只要依照以上設定與流程，即可複製此次成功的 Teams 自動登入體驗；若需擴充至其他功能，可在此基礎上延伸。提交此文件即可讓未來的專案快速對焦於成功配置。

