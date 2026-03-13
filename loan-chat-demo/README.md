# Loan Support Assistant — Web App with Microsoft Authentication

A simple web application that requires users to sign in with Microsoft (Azure AD) before they can use the embedded Copilot Studio chatbot. Built with HTML, CSS, JavaScript, and MSAL.js.

## Features

- **Microsoft sign-in** using Microsoft Identity Platform (Azure AD) via MSAL.js
- **Protected chat page**: only authenticated users can access the chatbot
- **Session handling**: sign-in state in localStorage (shared with popup); redirect to login when not authenticated
- **Token for Copilot**: the same Microsoft token is sent to the Copilot Studio agent (Conversations API) so the agent runs in authenticated mode
- **Mobile-friendly** layout and touch-friendly controls

## Project Structure

```
loan-chat-demo/
├── index.html      # Chat page (protected; redirects to login if not signed in)
├── login.html      # Sign-in page (opens popup)
├── auth-popup.html # Popup: does Microsoft sign-in, then redirects opener to index
├── auth-popup.js   # Popup logic (MSAL in popup)
├── config.js       # Copilot API base URL + MSAL settings (edit before run)
├── auth.js         # MSAL handleRedirect, signOut (cache in localStorage)
├── script.js       # Login button (popup), chat protection, programmatic chat
├── styles.css      # Shared styles
└── README.md       # This file
```

## 1. Azure App Registration (Microsoft Entra ID)

You need a **Single-page application (SPA)** registration so the web app can sign users in and get a token for the Power Platform / Copilot API.

### Steps

1. Go to [Microsoft Entra ID (Azure Portal)](https://entra.microsoft.com/) → **App registrations** → **New registration**.

2. **Name**: e.g. `Loan Support Web App`.

3. **Supported account types**: choose **Accounts in any organizational directory and personal Microsoft accounts** (or **Single tenant** if you only want your org).

4. **Redirect URIs** (add all that you use):
   - Platform: **Single-page application (SPA)**.
   - Add **both**:
     - Chat page: `http://localhost:8080/loan-chat-demo/index.html` (local) and/or `https://yourdomain.com/loan-chat-demo/index.html` (production).
     - **Popup sign-in** (recommended): `http://localhost:8080/loan-chat-demo/auth-popup.html` and/or `https://yourdomain.com/loan-chat-demo/auth-popup.html`.  
   The app uses a **popup** for sign-in by default so you land back on the same site (localhost or Vercel).

5. Click **Register**.

6. On the app’s **Overview** page, copy the **Application (client) ID**. You will put this in `config.js` as `MSAL.clientId`.

7. **Optional (if your Copilot/API requires it):**  
   **Certificates & secrets** → **Client secrets** is **not** used for the SPA (MSAL uses the implicit/public flow). Leave it empty for the front-end.

8. **API permissions** (required for Conversations API):
   - **Add a permission** → **APIs my organization uses** → search for **Power Platform** (or App ID `8578e004-a5c6-46e7-913e-12f58912df43`).
   - Add **Delegated** permissions. Some agents need:
     - **CopilotStudio.Copilots.Invoke** — expand **CopilotStudio**, check this.
     - **All.All.ReadWrite** — if the API returns "InsufficientDelegatedPermissions" asking for this: expand **All** in the same Power Platform API permission list (scroll down), then check **All.ReadWrite**. If you don't see an "All" section, the permission may not be available in your tenant; ask your Power Platform admin or use an agent that only requires CopilotStudio.Copilots.Invoke.
   - Click **Add permissions**, then **Grant admin consent for &lt;your org&gt;** (if your tenant uses it).

## 2. Configuring the Web App (`config.js`)

Edit `config.js` and set:

| Setting | Description |
|--------|--------------|
| `CONVERSATIONS_API_BASE` | Base URL of your Copilot Studio connection string (no `/conversations`, no `?api-version=...`). Get it from Copilot Studio → Publish → Channels → connection string. |
| `MSAL.clientId` | Application (client) ID from the Azure app registration. |
| `MSAL.redirectUri` | **Exact** redirect URI you configured in Azure (must be **index.html**, i.e. the chat page). Example: `http://localhost:8080/loan-chat-demo/index.html`. |
| `MSAL.authority` | Usually `https://login.microsoftonline.com/common`. Use your tenant ID for single-tenant: `https://login.microsoftonline.com/<tenant-id>`. |
| `MSAL.scope` | Usually `https://api.powerplatform.com/.default` for Power Platform / Copilot. |

Example:

```javascript
var LOAN_CHAT_CONFIG = {
  CONVERSATIONS_API_BASE: 'https://YOUR_ENV.environment.api.powerplatform.com/.../bots/YOUR_BOT',
  API_VERSION: '2022-03-01-preview',
  MSAL: {
    clientId: 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx',
    authority: 'https://login.microsoftonline.com/common',
    redirectUri: 'http://localhost:8080/loan-chat-demo/index.html',
    scope: 'https://api.powerplatform.com/.default'
  }
};
```

## 3. Copilot Studio Configuration

The agent must be set up so that it **requires Microsoft authentication** and is callable via the **Conversations API** with a Bearer token.

1. **Agent authentication**
   - In Copilot Studio: **Settings** → **Security** → **Authentication**.
   - Set to **Authenticate with Microsoft** (or **Authenticate manually** with Microsoft Entra ID).
   - This ensures the agent expects a Microsoft token when called from the web app.

2. **Connection string (Conversations API)**
   - **Publish** → **Channels** → open the channel you use for the web (e.g. Custom website or the one that gives a connection string).
   - Copy the **connection string** or the Conversations API URL.
   - The base URL is the part before `/conversations?api-version=...`. Put that in `config.js` as `CONVERSATIONS_API_BASE`.

3. **Same tenant / app**
   - The Azure app registration used by the web app (MSAL) should be in the same tenant as the Copilot Studio environment, or configured as allowed for that environment, so the token is accepted by the Power Platform / Copilot API.

## 4. How to Run the Application

### Prerequisites

- A local web server (the app must be served over HTTP/HTTPS; `file://` will not work with redirects and cookies/storage).
- `config.js` updated with your Azure client ID, redirect URI, and Copilot API base URL.

### Option A: VS Code Live Server

1. Open the `loan-chat-demo` folder in VS Code.
2. Install “Live Server” if needed.
3. Right-click `index.html` or `login.html` → **Open with Live Server**.
4. Note the URL (e.g. `http://127.0.0.1:5500/loan-chat-demo/login.html`).
5. In Azure, set the SPA redirect URI to your **chat page** (e.g. `http://127.0.0.1:5500/loan-chat-demo/index.html`).
6. In `config.js`, set `MSAL.redirectUri` to that same URL.

### Option B: Node.js (http-server)

```bash
cd loan-chat-demo
npx http-server -p 8080 -c-1
```

- Open `http://localhost:8080/loan-chat-demo/login.html`.
- Redirect URI in Azure and in `config.js`: `http://localhost:8080/loan-chat-demo/index.html`.

### Option C: Python

```bash
cd loan-chat-demo
# Python 3
python -m http.server 8080
```

- Open `http://localhost:8080/login.html` (path may vary by how you serve the folder).
- Set Azure and `config.js` redirect URI to match the full URL of `index.html`.

### Authentication Flow (Step-by-Step)

1. User opens the app (e.g. `.../loan-chat-demo/index.html` or `.../loan-chat-demo/login.html`).
2. If they open **index.html** and are **not** signed in, they are redirected to **login.html**.
3. On **login.html**, user clicks **Sign in with Microsoft**.
4. They are redirected to Microsoft login; after success, they are sent back to **index.html** (your configured redirect URI).
5. **index.html** runs `handleRedirectPromise()`, stores the account in MSAL cache (session storage), and shows the chat UI.
6. When the user sends a message, the app gets an access token via MSAL and calls the Copilot Studio Conversations API with `Authorization: Bearer <token>`.
7. The agent responds; messages are shown in the same page.
8. **Sign out** clears the session and sends the user to **login.html** (or Microsoft logout, depending on `auth.js`).

## 5. Making the Authentication Flow Work Correctly

- **Redirect URI must match exactly** (including path and trailing `index.html` if used). No extra query string unless you added it in Azure.
- Use **HTTPS** in production; some browsers restrict auth on plain HTTP.
- If you use **localhost**, use the same host in Azure and in `config.js` (e.g. `http://localhost:8080/...`).
- If the agent returns 401, the token may be for the wrong audience or scope; confirm `MSAL.scope` and that the Copilot channel is set to “Authenticate with Microsoft” (or equivalent).
- For “AADSTS50011: Reply URL mismatch”, fix the SPA redirect URI in Azure to exactly match `MSAL.redirectUri` in `config.js`.

## 6. Mobile and Production

- The UI is responsive; test on a real device or emulator.
- For production, deploy the folder to a web server, set the Azure redirect URI to your production `index.html` URL, and set `MSAL.redirectUri` in `config.js` to that same URL.

## Summary

| Item | Where |
|------|--------|
| Azure app registration | Microsoft Entra ID → App registrations → SPA redirect URI = `index.html` URL |
| Client ID & redirect URI | `config.js` → `MSAL.clientId`, `MSAL.redirectUri` |
| Copilot API base URL | `config.js` → `CONVERSATIONS_API_BASE` |
| Agent auth | Copilot Studio → Settings → Security → Authenticate with Microsoft |
| Run | Serve `loan-chat-demo` over HTTP/HTTPS and open `login.html` or `index.html` |
