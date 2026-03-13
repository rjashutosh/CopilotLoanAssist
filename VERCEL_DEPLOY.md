# Deploy to Vercel

## 1. Deploy from GitHub

1. Go to **[vercel.com](https://vercel.com)** and sign in (use **Continue with GitHub**).
2. Click **Add New…** → **Project**.
3. **Import** the repo **`rjashutosh/CopilotLoanAssist`** (or your fork).
4. Leave **Root Directory** empty (so the repo root is used).
5. Click **Deploy**. Wait for the build to finish.
6. Note your deployment URL, e.g. **`https://copilot-loan-assist.vercel.app`**.

## 2. App URLs

After deployment, use:

- **Login:** `https://<your-project>.vercel.app/loan-chat-demo/login.html`
- **Chat (after sign-in):** `https://<your-project>.vercel.app/loan-chat-demo/index.html`

## 3. Azure redirect URI

1. Open **Microsoft Entra** → **App registrations** → your app → **Authentication**.
2. Under **Single-page application**, add:
   - **`https://<your-project>.vercel.app/loan-chat-demo/index.html`**
   (Replace `<your-project>` with your actual Vercel project name, e.g. `copilot-loan-assist`.)
3. Save.

## 4. Update config and redeploy

1. In **`loan-chat-demo/config.js`**, set:
   ```javascript
   redirectUri: 'https://<your-project>.vercel.app/loan-chat-demo/index.html',
   ```
2. Commit and push to `main`. Vercel will redeploy automatically.

## 5. Test

Open **`https://<your-project>.vercel.app/loan-chat-demo/login.html`**, click **Sign in with Microsoft**, complete sign-in, and confirm you land on the chat page.
