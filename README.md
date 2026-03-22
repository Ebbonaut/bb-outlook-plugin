# BLOCKBRAIN Outlook Plugin

Outlook Add-in that generates AI-powered email replies via the BlockBrain API at the click of a button.

## 🏗 Project Structure

```
bb-outlook-plugin/
├── manifest.xml              ← Outlook Add-in manifest
├── src/
│   └── taskpane/
│       ├── taskpane.html     ← Taskpane UI
│       ├── taskpane.css      ← Styling
│       └── taskpane.js       ← Logic (API calls)
├── assets/                   ← Icons (16/32/64/80/128px) & logos
├── webpack.config.js         ← Build configuration
└── package.json
```

## 🚀 Setup

### 1. Install dependencies

```bash
npm install
```

### 2. Start dev server

```bash
npm run dev
```

The server starts on `https://localhost:3000` with HTTPS (required by Outlook).

### 3. Configure the plugin

API URL, Bearer Token, and Bot ID are configured directly in the plugin's **Settings** panel — no `.env` file needed. Settings are persisted in `localStorage`.

### 4. Load add-in in Outlook (Sideloading)

#### Option A: Outlook on the Web (easiest)

1. Open [Outlook Web](https://outlook.office.com)
2. Open any email
3. Click **⋯** (More actions) → **Get Add-ins**
4. Click **My add-ins** → **Add a custom add-in** → **Add from file**
5. Select the `manifest.xml` from this project

#### Option B: Outlook Desktop (Windows)

1. Open Outlook
2. Go to **File** → **Manage Add-ins** (or **Get Add-ins**)
3. Click **My add-ins** → **Add a custom add-in** → **Add from file**
4. Select the `manifest.xml`

#### Option C: Microsoft 365 Admin Center

1. Go to [admin.microsoft.com](https://admin.microsoft.com)
2. **Settings** → **Integrated Apps** → **Upload Custom Apps**
3. Upload the `manifest.xml`

### 5. Usage

1. Open an email in Outlook
2. Click the **BLOCKBRAIN** icon in the header/ribbon
3. Optionally enter additional hints (e.g. "Reply friendly but firm")
4. Click **Generate reply** → The AI creates a draft
5. Copy the draft or click **Open reply** to insert it into a Reply All form

## 🔧 Commands

| Command | Description |
|---------|-------------|
| `npm run dev` | Start dev server with hot reload |
| `npm run build` | Create production build |

## 📦 Production / Deployment

This project is deployed automatically via **GitHub Actions** to **GitHub Pages** on every push to `main`.

Live URL: `https://ebbonaut.github.io/bb-outlook-plugin/`

For a custom deployment:

1. Replace all URLs in `manifest.xml` with your hosted URL
2. Run `npm run build`
3. Deploy the `dist/` folder to a web server with HTTPS
4. Upload the updated `manifest.xml` in the Microsoft 365 Admin Center

## ⚠️ Notes

- **HTTPS required**: Outlook Add-ins only work over HTTPS
- **Token security**: The Bearer Token is stored in `localStorage` on the client. For production, consider using a backend proxy
- **System prompt**: Can be toggled on/off in Settings. Disable it if the bot already has its own Initial Instructions configured
