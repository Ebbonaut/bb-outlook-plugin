# BB Outlook Plugin – BlockBrain AI Draft

Outlook Add-in, das per Klick eine KI-generierte E-Mail-Antwort über die BlockBrain API erstellt.

## 🏗 Projekt-Struktur

```
bb-outlook-plugin/
├── manifest.xml              ← Outlook Add-in Manifest
├── src/
│   └── taskpane/
│       ├── taskpane.html     ← Taskpane UI
│       ├── taskpane.css      ← Styling
│       └── taskpane.js       ← Logik (API-Aufrufe)
├── assets/                   ← Icons (16/32/64/80/128px)
├── .env                      ← API-Konfiguration (nicht committen!)
├── webpack.config.js         ← Build-Konfiguration
└── package.json
```

## 🚀 Setup

### 1. Abhängigkeiten installieren

```bash
npm install
```

### 2. API-Token konfigurieren

Bearbeite die `.env` Datei:

```env
BB_BASE_URL=https://blocky.qa.theblockbrain.io
BB_BEARER_TOKEN=DEIN_ECHTER_TOKEN
BB_GENERAL_BOT_ID=69b7e7e7d54d83a12f86a13b
```

### 3. Dev-Server starten

```bash
npm run dev
```

Der Server startet auf `https://localhost:3000` mit HTTPS (erforderlich für Outlook).

### 4. Add-in in Outlook laden (Sideloading)

#### Option A: Outlook im Web (am einfachsten)

1. Öffne [Outlook Web](https://outlook.office.com)
2. Öffne eine beliebige E-Mail
3. Klicke auf **⋯** (Weitere Aktionen) → **Add-ins abrufen**
4. Klicke auf **Meine Add-Ins** → **Benutzerdefinierte Add-ins hinzufügen** → **Aus Datei hinzufügen**
5. Wähle die `manifest.xml` aus diesem Projekt

#### Option B: Outlook Desktop (Windows)

1. Öffne Outlook
2. Gehe zu **Datei** → **Add-Ins verwalten** (oder **Get Add-Ins**)
3. Klicke auf **Meine Add-Ins** → **Benutzerdefiniertes Add-In hinzufügen** → **Aus Datei hinzufügen**
4. Wähle die `manifest.xml`

#### Option C: Microsoft 365 Admin Center

1. Gehe zu [admin.microsoft.com](https://admin.microsoft.com)
2. **Einstellungen** → **Integrierte Apps** → **Benutzerdefinierte Apps hochladen**
3. Lade die `manifest.xml` hoch

### 5. Verwenden

1. Öffne eine E-Mail in Outlook
2. Klicke auf das **BB Draft** Icon im Header/Ribbon
3. Klicke **Antwort generieren** → Die KI erstellt einen Entwurf
4. Klicke **Entwurf öffnen** → Antwort wird in "Allen antworten" eingefügt

## 🔧 Befehle

| Befehl | Beschreibung |
|--------|-------------|
| `npm run dev` | Dev-Server mit Hot Reload starten |
| `npm run build` | Produktions-Build erstellen |
| `npm run validate` | Manifest validieren |

## 📦 Produktion / Deployment

Für den produktiven Einsatz:

1. Ersetze `https://localhost:3000` in der `manifest.xml` durch deine gehostete URL
2. Erstelle echte PNG-Icons (ersetze die Platzhalter in `assets/`)
3. Ändere die Add-in ID in der `manifest.xml` (neue GUID generieren)
4. Baue das Projekt: `npm run build`
5. Deploye den `dist/` Ordner auf einen Webserver mit HTTPS
6. Lade die aktualisierte `manifest.xml` im Microsoft 365 Admin Center hoch

## ⚠️ Hinweise

- **HTTPS erforderlich**: Outlook Add-ins funktionieren nur über HTTPS
- **Token-Sicherheit**: Der Bearer Token steht aktuell im Frontend-Code. Für Produktion sollte ein Backend-Proxy verwendet werden
- **Icons**: Die generierten Icons sind SVG-Platzhalter. Für Produktion echte PNGs erstellen
