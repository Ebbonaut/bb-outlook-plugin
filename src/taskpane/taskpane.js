import "./taskpane.css";

/* global Office */

// --- KONFIGURATION (aus .env oder hier direkt) ---
const BB_BASE_URL = process.env.BB_BASE_URL || "https://blocky.qa.theblockbrain.io";
const BB_BEARER_TOKEN = process.env.BB_BEARER_TOKEN || "MYTOKEN";
const BB_GENERAL_BOT_ID = process.env.BB_GENERAL_BOT_ID || "69b7e7e7d54d83a12f86a13b";

// Globale Variable für zwischengespeicherte KI-Antwort
let cachedAiContent = "";

// --- OFFICE INITIALISIERUNG ---
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("btn-generate").addEventListener("click", handleGenerate);
    document.getElementById("btn-open-draft").addEventListener("click", handleOpenDraft);
  }
});

// --- ANTWORT GENERIEREN ---
async function handleGenerate() {
  const item = Office.context.mailbox.item;
  const btnGenerate = document.getElementById("btn-generate");
  const btnOpenDraft = document.getElementById("btn-open-draft");
  const statusSection = document.getElementById("status-section");
  const statusIcon = document.getElementById("status-icon");
  const statusText = document.getElementById("status-text");

  // UI: Lade-Zustand
  btnGenerate.classList.add("ms-hidden");
  btnOpenDraft.classList.add("ms-hidden");
  statusSection.classList.remove("ms-hidden");
  statusIcon.className = "status-icon";
  statusText.innerText = "KI generiert Antwort...";

  try {
    const bodyText = await getMailBody(item);

    const prompt = `Analysiere den folgenden E-Mail-Verlauf. Identifiziere die ZULETZT eingegangene Nachricht und wer der Absender war.
DEINE ROLLE: Du bist der Empfänger dieser letzten Nachricht. Antworte in seinem Namen professionell und passend auf den Inhalt.

AUSGABE-REGELN:
- NUR den reinen Antworttext ausgeben.
- Kein Markdown, keine Einleitung (wie "Hier ist die Antwort:"), keine Metadaten.
- Sprache: Deutsch (oder die Sprache der Mail).
- Falls die letzte Nachricht vom Kundenservice kam, antworte als Kunde. Falls sie von einem Kunden kam, antworte als Service.

MAIL-VERLAUF:
Betreff: ${item.subject}
Inhalt: ${bodyText}`;

    // 1. Neue Konversation erstellen
    const convoData = await createConvo(BB_GENERAL_BOT_ID);
    const convoId = convoData.body.dataRoomId;

    // 2. Prompt senden & Antwort erhalten
    const answer = await sendUserInput(convoId, prompt);

    if (answer && answer.body && answer.body.content) {
      cachedAiContent = answer.body.content;

      // UI: Erfolg
      statusIcon.className = "status-icon success";
      statusText.innerText = "Antwort bereit!";
      btnOpenDraft.classList.remove("ms-hidden");
    } else {
      showError("Keine Antwort von der KI erhalten.");
    }
  } catch (error) {
    console.error("BlockBrain API Error:", error);
    showError("Fehler bei der API-Kommunikation.");
  }

  function showError(msg) {
    statusIcon.className = "status-icon error";
    statusText.innerText = msg;
    setTimeout(() => {
      statusSection.classList.add("ms-hidden");
      btnGenerate.classList.remove("ms-hidden");
    }, 3000);
  }
}

// --- ENTWURF ÖFFNEN ---
function handleOpenDraft() {
  if (!cachedAiContent) return;

  const item = Office.context.mailbox.item;
  const htmlPayload = cachedAiContent.replace(/\n/g, "<br>");

  item.displayReplyAllForm(htmlPayload);

  // UI zurücksetzen
  cachedAiContent = "";
  document.getElementById("btn-open-draft").classList.add("ms-hidden");
  document.getElementById("btn-generate").classList.remove("ms-hidden");
  document.getElementById("status-section").classList.add("ms-hidden");
}

// --- HILFSFUNKTIONEN ---

/**
 * Liest den Mail-Body als Text (Promise-Wrapper).
 */
function getMailBody(item) {
  return new Promise((resolve, reject) => {
    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error("Fehler beim Lesen der Mail: " + result.error.message));
      }
    });
  });
}

/**
 * Erstellt eine neue Konversation bei BlockBrain.
 */
async function createConvo(botId) {
  const formattedDate = new Date().toISOString().replace("T", " ").split(".")[0];
  const response = await fetch(`${BB_BASE_URL}/cortex/active-bot/${botId}/convo`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Accept: "application/json",
      Authorization: `Bearer ${BB_BEARER_TOKEN}`,
    },
    body: JSON.stringify({ convoName: formattedDate }),
  });

  if (!response.ok) {
    throw new Error(`Convo-Erstellung fehlgeschlagen: ${response.status}`);
  }

  return await response.json();
}

/**
 * Sendet den Prompt an BlockBrain und erhält die KI-Antwort.
 */
async function sendUserInput(convoId, content) {
  const response = await fetch(`${BB_BASE_URL}/cortex/completions/v2/user-input`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Accept: "application/json",
      Authorization: `Bearer ${BB_BEARER_TOKEN}`,
    },
    body: JSON.stringify({
      convoId,
      content,
      sessionId: "c1b3c8e3-beb1-4c34-aa24-454aabf12709",
      messageType: "user-question",
      enableStreaming: false,
    }),
  });

  if (!response.ok) {
    throw new Error(`API-Anfrage fehlgeschlagen: ${response.status}`);
  }

  return await response.json();
}
