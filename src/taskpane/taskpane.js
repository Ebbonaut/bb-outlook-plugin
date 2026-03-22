import "./taskpane.css";

/* global Office */

// --- DEFAULTS (Fallback wenn nichts in localStorage) ---
var DEFAULTS = {
  bbUrl: "https://blocky.theblockbrain.ai",
  bbToken: "",
  bbBotId: "",
  useSystemPrompt: false,
};

// --- SETTINGS aus localStorage laden ---
function loadSettings() {
  try {
    var saved = localStorage.getItem("bb_settings");
    if (saved) {
      var parsed = JSON.parse(saved);
      return {
        bbUrl: parsed.bbUrl || DEFAULTS.bbUrl,
        bbToken: parsed.bbToken || DEFAULTS.bbToken,
        bbBotId: parsed.bbBotId || DEFAULTS.bbBotId,
        useSystemPrompt: parsed.useSystemPrompt !== undefined ? parsed.useSystemPrompt : DEFAULTS.useSystemPrompt,
      };
    }
  } catch (e) {
    // ignore
  }
  return {
    bbUrl: DEFAULTS.bbUrl,
    bbToken: DEFAULTS.bbToken,
    bbBotId: DEFAULTS.bbBotId,
    useSystemPrompt: DEFAULTS.useSystemPrompt,
  };
}

function saveSettings(url, token, botId, useSystemPrompt) {
  var settings = {
    bbUrl: url,
    bbToken: token,
    bbBotId: botId,
    useSystemPrompt: useSystemPrompt,
  };
  localStorage.setItem("bb_settings", JSON.stringify(settings));
}

// --- OFFICE INITIALISIERUNG ---
Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    initUI();
  }
});

function initUI() {
  // Settings-Felder befüllen
  var settings = loadSettings();
  document.getElementById("cfg-url").value = settings.bbUrl;
  document.getElementById("cfg-token").value = settings.bbToken;
  document.getElementById("cfg-bot").value = settings.bbBotId;
  document.getElementById("cfg-system-prompt").checked = settings.useSystemPrompt;

  // Wenn Token oder Bot ID leer → Settings automatisch öffnen
  if (!settings.bbToken || !settings.bbBotId) {
    document.getElementById("settings-details").open = true;
  }

  // Event-Listener
  document.getElementById("btn-generate").addEventListener("click", handleGenerate);
  document.getElementById("btn-save").addEventListener("click", handleSave);
  document.getElementById("btn-copy").addEventListener("click", handleCopy);
  document.getElementById("btn-open-reply").addEventListener("click", handleOpenReply);
}

// --- SETTINGS SPEICHERN ---
function handleSave() {
  var url = document.getElementById("cfg-url").value.trim();
  var token = document.getElementById("cfg-token").value.trim();
  var botId = document.getElementById("cfg-bot").value.trim();
  var useSystemPrompt = document.getElementById("cfg-system-prompt").checked;

  saveSettings(url, token, botId, useSystemPrompt);

  // Kurz "Gespeichert" anzeigen
  var statusEl = document.getElementById("save-status");
  statusEl.classList.remove("hidden");
  setTimeout(function () {
    statusEl.classList.add("hidden");
  }, 2000);
}

// --- Gespeicherter Draft für Reply ---
var lastDraftHtml = "";

// --- DRAFT GENERIEREN ---
async function handleGenerate() {
  var settings = loadSettings();

  // Validierung
  if (!settings.bbToken || !settings.bbBotId) {
    document.getElementById("settings-details").open = true;
    if (!settings.bbToken) {
      document.getElementById("cfg-token").focus();
    } else {
      document.getElementById("cfg-bot").focus();
    }
    return;
  }

  var btnGenerate = document.getElementById("btn-generate");
  var btnText = document.getElementById("btn-text");
  var btnIcon = document.getElementById("btn-icon");
  var statusEl = document.getElementById("status");
  var statusText = document.getElementById("status-text");
  var spinnerEl = document.getElementById("spinner");
  var resultEl = document.getElementById("result");
  var draftTextEl = document.getElementById("draft-text");
  var hints = document.getElementById("hints").value.trim();

  // Ergebnis-Bereich verstecken bei neuer Generierung
  resultEl.classList.add("hidden");
  lastDraftHtml = "";

  // UI: Lade-Zustand
  btnGenerate.disabled = true;
  btnIcon.innerText = "⏳";
  btnText.innerText = "Wird generiert...";
  statusEl.classList.remove("hidden");
  spinnerEl.className = "spinner";
  statusText.innerText = "Lese E-Mail...";

  try {
    var item = Office.context.mailbox.item;

    // 1. Mail-Body lesen
    var bodyText = await getMailBody(item);
    statusText.innerText = "Bot generiert Antwort...";

    // 2. Prompt bauen
    var prompt = "";

    if (settings.useSystemPrompt) {
      prompt +=
        "Analysiere den folgenden E-Mail-Verlauf. Identifiziere die ZULETZT eingegangene Nachricht " +
        "sowie deren Absender und Kontext (z.B. privat, geschäftlich, Chef, Kollege, Kunde, Partner, Freund).\n\n" +
        "DEINE ROLLE: Du bist der Empfänger der letzten Nachricht. Verfasse eine passende Antwort in seinem Namen.\n\n" +
        "AUSGABE-REGELN:\n" +
        "- NUR den reinen Antworttext ausgeben – keine Einleitung, kein Markdown, keine Metadaten.\n" +
        "- Die Antwort MUSS folgende Struktur haben:\n" +
        "  1. Anrede (z.B. 'Sehr geehrter Herr/Frau ...', 'Lieber ...', 'Hallo ...', 'Hi ...' – passend zum Kontext und Tonalität der E-Mail)\n" +
        "  2. Inhalt der Antwort (auf alle angesprochenen Punkte eingehen)\n" +
        "  3. Grußformel mit Name (z.B. 'Mit freundlichen Grüßen', 'Beste Grüße', 'Viele Grüße' – passend zum Kontext)\n" +
        "- Ton und Stil passen sich automatisch dem Kontext an.\n" +
        "- Sprache: die Sprache der zuletzt eingegangenen Mail.\n\n";
    }

    if (hints) {
        prompt += "ZUSÄTZLICHE HINWEISE VOM BENUTZER:\n" + hints + "\n\n";
    }

    prompt +=
        "MAIL-VERLAUF:\n" +
        "Betreff: " + item.subject + "\n" +
        "Inhalt: " + bodyText;

    // 3. Neue Konversation erstellen
    statusText.innerText = "Verbinde mit BLOCKBRAIN...";
    var convoData = await createConvo(settings.bbUrl, settings.bbToken, settings.bbBotId, item.subject);
    var convoId = convoData.body.dataRoomId;

    // 4. Prompt senden
    statusText.innerText = "Bot schreibt Antwort...";
    var answer = await sendUserInput(settings.bbUrl, settings.bbToken, convoId, prompt);

    if (answer && answer.body && answer.body.content) {
      var draftContent = answer.body.content;
      lastDraftHtml = draftContent.replace(/\n/g, "<br>");

      // 5. Draft im Sidebar anzeigen
      draftTextEl.value = draftContent;
      resultEl.classList.remove("hidden");

      spinnerEl.className = "spinner done";
      statusText.innerText = "✅ Antwort generiert!";
    } else {
      spinnerEl.className = "spinner done";
      statusText.innerText = "❌ Keine Antwort vom Bot erhalten.";
    }
  } catch (error) {
    console.error("BlockBrain Error:", error);
    spinnerEl.className = "spinner done";
    statusText.innerText = "❌ Fehler: " + error.message;
  }

  // UI zurücksetzen
  btnGenerate.disabled = false;
  btnIcon.innerText = "✨";
  btnText.innerText = "Antwort generieren";
}

// --- REPLY ÖFFNEN (mit Retry-Logik) ---
function tryOpenReply() {
  if (!lastDraftHtml) return;

  var item = Office.context.mailbox.item;
  try {
    // displayReplyAllFormAsync (neuere API) mit Callback
    if (typeof item.displayReplyAllFormAsync === "function") {
      item.displayReplyAllFormAsync(
        { htmlBody: lastDraftHtml },
        function (asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.warn("displayReplyAllFormAsync failed:", asyncResult.error);
            // Fallback: Altes API
            try { item.displayReplyAllForm(lastDraftHtml); } catch (e) {
              console.warn("displayReplyAllForm fallback failed:", e);
            }
          }
        }
      );
    } else {
      // Fallback: Altes API
      item.displayReplyAllForm(lastDraftHtml);
    }
  } catch (e) {
    console.warn("Reply open failed:", e);
    // Draft ist im Sidebar sichtbar → User kann manuell kopieren
  }
}

// --- MANUELL REPLY ÖFFNEN ---
function handleOpenReply() {
  tryOpenReply();
}

// --- IN ZWISCHENABLAGE KOPIEREN ---
async function handleCopy() {
  var draftTextEl = document.getElementById("draft-text");
  var copyTextEl = document.getElementById("copy-text");
  try {
    await navigator.clipboard.writeText(draftTextEl.value);
    copyTextEl.innerText = "✅ Kopiert!";
  } catch (e) {
    // Fallback: select + execCommand
    draftTextEl.select();
    document.execCommand("copy");
    copyTextEl.innerText = "✅ Kopiert!";
  }
  setTimeout(function () {
    copyTextEl.innerText = "Kopieren";
  }, 2000);
}

// --- HILFSFUNKTIONEN ---

function getMailBody(item) {
  return new Promise(function (resolve, reject) {
    item.body.getAsync(Office.CoercionType.Text, function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error(result.error.message));
      }
    });
  });
}

async function createConvo(baseUrl, token, botId, convoName = "Outlook Plugin Conversation") {
  var response = await fetch(baseUrl + "/cortex/active-bot/" + botId + "/convo", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Accept: "application/json",
      Authorization: "Bearer " + token,
    },
    body: JSON.stringify({ convoName }),
  });

  if (!response.ok) {
    throw new Error("API " + response.status);
  }
  return await response.json();
}

async function sendUserInput(baseUrl, token, convoId, content) {
  var response = await fetch(baseUrl + "/cortex/completions/v2/user-input", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Accept: "application/json",
      Authorization: "Bearer " + token,
    },
    body: JSON.stringify({
      convoId: convoId,
      content: content,
      sessionId: "c1b3c8e3-beb1-4c34-aa24-454aabf12709",
      messageType: "user-question",
      enableStreaming: false,
    }),
  });

  if (!response.ok) {
    throw new Error("API " + response.status);
  }
  return await response.json();
}
