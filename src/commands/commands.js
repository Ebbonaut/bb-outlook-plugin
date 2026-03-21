/* global Office */

// --- KONFIGURATION ---
var BB_BASE_URL = process.env.BB_BASE_URL || "https://blocky.qa.theblockbrain.io";
var BB_BEARER_TOKEN = process.env.BB_BEARER_TOKEN || "MYTOKEN";
var BB_GENERAL_BOT_ID = process.env.BB_GENERAL_BOT_ID || "69b7e7e7d54d83a12f86a13b";

/**
 * WICHTIG: Funktion SOFORT global registrieren – BEVOR Office.onReady() läuft.
 * Wenn das erst danach passiert, findet Office die Funktion beim ersten Klick nicht.
 */
function generateDraft(event) {
  // Alles in eine async IIFE wrappen, damit event.completed() IMMER aufgerufen wird
  (async function () {
    try {
      var item = Office.context.mailbox.item;

      // 1. Mail-Body lesen
      var bodyText = await getMailBody(item);

      // 2. Prompt bauen
      var prompt =
        "Analysiere den folgenden E-Mail-Verlauf. Identifiziere die ZULETZT eingegangene Nachricht und wer der Absender war.\n" +
        "DEINE ROLLE: Du bist der Empfänger dieser letzten Nachricht. Antworte in seinem Namen professionell und passend auf den Inhalt.\n\n" +
        "AUSGABE-REGELN:\n" +
        "- NUR den reinen Antworttext ausgeben.\n" +
        '- Kein Markdown, keine Einleitung (wie "Hier ist die Antwort:"), keine Metadaten.\n' +
        "- Sprache: Deutsch (oder die Sprache der Mail).\n" +
        "- Falls die letzte Nachricht vom Kundenservice kam, antworte als Kunde. Falls sie von einem Kunden kam, antworte als Service.\n\n" +
        "MAIL-VERLAUF:\n" +
        "Betreff: " + item.subject + "\n" +
        "Inhalt: " + bodyText;

      // 3. Neue Konversation erstellen
      var convoData = await createConvo(BB_GENERAL_BOT_ID);
      var convoId = convoData.body.dataRoomId;

      // 4. Prompt senden & Antwort erhalten
      var answer = await sendUserInput(convoId, prompt);

      if (answer && answer.body && answer.body.content) {
        // 5. Reply-All-Fenster mit KI-Antwort öffnen
        var htmlPayload = answer.body.content.replace(/\n/g, "<br>");
        item.displayReplyAllForm(htmlPayload);
      }
    } catch (error) {
      console.error("BlockBrain API Error:", error);
    }

    // Delay damit displayReplyAllForm Zeit hat, das Fenster zu öffnen,
    // bevor Office die Runtime beendet
    await new Promise(function (r) { setTimeout(r, 2000); });
    event.completed();
  })();
}

// SOFORT global setzen – noch vor Office.onReady
window.generateDraft = generateDraft;

// Office initialisieren (muss aufgerufen werden, aber Funktion ist schon registriert)
Office.onReady(function () {
  // Runtime ist bereit
});

// --- HILFSFUNKTIONEN ---

function getMailBody(item) {
  return new Promise(function (resolve, reject) {
    item.body.getAsync(Office.CoercionType.Text, function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error("Fehler beim Lesen der Mail: " + result.error.message));
      }
    });
  });
}

async function createConvo(botId) {
  var formattedDate = new Date().toISOString().replace("T", " ").split(".")[0];
  var response = await fetch(BB_BASE_URL + "/cortex/active-bot/" + botId + "/convo", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Accept: "application/json",
      Authorization: "Bearer " + BB_BEARER_TOKEN,
    },
    body: JSON.stringify({ convoName: formattedDate }),
  });

  if (!response.ok) {
    throw new Error("Convo-Erstellung fehlgeschlagen: " + response.status);
  }

  return await response.json();
}

async function sendUserInput(convoId, content) {
  var response = await fetch(BB_BASE_URL + "/cortex/completions/v2/user-input", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Accept: "application/json",
      Authorization: "Bearer " + BB_BEARER_TOKEN,
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
    throw new Error("API-Anfrage fehlgeschlagen: " + response.status);
  }

  return await response.json();
}

// Funktion GLOBAL registrieren – das ist der zuverlässigste Weg für ExecuteFunction.
// Office sucht per FunctionName auf dem globalen window-Objekt.
window.generateDraft = generateDraft;
