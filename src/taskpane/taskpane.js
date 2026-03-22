import "./taskpane.css";

/* global Office */

// --- DEFAULTS (fallback if nothing in localStorage) ---
var DEFAULTS = {
  bbUrl: "https://blocky.theblockbrain.ai",
  bbToken: "",
  bbBotId: "",
  useSystemPrompt: false,
};

// --- LOAD SETTINGS from localStorage ---
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

// --- OFFICE INITIALIZATION ---
Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    initUI();
  }
});

function initUI() {
  // Populate settings fields
  var settings = loadSettings();
  document.getElementById("cfg-url").value = settings.bbUrl;
  document.getElementById("cfg-token").value = settings.bbToken;
  document.getElementById("cfg-bot").value = settings.bbBotId;
  document.getElementById("cfg-system-prompt").checked = settings.useSystemPrompt;

  // If token or bot ID empty → auto-open settings
  if (!settings.bbToken || !settings.bbBotId) {
    document.getElementById("settings-details").open = true;
  }

  // Event-Listener
  document.getElementById("btn-generate").addEventListener("click", handleGenerate);
  document.getElementById("btn-save").addEventListener("click", handleSave);
  document.getElementById("btn-copy").addEventListener("click", handleCopy);
  document.getElementById("btn-open-reply").addEventListener("click", handleOpenReply);
}

// --- SAVE SETTINGS ---
function handleSave() {
  var url = document.getElementById("cfg-url").value.trim();
  var token = document.getElementById("cfg-token").value.trim();
  var botId = document.getElementById("cfg-bot").value.trim();
  var useSystemPrompt = document.getElementById("cfg-system-prompt").checked;

  saveSettings(url, token, botId, useSystemPrompt);

  // Briefly show "Saved" status
  var statusEl = document.getElementById("save-status");
  statusEl.classList.remove("hidden");
  setTimeout(function () {
    statusEl.classList.add("hidden");
  }, 2000);
}

// --- Stored draft for reply ---
var lastDraftHtml = "";

// --- GENERATE DRAFT ---
async function handleGenerate() {
  var settings = loadSettings();

  // Validation
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

  // Hide result section on new generation
  resultEl.classList.add("hidden");
  lastDraftHtml = "";

  // UI: loading state
  btnGenerate.disabled = true;
  btnIcon.innerText = "⏳";
  btnText.innerText = "Generating...";
  statusEl.classList.remove("hidden");
  spinnerEl.className = "spinner";
  statusText.innerText = "Reading email...";

  try {
    var item = Office.context.mailbox.item;

    // 1. Read mail body
    var bodyText = await getMailBody(item);
    statusText.innerText = "Bot is generating reply...";

    // 2. Build prompt
    var prompt = "";

    if (settings.useSystemPrompt) {
      prompt +=
        "Analyze the following email thread. Identify the MOST RECENT incoming message " +
        "as well as its sender and context (e.g. private, business, boss, colleague, customer, partner, friend).\n\n" +
        "YOUR ROLE: You are the recipient of the most recent message. Compose an appropriate reply on their behalf.\n\n" +
        "OUTPUT RULES:\n" +
        "- Output ONLY the plain reply text – no introduction, no Markdown, no metadata.\n" +
        "- The reply MUST follow this structure:\n" +
        "  1. Greeting (e.g. 'Dear Mr/Ms ...', 'Hi ...', 'Hello ...' – matching the context and tone of the email)\n" +
        "  2. Body of the reply (address all points raised)\n" +
        "  3. Closing with name (e.g. 'Best regards', 'Kind regards', 'Cheers' – matching the context)\n" +
        "- Tone and style adapt automatically to the context.\n" +
        "- Language: use the language of the most recent incoming email.\n\n";
    }

    if (hints) {
        prompt += "ADDITIONAL HINTS FROM USER:\n" + hints + "\n\n";
    }

    prompt +=
        "EMAIL THREAD:\n" +
        "Subject: " + item.subject + "\n" +
        "Content: " + bodyText;

    // 3. Create new conversation
    statusText.innerText = "Connecting to BLOCKBRAIN...";
    var convoData = await createConvo(settings.bbUrl, settings.bbToken, settings.bbBotId, item.subject);
    var convoId = convoData.body.dataRoomId;

    // 4. Send prompt
    statusText.innerText = "Bot is writing reply...";
    var answer = await sendUserInput(settings.bbUrl, settings.bbToken, convoId, prompt);

    if (answer && answer.body && answer.body.content) {
      var draftContent = answer.body.content;
      lastDraftHtml = draftContent.replace(/\n/g, "<br>");

      // 5. Show draft in sidebar
      draftTextEl.value = draftContent;
      resultEl.classList.remove("hidden");

      spinnerEl.className = "spinner done";
      statusText.innerText = "✅ Reply generated!";
    } else {
      spinnerEl.className = "spinner done";
      statusText.innerText = "❌ No response received from bot.";
    }
  } catch (error) {
    console.error("BlockBrain Error:", error);
    spinnerEl.className = "spinner done";
    statusText.innerText = "❌ Error: " + error.message;
  }

  // Reset UI
  btnGenerate.disabled = false;
  btnIcon.innerText = "✨";
  btnText.innerText = "Generate reply";
}

// --- OPEN REPLY (with retry logic) ---
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
            // Fallback: old API
            try { item.displayReplyAllForm(lastDraftHtml); } catch (e) {
              console.warn("displayReplyAllForm fallback failed:", e);
            }
          }
        }
      );
    } else {
      // Fallback: old API
      item.displayReplyAllForm(lastDraftHtml);
    }
  } catch (e) {
    console.warn("Reply open failed:", e);
    // Draft is visible in sidebar → user can copy manually
  }
}

// --- MANUALLY OPEN REPLY ---
function handleOpenReply() {
  tryOpenReply();
}

// --- COPY TO CLIPBOARD ---
async function handleCopy() {
  var draftTextEl = document.getElementById("draft-text");
  var copyTextEl = document.getElementById("copy-text");
  try {
    await navigator.clipboard.writeText(draftTextEl.value);
    copyTextEl.innerText = "✅ Copied!";
  } catch (e) {
    // Fallback: select + execCommand
    draftTextEl.select();
    document.execCommand("copy");
    copyTextEl.innerText = "✅ Copied!";
  }
  setTimeout(function () {
    copyTextEl.innerText = "Copy";
  }, 2000);
}

// --- HELPER FUNCTIONS ---

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
