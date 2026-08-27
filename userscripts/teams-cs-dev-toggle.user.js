// ==UserScript==
// @name         Teams CS dev toggle
// @match        https://teams.microsoft.com/v2/*
// @match        https://teams.cloud.microsoft/v2/*
// @match        https://local.teams.office.com/v2/*
// @run-at       document-idle
// ==/UserScript==

(() => {
  if (window.top !== window) return;

  globalThis.__teamsCsDevToggle?.destroy();

  const TOGGLE_ID = "teams-cs-dev-toggle";
  const SETTINGS_KEY = "tmp.settings";
  const OVERRIDE_URL = "https://api.conv-dev.skype.net/conv/";
  document.getElementById(TOGGLE_ID)?.remove();

  function readSettings() {
    return JSON.parse(localStorage.getItem(SETTINGS_KEY) ?? "{}");
  }

  function isEnabled(settings) {
    return Object.hasOwn(
      settings.calling ?? {},
      "conversationServiceUrlOverride"
    );
  }

  async function waitFor(getValue, timeoutMs = 5000) {
    const deadline = Date.now() + timeoutMs;
    while (Date.now() < deadline) {
      const value = getValue();
      if (value) return value;
      await new Promise(resolve => setTimeout(resolve, 50));
    }
    throw new Error("Timed out waiting for restart control");
  }

  async function restartTeams() {
    const restartSelector =
      "button[aria-label='Restart (also exits container)']";
    let restartButton = document.querySelector(restartSelector);
    if (!restartButton) {
      const ringElement = document.querySelector(
        "[data-tid='titlebar-status-indicator']"
      );
      if (!ringElement) throw new Error("Could not find the ring control");
      ringElement.click();
      restartButton = await waitFor(() =>
        document.querySelector(restartSelector)
      );
    }
    restartButton.click();
  }

  async function updateSettings(enabled) {
    const settings = readSettings();
    if (enabled) {
      settings.calling ??= {};
      settings.calling.conversationServiceUrlOverride = {
        url: OVERRIDE_URL,
      };
    } else {
      delete settings.calling?.conversationServiceUrlOverride;
    }
    localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings));
    await restartTeams();
  }

  function createToggle(enabled) {
    const container = document.createElement("div");
    container.id = TOGGLE_ID;
    Object.assign(container.style, {
      alignItems: "center",
      color: "var(--colorNeutralForeground1, #242424)",
      cursor: "pointer",
      display: "inline-flex",
      font: "600 11px/20px -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif",
      gap: "5px",
      height: "20px",
      marginInlineStart: "8px",
      whiteSpace: "nowrap",
    });

    const label = document.createElement("span");
    label.textContent = "cs-dev";

    const button = document.createElement("button");
    button.type = "button";
    button.setAttribute("role", "switch");
    button.setAttribute("aria-label", "cs-dev");
    button.setAttribute("aria-checked", String(enabled));
    button.title = "Conversation Service development override";
    Object.assign(button.style, {
      alignItems: "center",
      background: enabled
        ? "var(--colorBrandBackground, #6264a7)"
        : "var(--colorNeutralBackground1, #fff)",
      border: enabled
        ? "1px solid var(--colorBrandBackground, #6264a7)"
        : "1px solid var(--colorNeutralStroke1, #616161)",
      borderRadius: "10px",
      cursor: "pointer",
      display: "inline-flex",
      height: "20px",
      justifyContent: enabled ? "flex-end" : "flex-start",
      padding: "2px",
      width: "36px",
    });

    const thumb = document.createElement("span");
    Object.assign(thumb.style, {
      background: enabled
        ? "var(--colorNeutralForegroundOnBrand, #fff)"
        : "var(--colorNeutralForeground3, #616161)",
      borderRadius: "50%",
      display: "block",
      height: "14px",
      width: "14px",
    });
    button.append(thumb);
    button.addEventListener("click", async event => {
      event.stopPropagation();
      button.disabled = true;
      try {
        await updateSettings(!isEnabled(readSettings()));
      } catch (error) {
        button.disabled = false;
        console.error("[Teams CS dev toggle]", error);
      }
    });
    container.append(label, button);
    return container;
  }

  function installToggle() {
    if (document.getElementById(TOGGLE_ID)) return;
    const ringBadge = document.querySelector(
      "[data-tid='titlebar-status-indicator']"
    );
    if (!ringBadge?.parentElement) return;
    const enabled = isEnabled(readSettings());
    ringBadge.parentElement.after(createToggle(enabled));
  }

  installToggle();
  const observer = new MutationObserver(installToggle);
  observer.observe(document.documentElement, { childList: true, subtree: true });
  globalThis.__teamsCsDevToggle = {
    destroy() {
      observer.disconnect();
      document.getElementById(TOGGLE_ID)?.remove();
      delete globalThis.__teamsCsDevToggle;
    },
  };
})();
