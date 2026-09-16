// ==UserScript==
// @name         Teams Focus mode
// @version      1.0.3
// @description  Reduce visual emphasis for unread messages and notifications.
// @match        https://teams.microsoft.com/v2/*
// @match        https://teams.cloud.microsoft/v2/*
// @match        https://local.teams.office.com/v2/*
// @run-at       document-idle
// ==/UserScript==

(() => {
  if (window.top !== window) return;

  globalThis.__teamsFocusMode?.destroy();

  const TOGGLE_ID = "teams-focus-mode-toggle";
  const STYLE_ID = "teams-focus-mode-style";
  const ROOT_ATTRIBUTE = "data-teams-focus-mode";
  const SETTINGS_KEY = "teamsmonkey.focus-mode.enabled";

  document.getElementById(TOGGLE_ID)?.remove();
  document.getElementById(STYLE_ID)?.remove();

  const style = document.createElement("style");
  style.id = STYLE_ID;
  style.textContent = `
    html[${ROOT_ATTRIBUTE}] [data-testid="hamburger-button-badge"],
    html[${ROOT_ATTRIBUTE}] [data-testid^="badge-count-indicator"],
    html[${ROOT_ATTRIBUTE}] [data-testid="dot-badge-container"],
    html[${ROOT_ATTRIBUTE}] [data-testid*="badge"],
    html[${ROOT_ATTRIBUTE}] [data-id^="conversation-folder-header-"],
    html[${ROOT_ATTRIBUTE}] [class*="Badge"] {
      display: none !important;
    }

    #${TOGGLE_ID} button:focus,
    #${TOGGLE_ID} button:focus-visible {
      border: 0 !important;
      box-shadow: none !important;
      outline: none !important;
    }

    html[${ROOT_ATTRIBUTE}] [data-testid="simple-collab-left-rail-sticky-filter-toggle-button-TEAMS_AND_CHANNELS"],
    html[${ROOT_ATTRIBUTE}] [data-testid="simple-collab-left-rail-sticky-filter-toggle-button-TEAMS_AND_CHANNELS"] *,
    html[${ROOT_ATTRIBUTE}] [data-id="conversation-folder-header-TEAMS_AND_CHANNELS"],
    html[${ROOT_ATTRIBUTE}] [data-id="conversation-folder-header-TEAMS_AND_CHANNELS"] *,
    html[${ROOT_ATTRIBUTE}] [data-testid="conversation-folder-header"],
    html[${ROOT_ATTRIBUTE}] [data-testid="conversation-folder-header"] *,
    html[${ROOT_ATTRIBUTE}] [data-id^="conversation-folder-header-"] *,
    html[${ROOT_ATTRIBUTE}] [role="treeitem"] [role="text"],
    html[${ROOT_ATTRIBUTE}] [data-testid="rail-jumper-toolbar-primary-button"] [aria-label="Channels"],
    html[${ROOT_ATTRIBUTE}] [data-testid="rail-jumper-toolbar-primary-button"][aria-label="Channels"] *,
    html[${ROOT_ATTRIBUTE}] [class*="RailJumperToolbarPrimaryButton"] {
      font-weight: 400 !important;
    }

    html[${ROOT_ATTRIBUTE}] [data-testid^="simple-collab-left-rail-sticky-filter-toggle-button-"] [data-testid],
    html[${ROOT_ATTRIBUTE}] [data-testid^="simple-collab-left-rail-sticky-filter-toggle-button-"] [id^="badge-"] {
      display: none !important;
    }

    html[${ROOT_ATTRIBUTE}] [data-testid="hamburger-button-badge"] {
      background: transparent !important;
      border: 0 !important;
      border-radius: 0 !important;
      box-shadow: none !important;
      color: var(--colorNeutralForeground3, #616161) !important;
      font-size: 11px !important;
      font-weight: 400 !important;
      line-height: 16px !important;
      min-width: auto !important;
      padding: 0 !important;
    }

    html[${ROOT_ATTRIBUTE}] [data-testid="hamburger-button-badge"] {
      position: static !important;
    }

    html[${ROOT_ATTRIBUTE}] [data-testid="slice-list-item-slice-all-messages"] [id^="title-"],
    html[${ROOT_ATTRIBUTE}] [role="treeitem"] [id^="title-"] {
      font-weight: 400 !important;
    }

    html[${ROOT_ATTRIBUTE}] [data-testid*="unread-indicator"],
    html[${ROOT_ATTRIBUTE}] [data-tid*="unread"],
    html[${ROOT_ATTRIBUTE}] [id*="unread"] {
      display: none !important;
    }
  `;
  document.head.append(style);

  function readEnabled() {
    return localStorage.getItem(SETTINGS_KEY) === "true";
  }

  function setEnabled(enabled) {
    localStorage.setItem(SETTINGS_KEY, String(enabled));
    document.documentElement.toggleAttribute(ROOT_ATTRIBUTE, enabled);
    const button = document.querySelector(`#${TOGGLE_ID} button`);
    if (button) {
      button.setAttribute("aria-checked", String(enabled));
      button.style.background = enabled
        ? "var(--colorBrandBackground, #6264a7)"
        : "var(--colorNeutralBackground1, #fff)";
      button.style.border = "0";
      button.style.boxShadow = "none";
      button.style.outline = "none";
      button.style.justifyContent = enabled ? "flex-end" : "flex-start";
      const thumb = button.firstElementChild;
      if (thumb) {
        thumb.style.background = enabled
          ? "var(--colorNeutralForegroundOnBrand, #fff)"
          : "var(--colorNeutralForeground3, #616161)";
      }
    }
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
    label.textContent = "Focus";

    const button = document.createElement("button");
    button.type = "button";
    button.setAttribute("role", "switch");
    button.setAttribute("aria-label", "Focus mode");
    button.title = "Reduce visual emphasis for unread messages and notifications";
    Object.assign(button.style, {
      alignItems: "center",
      borderRadius: "10px",
      border: "0",
      boxShadow: "none",
      cursor: "pointer",
      display: "inline-flex",
      height: "20px",
      outline: "none",
      padding: "2px",
      width: "36px",
    });

    const thumb = document.createElement("span");
    Object.assign(thumb.style, {
      borderRadius: "50%",
      display: "block",
      height: "14px",
      width: "14px",
    });
    button.append(thumb);
    button.addEventListener("click", event => {
      event.stopPropagation();
      setEnabled(!readEnabled());
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
    ringBadge.parentElement.after(createToggle(readEnabled()));
    setEnabled(readEnabled());
  }

  setEnabled(readEnabled());
  installToggle();
  const observer = new MutationObserver(installToggle);
  observer.observe(document.documentElement, { childList: true, subtree: true });

  globalThis.__teamsFocusMode = {
    destroy() {
      observer.disconnect();
      document.getElementById(TOGGLE_ID)?.remove();
      document.getElementById(STYLE_ID)?.remove();
      document.documentElement.removeAttribute(ROOT_ATTRIBUTE);
      delete globalThis.__teamsFocusMode;
    },
  };
})();
