// ==UserScript==
// @name         Teams User extensions
// @match        https://teams.microsoft.com/v2/*
// @match        https://teams.cloud.microsoft/v2/*
// @match        https://local.teams.office.com/v2/*
// @run-at       document-idle
// @toggleable   false
// ==/UserScript==

(() => {
  if (window.top !== window) return;

  globalThis.__teamsUserExtensions?.destroy();

  const DISABLED_STORAGE_KEY = "teams.userscripts.disabled";
  const MENU_ITEM_ID = "teams-user-extensions-menu-item";
  const MODAL_ID = "teams-user-extensions-modal";

  function readDisabledExtensions() {
    try {
      const value = JSON.parse(
        localStorage.getItem(DISABLED_STORAGE_KEY) ?? "[]"
      );
      return Array.isArray(value) ? new Set(value) : new Set();
    } catch (error) {
      console.error("[Teams User extensions]", error);
      return new Set();
    }
  }

  function writeDisabledExtensions(disabledExtensions) {
    localStorage.setItem(
      DISABLED_STORAGE_KEY,
      JSON.stringify([...disabledExtensions].sort())
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

  function closeModal() {
    document.getElementById(MODAL_ID)?.remove();
  }

  function openExtensionSettings(extension) {
    const settings = globalThis.__teamsUserscriptSettings?.[extension.name];
    if (typeof settings !== "function") return;
    closeModal();
    settings();
  }

  function createSwitch(extension, enabled) {
    const control = document.createElement("button");
    control.type = "button";
    control.setAttribute("role", "switch");
    control.setAttribute("aria-label", extension.name);
    control.setAttribute("aria-checked", String(enabled));
    Object.assign(control.style, {
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
      flex: "0 0 auto",
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
    control.append(thumb);
    control.addEventListener("click", async () => {
      control.disabled = true;
      const disabledExtensions = readDisabledExtensions();
      if (enabled) disabledExtensions.add(extension.name);
      else disabledExtensions.delete(extension.name);
      writeDisabledExtensions(disabledExtensions);
      try {
        await restartTeams();
      } catch (error) {
        control.disabled = false;
        console.error("[Teams User extensions]", error);
      }
    });
    return control;
  }

  function openModal() {
    closeModal();
    const disabledExtensions = readDisabledExtensions();
    const extensions = (globalThis.__teamsUserscriptManifest ?? []).filter(
      extension => extension.toggleable
    );

    const backdrop = document.createElement("div");
    backdrop.id = MODAL_ID;
    Object.assign(backdrop.style, {
      alignItems: "center",
      background: "rgba(0, 0, 0, 0.4)",
      display: "flex",
      inset: "0",
      justifyContent: "center",
      position: "fixed",
      zIndex: "2147483647",
    });

    const dialog = document.createElement("div");
    dialog.setAttribute("role", "dialog");
    dialog.setAttribute("aria-modal", "true");
    dialog.setAttribute("aria-labelledby", `${MODAL_ID}-title`);
    Object.assign(dialog.style, {
      background: "var(--colorNeutralBackground1, #fff)",
      borderRadius: "8px",
      boxShadow: "0 8px 32px rgba(0, 0, 0, 0.28)",
      color: "var(--colorNeutralForeground1, #242424)",
      font: "14px/20px -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif",
      maxHeight: "70vh",
      minWidth: "420px",
      overflow: "auto",
      padding: "20px",
    });

    const header = document.createElement("div");
    Object.assign(header.style, {
      alignItems: "center",
      display: "flex",
      justifyContent: "space-between",
      marginBottom: "16px",
    });
    const title = document.createElement("h2");
    title.id = `${MODAL_ID}-title`;
    title.textContent = "User extensions";
    Object.assign(title.style, {
      fontSize: "20px",
      lineHeight: "28px",
      margin: "0",
    });
    const closeButton = document.createElement("button");
    closeButton.type = "button";
    closeButton.setAttribute("aria-label", "Close");
    closeButton.textContent = "×";
    Object.assign(closeButton.style, {
      background: "transparent",
      border: "0",
      color: "inherit",
      cursor: "pointer",
      fontSize: "24px",
      height: "32px",
      lineHeight: "24px",
      width: "32px",
    });
    closeButton.addEventListener("click", closeModal);
    header.append(title, closeButton);
    dialog.append(header);

    if (!extensions.length) {
      const empty = document.createElement("p");
      empty.textContent = "No user extensions are available.";
      dialog.append(empty);
    } else {
      for (const extension of extensions) {
        const enabled = !disabledExtensions.has(extension.name);
        const row = document.createElement("div");
        Object.assign(row.style, {
          alignItems: "center",
          borderTop: "1px solid var(--colorNeutralStroke2, #e0e0e0)",
          display: "flex",
          gap: "24px",
          justifyContent: "space-between",
          minHeight: "52px",
        });
        const label = document.createElement("span");
        label.textContent = extension.name;
        const actions = document.createElement("span");
        Object.assign(actions.style, {
          alignItems: "center",
          display: "inline-flex",
          gap: "12px",
        });
        const settings = globalThis.__teamsUserscriptSettings?.[extension.name];
        if (typeof settings === "function") {
          const settingsLink = document.createElement("button");
          settingsLink.type = "button";
          settingsLink.textContent = "Settings";
          Object.assign(settingsLink.style, {
            background: "transparent",
            border: "0",
            color: "var(--colorBrandForeground1, #5b5fc7)",
            cursor: "pointer",
            font: "inherit",
            padding: "4px 0",
            textDecoration: "underline",
          });
          settingsLink.addEventListener("click", () =>
            openExtensionSettings(extension)
          );
          actions.append(settingsLink);
        }
        actions.append(createSwitch(extension, enabled));
        row.append(label, actions);
        dialog.append(row);
      }
    }

    backdrop.append(dialog);
    backdrop.addEventListener("click", event => {
      if (event.target === backdrop) closeModal();
    });
    document.body.append(backdrop);
    closeButton.focus();
  }

  function createMenuItem(referenceItem) {
    const item = document.createElement("div");
    item.id = MENU_ITEM_ID;
    item.className = referenceItem.className;
    item.setAttribute("role", "menuitem");
    item.tabIndex = 0;
    item.textContent = "User extensions";
    item.addEventListener("click", openModal);
    item.addEventListener("keydown", event => {
      if (event.key !== "Enter" && event.key !== " ") return;
      event.preventDefault();
      openModal();
    });
    return item;
  }

  function installMenuItem() {
    if (document.getElementById(MENU_ITEM_ID)) return;
    const ringMenuItem = [
      ...document.querySelectorAll("[data-tid='ringswitcher']"),
    ].find(
      element =>
        element.closest("[role='menu']") && element.getClientRects().length
    );
    const menu = ringMenuItem?.closest("[role='menu']");
    if (!ringMenuItem || !menu) return;
    const referenceItem =
      menu.querySelector("[data-tid='settings-button-menu']") ?? ringMenuItem;
    ringMenuItem.before(createMenuItem(referenceItem));
  }

  function handleKeydown(event) {
    if (event.key === "Escape" && document.getElementById(MODAL_ID)) {
      event.preventDefault();
      event.stopImmediatePropagation();
      closeModal();
    }
  }

  const observer = new MutationObserver(installMenuItem);
  observer.observe(document.documentElement, { childList: true, subtree: true });
  window.addEventListener("keydown", handleKeydown, true);
  installMenuItem();

  globalThis.__teamsUserExtensions = {
    destroy() {
      observer.disconnect();
      window.removeEventListener("keydown", handleKeydown, true);
      document.getElementById(MENU_ITEM_ID)?.remove();
      closeModal();
      delete globalThis.__teamsUserExtensions;
    },
  };
})();
