// ==UserScript==
// @name         Re-color
// @match        https://teams.microsoft.com/v2/*
// @match        https://teams.cloud.microsoft/v2/*
// @match        https://local.teams.office.com/v2/*
// @match        https://outlook.office.com/hosted/calendar/*
// @run-at       document-idle
// ==/UserScript==

(() => {
  const isCalendar = window.location.hostname === "outlook.office.com";
  if (!isCalendar && window.top !== window) return;

  if (!isCalendar) globalThis.__teamsReColor?.destroy();

  const NAME = "Re-color";
  const STORAGE_KEY = "teams.userscripts.re-color.primary";
  const ROOT_ATTRIBUTE = "data-teams-re-color";
  const MODAL_ID = "teams-re-color-settings";
  const DEFAULT_PRIMARY = "#6264a7";
  let overridePrimary;

  function readPrimary() {
    const value =
      overridePrimary ?? localStorage.getItem(STORAGE_KEY) ?? DEFAULT_PRIMARY;
    return /^#[0-9a-f]{6}$/i.test(value) ? value : DEFAULT_PRIMARY;
  }

  function secondaryColour(primary) {
    const value = Number.parseInt(primary.slice(1), 16);
    const lighten = channel => Math.round(channel + (255 - channel) * 0.9);
    const red = lighten((value >> 16) & 255);
    const green = lighten((value >> 8) & 255);
    const blue = lighten(value & 255);
    return `#${[red, green, blue]
      .map(channel => channel.toString(16).padStart(2, "0"))
      .join("")}`;
  }

  function applyColours() {
    const primary = readPrimary();
    const secondary = secondaryColour(primary);
    document.documentElement.style.setProperty("--teams-re-color-primary", primary);
    document.documentElement.style.setProperty(
      "--teams-re-color-secondary",
      secondary
    );
    if (isCalendar) {
      document.documentElement.style.removeProperty("--colorBrandBackground");
      document.documentElement.style.removeProperty("--colorBrandBackgroundHover");
      document.documentElement.style.removeProperty("--colorBrandBackgroundPressed");
      document.documentElement.style.removeProperty("--colorBrandBackground2");
      document.documentElement.style.removeProperty("--colorBrandBackground2Hover");
      for (const element of document.querySelectorAll("body *")) {
        if (
          element.style.getPropertyPriority("background-color") === "important" &&
          !element.matches(
            ".fui-CalendarDayGrid__daySelected, .fui-CalendarDayGrid__daySelected *, .czFpU, .lJS6W"
          )
        ) {
          element.style.removeProperty("background-color");
        }
      }
    }
    if (!isCalendar) {
      document.documentElement.style.setProperty("--colorBrandStroke1", primary);
      document.documentElement.style.setProperty("--colorBrandForeground1", primary);
    } else {
      document.documentElement.style.removeProperty("--colorBrandStroke1");
      document.documentElement.style.removeProperty("--colorBrandForeground1");
    }
    if (!isCalendar) {
      for (const frame of document.querySelectorAll("iframe")) {
        frame.contentWindow?.postMessage({ type: "teams-re-color", primary }, "*");
      }
    } else {
      for (const element of document.querySelectorAll(
        ".fui-CalendarDayGrid__daySelected"
      )) {
        element.style.setProperty("background-color", secondary, "important");
        element.style.setProperty("border-color", primary, "important");
        element.querySelector("button")?.style.setProperty(
          "background-color",
          secondary,
          "important"
        );
        element.querySelector("button")?.style.setProperty(
          "color",
          primary,
          "important"
        );
      }
      for (const element of document.querySelectorAll(".czFpU, .lJS6W")) {
        element.style.setProperty("background-color", primary, "important");
        element.style.setProperty("border-color", primary, "important");
      }
      for (const element of document.querySelectorAll("body *")) {
        if (
          !element.matches(
            ".fui-CalendarDayGrid__daySelected, .fui-CalendarDayGrid__daySelected *, .czFpU, .lJS6W"
          ) &&
          (element.style.backgroundColor === primary ||
            element.style.backgroundColor === secondary)
        ) {
          element.style.removeProperty("background-color");
        }
      }
    }
    for (const element of document.querySelectorAll("body *")) {
      const computed = getComputedStyle(element);
      if (
        computed.color === "rgb(98, 100, 167)" ||
        computed.color === "rgb(91, 95, 199)" ||
        computed.color === "rgb(68, 71, 145)"
      ) {
        element.style.setProperty("color", primary, "important");
      }
      if (
        computed.borderTopColor === "rgb(98, 100, 167)" ||
        computed.borderTopColor === "rgb(91, 95, 199)" ||
        computed.borderTopColor === "rgb(68, 71, 145)"
      ) {
        element.style.setProperty("border-color", primary, "important");
      }
      if (
        !isCalendar &&
        (computed.fill === "rgb(91, 95, 199)" ||
          computed.fill === "rgb(68, 71, 145)" ||
          computed.stroke === "rgb(91, 95, 199)" ||
          computed.stroke === "rgb(68, 71, 145)")
      ) {
        element.style.setProperty("fill", primary, "important");
        element.style.setProperty("stroke", primary, "important");
      }
      if (
        computed.fontFamily.includes("FluentSystemIcons") ||
        element.className?.toString().includes("Icon-font")
      ) {
        element.style.setProperty("font-weight", "400", "important");
      }
    }
  }

  function closeSettings() {
    document.getElementById(MODAL_ID)?.remove();
  }

  function openSettings() {
    closeSettings();
    const primary = readPrimary();
    const secondary = secondaryColour(primary);
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
    dialog.setAttribute("aria-label", "Re-color settings");
    Object.assign(dialog.style, {
      background: "var(--colorNeutralBackground1, #fff)",
      borderRadius: "8px",
      boxShadow: "0 8px 32px rgba(0, 0, 0, 0.28)",
      color: "var(--colorNeutralForeground1, #242424)",
      font: "14px/20px -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif",
      minWidth: "360px",
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
    title.textContent = "Re-color settings";
    title.style.margin = "0";
    const close = document.createElement("button");
    close.type = "button";
    close.textContent = "×";
    close.setAttribute("aria-label", "Close");
    close.style.cssText = "background:transparent;border:0;color:inherit;cursor:pointer;font-size:24px";
    close.addEventListener("click", closeSettings);
    header.append(title, close);
    const label = document.createElement("label");
    label.textContent = "Primary colour";
    label.style.display = "block";
    const input = document.createElement("input");
    input.type = "color";
    input.value = primary;
    input.style.cssText = "cursor:pointer;display:block;height:40px;margin-top:8px;width:100%";
    const preview = document.createElement("div");
    preview.textContent = `Secondary colour: ${secondary}`;
    preview.style.cssText = "border-radius:4px;margin-top:12px;padding:10px";
    const update = () => {
      localStorage.setItem(STORAGE_KEY, input.value);
      applyColours();
      preview.textContent = `Secondary colour: ${secondaryColour(input.value)}`;
      preview.style.background = secondaryColour(input.value);
      preview.style.color = "#fff";
      for (const frame of document.querySelectorAll("iframe")) {
        frame.contentWindow?.postMessage(
          { type: "teams-re-color", primary: input.value },
          "*"
        );
      }
    };
    input.addEventListener("input", update);
    dialog.append(header, label, input, preview);
    backdrop.append(dialog);
    backdrop.addEventListener("click", event => {
      if (event.target === backdrop) closeSettings();
    });
    document.body.append(backdrop);
    close.focus();
    update();
  }

  const style = document.createElement("style");
  style.setAttribute(ROOT_ATTRIBUTE, "style");
  style.textContent = `
    :root {
      --teams-re-color-primary: ${readPrimary()};
      --teams-re-color-secondary: ${secondaryColour(readPrimary())};
      --colorBrandBackgroundSelected: var(--teams-re-color-secondary) !important;
      --colorBrandStroke1: var(--teams-re-color-primary) !important;
      --colorBrandStroke2: var(--teams-re-color-primary) !important;
      --colorBrandForeground1: var(--teams-re-color-primary) !important;
      --colorBrandForeground2: var(--teams-re-color-primary) !important;
      --colorCompoundBrandBackground: var(--teams-re-color-primary) !important;
      --colorCompoundBrandBackgroundHover: var(--teams-re-color-primary) !important;
      --colorCompoundBrandBackgroundPressed: var(--teams-re-color-primary) !important;
      --colorCompoundBrandForeground1: var(--teams-re-color-primary) !important;
      --colorCompoundBrandStroke: var(--teams-re-color-primary) !important;
    }
    [data-tid='titlebar-status-indicator'],
    [data-tid='titlebar-status-indicator'] + *,
    [data-tid='selected'],
    [aria-selected='true'],
    [role='tab'][aria-selected='true'] {
      accent-color: var(--teams-re-color-primary);
    }
    [role='tab'][aria-selected='true']::after,
    [aria-selected='true']::after {
      background-color: var(--teams-re-color-primary) !important;
    }
    [role='tab'][aria-selected='false'],
    [role='tab']:not([aria-selected='true']) {
      color: var(--colorNeutralForeground1, #242424) !important;
      font-weight: 400 !important;
    }
    [aria-label^='Activity'],
    [aria-label^='Chat'],
    [aria-label^='Calendar'],
    [aria-label^='Calls'],
    [aria-label^='OneDrive'],
    [aria-label^='Copilot'],
    [aria-label^='Approvals'],
    [aria-label^='Events'] {
      font-weight: 400 !important;
    }
    [aria-label^='Activity']::before,
    [aria-label^='Chat']::before,
    [aria-label^='Calls']::before,
    [aria-label^='OneDrive']::before,
    [aria-label^='Copilot']::before,
    [aria-label^='Approvals']::before,
    [aria-label^='Events']::before,
    [aria-label^='Store']::before {
      background-color: transparent !important;
      border-color: transparent !important;
    }
    [aria-label^='Calendar']::before {
      background-color: var(--teams-re-color-primary) !important;
      border-color: var(--teams-re-color-primary) !important;
    }
    [role='tab']::before,
    [role='tab']::after,
    [aria-selected='true']::before,
    [aria-selected='true']::after {
      border-color: var(--teams-re-color-primary) !important;
    }
    .fui-Icon-font,
    [class*='Icon-font'],
    [class*='iconFont'],
    [class*='IconFont'],
    .fui-Button__icon,
    .fui-MenuButton__menuIcon,
    [class*='icon-'],
    .PVqL_ {
      font-weight: 400 !important;
    }
    [data-tid='chat-list-item'][aria-selected='true'],
    [data-tid='chat-list-item'][data-is-selected='true'],
    [data-tid='chat-pane-item'][aria-selected='true'],
    [data-tid='chat-pane-item'][data-is-selected='true'],
    [data-tabster*='LeftRailSelectedItem'],
    [role='treeitem'][data-item-type='chat'][tabindex='0'],
    .fui-ChatMyMessage__body,
    [data-tid='chat-message'][data-is-selected='true'],
    [data-tid='message-bubble'][data-is-selected='true'] {
      background-color: var(--teams-re-color-secondary) !important;
    }
    a[href],
    [role='link'],
    [data-tid='last-read-marker'],
    [data-tid='last-read-marker']::before,
    [data-tid='last-read-marker']::after,
    [data-tid='last-read-line'],
    [data-tid='last-read-line']::before,
    [data-tid='last-read-line']::after,
    [data-tid='last-read-line'] > div {
      color: var(--teams-re-color-primary) !important;
    }
    a[href]:visited,
    [role='link']:visited {
      color: var(--teams-re-color-primary) !important;
    }
    [role='switch'][aria-checked='true'] {
      background-color: var(--teams-re-color-primary) !important;
      border-color: var(--teams-re-color-primary) !important;
    }
    [role='switch'][aria-checked='true'] > *,
    [role='switch'][aria-checked='true']::after {
      background-color: #fff !important;
    }
    [role='switch'][aria-checked='false'] {
      background-color: #fff !important;
      border-color: var(--colorNeutralStroke1, #616161) !important;
    }
    [role='switch'][aria-checked='false'] > *,
    [role='switch'][aria-checked='false']::after {
      background-color: var(--colorNeutralForeground3, #616161) !important;
    }
    [data-tid='unread-count'],
    [data-tid='draft-badge'],
    .fui-Badge,
    [data-tid='titlebar-status-indicator'] {
      background-color: var(--teams-re-color-primary) !important;
    }
    [aria-current='page'],
    [aria-current='page']::after {
      color: var(--teams-re-color-primary) !important;
    }
    [data-tid='last-read-marker'] {
      border-color: var(--teams-re-color-primary) !important;
    }
    [data-tid='last-read-line']::before,
    [data-tid='last-read-line']::after {
      background-color: var(--teams-re-color-primary) !important;
    }
    [role='treeitem'][data-item-type='chat'] svg,
    [role='treeitem'][data-item-type='chat'] [data-inp] svg {
      color: var(--teams-re-color-primary) !important;
    }
    .fui-CalendarDayGrid__daySelected {
      background-color: var(--teams-re-color-secondary) !important;
      border-color: var(--teams-re-color-primary) !important;
      border-top-color: var(--teams-re-color-primary) !important;
    }
    .fui-CalendarDayGrid__daySelected > .fui-CalendarDayGrid__dayButton {
      background-color: var(--teams-re-color-secondary) !important;
    }
    .fui-CalendarDayGrid__daySelected .fui-CalendarDayGrid__dayButton {
      color: var(--teams-re-color-primary) !important;
    }
    [aria-label*='September'],
    [aria-label*='October'],
    [aria-label*='November'],
    [aria-label*='December'],
    [aria-label*='January'],
    [aria-label*='February'],
    [aria-label*='March'],
    [aria-label*='April'],
    [aria-label*='May'],
    [aria-label*='June'],
    [aria-label*='July'],
    [aria-label*='August'] {
      --calendar-re-color-primary: var(--teams-re-color-primary);
    }
    .fui-CalendarDayGrid__dayButton[aria-current='date'],
    .fui-CalendarDayGrid__dayButton[aria-selected='true'] {
      background-color: var(--teams-re-color-primary) !important;
      color: #fff !important;
    }
    button[aria-label='Date selector'],
    .fui-CalendarDay__headerIconButton,
    .fui-CalendarDayGrid__dayButton[aria-current='date'],
    .fui-CalendarDayGrid__dayButton[aria-selected='true'],
    .fui-CalendarDay__monthAndYear,
    .fui-CalendarDayGrid__daySelected .fui-CalendarDayGrid__dayNumber {
      color: var(--teams-re-color-primary) !important;
    }
    .fui-CalendarDayGrid__daySelected::before,
    .fui-CalendarDayGrid__daySelected::after,
    .fui-CalendarDayGrid__dayButton[aria-current='date']::before,
    .fui-CalendarDayGrid__dayButton[aria-selected='true']::before {
      background-color: var(--teams-re-color-secondary) !important;
      border-color: var(--teams-re-color-primary) !important;
    }
    .fui-CalendarDayGrid__daySelected::before {
      background-color: var(--teams-re-color-primary) !important;
    }
    .czFpU,
    .lJS6W {
      background-color: var(--teams-re-color-primary) !important;
      border-color: var(--teams-re-color-primary) !important;
    }
    .fui-CalendarDayGrid__dayButton,
    .fui-CalendarDay__headerIconButton,
    .fui-CalendarDay__monthAndYear,
    [aria-label*='calendar' i] svg,
    [aria-label*='month' i] svg,
    [aria-label*='week' i] svg,
    [aria-label*='today' i] svg,
    [aria-label*='previous' i] svg,
    [aria-label*='next' i] svg,
    [aria-label*='date selector' i] svg {
      accent-color: var(--teams-re-color-primary);
    }
    .BD_hL,
    .PVqL_,
    .DxDL9,
    .B5Pd6,
    .k65Bx,
    .nUQ_t,
    .Hx8ak,
    .bIymH,
    .NPc97,
    .zk3dz,
    .E2drh,
    .C8Gpn,
    .FLwLv,
    .Q0K3G,
    .fui-CalendarDayGrid__dayNumber,
    .fui-CalendarDayGrid__daySelected,
    .fui-CalendarDayGrid__daySelected * {
      color: var(--teams-re-color-primary) !important;
    }
    .BD_hL {
      border-color: var(--teams-re-color-primary) !important;
    }
    .fui-CalendarDayGrid__daySelected,
    .fui-CalendarDayGrid__daySelected > *,
    .fui-CalendarDayGrid__daySelected::before,
    .fui-CalendarDayGrid__daySelected::after {
      background-color: var(--teams-re-color-secondary) !important;
    }
    [aria-current='page'] svg,
    [aria-selected='true'] svg,
    [data-tid] svg {
      color: var(--teams-re-color-primary) !important;
      fill: currentColor !important;
      stroke: currentColor !important;
    }
    [class*='___5pnvcg0'] {
      color: var(--teams-re-color-primary) !important;
    }
    i[class*='Icon-font'],
    [class*='Icon-font'],
    [class*='iconFont'] {
      font-weight: 400 !important;
      font-style: normal !important;
    }
    .BD_hL::before,
    .BD_hL::after,
    .czFpU::before,
    .czFpU::after,
    .lJS6W::before,
    .lJS6W::after,
    .fui-CalendarDayGrid__daySelected::before,
    .fui-CalendarDayGrid__daySelected::after {
      background-color: var(--teams-re-color-primary) !important;
      border-color: var(--teams-re-color-primary) !important;
    }
    [style*='color: rgb(91, 95, 199)'],
    [style*='color: rgb(68, 71, 145)'],
    [style*='color: rgb(98, 100, 167)'] {
      color: var(--teams-re-color-primary) !important;
    }
    button svg,
    [role='button'] svg,
    [role='link'] svg,
    [role='tab'] svg {
      color: currentColor;
    }
    button svg[fill]:not([fill='none']),
    [role='button'] svg[fill]:not([fill='none']),
    [role='link'] svg[fill]:not([fill='none']),
    button svg [fill]:not([fill='none']),
    [role='button'] svg [fill]:not([fill='none']),
    [role='link'] svg [fill]:not([fill='none']) {
      fill: currentColor !important;
    }
    button svg[stroke]:not([stroke='none']),
    [role='button'] svg[stroke]:not([stroke='none']),
    [role='link'] svg[stroke]:not([stroke='none']),
    button svg [stroke]:not([stroke='none']),
    [role='button'] svg [stroke]:not([stroke='none']),
    [role='link'] svg [stroke]:not([stroke='none']) {
      stroke: currentColor !important;
    }
  `;
  document.head.append(style);
  applyColours();
  const syncTimer = window.setInterval(applyColours, 1000);
  window.addEventListener("message", event => {
    if (
      isCalendar &&
      event.data?.type === "teams-re-color" &&
      typeof event.data.primary === "string" &&
      /^#[0-9a-f]{6}$/i.test(event.data.primary)
    ) {
      overridePrimary = event.data.primary;
      applyColours();
    }
  });
  const observer = new MutationObserver(applyColours);
  observer.observe(document.documentElement, { childList: true, subtree: true });
  if (!isCalendar) {
    globalThis.__teamsUserscriptSettings ??= {};
    globalThis.__teamsUserscriptSettings[NAME] = openSettings;
  }
  globalThis.__teamsReColor = {
    destroy() {
      observer.disconnect();
      if (syncTimer !== undefined) window.clearInterval(syncTimer);
      style.remove();
      if (!isCalendar) {
        delete globalThis.__teamsUserscriptSettings?.[NAME];
        delete globalThis.__teamsReColor;
      }
    },
  };
})();
