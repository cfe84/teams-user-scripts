// ==UserScript==
// @name         Teams Vimium navigation
// @match        https://teams.microsoft.com/*
// @match        https://teams.cloud.microsoft/*
// @match        https://local.teams.office.com/*
// @match        https://outlook.office.com/hosted/calendar/*
// @match        https://onedrive.cloud.microsoft/*
// @run-at       document-idle
// ==/UserScript==

globalThis.__teamsVimium?.destroy();
document
  .querySelectorAll("[data-teams-vimium='root']")
  .forEach(element => element.remove());

(() => {
  const HINT_KEYS = "asdfghjklqwertyuiopzxcvbnm";
  const OWN_ATTRIBUTE = "data-teams-vimium";
  const FIND_HIGHLIGHT = "teams-vimium-find";
  const CURRENT_FIND_HIGHLIGHT = "teams-vimium-find-current";
  const NAVIGATION_SEQUENCE_TIMEOUT = 700;
  const SCROLL_STEP = 80;
  const IS_HOSTED_CALENDAR =
    location.hostname === "outlook.office.com" &&
    location.pathname.startsWith("/hosted/calendar/");
  const BASE_HINT_FIRST_KEYS = IS_HOSTED_CALENDAR
    ? HINT_KEYS.slice(Math.ceil(HINT_KEYS.length / 2))
    : HINT_KEYS.slice(0, Math.ceil(HINT_KEYS.length / 2));
  const CLICKABLE_SELECTOR = [
    "a[href]",
    "button",
    "input:not([type='hidden'])",
    "select",
    "textarea",
    "[contenteditable='true']",
    "[role='button']",
    "[role='checkbox']",
    "[role='combobox']",
    "[role='link']",
    "[role='menuitem']",
    "[role='menuitemcheckbox']",
    "[role='menuitemradio']",
    "[role='option']",
    "[role='radio']",
    "[role='searchbox']",
    "[role='switch']",
    "[role='tab']",
    "[role='textbox']",
    "[tabindex]:not([tabindex='-1'])",
  ].join(",");
  const TEXT_INPUT_SELECTOR = [
    "input:not([type='button']):not([type='checkbox']):not([type='radio']):not([type='submit'])",
    "textarea",
    "[contenteditable='true']",
    "[role='textbox']",
    "[role='searchbox']",
    "[role='combobox']",
  ].join(",");

  const state = {
    mode: "normal",
    pendingNavigationKey: "",
    pendingNavigationTimer: 0,
    activePane: null,
    activeChatRow: null,
    activeChatKey: null,
    activeCalendarItem: null,
    find: {
      query: "",
      ranges: [],
      currentIndex: -1,
    },
    hints: [],
    hintInput: "",
    relayedKey: false,
    hintRelayKeysRemaining: 0,
    vomnibarEntries: [],
    vomnibarResults: [],
    vomnibarIndex: 0,
    vomnibarRequest: 0,
  };

  const cleanupTasks = [];
  const rootHost = document.createElement("div");
  rootHost.setAttribute(OWN_ATTRIBUTE, "root");
  document.documentElement.append(rootHost);
  const root = rootHost.attachShadow({ mode: "open" });

  const style = document.createElement("style");
  style.textContent = `
      :host {
        all: initial;
        color-scheme: light dark;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      }
      * { box-sizing: border-box; }
      .hidden { display: none !important; }
      .mode {
        position: fixed;
        right: 12px;
        bottom: 12px;
        z-index: 2147483647;
        border: 1px solid rgba(255, 255, 255, 0.3);
        border-radius: 4px;
        padding: 3px 7px;
        color: white;
        background: #292929;
        font: 600 11px/16px ui-monospace, SFMono-Regular, Consolas, monospace;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.35);
        pointer-events: none;
      }
      .find {
        position: fixed;
        right: 12px;
        bottom: 12px;
        z-index: 2147483647;
        display: flex;
        align-items: center;
        gap: 8px;
        width: min(420px, calc(100vw - 24px));
        border: 1px solid #777;
        border-radius: 6px;
        padding: 7px 10px;
        color: CanvasText;
        background: Canvas;
        box-shadow: 0 5px 24px rgba(0, 0, 0, 0.35);
      }
      .find input, .vomnibar input {
        min-width: 0;
        flex: 1;
        border: 0;
        outline: 0;
        color: CanvasText;
        background: transparent;
        font: 14px/22px -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      }
      .count {
        color: GrayText;
        font: 12px/18px ui-monospace, SFMono-Regular, Consolas, monospace;
        white-space: nowrap;
      }
      .hint {
        position: fixed;
        z-index: 2147483646;
        min-width: 18px;
        border: 1px solid #765500;
        border-radius: 3px;
        padding: 1px 3px;
        color: #171000;
        background: #ffdf58;
        font: 700 11px/14px ui-monospace, SFMono-Regular, Consolas, monospace;
        text-align: center;
        text-transform: uppercase;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.45);
        pointer-events: none;
      }
      .hint .matched { color: #a4262c; }
      .vomnibar {
        position: fixed;
        top: 12%;
        left: 50%;
        z-index: 2147483647;
        width: min(680px, calc(100vw - 32px));
        max-height: 70vh;
        overflow: hidden;
        border: 1px solid #777;
        border-radius: 8px;
        color: CanvasText;
        background: Canvas;
        box-shadow: 0 12px 48px rgba(0, 0, 0, 0.45);
        transform: translateX(-50%);
      }
      .vomnibar-input {
        display: flex;
        gap: 8px;
        border-bottom: 1px solid #777;
        padding: 11px 14px;
      }
      .vomnibar-list {
        max-height: calc(70vh - 46px);
        overflow: auto;
      }
      .vomnibar-item {
        display: grid;
        grid-template-columns: minmax(0, 1fr) auto;
        gap: 12px;
        padding: 8px 14px;
        font: 14px/20px -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      }
      .vomnibar-item.selected {
        color: white;
        background: #5b5fc7;
      }
      .vomnibar-label {
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
      }
      .vomnibar-kind {
        color: GrayText;
        font-size: 11px;
        text-transform: uppercase;
      }
      .vomnibar-item.selected .vomnibar-kind { color: #eee; }
      .help-backdrop {
        position: fixed;
        inset: 0;
        z-index: 2147483647;
        display: grid;
        place-items: center;
        padding: 24px;
        background: rgba(0, 0, 0, 0.45);
      }
      .help {
        width: min(760px, 100%);
        max-height: min(720px, 90vh);
        overflow: auto;
        border: 1px solid #777;
        border-radius: 10px;
        padding: 20px 24px;
        color: CanvasText;
        background: Canvas;
        box-shadow: 0 12px 48px rgba(0, 0, 0, 0.5);
      }
      .help h1 {
        margin: 0 0 16px;
        font: 600 20px/26px -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      }
      .help-grid {
        display: grid;
        grid-template-columns: max-content minmax(140px, 1fr) max-content minmax(140px, 1fr);
        gap: 7px 12px;
        font: 13px/20px -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      }
      kbd {
        align-self: start;
        border: 1px solid #888;
        border-radius: 3px;
        padding: 0 5px;
        background: color-mix(in srgb, Canvas 90%, CanvasText);
        font: 12px/18px ui-monospace, SFMono-Regular, Consolas, monospace;
        white-space: nowrap;
      }
      @media (max-width: 620px) {
        .help-grid { grid-template-columns: max-content 1fr; }
      }
  `;

  function createElement(tagName, options = {}, children = []) {
    const element = document.createElement(tagName);
    if (options.id) element.id = options.id;
    if (options.className) element.className = options.className;
    if (options.text != null) element.textContent = options.text;
    for (const [name, value] of Object.entries(options.attributes ?? {})) {
      element.setAttribute(name, value);
    }
    element.append(...children);
    return element;
  }

  const mode = createElement("div", {
    id: "mode",
    className: "mode",
    text: "NORMAL",
  });
  const findInput = createElement("input", {
    id: "find-input",
    attributes: {
      type: "text",
      "aria-label": "Find in Teams",
      autocomplete: "off",
    },
  });
  const findCount = createElement("span", {
    id: "find-count",
    className: "count",
    text: "0/0",
  });
  const find = createElement(
    "div",
    { id: "find", className: "find hidden" },
    [createElement("span", { text: "/" }), findInput, findCount]
  );
  const hints = createElement("div", { id: "hints" });
  const vomnibarInput = createElement("input", {
    id: "vomnibar-input",
    attributes: {
      type: "text",
      "aria-label": "Open Teams item",
      autocomplete: "off",
    },
  });
  const vomnibarCount = createElement("span", {
    id: "vomnibar-count",
    className: "count",
    text: "0",
  });
  const vomnibarList = createElement("div", {
    id: "vomnibar-list",
    className: "vomnibar-list",
  });
  const vomnibar = createElement(
    "div",
    { id: "vomnibar", className: "vomnibar hidden" },
    [
      createElement("div", { className: "vomnibar-input" }, [
        createElement("span", { text: "Open" }),
        vomnibarInput,
        vomnibarCount,
      ]),
      vomnibarList,
    ]
  );
  const helpEntries = [
    ["Esc", "Normal mode / unfocus"],
    ["j / Ctrl+e", IS_HOSTED_CALENDAR ? "Next event down" : "Scroll down"],
    ["k / Ctrl+y", IS_HOSTED_CALENDAR ? "Next event up" : "Scroll up"],
    ["d / u", "Half page down / up"],
    [
      "h / l",
      IS_HOSTED_CALENDAR ? "Next event left / right" : "Select left / right pane",
    ],
    [
      "[[ / ]]",
      IS_HOSTED_CALENDAR ? "Previous / next period" : "Previous / next screen",
    ],
    ["x", "Close top modal or menu"],
    ["i", "Focus main text box"],
    ["/", "Find text"],
    ["n / N", "Next / previous find"],
    ["f / Ctrl+F", "Open link hints"],
    ["o", "Open vomnibar"],
    ["?", "Toggle this help"],
    ["Up / Down", "Navigate vomnibar"],
    ["Enter", "Activate selection"],
  ];
  const helpGrid = createElement(
    "div",
    { className: "help-grid" },
    helpEntries.flatMap(([key, description]) => [
      createElement("kbd", { text: key }),
      createElement("span", { text: description }),
    ])
  );
  const helpBackdrop = createElement(
    "div",
    { id: "help-backdrop", className: "help-backdrop hidden" },
    [
      createElement(
        "section",
        {
          className: "help",
          attributes: {
            role: "dialog",
            "aria-modal": "true",
            "aria-label": "Teams Vimium help",
          },
        },
        [
          createElement("h1", { text: "Teams Vimium navigation" }),
          helpGrid,
        ]
      ),
    ]
  );
  root.append(style, mode, find, hints, vomnibar, helpBackdrop);

  const ui = {
    mode,
    find,
    findInput,
    findCount,
    hints,
    vomnibar,
    vomnibarInput,
    vomnibarCount,
    vomnibarList,
    helpBackdrop,
  };

  function listen(target, type, handler, options) {
    target.addEventListener(type, handler, options);
    cleanupTasks.push(() => target.removeEventListener(type, handler, options));
  }

  function isVisible(element) {
    if (!(element instanceof Element) || element.closest(`[${OWN_ATTRIBUTE}]`)) {
      return false;
    }
    const style = getComputedStyle(element);
    if (
      style.display === "none" ||
      style.visibility === "hidden" ||
      style.pointerEvents === "none" ||
      (Number(style.opacity) === 0 && !element.matches("[role='switch']"))
    ) {
      return false;
    }
    const rect = element.getBoundingClientRect();
    return (
      rect.width > 1 &&
      rect.height > 1 &&
      rect.bottom > 0 &&
      rect.right > 0 &&
      rect.top < innerHeight &&
      rect.left < innerWidth
    );
  }

  function isTextInput(element) {
    return (
      element instanceof Element &&
      (element.matches(TEXT_INPUT_SELECTOR) ||
        Boolean(element.closest(TEXT_INPUT_SELECTOR)))
    );
  }

  function normaliseText(value) {
    return value?.replace(/\s+/g, " ").trim() ?? "";
  }

  function elementLabel(element) {
    const labelledBy = element.getAttribute("aria-labelledby");
    const labelledText = labelledBy
      ?.split(/\s+/)
      .map(id => document.getElementById(id)?.textContent)
      .filter(Boolean)
      .join(" ");
    return normaliseText(
      element.getAttribute("aria-label") ||
        labelledText ||
        element.getAttribute("title") ||
        element.getAttribute("alt") ||
        element.getAttribute("data-tid")?.replaceAll("-", " ") ||
        element.textContent ||
        element.getAttribute("placeholder") ||
        element.getAttribute("name")
    ).slice(0, 180);
  }

  function focusElement(element) {
    if (!(element instanceof HTMLElement)) return;
    element.scrollIntoView({ block: "center", inline: "nearest" });
    if (!element.hasAttribute("tabindex") && element.tabIndex < 0) {
      element.setAttribute("tabindex", "-1");
      element.dataset.teamsVimiumTemporaryTabindex = "true";
    }
    element.focus({ preventScroll: true });
  }

  function activateElement(element) {
    if (!(element instanceof HTMLElement)) return;
    element.scrollIntoView({ block: "center", inline: "nearest" });
    if (isTextInput(element)) {
      focusElement(element);
      if (element instanceof HTMLInputElement || element instanceof HTMLTextAreaElement) {
        element.select();
      }
      setMode("insert");
      return;
    }
    element.click();
  }

  function setMode(mode) {
    state.mode = mode;
    ui.mode.textContent = mode.toUpperCase();
    ui.mode.classList.toggle("hidden", mode !== "normal" && mode !== "insert");
  }

  function clearPendingNavigation() {
    clearTimeout(state.pendingNavigationTimer);
    state.pendingNavigationKey = "";
    state.pendingNavigationTimer = 0;
  }

  function leaveOverlayMode({ preserveFind = true } = {}) {
    clearPendingNavigation();
    clearHints();
    ui.find.classList.add("hidden");
    ui.vomnibar.classList.add("hidden");
    ui.helpBackdrop.classList.add("hidden");
    if (!preserveFind) clearFind();
    setMode("normal");
  }

  function enterNormalMode() {
    leaveOverlayMode();
    const activeElement =
      document.activeElement?.shadowRoot?.activeElement ?? document.activeElement;
    if (activeElement instanceof HTMLElement) activeElement.blur();
    document.getSelection()?.removeAllRanges();
  }

  function getScrollableAncestors(element) {
    const candidates = [];
    for (let current = element; current instanceof HTMLElement; current = current.parentElement) {
      const style = getComputedStyle(current);
      if (
        /(auto|scroll|overlay)/.test(style.overflowY) &&
        current.scrollHeight > current.clientHeight + 2
      ) {
        candidates.push(current);
      }
    }
    return candidates;
  }

  function scrollableDescendant(element) {
    if (!(element instanceof HTMLElement)) return null;
    return [element, ...element.querySelectorAll("*")]
      .filter(candidate => {
        if (!(candidate instanceof HTMLElement) || !isVisible(candidate)) {
          return false;
        }
        const style = getComputedStyle(candidate);
        return (
          /(auto|scroll|overlay)/.test(style.overflowY) &&
          candidate.scrollHeight > candidate.clientHeight + 2
        );
      })
      .sort(
        (left, right) =>
          right.clientWidth * right.clientHeight -
          left.clientWidth * left.clientHeight
      )[0] ?? null;
  }

  function activeScroller() {
    if (state.activePane && document.contains(state.activePane)) {
      return (
        scrollableDescendant(state.activePane) ||
        getScrollableAncestors(state.activePane)[0] ||
        (state.activePane.scrollHeight > state.activePane.clientHeight
          ? state.activePane
          : null)
      );
    }

    const active = document.activeElement;
    const ancestor = getScrollableAncestors(active)[0];
    if (ancestor) return ancestor;

    const center = document.elementFromPoint(innerWidth / 2, innerHeight / 2);
    return (
      getScrollableAncestors(center)[0] ||
      document.scrollingElement ||
      document.documentElement
    );
  }

  function scrollByAmount(amount) {
    activeScroller()?.scrollBy({ top: amount, behavior: "smooth" });
  }

  function isChatListActive() {
    return state.activePane?.matches(
      "[data-testid='simple-collab-rail'], [data-tid='simple-collab-rail']"
    );
  }

  function chatRowKey(row) {
    return row?.getAttribute("data-fui-tree-item-value");
  }

  function focusChatRow(row) {
    if (!(row instanceof HTMLElement)) return;
    state.activeChatRow = row;
    state.activeChatKey = chatRowKey(row);
    row.focus({ preventScroll: true });
    row.scrollIntoView({ block: "nearest", inline: "nearest" });
    showPaneIndicator(row);
  }

  async function moveChatSelection(direction, distance = 1) {
    const scroller = chatRailScroller();
    if (!scroller) return;

    let rows = visibleChatRows().sort(
      (left, right) =>
        left.getBoundingClientRect().top - right.getBoundingClientRect().top
    );
    if (!rows.length) return;

    const focusedRow = document.activeElement?.closest?.(
      "[data-testid='list-item']"
    );
    let current =
      rows.find(row => chatRowKey(row) === state.activeChatKey) ||
      (rows.includes(focusedRow) && focusedRow) ||
      (rows.includes(state.activeChatRow) && state.activeChatRow) ||
      rows.find(
        row =>
          row.getAttribute("aria-selected") === "true" ||
          row.getAttribute("aria-current") === "true"
      );

    if (!current) {
      focusChatRow(direction > 0 ? rows[0] : rows.at(-1));
      return;
    }

    let index = rows.indexOf(current);
    const targetIndex = index + direction * distance;
    if (rows[targetIndex]) {
      focusChatRow(rows[targetIndex]);
      return;
    }

    const currentKey = chatRowKey(current);
    const rowHeight = Math.max(1, current.getBoundingClientRect().height);
    scroller.scrollTop += direction * distance * rowHeight;
    await nextRender();
    rows = visibleChatRows().sort(
      (left, right) =>
        left.getBoundingClientRect().top - right.getBoundingClientRect().top
    );
    index = rows.findIndex(row => chatRowKey(row) === currentKey);
    focusChatRow(
      index >= 0
        ? rows[index + direction * distance] ??
          (direction > 0 ? rows.at(-1) : rows[0])
        : direction > 0
          ? rows[0]
          : rows.at(-1)
    );
  }

  function activateChatSelection() {
    if (!isChatListActive()) return false;
    const row =
      visibleChatRows().find(
        candidate => chatRowKey(candidate) === state.activeChatKey
      ) ??
      document.activeElement?.closest?.("[data-testid='list-item']") ??
      state.activeChatRow;
    if (!(row instanceof HTMLElement) || !document.contains(row)) return false;
    activateElement(row);
    return true;
  }

  function calendarItems() {
    if (!IS_HOSTED_CALENDAR) return [];
    return [...document.querySelectorAll("[role='button'][title]")].filter(
      element => element.getAttribute("title")?.includes("\n") && isVisible(element)
    );
  }

  function calendarItemCentre(element) {
    const rect = element.getBoundingClientRect();
    return {
      x: rect.left + rect.width / 2,
      y: rect.top + rect.height / 2,
    };
  }

  function focusCalendarItem(element) {
    if (!(element instanceof HTMLElement)) return;
    state.activeCalendarItem = element;
    focusElement(element);
    showPaneIndicator(element);
  }

  function moveCalendarSelection(horizontal, vertical) {
    const items = calendarItems();
    if (!items.length) return false;

    const active =
      (items.includes(document.activeElement) && document.activeElement) ||
      (items.includes(state.activeCalendarItem) && state.activeCalendarItem);
    if (!active) {
      const viewportCentre = { x: innerWidth / 2, y: innerHeight / 2 };
      focusCalendarItem(
        items.sort((left, right) => {
          const leftCentre = calendarItemCentre(left);
          const rightCentre = calendarItemCentre(right);
          return (
            Math.hypot(
              leftCentre.x - viewportCentre.x,
              leftCentre.y - viewportCentre.y
            ) -
            Math.hypot(
              rightCentre.x - viewportCentre.x,
              rightCentre.y - viewportCentre.y
            )
          );
        })[0]
      );
      return true;
    }

    const origin = calendarItemCentre(active);
    const candidates = items
      .filter(element => element !== active)
      .map(element => {
        const centre = calendarItemCentre(element);
        const x = centre.x - origin.x;
        const y = centre.y - origin.y;
        const primary = horizontal ? x * horizontal : y * vertical;
        const cross = horizontal ? Math.abs(y) : Math.abs(x);
        return { element, primary, score: primary + cross * 1.5 };
      })
      .filter(candidate => candidate.primary > 4)
      .sort(
        (left, right) =>
          left.score - right.score || left.primary - right.primary
      );
    if (candidates[0]) focusCalendarItem(candidates[0].element);
    return true;
  }

  function activateCalendarSelection() {
    if (!IS_HOSTED_CALENDAR) return false;
    const item =
      (calendarItems().includes(document.activeElement) &&
        document.activeElement) ||
      state.activeCalendarItem;
    if (!(item instanceof HTMLElement) || !document.contains(item)) return false;
    activateElement(item);
    return true;
  }

  function navigateCalendarPeriod(direction) {
    if (!IS_HOSTED_CALENDAR) return false;
    const prefix = direction < 0 ? "go to previous " : "go to next ";
    const button = [...document.querySelectorAll("button")].find(
      element =>
        isVisible(element) &&
        normaliseText(element.getAttribute("aria-label"))
          .toLocaleLowerCase()
          .startsWith(prefix)
    );
    if (!(button instanceof HTMLElement)) return false;
    state.activeCalendarItem = null;
    button.click();
    return true;
  }

  function navigateVertically(direction, fallbackAmount, chatDistance = 1) {
    if (isChatListActive()) {
      void moveChatSelection(direction, chatDistance);
    } else {
      scrollByAmount(fallbackAmount);
    }
  }

  function paneCandidates() {
    const chatList = document.querySelector(
      "[data-testid='simple-collab-rail'], [data-tid='simple-collab-rail']"
    );
    const messageView = document.querySelector(
      "[data-tid='message-pane-layout'], [data-testid='message-pane-layout']"
    );
    const subNavigation = document.querySelector(
      "[data-tid='app-layout-area--sub-nav'], [data-testid='app-layout-area--sub-nav']"
    );
    const main = document.querySelector(
      "[data-tid='app-layout-area--main'], [data-testid='app-layout-area--main']"
    );

    return [chatList, subNavigation, isVisible(messageView) ? messageView : main]
      .filter(
        element =>
          element instanceof HTMLElement &&
          isVisible(element) &&
          (element === chatList ||
            element.getBoundingClientRect().left >
              (chatList?.getBoundingClientRect().left ?? -1)) &&
          element.getBoundingClientRect().width >= 120 &&
          element.getBoundingClientRect().height >= 160
      )
      .sort(
        (left, right) =>
          left.getBoundingClientRect().left - right.getBoundingClientRect().left
      );
  }

  function selectPane(direction) {
    const panes = paneCandidates();
    if (!panes.length) return;
    let index = panes.indexOf(state.activePane);
    if (index < 0) {
      const focused = document.activeElement;
      index = panes.findIndex(pane => pane.contains(focused));
    }
    if (index < 0) {
      index = direction < 0 ? 0 : panes.length - 1;
    } else {
      index = Math.max(0, Math.min(panes.length - 1, index + direction));
    }
    state.activePane?.removeAttribute("data-teams-vimium-active-pane");
    state.activePane = panes[index];
    state.activePane.setAttribute("data-teams-vimium-active-pane", "true");
    focusElement(state.activePane);
    showPaneIndicator(state.activePane);
  }

  function showPaneIndicator(pane) {
    const existing = root.getElementById("pane-indicator");
    existing?.remove();
    const rect = pane.getBoundingClientRect();
    const indicator = document.createElement("div");
    indicator.id = "pane-indicator";
    Object.assign(indicator.style, {
      position: "fixed",
      zIndex: "2147483645",
      left: `${Math.max(0, rect.left)}px`,
      top: `${Math.max(0, rect.top)}px`,
      width: `${Math.min(innerWidth - Math.max(0, rect.left), rect.width)}px`,
      height: `${Math.min(innerHeight - Math.max(0, rect.top), rect.height)}px`,
      border: "2px solid #5b5fc7",
      borderRadius: "4px",
      pointerEvents: "none",
      boxSizing: "border-box",
    });
    root.append(indicator);
    setTimeout(() => indicator.remove(), 550);
  }

  function focusMainTextBox() {
    const candidates = [...document.querySelectorAll(TEXT_INPUT_SELECTOR)].filter(
      isVisible
    );
    if (!candidates.length) return;

    const scored = candidates
      .map(element => {
        const label = `${elementLabel(element)} ${element.getAttribute("data-tid") ?? ""}`.toLowerCase();
        const rect = element.getBoundingClientRect();
        let score = rect.top / innerHeight;
        if (/message|compose|type a message|reply|new message|send/.test(label)) {
          score += 100;
        }
        if (element.getAttribute("contenteditable") === "true") score += 20;
        if (element.closest("[role='dialog']")) score += 10;
        return { element, score };
      })
      .sort((left, right) => right.score - left.score);

    focusElement(scored[0].element);
    setMode("insert");
  }

  function closeTopOverlay() {
    const overlays = [
      ...document.querySelectorAll(
        "[role='dialog'], [role='menu'], [role='listbox'], [aria-modal='true']"
      ),
    ].filter(isVisible);
    const overlay = overlays.at(-1);
    if (!overlay) {
      document.dispatchEvent(
        new KeyboardEvent("keydown", {
          key: "Escape",
          code: "Escape",
          bubbles: true,
          cancelable: true,
        })
      );
      return;
    }
    const closeButton = [...overlay.querySelectorAll(CLICKABLE_SELECTOR)].find(
      element =>
        /^(close|dismiss|cancel|back)$|close|dismiss/i.test(elementLabel(element))
    );
    if (closeButton instanceof HTMLElement) {
      closeButton.click();
    } else {
      overlay.dispatchEvent(
        new KeyboardEvent("keydown", {
          key: "Escape",
          code: "Escape",
          bubbles: true,
          cancelable: true,
        })
      );
    }
  }

  function textNodes() {
    const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, {
      acceptNode(node) {
        const parent = node.parentElement;
        if (
          !parent ||
          !node.nodeValue?.trim() ||
          parent.closest(`[${OWN_ATTRIBUTE}], script, style, noscript, template`) ||
          !isVisible(parent)
        ) {
          return NodeFilter.FILTER_REJECT;
        }
        return NodeFilter.FILTER_ACCEPT;
      },
    });
    const nodes = [];
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      nodes.push(node);
    }
    return nodes;
  }

  function clearFind() {
    state.find.query = "";
    state.find.ranges = [];
    state.find.currentIndex = -1;
    if (globalThis.CSS?.highlights) {
      CSS.highlights.delete(FIND_HIGHLIGHT);
      CSS.highlights.delete(CURRENT_FIND_HIGHLIGHT);
    }
    ui.findInput.value = "";
    ui.findCount.textContent = "0/0";
  }

  function updateFind(query) {
    state.find.query = query;
    state.find.ranges = [];
    state.find.currentIndex = -1;
    if (globalThis.CSS?.highlights) {
      CSS.highlights.delete(FIND_HIGHLIGHT);
      CSS.highlights.delete(CURRENT_FIND_HIGHLIGHT);
    }
    if (!query) {
      ui.findCount.textContent = "0/0";
      return;
    }

    const needle = query.toLocaleLowerCase();
    for (const node of textNodes()) {
      const haystack = node.nodeValue.toLocaleLowerCase();
      for (let index = haystack.indexOf(needle); index >= 0; index = haystack.indexOf(needle, index + Math.max(1, needle.length))) {
        const range = new Range();
        range.setStart(node, index);
        range.setEnd(node, index + needle.length);
        state.find.ranges.push(range);
      }
    }

    if (globalThis.Highlight && globalThis.CSS?.highlights) {
      CSS.highlights.set(FIND_HIGHLIGHT, new Highlight(...state.find.ranges));
    }
    ui.findCount.textContent = `0/${state.find.ranges.length}`;
  }

  function selectFind(index) {
    const { ranges } = state.find;
    if (!ranges.length) return;
    state.find.currentIndex = (index + ranges.length) % ranges.length;
    const range = ranges[state.find.currentIndex];
    if (globalThis.Highlight && globalThis.CSS?.highlights) {
      CSS.highlights.set(CURRENT_FIND_HIGHLIGHT, new Highlight(range));
    }
    ui.findCount.textContent = `${state.find.currentIndex + 1}/${ranges.length}`;
    const element = range.startContainer.parentElement;
    if (element) {
      element.scrollIntoView({ block: "center", inline: "nearest", behavior: "smooth" });
      focusElement(element);
    }
  }

  function openFind() {
    leaveOverlayMode();
    setMode("find");
    ui.find.classList.remove("hidden");
    ui.findInput.value = state.find.query;
    ui.findInput.focus();
    ui.findInput.select();
  }

  function nextFind(direction) {
    if (!state.find.ranges.length && state.find.query) {
      updateFind(state.find.query);
    }
    selectFind(
      state.find.currentIndex < 0
        ? direction > 0
          ? 0
          : state.find.ranges.length - 1
        : state.find.currentIndex + direction
    );
  }

  function hintFirstKeys() {
    const context = globalThis.__teamsVimiumHintContext ?? {
      index: 0,
      count: 1,
    };
    return [...BASE_HINT_FIRST_KEYS].filter(
      (_, keyIndex) => keyIndex % context.count === context.index
    );
  }

  function hintCode(index, useLongCodes) {
    const firstKeys = hintFirstKeys();
    if (useLongCodes) {
      const suffixCapacity = HINT_KEYS.length * HINT_KEYS.length;
      return `${
        firstKeys[Math.floor(index / suffixCapacity)]
      }${HINT_KEYS[Math.floor(index / HINT_KEYS.length) % HINT_KEYS.length]}${
        HINT_KEYS[index % HINT_KEYS.length]
      }`;
    }
    return `${firstKeys[Math.floor(index / HINT_KEYS.length)]}${
      HINT_KEYS[index % HINT_KEYS.length]
    }`;
  }

  function clickableElements(scope = document) {
    const candidates = new Set(scope.querySelectorAll(CLICKABLE_SELECTOR));
    for (const element of scope.querySelectorAll("*")) {
      if (getComputedStyle(element).cursor === "pointer") {
        candidates.add(element.closest(CLICKABLE_SELECTOR) ?? element);
      }
    }

    const normalisedCandidates = new Set();
    for (const candidate of candidates) {
      if (
        candidate.matches(
          "[role='group'], [data-conversation-folder='true']"
        )
      ) {
        continue;
      }
      const calendarDayControl = IS_HOSTED_CALENDAR
        ? candidate.closest("[data-telemetry-id='WorkPlanButton']") ??
          candidate.querySelector?.("[data-telemetry-id='WorkPlanButton']")
        : null;
      const calendarDateCell = IS_HOSTED_CALENDAR
        ? candidate.closest("[role='gridcell'][data-cell-date]")
        : null;
      const tab = candidate.closest("[role='tab']");
      const messageLink =
        candidate.matches("a[href], [role='link']") &&
        candidate.closest("[data-tid='chat-pane-item']")
          ? candidate
          : null;
      const messageReaction =
        candidate.matches(
          "[data-tid='diverse-reaction-pill-button'], [data-tid='add-reaction-picker-entry-point-button']"
        ) && candidate.closest("[data-tid='chat-pane-item']")
          ? candidate
          : null;
      const messageImage = candidate.closest(
        "[data-tid='chat-pane-item'] img[data-gallery-id]"
      );
      const message = messageLink || messageReaction || messageImage
        ? null
        : candidate.closest("[data-tid='chat-pane-item']");
      const channel = candidate.closest(
        "[role='treeitem'][data-item-type='channel']"
      );
      const contact = candidate.closest("[data-testid='list-item']");
      const quickView = candidate.closest(
        "[data-tid^='slice-list-item-'], [data-testid^='slice-list-item-']"
      );
      const folderHeader = candidate.closest(
        "[data-tid='conversation-folder-header'], [data-testid='conversation-folder-header']"
      );
      normalisedCandidates.add(
        calendarDateCell ??
          calendarDayControl ??
          tab ??
        messageLink ??
        messageReaction ??
        messageImage ??
        message ??
        channel ??
        contact ??
          quickView ??
          folderHeader?.firstElementChild ??
          candidate
      );
    }

    const seenRects = new Set();
    return [...normalisedCandidates].filter(element => {
      if (
        (scope instanceof Element && !scope.contains(element)) ||
        !isVisible(element) ||
        element.closest("[inert]")
      ) {
        return false;
      }
      const rect = element.getBoundingClientRect();
      const key = `${Math.round(rect.left)}:${Math.round(rect.top)}:${Math.round(rect.width)}:${Math.round(rect.height)}`;
      if (seenRects.has(key)) return false;
      seenRects.add(key);
      return true;
    });
  }

  function revealMessageActions(message) {
    const messageBody =
      message.querySelector("[data-tid='chat-pane-message']") ?? message;
    const rect = messageBody.getBoundingClientRect();
    const eventInit = {
      bubbles: true,
      clientX: rect.left + rect.width / 2,
      clientY: rect.top + rect.height / 2,
    };
    messageBody.dispatchEvent(
      new PointerEvent("pointerover", {
        ...eventInit,
        pointerType: "mouse",
      })
    );
    messageBody.dispatchEvent(
      new MouseEvent("mouseover", eventInit)
    );
    focusElement(messageBody);
    const menuButton = message.querySelector(
      "[data-tid='message-actions-menu-hidden-button']"
    );
    if (!menuButton) return;

    // Teams ignores untrusted clicks on this hidden control, so invoke the
    // React handler directly when its private props are available.
    const reactPropsKey = Object.keys(menuButton).find(key =>
      key.startsWith("__reactProps$")
    );
    const onClick = reactPropsKey
      ? menuButton[reactPropsKey]?.onClick
      : undefined;
    if (typeof onClick === "function") {
      onClick({
        currentTarget: menuButton,
        target: menuButton,
        preventDefault() {},
        stopPropagation() {},
        nativeEvent: {},
      });
    } else {
      menuButton.click();
    }
  }

  function hintTarget(element) {
    if (element.matches("[role='gridcell'][data-cell-date]")) {
      const day = Number(element.getAttribute("data-cell-date")?.split("-").at(-1));
      return {
        element,
        anchor: element,
        centre: true,
        calendarDay: day,
        calendarDayPrimary: !element.classList.contains(
          "fui-CalendarDayGrid__dayOutsideNavigatedMonth"
        ),
        activate: () =>
          activateElement(element.querySelector("button") ?? element),
      };
    }
    if (element.matches("[data-telemetry-id='WorkPlanButton']")) {
      return {
        element,
        anchor: element,
        centre: true,
        activate: () => activateElement(element),
      };
    }
    if (element.matches("[role='tab']")) {
      return {
        element,
        anchor: element,
        centre: true,
        activate: () => activateElement(element),
      };
    }
    if (
      IS_HOSTED_CALENDAR &&
      element.matches("[role='button'][title]") &&
      element.getAttribute("title")?.includes("\n")
    ) {
      return {
        element,
        anchor: element,
        centre: true,
        openMeetingCardHints: true,
        activate: () => activateElement(element),
      };
    }
    if (
      element.parentElement?.matches(
        "[data-tid='conversation-folder-header'], [data-testid='conversation-folder-header']"
      ) &&
      element === element.parentElement.firstElementChild
    ) {
      return {
        element,
        anchor: element,
        centre: true,
        activate: () => activateElement(element),
      };
    }
    if (element.matches("[data-tid='chat-pane-item']")) {
      const anchor =
        element.querySelector("[data-tid='chat-pane-message']") ?? element;
      return {
        element,
        anchor,
        centre: true,
        activate: () => revealMessageActions(element),
      };
    }
    if (element.matches("img[data-gallery-id]")) {
      return {
        element,
        anchor: element,
        centre: true,
        activate: () => activateElement(element),
      };
    }
    if (element.matches("[role='treeitem'][data-item-type='channel']")) {
      return {
        element,
        anchor:
          element.querySelector("[id^='title-channel-list-item-']") ?? element,
        centre: true,
        activate: () => activateElement(element),
      };
    }
    if (element.matches("[data-testid='list-item']")) {
      return {
        element,
        anchor: element,
        centre: true,
        activate: () => activateElement(element),
      };
    }
    if (
      element.matches(
        "[data-tid^='slice-list-item-'], [data-testid^='slice-list-item-']"
      )
    ) {
      return {
        element,
        anchor: element.querySelector("[role='text']") ?? element,
        centre: true,
        activate: () => activateElement(element),
      };
    }
    return {
      element,
      anchor: element,
      centre: false,
      activate: () => activateElement(element),
    };
  }

  function clearHints() {
    state.hints = [];
    state.hintInput = "";
    ui.hints.replaceChildren();
  }

  function renderHints() {
    for (const hint of state.hints) {
      const matches = hint.code.startsWith(state.hintInput);
      hint.label.classList.toggle("hidden", !matches);
      if (matches) {
        const matched = hint.code.slice(0, state.hintInput.length);
        const remaining = hint.code.slice(state.hintInput.length);
        const matchedLabel = document.createElement("span");
        matchedLabel.className = "matched";
        matchedLabel.textContent = matched;
        hint.label.replaceChildren(matchedLabel, remaining);
      }
    }
  }

  function meetingCard() {
    const viewEventButton = [...document.querySelectorAll("button")].find(
      element =>
        element.getAttribute("aria-label") === "View event" &&
        isVisible(element)
    );
    if (!viewEventButton) return null;
    for (
      let candidate = viewEventButton.parentElement;
      candidate instanceof HTMLElement;
      candidate = candidate.parentElement
    ) {
      const rect = candidate.getBoundingClientRect();
      if (
        rect.width >= 280 &&
        rect.height >= 100 &&
        candidate.querySelectorAll("button, [role='button']").length >= 5
      ) {
        return candidate;
      }
    }
    return null;
  }

  async function openMeetingCardHints(meetingTitle) {
    let previousRect;
    let stableFrames = 0;
    for (let attempt = 0; attempt < 60; attempt++) {
      await nextRender();
      const card = meetingCard();
      if (!card || !normaliseText(card.textContent).includes(meetingTitle)) {
        previousRect = undefined;
        stableFrames = 0;
        continue;
      }
      const rect = card.getBoundingClientRect();
      const currentRect = [rect.left, rect.top, rect.width, rect.height];
      const isStable =
        previousRect &&
        currentRect.every(
          (value, index) => Math.abs(value - previousRect[index]) < 0.5
        );
      stableFrames = isStable ? stableFrames + 1 : 0;
      previousRect = currentRect;
      if (stableFrames >= 3) {
        openHints(card);
        return;
      }
    }
  }

  function openHints(scope = document) {
    leaveOverlayMode();
    const targets = clickableElements(scope).map(hintTarget);
    if (!targets.length) return;
    const firstKeyCount = hintFirstKeys().length;
    const twoCharacterCapacity = firstKeyCount * HINT_KEYS.length;
    const threeCharacterCapacity =
      twoCharacterCapacity * HINT_KEYS.length;
    if (!firstKeyCount || targets.length > threeCharacterCapacity) {
      console.warn("Too many targets for Teams Vimium hints");
      return;
    }
    const useLongCodes = targets.length > twoCharacterCapacity;
    setMode("hint");
    const calendarDayIndexes = new Map();
    let generatedCodeIndex = 0;
    state.hints = targets.map(target => {
      const rect = target.anchor.getBoundingClientRect();
      const label = document.createElement("span");
      label.className = "hint";
      let code;
      if (target.calendarDay) {
        const hintContext = globalThis.__teamsVimiumHintContext ?? {
          index: 0,
          count: 1,
        };
        const windowSuffix =
          hintContext.count > 1 ? HINT_KEYS[hintContext.index] : "";
        if (target.calendarDayPrimary) {
          code =
            target.calendarDay <= 3
              ? String(target.calendarDay).padStart(2, "0")
              : String(target.calendarDay);
          code += windowSuffix;
        } else {
          const conflictIndex =
            calendarDayIndexes.get(target.calendarDay) ?? 0;
          calendarDayIndexes.set(target.calendarDay, conflictIndex + 1);
          code = `${target.calendarDay}${String.fromCharCode(
            "a".charCodeAt(0) + conflictIndex
          )}${windowSuffix}`;
        }
      } else {
        code = hintCode(generatedCodeIndex++, useLongCodes);
      }
      label.textContent = code;
      if (target.centre) {
        label.style.left = `${rect.left + rect.width / 2}px`;
        label.style.top = `${rect.top + rect.height / 2}px`;
        label.style.transform = "translate(-50%, -50%)";
      } else {
        label.style.left = `${Math.max(0, Math.min(innerWidth - 30, rect.left + Math.min(8, rect.width / 2)))}px`;
        label.style.top = `${Math.max(0, Math.min(innerHeight - 18, rect.top + Math.min(8, rect.height / 2)))}px`;
      }
      ui.hints.append(label);
      return { code, ...target, label };
    });
  }

  function handleHintKey(key) {
    if (key === "Backspace") {
      state.hintInput = state.hintInput.slice(0, -1);
      renderHints();
      return true;
    }
    if (!/^[a-z0-9]$/i.test(key)) return false;
    state.hintInput += key.toLowerCase();
    const matches = state.hints.filter(hint =>
      hint.code.startsWith(state.hintInput)
    );
    if (matches.length === 1 && matches[0].code === state.hintInput) {
      const { activate, openMeetingCardHints: shouldOpenMeetingCardHints } =
        matches[0];
      const meetingTitle = shouldOpenMeetingCardHints
        ? normaliseText(matches[0].element.getAttribute("title")?.split("\n")[0])
        : "";
      clearHints();
      setMode("normal");
      relayHintComplete();
      activate();
      if (shouldOpenMeetingCardHints) {
        void openMeetingCardHints(meetingTitle);
      }
    } else if (!matches.length) {
      clearHints();
      setMode("normal");
    } else {
      renderHints();
    }
    return true;
  }

  function entryKind(element) {
    const role = element.getAttribute("role");
    const tid = element.getAttribute("data-tid")?.toLowerCase() ?? "";
    const context = `${role ?? ""} ${tid} ${elementLabel(element)}`.toLowerCase();
    if (
      (IS_HOSTED_CALENDAR &&
        element.matches("[role='button'][title]") &&
        element.getAttribute("title")?.includes("\n")) ||
      /meeting|calendar/.test(context)
    ) {
      return "meeting";
    }
    if (/channel/.test(context)) return "channel";
    if (/chat|contact|person|conversation/.test(context)) return "chat";
    if (/filter/.test(context)) return "filter";
    if (/quick|view/.test(context)) return "quick view";
    if (role === "tab" || /nav|rail/.test(context)) return "navigation";
    if (isTextInput(element)) return "text box";
    return role || element.tagName.toLowerCase();
  }

  function nextRender() {
    return new Promise(resolve =>
      requestAnimationFrame(() => requestAnimationFrame(resolve))
    );
  }

  function visibleChatRows() {
    return [...document.querySelectorAll("[data-testid='list-item']")].filter(
      isVisible
    );
  }

  function chatRailScroller() {
    const row = visibleChatRows()[0];
    return row ? getScrollableAncestors(row)[0] : null;
  }

  async function visitVirtualizedChatRows(
    scroller,
    visitor,
    { restoreScroll = true } = {}
  ) {
    const originalScrollTop = scroller.scrollTop;
    const step = Math.max(1, Math.floor(scroller.clientHeight * 0.75));
    try {
      for (
        let scrollTop = 0;
        scrollTop <= scroller.scrollHeight;
        scrollTop += step
      ) {
        scroller.scrollTop = Math.min(
          scrollTop,
          scroller.scrollHeight - scroller.clientHeight
        );
        await nextRender();
        if (visitor(visibleChatRows()) === false) return;
        if (
          scroller.scrollTop >=
          scroller.scrollHeight - scroller.clientHeight - 1
        ) {
          return;
        }
      }
    } finally {
      if (restoreScroll) {
        scroller.scrollTop = originalScrollTop;
        await nextRender();
      }
    }
  }

  async function activateVirtualizedChat(scroller, label) {
    await visitVirtualizedChatRows(
      scroller,
      rows => {
        const target = rows.find(row => elementLabel(row) === label);
        if (!target) return true;
        activateElement(target);
        return false;
      },
      { restoreScroll: false }
    );
  }

  async function collectVirtualizedChatEntries() {
    const scroller = chatRailScroller();
    if (!scroller) return [];
    const entries = new Map();
    await visitVirtualizedChatRows(scroller, rows => {
      for (const row of rows) {
        const label = elementLabel(row);
        if (!label || label.length > 180 || entries.has(label)) continue;
        entries.set(label, {
          label,
          kind: "chat",
          searchText: `${label} chat`.toLocaleLowerCase(),
          activate: () => activateVirtualizedChat(scroller, label),
        });
      }
    });
    return [...entries.values()];
  }

  function collectCalendarEntries() {
    const entries = [];
    const seen = new Set();
    for (const element of document.querySelectorAll("[role='button'][title]")) {
      if (!element.getAttribute("title")?.includes("\n")) continue;
      const label = elementLabel(element);
      if (!label || seen.has(label)) continue;
      seen.add(label);
      entries.push({
        element,
        label,
        kind: "meeting",
        searchText: `${label} meeting calendar`.toLocaleLowerCase(),
        activate: () => activateElement(element),
      });
    }
    return entries;
  }

  function collectVomnibarEntries(additionalEntries = []) {
    const entries = [];
    const seen = new Set();
    for (const entry of additionalEntries) {
      seen.add(`${entry.kind}:${entry.label.toLocaleLowerCase()}`);
      entries.push(entry);
    }
    for (const element of clickableElements()) {
      const label = elementLabel(element);
      if (!label || label.length > 180) continue;
      const kind = entryKind(element);
      const key = `${kind}:${label.toLocaleLowerCase()}`;
      if (seen.has(key)) continue;
      seen.add(key);
      entries.push({
        element,
        label,
        kind,
        searchText: `${label} ${kind}`.toLocaleLowerCase(),
        activate: () => activateElement(element),
      });
    }
    return entries.sort((left, right) => {
      if (!left.element || !right.element) {
        return left.kind.localeCompare(right.kind) ||
          left.label.localeCompare(right.label);
      }
      const leftRect = left.element.getBoundingClientRect();
      const rightRect = right.element.getBoundingClientRect();
      return leftRect.top - rightRect.top || leftRect.left - rightRect.left;
    });
  }

  function fuzzyScore(query, text) {
    if (!query) return 1;
    const exactIndex = text.indexOf(query);
    if (exactIndex >= 0) return 1000 - exactIndex * 2 - text.length / 100;
    let queryIndex = 0;
    let score = 0;
    let lastMatch = -2;
    for (let index = 0; index < text.length && queryIndex < query.length; index++) {
      if (text[index] !== query[queryIndex]) continue;
      score += index === lastMatch + 1 ? 8 : 2;
      lastMatch = index;
      queryIndex++;
    }
    return queryIndex === query.length ? score : -1;
  }

  function renderVomnibar() {
    const query = ui.vomnibarInput.value.trim().toLocaleLowerCase();
    state.vomnibarResults = state.vomnibarEntries
      .map(entry => ({ entry, score: fuzzyScore(query, entry.searchText) }))
      .filter(result => result.score >= 0)
      .sort((left, right) => right.score - left.score)
      .map(result => result.entry);
    state.vomnibarIndex = Math.max(
      0,
      Math.min(state.vomnibarIndex, state.vomnibarResults.length - 1)
    );
    ui.vomnibarCount.textContent = String(state.vomnibarResults.length);
    ui.vomnibarList.replaceChildren(
      ...state.vomnibarResults.map((entry, index) => {
        const row = document.createElement("div");
        row.className = `vomnibar-item${index === state.vomnibarIndex ? " selected" : ""}`;
        row.dataset.index = String(index);
        const label = document.createElement("span");
        label.className = "vomnibar-label";
        label.textContent = entry.label;
        const kind = document.createElement("span");
        kind.className = "vomnibar-kind";
        kind.textContent = entry.kind;
        row.append(label, kind);
        return row;
      })
    );
    ui.vomnibarList
      .querySelector(".selected")
      ?.scrollIntoView({ block: "nearest" });
  }

  async function openVomnibar() {
    leaveOverlayMode();
    const request = ++state.vomnibarRequest;
    state.vomnibarIndex = 0;
    setMode("vomnibar");
    ui.vomnibar.classList.remove("hidden");
    ui.vomnibarInput.value = "";
    ui.vomnibarCount.textContent = "Loading…";
    ui.vomnibarList.replaceChildren();
    ui.vomnibarInput.focus();
    const additionalEntries = IS_HOSTED_CALENDAR
      ? collectCalendarEntries()
      : await collectVirtualizedChatEntries();
    if (request !== state.vomnibarRequest || state.mode !== "vomnibar") return;
    state.vomnibarEntries = IS_HOSTED_CALENDAR
      ? additionalEntries
      : collectVomnibarEntries(additionalEntries);
    renderVomnibar();
    ui.vomnibarInput.focus();
  }

  function selectVomnibarEntry() {
    const entry = state.vomnibarResults[state.vomnibarIndex];
    if (!entry) return;
    leaveOverlayMode();
    entry.activate();
  }

  function openHelp() {
    leaveOverlayMode();
    setMode("help");
    ui.helpBackdrop.classList.remove("hidden");
  }

  function handleNavigationSequence(key) {
    if (key !== "[" && key !== "]") {
      clearPendingNavigation();
      return false;
    }
    if (state.pendingNavigationKey === key) {
      clearPendingNavigation();
      if (!navigateCalendarPeriod(key === "[" ? -1 : 1)) {
        if (key === "[") history.back();
        else history.forward();
      }
      return true;
    }
    clearPendingNavigation();
    state.pendingNavigationKey = key;
    state.pendingNavigationTimer = setTimeout(
      clearPendingNavigation,
      NAVIGATION_SEQUENCE_TIMEOUT
    );
    return true;
  }

  function prevent(event) {
    event.preventDefault();
    event.stopImmediatePropagation();
  }

  function relayKey(event) {
    if (
      state.relayedKey ||
      typeof globalThis.__teamsVimiumRelay !== "function"
    ) {
      return;
    }
    globalThis.__teamsVimiumRelay(
      JSON.stringify({
        type: "key",
        key: event.key,
        ctrlKey: event.ctrlKey,
        altKey: event.altKey,
        metaKey: event.metaKey,
        shiftKey: event.shiftKey,
      })
    );
  }

  function relayHintComplete() {
    if (typeof globalThis.__teamsVimiumRelay !== "function") return;
    globalThis.__teamsVimiumRelay(
      JSON.stringify({ type: "hint-complete" })
    );
  }

  function relayCalendarVomnibar() {
    if (typeof globalThis.__teamsVimiumRelay !== "function") return;
    globalThis.__teamsVimiumRelay(
      JSON.stringify({ type: "calendar-vomnibar" })
    );
  }

  function hostsVisibleCalendar() {
    return [...document.querySelectorAll("iframe")].some(iframe => {
      const rect = iframe.getBoundingClientRect();
      return (
        iframe.src.startsWith("https://outlook.office.com/hosted/calendar/") &&
        rect.width > 0 &&
        rect.height > 0
      );
    });
  }

  function openContextVomnibar() {
    if (!IS_HOSTED_CALENDAR && hostsVisibleCalendar()) {
      relayCalendarVomnibar();
      return;
    }
    void openVomnibar();
  }

  function handleGlobalKeydown(event) {
    if (event.defaultPrevented || event.isComposing) return;
    const key = event.key;

    if (key === "Escape") {
      relayKey(event);
      state.hintRelayKeysRemaining = 0;
      prevent(event);
      enterNormalMode();
      return;
    }

    if (
      event.ctrlKey &&
      !event.altKey &&
      !event.metaKey &&
      !event.shiftKey &&
      key.toLowerCase() === "f"
    ) {
      state.hintRelayKeysRemaining = 3;
      relayKey(event);
      prevent(event);
      enterNormalMode();
      openHints();
      return;
    }

    if (state.mode !== "hint" && state.hintRelayKeysRemaining > 0) {
      relayKey(event);
      if (/^[a-z0-9]$/i.test(key)) {
        state.hintRelayKeysRemaining--;
      }
      prevent(event);
      return;
    }

    if (state.mode === "find") {
      if (key === "Enter") {
        prevent(event);
        nextFind(event.shiftKey ? -1 : 1);
        ui.find.classList.add("hidden");
        setMode("normal");
      }
      return;
    }

    if (state.mode === "hint") {
      if (state.hintRelayKeysRemaining > 0) {
        relayKey(event);
        if (/^[a-z0-9]$/i.test(key)) {
          state.hintRelayKeysRemaining--;
        }
      }
      prevent(event);
      handleHintKey(key);
      return;
    }

    if (state.mode === "vomnibar") {
      if (key === "ArrowDown" || (event.ctrlKey && key.toLowerCase() === "n")) {
        prevent(event);
        state.vomnibarIndex = Math.min(
          state.vomnibarResults.length - 1,
          state.vomnibarIndex + 1
        );
        renderVomnibar();
      } else if (
        key === "ArrowUp" ||
        (event.ctrlKey && key.toLowerCase() === "p")
      ) {
        prevent(event);
        state.vomnibarIndex = Math.max(0, state.vomnibarIndex - 1);
        renderVomnibar();
      } else if (key === "Enter") {
        prevent(event);
        selectVomnibarEntry();
      }
      return;
    }

    if (state.mode === "help") return;
    if (isTextInput(event.target) || state.mode === "insert") return;
    if (event.metaKey || event.altKey) return;

    if (key.toLowerCase() === "f") {
      state.hintRelayKeysRemaining = 3;
      relayKey(event);
    }

    if (handleNavigationSequence(key)) {
      prevent(event);
      return;
    }

    if (
      key === "Enter" &&
      (activateCalendarSelection() || activateChatSelection())
    ) {
      prevent(event);
      return;
    }

    const lowerKey = key.toLowerCase();
    if (event.ctrlKey) {
      if (lowerKey === "e") {
        prevent(event);
        scrollByAmount(SCROLL_STEP);
      } else if (lowerKey === "y") {
        prevent(event);
        scrollByAmount(-SCROLL_STEP);
      }
      return;
    }

    const actions = {
      j: () =>
        IS_HOSTED_CALENDAR
          ? moveCalendarSelection(0, 1)
          : navigateVertically(1, SCROLL_STEP),
      k: () =>
        IS_HOSTED_CALENDAR
          ? moveCalendarSelection(0, -1)
          : navigateVertically(-1, -SCROLL_STEP),
      d: () =>
        navigateVertically(
          1,
          innerHeight / 2,
          Math.max(1, Math.floor(visibleChatRows().length / 2))
        ),
      u: () =>
        navigateVertically(
          -1,
          -innerHeight / 2,
          Math.max(1, Math.floor(visibleChatRows().length / 2))
        ),
      h: () =>
        IS_HOSTED_CALENDAR
          ? moveCalendarSelection(-1, 0)
          : selectPane(-1),
      l: () =>
        IS_HOSTED_CALENDAR
          ? moveCalendarSelection(1, 0)
          : selectPane(1),
      x: closeTopOverlay,
      i: focusMainTextBox,
      "/": openFind,
      n: () => nextFind(1),
      N: () => nextFind(-1),
      f: openHints,
      o: openContextVomnibar,
      "?": openHelp,
    };
    const action = actions[key] ?? actions[lowerKey];
    if (action) {
      prevent(event);
      action();
    }
  }

  function handleFocusIn(event) {
    if (
      isTextInput(event.target) &&
      !root.contains(event.target) &&
      state.mode === "normal"
    ) {
      setMode("insert");
    }
  }

  function handleFocusOut() {
    queueMicrotask(() => {
      const active = document.activeElement;
      if (state.mode === "insert" && !isTextInput(active)) setMode("normal");
    });
  }

  function installFindHighlightStyles() {
    const style = document.createElement("style");
    style.setAttribute(OWN_ATTRIBUTE, "highlight-style");
    style.textContent = `
      ::highlight(${FIND_HIGHLIGHT}) {
        color: #111;
        background: #ffe16b;
      }
      ::highlight(${CURRENT_FIND_HIGHLIGHT}) {
        color: white;
        background: #c239b3;
      }
    `;
    document.head.append(style);
    cleanupTasks.push(() => style.remove());
  }

  listen(document, "keydown", handleGlobalKeydown, true);
  listen(document, "focusin", handleFocusIn, true);
  listen(document, "focusout", handleFocusOut, true);
  listen(ui.findInput, "input", () => updateFind(ui.findInput.value));
  listen(ui.vomnibarInput, "input", () => {
    state.vomnibarIndex = 0;
    renderVomnibar();
  });
  listen(ui.vomnibarList, "mousemove", event => {
    const item = event.target.closest?.(".vomnibar-item");
    if (!item) return;
    state.vomnibarIndex = Number(item.dataset.index);
    renderVomnibar();
  });
  listen(ui.vomnibarList, "click", event => {
    const item = event.target.closest?.(".vomnibar-item");
    if (!item) return;
    state.vomnibarIndex = Number(item.dataset.index);
    selectVomnibarEntry();
  });
  listen(ui.helpBackdrop, "click", event => {
    if (event.target === ui.helpBackdrop) leaveOverlayMode();
  });
  listen(window, "blur", () => {
    if (state.hintRelayKeysRemaining === 0) clearHints();
  });
  listen(window, "resize", () => {
    if (state.mode === "hint") openHints();
  });

  installFindHighlightStyles();
  setMode("normal");

  globalThis.__teamsVimium = {
    receiveRelayedKey(payload) {
      const eventData = JSON.parse(payload);
      if (eventData.type === "hint-complete") {
        state.hintRelayKeysRemaining = 0;
        if (state.mode === "hint") leaveOverlayMode();
        return;
      }
      if (eventData.type === "calendar-vomnibar") {
        if (IS_HOSTED_CALENDAR) void openVomnibar();
        return;
      }
      state.relayedKey = true;
      try {
        handleGlobalKeydown(
          new KeyboardEvent("keydown", {
            ...eventData,
            bubbles: true,
            cancelable: true,
          })
        );
      } finally {
        state.relayedKey = false;
      }
    },
    destroy() {
      clearFind();
      clearHints();
      clearPendingNavigation();
      state.activePane?.removeAttribute("data-teams-vimium-active-pane");
      for (const element of document.querySelectorAll(
        "[data-teams-vimium-temporary-tabindex='true']"
      )) {
        element.removeAttribute("tabindex");
        delete element.dataset.teamsVimiumTemporaryTabindex;
      }
      for (const cleanup of cleanupTasks.reverse()) cleanup();
      rootHost.remove();
      delete globalThis.__teamsVimium;
    },
  };
})();
