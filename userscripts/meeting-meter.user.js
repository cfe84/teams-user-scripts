// ==UserScript==
// @name         Meeting meter
// @version      1.1.2
// @match        https://teams.microsoft.com/v2/*
// @match        https://teams.cloud.microsoft/v2/*
// @match        https://local.teams.office.com/v2/*
// @run-at       document-idle
// ==/UserScript==

(() => {
  globalThis.__teamsMeetingMeter?.destroy();

  const METER_ID = "teams-meeting-meter";
  const DEFAULT_EMPLOYEE_COST = 148;
  const STORAGE_KEY = "teams.meeting-meter.settings";
  const MEETING_STATE_KEY = "teams.meeting-meter.meeting";

  let employeeCost = readNumber(STORAGE_KEY, DEFAULT_EMPLOYEE_COST);
  let missedParticipantSeconds = 0;
  let actualParticipantSeconds = 0;
  let missedStartTime = new Date();
  let missedParticipantCount = 0;
  let currentParticipantCount = 0;
  let lastTick = performance.now();
  let activeMeetingKey;
  let meterRoot;
  let timeDisplay;
  let costDisplay;
  let costPerMinuteDisplay;
  let observer;
  let timer;
  let modal;

  function readNumber(key, fallback) {
    const storedValue = localStorage.getItem(key);
    if (storedValue === null) return fallback;
    const value = Number(storedValue);
    return Number.isFinite(value) && value >= 0 ? value : fallback;
  }

  function formatCost(value, fractionDigits = 2) {
    return new Intl.NumberFormat("en-US", {
      currency: "USD",
      maximumFractionDigits: fractionDigits,
      minimumFractionDigits: fractionDigits,
      style: "currency",
    }).format(value);
  }

  function parseDuration(value) {
    const parts = value.trim().split(":").map(Number);
    if (parts.some(Number.isNaN)) return 0;
    return parts.length === 2
      ? parts[0] * 60 + parts[1]
      : parts.length === 3
        ? parts[0] * 60 * 60 + parts[1] * 60 + parts[2]
        : 0;
  }

  function formatDuration(totalSeconds) {
    const roundedSeconds = Math.max(0, Math.floor(totalSeconds));
    const hours = Math.floor(roundedSeconds / 60 / 60);
    const minutes = Math.floor((roundedSeconds % (60 * 60)) / 60);
    const seconds = roundedSeconds % 60;
    return [hours, minutes, seconds]
      .map(value => value.toString().padStart(2, "0"))
      .join(":");
  }

  function dateTimeInputValue(date) {
    const local = new Date(date.getTime() - date.getTimezoneOffset() * 60 * 1000);
    return local.toISOString().slice(0, 16);
  }

  function parseParticipantCount() {
    const rosterButton = document.querySelector(
      "button[data-inp='roster-button']"
    );
    if (!rosterButton) return 0;
    const ariaLabel = rosterButton.getAttribute("aria-label") ?? "";
    const ariaMatch = ariaLabel.match(/(\d+)\s+participants?/i);
    if (ariaMatch) return Number(ariaMatch[1]);
    const tile = rosterButton.querySelector("[data-tid='roster-button-tile']");
    const tileMatch = tile?.textContent?.match(/\d+/);
    return tileMatch ? Number(tileMatch[0]) : 0;
  }

  function meetingStateKey(startTime) {
    const meetingStart = new Date(startTime);
    meetingStart.setSeconds(0, 0);
    return `${MEETING_STATE_KEY}:${location.pathname}${location.search}:${meetingStart.toISOString()}`;
  }

  function saveMeetingState() {
    localStorage.setItem(
      activeMeetingKey,
      JSON.stringify({
        actualParticipantSeconds,
        missedParticipantCount,
        missedParticipantSeconds,
        missedStartTime: missedStartTime.toISOString(),
      })
    );
  }

  function loadMeetingState(startTime, durationSeconds, participantCount) {
    activeMeetingKey = meetingStateKey(startTime);
    missedStartTime = startTime;
    missedParticipantCount = participantCount;
    missedParticipantSeconds = durationSeconds * participantCount;
    actualParticipantSeconds = 0;

    const saved = localStorage.getItem(meetingStateKey(startTime));
    if (!saved) {
      saveMeetingState();
      return;
    }
    try {
      const state = JSON.parse(saved);
      if (Number.isFinite(state.actualParticipantSeconds)) {
        actualParticipantSeconds = Math.max(0, state.actualParticipantSeconds);
      }
      if (Number.isFinite(state.missedParticipantCount)) {
        missedParticipantCount = Math.max(0, state.missedParticipantCount);
      }
      if (Number.isFinite(state.missedParticipantSeconds)) {
        missedParticipantSeconds = Math.max(0, state.missedParticipantSeconds);
      }
      if (state.missedStartTime) {
        missedStartTime = new Date(state.missedStartTime);
      }
    } catch {
      saveMeetingState();
    }
  }

  function totalParticipantSeconds() {
    return missedParticipantSeconds + actualParticipantSeconds;
  }

  function totalCost() {
    return totalParticipantSeconds() * employeeCost / (60 * 60);
  }

  function updateMeter() {
    const now = performance.now();
    actualParticipantSeconds +=
      ((now - lastTick) / 1000) * currentParticipantCount;
    lastTick = now;
    currentParticipantCount = parseParticipantCount();
    saveMeetingState();
    if (!meterRoot) return;
    timeDisplay.textContent = formatDuration(totalParticipantSeconds());
    const cost = totalCost();
    costDisplay.textContent = formatCost(cost);
    const currentCostPerMinute = currentParticipantCount * employeeCost / 60;
    costPerMinuteDisplay.textContent = `(${formatCost(
      currentCostPerMinute,
      1
    )}/min)`;
    meterRoot.title = `${currentParticipantCount} participant${
      currentParticipantCount === 1 ? "" : "s"
    } currently counted at ${formatCost(employeeCost)}/hour each`;
  }

  function closeModal() {
    modal?.remove();
    modal = undefined;
  }

  function openModal(includeMissedFields = true) {
    if (modal) return;

    modal = document.createElement("div");
    modal.setAttribute("role", "presentation");
    Object.assign(modal.style, {
      alignItems: "center",
      background: "rgba(0, 0, 0, 0.35)",
      display: "flex",
      inset: "0",
      justifyContent: "center",
      position: "fixed",
      zIndex: "100000",
    });

    const dialog = document.createElement("div");
    dialog.setAttribute("aria-labelledby", "teams-meeting-meter-title");
    dialog.setAttribute("role", "dialog");
    Object.assign(dialog.style, {
      background: "var(--colorNeutralBackground1, #fff)",
      borderRadius: "8px",
      boxShadow: "0 8px 32px rgba(0, 0, 0, 0.25)",
      color: "var(--colorNeutralForeground1, #242424)",
      maxWidth: "min(420px, calc(100vw - 32px))",
      padding: "24px",
      width: "420px",
    });

    const title = document.createElement("h2");
    title.id = "teams-meeting-meter-title";
    title.textContent = includeMissedFields
      ? "Meeting cost"
      : "Meeting Meter settings";
    title.style.margin = "0 0 12px";
    title.style.fontSize = "20px";

    const explanation = document.createElement("p");
    explanation.textContent =
      includeMissedFields
        ? "Cost is estimated from participant-seconds. The missed portion assumes the configured number of participants were present from the selected start time; the live portion counts participants every second after you joined."
        : "Set the average employee cost used to calculate meeting cost and the current cost per minute.";
    Object.assign(explanation.style, {
      fontSize: "14px",
      lineHeight: "1.4",
      margin: "0 0 20px",
    });

    const form = document.createElement("form");
    const fields = [
      ["Average employee cost to company (USD/hour)", "number", employeeCost, "0.01"],
    ];
    if (includeMissedFields) {
      fields.push(
        ["Missed portion meeting start", "datetime-local", dateTimeInputValue(missedStartTime), undefined],
        ["Missed portion participants", "number", missedParticipantCount, "1"]
      );
    }
    const inputs = fields.map(([labelText, type, value, step]) => {
      const label = document.createElement("label");
      label.textContent = labelText;
      Object.assign(label.style, {
        display: "block",
        fontSize: "14px",
        fontWeight: "600",
        marginBottom: "14px",
      });
      const input = document.createElement("input");
      input.type = type;
      input.value = String(value);
      if (step) input.step = step;
      input.min = "0";
      Object.assign(input.style, {
        boxSizing: "border-box",
        display: "block",
        font: "inherit",
        marginTop: "6px",
        padding: "7px",
        width: "100%",
      });
      label.append(input);
      form.append(label);
      return input;
    });

    const actions = document.createElement("div");
    Object.assign(actions.style, {
      display: "flex",
      gap: "8px",
      justifyContent: "flex-end",
      marginTop: "8px",
    });
    const cancel = document.createElement("button");
    cancel.type = "button";
    cancel.textContent = "Cancel";
    const save = document.createElement("button");
    save.type = "submit";
    save.textContent = "Save";
    actions.append(cancel, save);
    form.append(actions);
    dialog.append(title, explanation, form);
    modal.append(dialog);
    document.body.append(modal);

    cancel.addEventListener("click", closeModal);
    modal.addEventListener("click", event => {
      if (event.target === modal) closeModal();
    });
    form.addEventListener("submit", event => {
      event.preventDefault();
      const nextEmployeeCost = Number(inputs[0].value);
      if (!Number.isFinite(nextEmployeeCost) || nextEmployeeCost < 0) {
        return;
      }
      employeeCost = nextEmployeeCost;
      if (includeMissedFields) {
        const nextStart = new Date(inputs[1].value);
        const nextParticipants = Number(inputs[2].value);
        if (
          !Number.isFinite(nextParticipants) ||
          nextParticipants < 0 ||
          Number.isNaN(nextStart.getTime())
        ) {
          return;
        }
        missedParticipantCount = nextParticipants;
        missedStartTime = nextStart;
        missedParticipantSeconds =
          Math.max(0, (Date.now() - nextStart.getTime()) / 1000) *
          nextParticipants;
      }
      localStorage.setItem(STORAGE_KEY, String(employeeCost));
      saveMeetingState();
      updateMeter();
      closeModal();
    });
    inputs[0].focus();
  }

  function installMeter() {
    const duration = document.querySelector("span[data-tid='call-duration']");
    if (!duration?.parentElement) return;
    if (document.getElementById(METER_ID)) return;

    const meetingDurationSeconds = parseDuration(duration.textContent ?? "");
    currentParticipantCount = parseParticipantCount();
    const inferredStart = new Date(Date.now() - meetingDurationSeconds * 1000);
    loadMeetingState(inferredStart, meetingDurationSeconds, currentParticipantCount);

    meterRoot = document.createElement("span");
    meterRoot.id = METER_ID;
    meterRoot.setAttribute("aria-label", "Estimated meeting cost and person time");
    Object.assign(meterRoot.style, {
      alignItems: "center",
      display: "inline-flex",
      gap: "0",
      marginInlineStart: "10px",
      verticalAlign: "middle",
    });

    const labelDisplay = document.createElement("span");
    labelDisplay.textContent = "Meeting cost:";
    labelDisplay.setAttribute("aria-hidden", "true");
    Object.assign(labelDisplay.style, {
      color: "var(--colorNeutralForeground2, #616161)",
      font: "inherit",
      whiteSpace: "nowrap",
    });

    timeDisplay = document.createElement("span");
    timeDisplay.className = "teams-meeting-meter-time";
    timeDisplay.setAttribute("aria-label", "Person time");
    Object.assign(timeDisplay.style, {
      color: "var(--colorNeutralForeground2, #616161)",
      font: "inherit",
      fontSize: "14px",
      fontVariantNumeric: "tabular-nums",
      marginInlineStart: "6px",
      textAlign: "right",
      width: "72px",
      whiteSpace: "nowrap",
    });

    costDisplay = document.createElement("button");
    costDisplay.className = "teams-meeting-meter-cost";
    costDisplay.type = "button";
    costDisplay.setAttribute("aria-label", "Estimated meeting cost. Open meeting cost settings");
    Object.assign(costDisplay.style, {
      background: "none",
      border: "0",
      color: "var(--colorNeutralForeground2, #616161)",
      cursor: "pointer",
      font: "inherit",
      fontVariantNumeric: "tabular-nums",
      marginInlineStart: "20px",
      padding: "0",
      whiteSpace: "nowrap",
    });
    costDisplay.addEventListener("click", openModal);

    costPerMinuteDisplay = document.createElement("span");
    costPerMinuteDisplay.setAttribute("aria-label", "Estimated cost per minute");
    costPerMinuteDisplay.textContent = "($0.0/min)";
    Object.assign(costPerMinuteDisplay.style, {
      color: "var(--colorNeutralForeground2, #616161)",
      font: "inherit",
      fontVariantNumeric: "tabular-nums",
      marginInlineStart: "6px",
      whiteSpace: "nowrap",
    });

    meterRoot.append(
      labelDisplay,
      timeDisplay,
      costDisplay,
      costPerMinuteDisplay
    );
    duration.parentElement.append(meterRoot);
    updateMeter();
  }

  installMeter();
  timer = window.setInterval(updateMeter, 100);
  observer = new MutationObserver(installMeter);
  observer.observe(document.documentElement, { childList: true, subtree: true });

  const settings = () => openModal(false);
  const settingsRegistry = (globalThis.__teamsUserscriptSettings ??= {});
  settingsRegistry["Meeting meter"] = settings;

  globalThis.__teamsMeetingMeter = {
    destroy() {
      observer?.disconnect();
      window.clearInterval(timer);
      closeModal();
      meterRoot?.remove();
      if (settingsRegistry["Meeting meter"] === settings) {
        delete settingsRegistry["Meeting meter"];
      }
      delete globalThis.__teamsMeetingMeter;
    },
  };
})();
