// ==UserScript==
// @name         Meeting meter
// @version      1.0.0
// @match        https://teams.microsoft.com/v2/*
// @match        https://teams.cloud.microsoft/v2/*
// @match        https://local.teams.office.com/v2/*
// @run-at       document-idle
// ==/UserScript==

(() => {
  globalThis.__teamsMeetingMeter?.destroy();

  const METER_ID = "teams-meeting-meter";
  const PARTICIPANT_COST_USD_PER_HOUR = 148;
  const COST_PER_SECOND = PARTICIPANT_COST_USD_PER_HOUR / (60 * 60);

  let totalCost = 0;
  let totalParticipantSeconds = 0;
  let lastTick = performance.now();
  let currentParticipantCount = 0;
  let meetingDurationSeconds = 0;
  let meterRoot;
  let timeDisplay;
  let costDisplay;
  let observer;
  let timer;

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

  function formatCost(value) {
    return new Intl.NumberFormat("en-US", {
      currency: "USD",
      maximumFractionDigits: 2,
      minimumFractionDigits: 2,
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

  function updateMeter() {
    const now = performance.now();
    totalCost +=
      ((now - lastTick) / 1000) *
      currentParticipantCount *
      COST_PER_SECOND;
    totalParticipantSeconds +=
      ((now - lastTick) / 1000) * currentParticipantCount;
    lastTick = now;
    currentParticipantCount = parseParticipantCount();
    if (!meterRoot) return;
    timeDisplay.textContent = formatDuration(totalParticipantSeconds);
    costDisplay.textContent = formatCost(totalCost);
    meterRoot.title = `${currentParticipantCount} participant${
      currentParticipantCount === 1 ? "" : "s"
    } at $${PARTICIPANT_COST_USD_PER_HOUR}/hour fully loaded cost each`;
  }

  function installMeter() {
    const duration = document.querySelector("span[data-tid='call-duration']");
    if (!duration?.parentElement) return;
    if (document.getElementById(METER_ID)) return;

    meetingDurationSeconds = parseDuration(duration.textContent ?? "");
    currentParticipantCount = parseParticipantCount();
    totalParticipantSeconds =
      meetingDurationSeconds * currentParticipantCount;
    totalCost =
      meetingDurationSeconds * currentParticipantCount * COST_PER_SECOND;

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
      flex: "0 0 60px",
      marginInlineStart: "4px",
      width: "60px",
      whiteSpace: "nowrap",
    });

    costDisplay = document.createElement("span");
    costDisplay.className = "teams-meeting-meter-cost";
    costDisplay.setAttribute("aria-label", "Estimated meeting cost");
    Object.assign(costDisplay.style, {
      color: "var(--colorNeutralForeground2, #616161)",
      font: "inherit",
      fontVariantNumeric: "tabular-nums",
      marginInlineStart: "12px",
      whiteSpace: "nowrap",
    });

    meterRoot.append(labelDisplay, timeDisplay, costDisplay);
    duration.parentElement.append(meterRoot);
    updateMeter();
  }

  installMeter();
  timer = window.setInterval(updateMeter, 100);
  observer = new MutationObserver(() => {
    installMeter();
  });
  observer.observe(document.documentElement, { childList: true, subtree: true });

  globalThis.__teamsMeetingMeter = {
    destroy() {
      observer?.disconnect();
      window.clearInterval(timer);
      meterRoot?.remove();
      delete globalThis.__teamsMeetingMeter;
    },
  };
})();
