// ==UserScript==
// @name         Meeting meter
// @match        https://teams.microsoft.com/v2/*
// @match        https://teams.cloud.microsoft/v2/*
// @match        https://local.teams.office.com/v2/*
// @run-at       document-idle
// ==/UserScript==

(() => {
  globalThis.__teamsMeetingMeter?.destroy();

  const METER_ID = "teams-meeting-meter";
  const PARTICIPANT_COST_USD = 200_000;
  const WORKING_HOURS_PER_YEAR = 2_080;
  const COST_PER_SECOND =
    PARTICIPANT_COST_USD / (WORKING_HOURS_PER_YEAR * 60 * 60);

  let totalCost = 0;
  let lastTick = performance.now();
  let currentParticipantCount = 0;
  let meter;
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

  function updateMeter() {
    const now = performance.now();
    totalCost +=
      ((now - lastTick) / 1000) *
      currentParticipantCount *
      COST_PER_SECOND;
    lastTick = now;
    currentParticipantCount = parseParticipantCount();
    if (!meter) return;
    meter.textContent = formatCost(totalCost);
    meter.title = `${currentParticipantCount} participant${
      currentParticipantCount === 1 ? "" : "s"
    } at $${PARTICIPANT_COST_USD.toLocaleString("en-US")} fully loaded annual cost each`;
  }

  function installMeter() {
    const duration = document.querySelector("span[data-tid='call-duration']");
    if (!duration?.parentElement) return;
    if (document.getElementById(METER_ID)) return;

    meter = document.createElement("span");
    meter.id = METER_ID;
    meter.setAttribute("aria-label", "Estimated meeting cost");
    Object.assign(meter.style, {
      color: "var(--colorNeutralForeground2, #616161)",
      font: "inherit",
      marginInlineStart: "8px",
      whiteSpace: "nowrap",
    });
    duration.parentElement.append(meter);
    currentParticipantCount = parseParticipantCount();
    totalCost =
      parseDuration(duration.textContent ?? "") *
      currentParticipantCount *
      COST_PER_SECOND;
    updateMeter();
  }

  installMeter();
  timer = window.setInterval(updateMeter, 1000);
  observer = new MutationObserver(() => {
    installMeter();
    currentParticipantCount = parseParticipantCount();
  });
  observer.observe(document.documentElement, { childList: true, subtree: true });

  globalThis.__teamsMeetingMeter = {
    destroy() {
      observer?.disconnect();
      window.clearInterval(timer);
      meter?.remove();
      delete globalThis.__teamsMeetingMeter;
    },
  };
})();
