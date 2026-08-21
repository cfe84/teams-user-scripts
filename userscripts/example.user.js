// ==UserScript==
// @name         Teams userscript proof of concept
// @match        https://teams.microsoft.com/*
// @match        https://teams.cloud.microsoft/*
// @match        https://local.teams.office.com/*
// @run-at       document-idle
// ==/UserScript==

const marker = document.createElement("div");
marker.textContent = "Userscript loaded";
Object.assign(marker.style, {
  position: "fixed",
  right: "12px",
  bottom: "12px",
  zIndex: "2147483647",
  padding: "6px 10px",
  borderRadius: "4px",
  color: "white",
  background: "#5b5fc7",
  font: "12px sans-serif",
});
document.body.append(marker);
console.log("YEEYEYEYYEYEYEYYE");
