function isDevModeAuthStorage() {
  return localStorage.getItem("lspd_devmode_active") === "1";
}

function storedAuthToken() {
  return isDevModeAuthStorage() ? sessionStorage.getItem("lspd_token_dev") : localStorage.getItem("lspd_token");
}

function storeAuthToken(token) {
  if (isDevModeAuthStorage()) {
    sessionStorage.setItem("lspd_token_dev", token);
    localStorage.removeItem("lspd_token");
  } else {
    localStorage.setItem("lspd_token", token);
    sessionStorage.removeItem("lspd_token_dev");
  }
}

function clearAuthToken() {
  if (isDevModeAuthStorage()) sessionStorage.removeItem("lspd_token_dev");
  else localStorage.removeItem("lspd_token");
}

function installInspectGuard() {
  const blocker = $("#inspectBlocker");
  let lastReportAt = 0;
  const reportAttempt = (reason) => {
    const now = Date.now();
    if (now - lastReportAt < 2500) return;
    lastReportAt = now;
    fetch("/api/security/inspect-attempt", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        ...(state.token ? { Authorization: `Bearer ${state.token}` } : {})
      },
      body: JSON.stringify({ reason, page: state.page || document.title })
    }).catch(() => {});
  };
  const showBlocker = (reason) => {
    if (!blocker) return;
    blocker.classList.remove("hidden");
    window.setTimeout(() => blocker.classList.add("hidden"), 2200);
    reportAttempt(reason);
  };
  document.addEventListener("contextmenu", (event) => {
    event.preventDefault();
    showBlocker("Rechtsklick / Kontextmenü");
  });
  document.addEventListener("keydown", (event) => {
    const key = event.key.toLowerCase();
    const blocked = event.key === "F12"
      || (event.ctrlKey && event.shiftKey && ["i", "j", "c"].includes(key))
      || (event.ctrlKey && ["u", "s"].includes(key));
    if (!blocked) return;
    event.preventDefault();
    event.stopPropagation();
    showBlocker(`Tastenkürzel ${event.ctrlKey ? "Ctrl+" : ""}${event.shiftKey ? "Shift+" : ""}${event.key}`);
  }, true);
}

const state = {
  token: storedAuthToken(),
  currentUser: null,
  users: [],
  archivedUsers: [],
  ranks: [],
  roles: [],
  departmentPositions: [],
  settings: null,
  notes: [],
  duty: [],
  dutyHistory: [],
  logs: [],
  disciplinary: [],
  departments: [],
  customPages: [],
  page: localStorage.getItem("lspd_page") || "Dienstblatt",
  directionTab: localStorage.getItem("lspd_direction_tab") || "overview",
  profileTab: localStorage.getItem("lspd_profile_tab") || "Ausbildung",
  departmentTabs: JSON.parse(localStorage.getItem("lspd_department_tabs") || "{}")
};

const DISCORD_PENDING_TOKEN_KEY = "lspd_pending_discord_token";

const pages = [
  "Dienstblatt",
  "Einsatzzentrale",
  "Beschlagnahmung",
  "Informationen",
  "Ausbilderübersicht",
  "Meine Lernkontrollen",
  "Abteilungen",
  "Mitglieder",
  "Mitgliederfluktation",
  "Changelog",
  "Postfach",
  "Profil",
  "Kalender"
];

const pageIcons = {
  "Dienstblatt": "▣",
  "Einsatzzentrale": "◉",
  "Beschlagnahmung": "◇",
  "Informationen": "ⓘ",
  "Ausbilderübersicht": "□",
  "Meine Lernkontrollen": "✺",
  "Abteilungen": "▦",
  "Mitglieder": "♙",
  "Mitgliederfluktation": "↔",
  "Changelog": "☰",
  "Postfach": "✉",
  "Profil": "♙",
  "Kalender": "□",
  "Direktion": "◆"
};

const adminPages = ["IT", "Direktion"];
const positionOrder = { "Direktion": 5, "Leitung": 4, "Stv. Leitung": 3, "Mitglied": 2, "Anwärter": 1 };
const trainingGroups = [
  ["EST", "Wissen", "Fahren", "Schießen", "Verhalten", "Undercover", "Wanted"],
  ["EL", "Officer Prüfung", "Prak. VHF", "Prak. EL I", "Führung", "Prak. EL II"],
  ["Air Support", "Riot", "Coquette"]
];
const trainings = trainingGroups.flat();
const expandedDepartments = new Set(JSON.parse(localStorage.getItem("lspd_expanded_departments") || "[]"));
let calendarCursor = new Date(new Date().getFullYear(), new Date().getMonth(), 1);
let selectedCalendarDate = isoDateLocal(new Date());
let trainingTimerInterval = null;
const dutyOptions = [
  { title: "Innendienst", description: "Büro, Verwaltung, Leitstelle", icon: "Abteilungen", tone: "inside" },
  { title: "Außendienst", description: "Streife, Einsatz und Außendienst", icon: "Kalender", tone: "outside" },
  { title: "Undercover Dienst", description: "Zivile Arbeit und verdeckte Maßnahmen", icon: "Mitglieder", tone: "undercover" },
  { title: "Admin Dienst", description: "Teamler / administrative Tätigkeiten", icon: "IT", teamlerOnly: true, tone: "admin" }
];

function availableDutyOptions() {
  return dutyOptions;
}

const pageDescriptions = {
  "Einsatzzentrale": "Koordination laufender Einsätze und operativer Meldungen",
  "Beschlagnahmung": "Erfassung und Verwaltung beschlagnahmter Gegenstände",
  "Informationen": "Zentrale Informationen und Bewerbungsstatus verwalten",
  "Ausbilderübersicht": "Übersicht aller Ausbildungsmodule und Prüfungen",
  "Meine Lernkontrollen": "Eigene Lernkontrollen und Prüfungsstände einsehen",
  "Abteilungen": "Übersicht aller Abteilungen und Personal",
  "Mitglieder": "Übersicht aller aktiven Mitglieder und Ausbildungen",
  "Mitgliederfluktation": "Übersicht über Einstellungen und Kündigungen",
  "Changelog": "Änderungen und Neuerungen im Dienstblatt einsehen",
  "Postfach": "Interne Nachrichten und Mitteilungen verwalten",
  "Profil": "Eigene Accountdaten, Avatar und Passwort verwalten",
  "Kalender": "Termine, Dienste und wichtige Ereignisse planen",
  "Direktion": "Leitung, Verwaltung und Mitgliedersteuerung",
  "IT": "Systemeinstellungen, Reiter und Ränge verwalten"
};

const $ = (selector) => document.querySelector(selector);
const content = $("#content");
const modalRoot = $("#modalRoot");
const notifyRoot = $("#notifyRoot");
const warmedAvatarUrls = new Set();

function warmAvatarCache() {
  ["/assets/lspd-logo-20260515.png", ...(state.users || []).map((user) => user.avatarUrl).filter(Boolean)].forEach((url) => {
    if (warmedAvatarUrls.has(url)) return;
    warmedAvatarUrls.add(url);
    const image = new Image();
    image.decoding = "async";
    image.src = url;
  });
}

function hasRole(minRole) {
  const power = { User: 1, Supervisor: 2, Direktion: 3, IT: 4, "IT-Leitung": 5 };
  return (power[state.currentUser?.role] || 0) >= (power[minRole] || 0);
}

function canManageFluctuation() {
  return state.currentUser?.role === "IT-Leitung";
}

function permissionAllows(rule, user = state.currentUser) {
  if (!rule) return false;
  if (hasRole("IT")) return true;
  if (user.role === "Direktion") return true;
  const departmentMatch = (rule.departments || []).some((departmentId) => {
    const department = state.departments.find((item) => item.id === departmentId);
    return department?.members?.some((member) => member.userId === user.id);
  });
  const positionMatch = (rule.positions || []).some((positionKey) => {
    const [departmentId, position] = String(positionKey).split(":");
    const department = state.departments.find((item) => item.id === departmentId);
    return department?.members?.some((member) => member.userId === user.id && member.position === position);
  });
  return Boolean(rule.all) || (rule.users || []).includes(user.id) || (rule.roles || []).includes(user.role) || (rule.ranks || []).map(Number).includes(Number(user.rank)) || departmentMatch || positionMatch;
}

function canAccess(area, key, fallbackRole = "IT") {
  if (hasRole("IT")) return true;
  const rule = state.settings?.permissions?.[area]?.[key];
  return rule ? permissionAllows(rule) : hasRole(fallbackRole);
}

function canSeeDepartment(page) {
  if (hasRole("IT")) return true;
  if (page === "IT") return hasRole("IT");
  if (state.currentUser?.role === "Direktion") return true;
  const rule = state.settings?.permissions?.pages?.[page];
  if (rule) return permissionAllows(rule);
  if (page === "Direktion") return hasRole("Direktion");
  return true;
}

function isPageItOnlyVisible(page) {
  const rule = state.settings?.permissions?.pages?.[page];
  if (!rule || rule.all) return false;
  const roles = rule.roles || [];
  const onlyItRole = roles.every((role) => ["IT", "IT-Leitung"].includes(role));
  const hasOtherSelectors = Boolean((rule.users || []).length || (rule.ranks || []).length || (rule.departments || []).length || (rule.positions || []).length);
  return onlyItRole && !hasOtherSelectors;
}

function isDepartmentPage(page) {
  return page.startsWith("dept:");
}

function departmentByPage(page) {
  return state.departments.find((department) => `dept:${department.id}` === page);
}

function fullName(user = state.currentUser) {
  return `${user?.firstName || ""} ${user?.lastName || ""}`.trim();
}

function avatarMarkup(user = state.currentUser, size = "md") {
  if (user?.avatarUrl) {
    return `<img class="avatar ${size}" src="${escapeHtml(user.avatarUrl)}" alt="Avatar" loading="eager" decoding="async" fetchpriority="high">`;
  }
  return `<img class="avatar ${size}" src="/assets/lspd-logo-20260515.png" alt="LSPD" loading="eager" decoding="async" fetchpriority="high">`;
}

function rankLabel(rank) {
  const found = state.ranks.find((item) => Number(item.value) === Number(rank));
  const label = found ? found.label : `Template ${rank} - Rang ${rank}`;
  const clean = String(label).replace(/^\s*\(?\d+\)?\s*/, "").trim();
  return `(${Number(rank)}) ${clean || `Rang ${rank}`}`;
}

function rankOptionLabel(rank) {
  return rankLabel(rank.value);
}

function navLabel(page) {
  if (isDepartmentPage(page)) return departmentByPage(page)?.name || page;
  const custom = state.customPages?.find((item) => item.key === page);
  if (custom) return state.settings?.navLabels?.[page] || custom.name || page;
  return state.settings?.navLabels?.[page] || page;
}

function pageDescription(page) {
  if (isDepartmentPage(page)) {
    const department = departmentByPage(page);
    return department ? `${department.name} - ${department.description}` : "Abteilungsübersicht und interne Notizen";
  }
  return pageDescriptions[page] || "Diese Seite kann später weiter ausgebaut werden";
}

function iconSvg(page) {
  if (isDepartmentPage(page)) page = "Direktion";
  if (page === "Mitgliederfluktation") return `<img class="asset-icon asset-icon-fluctuation" src="/fluctuation-icon.svg" alt="">`;
  const icons = {
    "Dienstblatt": '<path d="M8 4h8l2 2v14H6V4Z"/><path d="M9 9h6M9 13h6M9 17h4"/>',
    "Einsatzzentrale": '<path d="M4 12a8 8 0 0 1 16 0"/><path d="M7 12a5 5 0 0 1 10 0"/><path d="M12 12v5"/><path d="M9 17h6"/>',
    "Beschlagnahmung": '<path d="m12 3 8 4.5v9L12 21l-8-4.5v-9L12 3Z"/><path d="M12 12 4.5 7.7M12 12l7.5-4.3M12 12v8.5"/>',
    "Informationen": '<circle cx="12" cy="12" r="9"/><path d="M12 10v6M12 7h.01"/>',
    "Ausbilderübersicht": '<path d="M4 5h7a3 3 0 0 1 3 3v11a3 3 0 0 0-3-3H4V5Z"/><path d="M20 5h-6a3 3 0 0 0-3 3"/>',
    "Meine Lernkontrollen": '<path d="M8 3a4 4 0 0 0-4 4v1a3 3 0 0 0 0 6v1a4 4 0 0 0 4 4"/><path d="M16 3a4 4 0 0 1 4 4v1a3 3 0 0 1 0 6v1a4 4 0 0 1-4 4"/><path d="M8 12h8"/>',
    "Abteilungen": '<rect x="4" y="4" width="16" height="16" rx="2"/><path d="M8 8h3v3H8zM13 8h3v3h-3zM8 13h3v3H8zM13 13h3v3h-3z"/>',
    "Mitglieder": '<path d="M16 21v-2a4 4 0 0 0-8 0v2"/><circle cx="12" cy="7" r="4"/><path d="M20 21v-2a3 3 0 0 0-2-2.8M4 21v-2a3 3 0 0 1 2-2.8"/>',
    "Mitgliederfluktation": '<path d="M7 7h11l-3-3M17 17H6l3 3"/>',
    "Changelog": '<path d="M8 4h8l2 2v14H6V4Z"/><path d="M9 9h6M9 13h6M9 17h3"/>',
    "Postfach": '<path d="M4 5h16v12H7l-3 3V5Z"/>',
    "Profil": '<circle cx="12" cy="8" r="4"/><path d="M6 21a6 6 0 0 1 12 0"/>',
    "Kalender": '<rect x="4" y="5" width="16" height="15" rx="2"/><path d="M16 3v4M8 3v4M4 10h16"/>',
    "Settings": '<circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.8 1.8 0 0 0 .4 2l.1.1-2 3.4-.2-.1a1.8 1.8 0 0 0-2 .4l-.2.2-3.8-2.2.1-.3a1.8 1.8 0 0 0-.8-1.8 1.8 1.8 0 0 0-2 .1l-.3.2-3.3-2 .1-.3a1.8 1.8 0 0 0-.4-2l-.2-.2 2-3.4.3.1a1.8 1.8 0 0 0 2-.4l.2-.2 3.8 2.2-.1.3a1.8 1.8 0 0 0 .8 1.8 1.8 1.8 0 0 0 2-.1l.3-.2 3.3 2Z"/>',
    "ChevronUp": '<path d="m18 15-6-6-6 6"/>',
    "ChevronDown": '<path d="m6 9 6 6 6-6"/>',
    "Plus": '<path d="M12 5v14M5 12h14"/>',
    "Lock": '<rect x="5" y="10" width="14" height="10" rx="2"/><path d="M8 10V7a4 4 0 0 1 8 0v3"/>',
    "EyeOff": '<path d="M3 3l18 18"/><path d="M10.6 10.6A2 2 0 0 0 13.4 13.4"/><path d="M9.9 5.1A9.8 9.8 0 0 1 12 5c5 0 8.5 4.4 9.7 6.1a1.6 1.6 0 0 1 0 1.8 16.5 16.5 0 0 1-2.5 2.9"/><path d="M6.2 6.2A16.8 16.8 0 0 0 2.3 11.1a1.6 1.6 0 0 0 0 1.8C3.5 14.6 7 19 12 19a9.7 9.7 0 0 0 4-.8"/>',
    "Direktion": '<path d="M12 3 20 7v6c0 5-3.5 7.5-8 8-4.5-.5-8-3-8-8V7l8-4Z"/><path d="M9 12l2 2 4-4"/>',
    "IT": '<path d="M8 18h8M10 22h4"/><rect x="3" y="4" width="18" height="14" rx="2"/><path d="M8 9h8M8 13h5"/>',
    "Logout": '<path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><path d="M16 17l5-5-5-5"/><path d="M21 12H9"/>'
  };
  return `<svg viewBox="0 0 24 24" aria-hidden="true">${icons[page] || icons.Dienstblatt}</svg>`;
}

function actionIcon(type) {
  const src = type === "delete" ? "/loeschen.png" : "/bearbeiten.png";
  return `<img class="asset-action-icon" src="${src}" alt="" draggable="false">`;
}

function formatDate(value) {
  if (!value) return "-";
  return new Date(value).toLocaleDateString("de-DE");
}

function formatTime(value) {
  if (!value) return "-";
  return new Date(value).toLocaleTimeString("de-DE", { hour: "2-digit", minute: "2-digit" });
}

function formatDateTime(value) {
  if (!value) return "-";
  return `${formatDate(value)} ${formatTime(value)}`;
}

function isoDateLocal(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function monthName(date) {
  return date.toLocaleDateString("de-DE", { month: "long", year: "numeric" });
}

function calendarDayTitle(value) {
  if (!value) return "-";
  return new Date(`${value}T00:00`).toLocaleDateString("de-DE", { day: "numeric", month: "long" });
}

function durationMs(entry) {
  const start = new Date(entry.startedAt).getTime();
  const end = entry.endedAt ? new Date(entry.endedAt).getTime() : Date.now();
  return Math.max(0, end - start);
}

function formatDuration(ms) {
  const minutes = Math.floor(ms / 60000);
  const hours = Math.floor(minutes / 60);
  const rest = minutes % 60;
  return `${hours}h ${String(rest).padStart(2, "0")}m`;
}

function wrapNameForTable(name) {
  return escapeHtml(name).replace(/-/g, "-<wbr>");
}

function rangeStart(range) {
  const now = new Date();
  if (range === "Woche") {
    const week = new Date(now);
    week.setDate(now.getDate() - 7);
    return week;
  }
  if (range === "Monat") return new Date(now.getFullYear(), now.getMonth(), 1);
  if (range === "Heute") return new Date(now.getFullYear(), now.getMonth(), now.getDate());
  return null;
}

function roleClass(role) {
  return `role-${String(role || "User").toLowerCase().replace(/[^a-z0-9]+/g, "-")}`;
}

function roleBadges(user) {
  const baseRole = user?.baseRole || (["IT", "IT-Leitung"].includes(user?.role) ? "Direktion" : user?.role || "User");
  const roles = user?.role === "IT-Leitung" ? [baseRole, "IT", "IT-Leitung"] : user?.role === "IT" ? [baseRole, "IT"] : [baseRole];
  if (user?.teamler) roles.push("Teamler");
  return roles.map((role) => `<span class="role-pill ${roleClass(role)}">${escapeHtml(role)}</span>`).join("");
}

function cleanText(value) {
  let text = String(value ?? "");
  const decodeOnce = (input) => {
    if (!/[ÃÂ]/.test(input)) return input;
    try {
      return decodeURIComponent(Array.from(input, (char) => {
        const code = char.charCodeAt(0);
        return code <= 255 ? `%${code.toString(16).padStart(2, "0")}` : encodeURIComponent(char);
      }).join(""));
    } catch {
      return input;
    }
  };
  for (let index = 0; index < 3; index += 1) {
    const decoded = decodeOnce(text);
    if (decoded === text) break;
    text = decoded;
  }
  return text
    .replaceAll("Â·", "·")
    .replaceAll("Â", "")
    .replaceAll("Verk\u003frzung", "Verkürzung")
    .replaceAll("\u003fbersicht", "Übersicht")
    .replaceAll("\u003fber", "Über")
    .replaceAll("\u003fndern", "Ändern")
    .replaceAll("\u003fnderung", "Änderung")
    .replaceAll("\u003fnderungen", "Änderungen");
}

function describeLogDetails(log) {
  const details = log.details || {};
  const action = cleanText(log.action);
  if (details.description) return cleanText(details.description);
  if (action === "Avatar ge\u00e4ndert") return "Avatar wurde ge\u00e4ndert.";
  if (action === "Passwort ge\u00e4ndert") return "Passwort wurde ge\u00e4ndert.";
  if (action === "Login") return "Angemeldet.";
  if (action === "Logout") return "Abgemeldet.";
  if (action.includes("Aktennotiz")) return `Notiz ${action.includes("bearbeitet") ? "bearbeitet" : action.includes("entfernt") ? "entfernt" : "hinzugef\u00fcgt"}${details.reason ? `: ${cleanText(details.reason)}` : ""}`;
  if (action.includes("Sanktion")) return `Sanktion ${action.includes("archiviert") ? "archiviert" : "hinzugef\u00fcgt"}${details.reason ? `: ${cleanText(details.reason)}` : ""}${details.amount ? ` \u00b7 ${details.amount}$` : ""}`;
  if (action.includes("Geldstrafe")) return `Geldstrafe ${action.includes("bezahlt") ? "als bezahlt markiert" : "bearbeitet"}.`;
  if (action.includes("Uprank")) return details.reason ? `Uprank: ${cleanText(details.reason)}` : "Uprank wurde dokumentiert.";
  if (details.before && details.after) {
    const changeText = describeObjectChanges(details.before, details.after);
    if (changeText) return changeText;
  }
  if (details.reason) return `Grund: ${cleanText(details.reason)}`;
  if (action.includes("Dienst gestartet")) return `In Dienst: ${cleanText(log.target)}`;
  if (action.includes("Dienst beendet")) return `Au\u00dfer Dienst: ${cleanText(details.before?.status || log.target || "")}`;
  if (action.includes("erstellt")) return "Eintrag wurde erstellt.";
  if (action.includes("gel\u00f6scht")) return "Eintrag wurde gel\u00f6scht.";
  if (action.includes("bearbeitet") || action.includes("ge\u00e4ndert")) return "Eintrag wurde ge\u00e4ndert.";
  return "";
}
function renderLogDetails(log) {
  const text = cleanText(describeLogDetails(log));
  if (!text) return `<span class="muted">-</span>`;
  const parts = text.split(";").map((part) => part.trim()).filter(Boolean);
  return `<div class="log-detail-text">${parts.map((part) => renderLogDetailPart(cleanText(part))).join("<span class=\"log-detail-separator\">;</span> ")} </div>`;
}

function renderLogDetailPart(part) {
  const match = part.match(/^(Ausbildung)\s+(.+?)\s+(hinzugef\u00fcgt|entfernt)$/i);
  if (match) {
    const tone = match[3] === "entfernt" ? "bad" : "good";
    return `${escapeHtml(match[1])} <strong>${escapeHtml(match[2].toUpperCase())}</strong> <mark class="${tone}">${escapeHtml(match[3])}</mark>`;
  }
  const changeMatch = part.match(/^([^:]+):\s*(.*?)\s*->\s*(.*)$/);
  if (changeMatch) {
    return `<strong>${escapeHtml(changeMatch[1])}</strong>: ${escapeHtml(changeMatch[2] || "-")} <span class="log-arrow">\u2192</span> ${escapeHtml(changeMatch[3] || "-")}`;
  }
  const actionMatch = part.match(/(hinzugef\u00fcgt|entfernt|bearbeitet|erstellt|gel\u00f6scht|bezahlt|Uprank|gesperrt|entlassen)/i);
  if (!actionMatch) return escapeHtml(part);
  const before = part.slice(0, actionMatch.index);
  const action = actionMatch[0];
  const after = part.slice(actionMatch.index + action.length);
  const tone = /entfernt|gel\u00f6scht|gesperrt|entlassen/i.test(action) ? "bad" : /hinzugef\u00fcgt|erstellt|bezahlt|Uprank/i.test(action) ? "good" : "neutral";
  return `${escapeHtml(before)}<mark class="${tone}">${escapeHtml(action)}</mark>${escapeHtml(after)}`;
}
function describeObjectChanges(before = {}, after = {}) {
  const changes = [];
  const fields = [
    ["firstName", "Vorname"],
    ["lastName", "Nachname"],
    ["phone", "Telefon"],
    ["dn", "Dienstnummer"],
    ["role", "Rolle"],
    ["title", "Titel"],
    ["priority", "Priorität"],
    ["text", "Text"],
    ["reason", "Grund"],
    ["status", "Status"]
  ];
  fields.forEach(([key, label]) => {
    const oldValue = String(before?.[key] ?? "");
    const newValue = String(after?.[key] ?? "");
    if (oldValue !== newValue) changes.push(`${label}: ${oldValue || "-"} -> ${newValue || "-"}`);
  });
  if ((before?.rank !== undefined || after?.rank !== undefined) && Number(before?.rank) !== Number(after?.rank)) {
    changes.push(`Rang: ${rankLabel(before?.rank)} -> ${rankLabel(after?.rank)}`);
  }
  trainings.forEach((training) => {
    const had = Boolean(before?.trainings?.[training]);
    const has = Boolean(after?.trainings?.[training]);
    if (had !== has) changes.push(`Ausbildung ${cleanText(training)} ${has ? "hinzugef\u00fcgt" : "entfernt"}`);
  });
  return changes.join("; ");
}

function logTone(action) {
  action = cleanText(action);
  if (/hinzugefügt|Uprank/i.test(action)) return "log-good";
  if (/erstellt|gestartet|hinzugefügt|Login|eingestellt/i.test(action)) return "log-good";
  if (/gelöscht|entlassen|gesperrt|beendet|Logout|Kündigung/i.test(action)) return "log-bad";
  if (/geändert|bearbeitet|aktualisiert/i.test(action)) return "log-warn";
  return "";
}

function escapeHtml(value) {
  return cleanText(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

async function api(path, options = {}) {
  const method = String(options.method || "GET").toUpperCase();
  const shouldNotify = !options.silent && method !== "GET";
  const response = await fetch(path, {
    ...options,
    headers: {
      "Content-Type": "application/json",
      ...(state.token ? { Authorization: `Bearer ${state.token}` } : {}),
      ...(options.headers || {})
    }
  });
  const data = await response.json().catch(() => ({}));
  if (!response.ok) {
    const message = data.error || "Aktion fehlgeschlagen.";
    if (shouldNotify) showNotify(message, "error");
    throw new Error(message);
  }
  if (shouldNotify) showNotify(successMessage(path, method), "success");
  return data;
}

async function startDiscordOAuth(mode = "login") {
  const targetError = $("#loginError");
  const button = mode === "login" ? $("#discordLoginBtn") : null;
  try {
    if (button) {
      button.disabled = true;
      button.classList.add("loading");
      button.textContent = "Weiter zu Discord...";
    }
    const config = await api("/api/discord/oauth-config", { silent: true });
    if (!config.applicationId) throw new Error("Discord Login ist noch nicht eingerichtet.");
    const oauthState = crypto.randomUUID ? crypto.randomUUID() : String(Date.now());
    sessionStorage.setItem("lspd_discord_oauth_state", JSON.stringify({ state: oauthState, mode }));
    const redirectUri = config.oauthRedirectUrl || `${window.location.origin}/`;
    const params = new URLSearchParams({
      client_id: config.applicationId,
      redirect_uri: redirectUri,
      response_type: "token",
      scope: "identify",
      state: oauthState,
      prompt: "consent"
    });
    window.location.href = `https://discord.com/oauth2/authorize?${params.toString()}`;
  } catch (error) {
    if (button) {
      button.disabled = false;
      button.classList.remove("loading");
      button.textContent = "Mit Discord einloggen";
    }
    if (targetError) targetError.textContent = error.message;
    else showNotify(error.message, "error");
  }
}

async function completeDiscordOAuth(accessToken, mode) {
  if (mode === "link" || state.token) {
    if (!state.token) {
      sessionStorage.setItem(DISCORD_PENDING_TOKEN_KEY, accessToken);
      showLogin();
      $("#loginError").textContent = "Discord erkannt. Bitte melde dich zuerst normal an. Nach dem Passwortwechsel kannst du Discord im Profil verknüpfen.";
      return;
    }
    const data = await api("/api/discord/link", { method: "POST", body: JSON.stringify({ accessToken }) });
    state.currentUser = data.user || state.currentUser;
    await bootstrap();
    showNotify(`Discord verknüpft: ${data.discordUser?.globalName || data.discordUser?.username || "Account"}`);
    return;
  }
  try {
    const data = await api("/api/discord/login", { method: "POST", body: JSON.stringify({ accessToken }) });
    state.token = data.token;
    storeAuthToken(state.token);
    await bootstrap();
  } catch (error) {
    sessionStorage.setItem(DISCORD_PENDING_TOKEN_KEY, accessToken);
    showLogin();
    $("#loginError").textContent = "Discord ist noch nicht verknüpft. Melde dich zuerst normal an und verknüpfe Discord danach im Profil.";
  }
}

async function handleDiscordOAuthRedirect() {
  if (!window.location.hash.includes("access_token=")) return false;
  const params = new URLSearchParams(window.location.hash.slice(1));
  const accessToken = params.get("access_token") || "";
  const returnedState = params.get("state") || "";
  const stored = JSON.parse(sessionStorage.getItem("lspd_discord_oauth_state") || "{}");
  window.history.replaceState(null, document.title, `${window.location.pathname}${window.location.search}`);
  if (!accessToken || !stored.state || returnedState !== stored.state) {
    showLogin();
    $("#loginError").textContent = "Discord Login konnte nicht geprüft werden.";
    return true;
  }
  sessionStorage.removeItem("lspd_discord_oauth_state");
  await completeDiscordOAuth(accessToken, stored.mode || "login");
  return true;
}

async function linkPendingDiscordAccount() {
  const accessToken = sessionStorage.getItem(DISCORD_PENDING_TOKEN_KEY);
  if (!accessToken || !state.token) return;
  if (state.currentUser?.mustChangePassword) return;
  try {
    const data = await api("/api/discord/link", { method: "POST", body: JSON.stringify({ accessToken }) });
    sessionStorage.removeItem(DISCORD_PENDING_TOKEN_KEY);
    state.currentUser = data.user || state.currentUser;
    showNotify(`Discord verknüpft: ${data.discordUser?.globalName || data.discordUser?.username || "Account"}`);
  } catch (error) {
    $("#loginError").textContent = error.message;
  }
}

function successMessage(path, method) {
  if (path.includes("/login")) return "Erfolgreich angemeldet.";
  if (path.includes("/logout")) return "Erfolgreich abgemeldet.";
  if (path.includes("/duty/start")) return "Dienst gestartet.";
  if (path.includes("/duty/stop-all")) return "Alle Officer wurden ausgetragen.";
  if (path.includes("/duty/stop")) return "Dienst beendet.";
  if (path.includes("/notes") && method === "POST") return "Notiz erstellt.";
  if (path.includes("/notes") && method === "PATCH") return "Notiz aktualisiert.";
  if (path.includes("/notes") && method === "DELETE") return "Notiz gelöscht.";
  if (path.includes("/departments") && path.includes("/members") && method === "POST") return "Person hinzugefügt.";
  if (path.includes("/departments") && path.includes("/members") && method === "PATCH") return "Position aktualisiert.";
  if (path.includes("/departments") && path.includes("/members") && method === "DELETE") return "Person entfernt.";
  if (path.includes("/departments") && path.includes("/info")) return "Abteilungsinfos gespeichert.";
  if (path.includes("/settings/defcon")) return "DEFCON aktualisiert.";
  if (path.includes("/profile/password")) return "Passwort geändert.";
  if (path.includes("/profile/avatar")) return "Avatar gespeichert.";
  if (path.includes("/file")) return method === "DELETE" ? "Akteneintrag entfernt." : "Akteneintrag gespeichert.";
  if (path.includes("/seizures") && method === "POST") return "Beschlagnahmung eingetragen.";
  if (path.includes("/seizures") && method === "PATCH") return "Beschlagnahmung gespeichert.";
  if (path.includes("/seizures") && method === "DELETE") return "Beschlagnahmung gelöscht.";
  if (path.includes("/settings/fluctuation") && method === "PATCH") return "Fluktuationseintrag gespeichert.";
  if (path.includes("/settings/fluctuation") && method === "DELETE") return "Fluktuationseintrag gelöscht.";
  if (path.includes("/suspend")) return "Mitglied suspendiert.";
  if (path.includes("/dismiss")) return "Mitglied entlassen.";
  if (path.includes("/users") && method === "POST") return "Mitglied eingestellt.";
  if (path.includes("/users") && method === "PATCH") return "Account aktualisiert.";
  if (path.includes("/users") && method === "DELETE") return "Account gelöscht.";
  if (path.includes("/it/ranks")) return "Ränge gespeichert.";
  if (path.includes("/it/nav-labels")) return "Reiter gespeichert.";
  if (path.includes("/it/permissions")) return "Rechte gespeichert.";
  if (path.includes("/it/default-password")) return "Standardpasswort gespeichert.";
  if (path.includes("/reset-password")) return "Passwort zurückgesetzt.";
  if (path.includes("/information")) return "Informationen gespeichert.";
  return "Aktion erfolgreich.";
}

function showNotify(message, type = "success") {
  if (type === "success" && /(gelöscht|löschen|entfernt|entfernen|abgelehnt|fehlgeschlagen|fehler)/i.test(cleanText(message))) {
    type = /(fehlgeschlagen|fehler)/i.test(cleanText(message)) ? "error" : "danger";
  }
  const duration = Math.min(5000, Math.max(3000, 2200 + String(message).length * 35));
  const item = document.createElement("div");
  item.className = `notify ${type}`;
  item.style.setProperty("--notify-duration", `${duration}ms`);
  item.innerHTML = `
    <div class="notify-icon">${type === "success" ? "✓" : type === "danger" ? "×" : "!"}</div>
    <button class="notify-close" type="button" aria-label="Benachrichtigung schliessen">&times;</button>
    <div class="notify-copy">
      <strong>${type === "success" ? "Erfolg" : type === "danger" ? "Gelöscht" : "Fehler"}</strong>
      <span>${escapeHtml(message)}</span>
    </div>
    <div class="notify-progress"></div>
  `;
  let remaining = duration;
  let startedAt = 0;
  let closeTimer = null;
  let removeTimer = null;
  let closed = false;
  const clearTimers = () => {
    window.clearTimeout(closeTimer);
    window.clearTimeout(removeTimer);
  };
  const close = () => {
    if (closed) return;
    closed = true;
    clearTimers();
    item.classList.remove("paused");
    item.classList.add("leaving");
    removeTimer = window.setTimeout(() => item.remove(), 280);
  };
  const armTimer = () => {
    startedAt = performance.now();
    closeTimer = window.setTimeout(close, Math.max(120, remaining));
  };
  item.addEventListener("mouseenter", () => {
    if (closed) return;
    window.clearTimeout(closeTimer);
    remaining = Math.max(120, remaining - (performance.now() - startedAt));
    item.classList.add("paused");
  });
  item.addEventListener("mouseleave", () => {
    if (closed) return;
    item.classList.remove("paused");
    armTimer();
  });
  item.addEventListener("click", (event) => {
    if (event.target.closest(".notify-close")) return;
    close();
  });
  item.querySelector(".notify-close")?.addEventListener("click", (event) => {
    event.stopPropagation();
    close();
  });
  notifyRoot.appendChild(item);
  armTimer();
}

async function bootstrap() {
  const data = await api("/api/bootstrap");
  Object.assign(state, data);
  warmAvatarCache();
  syncDevModeAuthStorage();
  const visiblePages = getVisiblePages();
  if (!visiblePages.includes(state.page)) {
    state.page = "Dienstblatt";
    localStorage.setItem("lspd_page", state.page);
  }
  renderApp();
}

function showLogin() {
  $("#loadingView")?.classList.add("hidden");
  $("#loginView").classList.remove("hidden");
  $("#appView").classList.add("hidden");
}

function showApp() {
  $("#loadingView")?.classList.add("hidden");
  $("#loginView").classList.add("hidden");
  $("#appView").classList.remove("hidden");
}

function renderApp() {
  showApp();
  if (state.currentUser?.mustChangePassword) {
    renderPasswordChangeRequired();
    return;
  }
  renderNavigation();
  renderTopbar();
  renderDevModeBanner();
  renderPage();
}

function renderPasswordChangeRequired() {
  $(".profile-card").innerHTML = `
    ${avatarMarkup(state.currentUser, "lg")}
    <div class="profile-copy">
      <strong>${escapeHtml(fullName())}</strong>
      <span>Passwortwechsel erforderlich</span>
      <em class="off">Gesperrte Ansicht</em>
    </div>
  `;
  $("#navigation").innerHTML = "";
  $("#pageTitle").textContent = "Passwort ändern";
  $("#rankLine").textContent = "Sicherheitsprüfung";
  const description = $("#pageDescription");
  if (description) description.textContent = "Du musst dein Passwort ändern, bevor du das Dienstblatt nutzen kannst.";
  $("#serviceStatus").textContent = "Passwortwechsel erforderlich";
  $("#serviceStatus").className = "service-pill off";
  $("#headerIcon").innerHTML = iconSvg("IT");
  $("#headerIcon").classList.remove("hidden");
  renderDevModeBanner();
  content.innerHTML = `
    <section class="force-password-stage">
      <div class="force-password-brand">
        <img src="/assets/lspd-logo-20260515.png" alt="LSPD">
        <span>LSPD Dienstblatt</span>
      </div>
      <div class="panel force-password-panel">
        <span class="login-kicker">Sicherheitsprüfung</span>
        <h3>Passwort ändern</h3>
        <p class="muted">Du bist mit dem Standardpasswort angemeldet. Aus Sicherheitsgründen musst du jetzt ein eigenes Passwort festlegen, bevor du das Dienstblatt sehen kannst.</p>
        <div class="security-note-box">
          <strong>Passwörter sind geschützt und nicht einsehbar.</strong>
          <span>Auch die IT kann dein Passwort nicht auslesen. Es kann nur auf das Standardpasswort zurückgesetzt und danach von dir neu gesetzt werden.</span>
        </div>
        <label>Neues Passwort<input type="password" id="forcedNewPassword" autocomplete="new-password" required></label>
        <label>Neues Passwort wiederholen<input type="password" id="forcedRepeatPassword" autocomplete="new-password" required></label>
        <p id="forcedPasswordError" class="form-error"></p>
        <button class="orange-btn" id="saveForcedPassword" type="button">Passwort speichern</button>
      </div>
    </section>
  `;
  $("#saveForcedPassword")?.addEventListener("click", saveForcedPassword);
}

async function saveForcedPassword() {
  const newPassword = $("#forcedNewPassword")?.value || "";
  const repeatPassword = $("#forcedRepeatPassword")?.value || "";
  if (!newPassword) {
    $("#forcedPasswordError").textContent = "Bitte ein neues Passwort eintragen.";
    return;
  }
  if (newPassword !== repeatPassword) {
    $("#forcedPasswordError").textContent = "Die neuen Passwörter stimmen nicht überein.";
    return;
  }
  try {
    await api("/api/profile/password", { method: "PATCH", body: JSON.stringify({ newPassword }) });
    await bootstrap();
  } catch (error) {
    $("#forcedPasswordError").textContent = error.message;
  }
}

function syncDevModeAuthStorage() {
  const active = Boolean(state.settings?.devMode);
  localStorage.setItem("lspd_devmode_active", active ? "1" : "0");
  if (active) {
    if (state.token) sessionStorage.setItem("lspd_token_dev", state.token);
    localStorage.removeItem("lspd_token");
  } else {
    if (state.token) localStorage.setItem("lspd_token", state.token);
    sessionStorage.removeItem("lspd_token_dev");
  }
}

function renderDevModeBanner() {
  const banner = $("#devModeBanner");
  if (!banner) return;
  banner.classList.toggle("hidden", !state.settings?.devMode);
  banner.innerHTML = state.settings?.devMode ? `
    <strong>DEVMODE</strong>
    <span>aktiv</span>
  ` : "";
}

function renderNavigation() {
  const visiblePages = getVisiblePages();
  const myDuty = state.duty.find((entry) => entry.userId === state.currentUser.id);

  $(".profile-card").innerHTML = `
    ${avatarMarkup(state.currentUser, "lg")}
    <div class="profile-copy">
      <strong>${escapeHtml(fullName())}</strong>
      <span>${escapeHtml(rankLabel(state.currentUser.rank))}</span>
      <em class="${myDuty ? "on" : "off"}">${myDuty ? "Im Dienst" : "Außer Dienst"}</em>
    </div>
  `;

  $("#navigation").innerHTML = visiblePages.map((page) => `
    <button class="nav-btn ${state.page === page ? "active" : ""}" data-page="${escapeHtml(page)}">
      <span class="nav-icon">${iconSvg(page)}</span>
      <span class="nav-label">${escapeHtml(navLabel(page))}</span>
      ${restrictedPageIcon(page)}
    </button>
  `).join("");

  document.querySelectorAll(".nav-btn").forEach((button) => {
    button.addEventListener("click", () => {
      if (button.dataset.page === "IT" && state.page !== "IT") localStorage.setItem("lspd_it_tab", "overview");
      state.page = button.dataset.page;
      localStorage.setItem("lspd_page", state.page);
      renderApp();
    });
  });
}

function getVisiblePages() {
  const departmentNav = (state.departments || [])
    .filter((department) => department.id !== "direktion" && department.canOpen)
    .map((department) => `dept:${department.id}`);
  const basePages = [
    ...pages.filter(canSeeDepartment),
    ...adminPages.filter(canSeeDepartment),
    ...(state.customPages || []).map((page) => page.key).filter(canSeeDepartment),
    ...departmentNav.filter(canSeeDepartment)
  ];
  return orderPages(basePages);
}

function renderTopbar() {
  const myDuty = state.duty.find((entry) => entry.userId === state.currentUser.id);
  $("#pageTitle").textContent = navLabel(state.page);
  $("#rankLine").textContent = state.page === "Dienstblatt" ? `Willkommen zurück, ${fullName()}` : pageDescription(state.page);
  $("#headerIcon").innerHTML = state.page === "Dienstblatt" ? "" : iconSvg(state.page);
  $("#headerIcon").classList.toggle("hidden", state.page === "Dienstblatt");
  $("#headerTitleBlock").classList.toggle("with-icon", state.page !== "Dienstblatt");
  $("#serviceStatus").textContent = myDuty ? "Im Dienst" : "Außer Dienst";
  $("#serviceStatus").className = `service-pill ${myDuty ? "on" : "off"}`;
}

function renderPage() {
  if (state.page === "Dienstblatt") return renderDienstblatt();
  if (state.page === "Mitglieder") return renderMembers();
  if (state.page === "Mitgliederfluktation") return renderFluctuation();
  if (state.page === "Beschlagnahmung") return renderSeizures();
  if (state.page === "Kalender") return renderCalendar();
  if (state.page === "Informationen") return renderInformation();
  if (state.page === "Direktion") return renderDirektion();
  if (state.page === "IT") return renderIT();
  if (state.page === "Abteilungen") return renderDepartmentsOverview();
  if (isDepartmentPage(state.page)) return renderDepartmentPage(departmentByPage(state.page));
  if (state.page === "Profil") return renderProfile();
  return renderTemplate(state.page);
}

function renderDienstblatt() {
  const agents = state.duty.length;
  const undercover = state.duty.filter((entry) => entry.status === "Undercover Dienst").length;
  const outside = state.duty.filter((entry) => entry.status === "Außendienst").length;
  const inside = state.duty.filter((entry) => entry.status === "Innendienst").length;
  const adminDuty = state.duty.filter((entry) => entry.status === "Admin Dienst").length;
  const myDuty = state.duty.find((entry) => entry.userId === state.currentUser.id);

  content.innerHTML = `
    <section class="panel defcon-panel ${defconClass(state.settings.defcon)}">
      <div>
        <div class="defcon-value">${escapeHtml(state.settings.defcon)}</div>
        ${state.settings.defconText ? `<p>${escapeHtml(state.settings.defconText)}</p>` : ""}
      </div>
      <div class="defcon-meta">
        <div>Aktualisiert von ${escapeHtml(state.settings.defconUpdatedBy)} - ${formatDate(state.settings.defconUpdatedAt)} - ${formatTime(state.settings.defconUpdatedAt)}</div>
      </div>
      ${canAccess("actions", "editDefcon", "Supervisor") ? `<button class="icon-btn" id="defconBtn" title="DEFCON bearbeiten">⚙</button>` : ""}
    </section>

    <section class="grid-4 dashboard-stats">
      <div class="stat-card"><span>Aktive Officer</span><i>${iconSvg("Einsatzzentrale")}</i><strong>${agents}</strong><small>Im Einsatz</small></div>
      <div class="stat-card"><span>Außendienst</span><i>${iconSvg("Kalender")}</i><strong>${outside}</strong><small>Auf Streife</small></div>
      <div class="stat-card"><span>Undercover Dienst</span><i>${iconSvg("Mitglieder")}</i><strong>${undercover}</strong><small>Zivil Einheit</small></div>
      <div class="stat-card"><span>Innendienst ${adminDuty ? `<em class="admin-duty-count">(${adminDuty})</em>` : ""}</span><i>${iconSvg("Abteilungen")}</i><strong>${inside}</strong><small>Im Büro${adminDuty ? " · Admin Dienst" : ""}</small></div>
    </section>

    <section class="panel">
      <div class="panel-header">
        <h3><span class="section-icon">▣</span>Dienstblatt-Notizen</h3>
        ${canAccess("actions", "manageNotes", "Supervisor") ? `<button class="blue-btn" id="addNoteBtn"><span>+</span> Notiz hinzufügen</button>` : ""}
      </div>
      <div class="note-list">
        ${state.notes.length ? state.notes.map(renderNote).join("") : `<p class="muted">Noch keine Notizen vorhanden.</p>`}
      </div>
    </section>

    <section class="panel">
      <div class="panel-header">
        <h3><span class="section-icon">♙</span>Aktive Officer</h3>
        <div class="button-row">
          ${myDuty ? `<button class="ghost-btn action-btn" id="switchDutyBtn"><span>${iconSvg("Einsatzzentrale")}</span> Umtragen</button>` : ""}
          <button class="blue-btn action-btn" id="startDutyBtn"><span>+</span> Eintragen</button>
          <button class="red-btn action-btn" id="stopDutyBtn"><span>${iconSvg("Profil")}</span> Austragen</button>
          ${canAccess("actions", "stopAllDuty", "Direktion") ? `<button class="orange-btn action-btn" id="stopAllDutyBtn"><span>${iconSvg("Mitglieder")}</span> Alle Austragen</button>` : ""}
        </div>
      </div>
      ${renderDutyTable()}
    </section>
  `;

  $("#defconBtn")?.addEventListener("click", openDefconModal);
  $("#addNoteBtn")?.addEventListener("click", () => openNoteModal());
  $("#startDutyBtn").addEventListener("click", openStartDutyModal);
  $("#switchDutyBtn")?.addEventListener("click", openSwitchDutyModal);
  $("#stopDutyBtn").addEventListener("click", () => openStopDutyModal(myDuty));
  $("#stopAllDutyBtn")?.addEventListener("click", openStopAllDutyModal);
}

function renderNote(note) {
  const className = note.priority.toLowerCase();
  return `
    <article class="note-card" data-note-id="${escapeHtml(note.id)}">
      <div class="note-top">
        <div class="note-title">
          <strong>${escapeHtml(note.title)}</strong>
          <span class="badge ${className}">${escapeHtml(note.priority)}</span>
        </div>
        ${canAccess("actions", "manageNotes", "Supervisor") ? `<div class="note-actions">
          <button class="mini-icon edit-note" data-note-id="${escapeHtml(note.id)}" title="Notiz bearbeiten">${actionIcon("edit")}</button>
          <button class="mini-icon danger delete-note" data-note-id="${escapeHtml(note.id)}" title="Notiz löschen">${actionIcon("delete")}</button>
        </div>` : ""}
      </div>
      <p>${escapeHtml(note.text)}</p>
      <small class="muted">${escapeHtml(note.authorName)} · ${formatDate(note.createdAt)} ${formatTime(note.createdAt)}</small>
    </article>
  `;
}

function defconClass(defcon) {
  const value = Number(String(defcon).replace(/\D/g, ""));
  if (value === 1) return "defcon-1";
  if (value === 2) return "defcon-2";
  if (value === 3) return "defcon-3";
  if (value === 4) return "defcon-4";
  return "defcon-5";
}

function renderDutyTable() {
  if (!state.duty.length) return `<p class="muted">Aktuell ist niemand im Dienst.</p>`;
  const sortedDuty = [...state.duty].sort((a, b) => {
    const userA = a.user || state.users.find((item) => item.id === a.userId) || {};
    const userB = b.user || state.users.find((item) => item.id === b.userId) || {};
    return Number(userB.rank || 0) - Number(userA.rank || 0)
      || fullName(userA).localeCompare(fullName(userB), "de");
  });
  return `
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Rang</th><th>Name</th><th>Dienststart</th><th>Telefon</th><th>Status</th><th>Aktionen</th>
          </tr>
        </thead>
        <tbody>
          ${sortedDuty.map((entry) => {
            const user = entry.user || state.users.find((item) => item.id === entry.userId);
            return `
              <tr>
                <td>${escapeHtml(rankLabel(user?.rank))}</td>
                <td><span class="member-name">${avatarMarkup(user, "sm")}${escapeHtml(fullName(user))}</span></td>
                <td>${formatTime(entry.startedAt)}</td>
                <td>${escapeHtml(user?.phone || "-")}</td>
                <td><span class="status-chip ${statusClass(entry.status)}">${escapeHtml(entry.status)}</span></td>
                <td><button class="agent-action remove-duty" data-user-id="${entry.userId}" title="Person austragen">${iconSvg("Profil")}</button></td>
              </tr>
            `;
          }).join("")}
        </tbody>
      </table>
    </div>
  `;
}

function statusClass(status) {
  if (status === "Innendienst") return "status-inside";
  if (status === "Admin Dienst") return "status-admin";
  if (status === "Außendienst") return "status-outside";
  if (status === "Undercover Dienst") return "status-undercover";
  return "";
}

function renderMembers() {
  const rows = [...state.users].sort((a, b) => b.rank - a.rank || a.lastName.localeCompare(b.lastName));
  const search = localStorage.getItem("lspd_members_search") || "";
  content.innerHTML = `
    <section class="panel">
      <div class="panel-header">
        <h3>Mitglieder</h3>
        <span class="muted">${rows.length} Einträge</span>
      </div>
      <div class="filter-row members-search-row">
        <input id="membersSearch" value="${escapeHtml(search)}" placeholder="Mitglied, DN, Rang oder Ausbildung suchen">
      </div>
      <div class="table-wrap">
        <table class="members-table">
          <thead>
            <tr>
              <th class="member-name-col text-left">Name</th>
              <th class="text-center">Telefon</th>
              <th class="text-left">DN</th>
              <th class="member-rank-col text-center">Rang</th>
              <th class="text-left">Beitritt</th>
              <th class="text-left">Letzte Beförderung</th>
              ${trainingGroups.map((group) => group.map((item, index) => `<th class="${index === 0 ? "training-group-start" : ""} text-center">${escapeHtml(item)}</th>`).join("")).join("")}
            </tr>
          </thead>
          <tbody>
            ${rows.map((user) => `
              <tr class="filterable-row ${user.id === state.currentUser?.id ? "member-row-self" : ""}">
                <td class="member-name-col text-left"><span class="member-name member-name-wrap">${avatarMarkup(user, "sm")}<span>${wrapNameForTable(fullName(user))}</span></span></td>
                <td class="text-center">${escapeHtml(user.phone)}</td>
                <td class="text-left">${escapeHtml(user.dn)}</td>
                <td class="member-rank-col text-center"><span class="rank-number" data-rank-label="${escapeHtml(rankLabel(user.rank))}">${escapeHtml(user.rank)}</span></td>
                <td class="text-left">${formatDate(user.joinedAt)}</td>
                <td class="text-left">${formatDate(user.lastPromotionAt)}</td>
                ${trainingGroups.map((group) => group.map((training) => {
                  const hasTraining = Boolean(user.trainings?.[training]);
                  return `<td class="text-center ${hasTraining ? "training-yes" : "training-no"}">${hasTraining ? "✓" : "X"}</td>`;
                }).join("")).join("")}
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>
    </section>
  `;
  setupTableFilter("#membersSearch");
  $("#membersSearch")?.addEventListener("input", (event) => localStorage.setItem("lspd_members_search", event.target.value));
  if (search) $("#membersSearch")?.dispatchEvent(new Event("input"));
}

function renderInformation() {
  const links = state.settings.informationLinks || [];
  const permits = state.settings.informationPermits || [];
  const factions = state.settings.informationFactions || [];
  content.innerHTML = `
    <section class="department-info-view information-admin-view">
      <div class="info-box full information-card">
        <div class="department-modal-heading">
          <h4>Rechte Definition</h4>
          ${canAccess("actions", "manageInformation", "Direktion") ? `<button class="blue-btn" id="editInformationRights">${actionIcon("edit")} Bearbeiten</button>` : ""}
        </div>
        <div class="rich-text-view">${formatDepartmentText(state.settings.informationRightsText)}</div>
      </div>
      <div class="info-box full information-card redirects-card">
        <div class="department-modal-heading">
          <h4>Weiterleitungen</h4>
          ${canAccess("actions", "manageInformation", "Direktion") ? `<button class="blue-btn" id="addInformationLink">${iconSvg("Plus")} Hinzufügen</button>` : ""}
        </div>
        <div class="link-card-grid">${links.map((link) => `
          <article class="small-link-card">
            <strong>${escapeHtml(link.title)}</strong>
            <span class="link-label">Link:</span>
            <a href="${escapeHtml(link.url)}" target="_blank" rel="noreferrer">${escapeHtml(link.url)}</a>
            ${canAccess("actions", "manageInformation", "Direktion") ? `<span class="button-row"><button class="blue-btn compact-action edit-info-link" data-id="${link.id}" title="Bearbeiten">${actionIcon("edit")} Bearbeiten</button><button class="mini-icon danger delete-info-link" data-id="${link.id}" title="Löschen">${actionIcon("delete")}</button></span>` : ""}
          </article>
        `).join("") || `<p class="muted">Noch keine Weiterleitungen.</p>`}</div>
      </div>
      <div class="info-box full information-card">
        <div class="department-modal-heading">
          <h4>Sondergenehmigungen</h4>
          ${canAccess("actions", "manageInformation", "Direktion") ? `<button class="blue-btn" id="addInformationPermit">${iconSvg("Plus")} Hinzufügen</button>` : ""}
        </div>
        <div class="table-wrap compact-table">
          <table>
            <thead><tr><th>Vor- und Nachname</th><th>Beschreibung</th><th>Gültig Bis</th><th>Aktionen</th></tr></thead>
            <tbody>${permits.map((permit) => `<tr><td>${escapeHtml(permit.name)}</td><td>${escapeHtml(permit.description)}</td><td>${formatDate(permit.validUntil)}</td><td>${canAccess("actions", "manageInformation", "Direktion") ? `<button class="mini-icon edit-info-permit" data-id="${permit.id}">${actionIcon("edit")}</button><button class="mini-icon danger delete-info-permit" data-id="${permit.id}">${actionIcon("delete")}</button>` : ""}</td></tr>`).join("") || `<tr><td colspan="4" class="muted">Keine Sondergenehmigungen.</td></tr>`}</tbody>
          </table>
        </div>
      </div>
      <div class="info-box full information-card">
        <div class="department-modal-heading">
          <h4>Fraktionen</h4>
          ${canAccess("actions", "manageInformation", "Direktion") ? `<button class="blue-btn" id="addInformationFaction">${iconSvg("Plus")} Hinzufügen</button>` : ""}
        </div>
        <div class="table-wrap compact-table">
          <table>
            <thead><tr><th>Organisation</th><th>Status</th><th>Aktionen</th></tr></thead>
            <tbody>${factions.map((faction) => `<tr><td>${escapeHtml(faction.organization)}</td><td><span class="status-label">${renderStatusDot(faction.status)}</span></td><td>${canAccess("actions", "manageInformation", "Direktion") ? `<button class="mini-icon edit-info-faction" data-id="${faction.id}">${actionIcon("edit")}</button><button class="mini-icon danger delete-info-faction" data-id="${faction.id}">${actionIcon("delete")}</button>` : ""}</td></tr>`).join("") || `<tr><td colspan="3" class="muted">Keine Fraktionen.</td></tr>`}</tbody>
          </table>
        </div>
      </div>
    </section>
    <section class="panel">
      <div class="panel-header">
        <h3><span class="section-icon">${iconSvg("Informationen")}</span>Informationen</h3>
        ${canAccess("actions", "manageInformation", "Direktion") ? `<button class="blue-btn" id="editInformation">${actionIcon("edit")} Bearbeiten</button>` : ""}
      </div>
      <div class="info-box">
        <strong>Bewerbungsstatus</strong>
        <p><span class="application-pill ${state.settings.applicationStatus === "Offen" ? "open" : "closed"}">${escapeHtml(state.settings.applicationStatus)}</span></p>
      </div>
      <div class="info-box">
        <strong>Beschreibung</strong>
        <p>${escapeHtml(state.settings.informationText)}</p>
      </div>
    </section>
  `;
  $("#editInformation")?.addEventListener("click", openInformationEditModal);
  $("#editInformationRights")?.addEventListener("click", openInformationRightsModal);
  $("#addInformationLink")?.addEventListener("click", () => openInformationLinkModal());
  $("#addInformationPermit")?.addEventListener("click", () => openInformationPermitModal());
  $("#addInformationFaction")?.addEventListener("click", () => openInformationFactionModal());
  document.querySelectorAll(".edit-info-link").forEach((button) => button.addEventListener("click", () => openInformationLinkModal(links.find((item) => item.id === button.dataset.id))));
  document.querySelectorAll(".delete-info-link").forEach((button) => button.addEventListener("click", () => deleteInformationItem("informationLinks", button.dataset.id)));
  document.querySelectorAll(".edit-info-permit").forEach((button) => button.addEventListener("click", () => openInformationPermitModal(permits.find((item) => item.id === button.dataset.id))));
  document.querySelectorAll(".delete-info-permit").forEach((button) => button.addEventListener("click", () => deleteInformationItem("informationPermits", button.dataset.id)));
  document.querySelectorAll(".edit-info-faction").forEach((button) => button.addEventListener("click", () => openInformationFactionModal(factions.find((item) => item.id === button.dataset.id))));
  document.querySelectorAll(".delete-info-faction").forEach((button) => button.addEventListener("click", () => deleteInformationItem("informationFactions", button.dataset.id)));
}

function openInformationEditModal() {
  openModal(`
    <h3>Informationen bearbeiten</h3>
    <label>Beschreibung<textarea id="informationText">${escapeHtml(state.settings.informationText)}</textarea></label>
    <label>Bewerbungsstatus
      <select id="informationApplicationStatus">
        <option ${state.settings.applicationStatus === "Offen" ? "selected" : ""}>Offen</option>
        <option ${state.settings.applicationStatus === "Geschlossen" ? "selected" : ""}>Geschlossen</option>
      </select>
    </label>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions">
      <button class="ghost-btn" data-close>Abbrechen</button>
      <button class="blue-btn" id="saveInformation">Speichern</button>
    </div>
  `, (modal) => {
    modal.querySelector("#saveInformation").addEventListener("click", async () => {
      try {
        await saveInformationPatch({ informationText: $("#informationText").value, applicationStatus: $("#informationApplicationStatus").value });
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
  });
}

async function saveInformationPatch(patch) {
  await api("/api/information", {
    method: "PATCH",
    body: JSON.stringify({
      informationText: state.settings.informationText,
      applicationStatus: state.settings.applicationStatus,
      informationRightsText: state.settings.informationRightsText || "",
      informationLinks: state.settings.informationLinks || [],
      informationDocs: state.settings.informationDocs || [],
      informationDocChanges: state.settings.informationDocChanges || [],
      informationPermits: state.settings.informationPermits || [],
      informationFactions: state.settings.informationFactions || [],
      ...patch
    })
  });
  closeModal();
  await bootstrap();
}

function openInformationRightsModal() {
  openModal(`
    <h3>Rechte Definition bearbeiten</h3>
    <label>Text<textarea id="informationRightsText" rows="12">${escapeHtml(state.settings.informationRightsText || "")}</textarea></label>
    <p class="muted">Überschriften mit ##, dicke Schrift mit **Text**.</p>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions"><button class="ghost-btn" data-close>Abbrechen</button><button class="blue-btn" id="saveInformationRights">Speichern</button></div>
  `, (modal) => {
    modal.querySelector("#saveInformationRights").addEventListener("click", async () => {
      try {
        await saveInformationPatch({ informationRightsText: $("#informationRightsText").value });
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
  });
}

function openInformationLinkModal(link = null) {
  openModal(`
    <h3>${link ? "Weiterleitung bearbeiten" : "Weiterleitung hinzufügen"}</h3>
    <label>Titel<input id="informationLinkTitle" value="${escapeHtml(link?.title || "")}"></label>
    <label>Link<input id="informationLinkUrl" value="${escapeHtml(link?.url || "")}" placeholder="https://..."></label>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions"><button class="ghost-btn" data-close>Abbrechen</button><button class="blue-btn" id="saveInformationLink">Speichern</button></div>
  `, (modal) => {
    modal.querySelector("#saveInformationLink").addEventListener("click", async () => {
      try {
        const title = $("#informationLinkTitle").value.trim();
        const url = $("#informationLinkUrl").value.trim();
        if (!title || !url) throw new Error("Titel und Link sind erforderlich.");
        await saveInformationPatch({ informationLinks: upsertById(state.settings.informationLinks, { id: link?.id, title, url }) });
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
  });
}

function openInformationPermitModal(permit = null) {
  openModal(`
    <h3>${permit ? "Sondergenehmigung bearbeiten" : "Sondergenehmigung hinzufügen"}</h3>
    <label>Vor- und Nachname<input id="informationPermitName" value="${escapeHtml(permit?.name || "")}"></label>
    <label>Beschreibung<textarea id="informationPermitDescription">${escapeHtml(permit?.description || "")}</textarea></label>
    <label>Gültig Bis<input id="informationPermitValidUntil" type="date" value="${escapeHtml(permit?.validUntil || "")}"></label>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions"><button class="ghost-btn" data-close>Abbrechen</button><button class="blue-btn" id="saveInformationPermit">Speichern</button></div>
  `, (modal) => {
    modal.querySelector("#saveInformationPermit").addEventListener("click", async () => {
      try {
        const name = $("#informationPermitName").value.trim();
        const description = $("#informationPermitDescription").value.trim();
        const validUntil = $("#informationPermitValidUntil").value;
        if (!name || !description || !validUntil) throw new Error("Alle Felder sind erforderlich.");
        await saveInformationPatch({ informationPermits: upsertById(state.settings.informationPermits, { id: permit?.id, name, description, validUntil }) });
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
  });
}

function openInformationFactionModal(faction = null) {
  openModal(`
    <h3>${faction ? "Fraktion bearbeiten" : "Fraktion hinzufügen"}</h3>
    <label>Organisation<input id="informationFactionOrganization" value="${escapeHtml(faction?.organization || "")}"></label>
    <label>Status
      <select id="informationFactionStatus">
        ${["Normal", "Mittel", "Hoch"].map((status) => `<option ${faction?.status === status ? "selected" : ""}>${status}</option>`).join("")}
      </select>
    </label>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions"><button class="ghost-btn" data-close>Abbrechen</button><button class="blue-btn" id="saveInformationFaction">Speichern</button></div>
  `, (modal) => {
    modal.querySelector("#saveInformationFaction").addEventListener("click", async () => {
      try {
        const organization = $("#informationFactionOrganization").value.trim();
        const status = $("#informationFactionStatus").value;
        if (!organization) throw new Error("Organisation ist erforderlich.");
        await saveInformationPatch({ informationFactions: upsertById(state.settings.informationFactions, { id: faction?.id, organization, status }) });
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
  });
}

async function deleteInformationItem(key, id) {
  try {
    await saveInformationPatch({ [key]: (state.settings[key] || []).filter((item) => item.id !== id) });
  } catch (error) {
    showNotify(error.message, "error");
  }
}

function openDeleteInformationConfirm(key, id, title = "Eintrag löschen?") {
  openModal(`
    <h3>${escapeHtml(title)}</h3>
    <p class="muted">Dieser Eintrag wird dauerhaft entfernt.</p>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions">
      <button class="ghost-btn" data-close>Abbrechen</button>
      <button class="red-btn" id="confirmInformationDelete">Löschen</button>
    </div>
  `, (modal) => {
    modal.querySelector("#confirmInformationDelete").addEventListener("click", async () => {
      try {
        await deleteInformationItem(key, id);
        closeModal();
        renderInformation();
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
  });
}

function renderDirektion() {
  if (!canSeeDepartment("Direktion")) {
    content.innerHTML = `<section class="panel"><h3>Kein Zugriff</h3><p class="muted">Dieser Bereich ist nur für die Direktion sichtbar.</p></section>`;
    return;
  }
  const directionDepartment = state.departments.find((department) => department.id === "direktion");
  const directionMembersCount = directionDepartment?.members.length || 0;
  const activeStrikeCount = (state.disciplinary || []).filter((entry) => (entry.type === "Strike" || entry.sanctionType === "Strike") && !entry.archivedAt && (!entry.expiresAt || new Date(entry.expiresAt) > new Date())).length;
  const monthStart = new Date(new Date().getFullYear(), new Date().getMonth(), 1);
  const monthFines = (state.disciplinary || []).filter((entry) => entry.sanctionType === "Geldstrafe" && new Date(entry.createdAt) >= monthStart);
  const openFines = monthFines.filter((entry) => !entry.paidAt);
  const monthFineAmount = monthFines.reduce((sum, entry) => sum + Number(entry.amount || 0), 0);
  const tabs = [
    ["overview", "Übersicht"],
    ["members", "Mitglieder Verwaltung"],
    ["fluctuation", "Mitgliederfluktation"],
    ["upranks", "Upranks"],
    ["uprankRules", "Uprank Voraussetzungen"],
    ["hours", "Dienstzeiten"],
    ["logs", "Logs"]
  ];

  content.innerHTML = `
    <section class="internal-subhead department-overview-head">
      <h2>Direktion Abteilung</h2>
      <div class="department-control-row">
        <div class="tabs-row direction-tabs">
          ${tabs.map(([id, label]) => `<button class="${state.directionTab === id ? "tab-active" : ""}" data-direction-tab="${id}">${label}</button>`).join("")}
        </div>
        <button class="blue-btn vote-btn">${iconSvg("Abteilungen")} Abstimmung</button>
      </div>
      ${state.directionTab === "overview" ? `
      <div class="direction-overview-focus">
        <article class="direction-focus-card">
          <span>Mitglieder</span>
          <strong>${state.users.length}</strong>
          <small>Aktive Mitglieder im Dienstblatt</small>
        </article>
        <article class="direction-focus-card muted-card">
          <span>Abmeldungen</span>
          <strong>0</strong>
          <small>Counter vorbereitet</small>
        </article>
        <article class="direction-focus-card strike-card">
          <span>Aktive Strikes</span>
          <strong>${activeStrikeCount}</strong>
          <small>Nicht archiviert oder abgelaufen</small>
        </article>
        <article class="direction-focus-card fine-card">
          <span>Geldstrafen Monat</span>
          <strong>${monthFines.length}</strong>
          <small>${openFines.length} offen / ${monthFineAmount.toLocaleString("de-DE")} $ gesamt</small>
        </article>
      </div>
      ${renderDirectionDepartmentContent(directionDepartment, directionMembersCount)}
      ` : ""}
      ${state.directionTab === "members" ? renderDirectionMembersPanel() : ""}
      ${state.directionTab === "fluctuation" ? renderDirectionFluctuationPanel() : ""}
      ${state.directionTab === "upranks" ? renderDirectionUpranksPanel() : ""}
      ${state.directionTab === "uprankRules" ? renderDirectionUprankRulesPanel() : ""}
      ${state.directionTab === "hours" ? renderDirectionHoursPanel() : ""}
      ${state.directionTab === "logs" ? renderLogsPanel() : ""}
    </section>
  `;

  document.querySelectorAll("[data-direction-tab]").forEach((button) => {
    button.addEventListener("click", () => {
      state.directionTab = button.dataset.directionTab;
      localStorage.setItem("lspd_direction_tab", state.directionTab);
      renderDirektion();
    });
  });
  $("#createUserBtn")?.addEventListener("click", () => openUserModal());
  $("#addManualDutyBtn")?.addEventListener("click", openManualDutyModal);
  $("#hoursUserSelect")?.addEventListener("change", (event) => {
    localStorage.setItem("lspd_hours_user", event.target.value);
    renderDirektion();
  });
  document.querySelectorAll(".department-add").forEach((button) => button.addEventListener("click", () => openDepartmentMemberModal(directionDepartment)));
  document.querySelectorAll(".dept-note-add").forEach((button) => button.addEventListener("click", () => openDepartmentNoteModal(directionDepartment)));
  document.querySelectorAll(".user-actions").forEach((button) => button.addEventListener("click", () => openUserActionsModal(state.users.find((user) => user.id === button.dataset.id))));
  document.querySelectorAll(".archive-delete").forEach((button) => button.addEventListener("click", () => openDeleteUserModal(button.dataset.id)));
  document.querySelectorAll(".archive-rehire").forEach((button) => button.addEventListener("click", () => openRehireUserModal(findAnyUser(button.dataset.id))));
  document.querySelectorAll(".uprank-run").forEach((button) => button.addEventListener("click", () => openUprankModal(findAnyUser(button.dataset.id), button.dataset.special === "true")));
  document.querySelectorAll(".uprank-shorten").forEach((button) => button.addEventListener("click", () => openUprankAdjustmentModal(findAnyUser(button.dataset.id), "Verkürzung")));
  document.querySelectorAll(".uprank-special").forEach((button) => button.addEventListener("click", () => openUprankAdjustmentModal(findAnyUser(button.dataset.id), "Sonderuprank")));
  $("#uprankSearch")?.addEventListener("input", (event) => {
    localStorage.setItem("lspd_uprank_search", event.target.value);
    updateUprankList();
  });
  $("#uprankRulesForm")?.addEventListener("submit", saveUprankRules);
  bindFluctuationActions();
  setupTableFilter("#logSearch");
  setupTableFilter("#hoursSearch");
  setupTableFilter("#directionFluctuationSearch");
}

function renderDirectionDepartmentContent(department, memberCount = department?.members?.length || 0) {
  if (!department) return "";
  const canMembers = departmentActionAllowed(department, "departmentMembers");
  const canNotes = departmentActionAllowed(department, "departmentNotes");
  return `
    <div class="department-layout department-overview-content">
      <div class="panel">
        <div class="panel-header">
          <h3><span class="section-icon">${iconSvg("Mitglieder")}</span>Abteilungsmitglieder <span class="heading-count">${memberCount}</span></h3>
          ${canMembers ? `<button class="blue-btn department-add" data-department-id="${escapeHtml(department.id)}">${iconSvg("Mitglieder")} Person hinzufügen</button>` : ""}
        </div>
        ${renderDepartmentMemberTable(department)}
      </div>
      <div class="panel">
        <div class="panel-header">
          <h3><span class="section-icon">${iconSvg("Einsatzzentrale")}</span>Notizen</h3>
          ${canNotes ? `<button class="blue-btn dept-note-add" data-department-id="${escapeHtml(department.id)}">+ Neue Notiz</button>` : ""}
        </div>
        <div class="note-list">
          ${department.notes.length ? department.notes.map((note) => renderDepartmentNote(department, note)).join("") : `<p class="muted">Noch keine Notizen vorhanden.</p>`}
        </div>
      </div>
    </div>
  `;
}

function renderDirectionMembersPanel() {
  const archiveRows = state.archivedUsers || [];
  return `
    <div class="panel department-overview-content">
      <div class="panel-header">
        <h3>Mitgliederverwaltung</h3>
        <button class="blue-btn" id="createUserBtn">Neues Mitglied einstellen</button>
      </div>
      <div class="table-wrap">
        <table>
          <thead><tr><th>Name</th><th>Telefon</th><th>DN</th><th>Rang</th><th>Rolle</th><th>Status</th><th>Aktionen</th></tr></thead>
          <tbody>
            ${state.users.map((user) => `
              <tr class="${userStatusRowClass(user)}">
                <td><strong>${escapeHtml(fullName(user))}</strong><small class="table-subline">Einstellung: ${formatDate(user.joinedAt)}</small></td>
                <td>${escapeHtml(user.phone)}</td>
                <td>${escapeHtml(user.dn)}</td>
                <td>${escapeHtml(rankLabel(user.rank))}</td>
                <td>${roleBadges(user)}</td>
                <td>${renderAccountStatus(user)}</td>
                <td><button class="mini-icon user-actions" data-id="${user.id}" title="Aktionen">${actionIcon("edit")}</button></td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>
    </div>
    <div class="panel department-overview-content">
      <div class="panel-header">
        <h3>Archiv</h3>
        <span class="muted">${archiveRows.length} entlassene Accounts</span>
      </div>
      <div class="table-wrap">
        <table>
          <thead><tr><th>Name</th><th>Kündigungsgrund</th><th>Alte DN</th><th>Alter Rang</th><th>Alte Ausbildungen</th><th>Entlassen am</th><th>Aktionen</th></tr></thead>
          <tbody>
            ${archiveRows.map((user) => {
              const info = terminationInfo(user);
              return `
              <tr>
                <td>${escapeHtml(fullName(user))}</td>
                <td>${escapeHtml(info.reason || "-")}</td>
                <td>${escapeHtml(info.oldDn || "-")}</td>
                <td>${escapeHtml(rankLabel(info.oldRank))}</td>
                <td class="archive-training-list">${renderTrainingSummary(info.oldTrainings)}</td>
                <td>${formatDateTime(info.terminatedAt)}</td>
                <td>
                  <div class="button-row">
                    <button class="blue-btn archive-rehire" data-id="${escapeHtml(user.id)}">Wiedereinstellen</button>
                    <button class="mini-icon danger archive-delete" data-id="${escapeHtml(user.id)}" title="Löschen">${actionIcon("delete")}</button>
                  </div>
                </td>
              </tr>`;
            }).join("") || `<tr><td colspan="7" class="muted">Noch keine entlassenen Personen im Archiv.</td></tr>`}
          </tbody>
        </table>
      </div>
    </div>
  `;
}

function terminationInfo(user) {
  const fallback = (state.settings.fluctuation || []).find((row) => row.userId === user.id && row.type === "Kündigung") || {};
  return {
    reason: user.termination?.reason || fallback.reason || "",
    oldRank: user.termination?.oldRank ?? fallback.rank ?? user.rank,
    oldDn: user.termination?.oldDn ?? fallback.dn ?? user.dn,
    oldTrainings: user.termination?.oldTrainings || user.trainings || {},
    terminatedAt: user.termination?.terminatedAt || fallback.createdAt || user.updatedAt
  };
}

function userAccountStatus(user) {
  return user?.accountStatus || (user?.terminated ? "Entlassen" : user?.locked ? "Gesperrt" : "Aktiv");
}

function userStatusRowClass(user) {
  const status = userAccountStatus(user);
  if (status === "Suspendiert") return "member-row-suspended";
  if (status === "Gesperrt") return "member-row-locked";
  return "";
}

function renderAccountStatus(user) {
  const status = userAccountStatus(user);
  const className = status === "Aktiv" ? "active" : status === "Suspendiert" ? "suspended" : "locked";
  return `<span class="account-status-chip ${className}">${escapeHtml(status)}</span>`;
}

function dnConflictFor(dn, currentUserId = "") {
  const value = String(dn || "").trim();
  if (!value) return null;
  return [...(state.users || []), ...(state.archivedUsers || [])].find((item) => item.id !== currentUserId && String(item.dn || "") === value);
}

function renderDnConflictBox(holder, dn) {
  if (!holder) return "";
  const info = terminationInfo(holder);
  const status = userAccountStatus(holder);
  return `
    <div class="info-box full dn-conflict-box">
      <strong>Dienstnummer bereits vergeben</strong>
      <p>DN ${escapeHtml(dn)} ist vergeben an ${escapeHtml(fullName(holder))} - Status: ${escapeHtml(status)}${holder.terminated ? ` - Entlassen am ${formatDateTime(info.terminatedAt)}` : ""}</p>
      ${holder.terminated ? `<label class="checkbox-line">Dienstnummer überschreiben und beim archivierten Account entfernen<input type="checkbox" id="overwriteDn"></label>` : `<p class="form-error">Aktive Mitglieder können nicht überschrieben werden.</p>`}
    </div>
  `;
}

function renderTrainingSummary(trainingsMap = {}) {
  const done = trainings.filter((training) => trainingsMap?.[training]);
  return done.length ? done.map((training) => `<span class="training-mini">${escapeHtml(training)}</span>`).join("") : `<span class="muted">Keine</span>`;
}

function renderTrainingPicker(selectedTrainings = {}) {
  return `
    <div class="training-picker">
      ${trainingGroups.map((group, index) => `
        <section class="training-picker-group">
          <div class="training-picker-title">${["Grundausbildung", "Führung / EL", "Spezialisierungen"][index] || "Ausbildungen"}</div>
          <div class="training-picker-grid">
            ${group.map((training) => `
              <label class="training-toggle">
                <input type="checkbox" name="training_${training}" ${selectedTrainings[training] ? "checked" : ""}>
                <span>${escapeHtml(training)}</span>
              </label>
            `).join("")}
          </div>
        </section>
      `).join("")}
    </div>
  `;
}

function renderItToggleOld(checked = false) {
  return `
    <label class="it-toggle">
      <input type="checkbox" name="isIT" ${checked ? "checked" : ""}>
      <span class="it-toggle-ui"></span>
      <span><b>IT-Rolle</b><small>Zusätzliche Systemrechte vergeben</small></span>
    </label>
  `;
}

function editableRoleOptions(user = null) {
  return state.roles.filter((role) => !["IT", "IT-Leitung"].includes(role));
}

function baseRoleForUser(user = null) {
  return user?.baseRole || (["IT", "IT-Leitung"].includes(user?.role) ? "Direktion" : user?.role || "User");
}

function renderItRoleControls(user = null) {
  const isIt = ["IT", "IT-Leitung"].includes(user?.role);
  const isLead = user?.role === "IT-Leitung";
  const disabled = canGrantItRoles() ? "" : "disabled";
  return `
    <div class="it-role-controls full">
      <label class="it-toggle">
        <input type="checkbox" name="isIT" ${isIt ? "checked" : ""} ${disabled}>
        <span class="it-toggle-ui"></span>
        <span><b>IT</b><small>Zusätzliche IT-Rechte</small></span>
      </label>
      <label class="it-toggle">
        <input type="checkbox" name="isITLead" ${isLead ? "checked" : ""} ${disabled}>
        <span class="it-toggle-ui"></span>
        <span><b>IT-Leitung</b><small>Darf IT-Rollen vergeben</small></span>
      </label>
    </div>
  `;
}

function renderTeamlerControl(user = null) {
  return `
    <label class="it-toggle">
      <input type="checkbox" name="teamler" ${user?.teamler ? "checked" : ""}>
      <span class="it-toggle-ui"></span>
      <span><b>Teamler</b><small>Darf Admin Dienst stempeln</small></span>
    </label>
  `;
}

function renderDirectionFluctuationPanel() {
  const rows = state.settings.fluctuation || [];
  const canManage = canManageFluctuation();
  return `
    <div class="panel department-overview-content">
      <div class="panel-header"><h3>Mitgliederfluktation</h3><input id="directionFluctuationSearch" class="compact-input" placeholder="Suchen"></div>
      <div class="table-wrap">
        <table>
          <thead><tr><th>Name</th><th>DN</th><th>Rang</th><th>Bearbeitet von</th><th>Typ</th><th>Grund</th><th>Datum</th>${canManage ? "<th>Aktionen</th>" : ""}</tr></thead>
          <tbody>
            ${rows.map((row) => `
              <tr class="filterable-row">
                <td>${escapeHtml(row.name)}</td>
                <td>${escapeHtml(row.dn || "-")}</td>
                <td>${escapeHtml(rankLabel(row.rank))}</td>
                <td>${escapeHtml(row.actorName || "-")}</td>
                <td><span class="fluctuation-chip ${fluctuationTypeClass(row)}">${escapeHtml(row.type)}</span></td>
                <td>${escapeHtml(row.reason || "-")}</td>
                <td>${formatDateTime(row.createdAt)}</td>
                ${canManage ? `<td><span class="button-row"><button class="mini-icon edit-fluctuation" data-id="${escapeHtml(row.id)}" title="Bearbeiten">${actionIcon("edit")}</button><button class="mini-icon danger delete-fluctuation" data-id="${escapeHtml(row.id)}" title="Löschen">${actionIcon("delete")}</button></span></td>` : ""}
              </tr>
            `).join("") || `<tr><td colspan="${canManage ? 8 : 7}" class="muted">Noch keine Einträge.</td></tr>`}
          </tbody>
        </table>
      </div>
    </div>
  `;
}

function isDismissedFluctuation(row) {
  return row?.type === "Kündigung" || row?.type === "KÃ¼ndigung";
}

function fluctuationTypeClass(row) {
  return row?.type === "Eingestellt" ? "hired" : "dismissed";
}

function fluctuationById(id) {
  return (state.settings.fluctuation || []).find((row) => row.id === id);
}

function datetimeLocalValue(value) {
  if (!value) return "";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "";
  const local = new Date(date.getTime() - date.getTimezoneOffset() * 60000);
  return local.toISOString().slice(0, 16);
}

function bindFluctuationActions() {
  if (!canManageFluctuation()) return;
  document.querySelectorAll(".edit-fluctuation").forEach((button) => {
    button.addEventListener("click", () => openFluctuationModal(fluctuationById(button.dataset.id)));
  });
  document.querySelectorAll(".delete-fluctuation").forEach((button) => {
    button.addEventListener("click", () => openDeleteFluctuationModal(fluctuationById(button.dataset.id)));
  });
}

function openFluctuationModal(row) {
  if (!row || !canManageFluctuation()) return;
  const selectedType = isDismissedFluctuation(row) ? "Kündigung" : "Eingestellt";
  openModal(`
    <h3>Fluktuationseintrag bearbeiten</h3>
    <form id="fluctuationForm" class="modal-form">
      <label>Name<input id="fluctuationName" required value="${escapeHtml(row.name || "")}"></label>
      <label>DN<input id="fluctuationDn" value="${escapeHtml(row.dn || "")}"></label>
      <label>Rang<select id="fluctuationRank">${state.ranks.map((rank) => `<option value="${rank.level}" ${Number(row.rank) === Number(rank.level) ? "selected" : ""}>${escapeHtml(rankOptionLabel(rank))}</option>`).join("")}</select></label>
      <label>Bearbeitet von<input id="fluctuationActor" value="${escapeHtml(row.actorName || "")}"></label>
      <label>Typ<select id="fluctuationType"><option ${selectedType === "Eingestellt" ? "selected" : ""}>Eingestellt</option><option ${selectedType === "Kündigung" ? "selected" : ""}>Kündigung</option></select></label>
      <label>Grund<textarea id="fluctuationReason" rows="4">${escapeHtml(row.reason || "")}</textarea></label>
      <label>Datum<input id="fluctuationCreatedAt" type="datetime-local" value="${datetimeLocalValue(row.createdAt)}"></label>
      <div class="modal-actions">
        <button type="button" class="ghost-btn" onclick="closeModal()">Abbrechen</button>
        <button class="blue-btn">Speichern</button>
      </div>
    </form>
  `, () => {
    $("#fluctuationForm").addEventListener("submit", async (event) => {
      event.preventDefault();
      const data = await api(`/api/settings/fluctuation/${row.id}`, {
        method: "PATCH",
        body: JSON.stringify({
          name: $("#fluctuationName").value,
          dn: $("#fluctuationDn").value,
          rank: Number($("#fluctuationRank").value),
          actorName: $("#fluctuationActor").value,
          type: $("#fluctuationType").value,
          reason: $("#fluctuationReason").value,
          createdAt: $("#fluctuationCreatedAt").value
        })
      });
      state.settings.fluctuation = data.fluctuation || state.settings.fluctuation;
      closeModal();
      renderApp();
    });
  });
}

function openDeleteFluctuationModal(row) {
  if (!row || !canManageFluctuation()) return;
  openModal(`
    <h3>Fluktuationseintrag löschen</h3>
    <p class="muted">Dieser Eintrag wird dauerhaft aus Direktion und aus dem Reiter Mitgliederfluktation entfernt.</p>
    <div class="profile-summary">
      <strong>${escapeHtml(row.name || "-")}</strong>
      <span>${escapeHtml(row.type || "-")} · ${formatDateTime(row.createdAt)}</span>
    </div>
    <div class="modal-actions">
      <button type="button" class="ghost-btn" onclick="closeModal()">Abbrechen</button>
      <button id="confirmDeleteFluctuation" class="red-btn">Löschen</button>
    </div>
  `, () => {
    $("#confirmDeleteFluctuation").addEventListener("click", async () => {
      const data = await api(`/api/settings/fluctuation/${row.id}`, { method: "DELETE" });
      state.settings.fluctuation = data.fluctuation || [];
      closeModal();
      renderApp();
    });
  });
}

function renderDirectionHoursPanel() {
  const rows = state.dutyHistory || [];
  const selectedUserId = localStorage.getItem("lspd_hours_user") || "all";
  const scopedRows = selectedUserId === "all" ? rows : rows.filter((entry) => entry.userId === selectedUserId);
  const now = new Date();
  const dayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const weekStart = new Date(now);
  weekStart.setDate(now.getDate() - 7);
  const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
  const sumFrom = (from = null) => scopedRows.filter((entry) => !from || new Date(entry.startedAt) >= from).reduce((sum, entry) => sum + durationMs(entry), 0);
  return `
    <div class="panel department-overview-content">
      <div class="panel-header">
        <h3>Dienstzeiten Verwaltung</h3>
        <div class="button-row">
          <select id="hoursUserSelect" class="compact-input">
            <option value="all">Alle Mitglieder</option>
            ${state.users.map((user) => `<option value="${user.id}" ${selectedUserId === user.id ? "selected" : ""}>${escapeHtml(fullName(user))}</option>`).join("")}
          </select>
          <button class="blue-btn" id="addManualDutyBtn">Stunden hinzufügen</button>
        </div>
      </div>
      <div class="grid-4 compact-stats">
        <div class="stat-card"><span>Heute</span><strong>${formatDuration(sumFrom(dayStart))}</strong><small>${selectedUserId === "all" ? "Alle Mitglieder" : "Ausgewähltes Mitglied"}</small></div>
        <div class="stat-card"><span>Woche</span><strong>${formatDuration(sumFrom(weekStart))}</strong><small>Letzte 7 Tage</small></div>
        <div class="stat-card"><span>Monat</span><strong>${formatDuration(sumFrom(monthStart))}</strong><small>Aktueller Monat</small></div>
        <div class="stat-card"><span>Gesamt</span><strong>${formatDuration(sumFrom())}</strong><small>Alle Zeiten</small></div>
      </div>
      <div class="filter-row">
        <input id="hoursSearch" placeholder="Name, Diensttyp oder Status suchen">
      </div>
      <div class="table-wrap">
        <table>
          <thead><tr><th>Name</th><th>Dienstbeginn</th><th>Dienstende</th><th>Diensttyp</th><th>Dauer</th><th>Status</th></tr></thead>
          <tbody>
            ${scopedRows.map((entry) => `
              <tr class="filterable-row">
                <td>${escapeHtml(fullName(entry.user || findAnyUser(entry.userId)) || "-")}</td>
                <td>${formatDateTime(entry.startedAt)}</td>
                <td>${entry.endedAt ? formatDateTime(entry.endedAt) : "Läuft noch"}</td>
                <td>${escapeHtml(entry.status)}</td>
                <td>${formatDuration(durationMs(entry))}</td>
                <td>${entry.endedAt ? "Beendet" : "Aktiv"}</td>
              </tr>
            `).join("") || `<tr><td colspan="6" class="muted">Noch keine Dienstzeiten.</td></tr>`}
          </tbody>
        </table>
      </div>
    </div>
  `;
}

function uprankRules() {
  return Array.isArray(state.settings.uprankRules) ? state.settings.uprankRules : [];
}

function uprankAdjustmentsFor(userId, targetRank, type = "") {
  return (state.settings.uprankAdjustments || []).filter((item) =>
    item.userId === userId &&
    Number(item.targetRank) === Number(targetRank) &&
    (!type || item.type === type)
  );
}

function daysSince(dateValue) {
  const time = new Date(dateValue || Date.now()).getTime();
  return Math.max(0, Math.floor((Date.now() - time) / 86400000));
}

function dutySumForUser(userId, from) {
  return (state.dutyHistory || [])
    .filter((entry) => entry.userId === userId && (!from || new Date(entry.startedAt) >= from))
    .reduce((sum, entry) => sum + durationMs(entry), 0);
}

function evaluateUprank(user) {
  const currentRank = Number(user.rank || 0);
  const targetRank = currentRank + 1;
  const rule = uprankRules().find((item) => Number(item.targetRank) === targetRank) || { targetRank, minDays: 14, trainings: [], specialOnly: targetRank >= 7 };
  const reduction = uprankAdjustmentsFor(user.id, targetRank, "Verkürzung").reduce((sum, item) => sum + Number(item.days || 0), 0);
  const effectiveDays = Math.max(0, Number(rule.minDays || 0) - reduction);
  const daysOnRank = daysSince(user.lastPromotionAt || user.joinedAt);
  const missingDays = Math.max(0, effectiveDays - daysOnRank);
  const missingTrainings = (rule.trainings || []).filter((training) => !user.trainings?.[training]);
  const hasSpecial = uprankAdjustmentsFor(user.id, targetRank, "Sonderuprank").length > 0;
  const hiddenSpecialRank = targetRank >= 10;
  const regularReady = missingDays === 0 && missingTrainings.length === 0;
  return {
    user,
    targetRank,
    rule,
    reduction,
    effectiveDays,
    daysOnRank,
    missingDays,
    missingTrainings,
    hasSpecial,
    hiddenSpecialRank,
    regularReady,
    ready: !hiddenSpecialRank && ((regularReady && !rule.specialOnly) || hasSpecial),
    needsSpecial: !hiddenSpecialRank && regularReady && rule.specialOnly && !hasSpecial
  };
}
function renderDirectionUpranksPanel() {
  const now = new Date();
  const weekStart = new Date(now);
  weekStart.setDate(now.getDate() - 7);
  const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
  const searchValue = localStorage.getItem("lspd_uprank_search") || "";
  const searchTerm = searchValue.trim().toLowerCase();
  const allRows = state.users
    .filter((user) => !user.terminated && Number(user.rank) < Math.max(...state.ranks.map((rank) => Number(rank.value))))
    .map(evaluateUprank);
  const visibleRows = allRows
    .filter((row) => {
      const searchText = `${fullName(row.user)} ${row.user.dn || ""} ${rankLabel(row.user.rank)} ${rankLabel(row.targetRank)} ${row.missingTrainings.join(" ")}`.toLowerCase();
      if (searchTerm) return searchText.includes(searchTerm);
      return row.ready || row.needsSpecial;
    })
    .sort((a, b) => {
      const specialA = a.rule.specialOnly || a.hasSpecial || a.needsSpecial;
      const specialB = b.rule.specialOnly || b.hasSpecial || b.needsSpecial;
      if (specialA !== specialB) return Number(specialA) - Number(specialB);
      return a.targetRank - b.targetRank || a.user.lastName.localeCompare(b.user.lastName);
    });
  const readyCount = allRows.filter((row) => row.ready).length;
  const specialCount = allRows.filter((row) => row.needsSpecial).length;
  return `
    <div class="panel department-overview-content">
      <div class="panel-header">
        <h3>Upranks</h3>
        <span class="muted">${readyCount} bereit \u00b7 ${specialCount} Sonderuprank n\u00f6tig</span>
      </div>
      <div class="uprank-search-row">
        <input id="uprankSearch" placeholder="Person suchen, um Uprank-Status zu pr\u00fcfen" value="${escapeHtml(searchValue)}">
        <small>${searchTerm ? "Suchmodus: alle passenden Personen werden angezeigt." : "Standard: nur berechtigte oder Sonderuprank-relevante Personen."}</small>
      </div>
      <div class="uprank-list">
        ${visibleRows.map((row) => renderUprankCard(row, weekStart, monthStart, Boolean(searchTerm))).join("") || `<p class="muted">${searchTerm ? "Keine Person gefunden." : "Keine Uprank-Kandidaten vorhanden."}</p>`}
      </div>
    </div>
  `;
}

function currentUprankRows(searchTerm = "") {
  const allRows = state.users
    .filter((user) => !user.terminated && Number(user.rank) < Math.max(...state.ranks.map((rank) => Number(rank.value))))
    .map(evaluateUprank);
  return allRows
    .filter((row) => {
      const searchText = `${fullName(row.user)} ${row.user.dn || ""} ${rankLabel(row.user.rank)} ${rankLabel(row.targetRank)} ${row.missingTrainings.join(" ")}`.toLowerCase();
      if (searchTerm) return searchText.includes(searchTerm);
      return row.ready || row.needsSpecial;
    })
    .sort((a, b) => {
      const specialA = a.rule.specialOnly || a.hasSpecial || a.needsSpecial;
      const specialB = b.rule.specialOnly || b.hasSpecial || b.needsSpecial;
      if (specialA !== specialB) return Number(specialA) - Number(specialB);
      return a.targetRank - b.targetRank || a.user.lastName.localeCompare(b.user.lastName);
    });
}

function updateUprankList() {
  const input = $("#uprankSearch");
  const list = document.querySelector(".uprank-list");
  if (!input || !list) return;
  const searchTerm = input.value.trim().toLowerCase();
  const now = new Date();
  const weekStart = new Date(now);
  weekStart.setDate(now.getDate() - 7);
  const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
  const rows = currentUprankRows(searchTerm);
  list.innerHTML = rows.map((row) => renderUprankCard(row, weekStart, monthStart, Boolean(searchTerm))).join("") || `<p class="muted">${searchTerm ? "Keine Person gefunden." : "Keine Uprank-Kandidaten vorhanden."}</p>`;
  const hint = document.querySelector(".uprank-search-row small");
  if (hint) hint.textContent = searchTerm ? "Suchmodus: alle passenden Personen werden angezeigt." : "Standard: nur berechtigte oder Sonderuprank-relevante Personen.";
  document.querySelectorAll(".uprank-run").forEach((button) => button.addEventListener("click", () => openUprankModal(findAnyUser(button.dataset.id), button.dataset.special === "true")));
  document.querySelectorAll(".uprank-shorten").forEach((button) => button.addEventListener("click", () => openUprankAdjustmentModal(findAnyUser(button.dataset.id), "Verkürzung")));
  document.querySelectorAll(".uprank-special").forEach((button) => button.addEventListener("click", () => openUprankAdjustmentModal(findAnyUser(button.dataset.id), "Sonderuprank")));
}

function renderUprankCard(row, weekStart, monthStart, isSearchMode) {
  const weekHours = formatDuration(dutySumForUser(row.user.id, weekStart));
  const monthHours = formatDuration(dutySumForUser(row.user.id, monthStart));
  const cardClass = row.ready ? "ready" : row.needsSpecial ? "special-required" : "";
  const daysText = row.effectiveDays === 0
    ? `${row.daysOnRank} Tage auf Rang${row.reduction ? ` \u00b7 ${row.reduction} Tage Verk\u00fcrzung` : ""}`
    : `${Math.min(row.daysOnRank, row.effectiveDays)}/${row.effectiveDays} Tage auf Rang${row.reduction ? ` \u00b7 ${row.reduction} Tage Verk\u00fcrzung` : ""}`;
  const statusText = row.hiddenSpecialRank
    ? "Nur auf Direktion/Sonderfreigabe"
    : row.ready
      ? "Uprank bereit"
      : row.needsSpecial
        ? "Sonderuprank n\u00f6tig"
        : "Noch nicht berechtigt";
  const canRun = row.ready || row.needsSpecial || row.hasSpecial;
  const missingItems = [
    row.hiddenSpecialRank ? "nur Sonderfreigabe möglich" : "",
    row.missingDays ? `${row.missingDays} Tage fehlen` : "",
    row.missingTrainings.length ? `Ausbildungen fehlen: ${row.missingTrainings.join(", ")}` : "",
    row.rule.specialOnly && !row.hasSpecial ? "Sonderuprank-Freigabe fehlt" : ""
  ].filter(Boolean);
  return `
    <article class="uprank-card ${cardClass}">
      <div class="uprank-main">
        <strong>${escapeHtml(fullName(row.user))}</strong>
        <span>${escapeHtml(rankLabel(row.user.rank))} \u2192 ${escapeHtml(rankLabel(row.targetRank))}</span>
        <small>${daysText}</small>
      </div>
      <div class="uprank-facts">
        <span class="requirement-chip ${row.hiddenSpecialRank ? "special" : row.missingDays ? "missing" : "ok"}">${row.hiddenSpecialRank ? "Kein Tages-System" : row.missingDays ? `${row.missingDays} Tage fehlen` : "Dauer erf\u00fcllt"}</span>
        <span class="requirement-chip ${row.missingTrainings.length ? "missing" : "ok"}">${row.missingTrainings.length ? `Fehlt: ${escapeHtml(row.missingTrainings.join(", "))}` : "Ausbildungen erf\u00fcllt"}</span>
        <span class="requirement-chip ${row.rule.specialOnly || row.hiddenSpecialRank ? "special" : "ok"}">${row.rule.specialOnly || row.hiddenSpecialRank ? "Nur Sonderuprank" : "Regul\u00e4r m\u00f6glich"}</span>
        <span class="requirement-chip ${row.ready ? "ok" : row.needsSpecial ? "special" : ""}">${statusText}</span>
        <span class="requirement-chip">Woche ${weekHours}</span>
        <span class="requirement-chip">Monat ${monthHours}</span>
        <span class="requirement-chip">${escapeHtml(userAccountStatus(row.user))}</span>
      </div>
      <div class="uprank-actions">
        ${canRun ? `<button class="blue-btn uprank-run" data-id="${escapeHtml(row.user.id)}" data-special="${row.hasSpecial || row.rule.specialOnly}">${row.ready ? "Befördern" : "Prüfen"}</button>` : isSearchMode ? `<div class="uprank-missing-box"><strong>Prüfen</strong><span>${escapeHtml(missingItems.join(" · ") || "Keine offene Voraussetzung gefunden.")}</span></div>` : ""}
        <button class="ghost-btn uprank-shorten" data-id="${escapeHtml(row.user.id)}">Verk\u00fcrzung</button>
        <button class="orange-btn uprank-special" data-id="${escapeHtml(row.user.id)}">Sonderuprank</button>
      </div>
    </article>
  `;
}
function renderDirectionUprankRulesPanel() {
  const rules = uprankRules();
  return `
    <form id="uprankRulesForm" class="panel department-overview-content">
      <div class="panel-header">
        <h3>Uprank Voraussetzungen</h3>
        <button class="blue-btn" type="submit">Voraussetzungen speichern</button>
      </div>
      <div class="uprank-rule-list">
        ${rules.map((rule) => `
          <section class="uprank-rule-card" data-target-rank="${rule.targetRank}">
            <div>
              <strong>${escapeHtml(rankLabel(Number(rule.targetRank) - 1))} → ${escapeHtml(rankLabel(rule.targetRank))}</strong>
              <small>Zielrang ${rule.targetRank}</small>
            </div>
            <label>Min. Tage auf Rang<input type="number" min="0" name="minDays_${rule.targetRank}" value="${Number(rule.minDays || 0)}"></label>
            <label class="it-toggle compact-rule-toggle">
              <input type="checkbox" name="specialOnly_${rule.targetRank}" ${rule.specialOnly ? "checked" : ""}>
              <span class="it-toggle-ui"></span>
              <span><b>Nur Sonderuprank</b><small>Reguläre Dauer/Ausbildung wird nicht automatisch vorgeschlagen.</small></span>
            </label>
            <div class="rule-training-grid">
              ${trainings.map((training) => `
                <label class="training-toggle">
                  <input type="checkbox" name="rule_${rule.targetRank}_${training}" ${rule.trainings?.includes(training) ? "checked" : ""}>
                  <span>${escapeHtml(training)}</span>
                </label>
              `).join("")}
            </div>
          </section>
        `).join("")}
      </div>
      <p id="uprankRulesError" class="form-error"></p>
    </form>
  `;
}

function renderLogsPanel() {
  const rows = state.logs || [];
  return `
    <div class="panel department-overview-content">
      <div class="panel-header"><h3>Website Logs</h3><span class="muted">${rows.length} Einträge</span></div>
      <div class="filter-row">
        <input id="logSearch" placeholder="Aktion, Person, Ziel oder Änderung suchen">
      </div>
      <div class="table-wrap">
        <table>
          <thead><tr><th>Zeit</th><th>Wer</th><th>Aktion</th><th>Ziel</th><th>Beschreibung</th></tr></thead>
          <tbody>
            ${rows.map((log) => `
              <tr class="filterable-row ${logTone(log.action)}">
                <td>${formatDateTime(log.createdAt)}</td>
                <td>${escapeHtml(log.actorName || "-")}</td>
                <td><span class="log-action-chip">${escapeHtml(cleanText(log.action))}</span></td>
                <td>${escapeHtml(cleanText(log.target || "-"))}</td>
                <td>${renderLogDetails(log)}</td>
              </tr>
            `).join("") || `<tr><td colspan="5" class="muted">Noch keine Logs.</td></tr>`}
          </tbody>
        </table>
      </div>
    </div>
  `;
}

function renderFluctuation() {
  const rows = state.settings.fluctuation || [];
  const selectedRange = localStorage.getItem("lspd_fluctuation_range") || "Monat";
  const from = rangeStart(selectedRange);
  const rangeRows = rows.filter((row) => !from || new Date(row.createdAt) >= from);
  const monthLabel = new Date().toLocaleDateString("de-DE", { month: "long", year: "numeric" });
  const grouped = rangeRows.reduce((acc, row) => {
    const key = new Date(row.createdAt).toLocaleDateString("de-DE", { month: "long", year: "numeric" });
    acc[key] ||= { hired: 0, dismissed: 0 };
    if (row.type === "Eingestellt") acc[key].hired += 1;
    if (isDismissedFluctuation(row)) acc[key].dismissed += 1;
    return acc;
  }, {});
  const summary = Object.entries(grouped).length ? Object.entries(grouped) : [[monthLabel, { hired: 0, dismissed: 0 }]];
  const totalHired = rangeRows.filter((row) => row.type === "Eingestellt").length;
  const totalDismissed = rangeRows.filter(isDismissedFluctuation).length;
  content.innerHTML = `
    <section class="panel fluctuation-summary">
      <div class="panel-header">
        <h3>Fluktuation Statistik</h3>
        <select id="fluctuationRange" class="compact-input">
          ${["Heute", "Woche", "Monat", "Gesamt"].map((range) => `<option ${selectedRange === range ? "selected" : ""}>${range}</option>`).join("")}
        </select>
      </div>
      <div class="fluctuation-summary-grid">
        <div class="fluctuation-total summary-green"><span>Einstellungen</span><strong>${totalHired}</strong></div>
        <div class="fluctuation-total summary-red"><span>Kündigungen</span><strong>${totalDismissed}</strong></div>
        <div class="fluctuation-summary-list">
          ${summary.map(([label, item]) => `
            <div class="fluctuation-summary-row">
              <strong>${escapeHtml(label)}</strong>
              <span class="summary-green">+ ${item.hired}</span>
              <span class="summary-red">- ${item.dismissed}</span>
            </div>
          `).join("")}
        </div>
      </div>
    </section>
    <section class="panel">
      <div class="panel-header"><h3>Mitgliederfluktation</h3><span class="muted">${rangeRows.length} Einträge</span></div>
      <div class="table-wrap">
        <table>
          <thead><tr><th>Name</th><th>Typ</th><th>Grund</th><th>Datum</th></tr></thead>
          <tbody>
            ${rangeRows.length ? rangeRows.map((row) => `
              <tr class="filterable-row">
                <td>${escapeHtml(row.name)}</td>
                <td><span class="fluctuation-chip ${fluctuationTypeClass(row)}">${escapeHtml(row.type)}</span></td>
                <td>${escapeHtml(row.reason || "-")}</td>
                <td>${formatDateTime(row.createdAt)}</td>
              </tr>
            `).join("") : `<tr><td colspan="4" class="muted">Noch keine Einträge.</td></tr>`}
          </tbody>
        </table>
      </div>
    </section>
  `;
  $("#fluctuationRange").addEventListener("change", (event) => {
    localStorage.setItem("lspd_fluctuation_range", event.target.value);
    renderFluctuation();
  });
}

function editableItPages() {
  return orderPages([...pages, ...adminPages, ...(state.customPages || []).map((page) => page.key), ...state.departments.filter((department) => department.id !== "direktion").map((department) => `dept:${department.id}`)]);
}

function orderPages(items) {
  const unique = [...new Set(items.filter(Boolean))];
  const order = state.settings?.pageOrder || [];
  const weight = new Map(order.map((page, index) => [page, index]));
  return unique.sort((a, b) => {
    const aWeight = weight.has(a) ? weight.get(a) : 10000 + unique.indexOf(a);
    const bWeight = weight.has(b) ? weight.get(b) : 10000 + unique.indexOf(b);
    return aWeight - bWeight;
  });
}

function isCustomPage(page) {
  return (state.customPages || []).some((item) => item.key === page);
}

function isInternalSheetPage(page) {
  return page === "IT" || page === "Direktion" || isDepartmentPage(page);
}

function permissionRule(area, key) {
  return state.settings.permissions?.[area]?.[key] || { roles: [], ranks: [], users: [] };
}

function isPageViewRestricted(page) {
  return !Boolean(permissionRule("pages", page).all);
}

function restrictedPageIcon(page) {
  if (!isPageViewRestricted(page)) return "";
  const icon = isInternalSheetPage(page) ? "Lock" : "EyeOff";
  const title = isInternalSheetPage(page) ? "Geschütztes internes Blatt" : "Ausgeblendeter Reiter";
  return `<span class="nav-hidden-eye" title="${title}">${iconSvg(icon)}</span>`;
}

function restrictedPageEditIcon(page) {
  if (!isPageViewRestricted(page)) return "";
  const icon = isInternalSheetPage(page) ? "Lock" : "EyeOff";
  const title = isInternalSheetPage(page) ? "Ansehen ist eingeschränkt" : "Reiter ist ausgeblendet";
  return `<span class="page-lock" title="${title}">${iconSvg(icon)}</span>`;
}

function departmentActionAllowed(department, action) {
  if (!department) return false;
  const key = `${action}:${department.id}`;
  const rule = state.settings?.permissions?.actions?.[key];
  if (!rule && action === "departmentLeadership") {
    if (hasRole("Direktion")) return true;
    const membership = department.members.find((member) => member.userId === state.currentUser?.id);
    return departmentLeaderPositionsFor(department).includes(membership?.position);
  }
  return rule ? canAccess("actions", key, "IT") : Boolean(department.canManage);
}

function departmentPositionsFor(department) {
  const positions = Array.isArray(department?.positions) && department.positions.length ? department.positions : state.departmentPositions;
  return [...new Set(positions || [])];
}

function departmentLeaderPositionsFor(department) {
  const positions = departmentPositionsFor(department);
  const fallback = positions.filter((position) => ["Direktion", "Leitung", "Stv. Leitung"].includes(position));
  const leaders = Array.isArray(department?.leaderPositions) ? department.leaderPositions.filter((position) => positions.includes(position)) : fallback;
  return [...new Set(leaders.length ? leaders : fallback)];
}

function defaultPositionColor(position) {
  if (position === "Direktion" || position === "Anwärter") return "green";
  if (position === "Leitung") return "red";
  if (position === "Stv. Leitung") return "orange";
  if (position === "Mitglied") return "blue";
  return "blue";
}

function positionColorFor(department, position) {
  const color = department?.positionColors?.[position] || defaultPositionColor(position);
  return ["green", "red", "orange", "blue"].includes(color) ? color : defaultPositionColor(position);
}

function positionPowerFor(department, position) {
  const positions = departmentPositionsFor(department);
  const index = positions.indexOf(position);
  return index === -1 ? 0 : positions.length - index;
}

function canGrantItRoles() {
  return state.currentUser?.role === "IT-Leitung";
}

function departmentTab(department) {
  const selected = state.departmentTabs?.[department.id] || "overview";
  if (selected === "members") return "overview";
  if (selected === "leadership" && !departmentActionAllowed(department, "departmentLeadership")) return "overview";
  if (selected === "estExam" && !isHumanResourcesDepartmentSheet(department)) return "overview";
  if (selected === "moduleExam" && !isTrainingDepartmentSheet(department)) return "overview";
  return selected;
}

function setDepartmentTab(department, tab) {
  state.departmentTabs = { ...(state.departmentTabs || {}), [department.id]: tab };
  localStorage.setItem("lspd_department_tabs", JSON.stringify(state.departmentTabs));
}

function isTrainingDepartmentSheet(department) {
  const name = cleanText(department?.name || "");
  return /(training|ausbildung)/i.test(name) && !/(human|humane|ressource|resource)/i.test(name);
}

function isHumanResourcesDepartmentSheet(department) {
  const name = cleanText(department?.name || "");
  return /(human|humane|ressource|resource|hr)/i.test(name);
}

function departmentsForOverview() {
  const departments = [...state.departments];
  const pageOrder = state.settings?.pageOrder || [];
  const originalIndex = new Map(departments.map((department, index) => [department.id, index]));
  return departments.sort((a, b) => {
    if (a.id === "direktion") return -1;
    if (b.id === "direktion") return 1;
    const aOrder = pageOrder.indexOf(`dept:${a.id}`);
    const bOrder = pageOrder.indexOf(`dept:${b.id}`);
    if (aOrder !== -1 || bOrder !== -1) return (aOrder === -1 ? 10000 : aOrder) - (bOrder === -1 ? 10000 : bOrder);
    return (originalIndex.get(a.id) || 0) - (originalIndex.get(b.id) || 0);
  });
}

function renderPermissionPickList(type, items, selected = []) {
  const placeholders = {
    role: "z.B. User, Supervisor, Direktion",
    department: "z.B. SWAT, Training, Metro",
    position: "z.B. Leitung, Stv. Leitung, Anwärter",
    rank: "z.B. 0, 5, Sergeant, Director",
    user: "z.B. Name, Dienstnummer, Alexa"
  };
  return `
    <div class="permission-picker" data-perm-picker="${type}">
      <input class="permission-search" placeholder="${escapeHtml(placeholders[type] || "Suchen und hinzufügen")}">
      <small class="permission-hint">${escapeHtml(type === "user" ? "Nach Name oder DN suchen und dann aktivieren." : "Suchen, Vorschlag auswählen und per Schalter aktivieren.")}</small>
      <div class="permission-checks">
        ${items.map((item) => {
          const isSelected = selected.includes(item.value);
          return `<label class="permission-toggle ${isSelected ? "selected" : "suggestion-hidden"}"><input type="checkbox" data-perm-${type}="${escapeHtml(item.value)}" ${isSelected ? "checked" : ""}><span class="permission-switch"></span><span>${escapeHtml(item.label)}</span></label>`;
        }).join("")}
      </div>
    </div>
  `;
}

function permissionSummary(area, key) {
  const rule = permissionRule(area, key);
  if (rule.all) return "Alle";
  const parts = [];
  if (rule.roles?.length) parts.push(`${rule.roles.length} Gruppen`);
  if (rule.departments?.length) parts.push(`${rule.departments.length} Abteilungen`);
  if (rule.positions?.length) parts.push(`${rule.positions.length} Positionen`);
  if (rule.ranks?.length) parts.push(`${rule.ranks.length} Ränge`);
  if (rule.users?.length) parts.push(`${rule.users.length} Personen`);
  return parts.length ? parts.join(" / ") : "Nur Standardrechte";
}

function renderPermissionEditor(area, key, label, description = "") {
  const rule = permissionRule(area, key);
  const rankItems = [...state.ranks].sort((a, b) => Number(a.value) - Number(b.value)).map((rank) => ({ value: String(rank.value), label: rankOptionLabel(rank) }));
  const roleItems = state.roles.map((role) => ({ value: role, label: role }));
  const userItems = state.users.map((user) => ({ value: user.id, label: `${fullName(user)} - DN ${user.dn || "-"}` }));
  const departmentItems = state.departments.filter((department) => department.id !== "direktion").map((department) => ({ value: department.id, label: department.name }));
  const positionItems = state.departments.flatMap((department) => departmentPositionsFor(department).map((position) => ({ value: `${department.id}:${position}`, label: `${department.name} - ${position}` })));
  return `
    <article class="permission-row" data-permission-area="${area}" data-permission-key="${escapeHtml(key)}">
      <div class="permission-row-head">
        <div class="permission-copy">
          <strong>${escapeHtml(label)}</strong>
          ${description ? `<small>${escapeHtml(description)}</small>` : ""}
        </div>
        <span class="permission-summary">${escapeHtml(permissionSummary(area, key))}</span>
      </div>
      <label class="permission-all-control"><input type="checkbox" data-perm-all ${rule.all ? "checked" : ""}><span class="permission-switch"></span><span>Alle erlauben</span></label>
      <div class="permission-controls">
        <div><span>Gruppen</span>${renderPermissionPickList("role", roleItems, rule.roles || [])}</div>
        <div><span>Abteilungen</span>${renderPermissionPickList("department", departmentItems, rule.departments || [])}</div>
        <div><span>Positionen</span>${renderPermissionPickList("position", positionItems, rule.positions || [])}</div>
        <div><span>Ränge</span>${renderPermissionPickList("rank", rankItems, (rule.ranks || []).map(String))}</div>
        <div><span>Personen</span>${renderPermissionPickList("user", userItems, rule.users || [])}</div>
      </div>
    </article>
  `;
}
function collectPermissionEditors() {
  const permissions = {
    pages: { ...(state.settings.permissions?.pages || {}) },
    actions: { ...(state.settings.permissions?.actions || {}) }
  };
  document.querySelectorAll("[data-permission-area][data-permission-key]").forEach((row) => {
    permissions[row.dataset.permissionArea][row.dataset.permissionKey] = {
      all: Boolean(row.querySelector("[data-perm-all]")?.checked),
      roles: Array.from(row.querySelectorAll("[data-perm-role]:checked")).map((input) => input.dataset.permRole),
      ranks: Array.from(row.querySelectorAll("[data-perm-rank]:checked")).map((input) => Number(input.dataset.permRank)),
      users: Array.from(row.querySelectorAll("[data-perm-user]:checked")).map((input) => input.dataset.permUser),
      departments: Array.from(row.querySelectorAll("[data-perm-department]:checked")).map((input) => input.dataset.permDepartment),
      positions: Array.from(row.querySelectorAll("[data-perm-position]:checked")).map((input) => input.dataset.permPosition)
    };
  });
  return permissions;
}

function renderITOverviewPanel(editablePages) {
  const activeDuty = state.duty.length;
  const protectedPages = editablePages.filter((page) => isInternalSheetPage(page) && isPageViewRestricted(page)).length;
  const restartTimes = state.settings.restartTimes || [];
  return `
    <div class="panel it-section-card it-overview-start it-overview-redesign">
      <div class="it-overview-headline">
        <div>
          <h3>IT Übersicht</h3>
          <p class="muted">Zentrale Verwaltung für Sicherungen, Struktur, Sessions und Restarts.</p>
        </div>
        <div class="it-status-strip">
          <span><b>${editablePages.length}</b> Reiter</span>
          <span><b>${state.departments.length}</b> Abteilungen</span>
          <span><b>${activeDuty}</b> im Dienst</span>
          <span><b>${protectedPages}</b> geschützt</span>
        </div>
      </div>
      <div class="it-overview-grid">
        <section class="it-overview-block">
          <div><strong>Daten & Sessions</strong><small>Sichern, importieren und aktive Logins steuern.</small></div>
          <div class="it-action-grid compact-actions">
            <button class="it-tool" id="overviewExportData"><strong>Datensicherung</strong><span>JSON exportieren</span></button>
            <button class="it-tool" id="overviewImportData"><strong>Datenimport</strong><span>Backup wiederherstellen</span></button>
            <button class="it-tool" id="overviewClearSessions"><strong>Sessions</strong><span>Andere Logins abmelden</span></button>
          </div>
        </section>
        <section class="it-overview-block">
          <div><strong>Struktur</strong><small>Neue Blätter anlegen und schnell in die Reiterverwaltung springen.</small></div>
          <div class="it-action-grid compact-actions">
            <button class="it-tool" id="overviewCreatePage"><strong>Reiter erstellen</strong><span>Leeres Template-Blatt</span></button>
            <button class="it-tool" id="overviewCreateDepartment"><strong>Abteilung erstellen</strong><span>Abteilungs-Template</span></button>
            <button class="it-tool ${state.settings?.devMode ? "devmode-on" : ""}" id="overviewToggleDevMode"><strong>Devmode</strong><span>${state.settings?.devMode ? "Aktiv" : "Aus"}</span></button>
          </div>
        </section>
        <section class="it-overview-block it-overview-restarts">
          <div><strong>Restarts</strong><small>${restartTimes.length ? `${restartTimes.length} Restartzeit${restartTimes.length === 1 ? "" : "en"} aktiv` : "Noch keine Restartzeit angelegt"}</small></div>
          <div class="restart-editor">
            <input id="restartTimeInput" type="time" value="00:00">
            <button class="blue-btn" id="addRestartTime" type="button">Hinzufügen</button>
          </div>
          <div class="restart-list">
            ${restartTimes.map((time) => `
              <span class="restart-chip"><b>${escapeHtml(time)}</b><button class="mini-icon delete-restart-time" type="button" data-time="${escapeHtml(time)}" title="Löschen">${actionIcon("delete")}</button></span>
            `).join("") || `<p class="muted">Noch keine Restartzeiten angelegt.</p>`}
          </div>
        </section>
      </div>
    </div>
  `;
}
function renderITDepartmentPositionsPanel() {
  const departments = state.departments.filter((department) => department.id !== "direktion");
  return `
    <div class="panel it-section-card it-department-positions-card">
      <div class="it-section-title">
        <span>04</span>
        <div><h3>Interne Abteilungsränge</h3><p class="muted">Positionen wie Leitung, Stv. Leitung oder eigene interne Ränge pro Abteilung bearbeiten.</p></div>
      </div>
      <div class="it-department-position-grid">
        ${departments.map((department) => `
          <article class="department-position-shortcut">
            <strong>${escapeHtml(department.name)}</strong>
            <small>${departmentPositionsFor(department).map(escapeHtml).join(" / ")}</small>
            <button class="ghost-btn edit-department-positions" type="button" data-page-key="dept:${escapeHtml(department.id)}">Positionen bearbeiten</button>
          </article>
        `).join("") || `<p class="muted">Keine Abteilungen vorhanden.</p>`}
      </div>
    </div>
  `;
}

function discordRoleColor(role) {
  const color = Number(role?.color || 0);
  return color ? `#${color.toString(16).padStart(6, "0")}` : "#99aab5";
}

function discordSelectedRoleIds(selectedRoleIds = []) {
  const ids = Array.isArray(selectedRoleIds) ? selectedRoleIds : String(selectedRoleIds || "").split(",");
  return [...new Set(ids.map((roleId) => String(roleId || "").trim()).filter(Boolean))];
}

function renderDiscordRolePicker(attribute, key, selectedRoleIds) {
  const roles = state.settings?.discordSync?.importedRoles || [];
  const selected = discordSelectedRoleIds(selectedRoleIds);
  if (!roles.length) {
    return `<div class="discord-role-picker disabled" ${attribute}="${escapeHtml(key)}" data-selected=""><span class="muted">Bitte zuerst Server-Rollen importieren.</span></div>`;
  }
  const selectedRoles = selected.map((roleId) => roles.find((role) => String(role.id) === String(roleId))).filter(Boolean);
  return `
    <div class="discord-role-picker" ${attribute}="${escapeHtml(key)}" data-selected="${escapeHtml(selected.join(","))}">
      <div class="discord-role-chip-list">
        ${selectedRoles.map((role) => `
          <span class="discord-role-chip" style="--role-color:${discordRoleColor(role)}" data-role-id="${escapeHtml(role.id)}">
            <b>@${escapeHtml(role.name)}</b>
            <button type="button" class="discord-role-remove" data-role-id="${escapeHtml(role.id)}">×</button>
          </span>
        `).join("") || `<span class="discord-role-empty">Keine Rolle ausgewählt</span>`}
      </div>
      <input class="discord-role-search" type="text" autocomplete="off" placeholder="@rolle suchen">
      <div class="discord-role-menu hidden">
        ${roles.map((role) => `
          <button type="button" class="discord-role-option ${selected.includes(String(role.id)) ? "selected" : ""}" data-role-id="${escapeHtml(role.id)}" data-role-name="${escapeHtml(role.name.toLowerCase())}" style="--role-color:${discordRoleColor(role)}">
            <span>@${escapeHtml(role.name)}</span>
            <small>${role.managed ? "verwaltet" : `Position ${escapeHtml(role.position)}`}</small>
          </button>
        `).join("")}
      </div>
    </div>
  `;
}

function renderDiscordSyncPanel() {
  const sync = state.settings?.discordSync || {};
  const rankRoles = sync.rankRoles || {};
  const departmentRoles = sync.departmentRoles || {};
  const importedRoles = sync.importedRoles || [];
  const sortedRanks = [...state.ranks].sort((a, b) => b.value - a.value);
  const orderedDepartmentKeys = orderPages((state.departments || []).map((department) => `dept:${department.id}`));
  const departments = orderedDepartmentKeys
    .map((key) => (state.departments || []).find((department) => `dept:${department.id}` === key))
    .filter(Boolean);
  return `
    <div class="panel it-section-card it-discord-card">
      <div class="it-section-title">
        <div><h3>Discord Sync</h3><p class="muted">Rang- und Abteilungsrollen vorbereiten, damit Discord-Rollen passend zum Dienstblatt vergeben werden koennen.</p></div>
        <div class="button-row">
          <button class="ghost-btn" id="linkOwnDiscord" type="button">Meinen Discord verknüpfen</button>
          <button class="ghost-btn" id="importDiscordRoles" type="button">Server-Rollen importieren</button>
          <button class="ghost-btn" id="testDiscordSync" type="button">Verbindung testen</button>
          <button class="ghost-btn" id="runDiscordSync" type="button">Jetzt synchronisieren</button>
          <button class="blue-btn" id="saveDiscordSync" type="button">Discord Sync speichern</button>
        </div>
      </div>
      <div class="discord-sync-layout">
        <div class="discord-sync-config">
          <label class="switch-line"><input id="discordSyncEnabled" type="checkbox" ${sync.enabled ? "checked" : ""}><span>Discord Sync aktivieren</span></label>
          <label>Anwendungs-ID<input id="discordApplicationId" inputmode="numeric" autocomplete="off" value="${escapeHtml(sync.applicationId || "")}" placeholder="Discord Anwendungs-ID"></label>
          <label>Öffentlicher Schlüssel<input id="discordPublicKey" autocomplete="off" value="${escapeHtml(sync.publicKey || "")}" placeholder="Discord Public Key"></label>
          <label>OAuth Redirect URL<input id="discordOauthRedirectUrl" autocomplete="off" value="${escapeHtml(sync.oauthRedirectUrl || `${window.location.origin}/`)}" placeholder="Exakt wie im Discord Developer Portal"></label>
          <label>Server ID<input id="discordServerId" inputmode="numeric" autocomplete="off" value="${escapeHtml(sync.serverId || "")}" placeholder="Discord Server ID"></label>
          <label>Bot Token<input id="discordBotToken" type="password" autocomplete="new-password" data-lpignore="true" data-1p-ignore="true" placeholder="${sync.botTokenSet ? "Token ist gespeichert - leer lassen zum Behalten" : "Bot Token eintragen"}"></label>
          <label class="switch-line"><input id="clearDiscordBotToken" type="checkbox"><span>Gespeicherten Bot Token entfernen</span></label>
          <p class="muted">Für Rollen-Sync muss in der Discord Application ein Bot-User existieren und mit bot-Scope auf dem Server sein. Der Bot braucht Rollen verwalten und seine höchste Rolle muss über den Rollen liegen, die hier vergeben werden sollen.</p>
          <p class="discord-import-status">${importedRoles.length ? `${importedRoles.length} Server-Rollen importiert.` : "Noch keine Server-Rollen importiert."}</p>
        </div>
        <div class="discord-sync-section">
          <div><strong>Ränge</strong><small>Jeder Dienstblatt-Rang kann mehrere Discord-Rollen bekommen.</small></div>
          <div class="discord-role-grid">
            ${sortedRanks.map((rank) => `
              <div class="discord-role-row">
                <span>${escapeHtml(rankOptionLabel(rank))}</span>
                ${renderDiscordRolePicker("data-discord-rank-role", rank.value, rankRoles[String(rank.value)] || [])}
              </div>
            `).join("")}
          </div>
        </div>
        <div class="discord-sync-section">
          <div><strong>Abteilungsrollen</strong><small>Pro Abteilung werden nur Leader-Rollen einzeln und eine allgemeine Mitgliederrolle gepflegt.</small></div>
          <div class="discord-department-list">
            ${departments.map((department) => `
              <article class="discord-department-card">
                <div><strong>${escapeHtml(department.name)}</strong><small>${departmentLeaderPositionsFor(department).filter((position) => position !== "Direktion").length} Leader-Rollen + Mitglieder</small></div>
                <div class="discord-role-grid">
                  ${[
                    ...departmentLeaderPositionsFor(department).filter((position) => position !== "Direktion"),
                    "__member"
                  ].map((position) => {
                    const key = `${department.id}:${position}`;
                    const label = position === "__member" ? `${department.name} Mitglieder` : `${department.name} ${position} Leader`;
                    return `
                      <div class="discord-role-row">
                        <span>${escapeHtml(label)}</span>
                        ${renderDiscordRolePicker("data-discord-dept-role", key, departmentRoles[key] || [])}
                      </div>
                    `;
                  }).join("")}
                </div>
              </article>
            `).join("") || `<p class="muted">Keine Abteilungen vorhanden.</p>`}
          </div>
        </div>
      </div>
      <p id="discordSyncMessage" class="muted"></p>
    </div>
  `;
}

function renderIT() {
  if (!hasRole("IT")) {
    content.innerHTML = `<section class="panel"><h3>Kein Zugriff</h3><p class="muted">Dieser Bereich ist nur für IT sichtbar.</p></section>`;
    return;
  }

  const editablePages = editableItPages();
  const sortedRanks = [...state.ranks].sort((a, b) => b.value - a.value);
  const storedItTab = localStorage.getItem("lspd_it_tab") || "overview";
  const itTabs = [["overview", "Übersicht"], ["pages", "Reiter"], ["members", "Mitglieder"], ["ranks", "Ränge"], ["discord", "Discord Sync"]];
  const visibleItTabs = itTabs;
  const itTab = visibleItTabs.some(([id]) => id === storedItTab) ? storedItTab : "overview";
  content.innerHTML = `
    <section class="it-command-center">
      <div class="panel it-hero-panel it-overview-card">
        <div>
          <h3><span class="section-icon">${iconSvg("IT")}</span>IT Verwaltung</h3>
          <p class="muted">Systemsteuerung, Rechte, Mitglieder und Ränge übersichtlich getrennt.</p>
        </div>
        <div class="tabs-row it-tabs">
          ${visibleItTabs.map(([id, label]) => `<button class="${itTab === id ? "tab-active" : ""}" data-it-tab="${id}">${label}</button>`).join("")}
        </div>
      </div>
    </section>

    <section class="it-workbench">
      ${itTab === "overview" ? renderITOverviewPanel(editablePages) : ""}
      <div class="panel it-section-card it-pages-card ${itTab === "pages" ? "" : "hidden"}">
        <div class="it-section-title">
          <span>03</span>
          <div><h3>Reiter & Rechte</h3><p class="muted">Namen ändern und Rechte direkt pro Reiter öffnen.</p></div>
          <div class="button-row">
            <button class="ghost-btn" id="createCustomPage" type="button">Reiter erstellen</button>
            <button class="ghost-btn" id="createDepartmentPage" type="button">Abteilung erstellen</button>
            <button class="blue-btn" id="saveNavLabels" type="button">Speichern</button>
          </div>
        </div>
        <div class="edit-list it-compact-list">
          ${editablePages.map((page, index) => `
            ${isInternalSheetPage(page) && !isInternalSheetPage(editablePages[index - 1] || "") ? `<div class="edit-section-divider"><span>Abteilungsblätter</span><small>Direktion, IT und Abteilungen mit eigenen Rechten für Ansicht, Personal, Notizen und interne Buttons.</small></div>` : ""}
            <label class="edit-row">
              <span class="edit-icon">${iconSvg(page)}</span>
              <span class="edit-name">${restrictedPageEditIcon(page)}${escapeHtml(isDepartmentPage(page) ? navLabel(page) : page)}</span>
              <input data-nav-key="${escapeHtml(page)}" value="${escapeHtml(navLabel(page))}">
              <span class="page-order-controls">
                <button class="mini-icon page-move" type="button" data-page-key="${escapeHtml(page)}" data-direction="-1" title="Nach oben">${iconSvg("ChevronUp")}</button>
                <button class="mini-icon page-move" type="button" data-page-key="${escapeHtml(page)}" data-direction="1" title="Nach unten">${iconSvg("ChevronDown")}</button>
              </span>
              <button class="mini-icon page-permission-open" type="button" data-page-key="${escapeHtml(page)}" title="Rechte verwalten" aria-label="Rechte verwalten">${actionIcon("edit")}</button>
            </label>
          `).join("")}
        </div>
        <p id="navSaveMessage" class="muted"></p>
      </div>
      ${itTab === "pages" ? renderITDepartmentPositionsPanel() : ""}

      <div class="panel it-section-card it-members-card ${itTab === "members" ? "" : "hidden"}">
        <div class="it-section-title">
          <span>04</span>
          <div><h3>Mitglieder</h3><p class="muted">Accounts, Ränge und IT-Zugänge direkt im IT-Blatt bearbeiten.</p></div>
          <button class="blue-btn" id="itCreateMember" type="button">Neues Mitglied einstellen</button>
        </div>
        <div id="defaultCredentialPanel" class="it-password-panel" autocomplete="off">
          <div>
            <strong>Standardpasswort</strong>
            <small>Neue Accounts bekommen dieses Passwort automatisch. Einzelne Accounts kannst du unten darauf zurücksetzen.</small>
          </div>
          <input id="defaultCredentialValue" type="text" autocomplete="off" autocapitalize="off" autocorrect="off" spellcheck="false" data-lpignore="true" data-1p-ignore="true" placeholder="Neues Standardpasswort">
          <button class="blue-btn" id="saveDefaultCredential" type="button">Standardpasswort speichern</button>
          <p id="defaultCredentialMessage" class="muted"></p>
        </div>
        <div class="it-member-list">
          ${state.users.map((user) => `
            <div class="it-member-row">
              <span>${avatarMarkup(user, "sm")}<span><strong>${escapeHtml(fullName(user))}</strong><small>DN ${escapeHtml(user.dn || "-")} · ${escapeHtml(rankLabel(user.rank))}</small></span></span>
              <span class="it-member-roles">${roleBadges(user)}</span>
              <button class="ghost-btn reset-member-password" type="button" data-user-id="${escapeHtml(user.id)}">Passwort Reset</button>
              <button class="mini-icon it-edit-member" type="button" data-user-id="${escapeHtml(user.id)}" title="Mitglied bearbeiten">${actionIcon("edit")}</button>
            </div>
          `).join("")}
        </div>
      </div>

      <div class="panel it-section-card it-ranks-card ${itTab === "ranks" ? "" : "hidden"}">
        <div class="it-section-title">
          <span>05</span>
          <div><h3>Ränge</h3><p class="muted">Rangnamen bearbeiten, hinzufügen oder entfernen.</p></div>
          <div class="button-row">
            <button class="ghost-btn" id="addRank" type="button">Rang hinzufügen</button>
            <button class="red-btn" id="removeRank" type="button">Rang entfernen</button>
            <button class="blue-btn" id="saveRanks" type="button">Speichern</button>
          </div>
        </div>
        <div class="edit-list rank-edit-list it-compact-list">
          ${sortedRanks.map((rank) => `
            <label class="edit-row">
              <span class="rank-number">Rang ${rank.value}</span>
              <input data-rank-value="${rank.value}" value="${escapeHtml(rank.label)}">
              <span class="edit-pencil">${actionIcon("edit")}</span>
            </label>
          `).join("")}
        </div>
        <p id="rankSaveMessage" class="muted"></p>
      </div>
      ${itTab === "discord" ? renderDiscordSyncPanel() : ""}
    </section>

  `;

  document.querySelectorAll("[data-it-tab]").forEach((button) => {
    button.addEventListener("click", () => {
      localStorage.setItem("lspd_it_tab", button.dataset.itTab);
      renderIT();
    });
  });

  $("#saveNavLabels")?.addEventListener("click", async () => {
    const navLabels = {};
    document.querySelectorAll("[data-nav-key]").forEach((input) => {
      navLabels[input.dataset.navKey] = input.value;
    });
    try {
      const data = await api("/api/it/nav-labels", { method: "PATCH", body: JSON.stringify({ navLabels }) });
      state.settings.navLabels = data.navLabels;
      if (Array.isArray(data.departments)) state.departments = data.departments;
      $("#navSaveMessage").textContent = "Reiter gespeichert.";
      renderNavigation();
      renderTopbar();
    } catch (error) {
      $("#navSaveMessage").textContent = error.message;
      $("#navSaveMessage").className = "form-error";
    }
  });

  document.querySelectorAll(".page-move").forEach((button) => button.addEventListener("click", () => movePageOrder(button.dataset.pageKey, Number(button.dataset.direction))));
  $("#createCustomPage")?.addEventListener("click", () => openCreatePageModal("custom"));
  $("#createDepartmentPage")?.addEventListener("click", () => openCreatePageModal("department"));
  $("#overviewCreatePage")?.addEventListener("click", () => openCreatePageModal("custom"));
  $("#overviewCreateDepartment")?.addEventListener("click", () => openCreatePageModal("department"));
  document.querySelectorAll(".edit-department-positions").forEach((button) => button.addEventListener("click", () => openPagePermissionModal(button.dataset.pageKey)));

  $("#saveRanks")?.addEventListener("click", async () => {
    const ranks = Array.from(document.querySelectorAll("[data-rank-value]"))
      .sort((a, b) => Number(a.dataset.rankValue) - Number(b.dataset.rankValue))
      .map((input) => ({ value: Number(input.dataset.rankValue), label: input.value }));
    try {
      const data = await api("/api/it/ranks", { method: "PATCH", body: JSON.stringify({ ranks }) });
      state.ranks = data.ranks;
      $("#rankSaveMessage").textContent = "Ränge gespeichert.";
      renderNavigation();
      renderTopbar();
    } catch (error) {
      $("#rankSaveMessage").textContent = error.message;
      $("#rankSaveMessage").className = "form-error";
    }
  });

  $("#addRank")?.addEventListener("click", openAddRankModal);
  $("#removeRank")?.addEventListener("click", openRemoveRankModal);
  $("#itCreateMember")?.addEventListener("click", () => openUserModal());
  $("#saveDefaultCredential")?.addEventListener("click", saveDefaultPassword);
  $("#saveDiscordSync")?.addEventListener("click", saveDiscordSyncSettings);
  $("#importDiscordRoles")?.addEventListener("click", importDiscordRoles);
  $("#testDiscordSync")?.addEventListener("click", testDiscordSync);
  $("#runDiscordSync")?.addEventListener("click", runDiscordSync);
  $("#linkOwnDiscord")?.addEventListener("click", () => startDiscordOAuth("link"));
  setupDiscordRolePickers();
  document.querySelectorAll(".it-edit-member").forEach((button) => button.addEventListener("click", () => openUserModal(state.users.find((user) => user.id === button.dataset.userId))));
  document.querySelectorAll(".reset-member-password").forEach((button) => button.addEventListener("click", () => openResetPasswordModal(state.users.find((user) => user.id === button.dataset.userId))));
  document.querySelectorAll(".page-permission-open").forEach((button) => button.addEventListener("click", () => openPagePermissionModal(button.dataset.pageKey)));
  setupPermissionSearch(document);
  document.querySelectorAll(".it-section summary button").forEach((button) => {
    button.addEventListener("click", (event) => event.stopPropagation());
  });

  $("#exportDataBtn")?.addEventListener("click", async () => {
    const response = await fetch("/api/it/export", { headers: { Authorization: `Bearer ${state.token}` } });
    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "lspd-dienstblatt-export.json";
    link.click();
    URL.revokeObjectURL(url);
  });
  $("#overviewExportData")?.addEventListener("click", async () => {
    const response = await fetch("/api/it/export", { headers: { Authorization: `Bearer ${state.token}` } });
    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "lspd-dienstblatt-export.json";
    link.click();
    URL.revokeObjectURL(url);
  });

  $("#importDataBtn")?.addEventListener("click", openDataImportModal);
  $("#overviewImportData")?.addEventListener("click", openDataImportModal);

  $("#clearSessionsBtn")?.addEventListener("click", async () => {
    await api("/api/it/clear-sessions", { method: "POST", body: "{}" });
  });
  $("#overviewClearSessions")?.addEventListener("click", async () => {
    await api("/api/it/clear-sessions", { method: "POST", body: "{}" });
    showNotify("Andere Sessions wurden abgemeldet.");
  });

  $("#toggleDevModeBtn")?.addEventListener("click", async () => {
    const data = await api("/api/it/devmode", {
      method: "PATCH",
      body: JSON.stringify({ devMode: !state.settings?.devMode })
    });
    state.settings = data.settings;
    syncDevModeAuthStorage();
    renderApp();
  });
  $("#overviewToggleDevMode")?.addEventListener("click", async () => {
    const data = await api("/api/it/devmode", {
      method: "PATCH",
      body: JSON.stringify({ devMode: !state.settings?.devMode })
    });
    state.settings = data.settings;
    syncDevModeAuthStorage();
    renderApp();
  });
  $("#addRestartTime")?.addEventListener("click", () => saveRestartTimes([...(state.settings.restartTimes || []), $("#restartTimeInput").value]));
  document.querySelectorAll(".delete-restart-time").forEach((button) => {
    button.addEventListener("click", () => saveRestartTimes((state.settings.restartTimes || []).filter((time) => time !== button.dataset.time)));
  });
}

async function saveRestartTimes(times) {
  const data = await api("/api/it/restarts", {
    method: "PATCH",
    body: JSON.stringify({ restartTimes: Array.from(new Set(times.filter(Boolean))).sort() })
  });
  state.settings = data.settings;
  renderIT();
}

function updateDiscordRolePicker(picker) {
  const roles = state.settings?.discordSync?.importedRoles || [];
  const selected = discordSelectedRoleIds(picker.dataset.selected || "");
  const chipList = picker.querySelector(".discord-role-chip-list");
  if (chipList) {
    chipList.innerHTML = selected.map((roleId) => roles.find((role) => String(role.id) === String(roleId))).filter(Boolean).map((role) => `
      <span class="discord-role-chip" style="--role-color:${discordRoleColor(role)}" data-role-id="${escapeHtml(role.id)}">
        <b>@${escapeHtml(role.name)}</b>
        <button type="button" class="discord-role-remove" data-role-id="${escapeHtml(role.id)}">×</button>
      </span>
    `).join("") || `<span class="discord-role-empty">Keine Rolle ausgewählt</span>`;
  }
  picker.querySelectorAll(".discord-role-option").forEach((option) => {
    option.classList.toggle("selected", selected.includes(String(option.dataset.roleId)));
  });
}

function filterDiscordRolePicker(picker) {
  const query = (picker.querySelector(".discord-role-search")?.value || "").replace(/^@/, "").trim().toLowerCase();
  picker.querySelectorAll(".discord-role-option").forEach((option) => {
    const match = !query || (option.dataset.roleName || "").includes(query);
    option.classList.toggle("hidden", !match);
  });
}

function setupDiscordRolePickers() {
  document.querySelectorAll(".discord-role-picker:not(.disabled)").forEach((picker) => {
    const input = picker.querySelector(".discord-role-search");
    const menu = picker.querySelector(".discord-role-menu");
    input?.addEventListener("focus", () => {
      menu?.classList.remove("hidden");
      filterDiscordRolePicker(picker);
    });
    input?.addEventListener("input", () => {
      menu?.classList.remove("hidden");
      filterDiscordRolePicker(picker);
    });
    picker.querySelectorAll(".discord-role-option").forEach((option) => {
      option.addEventListener("click", () => {
        const selected = discordSelectedRoleIds(picker.dataset.selected || "");
        const roleId = String(option.dataset.roleId || "");
        if (roleId && !selected.includes(roleId)) selected.push(roleId);
        picker.dataset.selected = selected.join(",");
        if (input) input.value = "";
        updateDiscordRolePicker(picker);
        filterDiscordRolePicker(picker);
        input?.focus();
      });
    });
    picker.addEventListener("click", (event) => {
      const remove = event.target.closest(".discord-role-remove");
      if (!remove) return;
      const selected = discordSelectedRoleIds(picker.dataset.selected || "").filter((roleId) => roleId !== String(remove.dataset.roleId || ""));
      picker.dataset.selected = selected.join(",");
      updateDiscordRolePicker(picker);
      filterDiscordRolePicker(picker);
    });
  });
  if (!window.discordRolePickerOutsideCloseInstalled) {
    window.discordRolePickerOutsideCloseInstalled = true;
    document.addEventListener("click", (event) => {
      document.querySelectorAll(".discord-role-picker .discord-role-menu").forEach((menu) => {
        if (!menu.closest(".discord-role-picker")?.contains(event.target)) menu.classList.add("hidden");
      });
    });
  }
}

async function saveDefaultPassword() {
  const input = $("#defaultCredentialValue");
  const message = $("#defaultCredentialMessage");
  const defaultPassword = input?.value.trim() || "";
  if (!defaultPassword) {
    if (message) {
      message.textContent = "Bitte ein Standardpasswort eintragen.";
      message.className = "form-error";
    }
    return;
  }
  try {
    const data = await api("/api/it/default-password", { method: "PATCH", body: JSON.stringify({ defaultPassword }) });
    state.settings = data.settings || state.settings;
    if (input) input.value = "";
    if (message) {
      message.textContent = "Standardpasswort gespeichert. Neue Accounts und Resets nutzen es ab jetzt.";
      message.className = "muted";
    }
  } catch (error) {
    if (message) {
      message.textContent = error.message;
      message.className = "form-error";
    }
  }
}

async function saveDiscordSyncSettings(options = {}) {
  const skipNotify = Boolean(options.skipNotify);
  const rethrow = Boolean(options.rethrow);
  const message = $("#discordSyncMessage");
  const rankRoles = {};
  document.querySelectorAll("[data-discord-rank-role]").forEach((picker) => {
    const roleIds = discordSelectedRoleIds(picker.dataset.selected || "");
    if (roleIds.length) rankRoles[picker.dataset.discordRankRole] = roleIds;
  });
  const departmentRoles = {};
  document.querySelectorAll("[data-discord-dept-role]").forEach((picker) => {
    const roleIds = discordSelectedRoleIds(picker.dataset.selected || "");
    if (roleIds.length) departmentRoles[picker.dataset.discordDeptRole] = roleIds;
  });
  const discordSync = {
    enabled: $("#discordSyncEnabled")?.checked || false,
    applicationId: $("#discordApplicationId")?.value.trim() || "",
    publicKey: $("#discordPublicKey")?.value.trim() || "",
    oauthRedirectUrl: $("#discordOauthRedirectUrl")?.value.trim() || "",
    serverId: $("#discordServerId")?.value.trim() || "",
    botToken: $("#discordBotToken")?.value.trim() || "",
    clearBotToken: $("#clearDiscordBotToken")?.checked || false,
    rankRoles,
    departmentRoles
  };
  try {
    const data = await api("/api/it/discord-sync", { method: "PATCH", body: JSON.stringify({ discordSync }) });
    state.settings = data.settings || state.settings;
    renderIT();
    if (!skipNotify) showNotify("Discord Sync gespeichert.");
  } catch (error) {
    if (message) {
      message.textContent = error.message;
      message.className = "form-error";
    }
    if (rethrow) throw error;
  }
}

async function importDiscordRoles() {
  const message = $("#discordSyncMessage");
  try {
    await saveDiscordSyncSettings({ skipNotify: true, rethrow: true });
    const data = await api("/api/it/discord-sync/import-roles", { method: "POST", body: "{}" });
    state.settings = data.settings || state.settings;
    renderIT();
    showNotify(`${data.roles?.length || 0} Discord Rollen importiert.`);
  } catch (error) {
    if (message) {
      message.textContent = error.message;
      message.className = "form-error";
    }
  }
}

async function testDiscordSync() {
  const message = $("#discordSyncMessage");
  try {
    await saveDiscordSyncSettings({ skipNotify: true, rethrow: true });
    const data = await api("/api/it/discord-sync/test", { method: "POST", body: "{}" });
    const guildText = data.guildName ? ` / Server: ${data.guildName}` : "";
    showNotify(`Discord Verbindung OK: ${data.botName || "Bot"}${guildText}`);
  } catch (error) {
    if (message) {
      message.textContent = error.message;
      message.className = "form-error";
    }
  }
}

async function runDiscordSync() {
  const message = $("#discordSyncMessage");
  try {
    const data = await api("/api/it/discord-sync/run", { method: "POST", body: "{}" });
    showNotify(`Discord Sync fuer ${data.synced || 0} Accounts gestartet.`);
  } catch (error) {
    if (message) {
      message.textContent = error.message;
      message.className = "form-error";
    }
  }
}

function openResetPasswordModal(user) {
  if (!user) return;
  openModal(`
    <h3>Passwort zurücksetzen</h3>
    <p class="muted">${escapeHtml(fullName(user))} kann sich danach wieder mit dem aktuellen Standardpasswort anmelden.</p>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions">
      <button class="ghost-btn" type="button" data-close>Abbrechen</button>
      <button class="orange-btn" id="confirmPasswordReset" type="button">Auf Standardpasswort setzen</button>
    </div>
  `, (modal) => {
    modal.querySelector("#confirmPasswordReset").addEventListener("click", async () => {
      try {
        await api(`/api/it/users/${user.id}/reset-password`, { method: "POST", body: "{}" });
        closeModal();
        await bootstrap();
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
  });
}

async function movePageOrder(page, direction) {
  const list = editableItPages();
  const index = list.indexOf(page);
  const nextIndex = index + direction;
  if (index < 0 || nextIndex < 0 || nextIndex >= list.length) return;
  const next = [...list];
  [next[index], next[nextIndex]] = [next[nextIndex], next[index]];
  const data = await api("/api/it/page-order", { method: "PATCH", body: JSON.stringify({ pageOrder: next }) });
  state.settings = data.settings;
  renderNavigation();
  renderIT();
}

function openCreatePageModal(type) {
  const isDepartment = type === "department";
  openModal(`
    <h3>${isDepartment ? "Abteilung erstellen" : "Reiter erstellen"}</h3>
    <p class="muted">${isDepartment ? "Erstellt ein leeres Abteilungsblatt mit Übersicht, Leitung und Notizen." : "Erstellt einen leeren Template-Reiter."}</p>
    <label>Name<input id="newPageName" placeholder="${isDepartment ? "z.B. Detective" : "z.B. Dienstanweisungen"}"></label>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions">
      <button class="ghost-btn" data-close>Abbrechen</button>
      <button class="blue-btn" id="confirmCreatePage">Erstellen</button>
    </div>
  `, (modal) => {
    modal.querySelector("#confirmCreatePage").addEventListener("click", async () => {
      try {
        const name = modal.querySelector("#newPageName").value.trim();
        const data = await api(isDepartment ? "/api/it/departments" : "/api/it/custom-pages", {
          method: "POST",
          body: JSON.stringify({ name })
        });
        state.settings = data.settings || state.settings;
        if (Array.isArray(data.departments)) state.departments = data.departments;
        if (Array.isArray(data.settings?.customPages)) state.customPages = data.settings.customPages;
        else if (!isDepartment && data.page) state.customPages = [...(state.customPages || []), data.page];
        closeModal();
        renderNavigation();
        renderIT();
      } catch (error) {
        modal.querySelector("#modalError").textContent = error.message;
      }
    });
  });
}

function openDataImportModal() {
  openModal(`
    <h3>Daten importieren</h3>
    <p class="muted">Importiert eine vollständige Dienstblatt-Datensicherung und ersetzt alle aktuellen Online-Daten. Danach wirst du abgemeldet und meldest dich mit den importierten Accounts neu an.</p>
    <label>Datensicherung auswählen<input id="dataImportFile" type="file" accept="application/json,.json"></label>
    <label class="checkbox-line">Ich verstehe, dass die aktuellen Online-Daten ersetzt werden.<input type="checkbox" id="confirmDataImport"></label>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions">
      <button class="ghost-btn" data-close>Abbrechen</button>
      <button class="red-btn" id="runDataImport">Importieren</button>
    </div>
  `, (modal) => {
    modal.querySelector("#runDataImport").addEventListener("click", async () => {
      const file = modal.querySelector("#dataImportFile").files?.[0];
      const confirmed = modal.querySelector("#confirmDataImport").checked;
      if (!file) {
        modal.querySelector("#modalError").textContent = "Bitte eine JSON-Datei auswählen.";
        return;
      }
      if (!confirmed) {
        modal.querySelector("#modalError").textContent = "Bitte die Sicherheitsabfrage bestätigen.";
        return;
      }
      try {
        const text = await file.text();
        const db = JSON.parse(text);
        const data = await api("/api/it/import", { method: "POST", body: JSON.stringify({ db }) });
        state.token = "";
        state.currentUser = null;
        authStorage().removeItem("lspdToken");
        closeModal();
        showNotify(`Import abgeschlossen: ${data.users} Benutzer importiert. Bitte neu einloggen.`, "success");
        window.setTimeout(() => window.location.reload(), 900);
      } catch (error) {
        modal.querySelector("#modalError").textContent = error.message || "Import fehlgeschlagen.";
      }
    });
  });
}

async function saveItPermissions() {
  try {
    const data = await api("/api/it/permissions", { method: "PATCH", body: JSON.stringify({ permissions: collectPermissionEditors() }) });
    state.settings.permissions = data.permissions;
    const message = $("#permissionSaveMessage");
    if (message) {
      message.textContent = "Rechte gespeichert.";
      message.className = "muted";
    }
    renderNavigation();
    return data.permissions;
  } catch (error) {
    const message = $("#permissionSaveMessage");
    if (message) {
      message.textContent = error.message;
      message.className = "form-error";
    }
    throw error;
  }
}

function setupPermissionSearch(root = document) {
  root.querySelectorAll(".permission-search").forEach((input) => {
    const picker = input.closest(".permission-picker");
    const checks = picker.querySelector(".permission-checks");
    const syncPicker = (resetScroll = false) => {
      const term = input.value.toLowerCase().trim();
      checks.querySelectorAll("label").forEach((label) => {
        const checked = Boolean(label.querySelector("input")?.checked);
        const matches = label.textContent.toLowerCase().includes(term);
        label.classList.toggle("selected", checked);
        label.classList.toggle("suggestion-hidden", !checked && (!term || !matches));
      });
      if (resetScroll) checks.scrollTop = 0;
    };
    input.addEventListener("input", () => {
      syncPicker();
    });
    checks.querySelectorAll("input").forEach((checkbox) => {
      checkbox.addEventListener("change", () => {
        syncPicker(true);
      });
    });
    syncPicker();
  });
}

function pagePermissionActions(page) {
  if (isDepartmentPage(page)) {
    const department = departmentByPage(page);
    return [
      [`departmentMembers:${department?.id}`, "Personal verwalten", "Person hinzufügen, Personal verwalten und Positionen ändern."],
      [`departmentNotes:${department?.id}`, "Notizen verwalten", "Notizen erstellen, bearbeiten und löschen."],
      [`departmentInfo:${department?.id}`, "Informationen und interne Buttons", "Abteilungsinformationen, Weiterleitungen, Sondergenehmigungen und weitere interne Buttons bearbeiten."],
      [`departmentLeadership:${department?.id}`, "Leitung-Bereich", "Internen Leitungstab sehen und Mitgliedsnotizen anlegen."]
    ];
  }
  const map = {
    Dienstblatt: [["editDefcon", "DEFCON anpassen", "Zahnrad und DEFCON-Stufe."], ["manageNotes", "Dienstblatt-Notizen", "Notizen schreiben/bearbeiten/löschen."], ["stopAllDuty", "Alle austragen", "Alle Dienste beenden."]],
    Informationen: [["manageInformation", "Informationen bearbeiten", "Weiterleitungen, Sondergenehmigungen und Fraktionen."]],
    Direktion: [["manageMembers", "Mitgliederverwaltung", "Accounts und Archiv verwalten."], ["manageDutyHours", "Dienstzeiten verwalten", "Stunden hinzufügen/entfernen."], ["viewLogs", "Logs sehen", "Logs im Direktionsbereich."]]
  };
  return map[page] || [];
}

function renderDepartmentPositionManager(department) {
  if (!department) return "";
  const leaderPositions = departmentLeaderPositionsFor(department);
  const colorOptions = [["green", "Grün"], ["red", "Rot"], ["orange", "Orange"], ["blue", "Blau"]];
  return `
    <section class="department-position-manager">
      <div class="permission-row-head">
        <div class="permission-copy">
          <strong>Interne Ränge / Positionen</strong>
          <small>Abteilungsränge sortieren, umbenennen und als Leader markieren. Leader haben Zugriff auf Leitung, Notizen und Personalverwaltung.</small>
        </div>
        <button class="ghost-btn" type="button" id="addDepartmentPosition">+ Rang hinzufügen</button>
      </div>
      <div class="department-position-list" id="departmentPositionList">
        ${departmentPositionsFor(department).map((position) => `
          <label class="department-position-row">
            <span>${escapeHtml(position)}</span>
            <input data-dept-position-old="${escapeHtml(position)}" value="${escapeHtml(position)}" ${position === "Direktion" ? "readonly" : ""}>
            <select data-dept-position-color class="position-color-select">
              ${colorOptions.map(([value, label]) => `<option value="${value}" ${positionColorFor(department, position) === value ? "selected" : ""}>${label}</option>`).join("")}
            </select>
            <label class="leader-position-toggle"><input type="checkbox" data-dept-position-leader ${leaderPositions.includes(position) || position === "Direktion" ? "checked" : ""} ${position === "Direktion" ? "disabled" : ""}><span>Leader</span></label>
            <span class="position-order-controls">
              <button class="mini-icon move-department-position" type="button" data-direction="-1" title="Nach oben">${iconSvg("ChevronUp")}</button>
              <button class="mini-icon move-department-position" type="button" data-direction="1" title="Nach unten">${iconSvg("ChevronDown")}</button>
            </span>
            <button class="mini-icon danger remove-department-position" type="button" ${position === "Direktion" ? "disabled" : ""}>${actionIcon("delete")}</button>
          </label>
        `).join("")}
      </div>
    </section>
  `;
}
function collectDepartmentPositions(modal) {
  return Array.from(modal.querySelectorAll("[data-dept-position-old]")).map((input) => ({
    old: input.dataset.deptPositionOld,
    label: input.value.trim(),
    leader: Boolean(input.closest(".department-position-row")?.querySelector("[data-dept-position-leader]")?.checked),
    color: input.closest(".department-position-row")?.querySelector("[data-dept-position-color]")?.value || defaultPositionColor(input.value.trim())
  })).filter((item) => item.label);
}

function openPagePermissionModal(page) {
  const actions = pagePermissionActions(page);
  const department = isDepartmentPage(page) ? departmentByPage(page) : null;
  openModal(`
    <h3>Rechte: ${escapeHtml(navLabel(page))}</h3>
    <p class="muted">Hier stellst du Ansehen und wichtige interne Funktionen für dieses Blatt ein. IT und Direktion bleiben berechtigt, nur der IT-Reiter bleibt ausschließlich IT.</p>
    <div class="permission-list modal-permission-list">
      ${department ? renderDepartmentPositionManager(department) : ""}
      ${renderPermissionEditor("pages", page, "Blatt ansehen", pageDescription(page))}
      ${actions.map(([key, label, description]) => renderPermissionEditor("actions", key, label, description)).join("")}
    </div>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions">
      <button class="ghost-btn" data-close>Abbrechen</button>
      <button class="blue-btn" id="savePagePermissions">Rechte speichern</button>
    </div>
  `, (modal) => {
    modal.classList.add("permission-modal");
    setupPermissionSearch(modal);
    modal.querySelector("#addDepartmentPosition")?.addEventListener("click", () => {
      modal.querySelector("#departmentPositionList")?.insertAdjacentHTML("beforeend", `
        <label class="department-position-row">
          <span>Neu</span>
          <input data-dept-position-old="" value="" placeholder="Name des neuen Rangs">
          <select data-dept-position-color class="position-color-select">
            <option value="green">Grün</option>
            <option value="red">Rot</option>
            <option value="orange">Orange</option>
            <option value="blue" selected>Blau</option>
          </select>
          <label class="leader-position-toggle"><input type="checkbox" data-dept-position-leader><span>Leader</span></label>
          <span class="position-order-controls">
            <button class="mini-icon move-department-position" type="button" data-direction="-1" title="Nach oben">${iconSvg("ChevronUp")}</button>
            <button class="mini-icon move-department-position" type="button" data-direction="1" title="Nach unten">${iconSvg("ChevronDown")}</button>
          </span>
          <button class="mini-icon danger remove-department-position" type="button">${actionIcon("delete")}</button>
        </label>
      `);
    });
    modal.addEventListener("click", (event) => {
      const removeButton = event.target.closest(".remove-department-position");
      if (removeButton && !removeButton.disabled) removeButton.closest(".department-position-row")?.remove();
      const moveButton = event.target.closest(".move-department-position");
      if (moveButton) {
        const row = moveButton.closest(".department-position-row");
        const direction = Number(moveButton.dataset.direction || 0);
        if (direction < 0 && row?.previousElementSibling) row.parentElement.insertBefore(row, row.previousElementSibling);
        if (direction > 0 && row?.nextElementSibling) row.parentElement.insertBefore(row.nextElementSibling, row);
      }
    });
    modal.querySelector("#savePagePermissions").addEventListener("click", async () => {
      try {
        if (department) {
          await api(`/api/departments/${department.id}/positions`, {
            method: "PATCH",
            body: JSON.stringify({ positions: collectDepartmentPositions(modal) })
          });
        }
        await saveItPermissions();
        closeModal();
        await bootstrap();
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
  });
}

function openAddRankModal() {
  const next = Math.max(...state.ranks.map((rank) => Number(rank.value)), -1) + 1;
  openModal(`
    <h3>Rang hinzufügen</h3>
    <label>Rangnummer<input id="newRankValue" type="number" min="0" value="${next}"></label>
    <label>Rangname<input id="newRankLabel" placeholder="Name des Rangs"></label>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions">
      <button class="ghost-btn" data-close>Abbrechen</button>
      <button class="blue-btn" id="confirmAddRank">Hinzufügen</button>
    </div>
  `, (modal) => {
    modal.querySelector("#confirmAddRank").addEventListener("click", () => {
      const value = Number($("#newRankValue").value);
      const label = $("#newRankLabel").value.trim();
      if (!Number.isInteger(value) || value < 0 || !label) {
        $("#modalError").textContent = "Bitte Rangnummer und Rangname angeben.";
        return;
      }
      if (state.ranks.some((rank) => Number(rank.value) === value)) {
        $("#modalError").textContent = "Diese Rangnummer existiert bereits.";
        return;
      }
      state.ranks.push({ value, label });
      closeModal();
      renderIT();
    });
  });
}

function openRemoveRankModal() {
  openModal(`
    <h3>Rang entfernen</h3>
    <label>Rang auswählen
      <select id="removeRankValue">
        ${[...state.ranks].sort((a, b) => Number(a.value) - Number(b.value)).map((rank) => `<option value="${rank.value}">Rang ${rank.value} - ${escapeHtml(rank.label)}</option>`).join("")}
      </select>
    </label>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions">
      <button class="ghost-btn" data-close>Abbrechen</button>
      <button class="red-btn" id="confirmRemoveRank">Löschen</button>
    </div>
  `, (modal) => {
    modal.querySelector("#confirmRemoveRank").addEventListener("click", () => {
      const value = Number($("#removeRankValue").value);
      state.ranks = state.ranks.filter((rank) => Number(rank.value) !== value);
      closeModal();
      renderIT();
    });
  });
}

function renderDepartmentsOverview() {
  const departments = departmentsForOverview();
  content.innerHTML = `
    <section class="department-grid">
      ${departments.map((department) => renderDepartmentCard(department)).join("")}
    </section>
  `;

  document.querySelectorAll(".department-info").forEach((button) => {
    button.addEventListener("click", () => openDepartmentInfoModal(state.departments.find((department) => department.id === button.dataset.departmentId)));
  });
  document.querySelectorAll(".department-add").forEach((button) => {
    button.addEventListener("click", () => openDepartmentMemberModal(state.departments.find((department) => department.id === button.dataset.departmentId)));
  });
  document.querySelectorAll(".department-manage").forEach((button) => {
    button.addEventListener("click", () => openDepartmentManageModal(state.departments.find((department) => department.id === button.dataset.departmentId)));
  });
  document.querySelectorAll(".department-expand").forEach((button) => {
    button.addEventListener("click", () => {
      const id = button.dataset.departmentId;
      if (expandedDepartments.has(id)) expandedDepartments.delete(id);
      else expandedDepartments.add(id);
      localStorage.setItem("lspd_expanded_departments", JSON.stringify([...expandedDepartments]));
      renderDepartmentsOverview();
    });
  });
}

function renderDepartmentCard(department) {
  const members = [...department.members].sort((a, b) => positionPowerFor(department, b.position) - positionPowerFor(department, a.position) || b.user.rank - a.user.rank);
  const isExpanded = expandedDepartments.has(department.id);
  const visibleMembers = isExpanded ? members : members.slice(0, 5);
  const hiddenCount = Math.max(0, members.length - 5);
  return `
    <article class="department-card">
      <div class="department-card-head">
        <strong>${iconSvg("Direktion")} ${escapeHtml(department.name)} <span class="department-member-count">${members.length}</span></strong>
        <span class="application-pill ${department.applicationStatus === "Offen" ? "open" : "closed"}">${escapeHtml(department.applicationStatus)}</span>
      </div>
      <button class="department-info info-strip" data-department-id="${escapeHtml(department.id)}">${iconSvg("Informationen")} Informationen</button>
      <table class="mini-table">
        <thead><tr><th>Position</th><th>Rang</th><th>Name</th></tr></thead>
        <tbody>
          ${visibleMembers.length ? visibleMembers.map((member) => `
            <tr>
              <td><span class="position-chip ${positionClass(member.position, department)}">${escapeHtml(member.position)}</span></td>
              <td>${escapeHtml(member.user.rank)}</td>
              <td class="dept-card-name"><span class="online-dot ${member.isOnDuty ? "online" : ""}"></span><span>${escapeHtml(fullName(member.user))}</span></td>
            </tr>
          `).join("") : `<tr><td colspan="3" class="muted">Noch keine Mitglieder.</td></tr>`}
        </tbody>
      </table>
      ${departmentActionAllowed(department, "departmentMembers") ? `<button class="blue-btn department-manage" data-department-id="${escapeHtml(department.id)}">${iconSvg("Mitglieder")} Personal verwalten</button>` : ""}
      ${hiddenCount ? `<button class="blue-btn department-expand" data-department-id="${escapeHtml(department.id)}">${iconSvg("ChevronDown")} ${isExpanded ? "Weniger anzeigen" : `${hiddenCount} weitere anzeigen`}</button>` : ""}
    </article>
  `;
}

function renderDepartmentPage(department) {
  if (!department) {
    renderTemplate("Abteilung");
    return;
  }
  const leaders = department.members.filter((member) => ["Leitung", "Stv. Leitung"].includes(member.position));
  const leaderText = leaders.map((member) => fullName(member.user)).join(", ") || "-";
  const canMembers = departmentActionAllowed(department, "departmentMembers");
  const canNotes = departmentActionAllowed(department, "departmentNotes");
  const canInfo = departmentActionAllowed(department, "departmentInfo");
  const canLeadership = departmentActionAllowed(department, "departmentLeadership");
  const isTrainingDepartment = isTrainingDepartmentSheet(department);
  const isHumanResourcesDepartment = isHumanResourcesDepartmentSheet(department);
  const tab = departmentTab(department);
  content.innerHTML = `
    <section class="internal-subhead department-overview-head">
      <h2>${escapeHtml(department.name)} Abteilung</h2>
      <div class="department-control-row">
        <div class="tabs-row department-tabs">
          <button class="${tab === "overview" ? "tab-active" : ""}" data-department-tab="overview">\u00dcbersicht</button>
          ${canLeadership ? `<button class="${tab === "leadership" ? "tab-active" : ""}" data-department-tab="leadership">Leitung</button>` : ""}
          ${isHumanResourcesDepartment ? `<button class="${tab === "estExam" ? "tab-active" : ""}" data-department-tab="estExam">EST Prüfung</button>` : ""}
          ${isTrainingDepartment ? `<button class="${tab === "moduleExam" ? "tab-active" : ""}" data-department-tab="moduleExam">Ausbildungen</button>` : ""}
        </div>
        ${canInfo ? `<button class="blue-btn vote-btn">${iconSvg("Abteilungen")} Abstimmung</button>` : ""}
      </div>
      ${tab !== "estExam" ? `<div class="grid-3 internal-stats">
        <div class="stat-card internal-stat-card"><span>Mitglieder</span><i>${iconSvg("Mitglieder")}</i><strong>${department.members.length}</strong><small>Aktive Mitarbeiter</small></div>
        <div class="stat-card internal-stat-card"><span>Leitung / Stv. Leitung</span><i>${iconSvg("Direktion")}</i><strong>${escapeHtml(leaderText === "-" ? "-" : leaders.length)}</strong><small>${escapeHtml(leaderText)}</small></div>
      </div>` : ""}
      ${tab === "overview" ? renderDepartmentOverviewPanels(department, canMembers, canNotes) : ""}
      ${tab === "leadership" && canLeadership ? renderDepartmentLeadershipPanel(department) : ""}
      ${tab === "estExam" && isHumanResourcesDepartment ? renderEstExamPanel(department) : ""}
      ${tab === "moduleExam" && isTrainingDepartment ? renderModuleExamPanel(department) : ""}
    </section>
  `;
  document.querySelectorAll("[data-department-tab]").forEach((button) => button.addEventListener("click", () => {
    setDepartmentTab(department, button.dataset.departmentTab);
    renderDepartmentPage(department);
  }));
  document.querySelectorAll(".department-add").forEach((button) => button.addEventListener("click", () => openDepartmentMemberModal(department)));
  document.querySelectorAll(".dept-note-add").forEach((button) => button.addEventListener("click", () => openDepartmentNoteModal(department)));
  document.querySelectorAll(".dept-member-note-add").forEach((button) => button.addEventListener("click", () => openDepartmentMemberNoteModal(department, button.dataset.userId)));
  $("#leadershipSearch")?.addEventListener("input", (event) => {
    localStorage.setItem(`lspd_leadership_search_${department.id}`, event.target.value);
    renderDepartmentPage(department);
    const input = $("#leadershipSearch");
    input?.focus();
    input?.setSelectionRange(input.value.length, input.value.length);
  });
  $("#leadershipRange")?.addEventListener("change", (event) => {
    localStorage.setItem(`lspd_leadership_range_${department.id}`, event.target.value);
    renderDepartmentPage(department);
  });
  $("#startEstExam")?.addEventListener("click", () => {
    const candidateId = userIdFromExamInput("estCandidateInput", "estCandidateList");
    if (!candidateId) {
      showNotify("Bitte zuerst einen Prüfling auswählen.", "error");
      return;
    }
    const exam = createTrainingExam("est", candidateId, "", trainingStore().estModules);
    saveActiveTrainingExam(exam);
    renderDepartmentPage(department);
    openTrainingExamModal(exam.id);
    showNotify("EST Prüfung angelegt.", "success");
  });
  $("#continueEstExam")?.addEventListener("click", () => {
    openTrainingExamModal($("#continueEstExam")?.dataset.examId);
  });
  $("#startModuleExam")?.addEventListener("click", () => {
    const candidateId = userIdFromExamInput("moduleCandidateInput", "moduleCandidateList");
    const selectedModules = Array.from($("#moduleExamSelect")?.selectedOptions || []).map((option) => option.value);
    if (!candidateId || !selectedModules.length) {
      showNotify("Bitte Prüfling und Modul auswählen.", "error");
      return;
    }
    const store = trainingStore();
    const modules = store.moduleModules.filter((module) => selectedModules.includes(module.id));
    const exam = createTrainingExam("module", candidateId, "", modules);
    saveActiveTrainingExam(exam);
    renderDepartmentPage(department);
    openTrainingExamModal(exam.id);
    showNotify("Modul Prüfung angelegt.", "success");
  });
  document.querySelectorAll(".training-question-add").forEach((button) => button.addEventListener("click", () => openTrainingQuestionModal(button.dataset.bank, button.dataset.moduleId)));
  document.querySelectorAll(".training-question-edit").forEach((button) => button.addEventListener("click", () => openTrainingQuestionModal(button.dataset.bank, button.dataset.moduleId, button.dataset.questionId)));
  document.querySelectorAll(".training-question-delete").forEach((button) => button.addEventListener("click", () => openDeleteTrainingQuestionModal(button.dataset.bank, button.dataset.moduleId, button.dataset.questionId, department)));
  document.querySelectorAll(".training-module-add").forEach((button) => button.addEventListener("click", () => openTrainingModuleModal()));
  document.querySelectorAll(".training-module-edit").forEach((button) => button.addEventListener("click", () => openTrainingModuleModal(button.dataset.moduleId)));
  document.querySelectorAll(".training-module-delete").forEach((button) => button.addEventListener("click", () => openDeleteTrainingModuleModal(button.dataset.moduleId, department)));
  document.querySelectorAll(".training-exam-open").forEach((button) => button.addEventListener("click", () => openTrainingExamModal(button.dataset.examId, button.dataset.readonly === "true")));
  document.querySelectorAll(".training-exam-archive").forEach((button) => button.addEventListener("click", () => archiveTrainingExam(button.dataset.examId, department)));
  document.querySelectorAll(".training-exam-delete").forEach((button) => button.addEventListener("click", () => openDeleteTrainingExamModal(button.dataset.examId, department)));
  document.querySelectorAll(".training-exam-pause").forEach((button) => button.addEventListener("click", () => pauseTrainingExam(button.dataset.examId, department)));
  document.querySelectorAll(".training-exam-stop").forEach((button) => button.addEventListener("click", () => archiveTrainingExam(button.dataset.examId, department)));
  document.querySelectorAll(".training-est-grant").forEach((button) => button.addEventListener("click", () => grantEstTrainingFromArchive(button.dataset.userId, department)));
  setupExamUserPickers(document);
  document.querySelectorAll(".training-archive-search").forEach((input) => input.addEventListener("input", () => {
    const term = input.value.toLowerCase();
    input.closest(".training-archive-card")?.querySelectorAll(".training-archive-row").forEach((row) => row.classList.toggle("hidden", !row.textContent.toLowerCase().includes(term)));
  }));
  if (trainingTimerInterval) window.clearInterval(trainingTimerInterval);
  trainingTimerInterval = window.setInterval(() => {
    document.querySelectorAll(".exam-live-timer[data-started-at]").forEach((item) => {
      const startedAt = item.dataset.startedAt;
      item.textContent = item.dataset.paused === "true" ? "Pausiert" : startedAt ? formatDuration(Date.now() - new Date(startedAt).getTime()) : "Noch nicht gestartet";
    });
  }, 1000);
}

function renderDepartmentOverviewPanels(department, canMembers, canNotes) {
  return `
    <div class="department-layout department-overview-content">
      <div class="panel">
        <div class="panel-header">
          <h3><span class="section-icon">${iconSvg("Mitglieder")}</span>Abteilungsmitglieder</h3>
          ${canMembers ? `<button class="blue-btn department-add" data-department-id="${escapeHtml(department.id)}">${iconSvg("Mitglieder")} Person hinzuf\u00fcgen</button>` : ""}
        </div>
        ${renderDepartmentMemberTable(department)}
      </div>
      <div class="panel">
        <div class="panel-header">
          <h3><span class="section-icon">${iconSvg("Einsatzzentrale")}</span>Notizen</h3>
          ${canNotes ? `<button class="blue-btn dept-note-add" data-department-id="${escapeHtml(department.id)}">+ Neue Notiz</button>` : ""}
        </div>
        <div class="note-list">
          ${department.notes.length ? department.notes.map((note) => renderDepartmentNote(department, note)).join("") : `<p class="muted">Noch keine Notizen vorhanden.</p>`}
        </div>
      </div>
    </div>
  `;
}

function activeEstExamOld() {
  try {
    return JSON.parse(localStorage.getItem("lspd_active_est_exam") || "null");
  } catch {
    return null;
  }
}

function legacyEstModules() {
  return [
    { name: "Rechtskunde", description: "Rechtsfragen und Grundlagen", questions: ["Wann darf eine Person durchsucht werden?", "Welche Rechte gelten bei einer Festnahme?"] },
    { name: "Dienstvorschriften", description: "Interne Regeln und Vorgehen", questions: ["Wie wird ein Einsatzbericht dokumentiert?", "Wann wird eine Leitung informiert?"] },
    { name: "Ortskunde", description: "Orte, Wege und Zuständigkeiten", questions: ["Wo befindet sich der Sammelpunkt?", "Welche Route führt zum Vespucci PD?"] }
  ];
}

const TRAINING_STORE_KEY = "lspd_training_exam_store";

const EST_LOCATION_PROMPTS = [
  "Würfelpark",
  "LSPD HQ",
  "Vespucci Kleidungsladen",
  "EKZ",
  "Ententeich",
  "Alamosee",
  "Schweinefarm",
  "Pferderanch",
  "Casino",
  "Container Hafen",
  "Missionrow PD",
  "Tequilala Bar"
];

function makeTrainingId(prefix) {
  return `${prefix}_${Date.now()}_${Math.random().toString(16).slice(2)}`;
}

async function grantEstTrainingFromArchive(userId, department) {
  try {
    await api(`/api/training/est/${userId}`, { method: "POST" });
    await bootstrap();
    renderDepartmentPage(department);
    showNotify("EST Haken vergeben.", "success");
  } catch (error) {
    showNotify(error.message, "error");
  }
}

function defaultTrainingQuestion(prompt, type = "manual", maxPoints = null) {
  const resolvedMaxPoints = maxPoints ?? (type === "location" ? 1 : type === "scenario" ? 10 : 3);
  return {
    id: makeTrainingId("question"),
    prompt,
    type,
    solution: type === "manual" ? "Musterlösung für den Prüfer eintragen." : "",
    answers: type === "choice" ? ["Antwortmöglichkeit"] : [],
    correctAnswers: [],
    wrongAnswers: [],
    image: "",
    scenarioInfo: "",
    fileAction: "",
    stationType: "",
    targetSeconds: 0,
    timeSeconds: 0,
    maxPoints: resolvedMaxPoints
  };
}

function defaultTrainingStore() {
  return {
    estModules: [
      { id: "est-law", name: "Rechtskunde", description: "Rechtsfragen und Grundlagen", phase: 1, questions: [defaultTrainingQuestion("Wann darf eine Person durchsucht werden?", "manual", 3), defaultTrainingQuestion("Welche Rechte gelten bei einer Festnahme?", "manual", 3)] },
      { id: "est-location", name: "Ortskunde", description: "Orte, Wege und Zuständigkeiten", questions: EST_LOCATION_PROMPTS.map((place) => defaultTrainingQuestion(place, "location")) },
      { id: "est-scenario", name: "Szenario", description: "10-80 / praktisches Szenario mit Akten-/Prüferinfos", phase: 2, questions: [defaultTrainingQuestion("10-80 Szenario", "scenario", 10)] },
      { id: "est-rules", name: "Dienstvorschriften", description: "Interne Regeln und Vorgehen", phase: 3, questions: [defaultTrainingQuestion("Wie wird ein Einsatzbericht dokumentiert?", "manual", 3)] },
      { id: "est-drive", name: "Fahrstrecke", description: "Fahrroute mit Bild und automatischer Zeitwertung", questions: [defaultTrainingQuestion("Fahrstrecke 1", "location", 10)] },
      { id: "est-heli", name: "Helistrecke", description: "Helikopterroute und Landedächer mit Bild und Zeitwertung", phase: 4, questions: [defaultTrainingQuestion("Helistrecke Route", "location", 10), defaultTrainingQuestion("Dachlandung 1", "location", 10)] }
    ],
    moduleModules: trainings.filter((training) => training !== "EST").slice(0, 6).map((training) => ({
      id: `module-${training.toLowerCase().replace(/[^a-z0-9]+/g, "-")}`,
      name: training,
      description: "Modulprüfung vorbereiten",
      questions: [defaultTrainingQuestion(`${training}: Prüffrage eintragen.`, "manual", 3)]
    })),
    activeExams: []
  };
}

function trainingStore() {
  try {
    const stored = JSON.parse(localStorage.getItem(TRAINING_STORE_KEY) || "null");
    if (stored?.estModules?.length) return normalizeTrainingStore({ ...defaultTrainingStore(), ...stored, activeExams: stored.activeExams || [] });
  } catch {
    return defaultTrainingStore();
  }
  const defaults = defaultTrainingStore();
  localStorage.setItem(TRAINING_STORE_KEY, JSON.stringify(defaults));
  return defaults;
}

function normalizeTrainingStore(store) {
  const defaults = defaultTrainingStore();
  const normalizedQuestionMaxPoints = (module, question) => {
    if (module.id === "est-location") return 1;
    if (["est-drive", "est-heli"].includes(module.id)) return Math.min(10, Math.max(1, Number(question.maxPoints || 10)));
    if (question.type === "scenario" || module.id === "est-scenario") return Math.min(10, Math.max(5, Number(question.maxPoints || 10)));
    return Math.min(10, Math.max(3, Number(question.maxPoints || 3)));
  };
  const mergeModules = (current, fallback) => {
    const modules = [...(current || [])];
    fallback.forEach((module) => {
      if (!modules.some((item) => item.id === module.id || item.name === module.name)) modules.push(module);
    });
    return modules.map((module) => ({
      ...module,
      questions: (module.questions || []).map((question) => ({
        ...question,
        answers: Array.isArray(question.answers) ? question.answers : [...(question.correctAnswers || []), ...(question.wrongAnswers || [])].filter(Boolean),
        correctAnswers: [],
        wrongAnswers: [],
        image: question.image || "",
        scenarioInfo: question.scenarioInfo || "",
        fileAction: question.fileAction || "",
        stationType: question.stationType || (module.id === "est-heli" && /dach|landung|combat/i.test(question.prompt || "") ? "combat" : module.id === "est-heli" ? "route" : ""),
        targetSeconds: Number(question.targetSeconds || 0),
        timeSeconds: Number(question.timeSeconds || 0),
        maxPoints: normalizedQuestionMaxPoints(module, question),
        penaltyPoints: 0,
        questionPenalty: false
      }))
    }));
  };
  return {
    ...defaults,
    ...store,
    estModules: mergeModules(store.estModules, defaults.estModules),
    moduleModules: mergeModules(store.moduleModules, defaults.moduleModules),
    activeExams: store.activeExams || []
  };
}

function saveTrainingStore(store) {
  localStorage.setItem(TRAINING_STORE_KEY, JSON.stringify(store));
}

function activeEstExam() {
  return trainingStore().activeExams.find((exam) => exam.kind === "est" && !["Vorbereitung", "Abgeschlossen", "Archiviert"].includes(exam.status));
}

function activeExamItems(kind) {
  return trainingStore().activeExams
    .filter((exam) => exam.kind === kind && !["Vorbereitung", "Abgeschlossen", "Archiviert"].includes(exam.status))
    .sort((a, b) => new Date(b.createdAt || b.startedAt || 0) - new Date(a.createdAt || a.startedAt || 0));
}

function examElapsedText(exam) {
  if (exam.status === "Pausiert") return "Pausiert";
  if (!exam.startedAt) return "Noch nicht gestartet";
  const pausedMs = Number(exam.pausedTotalMs || 0);
  return formatDuration(Date.now() - new Date(exam.startedAt).getTime() - pausedMs);
}

function saveActiveTrainingExam(exam) {
  const store = trainingStore();
  store.activeExams = [exam, ...store.activeExams.filter((item) => item.id !== exam.id)];
  saveTrainingStore(store);
}

function createTrainingExam(kind, candidateId, secondExaminerId, modules) {
  return {
    id: makeTrainingId("exam"),
    kind,
    candidateId,
    examinerId: state.currentUser?.id,
    secondExaminerId,
    status: "Vorbereitung",
    moduleIndex: 0,
    questionIndex: 0,
    reviewMode: false,
    finalResult: null,
    createdAt: new Date().toISOString(),
    startedAt: null,
    modules: modules.map((module) => ({
      id: module.id,
      name: module.name,
      description: module.description,
      questions: module.questions.map((question) => ({ ...question, result: null, traineeAnswer: "", selectedCorrect: [], selectedWrong: [], manualPoints: 0 }))
    }))
  };
}

function examCurrentModule(exam) {
  return exam.modules[exam.moduleIndex] || null;
}

function examCurrentQuestion(exam) {
  return examCurrentModule(exam)?.questions?.[exam.questionIndex] || null;
}

function examUserOptionLabel(user) {
  return `${fullName(user)} - DN ${user.dn || "-"} - ${rankLabel(user.rank)}`;
}

function renderExamUserPicker(id, listId, users, placeholder) {
  return `
    <div class="exam-user-picker" data-exam-picker="${id}">
      <input id="${id}" value="" placeholder="${escapeHtml(placeholder)}" autocomplete="off">
      <input id="${id}Value" type="hidden" value="">
      <div class="exam-user-options" id="${listId}">
        ${users.map((user) => `<button type="button" data-user-id="${escapeHtml(user.id)}" data-label="${escapeHtml(examUserOptionLabel(user))}">${escapeHtml(examUserOptionLabel(user))}</button>`).join("") || `<span class="muted">Keine passenden Mitglieder.</span>`}
      </div>
    </div>
  `;
}

function userIdFromExamInput(inputId, listId) {
  const selectedId = $(`#${inputId}Value`)?.value || "";
  if (selectedId) return selectedId;
  const value = $(`#${inputId}`)?.value.trim() || "";
  if (!value) return "";
  const option = Array.from(document.querySelectorAll(`#${listId} button`)).find((item) => item.dataset.label === value);
  return option?.dataset.userId || "";
}

function setupExamUserPickers(root = document) {
  root.querySelectorAll(".exam-user-picker").forEach((picker) => {
    const input = picker.querySelector("input:not([type='hidden'])");
    const hidden = picker.querySelector("input[type='hidden']");
    const options = picker.querySelector(".exam-user-options");
    const sync = () => {
      const term = input.value.toLowerCase().trim();
      hidden.value = "";
      options.querySelectorAll("button").forEach((button) => {
        button.classList.toggle("hidden", term && !button.dataset.label.toLowerCase().includes(term));
      });
    };
    input.addEventListener("focus", () => picker.classList.add("open"));
    input.addEventListener("click", () => picker.classList.add("open"));
    input.addEventListener("input", sync);
    input.addEventListener("keydown", (event) => {
      if (event.key === "Escape") picker.classList.remove("open");
    });
    options.querySelectorAll("button").forEach((button) => button.addEventListener("click", () => {
      input.value = button.dataset.label;
      hidden.value = button.dataset.userId;
      picker.classList.remove("open");
    }));
  });
}

function examArchiveItems(kind) {
  return trainingStore().activeExams
    .filter((exam) => exam.kind === kind && ["Archiviert", "Abgeschlossen"].includes(exam.status))
    .sort((a, b) => new Date(b.archivedAt || b.startedAt || 0) - new Date(a.archivedAt || a.startedAt || 0));
}

function renderTrainingExamArchive(kind, department) {
  const rows = examArchiveItems(kind);
  const canManageArchive = departmentActionAllowed(department, "departmentLeadership");
  return `
    <section class="panel training-archive-card">
      <div class="panel-header"><div><h3>${kind === "est" ? "EST Prüfungsarchiv" : "Modul Prüfungsarchiv"}</h3><p class="muted">${rows.length} archivierte Prüfungen</p></div><input class="compact-input training-archive-search" placeholder="Archiv durchsuchen"></div>
      <div class="training-archive-list">
        ${rows.length ? rows.map((exam) => renderTrainingExamArchiveRow(exam, canManageArchive)).join("") : `<p class="muted">Noch keine archivierten Prüfungen.</p>`}
      </div>
    </section>
  `;
}

function renderActiveTrainingExams(kind, department) {
  const rows = activeExamItems(kind);
  const canManage = departmentActionAllowed(department, "departmentLeadership");
  return `
    <section class="panel training-active-card">
      <div class="panel-header"><div><h3>Aktive Prüfungen</h3><p class="muted">${rows.length} laufende oder vorbereitete Prüfungen</p></div></div>
      <div class="training-archive-list">
        ${rows.length ? rows.map((exam) => renderActiveTrainingExamRow(exam, canManage)).join("") : `<p class="muted">Keine aktive Prüfung vorhanden.</p>`}
      </div>
    </section>
  `;
}

function renderActiveTrainingExamRow(exam, canManage) {
  const candidate = state.users.find((user) => user.id === exam.candidateId);
  const examiner = state.users.find((user) => user.id === exam.examinerId);
  return `
    <article class="training-archive-row">
      <div>
        <strong>${escapeHtml(candidate ? fullName(candidate) : "Unbekannter Prüfling")}</strong>
        <small>${exam.kind === "est" ? "EST Prüfung" : "Modul Prüfung"} · ${escapeHtml(exam.status)}</small>
      </div>
      <span><b>Prüfer</b>${escapeHtml(examiner ? fullName(examiner) : "-")}</span>
      <span><b>Dauer</b><span class="exam-live-timer" data-started-at="${escapeHtml(exam.startedAt || "")}">${escapeHtml(examElapsedText(exam))}</span></span>
      <div class="button-row">
        <button class="blue-btn training-exam-open" data-exam-id="${escapeHtml(exam.id)}" type="button">Öffnen</button>
        <button class="ghost-btn training-exam-pause" data-exam-id="${escapeHtml(exam.id)}" type="button">Pausieren</button>
        <button class="ghost-btn training-exam-stop" data-exam-id="${escapeHtml(exam.id)}" type="button">Stoppen</button>
        ${canManage ? `<button class="mini-icon danger training-exam-delete" data-exam-id="${escapeHtml(exam.id)}" type="button" title="Löschen">${actionIcon("delete")}</button>` : ""}
      </div>
    </article>
  `;
}

function renderTrainingExamArchiveRow(exam, canManageArchive) {
  const candidate = state.users.find((user) => user.id === exam.candidateId);
  const examiner = state.users.find((user) => user.id === exam.examinerId);
  const secondExaminer = state.users.find((user) => user.id === exam.secondExaminerId);
  const percent = Number(exam.finalResult?.percent || 0);
  const passed = Boolean(exam.finalResult) && percent >= 70;
  const alreadyHasEst = Boolean(candidate?.trainings?.EST);
  const result = exam.finalResult ? `${passed ? "Bestanden" : "Nicht bestanden"} · ${percent}% · ${exam.finalResult.points}/${exam.finalResult.total} Punkte` : "Ohne finale Auswertung";
  return `
    <article class="training-archive-row ${exam.kind === "est" && exam.finalResult ? passed ? "exam-passed" : "exam-failed" : ""}">
      <div>
        <strong>${escapeHtml(candidate ? fullName(candidate) : "Unbekannter Prüfling")}</strong>
        <small>${exam.kind === "est" ? "EST Prüfung" : "Modul Prüfung"} · ${escapeHtml(exam.status)} · ${formatDateTime(exam.archivedAt || exam.startedAt)}</small>
      </div>
      <span><b>Prüfer</b>${escapeHtml(examiner ? fullName(examiner) : "-")}${secondExaminer ? ` / ${escapeHtml(fullName(secondExaminer))}` : ""}</span>
      <span><b>Ergebnis</b>${escapeHtml(result)}</span>
      <div class="button-row">
        <button class="blue-btn training-exam-open" data-exam-id="${escapeHtml(exam.id)}" data-readonly="true" type="button">Verlauf öffnen</button>
        ${exam.kind === "est" && passed && candidate && !alreadyHasEst && canManageArchive ? `<button class="green-btn training-est-grant" data-user-id="${escapeHtml(candidate.id)}" type="button">EST Haken vergeben</button>` : ""}
        ${exam.kind === "est" && alreadyHasEst ? `<span class="requirement-chip ok">EST vergeben</span>` : ""}
        ${canManageArchive ? `<button class="mini-icon danger training-exam-delete" data-exam-id="${escapeHtml(exam.id)}" type="button" title="Archiv löschen">${actionIcon("delete")}</button>` : ""}
      </div>
    </article>
  `;
}

function legacyRenderEstExamPanel(department) {
  const candidates = state.users.filter((user) => !user.trainings?.EST);
  const activeExam = activeEstExam();
  const candidate = activeExam ? state.users.find((user) => user.id === activeExam.candidateId) : null;
  const secondExaminer = activeExam ? state.users.find((user) => user.id === activeExam.secondExaminerId) : null;
  return `
    <div class="training-exam-layout department-overview-content">
      ${renderActiveTrainingExams("est", department)}
      <section class="panel training-exam-card compact-est-panel">
        <div class="panel-header">
          <div><h3>EST Prüfung</h3><p class="muted">Vorlage für EST-Prüfungen mit Prüfer, Modulen, Pause und Auswertung ab 75%.</p></div>
          ${activeExam ? `<span class="requirement-chip ${activeExam.status === "Pausiert" ? "special" : "ok"}">${escapeHtml(activeExam.status)}</span>` : ""}
        </div>
        <div class="exam-start-grid compact-exam-start est-create-row">
          <label>Prüfling ohne EST
            <select id="estCandidateSelect"><option value="">Prüfling auswählen</option>${candidates.map((user) => `<option value="${user.id}">${escapeHtml(fullName(user))} - DN ${escapeHtml(user.dn || "-")}</option>`).join("")}</select>
          </label>
          <label>2. Prüfer optional
            <select id="estSecondExaminer"><option value=""></option>${state.users.map((user) => `<option value="${user.id}">${escapeHtml(fullName(user))}</option>`).join("")}</select>
          </label>
          <button class="blue-btn" id="startEstExam" type="button" ${candidates.length ? "" : "disabled"}>EST Prüfung starten</button>
        </div>
        ${activeExam ? `
          <div class="active-exam-box">
            <div>
              <strong>${escapeHtml(candidate ? fullName(candidate) : "Unbekannter Prüfling")}</strong>
              <small>Prüfer: ${escapeHtml(fullName(state.currentUser))}${secondExaminer ? ` · 2. Prüfer: ${escapeHtml(fullName(secondExaminer))}` : ""}</small>
            </div>
            <div class="button-row">
              ${activeExam.status === "Pausiert" ? `<button class="blue-btn" id="continueEstExam" type="button">Fortsetzen</button>` : `<button class="ghost-btn" id="pauseEstExam" type="button">Pausieren</button>`}
            </div>
          </div>
          <div class="exam-module-grid">
            ${estModules().map((module, moduleIndex) => `
              <article class="exam-module-card">
                <span>Modul ${moduleIndex + 1}</span>
                <strong>${escapeHtml(module.name)}</strong>
                <small>${escapeHtml(module.description)}</small>
                <div class="exam-question-list">
                  ${module.questions.map((question, index) => `
                    <label class="exam-question-row">
                      <input type="checkbox">
                      <span><b>Frage ${index + 1}</b>${escapeHtml(question)}<small>Multiple Choice oder Textbewertung wird später aus dem Fragenpool geladen.</small></span>
                    </label>
                  `).join("")}
                </div>
              </article>
            `).join("")}
          </div>
          <div class="exam-result-preview">
            <span><b>Auswertung</b>Ab 75% bestanden</span>
            <span class="result-pass">Bestanden</span>
            <span class="result-fail">Nicht bestanden</span>
          </div>
        ` : `<p class="muted">Wähle einen Prüfling aus, um die EST-Prüfung als laufende Vorlage zu starten.</p>`}
      </section>
    </div>
  `;
}

function legacyRenderModuleExamPanel(department) {
  const moduleOptions = trainings.filter((training) => training !== "EST");
  return `
    <div class="training-exam-layout department-overview-content">
      <section class="panel training-exam-card">
        <div class="panel-header"><div><h3>Ausbildungen</h3><p class="muted">Vorlage für spätere Modulprüfungen aus offenen Ausbildungen.</p></div></div>
        <div class="exam-start-grid">
          <label>Mitglied
            <select>${state.users.map((user) => `<option>${escapeHtml(fullName(user))} - DN ${escapeHtml(user.dn || "-")}</option>`).join("")}</select>
          </label>
          <label>Module auswählen
            <select multiple>${moduleOptions.map((training) => `<option>${escapeHtml(training)}</option>`).join("")}</select>
          </label>
          <button class="blue-btn" type="button">Modulprüfung vorbereiten</button>
        </div>
        <div class="exam-module-grid">
          ${moduleOptions.slice(0, 6).map((training) => `
            <article class="exam-module-card">
              <span>Modul</span>
              <strong>${escapeHtml(training)}</strong>
              <small>Fragen, Punkte und Antworten werden später über die Modul Verwaltung gepflegt.</small>
            </article>
          `).join("")}
        </div>
      </section>
    </div>
  `;
}

function legacyRenderTrainingManagementPanels() {
  return `
    <section class="training-management-grid">
      <article class="panel training-manage-card">
        <div class="panel-header"><div><h3>EST Verwaltung</h3><p class="muted">Fragenpool für Rechtskunde, Dienstvorschriften und Ortskunde.</p></div><button class="blue-btn" type="button">Frage hinzufügen</button></div>
        <div class="exam-module-grid">
          ${estModules().map((module) => `
            <div class="exam-module-card">
              <strong>${escapeHtml(module.name)}</strong>
              <small>Multiple Choice oder Textfrage · Punktevergabe · richtige und falsche Antworten</small>
              <label>Beispielfrage<input placeholder="Frage eingeben"></label>
              <label>Antworten<textarea placeholder="Richtige und falsche Antworten hinterlegen"></textarea></label>
            </div>
          `).join("")}
        </div>
      </article>
      <article class="panel training-manage-card">
        <div class="panel-header"><div><h3>Modul Verwaltung</h3><p class="muted">Eigene Module für weitere Ausbildungen anlegen und vorbereiten.</p></div><button class="blue-btn" type="button">Modul erstellen</button></div>
        <div class="module-template-list">
          ${trainings.filter((training) => training !== "EST").slice(0, 9).map((training) => `<span>${escapeHtml(training)}<small>Fragenpool vorbereiten</small></span>`).join("")}
        </div>
      </article>
    </section>
  `;
}

function legacyActiveRenderEstExamPanel(department) {
  const store = trainingStore();
  const candidates = state.users.filter((user) => !user.trainings?.EST);
  const activeExam = activeEstExam();
  const candidate = activeExam ? state.users.find((user) => user.id === activeExam.candidateId) : null;
  const secondExaminer = activeExam ? state.users.find((user) => user.id === activeExam.secondExaminerId) : null;
  return `
    <div class="training-exam-layout department-overview-content">
      <section class="panel training-exam-card">
        <div class="panel-header">
          <div><h3>EST Prüfung</h3><p class="muted">Prüfung erstellen und danach in einem eigenen Fenster durchführen.</p></div>
          ${activeExam ? `<span class="requirement-chip ${activeExam.status === "Pausiert" ? "special" : "ok"}">${escapeHtml(activeExam.status)}</span>` : ""}
        </div>
        <div class="exam-start-grid">
          <label>Prüfling ohne EST
            <select id="estCandidateSelect">${candidates.map((user) => `<option value="${user.id}">${escapeHtml(fullName(user))} - DN ${escapeHtml(user.dn || "-")}</option>`).join("")}</select>
          </label>
          <button class="blue-btn" id="startEstExam" type="button" ${candidates.length ? "" : "disabled"}>EST Prüfung erstellen</button>
        </div>
        ${activeExam ? `
          <div class="active-exam-box">
            <div>
              <strong>${escapeHtml(candidate ? fullName(candidate) : "Unbekannter Prüfling")}</strong>
              <small>Aktive EST-Prüfung · Prüfer: ${escapeHtml(fullName(state.currentUser))}${secondExaminer ? ` · 2. Prüfer: ${escapeHtml(fullName(secondExaminer))}` : ""}</small>
            </div>
            <button class="blue-btn" id="continueEstExam" data-exam-id="${escapeHtml(activeExam.id)}" type="button">Prüfungsfenster öffnen</button>
          </div>
        ` : `<p class="muted">Wähle einen Prüfling aus. Die Durchführung öffnet sich danach als eigenes Fenster.</p>`}
        <div class="est-module-strip">
          ${store.estModules.map((module, moduleIndex) => `
            <article class="est-module-chip">
              <span>Modul ${moduleIndex + 1}</span>
              <strong>${escapeHtml(module.name)}</strong>
            </article>
          `).join("")}
        </div>
      </section>
    </div>
  `;
}

function legacyActiveRenderModuleExamPanel(department) {
  const store = trainingStore();
  return `
    <div class="training-exam-layout department-overview-content">
      <section class="panel training-exam-card">
        <div class="panel-header"><div><h3>Ausbildungen</h3><p class="muted">Modulprüfung erstellen und in einem eigenen Fenster durchführen.</p></div></div>
        <div class="exam-start-grid">
          <label>Mitglied
            <select id="moduleCandidateSelect"><option value="">Mitglied auswählen</option>${state.users.map((user) => `<option value="${user.id}">${escapeHtml(fullName(user))} - DN ${escapeHtml(user.dn || "-")}</option>`).join("")}</select>
          </label>
          <label>2. Prüfer optional
            <select id="moduleSecondExaminer"><option value=""></option>${state.users.map((user) => `<option value="${user.id}">${escapeHtml(fullName(user))}</option>`).join("")}</select>
          </label>
          <label>Module auswählen
            <select id="moduleExamSelect" multiple>${store.moduleModules.map((module) => `<option value="${module.id}">${escapeHtml(module.name)}</option>`).join("")}</select>
          </label>
          <button class="blue-btn" id="startModuleExam" type="button">Modulprüfung erstellen</button>
        </div>
        <div class="exam-module-grid">
          ${store.moduleModules.map((module) => `
              <article class="exam-module-card">
                <span>Modul</span>
                <strong>${escapeHtml(module.name)}</strong>
                <small>${escapeHtml(module.description || "Prüfungsmodul")}</small>
              </article>
          `).join("")}
        </div>
      </section>
    </div>
  `;
}

function legacyLargeRenderEstExamPanel(department) {
  const store = trainingStore();
  const candidates = state.users.filter((user) => !user.trainings?.EST);
  const activeExam = activeEstExam();
  const candidate = activeExam ? state.users.find((user) => user.id === activeExam.candidateId) : null;
  const secondExaminer = activeExam ? state.users.find((user) => user.id === activeExam.secondExaminerId) : null;
  return `
    <div class="training-exam-layout department-overview-content">
      <section class="panel training-exam-card">
        <div class="panel-header">
          <div><h3>EST Prüfung</h3><p class="muted">Prüfung erstellen und danach in einem eigenen Fenster durchführen.</p></div>
          ${activeExam ? `<span class="requirement-chip ${activeExam.status === "Pausiert" ? "special" : "ok"}">${escapeHtml(activeExam.status)}</span>` : ""}
        </div>
        <div class="exam-start-grid">
          <label>Prüfling ohne EST
            ${renderExamUserPicker("estCandidateInput", "estCandidateList", candidates, "Prüfling auswählen")}
          </label>
          <label>2. Prüfer optional
            <select id="estSecondExaminer"><option value=""></option>${state.users.map((user) => `<option value="${user.id}">${escapeHtml(fullName(user))}</option>`).join("")}</select>
          </label>
          <button class="blue-btn" id="startEstExam" type="button" ${candidates.length ? "" : "disabled"}>EST Prüfung erstellen</button>
        </div>
        ${activeExam ? `
          <div class="active-exam-box">
            <div>
              <strong>${escapeHtml(candidate ? fullName(candidate) : "Unbekannter Prüfling")}</strong>
              <small>Aktive EST-Prüfung · Prüfer: ${escapeHtml(fullName(state.currentUser))}${secondExaminer ? ` · 2. Prüfer: ${escapeHtml(fullName(secondExaminer))}` : ""}</small>
            </div>
            <span class="button-row"><button class="blue-btn" id="continueEstExam" data-exam-id="${escapeHtml(activeExam.id)}" type="button">Prüfungsfenster öffnen</button><button class="ghost-btn training-exam-archive" data-exam-id="${escapeHtml(activeExam.id)}" type="button">Archivieren</button></span>
          </div>
        ` : `<p class="muted">Wähle einen Prüfling aus. Die Durchführung öffnet sich danach als eigenes Fenster.</p>`}
        <div class="exam-module-grid">
          ${store.estModules.map((module, moduleIndex) => `
            <article class="exam-module-card">
              <span>Modul ${moduleIndex + 1}</span>
              <strong>${escapeHtml(module.name)}</strong>
              <small>${escapeHtml(module.description || "Prüfungsmodul")}</small>
            </article>
          `).join("")}
        </div>
      </section>
      ${renderTrainingExamArchive("est", department)}
    </div>
  `;
}

function renderEstExamPanel(department) {
  const store = trainingStore();
  const candidates = state.users.filter((user) => !user.trainings?.EST);
  const activeExam = activeEstExam();
  const candidate = activeExam ? state.users.find((user) => user.id === activeExam.candidateId) : null;
  const secondExaminer = activeExam ? state.users.find((user) => user.id === activeExam.secondExaminerId) : null;
  return `
    <div class="training-exam-layout department-overview-content">
      ${renderActiveTrainingExams("est", department)}
      <section class="panel training-exam-card compact-est-panel">
        <div class="panel-header">
          <div><h3>EST Prüfung</h3><p class="muted">Prüfling auswählen und Prüfung vorbereiten.</p></div>
          ${activeExam ? `<span class="requirement-chip ${activeExam.status === "Pausiert" ? "special" : "ok"}">${escapeHtml(activeExam.status)}</span>` : ""}
        </div>
        <div class="exam-start-grid compact-exam-start est-create-row">
          <label>Prüfling ohne EST
            ${renderExamUserPicker("estCandidateInput", "estCandidateList", candidates, "Prüfling auswählen")}
          </label>
          <button class="blue-btn" id="startEstExam" type="button" ${candidates.length ? "" : "disabled"}>EST Prüfung erstellen</button>
        </div>
        ${activeExam ? `
          <div class="active-exam-box compact-active-exam">
            <div>
              <strong>${escapeHtml(candidate ? fullName(candidate) : "Unbekannter Prüfling")}</strong>
              <small>Aktive EST-Prüfung · Prüfer: ${escapeHtml(fullName(state.currentUser))}${secondExaminer ? ` · 2. Prüfer: ${escapeHtml(fullName(secondExaminer))}` : ""}</small>
            </div>
            <span class="button-row"><button class="blue-btn" id="continueEstExam" data-exam-id="${escapeHtml(activeExam.id)}" type="button">Prüfungsfenster öffnen</button><button class="ghost-btn training-exam-archive" data-exam-id="${escapeHtml(activeExam.id)}" type="button">Archivieren</button></span>
          </div>
        ` : `<p class="muted">Wähle einen Prüfling aus. Die Durchführung öffnet sich danach als eigenes Fenster.</p>`}
        <div class="est-module-strip">
          ${store.estModules.map((module, moduleIndex) => `
            <article class="est-module-chip">
              <span>Modul ${moduleIndex + 1}</span>
              <strong>${escapeHtml(module.name)}</strong>
            </article>
          `).join("")}
        </div>
      </section>
      ${renderTrainingExamArchive("est", department)}
    </div>
  `;
}

function renderModuleExamPanel(department) {
  const store = trainingStore();
  return `
    <div class="training-exam-layout department-overview-content">
      ${renderActiveTrainingExams("module", department)}
      <section class="panel training-exam-card">
        <div class="panel-header"><div><h3>Ausbildungen</h3><p class="muted">Mitglied und Modul auswählen, danach Prüfung öffnen und starten.</p></div></div>
        <div class="module-start-card">
          <label>Mitglied
            ${renderExamUserPicker("moduleCandidateInput", "moduleCandidateList", state.users, "Mitglied auswählen")}
          </label>
          <label>Module auswählen
            <select id="moduleExamSelect">${store.moduleModules.map((module) => `<option value="${module.id}">${escapeHtml(module.name)}</option>`).join("")}</select>
          </label>
          <button class="blue-btn" id="startModuleExam" type="button">Modulprüfung erstellen</button>
        </div>
      </section>
      ${renderTrainingExamArchive("module", department)}
    </div>
  `;
}

function renderTrainingManagementPanels() {
  const store = trainingStore();
  return `
    <section class="training-management-grid">
      <article class="panel training-manage-card">
        <div class="panel-header"><div><h3>EST Verwaltung</h3><p class="muted">Fragenpool für Rechtskunde, Dienstvorschriften, Ortskunde, Helistrecke, Fahrstrecke und Szenario.</p></div></div>
        <div class="exam-module-grid">
          ${store.estModules.map((module) => renderTrainingModuleAdmin("est", module, false)).join("")}
        </div>
      </article>
      <article class="panel training-manage-card">
        <div class="panel-header"><div><h3>Modul Verwaltung</h3><p class="muted">Eigene Module für weitere Ausbildungen anlegen und vorbereiten.</p></div><button class="blue-btn training-module-add" type="button">Modul erstellen</button></div>
        <div class="exam-module-grid">
          ${store.moduleModules.map((module) => renderTrainingModuleAdmin("module", module, true)).join("")}
        </div>
      </article>
    </section>
  `;
}

function renderTrainingModuleAdmin(bank, module, editableModule) {
  return `
    <article class="exam-module-card training-admin-module">
      <div class="training-module-head">
        <div><span>${bank === "est" ? "EST Modul" : "Modul"}</span><strong>${escapeHtml(module.name)}</strong><small>${escapeHtml(module.description || "")}</small></div>
        <div class="button-row">
          ${editableModule ? `<button class="mini-icon training-module-edit" data-module-id="${escapeHtml(module.id)}" type="button" title="Modul bearbeiten">${actionIcon("edit")}</button><button class="mini-icon danger training-module-delete" data-module-id="${escapeHtml(module.id)}" type="button" title="Modul löschen">${actionIcon("delete")}</button>` : ""}
          <button class="blue-btn training-question-add" data-bank="${bank}" data-module-id="${escapeHtml(module.id)}" type="button">Frage hinzufügen</button>
        </div>
      </div>
      <div class="training-question-admin-list">
        ${module.questions.length ? module.questions.map((question) => `
          <div class="training-question-admin-row">
            <span><b>${escapeHtml(question.prompt)}</b><small>${question.type === "location" ? question.stationType === "combat" ? "Combat-Landung / Ort" : module.id === "est-location" ? "Ort" : "Strecke" : question.type === "scenario" ? "Szenario" : "Musterlösung"} · max. ${escapeHtml(question.maxPoints)} Punkt</small></span>
            <button class="mini-icon training-question-edit" data-bank="${bank}" data-module-id="${escapeHtml(module.id)}" data-question-id="${escapeHtml(question.id)}" type="button">${actionIcon("edit")}</button>
            <button class="mini-icon danger training-question-delete" data-bank="${bank}" data-module-id="${escapeHtml(module.id)}" data-question-id="${escapeHtml(question.id)}" type="button">${actionIcon("delete")}</button>
          </div>
        `).join("") : `<p class="muted">Noch keine Fragen.</p>`}
      </div>
    </article>
  `;
}

function findTrainingModule(store, bank, moduleId) {
  const list = bank === "est" ? store.estModules : store.moduleModules;
  return list.find((module) => module.id === moduleId);
}

function openTrainingQuestionModal(bank, moduleId, questionId = null) {
  const store = trainingStore();
  const module = findTrainingModule(store, bank, moduleId);
  const question = module?.questions.find((item) => item.id === questionId);
  if (!module) return;
  const moduleName = cleanText(module.name || "");
  const isOrtskunde = /ortskunde/i.test(moduleName) || module.id === "est-location";
  const isFahrstrecke = /fahrstrecke/i.test(moduleName) || module.id === "est-drive";
  const isHelistrecke = /helistrecke/i.test(moduleName) || module.id === "est-heli";
  const imageEnabled = isOrtskunde || isFahrstrecke || isHelistrecke || question?.type === "location";
  const scenarioEnabled = /szenario/i.test(moduleName) || question?.type === "scenario";
  const defaultType = imageEnabled ? "location" : scenarioEnabled ? "scenario" : "manual";
  const isTimedRoute = isFahrstrecke || isHelistrecke;
  const stationType = question?.stationType || (isHelistrecke && /dach|landung|combat/i.test(question?.prompt || "") ? "combat" : "route");
  const typeLabel = isOrtskunde ? "Ortskunde · Ort mit Bild · max. 1 Punkt" : isFahrstrecke ? "Fahrstrecke · Strecke mit Sollzeit" : isHelistrecke ? "Helistrecke · Route oder Combat-Landung" : scenarioEnabled ? "Szenario" : "Frage mit Musterlösung";
  const promptLabel = isOrtskunde ? "Ort" : isFahrstrecke ? "Strecke" : isHelistrecke ? "Strecke / Combat-Landung" : "Frage";
  const maxPointsValue = isOrtskunde ? 1 : Math.min(10, Number(question?.maxPoints || (isTimedRoute || scenarioEnabled ? 10 : 3)));
  openModal(`
    <h3>${question ? "Frage bearbeiten" : "Frage erstellen"}</h3>
    <p class="muted">${escapeHtml(module.name)}</p>
    <form id="trainingQuestionForm" class="form-grid training-question-form">
      <label class="full">${escapeHtml(promptLabel)}<textarea name="prompt" required>${escapeHtml(question?.prompt || "")}</textarea></label>
      <input type="hidden" name="type" id="trainingQuestionType" value="${escapeHtml(defaultType)}">
      <div class="question-type-display"><span>Fragentyp</span><strong>${escapeHtml(typeLabel)}</strong></div>
      ${isOrtskunde ? `<input type="hidden" name="maxPoints" value="1">` : `<label>Max. Punkte<input name="maxPoints" type="number" min="0.5" max="10" step="0.5" value="${escapeHtml(maxPointsValue)}"></label>`}
      ${isHelistrecke ? `<label>Heli-Eintrag<select name="stationType"><option value="route" ${stationType === "route" ? "selected" : ""}>Strecke</option><option value="combat" ${stationType === "combat" ? "selected" : ""}>Combat-Landung / Ort</option></select></label>` : `<input type="hidden" name="stationType" value="">`}
      ${!imageEnabled ? `<label class="full manual-question-fields scenario-question-fields">Musterlösung / Prüferinfo<textarea name="solution" placeholder="Wird dem Prüfer während der Prüfung angezeigt.">${escapeHtml(question?.solution || "")}</textarea></label>` : `<input type="hidden" name="solution" value="${escapeHtml(question?.solution || "")}">`}
      <textarea name="answers" class="hidden">${escapeHtml((question?.answers || question?.correctAnswers || []).join("\n"))}</textarea>
      ${scenarioEnabled ? `<label class="full scenario-question-fields scenario-big-field">Szenario Ablauf / Prüferinfos<textarea name="scenarioInfo" placeholder="Beschreibe das Szenario ausführlich: Lage, Ablauf, erwartetes Verhalten, Hinweise für Prüfer...">${escapeHtml(question?.scenarioInfo || "")}</textarea></label><label class="full scenario-question-fields">Akte / Maßnahme<textarea name="fileAction" placeholder="Welche Akte, Maßnahme oder Sanktion soll vergeben werden?">${escapeHtml(question?.fileAction || "")}</textarea></label>` : `<textarea name="scenarioInfo" class="hidden">${escapeHtml(question?.scenarioInfo || "")}</textarea><textarea name="fileAction" class="hidden">${escapeHtml(question?.fileAction || "")}</textarea>`}
      ${isTimedRoute ? `<label class="image-question-fields">Sollzeit<input name="targetSeconds" value="${escapeHtml(formatSecondsInput(question?.targetSeconds || 0))}" placeholder="MM:SS oder Sekunden"></label>` : `<input type="hidden" name="targetSeconds" value="">`}
      ${imageEnabled ? `<div class="full image-upload-card">
        <label class="image-upload-drop" id="trainingQuestionImageDrop" title="Bild hochladen">
          <input id="trainingQuestionImage" type="file" accept="image/*">
          <span class="upload-icon">${iconSvg("Plus")}</span>
          <strong>Bild hochladen</strong>
          <small>Drag & Drop oder klicken</small>
        </label>
        <input name="image" type="hidden" value="${escapeHtml(question?.image || "")}">
        <div class="question-image-preview">${question?.image ? `<img src="${escapeHtml(question.image)}" alt="">` : `<span class="muted">Noch kein Bild hinterlegt.</span>`}</div>
      </div>` : `<input name="image" type="hidden" value="${escapeHtml(question?.image || "")}">`}
      <p id="modalError" class="form-error full"></p>
      <div class="modal-actions full">
        <button class="ghost-btn" type="button" data-close>Abbrechen</button>
        <button class="blue-btn" type="submit">Speichern</button>
      </div>
    </form>
  `, (modal) => {
    const handleQuestionImageFile = async (file) => {
      if (!file) return;
      const dataUrl = await readImageFileAsDataUrl(file);
      modal.querySelector("[name='image']").value = dataUrl;
      modal.querySelector(".question-image-preview").innerHTML = `<img src="${escapeHtml(dataUrl)}" alt="">`;
    };
    modal.querySelector("#trainingQuestionImage")?.addEventListener("change", async (event) => {
      await handleQuestionImageFile(event.target.files?.[0]);
    });
    const imageDrop = modal.querySelector("#trainingQuestionImageDrop");
    imageDrop?.addEventListener("dragover", (event) => {
      event.preventDefault();
      imageDrop.classList.add("drag-over");
    });
    imageDrop?.addEventListener("dragleave", () => imageDrop.classList.remove("drag-over"));
    imageDrop?.addEventListener("drop", async (event) => {
      event.preventDefault();
      imageDrop.classList.remove("drag-over");
      await handleQuestionImageFile(event.dataTransfer?.files?.[0]);
    });
    modal.querySelector("#trainingQuestionForm").addEventListener("submit", (event) => {
      event.preventDefault();
      const form = new FormData(event.currentTarget);
      const type = form.get("type");
      const nextQuestion = {
        id: question?.id || makeTrainingId("question"),
        prompt: String(form.get("prompt") || "").trim(),
        type,
        solution: String(form.get("solution") || "").trim(),
        answers: String(form.get("answers") || "").split("\n").map((item) => item.trim()).filter(Boolean),
        correctAnswers: [],
        wrongAnswers: [],
        image: String(form.get("image") || "").trim(),
        scenarioInfo: String(form.get("scenarioInfo") || "").trim(),
        fileAction: String(form.get("fileAction") || "").trim(),
        stationType: String(form.get("stationType") || "").trim(),
        targetSeconds: isTimedRoute ? secondsFromTimeInput(form.get("targetSeconds")) : 0,
        timeSeconds: Number(question?.timeSeconds || 0),
        maxPoints: isOrtskunde ? 1 : Math.min(10, Math.max(0.5, Number(form.get("maxPoints") || 1)))
      };
      if (!nextQuestion.prompt) {
        modal.querySelector("#modalError").textContent = "Bitte eine Frage eintragen.";
        return;
      }
      module.questions = question ? module.questions.map((item) => item.id === question.id ? nextQuestion : item) : [...module.questions, nextQuestion];
      saveTrainingStore(store);
      closeModal();
      renderDepartmentPage(departmentByPage(state.page));
    });
  });
}

function deleteTrainingQuestion(bank, moduleId, questionId, department) {
  const store = trainingStore();
  const module = findTrainingModule(store, bank, moduleId);
  if (!module) return;
  module.questions = module.questions.filter((question) => question.id !== questionId);
  saveTrainingStore(store);
  renderDepartmentPage(department);
  showNotify("Frage gelöscht.", "danger");
}

function openDeleteTrainingQuestionModal(bank, moduleId, questionId, department) {
  const store = trainingStore();
  const module = findTrainingModule(store, bank, moduleId);
  const question = module?.questions.find((item) => item.id === questionId);
  openConfirmModal({
    title: "Frage löschen",
    text: `${question?.prompt || "Diese Frage"} wirklich dauerhaft löschen?`,
    confirmText: "Frage löschen",
    onConfirm: () => deleteTrainingQuestion(bank, moduleId, questionId, department)
  });
}

function readImageFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    if (!file.type.startsWith("image/")) {
      reject(new Error("Bitte eine Bilddatei auswählen."));
      return;
    }
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(new Error("Bild konnte nicht gelesen werden."));
    reader.readAsDataURL(file);
  });
}

function openTrainingModuleModal(moduleId = null) {
  const store = trainingStore();
  const module = store.moduleModules.find((item) => item.id === moduleId);
  openModal(`
    <h3>${module ? "Modul bearbeiten" : "Modul erstellen"}</h3>
    <form id="trainingModuleForm" class="form-grid">
      <label>Name<input name="name" value="${escapeHtml(module?.name || "")}" required></label>
      <label class="full">Beschreibung<textarea name="description">${escapeHtml(module?.description || "")}</textarea></label>
      <p id="modalError" class="form-error full"></p>
      <div class="modal-actions full">
        <button class="ghost-btn" type="button" data-close>Abbrechen</button>
        <button class="blue-btn" type="submit">Speichern</button>
      </div>
    </form>
  `, (modal) => {
    modal.querySelector("#trainingModuleForm").addEventListener("submit", (event) => {
      event.preventDefault();
      const form = new FormData(event.currentTarget);
      const nextModule = {
        id: module?.id || makeTrainingId("module"),
        name: String(form.get("name") || "").trim(),
        description: String(form.get("description") || "").trim(),
        questions: module?.questions || []
      };
      if (!nextModule.name) {
        modal.querySelector("#modalError").textContent = "Bitte einen Modulnamen eintragen.";
        return;
      }
      store.moduleModules = module ? store.moduleModules.map((item) => item.id === module.id ? nextModule : item) : [...store.moduleModules, nextModule];
      saveTrainingStore(store);
      closeModal();
      renderDepartmentPage(departmentByPage(state.page));
    });
  });
}

function deleteTrainingModule(moduleId, department) {
  const store = trainingStore();
  store.moduleModules = store.moduleModules.filter((module) => module.id !== moduleId);
  saveTrainingStore(store);
  renderDepartmentPage(department);
  showNotify("Modul gelöscht.", "danger");
}

function openDeleteTrainingModuleModal(moduleId, department) {
  const store = trainingStore();
  const module = store.moduleModules.find((item) => item.id === moduleId);
  openConfirmModal({
    title: "Modul löschen",
    text: `${module?.name || "Dieses Modul"} wirklich dauerhaft löschen?`,
    confirmText: "Modul löschen",
    onConfirm: () => deleteTrainingModule(moduleId, department)
  });
}

function archiveTrainingExam(examId, department) {
  const store = trainingStore();
  const exam = store.activeExams.find((item) => item.id === examId);
  if (!exam) return;
  exam.status = "Archiviert";
  exam.archivedAt = new Date().toISOString();
  saveTrainingStore(store);
  renderDepartmentPage(department);
  showNotify("Prüfung archiviert.", "success");
}

function pauseTrainingExam(examId, department) {
  const store = trainingStore();
  const exam = store.activeExams.find((item) => item.id === examId);
  if (!exam) return;
  exam.status = "Pausiert";
  saveTrainingStore(store);
  renderDepartmentPage(department);
}

function deleteTrainingExam(examId, department) {
  const store = trainingStore();
  store.activeExams = store.activeExams.filter((exam) => exam.id !== examId);
  saveTrainingStore(store);
  renderDepartmentPage(department);
  showNotify("Prüfung gelöscht.", "danger");
}

function openDeleteTrainingExamModal(examId, department) {
  const store = trainingStore();
  const exam = store.activeExams.find((item) => item.id === examId);
  const candidate = state.users.find((user) => user.id === exam?.candidateId);
  openConfirmModal({
    title: "Prüfung löschen",
    text: `${candidate ? fullName(candidate) : "Diese Prüfung"} wirklich dauerhaft löschen?`,
    confirmText: "Prüfung löschen",
    onConfirm: () => deleteTrainingExam(examId, department)
  });
}

function examProgressText(exam) {
  const module = examCurrentModule(exam);
  return `${module?.name || "-"} · Frage ${exam.questionIndex + 1} von ${module?.questions.length || 0} · Modul ${exam.moduleIndex + 1} von ${exam.modules.length}`;
}

function renderExamQuestionControls(question) {
  if (!question) return "";
  if (question.type === "choice") {
    return `
      <div class="exam-answer-columns">
        <div>
          <strong>Richtige Antworten</strong>
          ${(question.correctAnswers || []).map((answer) => `<label class="exam-check"><input type="checkbox" name="correctAnswer" value="${escapeHtml(answer)}" ${question.selectedCorrect?.includes(answer) ? "checked" : ""}>${escapeHtml(answer)}</label>`).join("") || `<p class="muted">Keine richtigen Antworten hinterlegt.</p>`}
        </div>
        <div>
          <strong>Falsche / fehlende Antworten</strong>
          ${(question.wrongAnswers || []).map((answer) => `<label class="exam-check"><input type="checkbox" name="wrongAnswer" value="${escapeHtml(answer)}" ${question.selectedWrong?.includes(answer) ? "checked" : ""}>${escapeHtml(answer)}</label>`).join("") || `<p class="muted">Keine falschen Antworten hinterlegt.</p>`}
        </div>
      </div>
    `;
  }
  return `
    <div class="manual-solution-box">
      <strong>Musterlösung</strong>
      <p>${escapeHtml(question.solution || "Keine Musterlösung hinterlegt.")}</p>
    </div>
    <label class="full">Antwort des Prüflings<textarea id="examTraineeAnswer" placeholder="Antwort mitschreiben">${escapeHtml(question.traineeAnswer || "")}</textarea></label>
  `;
}

function renderExamReview(exam) {
  const total = exam.modules.flatMap((module) => module.questions).reduce((sum, question) => sum + Number(question.maxPoints || 1), 0);
  const scored = exam.modules.flatMap((module) => module.questions).reduce((sum, question) => sum + Number(question.manualPoints ?? question.result?.points ?? 0), 0);
  const percent = total ? Math.round((scored / total) * 100) : 0;
  return `
    <div class="exam-review-list">
      ${exam.modules.map((module) => `
        <section class="exam-review-module">
          <h4>${escapeHtml(module.name)}</h4>
          ${module.questions.map((question) => `
            <label class="exam-review-row">
              <span><b>${escapeHtml(question.prompt)}</b><small>Max. ${escapeHtml(question.maxPoints)} Punkte${question.type === "choice" ? " · automatisch vorbereitet" : " · manuell bewerten"}</small></span>
              <select data-review-score="${escapeHtml(question.id)}">
                ${[0, 0.5, 1, 1.5, 2].filter((value) => value <= Number(question.maxPoints || 1)).map((value) => `<option value="${value}" ${Number(question.manualPoints ?? question.result?.points ?? 0) === value ? "selected" : ""}>${value} Punkte</option>`).join("")}
              </select>
            </label>
          `).join("")}
        </section>
      `).join("")}
      <div class="exam-result-preview"><span><b>Zwischenstand</b>${scored} von ${total} Punkten</span><span class="${percent >= 75 ? "result-pass" : "result-fail"}">${percent}%</span></div>
    </div>
  `;
}

function openTrainingExamModal(examId) {
  const store = trainingStore();
  const exam = store.activeExams.find((item) => item.id === examId);
  if (!exam) return;
  const candidate = state.users.find((user) => user.id === exam.candidateId);
  const module = examCurrentModule(exam);
  const question = examCurrentQuestion(exam);
  const review = Boolean(exam.reviewMode);
  openModal(`
    <h3>${exam.kind === "est" ? "EST Prüfung" : "Modul Prüfung"}</h3>
    <p class="muted">${escapeHtml(candidate ? fullName(candidate) : "Unbekannter Prüfling")} · ${escapeHtml(examProgressText(exam))}</p>
    ${exam.finalResult ? `
      <div class="exam-result-preview">
        <span><b>Finales Ergebnis</b>${exam.finalResult.points} von ${exam.finalResult.total} Punkten</span>
        <span class="${exam.finalResult.percent >= 75 ? "result-pass" : "result-fail"}">${exam.finalResult.percent}% · ${exam.finalResult.percent >= 75 ? "Bestanden" : "Nicht bestanden"}</span>
      </div>
    ` : review ? renderExamReview(exam) : `
      <section class="exam-runner-card">
        <span>${escapeHtml(module?.name || "-")}</span>
        <h4>${escapeHtml(question?.prompt || "Keine Frage vorhanden")}</h4>
        ${renderExamQuestionControls(question)}
      </section>
    `}
    <div class="modal-actions">
      <button class="ghost-btn" id="pauseExamRunner" type="button">Zwischenspeichern & schließen</button>
      ${!exam.finalResult && !review ? `<button class="blue-btn" id="nextExamQuestion" type="button">Frage speichern / weiter</button><button class="orange-btn" id="startExamReview" type="button">Auswerten</button>` : ""}
      ${review ? `<button class="blue-btn" id="finishExamReview" type="button">Final auswerten</button>` : ""}
    </div>
  `, (modal) => {
    modal.classList.add("exam-modal");
    modal.querySelector("#examSecondExaminer")?.addEventListener("change", (event) => {
      exam.secondExaminerId = event.target.value;
      saveActiveTrainingExam(exam);
    });
    modal.querySelector("#beginExamRunner")?.addEventListener("click", () => {
      exam.secondExaminerId = modal.querySelector("#examSetupSecondExaminer")?.value || "";
      exam.status = "Laufend";
      exam.startedAt = new Date().toISOString();
      saveActiveTrainingExam(exam);
      openTrainingExamModal(exam.id);
      renderDepartmentPage(departmentByPage(state.page));
    });
    modal.querySelector("#pauseExamRunner")?.addEventListener("click", () => {
      exam.status = "Pausiert";
      saveActiveTrainingExam(exam);
      closeModal();
      renderDepartmentPage(departmentByPage(state.page));
    });
    const saveQuestion = () => {
      if (!question) return;
      if (question.type === "choice") {
        question.selectedCorrect = Array.from(modal.querySelectorAll("[name='correctAnswer']:checked")).map((input) => input.value);
        question.selectedWrong = Array.from(modal.querySelectorAll("[name='wrongAnswer']:checked")).map((input) => input.value);
        const correctCount = question.correctAnswers?.length || 0;
        const hitCount = question.selectedCorrect.length;
        const wrongCount = question.selectedWrong.length;
        const ratio = correctCount ? Math.max(0, (hitCount - wrongCount * 0.5) / correctCount) : 0;
        question.manualPoints = Math.min(Number(question.maxPoints || 1), Math.round(ratio * Number(question.maxPoints || 1) * 2) / 2);
        question.result = { points: question.manualPoints };
      } else {
        question.traineeAnswer = modal.querySelector("#examTraineeAnswer")?.value || "";
      }
      exam.status = "Laufend";
      saveActiveTrainingExam(exam);
    };
    modal.querySelector("#nextExamQuestion")?.addEventListener("click", () => {
      saveQuestion();
      if (exam.questionIndex < (examCurrentModule(exam)?.questions.length || 0) - 1) exam.questionIndex += 1;
      else if (exam.moduleIndex < exam.modules.length - 1) {
        exam.moduleIndex += 1;
        exam.questionIndex = 0;
      } else {
        exam.reviewMode = true;
      }
      saveActiveTrainingExam(exam);
      openTrainingExamModal(exam.id);
    });
    modal.querySelector("#startExamReview")?.addEventListener("click", () => {
      saveQuestion();
      exam.reviewMode = true;
      saveActiveTrainingExam(exam);
      openTrainingExamModal(exam.id);
    });
    modal.querySelector("#finishExamReview")?.addEventListener("click", () => {
      modal.querySelectorAll("[data-review-score]").forEach((select) => {
        exam.modules.forEach((reviewModule) => reviewModule.questions.forEach((reviewQuestion) => {
          if (reviewQuestion.id === select.dataset.reviewScore) reviewQuestion.manualPoints = Number(select.value);
        }));
      });
      const questions = exam.modules.flatMap((reviewModule) => reviewModule.questions);
      const total = questions.reduce((sum, item) => sum + Number(item.maxPoints || 1), 0);
      const points = questions.reduce((sum, item) => sum + Number(item.manualPoints || 0), 0);
      exam.finalResult = { total, points, percent: total ? Math.round((points / total) * 100) : 0 };
      exam.status = "Abgeschlossen";
      saveActiveTrainingExam(exam);
      openTrainingExamModal(exam.id);
      renderDepartmentPage(departmentByPage(state.page));
    });
  });
}

function renderExamArchiveDetail(exam) {
  return `
    <div class="exam-review-list">
      ${exam.modules.map((module) => `
        <section class="exam-review-module">
          <h4>${escapeHtml(module.name)}</h4>
          ${module.questions.map((question) => `
            <div class="exam-review-row">
              <span>
                <b>${escapeHtml(question.prompt)}</b>
                <small>${question.type === "choice" ? `Auswahl richtig: ${(question.selectedCorrect || []).join(", ") || "-"} · Auswahl falsch/fehlend: ${(question.selectedWrong || []).join(", ") || "-"}` : `Antwort: ${question.traineeAnswer || "-"}`}</small>
              </span>
              <strong>${escapeHtml(question.manualPoints ?? question.result?.points ?? 0)} / ${escapeHtml(question.maxPoints || 1)} Punkte</strong>
            </div>
          `).join("")}
        </section>
      `).join("")}
    </div>
  `;
}

function openTrainingExamModal(examId, readOnly = false) {
  const store = trainingStore();
  const exam = store.activeExams.find((item) => item.id === examId);
  if (!exam) return;
  const candidate = state.users.find((user) => user.id === exam.candidateId);
  const module = examCurrentModule(exam);
  const question = examCurrentQuestion(exam);
  const review = Boolean(exam.reviewMode);
  const archiveView = readOnly || exam.status === "Archiviert";
  const isSetup = exam.status === "Vorbereitung";
  openModal(`
    <h3>${exam.kind === "est" ? "EST Prüfung" : "Modul Prüfung"}</h3>
    <p class="muted">${escapeHtml(candidate ? fullName(candidate) : "Unbekannter Prüfling")} · ${escapeHtml(examProgressText(exam))}</p>
    ${!archiveView && !isSetup ? `<div class="exam-runner-meta"><label>2. Prüfer<select id="examSecondExaminer"><option value=""></option>${state.users.map((user) => `<option value="${user.id}" ${exam.secondExaminerId === user.id ? "selected" : ""}>${escapeHtml(fullName(user))}</option>`).join("")}</select></label><span><b>Dauer</b><i class="exam-live-timer" data-started-at="${escapeHtml(exam.startedAt || "")}">${escapeHtml(examElapsedText(exam))}</i></span></div>` : ""}
    ${archiveView ? renderExamArchiveDetail(exam) : isSetup ? `
      <section class="exam-runner-card">
        <span>Prüfung vorbereiten</span>
        <h4>${escapeHtml(candidate ? fullName(candidate) : "Unbekannter Prüfling")}</h4>
        <label>2. Prüfer optional<select id="examSetupSecondExaminer"><option value=""></option>${state.users.map((user) => `<option value="${user.id}" ${exam.secondExaminerId === user.id ? "selected" : ""}>${escapeHtml(fullName(user))}</option>`).join("")}</select></label>
        <p class="muted">Erst nach dem Start werden Fragen angezeigt und der Timer läuft.</p>
      </section>
    ` : exam.finalResult ? `
      <div class="exam-result-preview">
        <span><b>Finales Ergebnis</b>${exam.finalResult.points} von ${exam.finalResult.total} Punkten</span>
        <span class="${exam.finalResult.percent >= 75 ? "result-pass" : "result-fail"}">${exam.finalResult.percent}% · ${exam.finalResult.percent >= 75 ? "Bestanden" : "Nicht bestanden"}</span>
      </div>
      ${renderExamArchiveDetail(exam)}
    ` : review ? renderExamReview(exam) : `
      <section class="exam-runner-card">
        <span>${escapeHtml(module?.name || "-")}</span>
        <h4>${escapeHtml(question?.prompt || "Keine Frage vorhanden")}</h4>
        ${renderExamQuestionControls(question)}
      </section>
    `}
    <div class="modal-actions">
      <button class="ghost-btn" id="pauseExamRunner" type="button">${archiveView ? "Schließen" : isSetup ? "Abbrechen" : "Zwischenspeichern & schließen"}</button>
      ${!archiveView && isSetup ? `<button class="blue-btn" id="beginExamRunner" type="button">Prüfung starten</button>` : ""}
      ${!archiveView && !isSetup && !exam.finalResult && !review ? `<button class="blue-btn" id="nextExamQuestion" type="button">Frage speichern / weiter</button><button class="orange-btn" id="startExamReview" type="button">Auswerten</button>` : ""}
      ${!archiveView && review ? `<button class="blue-btn" id="finishExamReview" type="button">Final auswerten</button>` : ""}
    </div>
  `, (modal) => {
    modal.classList.add("exam-modal");
    if (isSetup) modal.classList.add("setup-exam-modal");
    modal.querySelector("#beginExamRunner")?.addEventListener("click", () => {
      exam.secondExaminerId = modal.querySelector("#examSetupSecondExaminer")?.value || "";
      exam.status = "Laufend";
      exam.startedAt = new Date().toISOString();
      saveActiveTrainingExam(exam);
      openTrainingExamModal(exam.id);
      renderDepartmentPage(departmentByPage(state.page));
    });
    modal.querySelector("#pauseExamRunner")?.addEventListener("click", () => {
      if (archiveView) {
        closeModal();
        return;
      }
      if (isSetup) {
        store.activeExams = store.activeExams.filter((item) => item.id !== exam.id);
        saveTrainingStore(store);
        closeModal();
        renderDepartmentPage(departmentByPage(state.page));
        return;
      }
      exam.status = "Pausiert";
      saveActiveTrainingExam(exam);
      closeModal();
      renderDepartmentPage(departmentByPage(state.page));
    });
    const saveQuestion = () => {
      if (!question) return;
      if (question.type === "choice") {
        question.selectedCorrect = Array.from(modal.querySelectorAll("[name='correctAnswer']:checked")).map((input) => input.value);
        question.selectedWrong = Array.from(modal.querySelectorAll("[name='wrongAnswer']:checked")).map((input) => input.value);
        const correctCount = question.correctAnswers?.length || 0;
        const hitCount = question.selectedCorrect.length;
        const wrongCount = question.selectedWrong.length;
        const ratio = correctCount ? Math.max(0, (hitCount - wrongCount * 0.5) / correctCount) : 0;
        question.manualPoints = Math.min(Number(question.maxPoints || 1), Math.round(ratio * Number(question.maxPoints || 1) * 2) / 2);
        question.result = { points: question.manualPoints };
      } else {
        question.traineeAnswer = modal.querySelector("#examTraineeAnswer")?.value || "";
      }
      exam.status = "Laufend";
      saveActiveTrainingExam(exam);
    };
    modal.querySelector("#nextExamQuestion")?.addEventListener("click", () => {
      saveQuestion();
      if (exam.questionIndex < (examCurrentModule(exam)?.questions.length || 0) - 1) exam.questionIndex += 1;
      else if (exam.moduleIndex < exam.modules.length - 1) {
        exam.moduleIndex += 1;
        exam.questionIndex = 0;
      } else {
        exam.reviewMode = true;
      }
      saveActiveTrainingExam(exam);
      openTrainingExamModal(exam.id);
    });
    modal.querySelector("#startExamReview")?.addEventListener("click", () => {
      saveQuestion();
      exam.reviewMode = true;
      saveActiveTrainingExam(exam);
      openTrainingExamModal(exam.id);
    });
    modal.querySelector("#finishExamReview")?.addEventListener("click", () => {
      modal.querySelectorAll("[data-review-score]").forEach((select) => {
        exam.modules.forEach((reviewModule) => reviewModule.questions.forEach((reviewQuestion) => {
          if (reviewQuestion.id === select.dataset.reviewScore) reviewQuestion.manualPoints = Number(select.value);
        }));
      });
      const questions = exam.modules.flatMap((reviewModule) => reviewModule.questions);
      const total = questions.reduce((sum, item) => sum + Number(item.maxPoints || 1), 0);
      const points = questions.reduce((sum, item) => sum + Number(item.manualPoints || 0), 0);
      exam.finalResult = { total, points, percent: total ? Math.round((points / total) * 100) : 0 };
      exam.status = "Abgeschlossen";
      exam.archivedAt = new Date().toISOString();
      saveActiveTrainingExam(exam);
      openTrainingExamModal(exam.id);
      renderDepartmentPage(departmentByPage(state.page));
    });
  });
}

function ensureExamModuleState(exam) {
  if (!exam || !Array.isArray(exam.modules)) return exam;
  exam.moduleIndex = Math.max(0, Math.min(Number(exam.moduleIndex || 0), Math.max(0, exam.modules.length - 1)));
  exam.questionIndex = Math.max(0, Number(exam.questionIndex || 0));
  exam.modules.forEach((module, moduleIndex) => {
    if (!module.status) {
      if (exam.status === "Vorbereitung") module.status = "Offen";
      else if (moduleIndex < exam.moduleIndex) module.status = "Abgeschlossen";
      else if (moduleIndex === exam.moduleIndex && exam.reviewMode) module.status = "Auswertung";
      else if (moduleIndex === exam.moduleIndex) module.status = "Laufend";
      else module.status = "Offen";
    }
    module.questions = module.questions || [];
    module.questions.forEach((question) => {
      if (!Array.isArray(question.selectedAnswers)) {
        question.selectedAnswers = [...(question.selectedCorrect || []), ...(question.selectedWrong || [])];
      }
      question.questionPenalty = Boolean(question.questionPenalty);
      question.traineeAnswer = question.traineeAnswer || "";
      question.manualPoints = Number(question.manualPoints ?? question.result?.points ?? 0);
    });
  });
  return exam;
}

function currentManagedExamModule(exam) {
  ensureExamModuleState(exam);
  return exam?.modules?.[exam.moduleIndex] || null;
}

function currentManagedExamQuestion(exam) {
  const module = currentManagedExamModule(exam);
  const questionCount = module?.questions?.length || 0;
  exam.questionIndex = Math.max(0, Math.min(Number(exam.questionIndex || 0), Math.max(0, questionCount - 1)));
  return module?.questions?.[exam.questionIndex] || null;
}

function examModuleTotal(module) {
  return (module?.questions || []).reduce((sum, question) => sum + Number(question.maxPoints || 1), 0);
}

function examModulePoints(module) {
  return (module?.questions || []).reduce((sum, question) => sum + Number(question.manualPoints ?? question.result?.points ?? 0), 0);
}

function examModulePercent(module) {
  const total = examModuleTotal(module);
  return total ? Math.round((examModulePoints(module) / total) * 100) : 0;
}

function normalizeChoiceAnswers(question) {
  return Array.from(new Set([...(question.correctAnswers || []), ...(question.wrongAnswers || [])].filter(Boolean)));
}

function scoreChoiceQuestion(question) {
  const maxPoints = Number(question.maxPoints || 1);
  const correctAnswers = question.correctAnswers || [];
  const selectedAnswers = question.selectedAnswers || [];
  const hits = correctAnswers.filter((answer) => selectedAnswers.includes(answer)).length;
  const base = correctAnswers.length ? (hits / correctAnswers.length) * maxPoints : 0;
  const points = Math.round((base - (question.questionPenalty ? 1 : 0)) * 2) / 2;
  return Math.max(-1, Math.min(maxPoints, points));
}

function renderExamModuleStepper(exam) {
  ensureExamModuleState(exam);
  return `
    <div class="exam-module-stepper">
      ${exam.modules.map((module, index) => `
        <span class="${index === exam.moduleIndex ? "active" : ""} ${module.status === "Abgeschlossen" ? "done" : ""}">
          <b>${escapeHtml(module.name)}</b>
          <small>${escapeHtml(module.status || "Offen")}${module.result ? ` · ${escapeHtml(module.result.percent)}%` : ""}</small>
        </span>
      `).join("")}
    </div>
  `;
}

function renderExamQuestionControls(question) {
  if (!question) return `<p class="muted">Keine Frage vorhanden.</p>`;
  if (question.type === "choice") {
    const answers = normalizeChoiceAnswers(question);
    return `
      <div class="exam-answer-list">
        <strong>Antworten des Prüflings markieren</strong>
        ${answers.length ? answers.map((answer) => `
          <label class="exam-check">
            <input data-autosave-exam type="checkbox" name="answerOption" value="${escapeHtml(answer)}" ${question.selectedAnswers?.includes(answer) ? "checked" : ""}>
            ${escapeHtml(answer)}
          </label>
        `).join("") : `<p class="muted">Keine Antwortmöglichkeiten hinterlegt.</p>`}
        <label class="exam-check penalty">
          <input data-autosave-exam type="checkbox" name="questionPenalty" ${question.questionPenalty ? "checked" : ""}>
          Frage falsch beantwortet (-1 Punkt)
        </label>
      </div>
    `;
  }
  return `
    <div class="manual-solution-box">
      <strong>Musterlösung für den Prüfer</strong>
      <p>${escapeHtml(question.solution || "Keine Musterlösung hinterlegt.")}</p>
    </div>
    <label class="full">Antwort des Prüflings<textarea data-autosave-exam id="examTraineeAnswer" placeholder="Antwort mitschreiben">${escapeHtml(question.traineeAnswer || "")}</textarea></label>
  `;
}

function renderExamAnswerSummary(question) {
  if (question.type === "choice") {
    const selected = question.selectedAnswers || [];
    const missing = (question.correctAnswers || []).filter((answer) => !selected.includes(answer));
    return `
      <small><b>Ausgewählt:</b> ${escapeHtml(selected.join(", ") || "-")}</small>
      <small><b>Nicht genannt:</b> ${escapeHtml(missing.join(", ") || "-")}</small>
      <small><b>Frage falsch beantwortet:</b> ${question.questionPenalty ? "Ja (-1 Punkt)" : "Nein"}</small>
    `;
  }
  return `
    <small><b>Antwort Prüfling:</b> ${escapeHtml(question.traineeAnswer || "-")}</small>
    <small><b>Musterlösung:</b> ${escapeHtml(question.solution || "-")}</small>
  `;
}

function renderExamReview(exam) {
  ensureExamModuleState(exam);
  const module = currentManagedExamModule(exam);
  const total = examModuleTotal(module);
  const scored = examModulePoints(module);
  const percent = examModulePercent(module);
  return `
    ${renderExamModuleStepper(exam)}
    <div class="exam-review-list">
      <section class="exam-review-module">
        <h4>${escapeHtml(module?.name || "Modul")}</h4>
        ${(module?.questions || []).map((question, index) => `
          <label class="exam-review-row detailed">
            <span>
              <b>${index + 1}. ${escapeHtml(question.prompt)}</b>
              ${renderExamAnswerSummary(question)}
              <small>Max. ${escapeHtml(question.maxPoints || 1)} Punkte</small>
            </span>
            <select data-review-score="${escapeHtml(question.id)}">
              ${[-1, 0, 0.5, 1, 1.5, 2].filter((value) => value <= Number(question.maxPoints || 1)).map((value) => `<option value="${value}" ${Number(question.manualPoints ?? question.result?.points ?? 0) === value ? "selected" : ""}>${value} Punkte</option>`).join("")}
            </select>
          </label>
        `).join("") || `<p class="muted">Keine Fragen in diesem Modul.</p>`}
      </section>
      <div class="exam-result-preview"><span><b>Modul-Zwischenstand</b>${scored} von ${total} Punkten</span><span class="${percent >= 75 ? "result-pass" : "result-fail"}">${percent}%</span></div>
    </div>
  `;
}

function renderExamArchiveDetail(exam) {
  ensureExamModuleState(exam);
  return `
    <div class="exam-review-list archive-detail">
      ${exam.modules.map((module, moduleIndex) => {
        const points = module.result?.points ?? examModulePoints(module);
        const total = module.result?.total ?? examModuleTotal(module);
        const percent = module.result?.percent ?? examModulePercent(module);
        return `
          <section class="exam-review-module archive-module-block">
            <div class="archive-module-head">
              <h4>Modul ${moduleIndex + 1}: ${escapeHtml(module.name)}</h4>
              <span class="${percent >= 75 ? "result-pass" : "result-fail"}">${escapeHtml(percent)}% · ${escapeHtml(points)} / ${escapeHtml(total)} Punkte</span>
            </div>
            ${(module.questions || []).map((question, questionIndex) => `
              <div class="exam-review-row detailed archive-question-row">
                <span>
                  <b>${questionIndex + 1}. ${escapeHtml(question.prompt)}</b>
                  ${renderExamAnswerSummary(question)}
                </span>
                <strong>${escapeHtml(question.manualPoints ?? question.result?.points ?? 0)} / ${escapeHtml(question.maxPoints || 1)} Punkte</strong>
              </div>
            `).join("") || `<p class="muted">Keine Fragen gespeichert.</p>`}
          </section>
        `;
      }).join("")}
    </div>
  `;
}

function examProgressText(exam) {
  ensureExamModuleState(exam);
  const module = currentManagedExamModule(exam);
  const questionCount = module?.questions?.length || 0;
  const status = module?.status || exam.status || "-";
  return `${module?.name || "-"} · ${status} · Frage ${Math.min(Number(exam.questionIndex || 0) + 1, Math.max(1, questionCount))} von ${questionCount} · Modul ${Number(exam.moduleIndex || 0) + 1} von ${exam.modules?.length || 0}`;
}

function renderExamModuleStart(exam, candidate) {
  const module = currentManagedExamModule(exam);
  const completed = exam.modules.filter((item) => item.status === "Abgeschlossen").length;
  return `
    ${renderExamModuleStepper(exam)}
    <section class="exam-runner-card exam-module-start-card">
      <span>${exam.status === "Vorbereitung" ? "Prüfung vorbereiten" : "Nächstes Modul bereit"}</span>
      <h4>${escapeHtml(module?.name || "Modul")}</h4>
      <p class="muted">${escapeHtml(candidate ? fullName(candidate) : "Unbekannter Prüfling")} · ${completed} von ${exam.modules.length} Modulen abgeschlossen</p>
      <label>2. Prüfer optional<select id="examSetupSecondExaminer"><option value=""></option>${state.users.map((user) => `<option value="${user.id}" ${exam.secondExaminerId === user.id ? "selected" : ""}>${escapeHtml(fullName(user))}</option>`).join("")}</select></label>
      <p class="muted">Erst nach dem Modulstart werden Fragen angezeigt. Jedes Modul wird separat ausgewertet und danach manuell fortgesetzt.</p>
    </section>
  `;
}

function openTrainingExamModal(examId, readOnly = false) {
  const store = trainingStore();
  const exam = store.activeExams.find((item) => item.id === examId);
  if (!exam) return;
  ensureExamModuleState(exam);
  const candidate = state.users.find((user) => user.id === exam.candidateId);
  const module = currentManagedExamModule(exam);
  const question = currentManagedExamQuestion(exam);
  const archiveView = readOnly || exam.status === "Archiviert";
  const isFinal = exam.status === "Abgeschlossen" || Boolean(exam.finalResult);
  const isSetup = !archiveView && !isFinal && (exam.status === "Vorbereitung" || module?.status === "Offen" || exam.status === "Modul bereit");
  const isReview = !archiveView && !isFinal && (exam.reviewMode || module?.status === "Auswertung");
  const isActive = !archiveView && !isFinal && !isSetup && !isReview;
  openModal(`
    <h3>${exam.kind === "est" ? "EST Prüfung" : "Modul Prüfung"}</h3>
    <p class="muted">${escapeHtml(candidate ? fullName(candidate) : "Unbekannter Prüfling")} · ${escapeHtml(examProgressText(exam))}</p>
    ${!archiveView && !isSetup && !isFinal ? `<div class="exam-runner-meta"><label>2. Prüfer<select id="examSecondExaminer"><option value=""></option>${state.users.map((user) => `<option value="${user.id}" ${exam.secondExaminerId === user.id ? "selected" : ""}>${escapeHtml(fullName(user))}</option>`).join("")}</select></label><span><b>Dauer</b><i class="exam-live-timer" data-started-at="${escapeHtml(exam.startedAt || "")}">${escapeHtml(examElapsedText(exam))}</i></span></div>` : ""}
    ${archiveView ? renderExamArchiveDetail(exam) : isSetup ? renderExamModuleStart(exam, candidate) : isFinal ? `
      <div class="exam-result-preview">
        <span><b>Finales Ergebnis</b>${exam.finalResult.points} von ${exam.finalResult.total} Punkten</span>
        <span class="${exam.finalResult.percent >= 75 ? "result-pass" : "result-fail"}">${exam.finalResult.percent}% · ${exam.finalResult.percent >= 75 ? "Bestanden" : "Nicht bestanden"}</span>
      </div>
      ${renderExamArchiveDetail(exam)}
    ` : isReview ? renderExamReview(exam) : `
      ${renderExamModuleStepper(exam)}
      <section class="exam-runner-card">
        <span>${escapeHtml(module?.name || "-")}</span>
        <h4>${escapeHtml(question?.prompt || "Keine Frage vorhanden")}</h4>
        ${renderExamQuestionControls(question)}
      </section>
    `}
    <div class="modal-actions">
      <button class="ghost-btn" id="pauseExamRunner" type="button">${archiveView || isFinal ? "Schließen" : isSetup ? "Abbrechen" : "Schließen"}</button>
      ${isSetup ? `<button class="blue-btn" id="beginExamRunner" type="button">${exam.status === "Vorbereitung" ? "Prüfung starten" : "Modul starten"}</button>` : ""}
      ${isActive ? `<button class="blue-btn" id="nextExamQuestion" type="button">${exam.questionIndex < (module?.questions.length || 0) - 1 ? "Frage speichern / weiter" : "Modul auswerten"}</button>` : ""}
      ${isReview ? `<button class="blue-btn" id="finishExamReview" type="button">Modul final auswerten</button>` : ""}
    </div>
  `, (modal) => {
    modal.classList.add("exam-modal");
    if (isSetup) modal.classList.add("setup-exam-modal");
    const persist = () => {
      ensureExamModuleState(exam);
      saveActiveTrainingExam(exam);
    };
    const saveQuestion = () => {
      const activeQuestion = currentManagedExamQuestion(exam);
      if (!activeQuestion) return;
      if (activeQuestion.type === "choice") {
        activeQuestion.selectedAnswers = Array.from(modal.querySelectorAll("[name='answerOption']:checked")).map((input) => input.value);
        activeQuestion.questionPenalty = Boolean(modal.querySelector("[name='questionPenalty']")?.checked);
        activeQuestion.manualPoints = scoreChoiceQuestion(activeQuestion);
        activeQuestion.result = { points: activeQuestion.manualPoints };
      } else {
        activeQuestion.traineeAnswer = modal.querySelector("#examTraineeAnswer")?.value || "";
      }
      if (module) module.status = "Laufend";
      if (!exam.startedAt) exam.startedAt = new Date().toISOString();
      exam.status = "Laufend";
      persist();
    };
    modal.querySelector("#examSecondExaminer")?.addEventListener("change", (event) => {
      exam.secondExaminerId = event.target.value;
      persist();
    });
    modal.querySelector("#beginExamRunner")?.addEventListener("click", () => {
      exam.secondExaminerId = modal.querySelector("#examSetupSecondExaminer")?.value || "";
      exam.status = "Laufend";
      if (!exam.startedAt) exam.startedAt = new Date().toISOString();
      if (module) {
        module.status = "Laufend";
        module.startedAt = module.startedAt || new Date().toISOString();
      }
      exam.reviewMode = false;
      exam.questionIndex = 0;
      persist();
      openTrainingExamModal(exam.id);
      renderDepartmentPage(departmentByPage(state.page));
    });
    modal.querySelector("#pauseExamRunner")?.addEventListener("click", () => {
      if (archiveView || isFinal) {
        closeModal();
        return;
      }
      if (isSetup && exam.status === "Vorbereitung") {
        store.activeExams = store.activeExams.filter((item) => item.id !== exam.id);
        saveTrainingStore(store);
        closeModal();
        renderDepartmentPage(departmentByPage(state.page));
        return;
      }
      if (isActive) saveQuestion();
      closeModal();
      renderDepartmentPage(departmentByPage(state.page));
    });
    modal.querySelectorAll("[data-autosave-exam]").forEach((input) => {
      input.addEventListener(input.tagName === "TEXTAREA" ? "input" : "change", saveQuestion);
    });
    modal.querySelector("#nextExamQuestion")?.addEventListener("click", () => {
      saveQuestion();
      if (exam.questionIndex < (module?.questions.length || 0) - 1) {
        exam.questionIndex += 1;
      } else if (module) {
        module.status = "Auswertung";
        exam.reviewMode = true;
      }
      persist();
      openTrainingExamModal(exam.id);
    });
    modal.querySelector("#finishExamReview")?.addEventListener("click", () => {
      modal.querySelectorAll("[data-review-score]").forEach((select) => {
        const reviewQuestion = (module?.questions || []).find((item) => item.id === select.dataset.reviewScore);
        if (reviewQuestion) {
          reviewQuestion.manualPoints = Number(select.value);
          reviewQuestion.result = { points: reviewQuestion.manualPoints };
        }
      });
      if (module) {
        module.status = "Abgeschlossen";
        module.completedAt = new Date().toISOString();
        module.result = {
          total: examModuleTotal(module),
          points: examModulePoints(module),
          percent: examModulePercent(module)
        };
      }
      const nextIndex = exam.modules.findIndex((item) => item.status !== "Abgeschlossen");
      exam.reviewMode = false;
      if (nextIndex >= 0) {
        exam.moduleIndex = nextIndex;
        exam.questionIndex = 0;
        exam.status = "Modul bereit";
        exam.modules[nextIndex].status = "Offen";
      } else {
        const total = exam.modules.reduce((sum, item) => sum + examModuleTotal(item), 0);
        const points = exam.modules.reduce((sum, item) => sum + examModulePoints(item), 0);
        exam.finalResult = { total, points, percent: total ? Math.round((points / total) * 100) : 0 };
        exam.status = "Abgeschlossen";
        exam.archivedAt = new Date().toISOString();
      }
      persist();
      openTrainingExamModal(exam.id);
      renderDepartmentPage(departmentByPage(state.page));
    });
  });
}

function shuffledItems(items) {
  return [...items].sort(() => Math.random() - 0.5);
}

function isEstLocationModule(module) {
  return /ortskunde|fahrstrecke/i.test(cleanText(module?.name || "")) || ["est-location", "est-drive"].includes(module?.id);
}

function isLocationQuestion(question, side = "main") {
  return side === "location" || question?.type === "location";
}

function orderedEstModules(modules = []) {
  const order = ["est-law", "est-location", "est-scenario", "est-rules", "est-drive", "est-heli"];
  return [...modules].sort((a, b) => {
    const ai = order.indexOf(a.id);
    const bi = order.indexOf(b.id);
    return (ai === -1 ? 999 : ai) - (bi === -1 ? 999 : bi);
  });
}

function createTrainingExam(kind, candidateId, secondExaminerId, modules) {
  const exam = {
    id: makeTrainingId("exam"),
    kind,
    candidateId,
    examinerId: state.currentUser?.id,
    secondExaminerId,
    status: "Vorbereitung",
    moduleIndex: 0,
    questionIndex: 0,
    reviewMode: false,
    finalResult: null,
    createdAt: new Date().toISOString(),
    startedAt: null,
    activeMainModuleId: "",
    modules: (kind === "est" ? orderedEstModules(modules) : modules).map((module) => ({
      id: module.id,
      name: module.name,
      description: module.description,
      phase: module.phase || 0,
      status: "Offen",
      questions: module.questions.map((question) => ({ ...question, result: null, traineeAnswer: "", selectedCorrect: [], selectedWrong: [], selectedAnswers: [], manualPoints: 0, stationType: question.stationType || "", timeSeconds: Number(question.timeSeconds || 0), targetSeconds: Number(question.targetSeconds || 0), questionPenalty: false, penaltyPoints: 0, skipped: false }))
    }))
  };
  if (kind === "est") prepareEstExamModules(exam);
  return exam;
}

function prepareEstExamModules(exam) {
  if (!exam || exam.kind !== "est" || exam.locationRandomized) return exam;
  exam.modules.filter((module) => isEstLocationModule(module) || module.id === "est-heli").forEach((sideModule) => {
    if (!sideModule.questions?.length && sideModule.id === "est-location") {
      sideModule.questions = EST_LOCATION_PROMPTS.map((place) => defaultTrainingQuestion(place, "location"));
    }
    const timed = ["est-drive", "est-heli"].includes(sideModule.id);
    sideModule.questions = shuffledItems(sideModule.questions || []).map((question) => ({
      ...question,
      type: "location",
      stationType: question.stationType || (sideModule.id === "est-heli" && /dach|landung|combat/i.test(question.prompt || "") ? "combat" : sideModule.id === "est-heli" ? "route" : ""),
      maxPoints: timed ? Number(question.maxPoints || 10) : 1,
      targetSeconds: Number(question.targetSeconds || 0),
      timeSeconds: Number(question.timeSeconds || 0),
      manualPoints: Number(question.manualPoints || 0),
      traineeAnswer: "",
      selectedAnswers: [],
      questionPenalty: false,
      penaltyPoints: 0,
      skipped: false
    }));
  });
  exam.locationRandomized = true;
  return exam;
}

function ensureExamModuleState(exam) {
  if (!exam || !Array.isArray(exam.modules)) return exam;
  if (exam.kind === "est") prepareEstExamModules(exam);
  exam.moduleIndex = Math.max(0, Math.min(Number(exam.moduleIndex || 0), Math.max(0, exam.modules.length - 1)));
  exam.questionIndex = Math.max(0, Number(exam.questionIndex || 0));
  exam.modules.forEach((module, moduleIndex) => {
    if (!module.status) {
      if (exam.status === "Vorbereitung") module.status = "Offen";
      else if (moduleIndex < exam.moduleIndex) module.status = "Abgeschlossen";
      else if (moduleIndex === exam.moduleIndex && exam.reviewMode) module.status = "Auswertung";
      else if (moduleIndex === exam.moduleIndex) module.status = "Laufend";
      else module.status = "Offen";
    }
    module.questions = module.questions || [];
    module.questions.forEach((question) => {
      if (!Array.isArray(question.selectedAnswers)) {
        question.selectedAnswers = [...(question.selectedCorrect || []), ...(question.selectedWrong || [])];
      }
      question.questionPenalty = false;
      question.penaltyPoints = 0;
      question.skipped = Boolean(question.skipped);
      question.traineeAnswer = question.traineeAnswer || "";
      question.manualPoints = Number(question.manualPoints ?? question.result?.points ?? 0);
      if (module.id === "est-location") question.maxPoints = 1;
      else if (["est-drive", "est-heli"].includes(module.id)) question.maxPoints = Math.min(10, Math.max(1, Number(question.maxPoints || 10)));
      else if (question.type === "scenario" || module.id === "est-scenario") question.maxPoints = Math.min(10, Math.max(5, Number(question.maxPoints || 10)));
      else question.maxPoints = Math.min(10, Math.max(3, Number(question.maxPoints || 3)));
    });
  });
  if (exam.kind === "est" && !exam.activeMainModuleId) {
    const firstMain = exam.modules.find((module) => !isEstLocationModule(module) && module.status !== "Abgeschlossen");
    exam.activeMainModuleId = firstMain?.id || "";
  }
  return exam;
}

function estLocationModule(exam) {
  ensureExamModuleState(exam);
  return exam.modules.find(isEstLocationModule) || null;
}

function estSideModules(exam) {
  ensureExamModuleState(exam);
  return exam.modules.filter(isEstLocationModule);
}

function estSideModulesForMain(exam, mainModule = currentManagedExamModule(exam)) {
  ensureExamModuleState(exam);
  const map = {
    "est-law": ["est-location"],
    "est-rules": ["est-drive"]
  };
  return exam.modules.filter((module) => (map[mainModule?.id] || []).includes(module.id));
}

function estMainModules(exam) {
  ensureExamModuleState(exam);
  return exam.modules.filter((module) => !isEstLocationModule(module));
}

function currentManagedExamModule(exam) {
  ensureExamModuleState(exam);
  if (exam.kind === "est" && exam.activeMainModuleId) return exam.modules.find((module) => module.id === exam.activeMainModuleId) || exam.modules[exam.moduleIndex] || null;
  return exam?.modules?.[exam.moduleIndex] || null;
}

function currentManagedExamQuestion(exam) {
  const module = currentManagedExamModule(exam);
  const questionCount = module?.questions?.length || 0;
  exam.questionIndex = Math.max(0, Math.min(Number(exam.questionIndex || 0), Math.max(0, questionCount - 1)));
  return module?.questions?.[exam.questionIndex] || null;
}

function examModuleTotal(module) {
  return (module?.questions || []).reduce((sum, question) => sum + Number(question.maxPoints || 1), 0);
}

function examModulePoints(module) {
  return (module?.questions || []).reduce((sum, question) => sum + Number(question.manualPoints ?? question.result?.points ?? 0), 0);
}

function examModulePercent(module) {
  const total = examModuleTotal(module);
  return total ? Math.round((examModulePoints(module) / total) * 100) : 0;
}

function examModuleTone(module) {
  const percent = examModulePercent(module);
  if (percent >= 75) return "good";
  if (percent >= 65) return "warn";
  return "bad";
}

function normalizeChoiceAnswers(question) {
  return Array.from(new Set([...(question.correctAnswers || []), ...(question.wrongAnswers || [])].filter(Boolean)));
}

function scoreChoiceQuestion(question) {
  if (question.skipped) return 0;
  return Math.max(0, Math.min(Number(question.maxPoints || 1), Number(question.manualPoints || 0)));
}

function scoreOptionsForQuestion(question, locationSide = false) {
  const maxPoints = Number(question.maxPoints || 1);
  const step = locationSide && maxPoints <= 1 ? 0.5 : 0.5;
  const values = [];
  for (let value = 0; value <= maxPoints + 0.001; value += step) {
    values.push(Math.round(value * 10) / 10);
  }
  return values;
}

function timedQuestionPoints(question) {
  const target = Number(question.targetSeconds || 0);
  const actual = Number(question.timeSeconds || 0);
  const max = Number(question.maxPoints || 10);
  if (!target || !actual) return Number(question.manualPoints || 0);
  if (actual <= target) return max;
  const overRatio = (actual - target) / target;
  return Math.max(0, Math.round((max * Math.max(0, 1 - overRatio)) * 10) / 10);
}

function formatSecondsInput(seconds) {
  const value = Number(seconds || 0);
  if (!value) return "";
  const minutes = Math.floor(value / 60);
  const rest = value % 60;
  return `${String(minutes).padStart(2, "0")}:${String(rest).padStart(2, "0")}`;
}

function secondsFromTimeInput(value) {
  const text = String(value || "").trim();
  if (!text) return 0;
  if (!text.includes(":")) return Number(text) || 0;
  const [minutes, seconds] = text.split(":").map((part) => Number(part) || 0);
  return minutes * 60 + seconds;
}

function renderEstExamPanel(department) {
  const candidates = state.users.filter((user) => !user.trainings?.EST);
  return `
    <div class="training-exam-layout department-overview-content est-dashboard">
      ${renderActiveTrainingExams("est", department)}
      <section class="panel training-exam-card compact-est-start est-create-panel">
        <div class="panel-header"><div><h3>EST Prüfung starten</h3><p class="muted">Prüfling auswählen und die Prüfung vorbereiten.</p></div></div>
        <div class="est-create-box">
          <label>Prüfling ohne EST ${renderExamUserPicker("estCandidateInput", "estCandidateList", candidates, "Prüfling suchen und auswählen")}</label>
          <button class="blue-btn" id="startEstExam" type="button">Prüfung vorbereiten</button>
        </div>
      </section>
      ${renderTrainingExamArchive("est", department)}
    </div>
  `;
}

function estCompletedExamItems() {
  return trainingStore().activeExams
    .filter((exam) => exam.kind === "est" && exam.status === "Abgeschlossen")
    .sort((a, b) => new Date(b.completedAt || b.archivedAt || b.startedAt || 0) - new Date(a.completedAt || a.archivedAt || a.startedAt || 0));
}

function activeExamItems(kind) {
  return trainingStore().activeExams
    .filter((exam) => exam.kind === kind && !["Vorbereitung", "Abgeschlossen", "Archiviert"].includes(exam.status))
    .sort((a, b) => new Date(b.createdAt || b.startedAt || 0) - new Date(a.createdAt || a.startedAt || 0));
}

function examArchiveItems(kind) {
  return trainingStore().activeExams
    .filter((exam) => exam.kind === kind && ["Archiviert", "Abgeschlossen"].includes(exam.status))
    .sort((a, b) => new Date(b.archivedAt || b.startedAt || 0) - new Date(a.archivedAt || a.startedAt || 0));
}

function renderActiveTrainingExams(kind, department) {
  const activeRows = activeExamItems(kind);
  const canManage = departmentActionAllowed(department, "departmentLeadership");
  return `
    <section class="panel training-active-card">
      <div class="panel-header"><div><h3>Aktive Prüfungen</h3><p class="muted">${activeRows.length} gestartete oder pausierte Prüfungen</p></div></div>
      <div class="training-active-grid">
        ${activeRows.length ? activeRows.map((exam) => renderActiveTrainingExamRow(exam, canManage)).join("") : `<p class="muted">Keine aktive Prüfung vorhanden.</p>`}
      </div>
    </section>
  `;
}

function renderCompletedTrainingExams(department) {
  const completedRows = estCompletedExamItems();
  const canManage = departmentActionAllowed(department, "departmentLeadership");
  return `
    <section class="panel training-completed-card">
      <div class="panel-header"><div><h3>Abgeschlossene EST Prüfungen</h3><p class="muted">${completedRows.length} fertig ausgewertete Prüfungen</p></div></div>
      <div class="training-archive-list">
        ${completedRows.length ? completedRows.map((exam) => renderCompletedTrainingExamRow(exam, canManage)).join("") : `<p class="muted">Noch keine abgeschlossene EST Prüfung.</p>`}
      </div>
    </section>
  `;
}

function renderActiveTrainingExamRow(exam, canManage) {
  ensureExamModuleState(exam);
  const candidate = state.users.find((user) => user.id === exam.candidateId);
  const examiner = state.users.find((user) => user.id === exam.examinerId);
  const activeModule = currentManagedExamModule(exam);
  const completedCount = exam.modules.filter((module) => module.status === "Abgeschlossen").length;
  const moduleBadges = exam.modules.map((module) => {
    const tone = module.status === "Abgeschlossen" ? examModuleTone(module) : module.id === activeModule?.id ? "active" : "";
    return `<span class="training-module-pill ${tone}"><b>${escapeHtml(module.name)}</b><small>${escapeHtml(module.status || "Offen")}${module.result ? ` · ${escapeHtml(module.result.percent)}%` : ""}</small></span>`;
  }).join("");
  return `
    <article class="training-active-exam-card">
      <div class="training-active-head">
        <div>
          <strong>${escapeHtml(candidate ? fullName(candidate) : "Unbekannter Prüfling")}</strong>
          <small>${exam.kind === "est" ? "EST Prüfung" : "Modul Prüfung"} · ${escapeHtml(exam.status)} · ${completedCount}/${exam.modules.length} Module</small>
        </div>
        <span class="status-pill ${exam.status === "Pausiert" ? "warn" : exam.status === "Modul bereit" ? "success" : ""}">${escapeHtml(exam.status)}</span>
      </div>
      <div class="training-active-meta">
        <span><b>Prüfer</b>${escapeHtml(examiner ? fullName(examiner) : "-")}</span>
        <span><b>Dauer</b><span class="exam-live-timer" data-started-at="${escapeHtml(exam.startedAt || "")}">${escapeHtml(examElapsedText(exam))}</span></span>
        <span><b>Aktuelles Modul</b>${escapeHtml(activeModule?.name || "Noch nicht gewählt")}</span>
      </div>
      <div class="training-module-pill-row">${moduleBadges}</div>
      <div class="training-active-actions">
        <button class="blue-btn training-exam-open" data-exam-id="${escapeHtml(exam.id)}" type="button">${exam.status === "Modul bereit" ? "Nächstes Modul starten" : "Öffnen"}</button>
        <button class="ghost-btn training-exam-pause" data-exam-id="${escapeHtml(exam.id)}" type="button">${exam.status === "Pausiert" ? "Fortsetzen" : "Pausieren"}</button>
        ${canManage ? `<button class="mini-icon danger training-exam-delete" data-exam-id="${escapeHtml(exam.id)}" type="button" title="Löschen">${actionIcon("delete")}</button>` : ""}
      </div>
    </article>
  `;
}

function renderCompletedTrainingExamRow(exam, canManage) {
  const candidate = state.users.find((user) => user.id === exam.candidateId);
  const result = exam.finalResult ? `${exam.finalResult.percent}% · ${exam.finalResult.points}/${exam.finalResult.total} Punkte` : "Ohne Ergebnis";
  return `
    <article class="training-archive-row completed-exam-row">
      <div>
        <strong>${escapeHtml(candidate ? fullName(candidate) : "Unbekannter Prüfling")}</strong>
        <small>EST abgeschlossen · ${formatDateTime(exam.completedAt || exam.archivedAt || exam.startedAt)}</small>
      </div>
      <span><b>Ergebnis</b>${escapeHtml(result)}</span>
      <span><b>Status</b>${exam.finalResult?.percent >= 75 ? "Bestanden" : "Nicht bestanden"}</span>
      <div class="button-row">
        <button class="blue-btn training-exam-open" data-exam-id="${escapeHtml(exam.id)}" data-readonly="true" type="button">Verlauf öffnen</button>
        <button class="ghost-btn training-exam-archive" data-exam-id="${escapeHtml(exam.id)}" type="button">Archivieren</button>
        ${canManage ? `<button class="mini-icon danger training-exam-delete" data-exam-id="${escapeHtml(exam.id)}" type="button" title="Löschen">${actionIcon("delete")}</button>` : ""}
      </div>
    </article>
  `;
}

function renderExamModuleStepper(exam) {
  ensureExamModuleState(exam);
  return `
    <div class="exam-module-stepper">
      ${exam.modules.map((module) => {
        const tone = module.status === "Abgeschlossen" ? examModuleTone(module) : "";
        return `
          <button type="button" class="exam-module-tab ${exam.activeMainModuleId === module.id ? "active" : ""} ${tone}" data-start-module-id="${escapeHtml(module.id)}">
            <b>${escapeHtml(module.name)}</b>
            <small>${escapeHtml(module.status || "Offen")}${module.result ? ` · ${escapeHtml(module.result.percent)}%` : ""}</small>
          </button>
        `;
      }).join("")}
    </div>
  `;
}

function renderCatalogQuestion(question, index, side = "main") {
  const maxPoints = Number(question.maxPoints || 1);
  const scoreValues = side === "location" ? [0, 0.5, 1] : [-1, 0, 0.5, 1, 1.5, 2].filter((value) => value <= maxPoints);
  if (side === "location" || question.type === "location") {
    return `
      <article class="exam-catalog-question location-question" data-question-id="${escapeHtml(question.id)}">
        <div class="catalog-question-grid location-score-left">
          <select class="score-select score-${String(question.manualPoints || 0).replace(".", "-")}" data-exam-score="${escapeHtml(question.id)}">${scoreValues.map((value) => `<option value="${value}" ${Number(question.manualPoints || 0) === value ? "selected" : ""}>${value}</option>`).join("")}</select>
          <div class="catalog-question-body">
            <div class="catalog-question-head"><b>${index + 1}. ${escapeHtml(question.prompt)}</b><small>Ortskunde</small></div>
            ${question.image ? `<img class="location-question-image" src="${escapeHtml(question.image)}" alt="">` : ""}
          </div>
        </div>
      </article>
    `;
  }
  if (question.type === "choice") {
    const answers = normalizeChoiceAnswers(question);
    return `
      <article class="exam-catalog-question" data-question-id="${escapeHtml(question.id)}">
        <div class="catalog-question-grid">
          <div class="catalog-question-body">
            <div class="catalog-question-head"><b>${index + 1}. ${escapeHtml(question.prompt)}</b><small>Max. ${escapeHtml(maxPoints)} Punkte</small></div>
            ${question.solution ? `<div class="inline-solution">Musterlösung: ${escapeHtml(question.solution)}</div>` : ""}
            <div class="exam-answer-list neutral">
              ${answers.map((answer) => `
                <label class="exam-check">
                  <input data-autosave-exam type="checkbox" name="answerOption_${escapeHtml(question.id)}" value="${escapeHtml(answer)}" ${question.selectedAnswers?.includes(answer) ? "checked" : ""}>
                  ${escapeHtml(answer)}
                </label>
              `).join("") || `<p class="muted">Keine Antwortmöglichkeiten hinterlegt.</p>`}
              <label class="exam-check muted-check">
                <input data-autosave-exam type="checkbox" name="questionSkipped_${escapeHtml(question.id)}" ${question.skipped ? "checked" : ""}>
                Leer gelassen / nicht beantwortet
              </label>
            </div>
            <div class="penalty-line">
              <span>Fehlerpunkte</span>
              <select data-exam-penalty="${escapeHtml(question.id)}">${[0, 1, 2, 3].map((value) => `<option value="${value}" ${Number(question.penaltyPoints || 0) === value ? "selected" : ""}>-${value}</option>`).join("")}</select>
            </div>
          </div>
          <select class="score-select score-${String(scoreChoiceQuestion(question)).replace(".", "-")}" disabled><option>${escapeHtml(scoreChoiceQuestion(question))}</option></select>
        </div>
      </article>
    `;
  }
  return `
    <article class="exam-catalog-question" data-question-id="${escapeHtml(question.id)}">
      <div class="catalog-question-grid">
        <div class="catalog-question-body">
          <div class="catalog-question-head"><b>${index + 1}. ${escapeHtml(question.prompt)}</b><small>Max. ${escapeHtml(maxPoints)} Punkte</small></div>
          ${question.solution ? `<div class="inline-solution">Musterlösung: ${escapeHtml(question.solution)}</div>` : ""}
          <label>Antwort des Prüflings<textarea data-autosave-exam data-exam-answer="${escapeHtml(question.id)}" placeholder="Antwort mitschreiben">${escapeHtml(question.traineeAnswer || "")}</textarea></label>
        </div>
        <select class="score-select score-${String(question.manualPoints || 0).replace(".", "-")}" data-exam-score="${escapeHtml(question.id)}">${scoreValues.map((value) => `<option value="${value}" ${Number(question.manualPoints || 0) === value ? "selected" : ""}>${value}</option>`).join("")}</select>
      </div>
    </article>
  `;
}

function renderExamModuleStepper(exam) {
  ensureExamModuleState(exam);
  const module = currentManagedExamModule(exam);
  return `<div class="exam-current-module-chip"><span>Aktuelles Modul</span><strong>${escapeHtml(module?.name || "-")}</strong><small>${escapeHtml(module?.status || exam.status || "-")}</small></div>`;
}

function renderCatalogQuestion(question, index, side = "main") {
  const maxPoints = Number(question.maxPoints || 1);
  const scoreValues = isLocationQuestion(question, side) ? [0, 0.5, 1] : [0, 0.5, 1, 1.5, 2].filter((value) => value <= maxPoints);
  const scoreClass = (value) => `score-select score-${String(value || 0).replace(".", "-")}`;
  const scoreBlock = (html) => `<div class="question-score-row"><span>Bewertung</span>${html}</div>`;
  if (isLocationQuestion(question, side)) {
    return `
      <article class="exam-catalog-question location-question" data-question-id="${escapeHtml(question.id)}">
        <div class="catalog-question-body"><div class="catalog-question-head"><b>${index + 1}. ${escapeHtml(question.prompt)}</b><small>Praxis</small></div>${question.image ? `<img class="location-question-image" src="${escapeHtml(question.image)}" alt="">` : ""}${question.targetSeconds ? `<label>Sollzeit<input data-exam-target="${escapeHtml(question.id)}" value="${escapeHtml(formatSecondsInput(question.targetSeconds))}" placeholder="MM:SS"></label><label>Gefahrene Zeit<input data-exam-time="${escapeHtml(question.id)}" value="${escapeHtml(formatSecondsInput(question.timeSeconds || 0))}" placeholder="MM:SS"></label>` : ""}</div>
        ${scoreBlock(`<select class="${scoreClass(question.manualPoints)}" data-exam-score="${escapeHtml(question.id)}">${scoreValues.map((value) => `<option value="${value}" ${Number(question.manualPoints || 0) === value ? "selected" : ""}>${String(value).replace(".", ",")}</option>`).join("")}</select>`)}
      </article>
    `;
  }
  if (question.type === "scenario") {
    return `
      <article class="exam-catalog-question scenario-runner-question" data-question-id="${escapeHtml(question.id)}">
        <div class="catalog-question-body">
          <div class="catalog-question-head"><b>${index + 1}. ${escapeHtml(question.prompt)}</b><small>Max. ${escapeHtml(maxPoints)} Punkte</small></div>
          <label>Antwort / Ablauf des Prüflings<textarea data-autosave-exam data-exam-answer="${escapeHtml(question.id)}" placeholder="Ablauf, Entscheidungen und Antworten mitschreiben">${escapeHtml(question.traineeAnswer || "")}</textarea></label>
          ${question.solution ? `<div class="inline-solution">Musterlösung: ${escapeHtml(question.solution)}</div>` : ""}
        </div>
        ${scoreBlock(`<select class="${scoreClass(question.manualPoints)}" data-exam-score="${escapeHtml(question.id)}">${scoreValues.map((value) => `<option value="${value}" ${Number(question.manualPoints || 0) === value ? "selected" : ""}>${String(value).replace(".", ",")}</option>`).join("")}</select>`)}
      </article>
    `;
  }
  if (question.type === "choice") {
    const answers = normalizeChoiceAnswers(question);
    return `
      <article class="exam-catalog-question compact-choice-question" data-question-id="${escapeHtml(question.id)}">
        <div class="catalog-question-body">
          <div class="catalog-question-head"><b>${index + 1}. ${escapeHtml(question.prompt)}</b><small>Max. ${escapeHtml(maxPoints)} Punkte</small></div>
          <div class="exam-answer-list neutral compact-answer-list">
            ${answers.map((answer) => `<label class="exam-check compact-answer-row"><span>${escapeHtml(answer)}</span><input data-autosave-exam type="checkbox" name="answerOption_${escapeHtml(question.id)}" value="${escapeHtml(answer)}" ${question.selectedAnswers?.includes(answer) ? "checked" : ""}></label>`).join("") || `<p class="muted">Keine Antwortmöglichkeiten hinterlegt.</p>`}
            <label class="exam-check compact-answer-row muted-check"><span>Leer gelassen / nicht beantwortet</span><input data-autosave-exam type="checkbox" name="questionSkipped_${escapeHtml(question.id)}" ${question.skipped ? "checked" : ""}></label>
          </div>
        </div>
        ${scoreBlock(`<span class="auto-score ${scoreClass(scoreChoiceQuestion(question))}">${String(scoreChoiceQuestion(question)).replace(".", ",")}</span><label class="penalty-line"><span>Fehlerpunkte</span><select data-exam-penalty="${escapeHtml(question.id)}">${[0, 1, 2, 3].map((value) => `<option value="${value}" ${Number(question.penaltyPoints || 0) === value ? "selected" : ""}>-${value}</option>`).join("")}</select></label>`)}
      </article>
    `;
  }
  return `
    <article class="exam-catalog-question" data-question-id="${escapeHtml(question.id)}">
      <div class="catalog-question-body">
        <div class="catalog-question-head"><b>${index + 1}. ${escapeHtml(question.prompt)}</b><small>Max. ${escapeHtml(maxPoints)} Punkte</small></div>
        <label>Antwort des Prüflings<textarea data-autosave-exam data-exam-answer="${escapeHtml(question.id)}" placeholder="Antwort mitschreiben">${escapeHtml(question.traineeAnswer || "")}</textarea></label>
        ${question.solution ? `<div class="inline-solution">Musterlösung: ${escapeHtml(question.solution)}</div>` : ""}
      </div>
      ${scoreBlock(`<select class="${scoreClass(question.manualPoints)}" data-exam-score="${escapeHtml(question.id)}">${scoreValues.map((value) => `<option value="${value}" ${Number(question.manualPoints || 0) === value ? "selected" : ""}>${String(value).replace(".", ",")}</option>`).join("")}</select>`)}
    </article>
  `;
}

function renderEstCatalogRunner(exam) {
  const mainModule = currentManagedExamModule(exam);
  const sideModules = estSideModulesForMain(exam, mainModule);
  const isHeliModule = mainModule?.id === "est-heli";
  const mainQuestions = isHeliModule ? (mainModule?.questions || []).filter((question) => question.stationType !== "combat") : (mainModule?.questions || []);
  const heliSideQuestions = isHeliModule ? (mainModule?.questions || []).filter((question) => question.stationType === "combat") : [];
  return `
    ${renderExamModuleStepper(exam)}
    <div class="est-runner-shell">
      <section class="est-runner-main ${mainModule?.status === "Abgeschlossen" ? examModuleTone(mainModule) : ""}">
        <div class="panel-header slim"><div><h3>${escapeHtml(mainModule?.name || "Hauptmodul")}</h3><p class="muted">Fragenkatalog links · alle Eingaben speichern automatisch.</p></div></div>
        <div class="exam-catalog-list">
          ${mainQuestions.map((question, index) => renderCatalogQuestion(question, index, isHeliModule ? "location" : "main")).join("") || `<p class="muted">Keine Fragen in diesem Modul.</p>`}
        </div>
      </section>
      <aside class="est-location-side">
        <div class="panel-header slim"><div><h3>${mainModule?.id === "est-scenario" ? "Szenario-Infos" : mainModule?.id === "est-heli" ? "Dachlandungen" : mainModule?.id === "est-rules" ? "Fahrstrecke" : "Praxis"}</h3><p class="muted">${sideModules.length || mainModule?.id === "est-scenario" || isHeliModule ? "Parallel zum aktuellen Modul." : "Dieses Modul läuft ohne parallele Praxisstrecke."}</p></div></div>
        ${mainModule?.id === "est-scenario" ? renderScenarioSidePanel(mainModule) : ""}
        ${isHeliModule ? `
          <section class="est-side-module ${mainModule.status === "Abgeschlossen" ? examModuleTone(mainModule) : ""}">
            <div class="catalog-question-head"><b>Dächer / Landepunkte</b><small>${escapeHtml(mainModule.status || "Offen")}</small></div>
            <div class="exam-catalog-list location-list">
              ${heliSideQuestions.map((question, index) => renderCatalogQuestion(question, index, "location")).join("") || `<p class="muted">Keine Dachlandungen hinterlegt.</p>`}
            </div>
          </section>
        ` : ""}
        ${!isHeliModule && sideModules.length ? sideModules.map((sideModule) => `
          <section class="est-side-module ${sideModule.status === "Abgeschlossen" ? examModuleTone(sideModule) : ""}">
            <div class="catalog-question-head"><b>${escapeHtml(sideModule.name)}</b><small>${escapeHtml(sideModule.status || "Offen")}</small></div>
            <div class="exam-catalog-list location-list">
              ${(sideModule.questions || []).map((question, index) => renderCatalogQuestion(question, index, "location")).join("") || `<p class="muted">Keine Einträge hinterlegt.</p>`}
            </div>
          </section>
        `).join("") : (!isHeliModule && mainModule?.id !== "est-scenario" ? `<p class="muted">Keine parallele Praxisstrecke in diesem Abschnitt.</p>` : "")}
      </aside>
    </div>
  `;
}

function renderScenarioSidePanel(module) {
  return `
    <section class="est-side-module scenario-side-module">
      ${(module?.questions || []).map((question, index) => `
        <article class="scenario-side-card">
          <div class="catalog-question-head"><b>${index + 1}. ${escapeHtml(question.prompt)}</b><small>Prüferbereich</small></div>
          <div class="scenario-info-box">
            <strong>Szenario Ablauf / Prüferinfos</strong>
            <p>${escapeHtml(question.scenarioInfo || "Noch keine Szenario-Infos im Leitungsbereich hinterlegt.")}</p>
          </div>
          <div class="scenario-info-box">
            <strong>Akte / Maßnahme</strong>
            <p>${escapeHtml(question.fileAction || "Noch keine Akte oder Maßnahme hinterlegt.")}</p>
          </div>
        </article>
      `).join("") || `<p class="muted">Keine Szenario-Einträge hinterlegt.</p>`}
    </section>
  `;
}

function renderNextModuleMenu(exam) {
  const modules = (exam.kind === "est" ? estMainModules(exam) : exam.modules).filter((module) => module.status !== "Abgeschlossen");
  return `
    <div class="next-module-menu hidden" id="nextModuleMenu">
      <strong>Modul auswählen</strong>
      ${modules.map((module) => `
        <button type="button" class="ghost-btn next-module-pick" data-module-id="${escapeHtml(module.id)}">
          ${escapeHtml(module.name)} <small>${escapeHtml(module.status || "Offen")}</small>
        </button>
      `).join("")}
    </div>
  `;
}

function renderModuleCatalogRunner(exam) {
  const module = currentManagedExamModule(exam);
  return `
    ${renderExamModuleStepper(exam)}
    <section class="exam-runner-card">
      <div class="panel-header slim"><div><h3>${escapeHtml(module?.name || "Modul")}</h3><p class="muted">Fragenkatalog · alle Eingaben speichern automatisch.</p></div></div>
      <div class="exam-catalog-list">
        ${(module?.questions || []).map((question, index) => renderCatalogQuestion(question, index, "main")).join("") || `<p class="muted">Keine Fragen in diesem Modul.</p>`}
      </div>
    </section>
  `;
}

function renderExamModuleStart(exam, candidate) {
  const modules = exam.kind === "est" ? estMainModules(exam).filter((module) => module.status !== "Abgeschlossen") : exam.modules.filter((module) => module.status !== "Abgeschlossen");
  return `
    ${renderExamModuleStepper(exam)}
    <section class="exam-runner-card exam-module-start-card compact-start">
      <span>Prüfung vorbereiten</span>
      <h4>${escapeHtml(candidate ? fullName(candidate) : "Unbekannter Prüfling")}</h4>
      <label>2. Prüfer optional<select id="examSetupSecondExaminer"><option value=""></option>${state.users.map((user) => `<option value="${user.id}" ${exam.secondExaminerId === user.id ? "selected" : ""}>${escapeHtml(fullName(user))}</option>`).join("")}</select></label>
      <div class="module-start-choice">
        <strong>Startmodul auswählen</strong>
        ${modules.map((module) => `<button type="button" class="ghost-btn ${exam.activeMainModuleId === module.id ? "selected" : ""}" data-start-module-id="${escapeHtml(module.id)}">${escapeHtml(module.name)}</button>`).join("") || `<p class="muted">Alle Module sind abgeschlossen.</p>`}
      </div>
      <p class="muted">Ortskunde, Helistrecke und Fahrstrecke laufen rechts parallel mit und werden separat bewertet.</p>
    </section>
  `;
}

function renderExamAnswerSummary(question) {
  if (question.type === "choice") {
    const selected = question.selectedAnswers || [];
    return `
      <small><b>Markiert:</b> ${escapeHtml(selected.join(", ") || "-")}</small>
      <small><b>Antwort / Notizen:</b> ${escapeHtml(question.traineeAnswer || "-")}</small>
      <small><b>Leer:</b> ${question.skipped ? "Ja" : "Nein"}</small>
    `;
  }
  if (question.type === "scenario") {
    return `
      <small><b>Szenario:</b> ${escapeHtml(question.scenarioInfo || "-")}</small>
      <small><b>Akte / Maßnahme:</b> ${escapeHtml(question.fileAction || "-")}</small>
      <small><b>Antwort / Ablauf:</b> ${escapeHtml(question.traineeAnswer || "-")}</small>
    `;
  }
  return `
    <small><b>Antwort Prüfling:</b> ${escapeHtml(question.traineeAnswer || "-")}</small>
    <small><b>Musterlösung:</b> ${escapeHtml(question.solution || "-")}</small>
  `;
}

function renderExamArchiveDetail(exam) {
  ensureExamModuleState(exam);
  return `
    <div class="exam-review-list archive-detail">
      ${exam.modules.map((module, moduleIndex) => {
        const points = module.result?.points ?? examModulePoints(module);
        const total = module.result?.total ?? examModuleTotal(module);
        const percent = module.result?.percent ?? examModulePercent(module);
        const tone = examModuleTone(module);
        return `
          <section class="exam-review-module archive-module-block ${tone}">
            <div class="archive-module-head">
              <h4>Modul ${moduleIndex + 1}: ${escapeHtml(module.name)}</h4>
              <span>${escapeHtml(percent)}% · ${escapeHtml(points)} / ${escapeHtml(total)} Punkte</span>
            </div>
            ${(module.questions || []).map((question, questionIndex) => `
              <div class="exam-review-row detailed archive-question-row">
                <span>
                  <b>${questionIndex + 1}. ${escapeHtml(question.prompt)}</b>
                  ${renderExamAnswerSummary(question)}
                </span>
                <strong>${escapeHtml(question.manualPoints ?? question.result?.points ?? 0)} / ${escapeHtml(question.maxPoints || 1)} Punkte</strong>
              </div>
            `).join("") || `<p class="muted">Keine Fragen gespeichert.</p>`}
          </section>
        `;
      }).join("")}
    </div>
  `;
}

function examProgressText(exam) {
  ensureExamModuleState(exam);
  const completed = exam.modules.filter((module) => module.status === "Abgeschlossen").length;
  return `${completed} von ${exam.modules.length} Modulen abgeschlossen`;
}

function finalizeExamIfComplete(exam) {
  ensureExamModuleState(exam);
  if (!exam.modules.length || !exam.modules.every((module) => module.status === "Abgeschlossen")) return false;
  const total = exam.modules.reduce((sum, module) => sum + examModuleTotal(module), 0);
  const points = exam.modules.reduce((sum, module) => sum + examModulePoints(module), 0);
  exam.finalResult = { total, points, percent: total ? Math.round((points / total) * 100) : 0 };
  exam.status = "Abgeschlossen";
  exam.completedAt = new Date().toISOString();
  return true;
}

function openTrainingExamModal(examId, readOnly = false) {
  const store = trainingStore();
  const exam = store.activeExams.find((item) => item.id === examId);
  if (!exam) return;
  ensureExamModuleState(exam);
  const candidate = state.users.find((user) => user.id === exam.candidateId);
  const archiveView = readOnly || exam.status === "Archiviert" || exam.status === "Abgeschlossen";
  const isSetup = !archiveView && (exam.status === "Vorbereitung" || exam.status === "Modul bereit" || !exam.activeMainModuleId);
  const isPaused = exam.status === "Pausiert";
  const mainModule = currentManagedExamModule(exam);
  const sideModules = estSideModules(exam);
  openModal(`
    <div class="exam-modal-head">
      <div><h3>${exam.kind === "est" ? "EST Prüfung" : "Modul Prüfung"}</h3><p class="muted">${escapeHtml(candidate ? fullName(candidate) : "Unbekannter Prüfling")} · ${escapeHtml(examProgressText(exam))}</p></div>
      ${!archiveView && !isSetup ? `<div class="exam-top-actions"><span class="pause-pill ${isPaused ? "paused" : ""}">${isPaused ? "Pausiert" : "Läuft"}</span><button class="ghost-btn" id="pauseExamRunner" type="button">${isPaused ? "Fortsetzen" : "Pausieren"}</button><button class="ghost-btn" id="saveExamRunner" type="button">Speichern</button></div>` : ""}
    </div>
    ${!archiveView && !isSetup ? `<div class="exam-runner-meta compact-meta"><label>2. Prüfer<select id="examSecondExaminer"><option value=""></option>${state.users.map((user) => `<option value="${user.id}" ${exam.secondExaminerId === user.id ? "selected" : ""}>${escapeHtml(fullName(user))}</option>`).join("")}</select></label><span><b>Dauer</b><i class="exam-live-timer" data-paused="${isPaused ? "true" : "false"}" data-started-at="${escapeHtml(exam.startedAt || "")}">${escapeHtml(examElapsedText(exam))}</i></span></div>` : ""}
    ${archiveView ? `
      ${exam.finalResult ? `<div class="exam-result-preview"><span><b>Gesamtergebnis</b>${exam.finalResult.points} von ${exam.finalResult.total} Punkten</span><span class="${exam.finalResult.percent >= 75 ? "result-pass" : "result-fail"}">${exam.finalResult.percent}% · ${exam.finalResult.percent >= 75 ? "Bestanden" : "Nicht bestanden"}</span></div>` : ""}
      ${renderExamArchiveDetail(exam)}
    ` : isSetup ? renderExamModuleStart(exam, candidate) : exam.kind === "est" ? renderEstCatalogRunner(exam) : renderModuleCatalogRunner(exam)}
    <div class="modal-actions">
      <button class="ghost-btn" id="closeExamRunner" type="button">${archiveView ? "Schließen" : "Schließen"}</button>
      ${isSetup ? `<button class="blue-btn" id="beginExamRunner" type="button">${exam.status === "Modul bereit" ? "Modul starten" : "Prüfung starten"}</button>` : ""}
      ${!archiveView && !isSetup && exam.status !== "Modul bereit" ? `<button class="blue-btn" id="finishMainModule" type="button">Modul abschließen</button>` : ""}
      ${!archiveView && !isSetup && exam.status === "Modul bereit" ? `<button class="blue-btn" id="startAnotherModule" type="button">Nächstes Modul starten</button>` : ""}
    </div>
    ${!archiveView && !isSetup ? renderNextModuleMenu(exam) : ""}
  `, (modal) => {
    modal.classList.add("exam-modal", "catalog-exam-modal");
    if (isSetup) modal.classList.add("setup-exam-modal");
    if (isSetup && exam.status === "Vorbereitung" && !exam.startedAt) {
      modalRoot.dataset.discardTrainingExamId = exam.id;
    } else {
      delete modalRoot.dataset.discardTrainingExamId;
    }
    const persist = (message = "") => {
      ensureExamModuleState(exam);
      saveActiveTrainingExam(exam);
      if (message) showNotify(message, "success");
    };
    const saveQuestionFromCard = (card) => {
      const questionId = card?.dataset.questionId;
      const module = exam.modules.find((item) => item.questions.some((question) => question.id === questionId));
      const question = module?.questions.find((item) => item.id === questionId);
      if (!question) return;
      if (question.type === "choice" || question.type === "scenario") {
        question.skipped = Boolean(card.querySelector(`[name='questionSkipped_${CSS.escape(question.id)}']`)?.checked);
        question.penaltyPoints = 0;
        question.questionPenalty = false;
        question.selectedAnswers = question.skipped ? [] : Array.from(card.querySelectorAll(`[name='answerOption_${CSS.escape(question.id)}']:checked`)).map((input) => input.value);
        const answer = card.querySelector(`[data-exam-answer='${CSS.escape(question.id)}']`);
        const score = card.querySelector(`[data-exam-score='${CSS.escape(question.id)}']`);
        if (answer) question.traineeAnswer = answer.value || "";
        if (score) question.manualPoints = Number(score.value);
        question.result = { points: question.manualPoints };
      } else {
        const answer = card.querySelector(`[data-exam-answer='${CSS.escape(question.id)}']`);
        const score = card.querySelector(`[data-exam-score='${CSS.escape(question.id)}']`);
        const time = card.querySelector(`[data-exam-time='${CSS.escape(question.id)}']`);
        const target = card.querySelector(`[data-exam-target='${CSS.escape(question.id)}']`);
        if (answer) question.traineeAnswer = answer.value || "";
        if (time) question.timeSeconds = secondsFromTimeInput(time.value);
        if (target) question.targetSeconds = secondsFromTimeInput(target.value);
        if (Number(question.maxPoints || 1) > 1 || question.targetSeconds) {
          question.manualPoints = timedQuestionPoints(question);
        } else if (score) {
          question.manualPoints = Number(score.value);
        }
        question.result = { points: question.manualPoints };
      }
      if (!exam.startedAt) exam.startedAt = new Date().toISOString();
      if (module.status === "Offen") module.status = "Laufend";
      if (!["Abgeschlossen", "Archiviert", "Pausiert"].includes(exam.status)) exam.status = "Laufend";
      saveActiveTrainingExam(exam);
    };
    const saveAll = () => {
      modal.querySelectorAll(".exam-catalog-question").forEach(saveQuestionFromCard);
      saveActiveTrainingExam(exam);
    };
    modal.querySelectorAll("[data-start-module-id]").forEach((button) => button.addEventListener("click", () => {
      const module = exam.modules.find((item) => item.id === button.dataset.startModuleId);
      if (!module || isEstLocationModule(module)) return;
      button.closest(".module-start-choice, .exam-module-stepper")?.querySelectorAll("[data-start-module-id]").forEach((item) => item.classList.toggle("selected", item === button));
      exam.activeMainModuleId = module.id;
      exam.moduleIndex = exam.modules.findIndex((item) => item.id === module.id);
      persist();
      if (!isSetup) openTrainingExamModal(exam.id);
    }));
    modal.querySelector("#examSecondExaminer")?.addEventListener("change", (event) => {
      exam.secondExaminerId = event.target.value;
      persist("2. Prüfer gespeichert.");
    });
    modal.querySelector("#beginExamRunner")?.addEventListener("click", () => {
      const selected = modal.querySelector(".module-start-choice .selected")?.dataset.startModuleId || exam.activeMainModuleId || modal.querySelector("[data-start-module-id]")?.dataset.startModuleId;
      if (!selected) {
        showNotify("Bitte zuerst ein Startmodul auswählen.", "error");
        return;
      }
      delete modalRoot.dataset.discardTrainingExamId;
      const module = exam.modules.find((item) => item.id === selected);
      exam.secondExaminerId = modal.querySelector("#examSetupSecondExaminer")?.value || "";
      exam.activeMainModuleId = selected;
      exam.moduleIndex = exam.modules.findIndex((item) => item.id === selected);
      exam.status = "Laufend";
      exam.startedAt = exam.startedAt || new Date().toISOString();
      exam.modules.forEach((item) => {
        if (item.id !== selected && item.status !== "Abgeschlossen") item.status = "Offen";
      });
      if (module) {
        module.status = "Laufend";
        module.startedAt = module.startedAt || new Date().toISOString();
      }
      persist("Prüfung gestartet.");
      openTrainingExamModal(exam.id);
      renderDepartmentPage(departmentByPage(state.page));
    });
    modal.querySelectorAll("[data-autosave-exam], [data-exam-score]").forEach((input) => {
      input.addEventListener(input.tagName === "TEXTAREA" ? "input" : "change", () => {
        if (input.matches("[data-exam-score]")) {
          input.className = `score-select score-${String(input.value || 0).replace(".", "-")}`;
        }
        saveQuestionFromCard(input.closest(".exam-catalog-question"));
      });
    });
    modal.querySelector("#saveExamRunner")?.addEventListener("click", () => {
      saveAll();
      showNotify("Prüfung gespeichert.", "success");
    });
    modal.querySelector("#pauseExamRunner")?.addEventListener("click", () => {
      saveAll();
      if (exam.status === "Pausiert") {
        exam.pausedTotalMs = Number(exam.pausedTotalMs || 0) + (Date.now() - new Date(exam.pausedAt || Date.now()).getTime());
        exam.pausedAt = "";
        exam.status = "Laufend";
        if (mainModule?.status !== "Abgeschlossen") mainModule.status = "Laufend";
        showNotify("Prüfung fortgesetzt.", "success");
      } else {
        exam.status = "Pausiert";
        exam.pausedAt = new Date().toISOString();
        showNotify("Prüfung pausiert.", "success");
      }
      saveActiveTrainingExam(exam);
      renderDepartmentPage(departmentByPage(state.page));
      openTrainingExamModal(exam.id);
    });
    modal.querySelector("#finishMainModule")?.addEventListener("click", () => {
      saveAll();
      if (mainModule) {
        mainModule.status = "Abgeschlossen";
        mainModule.completedAt = new Date().toISOString();
        mainModule.result = { total: examModuleTotal(mainModule), points: examModulePoints(mainModule), percent: examModulePercent(mainModule) };
      }
      const remainingMain = estMainModules(exam).some((module) => module.status !== "Abgeschlossen");
      if (exam.kind === "est") {
        estSideModulesForMain(exam, mainModule).forEach((sideModule) => {
          sideModule.status = "Abgeschlossen";
          sideModule.completedAt = sideModule.completedAt || new Date().toISOString();
          sideModule.result = { total: examModuleTotal(sideModule), points: examModulePoints(sideModule), percent: examModulePercent(sideModule) };
        });
      }
      if (!remainingMain && exam.kind === "est") {
        estSideModules(exam).forEach((sideModule) => {
          if (sideModule.status === "Abgeschlossen") return;
          sideModule.status = "Abgeschlossen";
          sideModule.completedAt = sideModule.completedAt || new Date().toISOString();
          sideModule.result = { total: examModuleTotal(sideModule), points: examModulePoints(sideModule), percent: examModulePercent(sideModule) };
        });
      }
      const completedAll = finalizeExamIfComplete(exam);
      if (!completedAll) {
        const next = estMainModules(exam).find((module) => module.status !== "Abgeschlossen");
        if (next) {
          exam.activeMainModuleId = next.id;
          exam.moduleIndex = exam.modules.findIndex((module) => module.id === next.id);
          exam.modules.forEach((module) => {
            if (module.status !== "Abgeschlossen") module.status = module.id === next.id ? "Offen" : "Offen";
          });
        }
        exam.status = "Modul bereit";
      }
      saveActiveTrainingExam(exam);
      renderDepartmentPage(departmentByPage(state.page));
      showNotify(completedAll ? "EST Prüfung vollständig abgeschlossen." : "Modul abgeschlossen.", "success");
      if (completedAll) {
        openTrainingExamModal(exam.id, true);
      } else {
        closeModal();
      }
    });
    modal.querySelector("#startAnotherModule")?.addEventListener("click", () => {
      modal.querySelector("#nextModuleMenu")?.classList.toggle("hidden");
    });
    modal.querySelectorAll(".next-module-pick").forEach((button) => button.addEventListener("click", () => {
      saveAll();
      if (mainModule && mainModule.status !== "Abgeschlossen") {
        mainModule.status = "Abgeschlossen";
        mainModule.completedAt = new Date().toISOString();
        mainModule.result = { total: examModuleTotal(mainModule), points: examModulePoints(mainModule), percent: examModulePercent(mainModule) };
        estSideModulesForMain(exam, mainModule).forEach((sideModule) => {
          sideModule.status = "Abgeschlossen";
          sideModule.completedAt = sideModule.completedAt || new Date().toISOString();
          sideModule.result = { total: examModuleTotal(sideModule), points: examModulePoints(sideModule), percent: examModulePercent(sideModule) };
        });
      }
      const completedAll = finalizeExamIfComplete(exam);
      if (!completedAll) {
        const next = exam.modules.find((module) => module.id === button.dataset.moduleId);
        if (next) {
          exam.activeMainModuleId = next.id;
          exam.moduleIndex = exam.modules.findIndex((module) => module.id === next.id);
          next.status = "Laufend";
          exam.status = "Laufend";
        } else {
          exam.status = "Modul bereit";
        }
      }
      saveActiveTrainingExam(exam);
      renderDepartmentPage(departmentByPage(state.page));
      showNotify(completedAll ? "EST Prüfung vollständig abgeschlossen." : "Modul abgeschlossen.", "success");
      openTrainingExamModal(exam.id, completedAll);
    }));
    modal.querySelector("#closeExamRunner")?.addEventListener("click", () => {
      if (!archiveView && !isSetup) saveAll();
      closeModal();
      renderDepartmentPage(departmentByPage(state.page));
      if (!archiveView && !isSetup) showNotify("Prüfung automatisch gespeichert.", "success");
    });
  });
}

function defaultInformationDocs() {
  return ["Interne Vorschriften", "Kleiderordnung", "Fahrzeugregelung"].map((title) => ({
    id: makeTrainingId("infodoc"),
    title,
    body: `## ${title}\n\nText kann hier gepflegt werden.`,
    updatedAt: new Date().toISOString(),
    updatedBy: fullName(state.currentUser || {})
  }));
}

function informationDocs() {
  const docs = state.settings.informationDocs || [];
  return docs.length ? docs : defaultInformationDocs();
}

function unreadInformationChanges() {
  const myId = state.currentUser?.id || "";
  return (state.settings.informationDocChanges || []).filter((change) => !(change.acknowledgedBy || []).includes(myId));
}

function renderInformation() {
  const links = state.settings.informationLinks || [];
  const docs = informationDocs();
  const changes = state.settings.informationDocChanges || [];
  const unread = unreadInformationChanges();
  content.innerHTML = `
    <section class="department-info-view information-admin-view modern-info-view">
      ${unread.length ? `<div class="info-change-alert">Es gibt ${unread.length} Änderung${unread.length === 1 ? "" : "en"} bei internen Informationen.</div>` : ""}
      <div class="info-box full information-card redirects-card">
        <div class="department-modal-heading">
          <h4>Link Weiterleitungen</h4>
          ${canAccess("actions", "manageInformation", "Direktion") ? `<button class="blue-btn" id="addInformationLink">${iconSvg("Plus")} Hinzufügen</button>` : ""}
        </div>
        <div class="link-card-grid">${links.map((link) => `
          <article class="small-link-card">
            <strong>${escapeHtml(link.title)}</strong>
            <span class="link-label">Link:</span>
            <a href="${escapeHtml(link.url)}" target="_blank" rel="noreferrer">${escapeHtml(link.url)}</a>
            ${canAccess("actions", "manageInformation", "Direktion") ? `<span class="button-row"><button class="blue-btn compact-action edit-info-link" data-id="${link.id}" title="Bearbeiten">${actionIcon("edit")} Bearbeiten</button><button class="mini-icon danger delete-info-link" data-id="${link.id}" title="Löschen">${actionIcon("delete")}</button></span>` : ""}
          </article>
        `).join("") || `<p class="muted">Noch keine Weiterleitungen.</p>`}</div>
      </div>
      <div class="info-box full information-card internal-doc-card">
        <div class="department-modal-heading">
          <h4>Interne Weiterleitungen</h4>
          ${canAccess("actions", "manageInformation", "Direktion") ? `<button class="blue-btn" id="addInformationDoc">${iconSvg("Plus")} Dokument</button>` : ""}
        </div>
        <div class="internal-doc-grid">${docs.map((doc) => `
          <button class="internal-doc-tile" data-doc-id="${escapeHtml(doc.id)}">
            <strong>${escapeHtml(doc.title)}</strong>
            <small>Zuletzt geändert: ${formatDateTime(doc.updatedAt)}</small>
          </button>
        `).join("")}</div>
      </div>
      <div class="info-box full information-card">
        <div class="department-modal-heading"><h4>Changelog</h4></div>
        <div class="info-change-list">${changes.slice(0, 12).map((change) => `
          <article class="info-change-row">
            <strong>${escapeHtml(change.title || "Dokument")}</strong>
            <small>${escapeHtml(change.author || "-")} · ${formatDateTime(change.createdAt)}</small>
            <div><del>${escapeHtml(change.before || "-")}</del><ins>${escapeHtml(change.after || "-")}</ins></div>
          </article>
        `).join("") || `<p class="muted">Noch keine Änderungen.</p>`}</div>
      </div>
    </section>
  `;
  $("#addInformationLink")?.addEventListener("click", () => openInformationLinkModal());
  $("#addInformationDoc")?.addEventListener("click", () => openInformationDocModal());
  document.querySelectorAll(".edit-info-link").forEach((button) => button.addEventListener("click", () => openInformationLinkModal(links.find((item) => item.id === button.dataset.id))));
  document.querySelectorAll(".delete-info-link").forEach((button) => button.addEventListener("click", () => deleteInformationItem("informationLinks", button.dataset.id)));
  document.querySelectorAll(".internal-doc-tile").forEach((button) => button.addEventListener("click", () => openInformationDocView(button.dataset.docId)));
}

function formatInformationDocText(text) {
  return formatDepartmentText(text || "");
}

function openInformationDocView(docId) {
  const docs = informationDocs();
  const doc = docs.find((item) => item.id === docId);
  if (!doc) return;
  const canEdit = canAccess("actions", "manageInformation", "Direktion");
  openModal(`
    <div class="paper-doc-modal">
      <div class="paper-doc-head">
        <h3>${escapeHtml(doc.title)}</h3>
        <input id="docSearchInput" placeholder="Im Dokument suchen">
        ${canEdit ? `<button class="blue-btn" id="editInformationDoc">${actionIcon("edit")} Bearbeiten</button>` : ""}
      </div>
      <article class="paper-doc-page" id="paperDocPage">${formatInformationDocText(doc.body)}</article>
    </div>
  `, (modal) => {
    modal.classList.add("wide-doc-modal");
    modal.querySelector("#docSearchInput")?.addEventListener("input", (event) => {
      const term = event.target.value.trim().toLowerCase();
      modal.querySelectorAll("#paperDocPage *").forEach((node) => node.classList.toggle("search-hit", term && node.textContent.toLowerCase().includes(term)));
    });
    modal.querySelector("#editInformationDoc")?.addEventListener("click", () => openInformationDocModal(doc));
  });
}

function openInformationDocModal(doc = null) {
  openModal(`
    <h3>${doc ? "Internes Dokument bearbeiten" : "Internes Dokument erstellen"}</h3>
    <label>Titel<input id="informationDocTitle" value="${escapeHtml(doc?.title || "")}"></label>
    <div class="format-toolbar"><button type="button" data-format="## ">Überschrift</button><button type="button" data-format="**fett**">Fett</button><button type="button" data-format="<span style='color:#75ffad'>Grün</span>">Grün</button><button type="button" data-format="<span style='color:#ff9ca0'>Rot</span>">Rot</button></div>
    <label>Text<textarea id="informationDocBody" rows="14">${escapeHtml(doc?.body || "")}</textarea></label>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions"><button class="ghost-btn" data-close>Abbrechen</button><button class="blue-btn" id="saveInformationDoc">Speichern</button></div>
  `, (modal) => {
    modal.querySelectorAll("[data-format]").forEach((button) => button.addEventListener("click", () => {
      const area = modal.querySelector("#informationDocBody");
      const value = button.dataset.format;
      area.setRangeText(value, area.selectionStart, area.selectionEnd, "end");
      area.focus();
    }));
    modal.querySelector("#saveInformationDoc").addEventListener("click", async () => {
      try {
        const title = modal.querySelector("#informationDocTitle").value.trim();
        const body = modal.querySelector("#informationDocBody").value;
        if (!title) throw new Error("Titel ist erforderlich.");
        const before = doc?.body || "";
        const nextDoc = { id: doc?.id || makeTrainingId("infodoc"), title, body, updatedAt: new Date().toISOString(), updatedBy: fullName(state.currentUser) };
        const docs = upsertById(informationDocs(), nextDoc);
        const changes = [{ id: makeTrainingId("docchange"), docId: nextDoc.id, title, before, after: body, action: doc ? "geändert" : "erstellt", createdAt: new Date().toISOString(), author: fullName(state.currentUser), acknowledgedBy: [] }, ...(state.settings.informationDocChanges || [])];
        await saveInformationPatch({ informationDocs: docs, informationDocChanges: changes });
        openInformationDocView(nextDoc.id);
      } catch (error) {
        modal.querySelector("#modalError").textContent = error.message;
      }
    });
  });
}

function formatInformationDocText(text = "", searchTerm = "") {
  let escaped = escapeHtml(text || "Noch kein Inhalt vorhanden.");
  const term = String(searchTerm || "").trim();
  if (term) {
    const safeTerm = escapeHtml(term).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    escaped = escaped.replace(new RegExp(`(${safeTerm})`, "gi"), `<mark class="doc-search-mark">$1</mark>`);
  }
  return escaped
    .replace(/\*\*(.*?)\*\*/g, "<strong>$1</strong>")
    .replace(/^### (.*)$/gm, "<h4>$1</h4>")
    .replace(/^## (.*)$/gm, "<h3>$1</h3>")
    .replace(/\n/g, "<br>");
}

function informationDocChangesFor(docId) {
  return (state.settings.informationDocChanges || []).filter((change) => change.docId === docId);
}

async function saveInformationDocDirect(doc, title, body, closeAfter = false) {
  if (!title) throw new Error("Titel ist erforderlich.");
  const before = doc?.body || "";
  const nextDoc = { id: doc?.id || makeTrainingId("infodoc"), title, body, updatedAt: new Date().toISOString(), updatedBy: fullName(state.currentUser) };
  const changes = before === body ? (state.settings.informationDocChanges || []) : [{
    id: makeTrainingId("docchange"),
    docId: nextDoc.id,
    title,
    before,
    after: body,
    action: doc?.id ? "geändert" : "erstellt",
    createdAt: new Date().toISOString(),
    author: fullName(state.currentUser),
    acknowledgedBy: []
  }, ...(state.settings.informationDocChanges || [])];
  await saveInformationPatch({ informationDocs: upsertById(informationDocs(), nextDoc), informationDocChanges: changes });
  if (closeAfter) {
    closeModal();
    renderInformation();
  } else {
    openInformationDocView(nextDoc.id);
  }
}

function openInformationDocCloseConfirm(doc, title, before, after) {
  openModal(`
    <div class="doc-compare-head">
      <span class="doc-compare-kicker">Vorschrift speichern</span>
      <h3>Änderung prüfen</h3>
      <p>Vergleiche die bisherige und die neue Fassung, bevor du die Änderung ins Dienstblatt übernimmst.</p>
    </div>
    <div class="doc-compare-grid">
      <section class="doc-compare-panel before">
        <header><span>Vorher</span><small>Aktuell gespeichert</small></header>
        <article class="doc-save-preview before">${formatInformationDocText(before || "")}</article>
      </section>
      <section class="doc-compare-panel after">
        <header><span>Nachher</span><small>Neue Fassung</small></header>
        <article class="doc-save-preview after">${formatInformationDocText(after || "")}</article>
      </section>
    </div>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions">
      <button class="ghost-btn" id="backToDocEdit">Zurück</button>
      <button class="ghost-btn" id="discardDocChanges">Nicht speichern</button>
      <button class="blue-btn" id="confirmSaveDocChanges">Speichern</button>
    </div>
  `, (confirmModal) => {
    confirmModal.classList.add("doc-compare-modal");
    confirmModal.querySelector("#backToDocEdit").addEventListener("click", () => openInformationDocView(doc.id, { title, body: after }));
    confirmModal.querySelector("#discardDocChanges").addEventListener("click", closeModal);
    confirmModal.querySelector("#confirmSaveDocChanges").addEventListener("click", async () => {
      try {
        await saveInformationDocDirect(doc, title, after, true);
      } catch (error) {
        confirmModal.querySelector("#modalError").textContent = error.message;
      }
    });
  });
}

function openInformationDocView(docId, draft = null) {
  const existing = informationDocs().find((item) => item.id === docId);
  const doc = existing || { id: docId || makeTrainingId("infodoc"), title: draft?.title || "Neue Vorschrift", body: draft?.body || "", updatedAt: new Date().toISOString() };
  const canEdit = canAccess("actions", "manageInformation", "Direktion");
  const changes = informationDocChangesFor(doc.id);
  openModal(`
    <div class="paper-doc-modal">
      <div class="paper-doc-head">
        ${canEdit ? `<input id="paperDocTitle" value="${escapeHtml(draft?.title ?? doc.title)}">` : `<h3>${escapeHtml(doc.title)}</h3>`}
        <input id="docSearchInput" placeholder="Im Dokument suchen">
      </div>
      ${canEdit ? `
        <div class="format-toolbar"><button type="button" data-format="## ">Überschrift</button><button type="button" data-format="**fett**">Fett</button><button type="button" data-format="<span style='color:#75ffad'>Grün</span>">Grün</button><button type="button" data-format="<span style='color:#ff9ca0'>Rot</span>">Rot</button></div>
        <textarea class="paper-doc-page paper-doc-editor" id="paperDocEditor">${escapeHtml(draft?.body ?? doc.body ?? "")}</textarea>
      ` : `<article class="paper-doc-page" id="paperDocPage">${formatInformationDocText(doc.body)}</article>`}
      <details class="doc-change-details">
        <summary>Changelog (${changes.length})</summary>
        <div class="info-change-list">${changes.map((change) => `
          <article class="info-change-row">
            <strong>${escapeHtml(change.action || "geändert")}</strong>
            <small>${escapeHtml(change.author || "-")} · ${formatDateTime(change.createdAt)}</small>
            <div><del>${escapeHtml(change.before || "-")}</del><ins>${escapeHtml(change.after || "-")}</ins></div>
          </article>
        `).join("") || `<p class="muted">Noch keine Änderungen.</p>`}</div>
      </details>
    </div>
  `, (modal) => {
    modal.classList.add("wide-doc-modal");
    const initial = draft?.body ?? doc.body ?? "";
    const x = modal.querySelector(".modal-x");
    if (x && canEdit) {
      const clone = x.cloneNode(true);
      x.replaceWith(clone);
      clone.addEventListener("click", () => {
        const title = modal.querySelector("#paperDocTitle")?.value.trim() || doc.title;
        const current = modal.querySelector("#paperDocEditor")?.value || "";
        if (current !== initial || title !== doc.title) openInformationDocCloseConfirm(doc, title, initial, current);
        else closeModal();
      });
    }
    modal.querySelector("#docSearchInput")?.addEventListener("input", (event) => {
      const term = event.target.value.trim();
      const page = modal.querySelector("#paperDocPage");
      if (page) page.innerHTML = formatInformationDocText(doc.body, term);
      const editor = modal.querySelector("#paperDocEditor");
      if (editor && term) {
        const index = editor.value.toLowerCase().indexOf(term.toLowerCase());
        if (index >= 0) {
          editor.focus();
          editor.setSelectionRange(index, index + term.length);
        }
      }
    });
    modal.querySelectorAll("[data-format]").forEach((button) => button.addEventListener("click", () => {
      const area = modal.querySelector("#paperDocEditor");
      area.setRangeText(button.dataset.format, area.selectionStart, area.selectionEnd, "end");
      area.focus();
    }));
  });
}

function renderInformation() {
  const links = state.settings.informationLinks || [];
  const docs = informationDocs();
  const permits = state.settings.informationPermits || [];
  const factions = state.settings.informationFactions || [];
  const unread = unreadInformationChanges();
  content.innerHTML = `
    <section class="department-info-view information-admin-view modern-info-view">
      ${unread.length ? `<div class="info-change-alert">Es gibt ${unread.length} Änderung${unread.length === 1 ? "" : "en"} bei Vorschriften.</div>` : ""}
      <div class="info-box full information-card internal-doc-card">
        <div class="department-modal-heading">
          <h4>Vorschriften</h4>
          ${canAccess("actions", "manageInformation", "Direktion") ? `<button class="blue-btn" id="addInformationDoc">${iconSvg("Plus")} Dokument</button>` : ""}
        </div>
        <div class="internal-doc-grid">${docs.map((doc) => `
          <button class="internal-doc-tile" data-doc-id="${escapeHtml(doc.id)}"><strong>${escapeHtml(doc.title)}</strong><small>Zuletzt geändert: ${formatDateTime(doc.updatedAt)}</small></button>
        `).join("")}</div>
      </div>
      <div class="info-box full information-card redirects-card">
        <div class="department-modal-heading"><h4>Link Weiterleitungen</h4>${canAccess("actions", "manageInformation", "Direktion") ? `<button class="blue-btn" id="addInformationLink">${iconSvg("Plus")} Hinzufügen</button>` : ""}</div>
        <div class="link-card-grid">${links.map((link) => `<article class="small-link-card"><strong>${escapeHtml(link.title)}</strong><span class="link-label">Link:</span><a href="${escapeHtml(link.url)}" target="_blank" rel="noreferrer">${escapeHtml(link.url)}</a>${canAccess("actions", "manageInformation", "Direktion") ? `<span class="button-row"><button class="blue-btn compact-action edit-info-link" data-id="${link.id}" title="Bearbeiten">${actionIcon("edit")} Bearbeiten</button><button class="mini-icon danger delete-info-link" data-id="${link.id}" title="Löschen">${actionIcon("delete")}</button></span>` : ""}</article>`).join("") || `<p class="muted">Noch keine Weiterleitungen.</p>`}</div>
      </div>
      <div class="info-box full information-card"><div class="department-modal-heading"><h4>Rechte Definition</h4>${canAccess("actions", "manageInformation", "Direktion") ? `<button class="blue-btn" id="editInformationRights">${actionIcon("edit")} Bearbeiten</button>` : ""}</div><div class="rich-text-view">${formatDepartmentText(state.settings.informationRightsText)}</div></div>
      <div class="info-box full information-card"><div class="department-modal-heading"><h4>Sondergenehmigungen</h4>${canAccess("actions", "manageInformation", "Direktion") ? `<button class="blue-btn" id="addInformationPermit">${iconSvg("Plus")} Hinzufügen</button>` : ""}</div><div class="table-wrap compact-table"><table><thead><tr><th>Vor- und Nachname</th><th>Beschreibung</th><th>Gültig Bis</th><th>Aktionen</th></tr></thead><tbody>${permits.map((permit) => `<tr><td>${escapeHtml(permit.name)}</td><td>${escapeHtml(permit.description)}</td><td>${formatDate(permit.validUntil)}</td><td>${canAccess("actions", "manageInformation", "Direktion") ? `<button class="mini-icon edit-info-permit" data-id="${permit.id}">${actionIcon("edit")}</button><button class="mini-icon danger delete-info-permit" data-id="${permit.id}">${actionIcon("delete")}</button>` : ""}</td></tr>`).join("") || `<tr><td colspan="4" class="muted">Keine Sondergenehmigungen.</td></tr>`}</tbody></table></div></div>
      <div class="info-box full information-card"><div class="department-modal-heading"><h4>Fraktionen</h4>${canAccess("actions", "manageInformation", "Direktion") ? `<button class="blue-btn" id="addInformationFaction">${iconSvg("Plus")} Hinzufügen</button>` : ""}</div><div class="table-wrap compact-table"><table><thead><tr><th>Organisation</th><th>Status</th><th>Aktionen</th></tr></thead><tbody>${factions.map((faction) => `<tr><td>${escapeHtml(faction.organization)}</td><td><span class="status-label">${renderStatusDot(faction.status)}</span></td><td>${canAccess("actions", "manageInformation", "Direktion") ? `<button class="mini-icon edit-info-faction" data-id="${faction.id}">${actionIcon("edit")}</button><button class="mini-icon danger delete-info-faction" data-id="${faction.id}">${actionIcon("delete")}</button>` : ""}</td></tr>`).join("") || `<tr><td colspan="3" class="muted">Keine Fraktionen.</td></tr>`}</tbody></table></div></div>
    </section>
    <section class="panel"><div class="panel-header"><h3><span class="section-icon">${iconSvg("Informationen")}</span>Informationen</h3>${canAccess("actions", "manageInformation", "Direktion") ? `<button class="blue-btn" id="editInformation">${actionIcon("edit")} Bearbeiten</button>` : ""}</div><div class="info-box"><strong>Bewerbungsstatus</strong><p><span class="application-pill ${state.settings.applicationStatus === "Offen" ? "open" : "closed"}">${escapeHtml(state.settings.applicationStatus)}</span></p></div><div class="info-box"><strong>Beschreibung</strong><p>${escapeHtml(state.settings.informationText)}</p></div></section>
  `;
  $("#editInformation")?.addEventListener("click", openInformationEditModal);
  $("#editInformationRights")?.addEventListener("click", openInformationRightsModal);
  $("#addInformationLink")?.addEventListener("click", () => openInformationLinkModal());
  $("#addInformationDoc")?.addEventListener("click", () => openInformationDocView(makeTrainingId("infodoc")));
  $("#addInformationPermit")?.addEventListener("click", () => openInformationPermitModal());
  $("#addInformationFaction")?.addEventListener("click", () => openInformationFactionModal());
  document.querySelectorAll(".edit-info-link").forEach((button) => button.addEventListener("click", () => openInformationLinkModal(links.find((item) => item.id === button.dataset.id))));
  document.querySelectorAll(".delete-info-link").forEach((button) => button.addEventListener("click", () => deleteInformationItem("informationLinks", button.dataset.id)));
  document.querySelectorAll(".edit-info-permit").forEach((button) => button.addEventListener("click", () => openInformationPermitModal(permits.find((item) => item.id === button.dataset.id))));
  document.querySelectorAll(".delete-info-permit").forEach((button) => button.addEventListener("click", () => deleteInformationItem("informationPermits", button.dataset.id)));
  document.querySelectorAll(".edit-info-faction").forEach((button) => button.addEventListener("click", () => openInformationFactionModal(factions.find((item) => item.id === button.dataset.id))));
  document.querySelectorAll(".delete-info-faction").forEach((button) => button.addEventListener("click", () => deleteInformationItem("informationFactions", button.dataset.id)));
  document.querySelectorAll(".internal-doc-tile").forEach((button) => button.addEventListener("click", () => openInformationDocView(button.dataset.docId)));
}

function unreadMailboxItems() {
  const myId = state.currentUser?.id || "";
  return (state.settings.informationDocChanges || []).filter((change) => !(change.acknowledgedBy || []).includes(myId));
}

function mailboxUnreadCount() {
  return unreadMailboxItems().length;
}

function markInformationChangeRead(changeId) {
  const myId = state.currentUser?.id || "";
  const changes = (state.settings.informationDocChanges || []).map((change) => change.id === changeId
    ? { ...change, acknowledgedBy: Array.from(new Set([...(change.acknowledgedBy || []), myId])) }
    : change);
  return saveInformationPatch({ informationDocChanges: changes });
}

function renderPage() {
  if (state.page === "Dienstblatt") return renderDienstblatt();
  if (state.page === "Mitglieder") return renderMembers();
  if (state.page === "Mitgliederfluktation") return renderFluctuation();
  if (state.page === "Beschlagnahmung") return renderSeizures();
  if (state.page === "Kalender") return renderCalendar();
  if (state.page === "Informationen") return renderInformation();
  if (state.page === "Postfach") return renderPostfach();
  if (state.page === "Direktion") return renderDirektion();
  if (state.page === "IT") return renderIT();
  if (state.page === "Abteilungen") return renderDepartmentsOverview();
  if (isDepartmentPage(state.page)) return renderDepartmentPage(departmentByPage(state.page));
  if (state.page === "Profil") return renderProfile();
  return renderTemplate(state.page);
}

function renderPostfach() {
  const unread = unreadMailboxItems();
  const allChanges = state.settings.informationDocChanges || [];
  const rows = allChanges.length ? allChanges : [];
  content.innerHTML = `
    <section class="panel mailbox-page">
      <div class="panel-header"><div><h3>Postfach</h3><p class="muted">${unread.length} ungelesene Nachricht${unread.length === 1 ? "" : "en"}</p></div></div>
      <div class="mailbox-list">
        ${rows.map((change) => {
          const read = !unread.some((item) => item.id === change.id);
          return `
            <article class="mailbox-row ${read ? "read" : "unread"}">
              <div><strong>Änderung bei ${escapeHtml(change.title || "Vorschrift")}</strong><small>${escapeHtml(change.author || "-")} · ${formatDateTime(change.createdAt)}</small></div>
              <p>${escapeHtml(change.action || "geändert")} in einem Vorschriften-Dokument.</p>
              <div class="button-row">
                <button class="blue-btn open-mail-doc" data-doc-id="${escapeHtml(change.docId)}" data-change-id="${escapeHtml(change.id)}">Öffnen</button>
                ${read ? "" : `<button class="ghost-btn mark-mail-read" data-change-id="${escapeHtml(change.id)}">Als gelesen markieren</button>`}
              </div>
            </article>
          `;
        }).join("") || `<p class="muted">Keine Nachrichten vorhanden.</p>`}
      </div>
    </section>
  `;
  document.querySelectorAll(".open-mail-doc").forEach((button) => button.addEventListener("click", async () => {
    await markInformationChangeRead(button.dataset.changeId);
    renderNavigation();
    openInformationDocView(button.dataset.docId, null, button.dataset.changeId);
  }));
  document.querySelectorAll(".mark-mail-read").forEach((button) => button.addEventListener("click", async () => {
    await markInformationChangeRead(button.dataset.changeId);
    await bootstrap();
  }));
}

function renderPostfach() {
  const unread = unreadMailboxItems();
  const rows = state.settings.informationDocChanges || [];
  content.innerHTML = `
    <section class="panel mailbox-page">
      <div class="panel-header"><div><h3>Postfach</h3><p class="muted">${unread.length} ungelesene Nachricht${unread.length === 1 ? "" : "en"}</p></div></div>
      <div class="mailbox-list">
        ${rows.map((change) => {
          const read = !unread.some((item) => item.id === change.id);
          return `
            <article class="mailbox-row ${read ? "read" : "unread"}">
              <div class="mailbox-main">
                <strong>${escapeHtml(change.title || "Vorschrift")} wurde geändert</strong>
                <p>Es gibt eine neue Änderung bei ${escapeHtml(change.title || "einer Vorschrift")}.</p>
              </div>
              <div class="button-row">
                <button class="blue-btn open-mail-doc" data-doc-id="${escapeHtml(change.docId)}" data-change-id="${escapeHtml(change.id)}">Änderung öffnen</button>
                ${read ? "" : `<button class="ghost-btn mark-mail-read" data-change-id="${escapeHtml(change.id)}">Als gelesen markieren</button>`}
              </div>
              <footer>${escapeHtml(change.author || "-")} · ${formatDateTime(change.createdAt)}</footer>
            </article>
          `;
        }).join("") || `<p class="muted">Keine Nachrichten vorhanden.</p>`}
      </div>
    </section>
  `;
  document.querySelectorAll(".open-mail-doc").forEach((button) => button.addEventListener("click", async () => {
    await markInformationChangeRead(button.dataset.changeId);
    renderNavigation();
    openInformationDocView(button.dataset.docId, null, button.dataset.changeId);
  }));
  document.querySelectorAll(".mark-mail-read").forEach((button) => button.addEventListener("click", async () => {
    await markInformationChangeRead(button.dataset.changeId);
    await bootstrap();
  }));
}

function renderNavigation() {
  const visiblePages = getVisiblePages();
  const myDuty = state.duty.find((entry) => entry.userId === state.currentUser.id);
  const unreadMail = mailboxUnreadCount();

  $(".profile-card").innerHTML = `
    ${avatarMarkup(state.currentUser, "md")}
    <div>
      <strong>${escapeHtml(fullName())}</strong>
      <small>${escapeHtml(rankLabel(state.currentUser.rank))}</small>
      <span class="service-pill ${myDuty ? "on" : "off"}">${myDuty ? "Im Dienst" : "Nicht im Dienst"}</span>
    </div>
  `;

  $("#navigation").innerHTML = visiblePages.map((page) => `
    <button class="nav-btn ${state.page === page ? "active" : ""}" data-page="${escapeHtml(page)}">
      <span class="nav-icon">${iconSvg(page)}</span>
      <span class="nav-label">${escapeHtml(navLabel(page))}</span>
      ${page === "Postfach" && unreadMail ? `<span class="nav-badge">${unreadMail}</span>` : ""}
      ${restrictedPageIcon(page)}
    </button>
  `).join("");

  document.querySelectorAll(".nav-btn").forEach((button) => {
    button.addEventListener("click", () => {
      if (button.dataset.page === "IT" && state.page !== "IT") localStorage.setItem("lspd_it_tab", "overview");
      state.page = button.dataset.page;
      localStorage.setItem("lspd_page", state.page);
      renderApp();
    });
  });
}

function renderNavigation() {
  const visiblePages = getVisiblePages();
  const myDuty = state.duty.find((entry) => entry.userId === state.currentUser.id);
  const unreadMail = mailboxUnreadCount();

  $(".profile-card").innerHTML = `
    ${avatarMarkup(state.currentUser, "lg")}
    <div class="profile-copy">
      <strong>${escapeHtml(fullName())}</strong>
      <span>${escapeHtml(rankLabel(state.currentUser.rank))}</span>
      <em class="${myDuty ? "on" : "off"}">${myDuty ? "Im Dienst" : "Außer Dienst"}</em>
    </div>
  `;

  $("#navigation").innerHTML = visiblePages.map((page) => `
    <button class="nav-btn ${state.page === page ? "active" : ""}" data-page="${escapeHtml(page)}">
      <span class="nav-icon">${iconSvg(page)}${page === "Postfach" && unreadMail ? `<span class="nav-badge">${unreadMail}</span>` : ""}</span>
      <span class="nav-label">${escapeHtml(navLabel(page))}</span>
      ${restrictedPageIcon(page)}
    </button>
  `).join("");

  document.querySelectorAll(".nav-btn").forEach((button) => {
    button.addEventListener("click", () => {
      if (button.dataset.page === "IT" && state.page !== "IT") localStorage.setItem("lspd_it_tab", "overview");
      state.page = button.dataset.page;
      localStorage.setItem("lspd_page", state.page);
      renderApp();
    });
  });
}

function renderCatalogQuestion(question, index, side = "main") {
  const maxPoints = Number(question.maxPoints || 1);
  const scoreValues = side === "location" ? [0, 0.5, 1] : [0, 0.5, 1, 1.5, 2].filter((value) => value <= maxPoints);
  const scoreClass = (value) => `score-select score-${String(value || 0).replace(".", "-")}`;
  if (side === "location" || question.type === "location") {
    return `
      <article class="exam-catalog-question location-question" data-question-id="${escapeHtml(question.id)}">
        <div class="catalog-question-grid location-score-left">
          <select class="${scoreClass(question.manualPoints)}" data-exam-score="${escapeHtml(question.id)}">${scoreValues.map((value) => `<option value="${value}" ${Number(question.manualPoints || 0) === value ? "selected" : ""}>${value}</option>`).join("")}</select>
          <div class="catalog-question-body"><div class="catalog-question-head"><b>${index + 1}. ${escapeHtml(question.prompt)}</b><small>Ortskunde</small></div>${question.image ? `<img class="location-question-image" src="${escapeHtml(question.image)}" alt="">` : ""}</div>
        </div>
      </article>
    `;
  }
  if (question.type === "choice") {
    const answers = normalizeChoiceAnswers(question);
    return `
      <article class="exam-catalog-question compact-choice-question" data-question-id="${escapeHtml(question.id)}">
        <div class="catalog-question-grid">
          <div class="catalog-question-body">
            <div class="catalog-question-head"><b>${index + 1}. ${escapeHtml(question.prompt)}</b><small>Max. ${escapeHtml(maxPoints)} Punkte</small></div>
            <div class="exam-answer-list neutral compact-answer-list">
              ${answers.map((answer) => `
                <label class="exam-check compact-answer-row">
                  <span>${escapeHtml(answer)}</span>
                  <input data-autosave-exam type="checkbox" name="answerOption_${escapeHtml(question.id)}" value="${escapeHtml(answer)}" ${question.selectedAnswers?.includes(answer) ? "checked" : ""}>
                </label>
              `).join("") || `<p class="muted">Keine Antwortmöglichkeiten hinterlegt.</p>`}
              <label class="exam-check compact-answer-row muted-check"><span>Leer gelassen / nicht beantwortet</span><input data-autosave-exam type="checkbox" name="questionSkipped_${escapeHtml(question.id)}" ${question.skipped ? "checked" : ""}></label>
            </div>
            <div class="penalty-line"><span>Fehlerpunkte</span><select data-exam-penalty="${escapeHtml(question.id)}">${[0, 1, 2, 3].map((value) => `<option value="${value}" ${Number(question.penaltyPoints || 0) === value ? "selected" : ""}>-${value}</option>`).join("")}</select></div>
          </div>
          <select class="${scoreClass(scoreChoiceQuestion(question))}" disabled><option>${escapeHtml(scoreChoiceQuestion(question))}</option></select>
        </div>
      </article>
    `;
  }
  return `
    <article class="exam-catalog-question" data-question-id="${escapeHtml(question.id)}">
      <div class="catalog-question-grid">
        <div class="catalog-question-body">
          <div class="catalog-question-head"><b>${index + 1}. ${escapeHtml(question.prompt)}</b><small>Max. ${escapeHtml(maxPoints)} Punkte</small></div>
          <label>Antwort des Prüflings<textarea data-autosave-exam data-exam-answer="${escapeHtml(question.id)}" placeholder="Antwort mitschreiben">${escapeHtml(question.traineeAnswer || "")}</textarea></label>
          ${question.solution ? `<div class="inline-solution">Musterlösung: ${escapeHtml(question.solution)}</div>` : ""}
        </div>
        <select class="${scoreClass(question.manualPoints)}" data-exam-score="${escapeHtml(question.id)}">${scoreValues.map((value) => `<option value="${value}" ${Number(question.manualPoints || 0) === value ? "selected" : ""}>${value}</option>`).join("")}</select>
      </div>
    </article>
  `;
}

function formatInformationDocText(text = "", searchTerm = "") {
  const term = String(searchTerm || "").trim();
  const renderInline = (value) => {
    let html = escapeHtml(value || "");
    if (term) {
      const safeTerm = escapeHtml(term).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      html = html.replace(new RegExp(`(${safeTerm})`, "gi"), `<mark class="doc-search-mark">$1</mark>`);
    }
    return html
      .replace(/\*\*(.*?)\*\*/g, "<strong>$1</strong>")
      .replace(/\[(.*?)\]\((https?:\/\/.*?)\)/g, `<a href="$2" target="_blank" rel="noreferrer">$1</a>`);
  };
  return String(text || "Noch kein Inhalt vorhanden.")
    .split(/\n/)
    .map((line) => {
      if (line.startsWith("### ")) return `<h4>${renderInline(line.slice(4))}</h4>`;
      if (line.startsWith("## ")) return `<h3>${renderInline(line.slice(3))}</h3>`;
      if (line.startsWith("::center ")) return `<p class="doc-align-center">${renderInline(line.slice(9))}</p>`;
      if (line.startsWith("::green ")) return `<p class="doc-text-green">${renderInline(line.slice(8))}</p>`;
      if (line.startsWith("::red ")) return `<p class="doc-text-red">${renderInline(line.slice(6))}</p>`;
      if (/^\d+\.\s+/.test(line)) return `<p class="doc-list-item numbered">${renderInline(line)}</p>`;
      if (line.startsWith("- ")) return `<p class="doc-list-item">${renderInline(line.slice(2))}</p>`;
      return line.trim() ? `<p>${renderInline(line)}</p>` : "<br>";
    })
    .join("");
}

function openInformationDocView(docId, draft = null, focusChangeId = "") {
  const existing = informationDocs().find((item) => item.id === docId);
  const doc = existing || { id: docId || makeTrainingId("infodoc"), title: draft?.title || "Neue Vorschrift", body: draft?.body || "", updatedAt: new Date().toISOString() };
  const canEdit = canAccess("actions", "manageInformation", "Direktion");
  const changes = informationDocChangesFor(doc.id);
  const readBody = draft?.body ?? doc.body ?? "";
  openModal(`
    <div class="paper-doc-modal">
      <div class="paper-doc-head no-title">
        <input id="docSearchInput" placeholder="Im Dokument suchen">
        ${canEdit ? `<button class="mode-toggle-btn" id="toggleDocMode" type="button">Bearbeitermodus</button>` : ""}
        <button class="ghost-btn" id="toggleDocChanges" type="button">Changelog (${changes.length})</button>
      </div>
      <div class="doc-read-mode" id="docReadMode"><article class="paper-doc-page" id="paperDocPage">${formatInformationDocText(readBody)}</article></div>
      ${canEdit ? `<div class="doc-edit-mode hidden" id="docEditMode"><input id="paperDocTitle" value="${escapeHtml(draft?.title ?? doc.title)}"><div class="format-toolbar"><button type="button" data-format="## ">Überschrift</button><button type="button" data-format="**fett**">Fett</button><button type="button" data-format="<span style='color:#75ffad'>Grün</span>">Grün</button><button type="button" data-format="<span style='color:#ff9ca0'>Rot</span>">Rot</button></div><textarea class="paper-doc-page paper-doc-editor" id="paperDocEditor">${escapeHtml(readBody)}</textarea></div>` : ""}
      <div class="doc-change-details hidden" id="docChangePanel">
        <div class="info-change-list">${changes.map((change) => `<article class="info-change-row ${change.id === focusChangeId ? "focus-change" : ""}"><strong>${escapeHtml(change.action || "geändert")}</strong><small>${escapeHtml(change.author || "-")} · ${formatDateTime(change.createdAt)}</small><div><del>${escapeHtml(change.before || "-")}</del><ins>${escapeHtml(change.after || "-")}</ins></div></article>`).join("") || `<p class="muted">Noch keine Änderungen.</p>`}</div>
      </div>
    </div>
  `, (modal) => {
    modal.classList.add("wide-doc-modal");
    let editMode = false;
    const initial = readBody;
    const setMode = (next) => {
      editMode = next;
      modal.querySelector("#docReadMode")?.classList.toggle("hidden", editMode);
      modal.querySelector("#docEditMode")?.classList.toggle("hidden", !editMode);
      const btn = modal.querySelector("#toggleDocMode");
      if (btn) btn.textContent = editMode ? "Lesemodus" : "Bearbeitermodus";
      btn?.classList.toggle("active", editMode);
    };
    modal.querySelector("#toggleDocMode")?.addEventListener("click", () => setMode(!editMode));
    modal.querySelector("#toggleDocChanges")?.addEventListener("click", () => modal.querySelector("#docChangePanel")?.classList.toggle("hidden"));
    if (focusChangeId) modal.querySelector("#docChangePanel")?.classList.remove("hidden");
    const x = modal.querySelector(".modal-x");
    if (x && canEdit) {
      const clone = x.cloneNode(true);
      x.replaceWith(clone);
      clone.addEventListener("click", () => {
        const title = modal.querySelector("#paperDocTitle")?.value.trim() || doc.title;
        const current = modal.querySelector("#paperDocEditor")?.value || initial;
        if (current !== initial || title !== doc.title) openInformationDocCloseConfirm(doc, title, initial, current);
        else closeModal();
      });
    }
    modal.querySelector("#docSearchInput")?.addEventListener("input", (event) => {
      const term = event.target.value.trim();
      const page = modal.querySelector("#paperDocPage");
      if (page) page.innerHTML = formatInformationDocText(readBody, term);
      const editor = modal.querySelector("#paperDocEditor");
      if (editor && editMode && term) {
        const index = editor.value.toLowerCase().indexOf(term.toLowerCase());
        if (index >= 0) editor.setSelectionRange(index, index + term.length);
      }
    });
    modal.querySelectorAll("[data-format]").forEach((button) => button.addEventListener("click", () => {
      const area = modal.querySelector("#paperDocEditor");
      area.setRangeText(button.dataset.format, area.selectionStart, area.selectionEnd, "end");
      area.focus();
    }));
  });
}

function openInformationDocView(docId, draft = null, focusChangeId = "") {
  const existing = informationDocs().find((item) => item.id === docId);
  const doc = existing || { id: docId || makeTrainingId("infodoc"), title: draft?.title || "Neue Vorschrift", body: draft?.body || "", updatedAt: new Date().toISOString() };
  const canEdit = canAccess("actions", "manageInformation", "Direktion");
  const changes = informationDocChangesFor(doc.id);
  const readBody = draft?.body ?? doc.body ?? "";
  openModal(`
    <div class="paper-doc-modal side-toolbar-doc">
      <div class="doc-read-mode" id="docReadMode"><article class="paper-doc-page" id="paperDocPage">${formatInformationDocText(readBody)}</article></div>
      ${canEdit ? `<div class="doc-edit-mode hidden" id="docEditMode">
        <input id="paperDocTitle" value="${escapeHtml(draft?.title ?? doc.title)}">
        <div class="format-toolbar doc-editor-toolbar">
          <button type="button" data-doc-format="heading">Überschrift</button>
          <button type="button" data-doc-format="bold">Fett</button>
          <button type="button" data-doc-format="center">Zentriert</button>
          <button type="button" data-doc-format="bullet">Liste</button>
          <button type="button" data-doc-format="number">Nummeriert</button>
          <button type="button" data-doc-format="link">Link</button>
          <button type="button" data-doc-format="green">Grün</button>
          <button type="button" data-doc-format="red">Rot</button>
        </div>
        <textarea class="paper-doc-page paper-doc-editor" id="paperDocEditor">${escapeHtml(readBody)}</textarea>
      </div>` : ""}
      <aside class="paper-doc-tools">
        <div class="doc-search-control">
          <input id="docSearchInput" placeholder="Im Dokument suchen">
          <span id="docSearchCount">0 Treffer</span>
        </div>
        ${canEdit ? `<button class="mode-toggle-btn" id="toggleDocMode" type="button">Bearbeitermodus</button><button class="blue-btn hidden" id="saveDocFromEditor" type="button">Speichern</button>` : ""}
        <button class="ghost-btn" id="toggleDocChanges" type="button">Changelog (${changes.length})</button>
      </aside>
    </div>
  `, (modal) => {
    modal.classList.add("wide-doc-modal");
    let editMode = false;
    const initial = readBody;
    const area = () => modal.querySelector("#paperDocEditor");
    const setMode = (next) => {
      editMode = next;
      modal.querySelector("#docReadMode")?.classList.toggle("hidden", editMode);
      modal.querySelector("#docEditMode")?.classList.toggle("hidden", !editMode);
      modal.querySelector("#saveDocFromEditor")?.classList.toggle("hidden", !editMode);
      const btn = modal.querySelector("#toggleDocMode");
      if (btn) btn.textContent = editMode ? "Lesemodus" : "Bearbeitermodus";
      btn?.classList.toggle("active", editMode);
    };
    const insertFormat = (type) => {
      const editor = area();
      if (!editor) return;
      const start = editor.selectionStart;
      const end = editor.selectionEnd;
      const selected = editor.value.slice(start, end) || "Text";
      const linePrefix = start === 0 || editor.value[start - 1] === "\n" ? "" : "\n";
      const formats = {
        heading: `## ${selected}`,
        bold: `**${selected}**`,
        center: `::center ${selected}`,
        bullet: `- ${selected}`,
        number: `1. ${selected}`,
        link: `[${selected}](https://)`,
        green: `::green ${selected}`,
        red: `::red ${selected}`
      };
      const value = ["heading", "center", "bullet", "number", "green", "red"].includes(type) ? `${linePrefix}${formats[type]}` : formats[type];
      editor.setRangeText(value, start, end, "end");
      editor.focus();
    };
    modal.querySelector("#toggleDocMode")?.addEventListener("click", () => setMode(!editMode));
    modal.querySelector("#toggleDocChanges")?.addEventListener("click", () => openInformationDocChangelog(doc.id, focusChangeId));
    modal.querySelector("#saveDocFromEditor")?.addEventListener("click", async () => {
      try {
        const title = modal.querySelector("#paperDocTitle")?.value.trim() || doc.title;
        const current = area()?.value || "";
        await saveInformationDocDirect(doc, title, current, false);
      } catch (error) {
        showNotify(error.message, "error");
      }
    });
    if (focusChangeId) window.setTimeout(() => openInformationDocChangelog(doc.id, focusChangeId), 50);
    const x = modal.querySelector(".modal-x");
    if (x && canEdit) {
      const clone = x.cloneNode(true);
      x.replaceWith(clone);
      clone.addEventListener("click", () => {
        const title = modal.querySelector("#paperDocTitle")?.value.trim() || doc.title;
        const current = area()?.value || initial;
        if (current !== initial || title !== doc.title) openInformationDocCloseConfirm(doc, title, initial, current);
        else closeModal();
      });
    }
    modal.querySelector("#docSearchInput")?.addEventListener("input", (event) => {
      const term = event.target.value.trim();
      const page = modal.querySelector("#paperDocPage");
      if (page) page.innerHTML = formatInformationDocText(readBody, term);
      const editor = area();
      if (editor && editMode && term) {
        const index = editor.value.toLowerCase().indexOf(term.toLowerCase());
        if (index >= 0) editor.setSelectionRange(index, index + term.length);
      }
    });
    modal.querySelectorAll("[data-doc-format]").forEach((button) => button.addEventListener("click", () => insertFormat(button.dataset.docFormat)));
  });
}

function renderInformation() {
  const links = state.settings.informationLinks || [];
  const docs = informationDocs();
  const permits = state.settings.informationPermits || [];
  const factions = state.settings.informationFactions || [];
  content.innerHTML = `
    <section class="department-info-view information-admin-view modern-info-view">
      <div class="info-box full information-card internal-doc-card"><div class="department-modal-heading"><h4>Vorschriften</h4>${canAccess("actions", "manageInformation", "Direktion") ? `<button class="blue-btn add-doc-button" id="addInformationDoc">${iconSvg("Plus")} Neue Vorschrift hinzufügen</button>` : ""}</div><div class="internal-doc-grid">${docs.map((doc) => `<article class="internal-doc-tile-wrap"><button class="internal-doc-tile" data-doc-id="${escapeHtml(doc.id)}"><strong>${escapeHtml(doc.title)}</strong><small>Zuletzt geändert: ${formatDateTime(doc.updatedAt)}</small></button>${canAccess("actions", "manageInformation", "Direktion") ? `<button class="mini-icon danger delete-info-doc" type="button" data-id="${escapeHtml(doc.id)}" title="Vorschrift löschen">${actionIcon("delete")}</button>` : ""}</article>`).join("")}</div></div>
      <div class="info-box full information-card redirects-card"><div class="department-modal-heading"><h4>Link Weiterleitungen</h4>${canAccess("actions", "manageInformation", "Direktion") ? `<button class="blue-btn" id="addInformationLink">${iconSvg("Plus")} Hinzufügen</button>` : ""}</div><div class="link-card-grid">${links.map((link) => `<article class="small-link-card"><strong>${escapeHtml(link.title)}</strong><span class="link-label">Link:</span><a href="${escapeHtml(link.url)}" target="_blank" rel="noreferrer">${escapeHtml(link.url)}</a>${canAccess("actions", "manageInformation", "Direktion") ? `<span class="button-row"><button class="blue-btn compact-action edit-info-link" data-id="${link.id}" title="Bearbeiten">${actionIcon("edit")} Bearbeiten</button><button class="mini-icon danger delete-info-link" data-id="${link.id}" title="Löschen">${actionIcon("delete")}</button></span>` : ""}</article>`).join("") || `<p class="muted">Noch keine Weiterleitungen.</p>`}</div></div>
      <div class="info-box full information-card"><div class="department-modal-heading"><h4>Rechte Definition</h4>${canAccess("actions", "manageInformation", "Direktion") ? `<button class="blue-btn" id="editInformationRights">${actionIcon("edit")} Bearbeiten</button>` : ""}</div><div class="rich-text-view">${formatDepartmentText(state.settings.informationRightsText)}</div></div>
      <div class="info-box full information-card"><div class="department-modal-heading"><h4>Sondergenehmigungen</h4>${canAccess("actions", "manageInformation", "Direktion") ? `<button class="blue-btn" id="addInformationPermit">${iconSvg("Plus")} Hinzufügen</button>` : ""}</div><div class="table-wrap compact-table"><table><thead><tr><th>Vor- und Nachname</th><th>Beschreibung</th><th>Gültig Bis</th><th>Aktionen</th></tr></thead><tbody>${permits.map((permit) => `<tr><td>${escapeHtml(permit.name)}</td><td>${escapeHtml(permit.description)}</td><td>${formatDate(permit.validUntil)}</td><td>${canAccess("actions", "manageInformation", "Direktion") ? `<button class="mini-icon edit-info-permit" data-id="${permit.id}">${actionIcon("edit")}</button><button class="mini-icon danger delete-info-permit" data-id="${permit.id}">${actionIcon("delete")}</button>` : ""}</td></tr>`).join("") || `<tr><td colspan="4" class="muted">Keine Sondergenehmigungen.</td></tr>`}</tbody></table></div></div>
      <div class="info-box full information-card"><div class="department-modal-heading"><h4>Fraktionen</h4>${canAccess("actions", "manageInformation", "Direktion") ? `<button class="blue-btn" id="addInformationFaction">${iconSvg("Plus")} Hinzufügen</button>` : ""}</div><div class="table-wrap compact-table"><table><thead><tr><th>Organisation</th><th>Status</th><th>Aktionen</th></tr></thead><tbody>${factions.map((faction) => `<tr><td>${escapeHtml(faction.organization)}</td><td><span class="status-label">${renderStatusDot(faction.status)}</span></td><td>${canAccess("actions", "manageInformation", "Direktion") ? `<button class="mini-icon edit-info-faction" data-id="${faction.id}">${actionIcon("edit")}</button><button class="mini-icon danger delete-info-faction" data-id="${faction.id}">${actionIcon("delete")}</button>` : ""}</td></tr>`).join("") || `<tr><td colspan="3" class="muted">Keine Fraktionen.</td></tr>`}</tbody></table></div></div>
    </section>
  `;
  $("#editInformation")?.addEventListener("click", openInformationEditModal);
  $("#editInformationRights")?.addEventListener("click", openInformationRightsModal);
  $("#addInformationLink")?.addEventListener("click", () => openInformationLinkModal());
  $("#addInformationDoc")?.addEventListener("click", () => openInformationDocView(makeTrainingId("infodoc")));
  $("#addInformationPermit")?.addEventListener("click", () => openInformationPermitModal());
  $("#addInformationFaction")?.addEventListener("click", () => openInformationFactionModal());
  document.querySelectorAll(".edit-info-link").forEach((button) => button.addEventListener("click", () => openInformationLinkModal(links.find((item) => item.id === button.dataset.id))));
  document.querySelectorAll(".delete-info-link").forEach((button) => button.addEventListener("click", () => openDeleteInformationConfirm("informationLinks", button.dataset.id, "Weiterleitung löschen?")));
  document.querySelectorAll(".edit-info-permit").forEach((button) => button.addEventListener("click", () => openInformationPermitModal(permits.find((item) => item.id === button.dataset.id))));
  document.querySelectorAll(".delete-info-permit").forEach((button) => button.addEventListener("click", () => deleteInformationItem("informationPermits", button.dataset.id)));
  document.querySelectorAll(".edit-info-faction").forEach((button) => button.addEventListener("click", () => openInformationFactionModal(factions.find((item) => item.id === button.dataset.id))));
  document.querySelectorAll(".delete-info-faction").forEach((button) => button.addEventListener("click", () => deleteInformationItem("informationFactions", button.dataset.id)));
  document.querySelectorAll(".internal-doc-tile").forEach((button) => button.addEventListener("click", () => openInformationDocView(button.dataset.docId)));
  document.querySelectorAll(".delete-info-doc").forEach((button) => button.addEventListener("click", (event) => {
    event.stopPropagation();
    openDeleteInformationDocConfirm(button.dataset.id);
  }));
}

function openDeleteInformationDocConfirm(docId) {
  const doc = informationDocs().find((item) => item.id === docId);
  if (!doc || !canAccess("actions", "manageInformation", "Direktion")) return;
  openModal(`
    <h3>Vorschrift löschen?</h3>
    <p class="muted">Die Vorschrift <strong>${escapeHtml(doc.title)}</strong> wird dauerhaft entfernt.</p>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions">
      <button class="ghost-btn" data-close>Abbrechen</button>
      <button class="red-btn" id="confirmDeleteInfoDoc">Löschen</button>
    </div>
  `, (modal) => {
    modal.querySelector("#confirmDeleteInfoDoc").addEventListener("click", async () => {
      try {
        await saveInformationPatch({
          informationDocs: informationDocs().filter((item) => item.id !== docId),
          informationDocChanges: (state.settings.informationDocChanges || []).filter((change) => change.docId !== docId)
        });
        closeModal();
        renderInformation();
      } catch (error) {
        modal.querySelector("#modalError").textContent = error.message;
      }
    });
  });
}

async function deleteInformationDocChange(changeId, docId) {
  const changes = (state.settings.informationDocChanges || []).filter((change) => change.id !== changeId);
  await saveInformationPatch({ informationDocChanges: changes });
  openInformationDocView(docId);
}

function openInformationDocChangelog(docId, focusChangeId = "") {
  const doc = informationDocs().find((item) => item.id === docId);
  if (!doc) return;
  const canManage = canAccess("actions", "manageInformation", "Direktion");
  const changes = informationDocChangesFor(docId);
  openModal(`
    <div class="doc-compare-head">
      <span class="doc-compare-kicker">Changelog</span>
      <h3>${escapeHtml(doc.title)}</h3>
      <p>Alle gespeicherten Änderungen dieser Vorschrift mit direktem Vorher/Nachher-Vergleich.</p>
    </div>
    <div class="doc-changelog-modal-list">
      ${changes.map((change) => `
        <article class="info-change-row ${change.id === focusChangeId ? "focus-change" : ""}">
          <div class="change-row-head">
            <span><strong>${escapeHtml(change.action || "geändert")}</strong><small>${escapeHtml(change.author || "-")} · ${formatDateTime(change.createdAt)}</small></span>
            ${canManage ? `<button class="mini-icon danger delete-doc-change" data-change-id="${escapeHtml(change.id)}" title="Changelog löschen">${actionIcon("delete")}</button>` : ""}
          </div>
          <div class="doc-compare-grid compact">
            <section class="doc-compare-panel before">
              <header><span>Vorher</span><small>Alte Version</small></header>
              <div class="doc-change-preview before">${formatInformationDocText(change.before || "")}</div>
            </section>
            <section class="doc-compare-panel after">
              <header><span>Nachher</span><small>Neue Version</small></header>
              <div class="doc-change-preview after">${formatInformationDocText(change.after || "")}</div>
            </section>
          </div>
        </article>
      `).join("") || `<p class="muted">Noch keine Änderungen vorhanden.</p>`}
    </div>
    <div class="modal-actions"><button class="ghost-btn" data-close>Schließen</button></div>
  `, (modal) => {
    modal.classList.add("doc-changelog-modal");
    modal.querySelectorAll(".delete-doc-change").forEach((button) => button.addEventListener("click", async () => {
      try {
        await deleteInformationDocChange(button.dataset.changeId, docId);
      } catch (error) {
        showNotify(error.message, "error");
      }
    }));
  });
}

function openInformationDocView(docId, draft = null, focusChangeId = "") {
  const existing = informationDocs().find((item) => item.id === docId);
  const doc = existing || { id: docId || makeTrainingId("infodoc"), title: draft?.title || "Neue Vorschrift", body: draft?.body || "", updatedAt: new Date().toISOString() };
  const canEdit = canAccess("actions", "manageInformation", "Direktion");
  const changes = informationDocChangesFor(doc.id);
  const readBody = draft?.body ?? doc.body ?? "";
  openModal(`
    <div class="paper-doc-modal side-toolbar-doc">
      <div class="doc-read-mode" id="docReadMode"><article class="paper-doc-page" id="paperDocPage">${formatInformationDocText(readBody)}</article></div>
      ${canEdit ? `<div class="doc-edit-mode hidden" id="docEditMode"><input id="paperDocTitle" value="${escapeHtml(draft?.title ?? doc.title)}"><div class="format-toolbar"><button type="button" data-format="## ">Überschrift</button><button type="button" data-format="**fett**">Fett</button><button type="button" data-format="<span style='color:#75ffad'>Grün</span>">Grün</button><button type="button" data-format="<span style='color:#ff9ca0'>Rot</span>">Rot</button></div><textarea class="paper-doc-page paper-doc-editor" id="paperDocEditor">${escapeHtml(readBody)}</textarea></div>` : ""}
      <aside class="paper-doc-tools">
        <input id="docSearchInput" placeholder="Im Dokument suchen">
        ${canEdit ? `<button class="mode-toggle-btn" id="toggleDocMode" type="button">Bearbeitermodus</button>` : ""}
        <button class="ghost-btn" id="toggleDocChanges" type="button">Changelog (${changes.length})</button>
      </aside>
    </div>
  `, (modal) => {
    modal.classList.add("wide-doc-modal");
    let editMode = false;
    const initial = readBody;
    const setMode = (next) => {
      editMode = next;
      modal.querySelector("#docReadMode")?.classList.toggle("hidden", editMode);
      modal.querySelector("#docEditMode")?.classList.toggle("hidden", !editMode);
      const btn = modal.querySelector("#toggleDocMode");
      if (btn) btn.textContent = editMode ? "Lesemodus" : "Bearbeitermodus";
      btn?.classList.toggle("active", editMode);
    };
    modal.querySelector("#toggleDocMode")?.addEventListener("click", () => setMode(!editMode));
    modal.querySelector("#toggleDocChanges")?.addEventListener("click", () => openInformationDocChangelog(doc.id, focusChangeId));
    if (focusChangeId) window.setTimeout(() => openInformationDocChangelog(doc.id, focusChangeId), 50);
    const x = modal.querySelector(".modal-x");
    if (x && canEdit) {
      const clone = x.cloneNode(true);
      x.replaceWith(clone);
      clone.addEventListener("click", () => {
        const title = modal.querySelector("#paperDocTitle")?.value.trim() || doc.title;
        const current = modal.querySelector("#paperDocEditor")?.value || initial;
        if (current !== initial || title !== doc.title) openInformationDocCloseConfirm(doc, title, initial, current);
        else closeModal();
      });
    }
    modal.querySelector("#docSearchInput")?.addEventListener("input", (event) => {
      const term = event.target.value.trim();
      const page = modal.querySelector("#paperDocPage");
      if (page) page.innerHTML = formatInformationDocText(readBody, term);
      const editor = modal.querySelector("#paperDocEditor");
      if (editor && editMode && term) {
        const index = editor.value.toLowerCase().indexOf(term.toLowerCase());
        if (index >= 0) editor.setSelectionRange(index, index + term.length);
      }
    });
    modal.querySelectorAll("[data-format]").forEach((button) => button.addEventListener("click", () => {
      const area = modal.querySelector("#paperDocEditor");
      area.setRangeText(button.dataset.format, area.selectionStart, area.selectionEnd, "end");
      area.focus();
    }));
  });
}

function openInformationDocView(docId, draft = null, focusChangeId = "") {
  const existing = informationDocs().find((item) => item.id === docId);
  const doc = existing || { id: docId || makeTrainingId("infodoc"), title: draft?.title || "Neue Vorschrift", body: draft?.body || "", updatedAt: new Date().toISOString() };
  const canEdit = canAccess("actions", "manageInformation", "Direktion");
  const changes = informationDocChangesFor(doc.id);
  const readBody = draft?.body ?? doc.body ?? "";
  openModal(`
    <div class="paper-doc-modal side-toolbar-doc">
      <div class="doc-read-mode" id="docReadMode"><article class="paper-doc-page" id="paperDocPage">${formatInformationDocText(readBody)}</article></div>
      ${canEdit ? `<div class="doc-edit-mode hidden" id="docEditMode">
        <input id="paperDocTitle" value="${escapeHtml(draft?.title ?? doc.title)}">
        <div class="format-toolbar doc-editor-toolbar">
          <button type="button" data-doc-format="heading">Überschrift</button>
          <button type="button" data-doc-format="bold">Fett</button>
          <button type="button" data-doc-format="center">Zentriert</button>
          <button type="button" data-doc-format="bullet">Liste</button>
          <button type="button" data-doc-format="number">Nummeriert</button>
          <button type="button" data-doc-format="link">Link</button>
          <button type="button" data-doc-format="green">Grün</button>
          <button type="button" data-doc-format="red">Rot</button>
        </div>
        <textarea class="paper-doc-page paper-doc-editor" id="paperDocEditor">${escapeHtml(readBody)}</textarea>
      </div>` : ""}
      <aside class="paper-doc-tools">
        <input id="docSearchInput" placeholder="Im Dokument suchen">
        <small class="doc-search-count" id="docSearchCount">0 Treffer</small>
        ${canEdit ? `<button class="mode-toggle-btn" id="toggleDocMode" type="button">Bearbeitermodus</button><button class="blue-btn hidden" id="saveDocFromEditor" type="button">Speichern</button>` : ""}
        <button class="ghost-btn" id="toggleDocChanges" type="button">Changelog (${changes.length})</button>
      </aside>
    </div>
  `, (modal) => {
    modal.classList.add("wide-doc-modal");
    let editMode = false;
    const initial = readBody;
    const editor = () => modal.querySelector("#paperDocEditor");
    const setMode = (next) => {
      editMode = next;
      modal.querySelector("#docReadMode")?.classList.toggle("hidden", editMode);
      modal.querySelector("#docEditMode")?.classList.toggle("hidden", !editMode);
      modal.querySelector("#saveDocFromEditor")?.classList.toggle("hidden", !editMode);
      const btn = modal.querySelector("#toggleDocMode");
      if (btn) btn.textContent = editMode ? "Lesemodus" : "Bearbeitermodus";
      btn?.classList.toggle("active", editMode);
    };
    const insertFormat = (type) => {
      const area = editor();
      if (!area) return;
      const start = area.selectionStart;
      const end = area.selectionEnd;
      const selected = area.value.slice(start, end) || "Text";
      const linePrefix = start === 0 || area.value[start - 1] === "\n" ? "" : "\n";
      const formats = {
        heading: `## ${selected}`,
        bold: `**${selected}**`,
        center: `::center ${selected}`,
        bullet: `- ${selected}`,
        number: `1. ${selected}`,
        link: `[${selected}](https://)`,
        green: `::green ${selected}`,
        red: `::red ${selected}`
      };
      const value = ["heading", "center", "bullet", "number", "green", "red"].includes(type) ? `${linePrefix}${formats[type]}` : formats[type];
      area.setRangeText(value, start, end, "end");
      area.focus();
    };
    modal.querySelector("#toggleDocMode")?.addEventListener("click", () => setMode(!editMode));
    modal.querySelector("#toggleDocChanges")?.addEventListener("click", () => openInformationDocChangelog(doc.id, focusChangeId));
    modal.querySelector("#saveDocFromEditor")?.addEventListener("click", async () => {
      try {
        const title = modal.querySelector("#paperDocTitle")?.value.trim() || doc.title;
        const current = editor()?.value || "";
        if (current === initial && title === doc.title) {
          showNotify("Keine Änderungen vorhanden.", "info");
          return;
        }
        openInformationDocCloseConfirm(doc, title, initial, current);
      } catch (error) {
        showNotify(error.message, "error");
      }
    });
    if (focusChangeId) window.setTimeout(() => openInformationDocChangelog(doc.id, focusChangeId), 50);
    const x = modal.querySelector(".modal-x");
    if (x && canEdit) {
      const clone = x.cloneNode(true);
      x.replaceWith(clone);
      clone.addEventListener("click", closeModal);
    }
    modal.querySelector("#docSearchInput")?.addEventListener("input", (event) => {
      const term = event.target.value.trim();
      const page = modal.querySelector("#paperDocPage");
      if (page) page.innerHTML = formatInformationDocText(readBody, term);
      const area = editor();
      if (area && editMode && term) {
        const index = area.value.toLowerCase().indexOf(term.toLowerCase());
        if (index >= 0) area.setSelectionRange(index, index + term.length);
      }
    });
    modal.querySelectorAll("[data-doc-format]").forEach((button) => button.addEventListener("click", () => insertFormat(button.dataset.docFormat)));
  });
}

function sanitizeInformationHtml(html = "") {
  const template = document.createElement("template");
  template.innerHTML = String(html || "");
  template.content.querySelectorAll("script, style, iframe, object, embed").forEach((node) => node.remove());
  template.content.querySelectorAll("*").forEach((node) => {
    Array.from(node.attributes).forEach((attribute) => {
      const name = attribute.name.toLowerCase();
      const value = attribute.value || "";
      if (name.startsWith("on")) node.removeAttribute(attribute.name);
      if (["href", "src"].includes(name) && !/^(https?:|mailto:|tel:|#|\/)/i.test(value)) node.removeAttribute(attribute.name);
    });
  });
  return template.innerHTML;
}

function highlightInformationHtml(html = "", searchTerm = "") {
  const term = String(searchTerm || "").trim();
  if (!term) return html;
  const wrapper = document.createElement("div");
  wrapper.innerHTML = html;
  const matcher = new RegExp(`(${term.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")})`, "gi");
  const walker = document.createTreeWalker(wrapper, NodeFilter.SHOW_TEXT);
  const nodes = [];
  while (walker.nextNode()) nodes.push(walker.currentNode);
  nodes.forEach((node) => {
    if (!matcher.test(node.nodeValue || "")) return;
    matcher.lastIndex = 0;
    const span = document.createElement("span");
    span.innerHTML = escapeHtml(node.nodeValue || "").replace(matcher, '<mark class="doc-search-mark">$1</mark>');
    node.replaceWith(...Array.from(span.childNodes));
  });
  return wrapper.innerHTML;
}

function informationTextToHtml(text = "") {
  const raw = String(text || "").trim();
  if (!raw) return "<p>Noch kein Inhalt vorhanden.</p>";
  if (/<\/?(p|h[1-6]|strong|b|em|i|u|ul|ol|li|a|span|div|section|article|hr|br|blockquote|font)\b/i.test(raw)) {
    return sanitizeInformationHtml(raw);
  }
  const lines = String(text || "").split(/\n/);
  const html = [];
  let listMode = "";
  const closeList = () => {
    if (listMode) html.push(`</${listMode}>`);
    listMode = "";
  };
  const inline = (value) => escapeHtml(value || "")
    .replace(/\*\*(.*?)\*\*/g, "<strong>$1</strong>")
    .replace(/\[(.*?)\]\((https?:\/\/.*?)\)/g, '<a href="$2" target="_blank" rel="noreferrer">$1</a>');
  lines.forEach((line) => {
    const trimmed = line.trim();
    if (!trimmed) {
      closeList();
      html.push("<p><br></p>");
      return;
    }
    if (/^━{5,}$/.test(trimmed)) {
      closeList();
      html.push('<hr class="doc-section-divider">');
      return;
    }
    if (trimmed.startsWith("# ")) {
      closeList();
      html.push(`<h2>${inline(trimmed.slice(2))}</h2>`);
      return;
    }
    if (trimmed.startsWith("### ")) {
      closeList();
      html.push(`<h4>${inline(trimmed.slice(4))}</h4>`);
      return;
    }
    if (trimmed.startsWith("## ")) {
      closeList();
      html.push(`<h3>${inline(trimmed.slice(3))}</h3>`);
      return;
    }
    if (trimmed.startsWith("::center ")) {
      closeList();
      html.push(`<p class="doc-align-center">${inline(trimmed.slice(9))}</p>`);
      return;
    }
    if (trimmed.startsWith("::green ")) {
      closeList();
      html.push(`<p><span style="color: #75ffad;">${inline(trimmed.slice(8))}</span></p>`);
      return;
    }
    if (trimmed.startsWith("::red ")) {
      closeList();
      html.push(`<p><span style="color: #ff9ca0;">${inline(trimmed.slice(6))}</span></p>`);
      return;
    }
    if (trimmed.startsWith("- ")) {
      if (listMode !== "ul") {
        closeList();
        html.push("<ul>");
        listMode = "ul";
      }
      html.push(`<li>${inline(trimmed.slice(2))}</li>`);
      return;
    }
    if (/^\d+\.\s+/.test(trimmed)) {
      if (listMode !== "ol") {
        closeList();
        html.push("<ol>");
        listMode = "ol";
      }
      html.push(`<li>${inline(trimmed.replace(/^\d+\.\s+/, ""))}</li>`);
      return;
    }
    if (/^§\d+(?!\.)\s+/.test(trimmed)) {
      closeList();
      html.push(`<h3 class="doc-clause-heading">${inline(trimmed)}</h3>`);
      return;
    }
    if (/^§\d+(\.\d+)+/.test(trimmed)) {
      closeList();
      html.push(`<p class="doc-clause">${inline(trimmed)}</p>`);
      return;
    }
    if (/^[A-ZÄÖÜ][^.!?]{2,45}:$/.test(trimmed)) {
      closeList();
      html.push(`<h4 class="doc-subheading">${inline(trimmed.replace(/:$/, ""))}</h4>`);
      return;
    }
    closeList();
    html.push(`<p>${inline(trimmed)}</p>`);
  });
  closeList();
  return sanitizeInformationHtml(html.join(""));
}

function formatInformationDocText(text = "", searchTerm = "") {
  return highlightInformationHtml(informationTextToHtml(text), searchTerm);
}

function openInformationDocView(docId, draft = null, focusChangeId = "") {
  const existing = informationDocs().find((item) => item.id === docId);
  const doc = existing || { id: docId || makeTrainingId("infodoc"), title: draft?.title || "Neue Vorschrift", body: draft?.body || "", updatedAt: new Date().toISOString() };
  const canEdit = canAccess("actions", "manageInformation", "Direktion");
  const changes = informationDocChangesFor(doc.id);
  const readBody = draft?.body ?? doc.body ?? "";
  const editorHtml = informationTextToHtml(readBody);
  openModal(`
    <div class="paper-doc-modal side-toolbar-doc">
      <div class="doc-read-mode" id="docReadMode">
        <article class="paper-doc-page" id="paperDocPage">${formatInformationDocText(readBody)}</article>
      </div>
      ${canEdit ? `<div class="doc-edit-mode hidden information-editor-workspace" id="docEditMode">
        <div class="information-editor-sticky">
          <input id="paperDocTitle" value="${escapeHtml(draft?.title ?? doc.title)}" aria-label="Dokumenttitel">
          <div class="docs-editor-toolbar" aria-label="Dokumentformatierung">
            <div class="docs-toolbar-group">
              <span class="docs-toolbar-label">Text</span>
              <select id="docBlockStyle" title="Formatvorlage">
                <option value="P">Normaler Text</option>
                <option value="H2">Titel</option>
                <option value="H3">Überschrift</option>
                <option value="H4">Unterüberschrift</option>
              </select>
              <select id="docFontSize" title="Textgröße">
                <option value="">Textgröße</option>
                <option value="2">Klein</option>
                <option value="3">Normal</option>
                <option value="4">Groß</option>
                <option value="5">Sehr groß</option>
              </select>
              <button type="button" class="docs-tool-btn" data-wysiwyg="bold" title="Fett (Strg+B)"><strong>B</strong></button>
              <button type="button" class="docs-tool-btn" data-wysiwyg="italic" title="Kursiv"><em>I</em></button>
              <button type="button" class="docs-tool-btn" data-wysiwyg="underline" title="Unterstrichen"><u>U</u></button>
            </div>
            <div class="docs-toolbar-group">
              <span class="docs-toolbar-label">Farbe</span>
              <label class="docs-color-menu" title="Textfarbe"><span>A</span><input id="docTextColor" type="color" value="#dce8f8"></label>
              <label class="docs-color-menu" title="Markierung"><span>▰</span><input id="docHighlightColor" type="color" value="#1f4f9b"></label>
            </div>
            <div class="docs-toolbar-group">
              <span class="docs-toolbar-label">Absatz</span>
              <button type="button" class="docs-tool-btn" data-wysiwyg="justifyLeft" title="Linksbündig">☰</button>
              <button type="button" class="docs-tool-btn" data-wysiwyg="justifyCenter" title="Zentriert">≡</button>
              <button type="button" class="docs-tool-btn" data-wysiwyg="justifyRight" title="Rechtsbündig">☷</button>
              <button type="button" class="docs-tool-btn wide" data-wysiwyg="insertUnorderedList" title="Aufzählung">• Liste</button>
              <button type="button" class="docs-tool-btn wide" data-wysiwyg="insertOrderedList" title="Nummerierung">1. Liste</button>
            </div>
            <div class="docs-toolbar-group">
              <span class="docs-toolbar-label">Einfügen</span>
              <button type="button" class="docs-tool-btn wide accent" data-insert-card="info" title="Eigene Kachel einfügen">Kachel</button>
              <button type="button" class="docs-tool-btn wide" data-insert-card="warning" title="Warn-Kachel einfügen">Warnung</button>
              <button type="button" class="docs-tool-btn wide" data-wysiwyg="createLink" title="Link einfügen">Link</button>
            </div>
            <div class="docs-toolbar-group compact">
              <button type="button" class="docs-tool-btn wide" id="autoFormatDoc" title="Vorschriften automatisch formatieren">Auto-Format</button>
              <button type="button" class="docs-tool-btn" data-wysiwyg="removeFormat" title="Formatierung entfernen">Tx</button>
            </div>
          </div>
        </div>
        <article class="paper-doc-page paper-doc-editor wysiwyg-editor" id="paperDocEditor" contenteditable="true" spellcheck="true">${editorHtml}</article>
      </div>` : ""}
      <aside class="paper-doc-tools">
        <input id="docSearchInput" placeholder="Im Dokument suchen">
        <small class="doc-search-count" id="docSearchCount">0 Treffer</small>
        ${canEdit ? `<button class="mode-toggle-btn" id="toggleDocMode" type="button">Bearbeitermodus</button><button class="blue-btn hidden" id="saveDocFromEditor" type="button">Speichern</button>` : ""}
        <button class="ghost-btn" id="toggleDocChanges" type="button">Changelog (${changes.length})</button>
      </aside>
    </div>
  `, (modal) => {
    modal.classList.add("wide-doc-modal");
    let editMode = false;
    let savedDocSelection = null;
    const editor = () => modal.querySelector("#paperDocEditor");
    const initial = sanitizeInformationHtml(editor()?.innerHTML || editorHtml);
    const saveDocSelection = () => {
      const selection = window.getSelection();
      const currentEditor = editor();
      if (!selection || !selection.rangeCount || !currentEditor?.contains(selection.anchorNode)) return;
      savedDocSelection = selection.getRangeAt(0).cloneRange();
    };
    const restoreDocSelection = () => {
      const currentEditor = editor();
      if (!currentEditor) return;
      currentEditor.focus();
      if (!savedDocSelection) return;
      const selection = window.getSelection();
      selection.removeAllRanges();
      selection.addRange(savedDocSelection);
    };
    const setMode = (next) => {
      editMode = next;
      modal.querySelector("#docReadMode")?.classList.toggle("hidden", editMode);
      modal.querySelector("#docEditMode")?.classList.toggle("hidden", !editMode);
      modal.querySelector("#saveDocFromEditor")?.classList.toggle("hidden", !editMode);
      const btn = modal.querySelector("#toggleDocMode");
      if (btn) btn.textContent = editMode ? "Lesemodus" : "Bearbeitermodus";
      btn?.classList.toggle("active", editMode);
      if (editMode) window.setTimeout(() => editor()?.focus(), 20);
    };
    const runCommand = (commandSpec) => {
      const [command, value] = commandSpec.split(":");
      const currentEditor = editor();
      if (!currentEditor) return;
      restoreDocSelection();
      if (command === "createLink") {
        const url = window.prompt("Link einfügen", "https://");
        if (!url) return;
        document.execCommand(command, false, url);
        currentEditor.querySelectorAll("a").forEach((link) => {
          link.target = "_blank";
          link.rel = "noreferrer";
        });
        return;
      }
      document.execCommand(command, false, value || null);
    };
    const insertDocCard = (type = "info") => {
      const currentEditor = editor();
      if (!currentEditor) return;
      const warning = type === "warning";
      restoreDocSelection();
      const cardHtml = `
        <section class="doc-section-card ${warning ? "warning" : ""}">
          <h2>${warning ? "Wichtiger Hinweis" : "Neue Kachel"}</h2>
          <p>${warning ? "Hinweis eintragen..." : "Inhalt eintragen..."}</p>
        </section>
        <p><br></p>
      `;
      document.execCommand("insertHTML", false, cardHtml);
      currentEditor.focus();
      saveDocSelection();
    };
    const applyAutoFormat = () => {
      const currentEditor = editor();
      if (!currentEditor) return;
      const source = currentEditor.innerText || currentEditor.textContent || "";
      currentEditor.innerHTML = informationTextToHtml(source);
      currentEditor.focus();
      saveDocSelection();
    };
    modal.querySelector("#toggleDocMode")?.addEventListener("click", () => setMode(!editMode));
    modal.querySelector("#toggleDocChanges")?.addEventListener("click", () => openInformationDocChangelog(doc.id, focusChangeId));
    modal.querySelector("#saveDocFromEditor")?.addEventListener("click", () => {
      const title = modal.querySelector("#paperDocTitle")?.value.trim() || doc.title;
      const current = sanitizeInformationHtml(editor()?.innerHTML || "");
      if (current === initial && title === doc.title) {
        showNotify("Keine Änderungen vorhanden.", "info");
        return;
      }
      openInformationDocCloseConfirm(doc, title, initial, current);
    });
    if (focusChangeId) window.setTimeout(() => openInformationDocChangelog(doc.id, focusChangeId), 50);
    const x = modal.querySelector(".modal-x");
    if (x) {
      const clone = x.cloneNode(true);
      x.replaceWith(clone);
      clone.addEventListener("click", closeModal);
    }
    let searchIndex = -1;
    const findEditorMatches = (term) => {
      const currentEditor = editor();
      if (!currentEditor || !term) return [];
      const matches = [];
      const needle = term.toLowerCase();
      const walker = document.createTreeWalker(currentEditor, NodeFilter.SHOW_TEXT, {
        acceptNode(node) {
          return node.nodeValue?.trim() ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_REJECT;
        }
      });
      while (walker.nextNode()) {
        const node = walker.currentNode;
        const text = node.nodeValue || "";
        let start = text.toLowerCase().indexOf(needle);
        while (start >= 0) {
          const range = document.createRange();
          range.setStart(node, start);
          range.setEnd(node, start + term.length);
          matches.push(range);
          start = text.toLowerCase().indexOf(needle, start + term.length);
        }
      }
      return matches;
    };
    const updateDocSearch = (term, move = 0) => {
      const page = modal.querySelector("#paperDocPage");
      const counter = modal.querySelector("#docSearchCount");
      if (!page) return;
      if (editMode) {
        const matches = findEditorMatches(term);
        if (!matches.length) {
          searchIndex = -1;
          if (counter) counter.textContent = term ? "0 Treffer" : "0 Treffer";
          return;
        }
        if (!move) {
          searchIndex = -1;
          if (counter) counter.textContent = `${matches.length} Treffer`;
          return;
        }
        searchIndex = move ? (searchIndex + move + matches.length) % matches.length : 0;
        const selection = window.getSelection();
        selection.removeAllRanges();
        selection.addRange(matches[searchIndex]);
        const target = matches[searchIndex].commonAncestorContainer.parentElement || editor();
        target?.scrollIntoView({ block: "center", behavior: "smooth" });
        saveDocSelection();
        if (counter) counter.textContent = `${searchIndex + 1} / ${matches.length} Treffer`;
        return;
      }
      page.innerHTML = formatInformationDocText(readBody, term);
      const marks = [...page.querySelectorAll(".doc-search-mark")];
      if (!marks.length) {
        searchIndex = -1;
        if (counter) counter.textContent = term ? "0 Treffer" : "0 Treffer";
        return;
      }
      searchIndex = move ? (searchIndex + move + marks.length) % marks.length : 0;
      marks.forEach((mark) => mark.classList.remove("active-search-mark"));
      marks[searchIndex].classList.add("active-search-mark");
      marks[searchIndex].scrollIntoView({ block: "center", behavior: "smooth" });
      if (counter) counter.textContent = `${searchIndex + 1} / ${marks.length} Treffer`;
    };
    modal.querySelector("#docSearchInput")?.addEventListener("input", (event) => {
      const term = event.target.value.trim();
      updateDocSearch(term, 0);
    });
    modal.querySelector("#docSearchInput")?.addEventListener("keydown", (event) => {
      if (event.key !== "Enter") return;
      event.preventDefault();
      const term = event.target.value.trim();
      if (term) updateDocSearch(term, event.shiftKey ? -1 : 1);
    });
    modal.querySelectorAll("[data-wysiwyg]").forEach((button) => {
      button.addEventListener("mousedown", (event) => event.preventDefault());
      button.addEventListener("click", () => runCommand(button.dataset.wysiwyg));
    });
    modal.querySelectorAll("[data-insert-card]").forEach((button) => {
      button.addEventListener("mousedown", (event) => event.preventDefault());
      button.addEventListener("click", () => insertDocCard(button.dataset.insertCard));
    });
    modal.querySelector("#docBlockStyle")?.addEventListener("change", (event) => runCommand(`formatBlock:${event.target.value}`));
    modal.querySelector("#docFontSize")?.addEventListener("change", (event) => {
      if (event.target.value) runCommand(`fontSize:${event.target.value}`);
      event.target.value = "";
    });
    modal.querySelector("#docTextColor")?.addEventListener("input", (event) => runCommand(`foreColor:${event.target.value}`));
    modal.querySelector("#docHighlightColor")?.addEventListener("input", (event) => runCommand(`hiliteColor:${event.target.value}`));
    modal.querySelector("#autoFormatDoc")?.addEventListener("mousedown", (event) => event.preventDefault());
    modal.querySelector("#autoFormatDoc")?.addEventListener("click", applyAutoFormat);
    editor()?.addEventListener("paste", (event) => {
      const plain = event.clipboardData?.getData("text/plain") || "";
      if (!plain || !/(^#\s|^§\d|━{5,}|^- |\n#\s|\n§\d|\n- )/m.test(plain)) return;
      event.preventDefault();
      document.execCommand("insertHTML", false, informationTextToHtml(plain));
    });
    editor()?.addEventListener("keydown", (event) => {
      if (!(event.ctrlKey || event.metaKey) || event.key.toLowerCase() !== "b") return;
      event.preventDefault();
      document.execCommand("bold", false, null);
      saveDocSelection();
    });
    editor()?.addEventListener("keyup", saveDocSelection);
    editor()?.addEventListener("mouseup", saveDocSelection);
    editor()?.addEventListener("input", saveDocSelection);
  });
}

async function saveInformationPatchSilent(patch) {
  const data = await api("/api/information", {
    method: "PATCH",
    silent: true,
    body: JSON.stringify({
      informationText: state.settings.informationText,
      applicationStatus: state.settings.applicationStatus,
      informationRightsText: state.settings.informationRightsText || "",
      informationLinks: state.settings.informationLinks || [],
      informationDocs: state.settings.informationDocs || [],
      informationDocChanges: state.settings.informationDocChanges || [],
      informationPermits: state.settings.informationPermits || [],
      informationFactions: state.settings.informationFactions || [],
      ...patch
    })
  });
  state.settings = data.settings || state.settings;
}

async function markInformationChangeRead(changeId) {
  const myId = state.currentUser?.id || "";
  const changes = (state.settings.informationDocChanges || []).map((change) => change.id === changeId
    ? { ...change, acknowledgedBy: Array.from(new Set([...(change.acknowledgedBy || []), myId])) }
    : change);
  await saveInformationPatchSilent({ informationDocChanges: changes });
}

async function deleteMailboxMessage(changeId) {
  const changes = (state.settings.informationDocChanges || []).filter((change) => change.id !== changeId);
  await saveInformationPatchSilent({ informationDocChanges: changes });
}

function renderPostfach() {
  const unread = unreadMailboxItems();
  const rows = state.settings.informationDocChanges || [];
  content.innerHTML = `
    <section class="panel mailbox-page">
      <div class="panel-header"><div><h3>Postfach</h3><p class="muted">${unread.length} ungelesene Nachricht${unread.length === 1 ? "" : "en"}</p></div></div>
      <div class="mailbox-list">
        ${rows.map((change) => {
          const read = !unread.some((item) => item.id === change.id);
          return `
            <article class="mailbox-row ${read ? "read" : "unread"}">
              <div class="mailbox-main">
                <strong>${escapeHtml(change.title || "Vorschrift")} wurde geändert</strong>
                <p>Es gibt eine neue Änderung bei ${escapeHtml(change.title || "einer Vorschrift")}.</p>
              </div>
              <div class="button-row">
                <button class="blue-btn open-mail-doc" data-doc-id="${escapeHtml(change.docId)}" data-change-id="${escapeHtml(change.id)}">Änderung öffnen</button>
                ${read ? "" : `<button class="ghost-btn mark-mail-read" data-change-id="${escapeHtml(change.id)}">Als gelesen markieren</button>`}
                <button class="mini-icon danger delete-mail-message" data-change-id="${escapeHtml(change.id)}" title="Nachricht löschen">${actionIcon("delete")}</button>
              </div>
              <footer>${escapeHtml(change.author || "-")} · ${formatDateTime(change.createdAt)}</footer>
            </article>
          `;
        }).join("") || `<p class="muted">Keine Nachrichten vorhanden.</p>`}
      </div>
    </section>
  `;
  document.querySelectorAll(".open-mail-doc").forEach((button) => button.addEventListener("click", async () => {
    await markInformationChangeRead(button.dataset.changeId);
    renderNavigation();
    openInformationDocView(button.dataset.docId, null, button.dataset.changeId);
  }));
  document.querySelectorAll(".mark-mail-read").forEach((button) => button.addEventListener("click", async () => {
    await markInformationChangeRead(button.dataset.changeId);
    renderApp();
  }));
  document.querySelectorAll(".delete-mail-message").forEach((button) => button.addEventListener("click", async () => {
    await deleteMailboxMessage(button.dataset.changeId);
    renderApp();
  }));
}

function renderExamModuleStart(exam, candidate) {
  const nextModule = exam.kind === "est" ? estMainModules(exam).find((module) => module.status !== "Abgeschlossen") : exam.modules.find((module) => module.status !== "Abgeschlossen");
  if (nextModule) {
    exam.activeMainModuleId = nextModule.id;
    exam.moduleIndex = exam.modules.findIndex((module) => module.id === nextModule.id);
  }
  const flow = exam.kind === "est"
    ? [
      ["1", "Rechtskunde + Ortskunde", "Rechtsfragen links, Ortskunde rechts parallel."],
      ["2", "10-80 Szenario", "Großes Szenariofeld mit Prüferinfos und Akte/Maßnahme."],
      ["3", "Dienstvorschriften + Fahrstrecke", "Vorschriften links, Fahrstrecke rechts mit Zeitwertung."],
      ["4", "Helistrecke", "Route und Landedächer mit Bild, Zeit und Bewertung."]
    ]
    : [];
  return `
    <section class="exam-runner-card exam-module-start-card compact-start">
      <span>Prüfung vorbereiten</span>
      <h4>${escapeHtml(candidate ? fullName(candidate) : "Unbekannter Prüfling")}</h4>
      <div class="exam-setup-row">
        <label class="exam-setup-second">2. Prüfer optional<select id="examSetupSecondExaminer"><option value=""></option>${state.users.map((user) => `<option value="${user.id}" ${exam.secondExaminerId === user.id ? "selected" : ""}>${escapeHtml(fullName(user))}</option>`).join("")}</select></label>
      </div>
      <input type="hidden" data-start-module-id="${escapeHtml(nextModule?.id || "")}">
      <div class="est-fixed-flow">
        <strong>${exam.status === "Vorbereitung" ? "Startet automatisch mit Rechtskunde" : `Nächstes Modul: ${escapeHtml(nextModule?.name || "-")}`}</strong>
        ${flow.map(([nr, title, text]) => `<span class="${nextModule?.name && title.includes(nextModule.name) ? "active" : ""}"><b>${nr}</b><i>${escapeHtml(title)}</i><small>${escapeHtml(text)}</small></span>`).join("")}
      </div>
      <p class="muted">Die Reihenfolge ist fest. Das Startfenster wird erst gespeichert, wenn die Prüfung wirklich gestartet wurde.</p>
    </section>
  `;
}

function renderExamModuleStepper(exam) {
  ensureExamModuleState(exam);
  const module = currentManagedExamModule(exam);
  return `<div class="exam-current-module-chip"><span>Aktuelles Modul</span><strong>${escapeHtml(module?.name || "-")}</strong><small>${escapeHtml(module?.status || exam.status || "-")}</small></div>`;
}

function renderCatalogQuestion(question, index, side = "main") {
  const maxPoints = Number(question.maxPoints || 1);
  const timed = maxPoints > 1 || Number(question.targetSeconds || 0) > 0;
  const scoreValues = timed ? Array.from({ length: Math.floor(maxPoints) + 1 }, (_, value) => value) : scoreOptionsForQuestion(question, side === "location" || question.type === "location");
  const scoreClass = (value) => `score-select score-${String(value || 0).replace(".", "-")}`;
  const scorePanel = (html) => `<div class="question-score-row"><span>Bewertung</span>${html}</div>`;
  if (side === "location" || question.type === "location") {
    const actualPoints = timedQuestionPoints(question);
    return `
      <article class="exam-catalog-question score-left-question location-question" data-question-id="${escapeHtml(question.id)}">
        ${scorePanel(timed
          ? `<strong class="auto-time-score">${String(actualPoints).replace(".", ",")} / ${escapeHtml(maxPoints)}</strong><input type="hidden" data-exam-score="${escapeHtml(question.id)}" value="${escapeHtml(actualPoints)}">`
          : `<select class="${scoreClass(question.manualPoints)}" data-exam-score="${escapeHtml(question.id)}">${scoreValues.map((value) => `<option value="${value}" ${Number(question.manualPoints || 0) === value ? "selected" : ""}>${String(value).replace(".", ",")}</option>`).join("")}</select>`)}
        <div class="question-content-box">
          <div class="catalog-question-head"><b>${index + 1}. ${escapeHtml(question.prompt)}</b><small>${timed ? "Zeitwertung" : "Bild / Strecke"}</small></div>
          ${question.image ? `<img class="location-question-image" src="${escapeHtml(question.image)}" alt="">` : ""}
          ${timed ? `<div class="time-score-row"><label>Sollzeit<input data-autosave-exam data-exam-target="${escapeHtml(question.id)}" value="${escapeHtml(formatSecondsInput(question.targetSeconds || 0))}" placeholder="MM:SS"></label><label>Gefahrene Zeit<input data-autosave-exam data-exam-time="${escapeHtml(question.id)}" value="${escapeHtml(formatSecondsInput(question.timeSeconds || 0))}" placeholder="MM:SS"></label></div>` : ""}
          ${question.solution ? `<div class="inline-solution">${escapeHtml(question.solution)}</div>` : ""}
        </div>
      </article>
    `;
  }
  if (question.type === "choice" || question.type === "scenario") {
    const answers = question.type === "scenario" ? [] : (question.answers || normalizeChoiceAnswers(question));
    return `
      <article class="exam-catalog-question score-right-question compact-choice-question" data-question-id="${escapeHtml(question.id)}">
        <div class="question-content-box">
          <div class="catalog-question-head"><b>${index + 1}. ${escapeHtml(question.prompt)}</b><small>Max. ${escapeHtml(maxPoints)} Punkte</small></div>
          ${question.scenarioInfo ? `<div class="scenario-info-box"><strong>Szenario</strong><p>${escapeHtml(question.scenarioInfo)}</p></div>` : ""}
          ${question.fileAction ? `<div class="scenario-info-box"><strong>Akte / Maßnahme</strong><p>${escapeHtml(question.fileAction)}</p></div>` : ""}
          ${answers.length ? `<div class="exam-answer-list neutral compact-answer-list">${answers.map((answer) => `<label class="exam-check compact-answer-row"><span>${escapeHtml(answer)}</span><input data-autosave-exam type="checkbox" name="answerOption_${escapeHtml(question.id)}" value="${escapeHtml(answer)}" ${question.selectedAnswers?.includes(answer) ? "checked" : ""}></label>`).join("")}</div>` : ""}
          <label>Antwort / Notizen des Prüflings<textarea data-autosave-exam data-exam-answer="${escapeHtml(question.id)}" placeholder="Antwort oder Ablauf mitschreiben">${escapeHtml(question.traineeAnswer || "")}</textarea></label>
          ${question.solution ? `<div class="inline-solution">Musterlösung: ${escapeHtml(question.solution)}</div>` : ""}
        </div>
        ${scorePanel(`<select class="${scoreClass(question.manualPoints)}" data-exam-score="${escapeHtml(question.id)}">${scoreValues.map((value) => `<option value="${value}" ${Number(question.manualPoints || 0) === value ? "selected" : ""}>${String(value).replace(".", ",")}</option>`).join("")}</select>`)}
      </article>
    `;
  }
  return `
    <article class="exam-catalog-question score-right-question" data-question-id="${escapeHtml(question.id)}">
      <div class="question-content-box">
        <div class="catalog-question-head"><b>${index + 1}. ${escapeHtml(question.prompt)}</b><small>Max. ${escapeHtml(maxPoints)} Punkte</small></div>
        <label>Antwort des Prüflings<textarea data-autosave-exam data-exam-answer="${escapeHtml(question.id)}" placeholder="Antwort mitschreiben">${escapeHtml(question.traineeAnswer || "")}</textarea></label>
        ${question.solution ? `<div class="inline-solution">Musterlösung: ${escapeHtml(question.solution)}</div>` : ""}
      </div>
      ${scorePanel(`<select class="${scoreClass(question.manualPoints)}" data-exam-score="${escapeHtml(question.id)}">${scoreValues.map((value) => `<option value="${value}" ${Number(question.manualPoints || 0) === value ? "selected" : ""}>${String(value).replace(".", ",")}</option>`).join("")}</select>`)}
    </article>
  `;
}

function renderDepartmentLeadershipPanel(department) {
  const searchValue = localStorage.getItem(`lspd_leadership_search_${department.id}`) || "";
  const selectedRange = localStorage.getItem(`lspd_leadership_range_${department.id}`) || "Gesamt";
  const searchTerm = searchValue.trim().toLowerCase();
  const members = department.members.filter((member) => {
    const haystack = `${fullName(member.user)} ${member.position} ${rankLabel(member.user.rank)} ${member.user.dn || ""}`.toLowerCase();
    return !searchTerm || haystack.includes(searchTerm);
  });
  return `
    <div class="panel department-overview-content">
      <div class="panel-header"><h3>Leitung</h3><span class="muted">${cleanText("Interne Mitglieder\u00fcbersicht")}</span></div>
      <div class="leadership-toolbar">
        <input id="leadershipSearch" value="${escapeHtml(searchValue)}" placeholder="Name, DN, Position oder Rang suchen">
        <label>Zeitraum
          <select id="leadershipRange">
            ${["Heute", "Woche", "Monat", "Gesamt"].map((range) => `<option ${selectedRange === range ? "selected" : ""}>${range}</option>`).join("")}
          </select>
        </label>
      </div>
      <div class="leadership-member-list">
        ${members.length ? members.map((member) => renderLeadershipMemberCard(department, member, selectedRange)).join("") : `<p class="muted">Keine Mitglieder gefunden.</p>`}
      </div>
    </div>
    ${isTrainingDepartmentSheet(department) ? renderTrainingManagementPanels() : ""}
  `;
}

function dutyRangeStarts() {
  const now = new Date();
  const dayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const weekStart = new Date(now);
  weekStart.setDate(now.getDate() - 7);
  const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
  return { Heute: dayStart, Woche: weekStart, Monat: monthStart, Gesamt: null };
}

function dutyRangeSumForUser(userId, range) {
  return dutySumForUser(userId, dutyRangeStarts()[range] || null);
}

function renderLeadershipMemberCard(department, member, selectedRange = "Gesamt") {
  const notes = (department.memberNotes || []).filter((note) => note.userId === member.userId).sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
  const ranges = ["Heute", "Woche", "Monat", "Gesamt"];
  return `
    <article class="leadership-member-card">
      <div class="leadership-member-head">
        <span>${avatarMarkup(member.user, "sm")}<strong>${escapeHtml(fullName(member.user))}</strong></span>
        <button class="blue-btn dept-member-note-add" data-user-id="${escapeHtml(member.userId)}">+ Interne Notiz</button>
      </div>
      <div class="leadership-facts">
        <span><b>Position</b>${escapeHtml(member.position)}</span>
        <span><b>In Abteilung seit</b>${formatDate(member.joinedAt)}</span>
        <span><b>Aktuelle Rolle seit</b>${formatDate(member.positionSince || member.joinedAt)}</span>
      </div>
      <div class="leadership-hours">
        ${ranges.map((range) => `<span class="${selectedRange === range ? "active" : ""}"><b>${range}</b>${formatDuration(dutyRangeSumForUser(member.userId, range))}</span>`).join("")}
      </div>
      <div class="leadership-notes">
        ${notes.length ? notes.map((note) => `<div><p>${escapeHtml(note.text)}</p><small>${escapeHtml(note.authorName || "-")} \u00b7 ${formatDate(note.createdAt)}</small></div>`).join("") : `<p class="muted">Keine internen Notizen.</p>`}
      </div>
    </article>
  `;
}

function renderDepartmentMemberTable(department) {
  return `
    <div class="table-wrap">
      <table>
        <thead><tr><th>Mitglied</th><th>Position</th><th>Rang</th></tr></thead>
        <tbody>
          ${department.members.map((member) => `
            <tr>
              <td><span class="member-name truncate"><span class="online-dot ${member.isOnDuty ? "online" : ""}"></span>${avatarMarkup(member.user, "sm")}<span>${escapeHtml(fullName(member.user))}</span></span></td>
              <td><span class="position-chip ${positionClass(member.position, department)}">${escapeHtml(member.position)}</span></td>
              <td><span class="department-rank-label">${escapeHtml(rankLabel(member.user.rank))}</span></td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    </div>
  `;
}

function renderDepartmentNote(department, note) {
  const className = note.priority.toLowerCase();
  return `
    <article class="note-card">
      <div class="note-top">
        <div class="note-title"><strong>${escapeHtml(note.title)}</strong><span class="badge dept-${className}">${escapeHtml(note.priority)}</span></div>
        ${departmentActionAllowed(department, "departmentNotes") ? `<div class="note-actions">
          <button class="mini-icon edit-dept-note" data-department-id="${department.id}" data-note-id="${note.id}">${actionIcon("edit")}</button>
          <button class="mini-icon danger delete-dept-note" data-department-id="${department.id}" data-note-id="${note.id}">${actionIcon("delete")}</button>
        </div>` : ""}
      </div>
      <p>${escapeHtml(note.text)}</p>
      <small class="muted">${escapeHtml(note.authorName)} · ${formatDate(note.createdAt)}</small>
    </article>
  `;
}

function positionClass(position, department = null) {
  return positionColorFor(department, position);
}

function renderProfileTrainingPanel(user) {
  const groupTitles = ["Grundausbildung", "Führung / EL", "Spezialisierungen"];
  const renderTrainingTile = (training) => {
    const done = Boolean(user.trainings?.[training]);
    const meta = user.trainingMeta?.[training] || {};
    const metaText = meta.completedAt
      ? `${formatDateTime(meta.completedAt)} · ${escapeHtml(meta.completedBy || "Unbekannt")}`
      : "Vor Systemumstellung";
    return `
      <div class="profile-training-row ${done ? "done" : "open"}">
        <span>${escapeHtml(training)}</span>
        <b>${done ? "Abgeschlossen" : "Offen"}${done ? `<small>${metaText}</small>` : ""}</b>
      </div>
    `;
  };
  return `
    <div class="panel-header"><h3>Ausbildung</h3></div>
    <div class="profile-training-group-grid">
      ${trainingGroups.map((group, index) => `
        <section class="profile-training-group">
          <h4>${escapeHtml(groupTitles[index] || `Gruppe ${index + 1}`)}</h4>
          <div class="profile-training-grid flat-training-grid">
            ${group.map(renderTrainingTile).join("")}
          </div>
        </section>
      `).join("")}
    </div>
  `;
}

function renderProfile() {
  const user = state.currentUser;
  const profileTabs = ["Ausbildung", "Dienstzeiten", "Abmeldung", "Anmeldung Prüfung"];
  const myHistory = (state.dutyHistory || []).filter((entry) => entry.userId === user.id).sort((a, b) => new Date(b.startedAt) - new Date(a.startedAt));
  const trainingDone = trainings.filter((training) => Boolean(user.trainings?.[training])).length;
  const trainingTotal = trainings.length || 1;
  const trainingPercent = Math.round((trainingDone / trainingTotal) * 100);
  const sumDuty = (status = null) => myHistory
    .filter((entry) => !status || entry.status === status || (status === "Innendienst" && entry.status === "Admin Dienst"))
    .reduce((sum, entry) => sum + durationMs(entry), 0);
  content.innerHTML = `
    <section class="panel profile-hero">
      ${avatarMarkup(user, "xl")}
      <div class="profile-main">
        <strong>${escapeHtml(fullName(user))}</strong>
        <div class="profile-badges">
          <span class="rank-pill">${escapeHtml(rankLabel(user.rank))}</span>
          ${roleBadges(user)}
        </div>
        <div class="profile-inline-facts">
          <span><b>Dienstnummer</b>${escapeHtml(user.dn)}</span>
          <span><b>Telefon</b>${escapeHtml(user.phone)}</span>
          <span><b>Beitritt Datum</b>${formatDate(user.joinedAt)}</span>
        </div>
      </div>
      <div class="profile-actions">
        <button class="orange-btn action-btn" id="openPasswordModal">${iconSvg("IT")} Passwort ändern</button>
        <button class="blue-btn action-btn" id="avatarPickBtn">${iconSvg("Profil")} Avatar ändern</button>
        <input id="avatarFileInput" class="hidden" type="file" accept="image/*">
      </div>
    </section>
    <section class="panel profile-discord-card">
      <div>
        <h3>Discord Sync</h3>
        <p class="muted">${user.discordId ? `Verknüpft mit Discord ID ${escapeHtml(user.discordId)}${user.discordName ? ` (${escapeHtml(user.discordName)})` : ""}.` : "Noch kein Discord Account verknüpft. Nach der Verknüpfung kann Discord Login und Rollen-Sync genutzt werden."}</p>
      </div>
      <button class="discord-login-btn" id="profileDiscordLinkSecondary" type="button">${user.discordId ? "Discord Verbindung erneuern" : "Discord jetzt verknüpfen"}</button>
    </section>
    <section class="grid-4 profile-stat-grid">
      <div class="stat-card progress-stat">
        <span>Ausbildungsfortschritt</span>
        <strong>${trainingPercent}%</strong>
        <div class="progress-bar"><i style="width: ${trainingPercent}%"></i></div>
        <small>${trainingDone} von ${trainingTotal} abgeschlossen</small>
      </div>
      <div class="stat-card"><span>Dienststunden</span><strong>${formatDuration(sumDuty())}</strong><small>Alle Dienste</small></div>
      <div class="stat-card split-stat">
        <span>Außendienst</span>
        <div class="service-split">
          <span><b>Normal</b>${formatDuration(sumDuty("Außendienst"))}</span>
          <span><b>Undercover</b>${formatDuration(sumDuty("Undercover Dienst"))}</span>
        </div>
      </div>
      <div class="stat-card"><span>Innendienst</span><strong>${formatDuration(sumDuty("Innendienst"))}</strong><small>Büro & Verwaltung</small></div>
    </section>
    <section class="tabs-row profile-tabs">
      ${profileTabs.map((tab) => `<button class="${state.profileTab === tab ? "tab-active" : ""}" data-profile-tab="${escapeHtml(tab)}">${escapeHtml(tab)}</button>`).join("")}
    </section>
    <section class="panel">
      ${state.profileTab === "Ausbildung" ? renderProfileTrainingPanel(user) : state.profileTab === "Dienstzeiten" ? `
        <div class="panel-header"><h3>Dienstzeiten</h3><span class="muted">${myHistory.length} Einträge</span></div>
        <div class="table-wrap">
          <table>
            <thead><tr><th>Dienstbeginn</th><th>Dienstende</th><th>Diensttyp</th><th>Dauer</th><th>Status</th></tr></thead>
            <tbody>
              ${myHistory.map((entry) => `
                <tr>
                  <td>${formatDateTime(entry.startedAt)}</td>
                  <td>${entry.endedAt ? formatDateTime(entry.endedAt) : "Läuft noch"}</td>
                  <td>${escapeHtml(entry.status)}</td>
                  <td>${formatDuration(durationMs(entry))}</td>
                  <td>${entry.endedAt ? "Beendet" : "Aktiv"}</td>
                </tr>
              `).join("") || `<tr><td colspan="5" class="muted">Noch keine Dienstzeiten.</td></tr>`}
            </tbody>
          </table>
        </div>
      ` : `<div class="template-page"><h3>${escapeHtml(state.profileTab)}</h3><p class="muted">Dieser Bereich kann später erweitert werden.</p></div>`}
    </section>
  `;

  document.querySelectorAll("[data-profile-tab]").forEach((button) => {
    button.addEventListener("click", () => {
      state.profileTab = button.dataset.profileTab;
      localStorage.setItem("lspd_profile_tab", state.profileTab);
      renderProfile();
    });
  });
  $("#openPasswordModal").addEventListener("click", openPasswordModal);
  $("#profileDiscordLinkSecondary")?.addEventListener("click", () => startDiscordOAuth("link"));
  $("#avatarPickBtn").addEventListener("click", () => $("#avatarFileInput").click());
  $("#avatarFileInput").addEventListener("change", uploadAvatarFile);
}

function openPasswordModal() {
  openModal(`
    <h3>Passwort ändern</h3>
    <label>Altes Passwort<input type="password" id="oldPassword" required></label>
    <label>Neues Passwort<input type="password" id="newPassword" required></label>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions">
      <button class="ghost-btn" data-close>Abbrechen</button>
      <button class="orange-btn" id="savePassword">Speichern</button>
    </div>
  `, (modal) => {
    modal.querySelector("#savePassword").addEventListener("click", async () => {
      try {
        await api("/api/profile/password", { method: "PATCH", body: JSON.stringify({ oldPassword: $("#oldPassword").value, newPassword: $("#newPassword").value }) });
        closeModal();
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
  });
}

function resizeAvatarFile(file) {
  return new Promise((resolve, reject) => {
    if (!file.type.startsWith("image/")) {
      reject(new Error("Bitte wähle eine Bilddatei aus."));
      return;
    }
    const image = new Image();
    const objectUrl = URL.createObjectURL(file);
    image.onload = () => {
      URL.revokeObjectURL(objectUrl);
      const maxSize = 512;
      const scale = Math.min(1, maxSize / Math.max(image.width, image.height));
      const width = Math.max(1, Math.round(image.width * scale));
      const height = Math.max(1, Math.round(image.height * scale));
      const canvas = document.createElement("canvas");
      canvas.width = width;
      canvas.height = height;
      const ctx = canvas.getContext("2d");
      ctx.drawImage(image, 0, 0, width, height);
      resolve(canvas.toDataURL("image/jpeg", 0.86));
    };
    image.onerror = () => {
      URL.revokeObjectURL(objectUrl);
      reject(new Error("Das Bild konnte nicht gelesen werden."));
    };
    image.src = objectUrl;
  });
}

async function uploadAvatarFile(event) {
  const file = event.target.files?.[0];
  if (!file) return;
  try {
    openAvatarCropModal(file);
  } catch (error) {
    openModal(`<h3>Avatar konnte nicht gespeichert werden</h3><p class="form-error">${escapeHtml(error.message)}</p><div class="modal-actions"><button class="ghost-btn" data-close>Schließen</button></div>`);
  } finally {
    event.target.value = "";
  }
}

async function saveAvatarUrl(avatarUrl) {
  const data = await api("/api/profile/avatar", { method: "PATCH", body: JSON.stringify({ avatarUrl }) });
  state.currentUser = data.user;
  renderNavigation();
  renderProfile();
}

function openAvatarCropModal(file) {
  if (!file.type.startsWith("image/")) {
    openModal(`<h3>Avatar konnte nicht gelesen werden</h3><p class="form-error">Bitte wähle eine Bilddatei aus.</p><div class="modal-actions"><button class="ghost-btn" data-close>Schließen</button></div>`);
    return;
  }
  const objectUrl = URL.createObjectURL(file);
  openModal(`
    <h3>Avatar anpassen</h3>
    <div class="avatar-crop-layout">
      <div class="avatar-crop-frame" id="avatarCropFrame"><img id="avatarCropImage" src="${escapeHtml(objectUrl)}" alt="Avatar Vorschau" draggable="false"></div>
      <p class="muted">Bild direkt ziehen. Mit dem Mausrad zoomst du rein oder raus.</p>
    </div>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions">
      <button class="ghost-btn" data-close>Abbrechen</button>
      <button class="blue-btn" id="saveCroppedAvatar">Avatar speichern</button>
    </div>
  `, (modal) => {
    const image = modal.querySelector("#avatarCropImage");
    const frame = modal.querySelector("#avatarCropFrame");
    const crop = { zoom: 0.78, x: 0, y: 0, dragging: false, startX: 0, startY: 0, originX: 0, originY: 0 };
    const syncPreview = () => {
      image.style.transform = `translate(${crop.x}px, ${crop.y}px) scale(${crop.zoom})`;
    };
    frame.addEventListener("wheel", (event) => {
      event.preventDefault();
      const nextZoom = crop.zoom + (event.deltaY < 0 ? 0.08 : -0.08);
      crop.zoom = Math.min(3, Math.max(0.45, nextZoom));
      syncPreview();
    }, { passive: false });
    frame.addEventListener("pointerdown", (event) => {
      crop.dragging = true;
      crop.startX = event.clientX;
      crop.startY = event.clientY;
      crop.originX = crop.x;
      crop.originY = crop.y;
      frame.setPointerCapture(event.pointerId);
    });
    frame.addEventListener("pointermove", (event) => {
      if (!crop.dragging) return;
      crop.x = crop.originX + event.clientX - crop.startX;
      crop.y = crop.originY + event.clientY - crop.startY;
      syncPreview();
    });
    frame.addEventListener("pointerup", () => {
      crop.dragging = false;
    });
    syncPreview();
    modal.querySelector("#saveCroppedAvatar").addEventListener("click", async () => {
      try {
        const avatarUrl = cropAvatarImage(image, crop.zoom, crop.x, crop.y);
        URL.revokeObjectURL(objectUrl);
        await saveAvatarUrl(avatarUrl);
        closeModal();
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
  });
}

function cropAvatarImage(image, zoom, offsetX, offsetY) {
  if (!image.naturalWidth || !image.naturalHeight) throw new Error("Bild ist noch nicht geladen.");
  const size = 512;
  const canvas = document.createElement("canvas");
  canvas.width = size;
  canvas.height = size;
  const ctx = canvas.getContext("2d");
  ctx.fillStyle = "#0e1728";
  ctx.fillRect(0, 0, size, size);
  const baseScale = Math.max(size / image.naturalWidth, size / image.naturalHeight) * zoom;
  const drawWidth = image.naturalWidth * baseScale;
  const drawHeight = image.naturalHeight * baseScale;
  ctx.drawImage(image, (size - drawWidth) / 2 + offsetX * 2, (size - drawHeight) / 2 + offsetY * 2, drawWidth, drawHeight);
  return canvas.toDataURL("image/jpeg", 0.88);
}

function renderTemplate(page) {
  content.innerHTML = `
    <section class="panel template-page">
      <h3>${escapeHtml(navLabel(page))}</h3>
      <p class="muted">Template-Seite. Die Funktionen können hier als nächstes erweitert werden.</p>
    </section>
  `;
}

function seizureItems() {
  return [...(state.settings?.seizures || [])].sort((a, b) => String(b.createdAt || "").localeCompare(String(a.createdAt || "")));
}

function formatMoney(value) {
  const amount = Number(value || 0);
  return amount > 0 ? `${amount.toLocaleString("de-DE")}$` : "-";
}

function seizureEvidenceLinks(item) {
  const links = Array.isArray(item.evidenceLinks) ? item.evidenceLinks : [];
  const legacy = [item.weapons, item.drugs, item.other].map((value) => String(value || "").trim()).filter(Boolean);
  return [...links.map((value) => String(value || "").trim()).filter(Boolean), ...legacy];
}

function renderEvidenceLinks(item) {
  const links = seizureEvidenceLinks(item);
  if (!links.length) return "-";
  return `<div class="evidence-link-list">${links.map((link, index) => {
    const isUrl = /^https?:\/\//i.test(link);
    const isPrnt = /^https?:\/\/(?:www\.)?prnt\.sc\//i.test(link);
    const isGyazo = /^https?:\/\/(?:www\.)?gyazo\.com\//i.test(link);
    const isImgur = /^https?:\/\/(?:www\.)?imgur\.com\/(?:a|gallery)\//i.test(link);
    const previewSrc = isPrnt || isGyazo || isImgur ? `/api/evidence-preview?url=${encodeURIComponent(link)}` : escapeHtml(link);
    const isUploadedImage = /^data:image\/(?:png|jpe?g|webp|gif);base64,/i.test(link);
    const loading = index < 2 ? "eager" : "lazy";
    const priority = index < 2 ? ' fetchpriority="high"' : "";
    const imageAttrs = `width="152" height="86" loading="${loading}" decoding="async"${priority} onload="this.closest('.evidence-preview-card')?.classList.add('is-loaded')"`;
    if (isUploadedImage) {
      return `<div class="evidence-preview-card uploaded-preview"><button class="evidence-thumb-link evidence-preview-open" type="button" data-link="${escapeHtml(link)}"><img src="${escapeHtml(link)}" alt="Beweis ${index + 1}" ${imageAttrs}></button><span class="evidence-text-link">Hochgeladenes Bild</span></div>`;
    }
    return isUrl
      ? `<div class="evidence-preview-card ${isPrnt || isGyazo || isImgur ? "prnt-preview" : ""}"><button class="evidence-thumb-link evidence-preview-open" type="button" data-link="${escapeHtml(link)}"><img src="${previewSrc}" alt="Beweis ${index + 1}" ${imageAttrs} onerror="this.closest('.evidence-preview-card').classList.add('no-preview')"><span class="prnt-fallback">${isGyazo ? "GYAZO" : isPrnt ? "PRNT.SC" : isImgur ? "IMGUR" : "VORSCHAU"}</span></button><a class="evidence-text-link" href="${escapeHtml(link)}" target="_blank" rel="noopener">${escapeHtml(link)}</a></div>`
      : `<span>${escapeHtml(link)}</span>`;
  }).join("")}</div>`;
}

function renderSeizures() {
  const search = localStorage.getItem("lspd_seizure_search") || "";
  const statRange = localStorage.getItem("lspd_seizure_stat_range") || "Gesamt";
  const visibleLimit = Math.max(15, Number(localStorage.getItem("lspd_seizure_visible_limit") || 25));
  const items = seizureItems();
  const statStart = rangeStart(statRange);
  const statItems = statStart ? items.filter((item) => new Date(item.createdAt).getTime() >= statStart.getTime()) : items;
  const filtered = items.filter((item) => {
    const haystack = [
      item.suspect,
      item.location,
      seizureEvidenceLinks(item).join(" "),
      item.witness,
      item.sourceType,
      item.vehicleId,
      item.officerName
    ].join(" ").toLowerCase();
    return haystack.includes(search.toLowerCase());
  });
  const totalBlackMoney = statItems.reduce((sum, item) => sum + (Number(item.blackMoney) || 0), 0);
  const totalCrates = statItems.reduce((sum, item) => sum + (Number(item.crates) || 0), 0);
  const dealerCount = statItems.filter((item) => item.sourceType === "Dealer").length;
  const camperCount = statItems.filter((item) => item.sourceType === "Camper").length;
  const canDelete = hasRole("Direktion");
  const canEditAll = hasRole("Direktion");
  const ranges = ["Heute", "Woche", "Monat", "Gesamt"];
  const visibleRows = filtered.slice(0, visibleLimit);
  const hiddenRows = Math.max(0, filtered.length - visibleRows.length);

  content.innerHTML = `
    <section class="seizure-page">
      <div class="grid-4 seizure-stats">
        <article class="stat-card"><span>Einträge</span><strong>${statItems.length}</strong><small>${escapeHtml(statRange)} erfasst</small></article>
        <article class="stat-card"><span>Schwarzgeld</span><strong>${totalBlackMoney.toLocaleString("de-DE")}$</strong><small>Gesamtmenge</small></article>
        <article class="stat-card"><span>Kisten</span><strong>${totalCrates.toLocaleString("de-DE")}</strong><small>Gesamtmenge</small></article>
        <article class="stat-card"><span>Dealer / Camper</span><strong>${dealerCount} / ${camperCount}</strong><small>Besondere Fundarten</small></article>
      </div>
      <div class="seizure-stats-head">
        <span>Zeitraum</span>
        <select id="seizureStatRange" class="compact-input seizure-range-select">
          ${ranges.map((range) => `<option value="${range}" ${statRange === range ? "selected" : ""}>${range}</option>`).join("")}
        </select>
      </div>
      <section class="panel seizure-panel">
        <div class="panel-header">
          <div><h3>${iconSvg("Beschlagnahmung")} Beschlagnahmungen (${filtered.length})</h3><p class="muted">Suche nach Tatverdächtigem, Standort, Officer oder Beweis.</p></div>
          <button class="blue-btn" id="addSeizureBtn">${iconSvg("Plus")} Neue Beschlagnahmung</button>
        </div>
        <div class="seizure-search-row">
          <input id="seizureSearch" value="${escapeHtml(search)}" placeholder="Suche nach Tatverdächtiger, Standort, Officer oder Beweis...">
          <button class="blue-btn" id="runSeizureSearch">Suchen</button>
        </div>
        <div class="table-wrap seizure-table-wrap">
          <table class="seizure-table">
            <thead>
              <tr>
                <th>Tatverdächtiger</th>
                <th>Standort</th>
                <th>Beweise</th>
                <th>Schwarzgeld</th>
                <th>Kisten</th>
                <th>Art</th>
                <th>KFZ / Kennzeichen</th>
                <th>Zeuge</th>
                <th>Mord/Totschlag</th>
                <th>Zeitstempel</th>
                <th>Erfasst von</th>
                <th>Aktionen</th>
              </tr>
            </thead>
            <tbody>
              ${visibleRows.map((item) => `
                <tr>
                  <td><strong>${escapeHtml(item.suspect || "-")}</strong></td>
                  <td class="seizure-location-cell">${escapeHtml(item.location || "-")}</td>
                  <td>${renderEvidenceLinks(item)}</td>
                  <td>${formatMoney(item.blackMoney)}</td>
                  <td>${Number(item.crates || 0) || "-"}</td>
                  <td><span class="seizure-pill ${item.sourceType === "Dealer" ? "dealer" : item.sourceType === "Camper" ? "camper" : "normal"}">${escapeHtml(item.sourceType && item.sourceType !== "Normal" ? item.sourceType : "-")}</span></td>
                  <td>${escapeHtml(item.vehicleId || "-")}</td>
                  <td>${escapeHtml(item.witness || "-")}</td>
                  <td><span class="seizure-pill ${item.murder ? "yes" : "no"}">${item.murder ? "Ja" : "Nein"}</span></td>
                  <td>${formatDateTime(item.createdAt)}</td>
                  <td>${escapeHtml(item.officerName || "-")}</td>
                  <td>${canEditAll || item.officerId === state.currentUser.id ? `<button class="mini-icon seizure-actions gear-action" data-id="${escapeHtml(item.id)}" title="Aktionen">${iconSvg("Settings")}</button>` : `<span class="muted">-</span>`}</td>
                </tr>
              `).join("") || `<tr><td colspan="12" class="empty-table">Keine Beschlagnahmungen gefunden.</td></tr>`}
            </tbody>
          </table>
        </div>
        ${hiddenRows ? `<div class="seizure-more-row"><button class="ghost-btn" id="showMoreSeizures">${iconSvg("ChevronDown")} ${hiddenRows} weitere anzeigen</button></div>` : ""}
      </section>
    </section>
  `;

  $("#addSeizureBtn")?.addEventListener("click", () => openSeizureModal());
  $("#seizureStatRange")?.addEventListener("change", (event) => {
    localStorage.setItem("lspd_seizure_stat_range", event.target.value);
    renderSeizures();
  });
  document.querySelectorAll(".seizure-actions").forEach((button) => button.addEventListener("click", () => openSeizureActionsModal(button.dataset.id)));
  document.querySelectorAll(".evidence-preview-open").forEach((button) => button.addEventListener("click", () => openEvidencePreview(button.dataset.link)));
  $("#showMoreSeizures")?.addEventListener("click", () => {
    localStorage.setItem("lspd_seizure_visible_limit", String(visibleLimit + 25));
    renderSeizures();
  });
  $("#runSeizureSearch")?.addEventListener("click", () => {
    localStorage.setItem("lspd_seizure_search", $("#seizureSearch").value);
    renderSeizures();
  });
  $("#seizureSearch")?.addEventListener("input", (event) => {
    localStorage.setItem("lspd_seizure_search", event.target.value);
  });
  $("#seizureSearch")?.addEventListener("keydown", (event) => {
    if (event.key !== "Enter") return;
    localStorage.setItem("lspd_seizure_search", event.target.value);
    renderSeizures();
  });
}

function openEvidencePreview(link) {
  const isPrnt = /^https?:\/\/(?:www\.)?prnt\.sc\//i.test(link || "");
  const isGyazo = /^https?:\/\/(?:www\.)?gyazo\.com\//i.test(link || "");
  const preview = isPrnt || isGyazo ? `/api/evidence-preview?url=${encodeURIComponent(link)}` : link;
  openModal(`
    <h3>Beweisvorschau</h3>
    <div class="evidence-popup-preview">
      <img src="${escapeHtml(preview)}" alt="Beweisvorschau">
      <p class="muted evidence-preview-fallback">Falls die Vorschau nicht lädt, öffne den Link separat.</p>
    </div>
    <a class="blue-btn evidence-popup-link" href="${escapeHtml(link)}" target="_blank" rel="noopener">Link öffnen</a>
  `, (modal) => modal.classList.add("evidence-preview-modal"));
}

function openSeizureActionsModal(id) {
  const item = seizureItems().find((entry) => entry.id === id);
  if (!item) return;
  const canEdit = hasRole("Direktion") || item.officerId === state.currentUser.id;
  const canDelete = hasRole("Direktion");
  if (!canEdit && !canDelete) return;
  openModal(`
    <h3>Beschlagnahmung Aktionen</h3>
    <p class="muted">${escapeHtml(item.suspect || "-")} · ${escapeHtml(item.location || "-")}</p>
    <div class="choice-grid">
      ${canEdit ? `<button class="choice-card" id="editSeizureAction"><strong>Bearbeiten</strong><span>Eintrag anpassen und Beweise ergänzen.</span></button>` : ""}
      ${canDelete ? `<button class="choice-card danger-choice" id="deleteSeizureAction"><strong>Löschen</strong><span>Eintrag dauerhaft entfernen.</span></button>` : ""}
    </div>
  `, (modal) => {
    modal.querySelector("#editSeizureAction")?.addEventListener("click", () => openSeizureModal(item));
    modal.querySelector("#deleteSeizureAction")?.addEventListener("click", () => openDeleteSeizureModal(id));
  });
}

function openDeleteSeizureModal(id) {
  const item = seizureItems().find((entry) => entry.id === id);
  if (!item || !hasRole("Direktion")) return;
  openModal(`
    <h3>Beschlagnahmung löschen</h3>
    <p class="muted">Eintrag von <strong>${escapeHtml(item.suspect || "-")}</strong> wirklich löschen?</p>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions">
      <button class="ghost-btn" data-close>Abbrechen</button>
      <button class="red-btn" id="confirmDeleteSeizure">Löschen</button>
    </div>
  `, (modal) => {
    modal.querySelector("#confirmDeleteSeizure").addEventListener("click", async () => {
      try {
        const data = await api(`/api/seizures/${id}`, { method: "DELETE" });
        state.settings = data.settings || { ...state.settings, seizures: (state.settings.seizures || []).filter((entry) => entry.id !== id) };
        closeModal();
        renderSeizures();
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
  });
}

function openSeizureModal(item = null) {
  const isEdit = Boolean(item?.id);
  const evidenceLinks = seizureEvidenceLinks(item || {});
  const selectedSource = ["Dealer", "Camper"].includes(item?.sourceType) ? item.sourceType : "";
  const witnessOptions = state.users
    .filter((user) => !user.terminated)
    .sort((a, b) => fullName(a).localeCompare(fullName(b), "de"))
    .map((user) => `<option value="${escapeHtml(fullName(user))}" ${item?.witness === fullName(user) ? "selected" : ""}>${escapeHtml(fullName(user))} (${escapeHtml(user.dn || "-")})</option>`)
    .join("");
  openModal(`
    <div class="seizure-modal-head"><span>${iconSvg(isEdit ? "Settings" : "Plus")}</span><div><h3>${isEdit ? "Beschlagnahmung bearbeiten" : "Neue Beschlagnahmung"}</h3><p class="muted">Pflichtfelder ausfüllen, optionale Angaben nur bei Bedarf ergänzen.</p></div></div>
    <div class="seizure-modal-grid">
      <label><span class="required-label">Tatverdächtiger <b>*</b></span><input id="seizureSuspect" value="${escapeHtml(item?.suspect || "")}" placeholder="Name des Tatverdächtigen" required></label>
      <label><span class="required-label">Standort <b>*</b></span><input id="seizureLocation" value="${escapeHtml(item?.location || "")}" placeholder="Ort der Beschlagnahmung" required></label>
      <div class="seizure-source-field full">
        <div class="seizure-source-options">
          ${["Dealer", "Camper"].map((type) => `<label><input class="seizure-source-choice" type="checkbox" value="${type}" ${selectedSource === type ? "checked" : ""}><span>${type}</span></label>`).join("")}
        </div>
      </div>
      <div class="full evidence-field">
        <div class="field-title required-label">Beweise <b>*</b></div>
        <div id="evidenceLinkList" class="evidence-input-list">
          ${(evidenceLinks.length ? evidenceLinks : [""]).map((link, index) => index === 0
            ? `<input class="evidence-link-input" value="${escapeHtml(link)}" placeholder="Screenshot-Link / Beweis-Link">`
            : `<div class="evidence-input-row"><input class="evidence-link-input" value="${escapeHtml(link)}" placeholder="Weiterer Screenshot-Link / Beweis-Link"><button class="mini-icon remove-evidence-link" type="button" title="Entfernen">X</button></div>`).join("")}
        </div>
        <div class="evidence-actions-row">
          <button class="ghost-btn evidence-add-btn" type="button" id="addEvidenceLink">${iconSvg("Plus")} Weiteren Link hinzufügen</button>
          <button class="ghost-btn evidence-add-btn" type="button" id="uploadEvidenceImage">${iconSvg("Plus")} Bild hochladen</button>
          <input id="evidenceImageUpload" class="hidden" type="file" accept="image/png,image/jpeg,image/webp,image/gif">
        </div>
      </div>
      <label>Schwarzgeld Menge<input id="seizureBlackMoney" type="number" min="0" step="1" value="${escapeHtml(item?.blackMoney || "")}" placeholder="0"></label>
      <label>Kisten Menge<input id="seizureCrates" type="number" min="0" step="1" value="${escapeHtml(item?.crates || "")}" placeholder="0"></label>
      <label>KFZ ID / Kennzeichen<input id="seizureVehicleId" value="${escapeHtml(item?.vehicleId || "")}" placeholder="Optional"></label>
      <label>Zeuge / Officer
        <select id="seizureWitness">
          <option value="" ${item?.witness ? "" : "selected"}>Officer auswählen...</option>
          ${witnessOptions}
        </select>
      </label>
      <label class="it-toggle seizure-murder-toggle">
        <input id="seizureMurder" type="checkbox" ${item?.murder ? "checked" : ""}>
        <span class="it-toggle-ui"></span>
        <span>Mord/Totschlag</span>
      </label>
    </div>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions"><button class="blue-btn full-action" id="saveSeizure">${isEdit ? "Beschlagnahmung speichern" : "Beschlagnahmung eintragen"}</button></div>
  `, (modal) => {
    modal.classList.add("seizure-modal");
    modal.querySelectorAll(".seizure-source-choice").forEach((input) => {
      input.addEventListener("change", () => {
        if (!input.checked) return;
        modal.querySelectorAll(".seizure-source-choice").forEach((other) => {
          if (other !== input) other.checked = false;
        });
      });
    });
    modal.querySelectorAll(".remove-evidence-link").forEach((button) => button.addEventListener("click", () => button.closest(".evidence-input-row")?.remove()));
    modal.querySelector("#saveSeizure").addEventListener("click", async () => {
      try {
        const evidenceLinks = [...document.querySelectorAll(".evidence-link-input")].map((input) => input.value.trim()).filter(Boolean);
        if (!evidenceLinks.length) {
          $("#modalError").textContent = "Bitte mindestens einen Beweis-Link eintragen.";
          return;
        }
        const data = await api(isEdit ? `/api/seizures/${item.id}` : "/api/seizures", {
          method: isEdit ? "PATCH" : "POST",
          body: JSON.stringify({
            suspect: $("#seizureSuspect").value,
            location: $("#seizureLocation").value,
            evidenceLinks,
            witness: $("#seizureWitness").value,
            murder: $("#seizureMurder").checked,
            blackMoney: $("#seizureBlackMoney").value,
            crates: $("#seizureCrates").value,
            vehicleId: $("#seizureVehicleId").value,
            sourceType: document.querySelector(".seizure-source-choice:checked")?.value || ""
          })
        });
        state.settings = data.settings || { ...state.settings, seizures: [data.seizure, ...(state.settings.seizures || [])] };
        closeModal();
        renderSeizures();
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
    modal.querySelector("#addEvidenceLink").addEventListener("click", () => {
      const row = document.createElement("div");
      row.className = "evidence-input-row";
      row.innerHTML = `<input class="evidence-link-input" placeholder="Weiterer Screenshot-Link / Beweis-Link"><button class="mini-icon remove-evidence-link" type="button" title="Entfernen">X</button>`;
      modal.querySelector("#evidenceLinkList").appendChild(row);
      row.querySelector(".remove-evidence-link").addEventListener("click", () => row.remove());
      row.querySelector("input").focus();
    });
    modal.querySelector("#uploadEvidenceImage")?.addEventListener("click", () => modal.querySelector("#evidenceImageUpload")?.click());
    modal.querySelector("#evidenceImageUpload")?.addEventListener("change", async (event) => {
      const file = event.target.files?.[0];
      event.target.value = "";
      if (!file) return;
      if (!["image/png", "image/jpeg", "image/webp", "image/gif"].includes(file.type)) {
        $("#modalError").textContent = "Bitte nur Foto-Dateien hochladen (PNG, JPG, WEBP oder GIF).";
        return;
      }
      const dataUrl = await new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(String(reader.result || ""));
        reader.onerror = reject;
        reader.readAsDataURL(file);
      });
      const row = document.createElement("div");
      row.className = "evidence-input-row uploaded-evidence-row";
      row.innerHTML = `<input class="evidence-link-input" value="${escapeHtml(dataUrl)}" readonly data-uploaded-image="true"><span class="uploaded-evidence-name">${escapeHtml(file.name)}</span><button class="mini-icon remove-evidence-link" type="button" title="Entfernen">X</button>`;
      modal.querySelector("#evidenceLinkList").appendChild(row);
      row.querySelector(".remove-evidence-link").addEventListener("click", () => row.remove());
    });
  });
}

function renderCalendar() {
  const year = calendarCursor.getFullYear();
  const month = calendarCursor.getMonth();
  const first = new Date(year, month, 1);
  const startOffset = (first.getDay() + 6) % 7;
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const cells = [];
  for (let i = 0; i < startOffset; i++) cells.push(null);
  for (let day = 1; day <= daysInMonth; day++) cells.push(new Date(year, month, day));
  while (cells.length % 7 !== 0) cells.push(null);
  const events = [...(state.settings.calendarEvents || [])].sort((a, b) => `${a.startDate}T${a.startTime || "00:00"}`.localeCompare(`${b.startDate}T${b.startTime || "00:00"}`));
  const monthEvents = events.filter((event) => {
    const date = new Date(`${event.startDate}T00:00`);
    return date.getFullYear() === year && date.getMonth() === month;
  });
  const selectedEvents = events.filter((event) => event.startDate === selectedCalendarDate);
  content.innerHTML = `
    <section class="calendar-layout">
      <div class="panel calendar-panel">
        <div class="calendar-head">
          <h3>${escapeHtml(monthName(calendarCursor))}</h3>
          <div class="button-row">
            <button class="ghost-btn" id="calendarToday">Heute</button>
            <button class="icon-btn calendar-prev" id="calendarPrev">${iconSvg("ChevronDown")}</button>
            <button class="icon-btn calendar-next" id="calendarNext">${iconSvg("ChevronDown")}</button>
          </div>
        </div>
        <div class="calendar-weekdays">${["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"].map((day) => `<span>${day}</span>`).join("")}</div>
        <div class="calendar-grid">
          ${cells.map((date) => {
            if (!date) return `<div class="calendar-day muted-day"></div>`;
            const iso = isoDateLocal(date);
            const dayEvents = events.filter((event) => event.startDate === iso);
            return `<button class="calendar-day ${iso === isoDateLocal(new Date()) ? "today" : ""}" data-date="${iso}">
              <strong>${date.getDate()}</strong>
              ${dayEvents.slice(0, 3).map((event) => `<span class="calendar-event-pill ${calendarColorClass(event.color)}">${escapeHtml(event.title)}</span>`).join("")}
            </button>`;
          }).join("")}
        </div>
      </div>
      <aside class="calendar-side">
        <section class="panel">
          <h3>Neuer Termin</h3>
          <button class="blue-btn calendar-create-btn" id="createCalendarEvent">${iconSvg("Plus")} Termin erstellen</button>
        </section>
        <section class="panel selected-day-panel">
          <h3>${escapeHtml(calendarDayTitle(selectedCalendarDate))}</h3>
          <div class="upcoming-list">
            ${selectedEvents.length ? selectedEvents.map((event) => `
              <article class="selected-event-card ${calendarColorClass(event.color)}">
                <button class="calendar-event-settings" data-id="${escapeHtml(event.id)}" title="Termin verwalten">⚙</button>
                <strong>${escapeHtml(event.title)}</strong>
                <span class="event-meta-line"><img class="event-meta-icon" src="/uhr.png" alt="" draggable="false">${event.allDay ? "Ganztägig" : `${escapeHtml(event.startTime)} Uhr - ${escapeHtml(event.endTime || "")} Uhr`}</span>
                <small class="event-meta-line"><img class="event-meta-icon" src="/standort.png" alt="" draggable="false">${escapeHtml(event.location || "Kein Ort")}</small>
                ${event.description ? `<p>${escapeHtml(event.description)}</p>` : ""}
                <small class="event-author">Erstellt von: ${escapeHtml(event.authorName || "-")}</small>
              </article>
            `).join("") : `<p class="muted">Keine Termine an diesem Tag.</p>`}
          </div>
        </section>
        <section class="panel">
          <h3>Anstehende Termine</h3>
          <div class="upcoming-list">
            ${monthEvents.length ? monthEvents.slice(0, 8).map((event) => `
              <article class="upcoming-event ${calendarColorClass(event.color)}">
                <strong>${escapeHtml(event.title)}</strong>
                <span class="event-meta-line"><img class="event-meta-icon" src="/uhr.png" alt="" draggable="false">${formatDate(event.startDate)} - ${event.allDay ? "Ganztägig" : `${escapeHtml(event.startTime)} Uhr`}</span>
                <small class="event-meta-line"><img class="event-meta-icon" src="/standort.png" alt="" draggable="false">${escapeHtml(event.location || "Kein Ort")}</small>
              </article>
            `).join("") : `<p class="muted">Keine Termine in diesem Monat.</p>`}
          </div>
        </section>
      </aside>
    </section>
  `;
  $("#calendarToday").addEventListener("click", () => {
    const now = new Date();
    calendarCursor = new Date(now.getFullYear(), now.getMonth(), 1);
    renderCalendar();
  });
  $("#calendarPrev").addEventListener("click", () => {
    calendarCursor = new Date(year, month - 1, 1);
    renderCalendar();
  });
  $("#calendarNext").addEventListener("click", () => {
    calendarCursor = new Date(year, month + 1, 1);
    renderCalendar();
  });
  $("#createCalendarEvent").addEventListener("click", () => openCalendarEventModal(selectedCalendarDate));
  document.querySelectorAll(".calendar-day[data-date]").forEach((button) => button.addEventListener("click", () => {
    selectedCalendarDate = button.dataset.date;
    renderCalendar();
  }));
  document.querySelectorAll(".calendar-event-settings").forEach((button) => button.addEventListener("click", () => openCalendarEventActionsModal(events.find((event) => event.id === button.dataset.id))));
}

function calendarColorClass(color = "Blau") {
  return `calendar-color-${String(color).toLowerCase().replace("ü", "ue")}`;
}

function addOneHour(time = "10:00") {
  const [hour, minute] = time.split(":").map(Number);
  return `${String((hour + 1) % 24).padStart(2, "0")}:${String(minute || 0).padStart(2, "0")}`;
}

function openCalendarEventModal(date = isoDateLocal(new Date()), event = null) {
  const isEdit = Boolean(event);
  const startDate = event?.startDate || date || isoDateLocal(new Date());
  const startTime = event?.startTime || "10:00";
  const endTime = event?.endTime || addOneHour(startTime);
  openModal(`
    <h3>${isEdit ? "Termin bearbeiten" : "Neuer Termin"}</h3>
    <p class="muted">${isEdit ? "Bearbeite den Kalender-Termin" : "Erstelle einen neuen Kalender-Termin"}</p>
    <label>Titel *<input id="calendarTitle" value="${escapeHtml(event?.title || "")}" placeholder="z.B. Training Division Meeting"></label>
    <label>Beschreibung<textarea id="calendarDescription" placeholder="Weitere Details zum Termin...">${escapeHtml(event?.description || "")}</textarea></label>
    <label class="checkbox-line">Ganztägig<input type="checkbox" id="calendarAllDay" ${event?.allDay ? "checked" : ""}></label>
    <div class="form-grid">
      <label>Startdatum *<input id="calendarStartDate" type="date" value="${escapeHtml(startDate)}"></label>
      <label>Startzeit *<input id="calendarStartTime" type="time" value="${escapeHtml(startTime)}"></label>
      <label>Enddatum<input id="calendarEndDate" type="date" value="${escapeHtml(event?.endDate || startDate)}"></label>
      <label>Endzeit<input id="calendarEndTime" type="time" value="${escapeHtml(endTime)}"></label>
    </div>
    <label>Event-Typ<select id="calendarType">${["Allgemein", "Training", "Besprechung", "Einsatz", "Prüfung"].map((item) => `<option ${event?.type === item ? "selected" : ""}>${item}</option>`).join("")}</select></label>
    <label>Farbe<select id="calendarColor">${["Blau", "Grün", "Orange", "Rot", "Lila"].map((item) => `<option ${event?.color === item ? "selected" : ""}>${item}</option>`).join("")}</select></label>
    <label>Ort<input id="calendarLocation" value="${escapeHtml(event?.location || "")}" placeholder="z.B. Mission Row - Besprechungsraum"></label>
    <label>Erinnerung (Minuten vorher)<select id="calendarReminder">${["Keine", "10 Minuten", "30 Minuten", "1 Stunde", "1 Tag"].map((item) => `<option ${(event?.reminder || "30 Minuten") === item ? "selected" : ""}>${item}</option>`).join("")}</select></label>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions"><button class="ghost-btn" data-close>Abbrechen</button><button class="blue-btn" id="saveCalendarEvent">${isEdit ? "Speichern" : "Termin erstellen"}</button></div>
  `, (modal) => {
    let endTouched = isEdit;
    modal.querySelector("#calendarEndTime").addEventListener("input", () => { endTouched = true; });
    modal.querySelector("#calendarStartTime").addEventListener("input", () => {
      if (!endTouched) $("#calendarEndTime").value = addOneHour($("#calendarStartTime").value);
    });
    modal.querySelector("#saveCalendarEvent").addEventListener("click", async () => {
      try {
        await api(isEdit ? `/api/calendar/events/${event.id}` : "/api/calendar/events", {
          method: isEdit ? "PATCH" : "POST",
          body: JSON.stringify({
            title: $("#calendarTitle").value,
            description: $("#calendarDescription").value,
            allDay: $("#calendarAllDay").checked,
            startDate: $("#calendarStartDate").value,
            startTime: $("#calendarStartTime").value,
            endDate: $("#calendarEndDate").value,
            endTime: $("#calendarEndTime").value,
            type: $("#calendarType").value,
            color: $("#calendarColor").value,
            location: $("#calendarLocation").value,
            reminder: $("#calendarReminder").value
          })
        });
        closeModal();
        await bootstrap();
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
  });
}

function openCalendarEventActionsModal(event) {
  if (!event) return;
  openModal(`
    <h3>Termin verwalten</h3>
    <p class="muted">${escapeHtml(event.title)}</p>
    <div class="action-menu-list">
      <button class="ghost-btn action-menu-btn" id="editCalendarEvent">${actionIcon("edit")} Bearbeiten</button>
      <button class="red-btn action-menu-btn" id="deleteCalendarEvent">${actionIcon("delete")} Löschen</button>
    </div>
    <div class="modal-actions"><button class="ghost-btn" data-close>Schließen</button></div>
  `, (modal) => {
    modal.querySelector("#editCalendarEvent").addEventListener("click", () => {
      closeModal();
      openCalendarEventModal(null, event);
    });
    modal.querySelector("#deleteCalendarEvent").addEventListener("click", () => {
      closeModal();
      openDeleteCalendarEventModal(event);
    });
  });
}

function openDeleteCalendarEventModal(event) {
  if (!event) return;
  openModal(`
    <h3>Termin löschen</h3>
    <p class="muted">${escapeHtml(event.title)} wirklich löschen?</p>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions"><button class="ghost-btn" data-close>Abbrechen</button><button class="red-btn" id="confirmDeleteCalendarEvent">Löschen</button></div>
  `, (modal) => {
    modal.querySelector("#confirmDeleteCalendarEvent").addEventListener("click", async () => {
      try {
        await api(`/api/calendar/events/${event.id}`, { method: "DELETE" });
        closeModal();
        await bootstrap();
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
  });
}

function setupTableFilter(selector) {
  const input = $(selector);
  if (!input) return;
  input.addEventListener("input", () => {
    const term = input.value.toLowerCase().trim();
    document.querySelectorAll(".filterable-row").forEach((row) => {
      const matches = !term || row.textContent.toLowerCase().includes(term);
      row.classList.toggle("hidden", !matches);
      row.classList.toggle("search-match", Boolean(term && matches));
    });
  });
}

function openModal(html, onReady) {
  modalRoot.innerHTML = `<div class="modal"><button class="modal-x" type="button" data-close aria-label="Schließen">×</button>${html}</div>`;
  modalRoot.classList.remove("hidden");
  document.removeEventListener("keydown", handleModalEscape, true);
  document.addEventListener("keydown", handleModalEscape, true);
  modalRoot.querySelectorAll("[data-close]").forEach((button) => button.addEventListener("click", closeModal));
  onReady?.(modalRoot.querySelector(".modal"));
}

function openConfirmModal({ title = "Löschen bestätigen", text = "Diesen Eintrag wirklich löschen?", confirmText = "Löschen", onConfirm }) {
  openModal(`
    <h3>${escapeHtml(title)}</h3>
    <p class="muted">${escapeHtml(text)}</p>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions">
      <button class="ghost-btn" data-close>Abbrechen</button>
      <button class="red-btn" id="confirmGenericDelete">${escapeHtml(confirmText)}</button>
    </div>
  `, (modal) => {
    modal.querySelector("#confirmGenericDelete").addEventListener("click", async () => {
      try {
        await onConfirm?.();
        closeModal();
      } catch (error) {
        modal.querySelector("#modalError").textContent = error.message;
      }
    });
  });
}

function closeModal() {
  const discardTrainingExamId = modalRoot.dataset.discardTrainingExamId;
  if (discardTrainingExamId) {
    try {
      const store = trainingStore();
      store.activeExams = (store.activeExams || []).filter((exam) => exam.id !== discardTrainingExamId || exam.startedAt);
      saveTrainingStore(store);
      const department = departmentByPage?.(state.page);
      if (isTrainingDepartmentSheet(department) || isHumanResourcesDepartmentSheet(department)) window.setTimeout(() => renderDepartmentPage(department), 0);
    } catch {
      // Closing a modal should never block the UI.
    }
    delete modalRoot.dataset.discardTrainingExamId;
  }
  document.removeEventListener("keydown", handleModalEscape, true);
  modalRoot.classList.add("hidden");
  modalRoot.innerHTML = "";
}

function handleModalEscape(event) {
  if (event.key !== "Escape" || modalRoot.classList.contains("hidden")) return;
  event.preventDefault();
  const closeButton = modalRoot.querySelector(".modal-x");
  if (closeButton) closeButton.click();
  else closeModal();
}

function openDefconModal() {
  if (!hasRole("Supervisor")) {
    openModal(`<h3>Keine Berechtigung</h3><p class="muted">DEFCON kann ab Supervisor geändert werden.</p><div class="modal-actions"><button class="ghost-btn" data-close>Schließen</button></div>`);
    return;
  }
  openModal(`
    <h3>DEFCON bearbeiten</h3>
    <label>Stufe
      <select id="defconSelect">${[1, 2, 3, 4, 5].map((nr) => `<option ${state.settings.defcon === `DEFCON ${nr}` ? "selected" : ""}>DEFCON ${nr}</option>`).join("")}</select>
    </label>
    <label>Beschreibung<input id="defconText" value="${escapeHtml(state.settings.defconText ?? "Automatisch / Manuell aktualisierbar")}"></label>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions">
      <button class="ghost-btn" data-close>Abbrechen</button>
      <button class="blue-btn" id="saveDefcon">Bestätigen</button>
    </div>
  `, (modal) => {
    modal.querySelector("#saveDefcon").addEventListener("click", async () => {
      try {
        await api("/api/settings/defcon", { method: "PATCH", body: JSON.stringify({ defcon: $("#defconSelect").value, defconText: $("#defconText").value }) });
        closeModal();
        await bootstrap();
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
  });
}

function openNoteModal(note = null) {
  if (!note?.id) note = null;
  if (!hasRole("Supervisor")) {
    openModal(`<h3>Keine Berechtigung</h3><p class="muted">Notizen können ab Supervisor erstellt werden.</p><div class="modal-actions"><button class="ghost-btn" data-close>Schließen</button></div>`);
    return;
  }
  const isEdit = Boolean(note);
  openModal(`
    <h3>${isEdit ? "Notiz bearbeiten" : "Notiz hinzufügen"}</h3>
    <label>Titel<input id="noteTitle" value="${escapeHtml(note?.title || "")}" required></label>
    <label>Priorität
      <select id="notePriority">
        ${["Info", "Anweisung", "Direktion"].map((priority) => `<option ${note?.priority === priority ? "selected" : ""}>${priority}</option>`).join("")}
      </select>
    </label>
    <label>Text<textarea id="noteText" required>${escapeHtml(note?.text || "")}</textarea></label>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions">
      <button class="ghost-btn" data-close>Abbrechen</button>
      <button class="blue-btn" id="saveNote">${isEdit ? "Notiz aktualisieren" : "Absenden"}</button>
    </div>
  `, (modal) => {
    modal.querySelector("#saveNote").addEventListener("click", async () => {
      try {
        await api(isEdit ? `/api/notes/${note.id}` : "/api/notes", {
          method: isEdit ? "PATCH" : "POST",
          body: JSON.stringify({ title: $("#noteTitle").value, priority: $("#notePriority").value, text: $("#noteText").value })
        });
        closeModal();
        await bootstrap();
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
  });
}

function openStartDutyModal() {
  let selected = "";
  openModal(`
    <div class="duty-workflow">
      <div class="duty-workflow-head">
        <span class="duty-head-icon">${iconSvg("Dienstblatt")}</span>
        <div>
          <span class="duty-kicker">Dienststatus</span>
          <h3>Dienst eintragen</h3>
          <p>Wähle den Bereich aus, in dem du jetzt arbeitest.</p>
        </div>
      </div>
      <div class="duty-choice-grid">
      ${availableDutyOptions().map((option) => {
        const disabled = option.teamlerOnly && !state.currentUser.teamler && !hasRole("IT");
        return `
        <button class="duty-choice-card ${escapeHtml(option.tone || "default")}" data-status="${escapeHtml(option.title)}" ${disabled ? "disabled" : ""}>
          <span class="duty-card-accent"></span>
          <i>${iconSvg(option.icon)}</i>
          <span class="duty-card-copy"><strong>${escapeHtml(option.title)}</strong><small>${escapeHtml(disabled ? "Nur für Teamler freigegeben" : option.description)}</small></span>
          <span class="duty-card-check">✓</span>
        </button>
      `;}).join("")}
      </div>
      <p id="modalError" class="form-error"></p>
      <div class="duty-modal-actions">
        <button class="ghost-btn" data-close>Abbrechen</button>
        <button class="blue-btn" id="confirmDuty" disabled>Eintragen</button>
      </div>
    </div>
  `, (modal) => {
    modal.classList.add("duty-modal");
    modal.querySelectorAll(".duty-choice-card").forEach((button) => {
      button.addEventListener("click", () => {
        selected = button.dataset.status;
        modal.querySelectorAll(".duty-choice-card").forEach((item) => item.classList.remove("active"));
        button.classList.add("active");
        $("#confirmDuty").disabled = false;
      });
    });
    modal.querySelector("#confirmDuty").addEventListener("click", async () => {
      try {
        await api("/api/duty/start", { method: "POST", body: JSON.stringify({ status: selected }) });
        closeModal();
        await bootstrap();
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
  });
}

function openSwitchDutyModal() {
  let selected = "";
  const current = state.duty.find((entry) => entry.userId === state.currentUser.id)?.status || "";
  openModal(`
    <div class="duty-workflow">
      <div class="duty-workflow-head">
        <span class="duty-head-icon">${iconSvg("Einsatzzentrale")}</span>
        <div>
          <span class="duty-kicker">Dienstwechsel</span>
          <h3>Dienst umtragen</h3>
          <p>Aktuell: ${escapeHtml(current || "Nicht im Dienst")}</p>
        </div>
      </div>
      <div class="duty-choice-grid">
      ${availableDutyOptions().filter((option) => option.title !== current).map((option) => {
        const disabled = option.teamlerOnly && !state.currentUser.teamler && !hasRole("IT");
        return `<button class="duty-choice-card ${escapeHtml(option.tone || "default")}" data-status="${escapeHtml(option.title)}" ${disabled ? "disabled" : ""}><span class="duty-card-accent"></span><i>${iconSvg(option.icon)}</i><span class="duty-card-copy"><strong>${escapeHtml(option.title)}</strong><small>${escapeHtml(disabled ? "Nur für Teamler freigegeben" : option.description)}</small></span><span class="duty-card-check">✓</span></button>`;
      }).join("")}
      </div>
      <p id="modalError" class="form-error"></p>
      <div class="duty-modal-actions">
        <button class="ghost-btn" data-close>Abbrechen</button>
        <button class="blue-btn" id="confirmSwitchDuty" disabled>Umtragen</button>
      </div>
    </div>
  `, (modal) => {
    modal.classList.add("duty-modal");
    modal.querySelectorAll(".duty-choice-card").forEach((button) => button.addEventListener("click", () => {
      selected = button.dataset.status;
      modal.querySelectorAll(".duty-choice-card").forEach((item) => item.classList.toggle("active", item === button));
      $("#confirmSwitchDuty").disabled = false;
    }));
    modal.querySelector("#confirmSwitchDuty").addEventListener("click", async () => {
      try {
        await api("/api/duty/switch", { method: "POST", body: JSON.stringify({ status: selected }) });
        closeModal();
        await bootstrap();
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
  });
}

function openStopDutyModal(myDuty) {
  openModal(`
    <h3>Dienst austragen</h3>
    <p class="muted">Aktueller Status: ${escapeHtml(myDuty?.status || "Nicht im Dienst")}</p>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions">
      <button class="ghost-btn" data-close>Abbrechen</button>
      <button class="red-btn" id="confirmStopDuty" ${myDuty ? "" : "disabled"}>Dienst beenden</button>
    </div>
  `, (modal) => {
    modal.querySelector("#confirmStopDuty").addEventListener("click", async () => {
      try {
        await api("/api/duty/stop", { method: "POST", body: "{}" });
        closeModal();
        await bootstrap();
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
  });
}

function openStopAllDutyModal() {
  if (!canAccess("actions", "stopAllDuty", "Direktion")) {
    openModal(`<h3>Keine Berechtigung</h3><p class="muted">Alle austragen ist für dich nicht freigegeben.</p><div class="modal-actions"><button class="ghost-btn" data-close>Schließen</button></div>`);
    return;
  }
  openModal(`
    <h3>Alle Officer austragen</h3>
    <p class="muted">Damit werden alle aktiven Dienst-Einträge beendet.</p>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions">
      <button class="ghost-btn" data-close>Abbrechen</button>
      <button class="orange-btn" id="confirmStopAll">Alle Austragen</button>
    </div>
  `, (modal) => {
    modal.querySelector("#confirmStopAll").addEventListener("click", async () => {
      try {
        await api("/api/duty/stop-all", { method: "POST", body: "{}" });
        closeModal();
        await bootstrap();
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
  });
}

function openUserModal(user) {
  const isEdit = Boolean(user);
  const selectedTrainings = user?.trainings || {};
  const baseRoles = editableRoleOptions(user);
  const selectedRole = baseRoles.includes(baseRoleForUser(user)) ? baseRoleForUser(user) : "User";
  const initialDn = String(user?.dn || "");
  const initialDnConflict = dnConflictFor(initialDn, user?.id);
  openModal(`
    <h3>${isEdit ? "Mitglied bearbeiten" : "Neues Mitglied einstellen"}</h3>
    <form id="userForm" class="form-grid">
      <label>Name<input name="firstName" value="${escapeHtml(user?.firstName || "")}" required></label>
      <label>Nachname / Doppelname<input name="lastName" value="${escapeHtml(user?.lastName || "")}" required></label>
      <label>Telefonnummer<input name="phone" value="${escapeHtml(user?.phone || "")}" required></label>
      <label>DN<input name="dn" id="userDnInput" inputmode="numeric" pattern="[0-9]+" value="${escapeHtml(initialDn)}" required></label>
      <label>Discord User-ID<input name="discordId" inputmode="numeric" pattern="[0-9]*" value="${escapeHtml(user?.discordId || "")}" placeholder="Optional"></label>
      <label>Einstellungsdatum<input name="joinedAt" type="date" value="${escapeHtml((user?.joinedAt || new Date().toISOString()).slice(0, 10))}"></label>
      <div id="userDnConflict" class="full">${renderDnConflictBox(initialDnConflict, initialDn)}</div>
      <label>Rang
        <select name="rank">${state.ranks.map((rank) => `<option value="${rank.value}" ${Number(user?.rank ?? 0) === Number(rank.value) ? "selected" : ""}>${escapeHtml(rankOptionLabel(rank))}</option>`).join("")}</select>
      </label>
      <label>Rolle
        <select name="role">${baseRoles.map((role) => `<option ${selectedRole === role ? "selected" : ""}>${escapeHtml(role)}</option>`).join("")}</select>
      </label>
      ${renderTeamlerControl(user)}
      ${renderItRoleControls(user)}
      <div class="full">
        <p class="muted">Ausbildungen</p>
        ${renderTrainingPicker(selectedTrainings)}
      </div>
      <p id="modalError" class="form-error full"></p>
      <div class="modal-actions full">
        <button class="ghost-btn" type="button" data-close>Abbrechen</button>
        <button class="blue-btn" type="submit">${isEdit ? "Speichern" : "Einstellen"}</button>
      </div>
    </form>
  `, (modal) => {
    modal.querySelector("#userForm").addEventListener("submit", async (event) => {
      event.preventDefault();
      const form = new FormData(event.currentTarget);
      const body = Object.fromEntries(form.entries());
      body.baseRole = body.role;
      if (canGrantItRoles()) {
        body.role = form.get("isITLead") === "on" ? "IT-Leitung" : form.get("isIT") === "on" ? "IT" : body.baseRole;
      } else {
        body.role = user?.role || body.baseRole;
      }
      delete body.isIT;
      delete body.isITLead;
      body.departments = user?.departments || [];
      body.rank = Number(body.rank);
      body.teamler = form.get("teamler") === "on";
      body.overwriteDn = $("#overwriteDn")?.checked || false;
      body.trainings = Object.fromEntries(trainings.map((training) => [training, form.get(`training_${training}`) === "on"]));
      try {
        await api(isEdit ? `/api/users/${user.id}` : "/api/users", {
          method: isEdit ? "PATCH" : "POST",
          body: JSON.stringify(body)
        });
        closeModal();
        await bootstrap();
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
    modal.querySelector("#userDnInput").addEventListener("input", (event) => {
      const dn = event.target.value;
      modal.querySelector("#userDnConflict").innerHTML = renderDnConflictBox(dnConflictFor(dn, user?.id), dn);
    });
  });
}

function findAnyUser(userId) {
  return [...(state.users || []), ...(state.archivedUsers || [])].find((item) => item.id === userId);
}

async function saveUprankRules(event) {
  event.preventDefault();
  const form = new FormData(event.currentTarget);
  const rules = uprankRules().map((rule) => ({
    targetRank: Number(rule.targetRank),
    minDays: Number(form.get(`minDays_${rule.targetRank}`) || 0),
    specialOnly: form.get(`specialOnly_${rule.targetRank}`) === "on",
    trainings: trainings.filter((training) => form.get(`rule_${rule.targetRank}_${training}`) === "on")
  }));
  try {
    await api("/api/settings/uprank-rules", { method: "PATCH", body: JSON.stringify({ rules }) });
    await bootstrap();
  } catch (error) {
    $("#uprankRulesError").textContent = error.message;
  }
}

function openUprankModal(user, forceSpecial = false) {
  const evaluation = evaluateUprank(user);
  openModal(`
    <h3>Uprank durchf\u00fchren</h3>
    <p class="muted">${escapeHtml(fullName(user))} \u00b7 ${escapeHtml(rankLabel(user.rank))} \u2192 ${escapeHtml(rankLabel(evaluation.targetRank))}</p>
    <div class="uprank-modal-summary">
      <span class="requirement-chip ${evaluation.missingDays ? "missing" : "ok"}">${evaluation.missingDays ? `${evaluation.missingDays} Tage fehlen` : "Dauer erf\u00fcllt"}</span>
      <span class="requirement-chip ${evaluation.missingTrainings.length ? "missing" : "ok"}">${evaluation.missingTrainings.length ? `Fehlt: ${escapeHtml(evaluation.missingTrainings.join(", "))}` : "Ausbildungen erf\u00fcllt"}</span>
      <span class="requirement-chip ${forceSpecial ? "special" : "ok"}">${forceSpecial ? "Sonderuprank" : "Regul\u00e4rer Uprank"}</span>
    </div>
    <form id="uprankForm" class="form-grid">
      <label class="full">Begr\u00fcndung<textarea name="reason" placeholder="Kurz begr\u00fcnden, besonders bei Sonderupranks."></textarea></label>
      <label class="checkbox-line full">Ingame get\u00e4tigt<input type="checkbox" name="ingameDone" required></label>
      <label class="checkbox-line full">Discord get\u00e4tigt<input type="checkbox" name="discordDone" required></label>
      <p id="modalError" class="form-error full"></p>
      <div class="modal-actions full">
        <button class="ghost-btn" type="button" data-close>Abbrechen</button>
        <button class="blue-btn" type="submit">Uprank speichern</button>
      </div>
    </form>
  `, (modal) => {
    modal.querySelector("#uprankForm").addEventListener("submit", async (event) => {
      event.preventDefault();
      const form = new FormData(event.currentTarget);
      try {
        await api(`/api/users/${user.id}/uprank`, {
          method: "POST",
          body: JSON.stringify({
            targetRank: evaluation.targetRank,
            reason: form.get("reason"),
            ingameDone: form.get("ingameDone") === "on",
            discordDone: form.get("discordDone") === "on",
            special: forceSpecial
          })
        });
        closeModal();
        await bootstrap();
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
  });
}

function openUprankAdjustmentModal(user, type) {
  const evaluation = evaluateUprank(user);
  const isShortening = type === "Verkürzung";
  openModal(`
    <h3>${escapeHtml(type)} eintragen</h3>
    <p class="muted">${escapeHtml(fullName(user))} · Zielrang ${escapeHtml(rankLabel(evaluation.targetRank))}</p>
    <form id="uprankAdjustmentForm" class="form-grid">
      ${isShortening ? `<label>Tage Verkürzung<input type="number" name="days" min="1" value="7" required></label>` : ""}
      <label class="${isShortening ? "" : "full"}">Zielrang
        <select name="targetRank">
          ${state.ranks.filter((rank) => Number(rank.value) > Number(user.rank)).map((rank) => `<option value="${rank.value}" ${Number(rank.value) === evaluation.targetRank ? "selected" : ""}>${escapeHtml(rankOptionLabel(rank))}</option>`).join("")}
        </select>
      </label>
      <label class="full">Grund<textarea name="reason" required></textarea></label>
      <p id="modalError" class="form-error full"></p>
      <div class="modal-actions full">
        <button class="ghost-btn" type="button" data-close>Abbrechen</button>
        <button class="blue-btn" type="submit">Speichern</button>
      </div>
    </form>
  `, (modal) => {
    modal.querySelector("#uprankAdjustmentForm").addEventListener("submit", async (event) => {
      event.preventDefault();
      const form = new FormData(event.currentTarget);
      try {
        await api(`/api/users/${user.id}/uprank-adjustments`, {
          method: "POST",
          body: JSON.stringify({
            type,
            targetRank: Number(form.get("targetRank")),
            days: Number(form.get("days") || 0),
            reason: form.get("reason")
          })
        });
        closeModal();
        await bootstrap();
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
  });
}

function openDeleteUserModal(userId) {
  const user = findAnyUser(userId);
  openModal(`
    <h3>Account löschen</h3>
    <p class="muted">${escapeHtml(fullName(user))} wirklich löschen?</p>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions">
      <button class="ghost-btn" data-close>Abbrechen</button>
      <button class="red-btn" id="confirmDelete">Löschen</button>
    </div>
  `, (modal) => {
    modal.querySelector("#confirmDelete").addEventListener("click", async () => {
      try {
        await api(`/api/users/${userId}`, { method: "DELETE" });
        closeModal();
        await bootstrap();
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
  });
}

function openRehireUserModal(user) {
  const info = terminationInfo(user);
  const oldDn = String(info.oldDn || user.dn || "");
  const selectedTrainings = info.oldTrainings || user.trainings || {};
  const baseRoles = editableRoleOptions(user);
  const selectedRole = baseRoles.includes(baseRoleForUser(user)) ? baseRoleForUser(user) : "User";
  openModal(`
    <h3>Wiedereinstellen</h3>
    <div class="old-data-box">
      <div>
        <strong>Alte Daten</strong>
        <p>Name: ${escapeHtml(fullName(user))} · Telefon: ${escapeHtml(user.phone || "-")} · Dienstnummer: ${escapeHtml(oldDn || "-")} · Rang: ${escapeHtml(rankLabel(info.oldRank ?? user.rank))}</p>
      </div>
      <button class="ghost-btn" type="button" id="fillOldRehireData">Alte Daten übernehmen</button>
    </div>
    <div id="rehireDnConflict"></div>
    <form id="rehireUserForm" class="form-grid">
      <label>Name<input name="firstName" id="rehireFirstName" value="" required></label>
      <label>Nachname / Doppelname<input name="lastName" id="rehireLastName" value="" required></label>
      <label>Telefonnummer<input name="phone" id="rehirePhone" value="" required></label>
      <label>Dienstnummer<input name="dn" id="rehireDnInput" inputmode="numeric" pattern="[0-9]+" value="" required></label>
      <label>Einstellungsdatum<input name="joinedAt" id="rehireJoinedAt" type="date" value=""></label>
      <label>Rang
        <select name="rank" id="rehireRank">
          <option value="">Rang auswählen</option>
          ${state.ranks.map((rank) => `<option value="${rank.value}">${escapeHtml(rankOptionLabel(rank))}</option>`).join("")}
        </select>
      </label>
      <label>Rolle
        <select name="role">${baseRoles.map((role) => `<option ${selectedRole === role ? "selected" : ""}>${escapeHtml(role)}</option>`).join("")}</select>
      </label>
      ${renderTeamlerControl(user)}
      ${renderItRoleControls(user)}
      <label class="full">Grund der Wiedereinstellung<textarea name="reason">Wiedereinstellung</textarea></label>
      <div class="full">
        <p class="muted">Ausbildungen</p>
        ${renderTrainingPicker(selectedTrainings)}
      </div>
      <p id="modalError" class="form-error full"></p>
      <div class="modal-actions full">
        <button class="ghost-btn" type="button" data-close>Abbrechen</button>
        <button class="blue-btn" type="submit">Wiedereinstellen</button>
      </div>
    </form>
  `, (modal) => {
    modal.querySelector("#rehireUserForm").addEventListener("submit", async (event) => {
      event.preventDefault();
      const form = new FormData(event.currentTarget);
      const baseRole = form.get("role");
      const role = canGrantItRoles() ? (form.get("isITLead") === "on" ? "IT-Leitung" : form.get("isIT") === "on" ? "IT" : baseRole) : user.role;
      try {
        await api(`/api/users/${user.id}/rehire`, {
          method: "POST",
          body: JSON.stringify({
            dn: form.get("dn"),
            rank: Number(form.get("rank")),
            firstName: form.get("firstName"),
            lastName: form.get("lastName"),
            phone: form.get("phone"),
            joinedAt: form.get("joinedAt"),
            role,
            baseRole,
            teamler: form.get("teamler") === "on",
            reason: form.get("reason"),
            overwriteDn: $("#overwriteDn")?.checked || false,
            trainings: Object.fromEntries(trainings.map((training) => [training, form.get(`training_${training}`) === "on"]))
          })
        });
        closeModal();
        await bootstrap();
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
    modal.querySelector("#rehireDnInput").addEventListener("input", (event) => {
      const dn = event.target.value;
      modal.querySelector("#rehireDnConflict").innerHTML = renderDnConflictBox(dnConflictFor(dn, user.id), dn);
    });
    modal.querySelector("#fillOldRehireData").addEventListener("click", () => {
      modal.querySelector("#rehireFirstName").value = user.firstName || "";
      modal.querySelector("#rehireLastName").value = user.lastName || "";
      modal.querySelector("#rehirePhone").value = user.phone || "";
      modal.querySelector("#rehireDnInput").value = oldDn;
      modal.querySelector("#rehireJoinedAt").value = new Date().toISOString().slice(0, 10);
      modal.querySelector("#rehireRank").value = String(info.oldRank ?? user.rank ?? "");
      modal.querySelector("#rehireDnConflict").innerHTML = renderDnConflictBox(dnConflictFor(oldDn, user.id), oldDn);
    });
  });
}

function openUserActionsModal(user) {
  openModal(`
    <h3>Aktionen</h3>
    <p class="muted">${escapeHtml(fullName(user))}</p>
    <div class="action-menu-list">
      <button class="ghost-btn action-edit-btn" id="actionEditUser">${actionIcon("edit")} Bearbeiten</button>
      <button class="blue-btn action-menu-btn" id="actionOpenFile">Akte öffnen</button>
      <button class="orange-btn action-menu-btn" id="actionToggleLock">${user.locked ? "Entsperren" : "Sperren"}</button>
      <button class="orange-btn action-menu-btn" id="actionSuspendUser">Suspendieren</button>
      <button class="red-btn action-menu-btn" id="actionDismissUser">Entlassen</button>
      <button class="red-btn action-menu-btn" id="actionDeleteUser">Löschen</button>
    </div>
    <div class="modal-actions"><button class="ghost-btn" data-close>Schließen</button></div>
  `, (modal) => {
    modal.querySelector("#actionEditUser").addEventListener("click", () => {
      closeModal();
      openUserModal(user);
    });
    modal.querySelector("#actionOpenFile").addEventListener("click", () => {
      closeModal();
      openPersonnelFileModal(user);
    });
    modal.querySelector("#actionToggleLock").addEventListener("click", () => openReasonUserModal(user, user.locked ? "Entsperren" : "Sperren", `/api/users/${user.id}/lock`, "PATCH", { locked: !user.locked }));
    modal.querySelector("#actionSuspendUser").addEventListener("click", () => openSuspendUserModal(user));
    modal.querySelector("#actionDismissUser").addEventListener("click", () => openDismissUserModal(user));
    modal.querySelector("#actionDeleteUser").addEventListener("click", () => {
      closeModal();
      openDeleteUserModal(user.id);
    });
  });
}

function openSuspendUserModal(user) {
  openReasonUserModal(user, "Suspendieren", `/api/users/${user.id}/suspend`, "POST");
}

function openDismissUserModal(user) {
  openReasonUserModal(user, "Entlassen", `/api/users/${user.id}/dismiss`, "POST");
}

function openPersonnelFileModal(user) {
  const entries = (state.disciplinary || []).filter((entry) => entry.userId === user.id).sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
  const isStrike = (entry) => entry.type === "Strike" || (entry.type === "Sanktion" && entry.sanctionType === "Strike");
  const isExpired = (entry) => entry.expiresAt && new Date(entry.expiresAt) <= new Date();
  const isActiveStrike = (entry) => isStrike(entry) && !entry.archivedAt && !isExpired(entry);
  const strikeWeight = (entry) => Math.max(1, Number(entry.strikeCount || 1));
  const notes = entries.filter((entry) => entry.type === "Aktennotiz");
  const sanctions = entries.filter((entry) => entry.type === "Sanktion" || entry.type === "Strike");
  const fines = sanctions.filter((entry) => entry.sanctionType === "Geldstrafe");
  const openFines = fines.filter((entry) => !entry.paidAt);
  const openFineAmount = openFines.reduce((sum, entry) => sum + Number(entry.amount || 0), 0);
  const history = entries.filter((entry) => !["Aktennotiz", "Sanktion", "Strike"].includes(entry.type));
  const activeStrikes = sanctions.filter(isActiveStrike).reduce((sum, entry) => sum + strikeWeight(entry), 0);
  openModal(`
    <h3>Akte - ${escapeHtml(fullName(user))}</h3>
    <div class="file-summary-grid">
      <div class="info-box"><strong>Status</strong><p>${renderAccountStatus(user)}</p></div>
      <div class="info-box"><strong>Aktive Strikes</strong><p><span class="strike-counter ${activeStrikes >= 3 ? "danger" : activeStrikes >= 2 ? "warn" : ""}">${activeStrikes}/3</span></p></div>
      <div class="info-box"><strong>Offene Geldstrafen</strong><p><span class="file-pill open">${openFines.length} / ${openFineAmount.toLocaleString("de-DE")} $</span></p></div>
      <div class="info-box"><strong>DN</strong><p>${escapeHtml(user.dn || "-")}</p></div>
      <div class="info-box"><strong>Rang</strong><p>${escapeHtml(rankLabel(user.rank))}</p></div>
    </div>
    <div class="button-row file-action-row">
      <button class="blue-btn" id="addFileNote">Notiz hinzufügen</button>
      <button class="orange-btn" id="addFileSanction">Sanktion vergeben</button>
    </div>
    <div class="file-section-grid">
      <section class="file-section notes">
        <h4>Notizen</h4>
        <div class="personnel-file-list compact">${notes.map((entry) => renderFileEntry(entry)).join("") || `<p class="muted">Noch keine Notizen.</p>`}</div>
      </section>
      <section class="file-section sanctions">
        <h4>Sanktionen</h4>
        <div class="personnel-file-list compact">${sanctions.map((entry) => renderFileEntry(entry, isActiveStrike(entry), isExpired(entry))).join("") || `<p class="muted">Noch keine Sanktionen.</p>`}</div>
      </section>
      <section class="file-section fines">
        <h4>Geldstrafen</h4>
        <div class="personnel-file-list compact">${fines.map((entry) => renderFineEntry(entry)).join("") || `<p class="muted">Keine Geldstrafen.</p>`}</div>
      </section>
      <section class="file-section history">
        <h4>Verlauf</h4>
        <div class="personnel-file-list compact">${history.map((entry) => renderFileEntry(entry)).join("") || `<p class="muted">Noch kein Verlauf.</p>`}</div>
      </section>
    </div>
    <div class="modal-actions"><button class="ghost-btn" data-close>Schließen</button></div>
  `, (modal) => {
    modal.querySelector("#addFileNote").addEventListener("click", () => openPersonnelFileEntryModal(user, "Aktennotiz"));
    modal.querySelector("#addFileSanction").addEventListener("click", () => openPersonnelFileEntryModal(user, "Sanktion", activeStrikes));
    modal.querySelectorAll(".remove-file-entry").forEach((button) => button.addEventListener("click", async () => {
      try {
        await api(`/api/users/${user.id}/file/${button.dataset.id}`, { method: "DELETE" });
        await bootstrap();
        openPersonnelFileModal(findAnyUser(user.id));
      } catch (error) {
        openModal(`<h3>Akteneintrag konnte nicht entfernt werden</h3><p class="form-error">${escapeHtml(error.message)}</p><div class="modal-actions"><button class="ghost-btn" data-close>Schließen</button></div>`);
      }
    }));
    modal.querySelectorAll(".note-settings").forEach((button) => button.addEventListener("click", () => {
      openNoteActionModal(user, entries.find((entry) => entry.id === button.dataset.id));
    }));
    modal.querySelectorAll(".mark-fine-paid").forEach((button) => button.addEventListener("click", async () => {
      try {
        await api(`/api/users/${user.id}/file/${button.dataset.id}`, { method: "PATCH", body: JSON.stringify({ paid: true }) });
        await bootstrap();
        openPersonnelFileModal(findAnyUser(user.id));
      } catch (error) {
        openModal(`<h3>Geldstrafe konnte nicht markiert werden</h3><p class="form-error">${escapeHtml(error.message)}</p><div class="modal-actions"><button class="ghost-btn" data-close>Schließen</button></div>`);
      }
    }));
  });
}

function renderFileEntry(entry, activeStrike = false, expired = false) {
  const sanctionType = entry.sanctionType || (entry.type === "Strike" ? "Strike" : "");
  const className = entry.type === "Aktennotiz" ? "note" : sanctionType === "Geldstrafe" ? "fine" : sanctionType === "Strike" ? "strike" : ["Entlassen", "Sperre", "Suspendierung"].includes(entry.type) ? "danger" : "history";
  const archived = (entry.type === "Sanktion" || entry.type === "Strike") && (entry.archivedAt || expired);
  const title = entry.type === "Sanktion" || entry.type === "Strike" ? `${sanctionType}${entry.title && entry.title !== sanctionType ? ` - ${entry.title}` : ""}` : entry.type;
  return `
    <article class="file-entry ${className} ${archived ? "archived" : ""}">
      <div>
        <div class="file-entry-head">
          <strong>${escapeHtml(title)}</strong>
          ${activeStrike ? `<span class="file-pill active">Aktiv${Number(entry.strikeCount || 1) > 1 ? ` (${Number(entry.strikeCount)})` : ""}</span>` : ""}
          ${archived ? `<span class="file-pill archived">${entry.archivedAt ? "Archiviert" : "Abgelaufen"}</span>` : ""}
          ${entry.amount ? `<span class="file-pill fine">${Number(entry.amount).toLocaleString("de-DE")} $</span>` : ""}
          ${entry.paidAt ? `<span class="file-pill paid">Bezahlt</span>` : ""}
          ${entry.type === "Aktennotiz" ? `<button class="calendar-event-settings note-settings inline-note-settings" data-id="${escapeHtml(entry.id)}" title="Notiz verwalten">${iconSvg("Settings")}</button>` : ""}
        </div>
        <p>${escapeHtml(entry.reason || "-")}</p>
        <small>${formatDateTime(entry.createdAt)} - ${escapeHtml(entry.actorName || "-")}${entry.expiresAt ? ` - Ablauf: ${formatDate(entry.expiresAt)}` : ""}${entry.archivedBy ? ` - Archiviert von ${escapeHtml(entry.archivedBy)}` : ""}</small>
      </div>
      ${(entry.type === "Sanktion" || entry.type === "Strike") && !entry.archivedAt ? `<button class="ghost-btn remove-file-entry" data-id="${escapeHtml(entry.id)}">Archivieren</button>` : ""}
    </article>
  `;
}

function renderFineEntry(entry) {
  return `
    <article class="file-entry fine ${entry.paidAt ? "paid-entry" : ""}">
      <div>
        <div class="file-entry-head">
          <strong>${escapeHtml(entry.title || "Geldstrafe")}</strong>
          <span class="file-pill fine">${Number(entry.amount || 0).toLocaleString("de-DE")} $</span>
          <span class="file-pill ${entry.paidAt ? "paid" : "open"}">${entry.paidAt ? "Bezahlt" : "Offen"}</span>
        </div>
        <p>${escapeHtml(entry.reason || "-")}</p>
        <small>${formatDateTime(entry.createdAt)} - ${escapeHtml(entry.actorName || "-")}${entry.paidAt ? ` - Bezahlt: ${formatDateTime(entry.paidAt)} von ${escapeHtml(entry.paidBy || "-")}` : ""}</small>
      </div>
      ${entry.paidAt ? "" : `<button class="blue-btn mark-fine-paid" data-id="${escapeHtml(entry.id)}">Bezahlt</button>`}
    </article>
  `;
}

function openNoteActionModal(user, entry) {
  if (!entry) return;
  openModal(`
    <h3>Notiz verwalten</h3>
    <p class="muted">${escapeHtml(fullName(user))}</p>
    <div class="action-menu-list">
      <button class="ghost-btn action-menu-btn" id="editFileNote">${actionIcon("edit")} Bearbeiten</button>
      <button class="red-btn action-menu-btn" id="deleteFileNote">${actionIcon("delete")} Löschen</button>
    </div>
    <div class="modal-actions"><button class="ghost-btn" data-close>Schließen</button></div>
  `, (modal) => {
    modal.querySelector("#editFileNote").addEventListener("click", () => openEditFileNoteModal(user, entry));
    modal.querySelector("#deleteFileNote").addEventListener("click", async () => {
      try {
        await api(`/api/users/${user.id}/file/${entry.id}`, { method: "DELETE" });
        await bootstrap();
        openPersonnelFileModal(findAnyUser(user.id));
      } catch (error) {
        openModal(`<h3>Notiz konnte nicht gelöscht werden</h3><p class="form-error">${escapeHtml(error.message)}</p><div class="modal-actions"><button class="ghost-btn" data-close>Schließen</button></div>`);
      }
    });
  });
}

function openEditFileNoteModal(user, entry) {
  openModal(`
    <h3>Notiz bearbeiten</h3>
    <label>Notiz<textarea id="editFileNoteText" required>${escapeHtml(entry.reason || "")}</textarea></label>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions"><button class="ghost-btn" data-close>Abbrechen</button><button class="blue-btn" id="saveFileNoteEdit">Speichern</button></div>
  `, (modal) => {
    modal.querySelector("#saveFileNoteEdit").addEventListener("click", async () => {
      try {
        await api(`/api/users/${user.id}/file/${entry.id}`, { method: "PATCH", body: JSON.stringify({ reason: $("#editFileNoteText").value }) });
        await bootstrap();
        openPersonnelFileModal(findAnyUser(user.id));
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
  });
}

function openPersonnelFileEntryModal(user, type, activeStrikes = 0) {
  const catalog = [
    { title: "Respektloses Verhalten", sanctionType: "Geldstrafe", base: 1500 },
    { title: "Dienstpflicht verletzt", sanctionType: "Strike", base: 0 },
    { title: "Unangemessene Fahrweise", sanctionType: "Geldstrafe", base: 1000 }
  ];
  const rankFactor = Math.max(1, Number(user.rank || 0) + 1);
  openModal(`
    <h3>${type === "Sanktion" ? "Sanktion vergeben" : "Notiz hinzufügen"}</h3>
    <p class="muted">${escapeHtml(fullName(user))}</p>
    ${type === "Sanktion" ? `
      <label>Sanktionskatalog
        <select id="sanctionCatalog">
          <option value="custom">Custom Sanktion</option>
          ${catalog.map((item, index) => `<option value="${index}">${escapeHtml(item.title)}</option>`).join("")}
        </select>
      </label>
      <label>Sanktionsart
        <select id="sanctionType">
          <option>Geldstrafe</option>
          <option ${activeStrikes >= 3 ? "disabled" : ""}>Strike</option>
          <option>Custom</option>
        </select>
      </label>
      <label>Titel<input id="fileEntryTitle" placeholder="z.B. Dienstpflicht verletzt"></label>
      <label>Strikes
        <select id="sanctionStrikeCount">
          ${[1, 2, 3].map((count) => `<option value="${count}" ${activeStrikes + count > 3 ? "disabled" : ""}>${count} Strike${count > 1 ? "s" : ""}</option>`).join("")}
        </select>
      </label>
      <label>Geldstrafe<input id="sanctionAmount" type="number" min="0" step="100" value="${1000 * rankFactor}"></label>
      <label>Ablaufdatum für Strike<input id="sanctionExpiresAt" type="date"></label>
      <p class="muted">Aktive Strikes: ${activeStrikes}/3. Archivierte oder abgelaufene Strikes zählen nicht mehr aktiv.</p>
    ` : ""}
    <label>${type === "Sanktion" ? "Grund / Beschreibung" : "Notiz"}<textarea id="fileEntryReason" required></textarea></label>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions">
      <button class="ghost-btn" data-close>Abbrechen</button>
      <button class="${type === "Sanktion" ? "orange-btn" : "blue-btn"}" id="saveFileEntry">Speichern</button>
    </div>
  `, (modal) => {
    modal.querySelector("#sanctionCatalog")?.addEventListener("change", (event) => {
      const item = catalog[Number(event.target.value)];
      if (!item) return;
      modal.querySelector("#fileEntryTitle").value = item.title;
      modal.querySelector("#sanctionType").value = item.sanctionType;
      modal.querySelector("#sanctionAmount").value = item.sanctionType === "Geldstrafe" ? item.base * rankFactor : 0;
      modal.querySelector("#sanctionStrikeCount").value = "1";
    });
    modal.querySelector("#saveFileEntry").addEventListener("click", async () => {
      try {
        await api(`/api/users/${user.id}/file`, {
          method: "POST",
          body: JSON.stringify({
            type,
            reason: $("#fileEntryReason").value,
            title: $("#fileEntryTitle")?.value || "",
            sanctionType: $("#sanctionType")?.value || "",
            amount: $("#sanctionAmount")?.value || 0,
            strikeCount: $("#sanctionStrikeCount")?.value || 1,
            expiresAt: $("#sanctionExpiresAt")?.value || ""
          })
        });
        await bootstrap();
        openPersonnelFileModal(findAnyUser(user.id));
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
  });
}

function openReasonUserModal(user, title, path, method, extra = {}) {
  openModal(`
    <h3>${escapeHtml(title)}</h3>
    <p class="muted">${escapeHtml(fullName(user))}</p>
    <label>Grund<textarea id="actionReason" required></textarea></label>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions">
      <button class="ghost-btn" data-close>Abbrechen</button>
      <button class="${title === "Entlassen" ? "red-btn" : "orange-btn"}" id="confirmReasonAction">${escapeHtml(title)}</button>
    </div>
  `, (modal) => {
    modal.querySelector("#confirmReasonAction").addEventListener("click", async () => {
      try {
        await api(path, { method, body: JSON.stringify({ ...extra, reason: $("#actionReason").value }) });
        closeModal();
        await bootstrap();
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
  });
}

function openManualDutyModal() {
  openModal(`
    <h3>Dienstzeit hinzufügen</h3>
    <label>Mitglied<select id="manualDutyUser">${state.users.map((user) => `<option value="${user.id}">${escapeHtml(fullName(user))}</option>`).join("")}</select></label>
    <label>Diensttyp<input id="manualDutyStatus" value="Manuelle Korrektur"></label>
    <label>Beginn<input id="manualDutyStart" type="datetime-local"></label>
    <label>Ende<input id="manualDutyEnd" type="datetime-local"></label>
    <label>Grund<textarea id="manualDutyReason"></textarea></label>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions">
      <button class="ghost-btn" data-close>Abbrechen</button>
      <button class="blue-btn" id="saveManualDuty">Speichern</button>
    </div>
  `, (modal) => {
    modal.querySelector("#saveManualDuty").addEventListener("click", async () => {
      try {
        await api("/api/duty/manual", {
          method: "POST",
          body: JSON.stringify({
            userId: $("#manualDutyUser").value,
            status: $("#manualDutyStatus").value,
            startedAt: $("#manualDutyStart").value,
            endedAt: $("#manualDutyEnd").value,
            reason: $("#manualDutyReason").value
          })
        });
        closeModal();
        await bootstrap();
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
  });
}

function openDeleteNoteModal(noteId) {
  const note = state.notes.find((item) => item.id === noteId);
  openModal(`
    <h3>Notiz löschen</h3>
    <p class="muted">${escapeHtml(note?.title || "Diese Notiz")} wirklich löschen?</p>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions">
      <button class="ghost-btn" data-close>Abbrechen</button>
      <button class="red-btn" id="confirmDeleteNote">Notiz löschen</button>
    </div>
  `, (modal) => {
    modal.querySelector("#confirmDeleteNote").addEventListener("click", async () => {
      try {
        await api(`/api/notes/${noteId}`, { method: "DELETE" });
        closeModal();
        await bootstrap();
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
  });
}

function formatDepartmentText(text = "") {
  const escaped = escapeHtml(text || "Noch keine Rechte definiert.");
  return escaped
    .replace(/\*\*(.*?)\*\*/g, "<strong>$1</strong>")
    .replace(/^### (.*)$/gm, "<h4>$1</h4>")
    .replace(/^## (.*)$/gm, "<h3>$1</h3>")
    .replace(/\n/g, "<br>");
}

function renderStatusDot(status) {
  const className = status === "Hoch" ? "high" : status === "Mittel" ? "medium" : "normal";
  return `<span class="status-dot ${className}"></span>${escapeHtml(status)}`;
}

function openDepartmentInfoModal(department) {
  openModal(`
    <div class="department-modal-icon">${iconSvg("Einsatzzentrale")}</div>
    <h3>${escapeHtml(department.name)}</h3>
    <p class="muted">Detaillierte Informationen über die Abteilung</p>
    <div class="department-modal-heading">
      <h4>Abteilungsinformationen</h4>
      ${departmentActionAllowed(department, "departmentInfo") ? `<button class="blue-btn" id="editDepartmentInfo">${actionIcon("edit")} Bearbeiten</button>` : ""}
    </div>
    <div class="department-info-view">
      <div class="info-box full"><strong>${iconSvg("Informationen")} Beschreibung</strong><p>${escapeHtml(department.description)}</p></div>
      <div class="info-box"><strong>Bewerbungsstatus</strong><span class="application-pill ${department.applicationStatus === "Offen" ? "open" : "closed"}">${escapeHtml(department.applicationStatus)}</span></div>
      <div class="info-box"><strong>Voraussetzungen</strong><p><span class="requirements-pill">${escapeHtml(department.requirements)}</span></p></div>
      <div class="info-box full personnel-box"><strong>${iconSvg("Mitglieder")} Personal (${department.members.length})</strong>
        <div class="personnel-list">${department.members.map((member) => `<div class="personnel-row"><span><b>${escapeHtml(fullName(member.user))}</b><small>${escapeHtml(rankLabel(member.user.rank))}</small></span><span class="position-chip ${positionClass(member.position, department)}">${escapeHtml(member.position)}</span></div>`).join("") || "<p>Keine Mitglieder.</p>"}</div>
      </div>
    </div>
    <div class="modal-actions">
      <button class="ghost-btn" data-close>Schließen</button>
    </div>
  `, (modal) => {
    modal.querySelector("#editDepartmentInfo")?.addEventListener("click", () => openDepartmentInfoEditModal(department));
  });
}

function openDepartmentInfoEditModal(department) {
  openModal(`
    <h3>Abteilungsinformationen bearbeiten</h3>
    <p class="muted">${escapeHtml(department.name)}</p>
    <label>Beschreibung<textarea id="deptDescription">${escapeHtml(department.description)}</textarea></label>
    <label>Bewerbungsstatus
      <select id="deptApplicationStatus">
        <option ${department.applicationStatus === "Offen" ? "selected" : ""}>Offen</option>
        <option ${department.applicationStatus === "Geschlossen" ? "selected" : ""}>Geschlossen</option>
      </select>
    </label>
    <label>Voraussetzungen<input id="deptRequirements" value="${escapeHtml(department.requirements)}"></label>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions">
      <button class="ghost-btn" data-close>Abbrechen</button>
      <button class="blue-btn" id="saveDepartmentInfo">Speichern</button>
    </div>
  `, (modal) => {
    modal.querySelector("#saveDepartmentInfo").addEventListener("click", async () => {
      try {
        await api(`/api/departments/${department.id}/info`, {
          method: "PATCH",
          body: JSON.stringify({ description: $("#deptDescription").value, applicationStatus: $("#deptApplicationStatus").value, requirements: $("#deptRequirements").value })
        });
        closeModal();
        await bootstrap();
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
  });
}

async function saveDepartmentInfo(department, patch) {
  await api(`/api/departments/${department.id}/info`, {
    method: "PATCH",
    body: JSON.stringify({
      description: department.description,
      applicationStatus: department.applicationStatus,
      requirements: department.requirements,
      rightsText: department.rightsText || "",
      links: department.links || [],
      permits: department.permits || [],
      factions: department.factions || [],
      ...patch
    })
  });
  closeModal();
  await bootstrap();
}

function openDepartmentRightsModal(department) {
  openModal(`
    <h3>Rechte Definition bearbeiten</h3>
    <p class="muted">${escapeHtml(department.name)}</p>
    <label>Text<textarea id="rightsText" rows="12">${escapeHtml(department.rightsText || "")}</textarea></label>
    <p class="muted">Überschriften mit ##, dicke Schrift mit **Text**.</p>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions"><button class="ghost-btn" data-close>Abbrechen</button><button class="blue-btn" id="saveRightsText">Speichern</button></div>
  `, (modal) => {
    modal.querySelector("#saveRightsText").addEventListener("click", async () => {
      try {
        await saveDepartmentInfo(department, { rightsText: $("#rightsText").value });
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
  });
}

function upsertById(items, item) {
  const list = [...(items || [])];
  const nextItem = { ...item, id: item.id || `item_${Date.now()}_${Math.random().toString(16).slice(2)}` };
  const index = list.findIndex((entry) => entry.id === nextItem.id);
  if (index >= 0) list[index] = nextItem;
  else list.push(nextItem);
  return list;
}

function openDepartmentLinkModal(department, link = null) {
  openModal(`
    <h3>${link ? "Weiterleitung bearbeiten" : "Weiterleitung hinzufügen"}</h3>
    <label>Titel<input id="linkTitle" value="${escapeHtml(link?.title || "")}"></label>
    <label>Link<input id="linkUrl" value="${escapeHtml(link?.url || "")}" placeholder="https://..."></label>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions"><button class="ghost-btn" data-close>Abbrechen</button><button class="blue-btn" id="saveDepartmentLink">Speichern</button></div>
  `, (modal) => {
    modal.querySelector("#saveDepartmentLink").addEventListener("click", async () => {
      try {
        const title = $("#linkTitle").value.trim();
        const url = $("#linkUrl").value.trim();
        if (!title || !url) throw new Error("Titel und Link sind erforderlich.");
        await saveDepartmentInfo(department, { links: upsertById(department.links, { id: link?.id, title, url }) });
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
  });
}

function openDepartmentPermitModal(department, permit = null) {
  openModal(`
    <h3>${permit ? "Sondergenehmigung bearbeiten" : "Sondergenehmigung hinzufügen"}</h3>
    <label>Vor- und Nachname<input id="permitName" value="${escapeHtml(permit?.name || "")}"></label>
    <label>Beschreibung<textarea id="permitDescription">${escapeHtml(permit?.description || "")}</textarea></label>
    <label>Gültig Bis<input id="permitValidUntil" type="date" value="${escapeHtml(permit?.validUntil || "")}"></label>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions"><button class="ghost-btn" data-close>Abbrechen</button><button class="blue-btn" id="saveDepartmentPermit">Speichern</button></div>
  `, (modal) => {
    modal.querySelector("#saveDepartmentPermit").addEventListener("click", async () => {
      try {
        const name = $("#permitName").value.trim();
        const description = $("#permitDescription").value.trim();
        const validUntil = $("#permitValidUntil").value;
        if (!name || !description || !validUntil) throw new Error("Alle Felder sind erforderlich.");
        await saveDepartmentInfo(department, { permits: upsertById(department.permits, { id: permit?.id, name, description, validUntil }) });
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
  });
}

function openDepartmentFactionModal(department, faction = null) {
  openModal(`
    <h3>${faction ? "Fraktion bearbeiten" : "Fraktion hinzufügen"}</h3>
    <label>Organisation<input id="factionOrganization" value="${escapeHtml(faction?.organization || "")}"></label>
    <label>Status
      <select id="factionStatus">
        ${["Normal", "Mittel", "Hoch"].map((status) => `<option ${faction?.status === status ? "selected" : ""}>${status}</option>`).join("")}
      </select>
    </label>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions"><button class="ghost-btn" data-close>Abbrechen</button><button class="blue-btn" id="saveDepartmentFaction">Speichern</button></div>
  `, (modal) => {
    modal.querySelector("#saveDepartmentFaction").addEventListener("click", async () => {
      try {
        const organization = $("#factionOrganization").value.trim();
        const status = $("#factionStatus").value;
        if (!organization) throw new Error("Organisation ist erforderlich.");
        await saveDepartmentInfo(department, { factions: upsertById(department.factions, { id: faction?.id, organization, status }) });
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
  });
}

async function deleteDepartmentInfoItem(department, key, id) {
  try {
    await saveDepartmentInfo(department, { [key]: (department[key] || []).filter((item) => item.id !== id) });
  } catch (error) {
    showNotify(error.message, "error");
  }
}

function openDepartmentMemberModal(department, member = null) {
  const availableUsers = state.users.filter((user) => member || !department.members.some((item) => item.userId === user.id));
  const myDepartmentPosition = department.members.find((item) => item.userId === state.currentUser.id)?.position;
  const allowedPositions = departmentPositionsFor(department).filter((position) => {
    if (hasRole("Direktion")) return true;
    if (position === "Direktion") return false;
    return positionPowerFor(department, position) < positionPowerFor(department, myDepartmentPosition);
  });
  const defaultDepartmentPosition = member?.position || (allowedPositions.includes("Anwärter") ? "Anwärter" : allowedPositions.includes("Mitglied") ? "Mitglied" : allowedPositions[0] || "Mitglied");
  openModal(`
    <h3>${member ? "Position bearbeiten" : "Person hinzufügen"}</h3>
    <p class="muted">Wählen Sie eine Person aus, die zu ${escapeHtml(department.name)} hinzugefügt werden soll.</p>
    ${member ? `<p><strong>${escapeHtml(fullName(member.user))}</strong></p>` : `<label>Person suchen<input id="departmentUserSearch" placeholder="Name oder DN suchen"></label>
    <label>Person auswählen<select id="departmentUserSelect">${availableUsers.map((user) => `<option value="${user.id}">${escapeHtml(fullName(user))} - DN ${escapeHtml(user.dn)} - ${escapeHtml(rankLabel(user.rank))}</option>`).join("")}</select></label>`}
    <label>Position auswählen
      <select id="departmentPositionSelect">
        ${(allowedPositions.length ? allowedPositions : ["Mitglied"]).map((position) => `<option ${defaultDepartmentPosition === position ? "selected" : ""}>${escapeHtml(position)}</option>`).join("")}
      </select>
    </label>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions">
      <button class="ghost-btn" data-close>Abbrechen</button>
      <button class="blue-btn" id="saveDepartmentMember">${member ? "Speichern" : "Person hinzufügen"}</button>
    </div>
  `, (modal) => {
    const search = modal.querySelector("#departmentUserSearch");
    const select = modal.querySelector("#departmentUserSelect");
    search?.addEventListener("input", () => {
      const term = search.value.toLowerCase();
      select.innerHTML = availableUsers
        .filter((user) => fullName(user).toLowerCase().includes(term) || String(user.dn).includes(term))
        .map((user) => `<option value="${user.id}">${escapeHtml(fullName(user))} - DN ${escapeHtml(user.dn)} - ${escapeHtml(rankLabel(user.rank))}</option>`)
        .join("");
    });
    modal.querySelector("#saveDepartmentMember").addEventListener("click", async () => {
      try {
        await api(member ? `/api/departments/${department.id}/members/${member.userId}` : `/api/departments/${department.id}/members`, {
          method: member ? "PATCH" : "POST",
          body: JSON.stringify({ userId: member?.userId || select.value, position: $("#departmentPositionSelect").value })
        });
        closeModal();
        await bootstrap();
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
  });
}

function openDepartmentManageModal(department) {
  openModal(`
    <h3>Personal verwalten</h3>
    <p class="muted">${escapeHtml(department.name)}</p>
    <button class="blue-btn department-add full" data-department-id="${escapeHtml(department.id)}">${iconSvg("Mitglieder")} Person hinzufügen</button>
    <div class="manage-member-list">
      ${department.members.length ? department.members.map((member) => `
        <div class="manage-member-row">
          <span><strong>${escapeHtml(fullName(member.user))}</strong><small>${escapeHtml(member.position)} · ${escapeHtml(rankLabel(member.user.rank))}</small></span>
          <span class="button-row">
            <button class="mini-icon edit-dept-member" data-department-id="${department.id}" data-user-id="${member.userId}" title="Position bearbeiten">${actionIcon("edit")}</button>
            <button class="mini-icon danger remove-dept-member" data-department-id="${department.id}" data-user-id="${member.userId}" title="Entfernen">${actionIcon("delete")}</button>
          </span>
        </div>
      `).join("") : `<p class="muted">Noch keine Mitglieder.</p>`}
    </div>
    <div class="modal-actions"><button class="ghost-btn" data-close>Schließen</button></div>
  `, (modal) => {
    modal.querySelector(".department-add")?.addEventListener("click", () => {
      closeModal();
      openDepartmentMemberModal(department);
    });
  });
}

function openRemoveDepartmentMemberModal(department, member) {
  openModal(`
    <h3>Person entfernen</h3>
    <p class="muted">${escapeHtml(fullName(member.user))} aus ${escapeHtml(department.name)} entfernen?</p>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions">
      <button class="ghost-btn" data-close>Abbrechen</button>
      <button class="red-btn" id="confirmRemoveDepartmentMember">Entfernen</button>
    </div>
  `, (modal) => {
    modal.querySelector("#confirmRemoveDepartmentMember").addEventListener("click", async () => {
      try {
        await api(`/api/departments/${department.id}/members/${member.userId}`, { method: "DELETE" });
        closeModal();
        await bootstrap();
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
  });
}

function openDepartmentNoteModal(department, note = null) {
  const isEdit = Boolean(note);
  openModal(`
    <h3>${isEdit ? "Notiz bearbeiten" : "Neue Notiz"}</h3>
    <label>Titel<input id="deptNoteTitle" value="${escapeHtml(note?.title || "")}"></label>
    <label>Priorität<select id="deptNotePriority">${["Leitung", "Info", "Mitglied"].map((priority) => `<option ${note?.priority === priority ? "selected" : ""}>${priority}</option>`).join("")}</select></label>
    <label>Text<textarea id="deptNoteText">${escapeHtml(note?.text || "")}</textarea></label>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions">
      <button class="ghost-btn" data-close>Abbrechen</button>
      <button class="blue-btn" id="saveDepartmentNote">${isEdit ? "Notiz aktualisieren" : "Notiz erstellen"}</button>
    </div>
  `, (modal) => {
    modal.querySelector("#saveDepartmentNote").addEventListener("click", async () => {
      try {
        await api(isEdit ? `/api/departments/${department.id}/notes/${note.id}` : `/api/departments/${department.id}/notes`, {
          method: isEdit ? "PATCH" : "POST",
          body: JSON.stringify({ title: $("#deptNoteTitle").value, priority: $("#deptNotePriority").value, text: $("#deptNoteText").value })
        });
        closeModal();
        await bootstrap();
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
  });
}

function openDepartmentMemberNoteModal(department, userId) {
  const member = department.members.find((item) => item.userId === userId);
  if (!member) return;
  openModal(`
    <h3>Interne Mitgliedsnotiz</h3>
    <p class="muted">${escapeHtml(fullName(member.user))} - ${escapeHtml(department.name)}</p>
    <label>Notiz<textarea id="deptMemberNoteText" rows="6" placeholder="Interne Notiz für die Abteilungsleitung..."></textarea></label>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions">
      <button class="ghost-btn" data-close>Abbrechen</button>
      <button class="blue-btn" id="saveDeptMemberNote">Speichern</button>
    </div>
  `, (modal) => {
    modal.querySelector("#saveDeptMemberNote").addEventListener("click", async () => {
      try {
        await api(`/api/departments/${department.id}/member-notes`, {
          method: "POST",
          body: JSON.stringify({ userId, text: $("#deptMemberNoteText").value })
        });
        closeModal();
        await bootstrap();
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
  });
}

function openDeleteDepartmentNoteModal(department, note) {
  openModal(`
    <h3>Notiz löschen</h3>
    <p class="muted">${escapeHtml(note.title)} wirklich löschen?</p>
    <p id="modalError" class="form-error"></p>
    <div class="modal-actions">
      <button class="ghost-btn" data-close>Abbrechen</button>
      <button class="red-btn" id="confirmDeleteDepartmentNote">Notiz löschen</button>
    </div>
  `, (modal) => {
    modal.querySelector("#confirmDeleteDepartmentNote").addEventListener("click", async () => {
      try {
        await api(`/api/departments/${department.id}/notes/${note.id}`, { method: "DELETE" });
        closeModal();
        await bootstrap();
      } catch (error) {
        $("#modalError").textContent = error.message;
      }
    });
  });
}

document.addEventListener("click", (event) => {
  const editDeptMember = event.target.closest(".edit-dept-member");
  if (editDeptMember) {
    const department = state.departments.find((item) => item.id === editDeptMember.dataset.departmentId);
    const member = department?.members.find((item) => item.userId === editDeptMember.dataset.userId);
    if (department && member) openDepartmentMemberModal(department, member);
    return;
  }

  const removeDeptMember = event.target.closest(".remove-dept-member");
  if (removeDeptMember) {
    const department = state.departments.find((item) => item.id === removeDeptMember.dataset.departmentId);
    const member = department?.members.find((item) => item.userId === removeDeptMember.dataset.userId);
    if (department && member) openRemoveDepartmentMemberModal(department, member);
    return;
  }

  const editDeptNote = event.target.closest(".edit-dept-note");
  if (editDeptNote) {
    const department = state.departments.find((item) => item.id === editDeptNote.dataset.departmentId);
    const note = department?.notes.find((item) => item.id === editDeptNote.dataset.noteId);
    if (department && note) openDepartmentNoteModal(department, note);
    return;
  }

  const deleteDeptNote = event.target.closest(".delete-dept-note");
  if (deleteDeptNote) {
    const department = state.departments.find((item) => item.id === deleteDeptNote.dataset.departmentId);
    const note = department?.notes.find((item) => item.id === deleteDeptNote.dataset.noteId);
    if (department && note) openDeleteDepartmentNoteModal(department, note);
    return;
  }

  const editNoteButton = event.target.closest(".edit-note");
  if (editNoteButton) {
    const note = state.notes.find((item) => item.id === editNoteButton.dataset.noteId);
    if (note) openNoteModal(note);
    return;
  }

  const deleteNoteButton = event.target.closest(".delete-note");
  if (deleteNoteButton) {
    openDeleteNoteModal(deleteNoteButton.dataset.noteId);
    return;
  }

  const removeButton = event.target.closest(".remove-duty");
  if (removeButton) {
    const entry = state.duty.find((item) => item.userId === removeButton.dataset.userId);
    const user = entry?.user || state.users.find((item) => item.id === removeButton.dataset.userId);
    openModal(`
      <h3>Person austragen</h3>
      <p class="muted">${escapeHtml(fullName(user))} aus dem Dienst austragen?</p>
      <p id="modalError" class="form-error"></p>
      <div class="modal-actions">
        <button class="ghost-btn" data-close>Abbrechen</button>
        <button class="red-btn" id="confirmRemovePerson">Person austragen</button>
      </div>
    `, (modal) => {
      modal.querySelector("#confirmRemovePerson").addEventListener("click", async () => {
        try {
          await api(`/api/duty/stop/${removeButton.dataset.userId}`, { method: "POST", body: "{}" });
          closeModal();
          await bootstrap();
        } catch (error) {
          $("#modalError").textContent = error.message;
        }
      });
    });
  }
});

$("#loginForm").addEventListener("submit", async (event) => {
  event.preventDefault();
  $("#loginError").textContent = "";
  try {
    const data = await api("/api/login", {
      method: "POST",
      body: JSON.stringify({ name: $("#loginName").value, password: $("#loginPassword").value })
    });
    state.token = data.token;
    storeAuthToken(state.token);
    await bootstrap();
    await linkPendingDiscordAccount();
  } catch (error) {
    $("#loginError").textContent = error.message;
  }
});

$("#discordLoginBtn")?.addEventListener("click", () => startDiscordOAuth("login"));

async function logout() {
  try {
    await api("/api/logout", { method: "POST", body: "{}" });
  } catch (_error) {
  } finally {
    clearAuthToken();
    state.token = null;
    showLogin();
  }
}

$("#logoutBtn")?.addEventListener("click", logout);

installInspectGuard();

document.addEventListener("click", (event) => {
  document.querySelectorAll(".exam-user-picker.open").forEach((picker) => {
    if (!picker.contains(event.target)) picker.classList.remove("open");
  });
});

async function initApp() {
  const handledDiscordRedirect = await handleDiscordOAuthRedirect();
  if (handledDiscordRedirect) return;
  if (state.token) {
    bootstrap().catch(() => {
      clearAuthToken();
      showLogin();
    });
  } else {
    showLogin();
  }
}

initApp();
