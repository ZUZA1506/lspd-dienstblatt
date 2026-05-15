const crypto = require("crypto");
const fs = require("fs");
const https = require("https");
const path = require("path");
const express = require("express");

const app = express();
const PORT = Number(process.env.PORT || 3000);
const ROOT = path.resolve(__dirname, "..");
const STORAGE_DIR = process.env.DATA_DIR ? path.resolve(process.env.DATA_DIR) : path.join(ROOT, "storage");
const DB_FILE = path.join(STORAGE_DIR, "dienstblatt.json");
const PUBLIC_DIR = path.join(ROOT, "public");
const DEFAULT_PASSWORD = "LSPD12345";

const roles = ["User", "Supervisor", "Direktion", "IT", "IT-Leitung"];
const rolePower = { User: 1, Supervisor: 2, Direktion: 3, IT: 4, "IT-Leitung": 5 };
const ranks = Array.from({ length: 13 }, (_, index) => ({
  value: index,
  label: `Template ${index} - Rang ${index}`
}));

function defaultRanks() {
  return ranks.map((rank) => ({ ...rank }));
}

function defaultUprankRules() {
  return ranks.slice(1).map((rank) => ({
    targetRank: rank.value,
    minDays: rank.value === 1 ? 7 : 14,
    trainings: rank.value === 1 ? ["EST"] : [],
    specialOnly: rank.value >= 7
  }));
}

function defaultDepartments() {
  return [
    makeDepartment("direktion", "Direktion", "LSPD Direktion und administrative Leitung", "Offen"),
    makeDepartment("training-recruitment", "Training / Recruitment", "Ausbildung, Recruiting und Lernkontrollen", "Offen"),
    makeDepartment("department-corruptions", "Department of Corruptions", "Interne Ermittlungen und Korruptionsdelikte", "Offen"),
    makeDepartment("metro-taskforce", "Metro Taskforce", "Spezialeinsätze und operative Taskforce", "Offen"),
    makeDepartment("swat", "SWAT", "Taktische Einsätze und Zugriffslagen", "Offen")
  ];
}

function makeDepartment(id, name, description, applicationStatus) {
  return {
    id,
    name,
    description,
    applicationStatus,
    requirements: "Voraussetzungen werden später ergänzt.",
    rightsText: "",
    links: [],
    permits: [],
    factions: [],
    positions: [...departmentPositions],
    members: [],
    notes: [],
    memberNotes: []
  };
}

function defaultPermissions() {
  return {
    pages: {},
    actions: {
      editDefcon: { roles: ["Supervisor", "Direktion", "IT", "IT-Leitung"], ranks: [], users: [] },
      manageNotes: { roles: ["Supervisor", "Direktion", "IT", "IT-Leitung"], ranks: [], users: [] },
      stopAllDuty: { roles: ["Direktion", "IT", "IT-Leitung"], ranks: [], users: [] },
      manageInformation: { roles: ["Direktion", "IT", "IT-Leitung"], ranks: [], users: [] },
      manageDutyHours: { roles: ["Direktion", "IT", "IT-Leitung"], ranks: [], users: [] },
      manageDepartments: { roles: ["Direktion", "IT", "IT-Leitung"], ranks: [], users: [] },
      manageMembers: { roles: ["Direktion", "IT", "IT-Leitung"], ranks: [], users: [] },
      viewLogs: { roles: ["Direktion", "IT", "IT-Leitung"], ranks: [], users: [] }
    }
  };
}

function normalizePermissionRule(rule = {}) {
  return {
    all: Boolean(rule.all),
    roles: Array.isArray(rule.roles) ? rule.roles.filter((role) => roles.includes(role)) : [],
    ranks: Array.isArray(rule.ranks) ? rule.ranks.map(Number).filter((rank) => Number.isInteger(rank)) : [],
    users: Array.isArray(rule.users) ? rule.users.map(String).filter(Boolean) : [],
    departments: Array.isArray(rule.departments) ? rule.departments.map(String).filter(Boolean) : [],
    positions: Array.isArray(rule.positions) ? rule.positions.map(String).filter(Boolean) : []
  };
}

function normalizePermissions(value = {}) {
  const defaults = defaultPermissions();
  const pages = value.pages && typeof value.pages === "object" ? value.pages : {};
  const actions = value.actions && typeof value.actions === "object" ? value.actions : {};
  return {
    pages: Object.fromEntries(Object.entries(pages).map(([key, rule]) => [key, normalizePermissionRule(rule)])),
    actions: {
      ...Object.fromEntries(Object.entries(defaults.actions).map(([key, rule]) => [key, normalizePermissionRule(actions[key] || rule)])),
      ...Object.fromEntries(Object.entries(actions).filter(([key]) => !defaults.actions[key]).map(([key, rule]) => [key, normalizePermissionRule(rule)]))
    }
  };
}

const departmentPositions = ["Direktion", "Leitung", "Stv. Leitung", "Mitglied", "Anwärter"];
const positionPower = { "Direktion": 5, "Leitung": 4, "Stv. Leitung": 3, "Mitglied": 2, "Anwärter": 1 };
const trainingNames = ["EST", "Wissen", "Fahren", "Schießen", "Verhalten", "Undercover", "Wanted", "EL", "Officer Prüfung", "Prak. VHF", "Prak. EL I", "Führung", "Prak. EL II", "Air Support", "Riot", "Coquette"];

function nowIso() {
  return new Date().toISOString();
}

function todayIso() {
  return new Date().toISOString().slice(0, 10);
}

function actorName(actor) {
  return actor ? `${actor.firstName} ${actor.lastName}`.trim() : "System";
}

function hashPassword(password) {
  return crypto.createHash("sha256").update(password).digest("hex");
}

function makeId(prefix) {
  return `${prefix}_${crypto.randomUUID()}`;
}

function ensureStorage() {
  fs.mkdirSync(STORAGE_DIR, { recursive: true });
  if (fs.existsSync(DB_FILE)) return;

  const adminId = makeId("user");
  const createdAt = nowIso();
  const seed = {
    users: [
      {
        id: adminId,
        firstName: "Admin",
        lastName: "Direktion",
        phone: "0000",
        dn: "1",
        rank: 12,
        role: "IT",
        departments: ["Direktion"],
        trainings: Object.fromEntries(trainingNames.map((training) => [training, true])),
        joinedAt: todayIso(),
        lastPromotionAt: todayIso(),
        passwordHash: hashPassword(DEFAULT_PASSWORD),
        avatarUrl: "",
        locked: false,
        createdAt,
        updatedAt: createdAt
      }
    ],
    sessions: [],
    settings: {
      defcon: "DEFCON 3",
      defconUpdatedBy: "System",
      defconUpdatedAt: createdAt,
      ranks: defaultRanks(),
      navLabels: {},
      customPages: [],
      pageOrder: [],
      departments: defaultDepartments(),
      informationText: "Hier können später zentrale Informationen für alle Officer gepflegt werden.",
      applicationStatus: "Offen",
      calendarEvents: [],
      seizures: [],
      fluctuation: [],
      uprankRules: defaultUprankRules(),
      uprankAdjustments: [],
      permissions: defaultPermissions(),
      devMode: false,
      restartTimes: [],
      restartLastRun: {}
    },
    notes: [],
    duty: [],
    dutyHistory: [],
    logs: [],
    disciplinary: []
  };

  fs.writeFileSync(DB_FILE, JSON.stringify(seed, null, 2));
}

function readDb() {
  ensureStorage();
  const db = JSON.parse(fs.readFileSync(DB_FILE, "utf8"));
  if (!db.users.some((user) => user.role === "IT-Leitung")) {
    const firstItUser = db.users.find((user) => user.role === "IT");
    if (firstItUser) firstItUser.role = "IT-Leitung";
  }
  db.users.forEach((user) => {
    if (!user.baseRole) user.baseRole = ["IT", "IT-Leitung"].includes(user.role) ? "Direktion" : user.role || "User";
    user.teamler = Boolean(user.teamler);
    user.trainingMeta = user.trainingMeta && typeof user.trainingMeta === "object" ? user.trainingMeta : {};
    if (user.trainings?.Schiessen && !user.trainings["Schießen"]) {
      user.trainings["Schießen"] = true;
      delete user.trainings.Schiessen;
    }
    user.trainings = { ...Object.fromEntries(trainingNames.map((training) => [training, false])), ...(user.trainings || {}) };
  });
  db.settings = db.settings || {};
  db.settings.ranks = Array.isArray(db.settings.ranks) && db.settings.ranks.length ? db.settings.ranks : defaultRanks();
  db.settings.navLabels = db.settings.navLabels || {};
  db.settings.customPages = Array.isArray(db.settings.customPages) ? db.settings.customPages : [];
  db.settings.pageOrder = Array.isArray(db.settings.pageOrder) ? db.settings.pageOrder : [];
  db.settings.departments = normalizeDepartments(db.settings.departments);
  db.settings.informationText = db.settings.informationText || "Hier können später zentrale Informationen für alle Officer gepflegt werden.";
  db.settings.applicationStatus = db.settings.applicationStatus || "Offen";
  if (typeof db.settings.defconText !== "string") db.settings.defconText = "Automatisch / Manuell aktualisierbar";
  db.settings.calendarEvents = Array.isArray(db.settings.calendarEvents) ? db.settings.calendarEvents : [];
  db.settings.seizures = Array.isArray(db.settings.seizures) ? db.settings.seizures : [];
  db.settings.uprankRules = normalizeUprankRules(db.settings.uprankRules);
  db.settings.uprankAdjustments = Array.isArray(db.settings.uprankAdjustments) ? db.settings.uprankAdjustments : [];
  db.settings.permissions = normalizePermissions(db.settings.permissions);
  db.settings.devMode = Boolean(db.settings.devMode);
  db.settings.restartTimes = Array.isArray(db.settings.restartTimes) ? db.settings.restartTimes : [];
  db.settings.restartLastRun = db.settings.restartLastRun && typeof db.settings.restartLastRun === "object" ? db.settings.restartLastRun : {};
  db.settings.informationRightsText = String(db.settings.informationRightsText || "");
  db.settings.informationLinks = Array.isArray(db.settings.informationLinks) ? db.settings.informationLinks : [];
  db.settings.informationDocs = Array.isArray(db.settings.informationDocs) ? db.settings.informationDocs : [];
  db.settings.informationDocChanges = Array.isArray(db.settings.informationDocChanges) ? db.settings.informationDocChanges : [];
  db.settings.informationPermits = Array.isArray(db.settings.informationPermits) ? db.settings.informationPermits : [];
  db.settings.informationFactions = Array.isArray(db.settings.informationFactions) ? db.settings.informationFactions : [];
  db.settings.fluctuation = Array.isArray(db.settings.fluctuation) ? db.settings.fluctuation : [];
  db.dutyHistory = Array.isArray(db.dutyHistory) ? db.dutyHistory : [];
  db.logs = Array.isArray(db.logs) ? db.logs : [];
  db.disciplinary = Array.isArray(db.disciplinary) ? db.disciplinary : [];
  db.users.forEach((user) => {
    if (!user.accountStatus) {
      const latestStatusEntry = db.disciplinary.find((entry) => entry.userId === user.id && ["Suspendierung", "Sperre", "Entsperrt", "Entlassen"].includes(entry.type));
      user.accountStatus = user.terminated ? "Entlassen" : latestStatusEntry?.type === "Suspendierung" ? "Suspendiert" : user.locked ? "Gesperrt" : "Aktiv";
    }
  });
  return db;
}

function normalizeDepartments(existingDepartments) {
  const defaults = defaultDepartments();
  const existing = Array.isArray(existingDepartments) ? existingDepartments : [];
  const normalizedDefaults = defaults.map((department) => {
    const stored = existing.find((item) => item.id === department.id || item.name === department.name);
    return {
      ...department,
      ...(stored || {}),
      rightsText: String(stored?.rightsText || department.rightsText || ""),
      links: Array.isArray(stored?.links) ? stored.links : department.links,
      permits: Array.isArray(stored?.permits) ? stored.permits : department.permits,
      factions: Array.isArray(stored?.factions) ? stored.factions : department.factions,
      positions: normalizeDepartmentPositions(stored?.positions || department.positions),
      members: Array.isArray(stored?.members) ? stored.members : department.members,
      notes: Array.isArray(stored?.notes) ? stored.notes : department.notes,
      memberNotes: Array.isArray(stored?.memberNotes) ? stored.memberNotes : department.memberNotes
    };
  });
  const defaultIds = new Set(defaults.map((department) => department.id));
  const custom = existing
    .filter((department) => department?.id && !defaultIds.has(department.id))
    .map((department) => ({
      ...makeDepartment(department.id, department.name || "Neue Abteilung", department.description || "Leeres Abteilungsblatt", department.applicationStatus || "Offen"),
      ...department,
      rightsText: String(department.rightsText || ""),
      links: Array.isArray(department.links) ? department.links : [],
      permits: Array.isArray(department.permits) ? department.permits : [],
      factions: Array.isArray(department.factions) ? department.factions : [],
      positions: normalizeDepartmentPositions(department.positions || departmentPositions),
      members: Array.isArray(department.members) ? department.members : [],
      notes: Array.isArray(department.notes) ? department.notes : [],
      memberNotes: Array.isArray(department.memberNotes) ? department.memberNotes : []
    }));
  return [...normalizedDefaults, ...custom];
}

function normalizeDepartmentPositions(value) {
  const incoming = Array.isArray(value) ? value.map((item) => String(item || "").trim()).filter(Boolean) : [];
  return incoming.length ? [...new Set(incoming)] : [...departmentPositions];
}

function departmentPositionsFor(department) {
  return normalizeDepartmentPositions(department?.positions || departmentPositions);
}

function positionPowerFor(department, position) {
  const positions = departmentPositionsFor(department);
  const index = positions.indexOf(position);
  if (index === -1) return 0;
  return positions.length - index;
}

function normalizeUprankRules(existingRules) {
  const existing = Array.isArray(existingRules) ? existingRules : [];
  const defaults = defaultUprankRules();
  return defaults.map((rule) => {
    const stored = existing.find((item) => Number(item.targetRank) === Number(rule.targetRank));
    const trainings = Array.isArray(stored?.trainings) ? stored.trainings.filter((training) => trainingNames.includes(training)) : rule.trainings;
    return {
      targetRank: rule.targetRank,
      minDays: Math.max(0, Number.parseInt(stored?.minDays ?? rule.minDays, 10) || 0),
      trainings,
      specialOnly: Boolean(stored?.specialOnly ?? rule.specialOnly)
    };
  });
}

function writeDb(db) {
  fs.writeFileSync(DB_FILE, JSON.stringify(db, null, 2));
}

function publicUser(user) {
  if (!user) return null;
  const { passwordHash, ...safeUser } = user;
  return {
    ...safeUser,
    avatarUrl: user.avatarUrl || "",
    fullName: `${user.firstName} ${user.lastName}`.trim()
  };
}

function logFluctuation(db, user, type, actor) {
  db.settings.fluctuation.unshift({
    id: makeId("fluctuation"),
    type,
    userId: user.id,
    name: `${user.firstName} ${user.lastName}`.trim(),
    dn: user.dn,
    rank: user.rank,
    actorName: actorName(actor),
    reason: "",
    createdAt: nowIso()
  });
}

function logAction(db, actor, action, target = "", details = {}) {
  db.logs = Array.isArray(db.logs) ? db.logs : [];
  db.logs.unshift({
    id: makeId("log"),
    action,
    target,
    actorId: actor?.id || "",
    actorName: actorName(actor),
    details,
    createdAt: nowIso()
  });
}

function logDisciplinary(db, user, type, reason, actor) {
  db.disciplinary = Array.isArray(db.disciplinary) ? db.disciplinary : [];
  db.disciplinary.unshift({
    id: makeId("disciplinary"),
    type,
    userId: user.id,
    name: `${user.firstName} ${user.lastName}`.trim(),
    dn: user.dn,
    rank: user.rank,
    actorName: actorName(actor),
    reason,
    createdAt: nowIso()
  });
}

function isActiveStrike(entry) {
  const isStrike = entry.type === "Strike" || (entry.type === "Sanktion" && entry.sanctionType === "Strike");
  if (!isStrike || entry.archivedAt) return false;
  if (entry.expiresAt && new Date(entry.expiresAt) <= new Date()) return false;
  return true;
}

function activeStrikeCount(entries, userId) {
  return entries
    .filter((entry) => entry.userId === userId && isActiveStrike(entry))
    .reduce((sum, entry) => sum + Math.max(1, Number(entry.strikeCount || 1)), 0);
}

function setAccountStatus(user, status) {
  user.accountStatus = status;
  user.locked = ["Gesperrt", "Suspendiert", "Entlassen"].includes(status);
}

function requireAuth(req, res, next) {
  const token = (req.headers.authorization || "").replace(/^Bearer\s+/i, "");
  const db = readDb();
  const session = db.sessions.find((item) => item.token === token);
  if (!session) return res.status(401).json({ error: "Nicht angemeldet." });

  const user = db.users.find((item) => item.id === session.userId);
  if (!user || user.locked) return res.status(401).json({ error: "Account gesperrt oder nicht gefunden." });

  req.db = db;
  req.session = session;
  req.user = user;
  next();
}

function requireRole(minRole) {
  return (req, res, next) => {
    if ((rolePower[req.user.role] || 0) < rolePower[minRole]) {
      return res.status(403).json({ error: "Keine Berechtigung." });
    }
    next();
  };
}

function hasPermission(user, db, area, key, fallbackRole = "IT") {
  if ((rolePower[user.role] || 0) >= rolePower.IT) return true;
  if (user.role === "Direktion" && key !== "IT") return true;
  const rule = db.settings.permissions?.[area]?.[key];
  if (!rule) return (rolePower[user.role] || 0) >= (rolePower[fallbackRole] || 99);
  const departmentMatch = (rule.departments || []).some((departmentId) => {
    const department = db.settings.departments.find((item) => item.id === departmentId);
    return department?.members.some((member) => member.userId === user.id);
  });
  const positionMatch = (rule.positions || []).some((positionKey) => {
    const [departmentId, position] = String(positionKey).split(":");
    const department = db.settings.departments.find((item) => item.id === departmentId);
    return department?.members.some((member) => member.userId === user.id && member.position === position);
  });
  return rule.all || rule.users.includes(user.id) || rule.roles.includes(user.role) || rule.ranks.includes(Number(user.rank)) || departmentMatch || positionMatch;
}

function requirePermission(area, key, fallbackRole = "IT") {
  return (req, res, next) => {
    if (!hasPermission(req.user, req.db, area, key, fallbackRole)) {
      return res.status(403).json({ error: "Keine Berechtigung." });
    }
    next();
  };
}

function getDepartment(db, departmentId) {
  return db.settings.departments.find((department) => department.id === departmentId);
}

function syncDirektionMembership(db, user, options = {}) {
  const department = getDepartment(db, "direktion");
  if (!department || !user) return;
  const hasDirektionRole = !user.terminated && user.role === "Direktion";
  if (options.roleAssigned) user.direktionManualRemoved = false;
  if (hasDirektionRole) {
    if (user.direktionManualRemoved) return;
    if (!department.members.some((member) => member.userId === user.id)) {
      department.members.push({
        userId: user.id,
        position: "Direktion",
        joinedAt: todayIso(),
        positionSince: todayIso(),
        autoRoleDirektion: true
      });
    }
    return;
  }
  const beforeLength = department.members.length;
  department.members = department.members.filter((member) => member.userId !== user.id);
  if (beforeLength !== department.members.length) user.direktionManualRemoved = false;
}

function isDepartmentManager(user, department, db = null) {
  if (!department) return false;
  if ((rolePower[user.role] || 0) >= rolePower.IT || user.role === "Direktion") return true;
  if (db && hasPermission(user, db, "actions", `departmentManage:${department.id}`, "IT")) return true;
  const membership = department.members.find((member) => member.userId === user.id);
  return positionPowerFor(department, membership?.position) >= positionPowerFor(department, "Stv. Leitung");
}

function canManageDepartmentAction(user, department, db, action) {
  if (!department) return false;
  const key = `${action}:${department.id}`;
  const rule = db?.settings?.permissions?.actions?.[key];
  if (rule) return hasPermission(user, db, "actions", key, "IT");
  if (action === "departmentLeadership") {
    if ((rolePower[user.role] || 0) >= rolePower.Direktion) return true;
    const membership = department.members.find((member) => member.userId === user.id);
    return positionPowerFor(department, membership?.position) >= positionPowerFor(department, "Leitung");
  }
  return isDepartmentManager(user, department, db);
}

function canAssignDepartmentPosition(user, department, position, db = null) {
  if (!departmentPositionsFor(department).includes(position)) return false;
  if ((rolePower[user.role] || 0) >= rolePower.Direktion) return true;
  if (db && canManageDepartmentAction(user, department, db, "departmentMembers")) return position !== "Direktion";
  if (position === "Direktion") return false;
  const membership = department.members.find((member) => member.userId === user.id);
  const actorPower = positionPowerFor(department, membership?.position);
  return actorPower > positionPowerFor(department, position);
}

function canSeeDepartmentPage(user, department, db = null) {
  if (!department) return false;
  if ((rolePower[user.role] || 0) >= rolePower.Direktion) return true;
  if (db && hasPermission(user, db, "pages", `dept:${department.id}`, "IT")) return true;
  return department.members.some((member) => member.userId === user.id);
}

function publicDepartment(department, db, currentUser) {
  const dutyIds = new Set(db.duty.map((entry) => entry.userId));
  return {
    ...department,
    canManage: isDepartmentManager(currentUser, department, db),
    canOpen: canSeeDepartmentPage(currentUser, department, db),
    members: department.members
      .map((member) => {
        const user = db.users.find((item) => item.id === member.userId);
        return user && !user.terminated ? { ...member, user: publicUser(user), isOnDuty: dutyIds.has(user.id) } : null;
      })
      .filter(Boolean)
      .sort((a, b) => positionPowerFor(department, b.position) - positionPowerFor(department, a.position) || b.user.rank - a.user.rank)
  };
}

function canGrantItRoles(actor) {
  return actor?.role === "IT-Leitung";
}

function protectItRoleChange(actor, existingRole, requestedRole) {
  const before = existingRole || "User";
  const next = roles.includes(requestedRole) ? requestedRole : before;
  const touchesItRole = ["IT", "IT-Leitung"].includes(before) || ["IT", "IT-Leitung"].includes(next);
  if (touchesItRole && before !== next && !canGrantItRoles(actor)) {
    return { error: "Nur die IT-Leitung darf IT- oder IT-Leitung-Rollen vergeben oder entfernen." };
  }
  return { role: next };
}

function validateDigits(value, field) {
  if (!/^\d+$/.test(String(value || ""))) {
    return `${field} darf nur Zahlen enthalten.`;
  }
  return null;
}

function normalizeUserInput(body, existingUser) {
  const firstName = String(body.firstName || "").trim();
  const lastName = String(body.lastName || "").trim();
  const phone = String(body.phone || "").trim();
  const dn = String(body.dn || "").trim();
  const rankMatch = String(body.rank ?? "").match(/\d+/);
  const rank = rankMatch ? Number(rankMatch[0]) : NaN;
  const role = roles.includes(body.role) ? body.role : existingUser?.role || "User";
  const requestedBaseRole = String(body.baseRole || "").trim();
  const baseRole = roles.includes(requestedBaseRole) && !["IT", "IT-Leitung"].includes(requestedBaseRole)
    ? requestedBaseRole
    : existingUser?.baseRole || (["IT", "IT-Leitung"].includes(role) ? "Direktion" : role);
  const departments = Array.isArray(body.departments) ? body.departments.map(String) : existingUser?.departments || [];
  const trainings = body.trainings && typeof body.trainings === "object" ? body.trainings : existingUser?.trainings || {};
  const teamler = Boolean(body.teamler);
  const joinedAt = String(body.joinedAt || existingUser?.joinedAt || todayIso()).slice(0, 10);

  if (!firstName || !lastName || !phone || !dn || Number.isNaN(rank)) {
    return { error: "Name, Nachname, Telefon, DN und Rang sind Pflichtfelder." };
  }

  const dnError = validateDigits(dn, "DN");
  if (dnError) return { error: dnError };
  if (rank < 0) return { error: "Rang muss mindestens 0 sein." };

  return {
    value: {
      firstName,
      lastName,
      phone,
      dn,
      rank,
      role,
      baseRole,
      teamler,
      joinedAt,
      departments,
      trainings
    }
  };
}

function dnConflictMessage(user) {
  const status = user.terminated ? "Entlassen" : user.accountStatus || (user.locked ? "Gesperrt" : "Aktiv");
  const dateText = user.terminated && user.termination?.terminatedAt ? `, entlassen am ${new Date(user.termination.terminatedAt).toLocaleString("de-DE")}` : "";
  return `${actorName(user)} (${status}${dateText})`;
}

function resolveDnConflict(db, currentUserId, dn, overwriteDn) {
  const holder = db.users.find((item) => item.id !== currentUserId && item.dn === dn);
  if (!holder) return null;
  if (!holder.terminated) {
    return { error: `Diese Dienstnummer ist bereits durch ${dnConflictMessage(holder)} vergeben.` };
  }
  if (!overwriteDn) {
    return { error: `Diese Dienstnummer ist bereits durch ${dnConflictMessage(holder)} vergeben. Zum Überschreiben bitte bestätigen.` };
  }
  holder.dn = "";
  holder.updatedAt = nowIso();
  return { holder };
}

function rankText(db, rank) {
  return (db.settings.ranks || []).find((item) => Number(item.value) === Number(rank))?.label || `Rang ${rank}`;
}

function userChangeSummary(db, before, after) {
  const changes = [];
  const fields = [
    ["firstName", "Vorname"],
    ["lastName", "Nachname"],
    ["phone", "Telefon"],
    ["dn", "Dienstnummer"],
    ["joinedAt", "Einstellungsdatum"],
    ["role", "Rolle"]
  ];
  fields.forEach(([key, label]) => {
    if (String(before?.[key] ?? "") !== String(after?.[key] ?? "")) changes.push(`${label}: ${before?.[key] || "-"} -> ${after?.[key] || "-"}`);
  });
  if (Number(before?.rank) !== Number(after?.rank)) changes.push(`Rang: ${rankText(db, before?.rank)} -> ${rankText(db, after?.rank)}`);
  trainingNames.forEach((training) => {
    const had = Boolean(before?.trainings?.[training]);
    const has = Boolean(after?.trainings?.[training]);
    if (had !== has) changes.push(`Ausbildung ${training} ${has ? "hinzugefügt" : "entfernt"}`);
  });
  return changes.join("; ");
}

function updateTrainingMeta(user, beforeTrainings, afterTrainings, actor) {
  user.trainingMeta = user.trainingMeta && typeof user.trainingMeta === "object" ? user.trainingMeta : {};
  trainingNames.forEach((training) => {
    const had = Boolean(beforeTrainings?.[training]);
    const has = Boolean(afterTrainings?.[training]);
    if (has && !had) {
      user.trainingMeta[training] = { completedAt: nowIso(), completedBy: actorName(actor) };
    }
    if (!has && had) delete user.trainingMeta[training];
  });
}

function daysSince(dateValue) {
  const time = new Date(dateValue || Date.now()).getTime();
  return Math.max(0, Math.floor((Date.now() - time) / 86400000));
}

function evaluateUprank(db, user, targetRank) {
  const rule = normalizeUprankRules(db.settings.uprankRules).find((item) => Number(item.targetRank) === Number(targetRank)) || { minDays: 14, trainings: [], specialOnly: targetRank >= 7 };
  const adjustments = Array.isArray(db.settings.uprankAdjustments) ? db.settings.uprankAdjustments : [];
  const reduction = adjustments
    .filter((item) => item.userId === user.id && Number(item.targetRank) === Number(targetRank) && item.type === "Verkürzung")
    .reduce((sum, item) => sum + Number(item.days || 0), 0);
  const effectiveDays = Math.max(0, Number(rule.minDays || 0) - reduction);
  const missingDays = Math.max(0, effectiveDays - daysSince(user.lastPromotionAt || user.joinedAt));
  const missingTrainings = (rule.trainings || []).filter((training) => !user.trainings?.[training]);
  const hasSpecial = adjustments.some((item) => item.userId === user.id && Number(item.targetRank) === Number(targetRank) && item.type === "Sonderuprank");
  return {
    rule,
    missingDays,
    missingTrainings,
    hasSpecial,
    regularReady: missingDays === 0 && missingTrainings.length === 0
  };
}

app.use(express.json({ limit: "25mb" }));
app.use((req, res, next) => {
  const ext = path.extname(req.path).toLowerCase();
  if (req.path.startsWith("/api/") || !ext) {
    res.setHeader("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
    res.setHeader("Pragma", "no-cache");
    res.setHeader("Expires", "0");
  }
  next();
});
app.use(express.static(PUBLIC_DIR, {
  etag: true,
  lastModified: true,
  maxAge: "1d",
  setHeaders: (res, filePath) => {
    if (filePath.endsWith(".html")) {
      res.setHeader("Content-Type", "text/html; charset=utf-8");
      res.setHeader("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
      res.setHeader("Pragma", "no-cache");
      res.setHeader("Expires", "0");
    }
    if (filePath.endsWith(".js")) res.setHeader("Content-Type", "application/javascript; charset=utf-8");
    if (filePath.endsWith(".css")) res.setHeader("Content-Type", "text/css; charset=utf-8");
    if (/\.(png|jpg|jpeg|webp|gif|svg|ico)$/i.test(filePath)) {
      res.setHeader("Cache-Control", "public, max-age=31536000, immutable");
    }
  }
}));

app.get("/api/evidence-preview", (req, res) => {
  const url = String(req.query.url || "");
  if (!/^https:\/\/(?:www\.)?prnt\.sc\//i.test(url)) return res.status(400).end();
  https.get(url, { headers: { "User-Agent": "Mozilla/5.0" } }, (remote) => {
    let body = "";
    remote.setEncoding("utf8");
    remote.on("data", (chunk) => { body += chunk; });
    remote.on("end", () => {
      const match = body.match(/<meta[^>]+property=["']og:image["'][^>]+content=["']([^"']+)["']/i)
        || body.match(/<meta[^>]+content=["']([^"']+)["'][^>]+property=["']og:image["']/i);
      const imageUrl = match?.[1] || "";
      if (!/^https?:\/\//i.test(imageUrl) || /st\.prntscr\.com\//i.test(imageUrl)) return res.status(404).end();
      res.redirect(imageUrl);
    });
  }).on("error", () => res.status(404).end());
});

app.post("/api/login", (req, res) => {
  const db = readDb();
  const name = String(req.body.name || "").trim().toLowerCase();
  const passwordHash = hashPassword(String(req.body.password || ""));
  const user = db.users.find((item) => {
    const fullName = `${item.firstName} ${item.lastName}`.trim().toLowerCase();
    return !item.locked && fullName === name && item.passwordHash === passwordHash;
  });

  if (!user) return res.status(401).json({ error: "Login fehlgeschlagen." });

  const token = crypto.randomBytes(32).toString("hex");
  db.sessions.push({ token, userId: user.id, createdAt: nowIso() });
  logAction(db, user, "Login", `${user.firstName} ${user.lastName}`.trim());
  writeDb(db);
  res.json({ token, user: publicUser(user) });
});

app.post("/api/logout", requireAuth, (req, res) => {
  req.db.sessions = req.db.sessions.filter((item) => item.token !== req.session.token);
  logAction(req.db, req.user, "Logout", `${req.user.firstName} ${req.user.lastName}`.trim());
  writeDb(req.db);
  res.json({ ok: true });
});

app.post("/api/security/inspect-attempt", (req, res) => {
  const db = readDb();
  const token = (req.headers.authorization || "").replace(/^Bearer\s+/i, "");
  const session = db.sessions.find((item) => item.token === token);
  const user = db.users.find((item) => item.id === session?.userId);
  const actor = user || { firstName: "Unbekannter", lastName: "Besucher" };
  const reason = String(req.body.reason || "Untersuchen versucht").slice(0, 120);
  const page = String(req.body.page || "").slice(0, 80);
  logAction(db, actor, "Untersuchen blockiert", page || "Website", {
    reason,
    userAgent: String(req.headers["user-agent"] || "").slice(0, 180)
  });
  writeDb(db);
  res.json({ ok: true });
});

app.get("/api/bootstrap", requireAuth, (req, res) => {
  const sortedUsers = [...req.db.users].filter((user) => !user.terminated).sort((a, b) => b.rank - a.rank || a.lastName.localeCompare(b.lastName));
  const archivedUsers = [...req.db.users].filter((user) => user.terminated).sort((a, b) => new Date(b.termination?.terminatedAt || b.updatedAt || 0) - new Date(a.termination?.terminatedAt || a.updatedAt || 0));
  res.json({
    currentUser: publicUser(req.user),
    users: sortedUsers.map(publicUser),
    archivedUsers: archivedUsers.map(publicUser),
    ranks: req.db.settings.ranks,
    roles,
    departmentPositions,
    settings: req.db.settings,
    customPages: req.db.settings.customPages || [],
    departments: req.db.settings.departments.map((department) => publicDepartment(department, req.db, req.user)),
    notes: [...req.db.notes].sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt)),
    dutyHistory: (req.db.dutyHistory || []).map((entry) => ({
      ...entry,
      user: publicUser(req.db.users.find((user) => user.id === entry.userId))
    })),
    logs: (req.db.logs || []).slice(0, 1000),
    disciplinary: req.db.disciplinary || [],
    duty: req.db.duty.map((entry) => ({
      ...entry,
      user: publicUser(req.db.users.find((user) => user.id === entry.userId))
    }))
  });
});

function requireItLead(req, res, next) {
  if (req.user?.role !== "IT-Leitung") {
    return res.status(403).json({ error: "Nur die IT-Leitung darf Mitgliederfluktation bearbeiten." });
  }
  next();
}

app.patch("/api/settings/fluctuation/:id", requireAuth, requireItLead, (req, res) => {
  const rows = req.db.settings.fluctuation || [];
  const row = rows.find((item) => item.id === req.params.id);
  if (!row) return res.status(404).json({ error: "Fluktuationseintrag nicht gefunden." });
  const before = { ...row };
  const type = String(req.body.type || row.type || "").trim();
  if (!["Eingestellt", "Kündigung", "KÃ¼ndigung"].includes(type)) {
    return res.status(400).json({ error: "Ungültiger Typ." });
  }
  const createdAt = req.body.createdAt ? new Date(req.body.createdAt) : new Date(row.createdAt || nowIso());
  if (Number.isNaN(createdAt.getTime())) return res.status(400).json({ error: "Ungültiges Datum." });

  row.name = String(req.body.name || row.name || "").trim();
  row.dn = String(req.body.dn ?? row.dn ?? "").trim();
  row.rank = Number.isInteger(Number(req.body.rank)) ? Number(req.body.rank) : row.rank;
  row.actorName = String(req.body.actorName ?? row.actorName ?? "").trim();
  row.type = type === "KÃ¼ndigung" ? "Kündigung" : type;
  row.reason = String(req.body.reason ?? row.reason ?? "").trim();
  row.createdAt = createdAt.toISOString();
  if (!row.name) return res.status(400).json({ error: "Name ist erforderlich." });

  logAction(req.db, req.user, "Fluktuationseintrag bearbeitet", row.name, { before, after: { ...row } });
  writeDb(req.db);
  res.json({ fluctuation: req.db.settings.fluctuation });
});

app.delete("/api/settings/fluctuation/:id", requireAuth, requireItLead, (req, res) => {
  const rows = req.db.settings.fluctuation || [];
  const index = rows.findIndex((item) => item.id === req.params.id);
  if (index === -1) return res.status(404).json({ error: "Fluktuationseintrag nicht gefunden." });
  const [removed] = rows.splice(index, 1);
  logAction(req.db, req.user, "Fluktuationseintrag gelöscht", removed.name || removed.id, { removed });
  writeDb(req.db);
  res.json({ fluctuation: req.db.settings.fluctuation });
});

app.post("/api/users", requireAuth, requireRole("Direktion"), (req, res) => {
  const normalized = normalizeUserInput(req.body);
  if (normalized.error) return res.status(400).json({ error: normalized.error });
  const roleCheck = protectItRoleChange(req.user, "User", normalized.value.role);
  if (roleCheck.error) return res.status(403).json({ error: roleCheck.error });
  normalized.value.role = roleCheck.role;

  const dnConflict = resolveDnConflict(req.db, "", normalized.value.dn, Boolean(req.body.overwriteDn));
  if (dnConflict?.error) return res.status(400).json({ error: dnConflict.error });

  const createdAt = nowIso();
  const user = {
    id: makeId("user"),
    ...normalized.value,
    trainings: { ...Object.fromEntries(trainingNames.map((training) => [training, false])), ...normalized.value.trainings },
    lastPromotionAt: todayIso(),
    passwordHash: hashPassword(DEFAULT_PASSWORD),
    avatarUrl: "",
    locked: false,
    accountStatus: "Aktiv",
    terminated: false,
    trainingMeta: {},
    createdAt,
    updatedAt: createdAt
  };
  updateTrainingMeta(user, {}, user.trainings, req.user);
  syncDirektionMembership(req.db, user, { roleAssigned: user.role === "Direktion" });

  req.db.users.push(user);
  logFluctuation(req.db, user, "Eingestellt", req.user);
  logAction(req.db, req.user, "Mitglied eingestellt", `${user.firstName} ${user.lastName}`.trim(), { after: publicUser(user) });
  writeDb(req.db);
  res.status(201).json({ user: publicUser(user), defaultPassword: DEFAULT_PASSWORD });
});

app.patch("/api/users/:id", requireAuth, requireRole("Direktion"), (req, res) => {
  const user = req.db.users.find((item) => item.id === req.params.id);
  if (!user) return res.status(404).json({ error: "Benutzer nicht gefunden." });

  const normalized = normalizeUserInput(req.body, user);
  if (normalized.error) return res.status(400).json({ error: normalized.error });
  const roleCheck = protectItRoleChange(req.user, user.role, normalized.value.role);
  if (roleCheck.error) return res.status(403).json({ error: roleCheck.error });
  normalized.value.role = roleCheck.role;

  const dnConflict = resolveDnConflict(req.db, user.id, normalized.value.dn, Boolean(req.body.overwriteDn));
  if (dnConflict?.error) return res.status(400).json({ error: dnConflict.error });

  const before = publicUser(user);
  const previousRole = user.role;
  const rankChanged = Number(user.rank) !== Number(normalized.value.rank);
  const beforeTrainings = { ...(user.trainings || {}) };
  Object.assign(user, normalized.value, {
    lastPromotionAt: rankChanged ? todayIso() : user.lastPromotionAt,
    updatedAt: nowIso()
  });
  updateTrainingMeta(user, beforeTrainings, user.trainings, req.user);
  syncDirektionMembership(req.db, user, { roleAssigned: previousRole !== "Direktion" && user.role === "Direktion" });

  const after = publicUser(user);
  logAction(req.db, req.user, "Benutzer bearbeitet", `${user.firstName} ${user.lastName}`.trim(), { before, after, description: userChangeSummary(req.db, before, after) });
  writeDb(req.db);
  res.json({ user: publicUser(user) });
});

app.post("/api/training/est/:id", requireAuth, (req, res) => {
  const user = req.db.users.find((item) => item.id === req.params.id && !item.terminated);
  if (!user) return res.status(404).json({ error: "Benutzer nicht gefunden." });
  const trainingDepartment = getDepartment(req.db, "training-recruitment");
  if (!canManageDepartmentAction(req.user, trainingDepartment, req.db, "departmentLeadership")) {
    return res.status(403).json({ error: "Keine Berechtigung." });
  }
  if (user.trainings?.EST) return res.json({ user: publicUser(user) });
  const before = publicUser(user);
  const beforeTrainings = { ...(user.trainings || {}) };
  user.trainings = { ...Object.fromEntries(trainingNames.map((training) => [training, false])), ...(user.trainings || {}), EST: true };
  updateTrainingMeta(user, beforeTrainings, user.trainings, req.user);
  user.updatedAt = nowIso();
  const after = publicUser(user);
  logAction(req.db, req.user, "Ausbildung EST hinzugefügt", `${user.firstName} ${user.lastName}`.trim(), { before, after, description: "EST nach bestandener Prüfung vergeben" });
  writeDb(req.db);
  res.json({ user: after });
});

app.patch("/api/settings/uprank-rules", requireAuth, requireRole("Direktion"), (req, res) => {
  const before = normalizeUprankRules(req.db.settings.uprankRules);
  const incoming = Array.isArray(req.body.rules) ? req.body.rules : [];
  req.db.settings.uprankRules = normalizeUprankRules(incoming);
  logAction(req.db, req.user, "Uprank Voraussetzungen geändert", "Direktion", { before, after: req.db.settings.uprankRules });
  writeDb(req.db);
  res.json({ rules: req.db.settings.uprankRules });
});

app.post("/api/users/:id/uprank-adjustments", requireAuth, requireRole("Direktion"), (req, res) => {
  const user = req.db.users.find((item) => item.id === req.params.id && !item.terminated);
  if (!user) return res.status(404).json({ error: "Benutzer nicht gefunden." });
  const type = String(req.body.type || "");
  if (!["Verkürzung", "Sonderuprank"].includes(type)) return res.status(400).json({ error: "Ungültige Uprank-Art." });
  const targetRank = Number(req.body.targetRank || user.rank + 1);
  if (!Number.isInteger(targetRank) || targetRank <= Number(user.rank)) return res.status(400).json({ error: "Ungültiger Zielrang." });
  const days = type === "Verkürzung" ? Math.max(1, Number.parseInt(req.body.days, 10) || 0) : 0;
  const reason = String(req.body.reason || "").trim();
  if (!reason) return res.status(400).json({ error: "Bitte einen Grund angeben." });
  const adjustment = {
    id: makeId("uprank_adjustment"),
    userId: user.id,
    name: actorName(user),
    type,
    targetRank,
    days,
    reason,
    actorName: actorName(req.user),
    createdAt: nowIso()
  };
  req.db.settings.uprankAdjustments.unshift(adjustment);
  logDisciplinary(req.db, user, type, reason, req.user);
  logAction(req.db, req.user, `${type} eingetragen`, actorName(user), { after: adjustment, reason });
  writeDb(req.db);
  res.status(201).json({ adjustment });
});

app.post("/api/users/:id/uprank", requireAuth, requireRole("Direktion"), (req, res) => {
  const user = req.db.users.find((item) => item.id === req.params.id && !item.terminated);
  if (!user) return res.status(404).json({ error: "Benutzer nicht gefunden." });
  const targetRank = Number(req.body.targetRank || user.rank + 1);
  if (!Number.isInteger(targetRank) || targetRank <= Number(user.rank)) return res.status(400).json({ error: "Ungültiger Zielrang." });
  if (!req.body.ingameDone || !req.body.discordDone) return res.status(400).json({ error: "Ingame und Discord müssen abgehakt sein." });
  const reason = String(req.body.reason || "").trim();
  const evaluation = evaluateUprank(req.db, user, targetRank);
  const special = Boolean(req.body.special) || evaluation.hasSpecial;
  if (evaluation.rule.specialOnly && !special) return res.status(400).json({ error: "Dieser Rang ist nur per Sonderuprank möglich." });
  if (!special && (!evaluation.regularReady || evaluation.missingDays || evaluation.missingTrainings.length)) {
    return res.status(400).json({ error: "Die Uprank Voraussetzungen sind noch nicht erfüllt." });
  }
  if (special && !reason) return res.status(400).json({ error: "Bitte Sonderuprank begründen." });
  const before = publicUser(user);
  user.rank = targetRank;
  user.lastPromotionAt = todayIso();
  user.updatedAt = nowIso();
  const after = publicUser(user);
  logDisciplinary(req.db, user, "Uprank", reason || `Uprank auf ${rankText(req.db, targetRank)}`, req.user);
  logAction(req.db, req.user, "Uprank durchgeführt", actorName(user), {
    before,
    after,
    reason,
    ingameDone: true,
    discordDone: true,
    description: `Uprank: ${rankText(req.db, before.rank)} -> ${rankText(req.db, after.rank)}; Ingame erledigt; Discord erledigt${reason ? `; Grund: ${reason}` : ""}`
  });
  writeDb(req.db);
  res.json({ user: after });
});

app.patch("/api/users/:id/lock", requireAuth, requireRole("Direktion"), (req, res) => {
  const user = req.db.users.find((item) => item.id === req.params.id);
  if (!user) return res.status(404).json({ error: "Benutzer nicht gefunden." });
  const reason = String(req.body.reason || "").trim();
  if (!reason) return res.status(400).json({ error: "Bitte einen Grund angeben." });
  const before = publicUser(user);
  setAccountStatus(user, Boolean(req.body.locked) ? "Gesperrt" : "Aktiv");
  user.updatedAt = nowIso();
  req.db.sessions = req.db.sessions.filter((session) => session.userId !== user.id);
  logDisciplinary(req.db, user, user.locked ? "Sperre" : "Entsperrt", reason, req.user);
  logAction(req.db, req.user, user.locked ? "Benutzer gesperrt" : "Benutzer entsperrt", `${user.firstName} ${user.lastName}`.trim(), { reason, before, after: publicUser(user) });
  writeDb(req.db);
  res.json({ user: publicUser(user) });
});

app.post("/api/users/:id/suspend", requireAuth, requireRole("Direktion"), (req, res) => {
  const user = req.db.users.find((item) => item.id === req.params.id);
  if (!user) return res.status(404).json({ error: "Benutzer nicht gefunden." });
  const reason = String(req.body.reason || "").trim();
  if (!reason) return res.status(400).json({ error: "Bitte einen Grund angeben." });
  const before = publicUser(user);
  setAccountStatus(user, "Suspendiert");
  user.updatedAt = nowIso();
  req.db.sessions = req.db.sessions.filter((session) => session.userId !== user.id);
  logDisciplinary(req.db, user, "Suspendierung", reason, req.user);
  logAction(req.db, req.user, "Benutzer suspendiert", `${user.firstName} ${user.lastName}`.trim(), { reason, before, after: publicUser(user) });
  writeDb(req.db);
  res.json({ user: publicUser(user) });
});

app.post("/api/users/:id/dismiss", requireAuth, requireRole("Direktion"), (req, res) => {
  if (req.params.id === req.user.id) return res.status(400).json({ error: "Du kannst dich nicht selbst entlassen." });
  const user = req.db.users.find((item) => item.id === req.params.id);
  if (!user) return res.status(404).json({ error: "Benutzer nicht gefunden." });
  const reason = String(req.body.reason || "").trim();
  if (!reason) return res.status(400).json({ error: "Bitte einen Grund angeben." });
  const before = publicUser(user);
  const oldRank = user.rank;
  const oldDn = user.dn;
  const oldTrainings = { ...(user.trainings || {}) };
  setAccountStatus(user, "Entlassen");
  user.terminated = true;
  user.termination = {
    reason,
    oldRank,
    oldDn,
    oldTrainings,
    terminatedAt: nowIso(),
    actorName: actorName(req.user)
  };
  user.updatedAt = nowIso();
  req.db.duty = req.db.duty.filter((entry) => entry.userId !== user.id);
  req.db.sessions = req.db.sessions.filter((session) => session.userId !== user.id);
  req.db.settings.departments.forEach((department) => {
    department.members = department.members.filter((member) => member.userId !== user.id);
  });
  logFluctuation(req.db, user, "Kündigung", req.user);
  req.db.settings.fluctuation[0].reason = reason;
  logDisciplinary(req.db, user, "Entlassen", reason, req.user);
  logAction(req.db, req.user, "Benutzer entlassen", `${user.firstName} ${user.lastName}`.trim(), { reason, before, after: publicUser(user) });
  writeDb(req.db);
  res.json({ user: publicUser(user) });
});

app.post("/api/users/:id/rehire", requireAuth, requireRole("Direktion"), (req, res) => {
  const user = req.db.users.find((item) => item.id === req.params.id);
  if (!user) return res.status(404).json({ error: "Benutzer nicht gefunden." });
  if (!user.terminated) return res.status(400).json({ error: "Account ist nicht archiviert." });
  const dn = String(req.body.dn || user.termination?.oldDn || user.dn || "").trim();
  const dnError = validateDigits(dn, "Dienstnummer");
  if (dnError) return res.status(400).json({ error: dnError });
  const dnConflict = resolveDnConflict(req.db, user.id, dn, Boolean(req.body.overwriteDn));
  if (dnConflict?.error) return res.status(400).json({ error: dnConflict.error });
  const rank = Number(req.body.rank ?? user.termination?.oldRank ?? user.rank);
  if (!Number.isInteger(rank) || rank < 0) return res.status(400).json({ error: "Ungültiger Rang." });
  const firstName = String(req.body.firstName || user.firstName || "").trim();
  const lastName = String(req.body.lastName || user.lastName || "").trim();
  const phone = String(req.body.phone || user.phone || "").trim();
  const joinedAt = String(req.body.joinedAt || todayIso()).slice(0, 10);
  const requestedRole = roles.includes(req.body.role) ? req.body.role : user.role;
  const roleCheck = protectItRoleChange(req.user, user.role, requestedRole);
  if (roleCheck.error) return res.status(403).json({ error: roleCheck.error });
  const role = roleCheck.role;
  const requestedBaseRole = String(req.body.baseRole || "").trim();
  const baseRole = roles.includes(requestedBaseRole) && !["IT", "IT-Leitung"].includes(requestedBaseRole)
    ? requestedBaseRole
    : user.baseRole || (["IT", "IT-Leitung"].includes(role) ? "Direktion" : role);
  if (!firstName || !lastName || !phone) return res.status(400).json({ error: "Name, Nachname und Telefonnummer sind Pflichtfelder." });
  const before = publicUser(user);
  user.terminated = false;
  setAccountStatus(user, "Aktiv");
  user.firstName = firstName;
  user.lastName = lastName;
  user.phone = phone;
  user.role = role;
  user.baseRole = baseRole;
  user.teamler = Boolean(req.body.teamler);
  user.dn = dn;
  user.rank = rank;
  user.joinedAt = joinedAt;
  const beforeTrainings = { ...(user.trainings || {}) };
  user.trainings = { ...Object.fromEntries(trainingNames.map((training) => [training, false])), ...(req.body.trainings || user.termination?.oldTrainings || user.trainings || {}) };
  updateTrainingMeta(user, beforeTrainings, user.trainings, req.user);
  user.rehiredAt = nowIso();
  user.updatedAt = nowIso();
  syncDirektionMembership(req.db, user, { roleAssigned: role === "Direktion" });
  logFluctuation(req.db, user, "Eingestellt", req.user);
  req.db.settings.fluctuation[0].reason = String(req.body.reason || "Wiedereinstellung").trim() || "Wiedereinstellung";
  logDisciplinary(req.db, user, "Wiedereinstellung", req.db.settings.fluctuation[0].reason, req.user);
  logAction(req.db, req.user, "Benutzer wiedereingestellt", `${user.firstName} ${user.lastName}`.trim(), { before, after: publicUser(user) });
  writeDb(req.db);
  res.json({ user: publicUser(user) });
});

app.post("/api/users/:id/file", requireAuth, requireRole("Direktion"), (req, res) => {
  const user = req.db.users.find((item) => item.id === req.params.id);
  if (!user) return res.status(404).json({ error: "Benutzer nicht gefunden." });
  const type = String(req.body.type || "").trim();
  const reason = String(req.body.reason || "").trim();
  if (!["Aktennotiz", "Sanktion", "Strike"].includes(type)) return res.status(400).json({ error: "Ungueltiger Akteneintrag." });
  if (!reason) return res.status(400).json({ error: "Bitte einen Text oder Grund angeben." });
  const sanctionType = type === "Strike" ? "Strike" : String(req.body.sanctionType || "").trim();
  if (type === "Sanktion" && !["Strike", "Geldstrafe", "Custom"].includes(sanctionType)) return res.status(400).json({ error: "Bitte eine Sanktionsart auswaehlen." });
  const strikeCount = sanctionType === "Strike" || type === "Strike" ? Math.max(1, Math.min(3, Number(req.body.strikeCount || 1))) : 0;
  if ((type === "Strike" || sanctionType === "Strike") && activeStrikeCount(req.db.disciplinary, user.id) + strikeCount > 3) {
    return res.status(400).json({ error: "Dieses Mitglied hat bereits 3/3 aktive Strikes." });
  }
  const entry = {
    id: makeId("disciplinary"),
    type: type === "Aktennotiz" ? "Aktennotiz" : "Sanktion",
    sanctionType: type === "Aktennotiz" ? "" : sanctionType,
    userId: user.id,
    name: `${user.firstName} ${user.lastName}`.trim(),
    dn: user.dn,
    rank: user.rank,
    actorName: actorName(req.user),
    reason,
    title: String(req.body.title || sanctionType || type).trim(),
    amount: Number(req.body.amount || 0),
    strikeCount,
    expiresAt: String(req.body.expiresAt || "").trim(),
    createdAt: nowIso()
  };
  req.db.disciplinary.unshift(entry);
  logAction(req.db, req.user, `${entry.type} eingetragen`, `${user.firstName} ${user.lastName}`.trim(), { reason, sanctionType: entry.sanctionType, amount: entry.amount || "" });
  writeDb(req.db);
  res.status(201).json({ disciplinary: req.db.disciplinary });
});

app.delete("/api/users/:id/file/:entryId", requireAuth, requireRole("Direktion"), (req, res) => {
  const user = req.db.users.find((item) => item.id === req.params.id);
  if (!user) return res.status(404).json({ error: "Benutzer nicht gefunden." });
  const entry = (req.db.disciplinary || []).find((item) => item.id === req.params.entryId && item.userId === user.id);
  if (!entry) return res.status(404).json({ error: "Akteneintrag nicht gefunden." });
  if (entry.type === "Sanktion" || entry.type === "Strike") {
    entry.archivedAt = nowIso();
    entry.archivedBy = actorName(req.user);
    logAction(req.db, req.user, "Sanktion archiviert", `${user.firstName} ${user.lastName}`.trim(), { before: entry });
  } else if (entry.type === "Aktennotiz") {
    req.db.disciplinary = req.db.disciplinary.filter((item) => item.id !== entry.id);
    logAction(req.db, req.user, "Aktennotiz entfernt", `${user.firstName} ${user.lastName}`.trim(), { before: entry });
  } else {
    return res.status(400).json({ error: "Dieser Eintrag kann nicht entfernt werden." });
  }
  writeDb(req.db);
  res.json({ ok: true });
});

app.patch("/api/users/:id/file/:entryId", requireAuth, requireRole("Direktion"), (req, res) => {
  const user = req.db.users.find((item) => item.id === req.params.id);
  if (!user) return res.status(404).json({ error: "Benutzer nicht gefunden." });
  const entry = (req.db.disciplinary || []).find((item) => item.id === req.params.entryId && item.userId === user.id);
  if (!entry) return res.status(404).json({ error: "Akteneintrag nicht gefunden." });
  const before = { ...entry };
  if (entry.type === "Aktennotiz") {
    const reason = String(req.body.reason || "").trim();
    if (!reason) return res.status(400).json({ error: "Bitte eine Notiz angeben." });
    entry.reason = reason;
    entry.updatedAt = nowIso();
    entry.updatedBy = actorName(req.user);
    logAction(req.db, req.user, "Aktennotiz bearbeitet", `${user.firstName} ${user.lastName}`.trim(), { before, after: entry });
  } else if (entry.type === "Sanktion" && entry.sanctionType === "Geldstrafe") {
    entry.paidAt = nowIso();
    entry.paidBy = actorName(req.user);
    logAction(req.db, req.user, "Geldstrafe bezahlt", `${user.firstName} ${user.lastName}`.trim(), { before, after: entry });
  } else {
    return res.status(400).json({ error: "Dieser Eintrag kann nicht bearbeitet werden." });
  }
  writeDb(req.db);
  res.json({ entry });
});

app.delete("/api/users/:id", requireAuth, requireRole("Direktion"), (req, res) => {
  if (req.params.id === req.user.id) return res.status(400).json({ error: "Du kannst dich nicht selbst löschen." });
  req.db.users = req.db.users.filter((user) => user.id !== req.params.id);
  req.db.duty = req.db.duty.filter((entry) => entry.userId !== req.params.id);
  req.db.sessions = req.db.sessions.filter((session) => session.userId !== req.params.id);
  logAction(req.db, req.user, "Benutzer gelöscht", req.params.id);
  writeDb(req.db);
  res.json({ ok: true });
});

app.patch("/api/profile/password", requireAuth, (req, res) => {
  const oldPassword = String(req.body.oldPassword || "");
  const newPassword = String(req.body.newPassword || "");
  if (req.user.passwordHash !== hashPassword(oldPassword)) return res.status(400).json({ error: "Altes Passwort stimmt nicht." });
  if (newPassword.length < 6) return res.status(400).json({ error: "Neues Passwort muss mindestens 6 Zeichen haben." });
  req.user.passwordHash = hashPassword(newPassword);
  req.user.updatedAt = nowIso();
  logAction(req.db, req.user, "Passwort geändert", `${req.user.firstName} ${req.user.lastName}`.trim());
  writeDb(req.db);
  res.json({ ok: true });
});

app.patch("/api/profile/avatar", requireAuth, (req, res) => {
  const avatarUrl = String(req.body.avatarUrl || "").trim();
  if (avatarUrl && !/^https?:\/\//i.test(avatarUrl) && !avatarUrl.startsWith("/") && !avatarUrl.startsWith("data:image/")) {
    return res.status(400).json({ error: "Avatar muss eine http(s)- oder lokale URL sein." });
  }
  const before = req.user.avatarUrl || "";
  req.user.avatarUrl = avatarUrl;
  req.user.updatedAt = nowIso();
  logAction(req.db, req.user, "Avatar aktualisiert", `${req.user.firstName} ${req.user.lastName}`.trim(), { before: before ? "Avatar vorhanden" : "Kein Avatar", after: avatarUrl ? "Avatar hochgeladen" : "Avatar entfernt" });
  writeDb(req.db);
  res.json({ user: publicUser(req.user) });
});

app.patch("/api/settings/defcon", requireAuth, requirePermission("actions", "editDefcon", "Supervisor"), (req, res) => {
  const defcon = String(req.body.defcon || "");
  if (!/^DEFCON [1-5]$/.test(defcon)) return res.status(400).json({ error: "Ungueltige DEFCON-Stufe." });
  const before = { defcon: req.db.settings.defcon, defconText: req.db.settings.defconText };
  req.db.settings.defcon = defcon;
  req.db.settings.defconText = typeof req.body.defconText === "string" ? req.body.defconText.trim() : String(req.db.settings.defconText || "");
  req.db.settings.defconUpdatedBy = `${req.user.firstName} ${req.user.lastName}`.trim();
  req.db.settings.defconUpdatedAt = nowIso();
  logAction(req.db, req.user, "DEFCON geändert", "DEFCON", { before, after: { defcon, defconText: req.db.settings.defconText } });
  writeDb(req.db);
  res.json({ settings: req.db.settings });
});

app.patch("/api/information", requireAuth, requirePermission("actions", "manageInformation", "Direktion"), (req, res) => {
  const before = {
    informationText: req.db.settings.informationText,
    applicationStatus: req.db.settings.applicationStatus,
    informationRightsText: req.db.settings.informationRightsText,
    informationLinks: req.db.settings.informationLinks,
    informationDocs: req.db.settings.informationDocs,
    informationDocChanges: req.db.settings.informationDocChanges,
    informationPermits: req.db.settings.informationPermits,
    informationFactions: req.db.settings.informationFactions
  };
  req.db.settings.informationText = String(req.body.informationText || "").trim();
  req.db.settings.applicationStatus = ["Offen", "Geschlossen"].includes(req.body.applicationStatus) ? req.body.applicationStatus : "Offen";
  req.db.settings.informationRightsText = String(req.body.informationRightsText || "").trim();
  req.db.settings.informationLinks = Array.isArray(req.body.informationLinks) ? req.body.informationLinks.map((item) => ({ id: String(item.id || makeId("link")), title: String(item.title || "").trim(), url: String(item.url || "").trim() })).filter((item) => item.title && item.url) : [];
  req.db.settings.informationDocs = Array.isArray(req.body.informationDocs) ? req.body.informationDocs.map((item) => ({ id: String(item.id || makeId("doc")), title: String(item.title || "").trim(), body: String(item.body || "").trim(), updatedAt: String(item.updatedAt || new Date().toISOString()), updatedBy: String(item.updatedBy || "") })).filter((item) => item.title) : [];
  req.db.settings.informationDocChanges = Array.isArray(req.body.informationDocChanges) ? req.body.informationDocChanges.map((item) => ({ id: String(item.id || makeId("docchange")), docId: String(item.docId || ""), title: String(item.title || "").trim(), before: String(item.before || ""), after: String(item.after || ""), action: String(item.action || "geändert"), createdAt: String(item.createdAt || new Date().toISOString()), author: String(item.author || ""), acknowledgedBy: Array.isArray(item.acknowledgedBy) ? item.acknowledgedBy.map(String) : [] })) : [];
  req.db.settings.informationPermits = Array.isArray(req.body.informationPermits) ? req.body.informationPermits.map((item) => ({ id: String(item.id || makeId("permit")), name: String(item.name || "").trim(), description: String(item.description || "").trim(), validUntil: String(item.validUntil || "").trim() })).filter((item) => item.name && item.description && item.validUntil) : [];
  req.db.settings.informationFactions = Array.isArray(req.body.informationFactions) ? req.body.informationFactions.map((item) => ({ id: String(item.id || makeId("faction")), organization: String(item.organization || "").trim(), status: ["Normal", "Mittel", "Hoch"].includes(item.status) ? item.status : "Normal" })).filter((item) => item.organization) : [];
  logAction(req.db, req.user, "Informationen geändert", "Informationen", { before, after: {
    informationText: req.db.settings.informationText,
    applicationStatus: req.db.settings.applicationStatus,
    informationRightsText: req.db.settings.informationRightsText,
    informationLinks: req.db.settings.informationLinks,
    informationDocs: req.db.settings.informationDocs,
    informationDocChanges: req.db.settings.informationDocChanges,
    informationPermits: req.db.settings.informationPermits,
    informationFactions: req.db.settings.informationFactions
  } });
  writeDb(req.db);
  res.json({ settings: req.db.settings });
});

app.patch("/api/it/ranks", requireAuth, requireRole("IT"), (req, res) => {
  const nextRanks = Array.isArray(req.body.ranks) ? req.body.ranks : [];
  if (!nextRanks.length) return res.status(400).json({ error: "Es muss mindestens ein Rang vorhanden sein." });

  const before = req.db.settings.ranks;
  req.db.settings.ranks = nextRanks
    .map((rank) => ({
      value: Number(rank.value),
      label: String(rank.label || `Template ${rank.value} - Rang ${rank.value}`).trim()
    }))
    .filter((rank) => Number.isInteger(rank.value) && rank.value >= 0)
    .sort((a, b) => a.value - b.value);
  logAction(req.db, req.user, "Ränge geändert", "IT", { before, after: req.db.settings.ranks });
  writeDb(req.db);
  res.json({ ranks: req.db.settings.ranks });
});

app.patch("/api/it/nav-labels", requireAuth, requireRole("IT"), (req, res) => {
  const navLabels = req.body.navLabels && typeof req.body.navLabels === "object" ? req.body.navLabels : {};
  const before = {
    navLabels: req.db.settings.navLabels,
    departments: req.db.settings.departments.map((department) => ({ id: department.id, name: department.name }))
  };
  const nextNavLabels = {};
  Object.entries(navLabels).forEach(([key, value]) => {
    const label = String(value || key).trim();
    if (key.startsWith("dept:")) {
      const departmentId = key.slice(5);
      const department = req.db.settings.departments.find((item) => item.id === departmentId);
      if (department && label) department.name = label;
      return;
    }
    nextNavLabels[key] = label;
  });
  req.db.settings.navLabels = nextNavLabels;
  logAction(req.db, req.user, "Reiter geändert", "IT", {
    before,
    after: {
      navLabels: req.db.settings.navLabels,
      departments: req.db.settings.departments.map((department) => ({ id: department.id, name: department.name }))
    }
  });
  writeDb(req.db);
  res.json({
    navLabels: req.db.settings.navLabels,
    departments: req.db.settings.departments.map((department) => publicDepartment(department, req.db, req.user))
  });
});

app.patch("/api/it/page-order", requireAuth, requireRole("IT"), (req, res) => {
  const pageOrder = Array.isArray(req.body.pageOrder) ? req.body.pageOrder.map(String).filter(Boolean) : [];
  const before = req.db.settings.pageOrder || [];
  req.db.settings.pageOrder = [...new Set(pageOrder)];
  logAction(req.db, req.user, "Reiter sortiert", "IT", { before, after: req.db.settings.pageOrder });
  writeDb(req.db);
  res.json({ settings: req.db.settings });
});

app.post("/api/it/custom-pages", requireAuth, requireRole("IT"), (req, res) => {
  const name = String(req.body.name || "").trim();
  if (!name) return res.status(400).json({ error: "Name ist erforderlich." });
  req.db.settings.customPages = Array.isArray(req.db.settings.customPages) ? req.db.settings.customPages : [];
  const existingKeys = new Set([
    ...req.db.settings.customPages.map((page) => page.key),
    ...req.db.settings.departments.map((department) => `dept:${department.id}`)
  ]);
  let base = `custom:${slugify(name)}`;
  let key = base;
  let index = 2;
  while (existingKeys.has(key)) key = `${base}-${index++}`;
  const page = { key, name, createdAt: nowIso() };
  req.db.settings.customPages.push(page);
  req.db.settings.navLabels = { ...(req.db.settings.navLabels || {}), [key]: name };
  req.db.settings.permissions = normalizePermissions(req.db.settings.permissions || {});
  req.db.settings.permissions.pages[key] = { all: false, roles: ["IT", "IT-Leitung"], ranks: [], users: [], departments: [], positions: [] };
  req.db.settings.pageOrder = [...new Set([...(req.db.settings.pageOrder || []), key])];
  logAction(req.db, req.user, "Reiter erstellt", name, { after: page });
  writeDb(req.db);
  res.status(201).json({ page, settings: req.db.settings });
});

app.post("/api/it/departments", requireAuth, requireRole("IT"), (req, res) => {
  const name = String(req.body.name || "").trim();
  if (!name) return res.status(400).json({ error: "Name ist erforderlich." });
  const existingIds = new Set(req.db.settings.departments.map((department) => department.id));
  let base = slugify(name, "abteilung");
  let id = base;
  let index = 2;
  while (existingIds.has(id)) id = `${base}-${index++}`;
  const department = makeDepartment(id, name, "Leeres Abteilungsblatt", "Offen");
  req.db.settings.departments.push(department);
  req.db.settings.permissions = normalizePermissions(req.db.settings.permissions || {});
  req.db.settings.permissions.pages[`dept:${id}`] = { all: false, roles: ["IT", "IT-Leitung"], ranks: [], users: [], departments: [], positions: [] };
  req.db.settings.pageOrder = [...new Set([...(req.db.settings.pageOrder || []), `dept:${id}`])];
  logAction(req.db, req.user, "Abteilung erstellt", name, { after: department });
  writeDb(req.db);
  res.status(201).json({
    department: publicDepartment(department, req.db, req.user),
    settings: req.db.settings,
    departments: req.db.settings.departments.map((item) => publicDepartment(item, req.db, req.user))
  });
});

app.patch("/api/it/permissions", requireAuth, requireRole("IT"), (req, res) => {
  const before = req.db.settings.permissions || defaultPermissions();
  req.db.settings.permissions = normalizePermissions(req.body.permissions || {});
  logAction(req.db, req.user, "Berechtigungen geändert", "IT", { before, after: req.db.settings.permissions });
  writeDb(req.db);
  res.json({ permissions: req.db.settings.permissions });
});

app.patch("/api/it/devmode", requireAuth, requireRole("IT"), (req, res) => {
  const before = Boolean(req.db.settings.devMode);
  req.db.settings.devMode = Boolean(req.body.devMode);
  logAction(req.db, req.user, req.db.settings.devMode ? "Devmode aktiviert" : "Devmode deaktiviert", "IT", { before, after: req.db.settings.devMode });
  writeDb(req.db);
  res.json({ settings: req.db.settings });
});

app.patch("/api/it/restarts", requireAuth, requireRole("IT"), (req, res) => {
  const before = req.db.settings.restartTimes || [];
  const restartTimes = Array.isArray(req.body.restartTimes) ? req.body.restartTimes : [];
  req.db.settings.restartTimes = [...new Set(restartTimes
    .map((time) => String(time || "").trim())
    .filter((time) => /^\d{2}:\d{2}$/.test(time)))]
    .sort();
  logAction(req.db, req.user, "Restartzeiten geändert", "IT", { before, after: req.db.settings.restartTimes });
  writeDb(req.db);
  res.json({ settings: req.db.settings });
});

app.get("/api/it/export", requireAuth, requireRole("IT"), (req, res) => {
  const exportDb = { ...req.db, sessions: [] };
  res.setHeader("Content-Type", "application/json");
  res.setHeader("Content-Disposition", "attachment; filename=\"lspd-dienstblatt-export.json\"");
  res.send(JSON.stringify(exportDb, null, 2));
});

app.post("/api/it/import", requireAuth, requireRole("IT"), (req, res) => {
  const imported = req.body?.db || req.body;
  if (!imported || typeof imported !== "object") return res.status(400).json({ error: "Keine gültige JSON-Datei empfangen." });
  if (!Array.isArray(imported.users) || !imported.settings || typeof imported.settings !== "object") {
    return res.status(400).json({ error: "Die Datei ist keine gültige Dienstblatt-Datensicherung." });
  }
  if (!imported.users.length) return res.status(400).json({ error: "Die Datensicherung enthält keine Benutzer." });

  const nextDb = {
    ...imported,
    users: imported.users,
    sessions: [],
    settings: imported.settings,
    notes: Array.isArray(imported.notes) ? imported.notes : [],
    duty: Array.isArray(imported.duty) ? imported.duty : [],
    dutyHistory: Array.isArray(imported.dutyHistory) ? imported.dutyHistory : [],
    logs: Array.isArray(imported.logs) ? imported.logs : [],
    disciplinary: Array.isArray(imported.disciplinary) ? imported.disciplinary : []
  };
  writeDb(nextDb);
  res.json({ ok: true, users: nextDb.users.length });
});

app.post("/api/it/clear-sessions", requireAuth, requireRole("IT"), (req, res) => {
  req.db.sessions = req.db.sessions.filter((session) => session.userId === req.user.id);
  writeDb(req.db);
  res.json({ ok: true });
});

app.patch("/api/departments/:departmentId/positions", requireAuth, requireRole("IT"), (req, res) => {
  const department = getDepartment(req.db, req.params.departmentId);
  if (!department) return res.status(404).json({ error: "Abteilung nicht gefunden." });
  const incoming = Array.isArray(req.body.positions) ? req.body.positions : [];
  const normalized = incoming
    .map((item) => ({
      old: String(item.old || "").trim(),
      label: String(item.label || "").trim()
    }))
    .filter((item) => item.label);
  const nextPositions = [...new Set(normalized.map((item) => item.label))];
  if (!nextPositions.length) return res.status(400).json({ error: "Mindestens eine Position ist erforderlich." });
  if (!nextPositions.includes("Direktion")) return res.status(400).json({ error: "Die Position Direktion muss erhalten bleiben." });
  const removedPositions = departmentPositionsFor(department).filter((position) => !normalized.some((item) => item.old === position || item.label === position));
  const positionInUse = removedPositions.find((position) => department.members.some((member) => member.position === position));
  if (positionInUse) return res.status(400).json({ error: `Die Position ${positionInUse} ist noch vergeben und kann nicht entfernt werden.` });

  const before = [...departmentPositionsFor(department)];
  normalized.forEach((item) => {
    if (item.old && item.old !== item.label) {
      department.members.forEach((member) => {
        if (member.position === item.old) member.position = item.label;
      });
    }
  });
  department.positions = nextPositions;
  logAction(req.db, req.user, "Abteilungsränge geändert", department.name, { before, after: department.positions });
  writeDb(req.db);
  res.json({ department: publicDepartment(department, req.db, req.user) });
});

app.patch("/api/departments/:departmentId/info", requireAuth, (req, res) => {
  const department = getDepartment(req.db, req.params.departmentId);
  if (!canManageDepartmentAction(req.user, department, req.db, "departmentInfo")) return res.status(403).json({ error: "Keine Berechtigung." });
  const before = { description: department.description, applicationStatus: department.applicationStatus, requirements: department.requirements, rightsText: department.rightsText, links: department.links, permits: department.permits, factions: department.factions };
  department.description = String(req.body.description || "").trim();
  department.applicationStatus = ["Offen", "Geschlossen"].includes(req.body.applicationStatus) ? req.body.applicationStatus : "Offen";
  department.requirements = String(req.body.requirements || "").trim();
  department.rightsText = String(req.body.rightsText || "").trim();
  department.links = Array.isArray(req.body.links) ? req.body.links.map((item) => ({ id: String(item.id || makeId("link")), title: String(item.title || "").trim(), url: String(item.url || "").trim() })).filter((item) => item.title && item.url) : [];
  department.permits = Array.isArray(req.body.permits) ? req.body.permits.map((item) => ({ id: String(item.id || makeId("permit")), name: String(item.name || "").trim(), description: String(item.description || "").trim(), validUntil: String(item.validUntil || "").trim() })).filter((item) => item.name && item.description && item.validUntil) : [];
  department.factions = Array.isArray(req.body.factions) ? req.body.factions.map((item) => ({ id: String(item.id || makeId("faction")), organization: String(item.organization || "").trim(), status: ["Normal", "Mittel", "Hoch"].includes(item.status) ? item.status : "Normal" })).filter((item) => item.organization) : [];
  logAction(req.db, req.user, "Abteilungsinfos geändert", department.name, { before, after: { description: department.description, applicationStatus: department.applicationStatus, requirements: department.requirements, rightsText: department.rightsText, links: department.links, permits: department.permits, factions: department.factions } });
  writeDb(req.db);
  res.json({ department: publicDepartment(department, req.db, req.user) });
});

app.post("/api/departments/:departmentId/members", requireAuth, (req, res) => {
  const department = getDepartment(req.db, req.params.departmentId);
  if (!canManageDepartmentAction(req.user, department, req.db, "departmentMembers")) return res.status(403).json({ error: "Keine Berechtigung." });
  const userId = String(req.body.userId || "");
  const position = departmentPositionsFor(department).includes(req.body.position) ? req.body.position : "Mitglied";
  if (!canAssignDepartmentPosition(req.user, department, position, req.db)) return res.status(403).json({ error: "Diese Position darfst du nicht vergeben." });
  const addedUser = req.db.users.find((user) => user.id === userId);
  if (!addedUser) return res.status(404).json({ error: "Benutzer nicht gefunden." });
  if (department.members.some((member) => member.userId === userId)) return res.status(400).json({ error: "Person ist bereits in der Abteilung." });
  if (department.id === "direktion" && addedUser.role === "Direktion") addedUser.direktionManualRemoved = false;
  department.members.push({ userId, position, joinedAt: todayIso(), positionSince: todayIso() });
  logAction(req.db, req.user, "Abteilungsmitglied hinzugefügt", department.name, { userId, position });
  writeDb(req.db);
  res.status(201).json({ department: publicDepartment(department, req.db, req.user) });
});

app.patch("/api/departments/:departmentId/members/:userId", requireAuth, (req, res) => {
  const department = getDepartment(req.db, req.params.departmentId);
  if (!canManageDepartmentAction(req.user, department, req.db, "departmentMembers")) return res.status(403).json({ error: "Keine Berechtigung." });
  const member = department.members.find((item) => item.userId === req.params.userId);
  if (!member) return res.status(404).json({ error: "Mitglied nicht gefunden." });
  const before = { ...member };
  const nextPosition = departmentPositionsFor(department).includes(req.body.position) ? req.body.position : member.position;
  if (!canAssignDepartmentPosition(req.user, department, nextPosition, req.db)) return res.status(403).json({ error: "Diese Position darfst du nicht vergeben." });
  if (member.position !== nextPosition) member.positionSince = todayIso();
  member.position = nextPosition;
  logAction(req.db, req.user, "Abteilungsmitglied bearbeitet", department.name, { before, after: member });
  writeDb(req.db);
  res.json({ department: publicDepartment(department, req.db, req.user) });
});

app.delete("/api/departments/:departmentId/members/:userId", requireAuth, (req, res) => {
  const department = getDepartment(req.db, req.params.departmentId);
  if (!canManageDepartmentAction(req.user, department, req.db, "departmentMembers")) return res.status(403).json({ error: "Keine Berechtigung." });
  const removedUser = req.db.users.find((user) => user.id === req.params.userId);
  if (department?.id === "direktion" && removedUser?.role === "Direktion") removedUser.direktionManualRemoved = true;
  department.members = department.members.filter((member) => member.userId !== req.params.userId);
  logAction(req.db, req.user, "Abteilungsmitglied entfernt", department.name, { userId: req.params.userId });
  writeDb(req.db);
  res.json({ department: publicDepartment(department, req.db, req.user) });
});

app.post("/api/departments/:departmentId/notes", requireAuth, (req, res) => {
  const department = getDepartment(req.db, req.params.departmentId);
  if (!canManageDepartmentAction(req.user, department, req.db, "departmentNotes")) return res.status(403).json({ error: "Keine Berechtigung." });
  const title = String(req.body.title || "").trim();
  const priority = String(req.body.priority || "").trim();
  const text = String(req.body.text || "").trim();
  if (!title || !text || !["Leitung", "Info", "Mitglied"].includes(priority)) {
    return res.status(400).json({ error: "Titel, Priorität und Text sind erforderlich." });
  }
  const note = {
    id: makeId("dept_note"),
    title,
    priority,
    text,
    authorId: req.user.id,
    authorName: `${req.user.firstName} ${req.user.lastName}`.trim(),
    createdAt: nowIso()
  };
  department.notes.push(note);
  logAction(req.db, req.user, "Abteilungsnotiz erstellt", department.name, { after: note });
  writeDb(req.db);
  res.status(201).json({ department: publicDepartment(department, req.db, req.user) });
});

app.patch("/api/departments/:departmentId/notes/:noteId", requireAuth, (req, res) => {
  const department = getDepartment(req.db, req.params.departmentId);
  if (!canManageDepartmentAction(req.user, department, req.db, "departmentNotes")) return res.status(403).json({ error: "Keine Berechtigung." });
  const note = department.notes.find((item) => item.id === req.params.noteId);
  if (!note) return res.status(404).json({ error: "Notiz nicht gefunden." });
  const before = { ...note };
  note.title = String(req.body.title || "").trim();
  note.priority = ["Leitung", "Info", "Mitglied"].includes(req.body.priority) ? req.body.priority : note.priority;
  note.text = String(req.body.text || "").trim();
  note.updatedAt = nowIso();
  logAction(req.db, req.user, "Abteilungsnotiz geändert", department.name, { before, after: note });
  writeDb(req.db);
  res.json({ department: publicDepartment(department, req.db, req.user) });
});

app.delete("/api/departments/:departmentId/notes/:noteId", requireAuth, (req, res) => {
  const department = getDepartment(req.db, req.params.departmentId);
  if (!canManageDepartmentAction(req.user, department, req.db, "departmentNotes")) return res.status(403).json({ error: "Keine Berechtigung." });
  department.notes = department.notes.filter((note) => note.id !== req.params.noteId);
  logAction(req.db, req.user, "Abteilungsnotiz gelöscht", department.name, { noteId: req.params.noteId });
  writeDb(req.db);
  res.json({ department: publicDepartment(department, req.db, req.user) });
});

app.post("/api/departments/:departmentId/member-notes", requireAuth, (req, res) => {
  const department = getDepartment(req.db, req.params.departmentId);
  if (!canManageDepartmentAction(req.user, department, req.db, "departmentLeadership")) return res.status(403).json({ error: "Keine Berechtigung." });
  const userId = String(req.body.userId || "");
  const text = String(req.body.text || "").trim();
  if (!department.members.some((member) => member.userId === userId)) return res.status(404).json({ error: "Mitglied nicht gefunden." });
  if (!text) return res.status(400).json({ error: "Notiz ist erforderlich." });
  const note = {
    id: makeId("dept-member-note"),
    userId,
    text,
    authorId: req.user.id,
    authorName: `${req.user.firstName} ${req.user.lastName}`.trim(),
    createdAt: nowIso()
  };
  department.memberNotes = Array.isArray(department.memberNotes) ? department.memberNotes : [];
  department.memberNotes.push(note);
  logAction(req.db, req.user, "Interne Abteilungsnotiz erstellt", department.name, { after: note });
  writeDb(req.db);
  res.status(201).json({ department: publicDepartment(department, req.db, req.user) });
});

app.post("/api/notes", requireAuth, requirePermission("actions", "manageNotes", "Supervisor"), (req, res) => {
  const title = String(req.body.title || "").trim();
  const priority = String(req.body.priority || "").trim();
  const text = String(req.body.text || "").trim();
  if (!title || !text || !["Info", "Anweisung", "Direktion"].includes(priority)) {
    return res.status(400).json({ error: "Titel, Prioritaet und Text sind erforderlich." });
  }

  const note = {
    id: makeId("note"),
    title,
    priority,
    text,
    authorId: req.user.id,
    authorName: `${req.user.firstName} ${req.user.lastName}`.trim(),
    createdAt: nowIso()
  };
  req.db.notes.push(note);
  logAction(req.db, req.user, "Notiz erstellt", note.title, { after: note });
  writeDb(req.db);
  res.status(201).json({ note });
});

app.patch("/api/notes/:id", requireAuth, requirePermission("actions", "manageNotes", "Supervisor"), (req, res) => {
  const note = req.db.notes.find((item) => item.id === req.params.id);
  if (!note) return res.status(404).json({ error: "Notiz nicht gefunden." });

  const title = String(req.body.title || "").trim();
  const priority = String(req.body.priority || "").trim();
  const text = String(req.body.text || "").trim();
  if (!title || !text || !["Info", "Anweisung", "Direktion"].includes(priority)) {
    return res.status(400).json({ error: "Titel, Prioritaet und Text sind erforderlich." });
  }

  const before = { ...note };
  Object.assign(note, {
    title,
    priority,
    text,
    updatedBy: `${req.user.firstName} ${req.user.lastName}`.trim(),
    updatedAt: nowIso()
  });
  logAction(req.db, req.user, "Notiz geändert", note.title, { before, after: note });
  writeDb(req.db);
  res.json({ note });
});

app.delete("/api/notes/:id", requireAuth, requirePermission("actions", "manageNotes", "Supervisor"), (req, res) => {
  const note = req.db.notes.find((item) => item.id === req.params.id);
  req.db.notes = req.db.notes.filter((note) => note.id !== req.params.id);
  logAction(req.db, req.user, "Notiz gelöscht", note?.title || req.params.id, { before: note || null });
  writeDb(req.db);
  res.json({ ok: true });
});

app.post("/api/duty/start", requireAuth, (req, res) => {
  const status = String(req.body.status || "");
  if (!["Innendienst", "Außendienst", "Undercover Dienst", "Admin Dienst"].includes(status)) {
    return res.status(400).json({ error: "Ungueltiger Dienststatus." });
  }
  if (status === "Admin Dienst" && !req.user.teamler && (rolePower[req.user.role] || 0) < rolePower.IT) {
    return res.status(403).json({ error: "Admin Dienst ist nur für Teamler freigegeben." });
  }
  if (req.db.duty.some((entry) => entry.userId === req.user.id)) {
    return res.status(400).json({ error: "Du bist bereits im Dienst." });
  }
  const entry = {
    id: makeId("duty"),
    userId: req.user.id,
    status,
    startedAt: nowIso()
  };
  req.db.duty.push(entry);
  req.db.dutyHistory.push({ ...entry, endedAt: "", manual: false });
  logAction(req.db, req.user, "Dienst gestartet", status, { after: entry });
  writeDb(req.db);
  res.status(201).json({ entry });
});

app.post("/api/duty/switch", requireAuth, (req, res) => {
  const status = String(req.body.status || "");
  if (!["Innendienst", "Außendienst", "Undercover Dienst", "Admin Dienst"].includes(status)) {
    return res.status(400).json({ error: "Ungueltiger Dienststatus." });
  }
  if (status === "Admin Dienst" && !req.user.teamler && (rolePower[req.user.role] || 0) < rolePower.IT) {
    return res.status(403).json({ error: "Admin Dienst ist nur für Teamler freigegeben." });
  }
  const active = req.db.duty.find((entry) => entry.userId === req.user.id);
  if (!active) return res.status(400).json({ error: "Du bist aktuell nicht im Dienst." });
  const before = { ...active };
  active.status = status;
  active.switchedAt = nowIso();
  const history = req.db.dutyHistory.find((entry) => entry.id === active.id) || req.db.dutyHistory.find((entry) => entry.userId === req.user.id && !entry.endedAt);
  if (history) {
    history.status = status;
    history.switchedAt = active.switchedAt;
  }
  logAction(req.db, req.user, "Dienst umgetragen", status, { before, after: active });
  writeDb(req.db);
  res.json({ entry: active });
});

app.post("/api/duty/stop", requireAuth, (req, res) => {
  const active = req.db.duty.find((entry) => entry.userId === req.user.id);
  if (active) {
    const history = req.db.dutyHistory.find((entry) => entry.id === active.id) || req.db.dutyHistory.find((entry) => entry.userId === req.user.id && !entry.endedAt);
    if (history) history.endedAt = nowIso();
    else req.db.dutyHistory.push({ ...active, endedAt: nowIso(), manual: false });
    logAction(req.db, req.user, "Dienst beendet", active.status, { before: active, endedAt: nowIso() });
  }
  req.db.duty = req.db.duty.filter((entry) => entry.userId !== req.user.id);
  writeDb(req.db);
  res.json({ ok: true });
});

app.post("/api/duty/stop/:userId", requireAuth, requireRole("Supervisor"), (req, res) => {
  const active = req.db.duty.find((entry) => entry.userId === req.params.userId);
  if (active) {
    const history = req.db.dutyHistory.find((entry) => entry.id === active.id) || req.db.dutyHistory.find((entry) => entry.userId === req.params.userId && !entry.endedAt);
    if (history) history.endedAt = nowIso();
    else req.db.dutyHistory.push({ ...active, endedAt: nowIso(), manual: false });
    logAction(req.db, req.user, "Dienst beendet", active.status, { userId: req.params.userId, before: active, endedAt: nowIso() });
  }
  req.db.duty = req.db.duty.filter((entry) => entry.userId !== req.params.userId);
  writeDb(req.db);
  res.json({ ok: true });
});

app.post("/api/duty/stop-all", requireAuth, requirePermission("actions", "stopAllDuty", "Direktion"), (req, res) => {
  const endedAt = nowIso();
  req.db.duty.forEach((active) => {
    const history = req.db.dutyHistory.find((entry) => entry.id === active.id) || req.db.dutyHistory.find((entry) => entry.userId === active.userId && !entry.endedAt);
    if (history) history.endedAt = endedAt;
    else req.db.dutyHistory.push({ ...active, endedAt, manual: false });
  });
  logAction(req.db, req.user, "Alle Dienste beendet", "Dienstblatt", { count: req.db.duty.length });
  req.db.duty = [];
  writeDb(req.db);
  res.json({ ok: true });
});

app.post("/api/duty/manual", requireAuth, requirePermission("actions", "manageDutyHours", "Direktion"), (req, res) => {
  const userId = String(req.body.userId || "");
  const status = String(req.body.status || "Manuelle Korrektur").trim();
  const startedAt = String(req.body.startedAt || "").trim();
  const endedAt = String(req.body.endedAt || "").trim();
  const reason = String(req.body.reason || "").trim();
  if (!req.db.users.some((user) => user.id === userId)) return res.status(404).json({ error: "Benutzer nicht gefunden." });
  if (!startedAt || !endedAt || !reason) return res.status(400).json({ error: "Start, Ende und Grund sind Pflichtfelder." });
  const entry = { id: makeId("duty_manual"), userId, status, startedAt: new Date(startedAt).toISOString(), endedAt: new Date(endedAt).toISOString(), manual: true, reason, actorName: actorName(req.user) };
  req.db.dutyHistory.push(entry);
  logAction(req.db, req.user, "Dienstzeit hinzugefügt", status, { after: entry });
  writeDb(req.db);
  res.status(201).json({ entry });
});

app.delete("/api/duty/history/:id", requireAuth, requirePermission("actions", "manageDutyHours", "Direktion"), (req, res) => {
  res.status(403).json({ error: "Der Dienstzeiten-Log ist nicht löschbar." });
});

function endAllActiveDuty(db, actor, action = "Alle Dienste beendet") {
  const endedAt = nowIso();
  db.duty.forEach((active) => {
    const history = db.dutyHistory.find((entry) => entry.id === active.id) || db.dutyHistory.find((entry) => entry.userId === active.userId && !entry.endedAt);
    if (history) history.endedAt = endedAt;
    else db.dutyHistory.push({ ...active, endedAt, manual: false });
  });
  const count = db.duty.length;
  logAction(db, actor, action, "Dienstblatt", { count });
  db.duty = [];
  return count;
}

app.post("/api/seizures", requireAuth, (req, res) => {
  const suspect = String(req.body.suspect || "").trim();
  const location = String(req.body.location || "").trim();
  const numberValue = (value) => Math.max(0, Number(value || 0) || 0);
  const sourceType = ["Dealer", "Camper"].includes(String(req.body.sourceType || "").trim()) ? String(req.body.sourceType).trim() : "";
  const evidenceLinks = Array.isArray(req.body.evidenceLinks)
    ? req.body.evidenceLinks.map((item) => String(item || "").trim()).filter(Boolean)
    : String(req.body.evidenceLink || req.body.weapons || "").split("\n").map((item) => item.trim()).filter(Boolean);
  if (!suspect || !location || !evidenceLinks.length) {
    return res.status(400).json({ error: "Tatverdächtiger, Standort und mindestens ein Beweis sind Pflichtfelder." });
  }
  const entry = {
    id: makeId("seizure"),
    suspect,
    location,
    evidenceLinks,
    weapons: "",
    drugs: "",
    other: "",
    witness: String(req.body.witness || "").trim(),
    murder: Boolean(req.body.murder),
    blackMoney: numberValue(req.body.blackMoney),
    crates: numberValue(req.body.crates),
    sourceType,
    vehicleId: String(req.body.vehicleId || "").trim(),
    officerId: req.user.id,
    officerName: actorName(req.user),
    createdAt: nowIso()
  };
  req.db.settings.seizures = Array.isArray(req.db.settings.seizures) ? req.db.settings.seizures : [];
  req.db.settings.seizures.unshift(entry);
  logAction(req.db, req.user, "Beschlagnahmung erstellt", suspect, { after: entry });
  writeDb(req.db);
  res.status(201).json({ seizure: entry, settings: req.db.settings });
});

app.patch("/api/seizures/:id", requireAuth, (req, res) => {
  const entry = req.db.settings.seizures.find((item) => item.id === req.params.id);
  if (!entry) return res.status(404).json({ error: "Beschlagnahmung nicht gefunden." });
  const canEditAll = (rolePower[req.user.role] || 0) >= rolePower.Direktion;
  if (!canEditAll && entry.officerId !== req.user.id) return res.status(403).json({ error: "Keine Berechtigung." });
  const suspect = String(req.body.suspect || "").trim();
  const location = String(req.body.location || "").trim();
  const before = { ...entry };
  const numberValue = (value) => Math.max(0, Number(value || 0) || 0);
  const evidenceLinks = Array.isArray(req.body.evidenceLinks)
    ? req.body.evidenceLinks.map((item) => String(item || "").trim()).filter(Boolean)
    : [];
  if (!suspect || !location || !evidenceLinks.length) return res.status(400).json({ error: "Tatverdächtiger, Standort und mindestens ein Beweis sind Pflichtfelder." });
  Object.assign(entry, {
    suspect,
    location,
    evidenceLinks,
    weapons: "",
    drugs: "",
    other: "",
    witness: String(req.body.witness || "").trim(),
    murder: Boolean(req.body.murder),
    blackMoney: numberValue(req.body.blackMoney),
    crates: numberValue(req.body.crates),
    sourceType: ["Dealer", "Camper"].includes(String(req.body.sourceType || "").trim()) ? String(req.body.sourceType).trim() : "",
    vehicleId: String(req.body.vehicleId || "").trim(),
    updatedAt: nowIso(),
    updatedBy: actorName(req.user)
  });
  logAction(req.db, req.user, "Beschlagnahmung bearbeitet", suspect, { before, after: entry });
  writeDb(req.db);
  res.json({ seizure: entry, settings: req.db.settings });
});

app.delete("/api/seizures/:id", requireAuth, requireRole("Direktion"), (req, res) => {
  const entry = req.db.settings.seizures.find((item) => item.id === req.params.id);
  if (!entry) return res.status(404).json({ error: "Beschlagnahmung nicht gefunden." });
  req.db.settings.seizures = req.db.settings.seizures.filter((item) => item.id !== req.params.id);
  logAction(req.db, req.user, "Beschlagnahmung gelöscht", entry.suspect || req.params.id, { before: entry });
  writeDb(req.db);
  res.json({ ok: true, settings: req.db.settings });
});

app.post("/api/calendar/events", requireAuth, (req, res) => {
  const title = String(req.body.title || "").trim();
  const description = String(req.body.description || "").trim();
  const startDate = String(req.body.startDate || "").trim();
  const startTime = String(req.body.startTime || "").trim();
  const endDate = String(req.body.endDate || startDate).trim();
  const endTime = String(req.body.endTime || startTime).trim();
  const type = String(req.body.type || "Allgemein").trim();
  const color = String(req.body.color || "Blau").trim();
  const location = String(req.body.location || "").trim();
  const reminder = String(req.body.reminder || "30 Minuten").trim();
  const allDay = Boolean(req.body.allDay);
  if (!title || !startDate || (!allDay && !startTime)) return res.status(400).json({ error: "Titel, Startdatum und Startzeit sind Pflichtfelder." });
  const event = {
    id: makeId("calendar"),
    title,
    description,
    startDate,
    startTime: allDay ? "" : startTime,
    endDate: endDate || startDate,
    endTime: allDay ? "" : endTime,
    type,
    color,
    location,
    reminder,
    allDay,
    authorName: actorName(req.user),
    createdAt: nowIso()
  };
  req.db.settings.calendarEvents.unshift(event);
  logAction(req.db, req.user, "Kalendertermin erstellt", title, { after: event });
  writeDb(req.db);
  res.status(201).json({ event });
});

app.patch("/api/calendar/events/:id", requireAuth, (req, res) => {
  const event = req.db.settings.calendarEvents.find((item) => item.id === req.params.id);
  if (!event) return res.status(404).json({ error: "Termin nicht gefunden." });
  const before = { ...event };
  const title = String(req.body.title || "").trim();
  const startDate = String(req.body.startDate || "").trim();
  const startTime = String(req.body.startTime || "").trim();
  const allDay = Boolean(req.body.allDay);
  if (!title || !startDate || (!allDay && !startTime)) return res.status(400).json({ error: "Titel, Startdatum und Startzeit sind Pflichtfelder." });
  Object.assign(event, {
    title,
    description: String(req.body.description || "").trim(),
    startDate,
    startTime: allDay ? "" : startTime,
    endDate: String(req.body.endDate || startDate).trim(),
    endTime: allDay ? "" : String(req.body.endTime || startTime).trim(),
    type: String(req.body.type || "Allgemein").trim(),
    color: String(req.body.color || "Blau").trim(),
    location: String(req.body.location || "").trim(),
    reminder: String(req.body.reminder || "30 Minuten").trim(),
    allDay,
    updatedAt: nowIso()
  });
  logAction(req.db, req.user, "Kalendertermin bearbeitet", title, { before, after: event });
  writeDb(req.db);
  res.json({ event });
});

app.delete("/api/calendar/events/:id", requireAuth, (req, res) => {
  const event = req.db.settings.calendarEvents.find((item) => item.id === req.params.id);
  req.db.settings.calendarEvents = req.db.settings.calendarEvents.filter((item) => item.id !== req.params.id);
  logAction(req.db, req.user, "Kalendertermin gelöscht", event?.title || req.params.id, { before: event || null });
  writeDb(req.db);
  res.json({ ok: true });
});

app.get("*", (_req, res) => {
  res.setHeader("Content-Type", "text/html; charset=utf-8");
  res.sendFile(path.join(PUBLIC_DIR, "index.html"));
});

function currentBerlinRestartWindow() {
  const parts = new Intl.DateTimeFormat("de-DE", {
    timeZone: "Europe/Berlin",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false
  }).formatToParts(new Date()).reduce((acc, part) => ({ ...acc, [part.type]: part.value }), {});
  return {
    date: `${parts.year}-${parts.month}-${parts.day}`,
    time: `${parts.hour}:${parts.minute}`
  };
}

function slugify(value, prefix = "seite") {
  const slug = String(value || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
  return slug || prefix;
}

function runScheduledRestarts() {
  try {
    const db = readDb();
    const times = db.settings.restartTimes || [];
    if (!times.length) return;
    const { date, time } = currentBerlinRestartWindow();
    if (!times.includes(time)) return;
    db.settings.restartLastRun = db.settings.restartLastRun || {};
    if (db.settings.restartLastRun[time] === date) return;
    const count = endAllActiveDuty(db, { firstName: "System", lastName: "Restart" }, "Restart: Dienste automatisch beendet");
    db.settings.restartLastRun[time] = date;
    if (count > 0) writeDb(db);
    else writeDb(db);
  } catch (error) {
    console.error("Restart scheduler failed:", error);
  }
}

ensureStorage();
runScheduledRestarts();
setInterval(runScheduledRestarts, 30000);
app.listen(PORT, () => {
  console.log(`LSPD Dienstblatt laeuft auf http://localhost:${PORT}`);
});
