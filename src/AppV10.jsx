import React, { useState, useEffect, useMemo, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";

// Local persistence (replaces the sandbox storage API when running as a real site).
const store = {
  get: async (k) => { const v = localStorage.getItem(k); return v ? { value: v } : null; },
  set: async (k, v) => { localStorage.setItem(k, v); return { value: v }; },
};

// ————————————————————————————————————————————————————————————————
// All In Fish Tank — Weekly Settlement Engine (v4)
// Settlements · umbrella & SA reports · house-backed ledger
// (makeup nets = P&L + RB credit; action buys) · Ak-Jon recon ·
// light/dark theme. Config persists week to week.
// ————————————————————————————————————————————————————————————————

const PALETTES = {
  light: {
    "--paper": "#F7F2E8", "--card": "#FFFDF8", "--cream": "#F3EAD8",
    "--gold": "#B49A5E", "--goldAccent": "#8C7440", "--ink": "#2B241C", "--bar": "#332A22",
    "--green": "#1E7B34", "--red": "#C0392B", "--mute": "#8A7E6C", "--line": "#E4D9C2",
    "--rowAlt": "#FBF6EB", "--banner": "#FBF3D9", "--surface": "#FFFFFF", "--onGold": "#FFFFFF",
    "--barText": "#F2E7CC", "--barMute": "#C9BB9A", "--barSubtle": "#B4A585", "--barGold": "#F0C97A",
    "--barGreen": "#7CCB8B", "--barRed": "#F0958A",
    "--pillRedBg": "#F9E4E1", "--pillGreenBg": "#E3F1E5", "--pillBlueBg": "#E2ECF4", "--pillGoldBg": "#F1E8D3",
    "--pillBlueFg": "#31587A", "--chipOff": "#C9BFA9", "--errBg": "#F9E4E1",
  },
  dark: {
    "--paper": "#131009", "--card": "#1D180F", "--cream": "#262013", "--gold": "#D4B36A",
    "--goldAccent": "#E6C88A", "--ink": "#F3EAD6", "--bar": "#0B0906",
    "--green": "#6FD08C", "--red": "#EF8677", "--mute": "#A79A80", "--line": "#332B1C",
    "--rowAlt": "#231D10", "--banner": "#2E2614", "--surface": "#2A2313", "--onGold": "#171207",
    "--barText": "#F6ECD2", "--barMute": "#BCAD8C", "--barSubtle": "#A29476", "--barGold": "#F4CE7E",
    "--barGreen": "#84D695", "--barRed": "#F49C90",
    "--pillRedBg": "#43231C", "--pillGreenBg": "#1C3A24", "--pillBlueBg": "#1E2E3D", "--pillBlueFg": "#A6CBEA",
    "--pillGoldBg": "#3B3013", "--chipOff": "#5F5439", "--errBg": "#43231C",
  },
};
const C = {
  paper: "var(--paper)", card: "var(--card)", cream: "var(--cream)",
  gold: "var(--gold)", goldDark: "var(--goldAccent)", ink: "var(--ink)", bar: "var(--bar)",
  green: "var(--green)", red: "var(--red)", mute: "var(--mute)", line: "var(--line)",
  rowAlt: "var(--rowAlt)", banner: "var(--banner)", surface: "var(--surface)",
};

const DEFAULT_SA_DEALS = {
  "9416-1077": 80, "9956-9064": 75, "4590-9906": 75, "5207-4267": 82.5,
  "9812-1646": 80, "1788-2643": 75, "2340-6362": 65, "6340-0999": 75,
  "8775-1559": 75, "8389-0206": 80, "5532-5157": 80, "5349-7456": 100,
  "4243-3806": 80, "4617-6330": 90, "5805-5568": 80, "9818-6720": 75,
  "2981-4306": 65, "5349-5156": 90, "4091-6935": 80, "1581-1690": 75,
  "8474-9014": 80, "3011-0934": 70, "1652-5578": 90,
};
const DEFAULT_PLAYER_DEALS = {
  "2859-2602": 70, "1793-1073": 40, "3409-6651": 70,
  "8835-5140": 70, "7623-2110": 50, "7959-6530": 70,
};
const DEFAULT_NAMES = {
  "9416-1077": "SharpCheddar", "9956-9064": "FishSupport", "4590-9906": "CUMROCKET",
  "5207-4267": "H8Varience", "9812-1646": "Cashlover777", "1788-2643": "Animal3391",
  "2340-6362": "Humpback679", "6340-0999": "Leaderzay", "8775-1559": "Hannibal0",
  "8389-0206": "DiorSauvage", "5532-5157": "Rlawnsgud", "5349-7456": "punterx",
  "4243-3806": "shortbusbully88", "4617-6330": "BroadwaySupport", "5805-5568": "catdad777",
  "9818-6720": "wheatie DOG", "2981-4306": "WrongCalc", "5349-5156": "imnotpunting",
  "4091-6935": "GamadGadol", "1581-1690": "TheBettor", "8474-9014": "Ved_13",
  "3011-0934": "LuckilyLucky", "1652-5578": "Sus Moustache",
  "2859-2602": "krishdhawan", "1793-1073": "RobinhoodAP", "3409-6651": "syao12",
  "8835-5140": "Biggest Donk", "7623-2110": "soggy waffles", "7959-6530": "CorporalToenail",
};

const DEFAULT_CONFIG = {
  theme: "dark", themeV2: true,
  defaultTB: 80,
  saDeals: { ...DEFAULT_SA_DEALS },
  playerDeals: { ...DEFAULT_PLAYER_DEALS },
  // Action buys on regular players: house taxes pct% of wins / rebates pct% of losses.
  actionTax: {}, // memberId -> { pct, backer: 'split'|'ak'|'jon' }
  confirmedSAs: Object.fromEntries(Object.keys(DEFAULT_SA_DEALS).map((k) => [k, true])),
  confirmedPlayers: Object.fromEntries(Object.keys(DEFAULT_PLAYER_DEALS).map((k) => [k, true])),
  names: { ...DEFAULT_NAMES },
  ownAccounts: { ak: ["axe7777", "gimmezemoneys"], jon: ["Flashbrook123"] },
  // 'makeup': weekly net = P&L + RB credit. Above makeup → player paid their %
  //           of the excess, backer books the rest. Below → no cash, net accrues
  //           to makeup on the backer's book.
  // 'action': backer owns actionPct% of (P&L + RB); player settles the rest.
  backed: {
    dingleberry23: { name: "dingleberry23", deal: "makeup", rbNormal: 75, rbMakeup: 75, makeup: 0, playerProfitPct: 50, backer: "split" },
    wjewje12:      { name: "wjewje12", deal: "makeup", rbNormal: 65, rbMakeup: 100, makeup: 0, playerProfitPct: 50, backer: "split" },
    niceblufflol:  { name: "niceblufflol", deal: "makeup", rbNormal: 65, rbMakeup: 100, makeup: 0, playerProfitPct: 50, backer: "jon" },
    gigapuntwhale: { name: "gigapuntwhale", deal: "action", actionPct: 50, rbPct: 100, backer: "jon" },
    gruzzy:        { name: "gruzzy", deal: "action", actionPct: 75, rbPct: 100, backer: "jon" },
    jumpingguppy:  { name: "jumpingguppy", deal: "action", actionPct: 100, rbPct: 100, backer: "jon" },
  },
  umbrellas: [
    { id: "u-vivaan", name: "Vivaan Rastoghi", saIds: ["1788-2643", "4243-3806", "4617-6330", "9818-6720", "5349-5156", "4091-6935", "1581-1690", "1652-5578"] },
    { id: "u-penpaper", name: "Penpaper", saIds: ["9812-1646", "8389-0206", "5805-5568"] },
    { id: "u-snorlax", name: "Snorlax", saIds: ["4590-9906", "8775-1559", "5532-5157"] },
  ],
  assignments: {},
  fees: [
    { id: "acct", label: "Accountant", pct: 5, recipient: "external", paidBy: "split" },
    { id: "punterx", label: "punterx (open-sitting)", pct: 5, recipient: "external", paidBy: "split" },
    { id: "gruzzy", label: "gruzzy", pct: 1, recipient: "jon", paidBy: "split" },
  ],
  feeBase: "net",
  finalizedPeriods: {},
};

const fmt = (n, d = 2) =>
  (n < 0 ? "-" : "") + Math.abs(n).toLocaleString("en-US", { minimumFractionDigits: d, maximumFractionDigits: d });
const fmtI = (n) => n.toLocaleString("en-US");
const money = (n) => (
  <span style={{ color: n > 0.005 ? C.green : n < -0.005 ? C.red : C.ink, fontWeight: 600 }}>{fmt(n)}</span>
);

// ———————————————— Parsing ————————————————
function parseWorkbook(buf) {
  const wb = XLSX.read(buf, { type: "array" });
  const sheetName = wb.SheetNames.find((s) => s.toLowerCase().includes("club overview")) || wb.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: null });
  let headerIdx = -1, period = "";
  for (let i = 0; i < Math.min(rows.length, 20); i++) {
    const c0 = String(rows[i]?.[0] ?? "");
    if (c0.startsWith("Period")) period = c0.replace("Period :", "").trim();
    if (c0 === "No.") { headerIdx = i; break; }
  }
  if (headerIdx < 0) throw new Error("Couldn't find the 'No.' header row on the Club Overview sheet. Is this the standard weekly export?");
  const players = [];
  for (let i = headerIdx + 2; i < rows.length; i++) {
    const r = rows[i];
    // Skip blanks and summary rows (e.g. trailing TOTAL): row number must be
    // numeric and the member ID must be present.
    if (r?.[0] == null || !Number.isFinite(parseFloat(r[0]))) continue;
    const memberId = String(r[7] ?? "").trim();
    const name = String(r[8] ?? "").trim();
    if (!memberId || memberId === "-" || !name) continue;
    const num = (v) => { const x = parseFloat(v); return isNaN(x) ? 0 : x; };
    const p = {
      saId: String(r[1] ?? "-").trim(), saName: String(r[2] ?? "-").trim(),
      agentId: String(r[3] ?? "-").trim(), agentName: String(r[4] ?? "-").trim(),
      role: String(r[6] ?? "").trim(), memberId, name,
      hands: num(r[10]), fee: num(r[11]), pnl: num(r[19]),
    };
    if (p.hands > 0 || p.fee !== 0 || p.pnl !== 0) players.push(p);
  }
  if (!players.length) throw new Error("Parsed the sheet but found no active players (rows with hands, fee, or P&L).");
  return { players, period };
}

// ———————————————— Engine ————————————————
function buildModel(players, cfg, weekAdj, period) {
  // If this period was finalized, makeup math freezes at the balances it was
  // settled with (otherwise finalize would double-count the same week).
  const finEntry = period ? (cfg.finalizedPeriods || {})[period] : null;
  const snapMakeup = finEntry && typeof finEntry === "object" ? finEntry.snapshot || null : null;
  const ownMap = {};
  cfg.ownAccounts.ak.forEach((n) => (ownMap[n.trim().toLowerCase()] = "ak"));
  cfg.ownAccounts.jon.forEach((n) => (ownMap[n.trim().toLowerCase()] = "jon"));
  const backedMap = cfg.backed || {};

  const ownRows = [], backedRows = [], extRows = [];
  for (const p of players) {
    const lo = p.name.toLowerCase();
    if (ownMap[lo]) ownRows.push({ ...p, owner: ownMap[lo] });
    else if (backedMap[lo]) backedRows.push({ ...p, backedKey: lo });
    else extRows.push(p);
  }

  const splitTipback = (fee, adj) =>
    (Math.min(adj.amtA, fee) * adj.rateA) / 100 + (Math.max(0, fee - adj.amtA) * adj.rateB) / 100;

  // Owner accounts: 100% feeback.
  const own = ownRows.map((p) => ({ ...p, feeback: p.fee, position: p.pnl + p.fee }));

  // Normal external players (with optional action-buy tax/rebate on P&L)
  const ext = extRows.map((p) => {
    const tbPct = cfg.playerDeals[p.memberId] ?? (p.saId !== "-" ? cfg.saDeals[p.saId] : undefined) ?? cfg.defaultTB;
    const adj = weekAdj?.[`p:${p.memberId}`];
    const tipback = adj ? splitTipback(p.fee, adj) : (p.fee * tbPct) / 100;
    const at = (cfg.actionTax || {})[p.memberId];
    const net = p.pnl + tipback;
    const actionCut = at ? (net * at.pct) / 100 : 0; // pct% of net after rakeback: taxes wins, rebates losses
    return { ...p, tbPct, adjusted: !!adj, tipback, actionTaxPct: at ? at.pct : 0, actionBacker: at ? at.backer || "split" : null, actionCut, settlement: net - actionCut };
  });

  const saMap = new Map(); const individuals = [];
  for (const p of ext) {
    if (p.saId !== "-") {
      if (!saMap.has(p.saId)) saMap.set(p.saId, { key: `sa:${p.saId}`, type: "sa", id: p.saId, name: p.saName, members: [] });
      saMap.get(p.saId).members.push(p);
    } else {
      individuals.push({ key: `p:${p.memberId}`, type: "player", id: p.memberId, name: p.name, members: [p] });
    }
  }
  const agg = (e) => {
    const sum = (f) => e.members.reduce((a, m) => a + m[f], 0);
    return { ...e, hands: sum("hands"), pnl: sum("pnl"), fee: sum("fee"), tipback: sum("tipback"), settlement: sum("settlement") };
  };
  const saEntities = [...saMap.values()].map((e) => {
    let x = agg(e);
    const adj = weekAdj?.[e.key];
    if (adj) { const tb = splitTipback(x.fee, adj); x = { ...x, tipback: tb, settlement: x.pnl + tb, adjusted: true }; }
    return x;
  });
  const indEntities = individuals.map(agg).map((x) => ({ ...x, adjusted: x.members.some((m) => m.adjusted) }));

  // House-backed entities
  const backedEntities = backedRows.map((p) => {
    const b = backedMap[p.backedKey];
    const key = `b:${p.backedKey}`;
    const adj = weekAdj?.[key];
    if (b.deal === "action") {
      const rbTotal = adj ? splitTipback(p.fee, adj) : (p.fee * (b.rbPct ?? 100)) / 100;
      const gross = p.pnl + rbTotal;
      const backerShare = (gross * b.actionPct) / 100;
      const settlement = gross - backerShare;
      return {
        key, type: "backed", dealType: "action", id: p.memberId, name: p.name,
        members: [{ ...p, tbPct: b.rbPct ?? 100, tipback: rbTotal, settlement }],
        hands: p.hands, pnl: p.pnl, fee: p.fee, tipback: rbTotal, settlement,
        backer: b.backer, backerBook: backerShare, actionPct: b.actionPct, adjusted: !!adj,
      };
    }
    // Makeup deal: weekly net = P&L + RB credit.
    const entering = snapMakeup && snapMakeup[key.slice(2)] != null ? snapMakeup[key.slice(2)] : (b.makeup || 0);
    const inMakeup = entering > 0.005;
    const rb = inMakeup ? b.rbMakeup : b.rbNormal;
    const rbCredit = adj ? splitTipback(p.fee, adj) : (p.fee * rb) / 100;
    const net = p.pnl + rbCredit;
    const excess = Math.max(0, net - entering);
    const playerCash = (excess * (b.playerProfitPct ?? 50)) / 100;
    const makeupAfter = Math.max(0, entering - net);
    return {
      key, type: "backed", dealType: "makeup", id: p.memberId, name: p.name,
      members: [{ ...p, tbPct: rb, tipback: rbCredit, settlement: playerCash }],
      hands: p.hands, pnl: p.pnl, fee: p.fee, tipback: rbCredit, settlement: playerCash,
      backer: b.backer, backerBook: net - playerCash, net,
      rb, inMakeup, makeupBefore: entering, makeupAfter, adjusted: !!adj,
    };
  });

  // Umbrella merge
  const umbrellaOf = {}; (cfg.umbrellas || []).forEach((u) => u.saIds.forEach((id) => (umbrellaOf[id] = u)));
  const umbMap = new Map(); const looseSAs = [];
  for (const e of saEntities) {
    const u = umbrellaOf[e.id];
    if (u) {
      if (!umbMap.has(u.id)) umbMap.set(u.id, { key: `u:${u.id}`, type: "umbrella", id: u.id, name: u.name, subgroups: [], members: [] });
      const m = umbMap.get(u.id); m.subgroups.push(e); m.members.push(...e.members);
    } else looseSAs.push(e);
  }
  const umbEntities = [...umbMap.values()].map((e) => {
    const sum = (f) => e.subgroups.reduce((a, s) => a + s[f], 0);
    return { ...e, hands: sum("hands"), pnl: sum("pnl"), fee: sum("fee"), tipback: sum("tipback"), settlement: sum("settlement"), adjusted: e.subgroups.some((s) => s.adjusted) };
  });

  const entities = [...umbEntities, ...looseSAs, ...backedEntities, ...indEntities]
    .sort((a, b) => Math.abs(b.fee) - Math.abs(a.fee));

  // Club economics. For makeup players the RB credit leaves club profit and
  // moves onto the backer's book, so it counts as a tipback here.
  const clubRevenue = players.reduce((a, p) => a + p.fee, 0);
  const extTipbacks = ext.reduce((a, p) => a + p.tipback, 0);
  const backedRB = backedEntities.reduce((a, e) => a + e.tipback, 0);
  const ownFeeback = own.reduce((a, p) => a + p.feeback, 0);
  const tipbacksPaid = extTipbacks + backedRB + ownFeeback;
  const clubProfit = clubRevenue - tipbacksPaid;
  const feeRows = cfg.fees.map((f) => ({ ...f, amount: f.kind === "fixed" ? (f.amount || 0) : ((cfg.feeBase === "gross" ? clubRevenue : clubProfit) * f.pct) / 100 }));
  const totalFees = feeRows.reduce((a, f) => a + f.amount, 0);
  const netProfit = clubProfit - totalFees;

  const ownPosition = { ak: 0, jon: 0 };
  own.forEach((p) => (ownPosition[p.owner] += p.position));

  const backedBook = { ak: 0, jon: 0 };
  backedEntities.forEach((e) => {
    if (e.backer === "split") { backedBook.ak += e.backerBook / 2; backedBook.jon += e.backerBook / 2; }
    else backedBook[e.backer] += e.backerBook;
  });

  const taxBook = { ak: 0, jon: 0 };
  ext.forEach((p) => {
    if (!p.actionCut) return;
    if (p.actionBacker === "ak" || p.actionBacker === "jon") taxBook[p.actionBacker] += p.actionCut;
    else { taxBook.ak += p.actionCut / 2; taxBook.jon += p.actionCut / 2; }
  });

  const entitle = {
    ak: netProfit / 2 + ownPosition.ak + backedBook.ak + taxBook.ak,
    jon: netProfit / 2 + ownPosition.jon + backedBook.jon + taxBook.jon,
  };
  feeRows.forEach((f) => { if (f.recipient === "ak" || f.recipient === "jon") entitle[f.recipient] += f.amount; });

  const actual = { ak: 0, jon: 0 }; const unassigned = [];
  entities.forEach((e) => {
    const who = cfg.assignments[e.key];
    if (who === "ak" || who === "jon") actual[who] += -e.settlement;
    else unassigned.push(e);
  });
  feeRows.forEach((f) => {
    if (f.recipient === "external") {
      if (f.paidBy === "split") { actual.ak -= f.amount / 2; actual.jon -= f.amount / 2; }
      else actual[f.paidBy] -= f.amount;
    }
  });

  const akOwesJon = actual.ak - entitle.ak;
  const balanceOk = Math.abs(actual.ak + actual.jon - (entitle.ak + entitle.jon)) < 0.02 && unassigned.length === 0;

  const totals = {
    hands: entities.reduce((a, e) => a + e.hands, 0),
    pnl: entities.reduce((a, e) => a + e.pnl, 0),
    fee: entities.reduce((a, e) => a + e.fee, 0),
    tipback: entities.reduce((a, e) => a + e.tipback, 0),
    settlement: entities.reduce((a, e) => a + e.settlement, 0),
  };

  return { own, entities, backedEntities, saEntities, indEntities, umbEntities, looseSAs,
    clubRevenue, extTipbacks, backedRB, ownFeeback, tipbacksPaid, clubProfit, feeRows, totalFees, netProfit,
    ownPosition, backedBook, taxBook, entitle, actual, akOwesJon, unassigned, balanceOk, totals };
}

// ———————————————— Styled Excel export (ExcelJS) ————————————————
const n2 = (x) => Math.round((x + Number.EPSILON) * 100) / 100;
const MONEY_FMT = "#,##0.00;(#,##0.00)";
const XLC = { gold: "FFB49A5E", bar: "FF332A22", cream: "FFF3EAD8", rowAlt: "FFFBF6EB", white: "FFFFFFFF", ink: "FF2B241C", green: "FF1E7B34", red: "FFC0392B", mute: "FF8A7E6C" };
const FBASE = { name: "Arial", size: 10, color: { argb: XLC.ink } };
const fillOf = (argb) => ({ type: "pattern", pattern: "solid", fgColor: { argb } });

function xTitle(ws, text) {
  const r = ws.addRow([text]);
  r.getCell(1).font = { ...FBASE, size: 13, bold: true };
  ws.addRow([]);
}
function xHeader(ws, labels, leftCols = 4) {
  const r = ws.addRow(labels);
  r.eachCell((c, i) => {
    c.font = { name: "Arial", size: 9, bold: true, color: { argb: XLC.white } };
    c.fill = fillOf(XLC.gold);
    c.alignment = { horizontal: i <= leftCols ? "left" : "right" };
  });
  return r.number;
}
function xMoney(cell, v, { bold = false, white = false, colorSign = true } = {}) {
  cell.value = n2(v);
  cell.numFmt = MONEY_FMT;
  const color = white ? XLC.white : colorSign && v > 0.005 ? XLC.green : colorSign && v < -0.005 ? XLC.red : XLC.ink;
  cell.font = { ...FBASE, bold, color: { argb: color } };
  cell.alignment = { horizontal: "right" };
}
function xNum(cell, v, fmt, { bold = false, white = false } = {}) {
  cell.value = v; cell.numFmt = fmt;
  cell.font = { ...FBASE, bold, color: { argb: white ? XLC.white : XLC.ink } };
  cell.alignment = { horizontal: "right" };
}
function xText(cell, v, { bold = false, white = false, mute = false } = {}) {
  cell.value = v;
  cell.font = { ...FBASE, bold, color: { argb: white ? XLC.white : mute ? XLC.mute : XLC.ink } };
  cell.alignment = { horizontal: "left" };
}
const DEAL_COLS = ["Player", "Device ID", "Super Agent", "Agent", "Hands", "Winnings", "Tips", "TB %", "Tipback", "Settlement"];
const DEAL_W = [20, 12, 17, 17, 9, 12, 12, 8, 12, 13];

function memberRow(ws, m, alt) {
  const r = ws.addRow([]);
  xText(r.getCell(1), m.name, { bold: true });
  xText(r.getCell(2), m.memberId, { mute: true });
  xText(r.getCell(3), m.saName, { mute: true });
  xText(r.getCell(4), m.agentName, { mute: true });
  xNum(r.getCell(5), m.hands, "#,##0");
  xMoney(r.getCell(6), m.pnl, { colorSign: false });
  xMoney(r.getCell(7), m.fee, { colorSign: false });
  xNum(r.getCell(8), m.tbPct, '0.0"%"');
  xMoney(r.getCell(9), m.tipback, { colorSign: false });
  xMoney(r.getCell(10), m.settlement);
  if (alt) r.eachCell({ includeEmpty: true }, (c) => { if (!c.fill || !c.fill.fgColor) c.fill = fillOf(XLC.rowAlt); });
  return r;
}
function totalRow(ws, label, e, { dark = false } = {}) {
  const r = ws.addRow([]);
  xText(r.getCell(1), label, { bold: true, white: dark });
  xNum(r.getCell(5), e.hands, "#,##0", { bold: true, white: dark });
  xMoney(r.getCell(6), e.pnl, { bold: true, white: dark, colorSign: !dark });
  xMoney(r.getCell(7), e.fee, { bold: true, white: dark, colorSign: false });
  xNum(r.getCell(8), e.fee ? n2((e.tipback / e.fee) * 100) : 0, '0.0"%"', { bold: true, white: dark });
  xMoney(r.getCell(9), e.tipback, { bold: true, white: dark, colorSign: false });
  xMoney(r.getCell(10), e.settlement, { bold: true, white: dark, colorSign: !dark });
  r.eachCell({ includeEmpty: true }, (c) => { c.fill = fillOf(dark ? XLC.bar : XLC.cream); });
  for (let j = 1; j <= 10; j++) if (!r.getCell(j).value && r.getCell(j).value !== 0) r.getCell(j).fill = fillOf(dark ? XLC.bar : XLC.cream);
  return r;
}
function safeSheetName(base, wb) {
  let nm = String(base).replace(/[\\/?*[\]:]/g, " ").trim().slice(0, 28) || "Sheet";
  let out = nm, i = 2;
  while (wb.worksheets.some((w) => w.name === out)) out = `${nm.slice(0, 25)} ${i++}`;
  return out;
}
function addDealSheet(wb, e, period) {
  const ws = wb.addWorksheet(safeSheetName(e.name, wb));
  DEAL_W.forEach((w, i) => (ws.getColumn(i + 1).width = w));
  xTitle(ws, `${e.name} — ${period || "this week"}`);
  const hr = xHeader(ws, DEAL_COLS);
  ws.views = [{ state: "frozen", ySplit: hr }];
  if (e.type === "umbrella") {
    e.subgroups.forEach((s) => {
      const sub = ws.addRow([]);
      xText(sub.getCell(1), `${s.name} — settlement ${fmt(s.settlement)}`, { bold: true });
      sub.eachCell({ includeEmpty: true }, (c) => (c.fill = fillOf(XLC.rowAlt)));
      for (let j = 1; j <= 10; j++) sub.getCell(j).fill = fillOf(XLC.rowAlt);
      [...s.members].sort((a, b) => b.fee - a.fee).forEach((m, i) => memberRow(ws, m, false));
    });
  } else {
    [...e.members].sort((a, b) => b.fee - a.fee).forEach((m, i) => memberRow(ws, m, i % 2 === 1));
  }
  totalRow(ws, "TOTAL", e);
  return ws;
}
async function saveWb(wb, filename) {
  const buf = await wb.xlsx.writeBuffer();
  const url = URL.createObjectURL(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
  const a = document.createElement("a"); a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}
const periodSlug = (period) => (period || "week").replace(/[^\d]/g, "_").replace(/^_+|_+$/g, "") || "week";

async function downloadDealExcel(e, period) {
  const wb = new ExcelJS.Workbook();
  addDealSheet(wb, e, period);
  await saveWb(wb, `${e.name.replace(/[^\w]+/g, "_")}_${periodSlug(period)}.xlsx`);
}

async function downloadWorkbook(model, period, cfg) {
  const wb = new ExcelJS.Workbook();
  // — Settlements summary
  const ws = wb.addWorksheet("Settlements");
  [24, 14, 9, 13, 13, 9, 13, 14].forEach((w, i) => (ws.getColumn(i + 1).width = w));
  xTitle(ws, `All In Fish Tank — Settlements ${period || ""}`.trim());
  const hr = xHeader(ws, ["Deal", "Type", "Hands", "Winnings", "Tips", "Avg TB %", "Tipback", "Settlement"], 2);
  ws.views = [{ state: "frozen", ySplit: hr }];
  model.entities.forEach((e, i) => {
    const r = ws.addRow([]);
    xText(r.getCell(1), e.name, { bold: true });
    xText(r.getCell(2), e.type === "umbrella" ? "Umbrella" : e.type === "sa" ? "Super Agent" : e.type === "backed" ? (e.dealType === "action" ? "Action buy" : "Makeup deal") : "Player", { mute: true });
    xNum(r.getCell(3), e.hands, "#,##0");
    xMoney(r.getCell(4), e.pnl, { colorSign: false });
    xMoney(r.getCell(5), e.fee, { colorSign: false });
    xNum(r.getCell(6), e.fee ? n2((e.tipback / e.fee) * 100) : 0, '0.0"%"');
    xMoney(r.getCell(7), e.tipback, { colorSign: false });
    xMoney(r.getCell(8), e.settlement);
    if (i % 2 === 1) for (let j = 1; j <= 8; j++) r.getCell(j).fill = fillOf(XLC.rowAlt);
  });
  const t = model.totals;
  const gr = ws.addRow([]);
  xText(gr.getCell(1), "GRAND TOTAL", { bold: true, white: true });
  xNum(gr.getCell(3), t.hands, "#,##0", { bold: true, white: true });
  xMoney(gr.getCell(4), t.pnl, { bold: true, white: true, colorSign: false });
  xMoney(gr.getCell(5), t.fee, { bold: true, white: true, colorSign: false });
  xNum(gr.getCell(6), t.fee ? n2((t.tipback / t.fee) * 100) : 0, '0.0"%"', { bold: true, white: true });
  xMoney(gr.getCell(7), t.tipback, { bold: true, white: true, colorSign: false });
  xMoney(gr.getCell(8), t.settlement, { bold: true, white: true, colorSign: false });
  for (let j = 1; j <= 8; j++) gr.getCell(j).fill = fillOf(XLC.bar);

  // — one sheet per umbrella / loose SA
  [...model.umbEntities, ...model.looseSAs].forEach((e) => addDealSheet(wb, e, period));

  // — Individuals
  if (model.indEntities.length) {
    const wi = wb.addWorksheet("Individuals");
    DEAL_W.forEach((w, i) => (wi.getColumn(i + 1).width = w));
    xTitle(wi, `Individual players (no super agent) — ${period || "this week"}`);
    xHeader(wi, DEAL_COLS);
    model.indEntities.forEach((e, i) => memberRow(wi, { ...e.members[0], saName: "-" }, i % 2 === 1));
  }

  // — House-backed
  const bMk = model.backedEntities.filter((e) => e.dealType === "makeup");
  const bAc = model.backedEntities.filter((e) => e.dealType === "action");
  if (bMk.length || bAc.length) {
    const wbk = wb.addWorksheet("House-Backed");
    [17, 15, 12, 8, 9, 12, 12, 12, 12, 13, 13, 13].forEach((w, i) => (wbk.getColumn(i + 1).width = w));
    xTitle(wbk, `House-backed players — ${period || "this week"}`);
    if (bMk.length) {
      const note = wbk.addRow(["Makeup deals — net = P&L + RB credit; settlement = player % of net above makeup; shortfall accrues to makeup."]);
      note.getCell(1).font = { ...FBASE, size: 9, color: { argb: XLC.mute } };
      xHeader(wbk, ["Player", "Backer", "Makeup in", "RB %", "Hands", "P&L", "Tips", "RB credit", "Net", "Settlement", "Backer book", "Makeup after"], 2);
      bMk.forEach((e, i) => {
        const r = wbk.addRow([]);
        xText(r.getCell(1), e.name, { bold: true });
        xText(r.getCell(2), e.backer === "split" ? "Ak & Jon 50/50" : e.backer === "ak" ? "Ak" : "Jon", { mute: true });
        xMoney(r.getCell(3), e.makeupBefore, { colorSign: false });
        xNum(r.getCell(4), e.rb, '0.0"%"');
        xNum(r.getCell(5), e.hands, "#,##0");
        xMoney(r.getCell(6), e.pnl);
        xMoney(r.getCell(7), e.fee, { colorSign: false });
        xMoney(r.getCell(8), e.tipback, { colorSign: false });
        xMoney(r.getCell(9), e.net);
        xMoney(r.getCell(10), e.settlement, { colorSign: false });
        xMoney(r.getCell(11), e.backerBook);
        xMoney(r.getCell(12), e.makeupAfter, { colorSign: false });
        if (i % 2 === 1) for (let j = 1; j <= 12; j++) r.getCell(j).fill = fillOf(XLC.rowAlt);
      });
      wbk.addRow([]);
    }
    if (bAc.length) {
      const note = wbk.addRow(["Action buys — backer owns their % of (P&L + rakeback); player settles the rest."]);
      note.getCell(1).font = { ...FBASE, size: 9, color: { argb: XLC.mute } };
      xHeader(wbk, ["Player", "Backer", "Backer %", "RB %", "Hands", "P&L", "Tips", "Player settlement", "Backer book"], 2);
      bAc.forEach((e, i) => {
        const r = wbk.addRow([]);
        xText(r.getCell(1), e.name, { bold: true });
        xText(r.getCell(2), e.backer === "ak" ? "Ak" : "Jon", { mute: true });
        xNum(r.getCell(3), e.actionPct, '0.0"%"');
        xNum(r.getCell(4), e.members[0].tbPct, '0.0"%"');
        xNum(r.getCell(5), e.hands, "#,##0");
        xMoney(r.getCell(6), e.pnl);
        xMoney(r.getCell(7), e.fee, { colorSign: false });
        xMoney(r.getCell(8), e.settlement);
        xMoney(r.getCell(9), e.backerBook);
        if (i % 2 === 1) for (let j = 1; j <= 9; j++) r.getCell(j).fill = fillOf(XLC.rowAlt);
      });
    }
  }

  // — Ak / Jon
  const wa = wb.addWorksheet("Ak-Jon");
  wa.getColumn(1).width = 44; wa.getColumn(2).width = 15;
  xTitle(wa, `Ak / Jon reconciliation — ${period || "this week"}`);
  const m = model;
  const line = (label, v, { bold = false, sign = true } = {}) => {
    const r = wa.addRow([]);
    xText(r.getCell(1), label, { bold, mute: !bold });
    xMoney(r.getCell(2), v, { bold, colorSign: sign });
  };
  line("Total tips collected (all accounts)", m.clubRevenue, { sign: false });
  line("Tipbacks to agents & players", -m.extTipbacks, { sign: false });
  line("Backed players' RB (cash + credits)", -m.backedRB, { sign: false });
  line("Owner accounts' 100% feeback", -m.ownFeeback, { sign: false });
  line("Club profit", m.clubProfit, { bold: true });
  m.feeRows.forEach((f) => line(`${f.label} — ${f.kind === "fixed" ? "fixed" : f.pct + "% of profit"}` + (f.recipient !== "external" ? ` → ${f.recipient === "ak" ? "Ak" : "Jon"}` : ""), -f.amount, { sign: false }));
  line("Net profit to split", m.netProfit, { bold: true });
  line("Each owner's half", m.netProfit / 2, { sign: false });
  wa.addRow([]);
  line("Ak — own accounts (P&L + 100% feeback)", m.ownPosition.ak);
  line("Ak — backed books", m.backedBook.ak);
  line("Ak — action-buy tax book", m.taxBook.ak);
  line("Ak — half of net profit", m.netProfit / 2, { sign: false });
  line("AK ENTITLEMENT", m.entitle.ak, { bold: true });
  wa.addRow([]);
  line("Jon — own accounts (P&L + 100% feeback)", m.ownPosition.jon);
  line("Jon — backed books", m.backedBook.jon);
  line("Jon — action-buy tax book", m.taxBook.jon);
  line("Jon — half of net profit", m.netProfit / 2, { sign: false });
  m.feeRows.filter((f) => f.recipient === "jon").forEach((f) => line(`Jon — ${f.label} fee`, f.amount, { sign: false }));
  line("JON ENTITLEMENT", m.entitle.jon, { bold: true });

  // ——— who settles what, per owner, plus the transfer ———
  const assignments = cfg?.assignments || {};
  [["ak", "Ak"], ["jon", "Jon"]].forEach(([w, W]) => {
    wa.addRow([]);
    const hr = wa.addRow([`${W} settles these deals`, "Settlement"]);
    hr.eachCell((c2, i) => { c2.font = { name: "Arial", size: 10, bold: true, color: { argb: XLC.white } }; c2.fill = fillOf(XLC.gold); c2.alignment = { horizontal: i === 1 ? "left" : "right" }; });
    m.entities.filter((e) => assignments[e.key] === w).forEach((e, i) => {
      const r = wa.addRow([]);
      xText(r.getCell(1), `${e.name} — ${e.settlement > 0.005 ? `${W} pays them` : e.settlement < -0.005 ? `they pay ${W}` : "even"}`, {});
      xMoney(r.getCell(2), e.settlement);
      if (i % 2 === 1) for (let j = 1; j <= 2; j++) r.getCell(j).fill = fillOf(XLC.rowAlt);
    });
    m.feeRows.filter((f) => f.recipient === "external").forEach((f) => {
      const share = f.paidBy === "split" ? f.amount / 2 : f.paidBy === w ? f.amount : 0;
      if (share > 0.005) { const r = wa.addRow([]); xText(r.getCell(1), `${f.label} (paid out)`, { mute: true }); xMoney(r.getCell(2), -share); }
    });
    const s1 = wa.addRow([]); xText(s1.getCell(1), `${W} — actual cash after settling`, { bold: true }); xMoney(s1.getCell(2), m.actual[w], { bold: true });
    for (let j = 1; j <= 2; j++) s1.getCell(j).fill = fillOf(XLC.cream);
    const s2 = wa.addRow([]); xText(s2.getCell(1), `${W} — entitlement`, {}); xMoney(s2.getCell(2), m.entitle[w]);
  });
  wa.addRow([]);
  const tr = wa.addRow([]);
  const owe = m.akOwesJon;
  xText(tr.getCell(1), Math.abs(owe) < 0.005 ? "PERFECTLY EVEN — NO TRANSFER" : owe > 0 ? "AK PAYS JON" : "JON PAYS AK", { bold: true, white: true });
  xMoney(tr.getCell(2), Math.abs(owe), { bold: true, white: true, colorSign: false });
  for (let j = 1; j <= 2; j++) tr.getCell(j).fill = fillOf(XLC.bar);

  await saveWb(wb, `FishTank_Settlements_${periodSlug(period)}.xlsx`);
}

// ———————————————— Export (clipboard — downloads are blocked in this sandbox) ————————————————
const toTSV = (header, rows) => [header, ...rows].map((r) => r.map((v) => String(v ?? "")).join("\t")).join("\n");

function ExportModal({ data, onClose }) {
  const [copied, setCopied] = useState(false);
  const taRef = useRef(null);
  if (!data) return null;
  const copy = async () => {
    try { await navigator.clipboard.writeText(data.text); setCopied(true); }
    catch (e) {
      taRef.current?.select();
      try { document.execCommand("copy"); setCopied(true); } catch (e2) {}
    }
    setTimeout(() => setCopied(false), 1800);
  };
  return (
    <div onClick={onClose} style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.55)", zIndex: 50, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: C.card, borderRadius: 12, padding: "20px 22px", width: "min(760px, 94vw)", maxHeight: "84vh", display: "flex", flexDirection: "column", boxShadow: "0 8px 40px rgba(0,0,0,0.35)" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 10 }}>
          <div style={{ fontFamily: "Georgia, serif", fontSize: 17, color: C.ink }}>{data.title}</div>
          <div style={{ marginLeft: "auto", display: "flex", gap: 8 }}>
            <Btn tone="gold" small onClick={copy}>{copied ? "✓ Copied" : "Copy to clipboard"}</Btn>
            <Btn tone="ghost" small onClick={onClose}>Close</Btn>
          </div>
        </div>
        <div style={{ color: C.mute, fontSize: 12, marginBottom: 8 }}>
          Tab-separated — paste straight into Excel or Google Sheets and it lands in columns. (File downloads are blocked inside this app's sandbox.)
        </div>
        <textarea ref={taRef} readOnly value={data.text} onFocus={(e) => e.target.select()}
          style={{ ...inputS, width: "100%", boxSizing: "border-box", flex: 1, minHeight: 260, fontFamily: "ui-monospace, Menlo, monospace", fontSize: 11.5, whiteSpace: "pre", resize: "vertical" }} />
      </div>
    </div>
  );
}

// ———————————————— UI atoms ————————————————
const th = { fontSize: 10.5, letterSpacing: "0.09em", textTransform: "uppercase", color: C.goldDark, fontWeight: 700, padding: "8px 10px", textAlign: "right", whiteSpace: "nowrap" };
const td = { padding: "7px 10px", textAlign: "right", fontVariantNumeric: "tabular-nums", fontSize: 13.5, whiteSpace: "nowrap", color: C.ink };
const tdL = { ...td, textAlign: "left" };
const inputS = { padding: "5px 8px", border: `1px solid ${C.line}`, borderRadius: 5, fontSize: 13, color: C.ink, background: C.surface };

function Pill({ children, tone = "gold" }) {
  const bg = { red: "var(--pillRedBg)", green: "var(--pillGreenBg)", blue: "var(--pillBlueBg)", gold: "var(--pillGoldBg)" }[tone];
  const fg = { red: C.red, green: C.green, blue: "var(--pillBlueFg)", gold: C.goldDark }[tone];
  return <span style={{ background: bg, color: fg, borderRadius: 4, padding: "2px 8px", fontSize: 11, fontWeight: 700, letterSpacing: "0.04em", whiteSpace: "nowrap" }}>{children}</span>;
}
const typePill = (e) =>
  e.type === "umbrella" ? <Pill tone="blue">umbrella · {e.subgroups.length} SAs</Pill>
  : e.type === "sa" ? <Pill tone="gold">SA · {e.members.length}</Pill>
  : e.type === "backed" ? <Pill tone="red">{e.dealType === "action" ? "action buy" : "house-backed"}</Pill>
  : <Pill tone="green">player</Pill>;

function PctInput({ value, onChange, width = 64 }) {
  const [v, setV] = useState(value ?? "");
  useEffect(() => setV(value ?? ""), [value]);
  return (
    <span style={{ whiteSpace: "nowrap" }}>
      <input value={v} onChange={(e) => setV(e.target.value)}
        onBlur={() => { const n = parseFloat(v); onChange(isNaN(n) ? null : Math.max(0, Math.min(100, n))); }}
        onKeyDown={(e) => e.key === "Enter" && e.currentTarget.blur()}
        style={{ ...inputS, width, textAlign: "right", fontVariantNumeric: "tabular-nums", border: `1px solid ${C.gold}` }} />
      <span style={{ marginLeft: 4, color: C.mute, fontSize: 12 }}>%</span>
    </span>
  );
}
function NumInput({ value, onChange, width = 90 }) {
  const [v, setV] = useState(value ?? "");
  useEffect(() => setV(value ?? ""), [value]);
  return (
    <input value={v} onChange={(e) => setV(e.target.value)}
      onBlur={() => { const n = parseFloat(String(v).replace(/,/g, "")); onChange(isNaN(n) ? 0 : n); }}
      onKeyDown={(e) => e.key === "Enter" && e.currentTarget.blur()}
      style={{ ...inputS, width, textAlign: "right", fontVariantNumeric: "tabular-nums", border: `1px solid ${C.gold}` }} />
  );
}
function Btn({ children, onClick, tone = "dark", small, disabled }) {
  const s = { dark: { background: "var(--bar)", color: "var(--barText)" }, gold: { background: C.gold, color: "var(--onGold)" }, ghost: { background: "transparent", color: C.goldDark, border: `1px solid ${C.gold}` } }[tone];
  return (
    <button onClick={onClick} disabled={disabled} style={{ ...s, border: s.border || "none", borderRadius: 6, padding: small ? "5px 12px" : "9px 18px", fontSize: small ? 12 : 13.5, fontWeight: 700, letterSpacing: "0.03em", cursor: disabled ? "default" : "pointer", opacity: disabled ? 0.45 : 1 }}>
      {children}
    </button>
  );
}
const Card = ({ title, children, right }) => (
  <div style={{ background: C.card, borderRadius: 10, padding: "16px 20px", boxShadow: "0 1px 5px rgba(0,0,0,0.15)" }}>
    <div style={{ display: "flex", alignItems: "center", marginBottom: 10, flexWrap: "wrap", gap: 8 }}>
      <div style={{ fontFamily: "Georgia, serif", fontSize: 16, color: C.ink }}>{title}</div>
      {right && <div style={{ marginLeft: "auto" }}>{right}</div>}
    </div>
    {children}
  </div>
);

function RoleSplash({ theme, setTheme, onChoose }) {
  const option = (role, title, copy) => (
    <button onClick={() => onChoose(role)} style={{
      minHeight: 190, padding: "28px 30px", textAlign: "left", cursor: "pointer",
      border: `1px solid ${C.line}`, borderRadius: 12, background: C.card, color: C.ink,
      boxShadow: "0 5px 22px rgba(0,0,0,0.18)", transition: "transform .12s ease, border-color .12s ease",
    }}>
      <div style={{ color: C.goldDark, fontSize: 11, fontWeight: 800, letterSpacing: ".1em", textTransform: "uppercase", marginBottom: 16 }}>Open workspace</div>
      <div style={{ fontFamily: "Georgia, 'Times New Roman', serif", fontSize: 28, marginBottom: 9 }}>{title}</div>
      <div style={{ color: C.mute, fontSize: 13.5, lineHeight: 1.55 }}>{copy}</div>
      <div style={{ color: C.goldDark, fontWeight: 800, fontSize: 13, marginTop: 22 }}>Continue →</div>
    </button>
  );
  return (
    <div style={{ ...PALETTES[theme], minHeight: "100vh", background: C.paper, color: C.ink, fontFamily: "'Avenir Next', 'Segoe UI', system-ui, sans-serif", colorScheme: theme, display: "flex", flexDirection: "column" }}>
      <div style={{ background: C.bar, padding: "18px 26px", display: "flex", alignItems: "center" }}>
        <div style={{ fontFamily: "Georgia, 'Times New Roman', serif", fontSize: 24, color: "var(--barText)" }}>Ledger Pilot</div>
        <button onClick={() => setTheme(theme === "dark" ? "light" : "dark")} title="Toggle dark mode" style={{ marginLeft: "auto", background: "transparent", border: "1px solid var(--barMute)", borderRadius: 6, color: "var(--barText)", cursor: "pointer", padding: "4px 10px", fontSize: 14 }}>{theme === "dark" ? "☀" : "☾"}</button>
      </div>
      <div style={{ width: "min(820px, calc(100% - 40px))", margin: "auto", padding: "54px 0 80px" }}>
        <div style={{ textAlign: "center", marginBottom: 30 }}>
          <div style={{ fontFamily: "Georgia, 'Times New Roman', serif", fontSize: 38, marginBottom: 9 }}>Choose your workspace</div>
          <div style={{ color: C.mute, fontSize: 14 }}>Select how you use Ledger Pilot.</div>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(280px, 1fr))", gap: 16 }}>
          {option("owner", "Club Owner", "Open the complete settlement dashboard with Fish Tank, club accounting and running tabs.")}
          {option("agent", "Agent", "Create collection tasks for selected clubs, players and reporting periods.")}
        </div>
      </div>
    </div>
  );
}

// ———————————————— Main ————————————————
export default function App() {
  const [cfg, setCfg] = useState(DEFAULT_CONFIG);
  const [portal, setPortal] = useState(null);
  const [players, setPlayers] = useState(null);
  const [period, setPeriod] = useState("");
  const [weekAdj, setWeekAdj] = useState({});
  const [tab, setTab] = useState("settle");
  const [err, setErr] = useState("");
  const [saveNote, setSaveNote] = useState("");
  const [expanded, setExpanded] = useState({});
  const [loaded, setLoaded] = useState(false);
  const fileRef = useRef(null);

  useEffect(() => {
    (async () => {
      try {
        const c = await store.get("fishtank-config-v4");
        if (c?.value) {
          const s = JSON.parse(c.value);
          if (!s.themeV2) { s.theme = "dark"; s.themeV2 = true; }
          setCfg({ ...DEFAULT_CONFIG, ...s,
            ownAccounts: s.ownAccounts || DEFAULT_CONFIG.ownAccounts,
            fees: s.fees || DEFAULT_CONFIG.fees,
            backed: s.backed || DEFAULT_CONFIG.backed,
            umbrellas: s.umbrellas || DEFAULT_CONFIG.umbrellas });
        }
      } catch (e) {}
      try {
        const d = await store.get("fishtank-lastweek-v4");
        if (d?.value) { const s = JSON.parse(d.value); setPlayers(s.players); setPeriod(s.period); setWeekAdj(s.weekAdj || {}); }
      } catch (e) {}
      setLoaded(true);
    })();
  }, []);

  const persistCfg = useCallback(async (next) => {
    setCfg(next);
    try { await store.set("fishtank-config-v4", JSON.stringify(next)); setSaveNote(""); }
    catch (e) { setSaveNote("Change couldn't be saved — it still applies this session."); }
  }, []);
  const up = (patch) => persistCfg({ ...cfg, ...patch });

  const persistWeek = async (p, per, adj) => {
    try { await store.set("fishtank-lastweek-v4", JSON.stringify({ players: p, period: per, weekAdj: adj })); } catch (e) {}
  };
  const setAdj = (adj) => { setWeekAdj(adj); persistWeek(players, period, adj); };

  const onFile = async (file) => {
    setErr("");
    try {
      const buf = await file.arrayBuffer();
      const { players: p, period: per } = parseWorkbook(buf);
      const names = { ...cfg.names };
      p.forEach((x) => { names[x.memberId] = x.name; if (x.saId !== "-") names[x.saId] = x.saName; });
      await persistCfg({ ...cfg, names });
      setPlayers(p); setPeriod(per); setWeekAdj({}); setTab("settle");
      persistWeek(p, per, {});
    } catch (e) { setErr(e.message || String(e)); }
  };

  const model = useMemo(() => (players ? buildModel(players, cfg, weekAdj, period) : null), [players, cfg, weekAdj, period]);

  const needsSetup = useMemo(() => {
    if (!model) return { deals: [], assigns: [] };
    const deals = [
      ...model.looseSAs.filter((e) => !cfg.confirmedSAs[e.id]),
      ...model.umbEntities.flatMap((u) => u.subgroups.filter((s) => !cfg.confirmedSAs[s.id])),
      ...model.indEntities.filter((e) => !cfg.confirmedPlayers[e.id]),
    ];
    return { deals, assigns: model.unassigned };
  }, [model, cfg]);

  const finalized = cfg.finalizedPeriods?.[period];
  const finalizeWeek = () => {
    const backed = { ...cfg.backed };
    const snapshot = {};
    model.backedEntities.forEach((e) => {
      if (e.dealType !== "makeup") return;
      const k = e.key.slice(2);
      snapshot[k] = Math.round(e.makeupBefore * 100) / 100;
      if (backed[k]) backed[k] = { ...backed[k], makeup: Math.round(e.makeupAfter * 100) / 100 };
    });
    up({ backed, finalizedPeriods: { ...cfg.finalizedPeriods, [period]: { snapshot } } });
  };

  if (!loaded) return <div style={{ fontFamily: "Georgia, serif", padding: 40, color: "#8A7E6C" }}>Loading saved setup…</div>;

  const theme = cfg.theme === "dark" ? "dark" : "light";

  if (!portal) return <RoleSplash theme={theme} setTheme={(next) => up({ theme: next })} onChoose={(role) => {
    if (role === "owner" && cfg.mode === "collector") up({ mode: "fishtank" });
    setPortal(role);
  }} />;

  return (
    <div style={{ ...PALETTES[theme], minHeight: "100vh", background: C.paper, color: C.ink, fontFamily: "'Avenir Next', 'Segoe UI', system-ui, sans-serif", colorScheme: theme }}>
      <div style={{ background: C.bar, padding: "18px 26px", display: "flex", alignItems: "center", gap: 18, flexWrap: "wrap" }}>
        <button type="button" onClick={() => setPortal(null)} aria-label="Return to workspace selection" title="Back to workspace selection" style={{
          border: "none", padding: 0, background: "transparent", cursor: "pointer", textAlign: "left",
          fontFamily: "Georgia, 'Times New Roman', serif", fontSize: 24, color: "var(--barText)",
        }}>
          Ledger Pilot <span style={{ color: "var(--barGold)", fontSize: 15 }}>· {portal === "agent" ? "agent data" : "weekly settlements"}</span>
        </button>
        <div style={{ display: "flex", gap: 2, background: "rgba(255,255,255,0.08)", borderRadius: 7, padding: 3 }}>
          {(portal === "agent" ? [["collector", "Data Tasks"]] : [["fishtank", "Fish Tank"], ["agent", "My Clubs"], ["tabs", "Tabs"]]).map(([k, label]) => (
            <button key={k} onClick={() => up({ mode: k })} style={{
              border: "none", cursor: "pointer", borderRadius: 5, padding: "5px 14px", fontSize: 12.5, fontWeight: 700,
              background: portal === "agent" || (cfg.mode || "fishtank") === k ? "var(--gold)" : "transparent",
              color: portal === "agent" || (cfg.mode || "fishtank") === k ? "var(--onGold)" : "var(--barMute)" }}>
              {label}
            </button>
          ))}
        </div>
        <div style={{ marginLeft: "auto", display: "flex", gap: 10, alignItems: "center" }}>
          {portal === "owner" && (cfg.mode || "fishtank") === "fishtank" && period && <span style={{ color: "var(--barMute)", fontSize: 12.5 }}>{period}{finalized ? " · finalized" : ""}</span>}
          <button onClick={() => setPortal(null)} style={{ background: "transparent", border: "none", color: "var(--barMute)", cursor: "pointer", fontSize: 12.5 }}>Choose role</button>
          <button onClick={() => up({ theme: theme === "dark" ? "light" : "dark" })} title="Toggle dark mode"
            style={{ background: "transparent", border: "1px solid var(--barMute)", borderRadius: 6, color: "var(--barText)", cursor: "pointer", padding: "4px 10px", fontSize: 14 }}>
            {theme === "dark" ? "☀" : "☾"}
          </button>
          <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }}
            onChange={(e) => { const f = e.target.files?.[0]; if (f) onFile(f); e.target.value = ""; }} />
          {portal === "owner" && (cfg.mode || "fishtank") === "fishtank" && <Btn tone="gold" small onClick={() => fileRef.current?.click()}>Upload weekly export</Btn>}
        </div>
      </div>

      {portal === "agent" ? (
        <DataTasks />
      ) : (cfg.mode || "fishtank") === "agent" ? (
        <AgentClubs theme={theme} />
      ) : (cfg.mode || "fishtank") === "tabs" ? (
        <TabsLedger />
      ) : (
      <>
      <div style={{ display: "flex", gap: 4, padding: "10px 26px 0", borderBottom: `2px solid ${C.line}`, background: C.paper, flexWrap: "wrap" }}>
        {[["settle", "Settlements"], ["agents", "Agent & umbrella reports"], ["backed", "House-backed"], ["recon", "Ak / Jon"], ["deals", "Deals & setup"]].map(([k, label]) => (
          <button key={k} onClick={() => setTab(k)} style={{
            border: "none", cursor: "pointer", padding: "9px 16px", fontSize: 13.5, fontWeight: 700,
            background: tab === k ? C.card : "transparent", color: tab === k ? C.ink : C.mute,
            borderRadius: "8px 8px 0 0", marginBottom: -2,
            boxShadow: tab === k ? "0 -1px 4px rgba(0,0,0,0.1)" : "none" }}>
            {label}
            {k === "deals" && (needsSetup.deals.length + needsSetup.assigns.length > 0) && (
              <span style={{ marginLeft: 6, background: C.red, color: "#fff", borderRadius: 9, padding: "1px 7px", fontSize: 10.5 }}>
                {needsSetup.deals.length + needsSetup.assigns.length}
              </span>
            )}
          </button>
        ))}
      </div>

      <div style={{ padding: "20px 26px 60px", maxWidth: 1180, margin: "0 auto" }}>
        {err && <div style={{ background: "var(--errBg)", color: C.red, padding: "10px 14px", borderRadius: 6, marginBottom: 14, fontSize: 13.5 }}>{err}</div>}
        {saveNote && <div style={{ background: C.banner, color: C.goldDark, padding: "8px 14px", borderRadius: 6, marginBottom: 14, fontSize: 12.5 }}>{saveNote}</div>}

        {!players && (
          <div style={{ background: C.card, border: `1px dashed ${C.gold}`, borderRadius: 10, padding: "50px 30px", textAlign: "center" }}>
            <div style={{ fontFamily: "Georgia, serif", fontSize: 20, marginBottom: 8 }}>Start the week</div>
            <div style={{ color: C.mute, fontSize: 14, marginBottom: 18 }}>Upload the club's weekly .xlsx export. Deals, umbrellas, backed-player ledgers, and fees are saved and apply automatically.</div>
            <Btn onClick={() => fileRef.current?.click()}>Choose file</Btn>
          </div>
        )}

        {players && model && tab === "settle" && <SettleTab model={model} cfg={cfg} needsSetup={needsSetup} goDeals={() => setTab("deals")} period={period} />}
        {players && model && tab === "agents" && <AgentsTab model={model} expanded={expanded} setExpanded={setExpanded} period={period} />}
        {players && model && tab === "backed" && <BackedTab model={model} cfg={cfg} up={up} finalized={finalized} finalizeWeek={finalizeWeek} />}
        {players && model && tab === "recon" && <ReconTab model={model} cfg={cfg} up={up} period={period} />}
        {players && model && tab === "deals" && <DealsTab model={model} cfg={cfg} up={up} needsSetup={needsSetup} weekAdj={weekAdj} setAdj={setAdj} />}
      </div>
      </>
      )}
    </div>
  );
}

// ———————————————— Settlements ————————————————
function SettleTab({ model, cfg, needsSetup, goDeals, period }) {
  const t = model.totals;
  const pending = needsSetup.deals.length + needsSetup.assigns.length;
  const [exportData, setExportData] = useState(null);
  const exportCsv = () =>
    setExportData({
      title: `Settlements · ${period || "this week"}`,
      text: toTSV(["Deal", "Type", "Hands", "Winnings", "Tips", "Avg TB %", "Tipback", "Settlement"],
        model.entities.map((e) => [e.name, e.type, e.hands, e.pnl.toFixed(2), e.fee.toFixed(2), e.fee ? ((e.tipback / e.fee) * 100).toFixed(1) : "", e.tipback.toFixed(2), e.settlement.toFixed(2)])),
    });

  return (
    <div>
      <ExportModal data={exportData} onClose={() => setExportData(null)} />
      {pending > 0 && (
        <div style={{ background: C.banner, border: `1px solid ${C.gold}`, borderRadius: 8, padding: "12px 16px", marginBottom: 16, display: "flex", alignItems: "center", gap: 12 }}>
          <div style={{ fontSize: 13.5 }}>
            <b>{pending} item{pending > 1 ? "s" : ""} need review this week</b> — {needsSetup.deals.length > 0 && `${needsSetup.deals.length} deal${needsSetup.deals.length > 1 ? "s" : ""} to confirm`}{needsSetup.deals.length > 0 && needsSetup.assigns.length > 0 && ", "}{needsSetup.assigns.length > 0 && `${needsSetup.assigns.length} unassigned to Ak/Jon`}. Unconfirmed deals use the default {cfg.defaultTB}%.
          </div>
          <div style={{ marginLeft: "auto" }}><Btn tone="gold" small onClick={goDeals}>Review now</Btn></div>
        </div>
      )}

      <div style={{ display: "flex", alignItems: "baseline", marginBottom: 10 }}>
        <div style={{ fontFamily: "Georgia, serif", fontSize: 19 }}>What every deal owes</div>
        <div style={{ marginLeft: "auto", display: "flex", gap: 8 }}>
          <Btn tone="gold" small onClick={() => {
            if (model.unassigned.length > 0) {
              window.alert(`Assign ${model.unassigned.length} remaining deal${model.unassigned.length > 1 ? "s" : ""} to Ak or Jon first (Ak / Jon tab). The workbook includes each owner's collection list and the final Ak↔Jon transfer, so it needs every deal assigned.`);
              return;
            }
            downloadWorkbook(model, period, cfg);
          }}>Download Excel workbook</Btn>
          <Btn tone="ghost" small onClick={exportCsv}>Copy table</Btn>
        </div>
      </div>
      <div style={{ color: C.mute, fontSize: 12.5, marginBottom: 12 }}>
        <span style={{ color: C.green, fontWeight: 700 }}>Green</span> = you pay them · <span style={{ color: C.red, fontWeight: 700 }}>red</span> = they pay you.
        Umbrellas and super agents settle as one line; no-SA players settle individually. Makeup players' settlement is their share of profit above makeup only (RB is a credit inside their net); action buys settle the player's share of P&L + RB. Owner accounts are excluded (see Ak / Jon).
      </div>

      <div style={{ background: C.card, borderRadius: 10, overflow: "hidden", boxShadow: "0 1px 6px rgba(0,0,0,0.15)" }}>
        <div style={{ background: C.bar, color: "var(--barText)", display: "grid", gridTemplateColumns: "minmax(200px,1.5fr) repeat(6, 1fr)", padding: "13px 10px", alignItems: "center" }}>
          <div style={{ paddingLeft: 10, fontFamily: "Georgia, serif", fontSize: 16 }}>Grand Total</div>
          {[fmtI(t.hands), fmt(t.pnl), fmt(t.fee), t.fee ? ((t.tipback / t.fee) * 100).toFixed(0) + "%" : "—", fmt(t.tipback)].map((v, i) => (
            <div key={i} style={{ textAlign: "right", paddingRight: 10, fontVariantNumeric: "tabular-nums", fontSize: 14.5 }}>{v}</div>
          ))}
          <div style={{ textAlign: "right", paddingRight: 10, fontVariantNumeric: "tabular-nums", fontSize: 14.5, color: t.settlement >= 0 ? "var(--barGreen)" : "var(--barRed)", fontWeight: 700 }}>{fmt(t.settlement)}</div>
        </div>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead><tr style={{ background: C.cream }}>
            <th style={{ ...th, textAlign: "left" }}>Deal name</th>
            <th style={th}>Hands</th><th style={th}>Winnings</th><th style={th}>Tips</th><th style={th}>Avg TB %</th><th style={th}>Tipback</th><th style={th}>Settlement</th>
          </tr></thead>
          <tbody>
            {model.entities.map((e, i) => (
              <tr key={e.key} style={{ background: i % 2 ? C.rowAlt : C.card, borderTop: `1px solid ${C.line}` }}>
                <td style={tdL}>
                  <span style={{ fontWeight: 600 }}>{e.name}</span> {typePill(e)}
                  {e.adjusted && <span style={{ marginLeft: 6 }}><Pill tone="blue">mid-week deal</Pill></span>}
                  {e.type === "backed" && e.dealType === "makeup" && e.inMakeup && <span style={{ marginLeft: 6 }}><Pill tone="red">in makeup</Pill></span>}
                  {e.type === "player" && e.members[0].actionTaxPct ? <span style={{ marginLeft: 6 }}><Pill tone="blue">action {e.members[0].actionTaxPct}%</Pill></span> : null}
                </td>
                <td style={td}>{fmtI(e.hands)}</td>
                <td style={td}>{fmt(e.pnl)}</td>
                <td style={td}>{fmt(e.fee)}</td>
                <td style={td}>{e.fee ? ((e.tipback / e.fee) * 100).toFixed(0) + "%" : "—"}</td>
                <td style={td}>{fmt(e.tipback)}</td>
                <td style={td}>{money(e.settlement)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ———————————————— Agent & umbrella reports ————————————————
function AgentsTab({ model, expanded, setExpanded, period }) {
  const groups = [...model.umbEntities, ...model.looseSAs];
  const [exportData, setExportData] = useState(null);
  const exportOne = (e, members) => {
    const rows = members.map((m) => [m.name, m.memberId, m.saName, m.agentName, m.hands, m.pnl.toFixed(2), m.fee.toFixed(2), m.tbPct, m.tipback.toFixed(2), m.settlement.toFixed(2)]);
    const total = ["TOTAL", "", "", "", e.hands, e.pnl.toFixed(2), e.fee.toFixed(2), e.fee ? ((e.tipback / e.fee) * 100).toFixed(1) : "", e.tipback.toFixed(2), e.settlement.toFixed(2)];
    setExportData({
      title: `${e.name} · ${period || "this week"}`,
      text: toTSV(["Player", "Device ID", "Super Agent", "Agent", "Hands", "Winnings", "Tips", "TB %", "Tipback", "Settlement"], [...rows, total]),
    });
  };

  const memberTable = (members, total) => (
    <table style={{ width: "100%", borderCollapse: "collapse" }}>
      <thead><tr>
        <th style={{ ...th, textAlign: "left" }}>Player</th><th style={{ ...th, textAlign: "left" }}>Device ID</th><th style={{ ...th, textAlign: "left" }}>Agent</th>
        <th style={th}>Hands</th><th style={th}>Winnings</th><th style={th}>Tips</th><th style={th}>TB %</th><th style={th}>Tipback</th><th style={th}>Settlement</th>
      </tr></thead>
      <tbody>
        {[...members].sort((a, b) => b.fee - a.fee).map((m, i) => (
          <tr key={m.memberId + i} style={{ background: i % 2 ? C.rowAlt : C.card, borderTop: `1px solid ${C.line}` }}>
            <td style={{ ...tdL, fontWeight: 600 }}>{m.name}{m.actionTaxPct ? <span style={{ marginLeft: 6 }}><Pill tone="blue">action {m.actionTaxPct}%</Pill></span> : null}</td>
            <td style={{ ...tdL, color: C.mute, fontSize: 12 }}>{m.memberId}</td>
            <td style={{ ...tdL, color: C.mute, fontSize: 12 }}>{m.agentName}</td>
            <td style={td}>{fmtI(m.hands)}</td><td style={td}>{fmt(m.pnl)}</td><td style={td}>{fmt(m.fee)}</td>
            <td style={td}>{m.tbPct}%</td><td style={td}>{fmt(m.tipback)}</td><td style={td}>{money(m.settlement)}</td>
          </tr>
        ))}
        <tr style={{ background: C.cream, borderTop: `2px solid ${C.gold}` }}>
          <td style={{ ...tdL, fontWeight: 700 }} colSpan={3}>Total</td>
          <td style={{ ...td, fontWeight: 700 }}>{fmtI(total.hands)}</td>
          <td style={{ ...td, fontWeight: 700 }}>{fmt(total.pnl)}</td>
          <td style={{ ...td, fontWeight: 700 }}>{fmt(total.fee)}</td>
          <td style={td}>{total.fee ? ((total.tipback / total.fee) * 100).toFixed(0) + "%" : "—"}</td>
          <td style={{ ...td, fontWeight: 700 }}>{fmt(total.tipback)}</td>
          <td style={{ ...td, fontWeight: 700 }}>{money(total.settlement)}</td>
        </tr>
      </tbody>
    </table>
  );

  return (
    <div>
      <ExportModal data={exportData} onClose={() => setExportData(null)} />
      <div style={{ fontFamily: "Georgia, serif", fontSize: 19, marginBottom: 4 }}>Reports for each deal</div>
      <div style={{ color: C.mute, fontSize: 12.5, marginBottom: 14 }}>Click a row to expand; export a CSV to send them. Umbrella reports break out each super agent inside.</div>
      {groups.map((e) => {
        const open = expanded[e.key];
        return (
          <div key={e.key} style={{ background: C.card, borderRadius: 10, marginBottom: 10, overflow: "hidden", boxShadow: "0 1px 5px rgba(0,0,0,0.15)" }}>
            <div onClick={() => setExpanded({ ...expanded, [e.key]: !open })} style={{ display: "flex", alignItems: "center", gap: 14, padding: "12px 16px", cursor: "pointer", background: C.cream }}>
              <span style={{ color: C.goldDark, fontSize: 12, width: 12 }}>{open ? "▼" : "►"}</span>
              <span style={{ fontWeight: 700, fontSize: 14.5 }}>{e.name}</span>
              {typePill(e)}
              <span style={{ color: C.mute, fontSize: 12 }}>{e.members.length} player{e.members.length !== 1 ? "s" : ""} · {fmtI(e.hands)} hands</span>
              <span style={{ marginLeft: "auto", fontSize: 13.5 }}>settlement {money(e.settlement)}</span>
              <Btn tone="gold" small onClick={(ev) => { ev.stopPropagation(); downloadDealExcel(e, period); }}>Excel</Btn>
              <Btn tone="ghost" small onClick={(ev) => { ev.stopPropagation(); exportOne(e, e.members); }}>Copy</Btn>
            </div>
            {open && (e.type === "umbrella"
              ? e.subgroups.map((s) => (
                  <div key={s.key} style={{ borderTop: `1px solid ${C.line}` }}>
                    <div style={{ padding: "8px 16px", fontSize: 13, fontWeight: 700, color: C.goldDark, background: C.rowAlt }}>
                      {s.name} — settlement {money(s.settlement)}
                    </div>
                    {memberTable(s.members, s)}
                  </div>
                ))
              : memberTable(e.members, e))}
          </div>
        );
      })}
    </div>
  );
}

// ———————————————— House-backed tab ————————————————
function BackedTab({ model, cfg, up, finalized, finalizeWeek }) {
  const [newName, setNewName] = useState("");
  const [newDeal, setNewDeal] = useState("makeup");
  const addBacked = () => {
    const k = newName.trim().toLowerCase();
    if (!k) return;
    const base = newDeal === "action"
      ? { name: newName.trim(), deal: "action", actionPct: 50, rbPct: 100, backer: "jon" }
      : { name: newName.trim(), deal: "makeup", rbNormal: cfg.defaultTB, rbMakeup: 100, makeup: 0, playerProfitPct: 50, backer: "split" };
    up({ backed: { ...cfg.backed, [k]: base } });
    setNewName("");
  };
  const setB = (k, patch) => up({ backed: { ...cfg.backed, [k]: { ...cfg.backed[k], ...patch } } });
  const removeB = (k) => { const b = { ...cfg.backed }; delete b[k]; up({ backed: b }); };
  const findE = (k) => model.backedEntities.find((x) => x.key === `b:${k}`);
  const backerSel = (k, b, withSplit) => (
    <select value={b.backer} onChange={(e) => setB(k, { backer: e.target.value })} style={{ ...inputS, padding: "4px 6px", fontSize: 12 }}>
      {withSplit && <option value="split">Ak & Jon 50/50</option>}
      <option value="ak">Ak</option><option value="jon">Jon</option>
    </select>
  );

  const makeupPlayers = Object.entries(cfg.backed).filter(([, b]) => b.deal !== "action");
  const actionPlayers = Object.entries(cfg.backed).filter(([, b]) => b.deal === "action");

  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", marginBottom: 4, gap: 10, flexWrap: "wrap" }}>
        <div style={{ fontFamily: "Georgia, serif", fontSize: 19 }}>House-backed players</div>
        <div style={{ marginLeft: "auto" }}>
          <Btn tone={finalized ? "ghost" : "gold"} small disabled={finalized} onClick={finalizeWeek}>
            {finalized ? "Week finalized · makeup rolled" : "Finalize week — roll makeup forward"}
          </Btn>
        </div>
      </div>
      <div style={{ color: C.mute, fontSize: 12.5, marginBottom: 16 }}>
        <b>Makeup deals</b>: the week's net = P&L + RB credit. Above makeup, the player is paid their % of the excess and the backer books the rest; below, no cash moves and the net accrues to makeup on the backer's book. The margin on their fees stays in split club profit. RB rate follows makeup <b>entering</b> the week; Finalize once to roll it forward (locked per period).
        {" "}<b>Action buys</b>: the backer owns their % of (P&L + rakeback); the player settles the remainder.
      </div>

      <div style={{ fontSize: 13, fontWeight: 700, color: C.goldDark, textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 6 }}>Makeup deals</div>
      <div style={{ background: C.card, borderRadius: 10, overflow: "auto", boxShadow: "0 1px 6px rgba(0,0,0,0.15)", marginBottom: 20 }}>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead><tr style={{ background: C.cream }}>
            <th style={{ ...th, textAlign: "left" }}>Player</th><th style={{ ...th, textAlign: "left" }}>Backer</th>
            <th style={th}>Makeup entering</th><th style={{ ...th, textAlign: "center" }}>Status</th>
            <th style={th}>RB % normal</th><th style={th}>RB % makeup</th><th style={th}>Player profit %</th>
            <th style={th}>Tips</th><th style={th}>P&L</th><th style={th}>RB credit</th><th style={th}>Net</th><th style={th}>Settlement</th><th style={th}>Makeup after</th><th style={th}></th>
          </tr></thead>
          <tbody>
            {makeupPlayers.map(([k, b], i) => {
              const e = findE(k);
              return (
                <tr key={k} style={{ background: i % 2 ? C.rowAlt : C.card, borderTop: `1px solid ${C.line}` }}>
                  <td style={{ ...tdL, fontWeight: 600 }}>
                    {b.name}
                    {!e && <span style={{ marginLeft: 8 }}><Pill tone="gold">no play</Pill></span>}
                  </td>
                  <td style={tdL}>{backerSel(k, b, true)}</td>
                  <td style={td}><NumInput width={85} value={b.makeup} onChange={(v) => setB(k, { makeup: v })} /></td>
                  <td style={{ ...td, textAlign: "center" }}>{(b.makeup || 0) > 0.005 ? <Pill tone="red">in makeup</Pill> : <Pill tone="green">clear</Pill>}</td>
                  <td style={td}>
                    <span style={{ opacity: (b.makeup || 0) > 0.005 ? 0.35 : 1 }}>
                      <PctInput width={46} value={b.rbNormal} onChange={(v) => v != null && setB(k, { rbNormal: v })} />
                    </span>
                    {(b.makeup || 0) <= 0.005 && <div style={{ fontSize: 9.5, color: C.green, fontWeight: 700, marginTop: 2 }}>ACTIVE</div>}
                  </td>
                  <td style={td}>
                    <span style={{ opacity: (b.makeup || 0) > 0.005 ? 1 : 0.35 }}>
                      <PctInput width={46} value={b.rbMakeup} onChange={(v) => v != null && setB(k, { rbMakeup: v })} />
                    </span>
                    {(b.makeup || 0) > 0.005 && <div style={{ fontSize: 9.5, color: C.green, fontWeight: 700, marginTop: 2 }}>ACTIVE</div>}
                  </td>
                  <td style={td}><PctInput width={46} value={b.playerProfitPct ?? 50} onChange={(v) => v != null && setB(k, { playerProfitPct: v })} /></td>
                  <td style={td}>{e ? fmt(e.fee) : "—"}</td>
                  <td style={td}>{e ? money(e.pnl) : "—"}</td>
                  <td style={td}>{e ? <>{fmt(e.tipback)} <span style={{ color: C.mute, fontSize: 11 }}>@{e.rb}%</span></> : "—"}</td>
                  <td style={td}>{e ? money(e.net) : "—"}</td>
                  <td style={td}>{e ? <b>{fmt(e.settlement)}</b> : "—"}</td>
                  <td style={td}>{e ? fmt(e.makeupAfter) : fmt(b.makeup || 0)}</td>
                  <td style={td}><button onClick={() => removeB(k)} style={{ border: "none", background: "none", color: C.red, cursor: "pointer", fontSize: 15 }}>×</button></td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      <div style={{ fontSize: 13, fontWeight: 700, color: C.goldDark, textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 6 }}>Action buys</div>
      <div style={{ background: C.card, borderRadius: 10, overflow: "auto", boxShadow: "0 1px 6px rgba(0,0,0,0.15)", marginBottom: 20 }}>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead><tr style={{ background: C.cream }}>
            <th style={{ ...th, textAlign: "left" }}>Player</th><th style={{ ...th, textAlign: "left" }}>Backer</th>
            <th style={th}>Backer action %</th><th style={th}>RB %</th>
            <th style={th}>Tips</th><th style={th}>Week P&L</th><th style={th}>Player settlement</th><th style={th}>Backer book</th><th style={th}></th>
          </tr></thead>
          <tbody>
            {actionPlayers.map(([k, b], i) => {
              const e = findE(k);
              return (
                <tr key={k} style={{ background: i % 2 ? C.rowAlt : C.card, borderTop: `1px solid ${C.line}` }}>
                  <td style={{ ...tdL, fontWeight: 600 }}>
                    {b.name}
                    {!e && <span style={{ marginLeft: 8 }}><Pill tone="gold">no play</Pill></span>}
                  </td>
                  <td style={tdL}>{backerSel(k, b, false)}</td>
                  <td style={td}><PctInput width={46} value={b.actionPct} onChange={(v) => v != null && setB(k, { actionPct: v })} /></td>
                  <td style={td}><PctInput width={46} value={b.rbPct ?? 100} onChange={(v) => v != null && setB(k, { rbPct: v })} /></td>
                  <td style={td}>{e ? fmt(e.fee) : "—"}</td>
                  <td style={td}>{e ? money(e.pnl) : "—"}</td>
                  <td style={td}>{e ? <b>{fmt(e.settlement)}</b> : "—"}</td>
                  <td style={td}>{e ? money(e.backerBook) : "—"}</td>
                  <td style={td}><button onClick={() => removeB(k)} style={{ border: "none", background: "none", color: C.red, cursor: "pointer", fontSize: 15 }}>×</button></td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
        <input placeholder="Add backed player by exact nickname…" value={newName} onChange={(e) => setNewName(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && addBacked()} style={{ ...inputS, width: 260 }} />
        <select value={newDeal} onChange={(e) => setNewDeal(e.target.value)} style={{ ...inputS, fontSize: 12 }}>
          <option value="makeup">Makeup deal</option><option value="action">Action buy</option>
        </select>
        <Btn tone="ghost" small onClick={addBacked}>+ Add</Btn>
        <div style={{ marginLeft: "auto", fontSize: 13 }}>
          Backed books this week — Ak: <b>{money(model.backedBook.ak)}</b> · Jon: <b>{money(model.backedBook.jon)}</b>
        </div>
      </div>
    </div>
  );
}

// ———————————————— Ak / Jon reconciliation ————————————————
function ReconTab({ model, cfg, up, period }) {
  const m = model;
  const assign = (key, who) => up({ assignments: { ...cfg.assignments, [key]: who } });
  const assignAll = (who) => { const a = { ...cfg.assignments }; m.unassigned.forEach((e) => (a[e.key] = who)); up({ assignments: a }); };
  const owePos = m.akOwesJon > 0.005, oweNeg = m.akOwesJon < -0.005;

  const row = (label, val, opts = {}) => (
    <div style={{ display: "flex", padding: "6px 0", borderBottom: opts.rule ? `1px solid ${C.line}` : "none", fontSize: 13.5 }}>
      <span style={{ color: opts.bold ? C.ink : C.mute, fontWeight: opts.bold ? 700 : 400 }}>{label}</span>
      <span style={{ marginLeft: "auto", fontVariantNumeric: "tabular-nums", fontWeight: opts.bold ? 700 : 500 }}>{typeof val === "number" ? money(val) : val}</span>
    </div>
  );

  return (
    <div>
      <div style={{ background: C.bar, borderRadius: 12, padding: "26px 30px", marginBottom: 18, textAlign: "center", color: "var(--barText)" }}>
        <div style={{ fontSize: 12, letterSpacing: "0.12em", textTransform: "uppercase", color: "var(--barGold)", marginBottom: 6 }}>Owner balance · {period || "this week"}</div>
        {m.unassigned.length > 0 ? (
          <div style={{ fontSize: 17 }}>Assign the {m.unassigned.length} remaining deal{m.unassigned.length > 1 ? "s" : ""} below to get the final number.</div>
        ) : (
          <div style={{ fontFamily: "Georgia, serif", fontSize: 30 }}>
            {owePos && <>Ak pays Jon <span style={{ color: "var(--barGold)" }}>{fmt(m.akOwesJon)}</span></>}
            {oweNeg && <>Jon pays Ak <span style={{ color: "var(--barGold)" }}>{fmt(-m.akOwesJon)}</span></>}
            {!owePos && !oweNeg && <>Perfectly even — no transfer needed</>}
          </div>
        )}
        <div style={{ fontSize: 12, color: "var(--barSubtle)", marginTop: 8 }}>
          After this transfer each of you nets exactly: half of net club profit + your own accounts' P&L with 100% feeback + your backed books{m.feeRows.some((f) => f.recipient !== "external") ? " + any fee routed to you" : ""}.
          {m.balanceOk ? " Balance check: ✓ books tie out." : m.unassigned.length === 0 ? " ⚠ Balance check failed — review setup." : ""}
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(310px, 1fr))", gap: 14, marginBottom: 18 }}>
        <Card title="Club economics">
          {row("Total tips collected (all accounts)", m.clubRevenue)}
          {row("Tipbacks to agents & players", -m.extTipbacks)}
          {row("Backed players' RB (cash + credits)", -m.backedRB)}
          {row("Owner accounts' 100% feeback", -m.ownFeeback, { rule: true })}
          {row("Club profit", m.clubProfit, { bold: true })}
          {m.feeRows.map((f) => row(`${f.label} · ${f.kind === "fixed" ? "fixed" : `${f.pct}% of ${cfg.feeBase === "gross" ? "tips" : "profit"}`}${f.recipient !== "external" ? ` → ${f.recipient === "ak" ? "Ak" : "Jon"}` : ""}`, -f.amount))}
          <div style={{ borderTop: `1px solid ${C.line}` }} />
          {row("Net profit to split", m.netProfit, { bold: true })}
          {row("Each owner's half", m.netProfit / 2)}
        </Card>
        {["ak", "jon"].map((w) => (
          <Card key={w} title={`${w === "ak" ? "Ak" : "Jon"}'s position`}>
            {m.own.filter((p) => p.owner === w).map((p) => row(`${p.name} · P&L ${fmt(p.pnl)} + feeback ${fmt(p.feeback)}`, p.position))}
            {m.own.filter((p) => p.owner === w).length === 0 && <div style={{ color: C.mute, fontSize: 12.5 }}>No activity from these accounts this week.</div>}
            <div style={{ borderTop: `1px solid ${C.line}` }} />
            {row("Own accounts (P&L + 100% feeback)", m.ownPosition[w], { bold: true })}
            {row("Backed books", m.backedBook[w])}
            {row("Action-buy tax book", m.taxBook[w])}
            {row("Half of net club profit", m.netProfit / 2)}
            {row("Entitlement (all-in)", m.entitle[w], { bold: true })}
            {row("Actual cash from assigned settlements", m.actual[w])}
          </Card>
        ))}
      </div>

      <Card title="Who settles with whom" right={m.unassigned.length > 0 && (
        <span style={{ display: "flex", gap: 8 }}>
          <Btn tone="ghost" small onClick={() => assignAll("ak")}>Rest → Ak</Btn>
          <Btn tone="ghost" small onClick={() => assignAll("jon")}>Rest → Jon</Btn>
        </span>
      )}>
        <div style={{ color: C.mute, fontSize: 12.5, marginBottom: 10 }}>Mark who physically settles each deal. Remembered for future weeks.</div>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead><tr style={{ background: C.cream }}>
            <th style={{ ...th, textAlign: "left" }}>Deal</th><th style={th}>Settlement</th><th style={{ ...th, textAlign: "center" }}>Settled by</th>
          </tr></thead>
          <tbody>
            {m.entities.map((e, i) => {
              const who = cfg.assignments[e.key];
              return (
                <tr key={e.key} style={{ background: i % 2 ? C.rowAlt : C.card, borderTop: `1px solid ${C.line}` }}>
                  <td style={tdL}><b>{e.name}</b> {typePill(e)}</td>
                  <td style={td}>{money(e.settlement)}</td>
                  <td style={{ ...td, textAlign: "center" }}>
                    {["ak", "jon"].map((w) => (
                      <button key={w} onClick={() => assign(e.key, w)} style={{
                        margin: "0 3px", padding: "4px 14px", borderRadius: 5, cursor: "pointer", fontSize: 12, fontWeight: 700,
                        border: `1px solid ${who === w ? C.goldDark : C.line}`,
                        background: who === w ? C.gold : C.surface, color: who === w ? "var(--onGold)" : C.mute }}>
                        {w === "ak" ? "Ak" : "Jon"}
                      </button>
                    ))}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </Card>
    </div>
  );
}

// ———————————————— Deals & setup ————————————————
function DealsTab({ model, cfg, up, needsSetup, weekAdj, setAdj }) {
  const [showAllSA, setShowAllSA] = useState({});
  const [umbName, setUmbName] = useState("");
  const [adjTarget, setAdjTarget] = useState("");
  const [taxTarget, setTaxTarget] = useState("");
  const taxTargets = [...new Map([...model.saEntities, ...model.indEntities].flatMap((e) => e.members).map((m) => [m.memberId, m])).values()].sort((a, b) => b.fee - a.fee);
  const addTax = () => {
    if (!taxTarget) return;
    up({ actionTax: { ...(cfg.actionTax || {}), [taxTarget]: { pct: 20, backer: "split" } } });
    setTaxTarget("");
  };

  const allSAs = [...model.looseSAs, ...model.umbEntities.flatMap((u) => u.subgroups)].sort((a, b) => b.fee - a.fee);
  const inds = model.indEntities;
  const omit = (o, k) => { const x = { ...o }; delete x[k]; return x; };

  const setSA = (id, pct) => up({ saDeals: pct == null ? omit(cfg.saDeals, id) : { ...cfg.saDeals, [id]: pct }, confirmedSAs: { ...cfg.confirmedSAs, [id]: true } });
  const setPlayer = (id, pct, isInd) => {
    const patch = { playerDeals: pct == null ? omit(cfg.playerDeals, id) : { ...cfg.playerDeals, [id]: pct } };
    if (isInd) patch.confirmedPlayers = { ...cfg.confirmedPlayers, [id]: true };
    up(patch);
  };
  const confirmAll = () => {
    const cs = { ...cfg.confirmedSAs }, cp = { ...cfg.confirmedPlayers };
    allSAs.forEach((e) => (cs[e.id] = true)); inds.forEach((e) => (cp[e.id] = true));
    up({ confirmedSAs: cs, confirmedPlayers: cp });
  };

  const addUmbrella = () => { if (!umbName.trim()) return; up({ umbrellas: [...cfg.umbrellas, { id: "u" + Date.now(), name: umbName.trim(), saIds: [] }] }); setUmbName(""); };
  const toggleSAinUmb = (uid, saId) => {
    up({ umbrellas: cfg.umbrellas.map((u) => {
      if (u.id === uid) return { ...u, saIds: u.saIds.includes(saId) ? u.saIds.filter((x) => x !== saId) : [...u.saIds, saId] };
      return { ...u, saIds: u.saIds.filter((x) => x !== saId) };
    }) });
  };
  const saOptions = allSAs.map((e) => ({ id: e.id, name: e.name }));
  Object.entries(cfg.names).forEach(([id, name]) => {
    if (cfg.saDeals[id] !== undefined && !saOptions.find((o) => o.id === id)) saOptions.push({ id, name });
  });

  const adjTargets = [
    ...allSAs.map((e) => ({ key: e.key, label: `${e.name} (SA)`, fee: e.fee })),
    ...inds.map((e) => ({ key: e.key, label: `${e.name} (player)`, fee: e.fee })),
    ...model.backedEntities.map((e) => ({ key: e.key, label: `${e.name} (backed)`, fee: e.fee })),
  ];
  const addAdj = () => {
    if (!adjTarget) return;
    setAdj({ ...weekAdj, [adjTarget]: { amtA: 0, rateA: cfg.defaultTB, rateB: cfg.defaultTB } });
    setAdjTarget("");
  };
  const setAdjField = (key, patch) => setAdj({ ...weekAdj, [key]: { ...weekAdj[key], ...patch } });
  const removeAdj = (key) => { const a = { ...weekAdj }; delete a[key]; setAdj(a); };

  const newBadge = (isNew) => isNew && <Pill tone="red">confirm</Pill>;

  return (
    <div>
      {needsSetup.deals.length > 0 && (
        <div style={{ background: C.banner, border: `1px solid ${C.gold}`, borderRadius: 8, padding: "12px 16px", marginBottom: 16, display: "flex", alignItems: "center" }}>
          <div style={{ fontSize: 13.5 }}><b>New this week:</b> {needsSetup.deals.map((e) => e.name).join(", ")} — set or confirm their deals below.</div>
          <div style={{ marginLeft: "auto" }}><Btn tone="gold" small onClick={confirmAll}>Everything's right — confirm all</Btn></div>
        </div>
      )}

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(330px, 1fr))", gap: 14, marginBottom: 16 }}>
        <Card title="Defaults & fees">
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14, fontSize: 13.5 }}>
            Default tipback for anyone without a deal
            <span style={{ marginLeft: "auto" }}><PctInput value={cfg.defaultTB} onChange={(v) => v != null && up({ defaultTB: v })} /></span>
          </div>
          <div style={{ fontSize: 12, color: C.mute, marginBottom: 6, textTransform: "uppercase", letterSpacing: "0.08em" }}>Cut fees</div>
          {cfg.fees.map((f, i) => (
            <div key={f.id} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8, fontSize: 13 }}>
              <input value={f.label} onChange={(e) => { const fees = [...cfg.fees]; fees[i] = { ...f, label: e.target.value }; up({ fees }); }} style={{ ...inputS, flex: 1 }} />
              <select value={f.kind === "fixed" ? "fixed" : "pct"} onChange={(e) => { const fees = [...cfg.fees]; fees[i] = { ...f, kind: e.target.value === "fixed" ? "fixed" : "pct" }; up({ fees }); }} style={{ ...inputS, padding: "5px 6px", fontSize: 12 }}>
                <option value="pct">%</option><option value="fixed">$</option>
              </select>
              {f.kind === "fixed"
                ? <NumInput width={72} value={f.amount ?? 0} onChange={(v) => { const fees = [...cfg.fees]; fees[i] = { ...f, amount: v }; up({ fees }); }} />
                : <PctInput width={52} value={f.pct} onChange={(v) => { if (v == null) return; const fees = [...cfg.fees]; fees[i] = { ...f, pct: v }; up({ fees }); }} />}
              <select value={f.recipient} onChange={(e) => { const fees = [...cfg.fees]; fees[i] = { ...f, recipient: e.target.value }; up({ fees }); }} style={{ ...inputS, padding: "5px 6px", fontSize: 12 }}>
                <option value="external">→ outside</option><option value="ak">→ Ak</option><option value="jon">→ Jon</option>
              </select>
              <button onClick={() => up({ fees: cfg.fees.filter((_, j) => j !== i) })} style={{ border: "none", background: "none", color: C.red, cursor: "pointer", fontSize: 15 }}>×</button>
            </div>
          ))}
          <div style={{ display: "flex", gap: 10, alignItems: "center", marginTop: 8 }}>
            <Btn tone="ghost" small onClick={() => up({ fees: [...cfg.fees, { id: "f" + Date.now(), label: "New fee", pct: 1, recipient: "external", paidBy: "split" }] })}>+ Add fee</Btn>
            <label style={{ marginLeft: "auto", fontSize: 12, color: C.mute }}>
              Fees are % of{" "}
              <select value={cfg.feeBase} onChange={(e) => up({ feeBase: e.target.value })} style={{ ...inputS, padding: "3px 6px", fontSize: 12 }}>
                <option value="net">club profit (after tipbacks)</option>
                <option value="gross">gross tips collected</option>
              </select>
            </label>
          </div>
        </Card>

        <Card title="Owner accounts">
          <div style={{ fontSize: 12, color: C.mute, marginBottom: 10 }}>House play: 100% feeback, excluded from settlements, P&L + feeback credited to the owner. One nickname per line. Backed players are managed on the House-backed tab.</div>
          {["ak", "jon"].map((w) => (
            <div key={w} style={{ marginBottom: 10 }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: C.goldDark, textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 4 }}>{w === "ak" ? "Ak" : "Jon"}</div>
              <textarea defaultValue={cfg.ownAccounts[w].join("\n")}
                onBlur={(e) => up({ ownAccounts: { ...cfg.ownAccounts, [w]: e.target.value.split("\n").map((s) => s.trim()).filter(Boolean) } })}
                rows={2}
                style={{ ...inputS, width: "100%", boxSizing: "border-box", fontFamily: "inherit", resize: "vertical" }} />
            </div>
          ))}
        </Card>
      </div>

      <div style={{ marginBottom: 16 }}>
        <Card title="Umbrella groups" right={
          <span style={{ display: "flex", gap: 8 }}>
            <input placeholder="New umbrella name…" value={umbName} onChange={(e) => setUmbName(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addUmbrella()} style={{ ...inputS, width: 180 }} />
            <Btn tone="ghost" small onClick={addUmbrella}>+ Create</Btn>
          </span>
        }>
          <div style={{ color: C.mute, fontSize: 12.5, marginBottom: 10 }}>Group super agents under one name — they settle as a single line and get one combined report. Each SA keeps its own tipback rate.</div>
          {cfg.umbrellas.length === 0 && <div style={{ color: C.mute, fontSize: 13 }}>No umbrellas yet.</div>}
          {cfg.umbrellas.map((u) => (
            <div key={u.id} style={{ borderTop: `1px solid ${C.line}`, padding: "10px 0" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}>
                <input value={u.name} onChange={(e) => up({ umbrellas: cfg.umbrellas.map((x) => x.id === u.id ? { ...x, name: e.target.value } : x) })} style={{ ...inputS, fontWeight: 700, width: 200 }} />
                <Pill tone="blue">{u.saIds.length} SAs</Pill>
                <button onClick={() => up({ umbrellas: cfg.umbrellas.filter((x) => x.id !== u.id) })} style={{ marginLeft: "auto", border: "none", background: "none", color: C.red, cursor: "pointer", fontSize: 15 }}>× delete</button>
              </div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                {saOptions.map((o) => {
                  const inThis = u.saIds.includes(o.id);
                  const inOther = !inThis && cfg.umbrellas.some((x) => x.id !== u.id && x.saIds.includes(o.id));
                  return (
                    <button key={o.id} onClick={() => !inOther && toggleSAinUmb(u.id, o.id)} style={{
                      padding: "4px 12px", borderRadius: 14, fontSize: 12, fontWeight: 600, cursor: inOther ? "default" : "pointer",
                      border: `1px solid ${inThis ? C.goldDark : C.line}`,
                      background: inThis ? C.gold : C.surface, color: inThis ? "var(--onGold)" : inOther ? "var(--chipOff)" : C.mute, opacity: inOther ? 0.6 : 1 }}>
                      {o.name}
                    </button>
                  );
                })}
              </div>
            </div>
          ))}
        </Card>
      </div>

      <div style={{ marginBottom: 16 }}>
        <Card title="Mid-week deal changes (this week only)" right={
          <span style={{ display: "flex", gap: 8 }}>
            <select value={adjTarget} onChange={(e) => setAdjTarget(e.target.value)} style={{ ...inputS, fontSize: 12, maxWidth: 220 }}>
              <option value="">Pick a deal…</option>
              {adjTargets.filter((t) => !weekAdj[t.key]).map((t) => <option key={t.key} value={t.key}>{t.label}</option>)}
            </select>
            <Btn tone="ghost" small onClick={addAdj}>+ Add change</Btn>
          </span>
        }>
          <div style={{ color: C.mute, fontSize: 12.5, marginBottom: 10 }}>
            For the rare case a deal changes partway through the week: the first $X of their tips settles at the old rate, the rest at the new rate. These clear automatically when you upload the next week's file.
          </div>
          {Object.keys(weekAdj).length === 0 && <div style={{ color: C.mute, fontSize: 13 }}>None this week.</div>}
          {Object.entries(weekAdj).map(([key, a]) => {
            const t = adjTargets.find((x) => x.key === key);
            const tipback = Math.min(a.amtA, t?.fee || 0) * a.rateA / 100 + Math.max(0, (t?.fee || 0) - a.amtA) * a.rateB / 100;
            return (
              <div key={key} style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap", borderTop: `1px solid ${C.line}`, padding: "10px 0", fontSize: 13 }}>
                <b style={{ minWidth: 140 }}>{t?.label || key}</b>
                <span style={{ color: C.mute }}>first</span>
                <NumInput width={90} value={a.amtA} onChange={(v) => setAdjField(key, { amtA: v })} />
                <span style={{ color: C.mute }}>of tips @</span>
                <PctInput width={50} value={a.rateA} onChange={(v) => v != null && setAdjField(key, { rateA: v })} />
                <span style={{ color: C.mute }}>· remaining {t ? fmt(Math.max(0, t.fee - a.amtA)) : "—"} @</span>
                <PctInput width={50} value={a.rateB} onChange={(v) => v != null && setAdjField(key, { rateB: v })} />
                <span style={{ marginLeft: "auto", fontWeight: 700 }}>tipback → {fmt(tipback)}</span>
                <button onClick={() => removeAdj(key)} style={{ border: "none", background: "none", color: C.red, cursor: "pointer", fontSize: 15 }}>×</button>
              </div>
            );
          })}
        </Card>
      </div>

      <div style={{ marginBottom: 16 }}>
        <Card title="Action buys on players (tax wins / rebate losses)" right={
          <span style={{ display: "flex", gap: 8 }}>
            <select value={taxTarget} onChange={(e) => setTaxTarget(e.target.value)} style={{ ...inputS, fontSize: 12, maxWidth: 220 }}>
              <option value="">Pick a player…</option>
              {taxTargets.filter((t) => !(cfg.actionTax || {})[t.memberId]).map((t) => <option key={t.memberId} value={t.memberId}>{t.name}{t.saName !== "-" ? ` (${t.saName})` : ""}</option>)}
            </select>
            <Btn tone="ghost" small onClick={addTax}>+ Add</Btn>
          </span>
        }>
          <div style={{ color: C.mute, fontSize: 12.5, marginBottom: 10 }}>
            The set % applies to the player's net after rakeback (P&L + tipback): the house takes that % when the net is positive and gives back the same % when it's negative. The house's cut lands on the chosen book in Ak / Jon.
          </div>
          {Object.keys(cfg.actionTax || {}).length === 0 && <div style={{ color: C.mute, fontSize: 13 }}>None yet.</div>}
          {Object.entries(cfg.actionTax || {}).map(([mid, a]) => {
            const t = taxTargets.find((x) => x.memberId === mid);
            const nm = t?.name || cfg.names[mid] || mid;
            const cut = t ? ((t.pnl + t.tipback) * a.pct) / 100 : null;
            return (
              <div key={mid} style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap", borderTop: `1px solid ${C.line}`, padding: "9px 0", fontSize: 13 }}>
                <b style={{ minWidth: 130 }}>{nm}</b>
                {!t && <Pill tone="gold">no play this week</Pill>}
                <span style={{ color: C.mute }}>house %</span>
                <PctInput width={50} value={a.pct} onChange={(v) => v != null && up({ actionTax: { ...cfg.actionTax, [mid]: { ...a, pct: v } } })} />
                <span style={{ color: C.mute }}>book</span>
                <select value={a.backer || "split"} onChange={(e) => up({ actionTax: { ...cfg.actionTax, [mid]: { ...a, backer: e.target.value } } })} style={{ ...inputS, padding: "4px 6px", fontSize: 12 }}>
                  <option value="split">Ak & Jon 50/50</option><option value="ak">Ak</option><option value="jon">Jon</option>
                </select>
                {t && <span style={{ marginLeft: "auto" }}>net {money(t.pnl + t.tipback)} → house cut {money(cut)}</span>}
                <button onClick={() => { const at = { ...cfg.actionTax }; delete at[mid]; up({ actionTax: at }); }} style={{ border: "none", background: "none", color: C.red, cursor: "pointer", fontSize: 15 }}>×</button>
              </div>
            );
          })}
        </Card>
      </div>

      <div style={{ background: C.card, borderRadius: 10, padding: "16px 20px", boxShadow: "0 1px 5px rgba(0,0,0,0.15)", marginBottom: 16 }}>
        <div style={{ fontFamily: "Georgia, serif", fontSize: 16, marginBottom: 4 }}>Super agent deals</div>
        <div style={{ color: C.mute, fontSize: 12.5, marginBottom: 12 }}>The rate applies to every player under the super agent. Expand to override a specific player.</div>
        {allSAs.map((e) => {
          const isNew = !cfg.confirmedSAs[e.id];
          const open = showAllSA[e.id];
          const umb = cfg.umbrellas.find((u) => u.saIds.includes(e.id));
          return (
            <div key={e.id} style={{ borderTop: `1px solid ${C.line}`, padding: "10px 0" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <span style={{ fontWeight: 700, fontSize: 14 }}>{e.name}</span>
                {umb && <Pill tone="blue">{umb.name}</Pill>}
                <span style={{ color: C.mute, fontSize: 12 }}>{e.members.length} player{e.members.length !== 1 ? "s" : ""} · tips {fmt(e.fee)}</span>
                {newBadge(isNew)}
                <span style={{ marginLeft: "auto" }}>
                  <PctInput value={cfg.saDeals[e.id] ?? cfg.defaultTB} onChange={(v) => setSA(e.id, v ?? cfg.defaultTB)} />
                </span>
                <button onClick={() => setShowAllSA({ ...showAllSA, [e.id]: !open })} style={{ border: "none", background: "none", color: C.goldDark, cursor: "pointer", fontSize: 12.5, fontWeight: 700 }}>
                  {open ? "hide players" : "player overrides"}
                </button>
              </div>
              {open && (
                <div style={{ marginTop: 8, marginLeft: 10, display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(260px, 1fr))", gap: 6 }}>
                  {e.members.map((mm) => (
                    <div key={mm.memberId} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12.5, background: C.rowAlt, borderRadius: 6, padding: "5px 10px" }}>
                      <span>{mm.name}</span>
                      <span style={{ marginLeft: "auto" }}>
                        <PctInput width={50} value={cfg.playerDeals[mm.memberId] ?? ""} onChange={(v) => setPlayer(mm.memberId, v, false)} />
                      </span>
                    </div>
                  ))}
                </div>
              )}
            </div>
          );
        })}
      </div>

      <div style={{ background: C.card, borderRadius: 10, padding: "16px 20px", boxShadow: "0 1px 5px rgba(0,0,0,0.15)" }}>
        <div style={{ fontFamily: "Georgia, serif", fontSize: 16, marginBottom: 4 }}>Individual players (no super agent)</div>
        <div style={{ color: C.mute, fontSize: 12.5, marginBottom: 12 }}>Settled one by one; each has their own tipback deal. House-backed players are managed on their own tab.</div>
        {inds.map((e) => (
          <div key={e.id} style={{ display: "flex", alignItems: "center", gap: 10, borderTop: `1px solid ${C.line}`, padding: "9px 0" }}>
            <span style={{ fontWeight: 700, fontSize: 14 }}>{e.name}</span>
            <span style={{ color: C.mute, fontSize: 12 }}>{e.id} · tips {fmt(e.fee)}</span>
            {newBadge(!cfg.confirmedPlayers[e.id])}
            <span style={{ marginLeft: "auto" }}>
              <PctInput value={cfg.playerDeals[e.id] ?? cfg.defaultTB} onChange={(v) => setPlayer(e.id, v ?? cfg.defaultTB, true)} />
            </span>
          </div>
        ))}
      </div>
    </div>
  );
}


// ———— Cross-module helpers ————
async function loadFishTankModel() {
  try {
    const [c, d] = await Promise.all([store.get("fishtank-config-v4"), store.get("fishtank-lastweek-v4")]);
    if (!d?.value) return null;
    const wkData = JSON.parse(d.value);
    if (!wkData.players?.length) return null;
    const saved = c?.value ? JSON.parse(c.value) : {};
    const ftCfg = { ...DEFAULT_CONFIG, ...saved,
      ownAccounts: saved.ownAccounts || DEFAULT_CONFIG.ownAccounts,
      fees: saved.fees || DEFAULT_CONFIG.fees,
      backed: saved.backed || DEFAULT_CONFIG.backed,
      umbrellas: saved.umbrellas || DEFAULT_CONFIG.umbrellas };
    return { model: buildModel(wkData.players, ftCfg, wkData.weekAdj || {}, wkData.period || ""), period: wkData.period || "", cfg: ftCfg };
  } catch (e) { return null; }
}
// rows for a set of usernames out of the fish tank week (ext + backed players)
function ftRowsForNames(ft, nameSet) {
  if (!ft) return [];
  const rows = [];
  const label = `Fish Tank${ft.period ? ` (${ft.period})` : ""}`;
  ft.model.entities.forEach((e) => {
    if (e.type === "backed") {
      if (nameSet.has(e.name.trim().toLowerCase())) rows.push({ id: "ft-" + e.key, clubName: label, name: e.name, customDeal: e.dealType === "action" ? `action buy ${e.actionPct}%` : `makeup · RB ${e.rb}%`, pnl: e.pnl, tips: e.fee, tipback: e.tipback, settlement: e.settlement, margin: 0, played: true });
    } else {
      e.members.forEach((mm) => {
        if (nameSet.has(mm.name.trim().toLowerCase())) rows.push({ id: "ft-" + mm.memberId, clubName: label, name: mm.name, customDeal: `TB ${mm.tbPct}%${mm.actionTaxPct ? ` · action ${mm.actionTaxPct}%` : ""}`, pnl: mm.pnl, tips: mm.fee, tipback: mm.tipback, settlement: mm.settlement, margin: 0, played: true });
      });
    }
  });
  return rows;
}
const makeNameMapper = (personAliases) => {
  const map = {};
  (personAliases || []).forEach((per) => per.usernames.forEach((u) => (map[u.trim().toLowerCase()] = per.name)));
  return (name) => map[String(name).trim().toLowerCase()] || name;
};

// ════════════════════════════════════════════════════════════════
// MY CLUBS — personal agent downlines (manual weekly entry)
// ════════════════════════════════════════════════════════════════
// Per player:  tipback = tips × TB%   (if TR on)
//              net     = P&L + tipback
//              TR      = TR% × base   (if TR on) — taxes wins, rebates losses
//                        base: net (P&L+tipback) · P&L only · P&L+tips (gross)
//              settle  = (net − TR) × (1 − rebate%) × conversion
// Per club (what YOU get from the club, the revenue side): same shape with the
// club's TB% and TR% — your margin is the difference.

const uid = () => Math.random().toString(36).slice(2, 9);
const SEED_CLUB_NAMES = ["Pumpkin", "Blackwater", "Betflix", "Socal", "Vans", "Don't Tilt", "Rafolini", "Bouncy Castle", "Aces Fortune", "Straddle Up", "Straddle Up (Xander)", "Ace Chasers", "Pineapple PC", "Trap City", "TPA Tiny", "OneTime", "45th Street"];
const FORMULA_VARS = ["pnl", "tips", "tb", "tr", "rebate", "tipback", "net", "gross"];
const FORMULA_HELP = "pnl · tips · tb · tr · tipback (tips×tb%) · net (pnl+tipback) · gross (pnl+tips) — plus min, max, abs, round";
const DEFAULT_FORMULA = "net - tr/100*net";
const NEW_CLUB = (name) => ({
  id: uid(), name, conv: 100, actionBase: "net", useFormula: false, formula: DEFAULT_FORMULA,
  clubTB: 80, clubAction: 0,
  players: [],
});
// Older saves used toggles + a separate rebate; fold everything into tb/tr.
function normalizeAgent(a) {
  const clubs = (a.clubs || []).map((c) => {
    const clubTB = c.useClubTB === false ? 0 : c.clubTB || 0;
    const clubAction = c.useClubAction === undefined ? c.clubAction || 0 : (c.useClubAction ? c.clubAction || 0 : 0);
    const players = (c.players || []).map((p) => {
      const tb = p.useTB === false ? 0 : p.tb || 0;
      let tr = p.useAction === undefined ? p.tr ?? p.action ?? 0 : (p.useAction ? p.action || 0 : 0);
      if (!tr && p.rebate) tr = p.rebate; // same knob under the old split naming
      return { id: p.id, name: p.name, tb, tr };
    });
    const { def, useClubTB, useClubAction, ...rest } = c;
    return { ...rest, clubTB, clubAction, players };
  });
  return { ...a, clubs };
}
const AGENT_DEFAULT = { clubs: SEED_CLUB_NAMES.map((n) => ({ ...NEW_CLUB(n), id: n.toLowerCase().replace(/[^a-z0-9]+/g, "-") })), umbrellas: [], weeks: {}, currentWeek: "", personAliases: [], myAccounts: [] };

const dealLabel = (p) =>
  [(p.tb || 0) > 0 ? `TB ${p.tb}%` : "no TB", (p.tr || 0) > 0 ? `TR ${p.tr}%` : null].filter(Boolean).join(" · ");

const TR_BASES = [["net", "of net (P&L + tip back)"], ["pnl", "of P&L only"], ["gross", "of P&L + tips"]];
const trBaseLabel = (b) => (b === "pnl" ? "P&L" : b === "gross" ? "P&L+tips" : "net");

// Custom formula support. Variables are plain numbers; percentages come in as
// percent values (tb 75 means 75%). Returns the settlement before conversion.
function compileFormula(expr) {
  if (!expr || !expr.trim()) return null;
  if (!/^[0-9a-zA-Z_+\-*/().,%\s<>=?:&|!]+$/.test(expr)) throw new Error("Invalid character in formula");
  const allowed = new Set([...FORMULA_VARS, "min", "max", "abs", "round"]);
  for (const id of expr.match(/[a-zA-Z_][a-zA-Z0-9_]*/g) || []) {
    if (!allowed.has(id)) throw new Error(`Unknown name: ${id}`);
  }
  const fn = new Function(...FORMULA_VARS, "min", "max", "abs", "round", `"use strict"; return (${expr});`);
  const probe = fn(100, 100, 75, 10, 0, 75, 175, 200, Math.min, Math.max, Math.abs, Math.round);
  if (typeof probe !== "number" || !isFinite(probe)) throw new Error("Formula must produce a number");
  return (v) => fn(v.pnl, v.tips, v.tb, v.tr, v.rebate, v.tipback, v.net, v.gross, Math.min, Math.max, Math.abs, Math.round);
}

function settleLine(pnl, tips, d, club, conv) {
  const tipback = (tips * (d.tb || 0)) / 100;
  const net = pnl + tipback;
  const gross = pnl + tips;
  const tr = d.tr || 0;
  if (club._fn) {
    let s;
    try { s = club._fn({ pnl, tips, tb: d.tb || 0, tr, rebate: 0, tipback, net, gross }); }
    catch (e) { s = net; }
    if (!isFinite(s)) s = net;
    return { tipback, net, actionCut: net - s, settlement: s * conv };
  }
  const trBase = club.actionBase === "pnl" ? pnl : club.actionBase === "gross" ? gross : net;
  const actionCut = (trBase * tr) / 100;
  return { tipback, net, actionCut, settlement: (net - actionCut) * conv };
}

function computeAgent(acfg, week) {
  const entries = week?.entries || {};
  const adjustments = week?.adjustments || [];
  const myAccSet = new Set((acfg.myAccounts || []).map((n) => n.trim().toLowerCase()).filter(Boolean));
  const clubs = acfg.clubs.map((club) => {
    const conv = (club.conv ?? 100) / 100;
    const base = club.actionBase || "net";
    let _fn = null, formulaError = null;
    if (club.useFormula) {
      try { _fn = compileFormula(club.formula); } catch (e) { formulaError = e.message; }
    }
    const ctx = { ...club, actionBase: base, _fn };
    const players = club.players.map((p) => {
      const e = entries[p.id] || {};
      const pnl = +e.pnl || 0, tips = +e.tips || 0;
      const played = pnl !== 0 || tips !== 0;
      const isMine = myAccSet.has(p.name.trim().toLowerCase());
      const clubDeal = { tb: club.clubTB || 0, tr: club.clubAction || 0 };
      const eff = isMine ? clubDeal : p; // your own accounts ride the club's deal automatically
      const mine = settleLine(pnl, tips, eff, ctx, conv);
      // what the club pays you for this player's action
      const clubSide = settleLine(pnl, tips, clubDeal, ctx, conv);
      return { ...p, tb: eff.tb, tr: eff.tr, isMine, clubId: club.id, clubName: club.name, pnl, tips, played,
        tipback: mine.tipback, net: mine.net, actionCut: mine.actionCut * conv,
        settlement: mine.settlement, clubValue: clubSide.settlement,
        margin: clubSide.settlement - mine.settlement,
        rakeMargin: ((tips * ((club.clubTB || 0) - (eff.tb || 0))) / 100) * conv,
        // margin contribution: what you charge the player minus what the club charges you
        actionMargin: (mine.actionCut - clubSide.actionCut) * conv };
    });
    const sum = (f) => players.reduce((a, p) => a + p[f], 0);
    const clubAdj = adjustments.filter((a) => a.clubId === club.id).reduce((a, x) => a + (+x.amount || 0), 0);
    const clubSettlement = sum("clubValue") + clubAdj;
    return { ...club, conv, base, formulaError, playersC: players, pnl: sum("pnl"), tips: sum("tips"),
      settlements: sum("settlement"), clubValue: sum("clubValue"), rakeMargin: sum("rakeMargin"),
      actionMargin: sum("actionMargin"), margin: sum("margin"),
      clubAdj, clubSettlement, active: players.some((p) => p.played) || clubAdj !== 0 };
  });
  const allPlayers = clubs.flatMap((c) => c.playersC);
  const umbrellas = (acfg.umbrellas || []).map((u) => {
    const members = allPlayers.filter((p) => u.playerIds.includes(p.id));
    const played = members.filter((m) => m.played);
    return { ...u, members, played, settlement: played.reduce((a, m) => a + m.settlement, 0) };
  });
  const inUmbrella = new Set((acfg.umbrellas || []).flatMap((u) => u.playerIds));
  const globalAdj = adjustments.filter((a) => !a.clubId);
  const globalAdjTotal = globalAdj.reduce((a, x) => a + (+x.amount || 0), 0);
  const T = (f) => clubs.reduce((a, c) => a + c[f], 0);
  const totals = { pnl: T("pnl"), tips: T("tips"), settlements: T("settlements"), clubValue: T("clubValue"),
    rakeMargin: T("rakeMargin"), actionMargin: T("actionMargin"), margin: T("margin"),
    clubSettlements: T("clubSettlement"), globalAdjTotal };
  totals.cashNet = totals.clubSettlements - totals.settlements + totals.globalAdjTotal;
  return { clubs, allPlayers, umbrellas, inUmbrella, totals, globalAdj };
}

// v1 (flat fields) or v2 (deal models) → v3 inline fields
function migrateAgent(old) {
  const models = old.models || null;
  const clubs = (old.clubs || []).map((c) => {
    const players = (c.players || []).map((p) => {
      let tb = p.tb, action = p.action ?? 0, rebate = p.rebate ?? 0;
      if (models) {
        const m = models.find((x) => x.id === p.modelId) || models.find((x) => x.id === c.defaultModelId) || {};
        tb = p.tbOverride ?? m.tb ?? 75; action = m.action ?? 0; rebate = m.rebate ?? 0;
      }
      tb = tb ?? 75;
      return { id: p.id || uid(), name: p.name, tb, action, rebate, useTB: (tb || 0) > 0, useAction: (action || 0) > 0 };
    });
    return { ...NEW_CLUB(c.name), id: c.id, name: c.name, conv: c.conv ?? 100,
      clubTB: c.clubTB ?? 80, clubAction: 0, players };
  });
  return { clubs, umbrellas: old.umbrellas || [], weeks: old.weeks || {}, currentWeek: old.currentWeek || "" };
}

function AgentClubs({ theme }) {
  const [acfg, setAcfg] = useState(AGENT_DEFAULT);
  const [tab, setTab] = useState("entry");
  const [loaded, setLoaded] = useState(false);
  const [exportData, setExportData] = useState(null);
  const [ft, setFt] = useState(null);
  useEffect(() => { (async () => setFt(await loadFishTankModel()))(); }, []);

  useEffect(() => {
    (async () => {
      for (const [key, needsMig] of [["agentclubs-v3", false], ["agentclubs-v2", true], ["agentclubs-v1", true]]) {
        try {
          const c = await store.get(key);
          if (c?.value) {
            const parsed = JSON.parse(c.value);
            const next = normalizeAgent(needsMig ? migrateAgent(parsed) : { ...AGENT_DEFAULT, ...parsed });
            setAcfg(next);
            if (needsMig) { try { await store.set("agentclubs-v3", JSON.stringify(next)); } catch (e) {} }
            setLoaded(true); return;
          }
        } catch (e) {}
      }
      setLoaded(true);
    })();
  }, []);
  const save = async (next) => { setAcfg(next); try { await store.set("agentclubs-v3", JSON.stringify(next)); } catch (e) {} };
  const up = (patch) => save({ ...acfg, ...patch });

  const weekKeys = Object.keys(acfg.weeks);
  const wk = acfg.currentWeek && acfg.weeks[acfg.currentWeek] ? acfg.currentWeek : weekKeys[weekKeys.length - 1] || "";
  const week = acfg.weeks[wk];
  const model = useMemo(() => computeAgent(acfg, week), [acfg, week]);

  const newWeek = () => {
    const label = window.prompt("Week label (e.g. 07/20 - 07/26):");
    if (!label || acfg.weeks[label]) return;
    up({ weeks: { ...acfg.weeks, [label]: { entries: {}, adjustments: [] } }, currentWeek: label });
  };
  const setEntry = (pid, field, v) => {
    const w = acfg.weeks[wk]; if (!w) return;
    up({ weeks: { ...acfg.weeks, [wk]: { ...w, entries: { ...w.entries, [pid]: { ...(w.entries[pid] || {}), [field]: v } } } } });
  };
  const setAdjs = (adjustments) => {
    const w = acfg.weeks[wk]; if (!w) return;
    up({ weeks: { ...acfg.weeks, [wk]: { ...w, adjustments } } });
  };

  if (!loaded) return <div style={{ padding: 40, color: C.mute }}>Loading…</div>;

  return (
    <div>
      <ExportModal data={exportData} onClose={() => setExportData(null)} />
      <div style={{ display: "flex", gap: 4, padding: "10px 26px 0", borderBottom: `2px solid ${C.line}`, background: C.paper, flexWrap: "wrap", alignItems: "center" }}>
        {[["entry", "Weekly entry"], ["summary", "Summary & reports"], ["clubs", "Clubs & deals"]].map(([k, label]) => (
          <button key={k} onClick={() => setTab(k)} style={{
            border: "none", cursor: "pointer", padding: "9px 16px", fontSize: 13.5, fontWeight: 700,
            background: tab === k ? C.card : "transparent", color: tab === k ? C.ink : C.mute,
            borderRadius: "8px 8px 0 0", marginBottom: -2 }}>{label}</button>
        ))}
        <div style={{ marginLeft: "auto", display: "flex", gap: 8, alignItems: "center", paddingBottom: 6 }}>
          <select value={wk} onChange={(e) => up({ currentWeek: e.target.value })} style={{ ...inputS, fontSize: 12.5 }}>
            {weekKeys.length === 0 && <option value="">No weeks yet</option>}
            {weekKeys.map((k) => <option key={k} value={k}>{k}</option>)}
          </select>
          {wk && <button title="Delete this week" onClick={() => {
            if (!window.confirm(`Delete week "${wk}" and all its entered data? This can't be undone.`)) return;
            const weeks = { ...acfg.weeks }; delete weeks[wk];
            const rest = Object.keys(weeks);
            up({ weeks, currentWeek: rest[rest.length - 1] || "" });
          }} style={{ border: "none", background: "none", color: C.red, cursor: "pointer", fontSize: 15 }}>×</button>}
          <Btn tone="gold" small onClick={newWeek}>+ New week</Btn>
        </div>
      </div>
      <div style={{ padding: "20px 26px 60px", maxWidth: 1180, margin: "0 auto" }}>
        {!week && tab !== "clubs" && (
          <div style={{ background: C.card, border: `1px dashed ${C.gold}`, borderRadius: 10, padding: "44px 30px", textAlign: "center" }}>
            <div style={{ fontFamily: "Georgia, serif", fontSize: 20, marginBottom: 8 }}>Start a week</div>
            <div style={{ color: C.mute, fontSize: 14, marginBottom: 16 }}>Create a week, then type each player's P&L and tips.</div>
            <Btn onClick={newWeek}>+ New week</Btn>
          </div>
        )}
        {week && tab === "entry" && <AgentEntry model={model} setEntry={setEntry} setAdjs={setAdjs} week={week} acfg={acfg} />}
        {week && tab === "summary" && <AgentSummary model={model} wk={wk} setExportData={setExportData} acfg={acfg} up={up} ft={ft} />}
        {tab === "clubs" && <AgentClubsSetup acfg={acfg} up={up} />}
      </div>
    </div>
  );
}

function AgentEntry({ model, setEntry, setAdjs, week, acfg }) {
  const [adjClub, setAdjClub] = useState("");
  const adjustments = week.adjustments || [];
  return (
    <div>
      <div style={{ color: C.mute, fontSize: 12.5, marginBottom: 12 }}>
        Enter each player's <b>P&L</b> and <b>Tips</b> in club currency; blank = no play. <span style={{ color: C.green, fontWeight: 700 }}>Green</span> = you pay them · <span style={{ color: C.red, fontWeight: 700 }}>red</span> = they pay you. <b>Margin</b> is what you keep after the club pays you for that player.
      </div>
      {model.clubs.filter((c) => c.players.length > 0).map((c) => (
        <div key={c.id} style={{ background: C.card, borderRadius: 10, marginBottom: 10, overflow: "hidden", boxShadow: "0 1px 5px rgba(0,0,0,0.12)" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12, padding: "10px 16px", background: C.cream, flexWrap: "wrap" }}>
            <span style={{ fontWeight: 700, fontSize: 14.5 }}>{c.name}</span>
            {c.conv !== 1 && <Pill tone="blue">conv {n2(c.conv * 100)}%</Pill>}
            <span style={{ color: C.mute, fontSize: 12 }}>
              from club: TB {c.clubTB || 0}%{(c.clubAction || 0) > 0 ? ` · TR ${c.clubAction}%` : ""}
            </span>
            {c.useFormula && <Pill tone={c.formulaError ? "red" : "blue"}>{c.formulaError ? "formula error" : "custom formula"}</Pill>}
            <span style={{ marginLeft: "auto", fontSize: 12.5 }}>players {money(c.settlements)} · club {money(c.clubSettlement)} · margin {money(c.margin)}</span>
          </div>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr>
              <th style={{ ...th, textAlign: "left" }}>Player</th>
              <th style={{ ...th, textAlign: "left" }}>Deal</th>
              <th style={th}>P&L</th><th style={th}>Tips</th>
              <th style={th}>Tipback</th><th style={th}>Settlement</th><th style={th}>Your margin</th>
            </tr></thead>
            <tbody>
              {c.playersC.map((p, i) => (
                <tr key={p.id} style={{ background: i % 2 ? C.rowAlt : C.card, borderTop: `1px solid ${C.line}` }}>
                  <td style={{ ...tdL, fontWeight: 600 }}>{p.name}{model.inUmbrella.has(p.id) && <span style={{ marginLeft: 6 }}><Pill tone="blue">{(acfg.umbrellas.find((u) => u.playerIds.includes(p.id)) || {}).name}</Pill></span>}</td>
                  <td style={{ ...tdL, color: C.mute, fontSize: 11.5 }}>{dealLabel(p)}</td>
                  <td style={td}><NumInput width={92} value={(week.entries[p.id] || {}).pnl ?? ""} onChange={(v) => setEntry(p.id, "pnl", v)} /></td>
                  <td style={td}><NumInput width={82} value={(week.entries[p.id] || {}).tips ?? ""} onChange={(v) => setEntry(p.id, "tips", v)} /></td>
                  <td style={td}>{p.played ? fmt(p.tipback) : "—"}</td>
                  <td style={td}>{p.played ? money(p.settlement) : "—"}</td>
                  <td style={td}>{p.played ? money(p.margin) : "—"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ))}
      <Card title="Adjustments (stakes, transfers, one-offs)" right={
        <span style={{ display: "flex", gap: 8 }}>
          <select value={adjClub} onChange={(e) => setAdjClub(e.target.value)} style={{ ...inputS, fontSize: 12 }}>
            <option value="">Not club-specific</option>
            {acfg.clubs.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
          </select>
          <Btn tone="ghost" small onClick={() => setAdjs([...adjustments, { id: uid(), clubId: adjClub || null, label: "New item", amount: 0 }])}>+ Add</Btn>
        </span>
      }>
        {adjustments.length === 0 && <div style={{ color: C.mute, fontSize: 13 }}>None this week. Positive = money to you, negative = money out.</div>}
        {adjustments.map((a) => (
          <div key={a.id} style={{ display: "flex", gap: 10, alignItems: "center", borderTop: `1px solid ${C.line}`, padding: "8px 0" }}>
            <input value={a.label} onChange={(e) => setAdjs(adjustments.map((x) => x.id === a.id ? { ...x, label: e.target.value } : x))} style={{ ...inputS, flex: 1 }} />
            <span style={{ color: C.mute, fontSize: 12 }}>{a.clubId ? (acfg.clubs.find((c) => c.id === a.clubId)?.name || "?") : "general"}</span>
            <NumInput width={100} value={a.amount} onChange={(v) => setAdjs(adjustments.map((x) => x.id === a.id ? { ...x, amount: v } : x))} />
            <button onClick={() => setAdjs(adjustments.filter((x) => x.id !== a.id))} style={{ border: "none", background: "none", color: C.red, cursor: "pointer", fontSize: 15 }}>×</button>
          </div>
        ))}
      </Card>
    </div>
  );
}

function AgentSummary({ model, wk, setExportData, acfg, up, ft }) {
  const t = model.totals;
  const active = model.clubs.filter((c) => c.active);
  const [reportSel, setReportSel] = useState("");
  const [aliasOpen, setAliasOpen] = useState(false);
  const persons = acfg.personAliases || [];
  const bundled = new Set(persons.flatMap((p) => p.usernames.map((u) => u.trim().toLowerCase())));
  const ftNames = ft ? [...new Set(ft.model.entities.flatMap((e) => e.type === "backed" ? [e.name] : e.members.map((m) => m.name)))] : [];
  const rawNames = [...new Set([...model.allPlayers.map((p) => p.name.trim()), ...ftNames.map((n) => n.trim())])]
    .filter((n) => !bundled.has(n.toLowerCase())).sort((a, b) => a.localeCompare(b));

  // resolve selection → username set + display name
  let reportLabel = "", nameSet = new Set();
  if (reportSel.startsWith("person:")) {
    const per = persons.find((p) => p.id === reportSel.slice(7));
    if (per) { reportLabel = per.name; nameSet = new Set(per.usernames.map((u) => u.trim().toLowerCase())); }
  } else if (reportSel.startsWith("u:")) {
    reportLabel = reportSel.slice(2); nameSet = new Set([reportLabel.toLowerCase()]);
  }
  const mcRows = model.allPlayers.filter((p) => nameSet.has(p.name.trim().toLowerCase()) && p.played);
  const ftRows = ftRowsForNames(ft, nameSet);
  // clubs owned by this person also fold into their report (sign flipped: positive = you pay them)
  const clubRows = model.clubs.filter((c) => c.active && (c.owner || "").trim() && (nameSet.has(c.owner.trim().toLowerCase()) || c.owner.trim().toLowerCase() === reportLabel.toLowerCase()))
    .map((c) => ({ id: "club-" + c.id, clubName: c.name, name: c.owner, customDeal: "club settlement", pnl: c.pnl, tips: c.tips, tipback: 0, settlement: -c.clubSettlement, margin: 0, played: true }));
  const reportRows = [...ftRows, ...mcRows, ...clubRows];
  const reportTotal = reportRows.reduce((a, p) => a + p.settlement, 0);
  // my play + fish tank side for the total card
  const myAcc = new Set((acfg.myAccounts || []).map((n) => n.trim().toLowerCase()).filter(Boolean));
  const myPlay = model.allPlayers.filter((p) => p.played && myAcc.has(p.name.trim().toLowerCase()));
  const myPlayTotal = myPlay.reduce((a, p) => a + p.settlement, 0);
  const ftEntitle = ft ? ft.model.entitle.ak : null;
  const weekTotal = (ftEntitle || 0) + t.margin + myPlayTotal + t.globalAdjTotal;
  const row = (label, val, opts = {}) => (
    <div style={{ display: "flex", padding: "6px 0", fontSize: 13.5 }}>
      <span style={{ color: opts.bold ? C.ink : C.mute, fontWeight: opts.bold ? 700 : 400 }}>{label}</span>
      <span style={{ marginLeft: "auto", fontVariantNumeric: "tabular-nums", fontWeight: opts.bold ? 700 : 500 }}>{money(val)}</span>
    </div>
  );
  const copyRows = (title, rows) => setExportData({ title, text: toTSV(["Club", "Player", "Deal", "P&L", "Tips", "Tipback", "Settlement", "Your margin"], rows.map((p) => [p.clubName, p.name, p.customDeal || dealLabel(p), p.pnl.toFixed(2), p.tips.toFixed(2), p.tipback.toFixed(2), p.settlement.toFixed(2), (p.margin || 0).toFixed(2)])) });

  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12, flexWrap: "wrap" }}>
        <div style={{ fontFamily: "Georgia, serif", fontSize: 19 }}>Week summary · {wk}</div>
        <div style={{ marginLeft: "auto", display: "flex", gap: 8 }}>
          <Btn tone="gold" small onClick={() => downloadAgentWorkbook(model, wk)}>Download Excel</Btn>
          <Btn tone="ghost" small onClick={() => copyRows(`My clubs · ${wk}`, model.allPlayers.filter((p) => p.played))}>Copy</Btn>
        </div>
      </div>
      <div style={{ marginBottom: 14 }}>
        <Card title={`Total week P&L${ft?.period ? ` · Fish Tank ${ft.period}` : ""}`}>
          {ftEntitle == null
            ? <div style={{ color: C.mute, fontSize: 12.5, marginBottom: 4 }}>No Fish Tank week loaded — upload one in the Fish Tank mode and it appears here automatically.</div>
            : row("Fish Tank — your entitlement (own play + backed books + ½ profit)", ftEntitle)}
          {row("My Clubs — margin on players", t.margin)}
          {row(`My play on other clubs${myPlay.length ? ` (${myPlay.map((p) => p.name).join(", ")})` : ""}`, myPlayTotal)}
          {myAcc.size === 0 && <div style={{ color: C.mute, fontSize: 11.5, margin: "2px 0 4px" }}>List your own usernames under "My accounts" in Clubs & deals to count your personal play here.</div>}
          {row("General adjustments", t.globalAdjTotal)}
          <div style={{ borderTop: `1px solid ${C.line}` }} />
          {row("TOTAL", weekTotal, { bold: true })}
        </Card>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(300px, 1fr))", gap: 14, marginBottom: 16 }}>
        <Card title="Your week">
          {row("Club side (clubs pay you)", t.clubSettlements)}
          {row("Player settlements", t.settlements)}
          {row("General adjustments", t.globalAdjTotal)}
          <div style={{ borderTop: `1px solid ${C.line}` }} />
          {row("Cash net", t.cashNet, { bold: true })}
        </Card>
        <Card title="Where the margin came from">
          {row("TB margin (club TB − player TB)", t.rakeMargin)}
          {row("TR margin (player TR − club TR)", t.actionMargin)}
          <div style={{ borderTop: `1px solid ${C.line}` }} />
          {row("Total margin on players", t.margin, { bold: true })}
        </Card>
      </div>
      <div style={{ marginBottom: 16 }}>
        <Card title="Player report" right={
          <span style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <select value={reportSel} onChange={(e) => setReportSel(e.target.value)} style={{ ...inputS, fontSize: 12.5, minWidth: 180 }}>
              <option value="">Pick a player…</option>
              {persons.length > 0 && <optgroup label="People (bundles)">
                {persons.map((p) => <option key={p.id} value={"person:" + p.id}>{p.name} ({p.usernames.length})</option>)}
              </optgroup>}
              <optgroup label="Usernames">
                {rawNames.map((n) => <option key={n} value={"u:" + n}>{n}</option>)}
              </optgroup>
            </select>
            <Btn tone="ghost" small onClick={() => setAliasOpen(!aliasOpen)}>{aliasOpen ? "Hide bundles" : "Bundles"}</Btn>
            {reportRows.length > 0 && <Btn tone="gold" small onClick={() => downloadPlayerExcel(reportLabel, reportRows, wk)}>Excel</Btn>}
            {reportRows.length > 0 && <Btn tone="ghost" small onClick={() => copyRows(`${reportLabel} · ${wk}`, reportRows)}>Copy</Btn>}
          </span>
        }>
          {aliasOpen && <PersonBundles acfg={acfg} up={up} allNames={[...new Set([...model.allPlayers.map((p) => p.name.trim()), ...ftNames])]} />}
          {!reportSel && <div style={{ color: C.mute, fontSize: 13 }}>One report across Fish Tank and every club — pick a username or a bundled person. Clubs whose owner matches also fold in.</div>}
          {reportSel && reportRows.length === 0 && <div style={{ color: C.mute, fontSize: 13 }}>No activity recorded for {reportLabel} this week.</div>}
          {reportRows.length > 0 && (
            <table style={{ width: "100%", borderCollapse: "collapse" }}>
              <tbody>
                {reportRows.map((p, i) => (
                  <tr key={p.id} style={{ background: i % 2 ? C.rowAlt : C.card, borderTop: `1px solid ${C.line}` }}>
                    <td style={{ ...tdL, fontWeight: 600 }}>{p.clubName}{p.name && p.name.toLowerCase() !== reportLabel.toLowerCase() ? <span style={{ color: C.mute, fontWeight: 400, fontSize: 11 }}> · {p.name}</span> : null}</td>
                    <td style={{ ...tdL, color: C.mute, fontSize: 12 }}>{p.customDeal || dealLabel(p)}</td>
                    <td style={td}>P&L {fmt(p.pnl)}</td>
                    <td style={td}>tips {fmt(p.tips)}</td>
                    <td style={td}>{money(p.settlement)}</td>
                  </tr>
                ))}
                <tr style={{ background: C.cream, borderTop: `2px solid ${C.gold}` }}>
                  <td style={{ ...tdL, fontWeight: 700 }} colSpan={4}>TOTAL</td>
                  <td style={{ ...td, fontWeight: 700 }}>{money(reportTotal)}</td>
                </tr>
              </tbody>
            </table>
          )}
        </Card>
      </div>
      {model.umbrellas.filter((u) => u.played.length > 0).map((u) => (
        <div key={u.id} style={{ background: C.card, borderRadius: 10, marginBottom: 10, overflow: "hidden", boxShadow: "0 1px 5px rgba(0,0,0,0.12)" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12, padding: "10px 16px", background: C.cream }}>
            <span style={{ fontWeight: 700 }}>{u.name}</span>
            <Pill tone="blue">umbrella · {u.played.length} lines</Pill>
            <span style={{ marginLeft: "auto", fontSize: 12.5 }}>settles as one {money(u.settlement)}</span>
            <Btn tone="gold" small onClick={() => downloadPlayerExcel(u.name, u.played, wk)}>Excel</Btn>
          </div>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <tbody>
              {u.played.map((p, i) => (
                <tr key={p.id} style={{ background: i % 2 ? C.rowAlt : C.card, borderTop: `1px solid ${C.line}` }}>
                  <td style={{ ...tdL, fontWeight: 600 }}>{p.name}</td>
                  <td style={{ ...tdL, color: C.mute, fontSize: 12 }}>{p.clubName}</td>
                  <td style={td}>P&L {fmt(p.pnl)}</td>
                  <td style={td}>tips {fmt(p.tips)}</td>
                  <td style={td}>{money(p.settlement)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ))}
      {active.map((c) => (
        <div key={c.id} style={{ background: C.card, borderRadius: 10, marginBottom: 10, overflow: "hidden", boxShadow: "0 1px 5px rgba(0,0,0,0.12)" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12, padding: "10px 16px", background: C.cream, flexWrap: "wrap" }}>
            <span style={{ fontWeight: 700 }}>{c.name}</span>
            <span style={{ color: C.mute, fontSize: 12 }}>tips {fmt(c.tips)} · P&L {fmt(c.pnl)}</span>
            <span style={{ marginLeft: "auto", fontSize: 12.5 }}>club {money(c.clubSettlement)} · margin {money(c.margin)}</span>
          </div>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <tbody>
              {c.playersC.filter((p) => p.played).map((p, i) => (
                <tr key={p.id} style={{ background: i % 2 ? C.rowAlt : C.card, borderTop: `1px solid ${C.line}` }}>
                  <td style={{ ...tdL, fontWeight: 600, width: "24%" }}>{p.name}</td>
                  <td style={{ ...tdL, color: C.mute, fontSize: 12 }}>{dealLabel(p)}</td>
                  <td style={td}>P&L {fmt(p.pnl)}</td>
                  <td style={td}>tips {fmt(p.tips)}</td>
                  <td style={td}>{money(p.settlement)}</td>
                  <td style={td}>margin {money(p.margin)}</td>
                </tr>
              ))}
              {c.clubAdj !== 0 && (
                <tr style={{ borderTop: `1px solid ${C.line}` }}><td style={{ ...tdL, color: C.mute }} colSpan={5}>Club adjustments</td><td style={td}>{money(c.clubAdj)}</td></tr>
              )}
            </tbody>
          </table>
        </div>
      ))}
    </div>
  );
}

// ———— Clubs & deals: inline fields, toggles, club-side revenue ————
const Toggle = ({ on, onClick, label }) => (
  <button onClick={onClick} style={{
    border: `1px solid ${on ? C.goldDark : C.line}`, background: on ? C.gold : C.surface,
    color: on ? "var(--onGold)" : C.mute, borderRadius: 12, padding: "2px 10px",
    fontSize: 11, fontWeight: 700, cursor: "pointer", whiteSpace: "nowrap" }}>{label}</button>
);

function FormulaEditor({ club, setClub }) {
  const [txt, setTxt] = useState(club.formula || DEFAULT_FORMULA);
  useEffect(() => setTxt(club.formula || DEFAULT_FORMULA), [club.formula]);
  let err = null, preview = null;
  try {
    const fn = compileFormula(txt);
    if (fn) {
      const pnl = -1000, tips = 500, tb = club.clubTB || 0, tr = club.clubAction || 0;
      const tipback = (tips * tb) / 100;
      preview = fn({ pnl, tips, tb, tr, rebate: 0, tipback, net: pnl + tipback, gross: pnl + tips });
    }
  } catch (e) { err = e.message; }
  return (
    <div style={{ padding: "10px 16px", borderBottom: `1px solid ${C.line}`, background: C.rowAlt }}>
      <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
        <span style={{ fontSize: 12, fontWeight: 700, color: C.goldDark, textTransform: "uppercase", letterSpacing: "0.06em", minWidth: 120 }}>Formula</span>
        <input value={txt} onChange={(e) => setTxt(e.target.value)}
          onBlur={() => setClub(club.id, { formula: txt })}
          onKeyDown={(e) => e.key === "Enter" && e.currentTarget.blur()}
          spellCheck={false}
          style={{ ...inputS, flex: 1, minWidth: 260, fontFamily: "ui-monospace, Menlo, monospace", fontSize: 12.5, borderColor: err ? C.red : C.gold }} />
        {err
          ? <span style={{ color: C.red, fontSize: 11.5 }}>{err}</span>
          : <span style={{ color: C.mute, fontSize: 11.5 }}>P&L −1,000 / tips 500 → <b style={{ color: C.ink }}>{preview == null ? "—" : fmt(preview)}</b></span>}
      </div>
      <div style={{ color: C.mute, fontSize: 11, marginTop: 5 }}>
        Variables: {FORMULA_HELP}. This IS the club's TR formula — players and the club row run through it with their own tb / tr. Percentages are numbers, so use <code>tr/100</code>.
      </div>
      <div style={{ color: C.mute, fontSize: 11, marginTop: 3 }}>
        Examples: <code>net - tr/100*net</code> · <code>net - tr/100*gross</code> · <code>net - tr/100*max(0, net)</code>
      </div>
    </div>
  );
}

function PersonBundles({ acfg, up, allNames }) {
  const [newPerson, setNewPerson] = useState("");
  const [addName, setAddName] = useState({});
  const persons = acfg.personAliases || [];
  const setPersons = (personAliases) => up({ personAliases });
  return (
    <div style={{ background: C.rowAlt, borderRadius: 8, padding: "10px 14px", marginBottom: 12 }}>
      <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 8 }}>
        <span style={{ fontSize: 12, fontWeight: 700, color: C.goldDark, textTransform: "uppercase", letterSpacing: "0.06em" }}>Bundles — one person, many usernames</span>
        <input placeholder="New person name…" value={newPerson} onChange={(e) => setNewPerson(e.target.value)}
          onKeyDown={(e) => { if (e.key === "Enter" && newPerson.trim()) { setPersons([...persons, { id: uid(), name: newPerson.trim(), usernames: [] }]); setNewPerson(""); } }}
          style={{ ...inputS, width: 170, marginLeft: "auto" }} />
        <Btn tone="ghost" small onClick={() => { if (newPerson.trim()) { setPersons([...persons, { id: uid(), name: newPerson.trim(), usernames: [] }]); setNewPerson(""); } }}>+ Person</Btn>
      </div>
      {persons.length === 0 && <div style={{ color: C.mute, fontSize: 12.5 }}>e.g. create "Kevin" and add his three usernames — reports and Tabs then treat them as one.</div>}
      {persons.map((per) => (
        <div key={per.id} style={{ display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap", borderTop: `1px solid ${C.line}`, padding: "7px 0" }}>
          <input value={per.name} onChange={(e) => setPersons(persons.map((x) => x.id === per.id ? { ...x, name: e.target.value } : x))} style={{ ...inputS, width: 130, fontWeight: 700, fontSize: 12.5 }} />
          {per.usernames.map((u) => (
            <span key={u} style={{ background: C.surface, border: `1px solid ${C.gold}`, borderRadius: 12, padding: "2px 9px", fontSize: 11.5 }}>
              {u} <button onClick={() => setPersons(persons.map((x) => x.id === per.id ? { ...x, usernames: x.usernames.filter((y) => y !== u) } : x))} style={{ border: "none", background: "none", color: C.red, cursor: "pointer", fontSize: 11 }}>×</button>
            </span>
          ))}
          <input list={"bundle-names-" + per.id} placeholder="add username…" value={addName[per.id] || ""} onChange={(e) => setAddName({ ...addName, [per.id]: e.target.value })}
            onKeyDown={(e) => {
              if (e.key === "Enter") {
                const v = (addName[per.id] || "").trim();
                if (v && !per.usernames.includes(v)) setPersons(persons.map((x) => x.id === per.id ? { ...x, usernames: [...x.usernames, v] } : x));
                setAddName({ ...addName, [per.id]: "" });
              }
            }} style={{ ...inputS, width: 140, fontSize: 11.5 }} />
          <datalist id={"bundle-names-" + per.id}>{allNames.map((n) => <option key={n} value={n} />)}</datalist>
          <button onClick={() => setPersons(persons.filter((x) => x.id !== per.id))} style={{ marginLeft: "auto", border: "none", background: "none", color: C.red, cursor: "pointer", fontSize: 13 }}>× delete</button>
        </div>
      ))}
    </div>
  );
}

function AgentClubsSetup({ acfg, up }) {
  const [newClub, setNewClub] = useState("");
  const [umbName, setUmbName] = useState("");
  const setClub = (id, patch) => up({ clubs: acfg.clubs.map((c) => (c.id === id ? { ...c, ...patch } : c)) });
  const addClub = () => { if (!newClub.trim()) return; up({ clubs: [...acfg.clubs, NEW_CLUB(newClub.trim())] }); setNewClub(""); };
  const delClub = (cid) => { if (window.confirm("Delete this club and its roster?")) up({ clubs: acfg.clubs.filter((c) => c.id !== cid) }); };
  const addPlayer = (c) => setClub(c.id, { players: [...c.players, { id: uid(), name: "New player", tb: c.clubTB || 0, tr: c.clubAction || 0 }] });
  const setPlayer = (cid, pid, patch) => setClub(cid, { players: acfg.clubs.find((c) => c.id === cid).players.map((p) => (p.id === pid ? { ...p, ...patch } : p)) });
  const delPlayer = (cid, pid) => setClub(cid, { players: acfg.clubs.find((c) => c.id === cid).players.filter((p) => p.id !== pid) });
  const myAccSet = new Set((acfg.myAccounts || []).map((n) => n.trim().toLowerCase()).filter(Boolean));
  const addUmb = () => { if (!umbName.trim()) return; up({ umbrellas: [...(acfg.umbrellas || []), { id: uid(), name: umbName.trim(), playerIds: [] }] }); setUmbName(""); };
  const togglePlayerInUmb = (uid_, pid) => up({
    umbrellas: acfg.umbrellas.map((u) => u.id === uid_
      ? { ...u, playerIds: u.playerIds.includes(pid) ? u.playerIds.filter((x) => x !== pid) : [...u.playerIds, pid] }
      : { ...u, playerIds: u.playerIds.filter((x) => x !== pid) }) });
  const allP = acfg.clubs.flatMap((c) => c.players.map((p) => ({ ...p, clubName: c.name })));

  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: "2fr 1fr", gap: 14, marginBottom: 16, alignItems: "start" }}>
        <Card title="Umbrellas (bunch players into one settlement)" right={
          <span style={{ display: "flex", gap: 8 }}>
            <input placeholder="New umbrella name…" value={umbName} onChange={(e) => setUmbName(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addUmb()} style={{ ...inputS, width: 180 }} />
            <Btn tone="ghost" small onClick={addUmb}>+ Create</Btn>
          </span>
        }>
          {(acfg.umbrellas || []).length === 0 && <div style={{ color: C.mute, fontSize: 13 }}>None yet. Groups players across any clubs into one settlement and one report.</div>}
          {(acfg.umbrellas || []).map((u) => (
            <div key={u.id} style={{ borderTop: `1px solid ${C.line}`, padding: "10px 0" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}>
                <input value={u.name} onChange={(e) => up({ umbrellas: acfg.umbrellas.map((x) => x.id === u.id ? { ...x, name: e.target.value } : x) })} style={{ ...inputS, fontWeight: 700, width: 200 }} />
                <Pill tone="blue">{u.playerIds.length} players</Pill>
                <button onClick={() => up({ umbrellas: acfg.umbrellas.filter((x) => x.id !== u.id) })} style={{ marginLeft: "auto", border: "none", background: "none", color: C.red, cursor: "pointer", fontSize: 15 }}>× delete</button>
              </div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                {allP.map((p) => {
                  const inThis = u.playerIds.includes(p.id);
                  const inOther = !inThis && acfg.umbrellas.some((x) => x.id !== u.id && x.playerIds.includes(p.id));
                  return (
                    <button key={p.id} onClick={() => !inOther && togglePlayerInUmb(u.id, p.id)} style={{
                      padding: "3px 10px", borderRadius: 12, fontSize: 11.5, fontWeight: 600, cursor: inOther ? "default" : "pointer",
                      border: `1px solid ${inThis ? C.goldDark : C.line}`,
                      background: inThis ? C.gold : C.surface, color: inThis ? "var(--onGold)" : inOther ? "var(--chipOff)" : C.mute, opacity: inOther ? 0.55 : 1 }}>
                      {p.name} · {p.clubName}
                    </button>
                  );
                })}
              </div>
            </div>
          ))}
        </Card>
        <Card title="My accounts">
          <div style={{ fontSize: 12, color: C.mute, marginBottom: 8 }}>Your own usernames across these clubs — their play counts as "my play" in the week summary. One per line.</div>
          <textarea defaultValue={(acfg.myAccounts || []).join("\n")}
            onBlur={(e) => up({ myAccounts: e.target.value.split("\n").map((s) => s.trim()).filter(Boolean) })}
            rows={5} style={{ ...inputS, width: "100%", boxSizing: "border-box", fontFamily: "inherit", resize: "vertical" }} />
        </Card>
      </div>

      <div style={{ display: "flex", gap: 8, marginBottom: 12, alignItems: "center", flexWrap: "wrap" }}>
        <input placeholder="New club name…" value={newClub} onChange={(e) => setNewClub(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addClub()} style={{ ...inputS, width: 240 }} />
        <Btn tone="gold" small onClick={addClub}>+ Add club</Btn>
        <div style={{ marginLeft: "auto", color: C.mute, fontSize: 12 }}><b>TB</b> = tip back · <b>TR</b> = tax rebate (taxes wins, rebates losses). Toggle either off to drop it from the deal. Conversion: value of 1 club unit as % (100 = 1:1).</div>
      </div>

      {acfg.clubs.map((c) => {
        return (
          <div key={c.id} style={{ background: C.card, borderRadius: 10, marginBottom: 12, overflow: "hidden", boxShadow: "0 1px 5px rgba(0,0,0,0.12)" }}>
            <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap", padding: "10px 16px", background: C.cream }}>
              <input value={c.name} onChange={(e) => setClub(c.id, { name: e.target.value })} style={{ ...inputS, fontWeight: 700, width: 165 }} />
              <span style={{ fontSize: 12, color: C.mute }}>owner</span>
              <input value={c.owner || ""} placeholder="who's behind it" onChange={(e) => setClub(c.id, { owner: e.target.value })} style={{ ...inputS, width: 130, fontSize: 12 }} />
              <span style={{ fontSize: 12, color: C.mute }} title="1 club unit in USD, as a percent — 100 = 1:1">FX → USD</span>
              <NumInput width={64} value={c.conv} onChange={(v) => setClub(c.id, { conv: v || 100 })} />
              <span style={{ fontSize: 11.5, color: C.mute }}>%</span>
              <span style={{ fontSize: 12, color: C.mute }}>TR base</span>
              <select value={c.actionBase || "net"} disabled={!!c.useFormula} onChange={(e) => setClub(c.id, { actionBase: e.target.value })} style={{ ...inputS, padding: "4px 6px", fontSize: 12, opacity: c.useFormula ? 0.4 : 1 }}>
                {TR_BASES.map(([v, l]) => <option key={v} value={v}>{l}</option>)}
              </select>
              <Toggle on={!!c.useFormula} onClick={() => setClub(c.id, { useFormula: !c.useFormula, formula: c.formula || DEFAULT_FORMULA })} label="custom formula" />
              <button onClick={() => delClub(c.id)} style={{ marginLeft: "auto", border: "none", background: "none", color: C.red, cursor: "pointer", fontSize: 13 }}>× delete club</button>
            </div>

            {c.useFormula && <FormulaEditor club={c} setClub={setClub} />}

            {/* what the club gives YOU — type a number to turn it on, 0 = off */}
            <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap", padding: "9px 16px", borderBottom: `1px solid ${C.line}`, background: C.rowAlt }}>
              <span style={{ fontSize: 12, fontWeight: 700, color: C.goldDark, textTransform: "uppercase", letterSpacing: "0.06em", minWidth: 120 }}>From the club</span>
              <span style={{ fontSize: 11.5, color: C.mute }}>TB</span>
              <PctInput width={52} value={c.clubTB ?? 0} onChange={(v) => v != null && setClub(c.id, { clubTB: v })} />
              <span style={{ fontSize: 11.5, color: C.mute }}>TR</span>
              <PctInput width={52} value={c.clubAction ?? 0} onChange={(v) => v != null && setClub(c.id, { clubAction: v })} />
              <span style={{ marginLeft: "auto", color: C.mute, fontSize: 11.5 }}>
                your revenue side · new players start on this deal · your accounts always ride it{(c.clubAction || 0) > 0 ? ` · TR ${trBaseLabel(c.actionBase || "net")}` : ""}
              </span>
            </div>

            {c.players.length > 0 && (
              <table style={{ width: "100%", borderCollapse: "collapse" }}>
                <thead><tr>
                  <th style={{ ...th, textAlign: "left" }}>Player</th>
                  <th style={th}>TB %</th>
                  <th style={th}>TR %</th>
                  <th style={{ ...th, textAlign: "left" }}>Formula</th><th style={th}></th>
                </tr></thead>
                <tbody>
                  {c.players.map((p, i) => {
                    const isMine = myAccSet.has((p.name || "").trim().toLowerCase());
                    return (
                    <tr key={p.id} style={{ background: i % 2 ? C.rowAlt : C.card, borderTop: `1px solid ${C.line}` }}>
                      <td style={tdL}>
                        <input value={p.name} onChange={(e) => setPlayer(c.id, p.id, { name: e.target.value })} style={{ ...inputS, width: 150, fontSize: 12.5 }} />
                        {isMine && <span style={{ marginLeft: 6 }}><Pill tone="gold">you · club deal</Pill></span>}
                      </td>
                      <td style={td}>{isMine
                        ? <span style={{ color: C.mute, fontSize: 12 }}>{c.clubTB || 0}%</span>
                        : <PctInput width={52} value={p.tb ?? 0} onChange={(v) => v != null && setPlayer(c.id, p.id, { tb: v })} />}</td>
                      <td style={td}>{isMine
                        ? <span style={{ color: C.mute, fontSize: 12 }}>{c.clubAction || 0}%</span>
                        : <PctInput width={52} value={p.tr ?? 0} onChange={(v) => v != null && setPlayer(c.id, p.id, { tr: v })} />}</td>
                      <td style={{ ...tdL, color: C.goldDark, fontSize: 11, fontFamily: "ui-monospace, Menlo, monospace" }}>
                        {c.useFormula
                          ? c.formula
                          : `(P&L${(isMine ? c.clubTB : p.tb) ? ` + tips×${isMine ? c.clubTB : p.tb}%` : ""}${(isMine ? c.clubAction : p.tr) ? ` − ${isMine ? c.clubAction : p.tr}%×${trBaseLabel(c.actionBase || "net")}` : ""})`}
                      </td>
                      <td style={{ ...td, width: 36 }}><button onClick={() => delPlayer(c.id, p.id)} style={{ border: "none", background: "none", color: C.red, cursor: "pointer", fontSize: 15 }}>×</button></td>
                    </tr>
                  );})}
                </tbody>
              </table>
            )}
            <div style={{ padding: "8px 16px" }}>
              <button onClick={() => addPlayer(c)} style={{ border: `1px dashed ${C.gold}`, background: "transparent", color: C.goldDark, borderRadius: 6, padding: "6px 12px", cursor: "pointer", fontSize: 12.5, fontWeight: 700 }}>+ Add player</button>
            </div>
          </div>
        );
      })}
    </div>
  );
}

function playerSheetRows(ws, rows, withClub) {
  rows.forEach((p, i) => {
    const r = ws.addRow([]);
    xText(r.getCell(1), withClub ? p.clubName : p.name, { bold: true });
    xText(r.getCell(2), withClub ? p.name : (p.customDeal || dealLabel(p)), { mute: true });
    xText(r.getCell(3), withClub ? (p.customDeal || dealLabel(p)) : "", { mute: true });
    xMoney(r.getCell(4), p.pnl, { colorSign: false });
    xMoney(r.getCell(5), p.tips, { colorSign: false });
    xMoney(r.getCell(6), p.tipback, { colorSign: false });
    xMoney(r.getCell(7), p.settlement);
    xMoney(r.getCell(8), p.margin || 0);
    if (i % 2 === 1) for (let j = 1; j <= 8; j++) r.getCell(j).fill = fillOf(XLC.rowAlt);
  });
  const tr = ws.addRow([]);
  xText(tr.getCell(1), "TOTAL", { bold: true });
  xMoney(tr.getCell(7), rows.reduce((a, p) => a + p.settlement, 0), { bold: true });
  xMoney(tr.getCell(8), rows.reduce((a, p) => a + (p.margin || 0), 0), { bold: true });
  for (let j = 1; j <= 8; j++) tr.getCell(j).fill = fillOf(XLC.cream);
}
const AG_HEAD = ["Club", "Player", "Deal", "P&L", "Tips", "Tipback", "Settlement", "Your margin"];
async function downloadPlayerExcel(name, rows, wk) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(safeSheetName(name, wb));
  [18, 18, 22, 12, 12, 12, 13, 13].forEach((w, i) => (ws.getColumn(i + 1).width = w));
  xTitle(ws, `${name} — ${wk}`);
  xHeader(ws, AG_HEAD, 3);
  playerSheetRows(ws, rows, true);
  await saveWb(wb, `${name.replace(/[^\w]+/g, "_")}_${wk.replace(/[^\d]/g, "_").replace(/^_+|_+$/g, "") || "week"}.xlsx`);
}
async function downloadAgentWorkbook(model, wk) {
  const wb = new ExcelJS.Workbook();
  const active = model.clubs.filter((c) => c.active);
  const t = model.totals;
  const ws = wb.addWorksheet("Summary");
  [22, 13, 13, 15, 15, 13].forEach((w, i) => (ws.getColumn(i + 1).width = w));
  xTitle(ws, `My clubs — ${wk}`);
  const hr = xHeader(ws, ["Club", "Players P&L", "Tips", "Player settlements", "Club settlement", "Your margin"], 1);
  ws.views = [{ state: "frozen", ySplit: hr }];
  active.forEach((c, i) => {
    const r = ws.addRow([]);
    xText(r.getCell(1), c.name, { bold: true });
    xMoney(r.getCell(2), c.pnl, { colorSign: false });
    xMoney(r.getCell(3), c.tips, { colorSign: false });
    xMoney(r.getCell(4), c.settlements);
    xMoney(r.getCell(5), c.clubSettlement);
    xMoney(r.getCell(6), c.margin);
    if (i % 2 === 1) for (let j = 1; j <= 6; j++) r.getCell(j).fill = fillOf(XLC.rowAlt);
  });
  const gr = ws.addRow([]);
  xText(gr.getCell(1), "TOTAL", { bold: true, white: true });
  [t.pnl, t.tips, t.settlements, t.clubSettlements, t.margin].forEach((v, i) => xMoney(gr.getCell(i + 2), v, { bold: true, white: true, colorSign: false }));
  for (let j = 1; j <= 6; j++) gr.getCell(j).fill = fillOf(XLC.bar);
  ws.addRow([]);
  [["Cash net (club side − players + adjustments)", t.cashNet], ["TB margin", t.rakeMargin], ["TR margin", t.actionMargin]].forEach(([lbl, v]) => {
    const r = ws.addRow([]); xText(r.getCell(1), lbl, { bold: true }); xMoney(r.getCell(2), v, { bold: true });
  });
  model.umbrellas.filter((u) => u.played.length > 0).forEach((u) => {
    const w2 = wb.addWorksheet(safeSheetName(u.name, wb));
    [18, 18, 22, 12, 12, 12, 13, 13].forEach((w, i) => (w2.getColumn(i + 1).width = w));
    xTitle(w2, `${u.name} (umbrella) — ${wk}`);
    xHeader(w2, AG_HEAD, 3);
    playerSheetRows(w2, u.played, true);
  });
  active.forEach((c) => {
    const w2 = wb.addWorksheet(safeSheetName(c.name, wb));
    [18, 22, 12, 12, 12, 12, 13, 13].forEach((w, i) => (w2.getColumn(i + 1).width = w));
    xTitle(w2, `${c.name} — ${wk}`);
    xHeader(w2, ["Player", "Deal", "", "P&L", "Tips", "Tipback", "Settlement", "Your margin"], 3);
    playerSheetRows(w2, c.playersC.filter((p) => p.played), false);
    if (c.clubAdj !== 0) { const ar = w2.addRow([]); xText(ar.getCell(1), "Club adjustments", { mute: true }); xMoney(ar.getCell(7), c.clubAdj); }
    const cr = w2.addRow([]);
    xText(cr.getCell(1), "CLUB SETTLEMENT (you ↔ club)", { bold: true });
    xMoney(cr.getCell(7), c.clubSettlement, { bold: true });
  });
  await saveWb(wb, `MyClubs_${wk.replace(/[^\d]/g, "_").replace(/^_+|_+$/g, "") || "week"}.xlsx`);
}

// ════════════════════════════════════════════════════════════════
// DATA TASKS — configure read-only ClubGG collection jobs
// ════════════════════════════════════════════════════════════════

const DATA_TASKS_KEY = "clubgg-data-tasks-v1";
const TASK_FIELDS = [["hands", "Hands"], ["rake", "Rake"], ["pnl", "P&L"]];
const isoDate = (d = new Date()) => {
  const y = d.getFullYear(), m = String(d.getMonth() + 1).padStart(2, "0"), day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
};
const shiftedDate = (days) => { const d = new Date(); d.setDate(d.getDate() + days); return isoDate(d); };

async function loadDataTasks() {
  try {
    const saved = await store.get(DATA_TASKS_KEY);
    if (saved?.value) return JSON.parse(saved.value);
  } catch (e) {}
  return [];
}

function DataTasks() {
  const [agent, setAgent] = useState(AGENT_DEFAULT);
  const [tasks, setTasks] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [name, setName] = useState("");
  const [clubIds, setClubIds] = useState([]);
  const [playerIds, setPlayerIds] = useState([]);
  const [fields, setFields] = useState(TASK_FIELDS.map(([k]) => k));
  const [range, setRange] = useState("custom");
  const [from, setFrom] = useState(shiftedDate(-3));
  const [to, setTo] = useState(isoDate());
  const [note, setNote] = useState("");
  const [error, setError] = useState("");

  useEffect(() => {
    (async () => {
      try {
        const saved = await store.get("agentclubs-v3");
        if (saved?.value) setAgent(normalizeAgent({ ...AGENT_DEFAULT, ...JSON.parse(saved.value) }));
      } catch (e) {}
      setTasks(await loadDataTasks());
      setLoaded(true);
    })();
  }, []);

  const selectedClubs = agent.clubs.filter((c) => clubIds.includes(c.id));
  const availablePlayers = selectedClubs.flatMap((c) => c.players.map((p) => ({ ...p, clubId: c.id, clubName: c.name })));
  const selectedPlayers = availablePlayers.filter((p) => playerIds.includes(p.id));

  const chooseRange = (value) => {
    setRange(value);
    if (value === "today") { setFrom(isoDate()); setTo(isoDate()); }
    if (value === "past7") { setFrom(shiftedDate(-6)); setTo(isoDate()); }
  };
  const toggleClub = (id) => {
    const on = clubIds.includes(id);
    setClubIds(on ? clubIds.filter((x) => x !== id) : [...clubIds, id]);
    if (on) {
      const removed = new Set((agent.clubs.find((c) => c.id === id)?.players || []).map((p) => p.id));
      setPlayerIds(playerIds.filter((pid) => !removed.has(pid)));
    }
  };
  const saveTask = async () => {
    setError("");
    if (!clubIds.length) return setError("Select at least one club.");
    if (!playerIds.length) return setError("Select at least one player.");
    if (!fields.length) return setError("Select at least one data field.");
    if (!from || !to || from > to) return setError("Choose a valid date range.");
    const task = {
      id: uid(), name: name.trim() || `${selectedClubs.map((c) => c.name).join(", ")} · ${from} to ${to}`,
      status: "queued", createdAt: new Date().toISOString(), from, to, fields,
      clubs: selectedClubs.map((c) => ({ id: c.id, name: c.name })),
      players: selectedPlayers.map((p) => ({ id: p.id, name: p.name, clubId: p.clubId, clubName: p.clubName })),
      note: note.trim(),
    };
    const next = [task, ...tasks];
    setTasks(next);
    await store.set(DATA_TASKS_KEY, JSON.stringify(next));
    setName(""); setNote("");
  };

  if (!loaded) return <div style={{ padding: 40, color: C.mute }}>Loading data tasks…</div>;

  return (
    <div>
      <div style={{ padding: "16px 26px 14px", borderBottom: `2px solid ${C.line}`, background: C.paper }}>
        <div style={{ fontFamily: "Georgia, serif", fontSize: 20 }}>ClubGG data tasks</div>
        <div style={{ color: C.mute, fontSize: 12.5, marginTop: 3 }}>Choose exactly what Ledger Pilot should collect. Tasks are saved here ready for the ClubGG collector.</div>
      </div>
      <div style={{ padding: "20px 26px 60px", maxWidth: 1180, margin: "0 auto" }}>
        {error && <div style={{ background: "var(--errBg)", color: C.red, padding: "9px 13px", borderRadius: 6, marginBottom: 14, fontSize: 13 }}>{error}</div>}
        <div style={{ display: "grid", gridTemplateColumns: "minmax(0, 1.35fr) minmax(300px, .65fr)", gap: 14, alignItems: "start", marginBottom: 18 }}>
          <div>
            <div style={{ marginBottom: 14 }}><Card title="1 · Clubs">
              <div style={{ display: "flex", flexWrap: "wrap", gap: 7 }}>
                {agent.clubs.map((c) => {
                  const on = clubIds.includes(c.id);
                  return <button key={c.id} onClick={() => toggleClub(c.id)} style={{ border: `1px solid ${on ? C.goldDark : C.line}`, background: on ? C.gold : C.surface, color: on ? "var(--onGold)" : C.ink, borderRadius: 16, padding: "5px 11px", cursor: "pointer", fontSize: 12.5, fontWeight: on ? 700 : 500 }}>{c.name} <span style={{ opacity: .7 }}>· {c.players.length}</span></button>;
                })}
              </div>
            </Card></div>

            <div style={{ marginBottom: 14 }}><Card title="2 · Players" right={availablePlayers.length > 0 && <span style={{ display: "flex", gap: 7 }}><Btn tone="ghost" small onClick={() => setPlayerIds(availablePlayers.map((p) => p.id))}>Select all</Btn><Btn tone="ghost" small onClick={() => setPlayerIds([])}>Clear</Btn></span>}>
              {!clubIds.length && <div style={{ color: C.mute, fontSize: 13 }}>Select one or more clubs first.</div>}
              {!!clubIds.length && !availablePlayers.length && <div style={{ color: C.mute, fontSize: 13 }}>These clubs have no players configured yet. Add their rosters under My Clubs → Clubs & deals.</div>}
              {selectedClubs.map((club) => (
                <div key={club.id} style={{ marginBottom: 9 }}>
                  <div style={{ color: C.goldDark, fontSize: 11, fontWeight: 800, letterSpacing: ".07em", textTransform: "uppercase", marginBottom: 5 }}>{club.name}</div>
                  <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(190px, 1fr))", gap: 5 }}>
                    {club.players.map((p) => <label key={p.id} style={{ display: "flex", gap: 7, alignItems: "center", background: C.rowAlt, borderRadius: 6, padding: "6px 9px", fontSize: 12.5, cursor: "pointer" }}><input type="checkbox" checked={playerIds.includes(p.id)} onChange={() => setPlayerIds(playerIds.includes(p.id) ? playerIds.filter((x) => x !== p.id) : [...playerIds, p.id])} />{p.name}</label>)}
                  </div>
                </div>
              ))}
            </Card></div>

            <Card title="3 · Time period and data">
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 12 }}>
                {[['today','Today'],['past7','Past 7 days'],['custom','Custom']].map(([v,l]) => <button key={v} onClick={() => chooseRange(v)} style={{ border: `1px solid ${range === v ? C.goldDark : C.line}`, background: range === v ? C.gold : C.surface, color: range === v ? "var(--onGold)" : C.ink, borderRadius: 6, padding: "5px 10px", cursor: "pointer", fontSize: 12.5, fontWeight: 700 }}>{l}</button>)}
              </div>
              <div style={{ display: "flex", gap: 9, alignItems: "center", flexWrap: "wrap", marginBottom: 12 }}>
                <span style={{ color: C.mute, fontSize: 12 }}>From</span><input type="date" value={from} onChange={(e) => { setRange("custom"); setFrom(e.target.value); }} style={inputS} />
                <span style={{ color: C.mute, fontSize: 12 }}>to</span><input type="date" value={to} onChange={(e) => { setRange("custom"); setTo(e.target.value); }} style={inputS} />
              </div>
              <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>{TASK_FIELDS.map(([k, label]) => <label key={k} style={{ display: "flex", gap: 6, alignItems: "center", fontSize: 13, cursor: "pointer" }}><input type="checkbox" checked={fields.includes(k)} onChange={() => setFields(fields.includes(k) ? fields.filter((x) => x !== k) : [...fields, k])} />{label}</label>)}</div>
            </Card>
          </div>

          <div style={{ position: "sticky", top: 14 }}><Card title="Task summary">
            <input placeholder="Task name (optional)" value={name} onChange={(e) => setName(e.target.value)} style={{ ...inputS, width: "100%", boxSizing: "border-box", marginBottom: 10 }} />
            <div style={{ fontSize: 13, lineHeight: 1.8 }}>
              <div style={{ display: "flex" }}><span style={{ color: C.mute }}>Clubs</span><b style={{ marginLeft: "auto" }}>{clubIds.length}</b></div>
              <div style={{ display: "flex" }}><span style={{ color: C.mute }}>Players</span><b style={{ marginLeft: "auto" }}>{playerIds.length}</b></div>
              <div style={{ display: "flex" }}><span style={{ color: C.mute }}>Period</span><b style={{ marginLeft: "auto" }}>{from && to ? `${from} → ${to}` : "—"}</b></div>
              <div style={{ display: "flex" }}><span style={{ color: C.mute }}>Read</span><b style={{ marginLeft: "auto" }}>{fields.map((k) => TASK_FIELDS.find(([x]) => x === k)?.[1]).join(", ") || "—"}</b></div>
            </div>
            <textarea placeholder="Internal note (optional)" value={note} onChange={(e) => setNote(e.target.value)} rows={3} style={{ ...inputS, width: "100%", boxSizing: "border-box", resize: "vertical", margin: "10px 0" }} />
            <Btn tone="gold" onClick={saveTask}>Create collection task</Btn>
          </Card></div>
        </div>

        <Card title={`Collection queue · ${tasks.length}`}>
          {!tasks.length && <div style={{ color: C.mute, fontSize: 13 }}>No tasks yet. Create one above and it will appear here ready for the collector.</div>}
          {tasks.map((task) => <div key={task.id} style={{ display: "flex", gap: 12, alignItems: "center", borderTop: `1px solid ${C.line}`, padding: "10px 0", flexWrap: "wrap" }}>
            <div style={{ minWidth: 230, flex: 1 }}><div style={{ fontWeight: 700, fontSize: 14 }}>{task.name}</div><div style={{ color: C.mute, fontSize: 11.5 }}>{task.clubs.map((c) => c.name).join(", ")} · {task.players.length} players</div></div>
            <div style={{ fontSize: 12.5, color: C.mute }}>{task.from} → {task.to}</div>
            <Pill tone="gold">queued</Pill>
          </div>)}
        </Card>
      </div>
    </div>
  );
}

// ════════════════════════════════════════════════════════════════
// TABS — running ledger of who owes whom (players, clubs, other)
// ════════════════════════════════════════════════════════════════
// Convention everywhere: POSITIVE = they owe you · NEGATIVE = you owe them.
// Weekly settlements push in with one click; everything else is manual
// (crypto vig, sales, outside staking...). Entries flagged as P&L feed
// the side-income tracker.

const TABS_KEY = "tabs-v1";
const TABS_EMPTY = { counterparties: [], entries: [], pushed: {} };
const today = () => new Date().toISOString().slice(0, 10);

async function loadTabs() {
  try { const c = await store.get(TABS_KEY); if (c?.value) return { ...TABS_EMPTY, ...JSON.parse(c.value) }; } catch (e) {}
  return { ...TABS_EMPTY };
}
async function saveTabs(data) { try { await store.set(TABS_KEY, JSON.stringify(data)); } catch (e) {} }

function applyWeekToTabsData(data, sourceKey, weekLabel, items) {
  if (data.pushed[sourceKey]) return data;
  const byKey = {}; data.counterparties.forEach((cp) => (byKey[cp.kind + "|" + cp.name.toLowerCase()] = cp));
  const d = today();
  items.filter((it) => Math.abs(it.amount) > 0.005).forEach((it) => {
    const key = it.kind + "|" + it.name.toLowerCase();
    let cp = byKey[key];
    if (!cp) { cp = { id: uid(), name: it.name, kind: it.kind }; byKey[key] = cp; data.counterparties.push(cp); }
    data.entries.push({ id: uid(), date: d, cpId: cp.id, amount: Math.round(it.amount * 100) / 100, note: `${it.note} · ${weekLabel}`, source: "week", week: weekLabel });
  });
  data.pushed[sourceKey] = true;
  return data;
}

// Build the week's tab items from Fish Tank + My Clubs, with usernames folded
// into person bundles and clubs folded into their owners.
async function collectPendingWeeks() {
  const pending = [];
  let acfg = null;
  try { const c = await store.get("agentclubs-v3"); if (c?.value) acfg = normalizeAgent({ ...AGENT_DEFAULT, ...JSON.parse(c.value) }); } catch (e) {}
  const mapName = makeNameMapper(acfg?.personAliases);
  const ft = await loadFishTankModel();
  if (ft && ft.period) {
    const items = ft.model.entities.map((e) => ({ name: mapName(e.name), kind: "player", amount: -e.settlement, note: `Fish Tank · ${e.name}` }));
    pending.push({ sourceKey: `ft:${ft.period}`, label: `Fish Tank · ${ft.period}`, items });
  }
  if (acfg) {
    const weekKeys = Object.keys(acfg.weeks || {});
    const wk = acfg.currentWeek && acfg.weeks[acfg.currentWeek] ? acfg.currentWeek : weekKeys[weekKeys.length - 1];
    if (wk) {
      const m = computeAgent(acfg, acfg.weeks[wk]);
      const items = [];
      m.umbrellas.filter((u) => u.played.length > 0).forEach((u) => items.push({ name: u.name, kind: "player", amount: -u.settlement, note: `My Clubs umbrella` }));
      m.allPlayers.filter((p) => p.played && !m.inUmbrella.has(p.id)).forEach((p) => items.push({ name: mapName(p.name), kind: "player", amount: -p.settlement, note: `My Clubs · ${p.name} (${p.clubName})` }));
      m.clubs.filter((c) => c.active).forEach((c) => items.push({ name: (c.owner || "").trim() || c.name, kind: "club", amount: c.clubSettlement, note: `Club ${c.name}` }));
      pending.push({ sourceKey: `mc:${wk}`, label: `My Clubs · ${wk}`, items });
    }
  }
  return pending;
}

function TabsLedger() {
  const [data, setData] = useState(TABS_EMPTY);
  const [loaded, setLoaded] = useState(false);
  const [view, setView] = useState("balances");
  const [exportData, setExportData] = useState(null);
  const [filterCp, setFilterCp] = useState("");
  const [showZero, setShowZero] = useState(false);
  // new entry form
  const [fKind, setFKind] = useState("player");
  const [fName, setFName] = useState("");
  const [fAmount, setFAmount] = useState("");
  const [fNote, setFNote] = useState("");
  const [fDate, setFDate] = useState(today());
  const [fPnl, setFPnl] = useState(false);
  const [fCat, setFCat] = useState("");

  const [pending, setPending] = useState([]);
  useEffect(() => { (async () => { setData(await loadTabs()); setPending(await collectPendingWeeks()); setLoaded(true); })(); }, []);
  const save = async (next) => { setData(next); await saveTabs(next); };
  const acceptSources = async (sources) => {
    let next = { ...data, counterparties: [...data.counterparties], entries: [...data.entries], pushed: { ...data.pushed } };
    sources.forEach((s) => { next = applyWeekToTabsData(next, s.sourceKey, s.label, s.items); });
    await save(next);
  };

  const balances = useMemo(() => {
    const map = {};
    data.entries.forEach((e) => (map[e.cpId] = (map[e.cpId] || 0) + e.amount));
    return map;
  }, [data]);
  const cps = useMemo(() => [...data.counterparties].sort((a, b) => a.name.localeCompare(b.name)), [data]);
  const totalAll = Object.values(balances).reduce((a, v) => a + v, 0);
  const pnlEntries = data.entries.filter((e) => e.pnl);
  const pnlByCat = useMemo(() => {
    const m = {};
    pnlEntries.forEach((e) => { const k = e.category || "uncategorized"; m[k] = (m[k] || 0) + e.amount; });
    return m;
  }, [data]);
  const pnlTotal = pnlEntries.reduce((a, e) => a + e.amount, 0);

  const addEntry = async () => {
    const amt = parseFloat(String(fAmount).replace(/,/g, ""));
    if (!fName.trim() || isNaN(amt) || amt === 0) return;
    const key = fKind + "|" + fName.trim().toLowerCase();
    let cp = data.counterparties.find((c) => c.kind + "|" + c.name.toLowerCase() === key);
    const counterparties = cp ? data.counterparties : [...data.counterparties, (cp = { id: uid(), name: fName.trim(), kind: fKind })];
    const entry = { id: uid(), date: fDate || today(), cpId: cp.id, amount: Math.round(amt * 100) / 100, note: fNote.trim(), source: "manual", pnl: fPnl, category: fPnl ? (fCat.trim() || "uncategorized") : undefined };
    await save({ ...data, counterparties, entries: [...data.entries, entry] });
    setFName(""); setFAmount(""); setFNote(""); setFPnl(false); setFCat("");
  };
  const settleUp = async (cp) => {
    const bal = balances[cp.id] || 0;
    if (Math.abs(bal) < 0.005) return;
    if (!window.confirm(`Log a settling entry of ${fmt(-bal)} for ${cp.name}?`)) return;
    await save({ ...data, entries: [...data.entries, { id: uid(), date: today(), cpId: cp.id, amount: Math.round(-bal * 100) / 100, note: "Settled up", source: "manual" }] });
  };
  const delEntry = async (id) => await save({ ...data, entries: data.entries.filter((e) => e.id !== id) });

  const exportAll = () => {
    const cpName = (id) => cps.find((c) => c.id === id)?.name || "?";
    setExportData({ title: "Tabs ledger", text: toTSV(["Date", "Counterparty", "Kind", "Amount", "Note", "Source", "P&L category"],
      [...data.entries].reverse().map((e) => [e.date, cpName(e.cpId), cps.find((c) => c.id === e.cpId)?.kind || "", e.amount.toFixed(2), e.note || "", e.week || e.source, e.pnl ? (e.category || "uncategorized") : ""])) });
  };

  if (!loaded) return <div style={{ padding: 40, color: C.mute }}>Loading…</div>;

  const kindSection = (kind, title) => {
    const list = cps.filter((c) => c.kind === kind).filter((c) => showZero || Math.abs(balances[c.id] || 0) > 0.005);
    const subtotal = list.reduce((a, c) => a + (balances[c.id] || 0), 0);
    return (
      <Card title={`${title} · ${fmt(subtotal)}`}>
        {list.length === 0 && <div style={{ color: C.mute, fontSize: 13 }}>Nothing open.</div>}
        {list.map((c) => (
          <div key={c.id} style={{ display: "flex", alignItems: "center", gap: 8, borderTop: `1px solid ${C.line}`, padding: "7px 0", fontSize: 13.5 }}>
            <button onClick={() => { setFilterCp(c.id); setView("ledger"); }} style={{ border: "none", background: "none", cursor: "pointer", color: C.ink, fontWeight: 600, fontSize: 13.5, padding: 0, textAlign: "left" }}>{c.name}</button>
            <span style={{ marginLeft: "auto", fontVariantNumeric: "tabular-nums" }}>{money(balances[c.id] || 0)}</span>
            <button onClick={() => settleUp(c)} title="Log settling entry" style={{ border: `1px solid ${C.line}`, background: C.surface, color: C.mute, borderRadius: 5, cursor: "pointer", fontSize: 10.5, padding: "2px 8px" }}>settle</button>
          </div>
        ))}
      </Card>
    );
  };

  return (
    <div>
      <ExportModal data={exportData} onClose={() => setExportData(null)} />
      <div style={{ display: "flex", gap: 4, padding: "10px 26px 0", borderBottom: `2px solid ${C.line}`, background: C.paper, flexWrap: "wrap", alignItems: "center" }}>
        {[["balances", "Balances"], ["ledger", "Ledger"]].map(([k, label]) => (
          <button key={k} onClick={() => setView(k)} style={{
            border: "none", cursor: "pointer", padding: "9px 16px", fontSize: 13.5, fontWeight: 700,
            background: view === k ? C.card : "transparent", color: view === k ? C.ink : C.mute,
            borderRadius: "8px 8px 0 0", marginBottom: -2 }}>
            {label}
          </button>
        ))}
        <div style={{ marginLeft: "auto", display: "flex", gap: 8, alignItems: "center", paddingBottom: 6 }}>
          <span style={{ fontSize: 13, color: C.mute }}>Net position: <b style={{ color: totalAll >= 0 ? C.green : C.red }}>{fmt(totalAll)}</b></span>
          <Btn tone="gold" small onClick={() => downloadTabsExcel(data, cps, balances, pnlByCat, pnlTotal)}>Download Excel</Btn>
          <Btn tone="ghost" small onClick={exportAll}>Copy</Btn>
        </div>
      </div>
      <div style={{ padding: "20px 26px 60px", maxWidth: 1180, margin: "0 auto" }}>
        <div style={{ color: C.mute, fontSize: 12.5, marginBottom: 14 }}>
          <b style={{ color: C.green }}>Positive = they owe you</b> · <b style={{ color: C.red }}>negative = you owe them</b>. Finished weeks appear below for review — Accept folds them into the balances (once per week, no double-counting). Everything else is manual.
        </div>

        {(() => {
          const open = pending.filter((s) => !data.pushed[s.sourceKey] && s.items.some((it) => Math.abs(it.amount) > 0.005));
          if (!open.length) return null;
          return (
            <div style={{ marginBottom: 16 }}>
              <Card title="Weekly settlements ready to fold in" right={
                open.length > 1 ? <Btn tone="gold" small onClick={() => acceptSources(open)}>Accept all</Btn> : null
              }>
                {open.map((s) => {
                  const agg = {};
                  s.items.filter((it) => Math.abs(it.amount) > 0.005).forEach((it) => {
                    const k = it.kind + "|" + it.name.toLowerCase();
                    if (!agg[k]) agg[k] = { name: it.name, kind: it.kind, amount: 0, n: 0 };
                    agg[k].amount += it.amount; agg[k].n += 1;
                  });
                  const rows = Object.values(agg).sort((a, b) => Math.abs(b.amount) - Math.abs(a.amount));
                  const tot = rows.reduce((a, r2) => a + r2.amount, 0);
                  return (
                    <div key={s.sourceKey} style={{ borderTop: `1px solid ${C.line}`, padding: "10px 0" }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 6 }}>
                        <b style={{ fontSize: 13.5 }}>{s.label}</b>
                        <span style={{ color: C.mute, fontSize: 12 }}>{rows.length} counterparties · net {fmt(tot)}</span>
                        <span style={{ marginLeft: "auto" }}><Btn tone="gold" small onClick={() => acceptSources([s])}>Accept — update tabs</Btn></span>
                      </div>
                      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(230px, 1fr))", gap: "3px 18px" }}>
                        {rows.map((r2) => (
                          <div key={r2.kind + r2.name} style={{ display: "flex", fontSize: 12.5, padding: "2px 0" }}>
                            <span style={{ color: C.ink, fontWeight: 600 }}>{r2.name}</span>
                            <span style={{ color: C.mute, fontSize: 10.5, marginLeft: 5, alignSelf: "center" }}>{r2.kind}{r2.n > 1 ? ` ·${r2.n}` : ""}</span>
                            <span style={{ marginLeft: "auto", fontVariantNumeric: "tabular-nums" }}>{money(r2.amount)}</span>
                          </div>
                        ))}
                      </div>
                    </div>
                  );
                })}
              </Card>
            </div>
          );
        })()}

        {/* Add entry */}
        <div style={{ marginBottom: 16 }}>
          <Card title="Add entry">
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
              <select value={fKind} onChange={(e) => setFKind(e.target.value)} style={{ ...inputS, fontSize: 12.5 }}>
                <option value="player">Player</option><option value="club">Club</option><option value="other">Other</option>
              </select>
              <input list="tabs-cp-names" placeholder="Who…" value={fName} onChange={(e) => setFName(e.target.value)} style={{ ...inputS, width: 170 }} />
              <datalist id="tabs-cp-names">{cps.filter((c) => c.kind === fKind).map((c) => <option key={c.id} value={c.name} />)}</datalist>
              <NumInput width={100} value={fAmount} onChange={(v) => setFAmount(v)} />
              <input placeholder="Note (crypto vig, sold X, stake...)" value={fNote} onChange={(e) => setFNote(e.target.value)} style={{ ...inputS, flex: 1, minWidth: 180 }} />
              <input value={fDate} onChange={(e) => setFDate(e.target.value)} style={{ ...inputS, width: 100 }} />
              <label style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 12.5, color: C.mute }}>
                <input type="checkbox" checked={fPnl} onChange={(e) => setFPnl(e.target.checked)} /> count as P&L
              </label>
              {fPnl && <input placeholder="Category (vig, sales, staking…)" value={fCat} onChange={(e) => setFCat(e.target.value)} style={{ ...inputS, width: 170 }} />}
              <Btn tone="gold" small onClick={addEntry}>Add</Btn>
            </div>
          </Card>
        </div>

        {view === "balances" && (
          <>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(280px, 1fr))", gap: 14, marginBottom: 16 }}>
              {kindSection("player", "Player tabs")}
              {kindSection("club", "Club tabs")}
              {kindSection("other", "Other")}
            </div>
            <div style={{ display: "flex", gap: 14, alignItems: "flex-start", flexWrap: "wrap" }}>
              <div style={{ flex: 1, minWidth: 280 }}>
                <Card title={`Side P&L · ${fmt(pnlTotal)}`}>
                  {Object.keys(pnlByCat).length === 0 && <div style={{ color: C.mute, fontSize: 13 }}>Flag manual entries as P&L (vig, sales, outside staking) and their running totals show here by category.</div>}
                  {Object.entries(pnlByCat).sort((a, b) => Math.abs(b[1]) - Math.abs(a[1])).map(([cat, v]) => (
                    <div key={cat} style={{ display: "flex", padding: "6px 0", borderTop: `1px solid ${C.line}`, fontSize: 13.5 }}>
                      <span style={{ color: C.mute }}>{cat}</span>
                      <span style={{ marginLeft: "auto", fontVariantNumeric: "tabular-nums" }}>{money(v)}</span>
                    </div>
                  ))}
                </Card>
              </div>
              <label style={{ fontSize: 12.5, color: C.mute, display: "flex", gap: 6, alignItems: "center" }}>
                <input type="checkbox" checked={showZero} onChange={(e) => setShowZero(e.target.checked)} /> show zero balances
              </label>
            </div>
          </>
        )}

        {view === "ledger" && (
          <div style={{ background: C.card, borderRadius: 10, overflow: "hidden", boxShadow: "0 1px 5px rgba(0,0,0,0.12)" }}>
            <div style={{ display: "flex", gap: 10, alignItems: "center", padding: "10px 16px", background: C.cream }}>
              <span style={{ fontWeight: 700 }}>Ledger</span>
              <select value={filterCp} onChange={(e) => setFilterCp(e.target.value)} style={{ ...inputS, fontSize: 12.5 }}>
                <option value="">Everyone</option>
                {cps.map((c) => <option key={c.id} value={c.id}>{c.name} ({c.kind})</option>)}
              </select>
              {filterCp && <span style={{ fontSize: 13 }}>balance {money(balances[filterCp] || 0)}</span>}
            </div>
            <table style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead><tr style={{ background: C.cream }}>
                <th style={{ ...th, textAlign: "left" }}>Date</th><th style={{ ...th, textAlign: "left" }}>Who</th>
                <th style={th}>Amount</th><th style={{ ...th, textAlign: "left" }}>Note</th>
                <th style={{ ...th, textAlign: "left" }}>Source</th><th style={th}></th>
              </tr></thead>
              <tbody>
                {[...data.entries].reverse().filter((e) => !filterCp || e.cpId === filterCp).map((e, i) => {
                  const cp = cps.find((c) => c.id === e.cpId);
                  return (
                    <tr key={e.id} style={{ background: i % 2 ? C.rowAlt : C.card, borderTop: `1px solid ${C.line}` }}>
                      <td style={{ ...tdL, color: C.mute, fontSize: 12 }}>{e.date}</td>
                      <td style={{ ...tdL, fontWeight: 600 }}>{cp?.name || "?"} <span style={{ color: C.mute, fontWeight: 400, fontSize: 11 }}>({cp?.kind})</span></td>
                      <td style={td}>{money(e.amount)}</td>
                      <td style={{ ...tdL, fontSize: 12.5 }}>{e.note || ""}{e.pnl && <span style={{ marginLeft: 6 }}><Pill tone="green">P&L · {e.category}</Pill></span>}</td>
                      <td style={{ ...tdL, color: C.mute, fontSize: 11.5 }}>{e.source === "week" ? e.week : "manual"}</td>
                      <td style={{ ...td, width: 36 }}><button onClick={() => delEntry(e.id)} style={{ border: "none", background: "none", color: C.red, cursor: "pointer", fontSize: 14 }}>×</button></td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}

async function downloadTabsExcel(data, cps, balances, pnlByCat, pnlTotal) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Balances");
  [22, 10, 14].forEach((w, i) => (ws.getColumn(i + 1).width = w));
  xTitle(ws, `Tabs — balances as of ${today()}`);
  [["player", "Player tabs"], ["club", "Club tabs"], ["other", "Other"]].forEach(([kind, title]) => {
    const hr = ws.addRow([title]);
    hr.getCell(1).font = { name: "Arial", size: 11, bold: true, color: { argb: XLC.white } };
    for (let j = 1; j <= 3; j++) hr.getCell(j).fill = fillOf(XLC.gold);
    const list = cps.filter((c) => c.kind === kind && Math.abs(balances[c.id] || 0) > 0.005);
    list.forEach((c, i) => {
      const r = ws.addRow([]);
      xText(r.getCell(1), c.name, { bold: true });
      xMoney(r.getCell(3), balances[c.id] || 0);
      if (i % 2 === 1) for (let j = 1; j <= 3; j++) r.getCell(j).fill = fillOf(XLC.rowAlt);
    });
    const sub = ws.addRow([]);
    xText(sub.getCell(1), "Subtotal", { bold: true });
    xMoney(sub.getCell(3), list.reduce((a, c) => a + (balances[c.id] || 0), 0), { bold: true });
    for (let j = 1; j <= 3; j++) sub.getCell(j).fill = fillOf(XLC.cream);
    ws.addRow([]);
  });
  const tot = ws.addRow([]);
  xText(tot.getCell(1), "NET POSITION", { bold: true, white: true });
  xMoney(tot.getCell(3), Object.values(balances).reduce((a, v) => a + v, 0), { bold: true, white: true, colorSign: false });
  for (let j = 1; j <= 3; j++) tot.getCell(j).fill = fillOf(XLC.bar);

  const wl = wb.addWorksheet("Ledger");
  [11, 20, 9, 13, 34, 16, 16].forEach((w, i) => (wl.getColumn(i + 1).width = w));
  xTitle(wl, "Ledger — full history");
  xHeader(wl, ["Date", "Counterparty", "Kind", "Amount", "Note", "Source", "P&L category"], 3);
  [...data.entries].reverse().forEach((e, i) => {
    const cp = cps.find((c) => c.id === e.cpId);
    const r = wl.addRow([]);
    xText(r.getCell(1), e.date, { mute: true });
    xText(r.getCell(2), cp?.name || "?", { bold: true });
    xText(r.getCell(3), cp?.kind || "", { mute: true });
    xMoney(r.getCell(4), e.amount);
    xText(r.getCell(5), e.note || "", { mute: true });
    xText(r.getCell(6), e.source === "week" ? e.week || "week" : "manual", { mute: true });
    xText(r.getCell(7), e.pnl ? (e.category || "uncategorized") : "", { mute: true });
    if (i % 2 === 1) for (let j = 1; j <= 7; j++) r.getCell(j).fill = fillOf(XLC.rowAlt);
  });

  if (Object.keys(pnlByCat).length) {
    const wp = wb.addWorksheet("Side P&L");
    [22, 14].forEach((w, i) => (wp.getColumn(i + 1).width = w));
    xTitle(wp, "Side P&L by category");
    xHeader(wp, ["Category", "Total"], 1);
    Object.entries(pnlByCat).forEach(([cat, v], i) => {
      const r = wp.addRow([]);
      xText(r.getCell(1), cat, { bold: true });
      xMoney(r.getCell(2), v);
      if (i % 2 === 1) for (let j = 1; j <= 2; j++) r.getCell(j).fill = fillOf(XLC.rowAlt);
    });
    const tr = wp.addRow([]);
    xText(tr.getCell(1), "TOTAL", { bold: true });
    xMoney(tr.getCell(2), pnlTotal, { bold: true });
    for (let j = 1; j <= 2; j++) tr.getCell(j).fill = fillOf(XLC.cream);
  }
  await saveWb(wb, `Tabs_${today()}.xlsx`);
}
