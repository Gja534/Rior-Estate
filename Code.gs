/**
 * Rior Estate – backend në Google Apps Script.
 * Të dhënat jetojnë në Google Sheets, fotot në Google Drive, hyrja bëhet me llogari Google.
 * Claude përdoret vetëm për ta shkruar/ndryshuar këtë kod.
 */
const TAB = {set: "Cilësimet", ag: "Agjentët", pr: "Prona", msg: "Mesazhet", sum: "Përmbledhje"};
const HEAD = ["Titulli", "Qyteti", "Lloji", "Shitje/Qera", "Çmimi (€)", "m²", "Përshkrimi", "Foto (link Google Drive)", "Hiq? (PO)", "Statusi"];
const CITIES = ["Tiranë", "Durrës", "Vlorë", "Sarandë", "Shkodër", "Korçë", "Elbasan", "Fier", "Gjirokastër", "Berat"];
const SS = () => SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty("SS"));
const say = m => { try { SpreadsheetApp.getUi().alert(m); } catch (e) { Logger.log(m); } };

function onOpen() {
  SpreadsheetApp.getUi().createMenu("Rior Estate")
    .addItem("Krijo tabelat e agjentëve të rinj", "createAgents")
    .addItem("Përditëso përmbledhjen", "refreshSummary")
    .addItem("Rifresko faqen e internetit", "clearCache")
    .addToUi();
}

/** Ekzekutohet një herë: krijon tabelat e nevojshme. */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty("SS", ss.getId());
  const tab = n => ss.getSheetByName(n) || ss.insertSheet(n);
  const set = tab(TAB.set);
  if (set.getLastRow() < 1) set.getRange(1, 1, 8, 2).setValues([
    ["Çelësi", "Vlera"], ["phone", "+355 69 000 0000"], ["email", "info@rior-estate.al"],
    ["addr", "Rruga e Kavajës 45, Tiranë"], ["hours", "E hënë – e shtunë, 09:00 – 18:00"],
    ["wa", "355690000000"],
    ["lead", "Apartamente, vila dhe toka në Tiranë, Durrës, Vlorë, Sarandë dhe Shkodër. Çdo pronë kalon kontrollin e dokumenteve para se ta shihni."],
    ["notify", Session.getEffectiveUser().getEmail()]]);
  const ag = tab(TAB.ag);
  if (ag.getLastRow() < 1) {
    ag.getRange(1, 1, 1, 7).setValues([["Emri", "Roli", "Telefoni", "Email Google", "Aktiv (PO/JO)", "ID e tabelës", "Linku"]]);
    ag.setFrozenRows(1); ag.getRange(1, 1, 1, 7).setFontWeight("bold");
  }
  const pr = tab(TAB.pr);
  if (pr.getLastRow() < 1) {
    pr.getRange(1, 1, 1, 11).setValues([HEAD.concat(["Agjenti (emri)"])]);
    pr.getRange(2, 1, 1, 10).setValues([["Shembull: Apartament 2+1 pranë Liqenit", "Tiranë", "Apartament", "Shitje", 135000, 92, "2 dhoma, ballkon, ashensor", "", "", "Fshehur"]]);
  }
  format(pr);
  const ms = tab(TAB.msg);
  if (ms.getLastRow() < 1) ms.getRange(1, 1, 1, 4).setValues([["Data", "Emri", "Kontakti", "Mesazhi"]]);
  tab(TAB.sum);
  ["Sheet1", "Sheet 1", "Faqe1"].forEach(n => {
    const s = ss.getSheetByName(n);
    if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });
  refreshSummary();
  say("Gati. Tani bëni Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone).");
}

function format(sh) {
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 10)).setFontWeight("bold").setBackground("#111111").setFontColor("#ffffff");
  const dv = l => SpreadsheetApp.newDataValidation().requireValueInList(l, true).setAllowInvalid(false).build();
  sh.getRange("B2:B1000").setDataValidation(dv(CITIES));
  sh.getRange("C2:C1000").setDataValidation(dv(["Apartament", "Vilë", "Tokë", "Lokal"]));
  sh.getRange("D2:D1000").setDataValidation(dv(["Shitje", "Qera"]));
  sh.getRange("I2:I1000").setDataValidation(dv(["PO"]));
  sh.getRange("J2:J1000").setDataValidation(dv(["Publikuar", "Fshehur"]));
  sh.setColumnWidth(1, 280); sh.setColumnWidth(7, 280); sh.setColumnWidth(8, 240);
}

function agentExtras(sh) {
  sh.getRange("L1:M4").setValues([
    ["Gjithsej prona", "=COUNTA(A2:A)"],
    ["Të publikuara", '=COUNTIFS(J2:J,"Publikuar",I2:I,"<>PO")'],
    ["Në pritje të miratimit", '=COUNTA(A2:A)-COUNTIF(J2:J,"Publikuar")-COUNTIF(J2:J,"Fshehur")'],
    ["Kërkuar për heqje", '=COUNTIF(I2:I,"PO")']]);
  sh.getRange("L1:L4").setFontWeight("bold");
  sh.getRange("L6").setValue("Si funksionon: shkruani çdo pronë në një rresht. Kolona Statusi (J) vendoset nga administratori; kur bëhet \"Publikuar\" prona del në faqe. Për ta hequr një pronë, zgjidhni PO te kolona I. Për foto: ngarkojeni në Google Drive, vendosni \"Kushdo me linkun\" dhe ngjitni linkun në kolonën H.");
  sh.getRange("L6").setWrap(true); sh.setColumnWidth(12, 260);
}

/** Krijon tabelën personale për çdo agjent të ri në listën "Agjentët" dhe ia ndan me email. */
function createAgents() {
  const ss = SS(), s = ss.getSheetByName(TAB.ag), v = s.getDataRange().getValues(), me = Session.getEffectiveUser().getEmail();
  let n = 0;
  for (let i = 1; i < v.length; i++) {
    const name = v[i][0], email = String(v[i][3] || "").trim().toLowerCase(), id = v[i][5];
    if (!name || !email || id) continue;
    const nb = SpreadsheetApp.create("Prona – " + name), sh = nb.getSheets()[0];
    sh.setName("Pronat e mia");
    sh.getRange(1, 1, 1, 10).setValues([HEAD]);
    format(sh); agentExtras(sh);
    const p = sh.getRange("J2:J1000").protect().setDescription("Vetëm administratori");
    p.removeEditors(p.getEditors());
    if (p.canDomainEdit()) p.setDomainEdit(false);
    p.addEditor(me);
    DriveApp.getFileById(nb.getId()).setShareableByEditors(false);
    nb.addEditor(email);
    s.getRange(i + 1, 6, 1, 2).setValues([[nb.getId(), nb.getUrl()]]);
    if (!v[i][4]) s.getRange(i + 1, 5).setValue("PO");
    n++;
  }
  say(n ? n + " tabela u krijuan. Agjentët morën ftesë me email nga Google." : "Nuk ka agjentë të rinj. Shtoni një rresht me emrin dhe emailin Google te tabela Agjentët.");
  refreshSummary(); clearCache();
}

function refreshSummary() {
  const ss = SS(), s = ss.getSheetByName(TAB.sum);
  const rows = [["Agjenti", "Gjithsej", "Të publikuara", "Në pritje", "Kërkuar për heqje", "Tabela"]];
  ss.getSheetByName(TAB.ag).getDataRange().getValues().slice(1).forEach(r => {
    if (!r[0] || !r[5]) return;
    try {
      const v = SpreadsheetApp.openById(String(r[5])).getSheets()[0].getDataRange().getValues().slice(1).filter(x => x[0]);
      const rem = x => String(x[8]).toUpperCase() === "PO";
      rows.push([r[0], v.length, v.filter(x => x[9] === "Publikuar" && !rem(x)).length,
        v.filter(x => x[9] !== "Publikuar" && x[9] !== "Fshehur").length, v.filter(rem).length, r[6]]);
    } catch (e) { rows.push([r[0], "gabim", "", "", "", r[6]]); }
  });
  s.clear();
  s.getRange(1, 1, rows.length, 6).setValues(rows);
  s.getRange(1, 1, 1, 6).setFontWeight("bold");
}

function clearCache() { CacheService.getScriptCache().remove("data"); }

function driveImg(u) {
  u = String(u || "").trim();
  if (!u) return "";
  const m = u.match(/[-\w]{25,}/);
  if (/drive\.google\.com|docs\.google\.com/.test(u) && m) return "https://drive.google.com/thumbnail?id=" + m[0] + "&sz=w1000";
  return u.indexOf("https://") === 0 ? u : "";
}

function readProps(sh, by, admin) {
  const out = [], v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    const r = v[i];
    if (!r[0] || r[9] !== "Publikuar" || String(r[8]).toUpperCase() === "PO") continue;
    out.push({
      id: sh.getParent().getId().slice(-6) + "-" + (i + 1), t: String(r[0]), q: String(r[1]), l: String(r[2]), v: String(r[3]),
      c: Number(r[4]) || 0, m: Number(r[5]) || 0, d: String(r[6] || ""), img: driveImg(r[7]), by: by || (admin ? String(r[10] || "") : "")
    });
  }
  return out;
}

/** Thirret nga faqja publike. */
function getData() {
  const cache = CacheService.getScriptCache(), hit = cache.get("data");
  if (hit) return JSON.parse(hit);
  const ss = SS(), k = {};
  ss.getSheetByName(TAB.set).getDataRange().getValues().slice(1).forEach(r => { if (r[0]) k[String(r[0])] = String(r[1]); });
  delete k.notify;
  const props = readProps(ss.getSheetByName(TAB.pr), "", true), agents = [];
  ss.getSheetByName(TAB.ag).getDataRange().getValues().slice(1).forEach(r => {
    if (!r[0] || String(r[4]).toUpperCase() === "JO") return;
    agents.push({n: String(r[0]), r: String(r[1] || ""), ph: String(r[2] || "")});
    if (r[5]) { try { props.push.apply(props, readProps(SpreadsheetApp.openById(String(r[5])).getSheets()[0], String(r[0]), false)); } catch (e) {} }
  });
  const out = {k: k, agents: agents, properties: props};
  try { cache.put("data", JSON.stringify(out), 300); } catch (e) {}
  return out;
}

/** Forma e kontaktit: ruan mesazhin në tabelën Mesazhet dhe dërgon email. */
function sendMessage(name, contact, msg) {
  name = String(name || "").slice(0, 120); contact = String(contact || "").slice(0, 160); msg = String(msg || "").slice(0, 2000);
  if (!name || !contact) throw new Error("Plotësoni emrin dhe kontaktin.");
  const c = CacheService.getScriptCache(), n = Number(c.get("m") || 0);
  if (n >= 20) throw new Error("Shumë mesazhe. Provoni më vonë.");
  c.put("m", n + 1, 600);
  const ss = SS(), k = {};
  ss.getSheetByName(TAB.set).getDataRange().getValues().slice(1).forEach(r => { if (r[0]) k[r[0]] = r[1]; });
  ss.getSheetByName(TAB.msg).appendRow([new Date(), name, contact, msg]);
  try { MailApp.sendEmail(k.notify || ss.getOwner().getEmail(), "Mesazh i ri nga faqja – " + name, "Emri: " + name + "\nKontakti: " + contact + "\n\n" + msg); } catch (e) {}
  return true;
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("Rior Estate – Prona premium në Shqipëri")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
