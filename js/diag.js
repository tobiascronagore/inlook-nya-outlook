// Förprov: vilka Graph-behörigheter får användaren själv godkänna i tenanten,
// och fungerar NAA (inloggning utan server) i den här Outlook-värden?
(function () {
  var $ = function (id) { return document.getElementById(id); };
  var lines = [];

  function log(s) { lines.push(s); $("report").textContent = lines.join("\n"); }
  function row(name, text, ok) {
    var tr = document.createElement("tr");
    tr.innerHTML = "<td></td><td></td>";
    tr.children[0].textContent = name; tr.children[1].textContent = text;
    tr.children[1].style.color = ok ? "var(--ok)" : "var(--err)";
    if ($("out") && $("out").tBodies[0]) $("out").tBodies[0].appendChild(tr);
    log(name + ": " + text);
  }

  // Tolkar felet till "vem måste godkänna".
  function explain(e) {
    var m = (e && (e.errorMessage || e.message)) || String(e);
    var code = ((m.match(/AADSTS\d+/) || [])[0]) || (e && e.errorCode) || "";
    var admin = InlookAuth.isAdminErr(e);
    var cancelled = /user_cancelled|AADSTS65004|popup_window_error|empty_window/i.test(m + (e && e.errorCode || ""));
    var kind = admin ? "KRÄVER ADMINISTRATÖR" : /AADSTS65001|consent_required/i.test(m) ? "medgivande saknas (ej bekräftat admin)" : cancelled ? "avbröts/popup blockerad" : "fel";
    return kind + " (" + (code || (e && e.errorCode) || "okänd kod") + ")";
  }

  var G = function (scopes, method, path, opts) { return InlookGraph.call(scopes, method, path, opts); };

  // Graph-test: ett läsanrop som bara visar statuskod (inget innehåll).
  function gt(n, scope, path) {
    return { n: n, s: [scope], g: "low", f: async function () { return (await G([scope], "GET", path)).status; } };
  }
  // Resurs-test (Exchange): bara att token kan fås; visar resurs (aud) och scp, aldrig token.
  function rt(n, scope) {
    return { n: n, s: [scope], g: "ex", f: async function () {
      await InlookAuth.getToken([scope]);
      var inf = InlookAuth.info();
      return "token för resurs " + (inf.aud || "?") + ", scp=" + (inf.scp || "?");
    } };
  }

  // Varje test: namn, scope, grupp (base/low/ex) och en funktion som visar att token faktiskt fungerar.
  var tests = [
    { n: "User.Read", s: ["User.Read"], g: "base", f: async function () { return (await G(["User.Read"], "GET", "/me?$select=id")).status; } },
    { n: "Files.ReadWrite (OneDrive)", s: ["Files.ReadWrite"], g: "base", f: async function () {
        await InlookGraph.writeJson("_diag.json", { test: true }, null);
        await InlookGraph.deleteFile("_diag.json");
        return "skriv+radera OK";
      } },
    { n: "Notes.Read", s: ["Notes.Read"], g: "base", f: async function () { return (await G(["Notes.Read"], "GET", "/me/onenote/notebooks?$top=1&$select=id")).status; } },
    { n: "Notes.ReadWrite", s: ["Notes.ReadWrite"], g: "base", f: async function () { return (await G(["Notes.ReadWrite"], "GET", "/me/onenote/notebooks?$top=1&$select=id")).status; } },
    // Lägre Graph-behörigheter (kan vara klassade som låg påverkan och medgivbara av användaren).
    gt("Mail.ReadBasic", "Mail.ReadBasic", "/me/messages?$top=1&$select=id,receivedDateTime"),
    gt("Calendars.ReadBasic", "Calendars.ReadBasic", "/me/events?$top=1&$select=id,start"),
    gt("Calendars.Read.Shared", "Calendars.Read.Shared", "/me/calendars?$top=1&$select=id"),
    gt("Tasks.Read (To Do, läs)", "Tasks.Read", "/me/todo/lists?$top=1&$select=id"),
    gt("MailboxSettings.Read", "MailboxSettings.Read", "/me/mailboxSettings/timeZone"),
    gt("People.Read", "People.Read", "/me/people?$top=1&$select=id"),
    gt("Contacts.Read", "Contacts.Read", "/me/contacts?$top=1&$select=id"),
    // Exchange-resursscope (annan resurs än Graph). Bara tokenfåendet prövas, ingen brevlåda rörs.
    rt("EWS.AccessAsUser.All", "https://outlook.office.com/EWS.AccessAsUser.All"),
    rt("IMAP.AccessAsUser.All", "https://outlook.office.com/IMAP.AccessAsUser.All"),
    rt("SMTP.Send", "https://outlook.office.com/SMTP.Send")
  ];

  // Office.js-möjligheter i värden (ingen inloggning). Visar bara ja/nej och längd, aldrig innehåll.
  async function officeChecks() {
    if (!(window.Office && Office.context && Office.context.mailbox)) { log("Office.js-test: ej i Outlook, hoppas över"); return; }
    var rq = Office.context.requirements, mb = Office.context.mailbox, it = mb.item;
    [["1.3", "läsa brödtext (body.getAsync)"], ["1.8", "NAA-krav"], ["1.10", "händelsestyrd aktivering"], ["1.14", "nyare API"]].forEach(function (p) {
      var ok = rq.isSetSupported("Mailbox", p[0]);
      row("Mailbox " + p[0], p[1] + ": " + (ok ? "stöds" : "stöds ej"), ok);
    });
    var f1 = typeof mb.displayNewMessageForm === "function", f2 = typeof mb.displayNewAppointmentForm === "function";
    row("displayNewMessageForm", f1 ? "finns" : "saknas", f1);
    row("displayNewAppointmentForm", f2 ? "finns" : "saknas", f2);
    if (!it) { log("Inget öppet mejl: markera ett mejl och kör igen för test av brödtext/svarsformulär"); return; }
    var f3 = typeof it.displayReplyForm === "function";
    row("item.displayReplyForm", f3 ? "finns" : "saknas", f3);
    if (it.body && it.body.getAsync) {
      await new Promise(function (ok) {
        var h = setTimeout(function () { row("Läsa brödtext (öppet mejl)", "ingen respons inom 10 s", false); ok(); }, 10000);
        it.body.getAsync("text", function (r) {
          clearTimeout(h);
          var good = r.status === Office.AsyncResultStatus.Succeeded;
          row("Läsa brödtext (öppet mejl)", good ? "OK, " + String(r.value || "").length + " tecken" : "fel: " + (r.error && r.error.message), good);
          ok();
        });
      });
    } else { row("Läsa brödtext (öppet mejl)", "body.getAsync saknas (skrivläge/möte?)", false); }
  }

  // iCloud-ICS direkt från panelen: prövar bara om webbläsaren får läsa (CORS). Länken sparas aldrig.
  async function icsTest() {
    var u = ($("icsurl") && $("icsurl").value || "").trim().replace(/^webcal:/i, "https:");
    if (!/^https:\/\//i.test(u)) { row("iCloud-ICS (CORS)", "ange en https/webcal-länk först", false); return; }
    try {
      var r = await fetch(u, { method: "GET", mode: "cors", credentials: "omit" });
      var txt = await r.text();
      var isIcs = txt.indexOf("BEGIN:VCALENDAR") >= 0;
      row("iCloud-ICS (CORS)", "läsbar: HTTP " + r.status + ", " + (isIcs ? "ICS, " + (txt.match(/BEGIN:VEVENT/g) || []).length + " händelser" : "inte ICS"), r.ok && isIcs);
    } catch (e) {
      row("iCloud-ICS (CORS)", "blockerad av webbläsaren (CORS) eller nätfel: " + (e && e.message), false);
    }
  }

  function setDis(v) { ["run", "run1", "runLow", "runEx", "runOff", "runIcs"].forEach(function (i) { if ($(i)) $(i).disabled = v; }); }

  async function run(grp) {
    setDis(true);
    if ($("out") && $("out").tBodies[0]) $("out").tBodies[0].innerHTML = "";
    lines = [];
    try {
      log("Datum: " + new Date().toISOString().slice(0, 10));
      var inOffice = !!(window.Office && Office.context && Office.context.mailbox);
      log("Office-värd: " + (inOffice ? ((Office.context.mailbox.diagnostics || {}).hostName || "okänd") + " " + ((Office.context.mailbox.diagnostics || {}).hostVersion || "") : "ingen (vanlig webbläsare)"));
      if (inOffice && Office.context.requirements) {
        log("NestedAppAuth 1.1 stöds: " + Office.context.requirements.isSetSupported("NestedAppAuth", "1.1"));
        log("Mailbox 1.8 stöds: " + Office.context.requirements.isSetSupported("Mailbox", "1.8"));
        var it = Office.context.mailbox.item;
        log("Mejl-id tillgängligt: " + !!(it && it.itemId));
      }
      log("Dialog API stöds: " + !!(window.Office && Office.context && Office.context.ui && Office.context.ui.displayDialogAsync));
      var list = grp === "base1" ? tests.slice(0, 1) : grp === "off" ? [] : !grp ? tests : tests.filter(function (t) { return t.g === grp; });
      if (grp === "off" || !grp) await officeChecks();
      for (var i = 0; i < list.length; i++) {
        var t = list[i];
        try {
          // Varje test har egen tidsgräns så att en hängd inloggning syns.
          var res = await new Promise(function (ok, no) {
            var h = setTimeout(function () { no(new Error("Ingen respons inom 150 s (hängd inloggning/popup?)")); }, 150000);
            t.f().then(function (v) { clearTimeout(h); ok(v); }, function (e) { clearTimeout(h); no(e); });
          });
          var okk = (typeof res === "number") ? (res >= 200 && res < 300) : !/^skriv \d|fel/.test(String(res));
          var inf = InlookAuth.info();
          row(t.n, "token OK via " + inf.mode + "; Graph: " + res, okk);
          if (inf.scp) log("  scp: " + inf.scp);
          if (inf.naaError) log("  NAA-fel (fallback användes): " + inf.naaError);
        } catch (e) {
          var inf2 = InlookAuth.info();
          row(t.n, explain(e) + " | " + String((e && (e.errorMessage || e.message)) || e).slice(0, 300), false);
          if (inf2.naaError) log("  NAA-fel: " + inf2.naaError);
        }
      }
      log("Klart. Kopiera hela rapporten ovan.");
    } catch (e) { log("FEL I FÖRPROVET: " + (e && e.message || e)); }
    setDis(false);
  }

  // Alla fel som annars sväljs ska synas i rapporten.
  window.addEventListener("error", function (ev) { log("JS-fel: " + ev.message + " (" + (ev.filename || "").split("/").pop() + ":" + ev.lineno + ")"); });
  window.addEventListener("unhandledrejection", function (ev) { var r = ev.reason; log("Ohanterat löftesfel: " + ((r && (r.errorMessage || r.message)) || r)); });

  // Knapparna kopplas direkt, inte först i Office.onReady, som kan utebli.
  function init() {
    if ($("run")) $("run").onclick = function () { run(false); };
    if ($("run1")) $("run1").onclick = function () { run("base1"); };
    if ($("runLow")) $("runLow").onclick = function () { run("low"); };
    if ($("runEx")) $("runEx").onclick = function () { run("ex"); };
    if ($("runOff")) $("runOff").onclick = function () { run("off"); };
    if ($("runIcs")) $("runIcs").onclick = function () { icsTest(); };
    log("Version " + ((document.getElementById("ver") || {}).textContent || "okänd (gammal sida i cache?)"));
    log("Sidan laddad. msal: " + (typeof msal) + ", Office: " + (typeof Office));
    if (window.Office && Office.onReady) {
      var t = setTimeout(function () { log("Office.onReady kom inte inom 5 s (fortsätter ändå)."); }, 5000);
      Office.onReady(function (i) { clearTimeout(t); log("Office.onReady: host=" + (i && i.host) + ", plattform=" + (i && i.platform)); });
    }
  }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init); else init();
})();
