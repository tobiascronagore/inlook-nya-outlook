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
    $("out").tBodies[0].appendChild(tr);
    log(name + ": " + text);
  }

  // Tolkar felet till "vem måste godkänna".
  function explain(e) {
    var m = (e && (e.errorMessage || e.message)) || String(e);
    var code = ((m.match(/AADSTS\d+/) || [])[0]) || (e && e.errorCode) || "";
    var admin = /AADSTS(65001|90094|90008|650057|90093)|admin/i.test(m) && !/AADSTS65004/.test(m);
    var cancelled = /user_cancelled|AADSTS65004|popup_window_error|empty_window/i.test(m + (e && e.errorCode || ""));
    var kind = admin ? "KRÄVER ADMINISTRATÖR" : cancelled ? "avbröts/popup blockerad" : "fel";
    return kind + " (" + (code || (e && e.errorCode) || "okänd kod") + ")";
  }

  var G = function (scopes, method, path, opts) { return InlookGraph.call(scopes, method, path, opts); };

  // Varje test: namn, scope, och en funktion som visar att token faktiskt fungerar.
  var tests = [
    { n: "User.Read", s: ["User.Read"], f: async function () { return (await G(["User.Read"], "GET", "/me?$select=id")).status; } },
    { n: "Tasks.ReadWrite (läs+skriv)", s: ["Tasks.ReadWrite"], f: async function () {
        var r = await G(["Tasks.ReadWrite"], "GET", "/me/todo/lists");
        if (!r.ok) return r.status;
        var l = (await r.json()).value[0];
        if (!l) return "inga listor";
        var c = await G(["Tasks.ReadWrite"], "POST", "/me/todo/lists/" + l.id + "/tasks", { json: { title: "Inlook förprov (raderas)" } });
        if (!c.ok) return "skriv " + c.status;
        var t = await c.json();
        var d = await G(["Tasks.ReadWrite"], "DELETE", "/me/todo/lists/" + l.id + "/tasks/" + t.id);
        return "läs 200, skapa " + c.status + ", radera " + d.status;
      } },
    { n: "Files.ReadWrite (OneDrive)", s: ["Files.ReadWrite"], f: async function () {
        var w = await InlookGraph.writeJson("_diag.json", { test: true }, null);
        await InlookGraph.deleteFile("_diag.json");
        return "skriv+radera OK";
      } },
    { n: "Mail.Read", s: ["Mail.Read"], f: async function () { return (await G(["Mail.Read"], "GET", "/me/messages?$top=1&$select=id")).status; } },
    { n: "Mail.ReadWrite", s: ["Mail.ReadWrite"], f: async function () { return (await G(["Mail.ReadWrite"], "GET", "/me/messages?$top=1&$select=id")).status; } },
    { n: "Notes.Read", s: ["Notes.Read"], f: async function () { return (await G(["Notes.Read"], "GET", "/me/onenote/notebooks?$top=1&$select=id")).status; } },
    { n: "Notes.ReadWrite", s: ["Notes.ReadWrite"], f: async function () { return (await G(["Notes.ReadWrite"], "GET", "/me/onenote/notebooks?$top=1&$select=id")).status; } },
    { n: "Calendars.Read", s: ["Calendars.Read"], f: async function () { return (await G(["Calendars.Read"], "GET", "/me/calendars?$top=1&$select=id")).status; } },
    { n: "Calendars.ReadWrite", s: ["Calendars.ReadWrite"], f: async function () { return (await G(["Calendars.ReadWrite"], "GET", "/me/calendars?$top=1&$select=id")).status; } }
  ];

  async function run(only) {
    $("run").disabled = true; $("run1").disabled = true;
    $("out").tBodies[0].innerHTML = ""; lines = [];
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
      var list = only ? tests.slice(0, 1) : tests;
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
          if (inf.scp) log("  scp (redan medgivna för Graph): " + inf.scp);
          if (inf.naaError) log("  NAA-fel (fallback användes): " + inf.naaError);
        } catch (e) {
          var inf2 = InlookAuth.info();
          row(t.n, explain(e) + " | " + String((e && (e.errorMessage || e.message)) || e).slice(0, 300), false);
          if (inf2.naaError) log("  NAA-fel: " + inf2.naaError);
        }
      }
      log("Klart. Kopiera hela rapporten ovan.");
    } catch (e) { log("FEL I FÖRPROVET: " + (e && e.message || e)); }
    $("run").disabled = false; $("run1").disabled = false;
  }

  // Alla fel som annars sväljs ska synas i rapporten.
  window.addEventListener("error", function (ev) { log("JS-fel: " + ev.message + " (" + (ev.filename || "").split("/").pop() + ":" + ev.lineno + ")"); });
  window.addEventListener("unhandledrejection", function (ev) { var r = ev.reason; log("Ohanterat löftesfel: " + ((r && (r.errorMessage || r.message)) || r)); });

  // Knapparna kopplas direkt, inte först i Office.onReady, som kan utebli.
  function init() {
    $("run").onclick = function () { run(false); };
    $("run1").onclick = function () { run(true); };
    log("Sidan laddad. msal: " + (typeof msal) + ", Office: " + (typeof Office));
    if (window.Office && Office.onReady) {
      var t = setTimeout(function () { log("Office.onReady kom inte inom 5 s (fortsätter ändå)."); }, 5000);
      Office.onReady(function (i) { clearTimeout(t); log("Office.onReady: host=" + (i && i.host) + ", plattform=" + (i && i.platform)); });
    }
  }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init); else init();
})();
