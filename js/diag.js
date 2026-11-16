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

  async function run() {
    $("run").disabled = true;
    $("out").tBodies[0].innerHTML = ""; lines = [];
    var inOffice = !!(window.Office && Office.context && Office.context.mailbox);
    log("Datum: " + new Date().toISOString().slice(0, 10));
    log("Office-värd: " + (inOffice ? ((Office.context.mailbox.diagnostics || {}).hostName || "okänd") + " " + ((Office.context.mailbox.diagnostics || {}).hostVersion || "") : "ingen (vanlig webbläsare)"));
    if (inOffice && Office.context.requirements) {
      log("NestedAppAuth 1.1 stöds: " + Office.context.requirements.isSetSupported("NestedAppAuth", "1.1"));
      log("Mailbox 1.8 stöds: " + Office.context.requirements.isSetSupported("Mailbox", "1.8"));
      var it = Office.context.mailbox.item;
      log("Mejl-id tillgängligt: " + !!(it && it.itemId));
    }
    for (var i = 0; i < tests.length; i++) {
      var t = tests[i];
      try {
        var res = await t.f();
        var ok = (typeof res === "number") ? (res >= 200 && res < 300) : !/^skriv \d|fel/.test(String(res));
        row(t.n, "token OK; Graph: " + res, ok);
      } catch (e) {
        row(t.n, explain(e), false);
      }
    }
    log("Klart. Kopiera hela rapporten ovan.");
    $("run").disabled = false;
  }

  Office.onReady(function () { $("run").onclick = run; });
})();
