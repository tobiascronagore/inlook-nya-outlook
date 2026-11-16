// Tilläggets panel: uppgift och anteckning från valt mejl + lista anteckningar.
(function () {
  var $ = function (id) { return document.getElementById(id); };
  var lists = [];
  var current = { subject: "", from: "", link: "" };

  function status(msg, kind) {
    var el = $("status");
    el.textContent = msg || "";
    el.className = "status" + (kind ? " " + kind : "");
  }

  // Visar ett begripligt fel; tenantens medgivandekrav får en egen förklaring.
  function fail(e) {
    var m = (e && e.message) || String(e);
    if (/AADSTS(65001|90094|65002|650057)|admin|consent_required|interaction_required/i.test(m)) {
      m = "Medgivande saknas eller kräver administratör: " + m;
    }
    status(m, "err");
  }

  // Länk som öppnar mejlet i Outlook på webben. Bygger bara på item-id från
  // Office.js, så ingen Graph-mailbehörighet behövs. (Prövas i förprovet.)
  function mailLink(itemId) {
    if (!itemId) return "";
    return "https://outlook.office.com/owa/?ItemID=" + encodeURIComponent(itemId) +
      "&exvsurl=1&viewmodel=ReadMessageItem";
  }

  function readItem() {
    var it = Office.context.mailbox && Office.context.mailbox.item;
    if (!it) { current = { subject: "", from: "", link: "" }; }
    else {
      var f = it.from && (it.from.displayName || it.from.emailAddress) || "";
      current = { subject: it.subject || "", from: f, link: mailLink(it.itemId) };
    }
    $("subject").textContent = current.subject || "–";
    $("from").textContent = current.from;
    $("task-title").value = current.subject;
  }

  function listFor(kind) {
    var saved = {};
    try { saved = JSON.parse(localStorage.getItem("inlook.lists") || "{}"); } catch (e) { /* tom */ }
    if (saved[kind] && lists.some(function (l) { return l.id === saved[kind]; })) return saved[kind];
    var found;
    if (kind === "privat") found = lists.find(function (l) { return /priv/i.test(l.displayName); });
    else found = lists.find(function (l) { return l.wellknownListName === "defaultList"; });
    return (found || lists[0] || {}).id;
  }

  function fillSelect(sel, kind) {
    sel.innerHTML = "";
    lists.forEach(function (l) {
      var o = document.createElement("option");
      o.value = l.id; o.textContent = l.displayName; sel.appendChild(o);
    });
    sel.value = listFor(kind);
    sel.onchange = function () {
      var saved = {};
      try { saved = JSON.parse(localStorage.getItem("inlook.lists") || "{}"); } catch (e) { /* tom */ }
      saved[kind] = sel.value;
      localStorage.setItem("inlook.lists", JSON.stringify(saved));
    };
  }

  async function ensureLists() {
    if (lists.length) return;
    lists = await InlookGraph.todoLists();
    fillSelect($("list-jobb"), "jobb");
    fillSelect($("list-privat"), "privat");
  }

  async function onTask() {
    var btn = $("btn-task"); btn.disabled = true;
    try {
      status("Skapar uppgift…");
      await ensureLists();
      var kind = document.querySelector("input[name=kind]:checked").value;
      var listId = listFor(kind);
      if (!listId) throw new Error("Hittade ingen To Do-lista.");
      var title = $("task-title").value.trim() || current.subject || "(namnlös)";
      var body = "Från mejl: " + current.subject + (current.from ? "\nAvsändare: " + current.from : "") +
        (current.link ? "\nÖppna mejlet: " + current.link : "");
      await InlookGraph.createTask(listId, title, body, current.link);
      status("Uppgiften skapades i " + (kind === "privat" ? "Privat" : "Jobb") + ".", "ok");
    } catch (e) { fail(e); }
    btn.disabled = false;
  }

  async function onNote() {
    var text = $("note-body").value.trim();
    if (!text) { status("Skriv en anteckning först.", "err"); return; }
    var btn = $("btn-note"); btn.disabled = true;
    try {
      status("Sparar anteckning…");
      await InlookGraph.addNote(current.subject, text, current.link);
      $("note-body").value = "";
      status("Anteckningen sparades i OneDrive.", "ok");
    } catch (e) { fail(e); }
    btn.disabled = false;
  }

  function renderNotes(notes) {
    var ul = $("notes"); ul.innerHTML = "";
    notes = notes.slice().sort(function (a, b) { return (b.updated || "").localeCompare(a.updated || ""); });
    if (!notes.length) { ul.innerHTML = "<li class='muted'>Inga anteckningar än.</li>"; return; }
    notes.forEach(function (n) {
      var li = document.createElement("li");
      var t = document.createElement("div"); t.className = "t"; t.textContent = n.title || "(namnlös)";
      var m = document.createElement("div"); m.className = "muted"; m.textContent = (n.updated || "").replace("T", " ");
      var b = document.createElement("div"); b.className = "b"; b.textContent = n.body || "";
      li.append(t, m, b);
      if (n.link && /^https:/i.test(n.link)) {
        var a = document.createElement("a"); a.href = n.link; a.target = "_blank"; a.rel = "noopener";
        a.textContent = "Öppna länk"; li.appendChild(a);
      }
      var d = document.createElement("button"); d.className = "del"; d.textContent = "Ta bort";
      d.onclick = async function () {
        // Två steg i stället för confirm(), som kan vara blockerat i panelen.
        if (d.dataset.armed !== "1") { d.dataset.armed = "1"; d.textContent = "Klicka igen för att ta bort"; return; }
        try { renderNotes(await InlookGraph.deleteNote(n.id)); status("Borttagen.", "ok"); }
        catch (e) { fail(e); }
      };
      li.appendChild(d);
      ul.appendChild(li);
    });
  }

  async function loadNotes() {
    try { status("Läser anteckningar…"); renderNotes((await InlookGraph.loadNotes()).notes); status(""); }
    catch (e) { fail(e); }
  }

  async function onImport() {
    var f = $("import-file").files[0];
    if (!f) { status("Välj en fil först.", "err"); return; }
    try {
      var data = JSON.parse(await f.text());
      if (/dagbok/i.test(f.name)) {
        var cur = await InlookGraph.readJson("dagbok.json");
        if (cur.data && $("btn-import").dataset.armed !== "1") {
          $("btn-import").dataset.armed = "1";
          status("dagbok.json finns redan i OneDrive. Klicka Importera igen för att skriva över.", "err");
          return;
        }
        await InlookGraph.writeJson("dagbok.json", data, cur.etag);
        status("dagbok.json importerad till OneDrive/Inlook.", "ok");
      } else if (/notes/i.test(f.name) && Array.isArray(data)) {
        var n = await InlookGraph.importNotes(data);
        status(n + " anteckningar importerade (befintliga lämnades orörda).", "ok");
        loadNotes();
      } else {
        status("Okänd fil. Välj notes.json eller dagbok.json.", "err");
      }
      $("btn-import").dataset.armed = "";
    } catch (e) { fail(e); }
  }

  function showTab(name) {
    document.querySelectorAll(".tab").forEach(function (b) { b.classList.toggle("active", b.dataset.tab === name); });
    ["mail", "notes", "settings"].forEach(function (t) { $("tab-" + t).hidden = (t !== name); });
    status("");
    if (name === "notes") loadNotes();
    if (name === "settings") ensureLists().catch(fail);
  }

  Office.onReady(function () {
    document.querySelectorAll(".tab").forEach(function (b) { b.onclick = function () { showTab(b.dataset.tab); }; });
    $("btn-task").onclick = onTask;
    $("btn-note").onclick = onNote;
    $("btn-reload").onclick = loadNotes;
    $("btn-import").onclick = onImport;
    readItem();
    // Fastnålad panel: uppdatera när användaren byter mejl.
    try {
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, readItem);
    } catch (e) { /* äldre värd utan stöd */ }
  });
})();
