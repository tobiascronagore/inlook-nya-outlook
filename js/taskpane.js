// Tilläggets panel: uppgift och anteckning från valt mejl + lista anteckningar.
(function () {
  var $ = function (id) { return document.getElementById(id); };
  var current = { subject: "", from: "", link: "" };

  function status(msg, kind, details) {
    var el = $("status");
    el.textContent = msg || "";
    el.className = "status" + (kind ? " " + kind : "");
    if (details) {
      var d = document.createElement("details");
      var sm = document.createElement("summary"); sm.textContent = "Visa detaljer";
      var pre = document.createElement("pre"); pre.textContent = details;
      d.append(sm, pre); el.appendChild(d);
    }
  }

  // Kort, begripligt fel; rå text (max 500 tecken) bakom "Visa detaljer".
  function fail(e) {
    var raw = String((e && (e.errorMessage || e.message)) || e);
    var code = (raw.match(/AADSTS\d+/) || [])[0] || (e && e.errorCode) || "";
    var msg;
    if (InlookAuth.isAdminErr(e)) msg = "Den här behörigheten kräver administratörsgodkännande i din organisation.";
    else if (/user_cancelled|AADSTS65004/i.test(raw + (e && e.errorCode || ""))) msg = "Inloggningen avbröts.";
    else if (/AADSTS65001|consent_required|interaction_required/i.test(raw)) msg = "Medgivande saknas. Klicka igen och godkänn i fönstret som visas.";
    else msg = "Något gick fel" + (code ? " (" + code + ")" : "") + ".";
    status(msg, "err", raw.length > 500 ? raw.slice(0, 500) + "…" : raw);
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

  var todoState = { todo: null, pending: [] };

  function lists() { return (todoState.todo && todoState.todo.lists) || []; }

  function fillSelect(sel, withAll, kind) {
    var cur = sel.value;
    sel.innerHTML = "";
    if (withAll) { var o = document.createElement("option"); o.value = ""; o.textContent = "Alla listor"; sel.appendChild(o); }
    lists().forEach(function (l) { var o = document.createElement("option"); o.value = l.id; o.textContent = l.name; sel.appendChild(o); });
    if (cur && lists().some(function (l) { return l.id === cur; })) sel.value = cur;
    else if (!withAll) { var p = InlookTodo.pickList(lists(), kind); if (p) sel.value = p.id; }
  }

  function syncLabel() {
    var t = todoState.todo;
    $("sync-info").textContent = t && t.generatedAt
      ? "Senast synkad från To Do: " + new Date(t.generatedAt).toLocaleString("sv-SE") + ". Ändringar tar några minuter."
      : "Ingen To Do-data än: flödet \"To Do till Inlook\" har inte skrivit todo.json (se FLODEN.md).";
  }

  async function refreshTodo() {
    todoState = await InlookTodo.load();
    fillSelect($("task-list"), false, "jobb");
    fillSelect($("nt-list"), false, "jobb");
    fillSelect($("filt-list"), true);
    syncLabel();
  }

  async function onTask() {
    var btn = $("btn-task"); btn.disabled = true;
    try {
      status("Skickar uppgift…");
      if (!todoState.todo) await refreshTodo();
      var listId = $("task-list").value;
      if (!listId) throw new Error("Inga To Do-listor hittades. Flödet \"To Do till Inlook\" måste ha körts först.");
      var due = $("task-due").value;
      var title = $("task-title").value.trim() || current.subject || "(namnlös)";
      await InlookTodo.send({ action: "create", listId: listId, listName: $("task-list").selectedOptions[0].textContent,
        title: title, dueDateTime: due ? due + "T00:00:00Z" : "",
        body: (current.link ? "Öppna mejlet: " + current.link + "\n\n" : "") + "Från mejl: " + current.subject + (current.from ? "\nAvsändare: " + current.from : "") });
      status("Uppgiften skickades. Den hamnar i To Do inom några minuter (väntar på synk).", "ok");
    } catch (e) { fail(e); }
    btn.disabled = false;
  }

  function renderTasks() {
    var ul = $("tasks"); ul.innerHTML = "";
    var rows = InlookTodo.buildRows(todoState.todo, todoState.pending, { listId: $("filt-list").value, showDone: $("show-done").checked });
    if (!rows.length) { ul.innerHTML = "<li class='muted'>Inga uppgifter.</li>"; return; }
    rows.forEach(function (t) {
      var li = document.createElement("li");
      var cb = document.createElement("input"); cb.type = "checkbox"; cb.checked = !!t.done; cb.disabled = !!t.isNew || t.gone;
      cb.onchange = async function () {
        cb.disabled = true;
        try {
          await InlookTodo.send({ action: cb.checked ? "complete" : "reopen", todoId: t.id, listId: t.listId, title: t.title });
          todoState.pending = InlookTodo.resolvePending(InlookTodo.loadPending(), todoState.todo, {}, Date.now());
          renderTasks(); status("Skickat, väntar på synk.", "ok");
        } catch (e) { cb.checked = !cb.checked; cb.disabled = false; fail(e); }
      };
      var tt = document.createElement("span"); tt.className = "t"; tt.textContent = " " + t.title;
      if (t.done || t.gone) tt.style.textDecoration = "line-through";
      var m = document.createElement("div"); m.className = "muted";
      m.textContent = (t.listName || "") + (t.due ? " · förfaller " + t.due : "") + (t.gone ? " · tas bort" : "");
      li.append(cb, tt, m);
      if (t.badge) { var bd = document.createElement("div"); bd.className = "badge"; bd.textContent = "Väntar: " + t.badge; li.appendChild(bd); }
      if (t.link && /^https:/i.test(t.link)) {
        var a = document.createElement("a"); a.href = t.link; a.target = "_blank"; a.rel = "noopener"; a.textContent = "Öppna mejlet"; li.appendChild(a);
      }
      if (!t.isNew && !t.gone) {
        var d = document.createElement("button"); d.className = "del"; d.textContent = "Ta bort";
        d.onclick = async function () {
          if (d.dataset.armed !== "1") { d.dataset.armed = "1"; d.textContent = "Klicka igen för att ta bort"; return; }
          try {
            await InlookTodo.send({ action: "delete", todoId: t.id, listId: t.listId, title: t.title });
            todoState.pending = InlookTodo.resolvePending(InlookTodo.loadPending(), todoState.todo, {}, Date.now());
            renderTasks(); status("Borttagning skickad, väntar på synk.", "ok");
          } catch (e) { fail(e); }
        };
        li.appendChild(d);
      }
      ul.appendChild(li);
    });
  }

  async function loadTasks() {
    try { status("Läser uppgifter…"); await refreshTodo(); renderTasks(); status(""); }
    catch (e) { fail(e); }
  }

  async function onNewTask() {
    var title = $("nt-title").value.trim();
    if (!title) { status("Skriv en rubrik.", "err"); return; }
    try {
      var sel = $("nt-list");
      if (!sel.value) throw new Error("Inga To Do-listor hittades. Flödet \"To Do till Inlook\" måste ha körts först.");
      await InlookTodo.send({ action: "create", listId: sel.value, listName: sel.selectedOptions[0].textContent, title: title, body: "" });
      $("nt-title").value = "";
      todoState.pending = InlookTodo.resolvePending(InlookTodo.loadPending(), todoState.todo, {}, Date.now());
      renderTasks(); status("Skickad, väntar på synk.", "ok");
    } catch (e) { fail(e); }
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
    ["mail", "tasks", "notes", "settings"].forEach(function (t) { $("tab-" + t).hidden = (t !== name); });
    status("");
    if (name === "notes") loadNotes();
    if (name === "tasks") loadTasks();
  }

  Office.onReady(function () {
    document.querySelectorAll(".tab").forEach(function (b) { b.onclick = function () { showTab(b.dataset.tab); }; });
    $("btn-task").onclick = onTask;
    $("btn-note").onclick = onNote;
    $("btn-reload").onclick = loadNotes;
    $("btn-tasks-reload").onclick = loadTasks;
    $("btn-nt").onclick = onNewTask;
    $("show-done").onchange = renderTasks;
    $("filt-list").onchange = renderTasks;
    refreshTodo().catch(function () { /* listor hämtas igen vid knapptryck */ });
    $("btn-import").onclick = onImport;
    readItem();
    // Fastnålad panel: uppdatera när användaren byter mejl.
    try {
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, readItem);
    } catch (e) { /* äldre värd utan stöd */ }
  });
})();
