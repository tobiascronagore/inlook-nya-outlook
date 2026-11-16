// Microsoft To Do via Power Automate (se FLODEN.md). Tillägget har ingen Tasks-behörighet:
//  - skriver kommandofiler till OneDrive/Inlook/inkorg (flöde 1 utför dem i To Do)
//  - läser OneDrive/Inlook/todo.json (flöde 2 skriver den) och kvitton/ (flöde 1 skriver dem)
// Kommando: {action:"create"|"complete"|"reopen"|"delete", id (unikt kommando-id), listId, listName,
//   todoId (för complete/reopen/delete), title, body, dueDateTime, importance, createdAt}
(function (root) {
  var KEY = "inlook.pending.v1";
  var DAY = 86400000;
  var STUCK_MS = 15 * 60 * 1000;

  // ---------- Ren logik (testas i tests/) ----------

  // Ett kommando är klart när dess kvitto finns OCH todo.json skrevs efter kvittot.
  // Annars väntar det ("processed" = flödet har kört det, men listan är inte omläst).
  function resolvePending(pending, todo, receipts, now) {
    var gen = todo && todo.generatedAt ? Date.parse(todo.generatedAt) : 0;
    return pending.filter(function (p) {
      if (now - Date.parse(p.sentAt) > 2 * DAY) return false;
      var r = receipts[p.id];
      if (r && gen && Date.parse(r) <= gen) return false;
      return true;
    }).map(function (p) {
      var r = receipts[p.id];
      return Object.assign({}, p, { processed: !!r, stuck: !r && now - Date.parse(p.sentAt) > STUCK_MS });
    });
  }

  function firstLink(text) {
    var m = /https:\/\/[^\s"'<>)]+/i.exec(text || "");
    return m ? m[0].replace(/&amp;/g, "&") : "";
  }

  // Rader att visa: uppgifter från todo.json med väntande kommandon pålagda.
  function buildRows(todo, pending, opts) {
    opts = opts || {};
    var lists = (todo && todo.lists) || [];
    var names = {};
    lists.forEach(function (l) { names[l.id] = l.name; });
    var byTodo = {};
    pending.slice().sort(function (a, b) { return Date.parse(a.sentAt) - Date.parse(b.sentAt); })
      .forEach(function (p) { if (p.todoId) byTodo[p.todoId] = p; });
    var rows = [];
    ((todo && todo.tasks) || []).forEach(function (t) {
      var done = t.status === "completed", p = byTodo[t.id], badge = "", gone = false;
      if (p) {
        if (p.action === "complete") done = true;
        else if (p.action === "reopen") done = false;
        else if (p.action === "delete") gone = true;
        badge = p.processed ? "Bearbetad, väntar på nästa synk" : (p.stuck ? "Väntar på synk (ovanligt länge, kontrollera flödet)" : "Väntar på synk");
      }
      rows.push({ id: t.id, listId: t.listId, listName: names[t.listId] || "", title: t.title || "(namnlös)",
        done: done, due: (t.dueDateTime || "").slice(0, 10), link: firstLink(t.body), badge: badge,
        gone: gone, pending: !!p, created: t.createdAt || "" });
    });
    pending.forEach(function (p) {
      if (p.action !== "create") return;
      rows.push({ id: "", cmdId: p.id, listId: p.listId, listName: p.listName || names[p.listId] || "", title: p.title,
        done: false, due: (p.dueDateTime || "").slice(0, 10), link: firstLink(p.body), gone: false, pending: true, isNew: true,
        badge: p.processed ? "Bearbetad, väntar på nästa synk" : (p.stuck ? "Väntar på synk (ovanligt länge, kontrollera flödet)" : "Väntar på synk"),
        created: p.createdAt || p.sentAt });
    });
    rows = rows.filter(function (r) {
      return (!opts.listId || r.listId === opts.listId) && (opts.showDone || !r.done || r.pending);
    });
    rows.sort(function (a, b) {
      return (a.done - b.done) || (a.due || "9999").localeCompare(b.due || "9999") || (b.created || "").localeCompare(a.created || "");
    });
    return rows;
  }

  // Förvald lista: den vars namn matchar (t.ex. "Jobb"/"Privat") eller standardlistan/första.
  function pickList(lists, kind) {
    var re = kind === "privat" ? /privat/i : /jobb|arbete/i;
    for (var i = 0; i < lists.length; i++) if (re.test(lists[i].name || "")) return lists[i];
    return lists[0] || null;
  }

  // ---------- Lagring av väntande kommandon (lokalt i webbläsaren) ----------
  function loadPending() {
    try { return JSON.parse(root.localStorage.getItem(KEY) || "[]"); } catch (e) { return []; }
  }
  function savePending(list) {
    try { root.localStorage.setItem(KEY, JSON.stringify(list)); } catch (e) { /* ingen lagring */ }
  }

  // ---------- OneDrive ----------
  async function send(cmd) {
    var G = root.InlookGraph;
    cmd = Object.assign({ id: G.newId(), importance: "Normal", createdAt: new Date().toISOString() }, cmd);
    await G.writeJson("inkorg/cmd-" + cmd.id + ".json", cmd, null);
    var list = loadPending();
    list.push({ id: cmd.id, action: cmd.action, todoId: cmd.todoId || "", listId: cmd.listId || "", listName: cmd.listName || "",
      title: cmd.title || "", body: cmd.body || "", dueDateTime: cmd.dueDateTime || "", createdAt: cmd.createdAt, sentAt: new Date().toISOString() });
    savePending(list);
    return cmd;
  }

  // Läser todo.json + kvitton, rensar klara kommandon. Returnerar {todo, pending, syncedAt}.
  async function load() {
    var G = root.InlookGraph, now = Date.now();
    var r = await G.readJson("todo.json");
    var todo = r.data && Array.isArray(r.data.lists) ? r.data : null;
    var files = await G.listFolder("kvitton");
    var receipts = {};
    files.forEach(function (f) {
      var m = /^(.+)\.json$/.exec(f.name);
      if (m) receipts[m[1]] = f.modified;
    });
    var pending = resolvePending(loadPending(), todo, receipts, now);
    savePending(pending.map(function (p) { var q = Object.assign({}, p); delete q.processed; delete q.stuck; return q; }));
    // Städa gamla kvitton (>3 dygn), högst 10 per laddning.
    var old = files.filter(function (f) { return now - Date.parse(f.modified) > 3 * DAY; }).slice(0, 10);
    for (var i = 0; i < old.length; i++) { try { await G.deleteFile("kvitton/" + old[i].name); } catch (e) { /* nästa gång */ } }
    return { todo: todo, pending: pending, syncedAt: todo ? todo.generatedAt : "" };
  }

  root.InlookTodo = { resolvePending: resolvePending, buildRows: buildRows, pickList: pickList, firstLink: firstLink,
    send: send, load: load, loadPending: loadPending, savePending: savePending };
  if (typeof module !== "undefined") module.exports = root.InlookTodo;
})(typeof window !== "undefined" ? window : globalThis);
