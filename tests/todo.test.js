// Kör: node tests/todo.test.js   (mockad OneDrive + simulerade flöden, inget nätverk)
const fs = require("fs"), path = require("path"), assert = require("assert");
const root = path.join(__dirname, "..");
globalThis.window = globalThis;
const store = {}; globalThis.localStorage = { getItem: k => store[k] ?? null, setItem: (k, v) => { store[k] = v; } };
globalThis.crypto = require("crypto").webcrypto;
globalThis.InlookAuth = { getToken: async () => "t" };

// ---- Mock-OneDrive: {sökväg under Inlook/: {content, etag, modified}} ----
const drive = {}; let tick = 0, clock = Date.parse("2026-11-16T10:00:00Z");
const RealDate = Date;
globalThis.Date = class extends RealDate { constructor(...a) { a.length ? super(...a) : super(clock); } static now() { return clock; } };
const now = () => new Date(clock).toISOString();
function put(p, content) { drive[p] = { content, etag: "e" + (++tick), modified: now() }; }
globalThis.fetch = async (url, init) => {
  init = init || {}; const u = new URL(url); let p = decodeURIComponent(u.pathname.replace(/^\/v1\.0\/me\/drive\/root:\/Inlook\//, ""));
  const ok = (b, s) => ({ ok: !s || s < 300, status: s || 200, headers: { get: () => null }, json: async () => b, text: async () => b });
  if (p.endsWith(":/children")) {
    const dir = p.slice(0, -10) + "/";
    const v = Object.keys(drive).filter(k => k.startsWith(dir) && !k.slice(dir.length).includes("/"))
      .map(k => ({ name: k.slice(dir.length), lastModifiedDateTime: drive[k].modified }));
    return ok({ value: v });
  }
  if (p.endsWith(":/content")) {
    p = p.slice(0, -9);
    if (init.method === "PUT") { put(p, init.body); return ok({}); }
    return drive[p] ? ok(drive[p].content) : ok({ error: {} }, 404);
  }
  if (init.method === "DELETE") { if (!drive[p]) return ok({}, 404); delete drive[p]; return ok({}, 204); }
  return drive[p] ? ok({ eTag: drive[p].etag }) : ok({ error: {} }, 404);
};
for (const f of ["config", "graph", "todo"]) (0, eval)(fs.readFileSync(path.join(root, "js", f + ".js"), "utf8"));
const T = globalThis.InlookTodo, G = globalThis.InlookGraph;

// ---- Simulerade flöden (samma kontrakt som FLODEN.md) ----
const toDo = { lists: [{ id: "L1", name: "Jobb" }, { id: "L2", name: "Privat" }], tasks: [], n: 0 };
function flow1() { // "Inlook till To Do"
  Object.keys(drive).filter(k => /^inkorg\/.+\.json$/.test(k)).forEach(k => {
    const c = JSON.parse(drive[k].content);
    switch (c.action) {
      case "create": toDo.tasks.push({ id: "T" + (++toDo.n), listId: c.listId, title: c.title, status: "notStarted", dueDateTime: c.dueDateTime, body: c.body, createdAt: c.createdAt }); break;
      case "complete": toDo.tasks.find(t => t.id === c.todoId).status = "completed"; break;
      case "reopen": toDo.tasks.find(t => t.id === c.todoId).status = "notStarted"; break;
      case "delete": toDo.tasks = toDo.tasks.filter(t => t.id !== c.todoId); break;
    }
    put("kvitton/" + c.id + ".json", JSON.stringify({ id: c.id, action: c.action, ok: true }));
    delete drive[k];
  });
}
function flow2() { put("todo.json", JSON.stringify({ generatedAt: now(), lists: toDo.lists, tasks: JSON.parse(JSON.stringify(toDo.tasks)) })); }
const step = ms => { clock += ms; };

(async () => {
  // 1. Skriva kommando: hamnar i inkorg, väntar på synk
  flow2(); step(1000);
  let st = await T.load(); assert.equal(st.todo.lists.length, 2);
  assert.equal(T.pickList(st.todo.lists, "privat").id, "L2"); assert.equal(T.pickList(st.todo.lists, "jobb").id, "L1");
  await T.send({ action: "create", listId: "L1", listName: "Jobb", title: "Svara kund", dueDateTime: "2026-11-20T00:00:00Z",
    body: "Öppna mejlet: https://outlook.office.com/owa/?ItemID=a%3D&exvsurl=1\n\nFrån mejl: X" });
  assert.equal(Object.keys(drive).filter(k => k.startsWith("inkorg/cmd-")).length, 1);
  const cmd = JSON.parse(drive[Object.keys(drive).find(k => k.startsWith("inkorg/"))].content);
  assert.deepEqual(Object.keys(cmd).sort(), ["action", "body", "createdAt", "dueDateTime", "id", "importance", "listId", "listName", "title"]);
  st = await T.load(); let rows = T.buildRows(st.todo, st.pending, {}, st.created);
  assert.equal(rows.length, 1); assert.ok(rows[0].isNew && rows[0].badge === "Synkas…" && !rows[0].ok); assert.equal(rows[0].due, "2026-11-20");
  assert.equal(rows[0].link, "https://outlook.office.com/owa/?ItemID=a%3D&exvsurl=1");

  // 2. Flöde 1 körs, flöde 2 ännu inte: bearbetat men väntar fortfarande
  step(60000); flow1();
  st = await T.load(); rows = T.buildRows(st.todo, st.pending, {}, st.created);
  assert.equal(rows.length, 1); assert.ok(rows[0].isNew && rows[0].ok && rows[0].badge === ""); // tasks tom i todo.json: finns kvar lokalt

  // 3. Flöde 2 körs: kommandot är löst, uppgiften kommer från todo.json
  step(60000); flow2(); step(1000);
  st = await T.load(); rows = T.buildRows(st.todo, st.pending, {}, st.created);
  assert.equal(st.pending.length, 0); assert.equal(st.created.length, 0); assert.equal(rows.length, 1); assert.ok(!rows[0].isNew && !rows[0].badge && rows[0].id === "T1");

  // 4. Bocka av: optimistiskt + väntar, sedan klart
  await T.send({ action: "complete", todoId: "T1", listId: "L1" });
  st = await T.load(); rows = T.buildRows(st.todo, st.pending, { showDone: false });
  assert.equal(rows.length, 1); assert.ok(rows[0].done && /Väntar/.test(rows[0].badge));
  step(60000); flow1(); step(60000); flow2(); step(1000);
  st = await T.load(); assert.equal(st.pending.length, 0);
  assert.equal(T.buildRows(st.todo, st.pending, { showDone: false }).length, 0);
  rows = T.buildRows(st.todo, st.pending, { showDone: true }); assert.ok(rows[0].done && !rows[0].badge);

  // 5. Återöppna och ta bort
  await T.send({ action: "reopen", todoId: "T1", listId: "L1" });
  step(1000); flow1(); step(1000); flow2(); step(1000); st = await T.load();
  assert.ok(!T.buildRows(st.todo, st.pending)[0].done);
  await T.send({ action: "delete", todoId: "T1", listId: "L1" });
  st = await T.load(); assert.ok(T.buildRows(st.todo, st.pending)[0].gone);
  step(1000); flow1(); step(1000); flow2(); step(1000); st = await T.load();
  assert.equal(T.buildRows(st.todo, st.pending, { showDone: true }).length, 0);

  // 6. Flödet står still: "ovanligt länge" efter 15 min; tas bort efter 2 dygn
  await T.send({ action: "create", listId: "L2", listName: "Privat", title: "Fastnad" });
  step(16 * 60000); st = await T.load();
  assert.ok(/ovanligt länge/.test(T.buildRows(st.todo, st.pending, {}, st.created)[0].badge));
  step(2 * 86400000); st = await T.load(); assert.equal(st.pending.length, 0);

  // 6b. Optimistisk uppgift överlever todo.json utan tasks, och dedupas först när en matchande uppgift finns
  assert.equal(st.created.length, 1);
  const c0 = st.created[0];
  let td = { generatedAt: now(), lists: toDo.lists, tasks: [{ id: "OLD", listId: c0.listId, title: c0.title, status: "notStarted", createdAt: "2026-01-01T00:00:00Z" }] };
  assert.equal(T.updateCreated(st.created, td, {}, clock).length, 1); // äldre uppgift med samma rubrik räknas inte
  td.tasks[0].createdAt = now();
  assert.equal(T.updateCreated(st.created, td, {}, clock).length, 0);
  td.tasks[0] = { id: "X", listId: "L1", title: c0.title }; assert.equal(T.updateCreated(st.created, td, {}, clock).length, 1);
  step(8 * 86400000); st = await T.load(); assert.equal(st.created.length, 0); // 7 dygn

  // 7. Listfilter + tom todo.json
  assert.equal(T.buildRows(st.todo, [], { listId: "L2" }).length, 0);
  assert.equal(T.buildRows(null, []).length, 0);

  // 8. Gamla kvitton städas
  step(4 * 86400000); await T.load();
  assert.equal(Object.keys(drive).filter(k => k.startsWith("kvitton/")).length, 0);
  console.log("OK: alla tester gick igenom");
})().catch(e => { console.error(e); process.exit(1); });
