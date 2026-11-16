// Graph-anrop mot användarens eget konto: OneDrive-mappen "Inlook" (anteckningar och To Do-kommandon).
(function () {
  var cfg = window.INLOOK_CONFIG;
  var BASE = "https://graph.microsoft.com/v1.0";
  var S_FILES = ["Files.ReadWrite"];

  // Ett anrop med en enda omförsök vid strypning (429/503): hellre få anrop.
  async function call(scopes, method, path, opts) {
    opts = opts || {};
    var token = await InlookAuth.getToken(scopes);
    var headers = Object.assign({ Authorization: "Bearer " + token }, opts.headers || {});
    var init = { method: method, headers: headers };
    if (opts.json !== undefined) {
      headers["Content-Type"] = "application/json";
      init.body = JSON.stringify(opts.json);
    } else if (opts.body !== undefined) {
      init.body = opts.body;
    }
    var res = await fetch(BASE + path, init);
    if ((res.status === 429 || res.status === 503) && !opts._retried) {
      var wait = Math.min(parseInt(res.headers.get("Retry-After") || "2", 10) || 2, 10);
      await new Promise(function (r) { setTimeout(r, wait * 1000); });
      return call(scopes, method, path, Object.assign({}, opts, { _retried: true }));
    }
    return res;
  }

  async function mustOk(res, what) {
    if (res.ok) return res;
    var txt = "";
    try { txt = (await res.json()).error.message || ""; } catch (e) { /* ingen JSON */ }
    var err = new Error(what + " misslyckades (" + res.status + ")" + (txt ? ": " + txt : ""));
    err.status = res.status;
    throw err;
  }

  // ---------- OneDrive-JSON ----------
  // name kan innehålla undermappar ("inkorg/cmd-1.json"); varje del kodas för sig.
  function filePath(name) {
    return "/me/drive/root:/" + encodeURIComponent(cfg.driveFolder) + "/" +
      String(name).split("/").map(encodeURIComponent).join("/");
  }

  // Lista filer i en undermapp: [{name, modified}]. Saknas mappen: tom lista.
  async function listFolder(name) {
    var res = await call(S_FILES, "GET", filePath(name) + ":/children?$select=name,lastModifiedDateTime&$top=200");
    if (res.status === 404) return [];
    await mustOk(res, "Lista " + name);
    return ((await res.json()).value || []).map(function (f) { return { name: f.name, modified: f.lastModifiedDateTime }; });
  }

  // Returnerar {data, etag}; data=null om filen inte finns. ETag hämtas FÖRE
  // innehållet: hinner filen ändras emellan blir skrivningen en säker 412.
  async function readJson(name) {
    var meta = await call(S_FILES, "GET", filePath(name) + "?$select=eTag");
    if (meta.status === 404) return { data: null, etag: null };
    await mustOk(meta, "Läs metadata");
    var etag = (await meta.json()).eTag || null;
    var res = await mustOk(await call(S_FILES, "GET", filePath(name) + ":/content"), "Läs " + name);
    return { data: JSON.parse(await res.text()), etag: etag };
  }

  // If-Match skyddar mot att en annan flik/dator skrivit emellan; 412 = konflikt.
  async function writeJson(name, data, etag) {
    var headers = { "Content-Type": "application/json" };
    var q = "";
    if (etag) headers["If-Match"] = etag;
    else q = "?@microsoft.graph.conflictBehavior=replace";
    var res = await call(S_FILES, "PUT", filePath(name) + ":/content" + q,
      { body: JSON.stringify(data, null, 2), headers: headers });
    if (res.status === 412) {
      var e = new Error("Filen ändrades någon annanstans. Ladda om och försök igen.");
      e.status = 412;
      throw e;
    }
    await mustOk(res, "Skriv " + name);
    return res.json();
  }

  async function deleteFile(name) {
    var res = await call(S_FILES, "DELETE", filePath(name));
    if (res.status !== 404) await mustOk(res, "Radera " + name);
  }

  // ---------- Anteckningar (samma format som desktop/.data/notes.json) ----------
  function nowStamp() {
    var d = new Date(), p = function (n) { return String(n).padStart(2, "0"); };
    return d.getFullYear() + "-" + p(d.getMonth() + 1) + "-" + p(d.getDate()) + "T" + p(d.getHours()) + ":" + p(d.getMinutes());
  }
  function newId() {
    var a = new Uint8Array(16); crypto.getRandomValues(a);
    return Array.prototype.map.call(a, function (b) { return b.toString(16).padStart(2, "0"); }).join("");
  }

  async function loadNotes() {
    var r = await readJson("notes.json");
    var list = Array.isArray(r.data) ? r.data : [];
    list.sort(function (a, b) { return (b.updated || "").localeCompare(a.updated || ""); });
    return { notes: list, etag: r.etag };
  }

  // Läs-ändra-skriv med en enda omförsök vid konflikt.
  async function mutateNotes(fn) {
    for (var attempt = 0; attempt < 2; attempt++) {
      var r = await readJson("notes.json");
      var list = Array.isArray(r.data) ? r.data : [];
      list = fn(list);
      try { await writeJson("notes.json", list, r.etag); return list; }
      catch (e) { if (e.status !== 412 || attempt === 1) throw e; }
    }
  }

  function addNote(title, body, link) {
    return mutateNotes(function (list) {
      list.push({
        id: newId(), title: title || "(namnlös)", body: body || "", link: link || "",
        taskEntryId: null, taskSubject: "", pageId: null, created: nowStamp(), updated: nowStamp()
      });
      return list;
    });
  }

  function deleteNote(id) {
    return mutateNotes(function (list) { return list.filter(function (n) { return n.id !== id; }); });
  }

  // Engångsimport: slår ihop på id, befintliga anteckningar i molnet ändras inte.
  async function importNotes(incoming) {
    var added = 0;
    await mutateNotes(function (list) {
      var have = {};
      list.forEach(function (n) { have[n.id] = true; });
      incoming.forEach(function (n) {
        if (n && n.id && !have[n.id]) { list.push(n); added++; }
      });
      return list;
    });
    return added;
  }

  window.InlookGraph = {
    readJson: readJson, writeJson: writeJson, deleteFile: deleteFile,
    loadNotes: loadNotes, addNote: addNote, deleteNote: deleteNote, importNotes: importNotes,
    listFolder: listFolder, newId: newId, nowStamp: nowStamp,
    call: call
  };
})();
