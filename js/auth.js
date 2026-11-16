// Inloggning via MSAL.js. Första val: NAA (nested app authentication), där
// Outlook förmedlar token åt tillägget. Om NAA inte finns eller misslyckas av
// annat skäl än medgivande används en fallback: Office Dialog API öppnar
// dialog.html (samma origin) som loggar in med MSAL redirect och skickar
// tillbaka token i minnet. Inga hemligheter; allt är publik klient.
(function () {
  var cfg = window.INLOOK_CONFIG;
  var BASE = window.location.origin + window.location.pathname.replace(/[^/]*$/, "");
  var DIALOG_URL = BASE + "dialog.html";
  var pcaPromise = null, account = null;
  var mem = {};            // scope-nyckel -> {token, exp, scp}
  var lastInfo = { mode: "", scp: "" };

  function decodeScp(token) {
    try {
      var p = JSON.parse(atob(token.split(".")[1].replace(/-/g, "+").replace(/_/g, "/")));
      return p.scp || "";
    } catch (e) { return ""; }
  }

  function getPca() {
    if (!pcaPromise) {
      pcaPromise = msal.createNestablePublicClientApplication({
        auth: { clientId: cfg.clientId, authority: cfg.authority, supportsNestedAppAuth: true, redirectUri: BASE },
        cache: { cacheLocation: "localStorage" }
      });
    }
    return pcaPromise;
  }

  function isAdminErr(e) {
    var m = (e && (e.errorMessage || e.message)) || String(e);
    // Bara tydliga admin-krav. AADSTS65001/consent_required betyder "medgivande saknas"
    // och ska leda till ett interaktivt anrop, inte klassas som admin.
    return /AADSTS(90094|90008)|needs admin approval|admin.{0,20}approv/i.test(m);
  }

  async function viaNaa(scopes) {
    var pca = await getPca();
    var req = { scopes: scopes };
    if (account) req.account = account;
    var res;
    try { res = await pca.acquireTokenSilent(req); }
    catch (e) {
      if (isAdminErr(e)) throw e;
      // Tyst försök räcker inte (medgivande/inloggning saknas): fråga användaren.
      res = await pca.acquireTokenPopup(req);
    }
    account = res.account || account;
    return { token: res.accessToken, exp: res.expiresOn ? +new Date(res.expiresOn) : Date.now() + 3000000 };
  }

  // Fallback: Office-dialog -> dialog.html (MSAL redirect) -> token via messageParent.
  function viaDialog(scopes) {
    return new Promise(function (resolve, reject) {
      if (!(window.Office && Office.context && Office.context.ui && Office.context.ui.displayDialogAsync)) {
        return reject(new Error("Office Dialog API finns inte här."));
      }
      var url = DIALOG_URL + "?v=20261116h&scopes=" + encodeURIComponent(scopes.join(" "));
      Office.context.ui.displayDialogAsync(url, { height: 60, width: 30 }, function (r) {
        if (r.status !== Office.AsyncResultStatus.Succeeded) {
          return reject(new Error("Dialogen kunde inte öppnas: " + (r.error && r.error.message) + " (kod " + (r.error && r.error.code) + ")"));
        }
        var dlg = r.value, done = false;
        function finish(fn, v) { if (!done) { done = true; try { dlg.close(); } catch (e) { /* stängd */ } fn(v); } }
        dlg.addEventHandler(Office.EventType.DialogMessageReceived, function (a) {
          var m; try { m = JSON.parse(a.message); } catch (e) { return finish(reject, new Error("Ogiltigt dialogsvar")); }
          if (m.ok) finish(resolve, { token: m.token, exp: m.exp || Date.now() + 3000000 });
          else finish(reject, new Error(m.error || "Inloggning misslyckades"));
        });
        dlg.addEventHandler(Office.EventType.DialogEventReceived, function (a) {
          finish(reject, new Error("Dialogen stängdes eller fick fel (kod " + a.error + ")"));
        });
      });
    });
  }

  function withTimeout(p, ms, what) {
    return new Promise(function (res, rej) {
      var t = setTimeout(function () { rej(new Error(what + " svarade inte inom " + Math.round(ms / 1000) + " s")); }, ms);
      p.then(function (v) { clearTimeout(t); res(v); }, function (e) { clearTimeout(t); rej(e); });
    });
  }

  var queue = Promise.resolve(); // Office tillåter bara en dialog i taget

  // Token för EXAKT de scopes som anges (inkrementellt medgivande).
  async function getToken(scopes) {
    var key = scopes.slice().sort().join(" ");
    var run = queue.then(async function () {
      var c = mem[key];
      if (c && c.exp - Date.now() > 120000) { lastInfo = { mode: c.mode, scp: c.scp }; return c.token; }
      var r, mode;
      if (cfg.useNaa) {
        try { r = await withTimeout(viaNaa(scopes), 120000, "NAA"); mode = "NAA"; }
        catch (e) {
          if (isAdminErr(e) || /user_cancelled/i.test(e && e.errorCode || "")) throw e;
          lastInfo.naaError = (e && e.errorCode ? e.errorCode + ": " : "") + ((e && (e.errorMessage || e.message)) || "");
        }
      }
      if (!r) { r = await viaDialog(scopes); mode = "dialog"; }
      var scp = decodeScp(r.token);
      mem[key] = { token: r.token, exp: r.exp, scp: scp, mode: mode };
      lastInfo = { mode: mode, scp: scp, naaError: lastInfo.naaError };
      return r.token;
    });
    queue = run.catch(function () { /* kön ska fortsätta efter fel */ });
    return run;
  }

  window.InlookAuth = { isAdminErr: isAdminErr, getToken: getToken, getPca: getPca, info: function () { return lastInfo; } };
})();
