// Fallback-inloggning i Office-dialog: MSAL redirect på samma origin.
// Redirect-URI = den här sidans adress (måste vara registrerad som SPA).
(function () {
  var cfg = window.INLOOK_CONFIG;
  var redirectUri = window.location.origin + window.location.pathname;
  function send(o) { Office.context.ui.messageParent(JSON.stringify(o)); }
  function show(t) { document.getElementById("m").textContent = t; }

  Office.onReady(async function () {
    try {
      var pca = new msal.PublicClientApplication({
        auth: { clientId: cfg.clientId, authority: cfg.authority, redirectUri: redirectUri },
        cache: { cacheLocation: "localStorage" }
      });
      await pca.initialize();
      // Scopes överlever redirecten via sessionStorage (URL:en tappar sin query).
      var qs = new URLSearchParams(window.location.search).get("scopes");
      if (qs) sessionStorage.setItem("inlook.dlg.scopes", qs);
      var scopes = (sessionStorage.getItem("inlook.dlg.scopes") || "User.Read").split(" ");
      var res = await pca.handleRedirectPromise();
      if (!res) {
        var acc = pca.getAllAccounts()[0];
        if (acc) {
          try { res = await pca.acquireTokenSilent({ scopes: scopes, account: acc }); } catch (e) { res = null; }
        }
        if (!res) { show("Skickar till Microsoft…"); await pca.acquireTokenRedirect({ scopes: scopes, account: acc }); return; }
      }
      sessionStorage.removeItem("inlook.dlg.scopes");
      send({ ok: true, token: res.accessToken, exp: res.expiresOn ? +new Date(res.expiresOn) : 0 });
    } catch (e) {
      send({ ok: false, error: (e.errorCode ? e.errorCode + ": " : "") + (e.errorMessage || e.message || String(e)) });
    }
  });
})();
