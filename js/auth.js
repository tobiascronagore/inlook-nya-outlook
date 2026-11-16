// Inloggning via MSAL.js. I Outlook används NAA (nested app authentication):
// Outlook förmedlar token åt tillägget, ingen server och ingen hemlighet.
// Utanför Outlook (vanlig webbläsare) faller MSAL tillbaka på vanlig popup-
// inloggning, vilket är praktiskt för att prova Graph-delen separat.
(function () {
  var cfg = window.INLOOK_CONFIG;
  var pcaPromise = null;
  var account = null;

  function getPca() {
    if (!pcaPromise) {
      pcaPromise = msal.createNestablePublicClientApplication({
        auth: {
          clientId: cfg.clientId,
          authority: cfg.authority,
          supportsNestedAppAuth: true,
          redirectUri: window.location.origin + window.location.pathname.replace(/[^/]*$/, "")
        },
        cache: { cacheLocation: "localStorage" }
      });
    }
    return pcaPromise;
  }

  // Ger en token för EXAKT de scopes som anges. Medgivande (consent) frågas
  // först när scopet behövs, och tenantens regler avgör om användaren själv
  // får godkänna eller om administratör krävs (felkod visas då i klartext).
  async function getToken(scopes) {
    var pca = await getPca();
    var req = { scopes: scopes };
    if (account) req.account = account;
    var res;
    try {
      res = await pca.acquireTokenSilent(req);
    } catch (e) {
      // Ingen cachad token / medgivande saknas -> interaktivt.
      res = await pca.acquireTokenPopup(req);
    }
    account = res.account || account;
    return res.accessToken;
  }

  window.InlookAuth = { getToken: getToken, getPca: getPca };
})();
