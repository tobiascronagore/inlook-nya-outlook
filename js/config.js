// Offentlig konfiguration. Klient-ID och katalog (tenant) är inte hemligheter;
// ingen nyckel eller lösenord får någonsin läggas här.
window.INLOOK_CONFIG = {
  clientId: "21d98a1b-265e-4100-820b-bb1cfde4a1b0",
  authority: "https://login.microsoftonline.com/skandglas.se",
  // Mapp i användarens OneDrive där JSON-filerna ligger.
  driveFolder: "Inlook",
  // NAA ger bridge-fel i nya Outlook hos oss; dialog-vägen är standard.
  useNaa: false
};
