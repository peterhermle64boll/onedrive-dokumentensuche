// 1. Diese Datei in "config.js" umbenennen.
// 2. CLIENT_ID durch die Anwendungs-ID (Client) aus Microsoft Entra ersetzen.
// Die App ist für persönliche Microsoft-Konten (Outlook/Hotmail/Live/OneDrive Personal) eingerichtet.
window.APP_CONFIG = {
  clientId: "HIER-DEINE-ANWENDUNGS-ID-EINTRAGEN",
  authority: "https://login.microsoftonline.com/consumers",
  redirectUri: window.location.origin + window.location.pathname,
  postLogoutRedirectUri: window.location.origin + window.location.pathname
};
