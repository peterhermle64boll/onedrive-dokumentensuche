// 1. Diese Datei in "config.js" umbenennen.
// 2. CLIENT_ID durch die Anwendungs-ID (Client) aus Microsoft Entra ersetzen.
// Die App ist für persönliche Microsoft-Konten (Outlook/Hotmail/Live/OneDrive Personal) eingerichtet.
window.APP_CONFIG = {
  clientId: "e66ed6f7-fea2-4a4a-ab40-236c30dd9e14",
  authority: "https://login.microsoftonline.com/consumers",
  redirectUri: window.location.origin + window.location.pathname,
  postLogoutRedirectUri: window.location.origin + window.location.pathname
};
