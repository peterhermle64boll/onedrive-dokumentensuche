(() => {
  "use strict";

  const $ = (id) => document.getElementById(id);
  const config = window.APP_CONFIG || {};
  const missingConfig = !config.clientId || config.clientId.includes("HIER-DEINE");
  $("setupNotice").classList.toggle("hidden", !missingConfig);

  const scopes = ["User.Read", "Files.Read", "Files.ReadWrite.AppFolder"];
  let msalApp = null;
  let account = null;
  let currentFolder = null;
  let folderStack = [];
  let selected = null;
  let favorites = [];
  let searchMode = "AND";
  let currentTab = "browse";

  if (!missingConfig && window.msal) {
    msalApp = new msal.PublicClientApplication({
      auth: {
        clientId: config.clientId,
        authority: config.authority || "https://login.microsoftonline.com/consumers",
        redirectUri: config.redirectUri || window.location.href,
        postLogoutRedirectUri: config.postLogoutRedirectUri || window.location.href
      },
      cache: { cacheLocation: "sessionStorage", storeAuthStateInCookie: false }
    });
  }

  function status(message, kind = "") {
    $("status").textContent = message;
    $("status").className = "status " + kind;
  }
  function safeText(value) { return String(value ?? ""); }
  function escapeHtml(value) {
    return safeText(value).replace(/[&<>"']/g, c => ({ "&":"&amp;", "<":"&lt;", ">":"&gt;", '"':"&quot;", "'":"&#39;" }[c]));
  }
  function isDocx(item) {
    return item.file && (item.name || "").toLowerCase().endsWith(".docx");
  }

  async function login() {
    if (missingConfig) { alert("Bitte zuerst config.js mit deiner Microsoft-Anwendungs-ID konfigurieren."); return; }
    await msalApp.loginRedirect({ scopes });
  }

  async function logout() {
    if (!msalApp || !account) return;
    await msalApp.logoutRedirect({ account, postLogoutRedirectUri: config.postLogoutRedirectUri });
  }

  async function getToken() {
    const request = { scopes, account };
    try {
      const result = await msalApp.acquireTokenSilent(request);
      return result.accessToken;
    } catch (e) {
      await msalApp.acquireTokenRedirect(request);
      throw e;
    }
  }

  async function graph(path, options = {}) {
    const token = await getToken();
    const response = await fetch("https://graph.microsoft.com/v1.0" + path, {
      ...options,
      headers: { Authorization: "Bearer " + token, ...(options.headers || {}) }
    });
    if (!response.ok) {
      let detail = "";
      try { detail = (await response.json()).error?.message || ""; } catch {}
      throw new Error(`Microsoft Graph (${response.status}): ${detail || response.statusText}`);
    }
    return response;
  }

  async function initialize() {
    if (!msalApp) {
      $("loginBtn").disabled = true;
      return;
    }
    try {
      await msalApp.initialize();
      const redirectResult = await msalApp.handleRedirectPromise();
      account = redirectResult?.account || msalApp.getAllAccounts()[0] || null;
      if (account) {
        msalApp.setActiveAccount(account);
        updateSignedIn();
        await Promise.all([loadFolder(null), loadFavorites()]);
      }
    } catch (e) {
      status("Anmeldung konnte nicht abgeschlossen werden: " + e.message, "error");
    }
  }

  function updateSignedIn() {
    $("accountLabel").textContent = account?.username || "Angemeldet";
    $("loginBtn").classList.toggle("hidden", !!account);
    $("logoutBtn").classList.toggle("hidden", !account);
    $("refreshBtn").disabled = !account;
  }

  async function loadFolder(folderId) {
    if (!account) return;
    const path = folderId
      ? `/me/drive/items/${encodeURIComponent(folderId)}/children?$select=id,name,folder,file,size,lastModifiedDateTime&$top=200`
      : `/me/drive/root/children?$select=id,name,folder,file,size,lastModifiedDateTime&$top=200`;
    try {
      status("Lade OneDrive-Dateien …");
      const data = await (await graph(path)).json();
      currentFolder = folderId;
      renderFiles(data.value || []);
      renderBreadcrumbs();
      status("");
    } catch (e) { status("Dateiliste konnte nicht geladen werden: " + e.message, "error"); }
  }

  function renderBreadcrumbs() {
    const crumbs = [{ name: "OneDrive", id: null }, ...folderStack];
    $("breadcrumbs").replaceChildren();
    crumbs.forEach((crumb, index) => {
      const b = document.createElement("button");
      b.className = "crumb";
      b.textContent = crumb.name;
      b.addEventListener("click", () => {
        folderStack = folderStack.slice(0, index);
        loadFolder(crumb.id);
      });
      $("breadcrumbs").append(b);
      if (index < crumbs.length - 1) {
        const sep = document.createElement("span"); sep.textContent = "›"; sep.className = "muted"; $("breadcrumbs").append(sep);
      }
    });
  }

  function renderFiles(items) {
    const root = $("fileList"); root.replaceChildren();
    const sorted = [...items].sort((a,b) => {
      const af = !!a.folder, bf = !!b.folder;
      return af === bf ? a.name.localeCompare(b.name, "de") : (af ? -1 : 1);
    });
    const usable = sorted.filter(x => x.folder || isDocx(x));
    if (!usable.length) { root.innerHTML = '<div class="empty">Hier wurden keine Ordner oder DOCX-Dateien gefunden.</div>'; return; }
    usable.forEach(item => {
      const row = document.createElement("div"); row.className = "file-row";
      const main = document.createElement("button"); main.className = "file-main";
      main.innerHTML = `<span class="file-icon">${item.folder ? "📁" : "W"}</span><span class="file-name">${escapeHtml(item.name)}<small>${item.folder ? "Ordner" : "Word-Dokument · " + formatBytes(item.size)}</small></span><span class="chevron">${item.folder ? "›" : "Auswählen"}</span>`;
      main.addEventListener("click", () => item.folder ? enterFolder(item) : selectFile(item));
      row.append(main);
      if (!item.folder) {
        const star = document.createElement("button"); star.className = "star-btn"; star.textContent = favorites.some(f => f.id === item.id) ? "★" : "☆";
        star.title = "Favorit umschalten"; star.addEventListener("click", () => toggleFavorite(item)); row.append(star);
      }
      root.append(row);
    });
  }

  async function enterFolder(item) {
    folderStack.push({ name: item.name, id: item.id });
    await loadFolder(item.id);
  }

  async function selectFile(item) {
    selected = { id: item.id, name: item.name, webUrl: item.webUrl || "", size: item.size };
    $("selectedFile").classList.remove("hidden");
    $("selectedFile").innerHTML = `<span>Ausgewählt:</span> <strong>${escapeHtml(item.name)}</strong> <button id="favSelected" class="star-btn" aria-label="Favorit umschalten">${favorites.some(f => f.id === item.id) ? "★" : "☆"}</button>`;
    $("favSelected").addEventListener("click", () => toggleFavorite(item));
    $("queryInput").disabled = false; $("andBtn").disabled = false; $("orBtn").disabled = false; $("searchBtn").disabled = false;
    $("results").replaceChildren();
    status("Datei ausgewählt. Du kannst jetzt suchen.", "success");
  }

  async function loadFavorites() {
    try {
      const response = await graph("/me/drive/special/approot:/favorites.json:/content");
      favorites = await response.json();
      if (!Array.isArray(favorites)) favorites = [];
    } catch (e) {
      if (String(e.message).includes("(404)")) favorites = [];
      else { favorites = []; status("Favoriten konnten nicht geladen werden: " + e.message, "error"); }
    }
    renderFavorites();
  }

  async function saveFavorites() {
    const blob = new Blob([JSON.stringify(favorites, null, 2)], { type: "application/json" });
    await graph("/me/drive/special/approot:/favorites.json:/content", {
      method: "PUT", headers: { "Content-Type": "application/json" }, body: blob
    });
    renderFavorites();
  }

  async function toggleFavorite(item) {
    const exists = favorites.some(f => f.id === item.id);
    favorites = exists
      ? favorites.filter(f => f.id !== item.id)
      : [...favorites, { id: item.id, name: item.name, webUrl: item.webUrl || "", addedAt: new Date().toISOString() }];
    try {
      await saveFavorites();
      if (selected?.id === item.id) $("favSelected").textContent = exists ? "☆" : "★";
      if (currentTab === "browse") loadFolder(currentFolder);
      status(exists ? "Favorit entfernt." : "Favorit gespeichert und synchronisiert.", "success");
    } catch (e) {
      status("Favorit konnte nicht gespeichert werden: " + e.message, "error");
    }
  }

  function renderFavorites() {
    const root = $("favoriteList"); root.replaceChildren();
    if (!favorites.length) { root.innerHTML = '<div class="empty">Noch keine Favoriten gespeichert.</div>'; return; }
    favorites.forEach(f => {
      const row = document.createElement("div"); row.className = "file-row";
      const main = document.createElement("button"); main.className = "file-main";
      main.innerHTML = `<span class="file-icon">W</span><span class="file-name">${escapeHtml(f.name)}<small>OneDrive-Favorit</small></span><span class="chevron">Auswählen</span>`;
      main.addEventListener("click", async () => {
        try {
          const item = await (await graph(`/me/drive/items/${encodeURIComponent(f.id)}?$select=id,name,size,webUrl`)).json();
          await selectFile(item);
          switchTab("browse");
        } catch (e) {
          status("Favorit ist nicht erreichbar (möglicherweise verschoben oder gelöscht): " + e.message, "error");
        }
      });
      const remove = document.createElement("button"); remove.className = "star-btn"; remove.textContent = "★"; remove.title = "Favorit entfernen";
      remove.addEventListener("click", () => toggleFavorite(f));
      row.append(main, remove); root.append(row);
    });
  }

  function switchTab(tab) {
    currentTab = tab;
    document.querySelectorAll(".tab").forEach(b => b.classList.toggle("active", b.dataset.tab === tab));
    $("browseView").classList.toggle("hidden", tab !== "browse");
    $("favoritesView").classList.toggle("hidden", tab !== "favorites");
    $("selectedFile").classList.toggle("hidden", !selected);
  }

  async function downloadDocx(itemId) {
    const response = await graph(`/me/drive/items/${encodeURIComponent(itemId)}/content`);
    return await response.arrayBuffer();
  }

  function normalize(s) {
    return String(s).toLocaleLowerCase("de-DE").normalize("NFC").replace(/\s+/g, " ").trim();
  }

  function splitSentences(text) {
    const cleaned = text.replace(/\s+/g, " ").trim();
    if (!cleaned) return [];
    if (window.Intl && Intl.Segmenter) {
      const segmenter = new Intl.Segmenter("de", { granularity: "sentence" });
      return [...segmenter.segment(cleaned)].map(x => x.segment.trim()).filter(Boolean);
    }
    return cleaned.match(/[^.!?]+[.!?]+["'»)]*|[^.!?]+$/g)?.map(x => x.trim()).filter(Boolean) || [cleaned];
  }

  function extractDocument(html) {
    const doc = new DOMParser().parseFromString(html, "text/html");
    const blocks = [];
    let heading = "Ohne zugeordnete Überschrift";
    const nodes = doc.body.querySelectorAll("h1,h2,h3,h4,h5,h6,p,li");
    nodes.forEach(node => {
      const text = node.textContent.replace(/\s+/g, " ").trim();
      if (!text) return;
      if (/^H[1-6]$/.test(node.tagName)) {
        heading = text;
      } else {
        splitSentences(text).forEach(sentence => blocks.push({ sentence, heading }));
      }
    });
    return blocks;
  }

  function highlight(text, terms) {
    const escapedTerms = terms.filter(Boolean).sort((a,b) => b.length-a.length).map(t => t.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"));
    if (!escapedTerms.length) return escapeHtml(text);
    const regex = new RegExp("(" + escapedTerms.join("|") + ")", "ig");
    return escapeHtml(text).replace(regex, "<mark>$1</mark>");
  }

  async function runSearch() {
    if (!selected) return;
    const terms = $("queryInput").value.split(/[,;\n]+/).map(x => x.trim()).filter(Boolean);
    if (!terms.length) { status("Bitte mindestens einen Suchbegriff eingeben.", "error"); return; }
    $("searchBtn").disabled = true; $("results").replaceChildren();
    status("Lade und analysiere das Word-Dokument …");
    try {
      const arrayBuffer = await downloadDocx(selected.id);
      const converted = await mammoth.convertToHtml({ arrayBuffer });
      const sentences = extractDocument(converted.value);
      const normalizedTerms = terms.map(normalize);
      const hits = [];
      sentences.forEach((entry, i) => {
        const haystack = normalize(entry.sentence);
        const matches = normalizedTerms.map(t => haystack.includes(t));
        const found = searchMode === "AND" ? matches.every(Boolean) : matches.some(Boolean);
        if (found) hits.push({ index: i, entry, before: sentences[i-1]?.sentence || "Kein vorheriger Satz vorhanden.", after: sentences[i+1]?.sentence || "Kein nachfolgender Satz vorhanden." });
      });
      renderResults(hits, terms);
      status(`${hits.length} Fundstelle(n) in ${sentences.length} Sätzen gefunden.`, hits.length ? "success" : "");
    } catch (e) {
      status("Dokument konnte nicht durchsucht werden: " + e.message, "error");
    } finally { $("searchBtn").disabled = false; }
  }

  function renderResults(hits, terms) {
    const root = $("results"); root.replaceChildren();
    if (!hits.length) { root.innerHTML = '<div class="empty">Keine passenden Sätze gefunden. Prüfe die Suchbegriffe oder wechsle zwischen UND und ODER.</div>'; return; }
    hits.forEach((hit, index) => {
      const card = document.createElement("article"); card.className = "hit-card";
      card.innerHTML = `<div class="hit-top"><span class="hit-number">Treffer ${index+1}</span><span class="heading-label">${escapeHtml(hit.entry.heading)}</span></div>
        <div class="context-label">Satz davor</div><p class="context">${escapeHtml(hit.before)}</p>
        <div class="context-label found-label">Gefundener Satz</div><p class="found">${highlight(hit.entry.sentence, terms)}</p>
        <div class="context-label">Satz danach</div><p class="context">${escapeHtml(hit.after)}</p>`;
      root.append(card);
    });
  }

  function formatBytes(bytes) {
    if (!bytes) return "Größe unbekannt";
    const units = ["B","KB","MB","GB"]; let i=0, n=bytes;
    while (n >= 1024 && i < units.length-1) { n/=1024; i++; }
    return `${n.toFixed(i ? 1 : 0)} ${units[i]}`;
  }

  $("loginBtn").addEventListener("click", login);
  $("logoutBtn").addEventListener("click", logout);
  $("refreshBtn").addEventListener("click", () => currentTab === "favorites" ? loadFavorites() : loadFolder(currentFolder));
  $("searchBtn").addEventListener("click", runSearch);
  $("andBtn").addEventListener("click", () => setMode("AND"));
  $("orBtn").addEventListener("click", () => setMode("OR"));
  document.querySelectorAll(".tab").forEach(b => b.addEventListener("click", () => switchTab(b.dataset.tab)));
  function setMode(mode) {
    searchMode = mode;
    $("andBtn").classList.toggle("selected", mode === "AND");
    $("orBtn").classList.toggle("selected", mode === "OR");
  }

  if ("serviceWorker" in navigator && location.protocol === "https:") {
    window.addEventListener("load", () => navigator.serviceWorker.register("./service-worker.js").catch(console.warn));
  }
  initialize();
})();
