# OneDrive-Dokumentensuche – Einrichtungsanleitung

## Was die Anwendung kann

- Anmeldung mit einem persönlichen Microsoft-Konto (OneDrive Personal).
- Ordner und `.docx`-Dateien im OneDrive anzeigen.
- Eine Datei auswählen und als Favorit speichern.
- Favoriten im persönlichen OneDrive-App-Speicher (`Apps/OneDrive-Dokumentensuche`) ablegen. Dadurch stehen sie nach erneuter Anmeldung auch auf anderen Geräten zur Verfügung.
- In der ausgewählten Word-Datei suchen: UND (alle Begriffe) oder ODER (mindestens ein Begriff).
- Fundstelle mit Satz davor, Satz danach und zuletzt erkannter Word-Überschrift anzeigen.
- Dokumenteninhalt wird im Browser verarbeitet. Es wird kein eigener Suchserver verwendet.

## Voraussetzungen

- Ein persönliches Microsoft-Konto mit OneDrive.
- Ein Microsoft-Entra-App-Registrierungseintrag für persönliche Microsoft-Konten.
- Ein HTTPS-Webhosting. Für die PWA-Installation und den Service Worker ist HTTPS erforderlich (localhost ist für lokale Entwicklung eine Ausnahme).
- Ein aktueller Browser. Für iPhone: Safari.

## Teil A – Microsoft-Anwendung registrieren

1. Öffne https://entra.microsoft.com/ und melde dich mit dem Microsoft-Konto an.
2. Öffne **Identität > Anwendungen > App-Registrierungen** (die Bezeichnungen können je nach Portalansicht leicht abweichen).
3. Wähle **Neue Registrierung**.
4. Name: `OneDrive-Dokumentensuche`.
5. Unterstützte Kontotypen: **Persönliche Microsoft-Konten בלבד** / **Personal Microsoft accounts only**.
6. Redirect-URI zunächst noch nicht oder als Plattform **Single-page application (SPA)** mit der späteren Web-Adresse eintragen. Beispiel: `https://DEIN-BENUTZER.github.io/onedrive-dokumentensuche/`
7. Registrierung abschließen und die **Anwendungs-ID (Client)** kopieren. Kein Client-Geheimnis erstellen; eine Browser-App darf kein Client-Secret enthalten.
8. Unter **Authentifizierung** prüfen, dass die Web-Adresse als Plattform **Single-page application** eingetragen ist. Die Redirect-URI muss exakt mit der URL in `config.js` übereinstimmen.
9. Unter **API-Berechtigungen > Berechtigung hinzufügen > Microsoft Graph > Delegierte Berechtigungen** hinzufügen:
   - `User.Read`
   - `Files.Read`
   - `Files.ReadWrite.AppFolder`
10. Speichern. Beim ersten Anmelden wird die Zustimmung für diese Berechtigungen abgefragt.

Hinweis: Das Portal kann sich optisch ändern. Verwende ausdrücklich persönliche Microsoft-Konten und die Plattform SPA. Für ein privates Konto ist der Authority-Endpunkt `https://login.microsoftonline.com/consumers`.

## Teil B – Anwendung konfigurieren

1. Entpacke das ZIP auf deinem Computer.
2. Öffne den Ordner `onedrive-dokumentensuche`.
3. Benenne `config.example.js` in `config.js` um (die ZIP enthält bereits eine Beispiel-`config.js`; du kannst diese direkt bearbeiten).
4. Öffne `config.js` mit Editor/Notepad.
5. Ersetze `HIER-DEINE-ANWENDUNGS-ID-EINTRAGEN` durch die kopierte Anwendungs-ID.
6. Trage später deine endgültige HTTPS-Adresse als Redirect-URI im Microsoft-Portal ein. `config.js` bildet die Redirect-URI automatisch aus der aktuellen Seitenadresse.

Beispiel:
```js
window.APP_CONFIG = {
  clientId: "00000000-0000-0000-0000-000000000000",
  authority: "https://login.microsoftonline.com/consumers",
  redirectUri: window.location.origin + window.location.pathname,
  postLogoutRedirectUri: window.location.origin + window.location.pathname
};
```

## Teil C – Veröffentlichen mit GitHub Pages (kostenlose Möglichkeit)

1. Erstelle unter https://github.com/ ein Konto oder melde dich an.
2. Erstelle ein neues Repository, zum Beispiel `onedrive-dokumentensuche`.
3. Lade den **Inhalt** des entpackten Ordners in das Repository hoch. `index.html` muss im veröffentlichten Stammverzeichnis liegen.
4. Öffne im Repository **Settings > Pages**.
5. Bei **Build and deployment** als Quelle **Deploy from a branch** wählen.
6. Branch `main` und Ordner `/ (root)` wählen und speichern.
7. GitHub zeigt nach der Veröffentlichung eine HTTPS-Adresse an, z. B. `https://DEIN-NAME.github.io/onedrive-dokumentensuche/`.
8. Trage diese exakte Adresse im Microsoft-Portal unter **App-Registrierungen > Authentifizierung > Single-page application > Redirect-URIs** ein.
9. Öffne `config.js` im Repository und prüfe, dass die Client-ID stimmt. Änderungen speichern.
10. Warte auf die Veröffentlichung und öffne die HTTPS-Adresse.

**Wichtig:** `config.js` enthält nur die öffentliche Client-ID, kein Passwort und kein Client-Geheimnis. Die App darf niemals ein Client-Secret im Browsercode enthalten. Ein öffentliches Repository macht den Quellcode öffentlich. Wenn du das nicht möchtest, nutze ein privates Repository mit einem Hosting, das private Deployments unterstützt, oder einen anderen HTTPS-Webhoster.

## Teil D – Auf dem iPhone installieren

1. Öffne die veröffentlichte HTTPS-Adresse in **Safari**.
2. Melde dich mit deinem persönlichen Microsoft-Konto an.
3. Tippe auf **Teilen** (Quadrat mit Pfeil nach oben).
4. Wähle **Zum Home-Bildschirm**.
5. Aktiviere **Als Web-App öffnen**, falls diese Option angezeigt wird.
6. Tippe auf **Hinzufügen**.
7. Starte die App über das neue Symbol auf dem Home-Bildschirm.

Auf dem PC kann die Seite in Edge oder Chrome über das Installieren-Symbol in der Adressleiste bzw. über das Browser-Menü als App installiert werden.

## Teil E – Benutzung

1. App öffnen und **Mit Microsoft anmelden** wählen.
2. Den Microsoft-Zugriff bestätigen.
3. Unter **OneDrive durchsuchen** Ordner öffnen und eine `.docx`-Datei auswählen.
4. Mit dem Stern die Datei als Favorit speichern. Favoriten werden in deinem OneDrive-App-Ordner synchronisiert.
5. Unter **Favoriten** lässt sich eine gespeicherte Datei später wieder auswählen.
6. Suchbegriffe eingeben. Trenne mehrere Begriffe durch Komma, Semikolon oder Zeilenumbruch.
7. UND wählen, wenn alle Begriffe im selben Satz vorkommen müssen; ODER, wenn mindestens einer genügt.
8. **Suchen** drücken. Jeder Treffer zeigt die Abschnittsüberschrift, den vorherigen Satz, den gefundenen Satz und den nachfolgenden Satz.

## Hinweise und Grenzen

- Die aktuelle Version durchsucht jeweils **eine ausgewählte DOCX-Datei**. Sie ist keine Volltextsuche über das gesamte OneDrive.
- Unterstützt werden `.docx`-Dateien. Alte `.doc`-Dateien, PDF-Dateien und passwortgeschützte Dokumente werden nicht verarbeitet.
- Die Erkennung der Überschrift nutzt Word-Überschriftenstile, die in der DOCX-Struktur als Überschriften vorliegen. Wenn ein Dokument Überschriften nur optisch fett/größer formatiert, kann die Zuordnung fehlen.
- Die Satztrennung ist automatisiert. Abkürzungen, Aufzählungen und Tabellen können zu ungewöhnlichen Satzgrenzen führen.
- Tabelleninhalte werden abhängig von der DOCX-Struktur ausgelesen; Textfelder, Kommentare, Kopf-/Fußzeilen und eingebettete Objekte sind nicht zuverlässig abgedeckt.
- Seitenzahlen werden nicht angezeigt, weil eine DOCX-Datei ohne Word-Layout-Engine keine verlässlichen Seitenpositionen liefert.
- Sehr große Dateien können je nach iPhone-Speicher und Browser langsamer verarbeitet werden.
- Bei Umbenennen oder Verschieben bleibt die OneDrive-Element-ID normalerweise erhalten. Beim Löschen oder Verschieben in ein anderes Konto kann ein Favorit ungültig werden.
- Favoriten benötigen `Files.ReadWrite.AppFolder`. Das ist eine Schreibberechtigung, aber nur für den von dieser App verwendeten App-Ordner. Dokumentdateien werden nicht verändert.
- Bei Problemen zuerst abmelden und erneut anmelden. Prüfe außerdem, ob die Redirect-URI exakt mit der veröffentlichten Adresse übereinstimmt.

## Datenschutz und Sicherheit

Die Anwendung enthält keinen eigenen Server und keine Datenbank. Die Microsoft-Anmeldung erfolgt mit MSAL im Browser. Die App fragt delegierte Berechtigungen ab und handelt im Kontext des angemeldeten Benutzers. Die Dokumentdatei wird über Microsoft Graph geladen und im Browser ausgewertet. Favoriten werden als `favorites.json` im OneDrive-App-Ordner gespeichert.

Der Quellcode ist ein funktionsfähiger Ausgangspunkt, aber vor produktiver Nutzung mit vertraulichen Unterlagen sollte er geprüft und auf die eigene Umgebung getestet werden. Veröffentliche keine vertraulichen Dokumente oder Zugangsdaten im Webhosting.
