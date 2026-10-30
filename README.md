# KXD Office A

KXD Text, KXD Tabelle und KXD Folien im Browser – speichert in echten Office-Formaten
(.docx, .xlsx, .pptx) und als PDF. Läuft komplett im Gerät, ohne Server, ohne Konto.

## Herkunft und Lizenz

Dieses Paket ist eine veränderte Fassung von **ranuts/document** (https://github.com/ranuts/document),
das wiederum die ONLYOFFICE-Editoren (sdkjs, web-apps, © Ascensio System SIA) und den
Umwandlungs-Baustein onlyoffice-x2t-wasm (CryptPad) enthält. Lizenz: GNU AGPL v3, siehe `LICENSE` und `NOTICE`.

Quellcode: der unveränderte Bau von ranuts/document (Commit e975654679d2cc3ef61d67d93fdce2839327676c), darüber:
- `index.html` und `kxd/` – eigene Startseite, Ablage im Gerät, Name und Icon (in diesem Repository im Klartext)
- `icons/document*` – eigenes Icon statt des ursprünglichen
- `manifest.json` – eigener App-Name
- `sdkjs/common/wasm/x2t/x2t.wasm.br` – unverändert, nur entpackt (GitHub Pages kann keine Brotli-Kopfzeile senden)
- `sw.js` – eigener, einfacher Offline-Speicher nur für diesen Unterordner (der ursprüngliche hängt fest an der Wurzel der Adresse)
- `assets/` – gebaut aus ranuts/document mit drei kleinen Änderungen für den Betrieb im Unterordner /office/
  (Pfad des Umwandlungs-Bausteins, Adresse des Offline-Speichers, eine Weiterleitung). Die Änderungen stehen als
  Patch in der Datei `office-a-unterordner.patch`.

ONLYOFFICE ist eine Marke von Ascensio System SIA. Dieses Projekt ist kein offizielles
ONLYOFFICE-Produkt und steht in keiner Verbindung dazu. Microsoft, Word, Excel und PowerPoint
sind Marken der Microsoft Corporation; dieses Projekt steht in keiner Verbindung zu Microsoft.
