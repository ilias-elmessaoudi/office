// KXD Office A: Uebersicht + Editor (ONLYOFFICE-Kern aus ranuts/document im iframe, Einbettungs-Schnittstelle).
// Gespeichert wird immer im echten Office-Format; nach jeder Aenderung automatisch im Geraet.
const $ = (s) => document.querySelector(s);
const ARTEN = {
  docx: { name: 'KXD Text', klasse: 'text', symbol: 'T', neu: 'Dokument' },
  xlsx: { name: 'KXD Tabelle', klasse: 'tabelle', symbol: '#', neu: 'Tabelle' },
  pptx: { name: 'KXD Folien', klasse: 'folien', symbol: '▶', neu: 'Präsentation' },
};
const artVon = (name) => {
  const ext = name.split('.').pop().toLowerCase();
  if (['xlsx', 'xls', 'ods', 'csv'].includes(ext)) return 'xlsx';
  if (['pptx', 'ppt', 'odp'].includes(ext)) return 'pptx';
  if (ext === 'pdf') return 'pdf';
  return 'docx';
};
// Alte/fremde Formate werden beim Speichern ins passende Office-Format umgewandelt.
const zielEndung = (art) => (art === 'pdf' ? 'pdf' : art);
// In den Einzel-Apps (text/, tabelle/, folien/) gesetzt: zeigt nur diese Art. In der Gesamt-App leer.
const NUR_ART = window.KXD_ART || null;

let aktuell = null;          // { id, name, art }
let editorBereit = null;     // Promise, erfuellt bei document:ready
let wartend = new Map();     // Befehls-ID -> { ok, fehler }
let speicherTimer = null;
let ungespeichert = false;

function meldung(text, dauer = 2600) {
  const m = $('#meldung');
  m.textContent = text; m.hidden = false;
  clearTimeout(meldung.t); meldung.t = setTimeout(() => (m.hidden = true), dauer);
}
function datum(ms) {
  const d = new Date(ms), heute = new Date();
  const zeit = d.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' });
  return d.toDateString() === heute.toDateString() ? 'Heute, ' + zeit : d.toLocaleDateString('de-DE') + ', ' + zeit;
}

// ---------- Uebersicht ----------
async function liste() {
  const filter = $('#suche').value.trim().toLowerCase();
  const alle = (await Speicher.alle()).filter((e) => e.name.toLowerCase().includes(filter) && (!NUR_ART || e.art === NUR_ART));
  const ziel = $('#liste');
  ziel.replaceChildren();
  $('#leer').hidden = alle.length > 0 || !!filter;
  for (const e of alle) {
    const a = ARTEN[e.art] || { klasse: 'pdf', symbol: 'P' };
    const zeile = document.createElement('div');
    zeile.className = 'zeile';
    zeile.innerHTML = `<span class="mini ${a.klasse}">${a.symbol}</span><span class="titel"></span><span class="zeit">${datum(e.geaendert)}</span><button class="mehr" title="Löschen">✕</button>`;
    zeile.querySelector('.titel').textContent = e.name;
    zeile.onclick = () => bearbeiten(e.id);
    zeile.querySelector('.mehr').onclick = async (ev) => {
      ev.stopPropagation();
      if (confirm(`„${e.name}“ aus der Übersicht löschen?`)) { await Speicher.loeschen(e.id); liste(); }
    };
    ziel.append(zeile);
  }
}

async function neu(art) {
  const res = await fetch(`kxd/vorlagen/leer.${art}`);
  const name = `Neues ${ARTEN[art].neu}.${art}`;
  const id = crypto.randomUUID();
  await Speicher.sichern({ id, name, art, datei: await res.blob() });
  bearbeiten(id);
}

async function dateienAufnehmen(dateien) {
  let letzte;
  for (const f of dateien) {
    letzte = crypto.randomUUID();
    await Speicher.sichern({ id: letzte, name: f.name, art: artVon(f.name), datei: f });
  }
  if (dateien.length === 1) bearbeiten(letzte); else liste();
}

// Über Windows „Öffnen mit“ (installierte App) übergebene Dateien. Der Griff auf die Datei wird gemerkt,
// damit „Speichern“ wie in Word zurück in die Originaldatei schreibt (z. B. im Download-Ordner).
async function griffeAufnehmen(griffe) {
  let letzte;
  const alle = await Speicher.alle();
  for (const griff of griffe) {
    const f = await griff.getFile();
    let vorhanden = null;
    for (const e of alle) if (e.griff && (await e.griff.isSameEntry(griff).catch(() => false))) { vorhanden = e; break; }
    if (vorhanden) {
      letzte = vorhanden.id;
      // Datei wurde außerhalb geändert (z. B. in Word)? Dann deren Stand übernehmen.
      if (f.lastModified > (vorhanden.dateiStand || 0)) await Speicher.sichern({ ...vorhanden, datei: f, dateiStand: f.lastModified, griff });
    } else {
      letzte = crypto.randomUUID();
      await Speicher.sichern({ id: letzte, name: f.name, art: artVon(f.name), datei: f, griff, dateiStand: f.lastModified });
    }
  }
  if (griffe.length === 1) bearbeiten(letzte); else liste();
}
if ('launchQueue' in window) {
  launchQueue.setConsumer((start) => { if (start.files?.length) griffeAufnehmen(start.files); });
}

// Über die KXD-Brücke (Windows „Öffnen mit“, siehe projekte/office/kxd-oeffnen) übergebene Datei:
// ?lokal=<Kennung> – die Brücke auf diesem PC liefert die Datei und nimmt Gespeichertes zurück.
const BRUECKE = 'http://127.0.0.1:47321/datei/';
async function brueckeAufnehmen(kennung) {
  const r = await fetch(BRUECKE + kennung, { targetAddressSpace: 'loopback' });
  if (!r.ok) throw new Error((await r.json().catch(() => ({}))).fehler || 'Brücke antwortet nicht');
  const name = decodeURIComponent(r.headers.get('X-Dateiname'));
  const pfad = decodeURIComponent(r.headers.get('X-Pfad'));
  const stand = +r.headers.get('X-Geaendert');
  const datei = await r.blob();
  const vorhanden = (await Speicher.alle()).find((e) => e.lokal?.pfad === pfad);
  let id;
  if (vorhanden) {
    id = vorhanden.id;
    const neu = { ...vorhanden, lokal: { kennung, pfad } };
    // Datei wurde außerhalb geändert (z. B. in Word)? Dann deren Stand übernehmen.
    if (stand > (vorhanden.dateiStand || 0)) Object.assign(neu, { datei, dateiStand: stand });
    await Speicher.sichern(neu);
  } else {
    id = crypto.randomUUID();
    await Speicher.sichern({ id, name, art: artVon(name), datei, lokal: { kennung, pfad }, dateiStand: stand });
  }
  history.replaceState(null, '', location.pathname);
  bearbeiten(id);
}

// Zurück in die Originaldatei schreiben – nur wenn das Format gleich bleibt (aus .doc wird kein .doc).
async function inOriginalSchreiben(e, nachfragen) {
  if (e?.lokal) {
    if (!e.lokal.pfad.toLowerCase().endsWith('.' + zielEndung(e.art))) return false;
    const r = await fetch(BRUECKE + e.lokal.kennung, { method: 'PUT', body: e.datei, targetAddressSpace: 'loopback' });
    if (!r.ok) {
      meldung('In die Originaldatei konnte nicht gespeichert werden (Datei bitte erneut über „Öffnen mit“ öffnen). In KXD Office ist alles gespeichert.', 7000);
      return false;
    }
    const jetzt = await Speicher.holen(e.id);
    if (jetzt) await Speicher.uebernehmen({ ...jetzt, dateiStand: (await r.json()).geaendert });
    return true;
  }
  const griff = e?.griff;
  if (!griff || !griff.name.toLowerCase().endsWith('.' + zielEndung(e.art))) return false;
  let recht = await griff.queryPermission({ mode: 'readwrite' });
  if (recht === 'prompt' && nachfragen) recht = await griff.requestPermission({ mode: 'readwrite' });
  if (recht !== 'granted') { $('#original').hidden = false; return false; }
  const ziel = await griff.createWritable();
  await ziel.write(e.datei);
  await ziel.close();
  const stand = (await griff.getFile()).lastModified;
  const jetzt = await Speicher.holen(e.id);
  if (jetzt) await Speicher.uebernehmen({ ...jetzt, dateiStand: stand });
  $('#original').hidden = true;
  return true;
}

// ---------- Editor (iframe) ----------
function befehl(type, payload = {}) {
  const id = crypto.randomUUID();
  $('#editor').contentWindow.postMessage({ id, type, payload }, location.origin);
  return new Promise((ok, fehler) => wartend.set(id, { ok, fehler }));
}

addEventListener('message', (e) => {
  if (e.origin !== location.origin) return;
  const { id, type, payload } = e.data || {};
  if (!type?.startsWith('document:')) return;
  if (type === 'document:ready') editorBereit?.los();
  if (type === 'document:dirty-changed') {
    ungespeichert = payload.dirty;
    if (payload.dirty) { status('Nicht gespeichert'); planeSpeichern(); }
  }
  const w = id && wartend.get(id);
  if (w) {
    wartend.delete(id);
    type === 'document:error' ? w.fehler(new Error(payload?.message || 'Fehler')) : w.ok(payload);
  }
});

function status(t) { $('#status').textContent = t; }
function planeSpeichern() {
  clearTimeout(speicherTimer);
  speicherTimer = setTimeout(speichern, 2500);
}

async function speichern() {
  if (!aktuell || aktuell.art === 'pdf') return;
  clearTimeout(speicherTimer);
  status('Speichert…');
  try {
    const p = await befehl('document:save', { targetExt: zielEndung(aktuell.art).toUpperCase() });
    const name = aktuell.name.replace(/\.[^.]+$/, '') + '.' + zielEndung(aktuell.art);
    if (name !== aktuell.name) { aktuell.name = name; $('#name').value = name; }
    const vorher = await Speicher.holen(aktuell.id);
    await Speicher.sichern({ ...vorher, name: aktuell.name, datei: p.file });
    const inDatei = await inOriginalSchreiben(await Speicher.holen(aktuell.id), false).catch(() => false);
    status(p.dirty ? 'Nicht gespeichert' : inDatei ? 'In Datei gespeichert' : 'Gespeichert');
    if (p.dirty) planeSpeichern();
  } catch (err) {
    status('Speichern fehlgeschlagen');
    meldung('Speichern fehlgeschlagen: ' + err.message, 5000);
  }
}

async function bearbeiten(id) {
  const e = await Speicher.holen(id);
  if (!e) return;
  aktuell = { id: e.id, name: e.name, art: e.art };
  $('#original').hidden = true;
  $('#start').hidden = true; $('#bearbeiten').hidden = false;
  $('#name').value = e.name;
  document.title = e.name + ' – ' + (ARTEN[e.art]?.name || 'KXD Office A');
  status('Lädt…');
  editorBereit = (() => { let los; const p = new Promise((r) => (los = r)); p.los = los; return p; })();
  const rahmen = $('#editor');
  rahmen.src = 'editor.html?embed=1&locale=de-DE';
  if (history.state?.id !== id) history.pushState({ id }, '', location.pathname + '#' + id);
  // „document:ready“ kommt manchmal, bevor wir lauschen – darum nachfragen, bis der Editor antwortet.
  rahmen.onload = () => {
    const frage = setInterval(() => {
      if (!aktuell || rahmen.src.endsWith('about:blank')) return clearInterval(frage);
      befehl('document:get-state').then(() => { clearInterval(frage); editorBereit.los(); });
    }, 400);
    editorBereit.then(() => clearInterval(frage));
  };
  await editorBereit;
  await befehl('document:open-buffer', { fileName: e.name, buffer: await e.datei.arrayBuffer(), readonly: false });
  status('Gespeichert');
}

async function schliessen() {
  if (ungespeichert) await speichern();
  aktuell = null; ungespeichert = false;
  $('#editor').src = 'about:blank';
  $('#bearbeiten').hidden = true; $('#start').hidden = false;
  document.title = NUR_ART ? ARTEN[NUR_ART].name : 'KXD Office A';
  liste();
}

async function herunterladen(ext) {
  status(ext === 'pdf' ? 'PDF wird erstellt…' : 'Wird vorbereitet…');
  try {
    const p = await befehl('document:save', { targetExt: ext.toUpperCase() });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(p.file);
    a.download = aktuell.name.replace(/\.[^.]+$/, '') + '.' + ext;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 10000);
    if (ext !== 'pdf' && !p.dirty) { await Speicher.sichern({ ...(await Speicher.holen(aktuell.id)), datei: p.file }); }
    status('Gespeichert');
  } catch (err) { meldung('Fehlgeschlagen: ' + err.message, 5000); status(''); }
}

// ---------- Bedienung ----------
document.querySelectorAll('[data-neu]').forEach((k) => {
  k.onclick = () => neu(k.dataset.neu);
  if (NUR_ART && k.dataset.neu !== NUR_ART) k.hidden = true;
});
$('#original').onclick = async () => {
  try { if (await inOriginalSchreiben(await Speicher.holen(aktuell.id), true)) { status('In Datei gespeichert'); meldung('In der Originaldatei gespeichert'); } }
  catch (err) { meldung('Konnte nicht in die Datei schreiben: ' + err.message, 5000); }
};
$('#oeffnen-knopf').onclick = () => $('#datei').click();
$('#datei').onchange = (e) => { dateienAufnehmen([...e.target.files]); e.target.value = ''; };
$('#suche').oninput = liste;
$('#zurueck').onclick = () => history.back();
$('#herunterladen').onclick = () => herunterladen(zielEndung(aktuell.art));
$('#als-pdf').onclick = () => herunterladen('pdf');
$('#name').onchange = async () => {
  let n = $('#name').value.trim() || aktuell.name;
  if (!/\.[a-z0-9]+$/i.test(n)) n += '.' + zielEndung(aktuell.art);
  aktuell.name = n; $('#name').value = n;
  const e = await Speicher.holen(aktuell.id);
  await Speicher.sichern({ ...e, name: n });
};
addEventListener('popstate', () => { if (aktuell) schliessen(); });
addEventListener('keydown', (e) => {
  if (!aktuell && (e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'o') { e.preventDefault(); $('#datei').click(); }
});
addEventListener('beforeunload', (e) => { if (ungespeichert) { speichern(); e.preventDefault(); } });
document.addEventListener('dragover', (e) => e.preventDefault());
document.addEventListener('drop', (e) => {
  e.preventDefault();
  if (!aktuell && e.dataTransfer.files.length) dateienAufnehmen([...e.dataTransfer.files]);
});

// ---------- Abgleich (Heft-Konto) ----------
function wolkeZeigen(z) {
  const w = $('#wolke');
  w.classList.toggle('an', z.angemeldet && !z.fehler);
  w.classList.toggle('fehler', !!z.fehler);
  w.textContent = !z.angemeldet ? '☁ Anmelden' : z.fehler ? '☁ Fehler' : z.laeuft ? '☁ Gleicht ab…' : '☁ Abgeglichen';
  w.title = z.fehler ? 'Abgleich fehlgeschlagen: ' + z.fehler : z.angemeldet ? 'Angemeldet als ' + z.email : 'Mit dem Heft-Konto anmelden';
}
addEventListener('kxd-abgleich', (e) => wolkeZeigen(e.detail));
addEventListener('kxd-liste', () => { if (!aktuell) liste(); });
addEventListener('kxd-fern', (e) => {
  if (aktuell?.id === e.detail.id) meldung('Dieses Dokument wurde auf einem anderen Gerät geändert. Zur Übersicht und neu öffnen, um die neue Fassung zu sehen.', 8000);
});
$('#wolke').onclick = () => {
  const z = Abgleich.zustand, d = $('#anmeldung');
  $('#anmelde-fehler').hidden = true;
  $('#anmelde-felder').hidden = z.angemeldet;
  $('#anmelde-text').textContent = z.angemeldet
    ? `Angemeldet als ${z.email}.` + (z.fehler ? ` Letzter Fehler: ${z.fehler}` : z.letzter ? ` Zuletzt abgeglichen: ${datum(z.letzter)}.` : '')
    : 'Mit deinem Heft-Konto anmelden. Dann sind deine Dokumente auf PC, Laptop und iPad gleich.';
  $('#anmelde-los').textContent = z.angemeldet ? 'Abmelden' : 'Anmelden';
  d.showModal();
};
$('#anmelde-form').addEventListener('submit', async (e) => {
  if (e.submitter?.value !== 'los') return;
  e.preventDefault();
  try {
    if (Abgleich.zustand.angemeldet) { await Abgleich.abmelden(); meldung('Abgemeldet'); }
    else { await Abgleich.anmelden($('#mail').value.trim(), $('#passwort').value); meldung('Angemeldet – Abgleich läuft'); }
    $('#passwort').value = '';
    $('#anmeldung').close();
  } catch (err) { $('#anmelde-fehler').textContent = err.message; $('#anmelde-fehler').hidden = false; }
});
// Offline-Speicher nur fuer den eigenen Unterordner (sw.js neben index.html)
if ('serviceWorker' in navigator) addEventListener('load', () => navigator.serviceWorker.register('sw.js').catch((e) => console.warn('Offline-Speicher', e)));
Abgleich.starten().catch((e) => console.warn('Abgleich nicht verfügbar', e));

const lokal = new URLSearchParams(location.search).get('lokal');
if (lokal) brueckeAufnehmen(lokal).catch((e) => { meldung('Datei konnte nicht geladen werden: ' + e.message, 8000); liste(); });
const start = lokal ? '' : location.hash.slice(1);
start ? Speicher.holen(start).then((e) => (e ? bearbeiten(start) : liste())) : liste();
