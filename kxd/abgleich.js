// Abgleich zwischen Geraeten ueber das Heft-Konto (Supabase-Projekt „KXD HEFT“).
// Gleiche Tabelle „eintraege“ und gleicher Speicher „dateien“ wie Heft, aber mit art = 'office'
// und Dateipfad <Nutzer>/office-<id>, damit sich Heft und KXD Office nicht in die Quere kommen.
// Bei Konflikten gewinnt die neuere Aenderung.
const Abgleich = (() => {
  const KONFIG = {
    url: 'https://dhtblosutrzawchscgpt.supabase.co',
    // oeffentlicher „anon“-Schluessel; geschuetzt wird ueber Anmeldung + Zeilenrechte
    key: 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImRodGJsb3N1dHJ6YXdjaHNjZ3B0Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3OTMxMjg0ODgsImV4cCI6MjEwODcwNDQ4OH0.UXYM6Op7_SkKm9oqh8zX4en8lT7ibAuSbKOpQB5SCPE',
  };
  const TABELLE = 'eintraege', EIMER = 'dateien', ART = 'office';
  const zustand = { angemeldet: false, email: null, uid: null, laeuft: false, letzter: null, fehler: null };
  let client = null, laeuft = null, nochmal = false;
  const pfad = (id) => `${zustand.uid}/office-${id}`;
  const melden = () => dispatchEvent(new CustomEvent('kxd-abgleich', { detail: { ...zustand } }));

  async function starten() {
    // gleicher Speicherschluessel wie Heft: wer dort angemeldet ist, ist es hier auch (gleiche Adresse)
    client = supabase.createClient(KONFIG.url, KONFIG.key, { auth: { persistSession: true, storageKey: 'heft-anmeldung', autoRefreshToken: true } });
    const { data } = await client.auth.getSession();
    sitzung(data.session);
    client.auth.onAuthStateChange((_e, s) => sitzung(s));
    let timer;
    addEventListener('kxd-lokal', () => { clearTimeout(timer); timer = setTimeout(abgleichen, 2500); });
    document.addEventListener('visibilitychange', () => { if (!document.hidden) abgleichen(); });
    addEventListener('online', abgleichen);
    setInterval(() => { if (!document.hidden) abgleichen(); }, 30000);
  }
  function sitzung(s) {
    const vorher = zustand.angemeldet;
    zustand.angemeldet = !!s; zustand.email = s?.user?.email || null; zustand.uid = s?.user?.id || null;
    melden();
    if (!vorher && s) abgleichen();
  }
  async function anmelden(email, passwort) {
    const { error } = await client.auth.signInWithPassword({ email, password: passwort });
    if (error) throw new Error(/invalid login/i.test(error.message) ? 'E-Mail oder Passwort stimmt nicht.' : error.message);
  }
  async function abmelden() { await client.auth.signOut(); await Speicher.metaSetzen('stand', null); }

  function abgleichen() {
    if (!client || !zustand.angemeldet || !navigator.onLine) return Promise.resolve();
    if (laeuft) { nochmal = true; return laeuft; }
    laeuft = (async () => {
      zustand.laeuft = true; zustand.fehler = null; melden();
      try {
        do { nochmal = false; await hochladen(); await herunterladen(); } while (nochmal);
        zustand.letzter = Date.now();
      } catch (e) { console.warn('Abgleich', e); zustand.fehler = e.message || String(e); }
      finally { zustand.laeuft = false; laeuft = null; melden(); }
    })();
    return laeuft;
  }

  async function hochladen() {
    for (const g of await Speicher.grabsteine()) {
      await client.storage.from(EIMER).remove([pfad(g.id)]).catch(() => {});
      const { error } = await client.from(TABELLE).upsert({ besitzer: zustand.uid, art: ART, id: g.id, daten: null, geloescht: true, geaendert: g.zeit }, { onConflict: 'besitzer,art,id' });
      if (error) throw new Error('Löschen: ' + error.message);
      await Speicher.grabsteinWeg(g.id);
    }
    for (const e of await Speicher.alle()) {
      if ((e.hochgeladen || 0) >= e.geaendert) continue;
      const { error } = await client.storage.from(EIMER).upload(pfad(e.id), e.datei, { upsert: true, contentType: e.datei.type || 'application/octet-stream' });
      if (error) throw new Error('Datei hochladen: ' + error.message);
      const daten = { name: e.name, art: e.art, groesse: e.datei.size };
      const { error: f } = await client.from(TABELLE).upsert({ besitzer: zustand.uid, art: ART, id: e.id, daten, geloescht: false, geaendert: e.geaendert }, { onConflict: 'besitzer,art,id' });
      if (f) throw new Error('Hochladen: ' + f.message);
      // nur als hochgeladen markieren, wenn seitdem nicht erneut geaendert
      const jetzt = await Speicher.holen(e.id);
      if (jetzt && jetzt.geaendert === e.geaendert) await Speicher.uebernehmen({ ...jetzt, hochgeladen: e.geaendert });
    }
  }

  async function herunterladen() {
    const stand = (await Speicher.meta('stand')) || '1970-01-01T00:00:00Z';
    let neuerStand = stand, etwasNeu = false;
    // Stand, bis zu dem sicher alles übernommen ist. Eine Änderung am gerade OFFENEN Dokument wird nie
    // eingespielt (sonst gingen laufende Eingaben und Rückgängig verloren) – sie wartet, bis es zu ist,
    // und wird darum bei jedem Abgleich erneut gelesen.
    let halten = null;
    const offen = window.KXD_OFFEN?.() || null;
    for (;;) {
      const { data, error } = await client.from(TABELLE).select('*').eq('art', ART).gt('server_zeit', neuerStand).order('server_zeit').limit(200);
      if (error) throw new Error('Herunterladen: ' + error.message);
      for (const z of data) {
        const vorher = neuerStand;
        neuerStand = z.server_zeit;
        const lokal = await Speicher.holen(z.id);
        if (lokal && lokal.geaendert >= z.geaendert) continue;           // eigenes Echo oder eigene, neuere Änderung
        if (z.id === offen) { halten ??= vorher; continue; }             // offenes Dokument: später
        if (z.geloescht) {
          if (lokal) { await Speicher.entfernen(z.id); etwasNeu = true; }
          continue;
        }
        const { data: blob, error: e2 } = await client.storage.from(EIMER).download(pfad(z.id));
        if (e2) throw new Error('Datei laden: ' + e2.message);
        // Griff auf eine lokale Originaldatei (Öffnen mit) bleibt erhalten
        await Speicher.uebernehmen({ ...(lokal || {}), id: z.id, name: z.daten.name, art: z.daten.art, datei: blob, geaendert: z.geaendert, hochgeladen: z.geaendert });
        etwasNeu = true;
        dispatchEvent(new CustomEvent('kxd-fern', { detail: { id: z.id } }));
      }
      if (data.length < 200) break;
    }
    const sicher = halten ?? neuerStand;
    if (sicher !== stand) await Speicher.metaSetzen('stand', sicher);
    if (etwasNeu) dispatchEvent(new Event('kxd-liste'));
  }

  return { zustand, starten, anmelden, abmelden, abgleichen };
})();
