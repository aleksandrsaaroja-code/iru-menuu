// Iru köök: ühine salvestus. Google Apps Script (script.google.com -> Uus projekt -> kleebi see kood).
// Avalda: Deploy -> New deployment -> Web app -> Execute as: Me, Who has access: Anyone.
const PIN = '3405';   // sama PIN, mis äpi Seadetes; ainult PIN-iga saab kirjutada

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const all = PropertiesService.getScriptProperties().getProperties();
  const ovr = {};
  Object.keys(all).forEach(function (k) { if (k.indexOf('o|') === 0) ovr[k.substring(2)] = all[k]; });
  const meta = JSON.parse(all['meta'] || '{}');
  return out_({ ok: true, ver: meta.ver || 0, depsBase: meta.depsBase || null,
                deps: all['deps'] ? JSON.parse(all['deps']) : null, ovr: ovr });
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, err: 'json' }); }
  if (String(b.pin) !== PIN) return out_({ ok: false, err: 'pin' });
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = PropertiesService.getScriptProperties();
    const meta = JSON.parse(p.getProperty('meta') || '{}');
    if (b.action === 'ovr') {
      if (b.value === null || b.value === undefined) p.deleteProperty('o|' + b.key);
      else p.setProperty('o|' + b.key, String(b.value));
    } else if (b.action === 'deps') {
      p.setProperty('deps', JSON.stringify(b.deps));
      meta.depsBase = b.base;
    } else if (b.action === 'clearDeps') {
      p.deleteProperty('deps'); meta.depsBase = null;
    }
    // kustuta üle 14 päeva vanad muudatused (võtmed algavad kuupäevaga YYYY-MM-DD)
    const lim = Utilities.formatDate(new Date(Date.now() - 14 * 864e5), 'UTC', 'yyyy-MM-dd');
    const all = p.getProperties();
    Object.keys(all).forEach(function (k) { if (k.indexOf('o|') === 0 && k.substring(2, 12) < lim) p.deleteProperty(k); });
    meta.ver = (meta.ver || 0) + 1;
    p.setProperty('meta', JSON.stringify(meta));
    return out_({ ok: true, ver: meta.ver });
  } finally {
    lock.releaseLock();
  }
}
