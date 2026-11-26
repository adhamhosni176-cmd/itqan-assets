/* ---------- ترويسات الجهات: دمج ترويسة جهة (ملف Word على الدرايف) في أي ملف Word بيطلع من النظام ---------- */
const LHF = {
  'faculty': '1sFkQM-Pq4dB3W4GDxzERVgMAvWHhfNJl',
  'vd-education': '1Px_Oe-F5wXoJNWWNfUn-GRXP6m1RI5ba',
  'vd-research': '1ZmnUokfrgsrvP7rUQP-YgMQURwiY-wJG',
  'vd-community': '1Yslrv9s4XZIPZN7looRZlUKKzKP6jLvx',
  'quality': '1aHfL9S4EosoXgOf_pahw7sNoo-DfpDWQ',
  'dept-journalism': '13ZkVsEMcH0WYk5rX6zds9UraFVN8sTxY',
  'dept-radiotv': '1Vkew1glSOuFKvpC_GXbK0Tiujkz-7OLf',
  'dept-pr': '1d803CgzjTd-InDxylF2Nb3_twd1QvF_m',
  'english': ''
};
function lhAttrs_(tag){ const o = {}; String(tag).replace(/([\w:]+)="([^"]*)"/g, (m, k, v) => { o[k] = v; return m; }); return o; }
function lhRels_(xml){ return (String(xml).match(/<Relationship\b[^>]*\/>/g) || []).map(lhAttrs_); }
/* files: أجزاء ملف Word الأساسي (Blobs بأسماء المسارات)، lh: أجزاء ملف الترويسة. mode = 'none' يشيل الترويسة */
function applyLh_(files, lh, mode){
  const byName = (arr, n) => arr.filter(b => b.getName() === n)[0];
  const text = b => b ? b.getDataAsString('UTF-8') : '';
  const put = (n, s, type) => { const i = files.findIndex(b => b.getName() === n); const nb = Utilities.newBlob(s, type || 'application/xml', n); if (i >= 0) files[i] = nb; else files.push(nb); };
  let doc = text(byName(files, 'word/document.xml'));
  const stripRefs = s => s.replace(/<w:(header|footer)Reference\b[^>]*\/>/g, '').replace(/<w:titlePg\/>/g, '');
  if (mode === 'none' || !lh){ put('word/document.xml', doc.replace(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g, stripRefs)); return files; }
  const ldoc = text(byName(lh, 'word/document.xml')), lrels = lhRels_(text(byName(lh, 'word/_rels/document.xml.rels')));
  const sect = (ldoc.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g) || []).pop() || '';
  const refOf = kind => { const L = sect.match(new RegExp('<w:' + kind + 'Reference\\b[^>]*\\/>', 'g')) || []; const a = L.map(lhAttrs_); const d = a.filter(x => x['w:type'] === 'default')[0] || a[0]; return d && d['r:id']; };
  const ct0 = text(byName(files, '[Content_Types].xml'));
  let ct = ct0, rels = text(byName(files, 'word/_rels/document.xml.rels')), refs = '';
  ['header', 'footer'].forEach(kind => {
    const rid = refOf(kind); if (!rid) return;
    const r = lrels.filter(x => x.Id === rid)[0]; if (!r) return;
    const part = byName(lh, 'word/' + r.Target); if (!part) return;
    const nn = 'lh' + kind[0].toUpperCase() + kind.slice(1) + '.xml';
    let prels = text(byName(lh, 'word/_rels/' + r.Target + '.rels'));
    if (prels){
      prels = prels.replace(/Target="media\/([^"]+)"/g, (m, f) => {
        const src = byName(lh, 'word/media/' + f);
        if (src) files.push(src.copyBlob().setName('word/media/lh_' + f));
        return 'Target="media/lh_' + f + '"';
      });
      put('word/_rels/' + nn + '.rels', prels);
    }
    put('word/' + nn, text(part));
    const type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.' + kind + '+xml';
    if (ct.indexOf('/word/' + nn + '"') < 0) ct = ct.replace('</Types>', '<Override PartName="/word/' + nn + '" ContentType="' + type + '"/></Types>');
    const id = 'rIdLH' + kind[0].toUpperCase();
    rels = rels.replace(new RegExp('<Relationship\\b[^>]*Id="' + id + '"[^>]*\\/>', 'g'), '')
      .replace('</Relationships>', '<Relationship Id="' + id + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/' + kind + '" Target="' + nn + '"/></Relationships>');
    refs += '<w:' + kind + 'Reference w:type="default" r:id="' + id + '"/>';
  });
  [['png', 'image/png'], ['jpeg', 'image/jpeg'], ['jpg', 'image/jpeg'], ['gif', 'image/gif'], ['emf', 'image/x-emf'], ['wmf', 'image/x-wmf']].forEach(([e, t]) => {
    if (!new RegExp('Extension="' + e + '"', 'i').test(ct)) ct = ct.replace('<Override', '<Default Extension="' + e + '" ContentType="' + t + '"/><Override');
  });
  const pgSz = (sect.match(/<w:pgSz\b[^>]*\/>/) || [''])[0], pgMar = (sect.match(/<w:pgMar\b[^>]*\/>/) || [''])[0];
  doc = doc.replace(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g, s => {
    s = stripRefs(s).replace(/^(<w:sectPr\b[^>]*>)/, '$1' + refs);
    if (pgSz) s = s.replace(/<w:pgSz\b[^>]*\/>/, pgSz);
    if (pgMar) s = s.replace(/<w:pgMar\b[^>]*\/>/, pgMar);
    return s;
  });
  put('word/document.xml', doc); put('word/_rels/document.xml.rels', rels); put('[Content_Types].xml', ct);
  return files;
}
function lhBlobs_(key){
  const id = LHF[key] || LHF.faculty;
  return Utilities.unzip(DriveApp.getFileById(id).getBlob().setContentType('application/zip'));
}
