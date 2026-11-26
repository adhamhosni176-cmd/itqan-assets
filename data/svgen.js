/* ================= توليد مستندات الاستبيانات عند الطلب (إتقان) =================
   زرار «استخراج النتائج» في إتقان يكتب طلبًا في عمود «طلب التوليد» بالسجل،
   ودالة aeQueue (مؤقت كل ٥ دقائق) تنفّذ الطلبات فقط ولا تفتح أي استمارة بدون طلب. */
var SG = {
  REG: '1I-yFI8-dH7d2RnJDTplzJMQZVArqirpxqRODZA8F0I8', ROOT: '1mabqQbmUr7nkd9DwH68B4Da3RclUEChz',
  RES: 'نتائج العبارات', LIB: 'ربط العبارات بالمحاور', RULES: 'قواعد خطط التحسين', NAMES: 'أسماء المخاطبين',
  DOCS: 'مستندات الاستبيانات', TH: 'مكتبة الإجراءات',
  COL: {FORM: 22, REQ: 23, MSG: 24},
  LH: {quality: '1aHfL9S4EosoXgOf_pahw7sNoo-DfpDWQ', faculty: '1sFkQM-Pq4dB3W4GDxzERVgMAvWHhfNJl', 'vd-education': '1Px_Oe-F5wXoJNWWNfUn-GRXP6m1RI5ba',
       'vd-research': '1ZmnUokfrgsrvP7rUQP-YgMQURwiY-wJG', 'vd-community': '1Yslrv9s4XZIPZN7looRZlUKKzKP6jLvx', 'dept-journalism': '13ZkVsEMcH0WYk5rX6zds9UraFVN8sTxY',
       'dept-radiotv': '1Vkew1glSOuFKvpC_GXbK0Tiujkz-7OLf', 'dept-pr': '1d803CgzjTd-InDxylF2Nb3_twd1QvF_m', english: '1iQ5UnApbzRak3aYo41jNdPpOeOhHBtqT'}
};
// الجهة: [الصفة كما في أسماء المخاطبين، الترويسة، صوت الجهة]
var SG_B = {
  ve: ['وكيل الكلية لشئون التعليم والطلاب', 'vd-education', 'الوكالة'], vr: ['وكيل الكلية لشئون الدراسات العليا والبحوث', 'vd-research', 'الوكالة'],
  vc: ['وكيل الكلية لشئون خدمة المجتمع وتنمية البيئة', 'vd-community', 'الوكالة'], dn: ['عميد الكلية', 'faculty', 'عمادة الكلية'],
  dm: ['مدير كلية الإعلام', 'faculty', 'إدارة الكلية'], jr: ['رئيس قسم الصحافة', 'dept-journalism', 'القسم'], rt: ['رئيس قسم الإذاعة والتليفزيون', 'dept-radiotv', 'القسم'],
  pr: ['رئيس قسم العلاقات العامة والإعلان', 'dept-pr', 'القسم'], en: ['منسق الشعبة الإنجليزية', 'english', 'الشعبة'], qu: ['مدير وحدة ضمان الجودة والاعتماد', 'quality', 'الوحدة']
};
var SG_FB = {2:'dm',3:'ve',4:'ve',5:'ve',6:'ve',7:'ve',8:'ve',9:'ve',10:'dm',11:'dn',13:'dm',14:'dm',16:'ve',17:'ve',18:'ve',19:'dn',20:'dn',21:'dm',22:'qu',23:'qu',24:'qu',25:'qu',26:'vr',
             27:'dn',29:'dm',30:'dn',32:'dm',33:'dm',34:'dm',35:'dn',36:'dn',37:'qu',38:'dm',40:'dm',41:'vr',42:'vr',43:'ve',44:'dm',45:'dm',46:'vc',47:'vc',48:'vr',49:'vc',50:'vc',51:'vc'};
var SG_DEPT = [['الإنجليز', 'en'], ['العلاقات', 'pr'], ['الإذاعة', 'rt'], ['الصحافة', 'jr']];
var SG_PERDEPT = [12, 15, 28, 39];
var SG_R = {VE:'وكيل الكلية لشئون التعليم والطلاب',VR:'وكيل الكلية للدراسات العليا والبحوث',VC:'وكيل الكلية لشئون خدمة المجتمع وتنمية البيئة',DN:'عميد الكلية',DM:'مدير الكلية',HD:'رؤساء الأقسام العلمية',FM:'أعضاء هيئة التدريس',QU:'وحدة ضمان الجودة',TR:'لجنة التدريب ونشر ثقافة الجودة',SC:'لجنة الاستبيانات والتغذية الراجعة',LB:'أمين المكتبة',GU:'وحدة متابعة الخريجين',SA:'إدارة رعاية الطلاب',SS:'إدارة شئون الطلاب',EN:'الشئون الهندسية',IT:'مسئول المعامل وتكنولوجيا المعلومات',CR:'وحدة إدارة الأزمات والكوارث',HR:'شئون العاملين',AC:'لجنة الإرشاد الأكاديمي',PC:'منسق البرنامج',CT:'الكنترول',FIN:'الشئون المالية',WEB:'مسئول الموقع الإلكتروني'};
var SG_LVN = {1: 'البدء', 2: 'التثبيت', 3: 'المعالجة الجذرية'};
var SG_C = {NAVY: '14514A', GOLD: 'C9A961', TEAL: '1F6F63', LIGHT: 'EDF0F1', INK: '273234', BORDER: 'C9D3D5'};

/* ---------- أدوات ---------- */
function sg_ss_(){ return SpreadsheetApp.openById(SG.REG); }
function sg_sh_(n){ return sg_ss_().getSheetByName(n); }
function sg_ar_(s){ return String(s).replace(/[0-9]/g, function(d){ return '٠١٢٣٤٥٦٧٨٩'[+d]; }).replace(/\./g, '٫'); }
function sg_w_(s){ return String(s == null ? '' : s).replace(/[٠-٩]/g, function(d){ return '٠١٢٣٤٥٦٧٨٩'.indexOf(d); }); }
function sg_num_(s){ var x = parseFloat(sg_w_(s).replace(/[^\d.]/g, '')); return isNaN(x) ? 0 : x; }
function sg_pct_(x){ return sg_ar_((+x).toFixed(2)) + '٪'; }
function sg_esc_(s){ return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
function sg_clean_(s){ return String(s || '').replace(/\s+/g, ' ').trim(); }
function sg_safe_(s){ return String(s).replace(/[\\\/:*?"<>|]/g, '-').substring(0, 150).trim(); }
function sg_yfold_(y){ return 'العام الجامعي ' + String(y).replace(/[‏‎]/g, '').replace(/\s/g, '').replace('/', '-'); }
function sg_stdName_(s){ return sg_safe_(sg_clean_(s).replace(/^استبيانات\s*/, '')); }
function sg_svName_(s){ return sg_safe_(sg_clean_(s).replace(/^(استبيان|استبانة)\s*(ل(?=قياس|تقييم))?/, '')); }
function sg_sub_(p, n){ var it = p.getFoldersByName(n); return it.hasNext() ? it.next() : p.createFolder(n); }
function sg_folder_(row){ return sg_sub_(sg_sub_(sg_sub_(DriveApp.getFolderById(SG.ROOT), sg_yfold_(row[1])), sg_stdName_(row[2])), sg_svName_(row[3])); }
function sg_nz_(s){ return String(s).replace(/[إأآا]/g, 'ا').replace(/ة/g, 'ه').replace(/ى/g, 'ي').replace(/[^ء-ي ]/g, ' ').replace(/\s+/g, ' ').trim(); }
function sg_sim_(a, b){
  a = sg_nz_(a); b = sg_nz_(b); if (!a || !b) return 0;
  var g = function(s){ var m = {}; for (var i = 0; i < s.length - 1; i++){ var k = s.substr(i, 2); m[k] = (m[k] || 0) + 1; } return m; };
  var A = g(a), B = g(b), t = 0, na = 0, nb = 0, k;
  for (k in A) { na += A[k]; if (B[k]) t += Math.min(A[k], B[k]); } for (k in B) nb += B[k];
  return 2 * t / (na + nb);
}

/* ---------- مصادر البيانات ---------- */
function sg_lib_(){
  var v = sg_sh_(SG.LIB).getDataRange().getValues(), L = {};
  v.slice(1).forEach(function(r){
    var f = +sg_w_(r[0]); if (!f) return;
    (L[f] = L[f] || []).push({j: +sg_w_(r[3]), t: sg_clean_(r[4]), th: r[5], a: [[r[7], r[8], r[9]], [r[10], r[11], r[12]], [r[13], r[14], r[15]]]});
  });
  return L;
}
function sg_themes_(){ var m = {}; sg_sh_(SG.TH).getDataRange().getValues().slice(1).forEach(function(r){ if (r[0]) m[r[0]] = r[1]; }); return m; }
function sg_thr_(){ try { var v = sg_sh_(SG.RULES).getDataRange().getValues(); for (var i = 1; i < v.length; i++) if (/حد الضعف/.test(v[i][0])) return +v[i][1] || 75; } catch (e) {} return 75; }
function sg_officials_(){
  var v = sg_sh_(SG.NAMES).getDataRange().getValues().slice(1), ys = {}, m = {};
  v.forEach(function(r){ ys[sg_w_(r[0])] = 1; });
  var last = Object.keys(ys).sort().pop();
  v.forEach(function(r){ if (sg_w_(r[0]) == last) m[sg_clean_(r[1])] = sg_clean_(r[2]); });
  return m;
}
function sg_name_(off, role){
  if (off[role]) return off[role];
  for (var k in off) if (sg_sim_(k, role) > .85) return off[k];
  return '';
}
function sg_bodyOf_(form, title){
  if (SG_PERDEPT.indexOf(form) >= 0) for (var i = 0; i < SG_DEPT.length; i++) if (String(title).indexOf(SG_DEPT[i][0]) >= 0) return SG_DEPT[i][1];
  return SG_FB[form] || 'dn';
}

/* نتائج العبارات: م | العام | الجهة | المصدر | رقم النموذج | رقم العبارة | العبارة | النسبة | الاستجابات | حجم المجتمع | الملف */
function sg_resSheet_(){
  var ss = sg_ss_(), sh = ss.getSheetByName(SG.RES);
  if (!sh) { sh = ss.insertSheet(SG.RES); sh.setRightToLeft(true); sh.getRange(1, 1, 1, 11).setValues([['م في السجل', 'العام', 'الجهة', 'المصدر', 'رقم النموذج', 'رقم العبارة', 'العبارة', 'النسبة', 'الاستجابات', 'حجم المجتمع', 'الملف']]).setFontWeight('bold').setBackground('#1F6F63').setFontColor('#FFFFFF'); }
  return sh;
}
function sg_results_(){ return sg_resSheet_().getDataRange().getValues().slice(1).filter(function(r){ return r[0] !== ''; }); }
function sg_putResults_(n, rows){
  var sh = sg_resSheet_(), v = sh.getDataRange().getValues(), keep = [v[0]];
  v.slice(1).forEach(function(r){ if (String(r[0]) != String(n) && r[0] !== '') keep.push(r); });
  rows.forEach(function(r){ keep.push(r); });
  sh.clearContents(); sh.getRange(1, 1, keep.length, 11).setValues(keep);
}

/* ---------- عناصر Word ---------- */
function sg_run_(t, o){
  o = o || {}; var sz = o.sz || 28, c = o.color || SG_C.INK;
  var rpr = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>' + (o.b ? '<w:b/><w:bCs/>' : '') + '<w:color w:val="' + c + '"/><w:sz w:val="' + sz + '"/><w:szCs w:val="' + sz + '"/><w:rtl/>';
  return '<w:r><w:rPr>' + rpr + '</w:rPr>' + String(t).split('\n').map(function(p, k){ return (k ? '<w:br/>' : '') + '<w:t xml:space="preserve">' + sg_esc_(p) + '</w:t>'; }).join('') + '</w:r>';
}
function sg_p_(content, o){
  o = o || {}; var jc = o.jc === undefined ? 'both' : o.jc, line = o.line || 264;
  if (content.indexOf('<w:') !== 0) content = sg_run_(content);
  var ppr = '<w:bidi/>' + (o.keep ? '<w:keepNext/>' : '') + '<w:spacing w:before="' + (o.before || 0) + '" w:after="' + (o.after == null ? 80 : o.after) + '" w:line="' + line + '" w:lineRule="' + (o.exact ? 'exact' : 'auto') + '"/>' + (jc ? '<w:jc w:val="' + jc + '"/>' : '');
  return '<w:p><w:pPr>' + ppr + '</w:pPr>' + content + '</w:p>';
}
function sg_title_(t, sub){ return sg_p_(sg_run_(t, {b: 1, sz: 32, color: SG_C.NAVY}), {jc: 'center', after: 0, keep: 1}) + (sub ? sg_p_(sg_run_(sub, {b: 1, sz: 24, color: SG_C.TEAL}), {jc: 'center', after: 100, keep: 1}) : ''); }
function sg_h_(t){ return sg_p_(sg_run_(t, {b: 1, sz: 26, color: SG_C.NAVY}), {before: 100, after: 60, keep: 1}); }
function sg_cell_(d, w, hdr){
  var sz = d.sz || 23, b = d.b != null ? d.b : hdr, fill = d.fill || (hdr ? SG_C.TEAL : ''), color = d.color || (hdr ? 'FFFFFF' : null), jc = d.jc || (hdr ? 'center' : null);
  var tc = '<w:tcW w:w="' + w + '" w:type="dxa"/>' + (d.span > 1 ? '<w:gridSpan w:val="' + d.span + '"/>' : '') + (fill ? '<w:shd w:val="clear" w:color="auto" w:fill="' + fill + '"/>' : '') + '<w:vAlign w:val="center"/>';
  var body = String(d.t).split('\n').map(function(l){ return sg_p_(sg_run_(l, {b: b, sz: sz, color: color}), {jc: jc, after: 0, line: Math.round(sz * 12.5), exact: 1}); }).join('');
  return '<w:tc><w:tcPr>' + tc + '</w:tcPr>' + body + '</w:tc>';
}
function sg_table_(rows, widths, header){
  if (header == null) header = true;
  var bd = '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(function(s){ return '<w:' + s + ' w:val="single" w:sz="4" w:space="0" w:color="' + SG_C.BORDER + '"/>'; }).join('') + '</w:tblBorders>';
  var sum = widths.reduce(function(a, b){ return a + b; }, 0);
  var x = '<w:tbl><w:tblPr><w:bidiVisual/><w:tblW w:w="' + sum + '" w:type="dxa"/><w:jc w:val="center"/>' + bd + '<w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="40" w:type="dxa"/><w:left w:w="90" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/><w:right w:w="90" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>' + widths.map(function(w){ return '<w:gridCol w:w="' + w + '"/>'; }).join('') + '</w:tblGrid>';
  rows.forEach(function(r, ri){
    var hdr = header && ri == 0, ci = 0;
    x += '<w:tr><w:trPr><w:cantSplit/>' + (hdr ? '<w:tblHeader/>' : '') + '</w:trPr>';
    r.forEach(function(c){ var d = (c && typeof c == 'object') ? c : {t: c == null ? '' : c}; var sp = d.span || 1; var w = widths.slice(ci, ci + sp).reduce(function(a, b){ return a + b; }, 0); x += sg_cell_(d, w, hdr); ci += sp; });
    x += '</w:tr>';
  });
  return x + '</w:tbl>' + sg_p_(sg_run_('', {sz: 12}), {after: 0, line: 200});
}
function sg_signs_(items){
  if (items.length == 1) items = [['', '']].concat(items);
  var W = 9360, lens = items.map(function(x){ return Math.max(x[0].length, (x[1] || '').length) + 6; });
  if (items.length == 2 && !items[0][0]) lens = [55, 45];
  var s = lens.reduce(function(a, b){ return a + b; }, 0), ws = lens.map(function(l){ return Math.floor(W * l / s); }); ws[ws.length - 1] = W - ws.slice(0, -1).reduce(function(a, b){ return a + b; }, 0);
  var c = function(t, w, b, color, before, keep){ return '<w:tc><w:tcPr><w:tcW w:w="' + w + '" w:type="dxa"/></w:tcPr>' + sg_p_(sg_run_(t, {b: b, sz: 23, color: color}), {jc: 'center', before: before, after: 0, keep: keep}) + '</w:tc>'; };
  var t = '<w:tbl><w:tblPr><w:bidiVisual/><w:tblW w:w="' + W + '" w:type="dxa"/><w:jc w:val="center"/><w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(function(x){ return '<w:' + x + ' w:val="nil"/>'; }).join('') + '</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>' + ws.map(function(w){ return '<w:gridCol w:w="' + w + '"/>'; }).join('') + '</w:tblGrid>';
  t += '<w:tr><w:trPr><w:cantSplit/></w:trPr>' + items.map(function(x, i){ return c(x[0], ws[i], 1, SG_C.NAVY, 160, 1); }).join('') + '</w:tr>';
  t += '<w:tr><w:trPr><w:cantSplit/></w:trPr>' + items.map(function(x, i){ return c(x[1] || '', ws[i], 1, null, 380, 0); }).join('') + '</w:tr></w:tbl>';
  return t;
}
function sg_box_(t){
  var bd = ['top', 'left', 'bottom', 'right'].map(function(x){ return '<w:' + x + ' w:val="single" w:sz="8" w:space="0" w:color="' + SG_C.GOLD + '"/>'; }).join('');
  return '<w:tbl><w:tblPr><w:bidiVisual/><w:tblW w:w="4600" w:type="dxa"/><w:jc w:val="center"/><w:tblBorders>' + bd + '</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="4600"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4600" w:type="dxa"/></w:tcPr>' + sg_p_(sg_run_(t || 'لجنة الاستبيانات والتغذية الراجعة', {b: 1, sz: 26, color: SG_C.NAVY}), {jc: 'center', after: 0, keep: 1}) + '</w:tc></w:tr></w:tbl>' + sg_p_(sg_run_('', {sz: 10}), {after: 60, line: 200, keep: 1});
}
var SG_LHC = {};
function sg_build_(lh, body, name){
  if (!SG_LHC[lh]) SG_LHC[lh] = Utilities.unzip(DriveApp.getFileById(SG.LH[lh]).getBlob().setContentType('application/zip'));
  var files = SG_LHC[lh].map(function(x){
    if (x.getName() != 'word/document.xml') return x;
    var d = x.getDataAsString('UTF-8'), i = d.indexOf('<w:body>') + 8, j = d.lastIndexOf('<w:sectPr');
    return Utilities.newBlob(d.slice(0, i) + body + d.slice(j), 'application/xml', 'word/document.xml');
  });
  return Utilities.zip(files, name).setName(name).setContentType('application/vnd.openxmlformats-officedocument.wordprocessingml.document');
}
function sg_save_(folder, name, blob){
  var it = folder.getFilesByName(name); while (it.hasNext()) it.next().setTrashed(true);
  return folder.createFile(blob);
}

/* ---------- خطة التحسين ---------- */
function sg_plans_(n, row, res, lib, th, thr, off){
  var form = +sg_w_(row[SG.COL.FORM - 1]), NEED = [22, 23].indexOf(form) >= 0, bad = function(v){ return NEED ? v >= thr : v < thr; }, my = res.filter(function(r){ return String(r[0]) == String(n); }), plans = {};
  my.forEach(function(r){ var b = r[2]; var p = plans[b] = plans[b] || {b: b, items: {}, resp: 0, pop: 0, files: {}}; if (r[5] !== '') (p.items[+r[5]] = p.items[+r[5]] || []).push(+r[7]); p.resp = Math.max(p.resp, +r[8] || 0); p.pop = Math.max(p.pop, +r[9] || 0); });
  var yW = function(y){ return sg_w_(y).replace(/[^\d]/g, ''); }, yr = yW(row[1]);
  var hist = {}; res.forEach(function(r){ if (+r[4] == form && r[5] !== '') { var k = r[2] + '|' + r[5], y = yW(r[1]); (hist[k] = hist[k] || {})[y] = (hist[k][y] || []).concat([+r[7]]); } });
  var years = {}; res.forEach(function(r){ years[yW(r[1])] = 1; }); var prevYears = Object.keys(years).filter(function(y){ return y < yr; }).sort().reverse();
  var avgOf = function(a){ return a.reduce(function(x, y){ return x + y; }, 0) / a.length; };
  var level = function(b, j){ var lv = 1; for (var i = 0; i < prevYears.length; i++){ var h = (hist[b + '|' + j] || {})[prevYears[i]]; if (h && bad(avgOf(h))) lv++; else break; } return Math.min(lv, 3); };
  var prev = function(b, j){ var h = prevYears.length ? (hist[b + '|' + j] || {})[prevYears[0]] : null; return h ? avgOf(h) : null; };
  var L = {}; (lib[form] || []).forEach(function(x){ L[x.j] = x; });
  return Object.keys(plans).map(function(b){
    var p = plans[b], its = Object.keys(p.items).map(function(j){ return [+j, avgOf(p.items[j])]; }).filter(function(x){ return L[x[0]]; }).sort(function(a, c){ return NEED ? c[1] - a[1] : a[1] - c[1]; });
    if (!its.length) return null;
    p.its = its; p.avg = avgOf(its.map(function(x){ return x[1]; }));
    var weak = its.filter(function(x){ return bad(x[1]); }); p.forced = !weak.length; if (!weak.length) weak = its.slice(0, 3); p.need = NEED;
    if (NEED) { p.rows = weak.slice(0, 10).map(function(x){ var lv = p.forced ? 1 : level(b, x[0]), a = L[x[0]].a[lv - 1]; return {label: L[x[0]].t.replace(/\.$/, ''), stm: sg_pct_(x[1]), act: a[0], who: String(a[1]).split(/\s*–\s*/).join('، و'), ev: a[2], lv: lv, low: x[1]}; }); p.extra = []; p.improved = []; p.L = L; p.level = level; return p; }
    var groups = {}, order = [];
    weak.forEach(function(x){ var t = L[x[0]].th || 'T00'; if (!groups[t]) { groups[t] = []; order.push(t); } groups[t].push([x[0], x[1], p.forced ? 1 : level(b, x[0])]); });
    var rows = order.map(function(t){
      var g = groups[t].sort(function(a, c){ return (c[2] - a[2]) || (a[1] - c[1]); }), acts = [];
      for (var i = 0; i < g.length && acts.length < 3; i++){ var a = L[g[i][0]].a[g[i][2] - 1]; if (a && a[0] && acts.map(function(z){ return z[0]; }).indexOf(a[0]) < 0) acts.push(a); }
      var stm = g.slice(0, 3).map(function(x){ return '• ' + L[x[0]].t.replace(/\.$/, '') + ' – ' + sg_pct_(x[1]); }).join('\n') + (g.length > 3 ? '\n• و' + sg_ar_(g.length - 3) + ' عبارات أخرى في المحور نفسه' : '');
      var who = []; acts.forEach(function(a){ String(a[1]).split(/\s*–\s*/).forEach(function(w){ w = w == 'الجهة المرفوع لها التقرير' ? SG_B[b][0] : w; if (w && who.indexOf(w) < 0) who.push(w); }); });
      var ev = []; acts.forEach(function(a){ String(a[2]).split('،').forEach(function(e){ e = e.trim(); if (e && ev.indexOf(e) < 0) ev.push(e); }); });
      return {label: th[t] || 'مجال التحسين', stm: stm, act: acts.map(function(a){ return (acts.length > 1 ? '• ' : '') + a[0]; }).join('\n'), who: who.join('، و'), ev: ev.join('، '), lv: g[0][2], low: g[0][1]};
    }).sort(function(a, c){ return (c.lv - a.lv) || (a.low - c.low); });
    p.rows = rows.slice(0, 6); p.extra = rows.slice(6);
    p.improved = its.filter(function(x){ var pv = prev(b, x[0]); return !bad(x[1]) && pv != null && bad(pv); }).map(function(x){ return [x[0], prev(b, x[0]), x[1]]; });
    p.L = L; p.level = level; return p;
  }).filter(function(x){ return x; });
}

function sg_card_(row, plans, thr, form, nItems, off){
  var L = function(t){ return {t: t, b: 1, fill: SG_C.LIGHT, color: SG_C.NAVY, sz: 21}; };
  var resp = plans.reduce(function(a, p){ return a + (p.resp || 0); }, 0), pop = sg_num_(row[13]) || plans.reduce(function(a, p){ return a + (p.pop || 0); }, 0);
  var all = []; plans.forEach(function(p){ p.its.forEach(function(x){ all.push(x[1]); }); });
  var avg = all.reduce(function(a, b){ return a + b; }, 0) / all.length;
  var b = sg_box_() + sg_title_('بطاقة بيانات تطبيق الاستبانة', 'العام الجامعي ' + row[1]);
  b += sg_table_([
    [L('اسم الاستبانة'), {t: sg_svName_(row[3]), span: 3}],
    [L('المعيار'), {t: sg_stdName_(row[2]), span: 3}],
    [L('العام الجامعي'), {t: row[1], sz: 21}, L('أسلوب التطبيق'), {t: 'إلكتروني عبر نماذج جوجل', sz: 21}],
    [L('الفئة المستهدفة'), {t: row[12] || '—', span: 3}],
    [L('حجم المجتمع'), pop ? sg_ar_(pop) : '—', L('عدد الاستجابات'), resp ? sg_ar_(resp) + ' استجابة' : '—'],
    [L('نسبة الاستجابة'), resp && pop ? sg_pct_(resp / pop * 100) : '—', L('محتوى الأداة'), sg_ar_(nItems) + ' عبارة'],
    [L('متوسط الرضا العام'), sg_pct_(avg), L('حد التحسين'), 'أقل من ' + sg_ar_(thr) + '٪'],
    [L('فترة التطبيق'), {t: 'من ‏/‏    ‏/‏          إلى ‏/‏    ‏/‏', span: 3}],
    [L('يُرفع التقرير إلى'), {t: row[14] || '—', span: 3}],
    [L('خطط التحسين الواردة من'), {t: plans.map(function(p){ return SG_B[p.b][0]; }).join('، '), span: 3}]
  ], [2150, 2530, 2150, 2530], false);
  b += sg_p_(sg_run_('المرفقات: ', {b: 1, sz: 26, color: SG_C.NAVY}) + sg_run_('تحليل نتائج الاستبانة، وتقرير النتائج المرفوع إلى الجهة المختصة، ورد الجهة وخطة التحسين.', {sz: 24}), {after: 0, keep: 1});
  b += sg_signs_([['مسؤول لجنة الاستبيانات', sg_name_(off, 'مسؤول لجنة الاستبيانات') || 'د. كريم عادل'], ['مدير وحدة ضمان الجودة والاعتماد', sg_name_(off, 'مدير وحدة ضمان الجودة والاعتماد')], ['عميد الكلية', sg_name_(off, 'عميد الكلية')]]);
  return b;
}

function sg_term_(){ var y = new Date().getFullYear(), m = new Date().getMonth(); var a = m >= 8 ? y : y - 1; return 'الفصل الدراسي الثاني ' + sg_ar_(a) + '‏/‏' + sg_ar_(a + 1); }

function sg_plan_(row, p, thr, off){
  var B = SG_B[p.b], voice = B[2], nm = sg_svName_(row[3]), its = p.its, L = p.L, TERM = sg_term_();
  var st = function(j){ return L[j].t.replace(/\.$/, ''); };
  var b = sg_p_(sg_run_('السيد الأستاذ الدكتور/ مدير وحدة ضمان الجودة والاعتماد بالكلية', {b: 1, sz: 24}), {after: 40}) + sg_p_(sg_run_('تحية طيبة وبعد،', {sz: 24}), {after: 120});
  var rr = p.resp && p.pop ? '، والتي شارك فيها ' + sg_ar_(p.resp) + ' من أصل ' + sg_ar_(p.pop) + ' بنسبة استجابة ' + sg_pct_(p.resp / p.pop * 100) : '';
  b += sg_p_(sg_run_('بالإشارة إلى تقرير لجنة الاستبيانات والتغذية الراجعة بشأن نتائج ' + nm + ' للعام الجامعي ' + row[1] + rr + '، نفيد سيادتكم بأنه تم الاطلاع على التقرير ونتائجه، ونوافيكم بقراءة ' + voice + ' للنتائج وخطة التحسين على النحو التالي:', {sz: 24}), {after: 120});
  b += sg_h_('أولًا: قراءة النتائج');
  var nw = its.filter(function(x){ return x[1] < thr; }).length;
  b += sg_p_(sg_run_('بلغ المتوسط العام لنسب الرضا عن عبارات الاستبانة الـ' + sg_ar_(its.length) + ' ' + sg_pct_(p.avg) + (nw ? '، وجاءت ' + sg_ar_(nw) + ' عبارة دون حد التحسين المعتمد وهو ' + sg_ar_(thr) + '٪.' : '، ولم تأتِ أي عبارة دون حد التحسين المعتمد وهو ' + sg_ar_(thr) + '٪، ومع ذلك تعتبر ' + voice + ' أدنى العبارات رضا فرصًا للتحسين.'), {sz: 24}));
  var top = its.slice().sort(function(a, c){ return c[1] - a[1]; }).slice(0, 3);
  b += sg_p_(sg_run_('وتمثلت أبرز نقاط القوة في: ', {b: 1, sz: 24}) + sg_run_(top.map(function(x){ return st(x[0]) + ' بنسبة ' + sg_pct_(x[1]); }).join('، ') + '.', {sz: 24}));
  var rep = its.filter(function(x){ return x[1] < thr && p.level(p.b, x[0]) > 1; });
  if (rep.length && !p.forced) b += sg_p_(sg_run_('ويُلاحظ أن ' + sg_ar_(rep.length) + ' من هذه العبارات جاءت دون الحد في القياس السابق أيضًا، ولذا انتقلت إجراءاتها في هذه الخطة إلى مستوى «' + SG_LVN[Math.max.apply(null, rep.map(function(x){ return p.level(p.b, x[0]); }))] + '» بدل تكرار إجراءات العام السابق.', {sz: 24}));
  if (p.improved.length) b += sg_p_(sg_run_('ما تحقق: ', {b: 1, sz: 24}) + sg_run_('ارتفعت نسبة الرضا عن ' + p.improved.slice(0, 2).map(function(x){ return st(x[0]) + ' من ' + sg_pct_(x[1]) + ' إلى ' + sg_pct_(x[2]); }).join('، و') + '، وتحرص ' + voice + ' على الحفاظ على هذا المستوى.', {sz: 24}));
  if (p.resp && p.pop && p.resp / p.pop < .4) b += sg_p_(sg_run_('كما جاءت نسبة الاستجابة دون ٤٠٪، بما يستدعي العمل على رفعها في التطبيق القادم بالتعاون مع لجنة الاستبيانات.', {sz: 24}));
  b += sg_h_('ثانيًا: خطة التحسين');
  b += sg_p_(sg_run_('تُنفَّذ الإجراءات خلال ' + TERM + '، ويُعد الإجراء محققًا لهدفه إذا بلغت نسبة الرضا عن عباراته ' + sg_ar_(thr) + '٪ فأكثر في التطبيق القادم للاستبانة، وذلك على النحو التالي:', {sz: 24}), {after: 60, keep: 1});
  var rows = [['م', 'المحور', 'العبارات ونسب الرضا', 'الإجراء', 'المسئول عن التنفيذ', 'دليل التنفيذ'].map(function(h){ return {t: h, sz: 21}; })];
  p.rows.forEach(function(r, k){ rows.push([{t: sg_ar_(k + 1), jc: 'center', sz: 21}, {t: r.label, sz: 21, b: 1, color: SG_C.NAVY}, {t: r.stm, sz: 20}, {t: r.act + (r.lv > 1 ? '\n(مستوى ' + SG_LVN[r.lv] + ')' : ''), sz: 21}, {t: r.who, sz: 21}, {t: r.ev, sz: 21}]); });
  b += sg_table_(rows, [340, 1250, 2050, 3020, 1500, 1200]);
  if (p.extra.length) b += sg_p_(sg_run_('وتُدرج باقي المجالات دون الحد (' + p.extra.map(function(r){ return r.label; }).join('، ') + ') في خطة المتابعة التالية وفق الأولوية.', {sz: 22}));
  b += sg_h_('ثالثًا: المتابعة');
  b += sg_p_(sg_run_('توافي ' + voice + ' وحدة ضمان الجودة والاعتماد بما يتم تنفيذه من إجراءات هذه الخطة وأدلته في نهاية ' + TERM + '.', {sz: 24}), {keep: 1});
  b += sg_p_(sg_run_('وتفضلوا بقبول فائق الاحترام،', {sz: 24}), {jc: 'center', before: 120, keep: 1});
  b += sg_signs_([[B[0], sg_name_(off, B[0])]]);
  return b;
}

function sg_needsPlan_(row, p, thr, off){
  var b = sg_box_('لجنة التدريب ونشر ثقافة الجودة') + sg_title_('خطة التدريب المبنية على نتائج استبانة الاحتياجات التدريبية', sg_svName_(row[3]) + ' – العام الجامعي ' + row[1]);
  var rr = p.resp && p.pop ? '، والتي شارك فيها ' + sg_ar_(p.resp) + ' من أصل ' + sg_ar_(p.pop) : '';
  b += sg_p_(sg_run_('بناءً على نتائج ' + sg_svName_(row[3]) + ' للعام الجامعي ' + row[1] + rr + '، أعدت لجنة التدريب ونشر ثقافة الجودة الخطة التالية للاحتياجات الأعلى (' + sg_ar_(thr) + '٪ فأكثر)، مرتبة حسب درجة الاحتياج، وتُنفذ خلال العام الجامعي التالي، وتنتقل في الاحتياجات التي تكرر ارتفاعها عن قياس العام السابق إلى برامج متقدمة:', {sz: 24}), {after: 100});
  var rows = [['م', 'الاحتياج التدريبي', 'درجة الاحتياج', 'البرنامج التدريبي', 'الجهة المنفذة', 'دليل التنفيذ'].map(function(h){ return {t: h, sz: 21}; })];
  p.rows.forEach(function(r, k){ rows.push([{t: sg_ar_(k + 1), jc: 'center', sz: 21}, {t: r.label, sz: 20, b: 1, color: SG_C.NAVY}, {t: r.stm, jc: 'center', sz: 21}, {t: r.act + (r.lv > 1 ? '\n(برنامج متقدم)' : ''), sz: 20}, {t: r.who, sz: 20}, {t: r.ev, sz: 20}]); });
  b += sg_table_(rows, [340, 2200, 900, 3000, 1600, 1320]);
  b += sg_h_('آلية التنفيذ والمتابعة');
  b += sg_p_(sg_run_('تُعلن مواعيد البرامج في بداية كل فصل دراسي، ويُطبق تقييم لكل برنامج عقب انتهائه، وتُحفظ كشوف الحضور وتقييمات المتدربين دليلًا على التنفيذ، وتُرفع تقارير التنفيذ إلى وحدة ضمان الجودة والاعتماد في نهاية كل فصل.', {sz: 24}), {keep: 1});
  b += sg_signs_([['مدير وحدة ضمان الجودة والاعتماد', sg_name_(off, 'مدير وحدة ضمان الجودة والاعتماد')], ['يُعتمد، عميد الكلية', sg_name_(off, 'عميد الكلية')]]);
  return b;
}

function sg_follow_(row, p, off){
  var L = function(t){ return {t: t, b: 1, fill: SG_C.LIGHT, color: SG_C.NAVY, sz: 21}; };
  var b = sg_box_() + sg_title_('تقرير متابعة تنفيذ خطة التحسين', sg_svName_(row[3]) + ' – العام الجامعي ' + row[1]);
  b += sg_table_([[L('الجهة المسئولة عن الخطة'), SG_B[p.b][0], L('تاريخ ورود الخطة'), '‏/‏    ‏/‏'], [L('فترة المتابعة'), {t: 'من ‏/‏    ‏/‏          إلى ‏/‏    ‏/‏', span: 3}]], [2100, 2580, 2100, 2580], false);
  b += sg_h_('موقف تنفيذ إجراءات الخطة ونتيجة القياس');
  var rows = [['م', 'الإجراء', 'المسئول', 'حالة التنفيذ', 'النسبة في ' + row[1], 'النسبة في القياس التالي'].map(function(h){ return {t: h, sz: 20}; })];
  p.rows.forEach(function(r, k){ rows.push([{t: sg_ar_(k + 1), jc: 'center', sz: 20}, {t: r.act, sz: 19}, {t: r.who, sz: 19}, {t: '☐ تم\n☐ تم جزئيًا\n☐ لم يتم', sz: 20}, {t: sg_pct_(r.low), jc: 'center', sz: 20}, '']); });
  b += sg_table_(rows, [400, 3860, 1700, 1200, 1100, 1100]);
  b += sg_p_(sg_run_('ملاحظات اللجنة وتوصياتها: ', {b: 1, sz: 24, color: SG_C.NAVY}) + sg_run_(new Array(46).join('…'), {sz: 24}), {jc: null, after: 60, keep: 1});
  b += sg_signs_([['مسؤول لجنة الاستبيانات', sg_name_(off, 'مسؤول لجنة الاستبيانات') || 'د. كريم عادل'], ['مدير وحدة ضمان الجودة والاعتماد', sg_name_(off, 'مدير وحدة ضمان الجودة والاعتماد')], ['عميد الكلية', sg_name_(off, 'عميد الكلية')]]);
  return b;
}

/* ---------- توليد استبانة واحدة ---------- */
function sg_generate_(rowIndex, row, ctx){
  var n = row[0], form = +sg_w_(row[SG.COL.FORM - 1]), msg = [], links = {}, folder = sg_folder_(row);
  // ١) من الاستمارة: التحليل والتقرير وحفظ نسب العبارات
  var url = row[6], res = null;
  if (url) { try { var fm = FormApp.openByUrl(url); if (fm.getResponses().length) res = ae_compute_(fm); } catch (e) { msg.push('تعذّر فتح الاستمارة'); } }
  if (res && res.count) {
    var meta = {year: ae_clean_(row[1]), std: ae_clean_(row[2]), title: ae_clean_(row[3]), aud: ae_clean_(row[12]) || 'المشاركين', pop: row[13]};
    var base = ae_safe_(meta.title);
    var a = ae_put_(folder, 'تحليل نتائج ' + base + '.docx', ae_analysis_(ctx.tplA, meta, res)); ae_pdf_(folder, 'تحليل نتائج ' + base + '.pdf', a); links.al = a.getUrl();
    var tos = ae_clean_(row[14] || 'عميد الكلية').split(/[،,]\s*/).filter(String), rf = tos.length > 1 ? sg_sub_(folder, 'التقارير') : folder;
    tos.forEach(function(role){ var nm = ctx.names[meta.year + '|' + role]; var nmf = 'تقرير نتائج ' + base + (tos.length > 1 ? ' - ' + ae_safe_(role) : ''); var r = ae_put_(rf, nmf + '.docx', ae_report_(ctx.tplR, meta, res, role + (nm ? ' ' + nm : ''))); ae_pdf_(rf, nmf + '.pdf', r); });
    links.rl = rf.getUrl(); links.cnt = res.count;
    if (form && ctx.lib[form]) {
      var body = sg_bodyOf_(form, row[3]), out = [];
      res.sections.forEach(function(s){ s.rows.forEach(function(r){
        var best = null, bs = 0; ctx.lib[form].forEach(function(x){ var s2 = sg_sim_(r.text, x.t); if (s2 > bs) { bs = s2; best = x; } });
        out.push([n, row[1], body, 'استمارة', form, bs >= .6 ? best.j : '', r.text, Math.round(r.pct * 100) / 100, res.count, sg_num_(row[13]) || '', '']);
      }); });
      sg_putResults_(n, out); ctx.res = sg_results_();
    }
  }
  // ٢) البطاقة وخطط التحسين والمتابعة من نسب العبارات المحفوظة
  var made = 0;
  if (form && ctx.lib[form]) {
    var plans = sg_plans_(n, row, ctx.res, ctx.lib, ctx.th, ctx.thr, ctx.off);
    if (plans.length) {
      var docs = [];
      var f1 = sg_save_(folder, '١ بطاقة بيانات تطبيق الاستبانة.docx', sg_build_('quality', sg_card_(row, plans, ctx.thr, form, ctx.lib[form].length, ctx.off), '١ بطاقة بيانات تطبيق الاستبانة.docx')); docs.push([n, row[1], 'card', '', f1.getName(), f1.getUrl(), folder.getUrl()]);
      plans.forEach(function(p){
        var bt = SG_B[p.b][0];
        var n2 = '٢ رد الجهة وخطة التحسين – ' + bt + '.docx', n3 = '٣ تقرير متابعة تنفيذ خطة التحسين – ' + bt + '.docx';
        if (p.need) n2 = '٢ خطة التدريب المبنية على نتائج الاستبانة – ' + bt + '.docx';
        var f2 = sg_save_(folder, n2, sg_build_(SG_B[p.b][1], p.need ? sg_needsPlan_(row, p, ctx.thr, ctx.off) : sg_plan_(row, p, ctx.thr, ctx.off), n2));
        var f3 = sg_save_(folder, n3, sg_build_('quality', sg_follow_(row, p, ctx.off), n3));
        docs.push([n, row[1], 'plan', p.b, n2, f2.getUrl(), folder.getUrl()]); docs.push([n, row[1], 'follow', p.b, n3, f3.getUrl(), folder.getUrl()]);
      });
      made = docs.length;
      var dsh = sg_sh_(SG.DOCS), dv = dsh.getDataRange().getValues(), keep = [dv[0]];
      dv.slice(1).forEach(function(r){ if (String(r[0]) != String(n)) keep.push(r); }); docs.forEach(function(d){ keep.push(d); });
      dsh.clearContents(); dsh.getRange(1, 1, keep.length, 7).setValues(keep);
    }
  } else if (!form) msg.push('غير مرتبطة بنموذج في الدليل، فلم تُولَّد خطة');
  // ٣) تحديث السجل
  var sh = sg_sh_(SG.REG);
  if (links.cnt != null) sh.getRange(rowIndex, 16, 1, 4).setValues([[links.cnt, Utilities.formatDate(new Date(), 'Africa/Cairo', 'yyyy-MM-dd HH:mm'), links.al, links.rl]]);
  else sh.getRange(rowIndex, 17).setValue(Utilities.formatDate(new Date(), 'Africa/Cairo', 'yyyy-MM-dd HH:mm'));
  if ((links.cnt || made) && row[11] != 'تم رد الجهة') sh.getRange(rowIndex, 12).setValue('النتائج جاهزة');
  sh.getRange(rowIndex, 21).setValue(folder.getUrl());
  if (!links.cnt && !made) msg.push('لا توجد ردود ولا نتائج محفوظة');
  return (made ? 'تم: ' + (links.cnt ? 'التحليل والتقرير و' : '') + sg_ar_(made) + ' مستند' : 'لم يُولَّد شيء') + (msg.length ? ' – ' + msg.join('، ') : '');
}

/* ---------- طابور الطلبات ---------- */
function aeQueue(){
  var lock = LockService.getScriptLock(); if (!lock.tryLock(5000)) return;
  try {
    var sh = sg_sh_(SG.REG), last = sh.getLastRow(); if (last < 2) return;
    var q = sh.getRange(2, SG.COL.REQ, last - 1, 1).getValues(), todo = [];
    q.forEach(function(r, i){ if (r[0] !== '' && r[0] != null) todo.push(i + 2); });
    if (!todo.length) return;
    var t0 = Date.now(), ctx = {lib: sg_lib_(), th: sg_themes_(), thr: sg_thr_(), off: sg_officials_(), res: sg_results_(), names: ae_names_(), tplA: ae_tpl_(AE.TPL_A), tplR: ae_tpl_(AE.TPL_R)};
    for (var k = 0; k < todo.length; k++) {
      if (Date.now() - t0 > 4 * 60 * 1000) break;
      var ri = todo[k], row = sh.getRange(ri, 1, 1, SG.COL.MSG).getValues()[0], m;
      try { m = sg_generate_(ri, row, ctx); } catch (e) { m = 'تعذّر: ' + e.message; }
      sh.getRange(ri, SG.COL.REQ, 1, 2).setValues([['', m + ' (' + Utilities.formatDate(new Date(), 'Africa/Cairo', 'yyyy-MM-dd HH:mm') + ')']]);
    }
  } finally { lock.releaseLock(); }
}
function aeQueueInstall(){
  ScriptApp.getProjectTriggers().forEach(function(t){ if (/^ae(Run|Queue)$/.test(t.getHandlerFunction())) ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('aeQueue').timeBased().everyMinutes(5).create();
  Logger.log('queue trigger installed');
}
