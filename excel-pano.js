/* Satış Panosu · Performans — DİLİMLEYİCİLİ EXCEL DASHBOARD üretici (04.11.2026)
   Build yok, dış kütüphane yok (ES module; yalnız "Excel Dashboard" düğmesine basılınca yüklenir).
   Girdi: performans.html'in hazırladığı düz veri — yalnız oturumdaki kullanıcının yetkili olduğu markanın karneleri.
   Çıktı: .xlsx Blob. Dosya yapısı (tablo dilimleyicisi, x15) Excel'in kendi kaydettiği dosyadan birebir alınmıştır.
   Hesap mantığı: "Veri" tablosunda her satır bir karne × bir metrik. "Görünür" sütunu =SUBTOTAL(103,…) dilimleyici
   süzgecinde 1/0 olur; Pano'daki tüm formüller SUMIFS/COUNTIFS(…, Veri[Görünür], 1) ile yalnız seçili satırları sayar.
   Kurallar Performans Panosu ile aynı: "Etkisi Yok" aylar hedef/gerçekleşen toplamına DAHİL, Başarılı/Yetersiz
   sayımına HARİÇ; gerçekleşme bandı ≥%100 · %80–99 · <%80; hedef ve gerçekleşen ikisi de girilmiş satırlar toplanır. */

const AYLAR = ['Ocak', 'Şubat', 'Mart', 'Nisan', 'Mayıs', 'Haziran', 'Temmuz', 'Ağustos', 'Eylül', 'Ekim', 'Kasım', 'Aralık'];
const AY_KISA = ['Oca', 'Şub', 'Mar', 'Nis', 'May', 'Haz', 'Tem', 'Ağu', 'Eyl', 'Eki', 'Kas', 'Ara'];
const CEYREK = a => (Math.floor(a / 3) + 1) + '. Çeyrek';
const YZ = '%0';   /* Türkçe yüzde yazımı: %94 */
const DEG = { B: 'Başarılı', Y: 'Yetersiz', EY: 'Etkisi Yok' };
const KISA = { arac_satis: 'Araç Satış', takas: 'Takas Alım', car_care: 'Car Care', dis_talep: 'Dış Talep', lead_donusum: 'Lead Dönüşüm',
  giden_arama: 'Giden Arama', teklif: 'Teklif', nps_csi: 'NPS / CSI', gizli_musteri: 'Gizli Müşteri', sigorta: 'Sigorta' };

/* ---------------- XML yardımcıları ---------------- */
const X = s => String(s ?? '').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, '')
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const KOL = n => { let s = ''; n++; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
const H = (c, r) => KOL(c) + r;                       /* c: 0 tabanlı sütun, r: 1 tabanlı satır */
const A = (c, r) => '$' + KOL(c) + '$' + r;
const ARALIK = (c1, r1, c2, r2) => A(c1, r1) + ':' + A(c2, r2);
const XML_BAS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS = {
  main: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  x14: 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main',
  x15: 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main',
  xdr: 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  c: 'http://schemas.openxmlformats.org/drawingml/2006/chart'
};
const RT = {
  belge: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  core: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
  app: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
  sayfa: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
  stil: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
  tema: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  sst: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
  tablo: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table',
  cizim: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
  grafik: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
  dilim: 'http://schemas.microsoft.com/office/2007/relationships/slicer',
  dilimOnbellek: 'http://schemas.microsoft.com/office/2007/relationships/slicerCache'
};
const CT = {
  kitap: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
  sayfa: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
  stil: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
  tema: 'application/vnd.openxmlformats-officedocument.theme+xml',
  sst: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
  tablo: 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml',
  cizim: 'application/vnd.openxmlformats-officedocument.drawing+xml',
  grafik: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
  dilim: 'application/vnd.ms-excel.slicer+xml',
  dilimOnbellek: 'application/vnd.ms-excel.slicerCache+xml',
  core: 'application/vnd.openxmlformats-package.core-properties+xml',
  app: 'application/vnd.openxmlformats-officedocument.extended-properties+xml'
};
const iliskiler = list => XML_BAS + `<Relationships xmlns="${NS.rel}">` + list.map(([id, tip, hedef]) => `<Relationship Id="${id}" Type="${tip}" Target="${X(hedef)}"/>`).join('') + '</Relationships>';

/* Renkler (Kokpit paleti) */
const R = { LAC9: '132A46', LAC7: '1B3A5E', LAC5: '33649A', LAC2: 'BFCDE2', LAC1: 'E2E9F3', LAC05: 'F1F5FA', PIR: 'B98B37', PIR7: '7D5C1F',
  MET1: '151A23', MET2: '4F5868', MET3: '66707F', CIZ: 'D8DDE5', CIZ1: 'E4E8EE', SOLUK: '929CAB', SOLUK2: 'BCC4CF', BEYAZ: 'FFFFFF',
  OK: '23955F', OK7: '136B45', OK1: 'DCF1E5', UY: 'C9831C', UY7: '8A560D', UY1: 'FBEDD2', HT: 'CC4444', HT7: '9D2B2B', HT1: 'F9DFDD', EY: 'B5BCC7' };

/* ---------------- ZIP (deflate: tarayıcının CompressionStream'i; yoksa sıkıştırmasız) ---------------- */
const CRC_T = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
const crc32 = u => { let c = 0xFFFFFFFF; for (let i = 0; i < u.length; i++) c = CRC_T[(c ^ u[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };
async function sikistir(u) {
  if (typeof CompressionStream === 'undefined') return null;
  try { const cs = new CompressionStream('deflate-raw'); const w = cs.writable.getWriter(); w.write(u); w.close(); return new Uint8Array(await new Response(cs.readable).arrayBuffer()); }
  catch (e) { return null; }
}
async function zipYap(dosyalar) {
  const enc = new TextEncoder(), parca = [], merkez = []; let ofs = 0;
  const d = new Date(), saat = (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1), gun = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate();
  for (const [ad, metin] of dosyalar) {
    const adB = enc.encode(ad), ham = enc.encode(metin), crc = crc32(ham);
    let veri = await sikistir(ham), yontem = 8;
    if (!veri || veri.length >= ham.length) { veri = ham; yontem = 0; }
    const y = new DataView(new ArrayBuffer(30));
    y.setUint32(0, 0x04034b50, true); y.setUint16(4, 20, true); y.setUint16(6, 0, true); y.setUint16(8, yontem, true); y.setUint16(10, saat, true); y.setUint16(12, gun, true);
    y.setUint32(14, crc, true); y.setUint32(18, veri.length, true); y.setUint32(22, ham.length, true); y.setUint16(26, adB.length, true); y.setUint16(28, 0, true);
    parca.push(new Uint8Array(y.buffer), adB, veri);
    const m = new DataView(new ArrayBuffer(46));
    m.setUint32(0, 0x02014b50, true); m.setUint16(4, 20, true); m.setUint16(6, 20, true); m.setUint16(8, 0, true); m.setUint16(10, yontem, true); m.setUint16(12, saat, true); m.setUint16(14, gun, true);
    m.setUint32(16, crc, true); m.setUint32(20, veri.length, true); m.setUint32(24, ham.length, true); m.setUint16(28, adB.length, true);
    m.setUint16(30, 0, true); m.setUint16(32, 0, true); m.setUint16(34, 0, true); m.setUint16(36, 0, true); m.setUint32(38, 0, true); m.setUint32(42, ofs, true);
    merkez.push(new Uint8Array(m.buffer), adB);
    ofs += 30 + adB.length + veri.length;
  }
  const mBoy = merkez.reduce((s, x) => s + x.length, 0), son = new DataView(new ArrayBuffer(22));
  son.setUint32(0, 0x06054b50, true); son.setUint16(8, dosyalar.length, true); son.setUint16(10, dosyalar.length, true); son.setUint32(12, mBoy, true); son.setUint32(16, ofs, true);
  return new Blob([...parca, ...merkez, new Uint8Array(son.buffer)], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

/* ---------------- Stil kaydı (styles.xml) ---------------- */
function stilKaydi() {
  const fonts = ['<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>'];
  const fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
  const borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
  const xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
  const dxfs = [], ozelBicim = [];   /* özel sayı biçimleri (164+); dosyada biçim kodu yerel değil, Excel görüntülerken yerelleştirir */
  const yer = (arr, x) => { let i = arr.indexOf(x); if (i < 0) { arr.push(x); i = arr.length - 1; } return i; };
  const fontXml = f => `<font>${f.b ? '<b/>' : ''}${f.i ? '<i/>' : ''}<sz val="${f.sz || 11}"/><color rgb="FF${f.renk || R.MET1}"/><name val="${f.ad || 'Calibri'}"/><family val="${f.ad === 'Cambria' ? 1 : 2}"/>${(f.ad || 'Calibri') === 'Calibri' ? '<scheme val="minor"/>' : ''}</font>`;
  const kenar = k => { const s = (ad, v) => v ? `<${ad} style="${v[1] || 'thin'}"><color rgb="FF${v[0]}"/></${ad}>` : `<${ad}/>`;
    return `<border>${s('left', k.sol)}${s('right', k.sag)}${s('top', k.ust)}${s('bottom', k.alt)}<diagonal/></border>`; };
  function xf(o) {
    const fontId = o.font ? yer(fonts, fontXml(o.font)) : 0;
    const fillId = o.dolgu ? yer(fills, `<fill><patternFill patternType="solid"><fgColor rgb="FF${o.dolgu}"/><bgColor indexed="64"/></patternFill></fill>`) : 0;
    const borderId = o.kenar ? yer(borders, kenar(o.kenar)) : 0;
    let fmt = o.fmt || 0;
    if (typeof fmt === 'string') { let i = ozelBicim.indexOf(fmt); if (i < 0) { ozelBicim.push(fmt); i = ozelBicim.length - 1; } fmt = 164 + i; }
    const hiz = (o.h || o.v || o.sar || o.girinti) ? `<alignment${o.h ? ` horizontal="${o.h}"` : ''}${o.v ? ` vertical="${o.v}"` : ''}${o.sar ? ' wrapText="1"' : ''}${o.girinti ? ` indent="${o.girinti}"` : ''}/>` : '';
    const x = `<xf numFmtId="${fmt}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"${fmt ? ' applyNumberFormat="1"' : ''}${fontId ? ' applyFont="1"' : ''}${fillId ? ' applyFill="1"' : ''}${borderId ? ' applyBorder="1"' : ''}${hiz ? ' applyAlignment="1">' + hiz + '</xf>' : '/>'}`;
    return yer(xfs, x);
  }
  function dxf(o) {   /* CT_Dxf sırası: font, numFmt, fill */
    return yer(dxfs, `<dxf>${o.renk ? `<font>${o.b ? '<b/>' : ''}<color rgb="FF${o.renk}"/></font>` : ''}${o.bicim ? `<numFmt numFmtId="${o.bicim[0]}" formatCode="${X(o.bicim[1])}"/>` : ''}${o.dolgu ? `<fill><patternFill><bgColor rgb="FF${o.dolgu}"/></patternFill></fill>` : ''}</dxf>`);
  }
  function xml() {
    return XML_BAS + `<styleSheet xmlns="${NS.main}" xmlns:mc="${NS.mc}" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">`
      + (ozelBicim.length ? `<numFmts count="${ozelBicim.length}">${ozelBicim.map((k, i) => `<numFmt numFmtId="${164 + i}" formatCode="${X(k)}"/>`).join('')}</numFmts>` : '')
      + `<fonts count="${fonts.length}">${fonts.join('')}</fonts><fills count="${fills.length}">${fills.join('')}</fills><borders count="${borders.length}">${borders.join('')}</borders>`
      + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
      + `<cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>`
      + `<dxfs count="${dxfs.length}">${dxfs.join('')}</dxfs><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>`
      + `<extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="${NS.x14}"><x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext></extLst></styleSheet>`;
  }
  return { xf, dxf, xml };
}

/* ---------------- Paylaşılan metinler + sayfa yazıcı ---------------- */
function metinKaydi() {
  const list = [], map = new Map(); let say = 0;
  return {
    no(s) { s = String(s); say++; let i = map.get(s); if (i === undefined) { i = list.length; list.push(s); map.set(s, i); } return i; },
    xml() { return XML_BAS + `<sst xmlns="${NS.main}" count="${say}" uniqueCount="${list.length}">` + list.map(s => `<si><t xml:space="preserve">${X(s)}</t></si>`).join('') + '</sst>'; }
  };
}
function sayfaYazici(sst) {
  const hucre = new Map(), yuk = new Map(), gizli = new Set(), birlesik = [], kb = [];
  const koy = (c, r, x) => { if (!hucre.has(r)) hucre.set(r, new Map()); hucre.get(r).set(c, x); };
  const sAt = s => (s ? ` s="${s}"` : '');
  const W = {
    metin(c, r, v, s) { if (v === null || v === undefined || v === '') { W.bos(c, r, s); return; } koy(c, r, `<c r="${H(c, r)}"${sAt(s)} t="s"><v>${sst.no(v)}</v></c>`); },
    sayi(c, r, v, s) { if (v === null || v === undefined || v === '' || !Number.isFinite(+v)) { W.bos(c, r, s); return; } koy(c, r, `<c r="${H(c, r)}"${sAt(s)}><v>${+v}</v></c>`); },
    bos(c, r, s) { if (s) koy(c, r, `<c r="${H(c, r)}"${sAt(s)}/>`); },
    /* deger: sayı | metin | {hata:'#N/A'} | '' — Excel açılışta yeniden hesaplar; önbellek değeri önizlemeler içindir */
    formul(c, r, f, deger, s) {
      let t = '', v = '';
      if (deger && typeof deger === 'object' && deger.hata) { t = ' t="e"'; v = `<v>${deger.hata}</v>`; }
      else if (typeof deger === 'number') { if (Number.isFinite(deger)) v = `<v>${deger}</v>`; }
      else { t = ' t="str"'; v = `<v>${X(deger ?? '')}</v>`; }
      koy(c, r, `<c r="${H(c, r)}"${sAt(s)}${t}><f>${X(f)}</f>${v}</c>`);
    },
    yukseklik(r, px) { yuk.set(r, px); },
    gizle(r) { gizli.add(r); },
    birlestir(c1, r1, c2, r2, s) { birlesik.push(H(c1, r1) + ':' + H(c2, r2)); if (s) for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) if (!(hucre.get(r) || new Map()).has(c)) W.bos(c, r, s); },
    kosul(xml) { kb.push(xml); },
    satirlarXml() {
      const tum = new Set([...hucre.keys(), ...yuk.keys(), ...gizli]);
      return [...tum].sort((a, b) => a - b).map(r => {
        const h = yuk.get(r), cells = hucre.get(r);
        const ic = cells ? [...cells.keys()].sort((a, b) => a - b).map(c => cells.get(c)).join('') : '';
        return `<row r="${r}"${h ? ` ht="${(h * 0.75).toFixed(2)}" customHeight="1"` : ''}${gizli.has(r) ? ' hidden="1"' : ''}>${ic}</row>`;
      }).join('');
    },
    birlesikXml() { return birlesik.length ? `<mergeCells count="${birlesik.length}">${birlesik.map(m => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>` : ''; },
    kosulXml() { return kb.join(''); },
    yuk
  };
  return W;
}

/* ---------------- Tema (lacivert / pirinç vurgular: tablo ve dilimleyici stilleri bunları kullanır) ---------------- */
function temaXml() {
  const renk = (ad, v) => `<a:${ad}><a:srgbClr val="${v}"/></a:${ad}>`;
  const yazi = (latin) => `<a:latin typeface="${latin}"/><a:ea typeface=""/><a:cs typeface=""/>`;
  const gs = (pos, ic) => `<a:gs pos="${pos}"><a:schemeClr val="phClr">${ic}</a:schemeClr></a:gs>`;
  return XML_BAS + `<a:theme xmlns:a="${NS.a}" name="Inciroglu"><a:themeElements><a:clrScheme name="Inciroglu">`
    + '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
    + renk('dk2', R.LAC9) + renk('lt2', R.LAC05) + renk('accent1', R.LAC7) + renk('accent2', R.PIR) + renk('accent3', R.OK) + renk('accent4', R.UY)
    + renk('accent5', R.LAC5) + renk('accent6', '8B93A0') + renk('hlink', '25507E') + renk('folHlink', '7D5C1F') + '</a:clrScheme>'
    + `<a:fontScheme name="Office"><a:majorFont>${yazi('Calibri Light')}</a:majorFont><a:minorFont>${yazi('Calibri')}</a:minorFont></a:fontScheme>`
    + '<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    + `<a:gradFill rotWithShape="1"><a:gsLst>${gs(0, '<a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/>')}${gs(50000, '<a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/>')}${gs(100000, '<a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/>')}</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>`
    + `<a:gradFill rotWithShape="1"><a:gsLst>${gs(0, '<a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/>')}${gs(50000, '<a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/>')}${gs(100000, '<a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/>')}</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst>`
    + '<a:lnStyleLst>' + [6350, 12700, 19050].map(w => `<a:ln w="${w}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>`).join('') + '</a:lnStyleLst>'
    + '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst>'
    + '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill>'
    + `<a:gradFill rotWithShape="1"><a:gsLst>${gs(0, '<a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/>')}${gs(50000, '<a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/>')}${gs(100000, '<a:shade val="63000"/><a:satMod val="120000"/>')}</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme>`
    + '</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>';
}

/* ---------------- Grafik (DrawingML chart) yardımcıları ---------------- */
const txPr = (sz, renk, b) => `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${sz}" b="${b ? 1 : 0}"><a:solidFill><a:srgbClr val="${renk}"/></a:solidFill><a:latin typeface="Calibri"/></a:defRPr></a:pPr><a:endParaRPr lang="tr-TR"/></a:p></c:txPr>`;
const strRef = (f, vals) => `<c:strRef><c:f>${X(f)}</c:f><c:strCache><c:ptCount val="${vals.length}"/>${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${X(v)}</c:v></c:pt>`).join('')}</c:strCache></c:strRef>`;
const numRef = (f, vals) => `<c:numRef><c:f>${X(f)}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${vals.length}"/>${vals.map((v, i) => (typeof v === 'number' && Number.isFinite(v)) ? `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>` : '').join('')}</c:numCache></c:numRef>`;
const etiketler = e => `<c:dLbls><c:numFmt formatCode="${X(e.fmt || 'General')}" sourceLinked="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${txPr(e.sz || 900, e.renk || R.MET1, e.b !== false)}${e.konum ? `<c:dLblPos val="${e.konum}"/>` : ''}<c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>`;
const etiketYok = '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>';
const dolguSp = renk => renk ? `<c:spPr><a:solidFill><a:srgbClr val="${renk}"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr>` : '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>';
function cubukSeri(i, s) {
  return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:v>${X(s.ad)}</c:v></c:tx>${dolguSp(s.renk)}<c:invertIfNegative val="0"/>`
    + (s.etiket ? etiketler(s.etiket) : '') + `<c:cat>${strRef(s.katF, s.kat)}</c:cat><c:val>${numRef(s.degF, s.deg)}</c:val></c:ser>`;
}
function cizgiSeri(i, s) {
  const isaret = s.isaret ? `<c:marker><c:symbol val="circle"/><c:size val="7"/><c:spPr><a:solidFill><a:srgbClr val="${s.isaret}"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:marker>` : '<c:marker><c:symbol val="none"/></c:marker>';
  return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:v>${X(s.ad)}</c:v></c:tx>`
    + `<c:spPr><a:ln w="${s.kalinlik || 28575}" cap="rnd"><a:solidFill><a:srgbClr val="${s.renk}"/></a:solidFill>${s.kesikli ? '<a:prstDash val="dash"/>' : ''}<a:round/></a:ln></c:spPr>`
    + isaret + (s.etiket ? etiketler(s.etiket) : '') + `<c:cat>${strRef(s.katF, s.kat)}</c:cat><c:val>${numRef(s.degF, s.deg)}</c:val><c:smooth val="0"/></c:ser>`;
}
function eksenler(o) {
  const katTx = txPr(900, R.MET2, false);
  const kat = `<c:catAx><c:axId val="${o.katId}"/><c:scaling><c:orientation val="${o.ters ? 'maxMin' : 'minMax'}"/></c:scaling><c:delete val="0"/><c:axPos val="${o.katPos}"/>`
    + `<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>`
    + `<c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="${R.SOLUK2}"/></a:solidFill></a:ln></c:spPr>${katTx}<c:crossAx val="${o.degId}"/><c:crosses val="autoZero"/>`
    + '<c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>';
  const deg = `<c:valAx><c:axId val="${o.degId}"/><c:scaling><c:orientation val="minMax"/>${o.max != null ? `<c:max val="${o.max}"/>` : ''}${o.min != null ? `<c:min val="${o.min}"/>` : ''}</c:scaling>`
    + `<c:delete val="1"/><c:axPos val="${o.degPos}"/><c:numFmt formatCode="${X(o.fmt || 'General')}" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>`
    + `<c:crossAx val="${o.katId}"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
  return kat + deg;
}
function grafikKabuk(icerik, lejant, gizliLejant) {
  return XML_BAS + `<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}"><c:date1904 val="0"/><c:lang val="tr-TR"/><c:roundedCorners val="0"/>`
    + `<c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/>${icerik}<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:plotArea>`
    + (lejant ? `<c:legend><c:legendPos val="t"/>${(gizliLejant || []).map(i => `<c:legendEntry><c:idx val="${i}"/><c:delete val="1"/></c:legendEntry>`).join('')}<c:overlay val="0"/>${txPr(900, R.MET2, false)}</c:legend>` : '')
    + '<c:plotVisOnly val="0"/><c:dispBlanksAs val="gap"/></c:chart>'
    + `<c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="${R.CIZ}"/></a:solidFill></a:ln></c:spPr>`
    + `${txPr(900, R.MET2, false)}<c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings></c:chartSpace>`;
}

/* ---------------- Çizim (grafik + dilimleyici yerleşimi) ---------------- */
const EMU = 9525;
function capa(k) { /* k: {c, cOff, r, rOff} — c/r 0 tabanlı, ofsetler piksel */
  return `<xdr:col>${k.c}</xdr:col><xdr:colOff>${Math.round(k.cOff * EMU)}</xdr:colOff><xdr:row>${k.r}</xdr:row><xdr:rowOff>${Math.round(k.rOff * EMU)}</xdr:rowOff>`;
}
function grafikCapa(id, rid, ad, bas, son) {
  return `<xdr:twoCellAnchor editAs="oneCell"><xdr:from>${capa(bas)}</xdr:from><xdr:to>${capa(son)}</xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${X(ad)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>`
    + '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
    + `<c:chart xmlns:c="${NS.c}" xmlns:r="${NS.r}" r:id="${rid}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;
}
function dilimCapa(id, ad, bas, son, px) {
  return `<xdr:twoCellAnchor editAs="absolute"><xdr:from>${capa(bas)}</xdr:from><xdr:to>${capa(son)}</xdr:to>`
    + `<mc:AlternateContent xmlns:mc="${NS.mc}"><mc:Choice xmlns:sle15="http://schemas.microsoft.com/office/drawing/2012/slicer" Requires="sle15">`
    + `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${X(ad)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>`
    + `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2010/slicer"><sle:slicer xmlns:sle="http://schemas.microsoft.com/office/drawing/2010/slicer" name="${X(ad)}"/></a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice>`
    + `<mc:Fallback><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="0" name=""/><xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="${Math.round(px.x * EMU)}" y="${Math.round(px.y * EMU)}"/><a:ext cx="${Math.round(px.w * EMU)}" cy="${Math.round(px.h * EMU)}"/></a:xfrm>`
    + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:prstClr val="white"/></a:solidFill><a:ln w="1"><a:solidFill><a:prstClr val="green"/></a:solidFill></a:ln></xdr:spPr>'
    + '<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip"/><a:lstStyle/><a:p><a:r><a:rPr lang="tr-TR" sz="1100"/><a:t>Bu şekil bir tablo dilimleyicisidir. Dilimleyiciler Excel 2013 ve sonraki sürümlerde çalışır.</a:t></a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback></mc:AlternateContent><xdr:clientData/></xdr:twoCellAnchor>';
}

/* ================================================================================================
   ANA İŞLEV
   G = { marka, yil, hazir, metrikler:[{key,ad,tip}], danlar:[{ad,eski}], aylar:[0..11],
         satirlar:[{ay,dan,eski,met,h,f,d,taslak,manuel}], secim:{sütunAdı:[değerler]}, taslakVar, kaynakNot }
   ================================================================================================ */
export async function panoXlsx(G) {
  const stil = stilKaydi(), sst = metinKaydi();
  const sayiMi = v => typeof v === 'number' && Number.isFinite(v);
  /* ---- metrikler: sayısal hedefliler önce (grafik aralığı bitişik kalsın), kısa ad slicer'a sığsın ---- */
  const kisaAd = m => m.key === 'kasko' ? (/sigorta/i.test(m.ad) ? 'Kasko / Sigorta' : 'Kasko') : (KISA[m.key] || String(m.ad).slice(0, 18));
  const METS = [...G.metrikler.filter(m => m.tip !== 'ozel'), ...G.metrikler.filter(m => m.tip === 'ozel')].map(m => ({ ...m, kisa: kisaAd(m) }));
  const kullanilan = new Set(); METS.forEach(m => { let k = m.kisa, i = 2; while (kullanilan.has(k)) k = m.kisa + ' ' + (i++); m.kisa = k; kullanilan.add(k); });
  const MK = Object.fromEntries(METS.map(m => [m.key, m]));
  const nS = METS.filter(m => m.tip !== 'ozel').length;
  const danAd = d => d.ad + (d.eski ? ' (Eski)' : '');
  const DANS = G.danlar.map(d => ({ ad: danAd(d), eski: !!d.eski }));
  const AYV = G.aylar.slice().sort((a, b) => a - b);

  /* ---- Veri satırları ---- */
  const V = G.satirlar.filter(s => MK[s.met]).map(s => {
    const h = sayiMi(s.h) ? s.h : null, f = sayiMi(s.f) ? s.f : null;
    return { yil: G.yil, ceyrek: CEYREK(s.ay), ay: AYLAR[s.ay], ayNo: s.ay + 1, dan: danAd({ ad: s.dan, eski: s.eski }), calisan: s.eski ? 'Eski' : 'Aktif',
      met: MK[s.met].kisa, metAd: MK[s.met].ad, tur: MK[s.met].tip === 'ozel' ? 'Müdür değerlendirmesi' : 'Sayısal hedefli',
      h, f, oran: (h !== null && f !== null && h > 0) ? f / h : '', hf: (h !== null && f !== null) ? 1 : 0, deg: DEG[s.d] || 'Etkisi Yok',
      karne: s.taslak ? 'Taslak' : 'Kesinleşen', not: s.manuel ? 'Fiili elle girildi' : '' };
  });
  if (!V.length) throw new Error('Dashboard için karne verisi yok.');
  /* Açılış seçimi: G.secim = {ceyrek:[1-4], ay:[0-11], dan:[ad], met:[key], kesin:bool} → sütun değerleri.
     Tablo süzgecine (filterColumn) + gizli satırlara yazılır; dilimleyici seçimi bu süzgeçten okunur. */
  const KOLON = [['Yıl', 'yil'], ['Çeyrek', 'ceyrek'], ['Ay', 'ay'], ['Ay No', 'ayNo'], ['Danışman', 'dan'], ['Çalışan Durumu', 'calisan'], ['Metrik', 'met'], ['Metrik Adı', 'metAd'],
    ['Metrik Türü', 'tur'], ['Hedef', 'h'], ['Gerçekleşen', 'f'], ['Gerçekleşme', 'oran'], ['Değerlendirme', 'deg'], ['Karne Durumu', 'karne'], ['Not', 'not'], ['HF', 'hf'], ['Görünür', 'gor']];
  const KI = Object.fromEntries(KOLON.map(([ad], i) => [ad, i]));
  const SH = G.secim || {}, ekran = {}, karneSec = {};
  const ekle = (hedef, kol, degerler) => { const d = (degerler || []).map(String).filter(Boolean); if (d.length) hedef[kol] = new Set(d); };
  if (SH.ceyrek) ekle(ekran, 'Çeyrek', SH.ceyrek.map(q => q + '. Çeyrek'));
  if (SH.ay) ekle(ekran, 'Ay', SH.ay.map(a => AYLAR[a]));
  if (SH.dan) ekle(ekran, 'Danışman', SH.dan.map(ad => danAd(G.danlar.find(x => x.ad === ad) || { ad, eski: false })));
  if (SH.met) ekle(ekran, 'Metrik', SH.met.filter(k => MK[k]).map(k => MK[k].kisa));
  if (SH.kesin && G.taslakVar) ekle(karneSec, 'Karne Durumu', ['Kesinleşen']);
  const gorunurMu = (v, sec) => Object.entries(sec).every(([kol, set]) => set.has(String(v[KOLON[KI[kol]][1]])));
  /* Seçimle hiç satır kalmıyorsa önce ekran seçimi, sonra karne durumu bırakılır (dosya boş açılmasın) */
  let secim = Object.assign({}, ekran, karneSec), ekranDustu = false, karneDustu = false;
  if (!V.some(v => gorunurMu(v, secim))) { ekranDustu = Object.keys(ekran).length > 0; secim = Object.assign({}, karneSec); }
  if (!V.some(v => gorunurMu(v, secim))) { karneDustu = Object.keys(karneSec).length > 0; secim = {}; }
  V.forEach(v => { v.gor = gorunurMu(v, secim) ? 1 : 0; });
  /* Kesinleşmiş karne hiç yoksa dosya bunu açıkça söyler (denetim bulgusu: "yalnız kesinleşenler seçili" yazıp taslak gösteriyordu) */
  const KNOT = karneDustu ? 'kesinleşmiş karne yok — taslak karneler gösteriliyor' : G.kaynakNot;
  const GV = V.filter(v => v.gor);

  /* ---- Önbellek değerleri (Excel formüllerinin açılıştaki karşılığı) ---- */
  const topla = (liste, alan) => liste.reduce((s, v) => s + (v[alan] || 0), 0);
  const sayac = (liste, deg) => liste.filter(v => v.deg === deg).length;
  const hfTop = liste => { const l = liste.filter(v => v.hf === 1); return { h: topla(l, 'h'), f: topla(l, 'f') }; };
  const metHes = METS.map(m => { const l = GV.filter(v => v.met === m.kisa), t = hfTop(l);
    return { m, h: t.h, f: t.f, oran: m.tip !== 'ozel' && t.h > 0 ? t.f / t.h : '', B: sayac(l, 'Başarılı'), Y: sayac(l, 'Yetersiz'), EY: sayac(l, 'Etkisi Yok'), kayit: l.length,
      bayrak: m.tip !== 'ozel' && t.h > 0 ? 1 : 0 }; });
  const tekMetrik = metHes.filter(x => x.bayrak).length === 1 ? metHes.find(x => x.bayrak).m.kisa : '';
  const arac = metHes.find(x => x.m.key === 'arac_satis');
  const gos = tekMetrik || (arac && arac.bayrak ? arac.m.kisa : '');
  const gosT = gos ? hfTop(GV.filter(v => v.met === gos)) : { h: 0, f: 0 };
  const basari = (B, Y) => (B + Y) ? B / (B + Y) : '';
  const bant = o => o === '' ? '' : (o >= 1 ? 'Hedefte' : o >= 0.8 ? 'Hedefe yakın' : 'Hedefin altında');
  const ayHes = AYV.map(a => { const l = GV.filter(v => v.ayNo === a + 1), t = gos ? hfTop(l.filter(v => v.met === gos)) : null;
    return { a, h: t ? t.h : '', f: t ? t.f : '', oran: t && t.h > 0 ? t.f / t.h : '', B: sayac(l, 'Başarılı'), Y: sayac(l, 'Yetersiz'), EY: sayac(l, 'Etkisi Yok'), kayit: l.length }; });
  const danHes = DANS.map(d => { const l = GV.filter(v => v.dan === d.ad), t = gos ? hfTop(l.filter(v => v.met === gos)) : null;
    return { d, h: t ? t.h : '', f: t ? t.f : '', oran: t && t.h > 0 ? t.f / t.h : '', B: sayac(l, 'Başarılı'), Y: sayac(l, 'Yetersiz'), EY: sayac(l, 'Etkisi Yok'), kayit: l.length }; });
  const TB = topla(metHes, 'B'), TY = topla(metHes, 'Y'), TEY = topla(metHes, 'EY');

  /* ======================= STİLLER ======================= */
  const S = {
    bant: stil.xf({ dolgu: R.LAC9 }),
    baslik: stil.xf({ font: { ad: 'Cambria', sz: 20, b: 1, renk: R.BEYAZ }, dolgu: R.LAC9, v: 'center', girinti: 1 }),
    alt: stil.xf({ font: { sz: 10, renk: R.LAC2 }, dolgu: R.LAC9, v: 'center', girinti: 1 }),
    secim: stil.xf({ font: { sz: 10, b: 1, renk: R.LAC7 }, dolgu: R.LAC05, v: 'center', girinti: 1, kenar: { alt: [R.LAC1] } }),
    kpiEt: stil.xf({ font: { sz: 9, b: 1, renk: R.MET3 }, dolgu: R.LAC05, v: 'bottom', girinti: 1, kenar: { ust: [R.PIR, 'medium'], sol: [R.BEYAZ, 'thick'], sag: [R.BEYAZ, 'thick'] } }),
    kpiYz: stil.xf({ font: { sz: 22, b: 1, renk: R.LAC9 }, dolgu: R.LAC05, fmt: YZ, h: 'left', v: 'center', girinti: 1, kenar: { sol: [R.BEYAZ, 'thick'], sag: [R.BEYAZ, 'thick'] } }),
    kpiMt: stil.xf({ font: { sz: 22, b: 1, renk: R.LAC9 }, dolgu: R.LAC05, h: 'left', v: 'center', girinti: 1, kenar: { sol: [R.BEYAZ, 'thick'], sag: [R.BEYAZ, 'thick'] } }),
    kpiAlt: stil.xf({ font: { sz: 9, renk: R.MET2 }, dolgu: R.LAC05, v: 'top', girinti: 1, kenar: { sol: [R.BEYAZ, 'thick'], sag: [R.BEYAZ, 'thick'] } }),
    grB: stil.xf({ font: { sz: 12, b: 1, renk: R.LAC9 }, v: 'bottom' }),
    grA: stil.xf({ font: { sz: 9, renk: R.MET3 }, v: 'top' }),
    bolum: stil.xf({ font: { sz: 12, b: 1, renk: R.LAC9 }, v: 'bottom', kenar: { alt: [R.PIR, 'medium'] } }),
    th: stil.xf({ font: { sz: 9, b: 1, renk: R.BEYAZ }, dolgu: R.LAC7, h: 'center', v: 'center', sar: 1 }),
    thSol: stil.xf({ font: { sz: 9, b: 1, renk: R.BEYAZ }, dolgu: R.LAC7, h: 'left', v: 'center', girinti: 1 }),
    td: stil.xf({ font: { sz: 10, renk: R.MET1 }, v: 'center', girinti: 1, kenar: { alt: [R.CIZ1] } }),
    tdEski: stil.xf({ font: { sz: 10, renk: R.HT7 }, v: 'center', girinti: 1, kenar: { alt: [R.CIZ1] } }),
    tdSayi: stil.xf({ font: { sz: 10, renk: R.MET1 }, fmt: 3, h: 'right', v: 'center', girinti: 1, kenar: { alt: [R.CIZ1] } }),
    tdYz: stil.xf({ font: { sz: 10, b: 1, renk: R.MET1 }, fmt: YZ, h: 'center', v: 'center', kenar: { alt: [R.CIZ1] } }),
    tdOrta: stil.xf({ font: { sz: 10, renk: R.MET2 }, h: 'center', v: 'center', kenar: { alt: [R.CIZ1] } }),
    tdSoluk: stil.xf({ font: { sz: 10, renk: R.SOLUK }, h: 'center', v: 'center', kenar: { alt: [R.CIZ1] } }),
    topMt: stil.xf({ font: { sz: 10, b: 1, renk: R.LAC9 }, dolgu: R.LAC05, v: 'center', girinti: 1, kenar: { ust: [R.LAC7] } }),
    topSayi: stil.xf({ font: { sz: 10, b: 1, renk: R.LAC9 }, dolgu: R.LAC05, fmt: 3, h: 'right', v: 'center', girinti: 1, kenar: { ust: [R.LAC7] } }),
    topYz: stil.xf({ font: { sz: 10, b: 1, renk: R.LAC9 }, dolgu: R.LAC05, fmt: YZ, h: 'center', v: 'center', kenar: { ust: [R.LAC7] } }),
    topOrta: stil.xf({ font: { sz: 10, b: 1, renk: R.LAC9 }, dolgu: R.LAC05, h: 'center', v: 'center', kenar: { ust: [R.LAC7] } }),
    notB: stil.xf({ font: { sz: 10, b: 1, renk: R.LAC9 }, v: 'bottom' }),
    not: stil.xf({ font: { sz: 9, renk: R.MET2 }, v: 'top', sar: 1 }),
    gizliMt: stil.xf({ font: { sz: 9, renk: R.MET3 } }),
    gizliYz: stil.xf({ font: { sz: 9, renk: R.MET3 }, fmt: YZ }),
    vBas: stil.xf({ font: { sz: 14, b: 1, renk: R.LAC9 } }),
    vAlt: stil.xf({ font: { sz: 9, renk: R.MET3 } }),
    vYz: stil.xf({ fmt: YZ, h: 'center' }),
    vSayi: stil.xf({ h: 'right' })
  };
  const D = { ok: stil.dxf({ renk: R.OK7, dolgu: R.OK1, b: 1 }), uy: stil.dxf({ renk: R.UY7, dolgu: R.UY1, b: 1 }), ht: stil.dxf({ renk: R.HT7, dolgu: R.HT1, b: 1 }), soluk: stil.dxf({ renk: R.SOLUK2 }),
    ondalik: stil.dxf({ bicim: [200, '#,##0.0'] }) };   /* hedef/gerçekleşen tam sayı #.##0; ondalıklıysa koşullu olarak #.##0,0 */

  /* ======================= VERİ SAYFASI ======================= */
  const W2 = sayfaYazici(sst);
  const T0 = 4, TSON = T0 + V.length, NK = KOLON.length;
  W2.metin(0, 1, 'Veri — karnelere girilen metrik değerleri', S.vBas); W2.yukseklik(1, 26);
  W2.metin(0, 2, `${G.marka} · ${G.yil} · her satır bir karne × bir metrik. Pano'daki dilimleyiciler bu tabloyu süzer; gizlenen satırlar hesaba girmez. Değerleri değiştirmeyin — güncel veri için panodan yeniden indirin.`, S.vAlt);
  KOLON.forEach(([ad], i) => W2.metin(i, T0, ad));
  const tr = n => `Veri[[#This Row],[${n}]]`;
  const F_ORAN = `IF(AND(ISNUMBER(${tr('Hedef')}),ISNUMBER(${tr('Gerçekleşen')})),IF(${tr('Hedef')}>0,${tr('Gerçekleşen')}/${tr('Hedef')},""),"")`;
  const F_HF = `IF(AND(ISNUMBER(${tr('Hedef')}),ISNUMBER(${tr('Gerçekleşen')})),1,0)`;
  const F_GOR = `SUBTOTAL(103,${tr('Yıl')})`;   /* her satırda dolu sütun */
  V.forEach((v, i) => {
    const r = T0 + 1 + i;
    W2.sayi(0, r, v.yil); W2.metin(1, r, v.ceyrek); W2.metin(2, r, v.ay); W2.sayi(3, r, v.ayNo); W2.metin(4, r, v.dan); W2.metin(5, r, v.calisan);
    W2.metin(6, r, v.met); W2.metin(7, r, v.metAd); W2.metin(8, r, v.tur); W2.sayi(9, r, v.h, S.vSayi); W2.sayi(10, r, v.f, S.vSayi);
    W2.formul(11, r, F_ORAN, v.oran, S.vYz); W2.metin(12, r, v.deg); W2.metin(13, r, v.karne); W2.metin(14, r, v.not);
    W2.formul(15, r, F_HF, v.hf); W2.formul(16, r, F_GOR, v.gor);
    if (!v.gor) W2.gizle(r);
  });
  const VERI_GEN = [7, 12, 10, 7, 26, 15, 16, 34, 21, 10, 12, 13, 15, 15, 18, 6, 9];
  const filtreXml = Object.entries(secim).map(([kol, set]) => `<filterColumn colId="${KI[kol]}"><filters>${[...set].map(v => `<filter val="${X(v)}"/>`).join('')}</filters></filterColumn>`).join('');
  const tabloXml = XML_BAS + `<table xmlns="${NS.main}" id="1" name="Veri" displayName="Veri" ref="A${T0}:${KOL(NK - 1)}${TSON}" totalsRowShown="0">`
    + `<autoFilter ref="A${T0}:${KOL(NK - 1)}${TSON}">${filtreXml}</autoFilter><tableColumns count="${NK}">`
    + KOLON.map(([ad], i) => {
      const f = ad === 'Gerçekleşme' ? F_ORAN : ad === 'HF' ? F_HF : ad === 'Görünür' ? F_GOR : null;
      return `<tableColumn id="${i + 1}" name="${X(ad)}"${f ? `><calculatedColumnFormula>${X(f)}</calculatedColumnFormula></tableColumn>` : '/>'}`;
    }).join('') + '</tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>';
  const veriSayfaXml = XML_BAS + `<worksheet xmlns="${NS.main}" xmlns:r="${NS.r}"><sheetPr><tabColor rgb="FF${R.LAC5}"/></sheetPr><dimension ref="A1:${KOL(NK - 1)}${TSON}"/>`
    + `<sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane ySplit="${T0}" topLeftCell="A${T0 + 1}" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A${T0 + 1}" sqref="A${T0 + 1}"/></sheetView></sheetViews>`
    + '<sheetFormatPr defaultRowHeight="15"/>'
    + `<cols>${VERI_GEN.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"${i >= 15 ? ' hidden="1"' : ''}/>`).join('')}</cols>`
    + `<sheetData>${W2.satirlarXml()}</sheetData>`
    + '<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="landscape"/>'
    + '<tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>';

  /* ======================= PANO SAYFASI ======================= */
  const W = sayfaYazici(sst);
  /* Sütunlar (px): A kenar · B dilimleyici şeridi · C boşluk · D–O 12 × 92 ana alan · P kenar · Q–W yardımcı (gizli) */
  const SUT = [14, 238, 14, ...Array(12).fill(92), 14];
  const genislik = px => Math.floor(px / 7 * 256) / 256;
  const sx = c => SUT.slice(0, c).reduce((s, w) => s + w, 0);
  const D0 = 3, O0 = 14;                                /* D ve O sütun indeksleri */
  const Q = 16, RR = 17, SS = 18, TT = 19, UU = 20, VV = 21, WW = 22;
  const satirPx = r => W.yuk.get(r) || 20;
  const sy = r => { let s = 0; for (let i = 1; i < r; i++) s += satirPx(i); return s; };
  const pxSatir = y => { let r = 1, s = 0; while (s + satirPx(r) <= y) { s += satirPx(r); r++; } return { r: r - 1, off: y - s }; };
  const pxSutun = x => { let c = 0, s = 0; while (c < SUT.length && s + SUT[c] <= x) { s += SUT[c]; c++; } return { c, off: x - s }; };
  [[1, 8], [2, 46], [3, 22], [4, 26], [5, 12], [6, 22], [7, 42], [8, 22], [9, 16], [10, 26], [11, 20], [28, 16], [29, 26], [30, 20], [47, 20]].forEach(([r, h]) => W.yukseklik(r, h));

  /* Başlık bandı + seçim satırı */
  for (let c = 1; c <= O0; c++) { W.bos(c, 2, S.bant); W.bos(c, 3, S.bant); }
  W.metin(1, 2, `${G.marka} · Performans Metrik Panosu · ${G.yil}`, S.baslik); W.birlestir(1, 2, O0, 2, S.baslik);
  W.metin(1, 3, `Kaynak: karnelere girilen hedef ve gerçekleşenler · ${KNOT} · Hazırlanma: ${G.hazir} · Yalnız yetkili olduğunuz markanın verisi`, S.alt); W.birlestir(1, 3, O0, 3, S.alt);

  /* ---- Tablo yerleşimi ---- */
  const MT_B = 48, MT_H = 49, MT0 = 50, MT1 = MT0 + METS.length - 1, MT_T = MT1 + 1;
  const AY_B = MT_T + 2, AY_H = AY_B + 1, AY0 = AY_H + 1, AY1 = AY0 + AYV.length - 1, AY_T = AY1 + 1;
  const DN_B = AY_T + 2, DN_H = DN_B + 1, DN0 = DN_H + 1, DN1 = DN0 + DANS.length - 1;
  const NT_B = DN1 + 2;
  const MS1 = MT0 + nS - 1;                              /* sayısal metriklerin son satırı */
  /* Yardımcı hücreler (gizli Q sütunu): seçili tek metrik, gösterilen metrik, sayılar */
  const vf = (kol, ...kr) => `SUMIFS(Veri[${kol}],${kr.join(',')},Veri[Görünür],1)`;
  const cf = (...kr) => `COUNTIFS(${kr.join(',')},Veri[Görünür],1)`;
  const qTek = A(Q, 2), qGos = A(Q, 3), qGosH = A(Q, 4), qGosF = A(Q, 5), qKayit = A(Q, 6), qAy = A(Q, 7), qDan = A(Q, 8), qMet = A(Q, 9);
  W.formul(Q, 2, nS ? `IF(SUM(${ARALIK(VV, MT0, VV, MS1)})=1,INDEX(${ARALIK(Q, MT0, Q, MS1)},MATCH(1,${ARALIK(VV, MT0, VV, MS1)},0)),"")` : '""', tekMetrik, S.gizliMt);
  const aracSatir = arac ? MT0 + METS.indexOf(arac.m) : null;
  W.formul(Q, 3, aracSatir ? `IF(${qTek}<>"",${qTek},IF(${A(VV, aracSatir)}=1,${A(Q, aracSatir)},""))` : `${qTek}`, gos, S.gizliMt);
  W.formul(Q, 4, `IF(${qGos}="",0,${vf('Hedef', `Veri[Metrik],${qGos}`, 'Veri[HF],1')})`, gosT.h, S.gizliMt);
  W.formul(Q, 5, `IF(${qGos}="",0,${vf('Gerçekleşen', `Veri[Metrik],${qGos}`, 'Veri[HF],1')})`, gosT.f, S.gizliMt);
  W.formul(Q, 6, 'SUM(Veri[Görünür])', GV.length, S.gizliMt);
  W.formul(Q, 7, `COUNTIF(${ARALIK(O0, AY0, O0, AY1)},">0")`, ayHes.filter(x => x.kayit > 0).length, S.gizliMt);
  W.formul(Q, 8, `COUNTIF(${ARALIK(O0, DN0, O0, DN1)},">0")`, danHes.filter(x => x.kayit > 0).length, S.gizliMt);
  W.formul(Q, 9, `COUNTIF(${ARALIK(O0, MT0, O0, MT1)},">0")`, metHes.filter(x => x.kayit > 0).length, S.gizliMt);
  const SAYI = h => `IF(${h}=INT(${h}),FIXED(${h},0),FIXED(${h},1))`;

  /* Seçim satırı */
  const secimMetin = `Seçim: ${ayHes.filter(x => x.kayit > 0).length} ay · ${danHes.filter(x => x.kayit > 0).length} danışman · ${metHes.filter(x => x.kayit > 0).length} metrik · ${GV.length} kayıt`;
  W.formul(1, 4, `"Seçim: "&${qAy}&" ay · "&${qDan}&" danışman · "&${qMet}&" metrik · "&FIXED(${qKayit},0)&" kayıt"&IF(${qKayit}=0,"   ⚠ Bu seçimde veri yok — dilimleyicilerdeki filtreyi temizleyin.","   ·   Soldaki dilimleyicilerden seçin: birden çok seçim Ctrl + tık · tümüne dönmek için dilimleyicinin sağ üstündeki filtre temizleme simgesi")`,
    secimMetin + '   ·   Soldaki dilimleyicilerden seçin: birden çok seçim Ctrl + tık · tümüne dönmek için dilimleyicinin sağ üstündeki filtre temizleme simgesi', S.secim);
  W.birlestir(1, 4, O0, 4, S.secim);

  /* ---- KPI kartları (satır 6–8) ---- */
  const kart = (c0, etiketF, etiketV, degerF, degerV, yuzdeMi, altF, altV) => {
    const c1 = c0 + 2;
    if (etiketF) W.formul(c0, 6, etiketF, etiketV, S.kpiEt); else W.metin(c0, 6, etiketV, S.kpiEt);
    W.birlestir(c0, 6, c1, 6, S.kpiEt);
    W.formul(c0, 7, degerF, degerV, yuzdeMi ? S.kpiYz : S.kpiMt); W.birlestir(c0, 7, c1, 7, yuzdeMi ? S.kpiYz : S.kpiMt);
    if (altF) W.formul(c0, 8, altF, altV, S.kpiAlt); else W.metin(c0, 8, altV, S.kpiAlt);
    W.birlestir(c0, 8, c1, 8, S.kpiAlt);
  };
  const KB = ARALIK(10, MT0, 10, MT1), KY = ARALIK(11, MT0, 11, MT1), KE = ARALIK(12, MT0, 12, MT1);
  kart(D0, null, 'BAŞARILI DEĞERLENDİRME ORANI', `IFERROR(SUM(${KB})/(SUM(${KB})+SUM(${KY})),"—")`, (TB + TY) ? TB / (TB + TY) : '—', true,
    `FIXED(SUM(${KB}),0)&" başarılı · "&FIXED(SUM(${KY}),0)&" yetersiz · "&FIXED(SUM(${KE}),0)&" etkisi yok"`, `${TB} başarılı · ${TY} yetersiz · ${TEY} etkisi yok`);
  kart(D0 + 3, `"GERÇEKLEŞME · "&IF(${qGos}="","TEK METRİK SEÇİN",${qGos})`, 'GERÇEKLEŞME · ' + (gos || 'TEK METRİK SEÇİN'),
    `IF(${qGosH}>0,${qGosF}/${qGosH},"—")`, gosT.h > 0 ? gosT.f / gosT.h : '—', true,
    `IF(${qGos}="","Grafik ve tablolar için tek bir sayısal metrik seçin",${SAYI(qGosF)}&" / "&${SAYI(qGosH)}&" · gerçekleşen / hedef")`,
    gos ? `${Number.isInteger(gosT.f) ? gosT.f : gosT.f.toFixed(1)} / ${Number.isInteger(gosT.h) ? gosT.h : gosT.h.toFixed(1)} · gerçekleşen / hedef` : 'Grafik ve tablolar için tek bir sayısal metrik seçin');
  const IR = ARALIK(8, MT0, 8, Math.max(MT0, MS1));
  kart(D0 + 6, null, 'HEDEFİ TUTTURULAN METRİK', nS ? `COUNTIF(${IR},">=1")&" / "&COUNT(${IR})` : '"—"',
    nS ? `${metHes.filter(x => x.oran !== '' && x.oran >= 1).length} / ${metHes.filter(x => x.oran !== '').length}` : '—', false, null, 'sayısal hedefli metrikler · gerçekleşme ≥ %100');
  const DO = ARALIK(O0, DN0, O0, DN1), DG = ARALIK(D0 + 3, DN0, D0 + 3, DN1);
  kart(D0 + 9, null, 'DANIŞMAN', `COUNTIF(${DO},">0")`, danHes.filter(x => x.kayit > 0).length, false,
    `FIXED(COUNTIFS(${DG},"Aktif",${DO},">0"),0)&" aktif · "&FIXED(COUNTIFS(${DG},"Eski",${DO},">0"),0)&" eski çalışan"`,
    `${danHes.filter(x => x.kayit > 0 && !x.d.eski).length} aktif · ${danHes.filter(x => x.kayit > 0 && x.d.eski).length} eski çalışan`);

  /* ---- Grafik başlıkları (hücrede; grafikler altlarında) ---- */
  const grBaslik = (c0, r, f, v, alt) => {
    if (f) W.formul(c0, r, f, v, S.grB); else W.metin(c0, r, v, S.grB);
    if (alt && alt.f) W.formul(c0, r + 1, alt.f, alt.v, S.grA); else W.metin(c0, r + 1, alt, S.grA);
  };
  grBaslik(D0, 10, null, 'Metrik bazında gerçekleşme', 'Gerçekleşen ÷ hedef · yeşil ≥ %100 · sarı %80–99 · kırmızı < %80 · kesikli çizgi: hedef (%100)');
  grBaslik(D0 + 6, 10, null, 'Metrik bazında değerlendirme', 'Karnelerde Başarılı / Yetersiz / Etkisi Yok işaretlenme sayısı (çubuk içinde adet)');
  grBaslik(D0, 29, `IF(${qGos}="","Aylık hedef – gerçekleşen",${qGos}&" · aylık hedef – gerçekleşen")`, gos ? gos + ' · aylık hedef – gerçekleşen' : 'Aylık hedef – gerçekleşen',
    { f: `IF(${qGos}="","Tek bir sayısal metrik seçin (farklı birimler toplanmaz).",IF(${qTek}="","Birden çok metrik seçili: Araç Satış gösteriliyor · tek metrik seçerek değiştirin.","Seçili metriğin aylık toplamı (Etkisi Yok aylar dahil)."))`,
      v: !gos ? 'Tek bir sayısal metrik seçin (farklı birimler toplanmaz).' : (!tekMetrik ? 'Birden çok metrik seçili: Araç Satış gösteriliyor · tek metrik seçerek değiştirin.' : 'Seçili metriğin aylık toplamı (Etkisi Yok aylar dahil).') });
  grBaslik(D0 + 6, 29, null, 'Aylık başarı oranı', 'Başarılı ÷ (Başarılı + Yetersiz) · Etkisi Yok hariç · seçili tüm metrikler');

  /* ---- Tablo ortak başlık / satır yazımı ---- */
  const bolumBaslik = (r, f, v) => { if (f) W.formul(D0, r, f, v, S.bolum); else W.metin(D0, r, v, S.bolum); W.birlestir(D0, r, O0, r, S.bolum); W.yukseklik(r, 28); };
  const thSatir = (r, ilk, digerleri) => { W.metin(D0, r, ilk, S.thSol); W.birlestir(D0, r, D0 + 2, r, S.thSol); digerleri.forEach((t, i) => W.metin(D0 + 3 + i, r, t, S.th)); W.yukseklik(r, 34); };
  const ad3 = (r, metin, s) => { W.metin(D0, r, metin, s); W.birlestir(D0, r, D0 + 2, r, s); W.yukseklik(r, 22); };
  const bantF = h => `IF(ISNUMBER(${h}),IF(${h}>=1,"Hedefte",IF(${h}>=0.8,"Hedefe yakın","Hedefin altında")),"")`;
  const basF = (b, y) => `IF((${b}+${y})>0,${b}/(${b}+${y}),"")`;
  const kosulYuzde = (aralik, ilk) => W.kosul(`<conditionalFormatting sqref="${aralik}">`
    + `<cfRule type="expression" dxfId="${D.ok}" priority="1"><formula>${X(`AND(ISNUMBER(${ilk}),${ilk}>=1)`)}</formula></cfRule>`
    + `<cfRule type="expression" dxfId="${D.uy}" priority="2"><formula>${X(`AND(ISNUMBER(${ilk}),${ilk}>=0.8,${ilk}<1)`)}</formula></cfRule>`
    + `<cfRule type="expression" dxfId="${D.ht}" priority="3"><formula>${X(`AND(ISNUMBER(${ilk}),${ilk}<0.8)`)}</formula></cfRule></conditionalFormatting>`);
  const kosulOndalik = (aralik, ilk) => W.kosul(`<conditionalFormatting sqref="${aralik}"><cfRule type="expression" dxfId="${D.ondalik}" priority="5"><formula>${X(`AND(ISNUMBER(${ilk}),MOD(${ilk},1)<>0)`)}</formula></cfRule></conditionalFormatting>`);
  const kosulSoluk = (aralik, kayitHucre) => W.kosul(`<conditionalFormatting sqref="${aralik}"><cfRule type="expression" dxfId="${D.soluk}" priority="4"><formula>${X(kayitHucre + '=0')}</formula></cfRule></conditionalFormatting>`);
  const deger = v => (v === '' || v === null || v === undefined) ? '' : v;

  /* ---- 1) Metrik özeti ---- */
  bolumBaslik(MT_B, null, 'Metrik özeti — seçime göre');
  thSatir(MT_H, 'Metrik', ['Hedef', 'Gerçekleşen', 'Gerçekleşme', 'Durum', 'Başarılı', 'Yetersiz', 'Etkisi Yok', 'Başarı oranı', 'Kayıt']);
  metHes.forEach((x, i) => {
    const r = MT0 + i, m = x.m, kq = A(Q, r), say = m.tip !== 'ozel';
    ad3(r, m.ad, S.td);
    W.metin(Q, r, m.kisa, S.gizliMt);
    if (say) {
      W.formul(6, r, vf('Hedef', `Veri[Metrik],${kq}`, 'Veri[HF],1'), x.h, S.tdSayi);
      W.formul(7, r, vf('Gerçekleşen', `Veri[Metrik],${kq}`, 'Veri[HF],1'), x.f, S.tdSayi);
      W.formul(8, r, `IF(${H(6, r)}>0,${H(7, r)}/${H(6, r)},"")`, deger(x.oran), S.tdYz);
      W.formul(9, r, bantF(H(8, r)), bant(x.oran), S.tdOrta);
      /* grafik serileri: bant renkli üç seri (yalnız biri dolu, diğerleri 0 = görünmez) + etiket serisi + tek metrik bayrağı.
         #N/A kullanılmaz: sütun grafiğinde "#YOK" etiketi basıyor. */
      W.formul(RR, r, `IF(ISNUMBER(${H(8, r)}),IF(${H(8, r)}>=1,${H(8, r)},0),0)`, x.oran !== '' && x.oran >= 1 ? x.oran : 0, S.gizliYz);
      W.formul(SS, r, `IF(ISNUMBER(${H(8, r)}),IF(AND(${H(8, r)}>=0.8,${H(8, r)}<1),${H(8, r)},0),0)`, x.oran !== '' && x.oran >= 0.8 && x.oran < 1 ? x.oran : 0, S.gizliYz);
      W.formul(TT, r, `IF(ISNUMBER(${H(8, r)}),IF(${H(8, r)}<0.8,${H(8, r)},0),0)`, x.oran !== '' && x.oran < 0.8 ? x.oran : 0, S.gizliYz);
      W.formul(WW, r, `IF(ISNUMBER(${H(8, r)}),${H(8, r)},0)`, x.oran !== '' ? x.oran : 0, S.gizliYz);
      W.sayi(UU, r, 1, S.gizliYz);
      W.formul(VV, r, `IF(${H(6, r)}>0,1,0)`, x.bayrak, S.gizliMt);
    } else {
      W.metin(6, r, '—', S.tdSoluk); W.metin(7, r, '—', S.tdSoluk); W.metin(8, r, '—', S.tdSoluk); W.metin(9, r, 'Değerlendirme', S.tdSoluk);
    }
    W.formul(10, r, cf(`Veri[Metrik],${kq}`, 'Veri[Değerlendirme],"Başarılı"'), x.B, S.tdSayi);
    W.formul(11, r, cf(`Veri[Metrik],${kq}`, 'Veri[Değerlendirme],"Yetersiz"'), x.Y, S.tdSayi);
    W.formul(12, r, cf(`Veri[Metrik],${kq}`, 'Veri[Değerlendirme],"Etkisi Yok"'), x.EY, S.tdSayi);
    W.formul(13, r, basF(H(10, r), H(11, r)), deger(basari(x.B, x.Y)), S.tdYz);
    W.formul(14, r, cf(`Veri[Metrik],${kq}`), x.kayit, S.tdSayi);
  });
  ad3(MT_T, 'Toplam (değerlendirme)', S.topMt);
  [6, 7, 8, 9].forEach(c => W.metin(c, MT_T, '—', S.topOrta));
  [10, 11, 12, 14].forEach(c => W.formul(c, MT_T, `SUM(${ARALIK(c, MT0, c, MT1)})`, c === 10 ? TB : c === 11 ? TY : c === 12 ? TEY : topla(metHes, 'kayit'), S.topSayi));
  W.formul(13, MT_T, basF(H(10, MT_T), H(11, MT_T)), deger(basari(TB, TY)), S.topYz);
  W.yukseklik(MT_T, 24);
  if (nS) { kosulYuzde(`${H(8, MT0)}:${H(8, MS1)}`, H(8, MT0)); kosulOndalik(`${H(6, MT0)}:${H(7, MS1)}`, H(6, MT0)); }
  kosulSoluk(`${H(D0, MT0)}:${H(O0, MT1)}`, '$' + KOL(O0) + MT0);

  /* ---- 2) Aylık özet ---- */
  bolumBaslik(AY_B, `"Aylık özet — "&IF(${qGos}="","hedef / gerçekleşen için tek metrik seçin",${qGos}&" hedef / gerçekleşen")&" · başarı sayıları seçili tüm metriklerden"`,
    'Aylık özet — ' + (gos ? gos + ' hedef / gerçekleşen' : 'hedef / gerçekleşen için tek metrik seçin') + ' · başarı sayıları seçili tüm metriklerden');
  thSatir(AY_H, 'Ay', ['Hedef', 'Gerçekleşen', 'Gerçekleşme', 'Durum', 'Başarılı', 'Yetersiz', 'Etkisi Yok', 'Başarı oranı', 'Kayıt']);
  ayHes.forEach((x, i) => {
    const r = AY0 + i, an = A(UU, r);
    ad3(r, AYLAR[x.a] + ' ' + G.yil, S.td);
    W.metin(Q, r, AY_KISA[x.a], S.gizliMt); W.sayi(UU, r, x.a + 1, S.gizliMt);
    W.formul(6, r, `IF(${qGos}="","",${vf('Hedef', `Veri[Ay No],${an}`, `Veri[Metrik],${qGos}`, 'Veri[HF],1')})`, deger(x.h), S.tdSayi);
    W.formul(7, r, `IF(${qGos}="","",${vf('Gerçekleşen', `Veri[Ay No],${an}`, `Veri[Metrik],${qGos}`, 'Veri[HF],1')})`, deger(x.f), S.tdSayi);
    W.formul(8, r, `IF(N(${H(6, r)})>0,${H(7, r)}/${H(6, r)},"")`, deger(x.oran), S.tdYz);
    W.formul(9, r, bantF(H(8, r)), bant(x.oran), S.tdOrta);
    W.formul(10, r, cf(`Veri[Ay No],${an}`, 'Veri[Değerlendirme],"Başarılı"'), x.B, S.tdSayi);
    W.formul(11, r, cf(`Veri[Ay No],${an}`, 'Veri[Değerlendirme],"Yetersiz"'), x.Y, S.tdSayi);
    W.formul(12, r, cf(`Veri[Ay No],${an}`, 'Veri[Değerlendirme],"Etkisi Yok"'), x.EY, S.tdSayi);
    W.formul(13, r, basF(H(10, r), H(11, r)), deger(basari(x.B, x.Y)), S.tdYz);
    W.formul(14, r, cf(`Veri[Ay No],${an}`), x.kayit, S.tdSayi);
    /* grafik serileri: veri olmayan ay 0 (çubuk görünmez, etiket biçimi sıfırı gizler); başarı oranı çizgisinde #N/A = boşluk */
    const bosAy = `OR(${H(14, r)}=0,${qGos}="")`;
    W.formul(RR, r, `IF(${bosAy},0,N(${H(6, r)}))`, x.kayit && gos ? x.h : 0, S.gizliMt);
    W.formul(SS, r, `IF(${bosAy},0,N(${H(7, r)}))`, x.kayit && gos ? x.f : 0, S.gizliMt);
    W.formul(TT, r, `IF(ISNUMBER(${H(13, r)}),${H(13, r)},NA())`, basari(x.B, x.Y) === '' ? { hata: '#N/A' } : basari(x.B, x.Y), S.gizliYz);
  });
  const ayT = { h: topla(ayHes, 'h'), f: topla(ayHes, 'f'), B: topla(ayHes, 'B'), Y: topla(ayHes, 'Y'), EY: topla(ayHes, 'EY'), kayit: topla(ayHes, 'kayit') };
  ad3(AY_T, 'Dönem toplamı', S.topMt);
  W.formul(6, AY_T, `IF(${qGos}="","",SUM(${ARALIK(6, AY0, 6, AY1)}))`, gos ? ayT.h : '', S.topSayi);
  W.formul(7, AY_T, `IF(${qGos}="","",SUM(${ARALIK(7, AY0, 7, AY1)}))`, gos ? ayT.f : '', S.topSayi);
  W.formul(8, AY_T, `IF(N(${H(6, AY_T)})>0,${H(7, AY_T)}/${H(6, AY_T)},"")`, gos && ayT.h > 0 ? ayT.f / ayT.h : '', S.topYz);
  W.formul(9, AY_T, bantF(H(8, AY_T)), bant(gos && ayT.h > 0 ? ayT.f / ayT.h : ''), S.topOrta);
  [10, 11, 12, 14].forEach(c => W.formul(c, AY_T, `SUM(${ARALIK(c, AY0, c, AY1)})`, c === 10 ? ayT.B : c === 11 ? ayT.Y : c === 12 ? ayT.EY : ayT.kayit, S.topSayi));
  W.formul(13, AY_T, basF(H(10, AY_T), H(11, AY_T)), deger(basari(ayT.B, ayT.Y)), S.topYz);
  W.yukseklik(AY_T, 24);
  kosulYuzde(`${H(8, AY0)}:${H(8, AY_T)}`, H(8, AY0)); kosulOndalik(`${H(6, AY0)}:${H(7, AY_T)}`, H(6, AY0));
  kosulSoluk(`${H(D0, AY0)}:${H(O0, AY1)}`, '$' + KOL(O0) + AY0);

  /* ---- 3) Danışman özeti ---- */
  bolumBaslik(DN_B, `"Danışman özeti — "&IF(${qGos}="","başarı sayıları",${qGos}&" hedef / gerçekleşen ve başarı sayıları")`,
    'Danışman özeti — ' + (gos ? gos + ' hedef / gerçekleşen ve başarı sayıları' : 'başarı sayıları'));
  W.metin(D0, DN_H, 'Danışman', S.thSol); W.birlestir(D0, DN_H, D0 + 2, DN_H, S.thSol);
  ['Çalışan', 'Hedef', 'Gerçekleşen', 'Gerçekleşme', 'Başarılı', 'Yetersiz', 'Etkisi Yok', 'Başarı oranı', 'Kayıt'].forEach((t, i) => W.metin(D0 + 3 + i, DN_H, t, S.th));
  W.yukseklik(DN_H, 34);
  danHes.forEach((x, i) => {
    const r = DN0 + i, dq = A(D0, r);
    ad3(r, x.d.ad, x.d.eski ? S.tdEski : S.td);
    W.metin(6, r, x.d.eski ? 'Eski' : 'Aktif', S.tdOrta);
    W.formul(7, r, `IF(${qGos}="","",${vf('Hedef', `Veri[Danışman],${dq}`, `Veri[Metrik],${qGos}`, 'Veri[HF],1')})`, deger(x.h), S.tdSayi);
    W.formul(8, r, `IF(${qGos}="","",${vf('Gerçekleşen', `Veri[Danışman],${dq}`, `Veri[Metrik],${qGos}`, 'Veri[HF],1')})`, deger(x.f), S.tdSayi);
    W.formul(9, r, `IF(N(${H(7, r)})>0,${H(8, r)}/${H(7, r)},"")`, deger(x.oran), S.tdYz);
    W.formul(10, r, cf(`Veri[Danışman],${dq}`, 'Veri[Değerlendirme],"Başarılı"'), x.B, S.tdSayi);
    W.formul(11, r, cf(`Veri[Danışman],${dq}`, 'Veri[Değerlendirme],"Yetersiz"'), x.Y, S.tdSayi);
    W.formul(12, r, cf(`Veri[Danışman],${dq}`, 'Veri[Değerlendirme],"Etkisi Yok"'), x.EY, S.tdSayi);
    W.formul(13, r, basF(H(10, r), H(11, r)), deger(basari(x.B, x.Y)), S.tdYz);
    W.formul(14, r, cf(`Veri[Danışman],${dq}`), x.kayit, S.tdSayi);
  });
  kosulYuzde(`${H(9, DN0)}:${H(9, DN1)}`, H(9, DN0)); kosulOndalik(`${H(7, DN0)}:${H(8, DN1)}`, H(7, DN0));
  kosulSoluk(`${H(D0, DN0)}:${H(O0, DN1)}`, '$' + KOL(O0) + DN0);

  /* ---- Notlar ---- */
  const notlar = [
    `Kaynak: Satış Panosu karne kayıtları — karnelere girilen hedef ve gerçekleşenler (${G.hazir} itibarıyla). Yalnız yetkili olduğunuz markanın verisi; ${KNOT}.`,
    `Dilimleyiciler (solda): Çeyrek, Ay, Danışman, Metrik, Çalışan Durumu${G.taslakVar ? ', Karne Durumu' : ''}. Birden çok seçim için Ctrl tuşuyla tıklayın; tümüne dönmek için dilimleyicinin sağ üstündeki filtre temizleme simgesine basın. Göstergeler, grafikler ve tablolar seçime göre anında değişir.`,
    'Gerçekleşme = Gerçekleşen ÷ Hedef (yalnız hedefi ve gerçekleşeni birlikte girilmiş karneler). Renk: ≥ %100 yeşil (hedefte) · %80–99 sarı (hedefe yakın) · < %80 kırmızı — Satış Panosu ile aynı bantlar.',
    '“Etkisi Yok” işaretli aylar hedef / gerçekleşen toplamlarına DAHİLDİR, Başarılı / Yetersiz sayımına katılmaz. Başarı oranı = Başarılı ÷ (Başarılı + Yetersiz).',
    'Aylık grafik ve tablolardaki hedef / gerçekleşen: tek bir sayısal metrik seçiliyse o metrik, değilse Araç Satış gösterilir (farklı birimler toplanmaz). NPS / CSI ve Gizli Müşteri müdür değerlendirmesidir; sayısal hedefi yoktur.',
    'Eski çalışanlar kırmızı ve “(Eski)” ekiyle gösterilir; “Çalışan Durumu” dilimleyicisiyle ayrılabilir. Elle girilen araç satış fiilileri Veri sayfasında “Not” sütununda işaretlidir.',
    'Veri sayfası panoyu besler; değerleri değiştirmeyin. Güncel veri için Satış Panosu › Karne › Performans › Excel Dashboard ile yeniden indirin.'
  ];
  W.metin(D0, NT_B, 'Nasıl okunur', S.notB); W.yukseklik(NT_B, 26);
  notlar.forEach((n, i) => { const r = NT_B + 1 + i; W.metin(D0, r, '• ' + n, S.not); W.birlestir(D0, r, O0, r, S.not); W.yukseklik(r, n.length > 150 ? 32 : 18); });
  const SON_SATIR = NT_B + notlar.length + 1;

  /* ======================= GRAFİKLER ======================= */
  const grafikler = [];
  const sReft = (c, r1, r2) => `Pano!${ARALIK(c, r1, c, r2)}`;
  /* A) Metrik bazında gerçekleşme — sütun; bant renkli üç seri (yalnız biri dolu, diğerleri 0) + şeffaf etiket serisi
     (değer yazısı tek yerden, sıfır gizli) + hedef çizgisi */
  if (nS) {
    const kat = metHes.slice(0, nS).map(x => x.m.kisa), katF = sReft(Q, MT0, MS1);
    const seri = (ad, renk, kol, kosul) => ({ ad, renk, katF, kat, degF: sReft(kol, MT0, MS1), deg: metHes.slice(0, nS).map(x => x.oran !== '' && kosul(x.oran) ? x.oran : 0) });
    const cub = [seri('Hedefte (≥ %100)', R.OK, RR, o => o >= 1), seri('Hedefe yakın (%80–99)', R.UY, SS, o => o >= 0.8 && o < 1), seri('Hedefin altında (< %80)', R.HT, TT, o => o < 0.8)];
    cub.push({ ad: 'Gerçekleşme', renk: null, katF, kat, degF: sReft(WW, MT0, MS1), deg: metHes.slice(0, nS).map(x => x.oran !== '' ? x.oran : 0), etiket: { fmt: '%0;;;', konum: 'outEnd', sz: 900, renk: R.MET1 } });
    const ciz = { ad: 'Hedef (%100)', renk: R.PIR7, kesikli: true, kalinlik: 19050, katF, kat, degF: sReft(UU, MT0, MS1), deg: kat.map(() => 1) };
    grafikler.push(grafikKabuk(
      `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${cub.map((s, i) => cubukSeri(i, s)).join('')}${etiketYok}<c:gapWidth val="55"/><c:overlap val="100"/><c:axId val="510001"/><c:axId val="510002"/></c:barChart>`
      + `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${cizgiSeri(4, ciz)}${etiketYok}<c:marker val="1"/><c:axId val="510001"/><c:axId val="510002"/></c:lineChart>`
      + eksenler({ katId: 510001, degId: 510002, katPos: 'b', degPos: 'l', min: 0, fmt: YZ }), true, [3]));
  } else grafikler.push(null);
  /* B) Metrik bazında değerlendirme — yatay %100 yığılmış (Başarılı / Yetersiz / Etkisi Yok adetleri) */
  {
    const kat = metHes.map(x => x.m.kisa), katF = sReft(Q, MT0, MT1);
    const seri = (ad, renk, kol, alan, yaziRenk) => ({ ad, renk, katF, kat, degF: sReft(kol, MT0, MT1), deg: metHes.map(x => x[alan]), etiket: { fmt: '0;;;', konum: 'ctr', sz: 900, renk: yaziRenk } });
    const s3 = [seri('Başarılı', R.OK, 10, 'B', R.BEYAZ), seri('Yetersiz', R.HT, 11, 'Y', R.BEYAZ), seri('Etkisi Yok', R.EY, 12, 'EY', R.MET1)];
    grafikler.push(grafikKabuk(
      `<c:barChart><c:barDir val="bar"/><c:grouping val="percentStacked"/><c:varyColors val="0"/>${s3.map((s, i) => cubukSeri(i, s)).join('')}${etiketYok}<c:gapWidth val="45"/><c:overlap val="100"/><c:axId val="520001"/><c:axId val="520002"/></c:barChart>`
      + eksenler({ katId: 520001, degId: 520002, katPos: 'l', degPos: 'b', ters: true, min: 0, max: 1, fmt: YZ }), true));
  }
  /* C) Aylık hedef – gerçekleşen (gösterilen metrik) */
  {
    const kat = ayHes.map(x => AY_KISA[x.a]), katF = sReft(Q, AY0, AY1);
    const s2 = [
      { ad: 'Hedef', renk: R.LAC2, katF, kat, degF: sReft(RR, AY0, AY1), deg: ayHes.map(x => x.kayit && gos ? x.h : 0), etiket: { fmt: '[=0]"";#,##0', konum: 'outEnd', sz: 800, renk: R.MET3, b: false } },
      { ad: 'Gerçekleşen', renk: R.LAC7, katF, kat, degF: sReft(SS, AY0, AY1), deg: ayHes.map(x => x.kayit && gos ? x.f : 0), etiket: { fmt: '[=0]"";#,##0', konum: 'outEnd', sz: 900, renk: R.LAC9 } }];
    grafikler.push(grafikKabuk(
      `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${s2.map((s, i) => cubukSeri(i, s)).join('')}${etiketYok}<c:gapWidth val="70"/><c:overlap val="-8"/><c:axId val="530001"/><c:axId val="530002"/></c:barChart>`
      + eksenler({ katId: 530001, degId: 530002, katPos: 'b', degPos: 'l', min: 0 }), true));
  }
  /* D) Aylık başarı oranı */
  {
    const kat = ayHes.map(x => AY_KISA[x.a]), katF = sReft(Q, AY0, AY1);
    const s = { ad: 'Başarı oranı', renk: R.LAC7, isaret: R.PIR, kalinlik: 28575, katF, kat, degF: sReft(TT, AY0, AY1), deg: ayHes.map(x => basari(x.B, x.Y) === '' ? null : basari(x.B, x.Y)), etiket: { fmt: YZ, konum: 't', sz: 900, renk: R.LAC9 } };
    grafikler.push(grafikKabuk(
      `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${cizgiSeri(0, s)}${etiketYok}<c:marker val="1"/><c:axId val="540001"/><c:axId val="540002"/></c:lineChart>`
      + eksenler({ katId: 540001, degId: 540002, katPos: 'b', degPos: 'l', min: 0, max: 1.15, fmt: YZ }), false));
  }

  /* ======================= DİLİMLEYİCİLER ======================= */
  const DILIM = [
    { kol: 'Çeyrek', ad: 'Ceyrek', baslik: 'Çeyrek', sutun: 2, oge: [...new Set(V.map(v => v.ceyrek))].length },
    { kol: 'Ay', ad: 'Ay', baslik: 'Ay', sutun: 2, oge: [...new Set(V.map(v => v.ay))].length },
    { kol: 'Danışman', ad: 'Danisman', baslik: 'Danışman', sutun: 1, oge: DANS.length, enCok: 9 },
    { kol: 'Metrik', ad: 'Metrik', baslik: 'Metrik', sutun: 1, oge: METS.length, enCok: 11 },
    { kol: 'Çalışan Durumu', ad: 'Calisan', baslik: 'Çalışan Durumu', sutun: 2, oge: 2 }
  ];
  if (G.taslakVar) DILIM.push({ kol: 'Karne Durumu', ad: 'Karne', baslik: 'Karne Durumu', sutun: 2, oge: 2 });
  let yDil = sy(6);
  const dilimCizim = DILIM.map((d, i) => {
    const satir = Math.ceil(d.oge / d.sutun), gorunur = Math.min(satir, d.enCok || satir), yuk = 38 + gorunur * 28 + 6;
    const px = { x: sx(1) + 4, y: yDil, w: SUT[1] - 8, h: yuk }; yDil += yuk + 10;
    const bas = pxSatir(px.y), son = pxSatir(px.y + px.h), sb = pxSutun(px.x), ss = pxSutun(px.x + px.w);
    return dilimCapa(2 + i, d.ad, { c: sb.c, cOff: sb.off, r: bas.r, rOff: bas.off }, { c: ss.c, cOff: ss.off, r: son.r, rOff: son.off }, px);
  });
  const grafikYer = [[D0, 12, D0 + 5, 27, 'sol'], [D0 + 6, 12, O0, 27, 'sag'], [D0, 31, D0 + 5, 46, 'sol'], [D0 + 6, 31, O0, 46, 'sag']];
  const grafikCizim = [], grafikDosya = [];
  let gNo = 0;
  grafikler.forEach((g, i) => {
    if (!g) return;
    gNo++; grafikDosya.push([`xl/charts/chart${gNo}.xml`, g]);
    const [c1, r1, c2, r2, yan] = grafikYer[i];
    const bas = { c: c1, cOff: yan === 'sag' ? 6 : 0, r: r1 - 1, rOff: 0 }, son = { c: c2, cOff: SUT[c2] - (yan === 'sol' ? 6 : 0), r: r2 - 1, rOff: satirPx(r2) };
    grafikCizim.push(grafikCapa(20 + i, 'rId' + gNo, 'Grafik ' + (i + 1), bas, son));
  });
  const cizimXml = XML_BAS + `<xdr:wsDr xmlns:xdr="${NS.xdr}" xmlns:a="${NS.a}">${dilimCizim.join('')}${grafikCizim.join('')}</xdr:wsDr>`;
  const cizimRel = iliskiler(grafikDosya.map((g, i) => ['rId' + (i + 1), RT.grafik, `../charts/chart${i + 1}.xml`]));
  const dilimXml = XML_BAS + `<slicers xmlns="${NS.x14}" xmlns:mc="${NS.mc}" mc:Ignorable="x" xmlns:x="${NS.main}">`
    + DILIM.map(d => `<slicer name="${X(d.ad)}" cache="Dilimleyici_${d.ad}" caption="${X(d.baslik)}"${d.sutun > 1 ? ` columnCount="${d.sutun}"` : ''} style="SlicerStyleLight1" rowHeight="228600"/>`).join('') + '</slicers>';
  const onbellekXml = d => XML_BAS + `<slicerCacheDefinition xmlns="${NS.x14}" xmlns:mc="${NS.mc}" mc:Ignorable="x" xmlns:x="${NS.main}" name="Dilimleyici_${d.ad}" sourceName="${X(d.kol)}">`
    + `<extLst><x:ext uri="{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}" xmlns:x15="${NS.x15}"><x15:tableSlicerCache tableId="1" column="${KI[d.kol] + 1}"/></x:ext></extLst></slicerCacheDefinition>`;

  /* ======================= PANO SAYFA XML ======================= */
  const kolonlar = SUT.map((px, i) => `<col min="${i + 1}" max="${i + 1}" width="${genislik(px)}" customWidth="1"/>`).join('')
    + `<col min="${Q + 1}" max="${WW + 1}" width="14" hidden="1" customWidth="1"/>`;
  const panoXml = XML_BAS + `<worksheet xmlns="${NS.main}" xmlns:r="${NS.r}" xmlns:mc="${NS.mc}" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">`
    + `<sheetPr><tabColor rgb="FF${R.LAC9}"/><pageSetUpPr fitToPage="1"/></sheetPr><dimension ref="A1:${KOL(WW)}${SON_SATIR}"/>`
    + '<sheetViews><sheetView showGridLines="0" tabSelected="1" zoomScale="90" zoomScaleNormal="90" workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews>'
    + `<sheetFormatPr defaultRowHeight="15"/><cols>${kolonlar}</cols><sheetData>${W.satirlarXml()}</sheetData>${W.birlesikXml()}${W.kosulXml()}`
    + '<pageMargins left="0.3" right="0.3" top="0.4" bottom="0.4" header="0.2" footer="0.2"/><pageSetup paperSize="9" orientation="landscape" fitToHeight="0"/>'
    + '<drawing r:id="rId1"/>'
    + `<extLst><ext uri="{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}" xmlns:x15="${NS.x15}"><x14:slicerList xmlns:x14="${NS.x14}"><x14:slicer r:id="rId2"/></x14:slicerList></ext></extLst></worksheet>`;

  /* ======================= KİTAP ======================= */
  const dilimRid = DILIM.map((d, i) => 'rId' + (10 + i));
  const kitapXml = XML_BAS + `<workbook xmlns="${NS.main}" xmlns:r="${NS.r}" xmlns:mc="${NS.mc}" mc:Ignorable="x15" xmlns:x15="${NS.x15}">`
    + '<fileVersion appName="xl" lastEdited="6" lowestEdited="6" rupBuild="14420"/><workbookPr defaultThemeVersion="164011"/>'
    + '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="15600" activeTab="0"/></bookViews>'
    + `<sheets><sheet name="Pano" sheetId="1" r:id="rId1"/><sheet name="Veri" sheetId="2" r:id="rId2"/></sheets>`
    + '<definedNames>' + DILIM.map(d => `<definedName name="Dilimleyici_${d.ad}">#N/A</definedName>`).join('')
    + `<definedName name="_xlnm.Print_Area" localSheetId="0">Pano!$A$1:$${KOL(15)}$${SON_SATIR}</definedName></definedNames>`
    + '<calcPr calcId="191029" fullCalcOnLoad="1"/>'
    + `<extLst><ext uri="{46BE6895-7355-4a93-B00E-2C351335B9C9}" xmlns:x15="${NS.x15}"><x15:slicerCaches xmlns:x14="${NS.x14}">${dilimRid.map(id => `<x14:slicerCache r:id="${id}"/>`).join('')}</x15:slicerCaches></ext></extLst></workbook>`;
  const kitapRel = iliskiler([['rId1', RT.sayfa, 'worksheets/sheet1.xml'], ['rId2', RT.sayfa, 'worksheets/sheet2.xml'], ['rId3', RT.stil, 'styles.xml'], ['rId4', RT.tema, 'theme/theme1.xml'], ['rId5', RT.sst, 'sharedStrings.xml'],
    ...DILIM.map((d, i) => [dilimRid[i], RT.dilimOnbellek, `slicerCaches/slicerCache${i + 1}.xml`])]);
  const turler = XML_BAS + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
    + [['/xl/workbook.xml', CT.kitap], ['/xl/worksheets/sheet1.xml', CT.sayfa], ['/xl/worksheets/sheet2.xml', CT.sayfa], ['/xl/theme/theme1.xml', CT.tema], ['/xl/styles.xml', CT.stil], ['/xl/sharedStrings.xml', CT.sst],
      ['/xl/tables/table1.xml', CT.tablo], ['/xl/drawings/drawing1.xml', CT.cizim], ['/xl/slicers/slicer1.xml', CT.dilim], ...DILIM.map((d, i) => [`/xl/slicerCaches/slicerCache${i + 1}.xml`, CT.dilimOnbellek]),
      ...grafikDosya.map((g, i) => [`/xl/charts/chart${i + 1}.xml`, CT.grafik]), ['/docProps/core.xml', CT.core], ['/docProps/app.xml', CT.app]]
      .map(([p, t]) => `<Override PartName="${p}" ContentType="${t}"/>`).join('') + '</Types>';
  const simdi = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  const core = XML_BAS + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
    + `<dc:title>${X(G.marka + ' Performans Metrik Panosu ' + G.yil)}</dc:title><dc:creator>İnciroğlu Otomotiv · Satış Panosu</dc:creator><cp:lastModifiedBy>Satış Panosu</cp:lastModifiedBy>`
    + `<dcterms:created xsi:type="dcterms:W3CDTF">${simdi}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${simdi}</dcterms:modified></cp:coreProperties>`;
  const app = XML_BAS + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>'
    + '<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Çalışma Sayfaları</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4></vt:variant></vt:vector></HeadingPairs>'
    + '<TitlesOfParts><vt:vector size="2" baseType="lpstr"><vt:lpstr>Pano</vt:lpstr><vt:lpstr>Veri</vt:lpstr></vt:vector></TitlesOfParts><Company>İnciroğlu Otomotiv</Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion></Properties>';

  const dosyalar = [
    ['[Content_Types].xml', turler],
    ['_rels/.rels', iliskiler([['rId1', RT.belge, 'xl/workbook.xml'], ['rId2', RT.core, 'docProps/core.xml'], ['rId3', RT.app, 'docProps/app.xml']])],
    ['docProps/core.xml', core], ['docProps/app.xml', app],
    ['xl/workbook.xml', kitapXml], ['xl/_rels/workbook.xml.rels', kitapRel],
    ['xl/styles.xml', stil.xml()], ['xl/theme/theme1.xml', temaXml()],
    ['xl/worksheets/sheet1.xml', panoXml], ['xl/worksheets/_rels/sheet1.xml.rels', iliskiler([['rId1', RT.cizim, '../drawings/drawing1.xml'], ['rId2', RT.dilim, '../slicers/slicer1.xml']])],
    ['xl/worksheets/sheet2.xml', veriSayfaXml], ['xl/worksheets/_rels/sheet2.xml.rels', iliskiler([['rId1', RT.tablo, '../tables/table1.xml']])],
    ['xl/tables/table1.xml', tabloXml],
    ['xl/drawings/drawing1.xml', cizimXml], ['xl/drawings/_rels/drawing1.xml.rels', cizimRel],
    ...grafikDosya,
    ['xl/slicers/slicer1.xml', dilimXml],
    ...DILIM.map((d, i) => [`xl/slicerCaches/slicerCache${i + 1}.xml`, onbellekXml(d)]),
    ['xl/sharedStrings.xml', sst.xml()]
  ];
  const blob = await zipYap(dosyalar);
  /* Test/doğrulama için açılıştaki (önbellek) değerler */
  const ozet = { kayit: V.length, gorunur: GV.length, gos, tekMetrik, metrik: metHes.map(x => ({ key: x.m.key, kisa: x.m.kisa, h: x.h, f: x.f, oran: x.oran, B: x.B, Y: x.Y, EY: x.EY, kayit: x.kayit })),
    ay: ayHes.map(x => ({ a: x.a, h: x.h, f: x.f, B: x.B, Y: x.Y, EY: x.EY, kayit: x.kayit })), dan: danHes.map(x => ({ ad: x.d.ad, h: x.h, f: x.f, B: x.B, Y: x.Y, EY: x.EY, kayit: x.kayit })),
    ekranDustu, karneDustu, satir: { MT0, MT1, MT_T, AY0, AY1, AY_T, DN0, DN1, NT_B }, dilim: DILIM.map(d => d.baslik), secim: Object.fromEntries(Object.entries(secim).map(([k, s]) => [k, [...s]])) };
  return { blob, ozet };
}
