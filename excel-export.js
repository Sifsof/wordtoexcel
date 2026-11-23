/*
 * excel-export.js — génération du classeur .xlsx (SheetJS « xlsx-js-style » + JSZip).
 * Structure imposée : Direction | Bureau | Désignation | Code | Qte | État.
 * Codes gardés en texte (zéros initiaux conservés), quantités en nombres.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.ExcelExport = factory();
  }
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var HEADERS = ['Direction', 'Bureau', 'Désignation', 'Code', 'Qte', 'État'];
  var SHEET_NAME = 'Inventaire';
  var GOLD = 'D4AF37';
  var EDGE = { style: 'thin', color: { rgb: '7A6A2E' } };
  var BORDER = { top: EDGE, bottom: EDGE, left: EDGE, right: EDGE };
  var HEAD_STYLE = {
    font: { bold: true, color: { rgb: '0D0D0D' }, name: 'Calibri', sz: 11 },
    fill: { patternType: 'solid', fgColor: { rgb: GOLD } },
    alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
    border: BORDER
  };

  // Quantité : nombre si elle en est un (« 05 » → 5), sinon texte d'origine.
  function qteValue(raw) {
    var s = String(raw == null ? '' : raw).replace(/\s/g, '');
    if (s === '') return '';
    if (/^-?\d+([.,]\d+)?$/.test(s)) return Number(s.replace(',', '.'));
    return String(raw).trim();
  }

  // Une ligne sans aucune donnée (ligne ajoutée puis laissée vide) n'est pas exportée.
  function isBlank(it) {
    return !it.designation && !it.code && !it.qte && !it.etat && !it.obs;
  }

  function blankCount(project) {
    var n = 0;
    (project.tables || []).forEach(function (t) {
      (t.items || []).forEach(function (it) { if (isBlank(it)) n += 1; });
    });
    return n;
  }

  function rowsOf(project) {
    var out = [];
    (project.tables || []).forEach(function (t) {
      (t.items || []).forEach(function (it) {
        if (isBlank(it)) return;
        out.push([
          String(it.direction || ''),
          String(it.bureau || ''),
          String(it.designation || ''),
          String(it.code || ''),
          qteValue(it.qte),
          String(it.etat || '')
        ]);
      });
    });
    return out;
  }

  function colLetter(i) { return String.fromCharCode(65 + i); }

  function styleFor(value) {
    var numeric = typeof value === 'number';
    return {
      font: { name: 'Calibri', sz: 10, color: { rgb: '000000' } },
      alignment: {
        vertical: 'top',
        wrapText: true,
        horizontal: numeric ? 'right' : 'left'
      },
      border: BORDER
    };
  }

  function buildSheet(XLSX, rows) {
    var aoa = [HEADERS.slice()].concat(rows);
    var ws = XLSX.utils.aoa_to_sheet(aoa);
    var lastRow = aoa.length;

    // En-têtes
    for (var c = 0; c < HEADERS.length; c++) {
      var h = ws[colLetter(c) + '1'];
      if (h) h.s = HEAD_STYLE;
    }
    // Corps
    for (var r = 1; r < lastRow; r++) {
      for (var k = 0; k < HEADERS.length; k++) {
        var cell = ws[colLetter(k) + (r + 1)];
        if (!cell) continue;
        if (k === 3 && cell.t !== 's') { cell.t = 's'; cell.v = String(cell.v); }
        cell.s = styleFor(cell.v);
      }
    }
    // Largeurs adaptées au contenu
    ws['!cols'] = HEADERS.map(function (label, k) {
      var max = label.length;
      for (var i = 0; i < rows.length; i++) {
        var len = String(rows[i][k]).length;
        if (len > max) max = len;
      }
      var cap = k === 2 ? 60 : 30;
      return { wch: Math.min(cap, Math.max(10, max + 2)) };
    });
    // Filtres automatiques sur les six colonnes
    ws['!autofilter'] = { ref: 'A1:' + colLetter(HEADERS.length - 1) + lastRow };
    return ws;
  }

  // Excel attend la première ligne figée : ajoutée directement dans le XML de la feuille.
  function freezeFirstRow(JSZipLib, bytes) {
    return JSZipLib.loadAsync(bytes).then(function (zip) {
      var path = 'xl/worksheets/sheet1.xml';
      var file = zip.file(path);
      if (!file) return bytes;
      return file.async('string').then(function (xml) {
        var pane = '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>' +
          '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>';
        xml = xml.replace(/<sheetView([^>]*)\/>/, '<sheetView$1>' + pane + '</sheetView>');
        xml = xml.replace(/<sheetView([^>]*)>(?!<pane)/, '<sheetView$1>' + pane);
        zip.file(path, xml);
        return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
      });
    });
  }

  /*
   * buildXlsx(project, { XLSX, JSZip }) → Promise<Uint8Array>
   */
  function buildXlsx(project, deps) {
    var XLSX = deps && deps.XLSX ? deps.XLSX : (typeof window !== 'undefined' ? window.XLSX : null);
    var JSZipLib = deps && deps.JSZip ? deps.JSZip : (typeof JSZip !== 'undefined' ? JSZip : null);
    if (!XLSX || !JSZipLib) return Promise.reject(new Error('Bibliothèques Excel indisponibles.'));
    var rows = rowsOf(project);
    if (!rows.length) return Promise.reject(new Error('Aucune ligne à convertir.'));

    var wb = XLSX.utils.book_new();
    wb.Props = { Title: project.name || 'Inventaire' };
    XLSX.utils.book_append_sheet(wb, buildSheet(XLSX, rows), SHEET_NAME);
    var bytes = XLSX.write(wb, { bookType: 'xlsx', type: 'array', cellStyles: true });
    return freezeFirstRow(JSZipLib, new Uint8Array(bytes));
  }

  return {
    HEADERS: HEADERS,
    qteValue: qteValue,
    rowsOf: rowsOf,
    blankCount: blankCount,
    isBlank: isBlank,
    buildXlsx: buildXlsx
  };
}));
