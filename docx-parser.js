/*
 * docx-parser.js — extraction des tableaux d'un document Word (.docx).
 * Seule dépendance : JSZip (décompression du .docx). Hors ligne, sans serveur.
 *
 * Principe : aucune ligne n'est écartée silencieusement. Chaque ligne non vide
 * est classée comme en-tête, libellé Direction/Bureau, ligne sans donnée ou
 * ligne importée. Les cellules fusionnées (horizontalement et verticalement)
 * sont reproduites dans chaque case concernée.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.DocxParser = factory();
  }
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  var FIELDS = ['direction', 'bureau', 'designation', 'code', 'qte', 'etat', 'obs'];
  var PATTERNS = {
    direction: /^direction$/i,
    bureau: /^bureau$/i,
    designation: /d[ée]sign/i,
    code: /\bcode\b|\bn°\s*(d['’]\s*)?inv/i,
    qte: /^(qt[ée]|quantit)/i,
    etat: /^[ée]tat/i,
    obs: /observ|remarque/i
  };
  var DEFAULT_ORDER = ['code', 'designation', 'qte', 'etat', 'obs'];
  // Cellule qui sert de libellé : « Direction … », « Bureau : … », « Bureau » seul.
  var LABEL_CELL = /^(direction\b|bureau\b\s*(n°\s*\d*)?\s*[:\-–]|bureau\s*(n°\s*\d*)?\s*:?$)/i;

  function norm(s) {
    return String(s == null ? '' : s).replace(/\s+/g, ' ').trim();
  }

  var seq = 0;
  function uid() {
    seq += 1;
    return 'r' + Date.now().toString(36) + seq.toString(36) + Math.random().toString(36).slice(2, 6);
  }

  function attr(el, name) {
    return el.getAttributeNS(W, name) || el.getAttribute('w:' + name) || '';
  }

  function kids(el, local) {
    var out = [];
    for (var n = el.firstChild; n; n = n.nextSibling) {
      if (n.nodeType === 1 && n.localName === local) out.push(n);
    }
    return out;
  }

  // Texte d'un nœud Word : w:t, tabulations, sauts de ligne ; ignore les suppressions (w:delText).
  function textOf(node) {
    var out = '';
    (function walk(n) {
      for (var c = n.firstChild; c; c = c.nextSibling) {
        if (c.nodeType !== 1) continue;
        switch (c.localName) {
          case 't': out += c.textContent; break;
          case 'tab': out += '\t'; break;
          case 'br': case 'cr': out += ' '; break;
          case 'delText': case 'instrText':
          case 'pPr': case 'tblPr': case 'trPr': case 'tcPr': case 'sectPr':
            break;
          default: walk(c);
        }
      }
    })(node);
    return out;
  }

  // Texte d'un tableau imbriqué, rendu lisible (cellules séparées par « | »).
  function nestedTableText(tbl) {
    return kids(tbl, 'tr').map(function (tr) {
      return kids(tr, 'tc').map(function (tc) { return norm(textOf(tc)); }).join(' | ');
    }).join(' ; ');
  }

  function cellText(tc, warnings) {
    var parts = [];
    kids(tc, 'p').forEach(function (p) {
      var t = norm(textOf(p));
      if (t) parts.push(t);
    });
    kids(tc, 'tbl').forEach(function (nt) {
      var t = norm(nestedTableText(nt));
      if (t) parts.push(t);
      warnings.push('Tableau imbriqué dans une cellule : son contenu a été repris.');
    });
    return parts.join(' ');
  }

  // Grille logique : chaque ligne a une case par colonne, fusions comprises.
  function buildGrid(tbl, warnings) {
    var rows = [];
    var vprev = []; // texte de la cellule ouverte par une fusion verticale, par colonne
    kids(tbl, 'tr').forEach(function (tr) {
      var row = [];
      var col = 0;
      kids(tr, 'tc').forEach(function (tc) {
        var pr = kids(tc, 'tcPr')[0];
        var span = 1;
        var vm = '';
        if (pr) {
          var gs = kids(pr, 'gridSpan')[0];
          if (gs) span = parseInt(attr(gs, 'val'), 10) || 1;
          var vmEl = kids(pr, 'vMerge')[0];
          if (vmEl) vm = attr(vmEl, 'val') === 'restart' ? 'restart' : 'continue';
        }
        var own = cellText(tc, warnings);
        for (var s = 0; s < span; s++) {
          var c = col + s;
          if (vm === 'continue') {
            row[c] = { text: vprev[c] || '', merged: true };
          } else {
            row[c] = { text: own, merged: span > 1 };
          }
          if (vm === 'restart') vprev[c] = own;
          else if (vm !== 'continue') vprev[c] = null;
        }
        col += span;
      });
      rows.push(row);
    });
    var width = 0;
    rows.forEach(function (r) { width = Math.max(width, r.length); });
    rows.forEach(function (r) {
      for (var i = 0; i < width; i++) {
        if (!r[i]) r[i] = { text: '', merged: false };
      }
    });
    return rows;
  }

  // Ligne-titre fusionnée sur toute la largeur (une seule valeur, répétée dans chaque colonne couverte)
  // ex. « Mobilier — bureau 3 » : c'est le titre d'une section, pas une ligne d'inventaire.
  function isSpanTitle(row) {
    if (!row.length || !row[0].text) return false;
    return row.every(function (c) { return c.merged && c.text === row[0].text; });
  }

  function matchField(text, used) {
    if (text.length > 60) return null;
    for (var k = 0; k < FIELDS.length; k++) {
      var f = FIELDS[k];
      if (!used[f] && PATTERNS[f].test(text)) return f;
    }
    return null;
  }

  // Une ligne d'en-tête contient au moins deux intitulés reconnus, dont Désignation ou Code.
  function detectHeader(texts) {
    var map = {};
    var used = {};
    var score = 0;
    texts.forEach(function (t, i) {
      var f = t ? matchField(t, used) : null;
      if (f && map[f] === undefined) {
        map[f] = i;
        used[f] = true;
        score += 1;
      }
    });
    return (score >= 2 && (map.designation !== undefined || map.code !== undefined)) ? map : null;
  }

  function defaultMap(width) {
    var map = {};
    DEFAULT_ORDER.forEach(function (f, i) {
      if (i < width) map[f] = i;
    });
    return map;
  }

  function hasData(texts, map) {
    return ['code', 'qte', 'etat'].some(function (f) {
      return map[f] !== undefined && texts[map[f]];
    });
  }

  // Libellé d'une ligne : « Direction : X », « Bureau : Y », ou les deux sur une même ligne.
  function parseLabelLine(text) {
    var t = norm(text);
    var lab = {};
    var split = t.search(/\bbureau\b/i);
    if (/^direction\b/i.test(t)) {
      lab.hasDirection = true;
      lab.direction = norm((split > 0 ? t.slice(0, split) : t)
        .replace(/^direction\b\s*(n°\s*\d*)?\s*[:\-–]?/i, ''));
    }
    if (/^bureau\b/i.test(t)) {
      lab.hasBureau = true;
      lab.bureau = norm(t.replace(/^bureau\b\s*(n°\s*\d*)?\s*[:\-–]?/i, ''));
    } else if (split > 0 && lab.hasDirection) {
      lab.hasBureau = true;
      lab.bureau = norm(t.slice(split).replace(/^bureau\b\s*(n°\s*\d*)?\s*[:\-–]?/i, ''));
    }
    return lab;
  }

  function applyLabels(ctx, lab, warnings) {
    if (lab.hasDirection) {
      if (lab.direction) {
        if (lab.direction !== ctx.direction) ctx.bureau = '';
        ctx.direction = lab.direction;
      } else {
        warnings.push('Libellé « Direction » sans valeur détectée.');
      }
    }
    if (lab.hasBureau) {
      if (lab.bureau) ctx.bureau = lab.bureau;
      else warnings.push('Libellé « Bureau » sans valeur détectée.');
    }
  }

  function nextValue(texts, i) {
    for (var j = i + 1; j < texts.length; j++) {
      var t = texts[j];
      if (t && !LABEL_CELL.test(t)) return t;
    }
    return '';
  }

  // Libellés présents dans une ligne (cellule unique ou paire « Libellé | valeur »).
  function labelsFromRow(texts) {
    var labs = [];
    texts.forEach(function (t, i) {
      if (!t || !LABEL_CELL.test(t)) return;
      var lab = parseLabelLine(t);
      if (lab.hasDirection && !lab.direction) lab.direction = nextValue(texts, i);
      if (lab.hasBureau && !lab.bureau) lab.bureau = nextValue(texts, i);
      labs.push(lab);
    });
    return labs;
  }

  function parseTable(tbl, ctx, docName, index, warnings) {
    var tblWarnings = [];
    var rows = buildGrid(tbl, tblWarnings);
    var map = null;
    var headerFound = false;
    var items = [];
    var counts = { nonBlank: 0, headers: 0, labels: 0, titles: 0, imported: 0, noData: 0, blank: 0, mergedRows: 0 };
    var sectionTitle = '';

    rows.forEach(function (row, r) {
      var texts = row.map(function (c) { return c.text; });
      if (row.some(function (c) { return c.merged; })) counts.mergedRows += 1;
      if (!texts.some(Boolean)) { counts.blank += 1; return; }
      counts.nonBlank += 1;

      var header = detectHeader(texts);
      if (header) {
        map = header;
        headerFound = true;
        counts.headers += 1;
        return;
      }

      var labs = labelsFromRow(texts);
      if (labs.length && (!map || !hasData(texts, map))) {
        labs.forEach(function (lab) { applyLabels(ctx, lab, tblWarnings); });
        counts.labels += 1;
        return;
      }

      if (isSpanTitle(row)) {
        sectionTitle = texts[0];
        counts.titles += 1;
        return;
      }

      if (!map) map = defaultMap(texts.length);
      var get = function (f) { return map[f] !== undefined ? (texts[map[f]] || '') : ''; };
      var item = {
        id: uid(),
        direction: get('direction') || ctx.direction,
        bureau: get('bureau') || ctx.bureau,
        designation: get('designation'),
        code: get('code'),
        qte: get('qte'),
        etat: get('etat'),
        obs: get('obs'),
        flags: [],
        origin: { doc: docName, table: index, row: r + 1 }
      };
      if (!item.designation && !item.code && !item.qte && !item.etat && !item.obs) {
        counts.noData += 1;
        return;
      }
      items.push(item);
      counts.imported += 1;
    });

    if (!headerFound) {
      tblWarnings.push('En-tête non détecté : colonnes supposées Code, Désignation, Qté, État, Obs. À vérifier.');
    }
    warnings.push.apply(warnings, tblWarnings);

    return {
      id: uid(),
      title: sectionTitle || ctx.title || ('Tableau ' + index),
      docName: docName,
      index: index,
      columns: map ? FIELDS.filter(function (f) { return map[f] !== undefined; }) : [],
      items: items,
      meta: {
        headerFound: headerFound,
        physicalRows: rows.length,
        counts: counts,
        warnings: unique(tblWarnings)
      }
    };
  }

  function unique(arr) {
    var seen = {};
    return arr.filter(function (x) {
      if (seen[x]) return false;
      seen[x] = true;
      return true;
    });
  }

  // Parcourt le corps du document dans l'ordre, en entrant dans les blocs w:sdt.
  function walkBody(container, visit) {
    for (var n = container.firstChild; n; n = n.nextSibling) {
      if (n.nodeType !== 1) continue;
      if (n.localName === 'sdt') {
        kids(n, 'sdtContent').forEach(function (c) { walkBody(c, visit); });
        continue;
      }
      visit(n);
    }
  }

  function parseBody(doc, docName) {
    var body = doc.getElementsByTagNameNS(W, 'body')[0];
    if (!body) throw new Error('Structure du document Word non reconnue.');
    var ctx = { direction: '', bureau: '', title: '' };
    var warnings = [];
    var tables = [];
    var paragraphs = 0;

    walkBody(body, function (n) {
      if (n.localName === 'p') {
        var t = norm(textOf(n));
        if (!t) return;
        paragraphs += 1;
        if (/^(direction|bureau)\b/i.test(t) && t.length <= 160) {
          applyLabels(ctx, parseLabelLine(t), warnings);
        } else if (t.length <= 160) {
          ctx.title = t;
        }
      } else if (n.localName === 'tbl') {
        tables.push(parseTable(n, ctx, docName, tables.length + 1, warnings));
      }
    });

    if (!tables.length) warnings.push('Aucun tableau détecté dans ce document.');
    return {
      docName: docName,
      paragraphs: paragraphs,
      tables: tables,
      warnings: unique(warnings)
    };
  }

  /*
   * parseDocx(entrée, options) → Promise<{ docName, tables, warnings }>
   * entrée : ArrayBuffer ou Uint8Array du fichier .docx
   * options : { docName, JSZip, DOMParser } (les deux derniers sont détectés en navigateur)
   */
  function parseDocx(input, options) {
    options = options || {};
    var JSZipLib = options.JSZip || (typeof JSZip !== 'undefined' ? JSZip : null);
    var DP = options.DOMParser || (typeof DOMParser !== 'undefined' ? DOMParser : null);
    if (!JSZipLib || !DP) {
      return Promise.reject(new Error('Bibliothèques de lecture indisponibles.'));
    }
    var docName = options.docName || 'document.docx';
    return JSZipLib.loadAsync(input).then(function (zip) {
      var f = zip.file('word/document.xml');
      if (!f) throw new Error('Ce fichier n’est pas un document Word (.docx) valide.');
      return f.async('string');
    }).then(function (xml) {
      var doc = new DP().parseFromString(xml, 'application/xml');
      if (doc.getElementsByTagName('parsererror').length) {
        throw new Error('Le contenu du document « ' + docName + ' » est illisible.');
      }
      return parseBody(doc, docName);
    });
  }

  return {
    parseDocx: parseDocx,
    parseLabelLine: parseLabelLine,
    detectHeader: detectHeader,
    FIELDS: FIELDS
  };
}));
