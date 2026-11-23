/*
 * app.js — interface, stockage IndexedDB, annulation/rétablissement, archives JSON.
 * Tout fonctionne hors ligne ; aucune donnée ne quitte l'ordinateur.
 */
(function () {
  'use strict';

  var M = window.ProjectModel;
  var P = window.DocxParser;
  var X = window.ExcelExport;

  var CFG = window.CWE_CONFIG || { dbName: 'convertisseur-word-excel', dbVersion: 1, version: '' };
  var DB_NAME = CFG.dbName;
  var DB_VERSION = CFG.dbVersion;
  var MAX_HISTORY = 100;
  var FIELD_LABELS = {
    direction: 'Direction', bureau: 'Bureau', designation: 'Désignation', code: 'Code',
    qte: 'Qte', etat: 'État', obs: 'Observations'
  };
  var NAV = [
    { id: 'import', label: '1. Importer un document Word' },
    { id: 'preview', label: '2. Prévisualiser les tableaux' },
    { id: 'edit', label: '3. Modifier les données' },
    { id: 'excel', label: '4. Convertir vers Excel' },
    { id: 'export', label: '5. Exporter une archive JSON' },
    { id: 'restore', label: '6. Restaurer une archive JSON' },
    { id: 'recent', label: '7. Gérer les projets récents' }
  ];

  var state = {
    db: null,
    project: null,
    past: [],
    future: [],
    view: 'import',
    tableIndex: 0,
    selectedRow: -1,
    search: '',
    importReport: [],
    flagAck: false,
    recent: [],
    saveTimer: null,
    deferredInstall: null
  };

  /* ---------- Utilitaires ---------- */

  function h(tag, attrs, kids) {
    var el = document.createElement(tag);
    if (attrs) {
      Object.keys(attrs).forEach(function (k) {
        var v = attrs[k];
        if (v === false || v == null) return;
        if (k === 'class') el.className = v;
        else if (k === 'text') el.textContent = v;
        else if (/^on[a-z]+$/.test(k)) el.addEventListener(k.slice(2), v);
        else if (v === true) el.setAttribute(k, '');
        else el.setAttribute(k, String(v));
      });
    }
    (kids || []).forEach(function (c) {
      if (c == null || c === false) return;
      el.appendChild(typeof c === 'string' ? document.createTextNode(c) : c);
    });
    return el;
  }

  function $(id) { return document.getElementById(id); }

  function hhmm() { return new Date().toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit' }); }
  function fmtDate(iso) { return iso ? new Date(iso).toLocaleString('fr-FR') : ''; }
  function stamp() { return new Date().toISOString().slice(0, 16).replace(/[:T]/g, '-'); }
  function slug(name) {
    return String(name || 'inventaire').normalize('NFD').replace(/[̀-ͯ]/g, '')
      .replace(/[^A-Za-z0-9_-]+/g, '_').replace(/^_+|_+$/g, '').slice(0, 60) || 'inventaire';
  }
  function byUpdated(a, b) { return String(b.updatedAt).localeCompare(String(a.updatedAt)); }

  function setStatus(text, isErr) {
    var el = $('status');
    el.textContent = text;
    el.className = isErr ? 'status err' : 'status';
  }

  // opts : { sticky, ms, label, onAction } — un bouton d'action optionnel (ex. « Mettre à jour »).
  var toastTimer = null;
  function hideToast() {
    clearTimeout(toastTimer);
    $('toast').hidden = true;
  }
  function toast(msg, isErr, opts) {
    opts = opts || {};
    var el = $('toast');
    el.replaceChildren();
    el.appendChild(h('span', { class: 'toast-msg', text: msg }));
    if (opts.label && opts.onAction) {
      el.appendChild(h('button', { class: 'toast-act', text: opts.label, onclick: function () { hideToast(); opts.onAction(); } }));
    }
    el.appendChild(h('button', { class: 'toast-x', 'aria-label': 'Fermer', text: '×', onclick: hideToast }));
    el.className = isErr ? 'toast err' : 'toast';
    el.hidden = false;
    clearTimeout(toastTimer);
    if (!opts.sticky) toastTimer = setTimeout(hideToast, opts.ms || 3800);
  }

  function download(name, blob) {
    var url = URL.createObjectURL(blob);
    var a = h('a', { href: url, download: name });
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 1500);
  }

  function readBuffer(file) {
    if (file.arrayBuffer) return file.arrayBuffer();
    return new Promise(function (resolve, reject) {
      var r = new FileReader();
      r.onload = function () { resolve(r.result); };
      r.onerror = function () { reject(r.error); };
      r.readAsArrayBuffer(file);
    });
  }

  function readText(file) {
    return new Promise(function (resolve, reject) {
      var r = new FileReader();
      r.onload = function () { resolve(r.result); };
      r.onerror = function () { reject(r.error); };
      r.readAsText(file, 'utf-8');
    });
  }

  function clampIndex() {
    var n = state.project.tables.length;
    if (state.tableIndex >= n) state.tableIndex = Math.max(0, n - 1);
    if (state.tableIndex < 0) state.tableIndex = 0;
    return state.tableIndex;
  }

  function visibleFields(t) {
    return M.FIELDS.filter(function (f) {
      return t.items.some(function (it) { return it[f] !== '' && it[f] != null; });
    });
  }

  function sourceLine(t) {
    return (t.docName || '—') + ' · tableau n° ' + (t.index || '—') + ' · ' + t.items.length + ' ligne(s)';
  }

  function emptyState(title, text) {
    return h('div', { class: 'empty' }, [h('h3', { text: title }), h('p', { text: text })]);
  }

  /* ---------- Stockage IndexedDB ---------- */

  function openDB() {
    return new Promise(function (resolve, reject) {
      if (!window.indexedDB) return reject(new Error('IndexedDB indisponible.'));
      var req = indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = function () {
        var db = req.result;
        if (!db.objectStoreNames.contains('projects')) db.createObjectStore('projects', { keyPath: 'id' });
        if (!db.objectStoreNames.contains('meta')) db.createObjectStore('meta', { keyPath: 'key' });
      };
      req.onsuccess = function () { resolve(req.result); };
      req.onerror = function () { reject(req.error); };
    });
  }

  function dbRequest(store, mode, makeReq) {
    return new Promise(function (resolve, reject) {
      var t = state.db.transaction(store, mode);
      var req = makeReq(t.objectStore(store));
      t.oncomplete = function () { resolve(req ? req.result : undefined); };
      t.onerror = function () { reject(t.error); };
      t.onabort = function () { reject(t.error || new Error('Transaction annulée.')); };
    });
  }
  function dbGet(store, key) { return dbRequest(store, 'readonly', function (s) { return s.get(key); }); }
  function dbGetAll(store) { return dbRequest(store, 'readonly', function (s) { return s.getAll(); }); }
  function dbPut(store, value) { return dbRequest(store, 'readwrite', function (s) { return s.put(value); }); }
  function dbDelete(store, key) { return dbRequest(store, 'readwrite', function (s) { return s.delete(key); }); }
  function dbClear(store) { return dbRequest(store, 'readwrite', function (s) { return s.clear(); }); }

  function saveNow() {
    if (!state.db || !state.project) return Promise.resolve();
    var p = state.project;
    return dbPut('projects', p).then(function () {
      return dbPut('meta', { key: 'current', value: p.id });
    }).then(function () {
      setStatus('Enregistré à ' + hhmm());
    }).catch(function (e) {
      setStatus('Échec de l’enregistrement : ' + ((e && e.message) || e), true);
    });
  }

  function scheduleSave() {
    clearTimeout(state.saveTimer);
    setStatus('Modification non enregistrée…');
    state.saveTimer = setTimeout(function () { state.saveTimer = null; saveNow(); }, 300);
  }

  function flush() {
    clearTimeout(state.saveTimer);
    state.saveTimer = null;
    return saveNow();
  }

  function loadRecent() {
    if (!state.db) { state.recent = [state.project]; return Promise.resolve(); }
    return dbGetAll('projects').then(function (list) {
      state.recent = list.sort(byUpdated);
    });
  }

  /* ---------- Modifications et annulation ---------- */

  function touched() {
    state.project.updatedAt = new Date().toISOString();
    scheduleSave();
    render();
  }

  // Chaque modification passe ici : un instantané est gardé pour Annuler.
  function mutate(fn) {
    state.past.push(M.clone(state.project.tables));
    if (state.past.length > MAX_HISTORY) state.past.shift();
    state.future = [];
    fn(state.project);
    touched();
  }

  function undo() {
    if (!state.past.length) return toast('Rien à annuler.');
    state.future.push(M.clone(state.project.tables));
    state.project.tables = state.past.pop();
    touched();
  }

  function redo() {
    if (!state.future.length) return toast('Rien à rétablir.');
    state.past.push(M.clone(state.project.tables));
    state.project.tables = state.future.pop();
    touched();
  }

  function setCurrent(p) {
    state.project = p;
    state.past = [];
    state.future = [];
    state.tableIndex = 0;
    state.selectedRow = -1;
    state.search = '';
    state.flagAck = false;
    state.view = 'preview';
    saveNow();
    render();
  }

  /* ---------- Importation ---------- */

  function importFiles(files) {
    var report = [];
    var chain = Promise.resolve();
    files.forEach(function (file) {
      chain = chain.then(function () {
        return readBuffer(file)
          .then(function (buf) { return P.parseDocx(buf, { docName: file.name }); })
          .then(function (parsed) {
            var r = { name: file.name, tables: parsed.tables.length, rows: 0, warnings: parsed.warnings };
            mutate(function (p) {
              var info = M.addDocument(p, parsed);
              r.rows = info.rows;
            });
            report.push(r);
          })
          .catch(function (e) {
            report.push({ name: file.name, error: (e && e.message) || String(e), warnings: [] });
          });
      });
    });
    return chain.then(function () {
      state.importReport = report;
      state.view = 'import';
      render();
      var ok = report.filter(function (r) { return !r.error; }).length;
      toast(ok + ' document(s) importé(s) sur ' + report.length + '.', ok < report.length);
    });
  }

  /* ---------- Vues ---------- */

  function viewImport() {
    var input = h('input', {
      type: 'file', accept: '.docx', multiple: true, hidden: true,
      onchange: function (e) {
        var files = Array.prototype.slice.call(e.target.files || []);
        e.target.value = '';
        if (files.length) importFiles(files);
      }
    });
    var st = M.stats(state.project);
    var sources = state.project.sources || [];

    var report = state.importReport.map(function (r) {
      if (r.error) {
        return h('li', { class: 'err' }, ['✖ ' + r.name + ' — ' + r.error]);
      }
      return h('li', null, [
        h('strong', { text: '✔ ' + r.name }),
        ' — ' + r.tables + ' tableau(x), ' + r.rows + ' ligne(s) importée(s)',
        r.warnings.length ? h('ul', { class: 'warn' }, r.warnings.map(function (w) { return h('li', { text: w }); })) : null
      ]);
    });

    return h('div', { class: 'page' }, [
      h('h2', { text: 'Importer un document Word' }),
      h('p', { class: 'lead', text: 'Choisissez un ou plusieurs fichiers .docx. Les tableaux sont lus dans l’ordre du document. Rien n’est envoyé sur Internet.' }),
      h('div', { class: 'row' }, [
        h('button', { class: 'btn', onclick: function () { input.click(); }, text: 'Importer un fichier Word' }),
        h('button', { class: 'ghost', onclick: function () { setView('preview'); }, text: 'Ouvrir la prévisualisation' }),
        input
      ]),
      h('div', { class: 'card-plain' }, [
        h('strong', { text: 'Projet : ' + state.project.name }),
        h('p', { class: 'meta', text: st.tables + ' tableau(x) · ' + st.items + ' ligne(s) · ' + st.flagged + ' ligne(s) à vérifier' })
      ]),
      report.length ? h('div', null, [h('h3', { text: 'Dernier import' }), h('ul', { class: 'report' }, report)]) : null,
      sources.length ? h('div', null, [
        h('h3', { text: 'Documents du projet' }),
        h('table', { class: 'grid compact' }, [
          h('thead', null, [h('tr', null, ['Document', 'Importé le', 'Tableaux', 'Lignes'].map(function (t) { return h('th', { text: t }); }))]),
          h('tbody', null, sources.map(function (s) {
            return h('tr', null, [
              h('td', { text: s.name }), h('td', { text: fmtDate(s.importedAt) }),
              h('td', { text: String(s.tables) }), h('td', { text: String(s.rows) })
            ]);
          }))
        ])
      ]) : null
    ]);
  }

  function tableNav(active, onPick) {
    return h('aside', { class: 'tabs' }, state.project.tables.map(function (t, i) {
      return h('button', {
        class: i === active ? 'tab on' : 'tab',
        onclick: function () { onPick(i); },
        text: (i + 1) + '. ' + (t.title || 'Tableau ' + (i + 1)) + ' (' + t.items.length + ')'
      });
    }));
  }

  function viewPreview() {
    var tables = state.project.tables;
    if (!tables.length) {
      return emptyState('Aucun tableau pour l’instant', 'Importez un document Word pour commencer.');
    }
    var idx = clampIndex();
    var t = tables[idx];
    var fields = visibleFields(t);
    var table = h('table', { class: 'grid preview' }, [
      h('thead', null, [h('tr', null, [h('th', { text: '#' })].concat(fields.map(function (f) {
        return h('th', { text: FIELD_LABELS[f] });
      })))]),
      h('tbody', null, t.items.map(function (it, i) {
        return h('tr', { class: it.flags.length ? 'flag' : '', title: it.flags.join(' · ') },
          [h('td', { text: String(i + 1) })].concat(fields.map(function (f) {
            return h('td', { text: it[f] || '' });
          })));
      }))
    ]);
    var warnings = (t.meta && t.meta.warnings) || [];
    var flagged = t.items.filter(function (it) { return it.flags.length; }).length;
    return h('div', { class: 'split' }, [
      tableNav(idx, function (i) { state.tableIndex = i; render(); }),
      h('div', { class: 'pane' }, [
        h('h2', { text: t.title || 'Tableau ' + (idx + 1) }),
        h('p', { class: 'meta', text: sourceLine(t) + ' · ' + fields.length + ' colonne(s)' }),
        flagged ? h('p', { class: 'warn', text: '⚠ ' + flagged + ' ligne(s) à vérifier (surlignées).' }) : null,
        warnings.length ? h('ul', { class: 'warn' }, warnings.map(function (w) { return h('li', { text: w }); })) : null,
        t.items.length ? h('div', { class: 'scroll' }, [table]) : emptyState('Tableau vide', 'Ajoutez des lignes dans « Modifier les données ».')
      ])
    ]);
  }

  function viewEdit() {
    var tables = state.project.tables;
    if (!tables.length) {
      return h('div', { class: 'page' }, [
        emptyState('Aucun tableau à modifier', 'Importez un document Word ou créez un tableau vide.'),
        h('div', { class: 'row' }, [h('button', { class: 'btn', onclick: addTable, text: 'Créer un tableau vide' })])
      ]);
    }
    var ti = clampIndex();
    var t = tables[ti];
    var body = h('tbody');
    var count = h('span', { class: 'meta' });

    function fillBody() {
      var q = state.search.trim().toLowerCase();
      var rows = [];
      t.items.forEach(function (it, i) {
        if (q && !rowText(it).includes(q)) return;
        rows.push(editRow(t, it, i));
      });
      body.replaceChildren.apply(body, rows);
      count.textContent = rows.length + ' / ' + t.items.length + ' ligne(s)';
    }
    fillBody();

    var titleInput = h('input', {
      class: 'title-input', value: t.title || '', placeholder: 'Titre du tableau',
      onchange: function (e) {
        var v = e.target.value.trim();
        mutate(function (p) { p.tables[ti].title = v; });
      }
    });

    var search = h('input', {
      type: 'search', class: 'search', placeholder: 'Rechercher Direction, Bureau, désignation…',
      value: state.search,
      oninput: function (e) { state.search = e.target.value; fillBody(); }
    });

    var picker = h('select', {
      class: 'picker',
      onchange: function (e) { state.tableIndex = Number(e.target.value); state.selectedRow = -1; render(); }
    }, tables.map(function (x, i) {
      return h('option', { value: String(i), selected: i === ti,
        text: (i + 1) + '. ' + (x.title || 'Tableau ' + (i + 1)) + ' (' + x.items.length + ' ligne(s))' });
    }));

    return h('div', { class: 'page wide' }, [
      h('div', { class: 'row' }, [h('label', { class: 'meta', text: 'Tableau :' }), picker]),
      h('div', { class: 'pane' }, [
        h('div', { class: 'toolbar' }, [
          h('button', { class: 'ghost', onclick: addRow, text: '+ Ligne' }),
          h('button', { class: 'ghost', onclick: function () { moveRow(-1); }, text: '↑ Monter' }),
          h('button', { class: 'ghost', onclick: function () { moveRow(1); }, text: '↓ Descendre' }),
          h('button', { class: 'ghost danger', onclick: deleteRow, text: 'Supprimer la ligne' }),
          h('span', { class: 'sep' }),
          h('button', { class: 'ghost', onclick: addTable, text: '+ Tableau' }),
          h('button', { class: 'ghost danger', onclick: deleteTable, text: 'Supprimer le tableau' })
        ]),
        h('div', { class: 'row' }, [titleInput, search, count]),
        h('p', { class: 'meta', text: sourceLine(t) }),
        h('div', { class: 'scroll' }, [
          h('table', { class: 'grid edit' }, [
            h('thead', null, [h('tr', null, ['#', 'Direction', 'Bureau', 'Désignation', 'Code', 'Qte', 'État', 'Observations', 'Signalement']
              .map(function (t2) { return h('th', { text: t2 }); }))]),
            body
          ])
        ])
      ])
    ]);
  }

  function rowText(it) {
    return [it.direction, it.bureau, it.designation, it.code].join(' ').toLowerCase();
  }

  function editRow(t, it, i) {
    var cell = function (field) {
      return h('td', null, [h('input', {
        class: 'cell', value: it[field] || '',
        onchange: function (e) { setField(i, field, e.target.value); }
      })]);
    };
    var cls = (i === state.selectedRow ? 'sel ' : '') + (it.flags.length ? 'flag' : '');
    return h('tr', {
      class: cls.trim(),
      onclick: function (e) {
        if (e.target.tagName === 'INPUT') return;
        state.selectedRow = i;
        render();
      }
    }, [
      h('td', { class: 'num', text: String(i + 1) }),
      cell('direction'), cell('bureau'), cell('designation'), cell('code'), cell('qte'), cell('etat'), cell('obs'),
      h('td', { class: 'sig', text: it.flags.join(' · ') })
    ]);
  }

  function setField(i, field, value) {
    mutate(function (p) {
      var it = p.tables[state.tableIndex].items[i];
      it[field] = value.trim();
      it.flags = M.computeFlags(it);
    });
  }

  function addRow() {
    mutate(function (p) {
      var t = p.tables[state.tableIndex];
      var at = state.selectedRow >= 0 ? state.selectedRow + 1 : t.items.length;
      var base = t.items[at - 1] || t.items[t.items.length - 1] || {};
      t.items.splice(at, 0, M.newItem({ direction: base.direction, bureau: base.bureau }));
      state.selectedRow = at;
    });
  }

  function deleteRow() {
    if (state.selectedRow < 0) return toast('Sélectionnez d’abord une ligne (clic sur la ligne).', true);
    mutate(function (p) {
      var t = p.tables[state.tableIndex];
      t.items.splice(state.selectedRow, 1);
      state.selectedRow = Math.min(state.selectedRow, t.items.length - 1);
    });
    toast('Ligne supprimée. Annuler : Ctrl+Z.');
  }

  function moveRow(delta) {
    var t = state.project.tables[state.tableIndex];
    var from = state.selectedRow;
    var to = from + delta;
    if (from < 0) return toast('Sélectionnez d’abord une ligne.', true);
    if (to < 0 || to >= t.items.length) return;
    mutate(function (p) {
      var items = p.tables[state.tableIndex].items;
      var tmp = items[from];
      items[from] = items[to];
      items[to] = tmp;
      state.selectedRow = to;
    });
  }

  function addTable() {
    mutate(function (p) {
      p.tables.push(M.emptyTable('Nouveau tableau ' + (p.tables.length + 1)));
      state.tableIndex = p.tables.length - 1;
      state.selectedRow = -1;
    });
  }

  function deleteTable() {
    var t = state.project.tables[clampIndex()];
    if (!t) return;
    if (!window.confirm('Supprimer le tableau « ' + (t.title || '') + ' » et ses ' + t.items.length + ' ligne(s) ?\n(Annulable avec Ctrl+Z.)')) return;
    mutate(function (p) {
      p.tables.splice(state.tableIndex, 1);
      state.selectedRow = -1;
    });
  }

  function viewExcel() {
    var st = M.stats(state.project);
    var flagged = [];
    state.project.tables.forEach(function (t) {
      t.items.forEach(function (it) {
        if (it.flags.length && !X.isBlank(it)) flagged.push({ title: t.title, it: it });
      });
    });
    var items = X.rowsOf(state.project).length;
    var blanks = X.blankCount(state.project);
    var list = flagged.slice(0, 100).map(function (f) {
      var o = f.it.origin || {};
      return h('li', null, [
        h('strong', { text: (f.it.designation || '(sans désignation)') + ' — code ' + (f.it.code || '—') }),
        ' · ' + f.title + (o.row ? ' · ligne ' + o.row : ''),
        h('div', { class: 'warn', text: f.it.flags.join(' · ') })
      ]);
    });
    var canConvert = items > 0 && (!flagged.length || state.flagAck);

    return h('div', { class: 'page' }, [
      h('h2', { text: 'Convertir vers Excel' }),
      h('p', { class: 'lead', text: 'Un seul classeur regroupe tous les tableaux, dans leur ordre. Colonnes : Direction | Bureau | Désignation | Code | Qte | État.' }),
      h('div', { class: 'card-plain' }, [
        h('p', { text: st.tables + ' tableau(x) · ' + items + ' ligne(s) à exporter · ' + st.flagged + ' ligne(s) signalée(s)' }),
        blanks ? h('p', { class: 'meta', text: blanks + ' ligne(s) entièrement vide(s) ignorée(s) à l’export.' }) : null
      ]),
      flagged.length ? h('div', null, [
        h('h3', { text: 'Lignes à vérifier avant conversion' }),
        h('ul', { class: 'report' }, list),
        flagged.length > 100 ? h('p', { class: 'meta', text: '… et ' + (flagged.length - 100) + ' autre(s).' }) : null,
        h('label', { class: 'check' }, [
          h('input', {
            type: 'checkbox', checked: state.flagAck,
            onchange: function (e) { state.flagAck = e.target.checked; render(); }
          }),
          ' J’ai vérifié les lignes signalées'
        ])
      ]) : null,
      h('p', { class: 'meta', text: 'La colonne « Observations » est conservée dans l’archive JSON ; le classeur Excel suit la structure à six colonnes imposée.' }),
      h('div', { class: 'row' }, [
        h('button', { class: 'btn', disabled: !canConvert, onclick: convertToExcel, text: 'Convertir vers Excel' })
      ])
    ]);
  }

  function convertToExcel() {
    X.buildXlsx(state.project, { XLSX: window.XLSX, JSZip: window.JSZip })
      .then(function (bytes) {
        download(slug(state.project.name) + '.xlsx', new Blob([bytes], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }));
        toast('Fichier Excel généré.');
      })
      .catch(function (e) { toast((e && e.message) || 'Conversion impossible.', true); });
  }

  function viewExport() {
    return h('div', { class: 'page' }, [
      h('h2', { text: 'Exporter une archive JSON' }),
      h('p', { class: 'lead', text: 'Sauvegarde complète du projet : données importées, tableaux, modifications, Directions et Bureaux. Le fichier est autonome et peut être restauré sur un autre ordinateur.' }),
      h('div', { class: 'row' }, [h('button', { class: 'btn', onclick: exportArchive, text: 'Exporter une archive JSON' })]),
      state.project.lastExport
        ? h('p', { class: 'meta', text: 'Dernière sauvegarde de ce projet : ' + fmtDate(state.project.lastExport) })
        : h('p', { class: 'warn', text: 'Ce projet n’a encore jamais été sauvegardé en JSON. Pensez-y régulièrement.' })
    ]);
  }

  function exportArchive() {
    var text = JSON.stringify(M.createArchive(state.project), null, 2);
    download(slug(state.project.name) + '_' + stamp() + '.json', new Blob([text], { type: 'application/json' }));
    state.project.lastExport = new Date().toISOString();
    scheduleSave();
    render();
    toast('Archive JSON créée.');
  }

  function viewRestore() {
    var input = h('input', {
      type: 'file', accept: '.json,application/json', hidden: true,
      onchange: function (e) {
        var f = e.target.files && e.target.files[0];
        e.target.value = '';
        if (f) restoreFile(f);
      }
    });
    return h('div', { class: 'page' }, [
      h('h2', { text: 'Restaurer une archive JSON' }),
      h('p', { class: 'lead', text: 'L’archive est ouverte comme un nouveau projet. Aucun projet existant n’est remplacé.' }),
      h('div', { class: 'row' }, [
        h('button', { class: 'btn', onclick: function () { input.click(); }, text: 'Choisir une archive JSON' }),
        input
      ])
    ]);
  }

  function restoreFile(file) {
    readText(file)
      .then(function (text) {
        var p = M.restoreArchive(text, ' (restaurée)');
        if (!window.confirm('Ouvrir « ' + p.name + ' » comme nouveau projet ?')) return;
        return flush().then(function () { setCurrent(p); toast('Projet restauré.'); });
      })
      .catch(function (e) { toast((e && e.message) || 'Archive illisible.', true); });
  }

  function viewRecent() {
    var cards = state.recent.map(function (p) {
      var st = M.stats(p);
      var isCur = p.id === state.project.id;
      return h('div', { class: 'card' + (isCur ? ' cur' : '') }, [
        h('div', { class: 'grow' }, [
          h('strong', { text: p.name }),
          h('div', { class: 'meta', text: fmtDate(p.updatedAt) + ' · ' + st.tables + ' tableau(x) · ' + st.items + ' ligne(s)' })
        ]),
        isCur ? h('span', { class: 'badge', text: 'Ouvert' }) : h('button', { class: 'ghost', onclick: function () { openProject(p.id); }, text: 'Ouvrir' }),
        h('button', { class: 'ghost', onclick: function () { renameProject(p); }, text: 'Renommer' }),
        h('button', { class: 'ghost danger', onclick: function () { deleteProject(p); }, text: 'Supprimer' })
      ]);
    });
    return h('div', { class: 'page' }, [
      h('h2', { text: 'Projets récents' }),
      h('div', { class: 'row' }, [h('button', { class: 'btn', onclick: newProject, text: 'Nouveau projet' })]),
      h('div', { class: 'list' }, cards.length ? cards : [h('p', { class: 'meta', text: 'Aucun projet enregistré.' })])
    ]);
  }

  function openProject(id) {
    return flush()
      .then(function () { return dbGet('projects', id); })
      .then(function (p) { if (p) setCurrent(p); });
  }

  function newProject() {
    flush().then(function () { setCurrent(M.newProject()); toast('Nouveau projet créé.'); });
  }

  function renameProject(p) {
    var name = window.prompt('Nouveau nom du projet :', p.name);
    if (!name || !name.trim()) return;
    p.name = name.trim();
    p.updatedAt = new Date().toISOString();
    if (p.id === state.project.id) {
      state.project.name = p.name;
      scheduleSave();
      render();
    } else {
      dbPut('projects', p).then(function () { return loadRecent(); }).then(render);
    }
  }

  function deleteProject(p) {
    if (!window.confirm('Supprimer définitivement le projet « ' + p.name + ' » ?')) return;
    dbDelete('projects', p.id).then(function () {
      if (p.id === state.project.id) {
        state.project = M.newProject();
        state.past = [];
        state.future = [];
        return saveNow().then(function () { return loadRecent(); }).then(render);
      }
      return loadRecent().then(render);
    });
  }

  /* ---------- Rendu principal ---------- */

  function setView(name) {
    state.view = name;
    if (name === 'recent') {
      loadRecent().then(render);
      return;
    }
    render();
  }

  function renderNav() {
    var nav = $('nav');
    nav.replaceChildren.apply(nav, NAV.map(function (n) {
      return h('button', {
        class: n.id === state.view ? 'nav on' : 'nav',
        onclick: function () { setView(n.id); },
        text: n.label
      });
    }));
  }

  function render() {
    if (!state.project) return;
    renderNav();
    $('proj-name').value = state.project.name;
    $('btn-undo').disabled = !state.past.length;
    $('btn-redo').disabled = !state.future.length;

    var builders = {
      import: viewImport, preview: viewPreview, edit: viewEdit, excel: viewExcel,
      export: viewExport, restore: viewRestore, recent: viewRecent
    };
    var build = builders[state.view] || viewImport;
    $('view').replaceChildren(build());
  }

  /* ---------- Démarrage ---------- */

  function init() {
    $('btn-undo').addEventListener('click', undo);
    $('btn-redo').addEventListener('click', redo);
    $('proj-name').addEventListener('change', function (e) {
      var v = e.target.value.trim() || 'Projet sans titre';
      state.project.name = v;
      touched();
    });
    $('btn-clear').addEventListener('click', clearAll);
    $('btn-install').addEventListener('click', installHelp);
    var standalone = (window.matchMedia && window.matchMedia('(display-mode: standalone)').matches) || window.navigator.standalone === true;
    if (standalone) $('btn-install').hidden = true;
    var vEl = $('app-version');
    if (vEl && CFG.version) vEl.textContent = 'version ' + CFG.version;

    document.addEventListener('keydown', function (e) {
      if (e.target && e.target.matches && e.target.matches('input, textarea, select')) return;
      var k = e.key ? e.key.toLowerCase() : '';
      if ((e.ctrlKey || e.metaKey) && k === 'z' && !e.shiftKey) { e.preventDefault(); undo(); }
      else if ((e.ctrlKey || e.metaKey) && (k === 'y' || (k === 'z' && e.shiftKey))) { e.preventDefault(); redo(); }
    });
    window.addEventListener('pagehide', function () {
      if (state.saveTimer) flush();
    });
    window.addEventListener('beforeinstallprompt', function (e) {
      e.preventDefault();
      state.deferredInstall = e;
    });
    window.addEventListener('appinstalled', function () {
      state.deferredInstall = null;
      $('btn-install').hidden = true;
      toast('Application installée. Vous pouvez la lancer depuis le bureau ou l’écran d’accueil.', false, { ms: 6000 });
    });

    return openDB().then(function (db) {
      state.db = db;
      return Promise.all([dbGetAll('projects'), dbGet('meta', 'current')]);
    }).then(function (res) {
      var list = res[0];
      var cur = res[1];
      var found = cur ? list.filter(function (p) { return p.id === cur.value; })[0] : null;
      state.project = found || list.sort(byUpdated)[0] || M.newProject();
      return saveNow();
    }).catch(function (e) {
      if (!state.project) state.project = M.newProject();
      setStatus('Stockage local indisponible : ' + ((e && e.message) || e) + '. Les données ne seront pas conservées.', true);
    }).then(function () {
      render();
      registerSW();
      if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(function () {});
    });
  }

  // Bouton « Installer » : invite native si disponible, sinon marche à suivre selon l'appareil.
  function installHelp() {
    if (state.deferredInstall) {
      // Une invite native ne peut servir qu'une fois : on la consomme, puis le message d'aide prendra le relais.
      var ev = state.deferredInstall;
      state.deferredInstall = null;
      ev.prompt();
      return;
    }
    var ua = navigator.userAgent || '';
    var msg;
    if (/iPhone|iPad|iPod/.test(ua)) {
      msg = 'Sur iPhone ou iPad : ouvrez cette page dans Safari, touchez le bouton Partager, puis « Sur l’écran d’accueil ».';
    } else if (/Android/.test(ua)) {
      msg = 'Sur Android : menu ⋮ de Chrome, puis « Installer l’application » (ou « Ajouter à l’écran d’accueil »).';
    } else {
      msg = 'Sur ordinateur : cliquez sur l’icône d’installation dans la barre d’adresse de Chrome ou Edge, ou menu ⋮ puis « Installer Convertisseur Word vers Excel ».';
    }
    toast(msg, false, { ms: 12000 });
  }

  // Service worker : la mise à jour n'est appliquée qu'après accord de l'utilisateur.
  function registerSW() {
    if (!('serviceWorker' in navigator)) return;
    var wantUpdate = false;
    var reloading = false;
    navigator.serviceWorker.addEventListener('controllerchange', function () {
      if (wantUpdate && !reloading) {
        reloading = true;
        window.location.reload();
      }
    });
    navigator.serviceWorker.register('./sw.js', { scope: './' }).then(function (reg) {
      function offerUpdate() {
        toast('Une nouvelle version de l’application est disponible.', false, {
          sticky: true,
          label: 'Mettre à jour',
          onAction: function () {
            wantUpdate = true;
            flush().then(function () {
              if (reg.waiting) reg.waiting.postMessage({ type: 'SKIP_WAITING' });
            });
          }
        });
      }
      if (reg.waiting && navigator.serviceWorker.controller) offerUpdate();
      reg.addEventListener('updatefound', function () {
        var nw = reg.installing;
        if (!nw) return;
        nw.addEventListener('statechange', function () {
          if (nw.state === 'installed' && navigator.serviceWorker.controller) offerUpdate();
        });
      });
    }).catch(function (e) {
      setStatus('Mode hors ligne indisponible : ' + ((e && e.message) || e), true);
    });
  }

  function clearAll() {
    if (!window.confirm('Vider toutes les données locales ? Tous les projets seront effacés de cet ordinateur. Cette action est définitive.')) return;
    if (!state.db) { window.location.reload(); return; }
    Promise.all([dbClear('projects'), dbClear('meta')]).then(function () {
      window.location.reload();
    }).catch(function (e) { toast('Effacement impossible : ' + e, true); });
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();

  // Accès de diagnostic pour les tests automatisés.
  window.CWE = { state: state, render: render, importFiles: importFiles, setView: setView, undo: undo, redo: redo };
})();
