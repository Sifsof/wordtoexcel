/*
 * project-model.js — modèle de projet, contrôles de données et archive JSON autonome.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.ProjectModel = factory();
  }
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var APP = 'convertisseur-word-vers-excel';
  var FORMAT_VERSION = 1;
  var FIELDS = ['direction', 'bureau', 'designation', 'code', 'qte', 'etat', 'obs'];

  var seq = 0;
  function uid() {
    seq += 1;
    return 'p' + Date.now().toString(36) + seq.toString(36) + Math.random().toString(36).slice(2, 6);
  }
  function nowIso() { return new Date().toISOString(); }
  function str(v) { return v == null ? '' : String(v); }
  function clone(o) { return JSON.parse(JSON.stringify(o)); }

  function isNumeric(s) {
    return /^-?\d+([.,]\d+)?$/.test(str(s).replace(/\s/g, ''));
  }

  function computeFlags(it) {
    var f = [];
    if (!it.direction) f.push('Direction non détectée');
    if (!it.bureau) f.push('Bureau non détecté');
    if (!it.designation) f.push('Désignation manquante');
    if (!it.code) f.push('Code manquant');
    if (!it.qte) f.push('Quantité manquante');
    else if (!isNumeric(it.qte)) f.push('Quantité non numérique');
    if (/\btotal\b/i.test([it.direction, it.bureau, it.designation, it.code, it.etat].map(str).join(' '))) {
      f.push('Ligne possiblement un total');
    }
    return f;
  }

  function newProject(name) {
    var t = nowIso();
    return { id: uid(), name: name || 'Projet sans titre', createdAt: t, updatedAt: t, sources: [], tables: [] };
  }

  function newItem(fields) {
    var it = {
      id: uid(),
      direction: str(fields.direction), bureau: str(fields.bureau),
      designation: str(fields.designation), code: str(fields.code),
      qte: str(fields.qte), etat: str(fields.etat), obs: str(fields.obs),
      flags: [], origin: fields.origin || null
    };
    it.flags = computeFlags(it);
    return it;
  }

  // Ajoute au projet les tableaux d'un document importé, dans leur ordre d'apparition.
  function addDocument(project, parsed) {
    var total = 0;
    parsed.tables.forEach(function (t) {
      t.items = t.items.map(function (it) {
        var x = newItem(it);
        x.id = it.id || x.id;
        return x;
      });
      total += t.items.length;
      project.tables.push(t);
    });
    project.sources.push({
      name: parsed.docName,
      importedAt: nowIso(),
      tables: parsed.tables.length,
      rows: total,
      warnings: parsed.warnings || []
    });
    project.updatedAt = nowIso();
    return { tables: parsed.tables.length, rows: total, warnings: parsed.warnings || [] };
  }

  function emptyTable(title) {
    return {
      id: uid(), title: title || 'Nouveau tableau', docName: '(saisie manuelle)', index: 0,
      columns: FIELDS.slice(), items: [],
      meta: { headerFound: true, physicalRows: 0, counts: {}, warnings: [] }
    };
  }

  function stats(project) {
    var items = 0, flagged = 0;
    (project.tables || []).forEach(function (t) {
      t.items.forEach(function (it) {
        items += 1;
        if (it.flags && it.flags.length) flagged += 1;
      });
    });
    return { tables: project.tables.length, items: items, flagged: flagged };
  }

  function normalizeItem(it) {
    return newItem({
      direction: it.direction, bureau: it.bureau, designation: it.designation, code: it.code,
      qte: it.qte, etat: it.etat, obs: it.obs, origin: it.origin || null
    });
  }

  function normalizeTable(t) {
    return {
      id: str(t.id) || uid(),
      title: str(t.title),
      docName: str(t.docName),
      index: Number(t.index) || 0,
      columns: Array.isArray(t.columns) ? t.columns.map(str) : FIELDS.slice(),
      items: (Array.isArray(t.items) ? t.items : []).map(normalizeItem),
      meta: t.meta && typeof t.meta === 'object' ? t.meta : { headerFound: true, physicalRows: 0, counts: {}, warnings: [] }
    };
  }

  function createArchive(project) {
    return {
      app: APP,
      formatVersion: FORMAT_VERSION,
      exportedAt: nowIso(),
      project: {
        name: project.name,
        createdAt: project.createdAt,
        updatedAt: project.updatedAt,
        sources: clone(project.sources || []),
        tables: clone(project.tables || [])
      }
    };
  }

  // Lit une archive (texte ou objet). Crée toujours un nouveau projet : rien n'est écrasé.
  function restoreArchive(input, suffix) {
    var obj = input;
    if (typeof input === 'string') {
      try { obj = JSON.parse(input); } catch (e) { throw new Error('Le fichier n’est pas un JSON valide.'); }
    }
    if (!obj || obj.app !== APP) {
      throw new Error('Cette archive ne provient pas de l’application Convertisseur Word vers Excel.');
    }
    if (Number(obj.formatVersion) > FORMAT_VERSION) {
      throw new Error('Cette archive a été créée par une version plus récente de l’application.');
    }
    var p = obj.project;
    if (!p || !Array.isArray(p.tables)) throw new Error('Archive incomplète : aucun tableau trouvé.');
    var project = newProject((str(p.name) || 'Projet restauré') + (suffix || ''));
    project.createdAt = str(p.createdAt) || project.createdAt;
    project.sources = Array.isArray(p.sources) ? p.sources : [];
    project.tables = p.tables.map(normalizeTable);
    return project;
  }

  return {
    APP: APP,
    FORMAT_VERSION: FORMAT_VERSION,
    FIELDS: FIELDS,
    uid: uid,
    isNumeric: isNumeric,
    computeFlags: computeFlags,
    newProject: newProject,
    newItem: newItem,
    addDocument: addDocument,
    emptyTable: emptyTable,
    stats: stats,
    createArchive: createArchive,
    restoreArchive: restoreArchive,
    clone: clone
  };
}));
