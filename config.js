/*
 * config.js — paramètres de l'application (une seule source pour l'interface).
 * Lors d'une mise à jour, changer aussi CACHE_VERSION dans sw.js (voir README_GITHUB.md).
 */
(function (g) {
  'use strict';
  g.CWE_CONFIG = {
    appName: 'Convertisseur Word vers Excel',
    version: '1.0.1',
    dbName: 'convertisseur-word-excel',
    dbVersion: 1,
    developer: 'BOUAFIA Nourelabedine'
  };
})(typeof self !== 'undefined' ? self : window);
