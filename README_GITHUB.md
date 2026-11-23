# Convertisseur Word vers Excel — version GitHub Pages (1.0.1)

Adresse de l'application : **https://sifsof.github.io/wordtoexcel/**

Application web statique : aucun serveur, aucun Node.js, aucun PowerShell. Après une première ouverture avec Internet,
elle fonctionne hors ligne et s'installe comme une application (Windows, Android, iPhone/iPad).

> La version Windows (`Lancer.bat`, serveur PowerShell, `localhost:8765`) est un projet distinct, non modifié.

---

## 1. Règle essentielle : tous les fichiers à la racine

Le dépôt doit contenir **tous les fichiers à la racine, sans aucun sous-dossier**. Les pages du site appellent
les fichiers par leur nom simple (`app.js`, `style.css`…). Si un fichier se retrouve dans un sous-dossier,
la page s'affiche vide et la console signale des erreurs 404.

Une garde de démarrage affiche alors un encadré rouge nommant les fichiers introuvables.

## 2. Contenu (20 fichiers, à plat)

| Fichier | Rôle |
|---|---|
| `index.html` | page principale (références vérifiées) |
| `style.css`, `responsive.css` | apparence ; adaptation téléphone |
| `config.js` | nom, version (1.0.1), nom de la base locale |
| `app.js` | interface, stockage, annulation, archives JSON, installation |
| `docx-parser.js` | lecture des tableaux des documents Word (.docx) |
| `project-model.js` | projet, contrôles des données, archive JSON |
| `excel-export.js` | génération du classeur Excel (.xlsx) |
| `jszip.min.js` | décompression des .docx (JSZip 3.10.2) |
| `xlsx.bundle.js` | écriture des .xlsx (xlsx-js-style, bundle 1.2.0-beta) |
| `sw.js` | service worker : fonctionnement hors ligne (CACHE_VERSION 1.0.1) |
| `manifest.webmanifest` | fiche d'installation |
| `favicon.ico`, `icon-192.png`, `icon-512.png`, `icon-maskable-512.png` | icônes |
| `LICENSE-jszip.md`, `LICENSE-xlsx-js-style.txt` | licences des bibliothèques |
| `.nojekyll` | indique à GitHub Pages de ne pas traiter le site avec Jekyll |
| `README_GITHUB.md` | le présent document |

Le fichier `README.md` déjà présent dans le dépôt n'est pas fourni ici : **ne le remplacez pas et ne le supprimez pas**.

## 3. Publication sur le dépôt

Dépôt : **https://github.com/Sifsof/wordtoexcel**

### Méthode A — interface GitHub

1. Ouvrez le dépôt, puis **Add file › Upload files**.
2. Ouvrez le dossier `WordToExcel_GitHub_Pages_CORRIGE` (ou décompressez l'archive), sélectionnez **les fichiers**
   (Ctrl+A) et glissez-les dans la zone d'envoi.
   **Ne glissez jamais le dossier lui-même** : il créerait un sous-dossier.
3. Vérifiez que la liste d'envoi ne montre aucun dossier, puis **Commit changes**.
   Le fichier `.nojekyll` est caché sous Windows : activez « Éléments masqués » dans l'Explorateur si besoin.

### Méthode B — Git

```bash
git clone https://github.com/Sifsof/wordtoexcel.git
cd wordtoexcel
# copiez ici les FICHIERS (pas le dossier) de WordToExcel_GitHub_Pages_CORRIGE
git add .
git commit -m "Version 1.0.1 - fichiers à plat, chemins corrigés"
git push
```

### Fichiers à remplacer dans le dépôt existant

Le dépôt contient déjà la plupart des fichiers à la racine. Pour la correction 1.0.1, remplacez au minimum :
`index.html`, `sw.js`, `config.js`, `docx-parser.js`, `README_GITHUB.md`.
Les autres fichiers sont identiques à ceux déjà présents.

## 4. Paramètres GitHub Pages

1. **Settings › Pages**.
2. **Source** : *Deploy from a branch* ; **Branch** : `main` ; dossier : `/ (root)` ; **Save**.
3. Attendez 1 à 3 minutes, puis ouvrez https://sifsof.github.io/wordtoexcel/.

Avec un compte gratuit, GitHub Pages exige un dépôt **public**. Le dépôt ne contient que le code :
aucune donnée d'inventaire. Les données restent dans le navigateur de chaque utilisateur.

## 5. Vérifications après publication

- [ ] Le menu à 7 entrées s'affiche, avec « version 1.0.1 » en bas à gauche.
- [ ] Aucun encadré rouge « L’application n’a pas pu démarrer ».
- [ ] Importer un document Word affiche les tableaux.
- [ ] Convertir vers Excel télécharge un `.xlsx`.
- [ ] Après une première ouverture avec Internet, coupez le réseau et rechargez : la page s'ouvre toujours.
- [ ] Console du navigateur (F12) : aucune erreur 404.

## 6. Installation

- **Windows (Chrome ou Edge)** : icône d'installation dans la barre d'adresse, ou menu ⋮ › *Installer Convertisseur Word vers Excel*.
- **Android (Chrome)** : menu ⋮ › *Installer l'application*.
- **iPhone / iPad (Safari uniquement)** : Partager › *Sur l'écran d'accueil*.

## 7. Mise à jour

Changez **deux** numéros, sinon les navigateurs ne détectent pas la mise à jour :
1. `config.js` : `version` (par exemple `'1.0.2'`) ;
2. `sw.js` : `CACHE_VERSION` (par exemple `'wordtoexcel-gh-1.0.2'`).

Puis remplacez les fichiers modifiés et validez. L'utilisateur voit le bandeau « Mettre à jour » ; rien n'est appliqué sans son accord.

## 8. Données et sauvegardes

- Le site GitHub et la version Windows (`localhost:8765`) sont deux adresses différentes : **leurs données ne se voient pas**.
  Pour passer un projet de l'une à l'autre : **5. Exporter une archive JSON**, puis **6. Restaurer une archive JSON**.
- Vider les données du site efface tout. Exportez une archive JSON régulièrement.

## 9. Licences

| Bibliothèque | Version | Licence |
|---|---|---|
| JSZip | 3.10.2 | MIT (double licence MIT / GPL v3 ; MIT retenue) |
| xlsx-js-style (fork de SheetJS Community Edition) | 1.2.0 | Apache 2.0 |

## 10. Dépannage

| Symptôme | Cause probable | Solution |
|---|---|---|
| Encadré rouge « n’a pas pu démarrer » | Fichier dans un sous-dossier ou absent | Remettez le fichier nommé à la racine du dépôt |
| Page vide, erreurs 404 dans la console | Même cause | Idem ; vérifiez qu'aucun dossier n'existe dans le dépôt |
| Ancienne version affichée | Cache du navigateur | Acceptez « Mettre à jour » ou actualisez avec Ctrl+F5 |
| Pas de bouton d'installation | Navigateur non compatible | Chrome ou Edge ; Safari sur iPhone |
| Le site affiche la liste des fichiers | Pages non activé sur `main` | Refaites la section 4 |
