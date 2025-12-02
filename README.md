# Macro Word — Masquage Interactif (Active Recall)

Macro VBA pour Microsoft Word permettant de masquer/révéler du contenu à l’aide de rectangles interactifs (Active Recall). 

## Fonctionnalités principales

- Masquage de texte par rectangles de couleur (mode dessin)
- Révélation progressive (mode révision)
- Changement de couleur des masques
- Statistiques de progression
- Suppression de tous les masques

## Installation

1. Ouvrez Word, puis appuyez sur `ALT+F11` pour ouvrir l’éditeur VBA.
2. Faites `Fichier > Importer un fichier...` et sélectionnez `MasquageInteractif.bas`.
3. Copiez la partie “Document_Open” dans le module `ThisDocument` de votre document Word si besoin.
4. Enregistrez votre document au format `.docm` (Word avec macros).


## Conseils

- Créez des boutons personnalisés dans le ruban pour un accès rapide.

## Dépendances
- Microsoft Word (version supportant VBA)- marche uniquement sur Word
