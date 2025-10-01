# Macro Word — Masquage Interactif (Active Recall)

Macro VBA pour Microsoft Word permettant de masquer/révéler du contenu à l’aide de rectangles interactifs (Active Recall).  
Idéal pour la révision, l’auto-évaluation, les QCM, ou l’apprentissage progressif.

## Fonctionnalités principales

- Masquage de texte par rectangles de couleur (mode dessin)
- Révélation progressive (mode révision)
- Changement de couleur des masques
- Statistiques de progression
- Suppression de tous les masques
- Aide intégrée et instructions

## Installation

1. Ouvrez Word, puis appuyez sur `ALT+F11` pour ouvrir l’éditeur VBA.
2. Faites `Fichier > Importer un fichier...` et sélectionnez `MasquageInteractif.bas`.
3. Copiez la partie “Document_Open” dans le module `ThisDocument` de votre document Word si besoin.
4. Enregistrez votre document au format `.docm` (Word avec macros).

## Utilisation rapide

- Cliquez sur “Activer Mode Dessin”, puis utilisez `Insertion > Formes > Rectangle` pour dessiner sur les zones à masquer.
- Cliquez sur “Appliquer Masquage” : les rectangles deviennent interactifs.
- Pour révéler : double-cliquez sur un rectangle ou sélectionnez-le puis cliquez sur “Révéler Sélection”.
- Utilisez “Basculer Tous” pour tout révéler/masquer d’un coup.
- “Changer Couleur” : modifie la couleur de tous les masques.
- “Statistiques” : affiche la progression de votre révision.

## Conseils

- Créez des boutons personnalisés dans le ruban pour un accès rapide.
- Utilisez le mode révision pour tester vos connaissances sans tricher !

## Dépendances
- Microsoft Word (version supportant VBA)

## Auteur(e) / Licence

- Auteur : Yoxen
- Licence : Libre
