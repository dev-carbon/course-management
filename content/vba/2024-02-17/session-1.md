---
title: Introduction à VBA
students:
  - Ceren KARAYILAN
---

# Introduction à VBA

## Qu'est-ce que VBA ?

VBA (Visual Basic for Applications) est un langage de programmation orienté objet intégré à Microsoft Office.
VBA est principalement utilisé pour automatiser des tâches répétitives dans Excel, Word et Access.
VBA permet de créer des macros, d'interagir avec les objets des applications Office et d'améliorer l'efficacité des processus de travail.

### Pourquoi utiliser VBA ?

- Automatisation des tâches répétitives (mise en forme, calculs, extraction de données).
- Création d'interfaces utilisateur interactives (formulaires, boîtes de dialogue).
- Personnalisation des fonctionnalités d'Excel et des autres applications Office.
- Intégration avec d'autres applications (Outlook, Word, Access, etc.).

## Caractéristiques principales

- **Langage de programmation orienté objet** : VBA repose sur des objets comme les feuilles, les cellules et les classeurs.
- **Syntaxe simple et claire** : Basé sur Visual Basic, il est facile à apprendre et à comprendre.
- **Grande communauté active** : De nombreux forums et ressources en ligne pour apprendre et résoudre des problèmes.
- **Accès direct aux fonctionnalités d'Excel** : Manipulation des feuilles, des cellules et des graphiques via le code.
- **Exécution rapide des macros** : Permet d'automatiser des tâches en quelques secondes.
- **Compatibilité avec Office** : Fonctionne sur toutes les versions d’Excel et autres applications Office.
- **Structures de contrôle avancées** : Permet d'exécuter du code conditionnellement (`If...Then...Else`), de répéter des actions (`For...Next`, `Do...Loop`), de gérer des sélections multiples (`Select Case`) et de capturer des erreurs (`On Error Resume Next`).

## Terminologies importantes

1. **Macro**
   Une séquence d'instructions enregistrée dans VBA qui permet d'automatiser des tâches répétitives.

2. **Procédure**
   Une procédure est un bloc de code VBA exécutant une tâche spécifique. Elle peut être de type Sub (ne retourne pas de valeur) ou Function (retourne une valeur).

3. **Module**
   Un conteneur de code VBA où sont stockées les procédures et fonctions. Il peut être standard (Module) ou de classe (Class Module).

4. **Objet**
   Une entité manipulable en VBA, comme une feuille de calcul, une cellule ou un graphique. Chaque objet possède des propriétés et des méthodes.

5. **Propriété**
   Une caractéristique d'un objet (ex. : `Range("A1").Value` donne la valeur de la cellule A1).

6. **Méthode**
   Une action qu'un objet peut exécuter (ex. : `Range("A1").Select` sélectionne la cellule A1).

7. **Workbook (Classeur)**
   Représente un fichier Excel. Il contient une ou plusieurs feuilles de calcul (Sheets). Exemple d'accès à un classeur : `Workbooks("MonFichier.xlsx")`.

8. **Worksheet (Feuille de calcul)**
   Une feuille spécifique d'un classeur. Exemple d'accès : `Worksheets("Feuil1")` ou `ActiveSheet` pour la feuille active.

9. **Collection**
   Un ensemble d'objets similaires regroupés. Par exemple, Workbooks est une collection contenant tous les classeurs ouverts, et Worksheets est une collection contenant toutes les feuilles d'un classeur.

10. **Range**
    Une ou plusieurs cellules dans une feuille de calcul. Exemple : `Range("A1")` fait référence à la cellule A1, tandis que `Range("A1:B2")` sélectionne un ensemble de cellules.
## Premiers pas avec VBA

### Préparation de l'environnement

Avant de commencer, il est important d'ajouter l'onglet **Développeur** à l'interface Excel :

1. Ouvrir Excel et cliquer sur **Fichier** > **Options**.
2. Dans la fenêtre **Options Excel**, aller dans **Personnaliser le ruban**.
3. Dans la colonne de droite, cocher **Développeur**, puis cliquer sur **OK**.
4. L'onglet **Développeur** est maintenant visible dans le ruban.

### Accéder à l'éditeur VBA

1. Aller dans l'onglet **Développeur**.
2. Cliquer sur **Visual Basic** pour ouvrir l'éditeur VBA (**VBE**).
3. Explorer l'interface de l'éditeur **VBA**, notamment :
   - **Project Explorer** : affiche les modules, feuilles et objets du projet.
   - **Properties Window** : permet de modifier les propriétés des objets sélectionnés.
   - **Immediate Window** : une console pour exécuter du code et tester des commandes rapidement.
4. Ajouter des boutons utiles à la barre d'outils :
   - **Compile Project** : Vérifie la syntaxe et signale les erreurs.
   - **Comment Block** : Ajoute des commentaires à plusieurs lignes de code.
   - **UnComment Block** : Supprime les commentaires d'un bloc de code.

## Enregistrer et exécuter une macro

1. Aller dans l'onglet **Développeur** et cliquer sur **Enregistrer une macro**.
2. Donner un nom à la macro et choisir où la stocker.
3. Effectuer des actions dans Excel (ex. : mise en forme d’une cellule).
4. Arrêter l’enregistrement et exécuter la macro via **Macros**.

### Différence entre **Absolute Cell Reference** et **Relative Cell Reference** :

- **Absolute Cell Reference (Référence absolue)** : Les cellules sont enregistrées avec des références fixes (ex. `$A$1`). Lorsqu'on exécute la macro, elle applique les actions **exactement sur ces cellules**, quel que soit l'emplacement du curseur.
- **Relative Cell Reference (Référence relative)** : Les cellules sont enregistrées en fonction de la position actuelle (ex. `A1` sans `$`). Lorsqu'on exécute la macro, elle applique les actions **par rapport à la cellule active**, ce qui la rend plus flexible.

⚠️ **Lors de l'enregistrement d'une macro, il est possible de choisir entre mode absolu et mode relatif via l'option "Référence relative" dans l'onglet Développeur.**

## Exercice : (Fichier **vehicles.xlsx**)

### Objectif :

Créer une macro qui formate une feuille Excel en ajoutant des en-têtes et en appliquant du style aux cellules.

### Étapes :

1. **Ouvrir** le fichier **vehicles.xlsx** dans Excel.
2. Aller dans l'onglet **Développeur**.
3. Dans le premier groupe, cliquer sur **Record Macro**.
4. Dans la boîte de dialogue :
   - Donner un nom explicite à la macro (ex. : `FormatVehicles`).
   - Choisir où enregistrer la macro (**This Workbook**).
   - Ajouter une description si nécessaire.
   - Cliquer sur **OK** pour commencer l'enregistrement.
5. **Insérer une ligne** en haut :
   - Clic droit sur la ligne **1** → **Insert**.
6. **Ajouter les en-têtes** en remplissant les cellules suivantes (utiliser la touche **TAB** pour passer à la colonne suivante) :
   - **A1** : VIN
   - **B1** : Année
   - **C1** : Marque
   - **D1** : Modèle
   - **E1** : Classification
   - **F1** : Couleur
   - **G1** : Coût
   - **H1** : Prix Manufacture
7. **Appliquer le formatage** :
   - Sélectionner la ligne **1**.
   - Aller dans l'onglet **Accueil**.
   - Mettre le texte en **gras** et **centré**.
   - Sélectionner la colonne **G** et appliquer le format **monétaire (Dollar)**.
   - Sélectionner les colonnes **A à H**, puis ajuster leur largeur en cliquant sur **AutoFit Column Width** dans le groupe **Cells**.
8. **Finaliser la macro** :
   - Avant d'arrêter l'enregistrement, cliquer sur une cellule vide.
   - Retourner dans l'onglet **Développeur** et cliquer sur **Stop Recording**.

### Résultat attendu :

La feuille de calcul doit contenir les en-têtes correctement formatés, et la macro pourra être rejouée pour formater une autre feuille similaire.

⚠️ **Pour enregistrer le fichier avec les macros, il est important de choisir le type de fichier adéquat.**  
En effet, les macros VBA ne sont pas compatibles avec le format de fichier Excel standard (.xlsx). Il est donc nécessaire de sauvegarder le fichier avec l'extension **.xlsm** (Classeur Excel prenant en charge les macros). Ce format permet de conserver et d'exécuter les macros que vous avez créées, contrairement au format .xlsx, qui ne les enregistre pas.


## Tester la macro sur un autre fichier

1. **Copier** le fichier original **vehicles.xlsx** et le renommer (ex. : **vehicles_test.xlsx**).
2. **Ouvrir** ce nouveau fichier dans Excel.
3. **Accéder à l'onglet Développeur**.
4. **Cliquer sur le bouton "Macros"** pour afficher la liste des macros enregistrées.
5. **Sélectionner la macro précédemment créée** et cliquer sur **Exécuter**.
6. Vérifier que :
   - La ligne d'en-tête a bien été insérée et formatée correctement.
   - Les colonnes ont la bonne largeur.
   - Les valeurs monétaires sont bien en dollars.
7. **Corriger la macro si nécessaire** via l'éditeur VBA.

## Exercice : (Fichier **vehicles.xlsx**)

---

Ce premier module pose les bases du VBA. La prochaine session abordera les structures de contrôle et la manipulation des cellules avec VBA.
