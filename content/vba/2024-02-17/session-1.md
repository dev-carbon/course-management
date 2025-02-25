---
title: Introduction à VBA
students:
  - Ceren KARAYILAN
files:
  - 1. Font Styles.xlsx
  - 2. Debugging.xlsm
  - 3. Creating Macros from Scratch.xlsm
  - 4. Insert & Format Text.xlsx
---

# Introduction à VBA

## Qu'est-ce que VBA ?

- VBA (Visual Basic for Applications) est un langage de programmation orienté objet intégré à Microsoft Office.
- VBA est principalement utilisé pour automatiser des tâches répétitives dans Excel, Word et Access.
- VBA permet de créer des macros, d'interagir avec les objets des applications Office et d'améliorer l'efficacité des processus de travail.

## Pourquoi utiliser VBA ?

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

- **Macro** : Séquence d'instructions enregistrée dans VBA permettant d'automatiser des tâches répétitives.
- **Procédure** : Bloc de code VBA exécutant une tâche spécifique. Peut être de type **Sub** (ne retourne pas de valeur) ou **Function** (retourne une valeur).
- **Module** : Conteneur de code VBA où sont stockées les procédures et fonctions. Il peut être standard (**Module**) ou de classe (**Class Module**).
- **Objet** : Entité manipulable en VBA, comme une feuille de calcul, une cellule ou un graphique. Chaque objet possède des propriétés et des méthodes.
- **Propriété** : Caractéristique d’un objet (ex. : `Range("A1").Value` donne la valeur de la cellule A1).
- **Méthode** : Action qu'un objet peut exécuter (ex. : `Range("A1").Select` sélectionne la cellule A1).
- **Workbook (Classeur)** : Représente un fichier Excel contenant une ou plusieurs feuilles de calcul (**Sheets**). Exemple d'accès : `Workbooks("MonFichier.xlsx")`.
- **Worksheet (Feuille de calcul)** : Feuille spécifique d'un classeur. Exemple d'accès : `Worksheets("Feuil1")` ou `ActiveSheet` pour la feuille active.
- **Collection** : Ensemble d'objets similaires regroupés. Par exemple, **Workbooks** est une collection contenant tous les classeurs ouverts, et **Worksheets** est une collection contenant toutes les feuilles d'un classeur.
- **Range** : Représente une ou plusieurs cellules dans une feuille de calcul. Exemple : `Range("A1")` fait référence à la cellule A1, tandis que `Range("A1:B2")` sélectionne un ensemble de cellules.

## Premiers pas avec VBA

### Préparation de l'environnement

Avant de commencer, il est important d'ajouter l'onglet **Développeur** à l'interface Excel :

1. Ouvrir Excel et cliquer sur **Fichier** > **Options**.
2. Dans la fenêtre **Options Excel**, aller dans **Personnaliser le ruban**.
3. Dans la colonne de droite, cocher **Développeur**, puis cliquer sur **OK**.
4. L'onglet **Développeur** est maintenant visible dans le ruban.

### Accéder à l'éditeur VBA

- L'éditeur VBA, est aussi appelé **VBE** pour (_Visual Basic Editor_).
- C'est l'environnement dans lequel on écrit et exécute du code VBA.
- Appuyer sur **ALT + F11** pour ouvrir directement l'éditeur VBA.

#### Accès via l'interface

1. Aller dans l'onglet **Développeur**.
2. Cliquer sur **Visual Basic** pour ouvrir l'éditeur.

#### Explorer l'interface du VBE

- **Project Explorer** : Affiche la structure du projet, incluant les modules, feuilles et objets.
- **Properties Window** : Permet de modifier les propriétés des objets sélectionnés.
- **Immediate Window** : Console permettant d’exécuter du code et tester des commandes rapidement.
  - **Accès** : Si elle n'est pas visible, activer la fenêtre via **Affichage** → **Fenêtre Exécution** ou en appuyant sur `Ctrl + G`.

#### Personnalisation de la barre d'outils

Il est recommandé d’ajouter des boutons utiles pour optimiser le travail :

- **Compile Project** : Vérifie la syntaxe et signale les erreurs.
- **Comment Block** : Ajoute des commentaires à plusieurs lignes de code.
- **UnComment Block** : Supprime les commentaires d'un bloc de code.
- **Step Into (F8)** : Exécute le code ligne par ligne pour suivre son déroulement.

### Cas pratiques

---

#### Exemple : enregistre ta premiere macro

### Différence entre **Absolute Cell Reference** et **Relative Cell Reference** :

- **Absolute Cell Reference (Référence absolue)** : Les cellules sont enregistrées avec des références fixes (ex. `$A$1`). Lorsqu'on exécute la macro, elle applique les actions **exactement sur ces cellules**, quel que soit l'emplacement du curseur.
- **Relative Cell Reference (Référence relative)** : Les cellules sont enregistrées en fonction de la position actuelle (ex. `A1` sans `$`). Lorsqu'on exécute la macro, elle applique les actions **par rapport à la cellule active**, ce qui la rend plus flexible.

⚠️ **Lors de l'enregistrement d'une macro, il est possible de choisir entre mode absolu et mode relatif via l'option "Référence relative" dans l'onglet Développeur.**

1. Aller dans l'onglet **Développeur** et cliquer sur **Enregistrer une macro**.
2. Dans la fenêtre qui s'affiche :
   - Donner un nom à la macro, par exemple **"FormatCellule"**.
   - Choisir où la stocker (**Ce classeur** par défaut).
   - Optionnel : Ajouter un raccourci clavier (ex. `Ctrl + Shift + F`).
   - Cliquer sur **OK** pour commencer l'enregistrement.
3. Effectuer les actions suivantes dans Excel :
   - Sélectionner la cellule **A1**.
   - Saisir le texte **"Hello VBA"**.
   - Appliquer un format : **gras**, **italique**, et couleur de fond **jaune**.
4. Arrêter l'enregistrement en cliquant sur **Arrêter l'enregistrement** dans l'onglet **Développeur**.
5. Exécuter la macro via **Macros** → **Sélectionner "FormatCellule"** → **Exécuter**.

#### Analyse de l'anatomie du code VBA

Après l'enregistrement, on peut voir le code généré dans l'editeur VBA:

1. Ouvrir l'éditeur VBA (`ALT + F11`).
2. Aller dans **Modules** → **Module1** (ou un autre module créé).

- Le Module est cree directement apres l enregistrement de la macro

3. Voici le code enregistré :

### Introduction Fenêtre Immediate en VBA

---

La **fenêtre Immediate** permet d'exécuter des instructions en temps réel, sans modifier le code principal de vos macros.  
Elle est particulièrement utile pour le débogage, la vérification rapide de valeurs, et l'exécution de commandes VBA ponctuelles.

#### Accès à la fenêtre Immediate

- Activez-la via **Affichage** → **Fenêtre Exécution** ou utilisez le raccourci clavier `Ctrl + G`.

#### Pratiquons quelques commandes dans la fenêtre Immediate

### 1. Sélection de cellule ou plage de cellules

- **Utilisation de `Range` :** Sélectionne une cellule ou une plage définie par son adresse.

  ```vba
  Range("B2").Select
  Range("A1:C3").Select
  ```

- **Utilisation de `Range.Range` :** Sélectionne une sous-plage à l'intérieur d'une plage définie.

  ```vba
  Range("A1:C5").Range("B2").Select 'Sélectionne B2 dans A1:C5
  ```

- **Utilisation de `Range.Offset` :** Déplace la sélection par rapport à une cellule de référence.

  ```vba
  Range("B2").Offset(1, 2).Select 'Sélectionne la cellule située 1 ligne en dessous et 2 colonnes à droite de B2
  ```

- **Utilisation de `Cells` :** Sélectionne une cellule par ses coordonnées numériques (ligne, colonne).

  ```vba
  Cells(3, 2).Select 'Sélectionne la cellule B3
  ```

- **Utilisation de `[]` :** Alternative rapide à `Range`.

  ```vba
  [D5].Select
  ```

- **Utilisation de `Range(variable)` :** Permet d'utiliser des variables dynamiquement.

  ```vba
  Dim maPlage As String
  maPlage = "A1:D4"
  Range(maPlage).Select
  ```

### 2. L'objet `Selection` et Couleurs

- Représente la plage actuellement sélectionnée.
- Utile pour appliquer des actions à la sélection :

  ```vba
  Selection.Interior.Color = RGB(255, 255, 0) 'Applique une couleur jaune à la sélection
  Selection.Font.Bold = True 'Met le texte en gras
  ```

### 3. `Value` et `Clear`

- **`Value` :** Permet de lire ou modifier le contenu d'une cellule ou plage.

  ```vba
  Range("A1").Value = "Hello VBA"
  Debug.Print Range("A1").Value 'Affiche la valeur de A1 dans la fenêtre Immediate
  ```

- **`Clear` :** Efface le contenu (et éventuellement le format) d'une cellule ou plage.

  ```vba
  Range("A1:A5").Clear 'Efface tout (valeur et format)
  Range("A1:A5").ClearContents 'Efface uniquement la valeur
  ```

### 4. `ActiveSheet`, `Sheets` et `Name`

- **`ActiveSheet` :** Feuille active du classeur.

  ```vba
  Debug.Print ActiveSheet.Name 'Affiche le nom de la feuille active
  ```

- **`Sheets` :** Collection de toutes les feuilles du classeur.

  ```vba
  Sheets(1).Select 'Sélectionne la première feuille
  Sheets("Feuil2").Select 'Sélectionne la feuille nommée "Feuil2"
  ```

- **`Name` :** Permet de lire ou modifier le nom d'une feuille.

  ```vba
  ActiveSheet.Name = "NouvelleFeuille" 'Renomme la feuille active
  ```

### 5. `CurrentRegion`

- Définit la plage de cellules contiguë contenant la cellule spécifiée.
- Utile pour travailler sur des tableaux de données.

  ```vba
  Range("B2").CurrentRegion.Select 'Sélectionne toute la région autour de B2
  Debug.Print Range("B2").CurrentRegion.Address 'Affiche l'adresse de la région
  ```

### 6. Interroger la feuille de calcul en temps reel

- **1️⃣ Récupérer le nom de la feuille active**

```vba
? ActiveSheet.Name
```

- **Lire la valeur d'une cellule spécifique**

```vba
? Range("A1").Value
```

- **Lire la valeur de la cellule active**

```vba
? ActiveCell.Value
```

- **Vérifier si une cellule est vide**

```vba
? IIf(IsEmpty(Range("A1").Value), "La cellule A1 est vide", "La cellule A1 contient : " & Range("A1").Value)
```

- **Vérifier si une cellule contient une formule**

```vba
? IIf(Range("B2").HasFormula, "La cellule B2 contient une formule", "La cellule B2 ne contient pas de formule")
```

# Exercices Pratiques : Fenêtre Immediate en VBA

---

## **Exercice 1 : Modification et mise en forme des cellules via la fenêtre Immediate**

💡 **Objectif :** Modifier rapidement les valeurs de certaines cellules et ajuster leur mise en forme.

### **Instructions :**

1. Ouvrez un classeur Excel avec une feuille vide.
2. Dans la **fenêtre Immediate**, saisissez et exécutez les commandes VBA suivantes pour remplir des cellules :

```vba
Range("A1").Value = "Nom"
Range("B1").Value = "Prénom"
Range("C1").Value = "Âge"
```

3. Ajustez automatiquement la largeur des colonnes :

```vba
Columns("A:C").AutoFit
```

4. Alignez le texte au centre horizontalement et verticalement :

```vba
  Range("A1:C1").HorizontalAlignment = xlCenter
  Range("A1:C1").VerticalAlignment = xlCenter
```

5. Changez la couleur du texte en blanc et appliquez un fond bleu :

```vba
Range("A1:C1").Interior.Color = RGB(0, 0, 255) ' Bleu
Range("A1:C1").Font.Color = RGB(255, 255, 255) ' Blanc
```

6. Vérifiez les valeurs des cellules en les affichant dans la fenêtre Immediate :

```vba
Debug.Print Range("A1").Value, Range("B1").Value, Range("C1").Value
```

7. Modifiez la valeur de la cellule C1 pour afficher un âge différent :

```vba
Range("C1").Value = 30 ' Modifier l'âge
Debug.Print "Nouvelle valeur de C1 :", Range("C1").Value
```

## Exercice 2 : Application avancée du formatage

💡 **Objectif** : Appliquer du formatage sur des cellules sélectionnées, ajuster la taille de police et l'alignement.

## **Instructions :**

1. Sélectionnez une plage de cellules (par exemple, A1:C3).

2. Dans la fenêtre Immediate, appliquez un fond jaune et mettez le texte en gras :

```vba
Selection.Interior.Color = RGB(255, 255, 0) ' Fond jaune
Selection.Font.Bold = True ' Texte en gras
Selection.Font.Size = 14 ' Taille de police augmentée
```

3. Alignez le texte au centre :

```vba
Selection.HorizontalAlignment = xlCenter
Selection.VerticalAlignment = xlCenter
```

4. Ajustez automatiquement la largeur des colonnes et la hauteur des lignes :

```vba
Selection.EntireColumn.AutoFit
Selection.EntireRow.AutoFit
```

5. Vérifiez que les modifications sont bien appliquées.

6. Effacez uniquement la mise en forme (sans supprimer le contenu) avec :

```vba
Selection.ClearFormats
```

## Exercice 3 : Manipulation des feuilles et mise en forme avancée

💡 **Objectif** : Créer, renommer, formater et naviguer entre des feuilles.

## **Instructions :**

1. Dans la fenêtre Immediate, créez une nouvelle feuille et renommez-la en "Données" :

```vba
Sheets.Add(After:=Sheets(Sheets.Count)).Name = "Données"
```

2. Sélectionnez cette feuille nouvellement créée :

```vba
Sheets("Données").Select
```

3. Remplissez les cellules A1:C1 avec des titres :

```vba
Range("A1").Value = "ID"
Range("B1").Value = "Nom"
Range("C1").Value = "Score"
```

4. Appliquez un style de tableau :

```vba
Range("A1:C1").Font.Bold = True
Range("A1:C1").Interior.Color = RGB(0, 128, 0) ' Vert
Range("A1:C1").Font.Color = RGB(255, 255, 255) ' Blanc
```

5. Ajustez la largeur des colonnes :

```vba
Columns("A:C").AutoFit
```

6. Affichez dans la fenêtre Immediate le nombre total de feuilles présentes dans le classeur :

```vba
Debug.Print "Nombre de feuilles : " & Sheets.Count
```

7. Supprimez la feuille si nécessaire :

```vba
Application.DisplayAlerts = False
Sheets("Données").Delete
Application.DisplayAlerts = True
```
