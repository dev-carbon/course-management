---
title: Programmation VBA
students:
  - Ceren KARAYILAN
files:
  - 6. For Next Loop.xlsm
  - 7. Do Loop.xlsm
  - 8. Worksheets, Count, Offset, Copy.xlsm
  - 9. Duplicating Data.xlsm
  - 10. Address, Call, End.xlsm
  - 11. Report Generator.xlsm
---

# Variables

En VBA, une variable est un espace mémoire permettant de stocker des données. 

## Déclaration des variables
On utilise l'instruction `Dim` pour déclarer une variable :

```vba
Dim nomVariable As Type
```

Exemple :

```vba
Dim age As Integer
Dim nom As String
Dim estValide As Boolean
```

## Types de variables
- `Integer` : nombres entiers (-32,768 à 32,767)
- `Long` : nombres entiers longs (-2,147,483,648 à 2,147,483,647)
- `Single` : nombres décimaux (précision simple)
- `Double` : nombres décimaux (précision double)
- `String` : chaînes de caractères
- `Boolean` : valeurs `True` ou `False`

## Option Explicit

L'instruction `Option Explicit` force la déclaration explicite de toutes les variables utilisées dans un module VBA. Elle doit être placée en haut du module :

```vba
Option Explicit
```

### Pourquoi utiliser Option Explicit ?
- **Évite les erreurs de typographie** : Si une variable mal orthographiée est utilisée, une erreur sera levée.
- **Améliore la lisibilité et la maintenance du code** : Toutes les variables doivent être déclarées explicitement, ce qui clarifie le code.
- **Optimise l'utilisation de la mémoire** : Empêche la création implicite de variables de type `Variant`, qui sont plus gourmandes en ressources.

Exemple sans `Option Explicit` (peut causer des erreurs invisibles) :

```vba
Dim valeur As Integer
valeur = 10
valer = 20 ' Erreur de frappe, mais VBA ne lève pas d'erreur sans Option Explicit
MsgBox valeur ' Affiche 10 au lieu de 20
```

Exemple avec `Option Explicit` (génère une erreur détectable) :

```vba
Option Explicit
Dim valeur As Integer
valeur = 10
valer = 20 ' VBA affichera une erreur "Variable non définie"
```

# Conditions

Les conditions permettent d'exécuter du code selon des critères spécifiques.

## If...Then...Else

```vba
If condition Then
    ' Instructions
ElseIf autreCondition Then
    ' Autres instructions
Else
    ' Instructions par défaut
End If
```

Exemple :

```vba
If age >= 18 Then
    MsgBox "Vous êtes majeur."
Else
    MsgBox "Vous êtes mineur."
End If
```

## Select Case

```vba
Select Case variable
    Case valeur1
        ' Instructions
    Case valeur2
        ' Instructions
    Case Else
        ' Instructions par défaut
End Select
```

Exemple :

```vba
Select Case jour
    Case "Lundi"
        MsgBox "Début de semaine"
    Case "Vendredi"
        MsgBox "Fin de semaine"
    Case Else
        MsgBox "Jour normal"
End Select
```

# Boucles

Les boucles permettent d'exécuter un bloc de code plusieurs fois.

## For...Next

```vba
For i = 1 To 10
    MsgBox "Iteration " & i
Next i
```

## For Each

```vba
Dim ws As Worksheet
For Each ws In Worksheets
    MsgBox ws.Name
Next ws
```

## Do While

```vba
Dim compteur As Integer
compteur = 0
Do While compteur < 5
    MsgBox "Compteur: " & compteur
    compteur = compteur + 1
Loop
```

## Do Until

```vba
Dim x As Integer
x = 1
Do Until x > 10
    MsgBox "Valeur de x: " & x
    x = x + 1
Loop
```

# Instruction End, Call, Address

## End

La méthode `.End` permet de sélectionner la dernière cellule non vide dans une direction donnée.

```vba
Dim lastCell As Range
Set lastCell = Range("A1").End(xlDown)
lastCell.Select
```

## Call

Permet d'appeler une procédure :

```vba
Call MaProcedure
```

## Address

La méthode `.Address` permet de récupérer l'adresse d'une cellule.

```vba
Dim cell As Range
Set cell = Range("B2")
MsgBox cell.Address
```

# Pratique

## Exercice 1: For Next Loop
Écrire une macro qui affiche les nombres de 1 à 10 dans des MsgBox.

## Exercice 2: For Next Loop Doubled
Modifier l'exercice précédent pour afficher les nombres pairs de 2 à 20.

## Exercice 3: For Next Loop Triple
Afficher les multiples de 3 jusqu'à 30.

## Exercice 4: For Each Loop
Lister tous les noms des feuilles d'un classeur.

## Exercice 5: Exit For Statement
Interrompre une boucle For lorsque la valeur atteint 5.

## Exercice 6: Do While Loop
Créer une boucle qui compte jusqu'à 10 et affiche les valeurs.

## Exercice 7: Do While Not Empty Loop
Récupérer les valeurs d'une colonne tant qu'une cellule n'est pas vide.

## Exercice 8: Do Until Loop
Afficher un compteur jusqu'à ce qu'il atteigne 15.

## Exercice 9: Do Loop Until
Créer une boucle qui continue tant qu'une condition n'est pas remplie.
