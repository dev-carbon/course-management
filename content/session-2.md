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
