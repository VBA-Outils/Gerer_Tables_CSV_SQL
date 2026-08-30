# 📊 Gestion avancée de tables dans des tableaux structurés Excel (paramétrage, référentiels, clients, etc)

![Langage](https://img.shields.io/badge/langage-VBA-blue)
![Excel](https://img.shields.io/badge/Excel-Tables%20structurées-green)
![Licence](https://img.shields.io/badge/Licence-MIT-yellow)

Gerer-Tables-Parametrage est un outil Excel/VBA permettant de gérer des **tables structurées** réparties par onglet, avec un **contrat d’interface centralisé**, et des fonctions d’**import/export CSV** ainsi que la génération de **requêtes SQL `INSERT INTO`**.

Ce projet vise à offrir une manière fiable, cohérente et automatisée de manipuler des données tabulaires dans Excel.

Ce fichier peut être utilisé pour gérer des tables de paramétrages d'un projet avec la possibilité d'importer depuis des fichiers CSV le contenu des tables, de les exporter, de comparer un fichier CSV avec le contenu d'une table du fichier Excel.

---

## 🧩 Fonctionnalités principales

### ✔️ Création et mise-à-jour des tables automatisée à partir d'un contrat d'interface
- Création automatique des tables selon un contrat d’interface qui définit les types de données à l'instar de PostGreSql.
- Vérification dynamique des saisies du contrat d'interface lors de la saisie.

### ✔️ Importer un fichier CSV dans une table Excel
- Lecture d’un fichier CSV avec gestion des séparateurs de lignes, séparateurs CSV, encodage du fichier, présence d'en-tête, etc.
- Vérification de la conformité avec le contrat d'interface
- Insertion des données dans la table cible par annule et remplace.
- Rapport d’erreurs (colonnes manquantes, formats incorrects)

### ✔️ Exporter une table vers un fichier CSV
- Export propre et conforme au contrat d'interface
- Gestion des séparateurs (point-virgule, tabulation, etc), des délimiteurs de texte, de l'encodage, des délimiteurs de lignes (CR/LF)
- Option d’inclure ou non les en‑têtes

### ✔️ Exporter une table Excel sous forme de requête SQL "INSERT INTO"
- Génération automatique des requêtes
- Échappement des valeurs
- Gestion des types (texte, numérique, date)
- Export dans un fichier `.sql`

### ✔️ Comparer une table entre le fichier Excel et un fichier CSV
- Charger un fichier CSV dans un nouveau classeur
- Comparer le contenu des tables
- Rapport de comparaison (lignes en différence, écarts des valeurs dans une colonne)
- Réactualiser le contenu de la table du 1er fichier à partir du 2ème fichier (au choix : lignes ajoutées, lignes supprimées, valeurs différentes) avec colorisation possible des écarts.

### ✔️ Vérifier la saisie d'un table en fonction de l'exportation (CSV ou SQL)
- Contrôler le format des données, leur valeur par rapport au type de la donnée.
- Détecter les doublons dans les clés primaires.
- Contrôler que les clés étrangères existent.

### ✔️ Dupliquer un modèle de données entre 2 classeurs
- Cette option permet de gérer les montées de version des macros VBA. Au lieu d'insérer toutes les modifications à la main dans les fichiers existants, le modèle de données du classeur avec une ancienne version des macros peut être entièrement recopié dans le classeur avec la nouvelle version des macros.
- Recopier le contrat d'interface du classeur 1 vers le classeur 2.
- Créer les nouvelles tables dans le classeur 2.
- Recopier le contenu des tables.

---

## 📁 Structure du classeur

Le classeur contient plusieurs onglets, chacun dédié à une table structurée.  
Un onglet particulier, **`Contrat d'interface`**, décrit :

- le nom de chaque table,
- les colonnes,
- leur ordre,
- leur type ou format,
- les règles de validation.

Ce contrat sert de **référence unique** pour toutes les opérations d’import/export.

---

## 🚀 Démarrage rapide

### 1. Télécharger le fichier Excel  
Récupérez `GererTables.xlsm` depuis le dépôt GitHub.

### 2. Activer les macros  
Excel → Options → Centre de gestion de la confidentialité → Paramètres des macros.

### 3. Définir vos tables  
Dans l’onglet **`Contrat d'interface`**, définissez les colonnes et formats. Puis créez les feuilles qui vont contenir les données.

### 4. Importer un CSV  
Menu → `Importer CSV` → Sélectionner le fichier → Choisir la table cible.

### 5. Exporter un CSV  
Menu → `Exporter CSV` → Choisir la table → Sélectionner l’emplacement.

### 6. Générer du SQL  
Menu → `Exporter SQL` → Un fichier `.sql` est généré.

---

## 🧪 Exemples

### Table structurée

| ID | Nom | Age |
| --- | --- | --- |
| 1 | Dupont | 32 |
| 2 | Martin | 45 |


### SQL généré

INSERT INTO Clients (ID, Nom, Age) VALUES 
(1, 'Dupont', 32),
(2, 'Martin', 45);

---

## 📘 Mode d’emploi — Création et utilisation des tables

### 1️⃣ Préparer l’onglet **Contrat d’interface**

L’onglet *Contrat d’interface* constitue le schéma fonctionnel du classeur.  
Chaque ligne décrit une table ou une colonne.

Renseigner pour chaque entrée :

- **Nom de la table**
- **Nom de la colonne**
- **Format / Type de données** (Texte, Entier, Date, etc.)
- **Clé primaire** (si applicable)
- **Références** (liens vers d’autres tables)

Cet onglet sert de base à la génération automatique des tables.

---

### 2️⃣ Générer les tables

Une fois le contrat d’interface complété :

1. Cliquer sur le bouton **Réactualiser / créer des feuilles avec les tables**.
2. Le programme :
   - crée un **onglet par table** définie ;
   - insère un **tableau structuré** avec les colonnes du contrat ;
   - ajoute automatiquement sur la première ligne :
     - **Menu**
     - **Exporter CSV**
     - **Importer CSV**
     - **Exporter SQL**
   - inscrit le **nom de la table** en haut de l’onglet.

Les tables sont immédiatement opérationnelles pour la saisie, l’import/export et la génération SQL.

---

### 3️⃣ Définir les relations entre tables

Après la création des tables :

- La colonne **Table référencée** du contrat d’interface propose une **liste déroulante** contenant toutes les tables du classeur.
- Lorsque l’utilisateur sélectionne une table :
  - la colonne **Clé primaire référencée** se remplit automatiquement avec une **liste déroulante contenant les clés primaires de la table choisie**.

Ce mécanisme permet de définir facilement :

- des **relations 1‑N** (clé primaire → clé étrangère),
- des **références croisées**,
- des **contraintes de cohérence** entre tables.

Le contrat d’interface devient ainsi un véritable *catalogue de métadonnées* pilotant la structure du classeur.

---

## 🧩 Résultat : un système de tables

Grâce à ce fonctionnement :

- la structure des tables est définie **une seule fois** dans le contrat d’interface,
- les onglets sont générés automatiquement et restent synchronisés,
- les relations entre tables sont gérées via des listes déroulantes dynamiques,
- l’utilisateur dispose d’un environnement cohérent pour :
  - saisir des données,
  - importer/exporter en CSV,
  - générer des requêtes SQL,
  - maintenir des relations entre tables.

