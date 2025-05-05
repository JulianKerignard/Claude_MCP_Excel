# MCP pour Claude Desktop 💻

Ce document présente le MCP (Multi-Component Plugin) développé pour interagir avec Claude via l'application desktop.

## 1. Présentation 📜

Le MCP permet :

* Interagir avec Claude depuis l'app desktop.
* Générer des comptes-rendus textuels ✍️.
* Produire des graphiques en React 📊 (future prise en charge d'images directement modifiables 🖼️).
* Optimiser les requêtes grâce à la bibliothèque Pandas 🐼.

**Note** : Le MCP fonctionne *uniquement* avec Claude Desktop. L'API Claude n'est pas (encore) supportée. 🚫

## 2. Fonctionnalités ✨

* **Comptes-rendus**
    * Résumé de texte, analyses, synthèses.
* **Graphiques**
    * Composants React intégrés.
    * *Futur* : réception de graphiques sous forme d'images et édition directe dans l'app.
* **Optimisation des données**
    * Utilisation de Pandas pour le traitement et la préparation des données.

## 3. Installation et configuration 🛠️

1.  Clonez le dépôt dans votre répertoire de travail.
2.  Assurez-vous d'avoir Node.js, Python 3.x, et Pandas installés.
3.  Dans le dossier `%APPDATA%\Claude\` (ou le dossier équivalent sur macOS), créez un fichier nommé `claude_desktop_config.json` avec le contenu suivant :

    ```json
    {
      "mcpServers": {
        "excel-viz": {
          "command": "python",
          "args": ["CHEMIN\\VERS\\VOTRE\\excel-mcp-server\\py\\excel_viz_server.py"] // Adaptez ce chemin !
        }
      }
    }
    ```

**Note** :

* Le MCP est lancé *uniquement* depuis l'application Claude Desktop.
* L'API Claude n'est pas (encore) supportée pour ce plugin.

## 4. Utilisation ▶️

1.  Ouvrez l'application Claude Desktop (Windows ou macOS).
2.  Dans la barre latérale ou le menu des plugins, sélectionnez `excel-viz` (ou tout autre serveur configuré dans `mcpServers`).
3.  Envoyez vos requêtes directement depuis l’interface intégrée du MCP : comptes-rendus, graphiques, etc. 🚀
4.  Le MCP se charge automatiquement au démarrage de Claude Desktop, sans commande à lancer manuellement. ✅

## 5. Exemple de prompt 💡

### 1. Lister les feuilles d'un fichier Excel
```text
Peux-tu me lister toutes les feuilles du fichier Excel "budget.xlsx" dans mon dossier Documents ?
```
### 2. Lire le contenu d'une feuille
```text
Montre-moi le contenu de la feuille "Revenus" dans le fichier "finances2024.xlsx"
```
### 3. Obtenir un résumé statistique
```text
Génère un résumé statistique des données dans la feuille "Ventes" du fichier "rapport_trimestriel.xlsx"
```
### 4. Exécuter une requête sur les données
```text
Dans le fichier "employés.xlsx", trouve toutes les lignes où la colonne "Salaire" est supérieure à 50000
```

## 6. Points importants à retenir !!

### Chemins de fichiers : Vous pouvez utiliser des chemins relatifs (par rapport à votre dossier Documents par défaut) ou des chemins absolus.
```text
// Chemin relatif (à partir du dossier Documents)
budget.xlsx

// Sous-dossier dans Documents
Finances/budget.xlsx

// Chemin absolu
D:/MesDonnées/Excel/budget.xlsx
```
### Syntaxe des requêtes : Pour la fonction excel_query, utilisez la syntaxe similaire à celle de pandas :
```text
// Exemples de requêtes valides
"Âge > 30"
"Département == 'Marketing'"
"Ventes > 1000 and Région == 'Nord'"
```

### Nom des feuilles : Si vous ne spécifiez pas de nom de feuille, le serveur utilisera la première feuille du classeur par défaut.
### Grandes quantités de données : Si votre fichier Excel contient beaucoup de données, pensez à spécifier la feuille et à utiliser des requêtes ciblées pour éviter d'afficher des tableaux trop volumineux.

## 7. Dépannage courant
### Si Claude vous indique qu'il a des difficultés à accéder au fichier, vérifiez :

Que le chemin du fichier est correct
Que le fichier n'est pas ouvert dans Excel (ce qui peut parfois bloquer l'accès)
Que le nom de la feuille est bien orthographié et existe dans le fichier

