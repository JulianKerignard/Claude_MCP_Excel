from mcp.server.fastmcp import FastMCP
import pandas as pd
import os
import json
import base64
import io
from datetime import datetime
import matplotlib.pyplot as plt
import matplotlib

matplotlib.use('Agg')  # Nécessaire pour le mode non-interactif

# Initialiser le serveur MCP
mcp = FastMCP("excel-viz-server")

# Chemin par défaut pour les fichiers Excel
DEFAULT_EXCEL_DIR = os.path.expanduser("~\\Documents")


@mcp.tool()
def read_excel(file_path: str, sheet_name: str = None) -> str:
    """
    Lit un fichier Excel et retourne son contenu sous forme de texte.

    Args:
        file_path: Chemin du fichier Excel (peut être relatif au répertoire par défaut)
        sheet_name: Nom de la feuille à lire (si None, lit la première feuille)

    Returns:
        Contenu du fichier Excel sous forme tabulaire
    """
    try:
        # Gérer les chemins relatifs
        if not os.path.isabs(file_path):
            file_path = os.path.join(DEFAULT_EXCEL_DIR, file_path)

        # Vérifier que le fichier existe
        if not os.path.exists(file_path):
            return f"Erreur: Le fichier '{file_path}' n'existe pas."

        # Lire le fichier Excel avec pandas
        if sheet_name:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
        else:
            df = pd.read_excel(file_path)

        # Convertir en chaîne formatée
        result = df.to_string(index=False)

        # Ajouter des informations sur le fichier
        header = f"Fichier: {os.path.basename(file_path)}\n"
        header += f"Feuille: {sheet_name or 'par défaut'}\n"
        header += f"Dimensions: {df.shape[0]} lignes × {df.shape[1]} colonnes\n\n"

        return header + result
    except Exception as e:
        return f"Erreur lors de la lecture du fichier Excel: {str(e)}"


@mcp.tool()
def get_excel_sheets(file_path: str) -> str:
    """
    Liste toutes les feuilles dans un fichier Excel.

    Args:
        file_path: Chemin du fichier Excel (peut être relatif au répertoire par défaut)

    Returns:
        Liste des noms de feuilles dans le fichier Excel
    """
    try:
        # Gérer les chemins relatifs
        if not os.path.isabs(file_path):
            file_path = os.path.join(DEFAULT_EXCEL_DIR, file_path)

        # Vérifier que le fichier existe
        if not os.path.exists(file_path):
            return f"Erreur: Le fichier '{file_path}' n'existe pas."

        # Lire les noms de feuilles
        xls = pd.ExcelFile(file_path)
        sheets = xls.sheet_names

        # Formater la réponse
        result = f"Feuilles dans '{os.path.basename(file_path)}':\n\n"
        for i, sheet in enumerate(sheets, 1):
            result += f"{i}. {sheet}\n"

        return result
    except Exception as e:
        return f"Erreur lors de la lecture des feuilles: {str(e)}"


@mcp.tool()
def excel_summary(file_path: str, sheet_name: str = None) -> str:
    """
    Génère un résumé statistique des données dans un fichier Excel.

    Args:
        file_path: Chemin du fichier Excel (peut être relatif au répertoire par défaut)
        sheet_name: Nom de la feuille à analyser (si None, utilise la première feuille)

    Returns:
        Résumé statistique des données
    """
    try:
        # Gérer les chemins relatifs
        if not os.path.isabs(file_path):
            file_path = os.path.join(DEFAULT_EXCEL_DIR, file_path)

        # Vérifier que le fichier existe
        if not os.path.exists(file_path):
            return f"Erreur: Le fichier '{file_path}' n'existe pas."

        # Lire le fichier Excel avec pandas
        if sheet_name:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
        else:
            df = pd.read_excel(file_path)

        # Générer résumé statistique
        summary = df.describe(include='all').fillna("N/A").to_string()

        # Informations supplémentaires
        info = f"Fichier: {os.path.basename(file_path)}\n"
        info += f"Feuille: {sheet_name or 'par défaut'}\n"
        info += f"Dimensions: {df.shape[0]} lignes × {df.shape[1]} colonnes\n"
        info += f"Colonnes: {', '.join(df.columns.tolist())}\n\n"

        return info + summary
    except Exception as e:
        return f"Erreur lors de la génération du résumé: {str(e)}"


@mcp.tool()
def excel_query(file_path: str, query: str, sheet_name: str = None) -> str:
    """
    Exécute une requête simple sur les données Excel.

    Args:
        file_path: Chemin du fichier Excel (peut être relatif au répertoire par défaut)
        query: Requête sous forme de texte (ex: "colonne > 100")
        sheet_name: Nom de la feuille à analyser (si None, utilise la première feuille)

    Returns:
        Résultat de la requête
    """
    try:
        # Gérer les chemins relatifs
        if not os.path.isabs(file_path):
            file_path = os.path.join(DEFAULT_EXCEL_DIR, file_path)

        # Vérifier que le fichier existe
        if not os.path.exists(file_path):
            return f"Erreur: Le fichier '{file_path}' n'existe pas."

        # Lire le fichier Excel avec pandas
        if sheet_name:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
        else:
            df = pd.read_excel(file_path)

        # Exécuter la requête
        try:
            result_df = df.query(query)

            # Générer la réponse
            header = f"Résultat de la requête '{query}':\n"
            header += f"Trouvé {len(result_df)} lignes sur {len(df)} totales\n\n"

            if len(result_df) > 0:
                return header + result_df.to_string(index=False)
            else:
                return header + "Aucun résultat trouvé."

        except Exception as query_error:
            return f"Erreur dans la requête: {str(query_error)}"

    except Exception as e:
        return f"Erreur lors de l'exécution de la requête: {str(e)}"


@mcp.tool()
def create_bar_chart(file_path: str, x_column: str, y_column: str, sheet_name: str = None,
                     title: str = "Graphique à barres") -> str:
    """
    Crée un graphique à barres à partir de données Excel.

    Args:
        file_path: Chemin du fichier Excel
        x_column: Nom de la colonne pour l'axe X
        y_column: Nom de la colonne pour l'axe Y
        sheet_name: Nom de la feuille (optionnel)
        title: Titre du graphique

    Returns:
        Image du graphique encodée en base64
    """
    try:
        # Gérer les chemins relatifs
        if not os.path.isabs(file_path):
            file_path = os.path.join(DEFAULT_EXCEL_DIR, file_path)

        # Vérifier que le fichier existe
        if not os.path.exists(file_path):
            return f"Erreur: Le fichier '{file_path}' n'existe pas."

        # Lire le fichier Excel
        if sheet_name:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
        else:
            df = pd.read_excel(file_path)

        # Vérifier que les colonnes existent
        if x_column not in df.columns:
            return f"Erreur: La colonne '{x_column}' n'existe pas dans le fichier."
        if y_column not in df.columns:
            return f"Erreur: La colonne '{y_column}' n'existe pas dans le fichier."

        # Créer le graphique
        plt.figure(figsize=(10, 6))
        plt.bar(df[x_column], df[y_column])
        plt.title(title)
        plt.xlabel(x_column)
        plt.ylabel(y_column)
        plt.xticks(rotation=45)
        plt.tight_layout()

        # Convertir le graphique en image base64
        buffer = io.BytesIO()
        plt.savefig(buffer, format='png')
        buffer.seek(0)
        image_base64 = base64.b64encode(buffer.read()).decode('utf-8')
        plt.close()

        return f"data:image/png;base64,{image_base64}"

    except Exception as e:
        return f"Erreur lors de la création du graphique: {str(e)}"


@mcp.tool()
def create_pie_chart(file_path: str, labels_column: str, values_column: str, sheet_name: str = None,
                     title: str = "Graphique circulaire") -> str:
    """
    Crée un graphique circulaire à partir de données Excel.

    Args:
        file_path: Chemin du fichier Excel
        labels_column: Nom de la colonne pour les étiquettes
        values_column: Nom de la colonne pour les valeurs
        sheet_name: Nom de la feuille (optionnel)
        title: Titre du graphique

    Returns:
        Description textuelle du graphique
    """
    try:
        # Gérer les chemins relatifs
        if not os.path.isabs(file_path):
            file_path = os.path.join(DEFAULT_EXCEL_DIR, file_path)

        # Lire le fichier Excel
        if sheet_name:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
        else:
            df = pd.read_excel(file_path)

        # Agréger les données si nécessaire
        aggregated_df = df.groupby(labels_column)[values_column].sum().reset_index()

        # Calculer les pourcentages
        total = aggregated_df[values_column].sum()
        percentages = [(value / total * 100) for value in aggregated_df[values_column]]

        # Créer une description textuelle du graphique
        result = f"Graphique circulaire: {title}\n\n"
        result += f"Distribution de {values_column} par {labels_column}:\n"

        for i, (label, value) in enumerate(zip(aggregated_df[labels_column], aggregated_df[values_column])):
            result += f"- {label}: {value:,.2f} ({percentages[i]:.1f}%)\n"

        return result

    except Exception as e:
        return f"Erreur lors de la création du graphique: {str(e)}"


@mcp.tool()
def create_line_chart(file_path: str, x_column: str, y_columns: str, sheet_name: str = None,
                      title: str = "Graphique linéaire") -> str:
    """
    Crée un graphique linéaire à partir de données Excel.

    Args:
        file_path: Chemin du fichier Excel
        x_column: Nom de la colonne pour l'axe X
        y_columns: Noms des colonnes pour l'axe Y (séparés par des virgules)
        sheet_name: Nom de la feuille (optionnel)
        title: Titre du graphique

    Returns:
        Image du graphique encodée en base64
    """
    try:
        # Gérer les chemins relatifs
        if not os.path.isabs(file_path):
            file_path = os.path.join(DEFAULT_EXCEL_DIR, file_path)

        # Vérifier que le fichier existe
        if not os.path.exists(file_path):
            return f"Erreur: Le fichier '{file_path}' n'existe pas."

        # Lire le fichier Excel
        if sheet_name:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
        else:
            df = pd.read_excel(file_path)

        # Vérifier que la colonne X existe
        if x_column not in df.columns:
            return f"Erreur: La colonne '{x_column}' n'existe pas dans le fichier."

        # Traiter les colonnes Y
        y_columns_list = [col.strip() for col in y_columns.split(',')]
        for col in y_columns_list:
            if col not in df.columns:
                return f"Erreur: La colonne '{col}' n'existe pas dans le fichier."

        # Créer le graphique
        plt.figure(figsize=(12, 7))
        for col in y_columns_list:
            plt.plot(df[x_column], df[col], marker='o', label=col)

        plt.title(title)
        plt.xlabel(x_column)
        plt.ylabel("Valeurs")
        plt.legend()
        plt.grid(True, linestyle='--', alpha=0.7)
        plt.xticks(rotation=45)
        plt.tight_layout()

        # Convertir le graphique en image base64
        buffer = io.BytesIO()
        plt.savefig(buffer, format='png')
        buffer.seek(0)
        image_base64 = base64.b64encode(buffer.read()).decode('utf-8')
        plt.close()

        return f"data:image/png;base64,{image_base64}"

    except Exception as e:
        return f"Erreur lors de la création du graphique: {str(e)}"


@mcp.tool()
def create_scatter_plot(file_path: str, x_column: str, y_column: str, color_column: str = None, sheet_name: str = None,
                        title: str = "Nuage de points") -> str:
    """
    Crée un nuage de points à partir de données Excel.

    Args:
        file_path: Chemin du fichier Excel
        x_column: Nom de la colonne pour l'axe X
        y_column: Nom de la colonne pour l'axe Y
        color_column: Nom de la colonne pour la couleur des points (optionnel)
        sheet_name: Nom de la feuille (optionnel)
        title: Titre du graphique

    Returns:
        Image du graphique encodée en base64
    """
    try:
        # Gérer les chemins relatifs
        if not os.path.isabs(file_path):
            file_path = os.path.join(DEFAULT_EXCEL_DIR, file_path)

        # Vérifier que le fichier existe
        if not os.path.exists(file_path):
            return f"Erreur: Le fichier '{file_path}' n'existe pas."

        # Lire le fichier Excel
        if sheet_name:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
        else:
            df = pd.read_excel(file_path)

        # Vérifier que les colonnes existent
        if x_column not in df.columns:
            return f"Erreur: La colonne '{x_column}' n'existe pas dans le fichier."
        if y_column not in df.columns:
            return f"Erreur: La colonne '{y_column}' n'existe pas dans le fichier."
        if color_column and color_column not in df.columns:
            return f"Erreur: La colonne '{color_column}' n'existe pas dans le fichier."

        # Créer le graphique
        plt.figure(figsize=(10, 8))

        if color_column:
            scatter = plt.scatter(df[x_column], df[y_column], c=df[color_column], cmap='viridis', alpha=0.6)
            plt.colorbar(scatter, label=color_column)
        else:
            plt.scatter(df[x_column], df[y_column], alpha=0.6)

        plt.title(title)
        plt.xlabel(x_column)
        plt.ylabel(y_column)
        plt.grid(True, linestyle='--', alpha=0.3)
        plt.tight_layout()

        # Convertir le graphique en image base64
        buffer = io.BytesIO()
        plt.savefig(buffer, format='png')
        buffer.seek(0)
        image_base64 = base64.b64encode(buffer.read()).decode('utf-8')
        plt.close()

        return f"data:image/png;base64,{image_base64}"

    except Exception as e:
        return f"Erreur lors de la création du graphique: {str(e)}"


@mcp.tool()
def get_column_names(file_path: str, sheet_name: str = None) -> str:
    """
    Retourne la liste des noms de colonnes d'un fichier Excel.

    Args:
        file_path: Chemin du fichier Excel
        sheet_name: Nom de la feuille (optionnel)

    Returns:
        Liste des noms de colonnes
    """
    try:
        # Gérer les chemins relatifs
        if not os.path.isabs(file_path):
            file_path = os.path.join(DEFAULT_EXCEL_DIR, file_path)

        # Vérifier que le fichier existe
        if not os.path.exists(file_path):
            return f"Erreur: Le fichier '{file_path}' n'existe pas."

        # Lire le fichier Excel
        if sheet_name:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
        else:
            df = pd.read_excel(file_path)

        # Obtenir et formater les noms de colonnes
        columns = df.columns.tolist()
        result = f"Colonnes dans '{os.path.basename(file_path)}'"
        if sheet_name:
            result += f" (feuille '{sheet_name}')"
        result += ":\n\n"

        for i, col in enumerate(columns, 1):
            result += f"{i}. {col}\n"

        return result

    except Exception as e:
        return f"Erreur lors de la récupération des noms de colonnes: {str(e)}"


@mcp.resource("excel://{file_path}")
def excel_resource(file_path: str) -> str:
    """
    Ressource pour accéder aux fichiers Excel.
    Exemple d'accès: excel://mon_fichier.xlsx
    """
    try:
        # Gérer les chemins relatifs
        if not os.path.isabs(file_path):
            file_path = os.path.join(DEFAULT_EXCEL_DIR, file_path)

        # Vérifier que le fichier existe
        if not os.path.exists(file_path):
            return f"Erreur: Le fichier '{file_path}' n'existe pas."

        # Lire le fichier Excel avec pandas
        excel_file = pd.ExcelFile(file_path)

        # Obtenir la structure du fichier
        result = {}

        # Obtenir les informations sur toutes les feuilles
        result["file_name"] = os.path.basename(file_path)
        result["sheets"] = excel_file.sheet_names
        result["file_size"] = f"{os.path.getsize(file_path) / 1024:.2f} KB"
        result["last_modified"] = datetime.fromtimestamp(os.path.getmtime(file_path)).strftime('%Y-%m-%d %H:%M:%S')

        # Obtenir des informations de base sur chaque feuille
        sheet_info = {}
        for sheet in excel_file.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet)
            sheet_info[sheet] = {
                "rows": len(df),
                "columns": len(df.columns),
                "column_names": df.columns.tolist()
            }

        result["sheet_details"] = sheet_info

        # Formater le résultat en JSON pour lisibilité
        return json.dumps(result, indent=2)

    except Exception as e:
        return f"Erreur lors de l'accès au fichier Excel: {str(e)}"


if __name__ == "__main__":
    mcp.run()