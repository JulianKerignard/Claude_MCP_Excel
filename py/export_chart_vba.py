from mcp.server.fastmcp import FastMCP
import os

# Initialiser le serveur MCP
mcp = FastMCP("excel-chart-export")

@mcp.tool()
def generate_chart_export_vba(chart_sheet_name: str = "Graphique1", 
                              chart_title: str = "Mon Graphique") -> str:
    """
    Génère le code VBA pour exporter un graphique Excel vers PowerPoint/Word.
    L'utilisateur doit manuellement ajouter ce code à son fichier Excel.
    
    Args:
        chart_sheet_name: Nom de la feuille contenant le graphique
        chart_title: Titre du graphique à utiliser dans PowerPoint
        
    Returns:
        Code VBA prêt à être copié-collé dans un module VBA Excel
    """
    
    vba_code = f"""
' Code VBA pour l'exportation de graphiques Excel vers PowerPoint et Word
' À copier-coller dans un nouveau module VBA (ALT+F11 pour ouvrir l'éditeur VBA)

Option Explicit

' Exporter le graphique vers PowerPoint
Sub ExportChartToPowerPoint()
    Dim pptApp As Object
    Dim pptPres As Object
    Dim pptSlide As Object
    Dim chartObj As ChartObject
    Dim sheetName As String
    
    ' Nom de la feuille contenant le graphique
    sheetName = "{chart_sheet_name}"
    
    On Error Resume Next
    
    ' Créer une instance PowerPoint
    Set pptApp = CreateObject("PowerPoint.Application")
    If Err.Number <> 0 Then
        MsgBox "PowerPoint n'est pas disponible sur cet ordinateur.", vbExclamation
        Exit Sub
    End If
    
    ' Rendre PowerPoint visible
    pptApp.Visible = True
    
    ' Créer une nouvelle présentation
    Set pptPres = pptApp.Presentations.Add
    
    ' Ajouter une diapositive
    Set pptSlide = pptPres.Slides.Add(1, 11) ' 11 = ppLayoutTitleOnly
    
    ' Définir le titre
    pptSlide.Shapes.Title.TextFrame.TextRange.Text = "{chart_title}"
    
    ' Vérifier que la feuille existe
    On Error Resume Next
    Set chartObj = ThisWorkbook.Sheets(sheetName).ChartObjects(1)
    If Err.Number <> 0 Then
        MsgBox "Feuille ou graphique introuvable: " & sheetName, vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Copier le graphique
    chartObj.Chart.CopyPicture
    
    ' Coller dans PowerPoint
    pptSlide.Shapes.Paste
    
    ' Ajuster la position
    pptSlide.Shapes(pptSlide.Shapes.Count).Top = 150
    pptSlide.Shapes(pptSlide.Shapes.Count).Left = 150
    
    ' Nettoyer les objets
    Set pptSlide = Nothing
    Set pptPres = Nothing
    Set pptApp = Nothing
    
    ' Message de confirmation
    MsgBox "Graphique exporté avec succès vers PowerPoint", vbInformation
End Sub

' Exporter le graphique vers Word
Sub ExportChartToWord()
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim chartObj As ChartObject
    Dim sheetName As String
    
    ' Nom de la feuille contenant le graphique
    sheetName = "{chart_sheet_name}"
    
    On Error Resume Next
    
    ' Créer une instance Word
    Set wdApp = CreateObject("Word.Application")
    If Err.Number <> 0 Then
        MsgBox "Word n'est pas disponible sur cet ordinateur.", vbExclamation
        Exit Sub
    End If
    
    ' Rendre Word visible
    wdApp.Visible = True
    
    ' Créer un nouveau document
    Set wdDoc = wdApp.Documents.Add
    
    ' Ajouter un titre
    wdDoc.Content.InsertAfter "{chart_title}"
    wdDoc.Content.InsertParagraphAfter
    wdDoc.Content.InsertParagraphAfter
    
    ' Vérifier que la feuille existe
    On Error Resume Next
    Set chartObj = ThisWorkbook.Sheets(sheetName).ChartObjects(1)
    If Err.Number <> 0 Then
        MsgBox "Feuille ou graphique introuvable: " & sheetName, vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Copier le graphique
    chartObj.Chart.CopyPicture
    
    ' Coller dans Word
    wdApp.Selection.Paste
    
    ' Nettoyer les objets
    Set wdDoc = Nothing
    Set wdApp = Nothing
    
    ' Message de confirmation
    MsgBox "Graphique exporté avec succès vers Word", vbInformation
End Sub

' Ajouter des boutons d'exportation sur la feuille du graphique
Sub AddExportButtons()
    Dim ws As Worksheet
    Dim btnPPT As Button
    Dim btnWord As Button
    Dim sheetName As String
    
    ' Nom de la feuille contenant le graphique
    sheetName = "{chart_sheet_name}"
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    If Err.Number <> 0 Then
        MsgBox "Feuille introuvable: " & sheetName, vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Supprimer les boutons existants
    On Error Resume Next
    ws.Buttons.Delete
    On Error GoTo 0
    
    ' Ajouter le bouton PowerPoint
    Set btnPPT = ws.Buttons.Add(10, 10, 150, 30)
    btnPPT.Caption = "Exporter vers PowerPoint"
    btnPPT.OnAction = "ExportChartToPowerPoint"
    
    ' Ajouter le bouton Word
    Set btnWord = ws.Buttons.Add(10, 50, 150, 30)
    btnWord.Caption = "Exporter vers Word"
    btnWord.OnAction = "ExportChartToWord"
    
    MsgBox "Boutons d'exportation ajoutés avec succès!", vbInformation
End Sub
"""
    
    # Ajouter des instructions pour l'utilisateur
    instructions = """
=== INSTRUCTIONS POUR AJOUTER CE CODE VBA À EXCEL ===

1. Ouvrez votre fichier Excel contenant le graphique
2. Appuyez sur ALT+F11 pour ouvrir l'éditeur VBA
3. Dans l'explorateur de projet (à gauche), cliquez-droit sur "VBAProject"
4. Sélectionnez "Insérer" > "Module"
5. Copiez-collez le code VBA ci-dessous dans le nouveau module
6. Fermez l'éditeur VBA (ALT+Q)
7. Enregistrez votre fichier au format .xlsm (Excel avec macros)
8. Pour ajouter des boutons d'exportation, exécutez la macro "AddExportButtons"

Note: Vous devrez peut-être activer les macros dans Excel et autoriser le contenu.
"""
    
    return instructions + "\n\n" + vba_code

@mcp.tool()
def generate_dynamic_chart_vba(source_sheet_name: str = "Données", 
                              chart_sheet_name: str = "Graphique",
                              data_range: str = "A1:D10",
                              chart_title: str = "Graphique Dynamique") -> str:
    """
    Génère le code VBA pour créer un graphique dynamique et l'exporter.
    
    Args:
        source_sheet_name: Nom de la feuille contenant les données source
        chart_sheet_name: Nom de la feuille où créer le graphique
        data_range: Plage de données au format Excel (ex: "A1:D10")
        chart_title: Titre du graphique
        
    Returns:
        Code VBA pour créer et exporter un graphique dynamique
    """
    
    vba_code = f"""
' Code VBA pour créer et exporter un graphique dynamique
' À copier-coller dans un nouveau module VBA (ALT+F11 pour ouvrir l'éditeur VBA)

Option Explicit

' Créer un graphique dynamique
Sub CreateDynamicChart()
    Dim sourceWS As Worksheet
    Dim chartWS As Worksheet
    Dim chartObj As ChartObject
    Dim chartData As Range
    
    ' Configuration
    Dim sourceSheetName As String
    Dim chartSheetName As String
    Dim dataRange As String
    Dim chartTitle As String
    
    ' Paramètres
    sourceSheetName = "{source_sheet_name}"
    chartSheetName = "{chart_sheet_name}"
    dataRange = "{data_range}"
    chartTitle = "{chart_title}"
    
    ' Vérifier que la feuille source existe
    On Error Resume Next
    Set sourceWS = ThisWorkbook.Sheets(sourceSheetName)
    If Err.Number <> 0 Then
        MsgBox "Feuille source introuvable: " & sourceSheetName, vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Vérifier/créer la feuille du graphique
    On Error Resume Next
    Set chartWS = ThisWorkbook.Sheets(chartSheetName)
    If Err.Number <> 0 Then
        ' Créer une nouvelle feuille
        Set chartWS = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        chartWS.Name = chartSheetName
    End If
    On Error GoTo 0
    
    ' Supprimer les graphiques existants
    chartWS.ChartObjects.Delete
    
    ' Définir la plage de données
    Set chartData = sourceWS.Range(dataRange)
    
    ' Créer le graphique
    Set chartObj = chartWS.ChartObjects.Add(Left:=100, Top:=50, Width:=450, Height:=300)
    
    ' Configurer le graphique
    With chartObj.Chart
        .SetSourceData Source:=chartData
        .ChartType = xlColumnClustered  ' Vous pouvez changer le type ici
        
        ' Titre
        .HasTitle = True
        .ChartTitle.Text = chartTitle
        
        ' Axes
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Axe X"
        
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Axe Y"
        
        ' Légende
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
    End With
    
    ' Ajouter un lien dynamique vers les données source
    chartWS.Range("A1").Value = "Source:"
    chartWS.Range("B1").Value = sourceSheetName
    
    chartWS.Range("A2").Value = "Plage:"
    chartWS.Range("B2").Value = dataRange
    
    chartWS.Range("A3").Value = "Titre:"
    chartWS.Range("B3").Value = chartTitle
    
    ' Ajouter des boutons d'exportation
    AddExportButtonsToSheet chartSheetName
    
    ' Ajouter un bouton pour rafraîchir le graphique
    Dim btnRefresh As Button
    Set btnRefresh = chartWS.Buttons.Add(10, 130, 150, 30)
    btnRefresh.Caption = "Rafraîchir graphique"
    btnRefresh.OnAction = "RefreshDynamicChart"
    
    MsgBox "Graphique dynamique créé avec succès!", vbInformation
End Sub

' Rafraîchir le graphique dynamique
Sub RefreshDynamicChart()
    Dim chartWS As Worksheet
    Dim sourceWS As Worksheet
    Dim chartObj As ChartObject
    Dim sourceSheetName As String
    Dim dataRange As String
    
    ' Trouver la feuille active du graphique
    Set chartWS = ActiveSheet
    
    ' Lire les informations de source
    sourceSheetName = chartWS.Range("B1").Value
    dataRange = chartWS.Range("B2").Value
    
    ' Vérifier les informations
    If sourceSheetName = "" Or dataRange = "" Then
        MsgBox "Informations de source manquantes. Impossible de rafraîchir.", vbExclamation
        Exit Sub
    End If
    
    ' Récupérer la feuille source
    On Error Resume Next
    Set sourceWS = ThisWorkbook.Sheets(sourceSheetName)
    If Err.Number <> 0 Then
        MsgBox "Feuille source introuvable: " & sourceSheetName, vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' S'assurer qu'il y a un graphique
    If chartWS.ChartObjects.Count = 0 Then
        MsgBox "Aucun graphique trouvé dans cette feuille.", vbExclamation
        Exit Sub
    End If
    
    ' Mettre à jour le graphique
    Set chartObj = chartWS.ChartObjects(1)
    chartObj.Chart.SetSourceData Source:=sourceWS.Range(dataRange)
    
    MsgBox "Graphique rafraîchi avec succès!", vbInformation
End Sub

' Ajouter des boutons d'exportation
Sub AddExportButtonsToSheet(sheetName As String)
    Dim ws As Worksheet
    Dim btnPPT As Button
    Dim btnWord As Button
    
    ' Vérifier que la feuille existe
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    If Err.Number <> 0 Then
        MsgBox "Feuille introuvable: " & sheetName, vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Ajouter le bouton PowerPoint
    Set btnPPT = ws.Buttons.Add(10, 50, 150, 30)
    btnPPT.Caption = "Exporter vers PowerPoint"
    btnPPT.OnAction = "ExportChartToPowerPoint"
    
    ' Ajouter le bouton Word
    Set btnWord = ws.Buttons.Add(10, 90, 150, 30)
    btnWord.Caption = "Exporter vers Word"
    btnWord.OnAction = "ExportChartToWord"
End Sub

' Exporter le graphique vers PowerPoint
Sub ExportChartToPowerPoint()
    Dim pptApp As Object
    Dim pptPres As Object
    Dim pptSlide As Object
    Dim chartObj As ChartObject
    Dim chartTitle As String
    
    ' Obtenir le titre du graphique
    chartTitle = ActiveSheet.Range("B3").Value
    If chartTitle = "" Then
        chartTitle = "Graphique Dynamique"
    End If
    
    ' Vérifier qu'il y a un graphique
    If ActiveSheet.ChartObjects.Count = 0 Then
        MsgBox "Aucun graphique trouvé dans cette feuille.", vbExclamation
        Exit Sub
    End If
    
    ' Créer une instance PowerPoint
    On Error Resume Next
    Set pptApp = CreateObject("PowerPoint.Application")
    If Err.Number <> 0 Then
        MsgBox "PowerPoint n'est pas disponible sur cet ordinateur.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Rendre PowerPoint visible
    pptApp.Visible = True
    
    ' Créer une nouvelle présentation
    Set pptPres = pptApp.Presentations.Add
    
    ' Ajouter une diapositive
    Set pptSlide = pptPres.Slides.Add(1, 11) ' 11 = ppLayoutTitleOnly
    
    ' Définir le titre
    pptSlide.Shapes.Title.TextFrame.TextRange.Text = chartTitle
    
    ' Copier le graphique
    Set chartObj = ActiveSheet.ChartObjects(1)
    chartObj.Chart.CopyPicture
    
    ' Coller dans PowerPoint
    pptSlide.Shapes.Paste
    
    ' Ajuster la position
    pptSlide.Shapes(pptSlide.Shapes.Count).Top = 150
    pptSlide.Shapes(pptSlide.Shapes.Count).Left = 150
    
    ' Nettoyer les objets
    Set pptSlide = Nothing
    Set pptPres = Nothing
    Set pptApp = Nothing
    
    ' Message de confirmation
    MsgBox "Graphique exporté avec succès vers PowerPoint", vbInformation
End Sub

' Exporter le graphique vers Word
Sub ExportChartToWord()
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim chartObj As ChartObject
    Dim chartTitle As String
    
    ' Obtenir le titre du graphique
    chartTitle = ActiveSheet.Range("B3").Value
    If chartTitle = "" Then
        chartTitle = "Graphique Dynamique"
    End If
    
    ' Vérifier qu'il y a un graphique
    If ActiveSheet.ChartObjects.Count = 0 Then
        MsgBox "Aucun graphique trouvé dans cette feuille.", vbExclamation
        Exit Sub
    End If
    
    ' Créer une instance Word
    On Error Resume Next
    Set wdApp = CreateObject("Word.Application")
    If Err.Number <> 0 Then
        MsgBox "Word n'est pas disponible sur cet ordinateur.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Rendre Word visible
    wdApp.Visible = True
    
    ' Créer un nouveau document
    Set wdDoc = wdApp.Documents.Add
    
    ' Ajouter un titre
    wdDoc.Content.InsertAfter chartTitle
    wdDoc.Content.InsertParagraphAfter
    wdDoc.Content.InsertParagraphAfter
    
    ' Copier le graphique
    Set chartObj = ActiveSheet.ChartObjects(1)
    chartObj.Chart.CopyPicture
    
    ' Coller dans Word
    wdApp.Selection.Paste
    
    ' Nettoyer les objets
    Set wdDoc = Nothing
    Set wdApp = Nothing
    
    ' Message de confirmation
    MsgBox "Graphique exporté avec succès vers Word", vbInformation
End Sub
"""
    
    # Ajouter des instructions pour l'utilisateur
    instructions = """
=== INSTRUCTIONS POUR CRÉER UN GRAPHIQUE DYNAMIQUE AVEC VBA ===

1. Ouvrez votre fichier Excel contenant les données
2. Appuyez sur ALT+F11 pour ouvrir l'éditeur VBA
3. Dans l'explorateur de projet (à gauche), cliquez-droit sur "VBAProject"
4. Sélectionnez "Insérer" > "Module"
5. Copiez-collez le code VBA ci-dessous dans le nouveau module
6. Fermez l'éditeur VBA (ALT+Q)
7. Enregistrez votre fichier au format .xlsm (Excel avec macros)
8. Pour créer le graphique dynamique, exécutez la macro "CreateDynamicChart"

Note: Vous devrez peut-être activer les macros dans Excel et autoriser le contenu.
Pour personnaliser, modifiez les valeurs des variables en haut de la macro CreateDynamicChart.
"""
    
    return instructions + "\n\n" + vba_code

if __name__ == "__main__":
    mcp.run()