Attribute VB_Name = "Module_Graphiques"
'===============================================================================
' FINANCE TRACKER VBA - MODULE GRAPHIQUES
' Version: 1.0
' Description: Création et gestion des graphiques financiers dynamiques
' Fonction: Visualisation avancée des données avec graphiques interactifs
'===============================================================================

Option Explicit

' Constantes pour les graphiques
Private Const COULEUR_SERIE_REVENUS As Long = &H4472C4
Private Const COULEUR_SERIE_DEPENSES As Long = &HC5504B
Private Const COULEUR_SERIE_EPARGNE As Long = &H70AD47
Private Const COULEUR_SERIE_BUDGET As Long = &HFFC000

' Énumération des types de graphiques
Public Enum TypeGraphique
    EvolutionMensuelle = 1
    RepartitionDepenses = 2
    ComparaisonBudget = 3
    TendanceEpargne = 4
End Enum

'===============================================================================
' PROCEDURES PRINCIPALES DE CREATION DES GRAPHIQUES
'===============================================================================

Sub PreparerDonneesGraphiques(ws As Worksheet)
    '-------------------------------------------------------------------------
    ' Prépare toutes les données nécessaires pour les graphiques du tableau de bord
    '-------------------------------------------------------------------------
    
    On Error GoTo GestionErreur
    
    Application.ScreenUpdating = False
    
    ' Préparation des données d'évolution mensuelle
    Call PreparerDonneesEvolution(ws)
    
    ' Préparation des données de répartition
    Call PreparerDonneesRepartition(ws)
    
    ' Création des graphiques
    Call CreerTousLesGraphiques(ws)
    
    Application.ScreenUpdating = True
    Exit Sub
    
GestionErreur:
    Application.ScreenUpdating = True
    Call EnregistrerJournal("Erreur préparation graphiques: " & Err.Description, "ERREUR")
End Sub

Sub PreparerDonneesEvolution(ws As Worksheet)
    '-------------------------------------------------------------------------
    ' Prépare les données d'évolution sur 12 mois pour le graphique linéaire
    '-------------------------------------------------------------------------
    
    Dim i As Integer
    Dim MoisAnalyse As Date
    Dim PlageDebut As String
    
    PlageDebut = "J1" ' Zone de données cachée pour les graphiques
    
    With ws
        ' En-têtes des colonnes de données
        .Range(PlageDebut).Value = "Mois"
        .Range(PlageDebut).Offset(0, 1).Value = "Revenus"
        .Range(PlageDebut).Offset(0, 2).Value = "Dépenses"
        .Range(PlageDebut).Offset(0, 3).Value = "Épargne"
        .Range(PlageDebut).Offset(0, 4).Value = "Budget Prévu"
        
        ' Données des 12 derniers mois
        For i = 11 To 0 Step -1
            MoisAnalyse = DateAdd("m", -i, ObtenirMoisCourant())
            
            .Range(PlageDebut).Offset(12 - i, 0).Value = Format(MoisAnalyse, "mmm yy")
            .Range(PlageDebut).Offset(12 - i, 1).Value = CalculerRevenusMois(MoisAnalyse)
            .Range(PlageDebut).Offset(12 - i, 2).Value = CalculerDepensesMois(MoisAnalyse)
            .Range(PlageDebut).Offset(12 - i, 3).Value = CalculerRevenusMois(MoisAnalyse) - CalculerDepensesMois(MoisAnalyse)
            .Range(PlageDebut).Offset(12 - i, 4).Value = CalculerBudgetRevenusMois(MoisAnalyse, True) - CalculerBudgetDepensesMois(MoisAnalyse, True)
        Next i
    End With
    
End Sub

Sub PreparerDonneesRepartition(ws As Worksheet)
    '-------------------------------------------------------------------------
    ' Prépare les données de répartition des dépenses par catégorie
    '-------------------------------------------------------------------------
    
    Dim Categories As Variant
    Dim i As Integer
    Dim MoisActuel As Date
    Dim PlageDebut As String
    Dim MontantCategorie As Currency
    
    Categories = Array("Logement", "Alimentation", "Transport", "Loisirs", "Santé", "Vêtements", "Autres")
    PlageDebut = "O1" ' Zone de données pour graphique secteurs
    MoisActuel = ObtenirMoisCourant()
    
    With ws
        ' En-têtes pour le graphique secteurs
        .Range(PlageDebut).Value = "Catégorie"
        .Range(PlageDebut).Offset(0, 1).Value = "Montant"
        .Range(PlageDebut).Offset(0, 2).Value = "Pourcentage"
        
        ' Données par catégorie
        For i = 0 To UBound(Categories)
            MontantCategorie = ObtenirBudgetCategorie(Categories(i), MoisActuel, False)
            
            .Range(PlageDebut).Offset(i + 1, 0).Value = Categories(i)
            .Range(PlageDebut).Offset(i + 1, 1).Value = MontantCategorie
            .Range(PlageDebut).Offset(i + 1, 2).Value = "=SI(SOMME(P2:P8)>0,P" & (i + 2) & "/SOMME(P$2:P$8),0)"
        Next i
    End With
    
End Sub

Sub CreerTousLesGraphiques(ws As Worksheet)
    '-------------------------------------------------------------------------
    ' Crée tous les graphiques du tableau de bord
    '-------------------------------------------------------------------------
    
    ' Suppression des anciens graphiques
    Call SupprimerAncienGraphiques(ws)
    
    ' Création du graphique d'évolution mensuelle
    Call CreerGraphiqueEvolution(ws, "A13:D23")
    
    ' Création du graphique de répartition des dépenses
    Call CreerGraphiqueRepartition(ws, "E13:H23")
    
End Sub

Sub SupprimerAncienGraphiques(ws As Worksheet)
    '-------------------------------------------------------------------------
    ' Supprime les anciens graphiques présents sur la feuille
    '-------------------------------------------------------------------------
    
    Dim objGraphique As ChartObject
    
    For Each objGraphique In ws.ChartObjects
        objGraphique.Delete
    Next objGraphique
    
End Sub

'===============================================================================
' CREATION DES GRAPHIQUES SPECIFIQUES
'===============================================================================

Sub CreerGraphiqueEvolution(ws As Worksheet, PlageGraphique As String)
    '-------------------------------------------------------------------------
    ' Crée le graphique d'évolution mensuelle des revenus/dépenses
    '-------------------------------------------------------------------------
    
    Dim objGraphique As ChartObject
    Dim PlageData As Range
    
    ' Définition de la plage de données
    Set PlageData = ws.Range("J1:N13")
    
    ' Création du graphique
    Set objGraphique = ws.ChartObjects.Add(ws.Range(PlageGraphique).Left, _
                                          ws.Range(PlageGraphique).Top, _
                                          ws.Range(PlageGraphique).Width, _
                                          ws.Range(PlageGraphique).Height)
    
    With objGraphique.Chart
        .ChartType = xlLineMarkers
        .SetSourceData PlageData, xlColumns
        .HasTitle = True
        .ChartTitle.Text = "Évolution Financière (12 mois)"
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Bold = True
        
        ' Configuration de l'axe X (mois)
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Période"
        .Axes(xlCategory).AxisTitle.Font.Size = 9
        
        ' Configuration de l'axe Y (montants)
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Montant (€)"
        .Axes(xlValue).AxisTitle.Font.Size = 9
        .Axes(xlValue).DisplayUnit = xlThousands
        .Axes(xlValue).HasDisplayUnitLabel = True
        .Axes(xlValue).DisplayUnitLabel.Text = "k€"
        
        ' Formatage des séries
        Call FormaterSerieGraphique(.SeriesCollection(1), "Revenus", COULEUR_SERIE_REVENUS)
        Call FormaterSerieGraphique(.SeriesCollection(2), "Dépenses", COULEUR_SERIE_DEPENSES)
        Call FormaterSerieGraphique(.SeriesCollection(3), "Épargne", COULEUR_SERIE_EPARGNE)
        Call FormaterSerieGraphique(.SeriesCollection(4), "Budget Prévu", COULEUR_SERIE_BUDGET)
        
        ' Configuration générale
        .Legend.Position = xlLegendPositionBottom
        .Legend.Font.Size = 8
        .PlotArea.Interior.Color = RGB(248, 248, 248)
        .ChartArea.Border.LineStyle = xlContinuous
        .HasDataTable = False
    End With
    
    objGraphique.Name = "GraphiqueEvolution"
    
End Sub

Sub CreerGraphiqueRepartition(ws As Worksheet, PlageGraphique As String)
    '-------------------------------------------------------------------------
    ' Crée le graphique de répartition des dépenses en secteurs
    '-------------------------------------------------------------------------
    
    Dim objGraphique As ChartObject
    Dim PlageData As Range
    
    ' Définition de la plage de données (catégories et montants seulement)
    Set PlageData = ws.Range("O1:P8")
    
    ' Création du graphique
    Set objGraphique = ws.ChartObjects.Add(ws.Range(PlageGraphique).Left, _
                                          ws.Range(PlageGraphique).Top, _
                                          ws.Range(PlageGraphique).Width, _
                                          ws.Range(PlageGraphique).Height)
    
    With objGraphique.Chart
        .ChartType = xlDoughnut
        .SetSourceData PlageData, xlColumns
        .HasTitle = True
        .ChartTitle.Text = "Répartition des Dépenses"
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Bold = True
        
        ' Configuration de la série principale
        With .SeriesCollection(1)
            .HasDataLabels = True
            .DataLabels.ShowPercentage = True
            .DataLabels.ShowCategoryName = False
            .DataLabels.ShowValue = False
            .DataLabels.Position = xlLabelPositionBestFit
            .DataLabels.Font.Size = 8
            .DataLabels.Font.Bold = True
            
            ' Application des couleurs personnalisées
            Call AppliquerCouleursRepartition(.Points)
        End With
        
        ' Configuration générale
        .Legend.Position = xlLegendPositionRight
        .Legend.Font.Size = 8
        .PlotArea.Interior.Color = RGB(248, 248, 248)
        .ChartArea.Border.LineStyle = xlContinuous
    End With
    
    objGraphique.Name = "GraphiqueRepartition"
    
End Sub

Sub FormaterSerieGraphique(Serie As Series, NomSerie As String, CouleurSerie As Long)
    '-------------------------------------------------------------------------
    ' Formate une série de données dans un graphique linéaire
    '-------------------------------------------------------------------------
    
    With Serie
        .Name = NomSerie
        .Border.Color = CouleurSerie
        .Border.Weight = xlMedium
        .MarkerStyle = xlMarkerStyleCircle
        .MarkerSize = 5
        .MarkerBackgroundColor = CouleurSerie
        .MarkerForegroundColor = CouleurSerie
        
        ' Ligne de tendance pour l'épargne
        If NomSerie = "Épargne" Then
            .Trendlines.Add Type:=xlLinear
            .Trendlines(1).Border.Color = CouleurSerie
            .Trendlines(1).Border.LineStyle = xlDash
        End If
    End With
    
End Sub

Sub AppliquerCouleursRepartition(Points As Points)
    '-------------------------------------------------------------------------
    ' Applique des couleurs distinctes aux secteurs du graphique de répartition
    '-------------------------------------------------------------------------
    
    Dim CouleursCategories As Variant
    Dim i As Integer
    
    CouleursCategories = Array(RGB(68, 114, 196), RGB(112, 173, 71), RGB(255, 192, 0), _
                              RGB(196, 89, 17), RGB(91, 155, 213), RGB(237, 125, 49), _
                              RGB(165, 165, 165))
    
    For i = 1 To Points.Count
        If i <= UBound(CouleursCategories) + 1 Then
            Points(i).Interior.Color = CouleursCategories(i - 1)
        End If
    Next i
    
End Sub

'===============================================================================
' PROCEDURES DE MISE A JOUR DES GRAPHIQUES
'===============================================================================

Sub ActualiserGraphiques(ws As Worksheet)
    '-------------------------------------------------------------------------
    ' Met à jour tous les graphiques avec les dernières données
    '-------------------------------------------------------------------------
    
    On Error GoTo GestionErreur
    
    Application.ScreenUpdating = False
    
    ' Mise à jour des données sources
    Call PreparerDonneesEvolution(ws)
    Call PreparerDonneesRepartition(ws)
    
    ' Actualisation des graphiques existants
    Call ActualiserGraphiqueEvolution(ws)
    Call ActualiserGraphiqueRepartition(ws)
    
    Application.ScreenUpdating = True
    Exit Sub
    
GestionErreur:
    Application.ScreenUpdating = True
    Call EnregistrerJournal("Erreur actualisation graphiques: " & Err.Description, "ERREUR")
End Sub

Sub ActualiserGraphiqueEvolution(ws As Worksheet)
    '-------------------------------------------------------------------------
    ' Met à jour spécifiquement le graphique d'évolution
    '-------------------------------------------------------------------------
    
    Dim objGraphique As ChartObject
    
    On Error Resume Next
    Set objGraphique = ws.ChartObjects("GraphiqueEvolution")
    On Error GoTo 0
    
    If Not objGraphique Is Nothing Then
        With objGraphique.Chart
            .SetSourceData ws.Range("J1:N13"), xlColumns
            .ChartTitle.Text = "Évolution Financière - " & Format(ObtenirMoisCourant, "mmmm yyyy")
        End With
    End If
    
End Sub

Sub ActualiserGraphiqueRepartition(ws As Worksheet)
    '-------------------------------------------------------------------------
    ' Met à jour spécifiquement le graphique de répartition
    '-------------------------------------------------------------------------
    
    Dim objGraphique As ChartObject
    
    On Error Resume Next
    Set objGraphique = ws.ChartObjects("GraphiqueRepartition")
    On Error GoTo 0
    
    If Not objGraphique Is Nothing Then
        With objGraphique.Chart
            .SetSourceData ws.Range("O1:P8"), xlColumns
            .ChartTitle.Text = "Répartition Dépenses - " & Format(ObtenirMoisCourant, "mmmm yyyy")
        End With
    End If
    
End Sub

'===============================================================================
' CREATION DE GRAPHIQUES POUR LES RAPPORTS
'===============================================================================

Function CreerGraphiqueRapport(ws As Worksheet, TypeGraph As TypeGraphique, _
                              PlageDestination As String, PlageData As String) As ChartObject
    '-------------------------------------------------------------------------
    ' Crée un graphique spécialisé pour les rapports
    '-------------------------------------------------------------------------
    
    Dim objGraphique As ChartObject
    Dim PlageGraphique As Range
    
    Set PlageGraphique = ws.Range(PlageDestination)
    
    Set objGraphique = ws.ChartObjects.Add(PlageGraphique.Left, _
                                          PlageGraphique.Top, _
                                          PlageGraphique.Width, _
                                          PlageGraphique.Height)
    
    Select Case TypeGraph
        Case EvolutionMensuelle
            Call ConfigurerGraphiqueEvolutionRapport(objGraphique, ws.Range(PlageData))
            
        Case RepartitionDepenses
            Call ConfigurerGraphiqueRepartitionRapport(objGraphique, ws.Range(PlageData))
            
        Case ComparaisonBudget
            Call ConfigurerGraphiqueComparaisonRapport(objGraphique, ws.Range(PlageData))
            
        Case TendanceEpargne
            Call ConfigurerGraphiqueTendanceRapport(objGraphique, ws.Range(PlageData))
    End Select
    
    Set CreerGraphiqueRapport = objGraphique
    
End Function

Sub ConfigurerGraphiqueEvolutionRapport(objGraphique As ChartObject, PlageData As Range)
    '-------------------------------------------------------------------------
    ' Configure un graphique d'évolution pour rapport
    '-------------------------------------------------------------------------
    
    With objGraphique.Chart
        .ChartType = xlLineMarkers
        .SetSourceData PlageData, xlColumns
        .HasTitle = True
        .ChartTitle.Text = "Évolution des Finances sur 12 Mois"
        
        ' Style professionnel pour rapport
        .ChartArea.Interior.Color = RGB(255, 255, 255)
        .PlotArea.Interior.Color = RGB(250, 250, 250)
        .Legend.Position = xlLegendPositionBottom
        
        ' Formatage des axes
        .Axes(xlValue).TickLabels.NumberFormat = "#,##0 €"
        .Axes(xlCategory).TickLabelSpacing = 2
    End With
    
End Sub

Sub ConfigurerGraphiqueRepartitionRapport(objGraphique As ChartObject, PlageData As Range)
    '-------------------------------------------------------------------------
    ' Configure un graphique de répartition pour rapport
    '-------------------------------------------------------------------------
    
    With objGraphique.Chart
        .ChartType = xlPie
        .SetSourceData PlageData, xlColumns
        .HasTitle = True
        .ChartTitle.Text = "Répartition des Dépenses par Catégorie"
        
        ' Configuration des étiquettes
        With .SeriesCollection(1)
            .HasDataLabels = True
            .DataLabels.ShowPercentage = True
            .DataLabels.ShowCategoryName = True
            .DataLabels.Position = xlLabelPositionBestFit
        End With
        
        .Legend.Position = xlLegendPositionNone
    End With
    
End Sub

Sub ConfigurerGraphiqueComparaisonRapport(objGraphique As ChartObject, PlageData As Range)
    '-------------------------------------------------------------------------
    ' Configure un graphique de comparaison budget vs réel
    '-------------------------------------------------------------------------
    
    With objGraphique.Chart
        .ChartType = xlColumnClustered
        .SetSourceData PlageData, xlColumns
        .HasTitle = True
        .ChartTitle.Text = "Comparaison Budget Prévu vs Réalisé"
        
        ' Formatage des séries
        .SeriesCollection(1).Interior.Color = COULEUR_SERIE_BUDGET
        .SeriesCollection(2).Interior.Color = COULEUR_SERIE_DEPENSES
        
        .Legend.Position = xlLegendPositionTop
        .Axes(xlValue).TickLabels.NumberFormat = "#,##0 €"
    End With
    
End Sub

Sub ConfigurerGraphiqueTendanceRapport(objGraphique As ChartObject, PlageData As Range)
    '-------------------------------------------------------------------------
    ' Configure un graphique de tendance d'épargne
    '-------------------------------------------------------------------------
    
    With objGraphique.Chart
        .ChartType = xlAreaStacked
        .SetSourceData PlageData, xlColumns
        .HasTitle = True
        .ChartTitle.Text = "Tendance d'Épargne et Projections"
        
        ' Configuration des séries
        .SeriesCollection(1).Interior.Color = COULEUR_SERIE_EPARGNE
        If .SeriesCollection.Count > 1 Then
            .SeriesCollection(2).Interior.Color = RGB(144, 238, 144) ' Vert clair pour projection
        End If
        
        .Legend.Position = xlLegendPositionBottom
        .Axes(xlValue).TickLabels.NumberFormat = "#,##0 €"
    End With
    
End Sub

'===============================================================================
' UTILITAIRES DE GESTION DES GRAPHIQUES
'===============================================================================

Sub ExporterGraphique(NomGraphique As String, CheminExport As String)
    '-------------------------------------------------------------------------
    ' Exporte un graphique en image PNG
    '-------------------------------------------------------------------------
    
    Dim ws As Worksheet
    Dim objGraphique As ChartObject
    
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    
    On Error Resume Next
    Set objGraphique = ws.ChartObjects(NomGraphique)
    On Error GoTo 0
    
    If Not objGraphique Is Nothing Then
        objGraphique.Chart.Export CheminExport & "\" & NomGraphique & ".png", "PNG"
        Call EnregistrerJournal("Graphique exporté: " & NomGraphique, "INFO")
    End If
    
End Sub

Sub RedimensionnerGraphiques(ws As Worksheet)
    '-------------------------------------------------------------------------
    ' Redimensionne automatiquement tous les graphiques selon la taille de la fenêtre
    '-------------------------------------------------------------------------
    
    Dim objGraphique As ChartObject
    Dim FacteurEchelle As Double
    
    FacteurEchelle = Application.WindowState / xlMaximized
    
    For Each objGraphique In ws.ChartObjects
        With objGraphique
            .Width = .Width * FacteurEchelle
            .Height = .Height * FacteurEchelle
        End With
    Next objGraphique
    
End Sub

Function VerifierDonneesGraphique(PlageData As Range) As Boolean
    '-------------------------------------------------------------------------
    ' Vérifie si les données sont suffisantes pour créer un graphique
    '-------------------------------------------------------------------------
    
    Dim CellulesVides As Integer
    Dim TotalCellules As Integer
    
    TotalCellules = PlageData.Cells.Count
    CellulesVides = Application.WorksheetFunction.CountBlank(PlageData)
    
    ' Au moins 70% des données doivent être présentes
    VerifierDonneesGraphique = (CellulesVides / TotalCellules) < 0.3
    
End Function

Sub PersonnaliserStyleGraphique(objGraphique As ChartObject, StylePersonnalise As String)
    '-------------------------------------------------------------------------
    ' Applique un style personnalisé à un graphique
    '-------------------------------------------------------------------------
    
    With objGraphique.Chart
        Select Case StylePersonnalise
            Case "Professionnel"
                .ChartArea.Interior.Color = RGB(255, 255, 255)
                .PlotArea.Interior.Color = RGB(248, 248, 248)
                .ChartArea.Border.Color = RGB(128, 128, 128)
                
            Case "Moderne"
                .ChartArea.Interior.Color = RGB(45, 45, 45)
                .PlotArea.Interior.Color = RGB(60, 60, 60)
                .ChartTitle.Font.Color = RGB(255, 255, 255)
                
            Case "Coloré"
                .ChartArea.Interior.Color = RGB(240, 248, 255)
                .PlotArea.Interior.Color = RGB(255, 255, 255)
                .ChartArea.Border.Color = COULEUR_SERIE_REVENUS
        End Select
    End With
    
End Sub

'===============================================================================
' FIN DU MODULE GRAPHIQUES
'===============================================================================
