Attribute VB_Name = "Module_Dashboard"
'===============================================================================
' FINANCE TRACKER VBA - MODULE TABLEAU DE BORD
' Version: 1.0
' Description: Gestion complète du tableau de bord avec graphiques et indicateurs
' Fonction: Interface principale de visualisation des données financières
'===============================================================================

Option Explicit

' Constantes pour la mise en forme du tableau de bord
Private Const COULEUR_REVENUS As Long = &H4472C4     ' Bleu foncé
Private Const COULEUR_DEPENSES As Long = &HC5504B    ' Rouge
Private Const COULEUR_EPARGNE As Long = &H70AD47     ' Vert
Private Const COULEUR_BUDGET As Long = &HFFC000      ' Orange

'===============================================================================
' PROCEDURES PRINCIPALES DU TABLEAU DE BORD
'===============================================================================

Sub CreerTableauBord(ws As Worksheet)
    '-------------------------------------------------------------------------
    ' Crée la structure complète du tableau de bord principal
    '-------------------------------------------------------------------------
    
    On Error GoTo GestionErreur
    
    Application.ScreenUpdating = False
    
    With ws
        ' Configuration générale de la feuille
        .Cells.Clear
        .Tab.Color = RGB(68, 114, 196)
        
        ' En-tête principal
        Call CreerEntetePrincipal(ws)
        
        ' Section indicateurs clés
        Call CreerIndicateursCles(ws)
        
        ' Section graphiques
        Call CreerZoneGraphiques(ws)
        
        ' Section résumé mensuel
        Call CreerResumeMensuel(ws)
        
        ' Section alertes et notifications
        Call CreerZoneAlertes(ws)
        
        ' Formatage final
        Call AppliquerFormatageTableauBord(ws)
    End With
    
    Application.ScreenUpdating = True
    Exit Sub
    
GestionErreur:
    Application.ScreenUpdating = True
    Call EnregistrerJournal("Erreur création tableau de bord: " & Err.Description, "ERREUR")
End Sub

Sub CreerEntetePrincipal(ws As Worksheet)
    '-------------------------------------------------------------------------
    ' Crée l'en-tête principal du tableau de bord
    '-------------------------------------------------------------------------
    
    With ws
        ' Titre principal
        .Range("A1:H1").Merge
        .Range("A1").Value = "TABLEAU DE BORD FINANCIER"
        .Range("A1").Font.Size = 18
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Color = RGB(68, 114, 196)
        .Range("A1").HorizontalAlignment = xlCenter
        
        ' Sous-titre avec date
        .Range("A2:H2").Merge
        .Range("A2").Value = "Période: " & Format(ObtenirMoisCourant, "mmmm yyyy")
        .Range("A2").Font.Size = 12
        .Range("A2").HorizontalAlignment = xlCenter
        .Range("A2").Font.Color = RGB(89, 89, 89)
        
        ' Ligne de séparation
        .Range("A3:H3").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("A3:H3").Borders(xlEdgeBottom).Color = RGB(68, 114, 196)
        .Range("A3:H3").Borders(xlEdgeBottom).Weight = xlMedium
    End With
    
End Sub

Sub CreerIndicateursCles(ws As Worksheet)
    '-------------------------------------------------------------------------
    ' Crée les indicateurs clés de performance financière
    '-------------------------------------------------------------------------
    
    Dim PlageIndicateur As Range
    Dim i As Integer
    Dim TitresIndicateurs As Variant
    Dim ValeursIndicateurs As Variant
    
    TitresIndicateurs = Array("REVENUS DU MOIS", "DÉPENSES DU MOIS", "ÉPARGNE RÉALISÉE", "BUDGET RESTANT")
    
    With ws
        ' En-tête de la section
        .Range("A5").Value = "INDICATEURS CLÉS"
        .Range("A5").Font.Size = 14
        .Range("A5").Font.Bold = True
        .Range("A5").Font.Color = RGB(68, 114, 196)
        
        ' Création des cartes d'indicateurs
        For i = 0 To 3
            Set PlageIndicateur = .Range(.Cells(7, i * 2 + 1), .Cells(9, i * 2 + 2))
            Call CreerCarteIndicateur(PlageIndicateur, TitresIndicateurs(i), "=CalculerIndicateur(" & Chr(34) & TitresIndicateurs(i) & Chr(34) & ")", i)
        Next i
    End With
    
End Sub

Sub CreerCarteIndicateur(PlageIndicateur As Range, TitreIndicateur As String, FormuleValeur As String, IndexCouleur As Integer)
    '-------------------------------------------------------------------------
    ' Crée une carte individuelle pour un indicateur
    '-------------------------------------------------------------------------
    
    Dim CouleurFond As Long
    
    ' Définition des couleurs selon l'indicateur
    Select Case IndexCouleur
        Case 0: CouleurFond = COULEUR_REVENUS
        Case 1: CouleurFond = COULEUR_DEPENSES
        Case 2: CouleurFond = COULEUR_EPARGNE
        Case 3: CouleurFond = COULEUR_BUDGET
    End Select
    
    With PlageIndicateur
        .Merge
        .Value = TitreIndicateur & vbCrLf & FormuleValeur
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Font.Size = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = CouleurFond
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(255, 255, 255)
        .Borders.Weight = xlMedium
    End With
    
End Sub

Sub CreerZoneGraphiques(ws As Worksheet)
    '-------------------------------------------------------------------------
    ' Crée la zone dédiée aux graphiques de synthèse
    '-------------------------------------------------------------------------
    
    With ws
        ' En-tête de la section graphiques
        .Range("A11").Value = "ANALYSE GRAPHIQUE"
        .Range("A11").Font.Size = 14
        .Range("A11").Font.Bold = True
        .Range("A11").Font.Color = RGB(68, 114, 196)
        
        ' Zone pour graphique évolution mensuelle
        Call CreerZoneGraphique(ws, "A13:D23", "Évolution Revenus/Dépenses")
        
        ' Zone pour graphique répartition dépenses
        Call CreerZoneGraphique(ws, "E13:H23", "Répartition des Dépenses")
        
        ' Préparer les données pour les graphiques
        Call PreparerDonneesGraphiques(ws)
    End With
    
End Sub

Sub CreerZoneGraphique(ws As Worksheet, PlageZone As String, TitreGraphique As String)
    '-------------------------------------------------------------------------
    ' Crée une zone réservée pour un graphique spécifique
    '-------------------------------------------------------------------------
    
    Dim Plage As Range
    Set Plage = ws.Range(PlageZone)
    
    With Plage
        .Merge
        .Value = TitreGraphique & vbCrLf & vbCrLf & "[Graphique généré automatiquement]"
        .Font.Size = 10
        .Font.Color = RGB(89, 89, 89)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(248, 248, 248)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
    End With
    
End Sub

Sub CreerResumeMensuel(ws As Worksheet)
    '-------------------------------------------------------------------------
    ' Crée le tableau de résumé mensuel détaillé
    '-------------------------------------------------------------------------
    
    Dim TabResume As Range
    
    With ws
        ' En-tête de la section
        .Range("A25").Value = "RÉSUMÉ MENSUEL DÉTAILLÉ"
        .Range("A25").Font.Size = 14
        .Range("A25").Font.Bold = True
        .Range("A25").Font.Color = RGB(68, 114, 196)
        
        ' Création du tableau de résumé
        Set TabResume = .Range("A27:H35")
        Call CreerTableauResume(TabResume)
    End With
    
End Sub

Sub CreerTableauResume(TabResume As Range)
    '-------------------------------------------------------------------------
    ' Construit le tableau détaillé du résumé mensuel
    '-------------------------------------------------------------------------
    
    Dim EntetesColonnes As Variant
    Dim LignesResume As Variant
    Dim i As Integer, j As Integer
    
    EntetesColonnes = Array("CATÉGORIE", "BUDGET PRÉVU", "MONTANT RÉEL", "ÉCART", "ÉCART %", "STATUT", "TENDANCE", "ACTIONS")
    LignesResume = Array("Revenus Salaire", "Revenus Autres", "Logement", "Alimentation", "Transport", "Loisirs", "Épargne", "TOTAL")
    
    With TabResume
        ' En-têtes de colonnes
        For i = 0 To UBound(EntetesColonnes)
            .Cells(1, i + 1).Value = EntetesColonnes(i)
            .Cells(1, i + 1).Font.Bold = True
            .Cells(1, i + 1).Font.Color = RGB(255, 255, 255)
            .Cells(1, i + 1).Interior.Color = RGB(68, 114, 196)
        Next i
        
        ' Lignes de données
        For i = 0 To UBound(LignesResume)
            .Cells(i + 2, 1).Value = LignesResume(i)
            For j = 2 To 8
                .Cells(i + 2, j).Value = "=CalculerResume(" & Chr(34) & LignesResume(i) & Chr(34) & "," & j & ")"
            Next j
        Next i
        
        ' Formatage du tableau
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
        .Font.Size = 9
        .HorizontalAlignment = xlCenter
    End With
    
End Sub

Sub CreerZoneAlertes(ws As Worksheet)
    '-------------------------------------------------------------------------
    ' Crée la zone des alertes et notifications importantes
    '-------------------------------------------------------------------------
    
    With ws
        ' En-tête de la section
        .Range("A37").Value = "ALERTES ET NOTIFICATIONS"
        .Range("A37").Font.Size = 14
        .Range("A37").Font.Bold = True
        .Range("A37").Font.Color = RGB(196, 89, 17)
        
        ' Zone d'alerte principale
        .Range("A39:H42").Merge
        .Range("A39").Value = "=GenererAlertes()"
        .Range("A39").Font.Size = 10
        .Range("A39").VerticalAlignment = xlTop
        .Range("A39").Interior.Color = RGB(255, 242, 204)
        .Range("A39").Borders.LineStyle = xlContinuous
        .Range("A39").Borders.Color = RGB(196, 89, 17)
    End With
    
End Sub

'===============================================================================
' PROCEDURES DE MISE À JOUR DU TABLEAU DE BORD
'===============================================================================

Sub ActualiserTableauBord()
    '-------------------------------------------------------------------------
    ' Met à jour toutes les données du tableau de bord
    '-------------------------------------------------------------------------
    
    On Error GoTo GestionErreur
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    
    ' Mise à jour de la période affichée
    ws.Range("A2").Value = "Période: " & Format(ObtenirMoisCourant, "mmmm yyyy")
    
    ' Recalcul des indicateurs
    Call RecalculerIndicateurs(ws)
    
    ' Mise à jour des graphiques
    Call ActualiserGraphiques(ws)
    
    ' Mise à jour du résumé
    Call ActualiserResumeMensuel(ws)
    
    ' Mise à jour des alertes
    Call ActualiserAlertes(ws)
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Call EnregistrerJournal("Tableau de bord actualisé", "INFO")
    Exit Sub
    
GestionErreur:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Call EnregistrerJournal("Erreur actualisation tableau de bord: " & Err.Description, "ERREUR")
End Sub

Sub RecalculerIndicateurs(ws As Worksheet)
    '-------------------------------------------------------------------------
    ' Recalcule tous les indicateurs clés du tableau de bord
    '-------------------------------------------------------------------------
    
    Dim RevenusMois As Currency
    Dim DepensesMois As Currency
    Dim EpargneRealisee As Currency
    Dim BudgetRestant As Currency
    
    ' Récupération des données depuis les feuilles de données
    RevenusMois = CalculerRevenusMois(ObtenirMoisCourant)
    DepensesMois = CalculerDepensesMois(ObtenirMoisCourant)
    EpargneRealisee = RevenusMois - DepensesMois
    BudgetRestant = CalculerBudgetRestant(ObtenirMoisCourant)
    
    ' Mise à jour des cellules d'affichage
    With ws
        .Range("B8").Value = FormaterMontant(RevenusMois)
        .Range("D8").Value = FormaterMontant(DepensesMois)
        .Range("F8").Value = FormaterMontant(EpargneRealisee)
        .Range("H8").Value = FormaterMontant(BudgetRestant)
    End With
    
End Sub

'===============================================================================
' FONCTIONS DE CALCUL POUR LE TABLEAU DE BORD
'===============================================================================

Function CalculerRevenusMois(MoisReference As Date) As Currency
    '-------------------------------------------------------------------------
    ' Calcule le total des revenus pour un mois donné
    '-------------------------------------------------------------------------
    
    Dim ws As Worksheet
    Dim DerniereLigne As Long
    Dim i As Long
    Dim Total As Currency
    
    Set ws = ThisWorkbook.Worksheets("Donnees_Revenus")
    
    If ws.Cells(1, 1).Value = "" Then
        CalculerRevenusMois = 0
        Exit Function
    End If
    
    DerniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Total = 0
    
    For i = 2 To DerniereLigne
        If Month(ws.Cells(i, 1).Value) = Month(MoisReference) And _
           Year(ws.Cells(i, 1).Value) = Year(MoisReference) Then
            Total = Total + ws.Cells(i, 4).Value ' Colonne montant réel
        End If
    Next i
    
    CalculerRevenusMois = Total
    
End Function

Function CalculerDepensesMois(MoisReference As Date) As Currency
    '-------------------------------------------------------------------------
    ' Calcule le total des dépenses pour un mois donné
    '-------------------------------------------------------------------------
    
    Dim ws As Worksheet
    Dim DerniereLigne As Long
    Dim i As Long
    Dim Total As Currency
    
    Set ws = ThisWorkbook.Worksheets("Donnees_Depenses")
    
    If ws.Cells(1, 1).Value = "" Then
        CalculerDepensesMois = 0
        Exit Function
    End If
    
    DerniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Total = 0
    
    For i = 2 To DerniereLigne
        If Month(ws.Cells(i, 1).Value) = Month(MoisReference) And _
           Year(ws.Cells(i, 1).Value) = Year(MoisReference) Then
            Total = Total + ws.Cells(i, 4).Value ' Colonne montant réel
        End If
    Next i
    
    CalculerDepensesMois = Total
    
End Function

Function CalculerBudgetRestant(MoisReference As Date) As Currency
    '-------------------------------------------------------------------------
    ' Calcule le budget restant pour le mois en cours
    '-------------------------------------------------------------------------
    
    Dim BudgetTotal As Currency
    Dim DepensesReelles As Currency
    
    BudgetTotal = ObtenirBudgetMensuelTotal()
    DepensesReelles = CalculerDepensesMois(MoisReference)
    
    CalculerBudgetRestant = BudgetTotal - DepensesReelles
    
End Function

'===============================================================================
' PROCEDURES DE FORMATAGE
'===============================================================================

Sub AppliquerFormatageTableauBord(ws As Worksheet)
    '-------------------------------------------------------------------------
    ' Applique le formatage global du tableau de bord
    '-------------------------------------------------------------------------
    
    With ws
        ' Mise en forme générale
        .Cells.Font.Name = "Segoe UI"
        .Cells.Font.Size = 9
        
        ' Largeur des colonnes
        .Columns("A:H").ColumnWidth = 12
        
        ' Hauteur des lignes pour l'affichage optimal
        .Rows("7:9").RowHeight = 25
        .Rows("13:23").RowHeight = 18
        
        ' Protection des cellules avec formules
        .Cells.Locked = True
        .Range("A1:H50").Locked = False ' Zone de saisie déverrouillée si nécessaire
    End With
    
End Sub

'===============================================================================
' FIN DU MODULE TABLEAU DE BORD
'===============================================================================
