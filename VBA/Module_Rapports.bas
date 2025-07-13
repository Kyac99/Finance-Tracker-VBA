Attribute VB_Name = "Module_Rapports"
'===============================================================================
' FINANCE TRACKER VBA - MODULE RAPPORTS
' Version: 1.0
' Description: Génération automatisée de rapports financiers détaillés
' Fonction: Création de rapports mensuels, annuels et analyses personnalisées
'===============================================================================

Option Explicit

' Constantes pour les rapports
Private Const REPERTOIRE_RAPPORTS As String = "Rapports\"
Private Const FORMAT_DATE_RAPPORT As String = "yyyy-mm-dd"

' Énumération des types de rapports
Public Enum TypeRapport
    RapportMensuel = 1
    RapportAnnuel = 2
    RapportComparatif = 3
    RapportProjection = 4
    RapportPersonnalise = 5
End Enum

' Structure pour les paramètres de rapport
Public Type ParametresRapport
    TypeRap As TypeRapport
    DateDebut As Date
    DateFin As Date
    IncludeGraphiques As Boolean
    FormatSortie As String
    NomFichier As String
End Type

'===============================================================================
' PROCEDURES PRINCIPALES DE GENERATION DES RAPPORTS
'===============================================================================

Sub GenererRapportMensuel()
    '-------------------------------------------------------------------------
    ' Génère le rapport mensuel complet
    '-------------------------------------------------------------------------
    
    On Error GoTo GestionErreur
    
    Dim ws As Worksheet
    Dim MoisRapport As Date
    
    Set ws = ThisWorkbook.Worksheets("Rapports")
    MoisRapport = ObtenirMoisCourant()
    
    Application.ScreenUpdating = False
    
    ' Préparation de la feuille de rapport
    Call PreparerFeuilleRapport(ws, "Rapport Mensuel - " & Format(MoisRapport, "mmmm yyyy"))
    
    ' Génération des sections du rapport
    Call CreerSectionResumeMensuel(ws, MoisRapport)
    Call CreerSectionAnalyseDepenses(ws, MoisRapport)
    Call CreerSectionComparaisonBudget(ws, MoisRapport)
    Call CreerSectionRecommandations(ws, MoisRapport)
    Call CreerSectionGraphiquesRapport(ws, MoisRapport)
    
    ' Finalisation du rapport
    Call FinaliserRapport(ws)
    
    Application.ScreenUpdating = True
    
    MsgBox "Rapport mensuel généré avec succès !", vbInformation, "Génération Rapport"
    Call EnregistrerJournal("Rapport mensuel généré", "INFO")
    
    Exit Sub
    
GestionErreur:
    Application.ScreenUpdating = True
    MsgBox "Erreur lors de la génération du rapport: " & Err.Description, vbCritical, "Erreur"
    Call EnregistrerJournal("Erreur génération rapport: " & Err.Description, "ERREUR")
End Sub

Sub PreparerFeuilleRapport(ws As Worksheet, TitreRapport As String)
    '-------------------------------------------------------------------------
    ' Prépare la structure de base de la feuille de rapport
    '-------------------------------------------------------------------------
    
    With ws
        .Cells.Clear
        .Tab.Color = RGB(255, 192, 0)
        
        ' En-tête du rapport
        .Range("A1:J1").Merge
        .Range("A1").Value = TitreRapport
        .Range("A1").Font.Size = 18
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Color = RGB(68, 114, 196)
        .Range("A1").HorizontalAlignment = xlCenter
        
        ' Informations du rapport
        .Range("A2:J2").Merge
        .Range("A2").Value = "Généré le " & Format(Now, "dd/mm/yyyy à hh:mm") & " - Finance Tracker v" & VERSION_APP
        .Range("A2").Font.Size = 10
        .Range("A2").HorizontalAlignment = xlCenter
        .Range("A2").Font.Color = RGB(128, 128, 128)
        
        ' Ligne de séparation
        .Range("A3:J3").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("A3:J3").Borders(xlEdgeBottom).Color = RGB(68, 114, 196)
        .Range("A3:J3").Borders(xlEdgeBottom).Weight = xlMedium
    End With
    
End Sub

'===============================================================================
' SECTIONS DU RAPPORT MENSUEL
'===============================================================================

Sub CreerSectionResumeMensuel(ws As Worksheet, MoisRapport As Date)
    '-------------------------------------------------------------------------
    ' Crée la section de résumé exécutif du mois
    '-------------------------------------------------------------------------
    
    Dim LigneActuelle As Integer
    Dim RevenusMois As Currency, DepensesMois As Currency, EpargneMois As Currency
    Dim BudgetRevenusMois As Currency, BudgetDepensesMois As Currency
    
    LigneActuelle = 5
    
    ' Calcul des métriques principales
    RevenusMois = CalculerRevenusMois(MoisRapport)
    DepensesMois = CalculerDepensesMois(MoisRapport)
    EpargneMois = RevenusMois - DepensesMois
    BudgetRevenusMois = CalculerBudgetRevenusMois(MoisRapport, True)
    BudgetDepensesMois = CalculerBudgetDepensesMois(MoisRapport, True)
    
    With ws
        ' Titre de la section
        .Cells(LigneActuelle, 1).Value = "RÉSUMÉ EXÉCUTIF"
        .Cells(LigneActuelle, 1).Font.Size = 14
        .Cells(LigneActuelle, 1).Font.Bold = True
        .Cells(LigneActuelle, 1).Font.Color = RGB(68, 114, 196)
        
        LigneActuelle = LigneActuelle + 2
        
        ' Tableau des métriques principales
        Call CreerTableauMetriques(ws, LigneActuelle, RevenusMois, DepensesMois, EpargneMois, _
                                   BudgetRevenusMois, BudgetDepensesMois)
        
        LigneActuelle = LigneActuelle + 8
        
        ' Analyse textuelle
        .Cells(LigneActuelle, 1).Value = "ANALYSE:"
        .Cells(LigneActuelle, 1).Font.Bold = True
        LigneActuelle = LigneActuelle + 1
        
        .Cells(LigneActuelle, 1).Value = GenererAnalyseTextuelle(RevenusMois, DepensesMois, EpargneMois, _
                                                                 BudgetRevenusMois, BudgetDepensesMois)
        .Range(.Cells(LigneActuelle, 1), .Cells(LigneActuelle, 8)).Merge
        .Cells(LigneActuelle, 1).WrapText = True
        .Cells(LigneActuelle, 1).VerticalAlignment = xlVAlignTop
        .Rows(LigneActuelle).RowHeight = 60
    End With
    
End Sub

Sub CreerTableauMetriques(ws As Worksheet, LigneDebut As Integer, RevenusMois As Currency, _
                         DepensesMois As Currency, EpargneMois As Currency, _
                         BudgetRevenusMois As Currency, BudgetDepensesMois As Currency)
    '-------------------------------------------------------------------------
    ' Crée le tableau des métriques principales
    '-------------------------------------------------------------------------
    
    Dim PlageTableau As Range
    Set PlageTableau = ws.Range(ws.Cells(LigneDebut, 1), ws.Cells(LigneDebut + 6, 5))
    
    With PlageTableau
        ' En-têtes
        .Cells(1, 1).Value = "MÉTRIQUE"
        .Cells(1, 2).Value = "PRÉVU"
        .Cells(1, 3).Value = "RÉALISÉ"
        .Cells(1, 4).Value = "ÉCART"
        .Cells(1, 5).Value = "PERFORMANCE"
        
        ' Données
        .Cells(2, 1).Value = "Revenus totaux"
        .Cells(2, 2).Value = BudgetRevenusMois
        .Cells(2, 3).Value = RevenusMois
        .Cells(2, 4).Value = RevenusMois - BudgetRevenusMois
        .Cells(2, 5).Value = IIf(BudgetRevenusMois > 0, RevenusMois / BudgetRevenusMois, 0)
        
        .Cells(3, 1).Value = "Dépenses totales"
        .Cells(3, 2).Value = BudgetDepensesMois
        .Cells(3, 3).Value = DepensesMois
        .Cells(3, 4).Value = DepensesMois - BudgetDepensesMois
        .Cells(3, 5).Value = IIf(BudgetDepensesMois > 0, DepensesMois / BudgetDepensesMois, 0)
        
        .Cells(4, 1).Value = "Épargne nette"
        .Cells(4, 2).Value = BudgetRevenusMois - BudgetDepensesMois
        .Cells(4, 3).Value = EpargneMois
        .Cells(4, 4).Value = EpargneMois - (BudgetRevenusMois - BudgetDepensesMois)
        .Cells(4, 5).Value = IIf(RevenusMois > 0, EpargneMois / RevenusMois, 0)
        
        .Cells(5, 1).Value = "Taux d'épargne"
        .Cells(5, 2).Value = IIf(BudgetRevenusMois > 0, (BudgetRevenusMois - BudgetDepensesMois) / BudgetRevenusMois, 0)
        .Cells(5, 3).Value = IIf(RevenusMois > 0, EpargneMois / RevenusMois, 0)
        .Cells(5, 4).Value = .Cells(5, 3).Value - .Cells(5, 2).Value
        .Cells(5, 5).Value = DeterminerPerformanceEpargne(.Cells(5, 3).Value)
        
        ' Formatage
        .Rows(1).Font.Bold = True
        .Rows(1).Interior.Color = RGB(68, 114, 196)
        .Rows(1).Font.Color = RGB(255, 255, 255)
        
        .Columns(2).Resize(, 3).NumberFormat = "#,##0 €"
        .Columns(5).NumberFormat = "0.0%"
        .Rows(5).Range("B1:D1").NumberFormat = "0.0%"
        
        .Borders.LineStyle = xlContinuous
        .Font.Size = 9
    End With
    
End Sub

Sub CreerSectionAnalyseDepenses(ws As Worksheet, MoisRapport As Date)
    '-------------------------------------------------------------------------
    ' Crée la section d'analyse détaillée des dépenses
    '-------------------------------------------------------------------------
    
    Dim LigneActuelle As Integer
    Dim Categories As Variant
    Dim i As Integer
    
    LigneActuelle = 20
    Categories = Array("Logement", "Alimentation", "Transport", "Loisirs", "Santé", "Vêtements", "Autres")
    
    With ws
        ' Titre de la section
        .Cells(LigneActuelle, 1).Value = "ANALYSE DES DÉPENSES PAR CATÉGORIE"
        .Cells(LigneActuelle, 1).Font.Size = 14
        .Cells(LigneActuelle, 1).Font.Bold = True
        .Cells(LigneActuelle, 1).Font.Color = RGB(196, 89, 17)
        
        LigneActuelle = LigneActuelle + 2
        
        ' Tableau d'analyse par catégorie
        Call CreerTableauAnalyseDepenses(ws, LigneActuelle, Categories, MoisRapport)
        
        LigneActuelle = LigneActuelle + UBound(Categories) + 4
        
        ' Top 3 des postes de dépenses
        .Cells(LigneActuelle, 1).Value = "TOP 3 DES POSTES DE DÉPENSES:"
        .Cells(LigneActuelle, 1).Font.Bold = True
        .Cells(LigneActuelle, 1).Font.Color = RGB(196, 89, 17)
        
        LigneActuelle = LigneActuelle + 1
        Call CreerTop3Depenses(ws, LigneActuelle, Categories, MoisRapport)
    End With
    
End Sub

Sub CreerTableauAnalyseDepenses(ws As Worksheet, LigneDebut As Integer, Categories As Variant, MoisRapport As Date)
    '-------------------------------------------------------------------------
    ' Crée le tableau détaillé d'analyse des dépenses
    '-------------------------------------------------------------------------
    
    Dim i As Integer
    Dim MontantCategorie As Currency, BudgetCategorie As Currency
    Dim TotalDepenses As Currency
    
    TotalDepenses = CalculerDepensesMois(MoisRapport)
    
    With ws
        ' En-têtes du tableau
        .Cells(LigneDebut, 1).Value = "CATÉGORIE"
        .Cells(LigneDebut, 2).Value = "BUDGET"
        .Cells(LigneDebut, 3).Value = "RÉALISÉ"
        .Cells(LigneDebut, 4).Value = "ÉCART"
        .Cells(LigneDebut, 5).Value = "% TOTAL"
        .Cells(LigneDebut, 6).Value = "TENDANCE"
        .Cells(LigneDebut, 7).Value = "RECOMMANDATION"
        
        ' Formatage des en-têtes
        .Range(.Cells(LigneDebut, 1), .Cells(LigneDebut, 7)).Font.Bold = True
        .Range(.Cells(LigneDebut, 1), .Cells(LigneDebut, 7)).Interior.Color = RGB(196, 89, 17)
        .Range(.Cells(LigneDebut, 1), .Cells(LigneDebut, 7)).Font.Color = RGB(255, 255, 255)
        
        ' Données par catégorie
        For i = 0 To UBound(Categories)
            BudgetCategorie = ObtenirBudgetCategorie(Categories(i), MoisRapport, True)
            MontantCategorie = ObtenirBudgetCategorie(Categories(i), MoisRapport, False)
            
            .Cells(LigneDebut + i + 1, 1).Value = Categories(i)
            .Cells(LigneDebut + i + 1, 2).Value = BudgetCategorie
            .Cells(LigneDebut + i + 1, 3).Value = MontantCategorie
            .Cells(LigneDebut + i + 1, 4).Value = MontantCategorie - BudgetCategorie
            .Cells(LigneDebut + i + 1, 5).Value = IIf(TotalDepenses > 0, MontantCategorie / TotalDepenses, 0)
            .Cells(LigneDebut + i + 1, 6).Value = CalculerTendanceCategorie(Categories(i), MoisRapport)
            .Cells(LigneDebut + i + 1, 7).Value = GenererRecommandations(Categories(i), BudgetCategorie, MontantCategorie)
        Next i
        
        ' Formatage du tableau
        .Range(.Cells(LigneDebut, 1), .Cells(LigneDebut + UBound(Categories) + 1, 7)).Borders.LineStyle = xlContinuous
        .Columns(2).Resize(, 3).NumberFormat = "#,##0 €"
        .Columns(5).NumberFormat = "0.0%"
        .Font.Size = 9
    End With
    
End Sub

Sub CreerSectionComparaisonBudget(ws As Worksheet, MoisRapport As Date)
    '-------------------------------------------------------------------------
    ' Crée la section de comparaison avec le budget prévu
    '-------------------------------------------------------------------------
    
    Dim LigneActuelle As Integer
    Dim MetriquesPrecedentes As MetriquesFinancieres
    
    LigneActuelle = 40
    MetriquesPrecedentes = CalculerMetriquesFinancieres(3)
    
    With ws
        ' Titre de la section
        .Cells(LigneActuelle, 1).Value = "COMPARAISON ET ÉVOLUTION"
        .Cells(LigneActuelle, 1).Font.Size = 14
        .Cells(LigneActuelle, 1).Font.Bold = True
        .Cells(LigneActuelle, 1).Font.Color = RGB(112, 173, 71)
        
        LigneActuelle = LigneActuelle + 2
        
        ' Tableau de comparaison sur 3 mois
        Call CreerTableauComparaison3Mois(ws, LigneActuelle, MoisRapport)
        
        LigneActuelle = LigneActuelle + 6
        
        ' Métriques de performance
        .Cells(LigneActuelle, 1).Value = "MÉTRIQUES DE PERFORMANCE:"
        .Cells(LigneActuelle, 1).Font.Bold = True
        
        LigneActuelle = LigneActuelle + 1
        .Cells(LigneActuelle, 1).Value = "• Volatilité des revenus: " & Format(MetriquesPrecedentes.VolatiliteRevenus / MetriquesPrecedentes.RevenusMoyens, "0.0%")
        LigneActuelle = LigneActuelle + 1
        .Cells(LigneActuelle, 1).Value = "• Régularité des dépenses: " & Format(MetriquesPrecedentes.VolatiliteDepenses / MetriquesPrecedentes.DepensesMoyennes, "0.0%")
        LigneActuelle = LigneActuelle + 1
        .Cells(LigneActuelle, 1).Value = "• Tendance générale: " & MetriquesPrecedentes.TendanceEvolution
    End With
    
End Sub

Sub CreerTableauComparaison3Mois(ws As Worksheet, LigneDebut As Integer, MoisRapport As Date)
    '-------------------------------------------------------------------------
    ' Crée le tableau de comparaison sur les 3 derniers mois
    '-------------------------------------------------------------------------
    
    Dim i As Integer
    Dim MoisAnalyse As Date
    
    With ws
        ' En-têtes
        .Cells(LigneDebut, 1).Value = "PÉRIODE"
        .Cells(LigneDebut, 2).Value = "REVENUS"
        .Cells(LigneDebut, 3).Value = "DÉPENSES"
        .Cells(LigneDebut, 4).Value = "ÉPARGNE"
        .Cells(LigneDebut, 5).Value = "TAUX ÉPARGNE"
        
        ' Formatage des en-têtes
        .Range(.Cells(LigneDebut, 1), .Cells(LigneDebut, 5)).Font.Bold = True
        .Range(.Cells(LigneDebut, 1), .Cells(LigneDebut, 5)).Interior.Color = RGB(112, 173, 71)
        .Range(.Cells(LigneDebut, 1), .Cells(LigneDebut, 5)).Font.Color = RGB(255, 255, 255)
        
        ' Données des 3 derniers mois
        For i = 2 To 0 Step -1
            MoisAnalyse = DateAdd("m", -i, MoisRapport)
            
            .Cells(LigneDebut + (3 - i), 1).Value = Format(MoisAnalyse, "mmmm yyyy")
            .Cells(LigneDebut + (3 - i), 2).Value = CalculerRevenusMois(MoisAnalyse)
            .Cells(LigneDebut + (3 - i), 3).Value = CalculerDepensesMois(MoisAnalyse)
            .Cells(LigneDebut + (3 - i), 4).Value = CalculerRevenusMois(MoisAnalyse) - CalculerDepensesMois(MoisAnalyse)
            
            Dim RevenusMoisAnalyse As Currency
            RevenusMoisAnalyse = CalculerRevenusMois(MoisAnalyse)
            .Cells(LigneDebut + (3 - i), 5).Value = IIf(RevenusMoisAnalyse > 0, (.Cells(LigneDebut + (3 - i), 4).Value) / RevenusMoisAnalyse, 0)
        Next i
        
        ' Formatage
        .Range(.Cells(LigneDebut, 1), .Cells(LigneDebut + 3, 5)).Borders.LineStyle = xlContinuous
        .Columns(2).Resize(, 3).NumberFormat = "#,##0 €"
        .Columns(5).NumberFormat = "0.0%"
    End With
    
End Sub

Sub CreerSectionRecommandations(ws As Worksheet, MoisRapport As Date)
    '-------------------------------------------------------------------------
    ' Crée la section des recommandations et actions
    '-------------------------------------------------------------------------
    
    Dim LigneActuelle As Integer
    Dim Recommandations As String
    
    LigneActuelle = 55
    
    With ws
        ' Titre de la section
        .Cells(LigneActuelle, 1).Value = "RECOMMANDATIONS ET PLAN D'ACTION"
        .Cells(LigneActuelle, 1).Font.Size = 14
        .Cells(LigneActuelle, 1).Font.Bold = True
        .Cells(LigneActuelle, 1).Font.Color = RGB(255, 192, 0)
        
        LigneActuelle = LigneActuelle + 2
        
        ' Génération des recommandations personnalisées
        Recommandations = GenererRecommandationsPersonnalisees(MoisRapport)
        
        .Cells(LigneActuelle, 1).Value = Recommandations
        .Range(.Cells(LigneActuelle, 1), .Cells(LigneActuelle, 8)).Merge
        .Cells(LigneActuelle, 1).WrapText = True
        .Cells(LigneActuelle, 1).VerticalAlignment = xlVAlignTop
        .Rows(LigneActuelle).RowHeight = 100
        
        LigneActuelle = LigneActuelle + 6
        
        ' Objectifs pour le mois suivant
        .Cells(LigneActuelle, 1).Value = "OBJECTIFS POUR LE MOIS SUIVANT:"
        .Cells(LigneActuelle, 1).Font.Bold = True
        .Cells(LigneActuelle, 1).Font.Color = RGB(255, 192, 0)
        
        LigneActuelle = LigneActuelle + 1
        Call GenererObjectifsMoisSuivant(ws, LigneActuelle, MoisRapport)
    End With
    
End Sub

'===============================================================================
' FONCTIONS DE GENERATION DE CONTENU
'===============================================================================

Function GenererAnalyseTextuelle(RevenusMois As Currency, DepensesMois As Currency, EpargneMois As Currency, _
                                 BudgetRevenusMois As Currency, BudgetDepensesMois As Currency) As String
    '-------------------------------------------------------------------------
    ' Génère une analyse textuelle intelligente de la situation financière
    '-------------------------------------------------------------------------
    
    Dim Analyse As String
    Dim TauxEpargne As Double
    Dim EcartRevenusPourcent As Double, EcartDepensesPourcent As Double
    
    If RevenusMois > 0 Then
        TauxEpargne = EpargneMois / RevenusMois
    End If
    
    If BudgetRevenusMois > 0 Then
        EcartRevenusPourcent = (RevenusMois - BudgetRevenusMois) / BudgetRevenusMois
    End If
    
    If BudgetDepensesMois > 0 Then
        EcartDepensesPourcent = (DepensesMois - BudgetDepensesMois) / BudgetDepensesMois
    End If
    
    Analyse = "SYNTHÈSE: "
    
    ' Analyse de l'épargne
    If TauxEpargne >= 0.2 Then
        Analyse = Analyse & "Excellente performance d'épargne (" & Format(TauxEpargne, "0%") & "). "
    ElseIf TauxEpargne >= 0.1 Then
        Analyse = Analyse & "Bon taux d'épargne (" & Format(TauxEpargne, "0%") & "). "
    ElseIf TauxEpargne > 0 Then
        Analyse = Analyse & "Épargne faible (" & Format(TauxEpargne, "0%") & "), à améliorer. "
    Else
        Analyse = Analyse & "⚠ ALERTE: Situation déficitaire, action immédiate requise. "
    End If
    
    ' Analyse des revenus
    If EcartRevenusPourcent >= 0.05 Then
        Analyse = Analyse & "Revenus supérieurs aux prévisions (+" & Format(EcartRevenusPourcent, "0%") & "). "
    ElseIf EcartRevenusPourcent <= -0.05 Then
        Analyse = Analyse & "Revenus inférieurs aux prévisions (" & Format(EcartRevenusPourcent, "0%") & "). "
    Else
        Analyse = Analyse & "Revenus conformes aux prévisions. "
    End If
    
    ' Analyse des dépenses
    If EcartDepensesPourcent >= 0.1 Then
        Analyse = Analyse & "⚠ Dépassement significatif du budget dépenses (+" & Format(EcartDepensesPourcent, "0%") & ")."
    ElseIf EcartDepensesPourcent >= 0.05 Then
        Analyse = Analyse & "Léger dépassement du budget dépenses (+" & Format(EcartDepensesPourcent, "0%") & ")."
    Else
        Analyse = Analyse & "Dépenses maîtrisées par rapport au budget."
    End If
    
    GenererAnalyseTextuelle = Analyse
    
End Function

Function GenererRecommandationsPersonnalisees(MoisRapport As Date) As String
    '-------------------------------------------------------------------------
    ' Génère des recommandations personnalisées basées sur l'analyse
    '-------------------------------------------------------------------------
    
    Dim Recommandations As String
    Dim RevenusMois As Currency, DepensesMois As Currency
    Dim TauxEpargne As Double
    Dim NbDepassements As Integer
    
    RevenusMois = CalculerRevenusMois(MoisRapport)
    DepensesMois = CalculerDepensesMois(MoisRapport)
    
    If RevenusMois > 0 Then
        TauxEpargne = (RevenusMois - DepensesMois) / RevenusMois
    End If
    
    NbDepassements = CompterDepassementsBudget(MoisRapport)
    
    Recommandations = "RECOMMANDATIONS PRIORITAIRES:" & vbCrLf & vbCrLf
    
    ' Recommandations selon le taux d'épargne
    If TauxEpargne < 0 Then
        Recommandations = Recommandations & "🔴 URGENT - Réduire immédiatement les dépenses non essentielles" & vbCrLf
        Recommandations = Recommandations & "🔴 Chercher des sources de revenus complémentaires" & vbCrLf
    ElseIf TauxEpargne < 0.1 Then
        Recommandations = Recommandations & "🟡 Améliorer le taux d'épargne (objectif: 15% minimum)" & vbCrLf
        Recommandations = Recommandations & "🟡 Réviser le budget des catégories en dépassement" & vbCrLf
    Else
        Recommandations = Recommandations & "🟢 Maintenir la discipline budgétaire actuelle" & vbCrLf
        Recommandations = Recommandations & "🟢 Envisager d'augmenter les investissements" & vbCrLf
    End If
    
    ' Recommandations selon les dépassements
    If NbDepassements > 2 Then
        Recommandations = Recommandations & "🔴 Revoir complètement la répartition budgétaire" & vbCrLf
    ElseIf NbDepassements > 0 Then
        Recommandations = Recommandations & "🟡 Surveiller de près les catégories en dépassement" & vbCrLf
    End If
    
    ' Recommandations générales
    Recommandations = Recommandations & vbCrLf & "ACTIONS RECOMMANDÉES:" & vbCrLf
    Recommandations = Recommandations & "• Réviser le budget du mois prochain selon les réalisations" & vbCrLf
    Recommandations = Recommandations & "• Automatiser l'épargne (virement automatique)" & vbCrLf
    Recommandations = Recommandations & "• Suivre quotidiennement les dépenses importantes" & vbCrLf
    
    GenererRecommandationsPersonnalisees = Recommandations
    
End Function

Sub GenererObjectifsMoisSuivant(ws As Worksheet, LigneDebut As Integer, MoisRapport As Date)
    '-------------------------------------------------------------------------
    ' Génère les objectifs suggérés pour le mois suivant
    '-------------------------------------------------------------------------
    
    Dim RevenusMoisActuel As Currency, DepensesMoisActuel As Currency
    Dim ObjectifRevenus As Currency, ObjectifDepenses As Currency, ObjectifEpargne As Currency
    
    RevenusMoisActuel = CalculerRevenusMois(MoisRapport)
    DepensesMoisActuel = CalculerDepensesMois(MoisRapport)
    
    ' Calcul des objectifs basés sur les performances actuelles
    ObjectifRevenus = RevenusMoisActuel * 1.02 ' +2%
    ObjectifDepenses = DepensesMoisActuel * 0.98 ' -2%
    ObjectifEpargne = ObjectifRevenus - ObjectifDepenses
    
    With ws
        .Cells(LigneDebut, 1).Value = "OBJECTIF"
        .Cells(LigneDebut, 2).Value = "MONTANT CIBLE"
        .Cells(LigneDebut, 3).Value = "ÉVOLUTION"
        
        .Cells(LigneDebut + 1, 1).Value = "Revenus minimum"
        .Cells(LigneDebut + 1, 2).Value = ObjectifRevenus
        .Cells(LigneDebut + 1, 3).Value = "+2%"
        
        .Cells(LigneDebut + 2, 1).Value = "Dépenses maximum"
        .Cells(LigneDebut + 2, 2).Value = ObjectifDepenses
        .Cells(LigneDebut + 2, 3).Value = "-2%"
        
        .Cells(LigneDebut + 3, 1).Value = "Épargne visée"
        .Cells(LigneDebut + 3, 2).Value = ObjectifEpargne
        .Cells(LigneDebut + 3, 3).Value = Format(ObjectifEpargne / ObjectifRevenus, "0%")
        
        ' Formatage
        .Range(.Cells(LigneDebut, 1), .Cells(LigneDebut, 3)).Font.Bold = True
        .Range(.Cells(LigneDebut, 1), .Cells(LigneDebut + 3, 3)).Borders.LineStyle = xlContinuous
        .Columns(2).NumberFormat = "#,##0 €"
    End With
    
End Sub

'===============================================================================
' FONCTIONS UTILITAIRES POUR LES RAPPORTS
'===============================================================================

Function DeterminerPerformanceEpargne(TauxEpargne As Double) As String
    '-------------------------------------------------------------------------
    ' Détermine la performance du taux d'épargne
    '-------------------------------------------------------------------------
    
    Select Case TauxEpargne
        Case Is >= 0.25
            DeterminerPerformanceEpargne = "Excellent"
        Case 0.15 To 0.24
            DeterminerPerformanceEpargne = "Très bon"
        Case 0.1 To 0.14
            DeterminerPerformanceEpargne = "Correct"
        Case 0.05 To 0.09
            DeterminerPerformanceEpargne = "Faible"
        Case Else
            DeterminerPerformanceEpargne = "Insuffisant"
    End Select
    
End Function

Sub CreerTop3Depenses(ws As Worksheet, LigneDebut As Integer, Categories As Variant, MoisRapport As Date)
    '-------------------------------------------------------------------------
    ' Crée le classement des 3 principales catégories de dépenses
    '-------------------------------------------------------------------------
    
    Dim i As Integer, j As Integer
    Dim MontantsCategories() As Currency
    Dim CategoriesTriees() As String
    Dim TempMontant As Currency, TempCategorie As String
    
    ReDim MontantsCategories(UBound(Categories))
    ReDim CategoriesTriees(UBound(Categories))
    
    ' Récupération des montants
    For i = 0 To UBound(Categories)
        MontantsCategories(i) = ObtenirBudgetCategorie(Categories(i), MoisRapport, False)
        CategoriesTriees(i) = Categories(i)
    Next i
    
    ' Tri par ordre décroissant (bubble sort simple)
    For i = 0 To UBound(MontantsCategories) - 1
        For j = i + 1 To UBound(MontantsCategories)
            If MontantsCategories(i) < MontantsCategories(j) Then
                TempMontant = MontantsCategories(i)
                TempCategorie = CategoriesTriees(i)
                MontantsCategories(i) = MontantsCategories(j)
                CategoriesTriees(i) = CategoriesTriees(j)
                MontantsCategories(j) = TempMontant
                CategoriesTriees(j) = TempCategorie
            End If
        Next j
    Next i
    
    ' Affichage du top 3
    With ws
        For i = 0 To 2
            If i <= UBound(CategoriesTriees) Then
                .Cells(LigneDebut + i, 1).Value = (i + 1) & ". " & CategoriesTriees(i)
                .Cells(LigneDebut + i, 2).Value = MontantsCategories(i)
                .Cells(LigneDebut + i, 2).NumberFormat = "#,##0 €"
            End If
        Next i
    End With
    
End Sub

Sub FinaliserRapport(ws As Worksheet)
    '-------------------------------------------------------------------------
    ' Finalise la mise en forme du rapport
    '-------------------------------------------------------------------------
    
    With ws
        ' Police générale
        .Cells.Font.Name = "Segoe UI"
        .Cells.Font.Size = 9
        
        ' Ajustement automatique des colonnes
        .Columns("A:J").AutoFit
        
        ' Protection du rapport
        .Protect Password:="FinanceTracker2025", _
                DrawingObjects:=False, _
                Contents:=True, _
                Scenarios:=False
    End With
    
End Sub

Sub CreerSectionGraphiquesRapport(ws As Worksheet, MoisRapport As Date)
    '-------------------------------------------------------------------------
    ' Ajoute les graphiques au rapport
    '-------------------------------------------------------------------------
    
    Dim LigneActuelle As Integer
    
    LigneActuelle = 75
    
    With ws
        .Cells(LigneActuelle, 1).Value = "VISUALISATIONS GRAPHIQUES"
        .Cells(LigneActuelle, 1).Font.Size = 14
        .Cells(LigneActuelle, 1).Font.Bold = True
        .Cells(LigneActuelle, 1).Font.Color = RGB(68, 114, 196)
        
        LigneActuelle = LigneActuelle + 2
        
        ' Préparation des données pour graphiques
        Call PreparerDonneesGraphiquesRapport(ws, MoisRapport)
        
        ' Création des graphiques de rapport
        Call CreerGraphiqueRapport(ws, EvolutionMensuelle, "A" & LigneActuelle & ":E" & (LigneActuelle + 15), "L1:P13")
        Call CreerGraphiqueRapport(ws, RepartitionDepenses, "F" & LigneActuelle & ":J" & (LigneActuelle + 15), "R1:S8")
    End With
    
End Sub

Sub PreparerDonneesGraphiquesRapport(ws As Worksheet, MoisRapport As Date)
    '-------------------------------------------------------------------------
    ' Prépare les données spécifiques pour les graphiques de rapport
    '-------------------------------------------------------------------------
    
    Dim i As Integer
    Dim MoisAnalyse As Date
    
    ' Données d'évolution (colonne L à P)
    With ws
        .Range("L1").Value = "Mois"
        .Range("M1").Value = "Revenus"
        .Range("N1").Value = "Dépenses"
        .Range("O1").Value = "Épargne"
        .Range("P1").Value = "Objectif Épargne"
        
        For i = 11 To 0 Step -1
            MoisAnalyse = DateAdd("m", -i, MoisRapport)
            .Range("L" & (13 - i)).Value = Format(MoisAnalyse, "mmm")
            .Range("M" & (13 - i)).Value = CalculerRevenusMois(MoisAnalyse)
            .Range("N" & (13 - i)).Value = CalculerDepensesMois(MoisAnalyse)
            .Range("O" & (13 - i)).Value = CalculerRevenusMois(MoisAnalyse) - CalculerDepensesMois(MoisAnalyse)
            .Range("P" & (13 - i)).Value = CalculerRevenusMois(MoisAnalyse) * 0.2 ' Objectif 20%
        Next i
        
        ' Données de répartition (colonne R à S)
        Dim Categories As Variant
        Categories = Array("Logement", "Alimentation", "Transport", "Loisirs", "Santé", "Vêtements", "Autres")
        
        .Range("R1").Value = "Catégorie"
        .Range("S1").Value = "Montant"
        
        For i = 0 To UBound(Categories)
            .Range("R" & (i + 2)).Value = Categories(i)
            .Range("S" & (i + 2)).Value = ObtenirBudgetCategorie(Categories(i), MoisRapport, False)
        Next i
    End With
    
End Sub

'===============================================================================
' FIN DU MODULE RAPPORTS
'===============================================================================
