Attribute VB_Name = "Module_Calculs"
'===============================================================================
' FINANCE TRACKER VBA - MODULE CALCULS FINANCIERS
' Version: 1.0
' Description: Moteur de calculs et analyses financières avancées
' Fonction: Centralisation de tous les calculs et métriques financières
'===============================================================================

Option Explicit

' Constantes pour les calculs financiers
Private Const TAUX_EPARGNE_RECOMMANDE As Double = 0.2
Private Const SEUIL_ALERTE_BUDGET As Double = 0.9
Private Const NB_MOIS_HISTORIQUE As Integer = 12

' Structure pour les métriques financières
Public Type MetriquesFinancieres
    RevenusMoyens As Currency
    DepensesMoyennes As Currency
    EpargneMoyenne As Currency
    TauxEpargne As Double
    VolatiliteRevenus As Double
    VolatiliteDepenses As Double
    TendanceEvolution As String
End Type

'===============================================================================
' FONCTIONS DE CALCUL PRINCIPALES
'===============================================================================

Function CalculerIndicateur(TypeIndicateur As String) As String
    '-------------------------------------------------------------------------
    ' Calcule et formate un indicateur spécifique du tableau de bord
    '-------------------------------------------------------------------------
    
    Dim MoisActuel As Date
    Dim Resultat As Currency
    
    MoisActuel = ObtenirMoisCourant()
    
    Select Case UCase(TypeIndicateur)
        Case "REVENUS DU MOIS"
            Resultat = CalculerRevenusMois(MoisActuel)
            
        Case "DÉPENSES DU MOIS"
            Resultat = CalculerDepensesMois(MoisActuel)
            
        Case "ÉPARGNE RÉALISÉE"
            Resultat = CalculerRevenusMois(MoisActuel) - CalculerDepensesMois(MoisActuel)
            
        Case "BUDGET RESTANT"
            Resultat = CalculerBudgetRestant(MoisActuel)
            
        Case Else
            Resultat = 0
    End Select
    
    CalculerIndicateur = FormaterMontant(Resultat)
    
End Function

Function CalculerResume(Categorie As String, ColonneIndex As Integer) As Variant
    '-------------------------------------------------------------------------
    ' Calcule les valeurs du tableau de résumé mensuel détaillé
    '-------------------------------------------------------------------------
    
    Dim MoisActuel As Date
    Dim BudgetPrevu As Currency, MontantReel As Currency
    Dim Ecart As Currency, EcartPourcent As Double
    
    MoisActuel = ObtenirMoisCourant()
    
    ' Récupération des montants selon la catégorie
    BudgetPrevu = ObtenirBudgetCategorie(Categorie, MoisActuel, True)
    MontantReel = ObtenirBudgetCategorie(Categorie, MoisActuel, False)
    
    Select Case ColonneIndex
        Case 2: ' Budget prévu
            CalculerResume = FormaterMontant(BudgetPrevu)
            
        Case 3: ' Montant réel
            CalculerResume = FormaterMontant(MontantReel)
            
        Case 4: ' Écart absolu
            Ecart = MontantReel - BudgetPrevu
            CalculerResume = FormaterMontant(Ecart)
            
        Case 5: ' Écart pourcentage
            If BudgetPrevu <> 0 Then
                EcartPourcent = (MontantReel - BudgetPrevu) / BudgetPrevu
                CalculerResume = Format(EcartPourcent, "0.0%")
            Else
                CalculerResume = "N/A"
            End If
            
        Case 6: ' Statut
            CalculerResume = DeterminerStatutBudget(BudgetPrevu, MontantReel)
            
        Case 7: ' Tendance
            CalculerResume = CalculerTendanceCategorie(Categorie, MoisActuel)
            
        Case 8: ' Actions recommandées
            CalculerResume = GenererRecommandations(Categorie, BudgetPrevu, MontantReel)
            
        Case Else
            CalculerResume = ""
    End Select
    
End Function

'===============================================================================
' FONCTIONS DE CALCUL AVANCÉES
'===============================================================================

Function CalculerMetriquesFinancieres(NbMoisAnalyse As Integer) As MetriquesFinancieres
    '-------------------------------------------------------------------------
    ' Calcule les métriques financières sur une période donnée
    '-------------------------------------------------------------------------
    
    Dim Metriques As MetriquesFinancieres
    Dim i As Integer
    Dim MoisAnalyse As Date
    Dim TotalRevenus As Currency, TotalDepenses As Currency
    Dim RevenusMensuels() As Currency, DepensesMensuelles() As Currency
    
    ReDim RevenusMensuels(1 To NbMoisAnalyse)
    ReDim DepensesMensuelles(1 To NbMoisAnalyse)
    
    ' Collecte des données sur la période
    For i = 1 To NbMoisAnalyse
        MoisAnalyse = DateAdd("m", -(NbMoisAnalyse - i), ObtenirMoisCourant())
        RevenusMensuels(i) = CalculerRevenusMois(MoisAnalyse)
        DepensesMensuelles(i) = CalculerDepensesMois(MoisAnalyse)
        TotalRevenus = TotalRevenus + RevenusMensuels(i)
        TotalDepenses = TotalDepenses + DepensesMensuelles(i)
    Next i
    
    ' Calcul des moyennes
    Metriques.RevenusMoyens = TotalRevenus / NbMoisAnalyse
    Metriques.DepensesMoyennes = TotalDepenses / NbMoisAnalyse
    Metriques.EpargneMoyenne = Metriques.RevenusMoyens - Metriques.DepensesMoyennes
    
    ' Calcul du taux d'épargne
    If Metriques.RevenusMoyens > 0 Then
        Metriques.TauxEpargne = Metriques.EpargneMoyenne / Metriques.RevenusMoyens
    End If
    
    ' Calcul de la volatilité
    Metriques.VolatiliteRevenus = CalculerVolatilite(RevenusMensuels, Metriques.RevenusMoyens)
    Metriques.VolatiliteDepenses = CalculerVolatilite(DepensesMensuelles, Metriques.DepensesMoyennes)
    
    ' Détermination de la tendance
    Metriques.TendanceEvolution = DeterminerTendanceEvolution(RevenusMensuels, DepensesMensuelles)
    
    CalculerMetriquesFinancieres = Metriques
    
End Function

Function CalculerVolatilite(Valeurs() As Currency, Moyenne As Currency) As Double
    '-------------------------------------------------------------------------
    ' Calcule la volatilité (écart-type) d'une série de valeurs
    '-------------------------------------------------------------------------
    
    Dim i As Integer
    Dim SommeCarresEcarts As Double
    Dim NbValeurs As Integer
    
    NbValeurs = UBound(Valeurs) - LBound(Valeurs) + 1
    
    For i = LBound(Valeurs) To UBound(Valeurs)
        SommeCarresEcarts = SommeCarresEcarts + (Valeurs(i) - Moyenne) ^ 2
    Next i
    
    If NbValeurs > 1 Then
        CalculerVolatilite = Sqr(SommeCarresEcarts / (NbValeurs - 1))
    Else
        CalculerVolatilite = 0
    End If
    
End Function

Function DeterminerTendanceEvolution(RevenusMensuels() As Currency, DepensesMensuelles() As Currency) As String
    '-------------------------------------------------------------------------
    ' Détermine la tendance d'évolution financière
    '-------------------------------------------------------------------------
    
    Dim TendanceRevenus As String, TendanceDepenses As String
    Dim NbMois As Integer
    
    NbMois = UBound(RevenusMensuels) - LBound(RevenusMensuels) + 1
    
    If NbMois < 3 Then
        DeterminerTendanceEvolution = "Données insuffisantes"
        Exit Function
    End If
    
    ' Analyse de la tendance des revenus
    TendanceRevenus = AnalyserTendance(RevenusMensuels)
    
    ' Analyse de la tendance des dépenses
    TendanceDepenses = AnalyserTendance(DepensesMensuelles)
    
    ' Synthèse de la tendance globale
    If TendanceRevenus = "Croissante" And TendanceDepenses = "Décroissante" Then
        DeterminerTendanceEvolution = "Amélioration forte"
    ElseIf TendanceRevenus = "Croissante" And TendanceDepenses = "Stable" Then
        DeterminerTendanceEvolution = "Amélioration modérée"
    ElseIf TendanceRevenus = "Stable" And TendanceDepenses = "Décroissante" Then
        DeterminerTendanceEvolution = "Optimisation réussie"
    ElseIf TendanceRevenus = "Stable" And TendanceDepenses = "Stable" Then
        DeterminerTendanceEvolution = "Situation stable"
    ElseIf TendanceRevenus = "Décroissante" And TendanceDepenses = "Croissante" Then
        DeterminerTendanceEvolution = "Dégradation préoccupante"
    Else
        DeterminerTendanceEvolution = "Évolution mixte"
    End If
    
End Function

Function AnalyserTendance(Valeurs() As Currency) As String
    '-------------------------------------------------------------------------
    ' Analyse la tendance d'une série de valeurs
    '-------------------------------------------------------------------------
    
    Dim i As Integer
    Dim NbHausse As Integer, NbBaisse As Integer
    Dim SeuilVariation As Double
    
    SeuilVariation = 0.05 ' 5% de variation considérée comme significative
    
    For i = LBound(Valeurs) + 1 To UBound(Valeurs)
        If Valeurs(i) > Valeurs(i - 1) * (1 + SeuilVariation) Then
            NbHausse = NbHausse + 1
        ElseIf Valeurs(i) < Valeurs(i - 1) * (1 - SeuilVariation) Then
            NbBaisse = NbBaisse + 1
        End If
    Next i
    
    If NbHausse > NbBaisse Then
        AnalyserTendance = "Croissante"
    ElseIf NbBaisse > NbHausse Then
        AnalyserTendance = "Décroissante"
    Else
        AnalyserTendance = "Stable"
    End If
    
End Function

'===============================================================================
' FONCTIONS DE CALCUL BUDGÉTAIRE
'===============================================================================

Function ObtenirBudgetCategorie(Categorie As String, MoisReference As Date, EstPrevu As Boolean) As Currency
    '-------------------------------------------------------------------------
    ' Récupère le budget d'une catégorie pour un mois donné
    '-------------------------------------------------------------------------
    
    Dim ws As Worksheet
    Dim DerniereLigne As Long
    Dim i As Long
    Dim ColonneMontant As Integer
    
    ' Détermination de la feuille selon le type de catégorie
    If EstCategorieRevenu(Categorie) Then
        Set ws = ThisWorkbook.Worksheets("Donnees_Revenus")
    Else
        Set ws = ThisWorkbook.Worksheets("Donnees_Depenses")
    End If
    
    If ws.Cells(1, 1).Value = "" Then
        ObtenirBudgetCategorie = 0
        Exit Function
    End If
    
    ' Détermination de la colonne selon le type de montant
    If EstPrevu Then
        ColonneMontant = 3 ' Colonne montant prévu
    Else
        ColonneMontant = 4 ' Colonne montant réel
    End If
    
    DerniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To DerniereLigne
        If ws.Cells(i, 2).Value = Categorie And _
           Month(ws.Cells(i, 1).Value) = Month(MoisReference) And _
           Year(ws.Cells(i, 1).Value) = Year(MoisReference) Then
            ObtenirBudgetCategorie = ws.Cells(i, ColonneMontant).Value
            Exit Function
        End If
    Next i
    
    ObtenirBudgetCategorie = 0
    
End Function

Function EstCategorieRevenu(Categorie As String) As Boolean
    '-------------------------------------------------------------------------
    ' Détermine si une catégorie est un revenu ou une dépense
    '-------------------------------------------------------------------------
    
    Dim CategoriesRevenus As Variant
    Dim i As Integer
    
    CategoriesRevenus = Array("Revenus Salaire", "Salaire principal", "Salaire conjoint", _
                             "Primes/Bonus", "Revenus locatifs", "Investissements", _
                             "Revenus Autres", "Autres revenus")
    
    For i = 0 To UBound(CategoriesRevenus)
        If InStr(1, UCase(Categorie), UCase(CategoriesRevenus(i))) > 0 Then
            EstCategorieRevenu = True
            Exit Function
        End If
    Next i
    
    EstCategorieRevenu = False
    
End Function

Function ObtenirBudgetMensuelTotal() As Currency
    '-------------------------------------------------------------------------
    ' Calcule le budget mensuel total prévu
    '-------------------------------------------------------------------------
    
    Dim BudgetRevenus As Currency, BudgetDepenses As Currency
    
    BudgetRevenus = CalculerBudgetRevenusMois(ObtenirMoisCourant(), True)
    BudgetDepenses = CalculerBudgetDepensesMois(ObtenirMoisCourant(), True)
    
    ObtenirBudgetMensuelTotal = BudgetRevenus - BudgetDepenses
    
End Function

Function CalculerBudgetRevenusMois(MoisReference As Date, EstPrevu As Boolean) As Currency
    '-------------------------------------------------------------------------
    ' Calcule le budget total des revenus pour un mois
    '-------------------------------------------------------------------------
    
    Dim ws As Worksheet
    Dim DerniereLigne As Long
    Dim i As Long
    Dim ColonneMontant As Integer
    Dim Total As Currency
    
    Set ws = ThisWorkbook.Worksheets("Donnees_Revenus")
    
    If ws.Cells(1, 1).Value = "" Then
        CalculerBudgetRevenusMois = 0
        Exit Function
    End If
    
    If EstPrevu Then
        ColonneMontant = 3
    Else
        ColonneMontant = 4
    End If
    
    DerniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To DerniereLigne
        If Month(ws.Cells(i, 1).Value) = Month(MoisReference) And _
           Year(ws.Cells(i, 1).Value) = Year(MoisReference) Then
            Total = Total + ws.Cells(i, ColonneMontant).Value
        End If
    Next i
    
    CalculerBudgetRevenusMois = Total
    
End Function

Function CalculerBudgetDepensesMois(MoisReference As Date, EstPrevu As Boolean) As Currency
    '-------------------------------------------------------------------------
    ' Calcule le budget total des dépenses pour un mois
    '-------------------------------------------------------------------------
    
    Dim ws As Worksheet
    Dim DerniereLigne As Long
    Dim i As Long
    Dim ColonneMontant As Integer
    Dim Total As Currency
    
    Set ws = ThisWorkbook.Worksheets("Donnees_Depenses")
    
    If ws.Cells(1, 1).Value = "" Then
        CalculerBudgetDepensesMois = 0
        Exit Function
    End If
    
    If EstPrevu Then
        ColonneMontant = 3
    Else
        ColonneMontant = 4
    End If
    
    DerniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To DerniereLigne
        If Month(ws.Cells(i, 1).Value) = Month(MoisReference) And _
           Year(ws.Cells(i, 1).Value) = Year(MoisReference) Then
            Total = Total + ws.Cells(i, ColonneMontant).Value
        End If
    Next i
    
    CalculerBudgetDepensesMois = Total
    
End Function

'===============================================================================
' FONCTIONS D'ANALYSE ET RECOMMANDATIONS
'===============================================================================

Function DeterminerStatutBudget(BudgetPrevu As Currency, MontantReel As Currency) As String
    '-------------------------------------------------------------------------
    ' Détermine le statut d'un budget selon l'écart réel/prévu
    '-------------------------------------------------------------------------
    
    Dim RatioRealisation As Double
    
    If BudgetPrevu = 0 Then
        If MontantReel = 0 Then
            DeterminerStatutBudget = "Non applicable"
        Else
            DeterminerStatutBudget = "Hors budget"
        End If
        Exit Function
    End If
    
    RatioRealisation = MontantReel / BudgetPrevu
    
    Select Case RatioRealisation
        Case Is < 0.8
            DeterminerStatutBudget = "Sous-réalisé"
        Case 0.8 To 1.1
            DeterminerStatutBudget = "Conforme"
        Case 1.1 To 1.2
            DeterminerStatutBudget = "Léger dépassement"
        Case Is > 1.2
            DeterminerStatutBudget = "Dépassement important"
    End Select
    
End Function

Function CalculerTendanceCategorie(Categorie As String, MoisReference As Date) As String
    '-------------------------------------------------------------------------
    ' Calcule la tendance d'évolution d'une catégorie sur 3 mois
    '-------------------------------------------------------------------------
    
    Dim MontantsMois(1 To 3) As Currency
    Dim i As Integer
    Dim MoisAnalyse As Date
    
    ' Récupération des 3 derniers mois
    For i = 1 To 3
        MoisAnalyse = DateAdd("m", -(3 - i), MoisReference)
        MontantsMois(i) = ObtenirBudgetCategorie(Categorie, MoisAnalyse, False)
    Next i
    
    ' Analyse de la tendance
    If MontantsMois(3) > MontantsMois(2) And MontantsMois(2) > MontantsMois(1) Then
        CalculerTendanceCategorie = "↗ Hausse"
    ElseIf MontantsMois(3) < MontantsMois(2) And MontantsMois(2) < MontantsMois(1) Then
        CalculerTendanceCategorie = "↘ Baisse"
    ElseIf Abs(MontantsMois(3) - MontantsMois(1)) / MontantsMois(1) < 0.1 Then
        CalculerTendanceCategorie = "→ Stable"
    Else
        CalculerTendanceCategorie = "↕ Variable"
    End If
    
End Function

Function GenererRecommandations(Categorie As String, BudgetPrevu As Currency, MontantReel As Currency) As String
    '-------------------------------------------------------------------------
    ' Génère des recommandations personnalisées selon l'écart budgétaire
    '-------------------------------------------------------------------------
    
    Dim EcartPourcent As Double
    Dim Recommandation As String
    
    If BudgetPrevu = 0 Then
        GenererRecommandations = "Définir un budget"
        Exit Function
    End If
    
    EcartPourcent = (MontantReel - BudgetPrevu) / BudgetPrevu
    
    Select Case EcartPourcent
        Case Is < -0.2
            Recommandation = "Augmenter le budget"
        Case -0.2 To -0.05
            Recommandation = "Surveiller sous-réalisation"
        Case -0.05 To 0.05
            Recommandation = "Maintenir le cap"
        Case 0.05 To 0.15
            Recommandation = "Attention au dépassement"
        Case Is > 0.15
            Recommandation = "Réduire les dépenses"
    End Select
    
    GenererRecommandations = Recommandation
    
End Function

Function GenererAlertes() As String
    '-------------------------------------------------------------------------
    ' Génère les alertes automatiques basées sur l'analyse financière
    '-------------------------------------------------------------------------
    
    Dim Alertes As String
    Dim MoisActuel As Date
    Dim TauxEpargne As Double
    Dim DepassementsBudget As Integer
    
    MoisActuel = ObtenirMoisCourant()
    Alertes = ""
    
    ' Vérification du taux d'épargne
    Dim RevenusMois As Currency, DepensesMois As Currency
    RevenusMois = CalculerRevenusMois(MoisActuel)
    DepensesMois = CalculerDepensesMois(MoisActuel)
    
    If RevenusMois > 0 Then
        TauxEpargne = (RevenusMois - DepensesMois) / RevenusMois
        
        If TauxEpargne < 0 Then
            Alertes = Alertes & "⚠ ALERTE CRITIQUE: Déficit budgétaire détecté" & vbCrLf
        ElseIf TauxEpargne < 0.1 Then
            Alertes = Alertes & "⚠ Taux d'épargne faible (" & Format(TauxEpargne, "0.0%") & ")" & vbCrLf
        End If
    End If
    
    ' Vérification des dépassements de budget
    DepassementsBudget = CompterDepassementsBudget(MoisActuel)
    If DepassementsBudget > 0 Then
        Alertes = Alertes & "⚠ " & DepassementsBudget & " catégorie(s) en dépassement budgétaire" & vbCrLf
    End If
    
    ' Message par défaut si aucune alerte
    If Alertes = "" Then
        Alertes = "✓ Situation financière stable - Aucune alerte active"
    End If
    
    GenererAlertes = Alertes
    
End Function

Function CompterDepassementsBudget(MoisReference As Date) As Integer
    '-------------------------------------------------------------------------
    ' Compte le nombre de catégories en dépassement budgétaire
    '-------------------------------------------------------------------------
    
    Dim Categories As Variant
    Dim i As Integer
    Dim BudgetPrevu As Currency, MontantReel As Currency
    Dim NbDepassements As Integer
    
    Categories = Array("Logement", "Alimentation", "Transport", "Loisirs", "Santé", "Vêtements")
    
    For i = 0 To UBound(Categories)
        BudgetPrevu = ObtenirBudgetCategorie(Categories(i), MoisReference, True)
        MontantReel = ObtenirBudgetCategorie(Categories(i), MoisReference, False)
        
        If BudgetPrevu > 0 And MontantReel > BudgetPrevu * SEUIL_ALERTE_BUDGET Then
            NbDepassements = NbDepassements + 1
        End If
    Next i
    
    CompterDepassementsBudget = NbDepassements
    
End Function

'===============================================================================
' FONCTIONS DE PROJECTION ET PRÉVISION
'===============================================================================

Function ProjectionEpargneSixMois() As Currency
    '-------------------------------------------------------------------------
    ' Projette l'épargne cumulée sur les 6 prochains mois
    '-------------------------------------------------------------------------
    
    Dim MetriquesTroisMois As MetriquesFinancieres
    Dim EpargneMensuelleProjectee As Currency
    
    MetriquesTroisMois = CalculerMetriquesFinancieres(3)
    EpargneMensuelleProjectee = MetriquesTroisMois.EpargneMoyenne
    
    ProjectionEpargneSixMois = EpargneMensuelleProjectee * 6
    
End Function

Function EstimerBudgetOptimal(Categorie As String) As Currency
    '-------------------------------------------------------------------------
    ' Estime un budget optimal basé sur l'historique et les bonnes pratiques
    '-------------------------------------------------------------------------
    
    Dim MoyenneHistorique As Currency
    Dim VolatiliteHistorique As Double
    Dim BudgetOptimal As Currency
    Dim i As Integer
    Dim MoisAnalyse As Date
    Dim Total As Currency
    
    ' Calcul de la moyenne sur 6 mois
    For i = 1 To 6
        MoisAnalyse = DateAdd("m", -i, ObtenirMoisCourant())
        Total = Total + ObtenirBudgetCategorie(Categorie, MoisAnalyse, False)
    Next i
    
    MoyenneHistorique = Total / 6
    
    ' Application d'une marge de sécurité de 10%
    BudgetOptimal = MoyenneHistorique * 1.1
    
    EstimerBudgetOptimal = BudgetOptimal
    
End Function

'===============================================================================
' FIN DU MODULE CALCULS
'===============================================================================
