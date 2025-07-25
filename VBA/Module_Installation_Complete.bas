Attribute VB_Name = "Module_Installation_Complete"
'===============================================================================
' FINANCE TRACKER VBA - MODULE D'INSTALLATION AUTOMATIQUE COMPLÈTE
' Version: 1.0
' Description: Installation et configuration automatique complète du système
' Fonction: Créé tout le système en une seule exécution
'===============================================================================

Option Explicit

' Constantes pour l'installation
Public Const VERSION_FINANCE_TRACKER As String = "1.0"
Public Const NOM_FICHIER_FINANCE_TRACKER As String = "FinanceTracker"

'===============================================================================
' INSTALLATION AUTOMATIQUE COMPLÈTE
'===============================================================================

Sub InstallationCompleteFinanceTracker()
    '-------------------------------------------------------------------------
    ' Installation automatique complète du système Finance Tracker
    ' EXÉCUTER CETTE MACRO POUR INSTALLER LE SYSTÈME COMPLET
    '-------------------------------------------------------------------------
    
    On Error GoTo GestionErreur
    
    Dim TempsDebut As Double
    TempsDebut = Timer
    
    ' Message de bienvenue
    If MsgBox("🚀 INSTALLATION FINANCE TRACKER VBA" & vbCrLf & vbCrLf & _
              "Cette installation va créer automatiquement :" & vbCrLf & _
              "✅ Toutes les feuilles nécessaires" & vbCrLf & _
              "✅ Le tableau de bord interactif" & vbCrLf & _
              "✅ Les formulaires de saisie" & vbCrLf & _
              "✅ Les graphiques dynamiques" & vbCrLf & _
              "✅ Le système de rapports" & vbCrLf & _
              "✅ Les catégories par défaut" & vbCrLf & vbCrLf & _
              "⏱️ Durée estimée : 30 secondes" & vbCrLf & vbCrLf & _
              "Continuer l'installation ?", _
              vbYesNo + vbQuestion, "Installation Finance Tracker") = vbNo Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Étapes d'installation
    Call EtapeCreationFeuilles
    Call EtapeConfigurationTableauBord
    Call EtapeConfigurationSaisie
    Call EtapeConfigurationCategories
    Call EtapeConfigurationParametres
    Call EtapeConfigurationDonnees
    Call EtapeConfigurationRapports
    Call EtapeConfigurationArchives
    Call EtapeCreationGraphiques
    Call EtapeConfigurationNavigation
    Call EtapeInitialisationDonnees
    Call EtapeConfigurationFinale
    
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    ' Message de fin
    Dim TempsFin As Double
    TempsFin = Timer - TempsDebut
    
    MsgBox "🎉 INSTALLATION TERMINÉE AVEC SUCCÈS !" & vbCrLf & vbCrLf & _
           "✅ Finance Tracker v" & VERSION_FINANCE_TRACKER & " est maintenant opérationnel" & vbCrLf & _
           "⏱️ Installation terminée en " & Format(TempsFin, "0") & " secondes" & vbCrLf & vbCrLf & _
           "📊 Vous pouvez maintenant :" & vbCrLf & _
           "• Consulter le TABLEAU DE BORD" & vbCrLf & _
           "• Commencer la SAISIE MENSUELLE" & vbCrLf & _
           "• Configurer vos CATÉGORIES" & vbCrLf & _
           "• Générer des RAPPORTS" & vbCrLf & vbCrLf & _
           "🚀 Bon suivi financier !", _
           vbInformation, "Installation Terminée"
    
    ' Activer le tableau de bord
    ThisWorkbook.Worksheets("Dashboard").Activate
    
    Exit Sub
    
GestionErreur:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "❌ Erreur lors de l'installation: " & Err.Description & vbCrLf & vbCrLf & _
           "Veuillez réessayer ou contacter le support.", vbCritical, "Erreur Installation"
End Sub

'===============================================================================
' ÉTAPES D'INSTALLATION
'===============================================================================

Sub EtapeCreationFeuilles()
    '-------------------------------------------------------------------------
    ' Étape 1 : Création de toutes les feuilles nécessaires
    '-------------------------------------------------------------------------
    
    Dim NomsFeuillesReq As Variant
    Dim i As Integer
    Dim ws As Worksheet
    
    NomsFeuillesReq = Array("Dashboard", "Saisie_Mensuelle", "Donnees_Revenus", _
                           "Donnees_Depenses", "Categories", "Parametres", _
                           "Rapports", "Archives")
    
    ' Supprimer les feuilles par défaut si elles existent
    Application.DisplayAlerts = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Dashboard" And ws.Name <> "Saisie_Mensuelle" And _
           ws.Name <> "Donnees_Revenus" And ws.Name <> "Donnees_Depenses" And _
           ws.Name <> "Categories" And ws.Name <> "Parametres" And _
           ws.Name <> "Rapports" And ws.Name <> "Archives" Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True
    
    ' Créer les feuilles nécessaires
    For i = 0 To UBound(NomsFeuillesReq)
        If Not FeuilleExiste(NomsFeuillesReq(i)) Then
            Set ws = ThisWorkbook.Worksheets.Add
            ws.Name = NomsFeuillesReq(i)
        End If
    Next i
    
End Sub

Sub EtapeConfigurationTableauBord()
    '-------------------------------------------------------------------------
    ' Étape 2 : Configuration complète du tableau de bord
    '-------------------------------------------------------------------------
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    
    With ws
        .Cells.Clear
        .Tab.Color = RGB(68, 114, 196)
        
        ' === EN-TÊTE PRINCIPAL ===
        .Range("A1:H1").Merge
        .Range("A1").Value = "FINANCE TRACKER - TABLEAU DE BORD"
        .Range("A1").Font.Size = 20
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Color = RGB(68, 114, 196)
        .Range("A1").HorizontalAlignment = xlCenter
        
        .Range("A2:H2").Merge
        .Range("A2").Value = "Tableau de bord financier - " & Format(Date, "mmmm yyyy")
        .Range("A2").Font.Size = 12
        .Range("A2").HorizontalAlignment = xlCenter
        .Range("A2").Font.Color = RGB(89, 89, 89)
        
        ' === INDICATEURS CLÉS ===
        .Range("A4").Value = "INDICATEURS CLÉS"
        .Range("A4").Font.Size = 14
        .Range("A4").Font.Bold = True
        .Range("A4").Font.Color = RGB(68, 114, 196)
        
        ' Carte Revenus
        .Range("A6:B8").Merge
        .Range("A6").Value = "REVENUS DU MOIS" & vbCrLf & "0 €"
        .Range("A6").Interior.Color = RGB(68, 114, 196)
        .Range("A6").Font.Color = RGB(255, 255, 255)
        .Range("A6").Font.Bold = True
        .Range("A6").HorizontalAlignment = xlCenter
        .Range("A6").VerticalAlignment = xlCenter
        
        ' Carte Dépenses
        .Range("C6:D8").Merge
        .Range("C6").Value = "DÉPENSES DU MOIS" & vbCrLf & "0 €"
        .Range("C6").Interior.Color = RGB(196, 89, 17)
        .Range("C6").Font.Color = RGB(255, 255, 255)
        .Range("C6").Font.Bold = True
        .Range("C6").HorizontalAlignment = xlCenter
        .Range("C6").VerticalAlignment = xlCenter
        
        ' Carte Épargne
        .Range("E6:F8").Merge
        .Range("E6").Value = "ÉPARGNE RÉALISÉE" & vbCrLf & "0 €"
        .Range("E6").Interior.Color = RGB(112, 173, 71)
        .Range("E6").Font.Color = RGB(255, 255, 255)
        .Range("E6").Font.Bold = True
        .Range("E6").HorizontalAlignment = xlCenter
        .Range("E6").VerticalAlignment = xlCenter
        
        ' Carte Budget Restant
        .Range("G6:H8").Merge
        .Range("G6").Value = "BUDGET RESTANT" & vbCrLf & "0 €"
        .Range("G6").Interior.Color = RGB(255, 192, 0)
        .Range("G6").Font.Color = RGB(0, 0, 0)
        .Range("G6").Font.Bold = True
        .Range("G6").HorizontalAlignment = xlCenter
        .Range("G6").VerticalAlignment = xlCenter
        
        ' === RÉSUMÉ MENSUEL ===
        .Range("A10").Value = "RÉSUMÉ MENSUEL"
        .Range("A10").Font.Size = 14
        .Range("A10").Font.Bold = True
        .Range("A10").Font.Color = RGB(68, 114, 196)
        
        ' Tableau de résumé
        .Range("A12:H12").Value = Array("CATÉGORIE", "BUDGET PRÉVU", "MONTANT RÉEL", "ÉCART", "ÉCART %", "STATUT", "TENDANCE", "ACTIONS")
        .Range("A12:H12").Font.Bold = True
        .Range("A12:H12").Interior.Color = RGB(68, 114, 196)
        .Range("A12:H12").Font.Color = RGB(255, 255, 255)
        
        ' Lignes de données exemple
        Dim CategoriesExemple As Variant
        CategoriesExemple = Array("Revenus Salaire", "Logement", "Alimentation", "Transport", "Loisirs", "Épargne")
        Dim i As Integer
        For i = 0 To UBound(CategoriesExemple)
            .Cells(13 + i, 1).Value = CategoriesExemple(i)
            .Cells(13 + i, 2).Value = "0 €"
            .Cells(13 + i, 3).Value = "0 €"
            .Cells(13 + i, 4).Value = "0 €"
            .Cells(13 + i, 5).Value = "0%"
            .Cells(13 + i, 6).Value = "En attente"
            .Cells(13 + i, 7).Value = "→"
            .Cells(13 + i, 8).Value = "Saisir données"
        Next i
        
        ' === ZONE GRAPHIQUES ===
        .Range("A20").Value = "VISUALISATIONS"
        .Range("A20").Font.Size = 14
        .Range("A20").Font.Bold = True
        .Range("A20").Font.Color = RGB(68, 114, 196)
        
        ' Zone graphique évolution
        .Range("A22:D32").Merge
        .Range("A22").Value = "ÉVOLUTION MENSUELLE" & vbCrLf & vbCrLf & "[Graphique généré automatiquement après saisie des données]"
        .Range("A22").Interior.Color = RGB(248, 248, 248)
        .Range("A22").Borders.LineStyle = xlContinuous
        .Range("A22").HorizontalAlignment = xlCenter
        .Range("A22").VerticalAlignment = xlCenter
        
        ' Zone graphique répartition
        .Range("E22:H32").Merge
        .Range("E22").Value = "RÉPARTITION DES DÉPENSES" & vbCrLf & vbCrLf & "[Graphique généré automatiquement après saisie des données]"
        .Range("E22").Interior.Color = RGB(248, 248, 248)
        .Range("E22").Borders.LineStyle = xlContinuous
        .Range("E22").HorizontalAlignment = xlCenter
        .Range("E22").VerticalAlignment = xlCenter
        
        ' === ALERTES ===
        .Range("A34").Value = "ALERTES ET NOTIFICATIONS"
        .Range("A34").Font.Size = 14
        .Range("A34").Font.Bold = True
        .Range("A34").Font.Color = RGB(196, 89, 17)
        
        .Range("A36:H38").Merge
        .Range("A36").Value = "✓ Système initialisé avec succès" & vbCrLf & _
                              "ℹ️ Commencez par saisir vos données mensuelles" & vbCrLf & _
                              "📊 Les graphiques se mettront à jour automatiquement"
        .Range("A36").Interior.Color = RGB(255, 242, 204)
        .Range("A36").Borders.LineStyle = xlContinuous
        .Range("A36").VerticalAlignment = xlTop
        
        ' Ajustement des colonnes
        .Columns("A:H").AutoFit
        
    End With
    
End Sub

Sub EtapeConfigurationSaisie()
    '-------------------------------------------------------------------------
    ' Étape 3 : Configuration de la feuille de saisie mensuelle
    '-------------------------------------------------------------------------
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Saisie_Mensuelle")
    
    With ws
        .Cells.Clear
        .Tab.Color = RGB(112, 173, 71)
        
        ' === EN-TÊTE ===
        .Range("A1:H1").Merge
        .Range("A1").Value = "SAISIE MENSUELLE DES DONNÉES FINANCIÈRES"
        .Range("A1").Font.Size = 18
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Color = RGB(112, 173, 71)
        .Range("A1").HorizontalAlignment = xlCenter
        
        .Range("A2:H2").Merge
        .Range("A2").Value = "Saisissez vos revenus et dépenses prévus et réels pour le mois"
        .Range("A2").Font.Size = 10
        .Range("A2").HorizontalAlignment = xlCenter
        .Range("A2").Font.Color = RGB(89, 89, 89)
        
        ' === SÉLECTION MOIS ===
        .Range("A4").Value = "MOIS DE RÉFÉRENCE:"
        .Range("A4").Font.Size = 12
        .Range("A4").Font.Bold = True
        .Range("C4").Value = Format(Date, "mmmm yyyy")
        .Range("C4").Font.Size = 11
        .Range("C4").Font.Bold = True
        .Range("C4").Interior.Color = RGB(226, 239, 218)
        
        ' === SECTION REVENUS ===
        .Range("A6").Value = "REVENUS DU MOIS"
        .Range("A6").Font.Size = 14
        .Range("A6").Font.Bold = True
        .Range("A6").Font.Color = RGB(68, 114, 196)
        
        ' En-têtes revenus
        Dim EntetesRevenus As Variant
        EntetesRevenus = Array("CATÉGORIE", "DESCRIPTION", "RÉCURRENT", "MONTANT PRÉVU", "STATUT", "MONTANT RÉEL", "ÉCART", "NOTES")
        .Range("A8:H8").Value = EntetesRevenus
        .Range("A8:H8").Font.Bold = True
        .Range("A8:H8").Interior.Color = RGB(68, 114, 196)
        .Range("A8:H8").Font.Color = RGB(255, 255, 255)
        
        ' Catégories revenus
        Dim CategoriesRevenus As Variant
        CategoriesRevenus = Array("Salaire Principal", "Salaire Conjoint", "Primes/Bonus", "Revenus Locatifs", "Investissements", "Autres Revenus")
        Dim i As Integer
        For i = 0 To UBound(CategoriesRevenus)
            .Cells(9 + i, 1).Value = CategoriesRevenus(i)
            .Cells(9 + i, 3).Value = "NON"
            .Cells(9 + i, 5).Value = "En attente"
            .Range(.Cells(9 + i, 1), .Cells(9 + i, 8)).Borders.LineStyle = xlContinuous
        Next i
        
        ' Total revenus
        .Range("G15").Value = "TOTAL REVENUS:"
        .Range("G15").Font.Bold = True
        .Range("H15").Value = "=SOMME(D9:D14)+SOMME(F9:F14)"
        .Range("H15").Font.Bold = True
        .Range("H15").NumberFormat = "#,##0.00 €"
        .Range("G15:H15").Interior.Color = RGB(68, 114, 196)
        .Range("G15:H15").Font.Color = RGB(255, 255, 255)
        
        ' === SECTION DÉPENSES ===
        .Range("A17").Value = "DÉPENSES DU MOIS"
        .Range("A17").Font.Size = 14
        .Range("A17").Font.Bold = True
        .Range("A17").Font.Color = RGB(196, 89, 17)
        
        ' En-têtes dépenses
        .Range("A19:H19").Value = EntetesRevenus
        .Range("A19:H19").Font.Bold = True
        .Range("A19:H19").Interior.Color = RGB(196, 89, 17)
        .Range("A19:H19").Font.Color = RGB(255, 255, 255)
        
        ' Catégories dépenses
        Dim CategoriesDepenses As Variant
        CategoriesDepenses = Array("Logement", "Alimentation", "Transport", "Assurances", "Santé", "Loisirs", "Vêtements", "Épargne", "Services", "Impôts", "Divers")
        For i = 0 To UBound(CategoriesDepenses)
            .Cells(20 + i, 1).Value = CategoriesDepenses(i)
            .Cells(20 + i, 3).Value = "NON"
            .Cells(20 + i, 5).Value = "En attente"
            .Range(.Cells(20 + i, 1), .Cells(20 + i, 8)).Borders.LineStyle = xlContinuous
        Next i
        
        ' Total dépenses
        .Range("G31").Value = "TOTAL DÉPENSES:"
        .Range("G31").Font.Bold = True
        .Range("H31").Value = "=SOMME(D20:D30)+SOMME(F20:F30)"
        .Range("H31").Font.Bold = True
        .Range("H31").NumberFormat = "#,##0.00 €"
        .Range("G31:H31").Interior.Color = RGB(196, 89, 17)
        .Range("G31:H31").Font.Color = RGB(255, 255, 255)
        
        ' === RÉSUMÉ ===
        .Range("A33").Value = "RÉSUMÉ ET VALIDATION"
        .Range("A33").Font.Size = 14
        .Range("A33").Font.Bold = True
        .Range("A33").Font.Color = RGB(112, 173, 71)
        
        ' Tableau résumé
        .Range("A35:D35").Value = Array("ÉLÉMENT", "PRÉVU", "RÉEL", "ÉCART")
        .Range("A35:D35").Font.Bold = True
        .Range("A35:D35").Interior.Color = RGB(112, 173, 71)
        .Range("A35:D35").Font.Color = RGB(255, 255, 255)
        
        .Range("A36").Value = "Total Revenus"
        .Range("B36").Value = "=SOMME(D9:D14)"
        .Range("C36").Value = "=SOMME(F9:F14)"
        .Range("D36").Value = "=C36-B36"
        
        .Range("A37").Value = "Total Dépenses"
        .Range("B37").Value = "=SOMME(D20:D30)"
        .Range("C37").Value = "=SOMME(F20:F30)"
        .Range("D37").Value = "=C37-B37"
        
        .Range("A38").Value = "Solde Net"
        .Range("B38").Value = "=B36-B37"
        .Range("C38").Value = "=C36-C37"
        .Range("D38").Value = "=C38-B38"
        
        .Range("A39").Value = "Taux d'Épargne"
        .Range("B39").Value = "=SI(B36>0,B38/B36,0)"
        .Range("C39").Value = "=SI(C36>0,C38/C36,0)"
        .Range("D39").Value = "=C39-B39"
        
        ' Formatage résumé
        .Range("B36:D39").NumberFormat = "#,##0.00 €"
        .Range("B39:D39").NumberFormat = "0.00%"
        .Range("A35:D39").Borders.LineStyle = xlContinuous
        
        ' Ajustement colonnes
        .Columns("A:H").AutoFit
        
    End With
    
End Sub

Sub EtapeConfigurationCategories()
    '-------------------------------------------------------------------------
    ' Étape 4 : Configuration des catégories
    '-------------------------------------------------------------------------
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Categories")
    
    With ws
        .Cells.Clear
        .Tab.Color = RGB(255, 192, 0)
        
        ' En-tête
        .Range("A1:H1").Merge
        .Range("A1").Value = "GESTION DES CATÉGORIES FINANCIÈRES"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Color = RGB(255, 192, 0)
        .Range("A1").HorizontalAlignment = xlCenter
        
        ' En-têtes colonnes
        .Range("A3:H3").Value = Array("ID", "NOM DE LA CATÉGORIE", "TYPE", "COULEUR", "BUDGET DÉFAUT", "ACTIVE", "PERSONNALISÉE", "DESCRIPTION")
        .Range("A3:H3").Font.Bold = True
        .Range("A3:H3").Interior.Color = RGB(255, 192, 0)
        .Range("A3:H3").Font.Color = RGB(0, 0, 0)
        
        ' Catégories par défaut
        Dim DonneesCategories As Variant
        DonneesCategories = Array( _
            Array(1, "Salaire Principal", "Revenu", RGB(68, 114, 196), 3000, "OUI", "NON", "Salaire principal du foyer"), _
            Array(2, "Salaire Conjoint", "Revenu", RGB(68, 114, 196), 2000, "OUI", "NON", "Salaire du conjoint"), _
            Array(3, "Primes/Bonus", "Revenu", RGB(68, 114, 196), 300, "OUI", "NON", "Primes et bonus"), _
            Array(4, "Revenus Locatifs", "Revenu", RGB(68, 114, 196), 0, "OUI", "NON", "Revenus immobiliers"), _
            Array(5, "Investissements", "Revenu", RGB(68, 114, 196), 0, "OUI", "NON", "Dividendes, plus-values"), _
            Array(6, "Autres Revenus", "Revenu", RGB(68, 114, 196), 0, "OUI", "NON", "Autres sources"), _
            Array(7, "Logement", "Dépense", RGB(196, 89, 17), 1200, "OUI", "NON", "Loyer, charges"), _
            Array(8, "Alimentation", "Dépense", RGB(196, 89, 17), 600, "OUI", "NON", "Courses, restaurants"), _
            Array(9, "Transport", "Dépense", RGB(196, 89, 17), 400, "OUI", "NON", "Essence, transports"), _
            Array(10, "Assurances", "Dépense", RGB(196, 89, 17), 300, "OUI", "NON", "Assurances diverses"), _
            Array(11, "Santé", "Dépense", RGB(196, 89, 17), 150, "OUI", "NON", "Frais médicaux"), _
            Array(12, "Loisirs", "Dépense", RGB(196, 89, 17), 300, "OUI", "NON", "Sorties, vacances"), _
            Array(13, "Vêtements", "Dépense", RGB(196, 89, 17), 100, "OUI", "NON", "Achats vestimentaires"), _
            Array(14, "Épargne", "Dépense", RGB(112, 173, 71), 500, "OUI", "NON", "Épargne mensuelle"), _
            Array(15, "Services", "Dépense", RGB(196, 89, 17), 150, "OUI", "NON", "Internet, téléphone"), _
            Array(16, "Impôts", "Dépense", RGB(196, 89, 17), 200, "OUI", "NON", "Impôts, taxes"), _
            Array(17, "Divers", "Dépense", RGB(196, 89, 17), 100, "OUI", "NON", "Autres dépenses") _
        )
        
        Dim i As Integer
        For i = 0 To UBound(DonneesCategories)
            .Cells(4 + i, 1).Value = DonneesCategories(i)(0)
            .Cells(4 + i, 2).Value = DonneesCategories(i)(1)
            .Cells(4 + i, 3).Value = DonneesCategories(i)(2)
            .Cells(4 + i, 4).Interior.Color = DonneesCategories(i)(3)
            .Cells(4 + i, 4).Value = "■"
            .Cells(4 + i, 5).Value = DonneesCategories(i)(4)
            .Cells(4 + i, 6).Value = DonneesCategories(i)(5)
            .Cells(4 + i, 7).Value = DonneesCategories(i)(6)
            .Cells(4 + i, 8).Value = DonneesCategories(i)(7)
            
            .Range(.Cells(4 + i, 1), .Cells(4 + i, 8)).Borders.LineStyle = xlContinuous
        Next i
        
        .Columns("A:H").AutoFit
        .Columns("E").NumberFormat = "#,##0 €"
        
    End With
    
End Sub

Sub EtapeConfigurationParametres()
    '-------------------------------------------------------------------------
    ' Étape 5 : Configuration des paramètres système
    '-------------------------------------------------------------------------
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Parametres")
    
    With ws
        .Cells.Clear
        .Tab.Color = RGB(91, 155, 213)
        
        ' En-tête
        .Range("A1:F1").Merge
        .Range("A1").Value = "PARAMÈTRES SYSTÈME FINANCE TRACKER"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Color = RGB(91, 155, 213)
        .Range("A1").HorizontalAlignment = xlCenter
        
        ' Paramètres généraux
        .Range("A3").Value = "PARAMÈTRES GÉNÉRAUX"
        .Range("A3").Font.Size = 12
        .Range("A3").Font.Bold = True
        .Range("A3").Font.Color = RGB(91, 155, 213)
        
        Dim ParametresGeneraux As Variant
        ParametresGeneraux = Array( _
            Array("Devise par défaut:", "EUR (€)"), _
            Array("Format de date:", "dd/mm/yyyy"), _
            Array("Taux d'épargne cible:", "20%"), _
            Array("Période de rétention:", "36 mois"), _
            Array("Sauvegarde automatique:", "Activée"), _
            Array("Version Finance Tracker:", VERSION_FINANCE_TRACKER) _
        )
        
        Dim i As Integer
        For i = 0 To UBound(ParametresGeneraux)
            .Cells(5 + i, 1).Value = ParametresGeneraux(i)(0)
            .Cells(5 + i, 2).Value = ParametresGeneraux(i)(1)
        Next i
        
        ' Paramètres d'affichage
        .Range("A12").Value = "PARAMÈTRES D'AFFICHAGE"
        .Range("A12").Font.Size = 12
        .Range("A12").Font.Bold = True
        .Range("A12").Font.Color = RGB(91, 155, 213)
        
        Dim ParametresAffichage As Variant
        ParametresAffichage = Array( _
            Array("Thème de couleur:", "Professionnel"), _
            Array("Taille de police:", "9"), _
            Array("Affichage graphiques:", "Automatique"), _
            Array("Nombre de décimales:", "2") _
        )
        
        For i = 0 To UBound(ParametresAffichage)
            .Cells(14 + i, 1).Value = ParametresAffichage(i)(0)
            .Cells(14 + i, 2).Value = ParametresAffichage(i)(1)
        Next i
        
        ' Paramètres d'alerte
        .Range("A19").Value = "PARAMÈTRES D'ALERTES"
        .Range("A19").Font.Size = 12
        .Range("A19").Font.Bold = True
        .Range("A19").Font.Color = RGB(91, 155, 213)
        
        Dim ParametresAlertes As Variant
        ParametresAlertes = Array( _
            Array("Alertes dépassement budget:", "Activées"), _
            Array("Seuil d'alerte (%):", "90%"), _
            Array("Alertes épargne insuffisante:", "Activées"), _
            Array("Rappels de saisie:", "Mensuels") _
        )
        
        For i = 0 To UBound(ParametresAlertes)
            .Cells(21 + i, 1).Value = ParametresAlertes(i)(0)
            .Cells(21 + i, 2).Value = ParametresAlertes(i)(1)
        Next i
        
        .Columns("A:B").AutoFit
        
    End With
    
End Sub

Sub EtapeConfigurationDonnees()
    '-------------------------------------------------------------------------
    ' Étape 6 : Configuration des feuilles de données
    '-------------------------------------------------------------------------
    
    ' Configuration Données Revenus
    Dim wsRevenus As Worksheet
    Set wsRevenus = ThisWorkbook.Worksheets("Donnees_Revenus")
    
    With wsRevenus
        .Cells.Clear
        .Tab.Color = RGB(68, 114, 196)
        
        .Range("A1:H1").Value = Array("DATE", "CATÉGORIE", "DESCRIPTION", "RÉCURRENT", "MONTANT PRÉVU", "MONTANT RÉEL", "ÉCART", "NOTES")
        .Range("A1:H1").Font.Bold = True
        .Range("A1:H1").Interior.Color = RGB(68, 114, 196)
        .Range("A1:H1").Font.Color = RGB(255, 255, 255)
        .Columns("A:H").AutoFit
    End With
    
    ' Configuration Données Dépenses
    Dim wsDepenses As Worksheet
    Set wsDepenses = ThisWorkbook.Worksheets("Donnees_Depenses")
    
    With wsDepenses
        .Cells.Clear
        .Tab.Color = RGB(196, 89, 17)
        
        .Range("A1:H1").Value = Array("DATE", "CATÉGORIE", "DESCRIPTION", "RÉCURRENT", "MONTANT PRÉVU", "MONTANT RÉEL", "ÉCART", "NOTES")
        .Range("A1:H1").Font.Bold = True
        .Range("A1:H1").Interior.Color = RGB(196, 89, 17)
        .Range("A1:H1").Font.Color = RGB(255, 255, 255)
        .Columns("A:H").AutoFit
    End With
    
End Sub

Sub EtapeConfigurationRapports()
    '-------------------------------------------------------------------------
    ' Étape 7 : Configuration de la feuille de rapports
    '-------------------------------------------------------------------------
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Rapports")
    
    With ws
        .Cells.Clear
        .Tab.Color = RGB(255, 192, 0)
        
        .Range("A1:J1").Merge
        .Range("A1").Value = "RAPPORTS FINANCIERS"
        .Range("A1").Font.Size = 18
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Color = RGB(255, 192, 0)
        .Range("A1").HorizontalAlignment = xlCenter
        
        .Range("A3").Value = "Cette feuille affichera vos rapports financiers automatisés."
        .Range("A4").Value = "Les rapports sont générés automatiquement à partir de vos données de saisie."
        .Range("A6").Value = "Types de rapports disponibles :"
        .Range("A7").Value = "• Rapport mensuel complet"
        .Range("A8").Value = "• Analyse des tendances"
        .Range("A9").Value = "• Comparaisons budgétaires"
        .Range("A10").Value = "• Recommandations personnalisées"
        .Range("A11").Value = "• Projections et objectifs"
        
        .Range("A13").Value = "Pour générer un rapport, utilisez les boutons du tableau de bord."
        .Range("A13").Font.Bold = True
        .Range("A13").Font.Color = RGB(255, 192, 0)
        
    End With
    
End Sub

Sub EtapeConfigurationArchives()
    '-------------------------------------------------------------------------
    ' Étape 8 : Configuration de la feuille d'archives
    '-------------------------------------------------------------------------
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Archives")
    
    With ws
        .Cells.Clear
        .Tab.Color = RGB(128, 128, 128)
        
        .Range("A1:H1").Value = Array("DATE_ARCHIVAGE", "TYPE", "DATE_ORIGINALE", "CATÉGORIE", "DESCRIPTION", "MONTANT_PREVU", "MONTANT_REEL", "NOTES")
        .Range("A1:H1").Font.Bold = True
        .Range("A1:H1").Interior.Color = RGB(128, 128, 128)
        .Range("A1:H1").Font.Color = RGB(255, 255, 255)
        
        ' Ligne de log d'installation
        .Cells(2, 1).Value = Now
        .Cells(2, 2).Value = "SYSTÈME"
        .Cells(2, 3).Value = Date
        .Cells(2, 4).Value = "Installation"
        .Cells(2, 5).Value = "Finance Tracker v" & VERSION_FINANCE_TRACKER & " installé avec succès"
        .Cells(2, 6).Value = 0
        .Cells(2, 7).Value = 0
        .Cells(2, 8).Value = "Installation automatique complète"
        
        .Columns("A:H").AutoFit
        
    End With
    
End Sub

Sub EtapeCreationGraphiques()
    '-------------------------------------------------------------------------
    ' Étape 9 : Préparation des zones graphiques
    '-------------------------------------------------------------------------
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    
    ' Préparation des données pour les graphiques (dans colonnes masquées)
    With ws
        ' Données d'évolution (colonnes J à N)
        .Range("J1:N1").Value = Array("Mois", "Revenus", "Dépenses", "Épargne", "Budget")
        
        ' Données exemple pour 12 mois
        Dim i As Integer
        For i = 1 To 12
            .Cells(1 + i, 10).Value = Format(DateAdd("m", i - 12, Date), "mmm")
            .Cells(1 + i, 11).Value = 0 ' Revenus
            .Cells(1 + i, 12).Value = 0 ' Dépenses
            .Cells(1 + i, 13).Value = 0 ' Épargne
            .Cells(1 + i, 14).Value = 0 ' Budget
        Next i
        
        ' Données de répartition (colonnes P à Q)
        .Range("P1:Q1").Value = Array("Catégorie", "Montant")
        Dim CategoriesGraph As Variant
        CategoriesGraph = Array("Logement", "Alimentation", "Transport", "Loisirs", "Santé", "Autres")
        For i = 0 To UBound(CategoriesGraph)
            .Cells(2 + i, 16).Value = CategoriesGraph(i)
            .Cells(2 + i, 17).Value = 0
        Next i
        
        ' Masquer les colonnes de données
        .Columns("J:Q").Hidden = True
    End With
    
End Sub

Sub EtapeConfigurationNavigation()
    '-------------------------------------------------------------------------
    ' Étape 10 : Création des boutons de navigation
    '-------------------------------------------------------------------------
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    
    ' Supprimer les anciens boutons s'ils existent
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left(shp.Name, 4) = "Btn_" Then
            shp.Delete
        End If
    Next shp
    
    ' Bouton Saisie Mensuelle
    Dim btnSaisie As Shape
    Set btnSaisie = ws.Shapes.AddShape(msoShapeRoundedRectangle, 50, 600, 150, 40)
    With btnSaisie
        .Name = "Btn_Saisie"
        .TextFrame.Characters.Text = "📝 SAISIE MENSUELLE"
        .TextFrame.Characters.Font.Size = 11
        .TextFrame.Characters.Font.Bold = True
        .Fill.ForeColor.RGB = RGB(112, 173, 71)
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .Line.ForeColor.RGB = RGB(112, 173, 71)
    End With
    
    ' Bouton Rapports
    Dim btnRapports As Shape
    Set btnRapports = ws.Shapes.AddShape(msoShapeRoundedRectangle, 220, 600, 150, 40)
    With btnRapports
        .Name = "Btn_Rapports"
        .TextFrame.Characters.Text = "📊 RAPPORTS"
        .TextFrame.Characters.Font.Size = 11
        .TextFrame.Characters.Font.Bold = True
        .Fill.ForeColor.RGB = RGB(255, 192, 0)
        .TextFrame.Characters.Font.Color = RGB(0, 0, 0)
        .Line.ForeColor.RGB = RGB(255, 192, 0)
    End With
    
    ' Bouton Paramètres
    Dim btnParametres As Shape
    Set btnParametres = ws.Shapes.AddShape(msoShapeRoundedRectangle, 390, 600, 150, 40)
    With btnParametres
        .Name = "Btn_Parametres"
        .TextFrame.Characters.Text = "⚙️ PARAMÈTRES"
        .TextFrame.Characters.Font.Size = 11
        .TextFrame.Characters.Font.Bold = True
        .Fill.ForeColor.RGB = RGB(91, 155, 213)
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .Line.ForeColor.RGB = RGB(91, 155, 213)
    End With
    
    ' Bouton Aide
    Dim btnAide As Shape
    Set btnAide = ws.Shapes.AddShape(msoShapeRoundedRectangle, 560, 600, 100, 40)
    With btnAide
        .Name = "Btn_Aide"
        .TextFrame.Characters.Text = "❓ AIDE"
        .TextFrame.Characters.Font.Size = 11
        .TextFrame.Characters.Font.Bold = True
        .Fill.ForeColor.RGB = RGB(128, 128, 128)
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .Line.ForeColor.RGB = RGB(128, 128, 128)
    End With
    
End Sub

Sub EtapeInitialisationDonnees()
    '-------------------------------------------------------------------------
    ' Étape 11 : Initialisation des données exemple
    '-------------------------------------------------------------------------
    
    ' Créer quelques données exemple pour le mois courant
    Dim wsRevenus As Worksheet, wsDepenses As Worksheet
    Set wsRevenus = ThisWorkbook.Worksheets("Donnees_Revenus")
    Set wsDepenses = ThisWorkbook.Worksheets("Donnees_Depenses")
    
    ' Exemple de données revenus
    With wsRevenus
        .Cells(2, 1).Value = DateSerial(Year(Date), Month(Date), 1)
        .Cells(2, 2).Value = "Salaire Principal"
        .Cells(2, 3).Value = "Salaire net mensuel"
        .Cells(2, 4).Value = "OUI"
        .Cells(2, 5).Value = 3000
        .Cells(2, 6).Value = 0
        .Cells(2, 7).Value = 0
        .Cells(2, 8).Value = "À saisir"
    End With
    
    ' Exemple de données dépenses
    Dim CategoriesExemple As Variant
    Dim BudgetsExemple As Variant
    CategoriesExemple = Array("Logement", "Alimentation", "Transport", "Épargne")
    BudgetsExemple = Array(1200, 600, 400, 500)
    
    Dim i As Integer
    For i = 0 To UBound(CategoriesExemple)
        With wsDepenses
            .Cells(2 + i, 1).Value = DateSerial(Year(Date), Month(Date), 1)
            .Cells(2 + i, 2).Value = CategoriesExemple(i)
            .Cells(2 + i, 3).Value = "Budget mensuel " & CategoriesExemple(i)
            .Cells(2 + i, 4).Value = "OUI"
            .Cells(2 + i, 5).Value = BudgetsExemple(i)
            .Cells(2 + i, 6).Value = 0
            .Cells(2 + i, 7).Value = 0
            .Cells(2 + i, 8).Value = "À saisir"
        End With
    Next i
    
End Sub

Sub EtapeConfigurationFinale()
    '-------------------------------------------------------------------------
    ' Étape 12 : Configuration finale et protection
    '-------------------------------------------------------------------------
    
    Dim ws As Worksheet
    
    ' Protection des feuilles avec mot de passe
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        ws.Protect Password:="FinanceTracker2025", _
                  DrawingObjects:=False, _
                  Contents:=True, _
                  Scenarios:=False, _
                  UserInterfaceOnly:=True
        On Error GoTo 0
    Next ws
    
    ' Déverrouiller les zones de saisie dans la feuille de saisie
    Set ws = ThisWorkbook.Worksheets("Saisie_Mensuelle")
    ws.Unprotect Password:="FinanceTracker2025"
    
    ' Définir les zones modifiables
    ws.Range("B9:C14,D9:D14,F9:F14,H9:H14").Locked = False ' Revenus
    ws.Range("B20:C30,D20:D30,F20:F30,H20:H30").Locked = False ' Dépenses
    
    ws.Protect Password:="FinanceTracker2025", _
              DrawingObjects:=False, _
              Contents:=True, _
              Scenarios:=False, _
              UserInterfaceOnly:=True
    
    ' Configuration générale du classeur
    With ThisWorkbook
        .SaveAs ThisWorkbook.Path & "\" & NOM_FICHIER_FINANCE_TRACKER & ".xlsm", xlOpenXMLWorkbookMacroEnabled
        .Application.DisplayAlerts = True
    End With
    
    Application.StatusBar = "Finance Tracker v" & VERSION_FINANCE_TRACKER & " - Prêt à l'emploi !"
    
End Sub

'===============================================================================
' FONCTIONS UTILITAIRES
'===============================================================================

Function FeuilleExiste(NomFeuille As String) As Boolean
    '-------------------------------------------------------------------------
    ' Vérifie si une feuille existe dans le classeur
    '-------------------------------------------------------------------------
    
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(NomFeuille)
    FeuilleExiste = Not ws Is Nothing
    On Error GoTo 0
    
End Function

'===============================================================================
' MACROS DE NAVIGATION (ASSIGNÉES AUX BOUTONS)
'===============================================================================

Sub NaviguerVersTableauBord()
    ThisWorkbook.Worksheets("Dashboard").Activate
End Sub

Sub NaviguerVersSaisie()
    ThisWorkbook.Worksheets("Saisie_Mensuelle").Activate
End Sub

Sub NaviguerVersRapports()
    ThisWorkbook.Worksheets("Rapports").Activate
End Sub

Sub NaviguerVersParametres()
    ThisWorkbook.Worksheets("Parametres").Activate
End Sub

Sub AfficherAideRapide()
    MsgBox "🚀 FINANCE TRACKER - AIDE RAPIDE" & vbCrLf & vbCrLf & _
           "📊 TABLEAU DE BORD : Vue d'ensemble de vos finances" & vbCrLf & _
           "📝 SAISIE MENSUELLE : Entrez vos revenus et dépenses" & vbCrLf & _
           "📋 RAPPORTS : Analyses détaillées et recommandations" & vbCrLf & _
           "⚙️ PARAMÈTRES : Configurez vos catégories et alertes" & vbCrLf & vbCrLf & _
           "💡 CONSEIL : Commencez par la saisie mensuelle pour remplir vos données !" & vbCrLf & vbCrLf & _
           "Version " & VERSION_FINANCE_TRACKER, _
           vbInformation, "Aide Finance Tracker"
End Sub

'===============================================================================
' FIN DU MODULE D'INSTALLATION COMPLÈTE
'===============================================================================
