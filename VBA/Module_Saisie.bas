Attribute VB_Name = "Module_Saisie"
'===============================================================================
' FINANCE TRACKER VBA - MODULE SAISIE MENSUELLE
' Version: 1.0
' Description: Interface de saisie pour les données financières mensuelles
' Fonction: Gestion complète des revenus et dépenses avec validation
'===============================================================================

Option Explicit

' Constantes pour la validation des données
Private Const MONTANT_MINIMUM As Currency = 0
Private Const MONTANT_MAXIMUM As Currency = 999999.99
Private Const LONGUEUR_MAX_DESCRIPTION As Integer = 100

'===============================================================================
' PROCEDURES DE CREATION DE L'INTERFACE DE SAISIE
'===============================================================================

Sub CreerFeuilleSaisie(ws As Worksheet)
    '-------------------------------------------------------------------------
    ' Crée l'interface complète de saisie mensuelle
    '-------------------------------------------------------------------------
    
    On Error GoTo GestionErreur
    
    Application.ScreenUpdating = False
    
    With ws
        .Cells.Clear
        .Tab.Color = RGB(112, 173, 71)
        
        ' En-tête de la feuille de saisie
        Call CreerEnteteSaisie(ws)
        
        ' Section sélection du mois
        Call CreerSelectionMois(ws)
        
        ' Section saisie des revenus
        Call CreerSectionRevenus(ws)
        
        ' Section saisie des dépenses
        Call CreerSectionDepenses(ws)
        
        ' Section résumé et validation
        Call CreerSectionValidation(ws)
        
        ' Boutons d'action
        Call CreerBoutonsAction(ws)
        
        ' Formatage et protection
        Call AppliquerFormatageSaisie(ws)
    End With
    
    Application.ScreenUpdating = True
    Exit Sub
    
GestionErreur:
    Application.ScreenUpdating = True
    Call EnregistrerJournal("Erreur création feuille saisie: " & Err.Description, "ERREUR")
End Sub

Sub CreerEnteteSaisie(ws As Worksheet)
    '-------------------------------------------------------------------------
    ' Crée l'en-tête de la feuille de saisie mensuelle
    '-------------------------------------------------------------------------
    
    With ws
        ' Titre principal
        .Range("A1:H1").Merge
        .Range("A1").Value = "SAISIE MENSUELLE DES DONNÉES FINANCIÈRES"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Color = RGB(112, 173, 71)
        .Range("A1").HorizontalAlignment = xlCenter
        
        ' Instructions
        .Range("A2:H2").Merge
        .Range("A2").Value = "Saisissez vos revenus et dépenses prévus et réels pour le mois sélectionné"
        .Range("A2").Font.Size = 10
        .Range("A2").HorizontalAlignment = xlCenter
        .Range("A2").Font.Color = RGB(89, 89, 89)
        
        ' Ligne de séparation
        .Range("A3:H3").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("A3:H3").Borders(xlEdgeBottom).Color = RGB(112, 173, 71)
        .Range("A3:H3").Borders(xlEdgeBottom).Weight = xlMedium
    End With
    
End Sub

Sub CreerSelectionMois(ws As Worksheet)
    '-------------------------------------------------------------------------
    ' Crée la section de sélection du mois
    '-------------------------------------------------------------------------
    
    With ws
        ' Label de sélection
        .Range("A5").Value = "MOIS DE RÉFÉRENCE:"
        .Range("A5").Font.Size = 12
        .Range("A5").Font.Bold = True
        .Range("A5").Font.Color = RGB(112, 173, 71)
        
        ' Liste déroulante des mois
        .Range("C5").Value = Format(ObtenirMoisCourant, "mmmm yyyy")
        .Range("C5").Font.Size = 11
        .Range("C5").Font.Bold = True
        .Range("C5").Interior.Color = RGB(226, 239, 218)
        .Range("C5").Borders.LineStyle = xlContinuous
        
        ' Bouton de changement de mois
        Call CreerBoutonChangementMois(ws, "E5")
        
        ' Date de dernière modification
        .Range("G5").Value = "Dernière modification: " & Format(Now, "dd/mm/yyyy hh:mm")
        .Range("G5").Font.Size = 8
        .Range("G5").Font.Color = RGB(128, 128, 128)
    End With
    
End Sub

Sub CreerBoutonChangementMois(ws As Worksheet, Position As String)
    '-------------------------------------------------------------------------
    ' Crée le bouton pour changer le mois de référence
    '-------------------------------------------------------------------------
    
    Dim btnChangerMois As Shape
    
    Set btnChangerMois = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range(Position).Left, ws.Range(Position).Top, 80, 20)
    
    With btnChangerMois
        .Name = "Btn_ChangerMois"
        .TextFrame.Characters.Text = "Changer"
        .TextFrame.Characters.Font.Size = 8
        .TextFrame.Characters.Font.Bold = True
        .Fill.ForeColor.RGB = RGB(112, 173, 71)
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .OnAction = "OuvrirSelectionMois"
    End With
    
End Sub

Sub CreerSectionRevenus(ws As Worksheet)
    '-------------------------------------------------------------------------
    ' Crée la section de saisie des revenus
    '-------------------------------------------------------------------------
    
    With ws
        ' En-tête de la section revenus
        .Range("A7").Value = "REVENUS DU MOIS"
        .Range("A7").Font.Size = 14
        .Range("A7").Font.Bold = True
        .Range("A7").Font.Color = RGB(68, 114, 196)
        
        ' Création du tableau de saisie des revenus
        Call CreerTableauSaisie(ws, "A9:H16", "REVENUS")
        
        ' Total des revenus
        .Range("G17").Value = "TOTAL REVENUS:"
        .Range("G17").Font.Bold = True
        .Range("H17").Value = "=SOMME(D10:D16)+SOMME(F10:F16)"
        .Range("H17").Font.Bold = True
        .Range("H17").NumberFormat = "#,##0.00 €"
        .Range("G17:H17").Interior.Color = RGB(68, 114, 196)
        .Range("G17:H17").Font.Color = RGB(255, 255, 255)
    End With
    
End Sub

Sub CreerSectionDepenses(ws As Worksheet)
    '-------------------------------------------------------------------------
    ' Crée la section de saisie des dépenses
    '-------------------------------------------------------------------------
    
    With ws
        ' En-tête de la section dépenses
        .Range("A19").Value = "DÉPENSES DU MOIS"
        .Range("A19").Font.Size = 14
        .Range("A19").Font.Bold = True
        .Range("A19").Font.Color = RGB(196, 89, 17)
        
        ' Création du tableau de saisie des dépenses
        Call CreerTableauSaisie(ws, "A21:H35", "DEPENSES")
        
        ' Total des dépenses
        .Range("G36").Value = "TOTAL DÉPENSES:"
        .Range("G36").Font.Bold = True
        .Range("H36").Value = "=SOMME(D22:D35)+SOMME(F22:F35)"
        .Range("H36").Font.Bold = True
        .Range("H36").NumberFormat = "#,##0.00 €"
        .Range("G36:H36").Interior.Color = RGB(196, 89, 17)
        .Range("G36:H36").Font.Color = RGB(255, 255, 255)
    End With
    
End Sub

Sub CreerTableauSaisie(ws As Worksheet, PlageTableau As String, TypeSection As String)
    '-------------------------------------------------------------------------
    ' Crée un tableau de saisie standardisé pour revenus ou dépenses
    '-------------------------------------------------------------------------
    
    Dim PlageTab As Range
    Dim EntetesColonnes As Variant
    Dim CategoriesDefaut As Variant
    Dim i As Integer
    
    Set PlageTab = ws.Range(PlageTableau)
    EntetesColonnes = Array("CATÉGORIE", "DESCRIPTION", "RÉCURRENT", "MONTANT PRÉVU", "STATUT", "MONTANT RÉEL", "ÉCART", "NOTES")
    
    ' Définition des catégories selon le type
    If TypeSection = "REVENUS" Then
        CategoriesDefaut = Array("Salaire principal", "Salaire conjoint", "Primes/Bonus", "Revenus locatifs", "Investissements", "Autres revenus", "", "")
    Else
        CategoriesDefaut = Array("Logement", "Alimentation", "Transport", "Assurances", "Santé", "Loisirs", "Vêtements", "Épargne", "Impôts", "Divers", "", "", "", "", "")
    End If
    
    With PlageTab
        ' En-têtes de colonnes
        For i = 0 To UBound(EntetesColonnes)
            .Cells(1, i + 1).Value = EntetesColonnes(i)
            .Cells(1, i + 1).Font.Bold = True
            .Cells(1, i + 1).Font.Color = RGB(255, 255, 255)
            .Cells(1, i + 1).HorizontalAlignment = xlCenter
            
            If TypeSection = "REVENUS" Then
                .Cells(1, i + 1).Interior.Color = RGB(68, 114, 196)
            Else
                .Cells(1, i + 1).Interior.Color = RGB(196, 89, 17)
            End If
        Next i
        
        ' Lignes de données avec catégories prédéfinies
        For i = 0 To UBound(CategoriesDefaut)
            If i + 2 <= .Rows.Count Then
                .Cells(i + 2, 1).Value = CategoriesDefaut(i)
                
                ' Configuration des cellules de saisie
                .Cells(i + 2, 2).Locked = False ' Description
                .Cells(i + 2, 3).Locked = False ' Récurrent (Liste déroulante OUI/NON)
                .Cells(i + 2, 4).Locked = False ' Montant prévu
                .Cells(i + 2, 6).Locked = False ' Montant réel
                .Cells(i + 2, 8).Locked = False ' Notes
                
                ' Formules calculées
                .Cells(i + 2, 5).Value = "=SI(F" & (i + 2) & ">0,""Saisi"",""En attente"")" ' Statut
                .Cells(i + 2, 7).Value = "=F" & (i + 2) & "-D" & (i + 2) ' Écart
                
                ' Formatage conditionnel pour les montants
                .Cells(i + 2, 4).NumberFormat = "#,##0.00 €"
                .Cells(i + 2, 6).NumberFormat = "#,##0.00 €"
                .Cells(i + 2, 7).NumberFormat = "#,##0.00 €"
            End If
        Next i
        
        ' Bordures du tableau
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(128, 128, 128)
        .Font.Size = 9
    End With
    
End Sub

Sub CreerSectionValidation(ws As Worksheet)
    '-------------------------------------------------------------------------
    ' Crée la section de résumé et validation des données
    '-------------------------------------------------------------------------
    
    With ws
        ' En-tête de la section validation
        .Range("A38").Value = "RÉSUMÉ ET VALIDATION"
        .Range("A38").Font.Size = 14
        .Range("A38").Font.Bold = True
        .Range("A38").Font.Color = RGB(112, 173, 71)
        
        ' Tableau de résumé
        Call CreerTableauResumeMensuel(ws, "A40:D44")
        
        ' Indicateurs de validation
        Call CreerIndicateursValidation(ws, "F40:H44")
    End With
    
End Sub

Sub CreerTableauResumeMensuel(ws As Worksheet, PlageResume As String)
    '-------------------------------------------------------------------------
    ' Crée le tableau de résumé mensuel pour validation
    '-------------------------------------------------------------------------
    
    Dim PlageTab As Range
    Set PlageTab = ws.Range(PlageResume)
    
    With PlageTab
        ' En-têtes
        .Cells(1, 1).Value = "ÉLÉMENT"
        .Cells(1, 2).Value = "PRÉVU"
        .Cells(1, 3).Value = "RÉEL"
        .Cells(1, 4).Value = "ÉCART"
        
        ' Données
        .Cells(2, 1).Value = "Total Revenus"
        .Cells(2, 2).Value = "=SOMME(D10:D16)"
        .Cells(2, 3).Value = "=SOMME(F10:F16)"
        .Cells(2, 4).Value = "=C42-B42"
        
        .Cells(3, 1).Value = "Total Dépenses"
        .Cells(3, 2).Value = "=SOMME(D22:D35)"
        .Cells(3, 3).Value = "=SOMME(F22:F35)"
        .Cells(3, 4).Value = "=C43-B43"
        
        .Cells(4, 1).Value = "Solde Net"
        .Cells(4, 2).Value = "=B42-B43"
        .Cells(4, 3).Value = "=C42-C43"
        .Cells(4, 4).Value = "=C44-B44"
        
        .Cells(5, 1).Value = "Taux d'Épargne"
        .Cells(5, 2).Value = "=SI(B42>0,B44/B42,0)"
        .Cells(5, 3).Value = "=SI(C42>0,C44/C42,0)"
        .Cells(5, 4).Value = "=C45-B45"
        
        ' Formatage
        .Cells(1, 1).Resize(1, 4).Font.Bold = True
        .Cells(1, 1).Resize(1, 4).Interior.Color = RGB(112, 173, 71)
        .Cells(1, 1).Resize(1, 4).Font.Color = RGB(255, 255, 255)
        
        .Columns(2).Resize(, 3).NumberFormat = "#,##0.00 €"
        .Cells(5, 2).Resize(1, 3).NumberFormat = "0.00%"
        
        .Borders.LineStyle = xlContinuous
        .Font.Size = 9
    End With
    
End Sub

Sub CreerIndicateursValidation(ws As Worksheet, PlageIndicateurs As String)
    '-------------------------------------------------------------------------
    ' Crée les indicateurs visuels de validation des données
    '-------------------------------------------------------------------------
    
    Dim PlageTab As Range
    Set PlageTab = ws.Range(PlageIndicateurs)
    
    With PlageTab
        ' Titre
        .Cells(1, 1).Value = "STATUT VALIDATION"
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Size = 10
        
        ' Indicateurs
        .Cells(2, 1).Value = "Données complètes:"
        .Cells(2, 2).Value = "=SI(ET(NBVAL(D10:D16)>0,NBVAL(D22:D35)>0),""✓"",""✗"")"
        
        .Cells(3, 1).Value = "Budget équilibré:"
        .Cells(3, 2).Value = "=SI(C44>=0,""✓"",""✗"")"
        
        .Cells(4, 1).Value = "Épargne positive:"
        .Cells(4, 2).Value = "=SI(C45>0.1,""✓"",""✗"")"
        
        .Cells(5, 1).Value = "Prêt à sauvegarder:"
        .Cells(5, 2).Value = "=SI(ET(F42=TRUE,F43=TRUE),""OUI"",""NON"")"
        
        ' Formatage conditionnel
        .Font.Size = 9
        .Columns(2).HorizontalAlignment = xlCenter
    End With
    
End Sub

Sub CreerBoutonsAction(ws As Worksheet)
    '-------------------------------------------------------------------------
    ' Crée les boutons d'action pour la sauvegarde et navigation
    '-------------------------------------------------------------------------
    
    Dim btnSauvegarder As Shape, btnEffacer As Shape, btnRetour As Shape
    
    ' Bouton Sauvegarder
    Set btnSauvegarder = ws.Shapes.AddShape(msoShapeRoundedRectangle, 50, 320, 100, 30)
    With btnSauvegarder
        .Name = "Btn_Sauvegarder"
        .TextFrame.Characters.Text = "SAUVEGARDER"
        .TextFrame.Characters.Font.Size = 10
        .TextFrame.Characters.Font.Bold = True
        .Fill.ForeColor.RGB = RGB(112, 173, 71)
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .OnAction = "SauvegarderDonneesMensuelles"
    End With
    
    ' Bouton Effacer
    Set btnEffacer = ws.Shapes.AddShape(msoShapeRoundedRectangle, 160, 320, 100, 30)
    With btnEffacer
        .Name = "Btn_Effacer"
        .TextFrame.Characters.Text = "EFFACER"
        .TextFrame.Characters.Font.Size = 10
        .TextFrame.Characters.Font.Bold = True
        .Fill.ForeColor.RGB = RGB(196, 89, 17)
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .OnAction = "EffacerDonneesSaisie"
    End With
    
    ' Bouton Retour
    Set btnRetour = ws.Shapes.AddShape(msoShapeRoundedRectangle, 270, 320, 100, 30)
    With btnRetour
        .Name = "Btn_Retour"
        .TextFrame.Characters.Text = "TABLEAU DE BORD"
        .TextFrame.Characters.Font.Size = 10
        .TextFrame.Characters.Font.Bold = True
        .Fill.ForeColor.RGB = RGB(68, 114, 196)
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .OnAction = "AfficherTableauBord"
    End With
    
End Sub

'===============================================================================
' PROCEDURES D'ACTION ET VALIDATION
'===============================================================================

Sub ActualiserSaisieMensuelle()
    '-------------------------------------------------------------------------
    ' Met à jour la feuille de saisie avec les données du mois sélectionné
    '-------------------------------------------------------------------------
    
    On Error GoTo GestionErreur
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Saisie_Mensuelle")
    
    Application.ScreenUpdating = False
    
    ' Charger les données existantes pour le mois
    Call ChargerDonneesMensuelles(ws)
    
    ' Mettre à jour les formules et calculs
    ws.Calculate
    
    ' Actualiser la date de modification
    ws.Range("G5").Value = "Dernière modification: " & Format(Now, "dd/mm/yyyy hh:mm")
    
    Application.ScreenUpdating = True
    Exit Sub
    
GestionErreur:
    Application.ScreenUpdating = True
    Call EnregistrerJournal("Erreur actualisation saisie: " & Err.Description, "ERREUR")
End Sub

Sub SauvegarderDonneesMensuelles()
    '-------------------------------------------------------------------------
    ' Sauvegarde les données saisies dans les feuilles de données
    '-------------------------------------------------------------------------
    
    On Error GoTo GestionErreur
    
    ' Validation des données avant sauvegarde
    If Not ValiderDonneesSaisie() Then
        MsgBox "Les données saisies contiennent des erreurs. Veuillez corriger avant de sauvegarder.", vbExclamation, "Validation"
        Exit Sub
    End If
    
    ' Confirmation de sauvegarde
    If MsgBox("Confirmer la sauvegarde des données mensuelles ?", vbYesNo + vbQuestion, "Confirmation") = vbNo Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' Sauvegarde des revenus
    Call SauvegarderRevenus
    
    ' Sauvegarde des dépenses
    Call SauvegarderDepenses
    
    ' Mise à jour du tableau de bord
    Call ActualiserTableauBord
    
    Application.ScreenUpdating = True
    
    MsgBox "Données sauvegardées avec succès !", vbInformation, "Sauvegarde"
    Call EnregistrerJournal("Données mensuelles sauvegardées", "INFO")
    
    Exit Sub
    
GestionErreur:
    Application.ScreenUpdating = True
    MsgBox "Erreur lors de la sauvegarde: " & Err.Description, vbCritical, "Erreur"
    Call EnregistrerJournal("Erreur sauvegarde: " & Err.Description, "ERREUR")
End Sub

Function ValiderDonneesSaisie() As Boolean
    '-------------------------------------------------------------------------
    ' Valide la cohérence et complétude des données saisies
    '-------------------------------------------------------------------------
    
    Dim ws As Worksheet
    Dim i As Integer
    Dim MontantPrevu As Currency, MontantReel As Currency
    
    Set ws = ThisWorkbook.Worksheets("Saisie_Mensuelle")
    ValiderDonneesSaisie = True
    
    ' Validation des revenus
    For i = 10 To 16
        MontantPrevu = ws.Cells(i, 4).Value
        MontantReel = ws.Cells(i, 6).Value
        
        If MontantPrevu < 0 Or MontantReel < 0 Then
            ValiderDonneesSaisie = False
            Exit Function
        End If
        
        If MontantPrevu > MONTANT_MAXIMUM Or MontantReel > MONTANT_MAXIMUM Then
            ValiderDonneesSaisie = False
            Exit Function
        End If
    Next i
    
    ' Validation des dépenses
    For i = 22 To 35
        MontantPrevu = ws.Cells(i, 4).Value
        MontantReel = ws.Cells(i, 6).Value
        
        If MontantPrevu < 0 Or MontantReel < 0 Then
            ValiderDonneesSaisie = False
            Exit Function
        End If
        
        If MontantPrevu > MONTANT_MAXIMUM Or MontantReel > MONTANT_MAXIMUM Then
            ValiderDonneesSaisie = False
            Exit Function
        End If
    Next i
    
End Function

'===============================================================================
' PROCEDURES DE FORMATAGE
'===============================================================================

Sub AppliquerFormatageSaisie(ws As Worksheet)
    '-------------------------------------------------------------------------
    ' Applique le formatage final de la feuille de saisie
    '-------------------------------------------------------------------------
    
    With ws
        ' Police générale
        .Cells.Font.Name = "Segoe UI"
        .Cells.Font.Size = 9
        
        ' Largeur des colonnes optimisée
        .Columns("A").ColumnWidth = 15
        .Columns("B").ColumnWidth = 20
        .Columns("C").ColumnWidth = 10
        .Columns("D:F").ColumnWidth = 12
        .Columns("G").ColumnWidth = 10
        .Columns("H").ColumnWidth = 15
        
        ' Protection des cellules
        .Cells.Locked = True
        .Range("C5,B10:C16,D10:D16,F10:F16,H10:H16,B22:C35,D22:D35,F22:F35,H22:H35").Locked = False
    End With
    
End Sub

'===============================================================================
' FIN DU MODULE SAISIE
'===============================================================================
