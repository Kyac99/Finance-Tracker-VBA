Attribute VB_Name = "Module_Donnees"
'===============================================================================
' FINANCE TRACKER VBA - MODULE GESTION DES DONNÉES
' Version: 1.0
' Description: Gestion centralisée des données financières et persistance
' Fonction: CRUD des données revenus/dépenses et sauvegarde automatique
'===============================================================================

Option Explicit

' Constantes pour la gestion des données
Private Const CHEMIN_SAUVEGARDE As String = "Sauvegardes\"
Private Const EXTENSION_SAUVEGARDE As String = ".bak"

'===============================================================================
' PROCEDURES DE SAUVEGARDE DES DONNÉES
'===============================================================================

Sub SauvegarderRevenus()
    '-------------------------------------------------------------------------
    ' Sauvegarde les données de revenus depuis la feuille de saisie
    '-------------------------------------------------------------------------
    
    On Error GoTo GestionErreur
    
    Dim wsSaisie As Worksheet, wsDonnees As Worksheet
    Dim MoisReference As Date
    Dim i As Integer
    Dim Categorie As String, Description As String
    Dim MontantPrevu As Currency, MontantReel As Currency
    Dim EstRecurrent As Boolean
    
    Set wsSaisie = ThisWorkbook.Worksheets("Saisie_Mensuelle")
    Set wsDonnees = ThisWorkbook.Worksheets("Donnees_Revenus")
    
    ' Initialiser la feuille de données si nécessaire
    Call InitialiserFeuilleDonnees(wsDonnees, "REVENUS")
    
    ' Récupérer le mois de référence
    MoisReference = ObtenirMoisCourant()
    
    ' Supprimer les anciennes données du mois
    Call SupprimerDonneesMois(wsDonnees, MoisReference)
    
    ' Sauvegarder les nouvelles données (lignes 10 à 16 pour les revenus)
    For i = 10 To 16
        Categorie = wsSaisie.Cells(i, 1).Value
        Description = wsSaisie.Cells(i, 2).Value
        EstRecurrent = (wsSaisie.Cells(i, 3).Value = "OUI")
        MontantPrevu = wsSaisie.Cells(i, 4).Value
        MontantReel = wsSaisie.Cells(i, 6).Value
        
        ' Sauvegarder seulement si il y a des données
        If Len(Categorie) > 0 And (MontantPrevu > 0 Or MontantReel > 0) Then
            Call AjouterDonneeFinanciere(wsDonnees, MoisReference, Categorie, _
                                        Description, EstRecurrent, MontantPrevu, MontantReel)
        End If
    Next i
    
    Call EnregistrerJournal("Revenus sauvegardés pour " & Format(MoisReference, "mm/yyyy"), "INFO")
    Exit Sub
    
GestionErreur:
    Call EnregistrerJournal("Erreur sauvegarde revenus: " & Err.Description, "ERREUR")
End Sub

Sub SauvegarderDepenses()
    '-------------------------------------------------------------------------
    ' Sauvegarde les données de dépenses depuis la feuille de saisie
    '-------------------------------------------------------------------------
    
    On Error GoTo GestionErreur
    
    Dim wsSaisie As Worksheet, wsDonnees As Worksheet
    Dim MoisReference As Date
    Dim i As Integer
    Dim Categorie As String, Description As String
    Dim MontantPrevu As Currency, MontantReel As Currency
    Dim EstRecurrent As Boolean
    
    Set wsSaisie = ThisWorkbook.Worksheets("Saisie_Mensuelle")
    Set wsDonnees = ThisWorkbook.Worksheets("Donnees_Depenses")
    
    ' Initialiser la feuille de données si nécessaire
    Call InitialiserFeuilleDonnees(wsDonnees, "DÉPENSES")
    
    ' Récupérer le mois de référence
    MoisReference = ObtenirMoisCourant()
    
    ' Supprimer les anciennes données du mois
    Call SupprimerDonneesMois(wsDonnees, MoisReference)
    
    ' Sauvegarder les nouvelles données (lignes 22 à 35 pour les dépenses)
    For i = 22 To 35
        Categorie = wsSaisie.Cells(i, 1).Value
        Description = wsSaisie.Cells(i, 2).Value
        EstRecurrent = (wsSaisie.Cells(i, 3).Value = "OUI")
        MontantPrevu = wsSaisie.Cells(i, 4).Value
        MontantReel = wsSaisie.Cells(i, 6).Value
        
        ' Sauvegarder seulement si il y a des données
        If Len(Categorie) > 0 And (MontantPrevu > 0 Or MontantReel > 0) Then
            Call AjouterDonneeFinanciere(wsDonnees, MoisReference, Categorie, _
                                        Description, EstRecurrent, MontantPrevu, MontantReel)
        End If
    Next i
    
    Call EnregistrerJournal("Dépenses sauvegardées pour " & Format(MoisReference, "mm/yyyy"), "INFO")
    Exit Sub
    
GestionErreur:
    Call EnregistrerJournal("Erreur sauvegarde dépenses: " & Err.Description, "ERREUR")
End Sub

Sub InitialiserFeuilleDonnees(ws As Worksheet, TypeDonnees As String)
    '-------------------------------------------------------------------------
    ' Initialise la structure d'une feuille de données
    '-------------------------------------------------------------------------
    
    If ws.Cells(1, 1).Value = "" Then
        With ws
            .Cells(1, 1).Value = "DATE"
            .Cells(1, 2).Value = "CATÉGORIE"
            .Cells(1, 3).Value = "DESCRIPTION"
            .Cells(1, 4).Value = "RÉCURRENT"
            .Cells(1, 5).Value = "MONTANT PRÉVU"
            .Cells(1, 6).Value = "MONTANT RÉEL"
            .Cells(1, 7).Value = "ÉCART"
            .Cells(1, 8).Value = "NOTES"
            
            ' Formatage des en-têtes
            .Range("A1:H1").Font.Bold = True
            If TypeDonnees = "REVENUS" Then
                .Range("A1:H1").Interior.Color = RGB(68, 114, 196)
            Else
                .Range("A1:H1").Interior.Color = RGB(196, 89, 17)
            End If
            .Range("A1:H1").Font.Color = RGB(255, 255, 255)
            .Range("A1:H1").Borders.LineStyle = xlContinuous
            
            ' Ajustement des colonnes
            .Columns("A").ColumnWidth = 12
            .Columns("B").ColumnWidth = 20
            .Columns("C").ColumnWidth = 25
            .Columns("D").ColumnWidth = 10
            .Columns("E:G").ColumnWidth = 15
            .Columns("H").ColumnWidth = 30
        End With
    End If
    
End Sub

Sub AjouterDonneeFinanciere(ws As Worksheet, DateDonnee As Date, Categorie As String, _
                           Description As String, EstRecurrent As Boolean, _
                           MontantPrevu As Currency, MontantReel As Currency)
    '-------------------------------------------------------------------------
    ' Ajoute une ligne de données financières
    '-------------------------------------------------------------------------
    
    Dim DerniereLigne As Long
    
    DerniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    With ws
        .Cells(DerniereLigne, 1).Value = DateDonnee
        .Cells(DerniereLigne, 2).Value = Categorie
        .Cells(DerniereLigne, 3).Value = Description
        .Cells(DerniereLigne, 4).Value = IIf(EstRecurrent, "OUI", "NON")
        .Cells(DerniereLigne, 5).Value = MontantPrevu
        .Cells(DerniereLigne, 6).Value = MontantReel
        .Cells(DerniereLigne, 7).Value = MontantReel - MontantPrevu
        .Cells(DerniereLigne, 8).Value = ""
        
        ' Formatage
        .Cells(DerniereLigne, 1).NumberFormat = "dd/mm/yyyy"
        .Range(.Cells(DerniereLigne, 5), .Cells(DerniereLigne, 7)).NumberFormat = "#,##0.00 €"
        .Range(.Cells(DerniereLigne, 1), .Cells(DerniereLigne, 8)).Borders.LineStyle = xlContinuous
        .Range(.Cells(DerniereLigne, 1), .Cells(DerniereLigne, 8)).Font.Size = 9
    End With
    
End Sub

'===============================================================================
' PROCEDURES DE CHARGEMENT DES DONNÉES
'===============================================================================

Sub ChargerDonneesMensuelles(ws As Worksheet)
    '-------------------------------------------------------------------------
    ' Charge les données existantes dans la feuille de saisie
    '-------------------------------------------------------------------------
    
    On Error GoTo GestionErreur
    
    Dim MoisReference As Date
    
    MoisReference = ObtenirMoisCourant()
    
    ' Charger les revenus
    Call ChargerRevenus(ws, MoisReference)
    
    ' Charger les dépenses
    Call ChargerDepenses(ws, MoisReference)
    
    Call EnregistrerJournal("Données mensuelles chargées pour " & Format(MoisReference, "mm/yyyy"), "INFO")
    Exit Sub
    
GestionErreur:
    Call EnregistrerJournal("Erreur chargement données: " & Err.Description, "ERREUR")
End Sub

Sub ChargerRevenus(ws As Worksheet, MoisReference As Date)
    '-------------------------------------------------------------------------
    ' Charge les données de revenus pour un mois donné
    '-------------------------------------------------------------------------
    
    Dim wsDonnees As Worksheet
    Dim DerniereLigne As Long
    Dim i As Long, LigneSaisie As Integer
    Dim Categorie As String
    
    Set wsDonnees = ThisWorkbook.Worksheets("Donnees_Revenus")
    
    If wsDonnees.Cells(1, 1).Value = "" Then Exit Sub
    
    DerniereLigne = wsDonnees.Cells(wsDonnees.Rows.Count, 1).End(xlUp).Row
    
    ' Effacer les données actuelles de la saisie (revenus)
    ws.Range("B10:C16,D10:D16,F10:F16,H10:H16").ClearContents
    
    ' Charger les données du mois
    For i = 2 To DerniereLigne
        If Month(wsDonnees.Cells(i, 1).Value) = Month(MoisReference) And _
           Year(wsDonnees.Cells(i, 1).Value) = Year(MoisReference) Then
            
            Categorie = wsDonnees.Cells(i, 2).Value
            LigneSaisie = TrouverLigneCategorie(ws, Categorie, 10, 16)
            
            If LigneSaisie > 0 Then
                ws.Cells(LigneSaisie, 2).Value = wsDonnees.Cells(i, 3).Value ' Description
                ws.Cells(LigneSaisie, 3).Value = wsDonnees.Cells(i, 4).Value ' Récurrent
                ws.Cells(LigneSaisie, 4).Value = wsDonnees.Cells(i, 5).Value ' Montant prévu
                ws.Cells(LigneSaisie, 6).Value = wsDonnees.Cells(i, 6).Value ' Montant réel
                ws.Cells(LigneSaisie, 8).Value = wsDonnees.Cells(i, 8).Value ' Notes
            End If
        End If
    Next i
    
End Sub

Sub ChargerDepenses(ws As Worksheet, MoisReference As Date)
    '-------------------------------------------------------------------------
    ' Charge les données de dépenses pour un mois donné
    '-------------------------------------------------------------------------
    
    Dim wsDonnees As Worksheet
    Dim DerniereLigne As Long
    Dim i As Long, LigneSaisie As Integer
    Dim Categorie As String
    
    Set wsDonnees = ThisWorkbook.Worksheets("Donnees_Depenses")
    
    If wsDonnees.Cells(1, 1).Value = "" Then Exit Sub
    
    DerniereLigne = wsDonnees.Cells(wsDonnees.Rows.Count, 1).End(xlUp).Row
    
    ' Effacer les données actuelles de la saisie (dépenses)
    ws.Range("B22:C35,D22:D35,F22:F35,H22:H35").ClearContents
    
    ' Charger les données du mois
    For i = 2 To DerniereLigne
        If Month(wsDonnees.Cells(i, 1).Value) = Month(MoisReference) And _
           Year(wsDonnees.Cells(i, 1).Value) = Year(MoisReference) Then
            
            Categorie = wsDonnees.Cells(i, 2).Value
            LigneSaisie = TrouverLigneCategorie(ws, Categorie, 22, 35)
            
            If LigneSaisie > 0 Then
                ws.Cells(LigneSaisie, 2).Value = wsDonnees.Cells(i, 3).Value ' Description
                ws.Cells(LigneSaisie, 3).Value = wsDonnees.Cells(i, 4).Value ' Récurrent
                ws.Cells(LigneSaisie, 4).Value = wsDonnees.Cells(i, 5).Value ' Montant prévu
                ws.Cells(LigneSaisie, 6).Value = wsDonnees.Cells(i, 6).Value ' Montant réel
                ws.Cells(LigneSaisie, 8).Value = wsDonnees.Cells(i, 8).Value ' Notes
            End If
        End If
    Next i
    
End Sub

Function TrouverLigneCategorie(ws As Worksheet, Categorie As String, LigneDebut As Integer, LigneFin As Integer) As Integer
    '-------------------------------------------------------------------------
    ' Trouve la ligne correspondant à une catégorie dans la feuille de saisie
    '-------------------------------------------------------------------------
    
    Dim i As Integer
    
    For i = LigneDebut To LigneFin
        If ws.Cells(i, 1).Value = Categorie Then
            TrouverLigneCategorie = i
            Exit Function
        End If
    Next i
    
    TrouverLigneCategorie = 0
    
End Function

'===============================================================================
' PROCEDURES DE SUPPRESSION ET NETTOYAGE
'===============================================================================

Sub SupprimerDonneesMois(ws As Worksheet, MoisReference As Date)
    '-------------------------------------------------------------------------
    ' Supprime toutes les données d'un mois spécifique
    '-------------------------------------------------------------------------
    
    Dim DerniereLigne As Long
    Dim i As Long
    
    If ws.Cells(1, 1).Value = "" Then Exit Sub
    
    DerniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Parcourir de bas en haut pour éviter les problèmes de suppression
    For i = DerniereLigne To 2 Step -1
        If Month(ws.Cells(i, 1).Value) = Month(MoisReference) And _
           Year(ws.Cells(i, 1).Value) = Year(MoisReference) Then
            ws.Rows(i).Delete
        End If
    Next i
    
End Sub

Sub NettoyerDonneesAnciennes(NbMoisConserver As Integer)
    '-------------------------------------------------------------------------
    ' Nettoie les données antérieures à un nombre de mois donné
    '-------------------------------------------------------------------------
    
    On Error GoTo GestionErreur
    
    Dim DateLimite As Date
    Dim wsRevenus As Worksheet, wsDepenses As Worksheet
    
    Set wsRevenus = ThisWorkbook.Worksheets("Donnees_Revenus")
    Set wsDepenses = ThisWorkbook.Worksheets("Donnees_Depenses")
    
    DateLimite = DateAdd("m", -NbMoisConserver, ObtenirMoisCourant())
    
    ' Archiver les données avant suppression
    Call ArchiverDonneesAnciennes(DateLimite)
    
    ' Nettoyer les revenus
    Call SupprimerDonneesAvantDate(wsRevenus, DateLimite)
    
    ' Nettoyer les dépenses
    Call SupprimerDonneesAvantDate(wsDepenses, DateLimite)
    
    Call EnregistrerJournal("Nettoyage des données antérieures à " & Format(DateLimite, "mm/yyyy"), "INFO")
    Exit Sub
    
GestionErreur:
    Call EnregistrerJournal("Erreur nettoyage données: " & Err.Description, "ERREUR")
End Sub

Sub SupprimerDonneesAvantDate(ws As Worksheet, DateLimite As Date)
    '-------------------------------------------------------------------------
    ' Supprime les données antérieures à une date limite
    '-------------------------------------------------------------------------
    
    Dim DerniereLigne As Long
    Dim i As Long
    Dim CompteurSuppressions As Integer
    
    If ws.Cells(1, 1).Value = "" Then Exit Sub
    
    DerniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    CompteurSuppressions = 0
    
    For i = DerniereLigne To 2 Step -1
        If ws.Cells(i, 1).Value < DateLimite Then
            ws.Rows(i).Delete
            CompteurSuppressions = CompteurSuppressions + 1
        End If
    Next i
    
    If CompteurSuppressions > 0 Then
        Call EnregistrerJournal(CompteurSuppressions & " lignes supprimées de " & ws.Name, "INFO")
    End If
    
End Sub

'===============================================================================
' PROCEDURES D'ARCHIVAGE ET SAUVEGARDE
'===============================================================================

Sub ArchiverDonneesAnciennes(DateLimite As Date)
    '-------------------------------------------------------------------------
    ' Archive les données anciennes avant leur suppression
    '-------------------------------------------------------------------------
    
    On Error GoTo GestionErreur
    
    Dim wsArchives As Worksheet
    Dim wsRevenus As Worksheet, wsDepenses As Worksheet
    Dim DerniereLigne As Long, i As Long
    Dim LigneArchive As Long
    
    Set wsArchives = ThisWorkbook.Worksheets("Archives")
    Set wsRevenus = ThisWorkbook.Worksheets("Donnees_Revenus")
    Set wsDepenses = ThisWorkbook.Worksheets("Donnees_Depenses")
    
    ' Initialiser la feuille d'archives si nécessaire
    Call InitialiserFeuilleArchives(wsArchives)
    
    LigneArchive = wsArchives.Cells(wsArchives.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Archiver les revenus
    Call ArchiverDonneesFeuille(wsRevenus, wsArchives, DateLimite, LigneArchive, "REVENU")
    
    ' Archiver les dépenses
    LigneArchive = wsArchives.Cells(wsArchives.Rows.Count, 1).End(xlUp).Row + 1
    Call ArchiverDonneesFeuille(wsDepenses, wsArchives, DateLimite, LigneArchive, "DÉPENSE")
    
    Exit Sub
    
GestionErreur:
    Call EnregistrerJournal("Erreur archivage: " & Err.Description, "ERREUR")
End Sub

Sub InitialiserFeuilleArchives(ws As Worksheet)
    '-------------------------------------------------------------------------
    ' Initialise la structure de la feuille d'archives
    '-------------------------------------------------------------------------
    
    If ws.Cells(1, 1).Value = "" Then
        With ws
            .Cells(1, 1).Value = "DATE_ARCHIVAGE"
            .Cells(1, 2).Value = "TYPE"
            .Cells(1, 3).Value = "DATE_ORIGINALE"
            .Cells(1, 4).Value = "CATÉGORIE"
            .Cells(1, 5).Value = "DESCRIPTION"
            .Cells(1, 6).Value = "MONTANT_PREVU"
            .Cells(1, 7).Value = "MONTANT_REEL"
            .Cells(1, 8).Value = "NOTES"
            
            ' Formatage
            .Range("A1:H1").Font.Bold = True
            .Range("A1:H1").Interior.Color = RGB(128, 128, 128)
            .Range("A1:H1").Font.Color = RGB(255, 255, 255)
            .Range("A1:H1").Borders.LineStyle = xlContinuous
        End With
    End If
    
End Sub

Sub ArchiverDonneesFeuille(wsSource As Worksheet, wsArchives As Worksheet, _
                          DateLimite As Date, ByRef LigneArchive As Long, TypeDonnee As String)
    '-------------------------------------------------------------------------
    ' Archive les données d'une feuille spécifique
    '-------------------------------------------------------------------------
    
    Dim DerniereLigne As Long
    Dim i As Long
    
    If wsSource.Cells(1, 1).Value = "" Then Exit Sub
    
    DerniereLigne = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To DerniereLigne
        If wsSource.Cells(i, 1).Value < DateLimite Then
            With wsArchives
                .Cells(LigneArchive, 1).Value = Now ' Date d'archivage
                .Cells(LigneArchive, 2).Value = TypeDonnee
                .Cells(LigneArchive, 3).Value = wsSource.Cells(i, 1).Value ' Date originale
                .Cells(LigneArchive, 4).Value = wsSource.Cells(i, 2).Value ' Catégorie
                .Cells(LigneArchive, 5).Value = wsSource.Cells(i, 3).Value ' Description
                .Cells(LigneArchive, 6).Value = wsSource.Cells(i, 5).Value ' Montant prévu
                .Cells(LigneArchive, 7).Value = wsSource.Cells(i, 6).Value ' Montant réel
                .Cells(LigneArchive, 8).Value = wsSource.Cells(i, 8).Value ' Notes
                
                ' Formatage
                .Cells(LigneArchive, 1).NumberFormat = "dd/mm/yyyy hh:mm"
                .Cells(LigneArchive, 3).NumberFormat = "dd/mm/yyyy"
                .Range(.Cells(LigneArchive, 6), .Cells(LigneArchive, 7)).NumberFormat = "#,##0.00 €"
            End With
            
            LigneArchive = LigneArchive + 1
        End If
    Next i
    
End Sub

Sub CreerSauvegardeComplete()
    '-------------------------------------------------------------------------
    ' Crée une sauvegarde complète du fichier
    '-------------------------------------------------------------------------
    
    On Error GoTo GestionErreur
    
    Dim CheminSauvegarde As String
    Dim NomFichierSauvegarde As String
    
    NomFichierSauvegarde = "FinanceTracker_" & Format(Now, "yyyy-mm-dd_hh-mm") & EXTENSION_SAUVEGARDE
    CheminSauvegarde = ThisWorkbook.Path & "\" & CHEMIN_SAUVEGARDE & NomFichierSauvegarde
    
    ' Créer le répertoire de sauvegarde s'il n'existe pas
    If Dir(ThisWorkbook.Path & "\" & CHEMIN_SAUVEGARDE, vbDirectory) = "" Then
        MkDir ThisWorkbook.Path & "\" & CHEMIN_SAUVEGARDE
    End If
    
    ' Créer la sauvegarde
    ThisWorkbook.SaveCopyAs CheminSauvegarde
    
    Call EnregistrerJournal("Sauvegarde complète créée: " & NomFichierSauvegarde, "INFO")
    Exit Sub
    
GestionErreur:
    Call EnregistrerJournal("Erreur sauvegarde complète: " & Err.Description, "ERREUR")
End Sub

'===============================================================================
' PROCEDURES UTILITAIRES
'===============================================================================

Sub EffacerDonneesSaisie()
    '-------------------------------------------------------------------------
    ' Efface toutes les données de la feuille de saisie
    '-------------------------------------------------------------------------
    
    Dim ws As Worksheet
    Dim Reponse As Integer
    
    Set ws = ThisWorkbook.Worksheets("Saisie_Mensuelle")
    
    Reponse = MsgBox("Êtes-vous sûr de vouloir effacer toutes les données de saisie ?", _
                     vbYesNo + vbExclamation, "Confirmation")
    
    If Reponse = vbYes Then
        ws.Range("B10:C16,D10:D16,F10:F16,H10:H16").ClearContents ' Revenus
        ws.Range("B22:C35,D22:D35,F22:F35,H22:H35").ClearContents ' Dépenses
        
        Call EnregistrerJournal("Données de saisie effacées", "INFO")
        MsgBox "Données de saisie effacées.", vbInformation
    End If
    
End Sub

Sub CreerFeuilleGenerique(ws As Worksheet, TypeFeuille As String)
    '-------------------------------------------------------------------------
    ' Crée une feuille générique selon son type
    '-------------------------------------------------------------------------
    
    With ws
        .Cells.Clear
        .Range("A1").Value = "FEUILLE " & UCase(TypeFeuille)
        .Range("A1").Font.Size = 14
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Color = RGB(68, 114, 196)
        
        Select Case TypeFeuille
            Case "Donnees_Revenus"
                Call InitialiserFeuilleDonnees(ws, "REVENUS")
            Case "Donnees_Depenses"
                Call InitialiserFeuilleDonnees(ws, "DÉPENSES")
            Case "Archives"
                Call InitialiserFeuilleArchives(ws)
        End Select
    End With
    
End Sub

'===============================================================================
' FIN DU MODULE DONNÉES
'===============================================================================
