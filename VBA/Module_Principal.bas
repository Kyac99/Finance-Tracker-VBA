Attribute VB_Name = "Module_Principal"
'===============================================================================
' FINANCE TRACKER VBA - MODULE PRINCIPAL
' Version: 1.0
' Description: Module principal pour la gestion du système de finances personnelles
' Auteur: Système automatisé
' Date: Juillet 2025
'===============================================================================

Option Explicit

' Variables globales pour la configuration
Public Const VERSION_APP As String = "1.0"
Public Const NOM_FICHIER As String = "FinanceTracker"

' Énumérations pour les types de données
Public Enum TypeTransaction
    Revenu = 1
    Depense = 2
End Enum

Public Enum Frequence
    Mensuelle = 1
    Trimestrielle = 3
    Annuelle = 12
End Enum

' Structure pour les données financières
Public Type DonneesFinancieres
    Mois As Date
    CategorieID As Long
    TypeTrans As TypeTransaction
    MontantPrevu As Currency
    MontantReel As Currency
    Description As String
    EstRecurrent As Boolean
End Type

'===============================================================================
' PROCEDURES D'INITIALISATION
'===============================================================================

Sub InitialiserApplication()
    '-------------------------------------------------------------------------
    ' Initialise l'application et configure l'environnement de travail
    '-------------------------------------------------------------------------
    
    On Error GoTo GestionErreur
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Vérifier et créer la structure des feuilles
    Call CreerStructureFeuilles
    
    ' Initialiser les données de base
    Call InitialiserDonneesBase
    
    ' Configurer la navigation
    Call ConfigurerNavigation
    
    ' Activer le tableau de bord
    Call AfficherTableauBord
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    MsgBox "Application Finance Tracker initialisée avec succès !" & vbCrLf & _
           "Version: " & VERSION_APP, vbInformation, "Initialisation"
    
    Exit Sub
    
GestionErreur:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Erreur lors de l'initialisation: " & Err.Description, vbCritical, "Erreur"
End Sub

Sub CreerStructureFeuilles()
    '-------------------------------------------------------------------------
    ' Crée et configure toutes les feuilles nécessaires au système
    '-------------------------------------------------------------------------
    
    Dim ws As Worksheet
    Dim NomsFeuillesReq As Variant
    Dim i As Integer
    
    ' Liste des feuilles requises
    NomsFeuillesReq = Array("Dashboard", "Saisie_Mensuelle", "Donnees_Revenus", _
                           "Donnees_Depenses", "Categories", "Parametres", _
                           "Rapports", "Archives")
    
    ' Créer les feuilles si elles n'existent pas
    For i = 0 To UBound(NomsFeuillesReq)
        If Not FeuilleExiste(NomsFeuillesReq(i)) Then
            Set ws = ThisWorkbook.Worksheets.Add
            ws.Name = NomsFeuillesReq(i)
            Call ConfigurerFeuille(ws, NomsFeuillesReq(i))
        End If
    Next i
    
End Sub

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

Sub ConfigurerFeuille(ws As Worksheet, TypeFeuille As String)
    '-------------------------------------------------------------------------
    ' Configure une feuille selon son type
    '-------------------------------------------------------------------------
    
    With ws
        .Cells.Clear
        
        Select Case TypeFeuille
            Case "Dashboard"
                Call CreerTableauBord(ws)
            Case "Saisie_Mensuelle"
                Call CreerFeuilleSaisie(ws)
            Case "Categories"
                Call CreerFeuilleCategories(ws)
            Case "Parametres"
                Call CreerFeuilleParametres(ws)
            Case Else
                Call CreerFeuilleGenerique(ws, TypeFeuille)
        End Select
        
        ' Protection de base
        .Protect Password:="FinanceTracker2025", _
                DrawingObjects:=False, _
                Contents:=True, _
                Scenarios:=False, _
                UserInterfaceOnly:=True
    End With
    
End Sub

'===============================================================================
' PROCEDURES DE NAVIGATION
'===============================================================================

Sub ConfigurerNavigation()
    '-------------------------------------------------------------------------
    ' Configure les boutons de navigation et les raccourcis
    '-------------------------------------------------------------------------
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    
    ' Supprimer les anciens boutons de navigation s'ils existent
    Call SupprimerAnciensBoutons(ws)
    
    ' Créer les nouveaux boutons
    Call CreerBoutonsNavigation(ws)
    
End Sub

Sub SupprimerAnciensBoutons(ws As Worksheet)
    '-------------------------------------------------------------------------
    ' Supprime les anciens boutons de navigation
    '-------------------------------------------------------------------------
    
    Dim shp As Shape
    
    For Each shp In ws.Shapes
        If Left(shp.Name, 4) = "Btn_" Then
            shp.Delete
        End If
    Next shp
    
End Sub

Sub CreerBoutonsNavigation(ws As Worksheet)
    '-------------------------------------------------------------------------
    ' Crée les boutons de navigation sur le tableau de bord
    '-------------------------------------------------------------------------
    
    Dim btnSaisie As Shape, btnRapports As Shape, btnParametres As Shape
    
    With ws
        ' Bouton Saisie Mensuelle
        Set btnSaisie = .Shapes.AddShape(msoShapeRoundedRectangle, 50, 450, 120, 35)
        With btnSaisie
            .Name = "Btn_Saisie"
            .TextFrame.Characters.Text = "Saisie Mensuelle"
            .TextFrame.Characters.Font.Size = 10
            .TextFrame.Characters.Font.Bold = True
            .Fill.ForeColor.RGB = RGB(68, 114, 196)
            .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
            .OnAction = "OuvrirSaisieMensuelle"
        End With
        
        ' Bouton Rapports
        Set btnRapports = .Shapes.AddShape(msoShapeRoundedRectangle, 180, 450, 120, 35)
        With btnRapports
            .Name = "Btn_Rapports"
            .TextFrame.Characters.Text = "Rapports"
            .TextFrame.Characters.Font.Size = 10
            .TextFrame.Characters.Font.Bold = True
            .Fill.ForeColor.RGB = RGB(112, 173, 71)
            .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
            .OnAction = "OuvrirRapports"
        End With
        
        ' Bouton Paramètres
        Set btnParametres = .Shapes.AddShape(msoShapeRoundedRectangle, 310, 450, 120, 35)
        With btnParametres
            .Name = "Btn_Parametres"
            .TextFrame.Characters.Text = "Paramètres"
            .TextFrame.Characters.Font.Size = 10
            .TextFrame.Characters.Font.Bold = True
            .Fill.ForeColor.RGB = RGB(255, 192, 0)
            .TextFrame.Characters.Font.Color = RGB(0, 0, 0)
            .OnAction = "OuvrirParametres"
        End With
    End With
    
End Sub

'===============================================================================
' PROCEDURES DE NAVIGATION - ACTIONS
'===============================================================================

Sub AfficherTableauBord()
    '-------------------------------------------------------------------------
    ' Active et affiche le tableau de bord principal
    '-------------------------------------------------------------------------
    
    ThisWorkbook.Worksheets("Dashboard").Activate
    Call ActualiserTableauBord
    
End Sub

Sub OuvrirSaisieMensuelle()
    '-------------------------------------------------------------------------
    ' Ouvre la feuille de saisie mensuelle
    '-------------------------------------------------------------------------
    
    ThisWorkbook.Worksheets("Saisie_Mensuelle").Activate
    Call ActualiserSaisieMensuelle
    
End Sub

Sub OuvrirRapports()
    '-------------------------------------------------------------------------
    ' Ouvre la feuille des rapports
    '-------------------------------------------------------------------------
    
    ThisWorkbook.Worksheets("Rapports").Activate
    Call GenererRapportMensuel
    
End Sub

Sub OuvrirParametres()
    '-------------------------------------------------------------------------
    ' Ouvre la feuille des paramètres
    '-------------------------------------------------------------------------
    
    ThisWorkbook.Worksheets("Parametres").Activate
    
End Sub

'===============================================================================
' UTILITAIRES GENERAUX
'===============================================================================

Function ObtenirMoisCourant() As Date
    '-------------------------------------------------------------------------
    ' Retourne le premier jour du mois courant
    '-------------------------------------------------------------------------
    
    ObtenirMoisCourant = DateSerial(Year(Date), Month(Date), 1)
    
End Function

Function FormaterMontant(Montant As Currency) As String
    '-------------------------------------------------------------------------
    ' Formate un montant en devise avec symbole
    '-------------------------------------------------------------------------
    
    FormaterMontant = Format(Montant, "#,##0.00 €")
    
End Function

Sub EnregistrerJournal(Message As String, Optional TypeLog As String = "INFO")
    '-------------------------------------------------------------------------
    ' Enregistre un message dans le journal d'activité
    '-------------------------------------------------------------------------
    
    Dim ws As Worksheet
    Dim DerniereLigne As Long
    
    Set ws = ThisWorkbook.Worksheets("Archives")
    DerniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    With ws
        .Cells(DerniereLigne, 1).Value = Now
        .Cells(DerniereLigne, 2).Value = TypeLog
        .Cells(DerniereLigne, 3).Value = Message
    End With
    
End Sub

Sub InitialiserDonneesBase()
    '-------------------------------------------------------------------------
    ' Initialise les données de base si nécessaires
    '-------------------------------------------------------------------------
    
    Call InitialiserCategories
    Call InitialiserParametres
    
End Sub

'===============================================================================
' FIN DU MODULE PRINCIPAL
'===============================================================================
