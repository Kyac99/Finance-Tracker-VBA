Attribute VB_Name = "Module_Categories"
'===============================================================================
' FINANCE TRACKER VBA - MODULE CATÉGORIES ET PARAMÈTRES
' Version: 1.0
' Description: Gestion des catégories financières et paramètres système
' Fonction: Configuration personnalisée des catégories et réglages utilisateur
'===============================================================================

Option Explicit

' Structure pour une catégorie financière
Public Type CategorieFinanciere
    ID As Long
    Nom As String
    Type As TypeTransaction ' Revenu ou Dépense
    CouleurAffichage As Long
    BudgetDefaut As Currency
    EstActive As Boolean
    EstPersonnalisee As Boolean
    Description As String
End Type

' Constantes pour les paramètres système
Private Const MAX_CATEGORIES As Integer = 50
Private Const LONGUEUR_MAX_NOM_CATEGORIE As Integer = 30

'===============================================================================
' PROCEDURES D'INITIALISATION DES CATEGORIES
'===============================================================================

Sub InitialiserCategories()
    '-------------------------------------------------------------------------
    ' Initialise les catégories par défaut du système
    '-------------------------------------------------------------------------
    
    On Error GoTo GestionErreur
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Categories")
    
    ' Création de la structure de la feuille catégories
    Call CreerStructureFeuilleCategories(ws)
    
    ' Vérifier si les catégories existent déjà
    If ws.Cells(2, 1).Value = "" Then
        ' Initialiser les catégories par défaut
        Call CreerCategoriesDefaut(ws)
        Call EnregistrerJournal("Catégories par défaut initialisées", "INFO")
    End If
    
    Exit Sub
    
GestionErreur:
    Call EnregistrerJournal("Erreur initialisation catégories: " & Err.Description, "ERREUR")
End Sub

Sub CreerFeuilleCategories(ws As Worksheet)
    '-------------------------------------------------------------------------
    ' Crée la feuille de gestion des catégories
    '-------------------------------------------------------------------------
    
    Call CreerStructureFeuilleCategories(ws)
    Call CreerCategoriesDefaut(ws)
    
End Sub

Sub CreerStructureFeuilleCategories(ws As Worksheet)
    '-------------------------------------------------------------------------
    ' Crée la structure de la feuille de gestion des catégories
    '-------------------------------------------------------------------------
    
    With ws
        .Cells.Clear
        .Tab.Color = RGB(255, 192, 0)
        
        ' En-tête principal
        .Range("A1:H1").Merge
        .Range("A1").Value = "GESTION DES CATÉGORIES FINANCIÈRES"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Color = RGB(255, 192, 0)
        .Range("A1").HorizontalAlignment = xlCenter
        
        ' En-têtes des colonnes
        .Cells(3, 1).Value = "ID"
        .Cells(3, 2).Value = "NOM DE LA CATÉGORIE"
        .Cells(3, 3).Value = "TYPE"
        .Cells(3, 4).Value = "COULEUR"
        .Cells(3, 5).Value = "BUDGET DÉFAUT"
        .Cells(3, 6).Value = "ACTIVE"
        .Cells(3, 7).Value = "PERSONNALISÉE"
        .Cells(3, 8).Value = "DESCRIPTION"
        
        ' Formatage des en-têtes
        .Range("A3:H3").Font.Bold = True
        .Range("A3:H3").Interior.Color = RGB(255, 192, 0)
        .Range("A3:H3").Font.Color = RGB(0, 0, 0)
        .Range("A3:H3").Borders.LineStyle = xlContinuous
        
        ' Ajustement des largeurs de colonnes
        .Columns("A").ColumnWidth = 5
        .Columns("B").ColumnWidth = 25
        .Columns("C").ColumnWidth = 12
        .Columns("D").ColumnWidth = 10
        .Columns("E").ColumnWidth = 15
        .Columns("F").ColumnWidth = 8
        .Columns("G").ColumnWidth = 12
        .Columns("H").ColumnWidth = 30
    End With
    
End Sub

Function ObtenirListeCategories(TypeCategorie As TypeTransaction) As Variant
    '-------------------------------------------------------------------------
    ' Retourne la liste des catégories actives d'un type donné
    '-------------------------------------------------------------------------
    
    Dim ws As Worksheet
    Dim DerniereLigne As Long
    Dim i As Long
    Dim ListeCategories() As String
    Dim CompteurCategories As Integer
    Dim TypeRecherche As String
    
    Set ws = ThisWorkbook.Worksheets("Categories")
    
    If TypeCategorie = Revenu Then
        TypeRecherche = "Revenu"
    Else
        TypeRecherche = "Dépense"
    End If
    
    DerniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ReDim ListeCategories(1 To 1)
    CompteurCategories = 0
    
    For i = 4 To DerniereLigne
        If ws.Cells(i, 3).Value = TypeRecherche And ws.Cells(i, 6).Value = "OUI" Then
            CompteurCategories = CompteurCategories + 1
            ReDim Preserve ListeCategories(1 To CompteurCategories)
            ListeCategories(CompteurCategories) = ws.Cells(i, 2).Value
        End If
    Next i
    
    If CompteurCategories > 0 Then
        ObtenirListeCategories = ListeCategories
    Else
        ObtenirListeCategories = Array()
    End If
    
End Function

Sub CreerCategoriesDefaut(ws As Worksheet)
    '-------------------------------------------------------------------------
    ' Crée les catégories financières par défaut
    '-------------------------------------------------------------------------
    
    Dim LigneActuelle As Integer
    LigneActuelle = 4
    
    ' Catégories de revenus
    Call AjouterCategorie(ws, LigneActuelle, 1, "Salaire Principal", "Revenu", RGB(68, 114, 196), 3000, True, False, "Salaire principal du foyer")
    LigneActuelle = LigneActuelle + 1
    Call AjouterCategorie(ws, LigneActuelle, 2, "Salaire Conjoint", "Revenu", RGB(68, 114, 196), 2000, True, False, "Salaire du conjoint")
    LigneActuelle = LigneActuelle + 1
    Call AjouterCategorie(ws, LigneActuelle, 3, "Primes et Bonus", "Revenu", RGB(68, 114, 196), 300, True, False, "Primes, gratifications, bonus")
    LigneActuelle = LigneActuelle + 1
    
    ' Catégories de dépenses
    Call AjouterCategorie(ws, LigneActuelle, 4, "Logement", "Dépense", RGB(196, 89, 17), 1200, True, False, "Loyer, charges, maintenance")
    LigneActuelle = LigneActuelle + 1
    Call AjouterCategorie(ws, LigneActuelle, 5, "Alimentation", "Dépense", RGB(196, 89, 17), 600, True, False, "Courses, restaurants")
    LigneActuelle = LigneActuelle + 1
    Call AjouterCategorie(ws, LigneActuelle, 6, "Transport", "Dépense", RGB(196, 89, 17), 400, True, False, "Essence, transports publics")
    LigneActuelle = LigneActuelle + 1
    Call AjouterCategorie(ws, LigneActuelle, 7, "Loisirs", "Dépense", RGB(196, 89, 17), 300, True, False, "Sorties, vacances, activités")
    LigneActuelle = LigneActuelle + 1
    Call AjouterCategorie(ws, LigneActuelle, 8, "Épargne", "Dépense", RGB(112, 173, 71), 500, True, False, "Épargne mensuelle")
    
End Sub

Sub AjouterCategorie(ws As Worksheet, Ligne As Integer, ID As Long, Nom As String, _
                    TypeCat As String, Couleur As Long, BudgetDefaut As Currency, _
                    EstActive As Boolean, EstPersonnalisee As Boolean, Description As String)
    '-------------------------------------------------------------------------
    ' Ajoute une catégorie dans la feuille des catégories
    '-------------------------------------------------------------------------
    
    With ws
        .Cells(Ligne, 1).Value = ID
        .Cells(Ligne, 2).Value = Nom
        .Cells(Ligne, 3).Value = TypeCat
        .Cells(Ligne, 4).Interior.Color = Couleur
        .Cells(Ligne, 4).Value = "■"
        .Cells(Ligne, 5).Value = BudgetDefaut
        .Cells(Ligne, 6).Value = IIf(EstActive, "OUI", "NON")
        .Cells(Ligne, 7).Value = IIf(EstPersonnalisee, "OUI", "NON")
        .Cells(Ligne, 8).Value = Description
        
        ' Formatage
        .Cells(Ligne, 5).NumberFormat = "#,##0 €"
        .Range(.Cells(Ligne, 1), .Cells(Ligne, 8)).Borders.LineStyle = xlContinuous
        .Range(.Cells(Ligne, 1), .Cells(Ligne, 8)).Font.Size = 9
    End With
    
End Sub

'===============================================================================
' FIN DU MODULE CATEGORIES
'===============================================================================
