Sub EtapeCreationFeuilles()
    '-------------------------------------------------------------------------
    ' Étape 1 : Création de toutes les feuilles nécessaires
    '-------------------------------------------------------------------------
    
    Dim ws As Worksheet
    Dim FeuillesRequises As Variant
    Dim i As Integer
    
    ' Liste des feuilles à créer
    FeuillesRequises = Split("Dashboard,Saisie_Mensuelle,Donnees_Revenus,Donnees_Depenses,Categories,Parametres,Rapports,Archives", ",")
    
    ' Supprimer les feuilles par défaut (comme Feuil1, Feuil2, etc.)
    Application.DisplayAlerts = False
    On Error Resume Next
    
    For Each ws In ThisWorkbook.Worksheets
        ' Supprimer seulement les feuilles avec des noms par défaut
        If ws.Name Like "Feuil*" Or ws.Name Like "Sheet*" Or ws.Name Like "Classeur*" Then
            ws.Delete
        End If
    Next ws
    
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Créer les feuilles nécessaires si elles n'existent pas
    For i = LBound(FeuillesRequises) To UBound(FeuillesRequises)
        If Not FeuilleExiste(Trim(FeuillesRequises(i))) Then
            Set ws = ThisWorkbook.Worksheets.Add
            ws.Name = Trim(FeuillesRequises(i))
        End If
    Next i
    
End Sub

Function FeuilleExiste(NomFeuille As String) As Boolean
    '-------------------------------------------------------------------------
    ' Vérifie si une feuille existe dans le classeur
    '-------------------------------------------------------------------------
    
    Dim ws As Worksheet
    
    FeuilleExiste = False
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(NomFeuille)
    If Not ws Is Nothing Then
        FeuilleExiste = True
    End If
    On Error GoTo 0
    
    Set ws = Nothing
    
End Function
