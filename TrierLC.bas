Attribute VB_Name = "Module2"
Sub TrierEtSeparerCotesLC()
    Dim ws As Worksheet
    Dim i As Long, j As Long
    Dim valeur() As String
    Dim Colonne As String
    Dim Ligne As Long
    Dim ValCol As Long
    Dim DerniereLigne As Long
    Dim StartSeparationCol As Long
    Dim DerniereColonne As Long
    Dim PlageTri As Range
    
    Set ws = ActiveSheet
    
    ' --- Demande à l'utilisateur la colonne et ligne de départ ---
    Colonne = Application.InputBox("Entrer la colonne des cotes LC", Type:=2)
    If Colonne = "" Then Exit Sub
    Ligne = Application.InputBox("Entrer la ligne de la première cote LC", Type:=1)
    If Ligne = 0 Then Exit Sub
    
    ValCol = ws.Range(Colonne & 1).Column
    DerniereLigne = ws.Cells(ws.Rows.Count, ValCol).End(xlUp).Row
    
    ' --- Déterminer la première colonne pour la séparation ---
    StartSeparationCol = ws.Cells(Ligne, ws.Columns.Count).End(xlToLeft).Column + 1
    
    ' --- Effacer les anciennes colonnes de séparation ---
    For i = StartSeparationCol To ws.Columns.Count
        If ws.Cells(Ligne, i).Value <> "" Then
            ws.Columns(i).ClearContents
        Else
            Exit For
        End If
    Next i
    
    ' --- Séparer les valeurs selon les espaces ---
    For i = Ligne To DerniereLigne
        If Not IsEmpty(ws.Cells(i, ValCol)) Then
            ws.Cells(i, ValCol).Value = Replace(ws.Cells(i, ValCol).Value, Chr(160), " ")
            ws.Cells(i, ValCol).Value = Trim(ws.Cells(i, ValCol).Value)
            valeur = Split(ws.Cells(i, ValCol).Value, " ")
            
            ' Répartir dynamiquement dans les colonnes
            For j = 0 To UBound(valeur)
                ws.Cells(i, StartSeparationCol + j).Value = valeur(j)
            Next j
        End If
    Next i
    
    ' --- Définir la plage pour le tri ---
    DerniereColonne = ws.Cells(Ligne, ws.Columns.Count).End(xlToLeft).Column
    Set PlageTri = ws.Range(ws.Cells(Ligne, 1), ws.Cells(DerniereLigne, DerniereColonne))
    
    ' --- Tri par colonnes séparées ---
    With ws.Sort
        .SortFields.Clear
        For j = 0 To DerniereColonne - StartSeparationCol
            .SortFields.Add2 Key:=ws.Cells(Ligne, StartSeparationCol + j), Order:=xlAscending
        Next j
        .SetRange PlageTri
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
    
    MsgBox "Séparation et tri par cote LC effectués avec succès.", vbInformation
End Sub

