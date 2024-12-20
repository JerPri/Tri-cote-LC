Attribute VB_Name = "Module1"
Sub TrierEtSeparerCotesLC()
    Dim ws As Worksheet
    Dim i As Long
    Dim valeur() As String
    Dim Colonne As String
    Dim Ligne As Long
    Dim ValCol As Long
    Dim DerniereLigne As Long
    Dim DerniereColonne As Long
    Dim PlageTri As Range
    Dim StartSeparationCol As Long

    ' D�finir la feuille active
    Set ws = ActiveSheet

    ' Saisie de l'utilisateur
    Colonne = Application.InputBox("Entrer la colonne des cotes LC", Type:=2)
    If Colonne = "" Then Exit Sub
    Ligne = Application.InputBox("Entrer la ligne de la premi�re cote LC", Type:=1)
    If Ligne = 0 Then Exit Sub

    ' D�tecter la derni�re ligne utilis�e dans la colonne des cotes LC
    ValCol = ws.Range(Colonne & 1).Column
    DerniereLigne = ws.Cells(ws.Rows.Count, ValCol).End(xlUp).Row

    ' D�terminer o� commencent les colonnes pour la s�paration
    StartSeparationCol = ws.Cells(Ligne, ws.Columns.Count).End(xlToLeft).Column + 1

    ' Effacement s�curis� des anciennes colonnes de s�paration uniquement
    For i = StartSeparationCol To ws.Columns.Count
        If ws.Cells(Ligne, i).Value <> "" Then
            ws.Columns(i).ClearContents
        Else
            Exit For ' Arr�ter d�s qu'une colonne vide est trouv�e
        End If
    Next i

    ' S�paration des valeurs dans les colonnes � droite
    For i = Ligne To DerniereLigne
        If Not IsEmpty(ws.Cells(i, ValCol)) Then
            ws.Cells(i, ValCol).Value = Trim(ws.Cells(i, ValCol).Value)
            valeur = Split(ws.Cells(i, ValCol).Value)
            
            ' R�partir les valeurs dans les colonnes de s�paration
            If UBound(valeur) >= 0 Then ws.Cells(i, StartSeparationCol).Value = valeur(0)
            If UBound(valeur) >= 1 Then ws.Cells(i, StartSeparationCol + 1).Value = valeur(1)
            If UBound(valeur) >= 2 Then ws.Cells(i, StartSeparationCol + 2).Value = valeur(2)
            If UBound(valeur) >= 3 Then
                If Left(valeur(3), 1) = "1" Or Left(valeur(3), 1) = "2" Then
                    ws.Cells(i, StartSeparationCol + 4).Value = valeur(3)
                Else
                    ws.Cells(i, StartSeparationCol + 3).Value = valeur(3)
                End If
            End If
            If UBound(valeur) >= 4 Then ws.Cells(i, StartSeparationCol + 4).Value = valeur(4)
        End If
    Next i

    ' D�finir la plage pour le tri
    DerniereColonne = ws.Cells(Ligne, ws.Columns.Count).End(xlToLeft).Column
    Set PlageTri = ws.Range(ws.Cells(Ligne, 1), ws.Cells(DerniereLigne, DerniereColonne))

    ' Tri dynamique multi-colonnes
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=ws.Cells(Ligne, StartSeparationCol), Order:=xlAscending
        .SortFields.Add2 Key:=ws.Cells(Ligne, StartSeparationCol + 1), Order:=xlAscending
        .SortFields.Add2 Key:=ws.Cells(Ligne, StartSeparationCol + 2), Order:=xlAscending
        .SortFields.Add2 Key:=ws.Cells(Ligne, StartSeparationCol + 3), Order:=xlAscending
        .SortFields.Add2 Key:=ws.Cells(Ligne, StartSeparationCol + 4), Order:=xlAscending

        .SetRange PlageTri
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

    MsgBox "S�paration et tri par cote LC effectu�s avec succ�s.", vbInformation
End Sub

