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

    ' Définir la feuille active
    Set ws = ActiveSheet

    ' Saisie de l'utilisateur
    Colonne = Application.InputBox("Enter the column of the call letters", Type:=2)
    If Colonne = "" Then Exit Sub
    Ligne = Application.InputBox("Enter the first row of the call letters", Type:=1)
    If Ligne = 0 Then Exit Sub

    ' Détecter la dernière ligne utilisée dans la colonne des cotes LC
    ValCol = ws.Range(Colonne & 1).Column
    DerniereLigne = ws.Cells(ws.Rows.Count, ValCol).End(xlUp).Row

    ' Déterminer où commencent les colonnes pour la séparation
    StartSeparationCol = ws.Cells(Ligne, ws.Columns.Count).End(xlToLeft).Column + 1

    ' Effacement sécurisé des anciennes colonnes de séparation uniquement
    For i = StartSeparationCol To ws.Columns.Count
        If ws.Cells(Ligne, i).Value <> "" Then
            ws.Columns(i).ClearContents
        Else
            Exit For ' Arrêter dès qu'une colonne vide est trouvée
        End If
    Next i

    ' Séparation des valeurs dans les colonnes à droite
    For i = Ligne To DerniereLigne
        If Not IsEmpty(ws.Cells(i, ValCol)) Then
            ws.Cells(i, ValCol).Value = Trim(ws.Cells(i, ValCol).Value)
            valeur = Split(ws.Cells(i, ValCol).Value)
            
            ' Répartir les valeurs dans les colonnes de séparation
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

    ' Définir la plage pour le tri
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

    MsgBox "Sort done.", vbInformation
End Sub
