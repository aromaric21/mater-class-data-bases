Sub test_doublons()

Dim rng As Range
Dim cell As Range
Dim dict As Object
Dim doublons As Range

Set rng = Range("A2:A4844")
Set dict = CreateObject("Scripting.Dictionary")

'Teste les doublons pour chaque cellule dans la range
For Each cell In rng
    If dict.Exists(cell.Value) Then
        If doublons Is Nothing Then
            Set doublons = cell
        Else
            Set doublons = Union(doublons, cell)
        End If
    Else
        dict.Add cell.Value, 1
    End If
Next cell

'Msgbox pour l'utilisateur
If Not doublons Is Nothing Then
    'Sélectionne les cellules en doublon dans la range
    doublons.Select
    MsgBox "Doublons trouvés dans la plage " & rng.Address & "."
Else
    MsgBox "Aucun doublon trouvé dans la plage " & rng.Address & "."
End If

End Sub
