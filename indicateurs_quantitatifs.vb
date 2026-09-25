Sub indicateurs_quanti()
    Dim mode As String
    Dim plage As Range
    Dim cellule As Range
    Dim test As String

mode = WorksheetFunction.mode(Range("R2:R4844"))

'Utilisation de l'objet dictionary
Dim valeursUniques As Object
Set valeursUniques = CreateObject("Scripting.Dictionary")

'Teste chaque cellule pour trouver toutes les valeurs possibles
For Each cellule In Range("S2:S4844")
valeursUniques(cellule.Value) = Empty
Next cellule

'Msgbox des valeurs uniques
For Each valeurUnique In valeursUniques.keys
    MsgBox (valeurUnique)
Next valeurUnique

End Sub
