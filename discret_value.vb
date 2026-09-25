Sub indicateurs()

Dim nbre_nonvide As Integer
Dim nbre_vide As Double
Dim nbre_0 As Integer
Dim min As String
Dim max As Integer
Dim moyenne As String
Dim mediane As Integer
Dim test As String

nbre_nonvide = WorksheetFunction.CountA(Range("R2:R4844"))
nbre_vide = WorksheetFunction.CountBlank(Range("R2:R4844"))
min = WorksheetFunction.min(Range("R2:R4844"))
max = WorksheetFunction.max(Range("R2:R4844"))
moyenne = WorksheetFunction.Average(Range("R2:R4844"))
mediane = WorksheetFunction.Median(Range("R2:R4844"))

nbre_0 = 0
For i = 2 To 4844

    If Range("R" & i).Value = "0" Then
    nbre_0 = nbre_0 + 1
    End If
Next i

End Sub
