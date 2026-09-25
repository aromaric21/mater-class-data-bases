Sub valeurs_manquantes_departement()

    Dim ligne_max As String
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    ligne_max = ActiveCell.Row

'Teste toutes les cellules
For i = 1 To 4844
    If Range("L" & i) = "" Then
       Range("L" & i).Select
       'Écriture avec la méthode en VBA
        ActiveCell.FormulaR1C1 = Left(Range("J" & i).Value, 2)
       'Écriture avec la fonction Excel
        ActiveCell.FormulaR1C1 = "=LEFT(RC[-2],2)"
    End If
   Next i
   
End Sub
