Sub test_pearson()

Dim test_pearson As Single

test_pearson = WorksheetFunction.Pearson(Range("A4:A12"), Range("B4:B12"))
MsgBox ("le coefficient de Pearson est de : " & test_pearson)

End Sub
