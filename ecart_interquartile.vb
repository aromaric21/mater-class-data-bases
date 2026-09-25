Sub ecart_interquartile()

Dim Q1 As Double
Dim Q3 As Double
Dim IQR As Double
Dim borne_basse As Double
Dim borne_haute As Double

Q1 = WorksheetFunction.Quartile(Range("E2:E4844"), 1)
Q3 = WorksheetFunction.Quartile(Range("E2:E4844"), 3)
IQR = Q3 - Q1

borne_basse = Round(Q1 - (1.5 * IQR), 0)
borne_haute = Round(Q3 + (1.5 * IQR), 0)

If borne_basse < 0 Then
   borne_basse = 0
   MsgBox ("attention borne basse inférieure à 0")
   Range("$B$1:$V$5000").AutoFilter Field:=4, Criteria1:=">" & borne_haute, _
     Operator:=xlAnd
   test = ">" & borne_haute

End If
End Sub
