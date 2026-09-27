Sub extract_table()

Dim sTableName As String
Dim oWbk As Workbook
Dim sWbkName As String

'Ouvre la table sélectionnée dans un nouveau Workbook
sTableName = "Adresse"
Set oWbk = Workbooks.OpenDatabase(Filename:="I:\BDD Access\BDD - Reporting.accdb", CommandText:=Array(sTableName), CommandType:=xlCmdTable)

'Enregistre et ferme le fichier
sWbkName = "I:\BDD Access\test.xlsx"
oWbk.SaveAs sWbkName
oWbk.Close

End Sub
