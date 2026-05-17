Attribute VB_Name = "consulta"
Sub consultaSQL()

Dim cn          As New ADODB.Connection
Dim rs          As New ADODB.Recordset
Dim strSQL      As String
Dim header      As Integer
Dim path        As String
Dim arquivo     As String
Dim sht         As String


path = Planilha2.Range("A7").Text                       'caminho da planilha de consulta )base de dados)
arquivo = Planilha2.Range("B7").Text                    'nome do arquivo que está sendo consultado
sht = Planilha2.Range("C7").Text                        'nome da palnilha do arquivo BD

cn.ConnectionString = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" & path & arquivo
cn.Open

strSQL = Planilha2.Range("E7").Text                     'SELECT * FROM [DADOS$] <= nome da aba com $ no final


Set rs = cn.Execute(strSQL)

Planilha1.Rows("1:" & Rows.Count).Delete
Planilha1.[a2].CopyFromRecordset rs

For header = 0 To rs.Fields.Count - 1
    Cells(1, header + 1) = rs.Fields(header).Name
Next

rs.Close
Set rs = Nothing

cn.Close
Set cn = Nothing


End Sub
