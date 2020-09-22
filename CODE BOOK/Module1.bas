Attribute VB_Name = "Module1"
Global SR As Integer, SRPLUS As Integer, NR As Integer, NRPLUS As Integer
Global BillStat As String
Global CN            As ADODB.Connection
Public RSc           As ADODB.Recordset
Public RS            As ADODB.Recordset
Global RSJO          As ADODB.Recordset
Public Sub connectDB()
On Error GoTo ErrHandler
    Dim strCN As String, DBpass As String
    Set CN = New ADODB.Connection
    strCN = "Data Source=" & App.Path & "\USdb.mdb"
    CN.Provider = "Microsoft Jet 4.0 OLE DB Provider"
    CN.ConnectionString = strCN
    CN.Open
   Exit Sub
ErrHandler:
MsgBox "Database does not exist or"
CN.Close
End
End Sub


