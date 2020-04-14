Attribute VB_Name = "Module1"
Public db As New ADODB.Connection
Public d1 As New ADODB.Connection
Public db1 As New ADODB.Connection
Public rs As New ADODB.Recordset
Public rs1 As New ADODB.Recordset
Public rs2 As New ADODB.Recordset
Public rs3 As New ADODB.Recordset
Public rs4 As New ADODB.Recordset
Public rs5 As New ADODB.Recordset
Public sql As String
Public sherif As String
Public sql1 As String
Public sql2 As String
'dddddd
Public bibian As String
Public azeez As String
Public m1a As String
Public m2a As String
Public m3a As String
Public m4a As String
Public m5a As String
Public m6 As String
Public Sub main()
    On Error Resume Next
    'db.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\inventory.mdb;Persist Security Info=False"
    'db.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\accommodation\database\look.mdb;Persist Security Info=False"
    db.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & App.Path & "\look.accdb;Persist Security Info=False"
menu.Show
'Housefilter.Show

db.Open
End Sub
