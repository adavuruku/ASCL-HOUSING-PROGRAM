Attribute VB_Name = "Module1"
Global counter As Integer
Public db As New ADODB.Connection
Public d1 As New ADODB.Connection
Public db1 As New ADODB.Connection
Public rs As New ADODB.Recordset
Public rs1 As New ADODB.Recordset
Public rs2 As New ADODB.Recordset
Public rs3 As New ADODB.Recordset
Public sql As String
Public sql1 As String
Public sql2 As String
Public Sub main()
    On Error Resume Next

   ' db.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\inventory.mdb;Persist Security Info=False"
     db.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database\house.mdb;Persist Security Info=False"
     db.Open
    'asclstaffReg.Show
    Home.Show
   ' PTUncompleted.Show
End Sub
Public Sub code008()
   ' Dim rs As New ADODB.Recordset
   ' Dim db As New ADODB.Connection
    ' db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
                    App.Path & "\Database\house.mdb;Persist Security Info=False"
    If rs.State = adStateOpen Then rs.Close
        rs.Open "select * from occup_reg ORDER by estate ASC", db, adOpenDynamic, adLockOptimistic
        rs.Requery
        If Not rs.EOF Then
        With DataReport1.Sections("section1").Controls
            .Item("text1").DataField = rs("occp_name").Name
            .Item("text2").DataField = rs("tenant").Name
            .Item("text3").DataField = rs("estate").Name
            .Item("text4").DataField = rs("road_line").Name
            .Item("text5").DataField = rs("flat").Name
        End With
        Set DataReport1.DataSource = rs
        DataReport1.Show
        Else
        MsgBox "NO RECORD EXISTS", vbInformation, "Error Message: "
    End If
End Sub
Public Sub code009()
   ' Dim rs As New ADODB.Recordset
   ' Dim db As New ADODB.Connection
    '  db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
                    App.Path & "\Database\house.mdb;Persist Security Info=False"
    If rs.State = adStateOpen Then rs.Close
    Dim m1, m2, m3, m4, m5 As String
        estate = sorter.Combo1.Text
        road = sorter.Combo2.Text
        flat = sorter.Text1.Text
        staff = sorter.Combo3.Text
            m1 = "[estate]='" + estate + "'"
            m2 = "[road_line]='" + road + "'"
            m3 = "[flat]='" + flat + "'"
            m4 = "[tenant]='" + staff + "'"
                    If (estate = "" Or estate = "Select Estate") Then
                    MsgBox "SELECT AN ESTATE", vbInformation, "SELECT PARAMETER"
                    Exit Sub
                    End If
                            If (estate = "" Or estate = "Select Estate") Then
                                        MsgBox "SELECT AN ESTATE", vbInformation, "SELECT PARAMETER"
                                        Exit Sub
                            ElseIf (road = "" Or road = "Select Road/Line") And _
                                                (flat = "" Or flat = "Select Flat") And _
                                                                                    (staff = "") Then
                                                                        m5 = m1
                            ElseIf (road = "" Or road = "Select Road/Line") And _
                                                (flat = "" Or flat = "Select Flat") And _
                                                                                    (staff <> "") Then
                                                                        m5 = m1 & " AND " & m4
                            ElseIf (road = "" Or road = "Select Road/Line") And _
                                                (flat <> "" Or flat <> "Select Flat") And _
                                                                                    (staff = "") Then
                                                                        m5 = m1 & " AND " & m3
                            ElseIf (road <> "" Or road <> "Select Road/Line") And _
                                                (flat = "" Or flat = "Select Flat") And _
                                                                                    (staff = "") Then
                                                                        m5 = m1 & " AND " & m2
                            ElseIf (road = "" Or road = "Select Road/Line") And _
                                                (flat <> "" Or flat <> "Select Flat") And _
                                                                                    (staff <> "") Then
                                                                        m5 = m1 & " AND " & m4 & " AND " & m3
                            ElseIf (road <> "" Or road <> "Select Road/Line") And _
                                                (flat = "" Or flat = "Select Flat") And _
                                                                                    (staff <> "") Then
                                                                        m5 = m1 & " AND " & m4 & " AND " & m2
                            
                            ElseIf (road = "" Or road = "Select Road/Line") And _
                                                (flat <> "" Or flat <> "Select Flat") And _
                                                                                    (staff <> "") Then
                                                                        m5 = m1 & " AND " & m3 & " AND " & m4
                            ElseIf (road <> "" Or road <> "Select Road/Line") And _
                                                (flat <> "" Or flat <> "Select Flat") And _
                                                                                    (staff = "") Then
                                                                        m5 = m1 & " AND " & m3 & " AND " & m2
                            
                            ElseIf (road <> "" Or road <> "Select Road/Line") And _
                                                (flat = "" Or flat = "Select Flat") And _
                                                                                    (staff <> "") Then
                                                                        m5 = m1 & " AND " & m2 & " AND " & m4
                            ElseIf (road <> "" Or road <> "Select Road/Line") And _
                                                (flat <> "" Or flat <> "Select Flat") And _
                                                                                    (staff = "") Then
                                                                        m5 = m1 & " AND " & m2 & " AND " & m3
                            Else
                                m5 = m1 & " AND " & m2 & " AND " & m3 & " AND " & m4
                            End If
        rs.Open "select * from occup_reg where " & m5 & "ORDER by estate ASC", _
                                db, adOpenDynamic, adLockOptimistic
        rs.Requery
        If Not rs.EOF Then
                With DataReport1.Sections("section1").Controls
                    .Item("text1").DataField = rs("occp_name").Name
                    .Item("text2").DataField = rs("tenant").Name
                    .Item("text3").DataField = rs("estate").Name
                    .Item("text4").DataField = rs("road_line").Name
                    .Item("text5").DataField = rs("flat").Name
                End With
                Set DataReport1.DataSource = rs
                DataReport1.Show
            Else
            MsgBox "No Such Records Exists!", vbInformation, "ALERT"
        End If
End Sub
Public Sub report1()
On Error Resume Next
cPath = App.Path & "\Inventory.mdb"
Dim AccessApp As Access.Application
    
  Set appAccess = New Access.Application
Set appAccess = CreateObject("Access.Application")


  appAccess.OpenCurrentDatabase cPath
    appAccess.DoCmd.OpenReport "masterdel", acViewPreview
appAccess.Visible = True


End Sub
