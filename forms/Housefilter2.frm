VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Housefilter 
   Caption         =   "Form2"
   ClientHeight    =   8625
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9840
   LinkTopic       =   "Form2"
   ScaleHeight     =   8625
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   7800
      Visible         =   0   'False
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Print Record"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      TabIndex        =   13
      Top             =   7680
      Width           =   2415
   End
   Begin VB.ComboBox cmbroomno 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   3480
      TabIndex        =   10
      Top             =   3240
      Width           =   3735
   End
   Begin VB.ComboBox cmbflatno 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3480
      TabIndex        =   9
      Top             =   2040
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Populate List"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   4
      Top             =   3960
      Width           =   3615
   End
   Begin VB.ComboBox cmbestate 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3480
      TabIndex        =   3
      Top             =   840
      Width           =   3735
   End
   Begin VB.ComboBox cmbroad 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3480
      TabIndex        =   2
      Top             =   1440
      Width           =   3735
   End
   Begin VB.ComboBox cmbblockno 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   3480
      TabIndex        =   1
      Top             =   2640
      Width           =   3735
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   4920
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   4683
      _Version        =   393216
      BackColor       =   8438015
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   8040
      Width           =   5775
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "SEARCH THE ACCOMODATION RECORD USING DIFFERENT CRITERIALS AS YOU WISH FROM THE OPTIONS BELLOW"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   9495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   4
      FillColor       =   &H00404080&
      Height          =   855
      Index           =   2
      Left            =   3360
      Top             =   3840
      Width           =   4095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Room No: "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   3240
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Estate: "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Road/Line: "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Flat No: "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Block No: "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   2640
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000000&
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   4
      FillColor       =   &H00404080&
      FillStyle       =   0  'Solid
      Height          =   4095
      Index           =   0
      Left            =   3240
      Top             =   720
      Width           =   4335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   4
      FillColor       =   &H00404080&
      Height          =   4095
      Index           =   3
      Left            =   7560
      Top             =   720
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   4
      FillColor       =   &H00404080&
      Height          =   4095
      Index           =   1
      Left            =   0
      Top             =   720
      Width           =   3255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   4
      FillColor       =   &H00404080&
      Height          =   8535
      Index           =   4
      Left            =   0
      Top             =   0
      Width           =   9735
   End
End
Attribute VB_Name = "Housefilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public m1, m2, m3, m4, m5 As String
Public pro As Integer

Private Sub cmbblockno_Click()
If (cmbestate.Text = "") Then
MsgBox ("select an estate")
cmbblockno.Text = ""
cmbestate.SetFocus
Exit Sub
End If
m3a = ""
If (cmbblockno.Text <> "") Then
m3a = " and [blockno]='" + cmbblockno.Text + "'"
'sherif = sherif & " and " & m1
End If
'MsgBox sherif
End Sub

Private Sub cmbestate_Click()
m5a = "select * from HouseEnumeration where [estate] = '" + cmbestate.Text + "'"
End Sub

Private Sub cmdroad_LostFocus()

End Sub

Private Sub cmbflatno_Click()
If (cmbestate.Text = "") Then
MsgBox ("select an estate")
cmbflatno.Text = ""
cmbestate.SetFocus
Exit Sub
End If
m2a = ""
If (cmbflatno.Text <> "") Then
m2a = " and [flatno]='" + cmbflatno.Text + "'"
'sherif = sherif & " and " & m1
End If
'MsgBox sherif
End Sub

Private Sub cmbroad_Click()
If (cmbestate.Text = "") Then
MsgBox ("select an estate")
cmbroad.Text = ""
cmbestate.SetFocus
Exit Sub
End If
m1a = ""
If (cmbroad.Text <> "") Then
m1a = " and [road]='" + cmbroad.Text + "'"
'sherif = sherif & " and " & m1
End If
'MsgBox sherif
End Sub

Private Sub cmbroomno_Click()
If (cmbestate.Text = "") Then
MsgBox ("select an estate")
cmbroomno.Text = ""
cmbestate.SetFocus
Exit Sub
End If
m4a = ""
If (cmbroomno.Text <> "") Then
m4a = " and [Roomno]='" + cmbroomno.Text + "'"
'sherif = sherif & " and " & m1
End If
'MsgBox sherif

End Sub

Private Sub Command1_Click()

If (cmbestate.Text = "") Then
MsgBox ("select an estate")
cmbestate.SetFocus
Exit Sub
End If

'check for empty combo boxes to empty there values

If (cmbroad.Text = "") Then
m1a = ""
End If
If (cmbflatno.Text = "") Then
m2a = ""
End If
If (cmbroomno.Text = "") Then
m4a = ""
End If
If (cmbblockno.Text = "") Then
m3a = ""
End If

'combine the outcome of each combo boxes and save in sherif string

sherif = Trim(m5a & m1a & m2a & m3a & m4a)

'echo out the concatenated result

MsgBox sherif

'Dellete all the record formaly in temporary table to enable save your new search record
If rs4.State = adStateOpen Then rs4.Close
rs4.Open "DELETE * FROM TEMPACCOMODATION ", db, 3, 3
'rs4.Close
'Set rs4 = Nothing
'use the result of the concatenated string to search through the House Enumeration Table to retrieve the House code

If rs.State = adStateOpen Then rs.Close

            rs.Open sherif, db, adOpenKeyset, adLockReadOnly
            If rs.EOF Then
            Exit Sub
            Else
            sherif345 = rs.RecordCount
             ProgressBar1.Visible = True
            Label7.Caption = "Fetching result please wait for some time ..."
            ProgressBar1.Max = sherif345
            rs.MoveFirst
            Do Until rs.EOF
                    bibian = rs![housecode]
                    pro = pro + 1
                    ProgressBar1.Value = pro
                    ' use the housecode return from sherif to loop through the Allocation table to retrieve the TFileNo
                    
                    If rs1.State = adStateOpen Then rs1.Close
                    rs1.Open "SELECT * FROM [AllocationT] WHERE Housecode ='" + rs![housecode] + "'and TFileNo <> NULL", db, 3, 3
                    If Not rs1.EOF Then
                    'Resume Next
                    'Else
                        azeez = rs1!tfileno
                    
                        'search for the records in TenantT table with the Tfileno retrieved from AllocationT table
                    
                        If rs2.State = adStateOpen Then rs2.Close
                        rs2.Open "SELECT * FROM [TenantT] WHERE TFileNo ='" + rs1!tfileno + "'", db, 3, 3
                        If Not rs2.EOF Then
                        'MsgBox "i entered here " & azeez
                        
                        'insert into temporary table all the details you want
                        If rs3.State = adStateOpen Then rs3.Close
                        rs3.Open "SELECT * FROM TEMPACCOMODATION", db, adOpenDynamic, adLockOptimistic
                        rs3.AddNew
                        
                        'save table house enumeration where record march i.e  rs
                        
                        rs3!estate = rs!estate
                        rs3!housetype = rs!housetype
                        rs3!housecode = rs!housecode
                        rs3!ROOMNo = rs!ROOMNo
                        rs3!FlatNo = rs!FlatNo
                        rs3!BlockNo = rs!BlockNo
                        rs3!ROAD = rs!ROAD
                        rs3!housestatus = rs!housestatus
                        rs3!OCCUPIED = rs!OCCUPIED
                        rs3!ANNUALRENT = rs!ANNUALRENT
                        rs3!YEARCOMPLETED = rs!YEARCOMPLETED
'                        rs3![HOUSEREMARK] = rs![HOUSEREMARK]
                        
                         'save table AllocationT where record march i.e  rs1
                         'DONT RESAVE house code is alredy saved at the top in [HOUSE ENUMERATION]
                        
                        rs3!tfileno = rs1!tfileno
                        rs3!entrydate = rs1!entrydate
                        rs3!ExitDate = rs1!ExitDate
                        rs3!RentPayment = rs1!RentPayment
                        rs3!AllocationLetter = rs1!AllocationLetter
                        rs3!Clearance = rs1!Clearance
                        rs3!Comment = rs1!Comment
                        rs3!HouseUse = rs1!HouseUse
                        
                        'save table TenantT where record march i.e  rs2
                         'DONT RESAVE TFileNo is alredy saved at the top in [AllocationT]
                        
                        rs3!SURNAME = rs2!SURNAME
                        rs3![OTHER NAMES] = rs2![OTHER NAMES]
                        rs3![FORM No] = rs2![FORMNo]
                        rs3!ACTIVE = rs2!ACTIVE
                        rs3![COMPANY/ORGZTN] = rs2![COMPANY/ORGZTN]
                        
                        
                        rs3!phoneno = rs2!phoneno
                        rs3!TCATEGORY = rs2!TCATEGORY
                        rs3![WORK PLACE] = rs2![WORK PLACE]
                        rs3!emailaddress = rs2!emailaddress
                        rs3![DATE OF BIRTH] = rs2![DATE OF BIRTH]
                        
                     
                        rs3![MARITAL STATUS] = rs2![MARITAL STATUS]
                        rs3![DATE OF FIRST APPT] = rs2![DATE OF FIRST APPT]
                        rs3!DEPARTMENT = rs2!DEPARTMENT
                        rs3![GRADE LEVEL] = rs2![GRADE LEVEL]
                        rs3!step = rs2!step
                           
                      
                        rs3!DIVISION = rs2!DIVISION
                        rs3!designation = rs2!designation
                        rs3!ippisno = rs2!ippisno
                        rs3!OCCUPATION = rs2!OCCUPATION
                        rs3!LGA = rs2!LGA
                        rs3![STATE OF ORIGIN] = rs2![STATE OF ORIGIN]
                        
                        
                        rs3![HOME TOWN] = rs2![HOME TOWN]
                        rs3!NATIONALITY = rs2!NATIONALITY
                        rs3![TENANT REMARK] = rs2![TENANT REMARK]
                     
                        
                        
                        rs3.Update
                        rs3.Close
                        Set rs3 = Nothing
                        
                        
                        End If
                    End If
            rs.MoveNext
            Loop
           ' MsgBox "Record Saved"
            rs.Close
            Set rs = Nothing
            End If
                    'End If
                'populate the result from temp table to the grid view
                db.CursorLocation = adUseClient
                str1 = "select * from TEMPACCOMODATION "
                If rs5.State = adStateOpen Then rs5.Close
                rs5.Open str1, db, adOpenDynamic, adLockOptimistic
                
                Set DataGrid1.DataSource = rs5
               
                
                ProgressBar1.Value = 0
                ProgressBar1.Visible = False
                Label7.Caption = "Result Ready ... and" & " " & sherif345 & " Record found"
                pro = 0
End Sub

Private Sub Form_Load()
'   On Error Resume Next
                If rs.State = adStateOpen Then rs.Close
rs.Open "select distinct estate from [HouseEnumeration] ", db, adOpenDynamic, adLockOptimistic
                        rs.MoveFirst
                        Do While Not rs.EOF
                        cmbestate.AddItem rs![estate]
                       
                        rs.MoveNext
                        Loop
       If rs1.State = adStateOpen Then rs1.Close
rs1.Open "select distinct road from [HouseEnumeration] where road <> null", db, adOpenDynamic, adLockOptimistic
                        rs1.MoveFirst
                        Do While Not rs1.EOF
                        cmbroad.AddItem rs1![ROAD]
                       
                        rs1.MoveNext
                        Loop

 If rs2.State = adStateOpen Then rs2.Close
rs2.Open "select distinct blockno from [HouseEnumeration] where blockno <> null", db, adOpenDynamic, adLockOptimistic
                        rs2.MoveFirst
                        Do While Not rs2.EOF
                        cmbblockno.AddItem rs2![BlockNo]
                       
                        rs2.MoveNext
                        Loop
                        
If rs3.State = adStateOpen Then rs3.Close
rs3.Open "select distinct flatno from [HouseEnumeration] where flatno <> null", db, adOpenDynamic, adLockOptimistic
                        rs3.MoveFirst
                        Do While Not rs3.EOF
                        cmbflatno.AddItem rs3![FlatNo]
                       
                        rs3.MoveNext
                        Loop
                        
If rs4.State = adStateOpen Then rs4.Close
rs4.Open "select distinct roomno from [HouseEnumeration] where roomno <> null", db, adOpenDynamic, adLockOptimistic
                        rs4.MoveFirst
                        Do While Not rs4.EOF
                        cmbroomno.AddItem rs4![ROOMNo]
                       
                        rs4.MoveNext
                        Loop
                         With Me
        .Top = (Screen.Height - .Height) / 2
        .Left = (Screen.Width - .Width) / 2
    End With
End Sub


