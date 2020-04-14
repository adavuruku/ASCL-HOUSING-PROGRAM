VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   10845
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2415
      Left            =   240
      TabIndex        =   8
      Top             =   5280
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   4260
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16776960
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Reference Sans Serif"
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
   Begin VB.TextBox Text7 
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   19
      Top             =   3600
      Width           =   3255
   End
   Begin VB.TextBox Text6 
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   17
      Top             =   3120
      Width           =   3255
   End
   Begin VB.TextBox Text5 
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   15
      Top             =   2640
      Width           =   3255
   End
   Begin VB.CommandButton cmdsearch 
      Caption         =   "Search Records"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   7
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CommandButton cmdprevious 
      Caption         =   "<<Previous"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   6
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton cmdmovefirst 
      Caption         =   "Move First"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   5
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton next 
      Caption         =   "Move Next >>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   1680
      Width           =   3255
   End
   Begin VB.TextBox Text3 
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   2160
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2400
      TabIndex        =   1
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "List Of All Houses Occupied By The One With The Above Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   21
      Top             =   4920
      Width           =   8175
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Form No :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   20
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2895
      Left            =   7080
      Stretch         =   -1  'True
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "House Code :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   16
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Estate :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   14
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "House Type :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   13
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone_No :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   12
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "File_No :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   2175
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      BorderWidth     =   4
      FillColor       =   &H00FF80FF&
      Height          =   4215
      Left            =   120
      Top             =   600
      Width           =   2055
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      BorderWidth     =   4
      FillColor       =   &H00FF80FF&
      Height          =   495
      Left            =   120
      Top             =   120
      Width           =   10455
   End
   Begin VB.Label Label8 
      BackColor       =   &H00404040&
      Height          =   3135
      Left            =   6720
      TabIndex        =   18
      Top             =   720
      Width           =   3375
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      BorderWidth     =   4
      FillColor       =   &H00FF80FF&
      Height          =   4695
      Left            =   120
      Top             =   120
      Width           =   10455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      BorderWidth     =   4
      FillColor       =   &H00FF80FF&
      Height          =   7695
      Left            =   120
      Top             =   120
      Width           =   10455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoconn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset

Private Sub cmdmovefirst_Click()
rs.MoveFirst
Text1.Text = rs("tfileno")
    Text2.Text = rs("surname") & " " & rs("other names")
    Text3.Text = rs("estate")
    Text4.Text = rs("HOUSECODE")
    Text5.Text = rs("phoneno")
     Text6.Text = rs("housetype")
     Text7.Text = rs("FORMNo")
 Set Image1.Picture = LoadPicture(App.Path & "\photos\" & rs("tfileno") & ".jpg")
 
    
    adoconn.CursorLocation = adUseClient
    ' bring to paste
str1 = "SELECT [HouseEnumeration].[estate],[HouseEnumeration].[HOUSECODE],[HouseEnumeration].[FlatNo],[HouseEnumeration].[ROOMNo],[HouseEnumeration].[ROAD],[AllocationT].[EntryDate],[AllocationT].[RentPayment] FROM [HouseEnumeration], [AllocationT] where [HouseEnumeration].[housecode]=[AllocationT].[housecode] AND [AllocationT].[TfileNo]='" & Text1 & "'"
    If rs1.State = adStateOpen Then rs1.Close
    rs1.Open str1, adoconn, adOpenDynamic, adLockOptimistic
    
    Set DataGrid1.DataSource = rs1
 
End Sub

Private Sub cmdprevious_Click()
    rs.MovePrevious
    If rs.BOF = True Then
        MsgBox "This is the first record.", vbExclamation, "Note it..."
        rs.MoveFirst
    End If
 Text1.Text = rs("tfileno")
    Text2.Text = rs("surname") & " " & rs("other names")
    Text3.Text = rs("estate")
    Text4.Text = rs("HOUSECODE")
    Text5.Text = rs("phoneno")
     Text6.Text = rs("housetype")
     Text7.Text = rs("FORMNo")
 Set Image1.Picture = LoadPicture(App.Path & "\photos\" & rs("tfileno") & ".jpg")
 
    
    adoconn.CursorLocation = adUseClient
    ' bring to paste
str1 = "SELECT [HouseEnumeration].[ESTATE],[HouseEnumeration].[HOUSECODE],[HouseEnumeration].[FlatNo],[HouseEnumeration].[ROOMNo],[HouseEnumeration].[ROAD],[AllocationT].[EntryDate],[AllocationT].[RentPayment] FROM [HouseEnumeration], [AllocationT] where [HouseEnumeration].[housecode]=[AllocationT].[housecode] AND [AllocationT].[TfileNo]='" & Text1 & "'"
    If rs1.State = adStateOpen Then rs1.Close
    rs1.Open str1, adoconn, adOpenDynamic, adLockOptimistic
    
    Set DataGrid1.DataSource = rs1
End Sub

Private Sub cmdsearch_Click()
 Dim key As String, str, str1 As String
    key = InputBox("Enter the fILEnO No whose details u want to know: ", "Please Enter The File Number To Search :")
    Set rs = Nothing
    'str = "SELECT [HouseEnumeration].[estate],[HouseEnumeration].[HOUSECODE],[tenantT].[surname],[allocationT].[Tfileno] FROM [House Enumeration], [AllocationT], [TenantT] where [House Enumeration].[house code]=[AllocationT].[housecode] AND [AllocationT].[TfileNo]=[TenantT].[TfileNo] AND [allocationT].[TfileNo]='" & key & "'"
     str = "SELECT [HouseEnumeration].[estate],[HouseEnumeration].[HOUSECODE],[HouseEnumeration].[Housetype],[tenantT].[surname],[tenantT].[other names],[tenantT].[phoneno],[tenantT].[FORMNo],[allocationT].[Tfileno] FROM [HouseEnumeration], [AllocationT], [TenantT] where [HouseEnumeration].[housecode]=[AllocationT].[housecode] AND [AllocationT].[TfileNo]=[TenantT].[TfileNo] AND [allocationT].[TfileNo]='" & key & "'"
'   rs.Open str, adoconn, adOpenForwardOnly, adLockReadOnly
    rs.Open str, adoconn, adOpenDynamic, adLockOptimistic
If rs.EOF = True Then
Exit Sub
Else
    Text1.Text = rs("tfileno")
    Text2.Text = rs("surname") & " " & rs("other names")
    Text3.Text = rs("estate")
    Text4.Text = rs("HOUSECODE")
    Text5.Text = rs("phoneno")
     Text6.Text = rs("housetype")
     Text7.Text = rs("FORMNo")
 Set Image1.Picture = LoadPicture(App.Path & "\photos\" & rs("tfileno") & ".jpg")
 
    
    adoconn.CursorLocation = adUseClient
    ' bring to paste
str1 = "SELECT [HouseEnumeration].[estate],[HouseEnumeration].[HOUSECODE],[HouseEnumeration].[FlatNo],[HouseEnumeration].[ROOMNo],[HouseEnumeration].[ROAD],[AllocationT].[EntryDate],[AllocationT].[RentPayment] FROM [HouseEnumeration], [AllocationT] where [HouseEnumeration].[housecode]=[AllocationT].[housecode] AND [AllocationT].[TfileNo]='" & Text1 & "'"
    If rs1.State = adStateOpen Then rs1.Close
    rs1.Open str1, adoconn, adOpenDynamic, adLockOptimistic
    
    Set DataGrid1.DataSource = rs1
  End If
End Sub

Private Sub Form_Load()

    Dim str As String
    
    Set adoconn = Nothing
   ' adoconn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\look.accdb;Persist Security Info=False"
    adoconn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & App.Path & "\look.accdb;Persist Security Info=False"
 '   str = "SELECT * FROM ([House Enumeration] LEFT JOIN AllocationT ON [House Enumeration].[HOUSE CODE] = AllocationT.HouseCode) RIGHT JOIN TenantT ON AllocationT.TFileNo = TenantT.[OTHER NAMES]"
    
    str = "SELECT [HouseEnumeration].[estate],[HouseEnumeration].[Housetype],[HouseEnumeration].[HOUSECODE],[tenantT].[surname],[tenantT].[other names],[tenantT].[phoneno],[tenantT].[FORMNo],[allocationT].[Tfileno] FROM [HouseEnumeration], [AllocationT], [TenantT] where [HouseEnumeration].[housecode]=[AllocationT].[housecode] AND [AllocationT].[TfileNo]=[TenantT].[TfileNo]"
   If rs.State = adStateOpen Then rs.Close
    rs.Open str, adoconn, adOpenDynamic, adLockOptimistic
    rs.MoveFirst
    If rs.EOF = True Then
    
    MsgBox ("    kkkk ")
    Else
    
    
    Text1.Text = rs("tfileno")
    Text2.Text = rs("surname") & "  " & rs("other names")
    Text3.Text = rs("estate")
    Text4.Text = rs("HOUSECODE")
    Text5.Text = rs("phoneno")
     Text6.Text = rs("housetype")
     Text7.Text = rs("FORMNo")
 Set Image1.Picture = LoadPicture(App.Path & "\photos\" & rs("tfileno") & ".jpg")
      adoconn.CursorLocation = adUseClient
    ' bring to paste
str1 = "SELECT [HouseEnumeration].[ESTATE],[HouseEnumeration].[HOUSECODE],[HouseEnumeration].[FlatNo],[HouseEnumeration].[ROOMNo],[HouseEnumeration].[ROAD],[AllocationT].[EntryDate],[AllocationT].[RentPayment] FROM [HouseEnumeration], [AllocationT] where [HouseEnumeration].[housecode]=[AllocationT].[housecode] AND [AllocationT].[TfileNo]='" & Text1 & "'"
    If rs1.State = adStateOpen Then rs1.Close
    rs1.Open str1, adoconn, adOpenDynamic, adLockOptimistic
    
    Set DataGrid1.DataSource = rs1
    
   End If
   ' txtPhone.Text = rs(3)
    'txtdepartment.Text = rs("department")
    'myphoto = rs(6)
'   Set pic2.Picture = LoadPicture(App.Path & "\photos\" & myphoto & ".jpg")
    'Set pic2.Picture = LoadPicture(App.Path & "\photos\" & rs("imagepath") & ".jpg")
'dd = "SELECT * FROM ([House Enumeration] LEFT JOIN AllocationT ON [House Enumeration].[HOUSE CODE] = AllocationT.HouseCode) RIGHT JOIN TenantT ON AllocationT.TFileNo = TenantT.[OTHER NAMES]"




End Sub

Private Sub next_Click()
    rs.MoveNext
    If rs.EOF = True Then
        MsgBox "This is the last record.", vbExclamation, "Note it..."
        rs.MoveLast
    End If
 Text1.Text = rs("tfileno")
    Text2.Text = rs("surname") & " " & rs("other names")
    Text3.Text = rs("estate")
    Text4.Text = rs("HOUSECODE")
    Text5.Text = rs("phoneno")
     Text6.Text = rs("housetype")
     Text7.Text = rs("FORMNo")
 Set Image1.Picture = LoadPicture(App.Path & "\photos\" & rs("tfileno") & ".jpg")
 
    
    adoconn.CursorLocation = adUseClient
    ' bring to paste
str1 = "SELECT [HouseEnumeration].[estate],[HouseEnumeration].[HOUSECODE],[HouseEnumeration].[FlatNo],[HouseEnumeration].[ROOMNo],[HouseEnumeration].[ROAD],[AllocationT].[EntryDate],[AllocationT].[RentPayment] FROM [HouseEnumeration], [AllocationT] where [HouseEnumeration].[housecode]=[AllocationT].[housecode] AND [AllocationT].[TfileNo]='" & Text1 & "'"
    If rs1.State = adStateOpen Then rs1.Close
    rs1.Open str1, adoconn, adOpenDynamic, adLockOptimistic
    
    Set DataGrid1.DataSource = rs1
End Sub

