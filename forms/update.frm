VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmupdate 
   Caption         =   "update"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   1080
      Top             =   4800
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\accommodation\database\CHC2013.accdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\accommodation\database\CHC2013.accdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdreset 
      Caption         =   "Reset"
      Height          =   495
      Left            =   3000
      TabIndex        =   8
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Yes"
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox blockno 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2880
      Top             =   4920
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\accommodation\database\CHC2013.accdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\accommodation\database\CHC2013.accdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.ComboBox houseuse 
      Height          =   315
      ItemData        =   "update.frx":0000
      Left            =   1560
      List            =   "update.frx":0010
      TabIndex        =   6
      Top             =   3000
      Width           =   3135
   End
   Begin VB.TextBox Housecode 
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   1200
      Width           =   3135
   End
   Begin VB.TextBox Tfileno 
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   600
      Width           =   3135
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Yes"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   2520
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker dateentry 
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   1920
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   873
      _Version        =   393216
      Format          =   117243905
      CurrentDate     =   41673
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "Update"
      Height          =   495
      Left            =   1440
      TabIndex        =   7
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Block Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Rent Payment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   14
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "House Use"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   13
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Allocation Letter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Entry Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "House Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "File Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "frmupdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub blockno_LostFocus()
If blockno.Text = "" Then
Exit Sub
Else
           formno = blockno.Text
                m1 = "[blockno]='" + formno + "'"
                If rs.State = adStateOpen Then rs.Close
                rs.Open "select * from [housecode] where" & m1, db, adOpenDynamic, adLockOptimistic
                If Not rs.EOF Then
                    Housecode.Text = rs![house code]
                    Else
                 MsgBox "Block Number not Found", vbInformation
                                 
                    
                End If
End If

End Sub

Private Sub cmdupdate_Click()
If Tfileno = "" Or Housecode = "" Or houseuse = "" Then
MsgBox "Some important fields are missing"
Exit Sub
Else



                 Dim m1 As String
                formno = Housecode.Text
                m1 = "[housecode]='" + formno + "'"
                If rs.State = adStateOpen Then rs.Close
                rs.Open "select * from AllocationT where" & m1, db, adOpenDynamic, adLockOptimistic
               ' rs.Open "select * from AllocationT where itemcode='" & itemcode.Text & "'", db, adOpenDynamic, adLockOptimistic
               If Not rs.EOF Then
                    rs![Tfileno] = Tfileno.Text
                    rs![ENTRYDATE] = dateentry.Value
                    rs![houseuse] = houseuse.Text
                    If Check1.Value = 1 Then
                     rs![AllocationLetter] = True
                    Else
                    rs![AllocationLetter] = False
                    End If
                    If Check2.Value = 1 Then
                     rs![rentpayment] = True
                    Else
                    rs![rentpayment] = False
                    End If
                  rs.Update
                  Else
                    MsgBox "RECORD NOT FOUND"
                    End If
                     MsgBox "success"
                   
End If
End Sub






