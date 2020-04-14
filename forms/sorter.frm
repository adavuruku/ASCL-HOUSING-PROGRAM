VERSION 5.00
Begin VB.Form sorter 
   Caption         =   "Form1"
   ClientHeight    =   5280
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   Picture         =   "sorter.frx":0000
   ScaleHeight     =   5280
   ScaleWidth      =   9120
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo3 
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
      ItemData        =   "sorter.frx":21DAA
      Left            =   3480
      List            =   "sorter.frx":21DB7
      TabIndex        =   10
      Top             =   3480
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   2880
      Width           =   2535
   End
   Begin VB.ComboBox Combo2 
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
      Top             =   2400
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
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
      TabIndex        =   1
      Top             =   1920
      Width           =   2535
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
      Height          =   975
      Left            =   3480
      TabIndex        =   0
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000000&
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   4
      FillColor       =   &H00404080&
      Height          =   2055
      Index           =   1
      Left            =   600
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000000&
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   4
      FillColor       =   &H00404080&
      FillStyle       =   0  'Solid
      Height          =   3255
      Index           =   0
      Left            =   3360
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Staff Status: "
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
      Left            =   600
      TabIndex        =   9
      Top             =   3600
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "SORT OCCUPANTS"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   2760
      TabIndex        =   8
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label Label12 
      Caption         =   " UNIT REGISTRY SYSTEM "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Index           =   1
      Left            =   960
      TabIndex        =   7
      Top             =   0
      Width           =   7215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
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
      Height          =   615
      Left            =   600
      TabIndex        =   6
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
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
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   1920
      Width           =   2775
   End
End
Attribute VB_Name = "sorter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call code009
End Sub
Private Sub Form_Load()
     'Dim db As New ADODB.Connection
      '  Dim rs As New ADODB.Recordset
      '  db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
                    App.Path & "\Database\house.mdb;Persist Security Info=False"
            On Error Resume Next
                If rs.State = adStateOpen Then rs.Close
                rs.Open "select * from [lists]", db, adOpenDynamic, adLockOptimistic
                        rs.MoveFirst
                        Do While Not rs.EOF
                        Combo1.AddItem rs![estate]
                        rs.MoveNext
                        Loop
    With Me
        .Top = (Screen.Height - .Height) / 2
        .Left = (Screen.Width - .Width) / 2
    End With
End Sub
