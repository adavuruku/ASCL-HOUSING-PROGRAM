VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form PrivateTenantReg 
   BackColor       =   &H00404080&
   Caption         =   "Form1"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12795
   FillColor       =   &H00C0FFFF&
   ForeColor       =   &H8000000A&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   12795
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   240
      Top             =   8040
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
      Connect         =   ""
      OLEDBString     =   ""
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
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10560
      TabIndex        =   37
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9480
      TabIndex        =   36
      Top             =   7680
      Width           =   3015
   End
   Begin VB.ComboBox Combo9 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   9480
      TabIndex        =   35
      Text            =   "Select Purpose"
      Top             =   4800
      Width           =   3015
   End
   Begin VB.ComboBox Combo8 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   9480
      TabIndex        =   34
      Text            =   "Select House Type"
      Top             =   4320
      Width           =   3015
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9480
      TabIndex        =   33
      Top             =   6240
      Width           =   3015
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   32
      Top             =   7680
      Width           =   1455
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   3240
      TabIndex        =   30
      Top             =   5520
      Width           =   3975
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   29
      Top             =   3240
      Width           =   3975
   End
   Begin VB.ComboBox Combo7 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   9480
      TabIndex        =   28
      Text            =   "Select House Status"
      Top             =   6720
      Width           =   3015
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   9480
      TabIndex        =   27
      Top             =   5280
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   80936961
      CurrentDate     =   41638
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3240
      TabIndex        =   22
      Top             =   2760
      Width           =   3975
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   3240
      TabIndex        =   21
      Text            =   "Select State"
      Top             =   2280
      Width           =   3975
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   20
      Top             =   1320
      Width           =   3975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H000040C0&
      Height          =   2895
      Left            =   7440
      TabIndex        =   19
      Top             =   840
      Width           =   2895
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   2610
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2685
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   11280
      Top             =   9480
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
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
      Connect         =   ""
      OLEDBString     =   ""
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
   Begin VB.CommandButton Command4 
      Caption         =   "Browse"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   18
      Top             =   3840
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   10560
      TabIndex        =   17
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Reset"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10560
      TabIndex        =   16
      Top             =   2040
      Width           =   1815
   End
   Begin VB.ComboBox Combo6 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Left            =   3240
      TabIndex        =   14
      Text            =   "Select Estate"
      Top             =   7080
      Width           =   3975
   End
   Begin VB.ComboBox Combo5 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Left            =   3240
      TabIndex        =   13
      Text            =   "Select Workplace"
      Top             =   3720
      Width           =   3975
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Left            =   9480
      TabIndex        =   12
      Text            =   "Select Department"
      Top             =   7200
      Width           =   3015
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Left            =   9480
      TabIndex        =   11
      Text            =   "Select Tenant"
      Top             =   5760
      Width           =   3015
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Left            =   3240
      TabIndex        =   10
      Top             =   7680
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1005
      Left            =   3240
      TabIndex        =   9
      Top             =   4320
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   420
      Left            =   3240
      TabIndex        =   8
      Top             =   1800
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Left            =   3240
      TabIndex        =   7
      Top             =   840
      Width           =   3975
   End
   Begin VB.PictureBox cDlImage 
      Height          =   480
      Left            =   13200
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   47
      Top             =   9960
      Width           =   1200
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Flat: "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Index           =   2
      Left            =   4920
      TabIndex        =   46
      Top             =   7680
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "House Type: "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   7440
      TabIndex        =   45
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Department: "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   7440
      TabIndex        =   44
      Top             =   7320
      Width           =   1935
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "File Number: "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   7440
      TabIndex        =   43
      Top             =   7800
      Width           =   1935
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Purpose: "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   7440
      TabIndex        =   42
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Entry Date: "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   5
      Left            =   7440
      TabIndex        =   41
      Top             =   5400
      Width           =   1935
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tenant: "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Index           =   6
      Left            =   7440
      TabIndex        =   40
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Guarantor: "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Index           =   7
      Left            =   7440
      TabIndex        =   39
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "House Status: "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Index           =   8
      Left            =   7440
      TabIndex        =   38
      Top             =   6840
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      Height          =   7575
      Index           =   2
      Left            =   7320
      Top             =   720
      Width           =   5295
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      Height          =   7575
      Index           =   0
      Left            =   3120
      Top             =   720
      Width           =   4215
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Flat: "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   31
      Top             =   7680
      Width           =   855
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Business Location/ Address: "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Index           =   4
      Left            =   360
      TabIndex        =   26
      Top             =   5520
      Width           =   2775
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name of Org/Govt/Agency "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Index           =   3
      Left            =   840
      TabIndex        =   25
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Workplace : "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Index           =   2
      Left            =   1440
      TabIndex        =   24
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Occupation : "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   23
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   4
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   3255
      Index           =   0
      Left            =   10440
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NON ASCL STAFF DATA ENTRY"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   660
      Left            =   2520
      TabIndex        =   15
      Top             =   0
      Width           =   8535
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Road/Line: "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   6
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "L. G. A : "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Estate: "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   7080
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name(surname first): "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "State of Origin: "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No: "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Form Number: "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      Height          =   7575
      Index           =   1
      Left            =   360
      Top             =   720
      Width           =   2775
   End
End
Attribute VB_Name = "PrivateTenantReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
     
            On Error Resume Next
            If Text1.Text = "" Or Text2.Text = "" Or Text7.Text = "" Or _
                Text8.Text = "" Or Text3.Text = "" Or Text9.Text = "" Or _
                Text4.Text = "" Or Text10.Text = "" Or Text11.Text = "" Or _
                Text5.Text = "" Or Text6.Text = "" Or _
                    Combo1.Text = "Select State" Or Combo1.Text = "" Or _
                        Combo5.Text = "Select Workplace" Or Combo5.Text = "" Or _
                            Combo6.Text = "Select Estate" Or Combo6.Text = "" Or _
                                Combo8.Text = "Select House Type" Or Combo8.Text = "" Or _
                                    Combo9.Text = "Select Purpose" Or Combo9.Text = "" Or _
                                        Combo7.Text = "Select House Status" Or Combo7.Text = "" Or _
                                            Combo2.Text = "Select Tenant" Or Combo2.Text = "" Or _
                                                Combo3.Text = "Select Department" Or Combo3.Text = "" Then
                MsgBox "There are some missing field", vbInformation, "DENIED"
                    Exit Sub
            'ElseIf Image1.Picture = 0 Then
              '  MsgBox "Upload a passport", vbInformation, "SAVED PARAMETER"
              '  Exit Sub
            ElseIf Not IsNumeric(Text2.Text) Then
                MsgBox "Entry must be Numeric", vbInformation, "| - ALERT - |"
                Exit Sub
            ElseIf Len(Text2.Text) <> 11 Then
                MsgBox "Phone Number Must" & vbCr & _
                        "Be 11 Characters long", vbInformation, "Entry Error"
            
            Else
                'Call SavePicture(Image1.Picture, App.Path & "\tmpfile")
                'Dim strFromFile As String
                'Dim lngFileSize As Long
                'Dim FileNum As Integer
                'FileNum = FreeFile
                'lngFileSize = FileLen(App.Path & "\tmpfile")
                'strFromFile = String(lngFileSize, " ")
                'Open App.Path & "\tmpfile" For Binary As FileNum
                'Get FileNum, , strFromFile
                    
                    Dim m1, m2, m3, m4 As String
                    estate = Reg.Combo6.Text
                    ROAD = Reg.Text4.Text
                    flat = Reg.Text5.Text
        
            m1 = "[ESTATE]='" + estate + "'"
            m2 = "[road_line]='" + ROAD + "'"
            m3 = "[flat]='" + flat + "'"
            m4 = m1 & " AND " & m2 & " AND " & m3
            If rs.State = adStateOpen Then rs.Close
            rs.Open "select * from House where" & m4, db, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
                Counter = Counter + 1
                    MsgBox "THIS PARTICULAR FLAT HAS ALREADY" & vbCr & _
                        "BEEN REGISTERED TO " & rs![occp_name], vbCritical, "DUBLICATE ENTRY- ATTEMPT: " & Counter
                Else
                rs.AddNew
                    rs![form_no] = Text1.Text
                    rs![housestatus] = Combo7.Text
                    rs![estate] = Combo6.Text
                    rs![Road_line] = Text4.Text
                    rs![flat] = Text10.Text
                    rs![housetype] = Combo8.Text
                    rs![purpose] = Combo9.Text
                    
                    
                    rs![occp_name] = Text6.Text
                    rs![phone] = Text2.Text
                    rs![LGA] = Text7.Text
                    rs![OCCUPATION] = Text8.Text
                    rs![wpdesc] = Text3.Text
                    rs![wploc] = Text9.Text
                    rs![guarantor] = Text11.Text
                    rs![file_no] = Text5.Text
                    rs![entrydate] = DTPicker1.Value
                    rs![dept] = Combo3.Text
                    rs![workplace] = Combo5.Text
                    rs![originstate] = Combo1.Text
                    rs![tenant] = Combo2.Text
                    'rs![occp_photo] = strFromFile
                        Close FileNum
                        Kill App.Path & "\tmpfile"
                    rs.Update
                MsgBox "Record Save Successful!", vbInformation, "DATABASE MESSAGE"
                   
                End If
                
                If rs1.State = adStateOpen Then rs1.Close
            rs1.Open "select * from HouseMaster where" & m4, db, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
             rs1.AddNew
                    rs1![housecode] = Text1.Text
                    rs1![housestatus] = Combo7.Text
                    rs1![estate] = Combo6.Text
                    rs1![Road_line] = Text4.Text
                    rs1![flat] = Text10.Text
                    rs1![housetype] = Combo8.Text
                    rs1![purpose] = Combo9.Text
                    
            rs1.Update
            MsgBox "success"
             Call Command2_Click
            End If
            
            End If
End Sub
Private Sub Command2_Click()
    Text1.Text = ""
    Text2.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text3.Text = ""
    Text9.Text = ""
    Text4.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Combo1.Text = "Select State"
    Combo5.Text = "Select Workplace"
    Combo6.Text = "Select Estate"
    Combo8.Text = "Select House Type"
    Combo9.Text = "Select Purpose"
    Combo7.Text = "Select House Status"
    Combo2.Text = "Select Tenant"
    Combo3.Text = "Select Department"
    Set Me.Image1.Picture = Nothing
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub
Private Sub Command4_Click()
    Me.cDlImage.ShowOpen
    If Me.cDlImage.FileName & "" <> "" Then
        Set Me.Image1.Picture = LoadPicture(Me.cDlImage.FileName)
    End If
End Sub
Private Sub Form_Load()
     'Dim db As New ADODB.Connection
      '  Dim rs As New ADODB.Recordset
      '  db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
                    App.Path & "\Database\house.mdb;Persist Security Info=False"
   ' On Error Resume Next
   ' If rs.State = adStateOpen Then rs.Close
   ' rs.Open "select * from [lists]", db, adOpenDynamic, adLockOptimistic
       '     rs.MoveFirst
       '     Do While Not rs.EOF
       '     Combo3.AddItem rs![dept]
       '     Combo2.AddItem rs![tenant]
      '      Combo9.AddItem rs![purpose]
      '      Combo7.AddItem rs![housestatus]
       '     Combo1.AddItem rs![State]
        '    Combo6.AddItem rs![estate]
         '   Combo5.AddItem rs![workplace]
         '   Combo8.AddItem rs![housetype]
           ' rs.MoveNext
           ' Loop
    With Me
        .Top = (Screen.Height - .Height) / 2
        .Left = (Screen.Width - .Width) / 2
    End With
End Sub
