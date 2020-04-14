VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form PTUncompleted 
   Caption         =   "Private Tenant Uncompleted Registration Form"
   ClientHeight    =   8685
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12870
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   12870
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox qrtnameAddress 
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
      Height          =   885
      Left            =   9000
      TabIndex        =   38
      Top             =   6840
      Width           =   3015
   End
   Begin VB.TextBox Road_line 
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
      Left            =   2880
      TabIndex        =   35
      Top             =   7800
      Width           =   1455
   End
   Begin VB.ComboBox estate 
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
      Left            =   2880
      TabIndex        =   34
      Text            =   "Select Estate"
      Top             =   7320
      Width           =   3975
   End
   Begin VB.TextBox flat 
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
      Left            =   5400
      TabIndex        =   33
      Top             =   7800
      Width           =   1455
   End
   Begin VB.ComboBox nationality 
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
      ItemData        =   "PTUncompleted.frx":0000
      Left            =   2880
      List            =   "PTUncompleted.frx":0002
      TabIndex        =   30
      Text            =   "Select Nationality"
      Top             =   4680
      Width           =   3975
   End
   Begin VB.TextBox hometown 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   2880
      TabIndex        =   28
      Top             =   3360
      Width           =   3975
   End
   Begin VB.TextBox housecode 
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
      Left            =   2880
      TabIndex        =   14
      Top             =   960
      Width           =   3975
   End
   Begin VB.TextBox phoneno 
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
      Left            =   2880
      TabIndex        =   13
      Top             =   1920
      Width           =   3975
   End
   Begin VB.TextBox wpdesc 
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
      Height          =   885
      Left            =   2880
      TabIndex        =   12
      Top             =   5640
      Width           =   3975
   End
   Begin VB.ComboBox workplace 
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
      ItemData        =   "PTUncompleted.frx":0004
      Left            =   2880
      List            =   "PTUncompleted.frx":0006
      TabIndex        =   11
      Text            =   "Select Workplace"
      Top             =   5160
      Width           =   3975
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
      Left            =   10200
      TabIndex        =   10
      Top             =   2160
      Width           =   1815
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
      Left            =   10200
      TabIndex        =   9
      Top             =   3240
      Width           =   1815
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
      Height          =   495
      Left            =   7440
      TabIndex        =   8
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Frame Frame2 
      Height          =   2895
      Left            =   7080
      TabIndex        =   7
      Top             =   960
      Width           =   2895
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   2610
         Left            =   120
         Picture         =   "PTUncompleted.frx":0008
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2685
      End
   End
   Begin VB.TextBox occp_name 
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
      Left            =   2880
      TabIndex        =   6
      Top             =   1440
      Width           =   3975
   End
   Begin VB.ComboBox originstate 
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
      Left            =   2880
      TabIndex        =   5
      Text            =   "Select State"
      Top             =   2400
      Width           =   3975
   End
   Begin VB.TextBox lga 
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
      Left            =   2880
      TabIndex        =   4
      Top             =   2880
      Width           =   3975
   End
   Begin VB.ComboBox housestatus 
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
      Left            =   9000
      TabIndex        =   3
      Text            =   "Select House Status"
      Top             =   6360
      Width           =   3015
   End
   Begin VB.TextBox occupation 
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
      Left            =   2880
      TabIndex        =   2
      Top             =   4080
      Width           =   3975
   End
   Begin VB.TextBox wploc 
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
      Left            =   2880
      TabIndex        =   1
      Top             =   6840
      Width           =   3975
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
      Left            =   10200
      TabIndex        =   0
      Top             =   1080
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   360
      Top             =   6000
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
      Connect         =   $"PTUncompleted.frx":4A91
      OLEDBString     =   $"PTUncompleted.frx":4B1D
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   10920
      Top             =   9600
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
      Connect         =   $"PTUncompleted.frx":4BA9
      OLEDBString     =   $"PTUncompleted.frx":4C36
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
   Begin MSComDlg.CommonDialog cDlImage 
      Left            =   12840
      Top             =   10080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Quarter Name/ Address: "
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
      Height          =   600
      Index           =   5
      Left            =   7200
      TabIndex        =   37
      Top             =   7200
      Width           =   1770
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
      Height          =   495
      Left            =   -120
      TabIndex        =   36
      Top             =   7320
      Width           =   2775
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
      Left            =   0
      TabIndex        =   32
      Top             =   7800
      Width           =   2775
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nationality:"
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
      Index           =   10
      Left            =   -120
      TabIndex        =   31
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Home Town/Address"
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
      Index           =   9
      Left            =   -120
      TabIndex        =   29
      Top             =   3480
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "House Code: "
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
      Left            =   0
      TabIndex        =   27
      Top             =   960
      Width           =   2775
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
      Height          =   495
      Left            =   0
      TabIndex        =   26
      Top             =   1920
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
      Height          =   495
      Left            =   0
      TabIndex        =   25
      Top             =   2400
      Width           =   2775
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
      Height          =   495
      Left            =   0
      TabIndex        =   24
      Top             =   1440
      Width           =   2775
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
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   23
      Top             =   2880
      Width           =   2775
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TENANT-UNCOMPLETED BUILDING"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   855
      Left            =   960
      TabIndex        =   22
      Top             =   0
      Width           =   10455
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   4
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   3255
      Index           =   0
      Left            =   10080
      Top             =   960
      Width           =   2055
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
      Height          =   495
      Index           =   1
      Left            =   0
      TabIndex        =   21
      Top             =   4200
      Width           =   2775
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
      Height          =   495
      Index           =   2
      Left            =   0
      TabIndex        =   20
      Top             =   5160
      Width           =   2775
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name of Organisation:"
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
      Height          =   615
      Index           =   3
      Left            =   0
      TabIndex        =   19
      Top             =   5760
      Width           =   2775
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bisiness Location : "
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
      Left            =   0
      TabIndex        =   18
      Top             =   6840
      Width           =   2775
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
      Left            =   4440
      TabIndex        =   17
      Top             =   7800
      Width           =   855
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
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
      Height          =   495
      Index           =   8
      Left            =   7080
      TabIndex        =   16
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   5
      Height          =   7215
      Index           =   1
      Left            =   0
      Top             =   960
      Width           =   2775
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
      Index           =   2
      Left            =   4440
      TabIndex        =   15
      Top             =   7800
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      BorderWidth     =   5
      Height          =   7455
      Index           =   2
      Left            =   7080
      Top             =   840
      Width           =   5295
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      BorderWidth     =   5
      Height          =   7575
      Index           =   0
      Left            =   2760
      Top             =   840
      Width           =   4215
   End
End
Attribute VB_Name = "PTUncompleted"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
            If housecode.Text = "" Or occp_name.Text = "" Or _
                 estate.Text = "" Or phoneno.Text = "" Or _
                Road_line.Text = "" Or flat.Text = "" Or _
                 housestatus.Text = "" Or _
                        estate.Text = "Select Estate" Or estate.Text = "" Or _
                         nationality.Text = "Select Nationality" Or nationality.Text = "" Or _
                        qrtnameAddress.Text = "" Or _
                     originstate.Text = "Select State" Or originstate.Text = "" Then

                MsgBox "There are some missing field", vbInformation, "DENIED"
                    Exit Sub
            ElseIf Image1.Picture = 0 Then
                MsgBox "Upload a passport", vbInformation, "SAVED PARAMETER"
                Exit Sub
            ElseIf Not IsNumeric(phoneno.Text) Then
                MsgBox "Entry must be Numeric", vbInformation, "| - ALERT - |"
                Exit Sub
            ElseIf Len(phoneno.Text) <> 11 Then
                MsgBox "Phone Number Must" & vbCr & _
                        "Be 11 Characters long", vbInformation, "Entry Error"
            
            Else
                Call SavePicture(Image1.Picture, App.Path & "\tmpfile")
                Dim strFromFile As String
                Dim lngFileSize As Long
                Dim FileNum As Integer
                FileNum = FreeFile
                lngFileSize = FileLen(App.Path & "\tmpfile")
                strFromFile = String(lngFileSize, " ")
                Open App.Path & "\tmpfile" For Binary As FileNum
                Get FileNum, , strFromFile
                    
                    Dim m1, m2, m3, m4 As String
                    estate = estate.Text
                    road = Road_line.Text
                    flat = flat.Text
        
            m1 = "[estate]='" + estate + "'"
            m2 = "[road_line]='" + road + "'"
            m3 = "[flat]='" + flat + "'"
            m4 = m1 & " AND " & m2 & " AND " & m3
            If rs.State = adStateOpen Then rs.Close
            rs.Open "select * from occup_reg where" & m4, db, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
                counter = counter + 1
                    MsgBox "THIS PARTICULAR FLAT HAS ALREADY" & vbCr & _
                        "BEEN REGISTERED TO " & rs![occp_name], vbCritical, "DUBLICATE ENTRY- ATTEMPT: " & counter
                Else
                rs.AddNew
                    rs![form_no] = housecode.Text
                    rs![housestatus] = housestatus.Text
                    rs![estate] = estate.Text
                    rs![Road_line] = Road_line.Text
                    rs![flat] = flat.Text
                    rs![housetype] = "Unspecified"
                    rs![purpose] = "Residential"
                    
               '***********************************************"
                    rs![occp_name] = occp_name.Text
                    rs![phone] = phoneno.Text
                    rs![originstate] = originstate.Text
                    rs![lga] = lga.Text
                    rs![hometown] = hometown.Text
                    rs![nationality] = nationality.Text
                    rs![occupation] = occupation.Text
                    rs![workplace] = workplace.Text
                    rs![wpdesc] = wpdesc.Text
                    rs![wploc] = wploc.Text
                    rs![qrtnameAddress] = qrtnameAddress.Text
                    rs![tenant] = "UCB"
                    rs![occp_photo] = strFromFile
                        Close FileNum
                        Kill App.Path & "\tmpfile"
                    rs.Update
                MsgBox "Record Save Successful!", vbInformation, "DATABASE MESSAGE"
                   
                End If
                
                If rs1.State = adStateOpen Then rs1.Close
            rs1.Open "select * from HouseMaster where" & m4, db, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
             rs1.AddNew
                    rs1![housecode] = housecode.Text
                    rs1![housestatus] = housestatus.Text
                    rs1![estate] = estate.Text
                    rs1![Road_line] = Road_line.Text
                    rs1![flat] = flat.Text
                   ' rs1![housetype] = housetype.Text
                    'rs1![purpose] = purpose.Text
                    
            rs1.Update
            MsgBox "success"
             clear1
            End If
            End If
End Sub
Private Sub clear1()
                    housecode.Text = ""
                    housestatus.Text = ""
                    estate.Text = ""
                    Road_line.Text = ""
                    flat.Text = ""
                    occp_name.Text = ""
                    phoneno.Text = ""
                    originstate.Text = ""
                    lga.Text = ""
                    hometown.Text = ""
                    nationality.Text = ""
                    occupation.Text = ""
                    workplace.Text = ""
                    wpdesc.Text = ""
                    wploc.Text = ""
                    qrtnameAddress.Text = ""
End Sub
