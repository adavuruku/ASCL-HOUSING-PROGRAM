VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Update_Info 
   Caption         =   "Form1"
   ClientHeight    =   8685
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12405
   LinkTopic       =   "Form1"
   Picture         =   "Update_Info.frx":0000
   ScaleHeight     =   8685
   ScaleWidth      =   12405
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Browse"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7680
      TabIndex        =   47
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Retrieve"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   37
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10320
      TabIndex        =   36
      Top             =   2520
      Width           =   1815
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
      Left            =   3000
      TabIndex        =   24
      Top             =   1320
      Width           =   1815
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
      Left            =   3000
      TabIndex        =   23
      Top             =   2280
      Width           =   3975
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
      Height          =   2445
      Left            =   3000
      TabIndex        =   22
      Top             =   4680
      Width           =   3975
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
      Left            =   3000
      TabIndex        =   21
      Top             =   8160
      Width           =   1455
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
      ItemData        =   "Update_Info.frx":664BB
      Left            =   9240
      List            =   "Update_Info.frx":664BD
      TabIndex        =   20
      Text            =   "Select Tenant"
      Top             =   6120
      Width           =   3015
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
      Left            =   9240
      TabIndex        =   19
      Text            =   "Select Department"
      Top             =   7560
      Width           =   3015
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
      ItemData        =   "Update_Info.frx":664BF
      Left            =   3000
      List            =   "Update_Info.frx":664C1
      TabIndex        =   18
      Text            =   "Workplace"
      Top             =   4200
      Width           =   3975
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
      Left            =   3000
      TabIndex        =   17
      Text            =   "Select Estate"
      Top             =   7680
      Width           =   3975
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
      Left            =   10320
      TabIndex        =   16
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Height          =   2895
      Left            =   7200
      TabIndex        =   15
      Top             =   1440
      Width           =   2895
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   2610
         Left            =   120
         Picture         =   "Update_Info.frx":664C3
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2685
      End
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
      Left            =   3000
      TabIndex        =   14
      Top             =   1800
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
      Height          =   360
      Left            =   3000
      TabIndex        =   13
      Text            =   "Select State of Origin"
      Top             =   2760
      Width           =   3975
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
      Left            =   3000
      TabIndex        =   12
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
      Height          =   360
      Left            =   9240
      TabIndex        =   10
      Top             =   7080
      Width           =   3015
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
      Left            =   3000
      TabIndex        =   9
      Top             =   3720
      Width           =   3975
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
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   7200
      Width           =   3975
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
      Left            =   5520
      TabIndex        =   7
      Top             =   8160
      Width           =   1455
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
      Left            =   9240
      TabIndex        =   6
      Top             =   6600
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
      Height          =   360
      Left            =   9240
      TabIndex        =   5
      Top             =   4680
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
      Height          =   360
      Left            =   9240
      TabIndex        =   4
      Top             =   5160
      Width           =   3015
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
      Left            =   9240
      TabIndex        =   3
      Top             =   8040
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Update"
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
      Left            =   10320
      TabIndex        =   2
      Top             =   1560
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   9240
      TabIndex        =   11
      Top             =   5640
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
      Format          =   26869761
      CurrentDate     =   41638
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
      Left            =   7200
      TabIndex        =   46
      Top             =   7200
      Width           =   1935
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
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
      Height          =   495
      Index           =   7
      Left            =   7200
      TabIndex        =   45
      Top             =   6720
      Width           =   1935
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
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
      Height          =   495
      Index           =   6
      Left            =   7200
      TabIndex        =   44
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
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
      Height          =   495
      Index           =   5
      Left            =   7200
      TabIndex        =   43
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
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
      Height          =   495
      Left            =   7200
      TabIndex        =   42
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
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
      Left            =   7200
      TabIndex        =   41
      Top             =   8160
      Width           =   1935
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
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
      Height          =   495
      Left            =   7200
      TabIndex        =   40
      Top             =   7680
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
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
      Height          =   495
      Left            =   7200
      TabIndex        =   39
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      BorderWidth     =   5
      Height          =   7215
      Index           =   2
      Left            =   7080
      Top             =   1320
      Width           =   5295
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   5
      Height          =   7215
      Index           =   1
      Left            =   120
      Top             =   1320
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
      Left            =   4560
      TabIndex        =   38
      Top             =   8160
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      BorderWidth     =   5
      Height          =   7455
      Index           =   0
      Left            =   2880
      Top             =   1200
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      Height          =   495
      Left            =   120
      TabIndex        =   35
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
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
      Left            =   120
      TabIndex        =   34
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
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
      Left            =   120
      TabIndex        =   33
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
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
      Left            =   120
      TabIndex        =   32
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
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
      Left            =   120
      TabIndex        =   31
      Top             =   7680
      Width           =   2775
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
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
      Left            =   120
      TabIndex        =   30
      Top             =   3240
      Width           =   2775
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
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
      Left            =   120
      TabIndex        =   29
      Top             =   8160
      Width           =   2775
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   4
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   2895
      Index           =   0
      Left            =   10200
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
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
      Left            =   120
      TabIndex        =   28
      Top             =   3720
      Width           =   2775
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
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
      Left            =   120
      TabIndex        =   27
      Top             =   4200
      Width           =   2775
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Workplace Description : "
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
      Height          =   2535
      Index           =   3
      Left            =   120
      TabIndex        =   26
      Top             =   4680
      Width           =   2775
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Workplace Location : "
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
      Left            =   120
      TabIndex        =   25
      Top             =   7200
      Width           =   2775
   End
   Begin VB.Label Label14 
      Caption         =   "              RETRIEVE AND UPDATE RECORD        "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   720
      Width           =   7935
   End
   Begin VB.Label Label12 
      Caption         =   " UNIT REGISTRY SYSTEM "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   735
      Index           =   1
      Left            =   2040
      TabIndex        =   0
      Top             =   0
      Width           =   7935
   End
End
Attribute VB_Name = "Update_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Command4_Click()
'''' RETRIEVE RECORD
     'Dim db As New ADODB.Connection
       ' Dim rs As New ADODB.Recordset
        'db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
                    App.Path & "\Database\house.mdb;Persist Security Info=False"
            On Error Resume Next
            If Text1.Text = "" Then
            MsgBox "HouseCode must be entered", vbCritical, "ALERT"
            ElseIf Image1.Picture = 0 Then
                MsgBox "Upload a passport", vbInformation, "SAVED PARAMETER"
                Exit Sub
            Else
                
                Dim m1 As String
                formno = Text1.Text
                m1 = "[form_no]='" + formno + "'"
                If rs.State = adStateOpen Then rs.Close
                rs.Open "select * from occup_reg where" & m1, db, adOpenDynamic, adLockOptimistic
                If Not rs.EOF Then
                    Text1.Text = rs![form_no]
                    Text6.Text = rs![occp_name]
                    Text2.Text = rs![phone]
                    Text7.Text = rs![lga]
                    Text8.Text = rs![occupation]
                    Text3.Text = rs![wpdesc]
                    Text9.Text = rs![wploc]
                    Text4.Text = rs![Road_line]
                    Text10.Text = rs![flat]
                    Text11.Text = rs![guarantor]
                    Text5.Text = rs![file_no]
                    
                    DTPicker1.Value = rs![entrydate]
                    
                    Combo3.Text = rs![dept]
                    Combo5.Text = rs![workplace]
                    Combo1.Text = rs![originstate]
                    Combo2.Text = rs![tenant]
                    Combo8.Text = rs![housetype]
                    Combo9.Text = rs![purpose]
                    Combo6.Text = rs![estate]
                    Combo7.Text = rs![housestatus]
                    
                    
                    Dim data As String
                data = rs("occp_photo")
                Open App.Path & "\tempfile" For Binary As #1
                Put 1, , data
                Image1.Picture = LoadPicture(App.Path & "\tempfile")
                Close #1
                Kill App.Path & "\tempfile"
                MsgBox "Record Retrieved Successful!", vbInformation, "DATABASE MESSAGE"
                Else
                    MsgBox "Recorded doesn't exist" & vbCr & _
                    "The Record never existed or" & vbCr & _
                                "has been deleted", vbCritical, _
                                                    "DATABASE RE - QUERY"
                End If
            End If
End Sub
Private Sub Command1_Click()
''''' UPDATE RECORD
   'Dim db As New ADODB.Connection
        'Dim rs As New ADODB.Recordset
        'db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
                    App.Path & "\Database\house.mdb;Persist Security Info=False"
                    If Text1.Text = "" Then
                    MsgBox "Enter form number"
                    Exit Sub
                Else
        Dim m1 As String
            formno = Text1.Text
            m1 = "[form_no]='" + formno + "'"
            If rs.State = adStateOpen Then rs.Close
                    rs.Open "select * from occup_reg where" & m1, _
                                        db, adOpenDynamic, adLockOptimistic
                    
                    Call SavePicture(Image1.Picture, App.Path & "\tmpfile")
                    Dim strFromFile As String
                    Dim lngFileSize As Long
                    Dim FileNum As Integer
                    FileNum = FreeFile
                    lngFileSize = FileLen(App.Path & "\tmpfile")
                    strFromFile = String(lngFileSize, " ")
                    Open App.Path & "\tmpfile" For Binary As FileNum
                    Get FileNum, , strFromFile
                    
                    rs![form_no] = Update_Info.Text1.Text
                    rs![occp_name] = Text6.Text
                    rs![phone] = Text2.Text
                    rs![lga] = Text7.Text
                    rs![occupation] = Text8.Text
                    rs![wpdesc] = Text3.Text
                    rs![wploc] = Text9.Text
                    rs![Road_line] = Text4.Text
                    rs![flat] = Text10.Text
                    rs![guarantor] = Text11.Text
                    rs![file_no] = Text5.Text
                    rs![entrydate] = DTPicker1.Value
                    rs![dept] = Combo3.Text
                    rs![workplace] = Combo5.Text
                    rs![originstate] = Combo1.Text
                    rs![tenant] = Combo2.Text
                    rs![housetype] = Combo8.Text
                    rs![purpose] = Combo9.Text
                    rs![estate] = Combo6.Text
                    rs![housestatus] = Combo7.Text
                    
                    rs![occp_photo] = strFromFile
                        Close FileNum
                        Kill App.Path & "\tmpfile"
                    rs.Update
                MsgBox "Record Updated successfully", vbInformation, "DATABASE MESSAGE"
                Call reset
            End If
End Sub
Private Sub Command2_Click()
'delete record
 'Dim db As New ADODB.Connection
      '  Dim rs As New ADODB.Recordset
       ' db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
                    App.Path & "\Database\house.mdb;Persist Security Info=False"
                Dim m1 As String
            formno = Text1.Text
            m1 = "[form_no]='" + formno + "'"
            If formno = "" Then
                MsgBox "Form Number must be entered", vbInformation, "NOTE"
        Else
        Dim C As String
        C = MsgBox("Do you want to DELETE this RECORD", vbYesNo, "confirmation")
                Select Case C
                        Case vbNo
                Exit Sub
        Case vbYes
            rs.Open "delete * from occup_reg where" & m1, _
                                        db, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
            MsgBox "RECORD DELETED", vbInformation, "DONE"
            If rs.State = adStateOpen Then rs.Close
                Call reset
            Else
                MsgBox "This record does not exist", vbCritical, "ERROR MESSAGE"
                Exit Sub
            End If
       End Select
    End If
End Sub

Private Sub Form_Load()
  ' Dim db As New ADODB.Connection
     '   Dim rs As New ADODB.Recordset
       ' db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
                    App.Path & "\Database\house.mdb;Persist Security Info=False"
    On Error Resume Next
    If rs.State = adStateOpen Then rs.Close
    rs.Open "select * from [lists]", db, _
                            adOpenDynamic, adLockOptimistic
            rs.MoveFirst
            Do While Not rs.EOF
            Combo3.AddItem rs![dept]
            Combo2.AddItem rs![tenant]
            Combo9.AddItem rs![purpose]
            Combo7.AddItem rs![housestatus]
            Combo1.AddItem rs![State]
            Combo6.AddItem rs![estate]
            Combo5.AddItem rs![workplace]
            Combo8.AddItem rs![housetype]
            rs.MoveNext
            Loop
    With Me
        .Top = (Screen.Height - .Height) / 2
        .Left = (Screen.Width - .Width) / 2
    End With
End Sub
Private Sub reset()
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
