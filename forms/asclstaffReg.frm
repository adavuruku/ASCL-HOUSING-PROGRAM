VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form asclstaffReg 
   BackColor       =   &H000080FF&
   Caption         =   "Ascl Staff Registration form"
   ClientHeight    =   9225
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   13200
   StartUpPosition =   2  'CenterScreen
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
      Height          =   375
      Left            =   2640
      TabIndex        =   49
      Top             =   8520
      Width           =   4095
   End
   Begin VB.ComboBox Combo5 
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
      Left            =   9600
      TabIndex        =   46
      Text            =   "Select Purpose"
      Top             =   7560
      Width           =   3135
   End
   Begin VB.ComboBox Combo4 
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
      Left            =   2640
      TabIndex        =   45
      Text            =   "Select Purpose"
      Top             =   6960
      Width           =   4095
   End
   Begin VB.ComboBox Combo2 
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
      Left            =   2640
      TabIndex        =   42
      Text            =   "Select Purpose"
      Top             =   7920
      Width           =   4095
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
      Left            =   2640
      TabIndex        =   40
      Text            =   "Select Purpose"
      Top             =   7440
      Width           =   4095
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
      Height          =   375
      Left            =   9600
      TabIndex        =   35
      Top             =   6360
      Width           =   3135
   End
   Begin VB.CheckBox acomoNo 
      BackColor       =   &H000040C0&
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   13
      Top             =   6600
      Width           =   855
   End
   Begin VB.CheckBox acomoYes 
      BackColor       =   &H000040C0&
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   12
      Top             =   6600
      Width           =   975
   End
   Begin VB.TextBox spousefileno 
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   9600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   8070
      Width           =   3135
   End
   Begin VB.TextBox spousename 
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
      Left            =   2640
      TabIndex        =   15
      Top             =   2400
      Width           =   4095
   End
   Begin VB.TextBox file_no 
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
      Left            =   2640
      TabIndex        =   1
      Top             =   1200
      Width           =   4095
   End
   Begin VB.TextBox designation 
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
      Left            =   2640
      TabIndex        =   7
      Top             =   4800
      Width           =   4095
   End
   Begin VB.TextBox emailaddress 
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
      Left            =   9600
      TabIndex        =   11
      Top             =   4680
      Width           =   3135
   End
   Begin VB.TextBox ippisno 
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
      Left            =   9600
      TabIndex        =   10
      Top             =   5160
      Width           =   3135
   End
   Begin VB.ComboBox step 
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
      Left            =   5040
      TabIndex        =   9
      Text            =   "Step"
      Top             =   5400
      Width           =   1695
   End
   Begin VB.ComboBox dept 
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
      Left            =   2640
      TabIndex        =   5
      Text            =   "Select Dept"
      Top             =   3600
      Width           =   4095
   End
   Begin MSComCtl2.DTPicker dofa 
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   6000
      Width           =   4095
      _ExtentX        =   7223
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
      CalendarBackColor=   65535
      CalendarForeColor=   0
      CalendarTitleBackColor=   16776960
      CalendarTitleForeColor=   33023
      CalendarTrailingForeColor=   255
      Format          =   80871425
      CurrentDate     =   41658
   End
   Begin MSComCtl2.DTPicker dob 
      Height          =   375
      Left            =   9600
      TabIndex        =   2
      Top             =   5760
      Width           =   3135
      _ExtentX        =   5530
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
      CalendarBackColor=   65535
      CalendarForeColor=   0
      CalendarTitleBackColor=   16776960
      CalendarTitleForeColor=   33023
      CalendarTrailingForeColor=   255
      Format          =   80871425
      CurrentDate     =   41658
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
      Left            =   2640
      TabIndex        =   14
      Top             =   3000
      Width           =   4095
   End
   Begin VB.ComboBox gradelevel 
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
      Left            =   2640
      TabIndex        =   8
      Text            =   "grade"
      Top             =   5400
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
      Height          =   735
      Left            =   10920
      TabIndex        =   19
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
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
      Height          =   720
      Left            =   10920
      TabIndex        =   20
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00008000&
      Height          =   3255
      Left            =   7200
      TabIndex        =   17
      Top             =   960
      Width           =   3015
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   2850
         Left            =   240
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2565
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
      Left            =   2640
      TabIndex        =   0
      Top             =   1800
      Width           =   4095
   End
   Begin VB.ComboBox maritalstatus 
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
      Left            =   9600
      TabIndex        =   3
      Text            =   "Select Status"
      Top             =   6960
      Width           =   3135
   End
   Begin VB.TextBox divv 
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
      Left            =   2640
      TabIndex        =   6
      Top             =   4200
      Width           =   4095
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
      Height          =   735
      Left            =   10920
      TabIndex        =   18
      Top             =   1320
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   600
      Top             =   9840
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   11160
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
   Begin VB.PictureBox cDlImage 
      Height          =   480
      Left            =   13080
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   37
      Top             =   10080
      Width           =   1200
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Home Town :"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   840
      TabIndex        =   48
      Top             =   8520
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nationality :"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   840
      TabIndex        =   47
      Top             =   6960
      Width           =   1575
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "T Category :"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   7800
      TabIndex        =   44
      Top             =   7560
      Width           =   1695
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Local Govt : "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   960
      TabIndex        =   43
      Top             =   7920
      Width           =   1575
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "State : "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   1440
      TabIndex        =   41
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Form Number:"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   600
      TabIndex        =   39
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name(Othername):"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   285
      TabIndex        =   38
      Top             =   2400
      Width           =   2130
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Number"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   6
      Left            =   7920
      TabIndex        =   36
      Top             =   6480
      Width           =   1590
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tenants Remarks :"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Index           =   3
      Left            =   8160
      TabIndex        =   34
      Top             =   8040
      Width           =   1335
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Designation :"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   12
      Left            =   840
      TabIndex        =   33
      Top             =   4800
      Width           =   1500
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email address"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   11
      Left            =   8040
      TabIndex        =   32
      Top             =   4680
      Width           =   1530
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IPPIS NO"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   360
      Index           =   10
      Left            =   8520
      TabIndex        =   31
      Top             =   5160
      Width           =   1005
   End
   Begin VB.Line Line1 
      X1              =   4800
      X2              =   4800
      Y1              =   5400
      Y2              =   5880
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Division :"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   9
      Left            =   1320
      TabIndex        =   30
      Top             =   4200
      Width           =   1050
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date ofFirst Appt: "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   360
      TabIndex        =   29
      Top             =   6000
      Width           =   2160
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "File Number:"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   480
      TabIndex        =   28
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth: "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   7920
      TabIndex        =   27
      Top             =   5880
      Width           =   1650
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name(Surname ): "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   360
      Left            =   360
      TabIndex        =   26
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Marita Status"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   0
      Left            =   7920
      TabIndex        =   25
      Top             =   7080
      Width           =   1545
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASCL STAFF DATA ENTRY"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   900
      Left            =   1875
      TabIndex        =   24
      Top             =   0
      Width           =   9675
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      BorderWidth     =   4
      FillColor       =   &H00E0E0E0&
      Height          =   3135
      Index           =   0
      Left            =   10440
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Department :"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   1
      Left            =   840
      TabIndex        =   23
      Top             =   3600
      Width           =   1500
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grade Level/Step"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Index           =   2
      Left            =   360
      TabIndex        =   22
      Top             =   5400
      Width           =   1950
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Are You Accommodated?: "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   600
      Index           =   4
      Left            =   480
      TabIndex        =   21
      Top             =   6360
      Width           =   1965
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      BorderWidth     =   5
      Height          =   8175
      Index           =   1
      Left            =   120
      Top             =   960
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      BorderWidth     =   5
      Height          =   4575
      Index           =   2
      Left            =   7200
      Top             =   4440
      Width           =   5775
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000040C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      BorderWidth     =   5
      Height          =   8175
      Index           =   0
      Left            =   2400
      Top             =   960
      Width           =   10695
   End
End
Attribute VB_Name = "asclstaffReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub acomoNo_Click()
If acomoNo.Value = 1 Then
acomoYes.Value = 0
estate.Enabled = False
Road_line.Enabled = False
flat.Enabled = False
housetype.Enabled = False
entrydate.Enabled = False
housestatus.Enabled = False
purpose.Enabled = False
End If

End Sub

Private Sub acomoYes_Click()
If acomoYes.Value = 1 Then
acomoNo.Value = 0
estate.Enabled = True
Road_line.Enabled = True
flat.Enabled = True
housetype.Enabled = True
entrydate.Enabled = True
housestatus.Enabled = True
purpose.Enabled = True
End If
End Sub

Private Sub Command1_Click()
 On Error Resume Next
            If occp_name.Text = "" Or file_no.Text = "" Or _
                maritalstatus.Text = "" Or divv.Text = "" Or _
                designation.Text = "" Or gradelevel.Text = "" Or step.Text = "" Or _
                ippisno.Text = "" Or phoneno.Text = "" Or _
                Road_line.Text = "" Or _
                maritalstatus.Text = "Select Status" Or maritalstatus.Text = "" Or _
                  dept.Text = "Select Dept" Or dept.Text = "" Then
                MsgBox "There are some missing field", vbInformation, "DENIED"
                    Exit Sub
            'ElseIf Image1.Picture = 0 Then
             '   MsgBox "Upload a passport", vbInformation, "SAVED PARAMETER"
              '  Exit Sub
            ElseIf Not IsNumeric(phoneno.Text) Then
                MsgBox "Entry must be Numeric", vbInformation, "| - ALERT - |"
                Exit Sub
            ElseIf Len(phoneno.Text) <> 11 Then
                MsgBox "Phone Number Must" & vbCr & _
                        "Be 11 Characters long", vbInformation, "Entry Error"
            
            Else
             If rs.State = adStateOpen Then rs.Close
                        rs.Open "SELECT * FROM [TenantT]", db, 3, 3
                rs.AddNew
                    rs![FORMNO] = Road_line.Text
                   ' rs![housestatus] = housestatus.Text
                    rs![other names] = spousename.Text
                    rs![surname] = occp_name.Text
                    rs![TFileNo] = file_no.Text
                    rs![Department] = dept.Text
                    rs![DIVISION] = divv.Text
                    
               '***********************************************"
                    rs![designation] = designation.Text
                    rs![Grade Level] = gradelevel.Text
                    rs![step] = step.Text
                    rs![date of birth] = dob.Value
                    rs![marital status] = maritalstatus.Text
                    rs![date of first appt] = dofa.Value
                    rs![state of origin] = Combo1.Text
                    rs![LGA] = Combo2.Text
                    rs![emailaddress] = emailaddress.Text
                   rs![Home Town] = Text1.Text
                  rs![ippisno] = ippisno.Text
                   
                    rs![phoneno] = phoneno.Text
                    rs![Nationality] = Combo4.Text
                    rs![TCATEGORY] = Combo5.Text
                    rs![ippisno] = ippisno.Text
                    rs![emialaddress] = emailaddress.Text
                    rs![TENANT REMARK] = spousefileno.Text
                        'hard coded
                        rs![tenant] = "ASCL"
                        rs![occupation] = "Civil Servant"
                        rs![WORK PLACE] = "Ajaokuta"
                        
                    rs.Update
                MsgBox "Record Save Successful!", vbInformation, "DATABASE MESSAGE"
                   
                End If
                
                
             clearr
           ' End If
          '  End If
End Sub
Private Sub clearr()
                    ' housecode.Text = ""
                       ' housestatus.Text = ""
                    ' estate.Text = ""
                     Road_line.Text = ""
                   ' flat.Text = ""
                    'housetype.Text = ""
                   ' purpose.Text = ""
                    Combo5.Text = ""
                      Combo4.Text = ""
                        Combo1.Text = ""
                        Combo2.Text = ""
               '***********************************************"
                    occp_name.Text = ""
                    phoneno.Text = ""
                    file_no.Text = ""
'                    dob.Value = ""
                    maritalstatus.Text = ""
                    maritalstatus.Text = "Select Status"
                    'dofa.Value = ""
                    dept.Text = ""
                    divv.Text = ""
                    designation.Text = ""
                    'Text11.Text
                    'entrydate
                    gradelevel.Text = ""
                    step.Text = ""
                    ippisno.Text = ""
                    spousename = ""
                    spousefileno = ""
                    emailaddress.Text = ""
End Sub

Private Sub Command2_Click()
clearr
End Sub

Private Sub Command3_Click()
 dob.Value = "01/01/2014"
End Sub

Private Sub Command4_Click()

End Sub

Private Sub file_no_LostFocus()
On Error Resume Next
If rs.State = adStateOpen Then rs.Close
    rs.Open "SELECT * FROM [TenantT] where TFileNo='" + file_no.Text + "'", db, 3, 3
               ' rs.AddNew
               If rs.EOF Then
               Exit Sub
               Command3.Enabled = False
                Command1.Enabled = True
                      Command2.Enabled = True
               Else
                     Road_line.Text = rs![FORMNO]
                   ' rs![housestatus] = housestatus.Text
                     spousename.Text = rs![other names]
                   occp_name.Text = rs![surname]
                   ' rs![TFileNo] = file_no.Text
                   dept.Text = rs![Department]
                    divv.Text = rs![DIVISION]
                    
               '***********************************************"
                    designation.Text = rs![designation]
                     gradelevel.Text = rs![Grade Level]
                   step.Text = rs![step]
                     'dob.Value = rs![date of birth]
                    'Date dd = rs![date of birth]
                     
                    maritalstatus.Text = rs![marital status]
                    'dofa.Value = rs![date of first appt]
                    Combo1.Text = rs![state of origin]
                    Combo2.Text = rs![LGA]
                    emailaddress.Text = rs![emailaddress]
                   Text1.Text = rs![Home Town]
                ippisno.Text = rs![ippisno]
                   
                     phoneno.Text = rs![phoneno]
                     Combo4.Text = rs![Nationality]
                     Combo5.Text = rs![TCATEGORY]
                    ippisno.Text = rs![ippisno]
                     emailaddress.Text = rs![emailaddress]
                   spousefileno.Text = rs![TENANT REMARK]
                        'hard coded
                       ' rs![tenant] = "ASCL"
                       ' rs![occupation] = "Civil Servant"
                        'rs![WORK PLACE] = "Ajaokuta"
                        
                    'rs.Update
                    'MsgBox dd
                    Set Image1.Picture = LoadPicture(App.Path & "\photos\" & rs![TFileNo] & ".jpg")
                    Command3.Enabled = True
                     Command1.Enabled = False
                      Command2.Enabled = False
                MsgBox "Record Save Retrieve!", vbInformation, "DATABASE MESSAGE"
                   
                End If
                
                
            ' clearr

End Sub

Private Sub Form_Load()
Command3.Enabled = False
End Sub
