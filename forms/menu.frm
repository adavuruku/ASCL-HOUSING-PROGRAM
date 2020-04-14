VERSION 5.00
Begin VB.MDIForm menu 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3030
   ClientLeft      =   225
   ClientTop       =   1155
   ClientWidth     =   4560
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   10935
      Left            =   0
      Picture         =   "menu.frx":0000
      ScaleHeight     =   10875
      ScaleWidth      =   4500
      TabIndex        =   0
      Top             =   0
      Width           =   4560
   End
   Begin VB.Menu mmnreg 
      Caption         =   "mm"
   End
   Begin VB.Menu mnuascl 
      Caption         =   "ASCL STAFF DATA"
   End
   Begin VB.Menu mnunonascl 
      Caption         =   "NON ASCL STAFF DATA"
   End
   Begin VB.Menu mnuhouse 
      Caption         =   "House Data"
   End
End
Attribute VB_Name = "menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mmnreg_Click()
Housefilter.Show
End Sub

Private Sub mnuascl_Click()
asclstaffReg.Show
End Sub

Private Sub mnuhouse_Click()
Form1.Show
End Sub

Private Sub mnunonascl_Click()
PrivateTenantReg.Show
End Sub
