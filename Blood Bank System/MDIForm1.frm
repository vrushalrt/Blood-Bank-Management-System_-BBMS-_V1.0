VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "BLOOD BANK SYSTEM"
   ClientHeight    =   7815
   ClientLeft      =   6555
   ClientTop       =   3630
   ClientWidth     =   14385
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   WindowState     =   2  'Maximized
   Begin VB.Menu file 
      Caption         =   "File"
      Index           =   1
      Begin VB.Menu newdonar 
         Caption         =   "NewDonar"
         Shortcut        =   ^N
      End
      Begin VB.Menu logoff 
         Caption         =   "LogOff"
         Shortcut        =   ^L
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu edit 
      Caption         =   "Edit"
      Begin VB.Menu updatedonar 
         Caption         =   "UpdateDonar"
         Shortcut        =   ^U
      End
      Begin VB.Menu admin 
         Caption         =   "Admin"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu view 
      Caption         =   "View"
      Begin VB.Menu viewsheet 
         Caption         =   "ViewSheet"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu reports 
      Caption         =   "Reports"
      Begin VB.Menu search 
         Caption         =   "Search"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu about 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub editdonar_Click()
Form1.Show
End Sub

Private Sub about_Click()
Form4.Show
End Sub

Private Sub admin_Click()
Form7.Show

End Sub

Private Sub exit_Click()

End

End Sub

Private Sub logoff_Click()
Form2.Show
MDIForm1.Hide
Form1.Hide
Form3.Hide
Form4.Hide

End Sub

Private Sub newdonar_Click()
Form1.Show
End Sub

Private Sub sheet_Click()
Form5.Show

End Sub

Private Sub search_Click()
Form6.Show
'Form9.Show

End Sub

Private Sub updatedonar_Click()
Form3.Show
Form1.Hide
End Sub

Private Sub viewsheet_Click()
Form5.Show

End Sub
