VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Admin Login"
   ClientHeight    =   2550
   ClientLeft      =   1650
   ClientTop       =   3660
   ClientWidth     =   4395
   LinkTopic       =   "Form7"
   ScaleHeight     =   2550
   ScaleWidth      =   4395
   Begin VB.CommandButton Command2 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   1800
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox txtpassword1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox txtusername1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Set con = New ADODB.Connection
 With con
    .Open "Provider=SQLOLEDB.1;Integrated Security=SSPI; Persist Security Info=False;Initial Catalog=emp_samip;Data Source=PROJECT-1"
    End With
    
    Set oRs = New ADODB.Recordset
        sql = "select * from login"
        oRs.Open sql, con, adOpenDynamic, adLockOptimistic
        
        oRs.MoveFirst
        Do While Not oRs.EOF
        
        If oRs("username").Value = txtusername1.Text Then
            UNAME = True
            
            If oRs("password").Value = txtpassword1.Text Then
                UPASS = True
                
                
                'MsgBox "Login successful.Welcome To Blood Bank", vbInformation
               Form8.Show
               Unload Me
               
               
                    Else
                        UPASS = False
                        MsgBox "Invalid Login.Incorrect Password!", vbCritical
                        'txtpassword11.SetFocus
                        Exit Do
                        Exit Sub
                    End If
                    
                Else
                    UNAME = False
                    MsgBox "Invalid Login.User not found!", vbCritical
                   ' txtusername1.SetFocus
                    Exit Do
                    Exit Sub
                End If
                
                oRs.MoveNext
                Loop
                
                oRs.close
            con.close
            
        
End Sub

Private Sub Command2_Click()
Unload Me
MDIForm1.Show

End Sub

