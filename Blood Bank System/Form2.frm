VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000005&
   Caption         =   "BLOOD BANK SYSTEM"
   ClientHeight    =   7920
   ClientLeft      =   5295
   ClientTop       =   1965
   ClientWidth     =   6015
   FillColor       =   &H80000005&
   LinkTopic       =   "Form2"
   ScaleHeight     =   7920
   ScaleWidth      =   6015
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
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   7440
      Width           =   5055
   End
   Begin VB.TextBox txtusername 
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   6360
      Width           =   2295
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
      Left            =   480
      TabIndex        =   2
      Top             =   6960
      Width           =   5055
   End
   Begin VB.TextBox txtpassword 
      DataSource      =   "Adodc1"
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3240
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   6360
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   480
      Negotiate       =   -1  'True
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   1005.046
      ScaleMode       =   0  'User
      ScaleWidth      =   5055
      TabIndex        =   5
      Top             =   1440
      Width           =   5055
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "        BLOOD BANK                      SYSTEM"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1320
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
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
      Height          =   240
      Left            =   720
      TabIndex        =   3
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
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
      Left            =   3600
      TabIndex        =   4
      Top             =   6120
      Width           =   1455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim oRs As ADODB.Recordset
Dim sql As String
Dim UNAME, UPASS As String



Private Sub Command1_Click()
Set con = New ADODB.Connection
 With con
    .Open "Provider=SQLOLEDB.1;Integrated Security=SSPI; Persist Security Info=False;Initial Catalog=blood_bank_system;Data Source=PROJECT-1"
    End With
    
    Set oRs = New ADODB.Recordset
        sql = "select * from login"
        oRs.Open sql, con, adOpenDynamic, adLockOptimistic
        
        oRs.MoveFirst
        Do While Not oRs.EOF
        
        If oRs("username").Value = txtusername.Text Then
            UNAME = True
            
            If oRs("password").Value = txtpassword.Text Then
                UPASS = True
                
                
                MsgBox "Login successful.Welcome To Blood Bank", vbInformation
               MDIForm1.Show
               Unload Me
               
               
                    Else
                        UPASS = False
                        MsgBox "Invalid Login.Incorrect Password!", vbCritical
                        txtpassword.SetFocus
                        Exit Do
                        Exit Sub
                    End If
                    
                Else
                    UNAME = False
                    MsgBox "Invalid Login.User not found!", vbCritical
                    'txtusername.SetFocus
                    Exit Do
                    Exit Sub
                End If
                
                oRs.MoveNext
                Loop
                
                oRs.close
            con.close
            
        
End Sub


Private Sub Command2_Click()
answer = MsgBox("DO YOU REALLY WANTED TO EXIT", vbExclamation + vbYesNo)
If answer = vbYes Then
End
Command2 = False
End If

End Sub

