VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form8 
   Caption         =   "Admin Master"
   ClientHeight    =   4230
   ClientLeft      =   1455
   ClientTop       =   3465
   ClientWidth     =   7005
   LinkTopic       =   "Form8"
   ScaleHeight     =   4230
   ScaleWidth      =   7005
   Begin VB.CommandButton Command22 
      Caption         =   "CLOSE"
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   11
      Top             =   2280
      Width           =   3255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "DELETE"
      Height          =   375
      Left            =   3000
      TabIndex        =   10
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "UPDATE"
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ADD"
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   8
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command11 
      Caption         =   ">"
      Height          =   375
      Index           =   1
      Left            =   3840
      TabIndex        =   7
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   6
      Top             =   1920
      Width           =   495
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   480
      Top             =   3480
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=emp_samip;Data Source=PROJECT-1"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=emp_samip;Data Source=PROJECT-1"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "login"
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
   Begin VB.TextBox txtpassword2 
      DataField       =   "password"
      DataSource      =   "Adodc1"
      Height          =   375
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox txtpassword1 
      DataField       =   "password"
      DataSource      =   "Adodc1"
      Height          =   375
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox txtusername2 
      DataField       =   "username"
      DataSource      =   "Adodc1"
      Height          =   375
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   1920
      TabIndex        =   3
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Re-type Password"
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
      Top             =   1440
      Width           =   1695
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
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "New Username"
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
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command11_Click(Index As Integer)
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF Then
MsgBox "YOU ARE ALREADY ON THE LAST RECORD"
'Adodc1.Recordset.MoveLast
End If
End Sub
Private Sub Command1_Click(Index As Integer)
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF Then
MsgBox "YOU ARE ALREADY ON THE FIRST RECORD"
'Adodc1.Recordset.MoveFirst
End If
End Sub

Private Sub Command2_Click(Index As Integer)
Adodc1.Recordset.AddNew
MsgBox "Admin Added Successfully.", vbInformation
End Sub

Private Sub Command22_Click(Index As Integer)
Unload Me
MDIForm1.Show
End Sub

Private Sub Command3_Click()
Adodc1.Recordset.update
MsgBox "Account Updated", vbInformation

End Sub

Private Sub Command4_Click()
Adodc1.Recordset.Delete
MsgBox "Account Deleted", vbInformation



End Sub

Private Sub txtusername_Change(Index As Integer)

End Sub
