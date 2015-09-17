VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form6 
   Caption         =   "Serach Record"
   ClientHeight    =   9075
   ClientLeft      =   510
   ClientTop       =   825
   ClientWidth     =   12990
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form6"
   ScaleHeight     =   9075
   ScaleWidth      =   12990
   Begin VB.CommandButton Command5 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   30
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   29
      Top             =   6840
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      DataField       =   "dob"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1680
      TabIndex        =   26
      Top             =   3960
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   16056321
      CurrentDate     =   41549
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   1095
      Left            =   7800
      Top             =   5640
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1931
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
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=blood_bank_system;Data Source=PROJECT-1"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=blood_bank_system;Data Source=PROJECT-1"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "donarinfo"
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
   Begin VB.Frame Frame2 
      Caption         =   "search record "
      Height          =   2055
      Left            =   120
      TabIndex        =   20
      Top             =   360
      Width           =   6975
      Begin VB.CommandButton Command1 
         Caption         =   "CLOSE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   27
         Top             =   1440
         Width           =   5655
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         TabIndex        =   25
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   2655
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form6.frx":0000
         Left            =   4800
         List            =   "Form6.frx":0013
         TabIndex        =   23
         Text            =   "operator"
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form6.frx":0028
         Left            =   2880
         List            =   "Form6.frx":0041
         TabIndex        =   22
         Text            =   "search"
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3240
         TabIndex        =   21
         Top             =   720
         Width           =   2655
      End
   End
   Begin VB.TextBox txtlname 
      DataField       =   "lname"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   3480
      Width           =   2535
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   975
      Left            =   3840
      TabIndex        =   6
      Top             =   5160
      Width           =   375
   End
   Begin VB.TextBox txtphno 
      DataField       =   "phno"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   6360
      Width           =   2535
   End
   Begin VB.TextBox txtzip 
      DataField       =   "zip"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4320
      TabIndex        =   4
      Top             =   5640
      Width           =   1095
   End
   Begin VB.ComboBox Combo2 
      DataField       =   "bgroup"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   1680
      TabIndex        =   3
      Text            =   "Select Group"
      Top             =   4440
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "gender"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Text            =   "Select Gender"
      Top             =   4800
      Width           =   2535
   End
   Begin VB.TextBox txtdid 
      DataField       =   "did"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox txtfname 
      DataField       =   "fname"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   3000
      Width           =   2535
   End
   Begin VB.TextBox txtaddress 
      DataField       =   "address"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1680
      TabIndex        =   7
      Top             =   5160
      Width           =   2535
   End
   Begin VB.Label Label1 
      DataField       =   "age"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   28
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "First Name       :"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Last Name       :"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   18
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "  Date Of Birth    :"
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "(MM-DD-YY)"
      Height          =   255
      Left            =   4440
      TabIndex        =   16
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "  Gender            :"
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Age is"
      Height          =   255
      Left            =   6120
      TabIndex        =   14
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Address          :"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "  Phone No.      :"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Label Label9 
      Caption         =   "Zip Code"
      Height          =   255
      Left            =   4320
      TabIndex        =   11
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Blood Group    :"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label11 
      Caption         =   "Donar ID         :"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   1575
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FIELD
Dim Exp
Dim val
Private Sub Combo3_Change()
FIELD = Combo3.Text
End Sub

Private Sub Combo4_Change()
Exp = Combo4.Text
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF Then
MsgBox "YOU ARE ALREADY ON THE FIRST RECORD"
End If
End Sub

Private Sub Command3_Click()

val = "'" & Trim(Text9.Text) & "'"
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = ("select * from donarinfo where ") & FIELD & Exp & val
'Adodc1.refresh
End Sub

Private Sub Command4_Click()
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "select * from donarinfo"
Adodc1.refresh
End Sub

Private Sub Command5_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF Then
MsgBox "YOU ARE ALREADY ON THE LAST RECORD"
End If
End Sub

Private Sub Form_Load()
Combo3.AddItem "fname", 0
Combo3.AddItem "lname", 1
Combo3.AddItem "did", 2
Combo3.AddItem "age", 3
Combo3.AddItem "zip", 4
Combo3.AddItem "phno", 5

Combo4.AddItem "=", 0
Combo4.AddItem ">", 1
Combo4.AddItem "<", 2
Combo4.AddItem "<=", 3
Combo4.AddItem ">=", 4


End Sub

