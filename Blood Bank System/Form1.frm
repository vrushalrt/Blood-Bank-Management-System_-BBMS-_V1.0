VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   Caption         =   "Register New Donar"
   ClientHeight    =   7200
   ClientLeft      =   1080
   ClientTop       =   1395
   ClientWidth     =   10470
   FillColor       =   &H000000C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   10470
   Begin VB.TextBox txtfname 
      DataField       =   "fname"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1440
      TabIndex        =   24
      Top             =   1440
      Width           =   2535
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      DataField       =   "dob"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1440
      TabIndex        =   23
      Top             =   2400
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
      Format          =   16187393
      CurrentDate     =   41549
   End
   Begin VB.TextBox txtlname 
      DataField       =   "lname"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1440
      TabIndex        =   22
      Top             =   1920
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "gender"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1440
      TabIndex        =   20
      Text            =   "Select Gender"
      Top             =   3720
      Width           =   1935
   End
   Begin VB.ComboBox Combo2 
      DataField       =   "bgroup"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1440
      TabIndex        =   19
      Text            =   "Select Group"
      Top             =   2880
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   7440
      Top             =   4920
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
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
   Begin VB.TextBox txtzip 
      DataField       =   "zip"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   4080
      TabIndex        =   17
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton close 
      Caption         =   "CLOSE"
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
      Left            =   1680
      TabIndex        =   15
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton reset 
      Caption         =   "RESET"
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
      Left            =   2640
      TabIndex        =   14
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton submit 
      Caption         =   "ADD NEW"
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
      TabIndex        =   13
      Top             =   6240
      Width           =   1815
   End
   Begin VB.TextBox txtphno 
      DataField       =   "phno"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1440
      TabIndex        =   12
      Top             =   5640
      Width           =   2535
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   975
      Left            =   3600
      TabIndex        =   10
      Top             =   4440
      Width           =   375
   End
   Begin VB.TextBox txtaddress 
      DataField       =   "address"
      DataSource      =   "Adodc1"
      Height          =   975
      Left            =   1440
      TabIndex        =   9
      Top             =   4440
      Width           =   2535
   End
   Begin VB.TextBox txtdid 
      DataField       =   "did"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label lblage 
      DataField       =   "age"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   5760
      TabIndex        =   25
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label11 
      Caption         =   "Donar ID   :"
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
      Left            =   120
      TabIndex        =   21
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "Blood Group    :"
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
      TabIndex        =   18
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Zip Code"
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
      Left            =   4200
      TabIndex        =   16
      Top             =   4440
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "  Phone No.      :"
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
      Left            =   0
      TabIndex        =   11
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Address     :"
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
      Left            =   120
      TabIndex        =   8
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Age is"
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
      Left            =   5160
      TabIndex        =   7
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "  Gender            :"
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
      Left            =   0
      TabIndex        =   6
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "(DD-MM-YY)"
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
      Left            =   3960
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "  Date Of Birt :"
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
      Left            =   0
      TabIndex        =   4
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Last Name    :"
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
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "First Name   :"
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
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   "        NEW DONAR"
      BeginProperty Font 
         Name            =   "Chiller"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub close_Click()
Form1.Hide
MDIForm1.Show

End Sub

Private Sub Command1_Click()
Adodc1.Recordset.update
End Sub

Private Sub Form_Load()

txtdid.Text = ""
txtfname.Text = ""
txtlname.Text = ""
txtaddress.Text = ""
txtzip.Text = ""
txtphno.Text = ""


Combo1.AddItem "MALE", 0
Combo1.AddItem "FEMALE", 1
Combo1.AddItem "OTHER", 2

Combo2.AddItem "A+", 0
Combo2.AddItem "B+", 1
Combo2.AddItem "AB+", 2
Combo2.AddItem "O+", 3
Combo2.AddItem "A-", 4
Combo2.AddItem "B-", 5
Combo2.AddItem "AB-", 6
Combo2.AddItem "O-", 7

End Sub

Private Sub reset_Click()
'Adodc1.refresh
Form1.Cls
End Sub

Private Sub submit_Click()
Dim bday As Integer
Dim bmonth As Integer
Dim byear As Integer


bday = DTPicker1.Day
bmonth = DTPicker1.Month
byear = DTPicker1.Year

If Year(Now) <= DTPicker1.Year Then
lblage.Caption = 0
ElseIf (bmonth < Month(Now)) Or (bday < Day(Now)) Then
lblage.Caption = Abs((Year(Now) - DTPicker1.Year) - 1)
Else
lblage.Caption = Abs((Year(Now) - DTPicker1.Year))
End If

    If lblage.Caption <= 18 Then
    MsgBox "Donar is under age", vbCritical
    End If
    
    
'FieldClear
'lstload
If txtdid.Text = "" Then
    MsgBox "*DonarID Field Requird!", vbCritical
    ElseIf txtfname.Text = "" Then
    MsgBox "*FirstName Field Required!", vbCritical
    ElseIf txtlname.Text = "" Then
    MsgBox "*LastName Field Required!", vbCritical
    ElseIf txtaddress.Text = "" Then
    MsgBox "*Address Field Required!", vbCritical
    ElseIf txtzip.Text = "" Then
    MsgBox "*ZipCode Field Required!", vbCritical
    ElseIf txtphno.Text = "" Then
    MsgBox "*Phone no. field mandatory!", vbCritical
    ElseIf lblage.Caption <= 18 Then
        'MsgBox "Sorry Donar is Under AGE", vbOKOnly
               Else
                Adodc1.Recordset.AddNew
        MsgBox "New Donar added successfully.", vbInformation
    End If
'End If
txtdid.SetFocus
End Sub

Private Sub txtdid_Change()

'If txtdid.Text = "" Then
'MsgBox "field is mandatory", vbCritical
'txtdid.SetFocus
'End If

End Sub

Private Sub txtphno_Change()
If Not IsNumeric(txtphno.Text) Then
    'MsgBox "Value must be numeric", vbCritical
    txtphno.Text = ""
End If
End Sub

Private Sub txtzip_Change()
If Not IsNumeric(txtzip.Text) Then
    'MsgBox "Value must be numeric", vbCritical
    txtzip.Text = ""
End If
End Sub
