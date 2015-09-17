VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "About"
   ClientHeight    =   3780
   ClientLeft      =   6255
   ClientTop       =   825
   ClientWidth     =   6840
   LinkTopic       =   "Form4"
   ScaleHeight     =   3780
   ScaleWidth      =   6840
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H80000002&
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4560
      MaskColor       =   &H00FF0000&
      TabIndex        =   0
      Top             =   2160
      Width           =   1260
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "Application Title:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5565
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H80000004&
      Caption         =   "Application Description: This programme is designed to keep record of Blood Bank in Hospitals ranging from Blood Donor Record."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   1050
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   5565
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   120
      X2              =   5880
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3045
      TabIndex        =   1
      Top             =   480
      Width           =   1080
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "Copyright :This product is design and developed by Vrushal Raut 2013."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   5895
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
 Me.Caption = "About " & App.Title
    lblVersion.Caption = "Windows Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
End Sub

