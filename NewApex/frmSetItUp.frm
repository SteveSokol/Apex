VERSION 5.00
Begin VB.Form frmSetItUp 
   BackColor       =   &H00FF0000&
   Caption         =   "Apex Set Up"
   ClientHeight    =   4125
   ClientLeft      =   4320
   ClientTop       =   2070
   ClientWidth     =   6675
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4125
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkFullScreen 
      BackColor       =   &H00FF0000&
      Caption         =   "Operate in Full Screen Mode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1560
      TabIndex        =   7
      Top             =   3000
      Width           =   3555
   End
   Begin VB.CommandButton comSetItUp 
      Caption         =   "Continue"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   3540
      Width           =   1200
   End
   Begin VB.TextBox txtUserName 
      Height          =   360
      Index           =   0
      Left            =   2340
      TabIndex        =   0
      Top             =   1800
      Width           =   4215
   End
   Begin VB.TextBox txtUserName 
      Height          =   360
      Index           =   1
      Left            =   2340
      TabIndex        =   1
      Top             =   2400
      Width           =   4215
   End
   Begin VB.Label labSetItUp 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      Caption         =   "Your Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   2115
   End
   Begin VB.Label labSetItUp 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      Caption         =   "Serial Number:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   2115
   End
   Begin VB.Label labSetItUp 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Please verify your license agreement by entering the program serial number and your name."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   1
      Left            =   180
      TabIndex        =   4
      Top             =   960
      Width           =   6435
   End
   Begin VB.Label labSetItUp 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Welcome to APEX for Windows"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   6435
   End
End
Attribute VB_Name = "frmSetItUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkFullScreen_Click()

If chkFullScreen.Value = 1 Then
  FullScreen = True
Else
  FullScreen = False
End If

End Sub

Private Sub comSetItUp_Click()

If LTrim(RTrim(LCase(txtUserName(0)))) = UserName(0) Then
  X = 0
  While X < 4
    If LTrim(RTrim(LCase(txtUserName(1)))) = UserCompany(X) Then
      UserYes = 2
      X = 4
    Else
      UserYes = 0
      X = X + 1
    End If
  Wend
ElseIf LTrim(RTrim(LCase(txtUserName(0)))) = UserName(1) Then
  X = 0
  While X < 4
    If LTrim(RTrim(LCase(txtUserName(1)))) = UserCompany(X) Then
      UserYes = 2
      X = 4
    Else
      UserYes = 0
      X = X + 1
    End If
  Wend
ElseIf LTrim(RTrim(LCase(txtUserName(0)))) = UserName(2) Then
  X = 0
  While X < 4
    If LTrim(RTrim(LCase(txtUserName(1)))) = UserCompany(X) Then
      UserYes = 2
      X = 4
    Else
      UserYes = 0
      X = X + 1
    End If
  Wend
ElseIf LTrim(RTrim(LCase(txtUserName(0)))) = UserName(3) Then
  X = 0
  While X < 4
    If LTrim(RTrim(LCase(txtUserName(1)))) = UserCompany(X) Then
      UserYes = 2
      X = 4
    Else
      UserYes = 0
      X = X + 1
    End If
  Wend
End If

If UserYes = 0 Then
  txtUserName(0).Text = ""
  txtUserName(1).Text = ""
  labSetItUp(1).ForeColor = &HFF00&
  labSetItUp(1).Caption = "Unable to verify license agreement.  Please try again or contact Aventurine Engineering, Inc."
Else
  frmSetItUp.Hide
  frmInputMenu.Show
  frmIntro.Show
End If

End Sub

Private Sub Form_Load()

frmSetItUp.Top = (Screen.Height / 2) - (4545 / 2)
frmSetItUp.Left = (Screen.Width / 2) - (6690 / 2)

UserName(0) = ""
UserName(1) = "951139"
UserName(2) = "951039"
UserName(3) = "8000"

UserCompany(0) = "steve"
UserCompany(1) = "manami ikeda"
UserCompany(2) = "ikeda"
UserCompany(3) = "manami"

frmSetItUp.Show

End Sub

