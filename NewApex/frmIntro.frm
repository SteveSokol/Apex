VERSION 5.00
Begin VB.Form frmIntro 
   BackColor       =   &H00000000&
   Caption         =   "Apex For Windows"
   ClientHeight    =   4020
   ClientLeft      =   2430
   ClientTop       =   2355
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4020
   ScaleWidth      =   6375
   Begin VB.CommandButton comBeginProgram 
      Caption         =   "&Begin"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   795
   End
   Begin VB.Label labTitleThree 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "InfoMine USA, Inc./Aventurine Engineering, Inc."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2520
      TabIndex        =   4
      Top             =   3660
      Width           =   3795
   End
   Begin VB.Label labTitleTwo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 2015"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2520
      TabIndex        =   3
      Top             =   3360
      Width           =   3795
   End
   Begin VB.Label labTitleOne 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Version 3.15"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2520
      TabIndex        =   2
      Top             =   3060
      Width           =   3795
   End
   Begin VB.Label labIntroTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Apex For Windows"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Image imgIntro 
      Height          =   3915
      Left            =   60
      Picture         =   "frmIntro.frx":0000
      Stretch         =   -1  'True
      Top             =   60
      Width           =   6270
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub comBeginProgram_Click()

If UserYes = 0 Then End

frmIntro.Hide

End Sub

Private Sub Form_Load()

  frmIntro.Top = Screen.Height * 0.14
  frmIntro.Left = Screen.Width * 0.15
  frmIntro.Height = Screen.Height * 0.69
  frmIntro.Width = Screen.Width * 0.7

  labIntroTitle.Top = frmIntro.Height * 0.0271
  labIntroTitle.Left = frmIntro.ScaleWidth - 4720

  imgIntro.Height = frmIntro.ScaleHeight * 0.9782
  imgIntro.Width = frmIntro.ScaleWidth * 0.9835

  labTitleOne.Left = frmIntro.ScaleWidth - 3980
  labTitleTwo.Left = frmIntro.ScaleWidth - 3980
  labTitleThree.Left = frmIntro.ScaleWidth - 3980

  labTitleOne.Top = frmIntro.ScaleHeight - 900
  labTitleTwo.Top = frmIntro.ScaleHeight - 660
  labTitleThree.Top = frmIntro.ScaleHeight - 420

  comBeginProgram.Top = frmIntro.ScaleHeight - 480
  comBeginProgram.Left = frmIntro.ScaleWidth * 0.0128

End Sub

