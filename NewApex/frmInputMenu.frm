VERSION 5.00
Begin VB.Form frmInputMenu 
   BackColor       =   &H00000000&
   Caption         =   "Data Entry Menu"
   ClientHeight    =   5295
   ClientLeft      =   2235
   ClientTop       =   1545
   ClientWidth     =   3630
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5295
   ScaleWidth      =   3630
   Begin VB.Label labInputTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "&Smelting and Refining"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   540
      TabIndex        =   9
      Top             =   2400
      Width           =   1890
   End
   Begin VB.Image imgOnToAnalyses 
      Height          =   195
      Left            =   3060
      Picture         =   "frmInputMenu.frx":0000
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label labInputHeading 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Data Entry"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   210
      TabIndex        =   8
      Top             =   300
      Width           =   1785
   End
   Begin VB.Label labInputTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "&Taxes and Escalation"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   8
      Left            =   540
      TabIndex        =   7
      Top             =   3840
      Width           =   1875
   End
   Begin VB.Label labInputTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "&Royalties"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   7
      Left            =   540
      TabIndex        =   6
      Top             =   3480
      Width           =   810
   End
   Begin VB.Label labInputTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "&Financing"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   6
      Left            =   540
      TabIndex        =   5
      Top             =   3120
      Width           =   840
   End
   Begin VB.Label labInputTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "&Capital Costs"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   5
      Left            =   540
      TabIndex        =   4
      Top             =   2760
      Width           =   1170
   End
   Begin VB.Label labInputTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Proce&ssing"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   540
      TabIndex        =   3
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label labInputTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "&Mining"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   540
      TabIndex        =   2
      Top             =   1680
      Width           =   570
   End
   Begin VB.Label labInputTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Commodities and &Grades"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   540
      TabIndex        =   1
      Top             =   1320
      Width           =   2205
   End
   Begin VB.Label labInputTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "&Project Title"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   540
      TabIndex        =   0
      Top             =   960
      Width           =   1035
   End
End
Attribute VB_Name = "frmInputMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

Dim X As Integer
Dim temphigh As Integer
Dim tempwide As Integer
Dim findintro As String

Call loaditup

frmUtilitiesMenu.Show
frmSchedulesMenu.Show
frmAnalysesMenu.Show

frmInputMenu.Height = Screen.Height * 0.6
frmInputMenu.Width = Screen.Width * 0.3
frmInputMenu.Top = Screen.Height * 0.05
frmInputMenu.Left = Screen.Width * 0.05

tempwide = frmInputMenu.ScaleWidth
temphigh = frmInputMenu.ScaleHeight
  
For X = 0 To 8
  labInputTitle(X).Top = temphigh * (0.1813 + (X * 0.068))
  labInputTitle(X).Left = tempwide * 0.1488
Next X
  
labInputHeading.Top = temphigh * 0.0567
labInputHeading.Left = tempwide * 0.0661
 
imgOnToAnalyses.Top = temphigh * 0.9518
imgOnToAnalyses.Left = tempwide * 0.8345
imgOnToAnalyses.Width = tempwide * 0.141
  
frmInputMenu.Show

End Sub
Private Sub labProjectTitle_Click()

End Sub


Private Sub labInputAccess_Click(Index As Integer)

End Sub


Private Sub imgOnToAnalyses_Click()

  imgOnToAnalyses.Visible = False
  frmAnalysesMenu.imgOnToSchedules.Visible = True
  frmAnalysesMenu.imgBackToInput.Visible = True
  frmAnalysesMenu.Show

End Sub
Private Sub labInputTitle_Click(Index As Integer)
  
  Call InputMenuOutCalls
  
  Select Case Index
    Case 0
      frmProjectTitle.Show
    Case 1
      frmCommodities.Show
    Case 2
      frmMining.Show
     Case 3
      frmProcessingCost.Show
    Case 4
      frmSmelting.Show
    Case 5
      frmCapital.Show
    Case 6
      frmFinancing.Show
    Case 7
      frmRoyalties.Show
    Case 8
      frmTaxes.Show
  End Select
  
End Sub



Public Sub getsetup()

End Sub
