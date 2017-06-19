VERSION 5.00
Begin VB.Form frmSchedulesMenu 
   BackColor       =   &H00000000&
   Caption         =   "Schedules Menu"
   ClientHeight    =   5340
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   3555
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
   ScaleHeight     =   5340
   ScaleWidth      =   3555
   Begin VB.Label labSchedulesTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Capital &Deductions"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   1140
      TabIndex        =   5
      Top             =   2580
      Width           =   1635
   End
   Begin VB.Image imgBackToAnalyses 
      Height          =   195
      Left            =   60
      Picture         =   "frmSchedulesMenu.frx":0000
      Stretch         =   -1  'True
      Top             =   5100
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgOnToUtilities 
      Height          =   195
      Left            =   3000
      Picture         =   "frmSchedulesMenu.frx":0442
      Stretch         =   -1  'True
      Top             =   5100
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label labSchedulesHeading 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Schedules"
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
      Left            =   720
      TabIndex        =   4
      Top             =   480
      Width           =   1755
   End
   Begin VB.Label labSchedulesTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Com&posite"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   1140
      TabIndex        =   3
      Top             =   2220
      Width           =   930
   End
   Begin VB.Label labSchedulesTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Val&ues and Costs"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   1140
      TabIndex        =   2
      Top             =   1860
      Width           =   1545
   End
   Begin VB.Label labSchedulesTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "&Royalties"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   1140
      TabIndex        =   1
      Top             =   1500
      Width           =   810
   End
   Begin VB.Label labSchedulesTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "&Cash Flow"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   1140
      TabIndex        =   0
      Top             =   1140
      Width           =   915
   End
End
Attribute VB_Name = "frmSchedulesMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  
  Dim X As Integer
  Dim temphigh As Integer
  Dim tempwide As Integer
  
  frmSchedulesMenu.Height = Screen.Height * 0.6
  frmSchedulesMenu.Width = Screen.Width * 0.3
  frmSchedulesMenu.Top = Screen.Height * 0.23
  frmSchedulesMenu.Left = Screen.Width * 0.45
  
  temphigh = frmSchedulesMenu.ScaleHeight
  tempwide = frmSchedulesMenu.ScaleWidth
  
  For X = 0 To 4
    labSchedulesTitle(X).Top = temphigh * (0.1818 + (X * 0.068))
    labSchedulesTitle(X).Left = tempwide * 0.1488
  Next X
  
  labSchedulesHeading.Top = temphigh * 0.0567
  labSchedulesHeading.Left = tempwide * 0.0661

  imgBackToAnalyses.Top = temphigh * 0.9518
  imgBackToAnalyses.Left = tempwide * 0.0113
  imgBackToAnalyses.Width = tempwide * 0.141
  
  imgOnToUtilities.Top = temphigh * 0.9518
  imgOnToUtilities.Left = tempwide * 0.843
  imgOnToUtilities.Width = tempwide * 0.141

End Sub

Private Sub imgBackToAnalyses_Click()
  imgBackToAnalyses.Visible = False
  imgOnToUtilities.Visible = False
  frmAnalysesMenu.imgBackToInput.Visible = True
  frmAnalysesMenu.imgOnToSchedules.Visible = True
  frmAnalysesMenu.Show
End Sub

Private Sub imgOnToUtilities_Click()
  imgBackToAnalyses.Visible = False
  imgOnToUtilities.Visible = False
  frmUtilitiesMenu.imgBackToSchedules.Visible = True
  frmUtilitiesMenu.Show
End Sub


Private Sub labSchedulesTitle_Click(Index As Integer)

Call InputMenuOutCalls
job = Index + 1
frmCashFlow.Show

End Sub


