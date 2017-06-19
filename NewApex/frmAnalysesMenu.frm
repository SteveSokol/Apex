VERSION 5.00
Begin VB.Form frmAnalysesMenu 
   BackColor       =   &H00000000&
   Caption         =   "Analyses Menu"
   ClientHeight    =   5265
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   3510
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
   ScaleHeight     =   5265
   ScaleWidth      =   3510
   Begin VB.Image imgOnToSchedules 
      Height          =   195
      Left            =   2940
      Picture         =   "frmAnalysesMenu.frx":0000
      Stretch         =   -1  'True
      Top             =   4980
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBackToInput 
      Height          =   195
      Left            =   60
      Picture         =   "frmAnalysesMenu.frx":0442
      Stretch         =   -1  'True
      Top             =   4980
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label labAnalysesHeading 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Analyses"
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
      TabIndex        =   5
      Top             =   480
      Width           =   1590
   End
   Begin VB.Label labAnalysesTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "&Risk"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   1020
      TabIndex        =   4
      Top             =   2580
      Width           =   390
   End
   Begin VB.Label labAnalysesTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "&Sensitivity"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   1020
      TabIndex        =   3
      Top             =   2220
      Width           =   885
   End
   Begin VB.Label labAnalysesTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "&Break Even"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   1020
      TabIndex        =   2
      Top             =   1860
      Width           =   990
   End
   Begin VB.Label labAnalysesTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "&Cash Flow"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   1020
      TabIndex        =   1
      Top             =   1500
      Width           =   915
   End
   Begin VB.Label labAnalysesTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "&Parameters"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   1020
      TabIndex        =   0
      Top             =   1140
      Width           =   1005
   End
End
Attribute VB_Name = "frmAnalysesMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  
  Dim X As Integer
  Dim temphigh As Integer
  Dim tempwide As Integer

  frmAnalysesMenu.Height = Screen.Height * 0.6
  frmAnalysesMenu.Width = Screen.Width * 0.3
  frmAnalysesMenu.Top = Screen.Height * 0.14
  frmAnalysesMenu.Left = Screen.Width * 0.25
  
  tempwide = frmAnalysesMenu.ScaleWidth
  temphigh = frmAnalysesMenu.ScaleHeight
  
  For X = 0 To 4
    labAnalysesTitle(X).Top = temphigh * (0.1813 + (X * 0.068))
    labAnalysesTitle(X).Left = tempwide * 0.1488
  Next X
  
  labAnalysesHeading.Top = temphigh * 0.0567
  labAnalysesHeading.Left = tempwide * 0.0661
  
  imgBackToInput.Top = temphigh * 0.9518
  imgBackToInput.Left = tempwide * 0.0113
  imgBackToInput.Width = tempwide * 0.141
  
  imgOnToSchedules.Top = temphigh * 0.9518
  imgOnToSchedules.Left = tempwide * 0.843
  imgOnToSchedules.Width = tempwide * 0.141

End Sub


Private Sub imgOnToSchedule_Click()

End Sub

Private Sub imgBackToInput_Click()
  imgBackToInput.Visible = False
  imgOnToSchedules.Visible = False
  frmInputMenu.imgOnToAnalyses.Visible = True
  frmInputMenu.Show
End Sub


Private Sub imgOnToSchedules_Click()
  imgBackToInput.Visible = False
  imgOnToSchedules.Visible = False
  frmSchedulesMenu.imgBackToAnalyses.Visible = True
  frmSchedulesMenu.imgOnToUtilities.Visible = True
  frmSchedulesMenu.Show
End Sub


Private Sub labAnalysesTitle_Click(Index As Integer)
  
  Select Case Index
    Case 0
      Call InputMenuOutCalls
      ParamSet = True
      frmParameters.Show
    Case 1
      Call InputMenuOutCalls
      Be = "off"
      Call cflow5(1, 0)
      Call rateofreturn
      frmRateOfReturn.Show
    Case 2
      If nTag > 0 And ParamSet = True Then
        Call InputMenuOutCalls
        frmBreakEven.Show
      Else
        WarnNumber = 1
        frmWarnTheUser.Show
      End If
    Case 3
      If nTag > 0 And ParamSet = True Then
        Call InputMenuOutCalls
          frmSensitivity.Show
      Else
        WarnNumber = 2
        frmWarnTheUser.Show
      End If
    Case 4
      If nTag > 0 And ParamSet = True Then
        Call InputMenuOutCalls
        DontRisk = False
        frmRisk.Show
      Else
        WarnNumber = 3
        frmWarnTheUser.Show
      End If
   End Select

End Sub


