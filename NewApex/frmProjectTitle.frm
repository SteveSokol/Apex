VERSION 5.00
Begin VB.Form frmProjectTitle 
   BackColor       =   &H00000000&
   Caption         =   "Project Title"
   ClientHeight    =   6150
   ClientLeft      =   870
   ClientTop       =   870
   ClientWidth     =   9150
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6150
   ScaleWidth      =   9150
   Begin VB.TextBox txtProjectTitles 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   3060
      TabIndex        =   4
      Top             =   3900
      Width           =   4095
   End
   Begin VB.TextBox txtProjectTitles 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   3060
      TabIndex        =   3
      Top             =   3300
      Width           =   5475
   End
   Begin VB.TextBox txtProjectTitles 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   3060
      TabIndex        =   2
      Top             =   2700
      Width           =   5475
   End
   Begin VB.TextBox txtProjectTitles 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   3060
      TabIndex        =   1
      Top             =   2100
      Width           =   5475
   End
   Begin VB.TextBox txtProjectTitles 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   3060
      TabIndex        =   0
      Top             =   1500
      Width           =   5475
   End
   Begin VB.Label labTitleHelp 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8340
      TabIndex        =   12
      Top             =   5700
      Width           =   675
   End
   Begin VB.Label labBackToMenu 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   600
      TabIndex        =   11
      Top             =   5640
      Width           =   675
   End
   Begin VB.Label labProjectHeading 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Project Title"
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
      Left            =   630
      TabIndex        =   10
      Top             =   600
      Width           =   1965
   End
   Begin VB.Label labProjectTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   840
      TabIndex        =   9
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label labProjectTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Analysis"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   840
      TabIndex        =   8
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label labProjectTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   840
      TabIndex        =   7
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label labProjectTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   840
      TabIndex        =   6
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label labProjectTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Project Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   840
      TabIndex        =   5
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Image imgBackToMenu 
      Height          =   195
      Left            =   60
      Picture         =   "frmProjectTitle.frx":0000
      Stretch         =   -1  'True
      Top             =   5700
      Width           =   495
   End
End
Attribute VB_Name = "frmProjectTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()

  Dim i As Integer
  
  If IsHelpOn = True Then
    txtProjectTitles(LastCell).SetFocus
    IsHelpOn = False
  Else
    If txtProjectTitles(4).Text = "" Then
      txtProjectTitles(4).Text = Format(Date, "Long Date")
      Titler(4) = txtProjectTitles(4).Text
    End If
    If PageChange(0) = True Then
      For i = 0 To 4
        txtProjectTitles(i).Text = Titler(i)
      Next i
    End If
    LastCell = 0
    txtProjectTitles(0).SetFocus
  End If
  
End Sub

Private Sub Form_Load()

Dim X As Integer
Dim temphigh As Currency
Dim tempwide As Currency

If FullScreen = False Then
  frmProjectTitle.Top = (Screen.Height - (frmProjectTitle.Height + 350)) / 2
  frmProjectTitle.Left = (Screen.Width - frmProjectTitle.Width) / 2
Else
  frmProjectTitle.Top = 0
  frmProjectTitle.Left = 0
  frmProjectTitle.WindowState = 2
End If

If frmProjectTitle.Top < 0 Then frmProjectTitle.Top = 0
If frmProjectTitle.Left < 0 Then frmProjectTitle.Left = 0

temphigh = frmProjectTitle.ScaleHeight
tempwide = frmProjectTitle.ScaleWidth

For X = 0 To 4
  labProjectTitles(X).Top = (temphigh * (0.105)) + (temphigh * (0.15 + (X / 10)))
  labProjectTitles(X).Left = tempwide * 0.0748
  labProjectTitles(X).Width = tempwide * 0.2318
  txtProjectTitles(X).Top = (temphigh * (0.1)) + (temphigh * (0.15 + (X / 10)))
  txtProjectTitles(X).Left = tempwide * 0.32
  If X = 4 Then
    txtProjectTitles(X).Width = tempwide * 0.4619
  Else
    txtProjectTitles(X).Width = tempwide * 0.6176
  End If
Next X

  labProjectHeading.Left = tempwide * 0.0194
  labProjectHeading.Top = temphigh * 0.0334

  labBackToMenu.Left = tempwide * 0.08
  labBackToMenu.Top = temphigh * 0.9425

  imgBackToMenu.Left = tempwide * 0.012
  imgBackToMenu.Top = temphigh * 0.9496

  labTitleHelp.Top = temphigh * 0.9425
  labTitleHelp.Left = tempwide * 0.9115

End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  frmProjectTitle.Hide
  Call InputMenuAccess(1)

End Sub

Private Sub imgBackToMenu_Click()
  
  frmProjectTitle.Hide
  Call InputMenuAccess(1)
  
End Sub


Private Sub labBackToMenu_Click()
  
  frmProjectTitle.Hide
  Call InputMenuAccess(1)

End Sub

Private Sub labTitleHelp_Click()

Dim begin As Integer
Dim sendindex As Integer

begin = 0

sendindex = LastCell + 1

If LastCell = 3 Then
  sendindex = 52
ElseIf LastCell = 4 Then
  sendindex = LastCell
End If

WhichScreen = 0

Call frmApexHelp.gethelptext(sendindex, begin)
frmApexHelp.Show

End Sub

Private Sub txtProjectTitles_Change(Index As Integer)
  
  PageChange(0) = True
  
  Titler(Index) = txtProjectTitles(Index).Text
  
End Sub


Private Sub txtProjectTitles_GotFocus(Index As Integer)

LastCell = Index

End Sub


