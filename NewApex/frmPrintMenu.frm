VERSION 5.00
Begin VB.Form frmPrintMenu 
   BackColor       =   &H00000000&
   Caption         =   "Print Menu"
   ClientHeight    =   3735
   ClientLeft      =   2625
   ClientTop       =   2010
   ClientWidth     =   3315
   FillColor       =   &H00404040&
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3735
   ScaleWidth      =   3315
   Begin VB.CommandButton comGetOnWithIt 
      Caption         =   "Print"
      Height          =   315
      Left            =   1260
      TabIndex        =   9
      Top             =   3300
      Width           =   615
   End
   Begin VB.HScrollBar hscSets 
      Height          =   195
      Left            =   1260
      Max             =   25
      Min             =   1
      TabIndex        =   7
      Top             =   2940
      Value           =   1
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.OptionButton optPrintMenu 
      BackColor       =   &H00000000&
      Caption         =   "Production Set"
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
      Height          =   255
      Index           =   4
      Left            =   780
      TabIndex        =   4
      Top             =   2160
      Width           =   1875
   End
   Begin VB.OptionButton optPrintMenu 
      BackColor       =   &H00000000&
      Caption         =   "Royalty Data"
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
      Height          =   255
      Index           =   3
      Left            =   780
      TabIndex        =   3
      Top             =   1800
      Width           =   1455
   End
   Begin VB.OptionButton optPrintMenu 
      BackColor       =   &H00000000&
      Caption         =   "Capital Data"
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
      Height          =   255
      Index           =   2
      Left            =   780
      TabIndex        =   2
      Top             =   1440
      Width           =   1395
   End
   Begin VB.OptionButton optPrintMenu 
      BackColor       =   &H00000000&
      Caption         =   "Production Data"
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
      Height          =   255
      Index           =   1
      Left            =   780
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
   End
   Begin VB.OptionButton optPrintMenu 
      BackColor       =   &H00000000&
      Caption         =   "All Data"
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
      Height          =   255
      Index           =   0
      Left            =   780
      TabIndex        =   0
      Top             =   720
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.Label labSetNumber 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1740
      TabIndex        =   8
      Top             =   2940
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label labWhichSet 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Set Number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   2640
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label labPrintHeading 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Print Menu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   210
      TabIndex        =   5
      Top             =   180
      Width           =   1635
   End
End
Attribute VB_Name = "frmPrintMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label1_Click()

End Sub


Private Sub comGetOnWithIt_Click()
        
    If optPrintMenu(0).Value = True Then
      job = 20
    ElseIf optPrintMenu(1).Value = True Then
      job = 21
    ElseIf optPrintMenu(2).Value = True Then
      job = 22
    ElseIf optPrintMenu(3).Value = True Then
      job = 23
    ElseIf optPrintMenu(4).Value = True Then
      job = 24
    End If
    
    Call printstuffout(job)

    frmPrintMenu.Hide
    
End Sub

Private Sub hscSets_Change()

If DoNotChange = True Then Exit Sub

labSetNumber.Caption = LTrim(RTrim(Str(hscSets.Value)))

End Sub

Private Sub optPrintMenu_Click(Index As Integer)

If Index = 4 Then
  hscSets.Visible = True
  labSetNumber.Visible = True
  labWhichSet.Visible = True
Else
  hscSets.Visible = False
  labSetNumber.Visible = False
  labWhichSet.Visible = False
End If

End Sub


