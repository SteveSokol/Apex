VERSION 5.00
Begin VB.Form frmRateOfReturn 
   BackColor       =   &H00000000&
   Caption         =   "Cash Flow Analyses"
   ClientHeight    =   5940
   ClientLeft      =   2610
   ClientTop       =   1680
   ClientWidth     =   6690
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
   ScaleHeight     =   5940
   ScaleWidth      =   6690
   Begin VB.Label labRateInTitles 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Date"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   660
      TabIndex        =   29
      Top             =   1980
      Width           =   5355
   End
   Begin VB.Label labPrintScreen 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Print"
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
      Left            =   5400
      TabIndex        =   28
      Top             =   5550
      Width           =   615
   End
   Begin VB.Image imgBackToMenu 
      Height          =   195
      Left            =   60
      Picture         =   "frmRateOfReturn.frx":0000
      Stretch         =   -1  'True
      Top             =   5610
      Width           =   495
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
      Left            =   660
      TabIndex        =   27
      Top             =   5550
      Width           =   675
   End
   Begin VB.Label labRateEquals 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "="
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   5
      Left            =   4140
      TabIndex        =   26
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label labRateEquals 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "="
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   4
      Left            =   4140
      TabIndex        =   25
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label labRateEquals 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "="
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   3
      Left            =   4140
      TabIndex        =   24
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label labRateEquals 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "="
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   2
      Left            =   4140
      TabIndex        =   23
      Top             =   3780
      Width           =   255
   End
   Begin VB.Label labRateEquals 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "="
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   1
      Left            =   4140
      TabIndex        =   22
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label labRateEquals 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "="
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   0
      Left            =   4140
      TabIndex        =   21
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label labRateInTitles 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "* Multiple Rates of Return Possible *"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   5
      Left            =   1260
      TabIndex        =   20
      Top             =   5520
      Visible         =   0   'False
      Width           =   4155
   End
   Begin VB.Label labRateInTitles 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Analysis"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   660
      TabIndex        =   19
      Top             =   1680
      Width           =   5355
   End
   Begin VB.Label labRateInTitles 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Company"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   660
      TabIndex        =   18
      Top             =   1380
      Width           =   5355
   End
   Begin VB.Label labRateInTitles 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Secondary Title"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   660
      TabIndex        =   17
      Top             =   1080
      Width           =   5355
   End
   Begin VB.Label labRateInTitles 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Primary Title"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   660
      TabIndex        =   16
      Top             =   780
      Width           =   5355
   End
   Begin VB.Line linRateOfReturn 
      BorderColor     =   &H00FFFF00&
      X1              =   720
      X2              =   5940
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line linPayBack 
      BorderColor     =   &H00FFFF00&
      X1              =   720
      X2              =   5940
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line linPresentValues 
      BorderColor     =   &H00FFFF00&
      X1              =   720
      X2              =   5940
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line linBottomRight 
      BorderColor     =   &H00FFFF00&
      X1              =   6060
      X2              =   6060
      Y1              =   2460
      Y2              =   5460
   End
   Begin VB.Line linBottomBottom 
      BorderColor     =   &H00FFFF00&
      X1              =   540
      X2              =   6120
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line libBottomTop 
      BorderColor     =   &H00FFFF00&
      X1              =   540
      X2              =   6120
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line linBottomLeft 
      BorderColor     =   &H00FFFF00&
      X1              =   600
      X2              =   600
      Y1              =   2460
      Y2              =   5460
   End
   Begin VB.Line linTopRight 
      BorderColor     =   &H00FFFF00&
      X1              =   6060
      X2              =   6060
      Y1              =   600
      Y2              =   2400
   End
   Begin VB.Line linTopBottom 
      BorderColor     =   &H00FFFF00&
      X1              =   540
      X2              =   6120
      Y1              =   2340
      Y2              =   2340
   End
   Begin VB.Line linTopTop 
      BorderColor     =   &H00FFFF00&
      X1              =   540
      X2              =   6120
      Y1              =   660
      Y2              =   660
   End
   Begin VB.Line linTopLeft 
      BorderColor     =   &H00FFFF00&
      X1              =   600
      X2              =   600
      Y1              =   600
      Y2              =   2400
   End
   Begin VB.Label labRateUnits 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "%"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   5760
      TabIndex        =   15
      Top             =   5040
      Width           =   195
   End
   Begin VB.Label labRateUnits 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "years"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   5400
      TabIndex        =   14
      Top             =   4560
      Width           =   555
   End
   Begin VB.Label labRateValues 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0.00"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   5040
      TabIndex        =   13
      Top             =   5040
      Width           =   675
   End
   Begin VB.Label labRateValues 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0.00"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   4740
      TabIndex        =   12
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label labRateValues 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "$0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   4500
      TabIndex        =   11
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label labRateValues 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "$0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   4500
      TabIndex        =   10
      Top             =   3780
      Width           =   1455
   End
   Begin VB.Label labRateValues 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "$0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   4500
      TabIndex        =   9
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label labRateValues 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "$0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   4500
      TabIndex        =   8
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label labRateTitles 
      BackColor       =   &H00000000&
      Caption         =   "Internal Rate of Return (IROR)"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   780
      TabIndex        =   7
      Top             =   5040
      Width           =   2715
   End
   Begin VB.Label labRateTitles 
      BackColor       =   &H00000000&
      Caption         =   "Payback Period"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   780
      TabIndex        =   6
      Top             =   4560
      Width           =   2715
   End
   Begin VB.Label labRateTitles 
      BackColor       =   &H00000000&
      Caption         =   "@ 20.00% Discount Rate"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   4
      Left            =   1080
      TabIndex        =   5
      Top             =   4080
      Width           =   2715
   End
   Begin VB.Label labRateTitles 
      BackColor       =   &H00000000&
      Caption         =   "@ 15.00% Discount Rate"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   4
      Top             =   3780
      Width           =   2715
   End
   Begin VB.Label labRateTitles 
      BackColor       =   &H00000000&
      Caption         =   "@ 10.00% Discount Rate"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   3
      Top             =   3480
      Width           =   2715
   End
   Begin VB.Label labRateTitles 
      BackColor       =   &H00000000&
      Caption         =   "Present Values:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   780
      TabIndex        =   2
      Top             =   3120
      Width           =   2715
   End
   Begin VB.Label labRateTitles 
      BackColor       =   &H00000000&
      Caption         =   "Net Sum of Cash Flows"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   780
      TabIndex        =   1
      Top             =   2640
      Width           =   2715
   End
   Begin VB.Label labRateHeading 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cash Flow Summary"
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
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   3075
   End
End
Attribute VB_Name = "frmRateOfReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()

Dim i As Integer

ShowMenu = True

For i = 0 To 4
  labRateInTitles(i).Caption = Titler(i)
Next i

For i = 2 To 4
  labRateTitles(i).Caption = "@ " & LTrim(RTrim(Str(Sets(22 + i)))) & "% Discount Rate"
Next i

labRateValues(0).Caption = RTrim(Str(Pv0))
labRateValues(1).Caption = RTrim(Str(Pv1))
labRateValues(2).Caption = RTrim(Str(Pv2))
labRateValues(3).Caption = RTrim(Str(Pv3))
labRateValues(4).Caption = RTrim(Str(Pb))
labRateValues(5).Caption = RTrim(Str(Rot * 100))

For i = 0 To 3
  labRateValues(i).Caption = Format(labRateValues(i).Caption, "$##,###,###,###")
Next i

For i = 4 To 5
  labRateValues(i).Caption = Format(labRateValues(i).Caption, "##0.00")
Next i

labRateInTitles(5).Visible = False
If BadRor = 2 Then labRateInTitles(5).Visible = True

End Sub

Private Sub Form_Deactivate()
 
  If ShowMenu = True Then
    frmRateOfReturn.Hide
    Call InputMenuAccess(2)
  End If
  
End Sub

Private Sub Form_Load()

If FullScreen = False Then
  frmRateOfReturn.Top = (Screen.Height - (frmRateOfReturn.Height + 350)) / 2
  frmRateOfReturn.Left = (Screen.Width - frmRateOfReturn.Width) / 2
Else
  frmRateOfReturn.Top = 0
  frmRateOfReturn.Left = 0
  frmRateOfReturn.WindowState = 2
End If

If frmRateOfReturn.Top < 0 Then frmRateOfReturn.Top = 0
If frmRateOfReturn.Left < 0 Then frmRateOfReturn.Left = 0

End Sub

Private Sub Form_LostFocus()
  If ShowMenu = True Then
    frmRateOfReturn.Hide
    Call InputMenuAccess(2)
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

  frmRateOfReturn.Hide
  If ShowMenu = True Then Call InputMenuAccess(2)

End Sub

Private Sub imgBackToMenu_Click()

  frmRateOfReturn.Hide
  Call InputMenuAccess(2)

End Sub


Private Sub labBackToMenu_Click()

  frmRateOfReturn.Hide
  Call InputMenuAccess(2)

End Sub


Private Sub labPrintScreen_Click()

  job = 9
  ShowMenu = False
  Call printstuffout(job)

End Sub


