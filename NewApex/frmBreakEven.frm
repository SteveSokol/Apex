VERSION 5.00
Begin VB.Form frmBreakEven 
   BackColor       =   &H00000000&
   Caption         =   "Break Even"
   ClientHeight    =   6420
   ClientLeft      =   2190
   ClientTop       =   1320
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
   ScaleHeight     =   6420
   ScaleWidth      =   9150
   Begin VB.VScrollBar vscBreakEven 
      Height          =   2475
      Left            =   4680
      TabIndex        =   52
      Top             =   3480
      Width           =   195
   End
   Begin VB.CheckBox chkDiscount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Use Discount Rate when Determining Break-Even Value"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   600
      Width           =   5055
   End
   Begin VB.HScrollBar hscTagNumber 
      Height          =   195
      Left            =   1140
      Max             =   50
      Min             =   1
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1020
      Value           =   1
      Width           =   375
   End
   Begin VB.Label labBreakTitles 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Working - Please Wait"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Index           =   5
      Left            =   4920
      TabIndex        =   51
      Top             =   2040
      Width           =   2595
   End
   Begin VB.Label labUnit 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   10
      Left            =   6840
      TabIndex        =   50
      Top             =   5700
      Width           =   1575
   End
   Begin VB.Label labUnit 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   9
      Left            =   6840
      TabIndex        =   49
      Top             =   5460
      Width           =   1575
   End
   Begin VB.Label labValue 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   10
      Left            =   5040
      TabIndex        =   48
      Top             =   5700
      Width           =   1695
   End
   Begin VB.Label labValue 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   9
      Left            =   5040
      TabIndex        =   47
      Top             =   5460
      Width           =   1695
   End
   Begin VB.Label labItem 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   10
      Left            =   840
      TabIndex        =   46
      Top             =   5700
      Width           =   3675
   End
   Begin VB.Label labItem 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   9
      Left            =   840
      TabIndex        =   45
      Top             =   5460
      Width           =   3675
   End
   Begin VB.Line linLast 
      BorderColor     =   &H00FFFF00&
      X1              =   480
      X2              =   8700
      Y1              =   3420
      Y2              =   3420
   End
   Begin VB.Label labUnit 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   8
      Left            =   6840
      TabIndex        =   44
      Top             =   5220
      Width           =   1575
   End
   Begin VB.Label labUnit 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   7
      Left            =   6840
      TabIndex        =   43
      Top             =   4980
      Width           =   1575
   End
   Begin VB.Label labUnit 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   6
      Left            =   6840
      TabIndex        =   42
      Top             =   4740
      Width           =   1575
   End
   Begin VB.Label labUnit 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   5
      Left            =   6840
      TabIndex        =   41
      Top             =   4500
      Width           =   1575
   End
   Begin VB.Label labUnit 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   4
      Left            =   6840
      TabIndex        =   40
      Top             =   4260
      Width           =   1575
   End
   Begin VB.Label labUnit 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   3
      Left            =   6840
      TabIndex        =   39
      Top             =   4020
      Width           =   1575
   End
   Begin VB.Label labUnit 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   2
      Left            =   6840
      TabIndex        =   38
      Top             =   3780
      Width           =   1575
   End
   Begin VB.Label labUnit 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   1
      Left            =   6840
      TabIndex        =   37
      Top             =   3540
      Width           =   1575
   End
   Begin VB.Label labUnit 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   6840
      TabIndex        =   36
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label labValue 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   8
      Left            =   5040
      TabIndex        =   35
      Top             =   5220
      Width           =   1695
   End
   Begin VB.Label labValue 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   7
      Left            =   5040
      TabIndex        =   34
      Top             =   4980
      Width           =   1695
   End
   Begin VB.Label labValue 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   6
      Left            =   5040
      TabIndex        =   33
      Top             =   4740
      Width           =   1695
   End
   Begin VB.Label labValue 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   5
      Left            =   5040
      TabIndex        =   32
      Top             =   4500
      Width           =   1695
   End
   Begin VB.Label labValue 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   4
      Left            =   5040
      TabIndex        =   31
      Top             =   4260
      Width           =   1695
   End
   Begin VB.Label labValue 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   3
      Left            =   5040
      TabIndex        =   30
      Top             =   4020
      Width           =   1695
   End
   Begin VB.Label labValue 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   2
      Left            =   5040
      TabIndex        =   29
      Top             =   3780
      Width           =   1695
   End
   Begin VB.Label labValue 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   1
      Left            =   5040
      TabIndex        =   28
      Top             =   3540
      Width           =   1695
   End
   Begin VB.Label labValue 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   5040
      TabIndex        =   27
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label labItem 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   8
      Left            =   840
      TabIndex        =   26
      Top             =   5220
      Width           =   3675
   End
   Begin VB.Label labItem 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   7
      Left            =   840
      TabIndex        =   25
      Top             =   4980
      Width           =   3675
   End
   Begin VB.Label labItem 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   6
      Left            =   840
      TabIndex        =   24
      Top             =   4740
      Width           =   3675
   End
   Begin VB.Label labItem 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   5
      Left            =   840
      TabIndex        =   23
      Top             =   4500
      Width           =   3675
   End
   Begin VB.Label labItem 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   4
      Left            =   840
      TabIndex        =   22
      Top             =   4260
      Width           =   3675
   End
   Begin VB.Label labItem 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   3
      Left            =   840
      TabIndex        =   21
      Top             =   4020
      Width           =   3675
   End
   Begin VB.Label labItem 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   20
      Top             =   3780
      Width           =   3675
   End
   Begin VB.Label labItem 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   19
      Top             =   3540
      Width           =   3675
   End
   Begin VB.Label labItem 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   18
      Top             =   3120
      Width           =   3675
   End
   Begin VB.Label labSetNumber 
      Alignment       =   2  'Center
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
      Left            =   1320
      TabIndex        =   16
      Top             =   1620
      Width           =   255
   End
   Begin VB.Label labSetTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Set Number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   900
      TabIndex        =   15
      Top             =   1380
      Width           =   1095
   End
   Begin VB.Label labTagNumber 
      Alignment       =   2  'Center
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
      Left            =   1620
      TabIndex        =   14
      Top             =   1020
      Width           =   255
   End
   Begin VB.Label labTagTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Tag Number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   900
      TabIndex        =   12
      Top             =   720
      Width           =   1095
   End
   Begin VB.Line linTagRight 
      BorderColor     =   &H00FFFF00&
      X1              =   2220
      X2              =   2220
      Y1              =   600
      Y2              =   1980
   End
   Begin VB.Line linTagLeft 
      BorderColor     =   &H00FFFF00&
      X1              =   660
      X2              =   660
      Y1              =   600
      Y2              =   1980
   End
   Begin VB.Line linTagBottom 
      BorderColor     =   &H00FFFF00&
      X1              =   600
      X2              =   2280
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line linTagTop 
      BorderColor     =   &H00FFFF00&
      X1              =   600
      X2              =   2280
      Y1              =   660
      Y2              =   660
   End
   Begin VB.Label labBreakOut 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "0.00 percent"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   7500
      TabIndex        =   11
      Top             =   1020
      Width           =   1020
   End
   Begin VB.Label labBreakOut 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "0.00 percent"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   7500
      TabIndex        =   10
      Top             =   1440
      Width           =   1020
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
      Left            =   8040
      TabIndex        =   9
      Top             =   6075
      Width           =   615
   End
   Begin VB.Image imgBackToMenu 
      Height          =   195
      Left            =   60
      Picture         =   "frmBreakEven.frx":0000
      Stretch         =   -1  'True
      Top             =   6120
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
      TabIndex        =   8
      Top             =   6075
      Width           =   675
   End
   Begin VB.Label labBreakMisc 
      BackColor       =   &H00000000&
      Caption         =   "Break-Even Values"
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
      Left            =   300
      TabIndex        =   7
      Top             =   2340
      Width           =   1755
   End
   Begin VB.Label labBreakMisc 
      BackColor       =   &H00000000&
      Caption         =   "Parameters"
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
      Left            =   3360
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
   Begin VB.Line linValuesMiddle 
      BorderColor     =   &H00FFFF00&
      X1              =   480
      X2              =   8700
      Y1              =   2940
      Y2              =   2940
   End
   Begin VB.Line linValuesRight 
      BorderColor     =   &H00FFFF00&
      X1              =   8760
      X2              =   8760
      Y1              =   2340
      Y2              =   6060
   End
   Begin VB.Line linValuesLeft 
      BorderColor     =   &H00FFFF00&
      X1              =   420
      X2              =   420
      Y1              =   2340
      Y2              =   6060
   End
   Begin VB.Line linValuesBottom 
      BorderColor     =   &H00FFFF00&
      X1              =   360
      X2              =   8820
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line linValuesTop 
      BorderColor     =   &H00FFFF00&
      X1              =   360
      X2              =   8820
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line linBreakRight 
      BorderColor     =   &H00FFFF00&
      X1              =   8940
      X2              =   8940
      Y1              =   240
      Y2              =   1920
   End
   Begin VB.Line linBreakLeft 
      BorderColor     =   &H00FFFF00&
      X1              =   3480
      X2              =   3480
      Y1              =   240
      Y2              =   1920
   End
   Begin VB.Line linBreakBottom 
      BorderColor     =   &H00FFFF00&
      X1              =   3420
      X2              =   9000
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Line linBreakTop 
      BorderColor     =   &H00FFFF00&
      X1              =   3420
      X2              =   9000
      Y1              =   300
      Y2              =   300
   End
   Begin VB.Label labBreakTitles 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Unit"
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
      Left            =   6840
      TabIndex        =   5
      Top             =   2580
      Width           =   1095
   End
   Begin VB.Label labBreakTitles 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Value"
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
      Left            =   5640
      TabIndex        =   4
      Top             =   2580
      Width           =   1095
   End
   Begin VB.Label labBreakTitles 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Item"
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
      Left            =   2160
      TabIndex        =   3
      Top             =   2580
      Width           =   1035
   End
   Begin VB.Label labBreakTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Discount Rate:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   3720
      TabIndex        =   2
      Top             =   1020
      Width           =   3495
   End
   Begin VB.Label labBreakTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Percent Change Required for Break Even:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   1
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Label labBreakEvenHeading 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Break-Even Analysis"
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
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmBreakEven"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tempwide As Integer
Dim temphigh As Integer
Dim tempold(25) As String

Private Sub chkDiscount_Click()

If DoNotChange = True Then Exit Sub

If chkDiscount.Value = 0 Then
  labBreakOut(0).Caption = "0.00 percent"
Else
  labBreakOut(0).Caption = Format(LTrim(RTrim(Str(Sets(25)))), "##0.00") & " percent"
End If

labBreakTitles(5).Visible = True

Call getthebreak(hscTagNumber.Value, IndTagData(hscTagNumber.Value).SetNumber)

End Sub

Private Sub Form_Activate()

  DoNotChange = True
  
  ShowMenu = True
   
  vscBreakEven.Visible = False
  vscBreakEven.Value = 0
  
  hscTagNumber.Value = 1
  labSetNumber.Caption = LTrim(RTrim(Str(IndTagData(hscTagNumber.Value).SetNumber)))
  chkDiscount.Value = 0
    
  labBreakOut(0).Caption = "0.00 percent"
  
  labBreakTitles(5).Visible = True
  
  DoNotChange = False
  
  Call getthebreak(hscTagNumber.Value, IndTagData(hscTagNumber.Value).SetNumber)
  
End Sub

Private Sub Form_Deactivate()

If ShowMenu = True Then
  frmBreakEven.Hide
  Call InputMenuAccess(2)
End If
  
End Sub

Private Sub Form_Load()

If FullScreen = False Then
  frmBreakEven.Top = (Screen.Height - (frmBreakEven.Height + 350)) / 2
  frmBreakEven.Left = (Screen.Width - frmBreakEven.Width) / 2
Else
  frmBreakEven.Top = 0
  frmBreakEven.Left = 0
  frmBreakEven.WindowState = 2
End If

If frmBreakEven.Top < 0 Then frmBreakEven.Top = 0
If frmBreakEven.Left < 0 Then frmBreakEven.Left = 0

tempwide = frmBreakEven.ScaleWidth
temphigh = frmBreakEven.ScaleHeight

Call screenstuff

End Sub


Private Sub Form_Resize()

tempwide = frmBreakEven.ScaleWidth
temphigh = frmBreakEven.ScaleHeight

Call screenstuff

End Sub

Private Sub Form_Unload(Cancel As Integer)

  frmBreakEven.Hide
  If ShowMenu = True Then Call InputMenuAccess(2)

End Sub

Private Sub hscTagNumber_Change()

Dim i As Integer

labTagNumber.Caption = LTrim(RTrim(Str(hscTagNumber.Value)))
labSetNumber.Caption = LTrim(RTrim(Str(IndTagData(hscTagNumber.Value).SetNumber)))
If DoNotChange = True Then Exit Sub

DoNotChange = True

For i = 0 To 10
  labItem(i).Caption = ""
  labValue(i).Caption = ""
  labUnit(i).Caption = ""
Next i

labBreakTitles(5).Visible = True

DoNotChange = False

Call getthebreak(hscTagNumber.Value, IndTagData(hscTagNumber.Value).SetNumber)

End Sub

Private Sub imgBackToMenu_Click()

  frmBreakEven.Hide
  Call InputMenuAccess(2)

End Sub

Private Sub labBackToMenu_Click()
  
  frmBreakEven.Hide
  Call InputMenuAccess(2)

End Sub

Public Sub getthebreak(thetag As Integer, theset As Integer)

Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim l As Integer
Dim r As Integer
Dim depcount As Integer
Dim pvbe As Currency
Dim testbe(3) As Currency
Dim uptest As Currency
Dim lowtest As Currency
Dim halftest As Currency
Dim oldtest As Currency
Dim clos As Currency
Dim beperc As Currency
Dim oldvalue(50) As Currency
Dim ii As Integer
Dim testvalue(51) As Currency
Dim tempout As String
Dim tempsend As Integer
Dim test As Integer

vscBreakEven.Visible = False
vscBreakEven.Value = 0
  
For i = 1 To 10
  labItem(i).Caption = ""
  labValue(i).Caption = ""
  labUnit(i).Caption = ""
Next i

For i = 1 To 180
  If i < 131 Then
    If Tagged(theset, i).Independent = thetag Then
      oldvalue(1) = Primary(theset, i)
      ii = i
    End If
  Else
    If Tagged(theset, i).Independent = thetag Then
      oldvalue(1) = CapitalData(i - 131).PurchaseAmount
      ii = i
    End If
  End If
Next i

labItem(0).Caption = LTrim(RTrim(IndTagData(thetag).Title))
labUnit(0).Caption = LTrim(RTrim(IndTagData(thetag).Units))

If ii = 0 Then
  labBreakOut(1).Caption = "0.00 percent"
End If
If ii = 0 Then Exit Sub

depcount = 0

For i = 1 To 25
  If DepTagData(thetag, i).TheCell <> 0 Then
    depcount = depcount + 1
    If depcount > 10 Then vscBreakEven.Visible = True
    If DepTagData(thetag, i).TheCell < 131 Then
      oldvalue(depcount + 1) = Primary(DepTagData(thetag, i).SetNumber, DepTagData(thetag, i).TheCell)
    Else
      oldvalue(depcount + 1) = CapitalData(DepTagData(thetag, i).TheCell - 131).PurchaseAmount
    End If
    If depcount < 11 Then
      labItem(depcount).Caption = LTrim(RTrim(DepTagData(thetag, i).Title))
      labUnit(depcount).Caption = LTrim(RTrim(DepTagData(thetag, i).Units))
    End If
  End If
Next i

vscBreakEven.max = depcount - 10

If oldvalue(1) = 0 Then
  k = 1: l = 0
  GoSub passbe
End If

For i = 1 To 3
  If i = 1 Then
    testvalue(1) = 0
  ElseIf i = 2 Then
    testvalue(1) = oldvalue(1)
    lowtest = oldvalue(1)
  Else
    uptest = 2 * oldvalue(1)
    testvalue(1) = 2 * oldvalue(1)
  End If
  For j = 2 To depcount + 1
    If oldvalue(1) <> 0 Then testvalue(j) = oldvalue(j) * (testvalue(1) / oldvalue(1))
  Next j
  
  Call findvalues(ii, theset, thetag, testvalue, depcount)
  Call cflow5(1, 0)
  Call rateofreturn
  If chkDiscount.Value = 0 Then
    pvbe = Pv0
  Else
    pvbe = Pv2
  End If
  testbe(i) = pvbe
  If i = 2 Then
    If testbe(1) = testbe(2) Then
      k = 1: l = 0
    End If
  End If
Next i

If testbe(2) < testbe(1) Then
  If testbe(1) > 0 And testbe(2) <= 0 Then
    uptest = lowtest: lowtest = 0
  ElseIf testbe(2) > 0 And testbe(3) <= 0 Then
'    lowtest = lowtest: uptest = uptest
  Else
    k = 1: l = 0
  End If
  For test = 1 To 20
    halftest = (lowtest + uptest) / 2
    If test > 1 Then
      If halftest <> 0 Then clos = (halftest - oldtest) / halftest
      If Abs(clos) < BeClos Then r = 1
    End If
    If r = 0 Then
      testvalue(1) = halftest
      For j = 2 To depcount + 1
        If oldvalue(1) <> 0 Then testvalue(j) = oldvalue(j) * (testvalue(1) / oldvalue(1))
      Next j
  
      Call findvalues(ii, theset, thetag, testvalue, depcount)
      Call cflow5(1, 0)
      Call rateofreturn
      If chkDiscount.Value = 0 Then
        pvbe = Pv0
      Else
        pvbe = Pv2
      End If
      beperc = (testvalue(1) / oldvalue(1) - 1) * 100
      labBreakOut(1).Caption = Format(LTrim(RTrim(Str(beperc))), "##0.00") & " percent"
      If pvbe > 0 Then
        lowtest = halftest
      Else
        uptest = halftest
      End If
      oldtest = halftest
    Else
      r = 0
    End If
  Next test
ElseIf testbe(2) > testbe(1) Then
  If testbe(1) < 0 And testbe(2) >= 0 Then
    uptest = lowtest: lowtest = 0
  ElseIf testbe(2) < 0 And testbe(3) >= 0 Then
    lowtest = lowtest: uptest = uptest
  Else
    k = 1: l = 0
  End If
  For test = 1 To 20
    halftest = (lowtest + uptest) / 2
    If test > 1 Then
      If halftest <> 0 Then clos = (halftest - oldtest) / halftest
      If Abs(clos) < BeClos Then r = 1
    End If
    If r = 0 Then
      testvalue(1) = halftest
      For j = 2 To depcount + 1
        If oldvalue(1) <> 0 Then testvalue(j) = oldvalue(j) * (testvalue(1) / oldvalue(1))
      Next j
  
      Call findvalues(ii, theset, thetag, testvalue, depcount)
      Call cflow5(1, 0)
      Call rateofreturn
      If chkDiscount.Value = 0 Then
        pvbe = Pv0
      Else
        pvbe = Pv2
      End If
      beperc = (testvalue(1) / oldvalue(1) - 1) * 100
      labBreakOut(1).Caption = Format(LTrim(RTrim(Str(beperc))), "##0.00") & " percent"
      If pvbe < 0 Then
        lowtest = halftest
      Else
        uptest = halftest
      End If
      oldtest = halftest
    Else
      r = 0
    End If
  Next test
End If

passbe:

If k = 1 Then
  labBreakOut(1).Caption = "0.00 percent"
Else
  testvalue(1) = (lowtest + uptest) / 2
  For j = 2 To depcount + 1
    testvalue(j) = oldvalue(j) * (testvalue(1) / oldvalue(1))
  Next j
  
  Call findvalues(ii, theset, thetag, testvalue, depcount)
  Call cflow5(1, 0)
  Call rateofreturn
  If chkDiscount.Value = 0 Then
    pvbe = Pv0
  Else
    pvbe = Pv2
  End If
  beperc = (testvalue(1) / oldvalue(1) - 1) * 100
  labBreakOut(1).Caption = Format(LTrim(RTrim(Str(beperc))), "##0.00") & " percent"
End If
    
tempout = LTrim(RTrim(Str(oldvalue(1) * (1 + (beperc / 100)))))
tempsend = ii
Call findaformat(tempsend, tempout)
labValue(0).Caption = LTrim(RTrim(tempout))
For i = 2 To depcount + 1
  tempout = LTrim(RTrim(Str(oldvalue(i) * (1 + (beperc / 100)))))
  tempsend = DepTagData(thetag, i - 1).TheCell
  tempold(i) = tempout
  Call findaformat(tempsend, tempout)
  If i < 12 Then
    labValue(i - 1).Caption = LTrim(RTrim(tempout))
  End If
Next i

'Clean-up

If ii < 131 Then
  Primary(theset, ii) = oldvalue(1)
Else
  CapitalData(ii - 131).PurchaseAmount = oldvalue(1)
End If

For i = 1 To depcount
  If DepTagData(thetag, i).TheCell < 131 Then
    Primary(DepTagData(thetag, i).SetNumber, DepTagData(thetag, i).TheCell) = oldvalue(i + 1)
  Else
    CapitalData(DepTagData(thetag, i).TheCell - 131).PurchaseAmount = oldvalue(i + 1)
  End If
Next i

Erase oldvalue
Erase testvalue

labBreakTitles(5).Visible = False

End Sub

Public Sub findvalues(ii, theset, thetag, testvalue, depcount)

Dim i As Integer

If labBreakTitles(5).Visible = True Then
   labBreakTitles(5).Visible = False
Else
   labBreakTitles(5).Visible = True
End If

Call sleep(0.15)

If ii < 131 Then
  Primary(theset, ii) = testvalue(1)
Else
  CapitalData(ii - 131).PurchaseAmount = testvalue(1)
End If
Call recalc(theset, ii)
  
For i = 1 To depcount
  If DepTagData(thetag, i).TheCell < 131 Then
    Primary(DepTagData(thetag, i).SetNumber, DepTagData(thetag, i).TheCell) = testvalue(i + 1)
  Else
    CapitalData(DepTagData(thetag, i).TheCell - 131).PurchaseAmount = testvalue(i + 1)
  End If
  Call recalc(DepTagData(thetag, i).SetNumber, DepTagData(thetag, i).TheCell)
Next i

End Sub

Private Sub labtest_Click()

End Sub

Public Sub screenstuff()

  Dim X As Integer
   
  labBreakEvenHeading.Top = temphigh * 0.0187
  labBreakEvenHeading.Left = tempwide * 0.0131
  
  linTagTop.X1 = tempwide * 0.0656
  linTagTop.X2 = tempwide * 0.2492
  linTagTop.Y1 = temphigh * 0.1028
  linTagTop.Y2 = temphigh * 0.1028
  
  linTagLeft.X1 = tempwide * 0.0721
  linTagLeft.X2 = tempwide * 0.0721
  linTagLeft.Y1 = temphigh * 0.0935
  linTagLeft.Y2 = temphigh * 0.3084

  linTagRight.X1 = tempwide * 0.2426
  linTagRight.X2 = tempwide * 0.2426
  linTagRight.Y1 = temphigh * 0.0935
  linTagRight.Y2 = temphigh * 0.3048

  linTagBottom.X1 = tempwide * 0.0656
  linTagBottom.X2 = tempwide * 0.2492
  linTagBottom.Y1 = temphigh * 0.2991
  linTagBottom.Y2 = temphigh * 0.2991

  linBreakTop.X1 = tempwide * 0.3738
  linBreakTop.X2 = tempwide * 0.9836
  linBreakTop.Y1 = temphigh * 0.0467
  linBreakTop.Y2 = temphigh * 0.0467

  linBreakLeft.X1 = tempwide * 0.3803
  linBreakLeft.X2 = tempwide * 0.3803
  linBreakLeft.Y1 = temphigh * 0.0374
  linBreakLeft.Y2 = temphigh * 0.2991

  linBreakRight.X1 = tempwide * 0.977
  linBreakRight.X2 = tempwide * 0.977
  linBreakRight.Y1 = temphigh * 0.0374
  linBreakRight.Y2 = temphigh * 0.2991

  linBreakBottom.X1 = tempwide * 0.3738
  linBreakBottom.X2 = tempwide * 0.9836
  linBreakBottom.Y1 = temphigh * 0.2897
  linBreakBottom.Y2 = temphigh * 0.2897

  linValuesTop.X1 = tempwide * 0.0393
  linValuesTop.X2 = tempwide * 0.9639
  linValuesTop.Y1 = temphigh * 0.3738
  linValuesTop.Y2 = temphigh * 0.3738

  linValuesLeft.X1 = tempwide * 0.0459
  linValuesLeft.X2 = tempwide * 0.0459
  linValuesLeft.Y1 = temphigh * 0.3645
  linValuesLeft.Y2 = temphigh * 0.9439

  linValuesMiddle.X1 = tempwide * 0.0525
  linValuesMiddle.X2 = tempwide * 0.9508
  linValuesMiddle.Y1 = temphigh * 0.4579
  linValuesMiddle.Y2 = temphigh * 0.4579

  linLast.X1 = tempwide * 0.0525
  linLast.X2 = tempwide * 0.9508
  linLast.Y1 = temphigh * 0.5327
  linLast.Y2 = temphigh * 0.5327

  linValuesRight.X1 = tempwide * 0.9574
  linValuesRight.X2 = tempwide * 0.9574
  linValuesRight.Y1 = temphigh * 0.3645
  linValuesRight.Y2 = temphigh * 0.9439

  linValuesBottom.X1 = tempwide * 0.0393
  linValuesBottom.X2 = tempwide * 0.9639
  linValuesBottom.Y1 = temphigh * 0.9346
  linValuesBottom.Y2 = temphigh * 0.9346

  For X = 0 To 1
    labBreakTitles(X).Top = (X * 0.0654 * temphigh) + (temphigh * 0.1589)
    labBreakTitles(X).Left = tempwide * 0.4066
    labBreakTitles(X).Width = tempwide * 0.382
    labBreakOut(X).Top = (X * 0.0654 * temphigh) + (temphigh * 0.1589)
    labBreakOut(X).Left = tempwide * 0.8197
  Next X

  For X = 2 To 4
    labBreakTitles(X).Top = temphigh * 0.4019
    If X = 2 Then
      labBreakTitles(X).Left = tempwide * 0.2361
      labBreakTitles(X).Width = tempwide * 0.1131
    ElseIf X = 3 Then
      labBreakTitles(X).Left = tempwide * 0.6164
      labBreakTitles(X).Width = tempwide * 0.1852
    Else
      labBreakTitles(X).Left = tempwide * 0.7475
      labBreakTitles(X).Width = tempwide * 0.1197
    End If
  Next X

  labBreakTitles(5).Top = temphigh * 0.3148
  labBreakTitles(5).Left = tempwide * 0.4725
  labBreakTitles(5).Width = tempwide * 0.4198
  
  For X = 0 To 10
    labItem(X).Top = ((X - 1) * 0.0374 * temphigh) + (temphigh * 0.5514)
    labItem(X).Left = tempwide * 0.0918
    labItem(X).Width = tempwide * 0.4016
    labValue(X).Top = ((X - 1) * 0.0374 * temphigh) + (temphigh * 0.5514)
    labValue(X).Left = tempwide * 0.5508
    labValue(X).Width = tempwide * 0.1852
    labUnit(X).Top = ((X - 1) * 0.0374 * temphigh) + (temphigh * 0.5514)
    labUnit(X).Left = tempwide * 0.7475
    labUnit(X).Width = tempwide * 0.1721
  Next X

  labItem(0).Top = temphigh * 0.486
  labValue(0).Top = temphigh * 0.486
  labUnit(0).Top = temphigh * 0.486
  
  chkDiscount.Top = temphigh * 0.0935
  chkDiscount.Left = (tempwide * 0.6828) - 2528

  labBreakMisc(0).Top = temphigh * 0.0374
  labBreakMisc(0).Left = tempwide * 0.3672

  labBreakMisc(1).Top = temphigh * 0.3645
  labBreakMisc(1).Left = tempwide * 0.0328

  labTagTitle.Top = temphigh * 0.1121
  labTagTitle.Left = tempwide * 0.0984
  labTagTitle.Width = tempwide * 0.1197
  
  hscTagNumber.Top = temphigh * 0.1589
  hscTagNumber.Left = (tempwide * 0.1451) - 188
  
  vscBreakEven.Top = temphigh * 0.542
  vscBreakEven.Height = temphigh * 0.3855
  vscBreakEven.Left = tempwide * 0.5115
  vscBreakEven.Width = tempwide * 0.0213
  
  labTagNumber.Top = temphigh * 0.1558
  labTagNumber.Left = tempwide * 0.177
  labTagNumber.Width = tempwide * 0.0279
  
  labSetTitle.Top = temphigh * 0.215
  labSetTitle.Left = tempwide * 0.0984
  labSetTitle.Width = tempwide * 0.1197
  
  labSetNumber.Top = temphigh * 0.2523
  labSetNumber.Left = tempwide * 0.1443
  labSetNumber.Width = tempwide * 0.0279
  
  labBackToMenu.Top = temphigh * 0.9462
  labBackToMenu.Left = tempwide * 0.0656

  imgBackToMenu.Top = temphigh * 0.9555
  imgBackToMenu.Left = tempwide * 0.0066
  imgBackToMenu.Width = tempwide * 0.0541

  labPrintScreen.Top = temphigh * 0.9462
  labPrintScreen.Left = tempwide * 0.8787

End Sub

Private Sub labPrintScreen_Click()

ShowMenu = False
job = 10
Call printstuffout(job)

End Sub

Private Sub timFlash_Timer()

  If labBreakTitles(5).Visible = False Then
    labBreakTitles(5).Visible = True
  Else
    labBreakTitles(5).Visible = False
  End If

End Sub

Public Sub sleep(sngNumberOfSeconds As Single)

  Dim sngEndTime As Single
  sngEndTime = Timer + sngNumberOfSeconds
  Do
    DoEvents
  Loop Until Timer >= sngEndTime
  
End Sub

Private Sub vscBreakEven_Change()
Dim i As Integer
Dim tempsend As Integer

For i = 1 To 10
  labItem(i).Caption = LTrim(RTrim(DepTagData(hscTagNumber.Value, vscBreakEven.Value + i).Title))
  labUnit(i).Caption = LTrim(RTrim(DepTagData(hscTagNumber.Value, vscBreakEven.Value + i).Units))
  tempsend = DepTagData(hscTagNumber.Value, vscBreakEven.Value + i).TheCell
  Call findaformat(tempsend, tempold(vscBreakEven.Value + i + 1))
  labValue(i).Caption = LTrim(RTrim(tempold(vscBreakEven.Value + i + 1)))
Next i

End Sub
