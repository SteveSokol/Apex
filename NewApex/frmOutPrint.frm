VERSION 5.00
Begin VB.Form frmOutPrint 
   BackColor       =   &H00000000&
   Caption         =   "Printer Untilities"
   ClientHeight    =   5655
   ClientLeft      =   2670
   ClientTop       =   1485
   ClientWidth     =   6045
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
   ScaleHeight     =   5655
   ScaleWidth      =   6045
   Begin VB.CommandButton comLeavePrint 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   4680
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   4980
      Width           =   975
   End
   Begin VB.ListBox lstValueFontSize 
      Height          =   735
      ItemData        =   "frmOutPrint.frx":0000
      Left            =   4500
      List            =   "frmOutPrint.frx":000D
      TabIndex        =   23
      Top             =   3660
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox lstFontSize 
      Height          =   1635
      ItemData        =   "frmOutPrint.frx":001B
      Left            =   4500
      List            =   "frmOutPrint.frx":0034
      TabIndex        =   21
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton comPrintItOut 
      Caption         =   "&Print"
      Height          =   315
      Left            =   3660
      TabIndex        =   10
      Top             =   4980
      Width           =   975
   End
   Begin VB.CheckBox chkPrintToFile 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Print To File"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   420
      TabIndex        =   9
      Top             =   5100
      Width           =   1335
   End
   Begin VB.CheckBox chkPrintTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Underline"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   2
      Left            =   600
      TabIndex        =   5
      Top             =   2520
      Width           =   1155
   End
   Begin VB.CheckBox chkPrintTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Italic"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   1020
      TabIndex        =   4
      Top             =   2160
      Width           =   735
   End
   Begin VB.CheckBox chkPrintTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Bold"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   1020
      TabIndex        =   3
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox txtPrintOutItem 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   1560
      TabIndex        =   8
      Text            =   "4"
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox txtPrintOutItem 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   1560
      TabIndex        =   7
      Text            =   "9"
      Top             =   3660
      Width           =   375
   End
   Begin VB.TextBox txtPrintOutItem 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   1560
      TabIndex        =   6
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox txtPrintOutItem 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   1560
      TabIndex        =   2
      Text            =   "11"
      Top             =   1380
      Width           =   375
   End
   Begin VB.TextBox txtPrintOutItem 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   1560
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.ListBox lstFontList 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4050
      ItemData        =   "frmOutPrint.frx":0052
      Left            =   3660
      List            =   "frmOutPrint.frx":0054
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   660
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Font Sizes"
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
      Index           =   5
      Left            =   3720
      TabIndex        =   24
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Font Sizes"
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
      Index           =   4
      Left            =   3720
      TabIndex        =   22
      Top             =   780
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Available Fonts"
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
      Index           =   3
      Left            =   3720
      TabIndex        =   20
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Print To File"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   19
      Top             =   4740
      Width           =   1095
   End
   Begin VB.Line linPrintFileRight 
      BorderColor     =   &H00FFFF00&
      X1              =   3000
      X2              =   3000
      Y1              =   4740
      Y2              =   5460
   End
   Begin VB.Line linPrintFileBottom 
      BorderColor     =   &H00FFFF00&
      X1              =   240
      X2              =   3060
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line linPrintFileLeft 
      BorderColor     =   &H00FFFF00&
      X1              =   300
      X2              =   300
      Y1              =   4740
      Y2              =   5460
   End
   Begin VB.Line linPrintFileTop 
      BorderColor     =   &H00FFFF00&
      X1              =   240
      X2              =   3060
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Labels and Values"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   18
      Top             =   2940
      Width           =   1635
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Titles and Headings"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   17
      Top             =   660
      Width           =   1755
   End
   Begin VB.Line linPrintValueRight 
      BorderColor     =   &H00FFFF00&
      X1              =   3000
      X2              =   3000
      Y1              =   2940
      Y2              =   4680
   End
   Begin VB.Line linPrintValueBottom 
      BorderColor     =   &H00FFFF00&
      X1              =   240
      X2              =   3060
      Y1              =   4620
      Y2              =   4620
   End
   Begin VB.Line linPrintValueLeft 
      BorderColor     =   &H00FFFF00&
      X1              =   300
      X2              =   300
      Y1              =   2940
      Y2              =   4680
   End
   Begin VB.Line linPrintValueTop 
      BorderColor     =   &H00FFFF00&
      X1              =   240
      X2              =   3060
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line linPrintTitleRight 
      BorderColor     =   &H00FFFF00&
      X1              =   3000
      X2              =   3000
      Y1              =   660
      Y2              =   2880
   End
   Begin VB.Line linPrintTitleBottom 
      BorderColor     =   &H00FFFF00&
      X1              =   240
      X2              =   3060
      Y1              =   2820
      Y2              =   2820
   End
   Begin VB.Line linPrintTitleLeft 
      BorderColor     =   &H00FFFF00&
      X1              =   300
      X2              =   300
      Y1              =   660
      Y2              =   2880
   End
   Begin VB.Line linPrintTitleTop 
      BorderColor     =   &H00FFFF00&
      X1              =   240
      X2              =   3060
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label labPrintOutHeading 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Printer Utility"
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
      TabIndex        =   16
      Top             =   120
      Width           =   2115
   End
   Begin VB.Label labPrintOutItems 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Columns"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   540
      TabIndex        =   15
      Top             =   4140
      Width           =   915
   End
   Begin VB.Label labPrintOutItems 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Font Size"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   540
      TabIndex        =   14
      Top             =   3720
      Width           =   915
   End
   Begin VB.Label labPrintOutItems 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Font Type"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   540
      TabIndex        =   13
      Top             =   3300
      Width           =   915
   End
   Begin VB.Label labPrintOutItems 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Font Size"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   540
      TabIndex        =   12
      Top             =   1440
      Width           =   915
   End
   Begin VB.Label labPrintOutItems 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Font Type"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   540
      TabIndex        =   11
      Top             =   1020
      Width           =   915
   End
End
Attribute VB_Name = "frmOutPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkPrintTitle_Click(Index As Integer)

Select Case Index
  Case 0
    chkPrintTitle(1).SetFocus
    lstFontList.Visible = False
    Label1(3).Visible = False
  Case 1
    chkPrintTitle(2).SetFocus
  Case 2
    txtPrintOutItem(2).SetFocus
    lstValueFontSize.Visible = True
    Label1(5).Visible = True
End Select

End Sub

Private Sub chkPrintTitle_GotFocus(Index As Integer)

  Label1(3).Visible = False
  lstFontList.Visible = False
  Label1(4).Visible = False
  lstFontSize.Visible = False

End Sub

Private Sub chkPrintToFile_Click()
Dim i As Integer
If chkPrintToFile.Value = 1 Then
  For i = 0 To 4
    txtPrintOutItem(i).Enabled = False
  Next i
  GetPrintFileName = True
  frmApexFileMaker.filApexFile.Refresh
  frmApexFileMaker.filApexFile.Pattern = "*.txt"
  frmApexFileMaker.Command1.Enabled = False
  frmApexFileMaker.Command2.Enabled = False
  frmApexFileMaker.Command3.Enabled = True
  frmApexFileMaker.Show
  frmApexFileMaker.txtApexFile.Text = frmApexFileMaker.filApexFile.Path
  frmApexFileMaker.proFile = 0
  frmApexFileMaker.txtApexFile.SetFocus
ElseIf chkPrintToFile.Value = 0 Then
  For i = 0 To 4
    txtPrintOutItem(i).Enabled = True
  Next i
  GetPrintFileName = False
End If

End Sub

Private Sub comLeavePrint_Click()

  frmOutPrint.Hide

End Sub

Private Sub comPrintItOut_Click()

Select Case job
  Case 1 To 5
    Call cashflowstuff
  Case 6 To 8
    Call statisticalstuff
  Case 9 To 13
    Call analysisstuff
  Case 20 To 24
    Call datastuff
End Select

End Sub

Private Sub Form_Activate()

LastCell = 0

Dim i As Integer
For i = 0 To Printer.FontCount - 1
  lstFontList.AddItem Printer.Fonts(i)
Next i
chkPrintToFile.Value = 0
If job > 5 Then
  txtPrintOutItem(4).Enabled = False
  labPrintOutItems(4).Enabled = False
Else
  txtPrintOutItem(4).Enabled = True
  labPrintOutItems(4).Enabled = True
End If
If job = 7 Or job = 8 Then
  chkPrintToFile.Enabled = False
Else
  chkPrintToFile.Enabled = True
End If

End Sub

Private Sub Form_Load()

frmOutPrint.Top = (Screen.Height - (frmOutPrint.Height + 350)) / 2
frmOutPrint.Left = (Screen.Width - frmOutPrint.Width) / 2

If frmOutPrint.Top < 0 Then frmOutPrint.Top = 0
If frmOutPrint.Left < 0 Then frmOutPrint.Left = 0

End Sub

Private Sub lstFontList_Click()
  
If LastCell = 0 Then
  txtPrintOutItem(0).Font = lstFontList.List(lstFontList.ListIndex)
  txtPrintOutItem(0).Text = lstFontList.List(lstFontList.ListIndex)
  Label1(3).Visible = False
  lstFontList.Visible = False
  txtPrintOutItem(1).SetFocus
ElseIf LastCell = 2 Then
  txtPrintOutItem(2).Font = lstFontList.List(lstFontList.ListIndex)
  txtPrintOutItem(2).Text = lstFontList.List(lstFontList.ListIndex)
  Label1(3).Visible = False
  lstFontList.Visible = False
  txtPrintOutItem(3).SetFocus
End If

End Sub

Private Sub lstFontSize_Click()

txtPrintOutItem(1).Text = lstFontSize.List(lstFontSize.ListIndex)
Label1(3).Visible = False
Label1(4).Visible = False
lstFontList.Visible = False
lstFontSize.Visible = False
chkPrintTitle(0).SetFocus

End Sub

Private Sub lstValueFontSize_Click()
  
txtPrintOutItem(3).Text = lstValueFontSize.List(lstValueFontSize.ListIndex)
Label1(5).Visible = False
lstValueFontSize.Visible = False
If job < 6 Then
  txtPrintOutItem(4).SetFocus
Else
  txtPrintOutItem(0).SetFocus
End If

End Sub


Private Sub txtPrintOutItem_GotFocus(Index As Integer)
LastCell = Index

If Index = 0 Or Index = 2 Then
  Label1(3).Visible = True
  lstFontList.Visible = True
  Label1(4).Visible = False
  lstFontSize.Visible = False
  lstValueFontSize.Visible = False
ElseIf Index = 1 Then
  Label1(3).Visible = False
  lstFontList.Visible = False
  Label1(4).Visible = True
  lstFontSize.Visible = True
  lstValueFontSize.Visible = False
ElseIf Index = 3 Then
  Label1(3).Visible = False
  lstFontList.Visible = False
  Label1(4).Visible = False
  lstFontSize.Visible = False
  lstValueFontSize.Visible = True
Else
  Label1(3).Visible = False
  lstFontList.Visible = False
  Label1(4).Visible = False
  lstFontSize.Visible = False
  lstValueFontSize.Visible = False
End If

End Sub
