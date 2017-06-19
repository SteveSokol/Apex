VERSION 5.00
Begin VB.Form frmCapital 
   BackColor       =   &H00000000&
   Caption         =   "Capital Cost Data"
   ClientHeight    =   6420
   ClientLeft      =   1140
   ClientTop       =   1515
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
   FontTransparent =   0   'False
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6420
   ScaleWidth      =   9150
   Begin VB.CommandButton cmdCapitalList 
      Caption         =   "&Default"
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
      Left            =   3840
      TabIndex        =   13
      Top             =   5220
      Width           =   735
   End
   Begin VB.TextBox txtSpecialDepreciate 
      Height          =   330
      Left            =   4260
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2340
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton comDepTag 
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3300
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   5760
      Width           =   195
   End
   Begin VB.CommandButton comIndTag 
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   420
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   5760
      Width           =   195
   End
   Begin VB.ListBox lstDepreciationList 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      ItemData        =   "frmCapital.frx":0000
      Left            =   5280
      List            =   "frmCapital.frx":0028
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   4320
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton cmdCapitalList 
      Caption         =   "Re&place"
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
      Left            =   3060
      TabIndex        =   12
      Top             =   5220
      Width           =   735
   End
   Begin VB.CommandButton cmdCapitalList 
      Caption         =   "&Clear All"
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
      Left            =   2280
      TabIndex        =   11
      Top             =   5220
      Width           =   735
   End
   Begin VB.CommandButton cmdCapitalList 
      Caption         =   "&Remove"
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
      Index           =   1
      Left            =   1500
      TabIndex        =   10
      Top             =   5220
      Width           =   735
   End
   Begin VB.CommandButton cmdCapitalList 
      Caption         =   "&Add"
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
      Index           =   0
      Left            =   720
      TabIndex        =   9
      Top             =   5220
      Width           =   735
   End
   Begin VB.TextBox txtCapitalValues 
      Enabled         =   0   'False
      Height          =   330
      Index           =   7
      Left            =   2820
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   4500
      Width           =   1455
   End
   Begin VB.TextBox txtCapitalValues 
      Enabled         =   0   'False
      Height          =   330
      Index           =   6
      Left            =   2820
      Locked          =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   4140
      Width           =   1455
   End
   Begin VB.TextBox txtCapitalValues 
      Height          =   330
      Index           =   5
      Left            =   2820
      TabIndex        =   8
      Top             =   3420
      Width           =   1455
   End
   Begin VB.TextBox txtCapitalValues 
      Height          =   330
      Index           =   4
      Left            =   2820
      TabIndex        =   7
      Top             =   3060
      Width           =   1455
   End
   Begin VB.TextBox txtCapitalValues 
      Height          =   330
      Index           =   3
      Left            =   2820
      TabIndex        =   6
      Top             =   2700
      Width           =   1455
   End
   Begin VB.TextBox txtCapitalValues 
      Height          =   330
      Index           =   2
      Left            =   2460
      TabIndex        =   4
      Top             =   2340
      Width           =   2235
   End
   Begin VB.TextBox txtCapitalValues 
      Height          =   330
      Index           =   1
      Left            =   2820
      TabIndex        =   2
      Top             =   1980
      Width           =   1455
   End
   Begin VB.TextBox txtCapitalValues 
      Height          =   330
      Index           =   0
      Left            =   2820
      TabIndex        =   1
      Top             =   1620
      Width           =   1455
   End
   Begin VB.ListBox lstCapitalList 
      Height          =   3435
      ItemData        =   "frmCapital.frx":011B
      Left            =   5760
      List            =   "frmCapital.frx":011D
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   600
      Width           =   2775
   End
   Begin VB.TextBox txtCapitalItem 
      Height          =   330
      Left            =   1020
      TabIndex        =   0
      Top             =   900
      Width           =   3255
   End
   Begin VB.Label labCapitalHelp 
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
      Left            =   4560
      TabIndex        =   44
      Top             =   6060
      Width           =   675
   End
   Begin VB.Label labInsert 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Insert"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2340
      TabIndex        =   43
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label labCapitalMisc 
      BackColor       =   &H00000000&
      Caption         =   "Cost Tallies"
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
      Left            =   240
      TabIndex        =   40
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label labCheckTag 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
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
      Height          =   195
      Left            =   4500
      TabIndex        =   39
      Top             =   1680
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label labDepTag 
      BackColor       =   &H00000000&
      Caption         =   "Dependent Tag"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   38
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Label labIndTag 
      BackColor       =   &H00000000&
      Caption         =   "Independent Tag"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   37
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label labCapitalMisc 
      BackColor       =   &H00000000&
      Caption         =   "Command"
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
      Index           =   12
      Left            =   240
      TabIndex        =   36
      Top             =   4920
      Width           =   975
   End
   Begin VB.Line linLeftBoxLast 
      BorderColor     =   &H00FFFF00&
      X1              =   240
      X2              =   4920
      Y1              =   4980
      Y2              =   4980
   End
   Begin VB.Label labCapitalMisc 
      BackColor       =   &H00000000&
      Caption         =   "years"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   4320
      TabIndex        =   34
      Top             =   2760
      Width           =   435
   End
   Begin VB.Label labCapitalMisc 
      BackColor       =   &H00000000&
      Caption         =   "Treatment"
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
      Left            =   240
      TabIndex        =   33
      Top             =   1320
      Width           =   915
   End
   Begin VB.Label labCapitalMisc 
      BackColor       =   &H00000000&
      Caption         =   "Item List"
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
      Left            =   5460
      TabIndex        =   32
      Top             =   240
      Width           =   795
   End
   Begin VB.Label labCapitalMisc 
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
      Index           =   0
      Left            =   240
      TabIndex        =   31
      Top             =   600
      Width           =   435
   End
   Begin VB.Line linBox2Bottom 
      BorderColor     =   &H00FFFF00&
      X1              =   5520
      X2              =   8760
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line linBox2Middle 
      BorderColor     =   &H00FFFF00&
      X1              =   240
      X2              =   4920
      Y1              =   3900
      Y2              =   3900
   End
   Begin VB.Line linBox2Top 
      BorderColor     =   &H00FFFF00&
      X1              =   5520
      X2              =   8760
      Y1              =   300
      Y2              =   300
   End
   Begin VB.Line linBox2Right 
      BorderColor     =   &H00FFFF00&
      X1              =   8700
      X2              =   8700
      Y1              =   240
      Y2              =   4260
   End
   Begin VB.Line linBox2Left 
      BorderColor     =   &H00FFFF00&
      X1              =   5580
      X2              =   5580
      Y1              =   240
      Y2              =   4260
   End
   Begin VB.Line linBox1Bottom 
      BorderColor     =   &H00FFFF00&
      X1              =   240
      X2              =   5040
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line linBox1Middle 
      BorderColor     =   &H00FFFF00&
      X1              =   240
      X2              =   4920
      Y1              =   1380
      Y2              =   1380
   End
   Begin VB.Line linBox1Top 
      BorderColor     =   &H00FFFF00&
      X1              =   240
      X2              =   5040
      Y1              =   660
      Y2              =   660
   End
   Begin VB.Line linBox1Right 
      BorderColor     =   &H00FFFF00&
      X1              =   4980
      X2              =   4980
      Y1              =   600
      Y2              =   5700
   End
   Begin VB.Line linBox1Left 
      BorderColor     =   &H00FFFF00&
      X1              =   300
      X2              =   300
      Y1              =   600
      Y2              =   5700
   End
   Begin VB.Image imgBackToMenu 
      Height          =   195
      Left            =   60
      Picture         =   "frmCapital.frx":011F
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
      Left            =   600
      TabIndex        =   30
      Top             =   6075
      Width           =   675
   End
   Begin VB.Label labCapitalMisc 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "$"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   10
      Left            =   2640
      TabIndex        =   27
      Top             =   4560
      Width           =   135
   End
   Begin VB.Label labCapitalMisc 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "$"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   9
      Left            =   2640
      TabIndex        =   26
      Top             =   4200
      Width           =   135
   End
   Begin VB.Label labCapitalMisc 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "$"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   2640
      TabIndex        =   25
      Top             =   3480
      Width           =   135
   End
   Begin VB.Label labCapitalMisc 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "$"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   2640
      TabIndex        =   24
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label labCapitalMisc 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Tag"
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
      Height          =   225
      Index           =   6
      Left            =   4500
      TabIndex        =   23
      Top             =   1440
      Width           =   345
   End
   Begin VB.Label labCapitalMisc 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Salvage Values"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   5
      Left            =   600
      TabIndex        =   22
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label labCapitalMisc 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Capital Costs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   4
      Left            =   600
      TabIndex        =   21
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label labCapitalHeading 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Capital Cost Data"
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
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label labCapitalTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Salvage Value"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   600
      TabIndex        =   19
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label labCapitalTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Year Sold"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   600
      TabIndex        =   18
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label labCapitalTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Depreciation Period"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   17
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label labCapitalTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Treat As"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   16
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label labCapitalTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Year Invested"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   15
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label labCapitalTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Amount"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   14
      Top             =   1680
      Width           =   1695
   End
End
Attribute VB_Name = "frmCapital"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim AddedYet As Integer
Dim temphigh As Single
Dim tempwide As Single
Private Sub cmdCapitalList_Click(Index As Integer)

  Dim X As Integer
  Dim ind As Integer
  Dim tempcap As Currency
  Dim tempsalv As Currency
   
  PageChange(5) = True
  
  If Val(txtCapitalValues(1).Text) < (Sets(12) - 1) Or Val(txtCapitalValues(1).Text) > (50 + Sets(12) - 1) Then
    WarnNumber = 6
    ShowMenu = False
    frmWarnTheUser.Show
    Exit Sub
  End If
  If Val(txtCapitalValues(4).Text) < (Sets(12) - 1) Or Val(txtCapitalValues(4).Text) > (50 + Sets(12) - 1) Then
    WarnNumber = 7
    ShowMenu = False
    frmWarnTheUser.Show
    Exit Sub
  End If
  
  Select Case Index
    Case 0
      AddedYet = False
      lstCapitalList.AddItem txtCapitalItem.Text
      ind = lstCapitalList.ListCount - 1
      CapitalData(ind).Item = LTrim(txtCapitalItem.Text)
      CapitalData(ind).PurchaseAmount = CCur(Val(txtCapitalValues(0).Text))
      CapitalData(ind).InvestYear = CInt(Val(txtCapitalValues(1).Text)) - Sets(12) + 1
      CapitalData(ind).DepMethod = Left(LTrim(LCase(txtCapitalValues(2).Text)), 3)
      If Left(LTrim(LCase(txtCapitalValues(2).Text)), 10) = "straight a" Then
        CapitalData(ind).DepMethod = "amo"
      End If
      CapitalData(ind).DepPeriod = CInt(Val(txtCapitalValues(3).Text))
      CapitalData(ind).DmRate = CInt(Val(txtSpecialDepreciate.Text))
      CapitalData(ind).SoldYear = CInt(Val(txtCapitalValues(4).Text)) - Sets(12) + 1
      CapitalData(ind).SalvageAmount = CCur(Val(txtCapitalValues(5).Text))
      NumCap = NumCap + 1
    Case 1
      If lstCapitalList.ListIndex >= 0 Then
        ind = lstCapitalList.ListIndex
        lstCapitalList.RemoveItem lstCapitalList.ListIndex
        For X = ind To lstCapitalList.ListCount + 1
          CapitalData(X).Item = CapitalData(X + 1).Item
          CapitalData(X).PurchaseAmount = CapitalData(X + 1).PurchaseAmount
          CapitalData(X).InvestYear = CapitalData(X + 1).InvestYear
          CapitalData(X).DepMethod = CapitalData(X + 1).DepMethod
          CapitalData(X).DepPeriod = CapitalData(X + 1).DepPeriod
          CapitalData(X).DmRate = CapitalData(X + 1).DmRate
          CapitalData(X).Changed = CapitalData(X + 1).Changed
          CapitalData(X).SoldYear = CapitalData(X + 1).SoldYear
          CapitalData(X).SalvageAmount = CapitalData(X + 1).SalvageAmount
          Tagged(1, 131 + X).Independent = Tagged(1, 132 + X).Independent
          Tagged(1, 131 + X).Dependent = Tagged(1, 132 + X).Dependent
        Next X
        NumCap = NumCap - 1
        DoNotChange = True
          lstCapitalList.ListIndex = -1
        DoNotChange = False
      Else
        Beep
      End If
    Case 2
      NumCap = 0
      For X = 0 To lstCapitalList.ListCount
        CapitalData(X).Item = ""
        CapitalData(X).PurchaseAmount = 0
        CapitalData(X).InvestYear = 0
        CapitalData(X).DepMethod = ""
        CapitalData(X).DepPeriod = 0
        CapitalData(X).DmRate = 0
        CapitalData(X).Changed = False
        CapitalData(X).SoldYear = 0
        CapitalData(X).SalvageAmount = 0
        Tagged(1, 131 + X).Independent = 0
        Tagged(1, 131 + X).Dependent = 0
      Next X
      lstCapitalList.Clear
    Case 3
      If lstCapitalList.ListIndex >= 0 Then
        ind = lstCapitalList.ListIndex
        CapitalData(ind).Item = LTrim(txtCapitalItem.Text)
        CapitalData(ind).PurchaseAmount = CCur(Val(txtCapitalValues(0).Text))
        CapitalData(ind).InvestYear = CInt(Val(txtCapitalValues(1).Text)) - Sets(12) + 1
        CapitalData(ind).DepMethod = Left(LTrim(LCase(txtCapitalValues(2).Text)), 3)
        If Left(LTrim(LCase(txtCapitalValues(2).Text)), 10) = "straight a" Then
          CapitalData(ind).DepMethod = "amo"
        End If
        CapitalData(ind).DepPeriod = CInt(Val(txtCapitalValues(3).Text))
        CapitalData(ind).DmRate = CInt(Val(txtSpecialDepreciate.Text))
        CapitalData(ind).SoldYear = CInt(Val(txtCapitalValues(4).Text)) - Sets(12) + 1
        CapitalData(ind).SalvageAmount = CCur(Val(txtCapitalValues(5).Text))
        If ChangedFlag = True Then
          ChangedFlag = False
          CapitalData(ind).Changed = True
        End If
        DoNotChange = True
          lstCapitalList.ListIndex = -1
        DoNotChange = False
      Else
        Beep
      End If
    Case 4
      For X = 0 To lstCapitalList.ListCount
        CapitalData(X).Item = ""
        CapitalData(X).PurchaseAmount = 0
        CapitalData(X).InvestYear = 0
        CapitalData(X).DepMethod = ""
        CapitalData(X).DepPeriod = 0
        CapitalData(X).Changed = 0
        CapitalData(X).SoldYear = 0
        CapitalData(X).SalvageAmount = 0
        Tagged(1, 131 + X).Independent = 0
        Tagged(1, 131 + X).Dependent = 0
      Next X
      lstCapitalList.Clear
      CapitalData(0).Item = "Acquisition": CapitalData(0).DepMethod = "acq":
      CapitalData(1).Item = "Exploration": CapitalData(1).DepMethod = "dev": CapitalData(1).DepPeriod = 5
      CapitalData(2).Item = "Engineering & Construction": CapitalData(2).DepMethod = "dev": CapitalData(2).DepPeriod = 5
      CapitalData(3).Item = "Development": CapitalData(3).DepMethod = "dev": CapitalData(3).DepPeriod = 5
      CapitalData(4).Item = "Pre-Production Stripping": CapitalData(4).DepMethod = "dev": CapitalData(4).DepPeriod = 5
      CapitalData(5).Item = "Infrastructure": CapitalData(5).DepMethod = "dev": CapitalData(5).DepPeriod = 5
      CapitalData(6).Item = "Buildings": CapitalData(6).DepMethod = "str": CapitalData(6).DepPeriod = 32
      CapitalData(7).Item = "Mine Equipment": CapitalData(7).DepMethod = "mod": CapitalData(7).DepPeriod = 7
      CapitalData(8).Item = "Mill Equipment": CapitalData(8).DepMethod = "mod": CapitalData(8).DepPeriod = 7
      CapitalData(9).Item = "Working Capital": CapitalData(9).DepMethod = "wor"
      CapitalData(10).Item = "Contingency": CapitalData(10).DepMethod = "dev": CapitalData(10).DepPeriod = 5
      CapitalData(11).Item = "Reclamation": CapitalData(11).DepMethod = "rec"
      For X = 0 To 11
        lstCapitalList.AddItem CapitalData(X).Item
      Next X
      NumCap = 12
  End Select
  
  txtCapitalItem.Text = ""
  
  txtSpecialDepreciate.Visible = False
  
  DoNotChange = True
    For X = 0 To 5
      txtCapitalValues(X).Text = ""
    Next X
  DoNotChange = False
  
  labCheckTag.Caption = ""
  labCheckTag.Visible = False
  
  tempcap = 0
  tempsalv = 0
  
  For X = 0 To NumCap - 1
    tempcap = tempcap + CapitalData(X).PurchaseAmount
    tempsalv = tempsalv + CapitalData(X).SalvageAmount
  Next X
  
  txtCapitalValues(6).Text = Format(Str(tempcap), "###,###,###,###")
  txtCapitalValues(7).Text = Format(Str(tempsalv), "###,###,###,###")
    
  txtCapitalItem.SetFocus

End Sub

Private Sub cmdCapitalList_GotFocus(Index As Integer)
LastCell = Index + 7
End Sub

Private Sub comDepTag_Click()

If nTag = 0 Then
  WarnNumber = 4
  ShowMenu = False
  frmWarnTheUser.Show
Else
  If labCheckTag.Visible = False Then
    ParamSet = False
    dTag = dTag + 1
    labCheckTag.Visible = True
    labCheckTag.ForeColor = &HFFFF&
    labCheckTag.Caption = LTrim(Str(nTag))
    If AddedYet = True Then
      Tagged(1, lstCapitalList.ListIndex + 131).Dependent = nTag
      DepTagData(nTag, dTag).TheCell = lstCapitalList.ListIndex + 131
      DepTagData(nTag, dTag).Title = LTrim(RTrim(txtCapitalItem.Text))
      DepTagData(nTag, dTag).Units = ""
      DepTagData(nTag, dTag).SetNumber = 1
    End If
  End If
  txtCapitalValues(0).SetFocus
End If

End Sub

Private Sub comIndTag_Click()

If labCheckTag.Visible = False Then
  ParamSet = False
  nTag = nTag + 1
  dTag = 0
  labCheckTag.Visible = True
  labCheckTag.ForeColor = &HFF&
  labCheckTag.Caption = LTrim(Str(nTag))
  If AddedYet = True Then
    Tagged(1, lstCapitalList.ListIndex + 131).Independent = nTag
    IndTagData(nTag).TheCell = lstCapitalList.ListIndex + 131
    IndTagData(nTag).Title = LTrim(RTrim(txtCapitalItem.Text))
    IndTagData(nTag).Units = ""
    IndTagData(nTag).SetNumber = 1
  End If
End If

txtCapitalValues(0).SetFocus

End Sub

Private Sub Form_Activate()
  
  Dim i As Integer
  Dim tempcap As Currency
  Dim tempsalv As Currency
  
  If IsHelpOn = True Then
    If LastCell = 0 Then
      frmCapital.txtCapitalItem.SetFocus
    ElseIf LastCell < 5 Then
      frmCapital.txtCapitalValues(LastCell - 1).SetFocus
    ElseIf LastCell > 6 Then
      cmdCapitalList(LastCell - 7).SetFocus
    End If
    IsHelpOn = False
  Else
    AddedYet = False
    ShowMenu = True
    DoNotChange = True
      lstCapitalList.ListIndex = -1
      txtCapitalItem.Text = ""
      labCheckTag.Caption = ""
      For i = 0 To 5
        txtCapitalValues(i).Text = ""
      Next i
      For i = 0 To NumCap - 1
        If CapitalData(i).PurchaseAmount > 0 Then
          tempcap = tempcap + CapitalData(i).PurchaseAmount
          tempsalv = tempsalv + CapitalData(i).SalvageAmount
        End If
      Next i
      txtCapitalValues(6).Text = Format(Str(tempcap), "###,###,###,###")
      txtCapitalValues(7).Text = Format(Str(tempsalv), "###,###,###,###")
    DoNotChange = False
    LastCell = 0
    
    If InsertFlag = True Then
      labInsert.Caption = "Insert"
    Else
      labInsert.Caption = "Typeover"
    End If
    
    frmCapital.txtCapitalItem.SetFocus

  End If
  
End Sub

Private Sub Form_Deactivate()
 
  Dim ind As Integer
  
  ind = lstCapitalList.ListCount

  CapitalData(ind).Item = ""
  CapitalData(ind).PurchaseAmount = 0
  CapitalData(ind).InvestYear = 0
  CapitalData(ind).DepMethod = 0
  CapitalData(ind).DepMethod = ""
  CapitalData(ind).DepPeriod = 0
  CapitalData(ind).SoldYear = 0
  CapitalData(ind).SalvageAmount = 0
  
  If ShowMenu = True Then
    frmCapital.Hide
    Call InputMenuAccess(1)
  End If
  
End Sub

Private Sub Form_Load()

Dim i As Integer
Dim tempcap As Currency
Dim tempsalv As Currency

tempcap = 0
tempsalv = 0
NumCap = 0

If FullScreen = False Then
  frmCapital.Top = (Screen.Height - (frmCapital.Height + 350)) / 2
  frmCapital.Left = (Screen.Width - frmCapital.Width) / 2
Else
  frmCapital.Top = 0
  frmCapital.Left = 0
  frmCapital.WindowState = 2
End If

If frmCapital.Top < 0 Then frmCapital.Top = 0
If frmCapital.Left < 0 Then frmCapital.Left = 0

tempwide = frmCapital.ScaleWidth
temphigh = frmCapital.ScaleHeight

If PageChange(5) = True Then
  For i = 0 To 40
    If CapitalData(i).PurchaseAmount > 0 Then
      lstCapitalList.AddItem LTrim(CapitalData(i).Item)
      NumCap = NumCap + 1
    End If
  Next i
  For i = 0 To NumCap
    tempcap = tempcap + CapitalData(i).PurchaseAmount
    tempsalv = tempsalv + CapitalData(i).SalvageAmount
  Next i
  txtCapitalValues(6).Text = Format(Str(tempcap), "###,###,###,###")
  txtCapitalValues(7).Text = Format(Str(tempsalv), "###,###,###,###")
End If

Call screenstuff

End Sub
Private Sub Form_Resize()

tempwide = frmCapital.ScaleWidth
temphigh = frmCapital.ScaleHeight

Call screenstuff

End Sub
Private Sub Form_Unload(Cancel As Integer)

  Dim ind As Integer
  
  ind = lstCapitalList.ListCount

  CapitalData(ind).Item = ""
  CapitalData(ind).PurchaseAmount = 0
  CapitalData(ind).InvestYear = 0
  CapitalData(ind).DepMethod = 0
  CapitalData(ind).DepMethod = ""
  CapitalData(ind).DepPeriod = 0
  CapitalData(ind).SoldYear = 0
  CapitalData(ind).SalvageAmount = 0
  
  frmCapital.Hide
  If ShowMenu = True Then Call InputMenuAccess(1)

End Sub

Private Sub imgBackToMenu_Click()
  
  Dim ind As Integer
  
  ind = lstCapitalList.ListCount

  CapitalData(ind).Item = ""
  CapitalData(ind).PurchaseAmount = 0
  CapitalData(ind).InvestYear = 0
  CapitalData(ind).DepMethod = 0
  CapitalData(ind).DepMethod = ""
  CapitalData(ind).DepPeriod = 0
  CapitalData(ind).SoldYear = 0
  CapitalData(ind).SalvageAmount = 0
  
  frmCapital.Hide
  If ShowMenu = True Then Call InputMenuAccess(1)

End Sub

Private Sub labBackToMenu_Click()

  Dim ind As Integer
  
  ind = lstCapitalList.ListCount

  CapitalData(ind).Item = ""
  CapitalData(ind).PurchaseAmount = 0
  CapitalData(ind).InvestYear = 0
  CapitalData(ind).DepMethod = 0
  CapitalData(ind).DepMethod = ""
  CapitalData(ind).DepPeriod = 0
  CapitalData(ind).SoldYear = 0
  CapitalData(ind).SalvageAmount = 0

  frmCapital.Hide
  If ShowMenu = True Then Call InputMenuAccess(1)

End Sub

Private Sub labCapitalHelp_Click()
Dim begin As Integer
Dim sendindex As Integer
ShowMenu = False
begin = 131

If LastCell < 5 Then
  sendindex = LastCell
ElseIf LastCell = 5 Then
  sendindex = LastCell + 1
ElseIf LastCell > 6 Then
  sendindex = LastCell
Else
  sendindex = LastCell - 1
End If

WhichScreen = 5

Call frmApexHelp.gethelptext(sendindex, begin)
frmApexHelp.Show

End Sub

Private Sub lstCapitalList_Click()

  Dim theword As String
  
  If DoNotChange = True Then Exit Sub
  AddedYet = True
  
  txtCapitalItem.Text = lstCapitalList.List(lstCapitalList.ListIndex)
  txtCapitalValues(0).Text = CapitalData(lstCapitalList.ListIndex).PurchaseAmount
  txtCapitalValues(1).Text = (CapitalData(lstCapitalList.ListIndex).InvestYear + Sets(12) - 1)
  Call findaword(lstCapitalList.ListIndex, theword)
  txtCapitalValues(2).Text = theword
  If Left(LCase(LTrim(theword)), 3) = "dim" Then
    txtSpecialDepreciate.Visible = True
    txtSpecialDepreciate.Text = LTrim(RTrim(Str(CapitalData(lstCapitalList.ListIndex).DmRate)))
  Else
    txtSpecialDepreciate.Visible = False
  End If
  txtCapitalValues(3).Text = LTrim(RTrim(Str(CapitalData(lstCapitalList.ListIndex).DepPeriod)))
  txtCapitalValues(4).Text = LTrim(RTrim(Str((CapitalData(lstCapitalList.ListIndex).SoldYear + Sets(12) - 1))))
  txtCapitalValues(5).Text = Format(LTrim(RTrim(Str(CapitalData(lstCapitalList.ListIndex).SalvageAmount))), "############")
  
  If Tagged(1, lstCapitalList.ListIndex + 131).Independent > 0 Then
    labCheckTag.Visible = True
    labCheckTag.ForeColor = &HFF&
    labCheckTag.Caption = LTrim(Str(Tagged(1, lstCapitalList.ListIndex + 131).Independent))
  ElseIf Tagged(1, lstCapitalList.ListIndex + 131).Dependent > 0 Then
    labCheckTag.Visible = True
    labCheckTag.ForeColor = &HFFFF&
    labCheckTag.Caption = LTrim(Str(Tagged(1, lstCapitalList.ListIndex + 131).Dependent))
  Else
    labCheckTag.Visible = False
  End If

  txtCapitalItem.SetFocus

End Sub

Private Sub lstDepreciationList_Click()

If Left(LCase(LTrim(lstDepreciationList.List(lstDepreciationList.ListIndex))), 3) = "mod" Then
  txtCapitalValues(2).Text = "Modified Accelerated Cost Recovery System"
Else
  txtCapitalValues(2).Text = lstDepreciationList.List(lstDepreciationList.ListIndex)
End If

If Left(LCase(LTrim(lstDepreciationList.List(lstDepreciationList.ListIndex))), 3) = "dev" Then
  txtCapitalValues(3).Text = 5
End If

If Left(LCase(LTrim(lstDepreciationList.List(lstDepreciationList.ListIndex))), 3) = "dim" Then
  txtSpecialDepreciate.Visible = True
Else
  txtSpecialDepreciate.Visible = False
End If

txtCapitalValues(3).SetFocus

End Sub

Private Sub txtCapitalItem_GotFocus()
  
  lstDepreciationList.Visible = False
  
  LastCell = 0

End Sub

Private Sub txtCapitalValues_Change(Index As Integer)

Dim ind As Integer
Dim otherind As Integer
Dim argh As Currency

If DoNotChange = True Then Exit Sub

If Index = 0 Then
  If labCheckTag.Visible = True Then ParamSet = False
End If

If Index = 1 Then
  If Val(txtCapitalValues(1).Text) < (Sets(12) - 1) Or Val(txtCapitalValues(1).Text) > (50 + Sets(12) - 1) Then
    Exit Sub
  End If
ElseIf Index = 4 Then
  If Val(txtCapitalValues(4).Text) < (Sets(12) - 1) Or Val(txtCapitalValues(4).Text) > (50 + Sets(12) - 1) Then
    Exit Sub
  End If
End If

If Index = 4 Then
  DoNotChange = True
    If Val(txtCapitalValues(4).Text) > (50 + Sets(12) - 1) Then txtCapitalValues(4).Text = Str(50 + Sets(12) - 1)
  DoNotChange = False
End If

If Index < 6 Then
   If CInt(Val(txtCapitalValues(3).Text)) > 50 Then
     txtCapitalValues(3).Text = Str(50 - CInt(Val(txtCapitalValues(1).Text)))
  End If
  ind = lstCapitalList.ListCount
  If Index = 4 Or Index = 5 Then
    CapitalData(ind).Changed = True
    ChangedFlag = True
  End If
  CapitalData(ind).Item = LTrim(txtCapitalItem.Text)
  CapitalData(ind).PurchaseAmount = CCur(Val(txtCapitalValues(0).Text))
  CapitalData(ind).InvestYear = CInt(Val(txtCapitalValues(1).Text)) - Sets(12) + 1
  CapitalData(ind).DepMethod = Left(LTrim(LCase(txtCapitalValues(2).Text)), 3)
  If Left(LTrim(LCase(txtCapitalValues(2).Text)), 10) = "straight a" Then
    CapitalData(ind).DepMethod = "amo"
  End If
  CapitalData(ind).DepPeriod = CInt(Val(txtCapitalValues(3).Text))
  CapitalData(ind).SoldYear = CInt(Val(txtCapitalValues(4).Text)) - Sets(12) + 1
  CapitalData(ind).SalvageAmount = CCur(Val(txtCapitalValues(5).Text))
  NumCap = NumCap + 1
  If Val(txtCapitalValues(1).Text) >= (Sets(12) - 1) And Val(txtCapitalValues(1).Text) <= (50 + Sets(12) - 1) Then Call cflow5(1, 0)
  NumCap = NumCap - 1
  DoNotChange = True
    txtCapitalValues(4).Text = LTrim(RTrim(Str(CapitalData(ind).SoldYear))) + Sets(12) - 1
    txtCapitalValues(5).Text = Format(LTrim(RTrim(Str(CapitalData(ind).SalvageAmount))), "############")
  DoNotChange = False
End If

End Sub

Private Sub txtCapitalValues_GotFocus(Index As Integer)

If Index = 2 Then
  lstDepreciationList.Visible = True
Else
  lstDepreciationList.Visible = False
End If
  
LastCell = Index + 1

End Sub
Public Sub screenstuff()
  
  Dim X As Integer
  Dim Y As Currency
  
  labCapitalHeading.Top = temphigh * 0.0187
  labCapitalHeading.Left = tempwide * 0.0131
  
  linBox1Top.X1 = tempwide * 0.0721
  linBox1Top.X2 = tempwide * 0.5508
  linBox1Top.Y1 = temphigh * 0.1028
  linBox1Top.Y2 = temphigh * 0.1028
  
  linBox1Middle.X1 = tempwide * 0.0852
  linBox1Middle.X2 = tempwide * 0.5377
  linBox1Middle.Y1 = temphigh * 0.2149
  linBox1Middle.Y2 = temphigh * 0.2149
  
  linLeftBoxLast.X1 = tempwide * 0.0852
  linLeftBoxLast.X2 = tempwide * 0.5377
  linLeftBoxLast.Y1 = temphigh * 0.6075
  linLeftBoxLast.Y2 = temphigh * 0.6075
 
  linBox1Bottom.X1 = tempwide * 0.0262
  linBox1Bottom.X2 = tempwide * 0.5508
  linBox1Bottom.Y1 = temphigh * 0.8785
  linBox1Bottom.Y2 = temphigh * 0.8785
 
  linBox1Left.X1 = tempwide * 0.0327
  linBox1Left.X2 = tempwide * 0.0327
  linBox1Left.Y1 = temphigh * 0.0934
  linBox1Left.Y2 = temphigh * 0.8879

  linBox1Right.X1 = tempwide * 0.5443
  linBox1Right.X2 = tempwide * 0.5443
  linBox1Right.Y1 = temphigh * 0.0934
  linBox1Right.Y2 = temphigh * 0.8879
  
  linBox2Top.X1 = tempwide * 0.6033
  linBox2Top.X2 = tempwide * 0.9574
  linBox2Top.Y1 = temphigh * 0.0467
  linBox2Top.Y2 = temphigh * 0.0467
     
  linBox2Middle.X1 = tempwide * 0.0852
  linBox2Middle.X2 = tempwide * 0.5377
  linBox2Middle.Y1 = temphigh * 0.7757
  linBox2Middle.Y2 = temphigh * 0.7757

  linBox2Bottom.X1 = tempwide * 0.6033
  linBox2Bottom.X2 = tempwide * 0.954
  linBox2Bottom.Y1 = temphigh * 0.6542
  linBox2Bottom.Y2 = temphigh * 0.6542
  
  linBox2Left.X1 = tempwide * 0.6098
  linBox2Left.X2 = tempwide * 0.6098
  linBox2Left.Y1 = temphigh * 0.0374
  linBox2Left.Y2 = temphigh * 0.6636

  linBox2Right.X1 = tempwide * 0.9508
  linBox2Right.X2 = tempwide * 0.9508
  linBox2Right.Y1 = temphigh * 0.0374
  linBox2Right.Y2 = temphigh * 0.6636
  
  For X = 0 To 5
    labCapitalTitles(X).Top = (X * 0.0561 * temphigh) + (temphigh * 0.2617)
    labCapitalTitles(X).Left = tempwide * 0.0656
    labCapitalTitles(X).Width = tempwide * 0.1852
    txtCapitalValues(X).Top = (X * 0.0561 * temphigh) + (temphigh * 0.257)
    txtCapitalValues(X).Left = tempwide * 0.308
    txtCapitalValues(X).Width = tempwide * 0.159
    If X = 2 Then
      txtCapitalValues(X).Left = tempwide * 0.2689
      txtCapitalValues(X).Width = tempwide * 0.2442
    End If
  Next X

  For X = 0 To 1
    labCapitalMisc(X + 4).Top = (X * 0.0561 * temphigh) + (temphigh * 0.6542)
    labCapitalMisc(X + 4).Left = tempwide * 0.0656
    labCapitalMisc(X + 4).Width = tempwide * 0.1852
    txtCapitalValues(X + 6).Top = (X * 0.0561 * temphigh) + (temphigh * 0.6495)
    txtCapitalValues(X + 6).Left = tempwide * 0.3082
    txtCapitalValues(X + 6).Width = tempwide * 0.159
  Next X
  
  txtSpecialDepreciate.Top = temphigh * 0.3692
  txtSpecialDepreciate.Left = tempwide * 0.4656
  txtSpecialDepreciate.Width = tempwide * 0.0475
  
  labCapitalMisc(0).Top = temphigh * 0.0934
  labCapitalMisc(0).Left = tempwide * 0.0262
   
  labCapitalMisc(1).Top = temphigh * 0.0374
  labCapitalMisc(1).Left = tempwide * 0.5967
  
  labCapitalMisc(2).Top = temphigh * 0.2056
  labCapitalMisc(2).Left = tempwide * 0.0262
  
  labCapitalMisc(3).Top = temphigh * 0.5981
  labCapitalMisc(3).Left = tempwide * 0.0262
  
  labCapitalMisc(6).Top = temphigh * 0.2243
  labCapitalMisc(6).Left = tempwide * 0.4918
  labCapitalMisc(6).Width = tempwide * 0.0377
  
  For X = 7 To 10
    If X = 7 Then
      labCapitalMisc(X).Top = temphigh * 0.2617
    ElseIf X = 8 Then
      labCapitalMisc(X).Top = temphigh * 0.542
    ElseIf X = 9 Then
      labCapitalMisc(X).Top = temphigh * 0.6542
    Else
      labCapitalMisc(X).Top = temphigh * 0.7103
    End If
    labCapitalMisc(X).Left = tempwide * 0.2885
    labCapitalMisc(X).Width = tempwide * 0.0148
  Next X
    
  labCapitalMisc(11).Top = temphigh * 0.4299
  labCapitalMisc(11).Left = tempwide * 0.4721
    
  labCapitalMisc(12).Top = temphigh * 0.7663
  labCapitalMisc(12).Left = tempwide * 0.0262
    
  lstCapitalList.Top = temphigh * 0.0935
  lstCapitalList.Left = tempwide * 0.6295
  lstCapitalList.Height = temphigh * 0.535
  lstCapitalList.Width = tempwide * 0.3033
    
  lstDepreciationList.Top = temphigh * 0.6822
  lstDepreciationList.Left = tempwide * 0.577
  lstDepreciationList.Height = temphigh * 0.3037
  lstDepreciationList.Width = tempwide * 0.4148
  
  txtCapitalItem.Top = temphigh * 0.1402
  txtCapitalItem.Left = tempwide * 0.1115
  txtCapitalItem.Width = tempwide * 0.3557
    
  For X = 0 To 4
   cmdCapitalList(X).Left = tempwide * (0.0787 + (X * 0.0852))
   cmdCapitalList(X).Top = temphigh * 0.8131
   cmdCapitalList(X).Width = tempwide * 0.0803
  Next X
 
  comIndTag.Top = temphigh * 0.8988
  comIndTag.Left = tempwide * 0.0459
  
  labIndTag.Top = temphigh * 0.8972
  labIndTag.Left = tempwide * 0.0787
  
  comDepTag.Top = temphigh * 0.8988
  comDepTag.Left = tempwide * 0.3607
  
  labDepTag.Top = temphigh * 0.8972
  labDepTag.Left = tempwide * 0.3934
  
  labCheckTag.Top = temphigh * 0.2617
  labCheckTag.Left = tempwide * 0.4918
 
  labBackToMenu.Top = temphigh * 0.9532
  labBackToMenu.Left = tempwide * 0.0656

  imgBackToMenu.Top = temphigh * 0.9626
  imgBackToMenu.Left = tempwide * 0.0066
  imgBackToMenu.Width = tempwide * 0.0541
  
  labCapitalHelp.Top = temphigh * 0.9532
  labCapitalHelp.Left = tempwide * 0.4984
    
  labInsert.Top = temphigh * 0.9562
  labInsert.Left = tempwide * 0.2656
  labInsert.Width = tempwide * 0.1066

End Sub
Private Sub txtCapitalValues_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

If KeyCode = 45 Then
  If InsertFlag = True Then
    InsertFlag = False
    labInsert.Caption = "Typeover"
  Else
    InsertFlag = True
    labInsert.Caption = "Insert"
  End If
End If

If InsertFlag = False Then
  Select Case KeyCode
    Case 48 To 57, 190
      If KeyCode = 190 Then
        If InStr(txtCapitalValues(Index).Text, ".") = 0 Then
          SendKeys "{DELETE}", False
        End If
      Else
        SendKeys "{DELETE}", False
      End If
  End Select
End If

End Sub
Private Sub txtCapitalValues_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii = 46 Then
  If InStr(txtCapitalValues(Index).Text, ".") > 0 Then
    Beep
    KeyAscii = 0
  End If
End If

If KeyAscii = 44 Then
  Beep
  KeyAscii = 0
End If

End Sub
Private Sub txtCapitalValues_LostFocus(Index As Integer)

If Index = 1 And Val(txtCapitalValues(Index).Text) = 0 Then
    Beep
End If

End Sub


