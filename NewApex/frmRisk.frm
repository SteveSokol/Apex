VERSION 5.00
Begin VB.Form frmRisk 
   BackColor       =   &H00000000&
   Caption         =   "Risk Analysis"
   ClientHeight    =   6420
   ClientLeft      =   1080
   ClientTop       =   1455
   ClientWidth     =   9150
   FillColor       =   &H00404040&
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
   Begin VB.CheckBox chkStats 
      BackColor       =   &H00000000&
      Caption         =   "View Statistical Analysis"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6240
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Line linIterationMidRight 
      BorderColor     =   &H00FFFF00&
      X1              =   5820
      X2              =   8940
      Y1              =   1020
      Y2              =   1020
   End
   Begin VB.Label labIterVariables 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   5
      Left            =   7620
      TabIndex        =   78
      Top             =   3060
      Width           =   1275
   End
   Begin VB.Label labIterVariables 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   7620
      TabIndex        =   77
      Top             =   2520
      Width           =   1275
   End
   Begin VB.Label labIterVariables 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   7620
      TabIndex        =   76
      Top             =   2100
      Width           =   1275
   End
   Begin VB.Label labIterVariables 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   7620
      TabIndex        =   75
      Top             =   1680
      Width           =   1275
   End
   Begin VB.Label labIterVariables 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   7620
      TabIndex        =   74
      Top             =   1260
      Width           =   1275
   End
   Begin VB.Label labIterVariables 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   0
      Left            =   5700
      TabIndex        =   73
      Top             =   180
      Width           =   45
   End
   Begin VB.Label labIterDistribution 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   4500
      TabIndex        =   72
      Top             =   3600
      Width           =   1155
   End
   Begin VB.Label labIterDistribution 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   4500
      TabIndex        =   71
      Top             =   3360
      Width           =   1155
   End
   Begin VB.Label labIterDistribution 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   4500
      TabIndex        =   70
      Top             =   3120
      Width           =   1155
   End
   Begin VB.Label labIterDistribution 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   4500
      TabIndex        =   69
      Top             =   2880
      Width           =   1155
   End
   Begin VB.Label labIterDistribution 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   4500
      TabIndex        =   68
      Top             =   2640
      Width           =   1155
   End
   Begin VB.Label labIterDistribution 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   4500
      TabIndex        =   67
      Top             =   2400
      Width           =   1155
   End
   Begin VB.Label labIterDistribution 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   4500
      TabIndex        =   66
      Top             =   2160
      Width           =   1155
   End
   Begin VB.Label labIterDistribution 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   4500
      TabIndex        =   65
      Top             =   1920
      Width           =   1155
   End
   Begin VB.Label labIterDistribution 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   4500
      TabIndex        =   64
      Top             =   1680
      Width           =   1155
   End
   Begin VB.Label labIterDistribution 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   4500
      TabIndex        =   63
      Top             =   1440
      Width           =   1155
   End
   Begin VB.Label labIterDistribution 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   4500
      TabIndex        =   62
      Top             =   1200
      Width           =   1155
   End
   Begin VB.Label labIterValue 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   2520
      TabIndex        =   61
      Top             =   3600
      Width           =   1875
   End
   Begin VB.Label labIterValue 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   2520
      TabIndex        =   60
      Top             =   3360
      Width           =   1875
   End
   Begin VB.Label labIterValue 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   2520
      TabIndex        =   59
      Top             =   3120
      Width           =   1875
   End
   Begin VB.Label labIterValue 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   2520
      TabIndex        =   58
      Top             =   2880
      Width           =   1875
   End
   Begin VB.Label labIterValue 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   2520
      TabIndex        =   57
      Top             =   2640
      Width           =   1875
   End
   Begin VB.Label labIterValue 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   2520
      TabIndex        =   56
      Top             =   2400
      Width           =   1875
   End
   Begin VB.Label labIterValue 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   55
      Top             =   2160
      Width           =   1875
   End
   Begin VB.Label labIterValue 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   54
      Top             =   1920
      Width           =   1875
   End
   Begin VB.Label labIterValue 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   53
      Top             =   1680
      Width           =   1875
   End
   Begin VB.Label labIterValue 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   52
      Top             =   1440
      Width           =   1875
   End
   Begin VB.Label labIterValue 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   51
      Top             =   1200
      Width           =   1875
   End
   Begin VB.Label labIterItem 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   50
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label labIterItem 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   49
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label labIterItem 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   48
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label labIterItem 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   47
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label labIterItem 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   46
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label labIterItem 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   45
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label labIterItem 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   44
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label labIterItem 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   43
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label labIterItem 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   42
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label labIterItem 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   41
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label labIterItem 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   40
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label labIterationTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Discount Rate"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   5880
      TabIndex        =   39
      Top             =   3060
      Width           =   1635
   End
   Begin VB.Line linSeperatorTop 
      BorderColor     =   &H00FFFF00&
      X1              =   5760
      X2              =   5760
      Y1              =   660
      Y2              =   960
   End
   Begin VB.Line linSeperatorBottom 
      BorderColor     =   &H00FFFF00&
      X1              =   5760
      X2              =   5760
      Y1              =   1080
      Y2              =   3900
   End
   Begin VB.Line linIterationRight 
      BorderColor     =   &H00FFFF00&
      X1              =   9000
      X2              =   9000
      Y1              =   540
      Y2              =   4020
   End
   Begin VB.Line linIterationLeft 
      BorderColor     =   &H00FFFF00&
      X1              =   120
      X2              =   120
      Y1              =   540
      Y2              =   4020
   End
   Begin VB.Line linIterationBottom 
      BorderColor     =   &H00FFFF00&
      X1              =   60
      X2              =   9060
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line linIterationMiddle 
      BorderColor     =   &H00FFFF00&
      X1              =   180
      X2              =   5700
      Y1              =   1020
      Y2              =   1020
   End
   Begin VB.Line linIterationTop 
      BorderColor     =   &H00FFFF00&
      X1              =   60
      X2              =   9060
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label labIterationTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Internal ROR"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   5880
      TabIndex        =   38
      Top             =   2520
      Width           =   1635
   End
   Begin VB.Label labIterationTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Pay-Back Period"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   5880
      TabIndex        =   37
      Top             =   2100
      Width           =   1635
   End
   Begin VB.Label labIterationTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Net Present Values"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   5880
      TabIndex        =   36
      Top             =   1680
      Width           =   1635
   End
   Begin VB.Label labIterationTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Sum of Cash Flows"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   5880
      TabIndex        =   35
      Top             =   1260
      Width           =   1635
   End
   Begin VB.Label labIterationTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Results"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   6120
      TabIndex        =   34
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label labIterationTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Distribution"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   4500
      TabIndex        =   33
      Top             =   720
      Width           =   1155
   End
   Begin VB.Label labIterationTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Value Generated"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   32
      Top             =   720
      Width           =   1875
   End
   Begin VB.Label labIterationTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Item"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   31
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label labIterationTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Iteration Number:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   3960
      TabIndex        =   30
      Top             =   180
      Width           =   1515
   End
   Begin VB.Line linSummaryRight 
      BorderColor     =   &H00FFFF00&
      X1              =   8940
      X2              =   8940
      Y1              =   4020
      Y2              =   6060
   End
   Begin VB.Line linSummaryUpper 
      BorderColor     =   &H00FFFF00&
      X1              =   240
      X2              =   8880
      Y1              =   4380
      Y2              =   4380
   End
   Begin VB.Line linSummaryMiddle 
      BorderColor     =   &H00FFFF00&
      X1              =   240
      X2              =   8880
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line linSummaryBottom 
      BorderColor     =   &H00FFFF00&
      X1              =   120
      X2              =   9000
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line linSummaryLeft 
      BorderColor     =   &H00FFFF00&
      X1              =   180
      X2              =   180
      Y1              =   4020
      Y2              =   6060
   End
   Begin VB.Line linSummaryTop 
      BorderColor     =   &H00FFFF00&
      X1              =   120
      X2              =   9000
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label labVariableNote 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Note:  Print Screen or Refer to Analysis Parameters Screen for Range of Values"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   1320
      TabIndex        =   29
      Top             =   5760
      Width           =   6495
   End
   Begin VB.Label labRateWarning 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "* Multiple Rates of Return Possible"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2820
      TabIndex        =   28
      Top             =   5460
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label labSummaryTitles 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Internal ROR"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   7500
      TabIndex        =   27
      Top             =   4140
      Width           =   1335
   End
   Begin VB.Label labSummaryTitles 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Pay-Back Period"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   5880
      TabIndex        =   26
      Top             =   4140
      Width           =   1455
   End
   Begin VB.Label labSummaryTitles 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Net Present Value"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   4140
      TabIndex        =   25
      Top             =   4140
      Width           =   1575
   End
   Begin VB.Label labSummaryTitles 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Sum of Cash Flows"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   2280
      TabIndex        =   24
      Top             =   4140
      Width           =   1695
   End
   Begin VB.Label labSummaryTitles 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Item"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   300
      TabIndex        =   23
      Top             =   4140
      Width           =   1695
   End
   Begin VB.Label labStds 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   7500
      TabIndex        =   22
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label labStds 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   5880
      TabIndex        =   21
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label labStds 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   4140
      TabIndex        =   20
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label labStds 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   2280
      TabIndex        =   19
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label labMaxs 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   7500
      TabIndex        =   18
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label labMaxs 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   5880
      TabIndex        =   17
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label labMaxs 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   4140
      TabIndex        =   16
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label labMaxs 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   2280
      TabIndex        =   15
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label labMins 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   7500
      TabIndex        =   14
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label labMins 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   5880
      TabIndex        =   13
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label labMins 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   4140
      TabIndex        =   12
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label labMins 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   2280
      TabIndex        =   11
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label labMeans 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   7500
      TabIndex        =   10
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label labMeans 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   5880
      TabIndex        =   9
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label labMeans 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   4140
      TabIndex        =   8
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label labMeans 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   2280
      TabIndex        =   7
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label labSummaryItems 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Standard Deviations"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   300
      TabIndex        =   6
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label labSummaryItems 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Maximum Values"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   300
      TabIndex        =   5
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label labSummaryItems 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Minimum Values"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   300
      TabIndex        =   4
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label labSummaryItems 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Mean Values"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   300
      TabIndex        =   3
      Top             =   4440
      Width           =   1695
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
      Left            =   8160
      TabIndex        =   2
      Top             =   6075
      Width           =   675
   End
   Begin VB.Image imgBackToMenu 
      Height          =   195
      Left            =   60
      Picture         =   "frmRisk.frx":0000
      Stretch         =   -1  'True
      Top             =   6180
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
      TabIndex        =   1
      Top             =   6075
      Width           =   675
   End
   Begin VB.Label labRiskHeading 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Risk Analysis"
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
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2115
   End
End
Attribute VB_Name = "frmRisk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim temphigh As Integer
Dim tempwide As Integer

Private Sub chkStats_Click()

ShowMenu = False
frmRisk.Hide
frmRisk.chkStats.Value = 0
frmStats.Visible = True

End Sub

Private Sub Form_Activate()

Dim i As Integer

ShowMenu = True

If DontRisk = False Then
  
  For i = 0 To 4
    labIterVariables(i).Caption = ""
  Next i
 
  For i = 0 To 10
    labIterItem(i).Caption = ""
    labIterValue(i).Caption = ""
    labIterDistribution(i).Caption = ""
  Next i

  For i = 0 To 3
    labMeans(i).Caption = ""
    labMins(i).Caption = ""
    labMaxs(i).Caption = ""
    labStds(i).Caption = ""
  Next i

  Call gettherisk
  
End If

DontRisk = True

End Sub

Private Sub Form_Deactivate()
  
If ShowMenu = True Then
  frmRisk.Hide
  Call InputMenuAccess(2)
End If
  
End Sub

Private Sub Form_Load()

If FullScreen = False Then
  frmRisk.Top = (Screen.Height - (frmRisk.Height + 350)) / 2
  frmRisk.Left = (Screen.Width - frmRisk.Width) / 2
Else
  frmRisk.Top = 0
  frmRisk.Left = 0
  frmRisk.WindowState = 2
End If

If frmRisk.Top < 0 Then frmRisk.Top = 0
If frmRisk.Left < 0 Then frmRisk.Left = 0

tempwide = frmRisk.ScaleWidth
temphigh = frmRisk.ScaleHeight

Call screenstuff
 
End Sub

Private Sub Form_Resize()

tempwide = frmRisk.ScaleWidth
temphigh = frmRisk.ScaleHeight

Call screenstuff

End Sub


Private Sub Form_Unload(Cancel As Integer)
  
  frmRisk.Hide
  If ShowMenu = True Then Call InputMenuAccess(2)

End Sub

Private Sub imgBackToMenu_Click()

  frmRisk.Hide
  Call InputMenuAccess(2)

End Sub

Private Sub labBackToMenu_Click()
  
  frmRisk.Hide
  Call InputMenuAccess(2)

End Sub

Public Sub gettherisk()

Dim skewcount As Integer
Dim riskword As String * 20
Dim titel(4) As String
Dim tempunit(2500) As String
Dim oldvalue(51, 51) As Currency
Dim minnum(4) As Currency
Dim maxnum(4) As Currency
Dim meannum(4) As Currency
Dim rootnum(4) As Single
Dim sumnumone(4) As Single
Dim sumnumtwo(4) As Single
Dim stdnum(4) As Single
Dim totnum(4) As Currency
Dim setcount(2500) As Integer
Dim itemcount(2500) As Integer
Dim maxcount(51) As Integer
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim l As Integer
Dim n As Integer
Dim r As Integer
Dim X As Integer
Dim topcount As Integer
Dim depcount As Integer
Dim socf As Integer
Dim sopv As Integer
Dim sopb As Integer
Dim sorr As Integer
Dim txtfilenum As Integer
Dim prnfilenum As Integer
Dim oldsource As Currency
Dim newsource As Currency
Dim lowbound As Currency
Dim highbound As Currency
Dim xxx As Currency

skewcount = 0
socf = 1
sopv = 2
sopb = 3
sorr = 4
If riter = 0 Then
  riter = 1000
End If

txtfilenum = FreeFile
Open MainDir & "\risk.txt" For Output As #txtfilenum

prnfilenum = FreeFile
Open MainDir & "\risk.prn" For Output As #prnfilenum

For X = 1 To 4
  Select Case X
    Case 1
      titel(X) = "Sum of Cash Flows"
    Case 2
      titel(X) = "Net Present Value"
    Case 3
      titel(X) = "Pay-Back Period"
    Case 4
      titel(X) = "Internal Rate of Return"
  End Select
  Print #txtfilenum,
  Print #prnfilenum, Chr$(34); titel(X); Chr$(34); ",";
Next X

Print #txtfilenum, Format(riter, "########0")
Print #prnfilenum,
Print #prnfilenum, riter, Chr$(44);
Print #prnfilenum,

labIterVariables(5).Caption = Format(LTrim(RTrim(Str(Sets(25)))), "##0.00") & "%"

topcount = 0

For i = 1 To 10
  labIterItem(i).Caption = ""
  labIterItem(i).ForeColor = &HFF
  labIterDistribution(i).Caption = ""
Next i

For n = 1 To Npna
  For i = 1 To nTag
    depcount = 1
    For j = 1 To 180
      If j < 131 Then
        If Tagged(n, j).Independent = i Then
          oldvalue(i, 0) = Primary(n, j)
          If topcount < 11 Then
            labIterItem(topcount).Caption = LTrim(RTrim(IndTagData(i).Title))
            labIterDistribution(topcount).Caption = LTrim(RTrim(IndTagData(i).Distribution))
          End If
          tempunit(topcount) = LTrim(RTrim(IndTagData(i).Units))
          setcount(topcount) = n
          itemcount(topcount) = j
          topcount = topcount + 1
          For k = 1 To 25
            For l = 1 To 180
              If Tagged(k, l).Dependent = i Then
                If l < 131 Then
                  oldvalue(i, depcount) = Primary(DepTagData(i, depcount).SetNumber, l)
                  If topcount < 11 Then
                    labIterItem(topcount).Caption = LTrim(RTrim(DepTagData(i, depcount).Title))
                    labIterItem(topcount).ForeColor = &H80FFFF
                    labIterDistribution(topcount).Caption = LTrim(RTrim(DepTagData(i, depcount).Distribution))
                  End If
                  tempunit(topcount) = LTrim(RTrim(DepTagData(i, depcount).Units))
                  depcount = depcount + 1
                  setcount(topcount) = k
                  itemcount(topcount) = l
                  topcount = topcount + 1
                Else
                  oldvalue(i, depcount) = CapitalData(l - 131).PurchaseAmount
                  If topcount <= 11 Then
                    labIterItem(topcount).Caption = LTrim(RTrim(DepTagData(i, depcount).Title))
                    labIterItem(topcount).ForeColor = &H80FFFF
                    labIterDistribution(topcount).Caption = LTrim(RTrim(DepTagData(i, depcount).Distribution))
                  End If
                  tempunit(topcount) = LTrim(RTrim(DepTagData(i, depcount).Units))
                  depcount = depcount + 1
                  setcount(topcount) = k
                  itemcount(topcount) = l
                  topcount = topcount + 1
                End If
              End If
            Next l
          Next k
        End If
      Else
        If Tagged(n, j).Independent = i Then
          oldvalue(i, 0) = CapitalData(j - 131).PurchaseAmount
          If topcount < 11 Then
            labIterItem(topcount).Caption = LTrim(RTrim(IndTagData(i).Title))
            labIterDistribution(topcount).Caption = LTrim(RTrim(IndTagData(i).Distribution))
          End If
          tempunit(topcount) = LTrim(RTrim(IndTagData(i).Units))
          setcount(topcount) = n
          itemcount(topcount) = j
          topcount = topcount + 1
          For k = 1 To 25
            For l = 1 To 180
              If Tagged(k, l).Dependent = i Then
                If l < 131 Then
                  oldvalue(i, depcount) = Primary(DepTagData(i, depcount).SetNumber, l)
                  If topcount < 11 Then
                    labIterItem(topcount).Caption = LTrim(RTrim(DepTagData(i, depcount).Title))
                    labIterItem(topcount).ForeColor = &H80FFFF
                    labIterDistribution(topcount).Caption = LTrim(RTrim(DepTagData(i, depcount).Distribution))
                  End If
                  tempunit(topcount) = LTrim(RTrim(DepTagData(i, depcount).Units))
                  depcount = depcount + 1
                  setcount(topcount) = k
                  itemcount(topcount) = l
                  topcount = topcount + 1
                Else
                  oldvalue(i, depcount) = CapitalData(l - 131).PurchaseAmount
                  If topcount <= 11 Then
                    labIterItem(topcount).Caption = LTrim(RTrim(DepTagData(i, depcount).Title))
                    labIterItem(topcount).ForeColor = &H80FFFF
                    labIterDistribution(topcount).Caption = LTrim(RTrim(DepTagData(i, depcount).Distribution))
                  End If
                  tempunit(topcount) = LTrim(RTrim(DepTagData(i, depcount).Units))
                  depcount = depcount + 1
                  setcount(topcount) = k
                  itemcount(topcount) = l
                  topcount = topcount + 1
                End If
              End If
            Next l
          Next k
        End If
      End If
    Next j
    If maxcount(i) < depcount - 1 Then maxcount(i) = depcount - 1
  Next i
Next n

For X = 1 To 4
  sumnumone(X) = 0
  sumnumtwo(X) = 0
Next X

chkStats.Enabled = False
labBackToMenu.Enabled = False
imgBackToMenu.Enabled = False
labPrintScreen.Enabled = False

For r = 1 To CInt(riter)
  Call sleep(0.0001)
  depcount = 0
  For i = 1 To nTag
    For j = 0 To maxcount(i)
      If j = 0 Then
        oldsource = oldvalue(i, 0)
        lowbound = IndTagData(i).Minimum
        highbound = IndTagData(i).Maximum
        Call distselect(IndTagData(i).Distribution, lowbound, highbound, xxx, oldsource, skewcount)
        newsource = xxx
        depcount = depcount + 1
      Else
        If oldvalue(i, j) = 0 Then
          xxx = 0
        Else
          xxx = (newsource / oldsource) * oldvalue(i, j)
         depcount = depcount + 1
        End If
      End If
                
      If itemcount(depcount - 1) > 130 Then
        CapitalData((itemcount(depcount - 1) - 131)).PurchaseAmount = xxx
      Else
        Primary(setcount(depcount - 1), itemcount(depcount - 1)) = xxx
      End If
      If depcount - 1 < 11 Then labIterValue(depcount - 1).Caption = LTrim(RTrim(Str(xxx)))
    Next j
  Next i

  Call cflow5(1, 0)
  Call rateofreturn
  totnum(socf) = Pv0
  totnum(sopv) = Pv2
  totnum(sopb) = Pb
  totnum(sorr) = Rot * 100
  
  labIterVariables(0).Caption = LTrim(RTrim(Str(r)))
  
  For i = 1 To 4
    If i < 3 Then
      labIterVariables(i).Caption = Format(LTrim(RTrim(Str(totnum(i)))), "$###,###,###,###")
    ElseIf i = 3 Then
      labIterVariables(i).Caption = Format(LTrim(RTrim(Str(totnum(i)))), "##0.00") & " years"
    Else
      labIterVariables(i).Caption = Format(LTrim(RTrim(Str(totnum(i)))), "##0.00") & "%"
    End If
    sumnumone(i) = sumnumone(i) + totnum(i)
    sumnumtwo(i) = sumnumtwo(i) + totnum(i) ^ 2
  Next i
   
  Print #txtfilenum, Pv0, Pv2, Pb, Rot * 100
  Print #prnfilenum, Pv0; ","; Pv2; ","; Pb; ","; Rot * 100; ",";
  Print #prnfilenum,
  
  If r = 1 Then
    For i = 1 To 4
      minnum(i) = totnum(i)
      maxnum(i) = totnum(i)
    Next i
  Else
    For i = 1 To 4
      If totnum(i) < minnum(i) Then minnum(i) = totnum(i)
      If totnum(i) > maxnum(i) Then maxnum(i) = totnum(i)
    Next i
  End If
Next r

chkStats.Enabled = True
labBackToMenu.Enabled = True
imgBackToMenu.Enabled = True
labPrintScreen.Enabled = True

Close #txtfilenum
Close #prnfilenum

For i = 1 To 4
  meannum(i) = CCur(sumnumone(i) / CInt(riter))
  rootnum(i) = (riter * sumnumtwo(i) - sumnumone(i) ^ 2)
  If rootnum(i) < 0 Then rootnum(i) = 0
  stdnum(i) = (rootnum(i) / (riter * (riter - 1))) ^ 0.5
  If i < 3 Then
    labMins(i - 1).Caption = Format(LTrim(RTrim(Str(minnum(i)))), "$###,###,###,###")
    labMaxs(i - 1).Caption = Format(LTrim(RTrim(Str(maxnum(i)))), "$###,###,###,###")
    labMeans(i - 1).Caption = Format(LTrim(RTrim(Str(meannum(i)))), "$###,###,###,###")
    labStds(i - 1).Caption = Format(LTrim(RTrim(Str(stdnum(i)))), "$###,###,###,###")
  ElseIf i = 3 Then
    labMins(i - 1).Caption = Format(LTrim(RTrim(Str(minnum(i)))), "##0.00") & " years"
    labMaxs(i - 1).Caption = Format(LTrim(RTrim(Str(maxnum(i)))), "##0.00") & " years"
    labMeans(i - 1).Caption = Format(LTrim(RTrim(Str(meannum(i)))), "##0.00") & " years"
    labStds(i - 1).Caption = Format(LTrim(RTrim(Str(stdnum(i)))), "##0.00") & " years"
  Else
    labMins(i - 1).Caption = Format(LTrim(RTrim(Str(minnum(i)))), "##0.00") & "%"
    labMaxs(i - 1).Caption = Format(LTrim(RTrim(Str(maxnum(i)))), "##0.00") & "%"
    labMeans(i - 1).Caption = Format(LTrim(RTrim(Str(meannum(i)))), "##0.00") & "%"
    labStds(i - 1).Caption = Format(LTrim(RTrim(Str(stdnum(i)))), "##0.00") & "%"
  End If
Next i

For n = 1 To Npna
  For i = 1 To nTag
    depcount = 1
    For j = 1 To 180
      If j < 131 Then
        If Tagged(n, j).Independent = i Then
          Primary(n, j) = oldvalue(i, 0)
          For k = 1 To 25
            For l = 1 To 180
              If Tagged(k, l).Dependent = i Then
                If l < 131 Then
                  Primary(k, l) = oldvalue(i, depcount)
                  depcount = depcount + 1
                Else
                  CapitalData(l - 131).PurchaseAmount = oldvalue(i, depcount)
                  depcount = depcount + 1
                End If
              End If
            Next l
          Next k
        End If
      Else
        If Tagged(n, j).Independent = i Then
          CapitalData(j - 131).PurchaseAmount = oldvalue(i, 0)
          For k = 1 To 25
            For l = 1 To 180
              If Tagged(k, l).Dependent = i Then
                If l < 131 Then
                  Primary(k, l) = oldvalue(i, depcount)
                  depcount = depcount + 1
                Else
                  CapitalData(l - 131).PurchaseAmount = oldvalue(i, depcount)
                  depcount = depcount + 1
                End If
              End If
            Next l
          Next k
        End If
      End If
    Next j
  Next i
Next n

If BadRor = 2 Then labRateWarning.Visible = True

Erase titel
Erase tempunit
Erase oldvalue
Erase minnum
Erase maxnum
Erase meannum
Erase rootnum
Erase sumnumone
Erase sumnumtwo
Erase stdnum
Erase totnum
Erase setcount
Erase itemcount

End Sub
Public Sub distselect(Dist As String, lowbound As Currency, highbound As Currency, xxx As Currency, oldsource As Currency, skewcount As Integer)

Dim u1 As Single
Dim u2 As Double
Dim u3 As Single
Dim u4 As Double
Dim mean As Currency
Dim std As Currency
Dim rrr As Integer
Dim stdup As Currency
Dim stddn As Currency
Dim skew As Integer
Dim mult As Integer

u1 = 254
u3 = 2.405
Randomize Timer

Select Case LCase(Left(LTrim(Dist), 1))
  Case "u"
    rrr = Int(Rnd * (101))
    xxx = lowbound + (highbound - lowbound) * ((100 - rrr) / 100)
  Case "l"
    mean = (lowbound + highbound) / 2
    std = (mean - lowbound) / 3.8
    If u3 >= 2.40483 Then GoTo lg1
    xxx = u3
    u3 = 2.405
    GoTo lg2
lg1:
    Randomize Timer
    u3 = Rnd
    u3 = Sqr((-2! * Log(u3)))
    u4 = 6.2831852 * Rnd
    xxx = Abs(Log(Abs(u3 * Cos(u4))))
    u3 = Abs(Log(Abs(u3 * Sin(u4))))
lg2:
    xxx = lowbound + xxx * std
  Case "s"
    mean = oldsource
    stdup = (highbound - mean) / 3.8
    stddn = (mean - lowbound) / 3.8
    skew = (CInt(1000 * ((mean - lowbound) / (highbound - lowbound))) / 100)
    If skew = 0 Then skew = 1
    If u1 >= 254! Then GoTo skew1
    xxx = u1
    u1 = 255!
    GoTo skew2
skew1:
    Randomize Timer
    u1 = Rnd
    u1 = Sqr(-2! * Log(u1))
    u2 = 6.2831852 * Rnd
    xxx = u1 * Cos(u2)
    u1 = u1 * Sin(u2)
skew2:
    skewcount = skewcount + 1
    If skewcount / skew <= 1 Then
      mult = -1
      std = stddn
    Else
      mult = 1
      std = stdup
    End If
    xxx = mean + ((mult * Abs(xxx)) * std)
    If skewcount = 10 Then skewcount = 0
  Case Else
    mean = (lowbound + highbound) / 2
    std = (mean - lowbound) / 3.8
    If u1 >= 254! Then GoTo norm1
    xxx = u1
    u1 = 255!
    GoTo norm2
norm1:
    Randomize Timer
    u1 = Rnd
    u1 = Sqr(-2! * Log(u1))
    u2 = 6.2831853 * Rnd
    xxx = u1 * Cos(u2)
    u1 = u1 * Sin(u2)
norm2:
    xxx = mean + xxx * std
End Select

End Sub
Public Sub sleep(sngNumberOfSeconds As Single)

  Dim sngEndTime As Single
    sngEndTime = Timer + sngNumberOfSeconds
  Do
    DoEvents
  Loop Until Timer >= sngEndTime
  
End Sub

Public Sub screenstuff()

  Dim X As Integer
   
  labRiskHeading.Top = temphigh * 0.0164
  labRiskHeading.Height = temphigh * 0.0631
  labRiskHeading.Left = tempwide * 0.0262
  
  linIterationTop.X1 = tempwide * 0.0066
  linIterationTop.X2 = tempwide * 0.9902
  linIterationTop.Y1 = temphigh * 0.0935
  linIterationTop.Y2 = temphigh * 0.0935
  
  linSeperatorTop.X1 = tempwide * 0.6295
  linSeperatorTop.X2 = tempwide * 0.6295
  linSeperatorTop.Y1 = temphigh * 0.1028
  linSeperatorTop.Y2 = temphigh * 0.1495
  
  linIterationLeft.X1 = tempwide * 0.0131
  linIterationLeft.X2 = tempwide * 0.0131
  linIterationLeft.Y1 = temphigh * 0.0841
  linIterationLeft.Y2 = temphigh * 0.6262

  linIterationMiddle.X1 = tempwide * 0.0197
  linIterationMiddle.X2 = tempwide * 0.623
  linIterationMiddle.Y1 = temphigh * 0.1589
  linIterationMiddle.Y2 = temphigh * 0.1589
  
  linIterationMidRight.X1 = tempwide * 0.6361
  linIterationMidRight.X2 = tempwide * 0.977
  linIterationMidRight.Y1 = temphigh * 0.1589
  linIterationMidRight.Y2 = temphigh * 0.1589

  linIterationRight.X1 = tempwide * 0.9836
  linIterationRight.X2 = tempwide * 0.9836
  linIterationRight.Y1 = temphigh * 0.0841
  linIterationRight.Y2 = temphigh * 0.6262
  
  linSeperatorBottom.X1 = tempwide * 0.6295
  linSeperatorBottom.X2 = tempwide * 0.6295
  linSeperatorBottom.Y1 = temphigh * 0.1682
  linSeperatorBottom.Y2 = temphigh * 0.6075

  linIterationBottom.X1 = tempwide * 0.0066
  linIterationBottom.X2 = tempwide * 0.9902
  linIterationBottom.Y1 = temphigh * 0.6168
  linIterationBottom.Y2 = temphigh * 0.6168

  linSummaryTop.X1 = tempwide * 0.0131
  linSummaryTop.X2 = tempwide * 0.9836
  linSummaryTop.Y1 = temphigh * 0.6355
  linSummaryTop.Y2 = temphigh * 0.6355
  
  linSummaryBottom.X1 = tempwide * 0.0131
  linSummaryBottom.X2 = tempwide * 0.9836
  linSummaryBottom.Y1 = temphigh * 0.9346
  linSummaryBottom.Y2 = temphigh * 0.9346
  
  linSummaryUpper.X1 = tempwide * 0.0262
  linSummaryUpper.X2 = tempwide * 0.9705
  linSummaryUpper.Y1 = temphigh * 0.6822
  linSummaryUpper.Y2 = temphigh * 0.6822
  
  linSummaryMiddle.X1 = tempwide * 0.0262
  linSummaryMiddle.X2 = tempwide * 0.9705
  linSummaryMiddle.Y1 = temphigh * 0.8411
  linSummaryMiddle.Y2 = temphigh * 0.8411

  linSummaryLeft.X1 = tempwide * 0.0197
  linSummaryLeft.X2 = tempwide * 0.0197
  linSummaryLeft.Y1 = temphigh * 0.6262
  linSummaryLeft.Y2 = temphigh * 0.9439

  linSummaryRight.X1 = tempwide * 0.977
  linSummaryRight.X2 = tempwide * 0.977
  linSummaryRight.Y1 = temphigh * 0.6262
  linSummaryRight.Y2 = temphigh * 0.9439
  
  labIterationTitle(0).Top = temphigh * 0.028
  labIterationTitle(0).Left = tempwide * 0.4328
  labIterationTitle(0).Width = tempwide * 0.1656

  labIterVariables(0).Top = temphigh * 0.028
  labIterVariables(0).Left = tempwide * 0.623
  
  For X = 1 To 4
    labIterationTitle(X).Top = temphigh * 0.1121
    If X = 1 Then
      labIterationTitle(X).Left = tempwide * 0.0262
      labIterationTitle(X).Width = tempwide * 0.2377
    ElseIf X = 2 Then
      labIterationTitle(X).Left = tempwide * 0.2754
      labIterationTitle(X).Width = tempwide * 0.2049
    ElseIf X = 3 Then
      labIterationTitle(X).Left = tempwide * 0.4918
      labIterationTitle(X).Width = tempwide * 0.1262
    Else
      labIterationTitle(X).Left = tempwide * 0.6689
      labIterationTitle(X).Width = tempwide * 0.277
    End If
  Next X

  For X = 1 To 5
    If X < 5 Then
      labIterationTitle(X + 4).Top = ((X - 1) * 0.0654 * temphigh) + (temphigh * 0.1963)
      labIterVariables(X).Top = ((X - 1) * 0.0654 * temphigh) + (temphigh * 0.1963)
    Else
      labIterationTitle(X + 4).Top = temphigh * 0.4766
      labIterVariables(X).Top = temphigh * 0.4766
    End If
    labIterationTitle(X + 4).Left = tempwide * 0.6426
    labIterationTitle(X + 4).Width = tempwide * 0.1787
    labIterVariables(X).Left = tempwide * 0.8328
    labIterVariables(X).Width = tempwide * 0.1393
  Next X
  
  For X = 0 To 10
    labIterItem(X).Top = (X * 0.0374 * temphigh) + (temphigh * 0.1869)
    labIterValue(X).Top = (X * 0.0374 * temphigh) + (temphigh * 0.1869)
    labIterDistribution(X).Top = (X * 0.0374 * temphigh) + (temphigh * 0.1869)
    labIterItem(X).Left = tempwide * 0.0262
    labIterValue(X).Left = tempwide * 0.2754
    labIterDistribution(X).Left = tempwide * 0.4918
    labIterItem(X).Width = tempwide * 0.2377
    labIterValue(X).Width = tempwide * 0.2049
    labIterDistribution(X).Width = tempwide * 0.1262
  Next X

  For X = 0 To 4
    labSummaryTitles(X).Top = temphigh * 0.6449
    If X = 0 Then
      labSummaryTitles(X).Left = tempwide * 0.0328
      labSummaryTitles(X).Width = tempwide * 0.1852
    ElseIf X = 1 Then
      labSummaryTitles(X).Left = tempwide * 0.2492
      labSummaryTitles(X).Width = tempwide * 0.1852
    ElseIf X = 2 Then
      labSummaryTitles(X).Left = tempwide * 0.4525
      labSummaryTitles(X).Width = tempwide * 0.1721
    ElseIf X = 3 Then
      labSummaryTitles(X).Left = tempwide * 0.6426
      labSummaryTitles(X).Width = tempwide * 0.159
    Else
      labSummaryTitles(X).Left = tempwide * 0.8197
      labSummaryTitles(X).Width = tempwide * 0.1459
    End If
  Next X
  
  For X = 0 To 3
    labSummaryItems(X).Top = (X * 0.0374 * temphigh) + (temphigh * 0.6916)
    labSummaryItems(X).Left = tempwide * 0.0328
    labSummaryItems(X).Width = tempwide * 0.1852
    labMeans(X).Top = temphigh * 0.6916
    labMins(X).Top = temphigh * 0.729
    labMaxs(X).Top = temphigh * 0.7664
    labStds(X).Top = temphigh * 0.8037
    If X = 0 Then
      labMeans(X).Left = tempwide * 0.2492
      labMeans(X).Width = tempwide * 0.1852
      labMins(X).Left = tempwide * 0.2492
      labMins(X).Width = tempwide * 0.1852
      labMaxs(X).Left = tempwide * 0.2492
      labMaxs(X).Width = tempwide * 0.1852
      labStds(X).Left = tempwide * 0.2492
      labStds(X).Width = tempwide * 0.1852
    ElseIf X = 1 Then
      labMeans(X).Left = tempwide * 0.4525
      labMeans(X).Width = tempwide * 0.1721
      labMins(X).Left = tempwide * 0.4525
      labMins(X).Width = tempwide * 0.1721
      labMaxs(X).Left = tempwide * 0.4525
      labMaxs(X).Width = tempwide * 0.1721
      labStds(X).Left = tempwide * 0.4525
      labStds(X).Width = tempwide * 0.1721
    ElseIf X = 2 Then
      labMeans(X).Left = tempwide * 0.6426
      labMeans(X).Width = tempwide * 0.159
      labMins(X).Left = tempwide * 0.6426
      labMins(X).Width = tempwide * 0.159
      labMaxs(X).Left = tempwide * 0.6426
      labMaxs(X).Width = tempwide * 0.159
      labStds(X).Left = tempwide * 0.6426
      labStds(X).Width = tempwide * 0.159
    Else
      labMeans(X).Left = tempwide * 0.8197
      labMeans(X).Width = tempwide * 0.1459
      labMins(X).Left = tempwide * 0.8197
      labMins(X).Width = tempwide * 0.1459
      labMaxs(X).Left = tempwide * 0.8197
      labMaxs(X).Width = tempwide * 0.1459
      labStds(X).Left = tempwide * 0.8197
      labStds(X).Width = tempwide * 0.1459
    End If
  Next X
  
  chkStats.Top = temphigh * 0.5607
  chkStats.Left = (tempwide * 0.8074) - 1148
  
  labRateWarning.Top = temphigh * 0.8505
  labRateWarning.Left = tempwide * 0.3082
  labRateWarning.Width = tempwide * 0.382
  
  labVariableNote.Top = temphigh * 0.8972
  labVariableNote.Left = tempwide * 0.1422
  labVariableNote.Width = tempwide * 0.7104
  
  labBackToMenu.Top = temphigh * 0.9463
  labBackToMenu.Left = tempwide * 0.0656

  imgBackToMenu.Top = temphigh * 0.9532
  imgBackToMenu.Left = tempwide * 0.0066
  imgBackToMenu.Width = tempwide * 0.0541

  labPrintScreen.Top = temphigh * 0.9463
  labPrintScreen.Left = tempwide * 0.8918

End Sub

Private Sub labPrintScreen_Click()

ShowMenu = False
job = 13
Call printstuffout(job)

End Sub

