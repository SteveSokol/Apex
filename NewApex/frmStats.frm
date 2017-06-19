VERSION 5.00
Begin VB.Form frmStats 
   BackColor       =   &H00000000&
   Caption         =   "Statistical Analysis"
   ClientHeight    =   6420
   ClientLeft      =   1635
   ClientTop       =   1695
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
   Begin VB.OptionButton optDisplayType 
      BackColor       =   &H00000000&
      Caption         =   "Cumulative Probability"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   2
      Left            =   150
      TabIndex        =   6
      Top             =   4560
      Width           =   2155
   End
   Begin VB.OptionButton optDisplayType 
      BackColor       =   &H00000000&
      Caption         =   "Frequency Histogram"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   150
      TabIndex        =   5
      Top             =   4200
      Width           =   2115
   End
   Begin VB.OptionButton optDisplayType 
      BackColor       =   &H00000000&
      Caption         =   "Frequency Distribution"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   150
      TabIndex        =   4
      Top             =   3840
      Value           =   -1  'True
      Width           =   2155
   End
   Begin VB.CheckBox chkVariableType 
      BackColor       =   &H00000000&
      Caption         =   "Sum of Cash Flows"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CheckBox chkVariableType 
      BackColor       =   &H00000000&
      Caption         =   "Internal ROR"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CheckBox chkVariableType 
      BackColor       =   &H00000000&
      Caption         =   "Pay-Back Period"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   2040
      Width           =   1935
   End
   Begin VB.CheckBox chkVariableType 
      BackColor       =   &H00000000&
      Caption         =   "Net Present Value"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.Label labLastTitles 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Display"
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
      Index           =   1
      Left            =   840
      TabIndex        =   131
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label labLastTitles 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Analysis"
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
      Index           =   0
      Left            =   780
      TabIndex        =   130
      Top             =   1320
      Width           =   855
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
      Left            =   1320
      TabIndex        =   129
      Top             =   6075
      Width           =   615
   End
   Begin VB.Image imgBack 
      Height          =   195
      Left            =   60
      Picture         =   "frmStats.frx":0000
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   495
   End
   Begin VB.Label labBack 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Back"
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
      TabIndex        =   128
      Top             =   6075
      Width           =   615
   End
   Begin VB.Label labTo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   20
      Left            =   4020
      TabIndex        =   127
      Top             =   5520
      Width           =   1635
   End
   Begin VB.Label labFrom 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   20
      Left            =   2460
      TabIndex        =   126
      Top             =   5520
      Width           =   1395
   End
   Begin VB.Line linYTics 
      BorderColor     =   &H0000FFFF&
      Index           =   20
      Visible         =   0   'False
      X1              =   3900
      X2              =   3960
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Label labXLabel 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Percent Cumulative Probability"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3960
      TabIndex        =   125
      Top             =   6060
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Label labFreqTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Cumulative"
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
      Left            =   7800
      TabIndex        =   124
      Top             =   600
      Width           =   975
   End
   Begin VB.Label labFreqTitles 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Occurrences"
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
      Index           =   2
      Left            =   6420
      TabIndex        =   123
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label labFreqTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "To"
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
      Index           =   1
      Left            =   4140
      TabIndex        =   122
      Top             =   600
      Width           =   1155
   End
   Begin VB.Label labFreqTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "From"
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
      Index           =   0
      Left            =   2460
      TabIndex        =   121
      Top             =   600
      Width           =   1155
   End
   Begin VB.Line linMainTopMid 
      BorderColor     =   &H00FFFF00&
      X1              =   2460
      X2              =   8940
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label labTableTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Frequency Distribution Table - Net Present Value"
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
      Left            =   2520
      TabIndex        =   120
      Top             =   180
      Width           =   6375
   End
   Begin VB.Label labXTics 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   10
      Left            =   8580
      TabIndex        =   119
      Top             =   5820
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label labXTics 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "90"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   9
      Left            =   8100
      TabIndex        =   118
      Top             =   5820
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label labXTics 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "80"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   7620
      TabIndex        =   117
      Top             =   5820
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label labXTics 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "70"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   7140
      TabIndex        =   116
      Top             =   5820
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label labXTics 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "60"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   6660
      TabIndex        =   115
      Top             =   5820
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label labXTics 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   6180
      TabIndex        =   114
      Top             =   5820
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label labXTics 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "40"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   5700
      TabIndex        =   113
      Top             =   5820
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label labXTics 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "30"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   5220
      TabIndex        =   112
      Top             =   5820
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label labXTics 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   4740
      TabIndex        =   111
      Top             =   5820
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label labXTics 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   4260
      TabIndex        =   110
      Top             =   5820
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label labXTics 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   3780
      TabIndex        =   109
      Top             =   5820
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Line linXTic 
      BorderColor     =   &H0000FFFF&
      Index           =   9
      Visible         =   0   'False
      X1              =   8760
      X2              =   8760
      Y1              =   5760
      Y2              =   5820
   End
   Begin VB.Line linXTic 
      BorderColor     =   &H0000FFFF&
      Index           =   8
      Visible         =   0   'False
      X1              =   8280
      X2              =   8280
      Y1              =   5760
      Y2              =   5820
   End
   Begin VB.Line linXTic 
      BorderColor     =   &H0000FFFF&
      Index           =   7
      Visible         =   0   'False
      X1              =   7800
      X2              =   7800
      Y1              =   5760
      Y2              =   5820
   End
   Begin VB.Line linXTic 
      BorderColor     =   &H0000FFFF&
      Index           =   6
      Visible         =   0   'False
      X1              =   7320
      X2              =   7320
      Y1              =   5760
      Y2              =   5820
   End
   Begin VB.Line linXTic 
      BorderColor     =   &H0000FFFF&
      Index           =   5
      Visible         =   0   'False
      X1              =   6840
      X2              =   6840
      Y1              =   5760
      Y2              =   5820
   End
   Begin VB.Line linXTic 
      BorderColor     =   &H0000FFFF&
      Index           =   4
      Visible         =   0   'False
      X1              =   6360
      X2              =   6360
      Y1              =   5760
      Y2              =   5820
   End
   Begin VB.Line linXTic 
      BorderColor     =   &H0000FFFF&
      Index           =   3
      Visible         =   0   'False
      X1              =   5880
      X2              =   5880
      Y1              =   5760
      Y2              =   5820
   End
   Begin VB.Line linXTic 
      BorderColor     =   &H0000FFFF&
      Index           =   2
      Visible         =   0   'False
      X1              =   5400
      X2              =   5400
      Y1              =   5760
      Y2              =   5820
   End
   Begin VB.Line linXTic 
      BorderColor     =   &H0000FFFF&
      Index           =   1
      Visible         =   0   'False
      X1              =   4920
      X2              =   4920
      Y1              =   5760
      Y2              =   5820
   End
   Begin VB.Line linXTic 
      BorderColor     =   &H0000FFFF&
      Index           =   0
      Visible         =   0   'False
      X1              =   4440
      X2              =   4440
      Y1              =   5760
      Y2              =   5820
   End
   Begin VB.Label labFrom 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   2460
      TabIndex        =   108
      Top             =   960
      Width           =   1395
   End
   Begin VB.Line linYTics 
      BorderColor     =   &H0000FFFF&
      Index           =   19
      Visible         =   0   'False
      X1              =   3900
      X2              =   3960
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line linYTics 
      BorderColor     =   &H0000FFFF&
      Index           =   18
      Visible         =   0   'False
      X1              =   3900
      X2              =   3960
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line linYTics 
      BorderColor     =   &H0000FFFF&
      Index           =   17
      Visible         =   0   'False
      X1              =   3900
      X2              =   3960
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line linYTics 
      BorderColor     =   &H0000FFFF&
      Index           =   16
      Visible         =   0   'False
      X1              =   3900
      X2              =   3960
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line linYTics 
      BorderColor     =   &H0000FFFF&
      Index           =   15
      Visible         =   0   'False
      X1              =   3900
      X2              =   3960
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line linYTics 
      BorderColor     =   &H0000FFFF&
      Index           =   14
      Visible         =   0   'False
      X1              =   3900
      X2              =   3960
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line linYTics 
      BorderColor     =   &H0000FFFF&
      Index           =   13
      Visible         =   0   'False
      X1              =   3900
      X2              =   3960
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line linYTics 
      BorderColor     =   &H0000FFFF&
      Index           =   12
      Visible         =   0   'False
      X1              =   3900
      X2              =   3960
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line linYTics 
      BorderColor     =   &H0000FFFF&
      Index           =   11
      Visible         =   0   'False
      X1              =   3900
      X2              =   3960
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line linYTics 
      BorderColor     =   &H0000FFFF&
      Index           =   10
      Visible         =   0   'False
      X1              =   3900
      X2              =   3960
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line linYTics 
      BorderColor     =   &H0000FFFF&
      Index           =   9
      Visible         =   0   'False
      X1              =   3900
      X2              =   3960
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line linYTics 
      BorderColor     =   &H0000FFFF&
      Index           =   8
      Visible         =   0   'False
      X1              =   3900
      X2              =   3960
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line linYTics 
      BorderColor     =   &H0000FFFF&
      Index           =   7
      Visible         =   0   'False
      X1              =   3900
      X2              =   3960
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line linYTics 
      BorderColor     =   &H0000FFFF&
      Index           =   6
      Visible         =   0   'False
      X1              =   3900
      X2              =   3960
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line linYTics 
      BorderColor     =   &H0000FFFF&
      Index           =   5
      Visible         =   0   'False
      X1              =   3900
      X2              =   3960
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line linYTics 
      BorderColor     =   &H0000FFFF&
      Index           =   4
      Visible         =   0   'False
      X1              =   3900
      X2              =   3960
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line linYTics 
      BorderColor     =   &H0000FFFF&
      Index           =   3
      Visible         =   0   'False
      X1              =   3900
      X2              =   3960
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line linYTics 
      BorderColor     =   &H0000FFFF&
      Index           =   2
      Visible         =   0   'False
      X1              =   3900
      X2              =   3960
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line linYTics 
      BorderColor     =   &H0000FFFF&
      Index           =   1
      Visible         =   0   'False
      X1              =   3900
      X2              =   3960
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line linYTics 
      BorderColor     =   &H0000FFFF&
      Index           =   0
      Visible         =   0   'False
      X1              =   3900
      X2              =   3960
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line linGraphX 
      BorderColor     =   &H0000FFFF&
      Visible         =   0   'False
      X1              =   3960
      X2              =   8760
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line linGraphY 
      BorderColor     =   &H0000FFFF&
      Visible         =   0   'False
      X1              =   3960
      X2              =   3960
      Y1              =   720
      Y2              =   5820
   End
   Begin VB.Line linMainBottom 
      BorderColor     =   &H00FFFF00&
      X1              =   2340
      X2              =   9060
      Y1              =   6300
      Y2              =   6300
   End
   Begin VB.Line linMainRight 
      BorderColor     =   &H00FFFF00&
      X1              =   9000
      X2              =   9000
      Y1              =   60
      Y2              =   6360
   End
   Begin VB.Line linMainTop 
      BorderColor     =   &H00FFFF00&
      X1              =   2340
      X2              =   9060
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line linMainLeft 
      BorderColor     =   &H00FFFF00&
      X1              =   2400
      X2              =   2400
      Y1              =   60
      Y2              =   6360
   End
   Begin VB.Label labStatsHeading2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Analysis"
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
      Left            =   660
      TabIndex        =   107
      Top             =   540
      Width           =   1575
   End
   Begin VB.Line linAnalysisRight 
      BorderColor     =   &H00FFFF00&
      X1              =   2340
      X2              =   2340
      Y1              =   3540
      Y2              =   4980
   End
   Begin VB.Line linAnalysisBottom 
      BorderColor     =   &H00FFFF00&
      X1              =   0
      X2              =   2400
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line linAnalysisLeft 
      BorderColor     =   &H00FFFF00&
      X1              =   60
      X2              =   60
      Y1              =   3540
      Y2              =   4980
   End
   Begin VB.Line linAnalysisTop 
      BorderColor     =   &H00FFFF00&
      X1              =   0
      X2              =   2400
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line linVaryRight 
      BorderColor     =   &H00FFFF00&
      X1              =   2280
      X2              =   2280
      Y1              =   1380
      Y2              =   3180
   End
   Begin VB.Line linVaryBottom 
      BorderColor     =   &H00FFFF00&
      X1              =   60
      X2              =   2340
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line linVaryLeft 
      BorderColor     =   &H00FFFF00&
      X1              =   120
      X2              =   120
      Y1              =   1380
      Y2              =   3180
   End
   Begin VB.Label labCumPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   19
      Left            =   7800
      TabIndex        =   106
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label labCumPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   18
      Left            =   7800
      TabIndex        =   105
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label labCumPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   17
      Left            =   7800
      TabIndex        =   104
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label labCumPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   16
      Left            =   7800
      TabIndex        =   103
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label labCumPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   15
      Left            =   7800
      TabIndex        =   102
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label labCumPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   14
      Left            =   7800
      TabIndex        =   101
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label labCumPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   13
      Left            =   7800
      TabIndex        =   100
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label labCumPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   12
      Left            =   7800
      TabIndex        =   99
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label labCumPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   11
      Left            =   7800
      TabIndex        =   98
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label labCumPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   10
      Left            =   7800
      TabIndex        =   97
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label labCumPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   9
      Left            =   7800
      TabIndex        =   96
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label labCumPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   7800
      TabIndex        =   95
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label labCumPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   7800
      TabIndex        =   94
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label labCumPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   7800
      TabIndex        =   93
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label labCumPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   7800
      TabIndex        =   92
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label labCumPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   7800
      TabIndex        =   91
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label labCumPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   7800
      TabIndex        =   90
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label labCumPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   7800
      TabIndex        =   89
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label labCumPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   7800
      TabIndex        =   88
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label labCumPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   7800
      TabIndex        =   87
      Top             =   960
      Width           =   975
   End
   Begin VB.Label labPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   19
      Left            =   6780
      TabIndex        =   86
      Top             =   5520
      Width           =   735
   End
   Begin VB.Label labPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   18
      Left            =   6780
      TabIndex        =   85
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label labPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   17
      Left            =   6780
      TabIndex        =   84
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label labPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   16
      Left            =   6780
      TabIndex        =   83
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label labPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   15
      Left            =   6780
      TabIndex        =   82
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label labPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   14
      Left            =   6780
      TabIndex        =   81
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label labPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   13
      Left            =   6780
      TabIndex        =   80
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label labPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   12
      Left            =   6780
      TabIndex        =   79
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label labPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   11
      Left            =   6780
      TabIndex        =   78
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label labPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   10
      Left            =   6780
      TabIndex        =   77
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label labPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   9
      Left            =   6780
      TabIndex        =   76
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label labPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   6780
      TabIndex        =   75
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label labPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   6780
      TabIndex        =   74
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label labPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   6780
      TabIndex        =   73
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label labPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   6780
      TabIndex        =   72
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label labPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   6780
      TabIndex        =   71
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label labPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   6780
      TabIndex        =   70
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label labPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   6780
      TabIndex        =   69
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label labPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   6780
      TabIndex        =   68
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label labPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   6780
      TabIndex        =   67
      Top             =   960
      Width           =   735
   End
   Begin VB.Label labOccur 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   19
      Left            =   6000
      TabIndex        =   66
      Top             =   5520
      Width           =   615
   End
   Begin VB.Label labOccur 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   18
      Left            =   6000
      TabIndex        =   65
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label labOccur 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   17
      Left            =   6000
      TabIndex        =   64
      Top             =   5040
      Width           =   615
   End
   Begin VB.Label labOccur 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   16
      Left            =   6000
      TabIndex        =   63
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label labOccur 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   15
      Left            =   6000
      TabIndex        =   62
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label labOccur 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   14
      Left            =   6000
      TabIndex        =   61
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label labOccur 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   13
      Left            =   6000
      TabIndex        =   60
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label labOccur 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   12
      Left            =   6000
      TabIndex        =   59
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label labOccur 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   11
      Left            =   6000
      TabIndex        =   58
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label labOccur 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   10
      Left            =   6000
      TabIndex        =   57
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label labOccur 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   9
      Left            =   6000
      TabIndex        =   56
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label labOccur 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   6000
      TabIndex        =   55
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label labOccur 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   6000
      TabIndex        =   54
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label labOccur 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   6000
      TabIndex        =   53
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label labOccur 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   6000
      TabIndex        =   52
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label labOccur 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   6000
      TabIndex        =   51
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label labOccur 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   6000
      TabIndex        =   50
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label labOccur 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   6000
      TabIndex        =   49
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label labOccur 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   6000
      TabIndex        =   48
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label labOccur 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   6000
      TabIndex        =   47
      Top             =   960
      Width           =   615
   End
   Begin VB.Line linVaryTop 
      BorderColor     =   &H00FFFF00&
      X1              =   60
      X2              =   2340
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label labTo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   19
      Left            =   4020
      TabIndex        =   46
      Top             =   5280
      Width           =   1635
   End
   Begin VB.Label labTo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   18
      Left            =   4020
      TabIndex        =   45
      Top             =   5040
      Width           =   1635
   End
   Begin VB.Label labTo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   17
      Left            =   4020
      TabIndex        =   44
      Top             =   4800
      Width           =   1635
   End
   Begin VB.Label labTo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   16
      Left            =   4020
      TabIndex        =   43
      Top             =   4560
      Width           =   1635
   End
   Begin VB.Label labTo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   15
      Left            =   4020
      TabIndex        =   42
      Top             =   4320
      Width           =   1635
   End
   Begin VB.Label labTo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   14
      Left            =   4020
      TabIndex        =   41
      Top             =   4080
      Width           =   1635
   End
   Begin VB.Label labTo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   13
      Left            =   4020
      TabIndex        =   40
      Top             =   3840
      Width           =   1635
   End
   Begin VB.Label labTo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   12
      Left            =   4020
      TabIndex        =   39
      Top             =   3600
      Width           =   1635
   End
   Begin VB.Label labTo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   11
      Left            =   4020
      TabIndex        =   38
      Top             =   3360
      Width           =   1635
   End
   Begin VB.Label labTo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   10
      Left            =   4020
      TabIndex        =   37
      Top             =   3120
      Width           =   1635
   End
   Begin VB.Label labTo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   9
      Left            =   4020
      TabIndex        =   36
      Top             =   2880
      Width           =   1635
   End
   Begin VB.Label labTo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   4020
      TabIndex        =   35
      Top             =   2640
      Width           =   1635
   End
   Begin VB.Label labTo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   4020
      TabIndex        =   34
      Top             =   2400
      Width           =   1635
   End
   Begin VB.Label labTo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   4020
      TabIndex        =   33
      Top             =   2160
      Width           =   1635
   End
   Begin VB.Label labTo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   4020
      TabIndex        =   32
      Top             =   1920
      Width           =   1635
   End
   Begin VB.Label labTo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   4020
      TabIndex        =   31
      Top             =   1680
      Width           =   1635
   End
   Begin VB.Label labTo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   4020
      TabIndex        =   30
      Top             =   1440
      Width           =   1635
   End
   Begin VB.Label labTo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   4020
      TabIndex        =   29
      Top             =   1200
      Width           =   1635
   End
   Begin VB.Label labTo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   4020
      TabIndex        =   28
      Top             =   960
      Width           =   1635
   End
   Begin VB.Label labTo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   4020
      TabIndex        =   27
      Top             =   720
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label labFrom 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   19
      Left            =   2460
      TabIndex        =   26
      Top             =   5280
      Width           =   1395
   End
   Begin VB.Label labStatsHeading 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Statistical"
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
      TabIndex        =   25
      Top             =   120
      Width           =   1635
   End
   Begin VB.Label labFrom 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   18
      Left            =   2460
      TabIndex        =   24
      Top             =   5040
      Width           =   1395
   End
   Begin VB.Label labFrom 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   17
      Left            =   2460
      TabIndex        =   23
      Top             =   4800
      Width           =   1395
   End
   Begin VB.Label labFrom 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   16
      Left            =   2460
      TabIndex        =   22
      Top             =   4560
      Width           =   1395
   End
   Begin VB.Label labFrom 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   15
      Left            =   2460
      TabIndex        =   21
      Top             =   4320
      Width           =   1395
   End
   Begin VB.Label labFrom 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   14
      Left            =   2460
      TabIndex        =   20
      Top             =   4080
      Width           =   1395
   End
   Begin VB.Label labFrom 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   13
      Left            =   2460
      TabIndex        =   19
      Top             =   3840
      Width           =   1395
   End
   Begin VB.Label labFrom 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   12
      Left            =   2460
      TabIndex        =   18
      Top             =   3600
      Width           =   1395
   End
   Begin VB.Label labFrom 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   11
      Left            =   2460
      TabIndex        =   17
      Top             =   3360
      Width           =   1395
   End
   Begin VB.Label labFrom 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   10
      Left            =   2460
      TabIndex        =   16
      Top             =   3120
      Width           =   1395
   End
   Begin VB.Label labFrom 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   9
      Left            =   2460
      TabIndex        =   15
      Top             =   2880
      Width           =   1395
   End
   Begin VB.Label labFrom 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   2460
      TabIndex        =   14
      Top             =   2640
      Width           =   1395
   End
   Begin VB.Label labFrom 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   2460
      TabIndex        =   13
      Top             =   2400
      Width           =   1395
   End
   Begin VB.Label labFrom 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   2460
      TabIndex        =   12
      Top             =   2160
      Width           =   1395
   End
   Begin VB.Label labFrom 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   2460
      TabIndex        =   11
      Top             =   1920
      Width           =   1395
   End
   Begin VB.Label labFrom 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   2460
      TabIndex        =   10
      Top             =   1680
      Width           =   1395
   End
   Begin VB.Label labFrom 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   2460
      TabIndex        =   9
      Top             =   1440
      Width           =   1395
   End
   Begin VB.Label labFrom 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   2460
      TabIndex        =   8
      Top             =   1200
      Width           =   1395
   End
   Begin VB.Label labFrom 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   2460
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   1395
   End
End
Attribute VB_Name = "frmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tempwide As Single
Dim temphigh As Single
Private Sub chkVariableType_Click(Index As Integer)

Dim X As Integer
Dim Y As Integer
Dim z As Integer
Dim temptab As String

If DoNotChange = True Then Exit Sub

DoNotChange = True

For X = 0 To 3
  If X = Index Then
    chkVariableType(X).Value = 1
  Else
    chkVariableType(X).Value = 0
  End If
Next X

For Y = 0 To 2
  If optDisplayType(Y).Value = True Then z = Y
Next Y

DoNotChange = False

Select Case Index
  Case 0
    temptab = "Net Present Value"
    X = 1
  Case 1
    temptab = "Pay-Back Period"
    X = 2
  Case 2
    temptab = "Internal Rate of Return"
    X = 3
  Case 3
    temptab = "Sum of Cash Flows"
    X = 4
End Select

Select Case z
  Case 0
    labTableTitle = "Frequency Distribution Table - " & temptab
  Case 1
    labTableTitle = "Relative Frequency Histogram - " & temptab
  Case 2
    labTableTitle = "Cumulative Probability Curve - " & temptab
End Select

'optDisplayType(0).Value = True

Call getthestats(X, z)

End Sub

Private Sub Form_Activate()

ShowMenu = True

Dim X As Integer
chkVariableType(0).Value = 1
For X = 1 To 3
  chkVariableType(X).Value = 0
Next X

optDisplayType(0).Value = True

Call getthestats(1, 0)

End Sub

Private Sub Form_Deactivate()

If ShowMenu = True Then
  frmStats.Hide
  frmRisk.Show
End If

End Sub

Private Sub Form_Load()

If FullScreen = False Then
  frmStats.Top = (Screen.Height - (frmStats.Height + 350)) / 2
  frmStats.Left = (Screen.Width - frmStats.Width) / 2
Else
  frmStats.Top = 0
  frmStats.Left = 0
  frmStats.WindowState = 2
End If

If frmStats.Top < 0 Then frmStats.Top = 0
If frmStats.Left < 0 Then frmStats.Left = 0

tempwide = frmStats.ScaleWidth
temphigh = frmStats.ScaleHeight

Call screenstuff
  
End Sub


Private Sub Form_Resize()

tempwide = frmStats.ScaleWidth
temphigh = frmStats.ScaleHeight
Dim Y As Integer
Dim X As Integer

Call screenstuff

If chkVariableType(1).Value = 1 Then
  Y = 2
ElseIf chkVariableType(2).Value = 1 Then
  Y = 3
ElseIf chkVariableType(3).Value = 1 Then
  Y = 4
Else
  Y = 1
End If

If optDisplayType(1).Value = True Then
  X = 1
ElseIf optDisplayType(2).Value = True Then
  X = 2
Else
  X = 0
End If

Call getthestats(Y, X)

End Sub

Private Sub Form_Unload(Cancel As Integer)
If ShowMenu = True Then
  frmStats.Hide
  frmRisk.Show
End If
End Sub

Private Sub imgBack_Click()

frmStats.Hide
frmRisk.Show

End Sub

Private Sub labBack_Click()

frmStats.Hide
frmRisk.Show

End Sub

Private Sub labPrintScreen_Click()

If optDisplayType(0).Value = True Then
  job = 6
ElseIf optDisplayType(1).Value = True Then
  job = 7
Else
  job = 8
End If

ShowMenu = False
Call printstuffout(job)

End Sub

Private Sub optDisplayType_Click(Index As Integer)

Dim X As Integer
Dim Y As Integer

Dim tempstring As String
If chkVariableType(1).Value = 1 Then
  tempstring = " - Pay-Back Period"
  Y = 2
ElseIf chkVariableType(2).Value = 1 Then
  tempstring = " - Internal Rate of Return"
  Y = 3
ElseIf chkVariableType(3).Value = 1 Then
  tempstring = " - Sum of Cash Flows"
  Y = 4
Else
  tempstring = " - Net Present Value"
  Y = 1
End If

If Index = 0 Then
  labTableTitle.Caption = "Frequency Distribution Table" & tempstring
  For X = 0 To 3
    labFreqTitles(X).Visible = True
  Next X
  For X = 0 To 19
    labTo(X + 1).BackColor = &H0&
    labTo(X + 1).Width = 1635
    labOccur(X).Visible = True
    labPercent(X).Visible = True
    labCumPercent(X).Visible = True
  Next X
  For X = 0 To 20
    linYTics(X).Visible = False
  Next X
  For X = 0 To 9
    linXTic(X).Visible = False
  Next X
  For X = 0 To 10
    labXTics(X).Visible = False
  Next X
  labFrom(0).Visible = False
  labTo(0).Visible = False
  linGraphX.Visible = False
  linGraphY.Visible = False
  labXLabel.Visible = False
Else
  For X = 0 To 3
    labFreqTitles(X).Visible = False
  Next X
  For X = 0 To 19
    labOccur(X).Visible = False
    labPercent(X).Visible = False
    labCumPercent(X).Visible = False
  Next X
  For X = 0 To 20
    linYTics(X).Visible = True
  Next X
  labXLabel.Visible = True
  If Index = 1 Then
    labTableTitle.Caption = "Relative Frequency Histogram" & tempstring
    labXLabel.Top = temphigh * 0.9065
    labXLabel.Caption = "Relative Frequency"
    For X = 0 To 10
      labXTics(X).Visible = False
    Next X
    For X = 0 To 9
      linXTic(X).Visible = False
    Next X
  Else
    labTableTitle.Caption = "Cumulative Probability Curve" & tempstring
    labXLabel.Top = temphigh * 0.9439
    labXLabel.Caption = "Percent Cumulative Probability"
    For X = 0 To 9
      linXTic(X).Visible = True
    Next X
    For X = 0 To 10
      labXTics(X).Visible = True
    Next X
  End If
  labFrom(0).Visible = True
  labTo(0).Visible = True
  linGraphX.Visible = True
  linGraphY.Visible = True
End If

Call getthestats(Y, Index)

End Sub
Public Sub getthestats(j As Integer, k As Integer)

Dim stuff(10000) As Currency
Dim junk(40) As Integer
Dim txtfilenum As Integer
Dim proj As String
Dim cdes As String
Dim moredes As String
Dim dte As String
Dim n As Integer
Dim i As Integer
Dim count As Integer
Dim values(4) As Currency
Dim mean As Double
Dim var As Double
Dim min As Currency
Dim max As Currency
Dim inc As Currency
Dim zap As Currency
Dim intzap As Integer
Dim sumjunk As Integer
Dim tempinput As Currency

txtfilenum = FreeFile

Open MainDir & "\risk.txt" For Input As #txtfilenum

Input #txtfilenum, proj
Input #txtfilenum, cdes
Input #txtfilenum, moredes
Input #txtfilenum, dte
Input #txtfilenum, n

For i = 1 To n
  Input #txtfilenum, values(4), values(1), values(2), values(3)
  stuff(i) = values(j)
Next i

Close #txtfilenum

mean = 0
var = 0
min = stuff(1)
max = stuff(1)
i = 0
count = 0

If j = 2 Then
  Do
    i = i + 1
    min = stuff(i)
    max = stuff(i)
  Loop Until max < 50.01 Or i = CInt(riter)
End If

If j = 3 Then
  Do
    i = i + 1
    min = stuff(i)
    max = stuff(i)
  Loop Until min > -0.01 Or i = CInt(riter)
End If

For i = 1 To n
  If (j = 4) Or (j = 1) Or (j = 2 And stuff(i) < 50.01) Or (j = 3 And stuff(i) > -0.01) Then
    count = count + 1
    mean = mean + CDbl(stuff(i))
    If stuff(i) < min Then min = stuff(i)
    If stuff(i) > max Then max = stuff(i)
  End If
Next i

If count > 0 Then mean = mean / count

For i = 1 To n
  If (j = 4) Or (j = 1) Or (j = 2 And stuff(i) < 50.01) Or (j = 3 And stuff(i) > -0.01) Then
    var = var + (stuff(i) - CCur(mean)) ^ 2
  End If
Next i

If count > 1 Then var = (var / (count - 1)) ^ 0.5

For i = 1 To 40
  junk(i) = 0
Next i

inc = (max - min) / 20

For i = 1 To n
  If (j = 4) Or (j = 1) Or (j = 2 And stuff(i) < 50.01) Or (j = 3 And stuff(i) > -0.01) Then
    zap = (stuff(i) - min) / inc
    intzap = Int(zap) + 1
    junk(intzap) = junk(intzap) + 1
  End If
Next i

Select Case k
  Case 0
    sumjunk = 0
    For i = 0 To 19
      sumjunk = sumjunk + junk(i)
      tempinput = min + (i - 1) * inc
      If j = 1 Or j = 4 Then
        labFrom(i + 1).Caption = Format(LTrim(RTrim(Str(tempinput))), "$###,###,###,###")
      ElseIf j = 2 Then
        labFrom(i + 1).Caption = Format(LTrim(RTrim(Str(tempinput))), "##0.00") & " years"
      Else
        labFrom(i + 1).Caption = Format(LTrim(RTrim(Str(tempinput))), "##0.00") & "%"
      End If
      tempinput = min + i * inc
      If j = 1 Or j = 4 Then
        labTo(i + 1).Caption = Format(LTrim(RTrim(Str(tempinput))), "$###,###,###,###")
      ElseIf j = 2 Then
        labTo(i + 1).Caption = Format(LTrim(RTrim(Str(tempinput))), "##0.00") & " years"
      Else
        labTo(i + 1).Caption = Format(LTrim(RTrim(Str(tempinput))), "##0.00") & "%"
      End If
      tempinput = junk(i)
      labOccur(i).Caption = Format(LTrim(RTrim(Str(tempinput))), "#,##0")
      tempinput = junk(i) / n * 100
      labPercent(i).Caption = Format(LTrim(RTrim(Str(tempinput))), "###.00") & "%"
      tempinput = sumjunk / n * 100
      labCumPercent(i).Caption = Format(LTrim(RTrim(Str(tempinput))), "###.00") & "%"
    Next i
  Case 1
    sumjunk = 0
    For i = 0 To 20
      sumjunk = sumjunk + junk(i)
      tempinput = min + (i - 1) * inc
      If j = 1 Or j = 4 Then
        labFrom(i).Caption = Format(LTrim(RTrim(Str(tempinput))), "$###,###,###,###")
      ElseIf j = 2 Then
        labFrom(i).Caption = Format(LTrim(RTrim(Str(tempinput))), "##0.00") & " years"
      Else
        labFrom(i).Caption = Format(LTrim(RTrim(Str(tempinput))), "##0.00") & "%"
      End If
      labTo(i).Caption = ""
      If (min + (i - 1) * inc) <= 0 Then
        labTo(i).BackColor = &HC0&
      Else
        labTo(i).BackColor = &H8000&
      End If
      tempinput = junk(i) / n * 100
      labTo(i).Width = ((tempinput / 100) * ((tempwide / 12000) * 6155)) * 4
    Next i
  Case 2
    sumjunk = 0
    For i = 0 To 20
      sumjunk = sumjunk + junk(i)
      tempinput = min + (i - 1) * inc
      If j = 1 Or j = 4 Then
        labFrom(i).Caption = Format(LTrim(RTrim(Str(tempinput))), "$###,###,###,###")
      ElseIf j = 2 Then
        labFrom(i).Caption = Format(LTrim(RTrim(Str(tempinput))), "##0.00") & " years"
      Else
        labFrom(i).Caption = Format(LTrim(RTrim(Str(tempinput))), "##0.00") & "%"
      End If
      labTo(i).Caption = ""
      If (min + (i - 1) * inc) <= 0 Then
        labTo(i).BackColor = &HC0&
      Else
        labTo(i).BackColor = &H8000&
      End If
      tempinput = sumjunk / n * 100
      labTo(i).Width = ((tempwide / 12000) * 6155) - ((tempinput / 100) * ((tempwide / 12000) * 6155))
    Next i
End Select

End Sub

Public Sub screenstuff()
  
  Dim X As Integer
  Dim Y As Currency
   
  labStatsHeading.Top = temphigh * 0.0187
  labStatsHeading.Left = tempwide * 0.0131
  labStatsHeading2.Top = temphigh * 0.0818
  labStatsHeading2.Left = tempwide * 0.0721
  
  linVaryTop.X1 = tempwide * 0.0066
  linVaryTop.X2 = tempwide * 0.2557
  linVaryTop.Y1 = temphigh * 0.2243
  linVaryTop.Y2 = temphigh * 0.2243
  
  linVaryLeft.X1 = tempwide * 0.0131
  linVaryLeft.X2 = tempwide * 0.0131
  linVaryLeft.Y1 = temphigh * 0.215
  linVaryLeft.Y2 = temphigh * 0.4953

  linVaryRight.X1 = tempwide * 0.2492
  linVaryRight.X2 = tempwide * 0.2492
  linVaryRight.Y1 = temphigh * 0.215
  linVaryRight.Y2 = temphigh * 0.4953

  linVaryBottom.X1 = tempwide * 0.0066
  linVaryBottom.X2 = tempwide * 0.2557
  linVaryBottom.Y1 = temphigh * 0.486
  linVaryBottom.Y2 = temphigh * 0.486

  linAnalysisTop.X1 = tempwide * 0.0033
  linAnalysisTop.X2 = tempwide * 0.259
  linAnalysisTop.Y1 = temphigh * 0.5275
  linAnalysisTop.Y2 = temphigh * 0.5275
  
  linAnalysisLeft.X1 = tempwide * 0.0098
  linAnalysisLeft.X2 = tempwide * 0.0098
  linAnalysisLeft.Y1 = temphigh * 0.5187
  linAnalysisLeft.Y2 = temphigh * 0.7297

  linAnalysisRight.X1 = tempwide * 0.2525
  linAnalysisRight.X2 = tempwide * 0.2525
  linAnalysisRight.Y1 = temphigh * 0.5187
  linAnalysisRight.Y2 = temphigh * 0.7297

  linAnalysisBottom.X1 = tempwide * 0.0033
  linAnalysisBottom.X2 = tempwide * 0.259
  linAnalysisBottom.Y1 = temphigh * 0.7209
  linAnalysisBottom.Y2 = temphigh * 0.7209

  linMainTop.X1 = tempwide * 0.2557
  linMainTop.X2 = tempwide * 0.9902
  linMainTop.Y1 = temphigh * 0.0187
  linMainTop.Y2 = temphigh * 0.0187
  
  linMainTopMid.X1 = tempwide * 0.2689
  linMainTopMid.X2 = tempwide * 0.977
  linMainTopMid.Y1 = temphigh * 0.0748
  linMainTopMid.Y2 = temphigh * 0.0748

  linMainLeft.X1 = tempwide * 0.2623
  linMainLeft.X2 = tempwide * 0.2623
  linMainLeft.Y1 = temphigh * 0.0093
  linMainLeft.Y2 = temphigh * 0.9907

  linMainRight.X1 = tempwide * 0.9836
  linMainRight.X2 = tempwide * 0.9836
  linMainRight.Y1 = temphigh * 0.0093
  linMainRight.Y2 = temphigh * 0.9907

  linMainBottom.X1 = tempwide * 0.2557
  linMainBottom.X2 = tempwide * 0.9902
  linMainBottom.Y1 = temphigh * 0.9798
  linMainBottom.Y2 = temphigh * 0.9798
  
  For X = 0 To 3
    chkVariableType(X).Top = (X * 0.0561 * temphigh) + (temphigh * 0.2617)
    chkVariableType(X).Left = tempwide * 0.0262
  Next X
    
  For X = 0 To 2
    optDisplayType(X).Top = (X * 0.0527 * temphigh) + (temphigh * 0.5626)
    optDisplayType(X).Left = tempwide * 0.0164
  Next X
  
  For X = 0 To 20
    labFrom(X).Top = (X * 0.037383 * temphigh) + (temphigh * 0.1121)
    labFrom(X).Left = tempwide * 0.2689
    labFrom(X).Width = tempwide * 0.1525
    labTo(X).Top = (X * 0.037383 * temphigh) + (temphigh * 0.1121)
    labTo(X).Left = tempwide * 0.4393
    labTo(X).Width = tempwide * 0.1787
    linYTics(X).X1 = tempwide * 0.4262
    linYTics(X).X2 = tempwide * 0.4328
    linYTics(X).Y1 = (X * 0.037383 * temphigh) + (temphigh * 0.1308)
    linYTics(X).Y2 = (X * 0.037383 * temphigh) + (temphigh * 0.1308)
    If X < 20 Then
      labOccur(X).Top = (X * 0.037383 * temphigh) + (temphigh * 0.1495)
      labOccur(X).Left = tempwide * 0.6557
      labOccur(X).Width = tempwide * 0.0672
      labPercent(X).Top = (X * 0.037383 * temphigh) + (temphigh * 0.1495)
      labPercent(X).Left = tempwide * 0.741
      labPercent(X).Width = tempwide * 0.0803
      labCumPercent(X).Top = (X * 0.037383 * temphigh) + (temphigh * 0.1495)
      labCumPercent(X).Left = tempwide * 0.8525
      labCumPercent(X).Width = tempwide * 0.1066
    End If
  Next X
  
  For X = 0 To 10
    labXTics(X).Top = temphigh * 0.9065
    labXTics(X).Left = (X * 0.0525 * tempwide) + (tempwide * 0.4131)
    labXTics(X).Width = tempwide * 0.041
    If X < 10 Then
      linXTic(X).X1 = (X * 0.0525 * tempwide) + (tempwide * 0.4852)
      linXTic(X).X2 = (X * 0.0525 * tempwide) + (tempwide * 0.4852)
      linXTic(X).Y1 = temphigh * 0.8972
      linXTic(X).Y2 = temphigh * 0.9065
    End If
  Next X
  
  labTableTitle.Top = temphigh * 0.028
  labTableTitle.Left = tempwide * 0.2754
  labTableTitle.Width = tempwide * 0.6967
  
  labXLabel.Top = temphigh * 0.9439
  labXLabel.Left = tempwide * 0.4328
  labXLabel.Width = tempwide * 0.5262
  
  linGraphY.X1 = tempwide * 0.4328
  linGraphY.X2 = tempwide * 0.4328
  linGraphY.Y1 = temphigh * 0.1121
  linGraphY.Y2 = temphigh * 0.9065
  
  linGraphX.X1 = tempwide * 0.4328
  linGraphX.X2 = tempwide * 0.9574
  linGraphX.Y1 = temphigh * 0.8972
  linGraphX.Y2 = temphigh * 0.8972
  
  For X = 0 To 3
    labFreqTitles(X).Top = temphigh * 0.0935
    If X = 0 Then
      labFreqTitles(X).Left = tempwide * 0.2689
      labFreqTitles(X).Width = tempwide * 0.1262
    ElseIf X = 1 Then
      labFreqTitles(X).Left = tempwide * 0.4525
      labFreqTitles(X).Width = tempwide * 0.1262
    ElseIf X = 2 Then
      labFreqTitles(X).Left = tempwide * 0.7016
      labFreqTitles(X).Width = tempwide * 0.1197
    Else
      labFreqTitles(X).Left = tempwide * 0.8525
      labFreqTitles(X).Width = tempwide * 0.1066
    End If
  Next X
  
  labLastTitles(0).Top = temphigh * 0.2056
  labLastTitles(0).Left = tempwide * 0.0852
  labLastTitles(0).Width = tempwide * 0.0934
  
  labLastTitles(1).Top = temphigh * 0.5099
  labLastTitles(1).Left = tempwide * 0.0918
  labLastTitles(1).Width = tempwide * 0.0803
    
  labPrintScreen.Top = temphigh * 0.9439
  labPrintScreen.Left = tempwide * 0.1443
  
  labBack.Top = temphigh * 0.9439
  labBack.Left = tempwide * 0.0656

  imgBack.Top = temphigh * 0.9532
  imgBack.Left = tempwide * 0.0066
  imgBack.Width = tempwide * 0.0541

End Sub
