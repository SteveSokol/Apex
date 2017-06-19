VERSION 5.00
Begin VB.Form frmMining 
   BackColor       =   &H00000000&
   Caption         =   "Mining"
   ClientHeight    =   6420
   ClientLeft      =   1140
   ClientTop       =   1395
   ClientWidth     =   9150
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
   ScaleHeight     =   6420
   ScaleWidth      =   9150
   Begin VB.TextBox txtMinSetLabel 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1560
      TabIndex        =   19
      Top             =   6000
      Width           =   1095
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
      Left            =   6720
      TabIndex        =   97
      TabStop         =   0   'False
      Top             =   6060
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
      Left            =   3660
      TabIndex        =   96
      TabStop         =   0   'False
      Top             =   6060
      Width           =   195
   End
   Begin VB.HScrollBar hscSetNumbers 
      Height          =   195
      Index           =   5
      Left            =   1260
      Max             =   100
      Min             =   1
      TabIndex        =   18
      Top             =   5100
      Value           =   1
      Width           =   375
   End
   Begin VB.HScrollBar hscSetNumbers 
      Height          =   195
      Index           =   4
      Left            =   840
      Max             =   100
      TabIndex        =   17
      Top             =   4320
      Value           =   1
      Width           =   375
   End
   Begin VB.HScrollBar hscSetNumbers 
      Height          =   195
      Index           =   3
      Left            =   840
      Max             =   100
      TabIndex        =   16
      Top             =   3600
      Value           =   1
      Width           =   375
   End
   Begin VB.HScrollBar hscSetNumbers 
      Height          =   195
      Index           =   2
      Left            =   840
      Max             =   100
      TabIndex        =   15
      Top             =   2880
      Value           =   1
      Width           =   375
   End
   Begin VB.HScrollBar hscSetNumbers 
      Height          =   195
      Index           =   1
      Left            =   840
      Max             =   100
      TabIndex        =   14
      Top             =   2160
      Value           =   1
      Width           =   375
   End
   Begin VB.HScrollBar hscSetNumbers 
      Height          =   195
      Index           =   0
      Left            =   840
      Max             =   100
      TabIndex        =   13
      Top             =   1440
      Value           =   1
      Width           =   375
   End
   Begin VB.TextBox txtMiningValues 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   12
      Left            =   6060
      TabIndex        =   12
      Top             =   5340
      Width           =   1215
   End
   Begin VB.TextBox txtMiningValues 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   11
      Left            =   6060
      TabIndex        =   11
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox txtMiningValues 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   10
      Left            =   6060
      TabIndex        =   10
      Top             =   4740
      Width           =   1215
   End
   Begin VB.TextBox txtMiningValues 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   9
      Left            =   6060
      TabIndex        =   9
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox txtMiningValues 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   8
      Left            =   6060
      TabIndex        =   8
      Top             =   4140
      Width           =   1215
   End
   Begin VB.TextBox txtMiningValues 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   7
      Left            =   6060
      TabIndex        =   7
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txtMiningValues 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   6
      Left            =   6060
      TabIndex        =   6
      Top             =   3540
      Width           =   1215
   End
   Begin VB.TextBox txtMiningValues 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   5
      Left            =   6060
      TabIndex        =   5
      Text            =   "320"
      Top             =   2700
      Width           =   1215
   End
   Begin VB.TextBox txtMiningValues 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   4
      Left            =   6060
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtMiningValues 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   6060
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtMiningValues 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   6060
      TabIndex        =   2
      Text            =   "100"
      Top             =   1260
      Width           =   1215
   End
   Begin VB.TextBox txtMiningValues 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   6060
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtMiningValues 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   6060
      TabIndex        =   0
      Text            =   "2"
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label labMiningHelp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Height          =   240
      Left            =   8640
      TabIndex        =   106
      Top             =   6120
      Width           =   435
   End
   Begin VB.Label YikesLabel 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Remember to Activate the Set"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   105
      Top             =   5400
      Width           =   2775
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
      Left            =   5520
      TabIndex        =   104
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label labMiningLabels 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Set Label"
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
      Index           =   18
      Left            =   1560
      TabIndex        =   103
      Top             =   5700
      Width           =   1095
   End
   Begin VB.Label labSetLabels 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   1680
      TabIndex        =   102
      Top             =   4320
      Width           =   915
   End
   Begin VB.Label labSetLabels 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1680
      TabIndex        =   101
      Top             =   3600
      Width           =   915
   End
   Begin VB.Label labSetLabels 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   1680
      TabIndex        =   100
      Top             =   2880
      Width           =   915
   End
   Begin VB.Label labSetLabels 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   99
      Top             =   2160
      Width           =   915
   End
   Begin VB.Label labSetLabels 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   98
      Top             =   1440
      Width           =   915
   End
   Begin VB.Label labMiningUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Index           =   13
      Left            =   7320
      TabIndex        =   95
      Top             =   5760
      Width           =   75
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
      Index           =   12
      Left            =   8340
      TabIndex        =   94
      Top             =   5400
      Visible         =   0   'False
      Width           =   435
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
      Index           =   11
      Left            =   8340
      TabIndex        =   93
      Top             =   5100
      Visible         =   0   'False
      Width           =   435
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
      Index           =   10
      Left            =   8340
      TabIndex        =   92
      Top             =   4800
      Visible         =   0   'False
      Width           =   435
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
      Index           =   9
      Left            =   8340
      TabIndex        =   91
      Top             =   4500
      Visible         =   0   'False
      Width           =   435
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
      Index           =   8
      Left            =   8340
      TabIndex        =   90
      Top             =   4200
      Visible         =   0   'False
      Width           =   435
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
      Index           =   7
      Left            =   8340
      TabIndex        =   89
      Top             =   3900
      Visible         =   0   'False
      Width           =   435
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
      Index           =   6
      Left            =   8340
      TabIndex        =   88
      Top             =   3600
      Visible         =   0   'False
      Width           =   435
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
      Index           =   5
      Left            =   8340
      TabIndex        =   87
      Top             =   2760
      Visible         =   0   'False
      Width           =   435
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
      Index           =   4
      Left            =   8340
      TabIndex        =   86
      Top             =   2460
      Visible         =   0   'False
      Width           =   435
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
      Index           =   3
      Left            =   8340
      TabIndex        =   85
      Top             =   1620
      Visible         =   0   'False
      Width           =   435
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
      Index           =   2
      Left            =   8340
      TabIndex        =   84
      Top             =   1320
      Visible         =   0   'False
      Width           =   435
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
      Index           =   1
      Left            =   8340
      TabIndex        =   83
      Top             =   1020
      Visible         =   0   'False
      Width           =   435
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
      Index           =   0
      Left            =   8340
      TabIndex        =   82
      Top             =   420
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label labDepTag 
      BackColor       =   &H00000000&
      Caption         =   "Dependent Tag"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   6960
      TabIndex        =   81
      Top             =   6060
      Width           =   1335
   End
   Begin VB.Label labIndTag 
      BackColor       =   &H00000000&
      Caption         =   "Independent Tag"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   3900
      TabIndex        =   80
      Top             =   6060
      Width           =   1515
   End
   Begin VB.Label labMiningLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Index           =   17
      Left            =   5880
      TabIndex        =   79
      Top             =   5700
      Width           =   135
   End
   Begin VB.Label labMiningLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   16
      Left            =   5880
      TabIndex        =   78
      Top             =   5400
      Width           =   135
   End
   Begin VB.Label labMiningLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   15
      Left            =   5880
      TabIndex        =   77
      Top             =   5100
      Width           =   135
   End
   Begin VB.Label labMiningLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   14
      Left            =   5880
      TabIndex        =   76
      Top             =   4800
      Width           =   135
   End
   Begin VB.Label labMiningLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   13
      Left            =   5880
      TabIndex        =   75
      Top             =   4500
      Width           =   135
   End
   Begin VB.Label labMiningLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   12
      Left            =   5880
      TabIndex        =   74
      Top             =   3900
      Width           =   135
   End
   Begin VB.Label labMiningLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   11
      Left            =   5880
      TabIndex        =   73
      Top             =   3600
      Width           =   135
   End
   Begin VB.Label labScreenTotals 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   6120
      TabIndex        =   72
      Top             =   5700
      Width           =   1095
   End
   Begin VB.Label labScreenTotals 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   6120
      TabIndex        =   71
      Top             =   3060
      Width           =   1095
   End
   Begin VB.Label labScreenTotals 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   5880
      TabIndex        =   70
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label labScreenTotals 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   6120
      TabIndex        =   69
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label labMiningLabels 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Tags"
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
      Height          =   240
      Index           =   10
      Left            =   8340
      TabIndex        =   68
      Top             =   180
      Width           =   435
   End
   Begin VB.Label labMiningLabels 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Set Designations"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   240
      TabIndex        =   67
      Top             =   780
      Width           =   1575
   End
   Begin VB.Line LineSetRight 
      BorderColor     =   &H00FFFF00&
      X1              =   2700
      X2              =   2700
      Y1              =   840
      Y2              =   4740
   End
   Begin VB.Line LineSetTop 
      BorderColor     =   &H00FFFF00&
      X1              =   300
      X2              =   2760
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Line LineSetBottom 
      BorderColor     =   &H00FFFF00&
      X1              =   300
      X2              =   2760
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line LineSetLeft 
      BorderColor     =   &H00FFFF00&
      X1              =   360
      X2              =   360
      Y1              =   840
      Y2              =   4740
   End
   Begin VB.Label labSetNumbers 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   1680
      TabIndex        =   66
      Top             =   5100
      Width           =   315
   End
   Begin VB.Label labSetNumbers 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   1260
      TabIndex        =   65
      Top             =   4320
      Width           =   315
   End
   Begin VB.Label labSetNumbers 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   1260
      TabIndex        =   64
      Top             =   3600
      Width           =   315
   End
   Begin VB.Label labSetNumbers 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   1260
      TabIndex        =   63
      Top             =   2880
      Width           =   315
   End
   Begin VB.Label labSetNumbers 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   1260
      TabIndex        =   62
      Top             =   2160
      Width           =   315
   End
   Begin VB.Label labSetNumbers 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   1260
      TabIndex        =   61
      Top             =   1440
      Width           =   315
   End
   Begin VB.Label labMiningLabels 
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
      Height          =   225
      Index           =   9
      Left            =   1020
      TabIndex        =   60
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label labMiningLabels 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Processing Set"
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
      Height          =   240
      Index           =   8
      Left            =   480
      TabIndex        =   59
      Top             =   4020
      Width           =   1350
   End
   Begin VB.Label labMiningLabels 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Processing Set"
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
      Index           =   7
      Left            =   480
      TabIndex        =   58
      Top             =   3300
      Width           =   1350
   End
   Begin VB.Label labMiningLabels 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Processing Set"
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
      Left            =   480
      TabIndex        =   57
      Top             =   2580
      Width           =   1335
   End
   Begin VB.Label labMiningLabels 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Royalty Set"
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
      Index           =   5
      Left            =   480
      TabIndex        =   56
      Top             =   1860
      Width           =   1335
   End
   Begin VB.Label labMiningLabels 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Grade Set"
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
      Index           =   4
      Left            =   480
      TabIndex        =   55
      Top             =   1140
      Width           =   1350
   End
   Begin VB.Label labMiningLabels 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   " Mine Operating Costs "
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   2940
      TabIndex        =   54
      Top             =   3300
      Width           =   2010
   End
   Begin VB.Label labMiningLabels 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   " Production "
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   2940
      TabIndex        =   53
      Top             =   2160
      Width           =   1065
   End
   Begin VB.Label labMiningLabels 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   " Reserves "
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   2940
      TabIndex        =   52
      Top             =   60
      Width           =   945
   End
   Begin VB.Line LineRight 
      BorderColor     =   &H00FFFF00&
      X1              =   8820
      X2              =   8820
      Y1              =   60
      Y2              =   6060
   End
   Begin VB.Line LineBottom 
      BorderColor     =   &H00FFFF00&
      X1              =   3000
      X2              =   8880
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line LineBottomMiddle 
      BorderColor     =   &H00FFFF00&
      X1              =   3000
      X2              =   8760
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line LineTopMiddle 
      BorderColor     =   &H00FFFF00&
      X1              =   3000
      X2              =   8760
      Y1              =   2220
      Y2              =   2220
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFF00&
      X1              =   3000
      X2              =   8880
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line LineLeft 
      BorderColor     =   &H00FFFF00&
      X1              =   3060
      X2              =   3060
      Y1              =   60
      Y2              =   6060
   End
   Begin VB.Label labMiningHeading 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mining"
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
      Left            =   180
      TabIndex        =   51
      Top             =   180
      Width           =   1245
   End
   Begin VB.Label labBackToMenu 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Height          =   240
      Left            =   660
      TabIndex        =   50
      Top             =   6120
      Width           =   555
   End
   Begin VB.Image imgBackToMenu 
      Height          =   195
      Left            =   60
      Picture         =   "frmMining.frx":0000
      Stretch         =   -1  'True
      Top             =   6180
      Width           =   495
   End
   Begin VB.Label labMiningUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   12
      Left            =   7320
      TabIndex        =   49
      Top             =   5400
      Width           =   45
   End
   Begin VB.Label labMiningUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   11
      Left            =   7320
      TabIndex        =   48
      Top             =   5100
      Width           =   45
   End
   Begin VB.Label labMiningUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   10
      Left            =   7320
      TabIndex        =   47
      Top             =   4800
      Width           =   45
   End
   Begin VB.Label labMiningUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   9
      Left            =   7320
      TabIndex        =   46
      Top             =   4500
      Width           =   45
   End
   Begin VB.Label labMiningUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "/cubic yard"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   8
      Left            =   7320
      TabIndex        =   45
      Top             =   3900
      Width           =   870
   End
   Begin VB.Label labMiningUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "/day"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   7
      Left            =   7320
      TabIndex        =   44
      Top             =   3600
      Width           =   330
   End
   Begin VB.Label labMiningUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "years"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Index           =   6
      Left            =   7320
      TabIndex        =   43
      Top             =   3060
      Width           =   450
   End
   Begin VB.Label labMiningUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "days/year"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   5
      Left            =   7320
      TabIndex        =   42
      Top             =   2760
      Width           =   780
   End
   Begin VB.Label labMiningUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   4
      Left            =   7320
      TabIndex        =   41
      Top             =   2460
      Width           =   45
   End
   Begin VB.Label labMiningUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Index           =   3
      Left            =   7320
      TabIndex        =   40
      Top             =   1980
      Width           =   45
   End
   Begin VB.Label labMiningUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "percent"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   2
      Left            =   7320
      TabIndex        =   39
      Top             =   1620
      Width           =   615
   End
   Begin VB.Label labMiningUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "percent"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   7320
      TabIndex        =   38
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label labMiningUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   7320
      TabIndex        =   37
      Top             =   1020
      Width           =   45
   End
   Begin VB.Label labMiningTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Screen Total"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Index           =   16
      Left            =   3360
      TabIndex        =   36
      Top             =   5700
      Width           =   2295
   End
   Begin VB.Label labMiningTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "(Equipment Operating Cost)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   15
      Left            =   3360
      TabIndex        =   35
      Top             =   5400
      Width           =   2310
   End
   Begin VB.Label labMiningTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "(Supply Cost)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   14
      Left            =   3360
      TabIndex        =   34
      Top             =   5100
      Width           =   2310
   End
   Begin VB.Label labMiningTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "(Labor Cost)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   13
      Left            =   3360
      TabIndex        =   33
      Top             =   4800
      Width           =   2340
   End
   Begin VB.Label labMiningTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Variable Cost - Waste"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   12
      Left            =   3360
      TabIndex        =   32
      Top             =   4500
      Width           =   2310
   End
   Begin VB.Label labMiningTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Stripping Ratio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   11
      Left            =   3360
      TabIndex        =   31
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label labMiningTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Variable Cost - Ore"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   10
      Left            =   3360
      TabIndex        =   30
      Top             =   3900
      Width           =   2325
   End
   Begin VB.Label labMiningTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Fixed Cost"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   9
      Left            =   3360
      TabIndex        =   29
      Top             =   3600
      Width           =   2310
   End
   Begin VB.Label labMiningTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Production Life"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Index           =   8
      Left            =   3360
      TabIndex        =   28
      Top             =   3060
      Width           =   2310
   End
   Begin VB.Label labMiningTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Operating Schedule"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   7
      Left            =   3360
      TabIndex        =   27
      Top             =   2760
      Width           =   2310
   End
   Begin VB.Label labMiningTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Rate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   6
      Left            =   3360
      TabIndex        =   26
      Top             =   2460
      Width           =   2325
   End
   Begin VB.Label labMiningTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Total Production"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Index           =   5
      Left            =   3360
      TabIndex        =   25
      Top             =   1920
      Width           =   2310
   End
   Begin VB.Label labMiningTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Wallrock Dilution"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   4
      Left            =   3360
      TabIndex        =   24
      Top             =   1620
      Width           =   2340
   End
   Begin VB.Label labMiningTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Mine Recovery"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   3
      Left            =   3360
      TabIndex        =   23
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label labMiningTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Ore Reserves"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   2
      Left            =   3360
      TabIndex        =   22
      Top             =   1020
      Width           =   2280
   End
   Begin VB.Label labMiningTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Ending in Year"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Index           =   1
      Left            =   3360
      TabIndex        =   21
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label labMiningTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Initial Production in Year"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   3360
      TabIndex        =   20
      Top             =   420
      Width           =   2295
   End
End
Attribute VB_Name = "frmMining"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim temphigh As Single
Dim tempwide As Single

Private Sub comDepTag_Click()

If nTag = 0 Then
  WarnNumber = 4
  DoNotChange = True
  ShowMenu = False
  frmWarnTheUser.Show
Else
  DoNotChange = False
  ShowMenu = True
  If labCheckTag(LastCell).Visible = False Then
    ParamSet = False
    dTag = dTag + 1
    labCheckTag(LastCell).Visible = True
    labCheckTag(LastCell).ForeColor = &HFFFF&
    labCheckTag(LastCell).Caption = LTrim(Str(nTag))
    If LastCell = 0 Then
      Tagged(hscSetNumbers(5).Value, LastCell + 34).Dependent = nTag
      DepTagData(nTag, dTag).TheCell = LastCell + 34
      DepTagData(nTag, dTag).Title = labMiningTitles(LastCell).Caption
    ElseIf LastCell < 4 Then
      Tagged(hscSetNumbers(5).Value, LastCell + 35).Dependent = nTag
      DepTagData(nTag, dTag).TheCell = LastCell + 35
      DepTagData(nTag, dTag).Title = labMiningTitles(LastCell + 1).Caption
      DepTagData(nTag, dTag).Units = labMiningUnits(LastCell - 1).Caption
   ElseIf LastCell < 6 Then
      Tagged(hscSetNumbers(5).Value, LastCell + 36).Dependent = nTag
      DepTagData(nTag, dTag).TheCell = LastCell + 36
      DepTagData(nTag, dTag).Title = labMiningTitles(LastCell + 2).Caption
      If LastCell = 4 Then
        DepTagData(nTag, dTag).Title = "Production " & labMiningTitles(LastCell + 2).Caption
      End If
      DepTagData(nTag, dTag).Units = labMiningUnits(LastCell).Caption
    Else
      Tagged(hscSetNumbers(5).Value, LastCell + 39).Dependent = nTag
      DepTagData(nTag, dTag).TheCell = LastCell + 39
      Select Case LastCell
        Case 6
          DepTagData(nTag, dTag).Title = "Mining - " & labMiningTitles(LastCell + 3).Caption
          DepTagData(nTag, dTag).Units = labMiningUnits(LastCell + 1).Caption
        Case 7
          DepTagData(nTag, dTag).Title = "Variable Mining Cost - ore"
          DepTagData(nTag, dTag).Units = labMiningUnits(LastCell + 1).Caption
        Case 8
          DepTagData(nTag, dTag).Title = labMiningTitles(LastCell + 3).Caption
        Case 9
          DepTagData(nTag, dTag).Title = "Variable Mining Cost - waste"
          DepTagData(nTag, dTag).Units = labMiningUnits(LastCell).Caption
        Case Else
          DepTagData(nTag, dTag).Title = "Mining - " & labMiningTitles(LastCell + 3).Caption
          DepTagData(nTag, dTag).Units = labMiningUnits(LastCell).Caption
      End Select
    End If
    DepTagData(nTag, dTag).SetNumber = hscSetNumbers(5).Value
  End If
  txtMiningValues(LastCell).SetFocus
End If

End Sub

Private Sub comIndTag_Click()

If labCheckTag(LastCell).Visible = False Then
  ParamSet = False
  nTag = nTag + 1
  dTag = 0
  labCheckTag(LastCell).Visible = True
  labCheckTag(LastCell).ForeColor = &HFF&
  labCheckTag(LastCell).Caption = LTrim(Str(nTag))
  If LastCell = 0 Then
    Tagged(hscSetNumbers(5).Value, LastCell + 34).Independent = nTag
    IndTagData(nTag).TheCell = LastCell + 34
    IndTagData(nTag).Title = labMiningTitles(LastCell).Caption
  ElseIf LastCell < 4 Then
    Tagged(hscSetNumbers(5).Value, LastCell + 35).Independent = nTag
    IndTagData(nTag).TheCell = LastCell + 35
    IndTagData(nTag).Title = labMiningTitles(LastCell + 1).Caption
    IndTagData(nTag).Units = labMiningUnits(LastCell - 1).Caption
  ElseIf LastCell < 6 Then
    Tagged(hscSetNumbers(5).Value, LastCell + 36).Independent = nTag
    IndTagData(nTag).TheCell = LastCell + 36
    IndTagData(nTag).Title = labMiningTitles(LastCell + 2).Caption
    If LastCell = 4 Then
      IndTagData(nTag).Title = "Production " & labMiningTitles(LastCell + 2).Caption
    End If
    IndTagData(nTag).Units = labMiningUnits(LastCell).Caption
  Else
    Tagged(hscSetNumbers(5).Value, LastCell + 39).Independent = nTag
    IndTagData(nTag).TheCell = LastCell + 39
    Select Case LastCell
      Case 6
        IndTagData(nTag).Title = "Mining - " & labMiningTitles(LastCell + 3).Caption
        IndTagData(nTag).Units = labMiningUnits(LastCell + 1).Caption
      Case 7
        IndTagData(nTag).Title = "Variable Mining Cost - Ore"
        IndTagData(nTag).Units = labMiningUnits(LastCell + 1).Caption
      Case 8
        IndTagData(nTag).Title = labMiningTitles(LastCell + 3).Caption
      Case 9
        IndTagData(nTag).Title = "Variable Mining Cost - Waste"
        IndTagData(nTag).Units = labMiningUnits(LastCell).Caption
      Case Else
        IndTagData(nTag).Title = "Mining - " & labMiningTitles(LastCell + 3).Caption
        IndTagData(nTag).Units = labMiningUnits(LastCell).Caption
    End Select
  End If
  IndTagData(nTag).SetNumber = hscSetNumbers(5).Value
End If

txtMiningValues(LastCell).SetFocus

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()

Dim baseunit As String
Dim baselength As Integer
Dim i As Integer

If IsHelpOn = True Then
  If LastCell = 100 Then
    txtMinSetLabel.SetFocus
  ElseIf LastCell < 13 Then
    txtMiningValues(LastCell).SetFocus
  ElseIf LastCell < 18 Then
    hscSetNumbers(LastCell - 13).SetFocus
  End If
  IsHelpOn = False
Else
  baseunit = LTrim(RTrim(CommodityData(1, 0).reserves))
  If baseunit = "" Then baseunit = "tons"
  baselength = Len(baseunit) - 1
  baseunit = Left(baseunit, baselength)
  labMiningUnits(0).Caption = baseunit & "s"
  labMiningUnits(3).Caption = baseunit & "s"
  labMiningUnits(4).Caption = baseunit & "s/day"
  labMiningUnits(8).Caption = "/" & baseunit & " ore"
  labMiningUnits(9).Caption = "/" & baseunit & " waste"
  For i = 10 To 13
    labMiningUnits(i).Caption = "/" & baseunit & " ore"
  Next i

  hscSetNumbers(5).Value = 1

  For i = 0 To 4
    hscSetNumbers(i).Value = Primary(hscSetNumbers(5).Value, (27 + i))
    If Primary(hscSetNumbers(5).Value, 27 + i) <> 0 Then
      Select Case i
        Case 0
          labSetLabels(i).Caption = Pn1(3, Int(Primary(hscSetNumbers(i).Value, 27 + 1)))
        Case 1
          labSetLabels(i).Caption = Pn1(8, Int(Primary(hscSetNumbers(i).Value, 27 + 1)))
        Case Else
          labSetLabels(i).Caption = Pn1(5, Int(Primary(hscSetNumbers(i).Value, 27 + 1)))
      End Select
    End If
  Next i

  txtMinSetLabel.Text = Pn1(4, hscSetNumbers(5).Value)
  Call drawthevalues
  ShowMenu = True
  If InsertFlag = True Then
    labInsert.Caption = "Insert"
  Else
    labInsert.Caption = "Typeover"
  End If
    
  LastCell = 0
  txtMiningValues(0).SetFocus
End If

End Sub

Private Sub Form_Deactivate()
  
  If ShowMenu = True Then
    frmMining.Hide
    Call InputMenuAccess(1)
  End If
  
End Sub

Private Sub Form_Load()

Dim i As Integer
Dim X As Integer

If FullScreen = False Then
  frmMining.Top = (Screen.Height - (frmMining.Height + 350)) / 2
  frmMining.Left = (Screen.Width - frmMining.Width) / 2
Else
  frmMining.Top = 0
  frmMining.Left = 0
  frmMining.WindowState = 2
End If

If frmMining.Top < 0 Then frmMining.Top = 0
If frmMining.Left < 0 Then frmMining.Left = 0

tempwide = frmMining.ScaleWidth
temphigh = frmMining.ScaleHeight

DoNotChange = True

For i = 0 To 5
  If i < 3 Then
    hscSetNumbers(i).Value = 1
  Else
    hscSetNumbers(i).Value = 1
  End If
Next i

If PageChange(2) = True Then
  Call drawthevalues
End If

DoNotChange = False

Call screenstuff

End Sub

Private Sub Form_Resize()

tempwide = frmMining.ScaleWidth
temphigh = frmMining.ScaleHeight

Call screenstuff
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  frmMining.Hide
  If ShowMenu = True Then Call InputMenuAccess(1)

End Sub

Private Sub hscSetNumbers_Change(Index As Integer)

Dim i As Integer
Dim rez As Integer
Dim prebit As Integer
Dim bit As Integer
Dim tempind As Integer

If DoNotChange = True Then Exit Sub

If Index = 5 Then
  For i = 0 To 4
    hscSetNumbers(i).Value = Primary(hscSetNumbers(Index).Value, i + 27)
  Next i
End If

labSetNumbers(Index).Caption = LTrim(RTrim(Str(hscSetNumbers(Index).Value)))
txtMinSetLabel.Text = Pn1(4, hscSetNumbers(5).Value)

If hscSetNumbers(5).Value > Np(4) Then
  Np(4) = hscSetNumbers(5).Value
End If

If Np(4) > Npna Then Npna = Np(4)

If Index < 5 Then
  Primary(hscSetNumbers(5).Value, Index + 27) = CCur(hscSetNumbers(Index).Value)
  Call recalc(hscSetNumbers(5).Value, Index + 27)
End If

Call drawthevalues

For rez = 7 To 11
  Select Case rez
    Case 7
      prebit = 3
    Case 8
      prebit = 8
    Case 9
      prebit = 5
  End Select
  If Primary(hscSetNumbers(5).Value, 20 + rez) <> 0 Then
    bit = Int(Primary(hscSetNumbers(5).Value, 20 + rez))
    labSetLabels(rez - 7).Caption = Pn1(prebit, bit)
  ElseIf rez = 8 And Primary(hscSetNumbers(5).Value, 20 + rez) = 0 Then
    labSetLabels(rez - 7).Caption = "Inactive"
  Else
    labSetLabels(rez - 7).Caption = ""
  End If
Next rez

If labSetLabels(1).Caption = "Inactive" Then
    YikesLabel.Caption = "Remember to Activate this Set"
Else
    YikesLabel.Caption = ""
End If

txtMiningValues(0).SetFocus

End Sub

Private Sub hscSetNumbers_GotFocus(Index As Integer)
LastCell = Index + 13
End Sub

Private Sub imgBackToMenu_Click()

  frmMining.Hide
  If ShowMenu = True Then Call InputMenuAccess(1)

End Sub

Private Sub labBackToMenu_Click()

  frmMining.Hide
  If ShowMenu = True Then Call InputMenuAccess(1)

End Sub

Public Sub CalcScreenTotals(Index As Integer)

  Dim TempValue As Double
  Dim X As Integer
  
  Select Case Index
    Case 1, 2, 3
      TempValue = (CDbl(Val(txtMiningValues(3).Text)) / 100)
      If TempValue = 1 Then TempValue = 0.99
      TempValue = TempValue / (1 - TempValue) * (CDbl(Val(txtMiningValues(1).Text)) * CDbl(Val(txtMiningValues(2).Text)) / 100)
      TempValue = TempValue + (CDbl(Val(txtMiningValues(1).Text)) * CDbl(Val(txtMiningValues(2).Text)) / 100)
      If TempValue > 0 Then
        labScreenTotals(1).Caption = Str(TempValue)
      Else
        labScreenTotals(1).Caption = "0.00"
      End If
      FakeReserves = TempValue
  End Select

  Select Case Index
    Case 0, 1, 2, 3, 4, 5
      TempValue = CDbl(Val(txtMiningValues(4).Text)) * CDbl(Val(txtMiningValues(5).Text))
      labScreenTotals(1).Caption = Format(labScreenTotals(1).Caption, "############")
      If TempValue > 0 Then TempValue = CDbl(Val(labScreenTotals(1).Caption)) / TempValue
      If TempValue > 0 Then
        labScreenTotals(2).Caption = Str(TempValue)
        FakeLife = TempValue
        TempValue = TempValue + CDbl(Val(txtMiningValues(0).Text))
        labScreenTotals(0).Caption = Str(TempValue)
      Else
        labScreenTotals(2).Caption = "0.00"
        FakeLife = 0
        labScreenTotals(0).Caption = "0.00"
      End If
  End Select
  
  TempValue = 0
    
  Select Case Index
    Case 4, 6 To 12
      For X = 0 To 2
        labMiningTitles(X + 10).Enabled = True
        txtMiningValues(X + 7).Enabled = True
        If X < 2 Then labMiningUnits(X + 8).Enabled = True
        labMiningTitles(X + 13).Enabled = True
        txtMiningValues(X + 10).Enabled = True
        labMiningUnits(X + 10).Enabled = True
      Next X
      If Val(txtMiningValues(4).Text) > 0 Then TempValue = CDbl(Val(txtMiningValues(6).Text)) / CDbl(Val(txtMiningValues(4).Text))
      If (Val(txtMiningValues(7).Text) + Val(txtMiningValues(8).Text) + Val(txtMiningValues(9).Text)) > 0 Then
        For X = 0 To 2
          labMiningTitles(X + 13).Enabled = False
          txtMiningValues(X + 10).Enabled = False
          labMiningUnits(X + 10).Enabled = False
        Next X
        TempValue = TempValue + CDbl(Val(txtMiningValues(7).Text))
        TempValue = TempValue + (CDbl(Val(txtMiningValues(8).Text)) * CDbl(Val(txtMiningValues(9).Text)))
      ElseIf (Val(txtMiningValues(10).Text) + Val(txtMiningValues(11).Text) + Val(txtMiningValues(12).Text)) > 0 Then
        For X = 0 To 2
          labMiningTitles(X + 10).Enabled = False
          txtMiningValues(X + 7).Enabled = False
          If X < 2 Then labMiningUnits(X + 8).Enabled = False
        Next X
        TempValue = TempValue + CDbl(Val(txtMiningValues(10).Text)) + CCur(Val(txtMiningValues(11).Text)) + CDbl(Val(txtMiningValues(12).Text))
      End If
      If TempValue > 0 Then
        labScreenTotals(3).Caption = Str(TempValue)
      Else
        labScreenTotals(3).Caption = ""
      End If
  End Select

  labScreenTotals(0).Caption = Format(labScreenTotals(0).Caption, "######.00")
  Primary(hscSetNumbers(5).Value, 35) = CCur(Val(labScreenTotals(0).Caption))
  labScreenTotals(0).Caption = Format(labScreenTotals(0).Caption, "###,###.00")
  labScreenTotals(1).Caption = Format(labScreenTotals(1).Caption, "############")
  Primary(hscSetNumbers(5).Value, 39) = CCur(Val(labScreenTotals(1).Caption))
  labScreenTotals(1).Caption = Format(labScreenTotals(1).Caption, "###,###,###,###")
  labScreenTotals(2).Caption = Format(labScreenTotals(2).Caption, "######.00")
  Primary(hscSetNumbers(5).Value, 42) = CCur(Val(labScreenTotals(2).Caption))
  labScreenTotals(2).Caption = Format(labScreenTotals(2).Caption, "###,###.00")
  labScreenTotals(3).Caption = Format(labScreenTotals(3).Caption, "###,###.00")
  
End Sub

Private Sub Label1_Click()

PrintForm

End Sub

Private Sub labMiningHelp_Click()

Dim begin As Integer
Dim sendindex As Integer

begin = 0
ShowMenu = False
WhichScreen = 2
Select Case LastCell
  Case 0
    sendindex = LastCell + 34
  Case 1 To 3
    sendindex = LastCell + 35
  Case 4, 5
    sendindex = LastCell + 36
  Case Is < 13
    sendindex = LastCell + 39
  Case 100
    sendindex = LastCell - 67
  Case Else
    sendindex = LastCell + 14
End Select

Call frmApexHelp.gethelptext(sendindex, begin)
frmApexHelp.Show

End Sub

Private Sub txtMiningValues_Change(Index As Integer)
  
Dim X As Integer

If DoNotChange = True Then Exit Sub

PageChange(2) = True

If labCheckTag(Index).Visible = True Then ParamSet = False

Select Case Index
  Case 0
    X = 34
  Case 1 To 3
    X = 35
  Case 4, 5
    X = 36
  Case Else
    X = 39
End Select

'if Index = 3 Then
  'If Val(txtMiningValues(Index).Text) >= 100 Then txtMiningValues(Index).Text = 99
'End If

'Select Case X
'  Case 34, 39
    Primary(hscSetNumbers(5).Value, Index + X) = CCur(Val(txtMiningValues(Index).Text))
'End Select

If Index = 1 Then DidWeChange(1) = False

If Index = 4 And DidWeChange(1) = False Then
  Primary(hscSetNumbers(5).Value, 92) = CCur(Val(txtMiningValues(Index).Text))
End If

Call CalcScreenTotals(Index)
Call recalc(hscSetNumbers(5).Value, Index + X)

End Sub


Private Sub txtMiningValues_GotFocus(Index As Integer)

LastCell = Index

End Sub



Public Sub screenstuff()
 
  Dim X As Integer
  Dim Y As Currency
  
  labMiningHeading.Top = temphigh * 0.0334
  labMiningHeading.Left = tempwide * 0.0194
  
  LineSetTop.X1 = tempwide * 0.0328
  LineSetTop.X2 = tempwide * 0.301
  LineSetTop.Y1 = temphigh * 0.1308
  LineSetTop.Y2 = temphigh * 0.1308
  
  LineSetLeft.X1 = tempwide * 0.0393
  LineSetLeft.X2 = tempwide * 0.0393
  LineSetLeft.Y1 = temphigh * 0.1215
  LineSetLeft.Y2 = temphigh * 0.7383

  LineSetRight.X1 = tempwide * 0.2944
  LineSetRight.X2 = tempwide * 0.2944
  LineSetRight.Y1 = temphigh * 0.1215
  LineSetRight.Y2 = temphigh * 0.7383

  LineSetBottom.X1 = tempwide * 0.0328
  LineSetBottom.X2 = tempwide * 0.301
  LineSetBottom.Y1 = temphigh * 0.729
  LineSetBottom.Y2 = temphigh * 0.729

  labMiningLabels(3).Top = temphigh * 0.1215
  labMiningLabels(3).Left = tempwide * 0.0262
 
  For X = 0 To 4
    labSetNumbers(X).Top = (X * 0.1121 * temphigh) + (temphigh * 0.2103)
    labSetNumbers(X).Left = tempwide * 0.1377
    labSetNumbers(X).Width = tempwide * 0.0344
    hscSetNumbers(X).Top = (X * 0.1121 * temphigh) + (temphigh * 0.2149)
    hscSetNumbers(X).Left = (tempwide * 0.1123) - 188
    labMiningLabels(X + 4).Top = (X * 0.1121 * temphigh) + (temphigh * 0.1682)
    labMiningLabels(X + 4).Left = tempwide * 0.0525
    labMiningLabels(X + 4).Width = tempwide * 0.1475
    labSetLabels(X).Top = (X * 0.1121 * temphigh) + (temphigh * 0.2103)
    labSetLabels(X).Left = tempwide * 0.1832
    labSetLabels(X).Width = tempwide * 0.0998
  Next X
    
  YikesLabel.Top = temphigh * 0.835
  YikesLabel.Left = tempwide * 0.0214
  YikesLabel.Width = tempwide * 0.2965
    
  labMiningLabels(9).Top = temphigh * 0.7477
  labMiningLabels(9).Left = tempwide * 0.1115
  labMiningLabels(9).Width = tempwide * 0.1197
  
  labMiningLabels(18).Top = temphigh * 0.8866
  labMiningLabels(18).Left = tempwide * 0.1815
  labMiningLabels(18).Width = tempwide * 0.1197
  
  hscSetNumbers(5).Top = temphigh * 0.7945
  hscSetNumbers(5).Left = tempwide * 0.1377
  
  labSetNumbers(5).Top = temphigh * 0.7899
  labSetNumbers(5).Left = tempwide * 0.1836
  labSetNumbers(5).Width = tempwide * 0.0344
  
  txtMinSetLabel.Top = temphigh * 0.9301
  txtMinSetLabel.Left = tempwide * 0.1815
  txtMinSetLabel.Width = tempwide * 0.1197
  
  For X = 0 To 16
    If X < 6 Then
      Y = 0
    ElseIf X < 9 Then
      Y = 0.0374
    Else
      Y = 0.0748
    End If
    labMiningTitles(X).Top = (Y * temphigh) + (X * 0.0467 * temphigh) + (temphigh * 0.0654)
    labMiningTitles(X).Left = tempwide * 0.3664
    labMiningTitles(X).Width = tempwide * 0.2503
  Next X
  
  For X = 0 To 12
    If X < 1 Then
      Y = 0
    ElseIf X < 4 Then
      Y = 0.0467
    ElseIf X < 6 Then
      Y = 0.1308
    Else
      Y = 0.2149
    End If
    txtMiningValues(X).Top = (Y * temphigh) + (X * 0.0467 * temphigh) + (temphigh * 0.0584)
    txtMiningValues(X).Left = tempwide * 0.6623
    txtMiningValues(X).Width = tempwide * 0.1328
    labCheckTag(X).Top = (Y * temphigh) + (X * 0.0467 * temphigh) + (temphigh * 0.0654)
    labCheckTag(X).Left = tempwide * 0.9115
    labCheckTag(X).Width = tempwide * 0.0475
  Next X

  For X = 0 To 13
    If X < 4 Then
      Y = 0
    ElseIf X < 7 Then
      Y = 0.0374
    ElseIf X < 9 Then
      Y = 0.0748
    Else
      Y = 0.1215
    End If
    labMiningUnits(X).Top = (Y * temphigh) + (X * 0.0467 * temphigh) + (temphigh * 0.1589)
    labMiningUnits(X).Left = tempwide * 0.8
  Next X
  
  For X = 0 To 3
    If X = 0 Then
      Y = 0.1137
    ElseIf X = 1 Then
      Y = 0.3007
    ElseIf X = 2 Then
      Y = 0.4782
    Else
      Y = 0.8895
    End If
    labScreenTotals(X).Top = temphigh * Y
    If X = 1 Then
      labScreenTotals(X).Left = tempwide * 0.6426
      labScreenTotals(X).Width = tempwide * 0.1458
    Else
      labScreenTotals(X).Left = tempwide * 0.6689
      labScreenTotals(X).Width = tempwide * 0.1197
    End If
  Next X
  
  LineTop.X1 = tempwide * 0.3272
  LineTop.X2 = tempwide * 0.9705
  LineTop.Y1 = temphigh * 0.0187
  LineTop.Y2 = temphigh * 0.0187
  
  LineTopMiddle.X1 = tempwide * 0.3272
  LineTopMiddle.X2 = tempwide * 0.9574
  LineTopMiddle.Y1 = temphigh * 0.3458
  LineTopMiddle.Y2 = temphigh * 0.3458

  LineLeft.X1 = tempwide * 0.3337
  LineLeft.X2 = tempwide * 0.3337
  LineLeft.Y1 = temphigh * 0.0093
  LineLeft.Y2 = temphigh * 0.9439

  LineRight.X1 = tempwide * 0.9639
  LineRight.X2 = tempwide * 0.9639
  LineRight.Y1 = temphigh * 0.0093
  LineRight.Y2 = temphigh * 0.9439

  LineBottomMiddle.X1 = tempwide * 0.3272
  LineBottomMiddle.X2 = tempwide * 0.9574
  LineBottomMiddle.Y1 = temphigh * 0.5234
  LineBottomMiddle.Y2 = temphigh * 0.5234

  LineBottom.X1 = tempwide * 0.3272
  LineBottom.X2 = tempwide * 0.9705
  LineBottom.Y1 = temphigh * 0.9346
  LineBottom.Y2 = temphigh * 0.9346

  For X = 0 To 2
    If X = 0 Then
      Y = 0.0093
    ElseIf X = 1 Then
      Y = 0.3364
    Else
      Y = 0.514
    End If
    labMiningLabels(X).Top = (Y * temphigh)
    labMiningLabels(X).Left = tempwide * 0.3206
  Next X
    
  labMiningLabels(10).Top = temphigh * 0.028
  labMiningLabels(10).Left = tempwide * 0.9115
  labMiningLabels(10).Width = tempwide * 0.0475
 
  For X = 11 To 17
    If X < 13 Then
      Y = 0
    Else
      Y = 0.0467
    End If
    labMiningLabels(X).Top = (Y * temphigh) + ((X - 11) * 0.0467 * temphigh) + (temphigh * 0.5607)
    labMiningLabels(X).Left = tempwide * 0.6426
    labMiningLabels(X).Width = tempwide * 0.0148
  Next X
  
  comIndTag.Top = temphigh * 0.9455
  comIndTag.Left = tempwide * 0.4005
  
  labIndTag.Top = temphigh * 0.9439
  labIndTag.Left = tempwide * 0.4322
  
  comDepTag.Top = temphigh * 0.9455
  comDepTag.Left = tempwide * 0.7345
  
  labDepTag.Top = temphigh * 0.9439
  labDepTag.Left = tempwide * 0.7673
  
  labBackToMenu.Top = temphigh * 0.9532
  labBackToMenu.Left = tempwide * 0.0656

  imgBackToMenu.Top = temphigh * 0.9626
  imgBackToMenu.Left = tempwide * 0.0066
  imgBackToMenu.Width = tempwide * 0.0541
  
  labMiningHelp.Top = temphigh * 0.9532
  labMiningHelp.Left = tempwide * 0.9377
  
  labInsert.Top = temphigh * 0.9562
  labInsert.Left = tempwide * 0.6033
  labInsert.Width = tempwide * 0.1066
End Sub

Public Sub drawthevalues()

Dim i As Integer
Dim X As Integer

DoNotChange = True

For i = 0 To 12
  Select Case i
    Case 0
      X = 34
    Case 1 To 3
      X = 35
    Case 4, 5
      X = 36
    Case Else
      X = 39
  End Select
  
'      Primary(hscSetNumbers(5).Value, 36) = 2000000
'      Primary(hscSetNumbers(5).Value, 37) = 80
'      Primary(hscSetNumbers(5).Value, 38) = 20
'      Primary(hscSetNumbers(5).Value, 40) = 2500
   
  If i = 1 Or (i > 3 And i < 7) Then
    txtMiningValues(i).Text = Format(LTrim(Str(Primary(hscSetNumbers(5).Value, i + X))), "#########0")
  Else
    txtMiningValues(i).Text = Format(LTrim(Str(Primary(hscSetNumbers(5).Value, i + X))), "######0.00")
  End If
 
  Call CalcScreenTotals(i)
Next i

For i = 27 To 31
  labSetNumbers(i - 27).Caption = LTrim(RTrim(Str(Primary(hscSetNumbers(5).Value, i))))
Next i

For i = 34 To 51
  If i < 35 Or (i > 35 And i < 39) Or (i > 39 And i < 42) Or i > 44 Then
    Select Case i
      Case 34
        X = 0
      Case 36 To 38
        X = i - 35
      Case 40 To 41
        X = i - 36
      Case 45 To 51
        X = i - 39
    End Select
    labCheckTag(X).Visible = False
    If Tagged(hscSetNumbers(5).Value, i).Independent > 0 Then
      labCheckTag(X).Visible = True
      labCheckTag(X).ForeColor = &HFF&
      labCheckTag(X).Caption = LTrim(RTrim(Str(Tagged(hscSetNumbers(5).Value, i).Independent)))
    ElseIf Tagged(hscSetNumbers(5).Value, i).Dependent > 0 Then
      labCheckTag(X).Visible = True
      labCheckTag(X).ForeColor = &HFFFF&
      labCheckTag(X).Caption = LTrim(RTrim(Str(Tagged(hscSetNumbers(5).Value, i).Dependent)))
    Else
      labCheckTag(X).Caption = ""
    End If
  End If
Next i

DoNotChange = False

End Sub

Private Sub txtMiningValues_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

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
        If InStr(txtMiningValues(Index).Text, ".") = 0 Then
          SendKeys "{DELETE}", False
        End If
      Else
        SendKeys "{DELETE}", False
      End If
  End Select
End If

End Sub

Private Sub txtMiningValues_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii = 46 Then
  If InStr(txtMiningValues(Index).Text, ".") > 0 Then
    Beep
    KeyAscii = 0
  End If
End If

If KeyAscii = 44 Then
  Beep
  KeyAscii = 0
End If

End Sub

Private Sub txtMinSetLabel_Change()

Pn1(4, hscSetNumbers(5).Value) = txtMinSetLabel.Text

End Sub

Private Sub txtMinSetLabel_GotFocus()
LastCell = 100
End Sub

