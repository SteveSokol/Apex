VERSION 5.00
Begin VB.Form frmSmelting 
   BackColor       =   &H00000000&
   Caption         =   "Smelting and Refining"
   ClientHeight    =   6420
   ClientLeft      =   -1800
   ClientTop       =   1935
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
   Visible         =   0   'False
   Begin VB.TextBox txtComSetLabel 
      Height          =   330
      Left            =   600
      TabIndex        =   133
      TabStop         =   0   'False
      Top             =   3780
      Width           =   1035
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
      Left            =   6840
      TabIndex        =   132
      TabStop         =   0   'False
      Top             =   6120
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
      Left            =   3060
      TabIndex        =   131
      TabStop         =   0   'False
      Top             =   6120
      Width           =   195
   End
   Begin VB.HScrollBar hscSetNumbers 
      Height          =   195
      Left            =   780
      Max             =   25
      Min             =   1
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   3060
      Value           =   1
      Width           =   375
   End
   Begin VB.TextBox txtSmeltingValues 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   29
      Left            =   6720
      TabIndex        =   9
      Text            =   "0"
      Top             =   1680
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox txtSmeltingValues 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   28
      Left            =   6720
      TabIndex        =   8
      Text            =   "0"
      Top             =   1380
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox txtSmeltingValues 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   27
      Left            =   6720
      TabIndex        =   7
      Text            =   "0"
      Top             =   1080
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox txtSmeltingValues 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   26
      Left            =   6720
      TabIndex        =   6
      Text            =   "0"
      Top             =   780
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox txtSmeltingValues 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   25
      Left            =   6720
      TabIndex        =   5
      Text            =   "0"
      Top             =   480
      Width           =   675
   End
   Begin VB.TextBox txtSmeltingValues 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   24
      Left            =   4140
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox txtSmeltingValues 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   23
      Left            =   4140
      TabIndex        =   3
      Top             =   1380
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox txtSmeltingValues 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   22
      Left            =   4140
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox txtSmeltingValues 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   21
      Left            =   4140
      TabIndex        =   1
      Top             =   780
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox txtSmeltingValues 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   20
      Left            =   4140
      TabIndex        =   0
      Top             =   480
      Width           =   675
   End
   Begin VB.TextBox txtSmeltingValues 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   19
      Left            =   6720
      TabIndex        =   29
      Top             =   5640
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox txtSmeltingValues 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   18
      Left            =   6720
      TabIndex        =   28
      Top             =   5340
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox txtSmeltingValues 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   17
      Left            =   6720
      TabIndex        =   27
      Top             =   5040
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox txtSmeltingValues 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   16
      Left            =   6720
      TabIndex        =   26
      Top             =   4740
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox txtSmeltingValues 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   15
      Left            =   6720
      TabIndex        =   25
      Top             =   4440
      Width           =   675
   End
   Begin VB.TextBox txtSmeltingValues 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   14
      Left            =   4140
      TabIndex        =   24
      Top             =   5640
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox txtSmeltingValues 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   13
      Left            =   4140
      TabIndex        =   23
      Top             =   5340
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox txtSmeltingValues 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   12
      Left            =   4140
      TabIndex        =   22
      Top             =   5040
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox txtSmeltingValues 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   11
      Left            =   4140
      TabIndex        =   21
      Top             =   4740
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox txtSmeltingValues 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   10
      Left            =   4140
      TabIndex        =   20
      Top             =   4440
      Width           =   675
   End
   Begin VB.TextBox txtSmeltingValues 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   9
      Left            =   6720
      TabIndex        =   19
      Text            =   "100"
      Top             =   3660
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox txtSmeltingValues 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   8
      Left            =   6720
      TabIndex        =   18
      Text            =   "100"
      Top             =   3360
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox txtSmeltingValues 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   7
      Left            =   6720
      TabIndex        =   17
      Text            =   "100"
      Top             =   3060
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox txtSmeltingValues 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   6
      Left            =   6720
      TabIndex        =   16
      Text            =   "100"
      Top             =   2760
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox txtSmeltingValues 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   5
      Left            =   6720
      TabIndex        =   15
      Text            =   "100"
      Top             =   2460
      Width           =   675
   End
   Begin VB.TextBox txtSmeltingValues 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   4
      Left            =   4140
      TabIndex        =   14
      Text            =   "100"
      Top             =   3660
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox txtSmeltingValues 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   3
      Left            =   4140
      TabIndex        =   13
      Text            =   "100"
      Top             =   3360
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox txtSmeltingValues 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   2
      Left            =   4140
      TabIndex        =   12
      Text            =   "100"
      Top             =   3060
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox txtSmeltingValues 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   1
      Left            =   4140
      TabIndex        =   11
      Text            =   "100"
      Top             =   2760
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox txtSmeltingValues 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   0
      Left            =   4140
      TabIndex        =   10
      Text            =   "100"
      Top             =   2460
      Width           =   675
   End
   Begin VB.Label labSmeltingHelp 
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
      Left            =   8580
      TabIndex        =   137
      Top             =   6120
      Width           =   495
   End
   Begin VB.Label labinsert 
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
      Left            =   5340
      TabIndex        =   136
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label labSetNumberTitle 
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
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   135
      Top             =   3480
      Width           =   1035
   End
   Begin VB.Label labSmeltingHeading 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Refining"
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
      Index           =   2
      Left            =   720
      TabIndex        =   134
      Top             =   1080
      Width           =   1530
   End
   Begin VB.Label labDollars 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "$"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   6540
      TabIndex        =   130
      Top             =   5700
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label labDollars 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "$"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   6540
      TabIndex        =   129
      Top             =   5400
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label labDollars 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "$"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   6540
      TabIndex        =   128
      Top             =   5100
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label labDollars 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "$"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   6540
      TabIndex        =   127
      Top             =   4800
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label labDollars 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "$"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   6540
      TabIndex        =   126
      Top             =   4500
      Width           =   135
   End
   Begin VB.Label labDollars 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "$"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   3960
      TabIndex        =   125
      Top             =   1740
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label labDollars 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "$"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   3960
      TabIndex        =   124
      Top             =   1440
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label labDollars 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "$"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   3960
      TabIndex        =   123
      Top             =   1140
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label labDollars 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "$"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   3960
      TabIndex        =   122
      Top             =   840
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label labDollars 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "$"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   3960
      TabIndex        =   121
      Top             =   540
      Width           =   135
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
      Index           =   29
      Left            =   8580
      TabIndex        =   120
      Top             =   1740
      Visible         =   0   'False
      Width           =   315
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
      Index           =   28
      Left            =   8580
      TabIndex        =   119
      Top             =   1440
      Visible         =   0   'False
      Width           =   315
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
      Index           =   27
      Left            =   8580
      TabIndex        =   118
      Top             =   1140
      Visible         =   0   'False
      Width           =   315
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
      Index           =   26
      Left            =   8580
      TabIndex        =   117
      Top             =   840
      Visible         =   0   'False
      Width           =   315
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
      Index           =   25
      Left            =   8580
      TabIndex        =   116
      Top             =   540
      Visible         =   0   'False
      Width           =   315
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
      Index           =   24
      Left            =   6000
      TabIndex        =   115
      Top             =   1740
      Visible         =   0   'False
      Width           =   315
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
      Index           =   23
      Left            =   6000
      TabIndex        =   114
      Top             =   1440
      Visible         =   0   'False
      Width           =   315
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
      Index           =   22
      Left            =   6000
      TabIndex        =   113
      Top             =   1140
      Visible         =   0   'False
      Width           =   315
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
      Index           =   21
      Left            =   6000
      TabIndex        =   112
      Top             =   840
      Visible         =   0   'False
      Width           =   315
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
      Index           =   20
      Left            =   6000
      TabIndex        =   111
      Top             =   540
      Visible         =   0   'False
      Width           =   315
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
      Index           =   19
      Left            =   8580
      TabIndex        =   110
      Top             =   5700
      Visible         =   0   'False
      Width           =   315
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
      Index           =   18
      Left            =   8580
      TabIndex        =   109
      Top             =   5400
      Visible         =   0   'False
      Width           =   315
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
      Index           =   17
      Left            =   8580
      TabIndex        =   108
      Top             =   5100
      Visible         =   0   'False
      Width           =   315
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
      Index           =   16
      Left            =   8580
      TabIndex        =   107
      Top             =   4800
      Visible         =   0   'False
      Width           =   315
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
      Index           =   15
      Left            =   8580
      TabIndex        =   106
      Top             =   4500
      Visible         =   0   'False
      Width           =   315
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
      Index           =   14
      Left            =   6000
      TabIndex        =   105
      Top             =   5700
      Visible         =   0   'False
      Width           =   315
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
      Index           =   13
      Left            =   6000
      TabIndex        =   104
      Top             =   5400
      Visible         =   0   'False
      Width           =   315
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
      Left            =   6000
      TabIndex        =   103
      Top             =   5100
      Visible         =   0   'False
      Width           =   315
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
      Left            =   6000
      TabIndex        =   102
      Top             =   4800
      Visible         =   0   'False
      Width           =   315
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
      Left            =   6000
      TabIndex        =   101
      Top             =   4500
      Visible         =   0   'False
      Width           =   315
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
      Left            =   8580
      TabIndex        =   100
      Top             =   3720
      Visible         =   0   'False
      Width           =   315
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
      Left            =   8580
      TabIndex        =   99
      Top             =   3420
      Visible         =   0   'False
      Width           =   315
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
      Left            =   8580
      TabIndex        =   98
      Top             =   3120
      Visible         =   0   'False
      Width           =   315
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
      Left            =   8580
      TabIndex        =   97
      Top             =   2820
      Visible         =   0   'False
      Width           =   315
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
      Left            =   8580
      TabIndex        =   96
      Top             =   2520
      Visible         =   0   'False
      Width           =   315
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
      Left            =   6000
      TabIndex        =   95
      Top             =   3720
      Visible         =   0   'False
      Width           =   315
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
      Left            =   6000
      TabIndex        =   94
      Top             =   3420
      Visible         =   0   'False
      Width           =   315
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
      Left            =   6000
      TabIndex        =   93
      Top             =   3120
      Visible         =   0   'False
      Width           =   315
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
      Left            =   6000
      TabIndex        =   92
      Top             =   2820
      Visible         =   0   'False
      Width           =   315
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
      Left            =   6000
      TabIndex        =   91
      Top             =   2520
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label labSmeltingHeading 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "and"
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
      Index           =   1
      Left            =   420
      TabIndex        =   90
      Top             =   600
      Width           =   1515
   End
   Begin VB.Label labDepTag 
      BackColor       =   &H00000000&
      Caption         =   "Dependent Tag"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   7140
      TabIndex        =   89
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label labIndTag 
      BackColor       =   &H00000000&
      Caption         =   "Independent Tag"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   88
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label labSmeltingHeading 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Smelting"
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
      Index           =   0
      Left            =   120
      TabIndex        =   87
      Top             =   120
      Width           =   1515
   End
   Begin VB.Label labSetNumbers 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1200
      TabIndex        =   86
      Top             =   3060
      Width           =   255
   End
   Begin VB.Label labSetNumberTitle 
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
      Index           =   0
      Left            =   600
      TabIndex        =   84
      Top             =   2700
      Width           =   1035
   End
   Begin VB.Line lineBottom 
      BorderColor     =   &H00FFFF00&
      X1              =   2280
      X2              =   9060
      Y1              =   6060
      Y2              =   6060
   End
   Begin VB.Line lineMiddleBottom 
      BorderColor     =   &H00FFFF00&
      X1              =   2400
      X2              =   8940
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line lineMiddleTop 
      BorderColor     =   &H00FFFF00&
      X1              =   2400
      X2              =   8940
      Y1              =   2100
      Y2              =   2100
   End
   Begin VB.Line lineRight 
      BorderColor     =   &H00FFFF00&
      X1              =   9000
      X2              =   9000
      Y1              =   60
      Y2              =   6120
   End
   Begin VB.Line lineLeft 
      BorderColor     =   &H00FFFF00&
      X1              =   2340
      X2              =   2340
      Y1              =   60
      Y2              =   6120
   End
   Begin VB.Line lineTop 
      BorderColor     =   &H00FFFF00&
      X1              =   2280
      X2              =   9060
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label labCommodityTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   14
      Left            =   2400
      TabIndex        =   83
      Top             =   5700
      Width           =   1395
   End
   Begin VB.Label labCommodityTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   2400
      TabIndex        =   82
      Top             =   5400
      Width           =   1395
   End
   Begin VB.Label labCommodityTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   2400
      TabIndex        =   81
      Top             =   5100
      Width           =   1395
   End
   Begin VB.Label labCommodityTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   2400
      TabIndex        =   80
      Top             =   4800
      Width           =   1395
   End
   Begin VB.Label labCommodityTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   2400
      TabIndex        =   79
      Top             =   4500
      Width           =   1395
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
      TabIndex        =   78
      Top             =   6120
      Width           =   675
   End
   Begin VB.Image imgBackToMenu 
      Height          =   195
      Left            =   60
      Picture         =   "frmSmelting.frx":0000
      Stretch         =   -1  'True
      Top             =   6180
      Width           =   495
   End
   Begin VB.Label labSmeltingUnits 
      BackColor       =   &H00000000&
      Caption         =   "percent"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   29
      Left            =   7440
      TabIndex        =   77
      Top             =   1740
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label labSmeltingUnits 
      BackColor       =   &H00000000&
      Caption         =   "percent"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   28
      Left            =   7440
      TabIndex        =   76
      Top             =   1440
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label labSmeltingUnits 
      BackColor       =   &H00000000&
      Caption         =   "percent"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   27
      Left            =   7440
      TabIndex        =   75
      Top             =   1140
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label labSmeltingUnits 
      BackColor       =   &H00000000&
      Caption         =   "percent"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   26
      Left            =   7440
      TabIndex        =   74
      Top             =   840
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label labSmeltingUnits 
      BackColor       =   &H00000000&
      Caption         =   "percent"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   25
      Left            =   7440
      TabIndex        =   73
      Top             =   540
      Width           =   1035
   End
   Begin VB.Label labSmeltingUnits 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   24
      Left            =   4860
      TabIndex        =   72
      Top             =   1740
      Width           =   1035
   End
   Begin VB.Label labSmeltingUnits 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   23
      Left            =   4860
      TabIndex        =   71
      Top             =   1440
      Width           =   1035
   End
   Begin VB.Label labSmeltingUnits 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   22
      Left            =   4860
      TabIndex        =   70
      Top             =   1140
      Width           =   1035
   End
   Begin VB.Label labSmeltingUnits 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   21
      Left            =   4860
      TabIndex        =   69
      Top             =   840
      Width           =   1035
   End
   Begin VB.Label labSmeltingUnits 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   20
      Left            =   4860
      TabIndex        =   68
      Top             =   540
      Width           =   1035
   End
   Begin VB.Label labSmeltingUnits 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   19
      Left            =   7440
      TabIndex        =   67
      Top             =   5700
      Width           =   1035
   End
   Begin VB.Label labSmeltingUnits 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   18
      Left            =   7440
      TabIndex        =   66
      Top             =   5400
      Width           =   1035
   End
   Begin VB.Label labSmeltingUnits 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   17
      Left            =   7440
      TabIndex        =   65
      Top             =   5100
      Width           =   1035
   End
   Begin VB.Label labSmeltingUnits 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   16
      Left            =   7440
      TabIndex        =   64
      Top             =   4800
      Width           =   1035
   End
   Begin VB.Label labSmeltingUnits 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   15
      Left            =   7440
      TabIndex        =   63
      Top             =   4500
      Width           =   1035
   End
   Begin VB.Label labSmeltingUnits 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   14
      Left            =   4860
      TabIndex        =   62
      Top             =   5700
      Width           =   1035
   End
   Begin VB.Label labSmeltingUnits 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   4860
      TabIndex        =   61
      Top             =   5400
      Width           =   1035
   End
   Begin VB.Label labSmeltingUnits 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   4860
      TabIndex        =   60
      Top             =   5100
      Width           =   1035
   End
   Begin VB.Label labSmeltingUnits 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   4860
      TabIndex        =   59
      Top             =   4800
      Width           =   1035
   End
   Begin VB.Label labSmeltingUnits 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   4860
      TabIndex        =   58
      Top             =   4500
      Width           =   1035
   End
   Begin VB.Label labSmeltingUnits 
      BackColor       =   &H00000000&
      Caption         =   "percent"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   7440
      TabIndex        =   57
      Top             =   3720
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label labSmeltingUnits 
      BackColor       =   &H00000000&
      Caption         =   "percent"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   7440
      TabIndex        =   56
      Top             =   3420
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label labSmeltingUnits 
      BackColor       =   &H00000000&
      Caption         =   "percent"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   7440
      TabIndex        =   55
      Top             =   3120
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label labSmeltingUnits 
      BackColor       =   &H00000000&
      Caption         =   "percent"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   7440
      TabIndex        =   54
      Top             =   2820
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label labSmeltingUnits 
      BackColor       =   &H00000000&
      Caption         =   "percent"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   7440
      TabIndex        =   53
      Top             =   2520
      Width           =   1035
   End
   Begin VB.Label labSmeltingUnits 
      BackColor       =   &H00000000&
      Caption         =   "percent"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   4860
      TabIndex        =   52
      Top             =   3720
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label labSmeltingUnits 
      BackColor       =   &H00000000&
      Caption         =   "percent"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   4860
      TabIndex        =   51
      Top             =   3420
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label labSmeltingUnits 
      BackColor       =   &H00000000&
      Caption         =   "percent"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   4860
      TabIndex        =   50
      Top             =   3120
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label labSmeltingUnits 
      BackColor       =   &H00000000&
      Caption         =   "percent"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   4860
      TabIndex        =   49
      Top             =   2820
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label labSmeltingUnits 
      BackColor       =   &H00000000&
      Caption         =   "percent"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   4860
      TabIndex        =   48
      Top             =   2520
      Width           =   1035
   End
   Begin VB.Label labTagTitle 
      Alignment       =   2  'Center
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
      Height          =   255
      Index           =   1
      Left            =   8520
      TabIndex        =   47
      Top             =   240
      Width           =   435
   End
   Begin VB.Label labTagTitle 
      Alignment       =   2  'Center
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
      Height          =   255
      Index           =   0
      Left            =   5940
      TabIndex        =   46
      Top             =   240
      Width           =   435
   End
   Begin VB.Label labCommodityTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   2400
      TabIndex        =   45
      Top             =   3720
      Width           =   1395
   End
   Begin VB.Label labCommodityTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   2400
      TabIndex        =   44
      Top             =   3420
      Width           =   1395
   End
   Begin VB.Label labCommodityTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   2400
      TabIndex        =   43
      Top             =   3120
      Width           =   1395
   End
   Begin VB.Label labCommodityTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   2400
      TabIndex        =   42
      Top             =   2820
      Width           =   1395
   End
   Begin VB.Label labCommodityTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   2400
      TabIndex        =   41
      Top             =   2520
      Width           =   1395
   End
   Begin VB.Label labCommodityTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   2400
      TabIndex        =   40
      Top             =   1740
      Width           =   1395
   End
   Begin VB.Label labCommodityTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   2400
      TabIndex        =   39
      Top             =   1440
      Width           =   1395
   End
   Begin VB.Label labCommodityTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   2400
      TabIndex        =   38
      Top             =   1140
      Width           =   1395
   End
   Begin VB.Label labCommodityTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   37
      Top             =   840
      Width           =   1395
   End
   Begin VB.Label labCommodityTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   36
      Top             =   540
      Width           =   1395
   End
   Begin VB.Label labSubjectTitles 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Refining Cost"
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
      Index           =   5
      Left            =   6780
      TabIndex        =   35
      Top             =   4140
      Width           =   1665
   End
   Begin VB.Label labSubjectTitles 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Commodity Price"
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
      Index           =   4
      Left            =   4200
      TabIndex        =   34
      Top             =   180
      Width           =   1665
   End
   Begin VB.Label labSubjectTitles 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Smelter Deduction"
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
      Index           =   3
      Left            =   4200
      TabIndex        =   33
      Top             =   4140
      Width           =   1665
   End
   Begin VB.Label labSubjectTitles 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Smelter Payment"
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
      Index           =   2
      Left            =   6780
      TabIndex        =   32
      Top             =   2160
      Width           =   1665
   End
   Begin VB.Label labSubjectTitles 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Mill Recovery"
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
      Index           =   1
      Left            =   4200
      TabIndex        =   31
      Top             =   2160
      Width           =   1665
   End
   Begin VB.Label labSubjectTitles 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Price Escalation"
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
      Index           =   0
      Left            =   6780
      TabIndex        =   30
      Top             =   180
      Width           =   1665
   End
End
Attribute VB_Name = "frmSmelting"
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
  ShowMenu = False
  frmWarnTheUser.Show
Else
  If labCheckTag(LastCell).Visible = False Then
    ParamSet = False
    dTag = dTag + 1
    labCheckTag(LastCell).Visible = True
    labCheckTag(LastCell).ForeColor = &HFFFF&
    labCheckTag(LastCell).Caption = LTrim(Str(nTag))
    If LastCell < 5 Then
      Tagged(hscSetNumbers.Value, LastCell + 54).Dependent = nTag
      DepTagData(nTag, dTag).TheCell = LastCell + 54
      DepTagData(nTag, dTag).Title = "Mill Recovery - " & labCommodityTitle(LastCell).Caption
      DepTagData(nTag, dTag).Units = labSmeltingUnits(LastCell).Caption
    ElseIf LastCell < 10 Then
      Tagged(hscSetNumbers.Value, LastCell + 55).Dependent = nTag
      DepTagData(nTag, dTag).TheCell = LastCell + 55
      DepTagData(nTag, dTag).Title = "Smelter Payment - " & labCommodityTitle(LastCell).Caption
      DepTagData(nTag, dTag).Units = labSmeltingUnits(LastCell).Caption
    ElseIf LastCell < 15 Then
      Tagged(hscSetNumbers.Value, LastCell + 57).Dependent = nTag
      DepTagData(nTag, dTag).TheCell = LastCell + 57
      DepTagData(nTag, dTag).Title = "Smelter Deduction - " & labCommodityTitle(LastCell - 5).Caption
      DepTagData(nTag, dTag).Units = labSmeltingUnits(LastCell).Caption
    ElseIf LastCell < 20 Then
      Tagged(hscSetNumbers.Value, LastCell + 58).Dependent = nTag
      DepTagData(nTag, dTag).TheCell = LastCell + 58
      DepTagData(nTag, dTag).Title = "Refining Cost - " & labCommodityTitle(LastCell - 5).Caption
      DepTagData(nTag, dTag).Units = labSmeltingUnits(LastCell).Caption
   ElseIf LastCell < 25 Then
      Tagged(hscSetNumbers.Value, LastCell + 60).Dependent = nTag
      DepTagData(nTag, dTag).TheCell = LastCell + 60
      DepTagData(nTag, dTag).Title = "Price - " & labCommodityTitle(LastCell - 20).Caption
      DepTagData(nTag, dTag).Units = labSmeltingUnits(LastCell).Caption
    Else
      Tagged(hscSetNumbers.Value, LastCell + 61).Dependent = nTag
      DepTagData(nTag, dTag).TheCell = LastCell + 61
      DepTagData(nTag, dTag).Title = "Price Escalation - " & labCommodityTitle(LastCell - 15).Caption
      DepTagData(nTag, dTag).Units = labSmeltingUnits(LastCell).Caption
    End If
    DepTagData(nTag, dTag).SetNumber = hscSetNumbers.Value
  End If
  txtSmeltingValues(LastCell).SetFocus
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
  If LastCell < 5 Then
    Tagged(hscSetNumbers.Value, LastCell + 54).Independent = nTag
    IndTagData(nTag).TheCell = LastCell + 54
    IndTagData(nTag).Title = "Mill Recovery - " & labCommodityTitle(LastCell).Caption
    IndTagData(nTag).Units = labSmeltingUnits(LastCell).Caption
  ElseIf LastCell < 10 Then
    Tagged(hscSetNumbers.Value, LastCell + 55).Independent = nTag
    IndTagData(nTag).TheCell = LastCell + 55
    IndTagData(nTag).Title = "Smelter Payment - " & labCommodityTitle(LastCell).Caption
    IndTagData(nTag).Units = labSmeltingUnits(LastCell).Caption
  ElseIf LastCell < 15 Then
    Tagged(hscSetNumbers.Value, LastCell + 57).Independent = nTag
    IndTagData(nTag).TheCell = LastCell + 57
    IndTagData(nTag).Title = "Smelter Deduction - " & labCommodityTitle(LastCell - 5).Caption
    IndTagData(nTag).Units = labSmeltingUnits(LastCell).Caption
  ElseIf LastCell < 20 Then
    Tagged(hscSetNumbers.Value, LastCell + 58).Independent = nTag
    IndTagData(nTag).TheCell = LastCell + 58
    IndTagData(nTag).Title = "Refining Cost - " & labCommodityTitle(LastCell - 5).Caption
    IndTagData(nTag).Units = labSmeltingUnits(LastCell).Caption
  ElseIf LastCell < 25 Then
    Tagged(hscSetNumbers.Value, LastCell + 60).Independent = nTag
    IndTagData(nTag).TheCell = LastCell + 60
    IndTagData(nTag).Title = "Price - " & labCommodityTitle(LastCell - 20).Caption
    IndTagData(nTag).Units = labSmeltingUnits(LastCell).Caption
  Else
    Tagged(hscSetNumbers.Value, LastCell + 61).Independent = nTag
    IndTagData(nTag).TheCell = LastCell + 61
    IndTagData(nTag).Title = "Price Escalation - " & labCommodityTitle(LastCell - 15).Caption
    IndTagData(nTag).Units = labSmeltingUnits(LastCell).Caption
  End If
  IndTagData(nTag).SetNumber = hscSetNumbers.Value
End If

txtSmeltingValues(LastCell).SetFocus

End Sub

Private Sub Form_Activate()

  Dim X As Integer
  Dim Y As Integer
  
  hscSetNumbers.Value = 1
  
  If IsHelpOn = True Then
    If LastCell = 100 Then
      txtComSetLabel.SetFocus
    Else
      txtSmeltingValues(LastCell).SetFocus
    End If
    IsHelpOn = False
  Else
    For X = 0 To 4
      If LTrim(RTrim(CommodityData(hscSetNumbers.Value, X).Name)) > "                  " And X > 0 Then
        For Y = 0 To 25 Step 5
          txtSmeltingValues(Y + X).Visible = True
          labSmeltingUnits(Y + X).Visible = True
        Next Y
        For Y = 0 To 5 Step 5
          labDollars(Y + X).Visible = True
        Next Y
      ElseIf X <> 0 Then
        For Y = 0 To 25 Step 5
          txtSmeltingValues(Y + X).Visible = False
          labSmeltingUnits(Y + X).Visible = False
        Next Y
        For Y = 0 To 5 Step 5
          labDollars(Y + X).Visible = False
        Next Y
      End If
      labCommodityTitle(X).Caption = RTrim(CommodityData(hscSetNumbers.Value, X).Name)
      labCommodityTitle(X + 5).Caption = RTrim(CommodityData(hscSetNumbers.Value, X).Name)
      labCommodityTitle(X + 10).Caption = RTrim(CommodityData(hscSetNumbers.Value, X).Name)
      labSmeltingUnits(X + 10) = RTrim(CommodityData(hscSetNumbers.Value, X).Units)
      labSmeltingUnits(X + 15) = RTrim(CommodityData(hscSetNumbers.Value, X).Price)
      labSmeltingUnits(X + 20) = RTrim(CommodityData(hscSetNumbers.Value, X).Price)
    Next X

    hscSetNumbers.Value = 1

    If Pn1(5, hscSetNumbers.Value) > "" And Pn1(3, hscSetNumbers.Value) = "" Then
      Pn1(3, hscSetNumbers.Value) = Pn1(5, hscSetNumbers.Value)
    End If
    
    txtComSetLabel.Text = Pn1(3, hscSetNumbers.Value)
  
    Call drawthevalues

    ShowMenu = True
        
    If InsertFlag = True Then
       labInsert.Caption = "Insert"
    Else
      labInsert.Caption = "Typeover"
    End If
    
    LastCell = 20

    txtSmeltingValues(20).SetFocus
  End If

End Sub

Private Sub Form_Deactivate()

If ShowMenu = True Then
  frmSmelting.Hide
  Call InputMenuAccess(1)
End If

End Sub

Private Sub Form_Load()

Dim i As Integer
Dim X As Integer

If FullScreen = False Then
  frmSmelting.Top = (Screen.Height - (frmSmelting.Height + 350)) / 2
  frmSmelting.Left = (Screen.Width - frmSmelting.Width) / 2
Else
  frmSmelting.Top = 0
  frmSmelting.Left = 0
  frmSmelting.WindowState = 2
End If

If frmSmelting.Top < 0 Then frmSmelting.Top = 0
If frmSmelting.Left < 0 Then frmSmelting.Left = 0

tempwide = frmSmelting.ScaleWidth
temphigh = frmSmelting.ScaleHeight

If PageChange(4) = True Then

  For X = 0 To 4
    If LTrim(RTrim(CommodityData(1, X).Name)) > "                  " And X > 0 Then
      For i = 0 To 25 Step 5
        txtSmeltingValues(i + X).Visible = True
        labSmeltingUnits(i + X).Visible = True
      Next i
      For i = 0 To 5 Step 5
        labDollars(i + X).Visible = True
      Next i
    End If
    labCommodityTitle(X).Caption = RTrim(CommodityData(1, X).Name)
    labCommodityTitle(X + 5).Caption = RTrim(CommodityData(1, X).Name)
    labCommodityTitle(X + 10).Caption = RTrim(CommodityData(1, X).Name)
    labSmeltingUnits(X + 10) = RTrim(CommodityData(1, X).Units)
    labSmeltingUnits(X + 15) = RTrim(CommodityData(1, X).Price)
    labSmeltingUnits(X + 20) = RTrim(CommodityData(1, X).Price)
  Next X

  hscSetNumbers.Value = 1
  txtComSetLabel.Text = Pn1(3, hscSetNumbers.Value)

  Call drawthevalues
End If

Call screenstuff

End Sub
Private Sub Form_Resize()

tempwide = frmSmelting.ScaleWidth
temphigh = frmSmelting.ScaleHeight

Call screenstuff

End Sub


Private Sub Form_Unload(Cancel As Integer)

  frmSmelting.Hide
  If ShowMenu = True Then Call InputMenuAccess(1)

End Sub

Private Sub hscSetNumbers_Change()

  labSetNumbers.Caption = LTrim(RTrim(Str(hscSetNumbers.Value)))
  
  txtComSetLabel.Text = Pn1(5, hscSetNumbers.Value)
  
  If hscSetNumbers.Value > Np(5) Then
    Np(5) = hscSetNumbers.Value
  End If
  
  If Np(5) > Npna Then Npna = Np(5)
    
  Call drawthevalues
  
End Sub

Private Sub imgBackToMenu_Click()
  
  frmSmelting.Hide
  If ShowMenu = True Then Call InputMenuAccess(1)

End Sub


Private Sub labBackToMenu_Click()
  
  frmSmelting.Hide
  If ShowMenu = True Then Call InputMenuAccess(1)

End Sub


Private Sub ladDepTag_Click()

End Sub

Private Sub labSmeltingHelp_Click()

Dim begin As Integer
Dim sendindex As Integer
ShowMenu = False
WhichScreen = 4
begin = 2
Select Case LastCell
  Case 0 To 4
    sendindex = LastCell + 52
  Case 5 To 9
    sendindex = LastCell + 53
  Case 10 To 14
    sendindex = LastCell + 55
  Case 15 To 19
    sendindex = LastCell + 56
  Case 20 To 24
    sendindex = LastCell + 58
  Case 25 To 29
    sendindex = LastCell + 59
  Case 100
    sendindex = 18
    WhichScreen = 0
End Select

Call frmApexHelp.gethelptext(sendindex, begin)
frmApexHelp.Show

End Sub

Private Sub txtComSetLabel_Change()

Pn1(3, hscSetNumbers.Value) = txtComSetLabel.Text

End Sub

Private Sub txtComSetLabel_GotFocus()

LastCell = 100

End Sub


Private Sub txtSmeltingValues_Change(Index As Integer)

Dim X As Integer

If DoNotChange = True Then Exit Sub

PageChange(4) = True

Select Case Index
  Case 0 To 4
    X = 54
  Case 5 To 9
    X = 55
  Case 10 To 14
    X = 57
  Case 15 To 19
    X = 58
  Case 20 To 24
    X = 60
  Case 25 To 29
    X = 61
End Select
   
If labCheckTag(Index).Visible = True Then ParamSet = False

Primary(hscSetNumbers.Value, Index + X) = CCur(Val(txtSmeltingValues(Index).Text))

End Sub


Private Sub txtSmeltingValues_GotFocus(Index As Integer)

LastCell = Index

End Sub



Public Sub screenstuff()

  Dim X As Integer
  Dim Y As Currency
  
  For X = 0 To 2
    labSmeltingHeading(X).Top = (temphigh * 0.0187) + (X * temphigh * 0.0632)
    If X = 0 Then
      labSmeltingHeading(X).Left = tempwide * 0.0131
    ElseIf X = 1 Then
      labSmeltingHeading(X).Left = tempwide * 0.0459
    Else
      labSmeltingHeading(X).Left = tempwide * 0.0787
    End If
  Next X
  
  LineLeft.X1 = tempwide * 0.2557
  LineLeft.X2 = tempwide * 0.2557
  LineLeft.Y1 = temphigh * 0.0093
  LineLeft.Y2 = temphigh * 0.9533

  LineTop.X1 = tempwide * 0.2492
  LineTop.X2 = tempwide * 0.9902
  LineTop.Y1 = temphigh * 0.0187
  LineTop.Y2 = temphigh * 0.0187
  
  lineMiddleTop.X1 = tempwide * 0.2623
  lineMiddleTop.X2 = tempwide * 0.977
  lineMiddleTop.Y1 = temphigh * 0.3271
  lineMiddleTop.Y2 = temphigh * 0.3271
  
  lineMiddleBottom.X1 = tempwide * 0.2623
  lineMiddleBottom.X2 = tempwide * 0.977
  lineMiddleBottom.Y1 = temphigh * 0.6355
  lineMiddleBottom.Y2 = temphigh * 0.6355
  
  LineBottom.X1 = tempwide * 0.2492
  LineBottom.X2 = tempwide * 0.9902
  LineBottom.Y1 = temphigh * 0.9439
  LineBottom.Y2 = temphigh * 0.9439
  
  LineRight.X1 = tempwide * 0.9836
  LineRight.X2 = tempwide * 0.9836
  LineRight.Y1 = temphigh * 0.0093
  LineRight.Y2 = temphigh * 0.9533
  
  For X = 0 To 4
    txtSmeltingValues(X + 25).Top = (X * 0.0467 * temphigh) + (temphigh * 0.0771)
    txtSmeltingValues(X + 25).Left = tempwide * 0.7344
    txtSmeltingValues(X + 25).Width = tempwide * 0.0738
    labSmeltingUnits(X + 25).Top = (X * 0.0467 * temphigh) + (temphigh * 0.0841)
    labSmeltingUnits(X + 25).Left = tempwide * 0.8131
    labCheckTag(X + 25).Top = (X * 0.0467 * temphigh) + (temphigh * 0.0841)
    labCheckTag(X + 25).Left = tempwide * 0.9311
    labCheckTag(X + 25).Width = tempwide * 0.0475
  
    txtSmeltingValues(X).Top = (X * 0.0467 * temphigh) + (temphigh * 0.3855)
    txtSmeltingValues(X).Left = tempwide * 0.4525
    txtSmeltingValues(X).Width = tempwide * 0.0738
    labSmeltingUnits(X).Top = (X * 0.0467 * temphigh) + (temphigh * 0.3925)
    labSmeltingUnits(X).Left = tempwide * 0.5311
    labCheckTag(X).Top = (X * 0.0467 * temphigh) + (temphigh * 0.3925)
    labCheckTag(X).Left = tempwide * 0.6492
    labCheckTag(X).Width = tempwide * 0.0475
  
    txtSmeltingValues(X + 5).Top = (X * 0.0467 * temphigh) + (temphigh * 0.3855)
    txtSmeltingValues(X + 5).Left = tempwide * 0.7344
    txtSmeltingValues(X + 5).Width = tempwide * 0.0738
    labSmeltingUnits(X + 5).Top = (X * 0.0467 * temphigh) + (temphigh * 0.3925)
    labSmeltingUnits(X + 5).Left = tempwide * 0.8131
    labCheckTag(X + 5).Top = (X * 0.0467 * temphigh) + (temphigh * 0.3925)
    labCheckTag(X + 5).Left = tempwide * 0.9311
    labCheckTag(X + 5).Width = tempwide * 0.0475
  
    txtSmeltingValues(X + 10).Top = (X * 0.0467 * temphigh) + (temphigh * 0.6939)
    txtSmeltingValues(X + 10).Left = tempwide * 0.4525
    txtSmeltingValues(X + 10).Width = tempwide * 0.0738
    labSmeltingUnits(X + 10).Top = (X * 0.0467 * temphigh) + (temphigh * 0.7009)
    labSmeltingUnits(X + 10).Left = tempwide * 0.5311
    labCheckTag(X + 10).Top = (X * 0.0467 * temphigh) + (temphigh * 0.7009)
    labCheckTag(X + 10).Left = tempwide * 0.6492
    labCheckTag(X + 10).Width = tempwide * 0.0475

    txtSmeltingValues(X + 20).Top = (X * 0.0467 * temphigh) + (temphigh * 0.0771)
    txtSmeltingValues(X + 20).Left = tempwide * 0.4525
    txtSmeltingValues(X + 20).Width = tempwide * 0.0738
    labSmeltingUnits(X + 20).Top = (X * 0.0467 * temphigh) + (temphigh * 0.0841)
    labSmeltingUnits(X + 20).Left = tempwide * 0.5311
    labCheckTag(X + 20).Top = (X * 0.0467 * temphigh) + (temphigh * 0.0841)
    labCheckTag(X + 20).Left = tempwide * 0.6492
    labCheckTag(X + 20).Width = tempwide * 0.0475
    labDollars(X).Top = (X * 0.0467 * temphigh) + (temphigh * 0.0841)
    labDollars(X).Left = tempwide * 0.4328
    labDollars(X).Width = tempwide * 0.0148

    txtSmeltingValues(X + 15).Top = (X * 0.0467 * temphigh) + (temphigh * 0.6939)
    txtSmeltingValues(X + 15).Left = tempwide * 0.7344
    txtSmeltingValues(X + 15).Width = tempwide * 0.0738
    labSmeltingUnits(X + 15).Top = (X * 0.0467 * temphigh) + (temphigh * 0.7009)
    labSmeltingUnits(X + 15).Left = tempwide * 0.8131
    labCheckTag(X + 15).Top = (X * 0.0467 * temphigh) + (temphigh * 0.7009)
    labCheckTag(X + 15).Left = tempwide * 0.9311
    labCheckTag(X + 15).Width = tempwide * 0.0475
    labDollars(X + 5).Top = (X * 0.0467 * temphigh) + (temphigh * 0.7009)
    labDollars(X + 5).Left = tempwide * 0.7148
    labDollars(X + 5).Width = tempwide * 0.0148
  Next X

  For X = 0 To 14
    If X < 5 Then
      Y = 0
    ElseIf X < 10 Then
      Y = 0.0748
    Else
      Y = 0.1495
    End If
    labCommodityTitle(X).Top = (temphigh * Y) + (X * 0.0467 * temphigh) + (temphigh * 0.0841)
    labCommodityTitle(X).Left = tempwide * 0.2623
    labCommodityTitle(X).Width = tempwide * 0.1525
  Next X
  
  For X = 0 To 5
    Select Case X
      Case 0, 4
        labSubjectTitles(X).Top = temphigh * 0.028
      Case 1, 2
        labSubjectTitles(X).Top = temphigh * 0.3364
      Case 3, 5
          labSubjectTitles(X).Top = temphigh * 0.6449
    End Select
    Select Case X
      Case 1, 3, 4
        labSubjectTitles(X).Left = tempwide * 0.459
      Case 0, 2, 5
        labSubjectTitles(X).Left = tempwide * 0.741
    End Select
    labSubjectTitles(X).Width = tempwide * 0.182
  Next X
  
  For X = 0 To 1
    If X = 0 Then
      labTagTitle(X).Left = tempwide * 0.6492
    Else
      labTagTitle(X).Left = tempwide * 0.9311
    End If
    labTagTitle(X).Top = temphigh * 0.0374
    labTagTitle(X).Width = tempwide * 0.0475
  Next X
  
  labSetNumberTitle(0).Top = temphigh * 0.4112
  labSetNumberTitle(0).Left = tempwide * 0.0656
  labSetNumberTitle(0).Width = tempwide * 0.1131
   
  labSetNumberTitle(1).Top = temphigh * 0.5093
  labSetNumberTitle(1).Left = tempwide * 0.0656
  labSetNumberTitle(1).Width = tempwide * 0.1131
   
  hscSetNumbers.Top = temphigh * 0.4579
  hscSetNumbers.Left = (tempwide * 0.1058) - 188
  
  labSetNumbers.Top = temphigh * 0.4532
  labSetNumbers.Left = tempwide * 0.1443
  labSetNumbers.Width = tempwide * 0.0279
  
  txtComSetLabel.Top = temphigh * 0.5514
  txtComSetLabel.Left = tempwide * 0.0656
  txtComSetLabel.Width = tempwide * 0.1131
  
  comIndTag.Top = temphigh * 0.9549
  comIndTag.Left = tempwide * 0.3344
  
  labIndTag.Top = temphigh * 0.9533
  labIndTag.Left = tempwide * 0.3672
  
  comDepTag.Top = temphigh * 0.9549
  comDepTag.Left = tempwide * 0.7475
  
  labDepTag.Top = temphigh * 0.9533
  labDepTag.Left = tempwide * 0.7803

  labBackToMenu.Top = temphigh * 0.9532
  labBackToMenu.Left = tempwide * 0.0656

  imgBackToMenu.Top = temphigh * 0.9626
  imgBackToMenu.Left = tempwide * 0.0066
  imgBackToMenu.Width = tempwide * 0.0541

  labSmeltingHelp.Top = temphigh * 0.9532
  labSmeltingHelp.Left = tempwide * 0.9377

  labInsert.Top = temphigh * 0.9562
  labInsert.Left = tempwide * 0.5836
  labInsert.Width = tempwide * 0.1066

End Sub

Public Sub drawthevalues()

Dim i As Integer
Dim X As Integer

DoNotChange = True

For i = 0 To 29
  Select Case i
    Case 0 To 4
      X = 54
    Case 5 To 9
      X = 55
    Case 10 To 14
      X = 57
    Case 15 To 19
      X = 58
    Case 20 To 24
      X = 60
    Case 25 To 29
      X = 61
  End Select
  txtSmeltingValues(i).Text = LTrim(Str(Primary(hscSetNumbers.Value, i + X)))
  Select Case i
    Case 0 To 9, 15 To 29
      txtSmeltingValues(i).Text = Format(txtSmeltingValues(i).Text, "#####0.00")
    Case Else
      txtSmeltingValues(i).Text = Format(txtSmeltingValues(i).Text, "##0.0000")
  End Select
Next i

For i = 54 To 90
  X = -1
  Select Case i
    Case 54 To 58
      X = i - 54
    Case 60 To 64
      X = i - 55
    Case 67 To 71
      X = i - 57
    Case 73 To 77
      X = i - 58
    Case 80 To 84
      X = i - 60
    Case 86 To 90
      X = i - 61
  End Select
  If X >= 0 Then
    labCheckTag(X).Visible = False
    If Tagged(hscSetNumbers.Value, i).Independent > 0 Then
      labCheckTag(X).Visible = True
      labCheckTag(X).ForeColor = &HFF&
      labCheckTag(X).Caption = LTrim(RTrim(Str(Tagged(hscSetNumbers.Value, i).Independent)))
    ElseIf Tagged(hscSetNumbers.Value, i).Dependent > 0 Then
      labCheckTag(X).Visible = True
      labCheckTag(X).ForeColor = &HFFFF&
      labCheckTag(X).Caption = LTrim(RTrim(Str(Tagged(hscSetNumbers.Value, i).Dependent)))
    Else
      labCheckTag(X).Caption = ""
    End If
  End If
Next i

DoNotChange = False

End Sub

Private Sub txtSmeltingValues_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

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
        If InStr(txtSmeltingValues(Index).Text, ".") = 0 Then
          SendKeys "{DELETE}", False
        End If
      Else
        SendKeys "{DELETE}", False
      End If
  End Select
End If

End Sub


Private Sub txtSmeltingValues_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii = 46 Then
  If InStr(txtSmeltingValues(Index).Text, ".") > 0 Then
    Beep
    KeyAscii = 0
  End If
End If

If KeyAscii = 44 Then
  Beep
  KeyAscii = 0
End If

End Sub


