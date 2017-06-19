VERSION 5.00
Begin VB.Form frmCommodities 
   BackColor       =   &H00000000&
   Caption         =   "Commodities and Grades"
   ClientHeight    =   6420
   ClientLeft      =   1125
   ClientTop       =   1065
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
   Begin VB.CommandButton cmdCommodList 
      Caption         =   "R&emove"
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
      Left            =   3000
      TabIndex        =   9
      Top             =   3720
      Width           =   795
   End
   Begin VB.TextBox txtComSetLabel 
      Height          =   330
      Left            =   6780
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   5970
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
      Left            =   7260
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   5100
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
      Left            =   4620
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   5100
      Width           =   195
   End
   Begin VB.ListBox lstReservesUnit 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      ItemData        =   "frmCommodities.frx":0000
      Left            =   1740
      List            =   "frmCommodities.frx":0010
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   4380
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.ListBox lstCommodDepletion 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      ItemData        =   "frmCommodities.frx":003A
      Left            =   1560
      List            =   "frmCommodities.frx":0059
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   4380
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.ListBox lstCommodPrice 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      ItemData        =   "frmCommodities.frx":00CA
      Left            =   1380
      List            =   "frmCommodities.frx":00F2
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   4380
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.ListBox lstCommodUnit 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      ItemData        =   "frmCommodities.frx":0171
      Left            =   1200
      List            =   "frmCommodities.frx":01A5
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   4380
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.ListBox lstCommodType 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      ItemData        =   "frmCommodities.frx":024B
      Left            =   1020
      List            =   "frmCommodities.frx":024D
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   4380
      Width           =   1515
   End
   Begin VB.TextBox txtCommodInfo 
      Height          =   330
      Index           =   4
      Left            =   1980
      TabIndex        =   5
      Top             =   2940
      Width           =   1395
   End
   Begin VB.TextBox txtCommodInfo 
      Height          =   330
      Index           =   3
      Left            =   1980
      TabIndex        =   4
      Top             =   2520
      Width           =   1395
   End
   Begin VB.TextBox txtCommodInfo 
      Height          =   330
      Index           =   2
      Left            =   1980
      TabIndex        =   3
      Top             =   2220
      Width           =   1395
   End
   Begin VB.TextBox txtCommodInfo 
      Height          =   330
      Index           =   1
      Left            =   1980
      TabIndex        =   2
      Top             =   1920
      Width           =   1395
   End
   Begin VB.TextBox txtCommodInfo 
      Height          =   330
      Index           =   0
      Left            =   1980
      TabIndex        =   1
      Top             =   1620
      Width           =   1395
   End
   Begin VB.CommandButton cmdCommodList 
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
      Left            =   2160
      TabIndex        =   8
      Top             =   3720
      Width           =   795
   End
   Begin VB.CommandButton cmdCommodList 
      Caption         =   "&Replace"
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
      Left            =   1320
      TabIndex        =   7
      Top             =   3720
      Width           =   795
   End
   Begin VB.CommandButton cmdCommodList 
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
      Left            =   480
      TabIndex        =   6
      Top             =   3720
      Width           =   795
   End
   Begin VB.HScrollBar hscCommodityNumber 
      Height          =   255
      Left            =   1740
      Max             =   4
      TabIndex        =   57
      Top             =   1260
      Value           =   1
      Width           =   1875
   End
   Begin VB.HScrollBar hscSetNumber 
      Height          =   195
      Left            =   5820
      Max             =   25
      Min             =   1
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   6060
      Value           =   1
      Width           =   375
   End
   Begin VB.TextBox txtWallGrades 
      Height          =   330
      Index           =   4
      Left            =   6060
      MaxLength       =   7
      TabIndex        =   35
      Top             =   4440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtWallGrades 
      Height          =   330
      Index           =   3
      Left            =   6060
      MaxLength       =   7
      TabIndex        =   34
      Top             =   4140
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtWallGrades 
      Height          =   330
      Index           =   2
      Left            =   6060
      MaxLength       =   7
      TabIndex        =   33
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtWallGrades 
      Height          =   330
      Index           =   1
      Left            =   6060
      MaxLength       =   7
      TabIndex        =   32
      Top             =   3540
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtWallGrades 
      Height          =   330
      Index           =   0
      Left            =   6060
      MaxLength       =   7
      TabIndex        =   31
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtOreGrades 
      Height          =   330
      Index           =   4
      Left            =   6060
      MaxLength       =   7
      TabIndex        =   30
      Top             =   2220
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtOreGrades 
      Height          =   330
      Index           =   3
      Left            =   6060
      MaxLength       =   7
      TabIndex        =   29
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtOreGrades 
      Height          =   330
      Index           =   2
      Left            =   6060
      MaxLength       =   7
      TabIndex        =   28
      Top             =   1620
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtOreGrades 
      Height          =   330
      Index           =   1
      Left            =   6060
      MaxLength       =   7
      TabIndex        =   27
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtOreGrades 
      Height          =   330
      Index           =   0
      Left            =   6060
      MaxLength       =   7
      TabIndex        =   26
      Top             =   1020
      Width           =   1215
   End
   Begin VB.Label labGradeHelp 
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
      TabIndex        =   85
      Top             =   6120
      Width           =   495
   End
   Begin VB.Label labCommodityTitles 
      AutoSize        =   -1  'True
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
      Height          =   240
      Index           =   14
      Left            =   120
      TabIndex        =   84
      Top             =   3420
      Width           =   975
   End
   Begin VB.Line linBox2UpperMiddle 
      BorderColor     =   &H00FFFF00&
      X1              =   300
      X2              =   3960
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line2 
      X1              =   3960
      X2              =   5160
      Y1              =   3000
      Y2              =   3480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFF00&
      X1              =   0
      X2              =   3660
      Y1              =   0
      Y2              =   0
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
      Left            =   4080
      TabIndex        =   83
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label labCommodityTitles 
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
      Index           =   13
      Left            =   6780
      TabIndex        =   82
      Top             =   5700
      Width           =   1035
   End
   Begin VB.Label labCommodityTitles 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Units and Rates"
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
      Index           =   12
      Left            =   120
      TabIndex        =   81
      Top             =   4140
      Width           =   1455
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
      TabIndex        =   77
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
      TabIndex        =   76
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
      TabIndex        =   75
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
      TabIndex        =   74
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
      TabIndex        =   73
      Top             =   3300
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
      TabIndex        =   72
      Top             =   2280
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
      TabIndex        =   71
      Top             =   1980
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
      TabIndex        =   70
      Top             =   1680
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
      TabIndex        =   69
      Top             =   1380
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
      TabIndex        =   68
      Top             =   1080
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label labDepTag 
      BackColor       =   &H00000000&
      Caption         =   "Dependent Tag"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   7560
      TabIndex        =   67
      Top             =   5100
      Width           =   1335
   End
   Begin VB.Label labIndTag 
      BackColor       =   &H00000000&
      Caption         =   "Independent Tag"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   4920
      TabIndex        =   66
      Top             =   5100
      Width           =   1515
   End
   Begin VB.Label labCommodityTitles 
      BackColor       =   &H00000000&
      Caption         =   "Commodities"
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
      Index           =   11
      Left            =   120
      TabIndex        =   65
      Top             =   600
      Width           =   1215
   End
   Begin VB.Line linBox2Middle 
      BorderColor     =   &H00FFFF00&
      X1              =   300
      X2              =   3960
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line linBox2Right 
      BorderColor     =   &H00FFFF00&
      X1              =   4020
      X2              =   4020
      Y1              =   600
      Y2              =   6060
   End
   Begin VB.Line linBox2Bottom 
      BorderColor     =   &H00FFFF00&
      X1              =   180
      X2              =   4080
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line linBox2Left 
      BorderColor     =   &H00FFFF00&
      X1              =   240
      X2              =   240
      Y1              =   660
      Y2              =   6060
   End
   Begin VB.Line linBox2Top 
      BorderColor     =   &H00FFFF00&
      X1              =   180
      X2              =   4080
      Y1              =   660
      Y2              =   660
   End
   Begin VB.Label labCommodityTitles 
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
      Height          =   255
      Index           =   10
      Left            =   8340
      TabIndex        =   59
      Top             =   780
      Width           =   435
   End
   Begin VB.Label labCommodityTitles 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Grades"
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
      Index           =   9
      Left            =   4380
      TabIndex        =   58
      Top             =   360
      Width           =   675
   End
   Begin VB.Line linebottom 
      BorderColor     =   &H00FFFF00&
      X1              =   4440
      X2              =   9000
      Y1              =   4980
      Y2              =   4980
   End
   Begin VB.Line lineright 
      BorderColor     =   &H00FFFF00&
      X1              =   8940
      X2              =   8940
      Y1              =   360
      Y2              =   5040
   End
   Begin VB.Line linetop 
      BorderColor     =   &H00FFFF00&
      X1              =   4440
      X2              =   9000
      Y1              =   420
      Y2              =   420
   End
   Begin VB.Line lineleft 
      BorderColor     =   &H00FFFF00&
      X1              =   4500
      X2              =   4500
      Y1              =   360
      Y2              =   5040
   End
   Begin VB.Label labCommodityTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Price Units"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   420
      TabIndex        =   56
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label labCommodityTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Reserve Units"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   420
      TabIndex        =   55
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label labCommodityHeading 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Commodities and Grades"
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
      TabIndex        =   54
      Top             =   120
      Width           =   4095
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6240
      TabIndex        =   53
      Top             =   6030
      Width           =   255
   End
   Begin VB.Label labCommodityNumber 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   4
      Left            =   3180
      TabIndex        =   51
      Top             =   960
      Width           =   195
   End
   Begin VB.Label labCommodityNumber 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   3
      Left            =   2880
      TabIndex        =   50
      Top             =   960
      Width           =   195
   End
   Begin VB.Label labCommodityNumber 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   2
      Left            =   2580
      TabIndex        =   49
      Top             =   960
      Width           =   195
   End
   Begin VB.Label labCommodityNumber 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   1
      Left            =   2280
      TabIndex        =   48
      Top             =   960
      Width           =   195
   End
   Begin VB.Label labCommodityNumber 
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
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   0
      Left            =   1980
      TabIndex        =   47
      Top             =   960
      Width           =   195
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
      TabIndex        =   46
      Top             =   6120
      Width           =   615
   End
   Begin VB.Image imgBackToMenu 
      Height          =   195
      Left            =   60
      Picture         =   "frmCommodities.frx":024F
      Stretch         =   -1  'True
      Top             =   6180
      Width           =   495
   End
   Begin VB.Label labWallUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   4
      Left            =   7320
      TabIndex        =   45
      Top             =   4500
      Width           =   45
   End
   Begin VB.Label labWallUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   3
      Left            =   7320
      TabIndex        =   44
      Top             =   4200
      Width           =   45
   End
   Begin VB.Label labWallUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   2
      Left            =   7320
      TabIndex        =   43
      Top             =   3900
      Width           =   45
   End
   Begin VB.Label labWallUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   7320
      TabIndex        =   42
      Top             =   3600
      Width           =   45
   End
   Begin VB.Label labWallUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   7320
      TabIndex        =   41
      Top             =   3300
      Width           =   45
   End
   Begin VB.Label labOreUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   4
      Left            =   7320
      TabIndex        =   40
      Top             =   2280
      Width           =   45
   End
   Begin VB.Label labOreUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   3
      Left            =   7320
      TabIndex        =   39
      Top             =   1980
      Width           =   45
   End
   Begin VB.Label labOreUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   2
      Left            =   7320
      TabIndex        =   38
      Top             =   1680
      Width           =   45
   End
   Begin VB.Label labOreUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   7320
      TabIndex        =   37
      Top             =   1380
      Width           =   45
   End
   Begin VB.Label labOreUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   7320
      TabIndex        =   36
      Top             =   1080
      Width           =   45
   End
   Begin VB.Label labWallGrades 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   4
      Left            =   4680
      TabIndex        =   25
      Top             =   4500
      Width           =   1185
   End
   Begin VB.Label labWallGrades 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   3
      Left            =   4680
      TabIndex        =   24
      Top             =   4200
      Width           =   1185
   End
   Begin VB.Label labWallGrades 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   2
      Left            =   4680
      TabIndex        =   23
      Top             =   3900
      Width           =   1185
   End
   Begin VB.Label labWallGrades 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   4680
      TabIndex        =   22
      Top             =   3600
      Width           =   1185
   End
   Begin VB.Label labWallGrades 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   4680
      TabIndex        =   21
      Top             =   3300
      Width           =   1185
   End
   Begin VB.Label labOreGrades 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   4
      Left            =   4680
      TabIndex        =   20
      Top             =   2280
      Width           =   1185
   End
   Begin VB.Label labOreGrades 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   3
      Left            =   4680
      TabIndex        =   19
      Top             =   1980
      Width           =   1185
   End
   Begin VB.Label labOreGrades 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   2
      Left            =   4680
      TabIndex        =   18
      Top             =   1680
      Width           =   1185
   End
   Begin VB.Label labOreGrades 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
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
      Left            =   4680
      TabIndex        =   17
      Top             =   1380
      Width           =   1185
   End
   Begin VB.Label labOreGrades 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   4680
      TabIndex        =   16
      Top             =   1080
      Width           =   1185
   End
   Begin VB.Label labCommodityTitles 
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
      Index           =   6
      Left            =   5580
      TabIndex        =   15
      Top             =   5700
      Width           =   1035
   End
   Begin VB.Label labCommodityTitles 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Commodity Number"
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
      Left            =   1800
      TabIndex        =   14
      Top             =   720
      Width           =   1755
   End
   Begin VB.Label labCommodityTitles 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Wallrock"
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
      Index           =   8
      Left            =   5940
      TabIndex        =   13
      Top             =   2940
      Width           =   1455
   End
   Begin VB.Label labCommodityTitles 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Ore"
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
      Index           =   7
      Left            =   6180
      TabIndex        =   12
      Top             =   720
      Width           =   975
   End
   Begin VB.Label labCommodityTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Depletion Rate"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   420
      TabIndex        =   11
      Top             =   2580
      Width           =   1455
   End
   Begin VB.Label labCommodityTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Grade Units"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   420
      TabIndex        =   10
      Top             =   1980
      Width           =   1455
   End
   Begin VB.Label labCommodityTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Commodity"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   420
      TabIndex        =   0
      Top             =   1680
      Width           =   1455
   End
End
Attribute VB_Name = "frmCommodities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim temphigh As Single
Dim tempwide As Single

Private Sub checktag_Click(Index As Integer)

End Sub
Private Sub CheckToBail(X As Integer)
  X = 0
  Select Case LCase(LTrim(RTrim(txtCommodInfo(4).Text)))
    Case "tons"
      Select Case LCase(LTrim(RTrim(txtCommodInfo(1).Text)))
        Case "oz/ton", "ct/ton", "percent", "ppm", "percent lb", "ppm lb"
          X = 1
        Case Else
          WarnNumber = 10
      End Select
      Select Case LCase(LTrim(RTrim(txtCommodInfo(2).Text)))
        Case "/tonne", "/kilogram", "/metric carat", "/cubic yard", "/long ton"
          WarnNumber = 10
        Case Else
          X = 1
      End Select
    Case "tonnes"
      Select Case LCase(LTrim(RTrim(txtCommodInfo(1).Text)))
        Case "oz/tonne", "ct/tonne", "g/tonne", "percent", "ppm"
          X = 1
        Case Else
          WarnNumber = 10
      End Select
      Select Case LCase(LTrim(RTrim(txtCommodInfo(2).Text)))
        Case "/pound", "/ton", "/short ton unit", "/carat", "/cubic yard", "/long ton"
          WarnNumber = 10
        Case Else
          X = 1
      End Select
    Case "long tons"
      Select Case LCase(LTrim(RTrim(txtCommodInfo(1).Text)))
        Case "oz/long ton", "ct/long ton", "dwt/long ton", "percent", "ppm", "percent lb", "ppm lb"
          X = 1
        Case Else
          WarnNumber = 10
      End Select
      Select Case LCase(LTrim(RTrim(txtCommodInfo(2).Text)))
        Case "/short ton unit", "/tonne", "/kilogram", "/metric carat", "/cubic yard"
          WarnNumber = 10
        Case Else
          X = 1
      End Select
    Case "cubic yards"
      Select Case LCase(LTrim(RTrim(txtCommodInfo(1).Text)))
        Case "oz/cu.yd.", "ct/cu.yd.", "g/cu.yd.", "percent", "ppm", "lb/cu.yd."
          X = 1
        Case Else
          WarnNumber = 10
      End Select
      Select Case LCase(LTrim(RTrim(txtCommodInfo(2).Text)))
        Case "/short ton unit", "/tonne", "/kilogram", "/metric carat", "/long ton"
          WarnNumber = 10
        Case Else
          X = 1
      End Select
  End Select
  IsWarnOn = True
  If WarnNumber = 10 Then
    frmWarnTheUser.Show
    DoNotChange = True
    ShowMenu = False
  Else
    DoNotChange = False
    ShowMenu = True
  End If
End Sub
Private Sub cmdCommodList_Click(Index As Integer)

Dim UnitNumber As Integer
Dim PriceNumber As Integer
Dim ReserveNumber As Integer
Dim TestNumber As Integer
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim X As Integer

If Index = 0 Or Index = 1 Then
  Call CheckToBail(X)
End If

If DoNotChange = True Then Exit Sub

PageChange(1) = True

Select Case Index
  Case 0
    For i = 1 To 25
      CommodityData(i, hscCommodityNumber.Value).Name = txtCommodInfo(0).Text
      CommodityData(i, hscCommodityNumber.Value).Units = txtCommodInfo(1).Text
      CommodityData(i, hscCommodityNumber.Value).Price = txtCommodInfo(2).Text
      CommodityData(i, hscCommodityNumber.Value).Depletion = CCur(Val(txtCommodInfo(3).Text))
    Next i
    txtOreGrades(hscCommodityNumber.Value).Visible = True
    txtWallGrades(hscCommodityNumber.Value).Visible = True
    labOreGrades(hscCommodityNumber.Value).Caption = RTrim(CommodityData(hscSetNumber.Value, hscCommodityNumber.Value).Name)
    labOreUnits(hscCommodityNumber.Value).Caption = CommodityData(hscSetNumber.Value, hscCommodityNumber.Value).Units
    labWallGrades(hscCommodityNumber.Value).Caption = RTrim(CommodityData(hscSetNumber.Value, hscCommodityNumber.Value).Name)
    labWallUnits(hscCommodityNumber.Value).Caption = CommodityData(hscSetNumber.Value, hscCommodityNumber.Value).Units
    If hscCommodityNumber.Value = 0 Then
      For j = 1 To 25
        CommodityData(j, hscCommodityNumber.Value).reserves = txtCommodInfo(4).Text
        For i = 1 To 5
          CommodityData(j, i).reserves = txtCommodInfo(4).Text
        Next i
      Next j
    End If
    For j = 1 To 25
      Select Case Left(txtCommodInfo(3).Text, 2)
        Case "0 "
          Primary(j, hscCommodityNumber.Value + 15) = 0
        Case "5 "
          Primary(j, hscCommodityNumber.Value + 15) = 5
        Case "7."
          Primary(j, hscCommodityNumber.Value + 15) = 7.5
        Case "8 "
          Primary(j, hscCommodityNumber.Value + 15) = 8
        Case "10"
          Primary(j, hscCommodityNumber.Value + 15) = 10
        Case "11"
          Primary(j, hscCommodityNumber.Value + 15) = 11.2
        Case "14"
          Primary(j, hscCommodityNumber.Value + 15) = 14
        Case "15"
          Primary(j, hscCommodityNumber.Value + 15) = 15
        Case "22"
          Primary(j, hscCommodityNumber.Value + 15) = 22
      End Select
    Next j
  Case 1
    For j = 1 To 25
      CommodityData(j, hscCommodityNumber.Value).Name = txtCommodInfo(0).Text
      CommodityData(j, hscCommodityNumber.Value).Units = txtCommodInfo(1).Text
      CommodityData(j, hscCommodityNumber.Value).Price = txtCommodInfo(2).Text
      CommodityData(j, hscCommodityNumber.Value).Depletion = CCur(Val(txtCommodInfo(3).Text))
    Next j
    labOreGrades(hscCommodityNumber.Value).Caption = RTrim(CommodityData(hscSetNumber.Value, hscCommodityNumber.Value).Name)
    labOreUnits(hscCommodityNumber.Value).Caption = CommodityData(hscSetNumber.Value, hscCommodityNumber.Value).Units
    labWallGrades(hscCommodityNumber.Value).Caption = RTrim(CommodityData(hscSetNumber.Value, hscCommodityNumber.Value).Name)
    labWallUnits(hscCommodityNumber.Value).Caption = CommodityData(hscSetNumber.Value, hscCommodityNumber.Value).Units
    If hscCommodityNumber.Value = 0 Then
      For j = 1 To 25
        CommodityData(j, hscCommodityNumber.Value).reserves = txtCommodInfo(4).Text
        For i = 1 To 5
          CommodityData(j, i).reserves = txtCommodInfo(4).Text
        Next i
      Next j
    End If
    For j = 1 To 25
      'Primary(j, hscCommodityNumber.Value + 2) = 0
      'Primary(j, hscCommodityNumber.Value + 8) = 0
      
      'Primary(j, hscCommodityNumber.Value + 54) = 100
      'Primary(j, hscCommodityNumber.Value + 60) = 100
      'Primary(j, hscCommodityNumber.Value + 67) = 0
      'Primary(j, hscCommodityNumber.Value + 73) = 0
      'Primary(j, hscCommodityNumber.Value + 80) = 0
      'Primary(j, hscCommodityNumber.Value + 86) = 0
      
      Select Case Left(txtCommodInfo(3).Text, 2)
        Case "0 "
          Primary(j, hscCommodityNumber.Value + 15) = 0
        Case "5 "
          Primary(j, hscCommodityNumber.Value + 15) = 5
        Case "7."
          Primary(j, hscCommodityNumber.Value + 15) = 7.5
        Case "8 "
          Primary(j, hscCommodityNumber.Value + 15) = 8
        Case "10"
          Primary(j, hscCommodityNumber.Value + 15) = 10
        Case "11"
          Primary(j, hscCommodityNumber.Value + 15) = 11.2
        Case "14"
          Primary(j, hscCommodityNumber.Value + 15) = 14
        Case "15"
          Primary(j, hscCommodityNumber.Value + 15) = 15
        Case "22"
          Primary(j, hscCommodityNumber.Value + 15) = 22
      End Select
    Next j
  Case 2
    If Primary(hscSetNumber.Value, hscCommodityNumber.Value + 3) > 0 Then
      For i = hscCommodityNumber.Value + 1 To 5
        If Primary(hscSetNumber.Value, i + 2) > 0 Then
          txtOreGrades(i - 1).Visible = True
          txtWallGrades(i - 1).Visible = True
        Else
          txtOreGrades(i - 1).Visible = False
          txtWallGrades(i - 1).Visible = False
        End If
        For j = 1 To 25
          Primary(j, i + 1) = Primary(j, i + 2)
          Primary(j, i + 7) = Primary(j, i + 8)
        
          Primary(j, i + 53) = Primary(j, i + 54)
          Primary(j, i + 59) = Primary(j, i + 60)
          Primary(j, i + 66) = Primary(j, i + 67)
          Primary(j, i + 72) = Primary(j, i + 73)
          Primary(j, i + 79) = Primary(j, i + 80)
          Primary(j, i + 85) = Primary(j, i + 86)
          
          CommodityData(j, i - 1).Name = CommodityData(j, i).Name
          CommodityData(j, i - 1).Units = CommodityData(j, i).Units
          CommodityData(j, i - 1).Price = CommodityData(j, i).Price
          CommodityData(j, i - 1).Depletion = CommodityData(j, i).Depletion
        Next j
        
        txtOreGrades(i - 1).Text = Format(LTrim(Str(Primary(hscSetNumber.Value, i + 1))), "###0.0000")
        txtWallGrades(i - 1).Text = Format(LTrim(Str(Primary(hscSetNumber.Value, i + 7))), "###0.0000")
                
        labOreGrades(i - 1).Caption = RTrim(CommodityData(hscSetNumber.Value, i - 1).Name)
        labOreUnits(i - 1).Caption = CommodityData(hscSetNumber.Value, i - 1).Units
        labWallGrades(i - 1).Caption = RTrim(CommodityData(hscSetNumber.Value, i - 1).Name)
        labWallUnits(i - 1).Caption = CommodityData(hscSetNumber.Value, i - 1).Units
      
        Cf(i - 1) = Cf(i)
      
      Next i
      txtOreGrades(4).Visible = False
      txtWallGrades(4).Visible = False
            
      txtOreGrades(4).Text = ""
      txtWallGrades(4).Text = ""
      
      For j = 1 To 25
        Primary(j, 6) = 0
        Primary(j, 12) = 0
      
        Primary(j, 58) = 100
        Primary(j, 64) = 100
        Primary(j, 71) = 0
        Primary(j, 77) = 0
        Primary(j, 84) = 0
        Primary(j, 90) = 0
        
        CommodityData(j, 4).Name = ""
        CommodityData(j, 4).Units = ""
        CommodityData(j, 4).Price = ""
        CommodityData(j, 4).Depletion = 0
      Next j
      
      labOreGrades(4).Caption = ""
      labWallGrades(4).Caption = ""
      labOreUnits(4).Caption = ""
      labWallUnits(4).Caption = ""
    
    Else
      txtOreGrades(hscCommodityNumber.Value).Visible = False
      txtWallGrades(hscCommodityNumber.Value).Visible = False
      
      For j = 1 To 25
        Primary(j, hscCommodityNumber.Value + 2) = 0
        Primary(j, hscCommodityNumber.Value + 8) = 0
      
        Primary(j, hscCommodityNumber.Value + 54) = 100
        Primary(j, hscCommodityNumber.Value + 60) = 100
        Primary(j, hscCommodityNumber.Value + 67) = 0
        Primary(j, hscCommodityNumber.Value + 73) = 0
        Primary(j, hscCommodityNumber.Value + 80) = 0
        Primary(j, hscCommodityNumber.Value + 86) = 0
      
        txtOreGrades(hscCommodityNumber.Value).Text = ""
        txtWallGrades(hscCommodityNumber.Value).Text = ""
      
        CommodityData(j, hscCommodityNumber.Value).Name = ""
        CommodityData(j, hscCommodityNumber.Value).Units = ""
        CommodityData(j, hscCommodityNumber.Value).Price = ""
        CommodityData(j, hscCommodityNumber.Value).Depletion = 0
      Next j
      
      txtOreGrades(hscCommodityNumber.Value).Text = ""
      txtWallGrades(hscCommodityNumber.Value).Text = ""
      
      labOreGrades(hscCommodityNumber.Value).Caption = ""
      labWallGrades(hscCommodityNumber.Value).Caption = ""
      labOreUnits(hscCommodityNumber.Value).Caption = ""
      labWallUnits(hscCommodityNumber.Value).Caption = ""
    End If
    For i = 0 To 4
      If Primary(hscSetNumber.Value, i + 2) > 0 Then k = i
    Next i
    txtCommodInfo(0).Text = CommodityData(hscSetNumber.Value, k).Name
    txtCommodInfo(1).Text = CommodityData(hscSetNumber.Value, k).Units
    txtCommodInfo(2).Text = CommodityData(hscSetNumber.Value, k).Price
    txtCommodInfo(3).Text = LTrim(RTrim(Str(CommodityData(hscSetNumber.Value, k).Depletion)))
    hscCommodityNumber.Value = k
  Case 3
    For i = 0 To 4
      For k = 1 To 25
        CommodityData(k, i).Name = ""
        CommodityData(k, i).Units = ""
        CommodityData(k, i).Price = ""
        CommodityData(k, i).Depletion = 0
        CommodityData(k, i).reserves = ""
        Primary(k, i + 15) = 22
      Next k
      k = 0
      labOreGrades(i).Caption = CommodityData(1, i).Name
      labOreUnits(i).Caption = CommodityData(1, i).Units
      labWallGrades(i).Caption = CommodityData(1, i).Name
      labWallUnits(i).Caption = CommodityData(1, i).Units
      If i > 0 Then
        txtOreGrades(i).Visible = False
        txtWallGrades(i).Visible = False
      Else
        txtOreGrades(i).Text = "0.00"
        txtWallGrades(i).Text = "0.00"
      End If
      If i = 3 Then
        txtCommodInfo(i).Text = "0"
      Else
        txtCommodInfo(i).Text = ""
      End If
      For j = 1 To 25
        Primary(j, i + 2) = 0
        Primary(j, i + 8) = 0
        Primary(j, i + 54) = 100
        Primary(j, i + 60) = 100
        Primary(j, i + 67) = 0
        Primary(j, i + 73) = 0
        Primary(j, i + 80) = 0
        Primary(j, i + 86) = 0
      Next j
    Next i
  End Select
  
TestNumber = lstCommodUnit.ListIndex
  
Select Case LTrim(RTrim(LCase(txtCommodInfo(1).Text)))
  Case "oz/ton"
    UnitNumber = 0
  Case "oz/tonne"
    UnitNumber = 1
  Case "oz/long ton"
    UnitNumber = 2
  Case "oz/cu.yd."
    UnitNumber = 3
  Case "ct/ton"
    UnitNumber = 4
  Case "ct/tonne"
    UnitNumber = 5
  Case "ct/long ton"
    UnitNumber = 6
  Case "ct/cu.yd."
    UnitNumber = 7
  Case "g/tonne"
    UnitNumber = 8
  Case "g/cu.yd."
    UnitNumber = 9
  Case "dwt/long ton"
    UnitNumber = 10
  Case "percent"
    UnitNumber = 11
  Case "ppm"
    UnitNumber = 12
  Case "percent lb"
    UnitNumber = 13
  Case "ppm lb"
    UnitNumber = 14
  Case "lb/cu.yd."
    UnitNumber = 15
End Select
    
TestNumber = lstCommodPrice.ListIndex
    
Select Case LTrim(RTrim(LCase(txtCommodInfo(2).Text)))
  Case "/troy ounce"
    PriceNumber = 0
  Case "/pound"
    PriceNumber = 1
  Case "/ton"
    PriceNumber = 2
  Case "/short ton unit"
    PriceNumber = 3
  Case "/tonne"
    PriceNumber = 4
  Case "/carat"
    PriceNumber = 5
  Case "/kilogram"
    PriceNumber = 6
  Case "/metric carat"
    PriceNumber = 7
  Case "/dwt"
    PriceNumber = 8
  Case "/avdp"
    PriceNumber = 9
  Case "/cubic yard"
    PriceNumber = 10
  Case "/long ton"
    PriceNumber = 11
End Select

TestNumber = lstReservesUnit.ListIndex

Select Case LTrim(RTrim(LCase(txtCommodInfo(4).Text)))
  Case "tons"
    ReserveNumber = 0
  Case "tonnes"
    ReserveNumber = 1
  Case "long tons"
    ReserveNumber = 2
  Case "cubic yards"
    ReserveNumber = 3
End Select

Select Case UnitNumber
  Case 0 To 7, 15
    Cf(hscCommodityNumber.Value + 1) = 1
  Case 8
    Select Case PriceNumber
      Case 0
        Cf(hscCommodityNumber.Value + 1) = 0.0321543
      Case 7
        Cf(hscCommodityNumber.Value + 1) = 0.05
    End Select
  Case 9
    Select Case PriceNumber
      Case 0
        Cf(hscCommodityNumber.Value + 1) = 0.0321543
      Case 5
        Cf(hscCommodityNumber.Value + 1) = 0.05
    End Select
  Case 10
    Select Case PriceNumber
      Case 0
        Cf(hscCommodityNumber.Value + 1) = 0.05
      Case 8
        Cf(hscCommodityNumber.Value + 1) = 1
    End Select
  Case 11
    Select Case PriceNumber
      Case 1
        Cf(hscCommodityNumber.Value + 1) = 20
      Case 2, 4, 11
        Cf(hscCommodityNumber.Value + 1) = 0.01
      Case 3, 10
        Cf(hscCommodityNumber.Value + 1) = 1
      Case 6
        Cf(hscCommodityNumber.Value + 1) = 10
    End Select
  Case 12
    Select Case PriceNumber
      Case 0
        Select Case ReserveNumber
          Case 0
            Cf(hscCommodityNumber.Value + 1) = 0.02917
          Case 1
            Cf(hscCommodityNumber.Value + 1) = 0.0321543
          Case 2
            Cf(hscCommodityNumber.Value + 1) = 0.032671
        End Select
      Case 6
        Cf(hscCommodityNumber.Value + 1) = 0.001
    End Select
  Case 13
    Select Case ReserveNumber
      Case 0
        Cf(hscCommodityNumber.Value + 1) = 20
      Case 2
        Cf(hscCommodityNumber.Value + 1) = 22.4
    End Select
  Case 14
    Select Case ReserveNumber
      Case 0
        Cf(hscCommodityNumber.Value + 1) = 0.002
      Case 2
        Cf(hscCommodityNumber.Value + 1) = 22.4
    End Select
End Select

labOreGrades(hscCommodityNumber.Value).Caption = RTrim(CommodityData(hscSetNumber.Value, hscCommodityNumber.Value).Name)
labOreUnits(hscCommodityNumber.Value).Caption = CommodityData(hscSetNumber.Value, hscCommodityNumber.Value).Units
labWallGrades(hscCommodityNumber.Value).Caption = RTrim(CommodityData(hscSetNumber.Value, hscCommodityNumber.Value).Name)
labWallUnits(hscCommodityNumber.Value).Caption = CommodityData(hscSetNumber.Value, hscCommodityNumber.Value).Units
  
If txtOreGrades(k).Visible = True Then
  txtOreGrades(k).SetFocus
Else
  txtCommodInfo(0).SetFocus
End If

End Sub

Private Sub cmdCommodList_GotFocus(Index As Integer)
LastCell = Index + 15
End Sub

Private Sub comDepTag_Click()

If LastCell > 9 Then Exit Sub

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
    If LastCell < 5 Then
      Tagged(hscSetNumber.Value, LastCell + 2).Dependent = nTag
      DepTagData(nTag, dTag).TheCell = LastCell + 2
      DepTagData(nTag, dTag).Title = labOreGrades(LastCell).Caption & " - Ore Grade"
      DepTagData(nTag, dTag).Units = labOreUnits(LastCell).Caption
    Else
      Tagged(hscSetNumber.Value, LastCell + 3).Dependent = nTag
      DepTagData(nTag, dTag).TheCell = LastCell + 3
      DepTagData(nTag, dTag).Title = labWallGrades(LastCell - 5).Caption & " - Wallrock Grade"
      DepTagData(nTag, dTag).Units = labWallUnits(LastCell - 5).Caption
    End If
    DepTagData(nTag, dTag).SetNumber = hscSetNumber.Value
  End If
  If LastCell < 5 Then
    txtOreGrades(LastCell).SetFocus
  Else
    txtWallGrades(LastCell - 5).SetFocus
  End If
End If

End Sub

Private Sub comGradeHelp_Click()
End Sub

Private Sub comIndTag_Click()

If LastCell > 9 Then Exit Sub

If labCheckTag(LastCell).Visible = False Then
  ParamSet = False
  dTag = 0
  nTag = nTag + 1
  labCheckTag(LastCell).Visible = True
  labCheckTag(LastCell).ForeColor = &HFF&
  labCheckTag(LastCell).Caption = LTrim(Str(nTag))
  If LastCell < 5 Then
    Tagged(hscSetNumber.Value, LastCell + 2).Independent = nTag
    IndTagData(nTag).TheCell = LastCell + 2
    IndTagData(nTag).Title = labOreGrades(LastCell).Caption & " - Ore Grade"
    IndTagData(nTag).Units = labOreUnits(LastCell).Caption
  Else
    Tagged(hscSetNumber.Value, LastCell + 3).Independent = nTag
    IndTagData(nTag).TheCell = LastCell + 3
    IndTagData(nTag).Title = labWallGrades(LastCell - 5).Caption & " - Wallrock Grade"
    IndTagData(nTag).Units = labWallUnits(LastCell - 5).Caption
  End If
  IndTagData(nTag).SetNumber = hscSetNumber.Value
End If

If LastCell < 5 Then
  txtOreGrades(LastCell).SetFocus
Else
  txtWallGrades(LastCell - 5).SetFocus
End If

End Sub

Private Sub Form_Activate()

  Dim X As Integer
  
  If IsHelpOn = True Then
    If LastCell < 5 Then
      frmCommodities.txtOreGrades(LastCell).SetFocus
    ElseIf LastCell < 10 Then
      frmCommodities.txtWallGrades(LastCell - 5).SetFocus
    ElseIf LastCell < 15 Then
      frmCommodities.txtCommodInfo(LastCell - 10).SetFocus
    ElseIf LastCell = 100 Then
      frmCommodities.txtComSetLabel.SetFocus
    Else
      cmdCommodList(LastCell - 15).SetFocus
    End If
    IsHelpOn = False
  ElseIf IsWarnOn = True Then
    IsWarnOn = False
  Else
    hscCommodityNumber.Value = 0
    hscSetNumber.Value = 1
    txtComSetLabel.Text = Pn1(3, hscSetNumber.Value)
  
    DoNotChange = True
      ShowMenu = True
      Call drawthevalues
      If PageChange(1) = True Then
        txtCommodInfo(0).Text = CommodityData(hscSetNumber.Value, 0).Name
        txtCommodInfo(1).Text = CommodityData(hscSetNumber.Value, 0).Units
        txtCommodInfo(2).Text = CommodityData(hscSetNumber.Value, 0).Price
        txtCommodInfo(3).Text = LTrim(Str(CommodityData(hscSetNumber.Value, 0).Depletion))
        txtCommodInfo(4).Text = CommodityData(hscSetNumber.Value, 0).reserves
      End If
    
    DoNotChange = False
    
    LastCell = 0
    
    If InsertFlag = True Then
      labInsert.Caption = "Insert"
    Else
      labInsert.Caption = "Typeover"
    End If
    
   txtCommodInfo(0).SetFocus
  
  End If
  
End Sub

Private Sub txtCommodity_Change()

End Sub

Private Sub txtUnit_Change(Index As Integer)

End Sub

Private Sub Form_Deactivate()
 
  If ShowMenu = True Then
    frmCommodities.Hide
    Call InputMenuAccess(1)
  End If

End Sub

Private Sub Form_Load()

Dim X As Integer

If FullScreen = False Then
  frmCommodities.Top = (Screen.Height - (frmCommodities.Height + 350)) / 2
  frmCommodities.Left = (Screen.Width - frmCommodities.Width) / 2
Else
  frmCommodities.Top = 0
  frmCommodities.Left = 0
  frmCommodities.WindowState = 2
End If

If frmCommodities.Top < 0 Then frmCommodities.Top = 0
If frmCommodities.Left < 0 Then frmCommodities.Left = 0

tempwide = frmCommodities.ScaleWidth
temphigh = frmCommodities.ScaleHeight

Call drawthevalues

If PageChange(1) = True Then
  txtCommodInfo(0).Text = CommodityData(1, 0).Name
  txtCommodInfo(1).Text = CommodityData(1, 0).Units
  txtCommodInfo(2).Text = CommodityData(1, 0).Price
  txtCommodInfo(3).Text = LTrim(Str(CommodityData(1, 0).Depletion))
  txtCommodInfo(4).Text = CommodityData(1, 0).reserves
End If

End Sub

Private Sub Form_Resize()

tempwide = frmCommodities.ScaleWidth
temphigh = frmCommodities.ScaleHeight

Call screenstuff

End Sub

Private Sub Form_Unload(Cancel As Integer)
  
frmCommodities.Hide
If ShowMenu = True Then Call InputMenuAccess(1)
  
End Sub

Private Sub hscCommodityNumber_Change()
  
  Dim i As Integer
  For i = 0 To 4
    labCommodityNumber(i).ForeColor = &HFFFF00
  Next i
  
  txtCommodInfo(0).Text = CommodityData(hscSetNumber.Value, hscCommodityNumber.Value).Name
  txtCommodInfo(1).Text = CommodityData(hscSetNumber.Value, hscCommodityNumber.Value).Units
  txtCommodInfo(2).Text = CommodityData(hscSetNumber.Value, hscCommodityNumber.Value).Price
  txtCommodInfo(3).Text = LTrim(Str(CommodityData(hscSetNumber.Value, hscCommodityNumber.Value).Depletion))
  txtCommodInfo(4).Text = CommodityData(hscSetNumber.Value, hscCommodityNumber.Value).reserves
  
  labCommodityNumber(hscCommodityNumber.Value).ForeColor = &HFFFF&
  txtCommodInfo(0).SetFocus

End Sub

Private Sub hscSetNumber_Change()

  labSetNumber.Caption = LTrim(RTrim(Str(hscSetNumber.Value)))
  
  txtComSetLabel.Text = Pn1(3, hscSetNumber.Value)
  
  If hscSetNumber.Value > Np(3) Then
    Np(3) = hscSetNumber.Value
  End If
  
  If Np(3) > Npna Then Npna = Np(3)
  
  Call drawthevalues
  
  txtOreGrades(0).SetFocus

End Sub

Private Sub imgBackToMenu_Click()
  
  frmCommodities.Hide
  If ShowMenu = True Then Call InputMenuAccess(1)

End Sub

Private Sub labBackToMenu_Click()
  
  frmCommodities.Hide
  If ShowMenu = True Then Call InputMenuAccess(1)

End Sub

Private Sub lstCommodityNumber_Click()

End Sub

Private Sub VScroll1_Change()

End Sub

Private Sub labGradeHelp_Click()

Dim begin As Integer
Dim sendindex As Integer
ShowMenu = False
begin = 4

WhichScreen = 1

If LastCell < 10 Then
  sendindex = LastCell + 6
ElseIf LastCell < 15 Then
  sendindex = LastCell - 9
ElseIf LastCell = 100 Then
  WhichScreen = 0
  sendindex = 16
Else
  sendindex = LastCell + 2
End If

Call frmApexHelp.gethelptext(sendindex, begin)
frmApexHelp.Show
 
End Sub

Private Sub lstCommodDepletion_Click()

txtCommodInfo(3).Text = LTrim(lstCommodDepletion.List(lstCommodDepletion.ListIndex))

If hscCommodityNumber.Value < 1 Then
  txtCommodInfo(4).SetFocus
Else
  cmdCommodList(0).SetFocus
End If

End Sub

Private Sub lstCommodPrice_Click()

txtCommodInfo(2).Text = lstCommodPrice.List(lstCommodPrice.ListIndex)

txtCommodInfo(3).SetFocus

End Sub


Private Sub lstCommodType_Click()

txtCommodInfo(0).Text = lstCommodType.List(lstCommodType.ListIndex)

txtCommodInfo(1).SetFocus

End Sub

Private Sub lstCommodUnit_Click()

txtCommodInfo(1).Text = lstCommodUnit.List(lstCommodUnit.ListIndex)

txtCommodInfo(2).SetFocus

End Sub


Private Sub lstReservesUnit_Click()

txtCommodInfo(4).Text = lstReservesUnit.List(lstReservesUnit.ListIndex)

cmdCommodList(0).SetFocus

End Sub

Private Sub txtCommodInfo_GotFocus(Index As Integer)

If Index = 0 Then
  lstCommodType.Visible = True
Else
  lstCommodType.Visible = False
End If

If Index = 1 Then
  lstCommodUnit.Visible = True
Else
  lstCommodUnit.Visible = False
End If

If Index = 2 Then
  lstCommodPrice.Visible = True
Else
  lstCommodPrice.Visible = False
End If

If Index = 3 Then
  lstCommodDepletion.Visible = True
Else
  lstCommodDepletion.Visible = False
End If

If Index = 4 Then
  lstReservesUnit.Visible = True
Else
  lstReservesUnit.Visible = False
End If

LastCell = 10 + Index

End Sub


Private Sub txtCommodInfo_KeyPress(Index As Integer, KeyAscii As Integer)
  If Index > 0 Then
    If KeyAscii >= Asc("0") Or KeyAscii <= Asc("9") Then
      KeyAscii = 0
      Beep
    End If
  End If
End Sub


Private Sub txtComSetLabel_Change()

Pn1(3, hscSetNumber.Value) = txtComSetLabel.Text

End Sub

Private Sub txtComSetLabel_GotFocus()

LastCell = 100

End Sub


Private Sub txtOreGrades_Change(Index As Integer)
  
  If DoNotChange = True Then Exit Sub
  
  If labCheckTag(Index).Visible = True Then ParamSet = False
   
  Primary(hscSetNumber.Value, Index + 2) = CCur(Val(txtOreGrades(Index).Text))

End Sub

Private Sub txtOreGrades_GotFocus(Index As Integer)

labInsert.Visible = True
LastCell = Index

End Sub


Private Sub txtOreGrades_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

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
        If InStr(txtOreGrades(Index).Text, ".") = 0 Then
          SendKeys "{DELETE}", False
        End If
      Else
        SendKeys "{DELETE}", False
      End If
  End Select
End If

End Sub

Private Sub txtOreGrades_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii = 46 Then
  If InStr(txtOreGrades(Index).Text, ".") > 0 Then
    Beep
    KeyAscii = 0
  End If
End If

If KeyAscii = 44 Then
  Beep
  KeyAscii = 0
End If

End Sub

Private Sub txtOreGrades_LostFocus(Index As Integer)
If IsWarnOn = False Then labInsert.Visible = False
End Sub

Private Sub txtWallGrades_Change(Index As Integer)
  
  If DoNotChange = True Then Exit Sub
  
  If labCheckTag(Index + 5).Visible = True Then ParamSet = False
    
  Primary(hscSetNumber.Value, Index + 8) = CCur(Val(txtWallGrades(Index).Text))

End Sub

Private Sub txtWallGrades_GotFocus(Index As Integer)

labInsert.Visible = True
LastCell = Index + 5

End Sub

Public Sub screenstuff()
  
  Dim X As Integer
  Dim Y As Currency
  
  labCommodityHeading.Top = temphigh * 0.0334
  labCommodityHeading.Left = tempwide * 0.0194
  
  LineLeft.X1 = tempwide * 0.4918
  LineLeft.X2 = tempwide * 0.4918
  LineLeft.Y1 = temphigh * 0.0561
  LineLeft.Y2 = temphigh * 0.785

  LineTop.X1 = tempwide * 0.4852
  LineTop.X2 = tempwide * 0.9836
  LineTop.Y1 = temphigh * 0.0654
  LineTop.Y2 = temphigh * 0.0654
  
  LineBottom.X1 = tempwide * 0.4852
  LineBottom.X2 = tempwide * 0.9836
  LineBottom.Y1 = temphigh * 0.7757
  LineBottom.Y2 = temphigh * 0.7757
  
  LineRight.X1 = tempwide * 0.977
  LineRight.X2 = tempwide * 0.977
  LineRight.Y1 = temphigh * 0.0561
  LineRight.Y2 = temphigh * 0.785
  
  linBox2Left.X1 = tempwide * 0.0262
  linBox2Left.X2 = tempwide * 0.0262
  linBox2Left.Y1 = temphigh * 0.0935
  linBox2Left.Y2 = temphigh * 0.9439
  
  linBox2Right.X1 = tempwide * 0.4393
  linBox2Right.X2 = tempwide * 0.4393
  linBox2Right.Y1 = temphigh * 0.0935
  linBox2Right.Y2 = temphigh * 0.9439
  
  linBox2Top.X1 = tempwide * 0.0197
  linBox2Top.X2 = tempwide * 0.4459
  linBox2Top.Y1 = temphigh * 0.1028
  linBox2Top.Y2 = temphigh * 0.1028
  
  linBox2UpperMiddle.X1 = tempwide * 0.0328
  linBox2UpperMiddle.X2 = tempwide * 0.4328
  linBox2UpperMiddle.Y1 = temphigh * 0.5421
  linBox2UpperMiddle.Y2 = temphigh * 0.5421
  
  linBox2Middle.X1 = tempwide * 0.0328
  linBox2Middle.X2 = tempwide * 0.4328
  linBox2Middle.Y1 = temphigh * 0.6542
  linBox2Middle.Y2 = temphigh * 0.6542
  
  linBox2Bottom.X1 = tempwide * 0.0197
  linBox2Bottom.X2 = tempwide * 0.4459
  linBox2Bottom.Y1 = temphigh * 0.9346
  linBox2Bottom.Y2 = temphigh * 0.9346
    
  labCommodityHeading.Top = temphigh * 0.0187
  labCommodityHeading.Left = tempwide * 0.0197
  
  For X = 0 To 4
    txtOreGrades(X).Top = (X * 0.0467 * temphigh) + (temphigh * 0.1612)
    txtOreGrades(X).Left = tempwide * 0.6623
    txtOreGrades(X).Width = tempwide * 0.1328
    labOreGrades(X).Top = (X * 0.0467 * temphigh) + (temphigh * 0.1682)
    labOreGrades(X).Left = tempwide * 0.5115
    labOreGrades(X).Width = tempwide * 0.1295
    labOreUnits(X).Top = (X * 0.0467 * temphigh) + (temphigh * 0.1682)
    labOreUnits(X).Left = tempwide * 0.8
  
    txtWallGrades(X).Top = (X * 0.0467 * temphigh) + (temphigh * 0.507)
    txtWallGrades(X).Left = tempwide * 0.6623
    txtWallGrades(X).Width = tempwide * 0.1328
    labWallGrades(X).Top = (X * 0.0467 * temphigh) + (temphigh * 0.514)
    labWallGrades(X).Left = tempwide * 0.5115
    labWallGrades(X).Width = tempwide * 0.1295
    labWallUnits(X).Top = (X * 0.0467 * temphigh) + (temphigh * 0.514)
    labWallUnits(X).Left = tempwide * 0.8
  
    labCommodityTitles(X).Top = (X * 0.0467 * temphigh) + (temphigh * 0.2653)
    labCommodityTitles(X).Left = tempwide * 0.0459
    labCommodityTitles(X).Width = tempwide * 0.159
    txtCommodInfo(X).Top = (X * 0.0467 * temphigh) + (temphigh * 0.2583)
    txtCommodInfo(X).Left = tempwide * 0.2164
    txtCommodInfo(X).Width = tempwide * 0.1525
    
    If X = 4 Then
      labCommodityTitles(X).Top = temphigh * 0.4727
      txtCommodInfo(X).Top = temphigh * 0.4639
    End If
  
    labCommodityNumber(X).Top = temphigh * 0.1515
    labCommodityNumber(X).Left = (X * 0.0328 * tempwide) + (tempwide * 0.2164)
    labCommodityNumber(X).Width = tempwide * 0.0213
  Next X

  For X = 0 To 3
    cmdCommodList(X).Left = tempwide * (0.0525 + (X * 0.0918))
    cmdCommodList(X).Top = temphigh * 0.5814
    cmdCommodList(X).Width = tempwide * 0.0869
  Next X
  
  For X = 0 To 9
    If X < 5 Then
      Y = 0
    Else
      Y = 0.1121
    End If
    labCheckTag(X).Top = (X * 0.0467 * temphigh) + (temphigh * 0.1682) + (temphigh * Y)
    labCheckTag(X).Left = tempwide * 0.9115
    labCheckTag(X).Width = tempwide * 0.0475
  Next X
  
  labCommodityTitles(5).Top = temphigh * 0.1141
  labCommodityTitles(5).Left = tempwide * 0.1968
  labCommodityTitles(5).Width = tempwide * 0.1918
  
  labCommodityTitles(6).Top = temphigh * 0.8879
  labCommodityTitles(6).Left = tempwide * 0.6098
  labCommodityTitles(6).Width = tempwide * 0.1131
  
  labCommodityTitles(7).Top = temphigh * 0.1121
  labCommodityTitles(7).Left = tempwide * 0.6754
  labCommodityTitles(7).Width = tempwide * 0.1066
  
  labCommodityTitles(8).Top = temphigh * 0.4579
  labCommodityTitles(8).Left = tempwide * 0.6492
  labCommodityTitles(8).Width = tempwide * 0.159
  
  labCommodityTitles(9).Top = temphigh * 0.0561
  labCommodityTitles(9).Left = tempwide * 0.4787

  labCommodityTitles(10).Top = temphigh * 0.1215
  labCommodityTitles(10).Left = tempwide * 0.9115
  labCommodityTitles(10).Width = tempwide * 0.0475
  
  labCommodityTitles(11).Top = temphigh * 0.0935
  labCommodityTitles(11).Left = tempwide * 0.0131
  
  labCommodityTitles(12).Top = temphigh * 0.6449
  labCommodityTitles(12).Left = tempwide * 0.0131
  
  labCommodityTitles(13).Top = temphigh * 0.8879
  labCommodityTitles(13).Left = tempwide * 0.741
  labCommodityTitles(13).Width = tempwide * 0.1131
  
  labCommodityTitles(14).Top = temphigh * 0.5327
  labCommodityTitles(14).Left = tempwide * 0.0131
  
  lstCommodType.Top = temphigh * 0.6749
  lstCommodType.Height = temphigh * 0.2383
  lstCommodType.Left = tempwide * 0.1967
  lstCommodType.Width = tempwide * 0.1656
  
  lstCommodUnit.Top = temphigh * 0.6749
  lstCommodUnit.Height = temphigh * 0.2383
  lstCommodUnit.Left = tempwide * 0.2033
  lstCommodUnit.Width = tempwide * 0.1656
  
  lstCommodPrice.Top = temphigh * 0.6749
  lstCommodPrice.Height = temphigh * 0.2383
  lstCommodPrice.Left = tempwide * 0.2098
  lstCommodPrice.Width = tempwide * 0.1656
  
  lstCommodDepletion.Top = temphigh * 0.6749
  lstCommodDepletion.Height = temphigh * 0.2383
  lstCommodDepletion.Left = tempwide * 0.2164
  lstCommodDepletion.Width = tempwide * 0.1656
  
  lstReservesUnit.Top = temphigh * 0.6749
  lstReservesUnit.Height = temphigh * 0.2383
  lstReservesUnit.Left = tempwide * 0.223
  lstReservesUnit.Width = tempwide * 0.1656
  
  hscCommodityNumber.Top = temphigh * 0.1983
  hscCommodityNumber.Left = tempwide * 0.1902
  hscCommodityNumber.Width = tempwide * 0.2049
    
  hscSetNumber.Top = temphigh * 0.9439
  hscSetNumber.Left = tempwide * 0.636

  labSetNumber.Top = temphigh * 0.9392
  labSetNumber.Left = tempwide * 0.6819
  labSetNumber.Width = tempwide * 0.0279
  
  txtComSetLabel.Top = temphigh * 0.9299
  txtComSetLabel.Left = tempwide * 0.741
  txtComSetLabel.Width = tempwide * 0.1131
  
  comIndTag.Top = temphigh * 0.796
  comIndTag.Left = tempwide * 0.5049
  
  labIndTag.Top = temphigh * 0.7944
  labIndTag.Left = tempwide * 0.5377
  
  comDepTag.Top = temphigh * 0.796
  comDepTag.Left = tempwide * 0.7934
  
  labDepTag.Top = temphigh * 0.7944
  labDepTag.Left = tempwide * 0.8262
  
  labBackToMenu.Top = temphigh * 0.9532
  labBackToMenu.Left = tempwide * 0.0721

  imgBackToMenu.Top = temphigh * 0.9626
  imgBackToMenu.Left = tempwide * 0.0066
  imgBackToMenu.Width = tempwide * 0.0541

  labGradeHelp.Top = temphigh * 0.9532
  labGradeHelp.Left = tempwide * 0.9377

  labInsert.Top = temphigh * 0.9562
  labInsert.Left = tempwide * 0.4467
  labInsert.Width = tempwide * 0.1066

End Sub

Public Sub drawthevalues()
  
  Dim X As Integer
  
  DoNotChange = True

  For X = 0 To 4
    If X > 0 Then txtOreGrades(X).Visible = False
    txtOreGrades(X).Text = ""
    labOreGrades(X).Caption = ""
    labOreUnits(X).Caption = ""
    If X > 0 Then txtWallGrades(X).Visible = False
    txtWallGrades(X).Text = ""
    labWallGrades(X).Caption = ""
    labWallUnits(X).Caption = ""
  Next X
  
  For X = 0 To 4
    If CommodityData(1, X).Name <> "" Then
      txtOreGrades(X).Visible = True
      txtOreGrades(X).Text = Format(LTrim(Str(Primary(hscSetNumber.Value, X + 2))), "###0.0000")
      labOreGrades(X).Caption = RTrim(CommodityData(1, X).Name)
      labOreUnits(X).Caption = CommodityData(1, X).Units
      txtWallGrades(X).Visible = True
      txtWallGrades(X).Text = Format(LTrim(Str(Primary(hscSetNumber.Value, X + 8))), "###0.0000")
      labWallGrades(X).Caption = RTrim(CommodityData(1, X).Name)
      labWallUnits(X).Caption = CommodityData(1, X).Units
    End If
  Next X
    
  For X = 2 To 6
    labCheckTag(X - 2).Visible = False
    If Tagged(hscSetNumber.Value, X).Independent > 0 Then
      labCheckTag(X - 2).Visible = True
      labCheckTag(X - 2).ForeColor = &HFF&
      labCheckTag(X - 2).Caption = LTrim(RTrim(Str(Tagged(hscSetNumber.Value, X).Independent)))
    ElseIf Tagged(hscSetNumber.Value, X).Dependent > 0 Then
      labCheckTag(X - 2).Visible = True
      labCheckTag(X - 2).ForeColor = &HFFFF&
      labCheckTag(X - 2).Caption = LTrim(RTrim(Str(Tagged(hscSetNumber.Value, X).Dependent)))
    Else
      labCheckTag(X - 2).Caption = ""
    End If
    labCheckTag(X + 3).Visible = False
    If Tagged(hscSetNumber.Value, X + 6).Independent > 0 Then
      labCheckTag(X + 3).Visible = True
      labCheckTag(X + 3).ForeColor = &HFF&
      labCheckTag(X + 3).Caption = LTrim(RTrim(Str(Tagged(hscSetNumber.Value, X + 6).Independent)))
    ElseIf Tagged(hscSetNumber.Value, X + 6).Dependent > 0 Then
      labCheckTag(X + 3).Visible = True
      labCheckTag(X + 3).ForeColor = &HFFFF&
      labCheckTag(X + 3).Caption = LTrim(RTrim(Str(Tagged(hscSetNumber.Value, X + 6).Dependent)))
    Else
      labCheckTag(X + 3).Caption = ""
    End If
  Next X
    
  DoNotChange = False

End Sub

Private Sub txtWallGrades_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

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
        If InStr(txtWallGrades(Index).Text, ".") = 0 Then
          SendKeys "{DELETE}", False
        End If
      Else
        SendKeys "{DELETE}", False
      End If
  End Select
End If

End Sub


Private Sub txtWallGrades_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii = 46 Then
  If InStr(txtWallGrades(Index).Text, ".") > 0 Then
    Beep
    KeyAscii = 0
  End If
End If

If KeyAscii = 44 Then
  Beep
  KeyAscii = 0
End If

End Sub

Private Sub txtWallGrades_LostFocus(Index As Integer)
labInsert.Visible = False
End Sub
