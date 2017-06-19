VERSION 5.00
Begin VB.Form frmCashFlow 
   BackColor       =   &H00000000&
   Caption         =   "Schedules"
   ClientHeight    =   8040
   ClientLeft      =   -270
   ClientTop       =   1275
   ClientWidth     =   11880
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   -1  'True
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8200
   ScaleMode       =   0  'User
   ScaleWidth      =   11880
   Begin VB.HScrollBar hscSetNumber 
      Height          =   195
      Left            =   5700
      Max             =   25
      Min             =   1
      TabIndex        =   239
      TabStop         =   0   'False
      Top             =   7740
      Value           =   1
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.VScrollBar vscCashFlow 
      Height          =   6675
      Left            =   11520
      Max             =   28
      TabIndex        =   31
      Top             =   600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.HScrollBar hscCashFlow 
      Height          =   195
      Left            =   2520
      Max             =   18
      Min             =   1
      TabIndex        =   30
      Top             =   7380
      Value           =   1
      Width           =   8895
   End
   Begin VB.Label labNextSet 
      BackColor       =   &H00000000&
      Caption         =   "Next Set"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   6300
      TabIndex        =   238
      Top             =   7740
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label labPrevSet 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Previous Set"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   4380
      TabIndex        =   237
      Top             =   7740
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Line linHoriz 
      BorderColor     =   &H00FFFF00&
      Index           =   13
      X1              =   2520
      X2              =   11400
      Y1              =   7465.671
      Y2              =   7465.671
   End
   Begin VB.Label labFlowTitles 
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
      Left            =   180
      TabIndex        =   236
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label labFlowHeading 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Index           =   1
      Left            =   11655
      TabIndex        =   235
      Top             =   0
      Width           =   45
   End
   Begin VB.Label labYear 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "7"
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
      Height          =   195
      Index           =   6
      Left            =   10140
      TabIndex        =   234
      Top             =   300
      Width           =   1245
   End
   Begin VB.Label labYear 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "6"
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
      Height          =   195
      Index           =   5
      Left            =   8880
      TabIndex        =   233
      Top             =   300
      Width           =   1245
   End
   Begin VB.Label labYear 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "5"
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
      Height          =   195
      Index           =   4
      Left            =   7620
      TabIndex        =   232
      Top             =   300
      Width           =   1245
   End
   Begin VB.Label labYear 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "4"
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
      Height          =   195
      Index           =   3
      Left            =   6360
      TabIndex        =   231
      Top             =   300
      Width           =   1245
   End
   Begin VB.Label labYear 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "3"
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
      Height          =   195
      Index           =   2
      Left            =   5100
      TabIndex        =   230
      Top             =   300
      Width           =   1245
   End
   Begin VB.Label labYear 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "2"
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
      Height          =   195
      Index           =   1
      Left            =   3840
      TabIndex        =   229
      Top             =   300
      Width           =   1245
   End
   Begin VB.Label labYear 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "1"
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
      Height          =   195
      Index           =   0
      Left            =   2580
      TabIndex        =   228
      Top             =   300
      Width           =   1245
   End
   Begin VB.Label labSevenFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   27
      Left            =   10140
      TabIndex        =   227
      Top             =   7080
      Width           =   1245
   End
   Begin VB.Label labSevenFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   26
      Left            =   10140
      TabIndex        =   226
      Top             =   6840
      Width           =   1245
   End
   Begin VB.Label labSevenFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   25
      Left            =   10140
      TabIndex        =   225
      Top             =   6600
      Width           =   1245
   End
   Begin VB.Label labSevenFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   24
      Left            =   10140
      TabIndex        =   224
      Top             =   6360
      Width           =   1245
   End
   Begin VB.Label labSevenFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   23
      Left            =   10140
      TabIndex        =   223
      Top             =   6120
      Width           =   1245
   End
   Begin VB.Label labSevenFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   22
      Left            =   10140
      TabIndex        =   222
      Top             =   5880
      Width           =   1245
   End
   Begin VB.Label labSevenFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   21
      Left            =   10140
      TabIndex        =   221
      Top             =   5640
      Width           =   1245
   End
   Begin VB.Label labSevenFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   20
      Left            =   10140
      TabIndex        =   220
      Top             =   5400
      Width           =   1245
   End
   Begin VB.Label labSevenFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   19
      Left            =   10140
      TabIndex        =   219
      Top             =   5160
      Width           =   1245
   End
   Begin VB.Label labSevenFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   18
      Left            =   10140
      TabIndex        =   218
      Top             =   4920
      Width           =   1245
   End
   Begin VB.Label labSevenFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   17
      Left            =   10140
      TabIndex        =   217
      Top             =   4680
      Width           =   1245
   End
   Begin VB.Label labSevenFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   16
      Left            =   10140
      TabIndex        =   216
      Top             =   4440
      Width           =   1245
   End
   Begin VB.Label labSevenFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   15
      Left            =   10140
      TabIndex        =   215
      Top             =   4200
      Width           =   1245
   End
   Begin VB.Label labSevenFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   14
      Left            =   10140
      TabIndex        =   214
      Top             =   3960
      Width           =   1245
   End
   Begin VB.Label labSevenFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   13
      Left            =   10140
      TabIndex        =   213
      Top             =   3720
      Width           =   1245
   End
   Begin VB.Label labSevenFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   12
      Left            =   10140
      TabIndex        =   212
      Top             =   3480
      Width           =   1245
   End
   Begin VB.Label labSevenFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   11
      Left            =   10140
      TabIndex        =   211
      Top             =   3240
      Width           =   1245
   End
   Begin VB.Label labSevenFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   10140
      TabIndex        =   210
      Top             =   3000
      Width           =   1245
   End
   Begin VB.Label labSevenFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   10140
      TabIndex        =   209
      Top             =   2760
      Width           =   1245
   End
   Begin VB.Label labSevenFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   10140
      TabIndex        =   208
      Top             =   2520
      Width           =   1245
   End
   Begin VB.Label labSevenFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   10140
      TabIndex        =   207
      Top             =   2280
      Width           =   1245
   End
   Begin VB.Label labSevenFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   10140
      TabIndex        =   206
      Top             =   2040
      Width           =   1245
   End
   Begin VB.Label labSevenFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   10140
      TabIndex        =   205
      Top             =   1800
      Width           =   1250
   End
   Begin VB.Label labSevenFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   10140
      TabIndex        =   204
      Top             =   1560
      Width           =   1245
   End
   Begin VB.Label labSevenFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   10140
      TabIndex        =   203
      Top             =   1320
      Width           =   1245
   End
   Begin VB.Label labSevenFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   10140
      TabIndex        =   202
      Top             =   1080
      Width           =   1245
   End
   Begin VB.Label labSevenFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   10140
      TabIndex        =   201
      Top             =   840
      Width           =   1245
   End
   Begin VB.Label labSixFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   27
      Left            =   8880
      TabIndex        =   200
      Top             =   7080
      Width           =   1245
   End
   Begin VB.Label labSixFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   26
      Left            =   8880
      TabIndex        =   199
      Top             =   6840
      Width           =   1245
   End
   Begin VB.Label labSixFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   25
      Left            =   8880
      TabIndex        =   198
      Top             =   6600
      Width           =   1245
   End
   Begin VB.Label labSixFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   24
      Left            =   8880
      TabIndex        =   197
      Top             =   6360
      Width           =   1245
   End
   Begin VB.Label labSixFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   23
      Left            =   8880
      TabIndex        =   196
      Top             =   6120
      Width           =   1245
   End
   Begin VB.Label labSixFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   22
      Left            =   8880
      TabIndex        =   195
      Top             =   5880
      Width           =   1245
   End
   Begin VB.Label labSixFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   21
      Left            =   8880
      TabIndex        =   194
      Top             =   5640
      Width           =   1245
   End
   Begin VB.Label labSixFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   20
      Left            =   8880
      TabIndex        =   193
      Top             =   5400
      Width           =   1245
   End
   Begin VB.Label labSixFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   19
      Left            =   8880
      TabIndex        =   192
      Top             =   5160
      Width           =   1245
   End
   Begin VB.Label labSixFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   18
      Left            =   8880
      TabIndex        =   191
      Top             =   4920
      Width           =   1245
   End
   Begin VB.Label labSixFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   17
      Left            =   8880
      TabIndex        =   190
      Top             =   4680
      Width           =   1245
   End
   Begin VB.Label labSixFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   16
      Left            =   8880
      TabIndex        =   189
      Top             =   4440
      Width           =   1245
   End
   Begin VB.Label labSixFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   15
      Left            =   8880
      TabIndex        =   188
      Top             =   4200
      Width           =   1245
   End
   Begin VB.Label labSixFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   14
      Left            =   8880
      TabIndex        =   187
      Top             =   3960
      Width           =   1245
   End
   Begin VB.Label labSixFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   13
      Left            =   8880
      TabIndex        =   186
      Top             =   3720
      Width           =   1245
   End
   Begin VB.Label labSixFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   12
      Left            =   8880
      TabIndex        =   185
      Top             =   3480
      Width           =   1245
   End
   Begin VB.Label labSixFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   11
      Left            =   8880
      TabIndex        =   184
      Top             =   3240
      Width           =   1245
   End
   Begin VB.Label labSixFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   8880
      TabIndex        =   183
      Top             =   3000
      Width           =   1245
   End
   Begin VB.Label labSixFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   8880
      TabIndex        =   182
      Top             =   2760
      Width           =   1245
   End
   Begin VB.Label labSixFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   8880
      TabIndex        =   181
      Top             =   2520
      Width           =   1245
   End
   Begin VB.Label labSixFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   8880
      TabIndex        =   180
      Top             =   2280
      Width           =   1245
   End
   Begin VB.Label labSixFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   8880
      TabIndex        =   179
      Top             =   2040
      Width           =   1245
   End
   Begin VB.Label labSixFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   8880
      TabIndex        =   178
      Top             =   1800
      Width           =   1245
   End
   Begin VB.Label labSixFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   8880
      TabIndex        =   177
      Top             =   1560
      Width           =   1245
   End
   Begin VB.Label labSixFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   8880
      TabIndex        =   176
      Top             =   1320
      Width           =   1245
   End
   Begin VB.Label labSixFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   8880
      TabIndex        =   175
      Top             =   1080
      Width           =   1245
   End
   Begin VB.Label labSixFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   8880
      TabIndex        =   174
      Top             =   840
      Width           =   1245
   End
   Begin VB.Label labFiveFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   27
      Left            =   7620
      TabIndex        =   173
      Top             =   7080
      Width           =   1245
   End
   Begin VB.Label labFiveFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   26
      Left            =   7620
      TabIndex        =   172
      Top             =   6840
      Width           =   1245
   End
   Begin VB.Label labFiveFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   25
      Left            =   7620
      TabIndex        =   171
      Top             =   6600
      Width           =   1245
   End
   Begin VB.Label labFiveFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   24
      Left            =   7620
      TabIndex        =   170
      Top             =   6360
      Width           =   1245
   End
   Begin VB.Label labFiveFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   23
      Left            =   7620
      TabIndex        =   169
      Top             =   6120
      Width           =   1245
   End
   Begin VB.Label labFiveFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   22
      Left            =   7620
      TabIndex        =   168
      Top             =   5880
      Width           =   1245
   End
   Begin VB.Label labFiveFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   21
      Left            =   7620
      TabIndex        =   167
      Top             =   5640
      Width           =   1245
   End
   Begin VB.Label labFiveFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   20
      Left            =   7620
      TabIndex        =   166
      Top             =   5400
      Width           =   1245
   End
   Begin VB.Label labFiveFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   19
      Left            =   7620
      TabIndex        =   165
      Top             =   5160
      Width           =   1245
   End
   Begin VB.Label labFiveFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   18
      Left            =   7620
      TabIndex        =   164
      Top             =   4920
      Width           =   1245
   End
   Begin VB.Label labFiveFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   17
      Left            =   7620
      TabIndex        =   163
      Top             =   4680
      Width           =   1245
   End
   Begin VB.Label labFiveFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   16
      Left            =   7620
      TabIndex        =   162
      Top             =   4440
      Width           =   1245
   End
   Begin VB.Label labFiveFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   15
      Left            =   7620
      TabIndex        =   161
      Top             =   4200
      Width           =   1245
   End
   Begin VB.Label labFiveFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   14
      Left            =   7620
      TabIndex        =   160
      Top             =   3960
      Width           =   1245
   End
   Begin VB.Label labFiveFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   13
      Left            =   7620
      TabIndex        =   159
      Top             =   3720
      Width           =   1245
   End
   Begin VB.Label labFiveFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   12
      Left            =   7620
      TabIndex        =   158
      Top             =   3480
      Width           =   1245
   End
   Begin VB.Label labFiveFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   11
      Left            =   7620
      TabIndex        =   157
      Top             =   3240
      Width           =   1245
   End
   Begin VB.Label labFiveFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   7620
      TabIndex        =   156
      Top             =   3000
      Width           =   1245
   End
   Begin VB.Label labFiveFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   7620
      TabIndex        =   155
      Top             =   2760
      Width           =   1245
   End
   Begin VB.Label labFiveFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      TabIndex        =   154
      Top             =   2520
      Width           =   1245
   End
   Begin VB.Label labFiveFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   7620
      TabIndex        =   153
      Top             =   2280
      Width           =   1245
   End
   Begin VB.Label labFiveFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   7620
      TabIndex        =   152
      Top             =   2040
      Width           =   1245
   End
   Begin VB.Label labFiveFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   7620
      TabIndex        =   151
      Top             =   1800
      Width           =   1245
   End
   Begin VB.Label labFiveFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   7620
      TabIndex        =   150
      Top             =   1560
      Width           =   1245
   End
   Begin VB.Label labFiveFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   7620
      TabIndex        =   149
      Top             =   1320
      Width           =   1245
   End
   Begin VB.Label labFiveFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   7620
      TabIndex        =   148
      Top             =   1080
      Width           =   1245
   End
   Begin VB.Label labFiveFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   7620
      TabIndex        =   147
      Top             =   840
      Width           =   1245
   End
   Begin VB.Label labFourFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   27
      Left            =   6360
      TabIndex        =   146
      Top             =   7080
      Width           =   1245
   End
   Begin VB.Label labFourFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   26
      Left            =   6360
      TabIndex        =   145
      Top             =   6840
      Width           =   1245
   End
   Begin VB.Label labFourFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   25
      Left            =   6360
      TabIndex        =   144
      Top             =   6600
      Width           =   1245
   End
   Begin VB.Label labFourFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   24
      Left            =   6360
      TabIndex        =   143
      Top             =   6360
      Width           =   1245
   End
   Begin VB.Label labFourFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   23
      Left            =   6360
      TabIndex        =   142
      Top             =   6120
      Width           =   1245
   End
   Begin VB.Label labFourFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   22
      Left            =   6360
      TabIndex        =   141
      Top             =   5880
      Width           =   1245
   End
   Begin VB.Label labFourFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   21
      Left            =   6360
      TabIndex        =   140
      Top             =   5640
      Width           =   1245
   End
   Begin VB.Label labFourFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   20
      Left            =   6360
      TabIndex        =   139
      Top             =   5400
      Width           =   1245
   End
   Begin VB.Label labFourFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   19
      Left            =   6360
      TabIndex        =   138
      Top             =   5160
      Width           =   1245
   End
   Begin VB.Label labFourFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   18
      Left            =   6360
      TabIndex        =   137
      Top             =   4920
      Width           =   1245
   End
   Begin VB.Label labFourFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   17
      Left            =   6360
      TabIndex        =   136
      Top             =   4680
      Width           =   1245
   End
   Begin VB.Label labFourFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   16
      Left            =   6360
      TabIndex        =   135
      Top             =   4440
      Width           =   1245
   End
   Begin VB.Label labFourFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   15
      Left            =   6360
      TabIndex        =   134
      Top             =   4200
      Width           =   1245
   End
   Begin VB.Label labFourFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   14
      Left            =   6360
      TabIndex        =   133
      Top             =   3960
      Width           =   1245
   End
   Begin VB.Label labFourFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   13
      Left            =   6360
      TabIndex        =   132
      Top             =   3720
      Width           =   1245
   End
   Begin VB.Label labFourFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   12
      Left            =   6360
      TabIndex        =   131
      Top             =   3480
      Width           =   1245
   End
   Begin VB.Label labFourFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   11
      Left            =   6360
      TabIndex        =   130
      Top             =   3240
      Width           =   1245
   End
   Begin VB.Label labFourFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   6360
      TabIndex        =   129
      Top             =   3000
      Width           =   1245
   End
   Begin VB.Label labFourFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   6360
      TabIndex        =   128
      Top             =   2760
      Width           =   1245
   End
   Begin VB.Label labFourFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   6360
      TabIndex        =   127
      Top             =   2520
      Width           =   1245
   End
   Begin VB.Label labFourFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   6360
      TabIndex        =   126
      Top             =   2280
      Width           =   1245
   End
   Begin VB.Label labFourFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   6360
      TabIndex        =   125
      Top             =   2040
      Width           =   1245
   End
   Begin VB.Label labFourFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   6360
      TabIndex        =   124
      Top             =   1800
      Width           =   1245
   End
   Begin VB.Label labFourFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   6360
      TabIndex        =   123
      Top             =   1560
      Width           =   1245
   End
   Begin VB.Label labFourFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   6360
      TabIndex        =   122
      Top             =   1320
      Width           =   1245
   End
   Begin VB.Label labFourFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   6360
      TabIndex        =   121
      Top             =   1080
      Width           =   1245
   End
   Begin VB.Label labFourFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   6360
      TabIndex        =   120
      Top             =   840
      Width           =   1245
   End
   Begin VB.Label labThreeFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   27
      Left            =   5100
      TabIndex        =   119
      Top             =   7080
      Width           =   1245
   End
   Begin VB.Label labThreeFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   26
      Left            =   5100
      TabIndex        =   118
      Top             =   6840
      Width           =   1245
   End
   Begin VB.Label labThreeFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   25
      Left            =   5100
      TabIndex        =   117
      Top             =   6600
      Width           =   1245
   End
   Begin VB.Label labThreeFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   24
      Left            =   5100
      TabIndex        =   116
      Top             =   6360
      Width           =   1245
   End
   Begin VB.Label labThreeFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   23
      Left            =   5100
      TabIndex        =   115
      Top             =   6120
      Width           =   1245
   End
   Begin VB.Label labThreeFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   22
      Left            =   5100
      TabIndex        =   114
      Top             =   5880
      Width           =   1245
   End
   Begin VB.Label labThreeFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   21
      Left            =   5100
      TabIndex        =   113
      Top             =   5640
      Width           =   1245
   End
   Begin VB.Label labThreeFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   20
      Left            =   5100
      TabIndex        =   112
      Top             =   5400
      Width           =   1245
   End
   Begin VB.Label labThreeFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   19
      Left            =   5100
      TabIndex        =   111
      Top             =   5160
      Width           =   1245
   End
   Begin VB.Label labThreeFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   18
      Left            =   5100
      TabIndex        =   110
      Top             =   4920
      Width           =   1245
   End
   Begin VB.Label labThreeFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   17
      Left            =   5100
      TabIndex        =   109
      Top             =   4680
      Width           =   1245
   End
   Begin VB.Label labThreeFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   16
      Left            =   5100
      TabIndex        =   108
      Top             =   4440
      Width           =   1245
   End
   Begin VB.Label labThreeFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   15
      Left            =   5100
      TabIndex        =   107
      Top             =   4200
      Width           =   1245
   End
   Begin VB.Label labThreeFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   14
      Left            =   5100
      TabIndex        =   106
      Top             =   3960
      Width           =   1245
   End
   Begin VB.Label labThreeFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   13
      Left            =   5100
      TabIndex        =   105
      Top             =   3720
      Width           =   1245
   End
   Begin VB.Label labThreeFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   12
      Left            =   5100
      TabIndex        =   104
      Top             =   3480
      Width           =   1245
   End
   Begin VB.Label labThreeFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   11
      Left            =   5100
      TabIndex        =   103
      Top             =   3240
      Width           =   1245
   End
   Begin VB.Label labThreeFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   5100
      TabIndex        =   102
      Top             =   3000
      Width           =   1245
   End
   Begin VB.Label labThreeFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   5100
      TabIndex        =   101
      Top             =   2760
      Width           =   1245
   End
   Begin VB.Label labThreeFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   5100
      TabIndex        =   100
      Top             =   2520
      Width           =   1245
   End
   Begin VB.Label labThreeFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   5100
      TabIndex        =   99
      Top             =   2280
      Width           =   1245
   End
   Begin VB.Label labThreeFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   5100
      TabIndex        =   98
      Top             =   2040
      Width           =   1245
   End
   Begin VB.Label labThreeFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   5100
      TabIndex        =   97
      Top             =   1800
      Width           =   1245
   End
   Begin VB.Label labThreeFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   5100
      TabIndex        =   96
      Top             =   1560
      Width           =   1245
   End
   Begin VB.Label labThreeFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   5100
      TabIndex        =   95
      Top             =   1320
      Width           =   1245
   End
   Begin VB.Label labThreeFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   5100
      TabIndex        =   94
      Top             =   1080
      Width           =   1245
   End
   Begin VB.Label labThreeFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   5100
      TabIndex        =   93
      Top             =   840
      Width           =   1245
   End
   Begin VB.Label labSevenFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   10140
      TabIndex        =   92
      Top             =   600
      Width           =   1240
   End
   Begin VB.Label labSixFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   8880
      TabIndex        =   91
      Top             =   600
      Width           =   1245
   End
   Begin VB.Label labFiveFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   7620
      TabIndex        =   90
      Top             =   600
      Width           =   1245
   End
   Begin VB.Label labFourFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   6360
      TabIndex        =   89
      Top             =   600
      Width           =   1245
   End
   Begin VB.Label labThreeFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   5100
      TabIndex        =   88
      Top             =   600
      Width           =   1245
   End
   Begin VB.Label labTwoFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   27
      Left            =   3840
      TabIndex        =   87
      Top             =   7080
      Width           =   1245
   End
   Begin VB.Label labTwoFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   26
      Left            =   3840
      TabIndex        =   86
      Top             =   6840
      Width           =   1245
   End
   Begin VB.Label labTwoFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   25
      Left            =   3840
      TabIndex        =   85
      Top             =   6600
      Width           =   1245
   End
   Begin VB.Label labTwoFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   24
      Left            =   3840
      TabIndex        =   84
      Top             =   6360
      Width           =   1245
   End
   Begin VB.Label labTwoFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   23
      Left            =   3840
      TabIndex        =   83
      Top             =   6120
      Width           =   1245
   End
   Begin VB.Label labTwoFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   22
      Left            =   3840
      TabIndex        =   82
      Top             =   5880
      Width           =   1245
   End
   Begin VB.Label labTwoFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   21
      Left            =   3840
      TabIndex        =   81
      Top             =   5640
      Width           =   1245
   End
   Begin VB.Label labTwoFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   20
      Left            =   3840
      TabIndex        =   80
      Top             =   5400
      Width           =   1245
   End
   Begin VB.Label labTwoFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   19
      Left            =   3840
      TabIndex        =   79
      Top             =   5160
      Width           =   1245
   End
   Begin VB.Label labTwoFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   18
      Left            =   3840
      TabIndex        =   78
      Top             =   4920
      Width           =   1245
   End
   Begin VB.Label labTwoFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   17
      Left            =   3840
      TabIndex        =   77
      Top             =   4680
      Width           =   1245
   End
   Begin VB.Label labTwoFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   16
      Left            =   3840
      TabIndex        =   76
      Top             =   4440
      Width           =   1245
   End
   Begin VB.Label labTwoFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   15
      Left            =   3840
      TabIndex        =   75
      Top             =   4200
      Width           =   1245
   End
   Begin VB.Label labTwoFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   14
      Left            =   3840
      TabIndex        =   74
      Top             =   3960
      Width           =   1245
   End
   Begin VB.Label labTwoFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   13
      Left            =   3840
      TabIndex        =   73
      Top             =   3720
      Width           =   1245
   End
   Begin VB.Label labTwoFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   12
      Left            =   3840
      TabIndex        =   72
      Top             =   3480
      Width           =   1245
   End
   Begin VB.Label labTwoFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   11
      Left            =   3840
      TabIndex        =   71
      Top             =   3240
      Width           =   1245
   End
   Begin VB.Label labTwoFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   3840
      TabIndex        =   70
      Top             =   3000
      Width           =   1245
   End
   Begin VB.Label labTwoFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   3840
      TabIndex        =   69
      Top             =   2760
      Width           =   1245
   End
   Begin VB.Label labTwoFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   3840
      TabIndex        =   68
      Top             =   2520
      Width           =   1245
   End
   Begin VB.Label labTwoFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   3840
      TabIndex        =   67
      Top             =   2280
      Width           =   1245
   End
   Begin VB.Label labTwoFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   3840
      TabIndex        =   66
      Top             =   2040
      Width           =   1245
   End
   Begin VB.Label labTwoFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   3840
      TabIndex        =   65
      Top             =   1800
      Width           =   1245
   End
   Begin VB.Label labTwoFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   3840
      TabIndex        =   64
      Top             =   1560
      Width           =   1245
   End
   Begin VB.Label labTwoFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   3840
      TabIndex        =   63
      Top             =   1320
      Width           =   1245
   End
   Begin VB.Label labTwoFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   3840
      TabIndex        =   62
      Top             =   1080
      Width           =   1245
   End
   Begin VB.Label labTwoFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   3840
      TabIndex        =   61
      Top             =   840
      Width           =   1245
   End
   Begin VB.Label labTwoFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   3840
      TabIndex        =   60
      Top             =   600
      Width           =   1245
   End
   Begin VB.Label labOneFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   27
      Left            =   2580
      TabIndex        =   59
      Top             =   7080
      Width           =   1245
   End
   Begin VB.Label labOneFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   26
      Left            =   2580
      TabIndex        =   58
      Top             =   6840
      Width           =   1245
   End
   Begin VB.Label labOneFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   25
      Left            =   2580
      TabIndex        =   57
      Top             =   6600
      Width           =   1245
   End
   Begin VB.Label labOneFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   24
      Left            =   2580
      TabIndex        =   56
      Top             =   6360
      Width           =   1245
   End
   Begin VB.Label labOneFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   23
      Left            =   2580
      TabIndex        =   55
      Top             =   6120
      Width           =   1245
   End
   Begin VB.Label labOneFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   22
      Left            =   2580
      TabIndex        =   54
      Top             =   5880
      Width           =   1245
   End
   Begin VB.Label labOneFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   21
      Left            =   2580
      TabIndex        =   53
      Top             =   5640
      Width           =   1245
   End
   Begin VB.Label labOneFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   20
      Left            =   2580
      TabIndex        =   52
      Top             =   5400
      Width           =   1245
   End
   Begin VB.Label labOneFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   19
      Left            =   2580
      TabIndex        =   51
      Top             =   5160
      Width           =   1245
   End
   Begin VB.Label labOneFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   18
      Left            =   2580
      TabIndex        =   50
      Top             =   4920
      Width           =   1245
   End
   Begin VB.Label labOneFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   17
      Left            =   2580
      TabIndex        =   49
      Top             =   4680
      Width           =   1245
   End
   Begin VB.Label labOneFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   16
      Left            =   2580
      TabIndex        =   48
      Top             =   4440
      Width           =   1245
   End
   Begin VB.Label labOneFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   15
      Left            =   2580
      TabIndex        =   47
      Top             =   4200
      Width           =   1245
   End
   Begin VB.Label labOneFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   14
      Left            =   2580
      TabIndex        =   46
      Top             =   3960
      Width           =   1245
   End
   Begin VB.Label labOneFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   13
      Left            =   2580
      TabIndex        =   45
      Top             =   3720
      Width           =   1245
   End
   Begin VB.Label labOneFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   12
      Left            =   2580
      TabIndex        =   44
      Top             =   3480
      Width           =   1245
   End
   Begin VB.Label labOneFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   11
      Left            =   2580
      TabIndex        =   43
      Top             =   3240
      Width           =   1245
   End
   Begin VB.Label labOneFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   2580
      TabIndex        =   42
      Top             =   3000
      Width           =   1245
   End
   Begin VB.Label labOneFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   2580
      TabIndex        =   41
      Top             =   2760
      Width           =   1245
   End
   Begin VB.Label labOneFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   2580
      TabIndex        =   40
      Top             =   2520
      Width           =   1245
   End
   Begin VB.Label labOneFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   2580
      TabIndex        =   39
      Top             =   2280
      Width           =   1245
   End
   Begin VB.Label labOneFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   2580
      TabIndex        =   38
      Top             =   2040
      Width           =   1245
   End
   Begin VB.Label labOneFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   2580
      TabIndex        =   37
      Top             =   1800
      Width           =   1245
   End
   Begin VB.Label labOneFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   2580
      TabIndex        =   36
      Top             =   1560
      Width           =   1245
   End
   Begin VB.Label labOneFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   2580
      TabIndex        =   35
      Top             =   1320
      Width           =   1245
   End
   Begin VB.Label labOneFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   2580
      TabIndex        =   34
      Top             =   1080
      Width           =   1245
   End
   Begin VB.Label labOneFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   2580
      TabIndex        =   33
      Top             =   840
      Width           =   1245
   End
   Begin VB.Label labOneFlow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   2580
      TabIndex        =   32
      Top             =   600
      Width           =   1245
   End
   Begin VB.Line linVert 
      BorderColor     =   &H00FFFF00&
      Index           =   2
      X1              =   11460
      X2              =   11460
      Y1              =   611.94
      Y2              =   7404.478
   End
   Begin VB.Line linHoriz 
      BorderColor     =   &H00FFFF00&
      Index           =   12
      X1              =   2520
      X2              =   11400
      Y1              =   6945.522
      Y2              =   6945.522
   End
   Begin VB.Line linVert 
      BorderColor     =   &H00FFFF00&
      Index           =   1
      X1              =   2460
      X2              =   2460
      Y1              =   611.94
      Y2              =   7404.478
   End
   Begin VB.Line linVert 
      BorderColor     =   &H00FFFF00&
      Index           =   0
      X1              =   2460
      X2              =   2460
      Y1              =   305.97
      Y2              =   489.552
   End
   Begin VB.Line linHoriz 
      BorderColor     =   &H00FFFF00&
      Index           =   11
      X1              =   2520
      X2              =   11400
      Y1              =   6455.97
      Y2              =   6455.97
   End
   Begin VB.Line linHoriz 
      BorderColor     =   &H00FFFF00&
      Index           =   10
      X1              =   2520
      X2              =   11400
      Y1              =   4497.761
      Y2              =   4497.761
   End
   Begin VB.Line linHoriz 
      BorderColor     =   &H00FFFF00&
      Index           =   9
      X1              =   2520
      X2              =   11400
      Y1              =   4008.209
      Y2              =   4008.209
   End
   Begin VB.Line linHoriz 
      BorderColor     =   &H00FFFF00&
      Index           =   8
      X1              =   2520
      X2              =   11400
      Y1              =   3029.104
      Y2              =   3029.104
   End
   Begin VB.Line linHoriz 
      BorderColor     =   &H00FFFF00&
      Index           =   7
      X1              =   2520
      X2              =   11400
      Y1              =   550.746
      Y2              =   550.746
   End
   Begin VB.Line linHoriz 
      BorderColor     =   &H00FFFF00&
      Index           =   6
      X1              =   180
      X2              =   2400
      Y1              =   7465.671
      Y2              =   7465.671
   End
   Begin VB.Line linHoriz 
      BorderColor     =   &H00FFFF00&
      Index           =   5
      X1              =   180
      X2              =   2400
      Y1              =   6945.522
      Y2              =   6945.522
   End
   Begin VB.Line linHoriz 
      BorderColor     =   &H00FFFF00&
      Index           =   4
      X1              =   180
      X2              =   2400
      Y1              =   6455.97
      Y2              =   6455.97
   End
   Begin VB.Line linHoriz 
      BorderColor     =   &H00FFFF00&
      Index           =   3
      X1              =   180
      X2              =   2400
      Y1              =   4497.761
      Y2              =   4497.761
   End
   Begin VB.Line linHoriz 
      BorderColor     =   &H00FFFF00&
      Index           =   2
      X1              =   180
      X2              =   2400
      Y1              =   4008.209
      Y2              =   4008.209
   End
   Begin VB.Line linHoriz 
      BorderColor     =   &H00FFFF00&
      Index           =   1
      X1              =   180
      X2              =   2400
      Y1              =   3029.104
      Y2              =   3029.104
   End
   Begin VB.Line linHoriz 
      BorderColor     =   &H00FFFF00&
      Index           =   0
      X1              =   180
      X2              =   2400
      Y1              =   550.746
      Y2              =   550.746
   End
   Begin VB.Label labFlowTitles 
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
      Index           =   27
      Left            =   180
      TabIndex        =   29
      Top             =   7080
      Width           =   2175
   End
   Begin VB.Label labFlowTitles 
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
      Index           =   26
      Left            =   180
      TabIndex        =   28
      Top             =   6840
      Width           =   2175
   End
   Begin VB.Label labFlowTitles 
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
      Index           =   25
      Left            =   180
      TabIndex        =   27
      Top             =   6600
      Width           =   2175
   End
   Begin VB.Label labFlowTitles 
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
      Index           =   24
      Left            =   180
      TabIndex        =   26
      Top             =   6360
      Width           =   2175
   End
   Begin VB.Label labFlowTitles 
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
      Index           =   23
      Left            =   180
      TabIndex        =   25
      Top             =   6120
      Width           =   2175
   End
   Begin VB.Label labFlowTitles 
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
      Index           =   22
      Left            =   180
      TabIndex        =   24
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Label labFlowTitles 
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
      Index           =   21
      Left            =   180
      TabIndex        =   23
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Label labFlowTitles 
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
      Index           =   20
      Left            =   180
      TabIndex        =   22
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Label labFlowTitles 
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
      Index           =   19
      Left            =   180
      TabIndex        =   21
      Top             =   5160
      Width           =   2175
   End
   Begin VB.Label labFlowTitles 
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
      Index           =   18
      Left            =   180
      TabIndex        =   20
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label labFlowTitles 
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
      Index           =   17
      Left            =   180
      TabIndex        =   19
      Top             =   4680
      Width           =   2175
   End
   Begin VB.Label labFlowTitles 
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
      Index           =   16
      Left            =   180
      TabIndex        =   18
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label labFlowTitles 
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
      Index           =   15
      Left            =   180
      TabIndex        =   17
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label labFlowTitles 
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
      Index           =   14
      Left            =   180
      TabIndex        =   16
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label labFlowTitles 
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
      Index           =   13
      Left            =   180
      TabIndex        =   15
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label labFlowTitles 
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
      Index           =   12
      Left            =   180
      TabIndex        =   14
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label labFlowTitles 
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
      Index           =   11
      Left            =   180
      TabIndex        =   13
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label labFlowTitles 
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
      Index           =   10
      Left            =   180
      TabIndex        =   12
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label labFlowTitles 
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
      Index           =   9
      Left            =   180
      TabIndex        =   11
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label labFlowTitles 
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
      Index           =   8
      Left            =   180
      TabIndex        =   10
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label labFlowTitles 
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
      Index           =   7
      Left            =   180
      TabIndex        =   9
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label labFlowTitles 
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
      Index           =   6
      Left            =   180
      TabIndex        =   8
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label labFlowTitles 
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
      Index           =   5
      Left            =   180
      TabIndex        =   7
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label labFlowTitles 
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
      Index           =   4
      Left            =   180
      TabIndex        =   6
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label labFlowTitles 
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
      Left            =   180
      TabIndex        =   5
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label labFlowTitles 
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
      Left            =   180
      TabIndex        =   4
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label labFlowTitles 
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
      Index           =   0
      Left            =   180
      TabIndex        =   3
      Top             =   600
      Width           =   2175
   End
   Begin VB.Image imgBackToMenu 
      Height          =   150
      Left            =   60
      Picture         =   "frmCashFlow.frx":0000
      Stretch         =   -1  'True
      Top             =   7740
      Width           =   390
   End
   Begin VB.Label labPrintScreen 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   11160
      TabIndex        =   2
      Top             =   7680
      Width           =   555
   End
   Begin VB.Label labBackToMenu 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Top             =   7680
      Width           =   615
   End
   Begin VB.Label labFlowHeading 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Cash Flow"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   915
   End
   Begin VB.Line linBottomSpread 
      BorderColor     =   &H00FFFF00&
      X1              =   60
      X2              =   11820
      Y1              =   7771.642
      Y2              =   7771.642
   End
   Begin VB.Line linRightSpread 
      BorderColor     =   &H00FFFF00&
      X1              =   11760
      X2              =   11760
      Y1              =   183.582
      Y2              =   7832.835
   End
   Begin VB.Line linLeftSpread 
      BorderColor     =   &H00FFFF00&
      X1              =   120
      X2              =   120
      Y1              =   183.582
      Y2              =   7832.835
   End
   Begin VB.Line linTopSpread 
      BorderColor     =   &H00FFFF00&
      X1              =   60
      X2              =   11820
      Y1              =   244.776
      Y2              =   244.776
   End
End
Attribute VB_Name = "frmCashFlow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim howwide As Integer
Dim howhigh As Integer
Dim temprow As Integer
Dim temptit(27) As String
Dim tempdeptit(50) As String

Private Sub Form_Activate()

Dim i As Integer
Dim j As Integer
Dim block As Integer
Dim propna As Integer
Dim orena As Integer
Dim maxyr1 As Integer
Dim endyr As Integer

ShowMenu = True

DoNotChange = True
'If FullScreen = False Then
  frmCashFlow.WindowState = 0
'Else
'  frmCashFlow.WindowState = 2
'End If
DoNotChange = False

frmCashFlow.Top = 0
frmCashFlow.Left = 0

For i = 0 To 27
  labFlowTitles(i).Visible = True
  labOneFlow(i).Visible = True
  labTwoFlow(i).Visible = True
  labThreeFlow(i).Visible = True
  labFourFlow(i).Visible = True
  labFiveFlow(i).Visible = True
  labSixFlow(i).Visible = True
  labSevenFlow(i).Visible = True
  labFlowTitles(i).Caption = ""
  labOneFlow(i).Caption = ""
  labTwoFlow(i).Caption = ""
  labThreeFlow(i).Caption = ""
  labFourFlow(i).Caption = ""
  labFiveFlow(i).Caption = ""
  labSixFlow(i).Caption = ""
  labSevenFlow(i).Caption = ""
Next i

For i = 0 To 6
  linHoriz(i).Visible = True
  linHoriz(i + 7).Visible = True
Next i

linVert(1).Visible = True
linVert(2).Visible = True

vscCashFlow.Visible = True
hscCashFlow.Visible = True

DoNotChange = True
vscCashFlow.Value = 0
hscCashFlow.Value = 1
DoNotChange = False

block = 1

If job = 2 Then propna = block
If job > 2 Then
  propna = Primary(block, 28)
  orena = Primary(block, 27)
End If

If job = 4 Then
  For i = 1 To 50
    For j = 1 To 15
      Ore(i, j) = 0
    Next j
  Next i
End If

If job < 4 Then
  Call cflow5(job, block)
ElseIf job = 5 Then
  Call cflow5(1, block)
End If

If job = 4 Then
 For block = 1 To Np(4)
    If Primary(block, 28) > 0 Then
      Call cflow5(job, block)
    End If
  Next block
  For i = 1 To MaxYr
    For j = 2 To 15
      If Ore(i, 1) <> 0 Then
        Ore(i, j) = Ore(i, j) / Ore(i, 1)
      End If
    Next j
  Next i
  block = 1
End If

maxyr1 = MaxYr
If MaxYr = 50 Then maxyr1 = 49
endyr = maxyr1 + 1

For j = 1 To 28
  Secondary(endyr, j) = 0
  For i = 1 To maxyr1
    Secondary(endyr, j) = Secondary(endyr, j) + Secondary(i, j)
  Next i
Next j

For j = 1 To 13
  Prop(endyr, j) = 0
  For i = 1 To maxyr1
    Prop(endyr, j) = Prop(endyr, j) + Prop(i, j)
  Next i
Next j

If job = 2 Then
  If Np(8) > 1 Then
    labPrevSet.Visible = True
    labNextSet.Visible = True
    hscSetNumber.Visible = True
  End If
ElseIf job = 3 Then
  If Np(4) > 1 Then
    labPrevSet.Visible = True
    labNextSet.Visible = True
    hscSetNumber.Visible = True
  End If
Else
  labPrevSet.Visible = False
  labNextSet.Visible = False
  hscSetNumber.Visible = False
End If

If job = 1 Then
  Call drawcash
ElseIf job = 2 Then
  Call drawroyalty
ElseIf job = 5 Then
  Call drawdepreciate
Else
  Call drawvalue
End If
Call flowspread(1, 0)

End Sub
Private Sub Form_Deactivate()
      
  If ShowMenu = True Then
    frmCashFlow.Hide
    Call InputMenuAccess(3)
  End If
  
End Sub

Private Sub Form_Load()

Dim i As Integer
Dim X As Integer
Dim Y As Integer
Dim temphigh As Single
Dim tempright As Integer
Dim tempwide As Single

X = 0
While Y = 0
  Select Case X
    Case 0
      tempright = labOneFlow(0).Left + labOneFlow(0).Width
    Case 1
      tempright = labTwoFlow(0).Left + labTwoFlow(0).Width
    Case 2
      tempright = labThreeFlow(0).Left + labThreeFlow(0).Width
    Case 3
      tempright = labFourFlow(0).Left + labFourFlow(0).Width
    Case 4
      tempright = labFiveFlow(0).Left + labFiveFlow(0).Width
    Case 5
      tempright = labSixFlow(0).Left + labSixFlow(0).Width
    Case 6
      tempright = labSevenFlow(0).Left + labSevenFlow(0).Width
    Case Else
      tempright = frmCashFlow.ScaleWidth
  End Select
  If tempright >= frmCashFlow.ScaleWidth - 460 Then
    howwide = X + 1
    Y = 1
  Else
    X = X + 1
  End If
Wend

frmCashFlow.Height = Screen.Height * (frmCashFlow.ScaleHeight / frmCashFlow.Height)
If Screen.Width < 12000 Then
  frmCashFlow.Width = Screen.Width
Else
  frmCashFlow.Width = 12000
End If
    
frmCashFlow.Top = 0
frmCashFlow.Left = 0

temphigh = frmCashFlow.ScaleHeight
tempwide = frmCashFlow.ScaleWidth
  
linBottomSpread.X1 = tempwide * 0.0066
linBottomSpread.X2 = tempwide * 0.9902
linBottomSpread.Y1 = temphigh * 0.9689
linBottomSpread.Y2 = temphigh * 0.9689
  
linRightSpread.X1 = tempwide * 0.9827
linRightSpread.X2 = tempwide * 0.9827

End Sub
Public Sub flowspread(startrow, startcol)

Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim l As Integer
Dim m As Integer
Dim addit As Integer
Dim endrow As Integer
Dim endcol As Integer
Dim whichset As Integer

On Error Resume Next

If job = 1 Then
  howhigh = temprow
  endcol = startcol + howhigh
ElseIf job = 2 Then
  howhigh = 13
  endcol = startcol + howhigh
ElseIf job = 5 Then
  howhigh = temprow
  If NumCap = 0 Then howhigh = 4
  endcol = startcol + howhigh
Else
  If whichset = 0 Then whichset = 1
  ValueCostUpper = 0
  While CommodityData(whichset, ValueCostUpper).Name <> ""
    ValueCostUpper = ValueCostUpper + 1
  Wend
  howhigh = 15 - ((9 - ValueCostUpper) - SecondUpper)
  endcol = startcol + howhigh
End If

endrow = startrow + 6
If job = 3 Or job = 4 Then
  If endrow > MaxYr Then endrow = MaxYr
Else
  If endrow > MaxYr + 1 Then endrow = MaxYr + 1
End If
endcol = startcol + 27
If endcol > 27 Then endcol = 27

For i = 0 To 6
  labYear(i).Caption = ""
Next i

For i = startrow To endrow
  l = 0
  TempDepTot(i) = 0
  labYear(i - startrow).Caption = LTrim(RTrim(Str(i + Sets(12) - 1)))
  If i = MaxYr + 1 Then
    labYear(i - startrow).Caption = "Cumulative"
  End If
  For j = 0 To howhigh
    If job = 1 Then
      labFlowTitles(j).Caption = temptit(j + startcol)
      For k = 1 To 8 Step 7
        linHoriz(k).Y1 = 2970 - (startcol * 240)
        linHoriz(k).Y2 = 2970 - (startcol * 240)
        If linHoriz(k).Y1 > frmCashFlow.ScaleHeight - 720 Then
           linHoriz(k).Visible = False
        End If
        linHoriz(k + 1).Y1 = 3930 - (startcol * 240)
        linHoriz(k + 1).Y2 = 3930 - (startcol * 240)
        If linHoriz(k + 1).Y1 > frmCashFlow.ScaleHeight - 720 Then
           linHoriz(k + 1).Visible = False
        End If
        linHoriz(k + 2).Y1 = 4410 - (startcol * 240)
        linHoriz(k + 2).Y2 = 4410 - (startcol * 240)
        If linHoriz(k + 2).Y1 > frmCashFlow.ScaleHeight - 720 Then
           linHoriz(k + 2).Visible = False
        End If
        linHoriz(k + 3).Y1 = 6330 - (startcol * 240)
        linHoriz(k + 3).Y2 = 6330 - (startcol * 240)
        If linHoriz(k + 3).Y1 > frmCashFlow.ScaleHeight - 720 Then
           linHoriz(k + 3).Visible = False
        Else
           linHoriz(k + 3).Visible = True
        End If
        linHoriz(k + 4).Y1 = 6810 - (startcol * 240)
        linHoriz(k + 4).Y2 = 6810 - (startcol * 240)
        If linHoriz(k + 4).Y1 > frmCashFlow.ScaleHeight - 720 Then
           linHoriz(k + 4).Visible = False
        Else
           linHoriz(k + 4).Visible = True
        End If
      Next k
      If i = startrow Then
        If j < 9 - startcol Then
          labOneFlow(j).Caption = Format(Str(Secondary(i, j + startcol + 1)), "$##,###,###,###")
        Else
          labOneFlow(j).Caption = Format(Str(Secondary(i, j + startcol + 2)), "$##,###,###,###")
        End If
      ElseIf i = startrow + 1 Then
        If j < 9 - startcol Then
          labTwoFlow(j).Caption = Format(Str(Secondary(i, j + startcol + 1)), "$##,###,###,###")
        Else
          labTwoFlow(j).Caption = Format(Str(Secondary(i, j + startcol + 2)), "$#,###,###,###")
        End If
      ElseIf i = startrow + 2 Then
        If j < 9 - startcol Then
          labThreeFlow(j).Caption = Format(Str(Secondary(i, j + startcol + 1)), "$#,###,###,###")
        Else
          labThreeFlow(j).Caption = Format(Str(Secondary(i, j + startcol + 2)), "$#,###,###,###")
        End If
      ElseIf i = startrow + 3 Then
        If j < 9 - startcol Then
          labFourFlow(j).Caption = Format(Str(Secondary(i, j + startcol + 1)), "$#,###,###,###")
        Else
          labFourFlow(j).Caption = Format(Str(Secondary(i, j + startcol + 2)), "$#,###,###,###")
        End If
      ElseIf i = startrow + 4 Then
        If j < 9 - startcol Then
          labFiveFlow(j).Caption = Format(Str(Secondary(i, j + startcol + 1)), "$#,###,###,###")
        Else
          labFiveFlow(j).Caption = Format(Str(Secondary(i, j + startcol + 2)), "$##,###,###,###")
        End If
      ElseIf i = startrow + 5 Then
        If j < 9 - startcol Then
          labSixFlow(j).Caption = Format(Str(Secondary(i, j + startcol + 1)), "$##,###,###,###")
        Else
          labSixFlow(j).Caption = Format(Str(Secondary(i, j + startcol + 2)), "$##,###,###,###")
        End If
      ElseIf i = startrow + 6 Then
        If j < 9 - startcol Then
          labSevenFlow(j).Caption = Format(Str(Secondary(i, j + startcol + 1)), "$##,###,###,###")
        Else
          labSevenFlow(j).Caption = Format(Str(Secondary(i, j + startcol + 2)), "$##,###,###,###")
        End If
      End If
    ElseIf job = 2 Then
      If i = startrow Then
        labOneFlow(j).Caption = Format(Str(Prop(i, j + 1)), "$##,###,###,###")
      ElseIf i = startrow + 1 Then
        labTwoFlow(j).Caption = Format(Str(Prop(i, j + 1)), "$##,###,###,###")
      ElseIf i = startrow + 2 Then
        labThreeFlow(j).Caption = Format(Str(Prop(i, j + 1)), "$##,###,###,###")
      ElseIf i = startrow + 3 Then
        labFourFlow(j).Caption = Format(Str(Prop(i, j + 1)), "$##,###,###,###")
      ElseIf i = startrow + 4 Then
        labFiveFlow(j).Caption = Format(Str(Prop(i, j + 1)), "$##,###,###,###")
      ElseIf i = startrow + 5 Then
        labSixFlow(j).Caption = Format(Str(Prop(i, j + 1)), "$##,###,###,###")
      ElseIf i = startrow + 6 Then
        labSevenFlow(j).Caption = Format(Str(Prop(i, j + 1)), "$##,###,###,###")
      End If
    ElseIf job = 5 Then
      If Deprecy(j + l, MaxYr + 1) > 0 Or j = NumCap Then
        If howhigh > 0 And NumCap > 0 And j <= NumCap Then
          While j < howhigh And Deprecy(j + l, MaxYr + 1) = 0
            If Deprecy(j + 1, MaxYr + 1) = 0 Then l = l + 1
          Wend
        End If
        If j < howhigh And Deprecy(j + l, MaxYr + 1) > 0 Then
          labFlowTitles(j).Caption = tempdeptit(j + startcol)
        End If
        TempDepTot(i) = 0
        For m = 0 To NumCap - 1
          TempDepTot(i) = TempDepTot(i) + Deprecy(m, i)
        Next m
        If i = startrow Then
          If j < howhigh And NumCap > 0 Then
            labOneFlow(j).Caption = Format(Str(Deprecy(j + l + startcol, i)), "$##,###,###,###")
          Else
            labOneFlow(j).Caption = Format(Str(TempDepTot(i)), "$##,###,###,###")
          End If
        ElseIf i = startrow + 1 Then
          If j < howhigh And NumCap > 0 Then
            labTwoFlow(j).Caption = Format(Str(Deprecy(j + l + startcol, i)), "$##,###,###,###")
          Else
            labTwoFlow(j).Caption = Format(Str(TempDepTot(i)), "$##,###,###,###")
          End If
        ElseIf i = startrow + 2 Then
          If j < howhigh And NumCap > 0 Then
            labThreeFlow(j).Caption = Format(Str(Deprecy(j + l + startcol, i)), "$##,###,###,###")
          Else
            labThreeFlow(j).Caption = Format(Str(TempDepTot(i)), "$##,###,###,###")
          End If
        ElseIf i = startrow + 3 Then
          If j < howhigh And NumCap > 0 Then
            labFourFlow(j).Caption = Format(Str(Deprecy(j + l + startcol, i)), "$##,###,###,###")
          Else
            labFourFlow(j).Caption = Format(Str(TempDepTot(i)), "$##,###,###,###")
          End If
        ElseIf i = startrow + 4 Then
          If j < howhigh And NumCap > 0 Then
            labFiveFlow(j).Caption = Format(Str(Deprecy(j + l + startcol, i)), "$##,###,###,###")
          Else
            labFiveFlow(j).Caption = Format(Str(TempDepTot(i)), "$##,###,###,###")
          End If
        ElseIf i = startrow + 5 Then
          If j < howhigh And NumCap > 0 Then
            labSixFlow(j).Caption = Format(Str(Deprecy(j + l + startcol, i)), "$##,###,###,###")
          Else
            labSixFlow(j).Caption = Format(Str(TempDepTot(i)), "$##,###,###,###")
          End If
        ElseIf i = startrow + 6 Then
          If j < howhigh And NumCap > 0 Then
            labSevenFlow(j).Caption = Format(Str(Deprecy(j + l + startcol, i)), "$##,###,###,###")
          Else
            labSevenFlow(j).Caption = Format(Str(TempDepTot(i)), "$##,###,###,###")
          End If
        End If
      End If
    Else
      If j >= ValueCostUpper + SecondUpper + 2 Then
        addit = (9 - ValueCostUpper - SecondUpper)
      ElseIf j >= ValueCostUpper + 1 Then
        addit = (6 - ValueCostUpper)
      Else
        addit = 1
      End If
      If i = startrow Then
        If Ore(i, 1) > 0 Then
          If j = 0 Then
            labOneFlow(j).Caption = Format(Str(Ore(i, j + addit)), "##,###,###,###")
          Else
            labOneFlow(j).Caption = Format(Str(Ore(i, j + addit)), "$#,##0.00")
          End If
        Else
          If j = 0 Then
            labOneFlow(j).Caption = Format(0, "#,###,###,###")
          Else
            labOneFlow(j).Caption = Format(0, "$#,##0.00")
          End If
        End If
      ElseIf i = startrow + 1 Then
        If Ore(i, 1) > 0 Then
          If j = 0 Then
            labTwoFlow(j).Caption = Format(Str(Ore(i, j + addit)), "##,###,###,###")
          Else
            labTwoFlow(j).Caption = Format(Str(Ore(i, j + addit)), "$#,##0.00")
          End If
        Else
          If j = 0 Then
            labTwoFlow(j).Caption = Format(0, "##,###,###,###")
          Else
            labTwoFlow(j).Caption = Format(0, "$#,##0.00")
          End If
        End If
      ElseIf i = startrow + 2 Then
        If Ore(i, 1) > 0 Then
          If j = 0 Then
            labThreeFlow(j).Caption = Format(Str(Ore(i, j + addit)), "##,###,###,###")
          Else
            labThreeFlow(j).Caption = Format(Str(Ore(i, j + addit)), "$#,##0.00")
          End If
        Else
          If j = 0 Then
            labThreeFlow(j).Caption = Format(0, "##,###,###,###")
          Else
            labThreeFlow(j).Caption = Format(0, "$#,##0.00")
          End If
        End If
      ElseIf i = startrow + 3 Then
        If Ore(i, 1) > 0 Then
          If j = 0 Then
            labFourFlow(j).Caption = Format(Str(Ore(i, j + addit)), "##,###,###,###")
          Else
            labFourFlow(j).Caption = Format(Str(Ore(i, j + addit)), "$#,##0.00")
          End If
        Else
          If j = 0 Then
            labFourFlow(j).Caption = Format(0, "##,###,###,###")
          Else
            labFourFlow(j).Caption = Format(0, "$#,##0.00")
          End If
        End If
      ElseIf i = startrow + 4 Then
        If Ore(i, 1) > 0 Then
          If j = 0 Then
            labFiveFlow(j).Caption = Format(Str(Ore(i, j + addit)), "##,###,###,###")
          Else
            labFiveFlow(j).Caption = Format(Str(Ore(i, j + addit)), "$#,##0.00")
          End If
        Else
          If j = 0 Then
            labFiveFlow(j).Caption = Format(0, "##,###,###,###")
          Else
            labFiveFlow(j).Caption = Format(0, "$#,##0.00")
          End If
        End If
      ElseIf i = startrow + 5 Then
        If Ore(i, 1) > 0 Then
          If j = 0 Then
            labSixFlow(j).Caption = Format(Str(Ore(i, j + addit)), "##,###,###,###")
          Else
            labSixFlow(j).Caption = Format(Str(Ore(i, j + addit)), "$#,##0.00")
          End If
        Else
          If j = 0 Then
            labSixFlow(j).Caption = Format(0, "##,###,###,###")
          Else
            labSixFlow(j).Caption = Format(0, "$#,##0.00")
          End If
        End If
      ElseIf i = startrow + 6 Then
        If Ore(i, 1) > 0 Then
          If j = 0 Then
            labSevenFlow(j).Caption = Format(Str(Ore(i, j + addit)), "##,###,###,###")
          Else
            labSevenFlow(j).Caption = Format(Str(Ore(i, j + addit)), "$#,##0.00")
          End If
        Else
          If j = 0 Then
            labSevenFlow(j).Caption = Format(0, "##,###,###,###")
          Else
            labSevenFlow(j).Caption = Format(0, "$#,##0.00")
          End If
        End If
      End If
   
   End If
 Next j
 Next i

End Sub

Private Sub HScroll1_Change()

End Sub

Private Sub Form_Resize()

If DoNotChange = True Then Exit Sub

If frmCashFlow.ScaleHeight > 0 And frmCashFlow.ScaleWidth > 0 Then

  Call screenstuff
  
  If job = 2 Then
    If Pn1(8, hscSetNumber.Value) <> "" Then
      labFlowHeading(1).Caption = "Royalty Set Number " & LTrim(RTrim(Str(hscSetNumber.Value))) & " - " & LTrim(RTrim(Pn1(8, hscSetNumber.Value)))
    Else
      labFlowHeading(1).Caption = "Royalty Set Number " & LTrim(RTrim(Str(hscSetNumber.Value)))
    End If
  ElseIf job = 3 Then
    If Pn1(4, hscSetNumber.Value) <> "" Then
      labFlowHeading(1).Caption = "Production Set Number " & LTrim(RTrim(Str(hscSetNumber.Value))) & " - " & LTrim(RTrim(Pn1(4, hscSetNumber.Value)))
    Else
      labFlowHeading(1).Caption = "Production Set Number " & LTrim(RTrim(Str(hscSetNumber.Value)))
    End If
  ElseIf job = 4 Then
    labFlowHeading(1).Caption = "Composite - All Production Sets"
  End If
  labFlowHeading(1).Left = frmCashFlow.ScaleWidth - (TextWidth(labFlowHeading(1).Caption) + 360)
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

  frmCashFlow.Hide
  If ShowMenu = True Then Call InputMenuAccess(3)

End Sub

Private Sub hscCashFlow_Change()

If DoNotChange = True Then Exit Sub

Dim startrow As Integer
Dim startcol As Integer

startrow = hscCashFlow.Value
startcol = vscCashFlow.Value

Call flowspread(startrow, startcol)

End Sub

Private Sub hscSetNumber_Change()

If job = 2 Then
  If Pn1(8, hscSetNumber.Value) <> "" Then
    labFlowHeading(1).Caption = "Royalty Set Number " & LTrim(RTrim(Str(hscSetNumber.Value))) & " - " & LTrim(RTrim(Pn1(8, hscSetNumber.Value)))
  Else
    labFlowHeading(1).Caption = "Royalty Set Number " & LTrim(RTrim(Str(hscSetNumber.Value)))
  End If
  labFlowHeading(1).Left = frmCashFlow.ScaleWidth - (TextWidth(labFlowHeading(1).Caption) + 360)
  If hscSetNumber.Value <= Np(8) Then
    Call cflow5(job, hscSetNumber.Value)
    Call flowspread(1, 0)
  End If
ElseIf job = 3 Then
  If Pn1(4, hscSetNumber.Value) <> "" Then
    labFlowHeading(1).Caption = "Production Set Number " & LTrim(RTrim(Str(hscSetNumber.Value))) & " - " & LTrim(RTrim(Pn1(4, hscSetNumber.Value)))
  Else
    labFlowHeading(1).Caption = "Production Set Number " & LTrim(RTrim(Str(hscSetNumber.Value)))
  End If
  labFlowHeading(1).Left = frmCashFlow.ScaleWidth - (TextWidth(labFlowHeading(1).Caption) + 120)
  If hscSetNumber.Value <= Np(4) Then
    Call cflow5(job, hscSetNumber.Value)
    Call flowspread(1, 0)
  End If
End If

End Sub

Private Sub imgBackToMenu_Click()
  
  temprow = 1
  hscSetNumber.Value = 1
  frmCashFlow.Hide
  Call InputMenuAccess(3)
 
End Sub

Private Sub labBackToMenu_Click()
  
  temprow = 1
  hscSetNumber.Value = 1
  frmCashFlow.Hide
  Call InputMenuAccess(3)

End Sub


Private Sub labCashFlow_Click(Index As Integer)

End Sub

Public Sub screenstuff()

Dim i As Integer
Dim X As Integer
Dim Y As Integer
Dim tempright As Integer
Dim tempcol As Integer
Dim tempbottom As Integer

X = 0
While Y = 0
  Select Case X
    Case 0
      tempright = labOneFlow(0).Left + labOneFlow(0).Width
    Case 1
      tempright = labTwoFlow(0).Left + labTwoFlow(0).Width
    Case 2
      tempright = labThreeFlow(0).Left + labThreeFlow(0).Width
    Case 3
      tempright = labFourFlow(0).Left + labFourFlow(0).Width
    Case 4
      tempright = labFiveFlow(0).Left + labFiveFlow(0).Width
    Case 5
      tempright = labSixFlow(0).Left + labSixFlow(0).Width
    Case 6
      tempright = labSevenFlow(0).Left + labSevenFlow(0).Width
    Case Else
      tempright = frmCashFlow.ScaleWidth
  End Select
  If tempright >= frmCashFlow.ScaleWidth - 460 Then
    tempcol = X
    howwide = tempcol + 1
    Y = 1
  Else
    X = X + 1
  End If
Wend

Y = 0
X = 10
If job = 5 Then
  X = 0
ElseIf job = 3 Or job = 4 Then
  X = ValueCostUpper + 8
ElseIf job = 2 Then
  X = 13
End If

While Y = 0
  If X < 27 Then
    tempbottom = labOneFlow(X + 1).Top + labOneFlow(X + 1).Height + 5
  End If
  If tempbottom >= frmCashFlow.ScaleHeight - 720 Or X = 27 Then
    temprow = X
    Y = 1
  Else
    X = X + 1
  End If
Wend

For i = 0 To 27
  If i <= temprow Then
    labFlowTitles(i).Visible = True
Else
    labFlowTitles(i).Visible = False
  End If
Next i

For X = 0 To 6
  If X < tempcol Then
    labYear(X).Visible = True
  Else
    labYear(X).Visible = False
  End If
Next X

If tempcol >= 1 Then
  For i = 0 To 27
    If i <= temprow Then
      labOneFlow(i).Visible = True
    Else
      labOneFlow(i).Visible = False
    End If
  Next i
Else
  For i = 0 To 27
    labOneFlow(i).Visible = False
  Next i
End If
      
If tempcol >= 2 Then
  For i = 0 To 27
    If i <= temprow Then
      labTwoFlow(i).Visible = True
    Else
      labTwoFlow(i).Visible = False
    End If
  Next i
Else
  For i = 0 To 27
    labTwoFlow(i).Visible = False
  Next i
End If

If tempcol >= 3 Then
  For i = 0 To 27
    If i <= temprow Then
      labThreeFlow(i).Visible = True
    Else
      labThreeFlow(i).Visible = False
    End If
  Next i
Else
  For i = 0 To 27
    labThreeFlow(i).Visible = False
  Next i
End If

If tempcol >= 4 Then
  For i = 0 To 27
    If i <= temprow Then
      labFourFlow(i).Visible = True
    Else
      labFourFlow(i).Visible = False
    End If
  Next i
Else
  For i = 0 To 27
    labFourFlow(i).Visible = False
  Next i
End If

If tempcol >= 5 Then
  For i = 0 To 27
    If i <= temprow Then
      labFiveFlow(i).Visible = True
    Else
      labFiveFlow(i).Visible = False
    End If
  Next i
Else
  For i = 0 To 27
    labFiveFlow(i).Visible = False
  Next i
End If

If tempcol >= 6 Then
  For i = 0 To 27
    If i <= temprow Then
      labSixFlow(i).Visible = True
    Else
      labSixFlow(i).Visible = False
    End If
  Next i
Else
  For i = 0 To 27
    labSixFlow(i).Visible = False
  Next i
End If

If tempcol >= 7 Then
  For i = 0 To 27
    If i <= temprow Then
      labSevenFlow(i).Visible = True
    Else
      labSevenFlow(i).Visible = False
    End If
  Next i
Else
  For i = 0 To 27
    labSevenFlow(i).Visible = False
  Next i
End If

If MaxYr > (tempcol - 1) Then
  If job = 3 Or job = 4 Then
    hscCashFlow.max = (MaxYr - tempcol) + 1
  Else
    hscCashFlow.max = (MaxYr - tempcol) + 2
  End If
ElseIf MaxYr = 2 Then
  hscCashFlow.max = 25
Else
  hscCashFlow.Visible = False
End If

If hscCashFlow.max = 1 Then hscCashFlow.Visible = False

vscCashFlow.max = 27 - temprow

If job = 5 Then
  If NumCap - 1 > 27 Then
    vscCashFlow.max = NumCap - temprow
  Else
    vscCashFlow.max = 0
  End If
End If

If job = 1 Then
  If temprow < 27 Then
    vscCashFlow.Visible = True
  Else
    vscCashFlow.Visible = False
  End If
ElseIf job = 5 Then
  If temprow < NumCap - 1 Then
    vscCashFlow.Visible = True
  Else
    vscCashFlow.Visible = False
  End If
Else
  vscCashFlow.Visible = False
End If

hscCashFlow.Top = frmCashFlow.ScaleHeight - 660
vscCashFlow.Left = frmCashFlow.ScaleWidth - 360

For i = 7 To 13
  linHoriz(i).X2 = frmCashFlow.ScaleWidth - 480
Next i

For i = 1 To 8 Step 7
  linHoriz(i).Visible = True
  linHoriz(i + 1).Visible = True
  linHoriz(i + 2).Visible = True
  If linHoriz(i + 3).Y1 > frmCashFlow.ScaleHeight - 720 Then
    linHoriz(i + 3).Visible = False
  Else
    linHoriz(i + 3).Visible = True
  End If
  If linHoriz(i + 4).Y1 > frmCashFlow.ScaleHeight - 720 Then
    linHoriz(i + 4).Visible = False
  Else
    linHoriz(i + 4).Visible = True
  End If
Next i

If job = 2 Then
  linHoriz(5).Visible = False
  linHoriz(12).Visible = False
ElseIf job = 5 Then
  linHoriz(2).Visible = False
  linHoriz(9).Visible = False
  linHoriz(3).Visible = False
  linHoriz(10).Visible = False
  linHoriz(4).Visible = False
  linHoriz(11).Visible = False
  linHoriz(5).Visible = False
  linHoriz(12).Visible = False
  If linHoriz(1).Y1 > frmCashFlow.ScaleHeight - 780 Then
    linHoriz(1).Visible = False
    linHoriz(8).Visible = False
  Else
    linHoriz(1).Visible = True
    linHoriz(8).Visible = True
  End If
End If

linHoriz(6).Y1 = frmCashFlow.ScaleHeight - 720
linHoriz(6).Y2 = frmCashFlow.ScaleHeight - 720
linHoriz(13).Y1 = frmCashFlow.ScaleHeight - 720
linHoriz(13).Y2 = frmCashFlow.ScaleHeight - 720

linVert(1).Y2 = frmCashFlow.ScaleHeight - 780
linVert(2).X1 = frmCashFlow.ScaleWidth - 420
linVert(2).X2 = frmCashFlow.ScaleWidth - 420
linVert(2).Y2 = frmCashFlow.ScaleHeight - 780

linRightSpread.X1 = frmCashFlow.ScaleWidth - 120
linRightSpread.X2 = frmCashFlow.ScaleWidth - 120
linLeftSpread.Y2 = frmCashFlow.ScaleHeight - 360
linRightSpread.Y2 = frmCashFlow.ScaleHeight - 360

linBottomSpread.Y1 = frmCashFlow.ScaleHeight - 420
linBottomSpread.Y2 = frmCashFlow.ScaleHeight - 420
linBottomSpread.X2 = frmCashFlow.ScaleWidth - 60

linTopSpread.X2 = frmCashFlow.ScaleWidth - 60

vscCashFlow.Height = (frmCashFlow.ScaleHeight - 780) - vscCashFlow.Top
hscCashFlow.Width = (frmCashFlow.ScaleWidth - 480) - hscCashFlow.Left

labPrintScreen.Top = frmCashFlow.ScaleHeight - 350
labPrintScreen.Left = frmCashFlow.ScaleWidth - 750

labBackToMenu.Top = frmCashFlow.ScaleHeight - 350
labBackToMenu.Left = frmCashFlow.ScaleWidth * 0.0512

imgBackToMenu.Top = frmCashFlow.ScaleHeight - 290
imgBackToMenu.Left = frmCashFlow.ScaleWidth * 0.0044
imgBackToMenu.Width = frmCashFlow.ScaleWidth * 0.0422

labPrevSet.Top = frmCashFlow.ScaleHeight - 330
labPrevSet.Left = (frmCashFlow.ScaleWidth / 2) - 1570

labNextSet.Top = frmCashFlow.ScaleHeight - 330
labNextSet.Left = (frmCashFlow.ScaleWidth / 2) + 370

hscSetNumber.Top = frmCashFlow.ScaleHeight - 300
hscSetNumber.Left = (frmCashFlow.ScaleWidth / 2) - 220

End Sub

Private Sub labPrintScreen_Click()

ShowMenu = False
Call printstuffout(job)

End Sub

Private Sub vscCashFlow_Change()

If DoNotChange = True Then Exit Sub

Dim startcol As Integer
Dim startrow As Integer

startcol = vscCashFlow.Value
startrow = hscCashFlow.Value

Call flowspread(startrow, startcol)

End Sub

Public Sub drawcash()

Dim i As Integer
Dim j As Integer

If frmCashFlow.ScaleHeight = 0 Or frmCashFlow.ScaleWidth = 0 Then Exit Sub
  
labFlowHeading(0).Caption = "Cash Flow Schedule"
labFlowHeading(1).Caption = ""

frmCashFlow.Top = 0
If Screen.Height - 415 < 8425 Then
  frmCashFlow.Height = Screen.Height - 415
Else
  frmCashFlow.Height = 8425
End If

For i = 0 To 13
  labFlowTitles(i).Font.Size = 9
Next i

temptit(0) = "Revenue"
temptit(1) = "Salvage"
temptit(2) = "Royalties"
temptit(3) = "Operating Costs"
temptit(4) = "Expensed Capital"
temptit(5) = "Depreciation"
temptit(6) = "Property Tax"
temptit(7) = "Severance Tax"
temptit(8) = "Interest Expense"
temptit(9) = "Reclamation"
temptit(10) = "Income Before Depletion"
temptit(11) = "Depletion"
temptit(12) = "Loss Carry Forward"
temptit(13) = "State Tax"
temptit(14) = "Net Taxable Income"
temptit(15) = "Federal Tax"
temptit(16) = "Net After Tax"
temptit(17) = "Depreciation"
temptit(18) = "Depletion"
temptit(19) = "Loss Forward"
temptit(20) = "Working Capital"
temptit(21) = "Total Capital"
temptit(22) = "Loan Principal"
temptit(23) = "Joint Venture Capital"
temptit(24) = "Net After Tax Cash Flow"
temptit(25) = "Partner's Share"
temptit(26) = "Net Cash Flow"
temptit(27) = "Cummulative"

For i = 0 To 27
  labFlowTitles(i).Caption = temptit(i)
Next i

For i = 0 To 27
  labFlowTitles(i).Top = 600 + (i * 240)
  labOneFlow(i).Top = 600 + (i * 240)
  labTwoFlow(i).Top = 600 + (i * 240)
  labThreeFlow(i).Top = 600 + (i * 240)
  labFourFlow(i).Top = 600 + (i * 240)
  labFiveFlow(i).Top = 600 + (i * 240)
  labSixFlow(i).Top = 600 + (i * 240)
  labSevenFlow(i).Top = 600 + (i * 240)
Next i

For i = 1 To 8 Step 7
  linHoriz(i).Y1 = 2970
  linHoriz(i).Y2 = 2970
  linHoriz(i + 1).Y1 = 3930
  linHoriz(i + 1).Y2 = 3930
  linHoriz(i + 2).Y1 = 4410
  linHoriz(i + 2).Y2 = 4410
  linHoriz(i + 3).Y1 = 6330
  linHoriz(i + 3).Y2 = 6330
  linHoriz(i + 4).Y1 = 6810
  linHoriz(i + 4).Y2 = 6810
Next i

For i = 1 To 5
  If linHoriz(6).Y1 < linHoriz(i).Y1 Then
    linHoriz(i).Visible = False
    linHoriz(i + 7).Visible = False
  Else
    linHoriz(i).Visible = True
    linHoriz(i + 7).Visible = True
  End If
Next i

Call screenstuff

End Sub

Public Sub drawvalue()

Dim i As Integer
Dim j As Integer
Dim whichset As Integer

If whichset = 0 Then whichset = 1

i = 0
While CommodityData(whichset, i).Name <> ""
  labFlowTitles(i + 1).Caption = LTrim(RTrim(CommodityData(whichset, i).Name)) & " Value"
  i = i + 1
Wend
ValueCostUpper = i

Call screenstuff

If job = 3 Then
  labFlowHeading(0).Caption = "Values and Costs Schedule"
  If Pn1(4, hscSetNumber.Value) <> "" Then
    labFlowHeading(1).Caption = "Production Set Number " & LTrim(RTrim(Str(hscSetNumber.Value))) & " - " & LTrim(RTrim(Pn1(4, hscSetNumber.Value)))
  Else
    labFlowHeading(1).Caption = "Production Set Number " & LTrim(RTrim(Str(hscSetNumber.Value)))
  End If
  hscSetNumber.max = Np(4)
ElseIf job = 4 Then
  labFlowHeading(0).Caption = "Values and Costs Schedule"
  labFlowHeading(1).Caption = "Composite - All Production Sets"
End If

vscCashFlow.Visible = False

labFlowHeading(1).Left = frmCashFlow.ScaleWidth - (TextWidth(labFlowHeading(1).Caption) + 120)

For i = 0 To 13
  labFlowTitles(i).Font.Size = 9
Next i

If whichset = 0 Then whichset = 1

labFlowTitles(0).Caption = "Tons Mined"
i = 0
While CommodityData(whichset, i).Name <> ""
  labFlowTitles(i + 1).Caption = LTrim(RTrim(CommodityData(whichset, i).Name)) & " Value"
  i = i + 1
Wend
ValueCostUpper = i
labFlowTitles(i + 1).Caption = "Gross Value"
i = i + 1
j = 1
While Primary(hscSetNumber.Value, 28 + j) <> 0
  labFlowTitles(i + 1).Caption = "Process " & LTrim(Str(j)) & " Value"
  i = i + 1
  j = j + 1
Wend
SecondUpper = j - 1
labFlowTitles(i + 1).Caption = "Mining Cost"
labFlowTitles(i + 2).Caption = "Milling Cost"
labFlowTitles(i + 3).Caption = "Transport and Smelting"
labFlowTitles(i + 4).Caption = "Refining Cost"
labFlowTitles(i + 5).Caption = "Total Cost"

If frmCashFlow.ScaleHeight = 0 Or frmCashFlow.ScaleWidth = 0 Then Exit Sub

frmCashFlow.Top = ((Screen.Height - 415) - (4200 + ((ValueCostUpper - 1) * 240))) / 2
frmCashFlow.Height = 4200 + ((ValueCostUpper + SecondUpper - 2) * 240)

For i = 1 To 8 Step 7
  linHoriz(i).Y1 = 840
  linHoriz(i).Y2 = 840
  linHoriz(i + 1).Y1 = 1140 + ((ValueCostUpper - 1) * 240)
  linHoriz(i + 1).Y2 = 1140 + ((ValueCostUpper - 1) * 240)
  linHoriz(i + 2).Y1 = 1420 + ((ValueCostUpper - 1) * 240)
  linHoriz(i + 2).Y2 = 1420 + ((ValueCostUpper - 1) * 240)
  linHoriz(i + 3).Y1 = 1740 + ((ValueCostUpper + SecondUpper - 2) * 240)
  linHoriz(i + 3).Y2 = 1740 + ((ValueCostUpper + SecondUpper - 2) * 240)
  linHoriz(i + 4).Y1 = 2760 + ((ValueCostUpper + SecondUpper - 2) * 240)
  linHoriz(i + 4).Y2 = 2760 + ((ValueCostUpper + SecondUpper - 2) * 240)
Next i

For i = 0 To (ValueCostUpper + SecondUpper + 8)
  If i = 0 Then
    j = 0
  ElseIf i = 1 Then
    j = 60
  ElseIf i = ValueCostUpper + 1 Then
    j = 120
  ElseIf i = ValueCostUpper + 2 Then
    j = 180
  ElseIf i = ValueCostUpper + SecondUpper + 2 Then
    j = 240
  ElseIf i = ValueCostUpper + SecondUpper + 6 Then
    j = 300
  End If
  labFlowTitles(i).Top = j + 600 + (i * 240)
  labOneFlow(i).Top = j + 600 + (i * 240)
  labTwoFlow(i).Top = j + 600 + (i * 240)
  labThreeFlow(i).Top = j + 600 + (i * 240)
  labFourFlow(i).Top = j + 600 + (i * 240)
  labFiveFlow(i).Top = j + 600 + (i * 240)
  labSixFlow(i).Top = j + 600 + (i * 240)
  labSevenFlow(i).Top = j + 600 + (i * 240)
Next i

For i = (ValueCostUpper + SecondUpper + 8) To 27
  labFlowTitles(i).Visible = False
  labOneFlow(i).Visible = False
  labTwoFlow(i).Visible = False
  labThreeFlow(i).Visible = False
  labFourFlow(i).Visible = False
  labFiveFlow(i).Visible = False
  labSixFlow(i).Visible = False
  labSevenFlow(i).Visible = False
Next i

For i = 1 To 5
  If linHoriz(6).Y1 < linHoriz(i).Y1 Then
    linHoriz(i).Visible = False
    linHoriz(i + 7).Visible = False
  Else
    linHoriz(i).Visible = True
    linHoriz(i + 7).Visible = True
  End If
Next i

End Sub

Public Sub drawroyalty()
  
Dim i As Integer
Dim j As Integer

Call screenstuff

labFlowHeading(0).Caption = "Royalty Schedule"
If Pn1(8, hscSetNumber.Value) <> "" Then
  labFlowHeading(1).Caption = "Royalty Set Number " & LTrim(RTrim(Str(hscSetNumber.Value))) & " - " & LTrim(RTrim(Pn1(8, hscSetNumber.Value)))
Else
  labFlowHeading(1).Caption = "Royalty Set Number " & LTrim(RTrim(Str(hscSetNumber.Value)))
End If
vscCashFlow.Visible = False

labFlowTitles(0).Caption = "Revenues"
labFlowTitles(1).Caption = "Transport/Smelting/Refining"
labFlowTitles(2).Caption = "Operating Costs"
labFlowTitles(3).Caption = "Depreciation"
labFlowTitles(4).Caption = "Property/Severance Tax"
labFlowTitles(5).Caption = "Net Profits"
labFlowTitles(6).Caption = "Net Profit Royalty Due"
labFlowTitles(7).Caption = "Exploration/Development"
labFlowTitles(8).Caption = "Recapture"
labFlowTitles(9).Caption = "Total NP and NSR Due"
labFlowTitles(10).Caption = "Advance Royalty Payment"
labFlowTitles(11).Caption = "Recapture"
labFlowTitles(12).Caption = "Net Royalty Payment"
labFlowTitles(13).Caption = "Cumulative Payments"

If frmCashFlow.ScaleHeight = 0 Or frmCashFlow.ScaleWidth = 0 Then Exit Sub

frmCashFlow.Height = 5340

labFlowHeading(1).Left = frmCashFlow.ScaleWidth - (TextWidth(labFlowHeading(1).Caption) + 120)

For i = 0 To 13
  labFlowTitles(i).Font.Size = 8
Next i

hscSetNumber.max = Np(8)

frmCashFlow.Top = ((Screen.Height - 415) - frmCashFlow.Height) / 2

For i = 1 To 8 Step 7
  linHoriz(i).Y1 = 1800
  linHoriz(i).Y2 = 1800
  linHoriz(i + 1).Y1 = 2100
  linHoriz(i + 1).Y2 = 2100
  linHoriz(i + 2).Y1 = 2880
  linHoriz(i + 2).Y2 = 2880
  linHoriz(i + 3).Y1 = 3660
  linHoriz(i + 3).Y2 = 3660
Next i

For i = 0 To 14
  If i < 5 Then
    j = 0
  ElseIf i < 6 Then
    j = 60
  ElseIf i < 9 Then
    j = 120
  ElseIf i < 12 Then
    j = 180
  Else
    j = 240
  End If
  labFlowTitles(i).Top = j + 600 + (i * 240)
  labOneFlow(i).Top = j + 600 + (i * 240)
  labTwoFlow(i).Top = j + 600 + (i * 240)
  labThreeFlow(i).Top = j + 600 + (i * 240)
  labFourFlow(i).Top = j + 600 + (i * 240)
  labFiveFlow(i).Top = j + 600 + (i * 240)
  labSixFlow(i).Top = j + 600 + (i * 240)
  labSevenFlow(i).Top = j + 600 + (i * 240)
Next i

For i = 14 To 27
  labFlowTitles(i).Visible = False
  labOneFlow(i).Visible = False
  labTwoFlow(i).Visible = False
  labThreeFlow(i).Visible = False
  labFourFlow(i).Visible = False
  labFiveFlow(i).Visible = False
  labSixFlow(i).Visible = False
  labSevenFlow(i).Visible = False
Next i

For i = 1 To 4
  If linHoriz(6).Y1 < linHoriz(i).Y1 Then
    linHoriz(i).Visible = False
    linHoriz(i + 7).Visible = False
  Else
    linHoriz(i).Visible = True
    linHoriz(i + 7).Visible = True
  End If
Next i

linHoriz(5).Visible = False
linHoriz(12).Visible = False

End Sub

Public Sub drawdepreciate()

Dim i As Integer
Dim j As Integer
Dim l As Integer
Dim whichset As Integer

labFlowHeading(0).Caption = "Capital Deduction Schedule"
labFlowHeading(1).Caption = ""

For i = 0 To 13
  labFlowTitles(i).Font.Size = 9
Next i

If Screen.Height - 415 < 8425 Then
  frmCashFlow.Height = Screen.Height - 415
Else
  frmCashFlow.Height = 8425
End If

j = 0
For i = 0 To NumCap - 1
  tempdeptit(j) = LTrim(RTrim(CapitalData(i).Item))
  j = j + 1
Next i
If j > 27 Then j = temprow
ValueCostUpper = j

For l = 0 To ValueCostUpper
  labFlowTitles(l).Caption = tempdeptit(l)
Next l

labFlowTitles(j).Caption = "Total"

If frmCashFlow.ScaleHeight = 0 Or frmCashFlow.ScaleWidth = 0 Then Exit Sub

frmCashFlow.Top = ((Screen.Height - 415) - (2020 + (ValueCostUpper * 280))) / 2
If frmCashFlow.Top < 0 Then frmCashFlow.Top = 0
frmCashFlow.Height = 2020 + (ValueCostUpper * 240)

For i = 1 To 8 Step 7
  linHoriz(i).Y1 = 840 + ((ValueCostUpper - 1) * 240)
  linHoriz(i).Y2 = 840 + ((ValueCostUpper - 1) * 240)
  linHoriz(i + 1).Visible = False
  linHoriz(i + 2).Visible = False
  linHoriz(i + 3).Visible = False
  linHoriz(i + 4).Visible = False
Next i

For i = 0 To ValueCostUpper
  If i = ValueCostUpper Then
    labFlowTitles(i).Top = 640 + (i * 240)
    labOneFlow(i).Top = 640 + (i * 240)
    labTwoFlow(i).Top = 640 + (i * 240)
    labThreeFlow(i).Top = 640 + (i * 240)
    labFourFlow(i).Top = 640 + (i * 240)
    labFiveFlow(i).Top = 640 + (i * 240)
    labSixFlow(i).Top = 640 + (i * 240)
    labSevenFlow(i).Top = 640 + (i * 240)
  Else
    labFlowTitles(i).Top = 600 + (i * 240)
    labOneFlow(i).Top = 600 + (i * 240)
    labTwoFlow(i).Top = 600 + (i * 240)
    labThreeFlow(i).Top = 600 + (i * 240)
    labFourFlow(i).Top = 600 + (i * 240)
    labFiveFlow(i).Top = 600 + (i * 240)
    labSixFlow(i).Top = 600 + (i * 240)
    labSevenFlow(i).Top = 600 + (i * 240)
  End If
Next i

For i = (ValueCostUpper + 1) To 27
  labFlowTitles(i).Visible = False
  labOneFlow(i).Visible = False
  labTwoFlow(i).Visible = False
  labThreeFlow(i).Visible = False
  labFourFlow(i).Visible = False
  labFiveFlow(i).Visible = False
  labSixFlow(i).Visible = False
  labSevenFlow(i).Visible = False
Next i

Call screenstuff

End Sub
