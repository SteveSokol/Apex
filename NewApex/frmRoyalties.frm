VERSION 5.00
Begin VB.Form frmRoyalties 
   BackColor       =   &H00000000&
   Caption         =   "Royalties"
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
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6420
   ScaleWidth      =   9150
   Begin VB.ListBox lstPayment 
      Height          =   735
      ItemData        =   "frmRoyalties.frx":0000
      Left            =   660
      List            =   "frmRoyalties.frx":000D
      TabIndex        =   64
      Top             =   4020
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.TextBox txtRoySetLabel 
      Height          =   330
      Left            =   960
      TabIndex        =   60
      Top             =   3360
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
      Left            =   6720
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   5700
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
      Left            =   3180
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   5700
      Width           =   195
   End
   Begin VB.HScrollBar hscSetNumbers 
      Height          =   195
      Left            =   1200
      Max             =   25
      Min             =   1
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   2520
      Value           =   1
      Width           =   375
   End
   Begin VB.TextBox txtRoyaltyValues 
      Height          =   330
      Index           =   10
      Left            =   5340
      TabIndex        =   22
      Text            =   "1"
      Top             =   5100
      Width           =   1215
   End
   Begin VB.TextBox txtRoyaltyValues 
      Height          =   330
      Index           =   9
      Left            =   5340
      TabIndex        =   21
      Text            =   "1"
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox txtRoyaltyValues 
      Height          =   330
      Index           =   8
      Left            =   5340
      TabIndex        =   20
      Text            =   "0"
      Top             =   4500
      Width           =   1215
   End
   Begin VB.TextBox txtRoyaltyValues 
      Height          =   330
      Index           =   7
      Left            =   5340
      TabIndex        =   19
      Top             =   4200
      Width           =   1515
   End
   Begin VB.TextBox txtRoyaltyValues 
      Height          =   330
      Index           =   6
      Left            =   5340
      TabIndex        =   18
      Text            =   "0"
      Top             =   3900
      Width           =   1215
   End
   Begin VB.TextBox txtRoyaltyValues 
      Height          =   330
      Index           =   5
      Left            =   5340
      TabIndex        =   17
      Text            =   "0"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtRoyaltyValues 
      Height          =   330
      Index           =   4
      Left            =   5340
      TabIndex        =   16
      Text            =   "1"
      Top             =   2460
      Width           =   1215
   End
   Begin VB.TextBox txtRoyaltyValues 
      Height          =   330
      Index           =   3
      Left            =   5340
      TabIndex        =   15
      Text            =   "0.00"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtRoyaltyValues 
      Height          =   330
      Index           =   2
      Left            =   5340
      TabIndex        =   14
      Text            =   "0.00"
      Top             =   1860
      Width           =   1215
   End
   Begin VB.TextBox txtRoyaltyValues 
      Height          =   330
      Index           =   1
      Left            =   5340
      TabIndex        =   13
      Text            =   "0.00"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtRoyaltyValues 
      Height          =   330
      Index           =   0
      Left            =   5340
      TabIndex        =   12
      Text            =   "0"
      Top             =   1260
      Width           =   1215
   End
   Begin VB.Label labRoyaltiesHelp 
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
      TabIndex        =   63
      Top             =   6120
      Width           =   495
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
      Left            =   5340
      TabIndex        =   62
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label labRoyaltyMisc 
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
      Index           =   8
      Left            =   960
      TabIndex        =   61
      Top             =   3060
      Width           =   1035
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
      Left            =   7920
      TabIndex        =   57
      Top             =   5160
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
      Left            =   7920
      TabIndex        =   56
      Top             =   4860
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
      Left            =   7920
      TabIndex        =   55
      Top             =   4560
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
      Left            =   7920
      TabIndex        =   54
      Top             =   4260
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
      Left            =   7920
      TabIndex        =   53
      Top             =   3960
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
      Left            =   7920
      TabIndex        =   52
      Top             =   3660
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
      Left            =   7920
      TabIndex        =   51
      Top             =   2520
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
      Left            =   7920
      TabIndex        =   50
      Top             =   2220
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
      Left            =   7920
      TabIndex        =   49
      Top             =   1920
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
      Left            =   7920
      TabIndex        =   48
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
      Index           =   0
      Left            =   7920
      TabIndex        =   47
      Top             =   1320
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label labIndTag 
      BackColor       =   &H00000000&
      Caption         =   "Independent Tag"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   3480
      TabIndex        =   46
      Top             =   5700
      Width           =   1455
   End
   Begin VB.Label labDepTag 
      BackColor       =   &H00000000&
      Caption         =   "Dependent Tag"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   7020
      TabIndex        =   45
      Top             =   5700
      Width           =   1335
   End
   Begin VB.Label labRoyaltyMisc 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "$"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   5160
      TabIndex        =   44
      Top             =   4560
      Width           =   135
   End
   Begin VB.Label labRoyaltyMisc 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "$"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   5160
      TabIndex        =   43
      Top             =   3960
      Width           =   135
   End
   Begin VB.Label labRoyaltyMisc 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "$"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   5160
      TabIndex        =   42
      Top             =   3660
      Width           =   135
   End
   Begin VB.Label labRoyaltyMisc 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "$"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   5160
      TabIndex        =   41
      Top             =   1620
      Width           =   135
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
      Left            =   1620
      TabIndex        =   40
      Top             =   2520
      Width           =   315
   End
   Begin VB.Label labRoyaltyMisc 
      BackColor       =   &H00000000&
      Caption         =   "Advanced/Fixed Royalties"
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
      Left            =   2940
      TabIndex        =   38
      Top             =   3120
      Width           =   2355
   End
   Begin VB.Label labRoyaltyMisc 
      BackColor       =   &H00000000&
      Caption         =   "Revenue/Production/Profit Royalties"
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
      Left            =   2940
      TabIndex        =   37
      Top             =   780
      Width           =   3195
   End
   Begin VB.Line lineMiddle 
      BorderColor     =   &H00FFFF00&
      X1              =   3120
      X2              =   8400
      Y1              =   3180
      Y2              =   3180
   End
   Begin VB.Line lineRight 
      BorderColor     =   &H00FFFF00&
      X1              =   8460
      X2              =   8460
      Y1              =   780
      Y2              =   5640
   End
   Begin VB.Line lineLeft 
      BorderColor     =   &H00FFFF00&
      X1              =   3060
      X2              =   3060
      Y1              =   780
      Y2              =   5640
   End
   Begin VB.Line lineBottom 
      BorderColor     =   &H00FFFF00&
      X1              =   3000
      X2              =   8520
      Y1              =   5580
      Y2              =   5580
   End
   Begin VB.Line lineTop 
      BorderColor     =   &H00FFFF00&
      X1              =   3000
      X2              =   8520
      Y1              =   840
      Y2              =   840
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
      TabIndex        =   36
      Top             =   6120
      Width           =   675
   End
   Begin VB.Image imgBackToMenu 
      Height          =   195
      Left            =   120
      Picture         =   "frmRoyalties.frx":0040
      Stretch         =   -1  'True
      Top             =   6180
      Width           =   495
   End
   Begin VB.Label labRoyaltyMisc 
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
      Index           =   3
      Left            =   7920
      TabIndex        =   35
      Top             =   1020
      Width           =   435
   End
   Begin VB.Label labRoyaltyMisc 
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
      Left            =   960
      TabIndex        =   34
      Top             =   2160
      Width           =   1035
   End
   Begin VB.Label labRoyaltyUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   7
      Left            =   6900
      TabIndex        =   33
      Top             =   4260
      Width           =   45
   End
   Begin VB.Label labRoyaltyUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   10
      Left            =   6600
      TabIndex        =   32
      Top             =   5160
      Width           =   45
   End
   Begin VB.Label labRoyaltyUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   9
      Left            =   6600
      TabIndex        =   31
      Top             =   4860
      Width           =   45
   End
   Begin VB.Label labRoyaltyUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "/year"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   8
      Left            =   6600
      TabIndex        =   30
      Top             =   4560
      Width           =   390
   End
   Begin VB.Label labRoyaltyUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   6
      Left            =   6600
      TabIndex        =   29
      Top             =   3960
      Width           =   45
   End
   Begin VB.Label labRoyaltyUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   5
      Left            =   6600
      TabIndex        =   28
      Top             =   3660
      Width           =   45
   End
   Begin VB.Label labRoyaltyUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   4
      Left            =   6600
      TabIndex        =   27
      Top             =   2520
      Width           =   45
   End
   Begin VB.Label labRoyaltyUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "percent"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   3
      Left            =   6600
      TabIndex        =   26
      Top             =   2220
      Width           =   615
   End
   Begin VB.Label labRoyaltyUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "percent"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   2
      Left            =   6600
      TabIndex        =   25
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label labRoyaltyUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   6600
      TabIndex        =   24
      Top             =   1620
      Width           =   45
   End
   Begin VB.Label labRoyaltyUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "percent"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   6600
      TabIndex        =   23
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label labRoyaltyTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "to Year"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   3180
      TabIndex        =   11
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label labRoyaltyTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Payment from Year"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   3180
      TabIndex        =   10
      Top             =   4860
      Width           =   1815
   End
   Begin VB.Label labRoyaltyTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Escalate Amount"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   3180
      TabIndex        =   9
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label labRoyaltyTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Payment Type"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   3180
      TabIndex        =   8
      Top             =   4260
      Width           =   1815
   End
   Begin VB.Label labRoyaltyTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Annual Payment"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   3180
      TabIndex        =   7
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label labRoyaltyTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Royalty Cap/Buy-Out"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   3180
      TabIndex        =   6
      Top             =   3660
      Width           =   1815
   End
   Begin VB.Label labRoyaltyTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Calculation Method"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   3180
      TabIndex        =   5
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label labRoyaltyTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Net Profit Interest"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   3180
      TabIndex        =   4
      Top             =   2220
      Width           =   1815
   End
   Begin VB.Label labRoyaltyTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Federal - U.S."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   3180
      TabIndex        =   3
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label labRoyaltyTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Production"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   3180
      TabIndex        =   2
      Top             =   1620
      Width           =   1815
   End
   Begin VB.Label labRoyaltyTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Net Smelter Return"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   3180
      TabIndex        =   1
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label labRoyaltyHeading 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Royalties"
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
      TabIndex        =   0
      Top             =   120
      Width           =   1635
   End
End
Attribute VB_Name = "frmRoyalties"
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
  Select Case LastCell
    Case 0, 1, 2, 3, 5, 6, 8
      If labCheckTag(LastCell).Visible = False Then
        ParamSet = False
        dTag = dTag + 1
        labCheckTag(LastCell).Visible = True
        labCheckTag(LastCell).ForeColor = &HFFFF&
        labCheckTag(LastCell).Caption = LTrim(Str(nTag))
        If LastCell < 5 Then
          Tagged(hscSetNumbers.Value, LastCell + 118).Dependent = nTag
          DepTagData(nTag, dTag).TheCell = LastCell + 118
          DepTagData(nTag, dTag).Title = labRoyaltyTitle(LastCell).Caption & " Royalty"
          DepTagData(nTag, dTag).Units = labRoyaltyUnits(LastCell).Caption
        Else
          Tagged(hscSetNumbers.Value, LastCell + 120).Dependent = nTag
          DepTagData(nTag, dTag).TheCell = LastCell + 120
          If LastCell = 5 Then
            DepTagData(nTag, dTag).Title = labRoyaltyTitle(LastCell).Caption
          ElseIf LastCell = 6 Then
            DepTagData(nTag, dTag).Title = "Annual Royalty Payment"
          ElseIf LastCell = 8 Then
            DepTagData(nTag, dTag).Title = "Escalate Royalty Amount"
          End If
          DepTagData(nTag, dTag).Units = labRoyaltyUnits(LastCell).Caption
        End If
        DepTagData(nTag, dTag).SetNumber = hscSetNumbers.Value
      End If
  End Select
  txtRoyaltyValues(LastCell).SetFocus
End If

End Sub

Private Sub comIndTag_Click()

Select Case LastCell
  Case 0, 1, 2, 3, 5, 6, 8
    If labCheckTag(LastCell).Visible = False Then
      ParamSet = False
      nTag = nTag + 1
      dTag = 0
      labCheckTag(LastCell).Visible = True
      labCheckTag(LastCell).ForeColor = &HFF&
      labCheckTag(LastCell).Caption = LTrim(Str(nTag))
      If LastCell < 5 Then
        Tagged(hscSetNumbers.Value, LastCell + 118).Independent = nTag
        IndTagData(nTag).TheCell = LastCell + 118
        IndTagData(nTag).Title = labRoyaltyTitle(LastCell).Caption & " Royalty"
        IndTagData(nTag).Units = labRoyaltyUnits(LastCell).Caption
      Else
        Tagged(hscSetNumbers.Value, LastCell + 120).Independent = nTag
        IndTagData(nTag).TheCell = LastCell + 120
        If LastCell = 5 Then
          IndTagData(nTag).Title = labRoyaltyTitle(LastCell).Caption
        ElseIf LastCell = 6 Then
          IndTagData(nTag).Title = "Annual Royalty Payment"
        ElseIf LastCell = 8 Then
          IndTagData(nTag).Title = "Escalate Royalty Amount"
        End If
        IndTagData(nTag).Units = labRoyaltyUnits(LastCell).Caption
      End If
      IndTagData(nTag).SetNumber = hscSetNumbers.Value
    End If
End Select

txtRoyaltyValues(LastCell).SetFocus

End Sub
Private Sub Form_Activate()

Dim baseunit As String
Dim baselength As Integer

If IsHelpOn = True Then
  ShowMenu = True
  If LastCell = 100 Then
    txtRoySetLabel.SetFocus
  Else
    txtRoyaltyValues(LastCell).SetFocus
  End If
  IsHelpOn = False
Else
  baseunit = LTrim(RTrim(CommodityData(1, 0).reserves))
  If baseunit = "" Then baseunit = "tons"
  baselength = Len(baseunit) - 1
  baseunit = Left(baseunit, baselength)
  labRoyaltyUnits(1).Caption = "/" & baseunit & " ore"

  hscSetNumbers.Value = 1
  txtRoySetLabel.Text = Pn1(8, hscSetNumbers.Value)
  ShowMenu = True
  Call drawthevalues
  If InsertFlag = True Then
    labInsert.Caption = "Insert"
  Else
    labInsert.Caption = "Typeover"
  End If
  lstPayment.Visible = False
  LastCell = 0
  txtRoyaltyValues(0).SetFocus
End If

End Sub

Private Sub Form_Deactivate()

If ShowMenu = True Then
  frmRoyalties.Hide
  Call InputMenuAccess(1)
End If

End Sub

Private Sub Form_Load()

Dim i As Integer

If FullScreen = False Then
  frmRoyalties.Top = (Screen.Height - (frmRoyalties.Height + 350)) / 2
  frmRoyalties.Left = (Screen.Width - frmRoyalties.Width) / 2
Else
  frmRoyalties.Top = 0
  frmRoyalties.Left = 0
  frmRoyalties.WindowState = 2
End If

If frmRoyalties.Top < 0 Then frmRoyalties.Top = 0
If frmRoyalties.Left < 0 Then frmRoyalties.Left = 0

tempwide = frmRoyalties.ScaleWidth
temphigh = frmRoyalties.ScaleHeight

If PageChange(7) = True Then
  Call drawthevalues
End If

Call screenstuff

End Sub

Private Sub Form_Resize()

tempwide = frmRoyalties.ScaleWidth
temphigh = frmRoyalties.ScaleHeight

Call screenstuff

End Sub

Private Sub Form_Unload(Cancel As Integer)

  frmRoyalties.Hide
  If ShowMenu = True Then Call InputMenuAccess(1)

End Sub

Private Sub hscSetNumbers_Change()

  labSetNumbers.Caption = LTrim(RTrim(Str(hscSetNumbers.Value)))
  
  txtRoySetLabel.Text = Pn1(8, hscSetNumbers.Value)
  
  If hscSetNumbers.Value > Np(8) Then
    Np(8) = hscSetNumbers.Value
  End If
  
  If Np(8) > Npna Then Npna = Np(8)
  
  Call drawthevalues
  
  txtRoyaltyValues(0).SetFocus

End Sub

Private Sub imgBackToMenu_Click()

  frmRoyalties.Hide
  If ShowMenu = True Then Call InputMenuAccess(1)

End Sub

Private Sub labBackToMenu_Click()

  frmRoyalties.Hide
  If ShowMenu = True Then Call InputMenuAccess(1)

End Sub

Private Sub labRoyaltiesHelp_Click()

Dim begin As Integer
Dim sendindex As Integer

begin = 0
ShowMenu = False
WhichScreen = 7
If LastCell < 5 Then
  sendindex = LastCell + 118
ElseIf LastCell = 100 Then
  WhichScreen = 0
  sendindex = 20
Else
  sendindex = LastCell + 120
End If

Call frmApexHelp.gethelptext(sendindex, begin)
frmApexHelp.Show

End Sub

Private Sub lstPayment_Click()
  
  txtRoyaltyValues(7).Text = lstPayment.List(lstPayment.ListIndex)

  txtRoyaltyValues(8).SetFocus

End Sub

Private Sub txtRoyaltyValues_Change(Index As Integer)

If DoNotChange = True Then Exit Sub

PageChange(7) = True

If labCheckTag(Index).Visible = True Then ParamSet = False

If Index < 5 Then
  Primary(hscSetNumbers.Value, Index + 118) = CCur(Val(txtRoyaltyValues(Index).Text))
  Call recalc(hscSetNumbers.Value, Index + 118)
ElseIf Index = 7 Then
  Select Case LCase(LTrim(RTrim(txtRoyaltyValues(Index).Text)))
    Case "minimum royalty"
      Primary(hscSetNumbers.Value, Index + 120) = 2
    Case "lease-bonus"
      Primary(hscSetNumbers.Value, Index + 120) = 3
    Case Else
      Primary(hscSetNumbers.Value, Index + 120) = 1
  End Select
Else
  Primary(hscSetNumbers.Value, Index + 120) = CCur(Val(txtRoyaltyValues(Index).Text))
  Call recalc(hscSetNumbers.Value, Index + 120)
End If

End Sub

Private Sub txtRoyaltyValues_GotFocus(Index As Integer)

If Index = 7 Then
  lstPayment.Visible = True
Else
  lstPayment.Visible = False
End If

LastCell = Index

End Sub

Public Sub screenstuff()
 
  Dim X As Integer
  Dim Y As Currency
  
  labRoyaltyHeading.Top = temphigh * 0.0334
  labRoyaltyHeading.Left = tempwide * 0.0194
  
  LineTop.X1 = tempwide * 0.3279
  LineTop.X2 = tempwide * 0.9311
  LineTop.Y1 = temphigh * 0.1308
  LineTop.Y2 = temphigh * 0.1308
  
  lineMiddle.X1 = tempwide * 0.341
  lineMiddle.X2 = tempwide * 0.918
  lineMiddle.Y1 = temphigh * 0.4953
  lineMiddle.Y2 = temphigh * 0.4953
  
  LineBottom.X1 = tempwide * 0.3279
  LineBottom.X2 = tempwide * 0.9311
  LineBottom.Y1 = temphigh * 0.8692
  LineBottom.Y2 = temphigh * 0.8692
  
  LineLeft.X1 = tempwide * 0.3344
  LineLeft.X2 = tempwide * 0.3344
  LineLeft.Y1 = temphigh * 0.1215
  LineLeft.Y2 = temphigh * 0.8785

  LineRight.X1 = tempwide * 0.9246
  LineRight.X2 = tempwide * 0.9246
  LineRight.Y1 = temphigh * 0.1215
  LineRight.Y2 = temphigh * 0.8785
  
  For X = 0 To 10
    If X < 5 Then
      Y = 0
    Else
      Y = 0.1308
    End If
    labRoyaltyTitle(X).Top = (Y * temphigh) + (X * 0.0467 * temphigh) + (temphigh * 0.2056)
    labRoyaltyTitle(X).Left = tempwide * 0.3475
    labRoyaltyTitle(X).Width = tempwide * 0.1984
    txtRoyaltyValues(X).Top = (Y * temphigh) + (X * 0.0467 * temphigh) + (temphigh * 0.2009)
    txtRoyaltyValues(X).Left = tempwide * 0.5836
    labRoyaltyUnits(X).Top = (Y * temphigh) + (X * 0.0467 * temphigh) + (temphigh * 0.2056)
    labCheckTag(X).Top = (Y * temphigh) + (X * 0.0467 * temphigh) + (temphigh * 0.2056)
    labCheckTag(X).Left = tempwide * 0.8656
    labCheckTag(X).Width = tempwide * 0.0475
    If X = 7 Then
      txtRoyaltyValues(X).Width = tempwide * 0.1656
      labRoyaltyUnits(X).Left = tempwide * 0.7541
    Else
      txtRoyaltyValues(X).Width = tempwide * 0.1328
      labRoyaltyUnits(X).Left = tempwide * 0.7213
    End If
  Next X

  labRoyaltyMisc(0).Top = temphigh * 0.4112
  labRoyaltyMisc(0).Left = tempwide * 0.1049
  labRoyaltyMisc(0).Width = tempwide * 0.1131
  
  labRoyaltyMisc(1).Top = temphigh * 0.1215
  labRoyaltyMisc(1).Left = tempwide * 0.3213
    
  labRoyaltyMisc(2).Top = temphigh * 0.486
  labRoyaltyMisc(2).Left = tempwide * 0.3213
  
  labRoyaltyMisc(3).Top = temphigh * 0.1589
  labRoyaltyMisc(3).Left = tempwide * 0.8656
  labRoyaltyMisc(3).Width = tempwide * 0.0475
  
  For X = 4 To 7
    If X = 4 Then
      labRoyaltyMisc(X).Top = temphigh * 0.2523
    ElseIf X = 5 Then
      labRoyaltyMisc(X).Top = temphigh * 0.5701
    ElseIf X = 6 Then
      labRoyaltyMisc(X).Top = temphigh * 0.6168
    Else
      labRoyaltyMisc(X).Top = temphigh * 0.7103
    End If
    labRoyaltyMisc(X).Left = tempwide * 0.5639
    labRoyaltyMisc(X).Width = tempwide * 0.0148
  Next X

  labRoyaltyMisc(8).Top = temphigh * 0.5093
  labRoyaltyMisc(8).Left = tempwide * 0.1049
  labRoyaltyMisc(8).Width = tempwide * 0.1131
  
  hscSetNumbers.Top = temphigh * 0.4579
  hscSetNumbers.Left = tempwide * 0.1311
  
  labSetNumbers.Top = temphigh * 0.4532
  labSetNumbers.Left = tempwide * 0.177
  labSetNumbers.Width = tempwide * 0.0344
  
  txtRoySetLabel.Top = temphigh * 0.5514
  txtRoySetLabel.Left = tempwide * 0.1049
  txtRoySetLabel.Width = tempwide * 0.1131
  
  comIndTag.Top = temphigh * 0.8995
  comIndTag.Left = tempwide * 0.3535
  
  labIndTag.Top = temphigh * 0.8979
  labIndTag.Left = tempwide * 0.3863
  
  comDepTag.Top = temphigh * 0.8995
  comDepTag.Left = tempwide * 0.7344
  
  labDepTag.Top = temphigh * 0.8979
  labDepTag.Left = tempwide * 0.7672
  
  labBackToMenu.Top = temphigh * 0.9532
  labBackToMenu.Left = tempwide * 0.0656

  imgBackToMenu.Top = temphigh * 0.9626
  imgBackToMenu.Left = tempwide * 0.0066
  imgBackToMenu.Width = tempwide * 0.0541

  labRoyaltiesHelp.Top = temphigh * 0.9532
  labRoyaltiesHelp.Left = tempwide * 0.9377

  labInsert.Top = temphigh * 0.9562
  labInsert.Left = tempwide * 0.5836
  labInsert.Width = tempwide * 0.1066

End Sub

Public Sub drawthevalues()

Dim i As Integer
Dim X As Integer

DoNotChange = True

For i = 0 To 10
  If i < 5 Then
    txtRoyaltyValues(i).Text = LTrim(Str(Primary(hscSetNumbers.Value, i + 118)))
  ElseIf i = 7 Then
    Select Case Primary(hscSetNumbers.Value, i + 120)
      Case 2
        txtRoyaltyValues(i).Text = "Minimum Royalty"
      Case 3
        txtRoyaltyValues(i).Text = "Lease-Bonus"
      Case Else
        txtRoyaltyValues(i).Text = "Advance Royalty"
    End Select
  Else
    txtRoyaltyValues(i).Text = LTrim(Str(Primary(hscSetNumbers.Value, i + 120)))
  End If
  If i > 0 And i < 4 Then
    txtRoyaltyValues(i).Text = Format(txtRoyaltyValues(i).Text, "#######0.00")
  Else
    txtRoyaltyValues(i).Text = Format(txtRoyaltyValues(i).Text, "#########0")
  End If
Next i

For i = 118 To 130
  If i < 122 Then
    X = i - 118
  ElseIf i > 124 Then
    X = i - 120
  End If
  Select Case i
    Case 118 To 121, 125, 126, 128
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
  End Select
  
Next i

DoNotChange = False

End Sub

Private Sub txtRoyaltyValues_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

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
        If InStr(txtRoyaltyValues(Index).Text, ".") = 0 Then
          SendKeys "{DELETE}", False
        End If
      Else
        SendKeys "{DELETE}", False
      End If
  End Select
End If

End Sub

Private Sub txtRoyaltyValues_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii = 46 Then
  If InStr(txtRoyaltyValues(Index).Text, ".") > 0 Then
    Beep
    KeyAscii = 0
  End If
End If

If KeyAscii = 44 Then
  Beep
  KeyAscii = 0
End If

End Sub

Private Sub txtRoySetLabel_Change()

Pn1(8, hscSetNumbers.Value) = txtRoySetLabel.Text

End Sub


Private Sub txtRoySetLabel_GotFocus()

LastCell = 100

End Sub


