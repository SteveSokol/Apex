VERSION 5.00
Begin VB.Form frmFinancing 
   BackColor       =   &H00000000&
   Caption         =   "Financing"
   ClientHeight    =   6420
   ClientLeft      =   1140
   ClientTop       =   1380
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
   Begin VB.TextBox txtFinSetLabel 
      Height          =   330
      Left            =   780
      TabIndex        =   69
      Top             =   3540
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
      Left            =   6900
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   5940
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
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   5940
      Width           =   195
   End
   Begin VB.HScrollBar hscSetNumbers 
      Height          =   195
      Left            =   960
      Max             =   25
      Min             =   1
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   2940
      Value           =   1
      Width           =   375
   End
   Begin VB.TextBox txtFinancingValues 
      Height          =   330
      Index           =   12
      Left            =   5700
      TabIndex        =   12
      Text            =   "1"
      Top             =   5340
      Width           =   1155
   End
   Begin VB.TextBox txtFinancingValues 
      Height          =   330
      Index           =   11
      Left            =   5700
      TabIndex        =   11
      Text            =   "1"
      Top             =   5040
      Width           =   1155
   End
   Begin VB.TextBox txtFinancingValues 
      Height          =   330
      Index           =   10
      Left            =   5700
      TabIndex        =   10
      Text            =   "0.00"
      Top             =   4740
      Width           =   1155
   End
   Begin VB.TextBox txtFinancingValues 
      Height          =   330
      Index           =   9
      Left            =   5700
      TabIndex        =   9
      Text            =   "0.00"
      Top             =   4440
      Width           =   1155
   End
   Begin VB.TextBox txtFinancingValues 
      Height          =   330
      Index           =   8
      Left            =   5700
      TabIndex        =   8
      Text            =   "1"
      Top             =   4140
      Width           =   1155
   End
   Begin VB.TextBox txtFinancingValues 
      Height          =   330
      Index           =   7
      Left            =   5700
      TabIndex        =   7
      Text            =   "0"
      Top             =   3840
      Width           =   1155
   End
   Begin VB.TextBox txtFinancingValues 
      Height          =   330
      Index           =   6
      Left            =   5700
      TabIndex        =   6
      Text            =   "0.00"
      Top             =   3000
      Width           =   1155
   End
   Begin VB.TextBox txtFinancingValues 
      Height          =   330
      Index           =   5
      Left            =   5700
      TabIndex        =   5
      Text            =   "0.00"
      Top             =   2700
      Width           =   1155
   End
   Begin VB.TextBox txtFinancingValues 
      Height          =   330
      Index           =   4
      Left            =   5700
      TabIndex        =   4
      Text            =   "0"
      Top             =   2400
      Width           =   1155
   End
   Begin VB.TextBox txtFinancingValues 
      Height          =   330
      Index           =   3
      Left            =   5700
      TabIndex        =   3
      Text            =   "1"
      Top             =   2100
      Width           =   1155
   End
   Begin VB.TextBox txtFinancingValues 
      Height          =   330
      Index           =   2
      Left            =   5700
      TabIndex        =   2
      Text            =   "0.00"
      Top             =   1800
      Width           =   1155
   End
   Begin VB.TextBox txtFinancingValues 
      Height          =   330
      Index           =   1
      Left            =   5700
      TabIndex        =   1
      Text            =   "0"
      Top             =   1500
      Width           =   1155
   End
   Begin VB.TextBox txtFinancingValues 
      Height          =   330
      Index           =   0
      Left            =   5700
      TabIndex        =   0
      Text            =   "0"
      Top             =   1200
      Width           =   1155
   End
   Begin VB.Label labFinancingHelp 
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
      Height          =   225
      Left            =   8580
      TabIndex        =   72
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
      Height          =   315
      Left            =   5400
      TabIndex        =   71
      Top             =   6060
      Width           =   975
   End
   Begin VB.Label labFinancingMisc 
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
      Index           =   9
      Left            =   780
      TabIndex        =   70
      Top             =   3300
      Width           =   1035
   End
   Begin VB.Label labFinancingMisc 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "$"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   8
      Left            =   5520
      TabIndex        =   66
      Top             =   1260
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
      Index           =   12
      Left            =   8040
      TabIndex        =   65
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
      Left            =   8040
      TabIndex        =   64
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
      Left            =   8040
      TabIndex        =   63
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
      Left            =   8040
      TabIndex        =   62
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
      Left            =   8040
      TabIndex        =   61
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
      Left            =   8040
      TabIndex        =   60
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
      Left            =   8040
      TabIndex        =   59
      Top             =   3060
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
      Left            =   8040
      TabIndex        =   58
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
      Left            =   8040
      TabIndex        =   57
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
      Left            =   8040
      TabIndex        =   56
      Top             =   2160
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
      Left            =   8040
      TabIndex        =   55
      Top             =   1860
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
      Left            =   8040
      TabIndex        =   54
      Top             =   1560
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
      Left            =   8040
      TabIndex        =   53
      Top             =   1260
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label labDepTag 
      BackColor       =   &H00000000&
      Caption         =   "Dependent Tag"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   7200
      TabIndex        =   52
      Top             =   5940
      Width           =   1275
   End
   Begin VB.Label labIndTag 
      BackColor       =   &H00000000&
      Caption         =   "Independent Tag"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   3480
      TabIndex        =   51
      Top             =   5940
      Width           =   1455
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
      Left            =   1440
      TabIndex        =   50
      Top             =   2940
      Width           =   255
   End
   Begin VB.Label labFinancingMisc 
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
      Index           =   7
      Left            =   780
      TabIndex        =   48
      Top             =   2640
      Width           =   1035
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
      TabIndex        =   47
      Top             =   6120
      Width           =   615
   End
   Begin VB.Image imgBackToMenu 
      Height          =   195
      Left            =   60
      Picture         =   "frmFinancing.frx":0000
      Stretch         =   -1  'True
      Top             =   6180
      Width           =   495
   End
   Begin VB.Label labFinancingHeading 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Project Financing"
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
      TabIndex        =   46
      Top             =   120
      Width           =   2955
   End
   Begin VB.Label labFinancingMisc 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Height          =   225
      Index           =   6
      Left            =   8040
      TabIndex        =   45
      Top             =   1020
      Width           =   420
   End
   Begin VB.Label labFinancingMisc 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "$"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   5
      Left            =   5520
      TabIndex        =   44
      Top             =   4800
      Width           =   135
   End
   Begin VB.Label labFinancingMisc 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "$"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   4
      Left            =   5520
      TabIndex        =   43
      Top             =   3900
      Width           =   135
   End
   Begin VB.Label labFinancingMisc 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "$"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   3
      Left            =   5520
      TabIndex        =   42
      Top             =   3060
      Width           =   135
   End
   Begin VB.Label labFinancingMisc 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "$"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   2
      Left            =   5550
      TabIndex        =   41
      Top             =   2760
      Width           =   105
   End
   Begin VB.Label labFinancingMisc 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Joint Ventures"
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
      Left            =   2940
      TabIndex        =   40
      Top             =   3480
      Width           =   1275
   End
   Begin VB.Label labFinancingMisc 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Loans"
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
      Index           =   0
      Left            =   2940
      TabIndex        =   39
      Top             =   840
      Width           =   555
   End
   Begin VB.Line LineRight 
      BorderColor     =   &H00FFFF00&
      X1              =   8580
      X2              =   8580
      Y1              =   840
      Y2              =   5880
   End
   Begin VB.Line lineLeft 
      BorderColor     =   &H00FFFF00&
      X1              =   3060
      X2              =   3060
      Y1              =   840
      Y2              =   5880
   End
   Begin VB.Line lineBottom 
      BorderColor     =   &H00FFFF00&
      X1              =   3000
      X2              =   8640
      Y1              =   5820
      Y2              =   5820
   End
   Begin VB.Line lineMiddle 
      BorderColor     =   &H00FFFF00&
      X1              =   3120
      X2              =   8520
      Y1              =   3540
      Y2              =   3540
   End
   Begin VB.Line lineTop 
      BorderColor     =   &H00FFFF00&
      X1              =   3000
      X2              =   8640
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Label labFinancingUnits 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   6900
      TabIndex        =   38
      Top             =   5400
      Width           =   960
   End
   Begin VB.Label labFinancingUnits 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   6900
      TabIndex        =   37
      Top             =   5100
      Width           =   960
   End
   Begin VB.Label labFinancingUnits 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   6900
      TabIndex        =   36
      Top             =   4800
      Width           =   960
   End
   Begin VB.Label labFinancingUnits 
      BackColor       =   &H00000000&
      Caption         =   "percent"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   6900
      TabIndex        =   35
      Top             =   4500
      Width           =   960
   End
   Begin VB.Label labFinancingUnits 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   6900
      TabIndex        =   34
      Top             =   4200
      Width           =   960
   End
   Begin VB.Label labFinancingUnits 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   6900
      TabIndex        =   33
      Top             =   3900
      Width           =   960
   End
   Begin VB.Label labFinancingUnits 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   6900
      TabIndex        =   32
      Top             =   3060
      Width           =   960
   End
   Begin VB.Label labFinancingUnits 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   6900
      TabIndex        =   31
      Top             =   2760
      Width           =   960
   End
   Begin VB.Label labFinancingUnits 
      BackColor       =   &H00000000&
      Caption         =   "years"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   6900
      TabIndex        =   30
      Top             =   2460
      Width           =   960
   End
   Begin VB.Label labFinancingUnits 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   6900
      TabIndex        =   29
      Top             =   2160
      Width           =   960
   End
   Begin VB.Label labFinancingUnits 
      BackColor       =   &H00000000&
      Caption         =   "percent"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   6900
      TabIndex        =   28
      Top             =   1860
      Width           =   960
   End
   Begin VB.Label labFinancingUnits 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   6900
      TabIndex        =   27
      Top             =   1560
      Width           =   960
   End
   Begin VB.Label labFinancingUnits 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   6900
      TabIndex        =   26
      Top             =   1260
      Width           =   960
   End
   Begin VB.Label labFinancingTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "To Year"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   3180
      TabIndex        =   25
      Top             =   5400
      Width           =   2205
   End
   Begin VB.Label labFinancingTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "From Year"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   3180
      TabIndex        =   24
      Top             =   5100
      Width           =   2205
   End
   Begin VB.Label labFinancingTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Management Fee"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   3180
      TabIndex        =   23
      Top             =   4800
      Width           =   2205
   End
   Begin VB.Label labFinancingTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Partner's Share"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   3180
      TabIndex        =   22
      Top             =   4500
      Width           =   2205
   End
   Begin VB.Label labFinancingTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Invested in Year"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   3180
      TabIndex        =   21
      Top             =   4200
      Width           =   2205
   End
   Begin VB.Label labFinancingTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Partner's Investment"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   3180
      TabIndex        =   20
      Top             =   3900
      Width           =   2205
   End
   Begin VB.Label labFinancingTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Floor Price"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   3180
      TabIndex        =   19
      Top             =   3060
      Width           =   2205
   End
   Begin VB.Label labFinancingTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Margin-Free Limit"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   3180
      TabIndex        =   18
      Top             =   2760
      Width           =   2205
   End
   Begin VB.Label labFinancingTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Loan Period"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   3180
      TabIndex        =   17
      Top             =   2460
      Width           =   2205
   End
   Begin VB.Label labFinancingTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Borrowed in Year"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   3180
      TabIndex        =   16
      Top             =   2160
      Width           =   2205
   End
   Begin VB.Label labFinancingTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Interest Rate"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   3180
      TabIndex        =   15
      Top             =   1860
      Width           =   2205
   End
   Begin VB.Label labFinancingTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Amount - Commodity"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   3180
      TabIndex        =   14
      Top             =   1560
      Width           =   2205
   End
   Begin VB.Label labFinancingTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Amount - Conventional"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   3180
      TabIndex        =   13
      Top             =   1260
      Width           =   2205
   End
End
Attribute VB_Name = "frmFinancing"
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
    Case 0, 1, 2, 4, 5, 6, 7, 9, 10
      If labCheckTag(LastCell).Visible = False Then
        ParamSet = False
        dTag = dTag + 1
        labCheckTag(LastCell).Visible = True
        labCheckTag(LastCell).ForeColor = &HFFFF&
        labCheckTag(LastCell).Caption = LTrim(Str(nTag))
        Tagged(hscSetNumbers.Value, LastCell + 105).Dependent = nTag
        DepTagData(nTag, dTag).TheCell = LastCell + 105
        DepTagData(nTag, dTag).Title = "Loan " & LTrim(RTrim(labFinancingTitles(LastCell).Caption))
        If LastCell = 5 Or LastCell = 6 Then
          DepTagData(nTag, dTag).Title = "Commodity " & LTrim(RTrim(labFinancingTitles(LastCell).Caption))
        ElseIf LastCell = 7 Or LastCell = 9 Or LastCell = 10 Then
          DepTagData(nTag, dTag).Title = "Joint Venture " & LTrim(RTrim(labFinancingTitles(LastCell).Caption))
        End If
        DepTagData(nTag, dTag).Units = LTrim(RTrim(labFinancingUnits(LastCell).Caption))
        DepTagData(nTag, dTag).SetNumber = hscSetNumbers.Value
      End If
  End Select
  txtFinancingValues(LastCell).SetFocus
End If

End Sub


Private Sub comIndTag_Click()

Select Case LastCell
  Case 0, 1, 2, 4, 5, 6, 7, 9, 10
    If labCheckTag(LastCell).Visible = False Then
      ParamSet = False
      nTag = nTag + 1
      dTag = 0
      labCheckTag(LastCell).Visible = True
      labCheckTag(LastCell).ForeColor = &HFF&
      labCheckTag(LastCell).Caption = LTrim(Str(nTag))
      Tagged(hscSetNumbers.Value, LastCell + 105).Independent = nTag
      IndTagData(nTag).TheCell = LastCell + 105
      IndTagData(nTag).Title = "Loan " & LTrim(RTrim(labFinancingTitles(LastCell).Caption))
      If LastCell = 5 Or LastCell = 6 Then
        IndTagData(nTag).Title = "Commodity " & LTrim(RTrim(labFinancingTitles(LastCell).Caption))
      ElseIf LastCell = 7 Or LastCell = 9 Or LastCell = 10 Then
        IndTagData(nTag).Title = "Joint Venture " & LTrim(RTrim(labFinancingTitles(LastCell).Caption))
      End If
      IndTagData(nTag).Units = LTrim(RTrim(labFinancingUnits(LastCell).Caption))
      IndTagData(nTag).SetNumber = hscSetNumbers.Value
    End If
  End Select

txtFinancingValues(LastCell).SetFocus

End Sub

Private Sub Form_Activate()

Dim baseunit As String
Dim baselength As Integer

If IsHelpOn = True Then
  If LastCell = 100 Then
    txtFinSetLabel.SetFocus
  Else
    txtFinancingValues(LastCell).SetFocus
  End If
  IsHelpOn = False
Else
  baseunit = LTrim(RTrim(CommodityData(1, 0).reserves))
  If baseunit = "" Then baseunit = "tons"
  
  baselength = Len(baseunit) - 1
  baseunit = Left(baseunit, baselength)
  labFinancingUnits(10).Caption = "/" & baseunit & " ore"
  
  baseunit = LTrim(RTrim(CommodityData(1, 0).Price))
    labFinancingUnits(5).Caption = baseunit
  labFinancingUnits(6).Caption = baseunit
  
  baselength = Len(baseunit) - 1
  If baselength > 0 Then
    baseunit = Right(baseunit, baselength)
    labFinancingUnits(1).Caption = baseunit & "s"
  End If
  
  hscSetNumbers.Value = 1
  txtFinSetLabel.Text = Pn1(7, hscSetNumbers.Value)
  Call drawthevalues
  
  If InsertFlag = True Then
    labInsert.Caption = "Insert"
  Else
    labInsert.Caption = "Typeover"
  End If
  ShowMenu = True
  txtFinancingValues(0).SetFocus
  LastCell = 0
End If

End Sub

Private Sub Form_Deactivate()

If ShowMenu = True Then
  frmFinancing.Hide
  Call InputMenuAccess(1)
End If

End Sub

Private Sub Form_Load()

Dim i As Integer

If FullScreen = False Then
  frmFinancing.Top = (Screen.Height - (frmFinancing.Height + 350)) / 2
  frmFinancing.Left = (Screen.Width - frmFinancing.Width) / 2
Else
  frmFinancing.Top = 0
  frmFinancing.Left = 0
  frmFinancing.WindowState = 2
End If

If frmFinancing.Top < 0 Then frmFinancing.Top = 0
If frmFinancing.Left < 0 Then frmFinancing.Left = 0

tempwide = frmFinancing.ScaleWidth
temphigh = frmFinancing.ScaleHeight

If PageChange(6) = True Then Call drawthevalues

Call screenstuff
  
End Sub

Private Sub Form_Resize()

tempwide = frmFinancing.ScaleWidth
temphigh = frmFinancing.ScaleHeight

Call screenstuff

End Sub


Private Sub Form_Unload(Cancel As Integer)

  frmFinancing.Hide
  If ShowMenu = True Then Call InputMenuAccess(1)

End Sub

Private Sub hscSetNumbers_Change()
 
  labSetNumbers.Caption = LTrim(RTrim(Str(hscSetNumbers.Value)))
  
  txtFinSetLabel.Text = Pn1(7, hscSetNumbers.Value)
  
  If hscSetNumbers.Value > Np(7) Then
    Np(7) = hscSetNumbers.Value
  End If
  
  If Np(7) > Npna Then Npna = Np(7)
  
  Call drawthevalues
  
  txtFinancingValues(0).SetFocus
  
End Sub

Private Sub imgBackToMenu_Click()
  
  frmFinancing.Hide
  If ShowMenu = True Then Call InputMenuAccess(1)

End Sub

Private Sub labBackToMenu_Click()
  
  frmFinancing.Hide
  If ShowMenu = True Then Call InputMenuAccess(1)

End Sub

Private Sub labFinincingTitles_Click(Index As Integer)

End Sub

Private Sub labFinancingHelp_Click()
Dim begin As Integer
Dim sendindex As Integer
ShowMenu = False
begin = 0
sendindex = LastCell + 105

WhichScreen = 6

If LastCell = 100 Then
  WhichScreen = 0
  sendindex = 20
End If

Call frmApexHelp.gethelptext(sendindex, begin)
frmApexHelp.Show

End Sub

Private Sub txtFinancingValues_Change(Index As Integer)

If DoNotChange = True Then Exit Sub

PageChange(6) = True

If labCheckTag(Index).Visible = True Then ParamSet = False

Primary(hscSetNumbers.Value, Index + 105) = CCur(Val(txtFinancingValues(Index).Text))

Call recalc(hscSetNumbers.Value, Index + 105)

End Sub


Private Sub txtFinancingValues_GotFocus(Index As Integer)

LastCell = Index

End Sub



Public Sub screenstuff()
  
  Dim X As Integer
  Dim Y As Currency
  
  labFinancingHeading.Top = temphigh * 0.0334
  labFinancingHeading.Left = tempwide * 0.0194
  
  LineTop.X1 = tempwide * 0.3279
  LineTop.X2 = tempwide * 0.9443
  LineTop.Y1 = temphigh * 0.1402
  LineTop.Y2 = temphigh * 0.1402
  
  lineMiddle.X1 = tempwide * 0.341
  lineMiddle.X2 = tempwide * 0.9311
  lineMiddle.Y1 = temphigh * 0.5514
  lineMiddle.Y2 = temphigh * 0.5514
  
  LineBottom.X1 = tempwide * 0.3279
  LineBottom.X2 = tempwide * 0.9443
  LineBottom.Y1 = temphigh * 0.9065
  LineBottom.Y2 = temphigh * 0.9065
  
  LineLeft.X1 = tempwide * 0.3344
  LineLeft.X2 = tempwide * 0.3344
  LineLeft.Y1 = temphigh * 0.1308
  LineLeft.Y2 = temphigh * 0.9159

  LineRight.X1 = tempwide * 0.9377
  LineRight.X2 = tempwide * 0.9377
  LineRight.Y1 = temphigh * 0.1308
  LineRight.Y2 = temphigh * 0.9159
  
  For X = 0 To 12
    If X < 7 Then
      Y = 0
    Else
      Y = 0.0841
    End If
    labFinancingTitles(X).Top = (Y * temphigh) + (X * 0.0467 * temphigh) + (temphigh * 0.1963)
    labFinancingTitles(X).Left = tempwide * 0.3475
    labFinancingTitles(X).Width = tempwide * 0.241
    txtFinancingValues(X).Top = (Y * temphigh) + (X * 0.0467 * temphigh) + (temphigh * 0.1916)
    txtFinancingValues(X).Left = tempwide * 0.623
    txtFinancingValues(X).Width = tempwide * 0.1262
    labFinancingUnits(X).Top = (Y * temphigh) + (X * 0.0467 * temphigh) + (temphigh * 0.1963)
    labFinancingUnits(X).Left = tempwide * 0.754
    labCheckTag(X).Top = (Y * temphigh) + (X * 0.0467 * temphigh) + (temphigh * 0.1963)
    labCheckTag(X).Left = tempwide * 0.8787
    labCheckTag(X).Width = tempwide * 0.0475
  Next X

  labFinancingMisc(0).Top = temphigh * 0.1308
  labFinancingMisc(0).Left = tempwide * 0.3213
    
  labFinancingMisc(1).Top = temphigh * 0.5421
  labFinancingMisc(1).Left = tempwide * 0.3213
  
  For X = 2 To 5
    If X = 2 Then
      Y = 0.4299
    ElseIf X = 3 Then
      Y = 0.4766
    ElseIf X = 4 Then
      Y = 0.6075
    Else
      Y = 0.7477
    End If
    labFinancingMisc(X).Top = temphigh * Y
    labFinancingMisc(X).Left = tempwide * 0.6033
    labFinancingMisc(X).Width = tempwide * 0.0148
  Next X
  
  labFinancingMisc(6).Top = temphigh * 0.1589
  labFinancingMisc(6).Left = tempwide * 0.8787
  labFinancingMisc(6).Width = tempwide * 0.0459
  
  labFinancingMisc(7).Top = temphigh * 0.4112
  labFinancingMisc(7).Left = tempwide * 0.0852
  labFinancingMisc(7).Width = tempwide * 0.1131
  
  labFinancingMisc(8).Top = temphigh * 0.1963
  labFinancingMisc(8).Left = tempwide * 0.6033
  labFinancingMisc(8).Width = tempwide * 0.0148
  
  labFinancingMisc(9).Top = temphigh * 0.5093
  labFinancingMisc(9).Left = tempwide * 0.0852
  labFinancingMisc(9).Width = tempwide * 0.1131
  
  comIndTag.Top = temphigh * 0.9268
  comIndTag.Left = tempwide * 0.3535
  
  labIndTag.Top = temphigh * 0.9252
  labIndTag.Left = tempwide * 0.3863
  
  comDepTag.Top = temphigh * 0.9268
  comDepTag.Left = tempwide * 0.7541
  
  labDepTag.Top = temphigh * 0.9252
  labDepTag.Left = tempwide * 0.7869
  
  hscSetNumbers.Top = temphigh * 0.4579
  hscSetNumbers.Left = tempwide * 0.1115
  
  labSetNumbers.Top = temphigh * 0.4532
  labSetNumbers.Left = tempwide * 0.1574
  labSetNumbers.Width = tempwide * 0.0279
  
  txtFinSetLabel.Top = temphigh * 0.5514
  txtFinSetLabel.Left = tempwide * 0.0852
  txtFinSetLabel.Width = tempwide * 0.1131
  
  labBackToMenu.Top = temphigh * 0.9532
  labBackToMenu.Left = tempwide * 0.0656

  imgBackToMenu.Top = temphigh * 0.9626
  imgBackToMenu.Left = tempwide * 0.0066
  imgBackToMenu.Width = tempwide * 0.0541
  
  labFinancingHelp.Top = temphigh * 0.9532
  labFinancingHelp.Left = tempwide * 0.9377

  labInsert.Top = temphigh * 0.9562
  labInsert.Left = tempwide * 0.5902
  labInsert.Width = tempwide * 0.1066
    
End Sub

Public Sub drawthevalues()

Dim i As Integer

DoNotChange = True

  For i = 0 To 12
    Select Case i
      Case 0, 1, 3, 4, 7, 8, 11, 12
        txtFinancingValues(i).Text = Format(LTrim(Str(Primary(hscSetNumbers.Value, i + 105))), "#########0")
      Case Else
        txtFinancingValues(i).Text = Format(LTrim(Str(Primary(hscSetNumbers.Value, i + 105))), "#####0.00")
    End Select
  Next i

  For i = 105 To 117
    Select Case i - 105
      Case 0, 1, 2, 4, 5, 6, 7, 9, 10
        labCheckTag(i - 105).Visible = False
        If Tagged(hscSetNumbers.Value, i).Independent > 0 Then
          labCheckTag(i - 105).Visible = True
          labCheckTag(i - 105).ForeColor = &HFF&
          labCheckTag(i - 105).Caption = LTrim(RTrim(Str(Tagged(hscSetNumbers.Value, i).Independent)))
        ElseIf Tagged(hscSetNumbers.Value, i).Dependent > 0 Then
          labCheckTag(i - 105).Visible = True
          labCheckTag(i - 105).ForeColor = &HFFFF&
          labCheckTag(i - 105).Caption = LTrim(RTrim(Str(Tagged(hscSetNumbers.Value, i).Dependent)))
        Else
          labCheckTag(i - 105).Caption = ""
        End If
    End Select
  Next i
  
DoNotChange = False

End Sub

Private Sub txtFinancingValues_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

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
        If InStr(txtFinancingValues(Index).Text, ".") = 0 Then
          SendKeys "{DELETE}", False
        End If
      Else
        SendKeys "{DELETE}", False
      End If
  End Select
End If

End Sub

Private Sub txtFinancingValues_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii = 46 Then
  If InStr(txtFinancingValues(Index).Text, ".") > 0 Then
    Beep
    KeyAscii = 0
  End If
End If

If KeyAscii = 44 Then
  Beep
  KeyAscii = 0
End If

End Sub

Private Sub txtFinSetLabel_Change()

Pn1(7, hscSetNumbers.Value) = txtFinSetLabel.Text

End Sub


Private Sub txtFinSetLabel_GotFocus()

LastCell = 100

End Sub


