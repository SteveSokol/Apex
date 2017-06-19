VERSION 5.00
Begin VB.Form frmProcessingCost 
   BackColor       =   &H00000000&
   Caption         =   "Processing Costs"
   ClientHeight    =   6390
   ClientLeft      =   1545
   ClientTop       =   1860
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
   ScaleHeight     =   6390
   ScaleWidth      =   9150
   Begin VB.TextBox txtProSetLabel 
      Height          =   330
      Left            =   420
      TabIndex        =   67
      Top             =   3420
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
      Left            =   7020
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   5820
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
      Left            =   2040
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   5820
      Width           =   195
   End
   Begin VB.HScrollBar hscSetNumbers 
      Height          =   195
      Left            =   600
      Max             =   25
      Min             =   1
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   2640
      Value           =   1
      Width           =   375
   End
   Begin VB.TextBox txtProcessingValues 
      Height          =   330
      Index           =   8
      Left            =   4920
      TabIndex        =   20
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox txtProcessingValues 
      Height          =   330
      Index           =   7
      Left            =   4920
      TabIndex        =   19
      Top             =   4140
      Width           =   1095
   End
   Begin VB.TextBox txtProcessingValues 
      Height          =   330
      Index           =   6
      Left            =   4920
      TabIndex        =   18
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox txtProcessingValues 
      Height          =   330
      Index           =   5
      Left            =   4920
      TabIndex        =   17
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtProcessingValues 
      Height          =   330
      Index           =   4
      Left            =   4920
      TabIndex        =   16
      Top             =   2460
      Width           =   1095
   End
   Begin VB.TextBox txtProcessingValues 
      Height          =   330
      Index           =   3
      Left            =   4920
      TabIndex        =   15
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txtProcessingValues 
      Height          =   330
      Index           =   2
      Left            =   4920
      TabIndex        =   14
      Top             =   1860
      Width           =   1095
   End
   Begin VB.TextBox txtProcessingValues 
      Height          =   330
      Index           =   1
      Left            =   4920
      TabIndex        =   13
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtProcessingValues 
      Height          =   330
      Index           =   0
      Left            =   4920
      TabIndex        =   12
      Top             =   1260
      Width           =   1095
   End
   Begin VB.Label labProcessHelp 
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
      Left            =   8640
      TabIndex        =   70
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
      Left            =   4920
      TabIndex        =   69
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label labProcessingLabels 
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
      Index           =   14
      Left            =   420
      TabIndex        =   68
      Top             =   3180
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
      Index           =   8
      Left            =   8100
      TabIndex        =   64
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
      Index           =   7
      Left            =   8100
      TabIndex        =   63
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
      Index           =   6
      Left            =   8100
      TabIndex        =   62
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
      Index           =   5
      Left            =   8100
      TabIndex        =   61
      Top             =   2820
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
      Left            =   8100
      TabIndex        =   60
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
      Left            =   8100
      TabIndex        =   59
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
      Left            =   8100
      TabIndex        =   58
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
      Left            =   8100
      TabIndex        =   57
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
      Left            =   8100
      TabIndex        =   56
      Top             =   1320
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label labDepTag 
      BackColor       =   &H00000000&
      Caption         =   "Dependent Tag"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   7320
      TabIndex        =   55
      Top             =   5820
      Width           =   1275
   End
   Begin VB.Label labIndTag 
      BackColor       =   &H00000000&
      Caption         =   "Independent Tag"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   2340
      TabIndex        =   54
      Top             =   5820
      Width           =   1455
   End
   Begin VB.Label labScreenTotals 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0.00"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   4980
      TabIndex        =   53
      Top             =   5340
      Width           =   975
   End
   Begin VB.Label labScreenTotals 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0.00"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   4980
      TabIndex        =   52
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label labScreenTotals 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0.00"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   4980
      TabIndex        =   51
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label labProcessingLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "$"
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Index           =   13
      Left            =   4740
      TabIndex        =   50
      Top             =   5340
      Width           =   135
   End
   Begin VB.Label labProcessingLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "$"
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Index           =   12
      Left            =   4740
      TabIndex        =   49
      Top             =   4800
      Width           =   135
   End
   Begin VB.Label labProcessingLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "$"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   11
      Left            =   4740
      TabIndex        =   48
      Top             =   4200
      Width           =   135
   End
   Begin VB.Label labProcessingLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "$"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   10
      Left            =   4740
      TabIndex        =   47
      Top             =   3900
      Width           =   135
   End
   Begin VB.Label labProcessingLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "$"
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Index           =   9
      Left            =   4740
      TabIndex        =   46
      Top             =   3120
      Width           =   135
   End
   Begin VB.Label labProcessingLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "$"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   8
      Left            =   4740
      TabIndex        =   45
      Top             =   2820
      Width           =   135
   End
   Begin VB.Label labProcessingLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "$"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   7
      Left            =   4740
      TabIndex        =   44
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label labProcessingLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "$"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   6
      Left            =   4740
      TabIndex        =   43
      Top             =   2220
      Width           =   135
   End
   Begin VB.Label labProcessingLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "$"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   5
      Left            =   4740
      TabIndex        =   42
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label labProcessingLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "$"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   4
      Left            =   4740
      TabIndex        =   41
      Top             =   1620
      Width           =   135
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
      Left            =   555
      TabIndex        =   40
      Top             =   6120
      Width           =   675
   End
   Begin VB.Image imgBackToMenu 
      Height          =   195
      Left            =   60
      Picture         =   "frmProcessing.frx":0000
      Stretch         =   -1  'True
      Top             =   6180
      Width           =   495
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
      Left            =   1020
      TabIndex        =   39
      Top             =   2640
      Width           =   315
   End
   Begin VB.Label labProcessingLabels 
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
      Index           =   3
      Left            =   420
      TabIndex        =   37
      Top             =   2400
      Width           =   1035
   End
   Begin VB.Label labProcessingLabels 
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
      Index           =   2
      Left            =   8100
      TabIndex        =   36
      Top             =   1020
      Width           =   435
   End
   Begin VB.Label labProcessingLabels 
      BackColor       =   &H00000000&
      Caption         =   "Transportation and Smelting Costs"
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
      Left            =   1860
      TabIndex        =   35
      Top             =   3480
      Width           =   3075
   End
   Begin VB.Label labProcessingLabels 
      BackColor       =   &H00000000&
      Caption         =   "Mill Operating Costs"
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
      Left            =   1860
      TabIndex        =   34
      Top             =   840
      Width           =   1875
   End
   Begin VB.Line LineMiddle 
      BorderColor     =   &H00FFFF00&
      X1              =   2040
      X2              =   8580
      Y1              =   3540
      Y2              =   3540
   End
   Begin VB.Line LineBottom 
      BorderColor     =   &H00FFFF00&
      X1              =   1920
      X2              =   8700
      Y1              =   5700
      Y2              =   5700
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFF00&
      X1              =   1920
      X2              =   8700
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line LineRight 
      BorderColor     =   &H00FFFF00&
      X1              =   8640
      X2              =   8640
      Y1              =   900
      Y2              =   5760
   End
   Begin VB.Line LineLeft 
      BorderColor     =   &H00FFFF00&
      X1              =   1980
      X2              =   1980
      Y1              =   900
      Y2              =   5760
   End
   Begin VB.Label labProcessingUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Index           =   11
      Left            =   6060
      TabIndex        =   33
      Top             =   5340
      Width           =   45
   End
   Begin VB.Label labProcessingUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Index           =   10
      Left            =   6060
      TabIndex        =   32
      Top             =   4800
      Width           =   45
   End
   Begin VB.Label labProcessingUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   9
      Left            =   6060
      TabIndex        =   31
      Top             =   4500
      Width           =   45
   End
   Begin VB.Label labProcessingUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   8
      Left            =   6060
      TabIndex        =   30
      Top             =   4200
      Width           =   45
   End
   Begin VB.Label labProcessingUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   7
      Left            =   6060
      TabIndex        =   29
      Top             =   3900
      Width           =   45
   End
   Begin VB.Label labProcessingHeading 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Processing Costs"
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
      Left            =   90
      TabIndex        =   28
      Top             =   120
      Width           =   2880
   End
   Begin VB.Label labProcessingUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Index           =   6
      Left            =   6060
      TabIndex        =   27
      Top             =   3120
      Width           =   45
   End
   Begin VB.Label labProcessingUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   5
      Left            =   6060
      TabIndex        =   26
      Top             =   2820
      Width           =   45
   End
   Begin VB.Label labProcessingUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   4
      Left            =   6060
      TabIndex        =   25
      Top             =   2520
      Width           =   45
   End
   Begin VB.Label labProcessingUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   3
      Left            =   6060
      TabIndex        =   24
      Top             =   2220
      Width           =   45
   End
   Begin VB.Label labProcessingUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   2
      Left            =   6060
      TabIndex        =   23
      Top             =   1920
      Width           =   45
   End
   Begin VB.Label labProcessingUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "/day"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   6060
      TabIndex        =   22
      Top             =   1620
      Width           =   330
   End
   Begin VB.Label labProcessingUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   6060
      TabIndex        =   21
      Top             =   1320
      Width           =   45
   End
   Begin VB.Label labProcessingTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Screen Total"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   11
      Left            =   2040
      TabIndex        =   11
      Top             =   5340
      Width           =   2505
   End
   Begin VB.Label labProcessingTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Subtotal"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   10
      Left            =   2040
      TabIndex        =   10
      Top             =   4800
      Width           =   2505
   End
   Begin VB.Label labProcessingTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Concentration Ratio"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   2040
      TabIndex        =   9
      Top             =   4500
      Width           =   2505
   End
   Begin VB.Label labProcessingTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Smelting"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   2040
      TabIndex        =   8
      Top             =   4200
      Width           =   2505
   End
   Begin VB.Label labProcessingTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Transportation"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   2040
      TabIndex        =   7
      Top             =   3900
      Width           =   2505
   End
   Begin VB.Label labProcessingTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Subtotal"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   6
      Left            =   2100
      TabIndex        =   6
      Top             =   3120
      Width           =   2445
   End
   Begin VB.Label labProcessingTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "(Equipment Operating Cost)"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   2100
      TabIndex        =   5
      Top             =   2820
      Width           =   2445
   End
   Begin VB.Label labProcessingTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "(Supply Cost)"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   2100
      TabIndex        =   4
      Top             =   2520
      Width           =   2445
   End
   Begin VB.Label labProcessingTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "(Labor Cost)"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   2100
      TabIndex        =   3
      Top             =   2220
      Width           =   2445
   End
   Begin VB.Label labProcessingTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Variable Cost"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   2100
      TabIndex        =   2
      Top             =   1920
      Width           =   2445
   End
   Begin VB.Label labProcessingTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Fixed Cost"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   2100
      TabIndex        =   1
      Top             =   1620
      Width           =   2445
   End
   Begin VB.Label labProcessingTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Production Rate"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   2100
      TabIndex        =   0
      Top             =   1320
      Width           =   2445
   End
End
Attribute VB_Name = "frmProcessingCost"
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
    If LastCell < 6 Then
      Tagged(hscSetNumbers.Value, LastCell + 92).Dependent = nTag
      DepTagData(nTag, dTag).TheCell = LastCell + 92
      DepTagData(nTag, dTag).Title = "Milling - " & labProcessingTitles(LastCell).Caption
      If LastCell = 0 Then
        DepTagData(nTag, dTag).Title = "Mill " & labProcessingTitles(LastCell).Caption
      ElseIf LastCell = 2 Then
        DepTagData(nTag, dTag).Title = "Variable Milling Cost"
      End If
      DepTagData(nTag, dTag).Units = labProcessingUnits(LastCell).Caption
    Else
      Tagged(hscSetNumbers.Value, LastCell + 95).Dependent = nTag
      DepTagData(nTag, dTag).TheCell = LastCell + 95
      DepTagData(nTag, dTag).Title = labProcessingTitles(LastCell + 1).Caption
      DepTagData(nTag, dTag).Units = labProcessingUnits(LastCell + 1).Caption
    End If
    DepTagData(nTag, dTag).SetNumber = hscSetNumbers.Value
  End If
  txtProcessingValues(LastCell).SetFocus
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
  If LastCell < 6 Then
    Tagged(hscSetNumbers.Value, LastCell + 92).Independent = nTag
    IndTagData(nTag).TheCell = LastCell + 92
    IndTagData(nTag).Title = "Milling - " & labProcessingTitles(LastCell).Caption
    If LastCell = 0 Then
      IndTagData(nTag).Title = "Mill " & labProcessingTitles(LastCell).Caption
    ElseIf LastCell = 2 Then
      IndTagData(nTag).Title = "Variable Milling Cost"
    End If
    IndTagData(nTag).Units = labProcessingUnits(LastCell).Caption
  Else
    Tagged(hscSetNumbers.Value, LastCell + 95).Independent = nTag
    IndTagData(nTag).TheCell = LastCell + 95
    IndTagData(nTag).Title = labProcessingTitles(LastCell + 1).Caption
    IndTagData(nTag).Units = labProcessingUnits(LastCell + 1).Caption
  End If
  IndTagData(nTag).SetNumber = hscSetNumbers.Value
End If

txtProcessingValues(LastCell).SetFocus

End Sub


Private Sub Command1_Click()

If labCheckTag(LastCell).Visible = False Then
  dTag = dTag + 1
  labCheckTag(LastCell).Visible = True
  labCheckTag(LastCell).ForeColor = &HFFFF&
  labCheckTag(LastCell).Caption = LTrim(Str(nTag))
  If LastCell < 6 Then
    Tagged(hscSetNumbers.Value, LastCell + 92).Dependent = nTag
    DepTagData(nTag, dTag).Title = "Milling - " & labProcessingTitles(LastCell).Caption
    If LastCell = 0 Then
      DepTagData(nTag, dTag).Title = "Mill " & labProcessingTitles(LastCell).Caption
    ElseIf LastCell = 2 Then
      DepTagData(nTag, dTag).Title = "Variable Milling Cost"
    End If
    DepTagData(nTag, dTag).Units = labProcessingUnits(LastCell).Caption
  Else
    Tagged(hscSetNumbers.Value, LastCell + 95).Dependent = nTag
    DepTagData(nTag, dTag).Title = labProcessingTitles(LastCell + 1).Caption
    DepTagData(nTag, dTag).Units = labProcessingUnits(LastCell + 1).Caption
  End If
  DepTagData(nTag, dTag).SetNumber = hscSetNumbers.Value
End If

End Sub


Private Sub Form_Activate()

Dim baseunit As String
Dim baselength As Integer
Dim i As Integer

If IsHelpOn = True Then
  If LastCell = 100 Then
    txtProSetLabel.SetFocus
  Else
    txtProcessingValues(LastCell).SetFocus
  End If
  IsHelpOn = False
Else
  baseunit = LTrim(RTrim(CommodityData(1, 0).reserves))
  If baseunit = "" Then baseunit = "tons"
  baselength = Len(baseunit) - 1
  baseunit = Left(baseunit, baselength)

  DoNotChange = True
  CalcProcessScreen (1)
  labProcessingUnits(0).Caption = baseunit & "s/day"
  For i = 2 To 6
    labProcessingUnits(i).Caption = "/" & baseunit & " ore"
  Next i
  For i = 7 To 8
    labProcessingUnits(i).Caption = "/" & baseunit & " concentrate"
  Next i
  For i = 10 To 11
    labProcessingUnits(i).Caption = "/" & baseunit & " ore"
  Next i
  Call drawthevalues
  ShowMenu = True
  DoNotChange = False

  hscSetNumbers.Value = 1
  txtProSetLabel.Text = Pn1(5, hscSetNumbers.Value)
  
  If InsertFlag = True Then
    labInsert.Caption = "Insert"
  Else
    labInsert.Caption = "Typeover"
  End If
  
  LastCell = 0
  txtProcessingValues(0).SetFocus

End If

End Sub

Private Sub Form_Deactivate()
If ShowMenu = True Then
  frmProcessingCost.Hide
  Call InputMenuAccess(1)
End If
End Sub

Private Sub Form_Load()

Dim i As Integer

If FullScreen = False Then
  frmProcessingCost.Top = (Screen.Height - (frmProcessingCost.Height + 350)) / 2
  frmProcessingCost.Left = (Screen.Width - frmProcessingCost.Width) / 2
Else
  frmProcessingCost.Top = 0
  frmProcessingCost.Left = 0
  frmProcessingCost.WindowState = 2
End If

If frmProcessingCost.Top < 0 Then frmProcessingCost.Top = 0
If frmProcessingCost.Left < 0 Then frmProcessingCost.Left = 0

tempwide = frmProcessingCost.ScaleWidth
temphigh = frmProcessingCost.ScaleHeight

DoNotChange = True

If PageChange(3) = True Then
  Call drawthevalues
End If

DoNotChange = False

Call screenstuff

End Sub



Private Sub Form_Resize()

tempwide = frmProcessingCost.ScaleWidth
temphigh = frmProcessingCost.ScaleHeight

Call screenstuff

End Sub

Private Sub Form_Unload(Cancel As Integer)

  frmProcessingCost.Hide
  If ShowMenu = True Then Call InputMenuAccess(1)

End Sub

Private Sub hscSetNumbers_Change()

labSetNumbers.Caption = LTrim(RTrim(Str(hscSetNumbers.Value)))
 
txtProSetLabel.Text = Pn1(5, hscSetNumbers.Value)

If hscSetNumbers.Value > Np(5) Then
  Np(5) = hscSetNumbers.Value
End If

If Np(5) > Npna Then Npna = Np(5)

Call CalcProcessScreen(0)

Call drawthevalues

txtProcessingValues(0).SetFocus

End Sub

Private Sub imgBackToMenu_Click()

  frmProcessingCost.Hide
  If ShowMenu = True Then Call InputMenuAccess(1)

End Sub

Private Sub labBackToMenu_Click()

  frmProcessingCost.Hide
  If ShowMenu = True Then Call InputMenuAccess(1)

End Sub


Private Sub labProcessHelp_Click()

Dim begin As Integer
Dim sendindex As Integer

begin = 0
ShowMenu = False
WhichScreen = 3
If LastCell < 6 Then
  sendindex = LastCell + 92
ElseIf LastCell = 100 Then
  sendindex = 20
  WhichScreen = 0
Else
  sendindex = LastCell + 95
End If

Call frmApexHelp.gethelptext(sendindex, begin)
frmApexHelp.Show

End Sub

Private Sub txtProcessingValues_Change(Index As Integer)

If DoNotChange = True Then Exit Sub

PageChange(3) = True

If labCheckTag(Index).Visible = True Then ParamSet = False

If Index < 6 Then
  Primary(hscSetNumbers.Value, Index + 92) = CCur(Val(txtProcessingValues(Index).Text))
Else
  Primary(hscSetNumbers.Value, Index + 95) = CCur(Val(txtProcessingValues(Index).Text))
End If

If Index = 0 Then DidWeChange(1) = True

Call CalcProcessScreen(Index)

End Sub



Public Sub CalcProcessScreen(Index)

Dim starteq As Currency
Dim i As Integer

Select Case Index
  Case 1 To 5
    For i = 1 To 5
      labProcessingTitles(i).Enabled = True
      labProcessingLabels(i + 3).Enabled = True
      txtProcessingValues(i).Enabled = True
      labProcessingTitles(i).Enabled = True
    Next i
    If (Val(txtProcessingValues(1).Text) + Val(txtProcessingValues(2).Text)) > 0 Then
      For i = 3 To 5
        labProcessingTitles(i).Enabled = False
        labProcessingLabels(i + 3).Enabled = False
        txtProcessingValues(i).Enabled = False
        labProcessingUnits(i).Enabled = False
      Next i
    ElseIf (Val(txtProcessingValues(3).Text) + Val(txtProcessingValues(4).Text) + Val(txtProcessingValues(5).Text)) > 0 Then
      For i = 1 To 2
        labProcessingTitles(i).Enabled = False
        labProcessingLabels(i + 3).Enabled = False
        txtProcessingValues(i).Enabled = False
        labProcessingUnits(i).Enabled = False
      Next i
    End If
End Select

If Primary(hscSetNumbers.Value, 92) = 0 Then
  starteq = 0
Else
  starteq = Primary(hscSetNumbers.Value, 93) / Primary(hscSetNumbers.Value, 92)
End If

Primary(hscSetNumbers.Value, 98) = starteq + Primary(hscSetNumbers.Value, 94) + Primary(hscSetNumbers.Value, 95) + Primary(hscSetNumbers.Value, 96) + Primary(hscSetNumbers.Value, 97)

labScreenTotals(0).Caption = Str(RTrim(Primary(hscSetNumbers.Value, 98)))

If Primary(hscSetNumbers.Value, 103) = 0 Then Primary(hscSetNumbers.Value, 103) = 1

Primary(hscSetNumbers.Value, 104) = (Primary(hscSetNumbers.Value, 101) + Primary(hscSetNumbers.Value, 102)) / Primary(hscSetNumbers.Value, 103)

labScreenTotals(1).Caption = Str(RTrim(Primary(hscSetNumbers.Value, 104)))

labScreenTotals(2).Caption = Str(RTrim(Primary(hscSetNumbers.Value, 98) + Primary(hscSetNumbers.Value, 104)))

labScreenTotals(0).Caption = Format(labScreenTotals(0).Caption, "##,##0.00")
labScreenTotals(1).Caption = Format(labScreenTotals(1).Caption, "##,##0.00")
labScreenTotals(2).Caption = Format(labScreenTotals(2).Caption, "##,##0.00")

End Sub

Private Sub txtProcessingValues_GotFocus(Index As Integer)

LastCell = Index

End Sub



Public Sub screenstuff()

  Dim X As Integer
  Dim Y As Currency
  
  labProcessingHeading.Top = temphigh * 0.0187
  labProcessingHeading.Left = tempwide * 0.0131
  
  LineLeft.X1 = tempwide * 0.2164
  LineLeft.X2 = tempwide * 0.2164
  LineLeft.Y1 = temphigh * 0.1402
  LineLeft.Y2 = temphigh * 0.8972

  LineTop.X1 = tempwide * 0.2098
  LineTop.X2 = tempwide * 0.9508
  LineTop.Y1 = temphigh * 0.1495
  LineTop.Y2 = temphigh * 0.1495
  
  lineMiddle.X1 = tempwide * 0.223
  lineMiddle.X2 = tempwide * 0.9377
  lineMiddle.Y1 = temphigh * 0.5514
  lineMiddle.Y2 = temphigh * 0.5514
  
  LineBottom.X1 = tempwide * 0.2098
  LineBottom.X2 = tempwide * 0.9508
  LineBottom.Y1 = temphigh * 0.8879
  LineBottom.Y2 = temphigh * 0.8879
  
  LineRight.X1 = tempwide * 0.9443
  LineRight.X2 = tempwide * 0.9443
  LineRight.Y1 = temphigh * 0.1402
  LineRight.Y2 = temphigh * 0.8972
  
  For X = 0 To 11
    If X < 7 Then
      Y = 0
    ElseIf X < 11 Then
      Y = 0.0748
    Else
      Y = 0.1122
    End If
    labProcessingTitles(X).Top = (Y * temphigh) + (X * 0.0467 * temphigh) + (temphigh * 0.2056)
    labProcessingTitles(X).Left = tempwide * 0.2295
    labProcessingTitles(X).Width = tempwide * 0.2672
    labProcessingUnits(X).Top = (Y * temphigh) + (X * 0.0467 * temphigh) + (temphigh * 0.2056)
    labProcessingUnits(X).Left = tempwide * 0.6623
  Next X

  For X = 0 To 8
    If X < 6 Then
      Y = 0
    Else
      Y = 0.1215
    End If
    txtProcessingValues(X).Top = (Y * temphigh) + (X * 0.0467 * temphigh) + (temphigh * 0.1986)
    txtProcessingValues(X).Left = tempwide * 0.5377
    txtProcessingValues(X).Width = tempwide * 0.1197
    labCheckTag(X).Top = (Y * temphigh) + (X * 0.0467 * temphigh) + (temphigh * 0.2056)
    labCheckTag(X).Left = tempwide * 0.8852
    labCheckTag(X).Width = tempwide * 0.0475
  Next X
  
  For X = 0 To 2
    If X = 0 Then
      labScreenTotals(X).Top = 0.486 * temphigh
    ElseIf X = 1 Then
      labScreenTotals(X).Top = 0.7477 * temphigh
    Else
      labScreenTotals(X).Top = 0.8318 * temphigh
    End If
    labScreenTotals(X).Left = tempwide * 0.5443
    labScreenTotals(X).Width = tempwide * 0.1066
  Next X
  
  For X = 0 To 1
    If X = 0 Then
      labProcessingLabels(X).Top = temphigh * 0.1308
    Else
      labProcessingLabels(X).Top = temphigh * 0.5421
    End If
    labProcessingLabels(X).Left = tempwide * 0.2033
  Next X
    
  labProcessingLabels(2).Top = temphigh * 0.1589
  labProcessingLabels(2).Left = tempwide * 0.8852
  labProcessingLabels(2).Width = tempwide * 0.0475
 
  labProcessingLabels(3).Top = temphigh * 0.4112
  labProcessingLabels(3).Left = tempwide * 0.0459
  labProcessingLabels(3).Width = tempwide * 0.1131
  
  labProcessingLabels(14).Top = temphigh * 0.5093
  labProcessingLabels(14).Left = tempwide * 0.0459
  labProcessingLabels(14).Width = tempwide * 0.1131
  
  For X = 4 To 13
    If X < 10 Then
      Y = 0
    ElseIf X < 12 Then
      Y = 0.0748
    ElseIf X < 13 Then
      Y = 0.1215
    Else
      Y = 0.1589
    End If
    labProcessingLabels(X).Top = (Y * temphigh) + ((X - 4) * 0.0467 * temphigh) + (temphigh * 0.2523)
    labProcessingLabels(X).Left = tempwide * 0.518
    labProcessingLabels(X).Width = tempwide * 0.0148
  Next X
  
  hscSetNumbers.Top = temphigh * 0.4579
  hscSetNumbers.Left = (tempwide * 0.0861) - 188

  labSetNumbers.Top = temphigh * 0.4532
  labSetNumbers.Left = tempwide * 0.118
  labSetNumbers.Width = tempwide * 0.0344
  
  txtProSetLabel.Top = temphigh * 0.5514
  txtProSetLabel.Left = tempwide * 0.0459
  txtProSetLabel.Width = tempwide * 0.1131
  
  comIndTag.Top = temphigh * 0.9081
  comIndTag.Left = tempwide * 0.223
  
  labIndTag.Top = temphigh * 0.9065
  labIndTag.Left = tempwide * 0.2557
  
  comDepTag.Top = temphigh * 0.9081
  comDepTag.Left = tempwide * 0.7672
  
  labDepTag.Top = temphigh * 0.9065
  labDepTag.Left = tempwide * 0.8
  
  labBackToMenu.Top = temphigh * 0.9532
  labBackToMenu.Left = tempwide * 0.0656

  imgBackToMenu.Top = temphigh * 0.9626
  imgBackToMenu.Left = tempwide * 0.0066
  imgBackToMenu.Width = tempwide * 0.0541

  labProcessHelp.Top = temphigh * 0.9532
  labProcessHelp.Left = tempwide * 0.9377

  labInsert.Top = temphigh * 0.9562
  labInsert.Left = tempwide * 0.5378
  labInsert.Width = tempwide * 0.1066

End Sub

Public Sub drawthevalues()

Dim i As Integer
Dim X As Integer

DoNotChange = True

For i = 0 To 8
  If i < 6 Then
    txtProcessingValues(i).Text = LTrim(Str(Primary(hscSetNumbers.Value, i + 92)))
  Else
    txtProcessingValues(i).Text = LTrim(Str(Primary(hscSetNumbers.Value, i + 95)))
  End If
  If i < 2 Then
    txtProcessingValues(i).Text = Format(txtProcessingValues(i).Text, "#########0")
  Else
    txtProcessingValues(i).Text = Format(txtProcessingValues(i).Text, "######0.00")
  End If
Next i

For i = 92 To 103
  If i < 98 Or i > 100 Then
    Select Case i
      Case 92 To 97
        X = i - 92
      Case 101 To 103
        X = i - 95
    End Select
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

Private Sub txtProcessingValues_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

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
        If InStr(txtProcessingValues(Index).Text, ".") = 0 Then
          SendKeys "{DELETE}", False
        End If
      Else
        SendKeys "{DELETE}", False
      End If
  End Select
End If

End Sub

Private Sub txtProcessingValues_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii = 46 Then
  If InStr(txtProcessingValues(Index).Text, ".") > 0 Then
    Beep
    KeyAscii = 0
  End If
End If

If KeyAscii = 44 Then
  Beep
  KeyAscii = 0
End If

End Sub


Private Sub txtProSetLabel_Change()

Pn1(5, hscSetNumbers.Value) = txtProSetLabel.Text

End Sub


Private Sub txtProSetLabel_GotFocus()

LastCell = 100

End Sub


