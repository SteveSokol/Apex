VERSION 5.00
Begin VB.Form frmSensitivity 
   BackColor       =   &H00000000&
   Caption         =   "Sensitivity Analysis"
   ClientHeight    =   6420
   ClientLeft      =   2190
   ClientTop       =   1545
   ClientWidth     =   9150
   FillColor       =   &H00C0C0C0&
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
   Begin VB.VScrollBar vscDependent 
      Height          =   975
      Left            =   6360
      Max             =   1
      Min             =   1
      TabIndex        =   98
      Top             =   840
      Value           =   1
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.HScrollBar hscTagNumber 
      Height          =   195
      Left            =   1200
      Max             =   50
      Min             =   1
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   960
      Value           =   1
      Width           =   375
   End
   Begin VB.Label labOutPV 
      BackColor       =   &H00000000&
      Caption         =   "00.00%"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4860
      TabIndex        =   99
      Top             =   2340
      Width           =   1095
   End
   Begin VB.Line linLastLine 
      BorderColor     =   &H00FFFF00&
      X1              =   3120
      X2              =   8880
      Y1              =   780
      Y2              =   780
   End
   Begin VB.Label labDepUnits 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   7920
      TabIndex        =   97
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label labDepUnits 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   7920
      TabIndex        =   96
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label labDepUnits 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   7920
      TabIndex        =   95
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label labDepUnits 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   7920
      TabIndex        =   94
      Top             =   840
      Width           =   975
   End
   Begin VB.Label labDepMax 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   6600
      TabIndex        =   93
      Top             =   1560
      Width           =   1275
   End
   Begin VB.Label labDepMax 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   6600
      TabIndex        =   92
      Top             =   1320
      Width           =   1275
   End
   Begin VB.Label labDepMax 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   6600
      TabIndex        =   91
      Top             =   1080
      Width           =   1275
   End
   Begin VB.Label labDepMax 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   6600
      TabIndex        =   90
      Top             =   840
      Width           =   1275
   End
   Begin VB.Label labDepMin 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   5040
      TabIndex        =   89
      Top             =   1560
      Width           =   1275
   End
   Begin VB.Label labDepMin 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   5040
      TabIndex        =   88
      Top             =   1320
      Width           =   1275
   End
   Begin VB.Label labDepMin 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   5040
      TabIndex        =   87
      Top             =   1080
      Width           =   1275
   End
   Begin VB.Label labDepMin 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   5040
      TabIndex        =   86
      Top             =   840
      Width           =   1275
   End
   Begin VB.Label labDepItem 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   3
      Left            =   3180
      TabIndex        =   85
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label labDepItem 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   3180
      TabIndex        =   84
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label labDepItem 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   3180
      TabIndex        =   83
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label labDepItem 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   3180
      TabIndex        =   82
      Top             =   840
      Width           =   1815
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
      Left            =   8220
      TabIndex        =   81
      Top             =   6120
      Width           =   600
   End
   Begin VB.Label labIndTitles 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Units"
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
      Index           =   4
      Left            =   7980
      TabIndex        =   80
      Top             =   180
      Width           =   855
   End
   Begin VB.Label labIndTitles 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Maximum Value"
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
      Left            =   6540
      TabIndex        =   79
      Top             =   180
      Width           =   1335
   End
   Begin VB.Label labIndTitles 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "to"
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
      Left            =   6360
      TabIndex        =   78
      Top             =   180
      Width           =   195
   End
   Begin VB.Label labIndTitles 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Minimum Value"
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
      Left            =   4980
      TabIndex        =   77
      Top             =   180
      Width           =   1395
   End
   Begin VB.Label labIndTitles 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Variable"
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
      Left            =   3480
      TabIndex        =   76
      Top             =   180
      Width           =   1215
   End
   Begin VB.Label labIndItem 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   7920
      TabIndex        =   75
      Top             =   540
      Width           =   975
   End
   Begin VB.Label labIndItem 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   6600
      TabIndex        =   74
      Top             =   540
      Width           =   1275
   End
   Begin VB.Label labIndItem 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   5040
      TabIndex        =   73
      Top             =   540
      Width           =   1275
   End
   Begin VB.Label labIndItem 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   3180
      TabIndex        =   72
      Top             =   540
      Width           =   1815
   End
   Begin VB.Line linParametersRight 
      BorderColor     =   &H00FFFF00&
      X1              =   8940
      X2              =   8940
      Y1              =   60
      Y2              =   1920
   End
   Begin VB.Line linParametersLeft 
      BorderColor     =   &H00FFFF00&
      X1              =   3060
      X2              =   3060
      Y1              =   60
      Y2              =   1920
   End
   Begin VB.Line linParametersBottom 
      BorderColor     =   &H00FFFF00&
      X1              =   3000
      X2              =   9000
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Line linParametersTop 
      BorderColor     =   &H00FFFF00&
      X1              =   3000
      X2              =   9000
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line linTagRight 
      BorderColor     =   &H00FFFF00&
      X1              =   2220
      X2              =   2220
      Y1              =   540
      Y2              =   1920
   End
   Begin VB.Line linTagLeft 
      BorderColor     =   &H00FFFF00&
      X1              =   720
      X2              =   720
      Y1              =   540
      Y2              =   1920
   End
   Begin VB.Line linTagBottom 
      BorderColor     =   &H00FFFF00&
      X1              =   660
      X2              =   2280
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Line linTagTop 
      BorderColor     =   &H00FFFF00&
      X1              =   660
      X2              =   2280
      Y1              =   600
      Y2              =   600
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
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   960
      TabIndex        =   71
      Top             =   1560
      Width           =   1035
   End
   Begin VB.Label labSetTitle 
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
      Left            =   960
      TabIndex        =   70
      Top             =   1320
      Width           =   1035
   End
   Begin VB.Label labTagNumber 
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
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1620
      TabIndex        =   69
      Top             =   960
      Width           =   255
   End
   Begin VB.Label LabTagTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Tag Number"
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
      Left            =   960
      TabIndex        =   67
      Top             =   660
      Width           =   1035
   End
   Begin VB.Line linSenseLeft 
      BorderColor     =   &H00FFFF00&
      X1              =   240
      X2              =   240
      Y1              =   1980
      Y2              =   6090
   End
   Begin VB.Line linSenseRight 
      BorderColor     =   &H00FFFF00&
      X1              =   8940
      X2              =   8940
      Y1              =   1980
      Y2              =   6090
   End
   Begin VB.Line linSenseBottom 
      BorderColor     =   &H00FFFF00&
      X1              =   180
      X2              =   9000
      Y1              =   6030
      Y2              =   6030
   End
   Begin VB.Line linSenseMiddle 
      BorderColor     =   &H00FFFF00&
      X1              =   300
      X2              =   8880
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line linSenseTop 
      BorderColor     =   &H00FFFF00&
      X1              =   180
      X2              =   9000
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Image imgBackToMenu 
      Height          =   195
      Left            =   60
      Picture         =   "frmSensitivity.frx":0000
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
      Left            =   600
      TabIndex        =   66
      Top             =   6120
      Width           =   675
   End
   Begin VB.Label labSenseUnits 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "(percent)"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   7560
      TabIndex        =   65
      Top             =   2340
      Width           =   1335
   End
   Begin VB.Label labSenseUnits 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "(years)"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   6180
      TabIndex        =   64
      Top             =   2340
      Width           =   1215
   End
   Begin VB.Label labSenseUnits 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "@"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   4140
      TabIndex        =   63
      Top             =   2340
      Width           =   615
   End
   Begin VB.Label labSenseUnits 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Sum of Cash Flows"
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
      Left            =   2100
      TabIndex        =   62
      Top             =   2340
      Width           =   1815
   End
   Begin VB.Label labSenseUnits 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   300
      TabIndex        =   61
      Top             =   2340
      Width           =   1575
   End
   Begin VB.Label labSenseTitles 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Internal ROR"
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
      Left            =   7560
      TabIndex        =   60
      Top             =   2100
      Width           =   1335
   End
   Begin VB.Label labSenseTitles 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Pay Back"
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
      Left            =   6180
      TabIndex        =   59
      Top             =   2100
      Width           =   1215
   End
   Begin VB.Label labSenseTitles 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Present Value"
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
      Left            =   4140
      TabIndex        =   58
      Top             =   2100
      Width           =   1815
   End
   Begin VB.Label labSenseTitles 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
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
      Left            =   2100
      TabIndex        =   57
      Top             =   2100
      Width           =   1815
   End
   Begin VB.Label labSenseTitles 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Value"
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
      Left            =   480
      TabIndex        =   56
      Top             =   2100
      Width           =   1215
   End
   Begin VB.Label labSenseHeading 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sensitivity Analysis"
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
      Left            =   60
      TabIndex        =   55
      Top             =   60
      Width           =   2850
   End
   Begin VB.Label labReturn 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Rate of Return"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   7620
      TabIndex        =   54
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label labReturn 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Rate of Return"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   7620
      TabIndex        =   53
      Top             =   5460
      Width           =   1215
   End
   Begin VB.Label labReturn 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Rate of Return"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   7620
      TabIndex        =   52
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label labReturn 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Rate of Return"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   7620
      TabIndex        =   51
      Top             =   4860
      Width           =   1215
   End
   Begin VB.Label labReturn 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Rate of Return"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   7620
      TabIndex        =   50
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label labReturn 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Rate of Return"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   7620
      TabIndex        =   49
      Top             =   4260
      Width           =   1215
   End
   Begin VB.Label labReturn 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Rate of Return"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   7620
      TabIndex        =   48
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label labReturn 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Rate of Return"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   7620
      TabIndex        =   47
      Top             =   3660
      Width           =   1215
   End
   Begin VB.Label labReturn 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Rate of Return"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   7620
      TabIndex        =   46
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label labReturn 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Rate of Return"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   7620
      TabIndex        =   45
      Top             =   3060
      Width           =   1215
   End
   Begin VB.Label labReturn 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Rate of Return"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   7620
      TabIndex        =   44
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label labPayBack 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Pay Back"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   6180
      TabIndex        =   43
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label labPayBack 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Pay Back"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   6180
      TabIndex        =   42
      Top             =   5460
      Width           =   1215
   End
   Begin VB.Label labPayBack 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Pay Back"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   6180
      TabIndex        =   41
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label labPayBack 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Pay Back"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   6180
      TabIndex        =   40
      Top             =   4860
      Width           =   1215
   End
   Begin VB.Label labPayBack 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Pay Back"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   6180
      TabIndex        =   39
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label labPayBack 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Pay Back"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   6180
      TabIndex        =   38
      Top             =   4260
      Width           =   1215
   End
   Begin VB.Label labPayBack 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Pay Back"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   6180
      TabIndex        =   37
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label labPayBack 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Pay Back"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   6180
      TabIndex        =   36
      Top             =   3660
      Width           =   1215
   End
   Begin VB.Label labPayBack 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Pay Back"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   6180
      TabIndex        =   35
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label labPayBack 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Pay Back"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   6180
      TabIndex        =   34
      Top             =   3060
      Width           =   1215
   End
   Begin VB.Label labPayBack 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Pay Back"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   6180
      TabIndex        =   33
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label labPresentValues 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Present Values"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   4140
      TabIndex        =   32
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label labPresentValues 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Present Values"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   4140
      TabIndex        =   31
      Top             =   5460
      Width           =   1815
   End
   Begin VB.Label labPresentValues 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Present Values"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   4140
      TabIndex        =   30
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label labPresentValues 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Present Values"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   4140
      TabIndex        =   29
      Top             =   4860
      Width           =   1815
   End
   Begin VB.Label labPresentValues 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Present Values"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   4140
      TabIndex        =   28
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label labPresentValues 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Present Values"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   4140
      TabIndex        =   27
      Top             =   4260
      Width           =   1815
   End
   Begin VB.Label labPresentValues 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Present Values"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   4140
      TabIndex        =   26
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label labPresentValues 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Present Values"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   4140
      TabIndex        =   25
      Top             =   3660
      Width           =   1815
   End
   Begin VB.Label labPresentValues 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Present Values"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   4140
      TabIndex        =   24
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label labPresentValues 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Present Values"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   4140
      TabIndex        =   23
      Top             =   3060
      Width           =   1815
   End
   Begin VB.Label labPresentValues 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Present Values"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   4140
      TabIndex        =   22
      Top             =   2760
      Width           =   1515
   End
   Begin VB.Label labCashFlows 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Cash Flows"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   2100
      TabIndex        =   21
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label labCashFlows 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Cash Flows"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   2100
      TabIndex        =   20
      Top             =   5460
      Width           =   1815
   End
   Begin VB.Label labCashFlows 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Cash Flows"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   2100
      TabIndex        =   19
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label labCashFlows 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Cash Flows"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   2100
      TabIndex        =   18
      Top             =   4860
      Width           =   1815
   End
   Begin VB.Label labCashFlows 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Cash Flows"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   2100
      TabIndex        =   17
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label labCashFlows 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Cash Flows"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   2100
      TabIndex        =   16
      Top             =   4260
      Width           =   1815
   End
   Begin VB.Label labCashFlows 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Cash Flows"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   2100
      TabIndex        =   15
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label labCashFlows 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Cash Flows"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   2100
      TabIndex        =   14
      Top             =   3660
      Width           =   1815
   End
   Begin VB.Label labCashFlows 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Cash Flows"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   2100
      TabIndex        =   13
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label labCashFlows 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Cash Flows"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   2100
      TabIndex        =   12
      Top             =   3060
      Width           =   1815
   End
   Begin VB.Label labCashFlows 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Cash Flows"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   2100
      TabIndex        =   11
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label labValues 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Values"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   480
      TabIndex        =   10
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label labValues 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Values"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   480
      TabIndex        =   9
      Top             =   5460
      Width           =   1215
   End
   Begin VB.Label labValues 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Values"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   480
      TabIndex        =   8
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label labValues 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Values"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   480
      TabIndex        =   7
      Top             =   4860
      Width           =   1215
   End
   Begin VB.Label labValues 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Values"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   480
      TabIndex        =   6
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label labValues 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Values"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   480
      TabIndex        =   5
      Top             =   4260
      Width           =   1215
   End
   Begin VB.Label labValues 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Values"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   4
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label labValues 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Values"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   3
      Top             =   3660
      Width           =   1215
   End
   Begin VB.Label labValues 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Values"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   2
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label labValues 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Values"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   3060
      Width           =   1215
   End
   Begin VB.Label labValues 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Values"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   2760
      Width           =   1095
   End
End
Attribute VB_Name = "frmSensitivity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim temphigh As Single
Dim tempwide As Single
Dim totaldeps As Integer
Private Sub Form_Activate()

DoNotChange = True
ShowMenu = True
hscTagNumber.Value = 1
labSetNumber.Caption = LTrim(RTrim(Str(IndTagData(hscTagNumber.Value).SetNumber)))

DoNotChange = False

Call gettheanalysis(hscTagNumber.Value, IndTagData(hscTagNumber.Value).SetNumber)

End Sub

Private Sub Form_Deactivate()
  
  If ShowMenu = True Then
    frmSensitivity.Hide
    Call InputMenuAccess(2)
  End If
 
End Sub

Private Sub Form_Load()

If FullScreen = False Then
  frmSensitivity.Top = (Screen.Height - (frmSensitivity.Height + 350)) / 2
  frmSensitivity.Left = (Screen.Width - frmSensitivity.Width) / 2
Else
  frmSensitivity.Top = 0
  frmSensitivity.Left = 0
  frmSensitivity.WindowState = 2
End If

If frmSensitivity.Top < 0 Then frmSensitivity.Top = 0
If frmSensitivity.Left < 0 Then frmSensitivity.Left = 0

tempwide = frmSensitivity.ScaleWidth
temphigh = frmSensitivity.ScaleHeight

Call screenstuff

End Sub

Private Sub Form_Resize()

tempwide = frmSensitivity.ScaleWidth
temphigh = frmSensitivity.ScaleHeight

Call screenstuff

End Sub


Private Sub Form_Unload(Cancel As Integer)

  frmSensitivity.Hide
  If ShowMenu = True Then Call InputMenuAccess(2)

End Sub

Private Sub hscTagNumber_Change()

Dim i As Integer

labTagNumber.Caption = LTrim(RTrim(Str(hscTagNumber.Value)))
labSetNumber.Caption = LTrim(RTrim(Str(IndTagData(hscTagNumber.Value).SetNumber)))
If DoNotChange = True Then Exit Sub

For i = 0 To 3
  labDepItem(i).Caption = ""
  labDepMin(i).Caption = ""
  labDepMax(i).Caption = ""
  labDepUnits(i).Caption = ""
Next i

Call gettheanalysis(hscTagNumber.Value, IndTagData(hscTagNumber.Value).SetNumber)

End Sub

Private Sub imgBackToMenu_Click()
  
  frmSensitivity.Hide
  Call InputMenuAccess(2)

End Sub

Private Sub labBackToMenu_Click()
  
  frmSensitivity.Hide
  Call InputMenuAccess(2)

End Sub

Public Sub gettheanalysis(thetag As Integer, theset As Integer)

Dim i As Integer
Dim j As Integer
Dim ii As Integer
Dim taggo As Integer
Dim nowvalue As String
Dim nowwhich As Integer
Dim depcount As Integer
Dim oldvalue(51) As Currency
Dim interval(51) As Currency

labIndItem(0).Caption = LTrim(RTrim(IndTagData(thetag).Title))
labIndItem(1).Caption = LTrim(RTrim(Str(IndTagData(thetag).Minimum)))
labIndItem(2).Caption = LTrim(RTrim(Str(IndTagData(thetag).Maximum)))
labIndItem(3).Caption = LTrim(RTrim(IndTagData(thetag).Units))
labSenseUnits(0).Caption = "(" & LTrim(RTrim(IndTagData(thetag).Units)) & ")"

labOutPV.Caption = Format(LTrim(RTrim(Str(Sets(25)))), "#0.00") & "%"

For i = 0 To 3
  labDepItem(i) = ""
  labDepMin(i) = ""
  labDepMax(i) = ""
  labDepUnits(i) = ""
Next i

For i = 1 To 180
  If i < 131 Then
    If Tagged(theset, i).Independent = thetag Then
      oldvalue(1) = Primary(theset, i)
      ii = i
    End If
  Else
    If Tagged(theset, i).Independent = thetag Then
      oldvalue(1) = CapitalData(i - 131).PurchaseAmount
      ii = i
    End If
  End If
Next i

nowwhich = ii
nowvalue = labIndItem(1).Caption
Call findaformat(nowwhich, nowvalue)
labIndItem(1).Caption = nowvalue
nowvalue = labIndItem(2).Caption
Call findaformat(nowwhich, nowvalue)
labIndItem(2).Caption = nowvalue
 
taggo = 0

For i = 1 To 50
  If DepTagData(thetag, i).TheCell <> 0 Then
    taggo = taggo + 1
    If DepTagData(thetag, i).TheCell < 131 Then
      oldvalue(taggo + 1) = Primary(DepTagData(thetag, i).SetNumber, DepTagData(thetag, i).TheCell)
    Else
      oldvalue(taggo + 1) = CapitalData(DepTagData(thetag, i).TheCell - 131).PurchaseAmount
    End If
    If taggo < 5 Then
    labDepItem(taggo - 1).Caption = LTrim(RTrim(DepTagData(thetag, i).Title))
    labDepUnits(taggo - 1).Caption = LTrim(RTrim(DepTagData(thetag, i).Units))
    nowwhich = DepTagData(thetag, i).TheCell
    nowvalue = LTrim(RTrim(Str(DepTagData(thetag, i).Minimum)))
    Call findaformat(nowwhich, nowvalue)
    labDepMin(taggo - 1).Caption = nowvalue
    nowvalue = LTrim(RTrim(Str(DepTagData(thetag, i).Maximum)))
    Call findaformat(nowwhich, nowvalue)
    labDepMax(taggo - 1).Caption = nowvalue
    End If
  End If
Next i

If taggo > 4 Then
  vscDependent.Visible = True
  vscDependent.max = (taggo - 3)
Else
  vscDependent.Visible = False
End If

interval(1) = (IndTagData(thetag).Maximum - IndTagData(thetag).Minimum) / 10
For i = 1 To taggo
  interval(i + 1) = (DepTagData(thetag, i).Maximum - DepTagData(thetag, i).Minimum) / 10
Next i

If ii < 131 Then
  Primary(theset, ii) = IndTagData(thetag).Minimum
  nowvalue = LTrim(RTrim(Str(Primary(theset, ii))))
  nowwhich = ii
  Call findaformat(nowwhich, nowvalue)
Else
  CapitalData(ii - 131).PurchaseAmount = IndTagData(thetag).Minimum
  nowvalue = Format(LTrim(RTrim(Str(CapitalData(ii - 131).PurchaseAmount))), "$##,###,###,###")
End If
  
For i = 1 To taggo
  If DepTagData(thetag, i).TheCell < 131 Then
    Primary(DepTagData(thetag, i).SetNumber, DepTagData(thetag, i).TheCell) = DepTagData(thetag, i).Minimum
  Else
    CapitalData(DepTagData(thetag, i).TheCell - 131).PurchaseAmount = DepTagData(thetag, i).Minimum
  End If
Next i

For i = 0 To 10
  Call cflow5(1, 0)
  Call rateofreturn
  labValues(i).Caption = nowvalue
  labCashFlows(i) = Format(LTrim(RTrim(Str(Pv0))), "$###,###,###,###")
  labPresentValues(i) = Format(LTrim(RTrim(Str(Pv2))), "$###,###,###,###")
  labPayBack(i) = Format(LTrim(RTrim(Str(Pb))), "###0.00")
  labReturn(i) = Format(LTrim(RTrim(Str(Rot * 100))), "#0.00")
  If ii < 131 Then
    Primary(theset, ii) = Primary(theset, ii) + interval(1)
    nowvalue = LTrim(RTrim(Str(Primary(theset, ii))))
    nowwhich = ii
    Call findaformat(nowwhich, nowvalue)
  Else
    CapitalData(ii - 131).PurchaseAmount = CapitalData(ii - 131).PurchaseAmount + interval(1)
    nowvalue = Format(LTrim(RTrim(Str(CapitalData(ii - 131).PurchaseAmount))), "$#,###,###,###")
  End If
  For j = 1 To taggo
    If DepTagData(thetag, j).TheCell < 131 Then
      Primary(DepTagData(thetag, j).SetNumber, DepTagData(thetag, j).TheCell) = Primary(DepTagData(thetag, j).SetNumber, DepTagData(thetag, j).TheCell) + interval(j + 1)
    Else
      CapitalData(DepTagData(thetag, j).TheCell - 131).PurchaseAmount = CapitalData(DepTagData(thetag, j).TheCell - 131).PurchaseAmount + interval(j + 1)
    End If
  Next j
Next i

'Clean-up

If ii < 131 Then
  Primary(theset, ii) = oldvalue(1)
Else
  CapitalData(ii - 131).PurchaseAmount = oldvalue(1)
End If

For i = 1 To taggo
  If DepTagData(thetag, i).TheCell < 131 Then
    Primary(DepTagData(thetag, i).SetNumber, DepTagData(thetag, i).TheCell) = oldvalue(i + 1)
  Else
    CapitalData(DepTagData(thetag, i).TheCell - 131).PurchaseAmount = oldvalue(i + 1)
  End If
Next i

Erase oldvalue
Erase interval

End Sub

Private Sub labPrintScreen_Click()
  
  ShowMenu = False
  job = 11
  SendTag = hscTagNumber.Value
  Call printstuffout(job)

End Sub

Private Sub vscDependent_Change()

Dim i As Integer
Dim j As Integer
Dim k As Integer

For i = 1 To 50
  If DepTagData(hscTagNumber, i).TheCell <> 0 Then
    If j > (vscDependent.Value - 2) And j < (vscDependent.Value + 3) Then
      labDepItem(k).Caption = DepTagData(hscTagNumber.Value, i).Title
      labDepMin(k).Caption = LTrim(RTrim(Str(DepTagData(hscTagNumber.Value, i).Minimum)))
      labDepMax(k).Caption = LTrim(RTrim(Str(DepTagData(hscTagNumber.Value, i).Maximum)))
      labDepUnits(k).Caption = DepTagData(hscTagNumber.Value, i).Units
      k = k + 1
    End If
    j = j + 1
  End If
Next i

End Sub



Public Sub screenstuff()
  
  Dim X As Integer
   
  labSenseHeading.Top = temphigh * 0.0093
  labSenseHeading.Left = tempwide * 0.0066
  
  linTagTop.X1 = tempwide * 0.0721
  linTagTop.X2 = tempwide * 0.2492
  linTagTop.Y1 = temphigh * 0.0935
  linTagTop.Y2 = temphigh * 0.0935
  
  linTagLeft.X1 = tempwide * 0.0787
  linTagLeft.X2 = tempwide * 0.0787
  linTagLeft.Y1 = temphigh * 0.0841
  linTagLeft.Y2 = temphigh * 0.2991

  linTagRight.X1 = tempwide * 0.2426
  linTagRight.X2 = tempwide * 0.2426
  linTagRight.Y1 = temphigh * 0.0841
  linTagRight.Y2 = temphigh * 0.2991

  linTagBottom.X1 = tempwide * 0.0721
  linTagBottom.X2 = tempwide * 0.2492
  linTagBottom.Y1 = temphigh * 0.2897
  linTagBottom.Y2 = temphigh * 0.2897

  linParametersTop.X1 = tempwide * 0.3279
  linParametersTop.X2 = tempwide * 0.9836
  linParametersTop.Y1 = temphigh * 0.0187
  linParametersTop.Y2 = temphigh * 0.0187
  
  linParametersLeft.X1 = tempwide * 0.3344
  linParametersLeft.X2 = tempwide * 0.3344
  linParametersLeft.Y1 = temphigh * 0.0093
  linParametersLeft.Y2 = temphigh * 0.2991

  linParametersRight.X1 = tempwide * 0.977
  linParametersRight.X2 = tempwide * 0.977
  linParametersRight.Y1 = temphigh * 0.0093
  linParametersRight.Y2 = temphigh * 0.2991

  linParametersBottom.X1 = tempwide * 0.3279
  linParametersBottom.X2 = tempwide * 0.9836
  linParametersBottom.Y1 = temphigh * 0.2897
  linParametersBottom.Y2 = temphigh * 0.2897
  
  linLastLine.X1 = tempwide * 0.341
  linLastLine.X2 = tempwide * 0.9705
  linLastLine.Y1 = temphigh * 0.1215
  linLastLine.Y2 = temphigh * 0.1215
  
  linSenseTop.X1 = tempwide * 0.0197
  linSenseTop.X2 = tempwide * 0.9836
  linSenseTop.Y1 = temphigh * 0.3178
  linSenseTop.Y2 = temphigh * 0.3178
  
  linSenseLeft.X1 = tempwide * 0.0262
  linSenseLeft.X2 = tempwide * 0.0262
  linSenseLeft.Y1 = temphigh * 0.3084
  linSenseLeft.Y2 = temphigh * 0.9486

  linSenseMiddle.X1 = tempwide * 0.0328
  linSenseMiddle.X2 = tempwide * 0.9705
  linSenseMiddle.Y1 = temphigh * 0.4112
  linSenseMiddle.Y2 = temphigh * 0.4112

  linSenseRight.X1 = tempwide * 0.977
  linSenseRight.X2 = tempwide * 0.977
  linSenseRight.Y1 = temphigh * 0.3084
  linSenseRight.Y2 = temphigh * 0.9486

  linSenseBottom.X1 = tempwide * 0.0197
  linSenseBottom.X2 = tempwide * 0.9836
  linSenseBottom.Y1 = temphigh * 0.9392
  linSenseBottom.Y2 = temphigh * 0.9392

  For X = 0 To 4
    labSenseTitles(X).Top = temphigh * 0.3271
    labSenseUnits(X).Top = temphigh * 0.3645
    labIndTitles(X).Top = temphigh * 0.028
    If X = 0 Then
      labSenseTitles(X).Left = tempwide * 0.0525
      labSenseTitles(X).Width = tempwide * 0.1328
      labSenseUnits(X).Left = tempwide * 0.0328
      labSenseUnits(X).Width = tempwide * 0.1721
      labIndTitles(X).Left = tempwide * 0.3803
      labIndTitles(X).Width = tempwide * 0.1328
    ElseIf X = 1 Then
      labSenseTitles(X).Left = tempwide * 0.2295
      labSenseTitles(X).Width = tempwide * 0.1984
      labSenseUnits(X).Left = tempwide * 0.2295
      labSenseUnits(X).Width = tempwide * 0.1984
      labIndTitles(X).Left = tempwide * 0.5443
      labIndTitles(X).Width = tempwide * 0.1525
    ElseIf X = 2 Then
      labSenseTitles(X).Left = tempwide * 0.4525
      labSenseTitles(X).Width = tempwide * 0.1984
      labSenseUnits(X).Left = tempwide * 0.4525
      labSenseUnits(X).Width = tempwide * 0.0672
      labIndTitles(X).Left = tempwide * 0.6951
      labIndTitles(X).Width = tempwide * 0.0213
    ElseIf X = 3 Then
      labSenseTitles(X).Left = tempwide * 0.6754
      labSenseTitles(X).Width = tempwide * 0.1328
      labSenseUnits(X).Left = tempwide * 0.6754
      labSenseUnits(X).Width = tempwide * 0.1328
      labIndTitles(X).Left = tempwide * 0.7148
      labIndTitles(X).Width = tempwide * 0.1459
    ElseIf X = 4 Then
      labSenseTitles(X).Left = tempwide * 0.8262
      labSenseTitles(X).Width = tempwide * 0.1459
      labSenseUnits(X).Left = tempwide * 0.8262
      labSenseUnits(X).Width = tempwide * 0.1459
      labIndTitles(X).Left = tempwide * 0.8721
      labIndTitles(X).Width = tempwide * 0.0934
    End If
  Next X
  
  labOutPV.Top = temphigh * 0.3645
  labOutPV.Left = tempwide * 0.5311
  labOutPV.Width = tempwide * 0.1197
  
  For X = 0 To 3
    labIndItem(X).Top = temphigh * 0.0841
    If X = 0 Then
      labIndItem(X).Left = tempwide * 0.3475
      labIndItem(X).Width = tempwide * 0.1984
    ElseIf X = 1 Then
      labIndItem(X).Left = tempwide * 0.5508
      labIndItem(X).Width = tempwide * 0.1393
    ElseIf X = 2 Then
      labIndItem(X).Left = tempwide * 0.7213
      labIndItem(X).Width = tempwide * 0.1393
    ElseIf X = 3 Then
      labIndItem(X).Left = tempwide * 0.8656
      labIndItem(X).Width = tempwide * 0.1066
    End If
  Next X
  
  For X = 0 To 3
    labDepItem(X).Top = (X * 0.0374 * temphigh) + (temphigh * 0.1308)
    labDepMin(X).Top = (X * 0.0374 * temphigh) + (temphigh * 0.1308)
    labDepMax(X).Top = (X * 0.0374 * temphigh) + (temphigh * 0.1308)
    labDepUnits(X).Top = (X * 0.0374 * temphigh) + (temphigh * 0.1308)
    labDepItem(X).Left = tempwide * 0.3475
    labDepMin(X).Left = tempwide * 0.5508
    labDepMax(X).Left = tempwide * 0.7213
    labDepUnits(X).Left = tempwide * 0.8656
    labDepItem(X).Width = tempwide * 0.1984
    labDepMin(X).Width = tempwide * 0.1393
    labDepMax(X).Width = tempwide * 0.1393
    labDepUnits(X).Width = tempwide * 0.1066
  Next X
    
  vscDependent.Top = temphigh * 0.1308
  vscDependent.Height = temphigh * 0.1519
  vscDependent.Left = (tempwide * 0.7058) - 98
    
  For X = 0 To 10
    labValues(X).Top = (X * 0.0467 * temphigh) + (temphigh * 0.4299)
    labValues(X).Left = tempwide * 0.0525
    labValues(X).Width = tempwide * 0.1197
    labCashFlows(X).Top = (X * 0.0467 * temphigh) + (temphigh * 0.4299)
    labCashFlows(X).Left = tempwide * 0.2295
    labCashFlows(X).Width = tempwide * 0.159
    labPresentValues(X).Top = (X * 0.0467 * temphigh) + (temphigh * 0.4299)
    labPresentValues(X).Left = tempwide * 0.4525
    labPresentValues(X).Width = tempwide * 0.1656
    labPayBack(X).Top = (X * 0.0467 * temphigh) + (temphigh * 0.4299)
    labPayBack(X).Left = tempwide * 0.6754
    labPayBack(X).Width = tempwide * 0.0934
    labReturn(X).Top = (X * 0.0467 * temphigh) + (temphigh * 0.4299)
    labReturn(X).Left = tempwide * 0.8262
    labReturn(X).Width = tempwide * 0.1066
  Next X
  
  labTagTitle.Top = temphigh * 0.1028
  labTagTitle.Left = tempwide * 0.1049
  labTagTitle.Width = tempwide * 0.1131
  
  hscTagNumber.Top = temphigh * 0.1495
  hscTagNumber.Left = (tempwide * 0.1512) - 188
  
  labTagNumber.Top = temphigh * 0.1464
  labTagNumber.Left = tempwide * 0.177
  labTagNumber.Width = tempwide * 0.0279
  
  labSetTitle.Top = temphigh * 0.2056
  labSetTitle.Left = tempwide * 0.1049
  labSetTitle.Width = tempwide * 0.1131
  
  labSetNumber.Top = temphigh * 0.243
  labSetNumber.Left = tempwide * 0.1049
  labSetNumber.Width = tempwide * 0.1131
  
  labBackToMenu.Top = temphigh * 0.9532
  labBackToMenu.Left = tempwide * 0.0656

  imgBackToMenu.Top = temphigh * 0.9626
  imgBackToMenu.Left = tempwide * 0.0066
  imgBackToMenu.Width = tempwide * 0.0541

  labPrintScreen.Top = temphigh * 0.9532
  labPrintScreen.Left = tempwide * 0.8984

End Sub
