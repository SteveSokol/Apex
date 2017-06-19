VERSION 5.00
Begin VB.Form frmParameters 
   BackColor       =   &H00000000&
   Caption         =   "Analysis Parameters"
   ClientHeight    =   6420
   ClientLeft      =   375
   ClientTop       =   1680
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
   Begin VB.TextBox txtParametersValues 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   4
      Left            =   7800
      TabIndex        =   58
      Text            =   "1,000"
      Top             =   5460
      Width           =   915
   End
   Begin VB.TextBox txtRangePercent 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   1
      Left            =   2760
      TabIndex        =   38
      Text            =   "25.00"
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox txtRangePercent 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   0
      Left            =   2760
      TabIndex        =   37
      Text            =   "-25.00"
      Top             =   4680
      Width           =   735
   End
   Begin VB.HScrollBar hscTags 
      Height          =   195
      Left            =   600
      Max             =   25
      Min             =   1
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4020
      Value           =   1
      Width           =   375
   End
   Begin VB.OptionButton optDistribution 
      BackColor       =   &H00000000&
      Caption         =   "Skewed"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   7620
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   4560
      Width           =   1275
   End
   Begin VB.OptionButton optDistribution 
      BackColor       =   &H00000000&
      Caption         =   "Log Normal"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   7620
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   4200
      Width           =   1275
   End
   Begin VB.OptionButton optDistribution 
      BackColor       =   &H00000000&
      Caption         =   "Uniform"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   7620
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1275
   End
   Begin VB.OptionButton optDistribution 
      BackColor       =   &H00000000&
      Caption         =   "Normal"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   7620
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   3480
      Value           =   -1  'True
      Width           =   1275
   End
   Begin VB.TextBox txtParametersValues 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   3
      Left            =   3240
      TabIndex        =   12
      Text            =   "20.00"
      Top             =   2100
      Width           =   795
   End
   Begin VB.TextBox txtParametersValues 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   2
      Left            =   3240
      TabIndex        =   11
      Text            =   "15.00"
      Top             =   1740
      Width           =   795
   End
   Begin VB.TextBox txtParametersValues 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   1
      Left            =   3240
      TabIndex        =   10
      Text            =   "10.00"
      Top             =   1380
      Width           =   795
   End
   Begin VB.TextBox txtParametersValues 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   0
      Left            =   3240
      TabIndex        =   0
      Text            =   "1"
      Top             =   1020
      Width           =   795
   End
   Begin VB.Line LinIteration 
      BorderColor     =   &H00FFFF00&
      X1              =   7560
      X2              =   8940
      Y1              =   4980
      Y2              =   4980
   End
   Begin VB.Label labParametersMisc 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Iterations"
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
      Index           =   14
      Left            =   7740
      TabIndex        =   57
      Top             =   5100
      Width           =   1035
   End
   Begin VB.Label labParametersHelp 
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
      TabIndex        =   56
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
      Left            =   3060
      TabIndex        =   55
      Top             =   6120
      Width           =   975
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
      Left            =   5820
      TabIndex        =   54
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label labParameters 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   1980
      TabIndex        =   53
      Top             =   3600
      Width           =   2475
   End
   Begin VB.Label labParametersHeading 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Analysis Parameters"
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
      Left            =   180
      TabIndex        =   52
      Top             =   180
      Width           =   3135
   End
   Begin VB.Label labParametersMisc 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Distribution"
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
      Left            =   7740
      TabIndex        =   51
      Top             =   3120
      Width           =   1035
   End
   Begin VB.Line linMidBoxRight 
      BorderColor     =   &H00FFFF00&
      X1              =   7500
      X2              =   7500
      Y1              =   3000
      Y2              =   5880
   End
   Begin VB.Line linMidBoxMiddle 
      BorderColor     =   &H00FFFF00&
      X1              =   1560
      X2              =   7440
      Y1              =   4140
      Y2              =   4140
   End
   Begin VB.Line linMidBoxLeft 
      BorderColor     =   &H00FFFF00&
      X1              =   1500
      X2              =   1500
      Y1              =   3180
      Y2              =   5880
   End
   Begin VB.Label labParameters 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Index           =   8
      Left            =   300
      TabIndex        =   50
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label labParameters 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Index           =   7
      Left            =   1020
      TabIndex        =   45
      Top             =   4020
      Width           =   105
   End
   Begin VB.Label labParameters 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   6060
      TabIndex        =   44
      Top             =   5340
      Width           =   1335
   End
   Begin VB.Label labParameters 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   6060
      TabIndex        =   43
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label labParameters 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   4380
      TabIndex        =   42
      Top             =   5340
      Width           =   1635
   End
   Begin VB.Label labParameters 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   4380
      TabIndex        =   41
      Top             =   4800
      Width           =   1635
   End
   Begin VB.Label labParameters 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   6060
      TabIndex        =   40
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label labParameters 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   4500
      TabIndex        =   39
      Top             =   3600
      Width           =   1515
   End
   Begin VB.Label labParametersMisc 
      BackColor       =   &H00000000&
      Caption         =   "%"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   12
      Left            =   3540
      TabIndex        =   36
      Top             =   5340
      Width           =   195
   End
   Begin VB.Label labParametersMisc 
      BackColor       =   &H00000000&
      Caption         =   "%"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   11
      Left            =   3540
      TabIndex        =   35
      Top             =   4740
      Width           =   195
   End
   Begin VB.Label labParametersMisc 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Range of Values"
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
      Left            =   5280
      TabIndex        =   34
      Top             =   4320
      Width           =   1395
   End
   Begin VB.Label labParametersMisc 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Percent Change"
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
      Index           =   9
      Left            =   2520
      TabIndex        =   33
      Top             =   4320
      Width           =   1395
   End
   Begin VB.Label labParametersMisc 
      BackColor       =   &H00000000&
      Caption         =   "Maximum:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   1620
      TabIndex        =   32
      Top             =   5340
      Width           =   975
   End
   Begin VB.Label labParametersMisc 
      BackColor       =   &H00000000&
      Caption         =   "Minimum:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   1620
      TabIndex        =   31
      Top             =   4740
      Width           =   975
   End
   Begin VB.Label labParametersMisc 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Base Value"
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
      Left            =   5460
      TabIndex        =   30
      Top             =   3180
      Width           =   1035
   End
   Begin VB.Label labParametersMisc 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Tagged Item"
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
      Left            =   2700
      TabIndex        =   29
      Top             =   3180
      Width           =   1035
   End
   Begin VB.Label labParametersMisc 
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
      Index           =   4
      Left            =   300
      TabIndex        =   28
      Top             =   4740
      Width           =   1095
   End
   Begin VB.Label labParametersMisc 
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
      Index           =   3
      Left            =   300
      TabIndex        =   27
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Image imgBackToMenu 
      Height          =   195
      Left            =   60
      Picture         =   "frmParameters.frx":0000
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
      TabIndex        =   25
      Top             =   6120
      Width           =   675
   End
   Begin VB.Label labParametersMisc 
      BackColor       =   &H00000000&
      Caption         =   "Tagged Variables"
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
      Left            =   60
      TabIndex        =   24
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label labParametersMisc 
      BackColor       =   &H00000000&
      Caption         =   "Number of Data Sets"
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
      Left            =   6060
      TabIndex        =   23
      Top             =   240
      Width           =   1875
   End
   Begin VB.Label labParametersMisc 
      BackColor       =   &H00000000&
      Caption         =   "Net Present Value"
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
      Left            =   660
      TabIndex        =   22
      Top             =   720
      Width           =   1695
   End
   Begin VB.Line linTaggedRight 
      BorderColor     =   &H00FFFF00&
      X1              =   9000
      X2              =   9000
      Y1              =   2880
      Y2              =   6000
   End
   Begin VB.Line linTaggedBottom 
      BorderColor     =   &H00FFFF00&
      X1              =   120
      X2              =   9060
      Y1              =   5940
      Y2              =   5940
   End
   Begin VB.Line linTaggedLeft 
      BorderColor     =   &H00FFFF00&
      X1              =   180
      X2              =   180
      Y1              =   2880
      Y2              =   6000
   End
   Begin VB.Line linTaggedTop 
      BorderColor     =   &H00FFFF00&
      X1              =   120
      X2              =   9060
      Y1              =   2940
      Y2              =   2940
   End
   Begin VB.Line linSetsRight 
      BorderColor     =   &H00FFFF00&
      X1              =   8400
      X2              =   8400
      Y1              =   240
      Y2              =   2640
   End
   Begin VB.Line linSetsBottom 
      BorderColor     =   &H00FFFF00&
      X1              =   6120
      X2              =   8460
      Y1              =   2580
      Y2              =   2580
   End
   Begin VB.Line linSetsLeft 
      BorderColor     =   &H00FFFF00&
      X1              =   6180
      X2              =   6180
      Y1              =   240
      Y2              =   2640
   End
   Begin VB.Line linSetsTop 
      BorderColor     =   &H00FFFF00&
      X1              =   6120
      X2              =   8460
      Y1              =   300
      Y2              =   300
   End
   Begin VB.Line linNPVRight 
      BorderColor     =   &H00FFFF00&
      X1              =   5040
      X2              =   5040
      Y1              =   720
      Y2              =   2640
   End
   Begin VB.Line linNPVBottom 
      BorderColor     =   &H00FFFF00&
      X1              =   720
      X2              =   5100
      Y1              =   2580
      Y2              =   2580
   End
   Begin VB.Line linNPVLeft 
      BorderColor     =   &H00FFFF00&
      X1              =   780
      X2              =   780
      Y1              =   720
      Y2              =   2640
   End
   Begin VB.Line linNPVTop 
      BorderColor     =   &H00FFFF00&
      X1              =   720
      X2              =   5100
      Y1              =   780
      Y2              =   780
   End
   Begin VB.Label labSetNumbers 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Height          =   255
      Index           =   4
      Left            =   7920
      TabIndex        =   21
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label labSetNumbers 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Height          =   255
      Index           =   3
      Left            =   7920
      TabIndex        =   20
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label labSetNumbers 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Height          =   255
      Index           =   2
      Left            =   7920
      TabIndex        =   19
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label labSetNumbers 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Height          =   255
      Index           =   1
      Left            =   7920
      TabIndex        =   18
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label labSetNumbers 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Height          =   255
      Index           =   0
      Left            =   7920
      TabIndex        =   17
      Top             =   720
      Width           =   135
   End
   Begin VB.Label labParametersUnits 
      BackColor       =   &H00000000&
      Caption         =   "percent"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   4080
      TabIndex        =   16
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label labParametersUnits 
      BackColor       =   &H00000000&
      Caption         =   "percent"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   4080
      TabIndex        =   15
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label labParametersUnits 
      BackColor       =   &H00000000&
      Caption         =   "percent"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   4080
      TabIndex        =   14
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label labParametersUnits 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   4080
      TabIndex        =   13
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label labParametersTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Royalties"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   6420
      TabIndex        =   9
      Top             =   2160
      Width           =   1035
   End
   Begin VB.Label labParametersTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Financing"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   6420
      TabIndex        =   8
      Top             =   1800
      Width           =   1035
   End
   Begin VB.Label labParametersTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Processing"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   6420
      TabIndex        =   7
      Top             =   1440
      Width           =   1035
   End
   Begin VB.Label labParametersTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Mining"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   6420
      TabIndex        =   6
      Top             =   1080
      Width           =   1035
   End
   Begin VB.Label labParametersTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Grades"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   6420
      TabIndex        =   5
      Top             =   720
      Width           =   1035
   End
   Begin VB.Label labParametersTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Third Discount Rate"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   4
      Top             =   2160
      Width           =   1995
   End
   Begin VB.Label labParametersTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Second Discount Rate"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   3
      Top             =   1800
      Width           =   1995
   End
   Begin VB.Label labParametersTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "First Discount Rate"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   2
      Top             =   1440
      Width           =   1995
   End
   Begin VB.Label labParametersTitles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Present Year"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   1
      Top             =   1080
      Width           =   1995
   End
End
Attribute VB_Name = "frmParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim temphigh As Integer
Dim tempwide As Integer
Dim newtag As Integer

Private Sub Form_Activate()

Dim i As Integer

If IsHelpOn = True Then
  If LastCell < 4 Then
    txtParametersValues(LastCell).SetFocus
  Else
    txtRangePercent(LastCell - 4).SetFocus
  End If
  IsHelpOn = False
Else
  DoNotChange = True
  ShowMenu = True
    For i = 0 To 3
      txtParametersValues(i).Text = LTrim(RTrim(Str(Sets(23 + i))))
    Next i
    hscTags.Value = 1
    labParameters(8).Caption = LTrim(RTrim(Str(IndTagData(hscTags.Value).SetNumber)))
    optDistribution(0).Value = True
    For i = 0 To 4
      If i < 3 Then
        labSetNumbers(i) = Np(i + 3)
      Else
        labSetNumbers(i) = Np(i + 4)
      End If
    Next i
    If IndTagData(hscTags.Value).MinChange <> True Then
      IndTagData(hscTags.Value).MinPercent = -25
    End If
    txtRangePercent(0).Text = Format(LTrim(RTrim(Str(IndTagData(hscTags.Value).MinPercent))), "###.00")

    If IndTagData(hscTags.Value).MaxChange <> True Then
      IndTagData(hscTags.Value).MaxPercent = 25
    End If
    txtRangePercent(1).Text = Format(LTrim(RTrim(Str(IndTagData(hscTags.Value).MaxPercent))), "###.00")
  DoNotChange = False
  
  If InsertFlag = True Then
    labInsert.Caption = "Insert"
  Else
    labInsert.Caption = "Typeover"
  End If
  
  Call selecttag(hscTags.Value, IndTagData(hscTags.Value).SetNumber)

End If

End Sub

Private Sub Form_Deactivate()

  If ShowMenu = True Then
    frmParameters.Hide
    Call InputMenuAccess(2)
  End If
  
End Sub

Private Sub Form_Load()

If FullScreen = False Then
  frmParameters.Top = (Screen.Height - (frmParameters.Height + 350)) / 2
  frmParameters.Left = (Screen.Width - frmParameters.Width) / 2
Else
  frmParameters.Top = 0
  frmParameters.Left = 0
  frmParameters.WindowState = 2
End If
   
If frmParameters.Top < 0 Then frmParameters.Top = 0
If frmParameters.Left < 0 Then frmParameters.Left = 0

tempwide = frmParameters.ScaleWidth
temphigh = frmParameters.ScaleHeight

Call screenstuff
 
End Sub

Private Sub Form_Resize()

tempwide = frmParameters.ScaleWidth
temphigh = frmParameters.ScaleHeight

Call screenstuff

End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  frmParameters.Hide
  If ShowMenu = True Then Call InputMenuAccess(2)

End Sub

Private Sub hscTags_Change()

labParameters(7).Caption = LTrim(RTrim(Str(hscTags.Value)))
labParameters(8).Caption = LTrim(RTrim(Str(IndTagData(hscTags.Value).SetNumber)))

If IndTagData(hscTags.Value).MinChange <> True Then
  IndTagData(hscTags.Value).MinPercent = -25
End If
txtRangePercent(0).Text = Format(LTrim(RTrim(Str(IndTagData(hscTags.Value).MinPercent))), "####.00")

If IndTagData(hscTags.Value).MaxChange <> True Then
  IndTagData(hscTags.Value).MaxPercent = 25
End If
txtRangePercent(1).Text = Format(LTrim(RTrim(Str(IndTagData(hscTags.Value).MaxPercent))), "####.00")

Call selecttag(hscTags.Value, IndTagData(hscTags.Value).SetNumber)

txtRangePercent(0).SetFocus

End Sub

Private Sub imgBackToMenu_Click()

  frmParameters.Hide
  If ShowMenu = True Then Call InputMenuAccess(2)

End Sub

Private Sub labBackToMenu_Click()
  
  frmParameters.Hide
  If ShowMenu = True Then Call InputMenuAccess(2)

End Sub

Private Sub labparametermisc_Click()

End Sub
Private Sub labParametersHelp_Click()

Dim begin As Integer
Dim sendindex As Integer

begin = 160

sendindex = LastCell

WhichScreen = 9
ShowMenu = False

Call frmApexHelp.gethelptext(sendindex, begin)
frmApexHelp.Show

End Sub

Private Sub labPrintScreen_Click()

  ShowMenu = False
  job = 12
  Call printstuffout(job)

End Sub

Private Sub optDistribution_Click(Index As Integer)

Dim i As Integer

If DoNotChange = True Then Exit Sub

If optDistribution(1) = True Then
  IndTagData(hscTags.Value).Distribution = "Uniform"
  For i = 1 To 50
    DepTagData(hscTags.Value, i).Distribution = "Uniform"
  Next i
ElseIf optDistribution(2) = True Then
  IndTagData(hscTags.Value).Distribution = "Log Normal"
  For i = 1 To 50
    DepTagData(hscTags.Value, i).Distribution = "Log Normal"
  Next i
ElseIf optDistribution(3) = True Then
  IndTagData(hscTags.Value).Distribution = "Skewed"
  For i = 1 To 50
    DepTagData(hscTags.Value, i).Distribution = "Skewed"
  Next i
Else
  IndTagData(hscTags.Value).Distribution = "Normal"
  For i = 1 To 50
    DepTagData(hscTags.Value, i).Distribution = "Normal"
  Next i
End If

txtRangePercent(0).SetFocus

End Sub
Private Sub txtParametersValues_Change(Index As Integer)

If DoNotChange = True Then Exit Sub

If Index < 4 Then
  Sets(Index + 23) = CCur(Val(txtParametersValues(Index).Text))
Else
  If CCur(Val(txtParametersValues(Index).Text)) > 10000 Then
    riter = 10000
  Else
    riter = CCur(Val(txtParametersValues(Index).Text))
  End If
End If

Select Case Index
  Case 0
    If Sets(23) < Sets(12) - 20 Or Sets(23) > Sets(12) + 70 Then
      Sets(23) = Sets(12)
    End If
  Case 1
    If Sets(24) < 0 Or Sets(24) >= 250 Then
      Sets(24) = 10
    End If
  Case 2
    If Sets(25) < 0 Or Sets(25) >= 250 Then
      Sets(25) = 15
    End If
  Case 3
    If Sets(26) < 0 Or Sets(26) >= 250 Then
      Sets(26) = 20
    End If
End Select

End Sub

Public Sub selecttag(thetag As Integer, theset As Integer)

Dim i As Integer
Dim ii As Integer
Dim k As Integer
Dim depcount As Integer
Dim tempout As String

For i = 1 To 180
  If Tagged(theset, i).Independent = thetag Then
    ii = i
  End If
Next i

If Left(IndTagData(thetag).Distribution, 1) = "" Then
  IndTagData(thetag).Distribution = "Normal"
  For i = 1 To 50
    DepTagData(thetag, i).Distribution = "Normal"
  Next i
End If

Select Case Left(LCase(IndTagData(thetag).Distribution), 1)
  Case "u"
    optDistribution(1).Value = True
  Case "l"
    optDistribution(2).Value = True
  Case "s"
    optDistribution(3).Value = True
  Case Else
    optDistribution(0).Value = True
End Select

If ii = 0 Then
  labParameters(0).Caption = "Nothing Tagged"
  labParameters(1).Caption = "0.0"
ElseIf ii < 131 Then
  labParameters(0).Caption = LTrim(RTrim(IndTagData(thetag).Title))
  tempout = LTrim(Str(Primary(theset, ii)))
  Call findaformat(ii, tempout)
  labParameters(1).Caption = LTrim(tempout)
Else
  labParameters(0).Caption = LTrim(RTrim(IndTagData(thetag).Title))
  tempout = LTrim(Str(CapitalData(ii - 131).PurchaseAmount))
  Call findaformat(ii, tempout)
  labParameters(1).Caption = LTrim(tempout)
End If

labParameters(2).Caption = LTrim(RTrim(IndTagData(thetag).Units))

For k = 0 To 1
  If ii < 131 Then
    tempout = LTrim(Str(Primary(theset, ii) * CCur(Val(txtRangePercent(k))) / 100) + Primary(theset, ii))
  Else
    tempout = LTrim(Str(CapitalData(ii - 131).PurchaseAmount * CCur(Val(txtRangePercent(k))) / 100) + CapitalData(ii - 131).PurchaseAmount)
  End If
  Call findaformat(ii, tempout)
  labParameters(k + 3).Caption = tempout
  If k = 0 Then
    IndTagData(thetag).Minimum = CCur(Val(Format(labParameters(k + 3).Caption, "##########.####")))
    For i = 1 To 50
      If DepTagData(thetag, i).TheCell <> 0 And DepTagData(thetag, i).TheCell < 131 Then
        DepTagData(thetag, i).Minimum = ((Primary(DepTagData(thetag, i).SetNumber, DepTagData(thetag, i).TheCell)) * (CCur(Val(txtRangePercent(k))) / 100)) + Primary(DepTagData(thetag, i).SetNumber, DepTagData(thetag, i).TheCell)
      ElseIf DepTagData(thetag, i).TheCell <> 0 And DepTagData(thetag, i).TheCell > 130 Then
        DepTagData(thetag, i).Minimum = ((CapitalData(DepTagData(thetag, i).TheCell - 131).PurchaseAmount * (CCur(Val(txtRangePercent(k))) / 100)) + CapitalData(DepTagData(thetag, i).TheCell - 131).PurchaseAmount)
      End If
    Next i
  Else
    IndTagData(thetag).Maximum = CCur(Val(Format(labParameters(k + 3).Caption, "##########.####")))
    For i = 1 To 50
      If DepTagData(thetag, i).TheCell <> 0 And DepTagData(thetag, i).TheCell < 131 Then
         DepTagData(thetag, i).Maximum = ((Primary(DepTagData(thetag, i).SetNumber, DepTagData(thetag, i).TheCell) * (CCur(Val(txtRangePercent(k))) / 100)) + Primary(DepTagData(thetag, i).SetNumber, DepTagData(thetag, i).TheCell))
      ElseIf DepTagData(thetag, i).TheCell <> 0 And DepTagData(thetag, i).TheCell > 130 Then
        DepTagData(thetag, i).Maximum = ((CapitalData(DepTagData(thetag, i).TheCell - 131).PurchaseAmount * (CCur(Val(txtRangePercent(k))) / 100)) + CapitalData(DepTagData(thetag, i).TheCell - 131).PurchaseAmount)
      End If
    Next i
  End If
  labParameters(k + 5).Caption = LTrim(RTrim(IndTagData(thetag).Units))
Next k

End Sub

Private Sub txtParametersValues_GotFocus(Index As Integer)

LastCell = Index

End Sub

Private Sub txtParametersValues_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

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
        If InStr(txtParametersValues(Index).Text, ".") = 0 Then
          SendKeys "{DELETE}", False
        End If
      Else
        SendKeys "{DELETE}", False
      End If
  End Select
End If

End Sub

Private Sub txtParametersValues_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii = 46 Then
  If InStr(txtParametersValues(Index).Text, ".") > 0 Then
    Beep
    KeyAscii = 0
  End If
End If

If KeyAscii = 44 Then
  Beep
  KeyAscii = 0
End If

End Sub
Private Sub txtParametersValues_LostFocus(Index As Integer)
DoNotChange = True
If Index = 4 Then
  txtParametersValues(Index).Text = Format(riter, "###,###,##0")
End If
DoNotChange = False
End Sub

Private Sub txtRangePercent_Change(Index As Integer)

If Index = 0 Then
  IndTagData(hscTags).MinPercent = CCur(Val(txtRangePercent(0).Text))
  IndTagData(hscTags.Value).MinChange = True
Else
  IndTagData(hscTags).MaxPercent = CCur(Val(txtRangePercent(1).Text))
  IndTagData(hscTags.Value).MaxChange = True
End If

Call selecttag(hscTags.Value, IndTagData(hscTags.Value).SetNumber)

End Sub
Public Sub screenstuff()
  
  Dim X As Integer
  Dim Y As Currency
  
  labParametersHeading.Top = temphigh * 0.0187
  labParametersHeading.Left = tempwide * 0.0194
  
  linNPVTop.X1 = tempwide * 0.0787
  linNPVTop.X2 = tempwide * 0.5574
  linNPVTop.Y1 = temphigh * 0.1215
  linNPVTop.Y2 = temphigh * 0.1215
  
  linNPVBottom.X1 = tempwide * 0.0787
  linNPVBottom.X2 = tempwide * 0.5574
  linNPVBottom.Y1 = temphigh * 0.4019
  linNPVBottom.Y2 = temphigh * 0.4019
  
  linNPVLeft.X1 = tempwide * 0.0852
  linNPVLeft.X2 = tempwide * 0.0852
  linNPVLeft.Y1 = temphigh * 0.1121
  linNPVLeft.Y2 = temphigh * 0.4112

  linNPVRight.X1 = tempwide * 0.5508
  linNPVRight.X2 = tempwide * 0.5508
  linNPVRight.Y1 = temphigh * 0.1121
  linNPVRight.Y2 = temphigh * 0.4112
  
  linTaggedTop.X1 = tempwide * 0.0131
  linTaggedTop.X2 = tempwide * 0.9902
  linTaggedTop.Y1 = temphigh * 0.4579
  linTaggedTop.Y2 = temphigh * 0.4579
  
  linTaggedBottom.X1 = tempwide * 0.0131
  linTaggedBottom.X2 = tempwide * 0.9902
  linTaggedBottom.Y1 = temphigh * 0.9252
  linTaggedBottom.Y2 = temphigh * 0.9252
  
  linTaggedLeft.X1 = tempwide * 0.0197
  linTaggedLeft.X2 = tempwide * 0.0197
  linTaggedLeft.Y1 = temphigh * 0.4486
  linTaggedLeft.Y2 = temphigh * 0.9346

  linTaggedRight.X1 = tempwide * 0.9836
  linTaggedRight.X2 = tempwide * 0.9836
  linTaggedRight.Y1 = temphigh * 0.4486
  linTaggedRight.Y2 = temphigh * 0.9346
    
  linSetsTop.X1 = tempwide * 0.6689
  linSetsTop.X2 = tempwide * 0.9246
  linSetsTop.Y1 = temphigh * 0.0467
  linSetsTop.Y2 = temphigh * 0.0467
  
  linSetsBottom.X1 = tempwide * 0.6689
  linSetsBottom.X2 = tempwide * 0.9246
  linSetsBottom.Y1 = temphigh * 0.4019
  linSetsBottom.Y2 = temphigh * 0.4019
  
  linSetsLeft.X1 = tempwide * 0.6754
  linSetsLeft.X2 = tempwide * 0.6754
  linSetsLeft.Y1 = temphigh * 0.0374
  linSetsLeft.Y2 = temphigh * 0.4112

  linSetsRight.X1 = tempwide * 0.918
  linSetsRight.X2 = tempwide * 0.918
  linSetsRight.Y1 = temphigh * 0.0374
  linSetsRight.Y2 = temphigh * 0.4112
    
  linMidBoxLeft.X1 = tempwide * 0.1639
  linMidBoxLeft.X2 = tempwide * 0.1639
  linMidBoxLeft.Y1 = temphigh * 0.4953
  linMidBoxLeft.Y2 = temphigh * 0.9159
  
  linMidBoxMiddle.X1 = tempwide * 0.1705
  linMidBoxMiddle.X2 = tempwide * 0.8131
  linMidBoxMiddle.Y1 = temphigh * 0.6449
  linMidBoxMiddle.Y2 = temphigh * 0.6449
  
  linMidBoxRight.X1 = tempwide * 0.8197
  linMidBoxRight.X2 = tempwide * 0.8197
  linMidBoxRight.Y1 = temphigh * 0.4673
  linMidBoxRight.Y2 = temphigh * 0.9159
  
  LinIteration.X1 = tempwide * 0.8262
  LinIteration.X2 = tempwide * 0.977
  LinIteration.Y1 = temphigh * 0.7757
  LinIteration.Y2 = temphigh * 0.7757
  
  For X = 0 To 3
    labParametersTitles(X).Top = (X * 0.0561 * temphigh) + (temphigh * 0.1682)
    labParametersTitles(X).Left = tempwide * 0.118
    labParametersTitles(X).Width = tempwide * 0.218
    txtParametersValues(X).Top = (X * 0.0561 * temphigh) + (temphigh * 0.1636)
    txtParametersValues(X).Left = tempwide * 0.3541
    txtParametersValues(X).Width = tempwide * 0.0869
    labParametersUnits(X).Top = (X * 0.0561 * temphigh) + (temphigh * 0.1682)
    labParametersUnits(X).Left = tempwide * 0.4459
    optDistribution(X).Top = (X * 0.0561 * temphigh) + (temphigh * 0.5421)
    optDistribution(X).Left = tempwide * 0.8328
  Next X

  For X = 0 To 4
    labParametersTitles(X + 4).Top = (X * 0.0561 * temphigh) + (temphigh * 0.1121)
    labParametersTitles(X + 4).Left = tempwide * 0.7016
    labParametersTitles(X + 4).Width = tempwide * 0.1131
    labSetNumbers(X).Top = (X * 0.0561 * temphigh) + (temphigh * 0.1121)
    labSetNumbers(X).Left = tempwide * 0.8656
  Next X
      
  For X = 0 To 1
    labParametersMisc(X + 7).Top = (X * 0.0935 * temphigh) + (temphigh * 0.7383)
    labParametersMisc(X + 7).Left = tempwide * 0.177
    labParametersMisc(X + 11).Top = (X * 0.0935 * temphigh) + (temphigh * 0.7383)
    labParametersMisc(X + 11).Left = tempwide * 0.3869
    txtRangePercent(X).Top = (X * 0.0935 * temphigh) + (temphigh * 0.7336)
    txtRangePercent(X).Left = tempwide * 0.3016
    txtRangePercent(X).Width = tempwide * 0.0803
    labParameters(X + 3).Top = (X * 0.0935 * temphigh) + (temphigh * 0.7383)
    labParameters(X + 3).Left = tempwide * 0.4787
    labParameters(X + 3).Width = tempwide * 0.1787
    labParameters(X + 5).Top = (X * 0.0935 * temphigh) + (temphigh * 0.7383)
    labParameters(X + 5).Left = tempwide * 0.6623
    labParameters(X + 5).Width = tempwide * 0.1459
  Next X
      
  labParameters(0).Top = temphigh * 0.5607
  labParameters(0).Left = tempwide * 0.2164
  labParameters(0).Width = tempwide * 0.2705
  
  labParameters(1).Top = temphigh * 0.5607
  labParameters(1).Left = tempwide * 0.4918
  labParameters(1).Width = tempwide * 0.1656

  labParameters(2).Top = temphigh * 0.5607
  labParameters(2).Left = tempwide * 0.6623
  labParameters(2).Width = tempwide * 0.1459
      
  labParameters(7).Top = temphigh * 0.6215
  labParameters(7).Left = tempwide * 0.1115
  labParameters(7).Width = tempwide * 0.0115
  
  labParameters(8).Top = temphigh * 0.785
  labParameters(8).Left = tempwide * 0.0328
  labParameters(8).Width = tempwide * 0.1197
 
  labParametersMisc(0).Top = temphigh * 0.1121
  labParametersMisc(0).Left = tempwide * 0.0721
  
  labParametersMisc(1).Top = temphigh * 0.0374
  labParametersMisc(1).Left = tempwide * 0.6623

  labParametersMisc(2).Top = temphigh * 0.4486
  labParametersMisc(2).Left = tempwide * 0.0066
  
  labParametersMisc(3).Top = temphigh * 0.5794
  labParametersMisc(3).Left = tempwide * 0.0328
  labParametersMisc(3).Width = tempwide * 0.1197
    
  labParametersMisc(4).Top = temphigh * 0.7383
  labParametersMisc(4).Left = tempwide * 0.0328
  labParametersMisc(4).Width = tempwide * 0.1197

  labParametersMisc(5).Top = temphigh * 0.4953
  labParametersMisc(5).Left = tempwide * 0.2951
  labParametersMisc(5).Width = tempwide * 0.1131

  labParametersMisc(6).Top = temphigh * 0.4953
  labParametersMisc(6).Left = tempwide * 0.5967
  labParametersMisc(6).Width = tempwide * 0.1131
  
  labParametersMisc(9).Top = temphigh * 0.6729
  labParametersMisc(9).Left = tempwide * 0.2754
  labParametersMisc(9).Width = tempwide * 0.1525

  labParametersMisc(10).Top = temphigh * 0.6729
  labParametersMisc(10).Left = tempwide * 0.577
  labParametersMisc(10).Width = tempwide * 0.1525
 
  labParametersMisc(13).Top = temphigh * 0.4858
  labParametersMisc(13).Left = tempwide * 0.8459
  labParametersMisc(13).Width = tempwide * 0.1131
  
  labParametersMisc(14).Top = temphigh * 0.7944
  labParametersMisc(14).Left = tempwide * 0.8459
  labParametersMisc(14).Width = tempwide * 0.1131
  
  txtParametersValues(4).Top = temphigh * 0.8505
  txtParametersValues(4).Left = tempwide * 0.8524
  txtParametersValues(4).Width = tempwide * 0.1
 
  hscTags.Top = temphigh * 0.6262
  hscTags.Left = (tempwide * 0.0656) - 94
  
  labBackToMenu.Top = temphigh * 0.9532
  labBackToMenu.Left = tempwide * 0.0656

  imgBackToMenu.Top = temphigh * 0.9626
  imgBackToMenu.Left = tempwide * 0.0066
  imgBackToMenu.Width = tempwide * 0.0541

  labPrintScreen.Top = temphigh * 0.9532
  labPrintScreen.Left = tempwide * 0.6361
  labPrintScreen.Width = tempwide * 0.1066

  labParametersHelp.Top = temphigh * 0.9532
  labParametersHelp.Left = tempwide * 0.9377

  labInsert.Top = temphigh * 0.9562
  labInsert.Left = tempwide * 0.3344
  labInsert.Width = tempwide * 0.1066
  
End Sub
Private Sub txtRangePercent_GotFocus(Index As Integer)

LastCell = Index + 4

End Sub
Private Sub txtRangePercent_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

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
    Case 48 To 57, 189, 190
      If KeyCode = 190 Then
        If InStr(txtRangePercent(Index).Text, ".") = 0 Then
          SendKeys "{DELETE}", False
        End If
      Else
        SendKeys "{DELETE}", False
      End If
  End Select
End If

End Sub
Private Sub txtRangePercent_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii = 46 Then
  If InStr(txtRangePercent(Index).Text, ".") > 0 Then
    Beep
    KeyAscii = 0
  End If
End If

If KeyAscii = 44 Then
  Beep
  KeyAscii = 0
End If

End Sub


