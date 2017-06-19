VERSION 5.00
Begin VB.Form frmApexHelp 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Apex Help Window"
   ClientHeight    =   5115
   ClientLeft      =   1260
   ClientTop       =   1800
   ClientWidth     =   5595
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5115
   ScaleWidth      =   5595
   Begin VB.CommandButton comApexHelp 
      Caption         =   "&Return"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   4740
      Width           =   5475
   End
   Begin VB.Label labApexHelp 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   19
      Top             =   600
      Width           =   5355
   End
   Begin VB.Label labApexHelp 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   18
      Left            =   120
      TabIndex        =   18
      Top             =   4440
      Width           =   5355
   End
   Begin VB.Label labApexHelp 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   17
      Left            =   120
      TabIndex        =   17
      Top             =   4200
      Width           =   5355
   End
   Begin VB.Label labApexHelp 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   16
      Left            =   120
      TabIndex        =   16
      Top             =   3960
      Width           =   5355
   End
   Begin VB.Label labApexHelp 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   15
      Top             =   3720
      Width           =   5355
   End
   Begin VB.Label labApexHelp 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   14
      Top             =   3480
      Width           =   5355
   End
   Begin VB.Label labApexHelp 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   13
      Top             =   3240
      Width           =   5355
   End
   Begin VB.Label labApexHelp 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   12
      Top             =   3000
      Width           =   5355
   End
   Begin VB.Label labApexHelp 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   5355
   End
   Begin VB.Label labApexHelp 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   5355
   End
   Begin VB.Label labApexHelp 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   5355
   End
   Begin VB.Label labApexHelp 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   5355
   End
   Begin VB.Label labApexHelp 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   5355
   End
   Begin VB.Label labApexHelp 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   5355
   End
   Begin VB.Label labApexHelp 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   5355
   End
   Begin VB.Label labApexHelp 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   5355
   End
   Begin VB.Label labApexHelp 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   5355
   End
   Begin VB.Label labApexHelp 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   5355
   End
   Begin VB.Label labApexHelp 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5355
   End
End
Attribute VB_Name = "frmApexHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub comApexHelp_Click()

Dim r As Integer

IsHelpOn = True
ShowMenu = True

frmApexHelp.Hide

End Sub


Public Sub gethelptext(sendindex As Integer, begin As Integer)

Dim fileforhelp As Integer
Dim locat As Integer
Dim pathtohelp As String
Dim nameofhelp As String
Dim helptext As String
Dim temptext As String
Dim linetext As String
Dim lineword As String
Dim lineletter As String
Dim tempx As Integer
Dim start As Integer
Dim lastspace As Integer
Dim r As Integer
Dim hlpfil As String

For r = 0 To 18
    labApexHelp(r).Caption = ""
Next r

fileforhelp = FreeFile
pathtohelp = MainDir
nameofhelp = "apxhlp03.hlp"

hlpfil = pathtohelp & "\" & nameofhelp

Open hlpfil For Binary As fileforhelp

Dim gethelp As helpheader

locat = Len(gethelp) * (sendindex + begin - 1) + 1

If fileforhelp > 0 Then
  Get fileforhelp, locat, gethelp
  helptext = Space(gethelp.length)
  If gethelp.location > 0 Then
    Get fileforhelp, gethelp.location, helptext
    start = 0
    While Len(helptext) > 0
        lineletter = Left(helptext, 1)
        linetext = linetext & lineletter
        helptext = Mid(helptext, 2, Len(helptext) - 1)
        If lineletter = " " And helptext <> "" Then
            lineletter = Left(helptext, 1)
            temptext = helptext
            helptext = Mid(helptext, 2, Len(helptext) - 1)
            While lineletter <> " "
                lineword = lineword & lineletter
                lineletter = Left(helptext, 1)
                If lineletter <> " " Then
                    helptext = Mid(helptext, 2, Len(helptext) - 1)
                End If
            Wend
            If TextWidth(linetext & lineword) >= 5355 Then
                helptext = temptext
                labApexHelp(start).Caption = linetext
                linetext = ""
                lineword = ""
                start = start + 1
            Else
                linetext = linetext + lineword
                If Len(helptext) = 0 Then labApexHelp(start).Caption = linetext
                lineword = ""
            End If
        Else
           labApexHelp(start).Caption = linetext
        End If
     Wend
  End If
End If

End Sub

