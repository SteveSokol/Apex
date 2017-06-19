VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmApexFileMaker 
   BackColor       =   &H00000000&
   Caption         =   "File "
   ClientHeight    =   5070
   ClientLeft      =   3180
   ClientTop       =   1335
   ClientWidth     =   5370
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5070
   ScaleWidth      =   5370
   Begin VB.CommandButton Command3 
      Caption         =   "&Print"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4140
      TabIndex        =   12
      Top             =   2160
      Width           =   975
   End
   Begin VB.DirListBox dirApexFile 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2430
      Left            =   1860
      TabIndex        =   9
      Top             =   1920
      Width           =   2175
   End
   Begin VB.DriveListBox drvApexFile 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1860
      TabIndex        =   8
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton comFileMaker 
      Caption         =   "&Return"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4140
      TabIndex        =   4
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtApexFile 
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   420
      Width           =   4875
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Save"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4140
      TabIndex        =   2
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Open"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4140
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
   Begin VB.FileListBox filApexFile 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3030
      Left            =   240
      Pattern         =   "*.wax"
      TabIndex        =   0
      Top             =   1200
      Width           =   1515
   End
   Begin ComctlLib.ProgressBar proFile 
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   4620
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
      Max             =   1000
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Directory List"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1860
      TabIndex        =   11
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Drive List"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1860
      TabIndex        =   10
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "File List"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   1515
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "File Name "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   180
      Width           =   4875
   End
End
Attribute VB_Name = "frmApexFileMaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const shiftkey As Integer = 1
Dim shiftdown As Integer

Private Sub comFileMaker_Click()
  
If GetPrintFileName = True Then
  frmApexFileMaker.Hide
  frmOutPrint.Show
Else
  frmApexFileMaker.Hide
  Call InputMenuAccess(4)
End If

End Sub

Private Sub Command1_Click()
  Dim loadinput As Integer
  ActiveFile = txtApexFile.Text
  loadinput = 1
  Call filestuff(loadinput)
  frmApexFileMaker.Hide
End Sub

Private Sub Command2_Click()
    
  Dim saveinput As Integer
  Dim i As Integer
  Dim wordlength As Integer
  Dim theword As String
  Dim addext As Integer
  
  JumpShip = False
  wordlength = Len(txtApexFile.Text)
  
  For i = 1 To wordlength
    If Mid(txtApexFile.Text, i, 1) = "\" Then
      theword = ""
    Else
      theword = theword + Mid(txtApexFile.Text, i, 1)
    End If
  Next i
  
  wordlength = Len(theword)
    
  addext = True
  For i = 1 To wordlength
    If Mid(txtApexFile.Text, i, 1) = "." Then
      addext = False
    End If
  Next i
  
  If addext = True Then theword = theword & ".wax  "
  
  addext = False
  For i = 0 To filApexFile.ListCount
    If Left(theword, 8) = Left(filApexFile.List(i), 8) Then
      addext = True
    End If
  Next i
    
  If addext = True Then
     Call getthewarning
  Else
    ActiveFile = txtApexFile.Text
    saveinput = 2
    Call filestuff(saveinput)
    frmApexFileMaker.Hide
  End If
  
End Sub

Private Sub Command3_Click()

Dim saveinput As Integer
Dim i As Integer
Dim wordlength As Integer
Dim theword As String
Dim spreadtheword As String
Dim addext As Integer
Dim tempdir As String

If GetPrintFileName = False Then
  frmApexFileMaker.Hide
Else
  JumpShip = False
  wordlength = Len(txtApexFile.Text)
  
  For i = 1 To wordlength
    If Mid(txtApexFile.Text, i, 1) = "\" Then
      theword = ""
    Else
      theword = theword + Mid(txtApexFile.Text, i, 1)
    End If
  Next i
  
  wordlength = Len(theword)
    
  addext = True
  For i = 1 To wordlength
    If Mid(theword, i, 1) = "." Then
      addext = False
    End If
  Next i
  
  If addext = True Then
    spreadtheword = txtApexFile.Text & ".prn"
    theword = txtApexFile.Text & ".txt"
    addext = False
  Else
    spreadtheword = txtApexFile.Text
    theword = txtApexFile.Text
    addext = False
  End If
  
  For i = 0 To filApexFile.ListCount
    If Left(theword, 8) = Left(filApexFile.List(i), 8) Then
      addext = True
    End If
  Next i
    
  If addext = True Then
    Call getthewarning
  Else
    PrintFileName = theword
    LotusFileName = spreadtheword
    GetPrintFileName = False
    frmApexFileMaker.Hide
  End If
End If

If PrintFileName = "" Then
  WarnNumber = 9
  ShowMenu = False
  frmWarnTheUser.Show
End If
 
PrintFileNumber = FreeFile
Open PrintFileName For Output As PrintFileNumber

Select Case job
  Case 1 To 5
    LotusFileNumber = FreeFile
    Open LotusFileName For Output As LotusFileNumber
    Call filecashflowstuff
  Case 6 To 8
    Call filestatisticalstuff
  Case 9 To 13
    Call fileanalysisstuff
  Case 20 To 24
    Call filedatastuff
End Select
Close PrintFileNumber
Close LotusFileNumber

Select Case job
  Case 1 To 5
    frmCashFlow.Show
  Case 9
    frmRateOfReturn.Show
  Case 10
    frmBreakEven.Show
  Case 11
    frmSensitivity.Show
  Case 12
    frmParameters.Show
  Case 13
    frmRisk.Show
End Select

End Sub

Private Sub dirApexFile_Change()

filApexFile.Path = dirApexFile.Path
txtApexFile.Text = filApexFile.Path

End Sub

Private Sub drvApexFile_Change()

txtApexFile.Text = drvApexFile.Drive
dirApexFile.Path = drvApexFile.Drive

End Sub

Private Sub filApexFile_Click()

  If Right(filApexFile.Path, 1) = "\" Then
    txtApexFile.Text = filApexFile.Path & filApexFile.FileName
  Else
    txtApexFile.Text = filApexFile.Path & "\" & filApexFile.FileName
  End If

End Sub

Private Sub Form_Activate()
  
  txtApexFile.Text = filApexFile.Path
  ShowMenu = True
  If JumpShip = True Then
    JumpShip = False
    Exit Sub
  End If
  
  txtApexFile.SetFocus
  
End Sub

Public Sub whichway()
  frmApexFileMaker.Hide
End Sub

Private Sub Form_Load()

frmApexFileMaker.Top = (Screen.Height - (frmApexFileMaker.Height + 350)) / 2
frmApexFileMaker.Left = (Screen.Width - frmApexFileMaker.Width) / 2

If frmApexFileMaker.Top < 0 Then frmApexFileMaker.Top = 0
If frmApexFileMaker.Left < 0 Then frmApexFileMaker.Left = 0

End Sub

Public Sub getthewarning()

    WarnNumber = 5
    ShowMenu = False
    frmWarnTheUser.Show
   
End Sub

Private Sub txtApexFile_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 45 Then
  If InsertFlag = True Then
    InsertFlag = False
  Else
    InsertFlag = True
  End If
End If

If InsertFlag = False Then
  Select Case KeyCode
    Case 48 To 57, 65 To 90, 186 To 191
      SendKeys "{DELETE}", False
  End Select
End If

End Sub

