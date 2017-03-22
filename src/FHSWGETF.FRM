VERSION 4.00
Begin VB.Form frmGetFile 
   Caption         =   "FHSW - Identify Path/File"
   ClientHeight    =   2880
   ClientLeft      =   2580
   ClientTop       =   3045
   ClientWidth     =   4155
   Height          =   3285
   Left            =   2520
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   4155
   Top             =   2700
   Width           =   4275
   Begin VB.CommandButton cmdButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   280
      Index           =   1
      Left            =   3240
      TabIndex        =   6
      Top             =   100
      Width           =   855
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   280
      Index           =   0
      Left            =   2520
      TabIndex        =   5
      Top             =   100
      Width           =   615
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.FileListBox filFILE 
      Height          =   1620
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1695
   End
   Begin VB.DirListBox dirPATH 
      Height          =   1605
      Left            =   2040
      TabIndex        =   4
      Top             =   1080
      Width           =   2055
   End
   Begin VB.DriveListBox drvDRIVE 
      Height          =   315
      Left            =   2040
      TabIndex        =   2
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox txtTypeFile 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   720
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   1575
   End
   Begin VB.Line Line2 
      X1              =   2400
      X2              =   2400
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4200
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label 
      Caption         =   "Type:"
      Enabled         =   0   'False
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmGetFile"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim GetOpt$, FileType$, FileSufx$
Public Sub GetFile(FType$, Opt$)
FileType$ = FType$
GetOpt$ = Opt$
Select Case FileType$
   Case "REPORT":
      txtTypeFile = "Report File"
      drvDrive.Drive = FDFSetup.RPTDRV
      dirPath.Path = FDFSetup.RPTDRV + ":" + FDFSetup.RPTDIR
      txtFileName = FT$(14)
      FileSufx$ = "FIL"
      filFILE.FILENAME = "*." + FileSufx$
End Select
Select Case GetOpt$
   Case "PATH": drvDrive.Enabled = True
                dirPath.Enabled = True
                filFILE.Enabled = False
                txtFileName.Enabled = False
                drvDrive.SetFocus
   Case "FILE": drvDrive.Enabled = False
                dirPath.Enabled = False
                filFILE.Enabled = True
                txtFileName.Enabled = True
   Case "FULL": drvDrive.Enabled = True
                dirPath.Enabled = True
                filFILE.Enabled = True
                txtFileName.Enabled = True
End Select
Show 1
End Sub
Private Sub cmdButton_Click(Index As Integer)
OKAY = (Index = 0)
If OKAY Then
   Select Case GetOpt$
      Case "PATH": SysVAR.FILENAME = dirPath.Path
      Case Else:
         SysVAR.FILENAME = dirPath.Path + "\" + txtFileName.Text
   End Select
   Select Case FileType$
      Case "REPORT": LSet PRT.DEST = txtFileName.Text
   End Select
   Else
   A = 27: CA$ = Chr$(27)
End If
Unload frmGetFile
End Sub


Private Sub dirPath_Change()
  filFILE.FILENAME = dirPath.Path + "\*." + FileSufx$
End Sub


Private Sub drvDrive_Change()
dirPath.Path = drvDrive.Drive
End Sub

Private Sub filFILE_Click()
txtFileName = filFILE.List(filFILE.ListIndex)
End Sub


Private Sub txtFileName_Change()
X = txtFileName.SelStart
If InStr(txtFileName.Text, ".") = 0 Then
   txtFileName.Text = RTrim$(txtFileName.Text) + "." + FileSufx$
End If
txtFileName.SelStart = X
End Sub

Private Sub txtFileName_GotFocus()
    txtFileName.SelStart = 0
    txtFileName.SelLength = Len(txtFileName.Text)
End Sub

Private Sub txtFileName_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


