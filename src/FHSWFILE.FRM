VERSION 4.00
Begin VB.Form frmFHSWFile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FHSWFILE - Identify Path/File"
   ClientHeight    =   2805
   ClientLeft      =   2730
   ClientTop       =   3630
   ClientWidth     =   4200
   Height          =   3210
   Left            =   2670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   Top             =   3285
   Width           =   4320
   Begin VB.CommandButton cmdButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   280
      Index           =   1
      Left            =   3240
      TabIndex        =   7
      Top             =   100
      Width           =   855
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   280
      Index           =   0
      Left            =   2520
      TabIndex        =   6
      Top             =   100
      Width           =   615
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1695
   End
   Begin VB.FileListBox filFILE 
      Height          =   1620
      Left            =   120
      TabIndex        =   5
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
      TabIndex        =   3
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox txtTypeFile 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   1710
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmFHSWFile"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim GetOpt$, FileType$, FileSufx$, RWTYPE
Dim HoldDrv, HoldPath, HoldFile
Public Sub LOCATE(xType$, OPT$)
Select Case xType$
   Case "ANCESTOR": XWTYPE = 1
   Case "DESCENDANT": XWTYPE = 2
   Case "RELATIVE": XWTYPE = 3
End Select
If XWTYPE > 0 Then
   FileType$ = "RELATIVE": RWTYPE = XWTYPE
   Else
   FileType$ = xType$
End If
GetOpt$ = OPT$
Select Case FileType$
   Case "REPORT":
      txtTypeFile = "Report File"
      xfpfx$ = RPTFPFX$
      xfile$ = PRT.DEST
      FileSufx$ = "RPT"
   Case "RELATIVE":
      Select Case RWTYPE
         Case 1: xX$ = "Ancestor"
         Case 2: xX$ = "Descendant"
         Case 3: xX$ = "Relative"
      End Select
      txtTypeFile.TEXT = xX$ + " Work File"
      xfpfx$ = WRKFPFX$
      xfile$ = FT$(RWTYPE + 4)
      FileSufx$ = "WRK"
   Case "SELECT"
      txtTypeFile.TEXT = "Select Work File"
      xfpfx$ = WRKFPFX$
      xfile$ = FT$(10)
      FileSufx$ = "SEL"
   Case "NODEWORK"
      txtTypeFile.TEXT = "Node Work File"
      xfpfx$ = WRKFPFX$
      xfile$ = FT$(8)
      FileSufx$ = "WRK"
   Case "RELWORK"
      txtTypeFile.TEXT = "Relationship Work File"
      xfpfx$ = WRKFPFX$
      X = WRelType + 4
      If X < 5 Then X = 7
      xfile$ = FT$(X)
      FileSufx$ = "WRK"
   Case "GEDCOM"
      txtTypeFile.TEXT = "GEDCOM File"
      xfpfx$ = TRANPFX$
      xfile$ = FT$(11)
      FileSufx$ = "GED"
   Case "MAILMERG"
      txtTypeFile.TEXT = "MailMerge File"
      xfpfx$ = TRANPFX$
      xfile$ = FT$(12)
      FileSufx$ = "MMG"
   Case "FDF"
      txtTypeFile.TEXT = "Family Definition File"
      xfpfx$ = PGMPath
      xfile$ = CONFIG.FDF
      FileSufx$ = "FDF"
   Case "PDF"
      txtTypeFile.TEXT = "Printer Definition File"
      xfpfx$ = PGMPath
      xfile$ = CONFIG.PDF
      FileSufx$ = "PDF"
   Case "INDEX"
      txtTypeFile.TEXT = "Index File"
      xfpfx$ = FAMFPFX$
      xfile$ = FT$(4)
      FileSufx$ = "NDX"
   Case "FAMFPATH"
      txtTypeFile.TEXT = "Family File Path"
      xfpfx$ = FAMFPFX$
      xfile$ = ""
      FileSufx$ = "NAM"
   Case "WRKFPATH"
      txtTypeFile.TEXT = "Work File Path"
      xfpfx$ = WRKFPFX$
      xfile$ = ""
      FileSufx$ = "WRK"
   Case "TRNFPATH"
      txtTypeFile.TEXT = "Transfer File Path"
      xfpfx$ = TRANPFX$
      xfile$ = ""
      FileSufx$ = "GED"
   Case "RPTFPATH"
      txtTypeFile.TEXT = "Report File Path"
      xfpfx$ = RPTFPFX$
      xfile$ = ""
      FileSufx$ = "RPT"
End Select
X = InStr(xfpfx$, ":")
If X > 0 Then
   On Local Error GoTo BadDrive
   HoldDrv = drvDrive.Drive
   xdrv$ = Left$(xfpfx$, X - 1)
   If (xdrv$ <> " ") And (xdrv$ <> Chr$(0)) Then
      drvDrive.Drive = xdrv$: X = 0
   End If
End If
GoTo SkipBadDrive
BadDrive: Resume Next
SkipBadDrive:
HoldPath = dirPath.path
On Local Error GoTo BadPath
dirPath.path = MID$(xfpfx$, X + 1, 255)
GoTo SkipBadPath
BadPath: Resume Next
SkipBadPath:
X = InStr(xfile$, ".")
If X > 0 Then xfile$ = Left$(xfile$, X - 1)
txtFileName.TEXT = xfile$ + "." + FileSufx$
filFILE.FILENAME = "*." + FileSufx$
Select Case GetOpt$
   Case "PATH": drvDrive.Enabled = True
                dirPath.Enabled = True
                filFILE.Enabled = False
                txtFileName.Enabled = False
                'dirPath.SetFocus
   Case "FILE": drvDrive.Enabled = False
                dirPath.Enabled = False
                filFILE.Enabled = True
                txtFileName.Enabled = True
   Case "FULL": drvDrive.Enabled = True
                dirPath.Enabled = True
                filFILE.Enabled = True
                txtFileName.Enabled = True
End Select
a = 27: CA$ = Chr$(27)
Show 1
EarlyExit:
End Sub
Private Sub CMDButton_Click(Index As Integer)
OKAY = (Index = 0)
If OKAY Then
   Select Case GetOpt$
      Case "PATH": SysVAR.FILENAME = dirPath.path
      Case Else:
         SysVAR.FILENAME = dirPath.path + "\" + txtFileName.TEXT
   End Select
   Select Case FileType$
      Case "INDEX": FT$(4) = txtFileName.TEXT
      Case "REPORT": PRT.DEST = txtFileName.TEXT
                     FDFSetup.RPTDRV = Left$(drvDrive.Drive, 1)
                     FDFSetup.RPTDIR = RTrim$(MID$(dirPath.path, 3, 255)) + "\"
                     RPTFPFX$ = FDFSetup.RPTDRV + ":" + RTrim$(FDFSetup.RPTDIR)
      Case "RELATIVE": FT$(RWTYPE + 4) = txtFileName.TEXT
                     FDFSetup.WRKDRV = Left$(drvDrive.Drive, 1)
                     FDFSetup.WRKDIR = RTrim$(MID$(dirPath.path, 3, 255)) + "\"
                     WRKFPFX$ = FDFSetup.WRKDRV + ":" + RTrim$(FDFSetup.WRKDIR)
      Case "SELECT": FT$(10) = txtFileName.TEXT
                     FDFSetup.WRKDRV = Left$(drvDrive.Drive, 1)
                     FDFSetup.WRKDIR = RTrim$(MID$(dirPath.path, 3, 255)) + "\"
                     WRKFPFX$ = FDFSetup.TRANDRV + ":" + RTrim$(FDFSetup.WRKDIR)
      Case "GEDCOM": FT$(11) = txtFileName.TEXT
                     FDFSetup.TRANDRV = Left$(drvDrive.Drive, 1)
                     FDFSetup.TRANDIR = RTrim$(MID$(dirPath.path, 3, 255)) + "\"
                     TRANPFX$ = FDFSetup.TRANDRV + ":" + RTrim$(FDFSetup.TRANDIR)
      Case "MAILMERG": FT$(12) = txtFileName.TEXT
                     FDFSetup.TRANDRV = Left$(drvDrive.Drive, 1)
                     FDFSetup.TRANDIR = RTrim$(MID$(dirPath.path, 3, 255)) + "\"
                     TRANPFX$ = FDFSetup.TRANDRV + ":" + RTrim$(FDFSetup.TRANDIR)
      Case "PDF": X = InStr(txtFileName.TEXT, ".") - 1
                  If X < 0 Then X = Len(txtFileName.TEXT)
                  CONFIG.PDF = Left$(txtFileName.TEXT, X)
      Case "FDF": X = InStr(txtFileName.TEXT, ".") - 1
                  If X < 0 Then X = Len(txtFileName.TEXT)
                  CONFIG.FDF = Left$(txtFileName.TEXT, X)
      Case "FAMFPATH": FAMFPFX$ = UCase(RTrim$(dirPath.path))
      Case "WRKFPATH": WRKFPFX$ = UCase(RTrim$(dirPath.path))
      Case "TRNFPATH": TRANPFX$ = UCase(RTrim$(dirPath.path))
      Case "RPTFPATH": RPTFPFX$ = UCase(RTrim$(dirPath.path))
   End Select
   a = 0: CA$ = Chr$(0)
   Else
   a = 27: CA$ = Chr$(27)
End If
Unload frmFHSWFile
End Sub


Private Sub dirPath_Change()
  filFILE.FILENAME = dirPath.path + "\" + "*." + FileSufx$
End Sub


Private Sub drvDrive_Change()
On Local Error GoTo BadDrv
dirPath.path = drvDrive.Drive
HoldDrv = drvDrive.Drive
Exit Sub
BadDrv:
drvDrive.Drive = HoldDrv
Resume Next
End Sub

Private Sub filFILE_Click()
txtFileName = filFILE.List(filFILE.ListIndex)
End Sub


Private Sub filFILE_DblClick()
CMDButton_Click (0)
End Sub


Private Sub filFILE_PatternChange()
filFILE.Refresh
End Sub


Private Sub Form_Load()
FormLoaded(1) = FormLoaded(1) Or FileForm
' Set Default Form Location on Screen
frmHEIGHT = frmFHSWFile.Height
frmWIDTH = frmFHSWFile.WIDTH
t1 = GetSetting("FHSWIN", "frmFILE", "TOP", 0)
If t1 = 0 Then
   t00 = mnuFHSWMain.TOP: w00 = mnuFHSWMain.WIDTH
   l00 = mnuFHSWMain.Left
   t1 = t00 + 2615
   l1 = l00 + (w00 - frmWIDTH) / 2
   Else
   t1 = 1440: l1 = 1440
End If
frmFHSWFile.TOP = GetSetting("FHSWIN", "frmFILE", "Top", t1)
frmFHSWFile.Left = GetSetting("FHSWIN", "frmFILE", "Left", l1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
FormLoaded(1) = FormLoaded(1) And (255 - FileForm)
SaveSetting "FHSWIN", "frmFILE", "Top", frmFHSWFile.TOP
SaveSetting "FHSWIN", "frmFILE", "Left", frmFHSWFile.Left

End Sub


Private Sub txtFileName_Change()
X = txtFileName.SelStart
If InStr(txtFileName.TEXT, ".") = 0 Then
   txtFileName.TEXT = RTrim$(txtFileName.TEXT) + "." + FileSufx$
End If
txtFileName.SelStart = X
End Sub

Private Sub txtFileName_DblClick()
If filFILE.Pattern = "*.*" Then
   filFILE.Pattern = "*." + FileSufx$
   Else
   filFILE.Pattern = "*.*"
End If
End Sub


Private Sub txtFileName_GotFocus()
    txtFileName.SelStart = 0
    txtFileName.SelLength = Len(txtFileName.TEXT)
End Sub

Private Sub txtFileName_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


