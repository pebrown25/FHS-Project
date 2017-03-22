VERSION 4.00
Begin VB.Form frmFHSWNote 
   Caption         =   "FHSWNOTE - Comments"
   ClientHeight    =   3105
   ClientLeft      =   3420
   ClientTop       =   3060
   ClientWidth     =   3270
   Height          =   3510
   Left            =   3360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   Top             =   2715
   Width           =   3390
   Begin VB.TextBox txtNOTES 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS LineDraw"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   720
      Width           =   3255
   End
   Begin VB.Frame NoteMNUFrame 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      Begin VB.TextBox txtLCNT 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtLNUM 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdSAVE 
         Caption         =   "Save"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdRETURN 
         Cancel          =   -1  'True
         Caption         =   "Return"
         Height          =   255
         Left            =   2400
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmFHSWNote"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
DefInt A-Z
Dim COMBGN, SRTYPE, SRNO
Dim holdTEXT As String
Dim CTOT, ctcnt1, ctcnt2, ctmax, ctptr()
Private Sub SaveCOMMENTS()
If Not ENV.FFOPEN Then
   Call FAMOpen("WRITE")
   If Not OKAY Then Exit Sub
End If
If SRNO = 0 Then Exit Sub
' Convert CR+LF to "\"
CRLF$ = Chr$(13) + Chr$(10)
X = 1: Y = InStr(X, txtNOTES.TEXT, CRLF$)
Z = Len(txtNOTES.TEXT)
holdTEXT = ""
While Y <> 0
   holdTEXT = holdTEXT + MID$(txtNOTES.TEXT, X, Y - X) + "\"
   X = Y + 2
   Y = InStr(X, txtNOTES.TEXT, CRLF$)
Wend
If X < Z Then
   holdTEXT = holdTEXT + MID$(txtNOTES.TEXT, X, Z)
End If
' Convert holdTEXT to Comment Records
X = 1: nxtrno = COMBGN: FF3PFX.RTYPE = Chr$(5)
XTEXT$ = RTrim$(MID$(holdTEXT, 1, 38))
ctcnt2 = 0
While XTEXT$ <> ""
   Call FF3GetRec(nxtrno)
   ctcnt2 = ctcnt2 + 1
   If Not (ctcnt2 < ctmax) Then
      ctmax = ctmax + 100
      ReDim Preserve ctptr(ctmax)
   End If
   ctptr(ctcnt2) = FF3PFX.RNO
   FF3COMMENT.SRTYPE = Chr$(SRTYPE)
   FF3COMMENT.SRNO = SRNO
   FF3COMMENT.BPTR = ctptr(ctcnt2 - 1)
   FF3COMMENT.FPTR = ctptr(ctcnt2 + 1)
   FF3COMMENT.CCNT = ctcnt2
   LSet FF3COMMENT.TEXT = XTEXT$: X = X + 38
   Call FF3PutRec
   XTEXT$ = RTrim$(MID$(holdTEXT, X, 38))
   nxtrno = FF3COMMENT.FPTR
Wend
If ctcnt1 > ctcnt2 Then
   ' Delete original comment lines not needed now
   For X = ctcnt2 + 1 To ctcnt1
       Call FF3DelRec(ctptr(X))
       ctptr(X) = 0
       If Not OKAY Then X = ctcnt1
   Next X
   If ctcnt2 > 0 Then
      Call FF3GetRec(ctptr(ctcnt2))
      FF3COMMENT.FPTR = 0
      Call FF3PutRec
   End If
   Else
   ' Update FPTR on NEW comment records
   For X = ctcnt1 - (ctcnt1 = 0) To ctcnt2
       Call FF3GetRec(ctptr(X))
       FF3COMMENT.FPTR = ctptr(X + 1)
       Call FF3PutRec
   Next X
End If
ncombgn = ctptr(1)
If ncombgn <> 0 Then
   Call FF3GetRec(ncombgn)
   FF3COMMENT.CCNT = ctcnt2
   FF3COMMENT.BPTR = 0
   Call FF3PutRec
End If
Select Case SRTYPE
   Case 1: ' Name record comments
      Call FF1GetRec(SRNO)
      FF1Rec.COM = ncombgn
      Call FF1PutREC
   Case 2: ' Address record comments
      Call FF2GetRec(SRNO)
      FF2PFX.COM = ncombgn
      Call FF2PutRec
   Case 3: ' Spouse record comments
      Call FF3GetRec(SRNO)
      FF3SPOUSE.COM = ncombgn
      Call FF3PutRec
   Case Else: ' Misc Info record comments
      Call FF3GetRec(SRNO)
      FF3PNTRS.COM = ncombgn
      Call FF3PutRec
End Select
COMBGN = ncombgn
ShowCOMMENTS
Call FamCLOSE
End Sub

Private Sub ShowCNTS()
txtLNUM.TEXT = (txtNOTES.SelStart + 37) \ 38
txtLCNT.TEXT = (Len(txtNOTES.TEXT) + 37) \ 38
End Sub

Sub ViewCOMMENTS(X, Y, Z)
   COMBGN = X: SRTYPE = Y: SRNO = Z
   frmFHSWUPDT.Enabled = False
   Select Case SRTYPE
      Case 3: frmFHSWSpse.Enabled = False
   End Select
   Call ShowCOMMENTS
End Sub
Private Sub ShowCOMMENTS()
   OKAY = -1
   If Not ENV.FFOPEN Then
      Call FAMOpen("WRITE")
      If Not OKAY Then Exit Sub
   End If
BLDCPAGE:
   txtNOTES.TEXT = "": X = COMBGN: CTOT = 0
   ctmax = 100: ReDim ctptr(ctmax)
   ctcnt1 = 0: holdTEXT = ""
   While (X <> 0)
      Call FF3GetRec(X): If Not OKAY Then Exit Sub
      If CTOT = 0 Then CTOT = FF3COMMENT.CCNT
      ctcnt1 = ctcnt1 + 1
      If Not (ctcnt1 < ctmax) Then
         ctmax = ctmax + 1
         ReDim ctptr(ctmax)
      End If
      ctptr(ctcnt1) = X
      holdTEXT = holdTEXT + FF3COMMENT.TEXT
      X = FF3COMMENT.FPTR
   Wend
   ' Convert "\" character to CR+LF
   CRLF$ = Chr$(13) + Chr$(10)
   X = 1: Y = InStr(X, holdTEXT, "\")
   Z = Len(holdTEXT)
   While Y <> 0
      txtNOTES.TEXT = txtNOTES.TEXT + MID$(holdTEXT, X, Y - X) + CRLF$
      X = Y + 1
      Y = InStr(X, holdTEXT, "\")
   Wend
   If X < Z Then
      txtNOTES.TEXT = txtNOTES.TEXT + RTrim$(MID$(holdTEXT, X, Z))
   End If
   holdTEXT = txtNOTES.TEXT
   frmFHSWNote.Show
   ShowCNTS
End Sub

Private Sub cmdRETURN_Click()
Unload frmFHSWNote
End Sub

Private Sub cmdSAVE_Click()
Call SaveCOMMENTS
End Sub

Private Sub Form_Load()
FormLoaded(2) = FormLoaded(2) Or NotesForm
frmHEIGHT = 2655
frmWIDTH = 7900
t0 = frmFHSWUPDT.TOP: l0 = frmFHSWUPDT.Left
t1 = GetSetting("FHSWIN", "frmNOTE", "Top", 0)
If t1 = 0 Then
   t1 = t0 + 4485
   If (t1 + frmHEIGHT) > Screen.Height Then t1 = frmFHSWUPDT.TOP + 450
   l1 = l0 - (frmWIDTH - frmFHSWUPDT.WIDTH) / 2
   If l1 < 0 Then l1 = frmFHSWUPDT.Left: frmWIDTH = frmFHSWUPDT.WIDTH
   Else
   t1 = 1440: l1 = 1440
End If
frmFHSWNote.TOP = GetSetting("FHSWIN", "frmNote", "Top", t1)
frmFHSWNote.Left = GetSetting("FHSWIN", "frmNote", "Left", l1)
frmFHSWNote.Height = GetSetting("FHSWIN", "frmNote", "Height", frmHEIGHT)
frmFHSWNote.WIDTH = GetSetting("FHSWIN", "frmNote", "Width", frmWIDTH)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If holdTEXT <> RTrim$(txtNOTES.TEXT) Then
   Call PutMSG("Changes to NOTES have NOT been Saved...Continue? (Y/N)")
   Call GetRESP
   If CA$ <> SNGLKEY$(1) Then OKAY = 0: Cancel = -1
End If
End Sub


Private Sub Form_Resize()
If frmFHSWNote.Height < 1300 Then frmFHSWNote.Height = 1375
If frmFHSWNote.WIDTH < 3500 Then frmFHSWNote.WIDTH = 3500
NoteMNUFrame.WIDTH = frmFHSWNote.WIDTH - 165
txtNOTES.Height = frmFHSWNote.Height - 1200
txtNOTES.WIDTH = frmFHSWNote.WIDTH - 165
cmdRETURN.Left = frmFHSWNote.WIDTH - 1020
txtLCNT.Left = frmFHSWNote.WIDTH - 1620
txtLNUM.Left = frmFHSWNote.WIDTH - 2220
End Sub


Private Sub Form_Unload(Cancel As Integer)
FormLoaded(2) = FormLoaded(2) And (255 - NotesForm)
SaveSetting "FHSWIN", "frmNOTE", "Top", frmFHSWNote.TOP
SaveSetting "FHSWIN", "frmNOTE", "Left", frmFHSWNote.Left
SaveSetting "FHSWIN", "frmNOTE", "Height", frmFHSWNote.Height
SaveSetting "FHSWIN", "frmNOTE", "Width", frmFHSWNote.WIDTH
frmFHSWUPDT.Enabled = True
Select Case SRTYPE
   Case 3: frmFHSWSpse.Enabled = True
End Select
End Sub

Private Sub txtNOTES_Change()
ShowCNTS
End Sub

Private Sub txtNOTES_KeyUp(KeyCode As Integer, Shift As Integer)
ShowCNTS
End Sub

Private Sub txtNOTES_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ShowCNTS
End Sub

