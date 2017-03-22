VERSION 4.00
Begin VB.Form frmFHSWRPgm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FHSWRPT - Print Report/Chart"
   ClientHeight    =   3600
   ClientLeft      =   1950
   ClientTop       =   3210
   ClientWidth     =   5955
   Height          =   4005
   KeyPreview      =   -1  'True
   Left            =   1890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   Top             =   2865
   Width           =   6075
   Begin VB.CommandButton cmdCHANGE 
      Caption         =   "Change"
      Height          =   315
      Left            =   5160
      TabIndex        =   17
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton cmdRETURN 
      Caption         =   "Return"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox PRTDesc 
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS LineDraw"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      TabIndex        =   14
      Top             =   600
      Width           =   2535
   End
   Begin VB.ComboBox PRTSetup 
      Height          =   315
      Left            =   1080
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdRFView 
      Caption         =   "View"
      Height          =   285
      Left            =   5160
      TabIndex        =   12
      Top             =   1080
      Width           =   700
   End
   Begin VB.CommandButton cmdCANCEL 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   285
      Left            =   5160
      TabIndex        =   11
      Top             =   120
      Width           =   700
   End
   Begin VB.CommandButton cmdRFBrowse 
      Caption         =   "Browse"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4200
      TabIndex        =   6
      Top             =   1080
      Width           =   750
   End
   Begin VB.ComboBox PRTDest 
      Height          =   315
      ItemData        =   "FHSWRPGM.frx":0000
      Left            =   1080
      List            =   "FHSWRPGM.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox RFDest 
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS LineDraw"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      TabIndex        =   5
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdWorkFile 
      Caption         =   "WorkFile"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "Options"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox RPDesc 
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS LineDraw"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      TabIndex        =   3
      Top             =   120
      Width           =   2535
   End
   Begin VB.ComboBox RPSetup 
      Height          =   315
      ItemData        =   "FHSWRPGM.frx":0028
      Left            =   1080
      List            =   "FHSWRPGM.frx":002A
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Line Line5 
      X1              =   5040
      X2              =   5040
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Label lblPRTSetup 
      Caption         =   "Setup:"
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
      Left            =   240
      TabIndex        =   15
      Top             =   630
      Width           =   615
   End
   Begin VB.Line Line4 
      X1              =   50
      X2              =   5880
      Y1              =   1440
      Y2              =   1440
   End
   Begin MSGrid.Grid grdOPTIONS 
      Height          =   1950
      Left            =   1440
      TabIndex        =   10
      Top             =   1560
      Width           =   4395
      _Version        =   65536
      _ExtentX        =   7752
      _ExtentY        =   3440
      _StockProps     =   77
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS LineDraw"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Rows            =   10
      Cols            =   3
      ScrollBars      =   2
      HighLight       =   0   'False
      MouseIcon       =   "FHSWRPGM.frx":002C
   End
   Begin VB.Label lblPRTDest 
      Caption         =   "Print to:"
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
      Left            =   240
      TabIndex        =   1
      Tag             =   "p"
      Top             =   1110
      Width           =   855
   End
   Begin VB.Line Line3 
      X1              =   50
      X2              =   5880
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line2 
      X1              =   1320
      X2              =   1320
      Y1              =   1440
      Y2              =   3600
   End
   Begin VB.Line Line1 
      X1              =   50
      X2              =   5880
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label lblReport 
      Caption         =   "Report:"
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
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   150
      Width           =   735
   End
End
Attribute VB_Name = "frmFHSWRPgm"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
DefInt A-Z
Dim FileDEST$, xOPTIX(), cOPTIX, EDTIX
Dim OptVLth, OptRow, NewOptValu$, OldOptValu$
Dim xPRT As Variant, PRTCode As String











Private Sub BuildOptGrid()
grdOPTIONS.Enabled = False
scrollwid = 250
grdOPTIONS.Cols = 3
grdOPTIONS.Rows = RDFReport.OPTCNT + 1
grdOPTIONS.ColWidth(0) = 250
grdOPTIONS.ColWidth(1) = grdOPTIONS.WIDTH - 1180
grdOPTIONS.ColWidth(2) = 600
h = (grdOPTIONS.Height - 100) / 8
grdOPTIONS.Col = 0
For X = 0 To grdOPTIONS.Rows - 1
   grdOPTIONS.RowHeight(X) = h
   grdOPTIONS.Row = X
   If X > 0 Then
      grdOPTIONS.TEXT = X
   End If
Next X
grdOPTIONS.Row = 0
grdOPTIONS.Col = 1: grdOPTIONS.TEXT = "         Option"
grdOPTIONS.Col = 2: grdOPTIONS.TEXT = "Value"
Call RptOptDisplay
End Sub

Private Sub BuildPRTSetup()
PRTSetup.Clear
For X = 1 To PDFHdr.SETUPS
    PRTSetup.AddItem PSETUPS(X).NAME
    PRTSetup.ItemData(PRTSetup.NewIndex) = X
Next X
PRTSetup.TEXT = PDFSetup.NAME
'Call PRTSetup_Click
End Sub
Private Sub BuildPRTDest()
Select Case RTrim$(PDFSetup.DEST)
   Case "SCRN:": PRTDest.ListIndex = 0
   Case "FILE": PRTDest.ListIndex = 2
   Case Else: PRTDest.ListIndex = 1
End Select
PRTDest_CLICK
End Sub

Private Sub OptCellINIT()
ChkOPT: ' Look for Real Option Entry
cOPTIX = xOPTIX(OptRow)
EDTIX = Asc(MID$(RPTOPT$(cOPTIX), 9, 1))
If EDTIX = 255 Then
   OptVLth = 0: OldOptValu$ = "": NewOptValu$ = ""
   Else
   If (21 < EDTIX) And (EDTIX < 26) Then
      OptVLth = EDTIX - 20
      Else
      OptVLth = 1
   End If
   OldOptValu$ = LTrim$(grdOPTIONS.TEXT)
   NewOptValu$ = ""
End If
grdOPTIONS.Row = OptRow
End Sub


Sub RptOptUPDATE()
If (NewOptValu$ <> OldOptValu$) Then
   grdOPTIONS.Row = OptRow: grdOPTIONS.Col = 2
   Mid$(RPTOPT$(cOPTIX), 12, 5) = grdOPTIONS.TEXT
   Call RptOptEDIT(RPTOPT$(cOPTIX))
   If Not OKAY Then
      grdOPTIONS.TEXT = Right$("    " + OldOptValu$, 5)
      Mid$(RPTOPT$(cOPTIX), 12, 5) = grdOPTIONS.TEXT
      Else
      Call RptOptSET(RPTOPT$(cOPTIX), RptOPTION)
      RDFIndex(ENV.RPTIX).CHG = -1
      If RDFHdr.FType = "RDF" Then RDFHdr.FType = "CHG"
   End If
End If
End Sub


Private Sub RptOptDisplay()
   If (RptPGM = "LIST") Then
      If (RptCODE = "FSUM") Or (RptCODE = "INDX") Then
         cmdWorkFile.Enabled = False
         Else
         cmdWorkFile.Enabled = True
      End If
   End If
   OPTTable.FT = RDFIndex(ENV.RPTIX).OPTBGN
   OPTTable.XL = 0: OPTTable.SIZE = RDFReport.OPTCNT
   ReDim xOPTIX(RDFReport.OPTCNT)
   X = OPTTable.FT: xX$ = Space$(16)
   Y$ = Space$(4): ACHAR$ = "   "
   Z = 0
   While X > 0
      OPTTable.SIZE = OPTTable.SIZE + 1
      OPTTable.LT = X
      LSet xX$ = RPTOPT$(X)
      Call RptOptFIND(xX$)
      Call RptOptEXPAND(xX$)
      OPTIX = CVI(MID$(xX$, 7, 2))
      Z = Z + 1: grdOPTIONS.Row = Z
      xOPTIX(Z) = X
      grdOPTIONS.Col = 1
      If OPTIX > 0 Then
         grdOPTIONS.TEXT = MID$(TABDATA$(OPTIX), 10, 30)
         Else
         grdOPTIONS.TEXT = "Unknown Option"
      End If
      RPTOPT$(X) = xX$ + grdOPTIONS.TEXT
      grdOPTIONS.Col = 2
      grdOPTIONS.TEXT = MID$(xX$, 12, 5)
      X = OPTCHN(X).FWD
   Wend
   Call RptOptINIT
End Sub
Private Sub TurnOffOpts()
If grdOPTIONS.Enabled = True Then
   grdOPTIONS.Enabled = False
   grdOPTIONS.HighLight = False
   If NewOptValu$ <> "" Then RptOptUPDATE
   RPTOPTS(ENV.RPTIX) = RptOPTION
   'cmdCANCEL.DEFAULT = True
   'cmdCANCEL.Cancel = True
   grdOPTIONS.TopRow = 1
   cmdRETURN.Caption = "Return"
   cmdRETURN.DEFAULT = True
   cmdOptions.Caption = "Options"
End If
End Sub
Private Sub TurnONOpts()
If grdOPTIONS.Enabled = False Then
   NewOptValu$ = "": OldOptValu$ = ""
   cmdOptions.Caption = "Save Options"
   cmdRETURN.Caption = "End"
   cmdRETURN.DEFAULT = False
   Call RptOptINIT
   grdOPTIONS.TopRow = 1: grdOPTIONS.Row = 2
   grdOPTIONS.Col = 2: OptRow = grdOPTIONS.Row
   Call OptCellINIT
   grdOPTIONS.SelStartRow = OptRow: grdOPTIONS.SelStartCol = 2
   grdOPTIONS.SelEndRow = OptRow: grdOPTIONS.SelEndCol = 2
   grdOPTIONS.Enabled = True
   grdOPTIONS.HighLight = True
End If
End Sub


Private Sub cmdCANCEL_Click()
a = 27: CA$ = Chr$(27)
If grdOPTIONS.Enabled = True Then
   If (NewOptValu$ <> "") Then
      grdOPTIONS.TEXT = Right$("     " + OldOptValu$, 5)
      NewOptValu$ = ""
      Else
      TurnOffOpts
   End If
   Else
   Unload frmFHSWRPgm
End If
End Sub

Private Sub cmdCHANGE_Click()
Load frmFHSWPrinter
frmFHSWPrinter.cmdPrint.Enabled = False
frmFHSWPrinter.Show 1
PRTSetup.TEXT = PDFSetup.NAME
Call PRTSetup_Click
End Sub

Private Sub cmdOptions_Click()
If grdOPTIONS.Enabled = False Then
   TurnONOpts
   grdOPTIONS.SetFocus
   Else
   TurnOffOpts
   Call RptOptSAVE
End If
End Sub

Private Sub cmdRETURN_Click()
If grdOPTIONS.Enabled = True Then
   TurnOffOpts
   Else
   Unload frmFHSWRPgm
End If
End Sub

Private Sub cmdRFBrowse_Click()
TurnOffOpts
xX$ = PRT.DEST: PRT.DEST = FileDEST$
Call frmFHSWFile.LOCATE("REPORT", "FULL")
If OKAY Then
   RFDest.TEXT = PRT.DEST
   FileDEST$ = PRT.DEST
   Else
   PRT.DEST = xX$
End If
If RTrim$(xX$) <> RTrim$(PRT.DEST) Then PAGE.COUNT = 0
RFDest.SetFocus
End Sub


Private Sub cmdRFView_Click()
Call frmFHSWView.ReportFIL
End Sub

Private Sub Form_Load()
FormLoaded(1) = FormLoaded(1) Or RPgmForm
' Set Default Form Location on Screen
frmHEIGHT = frmFHSWRPgm.Height
frmWIDTH = frmFHSWRPgm.WIDTH
t1 = GetSetting("FHSWIN", "frmRPGM", "TOP", 0)
If t1 = 0 Then
   t00 = mnuFHSWMain.TOP: w00 = mnuFHSWMain.WIDTH
   l00 = mnuFHSWMain.Left
   t1 = (t00 + 1075)
   l1 = l00 + (w00 - frmWIDTH) / 2: If l1 < 0 Then l1 = 1
   Else
   t1 = 1440: l1 = 1440
End If
frmFHSWRPgm.TOP = GetSetting("FHSWIN", "frmRPGM", "Top", t1)
frmFHSWRPgm.Left = GetSetting("FHSWIN", "frmRPGM", "Left", l1)
If FormLoaded(1) And (PrntForm) Then
   cmdCHANGE.Enabled = False
End If
Select Case RptPGM
   Case "RELR": PgmLabel$ = "Ancestor/Descendant/Relative Report Program"
   Case "GRPR": PgmLabel$ = "Family Group Report Program"
                cmdWorkFile.Caption = "Select"
   Case "REGR": PgmLabel$ = "Register Report Program"
   Case "ACHT": PgmLabel$ = "Ancestor Chart Program"
   Case "VBOX": PgmLabel$ = "Vertical Box Chart Program"
                cmdWorkFile.Enabled = False
   Case "HBOX": PgmLabel$ = "Horizontal Box Chart Program"
                cmdWorkFile.Caption = "NodeFile"
   Case "LIST": PgmLabel$ = "Search/Select/LIST Program"
                cmdWorkFile.Caption = "Select"
   Case "TAFL": PgmLabel$ = "Tiny Tafel Program"
   Case "PRTC": PgmLabel$ = "Printer Setup Program"
                cmdWorkFile.Enabled = False
   Case "FILE": PgmLabel$ = "File Setup Program"
                cmdWorkFile.Enabled = False
   Case "MLMG": PgmLabel$ = "MailMerge Export Program"
                cmdWorkFile.Caption = "Select"
End Select
If PgmLabel$ <> "" Then
   frmFHSWRPgm.Caption = "FHSWRPT - " + PgmLabel$
End If
FileDEST$ = FT$(14): PRTCode = "FILE"
NewOptValu$ = ""
Call BuildPRTSetup
End Sub

Private Sub Form_Unload(Cancel As Integer)
For X = 0 To (RPSetup.ListCount - 1)
   Call RPTFree(RPSetup.List(X))
Next X
RptPGM = "": RptCODE = ""
FormLoaded(1) = FormLoaded(1) And (255 - RPgmForm)
SaveSetting "FHSWIN", "frmRPGM", "Top", frmFHSWRPgm.TOP
SaveSetting "FHSWIN", "frmRPGM", "Left", frmFHSWRPgm.Left
If FormLoaded(3) And IndexForm Then
   frmFHSWINDX.Enabled = True
   frmFHSWINDX.SetFocus
   Else
   If FormLoaded(1) And PrntForm Then
      frmFHSWPrinter.Enabled = True
      frmFHSWPrinter.SetFocus
      Else
      If FormLoaded(2) And FamDefForm Then
         frmFHSWFFDF.Enabled = True
         frmFHSWFFDF.SetFocus
         Else
         mnuFHSWMain.SetFocus
      End If
   End If
End If
End Sub


Private Sub cmdPrint_Click()
TurnOffOpts
Select Case RptPGM
   Case "RELR": Call RptRELR("PRINT")
   Case "GRPR": Call RptGRPR("PRINT")
   Case "REGR": Call RptREGR("PRINT")
   Case "ACHT": Call RptACHT("PRINT")
   Case "VBOX": Call RptVBOX("PRINT")
   Case "HBOX": Call RptHBOX("PRINT")
   Case "LIST": Call RptLIST("PRINT")
   Case "FILE": Call RptFile("PRINT")
   Case "PRTC": Call RptPRTC("PRINT")
End Select
End Sub

Private Sub grdOPTIONS_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
a = KeyAscii: CA$ = Chr$(a)
Select Case a
  Case 13: ' Enter key
   If (NewOptValu$ <> "") Then
      RptOptUPDATE
      If Not OKAY Then
         grdOPTIONS.Row = OptRow
         grdOPTIONS_SelChange
         Exit Sub
      End If
   End If
   If grdOPTIONS.Row < RDFReport.OPTCNT Then
      OptRow = OptRow + 1
      If Not (grdOPTIONS.RowIsVisible(OptRow)) Then
         grdOPTIONS.TopRow = grdOPTIONS.TopRow + 1
      End If
      grdOPTIONS.Row = OptRow
   End If
  Case 27: ' Escape key
      cmdCANCEL_Click
  Case Else:
   If OptVLth = 0 Then Exit Sub
   If (EDTIX > 20) Then
      If ((KeyAscii < 48) Or (KeyAscii > 57)) And (CA$ <> "+") Then
         Call ErrBEEP(0): Exit Sub
      End If
   End If
   If Len(NewOptValu$) >= OptVLth Then
      Call ErrBEEP(0): Exit Sub
   End If
   NewOptValu$ = NewOptValu$ + CA$
   grdOPTIONS.TEXT = Right$("    " + NewOptValu$, 5)
   Exit Sub
End Select
grdOPTIONS_SelChange
End Sub


Private Sub grdOPTIONS_RowColChange()
  If (grdOPTIONS.Row < 1) And (grdOPTIONS.Enabled) Then
     grdOPTIONS.Row = 1
  End If
  grdOPTIONS_SelChange
End Sub


Private Sub grdOPTIONS_SelChange()
If grdOPTIONS.Enabled = False Then Exit Sub
grdOPTIONS.Enabled = False
If OptRow = 0 Then
   OptRow = 1
   grdOPTIONS.Row = 1
End If
NextRow = grdOPTIONS.Row
grdOPTIONS.Col = 2
If NewOptValu$ <> "" Then Call RptOptUPDATE
OptRow = NextRow
grdOPTIONS.Row = OptRow
Call OptCellINIT
grdOPTIONS.SelStartRow = OptRow
grdOPTIONS.SelStartCol = 2
grdOPTIONS.SelEndRow = OptRow
grdOPTIONS.SelEndCol = 2
grdOPTIONS.Enabled = True
grdOPTIONS.SetFocus
End Sub

Private Sub PRTDest_CLICK()
TurnOffOpts
RFDest.Enabled = False: RFDest.TEXT = ""
cmdRFBrowse.Enabled = False
cmdRFView.Enabled = False
PRT.OPT = PRTDest.ListIndex + 1
Select Case PRTDest.ListIndex
   Case 0: PRT.DEST = "$PRTSCRN.RPT"
   Case 1: PRT.DEST = PDFSetup.DEST
   Case 2: PRT.DEST = FileDEST$
           cmdRFBrowse.Enabled = True
           cmdRFView.Enabled = True
           RFDest.Enabled = True
           RFDest.TEXT = PRT.DEST
End Select
End Sub

Private Sub PRTSetup_Click()
TurnOffOpts
If PRTSetup.TEXT <> PRTCode Then
   If PRTSetup.TEXT = "" Then PRTSetup.TEXT = PRTCode
   Call PDFGetSet(PRTSetup.TEXT)
   If OKAY Then
      PRTDesc.TEXT = PDFSetup.DESC
      PRTCode = PDFSetup.NAME
      Call BuildPRTDest
      Else
      PRTDesc.TEXT = ""
   End If
End If
If PRTDest.TEXT <> PDFSetup.DEST Then
   Call BuildPRTDest
End If
End Sub


Private Sub RFDest_Change()
TurnOffOpts
If RFDest.Enabled Then
   X = RFDest.SelStart
   If InStr(RFDest.TEXT, ".") = 0 Then
      RFDest.TEXT = RTrim$(RFDest.TEXT) + ".RPT"
   End If
   RFDest.SelStart = X
End If
End Sub

Private Sub RFDest_GotFocus()
RFDest.SelStart = 0
RFDest.SelLength = Len(RFDest.TEXT)
End Sub


Private Sub RFDest_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub RFDest_LostFocus()
FileDEST$ = RFDest.TEXT
End Sub

Public Sub RPSetup_Click()
TurnOffOpts
If RPSetup.TEXT = "" Then
   X = 0: FOUND = 0
   While (X < RPSetup.ListCount) And (Not FOUND)
      FOUND = (RptCODE = RPSetup.List(X))
      If Not FOUND Then X = X + 1
   Wend
   If FOUND Then RPSetup.ListIndex = X
End If
RptCODE = RPSetup.TEXT
Call RptLOAD(RptCODE)
If OKAY Then
   RPDesc.TEXT = RDFReport.NAME
   Call BuildOptGrid
   Else
   RPDesc.TEXT = ""
End If
End Sub


Private Sub cmdWorkFile_Click()
TurnOffOpts
Select Case cmdWorkFile.Caption
   Case "WorkFile": Call frmFHSWRelW.BUILD(WRelType)
   Case "Select": frmFHSWSlct.Show
   Case "NodeFile": frmFHSWNdWk.Show
End Select
End Sub


