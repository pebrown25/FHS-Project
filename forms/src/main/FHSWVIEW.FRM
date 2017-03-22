VERSION 4.00
Begin VB.Form frmFHSWView 
   Caption         =   "FHSWVIEW- View Report File"
   ClientHeight    =   2940
   ClientLeft      =   405
   ClientTop       =   3105
   ClientWidth     =   10815
   Height          =   3345
   Left            =   345
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   Top             =   2760
   Width           =   10935
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSYMBOLS 
      Caption         =   "@"
      Height          =   315
      Left            =   5300
      TabIndex        =   20
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdSTYLE 
      Caption         =   "+"
      Height          =   315
      Left            =   6480
      TabIndex        =   19
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox toPage 
      Height          =   315
      Left            =   2880
      TabIndex        =   17
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox fromPage 
      Height          =   315
      Left            =   2280
      TabIndex        =   16
      Top             =   0
      Width           =   495
   End
   Begin VB.ComboBox cboPrintWhat 
      Height          =   315
      ItemData        =   "FHSWVIEW.frx":0000
      Left            =   1080
      List            =   "FHSWVIEW.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox PrinterName 
      Enabled         =   0   'False
      Height          =   315
      Left            =   7320
      TabIndex        =   14
      Top             =   0
      Width           =   2415
   End
   Begin VB.CommandButton cmdPRINTER 
      Caption         =   ">"
      Height          =   315
      Left            =   6960
      TabIndex        =   13
      Top             =   0
      Width           =   375
   End
   Begin VB.ComboBox cboFontSize 
      Height          =   315
      ItemData        =   "FHSWVIEW.frx":002D
      Left            =   5760
      List            =   "FHSWVIEW.frx":004C
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   0
      Width           =   735
   End
   Begin VB.ComboBox cboFONT 
      Height          =   315
      Left            =   3480
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   0
      Width           =   1815
   End
   Begin VB.HScrollBar OffsetScroll 
      Height          =   315
      Left            =   6960
      TabIndex        =   8
      Top             =   2595
      Width           =   1095
   End
   Begin VB.TextBox PageOffset 
      Height          =   315
      Left            =   6240
      TabIndex        =   7
      Top             =   2595
      Width           =   615
   End
   Begin VB.HScrollBar PageScroll 
      Height          =   315
      Left            =   9360
      TabIndex        =   5
      Top             =   2595
      Width           =   1335
   End
   Begin VB.TextBox PageNumber 
      Height          =   315
      Left            =   8760
      TabIndex        =   4
      Top             =   2595
      Width           =   495
   End
   Begin VB.CommandButton cmdCANCEL 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Default         =   -1  'True
      Height          =   315
      Left            =   9840
      TabIndex        =   2
      Top             =   0
      Width           =   825
   End
   Begin VB.CommandButton cmdPRINT 
      Caption         =   "Print"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   885
   End
   Begin VB.TextBox boxTEXT 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS LineDraw"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   360
      Width           =   10575
   End
   Begin MSGrid.Grid Grid1 
      Height          =   1455
      Left            =   3120
      TabIndex        =   21
      Top             =   3120
      Width           =   2295
      _Version        =   65536
      _ExtentX        =   4048
      _ExtentY        =   2566
      _StockProps     =   77
      BackColor       =   16777215
      Enabled         =   0   'False
      Rows            =   17
      Cols            =   17
      ScrollBars      =   0
      MouseIcon       =   "FHSWVIEW.frx":0095
   End
   Begin VB.Label Label2 
      Caption         =   "--"
      Height          =   255
      Left            =   2760
      TabIndex        =   18
      Top             =   30
      Width           =   135
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   240
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblOFFSET 
      Caption         =   "Offset:"
      Height          =   255
      Left            =   5640
      TabIndex        =   10
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   5040
      TabIndex        =   9
      Top             =   2640
      Width           =   15
   End
   Begin VB.Label ViewFileStats 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   4575
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblPAGE 
      Caption         =   "Page:"
      Height          =   315
      Left            =   8160
      TabIndex        =   3
      Top             =   2640
      Width           =   495
   End
End
Attribute VB_Name = "frmFHSWView"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
DefInt A-Z
Dim CR$, LF$, FF$
Dim MAXPAGE, HIPAGE, MaxLineLth As Integer
Dim MAXLINES As Long, HILINE As Long, MaxPageLth As Long, PL1 As Long
Dim FirstLine As Long, LastLine As Long
Dim PageNo, PageOfst, FamFileOpen, RptFile$
Dim BGNPAGE#(), BGNLINE#()
Dim TextBuffer As String * 32766, BufOfst As Integer
Dim LineBuffer As String * 2048
Dim RptFileSize As Long, FileOfst As Long
Dim FontStyle As Integer
Private Sub BuildFonTable()
cboFONT.Clear: xfont = 0
For X = 1 To Printer.FontCount
   cboFONT.AddItem Printer.Fonts(X - 1)
   cboFONT.ItemData(cboFONT.NewIndex) = X - 1
   If Printer.Font = Printer.Fonts(X - 1) Then
      xfont = X - 1
   End If
Next X
X = 0: FOUND = 0
While (X < cboFONT.ListCount) And (Not FOUND)
   FOUND = (xfont = cboFONT.ItemData(X))
   If FOUND Then xfont = X Else X = X + 1
Wend
cboFONT.ListIndex = xfont
End Sub

Public Sub ReportFIL()
xx$ = "": a = 0
If (FDFSetup.RPTDRV <> " ") And (FDFSetup.RPTDRV <> Chr$(0)) Then
   xx$ = FDFSetup.RPTDRV + ":"
End If
xx$ = xx$ + RTrim$(FDFSetup.RPTDIR) + "\"
xx$ = xx$ + PRT.DEST
SysVAR.FILENAME = xx$
RptFile$ = xx$
'ViewFileName.TEXT = XX$
ViewFileStats.Caption = ""
'On Local Error GoTo BADOPEN
GoTo OpenRPTFile
BADOPEN: ' Error Opening Print
   Call ERRMessage("XXXXM030", ERR, ERL)
   a = 27: CA$ = Chr$(27): OKAY = 0
   Resume ViewExit
OpenRPTFile:
   Open RptFile$ For Binary Access Read As #4
   FamFileOpen = ENV.FFOPEN
   Show
   boxtext.MousePointer = 11
   PageOfst = 0
   Call ScanFile
   FromPage.TEXT = Str$(1): ToPage.TEXT = Str$(HIPAGE)
ViewExit:
   Call FamCLOSE
   boxtext.MousePointer = 0
   On Local Error GoTo 0
   If a = 27 Then Unload frmFHSWView
End Sub


Private Sub ScanFile()
On Local Error GoTo 0
MAXPAGE = 0: MAXLINES = 0
BufOfst = 1: FileOfst = 0: RptFileSize = LOF(4)
HILINE = 0: HIPAGE = 0: BOL = 1: PageLine = 0
MaxLineLth = 0: MaxPageLth = 0: PL1 = 0
BUFSIZE = Len(TextBuffer)
BgnOfPage = -1: LastBuf = 0
CRLF$ = Chr$(13) + Chr$(10): FF$ = Chr$(12)
While FileOfst < RptFileSize
   Get #4, FileOfst + 1, TextBuffer
   LastBuf = ((RptFileSize - FileOfst) <= BUFSIZE)
   BOL = 1: EOL = InStr(1, TextBuffer, CRLF$)
   If (EOL > 0) Then
      xCR = -1: xLF = -1: EOL = EOL + 2
      Else
      xCR = 0: xLF = 0
      If LastBuf Then EOL = RptFileSize - FileOfst + 1
   End If
   While EOL <> 0
      If (FileOfst + EOL) > RptFileSize Then
         EOL = RptFileSize - FileOfst
      End If
      EOP = InStr(MID$(TextBuffer, BOL, EOL - BOL), FF$)
      If (FileOfst + EOP) > RptFileSize Then EOP = 0
      If EOP > 0 Then
         EOL = BOL + EOP
         xCR = 0: xLF = 0
         BgnOfPage = -1
         GoSub AddLine
         If (HILINE - PL1) > MaxPageLth Then
            MaxPageLth = HILINE - PL1
         End If
         PL1 = HILINE - 1
         Else
         If BgnOfPage Then GoSub AddPage
         If xLF Then GoSub AddLine
      End If
      BOL = EOL
      EOL = InStr(BOL, TextBuffer, CRLF$)
      If (EOL > 0) Then
         xCR = -1: xLF = -1: EOL = EOL + 2
         Else
         xCR = 0: xLF = 0
         If LastBuf And ((FileOfst + BOL) < RptFileSize) Then
            EOL = RptFileSize - FileOfst + 1
         End If
      End If
   Wend
   If LastBuf Then
      FileOfst = RptFileSize
      Else
      FileOfst = FileOfst + BOL - 1
   End If
Wend
BOL = 0: FileOfst = RptFileSize
GoSub AddPage: GoSub AddLine
Close #4
If MaxPageLth = 0 Then MaxPageLth = HILINE
HIPAGE = HIPAGE - 1: HILINE = HILINE - 1: ShowStats
PageScroll.Min = 1: PageScroll.Max = HIPAGE
PageScroll.SmallChange = 1
PageScroll.LargeChange = 5
PageScroll.VALUE = 1
OffsetScroll.Min = 0: OffsetScroll.Max = MaxLineLth
OffsetScroll.SmallChange = 10
OffsetScroll.LargeChange = 100
OffsetScroll.VALUE = 0
Exit Sub

AddPage:
   HIPAGE = HIPAGE + 1
   If HIPAGE > MAXPAGE Then
      MAXPAGE = MAXPAGE + 100
      ReDim Preserve BGNPAGE#(MAXPAGE + 1)
   End If
   BGNPAGE#(HIPAGE) = FileOfst + BOL
   BgnOfPage = 0: PageLine = 0
   ShowStats
   Return
AddLine:
   HILINE = HILINE + 1
   If HILINE > MAXLINES Then
      MAXLINES = MAXLINES + 2000
      ReDim Preserve BGNLINE#(MAXLINES + 1)
   End If
   BGNLINE#(HILINE) = FileOfst + BOL
   'xyz$ = MID$(TextBuffer, BGNLINE#(HILINE - 1), BGNLINE#(HILINE) - BGNLINE#(HILINE - 1))
   'PageLine = PageLine + 1
   'If PageLine >= PRT.FLENGTH Then
   '   BgnOfPage = -1
   'End If
   X = EOL - BOL + xCR + xLF + (EOP > 0)
   If X > MaxLineLth Then MaxLineLth = X
   Return
End Sub

Private Sub ShowFontSIZE()
X = Printer.Font.SIZE
If X = 0 Then X = 8
cboFontSize.ListIndex = ((X * 100 - (X * 100 Mod 100)) / 100) - 4
End Sub

Private Sub ShowPrinterParms()
CRLF$ = Chr$(13) + Chr$(10)
xx$ = "Printer: Driver Name=" + Printer.DriverName
xx$ = xx$ + ", Device=" + Printer.DeviceName + CRLF$
xx$ = xx$ + "  FontName=" + Printer.FontName
xx$ = xx$ + ", FontSize=" + fnx$(Printer.FONTSIZE) + CRLF$
xx$ = xx$ + "Font Count=" + fnx$(Printer.FontCount) + CRLF$
For X = 1 To Printer.FontCount
   xx$ = xx$ + "   " + Str$(X) + "  " + Printer.Fonts(X) + CRLF$
Next X
boxtext.TEXT = xx$
boxtext.Refresh
End Sub

Private Sub ShowStats()
xx$ = "Total of:" + FNXL$(HILINE)
xx$ = xx$ + " Lines, " + fnx$(HIPAGE)
xx$ = xx$ + " Pages; "
xx$ = xx$ + "Page Width=" + fnx$(MaxLineLth)
xx$ = xx$ + ", Page Length=" + FNXL$(MaxPageLth)
xx$ = xx$ + ";  File Size=" + LTrim$(Str$(RptFileSize))
ViewFileStats.Caption = xx$
ViewFileStats.Refresh
End Sub

Private Sub ViewPage(Pgnum)
Dim xn As Long
CRLF$ = Chr$(13) + Chr$(10)
Open RptFile$ For Binary Access Read As #4
FirstLine = 1: FOUND = 0
While (FirstLine <= HILINE) And (Not FOUND)
   FOUND = ((BGNLINE#(FirstLine) = BGNPAGE#(Pgnum)))
   If Not FOUND Then FirstLine = FirstLine + 1
Wend
If FOUND Then
   FOUND = 0: LastLine = FirstLine - 1
   While (LastLine <= HILINE) And (Not FOUND)
      If Not FOUND Then LastLine = LastLine + 1
      FOUND = ((BGNLINE#(LastLine) = BGNPAGE#(Pgnum + 1)))
   Wend
End If
If FOUND Then
   BufOfst = 0: LastLine = LastLine - 1
   For xn = FirstLine To LastLine
       FileOfst = BGNLINE#(xn) + PageOfst
       LINELTH# = BGNLINE#(xn + 1) - FileOfst
       If LINELTH# < 1 Then
          LL = 2: xx$ = CRLF$
          Else
          LONGLINE = (LINELTH# > 2046)
          If LONGLINE Then LL = 2046 Else LL = LINELTH#
          xx$ = Space$(LL)
          Get #4, FileOfst, xx$
          z = 1
          While (z > 0) And (z < LL)
             z = InStr(z, xx$, Chr$(0))
             If z > 0 Then Mid$(xx$, z, 1) = " "
          Wend
          If LONGLINE Then xx$ = xx$ + CRLF$: LL = LL + 2
       End If
       If LL > (Len(TextBuffer) - BufOfst) Then
          LL = (Len(TextBuffer) - BufOfst)
       End If
       If (LL > 0) And (MID$(xx$, 1, 1) <> Chr$(12)) Then
          Mid$(TextBuffer, BufOfst + 1, LL) = xx$
          BufOfst = BufOfst + LL
          Else
          xn = LastLine
       End If
   Next xn
   boxtext.TEXT = MID$(TextBuffer, 1, BufOfst)
   Else
   If (Pgnum > 0) And (Pgnum <= HIPAGE) Then
      PLTH# = BGNPAGE#(Pgnum + 1) - BGNPAGE#(Pgnum)
      If PLTH# > Len(TextBuffer) Then
         PLTH# = Len(TextBuffer)
      End If
      Get #4, BGNPAGE#(Pgnum), TextBuffer
      boxtext.TEXT = MID$(TextBuffer, 1, PLTH#)
   End If
End If
boxtext.Refresh
PageNo = Pgnum
PageNumber.TEXT = SFILL$(PageNo, 6)
'PageScroll.VALUE = PageNo
PageOffset.TEXT = SFILL$(PageOfst, 6)
Close #4
End Sub

Private Sub cboFONT_Click()
Printer.Font = Printer.Fonts(cboFONT.ItemData(cboFONT.ListIndex))
boxtext.Font = Printer.Font
If Grid1.Visible = True Then
   Grid1.Font = boxtext.Font
End If
Call cboFontSize_Click
End Sub


Private Sub cboFontSize_Click()
Printer.Font.SIZE = cboFontSize.ListIndex + 4
boxtext.FONTSIZE = Printer.Font.SIZE
'boxtext.Refresh
End Sub


Private Sub cboPrintWhat_Click()
Select Case cboPrintWhat.ListIndex
   Case 2: FromPage.Enabled = True
           ToPage.Enabled = True
           FromPage.SetFocus
   Case Else: FromPage.Enabled = False
           ToPage.Enabled = False
End Select
End Sub


Private Sub cmdCANCEL_Click()
Close
Unload frmFHSWView
End Sub



Private Sub cmdPrint_Click()
a = 0: PRT.OPT = 2: Call PrintINIT
xpgnum = Val(PageNumber.TEXT)
Select Case cboPrintWhat.ListIndex
   Case 0: B = 1: E = HIPAGE
   Case 1: B = Val(PageNumber.TEXT)
           E = B
   Case 2: B = Val(FromPage.TEXT)
           E = Val(ToPage.TEXT)
End Select
For X = B To E
    ViewPage (X)
    Printer.Print boxtext.TEXT
    Printer.NewPage
    If a = 27 Then X = E
Next X
Printer.EndDoc: PRT.OPT = 1
ViewPage (xpgnum)
'ShowPrinterParms
End Sub

Private Sub cmdPRINTER_Click()
CommonDialog.ShowPrinter
PrinterName.TEXT = Printer.DeviceName
Call BuildFonTable
Call ShowFontSIZE
End Sub





Private Sub cmdSTYLE_Click()
FontStyle = (FontStyle + 1) Mod 4
Printer.Font.Bold = ((FontStyle And 1) > 0)
Printer.Font.Italic = ((FontStyle And 2) > 0)
boxtext.Font.Bold = Printer.Font.Bold
boxtext.Font.Italic = Printer.Font.Italic
End Sub


Private Sub cmdSYMBOLS_Click()
If Grid1.Visible = False Then
   Grid1.Font = boxtext.Font
   Grid1.Visible = True
   Grid1.Refresh
   Else
   Grid1.Visible = False
End If
End Sub

Private Sub Form_Load()
FormLoaded(1) = FormLoaded(1) Or ViewForm
frmFHSWView.Enabled = False
frmFHSWView.TOP = GetSetting("FHSWIN", "frmVIEW", "Top", 2000)
frmFHSWView.Left = GetSetting("FHSWIN", "frmVIEW", "Left", 500)
frmFHSWView.Height = GetSetting("FHSWIN", "frmVIEW", "Height", 3060)
frmFHSWView.WIDTH = GetSetting("FHSWIN", "frmVIEW", "Width", 9945)
frmFHSWView.WindowState = GetSetting("FHSWIN", "frmVIEW", "State", 2)
Call Form_Resize
PrinterName.TEXT = Printer.DeviceName
Call ShowFontSIZE
Call BuildFonTable
cboPrintWhat.ListIndex = 1
frmFHSWView.Enabled = True
Grid1.WIDTH = 3590: Grid1.Height = 4100
For X = 0 To 16: For Y = 0 To 16
   If Y = 0 Then
      Grid1.ColWidth(X) = 200
      If X > 0 Then
         Grid1.Row = 0: Grid1.Col = X
         Grid1.TEXT = MID$("0123456789ABCDEF", X, 1)
      End If
      Else
      If X = 0 Then
         Grid1.Row = Y: Grid1.Col = 0
         Grid1.TEXT = MID$("0123456789ABCDEF", Y, 1)
         Else
         Grid1.Row = Y: Grid1.Col = X
         Grid1.TEXT = Chr$((X - 1) * 16 + (Y - 1))
      End If
   End If
Next Y: Next X
Grid1.Visible = False: Grid1.Refresh
Grid1.TOP = cboFONT.TOP + cboFONT.Height
Grid1.Left = cboFONT.Left
End Sub

Private Sub Form_Resize()
boxtext.WIDTH = frmFHSWView.WIDTH - 345
boxtext.Height = frmFHSWView.Height - 1130
cmdCANCEL.Left = frmFHSWView.WIDTH - 1100
xTOP = frmFHSWView.Height - 700: xTOP2 = xTOP - 40
lblOFFSET.TOP = xTOP
lblOFFSET.Left = frmFHSWView.WIDTH - 5280
PageOffset.TOP = xTOP2
PageOffset.Left = frmFHSWView.WIDTH - 4680
OffsetScroll.TOP = xTOP2
OffsetScroll.Left = frmFHSWView.WIDTH - 3960
lblpage.TOP = xTOP
lblpage.Left = frmFHSWView.WIDTH - 2760
PageNumber.TOP = xTOP2
PageNumber.Left = frmFHSWView.WIDTH - 2160
PageScroll.TOP = xTOP2
PageScroll.Left = frmFHSWView.WIDTH - 1560
ViewFileStats.TOP = xTOP
Y = frmFHSWView.WIDTH - 5400
If Y > 0 Then
   ViewFileStats.WIDTH = Y
End If
'x = Len(RTrim$(ViewFileName.TEXT))
'Y = frmFHSWView.WIDTH - 2730
'If Y > 0 Then
'   ViewFileName.WIDTH = Y
'End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
Close
FormLoaded(1) = FormLoaded(1) And (255 - ViewForm)
If frmFHSWView.WindowState <> 2 Then
   SaveSetting "FHSWIN", "frmVIEW", "Top", frmFHSWView.TOP
   SaveSetting "FHSWIN", "frmVIEW", "Left", frmFHSWView.Left
   SaveSetting "FHSWIN", "frmVIEW", "Height", frmFHSWView.Height
   SaveSetting "FHSWIN", "frmVIEW", "Width", frmFHSWView.WIDTH
End If
SaveSetting "FHSWIN", "frmVIEW", "State", frmFHSWView.WindowState
End Sub


Private Sub fromPage_Change()
If Val(FromPage.TEXT) < 0 Then
   Call ErrBEEP(0)
   FromPage.TEXT = Str$(1)
End If
If Val(FromPage.TEXT) > Val(ToPage.TEXT) Then ToPage.TEXT = FromPage.TEXT
End Sub

Private Sub fromPage_GotFocus()
   FromPage.SelStart = 0
   FromPage.SelLength = Len(FromPage.TEXT)
End Sub


Private Sub Grid1_Click()
Grid1.Visible = False
End Sub

Private Sub OffsetScroll_Change()
PageOfst = OffsetScroll.VALUE
ViewPage (PageNo)
End Sub

Private Sub PageNumber_GotFocus()
    PageNumber.SelStart = 0
    PageNumber.SelLength = Len(PageNumber.TEXT)
    cmdCANCEL.DEFAULT = False
    cmdCANCEL.Cancel = False
End Sub

Private Sub PageNumber_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case 13:
        PageNumber_LostFocus
  Case Else:
        If ((KeyAscii < 48) Or (KeyAscii > 57)) Then
         Call ErrBEEP(0): Exit Sub
      End If
End Select
End Sub


Private Sub PageNumber_LostFocus()
 If (Val(PageNumber.TEXT) > 0) And (Val(PageNumber.TEXT) < HIPAGE) Then
    PageScroll.VALUE = PageNumber.TEXT
    Else
    PageNumber.TEXT = PageNo
 End If
 cmdCANCEL.DEFAULT = True
 cmdCANCEL.Cancel = True
End Sub


Private Sub PageOffset_LostFocus()
    cmdCANCEL.DEFAULT = True
    cmdCANCEL.Cancel = True
End Sub

Private Sub PageScroll_Change()
If (PageScroll.VALUE <= HIPAGE) And (PageScroll.VALUE > 0) Then
   Pgnum = PageScroll.VALUE
   ViewPage (Pgnum)
   Else
   PageNumber = Pgnum
End If
End Sub


Private Sub PageOffset_GotFocus()
    PageOffset.SelStart = 0
    PageOffset.SelLength = Len(PageOffset.TEXT)
    cmdCANCEL.DEFAULT = False
    cmdCANCEL.Cancel = False
End Sub


Private Sub PageOffset_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case 13: PageOfst = Val(PageOffset.TEXT)
           OffsetScroll.VALUE = PageOfst
  Case Else:
        If ((KeyAscii < 48) Or (KeyAscii > 57)) Then
         Call ErrBEEP(0): Exit Sub
      End If
End Select
End Sub


Private Sub toPage_Change()
If Val(ToPage.TEXT) < 0 Then ToPage.TEXT = FromPage.TEXT
If Val(ToPage.TEXT) > HIPAGE Then
   Call ErrBEEP(0)
   ToPage.TEXT = Str$(HIPAGE)
End If
If Val(ToPage.TEXT) < Val(FromPage.TEXT) Then FromPage.TEXT = ToPage.TEXT
End Sub


Private Sub toPage_GotFocus()
   ToPage.SelStart = 0
   ToPage.SelLength = Len(ToPage.TEXT)
End Sub



