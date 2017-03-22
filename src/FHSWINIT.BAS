Attribute VB_Name = "modFHSWInit"
DefInt A-Z
Dim ATR(11), FG(11), BG(11), MT$(12)

Sub Main()
PN$ = " FAMILY HISTORY SYSTEM EXTENSIONS - FHSINIT - Initialization Program"
CY$ = " (C) Copyright 1985-1998 by Phillip E. Brown"
 ' Prepare to get DATA from FHSWINIT.DAT
   On Error GoTo 7
   Open "FHSWINIT.DAT" For Input As #5
   GoTo 8
7  MsgBox "Error Code=" + Str$(ERR) + ", Line=" + Str$(ERL)
   Exit Sub
8  Close #5
   Call DataINIT("INIT")
   ReDim HLIT$(10)

10 CONFIG.DUR = 2: MO = 0: ENV.SYSMODE = 0
   SysVAR.PGMNAME = "FHSINIT"
   ENV.PGMNAME = "FHSINIT": ENV.PGMSFX = ".EXE"
   ENV.PDF = " ": ENV.FDF = " "
   PGMPath = CurDir + "\"
   If Environ$("FHSMEM") = "LARGE" Then ENV.LARGEMEM = -1
   ' Get Configuration Information
   CFGNAME$ = Environ$("FHSCONFG")
   If CFGNAME$ = "" Then CFGNAME$ = "FHSCONFG"
   Call GetCONFIG(CFGNAME$)
   FHSFILE$ = Environ$("FHSFILES")
   If FHSFILE$ <> "" Then
      LSet ENV.FDF = CONFIG.FDF
      CONFIG.FDF = FHSFILE$
    End If
   FHSPRINT$ = Environ$("FHSPRINT")
   If FHSPRINT$ <> "" Then
      LSet ENV.PDF = CONFIG.FDF
      CONFIG.PDF = FHSPRINT$
   End If
   LANG$ = Environ$("FHSLANG"): SL = Len(LANG$)
   S1 = 0: S2 = InStr(LANG$, ":"): If S2 = 0 Then S2 = SL + 1
   If S2 - S1 = 4 Then
      CONFIG.TDF = UCase$(Left$(LANG$, 3))
      CONFIG.SDF = CONFIG.TDF: CONFIG.MDF = CONFIG.TDF
      CONFIG.RDF = CONFIG.TDF
   End If
   S1 = S2: S2 = InStr(S1 + 1, LANG$, ":"): If S2 = 0 Then S2 = SL + 1
   If S2 - S1 = 4 Then CONFIG.TDF = UCase$(MID$(LANG$, S1 + 1, 3))
   S1 = S2: S2 = InStr(S1 + 1, LANG$, ":"): If S2 = 0 Then S2 = SL + 1
   If S2 - S1 = 4 Then CONFIG.SDF = UCase$(MID$(LANG$, S1 + 1, 3))
   S1 = S2: S2 = InStr(S1 + 1, LANG$, ":"): If S2 = 0 Then S2 = SL + 1
   If S2 - S1 = 4 Then CONFIG.MDF = UCase$(MID$(LANG$, S1 + 1, 3))
16 dfmt$ = Environ$("FHSDATE"): If dfmt$ = "" Then GoTo 19
   If dfmt$ = "MM-DD-YYYY" Then CONFIG.DFM = "1": CONFIG.VER = -1
   If dfmt$ = "DD.MM.YYYY" Then CONFIG.DFM = "2": CONFIG.VER = -1
   If dfmt$ = "YYYY.MM.DD" Then CONFIG.DFM = "3": CONFIG.VER = -1
19 BPLACE$ = Space$(22): DPLACE$ = BPLACE$: MPLACE$ = BPLACE$
   PLINE$ = Space$(255): CLINE$ = PLINE$
   WKF$ = Space$(14)
   NFMT$ = Space$(52): BNAME$ = NFMT$:
   XRELATION$ = Space$(12): C$ = " "
   ' Create strings for determining printer character size
   StdWidth = Space$(224)
   For X = 1 To 224
      Mid$(StdWidth, X, 1) = Chr$(X + 31)
   Next X
   StdHeight = String$(19, "x")
   For X = 1 To 9
      Mid$(StdHeight, 2 * X, 1) = Chr$(13)
   Next X

20 ' Build File Name Tables
   Call DataINIT("FILES")
   'Open "FHSWINIT.DAT" For Input As #5
   ENV.DSNCNT = 14
   ReDim FType$(ENV.DSNCNT + 1), FGrp$(ENV.DSNCNT + 1), FT$(ENV.DSNCNT + 1)
   For X = 0 To ENV.DSNCNT
      FType$(X) = Space$(8): FGrp$(X) = " ": FT$(X) = Space$(12)
      Input #5, xZ, Z1$, Z2$, Z3$
      LSet FType$(X) = Z1$: LSet FGrp$(X) = Z2$: LSet FT$(X) = Z3$
   Next X
   Close #5
   Call GetFileDEF

21 ' Build Printer DEF
   Call GetPrintDEF

500 ' Format Display and Branch to Main Menu
    GoSub TablesBLD: If Not OKAY Then End
    mnuFHSWMain.Show
    Exit Sub

700 ' Build SNGLKEY$() AND SKEG() - the Single Key Edit Group Table
    Call TableLOAD("SINGLKEY", Y, 0)
    Z = TABLE(Y).SIZE
    ReDim SNGLKEY$(Z)
    X = TABLE(Y).FT: SNGLKEY$(0) = Chr$(Z)
    While X > 0
       Y = Val(MID$(TABDATA$(X), 1, 2))
       SNGLKEY$(Y) = MID$(TABDATA$(X), 3, 1)
       X = TABDATACHN(X).FWD
    Wend
    Call DataINIT("SKEGTAB")
    'Open "FHSWINIT.DAT" For Input As #5
    'Seek #5, XTBGN(3)
    Input #5, Z
    ReDim SKEG(Z)
    For X = 0 To Z: Input #5, SKEG(X): Next X
    Close #5
    Return

' Create TABLES Tables
TablesBLD:
    ENV.TDF = CONFIG.TDF
    Call DataINIT("TABLES")
    'Open "FHSWINIT.DAT" For Input As #5
    'Seek #5, XTBGN(6)
    Input #5, TABMAX
    ReDim XTDF(TABMAX) As TableFileDIR, TABLE(TABMAX) As TablePARMS
    ReDim SNGLKEY$(0)
    For X = 1 To TABMAX
        Input #5, TDFDir.NAME, TDFDir.EntCNT, TDFDir.XL
        XTDF(X) = TDFDir
    Next X
    Close #5
    ReDim TDFIndex(TABMAX) As TableFileDIR
    For X = 1 To TABMAX: TDFIndex(X) = XTDF(X): Next X
    Call TableLOAD("INITIAL", T, 0)
    If Not OKAY Then Return
    For X = 1 To TABMAX
        If (TDFIndex(X).NAME = XTDF(X).NAME) Or (TDFIndex(X).NAME = "") Then
           TDFIndex(X).NAME = XTDF(X).NAME
           TDFIndex(X).XL = XTDF(X).XL
        End If
    Next X
    Call TabINIT
    GoSub 700    ' Build SNGLKEY$(), SKEG()

    ' Build MDFIndex(), SDFIndex()
    Call DataINIT("GROUPS")
    'Open "FHSWINIT.DAT" For Input As #5
    'Seek #5, XTBGN(7)
    Input #5, GRPMAX
    ReDim MDFIndex(GRPMAX) As GroupPARMS
    For X = 1 To GRPMAX
       Input #5, Group$
       MDFIndex(X).NAME = Left$(Group$, 4)
    Next X
    Close #5
    Call MsgGrpLOAD("INITIAL", 0)
    If Not OKAY Then Return
    'Call FmtGrpLOAD("INITIAL", 0)
    'If Not OKAY Then Return

    ' Build RDFIndex()
    Call DataINIT("REPORTS")
    'Open "FHSWINIT.DAT" For Input As #5
    'Seek #5, XTBGN(8)
    Input #5, RPTMAX
    ReDim RDFIndex(RPTMAX) As ReportINDEX, RDFReports(RPTMAX) As ReportDefREC
    ReDim RPTOPTS(RPTMAX) As RptOPTIONS, PagePARMS(RPTMAX) As PagePARMS
    ReDim RPTOPT$(0), OPTCHN(0) As ChainPTRS
    ReDim RMONTB$(0), RGENDR$(0), RRELCD$(0)
    OPTTable.TABNAME = "OPTIONS": OPTTable.FMTNAME = "XXXXO005"
    OPTTable.KEYNAME = "KEY"
    OPTTable.XL = 37: OPTTable.ML = OPTTable.XL
    For X = 1 To RPTMAX
       Input #5, REPORT$
       RDFIndex(X).CODE = Left$(REPORT$, 4)
    Next X
    Close #5
    Call RptLOAD("INITIAL")
    Call SetCalendarType(0, 0) ' Build Calendar Tables
    OKAY = -1
    Return
End Sub

Rem $DYNAMIC
Sub GetCONFIG(CFGNAME$)
'SHARED ATR()
   xC$ = " ": ReDim CFGAL$(5)
   On Local Error GoTo 90
   Open CFGNAME$ + ".CFG" For Input As #1
   Close #1
   Open CFGNAME$ + ".CFG" For Binary Access Read As #1
   Get #1, 1, CONFIG
   If (CONFIG.CFG = "FHSCONFG") And (CONFIG.VER = 1) Then
      CONFIG.CFG = CFGNAME$
      If Asc(CONFIG.VCMD) = 0 Then CONFIG.VCMD = "EDIT"
      For X = 1 To 5
         Get #1, , xC$
         L = Asc(xC$)
         If L > 0 Then
            CFGAL$(X) = String$(L, 32)
            Get #1, , CFGAL$(X)
         End If
      Next X
      If CONFIG.SCOLORS = String$(12, 0) Then GoTo INITATR
      For I = 0 To 11
          SATTR = Asc(MID$(CONFIG.SCOLORS, I + 1, 1))
          FG(I) = FNFG(SATTR): BG(I) = FNBG(SATTR):
      Next I
      GoTo 50
   End If
   ' Initialise CONFIG Values if File not found
   CONFIG.CFG = "INVALID"
INITCFG:
   CONFIG.FDF = "FHSFILES": CONFIG.PDF = "FHSPRINT"
   CONFIG.EDA = "Y": CONFIG.DFM = "1"
   CONFIG.TDF = "FHS"
   CONFIG.SDF = CONFIG.TDF: CONFIG.MDF = CONFIG.TDF: CONFIG.RDF = CONFIG.TDF
   CONFIG.VCMD = "EDIT"
   CONFIG.BP = 63: CONFIG.DUR = 1
   On Local Error GoTo 0
INITATR:
   For I = 0 To 11
       Mid$(CONFIG.SCOLORS, I + 1, 1) = Chr$(ATR(I))
   Next I
50 Close
   If CONFIG.BP = 0 Then CONFIG.BP = 63
   BP = CONFIG.BP: DUR = CONFIG.DUR
   Exit Sub
90 CONFIG.CFG = "NOTFOUND"
   Resume INITCFG
End Sub

Rem $STATIC
Sub GetFileDEF()
    OKAY = 0
    FSUTABLE.TABNAME = "FSETUPS"
    FSUTABLE.XL = 38: FSUTABLE.ML = 38
    FSUTABLE.KEYNAME = "KEY": FSUTABLE.KO = 1: FSUTABLE.KL = 8
    On Local Error GoTo GFDNOFDF
GFDOPEN:
    Open CONFIG.FDF + ".FDF" For Input As #1
    Close
    Open CONFIG.FDF + ".FDF" For Binary Access Read As #1
    Get #1, , FDFHDr
    If (FDFHDr.FType <> "FDF") Or (FDFHDr.VER <> Chr$(0)) Then GoTo GFDINIT
    Close
    Call FDFOpen("INITQ")
    GoTo GFDEXIT
GFDNOFDF:
    If (CONFIG.FDF <> ENV.FDF) Then CONFIG.FDF = ENV.FDF: Resume GFDOPEN
    If (CONFIG.FDF <> "FHSFILES") Then CONFIG.FDF = "FHSFILES": Resume GFDOPEN
    Resume GFDINIT
GFDINIT:
    FDFHDr.FType = "XXX"
    FDFHDr.VER = Chr$(255)
    FDFHDr.SETUPS = 1
    FDFHDr.FAMDSN = 4
    FDFHDr.OTHERDSN = 10
    FDFHDr.DEFAULT = 1
    FDFHDr.EOFRBA = 0
    LSet FDFSetup.NAME = "DEFAULT"
    LSet FDFSetup.DESC = "FHS Default Values"
    LSet FDFSetup.FAMDIR = "": LSet FDFSetup.WRKDIR = ""
    LSet FDFSetup.TRANDIR = "": LSet FDFSetup.RPTDIR = ""
    FSUTABLE.FT = 1: FSUTABLE.LT = FDFHDr.SETUPS
    FSUTABLE.FF = FDFHDr.SETUPS + 1
    FSUMAX = FDFHDr.SETUPS + 5
    FSUTABLE.SIZE = FSUMAX
    ReDim FDFSURBA(FSUMAX), FSETUP$(FSUMAX), FSUCHN(FSUMAX) As ChainPTRS
    ReDim FSETUPS(FSUMAX) As FileSetupREC, FFNAME$(FDFHDr.FAMDSN * FSUMAX)
    Call ChnINIT(FSUCHN(), 1, FSUMAX)
    FSUCHN(1).FWD = 0: FSUCHN(2).BWD = 0
    FSETUPS(1) = FDFSetup
    FSETUP$(1) = Space$(FSUTABLE.ML)
    LSet FSETUP$(1) = FDFSetup.NAME + FDFSetup.DESC
    For Z = 1 To FDFHDr.FAMDSN: FFNAME$(Z) = FT$(Z): Next Z
    ENV.FSETUP = FDFHDr.DEFAULT
    OKAY = -1
GFDEXIT:
    Close
End Sub

Sub GetPrintDEF()
    OKAY = 0
    PSUTABLE.TABNAME = "PSETUPS"
    PSUTABLE.XL = 33: PSUTABLE.ML = 33
    PSUTABLE.KEYNAME = "KEY": PSUTABLE.KO = 1: PSUTABLE.KL = 8
    On Local Error GoTo GPDNOPDF
GPDOPEN:
    Open CONFIG.PDF + ".PDF" For Input As #1
    Close
    Open CONFIG.PDF + ".PDF" For Binary Access Read As #1
    Get #1, , PDFHdr
    If (PDFHdr.FType <> "PDF") Or (PDFHdr.VER <> Chr$(1)) Then GoTo GPDINIT
    Close
    Call PDFOpen("INITQ")
    GoTo GPDEXIT
GPDNOPDF:
    If (CONFIG.PDF <> ENV.PDF) Then CONFIG.PDF = ENV.PDF: Resume GPDOPEN
    If (CONFIG.PDF <> "FHSWPRNT") Then
       CONFIG.PDF = "FHSWPRNT": ENV.PDF = CONFIG.PDF
       Resume GPDOPEN
    End If
    Resume GPDINIT
GPDINIT:
    On Local Error GoTo 0
    ReDim PSETUPS(1) As PrintSetupREC
    ReDim PFONT$(1), PRINTERS$(1)
    PDFHdr.FType = "XXX"
    PDFHdr.VER = Chr$(255)
    PDFHdr.SETUPS = 1
    PDFHdr.DEFAULT = 1
    PDFHdr.EOFRBA = 0
    PDFSetup.NAME = "FHSINIT"
    PDFSetup.DESC = "FHS Default Values"
    PDFSetup.DEST = "SCRN:"
    PDFSetup.SYMBOLS = "TYPEWRIT"
    PDFSetup.SET1 = Chr$(43) + Chr$(43) + Chr$(43) + Chr$(33) + Chr$(43) + Chr$(43) + Chr$(45) + Chr$(43) + Chr$(43) + Chr$(43) + Chr$(43)
    PDFSetup.SET2 = Chr$(43) + Chr$(43) + Chr$(43) + Chr$(33) + Chr$(43) + Chr$(43) + Chr$(45) + Chr$(43) + Chr$(43) + Chr$(43)
    PDFSetup.SET3 = Chr$(43) + Chr$(43) + Chr$(43) + Chr$(33) + Chr$(43) + Chr$(43) + Chr$(45) + Chr$(43) + Chr$(43) + Chr$(43)
    PDFSetup.FWIDTH = 850
    PDFSetup.FLENGTH = 1100
    PDFSetup.TMARGIN = 75
    PDFSetup.BMARGIN = 50
    PDFSetup.IMARGIN = 70
    PDFSetup.OMARGIN = 50
    PDFSetup.OPTIONS = 3
    PFONT$(1) = "MS LineDraw"
    PDFSetup.FONTSIZE = 10
    PSETUPS(1) = PDFSetup
    PSUTABLE.TABNAME = "PSETUPS"
    'PSUTABLE.FMTNAME = "XXXXS021"
    'PSUTABLE.FMTNUM = 0
    PSUTABLE.XL = 55: PSUTABLE.ML = 55
    PSUTABLE.KO = 1: PSUTABLE.KL = 8
    PSUTABLE.FT = 1: PSUTABLE.LT = PDFHdr.SETUPS
    PSUTABLE.FF = 0
    PSUMAX = 1
    PSUTABLE.SIZE = 1
    ReDim PSUCHN(1) As ChainPTRS
    ENV.PSETUP = 1
    OKAY = -1
GPDEXIT:
    Close
End Sub

Rem $DYNAMIC
Sub TabINIT()
    Call TableLOAD("ERRMSGS", T, 0)
    EMCNT = TABLE(T).SIZE: ReDim ERRMSG$(EMCNT)
    X = TABLE(T).FT: Z = 0
    While X > 0
       ERRMSG$(Z) = RTrim$(TABDATA$(X))
       Z = Z + 1: X = TABDATACHN(X).FWD
    Wend
    Call TableFREE("ERRMSGS")
    Call TableLOAD("RELRULES", T, 0): ts = TABLE(T).SIZE
    ReDim RELRULE$(ts)
    X = TABLE(T).FT: Z = 0
    While X > 0
       Z = Val(MID$(TABDATA$(X), 1, 2))
       If Z <= ts Then
          RELRULE$(Z) = RTrim$(MID$(TABDATA$(X), 3, 12))
       End If
       X = TABDATACHN(X).FWD
    Wend
    Call TableFREE("RELRULES")
    Call TableLOAD("RELCODES", T, 0): ts = TABLE(T).SIZE
    ReDim RELCODE$(ts)
    X = TABLE(T).FT: Z = 0
    While X > 0
       Z = Val(MID$(TABDATA$(X), 1, 2))
       If Z <= ts Then
          RELCODE$(Z) = RTrim$(MID$(TABDATA$(X), 1, 17))
       End If
       X = TABDATACHN(X).FWD
    Wend
    Call TableFREE("RELCODES")
    Call TableLOAD("GENDER", T, 0): ts = TABLE(T).SIZE
    ReDim FGENDR$(ts)
    X = TABLE(T).FT: Z = 0
    While X > 0
      Z = Val(MID$(TABDATA$(X), 1, 1))
      If Z <= ts Then
         FGENDR$(Z) = MID$(TABDATA$(X), 2, 1)
      End If
      X = TABDATACHN(X).FWD
    Wend
    Call TableFREE("GENDER")
    Call TableLOAD("REFERENC", T, 0): ts = TABLE(T).SIZE
    ReDim RXIMP(ts), RXSYM$(ts)
    X = TABLE(T).FT: Z = 0
    While X > 0
      Z = Val(MID$(TABDATA$(X), 1, 2))
      If Z <= ts Then
         RXIMP(Z) = Val(MID$(TABDATA$(X), 3, 2))
      End If
      X = TABDATACHN(X).FWD
    Wend
    Call TableFREE("REFERENC")
    Call TableLOAD("UCTRAN", T, 0): ts = TABLE(T).SIZE
    UCT1$ = String$(255, 0)
    ReDim UCTRAN$(ts)
    X = TABLE(T).FT: Y = 0: Z = 0
    While X > 0
      If Y < ts Then
         Y = Y + 1
         Z = Asc(MID$(TABDATA$(X), 1, 1))
         Mid$(UCT1$, Z, 1) = Chr$(Y)
         UCTRAN$(Y) = RTrim$(MID$(TABDATA$(X), 2, 3))
      End If
      X = TABDATACHN(X).FWD
    Wend
    Call TableFREE("UCTRAN")
    Call TableLOAD("SOUNDEX", T, 0)
    SNDXTB$ = String$(255, 255)
    SXTB$ = "01230120022455012623010202"
    xC$ = " "
    For X = 1 To 26
        LSet xC$ = MID$(SXTB$, X, 1)
        Mid$(SNDXTB$, 64 + X, 1) = xC$
        Mid$(SNDXTB$, 96 + X, 1) = xC$
    Next X
    X = TABLE(T).FT
    While X > 0
       If Len(TABDATA$(X)) = 2 Then
          Z1 = Asc(MID$(TABDATA$(X), 1, 1))
          Mid$(SNDXTB$, Z1, 1) = MID$(TABDATA$(X), 2, 1)
       End If
       X = TABDATACHN(X).FWD
    Wend
    Call TableFREE("SOUNDEX")
    ' Added 04/28/08 to Support EVENTS processing in FHSWUPDT
    Call TableLOAD("EVENTS", T, 0)
End Sub

