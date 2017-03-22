Attribute VB_Name = "dllFHSWRun2"
DefInt A-Z
Function CVI(X$)
   x1 = Asc(Right$(X$, 1)): X2 = Asc(X$): S1 = 1: S2 = 0
   If x1 = 255 Then
      x1 = x1 Xor 255: X2 = X2 Xor 255
      S1 = -1: S2 = 1
   End If
   CVI = S1 * (x1 * 256 + X2 + S2)
End Function

Function MKI$(X)
   MKI$ = Chr$(X And 255) + Chr$(X \ 256)
End Function

Rem $DYNAMIC
Sub GetRFLD(RFNAME$, RFIX)
    RF$ = Space$(5): LSet RF$ = RFNAME$
    RFIX = RDFIndex(ENV.RPTIX).FLDBGN: FOUND = 0
    While (RFIX > 0) And (Not FOUND)
       FOUND = (RFLDS(RFIX).NAME = RF$)
       If Not FOUND Then RFIX = RFLDCHN(RFIX).FWD
    Wend
    If FOUND Then Exit Sub
    RFIX = RDFIndex(0).FLDBGN
    While (RFIX > 0) And (Not FOUND)
       FOUND = (RFLDS(RFIX).NAME = RF$)
       If Not FOUND Then RFIX = RFLDCHN(RFIX).FWD
    Wend
End Sub

Sub GetRVAR(RVNAME$, RVIX)
    RV$ = Space$(8): LSet RV$ = RVNAME$
    RVIX = RDFIndex(ENV.RPTIX).VARBGN: FOUND = 0
    While (RVIX > 0) And (Not FOUND)
       FOUND = (RVARS(RVIX).NAME = RV$)
       If Not FOUND Then RVIX = RVARCHN(RVIX).FWD
    Wend
    If FOUND Then Exit Sub
    RVIX = RDFIndex(0).VARBGN
    While (RVIX > 0) And (Not FOUND)
       FOUND = (RVARS(RVIX).NAME = RV$)
       If Not FOUND Then RVIX = RVARCHN(RVIX).FWD
    Wend
End Sub

Sub HoldPRINT()
2550 ' Interrupt printing
     If a = 27 Then Exit Sub
     'SLIN = CSRLIN: SCOL = POS(0)
     If PRT.EOF Then
        Call PutMSG("XXXX" + RDFReport.ENDMSG)
        Else
        If PAGE.BREAK Then
           Call PutMSG("XXXX" + RDFReport.PAGEMSG)
           Else
           Call PutMSG("XXXX" + RDFReport.WAITMSG)
        End If
     End If
     If Not PRT.PAGEWAIT Then PAGE.WAIT = ((Len(CA$) = 2) And (a = 81))
     If a <> 27 Then Call PutMSG("XXXX" + RDFReport.PRTMSG)
     'LOCATE SLIN, SCOL, 0
End Sub

Rem $STATIC
Sub HTFBuild(ILINE$, LL)
    Call HTFExpand(ILINE$, 1)
    XLINE$ = Space$(LL): X3 = Len(RTrim$(ILINE$))
    x1 = InStr(ILINE$, "<<"): X2 = InStr(x1 + 1, ILINE$, ">>")
    If X2 = 0 Then X2 = X3 + 1
    If x1 = 0 Then x1 = X2
    If x1 > 1 Then Mid$(XLINE$, 1, x1 - 1) = MID$(ILINE$, 1, x1 - 1)
    If X2 > x1 Then
       X4 = X3 - X2 - 1
       If X4 > 0 Then
          If X4 > LL Then X4 = LL
          Mid$(XLINE$, LL - X4 + 1, X4) = MID$(ILINE$, X3 - X4 + 1, X4)
       End If
    End If
    If (X2 > x1) Then
       X4 = X2 - x1 - 2
       If X4 > 0 Then
          If X4 > LL Then X4 = LL
          X5 = (LL - X4 + 1) / 2: If X5 < 1 Then X5 = 1
          Mid$(XLINE$, X5, X4) = MID$(ILINE$, x1 + 2, X4)
       End If
    End If
    LSet ILINE$ = XLINE$
End Sub

Sub HTFExpand(ITXT$, xType)          ' XTYPE:0 Static, 1:Dynamic
   OTXT$ = Space$(PAGE.BWIDTH + 50): VDEPTH = 0
EXPTEXT:
   VCNT = 0: LSet OTXT$ = " "
   x1 = 1: Z1 = 1: ML = Len(RTrim$(ITXT$))
NXTMVAR:
   Y1 = InStr(x1, ITXT$, "["): If Y1 = 0 Then GoTo LASTMSEG
   Y2 = InStr(Y1, ITXT$, "]"): If Y2 = 0 Then GoTo LASTMSEG
   Mid$(OTXT$, Z1, Y1 - x1) = MID$(ITXT$, x1, Y1 - x1)
   Z1 = Z1 + Y1 - x1: x1 = Y2 + 1: VCNT = VCNT + 1
   MVAR$ = MID$(ITXT$, Y1 + 1, Y2 - Y1 - 1)
   FOUND = -1
   Select Case MVAR$
      Case "MVAR": If xType = 1 Then MVAL$ = RTrim$(SysVAR.MVAR) Else GoTo NOVAR
      Case "NVAR1": If xType = 1 Then MVAL$ = LTrim$(Str$(SysVAR.NVAR1)) Else GoTo NOVAR
      Case "NVAR2": If xType = 1 Then MVAL$ = LTrim$(Str$(SysVAR.NVAR2)) Else GoTo NOVAR
      Case "NVAR3": If xType = 1 Then MVAL$ = LTrim$(Str$(SysVAR.NVAR3)) Else GoTo NOVAR
      Case "NVAR4": If xType = 1 Then MVAL$ = LTrim$(Str$(SysVAR.NVAR4)) Else GoTo NOVAR
      Case "NVAR5": If xType = 1 Then MVAL$ = LTrim$(Str$(SysVAR.NVAR5)) Else GoTo NOVAR
      Case "NVAR6": If xType = 1 Then MVAL$ = LTrim$(Str$(SysVAR.NVAR6)) Else GoTo NOVAR
      Case "PAGE": If xType = 1 Then MVAL$ = LTrim$(Str$(PAGE.COUNT)) Else GoTo NOVAR
      Case "LINE": If xType = 1 Then MVAL$ = LTrim$(Str$(PAGE.LINE)) Else GoTo NOVAR
      Case "FAMFILE": MVAL$ = RTrim$(Left$(FSETUP$(ENV.FSETUP), 8))
      Case "FILENAME": MVAL$ = RTrim$(SysVAR.FILENAME)
      Case "INDEX": MVAL$ = RTrim$(FT$(4))
      Case "PGMNAME": MVAL$ = RTrim$(SysVAR.PGMNAME)
      Case "RPTNAME": MVAL$ = RTrim$(RDFReport.NAME)
      Case "SETUP": If xType = 1 Then MVAL$ = RTrim$(SysVAR.SETUP) Else GoTo NOVAR
      Case "SUBJECT": If xType = 1 Then MVAL$ = RTrim$(SysVAR.SUBJECT)
      Case "TITLE": If xType = 1 Then MVAL$ = RTrim$(ENV.RTITLE)
      Case "TABNAME": If xType = 1 Then MVAL$ = RTrim$(SysVAR.TABNAME) Else GoTo NOVAR
      Case "DATE": MVAL$ = SysVAR.DATE
      Case "TIME": MVAL$ = Time$
      Case "PSFX": MVAL$ = ENV.PGMSFX
      Case "CHARTID": MVAL$ = SysVAR.CHARTID
      Case Else: FOUND = 0
   End Select
   If FOUND Then GoTo MOVESEG:
   If Left$(MVAR$, 3) = ".AL" Then
      X = Val(MID$(MVAR$, 4, 1))
      If (X < 1) Or (X > 5) Then GoTo NOVAR
      MVAL$ = CFGAL$(X)
      GoTo MOVESEG
   End If
   If Left$(MVAR$, 3) = ".PP" Then
      If xType > 0 Then
         X = Val(MID$(MVAR$, 4, Len(MVAR$) - 3))
         If (X > 0) And (X < 256) Then Z1 = X
         GoTo NXTMVAR
         Else
         GoTo NOVAR
      End If
   End If
   If Left$(MVAR$, 3) = ".RS" Then
      X = InStr(MVAR$, " "): If X = 0 Then GoTo NOVAR
      If X > 4 Then NN = Val(MID$(MVAR$, 4, X - 4)) Else NN = 0
      While MID$(MVAR$, X + 1, 1) = " ": X = X + 1: Wend
      xX$ = MID$(MVAR$, X + 1, Len(MVAR$) - X): X = Len(xX$)
      If (X > 0) And (NN = 0) Then NN = PAGE.BWIDTH \ X
      While NN > 0: MVAL$ = MVAL$ + xX$: NN = NN - 1: Wend
      GoTo MOVESEG
   End If
   'CHECK for Report Variables
   If Left$(MVAR$, 3) = ".RV" Then
      xX$ = MID$(MVAR$, 5, Len(MVAR$) - 4)
      Call GetRVAR(xX$, X)
      If FOUND Then MVAL$ = RVAR$(X): GoTo MOVESEG
   End If
   'CHECK for Report Fields
   If Left$(MVAR$, 3) = ".RF" Then
      Y = InStr(MVAR$, " ")
      xX$ = MID$(MVAR$, Y + 1, Len(MVAR$) - Y)
      Call GetRFLD(xX$, X)
      If Not FOUND Then GoTo NOVAR
      RDFFld = RFLDS(X)
      If Y > 4 Then xX$ = MID$(MVAR$, 4, 1) Else xX$ = "V"
      l1 = Asc(RDFFld.ALTH): L2 = Asc(RDFFld.LLTH): L3 = Asc(RDFFld.HLTH)
      Select Case xX$
         Case "A": MVAL$ = MID$(RFLD$(X), 1, l1)
         Case "L": MVAL$ = MID$(RFLD$(X), l1 + 1, L2)
         Case "H": MVAL$ = MID$(RFLD$(X), l1 + L2 + 1, L3)
         Case Else: GoTo NOVAR
      End Select
      GoTo MOVESEG
   End If
   If xType < 1 Then GoTo NOVAR
   ' Check for PHRASE Variables
   Call GetRVAR(MVAR$, X)
   If FOUND Then MVAL$ = RVAR$(X): GoTo MOVESEG
NOVAR:
   MVAL$ = MID$(ITXT$, Y1, Y2 - Y1 + 1): VCNT = VCNT - 1
MOVESEG:
   Mid$(OTXT$, Z1, Len(MVAL$)) = MVAL$
   Z1 = Z1 + Len(MVAL$)
   GoTo NXTMVAR
LASTMSEG:
   LS = ML - x1 + 1
   If LS > 0 Then Mid$(OTXT$, Z1, LS) = MID$(ITXT$, x1, LS)
   LSet ITXT$ = OTXT$: VDEPTH = VDEPTH + 1
   If (VCNT > 0) And (VDEPTH < 5) Then GoTo EXPTEXT
End Sub

Sub HTFPrint(ITYPE$)
    T$ = Left$(ITYPE$, 1)
    Select Case T$
       Case "H": LB = PAGE.BLTOP: LA = PAGE.BLHEAD
       Case "T": LB = 0: LA = PAGE.BLTITLE
       Case "F": LB = PAGE.BLFOOT: LA = 0
       Case Else: Exit Sub
    End Select
    TSL$ = "   ": NOPRINT = (Right$(ITYPE$, 1) = "C")
    Z = 0: PTYPE = PAGE.COUNT Mod 2
    If Not NOPRINT Then
       While LB > 0
          Call WriteLINE
          LB = LB - 1
       Wend
    End If
    X = 0: ONLY1 = 0: S$ = " "
    While (X < HTFLMAX)
       X = X + 1
       LSet TSL$ = HTFLINE$(X)
       If MID$(TSL$, 1, 1) = T$ Then
          LSet S$ = MID$(TSL$, 2, 1)
          If (S$ = "0") Or (S$ = PRT.SECTION) Then
             Y = Val(MID$(TSL$, 3, 1))
             If ((Y = 0) Or ((Y = 1) And (PAGE.BGNPAGE = PAGE.COUNT)) Or ((Y = 2) And (PAGE.COUNT > PAGE.BGNPAGE)) Or ((Y = 3) And (PTYPE > 0)) Or ((Y = 4) And (PTYPE = 0))) Then
                If Not (ONLY1 And (Y > 1)) Then
                   If Not NOPRINT Then
                      LSet PLINE$ = MID$(HTFLINE$(X), 4, Len(HTFLINE$(X)))
                      Call HTFBuild(PLINE$, PAGE.BWIDTH)
                      If Y < 9 Then Call WriteLINE
                   End If
                   If Y < 9 Then Z = Z + 1: ONLY1 = (Y = 1)
                End If
             End If
          End If
       End If
    Wend
    If (Z > 0) Then
       Z = Z + LA
       If Not NOPRINT Then While LA > 0: Call WriteLINE: LA = LA - 1: Wend
    End If
    Select Case T$
       Case "H": PAGE.HLCNT = Z
       Case "T": PAGE.TLCNT = Z
       Case "F": PAGE.FLCNT = Z + LB
    End Select
End Sub

Rem $DYNAMIC
Sub PageBREAK()
1605 ' Page Break
     If (a = 27) Or (PAGE.LINE = 0) Then Exit Sub
     If (PAGE.COUNT < PRT.BGNPAGE) Or (PRT.CFORMS) Then PAGE.LINE = 0: Exit Sub
     HLINE$ = RTrim$(PLINE$): LSet PLINE$ = " "
     PAGE.BREAK = -1
     X = PAGE.BLENGTH - PAGE.LINE
     While X > 0: Call WriteLINE: X = X - 1: Wend
     PAGE.FOOT = -1
     If RptOPTION.PRTFOOT Then Call HTFPrint("F")
     If (a <> 27) Then
        If (((PAGE.COUNT Mod 2) + 1) And PRT.SNGLSIDE) > 0 Then
           On Local Error GoTo 1609
           If (PRT.FORMFEED) Then
              Select Case PRT.OPT
                 Case 2: Printer.NewPage
                 Case Else: Print #4, Chr$(12);
              End Select
              Else
              X = PRT.FLENGTH - PAGE.LINE
              While X > 0: Call WriteLINE: X = X - 1: Wend
           End If
        End If
     End If
1606 PAGE.FOOT = 0
     LSet PLINE$ = HLINE$
     If (((PAGE.COUNT Mod 2) + 1) And PRT.SNGLSIDE) > 0 Then
        If (PAGE.WAIT) Then Call HoldPRINT
     End If
     PAGE.BREAK = 0
     PAGE.LINE = 0
     Exit Sub

1609 ' Error during Print
     Call ERRMessage("XXXXM080", ERR, ERL)
     If a = 27 Then Resume 1606
     Resume

End Sub

Sub PageSetup()
    Printer.EndDoc
    Printer.Orientation = 1 - ((PDFSetup.OPTIONS And 32) > 0)
    Printer.Font = PFONT$(ENV.PSETUP)
    Printer.Font.SIZE = PDFSetup.FONTSIZE
    Printer.Font.Bold = ((PDFSetup.FONTSTYL And 1) > 0)
    Printer.Font.Italic = ((PDFSetup.FONTSTYL And 2) > 0)
    If PAGE.BGNPAGE = 0 Then PAGE.BGNPAGE = 1
    PAGE.TotLINES = 0
    PAGE.TOP = 0: PAGE.LINE = 0: PAGE.COUNT = PAGE.BGNPAGE - 1
    PAGE.BLTOP = LineCount(PDFSetup.TMARGIN)
    PAGE.BLBOT = LineCount(PDFSetup.BMARGIN)
    PAGE.IMARGIN = CharCount(PDFSetup.IMARGIN)
    PAGE.OMARGIN = CharCount(PDFSetup.OMARGIN)
    PRT.FWIDTH = CharCount(PDFSetup.FWIDTH)
    PRT.FLENGTH = LineCount(PDFSetup.FLENGTH)
    PRT.SNGLSIDE = ((PDFSetup.OPTIONS And 1) > 0)
    PRT.FORMFEED = ((PDFSetup.OPTIONS And 2) > 0)
    PRT.PAGEWAIT = ((PDFSetup.OPTIONS And 4) > 0)
End Sub

Rem $STATIC
Sub PDFGetSet(XSETUP$)
    OKAY = 0: X = 0: FOUND = 0
    SysVAR.MVAR = XSETUP$
    xname$ = Space$(15): LSet xname$ = XSETUP$
    If XSETUP$ = "DEFAULT" Then
       X = PDFHdr.DEFAULT
       Else
       X = PSUTABLE.FT
       While (X > 0) And (Not FOUND)
          FOUND = (xname$ = PSETUPS(X).NAME)
          If Not FOUND Then X = PSUCHN(X).FWD
       Wend
    End If
    If X = 0 Then
       Call PutMSG("XXXXM026")
       OKAY = 0
       Exit Sub
    End If
    ENV.PSETUP = X
    PDFSetup = PSETUPS(X)
    OKAY = -1
End Sub

Sub PDFOpen(OPT$)
    OKAY = 0: XOPT$ = OPT$
    QUIET = (MID$(OPT$, Len(OPT$), 1) = "Q")
    If QUIET Then XOPT$ = Left$(OPT$, Len(OPT$) - 1)
LocatePDF:
    If Not QUIET Then
       Call frmFHSWFile.LOCATE("PDF", "FILE")
       If Not OKAY Then GoTo PDFOpenEND
    End If
    LSet SysVAR.FILENAME = PGMPath + CONFIG.PDF + ".PDF"
    On Local Error GoTo NOPDF
CHKPDF:
    Open SysVAR.FILENAME For Input As #1
    Close #1
    If XOPT$ = "NEW" Then
       Call PutMSG("XXXXM002")
       If CA$ <> SNGLKEY$(1) Then GoTo PDFOpenEND
       Kill SysVAR.FILENAME
       GoTo NEWPDF
    End If
    GoTo OLDPDF
NOPDF:
    X = ERR: Close
    If XOPT$ = "NEW" Then Resume NEWPDF
    Call ERRMessage("XXXXM030", ERR, 0)
    If (a = 27) Or QUIET Then OKAY = 0: Resume PDFOpenEND
    Close
    Resume LocatePDF
OLDPDF:
    On Local Error GoTo PDFOpenERR
    Select Case XOPT$
       Case "INIT": Open SysVAR.FILENAME For Binary Access Read As #1
       Case "READ": Open SysVAR.FILENAME For Binary Access Read As #1
       Case "WRITE": Open SysVAR.FILENAME For Binary Access Read Write As #1
       Case Else: GoTo PDFOpenEND
    End Select
    Dim CHKHdr As PrintDefHDR
    Get #1, , CHKHdr
    LSet SysVAR.MVAR = "PDF"
    If CHKHdr.FType <> "PDF" Then
       Call PutMSG("XXXXM033")
       GoTo PDFOpenEND
    End If
    If CHKHdr.VER <> Chr$(1) Then
       If Not QUIET Then Call PutMSG("XXXXM037")
       GoTo PDFOpenEND
    End If
    PDFHdr = CHKHdr
    OKAY = -1
    If XOPT$ <> "INIT" Then GoTo PDFOpenEND
    PSUMAX = PDFHdr.SETUPS
    ReDim PSETUPS(PSUMAX) As PrintSetupREC, PSUCHN(PSUMAX) As ChainPTRS
    ReDim PRINTERS$(PSUMAX), PFONT$(PSUMAX)
    Call ChnINIT(PSUCHN(), 1, PSUMAX)
    PSUTABLE.FT = 1: PSUTABLE.LT = PSUMAX: PSUTABLE.CL = 0: PSUTABLE.FF = 0
    PSUTABLE.SIZE = PSUMAX
    X = 0: PL$ = " "
    While X < PDFHdr.SETUPS
       X = X + 1
       Get #1, , PDFSetup
       PSETUPS(X) = PDFSetup
       Get #1, , PL$
       PRINTERS$(X) = Space$(Asc(PL$))
       Get #1, , PRINTERS$(X)
       Get #1, , PL$
       PFONT$(X) = Space$(Asc(PL$))
       Get #1, , PFONT$(X)
    Wend
    Close
    Call PDFGetSet("DEFAULT")
    GoTo PDFOpenEND
NEWPDF:
    On Local Error GoTo PDFOpenERR
    Open SysVAR.FILENAME For Binary Access Read Write As #1
    OKAY = -1
    GoTo PDFOpenEND
PDFOpenERR:
    Call ERRMessage("XXXXM030", ERR, ERL)
    OKAY = 0
    Resume PDFOpenEND
PDFOpenEND:
    If Not OKAY Then Close Else ENV.PDF = CONFIG.PDF
    On Local Error GoTo 0
End Sub


Sub PrintCLOSE()
    If FormLoaded(9) And MSGForm Then
       Unload frmFHSWMsg
    End If
    If PRT.OPEN = Chr$(255) Then
       Select Case PRT.OPT
          Case 2: Printer.EndDoc
          Case Else: Close #4
       End Select
       PRT.OPEN = Chr$(0)
       If (PRT.OPT = 1) And (a <> 27) Then
          Call frmFHSWView.ReportFIL
       End If
    End If
End Sub

Sub PrintHEAD()
1740 ' PRINT Report Heading
     If (a = 27) Then Exit Sub
     PAGE.COUNT = PAGE.COUNT + 1
     If ((PAGE.COUNT Mod 2) > 0) Or (PRT.SNGLSIDE < 0) Then
        PAGE.MARGIN = PAGE.IMARGIN: PAGE.RMARGIN = PAGE.OMARGIN
        Else
        PAGE.MARGIN = PAGE.OMARGIN: PAGE.RMARGIN = PAGE.IMARGIN
     End If
     If Not PRT.EOF Then
        HLINE$ = RTrim$(PLINE$)
        LSet PLINE$ = " "
        Call HTFPrint("FC")
        PAGE.BLENGTH = PRT.FLENGTH + PAGE.FLCNT * RptOPTION.PRTFOOT - PAGE.BLBOT
        PAGE.TOP = -1
        If RptOPTION.PRTHEAD Then
           Call HTFPrint("H")
           Call HTFPrint("T")
        End If
        LSet PLINE$ = HLINE$
     End If
End Sub

Sub PrintTITLE()
   If PAGE.LINE = 0 Then Exit Sub
   Call HTFPrint("TC"): Call HTFPrint("FC")
   PAGE.BLENGTH = PRT.FLENGTH - PAGE.FLCNT - PAGE.BLBOT
   If (PAGE.BLENGTH - PAGE.LINE) < (PAGE.TLCNT + 2) Then
      Call PageBREAK
      Call PrintHEAD
   Else
      Call HTFPrint("T")
   End If
End Sub

Sub RDFOpen(OPT$)
    OKAY = 0: Close: ENV.FFOPEN = 0
    LSet SysVAR.FILENAME = PGMPath + CONFIG.RDF + "RPTS.RDF"
    On Local Error GoTo NORDF
CHKRDF:
    Open SysVAR.FILENAME For Input As #1
    Close #1
    If OPT$ = "NEW" Then
       Call PutMSG("XXXXM002")
       If CA$ <> SNGLKEY$(1) Then GoTo RDFOpenEND
       Kill SysVAR.FILENAME
       GoTo NEWRDF
    End If
    GoTo OLDRDF
NORDF:
    X = ERR: Close
    If OPT$ = "NEW" Then Resume NEWRDF
    Call ERRMessage("XXXXM030", X, 0)
    If a = 27 Then Resume RDFOpenEND
    Resume CHKRDF
OLDRDF:
    On Local Error GoTo RDFOpenERR
    Select Case OPT$
       Case "READ": Open SysVAR.FILENAME For Binary Access Read As #1
       Case "WRITE": Open SysVAR.FILENAME For Binary Access Read Write As #1
       Case Else: GoTo RDFOpenEND
    End Select
    If RDFHdr.FType <> "CHG" Then
       Get #1, , RDFHdr
       LSet SysVAR.MVAR = "RDF"
       If RDFHdr.FType <> "RDF" Then
          Call PutMSG("XXXXM033")
          GoTo RDFOpenEND
       End If
       If Asc(RDFHdr.VER) > 1 Then
          Call PutMSG("XXXXM037")
          If CA$ <> SNGLKEY$(1) Then GoTo RDFOpenEND
       End If
    End If
    If Asc(RDFHdr.VER) = 0 Then RDFHdr.OPTIONS = 0
    OKAY = -1
    GoTo RDFOpenEND
NEWRDF:
    On Local Error GoTo RDFOpenERR
    Open SysVAR.FILENAME For Binary Access Read Write As #1
    OKAY = -1
    GoTo RDFOpenEND
RDFOpenERR:
    Call ERRMessage("XXXXM030", ERR, ERL)
    Resume RDFOpenEND
RDFOpenEND:
    If Not OKAY Then Close Else ENV.RDF = CONFIG.RDF
    On Local Error GoTo 0
End Sub

Sub RPTFree(RNAME$)
    If ENV.LARGEMEM Then Exit Sub
    XRN$ = Space$(4): LSet XRN$ = RNAME$
    RN = RPTMAX: FOUND = 0
    While (RN > 0) And (Not FOUND)
       FOUND = RDFIndex(RN).CODE = XRN$
       If Not FOUND Then RN = RN - 1
    Wend
    If Not FOUND Then Exit Sub
    Dim RPTNDX As ReportINDEX
    RPTNDX = RDFIndex(RN)
    If RPTNDX.CHG < 0 Then Exit Sub
    X = RPTNDX.OPTBGN
    While X > 0
       RPTOPT$(X) = ""
       If OPTFREE > 0 Then OPTCHN(OPTFREE).BWD = X
       Y = OPTCHN(X).FWD
       OPTCHN(X).FWD = OPTFREE
       OPTFREE = X
       X = Y
    Wend
    If OPTFREE > 0 Then OPTCHN(OPTFREE).BWD = 0
    RPTNDX.OPTBGN = 0
    X = RPTNDX.HTFBGN
    While X > 0
       HTFTXT$(X) = ""
       If HTFFree > 0 Then HTFCHN(HTFFree).BWD = X
       Y = HTFCHN(X).FWD
       HTFCHN(X).FWD = HTFFree
       HTFFree = X
       X = Y
    Wend
    If HTFFree > 0 Then HTFCHN(HTFFree).BWD = 0
    RPTNDX.HTFBGN = 0
    X = RPTNDX.VARBGN
    While X > 0
       RVAR$(X) = ""
       If RVFree > 0 Then RVARCHN(RVFree).BWD = X
       Y = RVARCHN(X).FWD
       RVARCHN(X).FWD = RVFree
       RVFree = X
       X = Y
    Wend
    If RVFree > 0 Then RVARCHN(RVFree).BWD = 0
    RPTNDX.VARBGN = 0
    X = RPTNDX.FLDBGN
    While X > 0
       RFLD$(X) = ""
       If RFFree > 0 Then RFLDCHN(RFFree).BWD = X
       Y = RFLDCHN(X).FWD
       RFLDCHN(X).FWD = RFFree
       RFFree = X
       X = Y
    Wend
    If RFFree > 0 Then RFLDCHN(RFFree).BWD = 0
    RPTNDX.FLDBGN = 0
    RPTNDX.DEFIX = 0
    RPTNDX.OPEN = 0
    RDFIndex(RN) = RPTNDX
    OKAY = -1
End Sub

Sub RptLOAD(RNAME$)
    Dim RPTNDX As ReportINDEX
    RN$ = Space$(4): LSet RN$ = RNAME$: xZ$ = String$(5, 0): OPT$ = Space$(6)
    SysVAR.FILENAME = CONFIG.RDF + "RPTS.RDF"
    If (RNAME$ = "ALL") Or (RNAME$ = "INITIAL") Then GoTo RPTINIT
    If (CONFIG.RDF = ENV.RDF) Then GoTo RPTLKUP
    RDFCHANGE = -1
RPTINIT:
    LOADALL = ((RNAME$ = "ALL") Or ((RNAME$ = "INITIAL") And ENV.LARGEMEM))
    Call RDFOpen("READ")
    If Not OKAY Then
       RDFHdr.FType = "NEW"
       RDFHdr.VER = Chr$(255)
       RDFHdr.REPORTS = 0: RDFHdr.VARTOTL = 0: RDFHdr.FLDTOTL = 0
       RDFHdr.TABLES = 0: RDFHdr.VARIABLE = 0: RDFHdr.FIELDS = 0
       RDFHdr.HTFTOTL = 0: RDFHdr.OPTTOTL = 0: RDFHdr.OPTIONS = 0
       RDFHdr.EOFRBA = 0
       Else
       FILEOPEN = -1
    End If
    ENV.RPTIX = 0
    On Local Error GoTo RPTLERR
    For X = 1 To RPTMAX
       RPTNDX.CODE = RDFIndex(X).CODE
       RDFIndex(X) = RPTNDX
    Next X
    If LOADALL Then
       RDFMAX = RDFHdr.REPORTS
       HTFMAX = RDFHdr.HTFTOTL: OPTMAX = RDFHdr.OPTTOTL
       RVMAX = RDFHdr.VARTOTL: RFMAX = RDFHdr.FLDTOTL
       Else
       RDFMAX = 0: HTFMAX = 0: OPTMAX = RDFHdr.OPTIONS
       RVMAX = RDFHdr.VARIABLE: RFMAX = RDFHdr.FIELDS
    End If
    RDFFree = 0
    ReDim RDFReports(RDFMAX) As ReportDefREC, RDFChn(RDFMAX) As ChainPTRS
    ReDim RPTOPTS(RPTMAX) As RptOPTIONS
    ReDim RPTOPT$(OPTMAX), OPTCHN(OPTMAX) As ChainPTRS
    ReDim HTFTXT$(HTFMAX), HTFCHN(HTFMAX) As ChainPTRS
    ReDim RFLDS(RFMAX) As ReportFieldREC, RFLDCHN(RFMAX) As ChainPTRS, RFLD$(RFMAX)
    ReDim RVARS(RVMAX) As ReportVariableREC, RVARCHN(RVMAX) As ChainPTRS, RVAR$(RVMAX)
    Call ChnINIT(RDFChn(), 1, RDFMAX): RDFFree = (-(RDFMAX > 0))
    Call ChnINIT(OPTCHN(), 1, OPTMAX): OPTFREE = (-(OPTMAX > 0))
    Call ChnINIT(HTFCHN(), 1, HTFMAX): HTFFree = (-(HTFMAX > 0))
    Call ChnINIT(RFLDCHN(), 1, RFMAX): RFFree = (-(RFMAX > 0))
    Call ChnINIT(RVARCHN(), 1, RVMAX): RVFree = (-(RVMAX > 0))
    xX$ = Space$(11): RptOPTION = RPTOPTS(0)
    RptOPTION.INITCHK = -1
    Y = TABLE(ENV.OPTTAB).FT
    While Y > 0
       optp$ = MID$(TABDATA$(Y), 1, 4)
       If MID$(optp$, 1, 1) = "-" Then
          If optp$ = "-WF-" Then
             skipopt = True
             Else
             skipopt = False
          End If
       End If
       If skipopt Then
          OPTMAX = OPTMAX - 1
          Else
          Mid$(xX$, 1, 4) = optp$
          Call RptOptBLD(xX$, Y)
          Mid$(xX$, 5, 2) = MID$(xX$, 10, 2)
          Call RptOptSET(xX$, RptOPTION)
       End If
       Y = TABDATACHN(Y).FWD
    Wend
    X = 0: If Asc(RDFHdr.VER) = 0 Then Seek #1, 23
    While X < RDFHdr.REPORTS
        X = X + 1
        Get #1, , RDFDir
        Y = 0: FOUND = 0
        While (Y < RPTMAX) And (Not FOUND)
           Y = Y + 1
           FOUND = (RDFDir.CODE = RDFIndex(Y).CODE)
        Wend
        If FOUND Then
           RDFIndex(Y).CHG = RDFDir.CHG
           RDFIndex(Y).DEFRBA = RDFDir.DEFRBA
           Else
           SysVAR.MVAR = RDFDir.CODE
           Call PutMSG("XXXXM038")
        End If
    Wend
    If RDFHdr.TABLES > 0 Then
       L$ = " ": ReDim REVENT$(1)
       For X = 1 To RDFHdr.TABLES
           Get #1, , RDFTable
           If RDFTable.EntCNT > 0 Then
              Select Case RDFTable.NAME
                 Case "MONTHS  ": ReDim RMONTB$(RDFTable.EntCNT)
                                  RMONTB$(0) = Str$(RDFTable.EntCNT)
                 Case "GENDER  ": ReDim RGENDR$(RDFTable.EntCNT)
                                  RGENDR$(0) = Str$(RDFTable.EntCNT)
                 Case "RELATION": ReDim RRELCD$(RDFTable.EntCNT)
                                  RRELCD$(0) = Str$(RDFTable.EntCNT)
                 Case "REFERENC": ReDim RXSYM$(RDFTable.EntCNT)
                                  RXSYM$(0) = Str$(RDFTable.EntCNT)
                 Case "EVENTS  ": ReDim REVENT$(RDFTable.EntCNT)
                                  REVENT$(0) = Str$(RDFTable.EntCNT)
              End Select
              For Y = 1 To RDFTable.EntCNT
                  Get #1, , L$: XL = Asc(L$)
                  xX$ = Space$(XL)
                  Get #1, , xX$
                  Select Case RDFTable.NAME
                     Case "MONTHS  ": RMONTB$(Y) = xX$
                     Case "EVENTS  ": REVENT$(Y) = xX$
                     Case "GENDER  ": RGENDR$(Y) = xX$
                     Case "RELATION": RRELCD$(Y) = xX$
                     Case "REFERENC": Z = Val(MID$(xX$, 1, 2))
                                      If (Z > 0) And (Z <= RDFTable.EntCNT) Then
                                         RXSYM$(Z) = MID$(xX$, 3, 3)
                                      End If
                  End Select
              Next Y
           End If
       Next X
    End If
    If RDFHdr.VARIABLE > 0 Then
       RDFIndex(0).VARBGN = RVFree
       For X = 1 To RDFHdr.VARIABLE
           Get #1, , RDFVar
           RVARS(RVFree) = RDFVar
           RVAR$(RVFree) = Space$(Asc(RDFVar.LTH))
           Get #1, , RVAR$(RVFree)
           RVFree = RVARCHN(RVFree).FWD
       Next X
       If RVFree > 0 Then
          RVARCHN(RVFree - 1).FWD = 0
          RVARCHN(RVFree).BWD = 0
       End If
    End If
    If RDFHdr.FIELDS > 0 Then
       RDFIndex(0).FLDBGN = RFFree
       For X = 1 To RDFHdr.FIELDS
           Get #1, , RDFFld
           RFLDS(RFFree) = RDFFld
           XL = Asc(RDFFld.ALTH) + Asc(RDFFld.LLTH) + Asc(RDFFld.HLTH)
           RFLD$(RFFree) = Space$(XL)
           Get #1, , RFLD$(RFFree)
           RFFree = RFLDCHN(RFFree).FWD
       Next X
       If RFFree > 0 Then
          RFLDCHN(RFFree - 1).FWD = 0
          RFLDCHN(RFFree).BWD = 0
       End If
    End If
    If RDFHdr.OPTIONS > 0 Then
       RDFIndex(0).OPTBGN = OPTFREE
       For X = 1 To RDFHdr.OPTIONS
           Get #1, , OPT$
           RPTOPT$(OPTFREE) = OPT$ + xZ$
           Call RptOptSET(OPT$, RptOPTION)
           OPTFREE = OPTCHN(OPTFREE).FWD
       Next X
       If OPTFREE > 0 Then
          OPTCHN(OPTFREE - 1).FWD = 0
          OPTCHN(OPTFREE).BWD = 0
       End If
    End If
    RPTOPTS(0) = RptOPTION
    ENV.RDF = CONFIG.RDF
    OKAY = -1
    If LOADALL Then XRN = 1: GoTo RptLOAD
    If RNAME$ = "INITIAL" Then GoTo RPTLEND
    LSet RN$ = RNAME$: RDFCHANGE = 0
RPTLKUP:
    XRN = 0
    FOUND = 0
    While (XRN < RPTMAX) And (Not FOUND)
       XRN = XRN + 1: FOUND = (RDFIndex(XRN).CODE = RN$)
    Wend
    If (Not FOUND) Then
       SysVAR.MVAR = RN$
       Call PutMSG("XXXXM039")
       OKAY = 0: GoTo RPTLEND
    End If
    ENV.RPTIX = XRN: OKAY = -1
RptLOAD:
    If XRN > RPTMAX Then GoTo RPTLEND
    RPTNDX = RDFIndex(XRN)
    If RPTNDX.OPEN Or (RPTNDX.DEFIX > 0) Then
       If Not LOADALL Then RDFReport = RDFReports(RPTNDX.DEFIX)
       GoTo RPTLNEXT
    End If
    If RPTNDX.DEFRBA > 0 Then
       If Not FILEOPEN Then
          Call RDFOpen("READ")
          If Not OKAY Then GoTo RPTLEND
          FILEOPEN = -1
       End If
       Get #1, RPTNDX.DEFRBA, RDFReport
       Else
       If LOADALL Then GoTo RPTLNEXT
       Dim NewReport As ReportDefREC
       RDFReport = NewReport
       RDFReport.CODE = RPTNDX.CODE
       RDFReport.NAME = "*** " + RPTNDX.CODE + " ***"
       RDFReport.PRTMSG = "M015"
       RDFReport.WAITMSG = "M016"
       RDFReport.PAGEMSG = "M017"
       RDFReport.INTMSG = "M018"
       RDFReport.ENDMSG = "M019"
    End If
    If RDFFree = 0 Then
       RDFMAX = RDFMAX + 1
       ReDim Preserve RDFReports(RDFMAX) As ReportDefREC
       ReDim Preserve RDFChn(RDFMAX) As ChainPTRS
       RDFFree = RDFMAX
    End If
    RDFReports(RDFFree) = RDFReport
    RPTNDX.DEFIX = RDFFree
    RDFFree = RDFChn(RDFFree).FWD
    RPTNDX.CHG = RDFReport.CHG
    X = RDFReport.OPTCNT: Y = 0:
    If X > 0 Then
       RPTNDX.OPTBGN = OPTFREE
       While X > 0
           If OPTFREE = 0 Then
              OPTFREE = OPTMAX + 1
              OPTMAX = OPTMAX + X + 1
              ReDim Preserve RPTOPT$(OPTMAX), OPTCHN(OPTMAX) As ChainPTRS
              Call ChnINIT(OPTCHN(), OPTFREE, OPTMAX)
              If Y > 0 Then OPTCHN(Y).FWD = OPTFREE Else RPTNDX.OPTBGN = OPTFREE
           End If
           Get #1, , OPT$
           If MID$(OPT$, 1, 1) = "-" Then
              If MID$(OPT$, 1, 4) = "-WF-" Then
                 skipopt = True
                 Else
                 skipopt = False
              End If
           End If
           If skipopt Then
              RDFReport.OPTCNT = RDFReport.OPTCNT - 1
              Else
              RPTOPT$(OPTFREE) = OPT$ + xZ$
              OPTCHN(OPTFREE).BWD = Y: Y = OPTFREE
              OPTFREE = OPTCHN(OPTFREE).FWD
           End If
           X = X - 1
       Wend
       If Y > 0 Then OPTCHN(Y).FWD = 0
       RDFReports(RPTNDX.DEFIX) = RDFReport
    End If
    X = RDFReport.HTFCNT: Y = 0
    If X > 0 Then
       RPTNDX.HTFBGN = HTFFree
       While X > 0
           If HTFFree = 0 Then
              HTFFree = HTFMAX + 1
              HTFMAX = HTFMAX + X
              ReDim Preserve HTFTXT$(HTFMAX), HTFCHN(HTFMAX) As ChainPTRS
              Call ChnINIT(HTFCHN(), HTFFree, HTFMAX)
              If Y > 0 Then HTFCHN(Y).FWD = HTFFree Else RPTNDX.HTFBGN = HTFFree
           End If
           Get #1, , RDFHTFLine
           XL = Asc(RDFHTFLine.LTH)
           HTFX$ = Space$(XL)
           Get #1, , HTFX$
           HTFTXT$(HTFFree) = RDFHTFLine.TYPE + RDFHTFLine.SECTION + RDFHTFLine.USE + HTFX$
           HTFCHN(HTFFree).BWD = Y: Y = HTFFree
           HTFFree = HTFCHN(HTFFree).FWD
           X = X - 1
       Wend
       If Y > 0 Then HTFCHN(Y).FWD = 0
    End If
    X = RDFReport.VARDEFS: Y = 0
    If X > 0 Then
       RPTNDX.VARBGN = RVFree
       While (X > 0)
           If RVFree = 0 Then
              RVFree = RVMAX + 1
              RVMAX = RVMAX + X
              ReDim Preserve RVAR$(RVMAX), RVARCHN(RVMAX) As ChainPTRS
              ReDim Preserve RVARS(RVMAX) As ReportVariableREC
              Call ChnINIT(RVARCHN(), RVFree, RVMAX)
              If Y > 0 Then RVARCHN(Y).FWD = RVFree Else RPTNDX.VARBGN = RVFree
           End If
           Get #1, , RDFVar
           RVARS(RVFree) = RDFVar
           RVAR$(RVFree) = Space$(Asc(RDFVar.LTH))
           Get #1, , RVAR$(RVFree)
           RVARCHN(RVFree).BWD = Y: Y = RVFree
           RVFree = RVARCHN(RVFree).FWD
           X = X - 1
       Wend
       If Y > 0 Then RVARCHN(Y).FWD = 0
    End If
    X = RDFReport.FLDDEFS: Y = 0
    If X > 0 Then
       RPTNDX.FLDBGN = RFFree
       While (X > 0)
           If RFFree = 0 Then
              RFFree = RFMAX + 1
              RFMAX = RFMAX + X
              ReDim Preserve RFLD$(RFMAX), RFLDCHN(RFMAX) As ChainPTRS
              ReDim Preserve RFLDS(RFMAX) As ReportFieldREC
              Call ChnINIT(RFLDCHN(), RFFree, RFMAX)
              If Y > 0 Then RFLDCHN(Y).FWD = RFFree Else RPTNDX.FLDBGN = RFFree
           End If
           Get #1, , RDFFld
           RFLDS(RFFree) = RDFFld
           XL = Asc(RDFFld.ALTH) + Asc(RDFFld.LLTH) + Asc(RDFFld.HLTH)
           RFLD$(RFFree) = Space$(XL)
           Get #1, , RFLD$(RFFree)
           RFLDCHN(RFFree).BWD = Y: Y = RFFree
           RFFree = RFLDCHN(RFFree).FWD
           X = X - 1
       Wend
       If Y > 0 Then RFLDCHN(Y).FWD = 0
    End If
    RPTNDX.OPEN = -1
    RDFIndex(XRN) = RPTNDX
RPTLNEXT:
    If (LOADALL And OKAY) Then XRN = XRN + 1: GoTo RptLOAD
    GoTo RPTLEND
RPTLERR:
    SysVAR.MVAR = RNAME$
    Call ERRMessage("XXXXM040", ERR, ERL)
    OKAY = 0
    Resume RPTLEND
RPTLEND:
    On Local Error GoTo 0
    Close
End Sub

Sub RptOPEN(RNAME$)
    a = 0
    If RNAME$ <> "" Then
       Call RptLOAD(RNAME$): If Not OKAY Then GoTo RPTOEND
       Call RptOptINIT
    End If
    If (ENV.RPTIX = 0) Then GoTo RPTOEND
    Call RptOptGET("BGNP", X): PAGE.BGNPAGE = X
    Call PageSetup
    PAGE.FULLWIDTH = 0
    PAGE.BLHEAD = RDFReport.BLHEAD
    PAGE.BLTITLE = RDFReport.BLTITLE
    PAGE.BLBODY = RDFReport.BLBODY
    PAGE.BLFOOT = RDFReport.BLFOOT
    If RDFReport.PWIDTH > 0 Then
       PAGE.BWIDTH = RDFReport.PWIDTH
       Else
       PAGE.BWIDTH = PRT.FWIDTH - PAGE.IMARGIN - PAGE.OMARGIN
    End If
    If RNAME$ <> "4PRT" Then LSet ENV.RTITLE = RDFReport.NAME
    HTFLMAX = RDFReport.HTFCNT: HTFPFX$ = Space$(3)
    ReDim HTFLINE$(HTFLMAX)
    X = RDFIndex(ENV.RPTIX).HTFBGN: Y = 0
    While X > 0
       Y = Y + 1
       LSet HTFPFX$ = HTFTXT$(X)
       If HTFPFX$ = "T49" Then
          LSet ENV.RTITLE = MID$(RTrim$(HTFTXT$(X)), 4, 255)
          HTFLMAX = HTFLMAX - 1: Y = Y - 1
          Else
          HTFLINE$(Y) = RTrim$(HTFTXT$(X))
       End If
       X = HTFCHN(X).FWD
    Wend
    If RptOPTION.COPYRITE Then
       Call GetRVAR("COPYRITE", X)
       If FOUND Then
          HTFLMAX = HTFLMAX + 1
          ReDim Preserve HTFLINE$(HTFLMAX)
          XLINE$ = Space$(PAGE.BWIDTH)
          LSet XLINE$ = RVAR$(X): Call HTFExpand(XLINE$, 0)
          HTFLINE$(HTFLMAX) = "F00" + XLINE$
       End If
    End If
    RptOPTION.SNGLREF = 0
    OKAY = -1
RPTOEND:
    If a = 27 Then OKAY = 0
    If Not OKAY Then Close
End Sub
Sub RptOptBLD(OPT$, OPTIX)
9999 ' Line Marker
    EDTIX = Val(MID$(TABDATA$(OPTIX), 5, 2))
    X$ = LTrim$(MID$(TABDATA$(OPTIX), 7, 3))
    If EDTIX < 0 Then
       EDTIX = 255
       Else
       If ((EDTIX < 21) And (EDTIX > SKEG(0))) Or (EDTIX > 25) Then
          EDTIX = 0
       End If
    End If
    If EDTIX > 0 Then
       If EDTIX < 21 Then
          xX = 0: FOUND = 0
          While (xX < Asc(SNGLKEY$(0))) And (Not FOUND)
             xX = xX + 1: FOUND = (X$ = SNGLKEY$(xX))
          Wend
          If FOUND Then
             ADEF = xX
             Else
             ADEF = SKEG(SKEG(EDTIX))
          End If
          Else
          If EDTIX < 26 Then
             D = EDTIX - 20
             ADEF = Abs(Val(X$)): If ADEF > (10 ^ D - 1) Then ADEF = 0
             If X$ <> "" Then If (Asc(X$) = 43) Then ADEF = -ADEF
          End If
       End If
       Else
       ADEF = Asc(X$)
    End If
    Mid$(OPT$, 7, 2) = MKI$(OPTIX)
    Mid$(OPT$, 9, 1) = Chr$(EDTIX)
    Mid$(OPT$, 10, 2) = MKI$(ADEF)
End Sub

Sub RptOptEDIT(OPT$)
    xX$ = LTrim$(MID$(OPT$, 12, 5)): SysVAR.MVAR = ""
    EDTIX = Asc(MID$(OPT$, 9, 1))
    OKAY = -1
    If (EDTIX > 0) And (xX$ = "") Then
       Mid$(OPT$, 5, 2) = MID$(OPT$, 10, 2)
       Else
       If EDTIX < 21 Then
          If EDTIX = 0 Then SKIX = Asc(xX$)
          If EDTIX > 0 Then
             X = SKEG(EDTIX): Y = SKEG(EDTIX + 1)
             Z = 0: OKAY = 0
             While (X < Y) And (Not OKAY)
                xZ$ = SNGLKEY$(SKEG(X))
                OKAY = (xX$ = xZ$)
                If Not OKAY Then
                   X = X + 1: Mid$(SysVAR.MVAR, Z + 1, 2) = xZ$ + ","
                   Z = Z + 2
                End If
             Wend
             If Not OKAY Then
                Mid$(SysVAR.MVAR, Z, 1) = " "
                Call PutMSG("XXXXM041")
                Else
                SKIX = SKEG(X)
             End If
          End If
          Else
          w& = Val(xX$): If w& > 32767 Then w& = 32767
          SKIX = w&
          If EDTIX < 26 Then
             D = EDTIX - 20
             If SKIX > ((10 ^ D) - 1) Then
                SysVAR.MVAR = LTrim$(Str$(10 ^ D - 1))
                Call PutMSG("XXXXM042")
                OKAY = 0
                Else
                If Asc(xX$) = 43 Then SKIX = -SKIX
             End If
          End If
       End If
       If OKAY Then Mid$(OPT$, 5, 2) = MKI$(SKIX)
    End If
    Call RptOptEXPAND(OPT$)
End Sub

Sub RptOptEXPAND(OPT$)
    SKIX = CVI(MID$(OPT$, 5, 2))
    EDTIX = Asc(MID$(OPT$, 9, 1))
    xX$ = Space$(5)
    If EDTIX = 255 Then
       RSet xX$ = ""
       Else
       If EDTIX = 0 Then
          RSet xX$ = Chr$(SKIX)
          Else
          If EDTIX < 21 Then
             RSet xX$ = SNGLKEY$(SKIX)
             Else
             If (EDTIX < 26) And (SKIX < 0) Then
                RSet xX$ = "+" + LTrim$(Str$(-SKIX))
                Else
                RSet xX$ = LTrim$(Str$(SKIX))
             End If
          End If
       End If
    End If
    Mid$(OPT$, 12, 5) = xX$
End Sub

Sub RptOptFIND(OPT$)
    OPTIX = CVI(MID$(OPT$, 7, 2))
    If OPTIX = 0 Then
       EDTIX = 0: ADEF = 32
       OPTIX = TABLE(ENV.OPTTAB).FT
       Y$ = Left$(OPT$, 4)
       FOUND = 0
       While (OPTIX > 0) And (Not FOUND)
          FOUND = (Y$ = MID$(TABDATA$(OPTIX), 1, 4))
          If Not FOUND Then OPTIX = TABDATACHN(OPTIX).FWD
       Wend
       If FOUND Then
          Call RptOptBLD(OPT$, OPTIX)
          Else
          Mid$(OPT$, 7, 2) = MKI$(OPTIX)
          Mid$(OPT$, 9, 1) = Chr$(EDTIX)
          Mid$(OPT$, 10, 2) = MKI$(ADEF)
       End If
    End If
End Sub

Sub RptOptGET(OPT$, OVALU)
   OVALU = 0: XOPT$ = Space$(4): LSet XOPT$ = OPT$
   X = RDFIndex(ENV.RPTIX).OPTBGN: FOUND = 0
   While (X > 0) And (Not FOUND)
      FOUND = (XOPT$ = Left$(RPTOPT$(X), 4))
      If Not FOUND Then X = OPTCHN(X).FWD
   Wend
   If Not FOUND Then
      X = RDFIndex(0).OPTBGN
      While (X > 0) And (Not FOUND)
         FOUND = (XOPT$ = Left$(RPTOPT$(X), 4))
         If Not FOUND Then X = OPTCHN(X).FWD
      Wend
   End If
   If FOUND Then
      OVALU = CVI(MID$(RPTOPT$(X), 5, 2))
      Call RptOptSET(RPTOPT$(X), RptOPTION)
      Else
      XOPT$ = XOPT$ + String$(7, 0)
      Call RptOptFIND(XOPT$)
      If FOUND Then OVALU = CVI(MID$(XOPT$, 10, 2))
   End If
End Sub

Sub RptOptINIT()
    RptOPTION = RPTOPTS(ENV.RPTIX)
    If RptOPTION.INITCHK = 0 Then
       RptOPTION = RPTOPTS(0)
       X = RDFIndex(ENV.RPTIX).OPTBGN
       While X > 0
          Call RptOptSET(RPTOPT$(X), RptOPTION)
          X = OPTCHN(X).FWD
       Wend
       RPTOPTS(ENV.RPTIX) = RptOPTION
    End If
End Sub

Sub RptOptSAVE()
    Dim XXXINDX As ReportINDEX, XXXREPORT As ReportDefREC
    OKAY = 0: XXXINDX = RDFIndex(ENV.RPTIX)
    Call RDFOpen("WRITE")
    If OKAY Then
       If XXXINDX.DEFIX > 0 Then
          XXXREPORT = RDFReports(XXXINDX.DEFIX)
          Get #1, XXXINDX.DEFRBA, RDFReport
          OKAY = (RDFReport.CODE = XXXREPORT.CODE)
          'If OKAY Then OKAY = (RDFReport.OPTCNT = XXXREPORT.OPTCNT)
          If OKAY Then
             oc1 = RDFReport.OPTCNT: oc2 = XXXREPORT.OPTCNT
             RDFReport.CHG = (((Abs(RDFReport.CHG) And 255) + 1) And 255)
             RDFReport.OPTCNT = oc2
             Put #1, XXXINDX.DEFRBA, RDFReport
             RDFReport.OPTCNT = oc1
             Y = XXXINDX.OPTBGN: opt1$ = Space$(6): opt2$ = Space$(6)
             optrba# = XXXINDX.DEFRBA + Len(RDFReport): ocx = 0
             While (Y > 0) And (ocx < oc2)
                LSet opt1$ = RPTOPT$(Y): FOUND = 0
                While (Not FOUND) And (ocx < oc2)
                   Get #1, , opt2$
                   FOUND = (MID$(opt1$, 1, 4) = MID$(opt2$, 1, 4))
                   If Not FOUND Then ocx = ocx + 1: optrba# = optrba# + 6
                Wend
                If FOUND Then
                   If Left$(opt1$, 4) = "BGPP" Then Mid$(opt1$, 5, 2) = MKI$(1)
                   Put #1, optrba#, opt1$
                   optrba# = optrba# + 6
                End If
                Y = OPTCHN(Y).FWD
             Wend
          End If
       End If
    End If
    Close: ENV.FFOPEN = 0
End Sub

Sub RptOptSET(OPT$, XOPT As RptOPTIONS)
    O = CVI(MID$(OPT$, 5, 2))
    Select Case Left$(OPT$, 4)
       Case "AGE ": XOPT.PRAGE = -(O = 1) - 2 * (O = 3)
       Case "AREF": XOPT.BLDREF = (O = 1)
       Case "BGPP": PRT.BGNPAGE = O
       Case "BLKL": XOPT.BLNKSEP = (O = 1)
       Case "BRKT": XOPT.IGNBKT = (O = 2): XOPT.RMVBKT = (O = 1)
       Case "CMIN": PAGE.CINDENT = O: PAGE.CINDENT = 0
       Case "COMM": If O = 2 Then XOPT.COMMENTS = 0 Else XOPT.COMMENTS = O
       Case "CPRT": XOPT.COPYRITE = (O = 1)
       Case "DATE": XOPT.DateS = (O = 1)
       Case "EVPG": XOPT.EVENPAGE = (O = 1)
       Case "HEAD": XOPT.PRTHEAD = (O = 1)
       Case "FOOT": XOPT.PRTFOOT = (O = 1)
       Case "LREF": XOPT.LINEREF = (XOPT.LINEREF And 254) - (O = 1)
       Case "PLAC": XOPT.PLACES = (O = 1)
       Case "PRGL": XOPT.PRGL = (O = 1)
       Case "PRID": XOPT.PRID = (O = 1)
       Case "PIDH": XOPT.PRID = ((O = 1) - 2 * (O = 22))
       Case "REFL": XOPT.LINEREF = (XOPT.LINEREF And 1) - 2 * (O = 6) - 4 * (O = 7)
                    XOPT.LINEREF = XOPT.LINEREF - 8 * (O = 8) - 16 * (O = 9)
                    XOPT.LINEREF = XOPT.LINEREF - 6 * (O = 10)
       Case "RELA": XOPT.Relation = (O = 1)
       Case "RRUL": ENV.RELRULES = (ENV.RELRULES And 253) - (O = 2) - 2 * (O = 3)
       Case "SEX ": XOPT.PRSEX = (O = 1)
       Case "MARY": XOPT.SPOUSE = -(O = 1) - 2 * (O = 4)
       Case "SURF": XOPT.NAMEFMT = (XOPT.NAMEFMT And 254) - (O = 1)
       Case "SURU": XOPT.NAMEFMT = (XOPT.NAMEFMT And 253) - 2 * (O = 1)
       Case "SURH": ENV.HUSBSURNM = (O = 1)
    End Select
End Sub

Sub WriteLINE()
2500 ' Write REPORT line
     If a = 27 Then Exit Sub
     RPAGE = PAGE.RPAGE
     If (PAGE.LINE = 0) And (Not PAGE.TOP) Then Call PrintHEAD
     PAGE.LINE = PAGE.LINE + 1: PAGE.TotLINES = PAGE.TotLINES + 1
     If PAGE.COUNT < PRT.BGNPAGE Then GoTo 2516
     XLINE$ = Space$(PAGE.WIDTH): PAGE.LOFST = Len(RTrim$(PLINE$))
     If PAGE.LOFST > 0 Then Mid$(XLINE$, PAGE.MARGIN + 1, PAGE.LOFST) = PLINE$
     PWIDTH = PAGE.LOFST + PAGE.MARGIN
     If Not (PAGE.TOP Or PAGE.FOOT) Then
        If (RptOPTION.LINEREF And 1) And ((PAGE.LINE Mod 10) = 0) Then
           NUMED$ = RTrim$(LTrim$(Str$(PAGE.LINE))): X = Len(NUMED$)
           If (((RptOPTION.LINEREF And 2) > 0) Or ((((RptOPTION.LINEREF And 8) > 0) And (PAGE.COUNT Mod 2) > 0) Or (((RptOPTION.LINEREF And 16) > 0) And (PAGE.COUNT Mod 2) = 0))) Then
              If PAGE.MARGIN > 3 Then
                 Z = (PAGE.MARGIN - 4): If Z = 0 Then Z = 1
                 Mid$(XLINE$, Z, X) = NUMED$
                 Mid$(XLINE$, Z + X, 1) = ">"
              End If
           End If
           If (((RptOPTION.LINEREF And 4) > 0) Or ((((RptOPTION.LINEREF And 16) > 0) And (PAGE.COUNT Mod 2) > 0) Or (((RptOPTION.LINEREF And 8) > 0) And (PAGE.COUNT Mod 2) = 0))) Then
              If (PAGE.RMARGIN > 3) Then
                 Z = (PAGE.WIDTH - PAGE.RMARGIN + 2)
                 Mid$(XLINE$, Z, 1) = "<"
                 Mid$(XLINE$, Z + 1, X) = NUMED$
                 PWIDTH = PAGE.WIDTH
              End If
           End If
        End If
     End If
     PWIDTH = PWIDTH - PRT.FOFFSET
     If PRT.FWIDTH < PWIDTH Then PWIDTH = PRT.FWIDTH
     If (((PAGE.COUNT Mod 2) + 1) And PRT.SNGLSIDE) > 0 Then
        On Local Error GoTo 2519
        If PWIDTH > 0 Then
           Select Case PRT.OPT
              Case 2: Printer.Print MID$(XLINE$, 1 + PRT.FOFFSET, PWIDTH);
              Case Else: Print #4, MID$(XLINE$, 1 + PRT.FOFFSET, PWIDTH);
           End Select
        End If
        Select Case PRT.OPT
           Case 2: Printer.Print ""
           Case Else: Print #4, CRLF$;
        End Select
     End If
2516 If INKEY$ <> "" Then Call HoldPRINT
2518 LSet PLINE$ = " ": PAGE.LOFST = 0: PAGE.TOP = 0
     If RPAGE = 0 Then PAGE.RPAGE = PAGE.COUNT: PAGE.RLINE = PAGE.LINE
     If ((Not PRT.CFORMS) And (Not PAGE.BREAK) And (PAGE.LINE > (PAGE.BLENGTH - 1))) Then
        Call PageBREAK
     End If
     'If PRT.Opt = 1 Then LOCATE CSRLIN, 1, 0
     Exit Sub

2519 'Print ERROR
     Call ERRMessage("XXXXM080", ERR, ERL)
     If a = 27 Then Resume 2518
     Resume

End Sub

