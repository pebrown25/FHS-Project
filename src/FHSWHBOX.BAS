Attribute VB_Name = "modFHSWHBox"
DefInt A-Z

   Dim ANCSEARCH, BRID, BNAME$, RTYPE$, WALIGN, ALIGNTAB$()
   Dim SCHAR$, BOXWIDTH, ORIENT, ORIENT2

Private Function FNDS$(a#)
   FNDS$ = Str$(Val(Str$(a#)))
End Function
Private Function FNODD(a#)
   FNODD = (Val(Right$(FNDS$(a#), 1)) Mod 2)
End Function
Private Function FNCW(a, B)
   FNCW = BOXWIDTH + (B = 0) * (((BOXWIDTH + 1) \ 2) - ((a + 1) Mod 2) * (BOXWIDTH Mod 2))
End Function
Private Function FNCOL(a)
   FNCOL = a + (2 * a - NodeHDR.WFMAXCOL - 1) * (ORIENT2 = 2)
End Function

Public Sub RptHBOX(OPT$)
   PN$ = " FAMILY HISTORY SYSTEM - FHSWHBOX - Horizontal BOX Chart Program"
   CY$ = " (C) Copyright 1992-1998 by Phillip E. Brown"
   USENEST = 0: GLMAX = 0: PSPOUSE = 0: CRULE = 0: ALIN# = 0
   READ$ = "READQ"

   'Call TableLOAD("WORKTYPE", T, 39)
   'If OKAY Then
   '   x = TABLE(T).FT
   '   While x > 0
   '     Z = Val(MID$(TABDATA$(x), 1, 2))
   '     If (8 < Z) And (Z < 12) Then
   '       HCTYPE$(Z - 8) = RTrim$(MID$(TABDATA$(x), 15, 255))
   '     End If
   '     x = TABDATACHN(x).FWD
   '   Wend
   '   Call TableFREE("WORKTYPE")
   'End If
   'Call TableLOAD("ALIGN", T, 0)
   'ReDim ALIGNTAB$(TABLE(T).SIZE)
   'x = TABLE(T).FT
   'While x > 0
   '   Y = Val(MID$(TABDATA$(x), 1, 1))
   '   ALIGNTAB$(Y) = MID$(TABDATA$(x), 3, 8)
   '   x = TABDATACHN(x).FWD
   'Wend
   'Call TableFREE("ALIGN")
   Select Case OPT$
      Case "INIT": GoTo 40
      Case "PRINT": GoTo 100
   End Select

40 ' Format Display
   Call PgmPREP("FHSHBOX")
   Call RptLOAD(RptCODE): Call RptLOAD("HSUM")
   Load frmFHSWRPgm
   frmFHSWRPgm.RPSetup.AddItem RptCODE
   frmFHSWRPgm.RPSetup.AddItem "HSUM"
   frmFHSWRPgm.RPSetup_Click
   frmFHSWRPgm.Show
   Exit Sub

100 ' Check Files
    ANCSEARCH = (NWrkType = 1): FamPath = (NWrkType = 3)
    Call NodeWrkOPEN("READQ")
    If Not OKAY Then GoTo 190
    WALIGN = Asc(NodeHDR.WFOPT) And 15
    If WALIGN = 0 Then GoTo 190
    If (FamPath And ((Asc(NodeHDR.WFOPT) And 32) = 0)) Then GoTo 190
    If (ANCSEARCH And ((Asc(NodeHDR.WFOPT) And 16) = 0)) Then GoTo 190
    If ((Not ANCSEARCH) And ((Asc(NodeHDR.WFOPT) And 16) > 0)) Then GoTo 190
    Close
    GoTo 1000

190 ' Bad Node Work File
    Call PutMSG("HBOXM002")
    Close
    Exit Sub
    
1000 ' Create Chart Symbol String for Screen Prints
     SCHAR$ = Space$(31): DCHAR$ = SCHAR$
     'For x = 1 To 31: READ Z: Mid$(SCHAR$, x, 1) = Chr$(Z): Next x
     'Data 218, 193, 191, 179, 180, 195, 196, 192, 194, 217, 218, 193, 191, 179, 180, 195, 196, 192, 194, 217, 197, 201, 207, 187, 186, 182, 199, 205, 200, 209, 188
     Call DataINIT("SBOXCHAR")
     Input #5, Z
     For X = 1 To 31: Input #5, Z: Mid$(SCHAR$, X, 1) = Chr$(Z): Next X
     Close #5
     
     ' Select Full or Summary Chart
     'Call PutMSG("HBOXM003")
     'If CA$ = Chr$(27) Then Exit Sub
     'A = Val(CA$): If A < 1 Or A > 2 Then A = 1
     If RptCODE = "HSUM" Then a = 1 Else a = 2
     If a < 2 Then
        Call HBOXSumry(NWrkType)
        Else
        Call HBOXChart(NWrkType)
     End If
End Sub
Rem $DYNAMIC
Sub BldChartDEF()
      RFCNT = 10
      ReDim RFLTH(RFCNT), RFLBL$(RFCNT)

59510 ' Create Report Definition Tables
      X = 1: RFNAME$ = "01030": GoSub 59505    'BIRTH DATE
      X = 2: RFNAME$ = "01032": GoSub 59505    'BIRTH PLACE
      X = 3: RFNAME$ = "02011": GoSub 59505    'SPOUSE
      X = 4: RFNAME$ = "02020": GoSub 59505    'MARRIAGE DATE
      X = 5: RFNAME$ = "02021": GoSub 59505    'MARRIAGE PLACE
      X = 6: RFNAME$ = "01040": GoSub 59505    'DEATH DATE
      X = 7: RFNAME$ = "01042": GoSub 59505    'DEATH PLACE
      X = 8: RFNAME$ = "02032": GoSub 59505    'END MARRIAGE REASON
      X = 9: RFNAME$ = "02030": GoSub 59505    'END MARRIAGE DATE
      X = 10: RFNAME$ = "02031": GoSub 59505   'END MARRIAGE PLACE
      GoTo 59560

59505 ' Process Report Field
      Call GetRFLD(RFNAME$, FX)
      If FOUND Then
         RDFFld = RFLDS(FX)
         If RFLTH(X) = 0 Then RFLTH(X) = Asc(RDFFld.RLTH)
         ALTH = Asc(RDFFld.ALTH)
         RFLBL$(X) = MID$(RFLD$(FX), 1, ALTH)
      End If
      Return

59560 ' Return to Calling Program
      OKAY = -1
End Sub

Sub HBOXChart(CTYPE)
800: 'Print Ancestor/Descendant/Family Path CHART
    Call PutMSG("HBOXM020")
    If CA$ = Chr$(27) Then GoTo 825
    MinChart = (CA$ <> "2"): NFMT$ = Space$(52)
    BD$ = Space$(12): DD$ = BD$: BPL$ = Space$(41): DPL$ = BPL$
    xAGE$ = "": xSTAT$ = " "
    SCLR = 0: PRINTALL = 0
    CENTERED = -1
    NUMED$ = Space$(5)
    Select Case CTYPE
       Case 1: RPT$ = "AHBX": ANCSEARCH = -1
               FCODE$ = RTrim$(MID$(RRELCD$(3), 3, 3))
               MCODE$ = RTrim$(MID$(RRELCD$(4), 3, 3))
       Case 2: RPT$ = "DHBX"
       Case 3: RPT$ = "FPTH": FamPath = -1
    End Select
    Call RptOPEN(RPT$): If Not OKAY Then GoTo 825
    Call BldChartDEF
    If ANCSEARCH Then
       Call RptOptGET("PCOD", X): ShowPCODE = (X = 1)
       Call RptOptGET("ALIN", X): ShowLIN = (X = 1)
       Else
       If Not FamPath Then
          Call RptOptGET("CHNM", X): ShowChNUM = (X = 1)
       End If
    End If
    Call RptOptGET("HEAD", X): If FOUND Then PRTHEAD = (X = 1)
    If MinChart Then
       If ShowLIN Then
          DW = Len(LTrim$(Str$(2 ^ (NodeHDR.WFHIGL + 1) - 1))) + 1
          Else
          DW = 3 - 2 * (Not ShowPCODE) * ANCSEARCH
       End If
       SameAS$ = ""
       Else
       Call GetRVAR("UNKNOWN", X): If FOUND Then Unknown$ = RVAR$(X)
       Call RptOptGET("DWTH", DW): If DW = 0 Then DW = 15
       Call GetRVAR("SAMEAS", X): If FOUND Then SameAS$ = RVAR$(X) + " "
       Call RptOptGET("PRID", X): ShowID = (X = 1)
       Call RptOptGET("SURF", X): LASTFIRST = (X = 1)
       Call RptOptGET("SURU", X): UPCASE = (X = 1)
       Call RptOptGET("SPLT", X): SPLITNM = (X = 1)
       Call RptOptGET("DAT2", X)
       If OKAY Then
          DateS = ((X = 1) Or (X = 4)): AllDATES = (X = 4)
          Else
          DateS = RptOPTION.DateS: AllDATES = 0
       End If
       PLACES = RptOPTION.PLACES
       SPOUSE = (RptOPTION.SPOUSE <> 0): ALLSPOUSE = (RptOPTION.SPOUSE = 2)
       Call RptOptGET("CMBD", X)
       If FOUND Then CombineDATA = (X <> 2): CombineALL = (X = 4)
    End If
    X = RptOPTION.LINEREF
    SHOWREF = ((X And 1) > 0)
    If SHOWREF Then
       LEFTSIDE = ((X And 2) > 0)
       RIGHTSIDE = ((X And 4) > 0)
       INSIDE = ((X And 8) > 0)
       OUTSIDE = ((X And 16) > 0)
       BOTHSIDES = ((X And 6) = 6)
    End If
    Call RptOptGET("SPBX", X): BOXSPAN = (X = 1)
    Call RptOptGET("BSEP", BOXSEP)
    Call RptOptGET("STEM", STLTH): ST1LTH = STLTH: ST2LTH = STLTH
    BDATA = BOXSEP + 2: BOXWIDTH = BDATA + DW
    FDL = -(13 * DateS + 44 * PLACES + DateS * PLACES)
    Call FAMOpen("READ"): If Not OKAY Then GoTo 825
    BRID = 0
    Call NodeWrkOPEN("READQ")
    If (Not OKAY) Or (NodeHDR.WFBRID = 0) Then GoTo 825
    Select Case CTYPE
       Case 1: If (Asc(NodeHDR.WFOPT) And 16) = 0 Then GoTo 825
       Case 2: If (Asc(NodeHDR.WFOPT) And 48) > 0 Then GoTo 825
       Case 3: If (Asc(NodeHDR.WFOPT) And 32) = 0 Then GoTo 825
    End Select
    BRID = NodeHDR.WFBRID: Call FF1GetRec(BRID): BNAME$ = FMTNAME$(1)
    SysVAR.SUBJECT = BNAME$: Call HTFExpand(ENV.RTITLE, 1)
    WALIGN = Asc(NodeHDR.WFOPT) And 15
    Call ShowINFO
    If (WALIGN < 1) Or (WALIGN = 5) Then
       Call PutMSG("HBOXM021"): GoTo 825
    End If
    If (WALIGN = 3) Or (WALIGN = 7) Then BOXSPAN = -1
    Call PutMSG("HBOXM022")
    If CA$ = Chr$(27) Then GoTo 825
    ORIENT = (-(CA$ <> "2") - 2 * (CA$ = "2"))
    If NodeHDR.WFMAXCOL = 1 Then ORIENT2 = 1: GoTo 830
    If ANCSEARCH Then
       Call PutMSG("HBOXM023")
       Else
       Call PutMSG("HBOXM024")
    End If
    If CA$ = Chr$(27) Then GoTo 825
    ORIENT2 = (-(CA$ <> "2") - 2 * (CA$ = "2"))
    GoTo 830

825 ' Early EXIT
    Call FamCLOSE: Exit Sub

830 ' Prepare for Printing (Single/Multi Strip)
    Call PrintOPEN: If Not OKAY Then GoTo 825
    BGNPAGE = PAGE.BGNPAGE
    BLDREF = RptOPTION.BLDREF
    If RptOPTION.BLDREF Then
       RREFRPT.BGNPAGE = BGNPAGE
       RREFREC.AGL = Chr$(0): RREFREC.DGL = Chr$(0)
    End If
    PAGE.BWIDTH = PRT.FWIDTH: LW = PRT.FWIDTH
    FIRSTONE = -1: CST = 1: ts = 1: BDLNUM = 1: ALLREF = 0
    LS = LEFTSIDE Or INSIDE: RS = RIGHTSIDE Or OUTSIDE
    LSM1 = PAGE.IMARGIN: RSM1 = PAGE.OMARGIN
    PAGE.IMARGIN = 0: PAGE.OMARGIN = 0
    BW# = BOXWIDTH * ((NodeHDR.WFMAXCOL + 1) \ 2) + FNCW(NodeHDR.WFMAXCOL + 1, 0) * ((NodeHDR.WFMAXCOL + 1) Mod 2)
    TW# = BW# + LSM1 - 8 * (LS + RS)
    PRTSTRIP = (TW# > LW)
    If Not PRTSTRIP Then
       REFAREA = (LW - BW# - LSM1 - RSM1) \ 2
       If SHOWREF Then
          If REFAREA < 5 Then REFAREA = 5
          If Not LS Then LSM1 = LSM1 + REFAREA
          RM# = LSM1 - REFAREA * LS + BW# - (REFAREA - 5) * RS
          Else
          If REFAREA < 0 Then REFAREA = 0
          LSM1 = LSM1 + REFAREA
       End If
       MARGIN = LSM1: PO# = 0: GoTo 900
    End If
    If PRT.OPT < 3 Then GoTo 845

840 ' Check for Full Width printing to File
    Call PutMSG("XXXXM070")
    If CA$ = Chr$(27) Then GoTo 825
    PRTSTRIP = (Val(CA$) = 1): PAGE.FULLWIDTH = Not PRTSTRIP
    If PAGE.FULLWIDTH Then
       If SHOWREF Then
          REFAREA = 8
          RM# = LSM1 - REFAREA * LS + BW# - (REFAREA - 5) * RS
       End If
       PRT.CFORMS = -1: MARGIN = LSM1: PO# = 0: GoTo 900
    End If

845 ' Prepare for Strip Printing
    If BOXSPAN Then GoTo 850
847 ' Adjust for Boxes not Crossing Border of Strip
    REFAREA = -8 * SHOWREF: ALLREF = -1
    BIGMGN = LSM1 + (LSM1 - RSM1) * (RSM1 > LSM1)
    c1 = (LW - BIGMGN + REFAREA * (LS + RS)) \ BOXWIDTH
    If c1 < 1 Then
       If (LSM1 + RSM1) > 0 Then
          LSM1 = 0: RSM1 = 0: GoTo 847
          Else
          If SHOWREF Then
             SHOWREF = 0: LS = 0: RS = 0: GoTo 847
             Else
             BOXSPAN = -1: GoTo 850
          End If
       End If
    End If
    ts = 1 + (((NodeHDR.WFMAXCOL + 1) \ 2) - 1) \ c1
    If PRT.SNGLSIDE Then
       LSM2 = LSM1: RSM2 = RSM1
       Else
       LSM2 = RSM1: RSM2 = LSM1
    End If
    GoTo 855

850 'Compute Number of Strips & Margins if Spanned Boxes
    ALLREF = 0
    TW# = BW# + LSM1 - 5 * (LS + RS)
    ts = Fix((TW# + LW - 1) / LW)
    REFAREA = (ts * LW - BW# - LSM1 - RSM1) \ 2
    If REFAREA < 5 Then REFAREA = 5
    If Not LS Then LSM1 = LSM1 + REFAREA: TW# = TW# + REFAREA

855 ' Prepare for First Strip
    CST = 0: GoSub 980: If Not OKAY Then GoTo 825

900 ' Prepare To Print Chart
    If Not SHOWREF Then LEFTSIDE = 0: RIGHTSIDE = 0: INSIDE = 0: OUTSIDE = 0
    PDATA$ = Space$(100): DDATA$ = Space$(100)
    If ORIENT = 1 Then
       VMIN = 0: VMAX = NodeHDR.WFHIGL: VSTEP = 1
       Else
       VMIN = NodeHDR.WFHIGL: VMAX = 0: VSTEP = -1
    End If
    HMIN = 1: HMAX = NodeHDR.WFMAXCOL: HSTEP = 1
    X = HMAX + (HMAX - HMIN) * (HMIN > HMAX)
    ReDim GLLOCOL(NodeHDR.WFHIGL), GLHICOL(NodeHDR.WFHIGL)
    ReDim xPCOL(X), CHCNT(X)
    ReDim DIL(X + 1), DID(X), xPID(X), DCHNUM(X)
    ReDim DMRN(X), DMID(X), DMIL(X), NDMRN(X)
    If RptOPTION.BLDREF Then ReDim ColRefID(X)
    If ANCSEARCH Then ReDim XALIN#(X)
    ' Create BOX Symbol Set Variables
    'If PRT.OPT = 1 Then
    '   DCHAR$ = SCHAR$
    '   Else
       DCHAR$ = PDFSetup.SET2 + Left$(PDFSetup.SET1, 11) + PDFSetup.SET3
    'End If
    HBC$ = MID$(DCHAR$, 7)
    BOXTOP$ = String$(DW + 2, Asc(HBC$)): BOXBOT$ = BOXTOP$
    Mid$(BOXTOP$, 1, 1) = MID$(DCHAR$, 1, 1)
    Mid$(BOXTOP$, DW + 2, 1) = MID$(DCHAR$, 3, 1)
    Mid$(BOXBOT$, 1, 1) = MID$(DCHAR$, 8, 1)
    Mid$(BOXBOT$, DW + 2, 1) = MID$(DCHAR$, 10, 1)
    PS = BDATA + DW \ 2
    If ORIENT = 1 Then
       PC = 12: CS = PS: CC = 19: FCC = 11: LCC = 13
       Else
       PC = 19: CS = PS: CC = 12: FCC = 18: LCC = 10
    End If
    xPC$ = MID$(DCHAR$, PC, 1): xCC$ = MID$(DCHAR$, CC, 1)
    xFCC$ = MID$(DCHAR$, FCC, 1): xLCC$ = MID$(DCHAR$, LCC, 1)
    VBC$ = MID$(DCHAR$, 4, 1): VLC$ = MID$(DCHAR$, 14, 1)
    xC$ = MID$(DCHAR$, 21, 1): HLC$ = MID$(DCHAR$, 17, 1)
    HBC$ = MID$(DCHAR$, 28)
    PBOXTOP$ = String$(DW + 2, Asc(HBC$)): PBOXBOT$ = PBOXTOP$
    Mid$(PBOXTOP$, 1, 1) = MID$(DCHAR$, 22, 1)
    Mid$(PBOXTOP$, DW + 2, 1) = MID$(DCHAR$, 24, 1)
    Mid$(PBOXBOT$, 1, 1) = MID$(DCHAR$, 29, 1)
    Mid$(PBOXBOT$, DW + 2, 1) = MID$(DCHAR$, 31, 1)
    PVBC$ = MID$(DCHAR$, 25, 1)
    FATHER$ = RTrim$(MID$(RRELCD$(3), 3, 3))
    MOTHER$ = RTrim$(MID$(RRELCD$(4), 3, 3))
    For GL = 0 To NodeHDR.WFHIGL
        Call NodeWorkGET("GLREC", GL)
        GLLOCOL(GL) = NodeGLREC.WFLOCOL
        GLHICOL(GL) = NodeGLREC.WFHICOL
    Next GL

930 ' Prepare to Begin Printing
    PAGE.LOFST = 0: PAGE.COUNT = PAGE.BGNPAGE - 1: PAGE.LINE = 0
    PRT.EOF = 0: DLNUM = 0
    For X = HMIN To HMAX Step HSTEP
        DCHNUM(X) = 0: DIL(X) = 0: DMIL(X) = 0
    Next X
    ReDim REFLN(FF1Hdr.MAXID), REFCOL(FF1Hdr.MAXID)
    If Not PRT.CFORMS Then
       ReDim REFPG(FF1Hdr.MAXID)
    End If
    If PRTSTRIP Then
       SysVAR.NVAR1 = CST: SysVAR.NVAR2 = ts
       Call PutMSG("HBOXM027")
       Else
       SysVAR.MVAR = RTYPE$
       Call PutMSG("HBOXM029")
    End If
    IL = 1: GoSub 2200: If Not OKAY Then GoTo 990
    VX = VMIN: HXMAX = 0: FIRSTONE = 0


939 ' Prepare WREC, WCOL, PCOL, CHCNT
    Call NodeWorkGET("GLREC", VMIN)
    GLBGN = NodeGLREC.WFGLBGN: GLEND = NodeGLREC.WFGLEND
    If ORIENT2 = 1 Then
       HXMAX = FNCOL(NodeGLREC.WFHICOL)
       Else
       HXMAX = FNCOL(NodeGLREC.WFLOCOL)
    End If
    For X = GLBGN To GLEND
        Call NodeWorkGET("NODEREC", X)
        Col = FNCOL(NodeREC.WFCOL)
        xPCOL(Col) = 0: xPID(Col) = 0
        CHCNT(Col) = NodeREC.WFNUMCH: Z = NodeREC.WFCHNUM
        If (ORIENT2 = 2) And (Z > 255) And ((Z And 768) Xor (768)) Then Z = Z Xor 768
        DCHNUM(Col) = Z: RID = NodeREC.WFRID
        DID(Col) = RID: DIL(Col) = 1: PREC = NodeREC.WFPREC
        If ANCSEARCH Then XALIN#(Col) = ALIN#
        If PREC > 0 Then
           Call NodeWorkGET("NODEREC", PREC)
           xPID(Col) = NodeREC.WFRID: xPCOL(Col) = FNCOL(NodeREC.WFCOL)
        End If
        If SPOUSE Then
           Call FF1GetRec(RID): DMIL(Col) = -1
           NDMRN(Col) = FF1Rec.SPOUSE
           Else
           DMIL(Col) = 0
        End If
    Next X
    HXCNT = GLEND - GLBGN + 1

950 ' Print DATALINEs
    XCNT = HXCNT
955 HX = HMIN: GL = VX: PAGE.LOFST = 0
    While HX > -1: GoSub 1000: Wend
    GoSub 1165: If a = 27 Then GoTo 990
    If XCNT > 0 Then GoTo 955 'Are there unfinished Boxes?

960 ' Process Space Between Rows
    If ORIENT = 1 Then STEM = 1: XMAX = ST1LTH Else STEM = 2: XMAX = ST2LTH
    For xXC = 1 To XMAX
        HX = HMIN
        While HX > -1
           IL = DIL(HX): GoSub 1080: GoSub 1140
           If a = 27 Then
              HX = -1
              Else
              HX = HX + (1 - (IL <> 0)) * HSTEP
              If ((HSTEP = 1) And (HX > HXMAX)) Or ((HSTEP = -1) And (HX < HXMAX)) Then HX = -1
           End If
        Wend
        If a = 27 Then X = XMAX: Else GoSub 1165: If a = 27 Then X = XMAX
    Next xXC
    STEM = 0

970 ' Prepare for Next Row
    If a = 27 Then GoTo 990
    If VX = VMAX Then GoTo 990 Else VX = VX + VSTEP
    GoSub 1100
    If a = 27 Then GoTo 990 Else GoTo 950

980 ' Check for Strip to Print
    CST = CST + 1
    If CST > ts Then GoTo 989
    If Not (PRINTALL) Then
       If Not (EVENONLY Or ODDONLY) Then
          SysVAR.NVAR1 = CST: SysVAR.NVAR2 = ts
          Call PutMSG("HBOXM026")
          If CA$ = Chr$(27) Then GoTo 989
          If (CA$ = SNGLKEY$(2)) Then GoTo 980
          PRINTALL = (CA$ = SNGLKEY$(4))
          EVENONLY = (CA$ = SNGLKEY$(16)): ODDONLY = (CA$ = SNGLKEY$(17))
       End If
       If (EVENONLY And ((CST Mod 2) = 1)) Or (ODDONLY And ((CST Mod 2) = 0)) Then GoTo 980
    End If
    PRT.AGAIN = -1
    If BOXSPAN Then
       If CST = 1 Then
          MARGIN = LSM1: LS = LEFTSIDE Or INSIDE: RS = 0
          Else
          MARGIN = 0: LS = 0
          If CST = ts Then RS = RIGHTSIDE Or OUTSIDE Else RS = 0
       End If
       PO# = (CST - 1) * LW + (CST > 1) * (LSM1 - REFAREA * (LEFTSIDE Or INSIDE))
       If RS Then RM# = TW# - 5 - PO#
       Else
       If (CST Mod 2) = 1 Then
          MARGIN = LSM1
          LS = LEFTSIDE Or INSIDE: RS = RIGHTSIDE Or OUTSIDE
          Else
          MARGIN = LSM2
          LS = LEFTSIDE Or OUTSIDE: RS = RIGHTSIDE Or INSIDE
       End If
       PO# = (CST - 1) * c1 * BOXWIDTH
       RM# = MARGIN - REFAREA * LS + c1 * BOXWIDTH - (REFAREA - 5) * RS
    End If
989 Call PutMSG("")
    Return

990 ' End of Printing CHART
    If RptOPTION.BLDREF Then
       Call RREFClose: RptOPTION.BLDREF = 0
    End If
    Call PageBREAK: PRT.AGAIN = 0
    If PRTSTRIP Then
       GoSub 980
       If PRT.AGAIN Then GoTo 930
    End If
    If (Not MinChart) And (MXWIDTH > DW) Then
       SysVAR.NVAR1 = MXWIDTH: SysVAR.NVAR2 = DW
       Call PutMSG("HBOXM030")
    End If
    Call RptCLOSE
    Call FamCLOSE
    On Local Error GoTo 0
    Exit Sub

1000 ' Process GL Data
     IL = DIL(HX): RID = DID(HX):  xo = 1: LSet PDATA$ = " "
     If FamPath And (Not NOSIBLING) Then
        MBRPATH = (RID = NodeHDR.WFID1) Or (RID = NodeHDR.WFID2) Or (CHCNT(HX) > 0)
        Else
        MBRPATH = 0
     End If
     If (IL = -1) Then GoTo 1010
     If IL = 0 Then GoTo 1010  ' No more Data
     If IL = 1 Then GoSub 1090 ' Print BOX Top
     If IL = 2 Then GoSub 1200 ' Print NAME Line(1)
     If IL = 3 Then GoSub 1200 ' Print NAME Line(2)
     If IL = 4 Then GoSub 1220 ' Print BIRTH Date (After Check for Dup ID)
     If IL = 5 Then GoSub 1240 ' Print BIRTH Place
     If IL = 6 Then GoSub 1250 ' Print DEATH Date
     If IL = 7 Then GoSub 1260 ' Print DEATH Place
     If IL = 8 Then GoSub 1300 ' Print MARRIAGE Data
     If IL = 9 Then GoSub 1095 ' Print BOX Bottom
     If (IL > 1) And (IL < 9) Then GoSub 1125
1010 If IL > 10 Then LSet DDATA$ = " ": GoSub 1080
     GoSub 1140
     If DUPID Then
        IL = 8: DUPID = 0
        Else
        If ((IL > 0) And (IL < 8)) Or (IL = 10) Then IL = IL + 1
     End If
     If MinChart And (IL > 4) And (IL < 9) Then IL = 9
     DIL(HX) = IL
     HX = HX + (1 - (IL <> 0)) * HSTEP
     If ((HSTEP = 1) And (HX > HXMAX)) Or ((HSTEP = -1) And (HX < HXMAX)) Then HX = -1
     Return

1080 ' Insert DSEP$ and BOX Characters
     If IL = 0 Then
        Mid$(DDATA$, PS, 1) = " "
        Else
        If (ORIENT = 1) Then
           If (((STEM = 1) And (CHCNT(HX) > 0)) Or ((STEM = 2) And (xPCOL(HX) > 0))) Then
              Mid$(DDATA$, PS, 1) = VLC$
           End If
           Else
           If (((STEM = 2) And (xPCOL(HX) > 0)) Or ((STEM = 1) And (CHCNT(HX) > 0))) Then
              Mid$(DDATA$, PS, 1) = VLC$
           End If
        End If
     End If
     Return

1090 ' Print BOX Top
     If Not MBRPATH Then
        Mid$(DDATA$, BDATA - 1, DW + 2) = BOXTOP$
        Else
        Mid$(DDATA$, BDATA - 1, DW + 2) = PBOXTOP$
     End If
     If (xPCOL(HX) > 0) And (ORIENT = 1) Then
        If Not MBRPATH Then
           Mid$(DDATA$, PS, 1) = MID$(DCHAR$, 2, 1)
           Else
           Mid$(DDATA$, PS, 1) = MID$(DCHAR$, 23, 1)
        End If
     End If
     If (CHCNT(HX) > 0) And (ORIENT = 2) Then
        If Not MBRPATH Then
           Mid$(DDATA$, PS, 1) = MID$(DCHAR$, 2, 1)
           Else
           Mid$(DDATA$, PS, 1) = MID$(DCHAR$, 23, 1)
        End If
     End If
     Return

1095 ' Print BOX Bottom
     IL = 10
     If Not MBRPATH Then
        Mid$(DDATA$, BDATA - 1, DW + 2) = BOXBOT$
        Else
        Mid$(DDATA$, BDATA - 1, DW + 2) = PBOXBOT$
     End If
     If (CHCNT(HX) > 0) And (ORIENT = 1) Then
        STEM = 1
        If Not MBRPATH Then
           Mid$(DDATA$, PS, 1) = MID$(DCHAR$, 9, 1)
           Else
           Mid$(DDATA$, PS, 1) = MID$(DCHAR$, 30, 1)
        End If
     End If
     If (xPCOL(HX) > 0) And (ORIENT = 2) Then
        STEM = 2
        If Not MBRPATH Then
           Mid$(DDATA$, PS, 1) = MID$(DCHAR$, 9, 1)
           Else
           Mid$(DDATA$, PS, 1) = MID$(DCHAR$, 30, 1)
        End If
     End If
     XCNT = XCNT - 1
     Return

1100 ' Increment VX for Horizontal Arrangement
     If ORIENT = 2 Then GoSub 1400: GoSub 1165: If a = 27 Then GoTo 1109
     For HX = HMIN To HXMAX Step HSTEP
         DIL(HX) = 0: xPCOL(HX) = 0: CHCNT(HX) = 0: DCHNUM(HX) = 0
     Next HX
     Call NodeWorkGET("GLREC", VX)
     GLBGN = NodeGLREC.WFGLBGN: GLEND = NodeGLREC.WFGLEND
     If ORIENT2 = 1 Then
        HXMAX = FNCOL(NodeGLREC.WFHICOL)
        Else
        HXMAX = FNCOL(NodeGLREC.WFLOCOL)
     End If
     For X = GLBGN To GLEND
         Call NodeWorkGET("NODEREC", X)
         XCOL = NodeREC.WFCOL: Col = FNCOL(XCOL)
         xPCOL(Col) = 0
         CHCNT(Col) = NodeREC.WFNUMCH: Z = NodeREC.WFCHNUM
         If (ORIENT2 = 2) And (Z > 255) And ((Z And 768) Xor (768)) Then Z = Z Xor 768
         DCHNUM(Col) = Z: RID = NodeREC.WFRID: DID(Col) = RID
         DIL(Col) = 1: If ANCSEARCH Then XALIN#(Col) = ALIN#
         If SPOUSE Then
            Call FF1GetRec(RID)
            DMIL(Col) = -1: NDMRN(Col) = FF1Rec.SPOUSE
            Else
            DMIL(Col) = 0
         End If
         PREC = NodeREC.WFPREC
         If PREC > 0 Then
            Call NodeWorkGET("NODEREC", PREC)
            xPID(Col) = NodeREC.WFRID
            xPCOL(Col) = FNCOL(NodeREC.WFCOL)
            Else
            xPID(Col) = 0: xPCOL(Col) = 0
         End If
     Next X
     HXCNT = GLEND - GLBGN + 1
     If ORIENT = 1 Then GoSub 1400: GoSub 1165: If a = 27 Then GoTo 1109
     If ORIENT = 1 Then STEM = 2: XMAX = ST2LTH Else STEM = 1: XMAX = ST1LTH
     For xXC = 1 To XMAX
         HX = HMIN
         While HX > -1
            IL = DIL(HX): GoSub 1080: GoSub 1140
            If a = 27 Then
               HX = -1
               Else
               HX = HX + (1 - (IL <> 0)) * HSTEP
               If ((HSTEP = 1) And (HX > HXMAX)) Or ((HSTEP = -1) And (HX < HXMAX)) Then HX = -1
            End If
         Wend
         If a = 27 Then xXC = XMAX: Else GoSub 1165: If a = 27 Then xXC = XMAX
     Next xXC: STEM = 0
1109 LSet PDATA$ = " "
     Return

1125 ' Place Characters for Side of Box, and Place Data in Box
     If Not MBRPATH Then
        Mid$(DDATA$, BDATA - 1, 1) = VBC$
        Mid$(DDATA$, BDATA + DW, 1) = VBC$
        Else
        Mid$(DDATA$, BDATA - 1, 1) = PVBC$
        Mid$(DDATA$, BDATA + DW, 1) = PVBC$
     End If
     xs = BDATA
     If Not CENTERED Then
        XL = DW
        Else
        If FL > 0 Then
           XL = FL: FL = 0
           Else
           'XL = InStr(PDATA$, "   ") - 1
           XL = Len(LTrim$(RTrim$(PDATA$)))
        End If
        If (XL < 1) Or (XL > DW) Then
           XL = DW
           Else
           xs = xs + (DW - XL + 1) \ 2
        End If
     End If
     Mid$(DDATA$, xs, XL) = LTrim$(PDATA$)
     If (Not MinChart) Then
        XDW = Len(RTrim$(PDATA$))
        If (XDW > MXWIDTH) Then MXWIDTH = XDW
     End If
     Return

1140 ' Print Column Data in DDATA$
     If (a = 27) Or ((DLNUM + 1) < BDLNUM) Then GoTo 1149
     CWIDTH = FNCW(HX, IL)
     If Not PRTSTRIP Then
        x1 = 1: XL = CWIDTH
        Else
        x1 = 0
        X0# = ((HX - 1) \ 2) * (BOXWIDTH) + FNCW(HX - 1, 0) * ((HX - 1) Mod 2)
        If BOXSPAN Then
           XL = LW + (CST = 1) * (LSM1 - REFAREA * LS)
           Else
           XL = c1 * BOXWIDTH
        End If
        If (X0# > (PO# + XL)) Or ((X0# + CWIDTH - 1) < PO#) Then
           If (HX = 1) Or ((Not BOXSPAN) And SHOWREF) Then GoTo 1142
           If HX = HXMAX Then GoTo 1144
           GoTo 1149
        End If
        If X0# <= PO# Then x1 = PO# - X0# + 1 Else x1 = 1
        If (PO# + XL) >= (X0# + CWIDTH) Then
           XL = CWIDTH - x1 + 1
           Else
           XL = ((PO# + XL) - (X0# + x1 - 1))
        End If
     End If
     If XL = 0 Then GoTo 1149
1142 If (PAGE.LOFST = 0) And (MARGIN > 0) Then
        PAGE.LOFST = MARGIN
        If PAGE.FULLWIDTH Then
           Print #4, Left$(PLINE$, PAGE.LOFST);
           PO# = PAGE.LOFST
        End If
     End If
     If LS And (PAGE.LOFST = MARGIN) Then
        If (((DLNUM + 1) Mod 5) = 0) Then
           RSet NUMED$ = fnx$(DLNUM + 1)
           Mid$(PLINE$, PAGE.LOFST + 1, 5) = Right$(NUMED$, 4) + ">"
        End If
        PAGE.LOFST = PAGE.LOFST + REFAREA
        If PAGE.FULLWIDTH Then
           Print #4, MID$(PLINE$, MARGIN + 1, REFAREA);
           PO# = PO# + REFAREA
        End If
     End If
     If x1 > 0 Then
        If PAGE.FULLWIDTH Then
           Print #4, Left$(DDATA$, CWIDTH);
           PO# = PO# + CWIDTH
           Else
           Mid$(PLINE$, PAGE.LOFST + 1, XL) = MID$(DDATA$, x1, XL)
           PAGE.LOFST = PAGE.LOFST + XL
        End If
     End If
1144 If RS And (((DLNUM + 1) Mod 5) = 0) And (HX = HXMAX) Then
        If PAGE.FULLWIDTH Then
           X = RM# - PO#
           If X > 0 Then
              Print #4, Space$(X) + "<" + fnx$(DLNUM + 1);
           End If
           Else
           PAGE.LOFST = RM#
           Mid$(PLINE$, PAGE.LOFST + 1, 255) = "<" + fnx$(DLNUM + 1)
           PAGE.LOFST = PAGE.LOFST + 5
        End If
     End If
1149 LSet DDATA$ = " "
     Return

1165 ' Advance One Line Forward
     If (a <> 27) Then
        If (PAGE.LINE = 0) And (Not PAGETOP) Then GoSub 2200
        DLNUM = DLNUM + 1
        If PAGE.FULLWIDTH Then
           Print #4, ""
           PAGE.LINE = PAGE.LINE + 1:  LSet PLINE$ = ""
           If INKEY$ <> "" Then Call HoldPRINT
           PAGE.RLINE = DLNUM
           Else
           Call WriteLINE: BLEFT = BLEFT - 1
        End If
        PAGE.LOFST = 0
     End If
     If RptOPTION.BLDREF Then GoSub 1170
     Return

1170 ' Create RREF Entries
     If CTYPE = 1 Then AGL = GL: DGL = 0: Else AGL = 0: DGL = GL
     RPAGE = PAGE.RPAGE: RLINE = PAGE.RLINE
     For HX = HMIN To HXMAX Step HSTEP
         RID = ColRefID(HX)
         If (RID <> 0) Then
            PAGE.XID = Abs(RID)
            If RID > 0 Then
               PAGE.xType = 1
               RREFREC.AGL = Chr$(AGL): RREFREC.DGL = Chr$(DGL)
               Else
               PAGE.xType = 3
            End If
            RREFXData$ = fnx$(RPAGE) + "-" + fnx$(RLINE) + ":" + fnx$(REFCOL(PAGE.XID))
            PAGE.RPAGE = RPAGE: PAGE.RLINE = 0
            Call RREFAdd: Call RREFBld
            REFLN(PAGE.XID) = RLINE: If Not PRT.CFORMS Then REFPG(PAGE.XID) = RPAGE
         End If
         ColRefID(HX) = 0
     Next HX
     PAGE.RPAGE = 0
     Return

1200 ' Print Relative NAME Line
     LSet xSTAT$ = " "
     If RID > 0 Then
        Call FF1GetRec(RID)
        Else
        LSet FF1Rec.GIVEN = "(" + Unknown$ + ")": LSet FF1Rec.SURNM = ""
     End If
     If Not ((IL <> 2) Or (CSID > 0) Or (GL = 0)) Then
        CN = DCHNUM(HX) And 255
        If Not ANCSEARCH Then
           If ShowChNUM Then Mid$(PDATA$, xo, 3) = fnx$(CN)
           If xPID(HX) = FF1Rec.FID Then LSet xSTAT$ = FDSTAT$("FID") Else LSet xSTAT$ = FDSTAT$("MID")
           Else
           If ShowLIN Then
              Mid$(PDATA$, xo, 255) = LTrim$(Str$(XALIN#(HX)))
              Else
              If ShowPCODE Then
                 If MinChart Then xo = xo + 1
                 If CN = 1 Then
                    Mid$(PDATA$, xo, 1) = FCODE$
                    Else
                    Mid$(PDATA$, xo, 1) = MCODE$
                 End If
              End If
           End If
           If (DCHNUM(HX) And 1024) > 0 Then LSet xSTAT$ = "*"
           If (DCHNUM(HX) And 2048) > 0 Then LSet xSTAT$ = "?"
        End If
        If xSTAT$ <> " " Then
           xo = InStr(xo, PDATA$, " ")
           Mid$(PDATA$, xo, 1) = xSTAT$
        End If
        xo = InStr(xo, PDATA$, " "): If xo = 0 Then GoTo 1209
        If Not MinChart Then
           If ShowChNUM Or ShowPCODE Then
              Mid$(PDATA$, xo, 1) = ".": xo = xo + 1
           End If
        End If
     End If
     If MinChart Then
        If FamPath Then
           If FF1Rec.SEX = FGENDR$(1) Then
              X = 1
              Else
              If FF1Rec.SEX = FGENDR$(2) Then
                 X = 2
                 Else
                 X = 3
              End If
           End If
           Mid$(PDATA$, xo, 1) = MID$(RGENDR$(X), 2, 1)
        End If
        IL = IL - (IL = 2)
        GoTo 1209
     End If
     If (IL = 2) And ShowID Then
        Mid$(PDATA$, xo, 1) = "("
        Mid$(PDATA$, xo + 1, 5) = fnx$(RID)
        xo = InStr(xo, PDATA$, " ")
        If (Not ANCSEARCH) And (CSID = 0) And (xPID(HX) <> 0) And SPOUSE Then
           LSet xSTAT$ = " "
           If xPID(HX) = FF1Rec.FID Then
              PID = FF1Rec.MID: If PID <> 0 Then LSet xSTAT$ = FDSTAT$("MID")
              Else
              PID = FF1Rec.FID: If PID <> 0 Then LSet xSTAT$ = FDSTAT$("FID")
           End If
           If PID > 0 Then
              Mid$(PDATA$, xo, 1) = ":"
              Mid$(PDATA$, xo + 1, 5) = fnx$(PID)
              If xSTAT$ <> " " Then
                 xo = InStr(xo, PDATA$, " ")
                 Mid$(PDATA$, xo, 1) = xSTAT$
              End If
              xo = InStr(xo, PDATA$, " ")
           End If
        End If
        Mid$(PDATA$, xo, 1) = ")": xo = xo + 1
     End If
     If Not SPLITNM Then
        If (IL = 2) Then
           LSet NFMT$ = FMTNAME$(0)
           Else
           IL = IL + 1: GoTo 1209
        End If
        Else
        If (IL = 2) Then If LASTFIRST Then LSet NFMT$ = FF1Rec.SURNM Else LSet NFMT$ = FF1Rec.GIVEN
        If (IL = 3) Then If LASTFIRST Then LSet NFMT$ = FF1Rec.GIVEN Else LSet NFMT$ = FF1Rec.SURNM
        If UPCASE Then
           If (((IL = 2) And LASTFIRST) Or ((IL = 3) And (Not LASTFIRST))) Then
              LSet NFMT$ = UCX$(NFMT$)
           End If
        End If
     End If
     If Len(RTrim$(NFMT$)) > 0 Then
        Mid$(PDATA$, xo, 255) = NFMT$: xo = DW
     End If
1209 Return

1220 ' Check for Crossed Lines & Create REF Record
     If (RID = 0) Then GoTo 1230
     DUPID = (REFLN(RID) <> 0)
     If DUPID Then
        DUPCOL = REFCOL(RID): DUPLN = REFLN(RID)
        If MinChart And (DW < 7) Then
           Mid$(PDATA$, xo, 3) = "<->"
           Else
           xDATA$ = fnx$(DUPLN) + ":" + fnx$(DUPCOL)
           If Not PRT.CFORMS Then
              xDATA$ = fnx$(REFPG(RID)) + "-" + xDATA$
           End If
           Mid$(PDATA$, xo, 255) = "<" + SameAS$ + xDATA$ + ">"
           xo = DW
        End If
     End If
     REFCOL(RID) = ((HX + 1) \ 2)
     If PRT.CFORMS Then
        REFLN(RID) = (DLNUM + SPLITNM)
        Else
        REFPG(RID) = PAGE.COUNT - (BLEFT = 0)
        REFLN(RID) = PAGE.LINE + (PAGE.LINE - PAGE.HLCNT - 1) * (BLEFT = 0)
     End If
     If RptOPTION.BLDREF And (RID <> 0) Then
        If CSID > 0 Then
           ColRefID(HX) = -CSID
           Else
           ColRefID(HX) = RID
        End If
     End If
     If Not DUPID Then GoTo 1230
     Return

1230 ' Print BIRTH Date
     If MinChart Then IL = 9: Return
     If (DateS) And (RID <> 0) Then
        Call FF1GetRec(RID)
        Call GetBDEvents(BD$, BPL$, DD$, DPL$, xAGE$)
        LBD = Len(RTrim$(LTrim$(BD$)))
        If LBD > 0 Then
           Mid$(PDATA$, xo, 14) = RFLBL$(1)
           x1 = InStr(xo, PDATA$, "   ")
           If x1 > xo Then
              xo = x1 - (x1 > xo) + (MID$(PDATA$, x1, 1) = ".")
           End If
           Mid$(PDATA$, xo, 12) = LTrim$(BD$)
           xo = InStr(xo, PDATA$, "   ")
        End If
        GoSub 1270
     End If
     If xo = 1 Then IL = IL + 1
     Return

1240 ' Print BIRTH Place
     If (PLACES) And (RID <> 0) Then
        Call FF1GetRec(RID)
        Call GetBDEvents(BD$, BPL$, DD$, DPL$, xAGE$)
        LBP = Len(LTrim$(RTrim$(BPL$)))
        If LBP > 0 Then
           Mid$(PDATA$, xo, 3) = RFLBL$(2)
           x1 = InStr(xo, PDATA$, "   ")
           If x1 > xo Then
              xo = x1 - (x1 > xo) + (MID$(PDATA$, x1, 1) = ".")
           End If
           Mid$(PDATA$, xo, 255) = LTrim$(BPL$)
           xo = InStr(xo, PDATA$, "   ")
        End If
        If (CSID = 0) And (CombineALL) Then GoSub 1270
     End If
     If xo = 1 Then IL = IL + 1
     Return

1250 ' Print DEATH Date
     If (DateS) And (RID <> 0) Then
        Call FF1GetRec(RID)
        Call GetBDEvents(BD$, BPL$, DD$, DPL$, xAGE$)
        LDD = Len(LTrim$(RTrim$(DD$)))
        If LDD > 0 Then
           Mid$(PDATA$, xo, 14) = RFLBL$(6)
           x1 = InStr(xo, PDATA$, "   ")
           If x1 > xo Then
              xo = x1 - (x1 > xo) + (MID$(PDATA$, x1, 1) = ".")
           End If
           Mid$(PDATA$, xo, 12) = LTrim$(DD$)
           xo = InStr(xo, PDATA$, "   ")
        End If
        GoSub 1270
     End If
     If xo = 1 Then IL = IL + 1
     Return

1260 ' Print DEATH Place
     If PLACES And (RID <> 0) Then
        Call FF1GetRec(RID)
        Call GetBDEvents(BD$, BPL$, DD$, DPL$, xAGE$)
        If RTrim$(DPL$) <> "" Then
           Mid$(PDATA$, xo, 3) = RFLBL$(7)
           x1 = InStr(xo, PDATA$, "   ")
           If x1 > xo Then
              xo = x1 - (x1 > xo) + (MID$(PDATA$, x1, 1) = ".")
           End If
           Mid$(PDATA$, xo, 255) = LTrim$(DPL$)
           xo = InStr(xo, PDATA$, "   ")
        End If
     End If
     If xo = 1 Then IL = IL + 1
     Return

1270 ' Check for Room for More Data
     If Not CombineDATA Then Return
     xIL = IL: nxtDL = 0: xHIL = 6 - 2 * (CSID <> 0)
     LBP = 0: LDD = 0: LDP = 0: LDD = 0: LDP = 0
1275 Select Case xIL
        Case 4: If PLACES Then LBP = Len(RTrim$(LTrim$(BPL$)))
                If LBP > 0 Then nxtDL = Len(RFLBL$(2)) + LBP + 1
        Case 5: If (DateS And (CSID = 0)) Then LDD = Len(RTrim$(LTrim$(DD$)))
                If LDD > 0 Then nxtDL = Len(RFLBL$(6)) + LDD + 1
        Case 6: If (PLACES And ((IL > 4) Or (CSID = 0))) Then LDP = Len(RTrim$(LTrim$(DPL$)))
                If LDP > 0 Then nxtDL = Len(RFLBL$(7)) + LDP + 1
        Case 7: If DateS And (IL > 5) Then LDD = Len(RTrim$(LTrim$(DD$)))
                If LDD > 0 Then nxtDL = Len(RFLBL$(6)) + LDD + 1
        Case 8: If PLACES And ((IL > 5) Or CombineALL) Then LDP = Len(RTrim$(LTrim$(DPL$)))
                If LDP > 0 Then nxtDL = Len(RFLBL$(7)) + LDP + 1
     End Select
     If (nxtDL = 0) And (xIL < xHIL) Then
        xIL = xIL + 1
        GoTo 1275
     End If
     If (nxtDL > 0) And ((DW - xo) > nxtDL) Then
        IL = xIL + 1: If xo > 1 Then xo = xo + 1
     End If
     Return


1300 ' Process MARRIAGE Data
     CSID = DMID(HX): RID = CSID - (CSID < 0)
     DIL(HX) = IL: IL = DMIL(HX): MRN = DMRN(HX)
1305 If IL = -1 Then GoSub 1320 ' Get Next Marriage Record
     DUPID = (IL < 0): IL = Abs(IL)
     If IL = 0 Then DIL(HX) = DIL(HX) + 1: GoTo 1310 ' No more Data
     If IL = 1 Then GoSub 1330 ' Print Separator Line
     If IL = 2 Then GoSub 1360 ' Print NAME Line(1)
     If IL = 3 Then GoSub 1360 ' Print NAME Line(2)
     If IL = 4 Then GoSub 1220 ' Print BIRTH Date after create REF entry
     If IL = 5 Then GoSub 1240 ' Print BIRTH Place
     If IL = 6 Then GoSub 1340 ' Print MARRIAGE Date
     If IL = 7 Then GoSub 1350 ' Print MARRIAGE Place
     If IL = 8 Then GoSub 1345 ' Print End of MARRIAGE Date
     If IL = 9 Then GoSub 1355 ' Print End of MARRIAGE Place
     If IL = 10 Then GoSub 1250 ' Print DEATH Date
     If IL = 11 Then GoSub 1260 ' Print DEATH Place
     IL = IL + 1: If (IL > 11) Then IL = -1: DUPID = 0
     If DUPID Then
        If IL < 6 Then IL = -6 Else If IL > 9 Then IL = -1 Else IL = -IL
     End If
     If xo = 1 Then GoTo 1305
1310 DMIL(HX) = IL: IL = DIL(HX): CSID = 0
     Return

1320 ' Get Next MARRIAGE Record
     MRN = NDMRN(HX): DMRN(HX) = MRN
     If MRN = 0 Then IL = 0: Return
     Call FF3GetRec(MRN)
     If DID(HX) = FF3SPOUSE.SP1ID Then
        CSID = FF3SPOUSE.SP2ID: NDMRN(HX) = FF3SPOUSE.SP1NXT
        Else
        CSID = FF3SPOUSE.SP1ID: NDMRN(HX) = FF3SPOUSE.SP2NXT
     End If
     If Not ALLSPOUSE Then NDMRN(HX) = 0
     CSID = CSID + (CSID = 0): DMID(HX) = CSID
     RID = CSID - (CSID < 0): IL = 1
     Return

1330 ' Print Separator Line
1331 LSet PDATA$ = MID$(BOXTOP$, 2, DW): xo = DW
1332 Return

1340 ' Print MARRIAGE Date
     If DateS Then
        Call FF3GetRec(MRN): MLOC = FF3SPOUSE.MLOC
        dfmt$ = LTrim$(RTrim$(XRPTDate$(FF3SPOUSE.D1MD, FF3SPOUSE.D1Y, 1)))
        LMD = Len(dfmt$)
        If LMD > 0 Then
           xo = xo - (xo > 1)
           Mid$(PDATA$, xo, 14) = RFLBL$(4)
           xo = InStr(xo, PDATA$, "   ")
           If xo > 1 Then If MID$(PDATA$, xo - 1, 1) <> "." Then xo = xo + 1
           Mid$(PDATA$, xo, 12) = dfmt$
           xo = xo + LMD + 1
           If CombineDATA And PLACES And (MLOC <> 0) Then
              Call FF3GetRec(MLOC)
              LMP = Len(RTrim$(LTrim$(FF3PLACE.PLACE1)))
              If ((DW - xo) > (LMP + Len(RFLBL$(5)))) Then
                 IL = IL + 1
              End If
           End If
        End If
     End If
     If xo = 1 Then IL = IL + 1
     Return

1345 ' Print End of MARRIAGE Date
     If AllDATES Then
        Call FF3GetRec(MRN): MLOC = FF3SPOUSE.MLOC
        dfmt$ = LTrim$(RTrim$(XRPTDate$(FF3SPOUSE.D2MD, FF3SPOUSE.D2Y, 1)))
        LMD = Len(dfmt$)
        If LMD > 0 Then
           xo = xo - (xo > 1)
           Mid$(PDATA$, xo, 14) = RFLBL$(9)
           xo = InStr(xo, PDATA$, "   ")
           If xo > 1 Then If MID$(PDATA$, xo - 1, 1) <> "." Then xo = xo + 1
           Mid$(PDATA$, xo, 12) = dfmt$
           xo = xo + LMD + 1
           If CombineDATA And PLACES And (MLOC <> 0) Then
              Call FF3GetRec(MLOC)
              LMP = Len(RTrim$(LTrim$(FF3PLACE.PLACE2)))
              If ((DW - xo) > (LMP + Len(RFLBL$(5)))) Then
                 IL = IL + 1
              End If
           End If
        End If
     End If
     If xo = 1 Then IL = IL + 1
     Return

1350 ' Print MARRIAGE Place
     If PLACES Then
        Call FF3GetRec(MRN): MLOC = FF3SPOUSE.MLOC
        If MLOC <> 0 Then
           LSet MPLACE = ""
           Call FF3GetRec(MLOC)
           If LTrim$(MPLACE) <> "" Then
              xo = xo - (xo > 1)
              Mid$(PDATA$, xo, 3) = RFLBL$(5)
              xo = InStr(xo, PDATA$, "   ")
              If xo > 1 Then If MID$(PDATA$, xo - 1, 1) <> "." Then xo = xo + 1
              Mid$(PDATA$, xo, 255) = MPLACE
              xo = xo + Len(RTrim$(MPLACE)): FL = xo - 1
           End If
        End If
     End If
     If (xo = 1) Then IL = IL + 1
     Return

1355 ' Print End of MARRIAGE Place
     If AllDATES And PLACES Then
        Call FF3GetRec(MRN): MLOC = FF3SPOUSE.MLOC
        If MLOC <> 0 Then
           LSet DVPLACE = ""
           Call FF3GetRec(MLOC)
           If LTrim$(DVPLACE) <> "" Then
              xo = xo - (xo > 1)
              Mid$(PDATA$, xo, 3) = RFLBL$(10)
              xo = InStr(xo, PDATA$, "   ")
              If xo > 1 Then If MID$(PDATA$, xo - 1, 1) <> "." Then xo = xo + 1
              Mid$(PDATA$, xo, 255) = DVPLACE
              xo = xo + Len(RTrim$(DVPLACE)): FL = xo - 1
           End If
        End If
     End If
     If (xo = 1) Then IL = IL + 1
     Return

1360 ' Print SPOUSE Name
     If IL = 2 Then
        Mid$(PDATA$, xo, 2) = RFLBL$(3)
        xo = InStr(xo, PDATA$, "   ")
        If xo > 1 Then If MID$(PDATA$, xo - 1, 1) <> "." Then xo = xo + 1
     End If
     GoSub 1200
     Return

1400 ' Build NODE Line for Horizontal Arrangement
     LSet PDATA$ = String$(BOXWIDTH, Asc(HLC$)): PCOL = 0: CA$ = " "
     HX = HMIN
     While HX > -1: IL = DIL(HX)
     If IL < 1 Then
        If PCOL = 0 Then LSet DDATA$ = " ": GoTo 1420 Else LSet DDATA$ = PDATA$
        If (PCOL = HX - 1) Then
           PPS = PS - FNCW(HX - 1, 0)
           If PPS > 0 Then Mid$(DDATA$, PPS, 1) = xPC$: GoTo 1420
        End If
        If (PCOL = HX) Then If PS <= FNCW(HX, 0) Then Mid$(DDATA$, PS, 1) = xPC$
        GoTo 1420
     End If
     If (DCHNUM(HX) And 256) = 0 Then GoTo 1410 'If not the FIRST child then go on
     If (DCHNUM(HX) And 512) > 0 Then Mid$(DDATA$, PS, 1) = VLC$: GoTo 1420
     PCOL = xPCOL(HX)
     Mid$(DDATA$, PS + 1, BOXWIDTH - PS) = PDATA$
     If PCOL = HX Then Mid$(DDATA$, PS, 1) = MID$(DCHAR$, 16, 1): GoTo 1420
     If (PCOL = HX + 1) Then
        PPS = FNCW(HX, 0) + PS
        If PPS <= BOXWIDTH Then Mid$(DDATA$, PPS, 1) = xPC$
     End If
     Mid$(DDATA$, PS, 1) = xFCC$: GoTo 1420
1410 If (DCHNUM(HX) And 512) > 0 Then   ' If not the LAST child then go on
        Mid$(DDATA$, 1, PS - 1) = PDATA$: Mid$(DDATA$, PS, 1) = xLCC$
        If (PCOL = HX) Then
           Mid$(DDATA$, PS, 1) = MID$(DCHAR$, 15, 1)
           Else
           If (PCOL = HX - 1) Then
              PPS = PS - FNCW(HX - 1, 0)
              If PPS > 0 Then Mid$(DDATA$, PPS, 1) = xPC$
           End If
           Mid$(DDATA$, PS, 1) = xLCC$
        End If
        PCOL = 0: GoTo 1420
     End If
     LSet DDATA$ = PDATA$
     If PCOL = HX Then Mid$(DDATA$, PS, 1) = xC$: GoTo 1420
     If (PCOL = HX - 1) Then
        PPS = PS - FNCW(HX - 1, 0)
        If PPS > 0 Then Mid$(DDATA$, PPS, 1) = xPC$
     End If
     If (PCOL = HX + 1) Then
        PPS = FNCW(HX, 0) + PS
        If PPS <= BOXWIDTH Then Mid$(DDATA$, PPS, 1) = xPC$
     End If
     Mid$(DDATA$, PS, 1) = xCC$
1420 GoSub 1140
     HX = HX + (1 - (IL <> 0)) * HSTEP
     If ((HSTEP = 1) And (HX > HXMAX)) Or ((HSTEP = -1) And (HX < HXMAX)) Then HX = -1
     Wend
     LSet PDATA$ = " "
     Return

2200 ' Print Column Heading Line
     If PRTHEAD And (Not PAGETOP) And (a <> 27) Then
        Call PrintHEAD: PAGETOP = -1: BLEFT = PAGE.BLENGTH
        If a <> 27 Then
           HL$ = PLINE$: HOFST = PAGE.LOFST: LSet PLINE$ = "": PAGE.LOFST = 0
           HXMAX = NodeHDR.WFMAXCOL
           xB = (BOXSEP + 2) \ 2
           For HX = 1 To HXMAX Step 2
              RSet NUMED$ = fnx$((HX + 1) \ 2)
              XL = Len(LTrim$(NUMED$))
              x1 = xB + (BOXWIDTH - XL) \ 2
              If x1 < xB Then x1 = xB: XL = BOXWIDTH - xB
              Mid$(DDATA$, 1, BOXWIDTH) = String$(BOXWIDTH, Asc("-"))
              Mid$(DDATA$, xB, 1) = ":"
              Mid$(DDATA$, x1 + 1, XL) = Right$(NUMED$, XL)
              If (HX = 1) Then Mid$(DDATA$, 1, xB) = Space$(xB)
              GoSub 1140
           Next HX
           If (((NodeHDR.WFMAXCOL + 1) Mod 2) = 1) Then
              HX = NodeHDR.WFMAXCOL
              XL = FNCW(NodeHDR.WFMAXCOL + 1, 0)
              Mid$(DDATA$, 1, XL) = String$(XL, Asc("-"))
              Mid$(DDATA$, xB, 1) = ":"
              GoSub 1140
           End If
           GoSub 1165: X = BH + 1
           If a <> 27 Then
              While (X > 0) And (a <> 27)
                  Call WriteLINE
                  X = X - 1
              Wend
           End If
           BLEFT = BLEFT - BH - 1
           LSet PLINE$ = HL$: PAGE.LOFST = HOFST: PAGETOP = 0
        End If
     End If
     OKAY = (a <> 27)
     Return
End Sub

Sub HBOXSumry(CTYPE)
500 ' Produce Chart Summary Report
    Select Case CTYPE
       Case 1: RPT$ = "AHBX": ANCSEARCH = -1
       Case 2: RPT$ = "DHBX"
       Case 3: RPT$ = "FPTH": FamPath = -1
    End Select
    Call RptLOAD(RPT$): If Not OKAY Then GoTo 550
    Call RptOptGET("DWTH", DW): If DW = 0 Then DW = 15
    Call RptOptGET("BSEP", BOXSEP)
    BOXWIDTH = DW + 2 + BOXSEP
    Call RptOPEN("HSUM"): If Not OKAY Then GoTo 550
    Call GetRVAR("MULTIPLY", X): If FOUND Then FF# = Val(RVAR$(X))
    Call FAMOpen("READ"): If Not OKAY Then GoTo 550
    BRID = 0
    Call NodeWrkOPEN("READQ")
    If (Not OKAY) Or (NodeHDR.WFBRID = 0) Then GoTo 550
    Align = Asc(NodeHDR.WFOPT) And 15
    Select Case CTYPE
       Case 1: If (Asc(NodeHDR.WFOPT) And 16) = 0 Then GoTo 550
       Case 2: If (Asc(NodeHDR.WFOPT) And 48) > 0 Then GoTo 550
       Case 3: If (Asc(NodeHDR.WFOPT) And 32) = 0 Then GoTo 550
    End Select
    Align = Asc(NodeHDR.WFOPT) And 15
    If (Align < 1) Or (Align = 5) Then
       Call PutMSG("HBOXM021"): GoTo 550
    End If
    BRID = NodeHDR.WFBRID: Call FF1GetRec(BRID)
    BNAME$ = FMTNAME$(1): WALIGN = Align: Call ShowINFO
    SysVAR.NVAR1 = BRID: SysVAR.SUBJECT = BNAME$
    Call PutMSG("HBOXM022")
    If CA$ = Chr$(27) Then GoTo 550
    ORIENT = (-(CA$ <> "2") - 2 * (CA$ = "2"))
    If NodeHDR.WFMAXCOL = 1 Then
       ORIENT2 = 1
       Else
       If ANCSEARCH Then
          Call PutMSG("HBOXM023")
          Else
          Call PutMSG("HBOXM024")
       End If
       If CA$ = Chr$(27) Then GoTo 550
       ORIENT2 = (-(CA$ <> "2") - 2 * (CA$ = "2"))
    End If
    Call PrintOPEN
    If Not OKAY Then GoTo 550
    If Len(PLINE$) < 100 Then PLINE$ = Space$(100)

515 ' Print Report
    PRT.SECTION = " "
    Call PrintHEAD
    xCTYPE$ = "CTYPE" + fnx$(CTYPE)
    Call GetRVAR(xCTYPE$, X): If FOUND Then xCTYPE$ = RVAR$(X)
    If FamPath Then
       If NOSIBLING Then
          Call GetRVAR("WOUTSIB", X)
          Else
          Call GetRVAR("WITHSIB", X)
       End If
       If FOUND Then xCTYPE$ = xCTYPE$ + " " + RVAR$(X)
    End If
    SysVAR.MVAR = xCTYPE$:
    PRT.SECTION = "1": Call HTFPrint("T")
    Call WriteLINE
    If Not FamPath Then PRT.SECTION = "2" Else PRT.SECTION = "3"
    Call HTFPrint("T")
    If FamPath Then
       Call FF1GetRec(NodeHDR.WFID1)
       SysVAR.NVAR1 = NodeHDR.WFID1: SysVAR.SUBJECT = FMTNAME$(1)
       PRT.SECTION = "4": Call HTFPrint("T")
       Call FF1GetRec(NodeHDR.WFID2)
       SysVAR.NVAR1 = NodeHDR.WFID2: SysVAR.SUBJECT = FMTNAME$(1)
       PRT.SECTION = "5": Call HTFPrint("T")
    End If
    Call WriteLINE
    SysVAR.NVAR1 = NodeHDR.WFNUMREL: SysVAR.NVAR2 = NodeHDR.WFHIGL
    SysVAR.NVAR3 = (NodeHDR.WFMAXCOL + 1) \ 2
    PRT.SECTION = "A": Call HTFPrint("T")
    Call WriteLINE
    Call GetRVAR("GARANGE1", X)
    If FOUND Then
       SysVAR.MVAR = RVAR$(X)
       PRT.SECTION = "B": Call HTFPrint("T")
    End If
    SysVAR.MVAR = "": SysVAR.SUBJECT = ""
    xORIENT$ = "BASELOC" + fnx$(ORIENT): Call GetRVAR(xORIENT$, X)
    If FOUND Then SysVAR.MVAR = RVAR$(X)
    If ANCSEARCH Then
       xORIENT$ = "PARANGE" + fnx$(ORIENT2)
       PRT.SECTION = "C"
       Else
       xORIENT$ = "CARANGE" + fnx$(ORIENT2)
       PRT.SECTION = "D"
    End If
    Call GetRVAR(xORIENT$, X): If FOUND Then SysVAR.SUBJECT = RVAR$(X)
    Call HTFPrint("T")
    xALIGN$ = "ALIGN" + fnx$(Align): Call GetRVAR(xALIGN$, X)
    If FOUND Then
       SysVAR.MVAR = RVAR$(X)
       PRT.SECTION = "E": Call HTFPrint("T")
    End If
    Call WriteLINE
    'BW# = BOXWIDTH * ((NodeHDR.WFMAXCOL + 1) \ 2) + FNCW(NodeHDR.WFMAXCOL + 1, 0) * ((NodeHDR.WFMAXCOL + 1) Mod 2)
    BW# = (NodeHDR.WFMAXCOL + 1) \ 2
    BW# = BW# * BOXWIDTH
    xZ# = FNCW(NodeHDR.WFMAXCOL + 1, 0)
    xZ# = xZ# * ((NodeHDR.WFMAXCOL + 1) Mod 2)
    BW# = BW# + xZ#
    TW# = BW# + PAGE.IMARGIN + PAGE.OMARGIN
    ts = Fix((TW# + PRT.FWIDTH - 1) / PRT.FWIDTH)
    SysVAR.NVAR1 = DW: SysVAR.NVAR2 = BOXSEP: SysVAR.MVAR = LTrim$(Str$(BW#))
    PRT.SECTION = "F": Call HTFPrint("T")
    SysVAR.MVAR = LTrim$(Str$(TW# * FF#))
    PRT.SECTION = "G": Call HTFPrint("T")
    SysVAR.NVAR1 = PRT.FWIDTH
    SysVAR.MVAR = "(" + fnx$(PAGE.IMARGIN) + "," + fnx$(PAGE.OMARGIN) + ")"
    SysVAR.NVAR3 = ts
    PRT.SECTION = "H": Call HTFPrint("T")
    If (Align <> 3) And (Align <> 5) And (Align <> 7) Then
       c1 = (PRT.FWIDTH - PAGE.IMARGIN - PAGE.OMARGIN) \ BOXWIDTH
       TS2 = 1 + (((NodeHDR.WFMAXCOL + 1) \ 2) - 1) \ c1
       If ts <> TS2 Then
          SysVAR.NVAR1 = TS2
          PRT.SECTION = "I": Call HTFPrint("T")
       End If
    End If
    Call WriteLINE
    PRT.SECTION = "T": Call HTFPrint("T")
    xX$ = " "
    SCW = PRT.FWIDTH - PAGE.IMARGIN - PAGE.OMARGIN - 30
    If SCW < 20 Then SCW = 20
    ReDim GLBOX(SCW)
    CD# = (NodeHDR.WFMAXCOL + 1) / (2 * SCW)
    If CD# < 1 Then CD# = 1: SCW = (NodeHDR.WFMAXCOL + 1) \ 2
    If ORIENT = 1 Then
       BGNGL = 0: ENDGL = NodeHDR.WFHIGL: GLSTEP = 1
       Else
       BGNGL = NodeHDR.WFHIGL: ENDGL = 0: GLSTEP = -1
    End If
    NUMED$ = Space$(5)
    For GL = BGNGL To ENDGL Step GLSTEP
        Call NodeWorkGET("GLREC", GL)
        RSet NUMED$ = fnx$(NodeGLREC.WFGLNUM)
        Mid$(PLINE$, 2, 5) = NUMED$
        X = ((NodeGLREC.WFLOCOL + 1) / 2): Y = (NodeGLREC.WFHICOL + 1) / 2
        If ORIENT2 = 2 Then
           Z = X: X = 1 + ((NodeHDR.WFMAXCOL + 1) \ 2) - Y
           Y = 1 + ((NodeHDR.WFMAXCOL + 1) \ 2) - Z
        End If
        RSet NUMED$ = fnx$(X): Mid$(PLINE$, 8, 5) = NUMED$
        RSet NUMED$ = fnx$(Y): Mid$(PLINE$, 14, 5) = NUMED$
        RSet NUMED$ = fnx$(NodeGLREC.WFGLEND - NodeGLREC.WFGLBGN + 1)
        Mid$(PLINE$, 20, 5) = NUMED$
        xo = 30
        GLBGN = NodeGLREC.WFGLBGN: GLEND = NodeGLREC.WFGLEND: SCHI = 0
        For X = GLBGN To GLEND
            Call NodeWorkGET("NODEREC", X)
            Col = (NodeREC.WFCOL + 1) / (2 * CD#): If Col < 1 Then Col = 1
            If ORIENT2 = 2 Then Col = SCW - Col + 1
            If Col > SCHI Then SCHI = Col
            If GLBOX(Col) > -1 Then
               If NodeREC.WFCHBGN < 0 Then GLBOX(Col) = -1 Else GLBOX(Col) = 1
            End If
        Next X
        For X = 1 To SCHI
            If GLBOX(X) = -1 Then
               LSet xX$ = "x"
               Else
               If GLBOX(X) = 1 Then
                  LSet xX$ = "o"
                  Else
                  LSet xX$ = " "
               End If
            End If
            Mid$(PLINE$, X + 30, 1) = xX$: GLBOX(X) = 0
        Next X: Call PrintLINE
    Next GL
    Call RptCLOSE
    If PRT.AGAIN Then GoTo 515
    Erase GLBOX

550 ' End of Summary Report Option
    Call FamCLOSE
End Sub


Sub ShowINFO()
3900 ' Display BRID Information
    'CurFMT = FMT001: Fmt = FMTTAB(CurFMT)
    'x = Fmt.BGNFLD
    'While x > 0
    '   FFLD = FLDTAB(x)
    '   If FFLD.UTXT > 0 Then
    '      XX$ = UTXT$(FFLD.UTXT)
    '      Select Case FFLD.NAME
    '         Case "BRID": If BRID > 0 Then LSet XX$ = Str$(BRID) Else LSet XX$ = ""
    '         Case "NAME": If BRID > 0 Then LSet XX$ = BNAME$ Else LSet XX$ = ""
    '         Case "NMRL": If BRID > 0 Then LSet XX$ = Str$(NodeHDR.WFNUMREL) Else LSet XX$ = ""
    '         Case "HIGL": If BRID > 0 Then LSet XX$ = Str$(NodeHDR.WFHIGL) Else LSet XX$ = ""
    '         Case "HICO": If BRID > 0 Then LSet XX$ = Str$((NodeHDR.WFMAXCOL + 1) \ 2) Else LSet XX$ = ""
    '         Case "ALIG": LSet XX$ = RTrim$(ALIGNTAB$(WALIGN))
    '      End Select
    '      LSet UTXT$(FFLD.UTXT) = XX$
    '   End If
    '   x = FLDCHN(x).FWD
    'Wend
    'Call WinDisplayDATA(0)

End Sub

