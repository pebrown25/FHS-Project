DECLARE SUB BldChartDEF ()
DECLARE SUB BldNodeWRK (CTYPE%)
DECLARE SUB HBOXChart (CTYPE%)
DECLARE SUB HBOXSumry (CTYPE%)
DECLARE SUB NodeWorkGET (RTP$, RNUM%)
DECLARE SUB NodeWorkOPEN (OPT$)
DECLARE SUB NodeWorkPUT (RTP$, RNUM%)
DECLARE SUB ShowINFO ()
DECLARE FUNCTION DATEMD% (XDATE$, x%)
DECLARE FUNCTION DATEY% (XDATE$, x%)
DECLARE FUNCTION FDSTAT$ (x$)
DECLARE FUNCTION FMTNAME$ (NFMT%)
DECLARE FUNCTION SFILL$ (N%, L%)
DECLARE FUNCTION UCX$ (x$)
DECLARE FUNCTION XRPTDate$ (MD%, Y%, O%)
1  REM $INCLUDE: 'FHSCOMON.BAS'
   TYPE NodeFileHDR
      FTYPE    AS STRING * 1    ' "W" for "W"ork File
      WFTYPE   AS STRING * 1    ' 6=Node Work File
      WFVER    AS STRING * 1    ' 1
      WFOPT    AS STRING * 1    ' ....XXXX = ALIGN Option (0 < XXXX < 9)
                                ' ...1.... = Ancestor Search
                                ' ..1..... = Family Path
                                ' .1...... = No Siblings
      WFOPT2   AS STRING * 1    ' ...1.... = Adoptions Included
                                ' .......1 = Male Lines only
                                ' ......1. = Female Lines only
      WFCRYR   AS INTEGER
      WFCRMD   AS INTEGER
      WFUPDTY  AS INTEGER
      WFUPDTMD AS INTEGER
      ' Record 2
      REORGY   AS INTEGER
      REORGMD  AS INTEGER
      REORGTM  AS INTEGER
      WFID1    AS INTEGER       ' FamPath ID1
      WFID2    AS INTEGER       ' FamPath ID2
      FILLER1  AS INTEGER
      ' Record 3
      WFMAXGL  AS INTEGER
      WFHIGL   AS INTEGER
      WFBRID   AS INTEGER       ' Base Record ID or Common Ancestor
      WFNUMREL AS INTEGER
      WFMAXCOL AS INTEGER
      WFRGL1   AS STRING * 1    ' FamPath RGL1
      WFRGL2   AS STRING * 1    ' FamPath RGL2
   END TYPE
   TYPE NodeFileGLREC
      WFGLNUM  AS INTEGER
      WFGLBGN  AS INTEGER
      WFGLEND  AS INTEGER
      WFLOCOL  AS INTEGER
      WFHICOL  AS INTEGER
      FILLER   AS INTEGER
   END TYPE
   TYPE NodeFileREC
      WFPREC   AS INTEGER
      WFCHNUM  AS INTEGER
      WFRID    AS INTEGER
      WFCOL    AS INTEGER
      WFCHBGN  AS INTEGER
      WFNUMCH  AS INTEGER
   END TYPE

   PN$ = " FAMILY HISTORY SYSTEM - FHSHBOX - Horizontal BOX Chart Program"
   CY$ = " (C) Copyright 1992,1993,1994,1996 by Phillip E. Brown"
   IF CONFIG.BP = 0 THEN RUN "FHSINIT"
   DIM NodeHDR AS NodeFileHDR, NODEREC AS NodeFileREC, NodeGLREC AS NodeFileGLREC
   DEF FNX$ (A) = LTRIM$(STR$(A))
   DEF FNDS$ (A#) = STR$(VAL(STR$(A#)))
   DEF FNODD (A#) = (VAL(RIGHT$(FNDS$(A#), 1)) MOD 2)
   DEF FNCW (A, B) = BOXWIDTH + (B = 0) * (((BOXWIDTH + 1) \ 2) - ((A + 1) MOD 2) * (BOXWIDTH MOD 2))
   DEF FNCOL (A) = A + (2 * A - NodeHDR.WFMAXCOL - 1) * (ORIENT2 = 2)
   REDIM HCTYPE$(3)
   CALL TableLOAD("WORKTYPE", T, 39)
   IF OKAY THEN
      x = TABLE(T).FT
      WHILE x > 0
        Z = VAL(MID$(TABDATA$(x), 1, 2))
        IF (8 < Z) AND (Z < 12) THEN
          HCTYPE$(Z - 8) = RTRIM$(MID$(TABDATA$(x), 15, 255))
        END IF
        x = TABDATACHN(x).FWD
      WEND
      CALL TableFREE("WORKTYPE")
   END IF
   CALL TableLOAD("ALIGN", T, 0)
   REDIM ALIGNTAB$(TABLE(T).SIZE)
   x = TABLE(T).FT
   WHILE x > 0
      Y = VAL(MID$(TABDATA$(x), 1, 1))
      ALIGNTAB$(Y) = MID$(TABDATA$(x), 3, 8)
      x = TABDATACHN(x).FWD
   WEND
   CALL TableFREE("ALIGN")

   ' Create Chart Symbol String for Screen Prints
   SCHAR$ = SPACE$(31): DCHAR$ = SCHAR$
   FOR x = 1 TO 31: READ Z: MID$(SCHAR$, x, 1) = CHR$(Z): NEXT x
   DATA 218,193,191,179,180,195,196,192,194,217,218,193,191,179,180,195,196,192,194,217,197,201,207,187,186,182,199,205,200,209,188
   ' Prepare Reports & Screen Formats
   CALL RptLOAD("DHBX"): CALL RptLOAD("AHBX")
   CALL RptLOAD("HSUM"): CALL RptLOAD("FPTH")
   CTYPE = MO
   CALL FmtFIND("XXXXS010"): FMT010 = CURFMT
   CALL WinINIT("LR1")
   CALL WinPREP(1): PGMRTRN$ = "FHSMENU"
   FUN = 0: USENEST = 0: GLMAX = 0: PSPOUSE = 0: CRULE = 0: ALIN# = 0
   READ$ = "READQ"

40 ' Format Display
   CALL WinCLR
   CALL FmtFIND("HBOXS001"): FMT001 = CURFMT
   CALL WinFORMAT(FMT001): WIN.OFMT = FMT001
   CALL WinDisplayLABELS(FMT001)

45 ' Refresh Option DISPLAY
    SELECT CASE CTYPE
       CASE 1: RPT$ = "AHBX"
       CASE 2: RPT$ = "DHBX"
       CASE 3: RPT$ = "FPTH"
    END SELECT
    RTYPE$ = HCTYPE$(CTYPE)
    CALL RptLOAD(RPT$)
    CALL WinPREP(2): CALL WinCLR: OPTTable.CL = 0
    CALL RptOptDISPLAY(RPT$)
    CALL WinSWITCH(1)

50 ' Display File, Printer SETUPS
   CURFMT = FMT001: FMT = FMTTAB(CURFMT)
   x = FMT.BGNFLD
   WHILE x > 0
      FFLD = FLDTAB(x)
      IF FFLD.UTXT > 0 THEN
         x$ = UTXT$(FFLD.UTXT)
         SELECT CASE FFLD.NAME
            CASE "FSUP": LSET x$ = FDFSetup.NAME
            CASE "FSDS": LSET x$ = FDFSetup.DESC
            CASE "PSUP": LSET x$ = PDFSetup.NAME
            CASE "PSDS": LSET x$ = PDFSetup.DESC
            CASE "PRTR": LSET x$ = PDFSetup.PRINTER
            CASE "FMWD": LSET x$ = STR$(PDFForms.FWIDTH)
            CASE "FMLT": LSET x$ = STR$(PDFForms.FLENGTH)
         END SELECT
         LSET UTXT$(FFLD.UTXT) = x$
      END IF
      x = FLDCHN(x).FWD
   WEND

100 ' Check Files
    CALL FamOPEN(READ$): READ$ = "READQ"
    FFOKAY = OKAY: BRID = 0: BNAME$ = "": WALIGN = 0
    CALL FmtFindFLD("NMCT"): LSET UTXT$(FFLD.UTXT) = STR$(FF1HDR.MAXID)
    IF FFOKAY THEN
       CALL NodeWorkOPEN("READQ"): NFOKAY = OKAY
       IF NFOKAY AND (NodeHDR.WFBRID <> 0) THEN
          WFOPT = ASC(NodeHDR.WFOPT)
          WFOPT2 = ASC(NodeHDR.WFOPT2)
          IF ((CTYPE = 1) AND ((WFOPT AND 16) > 0)) OR ((CTYPE = 2) AND ((WFOPT AND 48) = 0)) OR ((CTYPE = 3) AND ((WFOPT AND 32) > 0)) THEN
             CALL FF1GetRec(NodeHDR.WFBRID)
             IF OKAY THEN
                WALIGN = WFOPT AND 15
                BRID = NodeHDR.WFBRID: BNAME$ = FMTNAME$(1)
             END IF
          END IF
       END IF
    END IF
    CALL FamCLOSE
    CALL ShowINFO

200 ' Get Processing Option
    CALL PUTMSG("")
    IF FUN$ <> "" THEN CALL OptHILITE(FUN$, FUN$, 1)
    CALL GetKEY
    IF LEN(A$) < 2 THEN SOUND BP, DUR: FUN = 0: GOTO 200:  ELSE FUN = A
    SELECT CASE FUN
       CASE 30: CTYPE = 2: GOTO 45
       CASE 32: CTYPE = 1: GOTO 45
       CASE 59: FUN$ = "F1"
       CASE 60: FUN$ = "F2"
       CASE 61: FUN$ = "F3"
       CASE 62: FUN$ = "F4"
       CASE 64: FUN$ = "F6"
       CASE 65: FUN$ = "F7"
       CASE 66: CALL DOSShell: GOTO 40
       CASE 67: FUN$ = "F9"
       CASE 94: FUN$ = "F1"
       CASE 95: FUN$ = "F2"
       CASE 109: FUN$ = "F6": PGMRTRN$ = "FHSRPTS"
       CASE ELSE: CALL ErrBEEP(0): FUN = 0: GOTO 200
    END SELECT
    CALL OptHILITE(FUN$, FUN$, 6)
    SELECT CASE FUN$
       CASE "F1": GOTO 300
       CASE "F2": GOTO 310
       CASE "F3": CALL BldNodeWRK(CTYPE): GOTO 100
       CASE "F4": GOTO 350
       CASE "F6": GOTO M1000
       CASE "F7": CTYPE = CTYPE + 1: IF CTYPE > 3 THEN CTYPE = 1
                  MO = CTYPE: GOTO 45
       CASE "F9": GOTO 250
    END SELECT
    GOTO 200

250 ' Return to FAMMENU.BAS
    CALL FamCLOSE
    REDIM RFLTH(1), RFLBL$(1)
    CALL PgmPREP(PGMRTRN$)
    IF (NOT OKAY) OR (A = 27) THEN LSET PGMRTRN$ = "FHSMENU": GOTO 40
    IF PGMRTRN$ <> "FHSRPTS" THEN
       CALL RptFREE("AHBX"): CALL RptFREE("DHBX")
       CALL RptFREE("HSUM"): CALL RptFREE("FPTH")
    END IF
    CHAIN PGMRTRN$

300 ' Change FILE Parameters
    IF A = 94 THEN LSET PGMRTRN$ = "FHSFILE": GOTO 250
    HSU = ENV.FSETUP
    CALL WinSWITCH(2)
    CALL FDFSelect: READ$ = "READ"
    IF HSU <> ENV.FSETUP THEN BRID = 0
    GOTO 45

310 ' Change PRINTER Parameters
    IF A = 95 THEN LSET PGMRTRN$ = "FHSPRTC": GOTO 250
    CALL WinSWITCH(2)
    CALL PDFSelect
    GOTO 45

350 ' Process Request to Update Options
    CALL PUTMSG("HBOXM001")
    IF A = 27 THEN GOTO 45
    CALL WinSWITCH(2)
    IF A$ = "1" THEN
       XRPT$ = RPT$
       ELSE
       XRPT$ = "HSUM": CALL WinCLR: CALL RptOptDISPLAY("HSUM")
    END IF
    CALL RptOptUPDATE(XRPT$)
    IF XRPT$ = "HSUM" THEN CALL WinCLR: CALL RptOptDISPLAY(RPT$)
    GOTO 45

M1000: 'PROCESS REPORT OPTIONS
     IF PGMRTRN$ = "FHSRPTS" THEN GOTO 250
     IF BRID = 0 THEN
        CALL PUTMSG("HBOXM002")
        GOTO 200
     END IF
     CALL PUTMSG("HBOXM003")
     IF A$ = CHR$(27) THEN 200
     A = VAL(A$): IF A < 1 OR A > 2 THEN A = 1
     IF A < 2 THEN
        CALL HBOXSumry(CTYPE)
        ELSE
        CALL HBOXChart(CTYPE)
     END IF
     IF ENV.ScrnCLR THEN GOTO 40 ELSE GOTO 100

REM $DYNAMIC
SUB BldChartDEF
      RFCNT = 10
      REDIM RFLTH(RFCNT), RFLBL$(RFCNT)

59510 ' Create Report Definition Tables
      x = 1: RFNAME$ = "01030": GOSUB 59505    'BIRTH DATE
      x = 2: RFNAME$ = "01032": GOSUB 59505    'BIRTH PLACE
      x = 3: RFNAME$ = "02011": GOSUB 59505    'SPOUSE
      x = 4: RFNAME$ = "02020": GOSUB 59505    'MARRIAGE DATE
      x = 5: RFNAME$ = "02021": GOSUB 59505    'MARRIAGE PLACE
      x = 6: RFNAME$ = "01040": GOSUB 59505    'DEATH DATE
      x = 7: RFNAME$ = "01042": GOSUB 59505    'DEATH PLACE
      x = 8: RFNAME$ = "02032": GOSUB 59505    'END MARRIAGE REASON
      x = 9: RFNAME$ = "02030": GOSUB 59505    'END MARRIAGE DATE
      x = 10: RFNAME$ = "02031": GOSUB 59505   'END MARRIAGE PLACE
      GOTO 59560

59505 ' Process Report Field
      CALL GetRFLD(RFNAME$, FX)
      IF FOUND THEN
         RDFFld = RFLDS(FX)
         IF RFLTH(x) = 0 THEN RFLTH(x) = ASC(RDFFld.RLTH)
         ALTH = ASC(RDFFld.ALTH)
         RFLBL$(x) = MID$(RFLD$(FX), 1, ALTH)
      END IF
      RETURN

59560 ' Return to Calling Program
      OKAY = -1
END SUB

REM $STATIC
SUB BldNodeWRK (CTYPE)
SHARED FMT001, ANCSEARCH, ALIN#, BRID, BNAME$, RTYPE$, WALIGN, ALIGNTAB$()
SHARED NodeHDR AS NodeFileHDR, NODEREC AS NodeFileREC, NodeGLREC AS NodeFileGLREC
    ANCSEARCH = (CTYPE = 1): FAMPATH = (CTYPE = 3)
600 ' Create NODE Work File
    CALL NodeWorkOPEN("WRITEQ"): IF NOT OKAY THEN 605
    CALL FamOPEN("READ"): IF NOT OKAY THEN 695
    WALIGN = ASC(NodeHDR.WFOPT) AND 15: IF WALIGN = 0 THEN 605
    IF (FAMPATH AND ((ASC(NodeHDR.WFOPT) AND 32) = 0)) THEN 605
    IF (ANCSEARCH AND ((ASC(NodeHDR.WFOPT) AND 16) = 0)) THEN 605
    IF ((NOT ANCSEARCH) AND ((ASC(NodeHDR.WFOPT) AND 16) > 0)) THEN 605
    CALL PUTMSG("HBOXM004")
    IF A$ = CHR$(27) THEN 695 ELSE A = VAL(A$)
    IF A = 2 THEN GOTO 700

605 ' Create NEW NodeWork File
    CALL FamCLOSE: AGLMAX = 0: DGLMAX = 0: WALIGN = 0: xSTAT$ = " "
    IF NOT FAMPATH THEN
       SELECT CASE CTYPE
          CASE 1: CALL RptLOAD("AHBX")
                  CALL RptOptGET("MXAL", AGLMAX)
                  CALL RptOptGET("ALNS", XLNS)
          CASE 2: CALL RptLOAD("DHBX")
                  CALL RptOptGET("MXDL", DGLMAX)
                  CALL RptOptGET("DLNS", XLNS)
       END SELECT
       IF (XLNS < 1) OR (XLNS > 3) THEN XLNS = 1
       CALL RptOptGET("ADOP", x): ADOPTOK = (x = 1)
       ELSE
       CALL RptLOAD("FPTH")
    END IF
    CALL NodeWorkOPEN("NEW"): IF NOT OKAY THEN 695
    CALL FamOPEN("READ"): IF NOT OKAY THEN 695
    IF FAMPATH THEN
       GOSUB 1900: DGLMAX = NodeHDR.WFMAXGL
       IF A$ = CHR$(27) THEN 695
       GOTO 615
    END IF
    CALL FmtFIND("XXXXS010")
    CALL FmtFindFLD("BRID")
    IF FOUND THEN RSET UTXT$(FFLD.UTXT) = FNX$(NodeHDR.WFBRID)
610 CALL MsgFORMAT: IF A = 27 THEN 695 ELSE x& = 0
    CALL FmtFindFLD("BRID"): IF FOUND THEN x& = VAL(UTXT$(FFLD.UTXT))
    IF x& = 0 THEN 695
    IF 1 > x& OR x& > FF1HDR.MAXID THEN CALL ErrBEEP(0): GOTO 610
    x = x&: CALL FF1GetRec(x): IF NOT OKAY THEN GOTO 610
    NodeHDR.WFBRID = x
    NodeHDR.WFOPT = CHR$(-64 * NOSIBLING - 32 * FAMPATH - 16 * ANCSEARCH + 5)
    NOSIBLING = 0: ID1 = 0: ID2 = 0
    NodeHDR.WFOPT2 = CHR$(-16 * ADOPTOK + XLNS - 1)
    IF ANCSEARCH THEN NodeHDR.WFMAXGL = AGLMAX ELSE NodeHDR.WFMAXGL = DGLMAX
    BRID = x: BNAME$ = FMTNAME$(1): CALL ShowINFO
615 ' Initialize GL Records
    NodeGLREC.WFGLBGN = 0: NodeGLREC.WFGLEND = 0
    NodeGLREC.WFLOCOL = 0: NodeGLREC.WFHICOL = 0
    FOR GL = 0 TO NodeHDR.WFMAXGL
        NodeGLREC.WFGLNUM = GL
        CALL NodeWorkPUT("GLREC", GL)
    NEXT GL
    REDIM GLFCH(NodeHDR.WFMAXGL + 1), GLLCH(NodeHDR.WFMAXGL + 1)
    REDIM GLCOL(NodeHDR.WFMAXGL + 1), NodePTR(FF1HDR.MAXID)
    GOSUB 670: IF A$ = CHR$(27) THEN 695
    WALIGN = ALIGN: CALL ShowINFO: WALIGN = 5

625 ' Create NODE Records for Relatives
    PREC = 0: ID = NodeHDR.WFBRID
    NUMCH = 0: GLCOL = 0: A = 0: GL = 0
    CALL FF1GetRec(ID)
    IF ANCSEARCH THEN ALIN# = 1: CHBGN = 0 ELSE GOSUB 1970
    GOSUB 660
    IF A$ = CHR$(27) THEN 695
    GLFCH(0) = 1
630 ' Process next GL of Relatives
    BGNREC = GLFCH(GL): IF BGNREC = 0 THEN 680
    GLCOL(GL) = GLCOL: ENDREC = NodeHDR.WFNUMREL + 1: GLLCH(GL) = ENDREC
    GLCOL = 0: GL = GL + 1: IF GL > NodeHDR.WFMAXGL THEN 680
    SYSVAR.MVAR = RTYPE$: SYSVAR.NVAR1 = GL
    CALL PUTMSG("HBOXM005")
    FOR PREC = BGNREC TO ENDREC
       CALL NodeWorkGET("NODEREC", PREC)
       ID = NODEREC.WFCHBGN: IF ID < 0 THEN 635
       NUMCH = 0: FCH = 0: PID = NODEREC.WFRID
       IF ANCSEARCH THEN
          XALIN# = ALIN#: GOSUB 650
          ELSE
          WHILE ID > 0: CALL FF1GetRec(ID)
             OKAY = (FAMPATH OR ADOPTOK)
             IF NOT OKAY THEN
                IF FF1REC.FID = PID THEN LSET xSTAT$ = FDSTAT$("FID") ELSE LSET xSTAT$ = FDSTAT$("MID")
                OKAY = (xSTAT$ <> "*")
             END IF
             IF OKAY THEN
                NUMCH = NUMCH + 1
                NodeHDR.WFNUMREL = NodeHDR.WFNUMREL + 1
                GOSUB 1970: GOSUB 660
                IF A = 27 THEN
                   ID = 0: PREC = ENDREC
                   ELSE
                   IF FCH = 0 THEN FCH = NodeHDR.WFNUMREL + 1
                   IF GLFCH(GL) = 0 THEN GLFCH(GL) = FCH
                   GOSUB 1980
                END IF
                ELSE
                GOSUB 1980
             END IF
          WEND
       END IF
       ' Update FCH and NUMCH in "Parent" Record
       CALL NodeWorkGET("NODEREC", PREC)
       NODEREC.WFNUMCH = NUMCH: NODEREC.WFCHBGN = FCH
       CALL NodeWorkPUT("NODEREC", PREC)
635 NEXT PREC: IF A = 27 THEN 695
    GOTO 630

650 ' Create PARENT Records for Ancestor Search
    CNT = 0: CHBGN = 0: LSET xSTAT$ = " "
    CALL FF1GetRec(PID)
    IF (GL = 1) OR (XLNS = 1) OR ((XLNS = 2) AND (FF1REC.SEX = FGENDR$(1))) OR ((XLNS = 3) AND (FF1REC.SEX = FGENDR$(2))) THEN
       ID = FF1REC.FID
       IF ID > 0 THEN
          NUMCH = 1 + 1024 * (ASC(FF1REC.STATUS) AND 3)
          ALIN# = 2 * XALIN#: GOSUB 655
       END IF
       ID = FF1REC.MID
       IF ID > 0 THEN
          NUMCH = 2 + 1024 * ((ASC(FF1REC.STATUS) \ 4) AND 3)
          ALIN# = 2 * XALIN# + 1: GOSUB 655
       END IF
    END IF
    NUMCH = CNT
    RETURN

655 ' Create Initial NODE Record for Ancestor
    OKAY = (ADOPTOK OR ((NUMCH AND 1024) = 0))
    IF OKAY THEN
       CNT = CNT + 1: NodeHDR.WFNUMREL = NodeHDR.WFNUMREL + 1
       IF FCH = 0 THEN FCH = NodeHDR.WFNUMREL + 1
       IF GLFCH(GL) = 0 THEN GLFCH(GL) = FCH
       GOSUB 660
    END IF
    RETURN

660 ' Create Initial NODE Record for Relative
    GLCOL = GLCOL + 1
    IF GLCOL > NodeHDR.WFMAXCOL THEN NodeHDR.WFMAXCOL = GLCOL
    NODEREC.WFPREC = PREC: NODEREC.WFCHNUM = NUMCH
    NODEREC.WFRID = ID
    NODEREC.WFCOL = GLCOL: NODEREC.WFNUMCH = 0
    IF GL = NodeHDR.WFMAXGL THEN
       NODEREC.WFCHBGN = 0
       ELSE
       NODEREC.WFCHBGN = CHBGN
    END IF
    CALL NodeWorkPUT("NODEREC", NodeHDR.WFNUMREL + 1)
    IF NodePTR(ID) > 0 THEN
       CALL NodeWorkGET("NODEREC", NodePTR(ID) + 1)
       HPTR = NODEREC.WFCHBGN: HCNT = NODEREC.WFNUMCH
       CALL NodeWorkGET("NODEREC", NodeHDR.WFNUMREL + 1)
       NODEREC.WFCHBGN = -NodePTR(ID)
       NODEREC.WFNUMCH = -(HCNT + 1) * (HPTR < 0)
       CALL NodeWorkPUT("NODEREC", NodeHDR.WFNUMREL + 1)
    END IF
    NodePTR(ID) = NodeHDR.WFNUMREL + 1: NodeHDR.WFHIGL = GL
    A$ = INKEY$: IF A$ <> "" THEN A = ASC(A$): CALL ShowINFO
    RETURN

670 ' Determine Alignment Type
    IF FAMPATH AND NOSIBLING AND ((RGL1 = 0) OR (RGL2 = 0)) THEN
       ALIGN = 1
       ELSE
       IF ANCSEARCH THEN
          CALL PUTMSG("HBOXM006")
          ELSE
          CALL PUTMSG("HBOXM007")
       END IF
       IF A$ <> CHR$(27) THEN
          ALIGN = VAL(A$) + 1
          ALIGN = ALIGN + (ALIGN > 1) * ANCSEARCH * 4
          IF (ALIGN < 1) OR (ALIGN > 8) THEN ALIGN = 1
       END IF
    END IF
    RETURN

680 ' Update Header Records
    NodeHDR.WFVER = CHR$(0)
    CALL NodeWorkPUT("HDR", 1)
    x = 1
    FOR GL = 0 TO NodeHDR.WFHIGL
        NodeGLREC.WFGLNUM = GL
        NodeGLREC.WFGLBGN = x: NodeGLREC.WFGLEND = GLLCH(GL)
        NodeGLREC.WFLOCOL = 1: NodeGLREC.WFHICOL = GLCOL(GL)
        x = x + GLCOL(GL)
        CALL NodeWorkPUT("GLREC", GL)
    NEXT GL: CALL ShowINFO
    IF A = 27 THEN 695
    ERASE GLFCH, GLLCH, GLCOL, NodePTR
    GOTO 700

695 ' Premature End to NODE Work Create
    CALL FamCLOSE
    EXIT SUB

700 ' Determine Placement of BOXes
    IF ALIGN = 0 THEN GOSUB 670: IF A$ = CHR$(27) THEN 799
    NodeHDR.WFOPT = CHR$((ASC(NodeHDR.WFOPT) AND 240) + 5)
    CALL NodeWorkPUT("HDR", 1)
    WALIGN = 5: CALL ShowINFO
    HIGL = NodeHDR.WFHIGL
    CALL PUTMSG("HBOXM008")
    REDIM GLFCH(HIGL + 1), GLLCH(HIGL + 1)
    REDIM GLLOCOL(HIGL + 1), GLHICOL(HIGL + 1)
    REDIM GLREC(HIGL + 1), GLNUMCH(HIGL + 1)
    GL = 0: GLREC(0) = 1: GLFCH(0) = 1: GLLCH(0) = 1: NodeHDR.WFMAXCOL = 0
    FOR x = 0 TO HIGL: GLHICOL(x) = -1: NEXT x

710 ' First Placement for NODE Rec at GL Level
    GLREC = GLREC(GL)
    CALL NodeWorkGET("NODEREC", GLREC)
    GLCOL = GLHICOL(GL) + 2
    IF GLCOL > NodeHDR.WFMAXCOL THEN NodeHDR.WFMAXCOL = GLCOL
    GLHICOL(GL) = GLCOL
    FCH = NODEREC.WFCHBGN: NUMCH = NODEREC.WFNUMCH
    IF GLNUMCH(GL) = 1 THEN NODEREC.WFCHNUM = NODEREC.WFCHNUM OR 256
    NODEREC.WFCOL = GLCOL
    CALL NodeWorkPUT("NODEREC", GLREC)
    A$ = INKEY$: IF A$ <> "" THEN IF A$ = CHR$(27) THEN 799 ELSE CALL ShowINFO

720 ' Move from GL to GL+1 (Process Children)
    IF FCH < 1 THEN 730
    GL = GL + 1: GLREC(GL) = FCH
    GLFCH(GL) = GLHICOL(GL) + 2: GLLCH(GL) = FCH + NUMCH - 1
    GLNUMCH(GL) = 1
    A$ = INKEY$: IF A$ <> "" THEN CALL ShowINFO: IF A$ = CHR$(27) THEN 799
    GOTO 710

730 ' Process next Sibling at Level GL
    GLREC = GLREC(GL): IF GLREC = GLLCH(GL) THEN 740
    GLREC(GL) = GLREC + 1: GLNUMCH(GL) = GLNUMCH(GL) + 1
    A$ = INKEY$: IF A$ <> "" THEN CALL ShowINFO: IF A$ = CHR$(27) THEN 799
    GOTO 710

740 ' Move Back to Prior GL (After Processing Children)
    IF GL = 0 THEN GOTO 790
    FCH = GLFCH(GL): LCH = GLHICOL(GL)
    CALL NodeWorkGET("NODEREC", GLREC)
    NODEREC.WFCHNUM = NODEREC.WFCHNUM OR 512
    CALL NodeWorkPUT("NODEREC", GLREC)
    GL = GL - 1: GLCOL = GLHICOL(GL): GLREC = GLREC(GL)
745 ON ALIGN GOTO 746, 747, 748, 749, 746, 747, 748, 749
746 SHIFT = ((GLCOL - FCH) * (GLCOL < FCH) + (GLCOL - LCH) * (GLCOL > LCH)): GOTO 750
747 SHIFT = (FCH - GLCOL): GOTO 750
748 SHIFT = (((LCH + FCH) / 2) - GLCOL): GOTO 750
749 SHIFT = (LCH - GLCOL)

750 ' Check for Required Adjustment
    IF SHIFT > 0 THEN
       GOSUB 760
       ELSE
       IF SHIFT < 0 THEN SHIFT = -SHIFT: BGL = GL: GOSUB 770
    END IF
    A$ = INKEY$: IF A$ <> "" THEN CALL ShowINFO: IF A$ = CHR$(27) THEN 799
    GOTO 730

760 ' Shift Parent to Right
    GLCOL = GLCOL + SHIFT
    GLHICOL(GL) = GLCOL
    IF GLCOL > NodeHDR.WFMAXCOL THEN NodeHDR.WFMAXCOL = GLCOL
    IF GLNUMCH(GL) = 1 THEN GLFCH(GL) = GLCOL
    CALL NodeWorkGET("NODEREC", GLREC)
    NODEREC.WFCOL = GLCOL
    CALL NodeWorkPUT("NODEREC", GLREC)
    RETURN

770 ' Shift Children to Right
    CALL NodeWorkGET("NODEREC", GLREC)
    FCH = NODEREC.WFCHBGN: NUMCH = NODEREC.WFNUMCH: IF FCH < 1 THEN 775
    GL = GL + 1: GLREC = FCH: GLREC(GL) = FCH: GLLCH(GL) = FCH + NUMCH - 1
    CALL NodeWorkGET("NODEREC", GLREC)
    GLCOL = NODEREC.WFCOL + SHIFT: NODEREC.WFCOL = GLCOL
    CALL NodeWorkPUT("NODEREC", GLREC)
    GLFCH(GL) = GLCOL: GLHICOL(GL) = GLCOL: GOTO 770

775 ' Process Next Child This Level
    IF GLREC(GL) = GLLCH(GL) THEN 780
    GLREC = GLREC(GL) + 1: GLREC(GL) = GLREC
    CALL NodeWorkGET("NODEREC", GLREC)
    GLCOL = NODEREC.WFCOL + SHIFT: NODEREC.WFCOL = GLCOL
    CALL NodeWorkPUT("NODEREC", GLREC)
    GLHICOL(GL) = GLCOL: GOTO 770

780 ' No More Children at This Level
    IF GLHICOL(GL) > NodeHDR.WFMAXCOL THEN NodeHDR.WFMAXCOL = GLHICOL(GL)
    GL = GL - 1: IF GL > BGL THEN 775
    RETURN

790 ' Finished Processing Placement of NODEs
    NodeHDR.WFOPT = CHR$((ASC(NodeHDR.WFOPT) AND 240) + ALIGN)
    CALL NodeWorkPUT("HDR", 1)
    CALL PUTMSG("HBOXM009")
    FOR GL = 0 TO HIGL
        CALL NodeWorkGET("GLREC", GL)
        GLLOCOL(GL) = NodeGLREC.WFGLBGN
    NEXT GL
    FOR GL = 0 TO HIGL
        CALL NodeWorkGET("NODEREC", GLLOCOL(GL))
        GLLOCOL(GL) = NODEREC.WFCOL
    NEXT GL
    FOR GL = 0 TO HIGL
        CALL NodeWorkGET("GLREC", GL)
        NodeGLREC.WFLOCOL = GLLOCOL(GL)
        NodeGLREC.WFHICOL = GLHICOL(GL)
        CALL NodeWorkPUT("GLREC", GL)
    NEXT GL
799 CALL FamCLOSE
    EXIT SUB

1900 ' Prepare for FAMILY PATH Search
     REDIM ANCESTOR(FF1HDR.MAXID), CHILD(FF1HDR.MAXID)
     REDIM BL1(101), BL2(101), CHCNT(101)
     BRID = 0: BNAME$ = "": CALL ShowINFO
1904 CALL FmtFIND("HBOXS011"): CALL FmtDREST
1905 CALL MsgFORMAT: IF A = 27 THEN 1959
     CALL FmtFindFLD("ID1"): ID1# = VAL(UTXT$(FFLD.UTXT))
     IF (ID1# < 1) OR (FF1HDR.MAXID < ID1#) THEN CALL ErrBEEP(0): GOTO 1905
     ID1 = ID1#: CALL FF1GetRec(ID1)
     BRID = ID1: BNAME$ = FMTNAME$(1): CALL ShowINFO
1909 CALL FmtFIND("HBOXS012"): CALL FmtDREST
1910 CALL MsgFORMAT: IF A = 27 THEN 1904
     CALL FmtFindFLD("ID2"): ID2# = VAL(UTXT$(FFLD.UTXT))
     IF (ID2# < 1) OR (FF1HDR.MAXID < ID2#) THEN CALL ErrBEEP(0): GOTO 1910
     ID2 = ID2#: CALL FF1GetRec(ID2)
     BRID = ID2: BNAME$ = FMTNAME$(1): CALL ShowINFO
     ANCESTOR(1) = ID1: CHILD(ID1) = ID1: CHILD(0) = -1
     AX = 1: IX = 1: COLOR FG(1), BG(1)
     CALL PUTMSG("HBOXM010")
1925 ID = ANCESTOR(IX): CALL FF1GetRec(ID)
     FID = FF1REC.FID: MID = FF1REC.MID
     IF CHILD(FID) = 0 THEN AX = AX + 1: ANCESTOR(AX) = FID: CHILD(FID) = ID
     IF CHILD(MID) = 0 THEN AX = AX + 1: ANCESTOR(AX) = MID: CHILD(MID) = ID
     IF IX < AX THEN IX = IX + 1: GOTO 1925
1930 RGL1 = 0: RGL2 = 0
     IF CHILD(ID2) > 0 THEN BL2(0) = ID2: AID = ID2: GOTO 1955
     ANCESTOR(1) = ID2: CHILD(ID2) = -ID2
     AX = 1: IX = 1
     CALL PUTMSG("HBOXM011")
1935 ID = ANCESTOR(IX): CALL FF1GetRec(ID)
     FID = FF1REC.FID: MID = FF1REC.MID
     IF CHILD(FID) > 0 THEN BL2(0) = FID: AID = ID: GOTO 1945
     IF CHILD(FID) = 0 THEN AX = AX + 1: ANCESTOR(AX) = FID: CHILD(FID) = -ID
     IF CHILD(MID) > 0 THEN BL2(0) = MID: AID = ID: GOTO 1945
     IF CHILD(MID) = 0 THEN AX = AX + 1: ANCESTOR(AX) = MID: CHILD(MID) = -ID
     IF IX < AX THEN IX = IX + 1: GOTO 1935
     A$ = INKEY$: WHILE A$ <> "": A$ = INKEY$: WEND
     CALL PUTMSG("HBOXM012")
     IF A$ <> CHR$(27) THEN 1909
     BRID = 0: BNAME$ = "": CALL ShowINFO
     A$ = CHR$(27): GOTO 1959
1945 RGL2 = RGL2 + 1: BL2(RGL2) = AID
     IF RGL2 < 100 THEN
        IF CHILD(AID) <> -AID THEN AID = -CHILD(AID): GOTO 1945
     END IF
     AID = BL2(0)
1955 BL1(RGL1) = AID
     IF RGL1 < 100 THEN
        IF CHILD(AID) <> AID THEN AID = CHILD(AID): RGL1 = RGL1 + 1: GOTO 1955
     END IF

     ' Found Common Ancestor
     NodeHDR.WFBRID = BL1(0): CALL FF1GetRec(NodeHDR.WFBRID)
     BRID = NodeHDR.WFBRID: BNAME$ = FMTNAME$(1): CALL ShowINFO
     CALL PUTMSG("HBOXM013")
     IF A$ = CHR$(27) THEN 1909
     FOR GL = RGL1 + 1 TO 101: BL1(GL) = 0: NEXT GL
     FOR GL = RGL2 + 1 TO 101: BL2(GL) = 0: NEXT GL
     NodeHDR.WFMAXGL = RGL1 - (RGL2 - RGL1) * (RGL2 > RGL1)
     A$ = INKEY$: WHILE A$ <> "": A$ = INKEY$: WEND
     CALL PUTMSG("HBOXM014")
     NOSIBLING = (A$ = SNGLKEY$(2))
     IF A$ <> CHR$(27) THEN
        NodeHDR.WFOPT = CHR$(-64 * NOSIBLING + 32 + 5)
        NodeHDR.WFID1 = ID1: NodeHDR.WFID2 = ID2
        NodeHDR.WFRGL1 = CHR$(RGL1): NodeHDR.WFRGL2 = CHR$(RGL2)
     END IF
1959 ERASE ANCESTOR, CHILD
     RETURN

1970 ' Determine First Child ID
     CHBGN = FF1REC.OLDCH
     IF FAMPATH THEN
        IF (BL1(GL) = ID) AND (BL1(GL + 1) <> 0) THEN
           IF NOSIBLING THEN CHBGN = BL1(GL + 1)
           ELSE
           IF (BL2(GL) = ID) AND (BL2(GL + 1) <> 0) THEN
              IF NOSIBLING THEN CHBGN = BL2(GL + 1)
              ELSE
              CHBGN = 0
           END IF
        END IF
        ELSE
        IF (GL > 0) AND (XLNS > 1) THEN
           IF FF1REC.SEX <> FGENDR$(XLNS - 1) THEN CHBGN = 0
        END IF
     END IF
     RETURN

1980 ' Determine Next Child ID
     IF (FAMPATH AND NOSIBLING) THEN
        IF (BL1(GL) = ID) AND (BL1(GL - 1) = BL2(GL - 1)) THEN
           ID = BL2(GL)
           ELSE
           ID = 0
        END IF
        ELSE
        IF FF1REC.FID = PID THEN ID = FF1REC.FCH ELSE ID = FF1REC.MCH
     END IF
     RETURN

END SUB

SUB HBOXChart (CTYPE)
SHARED NodeHDR AS NodeFileHDR, NODEREC AS NodeFileREC, NodeGLREC AS NodeFileGLREC
SHARED SCHAR$, ALIN#, BRID, BNAME$, RTYPE$, WALIGN, ALIGNTAB$()
SHARED BOXWIDTH, ORIENT, ORIENT2

800 : 'Print Ancestor/Descendant/Family Path CHART
    CALL PUTMSG("HBOXM020")
    IF A$ = CHR$(27) THEN 825
    MINCHART = (A$ <> "2"): NFMT$ = SPACE$(52)
    BD$ = SPACE$(12): DD$ = BD$: BPL$ = SPACE$(41): DPL$ = BPL$
    xAGE$ = "": xSTAT$ = " "
    SCLR = 0: PRINTALL = 0
    CENTERED = -1
    NUMED$ = SPACE$(5)
    SELECT CASE CTYPE
       CASE 1: RPT$ = "AHBX": ANCSEARCH = -1
               FCODE$ = RTRIM$(MID$(RRELCD$(3), 3, 3))
               MCODE$ = RTRIM$(MID$(RRELCD$(4), 3, 3))
       CASE 2: RPT$ = "DHBX"
       CASE 3: RPT$ = "FPTH": FAMPATH = -1
    END SELECT
    CALL RptOPEN(RPT$): IF NOT OKAY THEN 825
    CALL BldChartDEF
    IF ANCSEARCH THEN
       CALL RptOptGET("PCOD", x): ShowPCODE = (x = 1)
       CALL RptOptGET("ALIN", x): ShowLIN = (x = 1)
       ELSE
       IF NOT FAMPATH THEN
          CALL RptOptGET("CHNM", x): ShowCHNUM = (x = 1)
       END IF
    END IF
    CALL RptOptGET("HEAD", x): IF FOUND THEN PRTHEAD = (x = 1)
    IF MINCHART THEN
       IF ShowLIN THEN
          DW = LEN(LTRIM$(STR$(2 ^ (NodeHDR.WFHIGL + 1) - 1))) + 1
          ELSE
          DW = 3 - 2 * (NOT ShowPCODE) * ANCSEARCH
       END IF
       SameAS$ = ""
       ELSE
       CALL GetRVAR("UNKNOWN", x): IF FOUND THEN Unknown$ = RVAR$(x)
       CALL RptOptGET("DWTH", DW): IF DW = 0 THEN DW = 15
       CALL GetRVAR("SAMEAS", x): IF FOUND THEN SameAS$ = RVAR$(x) + " "
       CALL RptOptGET("PRID", x): ShowID = (x = 1)
       CALL RptOptGET("SURF", x): LASTFIRST = (x = 1)
       CALL RptOptGET("SURU", x): UPCASE = (x = 1)
       CALL RptOptGET("SPLT", x): SPLITNM = (x = 1)
       CALL RptOptGET("DAT2", x)
       IF OKAY THEN
          DATES = ((x = 1) OR (x = 4)): AllDATES = (x = 4)
          ELSE
          DATES = RPTOPTION.DATES: AllDATES = 0
       END IF
       PLACES = RPTOPTION.PLACES
       SPOUSE = (RPTOPTION.SPOUSE <> 0): ALLSPOUSE = (RPTOPTION.SPOUSE = 2)
       CALL RptOptGET("CMBD", x)
       IF FOUND THEN CombineDATA = (x <> 2): CombineALL = (x = 4)
    END IF
    x = RPTOPTION.LINEREF
    SHOWREF = ((x AND 1) > 0)
    IF SHOWREF THEN
       LEFTSIDE = ((x AND 2) > 0)
       RIGHTSIDE = ((x AND 4) > 0)
       INSIDE = ((x AND 8) > 0)
       OUTSIDE = ((x AND 16) > 0)
       BOTHSIDES = ((x AND 6) = 6)
    END IF
    CALL RptOptGET("SPBX", x): BOXSPAN = (x = 1)
    CALL RptOptGET("BSEP", BOXSEP)
    CALL RptOptGET("STEM", STLTH): ST1LTH = STLTH: ST2LTH = STLTH
    BDATA = BOXSEP + 2: BOXWIDTH = BDATA + DW
    FDL = -(13 * DATES + 44 * PLACES + DATES * PLACES)
    CALL FamOPEN("READ"): IF NOT OKAY THEN 825
    BRID = 0
    CALL NodeWorkOPEN("READQ")
    IF (NOT OKAY) OR (NodeHDR.WFBRID = 0) THEN 825
    SELECT CASE CTYPE
       CASE 1: IF (ASC(NodeHDR.WFOPT) AND 16) = 0 THEN 825
       CASE 2: IF (ASC(NodeHDR.WFOPT) AND 48) > 0 THEN 825
       CASE 3: IF (ASC(NodeHDR.WFOPT) AND 32) = 0 THEN 825
    END SELECT
    BRID = NodeHDR.WFBRID: CALL FF1GetRec(BRID): BNAME$ = FMTNAME$(1)
    SYSVAR.SUBJECT = BNAME$: CALL HTFExpand(ENV.RTITLE, 1)
    WALIGN = ASC(NodeHDR.WFOPT) AND 15
    CALL ShowINFO
    IF (WALIGN < 1) OR (WALIGN = 5) THEN
       CALL PUTMSG("HBOXM021"): GOTO 825
    END IF
    IF (WALIGN = 3) OR (WALIGN = 7) THEN BOXSPAN = -1
    CALL PUTMSG("HBOXM022")
    IF A$ = CHR$(27) THEN 825
    ORIENT = (-(A$ <> "2") - 2 * (A$ = "2"))
    IF NodeHDR.WFMAXCOL = 1 THEN ORIENT2 = 1: GOTO 830
    IF ANCSEARCH THEN
       CALL PUTMSG("HBOXM023")
       ELSE
       CALL PUTMSG("HBOXM024")
    END IF
    IF A$ = CHR$(27) THEN 825
    ORIENT2 = (-(A$ <> "2") - 2 * (A$ = "2"))
    GOTO 830

825 ' Early EXIT
    CALL FamCLOSE: EXIT SUB

830 ' Prepare for Printing (Single/Multi Strip)
    CALL PrintOPEN: IF NOT OKAY THEN 825
    BGNPAGE = PAGE.BGNPAGE
    BLDREF = RPTOPTION.BLDREF
    IF RPTOPTION.BLDREF THEN
       RREFRPT.BGNPAGE = BGNPAGE
       RREFRec.AGL = CHR$(0): RREFRec.DGL = CHR$(0)
    END IF
    PAGE.BWIDTH = PRT.FWIDTH: LW = PRT.FWIDTH
    FIRSTONE = -1: CST = 1: TS = 1: BDLNUM = 1: ALLREF = 0
    LS = LEFTSIDE OR INSIDE: RS = RIGHTSIDE OR OUTSIDE
    LSM1 = PAGE.IMARGIN: RSM1 = PAGE.OMARGIN
    PAGE.IMARGIN = 0: PAGE.OMARGIN = 0
    BW# = BOXWIDTH * ((NodeHDR.WFMAXCOL + 1) \ 2) + FNCW(NodeHDR.WFMAXCOL + 1, 0) * ((NodeHDR.WFMAXCOL + 1) MOD 2)
    TW# = BW# + LSM1 - 8 * (LS + RS)
    PRTSTRIP = (TW# > LW)
    IF NOT PRTSTRIP THEN
       REFAREA = (LW - BW# - LSM1 - RSM1) \ 2
       IF SHOWREF THEN
          IF REFAREA < 5 THEN REFAREA = 5
          IF NOT LS THEN LSM1 = LSM1 + REFAREA
          RM# = LSM1 - REFAREA * LS + BW# - (REFAREA - 5) * RS
          ELSE
          IF REFAREA < 0 THEN REFAREA = 0
          LSM1 = LSM1 + REFAREA
       END IF
       MARGIN = LSM1: PO# = 0: GOTO 900
    END IF
    IF PRT.OPT < 3 THEN 845

840 ' Check for Full Width printing to File
    CALL PUTMSG("XXXXM070")
    IF A$ = CHR$(27) THEN 825
    PRTSTRIP = (VAL(A$) = 1): PAGE.FULLWIDTH = NOT PRTSTRIP
    IF PAGE.FULLWIDTH THEN
       IF SHOWREF THEN
          REFAREA = 8
          RM# = LSM1 - REFAREA * LS + BW# - (REFAREA - 5) * RS
       END IF
       PRT.CFORMS = -1: MARGIN = LSM1: PO# = 0: GOTO 900
    END IF

845 ' Prepare for Strip Printing
    IF BOXSPAN THEN 850
847 ' Adjust for Boxes not Crossing Border of Strip
    REFAREA = -8 * SHOWREF: ALLREF = -1
    BIGMGN = LSM1 + (LSM1 - RSM1) * (RSM1 > LSM1)
    C1 = (LW - BIGMGN + REFAREA * (LS + RS)) \ BOXWIDTH
    IF C1 < 1 THEN
       IF (LSM1 + RSM1) > 0 THEN
          LSM1 = 0: RSM1 = 0: GOTO 847
          ELSE
          IF SHOWREF THEN
             SHOWREF = 0: LS = 0: RS = 0: GOTO 847
             ELSE
             BOXSPAN = -1: GOTO 850
          END IF
       END IF
    END IF
    TS = 1 + (((NodeHDR.WFMAXCOL + 1) \ 2) - 1) \ C1
    IF PRT.SNGLSIDE THEN
       LSM2 = LSM1: RSM2 = RSM1
       ELSE
       LSM2 = RSM1: RSM2 = LSM1
    END IF
    GOTO 855

850 'Compute Number of Strips & Margins if Spanned Boxes
    ALLREF = 0
    TW# = BW# + LSM1 - 5 * (LS + RS)
    TS = FIX((TW# + LW - 1) / LW)
    REFAREA = (TS * LW - BW# - LSM1 - RSM1) \ 2
    IF REFAREA < 5 THEN REFAREA = 5
    IF NOT LS THEN LSM1 = LSM1 + REFAREA: TW# = TW# + REFAREA

855 ' Prepare for First Strip
    CST = 0: GOSUB 980: IF NOT OKAY THEN 825

900 ' Prepare To Print Chart
    IF NOT SHOWREF THEN LEFTSIDE = 0: RIGHTSIDE = 0: INSIDE = 0: OUTSIDE = 0
    PDATA$ = SPACE$(100): DDATA$ = SPACE$(100)
    IF ORIENT = 1 THEN
       VMIN = 0: VMAX = NodeHDR.WFHIGL: VSTEP = 1
       ELSE
       VMIN = NodeHDR.WFHIGL: VMAX = 0: VSTEP = -1
    END IF
    HMIN = 1: HMAX = NodeHDR.WFMAXCOL: HSTEP = 1
    x = HMAX + (HMAX - HMIN) * (HMIN > HMAX)
    REDIM GLLOCOL(NodeHDR.WFHIGL), GLHICOL(NodeHDR.WFHIGL)
    REDIM PCOL(x), CHCNT(x)
    REDIM DIL(x + 1), DID(x), PID(x), DCHNUM(x)
    REDIM DMRN(x), DMID(x), DMIL(x), NDMRN(x)
    IF RPTOPTION.BLDREF THEN REDIM ColRefID(x)
    IF ANCSEARCH THEN REDIM XALIN#(x)
    ' Create BOX Symbol Set Variables
    IF PRT.OPT = 1 THEN
       DCHAR$ = SCHAR$
       ELSE
       DCHAR$ = PDFSymbols.SET2 + LEFT$(PDFSymbols.SET1, 11) + PDFSymbols.SET3
    END IF
    HBC$ = MID$(DCHAR$, 7)
    BOXTOP$ = STRING$(DW + 2, ASC(HBC$)): BOXBOT$ = BOXTOP$
    MID$(BOXTOP$, 1, 1) = MID$(DCHAR$, 1, 1)
    MID$(BOXTOP$, DW + 2, 1) = MID$(DCHAR$, 3, 1)
    MID$(BOXBOT$, 1, 1) = MID$(DCHAR$, 8, 1)
    MID$(BOXBOT$, DW + 2, 1) = MID$(DCHAR$, 10, 1)
    PS = BDATA + DW \ 2
    IF ORIENT = 1 THEN
       PC = 12: CS = PS: CC = 19: FCC = 11: LCC = 13
       ELSE
       PC = 19: CS = PS: CC = 12: FCC = 18: LCC = 10
    END IF
    PC$ = MID$(DCHAR$, PC, 1): CC$ = MID$(DCHAR$, CC, 1)
    FCC$ = MID$(DCHAR$, FCC, 1): LCC$ = MID$(DCHAR$, LCC, 1)
    VBC$ = MID$(DCHAR$, 4, 1): VLC$ = MID$(DCHAR$, 14, 1)
    XC$ = MID$(DCHAR$, 21, 1): HLC$ = MID$(DCHAR$, 17, 1)
    HBC$ = MID$(DCHAR$, 28)
    PBOXTOP$ = STRING$(DW + 2, ASC(HBC$)): PBOXBOT$ = PBOXTOP$
    MID$(PBOXTOP$, 1, 1) = MID$(DCHAR$, 22, 1)
    MID$(PBOXTOP$, DW + 2, 1) = MID$(DCHAR$, 24, 1)
    MID$(PBOXBOT$, 1, 1) = MID$(DCHAR$, 29, 1)
    MID$(PBOXBOT$, DW + 2, 1) = MID$(DCHAR$, 31, 1)
    PVBC$ = MID$(DCHAR$, 25, 1)
    FATHER$ = RTRIM$(MID$(RRELCD$(3), 3, 3))
    MOTHER$ = RTRIM$(MID$(RRELCD$(4), 3, 3))
    FOR GL = 0 TO NodeHDR.WFHIGL
        CALL NodeWorkGET("GLREC", GL)
        GLLOCOL(GL) = NodeGLREC.WFLOCOL
        GLHICOL(GL) = NodeGLREC.WFHICOL
    NEXT GL

930 ' Prepare to Begin Printing
    PAGE.LOFST = 0: PAGE.COUNT = PAGE.BGNPAGE - 1: PAGE.LINE = 0
    PRT.EOF = 0: DLNUM = 0
    FOR x = HMIN TO HMAX STEP HSTEP
        DCHNUM(x) = 0: DIL(x) = 0: DMIL(x) = 0
    NEXT x
    REDIM REFLN(FF1HDR.MAXID), REFCOL(FF1HDR.MAXID)
    IF NOT PRT.CFORMS THEN
       REDIM REFPG(FF1HDR.MAXID)
    END IF
    IF PRTSTRIP THEN
       SYSVAR.NVAR1 = CST: SYSVAR.NVAR2 = TS
       CALL PUTMSG("HBOXM027")
       ELSE
       SYSVAR.MVAR = RTYPE$
       CALL PUTMSG("HBOXM029")
    END IF
    IL = 1: GOSUB 2200: IF NOT OKAY THEN 990
    VX = VMIN: HXMAX = 0: FIRSTONE = 0


939 ' Prepare WREC, WCOL, PCOL, CHCNT
    CALL NodeWorkGET("GLREC", VMIN)
    GLBGN = NodeGLREC.WFGLBGN: GLEND = NodeGLREC.WFGLEND
    IF ORIENT2 = 1 THEN
       HXMAX = FNCOL(NodeGLREC.WFHICOL)
       ELSE
       HXMAX = FNCOL(NodeGLREC.WFLOCOL)
    END IF
    FOR x = GLBGN TO GLEND
        CALL NodeWorkGET("NODEREC", x)
        COL = FNCOL(NODEREC.WFCOL)
        PCOL(COL) = 0: PID(COL) = 0
        CHCNT(COL) = NODEREC.WFNUMCH: Z = NODEREC.WFCHNUM
        IF (ORIENT2 = 2) AND (Z > 255) AND ((Z AND 768) XOR (768)) THEN Z = Z XOR 768
        DCHNUM(COL) = Z: RID = NODEREC.WFRID
        DID(COL) = RID: DIL(COL) = 1: PREC = NODEREC.WFPREC
        IF ANCSEARCH THEN XALIN#(COL) = ALIN#
        IF PREC > 0 THEN
           CALL NodeWorkGET("NODEREC", PREC)
           PID(COL) = NODEREC.WFRID: PCOL(COL) = FNCOL(NODEREC.WFCOL)
        END IF
        IF SPOUSE THEN
           CALL FF1GetRec(RID): DMIL(COL) = -1
           NDMRN(COL) = FF1REC.SPOUSE
           ELSE
           DMIL(COL) = 0
        END IF
    NEXT x
    HXCNT = GLEND - GLBGN + 1

950 ' Print DATALINEs
    XCNT = HXCNT
955 HX = HMIN: GL = VX: PAGE.LOFST = 0
    WHILE HX > -1: GOSUB 1000: WEND
    GOSUB 1165: IF A = 27 THEN 990
    IF XCNT > 0 THEN 955  'Are there unfinished Boxes?

960 ' Process Space Between Rows
    IF ORIENT = 1 THEN STEM = 1: XMAX = ST1LTH ELSE STEM = 2: XMAX = ST2LTH
    FOR XC = 1 TO XMAX
        HX = HMIN
        WHILE HX > -1
           IL = DIL(HX): GOSUB 1080: GOSUB 1140
           IF A = 27 THEN
              HX = -1
              ELSE
              HX = HX + (1 - (IL <> 0)) * HSTEP
              IF ((HSTEP = 1) AND (HX > HXMAX)) OR ((HSTEP = -1) AND (HX < HXMAX)) THEN HX = -1
           END IF
        WEND
        IF A = 27 THEN x = XMAX:  ELSE GOSUB 1165: IF A = 27 THEN x = XMAX
    NEXT XC
    STEM = 0

970 ' Prepare for Next Row
    IF A = 27 THEN 990
    IF VX = VMAX THEN 990 ELSE VX = VX + VSTEP
    GOSUB 1100
    IF A = 27 THEN 990 ELSE GOTO 950

980 ' Check for Strip to Print
    CST = CST + 1
    IF CST > TS THEN 989
    IF NOT (PRINTALL) THEN
       IF NOT (EVENONLY OR ODDONLY) THEN
          SYSVAR.NVAR1 = CST: SYSVAR.NVAR2 = TS
          CALL PUTMSG("HBOXM026")
          IF A$ = CHR$(27) THEN 989
          IF (A$ = SNGLKEY$(2)) THEN 980
          PRINTALL = (A$ = SNGLKEY$(4))
          EVENONLY = (A$ = SNGLKEY$(16)): ODDONLY = (A$ = SNGLKEY$(17))
       END IF
       IF (EVENONLY AND ((CST MOD 2) = 1)) OR (ODDONLY AND ((CST MOD 2) = 0)) THEN 980
    END IF
    PRT.AGAIN = -1
    IF BOXSPAN THEN
       IF CST = 1 THEN
          MARGIN = LSM1: LS = LEFTSIDE OR INSIDE: RS = 0
          ELSE
          MARGIN = 0: LS = 0
          IF CST = TS THEN RS = RIGHTSIDE OR OUTSIDE ELSE RS = 0
       END IF
       PO# = (CST - 1) * LW + (CST > 1) * (LSM1 - REFAREA * (LEFTSIDE OR INSIDE))
       IF RS THEN RM# = TW# - 5 - PO#
       ELSE
       IF (CST MOD 2) = 1 THEN
          MARGIN = LSM1
          LS = LEFTSIDE OR INSIDE: RS = RIGHTSIDE OR OUTSIDE
          ELSE
          MARGIN = LSM2
          LS = LEFTSIDE OR OUTSIDE: RS = RIGHTSIDE OR INSIDE
       END IF
       PO# = (CST - 1) * C1 * BOXWIDTH
       RM# = MARGIN - REFAREA * LS + C1 * BOXWIDTH - (REFAREA - 5) * RS
    END IF
989 CALL PUTMSG("")
    RETURN

990 ' End of Printing CHART
    IF RPTOPTION.BLDREF THEN
       CALL RREFClose: RPTOPTION.BLDREF = 0
    END IF
    CALL PageBREAK: PRT.AGAIN = 0
    IF PRTSTRIP THEN
       GOSUB 980
       IF PRT.AGAIN THEN 930
    END IF
    IF (NOT MINCHART) AND (MXWIDTH > DW) THEN
       SYSVAR.NVAR1 = MXWIDTH: SYSVAR.NVAR2 = DW
       CALL PUTMSG("HBOXM030")
    END IF
    CALL RptCLOSE
    CALL FamCLOSE
    ON LOCAL ERROR GOTO 0
    EXIT SUB

1000 ' Process GL Data
     IL = DIL(HX): RID = DID(HX):  XO = 1: LSET PDATA$ = " "
     IF FAMPATH AND (NOT NOSIBLING) THEN
        MBRPATH = (RID = NodeHDR.WFID1) OR (RID = NodeHDR.WFID2) OR (CHCNT(HX) > 0)
        ELSE
        MBRPATH = 0
     END IF
     IF (IL = -1) THEN 1010
     IF IL = 0 THEN 1010       ' No more Data
     IF IL = 1 THEN GOSUB 1090 ' Print BOX Top
     IF IL = 2 THEN GOSUB 1200 ' Print NAME Line(1)
     IF IL = 3 THEN GOSUB 1200 ' Print NAME Line(2)
     IF IL = 4 THEN GOSUB 1220 ' Print BIRTH Date (After Check for Dup ID)
     IF IL = 5 THEN GOSUB 1240 ' Print BIRTH Place
     IF IL = 6 THEN GOSUB 1250 ' Print DEATH Date
     IF IL = 7 THEN GOSUB 1260 ' Print DEATH Place
     IF IL = 8 THEN GOSUB 1300 ' Print MARRIAGE Data
     IF IL = 9 THEN GOSUB 1095 ' Print BOX Bottom
     IF (IL > 1) AND (IL < 9) THEN GOSUB 1125
1010 IF IL > 10 THEN LSET DDATA$ = " ": GOSUB 1080
     GOSUB 1140
     IF DUPID THEN
        IL = 8: DUPID = 0
        ELSE
        IF ((IL > 0) AND (IL < 8)) OR (IL = 10) THEN IL = IL + 1
     END IF
     IF MINCHART AND (IL > 4) AND (IL < 9) THEN IL = 9
     DIL(HX) = IL
     HX = HX + (1 - (IL <> 0)) * HSTEP
     IF ((HSTEP = 1) AND (HX > HXMAX)) OR ((HSTEP = -1) AND (HX < HXMAX)) THEN HX = -1
     RETURN

1080 ' Insert DSEP$ and BOX Characters
     IF IL = 0 THEN
        MID$(DDATA$, PS, 1) = " "
        ELSE
        IF (ORIENT = 1) THEN
           IF (((STEM = 1) AND (CHCNT(HX) > 0)) OR ((STEM = 2) AND (PCOL(HX) > 0))) THEN
              MID$(DDATA$, PS, 1) = VLC$
           END IF
           ELSE
           IF (((STEM = 2) AND (PCOL(HX) > 0)) OR ((STEM = 1) AND (CHCNT(HX) > 0))) THEN
              MID$(DDATA$, PS, 1) = VLC$
           END IF
        END IF
     END IF
     RETURN

1090 ' Print BOX Top
     IF NOT MBRPATH THEN
        MID$(DDATA$, BDATA - 1, DW + 2) = BOXTOP$
        ELSE
        MID$(DDATA$, BDATA - 1, DW + 2) = PBOXTOP$
     END IF
     IF (PCOL(HX) > 0) AND (ORIENT = 1) THEN
        IF NOT MBRPATH THEN
           MID$(DDATA$, PS, 1) = MID$(DCHAR$, 2, 1)
           ELSE
           MID$(DDATA$, PS, 1) = MID$(DCHAR$, 23, 1)
        END IF
     END IF
     IF (CHCNT(HX) > 0) AND (ORIENT = 2) THEN
        IF NOT MBRPATH THEN
           MID$(DDATA$, PS, 1) = MID$(DCHAR$, 2, 1)
           ELSE
           MID$(DDATA$, PS, 1) = MID$(DCHAR$, 23, 1)
        END IF
     END IF
     RETURN

1095 ' Print BOX Bottom
     IL = 10
     IF NOT MBRPATH THEN
        MID$(DDATA$, BDATA - 1, DW + 2) = BOXBOT$
        ELSE
        MID$(DDATA$, BDATA - 1, DW + 2) = PBOXBOT$
     END IF
     IF (CHCNT(HX) > 0) AND (ORIENT = 1) THEN
        STEM = 1
        IF NOT MBRPATH THEN
           MID$(DDATA$, PS, 1) = MID$(DCHAR$, 9, 1)
           ELSE
           MID$(DDATA$, PS, 1) = MID$(DCHAR$, 30, 1)
        END IF
     END IF
     IF (PCOL(HX) > 0) AND (ORIENT = 2) THEN
        STEM = 2
        IF NOT MBRPATH THEN
           MID$(DDATA$, PS, 1) = MID$(DCHAR$, 9, 1)
           ELSE
           MID$(DDATA$, PS, 1) = MID$(DCHAR$, 30, 1)
        END IF
     END IF
     XCNT = XCNT - 1
     RETURN

1100 ' Increment VX for Horizontal Arrangement
     IF ORIENT = 2 THEN GOSUB 1400: GOSUB 1165: IF A = 27 THEN 1109
     FOR HX = HMIN TO HXMAX STEP HSTEP
         DIL(HX) = 0: PCOL(HX) = 0: CHCNT(HX) = 0: DCHNUM(HX) = 0
     NEXT HX
     CALL NodeWorkGET("GLREC", VX)
     GLBGN = NodeGLREC.WFGLBGN: GLEND = NodeGLREC.WFGLEND
     IF ORIENT2 = 1 THEN
        HXMAX = FNCOL(NodeGLREC.WFHICOL)
        ELSE
        HXMAX = FNCOL(NodeGLREC.WFLOCOL)
     END IF
     FOR x = GLBGN TO GLEND
         CALL NodeWorkGET("NODEREC", x)
         XCOL = NODEREC.WFCOL: COL = FNCOL(XCOL)
         PCOL(COL) = 0
         CHCNT(COL) = NODEREC.WFNUMCH: Z = NODEREC.WFCHNUM
         IF (ORIENT2 = 2) AND (Z > 255) AND ((Z AND 768) XOR (768)) THEN Z = Z XOR 768
         DCHNUM(COL) = Z: RID = NODEREC.WFRID: DID(COL) = RID
         DIL(COL) = 1: IF ANCSEARCH THEN XALIN#(COL) = ALIN#
         IF SPOUSE THEN
            CALL FF1GetRec(RID)
            DMIL(COL) = -1: NDMRN(COL) = FF1REC.SPOUSE
            ELSE
            DMIL(COL) = 0
         END IF
         PREC = NODEREC.WFPREC
         IF PREC > 0 THEN
            CALL NodeWorkGET("NODEREC", PREC)
            PID(COL) = NODEREC.WFRID
            PCOL(COL) = FNCOL(NODEREC.WFCOL)
            ELSE
            PID(COL) = 0: PCOL(COL) = 0
         END IF
     NEXT x
     HXCNT = GLEND - GLBGN + 1
     IF ORIENT = 1 THEN GOSUB 1400: GOSUB 1165: IF A = 27 THEN 1109
     IF ORIENT = 1 THEN STEM = 2: XMAX = ST2LTH ELSE STEM = 1: XMAX = ST1LTH
     FOR XC = 1 TO XMAX
         HX = HMIN
         WHILE HX > -1
            IL = DIL(HX): GOSUB 1080: GOSUB 1140
            IF A = 27 THEN
               HX = -1
               ELSE
               HX = HX + (1 - (IL <> 0)) * HSTEP
               IF ((HSTEP = 1) AND (HX > HXMAX)) OR ((HSTEP = -1) AND (HX < HXMAX)) THEN HX = -1
            END IF
         WEND
         IF A = 27 THEN XC = XMAX:  ELSE GOSUB 1165: IF A = 27 THEN XC = XMAX
     NEXT XC: STEM = 0
1109 LSET PDATA$ = " "
     RETURN

1125 ' Place Characters for Side of Box, and Place Data in Box
     IF NOT MBRPATH THEN
        MID$(DDATA$, BDATA - 1, 1) = VBC$
        MID$(DDATA$, BDATA + DW, 1) = VBC$
        ELSE
        MID$(DDATA$, BDATA - 1, 1) = PVBC$
        MID$(DDATA$, BDATA + DW, 1) = PVBC$
     END IF
     XS = BDATA
     IF NOT CENTERED THEN
        XL = DW
        ELSE
        IF FL > 0 THEN
           XL = FL: FL = 0
           ELSE
           'XL = INSTR(PDATA$, "   ") - 1
           XL = LEN(LTRIM$(RTRIM$(PDATA$)))
        END IF
        IF (XL < 1) OR (XL > DW) THEN
           XL = DW
           ELSE
           XS = XS + (DW - XL + 1) \ 2
        END IF
     END IF
     MID$(DDATA$, XS, XL) = PDATA$
     IF (NOT MINCHART) THEN
        XDW = LEN(RTRIM$(PDATA$))
        IF (XDW > MXWIDTH) THEN MXWIDTH = XDW
     END IF
     RETURN

1140 ' Print Column Data in DDATA$
     IF (A = 27) OR ((DLNUM + 1) < BDLNUM) THEN GOTO 1149
     CWIDTH = FNCW(HX, IL)
     IF NOT PRTSTRIP THEN
        X1 = 1: XL = CWIDTH
        ELSE
        X1 = 0
        X0# = ((HX - 1) \ 2) * (BOXWIDTH) + FNCW(HX - 1, 0) * ((HX - 1) MOD 2)
        IF BOXSPAN THEN
           XL = LW + (CST = 1) * (LSM1 - REFAREA * LS)
           ELSE
           XL = C1 * BOXWIDTH
        END IF
        IF (X0# > (PO# + XL)) OR ((X0# + CWIDTH - 1) < PO#) THEN
           IF (HX = 1) OR ((NOT BOXSPAN) AND SHOWREF) THEN 1142
           IF HX = HXMAX THEN 1144
           GOTO 1149
        END IF
        IF X0# <= PO# THEN X1 = PO# - X0# + 1 ELSE X1 = 1
        IF (PO# + XL) >= (X0# + CWIDTH) THEN
           XL = CWIDTH - X1 + 1
           ELSE
           XL = ((PO# + XL) - (X0# + X1 - 1))
        END IF
     END IF
     IF XL = 0 THEN 1149
1142 IF (PAGE.LOFST = 0) AND (MARGIN > 0) THEN
        PAGE.LOFST = MARGIN
        IF PAGE.FULLWIDTH THEN
           PRINT #4, LEFT$(PLINE$, PAGE.LOFST);
           PO# = PAGE.LOFST
        END IF
     END IF
     IF LS AND (PAGE.LOFST = MARGIN) THEN
        IF (((DLNUM + 1) MOD 5) = 0) THEN
           RSET NUMED$ = FNX$(DLNUM + 1)
           MID$(PLINE$, PAGE.LOFST + 1, 5) = RIGHT$(NUMED$, 4) + ">"
        END IF
        PAGE.LOFST = PAGE.LOFST + REFAREA
        IF PAGE.FULLWIDTH THEN
           PRINT #4, MID$(PLINE$, MARGIN + 1, REFAREA);
           PO# = PO# + REFAREA
        END IF
     END IF
     IF X1 > 0 THEN
        IF PAGE.FULLWIDTH THEN
           PRINT #4, LEFT$(DDATA$, CWIDTH);
           PO# = PO# + CWIDTH
           ELSE
           MID$(PLINE$, PAGE.LOFST + 1, XL) = MID$(DDATA$, X1, XL)
           PAGE.LOFST = PAGE.LOFST + XL
        END IF
     END IF
1144 IF RS AND (((DLNUM + 1) MOD 5) = 0) AND (HX = HXMAX) THEN
        IF PAGE.FULLWIDTH THEN
           x = RM# - PO#
           IF x > 0 THEN
              PRINT #4, SPACE$(x) + "<" + FNX$(DLNUM + 1);
           END IF
           ELSE
           PAGE.LOFST = RM#
           MID$(PLINE$, PAGE.LOFST + 1, 255) = "<" + FNX$(DLNUM + 1)
           PAGE.LOFST = PAGE.LOFST + 5
        END IF
     END IF
1149 LSET DDATA$ = " "
     RETURN

1165 ' Advance One Line Forward
     IF (A <> 27) THEN
        IF (PAGE.LINE = 0) AND (NOT PAGETOP) THEN GOSUB 2200
        DLNUM = DLNUM + 1
        IF PAGE.FULLWIDTH THEN
           PRINT #4, ""
           PAGE.LINE = PAGE.LINE + 1:  LSET PLINE$ = ""
           IF INKEY$ <> "" THEN CALL HoldPRINT
           PAGE.RLINE = DLNUM
           ELSE
           CALL WriteLINE: BLEFT = BLEFT - 1
        END IF
        PAGE.LOFST = 0
     END IF
     IF RPTOPTION.BLDREF THEN GOSUB 1170
     RETURN

1170 ' Create RREF Entries
     IF CTYPE = 1 THEN AGL = GL: DGL = 0:  ELSE AGL = 0: DGL = GL
     RPAGE = PAGE.RPAGE: RLINE = PAGE.RLINE
     FOR HX = HMIN TO HXMAX STEP HSTEP
         RID = ColRefID(HX)
         IF (RID <> 0) THEN
            PAGE.XID = ABS(RID)
            IF RID > 0 THEN
               PAGE.XTYPE = 1
               RREFRec.AGL = CHR$(AGL): RREFRec.DGL = CHR$(DGL)
               ELSE
               PAGE.XTYPE = 3
            END IF
            RREFXData$ = FNX$(RPAGE) + "-" + FNX$(RLINE) + ":" + FNX$(REFCOL(PAGE.XID))
            PAGE.RPAGE = RPAGE: PAGE.RLINE = 0
            CALL RRefADD: CALL RRefBLD
            REFLN(PAGE.XID) = RLINE: IF NOT PRT.CFORMS THEN REFPG(PAGE.XID) = RPAGE
         END IF
         ColRefID(HX) = 0
     NEXT HX
     PAGE.RPAGE = 0
     RETURN

1200 ' Print Relative NAME Line
     LSET xSTAT$ = " "
     IF RID > 0 THEN
        CALL FF1GetRec(RID)
        ELSE
        LSET FF1REC.GIVEN = "(" + Unknown$ + ")": LSET FF1REC.SURNM = ""
     END IF
     IF NOT ((IL <> 2) OR (CSID > 0) OR (GL = 0)) THEN
        CN = DCHNUM(HX) AND 255
        IF NOT ANCSEARCH THEN
           IF ShowCHNUM THEN MID$(PDATA$, XO, 3) = FNX$(CN)
           IF PID(HX) = FF1REC.FID THEN LSET xSTAT$ = FDSTAT$("FID") ELSE LSET xSTAT$ = FDSTAT$("MID")
           ELSE
           IF ShowLIN THEN
              MID$(PDATA$, XO, 255) = LTRIM$(STR$(XALIN#(HX)))
              ELSE
              IF ShowPCODE THEN
                 'IF MinChart THEN XO = XO + 1
                 IF CN = 1 THEN
                    MID$(PDATA$, XO, 1) = FCODE$
                    ELSE
                    MID$(PDATA$, XO, 1) = MCODE$
                 END IF
              END IF
           END IF
           IF (DCHNUM(HX) AND 1024) > 0 THEN LSET xSTAT$ = "*"
           IF (DCHNUM(HX) AND 2048) > 0 THEN LSET xSTAT$ = "?"
        END IF
        IF xSTAT$ <> " " THEN
           XO = INSTR(XO, PDATA$, " ")
           MID$(PDATA$, XO, 1) = xSTAT$
        END IF
        XO = INSTR(XO, PDATA$, " "): IF XO = 0 THEN 1209
        IF NOT MINCHART THEN
           IF ShowCHNUM OR ShowPCODE THEN
              MID$(PDATA$, XO, 1) = ".": XO = XO + 1
           END IF
        END IF
     END IF
     IF MINCHART THEN IL = IL - (IL = 2): GOTO 1209
     IF (IL = 2) AND ShowID THEN
        MID$(PDATA$, XO, 1) = "("
        MID$(PDATA$, XO + 1, 5) = FNX$(RID)
        XO = INSTR(XO, PDATA$, " ")
        IF (NOT ANCSEARCH) AND (CSID = 0) AND (PID(HX) <> 0) AND SPOUSE THEN
           LSET xSTAT$ = " "
           IF PID(HX) = FF1REC.FID THEN
              PID = FF1REC.MID: IF PID <> 0 THEN LSET xSTAT$ = FDSTAT$("MID")
              ELSE
              PID = FF1REC.FID: IF PID <> 0 THEN LSET xSTAT$ = FDSTAT$("FID")
           END IF
           IF PID > 0 THEN
              MID$(PDATA$, XO, 1) = ":"
              MID$(PDATA$, XO + 1, 5) = FNX$(PID)
              IF xSTAT$ <> " " THEN
                 XO = INSTR(XO, PDATA$, " ")
                 MID$(PDATA$, XO, 1) = xSTAT$
              END IF
              XO = INSTR(XO, PDATA$, " ")
           END IF
        END IF
        MID$(PDATA$, XO, 1) = ")": XO = XO + 1
     END IF
     IF NOT SPLITNM THEN
        IF (IL = 2) THEN
           LSET NFMT$ = FMTNAME$(0)
           ELSE
           IL = IL + 1: GOTO 1209
        END IF
        ELSE
        IF (IL = 2) THEN IF LASTFIRST THEN LSET NFMT$ = FF1REC.SURNM ELSE LSET NFMT$ = FF1REC.GIVEN
        IF (IL = 3) THEN IF LASTFIRST THEN LSET NFMT$ = FF1REC.GIVEN ELSE LSET NFMT$ = FF1REC.SURNM
        IF UPCASE THEN
           IF (((IL = 2) AND LASTFIRST) OR ((IL = 3) AND (NOT LASTFIRST))) THEN
              LSET NFMT$ = UCX$(NFMT$)
           END IF
        END IF
     END IF
     IF LEN(RTRIM$(NFMT$)) > 0 THEN
        MID$(PDATA$, XO, 255) = NFMT$: XO = DW
     END IF
1209 RETURN

1220 ' Check for Crossed Lines & Create REF Record
     IF (RID = 0) THEN 1230
     DUPID = (REFLN(RID) <> 0)
     IF DUPID THEN
        DUPCOL = REFCOL(RID): DUPLN = REFLN(RID)
        IF MINCHART AND (DW < 7) THEN
           MID$(PDATA$, XO, 3) = "<->"
           ELSE
           XDATA$ = FNX$(DUPLN) + ":" + FNX$(DUPCOL)
           IF NOT PRT.CFORMS THEN
              XDATA$ = FNX$(REFPG(RID)) + "-" + XDATA$
           END IF
           MID$(PDATA$, XO, 255) = "<" + SameAS$ + XDATA$ + ">"
           XO = DW
        END IF
     END IF
     REFCOL(RID) = ((HX + 1) \ 2)
     IF PRT.CFORMS THEN
        REFLN(RID) = (DLNUM + SPLITNM)
        ELSE
        REFPG(RID) = PAGE.COUNT - (BLEFT = 0)
        REFLN(RID) = PAGE.LINE + (PAGE.LINE - PAGE.HLCNT - 1) * (BLEFT = 0)
     END IF
     IF RPTOPTION.BLDREF AND (RID <> 0) THEN
        IF CSID > 0 THEN
           ColRefID(HX) = -CSID
           ELSE
           ColRefID(HX) = RID
        END IF
     END IF
     IF NOT DUPID THEN 1230
     RETURN

1230 ' Print BIRTH Date
     IF MINCHART THEN IL = 9: RETURN
     IF (DATES) AND (RID <> 0) THEN
        CALL FF1GetRec(RID)
        CALL GetBDEvents(BD$, BPL$, DD$, DPL$, xAGE$)
        LBD = LEN(RTRIM$(LTRIM$(BD$)))
        IF LBD > 0 THEN
           MID$(PDATA$, XO, 14) = RFLBL$(1)
           X1 = INSTR(XO, PDATA$, "   ")
           IF X1 > XO THEN
              XO = X1 - (X1 > XO) + (MID$(PDATA$, X1, 1) = ".")
           END IF
           MID$(PDATA$, XO, 12) = LTRIM$(BD$)
           XO = INSTR(XO, PDATA$, "   ")
        END IF
        GOSUB 1270
     END IF
     IF XO = 1 THEN IL = IL + 1
     RETURN

1240 ' Print BIRTH Place
     IF (PLACES) AND (RID <> 0) THEN
        CALL FF1GetRec(RID)
        CALL GetBDEvents(BD$, BPL$, DD$, DPL$, xAGE$)
        LBP = LEN(LTRIM$(RTRIM$(BPL$)))
        IF LBP > 0 THEN
           MID$(PDATA$, XO, 3) = RFLBL$(2)
           X1 = INSTR(XO, PDATA$, "   ")
           IF X1 > XO THEN
              XO = X1 - (X1 > XO) + (MID$(PDATA$, X1, 1) = ".")
           END IF
           MID$(PDATA$, XO, 255) = LTRIM$(BPL$)
           XO = INSTR(XO, PDATA$, "   ")
        END IF
        IF (CSID = 0) AND (CombineALL) THEN GOSUB 1270
     END IF
     IF XO = 1 THEN IL = IL + 1
     RETURN

1250 ' Print DEATH Date
     IF (DATES) AND (RID <> 0) THEN
        CALL FF1GetRec(RID)
        CALL GetBDEvents(BD$, BPL$, DD$, DPL$, xAGE$)
        LDD = LEN(LTRIM$(RTRIM$(DD$)))
        IF LDD > 0 THEN
           MID$(PDATA$, XO, 14) = RFLBL$(6)
           X1 = INSTR(XO, PDATA$, "   ")
           IF X1 > XO THEN
              XO = X1 - (X1 > XO) + (MID$(PDATA$, X1, 1) = ".")
           END IF
           MID$(PDATA$, XO, 12) = LTRIM$(DD$)
           XO = INSTR(XO, PDATA$, "   ")
        END IF
        GOSUB 1270
     END IF
     IF XO = 1 THEN IL = IL + 1
     RETURN

1260 ' Print DEATH Place
     IF PLACES AND (RID <> 0) THEN
        CALL FF1GetRec(RID)
        CALL GetBDEvents(BD$, BPL$, DD$, DPL$, xAGE$)
        IF RTRIM$(DPL$) <> "" THEN
           MID$(PDATA$, XO, 3) = RFLBL$(7)
           X1 = INSTR(XO, PDATA$, "   ")
           IF X1 > XO THEN
              XO = X1 - (X1 > XO) + (MID$(PDATA$, X1, 1) = ".")
           END IF
           MID$(PDATA$, XO, 255) = LTRIM$(DPL$)
           XO = INSTR(XO, PDATA$, "   ")
        END IF
     END IF
     IF XO = 1 THEN IL = IL + 1
     RETURN

1270 ' Check for Room for More Data
     IF NOT CombineDATA THEN RETURN
     xIL = IL: nxtDL = 0: xHIL = 6 - 2 * (CSID <> 0)
     LBP = 0: LDD = 0: LDP = 0: LDD = 0: LDP = 0
1275 SELECT CASE xIL
        CASE 4: IF PLACES THEN LBP = LEN(RTRIM$(LTRIM$(BPL$)))
                IF LBP > 0 THEN nxtDL = LEN(RFLBL$(2)) + LBP + 1
        CASE 5: IF (DATES AND (CSID = 0)) THEN LDD = LEN(RTRIM$(LTRIM$(DD$)))
                IF LDD > 0 THEN nxtDL = LEN(RFLBL$(6)) + LDD + 1
        CASE 6: IF (PLACES AND ((IL > 4) OR (CSID = 0))) THEN LDP = LEN(RTRIM$(LTRIM$(DPL$)))
                IF LDP > 0 THEN nxtDL = LEN(RFLBL$(7)) + LDP + 1
        CASE 7: IF DATES AND (IL > 5) THEN LDD = LEN(RTRIM$(LTRIM$(DD$)))
                IF LDD > 0 THEN nxtDL = LEN(RFLBL$(6)) + LDD + 1
        CASE 8: IF PLACES AND ((IL > 5) OR CombineALL) THEN LDP = LEN(RTRIM$(LTRIM$(DPL$)))
                IF LDP > 0 THEN nxtDL = LEN(RFLBL$(7)) + LDP + 1
     END SELECT
     IF (nxtDL = 0) AND (xIL < xHIL) THEN
        xIL = xIL + 1
        GOTO 1275
     END IF
     IF (nxtDL > 0) AND ((DW - XO) > nxtDL) THEN
        IL = xIL + 1: IF XO > 1 THEN XO = XO + 1
     END IF
     RETURN


1300 ' Process MARRIAGE Data
     CSID = DMID(HX): RID = CSID - (CSID < 0)
     DIL(HX) = IL: IL = DMIL(HX): MRN = DMRN(HX)
1305 IF IL = -1 THEN GOSUB 1320' Get Next Marriage Record
     DUPID = (IL < 0): IL = ABS(IL)
     IF IL = 0 THEN DIL(HX) = DIL(HX) + 1: GOTO 1310' No more Data
     IF IL = 1 THEN GOSUB 1330' Print Separator Line
     IF IL = 2 THEN GOSUB 1360' Print NAME Line(1)
     IF IL = 3 THEN GOSUB 1360' Print NAME Line(2)
     IF IL = 4 THEN GOSUB 1220' Print BIRTH Date after create REF entry
     IF IL = 5 THEN GOSUB 1240' Print BIRTH Place
     IF IL = 6 THEN GOSUB 1340' Print MARRIAGE Date
     IF IL = 7 THEN GOSUB 1350' Print MARRIAGE Place
     IF IL = 8 THEN GOSUB 1345' Print End of MARRIAGE Date
     IF IL = 9 THEN GOSUB 1355' Print End of MARRIAGE Place
     IF IL = 10 THEN GOSUB 1250' Print DEATH Date
     IF IL = 11 THEN GOSUB 1260' Print DEATH Place
     IL = IL + 1: IF (IL > 11) THEN IL = -1: DUPID = 0
     IF DUPID THEN
        IF IL < 6 THEN IL = -6 ELSE IF IL > 9 THEN IL = -1 ELSE IL = -IL
     END IF
     IF XO = 1 THEN 1305
1310 DMIL(HX) = IL: IL = DIL(HX): CSID = 0
     RETURN

1320 ' Get Next MARRIAGE Record
     MRN = NDMRN(HX): DMRN(HX) = MRN
     IF MRN = 0 THEN IL = 0: RETURN
     CALL FF3GetRec(MRN)
     IF DID(HX) = FF3SPOUSE.SP1ID THEN
        CSID = FF3SPOUSE.SP2ID: NDMRN(HX) = FF3SPOUSE.SP1NXT
        ELSE
        CSID = FF3SPOUSE.SP1ID: NDMRN(HX) = FF3SPOUSE.SP2NXT
     END IF
     IF NOT ALLSPOUSE THEN NDMRN(HX) = 0
     CSID = CSID + (CSID = 0): DMID(HX) = CSID
     RID = CSID - (CSID < 0): IL = 1
     RETURN

1330 ' Print Separator Line
1331 LSET PDATA$ = MID$(BOXTOP$, 2, DW): XO = DW
1332 RETURN

1340 ' Print MARRIAGE Date
     IF DATES THEN
        CALL FF3GetRec(MRN): MLOC = FF3SPOUSE.MLOC
        DFMT$ = LTRIM$(RTRIM$(XRPTDate$(FF3SPOUSE.D1MD, FF3SPOUSE.D1Y, 1)))
        LMD = LEN(DFMT$)
        IF LMD > 0 THEN
           XO = XO - (XO > 1)
           MID$(PDATA$, XO, 14) = RFLBL$(4)
           XO = INSTR(XO, PDATA$, "   ")
           IF XO > 1 THEN IF MID$(PDATA$, XO - 1, 1) <> "." THEN XO = XO + 1
           MID$(PDATA$, XO, 12) = DFMT$
           XO = XO + LMD + 1
           IF CombineDATA AND PLACES AND (MLOC <> 0) THEN
              CALL FF3GetRec(MLOC)
              LMP = LEN(RTRIM$(LTRIM$(FF3PLACE.PLACE1)))
              IF ((DW - XO) > (LMP + LEN(RFLBL$(5)))) THEN
                 IL = IL + 1
              END IF
           END IF
        END IF
     END IF
     IF XO = 1 THEN IL = IL + 1
     RETURN

1345 ' Print End of MARRIAGE Date
     IF AllDATES THEN
        CALL FF3GetRec(MRN): MLOC = FF3SPOUSE.MLOC
        DFMT$ = LTRIM$(RTRIM$(XRPTDate$(FF3SPOUSE.D2MD, FF3SPOUSE.D2Y, 1)))
        LMD = LEN(DFMT$)
        IF LMD > 0 THEN
           XO = XO - (XO > 1)
           MID$(PDATA$, XO, 14) = RFLBL$(9)
           XO = INSTR(XO, PDATA$, "   ")
           IF XO > 1 THEN IF MID$(PDATA$, XO - 1, 1) <> "." THEN XO = XO + 1
           MID$(PDATA$, XO, 12) = DFMT$
           XO = XO + LMD + 1
           IF CombineDATA AND PLACES AND (MLOC <> 0) THEN
              CALL FF3GetRec(MLOC)
              LMP = LEN(RTRIM$(LTRIM$(FF3PLACE.PLACE2)))
              IF ((DW - XO) > (LMP + LEN(RFLBL$(5)))) THEN
                 IL = IL + 1
              END IF
           END IF
        END IF
     END IF
     IF XO = 1 THEN IL = IL + 1
     RETURN

1350 ' Print MARRIAGE Place
     IF PLACES THEN
        CALL FF3GetRec(MRN): MLOC = FF3SPOUSE.MLOC
        IF MLOC <> 0 THEN
           LSET MPLACE = ""
           CALL FF3GetRec(MLOC)
           IF LTRIM$(MPLACE) <> "" THEN
              XO = XO - (XO > 1)
              MID$(PDATA$, XO, 3) = RFLBL$(5)
              XO = INSTR(XO, PDATA$, "   ")
              IF XO > 1 THEN IF MID$(PDATA$, XO - 1, 1) <> "." THEN XO = XO + 1
              MID$(PDATA$, XO, 255) = MPLACE
              XO = XO + LEN(RTRIM$(MPLACE)): FL = XO - 1
           END IF
        END IF
     END IF
     IF (XO = 1) THEN IL = IL + 1
     RETURN

1355 ' Print End of MARRIAGE Place
     IF AllDATES AND PLACES THEN
        CALL FF3GetRec(MRN): MLOC = FF3SPOUSE.MLOC
        IF MLOC <> 0 THEN
           LSET DVPLACE = ""
           CALL FF3GetRec(MLOC)
           IF LTRIM$(DVPLACE) <> "" THEN
              XO = XO - (XO > 1)
              MID$(PDATA$, XO, 3) = RFLBL$(10)
              XO = INSTR(XO, PDATA$, "   ")
              IF XO > 1 THEN IF MID$(PDATA$, XO - 1, 1) <> "." THEN XO = XO + 1
              MID$(PDATA$, XO, 255) = DVPLACE
              XO = XO + LEN(RTRIM$(DVPLACE)): FL = XO - 1
           END IF
        END IF
     END IF
     IF (XO = 1) THEN IL = IL + 1
     RETURN

1360 ' Print SPOUSE Name
     IF IL = 2 THEN
        MID$(PDATA$, XO, 2) = RFLBL$(3)
        XO = INSTR(XO, PDATA$, "   ")
        IF XO > 1 THEN IF MID$(PDATA$, XO - 1, 1) <> "." THEN XO = XO + 1
     END IF
     GOSUB 1200
     RETURN

1400 ' Build NODE Line for Horizontal Arrangement
     LSET PDATA$ = STRING$(BOXWIDTH, ASC(HLC$)): PCOL = 0: A$ = " "
     HX = HMIN
     WHILE HX > -1: IL = DIL(HX)
     IF IL < 1 THEN
        IF PCOL = 0 THEN LSET DDATA$ = " ": GOTO 1420 ELSE LSET DDATA$ = PDATA$
        IF (PCOL = HX - 1) THEN
           PPS = PS - FNCW(HX - 1, 0)
           IF PPS > 0 THEN MID$(DDATA$, PPS, 1) = PC$: GOTO 1420
        END IF
        IF (PCOL = HX) THEN IF PS <= FNCW(HX, 0) THEN MID$(DDATA$, PS, 1) = PC$
        GOTO 1420
     END IF
     IF (DCHNUM(HX) AND 256) = 0 THEN 1410 'If not the FIRST child then go on
     IF (DCHNUM(HX) AND 512) > 0 THEN MID$(DDATA$, PS, 1) = VLC$: GOTO 1420
     PCOL = PCOL(HX)
     MID$(DDATA$, PS + 1, BOXWIDTH - PS) = PDATA$
     IF PCOL = HX THEN MID$(DDATA$, PS, 1) = MID$(DCHAR$, 16, 1): GOTO 1420
     IF (PCOL = HX + 1) THEN
        PPS = FNCW(HX, 0) + PS
        IF PPS <= BOXWIDTH THEN MID$(DDATA$, PPS, 1) = PC$
     END IF
     MID$(DDATA$, PS, 1) = FCC$: GOTO 1420
1410 IF (DCHNUM(HX) AND 512) > 0 THEN   ' If not the LAST child then go on
        MID$(DDATA$, 1, PS - 1) = PDATA$: MID$(DDATA$, PS, 1) = LCC$
        IF (PCOL = HX) THEN
           MID$(DDATA$, PS, 1) = MID$(DCHAR$, 15, 1)
           ELSE
           IF (PCOL = HX - 1) THEN
              PPS = PS - FNCW(HX - 1, 0)
              IF PPS > 0 THEN MID$(DDATA$, PPS, 1) = PC$
           END IF
           MID$(DDATA$, PS, 1) = LCC$
        END IF
        PCOL = 0: GOTO 1420
     END IF
     LSET DDATA$ = PDATA$
     IF PCOL = HX THEN MID$(DDATA$, PS, 1) = XC$: GOTO 1420
     IF (PCOL = HX - 1) THEN
        PPS = PS - FNCW(HX - 1, 0)
        IF PPS > 0 THEN MID$(DDATA$, PPS, 1) = PC$
     END IF
     IF (PCOL = HX + 1) THEN
        PPS = FNCW(HX, 0) + PS
        IF PPS <= BOXWIDTH THEN MID$(DDATA$, PPS, 1) = PC$
     END IF
     MID$(DDATA$, PS, 1) = CC$
1420 GOSUB 1140
     HX = HX + (1 - (IL <> 0)) * HSTEP
     IF ((HSTEP = 1) AND (HX > HXMAX)) OR ((HSTEP = -1) AND (HX < HXMAX)) THEN HX = -1
     WEND
     LSET PDATA$ = " "
     RETURN

2200 ' Print Column Heading Line
     IF PRTHEAD AND (NOT PAGETOP) AND (A <> 27) THEN
        CALL PrintHEAD: PAGETOP = -1: BLEFT = PAGE.BLENGTH
        IF A <> 27 THEN
           HL$ = PLINE$: HOFST = PAGE.LOFST: LSET PLINE$ = "": PAGE.LOFST = 0
           HXMAX = NodeHDR.WFMAXCOL
           XB = (BOXSEP + 2) \ 2
           FOR HX = 1 TO HXMAX STEP 2
              RSET NUMED$ = FNX$((HX + 1) \ 2)
              XL = LEN(LTRIM$(NUMED$))
              X1 = XB + (BOXWIDTH - XL) \ 2
              IF X1 < XB THEN X1 = XB: XL = BOXWIDTH - XB
              MID$(DDATA$, 1, BOXWIDTH) = STRING$(BOXWIDTH, ASC("-"))
              MID$(DDATA$, XB, 1) = ":"
              MID$(DDATA$, X1 + 1, XL) = RIGHT$(NUMED$, XL)
              IF (HX = 1) THEN MID$(DDATA$, 1, XB) = SPACE$(XB)
              GOSUB 1140
           NEXT HX
           IF (((NodeHDR.WFMAXCOL + 1) MOD 2) = 1) THEN
              HX = NodeHDR.WFMAXCOL
              XL = FNCW(NodeHDR.WFMAXCOL + 1, 0)
              MID$(DDATA$, 1, XL) = STRING$(XL, ASC("-"))
              MID$(DDATA$, XB, 1) = ":"
              GOSUB 1140
           END IF
           GOSUB 1165: x = BH + 1
           IF A <> 27 THEN
              WHILE (x > 0) AND (A <> 27)
                  CALL WriteLINE
                  x = x - 1
              WEND
           END IF
           BLEFT = BLEFT - BH - 1
           LSET PLINE$ = HL$: PAGE.LOFST = HOFST: PAGETOP = 0
        END IF
     END IF
     OKAY = (A <> 27)
     RETURN
END SUB

SUB HBOXSumry (CTYPE)
SHARED NodeHDR AS NodeFileHDR, NodeGLREC AS NodeFileGLREC, NODEREC AS NodeFileREC
SHARED ALIGNTAB$(), WALIGN, RTYPE$, ALIN#
500 ' Produce Chart Summary Report
    SELECT CASE CTYPE
       CASE 1: RPT$ = "AHBX": ANCSEARCH = -1
       CASE 2: RPT$ = "DHBX"
       CASE 3: RPT$ = "FPTH": FAMPATH = -1
    END SELECT
    CALL RptLOAD(RPT$): IF NOT OKAY THEN 550
    CALL RptOptGET("DWTH", DW): IF DW = 0 THEN DW = 15
    CALL RptOptGET("BSEP", BOXSEP)
    BOXWIDTH = DW + 2 + BOXSEP
    CALL RptOPEN("HSUM"): IF NOT OKAY THEN 550
    CALL GetRVAR("MULTIPLY", x): IF FOUND THEN FF# = VAL(RVAR$(x))
    CALL FamOPEN("READ"): IF NOT OKAY THEN 550
    BRID = 0
    CALL NodeWorkOPEN("READQ")
    IF (NOT OKAY) OR (NodeHDR.WFBRID = 0) THEN 550
    ALIGN = ASC(NodeHDR.WFOPT) AND 15
    SELECT CASE CTYPE
       CASE 1: IF (ASC(NodeHDR.WFOPT) AND 16) = 0 THEN 550
       CASE 2: IF (ASC(NodeHDR.WFOPT) AND 48) > 0 THEN 550
       CASE 3: IF (ASC(NodeHDR.WFOPT) AND 32) = 0 THEN 550
    END SELECT
    ALIGN = ASC(NodeHDR.WFOPT) AND 15
    IF (ALIGN < 1) OR (ALIGN = 5) THEN
       CALL PUTMSG("HBOXM021"): GOTO 550
    END IF
    BRID = NodeHDR.WFBRID: CALL FF1GetRec(BRID)
    BNAME$ = FMTNAME$(1): WALIGN = ALIGN: CALL ShowINFO
    SYSVAR.NVAR1 = BRID: SYSVAR.SUBJECT = BNAME$
    CALL PUTMSG("HBOXM022")
    IF A$ = CHR$(27) THEN 550
    ORIENT = (-(A$ <> "2") - 2 * (A$ = "2"))
    IF NodeHDR.WFMAXCOL = 1 THEN
       ORIENT2 = 1
       ELSE
       IF ANCSEARCH THEN
          CALL PUTMSG("HBOXM023")
          ELSE
          CALL PUTMSG("HBOXM024")
       END IF
       IF A$ = CHR$(27) THEN 550
       ORIENT2 = (-(A$ <> "2") - 2 * (A$ = "2"))
    END IF
    CALL PrintOPEN
    IF NOT OKAY THEN 550
    IF LEN(PLINE$) < 100 THEN PLINE$ = SPACE$(100)

515 ' Print Report
    PRT.SECTION = " "
    CALL PrintHEAD
    CTYPE$ = "CTYPE" + FNX$(CTYPE)
    CALL GetRVAR(CTYPE$, x): IF FOUND THEN CTYPE$ = RVAR$(x)
    IF FAMPATH THEN
       IF NOSIBLING THEN
          CALL GetRVAR("WOUTSIB", x)
          ELSE
          CALL GetRVAR("WITHSIB", x)
       END IF
       IF FOUND THEN CTYPE$ = CTYPE$ + " " + RVAR$(x)
    END IF
    SYSVAR.MVAR = CTYPE$:
    PRT.SECTION = "1": CALL HTFPRINT("T")
    CALL WriteLINE
    IF NOT FAMPATH THEN PRT.SECTION = "2" ELSE PRT.SECTION = "3"
    CALL HTFPRINT("T")
    IF FAMPATH THEN
       CALL FF1GetRec(NodeHDR.WFID1)
       SYSVAR.NVAR1 = NodeHDR.WFID1: SYSVAR.SUBJECT = FMTNAME$(1)
       PRT.SECTION = "4": CALL HTFPRINT("T")
       CALL FF1GetRec(NodeHDR.WFID2)
       SYSVAR.NVAR1 = NodeHDR.WFID2: SYSVAR.SUBJECT = FMTNAME$(1)
       PRT.SECTION = "5": CALL HTFPRINT("T")
    END IF
    CALL WriteLINE
    SYSVAR.NVAR1 = NodeHDR.WFNUMREL: SYSVAR.NVAR2 = NodeHDR.WFHIGL
    SYSVAR.NVAR3 = (NodeHDR.WFMAXCOL + 1) \ 2
    PRT.SECTION = "A": CALL HTFPRINT("T")
    CALL WriteLINE
    CALL GetRVAR("GARANGE1", x)
    IF FOUND THEN
       SYSVAR.MVAR = RVAR$(x)
       PRT.SECTION = "B": CALL HTFPRINT("T")
    END IF
    SYSVAR.MVAR = "": SYSVAR.SUBJECT = ""
    ORIENT$ = "BASELOC" + FNX$(ORIENT): CALL GetRVAR(ORIENT$, x)
    IF FOUND THEN SYSVAR.MVAR = RVAR$(x)
    IF ANCSEARCH THEN
       ORIENT$ = "PARANGE" + FNX$(ORIENT2)
       PRT.SECTION = "C"
       ELSE
       ORIENT$ = "CARANGE" + FNX$(ORIENT2)
       PRT.SECTION = "D"
    END IF
    CALL GetRVAR(ORIENT$, x): IF FOUND THEN SYSVAR.SUBJECT = RVAR$(x)
    CALL HTFPRINT("T")
    ALIGN$ = "ALIGN" + FNX$(ALIGN): CALL GetRVAR(ALIGN$, x)
    IF FOUND THEN
       SYSVAR.MVAR = RVAR$(x)
       PRT.SECTION = "E": CALL HTFPRINT("T")
    END IF
    CALL WriteLINE
    BW# = (NodeHDR.WFMAXCOL + 1) \ 2
    BW# = BW# * BOXWIDTH
    x# = FNCW(NodeHDR.WFMAXCOL + 1, 0)
    x# = x# * ((NodeHDR.WFMAXCOL + 1) MOD 2)
    BW# = BW# + x#
    'BW# = BOXWIDTH * ((NodeHDR.WFMAXCOL + 1) \ 2) + FNCW(NodeHDR.WFMAXCOL + 1, 0) * ((NodeHDR.WFMAXCOL + 1) MOD 2)
    TW# = BW# + PDFForms.IMARGIN + PDFForms.OMARGIN
    TS = FIX((TW# + PDFForms.FWIDTH - 1) / PDFForms.FWIDTH)
    SYSVAR.NVAR1 = DW: SYSVAR.NVAR2 = BOXSEP: SYSVAR.MVAR = LTRIM$(STR$(BW#))
    PRT.SECTION = "F": CALL HTFPRINT("T")
    SYSVAR.MVAR = LTRIM$(STR$(TW# * FF#))
    PRT.SECTION = "G": CALL HTFPRINT("T")
    SYSVAR.NVAR1 = PDFForms.FWIDTH
    SYSVAR.MVAR = "(" + FNX$(PDFForms.IMARGIN) + "," + FNX$(PDFForms.OMARGIN) + ")"
    SYSVAR.NVAR3 = TS
    PRT.SECTION = "H": CALL HTFPRINT("T")
    IF (ALIGN <> 3) AND (ALIGN <> 5) AND (ALIGN <> 7) THEN
       C1 = (PDFForms.FWIDTH - PDFForms.IMARGIN - PDFForms.OMARGIN) \ BOXWIDTH
       TS2 = 1 + (((NodeHDR.WFMAXCOL + 1) \ 2) - 1) \ C1
       IF TS <> TS2 THEN
          SYSVAR.NVAR1 = TS2
          PRT.SECTION = "I": CALL HTFPRINT("T")
       END IF
    END IF
    CALL WriteLINE
    PRT.SECTION = "T": CALL HTFPRINT("T")
    x$ = " "
    SCW = PDFForms.FWIDTH - PDFForms.IMARGIN - PDFForms.OMARGIN - 30
    IF SCW < 20 THEN SCW = 20
    REDIM GLBOX(SCW)
    CD# = (NodeHDR.WFMAXCOL + 1) / (2 * SCW)
    IF CD# < 1 THEN CD# = 1: SCW = (NodeHDR.WFMAXCOL + 1) \ 2
    IF ORIENT = 1 THEN
       BGNGL = 0: ENDGL = NodeHDR.WFHIGL: GLSTEP = 1
       ELSE
       BGNGL = NodeHDR.WFHIGL: ENDGL = 0: GLSTEP = -1
    END IF
    NUMED$ = SPACE$(5)
    FOR GL = BGNGL TO ENDGL STEP GLSTEP
        CALL NodeWorkGET("GLREC", GL)
        RSET NUMED$ = FNX$(NodeGLREC.WFGLNUM)
        MID$(PLINE$, 2, 5) = NUMED$
        x = ((NodeGLREC.WFLOCOL + 1) / 2): Y = (NodeGLREC.WFHICOL + 1) / 2
        IF ORIENT2 = 2 THEN
           Z = x: x = 1 + ((NodeHDR.WFMAXCOL + 1) \ 2) - Y
           Y = 1 + ((NodeHDR.WFMAXCOL + 1) \ 2) - Z
        END IF
        RSET NUMED$ = FNX$(x): MID$(PLINE$, 8, 5) = NUMED$
        RSET NUMED$ = FNX$(Y): MID$(PLINE$, 14, 5) = NUMED$
        RSET NUMED$ = FNX$(NodeGLREC.WFGLEND - NodeGLREC.WFGLBGN + 1)
        MID$(PLINE$, 20, 5) = NUMED$
        XO = 30
        GLBGN = NodeGLREC.WFGLBGN: GLEND = NodeGLREC.WFGLEND: SCHI = 0
        FOR x = GLBGN TO GLEND
            CALL NodeWorkGET("NODEREC", x)
            COL = (NODEREC.WFCOL + 1) / (2 * CD#): IF COL < 1 THEN COL = 1
            IF ORIENT2 = 2 THEN COL = SCW - COL + 1
            IF COL > SCHI THEN SCHI = COL
            IF GLBOX(COL) > -1 THEN
               IF NODEREC.WFCHBGN < 0 THEN GLBOX(COL) = -1 ELSE GLBOX(COL) = 1
            END IF
        NEXT x
        FOR x = 1 TO SCHI
            IF GLBOX(x) = -1 THEN
               LSET x$ = "x"
               ELSE
               IF GLBOX(x) = 1 THEN
                  LSET x$ = "o"
                  ELSE
                  LSET x$ = " "
               END IF
            END IF
            MID$(PLINE$, x + 30, 1) = x$: GLBOX(x) = 0
        NEXT x: CALL PrintLINE
    NEXT GL
    CALL RptCLOSE
    IF PRT.AGAIN THEN GOTO 515
    ERASE GLBOX

550 ' End of Summary Report Option
    CALL FamCLOSE
END SUB

SUB NodeWorkGET (RTP$, RNUM)
SHARED NodeHDR AS NodeFileHDR, NODEREC AS NodeFileREC, NodeGLREC AS NodeFileGLREC
SHARED ALIN#
3100 ANCSEARCH = ((ASC(NodeHDR.WFOPT) AND 16) > 0)
    SELECT CASE RTP$
       CASE "HDR": GET #5, 1, NodeHDR
       CASE "GLREC": RBA& = LEN(NodeHDR)
                     RBA& = RBA& + RNUM * LEN(NodeGLREC) + 1
                     GET #5, RBA&, NodeGLREC
       CASE "NODEREC": RBA& = LEN(NodeHDR) + (NodeHDR.WFMAXGL + 1) * LEN(NodeGLREC)
                       x& = RNUM - 1: x& = x& * (LEN(NODEREC) - 8 * ANCSEARCH)
                       RBA& = RBA& + x& + 1
                       GET #5, RBA&, NODEREC
                       IF ANCSEARCH THEN GET #5, , ALIN#
    END SELECT
END SUB

SUB NodeWorkOPEN (OPT$)
SHARED NodeHDR AS NodeFileHDR, NODEREC AS NodeFileREC, NodeGLREC AS NodeFileGLREC
3000 ' Open Index REFERENCE File for Creation
    QUIET = (RIGHT$(OPT$, 1) = "Q"): x = LEN(RTRIM$(OPT$)) + QUIET
    XOPT$ = MID$(OPT$, 1, x): DSN$ = FT$(8): NODEXData$ = ""
NODEDSN:
    IF NOT QUIET THEN
       CALL FmtFIND("XXXXS017")
       CALL FmtFindFLD("DSNM"): LSET UTXT$(FFLD.UTXT) = DSN$
       CALL MsgFORMAT: IF A = 27 THEN GOTO NODEOpenBAD
       CALL FmtFindFLD("DSNM"): LSET DSN$ = UTXT$(FFLD.UTXT)
    END IF
    DSNAME$ = RTRIM$(WRKFPFX$ + DSN$)
    LSET SYSVAR.FILENAME = DSNAME$
    ON LOCAL ERROR GOTO NONODE
    OPEN DSNAME$ FOR INPUT AS #5
NODECHK:
    CLOSE #5
    ON LOCAL ERROR GOTO NODEOpenERR
    IF XOPT$ = "NEW" THEN
       CALL PUTMSG("XXXXM002")
       IF A$ <> SNGLKEY$(1) THEN GOTO NODEOpenBAD
    END IF
    IF (XOPT$ = "NEW") THEN
       KILL SYSVAR.FILENAME
       GOTO NODENEW
    END IF
    GOTO NODEOLD
NONODE:
    IF (XOPT$ = "NEW") THEN RESUME NODENEW
    IF QUIET THEN RESUME NODEOpenBAD
    RESUME NONODEMSG
NONODEMSG:
    CALL PUTMSG("XXXXM029"): IF A = 27 THEN GOTO NODEOpenBAD
    GOTO NODEDSN

NODEOLD:
    ON LOCAL ERROR GOTO NODEOpenERR
    SELECT CASE XOPT$
       CASE "INIT": OPEN SYSVAR.FILENAME FOR BINARY ACCESS READ AS #5: QUIET = 0
       CASE "UPDT": OPEN SYSVAR.FILENAME FOR BINARY ACCESS READ WRITE AS #5
       CASE "READ": OPEN SYSVAR.FILENAME FOR BINARY ACCESS READ AS #5
       CASE "WRITE": OPEN SYSVAR.FILENAME FOR BINARY ACCESS READ WRITE AS #5
       CASE ELSE: GOTO NODEOpenBAD
    END SELECT
    SYSVAR.MVAR = ".WRK"
    GET #5, 1, NodeHDR
    IF (NodeHDR.FTYPE <> "W") OR (NodeHDR.WFTYPE <> CHR$(6)) OR (NodeHDR.WFVER <> CHR$(0)) THEN
       CLOSE #5: IF QUIET THEN GOTO NODEOpenBAD
       CALL PUTMSG("XXXXM033")
       IF XOPT$ = "INIT" THEN GOTO NODEDSN
       GOTO NODEOpenBAD
    END IF
    IF (NodeHDR.REORGY <> FF1HDR.REORGY) OR (NodeHDR.REORGMD <> FF1HDR.REORGMD) OR (NodeHDR.REORGTM <> FF1HDR.REORGTM) THEN
       CLOSE #5: IF QUIET THEN GOTO NODEOpenBAD
       CALL PUTMSG("XXXXM034")
       IF XOPT$ = "INIT" THEN GOTO NODEDSN
       GOTO NODEOpenBAD
    END IF
NODEOLDEnd:
    OKAY = -1: GOTO NODEOpenEnd

NODENEW:
    ON LOCAL ERROR GOTO NODEOpenERR
    IF XOPT$ = "WRITE" THEN CLOSE #5: KILL SYSVAR.FILENAME
    OPEN SYSVAR.FILENAME FOR BINARY ACCESS READ WRITE AS #5
    DIM NEWHdr AS NodeFileHDR: NodeHDR = NEWHdr
    NodeHDR.FTYPE = "W": NodeHDR.WFTYPE = CHR$(6): NodeHDR.WFVER = CHR$(255)
    NodeHDR.WFOPT = CHR$(0)
    NodeHDR.REORGY = FF1HDR.REORGY
    NodeHDR.REORGMD = FF1HDR.REORGMD
    NodeHDR.REORGTM = FF1HDR.REORGTM
    NodeHDR.WFUPDTY = DATEY(DATE$, 1)
    NodeHDR.WFUPDTMD = DATEMD(DATE$, 1)
    PUT #5, 1, NodeHDR
    OKAY = -1: GOTO NODEOpenEnd

NODEOpenERR:
    'CALL PutMSG("XXXXM004")   ' "Unable to OPEN File [FILENAME]..."
    SYSVAR.NVAR1 = ERR
    CALL PUTMSG("XXXXM030")    ' "Error Opening [FILENAME]...ERR=[NVAR1]..."
    RESUME NODEOpenBAD
NODEOpenBAD: OKAY = 0

NODEOpenEnd:
    IF OKAY THEN LSET FT$(8) = DSN$
END SUB

SUB NodeWorkPUT (RTP$, RNUM)
SHARED NodeHDR AS NodeFileHDR, NODEREC AS NodeFileREC, NodeGLREC AS NodeFileGLREC
SHARED ALIN#
3200 ' Write record to Node Work file
    ANCSEARCH = ((ASC(NodeHDR.WFOPT) AND 16) > 0)
    SELECT CASE RTP$
       CASE "HDR": PUT #5, 1, NodeHDR
       CASE "GLREC": RBA& = LEN(NodeHDR)
                     RBA& = RBA& + RNUM * LEN(NodeGLREC) + 1
                     PUT #5, RBA&, NodeGLREC
       CASE "NODEREC": RBA& = LEN(NodeHDR) + (NodeHDR.WFMAXGL + 1) * LEN(NodeGLREC)
                       x& = RNUM - 1: x& = x& * (LEN(NODEREC) - 8 * ANCSEARCH)
                       RBA& = RBA& + x& + 1
                       PUT #5, RBA&, NODEREC
                       IF ANCSEARCH THEN PUT #5, , ALIN#
    END SELECT
    NodeHDR.WFUPDTY = DATEY(DATE$, 1)
    NodeHDR.WFUPDTMD = DATEMD(DATE$, 1)

END SUB

SUB ShowINFO
SHARED FMT001, BRID, BNAME$, WALIGN, ALIGNTAB$()
SHARED NodeHDR AS NodeFileHDR
3900 ' Display BRID Information
    CURFMT = FMT001: FMT = FMTTAB(CURFMT)
    x = FMT.BGNFLD
    WHILE x > 0
       FFLD = FLDTAB(x)
       IF FFLD.UTXT > 0 THEN
          x$ = UTXT$(FFLD.UTXT)
          SELECT CASE FFLD.NAME
             CASE "BRID": IF BRID > 0 THEN LSET x$ = STR$(BRID) ELSE LSET x$ = ""
             CASE "NAME": IF BRID > 0 THEN LSET x$ = BNAME$ ELSE LSET x$ = ""
             CASE "NMRL": IF BRID > 0 THEN LSET x$ = STR$(NodeHDR.WFNUMREL) ELSE LSET x$ = ""
             CASE "HIGL": IF BRID > 0 THEN LSET x$ = STR$(NodeHDR.WFHIGL) ELSE LSET x$ = ""
             CASE "HICO": IF BRID > 0 THEN LSET x$ = STR$((NodeHDR.WFMAXCOL + 1) \ 2) ELSE LSET x$ = ""
             CASE "ALIG": LSET x$ = RTRIM$(ALIGNTAB$(WALIGN))
          END SELECT
          LSET UTXT$(FFLD.UTXT) = x$
       END IF
       x = FLDCHN(x).FWD
    WEND
    CALL WinDisplayDATA(0)

END SUB

