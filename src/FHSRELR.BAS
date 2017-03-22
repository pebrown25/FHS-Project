DECLARE SUB BldDetDEF ()
DECLARE SUB BldDetTITLE ()
DECLARE SUB BldDupDEF ()
DECLARE SUB BldMLIST ()
DECLARE SUB BldSumDEF ()
DECLARE SUB BldSumTITLE ()
DECLARE SUB GetNextMARRIAGE ()
DECLARE SUB XDETReport (WTYPE%)
DECLARE SUB XDUPReport (WTYPE%)
DECLARE SUB XSUMReport (WTYPE%)
DECLARE FUNCTION BLOODLINE$ (OPT%)
DECLARE FUNCTION DATEDif$ (D1MD%, D1Y%, D2MD%, D2Y%)
DECLARE FUNCTION FDSTAT$ (X$)
DECLARE FUNCTION FMTNAME$ (NFMT%)
DECLARE FUNCTION FFRBA& (RNUM%, RLTH%)
DECLARE FUNCTION FFRNUM% (RBA&, RLTH%)
DECLARE FUNCTION Relation$ (RT$, RSEX$, AGL%, RGL%, RCODE$())
DECLARE FUNCTION SFILL$ (X%, L%)
DECLARE FUNCTION UCX$ (X$)
DECLARE FUNCTION ZFILL$ (X%, L%)
DECLARE FUNCTION XRPTDate$ (MD%, Y%, O%)
DECLARE FUNCTION XSCRNDate$ (D1MD%, D1Y%)
TYPE SUMREC
   TOTAL    AS INTEGER
   MALE     AS INTEGER
   FEMALE   AS INTEGER
   CLINE    AS INTEGER
   DUPID    AS INTEGER
   NOP      AS INTEGER
   BDT      AS INTEGER
   BPL      AS INTEGER
   LOBY     AS INTEGER
   HIBY     AS INTEGER
   MARRIED  AS INTEGER
   MARCNT   AS INTEGER
   MDT      AS INTEGER
   MPL      AS INTEGER
   LOMY     AS INTEGER
   HIMY     AS INTEGER
   DDT      AS INTEGER
   DDTV     AS INTEGER
   DPL      AS INTEGER
   LODY     AS INTEGER
   HIDY     AS INTEGER
END TYPE
1  REM $INCLUDE: 'FHSCOMON.BAS'
   PN$ = " FAMILY HISTORY SYSTEM - FHSRELR - Ancestor/Descendant/Relative Report Program"
   CY$ = " (C) Copyright 1986,1987,1988,1989,1990,1991,1992,1993,1994,1995,1997 by Phillip E. Brown"
   IF CONFIG.BP = 0 THEN RUN "FHSINIT"
   DEF FNX$ (A) = LTRIM$(STR$(A))
   DIM SUMBASE AS SUMREC, SUM AS SUMREC, TOT AS SUMREC
   REDIM PTRTAB&(10), RX(1), PRPAGE(1), PRLINE(1)
   REDIM DFNUM(1), DFPP(1), RFLTH(1), RFLBL$(1)
   MAXMarry = 10: REDIM MLIST(MAXMarry)
   D1Y = VAL(RIGHT$(DATE$, 4)): D1MD = (100 * VAL(LEFT$(DATE$, 2)) + VAL(MID$(DATE$, 4, 2))): CDATE$ = XSCRNDate$(D1MD, D1Y): ROPT$ = SPACE$(31)
   XRELATION$ = SPACE$(12): WTYPE = MO: BRID = 0: READ$ = "READQ"
   CALL WinINIT("LR1")
   CALL WinPREP(1): PGMRTRN$ = "FHSMENU"
33 FUN = 0: CALL RptFREE(SRPT$): CALL RptFREE(RRPT$)
   IF DRPT$ <> "" THEN CALL RptFREE(DRPT$)
40 ' Format Display
   CALL WinCLR
   CALL FmtFIND("RELRS001"): FMT001 = CURFMT
   CALL WinFORMAT(FMT001): WIN.OFMT = FMT001
   SELECT CASE WTYPE
      CASE 1: RFMT$ = "RELRS002": SRPT$ = "ASUM": RRPT$ = "ARPT": DRPT$ = "ADUP"
      CASE 2: RFMT$ = "RELRS003": SRPT$ = "DSUM": RRPT$ = "DRPT": DRPT$ = "DDUP"
      CASE 3: RFMT$ = "RELRS004": SRPT$ = "RSUM": RRPT$ = "RRPT": DRPT$ = ""
   END SELECT
   CALL FmtFIND(RFMT$): FMT00X = CURFMT
   CALL RptLOAD(SRPT$): CALL RptLOAD(RRPT$)
   IF DRPT$ <> "" THEN CALL RptLOAD(DRPT$)
   CALL WinDisplayLABELS(FMT001): CALL WinDisplayLABELS(FMT00X)
45 ' Refresh Option DISPLAY
   CALL WinPREP(2): OptTABLE.CL = 0
   CALL RptOptDISPLAY(RRPT$)
   CALL WinSWITCH(1)
50 ' Display File, Printer SETUPS
   CURFMT = FMT001: FMT = FMTTAB(CURFMT)
   X = FMT.BGNFLD
   WHILE X > 0
      FFLD = FLDTAB(X)
      IF FFLD.UTXT > 0 THEN
         X$ = UTXT$(FFLD.UTXT)
         SELECT CASE FFLD.NAME
            CASE "FSUP": LSET X$ = FDFSetup.NAME
            CASE "FSDS": LSET X$ = FDFSetup.DESC
            CASE "PSUP": LSET X$ = PDFSetup.NAME
            CASE "PSDS": LSET X$ = PDFSetup.DESC
            CASE "PRTR": LSET X$ = PDFSetup.PRINTER
            CASE "FMWD": LSET X$ = STR$(PDFFORMS.FWIDTH)
            CASE "FMLT": LSET X$ = STR$(PDFFORMS.FLENGTH)
         END SELECT
         LSET UTXT$(FFLD.UTXT) = X$
      END IF
      X = FLDCHN(X).FWD
   WEND
100 ' Check Files
    CALL FAMOpen(READ$): READ$ = "READQ"
    IF OKAY THEN CALL RWRKOpen("READQ", WTYPE, 0)
    WFOKAY = OKAY
    IF WFOKAY THEN
       BRID = WFANCREC.ANCID: CALL FF1GetRec(BRID): BNAME$ = FMTNAME(1)
    END IF
    CALL FAMClose
180 ' Display Work File Information
    CURFMT = FMT00X: FMT = FMTTAB(CURFMT)
    IF NOT WFOKAY THEN CALL FmtDREST: GOTO 181
    X = FMT.BGNFLD
    WHILE X > 0
       FFLD = FLDTAB(X)
       IF FFLD.UTXT > 0 THEN
          X$ = UTXT$(FFLD.UTXT)
          SELECT CASE FFLD.NAME
             CASE "WFDT": LSET X$ = XSCRNDate$(WFHDR.UPDTMD, WFHDR.UPDTY)
             CASE "WFSZ": LSET X$ = STR$(WFHDR.TOTREC)
             CASE "RULA": LSET X$ = FNX$(((ASC(WFHDR.WFRULES) AND 12) \ 4) + 1)
             CASE "RULD": LSET X$ = FNX$(((ASC(WFHDR.WFRULES) AND 48) \ 16) + 1)
             CASE "RULO": LSET X$ = FNX$(ASC(WFHDR.WFRULES) AND 3)
             CASE "HIAL": LSET X$ = STR$(ASC(WFHDR.AMAXLV))
             CASE "HIDL": LSET X$ = STR$(ASC(WFHDR.DMAXLV))
             CASE "NUMA": LSET X$ = STR$(WFHDR.NUMANC)
             CASE "NUMR": LSET X$ = STR$(WFHDR.NUMREL)
             CASE "BRID": LSET X$ = STR$(BRID)
             CASE "NAME": LSET X$ = BNAME$
          END SELECT
          LSET UTXT$(FFLD.UTXT) = X$
       END IF
       X = FLDCHN(X).FWD
    WEND
181 CALL WinDisplayDATA(0)
    CURFMT = FMT001: FMT = FMTTAB(CURFMT)
    CALL FmtFindFLD("NMCT")
    LSET UTXT$(FFLD.UTXT) = STR$(FF1HDR.MAXID)
    CALL WinDisplayDATA(FMT001)

200 ' Get Processing Option
    CALL PutMSG("")
    IF FUN$ <> "" THEN CALL OptHILITE(FUN$, FUN$, 1)
    CALL GetKEY
    IF LEN(A$) < 2 THEN SOUND BP, DUR: FUN = 0: GOTO 200:  ELSE FUN = A
    SELECT CASE FUN
       CASE 30: WTYPE = 1: GOTO 33
       CASE 32: WTYPE = 2: GOTO 33
       CASE 19: WTYPE = 3: GOTO 33
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
       CASE "F2": GOTO 400
       CASE "F3": CALL RptLOAD(RRPT$)
                  CALL RWRKBuild(WTYPE, BRID, 0, FMT00X)
                  GOTO 100
       CASE "F4": GOTO 800
       CASE "F6": GOTO 1000
       CASE "F7": WTYPE = WTYPE + 1: IF WTYPE > 3 THEN WTYPE = 1
                  GOTO 33
       CASE "F9": GOTO 250
    END SELECT
    GOTO 200
250 ' Return to FAMMENU.BAS
    CALL FAMClose
    REDIM RFLTH(1), RFLBL$(1)
    CALL PgmPREP(PGMRTRN$)
    IF (NOT OKAY) OR (A = 27) THEN LSET PGMRTRN$ = "FHSMENU": GOTO 40
    CALL RptFREE(SRPT$): CALL RptFREE(RRPT$): CALL RptFREE(DRPT$)
    CHAIN PGMRTRN$

300 ' Change FILE Parameters
    IF A = 94 THEN LSET PGMRTRN$ = "FHSFILE": GOTO 250
    CALL WinSWITCH(2)
    CALL FDFSelect: READ$ = "READ"
    GOTO 45

400 ' Change PRINTER Parameters
    IF A = 95 THEN LSET PGMRTRN$ = "FHSPRTC": GOTO 250
    CALL WinSWITCH(2)
    CALL PDFSelect
    GOTO 45

800 ' Process Request to Update Options
    CALL WinSWITCH(2)
    SELECT CASE WTYPE
       CASE 3: CALL PutMSG("RELRM010")
       CASE ELSE: CALL PutMSG("RELRM005")
    END SELECT
    IF A = 27 THEN 45
    R$ = A$
    SELECT CASE R$
       CASE "1": XRPT$ = RRPT$
       CASE "2": XRPT$ = SRPT$
       CASE "3": XRPT$ = DRPT$
       CASE ELSE: CALL ErrBEEP(0): GOTO 800
    END SELECT
    IF XRPT$ = "" THEN
       CALL ErrBEEP(0): GOTO 800
       ELSE
       CALL RptLOAD(XRPT$)
       IF R$ <> "1" THEN CALL WinCLR: CALL RptOptDISPLAY(XRPT$)
       CALL RptOptUPDATE(XRPT$)
    END IF
    IF R$ <> "1" THEN CALL WinCLR
    GOTO 45

1000 'PROCESS REPORT OPTIONS
     SELECT CASE WTYPE
        CASE 3: CALL PutMSG("RELRM010")
        CASE ELSE: CALL PutMSG("RELRM005")
     END SELECT
     IF A$ = CHR$(27) THEN 200
     IF PGMRTRN$ = "FHSRPTS" THEN
        SELECT CASE A$
           CASE "1": CALL RptLOAD(RRPT$)
           CASE "2": CALL RptLOAD(SRPT$)
           CASE "3": OKAY = (DRPT$ <> ""): IF OKAY THEN CALL RptLOAD(DRPT$)
           CASE ELSE: GOTO 200
        END SELECT
        IF NOT OKAY THEN 200 ELSE GOTO 250
     END IF
     SELECT CASE A$
       CASE CHR$(27): GOTO 200
       CASE "2": CALL XSUMReport(WTYPE)
       CASE "3": IF DRPT$ <> "" THEN CALL XDUPReport(WTYPE) ELSE GOTO 200
       CASE ELSE: CALL XDETReport(WTYPE)
     END SELECT
     GOTO 40

REM $DYNAMIC
SUB BldDetDEF
SHARED DFNUM(), DFPP()
      RFCNT = 24
      REDIM DFNUM(RFCNT), DFPP(RFCNT)
      REDIM RFNUM(RFCNT), RFSEP(RFCNT), RFLTH(RFCNT), RFLBL$(RFCNT)
      CALL RptOptGET("MNGP", MINGAP)
      IF (NOT OKAY) OR (MINGAP = 0) THEN MINGAP = 2
      CALL RptOptGET("MXGP", MAXGAP)
      IF (NOT OKAY) OR (MAXGAP = 0) THEN MAXGAP = 5
      CALL RptOptGET("MINM", MINMLTH)
      CALL RptOptGET("PNML", PLNMLTH)


59500 ' Select Report Fields - Based Upon Report Options
      ' First Select Prefix Fields
      OKAY = 0
      PAGE.DFCNT = 0: IDFX = 0: BFX = 0: MFX = 0: DFX = 0
      RPTOPTION.TWOLINES = 0: PRID2 = 0
      AMAXLV = ASC(WFHDR.AMAXLV): DMAXLV = ASC(WFHDR.DMAXLV)
      IF (PAGE.BWIDTH = 0) THEN
         BWIDTH = PDFFORMS.FWIDTH - PDFFORMS.IMARGIN - PDFFORMS.OMARGIN
         ELSE
         BWIDTH = PAGE.BWIDTH
      END IF
      IF AMAXLV > 0 THEN
         PAGE.DFCNT = PAGE.DFCNT + 1
         LINEAGE# = -(2 ^ (AMAXLV + 1) - 1#)
         DFNUM(PAGE.DFCNT) = 2
         RFLTH(2) = LEN(STR$(VAL(STR$(LINEAGE#))))
      END IF
      IF RPTOPTION.PRGL THEN
         IF (AMAXLV > 0) OR (DMAXLV > 0) THEN
            X = AMAXLV: IF X < DMAXLV THEN X = DMAXLV
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 24
            RFLTH(24) = LEN(RTRIM$(STR$(X)))
         END IF
      END IF
      IF RPTOPTION.Relation THEN
         PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 3
      END IF
      IF (DMAXLV > 0) AND (NOT RPTOPTION.OUTLINE) THEN
         PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 4
         RFLTH(4) = 3 * DMAXLV - 1
      END IF
      IF RPTOPTION.PRID AND (RPTOPTION.FREEFORM OR (RPTOPTION.COMMENTS = 0)) THEN
         IDFX = PAGE.DFCNT + 1: RFSEP(6) = MINGAP + 1
         DFNUM(IDFX) = 6: DFNUM(IDFX + 1) = 7: DFNUM(IDFX + 2) = 8
         PAGE.DFCNT = PAGE.DFCNT + 3
         IF ((RPTOPTION.SPOUSE AND 1) > 0) THEN
            DFNUM(IDFX + 3) = 9: PAGE.DFCNT = PAGE.DFCNT + 1
            ELSE
            RFLTH(9) = 0
         END IF
      END IF

      ' Select Data Fields
      RFSEP(DFNUM(1)) = 1: PAGE.PFCNT = PAGE.DFCNT
      PAGE.DFCNT = PAGE.DFCNT + 1
      DFNUM(PAGE.DFCNT) = 5: RFSEP(5) = MINGAP + 1
      RFLTH(5) = MINMLTH - RPTOPTION.OUTLINE * (3 * DMAXLV) - 20 * RPTOPTION.FREEFORM
      IF RPTOPTION.FREEFORM THEN 59510
      IF RPTOPTION.PRID AND RPTOPTION.COMMENTS THEN
         IDFX = PAGE.DFCNT + 1
         DFNUM(IDFX) = 6: DFNUM(IDFX + 1) = 7: DFNUM(IDFX + 2) = 8
         PAGE.DFCNT = PAGE.DFCNT + 3
         IF ((RPTOPTION.SPOUSE AND 1) > 0) THEN
            DFNUM(IDFX + 3) = 9: PAGE.DFCNT = PAGE.DFCNT + 1
            ELSE
            RFLTH(9) = 0
         END IF
      END IF
      IF RPTOPTION.TIMELINE THEN
         DFNUM(PAGE.DFCNT + 1) = 22: DFNUM(PAGE.DFCNT + 2) = 23
         PAGE.DFCNT = PAGE.DFCNT + 2
         GOTO 59510
      END IF
      IF RPTOPTION.PRSEX THEN PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 10
      IF RPTOPTION.PRAGE THEN PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 11
      IF RPTOPTION.DATES THEN
         PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 12
         BFX = PAGE.DFCNT
      END IF
      IF RPTOPTION.PLACES THEN
         PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 13
      END IF
      IF RPTOPTION.SPOUSE THEN
         IF RPTOPTION.DATES THEN
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 15: MFX = PAGE.DFCNT
         END IF
         IF RPTOPTION.PRAGE THEN
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 18
         END IF
         IF RPTOPTION.PLACES THEN
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 16
         END IF
      END IF
      IF RPTOPTION.DATES THEN
         PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 19: DFX = PAGE.DFCNT
      END IF
      IF RPTOPTION.PLACES THEN PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 20

59510 ' Create Report Definition Tables
      X = 1: RFNAME$ = "06010": GOSUB 59505     'COMMENTS
      X = 2: RFNAME$ = "07011": GOSUB 59505     'LNUM
      X = 3: RFNAME$ = "07004": GOSUB 59505     'RELATION
      X = 4: RFNAME$ = "07012": GOSUB 59505     'BLOODLINE
      X = 5: RFNAME$ = "01020": GOSUB 59505     'FULL NAME
      X = 6: RFNAME$ = "01010": GOSUB 59505     'RID
      X = 7: RFNAME$ = "01011": GOSUB 59505     'FID
      X = 8: RFNAME$ = "01012": GOSUB 59505     'MID
      X = 9: RFNAME$ = "02010": GOSUB 59505     'SPID
      X = 10: RFNAME$ = "01015": GOSUB 59505    'SEX CODE
      X = 11: RFNAME$ = "01016": GOSUB 59505    'AGE
      X = 12: RFNAME$ = "01030": GOSUB 59505    'BIRTH DATE
      X = 13: RFNAME$ = "01032": GOSUB 59505    'BIRTH PLACE
              RFLTH(13) = PLNMLTH
      X = 14: RFNAME$ = "01034": GOSUB 59505    'BIRTH DATE & PLACE
              RFLTH(14) = PLNMLTH
      X = 15: RFNAME$ = "02020": GOSUB 59505    'MARRIAGE DATE
      X = 16: RFNAME$ = "02021": GOSUB 59505    'MARRIAGE PLACE
              RFLTH(16) = PLNMLTH
      X = 17: RFNAME$ = "02024": GOSUB 59505    'MARRIAGE DATE & PLACE
              RFLTH(17) = PLNMLTH
      X = 18: RFNAME$ = "02015": GOSUB 59505    'Last ANNIVERSARY
      X = 19: RFNAME$ = "01040": GOSUB 59505    'DEATH DATE
      X = 20: RFNAME$ = "01042": GOSUB 59505    'DEATH PLACE
              RFLTH(20) = PLNMLTH
      X = 21: RFNAME$ = "01044": GOSUB 59505    'DEATH DATE & PLACE
              RFLTH(21) = PLNMLTH
      X = 22: RFNAME$ = "07020": GOSUB 59505    'LIFE SPAN
      X = 23: RFNAME$ = "07021": GOSUB 59505    'LIFE LINE
      X = 24                                    'AGL, DGL or RGL
      IF AMAXLV > 0 THEN
         IF DMAXLV = 0 THEN
            RFNAME$ = "07001"
            ELSE
            RFNAME$ = "07003"
         END IF
         ELSE
         RFNAME$ = "07002"
      END IF
      GOSUB 59505
      GOTO 59520

59505 ' Process Report Field
      CALL GetRFLD(RFNAME$, FX)
      IF RFSEP(X) = 0 THEN RFSEP(X) = MINGAP
      IF FOUND THEN
         RDFFLD = RFLDS(FX)
         IF RFLTH(X) = 0 THEN RFLTH(X) = ASC(RDFFLD.RLTH)
         ALTH = ASC(RDFFLD.ALTH): LLTH = ASC(RDFFLD.LLTH)
         HLTH = ASC(RDFFLD.HLTH)
         IF (X = 1) OR (RPTOPTION.FREEFORM AND ((X > 9) AND (X < 24))) THEN
            RFLBL$(X) = MID$(RFLD$(FX), ALTH + 1, LLTH)
            ELSE
            RFLBL$(X) = MID$(RFLD$(FX), ALTH + LLTH + 1, HLTH)
         END IF
      END IF
      RETURN

59520 ' Check Report Width
      RWIDTH = 0: FX = 0
      WHILE FX < PAGE.DFCNT
          FX = FX + 1: X = DFNUM(FX): FSEP = RFSEP(X)
          IF (X > 0) AND (FSEP > 0) THEN
             Y = RFLTH(X): Z = LEN(RFLBL$(X)): IF Y < Z THEN Y = Z
             RWIDTH = RWIDTH + FSEP + Y
          END IF
      WEND
      RWIDTH = RWIDTH - 1
      IF RWIDTH <= BWIDTH THEN 59540
      IF RPTOPTION.TIMELINE OR RPTOPTION.FREEFORM THEN 59530
      IF NOT (RPTOPTION.DATES AND RPTOPTION.PLACES) THEN 59530
      IF NOT RPTOPTION.TWOLINES THEN
         RPTOPTION.TWOLINES = -1
         DFNUM(BFX) = 14: DFNUM(BFX + 1) = 0
         DFNUM(DFX) = 21: DFNUM(DFX + 1) = 0
         IF RPTOPTION.SPOUSE THEN
            DFNUM(MFX) = 17: DFNUM(MFX + 1) = 0
            IF RPTOPTION.PRAGE THEN DFNUM(MFX + 2) = 0
         END IF
         GOTO 59520
      END IF
      IF RPTOPTION.PRID AND (NOT PRID2) THEN
         PRID2 = -1
         DFNUM(IDFX) = 0: DFNUM(IDFX + 1) = 0: DFNUM(IDFX + 2) = 0
         IF ((RPTOPTION.SPOUSE AND 1) > 0) THEN DFNUM(IDFX + 3) = 0
         X = RFLTH(6) + RFLTH(7) + RFLTH(8) + RFLTH(9) + 18
         IF X > RFLTH(5) THEN RFLTH(5) = X
         GOTO 59520
      END IF

59530 ' Final Check After Trying TWOLINES and PRID2 Options
      PWIDTH = BWIDTH + PDFFORMS.IMARGIN + PDFFORMS.OMARGIN
      RPWIDTH = RWIDTH + PDFFORMS.IMARGIN + PDFFORMS.OMARGIN
      IF PWIDTH < RPWIDTH THEN
         SYSVAR.NVAR1 = PWIDTH: SYSVAR.NVAR2 = RPWIDTH
         CALL PutMSG("XXXXM009")
         IF A$ <> SNGLKEY$(1) THEN 59560
         BWIDTH = RWIDTH
         IF LEN(PLINE$) < RPWIDTH THEN PLINE$ = SPACE$(RPWIDTH)
      END IF

59540 ' Adjust to Fit
      IF RPTOPTION.FREEFORM THEN 59550
      IF NOT (RWIDTH < BWIDTH) THEN 59550
      IF RPTOPTION.TIMELINE THEN
         RFLTH(23) = RFLTH(23) + (BWIDTH - RWIDTH): GOTO 59550
      END IF
      GAPS = 0: FX = PAGE.PFCNT
      WHILE FX <= PAGE.DFCNT
          X = DFNUM(FX)
          IF (X > 0) AND (RFSEP(X) > 0) AND (X <> 7) AND (X <> 8) THEN
             GAPS = GAPS + 1
          END IF
          FX = FX + 1
      WEND
      GAP = (BWIDTH - RWIDTH) \ GAPS
      GAP1 = (BWIDTH - RWIDTH) - GAPS * GAP
      IF GAP >= (MAXGAP - MINGAP) THEN GAP = MAXGAP - MINGAP: GAP1 = 0
      FX = PAGE.PFCNT
      WHILE FX <= PAGE.DFCNT
          X = DFNUM(FX): FSEP = RFSEP(X)
          IF (X > 0) AND (FSEP > 0) AND (X <> 7) AND (X <> 8) THEN
             RFSEP(X) = FSEP + GAP - (GAP1 > 0): GAP1 = GAP1 - 1
          END IF
          FX = FX + 1
      WEND

59550 ' Determine Field Print Positions
      PP = 0: FX = 0
      WHILE FX < PAGE.DFCNT
          FX = FX + 1: X = DFNUM(FX): FSEP = RFSEP(X)
          IF (X > 0) AND (FSEP > 0) THEN
             FL = RFLTH(X): HL = LEN(RFLBL$(X))
             Z = FL: IF Z < HL THEN Z = HL
             PP = PP + FSEP + (Z + 1) \ 2 - (FL + 1) \ 2
             DFPP(FX) = PP
             PP = PP + Z
          END IF
      WEND

59560 ' Return to Calling Program
      PAGE.BWIDTH = BWIDTH
      IF RPTOPTION.PRID THEN
         IF PRID2 THEN RPTOPTION.PRID = 2 ELSE RPTOPTION.PRID = 1
      END IF
      OKAY = -1
END SUB

REM $STATIC
SUB BldDetTITLE
SHARED DFNUM(), DFPP()
SHARED TMBY, TMEY
     ' Build Title Line
      TL$ = SPACE$(PAGE.BWIDTH)
      FOR FX = 1 TO PAGE.DFCNT: X = DFNUM(FX): PP = DFPP(FX)
         IF (X > 0) AND (PP > 0) THEN
            FL = RFLTH(X): HL = LEN(RFLBL$(X))
            IF X = 23 THEN
               MID$(TL$, PP, 4) = FNX$(TMBY)
               MID$(TL$, PP + FL - 5, 4) = SFILL$(TMEY, 4)
            END IF
            IF HL > 0 THEN
               IF (RPTOPTION.FREEFORM = 0) OR (FX <= PAGE.PFCNT) THEN
                  PP = PP + (FL + 1) \ 2 - (HL + 1) \ 2
                  IF PP > 0 THEN MID$(TL$, PP, HL) = RFLBL$(X)
               END IF
            END IF
         END IF
      NEXT FX
      X = LEN(RTRIM$(TL$))
      IF X > 0 THEN
         HTFLMAX = HTFLMAX + 1
         REDIM PRESERVE HTFLINE$(HTFLMAX)
         HTFLINE$(HTFLMAX) = "T00" + RTRIM$(TL$)
      END IF
      OKAY = -1
END SUB

SUB BldDupDEF
      PAGE.DFCNT = 4
      REDIM DFNUM(PAGE.DFCNT), DFPP(PAGE.DFCNT)
      REDIM RFNUM(PAGE.DFCNT), RFSEP(PAGE.DFCNT), RFLTH(PAGE.DFCNT), RFLBL$(PAGE.DFCNT)
      X = 1: RFNAME$ = "01010": GOSUB DupFLD     'ID
      X = 2: RFNAME$ = "01020": GOSUB DupFLD     'NAME
      X = 3: RFNAME$ = "07004": GOSUB DupFLD     'Relation
      X = 4: RFNAME$ = "07010": GOSUB DupFLD     'LINEAGE
      EXIT SUB

DupFLD: ' Process Report Field
      CALL GetRFLD(RFNAME$, FX)
      IF FOUND THEN
         RDFFLD = RFLDS(FX)
         IF RFLTH(X) = 0 THEN RFLTH(X) = ASC(RDFFLD.RLTH)
         ALTH = ASC(RDFFLD.ALTH): LLTH = ASC(RDFFLD.LLTH)
         RFLBL$(X) = MID$(RFLD$(FX), ALTH + 1, LLTH)
      END IF
      RETURN

END SUB

SUB BldMLIST
SHARED SPRESEQ, SPRNO, SPCNT, MAXMarry, MLIST()
   SPCNT = 0
   IF SPRESEQ THEN
      WHILE SPRNO <> 0
         SPCNT = SPCNT + 1
         IF SPCNT > MAXMarry THEN
            MAXMarry = MAXMarry + 5
            REDIM PRESERVE MLIST(MAXMarry)
         END IF
         MLIST(SPCNT) = SPRNO
         CALL FF3GetRec(SPRNO)
         IF FF1REC.RID = FF3SPOUSE.SP1ID THEN
            SPRNO = FF3SPOUSE.SP1NXT
            ELSE
            IF FF1REC.RID = FF3SPOUSE.SP2ID THEN
               SPRNO = FF3SPOUSE.SP2NXT
               ELSE
               SPRNO = 0
            END IF
         END IF
      WEND
      SPRNO = MLIST(SPCNT)
   END IF

END SUB

SUB BldSumDEF
SHARED DFNUM(), DFPP()
SHARED DATECNT, DATERNG, PLACCNT, MARRCNT, RWIDTH, VRTCL
      AMAXLV = ASC(WFHDR.AMAXLV): DMAXLV = ASC(WFHDR.DMAXLV)
      RELSUMOPT = (AMAXLV > 0) AND (DMAXLV > 0)
      IF RELSUMOPT THEN RFCNT = AMAXLV + 3 ELSE RFCNT = 23
      PAGE.DFCNT = 0
      REDIM DFNUM(RFCNT), DFPP(RFCNT)
      REDIM RFNUM(RFCNT), RFSEP(RFCNT), RFLTH(RFCNT), RFLBL$(RFCNT)
      CALL RptOptGET("MNGP", MINGAP)
      IF (NOT OKAY) OR (MINGAP = 0) THEN MINGAP = 2
      DFPP(0) = MINGAP
      CALL RptOptGET("MXGP", MAXGAP)
      IF (NOT OKAY) OR (MAXGAP = 0) THEN MAXGAP = 5
      IF RELSUMOPT THEN GOTO 58500
      IF DMAXLV = 0 THEN
         CALL RptOptGET("CNOP", X): CNTNOP = (X = 1)
      END IF
      CALL RptOptGET("CDUP", X): CNTDUP = (X = 1)
      CALL RptOptGET("CLIN", X): CNTCLN = (X = 1)
      CALL RptOptGET("DTCT", X): DATECNT = (X = 1)
      CALL RptOptGET("DTRG", X): DATERNG = (X = 1)
      CALL RptOptGET("PLCT", X): PLACCNT = (X = 1)
      CALL RptOptGET("MCNT", X): MARRCNT = (X = 1)

58500 ' Select Report Fields - Based Upon Report Options
      OKAY = 0
      IF (PAGE.BWIDTH = 0) THEN
         BWIDTH = PDFFORMS.FWIDTH - PDFFORMS.IMARGIN - PDFFORMS.OMARGIN
         ELSE
         BWIDTH = PAGE.BWIDTH
      END IF
      IF RELSUMOPT THEN GOTO 58520
      PAGE.DFCNT = 1: DFNUM(1) = 1
      IF RPTOPTION.Relation OR VRTCL THEN
         PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 2
      END IF
      IF CNTCLN THEN PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 4
      IF CNTNOP THEN PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 6
      IF CNTDUP THEN PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 5
      PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 7
      PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 8
      PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 3
      IF DATECNT THEN PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 9
      IF DATERNG THEN
         DFNUM(PAGE.DFCNT + 1) = 11: DFNUM(PAGE.DFCNT + 2) = 12
         PAGE.DFCNT = PAGE.DFCNT + 2
      END IF
      IF PLACCNT THEN PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 10
      IF MARRCNT THEN
         DFNUM(PAGE.DFCNT + 1) = 13: DFNUM(PAGE.DFCNT + 2) = 14
         PAGE.DFCNT = PAGE.DFCNT + 2
         IF DATECNT THEN PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 15
         IF DATERNG THEN
            DFNUM(PAGE.DFCNT + 1) = 17: DFNUM(PAGE.DFCNT + 2) = 18
            PAGE.DFCNT = PAGE.DFCNT + 2
         END IF
         IF PLACCNT THEN PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 16
      END IF
      IF DATECNT THEN
         DFNUM(PAGE.DFCNT + 1) = 19: DFNUM(PAGE.DFCNT + 2) = 20
         PAGE.DFCNT = PAGE.DFCNT + 2
      END IF
      IF DATERNG THEN
         DFNUM(PAGE.DFCNT + 1) = 22: DFNUM(PAGE.DFCNT + 2) = 23
         PAGE.DFCNT = PAGE.DFCNT + 2
      END IF
      IF PLACCNT THEN PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 21

58510 ' Create Report Definition Tables
      X = 1                                     'AGL or DGL
      IF (DMAXLV > 0) AND (AMAXLV = 0) THEN
         RFNAME$ = "07002"
         ELSE
         RFNAME$ = "07001"
      END IF
      GOSUB 58515
      X = 2: RFNAME$ = "07004": GOSUB 58515     'Relation
      X = 3: RFNAME$ = "08001": GOSUB 58515     'Total
      X = 4: RFNAME$ = "08002": GOSUB 58515     'Lines
      X = 5: RFNAME$ = "08003": GOSUB 58515     'Duplicate ID's
      X = 6: RFNAME$ = "08004": GOSUB 58515     'No Parents
      X = 7: RFNAME$ = "08006": GOSUB 58515     'Males
      X = 8: RFNAME$ = "08007": GOSUB 58515     'Females
      X = 9: RFNAME$ = "08015": GOSUB 58515     'BY>0
      X = 10: RFNAME$ = "08016": GOSUB 58515    '#BPL
      X = 11: RFNAME$ = "08017": GOSUB 58515    'LOBY
      X = 12: RFNAME$ = "08018": GOSUB 58515    'HIBY
      X = 13: RFNAME$ = "08020": GOSUB 58515    'MAR
      X = 14: RFNAME$ = "08021": GOSUB 58515    '#MAR
      X = 15: RFNAME$ = "08025": GOSUB 58515    'MY>0
      X = 16: RFNAME$ = "08026": GOSUB 58515    '#MPL
      X = 17: RFNAME$ = "08027": GOSUB 58515    'LOMY
      X = 18: RFNAME$ = "08028": GOSUB 58515    'HIMY
      X = 19: RFNAME$ = "08034": GOSUB 58515    'DY>0
      X = 20: RFNAME$ = "08035": GOSUB 58515    'DY<9
      X = 21: RFNAME$ = "08036": GOSUB 58515    '#DPL
      X = 22: RFNAME$ = "08037": GOSUB 58515    'LODY
      X = 23: RFNAME$ = "08038": GOSUB 58515    'HIDY
      GOTO 58525

58515 ' Process Report Field
      CALL GetRFLD(RFNAME$, FX)
      IF RFSEP(X) = 0 THEN RFSEP(X) = MINGAP
      IF FOUND THEN
         RDFFLD = RFLDS(FX)
         IF RFLTH(X) = 0 THEN RFLTH(X) = ASC(RDFFLD.RLTH)
         ALTH = ASC(RDFFLD.ALTH): LLTH = ASC(RDFFLD.LLTH)
         HLTH = ASC(RDFFLD.HLTH)
         IF VRTCL THEN
            RFLBL$(X) = MID$(RFLD$(FX), ALTH + 1, LLTH)
            ELSE
            RFLBL$(X) = MID$(RFLD$(FX), ALTH + LLTH + 1, HLTH)
         END IF
      END IF
      RETURN

58520 ' Relationship Summary Table
      PAGE.DFCNT = RFCNT
      DFNUM(1) = 1: RFSEP(1) = 1: RFLTH(1) = 8
      CALL GetRFLD("07003", FX)
      IF FOUND THEN
         RDFFLD = RFLDS(FX)
         O = ASC(RDFFLD.ALTH) + ASC(RDFFLD.LLTH): L = ASC(RDFFLD.HLTH)
         RFLBL$(1) = MID$(RFLD$(FX), O + 1, L)
      END IF
      FOR X = 2 TO RFCNT
          DFNUM(X) = X
          RFSEP(X) = MINGAP: RFLTH(X) = 6
          RFLBL$(X) = STR$(2 - X)
      NEXT X
      CALL GetRFLD("08001", FX)
      IF FOUND THEN
         RDFFLD = RFLDS(FX)
         O = ASC(RDFFLD.ALTH) + ASC(RDFFLD.LLTH): L = ASC(RDFFLD.HLTH)
         RFLBL$(RFCNT) = MID$(RFLD$(FX), O + 1, L)
      END IF

58525 ' Check Report Width
      GOSUB 58530
      IF RWIDTH <= BWIDTH THEN 58540
      GOTO 58535

58530 ' Determine Required Width
      RWIDTH = 1: FX = 0
      WHILE FX < PAGE.DFCNT
          FX = FX + 1: X = DFNUM(FX): FSEP = RFSEP(X)
          IF (X > 0) AND (FSEP > 0) THEN
             Y = RFLTH(X): Z = LEN(RFLBL$(X))
             IF VRTCL THEN
                IF RWIDTH < Z THEN RWIDTH = Z
                ELSE
                IF Y < Z THEN Y = Z
                RWIDTH = RWIDTH + FSEP + Y
             END IF
          END IF
      WEND
      IF VRTCL THEN
         RFLTH(0) = RWIDTH
         RWIDTH = RWIDTH + (AMAXLV + DMAXLV + 2) * (6 + DFPP(0))
      END IF
      RETURN

58535 ' Check Report Width
      RPWIDTH = RWIDTH + PDFFORMS.IMARGIN + PDFFORMS.OMARGIN
      PWIDTH = BWIDTH + PDFFORMS.IMARGIN + PDFFORMS.OMARGIN
      IF PWIDTH < RPWIDTH THEN
         SYSVAR.NVAR1 = PWIDTH: SYSVAR.NVAR2 = RPWIDTH
         CALL PutMSG("XXXXM009")
         IF A$ <> SNGLKEY$(1) THEN OKAY = 0: GOTO 58560
         BWIDTH = RWIDTH
         IF LEN(PLINE$) < RPWIDTH THEN PLINE$ = SPACE$(RPWIDTH)
      END IF
      IF VRTCL THEN OKAY = -1: GOTO 58560

58540 ' Adjust to Fit
      IF NOT (RWIDTH < BWIDTH) THEN 58550
      GAPS = 0: FX = 2
      WHILE FX <= PAGE.DFCNT
          X = DFNUM(FX)
          IF (X > 0) AND (RFSEP(X) > 0) THEN
             GAPS = GAPS + 1
          END IF
          FX = FX + 1
      WEND
      GAP = (BWIDTH - RWIDTH) \ GAPS
      GAP1 = (BWIDTH - RWIDTH) - GAPS * GAP
      IF GAP >= (MAXGAP - MINGAP) THEN GAP = MAXGAP - MINGAP: GAP1 = 0
      FX = 2
      WHILE FX <= PAGE.DFCNT
          X = DFNUM(FX): FSEP = RFSEP(X)
          IF (X > 0) AND (FSEP > 0) THEN
             RFSEP(X) = FSEP + GAP - (GAP1 > 0): GAP1 = GAP1 - 1
          END IF
          FX = FX + 1
      WEND
      GOSUB 58530

58550 ' Determine Field Print Positions
      FX = 0
      IF RWIDTH < BWIDTH THEN PP = (BWIDTH - RWIDTH) \ 2 ELSE PP = 0
      WHILE FX < PAGE.DFCNT
          FX = FX + 1: X = DFNUM(FX): FSEP = RFSEP(X)
          IF (X > 0) AND (FSEP > 0) THEN
             FL = RFLTH(X): HL = LEN(RFLBL$(X))
             Z = FL: IF Z < HL THEN Z = HL
             PP = PP + FSEP + (Z + 1) \ 2 - (FL + 1) \ 2
             DFPP(FX) = PP
             PP = PP + Z
          END IF
      WEND
      OKAY = -1

58560 ' Return to Calling Program
      PAGE.BWIDTH = BWIDTH
END SUB

SUB BldSumTITLE
SHARED DFNUM(), DFPP(), VRTCL
      AMAXLV = ASC(WFHDR.AMAXLV): DMAXLV = ASC(WFHDR.DMAXLV)
      RELSUMOPT = (AMAXLV > 0) AND (DMAXLV > 0)
     ' Build Summary Report Title Line
      TL$ = SPACE$(LEN(PLINE$)): Y = 0
      IF NOT VRTCL THEN
         FOR FX = 1 TO PAGE.DFCNT: X = DFNUM(FX): PP = DFPP(FX)
             IF (X > 0) AND (PP > 0) THEN
                FL = RFLTH(X): HL = LEN(RFLBL$(X))
                IF HL > 0 THEN
                   PP = PP + (FL + 1) \ 2 - (HL + 1) \ 2
                   IF PP > 0 THEN MID$(TL$, PP, HL) = RFLBL$(X)
                END IF
             END IF
         NEXT FX
         ELSE
         LSET TL$ = RFLBL$(1)
         TL2$ = SPACE$(PAGE.BWIDTH)
         PP = RFLTH(0) + DFPP(0)
         X = AMAXLV + DMAXLV
         Y = 0: Y$ = SPACE$(6): IF AMAXLV > 0 THEN Z = -1 ELSE Z = 1
         WHILE X > -1
            RSET Y$ = STR$(Y * Z)
            MID$(TL$, PP, 6) = Y$
            MID$(TL2$, PP, 6) = STRING$(6, "-")
            PP = PP + 6 + DFPP(0): Y = Y + 1: X = X - 1
         WEND
         CALL GetRFLD("08001", FX)
         IF FOUND THEN
            RDFFLD = RFLDS(FX)
            O = ASC(RDFFLD.ALTH) + ASC(RDFFLD.LLTH): L = ASC(RDFFLD.HLTH)
            RSET Y$ = MID$(RFLD$(FX), O + 1, L)
            MID$(TL$, PP, 6) = Y$
            MID$(TL2$, PP, 6) = STRING$(6, "-")
         END IF
      END IF
      X = LEN(RTRIM$(TL$))
      IF X > 0 THEN
         HTFLMAX = HTFLMAX + 1
         REDIM PRESERVE HTFLINE$(HTFLMAX)
         HTFLINE$(HTFLMAX) = "T00" + RTRIM$(TL$)
         IF VRTCL THEN
            HTFLMAX = HTFLMAX + 1
            REDIM PRESERVE HTFLINE$(HTFLMAX)
            HTFLINE$(HTFLMAX) = "T00" + RTRIM$(TL2$)
         END IF
      END IF
      OKAY = -1
END SUB

FUNCTION BLOODLINE$ (OPT)
   DGL = WFGLREC.DGL
   X = DGL
   IF OPT = 1 THEN REDIM GBL(DGL), PARPTR&(DGL)
   WHILE X > 0
      GBL(X) = ASC(WFRELREC.CHNUM)
      IF WFRELREC.PARPTR <> PARPTR&(X) THEN
         CALL RWRKGet("PTRREC", RC, WFRELREC.PARPTR)
         PARPTR&(X) = WFRELREC.PARPTR: X = X - 1
         ELSE
         X = 0
      END IF
   WEND
   IF DGL > 0 THEN
      X$ = SPACE$(4 * DGL + 2)
      IF OPT = 1 THEN MID$(X$, 1, 1) = "(": W = 2 ELSE W = 1
      WHILE X < DGL
         X = X + 1: Z = GBL(X) AND 255
         IF OPT = 1 THEN
            MID$(X$, W, 3) = LTRIM$(STR$(Z))
            W = INSTR(W, X$, " ")
            ELSE
            MID$(X$, W, 2) = SFILL$(Z, 2)
            W = W + 2
         END IF
         Z = GBL(X) \ 256
         IF Z > 0 THEN
            MID$(X$, W, 1) = MID$("*?", Z + 1, 1)
            IF OPT = 1 THEN W = W + 1
         END IF
         IF OPT = 1 THEN MID$(X$, W, 1) = ","
         W = W + 1
      WEND
      IF OPT = 1 THEN MID$(X$, W - 1, 1) = ")"
   END IF
   BLOODLINE$ = RTRIM$(X$)

END FUNCTION

SUB GetNextMARRIAGE
SHARED SPRESEQ, MLIST(), SPCNT, SPRNO, RELID
   IF SPRESEQ THEN
      IF SPCNT > 0 THEN SPCNT = SPCNT - 1
      SPRNO = MLIST(SPCNT)
      ELSE
      IF RELID = FF3SPOUSE.SP1ID THEN
         SPRNO = FF3SPOUSE.SP1NXT
         ELSE
         IF RELID = FF3SPOUSE.SP2ID THEN
            SPRNO = FF3SPOUSE.SP2NXT
            ELSE
            SPRNO = 0
         END IF
      END IF
   END IF

END SUB

SUB XDETReport (WTYPE)
SHARED DFNUM(), DFPP()
SHARED RELID, PRPAGE(), PRLINE(), RX()
SHARED TMBY, TMEY
SHARED SPRESEQ, SPCNT, SPRNO, MAXMarry, MLIST()
     BD$ = SPACE$(12): DD$ = SPACE$(12): AGE$ = "   ": ANVSRC$ = " "
     BPL$ = SPACE$(41): MPL$ = BPL$: DPL$ = BPL$
     D1D$ = BD$: D2D$ = BD$: D1PL$ = BPL$: D2PL$ = BPL$: AGE2$ = AGE$
     NUMED$ = SPACE$(6)
     SELECT CASE WTYPE
        CASE 1: XRPT$ = "ARPT"
        CASE 2: XRPT$ = "DRPT"
        CASE 3: XRPT$ = "RRPT"
     END SELECT
     CALL RptOPEN(XRPT$): IF NOT OKAY THEN 1040
     WOPEN$ = "READQ"
1010 ' Verify Files by Opening
     CALL FAMOpen("READ"): IF NOT OKAY THEN 1040
     CALL RWRKOpen(WOPEN$, WTYPE, 0): IF OKAY THEN 1050
     IF WOPEN$ = "READQ" THEN
        CALL FAMClose
        LSET WOPEN$ = "READ": GOTO 1010
     END IF

1040 ' Bad Return
     CLOSE : ENV.FFOPEN = 0: EXIT SUB

1050 ' Get Print Time Options
     DMAXLV = ASC(WFHDR.DMAXLV)
     CALL RWRKGet("BASE", RC, PTR&): IF RC <> 1 THEN 1040
     SYSVAR.SUBJECT = FMTNAME$(1)
     RPTOPTION.OUTLINE = 0
     IF DMAXLV = 0 THEN
        GENLIST = -1
        ELSE
        CALL PutMSG("RELRM011"): IF A = 27 THEN 1040
        GENLIST = (A$ = "1")
        IF NOT GENLIST THEN
           CALL PutMSG("RELRM012"): IF A = 27 THEN 1040
           RPTOPTION.OUTLINE = (A$ <> "2")
        END IF
     END IF

     CALL PutMSG("RELRM013"): IF A = 27 THEN 1040
     RPTOPTION.FREEFORM = (A$ = "2"): RPTOPTION.TIMELINE = (A$ = "3")

1060 IF RPTOPTION.TIMELINE THEN
        CALL FmtFIND("XXXXS030")
        CALL MsgFORMAT: IF A = 27 THEN 1040
        CALL FmtFindFLD("TLBY"): IF FOUND THEN TMBY = VAL(UTXT$(FFLD.UTXT))
        CALL FmtFindFLD("TLEY"): IF FOUND THEN TMEY = VAL(UTXT$(FFLD.UTXT))
        IF (TMEY - TMBY) < 1 THEN CALL ErrBEEP(0): GOTO 1060
     END IF

1200 ' Setup Report Tables
     CALL BldDetDEF: IF NOT OKAY THEN 1040
     PRDAGE = (RPTOPTION.PRAGE = 2)
     PRID2 = ((RPTOPTION.PRID AND 2) > 0)
     RPTOPTION.PRID = (RPTOPTION.PRID > 0)
     X = LEN(PLINE$)
     CALL PrintOPEN: IF NOT OKAY THEN 1040
     PLINE$ = SPACE$(X): PL2$ = PLINE$
     WFHDR.WFRULES = CHR$((ASC(WFHDR.WFRULES) AND 252) - (GENLIST <> 0) - 2 * (GENLIST = 0))
     RGL$ = SPACE$(RFLTH(24)): LINEAGE$ = SPACE$(RFLTH(2))
     XRELATION$ = SPACE$(RFLTH(3)): NFMT$ = SPACE$(RFLTH(5))
     AGE$ = SPACE$(RFLTH(11)): XAGE$ = AGE$: ANNV$ = SPACE$(RFLTH(18))
     CALL BldDetTITLE
     CALL GetRVAR("SPOUSE", X): IF FOUND THEN SPX$ = RVAR$(X)
     CALL GetRVAR("YEARS", X): IF FOUND THEN YEARS$ = RVAR$(X)
     CALL GetRVAR("ATAGE", X): IF FOUND THEN ATAGE$ = RVAR$(X)
     CALL GetRVAR("UNKNOWN", X): IF FOUND THEN Unknown$ = RVAR$(X)
     CALL GetRVAR("SEEPRIOR", X): IF FOUND THEN SeePrior$ = RVAR$(X)
     CALL GetRVAR("LINE", X): IF FOUND THEN LINEX$ = RVAR$(X)
     ASPOUSE = ((RPTOPTION.SPOUSE AND 1) > 0)
     ALLSPOUSE = ((RPTOPTION.SPOUSE AND 2) > 0)
     IF ALLSPOUSE THEN CALL RptOptGET("SP12", X): SPRESEQ = (X = 1)
     CALL RptOptGET("EXAD", X): ExtAdoptRule = (X = 1)
     IF NOT RPTOPTION.FREEFORM THEN
        CALL RptOptGET("PNML", PLNMLTH)
     END IF

     IF NOT RPTOPTION.TIMELINE THEN 1400
     TMCHR$ = "<    >"
     CALL RptOptGET("TLBL", X): TLBLCHR$ = CHR$(X)
     CALL RptOptGET("TLLL", X): TLLLCHR$ = CHR$(X)
     CALL RptOptGET("TLBD", X): TLBDCHR$ = CHR$(X)
     CALL RptOptGET("TLMD", X): TLMDCHR$ = CHR$(X)
     CALL RptOptGET("TLBC", X): TLBCCHR$ = CHR$(X)
     CALL RptOptGET("TLDD", X): TLDDCHR$ = CHR$(X)
     MID$(TMCHR$, 2, 2) = TLBDCHR$ + TLMDCHR$
     MID$(TMCHR$, 4, 2) = TLBCCHR$ + TLDDCHR$
     TMLTH = RFLTH(23): TMLINE$ = SPACE$(TMLTH)
     TMFILL$ = STRING$(TMLTH, ASC(TLBLCHR$))
     TMLIFE$ = STRING$(TMLTH, ASC(TLLLCHR$))
     TMDELT# = (TMEY - TMBY) / TMLTH

1400 ' Begin with Base Record
     REDIM PRPAGE(FF1HDR.MAXID), PRLINE(FF1HDR.MAXID)
     A = 0: LGL = -1: DUPID = 0
     PAGE.CSID = 0
     PLINEAGE# = 0: ANCID = 0
     CALL RWRKGet("BASE", RC, PTR&): IF RC <> 1 THEN 1040
     GOTO 1450

1415 ' Process next entry on chart
     CALL RWRKGet("NEXT", RC, PTR&): IF (RC < 1) OR (RC > 8) THEN 1490
     DUPID = (WFRELREC.ID < 0)

1450 ' Process Workfile record
     RELID = ABS(WFRELREC.ID)
     REPID = (PRPAGE(RELID) <> 0)
     AGL = -WFAGLREC.AGL: RGL = AGL + WFGLREC.DGL
     IF RC < 3 THEN
        PLINEAGE# = LINEAGE#: PANCID = ANCID
        LINEAGE# = WFANCREC.LINEAGE: ANCID = WFANCREC.ANCID
        ASTAT = WFANCREC.PSTAT AND 3
        IF ((DMAXLV > 0) OR (RPTOPTION.BLNKSEP)) AND (PAGE.LINE > 0) THEN CALL WriteLINE
        ELSE
        IF ((RPTOPTION.BLNKSEP) AND (PAGE.LINE > 0)) THEN CALL WriteLINE
     END IF
     ExtAdopt = ((WFRELREC.CSTAT AND 5) > 0)
     PAGE.RPAGE = 0: PAGE.RLINE = 0
     PAGE.XID = RELID: PAGE.XTYPE = 1
     RREFRec.AGL = CHR$(-AGL): RREFRec.DGL = CHR$(RGL - AGL)
     CALL RREFAdd
     ComPTR = FF1REC.COM
     CALL GetBDEvents(BD$, BPL$, DD$, DPL$, AGE$)
     BDY = FF1REC.BY: DDY = FF1REC.DY
     IF RPTOPTION.SPOUSE THEN
        PAGE.CSID = 0: MLOC = 0: MCOMPTR = 0
        SDY = 0: SDMD = 0: M2Y = 0: M2MD = 0
        S1DY = FF1REC.DY: S1DMD = FF1REC.DMD
        SPRNO = FF1REC.SPOUSE
        IF (ASPOUSE OR (DMAXLV = 0)) AND (FF1REC.SPOUSE <> 0) THEN
           GOSUB 1530
           ELSE
           IF ALLSPOUSE THEN CALL BldMLIST
        END IF
     END IF
     IF RPTOPTION.Relation THEN
        IF (REPID) AND (WFGLREC.DGL > 0) THEN
           XID = RELID: GOSUB 1550
           ELSE
           LSET XRELATION$ = Relation$("M", FF1REC.SEX, AGL, RGL, RRELCD$())
           IF ExtAdoptRule AND ExtAdopt THEN
              LSET XRELATION$ = "*" + XRELATION$
           END IF
        END IF
     END IF
     GOSUB 1800   ' Process DATA Fields
     IF A = 27 THEN 1490

     PRPAGE(RELID) = PAGE.RPAGE
     PRLINE(RELID) = PAGE.RLINE

     IF NOT ALLSPOUSE THEN 1470       ' Didn't Request ALL Marriages
     IF (DUPID) AND (WFGLREC.DGL > 0) THEN 1470
     IF DMAXLV = 0 THEN 1470          'no spouse info for ancestor rpt

     ' Process Marriage Records
     PAGE.XTYPE = 3
     WHILE SPRNO <> 0
        DUPID = 0: REPID = 0: PAGE.RPAGE = 0: PAGE.RLINE = 0
        GOSUB 1530
        PAGE.XID = PAGE.CSID: PAGE.XTYPE = 3: CALL RREFAdd
        BDY = FF1REC.BY: DDY = FF1REC.DY
        LSET BD$ = D1D$: LSET BPL$ = D1PL$
        LSET DD$ = D2D$: LSET DPL$ = D2PL$: LSET AGE$ = AGE2$
        GOSUB 1800
        IF (PAGE.RPAGE > 0) AND (PAGE.CSID > 0) THEN
           PRPAGE(PAGE.CSID) = -(PAGE.RPAGE) * (1 + 2 * REPID)
           PRLINE(PAGE.CSID) = PAGE.RLINE
        END IF
        IF A = 27 THEN SPRNO = 0
        SDY = 0: SDMD = 0
     WEND

1470 ' Finish Line and Go Get Another WF Record
     CALL PrintLINE
     IF A = 27 THEN 1490
     GOTO 1415

1490 ' END OF REPORT
     CALL RptCLOSE: IF PRT.AGAIN THEN 1400
     CALL FAMClose
     REDIM PRPAGE(1), PRLINE(1)
     EXIT SUB

1530 ' Get Information for SPRNO
     IF DMAXLV = 0 THEN GOSUB 1540: IF PAGE.CSID = 0 THEN RETURN
     IF (SPRNO <> 0) OR (PAGE.CSID > 0) THEN
        IF SPRNO <> 0 THEN CALL FF3GetRec(SPRNO)
        MCOMPTR = FF3SPOUSE.COM: MLOC = FF3SPOUSE.MLOC
        M1Y = FF3SPOUSE.D1Y: M1MD = FF3SPOUSE.D1MD
        M2Y = FF3SPOUSE.D2Y: M2MD = FF3SPOUSE.D2MD
        IF FF3SPOUSE.SP1ID = RELID THEN
           PAGE.CSID = FF3SPOUSE.SP2ID
           ELSE
           PAGE.CSID = FF3SPOUSE.SP1ID
        END IF
        IF ALLSPOUSE THEN CALL GetNextMARRIAGE ELSE SPRNO = 0
        IF PAGE.CSID > 0 THEN
           CALL FF1GetRec(PAGE.CSID)
           IF PAGE.XTYPE = 3 THEN ComPTR = FF1REC.COM
           CALL GetBDEvents(D1D$, D1PL$, D2D$, D2PL$, AGE2$)
           SDY = FF1REC.DY: SDMD = FF1REC.DMD
           IF PAGE.XTYPE = 1 THEN
              CALL FF1GetRec(RELID)
              ELSE
              REPID = (PRPAGE(PAGE.CSID) <> 0)
              IF RPTOPTION.Relation THEN
                 IF REPID THEN
                    XID = PAGE.CSID: GOSUB 1550
                    ELSE
                    LSET XRELATION$ = " "
                 END IF
              END IF
           END IF
        END IF
        IF RPTOPTION.PRAGE THEN GOSUB 1545
        IF RPTOPTION.PLACES THEN
           IF MLOC THEN
              CALL FF3GetRec(MLOC): LSET MPL$ = MPlace
              ELSE
              LSET MPL$ = ""
           END IF
        END IF
     END IF
     RETURN

1540 ' Locate Marriage of Ancestors
     PAGE.CSID = 0: MLOC = 0
     IF (FF1REC.SEX <> FGENDR$(2)) OR (PLINEAGE# <> (LINEAGE# - 1#)) THEN RETURN
     WHILE SPRNO <> 0: CALL FF3GetRec(SPRNO)
        IF (ANCID = FF3SPOUSE.SP1ID) THEN
           SPRNO = FF3SPOUSE.SP1NXT: PAGE.CSID = FF3SPOUSE.SP2ID
           ELSE
           SPRNO = FF3SPOUSE.SP2NXT: PAGE.CSID = FF3SPOUSE.SP1ID
        END IF
        IF PAGE.CSID = PANCID THEN SPRNO = 0 ELSE PAGE.CSID = 0
     WEND
     RETURN

1545 ' Determine Last Anniversary
     IF ((M2Y > 0 OR M2MD > 0) OR (S1DY = 0 AND SDY = 0)) THEN
        LSET ANVSRC$ = " "
        ELSE
        M2Y = SDY: M2MD = SDMD: LSET ANVSRC$ = "*"
        IF ((S1DY > 0) AND ((SDY = 0) OR ((S1DY < SDY) OR (S1DY = SDY AND S1DMD < SDMD)))) THEN M2Y = S1DY: M2MD = S1DMD
     END IF
     RSET ANNV$ = DATEDif$(M1MD, M1Y, M2MD, M2Y)
     RETURN

1550      ' Format BackREF in XRELATION$
     PREF = ABS(PRPAGE(XID)): LREF = PRLINE(XID)
     LSET XRELATION$ = ">"
     MID$(XRELATION$, 2, 4) = FNX$(PREF): L1 = INSTR(XRELATION$, " ")
     MID$(XRELATION$, L1 + 1, 4) = FNX$(LREF)
     RETURN

1800 ' PRINT Prefix Fields (Those Before NAME on Data Line)
     COMBGN = 0: PAGE.COFST = DFPP(PAGE.PFCNT + 1) ' Set Continuation Offset
     FX = 0
     WHILE FX < PAGE.PFCNT
        FX = FX + 1: DFN = DFNUM(FX): X = DFPP(FX)
        IF (DFN > 0) AND (X > 0) THEN
           FLTH = RFLTH(DFN)
           PAGE.LOFST = X
           IF DFN = 24 THEN
              GOSUB 1821
              ELSE
              ON DFN GOSUB 1821, 1822, 1823, 1824, 1825, 1826, 1826, 1826, 1826
           END IF
        END IF
     WEND

1805 ' Process DATA Items After NAME Field on DATA Line
     IF ((NOT REPID) OR (RPTOPTION.Relation AND ((WFGLREC.DGL > 0) OR (PAGE.CSID > 0)))) THEN
        IF RPTOPTION.FREEFORM THEN GOSUB 2800 ELSE GOSUB 2500
        ELSE
        PAGE.LOFST = DFPP(PAGE.PFCNT + 1): FLTH = RFLTH(5): GOSUB 1825
        IF RPTOPTION.FREEFORM THEN
           PAGE.LOFST = INSTR(PAGE.LOFST, PLINE$, "   ") + 1
           ELSE
           PAGE.LOFST = DFPP(PAGE.PFCNT + 2)
        END IF
        LSET CL$ = "(" + SeePrior$: X = INSTR(1, CL$, "   ")
        IF X > 0 THEN
           PREF = ABS(PRPAGE(FF1REC.RID)): LREF = PRLINE(FF1REC.RID)
           MID$(CL$, X + 1, 3) = FNX$(PREF): X = INSTR(X, CL$, "   ")
           MID$(CL$, X + 1, LEN(LINEX$)) = LINEX$: X = INSTR(X, CL$, "  ")
           MID$(CL$, X + 1, 3) = FNX$(LREF): X = INSTR(X, CL$, "   ")
           IF X > 0 THEN MID$(CL$, X, 1) = ")"
        END IF
        IF RPTOPTION.FREEFORM THEN
           CALL PutDATA
           ELSE
           IF PAGE.LOFST > 0 THEN MID$(PLINE$, PAGE.LOFST, 255) = CL$
        END IF
     END IF
     IF NOT RPTOPTION.FREEFORM THEN
        CALL PrintLINE
        IF RPTOPTION.TWOLINES THEN
           LSET PLINE$ = PL2$
           CALL PrintLINE: LSET PL2$ = " "
        END IF
     END IF

     ' Print Comments - Individual and Marriage as appropriate
     IF NOT (REPID OR DUPID) THEN
        IF RPTOPTION.COMMENTS THEN
           COMBGN = 0: PAGE.LWIDTH = PAGE.BWIDTH - PAGE.CINDENT
           IF (PAGE.XTYPE = 3) AND MCOMPTR THEN CALL ComPRINT(MCOMPTR, COMBGN)
           IF ComPTR THEN CALL ComPRINT(ComPTR, COMBGN)
           IF (PAGE.XTYPE = 1) AND MCOMPTR THEN CALL ComPRINT(MCOMPTR, COMBGN)
        END IF
     END IF
     CALL PrintLINE
     RETURN

1820 ' Routines for Processing PREFIX Data Fields
1821 ' AGL, DGL or RGL
     IF (PAGE.XTYPE = 1) AND (RGL <> LGL) THEN
        RSET RGL$ = STR$(RGL)
        MID$(PLINE$, PAGE.LOFST, FLTH) = RGL$
        LGL = RGL
     END IF
     GOTO 1829
1822 ' Lineage Number
     IF (PAGE.XTYPE = 1) AND (RC < 3) THEN
        RSET LINEAGE$ = STR$(LINEAGE#)
        MID$(PLINE$, PAGE.LOFST, FLTH) = LINEAGE$
        MID$(PLINE$, PAGE.LOFST + FLTH, 1) = CHR$(32 + (32 - ASC("*")) * ((ASTAT AND 1) > 0) + (32 - ASC("?")) * ((ASTAT AND 2) > 0))
     END IF
     GOTO 1829
1823 ' Relationship
     MID$(PLINE$, PAGE.LOFST, FLTH) = XRELATION$
     LSET XRELATION$ = " "
     GOTO 1829
1824 ' Format BLOODLINE on PLINE$
     IF (PAGE.XTYPE = 1) OR RPTOPTION.OUTLINE THEN
        L = 1: X = PAGE.LOFST
        WHILE (GBL(L) <> 0) AND (L <= WFGLREC.DGL)
           IF PAGE.XTYPE = 1 THEN
              C = GBL(L) AND 255: S = GBL(L) \ 256
              IF (RPTOPTION.OUTLINE AND L < WFGLREC.DGL) THEN
                 MID$(PLINE$, X + 1, 1) = "."
                 ELSE
                 RSET NUMED$ = STR$(C)
                 MID$(PLINE$, X, 2) = RIGHT$(NUMED$, 2)
                 MID$(PLINE$, X + 2, 1) = CHR$(32 + (32 - ASC("*")) * ((S AND 1) > 0) + (32 - ASC("?")) * ((S AND 2) > 0))
              END IF
           END IF
           X = X + 3: L = L + 1
        WEND
     END IF
     GOTO 1829
1825 ' Name - and Outline Child Numbers
     NMLTH = FLTH
     IF RPTOPTION.OUTLINE THEN
        GOSUB 1824
        IF NOT RPTOPTION.FREEFORM THEN NMLTH = FLTH - (X - PAGE.LOFST)
        PAGE.LOFST = X
     END IF
     IF PAGE.XTYPE = 3 THEN
        MID$(PLINE$, PAGE.LOFST, 255) = SPX$
        X = INSTR(PAGE.LOFST, PLINE$, "  ") + 1
        IF NOT RPTOPTION.FREEFORM THEN NMLTH = FLTH - (X - PAGE.LOFST)
        PAGE.LOFST = X
     END IF
     IF (PAGE.XTYPE = 1) OR (PAGE.CSID > 0) THEN
        LSET NFMT$ = FMTNAME$(0)
        ELSE
        LSET NFMT$ = "(" + Unknown$ + ")"
     END IF
     MID$(PLINE$, PAGE.LOFST, NMLTH) = NFMT$
     PAGE.COFST = PAGE.LOFST
     IF PRID2 THEN GOSUB 1828
     IF RPTOPTION.FREEFORM THEN
        PAGE.LOFST = INSTR(PAGE.LOFST, PLINE$, "   ")
     END IF
     GOTO 1829
1826 ' ID Numbers - First Data Line
     SELECT CASE DFN
        CASE 6: MID$(PLINE$, PAGE.LOFST, 5) = SFILL$(FF1REC.RID, 5)
        CASE 7: MID$(PLINE$, PAGE.LOFST, 5) = SFILL$(FF1REC.FID, 5)
                MID$(PLINE$, PAGE.LOFST + 5, 1) = FDSTAT$("FID")
        CASE 8: MID$(PLINE$, PAGE.LOFST, 5) = SFILL$(FF1REC.MID, 5)
                MID$(PLINE$, PAGE.LOFST + 5, 1) = FDSTAT$("MID")
        CASE 9: IF PAGE.CSID > 0 THEN MID$(PLINE$, PAGE.LOFST, 5) = SFILL$(PAGE.CSID, 5)
     END SELECT
     GOTO 1829
1828 ' ID Numbers - Second Data Line
     X = PAGE.LOFST
     MID$(PL2$, X, 255) = RFLBL$(6) + "="
     X = INSTR(X, PL2$, "   ")
     MID$(PL2$, X, 5) = SFILL$(FF1REC.RID, 5)
     MID$(PL2$, X + 7, 255) = RFLBL$(7) + "="
     X = INSTR(X + 6, PL2$, "   ")
     MID$(PL2$, X, 5) = SFILL$(FF1REC.FID, 5)
     MID$(PL2$, X + 5, 1) = FDSTAT$("FID")
     X = INSTR(X + 5, PL2$, "   ")
     MID$(PL2$, X + 2, 255) = RFLBL$(8) + "="
     X = INSTR(X, PL2$, "   ")
     MID$(PL2$, X, 5) = SFILL$(FF1REC.MID, 5)
     MID$(PL2$, X + 5, 1) = FDSTAT$("MID")
     GOTO 1829
1829 ' Return from DATA Item
     RETURN

2500 ' Print Information for Fixed Format Report - Fields After NAME
     FX = PAGE.PFCNT
     WHILE FX < PAGE.DFCNT
        FX = FX + 1: DFN = DFNUM(FX): Y = DFPP(FX)
        IF (DFN > 0) AND (Y > 0) THEN
           FLTH = RFLTH(DFN): PAGE.LOFST = Y
           ON DFN - 4 GOSUB 1825, 1826, 1826, 1826, 1826, 2501, 2502, 2510, 2515, 2510, 2520, 2525, 2520, 2521, 2530, 2535, 2530, 2600, 2610
           PAGE.LOFST = PAGE.LOFST + RFLTH(DFN)
        END IF
     WEND
     RETURN

2501 ' Gender
     IF (PAGE.XTYPE = 1) OR (PAGE.CSID > 0) THEN
        IF FF1REC.SEX = FGENDR$(1) THEN
           X = 1
           ELSE
           IF FF1REC.SEX = FGENDR$(2) THEN
              X = 2
              ELSE
              X = 3
           END IF
        END IF
        MID$(PLINE$, PAGE.LOFST, 1) = MID$(RGENDR$(X), 2, 1)
     END IF
     RETURN
2502 ' Age
     IF (PAGE.XTYPE = 1) OR (PAGE.CSID > 0) THEN
        IF (NOT (PRDAGE AND (DDY = 0))) THEN
           MID$(PLINE$, PAGE.LOFST, 10) = AGE$
        END IF
     END IF
     RETURN
2510 ' Birth Date
     IF (PAGE.XTYPE = 1) OR (PAGE.CSID > 0) THEN
        IF FF1REC.BY > 0 THEN
           MID$(PLINE$, PAGE.LOFST, 12) = BD$
        END IF
        IF RPTOPTION.TWOLINES THEN 2515
     END IF
     RETURN
2515 ' Birth Place
     IF (PAGE.XTYPE = 1) OR (PAGE.CSID > 0) THEN
        IF RPTOPTION.TWOLINES THEN
           MID$(PL2$, PAGE.LOFST, PLNMLTH) = BPL$
           ELSE
           MID$(PLINE$, PAGE.LOFST, PLNMLTH) = BPL$
        END IF
     END IF
     RETURN
2520 ' Marriage Date
     IF (PAGE.XTYPE = 3) OR (PAGE.CSID > 0) THEN
        IF M1Y > 0 THEN
           MID$(PLINE$, PAGE.LOFST, 11) = XRPTDate$(M1MD, M1Y, 1)
           IF RPTOPTION.PRAGE AND RPTOPTION.TWOLINES THEN
              IF (NOT PRDAGE) OR (M2Y > 0) THEN
                 MID$(PLINE$, PAGE.LOFST + 12, 3) = ANNV$
                 MID$(PLINE$, PAGE.LOFST + 15, 1) = ANVSRC$
                 MID$(PLINE$, PAGE.LOFST + 17, 5) = YEARS$
              END IF
           END IF
        END IF
        IF RPTOPTION.TWOLINES THEN 2525
     END IF
     RETURN
2521 ' Last Anniversary
     IF (PAGE.XTYPE = 3) OR (PAGE.CSID > 0) THEN
        IF (M1Y > 0) AND (NOT (PRDAGE AND (M2Y = 0))) THEN
           MID$(PLINE$, PAGE.LOFST, 10) = ANNV$ + ANVSRC$
        END IF
     END IF
     RETURN
2525 ' Marriage PLACE
     IF MLOC <> 0 THEN
        IF RPTOPTION.TWOLINES THEN
           MID$(PL2$, PAGE.LOFST, PLNMLTH) = MPL$
           ELSE
           MID$(PLINE$, PAGE.LOFST, PLNMLTH) = MPL$
        END IF
     END IF
     RETURN
2530 ' Death Date
     IF (PAGE.XTYPE = 1) OR (PAGE.CSID > 0) THEN
        IF (FF1REC.DY > 0) AND (FF1REC.DY < 9999) THEN
           MID$(PLINE$, PAGE.LOFST, 12) = DD$
        END IF
        IF RPTOPTION.TWOLINES THEN 2535
     END IF
     RETURN
2535 ' Death Place
     IF (PAGE.XTYPE = 1) OR (PAGE.CSID > 0) THEN
        IF RPTOPTION.TWOLINES THEN
           MID$(PL2$, PAGE.LOFST, PLNMLTH) = DPL$
           ELSE
           MID$(PLINE$, PAGE.LOFST, PLNMLTH) = DPL$
        END IF
     END IF
     RETURN

2600 ' Build Timeline
     IF BDY > 0 THEN MID$(PLINE$, PAGE.LOFST, 4) = SFILL$(BDY, 4)
     MID$(PLINE$, PAGE.LOFST + 4, 1) = "-"
     IF (DDY = 9999) THEN
        MID$(PLINE$, PAGE.LOFST + 5, 4) = " ???"
        ELSE
        IF (DDY > 0) THEN MID$(PLINE$, PAGE.LOFST + 5, 4) = SFILL$(DDY, 4)
     END IF
     RETURN
2610 ' Build Timeline
     LSET TMLINE$ = TMFILL$
     TMRID = FF1REC.RID: TMDYR = DDY
     TMYEAR = BDY: IF TMYEAR > 0 THEN TMTYPE = 1: GOSUB 2650
     IF MID$(TMCHR$, 4, 1) = " " THEN 2630
     TMTYPE = 3: TMCHID = FF1REC.OLDCH
     WHILE TMCHID <> 0: CALL FF1GetRec(TMCHID)
        CALL GetBDEvents("", "", "", "", "")
        TMYEAR = FF1REC.BY: IF TMYEAR > 0 THEN GOSUB 2650
        IF TMRID = FF1REC.FID THEN TMCHID = FF1REC.FCH:  ELSE TMCHID = FF1REC.MCH
     WEND
     IF TMRID <> FF1REC.RID THEN CALL FF1GetRec(TMRID)
2630 IF MID$(TMCHR$, 3, 1) = " " THEN 2640
     TMTYPE = 2: TMSPRNO = FF1REC.SPOUSE
     WHILE TMSPRNO <> 0: CALL FF3GetRec(TMSPRNO)
        TMYEAR = FF3SPOUSE.D1Y: IF TMYEAR > 0 THEN GOSUB 2650
        IF TMRID = FF3SPOUSE.SP1ID THEN TMSPRNO = FF3SPOUSE.SP1NXT:  ELSE TMSPRNO = FF3SPOUSE.SP2NXT
     WEND
2640 TMYEAR = TMDYR
     IF (TMYEAR > 0) AND (TMYEAR < 9999) THEN TMTYPE = 4: GOSUB 2650
     MID$(PLINE$, PAGE.LOFST, TMLTH) = TMLINE$
     RETURN

2650 ' Process Event
     IF TMYEAR > TMEY THEN
        TMSYM = 5: TMX = TMLTH
        ELSE
        TMSYM = TMTYPE
        TMX = (TMYEAR - TMBY) / TMDELT#: IF TMX < 0 THEN TMX = 1: TMSYM = 0
        IF TMX < 1 THEN
           TMX = 1
           ELSE
           IF TMX > TMLTH THEN TMSYM = 5: TMX = TMLTH
        END IF
     END IF
     MID$(TMLINE$, TMX, 1) = MID$(TMCHR$, TMSYM + 1, 1)
     IF (TMTYPE = 1) AND (TMX < TMLTH) AND (TMDYR < 9999) THEN
        MID$(TMLINE$, TMX + 1, TMLTH) = TMLIFE$
     END IF
     IF (TMTYPE = 4) AND (TMX < TMLTH) THEN
        MID$(TMLINE$, TMX + 1, TMLTH) = TMFILL$
     END IF
     RETURN

2800 ' Print Information for Free Format Report
     'IF (RPTOPTION.PRID AND RPTOPTION.COMMENTS THEN
     '   PAGE.LOFST = PAGE.LOFST + 1
     '   LSET CL$ = "(" + RFLBL$(6): X = INSTR(1, CL$, "   ")
     '   MID$(CL$, X, 5) = FNX$(FF1REC.RID): X = INSTR(X, CL$, "   ")
     '   MID$(CL$, X, 255) = "," + RFLBL$(7): X = INSTR(X, CL$, "   ")
     '   MID$(CL$, X, 5) = FNX$(FF1REC.FID): X = INSTR(X, CL$, "   ")
     '   MID$(CL$, X, 1) = FDSTAT$("FID"): X = INSTR(X, CL$, "   ")
     '   MID$(CL$, X, 255) = "," + RFLBL$(8): X = INSTR(X, CL$, "   ")
     '   MID$(CL$, X, 5) = FNX$(FF1REC.MID): X = INSTR(X, CL$, "   ")
     '   MID$(CL$, X, 1) = FDSTAT$("MID"): X = INSTR(X, CL$, "   ")
     '   MID$(CL$, X, 1) = ")"
     '   CALL PutDATA
     'END IF
     PAGE.LOFST = DFPP(PAGE.PFCNT + 1): FLTH = RFLTH(5): GOSUB 1825
     PAGE.LOFST = INSTR(PAGE.LOFST, PLINE$, "   ")
     IF (PAGE.XTYPE = 1) OR (PAGE.CSID > 0) THEN
        IF RPTOPTION.PRSEX THEN
           IF FF1REC.SEX = FGENDR$(1) THEN
              X = 1
              ELSE
              IF FF1REC.SEX = FGENDR$(2) THEN
                 X = 2
                 ELSE
                 X = 3
              END IF
           END IF
           LSET CL$ = ", " + MID$(RGENDR$(X), 3, 10)
           CALL PutDATA
        END IF
     END IF

     ' AGE - if Still Living
     IF (PAGE.XTYPE = 1) OR (PAGE.CSID > 0) THEN
        IF (RPTOPTION.PRAGE) AND (NOT PRDAGE) AND (DDY = 0) THEN
           LSET CL$ = ", " + RFLBL$(11) + " " + LTRIM$(AGE$)
           CALL PutDATA
        END IF
     END IF
  
     ' Birth Date/Place
     IF (PAGE.XTYPE = 1) OR (PAGE.CSID > 0) THEN
        IF RPTOPTION.PLACES THEN PLTH = LEN(RTRIM$(BPL$))
        IF ((BD$ <> SPACE$(12)) AND RPTOPTION.DATES) OR (PLTH > 0) THEN
           LSET CL$ = ", " + RFLBL$(12): X = 1
           IF RPTOPTION.DATES AND (BD$ <> SPACE$(12)) THEN
              X = INSTR(X, CL$, "   ")
              MID$(CL$, X + 1, 12) = LTRIM$(BD$)
           END IF
           IF PLTH > 0 THEN
              X = INSTR(X, CL$, "   ")
              MID$(CL$, X + 1, 255) = RFLBL$(13) + " " + BPL$
           END IF
           CALL PutDATA
        END IF
     END IF

     ' Marriage DATE/Place/Anniv
     IF (PAGE.XTYPE = 3) OR (PAGE.CSID > 0) THEN
        IF RPTOPTION.PLACES THEN PLTH = LEN(RTRIM$(MPL$))
        IF (((M1Y <> 0) OR (M1MD <> 0)) AND RPTOPTION.DATES) OR (PLTH > 0) THEN
           LSET CL$ = ", " + RFLBL$(15): X = 1
           IF RPTOPTION.DATES AND ((M1Y > 0) OR (M1MD > 0)) THEN
              X = INSTR(X, CL$, "   ")
              MID$(CL$, X + 1, 11) = XRPTDate$(M1MD, M1Y, 0)
           END IF
           IF PLTH > 0 THEN
              X = INSTR(X, CL$, "    ")
              MID$(CL$, X + 1, 255) = RFLBL$(16) + " " + MPL$
           END IF
           IF (M1Y > 0) AND (RPTOPTION.PRAGE) THEN
              IF NOT (PRDAGE AND (M2Y = 0)) THEN
                 X = INSTR(X, CL$, "  ")
                 MID$(CL$, X + 1, 255) = RFLBL$(18) + " " + LTRIM$(ANNV$)
                 X = INSTR(X, CL$, "  ")
                 MID$(CL$, X, 1) = ANVSRC$
                 X = INSTR(X, CL$, "  ")
                 MID$(CL$, X + 1, 255) = YEARS$
              END IF
           END IF
           CALL PutDATA
        END IF
     END IF

     ' Death DATE/Place/Age
     IF (PAGE.XTYPE = 1) OR (PAGE.CSID > 0) THEN
        IF RPTOPTION.PLACES THEN PLTH = LEN(RTRIM$(DPL$))
        IF (RPTOPTION.DATES AND (DD$ <> SPACE$(12))) OR (PLTH > 0) THEN
           LSET CL$ = ", " + RFLBL$(19): X = 1
           IF RPTOPTION.DATES AND (DD$ <> SPACE$(12)) THEN
              X = INSTR(X, CL$, "  ")
              MID$(CL$, X + 1, 12) = LTRIM$(DD$)
           END IF
           IF PLTH > 0 THEN
              X = INSTR(X, CL$, "  ")
              MID$(CL$, X + 1, 255) = RFLBL$(20) + " " + DPL$
           END IF
           IF (RPTOPTION.PRAGE) AND (AGE$ <> "   ") THEN
              X = INSTR(X, CL$, "  ")
              MID$(CL$, X + 1, 255) = ATAGE$ + " " + LTRIM$(AGE$)
           END IF
           CALL PutDATA
        END IF
     END IF
     RETURN

END SUB

SUB XDUPReport (WTYPE)
     SELECT CASE WTYPE
        CASE 1: DRPT$ = "ADUP"
        CASE 2: DRPT$ = "DDUP"
     END SELECT
     CALL RptOPEN(DRPT$): IF NOT OKAY THEN GOTO S2690
     WOPEN$ = "READQ"
S2610: ' Verify Files by Opening
     CALL FAMOpen("READ"): IF NOT OKAY THEN GOTO S2690
     CALL RWRKOpen(WOPEN$, WTYPE, 0): IF OKAY THEN GOTO S2615
     IF WOPEN$ = "READQ" THEN
        CALL FAMClose
        LSET WOPEN$ = "READ": GOTO S2610
        ELSE
        GOTO S2690
     END IF
S2615: ' Get BASE Record
     CALL PutMSG("RELRM016")
     CALL RWRKGet("BASE", RC, PTR&): IF RC <> 1 THEN GOTO S2690
     DupCNT1 = 0: DupCNT2 = 0: DUPMAX = 100
     REDIM DUPID(FF1HDR.MAXID), DUPPTR&(DUPMAX)
S2620: ' Check for Duplicate Record
     IF (WFRELREC.ID > 0) AND (WFRELREC.NREF <> 0) THEN
        DupCNT1 = DupCNT1 + 1
        IF DupCNT1 > DUPMAX THEN
           DUPMAX = DUPMAX + 20
           REDIM PRESERVE DUPPTR&(DUPMAX)
        END IF
        DUPID(WFRELREC.ID) = DupCNT1
        DUPPTR&(DupCNT1) = PTR&
     END IF
     IF (WFRELREC.ID < 0) OR (WFRELREC.NREF <> 0) THEN DupCNT2 = DupCNT2 + 1
     CALL RWRKGet("WRKREC", RC, PTR&)
     IF (0 < RC) AND (RC < 9) THEN GOTO S2620
     SYSVAR.NVAR1 = DupCNT1: SYSVAR.NVAR2 = DupCNT2
     IF SYSVAR.NVAR1 = 0 THEN
        CALL PutMSG("RELRM017"): GOTO S2690
     END IF
     CALL PrintOPEN: IF NOT OKAY THEN GOTO S2690

S2650: ' Print Report of Duplicate Ancestors/Descendants
     CALL BldDupDEF
     X1 = LEN(RFLBL$(1)): X2 = LEN(RFLBL$(2))
     X3 = LEN(RFLBL$(3)): L3 = RFLTH(3): X4 = LEN(RFLBL$(4))
     CL$ = SPACE$(4 * ASC(WFHDR.DMAXLV) + 2)
     PAGE.COFST = X1 + X3 + X4 + L3 + 13
     CALL RWRKGet("BASE", RC, PTR&): IF RC <> 1 THEN GOTO S2690
     SYSVAR.SUBJECT = FMTNAME$(1)
     SYSVAR.NVAR1 = DupCNT1
     FOR ID = 1 TO FF1HDR.MAXID
         DUPIX = DUPID(ID)
         IF DUPIX > 0 THEN
            RELPTR& = DUPPTR&(DUPIX)
            CALL FF1GetRec(ID)
            PAGE.RPAGE = 0: PAGE.RLINE = 0
            PAGE.XID = ID: PAGE.XTYPE = 1: CALL RREFAdd
            LSET PLINE$ = RFLBL$(1)
            MID$(PLINE$, X1 + 1, 6) = STR$(ID)
            MID$(PLINE$, X1 + 8, X2) = RFLBL$(2)
            MID$(PLINE$, X1 + X2 + 9, 52) = FMTNAME$(0)
            CALL PrintLINE: IF A = 27 THEN RELPTR& = 0
            WHILE RELPTR& <> 0:
               CALL RWRKGet("RELREC", RC, RELPTR&)
               RELPTR& = WFRELREC.NREF
               AGL = WFGLREC.AGL: DGL = WFGLREC.DGL: RGL = AGL + DGL
               MID$(PLINE$, X1 + 10, X3) = RFLBL$(3)
               MID$(PLINE$, X1 + X3 + 11, L3) = Relation$("M", FF1REC.SEX, AGL, RGL, RRELCD$())
               MID$(PLINE$, X1 + X3 + L3 + 12, X4) = RFLBL$(4)
               X = PAGE.COFST
               IF AGL <> 0 THEN
                  MID$(PLINE$, X, 255) = LTRIM$(STR$(WFANCREC.LINEAGE))
                  X = INSTR(X, PLINE$, " ")
               END IF
               IF DGL <> 0 THEN
                  LSET CL$ = BLOODLINE$(1): CALL PutDATA
               END IF
               CALL PrintLINE: IF A = 27 THEN RELPTR& = 0
            WEND: IF A = 27 THEN ID = FF1HDR.MAXID
         END IF
     NEXT ID
     CALL RptCLOSE: IF PRT.AGAIN THEN GOTO S2650
S2690: ' Return to Main Program
     CLOSE : ENV.FFOPEN = 0

END SUB

SUB XSUMReport (WTYPE)
SHARED DFNUM(), DFPP()
SHARED DATECNT, DATERNG, PLACCNT, MARRCNT, RWIDTH, VRTCL
SHARED SUMBASE AS SUMREC, SUM AS SUMREC, TOT AS SUMREC
     ' Produce Generation Summary Report
     RELSUMOPT = (WTYPE = 3): GENSUMOPT = (WTYPE < 3)
     IF GENSUMOPT THEN
        CALL PutMSG("XXXXM044")
        IF A = 27 THEN GOTO S1040
        VRTCL = (A$ = "2")
     END IF
     SELECT CASE WTYPE
        CASE 1: SRPT$ = "ASUM"
        CASE 2: SRPT$ = "DSUM"
        CASE 3: SRPT$ = "RSUM"
     END SELECT
     CALL RptOPEN(SRPT$): IF NOT OKAY THEN GOTO S1040
     WOPEN$ = "READQ"
S1010: ' Verify Files by Opening
     CALL FAMOpen("READ"): IF NOT OKAY THEN GOTO S1040
     CALL RWRKOpen(WOPEN$, WTYPE, 0): IF OKAY THEN GOTO S1050
     IF WOPEN$ = "READQ" THEN
        CALL FAMClose
        LSET WOPEN$ = "READ": GOTO S1010
     END IF

S1040: ' Bad Return
     CLOSE : ENV.FFOPEN = 0: EXIT SUB

S1050: ' Prepare to Print
     CALL RWRKGet("BASE", RC, PTR&): IF RC <> 1 THEN GOTO S1040
     SYSVAR.SUBJECT = FMTNAME$(1)
     IF DMAXLV = 0 THEN
        CALL RptOptGET("CNOP", X): COUNTNOP = (X = 1)
        IF NOT VRTCL THEN CALL RptOptGET("LNOP", X): LISTNODES = (X = 1)
        LALIN = LEN(STR$(2 ^ ASC(WFHDR.AMAXLV))): ALIN$ = SPACE$(LALIN)
     END IF
     CALL RptOptGET("CDUP", X): COUNTDUPS = (X = 1)
     CALL RptOptGET("CLIN", X): COUNTCLINE = (X = 1)
     CALL RptOptGET("DTCT", X): COUNTDATES = (X = 1)
     CALL RptOptGET("DTRG", X): SHOWRANGE = (X = 1)
     CALL RptOptGET("PLCT", X): COUNTPLACES = (X = 1)
     IF COUNTPLACES THEN BPL$ = " ": DPL$ = BPL$:  ELSE BPL$ = "": DPL$ = ""
     CALL RptOptGET("MCNT", X): COUNTMARRY = (X = 1)
     SHOWREL = RPTOPTION.Relation
     IF RELSUMOPT THEN
        LORGL = -ASC(WFHDR.AMAXLV): HIRGL = ASC(WFHDR.HIRGL)
        REDIM RELCNT(LORGL TO 0, LORGL TO HIRGL)
        ELSE
        HIRGL = ASC(WFHDR.AMAXLV) + ASC(WFHDR.DMAXLV)
        IF VRTCL THEN REDIM SUMTAB(HIRGL) AS SUMREC
     END IF
     IF NOT VRTCL THEN
        CALL GetRFLD("08001", FX)
        IF FOUND THEN
           RDFFLD = RFLDS(FX)
           O = ASC(RDFFLD.ALTH): L = ASC(RDFFLD.LLTH)
           TOTAL$ = MID$(RFLD$(FX), O + 1, L)
        END IF
        IF RELSUMOPT THEN
           CALL GetRFLD("08008", FX)
           IF FOUND THEN
              RDFFLD = RFLDS(FX)
              O = ASC(RDFFLD.ALTH): L = ASC(RDFFLD.LLTH)
              ACCUM$ = MID$(RFLD$(FX), O + 1, L)
           END IF
        END IF
     END IF

S1200: ' Setup Report Tables & Open Files
     CALL BldSumDEF: IF NOT OKAY THEN GOTO S1040
     SPL = LEN(PLINE$)
     IF NOT RELSUMOPT THEN
        CALL PrintOPEN: IF NOT OKAY THEN GOTO S1040
        PLINE$ = SPACE$(SPL)
        CALL BldSumTITLE
        ' IF Ancestor or Descendant Summary, Retrieve Records by Generation
        WFHDR.WFRULES = CHR$((ASC(WFHDR.WFRULES) AND 252) + 1)
     END IF
     RGL$ = SPACE$(RFLTH(1))

S1400: ' Begin with Base Record
     SUMBASE.LOBY = 10000: SUMBASE.LOMY = 10000: SUMBASE.LODY = 10000
     SUM = SUMBASE: TOT = SUMBASE
     L = 20: LAGL = 0: LRGL = 0: PAGE.COFST = 1
     A = 0: LGL = -1: DUPID = 0: DupCNT1 = 0: DupCNT2 = 0
     IF RELSUMOPT THEN CALL PutMSG("XXXXM064")
     CALL RWRKGet("BASE", RC, PTR&): IF RC <> 1 THEN GOTO S1040
     FIRSTONE = -1: GOTO S1450

S1415: ' Process next entry on chart
     IF RELSUMOPT THEN
        CALL RWRKGet("WRKREC", RC, PTR&)
        ELSE
        CALL RWRKGet("NEXT", RC, PTR&)
     END IF
     IF (RC < 1) OR (RC > 8) THEN GOTO S1455
     DUPID = ((WFRELREC.ID < 0) OR (WFRELREC.NREF <> 0))

S1450: ' Process Workfile record
     RELID = ABS(WFRELREC.ID)
     IF (WFRELREC.ID > 0) AND (WFRELREC.NREF <> 0) THEN DupCNT1 = DupCNT1 + 1
     IF (WFRELREC.ID < 0) OR (WFRELREC.NREF <> 0) THEN DupCNT2 = DupCNT2 + 1
     AGL = -WFAGLREC.AGL: RGL = AGL + WFGLREC.DGL
     IF RELSUMOPT THEN GOSUB S2200:  ELSE GOSUB S2400
     IF A <> 27 THEN GOTO S1415

S1455: ' Process Final Totals
     IF RELSUMOPT THEN
        GOSUB S2210
        ELSE
        GOSUB S2460
        IF VRTCL THEN GOSUB S2700
        CALL RptCLOSE
        IF PRT.AGAIN THEN
           IF VRTCL THEN GOTO S1455 ELSE GOTO S1400
        END IF
     END IF
     CALL FAMClose
     EXIT SUB

S2200: ' Process Record for RELATIVE Table
     IF WFRELREC.ID > 0 THEN
        RELCNT(AGL, RGL) = RELCNT(AGL, RGL) + 1
     END IF
     RETURN

S2210: ' Print Relationship Summary Table
      CALL PrintOPEN
      IF NOT OKAY THEN GOTO S1040
      FULLWIDTH = 0: PLINE$ = SPACE$(SPL)
      IF (PRT.OPT > 2) AND (PAGE.BWIDTH > PRT.FWIDTH - PDFFORMS.IMARGIN - PDFFORMS.OMARGIN) THEN
         CALL PutMSG("XXXXM070")
         IF ASC(A$) = 27 THEN GOTO S1040 ELSE FULLWIDTH = (A$ <> "1")
         IF FULLWIDTH THEN
            PRT.CFORMS = -1: NOSPAN = -1
            PRT.FWIDTH = PAGE.BWIDTH + PDFFORMS.IMARGIN + PDFFORMS.OMARGIN
            PLINE$ = SPACE$(PRT.FWIDTH)
         END IF
      END IF
      RELABEL$ = PLINE$: RELCNT$ = PLINE$
      CALL BldSumTITLE

S2215: ' Print Table (AGAIN?)
      REDIM AGLTOT(LORGL TO 0)
      RGLTOT = 0: RGLTOTAL = 0
      XRGL = LORGL: SREL$ = SPACE$(RFLTH(2))
      WHILE XRGL <= HIRGL: XAGL = 0: FX = 2
         LSET PLINE$ = "": CALL WriteLINE: RGLTOT = 0
         RSET RGL$ = STR$(XRGL): MID$(RELCNT$, DFPP(1), RFLTH(1)) = RGL$
         WHILE XAGL >= LORGL
            PP = DFPP(FX): FLTH = RFLTH(DFNUM(FX))
            RCNT = RELCNT(XAGL, XRGL)
            IF (XAGL > XRGL) OR (RCNT = 0) THEN
               MID$(RELCNT$, PP + FLTH - 1, 1) = "."
               ELSE
               IF SHOWREL THEN
                  RSET SREL$ = Relation$("S", "*", XAGL, XRGL, RRELCD$())
                  MID$(RELABEL$, PP, FLTH) = SREL$
               END IF
               IF RCNT > 0 THEN MID$(RELCNT$, PP, FLTH) = SFILL$(RCNT, FLTH)
               AGLTOT(XAGL) = AGLTOT(XAGL) + RCNT
               RGLTOT = RGLTOT + RCNT
            END IF
            FX = FX + 1: XAGL = XAGL - 1
         WEND
         MID$(RELCNT$, DFPP(FX), RFLTH(FX)) = SFILL$(RGLTOT, RFLTH(FX))
         IF SHOWREL THEN LSET PLINE$ = RELABEL$: CALL PrintLINE: LSET RELABEL$ = ""
         LSET PLINE$ = RELCNT$: CALL PrintLINE: LSET RELCNT$ = ""
         XRGL = XRGL + 1: RGLTOTAL = RGLTOTAL + RGLTOT
      WEND
      CALL WriteLINE
      MID$(PLINE$, DFPP(1), 255) = TOTAL$
      XAGL = 0: FX = 2
      WHILE XAGL >= LORGL
         PP = DFPP(FX): FLTH = RFLTH(DFNUM(FX))
         MID$(PLINE$, PP, FLTH) = SFILL$(AGLTOT(XAGL), FLTH)
         FX = FX + 1: XAGL = XAGL - 1
      WEND
      MID$(PLINE$, DFPP(FX), RFLTH(FX)) = SFILL$(RGLTOTAL, RFLTH(FX))
      CALL PrintLINE
      MID$(PLINE$, DFPP(1), 255) = ACCUM$
      XAGL = 0: FX = 2: AGLACCUM = 0
      WHILE XAGL >= LORGL
         AGLACCUM = AGLACCUM + AGLTOT(XAGL)
         PP = DFPP(FX): FLTH = RFLTH(DFNUM(FX))
         MID$(PLINE$, PP, FLTH) = SFILL$(AGLACCUM, FLTH)
         FX = FX + 1: XAGL = XAGL - 1
      WEND
      CALL PrintLINE
      CALL RptCLOSE
      IF PRT.AGAIN THEN GOTO S2215
      RETURN

S2400: ' Process Records for ANC/DESC Genlevel Summaries
     CALL FF1GetRec(RELID)
     IF (SUM.TOTAL = 0) OR (RGL = LRGL) THEN GOTO S2420
     IF NOT VRTCL THEN GOSUB S2550
     GOSUB S2570
     FIRSTONE = -1: LAGL = AGL: LRGL = RGL
     SUM = SUMBASE

S2420: ' Increment Summary Counters
     CALL GetBDEvents("", BPL$, "", DPL$, "")
     C = -((WFRELREC.ID > 0) OR (DUPID AND COUNTDUPS))
     IF (COUNTPLACES AND C) THEN
        IF RTRIM$(BPLACE) <> "" THEN SUM.BPL = SUM.BPL + 1
        IF RTRIM$(DPLACE) <> "" THEN SUM.DPL = SUM.DPL + 1
     END IF
     IF ((FF1REC.BY > 0) AND C) THEN
        SUM.BDT = SUM.BDT + 1
        IF (FF1REC.BY < SUM.LOBY) THEN SUM.LOBY = FF1REC.BY
        IF (FF1REC.BY > SUM.HIBY) THEN SUM.HIBY = FF1REC.BY
     END IF
     IF ((FF1REC.DY > 0) AND C) THEN
        SUM.DDT = SUM.DDT + 1
        IF FF1REC.DY < 9999 THEN
           SUM.DDTV = SUM.DDTV + 1
           IF (FF1REC.DY < SUM.LODY) THEN SUM.LODY = FF1REC.DY
           IF (FF1REC.DY > SUM.HIDY) THEN SUM.HIDY = FF1REC.DY
        END IF
     END IF
     IF (COUNTMARRY AND C) THEN
        SPREC = FF1REC.SPOUSE
        IF SPREC <> 0 THEN SUM.MARRIED = SUM.MARRIED + 1
        WHILE (SPREC <> 0)
           SUM.MARCNT = SUM.MARCNT + 1
           CALL FF3GetRec(SPREC)
           IF (FF3SPOUSE.D1Y > 0) THEN
              SUM.MDT = SUM.MDT + 1
              IF (FF3SPOUSE.D1Y < SUM.LOMY) THEN SUM.LOMY = FF3SPOUSE.D1Y
              IF (FF3SPOUSE.D1Y > SUM.HIMY) THEN SUM.HIMY = FF3SPOUSE.D1Y
           END IF
           IF (FF3SPOUSE.MLOC <> 0) THEN SUM.MPL = SUM.MPL + 1
           IF (FF3SPOUSE.SP1ID = RELID) THEN
              SPREC = FF3SPOUSE.SP1NXT
              ELSE
              SPREC = FF3SPOUSE.SP2NXT
           END IF
        WEND
     END IF
     IF ((COUNTNOP OR LISTNODES) AND C) THEN
        IF (FF1REC.FID = 0) AND (FF1REC.MID = 0) THEN
           SUM.NOP = SUM.NOP + 1
           IF LISTNODES THEN GOSUB S2500
        END IF
     END IF
     IF COUNTDUPS THEN SUM.DUPID = SUM.DUPID - DUPID
     IF COUNTCLINE THEN
        IF RC = 1 THEN SUM.CLINE = SUM.CLINE + WFAGLREC.CLINES
        IF RC = 3 THEN SUM.CLINE = SUM.CLINE + WFGLREC.CLINES
     END IF
     SUM.TOTAL = SUM.TOTAL + C
     IF FF1REC.SEX = FGENDR$(1) THEN
        SUM.MALE = SUM.MALE + C
        ELSE
        IF FF1REC.SEX = FGENDR$(2) THEN SUM.FEMALE = SUM.FEMALE + C
     END IF
     RETURN

S2460: ' Process last summary Line & Totals
     IF NOT VRTCL THEN GOSUB S2550
     IF FIRSTONE THEN GOSUB S2570
     FIRSTONE = 0
     IF NOT VRTCL THEN
        CALL WriteLINE
        MID$(PLINE$, DFPP(1), 5) = TOTAL$
        SUM = TOT: FINALTOT = -1
        GOSUB S2550
     END IF
     RETURN

S2500: ' Print Nodal Person
     IF FIRSTONE THEN
        RSET RGL$ = STR$(LRGL): MID$(PLINE$, DFPP(1), RFLTH(1)) = RGL$
        IF RPTOPTION.Relation THEN MID$(PLINE$, DFPP(2), RFLTH(2)) = Relation$("M", "*", LAGL, LRGL, RRELCD$())
        CALL PrintLINE: FIRSTONE = 0
     END IF
     FX = 2 - RPTOPTION.Relation: PP = DFPP(FX)
     RSET ALIN$ = STR$(WFANCREC.LINEAGE)
     MID$(PLINE$, PP, LALIN) = ALIN$: PP = PP + LALIN + 2
     IF FF1REC.BY > 0 THEN MID$(PLINE$, PP, 4) = ZFILL$(FF1REC.BY, 4)
     MID$(PLINE$, PP + 4, 1) = "-"
     IF FF1REC.DY > 0 THEN MID$(PLINE$, PP + 5, 4) = ZFILL$(FF1REC.DY, 4)
     MID$(PLINE$, PP + 11, 52) = FMTNAME$(0)
     CALL PrintLINE
     RETURN

S2550: ' Print Ancestor/Descendant Summary Line - Horizontal Style
     FX = 1
     WHILE FX <= PAGE.DFCNT
        X = DFNUM(FX)
        PP = DFPP(FX): FLTH = RFLTH(X)
        SELECT CASE X
           CASE 1:  IF FIRSTONE THEN RSET RGL$ = STR$(LRGL): MID$(PLINE$, PP, FLTH) = RGL$
           CASE 2:  IF FIRSTONE THEN MID$(PLINE$, PP, FLTH) = Relation$("M", "*", LAGL, LRGL, RRELCD$())
           CASE 3:  MID$(PLINE$, PP, FLTH) = SFILL$(SUM.TOTAL, FLTH)
                    IF SUM.TOTAL > (SUM.MALE + SUM.FEMALE) THEN MID$(PLINE$, PP + FLTH, 1) = "*"
           CASE 4:  IF NOT FINALTOT THEN MID$(PLINE$, PP, FLTH) = SFILL$(SUM.CLINE, FLTH)
           CASE 5:  MID$(PLINE$, PP, FLTH) = SFILL$(SUM.DUPID, FLTH)
           CASE 6:  MID$(PLINE$, PP, FLTH) = SFILL$(SUM.NOP, FLTH)
           CASE 7:  MID$(PLINE$, PP, FLTH) = SFILL$(SUM.MALE, FLTH)
           CASE 8:  MID$(PLINE$, PP, FLTH) = SFILL$(SUM.FEMALE, FLTH)
           CASE 9:  MID$(PLINE$, PP, FLTH) = SFILL$(SUM.BDT, FLTH)
           CASE 10: MID$(PLINE$, PP, FLTH) = SFILL$(SUM.BPL, FLTH)
           CASE 11: IF SUM.BDT > 0 THEN MID$(PLINE$, PP, FLTH) = ZFILL$(SUM.LOBY, FLTH)
           CASE 12: IF SUM.BDT > 0 THEN MID$(PLINE$, PP, FLTH) = ZFILL$(SUM.HIBY, FLTH)
           CASE 13: MID$(PLINE$, PP, FLTH) = SFILL$(SUM.MARRIED, FLTH)
           CASE 14: MID$(PLINE$, PP, FLTH) = SFILL$(SUM.MARCNT, FLTH)
           CASE 15: MID$(PLINE$, PP, FLTH) = SFILL$(SUM.MDT, FLTH)
           CASE 16: MID$(PLINE$, PP, FLTH) = SFILL$(SUM.MPL, FLTH)
           CASE 17: IF SUM.MDT > 0 THEN MID$(PLINE$, PP, FLTH) = ZFILL$(SUM.LOMY, FLTH)
           CASE 18: IF SUM.MDT > 0 THEN MID$(PLINE$, PP, FLTH) = ZFILL$(SUM.HIMY, FLTH)
           CASE 19: MID$(PLINE$, PP, FLTH) = SFILL$(SUM.DDT, FLTH)
           CASE 20: MID$(PLINE$, PP, FLTH) = SFILL$(SUM.DDTV, FLTH)
           CASE 21: MID$(PLINE$, PP, FLTH) = SFILL$(SUM.DPL, FLTH)
           CASE 22: IF SUM.DDTV > 0 THEN MID$(PLINE$, PP, FLTH) = ZFILL$(SUM.LODY, FLTH)
           CASE 23: IF SUM.DDTV > 0 THEN MID$(PLINE$, PP, FLTH) = ZFILL$(SUM.HIDY, FLTH)
        END SELECT
        FX = FX + 1
     WEND
     CALL PrintLINE
     RETURN

S2570: ' Update Overall Counters
     TOT.TOTAL = TOT.TOTAL + SUM.TOTAL
     TOT.NOP = TOT.NOP + SUM.NOP
     TOT.DUPID = TOT.DUPID + SUM.DUPID
     TOT.CLINE = TOT.CLINE + SUM.CLINE
     TOT.MALE = TOT.MALE + SUM.MALE
     TOT.FEMALE = TOT.FEMALE + SUM.FEMALE
     TOT.MARRIED = TOT.MARRIED + SUM.MARRIED
     TOT.MARCNT = TOT.MARCNT + SUM.MARCNT
     TOT.BPL = TOT.BPL + SUM.BPL: TOT.MPL = TOT.MPL + SUM.MPL
     TOT.DPL = TOT.DPL + SUM.DPL
     TOT.BDT = TOT.BDT + SUM.BDT: TOT.DDT = TOT.DDT + SUM.DDT
     TOT.DDTV = TOT.DDTV + SUM.DDTV: TOT.MDT = TOT.MDT + SUM.MDT
     IF (SUM.LOBY > 0) AND (SUM.LOBY < TOT.LOBY) THEN TOT.LOBY = SUM.LOBY
     IF SUM.HIBY > TOT.HIBY THEN TOT.HIBY = SUM.HIBY
     IF (SUM.LODY > 0) AND (SUM.LODY < TOT.LODY) THEN TOT.LODY = SUM.LODY
     IF SUM.HIDY > TOT.HIDY THEN TOT.HIDY = SUM.HIDY
     IF (SUM.LOMY > 0) AND (SUM.LOMY < TOT.LOMY) THEN TOT.LOMY = SUM.LOMY
     IF SUM.HIMY > TOT.HIMY THEN TOT.HIMY = SUM.HIMY
     IF VRTCL THEN SUMTAB(ABS(LRGL)) = SUM
     RETURN

S2700: ' Print Ancestor/Descendant Summary Line - Vertical Style
     FX = 3
     WHILE FX <= PAGE.DFCNT
        X = DFNUM(FX): MID$(PLINE$, 1, 255) = RFLBL$(X)
        FLTH = RFLTH(X)
        PP = RFLTH(0) + DFPP(0)
        FOR Y = 0 TO HIRGL + 1
            PP = PP + 6 - FLTH
            IF (Y <= HIRGL) THEN SUM = SUMTAB(Y) ELSE SUM = TOT
            SELECT CASE X
               CASE 3:  MID$(PLINE$, PP, FLTH) = SFILL$(SUM.TOTAL, FLTH)
                        IF SUM.TOTAL > (SUM.MALE + SUM.FEMALE) THEN MID$(PLINE$, PP + FLTH, 1) = "*"
               CASE 4:  MID$(PLINE$, PP, FLTH) = SFILL$(SUM.CLINE, FLTH)
               CASE 5:  MID$(PLINE$, PP, FLTH) = SFILL$(SUM.DUPID, FLTH)
               CASE 6:  MID$(PLINE$, PP, FLTH) = SFILL$(SUM.NOP, FLTH)
               CASE 7:  MID$(PLINE$, PP, FLTH) = SFILL$(SUM.MALE, FLTH)
               CASE 8:  MID$(PLINE$, PP, FLTH) = SFILL$(SUM.FEMALE, FLTH)
               CASE 9:  MID$(PLINE$, PP, FLTH) = SFILL$(SUM.BDT, FLTH)
               CASE 10: MID$(PLINE$, PP, FLTH) = SFILL$(SUM.BPL, FLTH)
               CASE 11: IF SUM.BDT > 0 THEN MID$(PLINE$, PP, FLTH) = ZFILL$(SUM.LOBY, FLTH)
               CASE 12: IF SUM.BDT > 0 THEN MID$(PLINE$, PP, FLTH) = ZFILL$(SUM.HIBY, FLTH)
               CASE 13: MID$(PLINE$, PP, FLTH) = SFILL$(SUM.MARRIED, FLTH)
               CASE 14: MID$(PLINE$, PP, FLTH) = SFILL$(SUM.MARCNT, FLTH)
               CASE 15: MID$(PLINE$, PP, FLTH) = SFILL$(SUM.MDT, FLTH)
               CASE 16: MID$(PLINE$, PP, FLTH) = SFILL$(SUM.MPL, FLTH)
               CASE 17: IF SUM.MDT > 0 THEN MID$(PLINE$, PP, FLTH) = ZFILL$(SUM.LOMY, FLTH)
               CASE 18: IF SUM.MDT > 0 THEN MID$(PLINE$, PP, FLTH) = ZFILL$(SUM.HIMY, FLTH)
               CASE 19: MID$(PLINE$, PP, FLTH) = SFILL$(SUM.DDT, FLTH)
               CASE 20: MID$(PLINE$, PP, FLTH) = SFILL$(SUM.DDTV, FLTH)
               CASE 21: MID$(PLINE$, PP, FLTH) = SFILL$(SUM.DPL, FLTH)
               CASE 22: IF SUM.DDTV > 0 THEN MID$(PLINE$, PP, FLTH) = ZFILL$(SUM.LODY, FLTH)
               CASE 23: IF SUM.DDTV > 0 THEN MID$(PLINE$, PP, FLTH) = ZFILL$(SUM.HIDY, FLTH)
            END SELECT
            PP = PP + FLTH + DFPP(0)
        NEXT Y
        CALL PrintLINE
        FX = FX + 1
     WEND
     RETURN

END SUB

