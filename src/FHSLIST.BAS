DECLARE SUB GetLastRES ()
DECLARE SUB BldMLIST ()
DECLARE SUB GetNextMARRIAGE ()
DECLARE SUB DETAILReport ()
DECLARE SUB BldDetDEF ()
DECLARE SUB BldDetTITLE ()
DECLARE SUB XSUMReport ()
DECLARE SUB BldSumTITLE ()
DECLARE SUB BldSumDEF (WTYPE%)
DECLARE SUB SLCTLoad ()
DECLARE SUB ChkRULE (XRULE AS ANY, R%)
DECLARE SUB SLCTSave ()
DECLARE SUB SelectFILE (OPT$)
DECLARE SUB SelectID ()
DECLARE SUB TEXTSearch ()
DECLARE SUB BLDListDEF ()
DECLARE SUB GetRES ()
DECLARE SUB LISTReport ()
DECLARE SUB SelectSET (OPT$)
DECLARE SUB ULFILL (X$)
DECLARE FUNCTION DATEDif$ (D1MD%, D1Y%, D2MD%, D2Y%)
DECLARE FUNCTION DATEY% (XDATE$, XDFMT%)
DECLARE FUNCTION DATEMD% (XDATE$, XDFMT%)
DECLARE FUNCTION TIMEHHMM% (XTIME$)
DECLARE FUNCTION FDSTAT$ (X$)
DECLARE FUNCTION FFRBA& (RNUM%, RLTH%)
DECLARE FUNCTION FFRNUM% (RBA&, RLTH%)
DECLARE FUNCTION FMTNAME$ (NFMT%)
DECLARE FUNCTION Relation$ (RT$, RSEX$, AGL%, RGL%, RCODE$())
DECLARE FUNCTION SFILL$ (X%, L%)
DECLARE FUNCTION ZFILL$ (X%, L%)
DECLARE FUNCTION UCX$ (X$)
DECLARE FUNCTION XRPTDate$ (MD%, Y%, O%)
DECLARE FUNCTION XSCRNDate$ (MD%, Y%)
TYPE SUMREC
   TOTAL    AS INTEGER
   MALE     AS INTEGER
   FEMALE   AS INTEGER
   GROUPS   AS INTEGER
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
   PN$ = " FAMILY HISTORY SYSTEM - FHSLIST - Search/Select/LIST Program"
   CY$ = " (C) Copyright 1986,1987,1988,1990,1991,1992,1993,1996 by Phillip E. Brown"
   IF CONFIG.BP = 0 THEN RUN "FHSINIT"
   REDIM STYPE$(5)
   STYPE$(1) = "Single ID": STYPE$(2) = "ID List"
   STYPE$(3) = "REL Work File": STYPE$(4) = "SELECT File"
   STYPE$(5) = "Full File"
   DEF FNX$ (A) = LTRIM$(STR$(A))
   REDIM DFNUM(1), DFPP(1)
   REDIM RELATE(1), RELPTR&(1)
   REDIM SELSET(1), NDXID(1)
   MAXMarry = 10: REDIM MLIST(MAXMarry)
   WTYPE = MO: BRID = 0
   CALL WinINIT("LR1")
   CALL WinPREP(1): PGMRTRN$ = "FHSMENU"
   FUN = 0: CALL RptLOAD("LIST")
   DetailRptTITLE$ = RTRIM$(RDFReport.NAME)
   REDIM RULES(10) AS SlctFileRULE, UFLD(10)
   REDIM SELSET(1), RELATE(1), RFLBL$(1), DFSEP(1), RFLTH(1), DFNUM(1)
   RULES(0).LOID = 1: RULES(0).HIID = 32767
   RULES(0).HIBY = 9999: RULES(0).HIBMD = 9999
   RULES(0).HIDY = 9999: RULES(0).HIDMD = 9999
   RULES(0).BDST = " ": RULES(0).DDST = " "
   RULES(0).FIDST = " ": RULES(0).MIDST = " "
   LSET RULES(0).GVNM = " ": LSET RULES(0).SRNM = " "
   LSET RULES(0).BPLC = " ": LSET RULES(0).DPLC = " "
   FOR I = 1 TO 10: RULES(I) = RULES(0): NEXT I
   SXK$ = " ": SNDX$ = "    ": SNDXCHK$ = SNDX$
   KEYWORD$ = SPACE$(30): SRCHWORD$ = SPACE$(30)
   SXTB$ = "01230120022455012623010202"
   CALL SelectSET("DEFAULT")

40 ' Format Display
   CALL WinCLR
   CALL FmtFIND("LISTS005"): FMT005 = CURFMT
   CALL FmtFIND("LISTS004"): FMT004 = CURFMT
   CALL FmtFIND("LISTS003"): FMT003 = CURFMT
   CALL FmtFIND("LISTS002"): FMT002 = CURFMT
   CALL FmtFIND("LISTS901"): FMT901 = CURFMT
   CALL FmtFIND("LISTS001"): FMT001 = CURFMT: WIN.OFMT = FMT001
   CALL WinDisplayLABELS(FMT001)

45 ' Refresh Option DISPLAY
   RRPT$ = "LIST"
   CALL WinPREP(2)
   OptTABLE.CL = 0
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
    CALL FAMOpen("READQ"): CALL FamCLOSE
    CURFMT = FMT001: FMT = FMTTAB(CURFMT)
    IF OKAY THEN
       CALL FmtFindFLD("NMCT")
       LSET UTXT$(FFLD.UTXT) = STR$(FF1Hdr.MAXID)
    END IF
    CALL WinDisplayDATA(FMT001)
    CALL SelectSET("STATS")

200 ' Get Processing Option
    CALL PutMSG("")
    IF FUN$ <> "" THEN CALL OptHILITE(FUN$, FUN$, 1)
    CALL GetKEY
    IF LEN(A$) < 2 THEN SOUND BP, DUR: FUN = 0: GOTO 200:  ELSE FUN = A
    SELECT CASE FUN
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
       CASE "F3": CALL SelectSET("CREATE"): GOTO 45
       CASE "F4": GOTO 400
       CASE "F6": IF PGMRTRN$ = "FHSMENU" THEN 450
                  IF PGMRTRN$ = "FHSRPTS" THEN 210
       CASE "F7": CALL SelectFILE("ANY")
       CASE "F9": GOTO 210
    END SELECT
    IF ENV.SCRNCLR THEN GOTO 40 ELSE GOTO 200

210 ' Return to FAMMENU.BAS
    CALL FamCLOSE
    IF NOT SAVEWORK THEN
       CALL PutMSG("LISTM001")
       IF A$ <> SNGLKEY$(1) THEN 200
    END IF
    REDIM RFLTH(1), RFLBL$(1)
    CALL PgmPREP(PGMRTRN$)
    IF (NOT OKAY) OR (A = 27) THEN LSET PGMRTRN$ = "FHSMENU": GOTO 40
    CALL RptFREE("LIST")
    CHAIN PGMRTRN$

300 ' Change FILE Parameters
    IF A = 94 THEN LSET PGMRTRN$ = "FHSFILE": GOTO 210
    CALL WinSWITCH(2)
    XSETUP = ENV.FSETUP
    CALL FDFSelect
    IF ENV.FSETUP <> XSETUP THEN
       CALL SelectSET("DEFAULT")
    END IF
    GOTO 45

310 ' Change PRINTER Parameters
    IF A = 95 THEN LSET PGMRTRN$ = "FHSPRTC": GOTO 210
    CALL WinSWITCH(2)
    CALL PDFSelect
    GOTO 45

400 ' Process Request to Update Options
    CALL PutMSG("LISTM002")
    IF A = 27 THEN 200
    IF A$ = "1" THEN
       LSET RRPT$ = "LIST"
       ELSE
       CALL PutMSG("LISTM003")
       IF A = 27 THEN GOTO 200
       IF (A$ = "1") THEN
          WTYPE = 4
          ELSE
          CALL PutMSG("LISTM004")
          IF A = 27 THEN 200
          WTYPE = VAL(A$)
          IF (WTYPE < 1) OR (WTYPE > 3) THEN WTYPE = 1
       END IF
       SELECT CASE WTYPE
          CASE 1: RRPT$ = "ASUM"
          CASE 2: RRPT$ = "DSUM"
          CASE 3: RRPT$ = "RSUM"
          CASE 4: RRPT$ = "ISUM"
       END SELECT
    END IF
    CALL WinSWITCH(2)
    IF RRPT$ <> "LIST" THEN
       CALL WinCLR
       CALL RptOptDISPLAY(RRPT$)
    END IF
    CALL RptOptUPDATE(RRPT$)
    GOTO 45

450 ' Print Reports
    IF (SlctCNT + SecSlctCNT) = 0 THEN
       CALL PutMSG("LISTM005")
       GOTO 200
    END IF
    IF (SlctCNT > 0) AND (SecSlctCNT > 0) THEN
       CALL PutMSG("LISTM006")
       IF A = 27 THEN 200
       xSEL = VAL(A$): IF (xSEL < 1) OR (xSEL > 3) THEN xSEL = 3
       ELSE
       xSEL = 3
    END IF
    CALL PutMSG("LISTM002")
    IF A = 27 THEN 200
    IF A$ = "2" THEN CALL XSUMReport ELSE CALL DETAILReport
    GOTO 40

REM $DYNAMIC
SUB BldDetDEF
SHARED DFNUM(), DFPP(), INDEXED
      RFCNT = 29
      REDIM DFNUM(RFCNT), DFPP(RFCNT)
      REDIM RFNUM(RFCNT), RFSEP(RFCNT), RFLTH(RFCNT), RFLBL$(RFCNT)
      CALL RptOptGET("MNGP", MINGAP)
      IF (NOT OKAY) OR (MINGAP = 0) THEN MINGAP = 2
      CALL RptOptGET("MXGP", MAXGAP)
      IF (NOT OKAY) OR (MAXGAP = 0) THEN MAXGAP = 5
      CALL RptOptGET("MINM", MINMLTH)
      IF (NOT OKAY) THEN MINMLTH = 35
      IF MINMLTH > 52 THEN MINMLTH = 52
      CALL RptOptGET("PNML", PLNMLTH)
      IF NOT OKAY THEN PLNMLTH = 22 - 8 * ENV.LPNSupport
      IF RPTOPTION.PRID THEN
         CALL RptOptGET("PPID", X): PrintPID = (X = 1)
      END IF
      CALL RptOptGET("SNDX", X): PrintSNDX = (X = 1)
      IF INDEXED THEN
         TS = ENV.SortOPTS AND 3: DS = (ENV.SortOPTS \ 4) AND 7
      '   FOR X = 1 TO SFCNT
      '       Y = SX(X) \ 256
      '       SELECT CASE Y
      '          CASE 1: X1 = X
      '          CASE 2: X2 = X
      '       END SELECT
      '   NEXT X
      '   IF X1 > 0 THEN
      '      RPTOPTION.NAMEFMT = (RPTOPTION.NAMEFMT AND 1) - 2 * (TS = 1)
      '      IF X2 > 0 THEN
      '         RPTOPTION.NAMEFMT = (RPTOPTION.NAMEFMT AND 2) - (X1 < X2)
      '      END IF
      '   END IF
      END IF

59500 ' Select Report Fields - Based Upon Report Options
      ' First Select Prefix Fields
      OKAY = 0
      PAGE.DFCNT = 0: RPTOPTION.TWOLINES = 0: PRID2 = 0
      IF (PAGE.BWIDTH = 0) THEN
         BWIDTH = PDFFORMS.FWIDTH - PDFFORMS.IMARGIN - PDFFORMS.OMARGIN
         ELSE
         BWIDTH = PAGE.BWIDTH
      END IF
      IF RPTOPTION.Relation THEN
         PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 3
      END IF
      IF NOT INDEXED THEN
         PSK = 9
         ELSE
         PSK = SX(1) \ 256
         PAGE.DFCNT = PAGE.DFCNT - ((PSK > 2) AND (PSK < 8))
         IF (TS > 1) AND ((PSK > 5) AND (PSK < 8)) THEN
            DFNUM(PAGE.DFCNT) = 24
            PAGE.DFCNT = PAGE.DFCNT + 1: PSNDX = PAGE.DFCNT
         END IF
         SELECT CASE PSK
            CASE 3: DFNUM(PAGE.DFCNT) = 12: BFX = PAGE.DFCNT
            CASE 4: DFNUM(PAGE.DFCNT) = 15: MFX = PAGE.DFCNT
            CASE 5: DFNUM(PAGE.DFCNT) = 19: DFX = PAGE.DFCNT
            CASE 6: DFNUM(PAGE.DFCNT) = 13: BFP = PAGE.DFCNT
            CASE 7: DFNUM(PAGE.DFCNT) = 20: DFP = PAGE.DFCNT
            CASE 8: DFNUM(PAGE.DFCNT) = 10: PSEX = PAGE.DFCNT
            CASE 9: INDEXED = 0
         END SELECT
      END IF
      IF RPTOPTION.PRID AND (RPTOPTION.FREEFORM OR (NOT INDEXED)) THEN
         PAGE.DFCNT = PAGE.DFCNT + 1: RFSEP(6) = MINGAP + 1
         DFNUM(PAGE.DFCNT) = 6: IDFX = PAGE.DFCNT
         IF ((RPTOPTION.SPOUSE AND 3) > 0) THEN
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 9
            ELSE
            RFLTH(9) = 0
         END IF
         IF PrintPID THEN
            DFNUM(PAGE.DFCNT + 1) = 7: DFNUM(PAGE.DFCNT + 2) = 8
            PAGE.DFCNT = PAGE.DFCNT + 2
         END IF
      END IF
      IF INDEXED AND (TS > 1) AND (PSK = 1) THEN
         PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 24
         PSNDX = PAGE.DFCNT
      END IF

      ' Select Data Fields
      RFSEP(DFNUM(1)) = 1: PAGE.PFCNT = PAGE.DFCNT
      IF (MINMLTH > 0) OR RPTOPTION.FREEFORM THEN
         PAGE.DFCNT = PAGE.DFCNT + 1
         DFNUM(PAGE.DFCNT) = 5: RFSEP(5) = MINGAP + 1
         RFLTH(5) = MINMLTH - (52 - MINMLTH) * RPTOPTION.FREEFORM
      END IF
      IF RPTOPTION.FREEFORM THEN 59510
      IF PrintSNDX AND (PSNDX = 0) THEN
         PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 24
      END IF
      IF RPTOPTION.PRID AND NOT ((PSK = 9)) THEN
         PAGE.DFCNT = PAGE.DFCNT + 1: IDFX = PAGE.DFCNT
         DFNUM(PAGE.DFCNT) = 6
         IF PrintPID THEN
            DFNUM(PAGE.DFCNT + 1) = 7
            DFNUM(PAGE.DFCNT + 2) = 8
            PAGE.DFCNT = PAGE.DFCNT + 2
         END IF
         IF ((RPTOPTION.SPOUSE AND 1) > 0) THEN
            PAGE.DFCNT = PAGE.DFCNT + 1
            DFNUM(PAGE.DFCNT) = 9
            ELSE
            RFLTH(9) = 0
         END IF
      END IF
      IF RPTOPTION.TIMELINE THEN
         DFNUM(PAGE.DFCNT + 1) = 22: DFNUM(PAGE.DFCNT + 2) = 23
         PAGE.DFCNT = PAGE.DFCNT + 2
         GOTO 59510
      END IF
      IF RPTOPTION.PRSEX AND (PSEX = 0) THEN
         IF SFX = 0 THEN
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 10
         END IF
      END IF
      CALL RptOptGET("#CH", X): PNUMChild = (X = 1)
      IF PNUMChild THEN
         PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 26
      END IF
      IF RPTOPTION.PRAGE THEN
         PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 11
      END IF
      IF RPTOPTION.DATES THEN
         IF BFX = 0 THEN
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 12: BFX = PAGE.DFCNT
         END IF
         IF RPTOPTION.SPOUSE THEN
            IF MFX = 0 THEN
               PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 15: MFX = PAGE.DFCNT
               IF RPTOPTION.PRAGE THEN
                  PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 18: AFX = PAGE.DFCNT
               END IF
            END IF
         END IF
         IF DFX = 0 THEN
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 19: DFX = PAGE.DFCNT
         END IF
      END IF
      IF RPTOPTION.PLACES THEN
         IF BFP = 0 THEN
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 13
            BFP = PAGE.DFCNT
         END IF
         IF RPTOPTION.SPOUSE THEN
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 16: MFP = PAGE.DFCNT
         END IF
         IF DFP = 0 THEN
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 20
            DFP = PAGE.DFCNT
         END IF
      END IF

59510 ' Create Report Definition Tables
      X = 1: RFNAME$ = "06010": GOSUB 59505     'COMMENTS
      X = 2: RFNAME$ = "07015": GOSUB 59505     'SOUNDEX
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
      X = 13: RFNAME$ = "01032": RFLTH(13) = PLNMLTH
              GOSUB 59505                       'BIRTH PLACE
      X = 14: RFNAME$ = "01034": RFLTH(14) = PLNMLTH
              GOSUB 59505    'BIRTH DATE & PLACE
      X = 15: RFNAME$ = "02020": GOSUB 59505    'MARRIAGE DATE
      X = 16: RFNAME$ = "02021": RFLTH(16) = PLNMLTH
              GOSUB 59505                       'MARRIAGE PLACE
      X = 17: RFNAME$ = "02024": RFLTH(17) = PLNMLTH
              GOSUB 59505                       'MARRIAGE DATE & PLACE
      X = 18: RFNAME$ = "02015": GOSUB 59505    'Last ANNIVERSARY
      X = 19: RFNAME$ = "01040": GOSUB 59505    'DEATH DATE
      X = 20: RFNAME$ = "01042": RFLTH(20) = PLNMLTH
              GOSUB 59505                       'DEATH PLACE
      X = 21: RFNAME$ = "01044": RFLTH(21) = PLNMLTH
              GOSUB 59505                       'DEATH DATE & PLACE
      X = 22: RFNAME$ = "07020": GOSUB 59505    'LIFE SPAN
      X = 23: RFNAME$ = "07021": GOSUB 59505    'LIFE LINE
      X = 24: RFNAME$ = "07015": GOSUB 59505    'SOUNDEX
      X = 25: RFNAME$ = "07003": GOSUB 59505    'RGL
      X = 26: RFNAME$ = "01027": GOSUB 59505    'Child COUNT
              IF RPTOPTION.FREEFORM AND FOUND THEN
                 RFLBL$(29) = MID$(RFLD$(FX), ALTH + LLTH + 1, HLTH)
              END IF
      X = 27: RFNAME$ = "03030": GOSUB 59505    'Phone
      X = 28: RFNAME$ = "03031": GOSUB 59505    'Address Line
      GOTO 59520

59505 ' Process Report Field
      CALL GetRFLD(RFNAME$, FX)
      IF RFSEP(X) = 0 THEN RFSEP(X) = MINGAP
      IF FOUND THEN
         RDFFLD = RFLDS(FX)
         IF RFLTH(X) = 0 THEN RFLTH(X) = ASC(RDFFLD.RLTH)
         ALTH = ASC(RDFFLD.ALTH): LLTH = ASC(RDFFLD.LLTH)
         HLTH = ASC(RDFFLD.HLTH)
         IF (X = 1) OR (X > 26) OR (RPTOPTION.FREEFORM AND (((X > 9) AND (X < 25)) OR (X = 26))) THEN
            RFLBL$(X) = MID$(RFLD$(FX), ALTH + 1, LLTH)
            ELSE
            IF X = 26 THEN
               RFLBL$(26) = MID$(RFLD$(FX), 1, ALTH)
               ELSE
               RFLBL$(X) = MID$(RFLD$(FX), ALTH + LLTH + 1, HLTH)
            END IF
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
      IF (NOT RPTOPTION.PLACES) OR (NOT RPTOPTION.DATES) THEN 59530
      IF (NOT RPTOPTION.TWOLINES) THEN
         IF (BFX > 2) AND (DFX > 2) AND (MFX > 2) THEN
            RPTOPTION.TWOLINES = -1
            DFNUM(BFX) = 14: DFNUM(BFP) = 0
            DFNUM(DFX) = 21: DFNUM(DFP) = 0
            IF RPTOPTION.SPOUSE THEN
               DFNUM(MFX) = 17: DFNUM(MFP) = 0
               IF RPTOPTION.PRAGE THEN DFNUM(AFX) = 0
            END IF
            GOTO 59520
         END IF
         GOTO 59530
      END IF
      IF RPTOPTION.PRID AND (NOT PRID2) AND (INDEXED) THEN
         PRID2 = -1
         DFNUM(IDFX) = 0: DFNUM(IDFX + 1) = 0: DFNUM(IDFX + 2) = 0
         IF ((RPTOPTION.SPOUSE AND 1) > 0) THEN DFNUM(IFDFX + 3) = 0
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

SUB BldSumDEF (WTYPE)
SHARED DFNUM(), DFPP()
SHARED DATECNT, DATERNG, PLACCNT, MARRCNT, RWIDTH, VRTCL
      IF WTYPE < 4 THEN
         AMAXLV = ASC(WFHDR.AMAXLV): DMAXLV = ASC(WFHDR.DMAXLV)
         RELSUMOPT = (AMAXLV > 0) AND (DMAXLV > 0)
      END IF
      IF RELSUMOPT THEN RFCNT = AMAXLV + 3 ELSE RFCNT = 33
      PAGE.DFCNT = 0
      REDIM DFNUM(RFCNT), DFPP(RFCNT)
      REDIM RFNUM(RFCNT), RFSEP(RFCNT), RFLTH(RFCNT), RFLBL$(RFCNT)
      CALL RptOptGET("MNGP", MINGAP)
      IF (NOT OKAY) OR (MINGAP = 0) THEN MINGAP = 2
      DFPP(0) = MINGAP
      CALL RptOptGET("MXGP", MAXGAP)
      IF (NOT OKAY) OR (MAXGAP = 0) THEN MAXGAP = 5
      IF RELSUMOPT THEN GOTO 58500
      IF WTYPE < 4 THEN
         IF DMAXLV = 0 THEN
            CALL RptOptGET("CNOP", X): CNTNOP = (X = 1)
         END IF
         CALL RptOptGET("CDUP", X): CNTDUP = (X = 1)
         CALL RptOptGET("CLIN", X): CNTCLN = (X = 1)
         ELSE
         CALL RptOptGET("SNDX", X): PRTSNDX = (X = 1)
      END IF
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
      IF WTYPE < 4 THEN             ' Relationship Summary Reports
         PAGE.DFCNT = 1: DFNUM(1) = 1
         IF RPTOPTION.Relation OR VRTCL THEN
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 2
         END IF
         IF CNTCLN THEN PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 4
         IF CNTNOP THEN PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 6
         IF CNTDUP THEN PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 5
         GOTO 58505
      END IF

      ' Select Prefix Fields for INDEXED Summary
      TS = ENV.SortOPTS AND 3: DS = (ENV.SortOPTS \ 4) AND 7
      PSK = SX(1) \ 256
      IF (2 < PSK) AND (PSK < 6) THEN
         SELECT CASE DS
            CASE 0: PFK1 = 26: PFK2 = 27: PFK3 = 28
            CASE 1: PFK1 = 27: PFK2 = 28: PFK3 = 26
            CASE 2: PFK1 = 28: PFK2 = 27: PFK3 = 26
            CASE 3: PFK1 = 26: PFK2 = 27
            CASE 4: PFK1 = 27: PFK2 = 28
            CASE 5: PFK1 = 26
            CASE 6: PFK1 = 27
            CASE 7: PFK1 = 28
         END SELECT
         ELSE
         IF PSK > 7 THEN
            PFK1 = 33
            ELSE
            IF TS > 1 THEN PFK1 = 24 ELSE IF PRTSNDX THEN PFK3 = 24
            SELECT CASE PSK
               CASE 1: PFK2 = 29
               CASE 2: PFK2 = 30
               CASE 6: PFK2 = 31
               CASE 7: PFK2 = 32
            END SELECT
            IF PFK2 > 0 THEN RFLTH(PFK2) = SX(1) AND 127
         END IF
      END IF
      IF PFK1 > 0 THEN
         PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = PFK1
      END IF
      IF PFK2 > 0 THEN
         PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = PFK2
      END IF
      IF PFK3 > 0 THEN
         PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = PFK3
      END IF

58505 ' Select Counter Fields
      IF NOT (PSK = 8) THEN
         PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 7
         PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 8
      END IF
      PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 3
      IF NOT (PSK = 3) THEN
         IF DATECNT THEN PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 9
         IF DATERNG THEN
            DFNUM(PAGE.DFCNT + 1) = 11: DFNUM(PAGE.DFCNT + 2) = 12
            PAGE.DFCNT = PAGE.DFCNT + 2
         END IF
      END IF
      IF NOT (PSK = 6) THEN
         IF PLACCNT THEN
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 10
         END IF
      END IF
      IF MARRCNT THEN
         IF NOT (PSK = 4) THEN
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 13
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 14
            IF DATECNT THEN
               PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 15
            END IF
            IF DATERNG THEN
               PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 17
               PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 18
            END IF
         END IF
         IF PLACCNT THEN
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 16
         END IF
      END IF
      IF NOT (PSK = 5) THEN
         IF DATECNT THEN
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 19
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 20
         END IF
         IF DATERNG THEN
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 22
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 23
         END IF
      END IF
      IF NOT (PSK = 7) THEN
         IF PLACCNT THEN
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 21
         END IF
      END IF

58510 ' Create Report Definition Tables
      X = 1                                     'AGL or DGL
      IF WTYPE < 4 THEN
         IF (DMAXLV > 0) AND (AMAXLV = 0) THEN
            RFNAME$ = "07002"
            ELSE
            RFNAME$ = "07001"
         END IF
         GOSUB 58515
      END IF
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
      X = 24: RFNAME$ = "07015": GOSUB 58515    'SOUNDEX
      X = 25: RFNAME$ = "08005": GOSUB 58515    'SURNMCNT
      X = 26: RFNAME$ = "08009": GOSUB 58515    'YEAR
      X = 27: RFNAME$ = "08010": GOSUB 58515    'MONTH
      X = 28: RFNAME$ = "08011": GOSUB 58515    'DAY
      X = 29: RFNAME$ = "01022": GOSUB 58515    'Surname
      X = 30: RFNAME$ = "01021": GOSUB 58515    'GivenName
      X = 31: RFNAME$ = "01032": GOSUB 58515    'BirthPlace
      X = 32: RFNAME$ = "01042": GOSUB 58515    'DeathPlace
      X = 33: RFNAME$ = "01015": GOSUB 58515    'Sex Code
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
      DFNUM(1) = 1: RFSEP(1) = 1: RFLTH(1) = 8: DFPP(1) = 1
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
SHARED DFNUM(), DFPP(), VRTCL, INDEXED
      AMAXLV = ASC(WFHDR.AMAXLV): DMAXLV = ASC(WFHDR.DMAXLV)
      RELSUMOPT = (AMAXLV > 0) AND (DMAXLV > 0)
     ' Build Summary Report Title Line
      TL$ = SPACE$(PAGE.BWIDTH): Y = 0
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
         HTFLINE$(HTFLMAX) = "T" + PRT.SECTION + "0" + RTRIM$(TL$)
         IF VRTCL THEN
            HTFLMAX = HTFLMAX + 1
            REDIM PRESERVE HTFLINE$(HTFLMAX)
            HTFLINE$(HTFLMAX) = "T00" + RTRIM$(TL2$)
         END IF
      END IF
      OKAY = -1
END SUB

SUB ChkRULE (XRULE AS SlctFileRULE, R)
SHARED UFLD(), RULES() AS SlctFileRULE

565 ' Check for Modified SubRules
    UFLD = ((XRULE.LOID <> RULES(0).LOID) OR (XRULE.HIID <> RULES(0).HIID))
    UFLD = UFLD + 2 * (XRULE.SEX <> RULES(0).SEX)
    UFLD = UFLD + 4 * (XRULE.SRNM <> RULES(0).SRNM)
    UFLD = UFLD + 8 * (XRULE.GVNM <> RULES(0).GVNM)
    UFLD = UFLD + 16 * ((XRULE.LOBY <> RULES(0).LOBY) OR (XRULE.LOBMD <> RULES(0).LOBMD) OR (XRULE.HIBY <> RULES(0).HIBY) OR (XRULE.HIBMD <> RULES(0).HIBMD))
    UFLD = UFLD + 32 * (XRULE.BPLC <> RULES(0).BPLC)
    UFLD = UFLD + 64 * ((XRULE.LODY <> RULES(0).LODY) OR (XRULE.LODMD <> RULES(0).LODMD) OR (XRULE.HIDY <> RULES(0).HIDY) OR (XRULE.HIDMD <> RULES(0).HIDMD))
    UFLD = UFLD + 128 * (XRULE.DPLC <> RULES(0).DPLC)
    UFLD = UFLD + 256 * (XRULE.YORN <> RULES(0).YORN)
    UFLD = UFLD + 512 * ((XRULE.BDST <> RULES(0).BDST) OR (XRULE.DDST <> RULES(0).DDST))
    UFLD = UFLD + 1024 * ((XRULE.FIDST <> RULES(0).FIDST) OR (XRULE.MIDST <> RULES(0).MIDST))
    UFLD(R) = -UFLD
END SUB

SUB DETAILReport
SHARED DFNUM(), DFPP()
SHARED RELID, TMBY, TMEY
SHARED xSEL, SELSET(), DFSEP()
SHARED INDEXED
SHARED SPRESEQ, MLIST(), SPCNT, SPRNO, INDID, ADRSPTR, ADRSSPID

     ' Produce LIST of Selected Individuals
     CALL RptOPEN("LIST"): IF NOT OKAY THEN 1040
1010 CALL PutMSG("LISTM007")
     IF A = 27 THEN 1040
     RPTOPTION.FREEFORM = (A$ = "2"): RPTOPTION.TIMELINE = (A$ = "3")
1020 IF RPTOPTION.TIMELINE THEN
        CALL FmtFIND("XXXXS030")
        CALL MsgFORMAT: IF A = 27 THEN 1010
        CALL FmtFindFLD("TLBY"): IF FOUND THEN TMBY = VAL(UTXT$(FFLD.UTXT))
        CALL FmtFindFLD("TLEY"): IF FOUND THEN TMEY = VAL(UTXT$(FFLD.UTXT))
        IF (TMEY - TMBY) < 1 THEN CALL ErrBEEP(0): GOTO 1020
     END IF

     CALL PutMSG("LISTM008")
     IF ASC(A$) = 27 THEN 1040
     A = VAL(A$): IF A < 1 OR A > 2 THEN A = 2
     INDEXED = (A = 1)
     IF NOT INDEXED THEN
        REDIM SX(1)
        SFCNT = 1: SX(1) = 9 * 256: NDXRECS = FF1Hdr.MAXID
        RPTOPTION.PRID = -1: SXK = 0
        ELSE
        CALL INDEXFile("READ")
        IF NOT OKAY THEN 1040
        NDXRECS = NDXHdr.NDXRECS
        TS = ENV.SortOPTS AND 3: DS = (ENV.SortOPTS \ 4) AND 7
        SXK = SX(1) \ 256
        IF ((SXK > 2) AND (SXK < 6)) OR (SXK > 7) THEN SXK = 1
     END IF

     IF RPTOPTION.Relation THEN
        CALL RptOptGET("EXAD", X): ExtAdoptRULE = (X = 1)
        CALL RWRKOpen("READ", 0, 0): RPTOPTION.Relation = OKAY
        IF OKAY THEN
           REDIM RELATE(FF1Hdr.MAXID)
           CALL PutMSG("XXXXM064")
           CALL RWRKGet("BASE", RC, PTR&)
           WHILE (0 < RC) AND (RC < 9)
              IF WFRELREC.ID > 0 THEN
                 X = RELATE(WFRELREC.ID)
                 IF (X = 0) OR (WFGLREC.DGL = 0) THEN
                    X = 256 * (-WFGLREC.AGL) + WFGLREC.DGL
                    IF ExtAdoptRULE THEN X = X - 128 * ((WFRELREC.CSTAT AND 5) > 0)
                    RELATE(WFRELREC.ID) = X
                 END IF
              END IF
              CALL RWRKGet("WRKREC", RC, PTR&)
           WEND
           CALL FamCLOSE
        END IF
     END IF
     CALL FmtFIND("LISTS009")
     CALL FmtFindFLD("TITL")
     IF FOUND THEN
        IF LEN(STITLE$) > 0 THEN
           LSET UTXT$(FFLD.UTXT) = STITLE$
           ELSE
           LSET UTXT$(FFLD.UTXT) = ENV.RTITLE
        END IF
        CALL MsgFORMAT
        STITLE$ = RTRIM$(UTXT$(FFLD.UTXT))
        IF (LEN(STITLE$) > 0) THEN
           LSET ENV.RTITLE = STITLE$
        END IF
     END IF
     BD$ = SPACE$(12): DD$ = SPACE$(12): AGE$ = "   "
     BPL$ = SPACE$(41): MPL$ = BPL$: DPL$ = BPL$: CL$ = SPACE$(255)
     D1D$ = BD$: D2D$ = BD$: D1PL$ = BPL$: D2PL$ = BPL$: AGE2$ = AGE$
     NUMED$ = SPACE$(6): SNDX$ = "    ": SNDXText$ = SPACE$(30)
     SURNM$ = SPACE$(20): ADRDATA$ = SPACE$(135)
     CALL FAMOpen("READ")
     AFMT = FF2HDR.REVLVL
     IF OKAY THEN 1200

1040 ' Bad Return
     CLOSE : ENV.FFOPEN = 0: EXIT SUB

1200 ' Setup Report Tables
     CALL BldDetDEF: IF NOT OKAY THEN 1040
     PRDAGE = (RPTOPTION.PRAGE = 2)
     PRID2 = ((RPTOPTION.PRID AND 2) > 0)
     RPTOPTION.PRID = (RPTOPTION.PRID <> 0)
     CALL RptOptGET("#CH", X): PNUMChild = (X = 1)
     CALL RptOptGET("LRES", X): PLastRES = (X = 1)
     CALL RptOptGET("EURA", X): EuroAdr = (X = 1)
     CALL RptOptGET("PNML", PLNMLTH)
     IF PLNMLTH = 0 THEN PLNMLTH = 22 - 8 * ENV.LPNSupport
     CALL PrintOPEN: IF NOT OKAY THEN 1040
     PLINE$ = SPACE$(255): PL2$ = PLINE$
     RGL$ = SPACE$(RFLTH(25)): LINEAGE$ = SPACE$(RFLTH(2))
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
     CINDENT = PAGE.CINDENT: IF RPTOPTION.FREEFORM THEN PAGE.CINDENT = 0

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

1400 ' Begin with First Record
     A = 0: PAGE.CSID = 0: xNDX = 0

1415 ' Process next entry on chart
     IF NOT (xNDX < NDXRECS) THEN 1490
     xNDX = xNDX + 1
     IF INDEXED THEN INDID = NDXID(xNDX) ELSE INDID = xNDX
     IF (SELSET(INDID) AND xSEL) = 0 THEN 1415

1450 ' Process Individual record
     CALL FF1GetRec(INDID)
     IF PLastRES THEN CALL GetLastRES
     IF RPTOPTION.Relation THEN
        IF RELATE(INDID) <> 0 THEN
           X = RELATE(INDID): Y = X AND 128
           AGL = -(X \ 256): DGL = X AND 127
           RGL = AGL + DGL
           LSET XRELATION$ = Relation$("M", FF1REC.SEX, AGL, RGL, RRELCD$())
           IF ExtAdoptRULE THEN
              IF (X AND 128) > 0 THEN LSET XRELATION$ = "*" + XRELATION$
           END IF
           ELSE
           AGL = 0: DGL = 0: RGL = 0: LSET XRELATION$ = ""
        END IF
     END IF
     IF ((RPTOPTION.BLNKSEP) AND (PAGE.LINE > 0)) THEN CALL WriteLINE
     PAGE.RPAGE = 0: PAGE.RLINE = 0
     PAGE.XID = INDID: PAGE.XTYPE = 1
     RREFRec.AGL = CHR$(ABS(AGL)): RREFRec.DGL = CHR$(DGL): CALL RREFAdd
     COMPTR = FF1REC.COM
     CALL GetBDEvents(BD$, BPL$, DD$, DPL$, AGE$)
     BDY = FF1REC.BY: DDY = FF1REC.DY
     IF PNUMChild THEN GOSUB 1510
     IF RPTOPTION.SPOUSE THEN
        PAGE.CSID = 0: MLOC = 0: MCOMPTR = 0
        SDY = 0: SDMD = 0: M2Y = 0: M2MD = 0
        SPRNO = FF1REC.SPOUSE: S1DY = FF1REC.DY: S1DMD = FF1REC.DMD
        IF ASPOUSE AND (SPRNO <> 0) THEN
           GOSUB 1530
           ELSE
           IF ALLSPOUSE THEN CALL BldMLIST
        END IF
     END IF
     GOSUB 1800   ' Process DATA Fields
     IF A = 27 THEN 1490

     IF NOT ALLSPOUSE THEN 1470       ' Didn't Request ALL Marriages

     ' Process Marriage Records
     WHILE SPRNO <> 0
        PAGE.RPAGE = 0: PAGE.RLINE = 0
        PAGE.XTYPE = 3: GOSUB 1530
        PAGE.XID = PAGE.CSID: CALL RREFAdd
        BDY = FF1REC.BY: DDY = FF1REC.DY
        LSET BD$ = D1D$: LSET BPL$ = D1PL$
        LSET DD$ = D2D$: LSET DPL$ = D2PL$: LSET AGE$ = AGE2$
        GOSUB 1800
        IF A = 27 THEN SPRNO = 0
        SDY = 0: SDMD = 0
     WEND

1470 ' Finish Line and Go Get Another WF Record
     CALL PRINTLINE
     IF A = 27 THEN 1490
     GOTO 1415

1490 ' END OF REPORT
     CALL RptCLOSE: IF PRT.AGAIN THEN 1400
     CALL FamCLOSE
     PAGE.CINDENT = CINDENT
     EXIT SUB

1510 ' Count Children
     X = FF1REC.OLDCH: NUMCH = 0
     WHILE X <> 0
        NUMCH = NUMCH + 1
        CALL FF1GetRec(X)
        IF FF1REC.FID = INDID THEN
           X = FF1REC.FCH
           ELSE
           IF FF1REC.MID = INDID THEN
              X = FF1REC.MCH
              ELSE
              X = 0
           END IF
        END IF
     WEND
     CALL FF1GetRec(INDID)
     RETURN

1530 ' Get Information for SPRNO
     IF (SPRNO <> 0) OR (PAGE.CSID > 0) THEN
        IF SPRNO <> 0 THEN CALL FF3GetRec(SPRNO)
        MCOMPTR = FF3SPOUSE.COM: MLOC = FF3SPOUSE.MLOC
        M1Y = FF3SPOUSE.D1Y: M1MD = FF3SPOUSE.D1MD
        M2Y = FF3SPOUSE.D2Y: M2MD = FF3SPOUSE.D2MD
        IF FF3SPOUSE.SP1ID = INDID THEN
           PAGE.CSID = FF3SPOUSE.SP2ID
           ELSE
           PAGE.CSID = FF3SPOUSE.SP1ID
        END IF
        IF ALLSPOUSE THEN CALL GetNextMARRIAGE ELSE SPRNO = 0
        IF PAGE.CSID > 0 THEN
           CALL FF1GetRec(PAGE.CSID)
           CALL GetBDEvents(D1D$, D1PL$, D2D$, D2PL$, AGE2$)
           IF ALLSPOUSE THEN COMPTR = FF1REC.COM
           SDY = FF1REC.DY: SDMD = FF1REC.DMD
           IF PAGE.XTYPE = 1 THEN
              CALL FF1GetRec(INDID)
              ELSE
              IF RPTOPTION.Relation THEN LSET XRELATION$ = " "
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

1545 ' Determine Last Anniversary
     IF ((M2Y > 0 OR M2MD > 0) OR (S1DY = 0 AND SDY = 0)) THEN
        LSET ANVSRC$ = " "
        ELSE
        M2Y = SDY: M2MD = SDMD: LSET ANVSRC$ = "*"
        IF ((S1DY > 0) AND ((SDY = 0) OR ((S1DY < SDY) OR (S1DY = SDY AND S1DMD < SDMD)))) THEN M2Y = S1DY: M2MD = S1DMD
     END IF
     RSET ANNV$ = DATEDif$(M1MD, M1Y, M2MD, M2Y)
     RETURN

1800 ' PRINT Prefix Fields (Those Before NAME on Data Line)
     IF (PAGE.XTYPE = 1) OR (PAGE.CSID > 0) THEN
        LSET NFMT$ = FMTNAME$(0)
        ELSE
        LSET NFMT$ = "(" + Unknown$ + ")"
     END IF
     COMBGN = 0
     FX = 0
     WHILE FX < PAGE.PFCNT
        FX = FX + 1: DFN = DFNUM(FX): X = DFPP(FX)
        IF (DFN > 0) AND (X > 0) THEN
           FLTH = RFLTH(DFN)
           PAGE.LOFST = X
           GOSUB 1810
        END IF
     WEND
                              
1805 ' Process DATA Items After NAME Field on DATA Line
     IF RPTOPTION.FREEFORM THEN GOSUB 2800 ELSE GOSUB 2500
     IF NOT RPTOPTION.FREEFORM THEN
        CALL PRINTLINE
        IF RPTOPTION.TWOLINES THEN
           LSET PLINE$ = PL2$
           CALL PRINTLINE: LSET PL2$ = " "
        END IF
     END IF

     ' Print Comments - Individual and Marriage as appropriate
     IF RPTOPTION.COMMENTS THEN
        COMBGN = 0: PAGE.LWIDTH = PAGE.BWIDTH - PAGE.CINDENT
        IF (PAGE.XTYPE = 3) AND MCOMPTR THEN CALL ComPRINT(MCOMPTR, COMBGN)
        IF COMPTR THEN CALL ComPRINT(COMPTR, COMBGN)
        IF (PAGE.XTYPE = 1) AND MCOMPTR THEN CALL ComPRINT(MCOMPTR, COMBGN)
     END IF
     CALL PRINTLINE
     'IF PLastRES AND (PAGE.CSID = 0) THEN
     IF (PLastRES AND (ADRSPTR <> 0)) THEN
        IF (ALLSPOUSE AND (ADRSSPID = PAGE.CSID)) OR (NOT ALLSPOUSE) THEN
           GOSUB 1900
        END IF
     END IF
     RETURN

1810 ' Process Data Fields by DFN
     IF DFN < 10 THEN
        SELECT CASE DFN
           CASE 3: GOTO 1823
           CASE 5: GOTO 1825
           CASE 6: GOTO 1826
           CASE 7: GOTO 1826
           CASE 8: GOTO 1826
           CASE 9: GOTO 1826
        END SELECT
        ELSE
        IF DFN < 21 THEN
           SELECT CASE DFN
              CASE 10: GOTO 2501
              CASE 11: GOTO 2502
              CASE 12: GOTO 2510
              CASE 13: GOTO 2515
              CASE 14: GOTO 2510
              CASE 15: GOTO 2520
              CASE 16: GOTO 2525
              CASE 17: GOTO 2520
              CASE 18: GOTO 2521
              CASE 19: GOTO 2530
              CASE 20: GOTO 2535
           END SELECT
           ELSE
           SELECT CASE DFN
              CASE 21: GOTO 2530
              CASE 22: GOTO 2600
              CASE 23: GOTO 2610
              CASE 24: GOTO 1822
              CASE 25: GOTO 1821
              CASE 26: GOTO 1824
           END SELECT
        END IF
     END IF
     RETURN

1820 ' Routines for Processing PREFIX Data Fields
1821 ' AGL, DGL or RGL
     RSET RGL$ = STR$(RGL)
     MID$(PLINE$, PAGE.LOFST, FLTH) = RGL$
     GOTO 1829
1822 ' Soundex
     IF (SXK = 0) OR (PAGE.XTYPE = 1) THEN
        SELECT CASE SXK
           CASE 0: LSET SNDXText$ = SURNM$
           CASE 1: LSET SNDXText$ = SURNM$
           CASE 2: LSET SNDXText$ = FF1REC.GIVEN
           CASE 6: LSET SNDXText$ = BPL$
           CASE 7: LSET SNDXText$ = DPL$
        END SELECT
        CALL ComputeSNDX(SNDXText$, SNDX$)
        MID$(PLINE$, PAGE.LOFST, FLTH) = SNDX$
     END IF
     GOTO 1829
1824 ' Count Children
     IF NUMCH > 0 THEN MID$(PLINE$, PAGE.LOFST, FLTH) = SFILL$(NUMCH, FLTH)
     GOTO 1829
1823 ' Relationship
     MID$(PLINE$, PAGE.LOFST, FLTH) = XRELATION$
     LSET XRELATION$ = " "
     GOTO 1829
1825 ' Name
     NMLTH = FLTH
     IF PAGE.XTYPE = 3 THEN
        IF NOT RPTOPTION.BLNKSEP THEN PAGE.LOFST = PAGE.LOFST + CINDENT
        IF LEN(SPX$) > 0 THEN
           MID$(PLINE$, PAGE.LOFST, 255) = SPX$
           X = INSTR(PAGE.LOFST, PLINE$, "  ") + 1
        END IF
        IF NOT RPTOPTION.FREEFORM THEN NMLTH = FLTH - (X - PAGE.LOFST)
        PAGE.LOFST = X
     END IF
     MID$(PLINE$, PAGE.LOFST, NMLTH) = NFMT$
     PAGE.COFST = PAGE.LOFST
     IF PAGE.XTYPE = 1 THEN
        PAGE.COFST = PAGE.COFST - CINDENT * RPTOPTION.FREEFORM
     END IF
     IF PRID2 THEN GOSUB 1828
     IF RPTOPTION.FREEFORM THEN
        PAGE.LOFST = INSTR(PAGE.LOFST, PLINE$, "   ")
     END IF
     GOTO 1829
1826 ' ID Numbers - First Data Line
     SELECT CASE DFN
        CASE 6: IF (ASPOUSE OR (PAGE.CSID = 0) OR (((SX(1) \ 256) <> 9) AND (NOT RPTOPTION.FREEFORM))) THEN MID$(PLINE$, PAGE.LOFST, 5) = SFILL$(FF1REC.RID, 5)
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

1900 ' Format ADDRESS Line
     IF (NOT RPTOPTION.BLNKSEP) AND (NOT RPTOPTION.FREEFORM) THEN PAGE.LOFST = PAGE.LOFST + CINDENT
     LSET ADRDATA$ = "": X = 1
     IF AFMT = 0 THEN
        IF (FF2SREC.FON1 <> 0) OR (FF2SREC.FON2 <> 0) OR (FF2SREC.FON3 <> 0) THEN
           MID$(ADRDATA$, 1, 14) = ZFILL$(FF2SREC.FON1, 3) + "-" + ZFILL$(FF2SREC.FON2, 3) + "-" + ZFILL$(FF2SREC.FON3, 4) + "  "
        END IF
        MID$(ADRDATA$, 15, 30) = LTRIM$(RTRIM$(FF2SREC.LINE1))
        MID$(ADRDATA$, 45, 30) = LTRIM$(RTRIM$(FF2SREC.LINE2))
        MID$(ADRDATA$, 75, 15) = LTRIM$(RTRIM$(FF2SREC.CITY))
        MID$(ADRDATA$, 95, 4) = LTRIM$(RTRIM$(FF2SREC.STATE))
        MID$(ADRDATA$, 107, 5) = LTRIM$(RTRIM$(FF2SREC.ZIP))
        ELSE
        IF (ASC(FF2LREC.PHONE) <> 0) AND (LEFT$(FF2LREC.PHONE, 13) <> "(000)000-0000") THEN
           MID$(ADRDATA$, 1, 14) = LTRIM$(RTRIM$(FF2LREC.PHONE))
        END IF
        MID$(ADRDATA$, 15, 30) = LTRIM$(RTRIM$(FF2LREC.LINE1))
        MID$(ADRDATA$, 45, 30) = LTRIM$(RTRIM$(FF2LREC.LINE2))
        MID$(ADRDATA$, 75, 20) = LTRIM$(RTRIM$(FF2LREC.CITY))
        MID$(ADRDATA$, 95, 12) = LTRIM$(RTRIM$(FF2LREC.STATE))
        MID$(ADRDATA$, 107, 12) = LTRIM$(RTRIM$(FF2LREC.ZIP))
        MID$(ADRDATA$, 119, 16) = LTRIM$(RTRIM$(FF2LREC.COUNTRY))
     END IF
     IF (FF2PFX.D2Y > 0) OR (FF2PFX.D2MD > 0) THEN
        MID$(CL$, X, 1) = "*": X = X + 1
     END IF
     IF LEFT$(ADRDATA$, 1) <> " " THEN
        MID$(CL$, X, 255) = RFLBL$(27): X = INSTR(CL$, "  ")
        MID$(CL$, X, 1) = ":": X = X + 1
        MID$(CL$, X, 255) = MID$(ADRDATA$, 1, 14): X = INSTR(X, CL$, "  ") + 2
     END IF
     MID$(CL$, X, 255) = RFLBL$(28): X = INSTR(X, CL$, "  ")
     MID$(CL$, X, 1) = ":": X = X + 1
     IF MID$(ADRDATA$, 15, 1) <> " " THEN
        MID$(CL$, X, 255) = MID$(ADRDATA$, 15, 30): X = INSTR(X, CL$, "  ")
        MID$(CL$, X, 1) = ",": X = X + 2
     END IF
     IF MID$(ADRDATA$, 45, 1) <> " " THEN
        MID$(CL$, X, 255) = MID$(ADRDATA$, 45, 30): X = INSTR(X, CL$, "  ")
        MID$(CL$, X, 1) = ",": X = X + 2
     END IF
     IF EuroAdr THEN
        IF MID$(ADRDATA$, 107, 1) <> " " THEN
           MID$(CL$, X, 255) = MID$(ADRDATA$, 107, 12): X = INSTR(X, CL$, "  ")
           X = X + 1
        END IF
     END IF
     IF MID$(ADRDATA$, 75, 1) <> " " THEN
        MID$(CL$, X, 255) = MID$(ADRDATA$, 75, 20): X = INSTR(X, CL$, "  ")
        MID$(CL$, X, 1) = ",": X = X + 2
     END IF
     IF MID$(ADRDATA$, 95, 1) <> " " THEN
        MID$(CL$, X, 255) = MID$(ADRDATA$, 95, 12): X = INSTR(X, CL$, "  ") + 1
        IF EuroAdr THEN MID$(CL$, X - 1, 1) = ",": X = X + 1
     END IF
     IF NOT EuroAdr THEN
        IF MID$(ADRDATA$, 107, 1) <> " " THEN
           MID$(CL$, X, 255) = MID$(ADRDATA$, 107, 12): X = INSTR(X, CL$, "  ")
           MID$(CL$, X, 1) = ",": X = X + 2
        END IF
     END IF
     IF MID$(ADRDATA$, 119, 1) <> " " THEN
        MID$(CL$, X, 255) = MID$(ADRDATA$, 119, 16): X = INSTR(X, CL$, "  ")
        MID$(CL$, X, 1) = ",": X = X + 2
     END IF
     IF X > 2 THEN IF MID$(CL$, X - 2, 1) = "," THEN MID$(CL$, X - 2, 1) = " "
     CALL PutData: CALL PRINTLINE
     IF RPTOPTION.COMMENTS THEN
        COMBGN = 0: PAGE.LWIDTH = PAGE.BWIDTH - PAGE.CINDENT
        IF (FF2PFX.COM <> 0) THEN CALL ComPRINT(FF2PFX.COM, COMBGN)
        CALL PRINTLINE
     END IF
     RETURN

2500 ' Print Information for Fixed Format Report - Fields After Prefix
     IF (PAGE.XTYPE = 1) OR (PAGE.CSID > 0) THEN
        LSET NFMT$ = FMTNAME$(0)
        ELSE
        LSET NFMT$ = "(" + Unknown$ + ")"
     END IF
     FX = PAGE.PFCNT
     WHILE FX < PAGE.DFCNT
        FX = FX + 1: DFN = DFNUM(FX): Y = DFPP(FX)
        IF (DFN > 0) AND (Y > 0) THEN
           FLTH = RFLTH(DFN): PAGE.LOFST = Y
           GOSUB 1810
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
           MID$(PLINE$, PAGE.LOFST, 10) = ANNV$
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
     IF (TMTYPE = 1) AND (TMX < TMLTH) AND (TMDYR < 9999) THEN MID$(TMLINE$, TMX + 1, TMLTH) = TMLIFE$
     IF (TMTYPE = 4) AND (TMX < TMLTH) THEN MID$(TMLINE$, TMX + 1, TMLTH) = TMFILL$
     RETURN

2800 ' Print Information for Free Format Report
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
           CALL PutData
        END IF
     END IF

     ' AGE - if Still Living
     IF (PAGE.XTYPE = 1) OR (PAGE.CSID > 0) THEN
        IF (RPTOPTION.PRAGE) AND (NOT PRDAGE) AND (DDY = 0) THEN
           LSET CL$ = ", " + RFLBL$(11) + " " + LTRIM$(AGE$)
           CALL PutData
        END IF
     END IF

     ' Birth Date/Place
     IF (PAGE.XTYPE = 1) OR (PAGE.CSID > 0) THEN
        IF RPTOPTION.PLACES THEN PLTH = LEN(RTRIM$(BPL$))
        IF (BD$ <> SPACE$(12)) OR (PLTH > 1) THEN
           LSET CL$ = ", " + RFLBL$(12): X = 1
           IF (BD$ <> SPACE$(12)) THEN
              X = INSTR(X, CL$, "   ")
              MID$(CL$, X + 1, 12) = LTRIM$(BD$)
           END IF
           IF PLTH > 0 THEN
              X = INSTR(X, CL$, "   ")
              MID$(CL$, X + 1, 255) = RFLBL$(13) + " " + BPL$
           END IF
           CALL PutData
        END IF
     END IF

     ' Marriage DATE/Place/Anniv
     IF (PAGE.XTYPE = 3) OR (PAGE.CSID > 0) THEN
        IF RPTOPTION.PLACES THEN PLTH = LEN(RTRIM$(MPL$))
        IF (M1Y <> 0) OR (M1MD <> 0) OR (PLTH > 1) THEN
           LSET CL$ = ", " + RFLBL$(15): X = 1
           IF (M1Y > 0) OR (M1MD > 0) THEN
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
           CALL PutData
        END IF
     END IF

     ' Death DATE/Place/Age
     IF (PAGE.XTYPE = 1) OR (PAGE.CSID > 0) THEN
        IF RPTOPTION.PLACES THEN PLTH = LEN(RTRIM$(DPL$))
        IF (DD$ <> SPACE$(12)) OR (PLTH > 1) THEN
           LSET CL$ = ", " + RFLBL$(19): X = 1
           IF (DD$ <> SPACE$(12)) THEN
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
           CALL PutData
        END IF
     END IF

     ' Number of Children
     IF PNUMChild AND (NUMCH > 0) THEN
        LSET CL$ = ";" + STR$(NUMCH)
        Z = INSTR(CL$, "  ") + 1
        IF NUMCH = 1 THEN
           MID$(CL$, Z, 255) = RFLBL$(26)
           ELSE
           MID$(CL$, Z, 255) = RFLBL$(29)
        END IF
        CALL PutData
     END IF
     RETURN

END SUB

SUB GetLastRES
SHARED ADRSPTR, ADRSSPID
    ' Get Residence
    D1Y = 0: D1MD = 0: ADRSPTR = FF1REC.ADRS: ADRSSPID = 0
    IF ADRSPTR <> 0 THEN
       CALL FF2GETREC(ADRSPTR)
       D1Y = FF2PFX.D1Y: D1MD = FF2PFX.D1MD
    END IF
    MPTR = FF1REC.SPOUSE: MADRS = 0
    WHILE (MPTR <> 0) AND (MADRS = 0)
       CALL FF3GetRec(MPTR): MADRS = FF3SPOUSE.LOC
       IF FF1REC.RID = FF3SPOUSE.SP1ID THEN
          SPID = FF3SPOUSE.SP2ID: MPTR = FF3SPOUSE.SP1NXT
          ELSE
          SPID = FF3SPOUSE.SP1ID: MPTR = FF3SPOUSE.SP2NXT
       END IF
       IF MADRS <> 0 THEN
          CALL FF2GETREC(MADRS)
          IF ADRSPTR <> 0 THEN
             D2Y = FF2PFX.D1Y: D2MD = FF2PFX.D1MD
             IF ((D1Y < D2Y) OR ((D1Y = D2Y) AND (D1MD <= D2MD))) THEN
                ADRSPTR = MADRS: ADRSSPID = SPID
                ELSE
                CALL FF2GETREC(ADRSPTR)
             END IF
             ELSE
             ADRSPTR = MADRS: ADRSSPID = SPID
          END IF
          ELSE
       END IF
    WEND

END SUB

SUB GetNextMARRIAGE
SHARED SPRESEQ, MLIST(), SPCNT, SPRNO, INDID
   IF SPRESEQ THEN
      IF SPCNT > 0 THEN SPCNT = SPCNT - 1
      SPRNO = MLIST(SPCNT)
      ELSE
      IF INDID = FF3SPOUSE.SP1ID THEN
         SPRNO = FF3SPOUSE.SP1NXT
         ELSE
         IF INDID = FF3SPOUSE.SP2ID THEN
            SPRNO = FF3SPOUSE.SP2NXT
            ELSE
            SPRNO = 0
         END IF
      END IF
   END IF

END SUB

SUB SelectFILE (OPT$)
SHARED SELSET(), LOID, HIID, SlctCNT, SecSlctCNT
SHARED SBASE, SECSLCT, SAVEWORK
SHARED SRuleCNT, UFLD(), RULES() AS SlctFileRULE

    SELECT CASE OPT$
       CASE "SAVE": A = 1: GOTO 820
       CASE "LOAD": A = 2: GOTO 820
       CASE "MERG": A = 3: GOTO 820
       CASE "AND": A = 4: GOTO 820
       CASE "EXCL": A = 5: GOTO 820
    END SELECT

810 ' Get Select FILE Processing Option
    CALL PutMSG("LISTM009")
    IF A = 27 THEN 899
    A = VAL(A$)
    IF A < 1 OR A > 5 THEN CALL ErrBEEP(0): GOTO 810
820 FO = A
    SELECT CASE FO
       CASE 1: GOSUB 830
       CASE ELSE: GOSUB 840
    END SELECT
    CLOSE : ENV.FFOPEN = 0: SAVEWORK = (FO < 3)
    CALL SelectSET("STATS")
    IF ((OPT$ = "") AND (A <> 27)) THEN 810
    GOTO 899

830 ' Used to Create SELECT File from SELSET(), RULES()
   CALL SWRKOpen("NEW")
   IF OKAY THEN
      CALL PutMSG("XXXXM060")
      SEEK #6, LEN(SlctHDR) + 1
      SlctHDR.SPGM = MID$(ENV.PGMNAME, 4, 4)
      SlctHDR.SRULCNT = CHR$(SRuleCNT)
      SlctHDR.SECSLCT = CHR$(SECSLCT)
      IF SRuleCNT > 0 THEN
         FOR X = 1 TO SRuleCNT: PUT #6, , RULES(X): NEXT X
      END IF
      SlctHDR.LOID = 32767: SlctHDR.HIID = 0
      X = 0: Y = 0: Z0$ = CHR$(0): Z$ = " "
      WHILE X < HIID
         X = X + 1: Z = SELSET(X)
         LSET Z$ = CHR$(Z): PUT #6, , Z$
         IF Z > 0 THEN
            SlctHDR.SlctCNT = SlctHDR.SlctCNT - ((Z AND 1) = 1)
            SlctHDR.SECSCNT = SlctHDR.SECSCNT - ((Z AND 3) = 2)
            IF X < SlctHDR.LOID THEN SlctHDR.LOID = X
            SlctHDR.HIID = X
         END IF
      WEND
      SlctHDR.SLCTTOT = SlctHDR.SlctCNT + SlctHDR.SECSCNT
      SlctHDR.WFVER = CHR$(2)
      PUT #6, 1, SlctHDR
   END IF
   RETURN

840 ' Select File = 1 Byte for Each ID in .NAM, <>CHR$(0) if selected
   CALL SWRKOpen("READ")
   IF OKAY THEN
      SECSLCT = ASC(SlctHDR.SECSLCT)
      SEEK #6, LEN(SlctHDR) + 1
      SRCNT = ASC(SlctHDR.SRULCNT)
      IF SRCNT > 0 THEN
         X = SRuleCNT
         WHILE X > 0: RULES(X) = RULES(0): UFLD(X) = 0: X = X - 1: WEND
         FOR X = 1 TO SRCNT
             GET #6, , RULES(X)
             CALL ChkRULE(RULES(X), X)
         NEXT X
         SRuleCNT = SRCNT
      END IF
      Z$ = " ": X = 0: Y = 0
      OHIID = HIID: LOID = 32767: HIID = 0
      HI = SlctHDR.HIID: IF HI > FF1Hdr.MAXID THEN HI = FF1Hdr.MAXID
      IF FO = 2 THEN REDIM SELSET(FF1Hdr.MAXID)
      SlctCNT = 0: SecSlctCNT = 0
      WHILE X < HI
         X = X + 1
         GET #6, , Z$: Z = ASC(Z$)
         SELECT CASE FO
            CASE 2:                                               'LOAD
            CASE 3: Z = SELSET(X) OR Z: IF (Z AND 1) THEN Z = 1   'MERG
            CASE 4: Z = SELSET(X) AND Z                           'AND
            CASE 5: Z = SELSET(X) AND (Z = 0)                     'EXCLUDE
         END SELECT
         SELSET(X) = Z
         IF Z > 0 THEN
            SlctCNT = SlctCNT - ((Z AND 1) = 1)
            SecSlctCNT = SecSlctCNT - ((Z AND 3) = 2)
            IF X < LOID THEN LOID = X
            IF X > HIID THEN HIID = X
         END IF
      WEND
      IF FO > 2 THEN
         WHILE X < OHIID
            X = X + 1
            Z = SELSET(X)
            IF FA = 4 THEN
               Z = 0
               ELSE
               IF Z > 0 THEN
                  SlctCNT = SlctCNT - ((Z AND 1) = 1)
                  SecSlctCNT = SecSlctCNT - ((Z AND 3) = 2)
                  IF X < LOID THEN LOID = X
                  IF X > HIID THEN HIID = X
               END IF
            END IF
         WEND
      END IF
      SELSET(0) = LOID
   END IF
   RETURN

899 ' Return to CALLER

END SUB

REM $DYNAMIC
SUB SelectID
SHARED LOID, HIID, RID, SlctCNT, SecSlctCNT, SELSET()
SHARED FMT002, FMT004, FMT005, FMT901
   CALL WinSWITCH(2): CALL WinCLR
   CALL WinDisplayLABELS(FMT004)
   CALL FAMOpen("READ"): IF NOT OKAY THEN 195
   CALL FmtFindFLD("RID"): IF NOT FOUND THEN 195
   RIDFLD = WIN.CURFLD: RIDUTXT = FFLD.UTXT
   RID = LOID: IF RID > FF1Hdr.MAXID THEN RID = 0
   BD$ = SPACE$(11): BPLACE$ = SPACE$(41)
   DD$ = BD$: DPLACE$ = BPLACE$: AGE$ = SPACE$(3)

90 ' Get ID and/or Function Key
   IF ((SlctCNT + SecSlctCNT) > 1) THEN
      CALL PutMSG("LISTM010")
      ELSE
      CALL PutMSG("")
   END IF
   GOSUB 185
   CALL WinUPDATE: IF A = 27 THEN 195
   ID = VAL(UTXT$(RIDUTXT))
   IF (ID < 1) OR (ID > FF1Hdr.MAXID) THEN
      IF A = 60 THEN 95 ELSE CALL ErrBEEP(0): GOTO 90
   END IF
   IF (SELSET(ID) = 0) THEN
      SlctCNT = SlctCNT + 1: SELSET(ID) = 1
      IF ID < LOID THEN LOID = ID
      IF ID > HIID THEN HIID = ID
   END IF
   RID = ID
95 SELECT CASE A
      CASE 13:
      CASE 59:  GOTO 195
      CASE 60:  GOSUB 130
      CASE 61:  GOSUB 140
      CASE 62:  SELSET(RID) = (SELSET(RID) XOR 3)
                IF (SELSET(RID) AND 1) THEN SlctCNT = SlctCNT + 1 ELSE SlctCNT = SlctCNT - 1
                IF (SELSET(RID) AND 2) THEN SecSlctCNT = SecSlctCNT + 1 ELSE SecSlctCNT = SecSlctCNT - 1
      CASE 73:  GOSUB 125
      CASE 81:  GOSUB 120
      CASE 118: RID = HIID
      CASE 132: RID = LOID
      CASE ELSE: CALL ErrBEEP(0): GOTO 90
   END SELECT
   GOTO 90

120 ' PGDN List of ID's
    IF RID < HIID THEN
       RID = RID + 1
       WHILE (RID < HIID) AND (SELSET(RID) = 0): RID = RID + 1: WEND
       ELSE
       CALL ErrBEEP(0)
    END IF
    RETURN

125 ' PGUP List of ID's
    IF RID > LOID THEN
       RID = RID - 1
       WHILE (RID > LOID) AND (SELSET(RID) = 0): RID = RID - 1: WEND
       ELSE
       CALL ErrBEEP(0)
    END IF
    RETURN

130 ' Enter RANGE of ID's
    WIN.MFMT = FMT005: CALL MsgFORMAT
    IF A <> 27 THEN
       CALL FmtFindFLD("ID1"): ID1 = VAL(UTXT$(FFLD.UTXT))
       CALL FmtFindFLD("ID2"): ID2 = VAL(UTXT$(FFLD.UTXT))
       IF ID2 < ID1 THEN SWAP ID1, ID2
       IF ID2 > FF1Hdr.MAXID THEN
          ID2 = FF1Hdr.MAXID
          IF ID1 > ID2 THEN ID1 = ID2
       END IF
       IF (ID1 < 1) THEN CALL ErrBEEP(0): GOTO 130
       FOR I = ID1 TO ID2
           IF SELSET(I) = 0 THEN SlctCNT = SlctCNT + 1: SELSET(I) = 1
       NEXT I
       IF (LOID > ID1) THEN LOID = ID1
       IF (HIID < ID2) THEN HIID = ID2
       RID = ID1
    END IF
    RETURN

140 ' Delete ID from SELSET()
    SlctCNT = SlctCNT + (SELSET(RID) = 1)
    SecSlctCNT = SecSlctCNT + (SELSET(RID) = 2)
    T = SlctCNT + SecSlctCNT
    SELSET(RID) = 0: ID = RID
    IF ID = LOID THEN
       IF T > 0 THEN GOSUB 120: LOID = RID ELSE LOID = 32767
    END IF
    IF ID = HIID THEN
       IF T > 0 THEN GOSUB 125: HIID = RID ELSE HIID = 0
    END IF
    IF (LOID < ID) AND (ID < HIID) THEN GOSUB 120
    IF T = 0 THEN RID = 0
    RETURN

185 ' Display Name Record Information
    CURFMT = FMT004: FMT = FMTTAB(CURFMT)
    OKAY = 0: IF (RID > 0) AND (RID <= FF1Hdr.MAXID) THEN CALL FF1GetRec(RID)
    IF NOT OKAY THEN CALL FmtDREST: CALL WinDisplayLABELS(0): GOTO 189
    CALL GetBDEvents(BD$, BPLACE$, DD$, DPLACE$, AGE$)
    X = FMT.BGNFLD
    WHILE X > 0
       FFLD = FLDTAB(X)
       IF FFLD.UTXT > 0 THEN X$ = UTXT$(FFLD.UTXT)
       SELECT CASE FFLD.NAME
          CASE "RID ": RSET X$ = LTRIM$(STR$(RID))
          CASE "GVNM": LSET X$ = FF1REC.GIVEN
          CASE "SRNM": LSET X$ = FF1REC.SURNM
          CASE "SEX ": RSET X$ = FF1REC.SEX
          CASE "BDT ": LSET X$ = BD$
          CASE "BDST": LSET X$ = FDSTAT$("BD")
          CASE "BPLC": LSET X$ = BPLACE$
          CASE "DDT ": LSET X$ = DD$
          CASE "DDST": LSET X$ = FDSTAT$("DD")
          CASE "DPLC": LSET X$ = DPLACE$
          CASE "AGE ": LSET X$ = AGE$
          CASE "FID ": LSET X$ = STR$(FF1REC.FID)
          CASE "FST ": LSET X$ = FDSTAT$("FID")
          CASE "MID ": LSET X$ = STR$(FF1REC.MID)
          CASE "MST ": LSET X$ = FDSTAT$("MID")
          CASE "PRIM": CALL FmtFindFLD("PRIM")
                       IF (SELSET(RID) AND 1) THEN COLOR FG(6), BG(6)
                       PRINT DTXT$(X);
          CASE "PSCT": RSET X$ = LTRIM$(STR$(SlctCNT))
          CASE "SECN": CALL FmtFindFLD("SECN")
                       IF (SELSET(RID) AND 2) THEN COLOR FG(6), BG(6)
                       PRINT DTXT$(X);
          CASE "SSCT": RSET X$ = LTRIM$(STR$(SecSlctCNT))
       END SELECT
       IF FFLD.UTXT > 0 THEN LSET UTXT$(FFLD.UTXT) = X$
       X = FLDCHN(X).FWD
    WEND
189 CALL WinDisplayDATA(0)
    RETURN

195 ' Return to Main Program
    CALL FamCLOSE: CALL WinCLR
    CALL WinFORMAT(FMT002)
    CALL OptDISPLAY(FMT901)
    CALL WinSWITCH(1)
    EXIT SUB

END SUB

REM $STATIC
SUB SelectSET (OPT$)
SHARED FMT001, FMT002, FMT003, FMT901
SHARED XBASE, SBASE, SlctCNT, SecSlctCNT, SRuleCNT, LOID, HIID, SAVEWORK
SHARED SECSLCT, SELSET(), UFLD(), RULES() AS SlctFileRULE
SHARED SNDX$, KEYWORD$, SRCHWORD$

    DIM XRULE AS SlctFileRULE
    SELECT CASE OPT$
       CASE "DEFAULT": GOSUB 170
       CASE "STATS": GOSUB 230
       CASE "CREATE": GOSUB 500
    END SELECT
    CALL FamCLOSE
    EXIT SUB

170 ' Prepare Default Selection Set
    XBASE = SBASE: SBASE = 0: SlctCNT = 0: SecSlctCNT = 0: SAVEWORK = -1
    CALL FAMOpen("READQ")
    IF OKAY THEN
       REDIM SELSET(FF1Hdr.MAXID)
       GOSUB 520: SBASE = 1
    END IF
    RETURN

230 ' Display Selection Type, Select Count
    CURFMT = FMT001: FMT = FMTTAB(CURFMT)
    CALL FmtFindFLD("RCNT")
    IF FOUND THEN
       LSET UTXT$(FFLD.UTXT) = STR$(SRuleCNT)
       PRINT UTXT$(FFLD.UTXT)
    END IF
    CALL FmtFindFLD("PSCT")
    IF FOUND THEN
       LSET UTXT$(FFLD.UTXT) = STR$(SlctCNT)
       PRINT UTXT$(FFLD.UTXT)
    END IF
    CALL FmtFindFLD("SSCT")
    IF FOUND THEN
       LSET UTXT$(FFLD.UTXT) = STR$(SecSlctCNT)
       PRINT UTXT$(FFLD.UTXT)
    END IF
    C = 1: C$ = CHR$(XBASE + 64): GOSUB 235
    C = 5: C$ = CHR$(SBASE + 64): GOSUB 235
    GOTO 239
235 CALL FmtFindFLD(C$)
    IF FOUND THEN COLOR FG(C), BG(C): PRINT C$;
239 RETURN

250 ' Place Selection Rule Data on Screen
    XRULE = RULES(R): YORN = XRULE.YORN
    CURFMT = FMT002: FMT = FMTTAB(CURFMT)
    X = FMT.BGNFLD
    WHILE X > 0
       FFLD = FLDTAB(X)
       IF FFLD.UTXT > 0 THEN
          X$ = UTXT$(FFLD.UTXT)
          SELECT CASE FFLD.NAME
             CASE "RNUM": LSET X$ = LTRIM$(STR$(R))
             CASE "RTOT": LSET X$ = LTRIM$(STR$(SRuleCNT))
             CASE "LOID": RSET X$ = LTRIM$(STR$(XRULE.LOID))
             CASE "HIID": RSET X$ = LTRIM$(STR$(XRULE.HIID))
             CASE "SEX ": LSET X$ = XRULE.SEX
             CASE "FST ": LSET X$ = XRULE.FIDST
             CASE "MST ": LSET X$ = XRULE.MIDST
             CASE "SRNM": LSET X$ = XRULE.SRNM
             CASE "GVNM": LSET X$ = XRULE.GVNM
             CASE "LOBD": LSET X$ = XSCRNDate$(XRULE.LOBMD, XRULE.LOBY)
             CASE "HIBD": LSET X$ = XSCRNDate$(XRULE.HIBMD, XRULE.HIBY)
             CASE "BDST": LSET X$ = XRULE.BDST
             CASE "BPLC": LSET X$ = XRULE.BPLC
             CASE "LODD": LSET X$ = XSCRNDate$(XRULE.LODMD, XRULE.LODY)
             CASE "HIDD": LSET X$ = XSCRNDate$(XRULE.HIDMD, XRULE.HIDY)
             CASE "DDST": LSET X$ = XRULE.DDST
             CASE "DPLC": LSET X$ = XRULE.DPLC
             CASE "MARY": LSET X$ = CHR$(-(xY * ((YORN AND 1) > 0) + xN * ((YORN AND 2) > 0) + xC * ((YORN AND 4) > 0) + 32 * ((YORN AND 7) = 0)))
             CASE "CHLD": LSET X$ = CHR$(-(xY * ((YORN AND 8) > 0) + xN * ((YORN AND 16) > 0) + 32 * ((YORN AND 24) = 0)))
             CASE "RESI": LSET X$ = CHR$(-(xY * ((YORN AND 32) > 0) + xN * ((YORN AND 64) > 0) + xC * ((YORN AND 128) > 0) + 32 * ((YORN AND 224) = 0)))
             CASE "COMM": LSET X$ = CHR$(-(xY * ((YORN AND 256) > 0) + xN * ((YORN AND 512) > 0) + 32 * ((YORN AND 768) = 0)))
          END SELECT
          LSET UTXT$(FFLD.UTXT) = X$
       END IF
       X = FLDCHN(X).FWD
    WEND
    CALL WinDisplayDATA(0)
    RETURN

500 ' Create Selection SET
    xY = ASC(SNGLKEY$(1)): xN = ASC(SNGLKEY$(2))
    xC = ASC(SNGLKEY$(18))
    CALL WinPREP(2)
    IF WIN.WFMT <> FMT002 THEN
       CALL WinCLR
       CALL WinDisplayLABELS(FMT002)
    END IF
    R = 1: GOSUB 250
    CALL OptDISPLAY(FMT901)
    CALL WinPREP(1)
    XBASE = SBASE
505 GOSUB 230
    CALL PutMSG("LISTM011")
    IF A = 27 THEN 549
    A = VAL(A$): IF A < 1 OR A > 5 THEN CALL ErrBEEP(0): GOTO 505
    ApplyRULE = (A = 3): SecondSELECT = (A = 4)
506 ON A GOTO 510, 550, 600, 600
    CALL SelectID: GOTO 505

510 ' Select BASE Set
    CALL PutMSG("LISTM012")
    IF A = 27 THEN 505
    IF A$ = "0" THEN X = 64 ELSE X = A + (A > 96) * 32
    IF (X < 64 OR X > 68) THEN CALL ErrBEEP(0): GOTO 510
    XBASE = SBASE: X = X - 64
    IF X = 0 THEN SBASE = XBASE ELSE SBASE = X
    GOSUB 230
    CALL FAMOpen("READ")
    IF OKAY THEN
       ON X + 1 GOSUB 515, 520, 530, 540, 545
    END IF
    IF NOT OKAY THEN A = SBASE: SBASE = XBASE: XBASE = A
    GOSUB 230: XBASE = 0
    CLOSE : ENV.FFOPEN = 0
    IF OKAY THEN 505
    GOTO 510

515 CALL PutMSG("LISTM013")
    IF A$ = CHR$(27) THEN RETURN
    A = VAL(A$): IF (A < 1) OR (A > 4) THEN CALL ErrBEEP(0): GOTO 515
    SlctCNT = 0: SecSlctCNT = 0
    LO = 32767: HI = 0
    IF A = 4 THEN GOTO 517

    ' Select only Primary/Secondary
    FOR I = LOID TO HIID
        X = SELSET(I) AND A
        IF X > 0 THEN
           SlctCNT = SlctCNT + 1: X = 1
           HI = I: IF I < LO THEN LO = I
        END IF
        SELSET(I) = X
    NEXT I
    GOTO 519

517 ' Select Complement
    FOR I = 1 TO FF1Hdr.MAXID
        X = SELSET(I)
        SELECT CASE X
           CASE 0: SlctCNT = SlctCNT + 1: X = 1
           CASE 1: X = 0
           CASE 2: SecSlctCNT = SecSlctCNT + 1: X = 2
        END SELECT
        SELSET(I) = X
        IF X <> 0 THEN HI = I: IF I < LO THEN LO = I
    NEXT I

519 ' Finish Current BASE Operation
    LOID = LO: HIID = HI: OKAY = -1
    IF A = 4 THEN GOSUB 230: GOTO 515
    RETURN

520 LOID = 1: HIID = FF1Hdr.MAXID         ' Init SELSET from Full File
    SlctCNT = FF1Hdr.MAXID: SecSlctCNT = 0: OKAY = -1
    FOR I = 1 TO FF1Hdr.MAXID: SELSET(I) = 1: NEXT I
    RETURN

530 CALL SelectFILE("ANY")                ' Init SELSET from Select File
    OKAY = (OKAY AND (A <> 27))
    IF OKAY THEN
       CALL WinSWITCH(2): R = 1: GOSUB 250
       CALL WinSWITCH(1)
    END IF
    RETURN

540 S = SlctCNT + SecSlctCNT               'Init SELSET from REL Work File
    IF S > 0 THEN
       CALL PutMSG("LISTM014")
       OKAY = (A <> 27): IF NOT OKAY THEN RETURN
       SOPT = VAL(A$): IF (SOPT < 1) OR (SOPT > 4) THEN SOPT = 1
       ELSE
       SOPT = 1
    END IF
    CALL RWRKOpen("READ", 0, 0): IF NOT OKAY THEN RETURN
    LOID = 32767: HIID = 0: SlctCNT = 0: SecSlctCNT = 0
    SELECT CASE SOPT
       CASE 1: REDIM SELSET(FF1Hdr.MAXID)
       CASE 3: FOR X = 1 TO FF1Hdr.MAXID: SELSET(X) = -SELSET(X): NEXT X
    END SELECT
    CALL RWRKGet("BASE", RC, PTR&)
    WHILE RC > 0 AND RC < 9
       ID = ABS(WFRELREC.ID): S = SELSET(ID)
       SELECT CASE SOPT
          CASE 3: IF S < 0 THEN S = -S
          CASE 4: IF S > 0 THEN S = 0
          CASE ELSE: S = 1
       END SELECT
       SELSET(ID) = S
       CALL RWRKGet("WRKREC", RC, PTR&)
    WEND
    FOR X = 1 TO FF1Hdr.MAXID
       S = SELSET(X)
       SELECT CASE SOPT
          CASE 3: IF S < 0 THEN S = 0: SELSET(X) = S
       END SELECT
       IF S > 0 THEN
          IF S = 1 THEN SlctCNT = SlctCNT + 1
          IF S = 2 THEN SecSlctCNT = SecSlctCNT + 1
          HIID = X: IF X < LOID THEN LOID = X
       END IF
    NEXT X
    RETURN

545 SlctCNT = 0: LOID = 32767: HIID = 0  ' Init SELSET for Empty Set
    SecSlctCNT = 0
    REDIM SELSET(FF1Hdr.MAXID)
    RETURN

549 ' Return from Creating SELSET()
    SELSET(0) = LOID
    EXIT SUB

550 ' Change Selection RULEs
    CALL WinSWITCH(2)
    CURFMT = FMT002: FMT = FMTTAB(CURFMT)
555 CALL WinUPDATE: GOSUB 560: GOSUB 570
    SELECT CASE A
       CASE 59: R = 1: GOSUB 250: GOTO 595
       CASE 60: RULES(R) = RULES(0): UFLD(R) = UFLD(0): GOSUB 575
       CASE 61: RULES(10) = RULES(R): UFLD(10) = UFLD(R)
       CASE 62: RULES(R) = RULES(10): UFLD(R) = UFLD(10): GOSUB 575
       CASE 73: IF (R > 1) THEN R = R - 1
       CASE 81: IF (R < 9) AND (NOT (R > SRuleCNT)) THEN R = R + 1
       CASE 105: R = 1: SRuleCNT = 0: RULES(1) = RULES(0): UFLD(1) = 0
       CASE ELSE: CALL ErrBEEP(0): GOTO 555
    END SELECT
    GOSUB 250
    GOTO 555

560 ' Read Updated RULE INFO from screen
    YORN = 0: DFM = VAL(CONFIG.DFM)
    X = FMT.BGNFLD
    WHILE X > 0
       FFLD = FLDTAB(X)
       IF FFLD.UTXT > 0 THEN
          X$ = UTXT$(FFLD.UTXT)
          SELECT CASE FFLD.NAME
             CASE "RNUM": R = VAL(X$)
             CASE "LOID": GOSUB 569: XRULE.LOID = V#
             CASE "HIID": GOSUB 569: XRULE.HIID = V#
             CASE "SEX ": XRULE.SEX = X$
             CASE "FST ": XRULE.FIDST = X$
             CASE "MST ": XRULE.MIDST = X$
             CASE "SRNM": LSET XRULE.SRNM = X$
             CASE "GVNM": LSET XRULE.GVNM = X$
             CASE "LOBD": XRULE.LOBMD = DATEMD%(X$, DFM)
                          XRULE.LOBY = DATEY%(X$, DFM)
             CASE "HIBD": XRULE.HIBMD = DATEMD%(X$, DFM)
                          XRULE.HIBY = DATEY%(X$, DFM)
             CASE "BDST": XRULE.BDST = X$
             CASE "BPLC": LSET XRULE.BPLC = X$
             CASE "LODD": XRULE.LODMD = DATEMD%(X$, DFM)
                          XRULE.LODY = DATEY%(X$, DFM)
             CASE "HIDD": XRULE.HIDMD = DATEMD%(X$, DFM)
                          XRULE.HIDY = DATEY%(X$, DFM)
             CASE "DDST": XRULE.DDST = X$
             CASE "DPLC": LSET XRULE.DPLC = X$
             CASE "MARY": YORN = YORN - (1 * (ASC(X$) = xY) + 2 * (ASC(X$) = xN) + 4 * (ASC(X$) = xC))
             CASE "CHLD": YORN = YORN - (8 * (ASC(X$) = xY) + 16 * (ASC(X$) = xN))
             CASE "RESI": YORN = YORN - (32 * (ASC(X$) = xY) + 64 * (ASC(X$) = xN) + 128 * (ASC(X$) = xC))
             CASE "COMM": YORN = YORN - (256 * (ASC(X$) = xY) + 512 * (ASC(X$) = xN))
          END SELECT
       END IF
       X = FLDCHN(X).FWD
    WEND
    XRULE.YORN = YORN
    RULES(R) = XRULE
    CALL ChkRULE(RULES(R), R)
    RETURN

569 ' CHECK FOR INTEGER TOO LARGE
    V# = VAL(X$): IF V# > 32767 THEN V# = 32767
    RETURN

570 ' Adjust Rule Count (if necessary)
    IF (R >= SRuleCNT) AND (UFLD(R) <> 0) THEN
       SRuleCNT = R
       IF (R < 9) THEN RULES(R + 1) = RULES(0): UFLD(R + 1) = 0
    END IF
575 IF (R = SRuleCNT) AND (UFLD(R) = 0) THEN SRuleCNT = SRuleCNT - 1
    RETURN

595 ' Eliminate Blank Rules
    R = SRuleCNT: J = 0: I = 1
    WHILE I <= SRuleCNT
       IF (UFLD(I) = 0 AND J = 0) THEN
          J = I
          ELSE
          IF (UFLD(I) > 0 AND J > 0) THEN
             RULES(J) = RULES(I): R = J: J = J + 1
          END IF
       END IF
       I = I + 1
    WEND: IF UFLD(1) = 0 THEN SRuleCNT = 0 ELSE SRuleCNT = R
    CALL WinSWITCH(1)
599 GOTO 505

600 ' Apply RULES or Secondary Selections to Selection BASE
    IF SBASE = 0 THEN
       CALL PutMSG("LISTM015")
       GOTO 695
    END IF
    IF FF1Hdr.MAXID = 0 THEN
       CALL PutMSG("LISTM016")
       GOTO 695
    END IF
    IF (SlctCNT > 0) AND (SecSlctCNT > 0) THEN
       CALL PutMSG("LISTM017")
       IF A = 27 THEN 695
       ST = VAL(A$): IF (ST < 1) OR (ST > 3) THEN ST = 3
       ELSE
       ST = 3
    END IF
    PCNT = -(SlctCNT * (ST AND 1) + SecSlctCNT * ((ST AND 2) > 0))
    IF PCNT = 0 THEN
       CALL PutMSG("LISTM018")
       GOTO 695
    END IF
    IF SecondSELECT THEN GOTO 610

605 ' Determine LOID and HIID in Selection Rules
    IF SRuleCNT = 0 THEN 690
    LO = 32767: HI = 0
    FOR R = 1 TO SRuleCNT
      IF LO > RULES(R).LOID THEN LO = RULES(R).LOID
      IF HI < RULES(R).HIID THEN HI = RULES(R).HIID
    NEXT R
    IF LOID < LO THEN LOID = LO
    IF HIID > HI THEN HIID = HI
    GOTO 630

610 ' Get Secondary Select Options
    OSECSLCT = SECSLCT
    CALL WinSWITCH(2): CALL WinCLR: CALL WinFORMAT(FMT003)
    FOR A = 1 TO 6: GOSUB 618: NEXT A
    CALL PutMSG("LISTM019")
615 CALL GETRESP: IF (A = 27 OR A = 13) THEN 619
    A = A - 64 + 32 * (A > 96): A = -A * (0 < A AND A < 7)
    IF A = 0 THEN CALL ErrBEEP(0): GOTO 615
    SECSLCT = SECSLCT XOR 2 ^ A
    GOSUB 618: GOTO 615

618 ' SHOW SECSLCT Option
    X$ = CHR$(A + 64): CALL FmtFindFLD(X$)
    IF (SECSLCT AND 2 ^ A) THEN COLOR FG(6), BG(6) ELSE COLOR FG(1), BG(1)
    PRINT X$;
    RETURN

619 ' Finished With Choosing Secondary Select Options
    CALL WinCLR: CALL WinFORMAT(FMT002): CALL OptDISPLAY(FMT901)
    CALL WinSWITCH(1)
    IF A = 27 THEN SECSLCT = OSECSLCT: GOTO 695
    IF SECSLCT = 0 THEN GOTO 695

630 ' Restrict to ID's between LOID and HIID
    SlctCNT = 0: SecSlctCNT = 0: CHKCNT = 0
    LO = 32767: HI = 0
    I = 1: WHILE I < LOID: SELSET(I) = 0: I = I + 1: WEND
    WHILE I < HIID + 1
       IF (SELSET(I) > 0) THEN
          IF (SELSET(I) AND ST) > 0 THEN
             SELSET(I) = SELSET(I) OR 8: CHKCNT = CHKCNT + 1
          END IF
          SlctCNT = SlctCNT - ((SELSET(I) AND 1) > 0)
          SecSlctCNT = SecSlctCNT - ((SELSET(I) AND 2) > 0)
          IF HI = 0 THEN LO = I
          HI = I
       END IF
       I = I + 1
    WEND
    WHILE I < FF1Hdr.MAXID + 1: SELSET(I) = 0: I = I + 1: WEND
    LOID = LO: HIID = HI: GOSUB 230

640 ' Process Selection Option
    IF CHKCNT = 0 THEN GOTO 695
    CALL FAMOpen("READ"): IF NOT OKAY THEN 695
    LO = FF1Hdr.MAXID: HI = 0
    IF ApplyRULE THEN CALL PutMSG("LISTM020") ELSE CALL PutMSG("LISTM022")
    FOR I = LOID TO HIID
    IF (SELSET(I) AND 8) > 0 THEN
       IF ApplyRULE THEN R = 1: GOSUB 650 ELSE GOSUB 700
    END IF

    ' Check for User KEY Press
    A$ = INKEY$
    IF A$ <> "" THEN
       A = ASC(RIGHT$(A$, 1))
       IF A = 27 THEN
          I = HIID
          ELSE
          SYSVAR.NVAR1 = I
          CALL PutMSG("LISTM021")
       END IF
    END IF
    NEXT I
    GOTO 690

650 ' Process Selection Rule "R"
    UFLD = UFLD(R): XRULE = RULES(R)
    IF (UFLD AND 1) THEN
       IF (XRULE.LOID > I) OR (XRULE.HIID < I) THEN
          OKAY = 0: GOTO 675
       END IF
    END IF
    CALL FF1GetRec(I): IF NOT OKAY THEN 679
    DY = FF1REC.DY: SPREC = FF1REC.SPOUSE: COMREC = FF1REC.COM
    OLDCH = FF1REC.OLDCH: ADRSPTR = FF1REC.ADRS
    IF UFLD AND 240 THEN
       IF UFLD AND 160 THEN BPL$ = " " ELSE BPL$ = ""
       CALL GetBDEvents(BD$, BPL$, DD$, DPL$, AGE$)
    END IF
    IF (UFLD AND 2) THEN IF XRULE.SEX <> FF1REC.SEX THEN 675
    IF (UFLD AND 4) THEN
       LSET KEYWORD$ = RTRIM$(XRULE.SRNM)
       LSET SRCHWORD$ = RTRIM$(FF1REC.SURNM)
       CALL TEXTSearch
       IF NOT OKAY THEN 675
    END IF
    IF (UFLD AND 8) THEN
       LSET KEYWORD$ = RTRIM$(XRULE.GVNM)
       LSET SRCHWORD$ = RTRIM$(FF1REC.GIVEN)
       CALL TEXTSearch
       IF NOT OKAY THEN 675
    END IF
    IF (UFLD AND 16) THEN
       IF (FF1REC.BY < XRULE.LOBY) OR (FF1REC.BY > XRULE.HIBY) THEN 675
       IF (FF1REC.BMD < XRULE.LOBMD) OR (FF1REC.BMD > XRULE.HIBMD) THEN 675
    END IF
    IF (UFLD AND 64) THEN
       IF (FF1REC.DY < XRULE.LODY) OR (FF1REC.DY > XRULE.HIDY) THEN 675
       IF (FF1REC.DMD < XRULE.LODMD) OR (FF1REC.DMD > XRULE.HIDMD) THEN 675
    END IF
655 IF (UFLD AND 256) = 0 THEN 660
    YORN = XRULE.YORN
    IF ((YORN AND 5) AND (SPREC = 0)) THEN 675
    IF (YORN AND 4) THEN
       IF DY > 0 THEN 675
       CALL FF3GetRec(SPREC)
       IF (FF3SPOUSE.D2Y > 0) OR (FF3SPOUSE.D2MD > 0) THEN 675
    END IF
    IF ((YORN AND 2) AND (SPREC <> 0)) THEN 675
    IF ((YORN AND 8) AND (OLDCH = 0)) THEN 675
    IF ((YORN AND 16) AND (OLDCH <> 0)) THEN 675
    IF ((YORN AND 256) AND (COMREC = 0)) THEN 675
    IF ((YORN AND 512) AND (COMREC <> 0)) THEN 675
    IF ((YORN AND 224) = 0) THEN 660
    IF (ADRSPTR = 0) AND (SPREC <> 0) THEN
       CALL FF3GetRec(SPREC): ADRSPTR = FF3SPOUSE.LOC
    END IF
    IF ((YORN AND 160) AND (ADRSPTR = 0)) THEN 675
    IF ((YORN AND 64) AND (ADRSPTR <> 0)) THEN 675
    IF ((YORN AND 128) = 0) THEN 660
    CALL GetLastRES: IF (FF2PFX.D2Y <> 0) OR (FF2PFX.D2MD <> 0) THEN 675
660 IF (UFLD AND 160) = 0 THEN 665
    IF (UFLD AND 32) THEN
       LSET KEYWORD$ = RTRIM$(XRULE.BPLC)
       LSET SRCHWORD$ = RTRIM$(BPLACE)
       CALL TEXTSearch
       IF NOT OKAY THEN 675
    END IF
    IF (UFLD AND 128) THEN
       LSET KEYWORD$ = RTRIM$(XRULE.DPLC)
       LSET SRCHWORD$ = RTRIM$(DPLACE)
       CALL TEXTSearch
    END IF
    IF NOT OKAY THEN 675
665 IF (UFLD AND 512) THEN
       OKAY = (XRULE.BDST = RULES(0).BDST)
       IF NOT OKAY THEN
          OKAY = (XRULE.BDST = FDSTAT$("BD"))
          IF NOT OKAY THEN 675
       END IF
       OKAY = (XRULE.DDST = RULES(0).DDST)
       IF NOT OKAY THEN OKAY = (XRULE.DDST = FDSTAT$("DD"))
    END IF
    IF (UFLD AND 1024) THEN
       OKAY = (XRULE.FIDST = RULES(0).FIDST)
       IF NOT OKAY THEN
          OKAY = (XRULE.FIDST = FDSTAT$("FID"))
          IF NOT OKAY THEN 675
       END IF
       OKAY = (XRULE.MIDST = RULES(0).MIDST)
       IF NOT OKAY THEN OKAY = (XRULE.MIDST = FDSTAT$("MID"))
       IF NOT OKAY THEN 675
    END IF
    IF OKAY THEN 679

675 ' increment rule count
    R = R + 1: IF R <= SRuleCNT THEN 650
    SlctCNT = SlctCNT - (SELSET(I) AND 1)
    SecSlctCNT = SecSlctCNT + ((SELSET(I) AND 2) > 0)
    SELSET(I) = 0
    OKAY = 0

679 ' Next ID
    RETURN

690 ' Compute SLCTCNT, SecSLCTCNT and LOID, HIID
    SlctCNT = 0: SecSlctCNT = 0: LOID = 32767: HIID = 0
    FOR I = 1 TO FF1Hdr.MAXID
        X = SELSET(I) AND 7
        IF X > 0 THEN
           IF (SELSET(I) AND 1) THEN
              SlctCNT = SlctCNT + 1: SELSET(I) = 1
              ELSE
              IF (SELSET(I) AND 6) THEN
                 SecSlctCNT = SecSlctCNT + 1: SELSET(I) = 2
              END IF
           END IF
           HIID = I: IF I < LOID THEN LOID = I
       END IF
    NEXT I
    GOSUB 230
    IF SlctCNT + SecSlctCNT = 0 THEN 695
    SAVEWORK = 0

695 ' Finished Applying Rules
    CLOSE : ENV.FFOPEN = 0
    GOTO 505

700 ' Mark Secondary Selections
    CALL FF1GetRec(I): SPREC = FF1REC.SPOUSE: OLDCH = FF1REC.OLDCH
    IF (SECSLCT AND 2) THEN
       ID = FF1REC.FID: GOSUB 790
       ID = FF1REC.MID: GOSUB 790
    END IF
    IF (SECSLCT AND 4) THEN
       FID = FF1REC.FID: MID = FF1REC.MID
       IF FID > 0 THEN
          CALL FF1GetRec(FID): ID = FF1REC.OLDCH
          WHILE ID > 0
             GOSUB 790: CALL FF1GetRec(ID): ID = FF1REC.FCH
          WEND
       END IF
       IF MID > 0 THEN
          CALL FF1GetRec(MID): ID = FF1REC.OLDCH
          WHILE ID > 0
             GOSUB 790: CALL FF1GetRec(ID): ID = FF1REC.MCH
          WEND
       END IF
    END IF
    IF (SECSLCT AND 72) THEN
       ID = OLDCH
       WHILE ID > 0
          CALL FF1GetRec(ID)
          IF (SECSLCT AND 8) THEN GOSUB 790
          IF NOT (((SECSLCT AND 64) = 0) OR (FF1REC.SPOUSE = 0)) THEN
             CALL FF3GetRec(FF1REC.SPOUSE)
             IF ID = FF3SPOUSE.SP1ID THEN
                ID = FF3SPOUSE.SP2ID
                ELSE
                ID = FF3SPOUSE.SP1ID
             END IF
             GOSUB 790
          END IF
          IF I = FF1REC.FID THEN ID = FF1REC.FCH ELSE ID = FF1REC.MCH
       WEND
    END IF
    IF (SECSLCT AND 48) THEN
       WHILE SPREC <> 0
          CALL FF3GetRec(SPREC)
          IF I = FF3SPOUSE.SP1ID THEN
             ID = FF3SPOUSE.SP2ID: SPREC = FF3SPOUSE.SP1NXT
             ELSE
             ID = FF3SPOUSE.SP1ID: SPREC = FF3SPOUSE.SP2NXT
          END IF
          IF ID <> 0 THEN
             IF (SECSLCT AND 16) THEN GOSUB 790
             IF (SECSLCT AND 32) THEN
                CALL FF1GetRec(ID)
                ID = FF1REC.FID: GOSUB 790
                ID = FF1REC.MID: GOSUB 790
             END IF
          END IF
       WEND
    END IF
    RETURN

790 ' Mark ID as Secondary Selection
    IF (ID <> 0) AND (SELSET(ID) = 0) THEN
       SecSlctCNT = SecSlctCNT + 1
       SELSET(ID) = SELSET(ID) OR 4
       IF ID < LO THEN LO = ID
       IF ID > HI THEN HI = ID
    END IF
    RETURN

END SUB

SUB TEXTSearch
SHARED SNDX$, KEYWORD$, SRCHWORD$

X1200: ' Compare KEYWORD$ and SRCHWORD$
     OKAY = 0
     KO = 1: KL = LEN(RTRIM$(KEYWORD$))
     L = LEN(RTRIM$(SRCHWORD$)): IF L = 0 THEN L = 1
     MISMATCH = (MID$(KEYWORD$, 1, 1) = "#"): KO = KO - MISMATCH
     LOSEARCH = (MID$(KEYWORD$, KO, 1) = "^"): KO = KO - LOSEARCH
     EXSEARCH = (MID$(KEYWORD$, KO, 1) = "^"): KO = KO - EXSEARCH
     EXMATCH = (MID$(KEYWORD$, KO, 1) = "="): KO = KO - EXMATCH
     IF MID$(KEYWORD$, KO, 1) = "~" THEN GOTO X1250
     IF (KL < KO) THEN KL = KO
     IF EXMATCH THEN
        OKAY = (LEFT$(SRCHWORD$, L) = MID$(KEYWORD$, KO, KL - KO + 1))
        GOTO 1290
     END IF
1215 OFST = INSTR(LEFT$(SRCHWORD$, L), MID$(KEYWORD$, KO, KL - KO + 1))
     OKAY = (OFST > 0): IF OKAY THEN 1290
     IF NOT LOSEARCH THEN 1290
     ASCA = ASC(MID$(KEYWORD$, KO, 1)): IF (ASCA > 95) THEN 1290
     MID$(KEYWORD$, KO, 1) = CHR$(ASCA + 32)
     GOTO 1215

X1250: ' Perform Soundex Search
1260 CALL ComputeSNDX(MID$(KEYWORD$, KO + 1, KL - KO), SNDX$)
     SNDXCHK$ = SNDX$: GOSUB X1300

1290 ' Return to Caller
     IF MISMATCH THEN OKAY = (NOT OKAY)
     EXIT SUB

X1300: ' Soundex Search
     OFST = 0: SXK$ = LEFT$(SNDXCHK$, 1): OKAY = 0: L = 4
     IF EXSEARCH THEN
        WHILE MID$(SNDXCHK$, L, 1) = "0": L = L - 1: WEND
     END IF
1310 OFST = INSTR(OFST + 1, SRCHWORD$, SXK$)
     IF OFST = 0 THEN 1320
     LSET KEYWORD$ = MID$(SRCHWORD$, OFST, KL)
     CALL ComputeSNDX(KEYWORD$, SNDX$)
     IF LEFT$(SNDX$, L) = LEFT$(SNDXCHK$, L) THEN OKAY = -1: GOTO 1340
     IF OFST < 30 THEN 1310 ELSE OFST = 0
1320 IF NOT LOSEARCH THEN 1340
     SXA = ASC(SXK$)
     IF (SXA < 91) THEN LSET SXK$ = CHR$(SXA + 32): GOTO 1310
1340 RETURN

END SUB

SUB XSUMReport
SHARED DFNUM(), DFPP(), SELSET()
SHARED DATECNT, DATERNG, PLACCNT, MARRCNT, RWIDTH, VRTCL, xSEL
SHARED SUMBASE AS SUMREC, SUM AS SUMREC, TOT AS SUMREC

     HDFM$ = CONFIG.DFM
     CALL PutMSG("LISTM003")
     IF A = 27 THEN GOTO S1040
     IF NOT (A$ = "1") THEN
        CALL PutMSG("LISTM004")
        IF A = 27 THEN GOTO S1040
        WTYPE = VAL(A$)
        IF (WTYPE < 1) OR (WTYPE > 3) THEN WTYPE = 1
        ELSE
        WTYPE = 4
     END IF

     ' Produce Summary Report
     RELSUMOPT = (WTYPE = 3): GENSUMOPT = (WTYPE < 3)
     INDEXOPT = (WTYPE = 4)
     IF GENSUMOPT THEN
        CALL PutMSG("XXXXM044")
        IF A = 27 THEN GOTO S1040
        VRTCL = (A$ = "2")
        ELSE
        VRTCL = 0
     END IF
     SELECT CASE WTYPE
        CASE 1: SRPT$ = "ASUM"
        CASE 2: SRPT$ = "DSUM"
        CASE 3: SRPT$ = "RSUM"
        CASE 4: SRPT$ = "ISUM"
     END SELECT
     CALL RptOPEN(SRPT$): IF NOT OKAY THEN GOTO S1040
     WOPEN$ = "READQ"
S1010: ' Verify Files by Opening
     CALL FAMOpen("READ"): IF NOT OKAY THEN GOTO S1040
     IF INDEXOPT THEN
        CALL INDEXFile("READ")
        IF (A = 27) OR (NOT OKAY) THEN GOTO S1040
        PSK = SX(1) \ 256
        IF (PSK) > 8 THEN
           CALL PutMSG("LISTM025")
           CALL FamCLOSE: GOTO S1010
        END IF
        IF (2 < PSK) AND (PSK < 6) THEN
           DS = (ENV.SortOPTS \ 4) AND 7
           ELSE
           IF PSK <> 8 THEN
              TS = ENV.SortOPTS AND 3
              IF TS < 2 THEN CALL RptOptGET("SNDX", X): PrntSNDX = (X = 1)
           END IF
        END IF
        X$ = "SFLD" + LTRIM$(STR$(PSK))
        CALL GetRVAR(X$, X): IF FOUND THEN SYSVAR.SUBJECT = RVAR$(X)
        GOTO S1050
        ELSE
        CALL RWRKOpen(WOPEN$, WTYPE, 0): IF OKAY THEN GOTO S1050
        IF WOPEN$ = "READQ" THEN
           CALL FamCLOSE
           LSET WOPEN$ = "READ": GOTO S1010
        END IF
     END IF

S1040: ' Bad Return
     CONFIG.DFM = HDFM$
     CLOSE : ENV.FFOPEN = 0: EXIT SUB

S1050: ' Prepare to Print
     IF INDEXOPT THEN
        RPTOPTION.NAMEFMT = (RPTOPTION.NAMEFMT AND 1) - 2 * (TS = 2)
        ELSE
        CALL RWRKGet("BASE", RC, PTR&): IF RC <> 1 THEN GOTO S1040
        SYSVAR.SUBJECT = FMTNAME$(1)
        IF DMAXLV = 0 THEN
           CALL RptOptGET("CNOP", X): COUNTNOP = (X = 1)
           IF NOT VRTCL THEN CALL RptOptGET("LNOP", X): LISTNODES = (X = 1)
           LALIN = LEN(STR$(2 ^ ASC(WFHDR.AMAXLV))): ALIN$ = SPACE$(LALIN)
        END IF
        CALL RptOptGET("CDUP", X): COUNTDUPS = (X = 1)
        CALL RptOptGET("CLIN", X): COUNTCLINE = (X = 1)
     END IF
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
        IF GENSUMOPT THEN
           HIRGL = ASC(WFHDR.AMAXLV) + ASC(WFHDR.DMAXLV)
           IF VRTCL THEN REDIM SUMTAB(HIRGL) AS SUMREC
        END IF
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
     IF INDEXOPT THEN
        IF (2 < PSK) AND (PSK < 6) THEN
           SELECT CASE DS
              CASE 0: CDFM$ = "3": DSKL = 11
              CASE 1: CDFM$ = "1": DSKL = 11
              CASE 2: CDFM$ = "2": DSKL = 11
              CASE 3: CDFM$ = "3": DSKL = 7
              CASE 4: CDFM$ = "1": DSKL = 5
              CASE 5: CDFM$ = "3": DSKL = 4
              CASE 6: CDFM$ = "1": DSKL = 2
              CASE 7: CDFM$ = "2": DSKL = 2
           END SELECT
           SELECT CASE CDFM$
              CASE "1": MMX = 1: DDX = 4: YYX = 7
              CASE "2": DDX = 1: MMX = 4: YYX = 7
              CASE "3": YYX = 1: MMX = 6: DDX = 9
           END SELECT
           CONFIG.DFM = CDFM$
        END IF
        SELECT CASE PSK
           CASE 1: PSKL = SX(1) AND 127
           CASE 2: PSKL = SX(1) AND 127
           CASE 3: PSKL = DSKL: CONFIG.DFM = CDFM$
           CASE 4: PSKL = DSKL: CONFIG.DFM = CDFM$
           CASE 5: PSKL = DSKL: CONFIG.DFM = CDFM$
           CASE 6: PSKL = SX(1) AND 127: BPL$ = SPACE$(PSKL)
           CASE 7: PSKL = SX(1) AND 127: DPL$ = SPACE$(PSKL)
           CASE 8: PSKL = 1
        END SELECT
     END IF

S1200: ' Setup Report Tables & Open Files
     CALL BldSumDEF(WTYPE): IF NOT OKAY THEN GOTO S1040
     IF NOT RELSUMOPT THEN
        CALL PrintOPEN: IF NOT OKAY THEN GOTO S1040
        PLINE$ = SPACE$(PAGE.BWIDTH + PDFFORMS.IMARGIN + PDFFORMS.OMARGIN)
        IF INDEXOPT THEN PRT.SECTION = "S"
        CALL BldSumTITLE
        IF GENSUMOPT THEN
        ' IF Ancestor or Descendant Summary, Retrieve Records by Generation
           WFHDR.WFRULES = CHR$((ASC(WFHDR.WFRULES) AND 252) + 1)
        END IF
     END IF

S1400: ' Begin with Base Record
     SUMBASE.LOBY = 10000: SUMBASE.LOMY = 10000: SUMBASE.LODY = 10000
     SUM = SUMBASE: TOT = SUMBASE
     L = 20: LAGL = 0: LRGL = 0
     A = 0: LGL = -1: FINALTOT = 0
     RGL$ = SPACE$(RFLTH(1))
     IF RELSUMOPT THEN CALL PutMSG("XXXXM064")
     IF INDEXOPT THEN
        LSK$ = SPACE$(PSKL): CSK$ = LSK$
        LSNDX$ = SPACE$(4): CSNDX$ = LSNDX$
        xNDX = 1: RC = 1: PRT.SECTION = "S"
        ELSE
        CALL RWRKGet("BASE", RC, PTR&): IF RC <> 1 THEN GOTO S1040
     END IF
     FIRSTONE = -1: GOTO S1450

S1415: ' Process next entry on chart
     IF INDEXOPT THEN
        xNDX = xNDX + 1: IF xNDX > NDXHdr.NDXRECS THEN RC = 9: GOTO S1455
        ELSE
        IF RELSUMOPT THEN
           CALL RWRKGet("WRKREC", RC, PTR&)
           ELSE
           CALL RWRKGet("NEXT", RC, PTR&)
        END IF
        IF (RC < 1) OR (RC > 8) THEN GOTO S1455
     END IF

S1450: ' Process Workfile record
     IF INDEXOPT THEN
        RELID = NDXID(xNDX)
        IF (SELSET(RELID) AND xSEL) = 0 THEN GOTO S1415
        ELSE
        RELID = ABS(WFRELREC.ID)
        IF (SELSET(RELID) AND xSEL) = 0 THEN GOTO S1415
        AGL = -WFAGLREC.AGL: RGL = AGL + WFGLREC.DGL
     END IF
     IF RELSUMOPT THEN GOSUB S2200:  ELSE GOSUB S2400
     IF A <> 27 THEN GOTO S1415

S1455: ' Process Final Totals
     IF RELSUMOPT THEN
        GOSUB S2210
        ELSE
        GOSUB S2460
        IF VRTCL THEN GOSUB S2700
        IF INDEXOPT THEN
           CALL WriteLINE
           PRT.SECTION = "T": SYSVAR.NVAR1 = TOT.GROUPS
           CALL PrintTITLE
        END IF
        CALL RptCLOSE
        IF PRT.AGAIN THEN
           IF VRTCL THEN GOTO S1455 ELSE GOTO S1400
        END IF
     END IF
     CONFIG.DFM = HDFM$
     CALL FamCLOSE
     EXIT SUB

S2200: ' Process Record for RELATIVE Table
     IF WFRELREC.ID > 0 THEN
        RELCNT(AGL, RGL) = RELCNT(AGL, RGL) + 1
     END IF
     RETURN

S2210: ' Print Relationship Summary Table
      CALL PrintOPEN
      IF NOT OKAY THEN GOTO S1040
      FULLWIDTH = 0: PLINE$ = SPACE$(PAGE.BWIDTH + PDFFORMS.IMARGIN + PDFFORMS.OMARGIN)
      IF (PRT.OPT > 2) AND (PAGE.BWIDTH > PRT.FWIDTH - PDFFORMS.IMARGIN - PDFFORMS.OMARGIN) THEN
         CALL PutMSG("XXXXM070")
         IF ASC(A$) = 27 THEN GOTO S1040 ELSE FULLWIDTH = (A$ <> "1")
         IF FULLWIDTH THEN
            PRT.CFORMS = -1: NOSPAN = -1
            PRT.FWIDTH = PAGE.BWIDTH + PDFFORMS.IMARGIN + PDFFORMS.OMARGIN
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
         IF SHOWREL THEN LSET PLINE$ = RELABEL$: CALL PRINTLINE: LSET RELABEL$ = ""
         LSET PLINE$ = RELCNT$: CALL PRINTLINE: LSET RELCNT$ = ""
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
      CALL PRINTLINE
      MID$(PLINE$, DFPP(1), 255) = ACCUM$
      XAGL = 0: FX = 2: AGLACCUM = 0
      WHILE XAGL >= LORGL
         AGLACCUM = AGLACCUM + AGLTOT(XAGL)
         PP = DFPP(FX): FLTH = RFLTH(DFNUM(FX))
         MID$(PLINE$, PP, FLTH) = SFILL$(AGLACCUM, FLTH)
         FX = FX + 1: XAGL = XAGL - 1
      WEND
      CALL PRINTLINE
      CALL RptCLOSE
      IF PRT.AGAIN THEN GOTO S2215
      RETURN

S2400: ' Process Records for Indexed & ANC/DESC Genlevel Summaries
     CALL FF1GetRec(RELID)
     IF (PSK = 1) THEN
        IF ENV.HusbSURNM THEN XNAME$ = FMTNAME$(0) ELSE SURNM$ = FF1REC.SURNM
     END IF
     CALL GetBDEvents("", BPL$, "", DPL$, "")
     IF INDEXOPT THEN
        SELECT CASE PSK
           CASE 1: LSET CSK$ = SURNM$
           CASE 2: LSET CSK$ = FF1REC.GIVEN
           CASE 3: LSET CSK$ = XSCRNDate$(FF1REC.BMD, FF1REC.BY)
           CASE 4: MRID = FF1REC.SPOUSE: LSET CSK$ = ""
                   IF MRID <> 0 THEN
                      CALL FF3GetRec(MRID)
                      IF OKAY THEN
                         LSET CSK$ = XSCRNDate$(FF3SPOUSE.D1MD, FF3SPOUSE.D1Y)
                      END IF
                   END IF
           CASE 5: LSET CSK$ = XSCRNDate$(FF1REC.DMD, FF1REC.DY)
           CASE 6: LSET CSK$ = BPL$
           CASE 7: LSET CSK$ = DPL$
           CASE 8: LSET CSK$ = FF1REC.SEX
        END SELECT
        IF TS = 1 THEN LSET CSK$ = UCX$(CSK$)
        IF PrntSNDX OR (TS > 1) THEN CALL ComputeSNDX(CSK$, CSNDX$)
     END IF
     IF (SUM.TOTAL = 0) THEN
        IF INDEXOPT THEN LSET LSK$ = CSK$: LSET LSNDX$ = CSNDX$
        GOTO S2420
     END IF
     IF INDEXOPT THEN
        MATCHED = -1
        IF TS > 1 THEN MATCHED = (CSNDX$ = LSNDX$)
        IF MATCHED AND (TS <> 2) THEN
           MATCHED = (MID$(CSK$, 1, PSKL) = MID$(LSK$, 1, PSKL))
        END IF
        ELSE
        MATCHED = (RGL = LRGL)
     END IF
     IF MATCHED THEN GOTO S2420
     IF NOT VRTCL THEN GOSUB S2550
     GOSUB S2570
     FIRSTONE = -1: LAGL = AGL: LRGL = RGL
     LSET LSK$ = CSK$: LSET LSNDX$ = CSNDX$
     SUM = SUMBASE

S2420: ' Increment Summary Counters
     IF RELSUMOPT THEN
        C = -(WFRELREC.ID > 0)
        ELSE
        C = 1
     END IF
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
        CALL WriteLINE: FIRSTONE = 0
     END IF
     FX = 2 - RPTOPTION.Relation: PP = DFPP(FX)
     RSET ALIN$ = STR$(WFANCREC.LINEAGE)
     MID$(PLINE$, PP, LALIN) = ALIN$: PP = PP + LALIN + 2
     IF FF1REC.BY > 0 THEN MID$(PLINE$, PP, 4) = ZFILL$(FF1REC.BY, 4)
     MID$(PLINE$, PP + 4, 1) = "-"
     IF FF1REC.DY > 0 THEN MID$(PLINE$, PP + 5, 4) = ZFILL$(FF1REC.DY, 4)
     MID$(PLINE$, PP + 11, 52) = FMTNAME$(0)
     CALL WriteLINE
     RETURN

S2550: ' Print Ancestor/Descendant Summary Line - Horizontal Style
     FX = 1
     WHILE FX <= PAGE.DFCNT
        X = DFNUM(FX)
        IF (NOT FINALTOT) OR (X < 24) THEN
           PP = DFPP(FX): FLTH = RFLTH(X)
           SELECT CASE X
              CASE 1: IF FIRSTONE THEN RSET RGL$ = STR$(LRGL): MID$(PLINE$, PP, FLTH) = RGL$
              CASE 2: IF FIRSTONE THEN MID$(PLINE$, PP, FLTH) = Relation$("M", "*", LAGL, LRGL, RRELCD$())
              CASE 3: MID$(PLINE$, PP, FLTH) = SFILL$(SUM.TOTAL, FLTH)
                      IF SUM.TOTAL > (SUM.MALE + SUM.FEMALE) THEN MID$(PLINE$, PP + FLTH, 1) = "*"
              CASE 4: IF NOT FINALTOT THEN MID$(PLINE$, PP, FLTH) = SFILL$(SUM.CLINE, FLTH)
              CASE 6: MID$(PLINE$, PP, FLTH) = SFILL$(SUM.NOP, FLTH)
              CASE 7: MID$(PLINE$, PP, FLTH) = SFILL$(SUM.MALE, FLTH)
              CASE 8: MID$(PLINE$, PP, FLTH) = SFILL$(SUM.FEMALE, FLTH)
              CASE 9: MID$(PLINE$, PP, FLTH) = SFILL$(SUM.BDT, FLTH)
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
              CASE 24: MID$(PLINE$, PP, FLTH) = LSNDX$
              CASE 25:
              CASE 26: MID$(PLINE$, PP, 4) = MID$(LSK$, YYX, 4)
              CASE 27: MM = VAL(MID$(LSK$, MMX, 2))
                       IF (0 < MM) AND (MM < 13) THEN
                          MID$(PLINE$, PP, 3) = MID$(RMONTB$(MM + 1), 3, 3)
                       END IF
              CASE 28: DD = VAL(MID$(LSK$, DDX, 2))
                       IF DD > 0 THEN MID$(PLINE$, PP, 2) = MID$(LSK$, DDX, 2)
              CASE ELSE: MID$(PLINE$, PP, FLTH) = LSK$
           END SELECT
        END IF
        FX = FX + 1
     WEND
     CALL PRINTLINE
     RETURN

S2570: ' Update Overall Counters
     TOT.GROUPS = TOT.GROUPS + 1
     TOT.TOTAL = TOT.TOTAL + SUM.TOTAL
     TOT.NOP = TOT.NOP + SUM.NOP
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
        CALL PRINTLINE
        FX = FX + 1
     WEND
     RETURN

END SUB

