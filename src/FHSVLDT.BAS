DECLARE SUB PtrCHK3 (OPT$, X&, B%)
DECLARE SUB BldMatchDEF ()
DECLARE SUB BldMatchTITLE ()
DECLARE SUB ChkForPAUSE ()
DECLARE SUB DataCHECK ()
DECLARE SUB DataERR (ERRGrp%, EMNUM%)
DECLARE SUB MatchCHECK ()
DECLARE SUB MatchGrp (N%, ID%(), G%)
DECLARE SUB PtrCHECK ()
DECLARE SUB PtrERR (ERRGrp%, EMNUM%)
DECLARE SUB ValiDATE (DTMD%, DTYR%)
DECLARE FUNCTION FMTNAME$ (NFMT%)
DECLARE FUNCTION FFRBA& (RNUM%, RLTH%)
DECLARE FUNCTION FFRNUM% (RBA&, RLTH%)
DECLARE FUNCTION RELATION$ (A$, B$, C%, D%, RELCD$())
DECLARE FUNCTION TextDATA$ (X$)
DECLARE FUNCTION SFILL$ (X%, L%)
DECLARE FUNCTION ZFILL$ (X%, L%)
DECLARE FUNCTION UCX$ (X$)
DECLARE FUNCTION XRPTDate$ (MD%, Y%, O%)
DECLARE FUNCTION XSCRNDate$ (D1MD%, D1Y%)
1  REM $INCLUDE: 'FHSCOMON.BAS'
   PN$ = " FAMILY HISTORY SYSTEM - FHSVLDT - File Validation Program"
   CY$ = " (C) Copyright 1985,1986,1987,1988,1990,1991,1992,1993,1996,1997 by Phillip E. Brown"
   IF CONFIG.BP = 0 THEN RUN "FHSINIT"
   DEF FNX$ (A) = LTRIM$(STR$(A))
   REDIM ERID(1), DFNUM(1), DFPP(1), RFABR$(1)
   CALL WinCLR
   CALL WinINIT("LR1")
   CALL WinPREP(1): PGMRTRN$ = "FHSMENU": READ$ = "READQ"
   FUN = 0: WorkSAVED = -1
   CALL RptLOAD("PVLD"): CALL RptLOAD("DVLD")
   CALL FMTFind("VLDTS001"): FMT001 = CURFMT
   CALL FMTFind("VLDTS002"): FMT002 = CURFMT

40 ' Format Display
   CALL WinSWITCH(1)
   CALL WinDisplayLABELS(FMT001)

50 ' Display File, Printer SETUPS
   CURFMT = FMT001: FMT = FMTTAB(CURFMT)
   FF = FMT.BGNFLD
   WHILE FF > 0
      FFLD = FLDTAB(FF)
      IF FFLD.UTXT > 0 THEN
         X$ = UTXT$(FFLD.UTXT)
         SELECT CASE FFLD.Name
            CASE "FSUP": LSET X$ = FDFSETUP.Name
            CASE "FSDS": LSET X$ = FDFSETUP.DESC
            CASE "PSUP": LSET X$ = PDFSetup.Name
            CASE "PSDS": LSET X$ = PDFSetup.DESC
            CASE "PRTR": LSET X$ = PDFSetup.PRINTER
            CASE "FMWD": LSET X$ = STR$(PDFForms.FWIDTH)
            CASE "FMLT": LSET X$ = STR$(PDFForms.FLENGTH)
         END SELECT
         LSET UTXT$(FFLD.UTXT) = X$
      END IF
      FF = FLDCHN(FF).FWD
   WEND

100 ' Check Files
    IF NOT FAMOKAY THEN
       CALL FamOPEN(READ$): FAMOKAY = ENV.FFOPEN: READ$ = "READQ"
       CALL FMTFind("XXXXS009")
       CALL FmtFindFLD("ID1"): RSET UTXT$(FFLD.UTXT) = LTRIM$(STR$(1))
       CALL FmtFindFLD("ID2"): RSET UTXT$(FFLD.UTXT) = LTRIM$(STR$(FF1HDR.MAXID))
    END IF
    CURFMT = FMT001: FMT = FMTTAB(CURFMT)
    CALL FmtFindFLD("NMCT")
    LSET UTXT$(FFLD.UTXT) = STR$(FF1HDR.MAXID)
    CALL WinDisplayDATA(FMT001): CALL FmtDREST

150 ' Get File Info
    FMT = FMTTAB(FMT002)
    CALL FmtDREST
    FF = FMT.BGNFLD
    WHILE FF > 0
       FFLD = FLDTAB(FF)
       IF FFLD.UTXT > 0 THEN
          X$ = UTXT$(FFLD.UTXT)
          SELECT CASE FFLD.Name
            CASE "FSET": LSET X$ = FDFSETUP.Name
            CASE "DESC": LSET X$ = FDFSETUP.DESC
            CASE "SUNM": LSET X$ = FNX$(ENV.FSETUP)
            CASE "SUCT": LSET X$ = FNX$(FDFHdr.SETUPS)
            CASE "NAME": LSET X$ = FTYPE$(1)
            CASE "ADDR": LSET X$ = FTYPE$(2)
            CASE "OTHR": LSET X$ = FTYPE$(3)
            CASE "STAT": LSET X$ = FSTAT$
          END SELECT
          IF FAMOKAY THEN
             SELECT CASE FFLD.Name
                CASE "NVER": RSET X$ = LTRIM$(STR$(FF1HDR.REVLVL AND 255))
                CASE "NLUP": LSET X$ = XSCRNDate$(FF1HDR.UPDTMD, FF1HDR.UPDTY)
                CASE "NREC": X& = FFRBA&(FF1HDR.MAXID, 1): RSET X$ = LTRIM$(STR$(X&))
                CASE "AVER": RSET X$ = LTRIM$(STR$(FF2HDR.REVLVL AND 255))
                CASE "ALUP": LSET X$ = XSCRNDate$(FF2HDR.UPDTMD, FF2HDR.UPDTY)
                CASE "AREC": X& = FFRBA&(FF2HDR.HIREC, 1): RSET X$ = LTRIM$(STR$(X&))
                CASE "AFRE": X& = FFRBA&(FF2HDR.FREECNT, 1): RSET X$ = LTRIM$(STR$(X&))
                CASE "OVER": RSET X$ = LTRIM$(STR$(FF3HDR.REVLVL AND 255))
                CASE "OLUP": LSET X$ = XSCRNDate$(FF3HDR.UPDTMD, FF3HDR.UPDTY)
                CASE "OREC": X& = FFRBA&(FF3HDR.HIREC, 1): RSET X$ = LTRIM$(STR$(X&))
                CASE "OFRE": X& = FFRBA&(FF3HDR.FREECNT, 1): RSET X$ = LTRIM$(STR$(X&))
             END SELECT
          END IF
          LSET UTXT$(FFLD.UTXT) = X$
       END IF
       FF = FLDCHN(FF).FWD
    WEND

159 ' Show File Info
    CLOSE : ENV.FFOPEN = 0
    CALL WinPREP(2)
    IF WIN.WFMT <> FMT002 THEN CALL WinCLR
    CALL WinFORMAT(FMT002)
    CALL WinSWITCH(1)
    IF FAMOKAY THEN
       IF (FF1HDR.REVLVL > 255) OR (FF2HDR.REVLVL > 255) OR (FF3HDR.REVLVL > 255) THEN
          CALL PutMSG("XXXXM073")
       END IF
    END IF

200 ' Get Processing Option
    CALL PutMSG("")
    IF FUN$ <> "" THEN CALL OptHILITE(FUN$, FUN$, 1)
    CALL GetKEY
    IF LEN(A$) < 2 THEN CALL ErrBEEP(0): FUN = 0: GOTO 200:  ELSE FUN = A
    SELECT CASE FUN
       CASE 59: FUN$ = "F1"
       CASE 60: FUN$ = "F2"
       CASE 62: FUN$ = "F4"
       CASE 64: FUN$ = "F6"
       CASE 65: FUN$ = "F7"
       CASE 66: CALL DOSShell: FAMOKAY = 0: GOTO 40
       CASE 67: FUN$ = "F9"
       CASE 94: FUN$ = "F1"
       CASE 95: FUN$ = "F2"
       CASE 109: FUN$ = "F6": PGMRTRN$ = "FHSRPTS"
       CASE ELSE: CALL ErrBEEP(0): FUN = 0: GOTO 200
    END SELECT
    WIN.OFMT = FMT001: CALL OptHILITE(FUN$, FUN$, 6)
    SELECT CASE FUN$
       CASE "F1": GOTO 300
       CASE "F2": GOTO 400
       CASE "F4": GOTO 500
       CASE "F6": GOTO 1000
       CASE "F7": GOTO 600
       CASE "F9": GOTO 250
    END SELECT
    GOTO 200

250 ' Return to FAMMENU.BAS
    IF (ERID(0) > 0) AND (NOT WorkSAVED) THEN
       CALL PutMSG("VLDTM021")
       IF A$ <> SNGLKEY$(1) THEN 200
    END IF
    CALL FamCLOSE
    ERASE RFLTH, RFLBL$
    CALL PgmPREP(PGMRTRN$)
    IF (NOT OKAY) OR (A = 27) THEN LSET PGMRTRN$ = "FHSMENU": GOTO 40
    CALL RptFREE(RRPT$)
    CHAIN PGMRTRN$

300 ' Change FILE Parameters
    IF A = 94 THEN LSET PGMRTRN$ = "FHSFILE": GOTO 250
    CALL WinSWITCH(2)
    CALL FDFSelect
    CALL WinSWITCH(1): READ$ = "READ"
    FAMOKAY = 0: GOTO 50

400 ' Change PRINTER Parameters
    IF A = 95 THEN LSET PGMRTRN$ = "FHSPRTC": GOTO 250
    CALL WinSWITCH(2)
    CALL PDFSelect
    CALL WinSWITCH(1)
    GOTO 50

500 ' Process Request to Update Options
    CALL PutMSG("VLDTM014")
    SELECT CASE A$
       CASE CHR$(27): GOTO 200
       CASE "1": RRPT$ = "PVLD"
       CASE "2": RRPT$ = "DVLD"
       CASE "3": RRPT$ = "MDAT"
       CASE ELSE: GOTO 500
    END SELECT
    CALL WinSWITCH(2): CALL WinCLR: OptTABLE.CL = 0
    CALL RptOptDISPLAY(RRPT$)
    CALL RptOptUPDATE(RRPT$)
    GOTO 159

600 ' Create SELECT.WRK of Error ID's
    IF ERID(0) = 0 THEN
       CALL PutMSG("VLDTM022")
       GOTO 200
    END IF
    CALL SWRKCreate(ERID(), FF1HDR.MAXID, 0)
    WorkSAVED = -1
    GOTO 200

1000 : 'PROCESS Validation Request
     IF NOT FAMOKAY THEN CALL PutMSG("VLDTM001"): GOTO 200
     IF FF1HDR.MAXID = 0 THEN CALL PutMSG("VLDTM015"): GOTO 200
1005 ' Get Type Validation to Perform
     CALL PutMSG("VLDTM014")
     SELECT CASE A$
        CASE CHR$(27): GOTO 200
        CASE "1": RRPT$ = "PVLD": EFMT$ = "VLDTS011"
        CASE "2": RRPT$ = "DVLD": EFMT$ = "VLDTS012"
        CASE "3": RRPT$ = "MDAT": EFMT$ = "VLDTS013"
        CASE ELSE: GOTO 1005
     END SELECT
     CALL RptLOAD(RRPT$): IF NOT OKAY THEN 200
     IF PGMRTRN$ = "FHSRPTS" THEN 250
     CALL WinSWITCH(2): CALL WinCLR
     CALL FMTFind(EFMT$): CALL WinDisplayLABELS(0)
     IF RRPT$ = "MDAT" THEN GOTO 1015
     CALL FMTFind("VLDTS019"): FMT019 = CURFMT
     CALL FmtDREST: CALL WinFORMAT(0)
1010 ' Check for Records to Be Examined
     CALL PutMSG("VLDTM016")
     ALLREC = (A$ <> "2")
     IF A$ = CHR$(27) THEN GOTO 159
     REDIM ERRCNT(12), ScrnERRCNT(12), ERID(FF1HDR.MAXID)
     REDIM SELSET(FF1HDR.MAXID)
     IF NOT ALLREC THEN
        CALL SWRKBldList(SELSET(), SMAX, 1)
        IF NOT OKAY THEN GOTO 1010
        SYSVAR.NVAR1 = SMAX: SYSVAR.NVAR2 = FF1HDR.MAXID
        CALL PutMSG("VLDTM017")
        IF A$ = CHR$(27) THEN GOTO 159
        CURFMT = FMT019: FMT = FMTTAB(CURFMT)
        ELSE
        SMAX = FF1HDR.MAXID
        X = SMAX: WHILE X > 0: SELSET(X) = X: X = X - 1: WEND
     END IF
1015 ' GO Print Report
     SELECT CASE RRPT$
        CASE "PVLD": CALL PtrCHECK
        CASE "DVLD": CALL DataCHECK
        CASE "MDAT": CALL MatchCHECK
     END SELECT
     WorkSAVED = (ERID(0) = 0)
     REDIM SELSET(0)
     GOTO 40

REM $DYNAMIC
SUB BldMatchDEF
SHARED DFNUM(), DFPP(), RFABR$(), MFCNT, SLINE$
      Page.DFCNT = 0: RFCNT = 9
      REDIM DFNUM(10), DFPP(10)
      REDIM RFSEP(RFCNT), RFLTH(RFCNT), RFLBL$(RFCNT), RFABR$(RFCNT)
      CALL RptOptGET("MNGP", MINGAP)
      IF (NOT OKAY) OR (MINGAP = 0) THEN MINGAP = 2
      CALL RptOptGET("MXGP", MAXGAP)
      IF (NOT OKAY) OR (MAXGAP = 0) THEN MAXGAP = 5

58500 ' Select Report Fields - Based Upon Report Options
      OKAY = 0: Z = 0
      IF (Page.BWIDTH = 0) THEN
         BWIDTH = PDFForms.FWIDTH - PDFForms.IMARGIN - PDFForms.OMARGIN
         ELSE
         BWIDTH = Page.BWIDTH
      END IF
      DFNUM(1) = 1: Z = 1
      FOR X = 1 TO MFCNT
         Z = Z + 1: DFNUM(Z) = (SX(X) \ 256) + 1
      NEXT X
      Page.DFCNT = Z: Page.PFCNT = 0

58510 ' Create Report Definition Tables
      X = 1: RFNAME$ = "08001": GOSUB 58515     'COUNT
      X = 2: RFNAME$ = "01022": GOSUB 58515     'Surname
      X = 3: RFNAME$ = "01021": GOSUB 58515     'Given Name
      X = 4: RFNAME$ = "01030": GOSUB 58515     'Birth Date
      X = 5: RFNAME$ = "02020": GOSUB 58515     'Marriage Date
      X = 6: RFNAME$ = "01040": GOSUB 58515     'Death Date
      X = 7: RFNAME$ = "01032": GOSUB 58515     'Birth Place
      X = 8: RFNAME$ = "01042": GOSUB 58515     'Death Place
      X = 9: RFNAME$ = "01015": GOSUB 58515     'Gender
      GOTO 58525

58515 ' Process Report Field
      CALL GetRFLD(RFNAME$, FX)
      IF RFSEP(X) = 0 THEN RFSEP(X) = MINGAP
      IF FOUND THEN
         RDFFLD = RFLDS(FX)
         IF RFLTH(X) = 0 THEN RFLTH(X) = ASC(RDFFLD.RLTH)
         ALTH = ASC(RDFFLD.ALTH): LLTH = ASC(RDFFLD.LLTH)
         HLTH = ASC(RDFFLD.HLTH)
         RFABR$(X) = MID$(RFLD$(FX), 1, ALTH)
         RFLBL$(X) = MID$(RFLD$(FX), ALTH + LLTH + 1, HLTH)
      END IF
      RETURN

58525 ' Check Report Width
      GOSUB 58530
      IF RWIDTH <= BWIDTH THEN 59540
      GOTO 58535

58530 ' Determine Required Width
      RWIDTH = 0: SWIDTH = 0: FX = 0
      WHILE FX < Page.DFCNT
          FX = FX + 1: X = DFNUM(FX): FSEP = RFSEP(X)
          IF (X > 0) AND (FSEP > 0) THEN
             Y = RFLTH(X)
             IF FX > 1 THEN SWIDTH = SWIDTH + LEN(RFABR$(X)) + FSEP + Y
             Z = LEN(RFLBL$(X)): IF Y < Z THEN Y = Z
             IF PRT.OPT > 1 THEN RWIDTH = RWIDTH + 2 + Y
          END IF
      WEND
      IF SWIDTH < 78 THEN SWIDTH = 78
      RETURN

58535 ' Check Report Width
      IF PRT.OPT > 1 THEN
         RPWIDTH = RWIDTH + PDFForms.IMARGIN + PDFForms.OMARGIN
         PWIDTH = BWIDTH + PDFForms.IMARGIN + PDFForms.OMARGIN
         IF PWIDTH < RPWIDTH THEN
            SYSVAR.NVAR1 = PWIDTH: SYSVAR.NVAR2 = RPWIDTH
            CALL PutMSG("XXXXM009")
            IF A$ <> SNGLKEY$(1) THEN OKAY = 0: GOTO 58560
            BWIDTH = RWIDTH
         END IF
         ELSE
         BWIDTH = 0: GOTO 58560
      END IF

59540 ' Adjust to Fit
      IF PRT.OPT = 1 THEN GOTO 58560
      IF NOT (RWIDTH < BWIDTH) THEN 59550
      GAPS = 0: FX = Page.PFCNT
      WHILE FX <= Page.DFCNT
          X = DFNUM(FX)
          IF (X > 0) AND (RFSEP(X) > 0) THEN
             GAPS = GAPS + 1
          END IF
          FX = FX + 1
      WEND
      GAP = (BWIDTH - RWIDTH) \ GAPS
      GAP1 = (BWIDTH - RWIDTH) - GAPS * GAP
      IF GAP >= (MAXGAP - MINGAP) THEN GAP = MAXGAP - MINGAP: GAP1 = 0
      FX = Page.PFCNT
      WHILE FX <= Page.DFCNT
          X = DFNUM(FX): FSEP = RFSEP(X)
          IF (X > 0) AND (FSEP > 0) THEN
             RFSEP(X) = FSEP + GAP - (GAP1 > 0): GAP1 = GAP1 - 1
          END IF
          FX = FX + 1
      WEND

59550 ' Determine Field Print Positions
      PP = 0: FX = 0
      WHILE FX < Page.DFCNT
          FX = FX + 1: X = DFNUM(FX): FSEP = RFSEP(X)
          IF (X > 0) AND (FSEP > 0) THEN
             FL = RFLTH(X): HL = LEN(RFLBL$(X))
             Z = FL: IF Z < HL THEN Z = HL
             PP = PP + FSEP + (Z + 1) \ 2 - (FL + 1) \ 2
             DFPP(FX) = PP
             PP = PP + Z
          END IF
      WEND

58560 ' Return to Calling Program
      PLINE$ = SPACE$(BWIDTH + PDFForms.IMARGIN + PDFForms.OMARGIN)
      SLINE$ = SPACE$(SWIDTH)
      Page.BWIDTH = BWIDTH
      OKAY = -1
END SUB

REM $STATIC
SUB BldMatchTITLE
SHARED DFNUM(), DFPP()
995   ' Build Title Line
      TL$ = SPACE$(Page.BWIDTH)
      FOR FX = 1 TO Page.DFCNT: X = DFNUM(FX): PP = DFPP(FX)
         IF (X > 0) AND (PP > 0) THEN
            FL = RFLTH(X): HL = LEN(RFLBL$(X))
            IF HL > 0 THEN
               IF (RPTOPTION.FREEFORM = 0) OR (FX <= Page.PFCNT) THEN
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
         HTFLINE$(HTFLMAX) = "T10" + RTRIM$(TL$)
      END IF
      OKAY = -1
END SUB

REM $DYNAMIC
SUB ChkForPAUSE
SHARED XRID$, UPDTFile, NOPAUSE
    IF A = 27 THEN EXIT SUB
    A$ = INKEY$
    IF A$ <> "" THEN
       CALL PutMSG("VLDTM012")
       COLOR FG(5), BG(5): LOCATE 25, 74, 0: PRINT XRID$;
       CALL GetKEY
       NOPAUSE = ((LEN(A$) = 2) AND (A = 81))
       IF A <> 27 THEN
          CALL PutMSG("VLDTM013")
          ELSE
          UPDTFile = 0
       END IF
    END IF
END SUB

SUB DataCHECK
SHARED LastRID, XRID$, NOPAUSE, ALLREC, SMAX
SHARED FMT019, ERRCNT(), ScrnERRCNT(), ERID(), SELSET()

1100 ' Prepare for Data Validation
     REDIM BYR(FF1HDR.MAXID), SPID(50)

1105 ' Open files for processing
     SYSVAR.SUBJECT = FDFSETUP.Name
     CALL RptOPEN("DVLD"): IF NOT OKAY THEN 1190
     CALL FamOPEN("READ"): IF NOT OKAY THEN 1190
     IF (CRID < 1) OR (CRID > MAXID) THEN CRID = 1
     LastRID = 0: XRID$ = SPACE$(5): NOPAUSE = 0
     FSNDX$ = SPACE$(4): SNDX$ = FSNDX$: Surnm$ = SPACE$(LEN(FF1REC.Surnm))
     ENV.ScrnCLR = -1
     CALL PrintOPEN: IF NOT OKAY THEN 1190
     PRT.FWIDTH = Page.WIDTH
     DIM RTYPE$(7)
     FOR RT = 1 TO 7
         X$ = "RT" + FNX$(RT): CALL GetRVAR(X$, X)
         IF FOUND THEN RTYPE$(RT) = RVAR$(X) + " "
     NEXT RT
     DIM OTYPE$(1)
     CALL GetRVAR("ADDRESS", X): IF FOUND THEN OTYPE$(1) = RVAR$(X) + " "
     DIM DTYPE$(4)
     FOR DT = 1 TO 4
         X$ = "DT" + FNX$(DT): CALL GetRVAR(X$, X)
         IF FOUND THEN DTYPE$(DT) = RVAR$(X) + " "
     NEXT DT
     CALL RptOptGET("CKND", X): VALNMDT = (X = 1)
     CALL RptOptGET("CKMD", X): VALMYDT = (X = 1)
     CALL RptOptGET("CKAD", X): VALADDT = (X = 1)
     CALL RptOptGET("CKOD", X): VALMSDT = (X = 1)
     CHKMREC = (VALADDT + VALMYDT)
     VALIDATES = (VALNMDT OR VALMYDT OR VALADDT OR VALMSDT)
     CALL RptOptGET("CKAG", X): CHKAGES = (X = 1)
     IF CHKAGES THEN
        CALL RptOptGET("CKMA", X): CHKMAGE = (X = 1)
        CALL RptOptGET("CKSA", X): CHKSPAGE = (X = 1)
        CALL RptOptGET("CKAB", X): CHKPAGE = (X = 1)
     END IF
     CALL RptOptGET("ESBY", X): ESTIMBYR = (X = 1)
     IF ESTIMBYR THEN
        CALL RptOptGET("ESMA", ESTMAGE)
        CALL RptOptGET("ESPA", ESTPAGE)
     END IF
     CALL RptOptGET("CKUP", X):
     CHKSPAR = (X = 4): CHKPMAR = (X = 1) OR CHKSPAR
     CALL RptOptGET("CKBM", X): CHKBLKMAR = (X = 1)
     CALL RptOptGET("CKFS", X): CHKSURNM = (X = 1)
     IF CHKSURNM THEN
        CALL RptOptGET("CKSN", X): CHKSNDX = (X = 1)
        CHKSMNM = NOT CHKSNDX
     END IF
     CALL RptOptGET("CKSX", X): CHKSEX = (X = 1)
     IF CHKSEX THEN
        CALL RptOptGET("CKPS", X): CHKPSEX = (X = 1)
        CALL RptOptGET("CKSS", X): CHKSPSEX = (X = 1)
     END IF
     CALL RptOptGET("HIAG", MAXAGE)
     CALL RptOptGET("MDIF", MDIF)
     CALL RptOptGET("LOMA", AGEYMAR)
     CALL RptOptGET("HIMA", AGEOMAR)
     CALL RptOptGET("LOAB", AGEYPAR)
     CALL RptOptGET("HIAB", AGEOPAR)
     CHKMREC = CHKMREC OR CHKMAGE OR CHKSPAGE OR CHKPMAR OR CHKBLKMAR OR CHKSPSEX
     BLDBYTBL = CHKAGES OR CHKMAGE OR CHKSPAGE OR CHKPAGE OR CHKPSEX OR CHKSPSEX
     XSEX$ = " ": MALE$ = FGENDR$(1): FEMALE$ = FGENDR$(2): UNKNOWN$ = FGENDR$(3)
     FATHER$ = RELATION$("L", MALE$, -1, -1, RRELCD$())
     MOTHER$ = RELATION$("L", FEMALE$, -1, -1, RRELCD$())
     CYR = VAL(RIGHT$(DATE$, 4)): NUMED$ = SPACE$(6)

     ' Print Validation Criteria in Report
     IF PRT.OPT > 1 THEN
        PRT.SECTION = "0": CALL PrintHEAD
        PRT.SECTION = "W": CALL PrintTITLE
        IF VALIDATES <> 0 THEN PRT.SECTION = "1": CALL PrintTITLE
        IF CHKAGES THEN
           SYSVAR.NVAR1 = MAXAGE
           PRT.SECTION = "2": CALL PrintTITLE
        END IF
        IF CHKMAGE THEN
           SYSVAR.NVAR1 = AGEYMAR: SYSVAR.NVAR2 = AGEOMAR
           PRT.SECTION = "3": CALL PrintTITLE
        END IF
        IF CHKSPAGE THEN
           SYSVAR.NVAR1 = MDIF
           PRT.SECTION = "4": CALL PrintTITLE
        END IF
        IF CHKPAGE THEN
           SYSVAR.NVAR1 = AGEYPAR: SYSVAR.NVAR2 = AGEOPAR
           PRT.SECTION = "5": CALL PrintTITLE
        END IF
        IF CHKPMAR THEN PRT.SECTION = "6": CALL PrintTITLE
        IF CHKBLKMAR THEN PRT.SECTION = "7": CALL PrintTITLE
        IF CHKSURNM THEN PRT.SECTION = "8": CALL PrintTITLE
        IF CHKSEX THEN PRT.SECTION = "9": CALL PrintTITLE
        IF CHKPSEX THEN PRT.SECTION = "A": CALL PrintTITLE
        IF CHKSPSEX THEN PRT.SECTION = "B": CALL PrintTITLE
        CALL WriteLINE
        PRT.SECTION = "X": CALL PrintTITLE
     END IF
     IF A = 27 THEN 1190

1110 ' Build BYR() Tables
     IF NOT BLDBYTBL THEN 1200
     IF CHKPSEX OR CHKSPSEX THEN
        ON LOCAL ERROR GOTO 1189
        SYSVAR.FILENAME = WRKFPFX$ + "FH$VDTMP"
        OPEN SYSVAR.FILENAME FOR BINARY ACCESS READ WRITE AS #5
     END IF
     CALL PutMSG("VLDTM018")
     FOR X = 1 TO FF1HDR.MAXID: CALL FF1GetRec(X)
         RSET XRID$ = FNX$(X)
         COLOR FG(5), BG(5): LOCATE 25, 74, 0: PRINT XRID$;
         A$ = INKEY$
         IF A$ <> "" THEN
            A = ASC(A$): IF A = 27 THEN ID = 0
         END IF
         BYR = FF1REC.BY: DYR = FF1REC.DY
         IF (BYR = 0) AND (ESTIMBYR > 0) AND (DYR > ESTAGE) AND (DYR < 9999) THEN
            BYR = -(DYR - ESTAGE)
         END IF
         BYR(X) = BYR
         IF CHKPSEX OR CHKSPSEX THEN PUT #5, X, FF1REC.SEX
     NEXT X: IF A$ = CHR$(27) THEN 1190
     GOTO 1200

1189 ' Error Opening FH$VLDT
     CALL ErrMessage("XXXXM097", ERR, ERL)
     RESUME 1190

1190 ' Early EXIT
     CALL FamCLOSE: EXIT SUB

1200 ' Begin Data Validation
     IDX = 0
1210 ' Select Next Name Record
     IDX = IDX + 1: IF IDX > SMAX THEN 1350
     ID = SELSET(IDX): RSET XRID$ = FNX$(ID)
     CALL ChkForPAUSE: IF A = 27 THEN 1350
     COLOR FG(5), BG(5): LOCATE 25, 74, 0: PRINT XRID$;
     CALL FF1GetRec(ID)
     BYR = FF1REC.BY: DYR = FF1REC.DY
     FID = FF1REC.FID: MID = FF1REC.MID
     IF (BYR = 0) AND BLDBYTBL THEN BYR = ABS(BYR(ID))
     SEX$ = FF1REC.SEX
     LSET Surnm$ = FF1REC.Surnm: CHILD = FF1REC.OLDCH
     MREC = FF1REC.SPOUSE
     IF (CHKSURNM AND CHKSNDX) AND (SEX$ = MALE$) AND (CHILD <> 0) THEN
        CALL ComputeSNDX(Surnm$, FSNDX$)
     END IF
     IF VALNMDT THEN
        CALL ValiDATE(FF1REC.BMD, FF1REC.BY)
        IF (NOT OKAY) THEN SYSVAR.SUBJECT = DTYPE$(3): CALL DataERR(1, 1)
        CALL ValiDATE(FF1REC.DMD, FF1REC.DY)
        IF (NOT OKAY) OR ((FF1REC.DY > 0) AND (FF1REC.DY < FF1REC.BY)) THEN
           SYSVAR.MVAR = XSCRNDate$(FF1REC.DMD, FF1REC.DY)
           SYSVAR.SUBJECT = DTYPE$(4): CALL DataERR(1, 1)
        END IF
     END IF
     IF CHKSEX THEN
        IF (SEX$ <> MALE$) AND (SEX$ <> FEMALE$) AND (SEX$ <> UNKNOWN$) THEN
           SYSVAR.MVAR = FF1REC.SEX: CALL DataERR(8, 9)
        END IF
     END IF
     IF CHKPSEX THEN
        IF FID > 0 THEN
           GET #5, FID, XSEX$
           IF XSEX$ <> MALE$ THEN
              SYSVAR.SUBJECT = FATHER$: SYSVAR.MVAR = XSEX$
              CALL DataERR(9, 10)
           END IF
        END IF
        IF MID > 0 THEN
           GET #5, MID, XSEX$
           IF XSEX$ <> FEMALE$ THEN
              SYSVAR.SUBJECT = MOTHER$: SYSVAR.MVAR = XSEX$
              CALL DataERR(9, 10)
           END IF
        END IF
     END IF
     IF CHKAGES AND (DYR < 9999) AND (BYR > 0) THEN
        IF DYR > 0 THEN AGE = DYR - BYR ELSE AGE = CYR - BYR
        IF (AGE < 0) OR (AGE > MAXAGE) THEN
           SYSVAR.NVAR1 = AGE: CALL DataERR(2, 2)
        END IF
     END IF
     IF CHKPAGE AND (BYR > 0) THEN
        PBYR = BYR(FID)
        IF (PBYR <> 0) AND (BYR > 0) THEN
           AGE = BYR - ABS(PBYR)
           IF ((AGE < AGEYPAR) OR (AGE > AGEOPAR)) THEN
              SYSVAR.SUBJECT = FATHER$: SYSVAR.NVAR1 = AGE
              CALL DataERR(4, 5)
           END IF
        END IF
        PBYR = BYR(MID)
        IF (PBYR <> 0) AND (BYR > 0) THEN
           AGE = BYR - ABS(PBYR)
           IF ((AGE < AGEYPAR) OR (AGE > AGEOPAR)) THEN
              SYSVAR.SUBJECT = MOTHER$: SYSVAR.NVAR1 = AGE
              CALL DataERR(4, 5)
           END IF
        END IF
     END IF
     IF (NOT CHKMREC) OR (MREC = 0) THEN 1300
     RT = 2: MCNT = 0
     WHILE (MREC <> 0) AND (MCNT < 50)
        MCNT = MCNT + 1: CALL FF3GetRec(MREC)
        IF FF3SPOUSE.SP1ID = ID THEN
           SPID = FF3SPOUSE.SP2ID: MREC = FF3SPOUSE.SP1NXT
           ELSE
           SPID = FF3SPOUSE.SP1ID: MREC = FF3SPOUSE.SP2NXT
        END IF
        GOSUB 1250
     WEND
     IF (NOT CHKPMAR) AND ((NOT CHKSURNM) OR (SEX$ <> MALE$)) THEN 1300
     WHILE CHILD > 0
        CALL FF1GetRec(CHILD)
        FID = FF1REC.FID: MID = FF1REC.MID: SYSVAR.NVAR1 = CHILD
        IF ID = FID THEN
           PID = MID: CHILD = FF1REC.FCH: SYSVAR.MVAR = MOTHER$
           ELSE
           PID = FID: CHILD = FF1REC.MCH: SYSVAR.MVAR = FATHER$
        END IF
        IF CHKPMAR AND (CHKSPAR OR (PID <> 0)) THEN
           FOUND = 0: X = 1
           WHILE (X < (MCNT + 1)) AND (NOT FOUND)
              FOUND = (SPID(X) = PID)
              IF NOT FOUND THEN X = X + 1
           WEND
           IF NOT FOUND THEN CALL DataERR(5, 6)
        END IF
        IF CHKSURNM AND (SEX$ = MALE$) AND ((ASC(FF1REC.STATUS) AND 1) = 0) THEN
           IF CHKSNDX THEN CALL ComputeSNDX(FF1REC.Surnm, SNDX$)
           IF ((CHKSMNM AND (Surnm$ <> FF1REC.Surnm)) OR (CHKSNDX AND (FSNDX$ <> SNDX$))) THEN
              CALL DataERR(7, 8)
           END IF
        END IF
     WEND
     GOTO 1300

1250 ' Check Marriage Record Dates, Age at Marriage & Spouse Sex Code
     SPID(MCNT) = SPID: SPBYR = BYR(SPID): MADR = FF3SPOUSE.LOC
     D1MD = FF3SPOUSE.D1MD: D1Y = FF3SPOUSE.D1Y
     D2MD = FF3SPOUSE.D2MD: D2Y = FF3SPOUSE.D2Y
     IF VALMYDT THEN OT = 0: GOSUB 1380
     IF CHKBLKMAR THEN
        IF (SPID = 0) AND (D1MD = 0) AND (D1Y = 0) AND (D2MD = 0) AND (D2Y = 0) AND (FF3SPOUSE.MLOC = 0) AND (FF3SPOUSE.COM = 0) AND (FF3SPOUSE.LOC = 0) THEN
           SYSVAR.NVAR1 = MCNT: CALL DataERR(6, 7)
        END IF
     END IF
     IF CHKMAGE AND (BYR > 0) AND (D1Y > 0) THEN
        MAGE = D1Y - BYR
        IF (MAGE < AGEYMAR) OR (MAGE > AGEOMAR) THEN
           SYSVAR.NVAR1 = SPID: SYSVAR.NVAR2 = MAGE
           CALL DataERR(3, 3)
        END IF
     END IF
     IF ((ID <= SPID) OR (SELSET(SPID) = 0)) THEN
        IF CHKSPAGE AND (BYR > 0) AND (SPBYR > 0) THEN
           MAGE = ABS(BYR - SPBYR)
           IF (MAGE > MDIF) THEN
              SYSVAR.NVAR1 = SPID: SYSVAR.NVAR2 = MAGE
              CALL DataERR(3, 4)
           END IF
        END IF
        IF CHKSPSEX AND (SPID > 0) THEN
           GET #5, SPID, XSEX$
           IF SEX$ = XSEX$ THEN SYSVAR.NVAR1 = SPID: CALL DataERR(8, 11)
        END IF
     END IF
     IF VALADDT AND (MADR <> 0) THEN
        ADR = MADR: GOSUB 1370
     END IF
     RETURN
    
1300 ' Validate Other Dates
     CALL FF1GetRec(ID): ADR = FF1REC.ADRS
     IF VALADDT AND (ADR <> 0) THEN RT = 1: GOSUB 1370
     IF VALMSDT THEN
        OREC = FF1REC.EDUC: IF OREC <> 0 THEN RT = 3: GOSUB 1360
        OREC = FF1REC.WORK: IF OREC <> 0 THEN RT = 4: GOSUB 1360
        OREC = FF1REC.MIL: IF OREC <> 0 THEN RT = 5: GOSUB 1360
        OREC = FF1REC.HLTH: IF OREC <> 0 THEN RT = 6: GOSUB 1360
        OREC = FF1REC.EVENT: IF OREC <> 0 THEN RT = 7: GOSUB 1360
     END IF

1350 ' Go Get  Next ID
     IF A = 27 THEN 2900
     IF (IDX < SMAX) THEN 1210
     GOTO 1900

1360 ' Validate Miscellaneous Record Dates
     WHILE OREC <> 0: CALL FF3GetRec(OREC)
        IF RT = 7 THEN
           D1MD = FF3EVENT.D1MD: D1Y = FF3EVENT.D1Y
           OREC = FF3EVENT.NXT
           ELSE
           D1MD = FF3Pntrs.D1MD: D1Y = FF3Pntrs.D1Y
           D2MD = FF3Pntrs.D2MD: D2Y = FF3Pntrs.D2Y
           OREC = FF3Pntrs.NXT
        END IF
        OT = 0: GOSUB 1380
        IF VALADDT AND (RT < 7) THEN
           ADR = FF3Pntrs.LOC: IF ADR <> 0 THEN GOSUB 1370
        END IF
     WEND
     RETURN

1370 ' Validate Address Dates
     OT = 1
     WHILE ADR > 0: CALL FF2GetRec(ADR)
        D1MD = FF2PFX.D1MD: D1Y = FF2PFX.D1Y
        D2MD = FF2PFX.D2MD: D2Y = FF2PFX.D2Y
        GOSUB 1380
        ADR = FF2PFX.NXT
     WEND
     RETURN

1380 ' Validate Begin, End Dates
     CALL ValiDATE(D1MD, D1YR)
     IF (NOT OKAY) OR (D1YR = 9999) THEN
        SYSVAR.SUBJECT = RTYPE$(RT) + OTYPE$(OT) + DTYPE$(1)
        CALL DataERR(1, 1)
     END IF
     IF RT <> 7 THEN
        CALL ValiDATE(D2MD, D2YR)
        IF (NOT OKAY) OR (D2Y > 0) AND (D2Y < D1Y) THEN
           SYSVAR.SUBJECT = RTYPE$(RT) + OTYPE$(OT) + DTYPE$(2)
           CALL DataERR(1, 1)
        END IF
     END IF
     RETURN

1900 ' Return from VALIDATION operation
     IF ERRCNT(12) > 0 THEN
        SYSVAR.NVAR1 = ERRCNT(12): SYSVAR.NVAR2 = ERRCNT(0)
        CALL PutMSG("VLDTM019")
     END IF

2500 ' END OF REPORT
     IF A = 27 THEN 2900
     IF PRT.OPT = 1 THEN 2950
     LSET PLINE$ = " ": CALL WriteLINE
     SYSVAR.NVAR1 = ERRCNT(12): SYSVAR.NVAR2 = ERRCNT(0)
     PRT.SECTION = "Y": CALL PrintTITLE
     IF ERRCNT(12) > 0 THEN
        CALL WriteLINE
        PRT.SECTION = "Z": CALL PrintTITLE
        FOR ET = 1 TO 11
            IF ERRCNT(ET) > 0 THEN
               EMSG$ = "ET" + ZFILL$(ET, 2)
               CALL GetRVAR(EMSG$, X)
               IF FOUND THEN
                  RSET NUMED$ = FNX$(ERRCNT(ET))
                  LSET PLINE$ = NUMED$ + "  " + RVAR$(X)
                  CALL PrintLINE: IF A = 27 THEN ET = 11
               END IF
            END IF
        NEXT ET
     END IF
     GOTO 2950

2900 ' Early Termination
     IF PRT.OPT = 1 THEN
        IF ID > 0 THEN
           SYSVAR.NVAR1 = ID
           CALL PutMSG("VLDTM010")
           ELSE
           CALL PutMSG("VLDTM011")
        END IF
        ELSE
        CALL RptCLOSE
     END IF
     GOTO 2990

2950 ' Validation Procedure Complete
     IF PRT.OPT = 1 THEN
        CALL PutMSG("VLDTM020")
        ELSE
        CALL RptCLOSE
     END IF

2990 ' Close Files
     CALL FamCLOSE
     CALL PutMSG("")
     IF CHKPSEX OR CHKSPSEX THEN
        ON LOCAL ERROR GOTO 2995
        SYSVAR.FILENAME = WRKFPFX$ + "FH$VDTMP"
        KILL SYSVAR.FILENAME
     END IF
     GOTO 2999
2995 RESUME 2999
2999 ' END OF ROUTINE

END SUB

SUB DataERR (ERRGrp, EMNUM)
SHARED NOPAUSE, XRID$, LastRID, FMT019, ERRCNT(), ScrnERRCNT(), ERID()
   IF A <> 27 THEN CALL ChkForPAUSE
   IF A = 27 THEN EXIT SUB
   IF EMNUM > 0 THEN ERRCNT(EMNUM) = ERRCNT(EMNUM) + 1
   IF ERRGrp <> 0 THEN
      ERRCNT(12) = ERRCNT(12) + 1
      ScrnERRCNT(ERRGrp) = ScrnERRCNT(ERRGrp) + 1
   END IF
   EFLD$ = "EC" + ZFILL$(ERRGrp, 2)
   IF CURFMT <> FMT019 THEN CURFMT = FMT019: FMT = FMTTAB(CURFMT)
   CALL FmtFindFLD(EFLD$)
   IF FOUND THEN RSET UTXT$(FFLD.UTXT) = FNX$(ScrnERRCNT(ERRGrp))
   CALL FmtFindFLD("ETOT")
   IF FOUND THEN
      RSET UTXT$(FFLD.UTXT) = FNX$(ERRCNT(12))
      CALL WinDisplayDATA(0)
   END IF
   XRID = VAL(XRID$)
   IF XRID > 0 THEN
      IF ERID(XRID) = 0 THEN ERID(0) = ERID(0) + 1
      ERID(XRID) = ERID(XRID) + 1
   END IF
   IF XRID <> LastRID THEN
      ERRCNT(0) = ERRCNT(0) + 1: LastRID = XRID
      LSET PLINE$ = XRID$
      ELSE
      LSET PLINE$ = ""
   END IF
   EMNM$ = "EM" + ZFILL$(EMNUM, 2)
   CALL GetRVAR(EMNM$, X)
   IF FOUND THEN
      ERRMSG$ = RVAR$(X)
      ELSE
      ERRMSG$ = "Error Message " + EMNM$ + " Not Found"
   END IF
   MID$(PLINE$, 8, 255) = ERRMSG$
   CALL HTFEXPAND(PLINE$, 1)
   COLOR FG(5), BG(5): LOCATE 25, 1, 0: PRINT XRID$;
   COLOR FG(1), BG(1): LOCATE 25, 7, 0: PRINT MID$(PLINE$, 8, 70); TAB(79);
   IF PRT.OPT <> 1 THEN
      CALL PrintLINE
      ELSE
      IF NOT NOPAUSE THEN
         CALL ErrBEEP(0): CALL GetKEY
         NOPAUSE = ((LEN(A$) = 2) AND (A = 81))
         IF NOT Page.WAIT THEN Page.WAIT = NOPAUSE
      END IF
   END IF
   EXIT SUB
END SUB

REM $STATIC
SUB MatchCHECK
SHARED DFNUM(), DFPP(), MFCNT, ERID(), RFABR$(), SLINE$, NOPAUSE

    CDFM$ = CONFIG.DFM
    STATFMT = CURFMT: CALL FmtDREST
    ON LOCAL ERROR GOTO MCError
MCGetIndex:
    CALL FMTFind("VLDTS005"): IF NOT FOUND THEN GOTO MC190
    CALL FmtFindFLD("DSNM"): Y = FFLD.UTXT: LSET UTXT$(Y) = FT$(4)
    CALL MsgFORMAT: IF A = 27 THEN GOTO MC190
    LSET FT$(4) = UTXT$(Y)
    CALL IndexFILE("READQ")
    IF NOT OKAY THEN GOTO MCGetIndex
    MFCNT = SFCNT - 1
    IF MFCNT = 0 THEN
       CALL PutMSG("VLDTM024")
       GOTO MC190
    END IF
    REDIM LMDATA$(MFCNT), MDATA$(MFCNT)
    FMT = FMTTAB(STATFMT)
    FOR MF = 1 TO MFCNT
        X = SX(MF) \ 256
        IF (X = 1) THEN SNCHK = -1
        IF (X = 3) OR (X = 5) THEN BDCHK = -1
        IF (X = 4) THEN MDCHK = -1
        IF (X = 6) OR (X = 7) THEN BPCHK = -1
        FLDNM$ = "FLD" + LTRIM$(STR$(X))
        CALL FmtFindFLD(FLDNM$)
        IF FOUND THEN LSET UTXT$(FFLD.UTXT) = "*"
    NEXT MF
    CALL WinDisplayDATA(STATFMT)
    REDIM SELSET1(FF1HDR.MAXID), SELSET2(FF1HDR.MAXID)
    CALL MatchGrp(1, SELSET1(), G1): IF NOT OKAY THEN GOTO MC190
    CALL MatchGrp(2, SELSET2(), G2): IF NOT OKAY THEN GOTO MC190
    FullFILE = (G1 = 1) OR (G2 = 1)
    TS = ENV.SortOPTS AND 3: DS = (ENV.SortOPTS \ 4) AND 7
    SELECT CASE DS          ' Determine Date Format and Date Sort Key Length
       CASE 0: CONFIG.DFM = "3": DSKL = 10
       CASE 1: CONFIG.DFM = "1": DSKL = 10
       CASE 2: CONFIG.DFM = "2": DSKL = 10
       CASE 3: CONFIG.DFM = "3": DSKL = 7
       CASE 4: CONFIG.DFM = "1": DSKL = 5
       CASE 5: CONFIG.DFM = "3": DSKL = 4
       CASE 6: CONFIG.DFM = "1": DSKL = 2
       CASE 7: CONFIG.DFM = "2": DSKL = 2
    END SELECT

    ' Open files for processing
    SYSVAR.SUBJECT = FDFSETUP.Name
    CALL RptOPEN("MDAT"): IF NOT OKAY THEN GOTO MC190
    CALL FamOPEN("READ"): IF NOT OKAY THEN GOTO MC190
    CALL RptOptGET("MNMC", MinMATCH)
    IF (NOT OKAY) OR (MinMATCH = 0) THEN MinMATCH = 2
    MaxMATCH = MinMATCH: MaxMATCH2 = MinMATCH
    CALL FmtFindFLD("MNMC")
    IF FOUND THEN
       RSET UTXT$(FFLD.UTXT) = LTRIM$(STR$(MinMATCH))
       CALL WinDisplayDATA(0)
    END IF
    CALL FmtFindFLD("DCNT"): IF FOUND THEN DCFX = FFLD.UTXT
    CALL FmtFindFLD("GCNT"): IF FOUND THEN GCFX = FFLD.UTXT
    ENV.ScrnCLR = -1: NOPAUSE = 0
    CALL PrintOPEN: IF NOT OKAY THEN GOTO MC190
    IF BDCHK THEN BDT$ = SPACE$(11): DDT$ = BDT$
    IF MDCHK THEN MDT$ = SPACE$(11)
    IF BPCHK THEN BPL$ = SPACE$(22): DPL$ = BPL$
    BDPCHK = (BDCHK OR BPCHK)
    CALL BldMatchDEF: IF PRT.OPT > 1 THEN CALL BldMatchTITLE: NOPAUSE = -1
    GOTO MCBeginReport

MC190: ' Early Exit
    CONFIG.DFM = CDFM$
    CALL FamCLOSE: EXIT SUB

MCBeginReport:
    SlctCNT = 0: SecSlctCNT = 0
    SX1 = 0: M1CNT = 0: SX2 = 0: M2CNT = 0: S2 = 0: GCNT = 0
    IDX = 0: XTEXT$ = SPACE$(57)
    REDIM ERID(FF1HDR.MAXID), MatchRID(MaxMATCH), MatchRID2(MaxMATCH2)
    IF (DCFX <> 0) THEN LSET UTXT$(DCFX) = LTRIM$(STR$(SlctCNT))
    IF (GCFX <> 0) THEN LSET UTXT$(GCFX) = LTRIM$(STR$(GCNT))
    CALL WinDisplayDATA(STATFMT)
    PRT.SECTION = "1"

MCGetNextRID:
    IDX = IDX + 1: IF IDX > FF1HDR.MAXID THEN GOTO MCChkLastGroup
    COLOR FG(5), BG(5): LOCATE 25, 74, 0: PRINT SFILL$(IDX, 5);
    A$ = INKEY$
    IF A$ <> "" THEN
       A = ASC(RIGHT$(A$, 1)): OKAY = (A <> 27)
       IF NOT OKAY THEN GOTO MCFinish
       NOPAUSE = 0
       SYSVAR.NVAR1 = IDX: SYSVAR.NVAR2 = FF1HDR.MAXID
       CALL PutMSG("VLDTM025")
    END IF
    ID = NDXID(IDX)
    IF (ID = 0) OR (SELSET1(ID) = 0) AND (SELSET2(ID) = 0) THEN GOTO MCGetNextRID
    CALL FF1GetRec(ID)
    IF SNCHK THEN
       IF ENV.HUSBSURNM THEN
          LSET XNAME$ = FMTNAME$(0)    ' Have Function Build SURNM$
          ELSE
          Surnm$ = FF1REC.Surnm
       END IF
    END IF
    IF BDPCHK THEN CALL GetBDEvents(BDT$, BPL$, DDT$, DPL$, AGE$)
    IF MDCHK THEN
       IF FF1REC.SPOUSE <> 0 THEN
          CALL FF3GetRec(FF1REC.SPOUSE)
          LSET MDT$ = LTRIM$(XSCRNDate$(FF3SPOUSE.D1MD, FF3SPOUSE.D1Y))
          ELSE
          LSET MDT$ = ""
       END IF
    END IF
    MATCH = -1
    FOR MF = 1 TO MFCNT
        X = SX(MF) \ 256
        SELECT CASE X
           CASE 1: LSET XTEXT$ = TextDATA$(Surnm$): MKL = 20
           CASE 2: LSET XTEXT$ = TextDATA$(FF1REC.Given): MKL = 30
           CASE 3: LSET XTEXT$ = MID$(BDT$, 1, DSKL)
           CASE 4: LSET XTEXT$ = MID$(MDT$, 1, DSKL)
           CASE 5: LSET XTEXT$ = MID$(DDT$, 1, DSKL)
           CASE 6: LSET XTEXT$ = TextDATA$(BPL$): MKL = 22
           CASE 7: LSET XTEXT$ = TextDATA$(DPL$): MKL = 22
           CASE 8: LSET XTEXT$ = FF1REC.SEX
        END SELECT
        IF (X > 2) AND (X < 6) THEN
           SKL = DSKL
           ELSE
           IF X < 8 THEN
              IF TS = 2 THEN
                 SKL = 4
                 ELSE
                 SKL = (SX(MF) AND 127)
                 MKL = MKL - 4 * (TS > 2)
                 IF SKL > MKL THEN SKL = MKL
              END IF
              ELSE
              SKL = 1
           END IF
        END IF
        MDATA$(MF) = RTRIM$(XTEXT$)
        IF MATCH THEN
           MATCH = (MID$(MDATA$(MF), 1, SKL) = MID$(LMDATA$(MF), 1, SKL))
        END IF
    NEXT MF
    IF MATCH THEN
       M1CNT = M1CNT - (SELSET1(ID) > 0)
       M2CNT = M2CNT - (SELSET2(ID) > 0)
       IF SELSET1(ID) > 0 THEN
          IF ((M1CNT + S2) < MinMATCH) OR (M2CNT = 0) THEN
             SX1 = SX1 + 1
             IF SX1 > MaxMATCH THEN
                MaxMATCH = MaxMATCH + 10
                REDIM PRESERVE MatchRID(MaxMATCH)
             END IF
             MatchRID(SX1) = ID
             ELSE
             ERID(ID) = 1
          END IF
       END IF
       IF SELSET2(ID) > 0 THEN
          IF ((M1CNT + S2) < MinMATCH) OR (M1CNT = 0) THEN
             IF SELSET1(ID) = 0 THEN
                S2 = S2 + 1: SX2 = SX2 + 1
                IF SX2 > MaxMATCH2 THEN
                   MaxMATCH2 = MaxMATCH2 + 10
                   REDIM PRESERVE MatchRID2(MaxMATCH2)
                END IF
                MatchRID2(SX2) = ID
             END IF
             ELSE
             IF SELSET1(ID) = 0 THEN
                ERID(ID) = 2: S2 = S2 + 1
             END IF
          END IF
       END IF
       GOTO MCGetNextRID
    END IF
MCChkLastGroup:
    IF ((M1CNT + S2) >= MinMATCH) AND (M1CNT > 0) AND (M2CNT > 0) THEN
       X = 0: WHILE X < SX1: X = X + 1: ERID(MatchRID(X)) = 1: WEND
       X = 0: WHILE X < SX2: X = X + 1: ERID(MatchRID2(X)) = 2: WEND
       SlctCNT = SlctCNT + M1CNT
       SecSlctCNT = SecSlctCNT + S2
       GCNT = GCNT + 1: LP = 2
       FOR X = 1 TO Page.DFCNT
           DFN = DFNUM(X)
           IF PRT.OPT > 1 THEN
              PP = DFPP(X): FLTH = RFLTH(DFN)
              IF DFN > 1 THEN
                 MID$(PLINE$, PP, FLTH) = LMDATA$(X - 1)
                 ELSE
                 IF FullFILE THEN
                    L = (FLTH \ 2) + 3
                    MID$(PLINE$, PP, FLTH) = SFILL$(M1CNT + S2, L)
                    ELSE
                    L = (FLTH - 1) \ 2
                    MID$(PLINE$, PP, FLTH) = SFILL$(M1CNT, L) + "," + LTRIM$(STR$(S2))
                 END IF
              END IF
           END IF
           IF X > 1 THEN
              IF X > 2 THEN MID$(SLINE$, LP, 1) = ";": LP = LP + 2
              MID$(SLINE$, LP, 255) = RFABR$(DFN) + LMDATA$(X - 1)
              LP = LEN(RTRIM$(SLINE$)) + 1
           END IF
       NEXT X
       IF DCFX <> 0 THEN
          IF FullFILE THEN
             LSET UTXT$(DCFX) = LTRIM$(STR$(SlctCNT + SecSlctCNT))
             ELSE
             LSET UTXT$(DCFX) = LTRIM$(STR$(SlctCNT)) + "," + LTRIM$(STR$(SecSlctCNT))
          END IF
       END IF
       IF (GCFX <> 0) THEN LSET UTXT$(GCFX) = LTRIM$(STR$(GCNT))
       CALL WinDisplayDATA(STATFMT)
       IF PRT.OPT > 1 THEN
          CALL PrintLINE: IF A = 27 THEN GOTO MCFinish
       END IF
       COLOR FG(5), BG(5): LOCATE 25, 1, 0: PRINT RFABR$(1);
       IF FullFILE THEN
          PRINT LTRIM$(STR$(M1CNT + S2));
          ELSE
          PRINT LTRIM$(STR$(M1CNT)) + "," + LTRIM$(STR$(S2));
       END IF
       COLOR FG(1), BG(1): X25 = POS(0)
       PRINT (LEFT$(SLINE$, 79 - X25));
       IF NOT NOPAUSE THEN
          COLOR FG(5), BG(5): LOCATE 25, 76, 0: PRINT "...";
          CALL ErrBEEP(0): CALL GetKEY
          NOPAUSE = ((LEN(A$) = 2) AND (A = 81))
          IF NOT Page.WAIT THEN Page.WAIT = NOPAUSE
          COLOR FG(1), BG(1): LOCATE 25, 76, 0
          PRINT (MID$(SLINE$, 77 - X25, 3));
       END IF
       LSET SLINE$ = ""
       IF A = 27 THEN GOTO MCFinish
    END IF
    IF IDX > FF1HDR.MAXID THEN GOTO MCFinish
    FOR MF = 1 TO MFCNT: LMDATA$(MF) = MDATA$(MF): NEXT MF
    M1CNT = -(SELSET1(ID) > 0): SX1 = 0
    M2CNT = -(SELSET2(ID) > 0): SX2 = 0
    IF M1CNT > 0 THEN
       MatchRID(1) = ID: SX1 = 1
       ELSE
       MatchRID2(1) = ID: SX2 = 1
    END IF
    S2 = SX2
    GOTO MCGetNextRID

MCError:
    ' Error During Processing
    CALL ErrMessage("VLDTM027", ERR, ERL)
    RESUME MCFinish

MCFinish:
    ERID(0) = SlctCNT
    IF PRT.OPT > 1 THEN
       PRT.SECTION = "2"
       CALL WriteLINE
       IF FullFILE THEN
          SYSVAR.MVAR = LTRIM$(STR$(SlctCNT + SecSlctCNT))
          ELSE
          SYSVAR.MVAR = "(" + LTRIM$(STR$(SlctCNT)) + "," + LTRIM$(STR$(SecSlctCNT)) + ")=" + LTRIM$(STR$(SlctCNT + SecSlctCNT))
       END IF
       SYSVAR.NVAR1 = GCNT
       CALL PrintTITLE
       CALL RptCLOSE: IF PRT.AGAIN THEN GOTO MCBeginReport
       ELSE
       CALL PutMSG("VLDTM026")
    END IF
    CONFIG.DFM = CDFM$
    CALL FamCLOSE

END SUB

SUB MatchGrp (N, ID(), G)

GetMatchGrp:
    SYSVAR.NVAR1 = N: OKAY = -1
    CALL PutMSG("VLDTM030")
    IF A = 27 THEN OKAY = 0: EXIT SUB
    G = VAL(A$): IF (G < 1) OR (G > 4) THEN G = 1
    SELECT CASE G
       CASE 1: FOR X = 1 TO FF1HDR.MAXID: ID(X) = 1: NEXT X
               XCNT = FF1HDR.MAXID
       CASE 2: CALL RWRKOpen("READ", 0, 0)
               IF OKAY THEN
                  CALL RWRKGet("BASE", RC, PTR&)
                  CALL PutMSG("XXXXM064")
                  WHILE (0 < RC) AND (RC < 9)
                     IF WFRELREC.ID > 0 THEN ID(WFRELREC.ID) = 1
                     CALL RWRKGet("WRKREC", RC, PTR&)
                  WEND
                  CLOSE #6
               END IF
       CASE 3: CALL SWRKBldList(ID(), X, 0)
       CASE 4: CALL FMTFind("XXXXS009")
               CALL FmtFindFLD("ID1"): X1 = FFLD.UTXT
               CALL FmtFindFLD("ID2"): X2 = FFLD.UTXT
               CALL MsgFORMAT(0): IF OKAY THEN OKAY = (A <> 27)
               IF OKAY THEN
                  ID1 = VAL(UTXT$(X1)): ID2 = VAL(UTXT$(X2))
                  IF (ID2 < 1) OR (ID2 > FF1HDR.MAXID) THEN ID2 = FF1HDR.MAXID
                  IF (ID1 < 1) OR (ID1 > ID2) THEN ID1 = 1
                  IF (ID1 > 1) THEN
                     XID1 = 1: XID2 = ID1 - 1
                     ELSE
                     IF ID2 < FF1HDR.MAXID THEN
                        XID1 = ID2 + 1: XID2 = FF1HDR.MAXID
                     END IF
                  END IF
                  IF XID1 > 0 THEN
                     RSET UTXT$(X1) = LTRIM$(STR$(XID1))
                     RSET UTXT$(X2) = LTRIM$(STR$(XID2))
                  END IF
                  FOR X = ID1 TO ID2: ID(X) = 1: NEXT X
                  XCNT = ID2 - ID1 + 1
               END IF
    END SELECT
    IF NOT OKAY THEN GOTO GetMatchGrp
    IF (1 < N) AND (N < 4) THEN
       XCNT = 0
       FOR X = 1 TO FF1HDR.MAXID
           IF ID(X) <> 0 THEN ID(X) = 1: XCNT = XCNT + 1
       NEXT X
    END IF
    IF CURFMT <> WIN.WFMT THEN
       CURFMT = WIN.WFMT: FMT = FMTTAB(CURFMT)
    END IF
    FLDNM$ = "CNT" + LTRIM$(STR$(N))
    CALL FmtFindFLD(FLDNM$)
    IF FOUND THEN
       IF FFLD.UTXT > 0 THEN
          RSET UTXT$(FFLD.UTXT) = LTRIM$(STR$(XCNT))
          CALL WinDisplayDATA(0)
       END IF
    END IF
    ID(0) = XCNT: OKAY = -1

END SUB

REM $DYNAMIC
SUB PtrCHECK
SHARED CRID, LastRID, UPDTFile, XRID$, NOPAUSE, ALLREC, SMAX, FMT019
SHARED ERRCNT(), ERID(), SELSET()

P1000: ' Validate File Headers and Pointers
     CALL PutMSG("VLDTM002")
     IF ASC(A$) = 27 THEN 1090
     UPDTFile = (A$ = SNGLKEY$(1))
     ON LOCAL ERROR GOTO 1050
     GOTO P1100

1050 ' Error During Processing
     ' SYSVAR.NVAR1 = CRID: CALL PutMSG("CRID=[NVAR1]..."): CALL GetRESP
     CALL ErrMessage("VLDTM027", ERR, ERL)
     RESUME 1090

1090 ' EARLY EXIT
     CALL FamCLOSE
     EXIT SUB

P1100: ' Open files for processing
     SYSVAR.SUBJECT = FDFSETUP.Name
     CALL RptOPEN("PVLD"): IF NOT OKAY THEN 1090
     IF UPDTFile THEN CALL FamOPEN("VLDT") ELSE CALL FamOPEN("READ")
     IF NOT OKAY THEN 1090
     MAXID = FF1HDR.MAXID: HIREC2 = FF2HDR.HIREC
     HIREC3& = FFRBA&(FF3HDR.HIREC, 1)
     IF (CRID < 1) OR (CRID > MAXID) THEN CRID = 1
     REDIM PTRCK1(MAXID), PTRCK2(HIREC2)
     CALL PtrCHK3("OPEN", HIREC3&, 0): IF NOT OKAY THEN 1090
     REDIM ERID(MAXID)
     MAXAGE = 9999: CYR = VAL(RIGHT$(DATE$, 4))
     CALL RptOptGET("CKDD", X): IF FOUND THEN FixMissingDD = (X = 1)
     IF FixMissingDD THEN
        CALL RptOptGET("HIAG", X): IF FOUND THEN MAXAGE = X
     END IF
     LastRID = 0: XRID$ = SPACE$(5): NOPAUSE = 0
     ENV.ScrnCLR = -1
     CALL PrintOPEN: IF NOT OKAY THEN 1090
     PRT.FWIDTH = Page.WIDTH: IF LEN(PLINE$) < 100 THEN PLINE$ = SPACE$(100)
     CALL PutMSG("VLDTM003")
     DIM RTYPE$(11)
     FOR RT = 1 TO 10
         X$ = "RT" + ZFILL$(RT, 2): CALL GetRVAR(X$, X)
         IF FOUND THEN RTYPE$(RT) = RVAR$(X) + " "
     NEXT RT
     OLDTYPES$ = "1234F6789ABCDE": MID$(OLDTYPES$, 6, 1) = CHR$(11)
     DIM OLDTYPE$(14)
     FOR N = 1 TO 14: OLDTYPE$(N) = MID$(OLDTYPES$, N, 1): NEXT N
     IF PRT.OPT > 1 THEN
        PRT.SECTION = " ": CALL PrintHEAD
        PRT.SECTION = "E": CALL PrintTITLE
     END IF
     IDX = 0

1150 ' Check for User Input
     IF A <> 27 THEN CALL ChkForPAUSE
     IF A = 27 THEN GOTO P2900

1155 ' Process Next Name Record
     IDX = IDX + 1: IF IDX > SMAX THEN 2000
     CRID = SELSET(IDX)
     RSET XRID$ = FNX$(CRID)
     COLOR FG(5), BG(5): LOCATE 25, 74, 0: PRINT XRID$;
     CHG1 = 0: CALL FF1GetRec(CRID)

1160 ' Validate Name record Header Info
     IF FF1REC.RTYPE = "1" THEN LSET FF1REC.RTYPE = CHR$(1): CHG1 = -1
     IF FF1REC.RTYPE <> CHR$(1) THEN
        SYSVAR.NVAR1 = ASC(FF1REC.RTYPE): CALL PtrERR(1, 1)
        LSET FF1REC.RTYPE = CHR$(1): CHG1 = -1
     END IF
     IF FF1REC.RID <> CRID THEN
        SYSVAR.NVAR1 = FF1REC.RID: CALL PtrERR(1, 2)
        FF1REC.RID = CRID: CHG1 = -1
     END IF
     IF FixMissingDD THEN
        IF (FF1REC.BY <> 0) AND (FF1REC.DY = 0) AND (FF1REC.DMD = 0) THEN
           SYSVAR.NVAR1 = CYR - FF1REC.BY
           IF SYSVAR.NVAR1 > MAXAGE THEN
              CALL PtrERR(1, 12)
              FF1REC.DY = 9999: CHG1 = -1
           END IF
        END IF
     END IF
     MRID = FF1REC.MID: FRID = FF1REC.FID
     FSRID = FF1REC.FCH: MSRID = FF1REC.MCH
     CHRID = FF1REC.OLDCH: CHCNT = 0: LASTCH = 0
     SPREC = FF1REC.SPOUSE: SPCNT = 0: LASTSP = 0
     COMREC = FF1REC.COM: ADREC = FF1REC.ADRS
     LSET MISCPTR$ = F1MISCPTR$: BLOC = FF1REC.BLOC
     IF (MRID < 0) OR (MRID > MAXID) THEN
        SYSVAR.NVAR1 = MRID: CALL PtrERR(1, 3)
        FF1REC.MID = 0: MRID = 0: CHG1 = -1
     END IF
     IF (FRID < 0) OR (FRID > MAXID) THEN
        SYSVAR.NVAR1 = FRID: CALL PtrERR(1, 4)
        FF1REC.FID = 0: FRID = 0: CHG1 = -1
     END IF
     IF (CHRID < 0) OR (CHRID > MAXID) THEN
        SYSVAR.NVAR1 = CHRID: CALL PtrERR(1, 5)
        FF1REC.OLDCH = 0: CHRID = 0: CHG1 = -1
     END IF
     IF (FSRID < 0) OR (FSRID > MAXID) THEN
        SYSVAR.NVAR1 = FSRID: CALL PtrERR(1, 6)
        FF1REC.FCH = 0: FSRID = 0: CHG1 = -1
     END IF
     IF (MSRID < 0) OR (MSRID > MAXID) THEN
        SYSVAR.NVAR1 = MSRID: CALL PtrERR(1, 7)
        FF1REC.MCH = 0: MSRID = 0: CHG1 = -1
     END IF
     IF (FRID = 0) AND (FSRID > 0) THEN
        SYSVAR.NVAR1 = FSRID: CALL PtrERR(1, 8)
        FF1REC.FCH = 0: FSRID = 0: CHG1 = -1
     END IF
     IF (MRID = 0) AND (MSRID > 0) THEN
        SYSVAR.NVAR1 = MSRID: CALL PtrERR(1, 9)
        FF1REC.MCH = 0: MSRID = 0: CHG1 = -1
     END IF
     PTR& = FFRBA&(SPREC, 1)
     IF (PTR& > HIREC3&) THEN
        SYSVAR.NVAR1 = PTR&: CALL PtrERR(1, 10)
        FF1REC.SPOUSE = 0: CHG1 = -1
     END IF
     X = INSTR(FF1REC.Given, CHR$(0)): Y = X
     IF X > 0 THEN
        CALL PtrERR(1, 70)
        IF UPDTFile THEN
           WHILE (X > 0) AND (X <= LEN(FF1REC.Given))
                 MID$(FF1REC.Given, X, 1) = " "
                 X = INSTR(X + 1, FF1REC.Given, CHR$(0))
           WEND
        END IF
        CHG1 = -1
     END IF
     X = INSTR(FF1REC.Surnm, CHR$(0))
     IF X > 0 THEN
        IF Y = 0 THEN CALL PtrERR(1, 70)
        IF UPDTFile THEN
           WHILE (X > 0) AND (X <= LEN(FF1REC.Surnm))
                 MID$(FF1REC.Surnm, X, 1) = " "
                 X = X + 1
           WEND
        END IF
        CHG1 = -1
     END IF
     IF (UPDTFile AND CHG1) THEN
        CALL FF1PutRec: CHG1 = 0
     END IF

1180 ' Check FID
     IF FRID = 0 THEN GOTO P1190
     IF FRID = MRID THEN
        SYSVAR.NVAR1 = FF1REC.FID: CALL PtrERR(1, 11)
     END IF
     CALL FF1GetRec(FRID)
     PTRCK1(CRID) = PTRCK1(CRID) OR 1

P1190: ' Check MID
     IF MRID = 0 THEN GOTO P1200
     CALL FF1GetRec(MRID)
     PTRCK1(CRID) = PTRCK1(CRID) OR 4

P1200: ' Check Child Chain
     IF A = 27 THEN GOTO P2900
     IF CHRID = 0 THEN GOTO P1300
     IF (CHRID < 0) OR (CHRID > MAXID) THEN
        SYSVAR.NVAR1 = CHCNT: SYSVAR.NVAR2 = LASTCH: CALL PtrERR(1, 14)
        GOTO 1240
     END IF
     CALL FF1GetRec(CHRID): CHCNT = CHCNT + 1
     IF CRID = FF1REC.FID THEN
        LASTCH = CHRID: CHRID = FF1REC.FCH
        PTRCK1(LASTCH) = PTRCK1(LASTCH) OR 2
        IF (CHRID > 0) AND (CHRID <= MAXID) THEN
           IF ((PTRCK1(CHRID) AND 2) = 2) THEN 1229
        END IF
        GOTO P1200
     END IF
     IF CRID = FF1REC.MID THEN
        LASTCH = CHRID: CHRID = FF1REC.MCH
        PTRCK1(LASTCH) = PTRCK1(LASTCH) OR 8
        IF (CHRID > 0) AND (CHRID <= MAXID) THEN
           IF ((PTRCK1(CHRID) AND 8) = 8) THEN 1229
        END IF
        GOTO P1200
     END IF
     SYSVAR.NVAR1 = CHCNT: SYSVAR.NVAR2 = CHRID
     SYSVAR.NVAR3 = FF1REC.MID: SYSVAR.NVAR4 = FF1REC.FID
     SYSVAR.NVAR5 = LASTCH: CALL PtrERR(1, 15)
     GOTO 1240
1229 SYSVAR.NVAR1 = CHCNT: SYSVAR.NVAR2 = LASTCH
     CALL PtrERR(1, 16)

1240 ' Update OLDCH/LASTCH PTRs
     IF UPDTFile THEN
        IF LASTCH = 0 THEN
           CALL FF1GetRec(CRID): FF1REC.OLDCH = 0: CALL FF1PutRec
           ELSE
           CALL FF1GetRec(LASTCH)
           IF FF1REC.MID = CRID THEN FF1REC.MCH = 0
           IF FF1REC.FID = CRID THEN FF1REC.FCH = 0
           CALL FF1PutRec
        END IF
     END IF

P1300: ' Check Birth/Death dates and Location Record
     IF A = 27 THEN GOTO P2900
     CALL FF1GetRec(CRID)
     IF FF1REC.BY < 0 THEN
        SYSVAR.NVAR1 = FF1REC.BY: CALL PtrERR(1, 17)
        FF1REC.BY = 0: CHG1 = -1
     END IF
     IF FF1REC.DY < 0 THEN
        SYSVAR.NVAR2 = FF1REC.DY: CALL PtrERR(1, 18)
        FF1REC.DY = 0: CHG1 = -1
     END IF
     IF (UPDTFile AND CHG1) THEN CALL FF1PutRec: CHG1 = 0
P1310: ' Check Birth/Death Place Name Record(s)
     IF BLOC = 0 THEN 1400
     PTR& = FFRBA&(BLOC, 1): SYSVAR.NVAR1 = PTR&
     IF (PTR& > HIREC3&) THEN
        CALL PtrERR(1, 19): GOTO P1350
     END IF
     CALL PtrCHK3("GET", PTR&, PTRCK3)
     IF PTRCK3 <> 0 THEN
        SYSVAR.NVAR2 = PTRCK3: CALL PtrERR(7, 20)
        GOTO P1350
     END IF
     CHG3 = 0: CALL FF3GetRec(BLOC)
     IF FF3PFX.RTYPE = OLDTYPE$(4) THEN
        FF3PFX.RTYPE = CHR$(4): CHG3 = -1
     END IF
     IF (FF3PFX.RTYPE <> CHR$(4)) AND (FF3PFX.RTYPE <> CHR$(12)) THEN
        SYSVAR.NVAR2 = ASC(FF3PFX.RTYPE): CALL PtrERR(6, 21)
        GOTO P1350
     END IF
     IF BLOC <> FF3PFX.RNO THEN
        SYSVAR.NVAR2 = FFRBA&(FF3PFX.RNO, 1): CALL PtrERR(6, 23)
        FF3PFX.RNO = BLOC: CHG3 = -1
     END IF
     IF FF3PFX.RTYPE = CHR$(4) THEN
        IF FF3PLACE.SRNO <> CRID THEN
           SYSVAR.NVAR2 = FFRBA&(FF3PLACE.SRNO, 1): CALL PtrERR(6, 22)
           GOTO P1350
        END IF
        IF FF3PLACE.SRTYPE = OLDTYPE$(1) THEN
           FF3PLACE.SRTYPE = CHR$(1): CHG3 = -1
        END IF
        IF FF3PLACE.SRTYPE <> CHR$(1) THEN
           SYSVAR.NVAR2 = ASC(FF3PLACE.SRTYPE): CALL PtrERR(6, 24)
           FF3PLACE.SRTYPE = CHR$(1): CHG3 = -1
        END IF
        IF (UPDTFile AND CHG3) THEN CALL FF3PutRec: CHG3 = 0
        PTRCK3 = 4 * 256 + 64 + 1
        CALL PtrCHK3("PUT", PTR&, PTRCK3)
        GOTO 1400
     END IF
     ' Check Long Place Name Record(S)
     IF FF3PFX.RTYPE = CHR$(12) THEN
        IF FF3PLACE2.SRNO <> CRID THEN
           SYSVAR.NVAR2 = FFRBA&(FF3PLACE2.SRNO, 1): CALL PtrERR(6, 22)
           GOTO P1350
        END IF
        IF FF3PLACE2.SRTYPE <> CHR$(1) THEN
           SYSVAR.NVAR2 = ASC(FF3PLACE2.SRTYPE): CALL PtrERR(6, 24)
           FF3PLACE2.SRTYPE = CHR$(1): CHG3 = -1
        END IF
        IF (UPDTFile AND CHG3) THEN CALL FF3PutRec: CHG3 = 0
        PTRCK3 = 12 * 256 + 64 + 1
        CALL PtrCHK3("PUT", PTR&, PTRCK3)
        BLOC = FF3PLACE2.PNEXT
        GOTO P1310
     END IF

P1350: ' ZERO Out Birth/Death PLACE Pointer
     IF UPDTFile THEN
        CALL FF1GetRec(CRID): FF1REC.BLOC = 0: CALL FF1PutRec
     END IF

1400 ' Check Individual's Comments
     IF A = 27 THEN GOTO P2900
     IF COMREC = 0 THEN 1450
     SRTYPE = 1: SRNO = CRID: COMCNT = 0: GOSUB P1900

1450 ' Check Individual's Residence Records
     IF A = 27 THEN GOTO P2900
     IF ADREC = 0 THEN 1500
     ASRTYPE = 1: ASRNO = CRID: ADRCNT = 0: LastADR = 0: GOSUB 1800

1500 ' Check Spouse Records
     IF A = 27 THEN GOTO P2900
     IF SPREC = 0 THEN 1600 ELSE SPNUM = 0
     SYSVAR.NVAR1 = SPCNT + 1
     PTR& = FFRBA&(SPREC, 1)
     IF (PTR& > HIREC3&) THEN
        SYSVAR.NVAR2 = PTR&: CALL PtrERR(2, 25)
        GOTO 1570
     END IF
     CHG3 = 0: SPCNT = SPCNT + 1
     CALL FF3GetRec(SPREC)
     IF FF3PFX.RTYPE = OLDTYPE$(3) THEN
        LSET FF3PFX.RTYPE = CHR$(3): CHG3 = -1
     END IF
     IF FF3PFX.RTYPE <> CHR$(3) THEN
        SYSVAR.NVAR2 = ASC(FF3PFX.RTYPE): CALL PtrERR(2, 26)
        GOTO 1570
     END IF
     IF SPREC <> FF3PFX.RNO THEN
        SYSVAR.NVAR2 = FFRBA&(FF3PFX.RNO, 1): CALL PtrERR(2, 27)
        FF3PFX.RNO = SPREC: CHG3 = -1
     END IF
     IF (UPDTFile AND CHG3) THEN CALL FF3PutRec: CHG3 = 0

1520 ' Check Marriage/Divorce dates and Location Record
     IF A = 27 THEN GOTO P2900
     CALL FF3GetRec(SPREC)
     IF FF3SPOUSE.D1Y < 0 THEN
        SYSVAR.NVAR2 = FF3SPOUSE.D1Y: CALL PtrERR(2, 28)
        FF3SPOUSE.D1Y = 0: CHG3 = -1
     END IF
     IF FF3SPOUSE.D2Y < 0 THEN
        SYSVAR.NVAR2 = FF3SPOUSE.D2Y: CALL PtrERR(2, 29)
        FF3SPOUSE.D2Y = 0: CHG3 = -1
     END IF
     IF (UPDTFile AND CHG3) THEN CALL FF3PutRec
     CHG3 = 0: MLOC = FF3SPOUSE.MLOC
P1525: ' Check Place Name Record
     IF MLOC = 0 THEN 1550
     PTR& = FFRBA&(MLOC, 1)
     IF (PTR& > HIREC3&) THEN
        SYSVAR.NVAR2 = PTR&: CALL PtrERR(2, 30)
        GOTO 1535
     END IF
     CHG3 = 0: CALL FF3GetRec(MLOC)
     IF (FF3PFX.RTYPE <> CHR$(4)) AND (FF3PFX.RTYPE <> CHR$(12)) THEN
        SYSVAR.NVAR2 = ASC(FF3PFX.RTYPE): CALL PtrERR(6, 31)
        GOTO 1535
     END IF
     IF MLOC <> FF3PFX.RNO THEN
        SYSVAR.NVAR2 = PTR&: SYSVAR.NVAR3 = FFRBA&(FF3PFX.RNO, 1)
        CALL PtrERR(6, 32)
        FF3PFX.RNO = MLOC: CHG3 = -1
     END IF
     IF FF3PFX.RTYPE = OLDTYPE$(4) THEN LSET FF3PFX.RTYPE = CHR$(4): CHG3 = -1
     IF FF3PFX.RTYPE = CHR$(4) THEN
        IF FF3PLACE.SRTYPE = OLDTYPE$(3) THEN FF3PLACE.SRTYPE = CHR$(3): CHG3 = -1
        IF FF3PLACE.SRTYPE <> CHR$(3) THEN
           SYSVAR.NVAR2 = ASC(FF3PLACE.SRTYPE): CALL PtrERR(6, 33)
           FF3PLACE.SRTYPE = CHR$(3): CHG3 = -1
        END IF
        IF FF3PLACE.SRNO <> SPREC THEN
           SYSVAR.NVAR2 = FFRBA&(SPREC, 1): SYSVAR.NVAR3 = FFRBA&(FF3PLACE.SRNO, 1)
           CALL PtrERR(6, 34)
           GOTO 1535
        END IF
        IF (UPDTFile AND CHG3) THEN CALL FF3PutRec: CHG3 = 0
        PTRCK3 = 4 * 256 + 32 + 1
        CALL PtrCHK3("PUT", PTR&, PTRCK3)
        GOTO 1550
     END IF
     ' Check Long Place Name Record
     IF FF3PFX.RTYPE = CHR$(12) THEN
        IF FF3PLACE2.SRTYPE <> CHR$(3) THEN
           SYSVAR.NVAR2 = ASC(FF3PLACE2.SRTYPE): CALL PtrERR(6, 33)
           FF3PLACE2.SRTYPE = CHR$(3): CHG3 = -1
        END IF
        IF FF3PLACE2.SRNO <> SPREC THEN
           SYSVAR.NVAR2 = FFRBA&(SPREC, 1): SYSVAR.NVAR3 = FFRBA&(FF3PLACE2.SRNO, 1)
           CALL PtrERR(6, 34)
           GOTO 1535
        END IF
        IF (UPDTFile AND CHG3) THEN CALL FF3PutRec: CHG3 = 0
        PTRCK3 = 4 * 256 + 32 + 1
        CALL PtrCHK3("PUT", PTR&, PTRCK3)
        IF FF3PLACE2.PNEXT <> 0 THEN
           MLOC = FF3PLACE2.PNEXT: GOTO P1525
        END IF
        GOTO 1550
     END IF

1535 ' ZERO Out Spouse PLACE Pointer
     IF UPDTFile THEN
        CALL FF3GetRec(SPREC): FF3SPOUSE.MLOC = 0: CALL FF3PutRec
     END IF

1550 ' Get Next Spouse Record Number
     IF A = 27 THEN GOTO P2900
     CALL FF3GetRec(SPREC): SPNUM = 0
     SYSVAR.NVAR2 = FF3SPOUSE.SP1ID: SYSVAR.NVAR3 = FF3SPOUSE.SP2ID
     IF CRID = FF3SPOUSE.SP1ID THEN
        LASTSP = SPREC: SPREC = FF3SPOUSE.SP1NXT: SPNUM = 1
        GOTO 1560
     END IF
     IF CRID = FF3SPOUSE.SP2ID THEN
        LASTSP = SPREC: SPREC = FF3SPOUSE.SP2NXT: SPNUM = 2
        GOTO 1560
     END IF
     CALL PtrERR(2, 35)
     GOTO 1570
1560 LSP& = FFRBA&(LASTSP, 1)
     CALL PtrCHK3("GET", LSP&, PTRCK3)
     IF (PTRCK3 AND SPNUM) > 0 THEN
        CALL PtrERR(2, 36)
        GOTO 1570
     END IF
     PTRCK3 = (((PTRCK3 OR SPNUM) OR 64) OR 3 * 256)
     CALL PtrCHK3("PUT", LSP&, PTRCK3)
     IF (PTRCK3 AND 15) <> SPNUM THEN 1500 ' Skip COM & ADR chk for 2nd spouse
     COMREC = FF3SPOUSE.COM: ADREC = FF3SPOUSE.LOC
     IF COMREC <> 0 THEN SRTYPE = 3: SRNO = LASTSP: COMCNT = 0: GOSUB P1900
     IF ADREC > 0 THEN ASRTYPE = 3: ASRNO = LASTSP: ADRCNT = 0: GOSUB 1800
     GOTO 1500

1570 ' Terminate Spouse Records with Last one
     IF UPDTFile THEN
        IF LASTSP = 0 THEN
           CALL FF1GetRec(CRID)
           FF1REC.SPOUSE = 0: CALL FF1PutRec
           ELSE
           CALL FF3GetRec(LASTSP)
           IF CRID = FF3SPOUSE.SP1ID THEN FF3SPOUSE.SP1NXT = 0
           IF CRID = FF3SPOUSE.SP2ID THEN FF3SPOUSE.SP2NXT = 0
           CALL FF3PutRec
        END IF
     END IF

1600 ' Check Miscellaneous Info
     FOR MRTYPE = 6 TO 10
         ET = 6
         SELECT CASE MRTYPE
            CASE 6: MSCREC = FF1REC.EVENT: ET = 5
            CASE 7: MSCREC = FF1REC.EDUC
            CASE 8: MSCREC = FF1REC.WORK
            CASE 9: MSCREC = FF1REC.MIL
            CASE 10: MSCREC = FF1REC.HLTH
         END SELECT
         IF MSCREC <> 0 THEN
            MSCCNT = 0: LASTMSC = 0: LastADR = 0
            GOSUB 1700
         END IF
     NEXT MRTYPE
     IF A = 27 THEN GOTO P2900
     IF UPDTFile AND XCHANGED THEN
        CALL FF1GetRec(CRID): CALL FF1PutRec
     END IF
     GOTO 1150

1700 ' Check Miscellaneous Information records
     IF A = 27 THEN RETURN
     IF MSCREC = 0 THEN 1799
     SYSVAR.MVAR = RTYPE$(MRTYPE)
     MSCCNT = MSCCNT + 1
     SYSVAR.NVAR2 = MSCCNT
     PTR& = FFRBA&(MSCREC, 1)
     IF (PTR& > HIREC3&) THEN
        SYSVAR.NVAR3 = PTR&: CALL PtrERR(ET, 37)
        GOTO 1730
     END IF
     CHG3 = 0: CALL FF3GetRec(MSCREC)
     SYSVAR.MVAR = RTYPE$(MRTYPE)
     CALL PtrCHK3("GET", PTR&, PTRCK3)
     IF PTRCK3 <> 0 THEN
        SYSVAR.NVAR3 = PTRCK3: CALL PtrERR(7, 38)
        GOTO 1730
     END IF
     IF FF3PFX.RTYPE = OLDTYPE$(MRTYPE) THEN
        FF3PFX.RTYPE = CHR$(MRTYPE): CHG3 = -1
     END IF
     IF FF3PFX.RTYPE <> CHR$(MRTYPE) THEN
        SYSVAR.NVAR3 = ASC(FF3PFX.RTYPE): CALL PtrERR(ET, 39)
        GOTO 1730
     END IF
     IF MSCREC <> FF3PFX.RNO THEN
        SYSVAR.NVAR3 = FFRBA&(MSCREC, 1): SYSVAR.NVAR4 = FFRBA&(FF3PFX.RNO, 1)
        CALL PtrERR(ET, 40)
        GOTO 1730
     END IF
     SELECT CASE MRTYPE
        CASE 6: SRTYPE$ = FF3EVENT.SRTYPE: MRSRNO = FF3EVENT.SRNO
        CASE ELSE: SRTYPE$ = FF3Pntrs.SRTYPE: MRSRNO = FF3Pntrs.SRNO
     END SELECT
     IF SRTYPE$ = OLDTYPE$(1) THEN LSET SRTYPE$ = CHR$(1): CHG3 = -1
     IF SRTYPE$ <> CHR$(1) THEN
        SYSVAR.NVAR3 = ASC(SRTYPE$): CALL PtrERR(ET, 41)
        LSET SRTYPE$ = CHR$(1): CHG3 = -1
     END IF
     IF MRSRNO <> CRID THEN
        SYSVAR.NVAR3 = MRSRNO: CALL PtrERR(ET, 42)
        GOTO 1730
     END IF
     IF (UPDTFile AND CHG3) THEN
        SELECT CASE MRTYPE
           CASE 6: FF3EVENT.SRTYPE = SRTYPE$
           CASE ELSE: FF3Pntrs.SRTYPE = SRTYPE$
        END SELECT
        CALL FF3PutRec: CHG3 = 0
     END IF
     PTRCK3 = MRTYPE * 256 + 64 + 1
     CALL PtrCHK3("PUT", PTR&, PTRCK3)
     LASTMSC = MSCREC
     SELECT CASE MRTYPE
        CASE 6: MSCREC = FF3EVENT.NXT: COMREC = FF3EVENT.COM: ADREC = 0
                IF FF3EVENT.LPLNM <> 0 THEN
                   PTRCK3 = 12 * 256 + 64 + 1
                   PTR& = FFRBA&(FF3EVENT.LPLNM, 1)
                   CALL PtrCHK3("PUT", PTR&, PTRCK3)
                END IF
        CASE ELSE: MSCREC = FF3Pntrs.NXT: COMREC = FF3Pntrs.COM
                   ADREC = FF3Pntrs.LOC
     END SELECT
     IF COMREC <> 0 THEN
        SRTYPE = MRTYPE: SRNO = LASTMSC: COMCNT = 0: GOSUB P1900
     END IF
     IF ADREC > 0 THEN
        ASRTYPE = MRTYPE: ASRNO = LASTMSC: ADRCNT = 0: GOSUB 1800
     END IF
     GOTO 1700

1730 ' Stop Miscellaneous Records with Lastmsc
     IF UPDTFile THEN
        MSCCNT = MSCCNT - 1
        IF MSCCNT > 0 THEN
           CALL FF3GetRec(LASTMSC)
           SELECT CASE MRTYPE
              CASE 6: FF3EVENT.NXT = 0
              CASE ELSE: FF3Pntrs.NXT = 0
           END SELECT
           CALL FF3PutRec
           ELSE
           CALL FF1GetRec(CRID)
           SELECT CASE MRTYPE
              CASE 6: FF1REC.EVENT = 0
              CASE 7: FF1REC.EDUC = 0
              CASE 8: FF1REC.WORK = 0
              CASE 9: FF1REC.MIL = 0
              CASE 10: FF1REC.HLTH = 0
           END SELECT
           CALL FF1PutRec
        END IF
     END IF
1799 RETURN

1800 ' Check Address Records
     IF A = 27 THEN RETURN
     IF ADREC = 0 THEN 1899
     ADRCNT = ADRCNT + 1: CHG2 = 0
     SYSVAR.MVAR = RTYPE$(ASRTYPE)
     SYSVAR.NVAR1 = ADRCNT: SYSVAR.NVAR2 = ASRTYPE
     IF (ADREC < 0) OR (ADREC > HIREC2) THEN
        SYSVAR.NVAR3 = ADREC: CALL PtrERR(3, 43)
        GOTO 1830
     END IF
     CALL FF2GetRec(ADREC)
     IF PTRCK2(ADREC) <> 0 THEN
        SYSVAR.NVAR3 = PTRCK2(ADREC): CALL PtrERR(7, 44)
        GOTO 1830
     END IF
     IF FF2PFX.RTYPE = OLDTYPE$(2) THEN FF2PFX.RTYPE = CHR$(2): CHG2 = -1
     IF FF2PFX.RTYPE <> CHR$(2) THEN
        SYSVAR.NVAR3 = ASC(FF2PFX.RTYPE): CALL PtrERR(3, 45)
        FF2PFX.RTYPE = CHR$(2): CHG2 = -1
     END IF
     IF ADREC <> FF2PFX.RNO THEN
        SYSVAR.NVAR3 = ADREC: SYSVAR.NVAR4 = FF2PFX.RNO
        CALL PtrERR(3, 46)
        FF2PFX.RNO = ADREC: CHG2 = -1
     END IF
     IF FF2PFX.SRTYPE = OLDTYPE$(ASRTYPE) THEN LSET FF2PFX.SRTYPE = CHR$(ASRTYPE): CHG2 = -1
     IF FF2PFX.SRTYPE <> CHR$(ASRTYPE) THEN
        SYSVAR.NVAR3 = ASC(FF2PFX.SRTYPE): CALL PtrERR(3, 47)
        FF2PFX.SRTYPE = CHR$(ASRTYPE): CHG2 = -1
     END IF
     IF FF2PFX.SRNO <> ASRNO THEN
        SYSVAR.NVAR3 = FFRBA&(ASRNO, 1): SYSVAR.NVAR4 = FFRBA&(FF2PFX.SRNO, 1)
        CALL PtrERR(3, 48)
        GOTO 1830
     END IF
     IF (UPDTFile AND CHG2) THEN CALL FF2PutRec: CHG2 = 0
     IF ASRTYPE = 1 THEN PTRCK = 64 ELSE PTRCK = 32
     PTRCK2(ADREC) = 2 * 256 + PTRCK + 1
     LastADR = ADREC: ADREC = FF2PFX.NXT
     COMREC = FF2PFX.COM
     IF COMREC <> 0 THEN
        SRTYPE = 2: SRNO = LastADR
        COMCNT = 0: GOSUB P1900
     END IF
     GOTO 1800

1830 ' Stop Address Records with Lastadr
     IF UPDTFile THEN
        ADRCNT = ADRCNT - 1
        IF LastADR <> 0 THEN
           CALL FF2GetRec(LastADR): FF2PFX.NXT = 0: CALL FF2PutRec
           ELSE
           IF ASRTYPE = 1 THEN
              CALL FF1GetRec(CRID): FF1REC.ADRS = 0: CALL FF1PutRec
              ELSE
              CALL FF3GetRec(ASRNO)
              SELECT CASE ASRTYPE
                 CASE 3: FF3SPOUSE.LOC = 0
                 CASE ELSE: FF3Pntrs.LOC = 0
              END SELECT
              CALL FF3PutRec
           END IF
        END IF
     END IF
1899 RETURN

P1900: ' Check Comments
     IF A = 27 THEN RETURN
     SYSVAR.MVAR = RTYPE$(SRTYPE)
     IF COMREC = 0 THEN 1940
     COMCNT = COMCNT + 1
     SYSVAR.NVAR1 = COMCNT: SYSVAR.NVAR2 = SRTYPE
     PTR& = FFRBA&(COMREC, 1)
     IF (PTR& > HIREC3&) THEN
        SYSVAR.NVAR3 = PTR&: CALL PtrERR(4, 49)
        GOTO 1930
     END IF
     CHG3 = 0: CALL FF3GetRec(COMREC)
     SYSVAR.MVAR = RTYPE$(SRTYPE)
     IF COMCNT = 1 THEN
        COMTOT = FF3COMMENT.CCNT: LASTCOM = 0: FIRSTCOM = COMREC
     END IF
     CALL PtrCHK3("GET", PTR&, PTRCK3)
     IF PTRCK3 <> 0 THEN
        SYSVAR.NVAR3 = PTRCK3: CALL PtrERR(7, 50)
        GOTO 1930
     END IF
     IF FF3PFX.RTYPE = OLDTYPE$(5) THEN FF3PFX.RTYPE = CHR$(5): CHG3 = -1
     IF FF3PFX.RTYPE <> CHR$(5) THEN
        SYSVAR.NVAR3 = ASC(FF3PFX.RTYPE): CALL PtrERR(4, 51)
        GOTO 1930
     END IF
     IF COMREC <> FF3PFX.RNO THEN
        SYSVAR.NVAR3 = FFRBA&(COMREC, 1): SYSVAR.NVAR4 = FFRBA&(FF3PFX.RNO, 1)
        CALL PtrERR(4, 52)
        FF3PFX.RNO = COMREC: CHG3 = -1
     END IF
     IF FF3COMMENT.SRTYPE = OLDTYPE$(SRTYPE) THEN
        FF3COMMENT.SRTYPE = CHR$(SRTYPE): CHG3 = -1
     END IF
     IF FF3COMMENT.SRTYPE <> CHR$(SRTYPE) THEN
        SYSVAR.NVAR3 = ASC(FF3COMMENT.SRTYPE)
        CALL PtrERR(4, 53)
        FF3COMMENT.SRTYPE = CHR$(SRTYPE): CHG3 = -1
     END IF
     IF FF3COMMENT.SRNO <> SRNO THEN
        SYSVAR.NVAR3 = FFRBA&(SRNO, 1): SYSVAR.NVAR4 = FFRBA&(FF3COMMENT.SRNO, 1)
        CALL PtrERR(4, 54)
        GOTO 1930
     END IF
     IF FF3COMMENT.BPTR <> LASTCOM THEN
        SYSVAR.NVAR3 = FFRBA&(FF3COMMENT.BPTR, 1)
        SYSVAR.NVAR4 = FFRBA&(LASTCOM, 1)
        CALL PtrERR(4, 55)
        FF3COMMENT.BPTR = LASTCOM: CHG3 = -1
     END IF
     IF (UPDTFile AND CHG3) THEN CALL FF3PutRec: CHG3 = 0
     IF SRTYPE = 1 THEN PTRCK = 64 ELSE IF SRTYPE = 2 THEN PTRCK = 16 ELSE PTRCK = 32
     PTRCK3 = 5 * 256 + PTRCK + 1
     CALL PtrCHK3("PUT", PTR&, PTRCK3)
     LASTCOM = COMREC: COMREC = FF3COMMENT.FPTR
     GOTO P1900

1930 ' Stop comments with Lastcom
     COMCNT = COMCNT - 1
     IF UPDTFile THEN
        IF COMCNT >= 1 THEN
           CALL FF3GetRec(LASTCOM): FF3COMMENT.FPTR = 0: CALL FF3PutRec
           ELSE
           SELECT CASE SRTYPE
              CASE 1: CALL FF1GetRec(CRID): FF1REC.COM = 0: CALL FF1PutRec
              CASE 2: CALL FF2GetRec(SRNO): FF2PFX.COM = 0: CALL FF2PutRec
              CASE 3: CALL FF3GetRec(SRNO): FF3SPOUSE.COM = 0: CALL FF3PutRec
              CASE 6: CALL FF3GetRec(SRNO): FF3EVENT.COM = 0: CALL FF3PutRec
              CASE 11: CALL FF3GetRec(SRNO): FF3EVENT.COM = 0: CALL FF3PutRec
              CASE ELSE: CALL FF3GetRec(SRNO): FF3Pntrs.COM = 0: CALL FF3PutRec
           END SELECT
        END IF
     END IF

1940 ' Correct Comment Record Count if necessary
     IF ((COMCNT < 1) OR (COMCNT = COMTOT)) THEN 1949
     SYSVAR.NVAR3 = COMTOT: SYSVAR.NVAR4 = COMCNT
     CALL PtrERR(4, 56)
     IF UPDTFile THEN
        CALL FF3GetRec(FIRSTCOM): FF3COMMENT.CCNT = COMCNT
        CALL FF3PutRec
     END IF
1949 RETURN

2000 ' Check for broken CHILD chains
     IF A = 27 THEN GOTO P2900
     IF NOT ALLREC THEN GOTO P2500
     CALL PutMSG("VLDTM004")
     FOR CRID = 1 TO MAXID
         IF PTRCK1(CRID) <> 0 THEN
            X = PTRCK1(CRID) AND 3: Y = PTRCK1(CRID) \ 4
            IF (X = 1) OR (Y = 1) THEN
               CALL FF1GetRec(CRID)
               IF X = 1 THEN
                  SYSVAR.NVAR1 = FF1REC.FID: CALL PtrERR(1, 57)
               END IF
               IF Y = 1 THEN
                  SYSVAR.NVAR1 = FF1REC.MID: CALL PtrERR(1, 58)
               END IF
               IF UPDTFile THEN
                  IF X = 1 THEN FF1REC.FID = 0: FF1REC.FCH = 0
                  IF Y = 1 THEN FF1REC.MID = 0: FF1REC.MCH = 0
                  CALL FF1PutRec
               END IF
            END IF
         END IF
     NEXT CRID

2100 ' Check ADDRESS File FREE Chain
     IF A = 27 THEN GOTO P2900
     CALL PutMSG("VLDTM005")
     FREEBGN = FF2HDR.FREEBGN: FREECNT = FF2HDR.FREECNT
     FCNT = 0: FREC = FREEBGN: LASTFREE = 0: CRID = 0
2120 IF FREC = 0 THEN 2150
     FCNT = FCNT + 1: SYSVAR.NVAR1 = FCNT
     IF (FREC < 0) OR (FREC > HIREC2) THEN
        SYSVAR.NVAR2 = FREC: CALL PtrERR(9, 59)
        GOTO 2140
     END IF
     IF PTRCK2(FREC) > 0 THEN
        SYSVAR.NVAR2 = PTRCK2(FREC): CALL PtrERR(7, 60)
        GOTO 2140
     END IF
     PTRCK2(FREC) = 129
     LASTFREE = FREC
     CALL FF2GetRec(FREC)
     IF UPDTFile THEN
        IF (FF2PFX.RTYPE <> CHR$(255)) OR (FF2PFX.RNO <> FREC) THEN
           FF2PFX.RNO = FREC
           FF2PFX.RTYPE = CHR$(255): CALL FF2PutRec
        END IF
     END IF
     FREC = FF2PFX.NXT
     GOTO 2120

2140 ' Terminate FREE Chain with LastFree
     FCNT = FCNT - 1
     IF LASTFREE = 0 THEN
        FF2HDR.FREEBGN = 0
        ELSE
        IF UPDTFile THEN
           CALL FF2GetRec(LASTFREE): FF2PFX.NXT = 0: CALL FF2PutRec
        END IF
     END IF

2150 ' Check ADDRESS File FREE REC Count
     IF FREECNT = FCNT THEN 2160
     SYSVAR.NVAR2 = FREECNT: CALL PtrERR(9, 61)
     FF2HDR.FREECNT = FCNT
     IF UPDTFile THEN
        FFCHANGED = -1
        FF2HDR.REVLVL = FF2HDR.REVLVL OR 256
     END IF

2160 ' Check for Isolated ADDRESS Records
     IF A = 27 THEN GOTO P2900
     CALL PutMSG("VLDTM006")
     ISOCNT = 0
     FOR ID = 1 TO HIREC2
         IF PTRCK2(ID) = 0 THEN
            CALL FF2GetRec(ID)
            SYSVAR.NVAR1 = ID: SYSVAR.NVAR2 = ASC(FF2PFX.RTYPE)
            SYSVAR.NVAR3 = FF2PFX.SRNO: SYSVAR.NVAR4 = ASC(FF2PFX.SRTYPE)
            CALL PtrERR(8, 62)
            ISOCNT = ISOCNT + 1
         END IF
     NEXT ID
     IF ISOCNT = 0 THEN 2200

2170 ' Rebuild ADDRESS File FREE Chain
     IF A = 27 THEN GOTO P2900
     SYSVAR.NVAR1 = ISOCNT: CALL PtrERR(0, 63)
     IF UPDTFile THEN
        FREEBGN = 0: FREECNT = 0: LASTFREE = 0
        FOR ID = 1 TO HIREC2
            IF ((PTRCK2(ID) < 1) OR (PTRCK2(ID) = 129)) THEN
               IF FREEBGN = 0 THEN FREEBGN = ID: LASTFREE = ID
               CALL FF2GetRec(LASTFREE)
               FF2PFX.RTYPE = CHR$(255): FF2PFX.RNO = LASTFREE
               IF LASTFREE <> ID THEN FF2PFX.NXT = ID
               CALL FF2PutRec
               FREECNT = FREECNT + 1: LASTFREE = ID
            END IF
        NEXT ID
        IF (FF2HDR.FREEBGN <> FREEBGN) OR (FF2HDR.FREECNT <> FREECNT) THEN
           FF2HDR.FREEBGN = FREEBGN: FF2HDR.FREECNT = FREECNT
           FFCHANGED = -1: FF2HDR.REVLVL = FF2HDR.REVLVL OR 256
        END IF
        IF LASTFREE <> 0 THEN
           CALL FF2GetRec(LASTFREE)
           FF2PFX.RTYPE = CHR$(255)
           FF2PFX.NXT = 0: CALL FF2PutRec
        END IF
     END IF

2200 ' Check MISCELLANEOUS File FREE Chain
     IF A = 27 THEN GOTO P2900
     CALL PutMSG("VLDTM007")
     FREEBGN = FF3HDR.FREEBGN
     PTR& = FFRBA&(FF3HDR.FREEBGN, 1)
     IF PTR& > HIREC3& THEN FREEBGN = 0
     FREECNT& = FFRBA&(FF3HDR.FREECNT, 1)
     FCNT& = 0: FREC = FREEBGN: LASTFREE = 0
2210 IF FREC = 0 THEN 2250
     FCNT& = FCNT& + 1: SYSVAR.NVAR1 = FCNT&
     PTR& = FFRBA&(FREC, 1)
     IF (PTR& > HIREC3&) THEN
        SYSVAR.NVAR2 = PTR&: CALL PtrERR(9, 64)
        GOTO 2240
     END IF
     CALL PtrCHK3("GET", PTR&, PTRCK3)
     IF PTRCK3 <> 0 THEN
        SYSVAR.NVAR2 = PTRCK3: CALL PtrERR(7, 65)
        GOTO 2240
     END IF
     CALL PtrCHK3("PUT", PTR&, 129)
     LASTFREE = FREC: CALL FF3GetRec(FREC)
     IF UPDTFile THEN
        IF (FF3PFX.RTYPE <> CHR$(255)) OR (FF3PFX.RNO <> FREC) THEN
           FF3PFX.RNO = FREC
           FF3PFX.RTYPE = CHR$(255): CALL FF3PutRec
        END IF
     END IF
     FREC = FF3Pntrs.NXT
     GOTO 2210

2240 ' Terminate FREE Chain with LastFree
     FCNT& = FCNT& - 1
     IF FCNT& = 0 THEN
        FREEBGN = 0
        ELSE
        IF (LASTFREE <> 0) AND UPDTFile THEN
           CALL FF3GetRec(LASTFREE)
           FF3PFX.RTYPE = CHR$(255): FF3Pntrs.NXT = 0
           CALL FF3PutRec
        END IF
     END IF

2250 ' Check MISCELLANEOUS File FREE REC Count
     IF A = 27 THEN GOTO P2900
     IF FREECNT& <> FCNT& THEN
        SYSVAR.NVAR2 = FREECNT&: CALL PtrERR(9, 66)
     END IF
     IF UPDTFile THEN
        IF (FREEBGN <> FF3HDR.FREEBGN) OR (FREECNT& <> FCNT&) THEN
           FF3HDR.FREEBGN = FREEBGN: FF3HDR.FREECNT = FFRNUM(FCNT&, 1)
           FFCHANGED = -1: FF3HDR.REVLVL = FF3HDR.REVLVL OR 256
           FREECNT& = FCNT&
        END IF
     END IF

2270 ' Check for Isolated MISC Records
     IF A = 27 THEN GOTO P2900
     CALL PutMSG("VLDTM008")
     ISOCNT& = 0: ID& = 0
     WHILE ID& < HIREC3&: ID& = ID& + 1
         xID = FFRNUM(ID&, 1): CALL FF3GetRec(xID)
         CALL PtrCHK3("GET", ID&, PTRCK3)
         IF (PTRCK3 <> 0) THEN
            IF ((PTRCK3 \ 256) = 3) THEN
               IF ((PTRCK3 AND 15) <> 3) THEN
                  IF (FF3SPOUSE.SP1ID <> 0) AND (FF3SPOUSE.SP2ID <> 0) THEN
                     SYSVAR.NVAR1 = FF3SPOUSE.SP1ID
                     SYSVAR.NVAR2 = FF3SPOUSE.SP2ID
                     SYSVAR.NVAR3 = (PTRCK3 AND 15)
                     CALL PtrERR(2, 67)
                  END IF
               END IF
            END IF
            ELSE
            RT = ASC(FF3PFX.RTYPE)
            SELECT CASE RT
               CASE 3: SRT = ASC(FF3SPOUSE.SRTYPE): SRN = FF3SPOUSE.SP1ID
                       IF SRN = 0 THEN SRN = FF3SPOUSE.SP2ID
               CASE 4: SRT = ASC(FF3PLACE.SRTYPE): SRN = FF3PLACE.SRNO
               CASE 5: SRT = ASC(FF3COMMENT.SRTYPE): SRN = FF3COMMENT.SRNO
               CASE 6: SRT = ASC(FF3EVENT.SRTYPE): SRN = FF3EVENT.SRNO
               CASE 11: SRT = ASC(FF3EVENT.SRTYPE): SRN = FF3EVENT.SRNO
               CASE 12: SRT = ASC(FF3PLACE2.SRTYPE): SRN = FF3PLACE2.SRNO
               CASE ELSE: SRT = ASC(FF3Pntrs.SRTYPE): SRN = FF3Pntrs.SRNO
            END SELECT
            SYSVAR.NVAR1 = ID&: SYSVAR.NVAR2 = RT
            SYSVAR.NVAR3 = SRN: SYSVAR.NVAR4 = SRT
            CALL PtrERR(8, 68)
            ISOCNT& = ISOCNT& + 1
         END IF
     WEND
     IF A = 27 THEN GOTO P2900
     IF ISOCNT& = 0 THEN GOTO P2500

2280 ' Rebuild MISC File FREE Chain
     SYSVAR.NVAR1 = ISOCNT&: CALL PtrERR(0, 69)
     IF UPDTFile THEN
        FREEBGN = 0: FCNT& = 0: LASTFREE = 0: ID& = 0
        WHILE ID& < HIREC3&: ID& = ID& + 1
            xID = FFRNUM(ID&, 1)
            CALL PtrCHK3("GET", ID&, PTRCK3)
            IF ((PTRCK3 = 0) OR (PTRCK3 = 129)) THEN
               IF FREEBGN = 0 THEN FREEBGN = xID
               IF LASTFREE <> 0 THEN
                  CALL FF3GetRec(LASTFREE)
                  FF3PFX.RTYPE = CHR$(255): FF3PFX.RNO = LASTFREE
                  FF3Pntrs.NXT = xID: CALL FF3PutRec
               END IF
               FCNT& = FCNT& + 1: LASTFREE = xID
            END IF
        WEND
        IF LASTFREE <> 0 THEN
           CALL FF3GetRec(LASTFREE)
           FF3PFX.RTYPE = CHR$(255): FF3PFX.RNO = LASTFREE
           FF3Pntrs.NXT = 0: CALL FF3PutRec
        END IF
        FF3HDR.FREEBGN = FREEBGN: FF3HDR.FREECNT = FFRNUM(FCNT&, 1)
        FFCHANGED = -1: FF3HDR.REVLVL = FF3HDR.REVLVL OR 256
     END IF

P2500: ' END OF REPORT
     IF A = 27 THEN GOTO P2900
     IF (PRT.OPT = 1) OR (ERRCNT(0) = 0) THEN GOTO P2950
     LSET PLINE$ = " ": CALL WriteLINE
     PRT.SECTION = "T": CALL PrintTITLE
     NUMED$ = SPACE$(5)
     FOR ET = 1 TO 9
        IF ERRCNT(ET) > 0 THEN
           EMSG$ = "ET" + ZFILL$(ET, 2)
           CALL GetRVAR(EMSG$, X)
           IF FOUND THEN
              RSET NUMED$ = FNX$(ERRCNT(ET))
              LSET PLINE$ = NUMED$ + "  " + RVAR$(X)
              CALL PrintLINE: IF A = 27 THEN ET = 11
           END IF
        END IF
     NEXT ET
     CALL GetRVAR("ETOT", X)
     IF FOUND THEN
        RSET NUMED$ = FNX$(ERRCNT(0))
        LSET PLINE$ = NUMED$ + "  " + RVAR$(X)
        CALL PrintLINE
     END IF
     GOTO P2950

P2900: ' Early Termination
     IF PRT.OPT = 1 THEN
        IF CRID > 0 THEN
           SYSVAR.NVAR1 = CRID + 1
           CALL PutMSG("VLDTM009")
           ELSE
           CALL PutMSG("VLDTM010")
        END IF
        ELSE
        CALL RptCLOSE
     END IF
     GOTO P2990

P2950: ' Validation Procedure Complete
     IF ERID(0) <> 0 THEN
        SYSVAR.NVAR1 = 0
        FOR ET = 1 TO 6: SYSVAR.NVAR1 = SYSVAR.NVAR1 + ERRCNT(ET): NEXT ET
        SYSVAR.NVAR2 = FFRBA&(ERID(0), 1)
        CALL PutMSG("VLDTM019")
     END IF
     IF PRT.OPT = 1 THEN
        CALL PutMSG("VLDTM011")
        ELSE
        CALL RptCLOSE
     END IF
     IF UPDTFile THEN FFCHANGED = -1

P2990: ' Return
     CALL FamCLOSE: CALL PtrCHK3("DLET", 0, 0)
     CALL PutMSG("")

END SUB

REM $STATIC
SUB PtrCHK3 (OPT$, X&, B)
   OKAY = -1
   SELECT CASE OPT$
      CASE "OPEN": GOSUB ptCK3Open
      CASE "PUT": GOSUB ptCK3Put
      CASE "GET": GOSUB ptCK3Get
      CASE "DLET": GOSUB ptCK3Dlet
   END SELECT
   EXIT SUB

ptCK3Dlet: ' Delete Work File
   X$ = WRKFPFX$ + "FH$PTCK3"
   ON LOCAL ERROR GOTO ptCK3DletX
   KILL X$
   RETURN
ptCK3DletX: RESUME ptCK3Rtrn

ptCK3Open: ' Open Work File
   GOSUB ptCK3Dlet
   ON LOCAL ERROR GOTO ptCK3ERR
   SYSVAR.FILENAME = X$
   OPEN X$ FOR BINARY ACCESS READ WRITE AS #6
   Y& = (X& \ 100) + 1: Y$ = STRING$(200, 0)
   WHILE Y& > 0: PUT #6, (200 * (Y& - 1) + 1), Y$: Y& = Y& - 1: WEND
   RETURN

ptCK3Put: ' Put Number in Work File
   ON LOCAL ERROR GOTO ptCK3ERR
   PUT #6, (2 * X& - 1), B
   RETURN

ptCK3Get: ' Get Number from Work File
   ON LOCAL ERROR GOTO ptCK3ERR
   GET #6, (2 * X& - 1), B
   RETURN

ptCK3ERR: ' Error Routine
   CALL ErrMessage("XXXXM097", ERR, ERL)
   OKAY = 0
   RESUME ptCK3Rtrn

ptCK3Rtrn: RETURN

END SUB

REM $DYNAMIC
SUB PtrERR (ERRGrp, EMNUM)
SHARED CRID, LastRID, UPDTFile, XRID$, NOPAUSE, FMT019
SHARED ERRCNT(), ERID()
   IF A <> 27 THEN CALL ChkForPAUSE
   IF A = 27 THEN EXIT SUB
   ERRCNT(ERRGrp) = ERRCNT(ERRGrp) + 1
   IF ERRGrp <> 0 THEN ERRCNT(12) = ERRCNT(12) + 1
   IF CURFMT <> FMT019 THEN CURFMT = FMT019: FMT = FMTTAB(CURFMT)
   EFLD$ = "EC" + ZFILL$(ERRGrp, 2)
   CALL FmtFindFLD(EFLD$)
   IF FOUND THEN RSET UTXT$(FFLD.UTXT) = FNX$(ERRCNT(ERRGrp))
   CALL FmtFindFLD("ETOT")
   IF FOUND THEN
      RSET UTXT$(FFLD.UTXT) = FNX$(ERRCNT(12))
      CALL WinDisplayDATA(0)
   END IF
   IF CRID > 0 THEN
      ERID(CRID) = ERID(CRID) + 1
      RSET XRID$ = FNX$(CRID)
      ELSE
      LSET XRID$ = "-----"
   END IF
   IF CRID <> LastRID THEN
      LastRID = CRID
      IF CRID <> 0 THEN ERID(0) = ERID(0) + 1
      LSET PLINE$ = XRID$
      ELSE
      LSET PLINE$ = ""
   END IF
   EMNM$ = "EM" + ZFILL$(EMNUM, 2)
   CALL GetRVAR(EMNM$, X)
   IF FOUND THEN
      ERRMSG$ = RVAR$(X)
      ELSE
      ERRMSG$ = "Error Message " + EMNM$ + " Not Found"
   END IF
   MID$(PLINE$, 8, 255) = ERRMSG$
   CALL HTFEXPAND(PLINE$, 1)
   COLOR FG(5), BG(5): LOCATE 25, 1, 0: PRINT XRID$;
   COLOR FG(1), BG(1): LOCATE 25, 7, 0: PRINT MID$(PLINE$, 8, 70); TAB(79);
   IF PRT.OPT <> 1 THEN
      CALL PrintLINE
      ELSE
      LSET PLINE$ = ""
      IF NOT NOPAUSE THEN
         CALL ErrBEEP(0): CALL GetKEY
         NOPAUSE = ((LEN(A$) = 2) AND (A = 81))
      END IF
   END IF
END SUB

REM $STATIC
FUNCTION TextDATA$ (X$)
    SELECT CASE (ENV.SortOPTS AND 3)
       CASE 0: TextDATA$ = X$
       CASE 1: TextDATA$ = UCX$(X$)
       CASE ELSE: TD$ = SPACE$(4): CALL ComputeSNDX(X$, TD$)
                  TextDATA$ = TD$ + X$
    END SELECT
END FUNCTION

SUB ValiDATE (DTMD, DTYR)

9200 ' Validate Dates
     IF (DTMD <> 0) OR (DTYR <> 0) THEN
        DTMM = DTMD \ 100: DTDD = DTMD MOD 100
        CYR = VAL(RIGHT$(DATE$, 4))
        IF ((DTYR < 9999) AND (DTYR > CYR)) OR (DTMM > 12) OR (DTDD > 31) THEN
           OKAY = 0
           ELSE
           IF DTMM = 1 OR DTMM = 3 OR DTMM = 5 OR DTMM = 7 OR DTMM = 8 OR DTMM = 10 OR DTMM = 12 THEN
              OKAY = -1
              ELSE
              IF DTDD > 30 THEN
                 OKAY = 0
                 ELSE
                 IF DTMM <> 2 THEN
                    OKAY = -1
                    ELSE
                    LPYR = ((DTYR MOD 4) = 0) AND (((DTYR MOD 400) = 0) OR ((DTYR MOD 100) > 0))
                    OKAY = (DTDD < (29 - LPYR))
                 END IF
              END IF
           END IF
        END IF
     END IF
     IF NOT OKAY THEN
        LSET SYSVAR.MVAR = XSCRNDate$(DTMD, DTYR)
     END IF
END SUB

