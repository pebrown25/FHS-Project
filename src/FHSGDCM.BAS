DECLARE SUB GetSPHONE (X$)
DECLARE SUB ConvertADRS ()
DECLARE SUB GEDCOMStats ()
DECLARE SUB ConvertNAME (XVALUE$, XSURNM$, XGIVEN$)
DECLARE SUB ConvertCASE (OPT$, TEXT$, S%)
DECLARE SUB ConvertDATE (XDATE$, XDY%, XDMD%, XST$)
DECLARE SUB CountRECS ()
DECLARE SUB GEDCOMExport ()
DECLARE SUB GEDCOMImport ()
DECLARE SUB GEDCOMLine (GDTAG%, XREF$, VALUE$)
DECLARE SUB GEDCOMList ()
DECLARE SUB GEDCOMOpen (OPT$)
DECLARE SUB GetLINE ()
DECLARE SUB GetTAGS (OPT$)
DECLARE SUB PointerFILES (OPT$)
DECLARE SUB ConvertTIME (XTIME$, XTIME%)
DECLARE FUNCTION DATEDif$ (D1MD%, D1Y%, D2MD%, D2Y%)
DECLARE FUNCTION FDSTAT$ (X$)
DECLARE FUNCTION FMTNAME$ (NFMT%)
DECLARE FUNCTION FFRBA& (RNUM%, RLTH%)
DECLARE FUNCTION FFRNUM% (RBA&, RLTH%)
DECLARE FUNCTION RELATION$ (A$, B$, C%, D%, RELCD$())
DECLARE FUNCTION SFILL$ (X%, L%)
DECLARE FUNCTION ZFILL$ (X%, L%)
DECLARE FUNCTION UCX$ (X$)
DECLARE FUNCTION XRPTDate$ (MD%, Y%, O%)
DECLARE FUNCTION XSCRNDate$ (D1MD%, D1Y%)

TYPE PtrWrkREC1    ' 20 Bytes
   WFID     AS INTEGER
   WFFID    AS INTEGER
   WFMID    AS INTEGER
   WFCH     AS INTEGER
   WFFSIB   AS INTEGER
   WFMSIB   AS INTEGER
   WFMREC   AS INTEGER
   WFPMREC  AS INTEGER
   WFP      AS INTEGER
   WFSEX    AS STRING * 1
   WFS1     AS STRING * 1
END TYPE
TYPE PtrWrkREC2    ' 14 Bytes
   WFMRN    AS INTEGER
   WFHID    AS INTEGER
   WFWID    AS INTEGER
   WFHOF    AS INTEGER
   WFWOF    AS INTEGER
   WF1CH    AS INTEGER
   WFSP     AS INTEGER
END TYPE

1  REM $INCLUDE: 'FHSCOMON.BAS'
   PN$ = " FAMILY HISTORY SYSTEM - FHSGEDCM - GEDCOM Export/Import Program"
   CY$ = " (C) Copyright 1989,1990,1991,1992,1993,1994,1996,1997 by Phillip E. Brown"
   IF CONFIG.BP = 0 THEN RUN "FHSINIT"
   DEF FNCHR (A$) = (((A$ > "@") AND (A$ < "[")) OR ((A$ > "`") AND (A$ < "{")))
   DEF FNNUM (A$) = ((A$ > "/") AND (A$ < ":"))
   DEF FNX$ (A) = LTRIM$(STR$(A))
   REDIM RECCNT(11), TAG$(140)
   GEDFile$ = FT$(11)
   CALL WinCLR
   CALL WinINIT("LR1")
   CALL WinPREP(1): PGMRTRN$ = "FHSMENU"
   FUN = 0: READ$ = "READQ"
   CALL RptLOAD("GDIM"): CALL RptLOAD("GDEX"): CALL RptLOAD("GDLS")
   RRPT$ = "GDEX"
   CALL FmtFIND("GDCMS001"): FMT001 = CURFMT: CALL FmtDREST
   CALL FmtFIND("GDCMS005")
   X = FMT.BGNFLD
   WHILE X > 0
      FFLD = FLDTAB(X)
      IF FFLD.UTXT <> 0 THEN
         X$ = UTXT$(FFLD.UTXT)
         SELECT CASE FFLD.NAME
            CASE "NAME": IF RTRIM$(X$) = "" THEN X$ = CFGAL$(1)
            CASE "ADR1": IF RTRIM$(X$) = "" THEN X$ = CFGAL$(2)
            CASE "ADR2": IF RTRIM$(X$) = "" THEN X$ = CFGAL$(3)
            CASE "ADR3": IF RTRIM$(X$) = "" THEN X$ = CFGAL$(4)
            CASE "SOUR": IF RTRIM$(X$) = "" THEN X$ = "FHS 97"
         END SELECT
      END IF
      LSET UTXT$(FFLD.UTXT) = X$
      X = FLDCHN(X).FWD
   WEND
   CALL FmtFIND("GDCMS010"): FMT010 = CURFMT
   CALL FmtFIND("GDCMS011"): FMT011 = CURFMT
   FF = FMT.BGNFLD
   WHILE FF > 0
      FFLD = FLDTAB(FF)
      IF FFLD.UTXT > 0 THEN RSET UTXT$(FFLD.UTXT) = SNGLKEY$(1)
      FF = FLDCHN(FF).FWD
   WEND

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
         SELECT CASE FFLD.NAME
            CASE "FSUP": LSET X$ = FDFSetup.NAME
            CASE "FSDS": LSET X$ = FDFSetup.DESC
            CASE "PSUP": LSET X$ = PDFSetup.NAME
            CASE "PSDS": LSET X$ = PDFSetup.DESC
            CASE "PRTR": LSET X$ = PDFSetup.PRINTER
            CASE "FMWD": LSET X$ = STR$(PDFFORMS.FWIDTH)
            CASE "FMLT": LSET X$ = STR$(PDFFORMS.FLENGTH)
            CASE "GDDS": LSET X$ = GEDFile$
            CASE "LINS": LSET X$ = STR$(GEDLine&)
            CASE "INDI": LSET X$ = STR$(INDI)
            CASE "FAMI": LSET X$ = STR$(FAMI)
         END SELECT
         LSET UTXT$(FFLD.UTXT) = X$
      END IF
      FF = FLDCHN(FF).FWD
   WEND

100 ' Check Files
    CALL FamOPEN(READ$): READ$ = "READQ": FAMOKAY = ENV.FFOPEN
    CURFMT = FMT001: FMT = FMTTAB(CURFMT)
    CALL FmtFindFLD("NMCT")
    IF ENV.FFOPEN THEN
       LSET UTXT$(FFLD.UTXT) = STR$(FF1HDR.MAXID)
       ELSE
       LSET UTXT$(FFLD.UTXT) = " "
    END IF
    CALL WinDisplayDATA(FMT001)
    CALL FamCLOSE

150 CALL WinPREP(2)
    IF RRPT$ <> "GDEX" THEN CALL WinCLR: RRPT$ = "GDEX"
    CALL RptOptDISPLAY("GDEX")
    CALL WinSWITCH(1)

200 ' Get Processing Option
    CALL PutMSG("")
    IF FUN$ <> "" THEN CALL OptHILITE(FUN$, FUN$, 1)
    CALL GetKEY
    IF LEN(A$) < 2 THEN CALL ErrBEEP(0): FUN = 0: GOTO 200:  ELSE FUN = A
    SELECT CASE FUN
       CASE 59: FUN$ = "F1"
       CASE 60: FUN$ = "F2"
       CASE 62: FUN$ = "F4"
       CASE 63: FUN$ = "F5"
       CASE 64: FUN$ = "F6"
       CASE 65: FUN$ = "F7"
       CASE 66: CALL DOSShell: GOTO 40
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
       CASE "F5": GOTO M600
       CASE "F6": CALL GEDCOMList: IF PRT.OPT = 2 THEN 200 ELSE GOTO 40
       CASE "F7": CALL PutMSG("GDCMM056")
                  SELECT CASE A$
                     CASE "1": CALL ConvertADRS
                     CASE "2": CALL FF1OPEN("WRITE"): IF NOT OKAY THEN 100
                               FF1HDR.REVLVL = 1: PUT #1, 1, FF1HDR
                               CALL FamCLOSE
                               SYSVAR.MVAR = FDFSetup.NAME
                               CALL PutMSG("GDCMM057")
                  END SELECT
                  GOTO 100
       CASE "F9": GOTO 250
    END SELECT
    GOTO 200

250 ' Return to FAMMENU.BAS
    CALL FamCLOSE
    ERASE RFLTH, RFLBL$
    CALL PgmPREP(PGMRTRN$)
    IF (NOT OKAY) OR (A = 27) THEN LSET PGMRTRN$ = "FHSMENU": GOTO 40
    CALL RptFREE("GDIM"): CALL RptFREE("GDEX"): CALL RptFREE("GDLS")
    CHAIN PGMRTRN$

300 ' Change FILE Parameters
    IF A = 94 THEN LSET PGMRTRN$ = "FHSFILE": GOTO 250
    CALL WinSWITCH(2)
    CALL FDFSelect: READ$ = "READ"
    CALL WinSWITCH(1)
    GOTO 50

400 ' Change PRINTER Parameters
    IF A = 95 THEN LSET PGMRTRN$ = "FHSPRTC": GOTO 250
    CALL WinSWITCH(2)
    CALL PDFSelect
    CALL WinSWITCH(1)
    GOTO 50

500 ' Process Request to Update Options
    CALL PutMSG("GDCMM001")
    IF A = 27 THEN 200
    SELECT CASE A$
       CASE "1": RRPT$ = "GDIM"
       CASE "2": RRPT$ = "GDEX"
       CASE ELSE: RRPT$ = "GDLS"
    END SELECT
    CALL WinSWITCH(2)
    IF RRPT$ <> "GDEX" THEN CALL WinCLR
    OptTABLE.CL = 0
    CALL RptOptDISPLAY(RRPT$)
    CALL RptOptUPDATE(RRPT$)
    GOTO 150

M600: ' Process GEDCOM Request
    CALL PutMSG("GDCMM002")
    IF A = 27 THEN 200
    IF A$ = "2" THEN
       CALL GEDCOMExport
       ELSE
       CALL GEDCOMImport
    END IF
    CALL WinPREP(2): CALL WinCLR
    GOTO 40

REM $DYNAMIC
SUB ConvertADRS
    CALL FF1OPEN("READ"): IF NOT OKAY THEN GOTO CADRSExit
    CALL FF2Open("READ"): IF NOT OKAY THEN GOTO CADRSExit
    CLOSE #2: OAF$ = RTRIM$(SYSVAR.FILENAME)
    OAFMT = FF2HDR.REVLVL AND 255
    IF OAFMT = 1 THEN
       CALL PutMSG("GDCMM050")
       GOTO CADRSExit
    END IF
    CALL PutMSG("GDCMM051")
    IF A$ <> SNGLKEY$(1) THEN GOTO CADRSExit
    X = INSTR(OAF$, ".")
    IF X = 0 THEN OAF$ = OAF$ + ".OAF" ELSE OAF$ = LEFT$(OAF$, X) + "OAF"
    ON LOCAL ERROR GOTO CADRSSkp1
    KILL OAF$
    GOTO CADRSSkp2
CADRSSkp1: RESUME CADRSSkp2
CADRSSkp2: ON LOCAL ERROR GOTO CADRSSkp3
    NAME SYSVAR.FILENAME AS OAF$
    GOTO CADRSSkp4
CADRSSkp3: CALL ERRMessage("GDCMM052", ERR, ERL)
    OKAY = 0
    RESUME CADRSExit
CADRSSkp4:
    ON LOCAL ERROR GOTO CADRSExit
    OPEN OAF$ FOR BINARY ACCESS READ AS #3
    FF2HDR.REVLVL = (FF2HDR.REVLVL \ 256) * 256 + 1
    OPEN SYSVAR.FILENAME FOR BINARY ACCESS READ WRITE AS #2
    PUT #2, 1, FF2HDR
    X$ = STRING$(133, 0)
    PUT #2, , X$
    X$ = SPACE$(108): GET #3, 1, X$
    CALL PutMSG("GDCMM053")
    FOR X = 1 TO FF2HDR.HIREC
        GET #3, , FF2PFX: PUT #2, , FF2PFX
        GET #3, , FF2SREC
        FF2LREC.PHONE = ZFILL$(FF2SREC.FON1, 3) + "-" + ZFILL$(FF2SREC.FON2, 3) + "-" + ZFILL$(FF2SREC.FON3, 4)
        FF2LREC.LINE1 = FF2SREC.LINE1
        FF2LREC.LINE2 = FF2SREC.LINE2
        LSET FF2LREC.CITY = FF2SREC.CITY
        LSET FF2LREC.STATE = FF2SREC.STATE
        LSET FF2LREC.ZIP = FF2SREC.ZIP
        LSET FF2LREC.COUNTRY = ""
        PUT #2, , FF2LREC
    NEXT
    CALL PutMSG("GDCMM054")
    OKAY = -1
CADRSExit:
    CLOSE
END SUB

SUB ConvertCASE (OPT$, TEXT$, S)
    ' OPT$: NAME, TEXT   C: 0=Begin Text, 1=Continued Text
    TextFLD = (OPT$ = "TEXT"): W = 0: IF NOT TextFLD THEN S = 0
    X = LEN(RTRIM$(TEXT$))
    IF X > 0 THEN
       FOR Z = 1 TO X
           A = ASC(MID$(TEXT$, Z, 1))
           ' End Sentence at .,!,?
           IF (A = 46) OR (A = 21) OR (A = 63) THEN
              S = 0: W = 0
              ELSE
              'Count characters since Beginning of Sentence
              S = S - (A <> 32)
              ' Non-Alpha Character: End of Word
              IF (A < 65) OR ((A > 90) AND (A < 97)) OR (A > 122) THEN
                 W = 0
                 ELSE
                 ' Lower-Case Character: Stop Conversion
                 IF (A > 96) THEN
                    Z = X
                    ELSE
                    ' Upper-Case Character
                    W = W + 1
                    IF (W > 1) OR (TextFLD AND (S > 1)) THEN
                       MID$(TEXT$, Z, 1) = CHR$(A + 32)
                    END IF
                 END IF
              END IF
           END IF
       NEXT Z
       A = 0
    END IF
END SUB

REM $STATIC
SUB ConvertDATE (XDATE$, XDY, XDMD, XST$)
SHARED GDCMRULE
9010 ' yyyymmdd-->d1y,d1md
     IF RTRIM$(XDATE$) = "" THEN
        XDMD = 0
        ELSE
        X = 1
        IF (GDCMRULE < 3) OR (GDCMRULE = 5) THEN 9030
        IF MID$(XDATE$, 1, 1) = "D" THEN X = X + 6
        DFMT$ = MID$(XDATE$, X, 8)
        XST$ = MID$(XDATE$, X + 8, 1)
        XDY = VAL(MID$(DFMT$, 1, 4))
        XDMD = 100 * VAL(MID$(DFMT$, 5, 2)) + VAL(MID$(DFMT$, 7, 2))
     END IF
     EXIT SUB

9030 ' dd mmm yyyy-->d1y,d1md
     DL = LEN(XDATE$)
     V1 = 0: V2 = 0: V3 = 0: XST$ = " "
     Y = X: XX$ = MID$(XDATE$, X, 2)
     IF (XX$ = "AB") OR (XX$ = "AF") OR (XX$ = "BE") THEN
        LSET ST$ = "A": GOSUB 9050: GOSUB 9052: Y = X
     END IF
     GOSUB 9051: L = X - Y
     IF L > 0 THEN
        V1 = VAL(MID$(XDATE$, Y, L)): X = X + 1
        IF ST$ = " " THEN LSET ST$ = MID$(XDATE$, X - 1, 1)
     END IF
     IF X > DL THEN 9045
     GOSUB 9052: Y = X
     GOSUB 9050: L = X - Y: IF L > 0 THEN GOSUB 9070: V2 = MM
     IF X > DL THEN 9045
     GOSUB 9052: Y = X
     GOSUB 9051: L = X - Y
     IF L > 0 THEN
        V3 = VAL(MID$(XDATE$, Y, L))
        LSET ST$ = MID$(XDATE$, X, 1)
     END IF
9045 IF (V2 = 0 AND V3 = 0) THEN V3 = V1: V1 = 0
     IF V3 > 0 THEN XDY = V3
     XDMD = 100 * V2 + V1
     EXIT SUB

9050 WHILE ((X < (DL + 1)) AND FNCHR(MID$(XDATE$, X, 1)))
        X = X + 1
     WEND: RETURN
9051 WHILE ((X < (DL + 1)) AND FNNUM(MID$(XDATE$, X, 1)))
        X = X + 1
     WEND: RETURN
9052 WHILE (X < (DL + 1)) AND (MID$(XDATE$, X, 1) = " ")
        X = X + 1
     WEND: RETURN


9070 ' mmm-->mm
     MON$ = UCX$(MID$(XDATE$, Y, 3))
     MM = 13
     WHILE (MM > 1) AND (MON$ <> MID$(RMONTB$(MM), 3, 3)): MM = MM - 1: WEND
     MM = MM - 1
     RETURN

END SUB

SUB ConvertNAME (XVALUE$, XSURNM$, XGIVEN$)

    'XVALUE$-->FF1REC.Surnm & FF1REC.Given
     L = LEN(RTRIM$(XVALUE$)): LSET XSURNM$ = "": LSET XGIVEN$ = ""
     IF L > 0 THEN
        X = INSTR(1, XVALUE$, "/")
        IF X > 0 THEN
           IF X > 1 THEN LSET XGIVEN$ = MID$(XVALUE$, 1, X - 1)
           Y = INSTR(X + 1, XVALUE$, "/")
           IF Y = 0 THEN Y = 255
           LSET XSURNM$ = MID$(XVALUE$, X + 1, Y - X - 1)
           ELSE
           X = INSTR(1, XVALUE$, ",")
           IF X > 1 THEN LSET XSURNM$ = MID$(XVALUE$, 1, X - 1)
           X = X + 1
           IF X < L THEN LSET XGIVEN$ = LTRIM$(MID$(XVALUE$, X, 255))
        END IF
        X = INSTR(1, XSURNM$, "_")
        WHILE X > 0
           MID$(XSURNM$, X, 1) = " ": X = INSTR(X + 1, XSURNM$, "_")
        WEND
     END IF

END SUB

SUB ConvertTIME (XTIME$, XTIME)
    X = INSTR(XTIME$, ":"): Y = LEN(RTRIM$(XTIME$))
    IF X = 0 THEN
       XTIME = VAL(LEFT$(XTIME$, Y))
       ELSE
       XTIME = 100 * (VAL(LEFT$(XTIME$, X - 1)))
       IF Y > X THEN XTIME = XTIME + VAL(MID$(XTIME$, X + 1, Y - X))
    END IF
END SUB

REM $DYNAMIC
SUB CountRECS
SHARED NOPAUSE, PrintERR, FIREF, IIREF, TLINE$, RT, RECCNT()
990 ' Process Record Counts
    IF A$ = CHR$(27) THEN EXIT SUB
    IF RT > 11 THEN RT = 11
    RECCNT(RT) = RECCNT(RT) + 1
    IF (NOT PrintERR) OR (PRT.OPT <> 1) THEN
       IF RT < 11 THEN
          COLOR FG(4), BG(4): LOCATE 13 + RT, 43, 0
          PRINT USING "#####"; RECCNT(RT);
          ELSE
          LOCATE 25, 1, 0: COLOR FG(1), BG(1)
          PRINT LEFT$(TLINE$, 73); TAB(74);
          COLOR FG(5), BG(5)
          IF NOT NOPAUSE THEN
             LOCATE 25, 74, 0: PRINT " ... ";
             CALL GetKEY
             IF (LEN(A$) = 2) AND (A = 81) THEN NOPAUSE = -1
             IF A$ <> CHR$(27) THEN A$ = ""
          END IF
          LOCATE 25, 74, 0: PRINT USING "#####"; RECCNT(11);
       END IF
    END IF
    IF (RT > 10) THEN
       IF PrintERR AND (A$ <> CHR$(27)) THEN
          MID$(PLINE$, 1, 5) = SFILL$(IIREF, 5)
          MID$(PLINE$, 7, 5) = SFILL$(FIREF, 5)
          MID$(PLINE$, 15, 255) = TLINE$
          CALL PrintLINE
          IF (LEN(A$) = 2) AND (A = 81) THEN NOPAUSE = -1
          IF A$ <> CHR$(27) THEN A$ = ""
       END IF
    END IF
END SUB

SUB GEDCOMExport
SHARED HIID, HIMRN, CNVTCASE, S, GEDLine&, FMT001, GEDFile$
SHARED LDiskNUM, DiskNUM, DISKEOF
SHARED INDI, FAMI, INFOSW, GDCMTYPE, GDCMRULE, RT, GL, RECCNT(), TAG$()

600 ' Export Information to GEDCOM formatted dataset
    ON LOCAL ERROR GOTO 601
    REDIM XMONTB$(14): FOR X = 1 TO 14: XMONTB$(X) = RMONTB$(X): NEXT X
    CALL RptLOAD("GDEX"): IF NOT OKAY THEN GOTO 750
    INFOSW = 0
    FOR X = 0 TO 9
        X$ = "GD" + ZFILL$(X, 2)
        CALL RptOptGET(X$, V): IF V = 1 THEN INFOSW = INFOSW OR (2 ^ X)
    NEXT X
    NFMT$ = SPACE$(55): DFMT$ = SPACE$(12)
    CLINE$ = SPACE$(150): XLINE$ = SPACE$(100)
    BLOC$ = SPACE$(45): DLOC$ = BLOC$: XTIM$ = SPACE$(5): XREF$ = SPACE$(10)
    AL1$ = SPACE$(30): AL2$ = AL1$: CITY$ = SPACE$(20): STATE$ = SPACE$(12)
    ZIP$ = SPACE$(12): PHONE$ = SPACE$(14): CNTRY$ = SPACE$(LEN(FF2LREC.COUNTRY))
    HIREC& = FFRBA&(FF3HDR.HIREC, 1)
    IF HIREC& > 32766 THEN
       MC1 = 32766: X& = HIREC& - 32766: IF X& > 32766 THEN X& = 32766
       MC2 = X&
       ELSE
       MC1 = HIREC&: MC2 = 0
    END IF
    REDIM LoMISCSEQ(MC1), HiMISCSEQ(MC2)
    GOTO 605

601 'Error Setting up for GEDCOM Export
    CALL ERRMessage("GDCMM016", ERR, ERL)
    REDIM LoMISCSEQ(1), HiMISCSEQ(1), NDXID(1), SELSET(1), RECCNT(1)
    RESUME 750

605 ' Determine Records to Process
    CALL PutMSG("GDCMM003")
    IF ASC(A$) = 27 THEN 750 ELSE ALLRECS = (A$ = "1")
    REDIM SELSET(FF1HDR.MAXID)
    IF ALLRECS THEN
       SLCTMAX = FF1HDR.MAXID
       FOR X = 1 TO SLCTMAX: SELSET(X) = -1: NEXT X
       ELSE
       CALL SWRKBldList(SELSET(), SLCTMAX, 0)
       IF NOT OKAY THEN GOTO 605
       FOR X = 1 TO SLCTMAX
           IF SELSET(X) <> 0 THEN SELSET(X) = -1
       NEXT X
    END IF
    SELSET(0) = 0

660 ' Determine Sequence of Exported Records
    CALL PutMSG("GDCMM004")
    IF ASC(A$) = 27 THEN 690
    IF A$ < "1" OR A$ > "3" THEN SOUND BP, DUR: GOTO 660
    RELSEQ = (A$ = "2"): INDEXED = (A$ = "3")
    REDIM NDXID(FF1HDR.MAXID)
    IF RELSEQ THEN
       CALL RWRKOpen("READ", 0, 0): IF NOT OKAY THEN GOTO 660
       NDXRECS = 0
       CALL PutMSG("XXXXM064")
       CALL RWRKGet("BASE", RC, PTR&)
       WHILE (0 < RC) AND (RC < 9)
          IF WFRELREC.ID > 0 THEN
             IF SELSET(WFRELREC.ID) < 0 THEN
                NDXRECS = NDXRECS + 1: NDXID(NDXRECS) = WFRELREC.ID
                SELSET(WFRELREC.ID) = 1
             END IF
          END IF
          CALL RWRKGet("WRKREC", RC, PTR&)
       WEND
       FOR X = 1 TO SLCTMAX
           Y = SELSET(X)
           IF Y < 0 THEN
              NDXRECS = NDXRECS + 1: NDXID(NDXRECS) = X
              ELSE
              IF Y > 0 THEN SELSET(X) = -1
           END IF
       NEXT X
       CALL PutMSG("")
       ELSE
       IF INDEXED THEN
          CALL INDEXFile("READ")
          IF NOT OKAY THEN 660
          NDXRECS = NDXHdr.NDXRECS
          ELSE
          FOR I = 1 TO FF1HDR.MAXID: NDXID(I) = I: NEXT I
          NDXRECS = FF1HDR.MAXID
       END IF
    END IF
    HIID = 0
    FOR X = 1 TO NDXRECS
        Y = NDXID(X)
        IF SELSET(Y) < 0 THEN HIID = HIID + 1: SELSET(Y) = HIID
    NEXT X
    REDIM PRESERVE NDXID(NDXRECS)
    FOR X = 1 TO SLCTMAX
        IF SELSET(X) < 0 THEN SELSET(X) = 0
    NEXT X
    CALL GEDCOMStats

665 ' Get GEDCOM Tags - TAG$()
    CALL GetTAGS("EXPORT"): IF NOT OKAY THEN 750
    IF GDCMTYPE = 5 THEN
       CALL PutMSG("GDCMM006"): IF A = 27 THEN 750
       GDCMRULE = VAL(A$)
       IF GDCMRULE = 4 THEN GDCMRULE = 5
       IF (GDCMRULE < 2) OR (GDCMRULE > 5) THEN GDCMRULE = 1
       ELSE
       GDCMRULE = GDCMTYPE
    END IF
    SELECT CASE GDCMRULE
       CASE 1: LDSv55 = -1: GDRULE$ = "LDSV55"
       CASE 2: PAFV21 = -1: GDRULE$ = "PAFV21"
       CASE 3: PAFV20 = -1: GDRULE$ = "PAFV20"
       CASE 4: FHSv1 = -1: GDRULE$ = "FHSV1"
       CASE 5: FHSv2 = -1: GDRULE$ = "FHSV2"
    END SELECT
    CUST = (GDCMTYPE = 5)
    IF (GDCMTYPE < 4) OR ((INFOSW AND 1) = 0) THEN
       IF (GDCMRULE = 2) THEN
          X$ = " 0    1JAN 2FEB 3MAR 4APR 5MAY 6JUN 7JUL 8AUG 9SEP10OCT11NOV12DEC13???"
          ELSE
          X$ = " 0    1Jan 2Feb 3Mar 4Apr 5May 6Jun 7Jul 8Aug 9Sep10Oct11Nov12Dec13???"
       END IF
       FOR X = 1 TO 14
           RMONTB$(X) = MID$(X$, 5 * (X - 1) + 1, 5)
       NEXT X
    END IF

    CALL PutMSG("GDCMM019")
    CNVTCASE = (A$ = SNGLKEY$(1))

    ' Create Pointer Workfile
    DIM PtrREC1 AS PtrWrkREC1, PtrREC2 AS PtrWrkREC2
    CALL PointerFILES("OPEN"): IF NOT OKAY THEN 690

670 ' Enter Header Information
    CALL WinPREP(1): CALL WinSWITCH(2): CALL WinCLR
    CALL FmtFIND("GDCMS005"): CALL WinFORMAT(0)
    CALL PutMSG("GDCMM007")
    CALL WinUPDATE: IF A = 27 THEN 690
    CALL PutMSG("")
    REDIM HLIT$(10)
    X = FMT.BGNFLD
    WHILE X > 0
       FFLD = FLDTAB(X)
       IF FFLD.UTXT <> 0 THEN
          X$ = UTXT$(FFLD.UTXT)
          SELECT CASE FFLD.NAME
             CASE "NAME": HLIT$(1) = X$
             CASE "ADR1": HLIT$(2) = X$
             CASE "ADR2": HLIT$(3) = X$
             CASE "ADR3": HLIT$(4) = X$
             CASE "PHON": HLIT$(5) = X$
             CASE "SOUR": HLIT$(6) = X$
             CASE "DEST": HLIT$(7) = X$
             CASE "COM1": HLIT$(8) = X$
             CASE "COM2": HLIT$(9) = X$
             CASE "COM3": HLIT$(10) = X$
          END SELECT
       END IF
       X = FLDCHN(X).FWD
    WEND
    GOTO 700

690 ' Early Return from EXPORT operation
    CALL FamCLOSE
    GOTO 735

700 ' Export Data for Selected Name Records
    CALL FamOPEN("READ"): IF NOT OKAY THEN GOTO 690
    CALL GEDCOMOpen("WRITE"): IF NOT OKAY THEN 750
    GOSUB 19100                                     ' Build POINTER Workfile
    IF (NOT OKAY) OR (A = 27) THEN 690
    CALL GEDCOMStats
    GL = 0: MISCSEQ = 0: GOSUB 800                  ' Create HEADER Record
    CALL WinCLR: CALL FmtFIND("GDCMS010"): CALL WinFORMAT(0)
    CALL PutMSG("GDCMM008")
    ON LOCAL ERROR GOTO 705
    'ON LOCAL ERROR GOTO 0
    GOTO 720

705 ' Error During Export
    XERR = ERR: XERL = ERL: CALL GEDCOMStats
    CALL ERRMessage("GDCMM017", XERR, XERL)
    OKAY = 0: SID = HIID: I = HIMRN: A = 27: A$ = CHR$(27)
    SELECT CASE ERRReturn
       CASE 1: RESUME 721
       CASE 2: RESUME 979
       CASE 3: COMPTR = 0: RESUME 981
       CASE 4: RESUME 750
    END SELECT
    RESUME 730

720 ' Process Records in Pointer Workfiles
    REDIM RECCNT(11): SELSET(0) = 0
    IF FHSv1 OR PAFV20 THEN IIPFX$ = "" ELSE IIPFX$ = "I"
    IF PAFV20 THEN FIPFX$ = "" ELSE IF FHSv1 THEN FIPFX$ = "M" ELSE FIPFX$ = "F"
    FOR SID = 1 TO HIID
        GET #6, SID, PtrREC1: ID = PtrREC1.WFID
        CALL FF1GetRec(ID): GOSUB 830
        IF (A$ = CHR$(27)) OR DISKEOF THEN
           SID = HIID
           ELSE
           A$ = INKEY$
           IF A$ <> "" THEN
              K$ = A$: A = ASC(K$): IF A = 27 THEN SID = HIID
           END IF
        END IF
        IF (K$ <> "") OR (GEDLine& - HGEDLine& > 200) THEN
           CALL GEDCOMStats: K$ = "": HGEDLine& = GEDLine&
        END IF
721 NEXT SID
    IF (A$ <> CHR$(27)) AND (NOT DISKEOF) THEN
       GOSUB 900    ' Create Family records
       GL = 0
       IF NOT DISKEOF THEN
          CALL GEDCOMLine(140, "", "")    ' Create Trailer Record
       END IF
    END IF

730 ' End of EXPORT
    ERRReturn = 4
    CALL GEDCOMStats
    CALL PutMSG("GDCMM009")
735 CALL PutMSG("")
    CALL WinPREP(2): CALL WinCLR
750 CALL FamCLOSE
    CALL PointerFILES("DELETE")
    FOR X = 1 TO 14: RMONTB$(X) = XMONTB$(X): NEXT X
    EXIT SUB

800 ' Create GEDCOM File Header Record
    GL = 0
    CALL GEDCOMLine(1, "", "")
    GL = GL + 1
    IF NOT (PAFV20 OR FHSv1) THEN
       HL = 6: CTAG = 10: GOSUB 810
       HL = 7: CTAG = 13: GOSUB 810
       LSET DFMT$ = DATE$: GOSUB 9090
       CALL GEDCOMLine(14, "", DFMT$)
       CALL GEDCOMLine(15, "", GEDFile$)
    END IF
    IF TAG$(16) <> TAG$(0) THEN CALL GEDCOMLine(16, "", LTRIM$(STR$(HIID)))
    IF TAG$(17) <> TAG$(0) THEN CALL GEDCOMLine(17, "", LTRIM$(STR$(HIMRN)))
    IF TAG$(20) <> TAG$(0) THEN CALL GEDCOMLine(20, "", GDRULE$)
    IF PAFV21 OR LDSv55 THEN
       GL = GL - 1: CALL GEDCOMLine(4, "@S1@", "")
       GL = GL + 1
    END IF
    IF PAFV20 THEN CALL GEDCOMLine(4, "", ""): GL = GL + 1
    HL = 1: CTAG = 5: GOSUB 810
    CTAG = 6: GOSUB 810
    IF NOT (FHSv1 OR PAFV20) THEN GL = GL + 1
    CTAG = 7: GOSUB 810: GOSUB 810
    IF NOT (FHSv1 OR PAFV20) THEN GL = GL - 1
    CTAG = 8: GOSUB 810
    HL = 8
    CTAG = 18: GOSUB 810
    IF NOT (FHSv1 OR PAFV20) THEN GL = GL + 1
    CTAG = 19: GOSUB 810: GOSUB 810
    IF (FHSv1 OR PAFV20) THEN
       IF NOT FHSv1 THEN GL = GL - 1: CALL GEDCOMLine(9, "", ""): GL = GL + 1
       HL = 6: CTAG = 10: GOSUB 810: CTAG = 13: GOSUB 810
    END IF
    GL = GL - 1
    RETURN

810 ' Create Heading Lines
    IF (RTRIM$(HLIT$(HL)) <> "") AND (HLIT$(HL) <> "(000) 000-0000") THEN
       CALL GEDCOMLine(CTAG, "", HLIT$(HL))
    END IF
    HL = HL + 1
    RETURN

830 ' Create NAME Record Lines
    GL = 0: RT = 1: CALL CountRECS
    LSET XREF$ = "@" + IIPFX$ + FNX$(SID) + "@"
    CALL GEDCOMLine(30, XREF$, "")
    GL = GL + 1
    CALL GEDCOMLine(31, "", FNX$(PtrREC1.WFID))
    GOSUB 9110: CALL GEDCOMLine(32, "", NFMT$)
    S1 = ASC(FF1REC.STATUS): S2 = ASC(FF1REC.STATUS2)
    IF ((S2 AND 128) > 0) AND (TAG$(33) <> TAG$(0)) THEN
       CALL GEDCOMLine(33, "", "Y")
    END IF
    IF FF1REC.SEX <> " " THEN CALL GEDCOMLine(34, "", FF1REC.SEX)

840 ' Birth Information
    BLOCPTR = FF1REC.BLOC: BPL = 0: DPL = 0
    IF (((INFOSW AND 2) > 0) AND (BLOCPTR <> 0)) THEN
       CALL FF3GetRec(BLOCPTR)
       IF (FF3PFX.RTYPE = CHR$(12)) AND (FF3Place2.PNEXT <> 0) THEN
          CALL FF3GetRec(FF3Place2.PNEXT)
       END IF
       LSET BLOC$ = BPLACE
       LSET DLOC$ = DPLACE
       ELSE
       LSET BLOC$ = " ": LSET DLOC$ = " "
    END IF
    D1Y = FF1REC.BY: D1MD = FF1REC.BMD: TIM = FF1REC.BTM
    IF FHSv1 THEN
       TD = 36: TT = 38: TP = 39
       ELSE
       TD = 75: TT = 77: TP = 78
    END IF
    BPL = LEN(RTRIM$(BLOC$))
    IF (D1Y = 0) AND (D1MD = 0) AND ((TIM = 0) OR (TAG$(TT) = TAG$(0))) AND (BPL = 0) THEN
       GOTO 850
    END IF
    IF NOT FHSv1 THEN CALL GEDCOMLine(35, "", ""): GL = GL + 1
    IF (D1Y > 0) OR (D1MD > 0) OR ((S1 AND 48) > 0) THEN
       GOSUB 9000: X = LEN(RTRIM$(DFMT$)) + 1
       IF PAFV20 THEN
          MID$(DFMT$, X, 1) = CHR$(32 + (32 - ASC("A")) * ((S1 AND 16) = 16))
          ELSE
          IF FHSv1 THEN
             MID$(DFMT$, X, 1) = FDSTAT$("BD")
          END IF
       END IF
       CALL GEDCOMLine(TD, "", DFMT$)
       IF NOT (PAFV20 OR FHSv1) THEN
          IF ((S1 AND 48) <> 0) AND (TAG$(76) <> TAG$(0)) THEN
             'IF (S1 AND 32) > 0 THEN X$ = "3" ELSE X$ = "0"
             X$ = FDSTAT$("BD")
             CALL GEDCOMLine(76, "", X$)
          END IF
       END IF
    END IF
    IF (TIM > 0) AND (TAG$(TT) <> TAG$(0)) THEN
       IF FHSv1 THEN
          LSET XTIM$ = ZFILL$(TIM, 4)
          ELSE
          LSET XTIM$ = ZFILL$(TIM \ 100, 2) + ":" + ZFILL$(TIM MOD 100, 2)
       END IF
       CALL GEDCOMLine(TT, "", XTIM$)
    END IF
    IF BPL > 0 THEN
       RT = 2: CALL CountRECS
       CALL GEDCOMLine(TP, "", BLOC$)
    END IF
    IF NOT FHSv1 THEN GL = GL - 1

850 ' Death Information
    IF FHSv1 THEN TD = 41: TT = 43: TP = 45
    D1Y = FF1REC.DY: D1MD = FF1REC.DMD: TIM = FF1REC.DTM
    DPL = LEN(RTRIM$(DLOC$))
    IF (D1Y = 0) AND (D1MD = 0) AND ((TIM = 0) OR (TAG$(TT) = TAG$(0))) AND (DPL = 0) AND ((S1 AND 192) = 0) THEN
       GOTO 860
    END IF
    IF NOT FHSv1 THEN CALL GEDCOMLine(40, "", ""): GL = GL + 1
    IF (((D1Y > 0) AND ((D1Y < 9999) OR (FHSv1 AND (D1Y = 9999)))) OR D1MD > 0 OR ((S1 AND 192) > 0)) THEN
       GOSUB 9000: X = LEN(RTRIM$(DFMT$)) + 1
       IF PAFV20 THEN
          MID$(DFMT$, X, 1) = CHR$(32 + (32 - ASC("A")) * ((S1 AND 64) = 64))
          ELSE
          IF FHSv1 THEN
             MID$(DFMT$, X, 1) = FDSTAT$("DD")
          END IF
       END IF
       CALL GEDCOMLine(TD, "", DFMT$)
       IF NOT (PAFV20 OR FHSv1) THEN
          IF ((S1 AND 192) <> 0) AND (TAG$(76) <> TAG$(0)) THEN
             'IF (S1 AND 128) > 0 THEN X$ = "3" ELSE X$ = "0"
             X$ = FDSTAT$("DD")
             CALL GEDCOMLine(76, "", X$)
          END IF
       END IF
    END IF
    IF (TIM > 0) AND (TAG$(TT) <> TAG$(0)) THEN
       IF FHSv1 THEN
          LSET XTIM$ = ZFILL$(TIM, 4)
          ELSE
          LSET XTIM$ = ZFILL$(TIM \ 100, 2) + ":" + ZFILL$(TIM MOD 100, 2)
       END IF
       CALL GEDCOMLine(TT, "", XTIM$)
    END IF
    IF DPL > 0 THEN
       RT = 2: CALL CountRECS
       CALL GEDCOMLine(TP, "", DLOC$)
    END IF
    IF NOT FHSv1 THEN GL = GL - 1

860 ' Name Record Relationships
    IF NOT (FHSv1 OR FHSv2) THEN 865
    ST$ = " "
    X = PtrREC1.WFFID: T = 61: LSET ST$ = FDSTAT$("FID"): GOSUB 861
    X = PtrREC1.WFMID: T = 63: LSET ST$ = FDSTAT$("MID"): GOSUB 861
    X = PtrREC1.WFFSIB: T = 62: GOSUB 861
    X = PtrREC1.WFMSIB: T = 64: GOSUB 861
    X = PtrREC1.WFCH: T = 60: GOSUB 861
    IF ((INFOSW AND 4) > 0) AND (FF1REC.SPOUSE <> 0) THEN
       X& = FFRBA&(FF1REC.SPOUSE, 1)
       IF X& < 32767 THEN
          MRN = LoMISCSEQ(X&)
          ELSE
          MRN = HiMISCSEQ(X& - 32766)
       END IF
       LSET XREF$ = "@M" + FNX$(MRN) + "@"
       CALL GEDCOMLine(51, "", XREF$)
       ELSE
       MRN = 0
    END IF
    GOTO 880

861 ' Process FHS Relationship Field
    IF X > 0 THEN
       LSET XREF$ = "@" + IIPFX$ + FNX$(X) + "@" + ST$
       CALL GEDCOMLine(T, "", XREF$)
       LSET ST$ = ""
    END IF
    RETURN

865 ' Parent Marriage Record
    MRN = PtrREC1.WFPMREC: IF MRN = 0 THEN GOTO 870
    LSET XREF$ = "@" + FIPFX$ + FNX$(MRN) + "@"
    IF NOT PAFV20 THEN
       CALL GEDCOMLine(54, "", XREF$)
       IF (LDSv55 AND ((S1 AND 5) > 0) AND (TAG$(55) <> TAG$(0))) THEN
          GL = GL + 1
          IF (S1 AND 5) = 1 THEN
             LSET XREF$ = "HUSB"
             ELSE
             IF (S1 AND 5) = 4 THEN
                LSET XREF$ = "WIFE"
                ELSE
                LSET XREF$ = "BOTH"
             END IF
          END IF
          CALL GEDCOMLine(55, "", XREF$)
          GL = GL - 1
       END IF
       GOTO 870
    END IF
    CALL GEDCOMLine(54, "", "")
    GL = GL + 1
    CALL GEDCOMLine(31, "", XREF$)
    IF ((S1 AND 1) > 0) AND (TAG$(57) <> TAG$(0)) THEN CALL GEDCOMLine(57, "", "")
    IF ((S1 AND 4) > 0) AND (TAG$(56) <> TAG$(0)) THEN CALL GEDCOMLine(56, "", "")
    GL = GL - 1

870 ' Own Marriage Record(s)
    MRN = ABS(PtrREC1.WFMREC)
871 IF MRN = 0 THEN 875
    LSET XREF$ = "@" + FIPFX$ + FNX$(MRN) + "@"
    IF PAFV20 THEN
       CALL GEDCOMLine(50, "", "")
       GL = GL + 1: CALL GEDCOMLine(31, "", XREF$)
       GL = GL - 1: GOTO 875
    END IF
    CALL GEDCOMLine(50, "", XREF$)
    GET #7, MRN, PtrREC2
    IF SID = PtrREC2.WFHID THEN
       MRN = PtrREC2.WFHOF
       ELSE
       MRN = PtrREC2.WFWOF
    END IF
    MRN = ABS(MRN)
    GOTO 871

875 ' Find Older Sibling with Same Parents (PAF 2.0)
    IF NOT PAFV20 THEN 880
    FID = ABS(PtrREC1.WFFID): MID = ABS(PtrREC1.WFMID): PID = FID
    IF PID > 0 THEN
       SIBPTR = PtrREC1.WFFSIB
       ELSE
       PID = MID
       IF PID > 0 THEN SIBPTR = PtrREC1.WFMSIB ELSE GOTO 880
    END IF
    XID = 0
    WHILE SIBPTR > 0: GET #6, SIBPTR, PtrREC1
       XMID = ABS(PtrREC1.WFMID): XFID = (PtrREC1.WFFID)
       IF (FID = XFID) AND (MID = XMID) THEN
          XID = SIBPTR: SIBPTR = 0
          ELSE
          IF PID = FID THEN SIBPTR = PtrREC1.WFFSIB ELSE SIBPTR = PtrREC1.WFMSIB
       END IF
    WEND
    IF XID = 0 THEN 880
    CALL GEDCOMLine(58, "", "")
    GL = GL + 1: LSET XREF$ = "@" + FNX$(XID) + "@"
    CALL GEDCOMLine(59, "", XREF$)
    GL = GL - 1

880 ' Process NAME Record Comments
    COMPTR = FF1REC.COM: CTAG = 79: GOSUB 980

    ' Process Address Records
    ADRPTR = FF1REC.ADRS: GOSUB 1000

    ' Process Events & Miscellaneous Records
    XPTR = FF1REC.EVENT: GOSUB 2010
    XPTR = FF1REC.EDUC: XTYPE = 1: GOSUB 2020
    XPTR = FF1REC.WORK: XTYPE = 2: GOSUB 2020
    XPTR = FF1REC.MIL: XTYPE = 3: GOSUB 2020
    XPTR = FF1REC.HLTH: XTYPE = 4: GOSUB 2020

899 ' Return from Name Record Routines
    GL = GL - 1
    RETURN

900 ' Produce Family Record Lines
    IF (HIMRN = 0) THEN RETURN
    ERRReturn = 2
    GL = 0
    FOR I = 1 TO HIMRN
    RT = 3: CALL CountRECS
    GET #7, I, PtrREC2: COMPTR = 0
    HID = PtrREC2.WFHID: WID = PtrREC2.WFWID
    HOF = ABS(PtrREC2.WFHOF): WOF = ABS(PtrREC2.WFWOF)
    MRN = PtrREC2.WFMRN
    IF MRN <> 0 THEN
       CALL FF3GetRec(MRN): COMPTR = FF3SPOUSE.COM
    END IF
    LSET XREF$ = "@" + FIPFX$ + FNX$(I) + "@"
    CALL GEDCOMLine(110, XREF$, "")
    GL = GL + 1
    IF (MRN = 0) AND (TAG$(111) <> TAG$(0)) THEN CALL GEDCOMLine(111, "", "")
    IF HID <> 0 THEN
       LSET XREF$ = "@" + IIPFX$ + FNX$(HID) + "@"
       IF PAFV20 THEN
          CALL GEDCOMLine(112, "", "")
          GL = GL + 1: CALL GEDCOMLine(113, "", XREF$)
          IF HOF > 0 THEN
             LSET XREF$ = "@" + FIPFX$ + FNX$(HOF)
             CALL GEDCOMLine(114, "", XREF$)
          END IF
          GL = GL - 1
          ELSE
          CALL GEDCOMLine(112, "", XREF$)
       END IF
    END IF
    IF WID <> 0 THEN
       LSET XREF$ = "@" + IIPFX$ + FNX$(WID) + "@"
       IF PAFV20 THEN
          CALL GEDCOMLine(115, "", "")
          GL = GL + 1: CALL GEDCOMLine(113, "", XREF$)
          IF WOF < 0 THEN WOF = -WOF
          IF WOF > 0 THEN
             LSET XREF$ = "@" + FIPFX$ + FNX$(WOF)
             CALL GEDCOMLine(114, "", XREF$)
          END IF
          GL = GL - 1
          ELSE
          CALL GEDCOMLine(115, "", XREF$)
       END IF
    END IF
    IF FHSv1 OR FHSv2 THEN
       IF HOF > 0 THEN
          LSET XREF$ = "@" + FIPFX$ + FNX$(HOF) + "@"
          CALL GEDCOMLine(130, "", XREF$)
       END IF
       IF WOF > 0 THEN
          LSET XREF$ = "@" + FIPFX$ + FNX$(WOF) + "@"
          CALL GEDCOMLine(131, "", XREF$)
       END IF
    END IF
    NXCH = PtrREC2.WF1CH
    IF NXCH = 0 THEN 920  ' FHS Rules will always have NXCH=0
    IF PAFV20 THEN
       CALL GEDCOMLine(132, "", "")
       LSET XREF$ = "@" + IIPFX$ + FNX$(NXCH) + "@"
       GL = GL + 1: CALL GEDCOMLine(133, "", XREF$)
       GL = GL - 1: GOTO 920
    END IF
    PID = HID: IF PID = 0 THEN PID = WID: IF PID = 0 THEN 920
    GET #6, NXCH, PtrREC1

910 ' Print Line for Each Child for PAF v2.1, LDS v5.5
    LSET XREF$ = "@" + IIPFX$ + FNX$(NXCH) + "@"
    CALL GEDCOMLine(132, "", XREF$)
    S1 = ASC(PtrREC1.WFS1)
    IF (PAFV21 AND ((S1 AND 5) > 0) AND (TAG$(134) <> TAG$(0))) THEN
       IF (S1 AND 5) = 1 THEN
          LSET XREF$ = "HUSB"
          ELSE
          IF (S1 AND 5) = 4 THEN
             LSET XREF$ = "WIFE"
             ELSE
             LSET XREF$ = "BOTH"
          END IF
       END IF
       GL = GL + 1: CALL GEDCOMLine(134, "", XREF$)
       GL = GL - 1
    END IF

915 ' Find Next Child
    IF PID = PtrREC1.WFFID THEN
       NXCH = PtrREC1.WFFSIB
       ELSE
       IF PID = PtrREC1.WFMID THEN NXCH = PtrREC1.WFMSIB ELSE NXCH = 0
    END IF
    IF NXCH > 0 THEN
       GET #6, NXCH, PtrREC1
       ' The Folowing Line assumes Husband<-->Father and Wife<-->Mother
       IF HID = (PtrREC1.WFFID) AND (WID = PtrREC1.WFMID) THEN 910
       GOTO 915
    END IF

920 ' Process Marriage Date & Place Fields
    IF MRN = 0 THEN 975
    MLOC = 0: LSET MPLACE = " ": LSET DVPLACE = " "
    D1Y = FF3SPOUSE.D1Y: D1MD = FF3SPOUSE.D1MD
    D2Y = FF3SPOUSE.D2Y: D2MD = FF3SPOUSE.D2MD
    IF (INFOSW AND 2) > 0 THEN
       MLOC = FF3SPOUSE.MLOC
       IF MLOC <> 0 THEN
          CALL FF3GetRec(MLOC)
          IF (FF3PFX.RTYPE = CHR$(12)) AND (FF3Place2.PNEXT <> 0) THEN
             CALL FF3GetRec(FF3Place2.PNEXT)
          END IF
       END IF
    END IF
    MPL = LEN(RTRIM$(MPLACE))
    DVPL = LEN(RTRIM$(DVPLACE))
    IF (FHSv1 OR FHSv2) THEN
       IF RTRIM$(FF3SPOUSE.STATUS1) <> "" THEN
          CALL GEDCOMLine(121, "", FF3SPOUSE.STATUS1)
       END IF
       IF (D1Y > 0) OR (D1MD > 0) THEN
          GOSUB 9000: CALL GEDCOMLine(122, "", DFMT$)
       END IF
       IF MPL > 0 THEN
          RT = 2: CALL CountRECS
          CALL GEDCOMLine(124, "", MPLACE)
       END IF
       IF RTRIM$(FF3SPOUSE.STATUS2) <> "" THEN
          CALL GEDCOMLine(127, "", FF3SPOUSE.STATUS2)
       END IF
       IF (D2Y > 0) OR (D2MD > 0) THEN
          D1Y = D2Y: D1MD = D2MD: GOSUB 9000
          CALL GEDCOMLine(128, "", DFMT$)
       END IF
       IF DVPL > 0 THEN
          RT = 2: CALL CountRECS
          CALL GEDCOMLine(129, "", DVPLACE)
       END IF
       GOTO 970
    END IF

    IF (D1Y > 0) OR (D1MD > 0) OR (MPL > 0) THEN
       IF LDSv55 THEN
          CALL GEDCOMLine(70, "", ""): GL = GL + 1
          IF RTRIM$(FF3SPOUSE.STATUS1) <> "" THEN
             CALL GEDCOMLine(71, "", FF3SPOUSE.STATUS1)
          END IF
          ELSE
          CALL GEDCOMLine(120, "", ""): GL = GL + 1
       END IF
       IF (D1Y > 0) OR (D1MD > 0) THEN
          GOSUB 9000: CALL GEDCOMLine(75, "", DFMT$)
       END IF
       IF (MPL > 0) THEN
          RT = 2: CALL CountRECS
          CALL GEDCOMLine(78, "", MPLACE)
       END IF
       GL = GL - 1
    END IF

    IF LDSv55 THEN
       IF (D2Y > 0) OR (D2MD > 0) OR (RTRIM$(FF3SPOUSE.STATUS2) <> "") OR (DVPL > 0) THEN
          CALL GEDCOMLine(70, "", ""): GL = GL + 1
          CALL GEDCOMLine(71, "", "TERM")
          IF RTRIM$(FF3SPOUSE.STATUS2) <> "" THEN
             CALL GEDCOMLine(71, "", FF3SPOUSE.STATUS2)
          END IF
          IF ((D2Y > 0) AND ((D2Y < 9999) OR (FHSv1 AND (D2Y = 9999)))) OR (D2MD > 0) THEN
             D1Y = D2Y: D1MD = D2MD
             GOSUB 9000: CALL GEDCOMLine(75, "", DFMT$)
          END IF
          IF (DVPL > 0) THEN
             RT = 2: CALL CountRECS
             CALL GEDCOMLine(78, "", DVPLACE)
          END IF
          GL = GL - 1
       END IF
       GOTO 970
    END IF

    ' Check Divorced Status v2.0-2.1
    X$ = UCX$(LEFT$(FF3SPOUSE.STATUS2, 7))
    DIVORCED = (X$ = "DIVORCE")
    IF NOT DIVORCED THEN 970
    IF (TAG$(125) = TAG$(0)) THEN
       IF TAG$(126) <> TAG$(0) THEN
          CALL GEDCOMLine(126, "", "Y")
       END IF
       GOTO 970
    END IF
    CALL GEDCOMLine(125, "", "")
    GL = GL + 1: D1Y = FF3SPOUSE.D2Y: D1MD = FF3SPOUSE.D2MD
    IF ((D2Y > 0) AND ((D2Y < 9999) OR (FHSv1 AND (D2Y = 9999)))) OR (D2MD > 0) THEN
       D1Y = D2Y: D1MD = D2MD: GOSUB 9000: CALL GEDCOMLine(75, "", DFMT$)
    END IF
    IF DVPL > 0 THEN
       RT = 2: CALL CountRECS
       CALL GEDCOMLine(78, "", DVPLACE)
    END IF
    GL = GL - 1

970 ' Marriage Record Comments
    CTAG = 135: GOSUB 980

    ' Marriage Record Addresses
    ADRPTR = FF3SPOUSE.LOC: GOSUB 1000

975 GL = GL - 1
    A$ = INKEY$
    IF A$ <> "" THEN
       K$ = A$
       IF A$ = CHR$(27) THEN
          I = HIMRN
          ELSE
          SYSVAR.NVAR1 = I: SYSVAR.NVAR2 = HIMRN
          CALL PutMSG("GDCMM010")
       END IF
    END IF
    IF (K$ <> "") OR (GEDLine& - HGEDLine& > 200) THEN
        CALL GEDCOMStats: K$ = "": HGEDLine& = GEDLine&
    END IF

979 NEXT I
    RETURN

980 ' Export Comment Lines
    IF ((INFOSW AND 2 ^ 9) = 0) OR (TAG$(CTAG) = TAG$(0)) OR (COMPTR = 0) THEN RETURN
    HGL = GL: NGL = GL + 1: NTAG = CTAG + 1: S = 0
    OFST = 1: LSET CLINE$ = " "
    ERRReturn = 3
    WHILE COMPTR <> 0
       RT = 10: CALL CountRECS
       CALL FF3GetRec(COMPTR): COMPTR = FF3COMMENT.FPTR
       MID$(CLINE$, OFST, 38) = FF3COMMENT.TEXT
       OFST = INSTR(OFST, CLINE$, "   ") - (MID$(FF3COMMENT.TEXT, 38, 1) = " ")
       IF OFST > 69 THEN
          OFST = 70: FOUND = 0
          WHILE (OFST > 0) AND (NOT FOUND)
             FOUND = (MID$(CLINE$, OFST, 1) = " ")
             IF NOT FOUND THEN OFST = OFST - 1
          WEND
          IF OFST = 0 THEN OFST = 70
          LSET XLINE$ = LEFT$(CLINE$, OFST)
          CALL GEDCOMLine(CTAG, "", XLINE$): S = 1
          LSET CLINE$ = MID$(CLINE$, OFST + 1, 255)
          OFST = INSTR(1, CLINE$, "   ") - (MID$(FF3COMMENT.TEXT, 38, 1) = " ")
          IF NOT (PAFV20 OR FHSv1) THEN GL = NGL
          CTAG = NTAG
       END IF
981 WEND
    IF A <> 27 THEN
       IF OFST > 0 THEN
          LSET XLINE$ = LTRIM$(LEFT$(CLINE$, OFST - 1))
          CALL GEDCOMLine(CTAG, "", XLINE$)
       END IF
    END IF
    GL = HGL
    RETURN

1000 ' Export Address Records
     IF ((INFOSW AND 2 ^ 8) = 0) OR (TAG$(100) = TAG$(0)) THEN RETURN
     AFMT = FF2HDR.REVLVL AND 1
1020 WHILE ADRPTR <> 0
        RT = 9: CALL CountRECS
        CALL GEDCOMLine(100, "", "")
        GL = GL + 1: XDATE$ = ""
        CALL FF2GetRec(ADRPTR): ADRPTR = FF2PFX.NXT
        D1Y = FF2PFX.D1Y: D1MD = FF2PFX.D1MD
1021    IF (D1Y > 0) OR (D1MD > 0) THEN
           GOSUB 9000
           IF FHSv1 OR FHSv2 THEN
              CALL GEDCOMLine(81, "", DFMT$)
              ELSE
              XDATE$ = "from " + LTRIM$(RTRIM$(DFMT$)) + " "
           END IF
        END IF
        D1Y = FF2PFX.D2Y: D1MD = FF2PFX.D2MD
1022    IF (D1Y > 0) OR (D1MD > 0) THEN
           GOSUB 9000
           IF FHSv1 OR FHSv2 THEN
              CALL GEDCOMLine(82, "", DFMT$)
              ELSE
              XDATE$ = XDATE$ + "until " + LTRIM$(RTRIM$(DFMT$))
           END IF
        END IF
        IF XDATE$ <> "" THEN CALL GEDCOMLine(75, "", XDATE$)
1023    SELECT CASE AFMT
           CASE 0: LSET AL1$ = FF2SREC.LINE1: LSET AL2$ = FF2SREC.LINE2
                   LSET CITY$ = FF2SREC.CITY: LSET STATE$ = FF2SREC.STATE
                   LSET ZIP$ = FF2SREC.ZIP: LSET CNTRY$ = "": GOSUB 9200
           CASE 1: LSET AL1$ = FF2LREC.LINE1: LSET AL2$ = FF2LREC.LINE2
                   LSET CITY$ = FF2LREC.CITY: LSET STATE$ = FF2LREC.STATE
                   LSET ZIP$ = FF2LREC.ZIP: LSET CNTRY$ = FF2LREC.COUNTRY
                   LSET PHONE$ = FF2LREC.PHONE
        END SELECT
1024    IF LEN(RTRIM$(AL1$)) > 0 THEN
           CALL GEDCOMLine(101, "", AL1$)
           IF (LEN(RTRIM$(AL2$)) > 0) THEN
              CALL GEDCOMLine(102, "", AL2$)
           END IF
        END IF
1025    IF LEN(RTRIM$(CITY$)) > 0 THEN
           CALL GEDCOMLine(103, "", CITY$)
        END IF
1026    IF LEN(RTRIM$(STATE$)) > 0 THEN
           CALL GEDCOMLine(104, "", STATE$)
        END IF
1027    IF LEN(RTRIM$(ZIP$)) > 0 THEN
           CALL GEDCOMLine(105, "", ZIP$)
        END IF
1028    IF LEN(RTRIM$(CNTRY$)) > 0 THEN
           CALL GEDCOMLine(106, "", CNTRY$)
        END IF
1029    IF LEN(RTRIM$(PHONE$)) > 0 THEN CALL GEDCOMLine(107, "", PHONE$)
        COMPTR = FF2PFX.COM: CTAG = 108: GOSUB 980
        GL = GL - 1
     WEND
     RETURN

2010 ' Export EVENT Records
     IF ((INFOSW AND 2 ^ 7) > 0) AND (TAG$(70) <> TAG$(0)) THEN
        WHILE XPTR <> 0
           RT = 8: CALL CountRECS
           CALL FF3GetRec(XPTR): XPTR = FF3EVENT.NXT
           CALL GEDCOMLine(70, "", ""): GL = GL + 1
           CALL GEDCOMLine(71, "", FF3EVENT.CODE)
           CALL GEDCOMLine(72, "", STR$(ASC(FF3EVENT.IMPORT)))
           IF (FF3EVENT.D1Y > 0) OR (FF3EVENT.D1MD > 0) THEN
              D1Y = FF3EVENT.D1Y: D1MD = FF3EVENT.D1MD: GOSUB 9000
              CALL GEDCOMLine(75, "", DFMT$)
              IF (FF3EVENT.DSTAT <> " ") AND (TAG$(76) <> TAG$(0)) THEN
                 CALL GEDCOMLine(76, "", FF3EVENT.DSTAT)
              END IF
           END IF
           IF RTRIM$(FF3EVENT.PLACE) <> "" THEN
              CALL GEDCOMLine(78, "", FF3EVENT.PLACE)
           END IF
           COMPTR = FF3EVENT.COM: CTAG = 79: GOSUB 980
           A$ = INKEY$
           IF A$ <> "" THEN K$ = A$: IF A$ = CHR$(27) THEN XPTR = 0
           GL = GL - 1
        WEND
     END IF
     RETURN

2020 ' Export EDUC,WORK,MIL,HLTH Records
     SELECT CASE XTYPE
        CASE 1: XTAG = 86
        CASE 2: XTAG = 90
        CASE 3: XTAG = 92
        CASE 4: XTAG = 95
     END SELECT
     IF ((INFOSW AND 2 ^ (XTYPE + 2)) > 0) AND (TAG$(XTAG) <> TAG$(0)) THEN
        WHILE XPTR <> 0
           RT = XTYPE + 3: CALL CountRECS
           CALL FF3GetRec(XPTR): XPTR = FF3PNTRS.NXT
           CALL GEDCOMLine(XTAG, "", ""): GL = GL + 1: XDATE$ = ""
           D1Y = FF3PNTRS.D1Y: D1MD = FF3PNTRS.D1MD
           IF (D1Y > 0) OR (D1MD > 0) THEN
              GOSUB 9000
              IF FHSv1 OR FHSv2 THEN
                 CALL GEDCOMLine(81, "", DFMT$)
                 ELSE
                 XDATE$ = "from " + LTRIM$(RTRIM$(DFMT$)) + " "
              END IF
           END IF
           D1Y = FF3PNTRS.D2Y: D1MD = FF3PNTRS.D2MD
           IF (D1Y > 0) OR (D1MD > 0) THEN
              GOSUB 9000
              IF FHSv1 OR FHSv2 THEN
                 CALL GEDCOMLine(82, "", DFMT$)
                 ELSE
                 XDATE$ = XDATE$ + "until " + LTRIM$(RTRIM$(DFMT$))
              END IF
           END IF
           IF XDATE$ <> "" THEN CALL GEDCOMLine(75, "", XDATE$)
           SELECT CASE XTYPE
              CASE 1: IF RTRIM$(FF3EDUC.EDLEV) <> "" THEN CALL GEDCOMLine(88, "", FF3EDUC.EDLEV)
                      IF RTRIM$(FF3EDUC.EDSUB1) <> "" THEN CALL GEDCOMLine(87, "", FF3EDUC.EDSUB1)
                      IF RTRIM$(FF3EDUC.EDSUB2) <> "" THEN CALL GEDCOMLine(87, "", FF3EDUC.EDSUB2)
                      IF RTRIM$(FF3EDUC.EDDEG) <> "" THEN CALL GEDCOMLine(89, "", FF3EDUC.EDDEG)
              CASE 2: IF RTRIM$(FF3WORK.WORKTP) <> "" THEN CALL GEDCOMLine(91, "", FF3WORK.WORKTP)
              CASE 3: IF RTRIM$(FF3MIL.MILRANK) <> "" THEN CALL GEDCOMLine(93, "", FF3MIL.MILRANK)
                      IF RTRIM$(FF3MIL.MILSTAT) <> "" THEN CALL GEDCOMLine(83, "", FF3MIL.MILSTAT)
              CASE 4: IF RTRIM$(FF3HEALTH.HLTHDIAG) <> "" THEN CALL GEDCOMLine(96, "", FF3HEALTH.HLTHDIAG)
                      IF RTRIM$(FF3HEALTH.HLTHSTAT) <> "" THEN CALL GEDCOMLine(83, "", FF3HEALTH.HLTHSTAT)
           END SELECT
           COMPTR = FF3PNTRS.COM: CTAG = 84: GOSUB 980
           ADRPTR = FF3PNTRS.LOC: GOSUB 1000
           A$ = INKEY$:
           IF A$ <> "" THEN K$ = A$: IF A$ = CHR$(27) THEN XPTR = 0
           GL = GL - 1
        WEND
     END IF
     RETURN

9000 'd1y,d1md-->YYYYMMDD
     IF NOT (PAFV20 OR FHSv1) THEN
        DFMT$ = XRPTDate$(D1MD, D1Y, 0)
        ELSE
        L = 8: LSET DFMT$ = "00000000"
        IF (D1Y > 0) OR (D1MD > 0) THEN
           DM = INT(D1MD / 100): DD = D1MD MOD 100
           IF D1Y > 0 THEN MID$(DFMT$, 1, 4) = RIGHT$("000" + FNX$(D1Y), 4)
           IF DM > 0 THEN MID$(DFMT$, 5, 2) = RIGHT$("0" + FNX$(DM), 2)
           IF DD > 0 THEN MID$(DFMT$, 7, 2) = RIGHT$("0" + FNX$(DD), 2)
        END IF
     END IF
     RETURN

9090 ' MM-DD-YYYY-->dd mmm yyyy
     D1Y = VAL(MID$(DFMT$, 7, 4))
     D1MD = 100 * (VAL(MID$(DFMT$, 1, 2))) + VAL(MID$(DFMT$, 4, 2))
     DFMT$ = XRPTDate$(D1MD, D1Y, 0)
     RETURN

9110 'nfmt$=f1given$/f1surnm$/
     IF CNVTCASE THEN
        CALL ConvertCASE("NAME", FF1REC.GIVEN, S)
        CALL ConvertCASE("NAME", FF1REC.SURNM, S)
     END IF
     IF NOT FHSv1 THEN
        LSET NFMT$ = FF1REC.GIVEN: Z = 1
        Z = INSTR(Z, NFMT$, "   ")
        MID$(NFMT$, Z, 1) = "/": MID$(NFMT$, Z + 1, 30) = FF1REC.SURNM
        NLTH = INSTR(Z, NFMT$, "   ")
        MID$(NFMT$, NLTH, 1) = "/"
        ELSE
        LSET NFMT$ = FF1REC.SURNM
        X = LEN(RTRIM$(NFMT$))
        MID$(NFMT$, X + 1, 255) = ", " + FF1REC.GIVEN
     END IF
     RETURN

9200 ' Format Phone Number (Short Address Format)
     F1 = FF2SREC.FON1: F2 = FF2SREC.FON2: F3 = FF2SREC.FON3
     IF (F1 > 0) OR (F2 > 0) OR (F3 > 0) THEN
        IF FHSv1 THEN
           LSET PHONE$ = "   -   -    "
           MID$(PHONE$, 1, 3) = ZFILL$(F1, 3)
           MID$(PHONE$, 5, 3) = ZFILL$(F2, 3)
           MID$(PHONE$, 9, 4) = ZFILL$(F3, 4)
           ELSE
           LSET PHONE$ = "(000) 000-0000"
           MID$(PHONE$, 2, 3) = ZFILL$(F1, 3)
           MID$(PHONE$, 7, 3) = ZFILL$(F2, 3)
           MID$(PHONE$, 11, 4) = ZFILL$(F3, 4)
        END IF
        ELSE
        LSET PHONE$ = ""
     END IF
     RETURN

19100 ' Build Pointer Workfile
      OKAY = -1: CALL PutMSG("GDCMM011")
      ON LOCAL ERROR GOTO 19105
      SELSET(0) = 0: HIMRN = 0: HIID = 0
      ERRReturn = 1
      GOTO 19110

19105 ' Error Routine
      CALL ERRMessage("GDCMM018", ERR, ERL)
      OKAY = 0
      SELECT CASE ERRReturn
         CASE 1: IX = NDXRECS: RESUME 19111
         CASE 2: SID = HIID: RESUME 19141
         CASE 3: SID = HIID: RESUME 19195
      END SELECT
      RESUME 19199

19110 FOR IX = 1 TO NDXRECS: ID = NDXID(IX)
          SID = SELSET(ID): IF SID = 0 THEN GOTO 19111
          CALL FF1GetRec(ID)
          A$ = INKEY$
          IF A$ <> "" THEN
             IF ASC(A$) = 27 THEN OKAY = 0: IX = NDXRECS: GOTO 19111
             SYSVAR.NVAR1 = IX: SYSVAR.NVAR2 = NDXRECS
             CALL PutMSG("GDCMM012")
          END IF
          FID = SELSET(FF1REC.FID): MID = SELSET(FF1REC.MID)
          SPPTR = FF1REC.SPOUSE
          SEX$ = FF1REC.SEX
          PtrREC1.WFID = ID
          PtrREC1.WFFID = FID: PtrREC1.WFMID = MID: PtrREC1.WFSEX = SEX$
          PtrREC1.WFS1 = FF1REC.STATUS
          FCH = FF1REC.FCH: MCH = FF1REC.MCH
          XID = FF1REC.OLDCH: PID = ID: GOSUB 19200: PtrREC1.WFCH = XID
          IF FID > 0 THEN
             XID = FCH: PID = FID: GOSUB 19200: FCH = XID
             ELSE
             FCH = 0
          END IF
          PtrREC1.WFFSIB = FCH
          IF MID > 0 THEN
             XID = MCH: PID = MID: GOSUB 19200: MCH = XID
             ELSE
             MCH = 0
          END IF
          PtrREC1.WFMSIB = MCH
          PUT #6, SID, PtrREC1: GOSUB 19230
          GET #6, SID, PtrREC1: PtrREC1.WFMREC = SPPTR
          PUT #6, SID, PtrREC1
          HIID = SID
19111 NEXT IX
      IF NOT OKAY THEN GOTO 19199
      IF FHSv1 OR FHSv2 THEN 19199

19140 'Reverse Directions of Sibling Pointers for PAF 2.0
      IF NOT PAFV20 THEN 19150
      CALL PutMSG("GDCMM013")
      ERRReturn = 2
      FOR SID = 1 TO HIID
          GET #6, SID, PtrREC1
          CHPTR = PtrREC1.WFCH: SIBPTR = 0
          WHILE CHPTR <> 0: XID = CHPTR: GET #6, XID, PtrREC1
             IF SID = PtrREC1.WFFID THEN
                CHPTR = PtrREC1.WFFSIB: PtrREC1.WFFSIB = SIBPTR
                ELSE
                CHPTR = PtrREC1.WFMSIB: PtrREC1.WFMSIB = SIBPTR
             END IF
             PUT #6, XID, PtrREC1: SIBPTR = XID
          WEND
          GET #6, SID, PtrREC1: PtrREC1.WFCH = SIBPTR
          PUT #6, SID, PtrREC1
          A$ = INKEY$
          IF A$ <> "" THEN
             IF ASC(A$) = 27 THEN OKAY = 0: SID = HIID
             SYSVAR.NVAR1 = SID: SYSVAR.NVAR2 = HIID
             CALL PutMSG("GDCMM012")
          END IF
19141 NEXT SID: IF NOT OKAY THEN RETURN

19150 ' Find Parent Family Record
      CFREC = 0: ERRReturn = 3
      IF (INFOSW AND 4) = 0 THEN 19199
      CALL PutMSG("GDCMM014"): SYSVAR.NVAR2 = HIID
      FOR SID = 1 TO HIID
          SYSVAR.NVAR1 = SID
          GET #6, SID, PtrREC1
          A$ = INKEY$
          IF A$ <> "" THEN
             IF ASC(A$) = 27 THEN OKAY = 0: SID = HIID: GOTO 19195
             CALL PutMSG("GDCMM012")
          END IF
          IF PtrREC1.WFPMREC <> 0 THEN 19195
          FID = PtrREC1.WFFID: MID = PtrREC1.WFMID
          PID = FID: IF PID = 0 THEN PID = MID: IF PID = 0 THEN 19195
          GET #6, PID, PtrREC1
          PFPTR = 0: XMRN = PtrREC1.WFMREC
          WHILE XMRN > 0
             GET #7, XMRN, PtrREC2
             IF ((FID = PtrREC2.WFHID) AND (MID = PtrREC2.WFWID)) OR ((FID = PtrREC2.WFWID) AND (MID = PtrREC2.WFHID)) THEN
                PFPTR = XMRN: XMRN = 0
                ELSE
                IF PID = PtrREC2.WFHID THEN XMRN = PtrREC2.WFHOF ELSE XMRN = PtrREC2.WFWOF
             END IF
             A$ = INKEY$: IF A$ = CHR$(27) THEN XMRN = 0
          WEND
          IF A$ = CHR$(27) THEN OKAY = 0: SID = HIID: GOTO 19195

          ' Create NEW Family Record for Unmarried Parents
          IF PFPTR = 0 THEN
             CFREC = CFREC + 1: HIMRN = HIMRN + 1
             PFPTR = HIMRN: HOF = 0: WOF = 0
             IF FID > 0 THEN
                GET #6, FID, PtrREC1: HOF = PtrREC1.WFMREC
                PtrREC1.WFMREC = PFPTR
                PUT #6, FID, PtrREC1
             END IF
             IF MID > 0 THEN
                GET #6, MID, PtrREC1: WOF = PtrREC1.WFMREC
                PtrREC1.WFMREC = PFPTR
                PUT #6, MID, PtrREC1
             END IF
             PtrREC2.WFMRN = 0
             PtrREC2.WFHID = FID: PtrREC2.WFWID = MID
             PtrREC2.WFHOF = HOF: PtrREC2.WFWOF = WOF
             PtrREC2.WF1CH = 0
             PUT #7, HIMRN, PtrREC2
          END IF

          ' Update Parent family record for all siblings
          GET #6, PID, PtrREC1: N1CH = 0: SIBPTR = PtrREC1.WFCH
          WHILE SIBPTR > 0
             GET #6, SIBPTR, PtrREC1
             IF ((FID = PtrREC1.WFFID) AND (MID = PtrREC1.WFMID)) OR ((MID = PtrREC1.WFFID) AND (FID = PtrREC1.WFMID)) THEN
                PtrREC1.WFPMREC = PFPTR: PUT #6, SIBPTR, PtrREC1
                IF N1CH = 0 THEN N1CH = SIBPTR
             END IF
             IF PID = PtrREC1.WFFID THEN
                SIBPTR = PtrREC1.WFFSIB
                ELSE
                SIBPTR = PtrREC1.WFMSIB
             END IF
             A$ = INKEY$: IF A$ = CHR$(27) THEN SIBPTR = 0
             IF A$ <> "" THEN
                SYSVAR.NVAR3 = SIBPTR
                CALL PutMSG("Processing [NVAR1] of [NVAR2], Sibling=[NVAR3]")
             END IF
          WEND
          IF A$ = CHR$(27) THEN OKAY = 0: SID = HIID: GOTO 19195

          ' Update Parent family record with first sibling (same FID, MID)
          IF N1CH > 0 THEN
             GET #7, PFPTR, PtrREC2: PtrREC2.WF1CH = N1CH
             PUT #7, PFPTR, PtrREC2
          END IF

19195 ' Next ENTRY
      NEXT SID: IF NOT OKAY THEN GOTO 19199
      IF CFREC > 0 THEN
         SYSVAR.NVAR1 = CFREC
         CALL PutMSG("GDCMM015")
      END IF

19199 ' Return from Building Pointer WorkFiles
      ON LOCAL ERROR GOTO 0
      RETURN

19200 ' Locate First Selected Sibling Parented by PID (Starting with XID)
      WHILE (XID <> 0) AND (SELSET(XID) = 0)
         CALL FF1GetRec(XID)
         IF PID = FF1REC.FID THEN
            XID = FF1REC.FCH
            ELSE
            IF PID = FF1REC.MID THEN
               XID = FF1REC.MCH
               ELSE
               XID = 0
            END IF
         END IF
      WEND
      XID = SELSET(XID)
      RETURN

19230 ' Create Pointer File Records for Existing Marriages
      ' - In Most Recent to Least Recent order for FHS
      ' - In Chronological sequence for PAF or LDS
      FFPTR = 0: LFPTR = 0: IF ((INFOSW AND 4) = 0) THEN 19245
19232 IF SPPTR = 0 THEN 19245
      X& = FFRBA&(SPPTR, 1)
      IF X& < 32767 THEN
         XMRN = LoMISCSEQ(X&)
         ELSE
         XMRN = HiMISCSEQ(X& - 32766)
      END IF
      IF XMRN > 0 THEN GOTO 19240
      HIMRN = HIMRN + 1
      IF X& < 32767 THEN
         LoMISCSEQ(X&) = HIMRN
         ELSE
         HiMISCSEQ(X& - 32766) = HIMRN
      END IF
      CALL FF3GetRec(SPPTR)
      OFPTR = 0: OSPID = 0: OSPOF = 0: HOF = 0: WOF = 0
      IF ID = FF3SPOUSE.SP1ID THEN
         OFPTR = FF3SPOUSE.SP1NXT: OSPID = FF3SPOUSE.SP2ID
         OSPOF = FF3SPOUSE.SP2NXT
         ELSE
         IF ID = FF3SPOUSE.SP2ID THEN
            OFPTR = FF3SPOUSE.SP2NXT: OSPID = FF3SPOUSE.SP1ID
            OSPOF = FF3SPOUSE.SP1NXT
         END IF
      END IF
      PtrREC2.WFMRN = SPPTR
      IF SEX$ = FGENDR$(1) THEN
         HID = SID: SPPTR = OFPTR
         IF (FHSv1 OR FHSv2) THEN HOF = OFPTR ELSE HOF = LFPTR
         WID = SELSET(OSPID): IF WID > 0 THEN WOF = OSPOF
         ELSE
         WID = SID: SPPTR = OFPTR
         IF (FHSv1 OR FHSv2) THEN WOF = OFPTR ELSE WOF = LFPTR
         HID = SELSET(OSPID): IF HID > 0 THEN HOF = OSPOF
      END IF
      PtrREC2.WFHID = HID: PtrREC2.WFWID = WID
      PtrREC2.WFHOF = HOF: PtrREC2.WFWOF = WOF
      PtrREC2.WF1CH = 0
      XMRN = HIMRN
      GOTO 19241

19240 ' Process Other Marriage Records
      GET #7, XMRN, PtrREC2
      IF SID = PtrREC2.WFHID THEN
         SPPTR = PtrREC2.WFHOF
         IF NOT (FHSv1 OR FHSv2) THEN PtrREC2.WFHOF = LFPTR
         ELSE
         IF SID = PtrREC2.WFWID THEN
            SPPTR = PtrREC2.WFWOF
            IF NOT (FHSv1 OR FHSv2) THEN PtrREC2.WFWOF = LFPTR
            ELSE
            SPPTR = 0
         END IF
      END IF
19241 PUT #7, XMRN, PtrREC2
      IF FFPTR = 0 THEN FFPTR = XMRN
      IF FHSv1 OR FHSv2 THEN
         IF LFPTR > 0 THEN
            GET #7, LFPTR, PtrREC2
            IF SID = PtrREC2.WFHID THEN
               PtrREC2.WFHOF = XMRN
               ELSE
               PtrREC2.WFWOF = XMRN
            END IF
            PUT #7, LFPTR, PtrREC2
         END IF
      END IF
      LFPTR = XMRN
      GOTO 19232

19245 ' Finished with Marriage Records for SID
      IF (FHSv1 OR FHSv2) THEN SPPTR = FFPTR ELSE SPPTR = LFPTR
      RETURN

END SUB

SUB GEDCOMImport
SHARED NOPAUSE, PrintERR, INFOSW, GDCMTYPE, GDCMRULE, RT, RECCNT(), TAG$()
SHARED INGL, INTAG$, XREF, XREF$, XID, ST$, XVALUE$, GEDFile$
SHARED TLINE$, MRN, FMT001, HIID, HIMRN, GEDLine&, FIREF, IIREF, INDI, FAMI

     ' Import Information from GEDCOM format dataset
     REDIM XMONTB$(14): FOR X = 1 TO 14: XMONTB$(X) = RMONTB$(X): NEXT X
     CALL GEDCOMOpen("READ"): IF OKAY THEN 1010

1009 ' Exit Before Beginning Import
     CALL PutMSG(""): CLOSE
     FOR X = 1 TO 14: RMONTB$(X) = XMONTB$(X): NEXT X
     EXIT SUB

1010 ' Prepare for Import
     F1RID = 0: F1MAXID = FF1HDR.MAXID: GL = 0: NAMECNT = 0: PrintERR = 0
     BLOC$ = SPACE$(50): DLOC$ = BLOC$: MLOC$ = BLOC$: DVLOC$ = BLOC$
     DFMT$ = SPACE$(15): CLINE$ = SPACE$(150)
     REDIM RECCNT(11), HLIT$(10)
     CALL RptOPEN("GDIM"): IF NOT OKAY THEN 1009
     CALL GetRVAR("SEENOTES", X)
     IF FOUND THEN SEENOTE$ = RVAR$(X) ELSE SEENOTE$ = "see notes"
     CALL GetRVAR("BORN", X)
     IF FOUND THEN BORN$ = RVAR$(X) + ":" ELSE BORN$ = "Born:"
     CALL GetRVAR("MARRIED", X)
     IF FOUND THEN MARRIED$ = RVAR$(X) ELSE MARRIED$ = "Married"
     CALL GetRVAR("DIED", X)
     IF FOUND THEN DIED$ = RVAR$(X) + ":" ELSE DIED$ = "Died:"
     CALL GetRVAR("PARNTREC", X)
     IF FOUND THEN PARNTREC$ = RVAR$(X) ELSE PARNTREC$ = "Parent REC"
     INFOSW = 0
     FOR X = 0 TO 9
         X$ = "GD" + ZFILL$(X, 2)
         CALL RptOptGET(X$, V)
         IF V = 1 THEN INFOSW = INFOSW OR (2 ^ X)
     NEXT X
     IF FF1HDR.MAXID > 0 THEN
        SYSVAR.NVAR1 = FF1HDR.MAXID + 1
        CALL PutMSG("GDCMM020")
        IF ASC(A$) = 27 THEN 1009
     END IF
     CALL GetTAGS("IMPORT"): IF NOT OKAY THEN 1009
     CUST = (GDCMTYPE = 5)

BeginIMPORT:
     ON LOCAL ERROR GOTO 1999
     INTAG$ = SPACE$(LEN(TAG$(0))): ST$ = " ": XVALUE$ = SPACE$(100)
     CALL FamOPEN("WRITE"): IF NOT OKAY THEN 1009
     AFMT = FF2HDR.REVLVL AND 1: XGEDLine& = GEDLine&
     CALL GEDCOMOpen("READQ"): IF (NOT OKAY) OR (A = 27) THEN 1009
     CALL PointerFILES("OPEN"): IF NOT OKAY THEN 1009
     GL = 0: CALL GetLINE: IF NOT OKAY THEN 1009
     IF (INGL > 0) OR (INTAG$ <> TAG$(1)) THEN
        CALL PutMSG("GDCMM022")
        IF (A$ <> SNGLKEY$(1)) THEN 1009
        GOTO 1100
     END IF
     FIREF = 0: IIREF = 0

1050 ' Get Header Lines
     IF A$ = CHR$(27) THEN 1009
     CALL GetLINE: IF (NOT OKAY) OR EOF(5) THEN 1009
     IF (INGL < (GL + 1)) AND (INTAG$ <> TAG$(4)) THEN 1060
     EndOfHDR& = SEEK(5): HDRLine& = GEDLine&
     IF (INTAG$ = TAG$(4)) OR (INTAG$ = TAG$(9)) OR (INTAG$ = "CHEC") THEN 1050
     IF (INTAG$ = TAG$(14)) OR (INTAG$ = TAG$(15)) THEN 1050
     SELECT CASE INTAG$
        CASE TAG$(5): HT = 1
        CASE TAG$(6): HT = 2
        CASE TAG$(8): HT = 5
        CASE TAG$(10): HT = 6
        CASE TAG$(13): HT = 7
        CASE TAG$(18): HT = 9
        CASE TAG$(19):
           IF (HT < 3) OR ((HT > 4) AND (HT < 9)) OR (HT > 10) THEN
              RT = 99: CALL CountRECS: GOTO 1050
           END IF
        CASE TAG$(16): MAXRID = VAL(RTRIM$(XVALUE$)): GOTO 1050
        CASE TAG$(17): MAXMRID = VAL(RTRIM$(XVALUE$)): GOTO 1050
        CASE TAG$(20):
             SELECT CASE RTRIM$(XVALUE$)
               CASE "LDSV55": GDCMRULE = 1
               CASE "PAFV21": GDCMRULE = 2
               CASE "PAFV20": GDCMRULE = 3
               CASE "FHSV1": GDCMRULE = 4
               CASE "FHSV2": GDCMRULE = 5
             END SELECT
        CASE ELSE: RT = 99: CALL CountRECS: GOTO 1050
     END SELECT
1055 ' Put Value in HLIT$()
     HLIT$(HT) = RTRIM$(XVALUE$)
     HT = HT + 1
     GOTO 1050

1060 CALL WinSWITCH(2)
     IF MAXRID > 0 THEN
        INDI = MAXRID: FAMI = MAXMRID: CALL GEDCOMStats
        INDI = 0: FAMI = 0
     END IF
    'Determine GEDCOM Rules
     IF GDCMRULE = 0 THEN
        IF GDCMTYPE = 5 THEN
           CALL PutMSG("GDCMM006"): IF A = 27 THEN 1009
           GDCMRULE = VAL(A$)
           IF GDCMRULE = 4 THEN GDCMRULE = 5
           IF (GDCMRULE < 1) OR (GDCMRULE > 4) THEN GDCMRULE = 1
           ELSE
           GDCMRULE = GDCMTYPE
        END IF
     END IF
     SELECT CASE GDCMRULE
        CASE 1: LDSv55 = -1
        CASE 2: PAFV21 = -1
        CASE 3: PAFV20 = -1
        CASE 4: FHSv1 = -1
        CASE 5: FHSv2 = -1
     END SELECT
     IF (GDCMTYPE < 4) OR ((INFOSW AND 1) = 0) THEN
        X$ = " 0    1JAN 2FEB 3MAR 4APR 5MAY 6JUN 7JUL 8AUG 9SEP10OCT11NOV12DEC13???"
        FOR X = 1 TO 14
            RMONTB$(X) = MID$(X$, 5 * (X - 1) + 1, 5)
        NEXT X
     END IF
     FOR X = 1 TO 14: RMONTB$(X) = UCX$(RMONTB$(X)): NEXT X

1065 ' Show Header Information
     CALL WinCLR
     CALL FmtFIND("GDCMS005")
     X = FMT.BGNFLD
     WHILE X > 0
        FFLD = FLDTAB(X)
        IF FFLD.UTXT <> 0 THEN
           X$ = UTXT$(FFLD.UTXT)
           SELECT CASE FFLD.NAME
              CASE "NAME": X$ = HLIT$(1)
              CASE "ADR1": X$ = HLIT$(2)
              CASE "ADR2": X$ = HLIT$(3)
              CASE "ADR3": X$ = HLIT$(4)
              CASE "PHON": X$ = HLIT$(5)
              CASE "SOUR": X$ = HLIT$(6)
              CASE "DEST": X$ = HLIT$(7)
              CASE "COM1": X$ = HLIT$(8)
              CASE "COM2": X$ = HLIT$(9)
              CASE "COM3": X$ = HLIT$(10)
           END SELECT
           LSET UTXT$(FFLD.UTXT) = X$
        END IF
        X = FLDCHN(X).FWD
     WEND
     CALL WinFORMAT(0)
     CALL PutMSG("GDCMM024")
     IF ((A$ = SNGLKEY$(2)) OR (ASC(A$) = 27)) THEN 1900
     IF NOT ENV.LPNSupport THEN
        IF (NOT (FHSv1 OR FHSv2)) AND ((INFOSW AND 2 ^ 8) > 0) THEN
           CALL PutMSG("GDCMM023")
           IF ASC(A$) = 27 THEN 1009
           NotePLACES = (A$ = SNGLKEY$(1))
        END IF
     END IF
     CALL PutMSG("GDCMM021")
     IF ASC(A$) = 27 THEN 1009
     PrintERR = (A$ = SNGLKEY$(1))
     IF PrintERR THEN CALL PrintOPEN: IF NOT OKAY THEN 1009

     CALL PutMSG("GDCMM025")
     DIM PtrREC1 AS PtrWrkREC1, PtrREC2 AS PtrWrkREC2
     DIM NewPtrREC1 AS PtrWrkREC1, NewPtrREC2 AS PtrWrkREC2
     REDIM IIRNO(MAXRID), FIRNO(MAXMRID)
     HIID = 0: HIMRN = 0: A$ = ""

1070 ' Scan GEDCOM File and Build Pointer Work Files
     IF (INGL > 0) THEN IF SUBSW = 0 THEN 1071 ELSE GOTO 1075
     IF XREF < 0 THEN 1071
     ' Record Level Entries
     IF INTAG$ = TAG$(30) THEN 1072   ' Individual Records
     IF INTAG$ = TAG$(110) THEN 1074  ' Family Records
     IF INTAG$ = TAG$(140) THEN 1080  ' Trailer Record
1071 ' Unprocessed Record Level Lines
     RT = 99: CALL CountRECS: IF (A$ = CHR$(27)) THEN 1995
     GOTO 1079

1072 ' New Individual Record
     IF XREF = 0 THEN 1071
     IF SUBSW > 0 THEN
        IF (HIID > 0) AND (SUBSW < 4) THEN PUT #6, HIID, PtrREC1
        IF (HIMRN > 0) AND (SUBSW > 4) THEN PUT #7, HIMRN, PtrREC2
        SUBSW = 0
     END IF
     IF XREF > MAXRID THEN
        X& = XREF + 100: IF X& > 32766 THEN X& = 32766
        MAXRID = X&: REDIM PRESERVE IIRNO(MAXRID)
     END IF
     IF IIRNO(XREF) <> 0 THEN 1071
     PtrREC1 = NewPtrREC1: PtrREC1.WFID = XREF
     HIID = HIID + 1: IIRNO(XREF) = HIID
     SUBSW = 1
     GOTO 1079

1074 ' New Family Record
     IF XREF = 0 THEN 1071
     IF SUBSW > 0 THEN
        IF (HIID > 0) AND (SUBSW < 4) THEN PUT #6, HIID, PtrREC1
        IF (HIMRN > 0) AND (SUBSW > 4) THEN PUT #7, HIMRN, PtrREC2
        SUBSW = 0
     END IF
     IF XREF > MAXMRID THEN
        X& = XREF + 100: IF X& > 32766 THEN X& = 32766
        MAXMRID = X&: REDIM PRESERVE FIRNO(MAXMRID)
     END IF
     IF FIRNO(XREF) <> 0 THEN 1071
     PtrREC2 = NewPtrREC2
     HIMRN = HIMRN + 1: FIRNO(XREF) = HIMRN
     SUBSW = 5
     GOTO 1079

1075 ' INDI Record Relationship TAGs
     IF SUBSW > 4 THEN 1077
     SELECT CASE INTAG$
        CASE TAG$(31):
             SELECT CASE SUBSW   ' v2.0 family pointers
                CASE 2: PtrREC1.WFMREC = XID
                CASE 3: PtrREC1.WFPMREC = XID
             END SELECT
        CASE TAG$(34): PtrREC1.WFSEX = LEFT$(XVALUE$, 1)
        CASE TAG$(50): IF PAFV20 THEN SUBSW = 2
        CASE TAG$(51): PtrREC1.WFMREC = XID
        CASE TAG$(54): IF (PAFV20 OR LDSv55) THEN SUBSW = 3
        CASE TAG$(55): ' Adoption Status (LDSv55)
                       S1 = ASC(PtrREC1.WFS1)
                       SELECT CASE LEFT$(XVALUE$, 4)
                          CASE "HUSB": S1 = S1 OR 1
                          CASE "WIFE": S1 = S1 OR 4
                          CASE "FATH": S1 = S1 OR 1
                          CASE "MOTH": S1 = S1 OR 4
                          CASE "BOTH": S1 = S1 OR 5
                       END SELECT
                       PtrREC1.WFS1 = CHR$(S1)
        CASE TAG$(56): PtrREC1.WFS1 = CHR$(ASC(PtrREC1.WFS1) OR 1)
        CASE TAG$(57): PtrREC1.WFS1 = CHR$(ASC(PtrREC1.WFS1) OR 4)
        CASE TAG$(59): ' v2.0 older sibling
                       PtrREC1.WFFSIB = XID: PtrREC1.WFMSIB = XID
        CASE TAG$(60): PtrREC1.WFCH = XID
        CASE TAG$(61): PtrREC1.WFFID = XID
                       S1 = ASC(PtrREC1.WFS1)
                       SELECT CASE ST$
                         CASE "*": S1 = S1 OR 1
                         CASE "?": S1 = S1 OR 2
                       END SELECT
                       PtrREC1.WFS1 = CHR$(S1)
        CASE TAG$(62): PtrREC1.WFFSIB = XID
        CASE TAG$(63): PtrREC1.WFMID = XID
                       S1 = ASC(PtrREC1.WFS1)
                       SELECT CASE ST$
                         CASE "*": S1 = S1 OR 4
                         CASE "?": S1 = S1 OR 8
                       END SELECT
                       PtrREC1.WFS1 = CHR$(S1)
        CASE TAG$(64): PtrREC1.WFMSIB = XID
     END SELECT
     GOTO 1079

1077 ' FAMI Record Relationship TAGs
     SELECT CASE INTAG$
        CASE TAG$(112): IF PAFV20 THEN SUBSW = 6 ELSE PtrREC2.WFHID = XID
        CASE TAG$(113): IF SUBSW = 6 THEN PtrREC2.WFHID = XID ELSE PtrREC2.WFWID = XID
        CASE TAG$(114): IF SUBSW = 6 THEN PtrREC2.WFHOF = XID ELSE PtrREC2.WFWOF = XID
        CASE TAG$(115): IF PAFV20 THEN SUBSW = 7 ELSE PtrREC2.WFWID = XID
        CASE TAG$(130): PtrREC2.WFHOF = XID
        CASE TAG$(131): PtrREC2.WFWOF = XID
        CASE TAG$(132): IF PAFV20 THEN 1079    ' CHIL tag
                        ' IF NOT PAFV20 then 1 Line for each CHILd
                        ' Link together Youngest-->Oldest (as with 2.0)
                        CRNO = IIRNO(XID)
                        IF CRNO <> 0 THEN
                           GET #6, CRNO, PtrREC1
                          'PtrREC1.WFFID = PtrREC2.WFHID
                          'PtrREC1.WFMID = PtrREC2.WFWID
                           IF PtrREC2.WFHID <> 0 THEN PtrREC1.WFFSIB = PtrREC2.WF1CH
                           IF PtrREC2.WFWID <> 0 THEN PtrREC1.WFMSIB = PtrREC2.WF1CH
                           PUT #6, CRNO, PtrREC1
                        END IF
                        PtrREC2.WF1CH = XID
        CASE TAG$(133): ' Youngest Child (PAFV20)
                        PtrREC2.WF1CH = XID
        CASE TAG$(134): ' Adoption Status (PAFV21)
                        IF CRNO <> 0 THEN
                           S1 = ASC(PtrREC1.WFS1)
                           SELECT CASE LEFT$(XVALUE$, 4)
                              CASE "HUSB": S1 = S1 OR 1
                              CASE "WIFE": S1 = S1 OR 4
                              CASE "FATH": S1 = S1 OR 1
                              CASE "MOTH": S1 = S1 OR 4
                              CASE "BOTH": S1 = S1 OR 5
                           END SELECT
                           PtrREC1.WFS1 = CHR$(S1)
                           PUT #6, CRNO, PtrREC1
                        END IF
     END SELECT

1079 ' Next Line
     CALL GetLINE: IF (A$ = CHR$(27)) THEN 1995
     IF DISKEOF OR (INTAG$ = TAG$(140)) THEN 1080
     IF A$ = "" THEN A$ = INKEY$
     IF A$ <> "" THEN
        CALL PutMSG("GDCMM025")
        NOPAUSE = 0: K$ = A$
        IF A$ = CHR$(27) THEN CALL GetRESP: IF A$ = CHR$(27) THEN 1995
        A$ = ""
     END IF
     IF (K$ <> "") OR (GEDLine& - HGEDLine& > 200) THEN
        IF (NOT PrintERR) OR (PRT.OPT <> 1) THEN CALL GEDCOMStats
        K$ = "": HGEDLine& = GEDLine&
     END IF
     IF NOT DISKEOF THEN 1070

1080 ' Finish Work Files
     IF SUBSW > 0 THEN
        IF (SUBSW < 5) THEN
           PUT #6, HIID, PtrREC1
           ELSE
           PUT #7, HIMRN, PtrREC2
        END IF
        SUBSW = 0
     END IF
     REDIM MISCSEQ(HIMRN)
     XGEDLine& = GEDLine&
     CALL GEDCOMStats

1090 ' Return to End Of Header Section
     CALL GEDCOMOpen("READQ"): IF NOT OKAY THEN 1995
     SEEK #5, EndOfHDR&
     GEDLine& = HDRLine&
     CALL GetLINE

     ' Display Format for Record Counts
     IF (NOT PrintERR) OR (PRT.OPT <> 1) THEN
        CALL WinCLR: CALL FmtFIND("GDCMS010"): CALL WinFORMAT(0)
     END IF
     HGEDLine& = 0
     CALL PutMSG("GDCMM030")

1100 ' Process Next GEDCOM Record
     IF (A$ = CHR$(27)) THEN 1995
     IF A$ = "" THEN A$ = INKEY$
     IF A$ <> "" THEN
        K$ = A$
        IF A$ = CHR$(27) THEN 1995
        IF (LEN(A$) = 2) AND (A = 81) THEN NOPAUSE = -1 ELSE NOPAUSE = 0
        IF PrintERR AND (PRT.OPT = 1) THEN
           SYSVAR.NVAR1 = II: SYSVAR.NVAR2 = MRID
           CALL PutMSG("GDCMM026")
        END IF
        A$ = ""
     END IF
     IF (K$ <> "") OR (GEDLine& - HGEDLine& > 200) THEN
       IF NOT (PrintERR AND (PRT.OPT = 1)) THEN CALL GEDCOMStats
       K$ = "": HGEDLine& = GEDLine&
     END IF
     IF DISKEOF THEN 1900
     IF INTAG$ = TAG$(140) THEN 1900  ' Trailer Record
     IF (INGL > 0) OR (XREF = 0) THEN 1190
     SUBSW = 0: OLDCH = 0: HID = 0: WID = 0: GL = 0
     IF INTAG$ = TAG$(110) THEN 1200  ' Family Records
     IF INTAG$ = TAG$(30) THEN 1300   ' Individual Records
1190 RT = 99: CALL CountRECS
     CALL GetLINE
     GOTO 1100

1200 ' Process Family Information Record
     IF ((XREF = 0) OR (XREF > MAXMRID) OR ((INFOSW AND 4) = 0)) THEN 1190
     GL = GL + 1
     FIREF = XREF: MRID = FIRNO(XREF): IF MRID = 0 THEN 1190
     GET #7, MRID, PtrREC2: MRN = PtrREC2.WFMRN
     FF3PFX.RTYPE = CHR$(3): CALL FF3GetRec(MRN): IF NOT OKAY THEN 1190
     IF MRN = 0 THEN
        RT = 3: CALL CountRECS
        MRN = FF3PFX.RNO: MISCSEQ(MRID) = MRN
        PtrREC2.WFMRN = MRN: PUT #7, MRID, PtrREC2
        FF3SPOUSE.SRTYPE = CHR$(1): LSET FF3SPOUSE.STATUS1 = MARRIED$
        LSET FF3SPOUSE.STATUS2 = " "
        CALL FF3PutRec
     END IF
     LSET MLOC$ = " ": LSET DVLOC$ = " ": MPL = 0: DVPL = 0
     LAREC = 0: LCOM = 0
     SRNO = MRN: SRTYPE = 3: LAREC = 0

1210 ' Process FAMI Record Lines
     CALL GetLINE
1215 IF INGL < GL THEN 1290
     IF INGL < GL + 1 THEN SUBSW = 0
     IF INGL > (GL - (SUBSW > 0)) THEN 1289
     IF INTAG$ = "CHEC" THEN 1210
     IF INTAG$ = TAG$(111) THEN
        LSET FF3SPOUSE.STATUS1 = PARNTREC$: GOTO 1210
     END IF
     IF (INTAG$ = TAG$(112)) OR (INTAG$ = TAG$(115)) THEN
        IF PAFV20 THEN SUBSW = 99
        GOTO 1210
     END IF
     IF (INTAG$ = TAG$(113)) OR (INTAG$ = TAG$(114)) THEN 1210
     IF FHSv1 OR FHSv2 THEN
        IF INTAG$ = TAG$(121) THEN
           LSET FF3SPOUSE.STATUS1 = XVALUE$: GOTO 1210
        END IF
        IF INTAG$ = TAG$(122) THEN
           FF3SPOUSE.D1Y = 0
           CALL ConvertDATE(XVALUE$, FF3SPOUSE.D1Y, FF3SPOUSE.D1MD, ST$)
           GOTO 1210
        END IF
        IF INTAG$ = TAG$(124) THEN
           MPL = LEN(RTRIM$(XVALUE$))
           IF (MPL = 0) OR ((INFOSW AND 2) = 0) THEN 1289
           LSET MLOC$ = XVALUE$
           GOTO 1210
        END IF
        IF INTAG$ = TAG$(127) THEN
           LSET FF3SPOUSE.STATUS2 = XVALUE$: GOTO 1210
        END IF
        IF INTAG$ = TAG$(128) THEN
           FF3SPOUSE.D2Y = 9999
           CALL ConvertDATE(XVALUE$, FF3SPOUSE.D2Y, FF3SPOUSE.D2MD, ST$)
           GOTO 1210
        END IF
        IF INTAG$ = TAG$(129) THEN
           DVPL = LEN(RTRIM$(XVALUE$))
           IF (DVPL = 0) OR ((INFOSW AND 2) = 0) THEN 1289
           LSET DVLOC$ = XVALUE$
           GOTO 1210
        END IF
     END IF

     ' Marriage Status, Date, Place
     IF INTAG$ = TAG$(70) THEN SUBSW = 13: GOTO 1210
     IF INTAG$ = TAG$(120) THEN SUBSW = 13: GOTO 1210
     IF SUBSW = 13 THEN
        IF INTAG$ = TAG$(71) THEN
           IF RTRIM$(XVALUE$) = "TERM" THEN SUBSW = 14: GOTO 1210
           LSET FF3SPOUSE.STATUS1 = XVALUE$: GOTO 1210
        END IF
        IF (INTAG$ = TAG$(75)) THEN
           FF3SPOUSE.D1Y = 0
           CALL ConvertDATE(XVALUE$, FF3SPOUSE.D1Y, FF3SPOUSE.D1MD, ST$)
           GOTO 1210
        END IF
        IF (INTAG$ = TAG$(78)) THEN
           MPL = LEN(RTRIM$(XVALUE$))
           IF (MPL = 0) OR ((INFOSW AND 2) = 0) THEN 1289
           LSET MLOC$ = XVALUE$
           IF (MPL < (23 - 19 * ENV.LPNSupport)) THEN 1210
           IF NotePLACES THEN GOSUB 1285: GOTO 1210
           GOTO 1289
        END IF
     END IF

     ' "Divorce" Status, Date, Place
     IF (INTAG$ = TAG$(125)) OR (INTAG$ = TAG$(126)) THEN
        SUBSW = 14: LSET FF3SPOUSE.STATUS2 = "Divorced"
        FF3SPOUSE.D2Y = 9999: GOTO 1210
     END IF
     IF SUBSW = 14 THEN
        IF (INTAG$ = TAG$(71)) THEN
           LSET FF3SPOUSE.STATUS2 = XVALUE$: GOTO 1210
        END IF
        IF (INTAG$ = TAG$(75)) THEN
           FF3SPOUSE.D1Y = 9999
           CALL ConvertDATE(XVALUE$, FF3SPOUSE.D1Y, FF3SPOUSE.D2MD, ST$)
           GOTO 1210
        END IF
        IF (INTAG$ = TAG$(78)) THEN
           DVPL = LEN(RTRIM$(XVALUE$))
           IF (DVPL = 0) OR ((INFOSW AND 2) = 0) THEN 1289
           LSET DVLOC$ = XVALUE$
           IF (DVPL < (23 - 19 * ENV.LPNSupport)) THEN 1210
           IF PLACES THEN GOSUB 1285: GOTO 1210
           GOTO 1289
        END IF
     END IF

     ' Next Marriage & Child tags Already Processed
     IF (INTAG$ = TAG$(130)) OR (INTAG$ = TAG$(131)) THEN 1210
     IF (INTAG$ = TAG$(132)) THEN SUBSW = 99: GOTO 1210
     IF (INTAG$ = TAG$(133)) OR (INTAG$ = TAG$(134)) THEN 1210

     ' Address Records
     IF INTAG$ = TAG$(100) THEN
        CALL FF3PutRec
        SUBSW = 15: GOSUB 1600
        CALL FF3GetRec(MRN)
        IF FF3SPOUSE.LOC = 0 THEN FF3SPOUSE.LOC = LAREC
        GOTO 1215
     END IF

1280 ' Family Record Comments
     IF INTAG$ <> TAG$(135) THEN 1289
     SUBSW = 0: IF ((INFOSW AND 2 ^ 8) = 0) THEN 1289
     GOSUB 1285: GOTO 1215

1285 ' Process Comment Line or Long Place Name
     CALL FF3PutRec: CBGN = FF3SPOUSE.COM: GOSUB 1750: LCOM = BPTR
     CALL FF3GetRec(MRN)
     IF CBGN <> 0 THEN FF3SPOUSE.COM = CBGN: CALL FF3PutRec
     RETURN

1289 ' Bad Family Record Line
     RT = 99: CALL CountRECS: IF A$ = CHR$(27) THEN GOTO 1995
     GOTO 1210

1290 ' Write Family Record & Create Place Record
     CALL FF3PutRec: MLOC = 0
     IF ((INFOSW AND 2) <> 0) THEN
        IF (MPL > 0) OR (DVPL > 0) THEN
           IF ((MPL > 22) OR (DVPL > 22)) AND ENV.LPNSupport THEN
              FF3PFX.RTYPE = CHR$(12)
              IF DVPL > 0 THEN
                 CALL FF3GetRec(0): IF NOT OKAY THEN 1399
                 MLOC = FF3PFX.RNO
                 FF3Place2.SRTYPE = CHR$(3): FF3Place2.SRNO = MRN
                 FF3Place2.PTYPE = CHR$(4)
                 LSET FF3Place2.PLACE = MLOC$
                 CALL FF3PutRec
              END IF
              IF MPL > 0 THEN
                 CALL FF3GetRec(0): IF NOT OKAY THEN 1399
                 FF3Place2.PNEXT = MLOC: MLOC = FF3PFX.RNO
                 FF3Place2.SRTYPE = CHR$(3): FF3Place2.SRNO = MRN
                 FF3Place2.PTYPE = CHR$(3)
                 LSET FF3Place2.PLACE = DVLOC$
                 CALL FF3PutRec
              END IF
              ELSE
              IF (MPL > 22) AND (NotePLACES) THEN LSET MLOC$ = SEENOTE$
              IF (DVPL > 22) AND (NotePLACES) THEN LSET DVLOC$ = SEENOTE$
              FF3PFX.RTYPE = CHR$(4): CALL FF3GetRec(0): IF NOT OKAY THEN 1399
              MLOC = FF3PFX.RNO
              FF3PLACE.SRTYPE = CHR$(3): FF3PLACE.SRNO = MRN
              LSET FF3PLACE.PLACE1 = MLOC$: LSET FF3PLACE.PLACE2 = DVLOC$
              CALL FF3PutRec
           END IF
           CALL FF3GetRec(MRN): FF3SPOUSE.MLOC = MLOC
           CALL FF3PutRec
           IF MPL > 0 THEN RT = 2: CALL CountRECS
           IF DVPL > 0 THEN RT = 2: CALL CountRECS
        END IF
     END IF
     GL = 0: GOTO 1100

1300 ' Process Individual Information Record
     IF (XREF > MAXRID) OR (FF1HDR.MAXID > 32765) THEN 1190
     CALL FF1GetNew: IF NOT OKAY THEN 1190
     RT = 1: FIREF = 0: IIREF = XREF: CALL CountRECS
     F1RID = FF1REC.RID: II = IIRNO(XREF)
     GL = GL + 1: BPL = 0: DPL = 0: LFAM = 0: LCOM = 0
     LADRS = 0: LEVNT = 0: LMIL = 0: LEDU = 0: LMED = 0: LOCC = 0
     LSET BLOC$ = " ": LSET DLOC$ = " "
     SRNO = FF1REC.RID: SRTYPE = 1: SUBSW = 0

1310 ' Process INDI Record Lines
     CALL GetLINE
1315 IF (INGL = 0) THEN 1400        ' Write Name Record to File
     IF INGL > (GL - (SUBSW > 0)) THEN 1399
     IF INGL < (GL - (SUBSW > 0)) THEN SUBSW = 0: GL = INGL
     IF INTAG$ = "CHEC" THEN 1310
     IF (INTAG$ = TAG$(31)) AND (INGL = 1) THEN 1310
     IF INTAG$ = TAG$(32) THEN
        CALL ConvertNAME(XVALUE$, FF1REC.SURNM, FF1REC.GIVEN): GOTO 1310
     END IF
     IF INTAG$ = TAG$(33) THEN
        FF1REC.STATUS2 = CHR$((ASC(FF1REC.STATUS2) OR 128)): GOTO 1310
     END IF
     IF INTAG$ = TAG$(34) THEN FF1REC.SEX = LEFT$(XVALUE$, 1): GOTO 1310

     IF FHSv1 THEN
        ' Birth Date, Time, Place
        IF INTAG$ = TAG$(36) THEN
           FF1REC.BY = 0
           CALL ConvertDATE(XVALUE$, FF1REC.BY, FF1REC.BMD, ST$)
           FF1REC.STATUS = CHR$(ASC(FF1REC.STATUS) OR (-16 * (ST$ = "?")) OR (-32 * (ST$ = "!")))
           GOTO 1310
        END IF
        IF INTAG$ = TAG$(38) THEN
           CALL ConvertTIME(XVALUE$, FF1REC.BTM)
           GOTO 1310
        END IF
        IF INTAG$ = TAG$(39) THEN
           BPL = LEN(RTRIM$(XVALUE$))
           IF (BPL = 0) OR ((INFOSW AND 2) = 0) THEN 1399
           LSET BLOC$ = XVALUE$
           GOTO 1310
        END IF
        ' Death Date, Time, Place
        IF INTAG$ = TAG$(41) THEN
           FF1REC.DY = 9999
           CALL ConvertDATE(XVALUE$, FF1REC.DY, FF1REC.DMD, ST$)
           FF1REC.STATUS = CHR$(ASC(FF1REC.STATUS) OR (-64 * (ST$ = "?")) OR (-128 * (ST$ = "!")))
           GOTO 1310
        END IF
        IF INTAG$ = TAG$(43) THEN
           CALL ConvertTIME(XVALUE$, FF1REC.DTM)
           GOTO 1310
        END IF
        IF INTAG$ = TAG$(45) THEN
           DPL = LEN(RTRIM$(XVALUE$))
           IF (DPL = 0) OR ((INFOSW AND 2) = 0) THEN 1399
           LSET DLOC$ = XVALUE$
           GOTO 1310
        END IF
     END IF
        
     ' Process BIRTh Date, Time, Place
     IF INTAG$ = TAG$(35) THEN SUBSW = 1: GOTO 1310
     IF SUBSW = 1 THEN
        IF INTAG$ = TAG$(75) THEN
           FF1REC.BY = 0
           CALL ConvertDATE(XVALUE$, FF1REC.BY, FF1REC.BMD, ST$)
           FF1REC.STATUS = CHR$(ASC(FF1REC.STATUS) OR (-16 * ((ST$ = "A") OR (ST$ = "B") OR (ST$ = "F"))))
           GOTO 1310
        END IF
        IF INTAG$ = TAG$(76) THEN
           ST = VAL(RTRIM$(XVALUE$))
           FF1REC.STATUS = CHR$(ASC(FF1REC.STATUS) OR (-32 * (ST = 3)) OR (-16 * (ST = 0)))
           GOTO 1310
        END IF
        IF INTAG$ = TAG$(77) THEN
           CALL ConvertTIME(XVALUE$, FF1REC.BTM)
           GOTO 1310
        END IF
        IF INTAG$ = TAG$(78) THEN
           BPL = LEN(RTRIM$(XVALUE$))
           IF (BPL = 0) OR ((INFOSW AND 2) = 0) THEN 1399
           LSET BLOC$ = XVALUE$
           IF (BPL < (23 - 19 * ENV.LPNSupport)) THEN 1310
           IF NotePLACES THEN
              CBGN = FF1REC.COM: GOSUB 1750: FF1REC.COM = CBGN
              GOTO 1310
           END IF
           GOTO 1399
        END IF
     END IF

     ' Process DEATh  Date, Time, Place TAGs
     IF INTAG$ = TAG$(40) THEN SUBSW = 2: FF1REC.DY = 9999: GOTO 1310
     IF SUBSW = 2 THEN
        IF INTAG$ = TAG$(75) THEN
           FF1REC.DY = 9999
           CALL ConvertDATE(XVALUE$, FF1REC.DY, FF1REC.DMD, ST$)
           FF1REC.STATUS = CHR$(ASC(FF1REC.STATUS) OR (-64 * ((ST$ = "A") OR (ST$ = "B") OR (ST$ = "F"))))
           GOTO 1310
        END IF
        IF INTAG$ = TAG$(76) THEN
           ST = VAL(RTRIM$(XVALUE$))
           FF1REC.STATUS = CHR$(ASC(FF1REC.STATUS) OR (-128 * (ST = 3)) OR (-64 * (ST = 0)))
           GOTO 1310
        END IF
        IF INTAG$ = TAG$(77) THEN
           CALL ConvertTIME(XVALUE$, FF1REC.DTM)
           GOTO 1310
        END IF
        IF INTAG$ = TAG$(78) THEN
           DPL = LEN(RTRIM$(XVALUE$))
           IF (DPL = 0) OR ((INFOSW AND 2) = 0) THEN 1399
           LSET DLOC$ = XVALUE$
           IF (DPL < (23 - 19 * ENV.LPNSupport)) THEN 1310
           IF NotePLACES THEN
              CBGN = FF1REC.COM: GOSUB 1750: FF1REC.COM = CBGN
              GOTO 1310
           END IF
           GOTO 1399
        END IF
     END IF

     ' Marriage XREF
     IF INTAG$ = TAG$(50) THEN
        IF PAFV20 THEN SUBSW = 3: GOTO 1310
        IF (PAFV21 OR LDSv55) THEN
           GET #6, II, PtrREC1
           X = FIRNO(XID)
           IF X <> 0 THEN
              GET #7, X, PtrREC2
              IF (PtrREC2.WFHID = PtrREC1.WFID) AND ((PtrREC2.WFSP AND 1) = 0) THEN
                 PtrREC2.WFHOF = PtrREC1.WFMREC
                 PtrREC2.WFSP = PtrREC2.WFSP OR 1
                 ELSE
                 IF (PtrREC2.WFWID = PtrREC1.WFID) AND ((PtrREC2.WFSP AND 2) = 0) THEN
                    PtrREC2.WFWOF = PtrREC1.WFMREC
                    PtrREC2.WFSP = PtrREC2.WFSP OR 2
                 END IF
              END IF
              PUT #7, X, PtrREC2
           END IF
           PtrREC1.WFMREC = XID
           PUT #6, II, PtrREC1
           GOTO 1310
        END IF
        GOTO 1399
     END IF
     IF (SUBSW = 3) AND (INTAG$ = TAG$(31)) THEN 1310

     ' Bypass Parent's Marriage Line(s)
     IF INTAG$ = TAG$(54) THEN
        IF (PAFV20 OR PAFV21 OR LDSv55) THEN SUBSW = 4: GOTO 1310
     END IF
     IF (SUBSW = 4) AND (INTAG$ = TAG$(31)) THEN 1310

     ' Check for Christening, Baptismal or Burial information
     IF LEFT$(INTAG$, 3) = "BAP" THEN SUBSW = 8: ECODE$ = "BAPTISM"
     IF LEFT$(INTAG$, 3) = "CHR" THEN SUBSW = 8: ECODE$ = "CHRISTEN"
     IF LEFT$(INTAG$, 3) = "BUR" THEN SUBSW = 8: ECODE$ = "BURIAL"
     IF LEFT$(INTAG$, 3) = "ENG" THEN SUBSW = 8: ECODE$ = "ENGAGED"
     IF SUBSW = 8 THEN
        IF ((INFOSW AND 2 ^ 7) <> 0) THEN
           FF3PFX.RTYPE = CHR$(6): CALL FF3GetRec(0)
           ELSE
           OKAY = 0
        END IF
        IF NOT OKAY THEN 1399
        LSET FF3EVENT.SRTYPE = CHR$(SRTYPE): FF3EVENT.SRNO = SRNO
        LSET FF3EVENT.CODE = ECODE$
        RT = 8: GOSUB 1555: GOTO 1315
     END IF

     ' Adoption TAGs
     IF INTAG$ = TAG$(55) THEN 1310
     IF PAFV20 THEN
        IF (INTAG$ = TAG$(56)) OR (INTAG$ = TAG$(57)) THEN 1310
     END IF
       

     ' Bypass PAFV20,FHSV1,FHSV2 Sibling Tags
     IF PAFV20 THEN
        IF (INTAG$ = TAG$(58)) THEN SUBSW = 99: GOTO 1310
        IF (INTAG$ = TAG$(59)) THEN 1310
     END IF
     IF FHSv1 OR FHSv2 THEN
        IF (INTAG$ = TAG$(60)) OR (INTAG$ = TAG$(62)) OR (INTAG$ = TAG$(64)) THEN 1310
        IF (INTAG$ = TAG$(51)) OR (INTAG$ = TAG$(61)) OR (INTAG$ = TAG$(63)) THEN 1310
     END IF

     ' Miscellaneous Records
     IF (INTAG$ = TAG$(86)) AND ((INFOSW AND 2 ^ 3) <> 0) THEN RT = 4
     IF (INTAG$ = TAG$(90)) AND ((INFOSW AND 2 ^ 4) <> 0) THEN RT = 5
     IF (INTAG$ = TAG$(92)) AND ((INFOSW AND 2 ^ 5) <> 0) THEN RT = 6
     IF (INTAG$ = TAG$(95)) AND ((INFOSW AND 2 ^ 6) <> 0) THEN RT = 7
     IF (RT > 3) AND (RT < 8) THEN GOSUB 1500: GOTO 1315

     ' Event Records
     IF (INTAG$ = TAG$(70)) AND ((INFOSW AND 2 ^ 7) <> 0) THEN
        RT = 8: GOSUB 1550: GOTO 1315
     END IF

     ' Address Records
     IF INTAG$ = TAG$(100) THEN
        SUBSW = 15: LAREC = LADRS: GOSUB 1600: LADRS = LAREC
        IF FF1REC.ADRS = 0 THEN FF1REC.ADRS = LADRS
        GOTO 1315
     END IF

1380 ' Process Name Record Comments
     IF INTAG$ <> TAG$(79) THEN 1399
     SUBSW = 0: IF ((INFOSW AND 2 ^ 8) = 0) THEN 1399
     CBGN = FF1REC.COM: GOSUB 1750: FF1REC.COM = CBGN
     GOTO 1315

1399 ' Error Tag in Name Record
     RT = 99: CALL CountRECS: IF A$ = CHR$(27) THEN 1995
     GOTO 1310

1400 ' Write Name Record to File & Create Place Name Record
     IF (SUBSW = 6) THEN CALL FF3PutRec: SUBSW = 0: OKAY = -1
     IF ((INFOSW AND 2) <> 0) THEN
        BLOC = 0
        IF (BPL > 0) OR (DPL > 0) THEN
           IF ((BPL > 22) OR (DPL > 22)) AND ENV.LPNSupport THEN
              FF3PFX.RTYPE = CHR$(12)
              IF DPL > 0 THEN
                 CALL FF3GetRec(0)
                 IF OKAY THEN
                    BLOC = FF3PFX.RNO
                    FF3Place2.SRTYPE = CHR$(1): FF3Place2.SRNO = F1RID
                    FF3Place2.PTYPE = CHR$(2)
                    LSET FF3Place2.PLACE = DLOC$
                    CALL FF3PutRec
                 END IF
              END IF
              IF BPL > 0 THEN
                 CALL FF3GetRec(0)
                 IF OKAY THEN
                    FF3Place2.PNEXT = BLOC: BLOC = FF3PFX.RNO
                    FF3Place2.SRTYPE = CHR$(1): FF3Place2.SRNO = F1RID
                    FF3Place2.PTYPE = CHR$(1)
                    LSET FF3Place2.PLACE = BLOC$
                    CALL FF3PutRec
                 END IF
              END IF
              ELSE
              IF (BPL > 22) AND (NotePLACES) THEN LSET BLOC$ = SEENOTE$
              IF (DPL > 22) AND (NotePLACES) THEN LSET DLOC$ = SEENOTE$
              FF3PFX.RTYPE = CHR$(4): CALL FF3GetRec(0)
              IF OKAY THEN
                 BLOC = FF3PFX.RNO
                 FF3PLACE.SRTYPE = CHR$(1): FF3PLACE.SRNO = F1RID
                 LSET FF3PLACE.PLACE1 = BLOC$: LSET FF3PLACE.PLACE2 = DLOC$
                 CALL FF3PutRec
              END IF
           END IF
           IF OKAY THEN
              IF BPL > 0 THEN RT = 2: CALL CountRECS
              IF DPL > 0 THEN RT = 2: CALL CountRECS
           END IF
        END IF
     END IF
     IF NOT OKAY THEN RT = 99: CALL CountRECS: BLOC = 0
     FF1REC.BLOC = BLOC
     CALL FF1PutRec
     GL = 0: GOTO 1100

1500 ' Process Miscellaneous Record Types
     FF3PFX.RTYPE = CHR$(RT + 3): CALL FF3GetRec(0)
     IF NOT OKAY THEN
        RT = 99: CALL CountRECS: CALL GetLINE: GOTO 1549
     END IF
     CALL CountRECS: XGL = INGL + 1: XSUB = 0: XRT = RT
     FF3PNTRS.SRTYPE = CHR$(SRTYPE): FF3PNTRS.SRNO = SRNO
     SRTYPE = ASC(FF3PFX.RTYPE): SRNO = FF3PFX.RNO
1510 CALL GetLINE
1511 IF INGL < XGL THEN 1520
     IF INTAG$ = TAG$(75) THEN
        X = INSTR(1, XVALUE$, "from")
        IF X > 0 THEN
           XDATE$ = MID$(XVALUE$, X + 5, 11)
           CALL ConvertDATE(XDATE$, FF3PNTRS.D1Y, FF3PNTRS.D1MD, ST$)
        END IF
        X = INSTR(1, XVALUE$, "until")
        IF X > 0 THEN
           XDATE$ = MID$(XVALUE$, X + 6, 11)
           CALL ConvertDATE(XDATE$, FF3PNTRS.D2Y, FF3PNTRS.D2MD, ST$)
        END IF
        GOTO 1510
     END IF
     IF INTAG$ = TAG$(81) THEN
        CALL ConvertDATE(XVALUE$, FF3PNTRS.D1Y, FF3PNTRS.D1MD, ST$): GOTO 1510
     END IF
     IF INTAG$ = TAG$(82) THEN
        CALL ConvertDATE(XVALUE$, FF3PNTRS.D2Y, FF3PNTRS.D2MD, ST$): GOTO 1510
     END IF
     IF XRT = 4 THEN
        IF INTAG$ = TAG$(87) THEN
           XSUB = XSUB + 1
           SELECT CASE XSUB
              CASE 1: LSET FF3EDUC.EDSUB1 = XVALUE$
              CASE 2: LSET FF3EDUC.EDSUB2 = XVALUE$
              CASE ELSE: GOTO 1515
           END SELECT
           GOTO 1510
        END IF
        IF INTAG$ = TAG$(88) THEN LSET FF3EDUC.EDLEV = XVALUE$: GOTO 1510
        IF INTAG$ = TAG$(89) THEN LSET FF3EDUC.EDDEG = XVALUE$: GOTO 1510
     END IF
     IF XRT = 5 THEN
        IF INTAG$ = TAG$(91) THEN LSET FF3WORK.WORKTP = XVALUE$: GOTO 1510
     END IF
     IF XRT = 6 THEN
        IF INTAG$ = TAG$(93) THEN LSET FF3MIL.MILRANK = XVALUE$: GOTO 1510
        IF INTAG$ = TAG$(83) THEN LSET FF3MIL.MILSTAT = XVALUE$: GOTO 1510
     END IF
     IF XRT = 7 THEN
        IF INTAG$ = TAG$(96) THEN LSET FF3HEALTH.HLTHDIAG = XVALUE$: GOTO 1510
        IF INTAG$ = TAG$(83) THEN LSET FF3HEALTH.HLTHSTAT = XVALUE$: GOTO 1510
     END IF

     ' Address Records
     IF INTAG$ = TAG$(100) THEN
        CALL FF3PutRec: LAREC = FF3PNTRS.LOC
        SUBSW = 15: GOSUB 1600
        CALL FF3GetRec(SRNO)
        IF FF3PNTRS.LOC = 0 THEN FF3PNTRS.LOC = LAREC
        GOTO 1511
     END IF

     'Comments
     IF (INTAG$ = TAG$(79)) AND ((INFOSW AND 2 ^ 8) <> 0) THEN
        CALL FF3PutRec
        CBGN = FF3PNTRS.COM: GOSUB 1750
        CALL FF3GetRec(SRNO): FF3PNTRS.COM = CBGN
        GOTO 1511
     END IF

1515 RT = 99: CALL CountRECS: GOTO 1510

1520 ' Write Miscellaneous Record
     CALL FF3PutRec
     SRTYPE = ASC(FF3PNTRS.SRTYPE): SRNO = FF3PNTRS.SRNO
     SELECT CASE XRT
        CASE 4: IF FF1REC.EDUC = 0 THEN FF1REC.EDUC = FF3PFX.RNO
                X = LEDU: LEDU = FF3PFX.RNO
        CASE 5: IF FF1REC.WORK = 0 THEN FF1REC.WORK = FF3PFX.RNO
                X = LOCC: LOCC = FF3PFX.RNO
        CASE 6: IF FF1REC.MIL = 0 THEN FF1REC.MIL = FF3PFX.RNO
                X = LMIL: LMIL = FF3PFX.RNO
        CASE 7: IF FF1REC.HLTH = 0 THEN FF1REC.HLTH = FF3PFX.RNO
                X = LMED: LMED = FF3PFX.RNO
     END SELECT
     IF X <> 0 THEN
        Y = FF3PFX.RNO
        CALL FF3GetRec(X): FF3PNTRS.NXT = Y: CALL FF3PutRec
     END IF
1549 ' Return to Main Procedure
     RETURN

1550 ' Process Event Record
     FF3PFX.RTYPE = CHR$(6): CALL FF3GetRec(0)
     IF NOT OKAY THEN
        RT = 99: CALL CountRECS: CALL GetLINE: GOTO 1599
     END IF
1555 CALL CountRECS: XGL = INGL + 1
     FF3EVENT.SRTYPE = CHR$(SRTYPE): FF3EVENT.SRNO = SRNO
     SRTYPE = ASC(FF3PFX.RTYPE): SRNO = FF3PFX.RNO
1560 CALL GetLINE
1561 IF INGL < XGL THEN 1570
     IF INTAG$ = TAG$(71) THEN LSET FF3EVENT.CODE = XVALUE$: GOTO 1560
     IF INTAG$ = TAG$(72) THEN
        X = VAL(XVALUE$)
        IF (X < 0) THEN X = 0 ELSE IF (X > 255) THEN X = 255
        LSET FF3EVENT.IMPORT = CHR$(X): GOTO 1560
     END IF
     IF INTAG$ = TAG$(75) THEN
        CALL ConvertDATE(XVALUE$, FF3EVENT.D1Y, FF3EVENT.D1MD, FF3EVENT.DSTAT)
        GOTO 1560
     END IF
     IF INTAG$ = TAG$(76) THEN LSET FF3EVENT.DSTAT = XVALUE$: GOTO 1560
     IF INTAG$ = TAG$(78) THEN
        EPL = LEN(RTRIM$(XVALUE$))
        IF (EPL < (23 - 19 * ENV.LPNSupport)) OR (NOT NotePLACES) THEN
           LSET FF3EVENT.PLACE = XVALUE$
           IF (EPL > 22) AND ENV.LPNSupport THEN
              CALL FF3PutRec
              FF3PFX.RTYPE = CHR$(12): CALL FF3GetRec(0)
              FF3Place2.SRTYPE = CHR$(SRTYPE): FF3Place2.SRNO = SRNO
              LSET FF3Place2.PLACE = XVALUE$
              CALL FF3PutRec: LPLNM = FF3PFX.RNO
              CALL FF3GetRec(SRNO)
              FF3EVENT.LPLNM = LPLNM
           END IF
           GOTO 1560
        END IF
        LSET FF3EVENT.PLACE = SEENOTE$
        LSET INTAG$ = TAG$(79)
     END IF
     IF (INTAG$ = TAG$(79)) AND ((INFOSW AND 2 ^ 8) <> 0) THEN
        CALL FF3PutRec
        CBGN = FF3EVENT.COM: GOSUB 1750
        CALL FF3GetRec(SRNO): FF3EVENT.COM = CBGN
        GOTO 1561
     END IF
     RT = 99: CALL CountRECS: GOTO 1560
1570 ' Write Event Record
     CALL FF3PutRec
     IF FF1REC.EVENT = 0 THEN FF1REC.EVENT = SRNO
     SRTYPE = ASC(FF3EVENT.SRTYPE): SRNO = FF3EVENT.SRNO
     X = FF3PFX.RNO
     IF LEVNT <> 0 THEN
        CALL FF3GetRec(LEVNT): FF3EVENT.NXT = X: CALL FF3PutRec
     END IF
     LEVNT = X
1599 ' Return to Main Procedure
     RETURN


1600 ' Process Address Record
     IF ((INFOSW AND 2 ^ 7) = 0) THEN
        RT = 99: CALL CountRECS: CALL GetLINE
        RETURN
     END IF
     RT = 9: CALL CountRECS
     CALL FF2GetRec(0)
     FF2PFX.SRTYPE = CHR$(SRTYPE): FF2PFX.SRNO = SRNO
     SRTYPE = ASC(FF2PFX.RTYPE): SRNO = FF2PFX.RNO
     XGL = INGL + 1
1610 CALL GetLINE
1611 IF INGL < XGL THEN 1620
     IF INTAG$ = TAG$(75) THEN
        X = INSTR(1, XVALUE$, "from")
        IF X > 0 THEN
           XDATE$ = MID$(XVALUE$, X + 5, 11)
           CALL ConvertDATE(XDATE$, FF2PFX.D1Y, FF2PFX.D1MD, ST$)
        END IF
        X = INSTR(1, XVALUE$, "until")
        IF X > 0 THEN
           XDATE$ = MID$(XVALUE$, X + 6, 11)
           CALL ConvertDATE(XDATE$, FF2PFX.D2Y, FF2PFX.D2MD, ST$)
        END IF
        GOTO 1610
     END IF
     IF INTAG$ = TAG$(81) THEN
        CALL ConvertDATE(XVALUE$, FF2PFX.D1Y, FF2PFX.D1MD, ST$): GOTO 1610
     END IF
     IF INTAG$ = TAG$(82) THEN
        CALL ConvertDATE(XVALUE$, FF2PFX.D2Y, FF2PFX.D2MD, ST$): GOTO 1610
     END IF
     IF INTAG$ = TAG$(101) THEN
        IF AFMT = 0 THEN
           LSET FF2SREC.LINE1 = XVALUE$
           ELSE
           LSET FF2LREC.LINE1 = XVALUE$
        END IF
        GOTO 1610
     END IF
     IF INTAG$ = TAG$(102) THEN
        IF AFMT = 0 THEN
           LSET FF2SREC.LINE2 = XVALUE$
           ELSE
           LSET FF2LREC.LINE2 = XVALUE$
        END IF
        GOTO 1610
     END IF
     IF INTAG$ = TAG$(103) THEN
        IF AFMT = 0 THEN
           LSET FF2SREC.CITY = XVALUE$
           ELSE
           LSET FF2LREC.CITY = XVALUE$
        END IF
        GOTO 1610
     END IF
     IF INTAG$ = TAG$(104) THEN
        IF AFMT = 0 THEN
           LSET FF2SREC.STATE = XVALUE$
           ELSE
           LSET FF2LREC.STATE = XVALUE$
        END IF
        GOTO 1610
     END IF
     IF INTAG$ = TAG$(105) THEN
        IF AFMT = 0 THEN
           LSET FF2SREC.ZIP = XVALUE$
           ELSE
           LSET FF2LREC.ZIP = XVALUE$
        END IF
        GOTO 1610
     END IF
     IF INTAG$ = TAG$(106) THEN
        IF AFMT = 0 THEN 1615
        LSET FF2LREC.COUNTRY = XVALUE$
        GOTO 1610
     END IF
     IF INTAG$ = TAG$(107) THEN
        IF AFMT = 0 THEN
           CALL GetSPHONE(XVALUE$)
           ELSE
           LSET FF2LREC.PHONE = XVALUE$
        END IF
        GOTO 1610
     END IF
     IF (INTAG$ = TAG$(108)) AND ((INFOSW AND 2 ^ 8) <> 0) THEN
        CBGN = FF2PFX.COM: GOSUB 1750: FF2PFX.COM = CBGN
        GOTO 1611
     END IF
1615 RT = 99: CALL CountRECS
     GOTO 1610
1620 ' Write Address Record
     CALL FF2PutRec
     IF LAREC <> 0 THEN
        X = FF2PFX.RNO
        CALL FF2GetRec(LAREC): FF2PFX.NXT = X: CALL FF2PutRec
        CALL FF2GetRec(X)
     END IF
     SRTYPE = ASC(FF2PFX.SRTYPE): SRNO = FF2PFX.SRNO
     LAREC = FF2PFX.RNO
     RETURN
     

1750 ' Process Comment Lines
     GOSUB 1780
     IF RTRIM$(XVALUE$) = "" THEN 1770
1755 ' Next Comment Line
     L = LEN(RTRIM$(XVALUE$))
     IF L > 0 THEN MID$(CLINE$, O, L) = XVALUE$
1760 O = LEN(RTRIM$(CLINE$)) + 1
     IF (O > 38) THEN
        GOSUB 1790: LSET CLINE$ = MID$(CLINE$, 39, 255)
        GOTO 1760
     END IF
1770 ' Next Comment Line
     ' Skip Ahead if processing Long Place Names
     IF ((SUBSW > 0) AND (SUBSW < 3)) OR ((SUBSW > 12) AND (SUBSW < 15)) THEN 1775
     CALL GetLINE
     IF (INTAG$ = TAG$(79)) OR (INTAG$ = TAG$(84)) OR (INTAG$ = TAG$(108)) OR (INTAG$ = TAG$(135)) THEN
        MID$(CLINE$, O, 1) = "\": O = O + 1
        ELSE
        IF ((INTAG$ = TAG$(80)) OR (INTAG$ = TAG$(85)) OR (INTAG$ = TAG$(109)) OR (INTAG$ = TAG$(136))) THEN
           IF O > 1 THEN O = O + 1
           ELSE
           GOTO 1775
        END IF
     END IF
     GOTO 1755

1775 ' Final Line of Comments
     GOSUB 1790: LSET CLINE$ = ""
     IF (CBGN <> 0) AND (CCNT > 1) THEN
        CALL FF3GetRec(CBGN)
        FF3COMMENT.CCNT = CCNT
        CALL FF3PutRec
     END IF
     RETURN

1780 ' Prepare to Create Comment Lines
     CCNT = 0: IF CBGN <> 0 THEN CALL FF3GetRec(CBGN): CCNT = FF3COMMENT.CCNT
     FPTR = CBGN: BPTR = 0: O = 1
     WHILE FPTR <> 0
        BPTR = FPTR: CALL FF3GetRec(FPTR): FPTR = FF3COMMENT.FPTR
     WEND
     IF INTAG$ = "PLAC" THEN
        IF SUBSW = 1 THEN MID$(CLINE$, O, 20) = BORN$: O = O + LEN(BORN$)
        IF SUBSW = 2 THEN MID$(CLINE$, O, 20) = DIED$: O = O + LEN(DIED$)
        IF SUBSW = 13 THEN MID$(CLINE$, O, 20) = MARRIED$ + ":": O = O + LEN(MARRIED$) + 1
     END IF
     RETURN

1790 ' Create Comment Record
     IF (LEFT$(CLINE$, 38) <> SPACE$(38)) THEN
        FF3PFX.RTYPE = CHR$(5): CALL FF3GetRec(0)
        IF OKAY THEN
           RT = 10: CALL CountRECS
           LSET FF3COMMENT.TEXT = CLINE$
           FF3COMMENT.SRTYPE = CHR$(SRTYPE): FF3COMMENT.SRNO = SRNO
           CCNT = CCNT + 1: FF3COMMENT.CCNT = CCNT
           IF CBGN = 0 THEN CBGN = FF3PFX.RNO
           FF3COMMENT.BPTR = BPTR: FF3COMMENT.FPTR = 0
           CALL FF3PutRec
           IF BPTR <> 0 THEN
              FPTR = FF3PFX.RNO
              CALL FF3GetRec(BPTR): FF3COMMENT.FPTR = FPTR
              CALL FF3PutRec: BPTR = FPTR
              ELSE
              BPTR = FF3PFX.RNO
           END IF
        END IF
     END IF
     RETURN

1900 ' Final Steps of Import Operation
     CALL GEDCOMStats
     IF (RECCNT(1) = 0) THEN GOSUB 1995
     IF INTAG$ <> TAG$(140) THEN
        CALL PutMSG("No Trailer Record...Continue? (Y/N)")
        CALL GetRESP: CALL PutMSG("")
     END IF

1950 ' Update Relationship & Family Record Pointers
     ON LOCAL ERROR GOTO 1996
     IF PAFV20 THEN
        CALL PutMSG("GDCMM032")
        FOR I = 1 TO HIID
            GET #6, I, PtrREC1
            X = PtrREC1.WFMREC: Y = FIRNO(X): Z = 0
            WHILE Y <> 0
               GET #7, Y, PtrREC2
               IF PtrREC1.WFID = PtrREC2.WFHID THEN
                  SWAP Z, PtrREC2.WFHOF
                  ELSE
                  IF PtrREC1.WFID = PtrREC2.WFWID THEN
                     SWAP Z, PtrREC2.WFWOF
                     ELSE
                     Z = 0
                  END IF
               END IF
               PUT #7, Y, PtrREC2
               Y = FIRNO(Z): SWAP Z, X
            WEND
            PtrREC1.WFMREC = Z
            PUT #6, I, PtrREC1
            A$ = INKEY$
            IF A$ <> "" THEN
               IF A$ = CHR$(27) THEN
                  I = HIID
                  ELSE
                  SYSVAR.NVAR1 = I: SYSVAR.NVAR2 = HIID
                  CALL PutMSG("GDCMM027")
               END IF
            END IF
        NEXT I
     END IF
     IF A$ = CHR$(27) THEN 1995
     IF NOT (FHSv1 OR FHSv2) THEN
        CALL PutMSG("GDCMM033")
        FOR I = 1 TO HIID
            GET #6, I, PtrREC1: ID = PtrREC1.WFID
            X = PtrREC1.WFMREC: Y = FIRNO(X): Z = 0
            WFP = PtrREC1.WFP
            WHILE Y <> 0
               GET #7, Y, PtrREC2
               Z1 = PtrREC2.WF1CH: Z2 = IIRNO(Z1)
               WHILE Z2 <> 0
                  SWAP Z, Z1
                  GET #6, Z2, PtrREC1
                  IF (ID = PtrREC2.WFHID) AND ((PtrREC1.WFP AND 1) = 0) THEN
                     PtrREC1.WFFID = ID: SWAP PtrREC1.WFFSIB, Z1
                     PtrREC1.WFP = PtrREC1.WFP OR 1
                     ELSE
                     IF (ID = PtrREC2.WFWID) AND ((PtrREC1.WFP AND 2) = 0) THEN
                        PtrREC1.WFMID = ID: SWAP PtrREC1.WFMSIB, Z1
                        PtrREC1.WFP = PtrREC1.WFP OR 2
                        ELSE
                        Z1 = 0
                     END IF
                  END IF
                  PUT #6, Z2, PtrREC1
                  Z2 = IIRNO(Z1)
               WEND
               IF (ID = PtrREC2.WFHID) AND ((WFP AND 4) = 0) THEN
                  X = PtrREC2.WFHOF
                  WFP = WFP OR 4
                  ELSE
                  IF (ID = PtrREC2.WFWID) AND ((WFP AND 8) = 0) THEN
                     X = PtrREC2.WFWOF
                     WFP = WFP OR 4
                     ELSE
                     X = 0
                  END IF
               END IF
               Y = FIRNO(X)
            WEND
            GET #6, I, PtrREC1
            PtrREC1.WFCH = Z: PtrREC1.WFP = WFP
            PUT #6, I, PtrREC1
            A$ = INKEY$
            IF A$ <> "" THEN
               IF A$ = CHR$(27) THEN
                  I = HIID
                  ELSE
                  K$ = A$
                  SYSVAR.NVAR1 = I: SYSVAR.NVAR2 = HIID
                  CALL PutMSG("GDCMM027")
               END IF
            END IF
        NEXT I
     END IF
     IF A$ = CHR$(27) THEN 1995
     CALL PutMSG("GDCMM034")
     FOR I = 1 TO HIID
         GET #6, I, PtrREC1
         ID = F1MAXID + I: CALL FF1GetRec(ID)
         X = IIRNO(PtrREC1.WFFID)
         IF X <> 0 THEN
            FF1REC.FID = X + F1MAXID
            X = IIRNO(PtrREC1.WFFSIB)
            IF X <> 0 THEN FF1REC.FCH = X + F1MAXID
         END IF
         X = IIRNO(PtrREC1.WFMID)
         IF X <> 0 THEN
            FF1REC.MID = X + F1MAXID
            X = IIRNO(PtrREC1.WFMSIB)
            IF X <> 0 THEN FF1REC.MCH = X + F1MAXID
         END IF
         X = IIRNO(PtrREC1.WFCH)
         IF X <> 0 THEN FF1REC.OLDCH = X + F1MAXID
         X = MISCSEQ(FIRNO(PtrREC1.WFMREC))
         IF X <> 0 THEN FF1REC.SPOUSE = X
         FF1REC.STATUS = CHR$(ASC(FF1REC.STATUS) OR ASC(PtrREC1.WFS1))
         CALL FF1PutRec
         A$ = INKEY$
         IF A$ <> "" THEN
            IF A$ = CHR$(27) THEN
               I = HIID
               ELSE
               K$ = A$
               SYSVAR.NVAR1 = I: SYSVAR.NVAR2 = HIID
               CALL PutMSG("GDCMM027")
            END IF
         END IF
     NEXT I
     IF A$ = CHR$(27) THEN 1995
     CALL PutMSG("GDCMM035")
     FOR I = 1 TO HIMRN
         GET #7, I, PtrREC2
         MRN = PtrREC2.WFMRN
         IF MRN <> 0 THEN
            CALL FF3GetRec(MRN)
            X = IIRNO(PtrREC2.WFHID)
            IF X <> 0 THEN
               FF3SPOUSE.SP1ID = X + F1MAXID
               X = MISCSEQ(FIRNO(PtrREC2.WFHOF))
               IF X <> 0 THEN FF3SPOUSE.SP1NXT = X
            END IF
            X = IIRNO(PtrREC2.WFWID)
            IF X <> 0 THEN
               FF3SPOUSE.SP2ID = X + F1MAXID
               X = MISCSEQ(FIRNO(PtrREC2.WFWOF))
               IF X <> 0 THEN FF3SPOUSE.SP2NXT = X
            END IF
            CALL FF3PutRec
            A$ = INKEY$
            IF A$ <> "" THEN
               IF A$ = CHR$(27) THEN
                  I = HIMRN
                  ELSE
                  SYSVAR.NVAR1 = I: SYSVAR.NVAR2 = HIMRN
                  CALL PutMSG("GDCMM027")
                  CALL GetRESP: CALL PutMSG("")
               END IF
            END IF
         END IF
     NEXT I

1995 ' Return
     IF PrintERR THEN CALL RptCLOSE
     CALL FamCLOSE
     ON LOCAL ERROR GOTO 0
     FOR X = 1 TO 14: RMONTB$(X) = XMONTB$(X): NEXT X
     IF XGEDLine& > GEDLine& THEN GEDLine& = XGEDLine&
     CALL PutMSG("GDCMM036")
     EXIT SUB

1996 ' Error Routines
     XERR = ERR: XERL = ERL: CALL GEDCOMStats
     CALL ERRMessage("GDCMM037", XERR, XERL)
     RESUME 1995
1999 XERR = ERR: XERL = ERL: CALL GEDCOMStats
     CALL ERRMessage("GDCMM038", XERR, XERL)
     RESUME 1995

END SUB

REM $STATIC
SUB GEDCOMLine (GDTAG, XREF$, VALUE$)
SHARED INDI, FAMI, GEDLine&, CNVTCASE, S, GDCMRULE, GL, TAG$()
SHARED LDiskNUM, DiskNUM, DISKEOF, GEDFile$

5000 ' Write Line to GEDCOM File
   IF GDCMRULE = 4 THEN X$ = "" ELSE X$ = " "
   IF (GDCMRULE < 3) AND (XREF$ <> "") THEN Y$ = " " ELSE Y$ = ""
   IF CNVTCASE THEN
      IF ((GDTAG > 38) AND ((GDTAG < 46) OR (GDTAG > 70)) AND ((GDTAG < 110) OR (GDTAG > 120)) AND (GDTAG < 136)) THEN
         SELECT CASE GDTAG
            CASE 39: X = 1
            CASE 45: X = 1
            CASE 71: X = 1
            CASE 78: X = 1
            CASE 79: X = 2
            CASE 80: X = 3
            CASE 83: X = 1
            CASE 84: X = 2
            CASE 85: X = 3
            CASE 87: X = 1
            CASE 88: X = 1
            CASE 89: X = 1
            CASE 91: X = 1
            CASE 93: X = 1
            CASE 94: X = 1
            CASE 96: X = 1
            CASE 97: X = 1
            CASE 101: X = 1
            CASE 102: X = 1
            CASE 103: X = 1
            CASE 104: X = 1
            CASE 105: X = 1
            CASE 106: X = 1
            CASE 108: X = 2
            CASE 109: X = 3
            CASE 121: X = 1
            CASE 124: X = 1
            CASE 129: X = 1
            CASE 135: X = 2
            CASE 136: X = 3
         END SELECT
         IF X > 0 THEN
            IF X = 1 THEN
               CALL ConvertCASE("NAME", VALUE$, S)
               ELSE
               IF X = 3 THEN S = 0
               CALL ConvertCASE("TEXT", VALUE$, S)
            END IF
         END IF
      END IF
   END IF
   ON LOCAL ERROR GOTO 5090
5010 ' Write Line
   PRINT #5, FNX$(GL); X$; RTRIM$(XREF$); Y$; TAG$(GDTAG); " "; RTRIM$(VALUE$)
   IF GDTAG = 30 THEN
      INDI = INDI + 1
      ELSE
      IF GDTAG = 110 THEN FAMI = FAMI + 1
   END IF
   GEDLine& = GEDLine& + 1: OKAY = -1
   GOTO 5099

5090 ' Error Writing to GEDCOM File
     XERR = ERR: DISKEOF = -1
     CALL ERRMessage("GDCMM042", ERR, ERL)
     FDISK = (MID$(GEDFile$, 2, 1) = ":")
     IF (XERR <> 61) OR (NOT FDISK) THEN RESUME 5099
     CALL GEDCOMStats
     CALL GEDCOMOpen("WRITENQ")
     IF NOT OKAY THEN RESUME 5099
     RESUME 5010

5099 ' Exit

END SUB

REM $DYNAMIC
SUB GEDCOMList
SHARED TLINE$, GEDFile$, GEDLine&, INDI, FAMI, TAG$()
SHARED LDiskNUM, DiskNUM, DISKEOF

2000 ' Print Exported Data File Listing
     TAG$(0) = ""
     CALL RptOPEN("GDLS"): IF (NOT OKAY) THEN 2069
     CALL GEDCOMOpen("READ"): IF (NOT OKAY) OR (A = 27) THEN EXIT SUB
     CALL PutMSG("GDCMM040")
     IF A$ = CHR$(27) THEN 2069
     OUTLINE = (A$ <> SNGLKEY$(2))
     CALL PrintOPEN: IF NOT OKAY THEN 2069
     C$ = " "
2060 IF NOT DISKEOF THEN
        CALL GetLINE
        IF OUTLINE THEN
           LSET C$ = TLINE$: O = 2 * VAL(C$)
           MID$(PLINE$, O + 1, 255) = TLINE$
           ELSE
           LSET PLINE$ = TLINE$
        END IF
        CALL PrintLINE: IF A <> 27 THEN 2060
     END IF
     CALL RptCLOSE
     IF PRT.AGAIN THEN
        CALL GEDCOMOpen("READQ"): IF OKAY THEN 2060
     END IF
2069 CLOSE
     EXIT SUB
END SUB

SUB GEDCOMOpen (OPT$)
SHARED GEDFile$, LDiskNUM, DiskNUM, DISKEOF, GEDLine&, INDI, FAMI, TAG$()

18000 ' Get GEDFile Prefix
      GED$ = GEDFile$: X = INSTR(GED$, ".")
      IF X > 0 THEN GED$ = LEFT$(GED$, X - 1)
      ' Make Sure File is Closed
      ON LOCAL ERROR GOTO 18001
      CLOSE #5
      GOTO 18002
18001 RESUME 18002
18002 OKAY = 0: XOPT$ = OPT$: L = LEN(XOPT$)
      QUIET = (MID$(XOPT$, L, 1) = "Q")
      IF QUIET THEN L = L - 1: XOPT$ = LEFT$(XOPT$, L)
      NxtDISK = (MID$(XOPT$, L, 1) = "N")
      IF NxtDISK THEN
         XOPT$ = LEFT$(XOPT$, L - 1)
         NDISK = DiskNUM + 1
         GED$ = GED$ + ".G" + ZFILL$(DiskNUM, 2)
         ELSE
         NDISK = 0: GED$ = GED$ + ".GED"
      END IF
      IF NOT QUIET THEN
         CALL FmtFIND("GDCMS004")
         CALL FmtFindFLD("DSNM"): Y = FFLD.UTXT: LSET UTXT$(Y) = GED$
         CALL MsgFORMAT: IF A = 27 THEN OKAY = 0: GOTO 18300
         GED$ = RTRIM$(UTXT$(Y))
      END IF
      FDISK = (MID$(GED$, 2, 1) = ":")
      IF NOT FDISK THEN
         LSET FT$(11) = GED$: DSNM$ = TRANPFX$ + GED$
         ELSE
         DSNM$ = GED$
      END IF
      FDISK = (MID$(DSNM$, 2, 1) = ":")
      IF FDISK THEN
         FDRV$ = LEFT$(DSNM$, 1): FDISK = ((FDRV$ = "A") OR (FDRV$ = "B"))
      END IF
18010 ' Check for More Input Diskettes
      IF FDISK AND NxtDISK AND (XOPT$ = "READ") THEN
         CALL PutMSG("GDCMM044")
         OKAY = (A$ = SNGLKEY$(1))
         IF NOT OKAY THEN
            IF (A$ = CHR$(27)) OR (A$ = SNGLKEY$(2)) THEN 18300
            GOTO 18010
         END IF
      END IF
      SYSVAR.FILENAME = DSNM$
      SELECT CASE XOPT$
         CASE "WRITE": GOTO 18100
         CASE "READ": GOTO 18200
         CASE ELSE: GOTO 18300
      END SELECT

18100 ' Open TRANSFER File for output
      ON LOCAL ERROR GOTO 18140
18110 'IF Floppy Disk, Prompt for New Diskette
      IF FDISK THEN
         LSET SYSVAR.MVAR = FDRV$: SYSVAR.NVAR1 = NDISK
         A$ = INKEY$: WHILE A$ <> "": A$ = INKEY$: WEND
         CALL PutMSG("GDCMM043")
         IF A$ = CHR$(27) THEN OKAY = 0: GOTO 18300
      END IF
      OPEN DSNM$ FOR INPUT AS #5: CLOSE #5
      CALL PutMSG("GDCMM055")
      IF A$ <> SNGLKEY$(1) THEN OKAY = 0: GOTO 18300
      KILL DSNM$: GOTO 18150
18140 RESUME 18150
18150 ON LOCAL ERROR GOTO 18240
      OPEN DSNM$ FOR OUTPUT AS #5: OKAY = -1: GOTO 18300

18200 ' Open Exported DATA File for Input
      ON LOCAL ERROR GOTO 18240
      IF FDISK THEN
         IF ((NOT QUIET) OR (NDISK > 0) OR (LDiskNUM > 0)) THEN
            IF NDISK = 0 THEN X$ = "GDCMM045" ELSE X$ = "GDCMM046"
            LSET SYSVAR.MVAR = FDRV$: SYSVAR.NVAR1 = NDISK
            CALL PutMSG(X$)
            IF A$ = CHR$(27) THEN OKAY = 0: GOTO 18300
         END IF
      END IF
      OPEN DSNM$ FOR INPUT AS #5
      IF EOF(5) THEN
        CALL ErrBEEP(0): CALL PutMSG("GDCMM041")
        GOTO 18000
      END IF
      OKAY = -1: GOTO 18300
18240 IF (NOT QUIET) OR FDISK THEN
         XERR = ERR
         CALL ERRMessage("XXXXM030", XERR, ERL)
         IF FDISK THEN RESUME 18000
      END IF
      OKAY = 0: RESUME 18300

18300 ' Return to Caller
      IF OKAY THEN
         IF (NDISK = 0) THEN
            GEDLine& = 0: IF TAG$(0) <> "" THEN INDI = 0: FAMI = 0
         END IF
         DiskNUM = NDISK: LDiskNUM = DiskNUM
         GEDFile$ = GED$
      END IF
      DISKEOF = NOT OKAY

END SUB

REM $STATIC
SUB GEDCOMStats
SHARED FMT001, INDI, FAMI, GEDLine&, GEDFile$, PrintERR

    ' Show TRANSFER Dataset Stats
    CURFMT = FMT001: FMT = FMTTAB(CURFMT)
    FF = FMT.BGNFLD
    WHILE FF > 0
       FFLD = FLDTAB(FF)
       IF FFLD.UTXT > 0 THEN
          X$ = UTXT$(FFLD.UTXT)
          SELECT CASE FFLD.NAME
             CASE "GDDS": LSET X$ = GEDFile$
             CASE "LINS": LSET X$ = STR$(GEDLine&)
             CASE "INDI": LSET X$ = STR$(INDI)
             CASE "FAMI": LSET X$ = STR$(FAMI)
          END SELECT
          LSET UTXT$(FFLD.UTXT) = X$
       END IF
       FF = FLDCHN(FF).FWD
    WEND
    IF NOT (PrintERR AND (PRT.OPT = 1)) THEN
       CALL WinPREP(1): CALL WinDisplayDATA(FMT001)
       CALL WinPREP(2)
    END IF

END SUB

SUB GetLINE
SHARED INDI, FAMI, INGL, INTAG$, XREF, XID, ST$, XVALUE$, TLINE$, TAG$()
SHARED LDiskNUM, DiskNUM, DISKEOF, GEDLine&, GEDFile$
1800 ' Get Next TRANSFER File Line
     IF (A$ = CHR$(27)) OR (A = 27) THEN INGL = 0: EXIT SUB
     ON LOCAL ERROR GOTO 1899
     OKAY = -1: INGL = 0: XREF = 0: XID = 0: Z = 0: X$ = " "
     LSET INTAG$ = " ": LSET ST$ = " ": LSET XVALUE$ = ""
     IF DISKEOF THEN OKAY = 0: GOTO 1809
     LINE INPUT #5, TLINE$
     GEDLine& = GEDLine& + 1
     TL = LEN(TLINE$)
     X = 1: WHILE (MID$(TLINE$, X, 1) = " ") AND (X < TL): X = X + 1: WEND
     INGL = VAL(MID$(TLINE$, X, 1))
     X = X + 1: GOSUB 1810: IF X > TL THEN Z = 0: GOTO 1809
     IF MID$(TLINE$, X, 1) = "@" THEN
        Y = INSTR(X + 1, TLINE$, "@"): IF Y = 0 THEN Y = TL + 1
        X = X + 1: A = ASC(MID$(TLINE$, X, 1))
        IF (A < 48) OR (A > 57) THEN X = X + 1
        X! = VAL(MID$(TLINE$, X, Y - X)): X = Y + 1
        IF X! > 32766 THEN X! = 0
        XREF = X!
        IF NOT (X > TL) THEN GOSUB 1810
     END IF
     IF NOT (X > TL) THEN
        Y = X: GOSUB 1820
        LSET INTAG$ = MID$(TLINE$, Y, X - Y): GOSUB 1810
     END IF
     IF NOT (X > TL) THEN
        IF MID$(TLINE$, X, 1) = "@" THEN
           Y = INSTR(X + 1, TLINE$, "@"): IF Y = 0 THEN Y = TL + 1
           X = X + 1: A = ASC(MID$(TLINE$, X, 1))
           IF (A < 48) OR (A > 57) THEN X = X + 1
           X! = VAL(MID$(TLINE$, X, Y - X))
           IF X! > 32766 THEN X! = 0
           XID = X!
           IF Y < TL THEN LSET ST$ = MID$(TLINE$, Y + 1, 1)
           X = Y: GOSUB 1810
        END IF
     END IF
     IF NOT (X > TL) THEN
        Z = X
        LSET XVALUE$ = MID$(TLINE$, X, TL - X + 1)
        ELSE
        Z = 0
     END IF
     IF EOF(5) THEN
        DISKEOF = -1
        FDISK = (MID$(GEDFile$, 2, 1) = ":")
        IF FDISK THEN
           FDRV$ = LEFT$(GEDFile$, 1)
           FDISK = ((FDRV$ = "A") OR (FDRV$ = "B"))
           IF FDISK THEN CALL GEDCOMOpen("READNQ")
        END IF
     END IF
     IF TAG$(0) <> "" THEN
        IF INTAG$ = TAG$(30) THEN
           INDI = INDI + 1
           ELSE
           IF INTAG$ = TAG$(110) THEN FAMI = FAMI + 1
        END IF
     END IF
1809 EXIT SUB

1810 WHILE (X < (TL + 1)) AND (MID$(TLINE$, X, 1) = " ")
        X = X + 1
     WEND: RETURN

1820 WHILE (X < (TL + 1)) AND ((MID$(TLINE$, X, 1) <> " ") AND (MID$(TLINE$, X, 1) <> "@"))
        X = X + 1
     WEND: RETURN

1899 ' Error During Input
     CALL ERRMessage("GDCMM039", ERR, ERL)
     OKAY = 0: RESUME 1809

END SUB

SUB GetSPHONE (X$)
    X = 1: L = LEN(X$) + 1
    FF2SREC.FON1 = 0: FF2SREC.FON2 = 0: FF2SREC.FON3 = 0
    IF X < L THEN
       Y = ASC(MID$(X$, X, 1))
       GOSUB GetSPH1: B = X: GOSUB GetSPH2
       IF X > B THEN
          X! = VAL(MID$(X$, B, X - B))
          IF X! < 32767 THEN FF2SREC.FON1 = X!
          GOSUB GetSPH1: B = X: GOSUB GetSPH2
          IF X > B THEN
             X! = VAL(MID$(X$, B, X - B))
             IF X! < 32767 THEN FF2SREC.FON2 = X!
             GOSUB GetSPH1: B = X: GOSUB GetSPH2
             IF X > B THEN
                X! = VAL(MID$(X$, B, X - B))
                IF X! < 32767 THEN FF2SREC.FON3 = X!
             END IF
          END IF
       END IF
    END IF
    EXIT SUB

GetSPH1:
   WHILE (X < L) AND ((Y < 48) OR (Y > 57))
      X = X + 1: IF X < L THEN Y = ASC(MID$(X$, X, 1))
   WEND
   RETURN

GetSPH2:
   WHILE (X < L) AND ((Y > 47) AND (Y < 58))
      X = X + 1: IF X < L THEN Y = ASC(MID$(X$, X, 1))
   WEND
   RETURN

END SUB

SUB GetTAGS (OPT$)
   SHARED INFOSW, GDCMTYPE, GDCMRULE, TAG$()
   CALL PutMSG("GDCMM005")
   OKAY = (A <> 27)
   IF OKAY THEN
      GDCMTYPE = VAL(A$)
      GOSUB BuildTAGS
   END IF
   EXIT SUB

BuildTAGS:
   DIM TagTABLE AS TablePARMS
   CALL TableLOAD("GEDCOM", TNUM, 22): IF NOT OKAY THEN RETURN
   TagTABLE = TABLE(TNUM)
   REDIM TAG$(140)
   SELECT CASE GDCMTYPE
      CASE 1: BT = 15: TL = 4
      CASE 2: BT = 11: TL = 4
      CASE 3: BT = 7: TL = 4
      CASE 4: BT = 5: TL = 2
      CASE 5: BT = 19: TL = 4
   END SELECT
   TAG$(0) = STRING$(TL, ".")
   Y = TagTABLE.FT
   WHILE Y > 0
       TAGNM$ = MID$(TABDATA$(Y), 1, 4)
       TAGVALUE$ = MID$(TABDATA$(Y), BT, TL)
       IF RTRIM$(TAGVALUE$) = "" THEN TAGVALUE$ = TAG$(0)
       IF OPT$ = "EXPORT" THEN TAGVALUE$ = RTRIM$(TAGVALUE$)
       SELECT CASE TAGNM$
          CASE "H000": TAG$(1) = TAGVALUE$   ' Header Record
          CASE "H020": TAG$(4) = TAGVALUE$   ' Submitter's Section
          CASE "H021": TAG$(5) = TAGVALUE$   ' Submitter's Name
          CASE "H030": TAG$(6) = TAGVALUE$   ' Submitter's Address Line 1
          CASE "H031": TAG$(7) = TAGVALUE$   ' Submitter's Address Line 2,3
          CASE "H032": TAG$(8) = TAGVALUE$   ' Submitter's Phone Number
          CASE "H040": TAG$(9) = TAGVALUE$   ' Source System Section (2.1)
          CASE "H041": TAG$(10) = TAGVALUE$  ' Source System Identifier (FHS)
          CASE "H042": TAG$(11) = TAGVALUE$  ' Source Version
          CASE "H043": TAG$(12) = TAGVALUE$  ' Source Full Name
          CASE "H045": TAG$(13) = TAGVALUE$  ' Dest System Identifier
          CASE "H050": TAG$(14) = TAGVALUE$  ' DATE GEDCOM File Built
          CASE "H060": TAG$(15) = TAGVALUE$  ' File Name
          CASE "H061": TAG$(16) = TAGVALUE$  ' Number of II records
          CASE "H062": TAG$(17) = TAGVALUE$  ' Number of II records"
          CASE "H063": TAG$(20) = TAGVALUE$  ' GEDCOM Rules Used
          CASE "H090": TAG$(18) = TAGVALUE$  ' Submitter's Comments
          CASE "H091": TAG$(19) = TAGVALUE$  ' Submitter's Comments (Cont)
          CASE "I000": TAG$(30) = TAGVALUE$  ' Individual Information Record
          CASE "I001": TAG$(31) = TAGVALUE$  ' ID# of Source Record in FHS
          CASE "I010": TAG$(32) = TAGVALUE$  ' Full Name (Given/Surname/)
          CASE "I011": TAG$(33) = TAGVALUE$  ' "Surname Use" Field
          CASE "I015": TAG$(34) = TAGVALUE$  ' Sex Code
          CASE "I020": TAG$(35) = TAGVALUE$  ' Birth Information
          CASE "I021": TAG$(36) = TAGVALUE$  ' Birth Date
          CASE "I023": TAG$(38) = TAGVALUE$  ' Birth Time
          CASE "I025": TAG$(39) = TAGVALUE$  ' Birth Place
          CASE "I030": TAG$(40) = TAGVALUE$  ' Death Information
          CASE "I031": TAG$(41) = TAGVALUE$  ' Death Date
          CASE "I033": TAG$(43) = TAGVALUE$  ' Death Time
          CASE "I035": TAG$(45) = TAGVALUE$  ' Death Place
          CASE "I040": TAG$(50) = TAGVALUE$  ' Marriage Information
          CASE "I041": TAG$(51) = TAGVALUE$  ' "Pointer" to most recent
          CASE "I044": TAG$(54) = TAGVALUE$  ' Parent's Family Record
          CASE "I045": TAG$(55) = TAGVALUE$  ' Adoption: HUSB, WIFE, BOTH
          CASE "I046": TAG$(56) = TAGVALUE$  ' Adoption by Mother (2.0 n.s.)
          CASE "I047": TAG$(57) = TAGVALUE$  ' Adoption by Father (2.0 n.s.)
          CASE "I048": TAG$(58) = TAGVALUE$  ' Sibling Section (2.0)
          CASE "I049": TAG$(59) = TAGVALUE$  ' Next Oldest Sibling
          CASE "I050": TAG$(60) = TAGVALUE$  ' Oldest Child
          CASE "I051": TAG$(61) = TAGVALUE$  ' Father Identifier
          CASE "I052": TAG$(62) = TAGVALUE$  ' Younger Sibling, Same Father
          CASE "I053": TAG$(63) = TAGVALUE$  ' Mother Identifier
          CASE "I054": TAG$(64) = TAGVALUE$  ' Younger Sibling, Same Mother
          CASE "E000": TAG$(70) = TAGVALUE$  ' Event Record
          CASE "E001": TAG$(71) = TAGVALUE$  ' Event Type
          CASE "E002": TAG$(72) = TAGVALUE$  ' Event Importance Level
          CASE "E003": TAG$(75) = TAGVALUE$  ' Event Date
          CASE "E004": TAG$(76) = TAGVALUE$  ' Date Status Field
          CASE "E005": TAG$(77) = TAGVALUE$  ' Event Time
          CASE "E006": TAG$(78) = TAGVALUE$  ' Event Place
          CASE "N000": TAG$(79) = TAGVALUE$  ' First Comment Line
          CASE "N001": TAG$(80) = TAGVALUE$  ' Continuation of Comments
                                             ' Fields for Extra Records
          CASE "X101": TAG$(81) = TAGVALUE$  ' Beginning Effective Date
          CASE "X103": TAG$(82) = TAGVALUE$  ' Ending Effective Date
          CASE "X104": TAG$(83) = TAGVALUE$  ' Ending Status
          CASE "X105": TAG$(84) = TAGVALUE$  ' First Comment Line (Misc Info)
          CASE "X106": TAG$(85) = TAGVALUE$  ' Continued Comments (Misc Info)
          CASE "X120": TAG$(86) = TAGVALUE$  ' Education Record
          CASE "X121": TAG$(87) = TAGVALUE$  ' Subject (may be two)
          CASE "X122": TAG$(88) = TAGVALUE$  ' Education Level
          CASE "X123": TAG$(89) = TAGVALUE$  ' Degree or Certificate Earned
          CASE "X130": TAG$(90) = TAGVALUE$  ' Work Record
          CASE "X131": TAG$(91) = TAGVALUE$  ' Work Type
          CASE "X140": TAG$(92) = TAGVALUE$  ' Military Record
          CASE "X141": TAG$(93) = TAGVALUE$  ' Military Rank
          CASE "X150": TAG$(95) = TAGVALUE$  ' Medical Record
          CASE "X151": TAG$(96) = TAGVALUE$  ' Medical Diagnosis
          CASE "A010": TAG$(100) = TAGVALUE$ ' Address Record
          CASE "A011": TAG$(101) = TAGVALUE$ ' Address Line 1
          CASE "A012": TAG$(102) = TAGVALUE$ ' Address Line(s) continued
          CASE "A013": TAG$(103) = TAGVALUE$ ' City
          CASE "A014": TAG$(104) = TAGVALUE$ ' State/Province
          CASE "A015": TAG$(105) = TAGVALUE$ ' Zip/Postal Code
          CASE "A016": TAG$(106) = TAGVALUE$ ' Country
          CASE "A017": TAG$(107) = TAGVALUE$ ' Phone Number
          CASE "A018": TAG$(108) = TAGVALUE$ ' First Comment Line (ADDRESS)
          CASE "A019": TAG$(109) = TAGVALUE$ ' Next Comment Line  (ADDRESS)
          CASE "F000": TAG$(110) = TAGVALUE$ ' Family/Marriage Record
          CASE "F001": TAG$(111) = TAGVALUE$ ' Created for Unmarried Parents
          CASE "F010": TAG$(112) = TAGVALUE$ ' Husband Identifier
          CASE "F011": TAG$(113) = TAGVALUE$ ' Husb/Wife II Number (2.0)
          CASE "F012": TAG$(114) = TAGVALUE$ ' Husb/Wife Other Marriages (2.0)
          CASE "F015": TAG$(115) = TAGVALUE$ ' Wife's Identifier
          CASE "F020": TAG$(120) = TAGVALUE$ ' Marriage Beginning Event
          CASE "F021": TAG$(121) = TAGVALUE$ ' Beginning Marriage Status
          CASE "F022": TAG$(122) = TAGVALUE$ ' Marriage Date
          CASE "F024": TAG$(124) = TAGVALUE$ ' Place of Marriage
          CASE "F025": TAG$(125) = TAGVALUE$ ' Divorce Event (2.0)
          CASE "F026": TAG$(126) = TAGVALUE$ ' Divorce Flag (2.1)
          CASE "F027": TAG$(127) = TAGVALUE$ ' Ending Status
          CASE "F028": TAG$(128) = TAGVALUE$ ' Marriage Termination Date
          CASE "F029": TAG$(129) = TAGVALUE$ ' Place of Marriage Termination
          CASE "F030": TAG$(130) = TAGVALUE$ ' Husband's next earlier marriage
          CASE "F031": TAG$(131) = TAGVALUE$ ' Wife's next earlier marriage
          CASE "F040": TAG$(132) = TAGVALUE$ ' Child Identifier
          CASE "F041": TAG$(133) = TAGVALUE$ ' Youngest Child's Identifier
          CASE "F042": TAG$(134) = TAGVALUE$ ' Adoption (2.1)
          CASE "F090": TAG$(135) = TAGVALUE$ ' First Comment Line   (FAMILY)
          CASE "F091": TAG$(136) = TAGVALUE$ ' Comment Continuation (FAMILY)
          CASE "T000": TAG$(140) = TAGVALUE$ ' Trailer (END) Record
       END SELECT
       Y = TABDATACHN(Y).FWD
   WEND
   CALL TableFREE("GEDCOM")
   RETURN

END SUB

SUB PointerFILES (OPT$)
   SELECT CASE OPT$
      CASE "OPEN": GOSUB 19000
      CASE "DELETE": GOSUB 19050
   END SELECT
   EXIT SUB

19000 ' Open Pointer Workfiles
      OKAY = 0
      WF1$ = WRKFPFX$ + "FH$PTR1": WF2$ = WRKFPFX$ + "FH$PTR2"
      GOSUB 19050
      ON LOCAL ERROR GOTO 19010
      SYSVAR.MVAR = WF1$
      OPEN WF1$ FOR RANDOM ACCESS READ WRITE AS #6 LEN = 20
      SYSVAR.MVAR = WF2$
      OPEN WF2$ FOR RANDOM ACCESS READ WRITE AS #7 LEN = 14
      OKAY = -1
      GOTO 19019
19010 ' Error Opening Work Files
      CALL ERRMessage("XXXXM030", ERR, ERL)
      IF A = 27 THEN RESUME 19019 ELSE RESUME 19000
19019 RETURN

19050 ' KILL Pointer Workfiles
19055 ON LOCAL ERROR GOTO 19056
      KILL WF1$: GOTO 19060
19056 RESUME 19060
19060 ON LOCAL ERROR GOTO 19061
      KILL WF2$: GOTO 19070
19061 RESUME 19070
19070 RETURN

END SUB

