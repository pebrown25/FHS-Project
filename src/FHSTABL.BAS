DECLARE FUNCTION SFILL$ (X%, Y%)
   REM $INCLUDE: 'FHSCOMON.BAS'
   PN$ = " FAMILY HISTORY SYSTEM - FHSTABL - Table Customization program"
   CY$ = " (C) Copyright 1995,1996,1997 by Phillip E. Brown"
   IF CONFIG.BP = 0 THEN RUN "FHSINIT"
   FUN = 0: FMTNAME$ = SPACE$(8): OKAY = 0
   TNCHG = 0: GS$ = " ": TS$ = "   ": TV$ = "   ": L$ = " "
   CALL TableLOAD("ALL", Z, 0)
   GOTO TSINIT0
   CALL TableLOAD("RELCODES", Z, 0)
   TABLE(Z).XL = TABLE(Z).XL + 1: TABLE(Z).ML = TABLE(Z).XL
   X = TABLE(Z).FT: X$ = SPACE$(TABLE(Z).XL): Z = 1
   WHILE X > 0
      LSET X$ = SFILL$(Z, 2): Z = Z + 1
      MID$(X$, 3, 3) = MID$(TABDATA$(X), 5, 3)
      MID$(X$, 6, 12) = MID$(TABDATA$(X), 8, 9)
      MID$(X$, 18, 45) = MID$(TABDATA$(X), 17, 45)
      TABDATA$(X) = X$
      X = TABDATACHN(X).FWD
   WEND
TSINIT0:
   DIM MASTTABLE AS TablePARMS
   MASTTABLE = TABLE(1)
   LSET MASTTABLE.FMTNAME = "TABLS011": MASTTABLE.FMTNUM = 0
   MASTTABLE.KEYNAME = "KEY"
   MASTTABLE.CL = 0: MASTTABLE.KO = 1: MASTTABLE.KL = 8
   REDIM MTDATA$(TABMAX), MTDATACHN(TABMAX) AS ChainPTRS
   CALL ChnINIT(MTDATACHN(), 1, TABMAX)
TSINIT:
   MTE$ = SPACE$(255)
   FOR X = 1 TO TABMAX
       LSET MTE$ = TABLE(X).TABNAME
       RSET TS$ = LTRIM$(STR$(TDFIndex(X).ENTCNT))
       MID$(MTE$, 13, 3) = TS$
       IF (TDFIndex(X).CHG AND 128) = 128 THEN
          LSET GS$ = "+"
       ELSE
          LSET GS$ = " "
       END IF
       MID$(MTE$, 12, 1) = GS$
       RSET TV$ = LTRIM$(STR$(TDFIndex(X).CHG AND 127))
       MID$(MTE$, 9, 3) = TV$
       Y = TABLE(1).FT: FOUND = 0
       WHILE (Y > 0) AND (NOT FOUND)
          FOUND = (TABLE(X).TABNAME = MID$(TABDATA$(Y), 1, 8))
          IF NOT FOUND THEN Y = TABDATACHN(Y).FWD
       WEND
       IF FOUND THEN MID$(MTE$, 16, 255) = MID$(TABDATA$(Y), 9, 255)
       MTDATA$(X) = RTRIM$(MTE$)
   NEXT X
40 ' Format Display
   IF FUN = 0 THEN CALL WinCLR
   CALL FmtFIND("TABLS010")
   CALL FmtFindFLD("TDF"): IF FOUND THEN UTXT$(FFLD.UTXT) = CONFIG.TDF
   CALL FmtFindFLD("ECT")
   IF FOUND THEN UTXT$(FFLD.UTXT) = LTRIM$(STR$(TDFHdr.ENTCNT))
   CALL FmtFindFLD("EOF")
   IF FOUND THEN UTXT$(FFLD.UTXT) = LTRIM$(STR$(TDFHdr.EOFRBA))
   CALL WinFORMAT(0)
   CALL TableOPEN(MASTTABLE)
   CALL TableSHOW(MASTTABLE, MTDATA$(), MTDATACHN())
200 ' Get Processing Option
    PGMRTRN$ = "FHSCUST"
    CALL PutMSG(" ")
    CALL FmtFIND("XXXXS091")
    CALL WinFORMAT(WIN.OFMT)
250 CALL TableSELECT(MASTTABLE, MTDATA$(), MTDATACHN())
    SELECT CASE A
       CASE 13: FUN$ = "SEL"
       CASE 59: FUN$ = "F1"
       CASE 64: FUN$ = "F6"
       CASE 66: CALL DOSShell: GOTO 40
       CASE 67: FUN$ = "F9"
       CASE 94: ENV.EDMAST = NOT ENV.EDMAST: GOTO 250
       CASE 99: FUN$ = "F6": PGMRTRN$ = "FHSPRTC"
       CASE 109: FUN$ = "F6": CALL RptLOAD("TABL"): PGMRTRN$ = "FHSRPTS"
       CASE ELSE: SOUND BP, DUR: GOTO 250
    END SELECT
    CALL OPTHilite(FUN$, "", 6)
    SELECT CASE FUN$
       CASE "F1": GOTO 500
       CASE "F6": IF PGMRTRN$ <> "FHSCUST" THEN GOTO 1000
       CASE "F9": GOTO 1000
    END SELECT
300 ' Update TABLE Data
    TN = MASTTABLE.CT: TNCHG = 0
    TNAME$ = TABLE(TN).TABNAME
    IF FUN$ = "F6" THEN 600     ' Process PRINT Option
    X = WIN.ULR: WIN.ULR = MASTTABLE.TL - 1: CALL WinCLR: WIN.ULR = X
    CALL FmtFIND("TABLS020")
    CALL FmtFindFLD("TBN"): IF FOUND THEN UTXT$(FFLD.UTXT) = TNAME$
    CALL WinFORMAT(0)
    TN$ = RIGHT$("0" + LTRIM$(STR$(TN)), 2)
    LSET FMTNAME$ = "TABLS1" + TN$
    CALL FmtFIND(FMTNAME$): CALL WinFORMAT(0)
    LSET FMTNAME$ = "TABLS2" + TN$
    TABLE(TN).FMTNAME = FMTNAME$: TABLE(TN).FMTNUM = 0
    TABLE(TN).KEYNAME = "KEY"
    CALL TableOPEN(TABLE(TN))
    CALL TableSHOW(TABLE(TN), TABDATA$(), TABDATACHN())
    IF (TNAME$ = "UCTRAN  ") OR (TNAME$ = "SOUNDEX ") OR (TNAME$ = "EVENTS  ") OR (TNAME$ = "SORTSEQ ") THEN
       XEM = ENV.EDMAST: ENV.EDMAST = -1
    END IF
340 ' SET Option Format
    IF ENV.EDMAST THEN LSET FMTNAME$ = "XXXXS093" ELSE LSET FMTNAME$ = "XXXXS092"
    CALL FmtFIND(FMTNAME$)
350 ' SELECT Entry For Update
    CALL OPTDisplay(0)
    IF TABLE(TN).SIZE = 0 THEN A = 255: GOTO 375
    CALL TableSELECT(TABLE(TN), TABDATA$(), TABDATACHN())
    FUN$ = ""
  '  IF ENV.EDMAST AND ((A > 61) AND (A < 63)) AND (TNAME$ = "TABLES  ") THEN
  '     SOUND BP, DUR
  '     CALL PutMSG("TABLM002")
  '     CALL GetRESP: CALL PutMSG("")
  '     GOTO 350
  '  END IF
375 ' Determine Response Option
    FUN$ = ""
    SELECT CASE A
       CASE 13: FUN$ = "CHG"
       CASE 27: FUN$ = "F9"
       CASE 67: FUN$ = "F9"
       CASE 59: IF ENV.EDMAST THEN FUN$ = "MOD"
       CASE 60: IF ENV.EDMAST THEN FUN$ = "ADD"
       CASE 61: IF ENV.EDMAST THEN FUN$ = "DEL"
       CASE 62: IF ENV.EDMAST THEN FUN$ = "MOV"
       CASE 94: ENV.EDMAST = NOT ENV.EDMAST: GOTO 340
       CASE 255: FUN$ = "ADD"
    END SELECT
    IF FUN$ = "" THEN SOUND BP, DUR: GOTO 350
    CALL OPTHilite(FUN$, "", 6)
    IF (FUN$ = "DEL") THEN
       CALL PutMSG("TABLM003")
       IF A$ <> SNGLKEY$(1) THEN GOTO 350
       ELSE
       IF (FUN$ = "ADD") THEN
          IF TABDATAFREE = 0 THEN
             TABDATAFREE = TABDATAMAX + 1
             TABDATAMAX = TABDATAFREE
             REDIM PRESERVE TABDATA$(TABDATAMAX), TABDATACHN(TABDATAMAX) AS ChainPTRS
          END IF
          ELSE
          IF FUN$ = "F9" THEN
             IF (TNAME$ = "UCTRAN  ") OR (TNAME$ = "SOUNDEX ") OR (TNAME$ = "EVENTS  ") OR (TNAME$ = "SORTSEQ ") THEN ENV.EDMAST = XEM
             IF TNCHG THEN GOTO TSINIT ELSE GOTO 40
          END IF
       END IF
    END IF
400 ' UPDATE TABLE Entry
    TABLE(TN).FF = TABDATAFREE
    CALL TableUPDT(FUN$, TABLE(TN), TABDATA$(), TABDATACHN())
    TABDATAFREE = TABLE(TN).FF
    MID$(MTDATA$(TN), 13, 3) = SFILL$(TABLE(TN).SIZE, 3)
    IF (A <> 27) OR (FUN$ = "MOV") THEN
       MID$(MTDATA$(TN), 12, 1) = "+"
       TDFIndex(TN).CHG = TDFIndex(TN).CHG OR 128
       TDFHdr.VER = CHR$(255)
       TDFIndex(TN).FIRST = 0
       IF TN = 1 THEN TNCHG = -1
    END IF
    TDFIndex(TN).ENTCNT = TABLE(TN).SIZE
    IF TABLE(TN).SIZE = 0 THEN GOTO 40
    GOTO 350

500 ' Process FILE Requests
    CONFIG.TDF = ENV.TDF
    CALL PutMSG("XXXXM005")
    IF A = 27 THEN GOTO 200
    IF (A$ < "0") OR (A$ > "3") THEN SOUND BP, DUR: GOTO 500
    R$ = A$: GOSUB 510: IF A = 27 THEN 500
    ZCHG = 0
    SELECT CASE R$
       CASE "0": ZCHG = -1: GOTO 520
       CASE "1": GOTO 520
       CASE "2": GOTO 530
       CASE "3": GOTO 540
    END SELECT

510 ' Get TDF File Prefix
    CALL FmtFIND("TABLS050")
    CALL FmtFindFLD("TDF")
    LSET UTXT$(FFLD.UTXT) = CONFIG.TDF
    CALL MsgFORMAT
    CALL FmtFindFLD("TDF"): CONFIG.TDF = UTXT$(FFLD.UTXT)
    IF CONFIG.TDF <> ENV.TDF THEN CONFIG.VER = -1
    RETURN

520 ' Save TABLE Changes Into TDF File
    CALL TDFOpen("NEW"): IF NOT OKAY THEN GOTO 500
    CALL PutMSG("TABLM006")
    TDFHdr.FTYPE = "TDF": TDFHdr.VER = CHR$(0)
    TDFHdr.TabCNT = 0: TDFHdr.ENTCNT = 0: TDFHdr.EOFRBA = LEN(TDFHdr)
    PUT #1, 1, TDFHdr
    FOR X = 1 TO TABMAX
        IF TDFIndex(X).ENTCNT > 0 THEN
           TDFDir = TDFIndex(X)
           TV = TDFDir.CHG: IF (TV AND 128) > 0 THEN TV = TV + 1
           TV = TV AND 127
           IF ZCHG THEN TV = 0
           TDFIndex(X).CHG = TV
           TDFDir.CHG = TV
           PUT #1, , TDFDir
           TDFHdr.ENTCNT = TDFHdr.ENTCNT + TDFDir.ENTCNT
           TDFHdr.TabCNT = TDFHdr.TabCNT + 1
        END IF
    NEXT X
    FOR X = 1 TO TABMAX
       TDFIndex(X).FIRST = 0
       IF TDFIndex(X).ENTCNT > 0 THEN
          TDFIndex(X).FIRST = SEEK(1)
          Z = TABLE(X).FT
          WHILE Z > 0
             L = LEN(RTRIM$(TABDATA$(Z))): L$ = CHR$(L)
             TXT$ = MID$(TABDATA$(Z), 1, L)
             PUT #1, , L$
             PUT #1, , TXT$
             Z = TABDATACHN(Z).FWD
          WEND
       END IF
    NEXT X
    TDFHdr.EOFRBA = SEEK(1) - 1
    IF TDFHdr.ENTCNT > 0 THEN
       PUT #1, 1, TDFHdr
       Y = 1
       FOR X = 1 TO TDFHdr.TabCNT
          O = SEEK(1)
          WHILE TDFIndex(Y).ENTCNT = 0: Y = Y + 1: WEND
          GET #1, , TDFDir
          TDFDir.FIRST = TDFIndex(Y).FIRST
          PUT #1, O, TDFDir
          Y = Y + 1
       NEXT X
    END IF
    CLOSE #1
    GOTO TSINIT

530 ' Load TABLE Data from TDF File
    IF ASC(TDFHdr.VER) = 255 THEN
       CALL PutMSG("XXXXM020")
       IF A$ <> SNGLKEY$(1) THEN GOTO 500
    END IF
    CALL TableLOAD("ALL", T, 0): IF NOT OKAY THEN GOTO 500
    GOTO TSINIT0

540 ' Merge FILE TABLES with TABLE()
    ' Only Merge Existing Entries of Existing Tables
    CALL TDFOpen("READ"): IF NOT OKAY THEN GOTO 500
    CALL PutMSG("TABLM009")
    FOR TN = 1 TO MASTTABLE.SIZE: TDFIndex(TN).FIRST = 0: NEXT TN
    O = LEN(TDFHdr) + 1: C$ = " "
    FOR DNUM = 1 TO TDFHdr.TabCNT
       GET #1, O, TDFDir: O = O + LEN(TDFDir)
       X = 1: FOUND = 0
       WHILE (X < TABMAX + 1) AND (NOT FOUND)
          FOUND = (TDFDir.NAME = TABLE(X).TABNAME)
          IF NOT FOUND THEN X = X + 1
       WEND
       IF FOUND THEN
          KO = TABLE(X).KO: KL = TABLE(X).KL
          IF (TDFDir.KL = KL) AND (TDFDir.KO = KO) THEN
             IF KO = 0 THEN KO = 1
             IF KL = 0 THEN KL = 8
             SEEK #1, TDFDir.FIRST
             FOR Y = 1 TO TDFDir.ENTCNT
                GET #1, , C$
                TXT$ = SPACE$(ASC(C$))
                GET #1, , TXT$
                KEY$ = MID$(TXT$, KO, KL)
                Z = TABLE(X).FT: FOUND = 0
                WHILE (Z > 0) AND (NOT FOUND)
                   FOUND = (KEY$ = MID$(TABDATA$(Z), KO, KL))
                   IF NOT FOUND THEN Z = TABDATACHN(Z).FWD
                WEND
                IF FOUND THEN
                   TABDATA$(Z) = TXT$
                END IF
             NEXT Y
          END IF
       END IF
    NEXT DNUM
    CLOSE : TDFHdr.VER = CHR$(255)
    GOTO TSINIT

600 ' Process PRINT Request
    LSET SYSVAR.MVAR = TNAME$
610 CALL PutMSG("TABLM004"): IF A = 27 THEN 200
    IF (A$ < "1") OR (A$ > "2") THEN SOUND BP, DUR: GOTO 610
    IF A$ = "1" THEN BT = TN: ET = TN:  ELSE BT = 1: ET = TABMAX
    CALL RPTOpen("TABL"): IF NOT OKAY THEN RETURN
    CALL PrintOPEN
    HTFLMAX = HTFLMAX + 1
    REDIM PRESERVE HTFLINE$(HTFLMAX)
620 ' Print AGAIN
    FOR TN = BT TO ET
        GOSUB 650
        IF A = 27 THEN TN = ET
    NEXT TN
    CALL RptCLOSE: IF PRT.AGAIN THEN 620
    GOTO TSINIT

650 ' Print TABLE #TN
    LSET TNAME$ = TABLE(TN).TABNAME
    SYSVAR.TABNAME = TNAME$
    SYSVAR.MVAR = TNAME$
    SYSVAR.FILENAME = CONFIG.TDF + "TABLES.TDF"
    'HTFMAX = 6: REDIM HTFTXT$(HTFMAX)
    'HTFTXT$(1) = "H01<Family History System>"
    'HTFTXT$(2) = "H01<[RPTNAME]>"
    'HTFTXT$(3) = "H01"
    'HTFTXT$(4) = "H02TABLE: [TABNAME]<>Table Definition File: [FILENAME]"
    'HTFTXT$(5) = "F02<Page [PAGE]>"
    TN$ = RIGHT$("0" + LTRIM$(STR$(TN)), 2)
    LSET FMTNAME$ = "TABLS1" + TN$: CALL FmtFIND(FMTNAME$)
    Z = FMT.BGNFLD: LSET PLINE$ = " ": FCOL = 0: FLTH = 0
    WHILE Z > 0
       FFLD = FLDTAB(Z)
       FCOL = FFLD.LLCC AND 255
       FLTH = FFLD.FLDL \ 256
       IF FFLD.UTXT > 0 THEN
          MID$(PLINE$, FCOL, FLTH) = UTXT$(FFLD.UTXT)
          ELSE
          MID$(PLINE$, FCOL, FLTH) = DTXT$(Z)
       END IF
       Z = FLDCHN(Z).FWD
    WEND
    HTFLINE$(HTFLMAX) = "T00" + LEFT$(PLINE$, FCOL + FLTH)
    CALL PrintTITLE
    ' Process Table Entries
    LSET FMTNAME$ = "TABLS2" + TN$: CALL FmtFIND(FMTNAME$)
    LSET PLINE$ = " "
    TLN = TABLE(TN).FT
    TE$ = SPACE$(TABLE(TN).ML)
    WHILE TLN > 0
       LSET TE$ = TABDATA$(TLN)
       Z = FMT.BGNFLD: O = 1
       WHILE Z > 0
          FFLD = FLDTAB(Z)
          FC = FFLD.CATR \ 256
          FL = FLDTAB(Z).FLDL \ 255
          FCOL = FLDTAB(Z).LLCC AND 255
          IF FC > 0 THEN MID$(PLINE$, FCOL, FL) = MID$(TE$, O, FL)
          O = O + FL
          Z = FLDCHN(Z).FWD
       WEND
       CALL PrintLINE
       TLN = TABDATACHN(TLN).FWD
    WEND
    CALL PageBREAK
    RETURN

1000 ' Return to FHSCUST.BAS
     IF ASC(TDFHdr.VER) = 255 THEN
        CALL PutMSG("XXXXM020")
        IF A$ <> SNGLKEY$(1) THEN GOTO 200
     END IF
     CALL PgmPREP(PGMRTRN$)
     IF NOT OKAY THEN GOTO 200
     CALL PutMSG("")
     CLOSE : ENV.EDMAST = 0
     CALL TableFREE("ALL")
1010 ' ReINIT Tables
     CALL TabINIT
     CHAIN PGMRTRN$

