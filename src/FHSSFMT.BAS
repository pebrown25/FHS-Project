DECLARE SUB ScrnBDRY (WNUM AS INTEGER)
DECLARE SUB xFmtGrpLOAD (GNAME$, GNUM%)
DECLARE SUB FileOPT (R$)
DECLARE SUB SFMTRpt (OPT$)
1  REM $INCLUDE: 'FHSCOMON.BAS'
   PN$ = " FAMILY HISTORY SYSTEM - FHSSFMT - Screen Format Customization program"
   CY$ = " (C) Copyright 1995, 1996 by Phillip E. Brown"
   IF CONFIG.BP = 0 THEN RUN "FHSINIT"
   FUN = 0: CNT$ = "   ": GS$ = " ": L$ = " ": XX$ = "  ": XXX$ = "   ": R$ = " "
   REDIM SLINE$(24)
   SB$ = WB$
   FMTNAME$ = SPACE$(8): GNAME$ = SPACE$(4): PGMRTRN$ = "FHSCUST"
   CALL TableLOAD("GROUPS", TGRP, 0)
   Y = TABLE(TGRP).FT
   WHILE Y > 0
      X = 0: FOUND = 0
      WHILE (NOT FOUND) AND (X < GRPMAX)
         X = X + 1
         FOUND = (SDFIndex(X).NAME = LEFT$(TABDATA$(Y), 4))
      WEND
      IF FOUND THEN SDFIndex(X).DESC = Y:  ELSE SDFIndex(X).DESC = 0
      Y = TABDATACHN(Y).FWD
   WEND
   BGNSDF$ = ENV.SDF
10 ' Prepare TABLES
   CALL FmtFIND("SFMTS035"): FMT035 = CurFMT
   CALL FmtFIND("SFMTS036"): FMT036 = CurFMT
   GrpTABLE = TABLE(TGRP)
   GrpTABLE.FMTNAME = "SFMTS015": GrpTABLE.FMTNUM = 0
   GrpTABLE.KEYNAME = "KEY"
   GrpTABLE.SIZE = GRPMAX
   REDIM GRPDATA$(GRPMAX), GRPDATACHN(GRPMAX) AS ChainPTRS
   CALL ChnINIT(GRPDATACHN(), 1, GRPMAX)
   X = 0: X$ = SPACE$(255)
   WHILE X < GRPMAX
      X = X + 1: Y = SDFIndex(X).DESC
      LSET X$ = SDFIndex(X).NAME
      IF Y > 0 THEN MID$(X$, 12, 255) = MID$(TABDATA$(Y), 5, 255)
      RSET CNT$ = LTRIM$(STR$(SDFIndex(X).CHG AND 127))
      MID$(X$, 5, 3) = CNT$
      IF (SDFIndex(X).CHG AND 128) > 0 THEN
         MID$(X$, 8, 1) = "+"
      ELSE
         MID$(X$, 8, 1) = " "
      END IF
      RSET CNT$ = LTRIM$(STR$(SDFIndex(X).SIZE))
      MID$(X$, 9, 3) = CNT$
      GRPDATA$(X) = RTRIM$(X$)
   WEND
   GrpTABLE.FT = 1: GrpTABLE.LT = GrpTABLE.SIZE
   GRPDATACHN(1).BWD = -TGRP
20 ' Open Message Group Table
   GrpTABLE.FMTNUM = 0
   CALL TableOPEN(GrpTABLE)

40 ' Format Display for Format Group List
   IF FUN = 0 THEN CALL WinCLR
   CALL FmtFIND("SFMTS010"): FMT010 = CurFMT
   CALL FmtFindFLD("SDF")
   IF FOUND THEN UTXT$(FFLD.UTXT) = CONFIG.SDF
   CALL FmtFindFLD("FMTS")
   IF FOUND THEN UTXT$(FFLD.UTXT) = LTRIM$(STR$(SDFHdr.FmtCNT))
   CALL FmtFindFLD("FLDS")
   IF FOUND THEN UTXT$(FFLD.UTXT) = LTRIM$(STR$(SDFHdr.FldCNT))
   CALL FmtFindFLD("EOF")
   IF FOUND THEN UTXT$(FFLD.UTXT) = LTRIM$(STR$(SDFHdr.EOFRBA))
   CALL WinFORMAT(0)
   CALL TableSHOW(GrpTABLE, GRPDATA$(), GRPDATACHN())
   CALL FmtFIND("XXXXS091")

200 ' Get Processing Option
    LSET PGMRTRN$ = "FHSCUST"
    CALL PutMSG(" ")
    CALL OptDisplay(0)
250 CALL TableSELECT(GrpTABLE, GRPDATA$(), GRPDATACHN())
    SELECT CASE A
       CASE 13: FUN$ = "SEL"
       CASE 59: FUN$ = "F1"
       CASE 64: FUN$ = "F6"
       CASE 66: FUN$ = "F8"
       CASE 67: FUN$ = "F9"
       CASE 94: ENV.EDMAST = NOT ENV.EDMAST: GOTO 250
       CASE 99: FUN$ = "F6": LSET PGMRTRN$ = "FHSPRTC"
       CASE 109: FUN$ = "F6": LSET PGMRTRN$ = "FHSRPTS": CALL RptLOAD("SFMT")
       CASE ELSE: SOUND BP, DUR: GOTO 250
    END SELECT
    CALL OPTHilite(FUN$, "", 6)
    SELECT CASE FUN$
       CASE "F1": CALL FileOPT(R$): IF R$ = "" THEN 200 ELSE GOTO 10
       CASE "F6": IF PGMRTRN$ <> "FHSCUST" THEN GOTO 1000
                  CALL SFMTRpt("MAIN")
                  IF PRT.OPT <> 2 THEN GOTO 40
                  GOTO 200
       CASE "F8": CALL DOSShell: GOTO 40
       CASE "F9": GOTO 1000
    END SELECT

300 ' List Formats in GROUP
    LSET GS$ = MID$(GRPDATA$(GrpTABLE.CT), 8, 1)
    LSET GNAME$ = GRPDATA$(GrpTABLE.CT)
    CALL FmtGrpLOAD(GNAME$, GNUM)
    IF NOT OKAY THEN
       SYSVAR.MVAR = GNAME$
       CALL PutMSG("SFMTM001")
       MID$(GRPDATA$(GrpTABLE.CT), 8, 1) = "?"
       GOTO 250
    END IF
    IF (FUN$ = "SEL") AND (SDFIndex(GNUM).SIZE = 0) AND (NOT ENV.EDMAST) THEN
       SOUND BP, DUR: GOTO 200
    END IF
    FmtCNT = SDFIndex(GNUM).SIZE
    FT = SDFIndex(GNUM).FIRST: LT = SDFIndex(GNUM).LAST
    DIM FMTTable AS TablePARMS
    REDIM FMTDATA$(FmtCNT), FMTDATACHN(FmtCNT) AS ChainPTRS, SFX(FmtCNT)
    FMTTable.TABNAME = "FORMATS"
    FMTTable.OPEN = SDFIndex(GNUM).OPEN
    FMTTable.SIZE = FmtCNT
    FMTTable.XL = 13: FMTTable.KO = 1: FMTTable.KL = 4
    FMTTable.FMTNAME = "SFMTS025": FMTTable.FMTNUM = 0
    FMTTable.KEYNAME = "KEY"
    FMTTable.XL = 0
    FMTTable.FT = -(FmtCNT > 0)
    FMTTable.LT = FmtCNT
    FMTTable.FF = 0
    F$ = SPACE$(FMTTable.XL)
    CALL ChnINIT(FMTDATACHN(), 1, FmtCNT)
    X = 0: Y = SDFIndex(GNUM).FIRST: F$ = SPACE$(14)
    WHILE Y > 0
       X = X + 1: SFX(X) = Y: Fmt = FMTTAB(Y)
       LSET F$ = MID$(Fmt.NAME, 5, 4)
       RSET XX$ = LTRIM$(STR$(Fmt.WIN)): MID$(F$, 5, 2) = XX$
       RSET L$ = LTRIM$(STR$(Fmt.COLOR)): MID$(F$, 7, 1) = L$
       RSET XX$ = LTRIM$(STR$(Fmt.TYPE)): MID$(F$, 8, 2) = XX$
       RSET XX$ = LTRIM$(STR$(Fmt.FLDCT)): MID$(F$, 10, 2) = XX$
       RSET XX$ = LTRIM$(STR$(Fmt.UFLDCT)): MID$(F$, 12, 2) = XX$
       FMTDATA$(X) = RTRIM$(F$)
       Y = FmtCHN(Y).FWD
    WEND
    CALL TableOPEN(FMTTable)
325 ' Format Screen for Group Format List
    X = WIN.ULR: WIN.ULR = GrpTABLE.TL - 1: CALL WinCLR: WIN.ULR = X
    CALL FmtFIND("SFMTS020")
    CALL FmtFindFLD("SDF")
    IF FOUND THEN UTXT$(FFLD.UTXT) = CONFIG.SDF
    CALL FmtFindFLD("GRP"): IF FOUND THEN UTXT$(FFLD.UTXT) = GNAME$
    CALL WinFORMAT(0)
    CALL TableSHOW(FMTTable, FMTDATA$(), FMTDATACHN())
340 ' SET Option Format
    IF ENV.EDMAST THEN LSET FMTNAME$ = "XXXXS093" ELSE LSET FMTNAME$ = "XXXXS094"
    CALL FmtFIND(FMTNAME$)
350 ' SELECT Option For Update
    CALL PutMSG("")
    CALL OptDisplay(0)
    IF FMTTable.SIZE = 0 THEN A = 255: GOTO 375
    CALL TableSELECT(FMTTable, FMTDATA$(), FMTDATACHN())
    IF (A = 27) OR ((LEN(A$) = 2) AND (A = 67)) THEN GOTO 40
    FUN$ = ""
375 ' Determine Response Option
    SELECT CASE A
       CASE 13: FUN$ = "SEL"
       CASE 27: FUN$ = "F9"
       CASE 67: FUN$ = "F9"
       CASE 59: IF ENV.EDMAST THEN FUN$ = "MOD"
       CASE 60: IF ENV.EDMAST THEN FUN$ = "ADD"
       CASE 61: IF ENV.EDMAST THEN FUN$ = "DEL"
       CASE 62: IF ENV.EDMAST THEN FUN$ = "MOV"
       CASE 64: FUN$ = "F6"
       CASE 94: ENV.EDMAST = NOT ENV.EDMAST: GOTO 340
       CASE 255: FUN$ = "ADD"
    END SELECT
    IF FUN$ = "" THEN SOUND BP, DUR: GOTO 350
    CALL OPTHilite(FUN$, "", 6)
    IF FUN$ = "F6" THEN     ' Process PRINT Option
       CALL SFMTRpt("GROUP")
       IF PRT.OPT <> 2 THEN CALL WinFORMAT(FMT010): GOTO 325
       GOTO 350
    END IF
    IF FUN$ = "F9" THEN
       IF (GNAME$ <> "XXXX") AND (GNAME$ <> "SFMT") THEN CALL FmtGrpFREE(GNAME$)
       GOTO 40
    END IF
    IF FUN$ = "SEL" THEN GOTO 400
    IF (FUN$ = "ADD") THEN
       IF FMTTable.FF = 0 THEN
          FMTTable.FF = FMTTable.SIZE + 1
          REDIM PRESERVE FMTDATA$(FMTTable.FF), FMTDATACHN(FMTTable.FF) AS ChainPTRS, SFX(FMTTable.FF)
       END IF
    END IF
    IF (FUN$ = "DEL") THEN
       CALL PutMSG("SFMTM006")
       IF A$ <> SNGLKEY$(1) THEN GOTO 350
       DT = FMTTable.CT
    END IF
    IF (FUN$ = "MOV") THEN MT = FMTTable.CT
    ' UPDATE TABLE Entry
    CALL TableUPDT(FUN$, FMTTable, FMTDATA$(), FMTDATACHN())
    IF (A <> 27) OR ((FUN$ = "MOV") AND (FMTTable.CT <> MT)) THEN
       MID$(GRPDATA$(GrpTABLE.CT), 8, 1) = "+"
       SDFIndex(GNUM).CHG = SDFIndex(GNUM).CHG OR 128
       ELSE
       IF FMTTable.SIZE = 0 THEN GOTO 40 ELSE GOTO 325
    END IF
    IF FUN$ = "DEL" THEN
       Y = SFX(DT)
       CALL FmtRELEASE(Y)
       GOTO 380
    END IF
    IF FUN$ = "MOV" THEN
       X = FMTTable.FT
       WHILE X > 0
          Y = SFX(X)
          X1 = FMTDATACHN(X).BWD: X2 = FMTDATACHN(X).FWD
          IF X1 > 0 THEN FmtCHN(Y).BWD = SFX(X1):  ELSE FmtCHN(Y).BWD = 0
          IF X2 > 0 THEN FmtCHN(Y).FWD = SFX(X2):  ELSE FmtCHN(Y).FWD = 0
          X = X2
       WEND
       GOTO 380
    END IF
    IF FUN$ = "ADD" THEN
       IF FMTFree = 0 THEN
          FMTMAX = FMTMAX + 1
          FMTFree = FMTMAX
          REDIM PRESERVE FMTTAB(FMTMAX) AS FmtENTRY, FmtCHN(FMTMAX) AS ChainPTRS
       END IF
       Y = FMTFree: FMTFree = FmtCHN(FMTFree).FWD
       IF FMTFree > 0 THEN FmtCHN(FMTFree).BWD = 0
       X = FMTTable.CT: X1 = FMTDATACHN(X).BWD: X2 = FMTDATACHN(X).FWD
       SFX(X) = Y
       IF X1 > 0 THEN Y1 = SFX(X1) ELSE Y1 = 0
       IF X2 > 0 THEN Y2 = SFX(X2) ELSE Y2 = 0
       FmtCHN(Y).BWD = Y1: FmtCHN(Y).FWD = Y2
       IF Y1 > 0 THEN FmtCHN(Y1).FWD = Y
       IF Y2 > 0 THEN FmtCHN(Y2).BWD = Y
       Fmt.FLDCT = 0: Fmt.BGNFLD = 0: Fmt.UFLDCT = 0
       ELSE
       Fmt = FMTTAB(SFX(FMTTable.CT))
    END IF
    LSET F$ = FMTDATA$(FMTTable.CT)
    IF MID$(F$, 1, 4) = "    " THEN MID$(F$, 1, 4) = "...."
    Fmt.NAME = GNAME$ + MID$(F$, 1, 4)
    Fmt.WIN = VAL(MID$(F$, 5, 2))
    Fmt.COLOR = VAL(MID$(F$, 7, 1))
    Fmt.TYPE = VAL(MID$(F$, 8, 2))
    FMTTAB(SFX(FMTTable.CT)) = Fmt
    RSET XX$ = LTRIM$(STR$(Fmt.WIN)): MID$(F$, 5, 2) = XX$
    RSET L$ = LTRIM$(STR$(Fmt.COLOR)): MID$(F$, 7, 1) = L$
    RSET XX$ = LTRIM$(STR$(Fmt.TYPE)): MID$(F$, 8, 2) = XX$
    RSET XX$ = LTRIM$(STR$(Fmt.FLDCT)): MID$(F$, 10, 2) = XX$
    RSET XX$ = LTRIM$(STR$(Fmt.UFLDCT)): MID$(F$, 12, 2) = XX$
    FMTDATA$(FMTTable.CT) = RTRIM$(F$)
380 ' Update SDFIndex()
    RSET CNT$ = LTRIM$(STR$(FMTTable.SIZE))
    MID$(GRPDATA$(GrpTABLE.CT), 9, 3) = CNT$
    SDFIndex(GNUM).OPEN = (FMTTable.SIZE > 0)
    SDFIndex(GNUM).SIZE = FMTTable.SIZE
    IF FMTTable.SIZE = 0 THEN
       SDFIndex(GNUM).FIRST = 0: SDFIndex(GNUM).LAST = 0
       GOTO 40
    ELSE
       SDFIndex(GNUM).FIRST = SFX(FMTTable.FT)
       SDFIndex(GNUM).LAST = SFX(FMTTable.LT)
    END IF
    GOTO 325

400 ' Display Format Field List
    Z = SFX(FMTTable.CT)
    IF (FUN$ = "SEL") AND (FMTTAB(Z).FLDCT = 0) AND (NOT ENV.EDMAST) THEN
       SOUND BP, DUR: GOTO 325
    END IF
    Fmt = FMTTAB(Z): LSET FMTNAME$ = Fmt.NAME
    FldCNT = Fmt.FLDCT
    FF = Fmt.BGNFLD
    DIM FLDTable AS TablePARMS
    REDIM FLDDATA$(FldCNT), FLDDATACHN(FldCNT) AS ChainPTRS, FTX(FldCNT)
    FLDTable.TABNAME = FMTNAME$
    FLDTable.OPEN = 0
    FLDTable.SIZE = FldCNT
    FLDTable.XL = 95: FLDTable.KO = 1: FLDTable.KL = 4
    FLDTable.FMTNAME = "SFMTS035": FLDTable.FMTNUM = 0
    FLDTable.KEYNAME = "KEY"
    FLDTable.FT = -(FldCNT > 0)
    FLDTable.LT = FldCNT
    FLDTable.FF = 0
    FF$ = SPACE$(FLDTable.XL)
    CALL ChnINIT(FLDDATACHN(), 1, FldCNT)
    X = 0: Y = Fmt.BGNFLD
    WHILE Y > 0
       X = X + 1: FTX(X) = Y: FFLD = FLDTAB(Y)
       IF FFLD.NAME <> "    " THEN LSET FF$ = FFLD.NAME ELSE LSET FF$ = "...."
       RSET L$ = LTRIM$(STR$(FFLD.CATR \ 256)): MID$(FF$, 5, 1) = L$
       RSET XX$ = LTRIM$(STR$(FFLD.CATR AND 255)): MID$(FF$, 6, 2) = XX$
       RSET XX$ = LTRIM$(STR$(FFLD.LLCC \ 256)): MID$(FF$, 8, 2) = XX$
       RSET XX$ = LTRIM$(STR$(FFLD.LLCC AND 255)): MID$(FF$, 10, 2) = XX$
       RSET XXX$ = LTRIM$(STR$(FFLD.FLDL AND 255)): MID$(FF$, 12, 3) = XXX$
       RSET XXX$ = LTRIM$(STR$(FFLD.FLDL \ 256)): MID$(FF$, 15, 3) = XXX$
       MID$(FF$, 18, 255) = DTXT$(Y)
       FLDDATA$(X) = RTRIM$(FF$)
       Y = FLDCHN(Y).FWD
    WEND
    CALL TableOPEN(FLDTable)
    X = WIN.ULR: WIN.ULR = 5: CALL WinCLR: WIN.ULR = X
420 ' Format Field List Display
    CALL FmtFIND("SFMTS030"): FMT030 = CurFMT
    CALL FmtFindFLD("SDF")
    IF FOUND THEN UTXT$(FFLD.UTXT) = CONFIG.SDF
    CALL FmtFindFLD("FMT"): IF FOUND THEN UTXT$(FFLD.UTXT) = FMTNAME$
425 ' Format FIELD Display
    CALL WinFORMAT(FMT030)
    CALL TableSHOW(FLDTable, FLDDATA$(), FLDDATACHN())
440 ' SET Option Format
    IF ENV.EDMAST THEN LSET FMTNAME$ = "XXXXS093" ELSE LSET FMTNAME$ = "SFMTS094"
    CALL FmtFIND(FMTNAME$)
    IF NOT FOUND THEN SOUND BP, DUR: GOTO 40
450 ' SELECT Option For Update
    CALL OptDisplay(0)
    IF FLDTable.SIZE = 0 THEN A = 255: GOTO 475
    CALL TableSELECT(FLDTable, FLDDATA$(), FLDDATACHN())
    IF (A = 27) OR ((LEN(A$) = 2) AND (A = 67)) THEN GOTO 325
    FUN$ = "": XCHG = 0
475 ' Determine Response Option
    SELECT CASE A
       CASE 28: FUN$ = "CHG": XCHG = -1
       CASE 13: FUN$ = "CHG"
       CASE 27: FUN$ = "F9"
       CASE 67: FUN$ = "F9"
       CASE 59: IF ENV.EDMAST THEN FUN$ = "MOD"
       CASE 60: IF ENV.EDMAST THEN FUN$ = "ADD"
       CASE 61: IF ENV.EDMAST THEN FUN$ = "DEL"
       CASE 62: IF ENV.EDMAST THEN FUN$ = "MOV"
       CASE 64: FUN$ = "F6"
       CASE 65: FUN$ = "F7"
       CASE 94: ENV.EDMAST = NOT ENV.EDMAST: GOTO 440
       CASE 255: FUN$ = "ADD"
    END SELECT
    IF FUN$ = "" THEN SOUND BP, DUR: GOTO 450
    CALL OPTHilite(FUN$, "", 6)
    IF FUN$ = "F9" THEN GOTO 380
    IF FUN$ = "F7" THEN GOTO 490
    IF FUN$ = "F6" THEN     ' Process PRINT Option
       CALL SFMTRpt("FORMAT")
       IF PRT.OPT <> 2 THEN CALL WinFORMAT(FMT010): GOTO 425
       GOTO 450
    END IF
    IF (FUN$ = "ADD") THEN
       IF FLDTable.FF = 0 THEN
          FLDTable.FF = FLDTable.SIZE + 1
          REDIM PRESERVE FLDDATA$(FLDTable.FF), FLDDATACHN(FLDTable.FF) AS ChainPTRS, FTX(FLDTable.FF)
       END IF
    END IF
    IF (FUN$ = "DEL") THEN
       CALL PutMSG("SFMTM007")
       IF A$ <> SNGLKEY$(1) THEN GOTO 450
       DT = FLDTable.CT
    END IF
    IF (FUN$ = "MOV") THEN MT = FLDTable.CT
    ' UPDATE TABLE Entry
    IF XCHG THEN FLDTable.FMTNUM = FMT036 ELSE FLDTable.FMTNUM = FMT035
    CALL TableUPDT(FUN$, FLDTable, FLDDATA$(), FLDDATACHN())
    IF (A <> 27) OR ((FUN$ = "MOV") AND (FMTTable.CT <> MT)) THEN
       MID$(GRPDATA$(GrpTABLE.CT), 8, 1) = "+"
       SDFIndex(GNUM).CHG = SDFIndex(GNUM).CHG OR 128
       ELSE
       IF FLDTable.SIZE = 0 THEN GOTO 325 ELSE GOTO 450
    END IF
    FFLD = FLDTAB(FTX(FLDTable.CT))
    U = (FFLD.UTXT > 0)
    IF FUN$ = "DEL" THEN
       Y = FTX(DT)
       CALL FldRELEASE(Y)
       GOTO 480
    END IF
    IF FUN$ = "MOV" THEN
       X = FLDTable.FT
       WHILE X > 0
          Y = FTX(X)
          X1 = FLDDATACHN(X).BWD: X2 = FLDDATACHN(X).FWD
          IF X1 > 0 THEN FLDCHN(Y).BWD = FTX(X1):  ELSE FLDCHN(Y).BWD = 0
          IF X2 > 0 THEN FLDCHN(Y).FWD = FTX(X2):  ELSE FLDCHN(Y).FWD = 0
          X = X2
       WEND
       U = 0: GOTO 480
    END IF
    IF FUN$ = "ADD" THEN
       IF FLDFree = 0 THEN
          FLDFree = FLDMAX + 1
          FLDMAX = FLDMAX + 1
          REDIM PRESERVE FLDTAB(FLDMAX) AS FldENTRY, FLDCHN(FLDMAX) AS ChainPTRS
          REDIM PRESERVE DTXT$(FLDMAX)
       END IF
       Y = FLDFree: FLDFree = FLDCHN(FLDFree).FWD
       IF FLDFree > 0 THEN FLDCHN(FLDFree).BWD = 0
       X = FLDTable.CT: X1 = FLDDATACHN(X).BWD: X2 = FLDDATACHN(X).FWD
       FTX(X) = Y
       IF X1 > 0 THEN Y1 = FTX(X1) ELSE Y1 = 0
       IF X2 > 0 THEN Y2 = FTX(X2) ELSE Y2 = 0
       FLDCHN(Y).BWD = Y1: FLDCHN(Y).FWD = Y2
       IF Y1 > 0 THEN FLDCHN(Y1).FWD = Y
       IF Y2 > 0 THEN FLDCHN(Y2).BWD = Y
       FFLD.CATR = 0: FFLD.FLDL = 0: FFLD.LLCC = 0
       FFLD.DO = 0: FFLD.UTXT = 0: U = 0
    END IF
    LSET FF$ = FLDDATA$(FLDTable.CT): FX = FTX(FLDTable.CT)
    FFLD.NAME = MID$(FF$, 1, 4): IF FFLD.NAME = "...." THEN FFLD.NAME = "    "
    C = VAL(MID$(FF$, 5, 1)): ATR = VAL(MID$(FF$, 6, 2))
    FFLD.CATR = 256 * C + ATR
    LL = VAL(MID$(FF$, 8, 2)): CC = VAL(MID$(FF$, 10, 2))
    FFLD.LLCC = 256 * LL + CC
    DL = VAL(MID$(FF$, 12, 3)): FL = VAL(MID$(FF$, 15, 3))
    XL = LEN(RTRIM$(MID$(FF$, 18, 255)))
    IF (C <> 0) AND ((ATR AND 19) = 0) AND ((FMTTAB(Z).TYPE AND 16) = 0) THEN
       FL = XL: DL = FL
    ELSE
       IF FL < XL THEN FL = XL
       IF FL < DL THEN FL = DL
       IF DL = 0 THEN DL = FL
    END IF
    FFLD.FLDL = 256 * FL + DL
    IF FFLD.NAME = "    " THEN
       MID$(FF$, 1, 4) = "...."
    ELSE
       MID$(FF$, 1, 4) = FFLD.NAME
    END IF
    DTXT$(FX) = RTRIM$(MID$(FF$, 18, FL))
    RSET L$ = LTRIM$(STR$(FFLD.CATR \ 256)): MID$(FF$, 5, 1) = L$
    RSET XX$ = LTRIM$(STR$(FFLD.CATR AND 255)): MID$(FF$, 6, 2) = XX$
    RSET XX$ = LTRIM$(STR$(FFLD.LLCC \ 256)): MID$(FF$, 8, 2) = XX$
    RSET XX$ = LTRIM$(STR$(FFLD.LLCC AND 255)): MID$(FF$, 10, 2) = XX$
    RSET XXX$ = LTRIM$(STR$(FFLD.FLDL AND 255)): MID$(FF$, 12, 3) = XXX$
    RSET XXX$ = LTRIM$(STR$(FFLD.FLDL \ 256)): MID$(FF$, 15, 3) = XXX$
    FLDDATA$(FLDTable.CT) = RTRIM$(FF$)
    IF (FFLD.UTXT > 0) AND ((ATR AND 17) = 0) THEN
       IF UTXTFree > 0 THEN UTXTCHN(UTXTFree).BWD = FFLD.UTXT
       UTXTCHN(FFLD.UTXT).FWD = UTXTFree
       UTXTCHN(FFLD.UTXT).BWD = 0
       UTXTFree = FFLD.UTXT
       UTXT$(FFLD.UTXT) = "": FFLD.UTXT = 0
    END IF
    IF (ATR AND 17) > 0 THEN
       U = 0
       IF (FFLD.UTXT = 0) THEN
          IF UTXTFree = 0 THEN
             UTXTFree = UTXTMAX + 1
             UTXTMAX = UTXTMAX + 1
             REDIM PRESERVE UTXT$(UTXTMAX), HTXT$(UTXTMAX)
             REDIM PRESERVE UTXTCHN(UTXTMAX) AS ChainPTRS
          END IF
          FFLD.UTXT = UTXTFree: UTXTFree = UTXTCHN(FFLD.UTXT).FWD
          U = 1
       END IF
       IF FFLD.UTXT > 0 THEN
          UTXT$(FFLD.UTXT) = SPACE$(FL)
          LSET UTXT$(FFLD.UTXT) = DTXT$(FTX(FLDTable.CT))
          HTXT$(FFLD.UTXT) = UTXT$(FFLD.UTXT)
       ELSE
          FFLD.CATR = FFLD.CATR AND (255 - 17)
       END IF
    END IF
    FLDTAB(FTX(FLDTable.CT)) = FFLD
480 ' Update FMTTAB(FMTTABLE.CT)
    Fmt = FMTTAB(SFX(FMTTable.CT))
    Fmt.FLDCT = FLDTable.SIZE
    IF Fmt.FLDCT > 0 THEN Fmt.BGNFLD = FTX(FLDTable.FT) ELSE Fmt.BGNFLD = 0
    Fmt.UFLDCT = Fmt.UFLDCT + U
    FMTTAB(SFX(FMTTable.CT)) = Fmt
    LSET F$ = FMTDATA$(FMTTable.CT)
    RSET XX$ = LTRIM$(STR$(Fmt.FLDCT))
    MID$(F$, 10, 2) = XX$
    RSET XX$ = LTRIM$(STR$(Fmt.UFLDCT))
    MID$(F$, 12, 2) = XX$
    FMTDATA$(FMTTable.CT) = RTRIM$(F$)
    IF FLDTable.SIZE = 0 THEN GOTO 325
    GOTO 425

490 ' Test Display Format
    CurFMT = SFX(FMTTable.CT)
    Fmt = FMTTAB(CurFMT)
    COLOR FG(5), BG(5): LOCATE 25, 71, 0: PRINT Fmt.NAME;
    WT$ = WINDOWS(Fmt.WIN).LOC
    CALL WinOPEN(WT$)
    IF (Fmt.TYPE AND 4) > 0 THEN
       CALL OptDisplay(CurFMT)
       CALL GetRESP
       ELSE
       CALL WinFORMAT(0)
       CALL GetUFLD("FIRST")
       IF NOT FOUND THEN
          CALL GetRESP: CALL PutMSG("")
          ELSE
          CALL WinUPDATE
       END IF
    END IF
    CALL PutMSG("")
    WNUM = WIN.WNUM
    CALL WinCLOSE(0, 0)
    IF WNUM <> ENV.WINMSG THEN CALL WinErase(0, WNUM)
    CALL WinBDRY(0, 0)
    GOTO 400

1000 ' Return to Calling Program
     XCHG = 0
     FOR X = 1 TO GRPMAX
         XCHG = XCHG - ((SDFIndex(X).CHG AND 128) > 0)
     NEXT X
     IF XCHG THEN
        CALL PutMSG("XXXXM020")
        IF A$ <> SNGLKEY$(1) THEN GOTO 200
     END IF
     IF ENV.SDF <> BGNSDF$ THEN CONFIG.VER = -1
     CALL TableFREE("GROUPS")
     IF (XCHG = 0) THEN CALL FmtGrpLOAD("INITIAL", Z)
     CALL PgmPREP(PGMRTRN$)
     IF (A = 27) OR (NOT OKAY) THEN GOTO 200
     CHAIN PGMRTRN$

REM $DYNAMIC
SUB FileOPT (R$)
SHARED GRPDATA$()
    DIM XXXHDR AS FmtFileHDR
    CURSDF$ = ENV.SDF: GN$ = "    ": CNT$ = "   "
500 ' Process FILE Requests
    CALL PutMSG("XXXXM005")
    IF A = 27 THEN GOTO 595 ELSE R$ = A$
    IF (R$ < "0") OR (R$ > "3") THEN CALL ErrBEEP(0): GOTO 500
    IF (R$ <> "3") THEN
       CONFIG.SDF = CURSDF$
       MFMT$ = "SFMTS050": GOSUB 510: IF A = 27 THEN 595
       NEWSDF$ = CONFIG.SDF
    END IF
    FOR X = 1 TO GRPMAX
        XCHG = XCHG - ((SDFIndex(X).CHG AND 128) > 0)
    NEXT X
    SELECT CASE R$
       CASE "0": ZCHG = -1: GOTO 520
       CASE "1": GOTO 520
       CASE "2": GOTO 530
       CASE "3": GOTO 540
    END SELECT
    GOTO 595

510 ' Get SDF File Prefix
    CALL FmtFIND(MFMT$)
    CALL FmtFindFLD("SDF"): X = FFLD.UTXT
    LSET UTXT$(X) = CONFIG.SDF
    CALL MsgFORMAT
    IF A <> 27 THEN CONFIG.SDF = UTXT$(X)
    RETURN

520 ' Save TABLE Changes Into SDF File
    CONFIG.SDF = ENV.SDF
    GOSUB 550    ' Open Temporary File $$$SFMT$.SDF
    SYSVAR.FILENAME = NEWSDF$ + "SFMTS.SDF"
    CALL PutMSG("SFMTM010")
    GOSUB 560    ' Prepare XXXHDR, XXXIndex()
    FOR GN = 1 TO GRPMAX
       IF SDFIndex(GN).SIZE > 0 THEN
          LSET GN$ = SDFIndex(GN).NAME
          IF (SDFIndex(GN).OPEN = 0) THEN
             CALL FmtGrpLOAD(GN$, X): IF NOT OKAY THEN 595
          END IF
          GOSUB 570  ' Write Format Group (Formats, Fields) to Output File
          IF (GN$ <> "XXXX") AND (GN$ <> "SFMT") THEN
             CALL FmtGrpFREE(GN$)
          END IF
       END IF
    NEXT GN
    GOSUB 580  ' Finish XXXHDR, GrpDIR Entries in Output File
    XCHG = 0

530 ' Load Screen Formats from SDF File
     IF XCHG THEN
        CALL PutMSG("XXXXM020")
        IF A$ <> SNGLKEY$(1) THEN GOTO 500
     END IF
     CALL PutMSG("SFMTM011")
     CONFIG.SDF = NEWSDF$
     CALL FmtGrpLOAD("INITIAL", Z): IF NOT OKAY THEN GOTO 500
     GOTO 595

540 ' Merge Formats from SDF File
    IF XCHG THEN
       CALL PutMSG("XXXXM020")
       IF A$ <> SNGLKEY$(1) THEN GOTO 500 ELSE XCHG = 0
       CALL FmtGrpLOAD("INITIAL", GN): IF NOT OKAY THEN 500
    END IF
    MFMT$ = "SFMTS051": GOSUB 510: IF A = 27 THEN 595
    MRGSDF$ = CONFIG.SDF
    MFMT$ = "SFMTS052": GOSUB 510: IF A = 27 THEN 595
    NEWSDF$ = CONFIG.SDF
    XXXHDR = SDFHdr
    GOSUB 550  ' Open Temp File $$$SFMT$.SDF
    CONFIG.SDF = MRGSDF$
    CALL SDFOpen("READ"): IF NOT OKAY THEN 595
    MRGGrpCnt = SDFHdr.GrpCnt: CLOSE #1
    OPEN SYSVAR.FILENAME FOR BINARY ACCESS READ AS #3
    SDFHdr = XXXHDR: CONFIG.SDF = CURSDF$: ENV.SDF = CONFIG.SDF
    GOSUB 560   ' Prepare XXXHDR, XXXIndex()
    CALL PutMSG("SFMTM012")
    ' Only Merge Formats that are in Current Group Table
    REDIM MRGDIR(GRPMAX)
    X = 0:  OD = LEN(SDFHdr) + 1
    WHILE X < MRGGrpCnt
       X = X + 1: GET #3, OD, SDFGrpDIR
       IF SDFGrpDIR.FmtCNT <> 0 THEN
          Y = 0: FOUND = 0
          WHILE (Y < GRPMAX) AND (NOT FOUND)
             Y = Y + 1: FOUND = (SDFGrpDIR.NAME = SDFIndex(Y).NAME)
          WEND
          IF FOUND THEN
             MRGDIR(Y) = OD
             ELSE
             SYSVAR.MVAR = SDFGrpDIR.NAME
             CALL PutMSG("PRTCM007")
          END IF
       END IF
       OD = OD + LEN(SDFGrpDIR)
    WEND
    FOR GN = 1 TO GRPMAX
       IF SDFIndex(GN).SIZE > 0 THEN
          LSET GN$ = SDFIndex(GN).NAME
          IF (SDFIndex(GN).OPEN = 0) THEN
             CALL FmtGrpLOAD(GN$, GX): IF NOT OKAY THEN 595
          END IF
          IF MRGDIR(GN) <> 0 THEN GOSUB 545
          GOSUB 570  ' Write Format Group (Formats, Fields) to Output File
          IF (GN$ <> "XXXX") AND (GN$ <> "SFMT") THEN
             CALL FmtGrpFREE(GN$)
          END IF
       END IF
    NEXT GN
    GOSUB 580  ' Finish XXXHDR, GrpDIR Entries in Output File
    GOTO 530

545 ' Merge Group with Current Group
    GET #3, MRGDIR(GN), SDFGrpDIR
    U = 0: SEEK #3, SDFGrpDIR.FIRST
    WHILE (U < SDFGrpDIR.FmtCNT)
       U = U + 1
       GET #3, , SDFFmtREC
       V = SDFIndex(GN).FIRST: FMTFOUND = 0
       WHILE (V > 0) AND (NOT FMTFOUND)
          FMTFOUND = (SDFFmtREC.NAME = FMTTAB(V).NAME)
          IF NOT FMTFOUND THEN V = FmtCHN(V).FWD
       WEND
       FF = FMTTAB(V).BGNFLD: Z = SDFFmtREC.FldCNT
       WHILE (Z > 0)
          GET #3, , SDFFldREC
          ZL = ASC(SDFFldREC.LTH)
          Z$ = SPACE$(ZL)
          GET #3, , Z$
          IF FMTFOUND THEN ' LOOK FOR Field in Current FLDTAB()
             W = FF: FOUND = 0
             WHILE (W > 0) AND (NOT FOUND)
                IF FLDTAB(W).NAME = SDFFldREC.NAME THEN
                   FFLD = FLDTAB(W)
                   LL = FFLD.LLCC \ 256: CC = FFLD.LLCC AND 127
                   MLL = ASC(SDFFldREC.LL): MCC = ASC(SDFFldREC.CC)
                   FOUND = ((FFLD.NAME <> "    ") OR ((LL = MLL) AND (CC = MCC)))
                END IF
                IF NOT FOUND THEN W = FLDCHN(W).FWD
             WEND
             IF FOUND THEN
                DTXT$(W) = Z$
                FFLD.LLCC = 256 * MLL + MCC
                FL = FFLD.FLDL \ 256: DL = FFLD.FLDL AND 255
                MFL = ASC(SDFFldREC.FL): MDL = ASC(SDFFldREC.DL)
                IF (FL > DL) AND (FL > MFL) THEN MFL = FL
                FFLD.FLDL = 256 * MFL + MDL
                FLDTAB(W) = FFLD
             END IF
          END IF
          Z = Z - 1
       WEND
    WEND
    RETURN

550 ' Open Temporary File $$$SFMT$.SDF
    ON LOCAL ERROR GOTO 551
    KILL "$$$SFMT$.SDF"
    GOTO 552
551 RESUME 552
552 ON LOCAL ERROR GOTO 553
    OPEN "$$$SFMT$.SDF" FOR BINARY ACCESS READ WRITE AS #2
    GOTO 554
553 SYSVAR.FILENAME = "$$$SFMT$.SDF"
    CALL PutMSG("XXXXM030")
    CLOSE : RESUME 595
554 ' Return to Caller
    'ON LOCAL ERROR GOTO 0
    ON LOCAL ERROR GOTO 590
    RETURN

560 ' Build XXXHDR, XXXIndex() Array
    XXXHDR.FTYPE = "SDF": XXXHDR.VER = CHR$(0)
    XXXHDR.GrpCnt = 0
    XXXHDR.FmtCNT = 0: XXXHDR.FldCNT = 0: XXXHDR.UFldCNT = 0
    XXXHDR.EOFRBA = LEN(XXXHDR) + 1
    PUT #2, 1, XXXHDR
    REDIM XXXIndex(GRPMAX) AS GroupPARMS
    SDFGrpDIR.FldCNT = 0: SDFGrpDIR.UFldCNT = 0: SDFGrpDIR.FIRST = 0
    FOR GN = 1 TO GRPMAX
       XXXIndex(GN) = SDFIndex(GN)
       IF XXXIndex(GN).SIZE > 0 THEN
          SDFGrpDIR.NAME = XXXIndex(GN).NAME
          CHG = XXXIndex(GN).CHG
          IF (CHG AND 128) > 0 THEN CHG = (CHG AND 127) + 1
          IF ZCHG THEN CHG = 0 ELSE CHG = CHG AND 127
          XXXIndex(GN).CHG = CHG
          SDFGrpDIR.CHG = CHG
          XXXIndex(GN).FDIR = XXXHDR.EOFRBA
          SDFGrpDIR.FmtCNT = XXXIndex(GN).SIZE
          PUT #2, , SDFGrpDIR
          XXXHDR.EOFRBA = XXXHDR.EOFRBA + LEN(SDFGrpDIR)
       END IF
    NEXT GN
    REDIM SGFC(GRPMAX), SGUFC(GRPMAX)
    RETURN

570 ' Write Format Group to Output SDF File
    XXXIndex(GN).BGNRBA = XXXHDR.EOFRBA
    XXXHDR.GrpCnt = XXXHDR.GrpCnt + 1
    XXXHDR.FmtCNT = XXXHDR.FmtCNT + XXXIndex(GN).SIZE
    X = SDFIndex(GN).FIRST
    WHILE X > 0
       Fmt = FMTTAB(X)
       XXXHDR.FldCNT = XXXHDR.FldCNT + Fmt.FLDCT
       XXXHDR.UFldCNT = XXXHDR.UFldCNT + Fmt.UFLDCT
       SGFC(GN) = SGFC(GN) + Fmt.FLDCT
       SGUFC(GN) = SGUFC(GN) + Fmt.UFLDCT
       SDFFmtREC.NAME = Fmt.NAME
       SDFFmtREC.TYPE = Fmt.TYPE
       SDFFmtREC.WINNUM = Fmt.WIN
       SDFFmtREC.WinBGColor = Fmt.COLOR
       SDFFmtREC.FldCNT = Fmt.FLDCT
       SDFFmtREC.FIRST = XXXHDR.EOFRBA + LEN(SDFFmtREC)
       PUT #2, , SDFFmtREC
       Y = Fmt.BGNFLD
       WHILE Y > 0
          FFLD = FLDTAB(Y)
          SDFFldREC.NAME = FFLD.NAME
          SDFFldREC.COLOR = CHR$(FFLD.CATR \ 256)
          SDFFldREC.ATR = CHR$(FFLD.CATR AND 127)
          SDFFldREC.LL = CHR$(FFLD.LLCC \ 256)
          SDFFldREC.CC = CHR$(FFLD.LLCC AND 127)
          SDFFldREC.DL = CHR$(FFLD.FLDL AND 127)
          SDFFldREC.FL = CHR$(FFLD.FLDL \ 256)
          Y$ = RTRIM$(DTXT$(Y))
          SDFFldREC.LTH = CHR$(LEN(Y$))
          PUT #2, , SDFFldREC
          PUT #2, , Y$
          Y = FLDCHN(Y).FWD
       WEND
       XXXHDR.EOFRBA = SEEK(2)
       X = FmtCHN(X).FWD
    WEND
    SDFIndex(GN).CHG = SDFIndex(GN).CHG AND 127
    RETURN

580 ' Finish Ouptut SDF File HDR, GrpDIR RecordS
    XXXHDR.EOFRBA = XXXHDR.EOFRBA - 1
    PUT #2, 1, XXXHDR
    O = LEN(XXXHDR) + 1
    FOR GN = 1 TO GRPMAX
       IF XXXIndex(GN).SIZE > 0 THEN
          GET #2, O, SDFGrpDIR
          SDFGrpDIR.FldCNT = SGFC(GN)
          SDFGrpDIR.UFldCNT = SGUFC(GN)
          SDFGrpDIR.FIRST = XXXIndex(GN).BGNRBA
          PUT #2, O, SDFGrpDIR
          O = O + LEN(SDFGrpDIR)
       END IF
    NEXT GN
    CLOSE
    ON LOCAL ERROR GOTO 585
    KILL NEWSDF$ + "SFMTS.SDF"
    GOTO 586
585 RESUME 586
586 ON LOCAL ERROR GOTO 0
    NAME "$$$SFMT$.SDF" AS NEWSDF$ + "SFMTS.SDF"
    RETURN

590 ' Error Building $$$SFMT$.SDF
    XERR = ERR: XERL = ERL: SYSVAR.FILENAME = "$$$SFMT$.SDF"
    CALL ERRMessage("PRTCM008", XERR, XERL)
    OKAY = 0: CONFIG.SDF = CURSDF$
    RESUME 595

595  ' Return
     CLOSE
     EXIT SUB

END SUB

SUB ScrnBDRY (WNUM AS INTEGER)
SHARED SLINE$(), SB$
5100 ' DRAW WINDOW BOUNDARIES
     W = WNUM
     ULR = WINDOWS(W).ULR: ULC = WINDOWS(W).ULC
     LRR = WINDOWS(W).LRR: LRC = WINDOWS(W).LRC
     LL = LRR - ULR: CC = LRC - ULC + 1
     MID$(SLINE$(0), 1, 1) = MID$(SB$, 1, 1)
     MID$(SLINE$(0), CC, 1) = MID$(SB$, 3, 1)
     MID$(SLINE$(LL), 1, 1) = MID$(SB$, 6, 1)
     MID$(SLINE$(LL), CC, 1) = MID$(SB$, 8, 1)
     FOR I = 2 TO CC - 1:
         MID$(SLINE$(0), I, 1) = MID$(SB$, 2, 1)
         MID$(SLINE$(LL), I, 1) = MID$(SB$, 7, 1)
     NEXT I
     FOR I = 1 TO LL - 1
       MID$(SLINE$(I), 1, 1) = MID$(SB$, 4, 1)
       MID$(SLINE$(I), CC, 1) = MID$(SB$, 5, 1)
     NEXT I
END SUB

SUB SFMTRpt (OPT$)
SHARED GN, GNUM, SFX(), GRPDATA$()
SHARED FMTTable AS TablePARMS, FLDTable AS TablePARMS
SHARED SB$, SLINE$()
    PRT.OPT = 0: CNT$ = "   ": GN$ = "    ": GNX$ = GN$
    SELECT CASE OPT$
       CASE "MAIN": GOTO RPTMAIN
       CASE "GROUP": GOTO RPTGRP
       CASE "FORMAT": GOTO RPTFMT
       CASE "SCREEN": GOTO RPTSCRN
    END SELECT
    GOTO RPTEXIT

RPTMAIN:
    CALL PutMSG("SFMTM002")
    R$ = A$: IF (R$ < "1") OR (R$ > "5") THEN GOTO RPTEXIT
    IF (R$ = "3") OR (R$ = "5") THEN
       CALL PutMSG("SFMTM004"): IF A = 27 THEN GOTO RPTEXIT
       ShowSCREEN = (A$ = SNGLKEY$(1))
    END IF
    IF ShowSCREEN OR (R$ = "4") OR (R$ = "5") THEN
       CALL PutMSG("SFMTM005"): IF A = 27 THEN GOTO RPTEXIT
       OmitBL = (A$ = SNGLKEY$(1))
    END IF
    GOSUB 600: IF (A = 27) OR (NOT OKAY) THEN GOTO RPTEXIT
    ALLRPTS = 0: GB = 1: GE = GRPMAX
    SELECT CASE R$
       CASE "1": GOSUB 620
       CASE "2": GOSUB 630
       CASE "3": GOSUB 660 ' Detail Reports
       CASE "4": GOSUB 660 ' Screen Displays
       CASE "5": ALLRPTS = -1: GOSUB 620
    END SELECT
    GOTO RPTEXIT

RPTGRP:
    CALL PutMSG("SFMTM003")
    R$ = A$: IF (R$ < "1") OR (R$ > "3") THEN GOTO RPTEXIT
    IF (R$ = "2") THEN
       CALL PutMSG("SFMTM004"): IF A = 27 THEN GOTO RPTEXIT
       ShowSCREEN = (A$ = SNGLKEY$(1))
    END IF
    IF ShowSCREEN OR (R$ = "3") THEN
       CALL PutMSG("SFMTM005"): IF A = 27 THEN GOTO RPTEXIT
       OmitBL = (A$ = SNGLKEY$(1))
    END IF
    GOSUB 600: IF (A = 27) OR (NOT OKAY) THEN GOTO RPTEXIT
    GB = GNUM: GE = GNUM
    SELECT CASE R$
       CASE "1": GOSUB 630 ' Group List
       CASE "2": R$ = "3": GOSUB 660 ' Detail Reports
       CASE "3": R$ = "4": GOSUB 660 ' Screen Displays
    END SELECT
    GOTO RPTEXIT

RPTFMT:
    Y = SFX(FMTTable.CT): Fmt = FMTTAB(Y)
    SYSVAR.MVAR = Fmt.NAME
    CALL PutMSG("SFMTM004"): IF A = 27 THEN GOTO RPTEXIT
    ShowSCREEN = (A$ = SNGLKEY$(1))
    IF ShowSCREEN THEN
       CALL PutMSG("SFMTM005"): IF A = 27 THEN GOTO RPTEXIT
       OmitBL = (A$ = SNGLKEY$(1))
    END IF
    GOSUB 600: IF (A = 27) OR (NOT OKAY) THEN GOTO RPTEXIT
    GN = GNUM: GOSUB 640
    GOTO RPTEXIT

RPTSCRN:
    Y = SFX(FMTTable.CT): Fmt = FMTTAB(Y)
    SYSVAR.MVAR = Fmt.NAME: SYSVAR.CHARTID = WINDOWS(Fmt.WIN).LOC
    CALL PutMSG("SFMTM005"): IF A = 27 THEN GOTO RPTEXIT
    OmitBL = (A$ = SNGLKEY$(1))
    GOSUB 600: IF (A = 27) OR (NOT OKAY) THEN GOTO RPTEXIT
    GN = GNUM: GOSUB 650
    GOTO RPTEXIT

RPTEXIT:
    CALL PutMSG("")
    EXIT SUB

600 ' Process PRINT Request
    A = 0
    'HTFMAX = 12: REDIM HTFTXT$(HTFMAX)
    'HTFTXT$(1) = "H01<Family History System>"
    'HTFTXT$(2) = "H01<[RPTNAME]>"
    'HTFTXT$(3) = "HG1<Format Group List>"
    'HTFTXT$(4) = "HF1<Group Format List>"
    'HTFTXT$(5) = "HD1<Format Field Detail Report>"
    'HTFTXT$(6) = "H01"
    'HTFTXT$(7) = "H01Screen Definition File: [FILENAME]"
    'HTFTXT$(8) = "TG2Group  CHG  FMTS         Description"
    'HTFTXT$(9) = "TF2Format List For Group: [MVAR]"
    'HTFTXT$(10) = "TF2CODE  WIN  COLOR  TYPE  FLDS  UFLDS"
    'HTFTXT$(11) = "TD2Group FMT   NAME  COLOR  ATR  LINE COL  DLTH FLTH      Default Text"
    'HTFTXT$(12) = "F02<Page [PAGE]>"
    CALL RptOPEN("SFMT"): IF OKAY THEN CALL PrintOPEN
    SYSVAR.FILENAME = ENV.SDF + "SFMT.SDF"
    X$ = SPACE$(255)
    IF PRT.OPT = 1 THEN
       RWB$ = WB$: RMB$ = "зд©ЁЁюды"
       ELSE
       RWB$ = "06233769": RMB$ = SB$
       FOR X = 1 TO 8
          Z = VAL(MID$(RWB$, X, 1)) + 1
          MID$(RWB$, X, 1) = MID$(PDFSymbols.SET3, Z, 1)
          MID$(RMB$, X, 1) = MID$(PDFSymbols.SET2, Z, 1)
       NEXT X
    END IF
    RETURN

620 ' Produce Format Group List
    PRT.SECTION = "G"
    FOR GN = 1 TO GRPMAX: LSET X$ = GRPDATA$(GN)
        MID$(PLINE$, 2, 4) = MID$(X$, 1, 4)
        MID$(PLINE$, 8, 4) = MID$(X$, 5, 4)
        MID$(PLINE$, 13, 3) = MID$(X$, 9, 3)
        MID$(PLINE$, 19, 50) = MID$(X$, 12, 50)
        CALL PrintLINE: IF A = 27 THEN GN = GRPMAX
    NEXT GN
    IF NOT ALLRPTS THEN
       CALL RptCLOSE: IF PRT.AGAIN THEN 620
       GOTO RPTEND
    END IF
    CALL PageBREAK
    GB = 1: GE = GRPMAX

630 ' Produce Group Format List
    PRT.SECTION = "F"
    FOR GN = GB TO GE
       LSET GN$ = SDFIndex(GN).NAME
       IF GB <> GE THEN CALL FmtGrpLOAD(GN$, GN)
       Y = SDFIndex(GN).FIRST
       IF (GN > GB) AND (Y > 0) THEN
          CALL WriteLINE
          SYSVAR.MVAR = GN$
          CALL PrintTITLE
       END IF
       WHILE Y > 0
          SYSVAR.MVAR = SDFIndex(GN).NAME
          Fmt = FMTTAB(Y)
          LSET PLINE$ = MID$(Fmt.NAME, 5, 4)
          RSET CNT$ = LTRIM$(STR$(Fmt.WIN))
          MID$(PLINE$, 6, 3) = CNT$
          RSET CNT$ = LTRIM$(STR$(Fmt.COLOR))
          MID$(PLINE$, 12, 3) = CNT$
          RSET CNT$ = LTRIM$(STR$(Fmt.TYPE))
          MID$(PLINE$, 19, 3) = CNT$
          RSET CNT$ = LTRIM$(STR$(Fmt.FLDCT))
          MID$(PLINE$, 25, 3) = CNT$
          RSET CNT$ = LTRIM$(STR$(Fmt.UFLDCT))
          MID$(PLINE$, 32, 3) = CNT$
          CALL PrintLINE
          IF A = 27 THEN Y = 0 ELSE Y = FmtCHN(Y).FWD
       WEND
       IF GB <> GE THEN
          IF (GN$ <> "XXXX") AND (GN$ <> "SFMT") THEN CALL FmtGrpFREE(GN$)
       END IF
       IF A = 27 THEN GN = GRPMAX
    NEXT GN
    CALL PageBREAK
    IF NOT ALLRPTS THEN
       CALL RptCLOSE: IF PRT.AGAIN THEN 630
       GOTO RPTEND
    END IF
    R$ = "3": GOTO 660

640 ' Produce Format Detail Report (Single Format)
    PRT.SECTION = "D"
    GOSUB FLDLIST
    CALL RptCLOSE: IF PRT.AGAIN THEN 640
    GOTO RPTEND

650 ' Produce Screen Display Section (Single Format)
    PRT.SECTION = "S"
    GOSUB FLDSCRN
    CALL RptCLOSE: IF PRT.AGAIN THEN 650
    GOTO RPTEND

660 ' Produce Format Detail Report AND Screen Display Report (All Formats)
    SELECT CASE R$
       CASE "3": PRT.SECTION = "D"
       CASE "4": PRT.SECTION = "S"
    END SELECT
    FOR GN = GB TO GE
       LSET PLINE$ = " "
       LSET GN$ = SDFIndex(GN).NAME: LSET GNX$ = GN$
       IF GB <> GE THEN CALL FmtGrpLOAD(GN$, GN)
       Y = SDFIndex(GN).FIRST
       WHILE Y > 0
          SELECT CASE R$
             CASE "3": GOSUB FLDLIST
             CASE "4": GOSUB FLDSCRN
          END SELECT
          IF A = 27 THEN Y = 0 ELSE Y = FmtCHN(Y).FWD
       WEND
       IF GB <> GE THEN
          IF (GN$ <> "XXXX") AND (GN$ <> "SFMT") THEN CALL FmtGrpFREE(GN$)
       END IF
       IF A = 27 THEN GN = GRPMAX
    NEXT GN
    CALL PageBREAK
    IF A <> 27 THEN
       IF ALLRPTS AND (R$ = "3") AND (NOT ShowSCREEN) THEN R$ = "4": GOTO 660
    END IF
    CALL RptCLOSE: IF NOT PRT.AGAIN THEN GOTO RPTEND
    IF ALLRPTS THEN 620 ELSE GOTO 660

RPTEND:
    CLOSE
    ALLRPTS = 0
    RETURN

FLDLIST:
    PRT.SECTION = "D"
    'Group  FMT  NAME  COLOR  ATR  LINE COL  DLTH FLTH      Default Text
    IF PAGE.LINE > 0 THEN CALL WriteLINE
    MID$(PLINE$, 2, 4) = GNX$
    IF NOT ShowSCREEN THEN LSET GNX$ = ""
    Fmt = FMTTAB(Y)
    Z = Fmt.BGNFLD
    MID$(PLINE$, 7, 4) = MID$(Fmt.NAME, 5, 4)
    WHILE Z > 0
       IF (PAGE.LINE = 0) THEN
          MID$(PLINE$, 2, 4) = GN$
          MID$(PLINE$, 7, 4) = MID$(Fmt.NAME, 5, 4)
       END IF
       FFLD = FLDTAB(Z)
       MID$(PLINE$, 13, 4) = FFLD.NAME
       RSET CNT$ = LTRIM$(STR$(FFLD.CATR \ 256))
       MID$(PLINE$, 19, 3) = CNT$
       RSET CNT$ = LTRIM$(STR$(FFLD.CATR AND 255))
       MID$(PLINE$, 26, 3) = CNT$
       RSET CNT$ = LTRIM$(STR$(FFLD.LLCC \ 256))
       MID$(PLINE$, 31, 3) = CNT$
       RSET CNT$ = LTRIM$(STR$(FFLD.LLCC AND 255))
       MID$(PLINE$, 36, 3) = CNT$
       RSET CNT$ = LTRIM$(STR$(FFLD.FLDL AND 255))
       MID$(PLINE$, 41, 3) = CNT$
       RSET CNT$ = LTRIM$(STR$(FFLD.FLDL \ 256))
       MID$(PLINE$, 46, 3) = CNT$
       MID$(PLINE$, 51, 255) = DTXT$(Z)
       CALL PrintLINE
       IF A = 27 THEN Z = 0 ELSE Z = FLDCHN(Z).FWD
    WEND
    IF ShowSCREEN THEN GOSUB FLDSCRN
    RETURN

FLDSCRN:
    IF NOT ShowSCREEN THEN PRT.SECTION = "S"
    Fmt = FMTTAB(Y): SYSVAR.MVAR = Fmt.NAME
    SYSVAR.CHARTID = WINDOWS(Fmt.WIN).LOC
    LL = WINDOWS(Fmt.WIN).LRR - WINDOWS(Fmt.WIN).ULR + 1
    CC = WINDOWS(Fmt.WIN).LRC - WINDOWS(Fmt.WIN).ULC + 1
    REDIM SLINE$(LL)
    FOR X = 0 TO LL: SLINE$(X) = SPACE$(CC): NEXT X
    IF WINDOWS(Fmt.WIN).LOC = "MSG " THEN LSET SB$ = RMB$ ELSE LSET SB$ = RWB$
    CALL ScrnBDRY(Fmt.WIN)
    Z = Fmt.BGNFLD: OKAY = -1
    WHILE Z > 0
       FFLD = FLDTAB(Z)
       L = FFLD.LLCC \ 256
       C = FFLD.LLCC AND 255
       DL = FFLD.FLDL AND 255
       IF C = 0 THEN
          X = LEN(DTXT$(Z)): C = (CC - X) \ 2: IF C < 1 THEN C = 1
       END IF
       IF C < CC THEN
          IF DTXT$(Z) <> "" THEN
             MID$(SLINE$(L), 1 + C, DL) = DTXT$(Z)
             ELSE
             MID$(SLINE$(L), 1 + C, DL) = STRING$(DL, "_")
          END IF
       END IF
       IF A = 27 THEN Z = 0 ELSE Z = FLDCHN(Z).FWD
    WEND
    IF PAGE.LINE > 0 THEN CALL WriteLINE
    IF NOT ShowSCREEN THEN CALL PrintTITLE
    FOR X = 0 TO LL
       IF OmitBL THEN OKAY = (RTRIM$(MID$(SLINE$(X), 2, CC - 2)) <> "")
       IF OKAY THEN LSET PLINE$ = SLINE$(X): CALL PrintLINE
    NEXT X
    RETURN

END SUB

