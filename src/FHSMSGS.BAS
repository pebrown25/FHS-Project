1  REM $INCLUDE: 'FHSCOMON.BAS'
   PN$ = " FAMILY HISTORY SYSTEM - FHSMSGS - Message Customization program"
   CY$ = " (C) Copyright 1995 by Phillip E. Brown"
   IF CONFIG.BP = 0 THEN RUN "FHSINIT"
   FUN = 0: CNT$ = "   ": GS$ = " ": L$ = " "
   FMTNAME$ = SPACE$(8): GNAME$ = SPACE$(4): PGMRTRN$ = "FHSCUST"
   IF NOT ENV.LARGEMEM THEN CALL MsgGrpLOAD("ALL", Z)
   'X = MSGMAX: XMSG$ = SPACE$(100)
   'WHILE X > 0
   '   LSET XMSG$ = MID$(MSGDATA$(X), 1, 6)
   '   MID$(XMSG$, 8, 255) = MID$(MSGDATA$(X), 7, 255)
   '   MSGDATA$(X) = RTRIM$(XMSG$)
   '   X = X - 1
   'WEND

10 ' Prepare TABLES
   CALL TableLOAD("GROUPS", TGRP, 0)
   GrpTABLE = TABLE(TGRP)
   GrpTABLE.FMTNAME = "MSGSS015": GrpTABLE.FMTNUM = 0
   GrpTABLE.KEYNAME = "KEY"
   GrpTABLE.SIZE = GRPMAX
   REDIM GRPDATA$(GRPMAX), GRPDATACHN(GRPMAX) AS ChainPTRS
   CALL ChnINIT(GRPDATACHN(), 1, GRPMAX)
   X = 0: Y = GrpTABLE.FT
   WHILE X < GRPMAX
      X = X + 1
      GRPDATA$(X) = MDFIndex(X).NAME + "       "
      RSET CNT$ = LTRIM$(STR$(MDFIndex(X).CHG AND 127))
      MID$(GRPDATA$(X), 5, 3) = CNT$
      RSET CNT$ = LTRIM$(STR$(MDFIndex(X).SIZE))
      MID$(GRPDATA$(X), 9, 3) = CNT$
      IF ((MDFIndex(X).CHG AND 128) > 0) THEN
         MID$(GRPDATA$(X), 8, 1) = "+"
      ELSE
         MID$(GRPDATA$(X), 8, 1) = " "
      END IF
   WEND
   X$ = SPACE$(255): Z = 0
   WHILE Y > 0
      X = Z: FOUND = 0
      WHILE (NOT FOUND) AND (X < GRPMAX)
         X = X + 1
         FOUND = (MDFIndex(X).NAME = LEFT$(TABDATA$(Y), 4))
      WEND
      IF FOUND THEN
         MDFIndex(X).DESC = Y
         LSET X$ = GRPDATA$(X) + MID$(TABDATA$(Y), 5, 255)
         GRPDATA$(X) = RTRIM$(X$)
      END IF
      Y = TABDATACHN(Y).FWD: Z = Z + 1
   WEND
   GrpTABLE.FT = 1: GrpTABLE.LT = GrpTABLE.SIZE
   GRPDATACHN(1).BWD = -TGRP
40 ' Format Display
   IF FUN = 0 THEN CALL WinCLR
   CALL FmtFIND("MSGSS010"): FMT010 = CURFMT
   CALL FmtFindFLD("MDF"): IF FOUND THEN DTXT$(WIN.CURFLD) = CONFIG.MDF
   CALL FmtFindFLD("MCT")
   IF FOUND THEN DTXT$(WIN.CURFLD) = LTRIM$(STR$(MDFHdr.MSGCNT))
   CALL FmtFindFLD("EOF")
   IF FOUND THEN DTXT$(WIN.CURFLD) = LTRIM$(STR$(MDFHdr.EOFRBA))
   CALL WinFORMAT(0)
   CALL TableOPEN(GrpTABLE)
   CALL TableSHOW(GrpTABLE, GRPDATA$(), GRPDATACHN())
200 ' Get Processing Option
    LSET PGMRTRN$ = "FHSCUST"
    CALL PutMSG(" ")
    CALL FmtFIND("XXXXS091"): CALL WinFORMAT(0)
    CALL WinFORMAT(WIN.OFMT)
250 CALL TableSELECT(GrpTABLE, GRPDATA$(), GRPDATACHN())
    SELECT CASE A
       CASE 13: FUN$ = "SEL"
       CASE 59: FUN$ = "F1"
       CASE 64: FUN$ = "F6"
       CASE 66: CALL DOSShell: GOTO 40
       CASE 67: FUN$ = "F9"
       CASE 94: ENV.EDMAST = NOT ENV.EDMAST: GOTO 200
       CASE 99: FUN$ = "F6": LSET PGMRTRN$ = "FHSPRTC"
       CASE 109: FUN$ = "F6": LSET PGMRTRN$ = "FHSRPTS": CALL RptLOAD("MSGS")
       CASE ELSE: SOUND BP, DUR: GOTO 250
    END SELECT
    CALL OPTHilite(FUN$, "", 6)
    SELECT CASE FUN$
       CASE "F1": GOTO 500
       CASE "F6": IF PGMRTRN$ <> "FHSCUST" THEN GOTO 1000
       CASE "F9": GOTO 1000
    END SELECT
300 ' Update Messages in GROUP
    LSET GS$ = MID$(GRPDATA$(GrpTABLE.CT), 8, 1)
    LSET GNAME$ = GRPDATA$(GrpTABLE.CT)
    CALL MsgGrpLOAD(GNAME$, GNUM)
    IF NOT OKAY THEN
       SYSVAR.MVAR = GNAME$
       CALL PutMSG("MSGSM002")
       MID$(GRPDATA$(GrpTABLE.CT), 8, 1) = "?"
       GOTO 40
    END IF
    IF (FUN$ = "SEL") AND (MDFIndex(GNUM).SIZE = 0) AND (NOT ENV.EDMAST) THEN
       SOUND BP, DUR: GOTO 200
    END IF
    IF FUN$ = "F6" THEN GOSUB 600: GOTO 40   ' Process PRINT Option
    MSGTable.TABNAME = MDFIndex(GNUM).NAME
    MSGTable.OPEN = MDFIndex(GNUM).OPEN
    MSGTable.SIZE = MDFIndex(GNUM).SIZE
    MSGTable.XL = 87: MSGTable.KO = 1: MSGTable.KL = 4
    MSGTable.FMTNAME = "MSGSS025": MSGTable.FMTNUM = 0
    MSGTable.KEYNAME = "KEY"
    MSGTable.ML = 0
    MSGTable.FT = MDFIndex(GNUM).FIRST
    MSGTable.LT = MDFIndex(GNUM).LAST
    MSGTable.FF = MSGDATAFREE
    X = WIN.ULR: WIN.ULR = GrpTABLE.TL - 1: CALL WinCLR: WIN.ULR = X
    CALL FmtFIND("MSGSS020"): FMT020 = CURFMT
    CALL FmtFindFLD("GRP"): IF FOUND THEN UTXT$(FFLD.UTXT) = MSGTable.TABNAME
    CALL TableOPEN(MSGTable)
320 ' Format Group Message Screen
    CALL WinFORMAT(FMT020)
    CALL TableSHOW(MSGTable, MSGDATA$(), MSGDATACHN())
340 ' SET Option Format
    IF ENV.EDMAST THEN LSET FMTNAME$ = "XXXXS093" ELSE LSET FMTNAME$ = "XXXXS094"
    CALL FmtFIND(FMTNAME$)
    IF NOT FOUND THEN SOUND BP, DUR: GOTO 40
350 ' SELECT Message For Update
    CALL OPTDisplay(0)
    IF MSGTable.SIZE = 0 THEN A = 255: GOTO 375
    CALL TableSELECT(MSGTable, MSGDATA$(), MSGDATACHN())
    IF (A = 27) OR ((LEN(A$) = 2) AND (A = 67)) THEN GOTO 40
    FUN$ = ""
    IF ENV.EDMAST AND (A = 61) THEN
       CALL PutMSG("MSGSM003")
       IF A$ <> SNGLKEY$(1) THEN GOTO 350
       A = 61
    END IF
375 ' Determine Response Option
    SELECT CASE A
       CASE 13: FUN$ = "CHG"
       CASE 27: FUN$ = "F9"
       CASE 64: FUN$ = "F6"
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
    IF FUN$ = "F6" THEN
       ONLY1 = -1: GOSUB 615
       IF PRT.OPT = 1 THEN
          CALL WinFORMAT(FMT010)
          GOTO 320
       END IF
       GOTO 350
    END IF
    IF FUN$ = "F9" THEN GOTO 40
    IF FUN$ = "ADD" THEN
       IF MSGDATAFREE = 0 THEN
          MSGDATAFREE = MSGMAX + 1
          MSGMAX = MSGMAX + 1
          REDIM PRESERVE MSGDATA$(MSGMAX), MSGDATACHN(MSGMAX) AS ChainPTRS
       END IF
    END IF
400 ' UPDATE TABLE Entry
    MSGTable.FF = MSGDATAFREE
    CALL TableUPDT(FUN$, MSGTable, MSGDATA$(), MSGDATACHN())
    MSGDATAFREE = MSGTable.FF
    MDFIndex(GNUM).OPEN = (MSGTable.SIZE > 0)
    MDFIndex(GNUM).SIZE = MSGTable.SIZE
    MDFIndex(GNUM).FIRST = MSGTable.FT
    MDFIndex(GNUM).LAST = MSGTable.LT
    RSET CNT$ = LTRIM$(STR$(MSGTable.SIZE))
    MID$(GRPDATA$(GrpTABLE.CT), 9, 3) = CNT$
    IF (A <> 27) OR (FUN$ = "MOV") THEN
       MID$(GRPDATA$(GrpTABLE.CT), 8, 1) = "+"
       MDFIndex(GNUM).CHG = MDFIndex(GNUM).CHG OR 128
       MDFHdr.VER = CHR$(255)
    END IF
    IF MSGTable.SIZE = 0 THEN GOTO 40
    GOTO 350

500 ' Process FILE Requests
    CALL PutMSG("XXXXM005")
    CONFIG.MDF = ENV.MDF
    IF A = 27 THEN GOTO 40 ELSE R$ = A$
    IF (R$ < "0") OR (R$ > "3") THEN SOUND BP, DUR: GOTO 500
    GOSUB 510: IF A = 27 THEN GOTO 500
    ZCHG = (R$ = "0")
    SELECT CASE R$
       CASE "0": GOTO 520
       CASE "1": GOTO 520
       CASE "2": GOTO 530
       CASE "3": GOTO 540
    END SELECT

510 ' Get MDF File Prefix
    CALL FmtFIND("MSGSS050")
    CALL FmtFindFLD("MDF")
    LSET UTXT$(FFLD.UTXT) = CONFIG.MDF
    CALL MsgFORMAT
    IF A = 27 THEN RETURN
    CALL FmtFindFLD("MDF"): CONFIG.MDF = UTXT$(FFLD.UTXT)
    IF CONFIG.MDF <> ENV.MDF THEN CONFIG.VER = -1
    RETURN

520 ' Save TABLE Changes Into MDF File
    CALL MDFOpen("NEW"): IF NOT OKAY THEN GOTO 500
    CALL PutMSG("MSGSM004")
    MDFHdr.FTYPE = "MDF": MDFHdr.VER = CHR$(0)
    X = 0: MDFHdr.MSGCNT = 0
    FOR GNUM = 1 TO GRPMAX
       IF MDFIndex(GNUM).SIZE > 0 THEN
          X = X + 1
          MDFHdr.MSGCNT = MDFHdr.MSGCNT + MDFIndex(GNUM).SIZE
       END IF
    NEXT GNUM
    MDFHdr.GRPCNT = X
    PUT #1, , MDFHdr
    FOR GNUM = 1 TO GRPMAX
       IF MDFIndex(GNUM).SIZE > 0 THEN
          LSET MDFGrpDIR.NAME = MDFIndex(GNUM).NAME
          CHG = MDFIndex(GNUM).CHG
          IF (CHG AND 128) > 0 THEN CHG = (CHG AND 127) + 1
          CHG = CHG AND 127
          IF ZCHG THEN CHG = 0
          MDFIndex(GNUM).CHG = CHG
          MDFGrpDIR.CHG = CHG
          MDFGrpDIR.MSGCNT = MDFIndex(GNUM).SIZE
          PUT #1, , MDFGrpDIR
       END IF
    NEXT GNUM
    FOR GNUM = 1 TO GRPMAX
       IF MDFIndex(GNUM).SIZE > 0 THEN
          MDFIndex(GNUM).BGNRBA = SEEK(1)
          X = MDFIndex(GNUM).FIRST
          WHILE X > 0
             L$ = CHR$(LEN(MSGDATA$(X)))
             PUT #1, , L$
             PUT #1, , MSGDATA$(X)
             X = MSGDATACHN(X).FWD
          WEND
       END IF
    NEXT GNUM
    MDFHdr.EOFRBA = SEEK(1) - 1
    PUT #1, 1, MDFHdr
    O = LEN(MDFHdr) + 1
    FOR GNUM = 1 TO GRPMAX
       IF MDFIndex(GNUM).SIZE > 0 THEN
          GET #1, O, MDFGrpDIR
          MDFGrpDIR.FIRST = MDFIndex(GNUM).BGNRBA
          PUT #1, O, MDFGrpDIR
          O = O + LEN(MDFGrpDIR)
       END IF
    NEXT GNUM
    CLOSE
    GOTO 590

530 ' Load MESSAGE Data from MDF File
     IF MDFHdr.VER = CHR$(255) THEN
        CALL PutMSG("XXXXM020")
        IF A$ <> SNGLKEY$(1) THEN GOTO 500
     END IF
     CALL PutMSG("MSGSM005")
     CALL MsgGrpLOAD("ALL", Z): IF NOT OKAY THEN GOTO 500
     GOTO 590

540 ' Merge MESSAGE Data from MDF File
     ' Only Merge Existing Messages
     CALL MDFOpen("READ"): IF NOT OKAY THEN GOTO 500
     CALL PutMSG("MSGSM006")
     X = 0: GCNT = MDFFileHDR.GrpCnt: OD = LEN(MDFHdr) + 1
     WHILE X < GCNT
        X = X + 1: GET #1, OD, MDFGrpDIR: OD = OD + LEN(MDFGrpDIR)
        Y = 0: FOUND = 0
        WHILE (Y < GRPMAX) AND (NOT FOUND)
           Y = Y + 1: FOUND = (MDFGrpDIR.NAME = MDFIndex(Y).NAME)
        WEND
        IF FOUND THEN
           W = 0: SEEK #1, MDFGrpDIR.FIRST
           WHILE W < MDFGrpDIR.MSGCNT
              GET #1, , L$: L = ASC(L$)
              MSG$ = SPACE$(L)
              GET #1, , MSG$
              LSET GNAME$ = MSG$
              Z = MDFIndex(Y).FIRST: FOUND = 0
              WHILE (Z < 0) AND (NOT FOUND)
                 Z = Z + 1: FOUND = (LEFT$(MSGDATA$(Z), 4) = GNAME$)
                 IF NOT FOUND THEN Z = MSGDATACHN(Z).FWD
              WEND
              IF FOUND AND (MSG$ <> MSGDATA$(Z)) THEN
                 'MSGDATA$(Z) = LEFT$(MSG$, 8) + MID$(MSGDATA$(X), 9, 1) + MID$(MSG$, 9, 255)
                 MSGDATA$(Z) = MSG$
                 MDFIndex(Y).CHG = MDFIndex(Y).CHG OR 128
                 MDFHdr.VER = CHR$(255)
              END IF
           WEND
        END IF
     WEND
     GOTO 590

590 ' RESET Count & Status Fields in GRPDATA$()
     Y = GRPMAX
     WHILE Y > 0
        RSET CNT$ = LTRIM$(STR$(MDFIndex(Y).CHG AND 127))
        MID$(GRPDATA$(Y), 5, 3) = CNT$
        IF (MDFIndex(Y).CHG AND 128) > 0 THEN
           MID$(GRPDATA$(Y), 8, 1) = "+"
        ELSE
           MID$(GRPDATA$(Y), 8, 1) = " "
        END IF
        RSET CNT$ = LTRIM$(STR$(MDFIndex(Y).SIZE))
        MID$(GRPDATA$(Y), 9, 3) = CNT$
        Y = Y - 1
     WEND
     GOTO 40

600 ' Process PRINT Request
    SYSVAR.MVAR = GNAME$
610 CALL PutMSG("MSGSM010"): IF A = 27 THEN RETURN
    IF (A$ < "1") OR (A$ > "2") THEN CALL ErrBEEP(0): GOTO 610
    ONLY1 = (A$ = "1")

615 ' Print Message Report
    'HTFMAX = 6: REDIM HTFTXT$(HTFMAX)
    'HTFTXT$(1) = "H01<Family History System>"
    'HTFTXT$(2) = "H01<[RPTNAME]>"
    'HTFTXT$(3) = "H01"
    'HTFTXT$(4) = "H01Message Definition File:[FILENAME]"
    'HTFTXT$(5) = "T02Group  CODE  R  D         Description"
    'HTFTXT$(6) = "F02<Page [PAGE]>"
    CALL RptOPEN("MSGS"): IF (NOT OKAY) THEN 625
    CALL PrintOPEN
    SYSVAR.FILENAME = ENV.MDF + "MSGS.MDF"
    IF ONLY1 THEN GB = GNUM: GE = GNUM:  ELSE GB = 1: GE = GRPMAX
620 FOR GN = GB TO GE
        IF MDFIndex(GN).SIZE > 0 THEN
           IF PAGE.LINE > 0 THEN CALL WriteLINE
           MID$(PLINE$, 2, 4) = MDFIndex(GN).NAME
           GOSUB 650
           IF A = 27 THEN GN = GE
        END IF
    NEXT GN
    CALL PageBREAK
625 ' End of Print
    CALL RptCLOSE: IF PRT.AGAIN THEN 620
    RETURN

650 ' Print Message Group #GN
    GMN = MDFIndex(GN).FIRST
    WHILE GMN > 0
       IF PAGE.LINE = 0 THEN MID$(PLINE$, 2, 4) = MDFIndex(GN).NAME
       MID$(PLINE$, 8, 4) = LEFT$(MSGDATA$(GMN), 4)
       MID$(PLINE$, 14, 1) = MID$(MSGDATA$(GMN), 5, 1)
       MID$(PLINE$, 17, 1) = MID$(MSGDATA$(GMN), 6, 1)
       MID$(PLINE$, 20, 1) = MID$(MSGDATA$(GMN), 7, 1)
       MID$(PLINE$, 23, 255) = MID$(MSGDATA$(GMN), 8, 255)
       CALL PrintLINE
       GMN = MSGDATACHN(GMN).FWD
    WEND
    RETURN

1000 ' Return to FAMMENU.BAS
     IF (MDFHdr.VER = CHR$(255)) THEN
        CALL PutMSG("XXXXM020")
        IF A$ <> SNGLKEY$(1) THEN GOTO 200
     END IF
     CALL PgmPREP(PGMRTRN$): IF NOT OKAY THEN GOTO 200
     IF NOT ENV.LARGEMEM THEN CALL MsgGrpLOAD("INITIAL", Z)
     CALL PutMSG("")
     CLOSE
     CHAIN PGMRTRN$

