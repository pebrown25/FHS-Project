DECLARE SUB FF2GetRec (RNUM%)
DECLARE SUB PrintOPEN ()
DECLARE SUB RptCLOSE ()
DECLARE SUB ChkID (X&)
DECLARE SUB EvenBREAK (X$)
DECLARE SUB FamCLOSE ()
DECLARE SUB FAMOpen (OPT$)
DECLARE SUB FDFGetSet (XSETUP$)
DECLARE SUB FDFOpen (OPT$)
DECLARE SUB FF1GetRec (RNUM%)
DECLARE SUB FF1Open (OPT$)
DECLARE SUB FF1PutREC ()
DECLARE SUB FF2Open (OPT$)
DECLARE SUB FF3GetRec (RNUM%)
DECLARE SUB FF3Open (OPT$)
DECLARE SUB FF3PutRec ()
DECLARE SUB PrintLINE ()
DECLARE SUB PutDATA ()
DECLARE SUB RREFAdd ()
DECLARE SUB RREFBld ()
DECLARE SUB RREFClose ()
DECLARE SUB RREFOpen (OPT$)
DECLARE SUB RWRKOpen (OPT$, XTYPE%, DSNFMT%)
DECLARE SUB SortStatACCUM ()
DECLARE SUB SortStatSHOW ()
DECLARE SUB SWRKOpen (XOPT$)
DECLARE FUNCTION XRPTDate$ (MD%, Y%, O%)
DECLARE FUNCTION Relation$ (RT$, RSEX$, AGL%, RGL%, RCODE$())
DECLARE FUNCTION FMTNAME$ (NFMT%)
DECLARE FUNCTION FFSTAT$ (ST$, S$)
DECLARE FUNCTION FDSTAT$ (ST$)
DECLARE FUNCTION SFILL$ (X%, L%)
DECLARE FUNCTION ZFILL$ (X%, L%)
DECLARE FUNCTION UCX$ (X$)
DECLARE FUNCTION SNX$ (A%)
DECLARE FUNCTION FFRNUM% (RBA&, RLTH%)
DECLARE FUNCTION FFRBA& (RNUM%, RLTH%)
DECLARE FUNCTION DateDif$ (D1MD%, D1Y%, D2MD%, D2Y%)
DECLARE FUNCTION XSCRNTime$ (HHMM%)
DECLARE FUNCTION XSCRNDate$ (MD%, Y%)
DECLARE FUNCTION TimeHHMM% (X$)
DECLARE FUNCTION DATEMD% (X$, Y%)
DECLARE FUNCTION DATEY% (X$, Y%)
DECLARE FUNCTION TIMEHM% (X$)
REM $INCLUDE: 'FHSCOMON.BAS'

REM $DYNAMIC
SUB ChkID (X&)
    IF X& > FF1HDR.MAXID THEN CALL ErrBEEP(0): OKAY = 0 ELSE OKAY = -1
END SUB

SUB ComPRINT (ComPTR, COMBGN)
2920 ' Print COMMENT Lines
     BGNBKT = 0
     COFST = PAGE.COFST: PAGE.COFST = COFST + PAGE.CINDENT
     BWIDTH = PAGE.BWIDTH: PAGE.BWIDTH = BWIDTH - PAGE.CINDENT
     IF PAGE.LOFST = 0 THEN PAGE.LOFST = COFST
     IF RPTOPTION.FREEFORM OR COMBGN THEN
        LSET CL$ = ";": OFST = 3
        ELSE
        LSET CL$ = RFLBL$(1)
        OFST = INSTR(CL$, " ") + 1
     END IF
     ' Print Comment Information
     WHILE ComPTR <> 0: CALL FF3GetRec(ComPTR): ComPTR = FF3COMMENT.FPTR
        MID$(CL$, OFST, 38) = FF3COMMENT.TEXT
        IF RPTOPTION.IGNBKT THEN
           ENDBKT = 1
           WHILE ENDBKT <> 0
              IF BGNBKT = 0 THEN BGNBKT = INSTR(OFST, CL$, "{")
              IF BGNBKT = 0 THEN
                 ENDBKT = 0
                 ELSE
                 OFST = BGNBKT
                 ENDBKT = INSTR(BGNBKT, CL$, "}")
                 IF ENDBKT = 0 THEN
                    MID$(CL$, BGNBKT, 38) = SPACE$(38)
                    ELSE
                    MID$(CL$, BGNBKT, 1) = " "
                    MID$(CL$, BGNBKT + 1, 38) = MID$(CL$, ENDBKT + 1, 38)
                    BGNBKT = 0
                 END IF
              END IF
           WEND
           ELSE
           IF RPTOPTION.RMVBKT THEN
              X = 1
              WHILE X > 0
                 X = INSTR(X, CL$, "{")
                 IF X > 0 THEN MID$(CL$, X, 1) = " "
              WEND
              X = 1
              WHILE X > 0
                 X = INSTR(X, CL$, "}")
                 IF X > 0 THEN MID$(CL$, X, 1) = " "
              WEND
           END IF
        END IF

        ' Process NEWLINE requests
        LCTRL = 1
        WHILE LCTRL > 0
           LCTRL = INSTR(1, CL$, "\")
           IF LCTRL > 0 THEN
              HLCNT = PAGE.LINE
              X$ = RTRIM$(MID$(CL$, LCTRL + 1, 255))
              LSET CL$ = LEFT$(CL$, LCTRL - 1): CALL PutDATA: COMBGN = 1
              IF (HLCNT = PAGE.LINE) THEN CALL WriteLINE: PAGE.LOFST = PAGE.COFST
              LSET CL$ = X$
           END IF
        WEND

        ' Finish Processing COMMENT Record
        X$ = RTRIM$(CL$)
        IF NOT ((X$ = "") OR (X$ = ";") OR (X$ = RFLBL$(1))) THEN
           CALL PutDATA: X$ = "": COMBGN = 1: IF BGNBKT > 0 THEN BGNBKT = 1
        END IF
        OFST = LEN(X$) + 1
        IF BGNBKT > 0 THEN
           BGNBKT = OFST
           ELSE
           IF MID$(FF3COMMENT.TEXT, 38, 1) = " " THEN OFST = OFST + 1
        END IF
        IF A = 27 THEN ComPTR = 0
     WEND

    ' Return to Caller
    PAGE.COFST = COFST: PAGE.BWIDTH = BWIDTH: PAGE.LWIDTH = BWIDTH
END SUB

SUB ComputeSNDX (T$, S$)

     ' Compute Soundex
     X = 1: TL = LEN(T$): Y = 1: SL = LEN(S$)
     SXA = 0: SXH = 0
     LSET S$ = STRING$(SL, ASC("0"))
SNDX1:
     IF (X > TL) OR (Y > SL) THEN GOTO SNDXE
     SXA = ASC(MID$(T$, X, 1)): X = X + 1
     SXT = ASC(MID$(SNDXTB$, SXA, 1))
     IF (SXT > 254) THEN IF Y = 1 THEN GOTO SNDX1 ELSE GOTO SNDXE
     IF Y = 1 THEN
        MID$(S$, 1, 1) = UCX$(CHR$(SXA))
        ELSE
        IF (SXT = 48) OR (SXT = SXH) THEN GOTO SNDX1
        MID$(S$, Y, 1) = CHR$(SXT)
        SXH = SXT
     END IF
     Y = Y + 1
     GOTO SNDX1
SNDXE: 'RETURN

END SUB

FUNCTION DateDif$ (D1MD, D1Y, D2MD, D2Y)
   MD1 = D1MD: Y1 = D1Y
   MD2 = D2MD: Y2 = D2Y
   IF (Y2 = 0) AND (MD2 = 0) THEN Y2 = DATEY(DATE$, 1): MD2 = DATEMD(DATE$, 1)
   IF (Y1 = 0) OR (Y2 = 0) OR (Y2 = 9999) THEN
      DateDif$ = " ??"
      ELSE
      IF MD2 < MD1 THEN DIF = Y2 - Y1 - 1 ELSE DIF = Y2 - Y1
      IF DIF < 0 THEN
         DateDif$ = " **"
         ELSE
         DateDif$ = LTRIM$(STR$(DIF))
      END IF
   END IF
END FUNCTION

FUNCTION DATEMD (X$, Y)
   M = VAL(MID$(X$, 1 - 3 * (Y = 2) - 5 * (Y = 3), 2))
   D = VAL(MID$(X$, 4 + 3 * (Y = 2) - 5 * (Y = 3), 2))
   DATEMD = 100 * M + D
END FUNCTION

FUNCTION DATEY (X$, Y)
    DATEY = VAL(MID$(X$, 7 + 6 * (Y = 3), 4))
END FUNCTION

REM $STATIC
SUB EvenBREAK (X$)
    IF (RPTOPTION.EVENPAGE) AND ((PAGE.COUNT MOD 2) > 0) THEN
       CALL PageBREAK: IF X$ = "EOF" THEN PRT.EOF = -1
       CALL GetRVAR("BLANKPG", X)
       BL$ = PLINE$
       IF FOUND THEN
          LSET BL$ = RVAR$(X)
          CALL HTFBuild(BL$, PAGE.BWIDTH)
          ELSE
          BL$ = ""
       END IF
       X = PAGE.BLENGTH / 3
       WHILE X > 0: CALL WriteLINE: X = X - 1: WEND
       LSET PLINE$ = BL$
       CALL WriteLINE
       ELSE
       IF X$ = "EOF" THEN PRT.EOF = -1
    END IF
    CALL PageBREAK
END SUB

REM $DYNAMIC
SUB FamCLOSE
   IF ENV.FFOPEN AND FFCHANGED THEN
      Y = DATEY(DATE$, 1): MD = DATEMD(DATE$, 1)
      IF FF1HDR.RevLVL > 255 THEN
         FF1HDR.UPDTY = Y: FF1HDR.UPDTMD = MD
         FF1HDR.RevLVL = FF1HDR.RevLVL AND 255: PUT #1, 1, FF1HDR
      END IF
      IF FF2HDR.RevLVL > 255 THEN
         FF2HDR.UPDTY = Y: FF2HDR.UPDTMD = MD
         FF2HDR.RevLVL = FF2HDR.RevLVL AND 255: PUT #2, 1, FF2HDR
      END IF
      IF FF3HDR.RevLVL > 255 THEN
         FF3HDR.UPDTY = Y: FF3HDR.UPDTMD = MD
         FF3HDR.RevLVL = FF3HDR.RevLVL AND 255: PUT #3, 1, FF3HDR
      END IF
   END IF
   CLOSE : ENV.FFOPEN = 0: FFCHANGED = 0: PRT.OPEN = CHR$(0)
END SUB

REM $STATIC
SUB FAMOpen (OPT$)
    IF ENV.FFOPEN THEN CALL FamCLOSE
    X$ = OPT$
    IF (X$ = "VLDT") OR (X$ = "VLDTQ") THEN VOPT = -1: X$ = "WRITE"
    CALL FF1Open(X$): IF NOT OKAY THEN GOTO FAMOpenEND
    CALL FF2Open(X$): IF NOT OKAY THEN GOTO FAMOpenEND
    CALL FF3Open(X$): IF NOT OKAY THEN GOTO FAMOpenEND
    IF (X$ <> "INIT") AND (X$ <> "INITQ") THEN GOTO FAMOpenEND
    CLOSE
    CALL FF1Open("NEW"): IF NOT OKAY THEN GOTO FAMOpenEND
    CALL FF2Open("NEW"): IF NOT OKAY THEN GOTO FAMOpenEND
    CALL FF3Open("NEW")
FAMOpenEND:
    ENV.FFOPEN = OKAY: FFCHANGED = 0
    IF NOT OKAY THEN
       CLOSE
       ELSE
       IF NOT VOPT THEN
          IF ((FF1HDR.RevLVL > 255) OR (FF2HDR.RevLVL > 255) OR (FF3HDR.RevLVL > 255)) THEN
             CALL PutMSG("XXXXM073")
          END IF
       END IF
    END IF
END SUB

SUB FDFGetSet (XSETUP$)
    OKAY = 0: SYSVAR.MVAR = XSETUP$
    IF XSETUP$ = "DEFAULT" THEN
       X = FDFHdr.DEFAULT
       IF (X < 1) OR (X > FDFHdr.SETUPS) THEN X = 1: FDFHdr.DEFAULT = 1
    ELSE
       ZSETUP$ = SPACE$(8): LSET ZSETUP$ = XSETUP$
       X = FSUTable.FT: FOUND = 0
       WHILE (X > 0) AND (NOT FOUND)
          FOUND = (MID$(FSETUP$(X), 1, 8) = ZSETUP$)
          IF NOT FOUND THEN X = FSUCHN(X).FWD
       WEND
       IF NOT FOUND THEN
          CALL PutMSG("XXXXM027")
          EXIT SUB
       END IF
       IF X = ENV.FSETUP THEN OKAY = -1: EXIT SUB
    END IF
    IF FDFSURBA(X) = 0 THEN
       CALL PutMSG("XXXXM028")
       EXIT SUB
    END IF
    CALL FDFOpen("READ")
    IF NOT OKAY THEN EXIT SUB
    ENV.FSETUP = X
    GET #1, FDFSURBA(X), FDFSetup
    FOR X = 1 TO FDFHdr.FAMDSN
        GET #1, , FT$(X)
    NEXT X
    FAMFPFX$ = "": WRKFPFX$ = "": TRANPFX$ = "": RPTFPFX$ = ""
    D$ = FDFSetup.FAMDRV: IF ("A" <= D$) AND (D$ <= "Z") THEN FAMFPFX$ = D$ + ":"
    IF MID$(FDFSetup.FAMDIR, 1, 1) <> "\" THEN
       SYSVAR.FILENAME = FAMFPFX$
       FAMFPFX$ = CURDIR$(FAMFPFX$)
       IF MID$(FAMFPFX$, LEN(FAMFPFX$), 1) <> "\" THEN FAMFPFX$ = FAMFPFX$ + "\"
    END IF
    FAMFPFX$ = RTRIM$(FAMFPFX$ + FDFSetup.FAMDIR)
    IF MID$(FAMFPFX$, LEN(FAMFPFX$), 1) <> "\" THEN FAMFPFX$ = FAMFPFX$ + "\"
    D$ = FDFSetup.WRKDRV: IF ("A" <= D$) AND (D$ <= "Z") THEN WRKFPFX$ = D$ + ":"
    IF MID$(FDFSetup.WRKDIR, 1, 1) <> "\" THEN
       SYSVAR.FILENAME = WRKFPFX$
       WRKFPFX$ = CURDIR$(WRKFPFX$)
       IF MID$(WRKFPFX$, LEN(WRKFPFX$), 1) <> "\" THEN WRKFPFX$ = WRKFPFX$ + "\"
    END IF
    WRKFPFX$ = RTRIM$(WRKFPFX$ + FDFSetup.WRKDIR)
    IF MID$(WRKFPFX$, LEN(WRKFPFX$), 1) <> "\" THEN WRKFPFX$ = WRKFPFX$ + "\"
    D$ = FDFSetup.TRANDRV: IF ("A" <= D$) AND (D$ <= "Z") THEN TRANPFX$ = D$ + ":"
    IF MID$(FDFSetup.TRANDIR, 1, 1) <> "\" THEN
       SYSVAR.FILENAME = TRANPFX$
       TRANPFX$ = CURDIR$(TRANPFX$)
       IF MID$(TRANPFX$, LEN(TRANPFX$), 1) <> "\" THEN TRANPFX$ = TRANPFX$ + "\"
    END IF
    TRANPFX$ = RTRIM$(TRANPFX$ + FDFSetup.TRANDIR)
    IF MID$(TRANPFX$, LEN(TRANPFX$), 1) <> "\" THEN TRANPFX$ = TRANPFX$ + "\"
    D$ = FDFSetup.RPTDRV: IF ("A" <= D$) AND (D$ <= "Z") THEN RPTFPFX$ = D$ + ":"
    IF MID$(FDFSetup.RPTDIR, 1, 1) <> "\" THEN
       SYSVAR.FILENAME = RPTFPFX$
       RPTFPFX$ = CURDIR$(RPTFPFX$)
       IF MID$(RPTFPFX$, LEN(RPTFPFX$), 1) <> "\" THEN RPTFPFX$ = RPTFPFX$ + "\"
    END IF
    RPTFPFX$ = RTRIM$(RPTFPFX$ + FDFSetup.RPTDIR)
    IF MID$(RPTFPFX$, LEN(RPTFPFX$), 1) <> "\" THEN RPTFPFX$ = RPTFPFX$ + "\"
    OKAY = -1: GOTO FDFGetSetExit

FDFGetSetErr:
    CALL ERRMessage("XXXXM098", ERR, ERL)
    OKAY = 0
    RESUME FDFGetSetExit

FDFGetSetExit:
    CALL FamCLOSE
END SUB

SUB FDFOpen (OPT$)
    OKAY = 0
    LSET SYSVAR.FILENAME = CONFIG.FDF + ".FDF"
    ON LOCAL ERROR GOTO NOFDF
CHKFDF:
    OPEN SYSVAR.FILENAME FOR INPUT AS #1
    CLOSE #1
    IF OPT$ = "NEW" THEN
       CALL PutMSG("XXXXM002")
       IF A$ <> SNGLKEY$(1) THEN GOTO FDFOpenEND
       KILL SYSVAR.FILENAME
       GOTO NEWFDF
    END IF
    GOTO OLDFDF
NOFDF:
    X = ERR: CALL FamCLOSE
    IF OPT$ = "NEW" THEN RESUME NEWFDF
    CALL ERRMessage("XXXXM030", X, 0)
    IF A$ = CHR$(27) THEN RESUME FDFOpenEND
    RESUME CHKFDF
OLDFDF:
    ON LOCAL ERROR GOTO FDFOpenERR
    SELECT CASE OPT$
       CASE "INIT": OPEN SYSVAR.FILENAME FOR BINARY ACCESS READ AS #1
       CASE "READ": OPEN SYSVAR.FILENAME FOR BINARY ACCESS READ AS #1
       CASE "WRITE": OPEN SYSVAR.FILENAME FOR BINARY ACCESS READ WRITE AS #1
       CASE ELSE: GOTO FDFOpenEND
    END SELECT
    GET #1, , FDFHdr
    LSET SYSVAR.MVAR = "FDF"
    IF FDFHdr.FTYPE <> "FDF" THEN
       CALL PutMSG("XXXXM033")
       GOTO FDFOpenEND
    END IF
    IF FDFHdr.VER <> CHR$(0) THEN
       CALL PutMSG("XXXXM036")
       GOTO FDFOpenEND
    END IF
    OKAY = -1
    IF OPT$ <> "INIT" GOTO FDFOpenEND
    FOR X = 1 TO (FDFHdr.FAMDSN + FDFHdr.OTHERDSN)
        GET #1, , FDFDSRec
        Y = 0: FOUND = 0
        WHILE (Y < ENV.DSNCNT) AND (NOT FOUND)
           Y = Y + 1: FOUND = (FTYPE$(Y) = FDFDSRec.TYPE)
        WEND
        IF FOUND THEN FT$(Y) = FDFDSRec.DSNNAME
    NEXT X
    FSUMAX = FDFHdr.SETUPS
    REDIM FDFSURBA(FSUMAX), FSETUP$(FSUMAX), FSUCHN(FSUMAX) AS ChainPTRS
    CALL ChnINIT(FSUCHN(), 1, FSUMAX)
    FSUTable.FT = 1: FSUTable.LT = FSUMAX: FSUTable.CL = 0
    FSUTable.SIZE = FSUMAX: FSUTable.FF = 0
    X = 0: RBA = SEEK(1)
    WHILE X < FDFHdr.SETUPS
       X = X + 1
       FDFSURBA(X) = RBA
       GET #1, RBA, FDFSetup
       FSETUP$(X) = SPACE$(FSUTable.XL)
       MID$(FSETUP$(X), 1, 8) = FDFSetup.NAME
       MID$(FSETUP$(X), 9, 255) = FDFSetup.DESC
       RBA = RBA + LEN(FDFSetup) + FDFHdr.FAMDSN * 12
    WEND
    CLOSE #1
    CALL FDFGetSet("DEFAULT")
    GOTO FDFOpenEND
NEWFDF:
    ON LOCAL ERROR GOTO FDFOpenERR
    OPEN SYSVAR.FILENAME FOR BINARY ACCESS READ WRITE AS #1
    OKAY = -1
    GOTO FDFOpenEND
FDFOpenERR:
    CALL ERRMessage("XXXXM030", ERR, ERL)
    RESUME FDFOpenEND
FDFOpenEND:
    IF NOT OKAY THEN CALL FamCLOSE ELSE ENV.FDF = CONFIG.FDF
    ON LOCAL ERROR GOTO 0
END SUB

SUB FDFSelect
   CALL FamCLOSE
   CALL WinCLR
   CALL FmtFIND("XXXXS020")
   CALL WinFORMAT(0)
   CALL FmtFIND("XXXXS097")
   CALL OptDISPLAY(0)
   IF FSUTable.FMTNAME <> "XXXXS021" THEN
      FSUTable.FMTNAME = "XXXXS021": FSUTable.FMTNUM = 0
   END IF
   CALL TableOPEN(FSUTable)
   CALL TableSHOW(FSUTable, FSETUP$(), FSUCHN())
   CALL TableSELECT(FSUTable, FSETUP$(), FSUCHN())
   IF A$ <> CHR$(27) THEN
      X$ = LEFT$(FSETUP$(FSUTable.CT), 8)
      CALL FDFGetSet(X$)
   END IF
   CALL WinCLR
END SUB

FUNCTION FDSTAT$ (ST$)
   X$ = " ": X = ASC(FF1REC.STATUS)
   SELECT CASE ST$
      CASE "FID": SELECT CASE (X AND 3)
                    CASE 1: LSET X$ = "*"
                    CASE 2: LSET X$ = "?"
                  END SELECT
      CASE "MID": SELECT CASE (X AND 12)
                    CASE 4: LSET X$ = "*"
                    CASE 8: LSET X$ = "?"
                  END SELECT
      CASE "BD": SELECT CASE (X AND 48)
                   CASE 16: LSET X$ = "?"
                   CASE 32: LSET X$ = "!"
                 END SELECT
      CASE "DD": SELECT CASE (X AND 192)
                   CASE 64: LSET X$ = "?"
                   CASE 128: LSET X$ = "!"
                 END SELECT
      CASE "SNU": IF ((ASC(FF1REC.STATUS2) AND 128) = 128) THEN LSET X$ = SNGLKEY$(1)
   END SELECT
   FDSTAT$ = X$
END FUNCTION

SUB FF1GetNew
   MAXID& = FFRBA&(FF1HDR.MAXID, 1) + 1
   MAXID = FFRNUM(MAXID&, 1)
   IF MAXID > 0 THEN
      FF1HDR.MAXID = MAXID
      RBA& = 100 * MAXID& + 1
      NF$ = STRING$(50, 0) + SPACE$(50)
      PUT #1, RBA&, NF$
      GET #1, RBA&, FF1REC
      FF1REC.RTYPE = CHR$(1)
      FF1REC.RID = MAXID
      FF1REC.SEX = " "
      FF1HDR.RevLVL = FF1HDR.RevLVL OR 256
      PUT #1, 1, FF1HDR
      PUT #1, RBA&, FF1REC
      OKAY = -1
   END IF
END SUB

REM $DYNAMIC
SUB FF1GetRec (RNUM)
   OKAY = 0
   IF ENV.FFOPEN THEN
      MAXID& = FFRBA&(FF1HDR.MAXID, 1)
      RBA& = FFRBA&(RNUM, 1)
      IF (RBA& < 1) OR (RBA& > MAXID&) THEN
         SYSVAR.MVAR = LTRIM$(STR$(RBA&))
         CALL PutMSG("XXXXM050")
         ELSE
         RBA& = RBA& * 100 + 1
         GET #1, RBA&, FF1REC
         OKAY = -1
      END IF
   END IF
END SUB

SUB FF1Open (OPT$)
    OKAY = 0: D$ = FDFSetup.NFDRV: A$ = ""
    QUIET = (RIGHT$(OPT$, 1) = "Q"): X = LEN(RTRIM$(OPT$)) + QUIET
    XOPT$ = MID$(OPT$, 1, X)
    IF (("A" <= D$) AND (D$ <= "Z")) AND (D$ <> FDFSetup.FAMDRV) THEN
       DSNAME$ = RTRIM$(D$ + ":" + FT$(1))
       ELSE
       DSNAME$ = RTRIM$(RTRIM$(FAMFPFX$) + FT$(1))
    END IF
    LSET SYSVAR.FILENAME = DSNAME$
    ON LOCAL ERROR GOTO NOFF1
CHKFF1:
    OPEN DSNAME$ FOR INPUT AS #1
    CLOSE #1
    IF XOPT$ = "INIT" THEN
       IF NOT QUIET THEN CALL PutMSG("XXXXM051")
       IF A$ <> SNGLKEY$(1) THEN GOTO FF1OpenEND
       GOTO OLDFF1
    END IF
    IF XOPT$ = "NEW" THEN
       IF NOT QUIET THEN CALL PutMSG("XXXXM002")
       IF A$ <> SNGLKEY$(1) THEN GOTO FF1OpenEND
       KILL SYSVAR.FILENAME
       GOTO NEWFF1
    END IF
    GOTO OLDFF1
NOFF1:
    X = ERR
    IF (XOPT$ = "NEW") THEN RESUME NEWFF1
    IF (XOPT$ = "INIT") THEN OKAY = -1: RESUME FF1OpenEND
    IF QUIET THEN RESUME FF1OpenEND
    CALL ERRMessage("XXXXM030", X, 0)
    IF (A$ = CHR$(27)) OR (X <> 71) THEN RESUME FF1OpenEND
    RESUME CHKFF1
OLDFF1:
    ON LOCAL ERROR GOTO FF1OpenERR
    SELECT CASE XOPT$
       CASE "INIT": OPEN SYSVAR.FILENAME FOR BINARY ACCESS READ AS #1
       CASE "READ": OPEN SYSVAR.FILENAME FOR BINARY ACCESS READ AS #1
       CASE "WRITE": OPEN SYSVAR.FILENAME FOR BINARY ACCESS READ WRITE AS #1
       CASE ELSE: GOTO FF1OpenEND
    END SELECT
    GET #1, , FF1HDR
    LSET SYSVAR.MVAR = "NAME"
    IF FF1HDR.FTYPE <> "N" THEN
       IF NOT QUIET THEN CALL PutMSG("XXXXM033")
       GOTO FF1OpenEND
    END IF
    IF ((FF1HDR.RevLVL AND 254) <> 0) THEN
       CALL PutMSG("XXXXM036")
       IF A$ <> SNGLKEY$(1) THEN GOTO FF1OpenEND
       IF XOPT$ = "WRITE" THEN
          FF1HDR.RevLVL = (FF1HDR.RevLVL AND 3) OR 256
          ELSE
          FF1HDR.RevLVL = (FF1HDR.RevLVL AND 3)
       END IF
    END IF
    OKAY = -1
    GOTO FF1OpenEND
NEWFF1:
    CALL PutMSG("XXXXM056")
    IF A$ = CHR$(27) THEN GOTO FF1OpenEND
    ENV.LPNSupport = (A$ = SNGLKEY$(1))
    ON LOCAL ERROR GOTO FF1OpenERR
    OPEN SYSVAR.FILENAME FOR BINARY ACCESS READ WRITE AS #1
    X$ = STRING$(100, 0): PUT #1, 1, X$: GET #1, 1, FF1HDR
    FF1HDR.FTYPE = "N"
    FF1HDR.RevLVL = -ENV.LPNSupport
    FF1HDR.REORGY = DATEY(DATE$, 1)
    FF1HDR.REORGMD = DATEMD(DATE$, 1)
    FF1HDR.REORGTM = TimeHHMM(TIME$)
    FF1HDR.UPDTY = FF1HDR.REORGY
    FF1HDR.UPDTMD = FF1HDR.REORGMD
    PUT #1, 1, FF1HDR
    OKAY = -1
    GOTO FF1OpenEND
FF1OpenERR:
    IF NOT QUIET THEN CALL ERRMessage("XXXXM030", ERR, ERL)
    RESUME FF1OpenEND
FF1OpenEND:
    ENV.LPNSupport = (FF1HDR.RevLVL > 0)
END SUB

REM $STATIC
SUB FF1PutREC
    HIRBA& = FFRBA&(FF1HDR.MAXID, 1)
    RBA& = FFRBA&(FF1REC.RID, 1)
    IF (RBA& > 0) AND (RBA& <= HIRBA&) THEN
       RBA& = RBA& * 100 + 1
       X = ASC(FF1REC.STATUS2)
       IF ((X AND 96) = 0) THEN FF1HDR.CHGCNT2 = FF1HDR.CHGCNT2 + 1
       IF XCHANGED THEN               ' Turn on record Change Bit
          X = X OR 32: XCHANGED = 0
          ELSE
          IF ((X AND 64) = 0) THEN FF1HDR.CHGCNT1 = FF1HDR.CHGCNT1 + 1
          X = X OR 64
       END IF
       FF1REC.STATUS2 = CHR$(X)
       PUT #1, RBA&, FF1REC
       IF FF1HDR.RevLVL < 256 THEN
          FF1HDR.RevLVL = FF1HDR.RevLVL OR 256
          PUT #1, 1, FF1HDR
       END IF
       OKAY = -1: FFCHANGED = -1
       ELSE
       CALL ErrBEEP(0): OKAY = 0
    END IF
END SUB

SUB FF2DelRec (RNUM)
    HIRBA& = FFRBA&(FF2HDR.HIREC, 1)
    RNO& = FFRBA&(RNUM, 1)
    IF (RNO& > 0) AND (RNO& <= HIRBA&) THEN
       AFMT = FF2HDR.RevLVL AND 255
       IF AFMT = 0 THEN RL = 108 ELSE RL = 152
       RBA& = RNO& * RL + 1
       GET #2, RBA&, FF2PFX
       IF FF2PFX.RTYPE <> CHR$(255) THEN
          FF2PFX.RTYPE = CHR$(255)
          FF2PFX.NXT = FF2HDR.FREEBGN
          PUT #2, RBA&, FF2PFX
          FF2HDR.FREEBGN = RNUM
          FF2HDR.FREECNT = FF2HDR.FREECNT + 1
          IF FF2HDR.RevLVL < 256 THEN
             FF2HDR.RevLVL = 256 + AFMT
             PUT #2, 1, FF2HDR
          END IF
          OKAY = -1: XCHANGED = -1: FFCHANGED = -1
       END IF
       ELSE
       OKAY = 0
    END IF
END SUB

SUB FF2GetRec (RNUM)
   OKAY = 0
   IF NOT ENV.FFOPEN THEN GOTO FF2GetEND
   AFMT = FF2HDR.RevLVL AND 255
   IF AFMT = 0 THEN RL = 108 ELSE RL = 152
   HIREC& = FFRBA&(FF2HDR.HIREC, 1)
   IF RNUM = 0 THEN GOTO FF2GetNew
   RNO& = FFRBA&(RNUM, 1)
   IF (RNO& < 1) OR (RNO& > HIREC&) THEN
      SYSVAR.MVAR = LTRIM$(STR$(RNO&))
      CALL PutMSG("XXXXM052")
      GOTO FF2GetEND
   END IF
   RBA& = RNO& * RL + 1
   GET #2, RBA&, FF2PFX
   SELECT CASE AFMT
      CASE 0: GET #2, , FF2SRec
      CASE ELSE: GET #2, , FF2LRec
   END SELECT
   OKAY = -1
   GOTO FF2GetEND
FF2GetNew:
   IF FF2HDR.FREECNT > 0 THEN
      RBA& = FFRBA&(FF2HDR.FREEBGN, RL) + 1
      IF RBA& > 1 THEN
         GET #2, RBA&, FF2PFX
         FF2HDR.FREEBGN = FF2PFX.NXT
         FF2HDR.FREECNT = FF2HDR.FREECNT - 1
         ELSE
         RBA& = 0
      END IF
      ELSE
      RBA& = 0
   END IF
   IF RBA& = 0 THEN
      HIREC& = HIREC& + 1
      HIREC = FFRNUM(HIREC&, 1)
      IF HIREC = 0 THEN GOTO FF2GetEND
      FF2HDR.HIREC = HIREC
      RBA& = HIREC& * RL + 1
   END IF
   RNO = FFRNUM(RBA& - 1, RL)
   FF2HDR.UPDTY = DATEY(DATE$, 1): FF2HDR.UPDTMD = DATEMD(DATE$, 1)
   FF2HDR.RevLVL = 256 + AFMT: Z = 18 - 6 * (AFMT = 0)
   AF$ = STRING$(Z, 0) + SPACE$(RL - Z)
   MID$(AF$, 1, 3) = CHR$(2) + MKI$(RNO)
   PUT #2, RBA&, AF$
   GET #2, RBA&, FF2PFX
   FF2PFX.RNO = RNO
   SELECT CASE AFMT
      CASE 0: GET #2, , FF2SRec
      CASE ELSE: GET #2, , FF2LRec
   END SELECT
   PUT #2, 1, FF2HDR
   OKAY = -1
FF2GetEND:
END SUB

REM $DYNAMIC
SUB FF2Open (OPT$)
    OKAY = 0: D$ = FDFSetup.NFDRV: A$ = ""
    QUIET = (RIGHT$(OPT$, 1) = "Q"): X = LEN(RTRIM$(OPT$)) + QUIET
    XOPT$ = MID$(OPT$, 1, X)
    IF (("A" <= D$) AND (D$ <= "Z")) AND (D$ <> FDFSetup.FAMDRV) THEN
       DSNAME$ = D$ + ":" + FT$(2)
       ELSE
       DSNAME$ = FAMFPFX$ + FT$(2)
    END IF
    LSET SYSVAR.FILENAME = DSNAME$
    ON LOCAL ERROR GOTO NOFF2
CHKFF2:
    OPEN DSNAME$ FOR INPUT AS #2
    CLOSE #2
    IF XOPT$ = "INIT" THEN
       IF NOT QUIET THEN CALL PutMSG("XXXXM051")
       IF A$ <> SNGLKEY$(1) THEN GOTO FF2OpenEND
       GOTO OLDFF2
    END IF
    IF XOPT$ = "NEW" THEN
       IF NOT QUIET THEN CALL PutMSG("XXXXM002")
       IF A$ <> SNGLKEY$(1) THEN GOTO FF2OpenEND
       KILL SYSVAR.FILENAME
       GOTO NEWFF2
    END IF
    GOTO OLDFF2
NOFF2:
    X = ERR
    IF (XOPT$ = "NEW") THEN RESUME NEWFF2
    IF (XOPT$ = "INIT") THEN OKAY = -1: RESUME FF2OpenEND
    IF NOT QUIET THEN CALL ERRMessage("XXXXM030", X, 0)
    IF (A$ = CHR$(27)) OR (X <> 71) THEN RESUME FF2OpenEND
    RESUME CHKFF2
OLDFF2:
    ON LOCAL ERROR GOTO FF2OpenERR
    SELECT CASE XOPT$
       CASE "INIT": OPEN SYSVAR.FILENAME FOR BINARY ACCESS READ AS #2
       CASE "READ": OPEN SYSVAR.FILENAME FOR BINARY ACCESS READ AS #2
       CASE "WRITE": OPEN SYSVAR.FILENAME FOR BINARY ACCESS READ WRITE AS #2
       CASE ELSE: GOTO FF2OpenEND
    END SELECT
    GET #2, , FF2HDR
    LSET SYSVAR.MVAR = "ADDRESS"
    IF FF2HDR.FTYPE <> "A" THEN
       CALL PutMSG("XXXXM033")
       IF A$ <> SNGLKEY$(1) THEN GOTO FF2OpenEND
       FF2HDR.FTYPE = "A"
    END IF
    AFMT = FF2HDR.RevLVL AND 255
    IF (AFMT > 1) THEN
       CALL PutMSG("XXXXM036")
       IF A$ <> SNGLKEY$(1) THEN GOTO FF2OpenEND
       IF XOPT$ = "WRITE" THEN
          FF2HDR.RevLVL = AFMT + 256
          ELSE
          FF2HDR.RevLVL = AFMT
       END IF
    END IF
    IF (FF2HDR.REORGY <> FF1HDR.REORGY) OR (FF2HDR.REORGMD <> FF1HDR.REORGMD) OR (FF2HDR.REORGTM <> FF1HDR.REORGTM) THEN
       CALL PutMSG("XXXXM034")
       IF A$ <> SNGLKEY$(1) THEN GOTO FF2OpenEND
       FF2HDR.REORGY = FF1HDR.REORGY: FF2HDR.REORGMD = FF1HDR.REORGMD
       FF2HDR.REORGTM = FF1HDR.REORGTM
       IF XOPT$ = "WRITE" THEN FF2HDR.RevLVL = FF2HDR.RevLVL OR 256
    END IF
    OKAY = -1
    GOTO FF2OpenEND
NEWFF2:
    CALL PutMSG("XXXXM053")
    IF A$ = CHR$(27) THEN GOTO FF2OpenEND
    SELECT CASE A$
       CASE "1": AFMT = 0
       CASE "2": AFMT = 1
       CASE ELSE: CALL ErrBEEP(0): GOTO NEWFF2
    END SELECT
    ON LOCAL ERROR GOTO FF2OpenERR
    OPEN SYSVAR.FILENAME FOR BINARY ACCESS READ WRITE AS #2
    FF2HDR.FTYPE = "A"
    FF2HDR.RevLVL = AFMT
    FF2HDR.HIREC = 0
    FF2HDR.REORGY = FF1HDR.REORGY
    FF2HDR.REORGMD = FF1HDR.REORGMD
    FF2HDR.REORGTM = FF1HDR.REORGTM
    FF2HDR.UPDTY = FF1HDR.UPDTY
    FF2HDR.UPDTMD = FF1HDR.UPDTMD
    FF2HDR.FREEBGN = 0: FF2HDR.FREECNT = 0
    PUT #2, 1, FF2HDR
    X$ = STRING$(89 - 44 * (AFMT = 1), 0)
    PUT #2, , X$
    OKAY = -1
    GOTO FF2OpenEND
FF2OpenERR:
    CALL ERRMessage("XXXXM030", ERR, ERL)
    RESUME FF2OpenEND
FF2OpenEND:
END SUB

REM $STATIC
SUB FF2PutRec
    HIRBA& = FFRBA&(FF2HDR.HIREC, 1)
    RBA& = FFRBA&(FF2PFX.RNO, 1)
    IF (RBA& > 0) AND (RBA& <= HIRBA&) THEN
       AFMT = FF2HDR.RevLVL AND 255
       IF AFMT = 0 THEN RL = 108 ELSE RL = 152
       RBA& = RBA& * RL + 1
       PUT #2, RBA&, FF2PFX
       XCHANGED = -1
       SELECT CASE AFMT
          CASE 0: PUT #2, , FF2SRec
          CASE 1: PUT #2, , FF2LRec
       END SELECT
       IF FF2HDR.RevLVL < 256 THEN
          FF2HDR.RevLVL = 256 + AFMT
          PUT #2, 1, FF2HDR
       END IF
       OKAY = -1: FFCHANGED = -1: XCHANGED = -1
       ELSE
       CALL ErrBEEP(0): OKAY = 0
    END IF
END SUB

SUB FF3DelRec (RNUM)
    HIRBA& = FFRBA&(FF3HDR.HIREC, 1)
    RBA& = FFRBA&(RNUM, 1)
    IF (RBA& > 0) AND (RBA& <= HIRBA&) THEN
       RBA& = RBA& * 50 + 1
       GET #3, RBA&, FF3PFX
       IF FF3PFX.RTYPE <> CHR$(255) THEN
          FF3PFX.RTYPE = CHR$(255)
          PUT #3, RBA&, FF3PFX
          GET #3, RBA& + 3, FF3PNTRS
          FF3PNTRS.NXT = FF3HDR.FREEBGN
          PUT #3, RBA& + 3, FF3PNTRS
          FF3HDR.FREEBGN = RNUM
          FF3HDR.FREECNT = FFRNUM(FFRBA&(FF3HDR.FREECNT, 1) + 1, 1)
          IF FF3HDR.RevLVL < 256 THEN
             FF3HDR.RevLVL = FF3HDR.RevLVL OR 256
             PUT #3, 1, FF3HDR
          END IF
          OKAY = -1: XCHANGED = -1: FFCHANGED = -1
       END IF
       ELSE
       OKAY = 0
    END IF
END SUB

SUB FF3GetRec (RNUM)
   OKAY = 0
   IF NOT ENV.FFOPEN THEN GOTO FF3GetEND
   HIREC& = FFRBA&(FF3HDR.HIREC, 1)
   IF RNUM = 0 THEN GOTO FF3GetNew
   RBA& = FFRBA&(RNUM, 1)
   LSET SYSVAR.MVAR = "MISCINFO"
   IF (RBA& < 1) OR (RBA& > HIREC&) THEN
      SYSVAR.MVAR = LTRIM$(STR$(RBA&))
      CALL PutMSG("XXXXM054")
      GOTO FF3GetEND
   END IF
   RBA& = RBA& * 50 + 1
FF3GetIT:
   GET #3, RBA&, FF3PFX
   SELECT CASE FF3PFX.RTYPE
      CASE CHR$(3): GET #3, , FF3SPOUSE
      CASE CHR$(4): GET #3, , FF3PLACE
                SELECT CASE FF3PLACE.SRTYPE
                   CASE CHR$(1): LSET BPLACE = FF3PLACE.PLACE1
                                 LSET DPLACE = FF3PLACE.PLACE2
                   CASE CHR$(3): LSET MPLACE = FF3PLACE.PLACE1
                                 LSET DVPLACE = FF3PLACE.PLACE2
                END SELECT
      CASE CHR$(5): GET #3, , FF3COMMENT
      CASE "3": GET #3, , FF3SPOUSE              ' Old SPOUSE Records
      CASE "4": GET #3, , FF3PLACE               ' Old PLACE Records
                SELECT CASE FF3PLACE.SRTYPE
                   CASE CHR$(1): LSET BPLACE = FF3PLACE.PLACE1
                                 LSET DPLACE = FF3PLACE.PLACE2
                   CASE CHR$(3): LSET MPLACE = FF3PLACE.PLACE1
                                 LSET DVPLACE = FF3PLACE.PLACE2
                END SELECT
      CASE "F": GET #3, , FF3COMMENT             ' Old COMMENT Records
      CASE CHR$(6): GET #3, , FF3EVENT           ' New EVENT Records
      CASE CHR$(11): GET #3, , FF3EVENT          ' Old EVENT Records
      CASE CHR$(12): GET #3, , FF3PLACE2
                SELECT CASE FF3PLACE2.PTYPE
                   CASE CHR$(1): LSET BPLACE = FF3PLACE2.PLACE
                   CASE CHR$(2): LSET DPLACE = FF3PLACE2.PLACE
                   CASE CHR$(3): LSET MPLACE = FF3PLACE2.PLACE
                   CASE CHR$(4): LSET DVPLACE = FF3PLACE2.PLACE
                END SELECT
      CASE ELSE: GET #3, , FF3PNTRS
                 SELECT CASE FF3PFX.RTYPE
                    CASE CHR$(7): GET #3, , FF3EDUC
                    CASE CHR$(8): GET #3, , FF3WORK
                    CASE CHR$(9): GET #3, , FF3MIL
                    CASE CHR$(10): GET #3, , FF3HEALTH
                    CASE "7": GET #3, , FF3EDUC    ' Old EDUC Records
                    CASE "8": GET #3, , FF3WORK    ' Old WORK Records
                    CASE "9": GET #3, , FF3MIL     ' Old MIL Records
                    CASE "A": GET #3, , FF3HEALTH  ' Old HLTH Records
                    CASE CHR$(255):
                    CASE ELSE: SYSVAR.MVAR = LTRIM$(STR$(ASC(FF3PFX.RTYPE)))
                               'CALL PutMSG("XXXXM055")
                               GOTO FF3GetEND
                 END SELECT
   END SELECT
   OKAY = -1
   GOTO FF3GetEND
FF3GetNew:
   IF FF3HDR.FREECNT <> 0 THEN
      RNO = FF3HDR.FREEBGN
      RBA& = FFRBA&(RNO, 50) + 1
      IF RBA& > 1 THEN
         GET #3, RBA& + 3, FF3PNTRS
         FF3HDR.FREEBGN = FF3PNTRS.NXT
         FF3HDR.FREECNT = FFRNUM(FFRBA&(FF3HDR.FREECNT, 1) - 1, 1)
      END IF
      ELSE
      RBA& = 0
   END IF
   IF RBA& = 0 THEN
      HIREC& = HIREC& + 1: HIREC = FFRNUM(HIREC&, 1)
      IF HIREC = 0 THEN GOTO FF3GetEND
      FF3HDR.HIREC = HIREC
      RNO = HIREC
      RBA& = HIREC& * 50 + 1
   END IF
   IF RBA& > 1 THEN
      FF3HDR.RevLVL = FF3HDR.RevLVL OR 256
      PUT #3, 1, FF3HDR
      SELECT CASE FF3PFX.RTYPE
         CASE CHR$(3): ZL = 23: XL = 24
         CASE CHR$(4): ZL = 3: XL = 44
         CASE CHR$(5): ZL = 9: XL = 38
         CASE CHR$(6): ZL = 16: XL = 31
         CASE CHR$(11): ZL = 16: XL = 31  ' Old Event Records
         CASE CHR$(12): ZL = 6: XL = 41   ' New PLACE Name Record
         CASE ELSE: ZL = 17: XL = 30
      END SELECT
      AF$ = FF3PFX.RTYPE + MKI$(RNO) + STRING$(ZL, 0) + SPACE$(XL)
      PUT #3, RBA&, AF$
      GOTO FF3GetIT
   END IF
FF3GetEND:
END SUB

REM $DYNAMIC
DEFSNG A-Z
SUB FF3Open (OPT$)
    OKAY = 0: D$ = FDFSetup.NFDRV: A$ = ""
    QUIET = (RIGHT$(OPT$, 1) = "Q"): X = LEN(RTRIM$(OPT$)) + QUIET
    XOPT$ = MID$(OPT$, 1, X)
    IF (("A" <= D$) AND (D$ <= "Z")) AND (D$ <> FDFSetup.FAMDRV) THEN
       DSNAME$ = D$ + ":" + FT$(3)
       ELSE
       DSNAME$ = FAMFPFX$ + FT$(3)
    END IF
    LSET SYSVAR.FILENAME = DSNAME$
    ON LOCAL ERROR GOTO NOFF3
CHKFF3:
    OPEN DSNAME$ FOR INPUT AS #3
    CLOSE #3
    IF XOPT$ = "INIT" THEN
       IF NOT QUIET THEN CALL PutMSG("XXXXM051")
       IF A$ <> SNGLKEY$(1) THEN GOTO FF3OpenEND
       GOTO OLDFF3
    END IF
    IF XOPT$ = "NEW" THEN
       IF NOT QUIET THEN CALL PutMSG("XXXXM002")
       IF A$ <> SNGLKEY$(1) THEN GOTO FF3OpenEND
       KILL SYSVAR.FILENAME
       GOTO NEWFF3
    END IF
    GOTO OLDFF3
NOFF3:
    X = ERR
    IF XOPT$ = "NEW" THEN RESUME NEWFF3
    IF XOPT$ = "INIT" THEN OKAY = -1: RESUME FF3OpenEND
    IF NOT QUIET THEN CALL ERRMessage("XXXXM030", X, 0)
    IF (A$ = CHR$(27)) OR (X <> 71) THEN RESUME FF3OpenEND
    RESUME CHKFF3
OLDFF3:
    ON LOCAL ERROR GOTO FF3OpenERR
    SELECT CASE XOPT$
       CASE "INIT": OPEN SYSVAR.FILENAME FOR BINARY ACCESS READ AS #3
       CASE "READ": OPEN SYSVAR.FILENAME FOR BINARY ACCESS READ AS #3
       CASE "WRITE": OPEN SYSVAR.FILENAME FOR BINARY ACCESS READ WRITE AS #3
       CASE ELSE: GOTO FF3OpenEND
    END SELECT
    GET #3, 1, FF3HDR
    LSET SYSVAR.MVAR = "MISCINFO"
    IF FF3HDR.FTYPE <> "M" THEN
       IF NOT QUIET THEN CALL PutMSG("XXXXM033")
       IF A$ <> SNGLKEY$(1) THEN GOTO FF3OpenEND
       FF3HDR.FTYPE = "M"
       IF XOPT$ = "WRITE" THEN FF3HDR.RevLVL = FF3HDR.RevLVL OR 256
    END IF
    IF ((FF3HDR.RevLVL AND 255) <> 0) THEN
       IF NOT QUIET THEN CALL PutMSG("XXXXM036")
       IF A$ <> SNGLKEY$(1) THEN GOTO FF3OpenEND
       FF3HDR.RevLVL = 0
       IF XOPT$ = "WRITE" THEN FF3HDR.RevLVL = FF3HDR.RevLVL OR 256
    END IF
    IF (FF3HDR.REORGY <> FF1HDR.REORGY) OR (FF3HDR.REORGMD <> FF1HDR.REORGMD) OR (FF3HDR.REORGTM <> FF1HDR.REORGTM) THEN
       IF NOT QUIET THEN CALL PutMSG("XXXXM034")
       IF A$ <> SNGLKEY$(1) THEN GOTO FF3OpenEND
       FF3HDR.REORGY = FF1HDR.REORGY: FF3HDR.REORGMD = FF1HDR.REORGMD
       FF3HDR.REORGTM = FF1HDR.REORGTM
       IF XOPT$ = "WRITE" THEN FF3HDR.RevLVL = FF3HDR.RevLVL OR 256
    END IF
    OKAY = -1
    GOTO FF3OpenEND
NEWFF3:
    ON LOCAL ERROR GOTO FF3OpenERR
    OPEN SYSVAR.FILENAME FOR BINARY ACCESS READ WRITE AS #3
    FF3HDR.FTYPE = "M"
    FF3HDR.RevLVL = 0
    FF3HDR.HIREC = 0
    FF3HDR.REORGY = FF1HDR.REORGY
    FF3HDR.REORGMD = FF1HDR.REORGMD
    FF3HDR.REORGTM = FF1HDR.REORGTM
    FF3HDR.UPDTY = FF1HDR.UPDTY
    FF3HDR.UPDTMD = FF1HDR.UPDTMD
    FF3HDR.FREEBGN = 0: FF3HDR.FREECNT = 0
    'FF3HDR.HDRFILL = STRING$(LEN(FF3HDR.HDRFILL), 0)
    PUT #3, 1, FF3HDR
    OKAY = -1
    GOTO FF3OpenEND
FF3OpenERR:
    IF NOT QUIET THEN CALL ERRMessage("XXXXM030", ERR, ERL)
    RESUME FF3OpenEND
FF3OpenEND:
END SUB

REM $STATIC
DEFINT A-Z
SUB FF3PutRec
    OKAY = 0
    HIRBA& = FFRBA&(FF3HDR.HIREC, 1)
    RBA& = FFRBA&(FF3PFX.RNO, 1)
    IF (RBA& > 0) AND (RBA& <= HIRBA&) THEN
       RBA& = RBA& * 50 + 1
       PUT #3, RBA&, FF3PFX
       SELECT CASE FF3PFX.RTYPE
          CASE CHR$(3): PUT #3, , FF3SPOUSE
          CASE CHR$(4): PUT #3, , FF3PLACE
          CASE CHR$(5): PUT #3, , FF3COMMENT
          CASE CHR$(6): PUT #3, , FF3EVENT
          CASE CHR$(11): PUT #3, , FF3EVENT     ' Old EVENT Records
          CASE CHR$(12): PUT #3, , FF3PLACE2
          CASE ELSE: PUT #3, , FF3PNTRS
                     SELECT CASE FF3PFX.RTYPE
                        CASE CHR$(7): PUT #3, , FF3EDUC
                        CASE CHR$(8): PUT #3, , FF3WORK
                        CASE CHR$(9): PUT #3, , FF3MIL
                        CASE CHR$(10): PUT #3, , FF3HEALTH
                        CASE CHR$(255):
                        CASE ELSE: GOTO FF3PutEND
                     END SELECT
       END SELECT
       IF FF3HDR.RevLVL < 256 THEN
          FF3HDR.RevLVL = FF3HDR.RevLVL OR 256
          PUT #3, 1, FF3HDR
       END IF
       OKAY = -1: XCHANGED = -1: FFCHANGED = -1
    END IF
FF3PutEND:
   IF NOT OKAY THEN CALL ErrBEEP(0)
END SUB

REM $DYNAMIC
FUNCTION FFRBA& (RNUM, RLTH)
   X& = RNUM
   ' New Routine - After May 20, 1997
     IF X& < 0 THEN X& = 65535 + X&: IF X& < 0 THEN X& = 0
   ' Old Routine - Prior to May 21, 1997
   ' IF X& < 0 THEN X& = 32766 - X&
   FFRBA& = X& * RLTH
END FUNCTION

FUNCTION FFRNUM (RBA&, RLTH)
   IF RLTH > 0 THEN X& = RBA& / RLTH ELSE X& = RBA&
   ' New Routine - After May 20,1997
     IF X& > 32767 THEN X& = X& - 65535: IF X& > 0 THEN X& = 0
     IF X& < -32767 THEN X& = 0
   ' Old Routine - prior to May 21, 1997
   ' IF X& > 32766 THEN X& = 32766 - X&
   ' IF X& < -32766 THEN X& = 0
   FFRNUM = X&
END FUNCTION

FUNCTION FFSTAT$ (ST$, S$)
   X = ASC(FF1REC.STATUS)
   SELECT CASE ST$
      CASE "FID": X = (X AND (255 - 3))
                  SELECT CASE S$
                    CASE "*": X = X OR 1
                    CASE "?": X = X OR 2
                  END SELECT
      CASE "MID": X = (X AND (255 - 12))
                  SELECT CASE S$
                    CASE "*": X = X OR 4
                    CASE "?": X = X OR 8
                  END SELECT
      CASE "BD": X = (X AND (255 - 48))
                 SELECT CASE S$
                   CASE "?": X = X OR 16
                   CASE "!": X = X OR 32
                 END SELECT
      CASE "DD": X = (X AND (255 - 192))
                 SELECT CASE S$
                   CASE "?": X = X OR 64
                   CASE "!": X = X OR 128
                 END SELECT
      CASE "SNU": X = ASC(FF1REC.STATUS2) AND 127: IF S$ = "Y" THEN X = X OR 128
   END SELECT
   FFSTAT$ = CHR$(X)
END FUNCTION

REM $STATIC
FUNCTION FMTNAME$ (XNFMT)
   NFMT = XNFMT: IF NFMT = 0 THEN NFMT = RPTOPTION.NAMEFMT + 1
   LASTFIRST = ((NFMT = 2) OR (NFMT = 4))
   UPCASE = ((NFMT = 3) OR (NFMT = 4))
   GIVEN$ = LTRIM$(RTRIM$(FF1REC.GIVEN))
   SURNM$ = LTRIM$(RTRIM$(FF1REC.SURNM))
   IF ENV.HUSBSURNM AND (FF1REC.SEX = FGENDR$(2)) THEN
      IF (FF1REC.SPOUSE <> 0) AND ((ASC(FF1REC.STATUS2) AND 128) = 0) THEN
         SPID = 0
         CALL FF3GetRec(FF1REC.SPOUSE)
         IF FF1REC.RID = FF3SPOUSE.SP1ID THEN
            SPID = FF3SPOUSE.SP2ID
            ELSE
            SPID = FF3SPOUSE.SP1ID
         END IF
         IF SPID > 0 THEN
            RID = FF1REC.RID: CALL FF1GetRec(SPID)
            SURNM$ = LTRIM$(RTRIM$(FF1REC.SURNM))
            CALL FF1GetRec(RID)
            GIVEN$ = GIVEN$ + " " + LTRIM$(RTRIM$(FF1REC.SURNM))
         END IF
      END IF
   END IF
   IF UPCASE THEN SURNM$ = UCX$(SURNM$)
   IF LASTFIRST THEN
     FMTNAME$ = SURNM$ + ", " + GIVEN$
     ELSE
     FMTNAME$ = GIVEN$ + " " + SURNM$
   END IF
END FUNCTION

SUB GetBDEvents (BD$, BPL$, DD$, DPL$, AGE$)
     ' Get Birth(Baptism)/Death(Burial)/Age Fields
     D1ST$ = FDSTAT$("BD"): D2ST$ = FDSTAT$("DD")
     LSET BPLACE = "": LSET DPLACE = ""
     IF BPL$ <> "" THEN
        IF FF1REC.BLOC <> 0 THEN
           CALL FF3GetRec(FF1REC.BLOC)
           IF FF3PFX.RTYPE = CHR$(12) THEN
              IF FF3PLACE2.PNEXT <> 0 THEN
                 CALL FF3GetRec(FF3PLACE2.PNEXT)
              END IF
           END IF
        END IF
     END IF
     EVNTPTR = FF1REC.EVENT
     WHILE ((FF1REC.BY = 0) OR (FF1REC.DY = 0)) AND (EVNTPTR <> 0)
        CALL FF3GetRec(EVNTPTR)
        IF FF1REC.BY = 0 THEN
           SELECT CASE FF3EVENT.CODE
              CASE "BIRTH   ": FF1REC.BY = FF3EVENT.D1Y
                               FF1REC.BMD = FF3EVENT.D1MD
                               LSET BPLACE = FF3EVENT.PLACE
                               D1ST$ = FF3EVENT.DSTAT
              CASE "BAPTISM ": FF1REC.BY = FF3EVENT.D1Y
                               FF1REC.BMD = FF3EVENT.D1MD
                               LSET BPLACE = FF3EVENT.PLACE
                               D1ST$ = "+"
           END SELECT
        END IF
        IF FF1REC.DY = 0 THEN
           SELECT CASE FF3EVENT.CODE
              CASE "DEATH   ": FF1REC.DY = FF3EVENT.D1Y
                               FF1REC.DMD = FF3EVENT.D1MD
                               DPLACE = FF3EVENT.PLACE
                               D2ST$ = FF3EVENT.DSTAT
              CASE "BURIAL  ": FF1REC.DY = FF3EVENT.D1Y
                               FF1REC.DMD = FF3EVENT.D1MD
                               DPLACE = FF3EVENT.PLACE
                               D2ST$ = "+"
           END SELECT
        END IF
        EVNTPTR = FF3EVENT.NXT
     WEND
     IF BD$ <> "" THEN
        SELECT CASE LEN(BD$)
           CASE 11: LSET BD$ = XSCRNDate$(FF1REC.BMD, FF1REC.BY): MID$(BD$, 11, 1) = D1ST$
                    LSET DD$ = XSCRNDate$(FF1REC.DMD, FF1REC.DY): MID$(DD$, 11, 1) = D2ST$
           CASE 12: LSET BD$ = XRPTDate$(FF1REC.BMD, FF1REC.BY, 1): MID$(BD$, 12, 1) = D1ST$
                    LSET DD$ = XRPTDate$(FF1REC.DMD, FF1REC.DY, 1): MID$(DD$, 12, 1) = D2ST$
        END SELECT
     END IF
     IF BPL$ <> "" THEN
        LSET BPL$ = BPLACE: LSET DPL$ = DPLACE
     END IF
     IF AGE$ <> "" THEN
        RSET AGE$ = DateDif$(FF1REC.BMD, FF1REC.BY, FF1REC.DMD, FF1REC.DY)
     END IF
END SUB

REM $DYNAMIC
SUB INDEXFile (OPT$)
    QUIET = (RIGHT$(OPT$, 1) = "Q"): X = LEN(RTRIM$(OPT$)) + QUIET
    XOPT$ = MID$(OPT$, 1, X)
    OKAY = 0: D$ = FDFSetup.NFDRV
    IF (("A" <= D$) AND (D$ <= "Z")) AND (D$ <> FDFSetup.FAMDRV) THEN
       NDXPFX$ = D$ + ":"
       ELSE
       NDXPFX$ = FAMFPFX$
    END IF
GETNDXNM:
    IF NOT QUIET THEN
       CALL FmtFIND("XXXXS012")
       CALL FmtFindFLD("DSNM"): LSET UTXT$(FFLD.UTXT) = FT$(4)
       CALL MsgFORMAT: IF A = 27 THEN GOTO NDXOpenBAD
       CALL FmtFindFLD("DSNM"): LSET FT$(4) = UTXT$(FFLD.UTXT)
    END IF
    DSNAME$ = RTRIM$(NDXPFX$ + FT$(4))
    LSET SYSVAR.FILENAME = DSNAME$
    ON LOCAL ERROR GOTO NONDX
CHKNDX:
    OPEN DSNAME$ FOR INPUT AS #4
    CLOSE #4
    IF XOPT$ = "CREATE" THEN
       A$ = ""
       WHILE A$ <> SNGLKEY$(1)
          CALL PutMSG("XXXXM002")
          IF (A$ = SNGLKEY$(2)) OR (A$ = CHR$(27)) THEN GOTO NDXOpenBAD
       WEND
       KILL SYSVAR.FILENAME
       GOTO NEWNDX
    END IF
    GOTO OLDNDX
NONDX:
    X = ERR
    IF (XOPT$ = "CREATE") THEN RESUME NEWNDX
    IF QUIET THEN RESUME NDXOpenBAD
    CALL ERRMessage("XXXXM030", X, 0)
    'IF A$ = CHR$(27) THEN RESUME GETNDXNM
    'RESUME CHKNDX
    RESUME NDXOpenBAD
OLDNDX:
    ON LOCAL ERROR GOTO NDXOpenERR
    SELECT CASE XOPT$
       CASE "CHECK": OPEN SYSVAR.FILENAME FOR BINARY ACCESS READ AS #4
       CASE "READ": OPEN SYSVAR.FILENAME FOR BINARY ACCESS READ AS #4
       CASE "WRITE": OPEN SYSVAR.FILENAME FOR BINARY ACCESS READ WRITE AS #4
       CASE ELSE: GOTO NDXOpenBAD
    END SELECT
    GET #4, , NDXHdr
    LSET SYSVAR.MVAR = "INDEX"
    IF NDXHdr.FTYPE <> "X" THEN
       IF NOT QUIET THEN CALL PutMSG("XXXXM033")
       GOTO NDXOpenBAD
    END IF
    IF NDXHdr.VER <> CHR$(1) THEN
       IF NOT QUIET THEN CALL PutMSG("XXXXM037")
       GOTO NDXOpenBAD
    END IF
    IF (NDXHdr.REORGY <> FF1HDR.REORGY) OR (NDXHdr.REORGMD <> FF1HDR.REORGMD) OR (NDXHdr.REORGTM <> FF1HDR.REORGTM) THEN
       IF NOT QUIET THEN CALL PutMSG("XXXXM034")
       GOTO NDXOpenBAD
    END IF
    IF (XOPT$ = "READ") AND (FF1HDR.MAXID > NDXHdr.NDXRECS) THEN
       CALL PutMSG("XXXXM074")
       IF A$ <> SNGLKEY$(1) THEN GOTO NDXOpenBAD
    END IF
    OKAY = -1
    SFCNT = NDXHdr.OPT AND 31
    REDIM SX(12)
    X = 0: WHILE X < SFCNT: X = X + 1: GET #4, , SX(X): WEND
    ENV.HUSBSURNM = ((NDXHdr.OPT AND 32) > 0)
    ENV.SortOPTS = (NDXHdr.OPT \ 64)
    IF XOPT$ = "CHECK" THEN GOTO NDXOpenEND
    IF NDXHdr.NDXRECS > FF1HDR.MAXID THEN NDXHdr.NDXRECS = FF1HDR.MAXID
    REDIM NDXID(FF1HDR.MAXID)
    C = 0: X = 0
    WHILE X < NDXHdr.NDXRECS
       X = X + 1: GET #4, , XID
       IF XID <= FF1HDR.MAXID THEN C = C + 1: NDXID(C) = XID
    WEND
    GOTO NDXOpenEND
NEWNDX:
    ON LOCAL ERROR GOTO NDXOpenERR
    IF ENV.SortSTAT THEN
       SortSTAT!(10) = 0: CALL SortStatSHOW: HP5 = SortPHASE
       SortPHASE = 5: CALL SortStatACCUM
    END IF
    OPEN SYSVAR.FILENAME FOR BINARY ACCESS READ WRITE AS #4
    NDXHdr.FTYPE = "X"
    NDXHdr.VER = CHR$(1)
    NDXHdr.OPT = (SFCNT - 32 * ENV.HUSBSURNM + 64 * ENV.SortOPTS)
    NDXHdr.REORGY = FF1HDR.REORGY
    NDXHdr.REORGMD = FF1HDR.REORGMD
    NDXHdr.REORGTM = FF1HDR.REORGTM
    NDXHdr.UPDTY = VAL(RIGHT$(DATE$, 4))
    NDXHdr.UPDTMD = 100 * (VAL(LEFT$(DATE$, 2))) + VAL(MID$(DATE$, 4, 2))
    NDXHdr.NDXRECS = FF1HDR.MAXID
    PUT #4, 1, NDXHdr
    X$ = " ": X = 0
    WHILE X < SFCNT: X = X + 1: PUT #4, , SX(X): WEND
    X = 0
    WHILE X < FF1HDR.MAXID
       X = X + 1: PUT #4, , NDXID(X)
       IF ENV.SortSTAT AND (X MOD 100 = 0) THEN
          SortSTAT!(10) = X: CALL SortStatSHOW
       END IF
    WEND
    IF ENV.SortSTAT THEN
       SortSTAT!(10) = X: CALL SortStatSHOW: SortPHASE = HP5
    END IF
    OKAY = -1
    GOTO NDXOpenEND
NDXOpenERR:
    CALL ERRMessage("XXXXM030", ERR, ERL)
    RESUME NDXOpenBAD
NDXOpenBAD:
    OKAY = 0
NDXOpenEND:
    IF NOT OKAY THEN CALL FamCLOSE ELSE CLOSE #4
END SUB

REM $STATIC
SUB PrintLINE
2100 ' Print Line of DATA
     IF A <> 27 THEN
        PAGE.LOFST = LEN(RTRIM$(PLINE$))
        IF PAGE.LOFST > 0 THEN
           PAGE.RPAGE = 0: PAGE.RLINE = 0: CALL WriteLINE
           IF RPTOPTION.BLDREF THEN
              IF RREFRpt.BGNPAGE = 0 THEN RREFRpt.BGNPAGE = PAGE.COUNT
              IF (PAGE.XREF > 0) THEN CALL RREFBld
           END IF
        END IF
     END IF
     PAGE.LOFST = PAGE.COFST: PAGE.LWIDTH = PAGE.BWIDTH
END SUB

SUB PrintOPEN
    MD = DATEMD(DATE$, 1): Y = DATEY(DATE$, 1)
    SYSVAR.DATE = XRPTDate$(MD, Y, 0)
    IF ENV.SYSMODE = 0 THEN
       RPTOPTION.BLDREF = 0
       ELSE
       IF RPTOPTION.BLDREF THEN
          CALL PutMSG("XXXXM061"): IF A = 27 THEN 1390
          RPTOPTION.BLDREF = (A$ = SNGLKEY$(1))
          IF RPTOPTION.BLDREF THEN
             IF RPTOPTION.SNGLREF THEN
                CALL PutMSG("XXXXM063"): IF A = 27 THEN 1390
                RPTOPTION.SNGLREF = VAL(A$) - 1
             END IF
             CALL RREFOpen("WRITE"): IF (NOT OKAY) OR (A = 27) THEN 1390
          END IF
       END IF
    END IF
    IF PAGE.BGNPAGE = 0 THEN PAGE.BGNPAGE = 1
    PAGE.TOP = 0: PAGE.LINE = 0: PAGE.COUNT = PAGE.BGNPAGE - 1
    PAGE.BLTOP = PDFFORMS.TMARGIN: PAGE.BLBOT = PDFFORMS.BMARGIN
    PAGE.IMARGIN = PDFFORMS.IMARGIN: PAGE.OMARGIN = PDFFORMS.OMARGIN
    PRT.FWIDTH = PDFFORMS.FWIDTH: PRT.FLENGTH = PDFFORMS.FLENGTH
    CALL RptOptGET("FOFS", X): PRT.FOFFSET = X
    PRT.SNGLSIDE = ((PDFFORMS.OPTIONS AND 1) > 0)
    PRT.FORMFEED = ((PDFFORMS.OPTIONS AND 2) > 0)
    PRT.PAGEWAIT = ((PDFFORMS.OPTIONS AND 4) > 0)
    CALL RptOptGET("CFRM", X): PRT.CFORMS = (X = 1)
    PRT.DEST = PDFSetup.DEST
    CALL RptOptGET("APST", X)
    PRT.ALL = (X = 4): PRT.AGAIN = ((X = 1) OR (PRT.ALL))
    PAGE.LOFST = 1: PAGE.COFST = 1
    PAGE.RLINE = 0: PAGE.RPAGE = 0: PAGE.WAIT = PRT.PAGEWAIT
    PRT.EOF = 0
    LF$ = ""
    IF PAGE.BWIDTH > 0 THEN
       PAGE.WIDTH = PAGE.BWIDTH + PAGE.IMARGIN + PAGE.OMARGIN
       ELSE
       PAGE.WIDTH = PDFFORMS.FWIDTH
       PAGE.BWIDTH = PAGE.WIDTH - PAGE.IMARGIN - PAGE.OMARGIN
    END IF
    PAGE.LWIDTH = PAGE.BWIDTH
    PLINE$ = SPACE$(255): CL$ = SPACE$(255)
    IF NOT PRT.SNGLSIDE THEN
       CALL PutMSG("XXXXM072")
       IF ASC(A$) = 27 THEN 1390
       PRT.SNGLSIDE = VAL(A$)
       IF (PRT.SNGLSIDE < 1) OR (PRT.SNGLSIDE > 2) THEN
          PRT.SNGLSIDE = 3
          ELSE
          PRT.SNGLSIDE = PRT.SNGLSIDE XOR 3
       END IF
       ELSE
       IF ((RPTOPTION.LINEREF AND 1) > 0) THEN
          IF ((RPTOPTION.LINEREF AND 24) > 0) THEN
             RPTOPTION.LINEREF = 1 + (RPTOPTION.LINEREF \ 4)
          END IF
       END IF
    END IF
    IF PDFSetup.DEST = "SCRN:" THEN OPT = 1: GOTO 1265
    IF PDFSetup.DEST = "FILE " THEN PRT.OPT = 3: GOTO 1275
1260 CALL PutMSG("XXXXM008")
     IF ASC(A$) = 27 THEN 1390
     OPT = VAL(A$): IF OPT < 1 OR OPT > 3 THEN CALL ErrBEEP(0): GOTO 1260
1265 PRT.OPT = OPT
     IF OPT = 1 THEN PRT.DEST = "SCRN:" ELSE PRT.DEST = PDFSetup.DEST
     IF PRT.OPT <> 3 THEN DEST$ = PRT.DEST: GOTO 1300
1275 ' File Output
     CALL FmtFIND("XXXXS011")
     CALL FmtFindFLD("FNAME"): NMFLD = WIN.CURFLD
     LSET UTXT$(FLDTAB(NMFLD).UTXT) = FT$(14)
     CALL MsgFORMAT
     IF A = 27 THEN 1260
     LSET PRT.DEST = UTXT$(FLDTAB(NMFLD).UTXT)
     IF INSTR(PT$(3), CHR$(10)) = 0 THEN LF$ = CHR$(10)
     DEST$ = RPTFPFX$ + RTRIM$(PRT.DEST)
1300 LSET SYSVAR.FILENAME = DEST$
     ON LOCAL ERROR GOTO 1330
     GOTO 1350
1330 ' Error Opening Print
     CALL ERRMessage("XXXXM030", ERR, ERL)
     IF PRT.OPT = 1 THEN A = 27: RESUME 1390 ELSE RESUME 1260
1350 OPEN DEST$ FOR OUTPUT AS #4
     SELECT CASE PRT.OPT
        CASE 1: WIDTH #4, 80
        CASE 2: WIDTH #4, 255
        CASE 3: LSET FT$(14) = PRT.DEST
     END SELECT
     IF OPT > 1 THEN PRINT #4, PT$(1); PT$(2); :  ELSE LOCATE 1, 1, 0: PRINT CHR$(13);
     ON LOCAL ERROR GOTO 0
     IF PRT.OPT = 1 THEN
        IF NOT ENV.ScrnCLR THEN CALL ScrnCLR
        PRT.FWIDTH = 79: PRT.FLENGTH = 23
        PRT.PAGEWAIT = -1: PAGE.WAIT = -1
        PRT.FORMFEED = 0: PAGE.BLTOP = 0: PAGE.BLBOT = 0
     END IF
     CALL PutMSG("XXXX" + RDFReport.PRTMSG)
     OKAY = -1: PRT.OPEN = CHR$(255): EXIT SUB
1390 ' Bad EXIT From Routine
     ON LOCAL ERROR GOTO 0
     CALL PutMSG("")
     OKAY = 0
END SUB

SUB PutDATA
2000 ' Place Data on Line
     IF PAGE.LOFST > (PAGE.LWIDTH + 1) THEN
        CALL PrintLINE: IF A = 27 THEN 2009
     END IF
     IF (RPTOPTION.BLDREF) AND (PAGE.XID > 0) THEN CALL RREFAdd
     IF PAGE.LOFST = 0 THEN PAGE.LOFST = PAGE.COFST
2005 ET = LEN(RTRIM$(CL$)): IF ET < 1 THEN 2009
     BT = 1
     IF (PAGE.LOFST <= PAGE.COFST) THEN
        C$ = MID$(CL$, 1, 1)
        WHILE (BT < ET + 1) AND ((C$ = " ") OR (C$ = ".") OR (C$ = ",") OR (C$ = ";"))
           BT = BT + 1: LSET C$ = MID$(CL$, BT, 1)
        WEND: IF ET < BT THEN LSET CL$ = "": GOTO 2009
     END IF
2006 X = BT + PAGE.LWIDTH - PAGE.LOFST + 1
     IF X <= ET THEN XL$ = MID$(CL$, X, ET - X + 1): ET = X - 1
     MID$(PLINE$, PAGE.LOFST, ET - BT + 1) = MID$(CL$, BT, ET - BT + 1)
     PAGE.LOFST = PAGE.LOFST + ET - BT + 1
2007 IF (XL$ <> "") THEN
        X = PAGE.LWIDTH: C$ = MID$(PLINE$, X, 1)
        WHILE (X > PAGE.COFST) AND (C$ <> " ") AND (C$ <> ",") AND (C$ <> ".") AND (C$ <> ";")
           X = X - 1: LSET C$ = MID$(PLINE$, X, 1)
        WEND: IF X = PAGE.COFST THEN X = PAGE.LWIDTH
2008    CW = PAGE.LWIDTH - X: LSET CL$ = MID$(PLINE$, X + 1, CW)
        MID$(CL$, CW + 1, 255) = XL$: XL$ = ""
        IF CW THEN MID$(PLINE$, X + 1, CW) = SPACE$(CW)
        CALL PrintLINE: IF A = 27 THEN 2009
        GOTO 2005
     END IF
2009 ' Finished
     LSET CL$ = ""
END SUB

FUNCTION Relation$ (RT$, RSEX$, AGL, RGL, RCODE$())
     RELWORK$ = SPACE$(50): L1 = 1
     LATREL = ((ENV.RELRULES AND 2) = 2)
     SELECT CASE RT$
        CASE "S": X1 = 3: Y1 = 3: X2 = 3: Y2 = 3
        CASE "L": X1 = 6: Y1 = 12: X2 = 6: Y2 = 12: LongREL = -1
        CASE ELSE: X1 = 3: Y1 = 3: X2 = 6: Y2 = 12
     END SELECT
     IF (RGL < AGL) THEN LSET RELWORK$ = MID$(RCODE$(1), X2, Y2): GOTO 9549
     IF (RGL = 0) AND (AGL = 0) THEN LSET RELWORK$ = MID$(RCODE$(2), X2, Y2): GOTO 9549
     IF RSEX$ = FGENDR$(1) THEN XGEN = 0 ELSE IF RSEX$ = FGENDR$(2) THEN XGEN = 1 ELSE XGEN = 2
     OGEN = XGEN + 2 * (XGEN > 1)
9505 ' Direct DESCENDANTS
     IF AGL <> 0 THEN 9510
     XGL = -RGL: XC = 6
     IF LongREL AND (RSEX$ = "X") THEN XC = 33: XGEN = 0
     IF LATREL AND (XGL < -1) THEN
        IF XC = 6 THEN XC = 28 ELSE XC = 34
     END IF
     GOTO 9535
9510 ' Direct ANCESTORS
     IF AGL <> RGL THEN 9520
     XGL = AGL: XC = 3
     IF LongREL AND (RSEX$ = "X") THEN XC = 31: XGEN = 0
     IF LATREL AND (XGL < -1) THEN
        IF XC = 3 THEN XC = 25 ELSE XC = 32
     END IF
     GOTO 9535
9520 ' Brothers and Sisters
     IF AGL <> -1 THEN 9530
     IF RGL > 0 THEN 9525
     XGL = AGL: XC = 9: GOTO 9535
9525 ' Nieces and Nephews
     XGL = -RGL: XC = 15: XG = 28: GOTO 9535
9530 ' Aunts and Uncles
     IF (RGL - AGL) > 1 THEN 9540
     XGL = AGL + 1: XC = 12: XG = 25
9535 ' Build Relation
     IF LATREL THEN
        IF XGL < -2 THEN MID$(RELWORK$, L1, 3) = LTRIM$(STR$(-(XGL + 1)))
        L1 = INSTR(L1, RELWORK$, "  ") - LongREL
        ELSE
        IF LongREL AND (XGL < -3) THEN
           MID$(RELWORK$, L1, 3) = LTRIM$(STR$(-(XGL + 2)))
           L1 = INSTR(L1, RELWORK$, "  ") + 1
        END IF
        IF XGL < -2 THEN
           MID$(RELWORK$, L1, Y1) = MID$(RCODE$(22 + OGEN), X1, Y1)
           L1 = INSTR(L1, RELWORK$, "  ") - LongREL
           IF (XGL < -3) AND (NOT LongREL) THEN
              MID$(RELWORK$, L1, 3) = LTRIM$(STR$(-(XGL + 2)))
              L1 = INSTR(L1, RELWORK$, "  ")
           END IF
        END IF
        IF (XGL < -1) THEN
           MID$(RELWORK$, L1, Y1) = MID$(RCODE$(20 + OGEN), X1, Y1)
           L1 = INSTR(RELWORK$, "  ") - LongREL
        END IF
     END IF
9536 MID$(RELWORK$, L1, Y2) = MID$(RCODE$(XC + XGEN), X2, Y2)
     L1 = INSTR(L1, RELWORK$, " ") - LongREL
     IF LATREL AND (XG > 0) THEN
        IF XGL < -1 THEN
           MID$(RELWORK$, L1, Y2) = MID$(RCODE$(XG + XGEN), X2, Y2)
        END IF
     END IF
     GOTO 9549
9540 ' Cousins
     CIVREL = ((ENV.RELRULES AND 3) > 0)
     CGL = ABS(AGL + (RGL) * (RGL < 0) + 1)
     IF CIVREL THEN CGL = (2 * CGL + ABS(RGL) - 1)
     LSET RELWORK$ = LTRIM$(STR$(CGL))
     L1 = INSTR(RELWORK$, "  ") - LongREL
     IF (RGL = 0) OR (CIVREL) THEN
        MID$(RELWORK$, L1, Y2) = MID$(RCODE$(18 + OGEN), X2, Y2)
        GOTO 9549
     END IF
     MID$(RELWORK$, L1, Y1) = MID$(RCODE$(18 + OGEN), X1, Y1)
     L1 = INSTR(L1, RELWORK$, " ") - LongREL
     MID$(RELWORK$, L1, 4) = LTRIM$(STR$(ABS(RGL)))
     L1 = INSTR(L1, RELWORK$, " ") - LongREL
     MID$(RELWORK$, L1, Y1) = MID$(RCODE$(24), X1, Y1)
9549 Relation$ = RTRIM$(RELWORK$)
END FUNCTION

REM $DYNAMIC
SUB RelCALC (BRID)
1500 ' RELationship CALCulator
     IF WINTAB(6).WFMT <> 0 THEN
        CALL WinSWITCH(6)
        CALL WinCLR
     END IF
     CALL WinSWITCH(4)
     CALL FmtFIND("UPDTS917"): FMT917 = CURFMT
     CALL FmtFIND("UPDTS032"): FMT032 = CURFMT
     CALL FMTDREST
     CALL FmtFIND("UPDTS031"): FMT031 = CURFMT
     CALL FMTDREST
     CALL FmtFIND("UPDTS030"): FMT030 = CURFMT
     CALL FmtFindFLD("BDT "): LSET DTXT$(WIN.CURFLD) = XSCRNDate$(0, 0)
     CALL FmtFindFLD("DDT "): LSET DTXT$(WIN.CURFLD) = XSCRNDate$(0, 0)
     CALL FMTDREST: GOSUB 1510
     CALL OptDISPLAY(FMT917)
     CALL WinFORMAT(FMT030)
     RELWORK$ = SPACE$(12)
     ID = BRID
     IF ID <> 0 THEN GOTO RCShowRID1:
RCGetRID1:
     CALL OptHILITE("F1", "", 6)
     CURFMT = FMT031: FMT = FMTTAB(CURFMT)
     IF OKAY THEN CALL FmtUHOLD
     CALL WinFORMAT(0)
RCGR1:
     WIN.CLIN = 0: CALL WinUPDATE
     IF A = 27 THEN
        IF RID1 = 0 THEN CALL WinCLR: EXIT SUB
        CALL FMTHREST: CALL WinFORMAT(0): GOTO RCGetOPT
     END IF
     CALL FmtFindFLD("RID1")
     X& = VAL(UTXT$(FFLD.UTXT)): CALL ChkID(X&)
     IF OKAY THEN ID = X& ELSE GOTO RCGR1
RCShowRID1:
     OKAY = ID > 0
     IF OKAY THEN CALL FF1GetRec(ID)
     IF NOT OKAY THEN CALL ErrBEEP(0): GOTO RCGR1
     RID1 = ID: CURFMT = FMT031: FMT = FMTTAB(CURFMT)
     CALL FmtFindFLD("RID1"): RSET UTXT$(FFLD.UTXT) = LTRIM$(STR$(RID1))
     CALL FmtFindFLD("NAM1"): LSET UTXT$(FFLD.UTXT) = FMTNAME$(1)
     CALL WinFORMAT(0)
     GOSUB 1520
     CALL OptDISPLAY(0)
     IF RID2 <> 0 THEN GOSUB 1550: GOTO RCGetOPT
RCGetRID2:
     CALL OptHILITE("F2", "", 6)
     CURFMT = FMT032: FMT = FMTTAB(CURFMT)
     IF OKAY THEN CALL FmtUHOLD
     CALL WinFORMAT(0)
RCGR2:
     WIN.CLIN = 0: CALL WinUPDATE
     IF A = 27 THEN
        IF RID2 = 0 THEN CALL WinCLR: EXIT SUB
        CALL FMTHREST: CALL WinFORMAT(0): GOTO RCGetOPT
     END IF
     CALL FmtFindFLD("RID2")
     X& = VAL(UTXT$(FFLD.UTXT)): CALL ChkID(X&)
     IF (NOT OKAY) OR (X& < 1) THEN GOTO RCGR2
     ID = X&: CALL FF1GetRec(ID)
     RID2 = ID: SEX2$ = FF1REC.SEX
     CALL FmtFindFLD("NAM2"): LSET UTXT$(FFLD.UTXT) = FMTNAME$(1)
     CALL WinFORMAT(0)
     GOSUB 1550
RCGetOPT:
     CALL OptDISPLAY(FMT917)
     CALL GetKEY
     SELECT CASE A
        CASE 27: FUN$ = "ESC"
        CASE 59: OKAY = -1: GOTO RCGetRID1
        CASE 60: OKAY = -1: GOTO RCGetRID2
        CASE 61: IF ID > 0 THEN FUN$ = "F3"
        CASE 62: FUN$ = "F4"
        CASE 63: FUN$ = "F5"
        CASE ELSE: CALL ErrBEEP(0): GOTO RCGetOPT
     END SELECT
     CALL OptHILITE(FUN$, "", 6)
     SELECT CASE FUN$
        CASE "ESC": CALL WinCLR: EXIT SUB
        CASE "F3": GOSUB 1555: GOTO RCGetOPT
        CASE "F4": ENV.RELRULES = ((RELRULE + 1) MOD 3) + RELTYPE
                   GOSUB 1510: GOSUB 1585
                   GOTO RCGetOPT
        CASE "F5": ENV.RELRULES = RELRULE + ((RELTYPE + 4) MOD 12)
                   GOSUB 1510
                   GOSUB 1520: GOSUB 1550
                   GOTO RCGetOPT
     END SELECT
     GOTO RCGetOPT

1510 ' Set RELRULES
     RELRULE = ENV.RELRULES AND 3: RELTYPE = ENV.RELRULES AND 12
     MATERNAL = ((RELTYPE AND 4) > 0)
     PATERNAL = ((RELTYPE AND 8) > 0)
     CURFMT = FMT030: FMT = FMTTAB(CURFMT)
     SELECT CASE RELRULE
        CASE 0: RRULE$ = RELRULE$(1)
        CASE 1: RRULE$ = RELRULE$(2)
        CASE 2: RRULE$ = RELRULE$(0)
     END SELECT
     CALL FmtFindFLD("RULE")
     LSET UTXT$(FFLD.UTXT) = RRULE$
     SELECT CASE RELTYPE
        CASE 0: RTYPE$ = RELRULE$(3)
        CASE 4: RTYPE$ = RELRULE$(4)
        CASE 8: RTYPE$ = RELRULE$(5)
     END SELECT
     CALL FmtFindFLD("TYPE")
     LSET UTXT$(FFLD.UTXT) = RTYPE$
     RETURN

1520 'Build Ancestor Table for ID #1
     CALL PutMSG("UPDTM011")
     AMAX1 = 100: AMAX2 = 100: Y2 = 0
     REDIM RELATE(FF1HDR.MAXID), ATAB(AMAX1), ARGL1(AMAX1), ALIN1(AMAX1) AS DOUBLE
     AGL = 0: X = 1: Y = 1: ATAB(1) = RID1: ALIN1(1) = 1
     RELATE(RID1) = 1: RELTABOK = 0
1525 AGL = AGL + 1                ' Process next generation of ancestors
     Z = Y
     WHILE (X <= Z)
        ID = ATAB(X): CALL FF1GetRec(ID)
        ID = FF1REC.FID
        IF (ID > 0) AND (NOT MATERNAL) THEN P2 = 1: GOSUB 1540
        ID = FF1REC.MID
        IF (ID > 0) AND (NOT PATERNAL) THEN P2 = 2: GOSUB 1540
        A$ = INKEY$: IF A$ = CHR$(27) THEN X = Z + 1 ELSE X = X + 1
     WEND
     IF A$ <> CHR$(27) THEN
        IF X <= Y THEN 1525          ' If any found, go process them
        RELTABOK = -1
     END IF
     CALL PutMSG("")
     RETURN

1540 IF RELATE(ID) = 0 THEN
        Y = Y + 1: RELATE(ID) = Y
        IF Y > AMAX1 THEN
           AMAX1 = AMAX1 + 100
           REDIM PRESERVE ATAB(AMAX1), ARGL1(AMAX1), ALIN1(AMAX1) AS DOUBLE
        END IF
        ATAB(Y) = ID: ARGL1(Y) = AGL
        ALIN1(Y) = 2 * ALIN1(X) + P2 - 1
     END IF
     RETURN

1550 ' Find Common Ancestor of RID2
     IF NOT RELTABOK THEN GOSUB 1520
     IF NOT RELTABOK THEN LSET RELWORK$ = "": GOTO 1585
     X = 1: WHILE X <= Y2: RELATE(ATAB(X)) = 0: X = X + 1: WEND
     AGL = 0: ID = 0: RLCT = 1
     X2 = 1: Y2 = 0: Z2 = 0
     IF RID2 = RID1 THEN AGL1 = 0: RGL = 0: GOTO 1585
     REDIM ATAB(AMAX2), ALIN2(AMAX2)  AS DOUBLE
     ATAB(1) = RID2: ALIN2(1) = 1
     ID = RID2
     FOUND = (RELATE(ID) > 0): IF FOUND THEN LIN2# = 1: GOTO 1585
     ' Look for Common Ancestor
     RLCT = 0: X2 = 1: Y2 = 1: RELATE(RID2) = -1
1552 AGL = AGL + 1: Z2 = Y2: P2 = 0 ' Process next generation of ancestors
1555 A$ = "": FOUND = 0
1556 WHILE (X2 <= Z2) AND (NOT FOUND)
        ID = ATAB(X2): CALL FF1GetRec(ID)
        IF ((P2 AND 1) = 0) THEN
           P2 = P2 OR 1
           IF (NOT MATERNAL) THEN
              ID = FF1REC.FID: IF ID > 0 THEN GOSUB 1570
           END IF
        END IF
        IF (NOT FOUND) THEN
           P2 = 0
           IF (NOT PATERNAL) THEN
              ID = FF1REC.MID: IF ID > 0 THEN GOSUB 1570
           END IF
        END IF
        IF (P2 = 0) THEN
           X2 = X2 + 1
           IF NOT FOUND THEN
              A$ = INKEY$: FOUND = (A$ = CHR$(27))
              IF FOUND THEN ID = 0: LSET RELWORK$ = ""
           END IF
        END IF
     WEND
     IF FOUND THEN RLCT = RLCT + 1: GOTO 1585
     IF X2 <= Y2 THEN 1552
     IF FUN$ = "F3" THEN FUN$ = "": GOTO 1550
     AGL1 = 1: RGL = 0: ID = 0
     GOTO 1585

1570 ' Check Ancestor of RID2
     LIN2# = 2 * ALIN2(X2) - (P2 = 0)
     IF (RELATE(ID) = 0) THEN
        Y2 = Y2 + 1: RELATE(ID) = -Y2
        IF Y2 > AMAX2 THEN
           AMAX2 = AMAX2 + 100
           REDIM PRESERVE ATAB(AMAX2), ALIN2(AMAX2) AS DOUBLE
        END IF
        ATAB(Y2) = ID
        ALIN2(Y2) = LIN2#
        ELSE
        FOUND = (RELATE(ID) > 0)
     END IF
     RETURN

1585 ' Update Screen Fields
     IF ID > 0 THEN
        LIN1# = ALIN1(RELATE(ID))
        AGL1 = -ARGL1(RELATE(ID)): AGL2 = -AGL
        RGL = AGL1 - AGL2
        CALL FF1GetRec(ID)
     END IF
     LSET RELWORK$ = Relation$("M", SEX2$, AGL1, RGL, RELCODE$())
     CURFMT = FMT030: FMT = FMTTAB(CURFMT)
     CALL FMTDREST
     WIN.CURFLD = FMT.BGNFLD
     WHILE WIN.CURFLD > 0
        CALL GetFldPARMS
        IF FFLD.UTXT <> 0 THEN
           X$ = UTXT$(FFLD.UTXT)
           SELECT CASE FFLD.NAME
              CASE "RULE": LSET X$ = RRULE$
              CASE "TYPE": LSET X$ = RTYPE$
              CASE "RLCT": RSET X$ = LTRIM$(STR$(RLCT))
              CASE "RELA": LSET X$ = RELWORK$
              CASE "AGL1": IF ID > 0 THEN RSET X$ = STR$(AGL1)
              CASE "LIN1": IF ID > 0 THEN LSET X$ = STR$(LIN1#)
              CASE "AGL2": IF ID > 0 THEN RSET X$ = STR$(AGL2)
              CASE "LIN2": IF ID > 0 THEN LSET X$ = STR$(LIN2#)
              CASE "RGL ": IF ID > 0 THEN LSET X$ = STR$(RGL)
              CASE "RID3": IF ID > 0 THEN RSET X$ = LTRIM$(STR$(ID))
              CASE "NAM3": IF ID > 0 THEN LSET X$ = FMTNAME$(1)
              CASE "BDT ": IF ID > 0 THEN LSET X$ = XSCRNDate$(FF1REC.BMD, FF1REC.BY)
              CASE "BDST": IF ID > 0 THEN LSET X$ = FDSTAT$("BD")
              CASE "DDT ": IF ID > 0 THEN LSET X$ = XSCRNDate$(FF1REC.DMD, FF1REC.DY)
              CASE "DDST": IF ID > 0 THEN LSET X$ = FDSTAT$("DD")
           END SELECT
           LSET UTXT$(FFLD.UTXT) = X$
        END IF
        WIN.CURFLD = FLDCHN(WIN.CURFLD).FWD
     WEND
     CALL WinDisplayDATA(0)
     RETURN
END SUB

REM $STATIC
SUB RptCLOSE
    IF A <> 27 THEN
       CALL EvenBREAK("EOF")
       IF LEN(RTRIM$(ENV.RTITLE)) > 0 THEN
          CALL HTFEXPAND(ENV.RTITLE, 1)
       END IF
       IF RPTOPTION.BLDREF THEN CALL RREFClose: RPTOPTION.BLDREF = 0
       ELSE
       CALL PutMSG("XXXX" + RDFReport.INTMSG)
    END IF
    IF (A <> 27) AND (PRT.AGAIN) THEN
       PRT.AGAIN = ((PRT.FOFFSET + PRT.FWIDTH) < PAGE.WIDTH)
       PRT.ALL = (PRT.AGAIN AND PRT.ALL)
       WHILE INKEY$ <> "": WEND      ' Clear Key Buffer
       IF (PRT.AGAIN) AND (NOT PRT.ALL) THEN
          PRT.FOFFSET = PRT.FOFFSET + PRT.FWIDTH
          SYSVAR.NVAR1 = PRT.FOFFSET
          CALL PutMSG("XXXXM007")
          PRT.ALL = (A$ = SNGLKEY$(4))
          PRT.AGAIN = ((A$ = SNGLKEY$(1)) OR PRT.ALL)
       END IF
       IF (NOT PRT.AGAIN) AND (PRT.SNGLSIDE = 2) THEN
          CALL PutMSG("XXXXM081")
          PRT.AGAIN = (A$ = SNGLKEY$(1))
          IF PRT.AGAIN THEN
             CALL RptOptGET("FOFS", X): PRT.FOFFSET = X
             PRT.SNGLSIDE = 1
          END IF
       END IF
       ELSE
       PRT.AGAIN = 0: PRT.ALL = 0
    END IF
    IF PRT.AGAIN THEN
       PAGE.COUNT = PAGE.BGNPAGE - 1
       CALL PutMSG("XXXX" + RDFReport.PRTMSG)
       ELSE
       CALL PrintCLOSE
    END IF
    PRT.EOF = 0
END SUB

SUB RREFAdd
    IF RPTOPTION.BLDREF THEN
       IF (PAGE.XID > 0) AND (PAGE.XREF < 100) THEN
          PAGE.XREF = PAGE.XREF + 1
          RREFXID(PAGE.XREF) = PAGE.XID: RREFTYPE(PAGE.XREF) = PAGE.XTYPE
       END IF
       PAGE.XID = 0
    END IF
END SUB

SUB RREFBld
     IF PAGE.XREF > 0 THEN
        Z = LEN(RREFXData$): Z$ = CHR$(Z): Z = Z - (Z > 0)
        FOR X = 1 TO PAGE.XREF: GOSUB 1920: NEXT X
        PAGE.XREF = 0
     END IF
     EXIT SUB

1920 ' Create Reference File Entry
     XID = RREFXID(X): XTYPE = RREFTYPE(X): XPTR& = RX&(XID)
     AGL = ASC(RREFRec.AGL): DGL = ASC(RREFRec.DGL)
     RREFRpt.NUMREFS = RREFRpt.NUMREFS + 1: OK = -1
     IF XPTR& <> 0 THEN
        GET #6, XPTR& + 1, RREFRec
        IF RPTOPTION.SNGLREF AND (XPTR& > RREFHdr.EOFRBA) THEN
           OK = 0
           X = RREFRec.REFTYPE: Y = XTYPE
           IF RPTOPTION.SNGLREF = 1 THEN
              P1 = 1: P2 = 0
              ELSE
              P1 = RXIMP(X): P2 = RXIMP(Y)
           END IF
           IF (P2 <= P1) THEN
              RREFRec.REFTYPE = XTYPE + (RREFRec.REFTYPE AND 1024)
              RREFRec.PGNUM = PAGE.RPAGE: RREFRec.LINENUM = PAGE.RLINE
           END IF
           ELSE
           RREFRec.NEXT = RREFRpt.EOFRBA
        END IF
        PUT #6, XPTR& + 1, RREFRec
        ELSE
        RREFRpt.NEWID = RREFRpt.NEWID + 1
     END IF
     IF OK THEN
        IF XPTR& < RREFHdr.EOFRBA THEN RREFRpt.NUMID = RREFRpt.NUMID + 1
        RX&(XID) = RREFRpt.EOFRBA
        RREFRec.RID = XID: RREFRec.AGL = CHR$(AGL): RREFRec.DGL = CHR$(DGL)
        RREFRec.REFTYPE = XTYPE - 1024 * (Z > 0)
        RREFRec.PGNUM = PAGE.RPAGE: RREFRec.LINENUM = PAGE.RLINE
        RREFRec.NEXT = 0
        PUT #6, RREFRpt.EOFRBA + 1, RREFRec
        IF Z > 0 THEN PUT #6, , Z$: PUT #6, , RREFXData$
        RREFRpt.EOFRBA = RREFRpt.EOFRBA + LEN(RREFRec) + Z
     END IF
     RREFRec.AGL = CHR$(0): RREFRec.DGL = CHR$(0)
     RETURN

END SUB

SUB RREFClose
    IF RREFRpt.BGNPAGE > 0 THEN
       RREFRpt.LASTPAGE = PAGE.COUNT
       RREFRpt.TITLE = ENV.RTITLE
       PUT #6, RREFHdr.EOFRBA + 1, RREFRpt
       RREFHdr.WFVER = CHR$(0)
       RREFHdr.RPTS = CHR$(ASC(RREFHdr.RPTS) + 1)
       RREFHdr.NUMID = RREFHdr.NUMID + RREFRpt.NEWID
       RREFHdr.NUMREFS = RREFHdr.NUMREFS + RREFRpt.NUMREFS
       IF RREFHdr.LASTPAGE < RREFRpt.LASTPAGE THEN RREFHdr.LASTPAGE = RREFRpt.LASTPAGE
       RREFHdr.EOFRBA = RREFRpt.EOFRBA
       PUT #6, 1, RREFHdr
       REDIM RX&(1): RREFXData$ = ""
    END IF
    CLOSE #6
END SUB

REM $DYNAMIC
SUB RREFOpen (OPT$)
    ' Open Index REFERENCE File for Creation
    QUIET = (RIGHT$(OPT$, 1) = "Q"): X = LEN(RTRIM$(OPT$)) + QUIET
    XOPT$ = MID$(OPT$, 1, X): DSN$ = FT$(9)
REFDSN:
    IF NOT QUIET THEN
       CALL FmtFIND("XXXXS017")
       CALL FmtFindFLD("DSNM"): LSET UTXT$(FFLD.UTXT) = DSN$
       CALL MsgFORMAT: IF A = 27 THEN GOTO REFOpenBAD
       CALL FmtFindFLD("DSNM"): LSET DSN$ = UTXT$(FFLD.UTXT)
    END IF
    DSNAME$ = RTRIM$(WRKFPFX$ + DSN$)
    LSET SYSVAR.FILENAME = DSNAME$
    ON LOCAL ERROR GOTO NOREF
    OPEN DSNAME$ FOR INPUT AS #6
REFCHK:
    CLOSE #6
    ON LOCAL ERROR GOTO REFOpenERR
    IF XOPT$ = "NEW" THEN
       CALL PutMSG("XXXXM002")
       IF A$ <> SNGLKEY$(1) THEN GOTO REFOpenBAD
    END IF
    IF (XOPT$ = "NEW") THEN
       KILL SYSVAR.FILENAME
       GOTO REFNEW
    END IF
    GOTO REFOLD
NOREF:
    X = ERR
    IF XOPT$ = "WRITE" THEN XOPT$ = "NEW"
    IF (XOPT$ = "NEW") THEN RESUME REFNEW
    IF QUIET THEN RESUME REFOpenBAD
    CALL ERRMessage("XXXXM030", X, 0)
    IF A = 27 THEN RESUME REFOpenBAD
    RESUME REFDSN
REFOLD:
    ON LOCAL ERROR GOTO REFOpenERR
    'ON LOCAL ERROR GOTO 0
    SELECT CASE XOPT$
       CASE "INIT": OPEN SYSVAR.FILENAME FOR BINARY ACCESS READ AS #6
       CASE "UPDT": OPEN SYSVAR.FILENAME FOR BINARY ACCESS READ WRITE AS #6
       CASE "READ": OPEN SYSVAR.FILENAME FOR BINARY ACCESS READ AS #6
       CASE "WRITE": OPEN SYSVAR.FILENAME FOR BINARY ACCESS READ WRITE AS #6
       CASE ELSE: GOTO REFOpenBAD
    END SELECT
    SYSVAR.MVAR = ".REF"
    GET #6, , RREFHdr
    IF (RREFHdr.FTYPE <> "W") OR (RREFHdr.WTYPE <> CHR$(4)) OR (RREFHdr.WFVER <> CHR$(0)) THEN
       IF XOPT$ = "WRITE" THEN XOPT$ = "NEW": GOTO REFCHK
       CLOSE #6: IF QUIET THEN GOTO REFOpenBAD
       CALL PutMSG("XXXXM033")
       IF XOPT$ = "INIT" THEN GOTO REFDSN
       GOTO REFOpenBAD
    END IF
    IF (RREFHdr.REORGY <> FF1HDR.REORGY) OR (RREFHdr.REORGMD <> FF1HDR.REORGMD) OR (RREFHdr.REORGTM <> FF1HDR.REORGTM) THEN
       IF XOPT$ = "WRITE" THEN XOPT$ = "NEW": GOTO REFCHK
       CLOSE #6: IF QUIET THEN GOTO REFOpenBAD
       CALL PutMSG("XXXXM034")
       IF XOPT$ = "INIT" THEN GOTO REFDSN
       GOTO REFOpenBAD
    END IF
    R = ASC(RREFHdr.RPTS)
    IF (XOPT$ = "WRITE") THEN
       IF R > 0 THEN
          SYSVAR.NVAR1 = R + 1
          CALL PutMSG("XXXXM062")
          IF A = 27 THEN GOTO REFOpenBAD
       END IF
       IF (R = 0) OR (A$ = "1") THEN GOTO REFNEW
    END IF
    X& = LEN(RREFHdr): Z$ = " "
    REDIM RX&(FF1HDR.MAXID)
    IF XOPT$ = "INIT" THEN GOTO REFOLDEnd
    WHILE X& < RREFHdr.EOFRBA
       GET #6, X& + 1, RREFRpt
       IF ASC(RREFRpt.RTYPE) AND 128 THEN
          GET #6, , Z$: Z = ASC(Z$)
          ELSE
          Z = 0
       END IF
       X& = X& + LEN(RREFRpt) + Z
       WHILE X& < RREFRpt.EOFRBA
          GET #6, X& + 1, RREFRec
          ' IF "READ" then RX&()=First REF, else RX&()=Last REF
          IF (RREFRec.RID > 0) AND (RREFRec.RID <= FF1HDR.MAXID) THEN
             IF (XOPT$ = "READ") THEN
                IF RX&(RREFRec.RID) = 0 THEN RX&(RREFRec.RID) = X&
                ELSE
                RX&(RREFRec.RID) = X&
             END IF
          END IF
          IF RREFRec.REFTYPE AND 1024 THEN
             GET #6, , Z$: Z = ASC(Z$) + 1
             ELSE
             Z = 0
          END IF
          X& = X& + LEN(RREFRec) + Z
       WEND
    WEND
    IF XOPT$ = "WRITE" THEN GOTO RPTNEW
REFOLDEnd:
    IF R > 0 THEN GET #6, LEN(RREFHdr) + 1, RREFRpt
    OKAY = -1: GOTO REFOpenEnd
REFNEW:
    ON LOCAL ERROR GOTO REFOpenERR
    IF XOPT$ = "WRITE" THEN CLOSE #6: KILL SYSVAR.FILENAME
    OPEN SYSVAR.FILENAME FOR BINARY ACCESS READ WRITE AS #6
    DIM NEWHdr AS RefFileHDR: RREFHdr = NEWHdr
    RREFHdr.FTYPE = "W": RREFHdr.WTYPE = CHR$(4): RREFHdr.WFVER = CHR$(255)
    RREFHdr.REORGY = FF1HDR.REORGY
    RREFHdr.REORGMD = FF1HDR.REORGMD
    RREFHdr.REORGTM = FF1HDR.REORGTM
    RREFHdr.UPDTY = DATEY(DATE$, 1)
    RREFHdr.UPDTMD = DATEMD(DATE$, 1)
    RREFHdr.EOFRBA = LEN(RREFHdr)
    PUT #6, 1, RREFHdr
    REDIM RX&(FF1HDR.MAXID)
RPTNEW:
    DIM NEWRpt AS RefFileRptREC: RREFRpt = NEWRpt
    Z = LEN(RREFXData$)
    RREFRpt.RTYPE = CHR$(RDFReport.TYPE - 128 * (Z > 0))
    RREFRpt.RCODE = RDFReport.CODE
    RREFRpt.TITLE = RDFReport.NAME
    RREFRpt.UPDTY = DATEY(DATE$, 1)
    RREFRpt.UPDTMD = DATEMD(DATE$, 1)
    RREFRpt.UPDTTM = TimeHHMM(TIME$)
    RREFRpt.EOFRBA = RREFHdr.EOFRBA + LEN(RREFRpt)
    PUT #6, RREFHdr.EOFRBA + 1, RREFRpt
    IF Z > 0 THEN
       Z$ = CHR$(Z): PUT #6, , Z$
       PUT #6, , RREFXData$: RREFXData$ = ""
    END IF
    REDIM RREFXID(100), RREFTYPE(100): PAGE.XREF = 0
    IF PAGE.BGNPAGE = 0 THEN PAGE.BGNPAGE = RREFHdr.LASTPAGE + 1
    OKAY = -1: GOTO REFOpenEnd
REFOpenERR:
    CALL ERRMessage("XXXXM030", ERR, ERL)
    RESUME REFOpenBAD
REFOpenBAD: OKAY = 0
REFOpenEnd:
    RREFRec.AGL = CHR$(0): RREFRec.DGL = CHR$(0)
    IF OKAY THEN LSET FT$(9) = DSN$
END SUB

SUB RWRKBuild (WTYPE, BRID, FMTA, FMTSTAT)
500 ' Create work file
    ON LOCAL ERROR GOTO 595
    CALL FAMOpen("READ"): IF NOT OKAY THEN 590
    FMTBRID = FMTA
    IF FMTBRID = 0 THEN
       CALL FmtFIND("XXXXS010")
       IF FOUND THEN FMTBRID = CURFMT ELSE GOTO 590
    END IF
    CURFMT = FMTBRID: FMT = FMTTAB(CURFMT)
    CALL FmtFindFLD("BRID"): IF NOT FOUND THEN 590
    BFLD = FFLD.UTXT: RSET UTXT$(BFLD) = SNX$(BRID)
501 WIN.MFMT = FMTBRID: CALL MsgFORMAT: IF A = 27 THEN 590
    X# = VAL(UTXT$(BFLD)): IF X# = 0 THEN 590
    IF (1 > X#) OR (X# > FF1HDR.MAXID) THEN CALL ErrBEEP(0): GOTO 501
    X = X#
    CALL FF1GetRec(X): IF NOT OKAY THEN GOTO 501
    BRID = X
    CALL RptOptGET("ADOP", X): AOK = (X = 1)
    AGLMAX = 0: DGLMAX = 0: DSRCH = 0
    IF WTYPE <> 2 THEN CALL RptOptGET("MXAL", AGLMAX)
    IF WTYPE > 1 THEN CALL RptOptGET("MXDL", DGLMAX)
    IF DGLMAX > 0 THEN
       CALL RptOptGET("ORDR", DSRCH)
       IF (DSRCH < 1) OR (DSRCH > 2) THEN DSRCH = 1
       CALL RptOptGET("DLNS", DSEXCHK)
       IF (DSEXCHK < 1) OR (DSEXCHK > 3) THEN DSEXCHK = 1
       DSEXCHK = DSEXCHK - 1
    END IF
    IF AGLMAX > 0 THEN
       CALL RptOptGET("ALNS", ASEXCHK)
       IF (ASEXCHK < 1) OR (ASEXCHK > 3) THEN ASEXCHK = 1
       ASEXCHK = ASEXCHK - 1
    END IF
    WFHDR.AMAXGL = AGLMAX: WFHDR.DMAXGL = DGLMAX
    WFHDR.AMAXLV = CHR$(0): WFHDR.DMAXLV = CHR$(0)
    WFHDR.HIRGL = CHR$(0): HIRGL = 0
    WFHDR.NUMANC = 0: WFHDR.NUMREL = 0
    GOSUB X80
    CALL RWRKOpen("NEW", WTYPE, 0): IF NOT OKAY THEN 590
    WFHDR.WFRULES = CHR$(DSRCH + 4 * ASEXCHK + 16 * DSEXCHK + 64 * ABS(AOK))
    WFHDR.TOTREC = 2
    PUT #5, 1, WFHDR
    WFANCREC.GACNT = 1
    WFANCREC.LINEAGE = 1: WFANCREC.ANCID = BRID
    WFANCREC.FIRSTGL = 0: WFANCREC.ANCCH = 0: WFANCREC.FIRSTANC = 0
    WFANCREC.PRVDUP = 0: WFANCREC.NXTDUP = 0
    PUT #5, , WFANCREC
    REDIM PTRTAB&(FF1HDR.MAXID)
    WFGLREC.ANCPTR = LEN(WFHDR): WFGLREC.AGL = 0: WFGLREC.DGL = 0
    WFGLREC.PRIOR = 0: WFGLREC.NEXT = 0
    WFGLREC.CLINES = 0: WFGLREC.TOTREC = 1
    HIWRKRBA& = LEN(WFHDR) + LEN(WFANCREC)

505 ' Create WFAGLREC
    IF WFGLREC.TOTREC = 0 THEN GOTO 540
    WFHDR.AMAXLV = CHR$(WFGLREC.AGL)
    ANCEND& = HIWRKRBA&
    HIWRKRBA& = HIWRKRBA& + LEN(WFGLREC)
    PUT #5, ANCEND& + 1, WFGLREC: WFHDR.TOTREC = WFHDR.TOTREC + 1
    IF WFGLREC.PRIOR <> 0 THEN
       X& = WFGLREC.PRIOR
       GET #5, X& + 1, WFGLREC: WFGLREC.NEXT = ANCEND&
       PUT #5, X& + 1, WFGLREC
       GET #5, ANCEND& + 1, WFGLREC
    END IF
    WFGLREC.PRIOR = ANCEND&
    ANCRBA& = WFGLREC.ANCPTR: WFGLREC.ANCPTR = 0
    WFGLREC.CLINES = 0: WFGLREC.TOTREC = 0: WFGLREC.NEXT = 0

510 ' Locate next generation of Ancestors
    WFGLREC.AGL = WFGLREC.AGL + 1
    IF WFGLREC.AGL > WFHDR.AMAXGL THEN 540
    SYSVAR.NVAR1 = WFGLREC.AGL
    CALL PutMSG("XXXXM065")
    A$ = INKEY$: IF A$ <> "" THEN IF ASC(A$) = 27 THEN 585
    GOSUB X80

520 ' Record parents of next Ancestor in this Generation
    GET #5, ANCRBA& + 1, WFANCREC
    LINEAGE# = WFANCREC.LINEAGE: FIRSTANC& = 0
    IF WFANCREC.ANCID < 0 THEN 530
    WFANCREC.ANCCH = ANCRBA&: ExtAdopt = ((WFANCREC.PSTAT AND 5) > 0)
    CALL FF1GetRec(WFANCREC.ANCID): C = 1
    IF (WFGLREC.AGL > 1) AND (ASEXCHK <> 0) AND (FF1REC.SEX <> FGENDR$(ASEXCHK)) THEN 530
    ANCID = FF1REC.FID: I = 0: S = ASC(FF1REC.STATUS) AND 15
    IF ((ANCID <> 0) AND (AOK OR ((S AND 1) = 0))) THEN GOSUB 535
    ANCID = FF1REC.MID: I = 1: S = S \ 4
    IF ((ANCID <> 0) AND (AOK OR ((S AND 1) = 0))) THEN GOSUB 535
530 ' GO TO Next Ancestor Record
    IF FIRSTANC& <> 0 THEN
       GET #5, ANCRBA& + 1, WFANCREC
       WFANCREC.FIRSTANC = FIRSTANC&
       PUT #5, ANCRBA& + 1, WFANCREC
    END IF
    ANCRBA& = ANCRBA& + LEN(WFANCREC)
    IF ANCRBA& >= ANCEND& THEN 505   ' finished this GEN, process Next GEN
    GOTO 520                        ' process Next ancestor, this GEN

535 ' Create Ancestor Record
    IF PTRTAB&(ANCID) = 0 THEN
       WFANCREC.ANCID = ANCID: WFHDR.NUMREL = WFHDR.NUMREL + 1
       ELSE
       WFANCREC.ANCID = -ANCID
    END IF
    WFANCREC.LINEAGE = 2 * LINEAGE# + I
    WFANCREC.PSTAT = ((S AND 3) - 4 * ExtAdopt)
    WFANCREC.PRVDUP = PTRTAB&(ANCID)
    WFHDR.NUMANC = WFHDR.NUMANC + 1
    IF C > 0 THEN WFGLREC.CLINES = WFGLREC.CLINES + 1: C = 0
    WFGLREC.TOTREC = WFGLREC.TOTREC + 1
    WFANCREC.GACNT = WFGLREC.TOTREC: WFANCREC.FIRSTANC = 0
    WFANCREC.NXTDUP = 0
    PUT #5, HIWRKRBA& + 1, WFANCREC: WFHDR.TOTREC = WFHDR.TOTREC + 1
    X& = PTRTAB&(ANCID)
    IF X& <> 0 THEN
       GET #5, X& + 1, WFANCREC
       WFANCREC.NXTDUP = HIWRKRBA&
       PUT #5, X& + 1, WFANCREC
       GET #5, HIWRKRBA& + 1, WFANCREC
    END IF
    PTRTAB&(ANCID) = HIWRKRBA&
    IF WFGLREC.ANCPTR = 0 THEN WFGLREC.ANCPTR = HIWRKRBA&
    IF (FIRSTANC& = 0) THEN FIRSTANC& = HIWRKRBA&
    HIWRKRBA& = HIWRKRBA& + LEN(WFANCREC)
    RETURN

540 ' Finished Locating Ancestors
    IF DGLMAX = 0 THEN 580
    WFAGLREC.NEXT = LEN(WFHDR) + LEN(WFANCREC): WFHDR.NUMREL = 0
    REDIM PTRTAB&(FF1HDR.MAXID)
    DIM GLS AS WRKGLStats
    IF WFHDR.AMAXGL > 0 THEN CALL PutMSG("XXXXM066")

550 ' Look for Descendants of Next Level of Ancestors
    IF WFAGLREC.NEXT = 0 THEN 580
    GOSUB X80
    GET #5, WFAGLREC.NEXT + 1, WFAGLREC
    AGL = -WFAGLREC.AGL: ANCNUM = 0
    ANCRBA& = WFAGLREC.ANCPTR - LEN(WFANCREC)

555 ' Locate Descendants of Next Ancestor
    ANCNUM = ANCNUM + 1
    IF ANCNUM > WFAGLREC.TOTREC THEN 550
    ANCRBA& = ANCRBA& + LEN(WFANCREC)
    GET #5, ANCRBA& + 1, WFANCREC
    S = (WFANCREC.PSTAT AND 3): ExtAdopt = ((WFANCREC.PSTAT AND 5) > 0)
    RELID = ABS(WFANCREC.ANCID)
    REDIM GLSTATS(WFHDR.DMAXGL + 1) AS WRKGLStats
    DGL = 0: PARPTR& = 0

560 ' Process Descendant RELID in Generation DGL
    GLS = GLSTATS(DGL)
    CALL FF1GetRec(RELID)
    IF (DGL > 0) THEN
       S = ASC(FF1REC.STATUS) AND 15
       IF (PARENT = FF1REC.MID) THEN S = S \ 4
       IF (NOT AOK) AND ((S AND 1) > 0) THEN FF1REC.OLDCH = 0: GOTO 566
    END IF

    ' First Create GLREC if this is First RELID in Generation DGL
    IF GLS.GLPTR = 0 THEN
       GLS.GLPTR = HIWRKRBA&
       IF DGL > 0 THEN
          X& = GLSTATS(DGL - 1).GLPTR
          GET #5, X& + 1, WFGLREC: WFGLREC.NEXT = GLS.GLPTR
          PUT #5, X& + 1, WFGLREC
          WFGLREC.PRIOR = GLSTATS(DGL - 1).GLPTR
          ELSE
          WFANCREC.FIRSTGL = GLS.GLPTR
          PUT #5, ANCRBA& + 1, WFANCREC
          WFGLREC.PRIOR = 0
       END IF
       WFGLREC.ANCPTR = ANCRBA&
       WFGLREC.AGL = AGL: WFGLREC.DGL = DGL
       WFGLREC.NEXT = 0
       WFGLREC.CLINES = 0: WFGLREC.TOTREC = 0
       PUT #5, HIWRKRBA& + 1, WFGLREC: WFHDR.TOTREC = WFHDR.TOTREC + 1
       HIWRKRBA& = HIWRKRBA& + LEN(WFGLREC)
    END IF

    ' Create Relative Record for RELID in DGL
    IF (DGL > 0) AND (GLS.CHNUM = 0) THEN
       GLS.CLINES = GLS.CLINES + 1
       GET #5, PARPTR& + 1, WFRELREC
       WFRELREC.OLDCH = HIWRKRBA&: ExtAdopt = ((WFRELREC.CSTAT AND 5) > 0)
       PUT #5, PARPTR& + 1, WFRELREC
    END IF
    IF GLS.RELREC <> 0 THEN
       GET #5, GLS.RELREC + 1, WFRELREC: WFRELREC.GLNEXT = HIWRKRBA&
       PUT #5, GLS.RELREC + 1, WFRELREC
    END IF
    GLS.RELID = RELID
    GLS.RELREC = HIWRKRBA&
    GLS.CHNUM = GLS.CHNUM + 1
    GLS.COUNT = GLS.COUNT + 1
    GLSTATS(DGL) = GLS
    WFRELREC.GLPTR = GLS.GLPTR
    WFRELREC.PARPTR = PARPTR&
    WFRELREC.ID = RELID
    WFRELREC.NREF = 0: WFRELREC.GLNEXT = 0
    WFRELREC.CHNUM = CHR$(GLS.CHNUM)
    WFRELREC.CSTAT = S - 4 * ExtAdopt
    WFRELREC.NUMCH = CHR$(0)
    IF PTRTAB&(RELID) = 0 THEN
       IF (DGL < WFHDR.DMAXGL) AND ((((DSEXCHK = 0) OR (FF1REC.SEX = FGENDR$(DSEXCHK))))) THEN
          WFRELREC.OLDCH = FF1REC.OLDCH
          ELSE
          WFRELREC.OLDCH = 0
       END IF
       ELSE
       WFRELREC.ID = -WFRELREC.ID
       WFRELREC.OLDCH = PTRTAB&(RELID)
       WFRELREC.NUMCH = CHR$(1)
    END IF
    PUT #5, HIWRKRBA& + 1, WFRELREC: WFHDR.TOTREC = WFHDR.TOTREC + 1
    IF PTRTAB&(RELID) = 0 THEN
       IF (AGL <> 0) OR (DGL <> 0) THEN WFHDR.NUMREL = WFHDR.NUMREL + 1
       ELSE
       X& = PTRTAB&(RELID)
       GET #5, X& + 1, WFRELREC: WFRELREC.NREF = HIWRKRBA&
       PUT #5, X& + 1, WFRELREC
       IF WFRELREC.ID < 0 THEN
          X = ASC(WFRELREC.NUMCH)
          GET #5, HIWRKRBA& + 1, WFRELREC
          WFRELREC.NUMCH = CHR$(X + 1)
          PUT #5, HIWRKRBA& + 1, WFRELREC
          ELSE
          GET #5, HIWRKRBA& + 1, WFRELREC
       END IF
    END IF
    PTRTAB&(RELID) = HIWRKRBA&
    HIWRKRBA& = HIWRKRBA& + LEN(WFRELREC)
    IF DGL > DMAXLV THEN DMAXLV = DGL: WFHDR.DMAXLV = CHR$(DGL)
    IF (AGL + DGL) > HIRGL THEN HIRGL = AGL + DGL: WFHDR.HIRGL = CHR$(HIRGL)

565 ' Select Next RELID to Process
    A$ = INKEY$
    IF A$ <> "" THEN IF ASC(A$) = 27 THEN 585 ELSE GOSUB X80
    IF (DSRCH = 2) OR (DGL = 0) THEN NEXTPAR& = GLS.RELREC: GOTO 567

566 ' Look for Younger Sibling of Child
    IF PARENT = FF1REC.FID THEN RELID = FF1REC.FCH ELSE RELID = FF1REC.MCH
    IF PARENT <> WFRELREC.ID THEN
       GET #5, PARPTR& + 1, WFRELREC
       ExtAdopt = ((WFRELREC.CSTAT AND 5) > 0)
    END IF
    IF RELID > 0 THEN 560
    'IF GLS.CHNUM > 0 THEN
       WFRELREC.NUMCH = CHR$(GLS.CHNUM)
       IF GLS.CHNUM = 0 THEN WFRELREC.OLDCH = 0
       PUT #5, PARPTR& + 1, WFRELREC
       GLS.CHNUM = 0
    'END IF
    GLSTATS(DGL) = GLS
    DGL = DGL - 1
    GLS = GLSTATS(DGL)
    IF (DSRCH = 2) THEN 569
    GET #5, GLS.RELREC + 1, WFRELREC
    NEXTPAR& = WFRELREC.GLNEXT

567 ' Look for a parent in DGL - FF1REC should be GLS.RELID Record
    WHILE NEXTPAR& <> 0
       IF GLS.RELREC <> NEXTPAR& THEN
          GLS.RELREC = NEXTPAR&
          GET #5, NEXTPAR& + 1, WFRELREC
       END IF
       IF WFRELREC.ID > 0 THEN
          RELID = WFRELREC.OLDCH
          IF RELID > 0 THEN
             PARENT = WFRELREC.ID: PARPTR& = GLS.RELREC
             GLS.RELID = PARENT
             GLSTATS(DGL) = GLS
             DGL = DGL + 1
             GLSTATS(DGL).CHNUM = 0
             GOTO 560
          END IF
       END IF
       NEXTPAR& = WFRELREC.GLNEXT
    WEND

    GLSTATS(DGL) = GLS
    IF (DSRCH = 2) THEN      ' Family Search
       IF DGL = 0 THEN 570   ' STOP - if we've returned to GL=0
       GOTO 566              ' ELSE Look for a Sibling of Childless Relative
    END IF

    ' Gen Search - Process Next Gen Level
    IF GLSTATS(DGL + 1).COUNT = 0 THEN 570
    DGL = DGL + 1: GLS = GLSTATS(DGL)
    NEXTPAR& = GLS.GLPTR + LEN(WFGLREC)  ' First RELREC in GL follows GLPTR
    GLS.RELREC = 0
    IF AGLMAX = 0 THEN
       GOSUB X80
       SYSVAR.NVAR1 = DGL
       CALL PutMSG("XXXXM067")
    END IF
    GOTO 567

569 ' FAM Search - Look for Sibling of Parent
    IF DGL < 1 THEN 570
    IF DGL = 1 THEN GOSUB X80
    PARPTR& = GLSTATS(DGL - 1).RELREC
    PARENT = GLSTATS(DGL - 1).RELID
    CALL FF1GetRec(GLS.RELID)
    GOTO 566

570 ' GEN Search - Update WFGLRECs for This Ancestor's Descendants
    DGL = 0: GLS = GLSTATS(0)
    WHILE GLS.GLPTR <> 0
       GET #5, GLS.GLPTR + 1, WFGLREC
       WFGLREC.CLINES = GLS.CLINES
       WFGLREC.TOTREC = GLS.COUNT
       PUT #5, GLS.GLPTR + 1, WFGLREC
       DGL = DGL + 1: GLS = GLSTATS(DGL)
       IF DGL > WFHDR.DMAXGL THEN GLS.GLPTR = 0
    WEND
    GOTO 555

580 ' WorkFile is Complete
    WFHDR.WFVER = CHR$(2): GOSUB X80
    CALL PutMSG("XXXXM068")

585 PUT #5, 1, WFHDR

590 ' Return to Calling Program
    ON ERROR GOTO 0
    CALL FamCLOSE
    EXIT SUB

595 ' Error Writing to Work File
    CALL ERRMessage("XXXXM032", ERR, ERL)
    RESUME 590

X80: ' Display Work File Information
    CURFMT = FMTSTAT: FMT = FMTTAB(CURFMT): X = FMT.BGNFLD
    WHILE X > 0
       FFLD = FLDTAB(X)
       IF FFLD.UTXT > 0 THEN
          X$ = UTXT$(FFLD.UTXT)
          IF WFHDR.NUMREL < 1 THEN
             SELECT CASE FFLD.NAME
                CASE "BRID": LSET X$ = STR$(BRID)
                CASE "NAME": LSET X$ = FMTNAME$(1)
                CASE "WFDT": LSET X$ = ""
                CASE "RULA": LSET X$ = SNX$(((ASC(WFHDR.WFRULES) AND 12) \ 4) + 1)
                CASE "RULD": LSET X$ = SNX$(((ASC(WFHDR.WFRULES) AND 48) \ 16) + 1)
                CASE "RULO": LSET X$ = SNX$(ASC(WFHDR.WFRULES) AND 3)
             END SELECT
          END IF
          SELECT CASE FFLD.NAME
             CASE "WFSZ": LSET X$ = STR$(WFHDR.TOTREC)
             CASE "HIAL": LSET X$ = STR$(ASC(WFHDR.AMAXLV))
             CASE "HIDL": LSET X$ = STR$(ASC(WFHDR.DMAXLV))
             CASE "NUMA": LSET X$ = STR$(WFHDR.NUMANC)
             CASE "NUMR": LSET X$ = STR$(WFHDR.NUMREL)
          END SELECT
          LSET UTXT$(FFLD.UTXT) = X$
       END IF
       X = FLDCHN(X).FWD
    WEND
    CALL WinDisplayDATA(0)
    RETURN

END SUB

REM $STATIC
SUB RWRKGet (OPT$, RC, PTR&)
    ' OPT$="NEXT" gets next Work Rec in standard sequence
    '     ="WRKREC" same as "NEXT" without getting NAME rec
    '     ="RELREC" gets Work Rec pointed to by PTR&
    '     ="PTRREC" same as "RELREC" without getting NAME rec
    OKAY = -1: IF (OPT$ <> "RELREC") AND (OPT$ <> "PTRREC") THEN PTR& = 0
    ON LOCAL ERROR GOTO RWRKGNErr
    SELECT CASE OPT$
       CASE "BASE": WFAGLREC.NEXT = LEN(WFHDR) + LEN(WFANCREC)
                    WFAGLREC.PRIOR = 0: WFAGLREC.TOTREC = 0
                    GOTO RWRKGNAnc
       CASE "NEXT":
       CASE "WRKREC":
       CASE "RELREC": GOTO RWRKGNPtr
       CASE "PTRREC": GOTO RWRKGNPtr
       CASE ELSE: GOTO RWRKGNBad
    END SELECT
    IF ASC(WFHDR.DMAXLV) > 0 THEN GOTO RWRKGNDesc ELSE GOTO RWRKGNAnc

RWRKGNPtr:
    X& = PTR&
    IF ASC(WFHDR.DMAXLV) = 0 THEN
       GET #5, X& + 1, WFANCREC
       Y& = X& - (WFANCREC.GACNT - 1) * LEN(WFANCREC) - LEN(WFAGLREC)
       GET #5, Y& + 1, WFAGLREC
       GET #5, WFAGLREC.NEXT + 1, WFAGLREC
       WFGLREC.AGL = -WFAGLREC.AGL: WFGLREC.DGL = 0
       WFRELREC.ID = WFANCREC.ANCID: WFRELREC.NREF = WFANCREC.NXTDUP
       WFRELREC.CSTAT = WFANCREC.PSTAT
       ELSE
       GET #5, X& + 1, WFRELREC: X& = WFRELREC.GLPTR
       IF OPT$ <> "PTRREC" THEN
          GET #5, X& + 1, WFGLREC: X& = WFGLREC.ANCPTR
          GET #5, X& + 1, WFANCREC
       END IF
    END IF
    RC = 0: GOTO RWRKGNEnd

RWRKGNAnc:
    IF WFAGLREC.PRIOR < WFAGLREC.TOTREC THEN
       PTR& = WFAGLREC.ANCPTR + WFAGLREC.PRIOR * LEN(WFANCREC)
       GET #5, PTR& + 1, WFANCREC
       WFAGLREC.PRIOR = WFAGLREC.PRIOR + 1
       RC = 2
       ELSE
       IF WFAGLREC.NEXT = 0 THEN
          RC = 0: GOTO RWRKGNEnd
          ELSE
          GET #5, WFAGLREC.NEXT + 1, WFAGLREC
          GET #5, WFAGLREC.ANCPTR + 1, WFANCREC
          PTR& = WFAGLREC.ANCPTR
          WFAGLREC.PRIOR = 1
          RC = 1
       END IF
    END IF
    IF ASC(WFHDR.DMAXLV) = 0 THEN
       WFGLREC.AGL = -WFAGLREC.AGL: WFGLREC.DGL = 0
       WFRELREC.ID = WFANCREC.ANCID: WFRELREC.NREF = WFANCREC.NXTDUP
       WFRELREC.CSTAT = WFANCREC.PSTAT
       ELSE
       X = ASC(WFHDR.DMAXLV) + 1
       REDIM GBL(X), GLNUMCH(X), PARPTR&(X)
       GET #5, WFANCREC.FIRSTGL + 1, WFGLREC
       GET #5, , WFRELREC
       PTR& = WFANCREC.FIRSTGL + LEN(WFGLREC)
       GLNUMCH(0) = ASC(WFRELREC.NUMCH)
    END IF
    GOTO RWRKGNEnd

RWRKGNDesc:
    IF (ASC(WFHDR.WFRULES) AND 1) THEN GOTO RWRKGNGRel
    IF (WFRELREC.ID < 0) OR (WFRELREC.OLDCH = 0) THEN GOTO RWRKGNSib
    PTR& = WFRELREC.OLDCH
    GET #5, WFGLREC.NEXT + 1, WFGLREC
    GET #5, WFRELREC.OLDCH + 1, WFRELREC
    IF GBL(WFGLREC.DGL) > 0 THEN RC = 4 ELSE RC = 3
    GBL(WFGLREC.DGL) = ASC(WFRELREC.CHNUM) + WFRELREC.CSTAT * 256
    GLNUMCH(WFGLREC.DGL) = ASC(WFRELREC.NUMCH)
    X = WFGLREC.DGL
    IF (X > 0) AND (PARPTR&(X) <> WFRELREC.PARPTR) THEN
       WHILE (X > 0) AND (PARPTR&(X) <> WFRELREC.PARPTR)
          PARPTR&(X) = WFRELREC.PARPTR
          GET #5, WFRELREC.PARPTR + 1, WFRELREC
          X = X - 1
          GBL(X) = ASC(WFRELREC.CHNUM) + WFRELREC.CSTAT * 256
          GLNUMCH(X) = ASC(WFRELREC.NUMCH)
       WEND
       GET #5, PTR& + 1, WFRELREC
       ELSE
       PARPTR&(X) = WFRELREC.PARPTR
    END IF
    GOTO RWRKGNEnd

RWRKGNSib:
    IF WFGLREC.DGL < 1 THEN GOTO RWRKGNAnc
    IF (GBL(WFGLREC.DGL) AND 255) < GLNUMCH(WFGLREC.DGL - 1) THEN GOTO RWRKGNGRel
    GBL(WFGLREC.DGL) = 0: GLNUMCH(WFGLREC.DGL) = 0: PARPTR&(WFGLREC.DGL) = 0
    GET #5, WFRELREC.PARPTR + 1, WFRELREC
    GET #5, WFRELREC.GLPTR + 1, WFGLREC
    GOTO RWRKGNSib

RWRKGNGRel:
    IF (WFRELREC.GLNEXT <> 0) THEN
       PTR& = WFRELREC.GLNEXT
       GET #5, PTR& + 1, WFRELREC
       'IF WFRELREC.ID < 0 THEN GOTO RWRKGNGRel
       IF GBL(WFGLREC.DGL) = 0 THEN RC = 3 ELSE RC = 4
       GBL(WFGLREC.DGL) = ASC(WFRELREC.CHNUM) + 256 * WFRELREC.CSTAT
       GLNUMCH(WFGLREC.DGL) = ASC(WFRELREC.NUMCH)
       IF (ASC(WFRELREC.CHNUM) > 1) THEN GOTO RWRKGNEnd
       X = WFGLREC.DGL
       IF (X > 0) AND (PARPTR&(X) <> WFRELREC.PARPTR) THEN
          WHILE (X > 0) AND (PARPTR&(X) <> WFRELREC.PARPTR)
             PARPTR&(X) = WFRELREC.PARPTR
             GET #5, WFRELREC.PARPTR + 1, WFRELREC
             X = X - 1
             GBL(X) = ASC(WFRELREC.CHNUM) + WFRELREC.CSTAT * 256
             GLNUMCH(X) = ASC(WFRELREC.NUMCH)
          WEND
          GET #5, PTR& + 1, WFRELREC
          ELSE
          PARPTR&(X) = WFRELREC.PARPTR
       END IF
       GOTO RWRKGNEnd
    END IF
    IF WFGLREC.NEXT = 0 THEN GOTO RWRKGNAnc
    WFRELREC.GLNEXT = WFGLREC.NEXT + LEN(WFGLREC)
    GET #5, WFGLREC.NEXT + 1, WFGLREC
    GOTO RWRKGNGRel

RWRKGNErr:
    CALL ERRMessage("XXXXM031", ERR, ERL)
    RESUME RWRKGNBad
RWRKGNBad:
    RC = 9: OKAY = 0
RWRKGNEnd:
    IF (RC > 0) AND (RC < 9) AND (OPT$ <> "WRKREC") AND (OPT$ <> "PTRREC") AND ENV.FFOPEN THEN
       IF WFRELREC.ID <> FF1REC.RID THEN CALL FF1GetRec(ABS(WFRELREC.ID))
    END IF
END SUB

SUB RWRKOpen (OPT$, XTYPE, DSNFMT)
    QUIET = (RIGHT$(OPT$, 1) = "Q"): X = LEN(RTRIM$(OPT$)) + QUIET
    XOPT$ = MID$(OPT$, 1, X)
    FTYPE = XTYPE
    IF FTYPE = 0 THEN FTYPE = -(ASC(WFHDR.AMAXLV) > 0) - 2 * (ASC(WFHDR.DMAXLV) > 0)
    IF ((FTYPE < 1) OR (FTYPE > 3)) THEN FTYPE = 3
    SELECT CASE FTYPE           ' Set File Name Table Index
       CASE 1: X = 5
       CASE 2: X = 6
       CASE 3: X = 7
       CASE ELSE: GOTO WFOpenBAD
    END SELECT
    IF NOT QUIET THEN
       IF DSNFMT <> 0 THEN
          WIN.MFMT = DSNFMT: CURFMT = DSNFMT: FMT = FMTTAB(CURFMT)
          ELSE
          CALL FmtFIND("XXXXS014")
       END IF
       CALL FmtFindFLD("DSNM"): Y = FFLD.UTXT: LSET UTXT$(Y) = FT$(X)
       CALL MsgFORMAT: IF A = 27 THEN GOTO WFOpenBAD
       IF XTYPE <> 0 THEN LSET FT$(X) = UTXT$(Y)
       DSNAME$ = RTRIM$(WRKFPFX$ + UTXT$(Y))
       ELSE
       DSNAME$ = RTRIM$(WRKFPFX$ + FT$(X))
    END IF
    LSET SYSVAR.FILENAME = DSNAME$
    ON LOCAL ERROR GOTO NOWF
CHKWF:
    OPEN SYSVAR.FILENAME FOR INPUT AS #5
    CLOSE #5
    IF XOPT$ = "NEW" THEN
       CALL PutMSG("XXXXM002")
       IF A$ <> SNGLKEY$(1) THEN GOTO WFOpenBAD
       KILL SYSVAR.FILENAME
       GOTO NEWWF
    END IF
    GOTO OLDWF
NOWF: X = ERR
    IF XOPT$ = "NEW" THEN RESUME NEWWF
    IF NOT QUIET THEN CALL ERRMessage("XXXXM030", X, 0)
    RESUME WFOpenBAD
OLDWF:
    ON LOCAL ERROR GOTO WFOpenERR
    SELECT CASE XOPT$
       CASE "READ": OPEN SYSVAR.FILENAME FOR BINARY ACCESS READ AS #5
       CASE "WRITE": OPEN SYSVAR.FILENAME FOR BINARY ACCESS READ WRITE AS #5
       CASE ELSE: GOTO WFOpenBAD
    END SELECT
    SYSVAR.MVAR = ".WRK"
    GET #5, , WFHDR
    IF (WFHDR.FTYPE <> "W") OR (WFHDR.WFTYPE <> CHR$(3)) OR (WFHDR.WFVER <> CHR$(2)) THEN
       IF NOT QUIET THEN CALL PutMSG("XXXXM033")
       GOTO WFOpenBAD
    END IF
    IF (WFHDR.REORGY <> FF1HDR.REORGY) OR (WFHDR.REORGMD <> FF1HDR.REORGMD) OR (WFHDR.REORGTM <> FF1HDR.REORGTM) THEN
       IF NOT QUIET THEN CALL PutMSG("XXXXM034")
       GOTO WFOpenBAD
    END IF
    GET #5, 1, WFHDR
    GET #5, , WFANCREC
    OKAY = -1
    GOTO WFOpenEND
NEWWF:
    ON LOCAL ERROR GOTO WFOpenERR
    CLOSE #5
    OPEN SYSVAR.FILENAME FOR BINARY ACCESS READ WRITE AS #5
    WFHDR.FTYPE = "W": WFHDR.WFTYPE = CHR$(3): WFHDR.WFVER = CHR$(255)
    WFHDR.REORGY = FF1HDR.REORGY
    WFHDR.REORGMD = FF1HDR.REORGMD
    WFHDR.REORGTM = FF1HDR.REORGTM
    WFHDR.UPDTY = VAL(RIGHT$(DATE$, 4))
    WFHDR.UPDTMD = 100 * (VAL(LEFT$(DATE$, 2))) + VAL(MID$(DATE$, 4, 2))
    PUT #5, 1, WFHDR
    OKAY = -1
    GOTO WFOpenEND
WFOpenERR:
    CALL ERRMessage("XXXXM030", ERR, ERL)
    RESUME WFOpenBAD
WFOpenBAD:
    OKAY = 0
WFOpenEND:
END SUB

FUNCTION SFILL$ (X, L)
   S$ = LTRIM$(STR$(FFRBA&(X, 1))): SL = LEN(S$)
   IF SL <= L THEN
      SFILL$ = STRING$(L, ASC(" "))
      MID$(SFILL$, L - SL + 1, SL) = S$
      ELSE
      SFILL$ = S$
   END IF
END FUNCTION

SUB SortStatACCUM
    IF SortPHASE <> 0 THEN
       SortSTAT!(SortPHASE) = SortSTAT!(SortPHASE) + TIMER - SortSTAT!(0)
    END IF
    SortSTAT!(0) = TIMER
END SUB

SUB SortStatSHOW
    CALL SortStatACCUM
    IF SortFMT <> 0 THEN
       IF (CURFMT <> SortFMT) THEN CURFMT = SortFMT: FMT = FMTTAB(SortFMT)
       X = FMT.BGNFLD
       WHILE X > 0
          FFLD = FLDTAB(X)
          IF FFLD.UTXT > 0 THEN
             X$ = UTXT$(FFLD.UTXT): XL = LEN(X$)
             SELECT CASE FFLD.NAME
                CASE "SS  ": RSET X$ = LTRIM$(STR$(SS))
                CASE "SEC1": RSET X$ = MID$(STR$(SortSTAT!(1)), 2, XL)
                CASE "SEC2": RSET X$ = MID$(STR$(SortSTAT!(2)), 2, XL)
                CASE "SEC3": RSET X$ = MID$(STR$(SortSTAT!(3)), 2, XL)
                CASE "SEC4": RSET X$ = MID$(STR$(SortSTAT!(4)), 2, XL)
                CASE "SEC5": RSET X$ = MID$(STR$(SortSTAT!(5)), 2, XL)
                CASE "CNT1": RSET X$ = LTRIM$(STR$(SortSTAT!(6)))
                CASE "CNT2": RSET X$ = LTRIM$(STR$(SortSTAT!(7)))
                CASE "CNT3": RSET X$ = LTRIM$(STR$(SortSTAT!(8)))
                CASE "CNT4": RSET X$ = LTRIM$(STR$(SortSTAT!(9)))
                CASE "CNTW": RSET X$ = LTRIM$(STR$(SortSTAT!(7) + SortSTAT!(8) + SortSTAT!(9)))
                CASE "CNT5": RSET X$ = LTRIM$(STR$(SortSTAT!(10)))
                CASE "TIM2": LSET X$ = TIME$
             END SELECT
             LSET UTXT$(FFLD.UTXT) = X$
          END IF
          X = FLDCHN(X).FWD
       WEND
       CALL WinDisplayDATA(0)
       SortSTAT!(0) = TIMER
    END IF
END SUB

REM $DYNAMIC
SUB SWRKBldList (IDX(), SMAX, LFMT)
   ' Select File = 1 Byte for Each ID in .NAM, <>CHR$(0) if selected
   ' IDX() = Array of Size FF1HDR.MAXID provided by Calling Progam
   ' SMAX=High Non Zero Element of IDX() generated by SWRKBldList
   ' LFMT=0 --> IDX() = Full ID list, <>0 if selected (value of Byte)
   ' LFMT=1 --> IDX() = Min ID List, = ID's Selected
   CALL SWRKOpen("READ")
   IF OKAY THEN
      DIM XRULE AS SlctFileRULE
      SRCNT = ASC(SlctHDR.SRULCNT)
      SEEK #6, LEN(SlctHDR) + SRCNT * LEN(XRULE) + 1
      Z$ = " ": X = 0: Y = 0: HI = 0: XMAX = SlctHDR.HIID
      IF XMAX > FF1HDR.MAXID THEN XMAX = FF1HDR.MAXID
      WHILE X < XMAX
         X = X + 1
         GET #6, , Z$: Z = ASC(Z$)
         SELECT CASE LFMT
            CASE 0: Y = Y + 1: IDX(Y) = Z: IF Z <> 0 THEN HI = Y
            CASE 1: IF Z <> 0 THEN Y = Y + 1: IDX(Y) = X
         END SELECT
      WEND
      IDX(0) = SlctHDR.LOID
      IF LFMT = 1 THEN SMAX = Y ELSE SMAX = HI
   END IF
   CALL FamCLOSE
END SUB

SUB SWRKCreate (IDX(), SMAX, LFMT)
   ' Used to Create SELECT File with NO Rules
   ' LFMT 0:IDX() Full Check List  1:IDX() Short ID List
   CALL SWRKOpen("NEW")
   IF OKAY THEN
      CALL PutMSG("XXXXM060")
      SlctHDR.SPGM = MID$(ENV.PGMNAME, 4, 4)
      SlctHDR.LOID = 32767: SlctHDR.HIID = 0
      SEEK #6, LEN(SlctHDR) + 1
      X = 0: Y = 0: Z0$ = CHR$(0): Z$ = " "
      WHILE X < SMAX
         X = X + 1: Z = IDX(X)
         SELECT CASE LFMT
            CASE 0: Z = Z AND 255: IF Z <> 0 THEN Y = X
            CASE 1: Y = Y + 1
                    WHILE Y < Z: Y = Y + 1: PUT #6, , Z0$: WEND
                    Z = 1
         END SELECT
         LSET Z$ = CHR$(Z): PUT #6, , Z$
         IF Z > 0 THEN
            SlctHDR.SLCTCNT = SlctHDR.SLCTCNT - ((Z AND 1) = 1)
            SlctHDR.SECSCNT = SlctHDR.SECSCNT - ((Z AND 3) = 2)
            IF Y < SlctHDR.LOID THEN SlctHDR.LOID = Y
            SlctHDR.HIID = Y
         END IF
      WEND
      SlctHDR.SLCTTOT = SlctHDR.SLCTCNT + SlctHDR.SECSCNT
      SlctHDR.WFVER = CHR$(2)
      PUT #6, 1, SlctHDR
   END IF
   CALL FamCLOSE
END SUB

SUB SWRKOpen (XOPT$)
    QUIET = (RIGHT$(XOPT$, 1) = "Q"): X = LEN(RTRIM$(XOPT$)) + QUIET
    OPT$ = MID$(XOPT$, 1, X)
    IF NOT QUIET THEN
       CALL FmtFIND("XXXXS013")
       CALL FmtFindFLD("DSNM"): LSET UTXT$(FFLD.UTXT) = FT$(10)
       CALL MsgFORMAT: IF A = 27 THEN GOTO SFOpenBAD
       CALL FmtFindFLD("DSNM"): LSET FT$(10) = UTXT$(FFLD.UTXT)
    END IF
    DSNAME$ = RTRIM$(WRKFPFX$ + FT$(10))
    LSET SYSVAR.FILENAME = DSNAME$
    ON LOCAL ERROR GOTO NOSF
CHKSF:
    OPEN DSNAME$ FOR INPUT AS #6
    CLOSE #6
    IF OPT$ = "NEW" THEN
       IF NOT QUIET THEN
          CALL PutMSG("XXXXM002")
          IF A$ <> SNGLKEY$(1) THEN GOTO SFOpenBAD
       END IF
       KILL SYSVAR.FILENAME
       GOTO NEWSF
    END IF
    GOTO OLDSF
NOSF:
    X = ERR
    IF OPT$ = "NEW" THEN RESUME NEWSF
    IF NOT QUIET THEN CALL ERRMessage("XXXXM030", X, 0)
    RESUME SFOpenBAD
OLDSF:
    ON LOCAL ERROR GOTO SFOpenERR
    SELECT CASE OPT$
       CASE "READ": OPEN SYSVAR.FILENAME FOR BINARY ACCESS READ AS #6
       CASE "WRITE": OPEN SYSVAR.FILENAME FOR BINARY ACCESS READ WRITE AS #6
       CASE ELSE: GOTO SFOpenBAD
    END SELECT
    SYSVAR.MVAR = "SELECT"
    GET #6, 1, SlctHDR
    IF (SlctHDR.FTYPE <> "W") OR (SlctHDR.WFTYPE <> CHR$(5)) OR (SlctHDR.WFVER <> CHR$(2)) THEN
       IF NOT QUIET THEN CALL PutMSG("XXXXM033")
       GOTO SFOpenBAD
    END IF
    IF (SlctHDR.REORGY <> FF1HDR.REORGY) OR (SlctHDR.REORGMD <> FF1HDR.REORGMD) OR (SlctHDR.REORGTM <> FF1HDR.REORGTM) THEN
       IF NOT QUIET THEN CALL PutMSG("XXXXM034")
       GOTO SFOpenBAD
    END IF
    OKAY = -1
    GOTO SFOpenEND
NEWSF:
    ON LOCAL ERROR GOTO SFOpenERR
    DIM NEWSHDR AS SlctFileHDR
    CLOSE #6
    OPEN SYSVAR.FILENAME FOR BINARY ACCESS READ WRITE AS #6
    NEWSHDR.FTYPE = "W": NEWSHDR.WFTYPE = CHR$(5): NEWSHDR.WFVER = CHR$(255)
    NEWSHDR.REORGY = FF1HDR.REORGY
    NEWSHDR.REORGMD = FF1HDR.REORGMD
    NEWSHDR.REORGTM = FF1HDR.REORGTM
    NEWSHDR.UPDTY = VAL(RIGHT$(DATE$, 4))
    NEWSHDR.UPDTMD = 100 * (VAL(LEFT$(DATE$, 2))) + VAL(MID$(DATE$, 4, 2))
    NEWSHDR.MAXID = FF1HDR.MAXID
    SlctHDR = NEWSHDR
    PUT #6, 1, SlctHDR
    OKAY = -1
    GOTO SFOpenEND
SFOpenERR:
    CALL ERRMessage("XXXXM030", ERR, ERL)
    RESUME SFOpenBAD
SFOpenBAD:
    OKAY = 0
SFOpenEND:
END SUB

REM $STATIC
FUNCTION TimeHHMM (X$)
   TimeHHMM = VAL(MID$(X$, 1, 2)) * 100 + VAL(MID$(X$, 4, 2))
END FUNCTION

FUNCTION UCX$ (X$)
   X = LEN(RTRIM$(X$)): Y = 0: Z = 0
   Y$ = SPACE$(100)
   WHILE Y < X
      Y = Y + 1: C = ASC(MID$(X$, Y, 1)): IF C = 0 THEN C = 32
      D = ASC(MID$(UCT1$, C, 1))
      IF D > 0 THEN
         MID$(Y$, Z + 1, 255) = UCTRAN$(D): Z = Z + LEN(UCTRAN$(D))
         ELSE
         Z = Z + 1: MID$(Y$, Z, 1) = UCASE$(CHR$(C))
      END IF
   WEND
   UCX$ = RTRIM$(Y$)
END FUNCTION

FUNCTION XRPTDate$ (MD, Y, O)
   X$ = SPACE$(11)
   IF (MD > 0) OR (Y > 0) THEN
      M = MD \ 100: D = MD MOD 100
      IF (MD = 0) AND (Y = 9999) THEN
         M = 13
         ELSE
         IF (M > 13) THEN M = 0
      END IF
      IF (0 < Y) AND (Y < 9999) THEN RSET X$ = STR$(Y)
      IF (MD > 0) OR (Y = 9999) THEN
         MID$(X$, 4, 3) = MID$(RMONTB$(M + 1), 3, 3)
      END IF
      IF D > 0 THEN MID$(X$, 1, 2) = SFILL$(D, 2)
   END IF
   IF O = 0 THEN
      XRPTDate$ = LTRIM$(RTRIM$(X$))
      ELSE
      XRPTDate$ = X$
   END IF
END FUNCTION

FUNCTION XSCRNDate$ (MD, Y)
   M = MD \ 100: D = MD MOD 100
   X$ = SPACE$(10):
   Y$ = RIGHT$("000" + LTRIM$(STR$(Y)), 4)
   M$ = RIGHT$("0" + LTRIM$(STR$(M)), 2)
   D$ = RIGHT$("0" + LTRIM$(STR$(D)), 2)
   SELECT CASE CONFIG.DFM
      CASE "1": LSET X$ = "  -  -": MID$(X$, 1, 2) = M$: MID$(X$, 4, 2) = D$
      CASE "2": LSET X$ = "  .  .": MID$(X$, 1, 2) = D$: MID$(X$, 4, 2) = M$
      CASE "3": LSET X$ = "    .  .": MID$(X$, 6, 2) = M$: MID$(X$, 9, 2) = D$
   END SELECT
   IF CONFIG.DFM = "3" THEN MID$(X$, 1, 4) = Y$ ELSE MID$(X$, 7, 4) = Y$
   XSCRNDate$ = X$
END FUNCTION

FUNCTION XSCRNTime$ (HHMM)
   HH$ = RIGHT$("0" + LTRIM$(STR$(HHMM \ 100)), 2)
   MM$ = RIGHT$("0" + LTRIM$(STR$(HHMM MOD 100)), 2)
   XSCRNTime$ = HH$ + ":" + MM$
END FUNCTION

FUNCTION ZFILL$ (X, L)
   Z$ = LTRIM$(STR$(FFRBA&(X, 1))): ZL = LEN(Z$)
   IF ZL <= L THEN
      ZFILL$ = STRING$(L, ASC("0"))
      MID$(ZFILL$, L - ZL + 1, ZL) = Z$
      ELSE
      ZFILL$ = Z$
   END IF
END FUNCTION

