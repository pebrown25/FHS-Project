DECLARE SUB MainSub ()
DECLARE SUB CHGReset ()
DECLARE SUB ReadNameDATA ()
DECLARE SUB SearchFile (OPT$)
DECLARE SUB SearchSELECT (TP AS ANY, TDATA$(), TCHN() AS ANY, SID%(), SMAX%)
DECLARE SUB ShowNameDATA ()
DECLARE SUB UpdtPLACE (XLOC%, SRTYPE%, SRNO%)
DECLARE SUB ViewADDRESS (ADRBGN%, SRTYPE%, SRNO%)
DECLARE SUB ViewCHILDREN (PID1%, PID2%)
DECLARE SUB ViewCOMMENTS (COMBGN%, SRTYPE%, SRNO%)
DECLARE SUB ViewEVENTS (EVNTBGN%, SRTYPE%, SRNO%)
DECLARE SUB ViewMARRIAGE (SPID%, MRNO%)
DECLARE SUB ViewMISC (ORTYPE%, BRID%)
DECLARE SUB ViewPARENTS (RID%)
DECLARE FUNCTION RELATION$ (RT$, RSEX$, AGL%, RGL%, RCODE$())
DECLARE FUNCTION FMTNAME$ (NFMT%)
DECLARE FUNCTION FDSTAT$ (ST$)
DECLARE FUNCTION FFSTAT$ (ST$, S$)
DECLARE FUNCTION SFILL$ (X%, L%)
DECLARE FUNCTION ZFILL$ (X%, L%)
DECLARE FUNCTION DateDif$ (D1MD%, D1Y%, D2MD%, D2Y%)
DECLARE FUNCTION DateMD% (X$, Y%)
DECLARE FUNCTION DateY% (X$, Y%)
DECLARE FUNCTION FFRBA& (RNUM%, RLTH%)
DECLARE FUNCTION FFRNUM% (RBA&, RLTH%)
DECLARE FUNCTION TimeHHMM% (X$)
DECLARE FUNCTION XSCRNDate$ (MD%, Y%)
DECLARE FUNCTION XSCRNTime$ (HHMM%)
   REM $INCLUDE: 'FHSCOMON.BAS'
   PN$ = " FAMILY HISTORY SYSTEM - FHSUPDT - File Update Program"
   CY$ = " (C) Copyright 1985-1998 by Phillip E. Brown"
   IF CONFIG.BP = 0 THEN RUN "FHSINIT"
   DFM = VAL(CONFIG.DFM)
   PGMRTRN$ = "FHSMENU": OPEN$ = "Open": NOTOPEN$ = "NOT Open"
   DIM OBPLACE AS STRING * 41, ODPLACE AS STRING * 41
   DIM OMPLACE AS STRING * 41, ODVPLACE AS STRING * 41
   DIM SRCHTable AS TablePARMS, CHLDTable AS TablePARMS
   SRCHTable.FMTNAME = "UPDTS005": SRCHTable.KEYNAME = "RID "
   CHLDTable.FMTNAME = "UPDTS009": CHLDTable.KEYNAME = "RID "
   REDIM SRCHID(0), SRCHLIST$(10), SRCHCHN(10) AS ChainPTRS
   REDIM CHLDLIST$(0), CHLDCHN(0) AS ChainPTRS
   CALL TableLOAD("FILESTAT", T, 13)
   IF OKAY THEN
      X = TABLE(T).FT
      WHILE X > 0
        Z = VAL(MID$(TABDATA$(X), 1, 1))
        IF Z = 1 THEN
           OPEN$ = RTRIM$(MID$(TABDATA$(X), 2, 12))
           ELSE
           IF Z = 2 THEN NOTOPEN$ = RTRIM$(MID$(TABDATA$(X), 2, 12))
        END IF
        X = TABDATACHN(X).FWD
      WEND
      CALL TableFREE("FILESTAT")
   END IF
   LORID = 1: HIRID = 32767
   CALL RptLOAD("SRCH")
   CALL RptLOAD("FSUM")
   CALL TableLOAD("EVENTS", TNUM, 0)
   ZERODATE$ = XSCRNDate$(0, 0)
   NINEDATE$ = XSCRNDate$(9999, 9999)
   ' Program INITialization
   CALL WinCLR
   CALL WinINIT("UL1"): CALL WinINIT("UR1")
   CALL WinINIT("LR1"): CALL WinINIT("LL1"): CALL WinINIT("LH1")
   CALL FmtFIND("UPDTS001"): FMT001 = CURFMT
   CALL FmtFIND("UPDTS02A"): FMT02A = CURFMT
   CALL FmtFIND("UPDTS003"): FMT003 = CURFMT
   CALL FmtFindFLD("LOBD"): LSET DTXT$(WIN.CURFLD) = ZERODATE$
   CALL FmtFindFLD("HIBD"): LSET DTXT$(WIN.CURFLD) = NINEDATE$
   CALL FmtFindFLD("LODD"): LSET DTXT$(WIN.CURFLD) = ZERODATE$
   CALL FmtFindFLD("HIDD"): LSET DTXT$(WIN.CURFLD) = NINEDATE$
   CALL FMTDREST
   CALL FmtFIND("UPDTS004"): FMT004 = CURFMT
   CALL FmtFIND("UPDTS005"): FMT005 = CURFMT
   CALL FmtFIND("UPDTS006"): FMT006 = CURFMT
   CALL FmtFindFLD("BDT"): LSET DTXT$(WIN.CURFLD) = ZERODATE$
   CALL FmtFindFLD("DDT"): LSET DTXT$(WIN.CURFLD) = ZERODATE$
   CALL FmtFIND("UPDTS007"): FMT007 = CURFMT
   CALL FmtFIND("UPDTS009"): FMT009 = CURFMT
   CALL FmtFIND("UPDTS900"): FMT900 = CURFMT
   CALL FmtFIND("UPDTS901"): FMT901 = CURFMT
   CALL FmtFIND("UPDTS902"): FMT902 = CURFMT
   CALL FmtFIND("UPDTS903"): FMT903 = CURFMT
   CALL FmtFIND("UPDTS904"): FMT904 = CURFMT
   CALL FmtFIND("UPDTS905"): FMT905 = CURFMT
   CALL FmtFIND("UPDTS906"): FMT906 = CURFMT
   CALL FmtFIND("UPDTS907"): FMT907 = CURFMT
   CALL FmtFIND("UPDTS911"): FMT911 = CURFMT
   CALL WinSWITCH(2)
   CALL WinFORMAT(FMT001)
   WIN.OFMT = FMT001
   CALL OptHILITE("F1", "F1", 6)
   CHKLVL = -1

100 ' Process FAMILY File Open/Close/INIT
   CALL WinSWITCH(3)
   IF WIN.WFMT <> FMT02A THEN
      CALL WinCLR
      CALL WinDisplayLABELS(FMT02A)
   END IF
   CALL OptDISPLAY(FMT900)
   IF NOT ENV.FFOPEN THEN CALL FamOPEN("VLDTQ")
   CALL PutMSG("")
101 ' Process File Info
   FMT = FMTTAB(FMT02A)
   IF ENV.FFOPEN THEN
      FSTAT$ = OPEN$: AFMT = FF2HDR.RevLVL AND 255
      ELSE
      FSTAT$ = NOTOPEN$: AFMT = 0
      CALL FMTDREST
   END IF
   FF = FMT.BGNFLD
   WHILE FF > 0
      FFLD = FLDTAB(FF)
      IF FFLD.UTXT > 0 THEN
         X$ = UTXT$(FFLD.UTXT)
         SELECT CASE FFLD.NAME
           CASE "FSET": LSET X$ = FDFSetup.NAME
           CASE "DESC": LSET X$ = FDFSetup.DESC
           CASE "NAME": LSET X$ = FTYPE$(1)
           CASE "ADDR": LSET X$ = FTYPE$(2)
           CASE "OTHR": LSET X$ = FTYPE$(3)
           CASE "STAT": LSET X$ = FSTAT$
         END SELECT
         IF ENV.FFOPEN THEN
           SELECT CASE FFLD.NAME
              CASE "NVER": RSET X$ = LTRIM$(STR$(FF1HDR.RevLVL AND 255))
              CASE "NLUP": LSET X$ = XSCRNDate$(FF1HDR.UPDTMD, FF1HDR.UPDTY)
              CASE "NREC": X& = FFRBA&(FF1HDR.MAXID, 1): RSET X$ = LTRIM$(STR$(X&))
              CASE "AVER": RSET X$ = LTRIM$(STR$(FF2HDR.RevLVL AND 255))
              CASE "ALUP": LSET X$ = XSCRNDate$(FF2HDR.UPDTMD, FF2HDR.UPDTY)
              CASE "AREC": X& = FFRBA&(FF2HDR.HIREC, 1): RSET X$ = LTRIM$(STR$(X&))
              CASE "AFRE": X& = FFRBA&(FF2HDR.FREECNT, 1): RSET X$ = LTRIM$(STR$(X&))
              CASE "OVER": RSET X$ = LTRIM$(STR$(FF3HDR.RevLVL AND 255))
              CASE "OLUP": LSET X$ = XSCRNDate$(FF3HDR.UPDTMD, FF3HDR.UPDTY)
              CASE "OREC": X& = FFRBA&(FF3HDR.HIREC, 1): RSET X$ = LTRIM$(STR$(X&))
              CASE "OFRE": X& = FFRBA&(FF3HDR.FREECNT, 1): RSET X$ = LTRIM$(STR$(X&))
           END SELECT
         END IF
         LSET UTXT$(FFLD.UTXT) = X$
      END IF
      FF = FLDCHN(FF).FWD
   WEND
   CALL WinDisplayDATA(FMT02A)
   IF ENV.FFOPEN AND CHKLVL THEN
      IF (FF1HDR.RevLVL > 255) OR (FF2HDR.RevLVL > 255) OR (FF3HDR.RevLVL > 255) THEN
         CALL PutMSG("XXXXM073")
      END IF
      CHKLVL = 0
      IF FUN$ = "" THEN GOTO 195
   END IF
   CALL GetKEY
   SELECT CASE A
      CASE 27: CALL WinCLR: GOTO 195
      CASE 59: CALL OptHILITE("F1", "", 6)
               CALL WinSWITCH(4)
               CALL FamCLOSE: CALL FDFSelect
               LORID = 1: HIRID = 32767: CHKLVL = -1
      CASE 60: CALL FamCLOSE: GOTO 101
      CASE 61: CALL OptHILITE("F3", "", 6)
               CALL FamOPEN("INIT")
               IF NOT OKAY THEN 100
               LORID = 1: HIRID = 32767: FUN$ = ""
      CASE 64: CALL OptHILITE("F6", "", 6)
               CALL FSUMReport
               IF PRT.OPT = 1 THEN
                  CALL WinPREP(6): WIN.WFMT = 0
                  CALL WinSWITCH(2): CALL WinFORMAT(FMT001)
                  CALL WinSWITCH(3): WIN.WFMT = 0
               END IF
      CASE 94: PGMRTRN$ = "FHSFILE": GOTO 49000
      CASE 109: CALL OptHILITE("F6", "", 6): CALL FamCLOSE
               CALL WinSWITCH(4)  ' ALT+F6
               CALL RptOptDISPLAY("FSUM")
               CALL RptOptUPDATE
               CALL WinCLR
               CALL WinSWITCH(3)
   END SELECT
   CALL PutMSG("")
   GOTO 100

195 ' Show Change DATE
    CURFMT = FMT001: FMT = FMTTAB(CURFMT)
    CALL FmtFindFLD("CHDT")
    IF FOUND THEN
       IF ENV.FFOPEN THEN
          LSET UTXT$(FFLD.UTXT) = XSCRNDate$(FF1HDR.CHGDTMD, FF1HDR.CHGDTY)
          ELSE
          LSET UTXT$(FFLD.UTXT) = ""
       END IF
    END IF

200 ' Main Menu Display
   CALL WinSWITCH(2): RGL = 0
   IF WIN.WFMT <> FMT001 THEN CALL WinCLR
   CALL WinFORMAT(FMT001)
   WIN.OFMT = FMT001
201 ' Get Next REQUEST
   CALL WinUPDATE
   SELECT CASE A
      CASE 13: FUN$ = "F3"
      CASE 59: FUN$ = "F1"
      CASE 60: FUN$ = "F2"
      CASE 61: FUN$ = "F3"
      CASE 62: FUN$ = "F4"
      CASE 63: FUN$ = "F5"
      CASE 64: FUN$ = "F6"
      CASE 65: FUN$ = "F7"
      CASE 67: FUN$ = "F9"
      CASE ELSE: CALL ErrBEEP(0): GOTO 201
   END SELECT
   CALL OptHILITE(FUN$, FUN$, 6)
   SELECT CASE FUN$
      CASE "F1": GOTO 100
      CASE "F2": IF ENV.FFOPEN THEN BRID = 0: GOTO 500
      CASE "F3": IF ENV.FFOPEN THEN CALL FmtFindFLD("RID")
                 X& = VAL(UTXT$(FFLD.UTXT)): CALL ChkID(X&)
                 IF OKAY THEN BRID = X&: GOTO 600:  ELSE GOTO 200
      CASE "F4": IF ENV.FFOPEN THEN SOPT$ = "SEARCH": GOTO 400
      CASE "F5": IF ENV.FFOPEN THEN CALL RelCALC(0): GOTO 200
      CASE "F6": IF ENV.FFOPEN THEN CALL CHGReset: GOTO 195
      CASE "F7": CALL FmtFindFLD("RID"): SHELL "DATECALC": GOTO 200
      CASE "F9": GOTO 49000
      CASE ELSE: CALL ErrBEEP(0): GOTO 200
   END SELECT
   CALL PutMSG("UPDTM016")
   GOTO 200

400 ' Search FILE
   CALL SearchFile(SOPT$)
   IF SMAX = 0 THEN GOTO 200
   IF BRID > 0 THEN GOTO 600
   GOTO 200

500 ' ADD NAME Record
   CALL WinSWITCH(3)
   IF WIN.WFMT <> FMT006 THEN
      CALL WinCLR: CALL WinDisplayLABELS(FMT006)
   END IF
   CALL OptDISPLAY(FMT904)
   CURFMT = WIN.WFMT: FMT = FMTTAB(CURFMT)
   CALL FMTDREST: NRID = 0: CFID = 0: CMID = 0
   CALL FmtFindFLD("RID"): RSET UTXT$(FFLD.UTXT) = LTRIM$(STR$(FF1HDR.MAXID + 1))
   CALL FmtFindFLD("RGL"): RSET UTXT$(FFLD.UTXT) = LTRIM$(STR$(RGL))
   CALL WinDisplayDATA(0)
501 WIN.CLIN = 0: CALL WinUPDATE
   IF (A = 27) OR (A = 60) THEN
      IF NRID > 0 THEN FF1HDR.MAXID = FF1HDR.MAXID - 1
      GOTO 600
   END IF
   IF A <> 59 THEN CALL ErrBEEP(0): GOTO 501
   IF NRID = 0 THEN
      CALL FF1GetNew: IF NOT OKAY THEN GOTO 600
      NRID = FF1REC.RID
      LSET OBPLACE = "": LSET ODPLACE = ""
   END IF
   NameREC = FF1REC
   GOSUB 660: IF NOT OKAY THEN GOTO 501
   BRID = NRID

600 ' Display NAME Record Number BRID
   IF BRID = 0 THEN CALL ErrBEEP(0): GOTO 200
   CALL FF1GetRec(BRID): IF NOT OKAY THEN GOTO 200
   NameREC = FF1REC
   LSET BPLACE = "": LSET DPLACE = ""
   IF FF1REC.BLOC <> 0 THEN
      CALL FF3GetRec(FF1REC.BLOC)
      IF FF3PFX.RTYPE = CHR$(12) THEN
         IF FF3PLACE2.PNEXT <> 0 THEN CALL FF3GetRec(FF3PLACE2.PNEXT)
      END IF
   END IF
   LSET OBPLACE = BPLACE: LSET ODPLACE = DPLACE
   CALL WinSWITCH(3)
   IF WIN.WFMT <> FMT006 THEN
      CALL WinCLR: CALL WinDisplayLABELS(FMT006)
   END IF
   CALL ShowNameDATA
601 CALL OptDISPLAY(FMT902)
   IF (FF1REC.FID <> 0) OR (FF1REC.MID <> 0) THEN CALL OptHILITE("F2", "", 7)
   IF FF1REC.OLDCH <> 0 THEN CALL OptHILITE("F3", "", 7)
   IF FF1REC.SPOUSE <> 0 THEN CALL OptHILITE("F4", "", 7)
   IF (FF1REC.ADRS <> 0) OR (FF1REC.COM <> 0) OR (FF1REC.EVENT <> 0) OR (FF1REC.EDUC <> 0) OR (FF1REC.WORK <> 0) OR (FF1REC.MIL <> 0) OR (FF1REC.HLTH <> 0) THEN
      CALL OptHILITE("F5", "", 7)
   END IF
   CALL WinDisplayDATA(FMT006)
   CALL GetKEY: FUN$ = ""
   SELECT CASE A
      CASE 27: IF CMAX > 0 THEN GOTO 800
               IF SMAX > 0 THEN SOPT$ = "LIST": GOTO 400
               GOTO 200
      CASE 59: FUN$ = "F1"
      CASE 60: IF (FF1REC.FID <> 0) OR (FF1REC.MID <> 0) THEN FUN$ = "F2"
      CASE 61: IF (FF1REC.OLDCH <> 0) THEN FUN$ = "F3"
      CASE 62: FUN$ = "F4"
      CASE 63: FUN$ = "F5"
   END SELECT
   IF FUN$ = "" THEN CALL ErrBEEP(0): GOTO 601
   CALL OptHILITE(FUN$, "", 6)
   SELECT CASE FUN$
      CASE "F1": CALL OptDISPLAY(FMT903)
                 CALL GetKEY
                 SELECT CASE A
                    CASE 59: GOTO 610
                    CASE 60: GOTO 500
                    CASE ELSE: CALL ErrBEEP(0): GOTO 601
                 END SELECT
      CASE "F2": GOTO 700
      CASE "F3": CMAX = 0: PID1 = BRID: PID2 = 0: CSRC = 1: GOTO 800
      CASE "F4": CMRNO = FF1REC.SPOUSE: GOTO 900
      CASE "F5": GOTO 603
      CASE ELSE: CALL ErrBEEP(0): GOTO 601
   END SELECT

603 ' More Name Record Options
   CALL OptDISPLAY(FMT911)
   IF FF1REC.ADRS <> 0 THEN CALL OptHILITE("F2", "", 7)
   IF FF1REC.COM <> 0 THEN CALL OptHILITE("F3", "", 7)
   IF FF1REC.EVENT <> 0 THEN CALL OptHILITE("F4", "", 7)
   IF (FF1REC.EDUC <> 0) OR (FF1REC.WORK <> 0) OR (FF1REC.MIL <> 0) OR (FF1REC.HLTH <> 0) THEN
      CALL OptHILITE("F5", "", 7)
   END IF
604 CALL GetKEY
   SELECT CASE A
      CASE 27: GOTO 601
      CASE 59: FUN$ = "F1"
      CASE 60: FUN$ = "F2"
      CASE 61: FUN$ = "F3"
      CASE 62: FUN$ = "F4"
      CASE 63: FUN$ = "F5"
      CASE ELSE: CALL ErrBEEP(0): GOTO 604
   END SELECT
   CALL OptHILITE(FUN$, "", 6)
   SELECT CASE FUN$
      CASE "F1": CALL RelCALC(BRID): CALL WinSWITCH(3): CALL FF1GetRec(BRID)
      CASE "F2": ADRBGN = FF1REC.ADRS
                 CALL ViewADDRESS(ADRBGN, 1, BRID)
                 CALL WinSWITCH(3)
                 IF XCHANGED THEN
                    FF1REC.ADRS = ADRBGN
                    CALL FF1PutREC
                 END IF
      CASE "F3": COMBGN = FF1REC.COM
                 CALL ViewCOMMENTS(COMBGN, 1, BRID)
                 CALL WinSWITCH(3)
                 IF XCHANGED THEN
                    FF1REC.COM = COMBGN
                    CALL FF1PutREC
                 END IF
      CASE "F4": EVNTBGN = FF1REC.EVENT
                 CALL ViewEVENTS(EVNTBGN, 1, BRID)
                 CALL WinSWITCH(3)
                 IF XCHANGED THEN
                    FF1REC.EVENT = EVNTBGN
                    CALL FF1PutREC
                 END IF
      CASE "F5": GOTO 605
   END SELECT
   GOTO 603

605 ' Still More Name Record Options
   CALL OptDISPLAY(FMT905)
   IF FF1REC.EDUC <> 0 THEN CALL OptHILITE("F1", "", 7)
   IF FF1REC.WORK <> 0 THEN CALL OptHILITE("F2", "", 7)
   IF FF1REC.MIL <> 0 THEN CALL OptHILITE("F3", "", 7)
   IF FF1REC.HLTH <> 0 THEN CALL OptHILITE("F4", "", 7)
606 CALL GetKEY
   SELECT CASE A
      CASE 27: GOTO 603
      CASE 59: FUN$ = "F1"
      CASE 60: FUN$ = "F2"
      CASE 61: FUN$ = "F3"
      CASE 62: FUN$ = "F4"
      CASE ELSE: CALL ErrBEEP(0): GOTO 606
   END SELECT
   CALL OptHILITE(FUN$, "", 6)
   SELECT CASE FUN$
      CASE "F1": CALL ViewMISC(7, BRID)
      CASE "F2": CALL ViewMISC(8, BRID)
      CASE "F3": CALL ViewMISC(9, BRID)
      CASE "F4": CALL ViewMISC(10, BRID)
   END SELECT
   GOTO 605

610 ' Update NAME Record
    CALL OptDISPLAY(FMT904)
    CURFMT = FMT006: FMT = FMTTAB(CURFMT)
611 WIN.CLIN = 0: CALL WinUPDATE
    IF (A = 27) OR (A = 60) THEN
       CALL FMTHREST
       FF1REC = NameREC: LSET BPLACE = OBPLACE: LSET DPLACE = ODPLACE
       GOTO 601
    END IF
    IF A <> 59 THEN CALL ErrBEEP(0): GOTO 610
    GOSUB 660: IF NOT OKAY THEN GOTO 611
    GOTO 601

660 ' Get Name Record DATA from Screen and Update File
   CALL ReadNameDATA
   IF (UFID! <> FF1REC.FID) OR (UMID! <> FF1REC.MID) OR ((UFID! > 0) AND (UFID! = UMID!)) THEN
      CALL PutMSG("UPDTM003"): OKAY = 0
      ELSE
      NameREC = FF1REC: UFID = UFID!: UMID = UMID!
      CALL FF1GetRec(FF1REC.RID): CFID = FF1REC.FID: CMID = FF1REC.MID
      BirthCHG = ((NameREC.BY <> FF1REC.BY) OR (NameREC.BMD <> FF1REC.BMD))
      IF (BirthCHG OR (CFID <> UFID)) AND (CFID > 0) THEN
         CALL UpdtRelPTRS("REM", CFID)
      END IF
      IF (BirthCHG OR (CMID <> UMID)) AND (CMID > 0) THEN
         CALL UpdtRelPTRS("REM", CMID)
      END IF
      IF (BirthCHG OR (CFID <> UFID)) AND (UFID > 0) THEN
         CALL UpdtRelPTRS("ADD", UFID)
      END IF
      IF (BirthCHG OR (CMID <> UMID)) AND (UMID > 0) THEN
         CALL UpdtRelPTRS("ADD", UMID)
      END IF
      CALL UpdtPLACE(NameREC.BLOC, 1, NameREC.RID)
      FF1REC = NameREC
      CALL FF1PutREC
      CALL ShowNameDATA
   END IF
   RETURN

700 ' Display PARENTs
   CALL ViewPARENTS(BRID)
   IF FUN$ <> "CHLD" THEN GOTO 600
   PCRID = BRID: PID1 = FF1REC.FID: PID2 = FF1REC.MID
   CSRC = 2: CMAX = 0

800 ' Display CHILDREN
   CALL ViewCHILDREN(PID1, PID2)
   IF FUN$ = "SEL" THEN BRID = CRID: GOTO 600
   ' ESCape From ViewCHILDREN -
   '   Return to ParentVIEW, MarriageVIEW or NameVIEW
   SELECT CASE CSRC
      CASE 1: BRID = PID1: GOTO 600
      CASE 2: BRID = PCRID: GOTO 700
      CASE 3: BRID = PID1: CMRNO = CPMAR: GOTO 900
   END SELECT
   GOTO 600

900 ' Process Family Records
   CALL ViewMARRIAGE(BRID, CMRNO)
   SELECT CASE FUN$
      CASE "SEL": GOTO 600
      CASE "CHLD": PID1 = BRID: CPMAR = CMRNO
                   IF BRID = FF3SPOUSE.SP1ID THEN
                      PID2 = FF3SPOUSE.SP2ID
                      ELSE
                      PID2 = FF3SPOUSE.SP1ID
                   END IF
                   CSRC = 3: CMAX = 0: GOTO 800
      CASE ELSE: GOTO 600
   END SELECT

49000 ' Program Termination
   CALL FamCLOSE: A = 0: CALL WinSWITCH(1)
   CALL PgmPREP(PGMRTRN$)
   IF (NOT OKAY) OR (A = 27) THEN PGMRTRN$ = "FHSMENU": GOTO 200
   CALL RptFREE("SRCH")
   CALL RptFREE("FSUM")
   CALL TableFREE("EVENTS", TNUM, 0)
   CHAIN PGMRTRN$

REM $DYNAMIC
SUB CHGReset
    IF ENV.FFOPEN THEN
       CALL PutMSG("UPDTM015")
       IF A$ <> SNGLKEY$(1) THEN EXIT SUB
       ID& = 0
       WHILE ID& < FF1HDR.MAXID
          ID& = ID& + 1: X& = 100 * ID& + 1
          GET #1, X&, FF1REC: Y = ASC(FF1REC.STATUS2)
          IF (Y AND 96) > 0 THEN
             FF1REC.STATUS2 = CHR$(Y AND 159)
             PUT #1, X&, FF1REC
          END IF
       WEND
       FF1HDR.CHGDTY = VAL(MID$(DATE$, 7, 4))
       FF1HDR.CHGDTMD = 100 * VAL(MID$(DATE$, 1, 2)) + VAL(MID$(DATE$, 4, 2))
       FF1HDR.CHGCNT1 = 0: FF1HDR.CHGCNT2 = 0
       FF1HDR.RevLVL = 256 + (FF1HDR.RevLVL AND 255)
       FFCHANGED = -1
    END IF
END SUB

REM $STATIC
SUB ReadNameDATA
SHARED FMT006, UFID!, UMID!

  ' Get Name Data From SCREEN
   CURFMT = FMT006: FMT = FMTTAB(CURFMT): OKAY = 0
   DFM = VAL(CONFIG.DFM)
   CALL GetUFLD("FIRST"): FF = WIN.CURFLD
   WHILE WIN.CURFLD > 0
      X$ = UTXT$(FFLD.UTXT)
      SELECT CASE FFLD.NAME
         CASE "SRNM": LSET FF1REC.SURNM = X$
         CASE "SNMU": IF X$ = SNGLKEY$(1) THEN X$ = "Y"
                      FF1REC.STATUS2 = FFSTAT$("SNU", X$)
         CASE "GVNM": LSET FF1REC.GIVEN = X$
         CASE "BDT ": FF1REC.BMD = DateMD(X$, DFM): FF1REC.BY = DateY(X$, DFM)
         CASE "BDST": FF1REC.STATUS = FFSTAT$("BD", X$)
         CASE "BTM ": FF1REC.BTM = TimeHHMM(X$)
         CASE "BPL ": LSET BPLACE = X$
         CASE "DDT ": FF1REC.DMD = DateMD(X$, DFM): FF1REC.DY = DateY(X$, DFM)
         CASE "DDST": FF1REC.STATUS = FFSTAT$("DD", X$)
         CASE "DTM ": FF1REC.DTM = TimeHHMM(X$)
         CASE "DPL ": LSET DPLACE = X$
         CASE "SEX ": LSET FF1REC.SEX = X$
         CASE "FID ": UFID! = VAL(X$)
         CASE "FST ": FF1REC.STATUS = FFSTAT$("FID", X$)
         CASE "MID ": UMID! = VAL(X$)
         CASE "MST ": FF1REC.STATUS = FFSTAT$("MID", X$)
      END SELECT
      CALL GetUFLD("NEXT")
      IF WIN.CURFLD = FF THEN WIN.CURFLD = 0
   WEND
   IF NOT ((UFID! > FF1HDR.MAXID) OR (UFID! = FF1REC.RID)) THEN
      FF1REC.FID = UFID!
   END IF
   IF NOT ((UMID! > FF1HDR.MAXID) OR (UMID! = FF1REC.RID)) THEN
      FF1REC.MID = UMID!
   END IF
   CALL FmtFindFLD("AGE")
   LSET UTXT$(FFLD.UTXT) = DateDif$(FF1REC.BMD, FF1REC.BY, FF1REC.DMD, FF1REC.DY)

END SUB

REM $DYNAMIC
SUB SearchFile (OPT$)
SHARED SRCHTable AS TablePARMS, SRCHLIST$(), SRCHCHN() AS ChainPTRS
SHARED SRCHID(), SMAX, BRID, LORID, HIRID, BGNID
SHARED FMT003, FMT004, FMT901, FMT907
   DFM = VAL(CONFIG.DFM)
   IF OPT$ = "LIST" THEN GOTO SearchLIST
SEARCHSetup: ' File SEARCH
   CALL WinSWITCH(2)
   IF WIN.WFMT <> FMT003 THEN
      CALL WinCLR
      CALL WinDisplayLABELS(FMT003)
   END IF
   CALL FmtFindFLD("LOID")
   IF FOUND THEN
      IF (LORID < FF1HDR.MAXID) THEN
         RSET UTXT$(FFLD.UTXT) = LTRIM$(STR$(LORID))
      END IF
   END IF
   CALL FmtFindFLD("HIID")
   IF (FF1HDR.MAXID < HIRID) THEN
      RSET UTXT$(FFLD.UTXT) = LTRIM$(STR$(FF1HDR.MAXID))
   END IF
   REDIM SRCHID(FF1HDR.MAXID)
   SMAX = 0
402 CALL OptDISPLAY(FMT901)
   CALL WinDisplayDATA(FMT003)
   WIN.CLIN = 0: CALL WinUPDATE
   SELECT CASE A
      CASE 27: SMAX = 0: EXIT SUB
      CASE 59: FUN$ = "F1"
      CASE 60: CALL FMTDREST: LORID = 32767: HIRID = 32767: GOTO SEARCHSetup
      CASE 61: FUN$ = "F3"
      CASE ELSE: CALL ErrBEEP(0): GOTO 402
   END SELECT
   CALL OptHILITE(FUN$, "", 6)
   IF FUN$ = "F3" THEN
      REDIM SRCHID(FF1HDR.MAXID)
      CALL SWRKBldList(SRCHID(), SMAX, 1): BOKAY = (OKAY AND (A <> 27))
      CALL FamOPEN("VLDT"): IF NOT BOKAY THEN GOTO 402
      CALL FMTDREST: CALL WinDisplayDATA(FMT003)
      BGNID = 1: LORID = SlctHDR.MAXID + 1: HIRID = SlctHDR.MAXID
      GOTO 403
   END IF
   FF = FMTTAB(FMT003).BGNFLD
   WHILE FF > 0
      FFLD = FLDTAB(FF)
      IF FFLD.UTXT > 0 THEN
         X$ = UTXT$(FFLD.UTXT)
         SELECT CASE FFLD.NAME
            CASE "LOID": X& = VAL(X$): CALL ChkID(X&)
                         IF NOT OKAY THEN X& = FF1HDR.MAXID
                         LORID = X&: HISrchID = LORID - 1
            CASE "HIID": X& = VAL(X$): CALL ChkID(X&)
                         IF NOT OKAY THEN X& = FF1HDR.MAXID
                         HIRID = X&
            CASE "SRNM": SURNAME$ = RTRIM$(X$)
            CASE "GVNM": GIVEN$ = RTRIM$(X$)
            CASE "SEX ": SEX$ = RTRIM$(X$)
            CASE "CHGS": CHGS = -(X$ = SNGLKEY$(1)) - 2 * (X$ = SNGLKEY$(2)) - 3 * (X$ = SNGLKEY$(4))
            CASE "BPL ": SBPLACE$ = RTRIM$(X$)
            CASE "DPL ": SDPLACE$ = RTRIM$(X$)
            CASE "LOBD": LOBMD = DateMD(X$, DFM): LOBY = DateY(X$, DFM)
            CASE "HIBD": HIBMD = DateMD(X$, DFM): HIBY = DateY(X$, DFM)
            CASE "LODD": LODMD = DateMD(X$, DFM): LODY = DateY(X$, DFM)
            CASE "HIDD": HIDMD = DateMD(X$, DFM): HIDY = DateY(X$, DFM)
            CASE "BDST": BDST$ = RTRIM$(X$)
            CASE "DDST": DDST$ = RTRIM$(X$)
            CASE "FST ": FIDST$ = RTRIM$(X$)
            CASE "MST ": MIDST$ = RTRIM$(X$)
         END SELECT
      END IF
      FF = FLDCHN(FF).FWD
   WEND
   BGNID = LORID
   BDCheck = ((LOBMD > 0) OR (LOBY > 0) OR (HIBMD < 9999) OR (HIBY < 9999))
   BDCheck = BDCheck OR ((LODMD > 0) OR (LODY > 0) OR (HIDMD < 9999) OR (HIDY < 9999))
   BDCheck = BDCheck OR ((SBPLACE$ <> "") OR (SDPLACE$ <> ""))
   IF NOT BDCheck THEN B = -1: C = -1: D = -1

403 ' Perform SEARCH
   OKAY = -1: COLOR FG(5), BG(5)
   CALL PutMSG("")
   WHILE (OKAY AND (LORID <= HIRID))
      CALL FF1GetRec(LORID)
      IF OKAY THEN
         A = ((INSTR(FF1REC.SURNM, SURNAME$) > 0) AND (INSTR(FF1REC.GIVEN, GIVEN$) > 0))
         IF BDCheck THEN
            CALL GetBDEvents(BDT$, BPLACE, DDT$, DPLACE, AGE$)
            B = ((LOBMD <= FF1REC.BMD) AND (LOBY <= FF1REC.BY) AND (FF1REC.BMD <= HIBMD) AND (FF1REC.BY <= HIBY))
            C = ((LODMD <= FF1REC.DMD) AND (LODY <= FF1REC.DY) AND (FF1REC.DMD <= HIDMD) AND (FF1REC.DY <= HIDY))
            D = (((SBPLACE$ = "") OR (INSTR(BPLACE, SBPLACE$) > 0)) AND ((SDPLACE$ = "") OR (INSTR(DPLACE, SDPLACE$) > 0)))
         END IF
         E = (SEX$ = "") OR (SEX$ = FF1REC.SEX)
         X = ASC(FF1REC.STATUS2)
         F = ((CHGS = 0) OR ((CHGS = 1) AND ((X AND 64) > 0)) OR ((CHGS = 2) AND ((X AND 64) = 0)) OR ((CHGS = 3) AND ((X AND 96) > 0)))
         G = (BDST$ = "") OR (BDST$ = FDSTAT$("BD"))
         H = (DDST$ = "") OR (DDST$ = FDSTAT$("DD"))
         I = (FIDST$ = "") OR (FIDST$ = FDSTAT$("FID"))
         J = (MIDST$ = "") OR (MIDST$ = FDSTAT$("MID"))
         IF (A AND B AND C AND D AND E AND F AND G AND H AND I AND J) THEN
            SMAX = SMAX + 1: SRCHID(SMAX) = LORID
         END IF
         LOCATE 25, 1, 0: PRINT STR$(LORID) + " " + STR$(SMAX);
         A$ = INKEY$: IF A$ <> "" THEN OKAY = 0
         LORID = LORID + 1
      END IF
   WEND
   IF SMAX = 0 THEN CALL PutMSG("UPDTM001"): GOTO 402
   SRCHID(0) = 1: SRCHTable.CL = 0
   SRCHTable.FF = 1: SRCHTable.FT = 0

SearchLIST:
   CALL WinSWITCH(6)
   IF WIN.WFMT <> FMT004 THEN
      CALL WinCLR
   END IF
   CALL WinFORMAT(FMT004)
   CALL OptDISPLAY(FMT907)
   IF LORID > HIRID THEN REDIM PRESERVE SRCHID(SMAX)
   IF SRCHTable.FT = 0 THEN SRCHTable.FT = 1: SRCHTable.FF = 0
   CALL TableOPEN(SRCHTable)

405 ' Display Results
   SYSVAR.NVAR1 = LORID - BGNID: SYSVAR.NVAR2 = HIRID - BGNID + 1
   SYSVAR.NVAR3 = SMAX
   CALL PutMSG("UPDTM002")
   CALL SearchSELECT(SRCHTable, SRCHLIST$(), SRCHCHN(), SRCHID(), SMAX)
   CALL PutMSG("")
   SELECT CASE A
      CASE 27: CALL WinCLR: REDIM SRCHID(0): SMAX = 0: GOTO SEARCHSetup
      CASE 13: BRID = VAL(MID$(SRCHLIST$(SRCHTable.CT), 1, 5))
               EXIT SUB
      CASE 59: CALL OptHILITE("F1", "", 6): SRCNT = 0
               CALL SWRKCreate(SRCHID(), SMAX, 1): CALL FamOPEN("VLDT")
               CALL OptHILITE("F1", "", 1): GOTO 405
      CASE 64: GOTO 410
      CASE 80: IF LORID < HIRID THEN GOTO 403
      CASE 81: IF LORID < HIRID THEN GOTO 403
      CASE 101: GOTO 403
   END SELECT
   CALL ErrBEEP(0): GOTO 405

410 ' Print SEARCH Results
   CALL OptHILITE("F6", "", 6)
   CALL FamCLOSE
   CALL RptOPEN("SRCH"): IF OKAY THEN CALL PrintOPEN
   IF NOT OKAY THEN GOTO 415
   IF LEN(PLINE$) < 133 THEN PLINE$ = SPACE$(113)
   CALL FamOPEN("VLDT")
411 Z = 0
   WHILE (Z < SMAX) AND OKAY
      Z = Z + 1
      CALL FF1GetRec(SRCHID(Z))
      IF OKAY THEN
         MID$(PLINE$, 1, 6) = SFILL$(FF1REC.RID, 6)
         MID$(PLINE$, 11, 6) = SFILL$(FF1REC.FID, 6)
         MID$(PLINE$, 17, 1) = FDSTAT$("FID")
         MID$(PLINE$, 18, 6) = SFILL$(FF1REC.MID, 6)
         MID$(PLINE$, 24, 1) = FDSTAT$("MID")
         X = VAL(DateDif$(FF1REC.BMD, FF1REC.BY, FF1REC.DMD, FF1REC.DY))
         MID$(PLINE$, 26, 4) = SFILL$(X, 4)
         MID$(PLINE$, 33, 1) = FF1REC.SEX
         MID$(PLINE$, 37, 10) = XSCRNDate$(FF1REC.BMD, FF1REC.BY)
         MID$(PLINE$, 47, 1) = FDSTAT$("BD")
         MID$(PLINE$, 49, 10) = XSCRNDate$(FF1REC.DMD, FF1REC.DY)
         MID$(PLINE$, 59, 1) = FDSTAT$("DD")
         MID$(PLINE$, 61, 20) = FF1REC.SURNM
         MID$(PLINE$, 83, 30) = FF1REC.GIVEN
         IF RPTOPTION.BLDREF THEN PAGE.XID = FF1REC.RID: PAGE.XTYPE = 1: CALL RREFAdd
         CALL PrintLINE
      END IF
      IF A = 27 THEN Z = SMAX + 1
   WEND
   CALL RptCLOSE: IF PRT.AGAIN THEN GOTO 411
415 CALL FamOPEN("VLDT")
   IF (PRT.OPT = 1) OR (PRT.OPT = 3) THEN
      WIN.WFMT = 0
      CALL WinSWITCH(2): CALL WinFORMAT(FMT003)
      CALL OptDISPLAY(FMT901): CALL OptHILITE("F1", "", 6)
      GOTO SearchLIST
   END IF
   GOTO 405

END SUB

SUB SearchSELECT (TP AS TablePARMS, TDATA$(), TCHN() AS ChainPTRS, SID(), SMAX)
    CURFMT = TP.FMTNUM: FMT = FMTTAB(CURFMT)
    IF TP.CL = 0 THEN TP.CL = TP.TL: TP.CT = TP.FS
    SL = TP.BL - TP.TL + 1
BLDSPAGE: ' Build PAGE of Results
   X = SID(0): Y = 0: X$ = SPACE$(TP.XL)
   CALL ChnINIT(TCHN(), 1, SL)
   WHILE (X <= SMAX) AND (Y < SL)
      Y = Y + 1: GOSUB GETSDATA: X = X + 1
   WEND
   TCHN(Y).FWD = 0
   TP.LT = Y: TP.SIZE = TP.LT
SHOWPAGE:
   CALL TableSHOW(TP, TDATA$(), TCHN())
   NLIN = 0: GOTO SHOWSKEY
GETSDATA:
   CALL FF1GetRec(SID(X))
   MID$(X$, 1, 5) = SFILL$(SID(X), 5)
   MID$(X$, 6, 5) = SFILL$(FF1REC.FID, 5)
   MID$(X$, 11, 1) = FDSTAT$("FID")
   MID$(X$, 12, 5) = SFILL$(FF1REC.MID, 5)
   MID$(X$, 17, 1) = FDSTAT$("MID")
   MID$(X$, 18, 1) = FF1REC.SEX
   MID$(X$, 19, 10) = XSCRNDate$(FF1REC.BMD, FF1REC.BY)
   MID$(X$, 29, 1) = FDSTAT$("BD")
   MID$(X$, 30, 10) = XSCRNDate$(FF1REC.DMD, FF1REC.DY)
   MID$(X$, 40, 1) = FDSTAT$("DD")
   MID$(X$, 41, 15) = FF1REC.SURNM
   MID$(X$, 56, 18) = FF1REC.GIVEN
   TDATA$(Y) = X$
   RETURN
SHOWSKEY:
    IF NLIN > 0 THEN
       CALL TableKEY("UNMARK", TP, TDATA$())
       TP.CL = NLIN: TP.CT = NTAB
    END IF: NLIN = 0
    IF TP.KEYFLD > 0 THEN
       CALL TableKEY("HILITE", TP, TDATA$())
       ELSE
       CALL TableKEY("MARK", TP, TDATA$())
    END IF
USERSCTL:
    CALL GetKEY: IF LEN(A$) > 1 THEN GOTO SFUNKEY
    IF (A = 27) OR (A = 13) THEN EXIT SUB
    CALL ErrBEEP(0): GOTO USERSCTL
SFUNKEY:
    IF A = 59 THEN EXIT SUB
    IF A = 72 THEN                            ' UP
       NLIN = TP.CL - 1: NTAB = TCHN(TP.CT).BWD
       IF NLIN >= TP.TL THEN GOTO SHOWSKEY
       IF SID(0) > 1 THEN
          SID(0) = SID(0) - 1
          IF TP.LT < SL THEN GOTO BLDSPAGE
          Y = SL: WHILE Y > 1: TDATA$(Y) = TDATA$(Y - 1): Y = Y - 1: WEND
          X = SID(0): GOSUB GETSDATA
          GOTO SHOWPAGE
       END IF
       CALL ErrBEEP(0): GOTO USERSCTL
    END IF
    IF A = 80 THEN                            ' DOWN
       NLIN = TP.CL + 1: NTAB = TCHN(TP.CT).FWD
       IF NLIN <= TP.LL THEN GOTO SHOWSKEY
       X = SID(0) + (TP.CL - TP.TL)
       IF X < SMAX THEN
          SID(0) = SID(0) + 1
          Y = 1: WHILE Y < SL: TDATA$(Y) = TDATA$(Y + 1): Y = Y + 1: WEND
          X = X + 1: GOSUB GETSDATA
          GOTO SHOWPAGE
       END IF
       EXIT SUB
    END IF
    IF A = 73 THEN                             ' PGUP
       X = SID(0) + (TP.CL - TP.TL) - (TP.BL - TP.TL) * (1 - (TP.CL = TP.BL))
       IF X < 1 THEN
          IF SID(0) > 1 THEN
             TP.CL = TP.TL + SID(0) - 1: SID(0) = 1: GOTO BLDSPAGE
             ELSE
             CALL ErrBEEP(0): GOTO USERSCTL
          END IF
       END IF
       SID(0) = X: TP.CL = TP.BL
       GOTO BLDSPAGE
    END IF
    IF A = 81 THEN                            ' PGDN
       X = SID(0) + TP.CL - TP.TL - (TP.BL - TP.TL) * (TP.CL = TP.TL)
       TP.CL = TP.TL
       IF X < SMAX THEN SID(0) = X: GOTO BLDSPAGE
       EXIT SUB
    END IF
    EXIT SUB
END SUB

REM $STATIC
SUB ShowNameDATA
SHARED FMT006, RGL, BRID, CFID, CMID
SHARED OBPLACE AS STRING * 41, ODPLACE AS STRING * 41
   ' Move Name Record DATA to Screen
   CURFMT = FMT006: FMT = FMTTAB(CURFMT)
   FF = FMT.BGNFLD
   WHILE FF > 0
      FFLD = FLDTAB(FF)
      IF FFLD.UTXT > 0 THEN
         X$ = UTXT$(FFLD.UTXT)
         SELECT CASE FFLD.NAME
            CASE "RID ": LSET X$ = STR$(BRID)
            CASE "RGL ": RSET X$ = LTRIM$(STR$(RGL))
            CASE "SRNM": LSET X$ = FF1REC.SURNM
            CASE "SNMU": LSET X$ = FDSTAT$("SNU")
            CASE "GVNM": LSET X$ = FF1REC.GIVEN
            CASE "BDT ": LSET X$ = XSCRNDate$(FF1REC.BMD, FF1REC.BY)
            CASE "BDST": LSET X$ = FDSTAT$("BD")
            CASE "BTM ": LSET X$ = XSCRNTime$(FF1REC.BTM)
            CASE "BPL ": LSET X$ = OBPLACE: FLDTAB(FF).DO = 0
            CASE "DDT ": LSET X$ = XSCRNDate$(FF1REC.DMD, FF1REC.DY)
            CASE "DDST": LSET X$ = FDSTAT$("DD")
            CASE "DTM ": LSET X$ = XSCRNTime$(FF1REC.DTM)
            CASE "DPL ": LSET X$ = ODPLACE: FLDTAB(FF).DO = 0
            CASE "SEX ": LSET X$ = FF1REC.SEX
            CASE "AGE ": LSET X$ = DateDif$(FF1REC.BMD, FF1REC.BY, FF1REC.DMD, FF1REC.DY)
            CASE "FID ": CFID = FF1REC.FID: RSET X$ = LTRIM$(STR$(CFID))
            CASE "FST ": LSET X$ = FDSTAT$("FID")
            CASE "MID ": CMID = FF1REC.MID: RSET X$ = LTRIM$(STR$(CMID))
            CASE "MST ": LSET X$ = FDSTAT$("MID")
         END SELECT
         LSET UTXT$(FFLD.UTXT) = X$
      END IF
      FF = FLDCHN(FF).FWD
   WEND
   CALL FmtUHOLD

END SUB

SUB UpdtPLACE (XLOC, SRTYPE, SRNO)
SHARED OBPLACE AS STRING * 41, ODPLACE AS STRING * 41
SHARED OMPLACE AS STRING * 41, ODVPLACE AS STRING * 41

670 ' Update PLACE Information
    SELECT CASE SRTYPE
       CASE 1: X1PL$ = RTRIM$(BPLACE): Y1PL$ = RTRIM$(DPLACE)
               X2PL$ = RTRIM$(OBPLACE): Y2PL$ = RTRIM$(ODPLACE)
       CASE 3: X1PL$ = RTRIM$(MPLACE): Y1PL$ = RTRIM$(DVPLACE)
               X2PL$ = RTRIM$(OMPLACE): Y2PL$ = RTRIM$(ODVPLACE)
    END SELECT
    IF (X1PL$ = X2PL$) AND (Y1PL$ = Y2PL$) THEN EXIT SUB
    IF XLOC <> 0 THEN
       CALL FF3GetRec(XLOC)
       IF FF3PFX.RTYPE = CHR$(12) THEN
          IF FF3PLACE2.PNEXT <> 0 THEN CALL FF3DelRec(FF3PLACE2.PNEXT)
       END IF
       CALL FF3DelRec(XLOC)
    END IF
    XLOC = 0: FF3PFX.RNO = 0
    PL1 = LEN(X1PL$): PL2 = LEN(Y1PL$)
    IF (PL1 > 0) OR (PL2 > 0) THEN
       IF ((PL1 > 22) OR (PL2 > 22)) AND ENV.LPNSupport THEN
          FF3PFX.RTYPE = CHR$(12)
          IF PL2 > 0 THEN
             CALL FF3GetRec(0)
             IF OKAY THEN
                FF3PLACE2.SRTYPE = CHR$(SRTYPE): FF3PLACE2.SRNO = SRNO
                FF3PLACE2.PTYPE = CHR$(SRTYPE + 1)
                LSET FF3PLACE2.PLACE = Y1PL$
                CALL FF3PutRec
                IF OKAY THEN XLOC = FF3PFX.RNO: FF3PFX.RNO = 0
             END IF
          END IF
          IF PL1 > 0 THEN
             CALL FF3GetRec(0)
             IF OKAY THEN
                FF3PLACE2.SRTYPE = CHR$(SRTYPE): FF3PLACE2.SRNO = SRNO
                FF3PLACE2.PTYPE = CHR$(SRTYPE)
                FF3PLACE2.PNEXT = XLOC
                LSET FF3PLACE2.PLACE = X1PL$
                CALL FF3PutRec
             END IF
          END IF
          ELSE
          FF3PFX.RTYPE = CHR$(4)
          CALL FF3GetRec(0)
          IF OKAY THEN
             FF3PLACE.SRTYPE = CHR$(SRTYPE): FF3PLACE.SRNO = SRNO
             LSET FF3PLACE.PLACE1 = X1PL$: LSET FF3PLACE.PLACE2 = Y1PL$
             LSET X1PL$ = FF3PLACE.PLACE1: LSET Y1PL$ = FF3PLACE.PLACE2
             CALL FF3PutRec
          END IF
       END IF
       IF OKAY THEN
          XLOC = FF3PFX.RNO
          SELECT CASE SRTYPE
             CASE 1: LSET OBPLACE = X1PL$: LSET ODPLACE = Y1PL$
             CASE 3: LSET OMPLACE = X1PL$: LSET ODVPLACE = Y1PL$
          END SELECT
       END IF
    END IF

END SUB

REM $DYNAMIC
SUB ViewADDRESS (ADRBGN, SRTYPE, SRNO)
3000 ' Process ADDRESS Information
    DFM = VAL(CONFIG.DFM)
    AFMT = FF2HDR.RevLVL AND 255
    ZERODATE$ = XSCRNDate$(0, 0)
    CALL FmtFIND("UPDTS909"): FMT909 = CURFMT
    CALL FmtFIND("UPDTS912"): FMT912 = CURFMT
    IF WINTAB(6).WFMT <> 0 THEN
       CALL WinSWITCH(6)
       CALL WinCLR
    END IF
    CALL WinSWITCH(5)
    CALL FmtFIND("UPDTS012"): FMT012 = CURFMT
    CALL FmtFindFLD("CTRY")
    IF FOUND THEN
       SELECT CASE AFMT
          CASE 0: LSET UTXT$(FFLD.UTXT) = ""
          CASE ELSE: LSET UTXT$(FFLD.UTXT) = DTXT$(WIN.CURFLD)
       END SELECT
    END IF
    CALL WinFORMAT(FMT012)
    SELECT CASE AFMT
       CASE 0: CALL FmtFIND("UPDTS013")
       CASE ELSE: CALL FmtFIND("UPDTS014")
    END SELECT
    ADRFMT = CURFMT
    CALL FmtFindFLD("DT1 "): LSET DTXT$(WIN.CURFLD) = ZERODATE$
    CALL FmtFindFLD("DT2 "): LSET DTXT$(WIN.CURFLD) = ZERODATE$
    ADRIX = 0: ADRHIX = 10
    REDIM ADRCHN(ADRHIX)
    ADRRNO = ADRBGN
    IF ADRRNO = 0 THEN
       CALL OptDISPLAY(FMT909)
       CALL OptHILITE("F2", "", 6)
       GOTO 3020
    END IF
    GOTO 3005

3002 ' Move Address INFO to Screen
    CURFMT = ADRFMT: FMT = FMTTAB(ADRFMT)
    CALL GetUFLD("FIRST"): FF = WIN.CURFLD
    IF AFMT = 0 THEN
       WHILE WIN.CURFLD > 0
          X$ = UTXT$(FFLD.UTXT)
          SELECT CASE FFLD.NAME
             CASE "DT1 ": LSET X$ = XSCRNDate$(FF2PFX.D1MD, FF2PFX.D1Y)
             CASE "DT2 ": LSET X$ = XSCRNDate$(FF2PFX.D2MD, FF2PFX.D2Y)
             CASE "FONE": MID$(X$, 2, 3) = ZFILL$(FF2SRec.FON1, 3)
                          MID$(X$, 6, 3) = ZFILL$(FF2SRec.FON2, 3)
                          MID$(X$, 10, 4) = ZFILL$(FF2SRec.FON3, 4)
             CASE "ADL1": LSET X$ = FF2SRec.LINE1
             CASE "ADL2": LSET X$ = FF2SRec.LINE2
             CASE "CITY": LSET X$ = FF2SRec.CITY
             CASE "STAT": LSET X$ = FF2SRec.STATE
             CASE "ZIP ": LSET X$ = FF2SRec.ZIP
          END SELECT
          LSET UTXT$(FFLD.UTXT) = X$
          CALL GetUFLD("NEXT")
          IF FF = WIN.CURFLD THEN WIN.CURFLD = 0
       WEND
       ELSE
       WHILE WIN.CURFLD > 0
          X$ = UTXT$(FFLD.UTXT)
          SELECT CASE FFLD.NAME
             CASE "DT1 ": LSET X$ = XSCRNDate$(FF2PFX.D1MD, FF2PFX.D1Y)
             CASE "DT2 ": LSET X$ = XSCRNDate$(FF2PFX.D2MD, FF2PFX.D2Y)
             CASE "FONE": LSET X$ = FF2LRec.PHONE
             CASE "ADL1": LSET X$ = FF2LRec.LINE1
             CASE "ADL2": LSET X$ = FF2LRec.LINE2
             CASE "CITY": LSET X$ = FF2LRec.CITY
             CASE "STAT": LSET X$ = FF2LRec.STATE
             CASE "ZIP ": LSET X$ = FF2LRec.ZIP
             CASE "CTRY": LSET X$ = FF2LRec.COUNTRY
          END SELECT
          LSET UTXT$(FFLD.UTXT) = X$
          CALL GetUFLD("NEXT")
          IF FF = WIN.CURFLD THEN WIN.CURFLD = 0
       WEND
    END IF
    CALL WinFORMAT(ADRFMT)
    AD1MD = FF2PFX.D1MD: AD1Y = FF2PFX.D1Y
    CALL FmtUHOLD
    RETURN

3004 ' Move ADDRESS Info from Screen
    CURFMT = ADRFMT: FMT = FMTTAB(ADRFMT)
    CALL GetUFLD("FIRST"): FF = WIN.CURFLD
    IF AFMT = 0 THEN
       WHILE WIN.CURFLD > 0
          X$ = UTXT$(FFLD.UTXT)
          SELECT CASE FFLD.NAME
             CASE "DT1 ": FF2PFX.D1MD = DateMD(X$, DFM)
                          FF2PFX.D1Y = DateY(X$, DFM)
             CASE "DT2 ": FF2PFX.D2MD = DateMD(X$, DFM)
                          FF2PFX.D2Y = DateY(X$, DFM)
             CASE "FONE": FF2SRec.FON1 = VAL(MID$(X$, 2, 3))
                          FF2SRec.FON2 = VAL(MID$(X$, 6, 3))
                          FF2SRec.FON3 = VAL(MID$(X$, 10, 4))
             CASE "ADL1": LSET FF2SRec.LINE1 = X$
             CASE "ADL2": LSET FF2SRec.LINE2 = X$
             CASE "CITY": LSET FF2SRec.CITY = X$
             CASE "STAT": LSET FF2SRec.STATE = X$
             CASE "ZIP ": LSET FF2SRec.ZIP = X$
          END SELECT
          CALL GetUFLD("NEXT")
          IF FF = WIN.CURFLD THEN WIN.CURFLD = 0
       WEND
       ELSE
       WHILE WIN.CURFLD > 0
          X$ = UTXT$(FFLD.UTXT)
          SELECT CASE FFLD.NAME
             CASE "DT1 ": FF2PFX.D1MD = DateMD(X$, DFM)
                          FF2PFX.D1Y = DateY(X$, DFM)
             CASE "DT2 ": FF2PFX.D2MD = DateMD(X$, DFM)
                          FF2PFX.D2Y = DateY(X$, DFM)
             CASE "FONE": LSET FF2LRec.PHONE = X$
             CASE "ADL1": LSET FF2LRec.LINE1 = X$
             CASE "ADL2": LSET FF2LRec.LINE2 = X$
             CASE "CITY": LSET FF2LRec.CITY = X$
             CASE "STAT": LSET FF2LRec.STATE = X$
             CASE "ZIP ": LSET FF2LRec.ZIP = X$
             CASE "CTRY": LSET FF2LRec.COUNTRY = X$
          END SELECT
          LSET UTXT$(FFLD.UTXT) = X$
          CALL GetUFLD("NEXT")
          IF FF = WIN.CURFLD THEN WIN.CURFLD = 0
       WEND
    END IF
    RETURN

3005 ' Process Existing ADDRESS Record
    CALL FF2GetRec(ADRRNO)
    ADRCHN(ADRIX) = ADRRNO
    ADRCHN(ADRIX + 1) = FF2PFX.NXT
    GOSUB 3002: NewADR = 0
3006 ' Get Next Request
    CALL OptDISPLAY(FMT912)
    IF FF2PFX.COM <> 0 THEN CALL OptHILITE("F2", "", 7)
    IF (ADRIX > 0) THEN
       IF (ADRCHN(ADRIX + 1) <> 0) THEN
          CALL PutMSG("UPDTM004")
          ELSE
          CALL PutMSG("UPDTM005")
       END IF
       ELSE
       IF ADRCHN(ADRIX + 1) <> 0 THEN
          CALL PutMSG("UPDTM006")
          ELSE
          CALL PutMSG("")
       END IF
    END IF
    CALL GetKEY: FUN$ = ""
    SELECT CASE A
       CASE 27: GOTO 3029
       CASE 59: FUN$ = "F1"
       CASE 60: FUN$ = "F2"
       CASE 73: FUN$ = "PGUP"
                IF ADRCHN(ADRIX + 1) <> 0 THEN
                   ADRIX = ADRIX + 1
                   ADRRNO = ADRCHN(ADRIX)
                   IF NOT (ADRIX < ADRHIX) THEN
                      ADRHIX = ADRHIX + 5
                      REDIM PRESERVE ADRCHN(ADRHIX)
                   END IF
                   GOTO 3005
                   ELSE
                   FUN$ = ""
                END IF
       CASE 81: FUN$ = "PGDN"
                IF ADRIX > 0 THEN
                   ADRIX = ADRIX - 1
                   ADRRNO = ADRCHN(ADRIX)
                   GOTO 3005
                   ELSE
                   FUN$ = ""
                END IF
    END SELECT
    IF FUN$ = "" THEN CALL ErrBEEP(0): GOTO 3006
    CALL OptHILITE(FUN$, "", 6)
    SELECT CASE FUN$
       CASE "F1": GOTO 3010
       CASE "F2": COMBGN = FF2PFX.COM
                  CALL ViewCOMMENTS(COMBGN, 2, ADRRNO)
                  CALL WinSWITCH(5)
                  IF COMBGN <> FF2PFX.COM THEN
                     FF2PFX.COM = COMBGN
                     CALL FF2PutRec
                  END IF
    END SELECT
    GOTO 3006

3010 ' Update Address Record
    CALL PutMSG("")
    CALL OptDISPLAY(FMT909)
    CALL GetKEY
    SELECT CASE A
       CASE 27: GOTO 3006
       CASE 59: FUN$ = "F1"
       CASE 60: FUN$ = "F2"
       CASE 61: FUN$ = "F3"
       CASE ELSE: CALL ErrBEEP(0): GOTO 3010
    END SELECT
    CALL OptHILITE(FUN$, "", 6)
    SELECT CASE FUN$
       CASE "F1": GOTO 3015   ' Change
       CASE "F2": GOTO 3020   ' Add
       CASE "F3": GOTO 3025   ' Delete
    END SELECT

3015 ' Modify Current Address Record
   CALL PutMSG("XXXXM012")
   CURFMT = ADRFMT: FMT = FMTTAB(CURFMT)
   WIN.CLIN = 0: CALL WinUPDATE
   IF (A = 27) OR (A = 60) THEN CALL FMTHREST: CALL WinFORMAT(0): GOTO 3010
   IF A <> 59 THEN CALL ErrBEEP(0): GOTO 3015

3016 ' Write Address Record to FILE
   CALL PutMSG("")
   GOSUB 3004
   CALL FF2PutRec
   IF (ADRBGN = 0) OR NewADR OR (FF2PFX.D1MD <> AD1MD) OR (FF2PFX.D1Y <> AD1Y) THEN
      REDIM ADRCHN(ADRHIX): ADRIX = 0
      ELSE
      GOTO 3006
   END IF
   IF ADRBGN = 0 THEN
      ADRBGN = ADRRNO:  ADRCHN(0) = ADRBGN
      ELSE
      AD1MD = FF2PFX.D1MD: AD1Y = FF2PFX.D1Y
      GOSUB 3017
      IF ADRBGN = 0 THEN
         ADRBGN = ADRRNO: ADRCHN(0) = ADRBGN
         ELSE
         X = ADRBGN: Y = 0
         WHILE X <> 0
            CALL FF2GetRec(X)
            IF OKAY THEN
               IF (FF2PFX.D1Y > AD1Y) OR ((FF2PFX.D1Y = AD1Y) AND (FF2PFX.D1MD > AD1MD)) THEN
                  ADRCHN(ADRIX) = X
                  ADRIX = ADRIX + 1
                  IF NOT (ADRIX < ADRHIX) THEN
                     ADRHIX = ADRHIX + 5
                     REDIM PRESERVE ADRCHN(ADRHIX)
                  END IF
                  Y = X: X = FF2PFX.NXT
                  ELSE
                  ADRCHN(ADRIX + 1) = X: X = 0
               END IF
               ELSE
               X = 0
            END IF
         WEND
         IF Y = 0 THEN
            ADRBGN = ADRRNO
            ELSE
            CALL FF2GetRec(Y)
            IF OKAY THEN
               FF2PFX.NXT = ADRRNO
               CALL FF2PutRec
            END IF
         END IF
         CALL FF2GetRec(ADRRNO)
         IF OKAY THEN
            IF FF2PFX.NXT <> ADRCHN(ADRIX + 1) THEN
               FF2PFX.NXT = ADRCHN(ADRIX + 1)
               CALL FF2PutRec
            END IF
         END IF
      END IF
   END IF
   GOTO 3005

3017 ' Remove ADRRNO From Address Chain
   CALL PutMSG("")
   Y = FF2PFX.NXT
   IF ADRBGN = ADRRNO THEN
      ADRBGN = Y
      ELSE
      X = ADRBGN
      WHILE (X <> 0)
         CALL FF2GetRec(X)
         IF OKAY THEN
            IF FF2PFX.NXT = ADRRNO THEN
               FF2PFX.NXT = Y
               CALL FF2PutRec
               X = 0
               ELSE
               X = FF2PFX.NXT
            END IF
            ELSE
            X = 0
         END IF
      WEND
   END IF
   CALL FF2GetRec(ADRRNO)
   IF OKAY THEN
      FF2PFX.NXT = 0
      CALL FF2PutRec
   END IF
   RETURN

3020 ' Add New Address Record
   CALL FF2GetRec(0): IF NOT OKAY THEN GOTO 3029
   CALL PutMSG("XXXXM012")
   CURFMT = ADRFMT: FMT = FMTTAB(CURFMT)
   CALL FMTDREST: CALL WinFORMAT(0): WIN.CLIN = 0
3021 ' Enter NEW Address Info
   CALL WinUPDATE
   IF (A = 27) OR (A = 60) THEN
      CALL FF2DelRec(FF2PFX.RNO)
      IF ADRRNO = 0 THEN
         GOTO 3029
         ELSE
         CALL FMTHREST: CALL WinFORMAT(ADRFMT)
         CALL FF2GetRec(ADRRNO)
         GOTO 3010
      END IF
   END IF
   IF A <> 59 THEN CALL ErrBEEP(0): GOTO 3021
   ADRRNO = FF2PFX.RNO
   FF2PFX.SRTYPE = CHR$(SRTYPE): FF2PFX.SRNO = SRNO
   NewADR = -1: GOTO 3016

3025 ' Delete Current ADDRESS Record
   CALL PutMSG("UPDTM007")
   IF A$ <> SNGLKEY$(1) THEN GOTO 3010
   ACOM = FF2PFX.COM
   IF ACOM <> 0 THEN
      CALL PutMSG("UPDTM009")
      IF A$ <> SNGLKEY$(1) THEN GOTO 3010
   END IF
   WHILE ACOM <> 0
       CALL FF3GetRec(ACOM)
       IF NOT OKAY THEN
          ACOM = 0
          ELSE
          X = FF3COMMENT.FPTR
          CALL FF3DelRec(ACOM)
          IF NOT OKAY THEN
             ACOM = 0
             ELSE
             ACOM = X
          END IF
       END IF
   WEND
   GOSUB 3017
   CALL FF2DelRec(ADRRNO)
   ADRCHN(ADRIX) = ADRCHN(ADRIX + 1)
   ADRRNO = ADRCHN(ADRIX)
   IF (ADRRNO = 0) AND (ADRIX > 0) THEN ADRIX = ADRIX - 1: ADRRNO = ADRCHN(ADRIX)
   IF ADRRNO <> 0 GOTO 3005

3029 ' Return to Calling Routine
   CALL PutMSG("")
   CALL WinCLR
   REDIM ADRCHN(0): ADRHIX = 0
END SUB

REM $STATIC
SUB ViewCHILDREN (PID1, PID2)
SHARED FUN$, RGL, CGL, CRID, CMAX, CTSIZE, CHLDTable AS TablePARMS
SHARED CHLDLIST$(), CHLDCHN() AS ChainPTRS
   CALL FmtFIND("UPDTS008"): FMT008 = CURFMT
   CALL FmtFIND("UPDTS916"): FMT916 = CURFMT
   CALL WinSWITCH(6)
   IF WIN.WFMT <> FMT008 THEN
      CALL WinCLR
   END IF
   CALL WinFORMAT(FMT008)
   IF CMAX > 0 THEN RGL = CGL: GOTO 801
   RGL = RGL + 1: CGL = RGL
   CALL FmtFindFLD("RGL"): IF FOUND THEN LSET UTXT$(FFLD.UTXT) = STR$(CGL)
   CALL WinDisplayDATA(0)
   CTSIZE = 10: REDIM CHLDLIST$(CTSIZE), CHLDCHN(CTSIZE) AS ChainPTRS
   CALL ChnINIT(CHLDCHN(), 1, 10): CF = 1
   CHLDTable.FT = 0: CHLDTable.LT = 0: CHLDTable.FF = 0: CHLDTable.SIZE = 0
   CALL TableOPEN(CHLDTable)
   IF PID1 > 0 THEN PRID = PID1: OPID = PID2 ELSE PRID = PID2: OPID = PID1
   CALL FF1GetRec(PRID)
   C = FF1REC.OLDCH: X$ = SPACE$(CHLDTable.XL)
   WHILE C > 0
      CALL FF1GetRec(C)
      IF NOT OKAY THEN
         C = 0
         ELSE
         IF (OPID = 0) OR (((FF1REC.FID = PID1) AND (FF1REC.MID = PID2)) OR ((FF1REC.MID = PID1) AND (FF1REC.FID = PID2))) THEN
            MID$(X$, 1, 5) = SFILL$(C, 5)
            MID$(X$, 6, 5) = SFILL$(FF1REC.FID, 5)
            MID$(X$, 11, 1) = FDSTAT$("FID")
            MID$(X$, 12, 5) = SFILL$(FF1REC.MID, 5)
            MID$(X$, 17, 1) = FDSTAT$("MID")
            MID$(X$, 18, 1) = FF1REC.SEX
            MID$(X$, 19, 10) = XSCRNDate$(FF1REC.BMD, FF1REC.BY)
            MID$(X$, 29, 1) = FDSTAT$("BD")
            MID$(X$, 30, 10) = XSCRNDate$(FF1REC.DMD, FF1REC.DY)
            MID$(X$, 40, 1) = FDSTAT$("DD")
            IF (FF1REC.DY = 9999) OR (FF1REC.BY = 0) THEN
               MID$(X$, 41, 4) = "  ??"
               ELSE
               X = VAL(DateDif$(FF1REC.BMD, FF1REC.BY, FF1REC.DMD, FF1REC.DY))
               MID$(X$, 41, 4) = SFILL$(X, 4)
            END IF
            MID$(X$, 46, 29) = FF1REC.GIVEN
            IF CF = 0 THEN
               CF = CTSIZE + 1: CTSIZE = CTSIZE + 5
               REDIM PRESERVE CHLDLIST$(CTSIZE), CHLDCHN(CTSIZE) AS ChainPTRS
               CALL ChnINIT(CHLDCHN(), CF, CTSIZE)
               CHLDCHN(CF).BWD = CF - 1: CHLDCHN(CF - 1).FWD = CF
            END IF
            CHLDLIST$(CF) = X$: CHLDTable.LT = CF
            CF = CHLDCHN(CF).FWD
         END IF
      END IF
      IF FF1REC.FID = PRID THEN C = FF1REC.FCH ELSE C = FF1REC.MCH
   WEND
   CMAX = CHLDTable.LT: CHLDCHN(CMAX).FWD = 0
   REDIM PRESERVE CHLDLIST$(CMAX), CHLDCHN(CMAX) AS ChainPTRS
   IF CMAX = 0 THEN
      CALL PutMSG("UPDTM018")
      CALL WinCLR: FUN$ = ""
      EXIT SUB
   END IF
   CHLDTable.FT = 1: CHLDTable.SIZE = CMAX: CHLDTable.LT = CMAX
   CHLDTable.CL = 0
   CALL TableOPEN(CHLDTable)
801 CALL TableSHOW(CHLDTable, CHLDLIST$(), CHLDCHN())
802 CALL OptDISPLAY(FMT916)
   IF CHLDTable.SIZE > 9 THEN
      CALL PutMSG("UPDTM017")
   END IF
   CALL TableSELECT(CHLDTable, CHLDLIST$(), CHLDCHN())
   CALL PutMSG("")
   SELECT CASE A
       CASE 13: FUN$ = "SEL"
       CASE 62: FUN$ = "F4"
       CASE 27: FUN$ = "ESC": A = 0
       CASE ELSE: CALL ErrBEEP(0): GOTO 802
   END SELECT
   CALL OptHILITE("FUN$", "", 6)
   SELECT CASE FUN$
       CASE "SEL": CRID = VAL(LEFT$(CHLDLIST$(CHLDTable.CT), 5))
                   EXIT SUB
       CASE "ESC": CALL WinCLR: CMAX = 0: CRID = 0: RGL = RGL - 1
                   REDIM CHLDLIST$(0), CHLDCHN(0) AS ChainPTRS
                   EXIT SUB
       CASE "F4": GOTO 805
       CASE ELSE: GOTO 802
   END SELECT
805 ' Move Child in List
   XCB = CHLDCHN(CHLDTable.CT).BWD
   CALL TableUPDT("MOV", CHLDTable, CHLDLIST$(), CHLDCHN())
   IF XCB = CHLDCHN(CHLDTable.CT).BWD THEN GOTO 802
   XCB = CHLDCHN(CHLDTable.CT).BWD
   IF XCB > 0 THEN XCB = VAL(MID$(CHLDLIST$(XCB), 1, 5))
   XRID = VAL(MID$(CHLDLIST$(CHLDTable.CT), 1, 5))
   CALL FF1GetRec(XRID)
   NameREC = FF1REC
   CFID = NameREC.FID: CMID = NameREC.MID
   IF (CFID <> PRID) AND (CFID <> OPID) THEN CFID = 0
   IF (CMID <> PRID) AND (CMID <> OPID) THEN CMID = 0
   IF (CFID > 0) THEN CALL UpdtRelPTRS("REM", CFID)
   IF (CMID > 0) THEN CALL UpdtRelPTRS("REM", CMID)
   IF (CFID > 0) THEN UXID = CFID: GOSUB 810
   IF (CMID > 0) THEN UXID = CMID: GOSUB 810
   FF1REC = NameREC
   CALL FF1PutREC
   GOTO 802

810 ' Place NameREC after XCB in child chain of UXID
   CALL FF1GetRec(UXID)
   IF XCB = 0 THEN
      IF UXID = CFID THEN
         NameREC.FCH = FF1REC.OLDCH
         ELSE
         NameREC.MCH = FF1REC.OLDCH
      END IF
      FF1REC.OLDCH = NameREC.RID: OKAY = -1
      ELSE
      X = FF1REC.OLDCH: FOUND = 0
      WHILE (X > 0) AND (NOT FOUND)
         CALL FF1GetRec(X)
         IF NOT OKAY THEN
            X = 0
            ELSE
            FOUND = (X = XCB)
            IF NOT FOUND THEN
               IF FF1REC.FID = UXID THEN
                  X = FF1REC.FCH
                  ELSE
                  X = FF1REC.MCH
               END IF
            END IF
         END IF
      WEND
      OKAY = FOUND
      IF FOUND THEN
         IF UXID = CFID THEN
            IF UXID = FF1REC.FID THEN
               NameREC.FCH = FF1REC.FCH
               FF1REC.FCH = NameREC.RID
               ELSE
               NameREC.FCH = FF1REC.MCH
               FF1REC.MCH = NameREC.RID
            END IF
            ELSE
            IF UXID = FF1REC.FID THEN
               NameREC.MCH = FF1REC.FCH
               FF1REC.FCH = NameREC.RID
               ELSE
               NameREC.MCH = FF1REC.MCH
               FF1REC.MCH = NameREC.RID
            END IF
         END IF
      END IF
   END IF
   IF OKAY THEN CALL FF1PutREC
   RETURN

CVEXIT:

END SUB

SUB ViewCOMMENTS (COMBGN, SRTYPE, SRNO)
   CALL FmtFIND("UPDTS015"): FMT015 = CURFMT
   IF SRTYPE = 6 THEN
      CALL FmtFIND("UPDTS017")
      ELSE
      CALL FmtFIND("UPDTS016")
   END IF
   COMFMT = CURFMT
   CALL FmtFIND("UPDTS915"): FMT915 = CURFMT
   CALL PutMSG("")
   IF SRTYPE = 6 THEN
      CALL WinSWITCH(2)
      IF WIN.WFMT <> 0 THEN CALL WinCLR
      ELSE
      IF WINTAB(6).WFMT <> 0 THEN
         CALL WinSWITCH(6)
         CALL WinCLR
      END IF
      CALL WinSWITCH(4)
   END IF
   TL = WIN.ULR + 2: BL = WIN.LRR - 2: SLC = BL - TL + 1: WIN.CLIN = 0
   CTMAX = 25: REDIM COMRNO(CTMAX), PTRCHG(CTMAX), COMPTR(CTMAX) AS ChainPTRS
   COMRNO(0) = 1: COMPTR(0).FWD = COMBGN: CTMAX = 0: R = 0
   IF COMBGN <> 0 THEN
      CALL FF3GetRec(COMBGN): IF NOT OKAY THEN GOTO COMURTRN
      CTOT = FF3COMMENT.CCNT
      ELSE
      CTOT = 0
   END IF
BLDCPAGE:
   Y = COMRNO(0): X = COMPTR(Y - 1).FWD: Z = 0
   CURFMT = FMT015: FMT = FMTTAB(CURFMT)
   CALL FmtFindFLD("TLNO"): IF FOUND THEN RSET UTXT$(FFLD.UTXT) = LTRIM$(STR$(Y))
   CALL FmtFindFLD("CTOT"): IF FOUND THEN RSET UTXT$(FFLD.UTXT) = LTRIM$(STR$(CTOT))
   CALL WinFORMAT(0)
   CURFMT = COMFMT: FMT = FMTTAB(CURFMT)
   IF NOT (Y + SLC) < CTMAX THEN
      CTMAX = CTMAX + 25
      REDIM PRESERVE COMRNO(CTMAX), PTRCHG(CTMAX), COMPTR(CTMAX) AS ChainPTRS
   END IF
   CALL FMTDREST
   CALL GetUFLD("FIRST"): FF = WIN.CURFLD
   WHILE (WIN.CURFLD > 0) AND (X <> 0)
      CALL FF3GetRec(X): IF NOT OKAY THEN GOTO COMURTRN
      COMRNO(Y) = X: PTRCHG(Y) = 0
      COMPTR(Y).BWD = FF3COMMENT.BPTR
      COMPTR(Y).FWD = FF3COMMENT.FPTR
      UTXT$(FFLD.UTXT) = FF3COMMENT.TEXT
      Z = Z + 1: Y = Y + 1
      X = FF3COMMENT.FPTR
      CALL GetUFLD("NEXT")
      IF WIN.CURFLD = FF THEN WIN.CURFLD = 0
   WEND
   IF Z > 0 THEN ZL = Z + TL - 1 ELSE ZL = TL
   WHILE Z <= SLC
      COMRNO(Y) = COMPTR(Y - 1).FWD: PTRCHG(Y) = 0
      COMPTR(Y).BWD = COMRNO(Y - 1): COMPTR(Y).FWD = 0
      Z = Z + 1: Y = Y + 1
   WEND
   ' Check for PGUP, PGDN at end of Previous WinUPDATE
   SELECT CASE R
      CASE 73: R = 0
               Y = COMRNO(0) + (WIN.CLIN - TL) - (BL - TL) * (1 - (WIN.CLIN = BL))
               IF Y < 1 THEN
                  IF COMRNO(0) > 1 THEN
                     WIN.CLIN = TL + COMRNO(0) - 1
                     COMRNO(0) = 1: GOTO BLDCPAGE
                     ELSE
                     CALL ErrBEEP(0)
                  END IF
                  ELSE
                  COMRNO(0) = Y: WIN.CLIN = BL
                  GOTO BLDCPAGE
               END IF
      CASE 81: R = 0
               Y = COMRNO(0) + (WIN.CLIN - TL) - (ZL - TL) * (WIN.CLIN = TL)
               IF NOT (COMRNO(Y) = 0) THEN
                  COMRNO(0) = Y: WIN.CLIN = TL
                  GOTO BLDCPAGE
                  ELSE
                  CALL ErrBEEP(0)
               END IF
   END SELECT
COMUPDATE:
   CALL FmtUHOLD
   CALL OptDISPLAY(FMT915)
   CALL WinFORMAT(COMFMT)
   CALL WinUPDATE: R = A
   ' Check for Line Deletion (F3)
   IF R = 61 THEN
      CALL GetUFLD("THIS")
      IF NOT FOUND THEN
         CALL ErrBEEP(0)
         ELSE
         COLOR FG(5), BG(5): LOCATE FLIN, BCOL, 0
         PRINT UTXT$(FFLD.UTXT);
         CALL PutMSG("UPDTM010")
         IF A$ = SNGLKEY$(1) THEN
            LSET UTXT$(FFLD.UTXT) = " "
            ELSE
            COLOR FG(FCOLOR), BG(FCOLOR)
            LOCATE FLIN, BCOL, 0: PRINT UTXT$(FFLD.UTXT);
         END IF
      END IF
   END IF
CHKCPAGE:
   ' First Get Record for NEW Lines and Delete Blank Lines
   Y = COMRNO(0): X = COMRNO(Y): Z = Y - 1
   CALL GetUFLD("FIRST"): FL = FFLD.FLDL \ 256
   WHILE (WIN.CURFLD > 0)
      IF X = 0 THEN
         IF (UTXT$(FFLD.UTXT) <> HTXT$(FFLD.UTXT)) THEN
            FF3PFX.RTYPE = CHR$(5)
            CALL FF3GetRec(0)
            IF OKAY THEN
               COMRNO(Y) = FF3PFX.RNO
               FF3COMMENT.SRTYPE = CHR$(SRTYPE)
               FF3COMMENT.SRNO = SRNO
               CALL FF3PutRec
               CTOT = CTOT + 1
            END IF
         END IF
         ELSE
         IF (UTXT$(FFLD.UTXT) = SPACE$(FL)) THEN
            CALL FF3DelRec(X): COMRNO(Y) = 0: CTOT = CTOT - 1
         END IF
      END IF
      IF COMRNO(Y) <> 0 THEN
         IF COMPTR(Z).FWD <> COMRNO(Y) THEN COMPTR(Z).FWD = COMRNO(Y): PTRCHG(Z) = -1
         IF Z > 0 THEN
            IF COMPTR(Y).BWD <> COMRNO(Z) THEN COMPTR(Y).BWD = COMRNO(Z): PTRCHG(Y) = -1
            ELSE
            IF COMPTR(Y).BWD <> 0 THEN COMPTR(Y).BWD = 0: PTRCHG(Y) = -1
         END IF
         Z = Y
      END IF
      Y = Y + 1: X = COMRNO(Y)
      CALL GetUFLD("NEXT")
      IF WIN.CURFLD = FF THEN WIN.CURFLD = 0
   WEND
   IF (COMPTR(Z).FWD <> X) THEN COMPTR(Z).FWD = X: PTRCHG(Z) = -1
   IF (X <> 0) AND (COMPTR(Y).BWD <> COMRNO(Z)) THEN COMPTR(Y).BWD = COMRNO(Z): PTRCHG(Y) = -1
   ' Next Update Text and Pointers in Records
   Y = COMRNO(0): X = COMRNO(Y): CLIN = WIN.CLIN
   IF (Y > 1) AND PTRCHG(Y - 1) THEN
      CALL FF3GetRec(COMRNO(Y - 1))
      IF OKAY THEN
         FF3COMMENT.FPTR = COMPTR(Y - 1).FWD
         CALL FF3PutRec: PTRCHG(Y - 1) = 0
      END IF
   END IF
   CALL GetUFLD("FIRST")
   WHILE (WIN.CURFLD > 0)
      IF X <> 0 THEN
         IF ((UTXT$(FFLD.UTXT) <> HTXT$(FFLD.UTXT)) OR PTRCHG(Y)) THEN
            CALL FF3GetRec(X): IF NOT OKAY THEN GOTO COMURTRN
            FF3COMMENT.BPTR = COMPTR(Y).BWD
            FF3COMMENT.FPTR = COMPTR(Y).FWD
            FF3COMMENT.TEXT = UTXT$(FFLD.UTXT)
            CALL FF3PutRec
            PTRCHG(Y) = 0
         END IF
         ELSE
         IF (FLIN < CLIN) AND (R <> 60) THEN WIN.CLIN = WIN.CLIN - 1
      END IF
      Y = Y + 1: X = COMRNO(Y)
      CALL GetUFLD("NEXT")
      IF WIN.CURFLD = FF THEN WIN.CURFLD = 0
   WEND
   IF (X <> 0) AND PTRCHG(Y) THEN
      CALL FF3GetRec(X): FF3COMMENT.BPTR = COMPTR(Y).BWD
      CALL FF3PutRec
      PTRCHG(Y) = 0
   END IF
   ' Check for Line Insertion (F2)
   IF R = 60 THEN
      Z = COMRNO(0) + BL - TL
      CALL GetUFLD("THIS")
      IF (NOT FOUND) OR ((COMRNO(Z) = 0) AND (COMRNO(Z + 1) <> 0)) THEN
         CALL ErrBEEP(0): GOTO COMUPDATE
         ELSE
         ZL = COMRNO(0) + WIN.CLIN - TL: Z = ZL
         ZTXT = FFLD.UTXT
         CALL GetUFLD("NEXT")
         IF WIN.CURFLD = FF THEN WIN.CURFLD = 0
         WHILE WIN.CURFLD <> 0
             Z = Z + 1
             SWAP COMRNO(ZL), COMRNO(Z)
             SWAP COMPTR(ZL), COMPTR(Z)
             SWAP UTXT$(ZTXT), UTXT$(FFLD.UTXT)
             CALL GetUFLD("NEXT")
             IF WIN.CURFLD = FF THEN WIN.CURFLD = 0
         WEND
         Z = Z + 1
         SWAP COMRNO(ZL), COMRNO(Z)
         SWAP COMPTR(ZL), COMPTR(Z)
         LSET HTXT$(ZTXT) = "": LSET UTXT$(ZTXT) = ""
         COMRNO(ZL) = 0: COMPTR(ZL).FWD = 0: COMPTR(ZL).BWD = 0
         GOTO COMUPDATE
      END IF
   END IF
   IF (R <> 27) AND (R <> 59) THEN GOTO BLDCPAGE
COMURTRN:
   COMBGN = COMPTR(0).FWD
   IF COMBGN <> 0 THEN
      CALL FF3GetRec(COMBGN): FF3COMMENT.CCNT = CTOT
      CALL FF3PutRec
   END IF
   CTMAX = 0: REDIM COMRNO(CTMAX), PTRCHG(CTMAX), COMPTR(CTMAX) AS ChainPTRS
   CALL WinCLR
END SUB

SUB ViewEVENTS (EVNTBGN, SRTYPE, SRNO)
    ' Process EVENT Information
    DFM = VAL(CONFIG.DFM)
    DIM EventTABLE AS TablePARMS, EDataTABLE AS TablePARMS
    EDataTABLE.FMTNAME = "UPDTS041": EDataTABLE.FMTNUM = 0
    EDataTABLE.XL = 90: EDATA$ = SPACE$(EDataTABLE.XL)
    TNUM = 0: FOUND = 0
    WHILE (TNUM < TABMAX) AND (NOT FOUND)
       TNUM = TNUM + 1: FOUND = TDFIndex(TNUM).NAME = "EVENTS  "
    WEND
    IF NOT FOUND THEN
       TNUM = 0
       ELSE
       EventTABLE = TABLE(TNUM)
       EventTABLE.FMTNAME = "UPDTS044": EventTABLE.FMTNUM = 0
    END IF
    CALL FmtFIND("UPDTS040"): FMT040 = CURFMT
    CALL FmtFIND("UPDTS041"): FMT041 = CURFMT
    CALL FmtFindFLD("DATE"): DTXT$(WIN.CURFLD) = XSCRNDate$(0, 0)
    CALL FmtFIND("UPDTS043"): FMT043 = CURFMT
    CALL FmtFIND("UPDTS044"): FMT044 = CURFMT
    CALL FmtFIND("XXXXS097"): FMTX97 = CURFMT
    CALL FmtFIND("UPDTS914"): FMT914 = CURFMT
    CALL FmtFIND("UPDTS918"): FMT918 = CURFMT

PrepEVENTS:
    EVTMAX = 10
    REDIM EVTCHN(EVTMAX) AS ChainPTRS, EVTDATA$(EVTMAX), EVNTRNO(EVTMAX)
    X = EVNTBGN: EVTIX = 0
    WHILE X <> 0
       EVTIX = EVTIX + 1
       CALL FF3GetRec(X)
       IF OKAY THEN
          EVNTRNO(EVTIX) = X
          LSET EDATA$ = ""
          GOSUB FindEVENT: GOSUB GetEDATA
          EVTDATA$(EVTIX) = EDATA$
          IF NOT (EVTIX < EVTMAX) THEN
             EVTMAX = EVTMAX + 20
             REDIM PRESERVE EVTCHN(EVTMAX) AS ChainPTRS, EVTDATA$(EVTMAX)
             REDIM PRESERVE EVNTRNO(EVTMAX)
          END IF
          X = FF3EVENT.NXT
          ELSE
          X = 0
       END IF
    WEND
    CALL ChnINIT(EVTCHN(), 1, EVTMAX)
    EDataTABLE.FT = -(EVTIX > 0)
    EDataTABLE.LT = EVTIX
    EDataTABLE.SIZE = EDataTABLE.LT
    EDataTABLE.FF = EVTIX + 1
    EVTCHN(EDataTABLE.SIZE).FWD = 0: EVTCHN(EDataTABLE.FF).BWD = 0
    GOTO ListEVENTS

GetEDATA:
    MID$(EDATA$, 1, 8) = FF3EVENT.CODE
    MID$(EDATA$, 9, 3) = SFILL$(ASC(FF3EVENT.IMPORT), 3)
    IF FF3EVENT.COM <> 0 THEN
       MID$(EDATA$, 12, 1) = SNGLKEY$(1)
       ELSE
       MID$(EDATA$, 12, 1) = " "
    END IF
    MID$(EDATA$, 13, 10) = XSCRNDate$(FF3EVENT.D1MD, FF3EVENT.D1Y)
    MID$(EDATA$, 23, 1) = FF3EVENT.DSTAT
    IF FF3EVENT.LPLNM = 0 THEN
       MID$(EDATA$, 24, 41) = FF3EVENT.PLACE
       ELSE
       CALL FF3GetRec(FF3EVENT.LPLNM)
       MID$(EDATA$, 24, 41) = FF3PLACE2.PLACE + SPACE$(19)
    END IF
    LSET EVTDATA$(EVTIX) = EDATA$
    RETURN

PutEDATA:
    LSET EDATA$ = EVTDATA$(EVTIX)
    FF3EVENT.CODE = MID$(EDATA$, 1, 8)
    X = VAL(MID$(EDATA$, 9, 3))
    IF X < 256 THEN FF3EVENT.IMPORT = CHR$(X)
    FF3EVENT.D1MD = DateMD(MID$(EDATA$, 13, 10), DFM)
    FF3EVENT.D1Y = DateY(MID$(EDATA$, 13, 10), DFM)
    FF3EVENT.DSTAT = MID$(EDATA$, 23, 1)
    EPLLTH = LEN(RTRIM$(MID$(EDATA$, 24, 41)))
    IF (EPLLTH > 22) AND (NOT ENV.LPNSupport) THEN EPLLTH = 22
    LSET FF3EVENT.PLACE = MID$(EDATA$, 24, 22)
    RETURN

FindEVENT:
    FOUND = 0
    MID$(EDATA$, 65, 25) = SPACE$(25)
    IF TNUM <> 0 THEN
       X = EventTABLE.FT
       WHILE (X > 0) AND (NOT FOUND)
          FOUND = (MID$(TABDATA$(X), 1, 8) = FF3EVENT.CODE)
          IF NOT FOUND THEN X = TABDATACHN(X).FWD
       WEND
    END IF
    IF FOUND THEN MID$(EDATA$, 65, 25) = MID$(TABDATA$(X), 13, 25)
    RETURN

GetEVENT:
    CALL WinSWITCH(2)
    CALL WinCLR: CALL WinFORMAT(FMT043)
    CALL OptDISPLAY(FMTX97)
    CALL TableOPEN(EventTABLE)
    CALL TableSHOW(EventTABLE, TABDATA$(), TABDATACHN())
    CALL TableSELECT(EventTABLE, TABDATA$(), TABDATACHN())
    IF A <> 27 THEN
       FF3EVENT.CODE = MID$(TABDATA$(EventTABLE.CT), 1, 8)
       MID$(EDATA$, 1, 8) = FF3EVENT.CODE
       GOSUB FindEVENT
    END IF
    CALL WinCLR
    CALL WinSWITCH(6)
    RETURN

ListEVENTS:
    CALL WinSWITCH(6)
    IF WINTAB(6).WFMT <> FMT041 THEN CALL WinCLR
    CURFMT = FMT040: FMT = FMTTAB(CURFMT)
    CALL FmtFindFLD("BRID"): IF FOUND THEN LSET UTXT$(FFLD.UTXT) = STR$(SRNO)
    CALL WinFORMAT(0): CALL OptDISPLAY(FMT918)
    EDataTABLE.CL = 0
    CALL TableOPEN(EDataTABLE)
    IF EVNTBGN = 0 THEN
       GOSUB AddEVENT
       IF EVNTBGN = 0 THEN GOTO EUpdtRETURN
    END IF
ListSHOW:
    CALL TableSHOW(EDataTABLE, EVTDATA$(), EVTCHN())
ListSELECT:
    CALL OptDISPLAY(FMT918)
    CALL TableSELECT(EDataTABLE, EVTDATA$(), EVTCHN())
    SELECT CASE A
       CASE 27: FUN$ = "ESC"
       CASE 59: FUN$ = "F1"
       CASE 60: FUN$ = "F2"
       CASE 61: FUN$ = "F3"
       CASE 62: FUN$ = "F4"
       CASE 63: FUN$ = "F5"
       CASE ELSE: CALL ErrBEEP(0): GOTO ListSELECT
    END SELECT
    CALL OptHILITE(FUN$, "", 6)
    EVTIX = EDataTABLE.CT
    EVTRNO = EVNTRNO(EVTIX)
    SELECT CASE FUN$
       CASE "ESC": GOTO EUpdtRETURN
       CASE "F1": GOTO UpdtEVENT
       CASE "F2": GOSUB AddEVENT: GOTO ListSHOW
       CASE "F3": GOSUB DeleteEVENT
                  IF EVTRNO = 0 THEN GOTO EUpdtRETURN
                  GOTO ListSHOW
       CASE "F4": GOSUB MoveEVENT: GOTO ListSELECT
       CASE "F5": CALL FF3GetRec(EVTRNO)
                  COMBGN = FF3EVENT.COM
                  CALL ViewCOMMENTS(COMBGN, 6, EVTRNO)
                  CALL WinSWITCH(6)
                  CALL FF3GetRec(EVTRNO)
                  IF COMBGN <> FF3EVENT.COM THEN
                     FF3EVENT.COM = COMBGN
                     CALL FF3PutRec
                  END IF
                  IF COMBGN <> 0 THEN
                     MID$(EVTDATA$(EDataTABLE.CT), 12, 1) = SNGLKEY$(1)
                     ELSE
                     MID$(EVTDATA$(EDataTABLE.CT), 12, 1) = " "
                  END IF
                  GOTO ListSHOW
    END SELECT
    GOTO ListSELECT

MoveEVENT:
   XBWD = EVTCHN(EDataTABLE.CT).BWD: XFWD = EVTCHN(EDataTABLE.CT).FWD
   CALL TableUPDT("MOV", EDataTABLE, EVTDATA$(), EVTCHN())
   IF XBWD = EVTCHN(EDataTABLE.CT).BWD THEN GOTO MoveRETURN
   IF XBWD > 0 THEN
      CALL FF3GetRec(EVNTRNO(XBWD))
      FF3EVENT.NXT = EVNTRNO(XFWD)
      CALL FF3PutRec
   END IF
   XBWD = EVTCHN(EDataTABLE.CT).BWD: XFWD = EVTCHN(EDataTABLE.CT).FWD
   IF XBWD > 0 THEN
      CALL FF3GetRec(EVNTRNO(XBWD))
      FF3EVENT.NXT = EVTRNO
      CALL FF3PutRec
   END IF
   CALL FF3GetRec(EVTRNO)
   FF3EVENT.NXT = EVNTRNO(XFWD)
   CALL FF3PutRec
   EVNTBGN = EVNTRNO(EDataTABLE.FT)
   EVTIX = EDataTABLE.CT
MoveRETURN:
   RETURN

UpdtEVENT:
   CALL OptDISPLAY(FMT914)
   CALL TableUPDT("CHG", EDataTABLE, EVTDATA$(), EVTCHN())
   SELECT CASE A
      CASE 13: FUN$ = "F1"
      CASE 27: FUN$ = "ESC"
      CASE 59: FUN$ = "F1"
      CASE 60: IF TNUM > 0 THEN FUN$ = "F2"
      CASE ELSE: CALL ErrBEEP(0): GOTO UpdtEVENT
   END SELECT
   CALL OptHILITE(FUN$, "", 6)
   SELECT CASE FUN$
      CASE "ESC": CALL FF3GetRec(EVTRNO): GOSUB GetEDATA: GOTO ListSHOW
      CASE "F1": GOSUB WriteEVENT: GOTO ListSELECT
      CASE "F2": LSET EDATA$ = EVTDATA$(EDataTABLE.CT)
                 GOSUB GetEVENT
                 LSET EVTDATA$(EDataTABLE.CT) = EDATA$
                 CALL TableSHOW(EDataTABLE, EVTDATA$(), EVTCHN())
   END SELECT
   GOTO UpdtEVENT

AddEVENT:
   CALL OptHILITE("F2", "", 6)
   FF3PFX.RTYPE = CHR$(6)
   CALL FF3GetRec(0): IF NOT OKAY THEN RETURN
   FF3EVENT.SRTYPE = CHR$(SRTYPE): FF3EVENT.SRNO = SRNO
   CALL FF3PutRec
   IF EDataTABLE.FF = 0 THEN
      EDataTABLE.FF = EVTMAX + 1: EVTMAX = EVTMAX + 20
      REDIM PRESERVE EVTCHN(EVTMAX) AS ChainPTRS, EVTDATA$(EVTMAX)
      REDIM PRESERVE EVNTRNO(EVTMAX)
      CALL ChnINIT(EVTCHN(), EDataTABLE.FF, EVTMAX)
   END IF
   CALL TableUPDT("NEW", EDataTABLE, EVTDATA$(), EVTCHN())
   CALL TableSHOW(EDataTABLE, EVTDATA$(), EVTCHN())
   CALL TableKEY("MARK", EDataTABLE, EVTDATA$())
   EVTIX = EDataTABLE.CT
   GOSUB GetEVENT: IF A = 27 THEN GOTO AddRETURN
   GOSUB GetEDATA
   CALL TableSHOW(EDataTABLE, EVTDATA$(), EVTCHN())
   CALL TableUPDT("CHG", EDataTABLE, EVTDATA$(), EVTCHN())
AddRETURN:
   IF A = 27 THEN
      CALL FF3DelRec(FF3PFX.RNO)
      CALL TableUPDT("REM", EDataTABLE, EVTDATA$(), EVTCHN())
      ELSE
      EVTRNO = FF3PFX.RNO: EVTIX = EDataTABLE.CT
      EVNTRNO(EVTIX) = EVTRNO
      GOSUB WriteEVENT
   END IF
   RETURN

DeleteEVENT:
   CALL PutMSG("UPDTM007")
   IF A$ <> SNGLKEY$(1) THEN GOTO DeleteRETURN
   CALL FF3GetRec(EVTRNO)
   ECOM = FF3EVENT.COM: LPLNM = FF3EVENT.LPLNM
   IF ECOM <> 0 THEN
      CALL PutMSG("UPDTM009")
      IF A$ <> SNGLKEY$(1) THEN GOTO DeleteRETURN
   END IF
   WHILE ECOM <> 0
       CALL FF3GetRec(ECOM)
       IF NOT OKAY THEN
          ECOM = 0
          ELSE
          X = FF3COMMENT.FPTR
          CALL FF3DelRec(ECOM)
          IF NOT OKAY THEN
             ECOM = 0
             ELSE
             ECOM = X
          END IF
       END IF
   WEND
   XBWD = EVTCHN(EDataTABLE.CT).BWD
   XFWD = EVTCHN(EDataTABLE.CT).FWD
   IF XBWD <> 0 THEN
      CALL FF3GetRec(EVNTRNO(XBWD))
      FF3EVENT.NXT = EVNTRNO(XFWD)
      CALL FF3PutRec
      ELSE
      EVNTBGN = EVNTRNO(XFWD)
   END IF
   IF LPLNM <> 0 THEN CALL FF3DelRec(LPLNM)
   CALL FF3DelRec(EVTRNO)
   EVTRNO = EVNTRNO(XFWD)
   IF EVTRNO = 0 THEN EVTRNO = EVNTRNO(XBWD)
   CALL TableUPDT("REM", EDataTABLE, EVTDATA$(), EVTCHN())
DeleteRETURN:
   RETURN

WriteEVENT:
   CALL FF3GetRec(EVTRNO)
   GOSUB PutEDATA
   XBWD = EVTCHN(EDataTABLE.CT).BWD
   XFWD = EVTCHN(EDataTABLE.CT).FWD
   FF3EVENT.NXT = EVNTRNO(XFWD)
   CALL FF3PutRec
   IF EPLLTH < 23 THEN
      IF FF3EVENT.LPLNM <> 0 THEN CALL FF3DelRec(FF3EVENT.LPLNM)
      ELSE
      FF3PFX.RTYPE = CHR$(12)
      CALL FF3GetRec(FF3EVENT.LPLNM)
      FF3PLACE2.SRTYPE = CHR$(6)
      FF3PLACE2.SRNO = EVTRNO
      LSET FF3PLACE2.PLACE = MID$(EDATA$, 24, 41)
      CALL FF3PutRec
      IF FF3EVENT.LPLNM = 0 THEN
         FF3EVENT.LPLNM = FF3PFX.RNO
         FF3PFX.RTYPE = CHR$(6): FF3PFX.RNO = EVTRNO
         CALL FF3PutRec
      END IF
   END IF
   IF XBWD <> 0 THEN
      CALL FF3GetRec(EVNTRNO(XBWD))
      FF3EVENT.NXT = EVTRNO
      CALL FF3PutRec
      ELSE
      EVNTBGN = EVTRNO
   END IF
   RETURN

EUpdtRETURN: ' Return to Calling Routine
   CALL PutMSG("")
   CALL WinCLR
END SUB

SUB ViewMARRIAGE (SPID1, MRNO)
SHARED FUN$, RGL
SHARED OMPLACE AS STRING * 41, ODVPLACE AS STRING * 41
   DIM OSPOUSE AS OthrFamilyREC
   CALL FmtFIND("UPDTS010"): FMT010 = CURFMT
   CALL FmtFIND("UPDTS011"): FMT011 = CURFMT
   CALL FmtFindFLD("BMDT"): LSET DTXT$(WIN.CURFLD) = XSCRNDate$(0, 0)
   CALL FmtFindFLD("EMDT"): LSET DTXT$(WIN.CURFLD) = XSCRNDate$(0, 0)
   CALL FmtFIND("UPDTS908"): FMT908 = CURFMT
   CALL FmtFIND("UPDTS909"): FMT909 = CURFMT
   CALL FmtFIND("UPDTS910"): FMT910 = CURFMT
   CALL WinSWITCH(2)
   CALL FF1GetRec(SPID1): SP1OLDCH = FF1REC.OLDCH
   DFM = VAL(CONFIG.DFM)
   DD1Y = FF1REC.DY
   IF DD1Y > 0 THEN IF DD1Y < 9999 THEN DD1MD = FF1REC.DMD ELSE DD1Y = 0
   IF WIN.WFMT <> FMT010 THEN CALL WinCLR
   CURFMT = FMT010: FMT = FMTTAB(CURFMT)
   CALL FmtFindFLD("RID"): DTXT$(WIN.CURFLD) = STR$(SPID1)
   CALL FmtFindFLD("RGL"): RSET DTXT$(WIN.CURFLD) = STR$(RGL)
   CALL WinDisplayLABELS(FMT010)
A900: ' Build List of Marriages
   MRIX = 0: MRHIX = 10: REDIM MRCHN(MRHIX)
   IF MRNO = 0 THEN GOTO 907
   X = FF1REC.SPOUSE: FOUND = 0
   WHILE (X <> 0) AND (NOT FOUND)
      FOUND = (X = MRNO)
      IF NOT FOUND THEN
         MRCHN(MRIX) = X: MRIX = MRIX + 1
         IF NOT (MRIX < MRHIX) THEN
            MRHIX = MRHIX + 5: REDIM PRESERVE MRCHN(MRHIX)
         END IF
         CALL FF3GetRec(X)
         IF NOT OKAY THEN
            X = 0
            ELSE
            IF SPID1 = FF3SPOUSE.SP1ID THEN
               X = FF3SPOUSE.SP1NXT
               ELSE
               IF SPID1 = FF3SPOUSE.SP2ID THEN
                  X = FF3SPOUSE.SP2NXT
                  ELSE
                  X = 0
               END IF
            END IF
         END IF
      END IF
   WEND
   IF NOT FOUND THEN GOTO 990
   GOTO 903

901 ' Display Spouse Information
   CURFMT = FMT010: FMT = FMTTAB(CURFMT)
   CALL FMTDREST
   IF SPID2 > 0 THEN
      CALL FF1GetRec(SPID2): SP2OLDCH = FF1REC.OLDCH: DDY = FF1REC.DY
      IF DDY > 0 THEN IF DDY < 9999 THEN DDMD = FF1REC.DMD ELSE DDY = 0
      ELSE
      DDY = 0: SP2OLDCH = 0
   END IF
   IF (0 < DD1Y) AND (DD1Y < DDY) THEN DDY = DD1Y: DDMD = DD1MD
   FF = FMT.BGNFLD
   WHILE FF > 0
      FFLD = FLDTAB(FF)
      IF FFLD.UTXT > 0 THEN
         X$ = UTXT$(FFLD.UTXT)
         SELECT CASE FFLD.NAME
            CASE "SPID": LSET X$ = STR$(SPID2)
            CASE "SPGV": IF SPID2 > 0 THEN LSET X$ = FF1REC.GIVEN
            CASE "SPSN": IF SPID2 > 0 THEN LSET X$ = FF1REC.SURNM
            CASE "AGE ": IF SPID2 > 0 THEN LSET X$ = DateDif$(FF1REC.BMD, FF1REC.BY, FF1REC.DMD, FF1REC.DY)
            CASE "BDT ": IF SPID2 > 0 THEN LSET X$ = XSCRNDate$(FF1REC.BMD, FF1REC.BY)
            CASE "BDST": IF SPID2 > 0 THEN LSET X$ = FDSTAT$("BD")
            CASE "DDT ": IF SPID2 > 0 THEN LSET X$ = XSCRNDate$(FF1REC.DMD, FF1REC.DY)
            CASE "DDST": IF SPID2 > 0 THEN LSET X$ = FDSTAT$("DD")
         END SELECT
         LSET UTXT$(FFLD.UTXT) = X$
      END IF
      FF = FLDCHN(FF).FWD
   WEND
   CALL WinDisplayDATA(FMT010)
   RETURN

902 ' Display Marriage STATUS Information from Family Record
   CURFMT = FMT011: FMT = FMTTAB(CURFMT)
   FF = FMT.BGNFLD
   WHILE FF > 0
      FFLD = FLDTAB(FF)
      IF FFLD.UTXT > 0 THEN
         X$ = UTXT$(FFLD.UTXT)
         SELECT CASE FFLD.NAME
            CASE "ANNV": D2MD = FF3SPOUSE.D2MD: D2Y = FF3SPOUSE.D2Y
                         IF D2Y = 0 THEN D2Y = DDY: IF D2Y = 0 THEN D2MD = 0
                         LSET X$ = DateDif$(FF3SPOUSE.D1MD, FF3SPOUSE.D1Y, D2MD, D2Y)
            CASE "BMST": LSET X$ = FF3SPOUSE.STATUS1
            CASE "BMDT": LSET X$ = XSCRNDate$(FF3SPOUSE.D1MD, FF3SPOUSE.D1Y)
            CASE "BMPL": LSET X$ = MPLACE: FLDTAB(FF).DO = 0
            CASE "EMST": LSET X$ = FF3SPOUSE.STATUS2
            CASE "EMDT": LSET X$ = XSCRNDate$(FF3SPOUSE.D2MD, FF3SPOUSE.D2Y)
            CASE "EMPL": LSET X$ = DVPLACE: FLDTAB(FF).DO = 0
         END SELECT
         LSET UTXT$(FFLD.UTXT) = X$
      END IF
      FF = FLDCHN(FF).FWD
   WEND
   CALL WinFORMAT(FMT011)
   RETURN

903 ' Options for Existing Family Record
   IF MRNO = 0 THEN
      IF SPID2 = 0 THEN CALL WinCLR
      GOTO 990
   END IF
   MRCHN(MRIX) = MRNO
   IF NOT (MRIX < MRHIX) THEN
      MRHIX = MRHIX + 5: REDIM PRESERVE MRCHN(MRHIX)
   END IF
   CALL FF3GetRec(MRNO): IF NOT OKAY THEN GOTO 990
   IF FF3SPOUSE.SP1ID = SPID1 THEN
      SPID2 = FF3SPOUSE.SP2ID: NM = FF3SPOUSE.SP1NXT
      ELSE
      SPID2 = FF3SPOUSE.SP1ID: NM = FF3SPOUSE.SP2NXT
   END IF
   MRCHN(MRIX + 1) = NM
   LSET MPLACE = "": LSET DVPLACE = ""
   IF FF3SPOUSE.MLOC <> 0 THEN
      CALL FF3GetRec(FF3SPOUSE.MLOC)
      IF FF3PFX.RTYPE = CHR$(12) THEN
         IF FF3PLACE2.PNEXT <> 0 THEN CALL FF3GetRec(FF3PLACE2.PNEXT)
      END IF
   END IF
   LSET OMPLACE = MPLACE: LSET ODVPLACE = DVPLACE
   OSPID = SPID2
   OSPOUSE = FF3SPOUSE
   D1MD = FF3SPOUSE.D1MD: D1Y = FF3SPOUSE.D1Y
   GOSUB 901: GOSUB 902

B903: ' Get Process Option
   XCHANGED = 0
   CALL OptDISPLAY(FMT908)
   IF (SP1OLDCH <> 0) AND (SP2OLDCH <> 0) THEN CALL OptHILITE("F3", "", 7)
   IF FF3SPOUSE.LOC <> 0 THEN CALL OptHILITE("F4", "", 7)
   IF FF3SPOUSE.COM <> 0 THEN CALL OptHILITE("F5", "", 7)
   IF (MRIX > 0) THEN
      IF (MRCHN(MRIX + 1) <> 0) THEN
         CALL PutMSG("UPDTM004")
         ELSE
         CALL PutMSG("UPDTM005")
      END IF
      ELSE
      IF MRCHN(MRIX + 1) <> 0 THEN
         CALL PutMSG("UPDTM006")
      END IF
   END IF
   CALL GetKEY: FUN$ = ""
   CALL PutMSG("")
   SELECT CASE A
      CASE 27: GOTO 990
      CASE 59: FUN$ = "F1"  ' Change Record
      CASE 60: IF SPID2 <> 0 THEN FUN$ = "F2"   ' Select SPOUSE for BRID
      CASE 61: IF (SP1OLDCH <> 0) AND (SP2OLDCH <> 0) THEN FUN$ = "F3"  ' Display Children
      CASE 62: FUN$ = "F4"  ' Process Addresses
      CASE 63: FUN$ = "F5"  ' Process Comments
      CASE 73: FUN$ = "PGUP"
               IF MRCHN(MRIX + 1) <> 0 THEN
                  MRIX = MRIX + 1
                  MRNO = MRCHN(MRIX)
                  GOTO 903
                  ELSE
                  FUN$ = ""
               END IF
      CASE 81: FUN$ = "PGDN"
               IF MRIX > 0 THEN
                  MRIX = MRIX - 1
                  MRNO = MRCHN(MRIX)
                  GOTO 903
                  ELSE
                  FUN$ = ""
               END IF
   END SELECT
   IF FUN$ = "" THEN CALL ErrBEEP(0): GOTO B903
   CALL OptHILITE(FUN$, "", 6)
   SELECT CASE FUN$
      CASE "F1": NMRNO = MRNO: GOTO 905
      CASE "F2": SPID1 = SPID2: CALL WinCLR: FUN$ = "SEL": EXIT SUB
      CASE "F3": FUN$ = "CHLD": EXIT SUB
      CASE "F4": ADRBGN = FF3SPOUSE.LOC
                 CALL ViewADDRESS(ADRBGN, 3, MRNO)
                 CALL WinSWITCH(2)
                 CALL FF3GetRec(MRNO)
                 IF ADRBGN <> FF3SPOUSE.LOC THEN
                    FF3SPOUSE.LOC = ADRBGN: CALL FF3PutRec
                 END IF
      CASE "F5": COMBGN = FF3SPOUSE.COM
                 CALL ViewCOMMENTS(COMBGN, 3, MRNO)
                 CALL WinSWITCH(2)
                 CALL FF3GetRec(MRNO)
                 IF COMBGN <> FF3SPOUSE.COM THEN
                    FF3SPOUSE.COM = COMBGN: CALL FF3PutRec
                 END IF
   END SELECT
   OSPOUSE = FF3SPOUSE
   GOSUB C903: GOTO B903

C903: ' Update Change FLAGs in Couple's Name Records
   IF XCHANGED THEN
      IF (SPID1 <> 0) THEN
         CALL FF1GetRec(SPID1): XCHANGED = -1: CALL FF1PutREC
      END IF
      IF (OSPID <> 0) AND (OSPID <> SPID2) AND (OSPID <> SPID1) THEN
         CALL FF1GetRec(OSPID): XCHANGED = -1: CALL FF1PutREC
      END IF
      IF (SPID2 <> 0) THEN
         CALL FF1GetRec(SPID2): XCHANGED = -1: CALL FF1PutREC
      END IF
   END IF
   XCHANGED = 0
   RETURN

905 ' Change OPTIONS
   CALL OptDISPLAY(FMT909)
   CALL GetKEY
   SELECT CASE A
      CASE 27: GOTO 903
      CASE 59: GOTO 906   ' Change
      CASE 60: GOTO 907   ' Add
      CASE 61: GOTO 908   ' Delete
      CASE ELSE: CALL ErrBEEP(0): GOTO 905
   END SELECT

906 ' Modify Current Marriage Record
   CALL OptDISPLAY(FMT910)
   GOSUB C906
   SELECT CASE A
      CASE 27: GOTO 903
      CASE 59: GOTO A906   ' Save Changes
      CASE 60: GOSUB B906: GOTO 906  ' Change Spouse ID
      CASE ELSE: CALL ErrBEEP(0): GOTO 906
   END SELECT

A906: ' Save Changes
   CALL OptHILITE("F1", "", 6)
   FF3PFX.RTYPE = CHR$(3): FF3PFX.RNO = NMRNO: CALL FF3PutRec
   XLOC = FF3SPOUSE.MLOC
   CALL UpdtPLACE(XLOC, 3, NMRNO)
   CALL FF3GetRec(NMRNO)
   FF3SPOUSE.MLOC = XLOC
   CALL FF3PutRec
   IF (NMRNO <> MRNO) OR (SPID2 <> OSPID) OR (FF3SPOUSE.D1MD <> D1MD) OR (FF3SPOUSE.D1Y <> D1Y) THEN
      IF (NMRNO = MRNO) THEN
         CALL UpdtFamPTRS("REM", SPID1, NMRNO)
         IF OSPID <> 0 THEN CALL UpdtFamPTRS("REM", OSPID, NMRNO)
      END IF
      IF SPID2 <> OSPID THEN
         CALL FF3GetRec(NMRNO)
         IF SPID1 = FF3SPOUSE.SP1ID THEN
            FF3SPOUSE.SP2ID = SPID2
            ELSE
            IF SPID1 = FF3SPOUSE.SP2ID THEN FF3SPOUSE.SP1ID = SPID2
         END IF
         CALL FF3PutRec
      END IF
      IF SPID2 <> 0 THEN
         CALL UpdtFamPTRS("ADD", SPID2, NMRNO)
      END IF
      CALL UpdtFamPTRS("ADD", SPID1, NMRNO)
      CALL FF1GetRec(SPID1)
      MRNO = NMRNO
   END IF
   GOSUB C903
   CALL FF1GetRec(SPID1)
   GOTO A900

B906: ' Modify SPOUSE ID
   CALL OptHILITE("F2", "", 6)
   CURFMT = FMT010: FMT = FMTTAB(CURFMT)
   CALL FmtUHOLD
B906B:
   CALL PutMSG("XXXXM011")
   WIN.CLIN = 0: CALL WinUPDATE
   CALL PutMSG("")
   CALL FmtFindFLD("SPID")
   IF A = 27 THEN
      CALL FMTHREST: CALL WinDisplayDATA(0): RETURN
      ELSE
      X& = VAL(UTXT$(FFLD.UTXT)): CALL ChkID(X&)
      IF NOT OKAY THEN GOTO B906B
      SPID2 = X&
      IF SPID2 = SPID1 THEN CALL PutMSG("UPDTM003"): GOTO B906B
      IF SPID2 > 0 THEN CALL FF1GetRec(SPID2): IF NOT OKAY THEN GOTO B906B
   END IF
   GOSUB 901
   RETURN

C906: ' Modify Marriage STATUS Info
   CURFMT = FMT011: FMT = FMTTAB(CURFMT)
   CALL FmtUHOLD
   CALL PutMSG("XXXXM012")
   WIN.CLIN = 0: CALL WinUPDATE
   CALL PutMSG("")
   IF A = 27 THEN CALL FMTHREST: CALL WinFORMAT(0): RETURN
   FF = FMT.BGNFLD
   WHILE FF > 0
      FFLD = FLDTAB(FF)
      IF FFLD.UTXT > 0 THEN
         X$ = UTXT$(FFLD.UTXT)
         SELECT CASE FFLD.NAME
            CASE "BMST": FF3SPOUSE.STATUS1 = X$
            CASE "BMDT": FF3SPOUSE.D1MD = DateMD(X$, DFM)
                         FF3SPOUSE.D1Y = DateY(X$, DFM)
            CASE "BMPL": MPLACE = X$
            CASE "EMST": FF3SPOUSE.STATUS2 = X$
            CASE "EMDT": FF3SPOUSE.D2MD = DateMD(X$, DFM)
                         FF3SPOUSE.D2Y = DateY(X$, DFM)
            CASE "EMPL": DVPLACE = X$
         END SELECT
      END IF
      FF = FLDCHN(FF).FWD
   WEND
   CALL FmtFindFLD("ANNV")
   IF FOUND THEN
      D2MD = FF3SPOUSE.D2MD: D2Y = FF3SPOUSE.D2Y
      IF D2Y = 0 THEN D2Y = DDY: IF D2Y = 0 THEN D2MD = 0
      LSET UTXT$(FFLD.UTXT) = DateDif$(FF3SPOUSE.D1MD, FF3SPOUSE.D1Y, D2MD, D2Y)
   END IF
   CALL WinFORMAT(0)
   RETURN

907 ' Add Marriage Record
   FF3PFX.RTYPE = CHR$(3)
   CALL FF3GetRec(0)
   IF NOT OKAY THEN GOTO 903
   NMRNO = FF3PFX.RNO
   FF3SPOUSE.SRTYPE = CHR$(1)
   FF3SPOUSE.SP1ID = SPID1
   CALL FF3PutRec
   CURFMT = FMT010: FMT = FMTTAB(CURFMT)
   CALL FMTDREST
   CURFMT = FMT011: FMT = FMTTAB(CURFMT)
   CALL FMTDREST: SPID2 = 0: OSPID = 0
   CALL WinDisplayDATA(FMT010)
   CALL WinDisplayDATA(FMT011)
   CALL OptDISPLAY(FMT910)
   GOSUB B906: IF A = 27 THEN CALL FF3DelRec(NMRNO): GOTO 903
A907:
   CALL OptDISPLAY(0)
   GOSUB C906
   SELECT CASE A
      CASE 27: CALL FF3DelRec(NMRNO): GOTO 903
      CASE 59: MRNO = 0: GOTO A906
      CASE 60: GOSUB B906: GOTO A907
      CASE ELSE: CALL ErrBEEP(0): GOTO A907
   END SELECT

908 ' Delete Marriage Record
   CALL OptHILITE("F3", "", 6)
   CALL PutMSG("UPDTM007")
   IF (A = 27) OR (A$ = SNGLKEY$(2)) THEN GOTO 905
   IF A$ <> SNGLKEY$(1) THEN CALL ErrBEEP(0): GOTO 908
   IF OSPID = 0 THEN
      MADR = FF3SPOUSE.LOC
      IF MADR <> 0 THEN
         CALL PutMSG("UPDTM008")
         IF A$ <> SNGLKEY$(1) THEN GOTO 905
      END IF
      MCOM = FF3SPOUSE.COM
      IF MCOM <> 0 THEN
         CALL PutMSG("UPDTM009")
         IF A$ <> SNGLKEY$(1) THEN GOTO 905
      END IF
   END IF
   IF OSPID = 0 THEN
      MADR = FF3SPOUSE.LOC: MCOM = FF3SPOUSE.COM
      IF FF3SPOUSE.MLOC <> 0 THEN
         CALL FF3GetRec(FF3SPOUSE.MLOC)
         IF FF3PFX.RTYPE = CHR$(12) THEN
            IF FF3PLACE2.PNEXT <> 0 THEN CALL FF3DelRec(FF3PLACE2.PNEXT)
         END IF
         CALL FF3DelRec(FF3SPOUSE.MLOC)
      END IF
      X = MADR
      WHILE X <> 0
         CALL FF2GetRec(X)
         IF NOT OKAY THEN
            X = 0
            ELSE
            Y = FF2PFX.COM
            WHILE Y <> 0
               CALL FF3GetRec(Y)
               IF NOT OKAY THEN
                  Y = 0
                  ELSE
                  NY = FF3COMMENT.FPTR
                  CALL FF3DelRec(Y)
                  IF NOT OKAY THEN NY = 0
                  Y = NY
               END IF
            WEND
            NX = FF2PFX.NXT
            CALL FF2DelRec(X)
            IF NOT OKAY THEN NX = 0
            X = NX
         END IF
      WEND
      X = MCOM
      WHILE X <> 0
         CALL FF3GetRec(X)
         IF NOT OKAY THEN
            X = 0
            ELSE
            NX = FF3COMMENT.FPTR
            CALL FF3DelRec(X)
            IF NOT OKAY THEN NX = 0
            X = NX
         END IF
      WEND
   END IF
   CALL UpdtFamPTRS("REM", SPID1, NMRNO)
   IF OSPID = 0 THEN
      CALL FF3DelRec(NMRNO)
      ELSE
      CALL FF3GetRec(NMRNO)
      IF FF3SPOUSE.SP1ID = SPID1 THEN
         FF3SPOUSE.SP1ID = 0: FF3SPOUSE.SP1NXT = 0
         ELSE
         IF FF3SPOUSE.SP2ID = SPID1 THEN
            FF3SPOUSE.SP2ID = 0: FF3SPOUSE.SP2NXT = 0
         END IF
      END IF
      CALL FF3PutRec
   END IF

   IF OKAY THEN
      MRNO = MRCHN(MRIX + 1)
      IF (MRNO = 0) AND (MRIX > 0) THEN
         MRIX = MRIX - 1: MRNO = MRCHN(MRIX)
      END IF
   END IF
   GOSUB C903
   GOTO 903

990 ' ESCAPE Exit
    'CALL WinCLR
    FUN$ = "ESC"

END SUB

SUB ViewMISC (ORTYPE, BRID)
6000 ' Process Miscellaneous Record Types
   DFM = VAL(CONFIG.DFM)
   CALL FmtFIND("UPDTS909"): FMT909 = CURFMT
   CALL FmtFIND("UPDTS913"): FMT913 = CURFMT
   SELECT CASE ORTYPE
      CASE 7: CALL FmtFIND("UPDTS021"): OBGN = FF1REC.EDUC
      CASE 8: CALL FmtFIND("UPDTS022"): OBGN = FF1REC.WORK
      CASE 9: CALL FmtFIND("UPDTS023"): OBGN = FF1REC.MIL
      CASE 10: CALL FmtFIND("UPDTS024"): OBGN = FF1REC.HLTH
      CASE ELSE: GOTO 6029
   END SELECT
   OTHFMT = CURFMT
   CALL FmtFindFLD("BRID"): DTXT$(WIN.CURFLD) = STR$(BRID)
   CALL FmtFindFLD("DAT1"): LSET DTXT$(WIN.CURFLD) = XSCRNDate$(0, 0)
   CALL FmtFindFLD("DAT2"): LSET DTXT$(WIN.CURFLD) = XSCRNDate$(0, 0)
   OTHIX = 0: OTHHIX = 10
   REDIM OTHCHN(OTHHIX)
   CALL WinSWITCH(2)
   CALL WinCLR
   CALL WinDisplayLABELS(OTHFMT)
   ORNO = OBGN
   IF ORNO = 0 THEN
      CALL OptDISPLAY(FMT909)
      CALL OptHILITE("F2", "", 6)
      GOTO 6020
   END IF
   GOTO 6005

6002 ' Move OTH Info to Screen
   CURFMT = OTHFMT: FMT = FMTTAB(OTHFMT)
   CALL GetUFLD("FIRST"): FF = WIN.CURFLD
   WHILE WIN.CURFLD > 0
      X$ = UTXT$(FFLD.UTXT)
      SELECT CASE FFLD.NAME
         CASE "BRID": LSET X$ = STR$(BRID)
         CASE "DAT1": LSET X$ = XSCRNDate$(FF3PNTRS.D1MD, FF3PNTRS.D1Y)
         CASE "DAT2": LSET X$ = XSCRNDate$(FF3PNTRS.D2MD, FF3PNTRS.D2Y)
         CASE "EDLV": LSET X$ = FF3EDUC.EDLEV
         CASE "EDEG": LSET X$ = FF3EDUC.EDDEG
         CASE "EDS1": LSET X$ = FF3EDUC.EDSUB1
         CASE "EDS2": LSET X$ = FF3EDUC.EDSUB2
         CASE "WORK": LSET X$ = FF3WORK.WORKTP
         CASE "RANK": LSET X$ = FF3MIL.MILRANK
         CASE "MLST": LSET X$ = FF3MIL.MILSTAT
         CASE "DIAG": LSET X$ = FF3HEALTH.HLTHDIAG
         CASE "HLST": LSET X$ = FF3HEALTH.HLTHSTAT
      END SELECT
      LSET UTXT$(FFLD.UTXT) = X$
      CALL GetUFLD("NEXT")
      IF FF = WIN.CURFLD THEN WIN.CURFLD = 0
   WEND
   CALL WinFORMAT(OTHFMT)
   OD1Y = FF3PNTRS.D1Y: OD1MD = FF3PNTRS.D1MD
   CALL FmtUHOLD
   RETURN

6004 ' Move OTH Info from Screen
   CURFMT = OTHFMT: FMT = FMTTAB(OTHFMT)
   CALL GetUFLD("FIRST"): FF = WIN.CURFLD
   WHILE WIN.CURFLD > 0
      X$ = UTXT$(FFLD.UTXT)
      SELECT CASE FFLD.NAME
         CASE "DAT1": FF3PNTRS.D1MD = DateMD(X$, DFM)
                      FF3PNTRS.D1Y = DateY(X$, DFM)
         CASE "DAT2": FF3PNTRS.D2MD = DateMD(X$, DFM)
                      FF3PNTRS.D2Y = DateY(X$, DFM)
         CASE "EDLV": LSET FF3EDUC.EDLEV = X$
         CASE "EDEG": LSET FF3EDUC.EDDEG = X$
         CASE "EDS1": LSET FF3EDUC.EDSUB1 = X$
         CASE "EDS2": LSET FF3EDUC.EDSUB2 = X$
         CASE "WORK": LSET FF3WORK.WORKTP = X$
         CASE "RANK": LSET FF3MIL.MILRANK = X$
         CASE "MLST": LSET FF3MIL.MILSTAT = X$
         CASE "DIAG": LSET FF3HEALTH.HLTHDIAG = X$
         CASE "HLST": LSET FF3HEALTH.HLTHSTAT = X$
      END SELECT
      CALL GetUFLD("NEXT")
      IF FF = WIN.CURFLD THEN WIN.CURFLD = 0
   WEND
   RETURN

6005 ' Process Existing Record
   CALL FF3GetRec(ORNO): IF NOT OKAY THEN GOTO 6029
   OTHCHN(OTHIX) = ORNO
   OTHCHN(OTHIX + 1) = FF3PNTRS.NXT
   GOSUB 6002
6006 ' Get Next Request
   CALL OptDISPLAY(FMT913)
   IF FF3PNTRS.LOC <> 0 THEN CALL OptHILITE("F2", "", 7)
   IF FF3PNTRS.COM <> 0 THEN CALL OptHILITE("F3", "", 7)
    IF (OTHIX > 0) THEN
       IF (OTHCHN(OTHIX + 1) <> 0) THEN
          CALL PutMSG("UPDTM004")
          ELSE
          CALL PutMSG("UPDTM005")
       END IF
       ELSE
       IF OTHCHN(OTHIX + 1) <> 0 THEN
          CALL PutMSG("UPDTM006")
          ELSE
          CALL PutMSG("")
       END IF
    END IF
    CALL GetKEY: FUN$ = ""
    SELECT CASE A
       CASE 27: GOTO 6029
       CASE 59: FUN$ = "F1"  ' Update
       CASE 60: FUN$ = "F2"  ' Addresses
       CASE 61: FUN$ = "F3"  ' CommentS
       CASE 73: FUN$ = "PGUP"
                IF OTHCHN(OTHIX + 1) <> 0 THEN
                   OTHIX = OTHIX + 1
                   ORNO = OTHCHN(OTHIX)
                   IF NOT (OTHIX < OTHHIX) THEN
                      OTHHIX = OTHHIX + 5
                      REDIM PRESERVE OTHCHN(OTHHIX)
                   END IF
                   GOTO 6005
                   ELSE
                   FUN$ = ""
                END IF
       CASE 81: FUN$ = "PGDN"
                IF OTHIX > 0 THEN
                   OTHIX = OTHIX - 1
                   ORNO = OTHCHN(OTHIX)
                   GOTO 6005
                   ELSE
                   FUN$ = ""
                END IF
    END SELECT
    IF FUN$ = "" THEN CALL ErrBEEP(0): GOTO 6006
    CALL OptHILITE(FUN$, "", 6)
    SELECT CASE FUN$
       CASE "F1": GOTO 6010
       CASE "F2": ABGN = FF3PNTRS.LOC
                  CALL ViewADDRESS(ABGN, ORTYPE, ORNO)
                  CALL WinSWITCH(2)
                  CALL FF3GetRec(ORNO)
                  IF ABGN <> FF3PNTRS.LOC THEN
                     FF3PNTRS.LOC = ABGN
                     CALL FF3PutRec
                  END IF
       CASE "F3": CBGN = FF3PNTRS.COM
                  CALL ViewCOMMENTS(CBGN, ORTYPE, ORNO)
                  CALL WinSWITCH(2)
                  CALL FF3GetRec(ORNO)
                  IF CBGN <> FF3PNTRS.COM THEN
                     FF3PNTRS.COM = CBGN
                     CALL FF3PutRec
                  END IF
    END SELECT
    GOTO 6006

6010 ' Update Miscellaneous Record
    CALL PutMSG("")
    CALL OptDISPLAY(FMT909)
    CALL GetKEY
    SELECT CASE A
       CASE 27: GOTO 6006
       CASE 59: FUN$ = "F1"
       CASE 60: FUN$ = "F2"
       CASE 61: FUN$ = "F3"
       CASE ELSE: CALL ErrBEEP(0): GOTO 6010
    END SELECT
    CALL OptHILITE(FUN$, "", 6)
    SELECT CASE FUN$
       CASE "F1": GOTO 6015   ' Change
       CASE "F2": GOTO 6020   ' Add
       CASE "F3": GOTO 6025   ' Delete
    END SELECT

6015 ' Modify Current Miscellaneous Record
   CALL PutMSG("XXXXM012")
   CURFMT = OTHFMT: FMT = FMTTAB(CURFMT)
   WIN.CLIN = 0: CALL WinUPDATE
   IF (A = 27) OR (A = 60) THEN CALL FMTHREST: CALL WinFORMAT(0): GOTO 6010
   IF A <> 59 THEN CALL ErrBEEP(0): GOTO 6015

6016 ' Write Miscellaneous Record to FILE
   CALL PutMSG("")
   GOSUB 6004
   CALL FF3PutRec
   IF (OBGN = 0) OR (FF3PNTRS.D1MD <> OD1MD) OR (FF3PNTRS.D1Y <> OD1Y) THEN
      REDIM OTHCHN(OTHHIX): OTHIX = 0
      ELSE
      GOTO 6005
   END IF
   IF OBGN = 0 THEN
      OBGN = ORNO:  OTHCHN(0) = OBGN
      ELSE
      OD1MD = FF3PNTRS.D1MD: OD1Y = FF3PNTRS.D1Y
      GOSUB 6017
      IF OBGN = 0 THEN
         OBGN = ORNO: OTHCHN(0) = OBGN
         ELSE
         X = OBGN: Y = 0
         WHILE X <> 0
            CALL FF3GetRec(X)
            IF OKAY THEN
               IF (FF3PNTRS.D1Y > OD1Y) OR ((FF3PNTRS.D1Y = OD1Y) AND (FF3PNTRS.D1MD > OD1MD)) THEN
                  OTHCHN(OTHIX) = X
                  OTHIX = OTHIX + 1
                  IF NOT (OTHIX < OTHHIX) THEN
                     OTHHIX = OTHHIX + 5
                     REDIM PRESERVE OTHCHN(OTHHIX)
                  END IF
                  Y = X: X = FF3PNTRS.NXT
                  ELSE
                  OTHCHN(OTHIX + 1) = X: X = 0
               END IF
               ELSE
               X = 0
            END IF
         WEND
         IF Y = 0 THEN
            OBGN = ORNO
            ELSE
            CALL FF3GetRec(Y)
            IF OKAY THEN
               FF3PNTRS.NXT = ORNO
               CALL FF3PutRec
            END IF
         END IF
         CALL FF3GetRec(ORNO)
         IF OKAY THEN
            IF FF3PNTRS.NXT <> OTHCHN(OTHIX + 1) THEN
               FF3PNTRS.NXT = OTHCHN(OTHIX + 1)
               CALL FF3PutRec
            OTHCHN(OTHIX) = ORNO
            END IF
         END IF
      END IF
   END IF
   GOSUB 6002
   GOTO 6006

6017 ' Remove OTH Record from Chain
   CALL PutMSG("")
   Y = FF3PNTRS.NXT
   IF OBGN = ORNO THEN
      OBGN = Y
      ELSE
      X = OBGN
      WHILE (X <> 0)
         CALL FF3GetRec(X)
         IF OKAY THEN
            IF FF3PNTRS.NXT = ORNO THEN
               FF3PNTRS.NXT = Y
               CALL FF3PutRec
               X = 0
               ELSE
               X = FF3PNTRS.NXT
            END IF
            ELSE
            X = 0
         END IF
      WEND
   END IF
   CALL FF3GetRec(ORNO)
   IF OKAY THEN
      FF3PNTRS.NXT = 0
      CALL FF3PutRec
   END IF
   RETURN

6020 ' Add New Record
   CALL PutMSG("XXXXM012")
   FF3PFX.RTYPE = CHR$(ORTYPE)
   CALL FF3GetRec(0): IF NOT OKAY THEN GOTO 6029
   CURFMT = OTHFMT: FMT = FMTTAB(CURFMT)
   CALL FMTDREST
   CALL WinFORMAT(0): WIN.CLIN = 0
6021 CALL WinUPDATE
   IF (A = 27) OR (A = 60) THEN
      CALL FF3DelRec(FF3PFX.RNO)
      IF ORNO = 0 THEN
         GOTO 6029
         ELSE
         CALL FMTHREST: CALL WinFORMAT(OTHFMT)
         CALL FF3GetRec(ORNO)
         GOTO 6010
      END IF
   END IF
   IF A <> 59 THEN CALL ErrBEEP(0): GOTO 6021
   ORNO = FF3PFX.RNO
   FF3PNTRS.SRTYPE = CHR$(1): FF3PNTRS.SRNO = BRID
   GOTO 6016

6025 ' Delete Miscellaneous Record
   CALL PutMSG("UPDTM007")
   IF A$ <> SNGLKEY$(1) THEN GOTO 6010
   MADR = FF3PNTRS.LOC
   IF MADR <> 0 THEN
      CALL PutMSG("UPDTM008")
      IF A$ <> SNGLKEY$(1) THEN GOTO 6010
   END IF
   MCOM = FF3PNTRS.COM
   IF MCOM <> 0 THEN
      CALL PutMSG("UPDTM009")
      IF A$ <> SNGLKEY$(1) THEN GOTO 6010
   END IF
   X = MADR
   WHILE X <> 0
      CALL FF2GetRec(X)
      IF NOT OKAY THEN
         X = 0
         ELSE
         Y = FF2PFX.COM
         WHILE Y <> 0
            CALL FF3GetRec(Y)
            IF NOT OKAY THEN
               Y = 0
               ELSE
               NY = FF3COMMENT.FPTR
               CALL FF3DelRec(Y)
               IF NOT OKAY THEN NY = 0
               Y = NY
            END IF
         WEND
         NX = FF2PFX.NXT
         CALL FF2DelRec(X)
         IF NOT OKAY THEN NX = 0
         X = NX
      END IF
   WEND
   X = MCOM
   WHILE X <> 0
      CALL FF3GetRec(X)
      IF NOT OKAY THEN
         X = 0
         ELSE
         NX = FF3COMMENT.FPTR
         CALL FF3DelRec(X)
         IF NOT OKAY THEN NX = 0
         X = NX
      END IF
   WEND
   CALL FF3GetRec(ORNO)
   GOSUB 6017
   CALL FF3DelRec(ORNO)
   OTHCHN(OTHIX) = OTHCHN(OTHIX + 1)
   ORNO = OTHCHN(OTHIX)
   IF (ORNO = 0) AND (OTHIX > 0) THEN OTHIX = OTHIX - 1: ORNO = OTHCHN(OTHIX)
   IF ORNO <> 0 THEN GOTO 6005

6029 ' Return from Miscellaneous Record
   CALL PutMSG(""): UPDT = 0
   SELECT CASE ORTYPE
      CASE 7: FF1REC.EDUC = OBGN
      CASE 8: FF1REC.WORK = OBGN
      CASE 9: FF1REC.MIL = OBGN
      CASE 10: FF1REC.HLTH = OBGN
   END SELECT
   IF XCHANGED THEN CALL FF1PutREC
   CALL WinCLR
   CALL WinSWITCH(3)

END SUB

SUB ViewPARENTS (RID)
SHARED FUN$, RGL
SHARED FMT007, FMT906
   RGL = RGL - 1
   CALL FF1GetRec(RID)
   NameREC = FF1REC
   CALL WinSWITCH(2)
   IF WIN.WFMT <> FMT007 THEN
      CALL WinCLR
      CALL WinDisplayLABELS(FMT007)
   END IF
   CURFMT = WIN.WFMT: FMT = FMTTAB(CURFMT)
   CALL FMTDREST
   CALL FmtFindFLD("RID"): LSET UTXT$(FFLD.UTXT) = STR$(RID)
   CALL FmtFindFLD("RGL"): RSET UTXT$(FFLD.UTXT) = LTRIM$(STR$(RGL))
   CALL FmtFindFLD("SFID"): LSET UTXT$(FFLD.UTXT) = STR$(NameREC.FID)
   IF NameREC.FID > 0 THEN
      CALL FF1GetRec(NameREC.FID): FOLDCH = FF1REC.OLDCH
      CALL FmtFindFLD("FGVN"): LSET UTXT$(FFLD.UTXT) = FF1REC.GIVEN
      CALL FmtFindFLD("FSNM"): LSET UTXT$(FFLD.UTXT) = FF1REC.SURNM
      CALL FmtFindFLD("FBDT"): LSET UTXT$(FFLD.UTXT) = XSCRNDate(FF1REC.BMD, FF1REC.BY)
      CALL FmtFindFLD("FBDS"): LSET UTXT$(FFLD.UTXT) = FDSTAT$("BD")
      CALL FmtFindFLD("FDDT"): LSET UTXT$(FFLD.UTXT) = XSCRNDate(FF1REC.DMD, FF1REC.DY)
      CALL FmtFindFLD("FDDS"): LSET UTXT$(FFLD.UTXT) = FDSTAT$("DD")
      CALL FmtFindFLD("FAGE"): LSET UTXT$(FFLD.UTXT) = DateDif$(FF1REC.BMD, FF1REC.BY, FF1REC.DMD, FF1REC.DY)
   END IF
   CALL FmtFindFLD("SMID"): LSET UTXT$(FFLD.UTXT) = STR$(NameREC.MID)
   IF NameREC.MID > 0 THEN
      CALL FF1GetRec(NameREC.MID): MOLDCH = FF1REC.OLDCH
      CALL FmtFindFLD("MGVN"): LSET UTXT$(FFLD.UTXT) = FF1REC.GIVEN
      CALL FmtFindFLD("MSNM"): LSET UTXT$(FFLD.UTXT) = FF1REC.SURNM
      CALL FmtFindFLD("MBDT"): LSET UTXT$(FFLD.UTXT) = XSCRNDate(FF1REC.BMD, FF1REC.BY)
      CALL FmtFindFLD("MBDS"): LSET UTXT$(FFLD.UTXT) = FDSTAT$("BD")
      CALL FmtFindFLD("MDDT"): LSET UTXT$(FFLD.UTXT) = XSCRNDate(FF1REC.DMD, FF1REC.DY)
      CALL FmtFindFLD("MDDS"): LSET UTXT$(FFLD.UTXT) = FDSTAT$("DD")
      CALL FmtFindFLD("MAGE"): LSET UTXT$(FFLD.UTXT) = DateDif$(FF1REC.BMD, FF1REC.BY, FF1REC.DMD, FF1REC.DY)
   END IF
701 : CALL WinDisplayDATA(0)
   CALL OptDISPLAY(FMT906)
   PCHILD = ((NameREC.FID = 0) OR (FOLDCH <> 0)) AND ((NameREC.MID = 0) OR (MOLDCH <> 0))
   IF PCHILD THEN CALL OptHILITE("F3", "", 7)
   CALL GetKEY: FUN$ = ""
   SELECT CASE A
      CASE 27: RGL = RGL + 1: GOTO 709
      CASE 59: IF NameREC.FID <> 0 THEN FUN$ = "F1"
      CASE 60: IF NameREC.MID <> 0 THEN FUN$ = "F2"
      CASE 61: IF PCHILD THEN FUN$ = "F3"
   END SELECT
   IF FUN$ = "" THEN CALL ErrBEEP(0): GOTO 701
   CALL OptHILITE(FUN$, "", 6)
   SELECT CASE FUN$
      CASE "F1": RID = NameREC.FID
      CASE "F2": RID = NameREC.MID
      CASE "F3": FUN$ = "CHLD"
   END SELECT
709 ' Return
   FF1REC = NameREC
END SUB

