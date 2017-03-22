DECLARE SUB ULFLDS (A%)
DECLARE SUB BldMLIST ()
DECLARE SUB GetNextMARRIAGE ()
DECLARE SUB BLDFGRPDef ()
DECLARE SUB FGRPReport ()
DECLARE SUB SelectID ()
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
DECLARE FUNCTION XRPTDate$ (MD%, Y%, O%)
DECLARE FUNCTION XSCRNDate$ (MD%, Y%)

TYPE NDATA      ' 174 Bytes
     RID      AS STRING * 5
     PID      AS STRING * 5
     NAME     AS STRING * 53
     BDATE    AS STRING * 12
     BPLACE   AS STRING * 41
     DDATE    AS STRING * 12
     DPLACE   AS STRING * 41
     SEX      AS STRING * 1
     NUMCH    AS STRING * 3
END TYPE

TYPE PDATA       ' 79 Bytes - Father, Mother
    NAME      AS STRING * 52
    BDATE     AS STRING * 12
    DDATE     AS STRING * 12
    AGE       AS STRING * 3
END TYPE

TYPE MDATA       ' 164 Bytes
     MDATE    AS STRING * 11
     MPLACE   AS STRING * 41
     DVDATE   AS STRING * 11
     DVPLACE  AS STRING * 41
     ANNIV    AS STRING * 3
     SPID     AS STRING * 5
     SPNAME   AS STRING * 52
END TYPE

TYPE EDATA       ' 79 Bytes
     CODE     AS STRING * 8
     DESC     AS STRING * 15
     DATE     AS STRING * 11
     DSTAT    AS STRING * 1
     PLACE    AS STRING * 41
     IMPORT   AS STRING * 3
END TYPE

1  REM $INCLUDE: 'FHSCOMON.BAS'
   PN$ = " FAMILY HISTORY SYSTEM - FHSGRPR - Family Group Report Program"
   CY$ = " (C) Copyright 1986,1987,1988,1989,1990,1991,1992,1993,1994,1995 by Phillip E. Brown"
   IF CONFIG.BP = 0 THEN RUN "FHSINIT"
   REDIM STYPE$(5)
   MAXMarry = 10: REDIM MLIST(MAXMarry)
   STYPE$(1) = "Single ID": STYPE$(2) = "ID List"
   STYPE$(3) = "REL Work File": STYPE$(4) = "SELECT File"
   STYPE$(5) = "Full File"
   CALL TableLOAD("WORKTYPE", T, 39)
   IF OKAY THEN
      X = TABLE(T).FT
      WHILE X > 0
        Z = VAL(MID$(TABDATA$(X), 1, 2))
        IF (3 < Z) AND (Z < 9) THEN
           STYPE$(Z - 3) = RTRIM$(MID$(TABDATA$(X), 15, 255))
        END IF
        X = TABDATACHN(X).FWD
      WEND
      CALL TableFREE("WORKTYPE")
   END IF
   NAME$ = SPACE$(52): IA$ = "": MA$ = "": READ$ = "READQ"
   DEF FNX$ (A) = LTRIM$(STR$(A))
   REDIM RELATE(1), RELPTR&(1)
   REDIM IDLIST(1), NDXID(1)
   WTYPE = MO: BRID = 0
   CALL WinINIT("LR1")
   CALL WinPREP(1): PGMRTRN$ = "FHSMENU"
   FUN = 0: CALL RptLOAD("FXGR"): CALL RptLOAD("FFGR"): CALL RptLOAD("BLGR")

40 ' Format Display
   CALL WinCLR
   CALL FmtFIND("GRPRS001"): FMT001 = CURFMT: WIN.OFMT = FMT001
   CALL FmtFIND("GRPRS002"): FMT002 = CURFMT
   CALL WinDisplayLABELS(FMT001)

45 ' Refresh Option DISPLAY
   IF MO = 1 THEN RRPT$ = "FXGR" ELSE RRPT$ = "FFGR"
   CALL WinPREP(2): CALL WinCLR: OptTABLE.CL = 0
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
    CALL FAMOpen(READ$): CALL FamCLOSE: READ$ = "READQ"
    CURFMT = FMT001: FMT = FMTTAB(CURFMT)
    CALL FmtFindFLD("NMCT")
    LSET UTXT$(FFLD.UTXT) = STR$(FF1HDR.MAXID)
    CALL WinDisplayDATA(FMT001)
    CALL WinDisplayDATA(FMT002)

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
       CASE "F2": GOTO 400
       CASE "F3": CALL SelectID
                  GOTO 100
       CASE "F4": GOTO 800
       CASE "F6": IF PGMRTRN$ = "FHSMENU" THEN CALL FGRPReport
                  IF PGMRTRN$ = "FHSRPTS" THEN 250
       CASE "F7": MO = MO + 1: IF MO > 2 THEN MO = 1
                  GOTO 45
       CASE "F9": GOTO 250
    END SELECT
    IF ENV.SCRNCLR THEN GOTO 40 ELSE GOTO 200

250 ' Return to FAMMENU.BAS
    CALL FamCLOSE
    REDIM RFLTH(1), RFLBL$(1)
    CALL PgmPREP(PGMRTRN$)
    IF (NOT OKAY) OR (A = 27) THEN LSET PGMRTRN$ = "FHSMENU": GOTO 40
    CALL RptFREE("FXGR"): CALL RptFREE("FFGR"): CALL RptFREE("BLGR")
    CHAIN PGMRTRN$

300 ' Change FILE Parameters
    IF A = 94 THEN LSET PGMRTRN$ = "FHSFILE": GOTO 250
    CALL WinSWITCH(2)
    XSETUP = ENV.FSETUP
    CALL FDFSelect
    IF ENV.FSETUP <> XSETUP THEN
       RID = 0: LSET NAME$ = "": SCNT = 0
       FMT = FMTTAB(FMT002): CALL FmtDREST
       CALL WinFORMAT(0): READ$ = "READ"
    END IF
    GOTO 45

400 ' Change PRINTER Parameters
    IF A = 95 THEN LSET PGMRTRN$ = "FHSPRTC": GOTO 250
    CALL WinSWITCH(2)
    CALL PDFSelect
    GOTO 45

800 ' Process Request to Update Options
    IF RRPT$ = "FXGR" THEN
       CALL PutMSG("GRPRM002")
       IF A = 27 THEN 200
       IF A$ = "1" THEN RRPT$ = "BLGR"
    END IF
    CALL WinSWITCH(2)
    IF RRPT$ = "BLGR" THEN
       CALL WinCLR: CALL RptOptDISPLAY(RRPT$)
    END IF
    CALL RptOptUPDATE(RRPT$)
    GOTO 45

REM $DYNAMIC
SUB BLDFGRPDef
SHARED IA$, MA$, BLNKFORM
      PAGE.DFCNT = 67
      REDIM RFLTH(PAGE.DFCNT), RFLBL$(PAGE.DFCNT)

58500 ' Select Report Fields - Based Upon Report Options
      IA$ = SPACE$(12): MA$ = SPACE$(24): X = 0: MX$ = "  "
      XI$ = "PRSP": XM$ = "NMSP": GOSUB 58505
      XI$ = "PRCH": XM$ = "NMCH": GOSUB 58505
      XI$ = "PRAD": XM$ = "NMAD": GOSUB 58505
      XI$ = "PRCM": XM$ = "NMCM": GOSUB 58505
      XI$ = "PRME": XM$ = "NMME": GOSUB 58505
      XI$ = "PRED": XM$ = "NMED": GOSUB 58505
      XI$ = "PRWK": XM$ = "NMWK": GOSUB 58505
      XI$ = "PRML": XM$ = "NMML": GOSUB 58505
      XI$ = "PRXA": XM$ = "NMXA": GOSUB 58505
      XI$ = "PRXC": XM$ = "NMXC": GOSUB 58505
      XI$ = "PREV": XM$ = "NMEV": GOSUB 58505
      XI$ = "CHCM": XM$ = "    ": GOSUB 58505
      GOTO 58510

58505 ' Process Information Option
      X = X + 1: CALL RptOptGET(XI$, FX)
      IF FOUND THEN MID$(IA$, X, 1) = LTRIM$(STR$(FX))
      IF NOT RptOPTION.FREEFORM THEN
         CALL RptOptGET(XM$, FX)
         IF FOUND THEN
            RSET MX$ = LTRIM$(STR$(ABS(FX)))
            IF FX < 0 THEN MID$(MX$, 1) = "+"
            ELSE
            RSET MX$ = " 0"
         END IF
         ELSE
         RSET MX$ = " 0"
      END IF
      MID$(MA$, 2 * X - 1, 2) = MX$
      RETURN

58510 ' Create Report Definition Tables
      IF BLNKFORM THEN RRPT$ = "FXGR": CALL RptLOAD(RRPT$)
      X = 1: RFNAME$ = "06010": GOSUB 58515       'Comments
      X = 2: RFNAME$ = "01010": GOSUB 58515       'RID
      X = 3: RFNAME$ = "01013": GOSUB 58515       'Parent ID
      X = 4: RFNAME$ = "02010": GOSUB 58515       'Spouse's ID
      X = 5: RFNAME$ = "01020": GOSUB 58515       'Fullname
      X = 6: RFNAME$ = "01023": GOSUB 58515       'Father's Name
      X = 7: RFNAME$ = "01024": GOSUB 58515       'Mother's Name
      X = 8: RFNAME$ = "01030": GOSUB 58515       'Birth Date
      X = 9: RFNAME$ = "01032": GOSUB 58515       'Birth Place
      X = 10: RFNAME$ = "01040": GOSUB 58515      'Death Date
      X = 11: RFNAME$ = "01042": GOSUB 58515      'Death Place
      X = 12: RFNAME$ = "01016": GOSUB 58515      'Age
      X = 13: RFNAME$ = "02011": GOSUB 58515      'Spouse's Name
      X = 14: RFNAME$ = "02020": GOSUB 58515      'Marriage Date
      X = 15: RFNAME$ = "02021": GOSUB 58515      'Marriage Place
      X = 16: RFNAME$ = "02030": GOSUB 58515      'Divorce Date
      X = 17: RFNAME$ = "02031": GOSUB 58515      'Divorce Place
      X = 18: RFNAME$ = "02015": GOSUB 58515      'Anniversary (for...years)
      X = 19: RFLBL$(19) = MID$(RFLD$(FX), ALTH + LLTH + 1, HLTH)
      X = 20: RFNAME$ = "01015": GOSUB 58515      'Child's Gender
      X = 21: RFNAME$ = "01026": GOSUB 58515      'Child's Name
      X = 22: RFNAME$ = "01034": GOSUB 58515      'Birth Date/Place
      X = 23: RFNAME$ = "01027": GOSUB 58515      'Number Children
              IF FOUND THEN                       ' one child
                 RFLBL$(67) = MID$(RFLD$(FX), ALTH + 1, LLTH)
                 RFLTH(67) = LLTH
              END IF
      X = 24: RFNAME$ = "01044": GOSUB 58515      'Death Date/Place
      X = 25: RFNAME$ = "01028": GOSUB 58515      'Married Date/Spouse
      X = 26: RFNAME$ = "01030": GOSUB 58515      'Birth Date Abbrev.
      X = 27: RFNAME$ = "01032": GOSUB 58515      'Birth Place Abbrev.
      X = 28: RFNAME$ = "01040": GOSUB 58515      'Death Date Abbrev.
      X = 29: RFNAME$ = "01042": GOSUB 58515      'Death Place Abbrev.
      X = 30: RFNAME$ = "02020": GOSUB 58515      'Marriage Date Abbrev.
      X = 31: RFNAME$ = "03001": GOSUB 58515      'Address Section Label
      X = 32: RFNAME$ = "03010": GOSUB 58515      'Address Begin Date
      X = 33: RFNAME$ = "03020": GOSUB 58515      'Address End Date
      X = 34: RFNAME$ = "03030": GOSUB 58515      'Phone
      X = 35: RFNAME$ = "03031": GOSUB 58515      'Address Lines
      X = 36: RFNAME$ = "03032": GOSUB 58515      'City
      X = 37: RFNAME$ = "03033": GOSUB 58515      'State
      X = 38: RFNAME$ = "03034": GOSUB 58515      'Zip Code
      X = 39: RFNAME$ = "03035": GOSUB 58515      'Province
      X = 40: RFNAME$ = "03037": GOSUB 58515      'Postal Code
      X = 41: RFNAME$ = "03036": GOSUB 58515      'Country
      X = 42: RFNAME$ = "04001": GOSUB 58515      'Event Information
      X = 43: RFNAME$ = "04010": GOSUB 58515      'Event Code
      X = 44: RFNAME$ = "04011": GOSUB 58515      'Event Date
      X = 45: RFNAME$ = "04012": GOSUB 58515      'Event Place
      X = 46: RFNAME$ = "04013": GOSUB 58515      'Event Import Level
      X = 47: RFNAME$ = "05010": GOSUB 58515      'Period Begin Date
      X = 48: RFNAME$ = "05011": GOSUB 58515      'Period End Date
      X = 49: RFNAME$ = "05030": GOSUB 58515      'Education Information
      X = 50: RFNAME$ = "05031": GOSUB 58515      'Level
      X = 66: RFLBL$(66) = MID$(RFLD$(FX), ALTH + LLTH + 1, HLTH)
              RFLTH(66) = HLTH
      X = 51: RFNAME$ = "05032": GOSUB 58515      'Subject1
      X = 52: RFNAME$ = "05033": GOSUB 58515      'Subject2
      X = 53: RFNAME$ = "05034": GOSUB 58515      'Degree
      X = 54: RFNAME$ = "05040": GOSUB 58515      'Medical Information
      X = 55: RFNAME$ = "05041": GOSUB 58515      'Diagnosis Date
      X = 56: RFNAME$ = "05042": GOSUB 58515      'Illness
      X = 57: RFNAME$ = "05043": GOSUB 58515      'Latest Status
      X = 58: RFNAME$ = "05044": GOSUB 58515      'Status Date
      X = 59: RFNAME$ = "05050": GOSUB 58515      'Military Information
      X = 60: RFNAME$ = "05051": GOSUB 58515      'Rank
      X = 61: RFNAME$ = "05052": GOSUB 58515      'Status
      X = 62: RFNAME$ = "05060": GOSUB 58515      'Occupation Information
      X = 63: RFNAME$ = "05061": GOSUB 58515      'Type Work
      X = 64: RFNAME$ = "07004": GOSUB 58515      'Relationship
      X = 65: RFNAME$ = "07010": GOSUB 58515      'Lineage
      GOTO 58530

58515 ' Process Report Field
      CALL GetRFLD(RFNAME$, FX)
      IF FOUND THEN
         RDFFLD = RFLDS(FX)
         ALTH = ASC(RDFFLD.ALTH): LLTH = ASC(RDFFLD.LLTH): HLTH = ASC(RDFFLD.HLTH)
         IF (X < 20) OR (X > 30) THEN
            RFLBL$(X) = MID$(RFLD$(FX), ALTH + 1, LLTH)
            ELSE
            IF (X < 26) THEN
               RFLBL$(X) = MID$(RFLD$(FX), ALTH + LLTH + 1, HLTH)
               ELSE
               RFLBL$(X) = MID$(RFLD$(FX), 1, ALTH)
            END IF
         END IF
      END IF
      RFLTH(X) = LEN(RFLBL$(X))
      RETURN

58530 ' Check Report Width
      OKAY = -1: RWIDTH = RDFReport.PWIDTH
      IF RWIDTH = 0 THEN RWIDTH = 125
      IF PAGE.BWIDTH = 0 THEN
         BWIDTH = PDFFORMS.FWIDTH - PDFFORMS.IMARGIN - PDFFORMS.OMARGIN
         ELSE
         BWIDTH = PAGE.BWIDTH
      END IF
      IF RptOPTION.FREEFORM THEN 58560
      PWIDTH = BWIDTH + PDFFORMS.IMARGIN + PDFFORMS.OMARGIN
      RPWIDTH = RWIDTH + PDFFORMS.IMARGIN + PDFFORMS.OMARGIN
      IF PWIDTH < RPWIDTH THEN
         SYSVAR.NVAR1 = PWIDTH: SYSVAR.NVAR2 = RPWIDTH
         CALL PutMSG("XXXXM009")
         IF A$ = SNGLKEY$(1) THEN BWIDTH = RWIDTH ELSE OKAY = 0
      END IF

58560 ' Return to Calling Program
      PAGE.BWIDTH = BWIDTH
END SUB

REM $STATIC
SUB BldMLIST
SHARED SPRESEQ, SC, SPMCNT, MAXMarry, MLIST()
   SPMCNT = 0: SC = FF1REC.SPOUSE
   IF SPRESEQ THEN
      WHILE SC <> 0
         SPMCNT = SPMCNT + 1
         IF SPMCNT > MAXMarry THEN
            MAXMarry = MAXMarry + 5
            REDIM PRESERVE MLIST(MAXMarry)
         END IF
         MLIST(SPMCNT) = SC
         CALL FF3GetRec(SC)
         IF FF1REC.RID = FF3SPOUSE.SP1ID THEN
            SC = FF3SPOUSE.SP1NXT
            ELSE
            IF FF1REC.RID = FF3SPOUSE.SP2ID THEN
               SC = FF3SPOUSE.SP2NXT
               ELSE
               SC = 0
            END IF
         END IF
      WEND
      SC = MLIST(SPMCNT)
   END IF

END SUB

REM $DYNAMIC
SUB FGRPReport
SHARED BLNKFORM, IA$, MA$, NDXID(), RELATE(), RELPTR&()
SHARED LOID, HIID, RID, STYPE, SCNT, IDLIST(), RRPT$
SHARED SPRESEQ, MLIST(), SPMCNT, SC
SHARED NameDATA AS NDATA, FatherDATA AS PDATA, MotherDATA AS PDATA
SHARED MarriageDATA AS MDATA, EventDATA AS EDATA

    RptOPTION.FREEFORM = (RRPT$ = "FFGR")
    IF NOT RptOPTION.FREEFORM THEN
       CALL PutMSG("GRPRM002"): IF ASC(A$) = 27 THEN 1090
       BLNKFORM = (A$ = "1"): IF BLNKFORM THEN RRPT$ = "BLGR"
    END IF

1000 'Process PRINT Requests
    CALL RptOPEN(RRPT$): IF NOT OKAY THEN GOTO 1090
    RptOPTION.FREEFORM = (RRPT$ = "FFGR")
    IF RptOPTION.FREEFORM THEN
       FC = 32: ULALL = 0
       ELSE
       CALL RptOptGET("ULBF", X): ULALL = (X = 4)
       IF X <> 2 THEN FC = 95 ELSE FC = 32
    END IF
    CALL RptOptGET("BLKL", X): BLNKBTWN = (X = 1): CHILDSEP = BLNKBTWN
    CALL RptOptGET("EURA", X): EuroADR = (X = 1)
    CALL BLDFGRPDef
    REDIM OA(4)
    RID$ = SPACE$(5)
    XRID$ = "(" + RFLBL$(2) + RID$ + ")"
    XOPID$ = RFLBL$(3) + RID$ + ")"
    XSPID$ = "(" + RFLBL$(4) + RID$ + ")"
    CDATE$ = SPACE$(11)
    DIM BNDATA AS NDATA, BPDATA AS PDATA, BMDATA AS MDATA, BEDATA AS EDATA
    DIM NameDATA AS NDATA, FatherDATA AS PDATA, MotherDATA AS PDATA
    DIM MarriageDATA AS MDATA, EventDATA AS EDATA
    BNDATA.RID = STRING$(LEN(BNDATA.RID), FC)
    BNDATA.PID = STRING$(LEN(BNDATA.PID), FC)
    BNDATA.NAME = STRING$(LEN(BNDATA.NAME), FC)
    BNDATA.BDATE = STRING$(LEN(BNDATA.BDATE), FC)
    BNDATA.BPLACE = STRING$(LEN(BNDATA.BPLACE), FC)
    BNDATA.DDATE = STRING$(LEN(BNDATA.DDATE), FC)
    BNDATA.DPLACE = STRING$(LEN(BNDATA.DPLACE), FC)
    BNDATA.SEX = STRING$(LEN(BNDATA.SEX), FC)
    BNDATA.NUMCH = STRING$(LEN(BNDATA.NUMCH), FC)
    BPDATA.NAME = STRING$(LEN(BPDATA.NAME), FC)
    BPDATA.BDATE = STRING$(LEN(BPDATA.BDATE), FC)
    BPDATA.DDATE = STRING$(LEN(BPDATA.DDATE), FC)
    BPDATA.AGE = STRING$(LEN(BPDATA.AGE), FC)
    BMDATA.MDATE = STRING$(LEN(BMDATA.MDATE), FC)
    BMDATA.MPLACE = STRING$(LEN(BMDATA.MPLACE), FC)
    BMDATA.DVDATE = STRING$(LEN(BMDATA.DVDATE), FC)
    BMDATA.DVPLACE = STRING$(LEN(BMDATA.DVPLACE), FC)
    BMDATA.ANNIV = STRING$(LEN(BMDATA.ANNIV), FC)
    BMDATA.SPID = STRING$(LEN(BMDATA.SPID), FC)
    BMDATA.SPNAME = STRING$(LEN(BMDATA.SPNAME), FC)
    BEDATA.CODE = STRING$(LEN(BEDATA.CODE), FC)
    BEDATA.DESC = STRING$(LEN(BEDATA.DESC), FC)
    BEDATA.DATE = STRING$(LEN(BEDATA.DATE), FC)
    BEDATA.DSTAT = STRING$(LEN(BEDATA.DSTAT), FC)
    BEDATA.PLACE = STRING$(LEN(BEDATA.PLACE), FC)
    BEDATA.IMPORT = STRING$(LEN(BEDATA.IMPORT), FC)

    'NameDATA$ = SPACE$(136)
    '  1   RID       5
    '  6   PID       5
    ' 11   NAME     53
    ' 63   BDATE    12
    ' 75   BPLACE   41
    ' 97   DDATE    12
    '109   DPLACE   41
    '131   SEX       1
    '132   NUMCH     2
    'FatherDATA$ = SPACE$(79): MotherDATA$ = FatherDATA$
    '  1   NAME     52
    ' 53   BDATE    12
    ' 65   DDATE    12
    ' 77   AGE       3
    'MarriageDATA$ = SPACE$(126)
    '  1   MDATE    11
    ' 12   MPLACE   41
    ' 34   DVDATE   11
    ' 45   DVPLACE  41
    ' 67   ANNIV     3
    ' 70   SPID      5
    ' 75   SPNAME   52
    'EventDATA$ = SPACE$(45)
    '  1   CODE      8
    '  9   DATE     11
    ' 20   DSTAT     1
    ' 21   PLACE    22
    ' 43   IMPORT    3
    ADRDATA$ = SPACE$(157): MAXADR = 30
    OTHDATA$ = SPACE$(52)

    ' Build Separator Lines from Report Variables
    IF RptOPTION.FREEFORM THEN SPW = PAGE.BWIDTH ELSE SPW = 125
    LINESEP1$ = SPACE$(SPW)
    CALL GetRVAR("LINESEP1", X)
    IF FOUND THEN LSET LINESEP1$ = RVAR$(X): CALL HTFBuild(LINESEP1$, PAGE.BWIDTH)
    LINESEP2$ = SPACE$(SPW)
    CALL GetRVAR("LINESEP2", X)
    IF FOUND THEN LSET LINESEP2$ = RVAR$(X): CALL HTFBuild(LINESEP2$, PAGE.BWIDTH)
    LINESEP3$ = SPACE$(SPW)
    CALL GetRVAR("LINESEP3", X)
    IF FOUND THEN LSET LINESEP3$ = RVAR$(X): CALL HTFBuild(LINESEP3$, PAGE.BWIDTH)
    ' Build Child Information Header
    IF RptOPTION.FREEFORM THEN
       ChildHDR$ = RFLBL$(21)
       ELSE
       LSET PLINE$ = " "
       MID$(PLINE$, 4, 255) = RFLBL$(20)
       MID$(PLINE$, 22 - LEN(RFLBL$(21)) \ 2, 255) = RFLBL$(21)
       MID$(PLINE$, 42, 255) = RFLBL$(22)
       MID$(PLINE$, 61, 3) = RFLBL$(23)
       MID$(PLINE$, 66, 255) = RFLBL$(24)
       MID$(PLINE$, 90, 35) = RFLBL$(25)
       ChildHDR$ = RTRIM$(PLINE$)
    END IF
    IF BLNKFORM THEN GOTO PRTBlankForm
    ' Get Misc Options
    PRTHDRID = 0
    IF RptOPTION.PRID THEN
       IF RptOPTION.PRID = 2 THEN RptOPTION.PRID = 0
       CALL GetRFLD("01010", X)
       IF FOUND THEN
          PRTHDRID = -1
          RIDLbl$ = MID$(RFLD$(X), ASC(RDFFLD.ALTH) + 1, ASC(RDFFLD.LLTH))
       END IF
    END IF
    CALL RptOptGET("PLIN", X): SHOWLIN = (X = 1)
    CALL RptOptGET("BYPS", X)
    BYPASS = ((X = 10) OR (X = 11)): BYPASS2 = ((X = 10) OR (X = 12))
    CALL RptOptGET("SP12", X): SPRESEQ = (X = 1)
    CALL GetRVAR("UNKNOWN", X): IF FOUND THEN Unknown$ = RVAR$(X)
    CALL FmtFIND("GRPRS015"): FMT015 = CURFMT
    CALL FmtFIND("GRPRS016"): FMT016 = CURFMT
    GOTO PRTFGrpReport

PRTBlankForm: ' Print Blank Family Group Report
    CLSCRN = 0: BF = -1: BFH = -1: FORMFD = -1: NBL = -1:
    A = 0: FC = 95: MLOC = 0: RptOPTION.BLDREF = 0
    IF RptOPTION.PRID = 2 THEN RptOPTION.PRID = 0
    PSTAT = 0: CSTAT = 0: STAT1 = 0
    CALL PutMSG("GRPRM003"): IF ASC(A$) = 27 THEN 1090
    COPIES = VAL(A$): IF COPIES = 0 THEN 1090
    IF ((MID$(IA$, 3, 1) = "1") AND (VAL(MID$(MA$, 5, 2)) > 0)) OR ((MID$(IA$, 9, 1) = "1") AND (VAL(MID$(MA$, 17, 2)) > 0)) THEN
       CALL PutMSG("GRPRM004"): IF ASC(A$) = 27 THEN 1090
       AFMT = VAL(A$) - 1: IF AFMT < 0 OR AFMT > 1 THEN AFMT = 0
    END IF
    CALL PrintOPEN
    PRT.SECTION = "F"
    LSET SYSVAR.SUBJECT = STRING$(52, FC)
    LSET CDATE$ = STRING$(11, FC)
    SHOWLIN = 0: HCOPY = COPIES
    ComPTR = 0: COMPTR2 = 0
PBFAgain:
    WHILE COPIES > 0
       PAGE.COUNT = PAGE.BGNPAGE - 1
       GOSUB 600: GOSUB 2300: IF A = 27 THEN CALL RptCLOSE: GOTO 1090
       IF MID$(IA$, 1, 1) = "1" THEN
          RH = VAL(MID$(MA$, 1, 2))
          IF RH <> 0 THEN
             SPCNT = 0
             WHILE SPCNT < RH
                GOSUB 2350: IF A = 27 THEN 1090
                SPCNT = SPCNT + 1
             WEND
          END IF
       END IF
       IF MID$(IA$, 2, 1) = "1" THEN
          RH = VAL(MID$(MA$, 3, 2))
          IF RH > 0 THEN
             CH = 0: GOSUB 2400: IF A = 27 THEN 1090
             NameDATA.NAME = STRING$(52, FC)
             MID$(NameDATA.BDATE, 1, RFLTH(26)) = RFLBL$(26)
             MID$(NameDATA.BPLACE, 1, RFLTH(27)) = RFLBL$(27)
             MID$(NameDATA.DDATE, 1, RFLTH(28)) = RFLBL$(28)
             MID$(NameDATA.DPLACE, 1, RFLTH(29)) = RFLBL$(29)
             WHILE RH <> 0
                GOSUB 2410: IF A = 27 THEN 1090
                RH = RH - 1
             WEND
          END IF
       END IF
       IF MID$(IA$, 3, 1) = "1" THEN
          RH = VAL(MID$(MA$, 5, 2))
          IF RH <> 0 THEN
             LSET ADRDATA$ = STRING$(LEN(ADRDATA$), FC)
             GOSUB PrintRESH: IF A = 27 THEN 1090
             WHILE RH <> 0
                GOSUB PrintRESI: IF A = 27 THEN 1090
                RH = RH - 1
             WEND
          END IF
       END IF
       IF MID$(IA$, 11, 1) = "1" THEN
          RH = VAL(MID$(MA$, 21, 2))
          IF RH <> 0 THEN
             GOSUB PrintEventH: IF A = 27 THEN 1090
             WHILE RH <> 0
                GOSUB PrintEVENT: IF A = 27 THEN 1090
                RH = RH - 1
             WEND
          END IF
       END IF
       LSET OTHDATA$ = STRING$(LEN(OTHDATA$), FC)
       FOR OT = 1 TO 4: OA(OT) = 0: GOSUB 2700: NEXT OT
       CALL PageBREAK
       COPIES = COPIES - 1
    WEND
    CALL RptCLOSE
    IF PRT.AGAIN THEN COPIES = HCOPY: GOTO PBFAgain
1090 CALL FamCLOSE
    EXIT SUB

PRTFGrpReport: ' Print Family Group Reports for Selected ID's
    IF SCNT = 0 THEN CALL PutMSG("GRPRM005"): GOTO 1090
    CALL FmtFIND("GRPRS002"): FMT002 = CURFMT
    CALL FmtFindFLD("ID"): IF FOUND THEN FFRID = FFLD.UTXT
    CALL FmtFindFLD("NAME"): IF FOUND THEN FFNAME = FFLD.UTXT
    CALL FAMOpen("READ"): IF NOT OKAY THEN GOTO 1090
    FORMFD = -1
    A = 0: BF = 0: BFH = 0: INDEXED = 0: RELSEQ = 0
    BD$ = SPACE$(12): BPL$ = SPACE$(41): DD$ = BD$: DPL$ = BPL$: AGE$ = "   "
    IF RptOPTION.FREEFORM THEN DS = 0 ELSE DS = 1
307 IF (SCNT > 1) THEN
       CALL PutMSG("GRPRM006"): IF ASC(A$) = 27 THEN 396
       RELSEQ = (A$ = "2"): INDEXED = (A$ = "3")
       IF RELSEQ THEN SHOWLIN = -1
    END IF
    NDXRECS = FF1HDR.MAXID
    IF INDEXED THEN
       REDIM NDXID(FF1HDR.MAXID)
       CALL INDEXFile("READ")
       IF ASC(A$) = 27 THEN 396 ELSE IF NOT OKAY THEN 307
    END IF
    IF RELSEQ OR SHOWLIN THEN
       IF RELSEQ THEN X = FMT016 ELSE X = FMT015
       CALL RWRKOpen("READ", 0, X): IF NOT OKAY THEN GOTO 396
       IF RELSEQ THEN NDXRECS = 0: REDIM NDXID(FF1HDR.MAXID)
       IF SHOWLIN THEN
          REDIM RELATE(FF1HDR.MAXID), RELPTR&(FF1HDR.MAXID)
          CALL RptOptGET("EXAD", X): ExtAdoptRule = (X = 1)
       END IF
       CALL PutMSG("XXXXM064")
       CALL RWRKGet("BASE", RC, PTR&)
       WHILE (0 < RC) AND (RC < 9)
          IF WFRELREC.ID > 0 THEN
             IF RELSEQ THEN
                IF IDLIST(WFRELREC.ID) > 0 THEN
                   NDXRECS = NDXRECS + 1: NDXID(NDXRECS) = WFRELREC.ID
                   IDLIST(WFRELREC.ID) = -1
                END IF
             END IF
             IF SHOWLIN THEN
                IF WFGLREC.DGL = 0 THEN
                   ExtAdopt = ((WFANCREC.PSTAT AND 5) > 0)
                   ELSE
                   ExtAdopt = ((WFRELREC.CSTAT AND 5) > 0)
                END IF
                X = RELATE(WFRELREC.ID)
                IF (X = 0) OR (ExtAdoptRule AND ((X AND 128) > 0) AND (NOT ExtAdopt)) THEN
                   X = 256 * (-WFGLREC.AGL) + WFGLREC.DGL
                   IF ExtAdoptRule THEN X = X - 128 * ExtAdopt
                   RELATE(WFRELREC.ID) = X
                   RELPTR&(WFRELREC.ID) = PTR&
                END IF
             END IF
          END IF
          CALL RWRKGet("WRKREC", RC, PTR&)
       WEND
       IF RELSEQ THEN
          FOR X = 1 TO FF1HDR.MAXID
              Y = IDLIST(X)
              IF Y > 0 THEN
                 NDXRECS = NDXRECS + 1: NDXID(NDXRECS) = X
                 ELSE
                 IF Y < 0 THEN IDLIST(X) = 1
              END IF
          NEXT X
          INDEXED = -1
       END IF
    END IF
    IF INDEXED THEN REDIM PRESERVE NDXID(NDXRECS)
    D1Y = VAL(RIGHT$(DATE$, 4))
    D1MD = (100 * VAL(LEFT$(DATE$, 2)) + VAL(MID$(DATE$, 4, 2)))
    CDATE$ = XRPTDate$(D1Y, D1MD, 0)
    IF (SCNT = 1) THEN
       FORMFD = -1
       ELSE
       CALL PutMSG("GRPRM008"): IF ASC(A$) = 27 THEN 396
       FORMFD = (A$ = "1")
       IF NOT FORMFD THEN LSET MA$ = " 0 0 0 0 0 0 0 0 0 0 0 0"
       IF FORMFD THEN
          CALL PutMSG("GRPRM011"): IF ASC(A$) = 27 THEN 396
          IF (A$ <> SNGLKEY$(2)) THEN
             ID = LOID
             ELSE
             GOSUB 350: IF NOT OKAY THEN 396
          END IF
       END IF
       RptOPTION.SNGLREF = -1  ' Allow to Ask About Limiting References Recorded
    END IF
    IF NOT FORMFD THEN POPT$ = SNGLKEY$(4): PRTHDRID = 0 ELSE POPT$ = " "
    CALL PrintOPEN: IF NOT OKAY THEN 1090
    IF LEN(PLINE$) < 132 THEN PLINE$ = SPACE$(132)
    AFMT = FF2Hdr.REVLVL
    IF (SCNT = 1) THEN
       PRT.SECTION = "F"
       ID = RID: GOSUB PrtReport: GOTO 395
       ELSE
       IF FORMFD THEN PRT.SECTION = "F" ELSE PRT.SECTION = "M"
    END IF
    FSID = SID: FRID = ID: FirstONE = -1: GOSUB 355

330 ' Print Family Group Report/Forms
    IF (NOT FORMFD) AND (A = 27) THEN 395
    IF NOT FirstONE THEN GOSUB 355: IF ID = 0 THEN 395
    FirstONE = 0
    IF POPT$ <> SNGLKEY$(4) THEN
       SYSVAR.NVAR1 = ID: CALL PutMSG("GRPRM012")
       POPT$ = A$: IF ASC(POPT$) = 27 THEN 395
       IF (POPT$ = SNGLKEY$(2)) THEN 330
    END IF
    IF (NOT FORMFD) AND (RID <> 0) THEN
       IF BLNKBTWN THEN CALL WriteLINE
       LSET PLINE$ = LINESEP1$: CALL WriteLINE
       IF BLNKBTWN THEN CALL WriteLINE
    END IF
335 RID = ID: GOSUB PrtReport
    IF (OPT < 2) OR (NOT FORMFD) THEN 330
    IF A <> 27 THEN 330
    SYSVAR.NVAR1 = RID
    CALL PutMSG("GRPRM013")
    POPT$ = A$: IF ASC(POPT$) = 27 THEN 395
    IF POPT$ = SNGLKEY$(2) THEN 330
    PRINT #4, PT$(1) + PT$(2);
    GOTO 335

350 ' Get First ID Number to Process
    CALL FmtFIND("GRPRS012"): OKAY = -1: ID = 0
    WHILE OKAY AND (ID = 0)
       CALL MsgFORMAT: IF OKAY THEN OKAY = (A <> 27)
       IF OKAY THEN
          CALL FmtFindFLD("ID")
          IF OKAY THEN
             ID = VAL(UTXT$(FFLD.UTXT)): OKAY = (ID <> 0)
             IF OKAY THEN
                IF (IDLIST(ID) < 0) THEN PAGE.COUNT = -IDLIST(ID)
             END IF
          END IF
       END IF
    WEND
    IF OKAY THEN
       IF INDEXED THEN
          SID = 1
          WHILE (SID < NDXRECS) AND (NDXID(SID) <> ID): SID = SID + 1: WEND
          OKAY = (NDXID(SID) = ID)
          ELSE
          SID = ID
       END IF
    END IF
    RETURN

355 ' Find next ID to Print
    IF SID > 0 THEN 362 ELSE RID = 0
    IF INDEXED THEN
       SID = 1: ID = NDXID(1)
       ELSE
       SID = LOID: ID = SID
    END IF
    GOTO 365
357 IF NOT INDEXED THEN
       IF RID = 0 THEN SID = IDLIST(0) ELSE SID = RID
       ID = SID: GOTO 365
    END IF
360 SID = SID + 1: IF SID > NDXRECS THEN ID = 0: GOTO 391
    ID = NDXID(SID)
    IF ((ID > 0) AND (RID = ID)) OR (RID = 0) THEN 365 ELSE GOTO 360
362 SID = SID + 1: IF SID > NDXRECS THEN ID = 0: GOTO 391
    IF NOT INDEXED THEN ID = SID ELSE ID = NDXID(SID)
365 IF (IDLIST(ID) = 0) OR (ID = 0) THEN 362
    CALL FF1GetRec(ID)
    IF PRT.OPT <> 1 THEN
       IF FFRID > 0 THEN RSET UTXT$(FFRID) = FNX$(ID)
       IF FFNAME > 0 THEN LSET UTXT$(FFNAME) = FMTNAME$(1)
       CALL WinDisplayDATA(FMT002)
    END IF
    IF ((MID$(IA$, 5, 1) = "1") AND (FF1REC.EDUC <> 0)) THEN 391
    IF ((MID$(IA$, 6, 1) = "1") AND (FF1REC.WORK <> 0)) THEN 391
    IF ((MID$(IA$, 7, 1) = "1") AND (FF1REC.MIL <> 0)) THEN 391
    IF ((MID$(IA$, 8, 1) = "1") AND (FF1REC.HLTH <> 0)) THEN 391
    IF ((MID$(IA$, 11, 1) = "1") AND (FF1REC.EVENT <> 0)) THEN 391
    IF (BYPASS) AND (MID$(IA$, 2, 1) <> "2") THEN
       FID = FF1REC.FID: MID = FF1REC.MID
       IF ((FID > 0) * (IDLIST(FID) <> 0)) + ((MID > 0) * (IDLIST(MID) <> 0)) = 0 THEN 381
       IF (FF1REC.OLDCH <> 0) THEN 381
       IF ((MID$(IA$, 3, 1) = "1") AND (FF1REC.ADRS <> 0)) THEN 381
       IF ((MID$(IA$, 4, 1) <> "2") AND (FF1REC.COM <> 0) AND (MID$(IA$, 12, 1) = "2")) THEN 381
       IF ((MID$(IA$, 1, 1) <> "1") OR (FF1REC.SPOUSE = 0)) THEN 362
       CALL FF3GetRec(FF1REC.SPOUSE)
       IF (FF3SPOUSE.MLOC <> 0) THEN 381
       IF ((MID$(IA$, 3, 1) = "1") AND (FF3SPOUSE.LOC <> 0)) THEN 381
       IF ((MID$(IA$, 4, 1) <> "2") AND (FF3SPOUSE.COM <> 0)) THEN 381
       IF ID = FF3SPOUSE.SP1ID THEN
          CSPID = FF3SPOUSE.SP2ID: SC = FF3SPOUSE.SP2NXT
          ELSE
          CSPID = FF3SPOUSE.SP1ID: SC = FF3SPOUSE.SP1NXT
       END IF
       IF SC <> 0 THEN 381
       IF CSPID = 0 THEN 362 ELSE CALL FF1GetRec(CSPID)
       IF (MID$(IA$, 4, 1) = "4") AND (FF1REC.COM <> 0) THEN 381
       FID = FF1REC.FID: MID = FF1REC.MID: IF (FID > 0) OR (MID > 0) THEN 381
       IF (FF1REC.BLOC = 0) AND (FF1REC.BY = 0) AND (FF1REC.BMD = 0) AND ((FF1REC.DY = 0) OR (FF1REC.DY = 9999)) AND (FF1REC.DMD = 0) THEN 362
    END IF
381 IF (BYPASS2) AND (MID$(IA$, 1, 1) = "1") THEN
       CALL FF1GetRec(ID)
       SC = FF1REC.SPOUSE: IF (FF1REC.SEX <> FGENDR$(2)) OR (SC = 0) THEN 391
       IF (MID$(IA$, 4, 1) <> "1") OR (FF1REC.COM = 0) THEN
          CALL FF3GetRec(SC)
          IF ID = FF3SPOUSE.SP1ID THEN
             CSPID = FF3SPOUSE.SP2ID: SC = FF3SPOUSE.SP1NXT
             ELSE
             CSPID = FF3SPOUSE.SP1ID: SC = FF3SPOUSE.SP2NXT
          END IF
          IF SC <> 0 THEN 391
          IF (CSPID = 0) OR (IDLIST(CSPID) = 0) THEN 391
          IF (MID$(IA$, 2, 1) = "2") OR (FF1REC.OLDCH = 0) THEN 362
          CP = FF1REC.OLDCH
          WHILE CP <> 0
             CALL FF1GetRec(CP)
             IF CSPID = FF1REC.FID THEN CP = FF1REC.MCH ELSE CP = 0
          WEND
          IF CSPID = FF1REC.FID THEN 362
       END IF
    END IF
391 RETURN

395 ' Finished Printing Reports
    PRT.AGAIN = 0: PRT.OPEN = CHR$(0): CALL RptCLOSE
    PRT.OPEN = CHR$(255)
    'IF PAGE.LINE > 0 THEN CALL PageBREAK
    'IF RPTOPTION.BLDREF THEN CALL RREFClose: RPTOPTION.BLDREF = 0
396 ' Reset IDLIST()
    IF STYPE > 1 THEN
       FOR X = 1 TO FF1HDR.MAXID
           IF IDLIST(X) <> 0 THEN
              IDLIST(X) = 1
           END IF
       NEXT X
    END IF
    IF NOT OKAY THEN 399
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
    IF PRT.AGAIN THEN
       PAGE.COUNT = PAGE.BGNPAGE - 1
       IF SCNT = 1 THEN GOSUB PrtReport: GOTO 395
       SID = FSID: RID = FRID: ID = RID: GOTO 330
    END IF
399 ' Return to Main Program
    CALL RptCLOSE
    CALL FamCLOSE
    EXIT SUB

PrtReport: ' Print Family Group Report
    IF FORMFD THEN PAGE.BGNPAGE = PAGE.COUNT + 1
    IF FF1REC.RID <> RID THEN CALL FF1GetRec(RID)
    REDIM APTR(MAXADR), AIX(MAXADR), ADRDATE$(MAXADR)
    AX = -1: APTR(0) = FF1REC.ADRS: IF APTR(0) <> 0 THEN AX = 0
    SC = FF1REC.SPOUSE: CC = FF1REC.OLDCH: ComPTR = FF1REC.COM
    CALL BldMLIST
    EVENTPTR = FF1REC.EVENT
    OA(1) = FF1REC.HLTH: OA(2) = FF1REC.EDUC
    OA(3) = FF1REC.WORK: OA(4) = FF1REC.MIL
    GOSUB 600: GOSUB 610: MCNT = 0
    LSET SYSVAR.SUBJECT = FMTNAME$(0)
    IF PRTHDRID THEN
       X = LEN(RTRIM$(SYSVAR.SUBJECT))
       IF X > 2 THEN
          MID$(SYSVAR.SUBJECT, X + 2, 255) = "(" + RIDLbl$ + FNX$(RID) + ")"
       END IF
    END IF
    GOSUB 2300: IF A = 27 THEN 590
    IF MID$(IA$, 1, 1) <> "2" THEN
       GOSUB 600
       WHILE SC <> 0
          CALL FF3GetRec(SC): CALL GetNextMARRIAGE
          MLOC = FF3SPOUSE.MLOC
          IF RID = FF3SPOUSE.SP1ID THEN
             CSPID = FF3SPOUSE.SP2ID
             ELSE
             CSPID = FF3SPOUSE.SP1ID
          END IF
          ComPTR = FF3SPOUSE.COM
          APTR(AX + 1) = FF3SPOUSE.LOC: IF APTR(AX + 1) > 0 THEN AX = AX + 1
          IF CSPID > 0 THEN
             CALL FF1GetRec(CSPID): COMPTR2 = FF1REC.COM
             GOSUB 610
             ELSE
             COMPTR2 = 0: NameDATA.NAME = "(" + Unknown$ + ")"
             FID = 0: MID = 0
          END IF
          GOSUB 630: GOSUB 2350: IF A = 27 THEN 590
          GOSUB 600
       WEND
       CSPID = 0: MID = 0: FID = 0
       IF MID$(MA$, 1, 1) = "+" THEN
          RH = VAL(MID$(MA$, 1, 2))
          ELSE
          RH = VAL(MID$(MA$, 1, 2)) - MCNT
       END IF
       MLOC = 0
       WHILE RH > 0
          GOSUB 2350: IF A = 27 THEN 590 ELSE RH = RH - 1
       WEND
    END IF
    RH = VAL(MID$(MA$, 3, 2))
    IF (MID$(IA$, 2, 1) <> "2") AND ((CC <> 0) OR (RH <> 0)) THEN
       GOSUB 2400: IF A = 27 THEN 590
       CH = 0
       WHILE CC <> 0
          CALL FF1GetRec(CC): ComPTR = FF1REC.COM
          STAT1 = ASC(FF1REC.STATUS)
          MSTAT = ((STAT1 \ 4) AND 3): FSTAT = (STAT1 AND 3)
          IF RID = FF1REC.MID THEN
             NCH = FF1REC.MCH: PID = FF1REC.FID: CSTAT = MSTAT: PSTAT = FSTAT
             ELSE
             NCH = FF1REC.FCH: PID = FF1REC.MID: CSTAT = FSTAT: PSTAT = MSTAT
          END IF
          NameDATA = BNDATA
          NameDATA.RID = SFILL$(CC, 5)
          NameDATA.PID = SFILL$(PID, 5)
          NameDATA.NAME = FMTNAME$(0): GOSUB 615: GOSUB 650
          MarriageDATA = BMDATA
          IF FF1REC.SPOUSE <> 0 THEN
             CALL FF3GetRec(FF1REC.SPOUSE)
             IF RptOPTION.FREEFORM THEN MLOC = FF3SPOUSE.MLOC ELSE MLOC = 0
             GOSUB 630
             IF CC = FF3SPOUSE.SP1ID THEN CSPID = FF3SPOUSE.SP2ID ELSE CSPID = FF3SPOUSE.SP1ID
             IF CSPID = 0 THEN
                MarriageDATA.SPNAME = "(" + Unknown$ + ")"
                ELSE
                CALL FF1GetRec(CSPID)
                MarriageDATA.SPID = SFILL$(CSPID, 5)
                MarriageDATA.SPNAME = FMTNAME$(0)
                IF ULALL THEN CALL ULFLDS(4)
             END IF
          END IF
          GOSUB 2410: IF A = 27 THEN 590
          CC = NCH
       WEND
       IF NOT RptOPTION.FREEFORM THEN
          IF MID$(MA$, 3, 1) <> "+" THEN RH = RH - CH
          IF RH > 0 THEN
             NameDATA = BNDATA
             MarriageDATA = BMDATA
             WHILE RH > 0
                GOSUB 2410: IF A = 27 THEN 590
                RH = RH - 1
             WEND
          END IF
       END IF
    END IF

460 ' Process RESIDENCEes
    NBL = -1
    RH = VAL(MID$(MA$, 5, 2))
    IF (MID$(IA$, 3, 1) <> "2") AND ((APTR(0) <> 0) OR (RH <> 0)) THEN
       GOSUB PrintRESH: IF A = 27 THEN 590
       Y = 0
       WHILE APTR(Y) <> 0
          IF NOT (AX < MAXADR) THEN
             MAXADR = MAXADR + 10
             REDIM PRESERVE APTR(MAXADR), AIX(MAXADR), ADRDATE$(MAXADR)
          END IF
          ADRDATE$(Y) = SPACE$(8)
          CALL FF2GetRec(APTR(Y))
          APTR(AX + 1) = FF2PFX.NXT
          IF (APTR(AX + 1) <> 0) AND (AX < MAXADR) THEN AX = AX + 1
          MID$(ADRDATE$(Y), 1, 4) = ZFILL$(FF2PFX.D1Y, 4)
          MID$(ADRDATE$(Y), 5, 4) = ZFILL$(FF2PFX.D1MD, 4)
          AIX(Y) = Y: Z = Y - 1
          WHILE (Z >= 0)
             IF LEFT$(ADRDATE$(AIX(Z)), 4) < LEFT$(ADRDATE$(AIX(Z + 1)), 4) THEN
                SWAP AIX(Z), AIX(Z + 1): Z = Z - 1
                ELSE
                Z = -1
             END IF
          WEND
          IF Y < MAXADR THEN Y = Y + 1 ELSE APTR(Y) = 0
       WEND
       Y = 0
       WHILE Y <= AX
          Z = AIX(Y)
          APTR = APTR(Z)
          GOSUB PrintRESI
          IF A = 27 THEN 590 ELSE Y = Y + 1
       WEND
       RH = VAL(MID$(MA$, 5, 2))
       IF MID$(MA$, 5, 1) <> "+" THEN RH = RH - AX
       IF RH > 0 THEN
          BFH = BF: BF = -1
          WHILE RH > 0
             GOSUB PrintRESI: IF A = 27 THEN 590
             RH = RH - 1
          WEND
          BF = BFH
       END IF
    END IF

470 ' Process EVENTS
    RH = VAL(MID$(MA$, 21, 2))
    IF (MID$(IA$, 11, 1) <> "2") AND ((EVENTPTR <> 0) OR (RH <> 0)) THEN
       GOSUB PrintEventH: IF A = 27 THEN 590
       WHILE EVENTPTR <> 0
          GOSUB PrintEVENT: EVENTPTR = FF3EVENT.NXT
       WEND
       IF MID$(MA$, 21, 1) <> "+" THEN RH = RH - EVNTCNT
       IF RH > 0 THEN
          BFH = BF: BF = -1
          WHILE RH > 0
             GOSUB PrintEVENT: IF A = 27 THEN 590
             RH = RH - 1
          WEND
          BF = BFH
       END IF
    END IF

500 ' Process OTHER Info
    FOR OT = 1 TO 4: GOSUB 2700: NEXT OT
590 IF FORMFD AND (PAGE.LINE > 0) THEN CALL PageBREAK
    IF A = 27 THEN LSET PLINE$ = " ": LOFFSET = TWIDTH
    RETURN

600 ' Initialise Data Areas - Individual
    NameDATA = BNDATA
    IF RptOPTION.PRID AND (NOT RptOPTION.FREEFORM) THEN
       X = LEN(XRID$): MID$(NameDATA.NAME, 53 - X, X) = XRID$
       MID$(NameDATA.NAME, 47, 5) = STRING$(5, FC)
    END IF
    FatherDATA = BPDATA
    IF RptOPTION.PRID AND (NOT RptOPTION.FREEFORM) THEN
       MID$(FatherDATA.NAME, 53 - X, X) = MID$(NameDATA.NAME, 53 - X, X)
    END IF
    MotherDATA = FatherDATA
    MarriageDATA = BMDATA
    RETURN

610 ' Place Information in Data Areas
    NameDATA.NAME = FMTNAME$(0)
    IF RptOPTION.PRID AND (NOT RptOPTION.FREEFORM) THEN
       L = LEN(XRID$): MID$(NameDATA.NAME, 53 - L, L) = XRID$
       RSET RID$ = FNX$(FF1REC.RID): NameDATA.RID = RID$
       MID$(NameDATA.NAME, 47, 5) = NameDATA.RID
    END IF
    FID = FF1REC.FID: MID = FF1REC.MID
    STAT1 = ASC(FF1REC.STATUS)
    MSTAT = ((STAT1 \ 4) AND 3): FSTAT = (STAT1 AND 3)
    GOSUB 620

615 ' Name Record Information Common to Subject, Spouse, Child
    CALL GetBDEvents(BD$, BPL$, DD$, DPL$, "")
    IF RTRIM$(BD$) <> "" THEN NameDATA.BDATE = BD$
    IF RTRIM$(DD$) <> "" THEN NameDATA.DDATE = DD$
    IF RTRIM$(BPL$) <> "" THEN NameDATA.BPLACE = BPL$
    IF RTRIM$(DPL$) <> "" THEN NameDATA.DPLACE = DPL$
    IF RTRIM$(FF1REC.SEX) <> "" THEN NameDATA.SEX = FF1REC.SEX
    IF ULALL THEN CALL ULFLDS(1)
    RETURN

620 ' Father, Mother Information
    XID = FF1REC.RID
    IF FID <> 0 THEN
       CALL FF1GetRec(FID): STAT1 = ASC(FF1REC.STATUS)
       FatherDATA.NAME = FMTNAME$(0)
       IF RptOPTION.PRID AND (NOT RptOPTION.FREEFORM) THEN
          L = LEN(XRID$)
          RSET RID$ = FNX$(FID): RSTAT = FSTAT
          MID$(FatherDATA.NAME, 53 - L, L) = XRID$
          MID$(FatherDATA.NAME, 47, 5) = RID$
       END IF
       CALL GetBDEvents(BD$, "", DD$, "", AGE$)
       IF RTRIM$(BD$) <> "" THEN FatherDATA.BDATE = BD$
       IF RTRIM$(DD$) <> "" THEN FatherDATA.DDATE = DD$
       IF (FF1REC.BY > 0) AND (FF1REC.DY > 0) AND (FF1REC.DY < 9999) THEN
          FatherDATA.AGE = LTRIM$(AGE$)
       END IF
       IF ULALL THEN CALL ULFLDS(2)
    END IF
    IF MID <> 0 THEN
       CALL FF1GetRec(MID): STAT1 = ASC(FF1REC.STATUS)
       RSET RID$ = FNX$(MID): RSTAT = MSTAT
       MotherDATA.NAME = FMTNAME$(0)
       IF RptOPTION.PRID AND (NOT RptOPTION.FREEFORM) THEN
          L = LEN(XRID$)
          MID$(MotherDATA.NAME, 53 - L, L) = XRID$
          MID$(MotherDATA.NAME, 47, 5) = RID$
       END IF
       CALL GetBDEvents(BD$, "", DD$, "", AGE$)
       IF RTRIM$(BD$) <> "" THEN MotherDATA.BDATE = BD$
       IF RTRIM$(DD$) <> "" THEN MotherDATA.DDATE = DD$
       IF (FF1REC.BY > 0) AND (FF1REC.DY > 0) AND (FF1REC.DY < 9999) THEN
          MotherDATA.AGE = LTRIM$(AGE$)
       END IF
       IF ULALL THEN CALL ULFLDS(3)
    END IF
    CALL FF1GetRec(XID)
    RETURN

630 ' Marriage Information
    IF (FF3SPOUSE.D1MD <> 0) OR (FF3SPOUSE.D1Y <> 0) THEN
       MarriageDATA.MDATE = XRPTDate$(FF3SPOUSE.D1MD, FF3SPOUSE.D1Y, DS)
    END IF
    IF (FF3SPOUSE.D2MD <> 0) OR (FF3SPOUSE.D2Y <> 0) THEN
       MarriageDATA.DVDATE = XRPTDate$(FF3SPOUSE.D2MD, FF3SPOUSE.D2Y, DS)
    END IF
    IF MLOC <> 0 THEN
       CALL FF3GetRec(MLOC)
       IF (FF3PFX.RTYPE = CHR$(12)) AND (FF3PLACE2.PNEXT <> 0) THEN
          CALL FF3GetRec(FF3PLACE2.PNEXT)
       END IF
       IF RTRIM$(MPLACE) <> "" THEN
          MarriageDATA.MPLACE = MPLACE
       END IF
       IF RTRIM$(DVPLACE) <> "" THEN
          MarriageDATA.DVPLACE = DVPLACE
       END IF
    END IF
    IF ULALL THEN CALL ULFLDS(4)
    RETURN

650 ' Count Children of Child
    GCH = 0
    IF CC <> 0 THEN
       NGCH = FF1REC.OLDCH
       WHILE NGCH <> 0
          GCH = GCH + 1: CALL FF1GetRec(NGCH): NGCH = 0
          IF CC = FF1REC.FID THEN
             NGCH = FF1REC.FCH
             ELSE
             IF CC = FF1REC.MID THEN NGCH = FF1REC.MCH
          END IF
       WEND
       CALL FF1GetRec(CC)
    END IF
    IF GCH > 0 THEN NameDATA.NUMCH = SFILL$(GCH, LEN(NameDATA.NUMCH))
    RETURN

2300 ' Print Subject Information
     IF SHOWLIN THEN GOSUB 62300: IF A = 27 THEN 2325
     IF RptOPTION.FREEFORM THEN GOTO FF2300
     MID$(PLINE$, 1, 255) = RFLBL$(5): Z = RFLTH(5) + 2 + (RFLTH(5) = 0)
     MID$(PLINE$, Z, 64 - Z) = NameDATA.NAME
     MID$(PLINE$, 65, 255) = RFLBL$(6): Z = 66 + RFLTH(6)
     MID$(PLINE$, Z, 52) = FatherDATA.NAME
     IF NOT BF THEN
        PAGE.XID = RID: PAGE.XTYPE = 2: CALL RREFAdd
        PAGE.XID = FID: PAGE.XTYPE = 5: CALL RREFAdd
     END IF
     CALL PrintLINE
     IF STYPE > 1 THEN IDLIST(RID) = -PAGE.COUNT
     MID$(PLINE$, 3, 255) = RFLBL$(8): Z = 4 + RFLTH(8)
     MID$(PLINE$, Z, 12) = NameDATA.BDATE: Z = Z + 14
     MID$(PLINE$, Z, 255) = RFLBL$(9): Z = Z + 1 + RFLTH(9)
     MID$(PLINE$, Z, 64 - Z) = NameDATA.BPLACE
     MID$(PLINE$, 67, 255) = RFLBL$(8): Z = 68 + RFLTH(8)
     MID$(PLINE$, Z, 12) = FatherDATA.BDATE: Z = Z + 14
     MID$(PLINE$, Z, 255) = RFLBL$(10): Z = Z + RFLTH(10) + 1
     MID$(PLINE$, Z, 12) = FatherDATA.DDATE
     Z = Z + LEN(FatherDATA.DDATE) + 1
     IF FatherDATA.AGE <> STRING$(3, FC) THEN
        MID$(PLINE$, Z, 255) = RFLBL$(12): Z = Z + RFLTH(12) + 1
        MID$(PLINE$, Z, 3) = FatherDATA.AGE
     END IF
     CALL PrintLINE: IF A = 27 THEN 2325
     MID$(PLINE$, 3, 255) = RFLBL$(10): Z = 4 + RFLTH(10)
     MID$(PLINE$, Z, 12) = NameDATA.DDATE: Z = Z + 14
     MID$(PLINE$, Z, 6) = RFLBL$(11): Z = Z + 1 + RFLTH(11)
     MID$(PLINE$, Z, 64 - Z) = NameDATA.DPLACE
     MID$(PLINE$, 65, 255) = RFLBL$(7): Z = 66 + RFLTH(7)
     MID$(PLINE$, Z, 52) = MotherDATA.NAME
     IF NOT BF THEN PAGE.XID = MID: PAGE.XTYPE = 6: CALL RREFAdd
     CALL PrintLINE
     MID$(PLINE$, 67, 255) = RFLBL$(8): Z = 68 + RFLTH(8)
     MID$(PLINE$, Z, 12) = MotherDATA.BDATE: Z = Z + 14
     MID$(PLINE$, Z, 255) = RFLBL$(10): Z = Z + 1 + RFLTH(10)
     MID$(PLINE$, Z, 12) = MotherDATA.DDATE
     Z = Z + LEN(MotherDATA.DDATE) + 1
     IF MotherDATA.AGE <> STRING$(3, FC) THEN
        MID$(PLINE$, Z, 255) = RFLBL$(12): Z = Z + RFLTH(12) + 1
        MID$(PLINE$, Z, 3) = MotherDATA.AGE
     END IF
     IF MID$(IA$, 4, 1) = "2" THEN CALL PrintLINE: GOTO 2325
     PAGE.LOFST = 0: PAGE.COFST = 3: CWIDTH = 58
     RI = VAL(MID$(MA$, 7, 2)): GOSUB PrintCOM
     IF A <> 27 THEN CALL PrintLINE
2325 RETURN

FF2300: ' Print FreeForm Subject Information
     PAGE.LOFST = 1: PAGE.COFST = 3
     LSET CL$ = RFLBL$(5): Z = RFLTH(5) + 1 + (RFLTH(5) = 0)
     IF RptOPTION.PRID THEN
        MID$(CL$, Z + 1, 255) = "(" + RFLBL$(2): Z = Z + RFLTH(2) + 2
        MID$(CL$, Z, 6) = FNX$(RID) + ")": Z = INSTR(Z, CL$, "   ")
     END IF
     MID$(CL$, Z + 1, 52) = NameDATA.NAME: Z = INSTR(Z, CL$, "   ")
     IF (RTRIM$(NameDATA.BDATE) <> "") OR (RTRIM$(NameDATA.BPLACE) <> "") THEN
        MID$(CL$, Z, 255) = RFLBL$(8): Z = Z + RFLTH(8)
        IF RTRIM$(NameDATA.BDATE) <> "" THEN
           MID$(CL$, Z + 1, 12) = LTRIM$(NameDATA.BDATE): Z = INSTR(Z, CL$, "   ")
        END IF
        IF RTRIM$(NameDATA.BPLACE) <> "" THEN
           MID$(CL$, Z, 255) = RFLBL$(9): Z = Z + RFLTH(9)
           MID$(CL$, Z + 1, 255) = NameDATA.BPLACE: Z = INSTR(Z, CL$, "   ")
        END IF
     END IF
     IF (RTRIM$(NameDATA.DDATE) <> "") OR (RTRIM$(NameDATA.DPLACE) <> "") THEN
        MID$(CL$, Z, 255) = RFLBL$(10): Z = Z + RFLTH(10)
        IF RTRIM$(NameDATA.DDATE) <> "" THEN
           MID$(CL$, Z + 1, 12) = LTRIM$(NameDATA.DDATE): Z = INSTR(Z, CL$, "   ")
        END IF
        IF RTRIM$(NameDATA.DPLACE) <> "" THEN
           MID$(CL$, Z, 255) = RFLBL$(11): Z = Z + RFLTH(11)
           MID$(CL$, Z + 1, 255) = NameDATA.DPLACE
        END IF
     END IF
     PAGE.XID = RID: PAGE.XTYPE = 2
     CALL PutDATA
     IF FID <> 0 THEN
        MID$(CL$, 1, 255) = RFLBL$(6): Z = 1 + RFLTH(6)
        IF RptOPTION.PRID THEN
           MID$(CL$, Z + 1, 255) = "(" + RFLBL$(2): Z = Z + 2 + RFLTH(2)
           MID$(CL$, Z, 6) = FNX$(FID) + ")": Z = INSTR(Z, CL$, "   ")
        END IF
        MID$(CL$, Z + 1, 52) = FatherDATA.NAME: Z = INSTR(Z, CL$, "   ")
        IF RTRIM$(FatherDATA.BDATE) <> "" THEN
           MID$(CL$, Z, 255) = RFLBL$(8): Z = Z + RFLTH(8)
           MID$(CL$, Z + 1, 12) = LTRIM$(FatherDATA.BDATE)
           Z = INSTR(Z + 1, CL$, "   ")
        END IF
        IF (RTRIM$(FatherDATA.DDATE) <> "") THEN
           MID$(CL$, Z, 255) = RFLBL$(10): Z = Z + RFLTH(10)
           MID$(CL$, Z + 1, 12) = LTRIM$(FatherDATA.DDATE)
           Z = INSTR(Z, CL$, "   ")
           IF FatherDATA.AGE <> STRING$(3, FC) THEN
              MID$(CL$, Z, 255) = RFLBL$(12): Z = Z + RFLTH(12)
              MID$(CL$, Z + 1, 3) = FatherDATA.AGE
           END IF
        END IF
        PAGE.XID = FID: PAGE.XTYPE = 5
        CALL PutDATA
     END IF
     IF MID <> 0 THEN
        MID$(CL$, 1, 255) = RFLBL$(7): Z = 1 + RFLTH(7)
        IF RptOPTION.PRID THEN
           MID$(CL$, Z + 1, 255) = "(" + RFLBL$(2): Z = Z + 2 + RFLTH(2)
           MID$(CL$, Z, 6) = FNX$(MID) + ")": Z = INSTR(Z, CL$, "   ")
        END IF
        MID$(CL$, Z + 1, 52) = MotherDATA.NAME: Z = INSTR(Z, CL$, "   ")
        IF LTRIM$(MotherDATA.BDATE) <> "" THEN
           MID$(CL$, Z, 255) = RFLBL$(8): Z = Z + RFLTH(8)
           MID$(CL$, Z + 1, 12) = LTRIM$(MotherDATA.BDATE)
           Z = INSTR(Z, CL$, "   ")
        END IF
        IF LTRIM$(MotherDATA.DDATE) <> "" THEN
           MID$(CL$, Z, 255) = RFLBL$(10): Z = Z + RFLTH(10)
           MID$(CL$, Z + 1, 12) = LTRIM$(MotherDATA.DDATE)
           Z = INSTR(Z, CL$, "   ")
           IF MotherDATA.AGE <> STRING$(3, FC) THEN
              MID$(CL$, Z, 255) = RFLBL$(12): Z = Z + RFLTH(12)
              MID$(CL$, Z + 1, 3) = MotherDATA.AGE
           END IF
        END IF
        PAGE.XID = MID: PAGE.XTYPE = 6
        CALL PutDATA
     END IF
     IF MID$(IA$, 4, 1) = "2" THEN CALL PrintLINE: GOTO 2325
     GOSUB PrintCOM
     CALL PrintLINE
FF2325: RETURN

2350 ' Print Spouse Info
     CALL WriteLINE: IF A = 27 THEN 2380
     IF RptOPTION.FREEFORM THEN GOTO FF2350
     MID$(PLINE$, 1, 255) = RFLBL$(13): Z = 2 + RFLTH(13)
     MID$(PLINE$, Z, 64 - Z) = NameDATA.NAME
     MID$(PLINE$, 65, 255) = RFLBL$(6): Z = 66 + RFLTH(6)
     MID$(PLINE$, Z, 52) = FatherDATA.NAME
     IF NOT BF THEN
        PAGE.XID = CSPID: PAGE.XTYPE = 3: CALL RREFAdd
        PAGE.XID = FID: PAGE.XTYPE = 7: CALL RREFAdd
     END IF
     CALL PrintLINE
     MID$(PLINE$, 3, 255) = RFLBL$(8): Z = 4 + RFLTH(8)
     MID$(PLINE$, Z, 12) = NameDATA.BDATE: Z = Z + 14
     MID$(PLINE$, Z, 255) = RFLBL$(9): Z = Z + 1 + RFLTH(9)
     MID$(PLINE$, Z, 64 - Z) = NameDATA.BPLACE
     MID$(PLINE$, 67, 255) = RFLBL$(8): Z = 68 + RFLTH(8)
     MID$(PLINE$, Z, 12) = FatherDATA.BDATE: Z = Z + 14
     MID$(PLINE$, Z, 255) = RFLBL$(10): Z = Z + 1 + RFLTH(10)
     MID$(PLINE$, Z, 12) = FatherDATA.DDATE
     Z = Z + LEN(FatherDATA.DDATE) + 1
     IF FatherDATA.AGE <> STRING$(3, FC) THEN
        MID$(PLINE$, Z, 255) = RFLBL$(12): Z = Z + RFLTH(12) + 1
        MID$(PLINE$, Z, 3) = FatherDATA.AGE
     END IF
     CALL PrintLINE: IF A = 27 THEN 2325
     MID$(PLINE$, 3, 255) = RFLBL$(14): Z = 4 + RFLTH(14)
     MID$(PLINE$, Z, 11) = MarriageDATA.MDATE: Z = Z + 14
     MID$(PLINE$, Z, 255) = RFLBL$(15): Z = Z + 1 + RFLTH(15)
     MID$(PLINE$, Z, 64 - Z) = MarriageDATA.MPLACE
     MID$(PLINE$, 65, 255) = RFLBL$(7): Z = 66 + RFLTH(7)
     MID$(PLINE$, Z, 52) = MotherDATA.NAME
     IF NOT BF THEN PAGE.XID = MID: PAGE.XTYPE = 8: CALL RREFAdd
     CALL PrintLINE
     IF ((MarriageDATA.DVDATE) <> STRING$(11, FC)) OR (CSPID = 0) OR BF THEN
        MID$(PLINE$, 5, 255) = RFLBL$(16): Z = 6 + RFLTH(16)
        MID$(PLINE$, Z, 11) = MarriageDATA.DVDATE: Z = Z + 14
        MID$(PLINE$, Z, 255) = RFLBL$(17): Z = Z + 1 + RFLTH(17)
        MID$(PLINE$, Z, 64 - Z) = MarriageDATA.DVPLACE
     END IF
     MID$(PLINE$, 67, 255) = RFLBL$(8): Z = 68 + RFLTH(8)
     MID$(PLINE$, Z, 12) = MotherDATA.BDATE: Z = Z + 14
     MID$(PLINE$, Z, 255) = RFLBL$(10): Z = Z + 1 + RFLTH(10)
     MID$(PLINE$, Z, 12) = MotherDATA.DDATE
     Z = Z + LEN(MotherDATA.DDATE) + 1
     IF MotherDATA.AGE <> STRING$(3, FC) THEN
        MID$(PLINE$, Z, 255) = RFLBL$(12): Z = INSTR(Z, PLINE$, "   ") + 1
        MID$(PLINE$, Z, 3) = MotherDATA.AGE
     END IF
     IF (MarriageDATA.DVDATE <> STRING$(11, FC)) OR (CSPID = 0) OR BF THEN
        CALL PrintLINE: IF A = 27 THEN 2380
     END IF
     MID$(PLINE$, 3, 255) = RFLBL$(10): Z = 4 + RFLTH(10)
     MID$(PLINE$, Z, 12) = NameDATA.DDATE: Z = Z + 14
     MID$(PLINE$, Z, 255) = RFLBL$(11): Z = Z + 1 + RFLTH(11)
     MID$(PLINE$, Z, 64 - Z) = NameDATA.DPLACE
     CALL PrintLINE: IF A = 27 THEN 2380
     IF MID$(IA$, 4, 1) <> "2" THEN
        PAGE.COFST = 3
        RI = VAL(MID$(MA$, 7, 2))
        IF (MID$(IA$, 4, 1) <> "4") THEN COMPTR2 = 0
        GOSUB PrintCOM
     END IF
     MCNT = MCNT + 1
2380 RETURN

FF2350: ' Print FreeForm Spouse Information
     PAGE.LOFST = 1: PAGE.COFST = 3
     MID$(CL$, 1, 255) = RFLBL$(14): Z = 1 + RFLTH(14)
     IF LTRIM$(MarriageDATA.MDATE) <> "" THEN
        MID$(CL$, Z + 1, 11) = MarriageDATA.MDATE: Z = INSTR(Z, CL$, "   ")
     END IF
     IF RTRIM$(MarriageDATA.MPLACE) <> "" THEN
        MID$(CL$, Z, 255) = RFLBL$(15): Z = Z + RFLTH(15)
        MID$(CL$, Z + 1, 255) = MarriageDATA.MPLACE: Z = INSTR(Z, CL$, "   ")
     END IF
     IF (RTRIM$(MarriageDATA.DVDATE) <> "") THEN
        MID$(CL$, Z, 255) = RFLBL$(16): Z = Z + RFLTH(16)
        MID$(CL$, Z + 1, 11) = MarriageDATA.DVDATE: Z = INSTR(Z, CL$, "   ")
        IF (RTRIM$(MarriageDATA.DVPLACE) <> "") THEN
           MID$(CL$, Z, 255) = RFLBL$(17): Z = Z + RFLTH(17)
           MID$(CL$, Z + 1, 255) = MarriageDATA.DVPLACE: Z = INSTR(Z, CL$, "   ")
        END IF
     END IF
     MID$(CL$, Z, 255) = RFLBL$(13): Z = Z + RFLTH(13)
     IF RptOPTION.PRID THEN
        MID$(CL$, Z + 1, 255) = "(" + RFLBL$(2): Z = Z + RFLTH(2) + 2
        MID$(CL$, Z, 6) = FNX$(CSPID) + ")": Z = INSTR(Z, CL$, "   ")
     END IF
     MID$(CL$, Z + 1, 52) = NameDATA.NAME: Z = INSTR(Z, CL$, "   ")
     IF (RTRIM$(NameDATA.BDATE) <> "") OR (RTRIM$(NameDATA.BPLACE) <> "") THEN
        MID$(CL$, Z, 255) = RFLBL$(8): Z = Z + RFLTH(8)
        IF RTRIM$(NameDATA.BDATE) <> "" THEN
           MID$(CL$, Z + 1, 12) = LTRIM$(NameDATA.BDATE)
           Z = INSTR(Z, CL$, "   ")
        END IF
        IF RTRIM$(NameDATA.BPLACE) <> "" THEN
           MID$(CL$, Z, 255) = RFLBL$(9): Z = Z + RFLTH(9)
           MID$(CL$, Z + 1, 255) = NameDATA.BPLACE: Z = INSTR(Z, CL$, "   ")
        END IF
     END IF
     IF (RTRIM$(NameDATA.DDATE) <> "") OR (RTRIM$(NameDATA.DPLACE) <> "") THEN
        MID$(CL$, Z, 255) = RFLBL$(10): Z = Z + RFLTH(10)
        IF RTRIM$(NameDATA.DDATE) <> "" THEN
           MID$(CL$, Z + 1, 12) = LTRIM$(NameDATA.DDATE)
           Z = INSTR(Z, CL$, "   ")
        END IF
        IF RTRIM$(NameDATA.DPLACE) <> "" THEN
           MID$(CL$, Z, 255) = RFLBL$(11): Z = Z + RFLTH(11)
           MID$(CL$, Z + 1, 255) = NameDATA.DPLACE
        END IF
     END IF
     PAGE.XID = CSPID: PAGE.XTYPE = 3: CALL PutDATA
     IF FID <> 0 THEN
        MID$(CL$, 1, 255) = RFLBL$(6): Z = 1 + RFLTH(6)
        IF RptOPTION.PRID THEN
           MID$(CL$, Z + 1, 255) = "(" + RFLBL$(2): Z = Z + 2 + RFLTH(2)
           MID$(CL$, Z, 6) = FNX$(FID) + ")": Z = INSTR(Z, CL$, "   ")
        END IF
        MID$(CL$, Z + 1, 52) = FatherDATA.NAME: Z = INSTR(Z, CL$, "   ")
        IF LTRIM$(FatherDATA.BDATE) <> "" THEN
           MID$(CL$, Z, 255) = RFLBL$(8): Z = Z + RFLTH(8)
           MID$(CL$, Z + 1, 12) = LTRIM$(FatherDATA.BDATE)
           Z = INSTR(Z + 1, CL$, "   ")
        END IF
        IF LTRIM$(FatherDATA.DDATE) <> "" THEN
           MID$(CL$, Z, 255) = RFLBL$(10): Z = Z + RFLTH(10)
           MID$(CL$, Z + 1, 12) = LTRIM$(FatherDATA.DDATE)
           Z = INSTR(Z, CL$, "   ")
           IF FatherDATA.AGE <> SPACE$(3) THEN
              MID$(CL$, Z, 255) = RFLBL$(12): Z = Z + RFLTH(12)
              MID$(CL$, Z + 1, 3) = FatherDATA.AGE
           END IF
        END IF
        PAGE.XID = FID: PAGE.XTYPE = 7: CALL PutDATA
     END IF
     IF MID <> 0 THEN
        MID$(CL$, 1, 255) = RFLBL$(7): Z = 1 + RFLTH(7)
        IF RptOPTION.PRID THEN
           MID$(CL$, Z + 1, 255) = "(" + RFLBL$(2): Z = Z + 2 + RFLTH(2)
           MID$(CL$, Z, 6) = FNX$(MID) + ")": Z = INSTR(Z, CL$, "   ")
        END IF
        MID$(CL$, Z + 1, 52) = MotherDATA.NAME: Z = INSTR(Z, CL$, "   ")
        IF LTRIM$(MotherDATA.BDATE) <> "" THEN
           MID$(CL$, Z, 255) = RFLBL$(8): Z = Z + RFLTH(8)
           MID$(CL$, Z + 1, 12) = LTRIM$(MotherDATA.BDATE)
           Z = INSTR(Z, CL$, "   ")
        END IF
        IF LTRIM$(MotherDATA.DDATE) <> "" THEN
           MID$(CL$, Z, 255) = RFLBL$(10): Z = Z + RFLTH(10)
           MID$(CL$, Z + 1, 12) = LTRIM$(MotherDATA.DDATE)
           Z = INSTR(Z, CL$, "   ")
           IF MotherDATA.AGE <> SPACE$(3) THEN
              MID$(CL$, Z, 255) = RFLBL$(12): Z = Z + RFLTH(12)
              MID$(CL$, Z + 1, 3) = MotherDATA.AGE
           END IF
        END IF
        PAGE.XID = MID: PAGE.XTYPE = 8: CALL PutDATA
     END IF
     IF MID$(IA$, 4, 1) <> "2" THEN
        IF (MID$(IA$, 4, 1) <> "4") THEN COMPTR2 = 0
        GOSUB PrintCOM
     END IF
FF2380: MCNT = MCNT + 1
        RETURN

2400 ' Print Child Information Header
     CALL WriteLINE
     IF A <> 27 THEN
        LSET PLINE$ = ChildHDR$
        CALL PrintLINE
     END IF
     RETURN

2410 ' Print Child Information
     CH = CH + 1: MID$(PLINE$, 1, 3) = RIGHT$("  " + STR$(CH), 2) + "."
     IF CSTAT > 0 THEN
        FF1REC.STATUS = CHR$(CSTAT)
        MID$(PLINE$, 3, 1) = FDSTAT$("FID"): CSTAT = 0
     END IF
     IF RptOPTION.FREEFORM THEN GOTO FF2410
     MID$(PLINE$, 5, 1) = NameDATA.SEX
     MID$(PLINE$, 8, 32) = NameDATA.NAME
     MID$(PLINE$, 42, 12) = NameDATA.BDATE
     MID$(PLINE$, 66, 12) = NameDATA.DDATE
     MID$(PLINE$, 61, 3) = NameDATA.NUMCH
     MID$(PLINE$, 90, 11) = MarriageDATA.MDATE
     IF RptOPTION.PRID THEN
        L = LEN(XSPID$)
        MID$(PLINE$, 125 - L, L) = XSPID$
        MID$(PLINE$, 119, 5) = MarriageDATA.SPID
     END IF
     IF NOT BF THEN PAGE.XID = CC: PAGE.XTYPE = 4: CALL RREFAdd
     CALL PrintLINE
     MID$(PLINE$, 42, 22) = NameDATA.BPLACE
     MID$(PLINE$, 66, 22) = NameDATA.DPLACE
     MID$(PLINE$, 90, 35) = MarriageDATA.SPNAME
     IF RptOPTION.PRID THEN
        L = LEN(XRID$)
        MID$(PLINE$, 8, L) = XRID$
        MID$(PLINE$, 2 + L, 5) = NameDATA.RID
        MID$(PLINE$, 7 + L, 1) = ";"
        L2 = LEN(XOPID$)
        MID$(PLINE$, 9 + L, L2) = XOPID$
        MID$(PLINE$, 3 + L + L2, 5) = NameDATA.PID
        IF PSTAT <> 0 THEN
           FF1REC.STATUS = CHR$(PSTAT)
           MID$(PLINE$, 4 + L + L2, 1) = FDSTAT$("FID"): PSTAT = 0
        END IF
     END IF
     IF NOT BF THEN PAGE.XID = CSPID: PAGE.XTYPE = 9: CALL RREFAdd
     CALL PrintLINE
     IF MID$(IA$, 12, 1) <> "2" THEN
        PAGE.COFST = 8: PAGE.LOFST = 0
        RI = VAL(MID$(MA$, 23, 2)): GOSUB PrintCOM
     END IF
     IF CHILDSEP AND ((NCH > 0) OR ((RH > 1) AND (NOT RptOPTION.FREEFORM))) THEN
        CALL WriteLINE
     END IF
2490 RETURN

FF2410: ' FreeForm Child Information
     PAGE.LOFST = 5: PAGE.COFST = 8
     MID$(CL$, 1, 1) = NameDATA.SEX: Z = 4
     IF RptOPTION.PRID THEN
        MID$(CL$, Z, 255) = "(" + RFLBL$(2) + LTRIM$(NameDATA.RID)
        Z = INSTR(Z, CL$, "   ")
        MID$(CL$, Z, 255) = RFLBL$(3) + LTRIM$(NameDATA.PID)
        Z = INSTR(Z, CL$, "   ")
        IF PSTAT <> 0 THEN
           FF1REC.STATUS = CHR$(PSTAT)
           MID$(CL$, Z, 1) = FDSTAT$("FID"): PSTAT = 0
           Z = INSTR(Z, CL$, " ")
        END IF
        MID$(CL$, Z, 1) = ")": Z = Z + 2
     END IF
     MID$(CL$, Z, 52) = NameDATA.NAME: Z = INSTR(Z, CL$, "   ")
     IF (LTRIM$(NameDATA.BDATE) <> "") OR (RTRIM$(NameDATA.BPLACE) <> "") THEN
        MID$(CL$, Z, 255) = RFLBL$(8): Z = Z + RFLTH(8)
        IF RTRIM$(NameDATA.BDATE) <> "" THEN
           MID$(CL$, Z + 1, 12) = LTRIM$(NameDATA.BDATE)
           Z = INSTR(Z, CL$, "   ")
        END IF
        IF RTRIM$(NameDATA.BPLACE) <> "" THEN
           MID$(CL$, Z, 255) = RFLBL$(9): Z = Z + RFLTH(9)
           MID$(CL$, Z + 1, 255) = NameDATA.BPLACE: Z = INSTR(Z, CL$, "   ")
        END IF
     END IF
     IF (LTRIM$(NameDATA.DDATE) <> "") OR (RTRIM$(NameDATA.DPLACE) <> "") THEN
        MID$(CL$, Z, 255) = RFLBL$(10): Z = Z + RFLTH(10)
        IF RTRIM$(NameDATA.DDATE) <> "" THEN
           MID$(CL$, Z + 1, 12) = LTRIM$(NameDATA.DDATE)
           Z = INSTR(Z, CL$, "   ")
        END IF
        IF RTRIM$(NameDATA.DPLACE) <> "" THEN
           MID$(CL$, Z, 255) = RFLBL$(11): Z = Z + RFLTH(11)
           MID$(CL$, Z + 1, 255) = LTRIM$(NameDATA.DPLACE): Z = INSTR(Z, CL$, "   ")
        END IF
     END IF
     IF NameDATA.NUMCH <> SPACE$(3) THEN
        MID$(CL$, Z, 1) = ","
        xNCH = VAL(NameDATA.NUMCH)
        MID$(CL$, Z + 1, 3) = STR$(xNCH)
        Z = INSTR(Z + 2, CL$, "   ") + 1
        IF xNCH > 1 THEN
           MID$(CL$, Z, 255) = RFLBL$(23)
           ELSE
           MID$(CL$, Z, 255) = RFLBL$(67)
        END IF
        Z = INSTR(Z, CL$, "   ")
     END IF
     PAGE.XID = CC: PAGE.XTYPE = 4: CALL PutDATA
     IF RTRIM$(MarriageDATA.SPNAME) <> "" THEN
        MID$(CL$, 1, 255) = RFLBL$(25): Z = 1 + RFLTH(25)
        IF RptOPTION.PRID THEN
           MID$(CL$, Z + 1, 255) = "(" + RFLBL$(2): Z = Z + RFLTH(2) + 2
           MID$(CL$, Z, 6) = LTRIM$(MarriageDATA.SPID) + ")"
           Z = INSTR(Z, CL$, "   ")
        END IF
        MID$(CL$, Z + 1, 52) = MarriageDATA.SPNAME: Z = INSTR(Z, CL$, "   ")
        IF RTRIM$(MarriageDATA.MDATE) <> "" THEN
           MID$(CL$, Z, 255) = RFLBL$(30): Z = Z + RFLTH(30)
           MID$(CL$, Z + 1, 11) = MarriageDATA.MDATE
           Z = INSTR(Z, CL$, "   ")
           IF RTRIM$(MarriageDATA.MPLACE) <> "" THEN
              MID$(CL$, Z, 255) = RFLBL$(15): Z = Z + RFLTH(15)
              MID$(CL$, Z + 1, 255) = MarriageDATA.MPLACE
           END IF
        END IF
        PAGE.XID = CSPID: PAGE.XTYPE = 9: CALL PutDATA
     END IF
     IF MID$(IA$, 12, 1) <> "2" THEN
        PAGE.COFST = 8: GOSUB PrintCOM
     END IF
     CALL PrintLINE
     IF CHILDSEP AND (NCH > 0) THEN CALL WriteLINE
     RETURN

PrintEventH:
     IF BLNKBTWN THEN CALL WriteLINE
     LSET PLINE$ = LINESEP2$: CALL PrintLINE: EVNTCNT = 0
     IF A <> 27 THEN LSET PLINE$ = RFLBL$(42): CALL PrintLINE
     RETURN

PrintEVENT:
     IF EVNTCNT > 0 THEN LSET PLINE$ = LINESEP3$: CALL PrintLINE
     EVNTCNT = EVNTCNT + 1: LSET EventDATA = BEDATA: ComPTR = 0
     IF NOT BF THEN
        CALL FF3GetRec(EVENTPTR): ComPTR = FF3EVENT.COM
        EventDATA.CODE = FF3EVENT.CODE
        M = VAL(REVENT$(0)): X = 0: FOUND = 0
        WHILE (X < M) AND (NOT FOUND)
           X = X + 1
           FOUND = (RTRIM$(EventDATA.CODE) = RTRIM$(MID$(REVENT$(X), 1, 8)))
        WEND
        IF FOUND THEN
           EventDATA.DESC = MID$(REVENT$(X), 9, 15)
           ELSE
           EventDATA.DESC = EventDATA.CODE
        END IF
        EventDATA.DATE = XRPTDate$(FF3EVENT.D1MD, FF3EVENT.D1Y, DS)
        EventDATA.DSTAT = RTRIM$(FF3EVENT.DSTAT)
        IF FF3EVENT.LPLNM <> 0 THEN
           CALL FF3GetRec(FF3EVENT.LPLNM)
           LSET EventDATA.PLACE = FF3PLACE2.PLACE
           ELSE
           LSET EventDATA.PLACE = FF3EVENT.PLACE
        END IF
        X = ASC(FF3EVENT.IMPORT)
        IF X > 0 THEN EventDATA.IMPORT = LTRIM$(STR$(X))
        IF ULALL THEN CALL ULFLDS(5)
     END IF

     ' Print Information Formatted into EventDATA
     PAGE.LOFST = 2: PAGE.COFST = 4: X = PAGE.LOFST
     MID$(PLINE$, X, 255) = RFLBL$(43): X = X + RFLTH(43)
     MID$(PLINE$, X + 1, 15) = EventDATA.DESC
     IF RptOPTION.FREEFORM THEN X = INSTR(X, PLINE$, "   ") ELSE X = X + 17
     IF (RTRIM$(EventDATA.DATE) <> "") OR (NOT RptOPTION.FREEFORM) THEN
        MID$(PLINE$, X, 255) = RFLBL$(44): X = X + RFLTH(44)
        MID$(PLINE$, X + 1, 12) = EventDATA.DATE
        IF RptOPTION.FREEFORM THEN X = INSTR(X, PLINE$, "   ") ELSE X = X + 12
        IF EventDATA.DSTAT <> "" THEN
           MID$(PLINE$, X, 3) = "(" + EventDATA.DSTAT + ")"
           IF RptOPTION.FREEFORM THEN X = X + 3
        END IF
        IF NOT RptOPTION.FREEFORM THEN X = X + 5
     END IF
     IF (RTRIM$(EventDATA.PLACE) <> "") OR (NOT RptOPTION.FREEFORM) THEN
        MID$(PLINE$, X, 255) = RFLBL$(45): X = X + RFLTH(45)
        MID$(PLINE$, X + 1, 255) = EventDATA.PLACE
        IF RptOPTION.FREEFORM THEN
           X = INSTR(X, PLINE$, "   ")
           ELSE
           X = X + 24 - 19 * ENV.LPNSupport
        END IF
     END IF
     IF (RTRIM$(EventDATA.IMPORT) <> "") OR (NOT RptOPTION.FREEFORM) THEN
        MID$(PLINE$, X, 255) = RFLBL$(46): X = X + RFLTH(46)
        MID$(PLINE$, X + 1, 3) = EventDATA.IMPORT
     END IF
     IF RptOPTION.FREEFORM THEN
        LSET CL$ = PLINE$: LSET PLINE$ = "": CALL PutDATA
        ELSE
        CALL PrintLINE
     END IF
     IF MID$(IA$, 10, 1) = "1" THEN
        RI = VAL(MID$(MA$, 19, 2)): GOSUB PrintCOM
     END IF
     RETURN

PrintRESH:
     IF BLNKBTWN THEN CALL WriteLINE
     LSET PLINE$ = LINESEP2$: CALL PrintLINE
     IF A <> 27 THEN MID$(PLINE$, 1, 255) = RFLBL$(31): CALL PrintLINE
     NBL = -1
     RETURN

PrintRESI: IF NOT NBL THEN CALL WriteLINE: IF A = 27 THEN 2595
     LSET ADRDATA$ = STRING$(157, FC): ComPTR = 0: NBL = 0
     IF NOT BF THEN
        CALL FF2GetRec(APTR): ComPTR = FF2PFX.COM
        IF (FF2PFX.D1MD <> 0) OR (FF2PFX.D1Y <> 0) THEN
           MID$(ADRDATA$, 1, 11) = XRPTDate$(FF2PFX.D1MD, FF2PFX.D1Y, DS)
        END IF
        IF (FF2PFX.D2MD <> 0) OR (FF2PFX.D2Y <> 0) THEN
           MID$(ADRDATA$, 12, 11) = XRPTDate$(FF2PFX.D2MD, FF2PFX.D2Y, DS)
        END IF
        IF AFMT = 0 THEN
           IF (FF2SREC.FON1 <> 0) OR (FF2SREC.FON2 <> 0) OR (FF2SREC.FON3 <> 0) THEN
              MID$(ADRDATA$, 23, 14) = "(" + ZFILL$(FF2SREC.FON1, 3) + ")" + ZFILL$(FF2SREC.FON2, 3) + "-" + ZFILL$(FF2SREC.FON3, 4) + " "
           END IF
           MID$(ADRDATA$, 37, 30) = FF2SREC.LINE1
           MID$(ADRDATA$, 67, 30) = FF2SREC.LINE2
           MID$(ADRDATA$, 97, 15) = FF2SREC.CITY
           MID$(ADRDATA$, 117, 4) = FF2SREC.STATE
           MID$(ADRDATA$, 129, 5) = FF2SREC.ZIP
           ELSE
           IF (ASC(FF2LREC.PHONE) <> 0) AND (RTRIM$(FF2LREC.PHONE) <> "") THEN
              IF (LEFT$(FF2LREC.PHONE, 12) <> "000-000-0000") AND (LEFT$(FF2LREC.PHONE, 13) <> "(000)000-0000") AND (LEFT$(FF2LREC.PHONE, 14) <> "(000) 000-0000") THEN
                 MID$(ADRDATA$, 23, 14) = FF2LREC.PHONE
              END IF
           END IF
           MID$(ADRDATA$, 37, 30) = FF2LREC.LINE1
           MID$(ADRDATA$, 67, 30) = FF2LREC.LINE2
           MID$(ADRDATA$, 97, 20) = FF2LREC.CITY
           MID$(ADRDATA$, 117, 12) = FF2LREC.STATE
           MID$(ADRDATA$, 129, 12) = FF2LREC.ZIP
           MID$(ADRDATA$, 141, 16) = FF2LREC.COUNTRY
        END IF
        IF ULALL THEN CALL ULFILL(ADRDATA$)
     END IF

     ' Print Information Formatted into ADRDATA$
     PAGE.LOFST = 2: PAGE.COFST = 4: X = 2: LSET CL$ = ""
     IF (LTRIM$(MID$(ADRDATA$, 1, 11)) <> "") OR (NOT RptOPTION.FREEFORM) THEN
        MID$(CL$, X, 255) = RFLBL$(32): X = X + RFLTH(32)
        MID$(CL$, X + 1, 11) = MID$(ADRDATA$, 1, 11)
        IF RptOPTION.FREEFORM THEN X = INSTR(X, CL$, "   ") ELSE X = X + 13
     END IF
     IF (LTRIM$(MID$(ADRDATA$, 12, 11)) <> "") OR (NOT RptOPTION.FREEFORM) THEN
        MID$(CL$, X, 255) = RFLBL$(33): X = X + RFLTH(33)
        MID$(CL$, X + 1, 11) = MID$(ADRDATA$, 12, 11)
        IF RptOPTION.FREEFORM THEN X = INSTR(X, CL$, "   ") ELSE X = X + 14
     END IF
     IF (LTRIM$(MID$(ADRDATA$, 23, 14)) <> "") OR (NOT RptOPTION.FREEFORM) THEN
        MID$(CL$, X, 255) = RFLBL$(34): X = X + RFLTH(34)
        MID$(CL$, X + 1, 14) = LTRIM$((MID$(ADRDATA$, 23, 14)))
     END IF
     IF NOT RptOPTION.FREEFORM THEN
        LSET PLINE$ = CL$: LSET CL$ = ""
        CALL PrintLINE: X = 4: IF A = 27 THEN 2595
        ELSE
        X = INSTR(X, CL$, "   ")
     END IF
     IF LTRIM$(MID$(ADRDATA$, 37, 30)) <> "" THEN
        MID$(CL$, X, 255) = RFLBL$(35): X = X + RFLTH(35)
        MID$(CL$, X + 1, 30) = LTRIM$(MID$(ADRDATA$, 37, 30))
     END IF
     X2 = X: X = INSTR(X, CL$, "   ")
     IF LTRIM$(MID$(ADRDATA$, 67, 30)) <> "" THEN
        IF (X > X2) THEN MID$(CL$, X, 1) = ",": X = X + 2
        MID$(CL$, X, 30) = LTRIM$(MID$(ADRDATA$, 67, 30))
     END IF
     IF (NOT RptOPTION.FREEFORM) THEN
        IF AFMT = 0 THEN
           X = X2 + 64
           IF EuroADR THEN
              MID$(CL$, X, 255) = RFLBL$(38): X = X + 1 + RFLTH(38)
              MID$(CL$, X, 5) = MID$(ADRDATA$, 129, 5): X = X + 7
           END IF
           MID$(CL$, X, 255) = RFLBL$(36): X = X + 1 + RFLTH(36)
           MID$(CL$, X, 15) = MID$(ADRDATA$, 97, 15): X = X + 17
           MID$(CL$, X, 255) = RFLBL$(37): X = X + 1 + RFLTH(37)
           MID$(CL$, X, 4) = MID$(ADRDATA$, 117, 4): X = X + 6
           IF NOT EuroADR THEN
              MID$(CL$, X, 255) = RFLBL$(38): X = X + 1 + RFLTH(38)
              MID$(CL$, X, 5) = MID$(ADRDATA$, 129, 5)
           END IF
           ELSE
           LSET PLINE$ = CL$: LSET CL$ = "": CALL PrintLINE: IF A = 27 THEN 2595
           X = 4
           IF EuroADR THEN
              MID$(CL$, X, 255) = RFLBL$(40): X = X + 1 + RFLTH(40)
              MID$(CL$, X, 12) = MID$(ADRDATA$, 129, 12): X = X + 14
           END IF
           MID$(CL$, X, 255) = RFLBL$(36): X = X + 1 + RFLTH(36)
           MID$(CL$, X, 20) = MID$(ADRDATA$, 97, 20): X = X + 22
           MID$(CL$, X, 255) = RFLBL$(39): X = X + 1 + RFLTH(39)
           MID$(CL$, X, 12) = MID$(ADRDATA$, 117, 12): X = X + 14
           IF NOT EuroADR THEN
              MID$(CL$, X, 255) = RFLBL$(40): X = X + 1 + RFLTH(40)
              MID$(CL$, X, 12) = MID$(ADRDATA$, 129, 12): X = X + 14
           END IF
           MID$(CL$, X, 255) = RFLBL$(41): X = X + 1 + RFLTH(41)
           MID$(CL$, X, 16) = MID$(ADRDATA$, 141, 16)
        END IF
        GOTO 2560
     END IF

     'Free Form Address for Merged Reports
     IF RptOPTION.FREEFORM THEN
        CALL PutDATA: X = 1
        ELSE
        IF AFMT = 1 THEN
           LSET PLINE$ = CL$: CALL PrintLINE: LSET CL$ = ""
           IF A = 27 THEN 2595 ELSE X = 11
        END IF
     END IF
     IF EuroADR THEN
        X2 = 129: ZL = 5 - 7 * (AFMT = 1): EuroZIP = 0
        IF LTRIM$(MID$(ADRDATA$, X2, ZL)) <> "" THEN
           MID$(CL$, X, 1) = ",": X = X + 2: EuroZIP = -1
           MID$(CL$, X, ZL) = LTRIM$(MID$(ADRDATA$, X2, ZL))
           X = INSTR(X, CL$, "   ")
        END IF
     END IF
     X2 = 97
     ZL = 15 - 5 * (AFMT = 1)
     IF LTRIM$(MID$(ADRDATA$, X2, ZL)) <> "" THEN
        IF NOT EuroZIP THEN MID$(CL$, X, 1) = ",": X = X + 1
        X = X + 1
        MID$(CL$, X, ZL) = LTRIM$(MID$(ADRDATA$, 97, ZL))
        X = INSTR(X, CL$, "   ")
     END IF
     X2 = 117: ZL = 4 - 8 * (AFMT = 1)
     IF LTRIM$(MID$(ADRDATA$, X2, ZL)) <> "" THEN
        MID$(CL$, X, 1) = ","
        MID$(CL$, X + 2, ZL) = LTRIM$(MID$(ADRDATA$, X2, ZL))
        X = INSTR(X + 2, CL$, "   ")
     END IF
     IF NOT EuroADR THEN
        X2 = 129: ZL = 5 - 7 * (AFMT = 1)
        IF LTRIM$(MID$(ADRDATA$, X2, ZL)) <> "" THEN
           MID$(CL$, X + 2, ZL) = LTRIM$(MID$(ADRDATA$, X2, ZL))
           X = INSTR(X + 2, CL$, "   ")
        END IF
     END IF
     IF AFMT = 1 THEN
        X2 = 141: ZL = 16
        IF LTRIM$(MID$(ADRDATA$, X2, ZL)) <> "" THEN
           MID$(CL$, X, 1) = ","
           MID$(CL$, X + 2, ZL) = LTRIM$(MID$(ADRDATA$, X2, ZL))
        END IF
     END IF
2560 IF RptOPTION.FREEFORM THEN
        CALL PutDATA
        ELSE
        LSET PLINE$ = CL$: LSET CL$ = ""
        CALL PrintLINE: IF A = 27 THEN 2595
     END IF
     IF MID$(IA$, 10, 1) = "1" THEN
        RI = VAL(MID$(MA$, 19, 2)): GOSUB PrintCOM
     END IF
     CALL PrintLINE
2595 RETURN

PrintCOM: ' Process COMMENTS
     IF RptOPTION.FREEFORM THEN RI = 0: PAGE.CINDENT = 0
     COMBGN = 0: PAGE.LWIDTH = CWIDTH: CWIDTH = 0
     IF PAGE.LWIDTH = 0 THEN PAGE.LWIDTH = PAGE.BWIDTH - PAGE.CINDENT
     IF ((ComPTR <> 0) OR (COMPTR2 <> 0)) THEN
        IF ComPTR <> 0 THEN CALL ComPRINT(ComPTR, COMBGN): ComPTR = 0
        IF COMPTR2 <> 0 THEN CALL ComPRINT(COMPTR2, COMBGN): COMPTR2 = 0
        ELSE
        IF RI > 0 THEN
           COFST = PAGE.COFST
           IF NOT COMBGN THEN LSET CL$ = RFLBL$(1): CALL PutDATA
           PAGE.COFST = PAGE.COFST + PAGE.CINDENT
           WHILE RI > 0
              LSET CL$ = STRING$(PAGE.LWIDTH - PAGE.LOFST, FC)
              CALL PutDATA: CALL PrintLINE
              RI = RI - 1
           WEND
           PAGE.COFST = COFST
        END IF
     END IF
     CALL PrintLINE
2695 RETURN

2700 ' Print OTHER Info
     OD = VAL(MID$(MA$, 7 + 2 * OT, 2))
     IF ((OA(OT) = 0) AND (OD = 0)) OR (MID$(IA$, OT + 4, 1) <> "1") THEN 2795
     IF BLNKBTWN THEN CALL WriteLINE
     LSET PLINE$ = LINESEP2$: CALL PrintLINE: IF A = 27 THEN 2790
     SELECT CASE OT
       CASE 1: X = 54
       CASE 2: X = 49
       CASE 3: X = 62
       CASE 4: X = 59
     END SELECT
     MID$(PLINE$, 1, 255) = RFLBL$(X)
     CALL PrintLINE: IF A = 27 THEN 2790
     Y = 0
2720 IF Y > 0 THEN LSET PLINE$ = LINESEP3$: CALL WriteLINE
     ComPTR = 0: APTR = 0: LSET OTHDATA$ = STRING$(LEN(OTHDATA$), FC)
     IF OA(OT) <> 0 THEN
        CALL FF3GetRec(OA(OT))
        ComPTR = FF3PNTRS.COM: APTR = FF3PNTRS.LOC
        IF (FF3PNTRS.D1MD <> 0) OR (FF3PNTRS.D1Y <> 0) THEN
           MID$(OTHDATA$, 1, 11) = XRPTDate$(FF3PNTRS.D1MD, FF3PNTRS.D1Y, DS)
        END IF
        IF (FF3PNTRS.D2MD <> 0) OR (FF3PNTRS.D2Y <> 0) THEN
           MID$(OTHDATA$, 12, 11) = XRPTDate$(FF3PNTRS.D2MD, FF3PNTRS.D2Y, DS)
        END IF
        SELECT CASE OT
           CASE 1: IF RTRIM$(FF3HEALTH.HLTHDIAG) <> "" THEN MID$(OTHDATA$, 23, 22) = FF3HEALTH.HLTHDIAG
                   IF RTRIM$(FF3HEALTH.HLTHSTAT) <> "" THEN MID$(OTHDATA$, 45, 8) = FF3HEALTH.HLTHSTAT
           CASE 2: IF RTRIM$(FF3EDUC.EDLEV) <> "" THEN MID$(OTHDATA$, 23, 8) = FF3EDUC.EDLEV
                   IF RTRIM$(FF3EDUC.EDDEG) <> "" THEN MID$(OTHDATA$, 31, 4) = FF3EDUC.EDDEG
                   IF RTRIM$(FF3EDUC.EDSUB1) <> "" THEN MID$(OTHDATA$, 35, 9) = FF3EDUC.EDSUB1
                   IF RTRIM$(FF3EDUC.EDSUB2) <> "" THEN MID$(OTHDATA$, 44, 9) = FF3EDUC.EDSUB2
           CASE 3: IF RTRIM$(FF3WORK.WORKTP) <> "" THEN MID$(OTHDATA$, 23, 30) = FF3WORK.WORKTP
           CASE 4: IF RTRIM$(FF3MIL.MILRANK) <> "" THEN MID$(OTHDATA$, 23, 22) = FF3MIL.MILRANK
                   IF RTRIM$(FF3MIL.MILSTAT) <> "" THEN MID$(OTHDATA$, 45, 8) = FF3MIL.MILSTAT
        END SELECT
     END IF
     IF ULALL THEN CALL ULFILL(OTHDATA$)

     ' Print Information Formatted in OTHDATA$
     PAGE.LOFST = 2: PAGE.COFST = 4: X = PAGE.LOFST
     IF OT > 1 THEN
        IF (RTRIM$(MID$(OTHDATA$, 1, 11)) <> "") OR (NOT RptOPTION.FREEFORM) THEN
           MID$(PLINE$, 2, 255) = RFLBL$(47): X = 3 + RFLTH(47)
           MID$(PLINE$, X, 11) = MID$(OTHDATA$, 1, 11)
           IF RptOPTION.FREEFORM THEN X = INSTR(X, PLINE$, "   ") ELSE X = X + 13
        END IF
        IF (RTRIM$(MID$(OTHDATA$, 12, 11)) <> "") OR (NOT RptOPTION.FREEFORM) THEN
           MID$(PLINE$, X, 255) = RFLBL$(48): X = X + 1 + RFLTH(48)
           MID$(PLINE$, X, 11) = MID$(OTHDATA$, 12, 11)
           IF RptOPTION.FREEFORM THEN X = INSTR(X, PLINE$, "   ") ELSE X = X + 13
        END IF
     END IF
     SELECT CASE OT
        CASE 1: ' Medical Information
                IF NOT RptOPTION.FREEFORM THEN
                   MID$(PLINE$, X, 255) = RFLBL$(55): X = X + RFLTH(55)
                   MID$(PLINE$, X + 1, 11) = MID$(OTHDATA$, 1, 11)
                   X = X + 13
                END IF
                IF (RTRIM$(MID$(OTHDATA$, 23, 11)) <> "") OR (NOT RptOPTION.FREEFORM) THEN
                   MID$(PLINE$, X, 255) = RFLBL$(56): X = X + RFLTH(56)
                   MID$(PLINE$, X + 1, 22) = LTRIM$(MID$(OTHDATA$, 23, 22))
                   IF RptOPTION.FREEFORM THEN X = INSTR(X, PLINE$, "   ") ELSE X = X + 24
                END IF
                IF RptOPTION.FREEFORM AND (RTRIM$(MID$(OTHDATA$, 1, 11)) <> "") THEN
                   MID$(PLINE$, X, 255) = RFLBL$(55): X = X + RFLTH(55)
                   MID$(PLINE$, X + 1, 11) = LTRIM$(MID$(OTHDATA$, 1, 11))
                   X = INSTR(X, PLINE$, "   ")
                END IF
                IF (RTRIM$(MID$(OTHDATA$, 45, 8)) <> "") OR (NOT RptOPTION.FREEFORM) THEN
                   MID$(PLINE$, X, 255) = RFLBL$(57): X = X + RFLTH(57)
                   MID$(PLINE$, X + 1, 8) = LTRIM$(MID$(OTHDATA$, 45, 8))
                   IF RptOPTION.FREEFORM THEN X = INSTR(X, PLINE$, "   ") ELSE X = X + 10
                END IF
                IF (RTRIM$(MID$(OTHDATA$, 12, 11)) <> "") OR (NOT RptOPTION.FREEFORM) THEN
                   MID$(PLINE$, X, 255) = RFLBL$(58): X = X + RFLTH(58)
                   MID$(PLINE$, X + 1, 11) = LTRIM$(MID$(OTHDATA$, 12, 11))
                END IF
        CASE 2: ' Education Information
                IF NOT RptOPTION.FREEFORM THEN
                   MID$(PLINE$, X, 255) = RFLBL$(50): X = X + RFLTH(50)
                   MID$(PLINE$, X + 1, 8) = MID$(OTHDATA$, 23, 8)
                   X = X + 10
                   MID$(PLINE$, X, 255) = RFLBL$(53): X = X + RFLTH(53)
                   MID$(PLINE$, X + 1, 4) = MID$(OTHDATA$, 31, 4)
                   X = X + 6
                END IF
                MID$(PLINE$, X, 255) = RFLBL$(51): X = X + RFLTH(51)
                IF (RTRIM$(MID$(OTHDATA$, 35, 9)) <> "") OR (NOT RptOPTION.FREEFORM) THEN
                   MID$(PLINE$, X + 1, 9) = LTRIM$(MID$(OTHDATA$, 35, 9))
                   IF RptOPTION.FREEFORM THEN X = INSTR(X, PLINE$, "   ") ELSE X = X + 11
                END IF
                IF (RTRIM$(MID$(OTHDATA$, 44, 9)) <> "") OR (NOT RptOPTION.FREEFORM) THEN
                   MID$(PLINE$, X, 255) = RFLBL$(52): X = X + RFLTH(52)
                   MID$(PLINE$, X + 1, 9) = LTRIM$(MID$(OTHDATA$, 44, 9))
                   IF RptOPTION.FREEFORM THEN X = INSTR(X, PLINE$, "   ")
                END IF
                IF RptOPTION.FREEFORM THEN
                   IF RTRIM$(MID$(OTHDATA$, 23, 8)) <> "" THEN
                      MID$(PLINE$, X, 255) = RFLBL$(50): X = X + RFLTH(50)
                      MID$(PLINE$, X + 1, 8) = LTRIM$(MID$(OTHDATA$, 23, 8))
                      X = INSTR(X, PLINE$, "   ")
                      MID$(PLINE$, X, 255) = RFLBL$(66): X = X + RFLTH(66)
                   END IF
                   IF RTRIM$(MID$(OTHDATA$, 31, 4)) <> "" THEN
                      MID$(PLINE$, X, 255) = RFLBL$(53): X = X + RFLTH(53)
                      MID$(PLINE$, X + 1, 4) = LTRIM$(MID$(OTHDATA$, 31, 4))
                   END IF
                END IF
        CASE 3: ' Occupation Information
                IF RTRIM$(MID$(OTHDATA$, 23, 30)) <> "" OR (NOT RptOPTION.FREEFORM) THEN
                   MID$(PLINE$, X, 255) = RFLBL$(63): X = X + RFLTH(63)
                   MID$(PLINE$, X + 1, 30) = LTRIM$(MID$(OTHDATA$, 23, 30))
                END IF
        CASE 4: ' Military Information
                IF RTRIM$(MID$(OTHDATA$, 23, 22)) <> "" OR (NOT RptOPTION.FREEFORM) THEN
                   MID$(PLINE$, X, 255) = RFLBL$(60): X = X + RFLTH(60)
                   MID$(PLINE$, X + 1, 22) = LTRIM$(MID$(OTHDATA$, 23, 22))
                   IF RptOPTION.FREEFORM THEN X = INSTR(X, PLINE$, "   ") ELSE X = X + 24
                END IF
                IF (RTRIM$(MID$(OTHDATA$, 45, 8)) <> "") OR (NOT RptOPTION.FREEFORM) THEN
                   MID$(PLINE$, X, 255) = RFLBL$(61): X = X + RFLTH(61)
                   MID$(PLINE$, X + 1, 8) = LTRIM$(MID$(OTHDATA$, 45, 8))
                END IF
     END SELECT
     IF RptOPTION.FREEFORM THEN
        LSET CL$ = PLINE$: LSET PLINE$ = "": CALL PutDATA
        ELSE
        CALL PrintLINE: IF A = 27 THEN 2790
     END IF
     Y = Y + 1: IF OA(OT) <> 0 THEN OA(OT) = FF3PNTRS.NXT
     IF MID$(IA$, 10, 1) = "1" THEN
        RI = VAL(MID$(MA$, 19, 2))
        GOSUB PrintCOM: IF A = 27 THEN 2790
     END IF
     IF MID$(IA$, 9, 1) = "1" THEN
        RH = VAL(MID$(MA$, 17, 2))
        IF (APTR <> 0) OR (RH > 0) THEN
           AX = 0
           WHILE APTR <> 0
              AX = AX + 1: GOSUB PrintRESI
              IF A = 27 THEN APTR = 0: RH = 0:  ELSE APTR = FF2PFX.NXT
           WEND
           IF MID$(MA$, 17, 1) <> "+" THEN RH = RH - AX
           IF RH > 0 THEN
              BFH = BF: BF = -1
              WHILE RH > 0
                 GOSUB PrintRESI: RH = RH - 1
                 IF A = 27 THEN RH = 0
              WEND
              BF = BFH
           END IF
           IF A = 27 THEN 2790
        END IF
     END IF
     IF OA(OT) <> 0 THEN 2720
     IF Y < OD THEN 2720 ELSE 2795
2790 OT = 99
2795 RETURN

62300 ' Format & Print Line showing RELATION & LINEAGE
      X = RELATE(RID)
      IF X = 0 THEN 62390
      AGL = -(X \ 256): DGL = X AND 127: RGL = AGL + DGL
      ExtAdopt = ((X AND 128) > 0)
      MID$(PLINE$, 1, 255) = RFLBL$(64): X = 2 + RFLTH(64)
      IF ExtAdopt THEN MID$(PLINE$, X, 1) = "*": X = X + 1
      MID$(PLINE$, X, 12) = Relation$("M", FF1REC.SEX, AGL, RGL, RRELCD$())
      X = X + 15
      MID$(PLINE$, X, 255) = RFLBL$(65): X = X + RFLTH(65)
      CALL RWRKGet("RELREC", RC, RELPTR&(RID))
      IF ASC(WFHDR.AMAXLV) > 0 THEN
         MID$(PLINE$, X, 255) = STR$(WFANCREC.LINEAGE)
         X = INSTR(X, PLINE$, "   ")
         IF NOT ExtAdoptRule THEN WFANCREC.PSTAT = WFANCREC.PSTAT AND 3
         IF WFANCREC.PSTAT > 0 THEN
            MID$(PLINE$, X, 1) = MID$("*? *", WFANCREC.PSTAT, 1)
            X = X + 1
         END IF
      END IF
      X = X + 1
      IF DGL = 0 THEN GOTO 62390
      Y = DGL
      WHILE Y > 0
         GBL(Y) = ASC(WFRELREC.CHNUM) + 256 * (WFRELREC.CSTAT AND 3)
         GET #5, WFRELREC.PARPTR + 1, WFRELREC
         Y = Y - 1
      WEND
      MID$(PLINE$, X, 1) = "(": X = X + 1
      FOR Y = 1 TO DGL
         MID$(PLINE$, X, 2) = FNX$(GBL(Y) AND 255): X = INSTR(X, PLINE$, " ")
         Z = GBL(Y) \ 256
         IF Z > 0 THEN MID$(PLINE$, X, 1) = MID$("*?", Z, 1): X = X + 1
         IF Y < DGL THEN MID$(PLINE$, X, 1) = "," ELSE MID$(PLINE$, X, 1) = ")"
         X = X + 1: IF X > PAGE.LWIDTH - 4 THEN CALL PrintLINE: X = 3
      NEXT Y
62390 CALL PrintLINE
      RETURN
END SUB

REM $STATIC
SUB GetNextMARRIAGE
SHARED SPRESEQ, MLIST(), SPMCNT, SC, RID
   IF SPRESEQ THEN
      IF SPMCNT > 0 THEN SPMCNT = SPMCNT - 1
      SC = MLIST(SPMCNT)
      ELSE
      IF RID = FF3SPOUSE.SP1ID THEN
         SC = FF3SPOUSE.SP1NXT
         ELSE
         IF RID = FF3SPOUSE.SP2ID THEN
            SC = FF3SPOUSE.SP2NXT
            ELSE
            SC = 0
         END IF
      END IF
   END IF

END SUB

REM $DYNAMIC
SUB SelectID
SHARED BRID, HIID, RID, NAME$, SCNT, STYPE, STYPE$(), IDLIST()
   CALL FmtFIND("GRPRS002"): FMT002 = CURFMT
   CALL FmtFindFLD("ID"): IF FOUND THEN FFRID = FFLD.UTXT
   CALL FmtFindFLD("NAME"): IF FOUND THEN FFNAME = FFLD.UTXT
   CALL FmtFindFLD("STYP"): IF FOUND THEN FFSTYPE = FFLD.UTXT
   CALL FmtFindFLD("SCNT"): IF FOUND THEN FFSCNT = FFLD.UTXT
   CALL FmtFIND("GRPRS010"): FMT010 = CURFMT
   CALL FmtFIND("XXXXS009"): FMT009 = CURFMT

90 ' Change Select Type or ID #'s
   X = 0: CALL FAMOpen("READ"): IF NOT OKAY THEN 195
   CALL PutMSG("GRPRM014")
   WHILE X = 0
      CALL GetRESP: IF A = 27 THEN 195 ELSE X = VAL(A$)
      IF (X < 0) OR (X > 5) THEN
         CALL ErrBEEP(0): X = 0
         ELSE
         CALL PutMSG("")
         IF (A <> 32) THEN
            SCNT = 0: ID = 0: LOID = 32767: HIID = 0
            STYPE = X: RID = 0
         END IF
         IF STYPE = 0 THEN STYPE = 2
         X = STYPE
      END IF
   WEND
   GOSUB 185
   IF SCNT > 1 THEN GOTO 105

   ' Get Single ID to Process
   IF STYPE = 1 THEN
      CALL PutMSG("GRPRM015")
      SCNT = 1: GOSUB 180
      IF RID = 0 THEN SCNT = 0: GOSUB 185
      GOTO 195
   END IF

   REDIM IDLIST(FF1HDR.MAXID)

   ' Process REL Work File
   IF STYPE = 3 THEN GOSUB 160: IF NOT WFOKAY THEN GOTO 90
   CALL FAMOpen("READ")

   ' Process SELECT File
   IF STYPE = 4 THEN
      CALL SWRKBldList(IDLIST(), X, 0): IF NOT OKAY THEN GOTO 90
      CALL FAMOpen("READ")
      HIID = SLCTHdr.HIID: LOID = SLCTHdr.LOID
      SCNT = SLCTHdr.SLCTTOT
      RID = LOID
   END IF

   ' Process Full File
   IF STYPE = 5 THEN
      LOID = 1: HIID = FF1HDR.MAXID
      I = 1: WHILE I <= FF1HDR.MAXID: IDLIST(I) = 1: I = I + 1: WEND
      RID = 1: SCNT = FF1HDR.MAXID
   END IF

105 ' Get List of ID's to process
    WIN.MFMT = FMT010: CALL MsgFORMAT

110 ' Get Next ID/Response
    GOSUB 185
    ID = RID: GOSUB 180
    IF (RID = 0) AND (RA <> 60) THEN 195
    SELECT CASE RA
       CASE 27:  GOTO 195
       CASE 59:  GOTO 195
       CASE 60:  GOSUB 130: GOTO 105
       CASE 61:  SCNT = SCNT - 1: IDLIST(RID) = 0
                 IF RID = LOID THEN
                    IF SCNT > 0 THEN GOSUB 120: LOID = RID ELSE LOID = 32767
                 END IF
                 IF RID = HIID THEN
                    IF SCNT > 0 THEN GOSUB 125: HIID = RID ELSE HIID = 0
                 END IF
                 IF (LOID < RID) AND (RID < HIID) THEN GOSUB 120
                 IF SCNT = 0 THEN RID = 0
       CASE 68:  SRCNT = 0: CALL SWRKCreate(IDLIST(), FF1HDR.MAXID, 0)
                 CALL FAMOpen("READ"): GOTO 105
       CASE 73:  GOSUB 125
       CASE 81:  GOSUB 120
       CASE 118: RID = HIID
       CASE 132: RID = LOID
       CASE ELSE: CALL ErrBEEP(0): GOTO 110
    END SELECT
    IDLIST(0) = LOID
    GOTO 110

120 ' PGDN List of ID's
    IF RID < HIID THEN
       RID = RID + 1
       WHILE (RID < HIID) AND (IDLIST(RID) = 0): RID = RID + 1: WEND
       ELSE
       CALL ErrBEEP(0)
    END IF
    RETURN

125 ' PGUP List of ID's
    IF RID > LOID THEN
       RID = RID - 1
       WHILE (RID > LOID) AND (IDLIST(RID) = 0): RID = RID - 1: WEND
       ELSE
       CALL ErrBEEP(0)
    END IF
    RETURN

130 ' Enter RANGE of ID's
    WIN.MFMT = FMT009: CALL MsgFORMAT
    CALL FmtFindFLD("ID1"): ID1 = VAL(UTXT$(FFLD.UTXT))
    CALL FmtFindFLD("ID2"): ID2 = VAL(UTXT$(FFLD.UTXT))
    IF ID2 < ID1 THEN SWAP ID1, ID2
    FOR I = ID1 TO ID2
       IF (I > 0) AND (I <= FF1HDR.MAXID) THEN
          IF IDLIST(I) = 0 THEN SCNT = SCNT + 1: IDLIST(I) = 1
       END IF
    NEXT I
    IF (LOID > ID1) THEN LOID = ID1
    IF (HIID < ID2) THEN HIID = ID2
    RID = ID2: IDLIST(0) = LOID
    RETURN

160 ' Load IDLIST from Relative WRK File
    CALL RWRKOpen("READ", 0, 0): WFOKAY = OKAY
    IF OKAY THEN
       CALL RWRKGet("BASE", RC, PTR&)
       CALL PutMSG("XXXXM064")
       WHILE (0 < RC) AND (RC < 9)
          IF WFRELREC.ID > 0 THEN
             RID = WFRELREC.ID
             IF IDLIST(RID) = 0 THEN
                SCNT = SCNT + 1: IDLIST(RID) = 1
                IF RID < LOID THEN LOID = RID
                IF HIID < RID THEN HIID = RID
             END IF
          END IF
          CALL RWRKGet("WRKREC", RC, PTR&)
       WEND
       IDLIST(0) = LOID
       RID = LOID
    END IF
    CALL FamCLOSE
    RETURN

180 ' Get Next ID Number to Process
    CURFMT = FMT002: FMT = FMTTAB(FMT002): CALL WinUPDATE
    RID = VAL(UTXT$(FFRID)): RA = A
    IF (RA = 27) THEN RID = ID: GOTO 185
    IF ((RA <> 13) AND (STYPE > 1) AND (RID = ID)) THEN GOTO 185
    IF (RID = 0) OR (RID > FF1HDR.MAXID) THEN CALL ErrBEEP(0): GOTO 180
    IF (STYPE > 1) THEN
       IF IDLIST(RID) = 0 THEN IDLIST(RID) = 1: SCNT = SCNT + 1
    END IF
    IF RID < LOID THEN LOID = RID
    IF RID > HIID THEN HIID = RID
    GOSUB 185: ID = RID
    IF STYPE > 1 THEN 180
    RETURN

185 ' Display Information
    LSET NAME$ = ""
    IF (RID > 0) AND (RID <= FF1HDR.MAXID) THEN
       CALL FF1GetRec(RID): IF OKAY THEN LSET NAME$ = FMTNAME$(1)
    END IF
    IF FFRID > 0 THEN RSET UTXT$(FFRID) = FNX$(RID)
    IF FFNAME > 0 THEN LSET UTXT$(FFNAME) = NAME$
    IF FFSTYPE > 0 THEN LSET UTXT$(FFSTYPE) = STYPE$(STYPE)
    IF FFSCNT > 0 THEN RSET UTXT$(FFSCNT) = FNX$(SCNT)
    CALL WinDisplayDATA(FMT002)
    RETURN

195 ' Return to Menu from F3
    CALL PutMSG("")
    CALL FamCLOSE
    ON ERROR GOTO 0
    EXIT SUB

END SUB

REM $STATIC
SUB ULFILL (X$)
    Y$ = CHR$(95): Z = INSTR(X$, " ")
    WHILE Z > 0: MID$(X$, Z, 1) = Y$: Z = INSTR(X$, " "): WEND
END SUB

SUB ULFLDS (A)
SHARED NameDATA AS NDATA, FatherDATA AS PDATA, MotherDATA AS PDATA
SHARED MarriageDATA AS MDATA, EventDATA AS EDATA
   SELECT CASE A
      CASE 1: CALL ULFILL(NameDATA.RID)
              CALL ULFILL(NameDATA.PID)
              CALL ULFILL(NameDATA.NAME)
              CALL ULFILL(NameDATA.BDATE)
              CALL ULFILL(NameDATA.BPLACE)
              CALL ULFILL(NameDATA.DDATE)
              CALL ULFILL(NameDATA.DPLACE)
              CALL ULFILL(NameDATA.SEX)
              CALL ULFILL(NameDATA.NUMCH)
      CASE 2: CALL ULFILL(FatherDATA.NAME)
              CALL ULFILL(FatherDATA.BDATE)
              CALL ULFILL(FatherDATA.DDATE)
              CALL ULFILL(FatherDATA.AGE)
      CASE 3: CALL ULFILL(MotherDATA.NAME)
              CALL ULFILL(MotherDATA.BDATE)
              CALL ULFILL(MotherDATA.DDATE)
              CALL ULFILL(MotherDATA.AGE)
      CASE 4: CALL ULFILL(MarriageDATA.MDATE)
              CALL ULFILL(MarriageDATA.MPLACE)
              CALL ULFILL(MarriageDATA.DVDATE)
              CALL ULFILL(MarriageDATA.DVPLACE)
              CALL ULFILL(MarriageDATA.ANNIV)
              CALL ULFILL(MarriageDATA.SPID)
              CALL ULFILL(MarriageDATA.SPNAME)
      CASE 5: CALL ULFILL(EventDATA.CODE)
              CALL ULFILL(EventDATA.DATE)
              CALL ULFILL(EventDATA.DSTAT)
              CALL ULFILL(EventDATA.PLACE)
              CALL ULFILL(EventDATA.IMPORT)
   END SELECT
END SUB

