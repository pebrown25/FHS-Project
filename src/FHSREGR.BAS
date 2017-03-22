DECLARE SUB BldMLIST ()
DECLARE SUB BLDREGRDef ()
DECLARE SUB FREGReport ()
DECLARE SUB LINEAGE (OPT$)
DECLARE SUB GetNextMARRIAGE ()
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
DECLARE FUNCTION ROMAN$ (N%)
DECLARE FUNCTION UCX$ (X$)
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
   PN$ = " FAMILY HISTORY SYSTEM - FHSREGR - Family Register Report Program"
   CY$ = " (C) Copyright 1997 by Phillip E. Brown"
   IF CONFIG.BP = 0 THEN RUN "FHSINIT"
   REDIM WTYPE$(3)
   WTYPE$(1) = "*ANCESTOR": WTYPE$(2) = "*DESCNDNT": WTYPE$(3) = "*RELATIVE"
   CALL TableLOAD("WORKTYPE", T, 14)
   IF OKAY THEN
      X = TABLE(T).FT
      WHILE X > 0
        Z = VAL(MID$(TABDATA$(X), 1, 2))
        IF (0 < Z) AND (Z < 4) THEN
           WTYPE$(Z) = RTRIM$(MID$(TABDATA$(X), 3, 12))
        END IF
        X = TABDATACHN(X).FWD
      WEND
      CALL TableFREE("WORKTYPE")
   END IF
   MAXMarry = 10: REDIM MLIST(MAXMarry)
   NAME$ = SPACE$(52): IA$ = "": MA$ = "": READ$ = "READQ"
   DEF FNX$ (A) = LTRIM$(STR$(A))
   REDIM RELATE(1), RELPTR&(1), LINTAB$(99)
   REDIM IDLIST(1), NDXID(1)
   WTYPE = MO: BRID = 0: ShowPNAME = 0
   CALL WinINIT("LR1")
   CALL WinPREP(1): PGMRTRN$ = "FHSMENU"
   FUN = 0: RRPT$ = "REGR": CALL RptLOAD(RRPT$)
   CALL FmtFIND("REGRS001"): FMT001 = CURFMT: WIN.OFMT = FMT001
   CALL FmtFIND("REGRS002"): FMT002 = CURFMT

40 ' Format Display
   CALL WinCLR
   CALL WinDisplayLABELS(FMT001)

45 ' Refresh Option DISPLAY
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
    CALL FAMOpen(READ$): FAMOKAY = OKAY: READ$ = "READQ"
    IF OKAY THEN CALL RWRKOpen("READQ", WTYPE, 0)
    WFOKAY = OKAY
    IF WFOKAY THEN
       BRID = WFANCREC.ANCID: CALL FF1GetRec(BRID): BNAME$ = FMTNAME(1)
    END IF
    CALL FamCLOSE
101 ' Display Work File Information
    CURFMT = FMT002: FMT = FMTTAB(CURFMT)
    IF NOT WFOKAY THEN CALL FmtDREST: GOTO 102
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
102 CALL FmtFindFLD("TYPE")
    IF FOUND THEN LSET UTXT$(FFLD.UTXT) = WTYPE$(WTYPE)
    CALL WinDisplayDATA(0)
    IF FAMOKAY THEN
       CURFMT = FMT001: FMT = FMTTAB(CURFMT)
       CALL FmtFindFLD("NMCT")
       LSET UTXT$(FFLD.UTXT) = STR$(FF1HDR.MAXID)
       CALL WinDisplayDATA(FMT001)
    END IF

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
       CASE "F3": CALL RWRKBuild(WTYPE, BRID, 0, FMT002)
                  GOTO 100
       CASE "F4": GOTO 800
       CASE "F6": IF PGMRTRN$ = "FHSMENU" THEN 1000
                  IF PGMRTRN$ = "FHSRPTS" THEN 250
       CASE "F7": WTYPE = WTYPE + 1: IF WTYPE > 3 THEN WTYPE = 1
                  GOTO 100
       CASE "F9": GOTO 250
    END SELECT
    IF ENV.SCRNCLR THEN GOTO 40 ELSE GOTO 200

250 ' Return to FAMMENU.BAS
    CALL FamCLOSE
    REDIM RFLTH(1), RFLBL$(1)
    CALL PgmPREP(PGMRTRN$)
    IF (NOT OKAY) OR (A = 27) THEN LSET PGMRTRN$ = "FHSMENU": GOTO 40
    CALL RptFREE(RRPT$)
    CHAIN PGMRTRN$

300 ' Change FILE Parameters
    IF A = 94 THEN LSET PGMRTRN$ = "FHSFILE": GOTO 250
    CALL WinSWITCH(2)
    XSETUP = ENV.FSETUP
    CALL FDFSelect: READ$ = "READ"
    IF ENV.FSETUP <> XSETUP THEN
       RID = 0: LSET NAME$ = "": SCNT = 0
       FMT = FMTTAB(FMT002): CALL FmtDREST
       CALL WinFORMAT(0)
    END IF
    GOTO 45

400 ' Change PRINTER Parameters
    IF A = 95 THEN LSET PGMRTRN$ = "FHSPRTC": GOTO 250
    CALL WinSWITCH(2)
    CALL PDFSelect
    GOTO 45

800 ' Process Request to Update Options
    CALL WinSWITCH(2)
    CALL RptOptUPDATE(RRPT$)
    GOTO 45

1000 ' Process Print Requests
    IF NOT FAMOKAY THEN
       CALL PutMSG("REGRM001")
       GOTO 200
    END IF
    IF NOT WFOKAY THEN
       CALL PutMSG("REGRM002")
       GOTO 200
    END IF
    CALL FREGReport
    'IF ENV.SCRNCLR THEN 40
    'GOTO 200
    GOTO 40

REM $DYNAMIC
SUB BldMLIST
SHARED SPReSeq, SC, SPMCNT, MAXMarry, MLIST()
   SPMCNT = 0: SC = FF1REC.SPOUSE
   IF SPReSeq THEN
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

SUB BLDREGRDef
SHARED IA$, MA$, RRPT$
      PAGE.DFCNT = 67
      REDIM RFLTH(PAGE.DFCNT), RFLBL$(PAGE.DFCNT)

58500 ' Select Report Fields - Based Upon Report Options
      IA$ = SPACE$(12): MA$ = SPACE$(24): X = 0
      XI$ = "PRSP": GOSUB 58505
      XI$ = "PRCH": GOSUB 58505
      XI$ = "PRAD": GOSUB 58505
      XI$ = "PRCM": GOSUB 58505
      XI$ = "PRME": GOSUB 58505
      XI$ = "PRED": GOSUB 58505
      XI$ = "PRWK": GOSUB 58505
      XI$ = "PRML": GOSUB 58505
      XI$ = "PRXA": GOSUB 58505
      XI$ = "PRXC": GOSUB 58505
      XI$ = "PREV": GOSUB 58505
      XI$ = "CHCM": GOSUB 58505
      GOTO 58510

58505 ' Process Information Option
      X = X + 1: CALL RptOptGET(XI$, FX)
      IF FOUND THEN MID$(IA$, X, 1) = LTRIM$(STR$(FX))
      RETURN

58510 ' Create Report Definition Tables
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
              IF FOUND THEN
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
      OKAY = -1: RWIDTH = RDFREPORT.PWIDTH
      IF RWIDTH = 0 THEN RWIDTH = 125
      IF PAGE.BWIDTH = 0 THEN
         BWIDTH = PDFFORMS.FWIDTH - PDFFORMS.IMARGIN - PDFFORMS.OMARGIN
         ELSE
         BWIDTH = PAGE.BWIDTH
      END IF

58560 ' Return to Calling Program
      PAGE.BWIDTH = BWIDTH
END SUB

SUB FREGReport
SHARED WTYPE, RRPT$, IA$, MA$, NDXID(), RELATE(), RELPTR&()
SHARED STYPE, SCNT, IDLIST(), LINTAB$()
SHARED SPReSeq, MLIST(), SPMCNT, SC
SHARED ShowREL, ShowChNUM, ShowRegNUM, ShowPNAME, ExtAdoptRule, RomanCAPS
SHARED LinCNT, NMOFST, RID

    ON LOCAL ERROR GOTO 900
    'Process PRINT Requests
    CALL RptOPEN(RRPT$): IF NOT OKAY THEN GOTO 399
    CALL RptOptGET("BLKL", X): BLNKBTWN = (X = 1): CHILDSEP = BLNKBTWN
    CALL RptOptGET("EURA", X): EuroAdr = (X = 1)
    CALL RptOptGET("URMN", X): RomanCAPS = (X = 1)
    CALL BLDREGRDef
    REDIM OA(4)
    RID$ = SPACE$(5)
    XRID$ = "(" + RFLBL$(2) + RID$ + ")"
    XOPID$ = RFLBL$(3) + RID$ + ")"
    XSPID$ = "(" + RFLBL$(4) + RID$ + ")"
    CDATE$ = SPACE$(11)
    DIM BNDATA AS NDATA, BPDATA AS PDATA, BMDATA AS MDATA, BEDATA AS EDATA
    DIM NameDATA AS NDATA, FatherDATA AS PDATA, MotherDATA AS PDATA
    DIM MarriageDATA AS MDATA, EventDATA AS EDATA
    BNDATA.RID = SPACE$(LEN(BNDATA.RID))
    BNDATA.PID = SPACE$(LEN(BNDATA.PID))
    BNDATA.NAME = SPACE$(LEN(BNDATA.NAME))
    BNDATA.BDATE = SPACE$(LEN(BNDATA.BDATE))
    BNDATA.BPLACE = SPACE$(LEN(BNDATA.BPLACE))
    BNDATA.DDATE = SPACE$(LEN(BNDATA.DDATE))
    BNDATA.DPLACE = SPACE$(LEN(BNDATA.DPLACE))
    BNDATA.SEX = SPACE$(LEN(BNDATA.SEX))
    BNDATA.NUMCH = SPACE$(LEN(BNDATA.NUMCH))
    BPDATA.NAME = SPACE$(LEN(BPDATA.NAME))
    BPDATA.BDATE = SPACE$(LEN(BPDATA.BDATE))
    BPDATA.DDATE = SPACE$(LEN(BPDATA.DDATE))
    BPDATA.AGE = SPACE$(LEN(BPDATA.AGE))
    BMDATA.MDATE = SPACE$(LEN(BMDATA.MDATE))
    BMDATA.MPLACE = SPACE$(LEN(BMDATA.MPLACE))
    BMDATA.DVDATE = SPACE$(LEN(BMDATA.DVDATE))
    BMDATA.DVPLACE = SPACE$(LEN(BMDATA.DVPLACE))
    BMDATA.ANNIV = SPACE$(LEN(BMDATA.ANNIV))
    BMDATA.SPID = SPACE$(LEN(BMDATA.SPID))
    BMDATA.SPNAME = SPACE$(LEN(BMDATA.SPNAME))
    BEDATA.CODE = SPACE$(LEN(BEDATA.CODE))
    BEDATA.DESC = SPACE$(LEN(BEDATA.DESC))
    BEDATA.DATE = SPACE$(LEN(BEDATA.DATE))
    BEDATA.DSTAT = SPACE$(LEN(BEDATA.DSTAT))
    BEDATA.PLACE = SPACE$(LEN(BEDATA.PLACE))
    BEDATA.IMPORT = SPACE$(LEN(BEDATA.IMPORT))

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
    'EventDATA$ = SPACE$(79)
    '  1   CODE      8
    '  9   DATE     11
    ' 20   DSTAT     1
    ' 21   PLACE    41
    ' 43   IMPORT    3
    ADRDATA$ = SPACE$(157): MAXADR = 30
    OTHDATA$ = SPACE$(52)
    ' Build Separator Lines from Report Variables
    LINESEP1$ = SPACE$(PAGE.BWIDTH)
    CALL GetRVAR("LINESEP1", X)
    IF FOUND THEN
       LSET LINESEP1$ = RVAR$(X): CALL HTFBuild(LINESEP1$, PAGE.BWIDTH)
    END IF
    LINESEP2$ = SPACE$(PAGE.BWIDTH)
    CALL GetRVAR("LINESEP2", X)
    IF FOUND THEN
       LSET LINESEP2$ = RVAR$(X): CALL HTFBuild(LINESEP2$, PAGE.BWIDTH)
    END IF
    LINESEP3$ = SPACE$(PAGE.BWIDTH)
    CALL GetRVAR("LINESEP3", X)
    IF FOUND THEN
       LSET LINESEP3$ = RVAR$(X): CALL HTFBuild(LINESEP3$, PAGE.BWIDTH)
    END IF
    ' Build Child Information Header
    ChildHDR$ = RFLBL$(21)
    ' Get Misc Options
    ShowREL = RptOPTION.Relation
    CALL RptOptGET("PLIN", X): SHOWLIN = (X = 1)
    CALL RptOptGET("BYPS", X)
    BYPASS = ((X = 10) OR (X = 11)): BYPASS2 = ((X = 10) OR (X = 12))
    CALL RptOptGET("SP12", X): SPReSeq = (X = 1)
    CALL RptOptGET("NGSA", X): CountALL = (X = 1)
    CALL RptOptGET("NGSC", X): ShowChNUM = (X = 1)
    CALL RptOptGET("NGSR", X): ShowRegNUM = (X = 1)
    CALL RptOptGET("NGSN", X): ShowPNAME = (X = 1)
    CALL RptOptGET("NGHL", X): RelInHEAD = (X = 1)
    IF SHOWLIN AND NOT (ShowRegNUM OR ShowPNAME) THEN ShowChNUM = -1
    CALL GetRVAR("UNKNOWN", X): IF FOUND THEN Unknown$ = RVAR$(X)

    ' Prepare Register Tables
    CALL FAMOpen("READ"): IF NOT OKAY THEN 399
    A = 0
    BD$ = SPACE$(12): BPL$ = SPACE$(41)
    DD$ = BD$: DPL$ = BPL$: AGE$ = "   "
    DS = 0: FC = ASC(" ")
    STYPE = 3: SCNT = 0: HIID = 0: LOID = FF1HDR.MAXID
    CALL RWRKOpen("READQ", 0, 0): IF NOT OKAY THEN GOTO 395
    NDXRECS = 0: REDIM NDXID(FF1HDR.MAXID), IDLIST(FF1HDR.MAXID)
    'IF SHOWLIN THEN
       REDIM RELATE(FF1HDR.MAXID), RELPTR&(FF1HDR.MAXID)
       CALL RptOptGET("EXAD", X): ExtAdoptRule = (X = 1)
    'END IF
    CALL PutMSG("XXXXM064")
    ' The Next Statement Forces to Read Workfile by Generations
    WFHDR.WFRULES = CHR$((ASC(WFHDR.WFRULES) AND 252) + 1)
    CALL RwrkGET("BASE", RC, PTR&)
    WHILE (0 < RC) AND (RC < 9)
       IF WFRELREC.ID > 0 THEN
          RID = WFRELREC.ID
          IF IDLIST(RID) = 0 THEN
             SCNT = SCNT + 1: IDLIST(RID) = 1
             IF RID < LOID THEN LOID = RID
             IF HIID < RID THEN HIID = RID
          END IF
          IF IDLIST(RID) > 0 THEN
             NDXRECS = NDXRECS + 1: NDXID(NDXRECS) = RID
             IDLIST(RID) = -1
          END IF
          'IF SHOWLIN THEN
             IF WFGLREC.DGL = 0 THEN
                ExtAdopt = ((WFANCREC.PSTAT AND 5) > 0)
                ELSE
                ExtAdopt = ((WFRELREC.CSTAT AND 5) > 0)
             END IF
             X = RELATE(RID)
             IF (X = 0) OR (ExtAdoptRule AND ((X AND 128) > 0) AND (NOT ExtAdopt)) THEN
                X = 256 * (-WFGLREC.AGL) + WFGLREC.DGL
                IF ExtAdoptRule THEN X = X - 128 * ExtAdopt
                RELATE(RID) = X
                RELPTR&(RID) = PTR&
             END IF
          'END IF
       END IF
       CALL RwrkGET("WRKREC", RC, PTR&)
    WEND
    IDLIST(0) = LOID
    RID = LOID
    IF SCNT = 0 THEN
       CALL PutMSG("REGRM003"): GOTO 399
    END IF
    REDIM PRESERVE NDXID(NDXRECS)
    
    ' Assign Sequence Numbers to Individuals
    ' (applying BYPASS Rules if in effect)
310 SID = 0: XID = 0
312 SID = SID + 1: IF SID > NDXRECS THEN ID = 0: GOTO 319
    ID = NDXID(SID)
315 IF (IDLIST(ID) = 0) OR (ID = 0) THEN 312
    CALL FF1GetRec(ID)
    IF ((MID$(IA$, 5, 1) = "1") AND (FF1REC.EDUC <> 0)) THEN 319
    IF ((MID$(IA$, 6, 1) = "1") AND (FF1REC.WORK <> 0)) THEN 319
    IF ((MID$(IA$, 7, 1) = "1") AND (FF1REC.MIL <> 0)) THEN 319
    IF ((MID$(IA$, 8, 1) = "1") AND (FF1REC.HLTH <> 0)) THEN 319
    IF ((MID$(IA$, 11, 1) = "1") AND (FF1REC.EVENT <> 0)) THEN 319
    IF (BYPASS) AND (MID$(IA$, 2, 1) <> "2") THEN
       FID = FF1REC.FID: MID = FF1REC.MID
       IF ((FID > 0) * (IDLIST(FID) <> 0)) + ((MID > 0) * (IDLIST(MID) <> 0)) = 0 THEN 317
       IF (FF1REC.OLDCH <> 0) THEN 317
       IF ((MID$(IA$, 3, 1) = "1") AND (FF1REC.ADRS <> 0)) THEN 317
       IF ((MID$(IA$, 4, 1) <> "2") AND (FF1REC.COM <> 0) AND (MID$(IA$, 12, 1) = "2")) THEN 317
       IF ((MID$(IA$, 1, 1) <> "1") OR (FF1REC.SPOUSE = 0)) THEN 318
       CALL FF3GetRec(FF1REC.SPOUSE)
       IF (FF3SPOUSE.MLOC <> 0) THEN 317
       IF ((MID$(IA$, 3, 1) = "1") AND (FF3SPOUSE.LOC <> 0)) THEN 317
       IF ((MID$(IA$, 4, 1) <> "2") AND (FF3SPOUSE.COM <> 0)) THEN 317
       IF ID = FF3SPOUSE.SP1ID THEN
          CSPID = FF3SPOUSE.SP2ID: SC = FF3SPOUSE.SP2NXT
          ELSE
          CSPID = FF3SPOUSE.SP1ID: SC = FF3SPOUSE.SP1NXT
       END IF
       IF SC <> 0 THEN 317
       IF CSPID = 0 THEN 318 ELSE CALL FF1GetRec(CSPID)
       IF (MID$(IA$, 4, 1) = "4") AND (FF1REC.COM <> 0) THEN 317
       FID = FF1REC.FID: MID = FF1REC.MID: IF (FID > 0) OR (MID > 0) THEN 317
       IF (FF1REC.BLOC = 0) AND (FF1REC.BY = 0) AND (FF1REC.BMD = 0) AND ((FF1REC.DY = 0) OR (FF1REC.DY = 9999)) AND (FF1REC.DMD = 0) THEN 318
    END IF
317 IF (BYPASS2) AND (MID$(IA$, 1, 1) = "1") THEN
       CALL FF1GetRec(ID)
       SC = FF1REC.SPOUSE: IF (FF1REC.SEX <> FGENDR$(2)) OR (SC = 0) THEN 319
       IF (MID$(IA$, 4, 1) <> "1") OR (FF1REC.COM = 0) THEN
          CALL FF3GetRec(SC)
          IF ID = FF3SPOUSE.SP1ID THEN
             CSPID = FF3SPOUSE.SP2ID: SC = FF3SPOUSE.SP1NXT
             ELSE
             CSPID = FF3SPOUSE.SP1ID: SC = FF3SPOUSE.SP2NXT
          END IF
          IF SC <> 0 THEN 319
          IF (CSPID = 0) OR (IDLIST(CSPID) = 0) THEN 319
          IF (MID$(IA$, 2, 1) = "2") OR (FF1REC.OLDCH = 0) THEN 318
          CP = FF1REC.OLDCH
          WHILE CP <> 0
             CALL FF1GetRec(CP)
             IF CSPID = FF1REC.FID THEN CP = FF1REC.MCH ELSE CP = 0
          WEND
          IF CSPID = FF1REC.FID THEN 318
       END IF
    END IF
    GOTO 319
318 ' Record Bypassed (Assign Register Number if CountALL Option)
    IF CountALL THEN
       XID = XID + 1: IDLIST(ID) = -XID
       ELSE
       IDLIST(ID) = 0
    END IF
    GOTO 312
319 ' Assign Register Number for Records to be Printed
    IF ID > 0 THEN
       XID = XID + 1: IDLIST(ID) = XID
       GOTO 312
    END IF

    ' Begin Printing Register Report
    D1Y = VAL(RIGHT$(DATE$, 4))
    D1MD = (100 * VAL(LEFT$(DATE$, 2)) + VAL(MID$(DATE$, 4, 2)))
    CDATE$ = XRPTDate$(D1Y, D1MD, 0)
    MA$ = " 0 0 0 0 0 0 0 0 0 0 0 0"
    AMAXLV = ASC(WFHDR.AMAXLV): DMAXLV = ASC(WFHDR.DMAXLV)
    IF DMAXLV = 0 THEN AncSECTION = -1
    IF (AMAXLV > 0) AND (DMAXLV > 0) THEN
       CALL PutMSG("REGRM004")
       IF A = 27 THEN 399
       AncONLY = (A$ = "1"): DescONLY = (A$ = "2")
       IF NOT DescONLY THEN
          CALL GetRVAR("ASECHDR", X): IF FOUND THEN ASECHDR$ = RVAR$(X)
       END IF
       IF NOT AncONLY THEN
          CALL GetRVAR("DSECHDR", X): IF FOUND THEN DSECHDR$ = RVAR$(X)
       END IF
    END IF
    RptOPTION.SNGLREF = -1  ' Allow to Ask About Limiting References Recorded
    POPT$ = SNGLKEY$(4): PRTHDRID = 0
    CALL PrintOPEN: IF NOT OKAY THEN 399
    IF LEN(PLINE$) < 255 THEN PLINE$ = SPACE$(255)
    CL$ = PLINE$: xLINEAGE$ = CL$
    AFMT = FF2Hdr.REVLVL
    AMAXLV = ASC(WFHDR.AMAXLV): DMAXLV = ASC(WFHDR.DMAXLV)
    IF DMAXLV = 0 THEN AncSECTION = -1
    SELECT CASE WTYPE
       CASE 1: PRT.SECTION = "A"
       CASE 2: PRT.SECTION = "D"
       CASE 3: PRT.SECTION = "R"
    END SELECT
320 IF (AMAXLV > 0) AND (DMAXLV > 0) AND (NOT DescONLY) THEN
       AncSECTION = -1: NoDetail = -1
    END IF
325 SID = 0: RID = 0: NMOFST = 4: SPW = PAGE.BWIDTH - NMOFST
    AGL = 0: DGL = 0

330 ' Print Next Register Entry
    SID = SID + 1: IF SID > NDXRECS THEN ID = 0: GOTO 370
    ID = NDXID(SID)
    IF (IDLIST(ID) < 1) OR (ID = 0) THEN 330
    CALL FF1GetRec(ID)
    LAGL = AGL: LDGL = DGL
    X = RELATE(ID)
    AGL = -(X \ 256): DGL = X AND 127: RGL = AGL + DGL
    IF AncSECTION AND (DGL > 0) THEN 330
    ExtAdopt = ((X AND 128) > 0)
    CALL RwrkGET("RELREC", RC, RELPTR&(ID))
    IF RID = 0 THEN
       LSET SYSVAR.SUBJECT = FMTNAME$(1)
       IF PRTHDRID THEN
          X = LEN(RTRIM$(SYSVAR.SUBJECT))
          IF X > 2 THEN
             MID$(SYSVAR.SUBJECT, X + 2, 255) = "(" + RIDLbl$ + FNX$(RID) + ")"
          END IF
       END IF
    END IF

350 ' Print Appropriate Section Headers
    IF AncSECTION AND NoDetail AND (RID = 0) THEN STYPE = 1: GOTO 355
    IF (AMAXLV > 0) AND (DGL = 0) AND (NOT AncSECTION) THEN
       STYPE = 2: GOTO 355
    END IF
    IF (AncSECTION AND (AGL <> LAGL)) THEN
       STYPE = 3: GOTO 355
    END IF
    IF ((NOT AncSECTION) AND (DGL <> LDGL)) THEN
       STYPE = 4: GOTO 355
    END IF
    IF (RID <> 0) THEN
       CALL WriteLINE
    END IF
    GOTO 360

355 ' Print Section Header
    IF (PAGE.BLENGTH - PAGE.LINE) < 5 THEN
       CALL PageBREAK
       ELSE
       CALL WriteLINE
    END IF
    SELECT CASE STYPE
       CASE 1: MID$(PLINE$, NMOFST, SPW) = LINESEP1$
       CASE 2: MID$(PLINE$, NMOFST, SPW) = LINESEP1$
       CASE ELSE: MID$(PLINE$, NMOFST, SPW) = LINESEP2$
    END SELECT
    CALL PrintLINE
    SELECT CASE STYPE
       CASE 1: LSET PLINE$ = ASECHDR$
       CASE 2: LSET SYSVAR.SUBJECT = FMTNAME$(1)
               LSET SYSVAR.MVAR = "[" + LTRIM$(STR$(WFANCREC.LINEAGE)) + "]"
               LSET PLINE$ = DSECHDR$
       CASE 3: LSET PLINE$ = "<<" + Relation$("L", "X", AGL, AGL, RRELCD$()) + ">>"
               LSET PLINE$ = UCX$(PLINE$)
       CASE 4: IF RelInHEAD THEN xAGL = AGL: xRGL = RGL ELSE xAGL = 0: xRGL = DGL
               LSET PLINE$ = "<<" + Relation$("L", "X", xAGL, xRGL, RRELCD$()) + ">>"
               LSET PLINE$ = UCX$(PLINE$)
    END SELECT
    CALL HTFBuild(PLINE$, PAGE.BWIDTH)
    CALL PrintLINE
    SELECT CASE STYPE
       CASE 1: MID$(PLINE$, NMOFST, SPW) = LINESEP1$
       CASE 2: MID$(PLINE$, NMOFST, SPW) = LINESEP1$
       CASE ELSE: MID$(PLINE$, NMOFST, SPW) = LINESEP2$
    END SELECT
    CALL PrintLINE
    CALL WriteLINE
    
360 ' Print Register Entry
    RID = ID: GOSUB PrtReport
    IF (A = 27) THEN 395
    GOTO 330

370 ' Check For Additional Sections
    IF (AMAXLV = 0) OR (DMAXLV = 0) THEN 395
    IF AncONLY OR (NOT AncSECTION) THEN 395
    CALL EvenBREAK("noEOF"): IF A = 27 THEN 395
    AncSECTION = 0: NoDetail = 0
    GOTO 325

395 ' Finished Printing Reports
    CALL RptCLOSE
    IF NOT OKAY THEN 399
    IF PRT.AGAIN THEN
       PAGE.COUNT = PAGE.BGNPAGE - 1
       GOTO 320
    END IF
399 ' Return to Main Program
    ON LOCAL ERROR GOTO 0
    CALL FamCLOSE
    EXIT SUB

900 ' Error Routine
    SYSVAR.NVAR1 = ERL
    CALL ERRMessage("XXXXM098", ERR, ERL)
    RESUME 399

PrtReport: ' Print Family Group Report
    IF FF1REC.RID <> RID THEN CALL FF1GetRec(RID)
    REDIM APTR(MAXADR), AIX(MAXADR), ADRDATE$(MAXADR)
    AX = -1: APTR(0) = FF1REC.ADRS: IF APTR(0) <> 0 THEN AX = 0
    SC = FF1REC.SPOUSE: CC = FF1REC.OLDCH: ComPTR = FF1REC.COM
    CALL BldMLIST
    EVENTPTR = FF1REC.EVENT
    OA(1) = FF1REC.HLTH: OA(2) = FF1REC.EDUC
    OA(3) = FF1REC.WORK: OA(4) = FF1REC.MIL
    GOSUB 600: GOSUB 610: MCNT = 0
    GOSUB 2300: IF A = 27 THEN 590
    IF NoDetail THEN 590
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
       MLOC = 0
    END IF
    IF (MID$(IA$, 2, 1) <> "2") AND ((CC <> 0)) THEN
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
             MLOC = FF3SPOUSE.MLOC
             GOSUB 630
             IF CC = FF3SPOUSE.SP1ID THEN CSPID = FF3SPOUSE.SP2ID ELSE CSPID = FF3SPOUSE.SP1ID
             IF CSPID = 0 THEN
                MarriageDATA.SPNAME = "(" + Unknown$ + ")"
                ELSE
                CALL FF1GetRec(CSPID)
                MarriageDATA.SPID = SFILL$(CSPID, 5)
                MarriageDATA.SPNAME = FMTNAME$(0)
             END IF
          END IF
          GOSUB 2410: IF A = 27 THEN 590
          CC = NCH
       WEND
    END IF

460 ' Process RESIDENCEes
    IF (MID$(IA$, 3, 1) <> "2") AND ((APTR(0) <> 0)) THEN
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
    END IF

470 ' Process EVENTS
    IF (MID$(IA$, 11, 1) <> "2") AND ((EVENTPTR <> 0)) THEN
       GOSUB PrintEventH: IF A = 27 THEN 590
       WHILE EVENTPTR <> 0
          GOSUB PrintEVENT: EVENTPTR = FF3EVENT.NXT
       WEND
    END IF

500 ' Process OTHER Info
    FOR OT = 1 TO 4: GOSUB 2700: NEXT OT
590 ' End of PrtReport
    IF A = 27 THEN LSET PLINE$ = " ": LOFFSET = TWIDTH
    RETURN

600 ' Initialise Data Areas - Individual
    NameDATA = BNDATA
    FatherDATA = BPDATA
    MotherDATA = BPDATA
    MarriageDATA = BMDATA
    RETURN

610 ' Place Information in Data Areas
    NameDATA.NAME = FMTNAME$(0)
    FID = FF1REC.FID: MID = FF1REC.MID
    STAT1 = ASC(FF1REC.STATUS)
    MSTAT = ((STAT1 \ 4) AND 3): FSTAT = (STAT1 AND 3)
    GOSUB 620

615 ' Name Record Information Common to Subject, Spouse, Child
    CALL GetBDEvents(BD$, BPL$, DD$, DPL$, "")
    NameDATA.BDATE = LTRIM$(BD$)
    NameDATA.DDATE = LTRIM$(DD$)
    NameDATA.BPLACE = BPL$
    NameDATA.DPLACE = DPL$
    NameDATA.SEX = FF1REC.SEX
    RETURN

620 ' Father, Mother Information
    XID = FF1REC.RID
    IF FID <> 0 THEN
       CALL FF1GetRec(FID): STAT1 = ASC(FF1REC.STATUS)
       FatherDATA.NAME = FMTNAME$(0)
       CALL GetBDEvents(BD$, "", DD$, "", AGE$)
       FatherDATA.BDATE = LTRIM$(BD$)
       FatherDATA.DDATE = LTRIM$(DD$)
       IF (FF1REC.BY > 0) AND (FF1REC.DY > 0) AND (FF1REC.DY < 9999) THEN
          FatherDATA.AGE = LTRIM$(AGE$)
       END IF
    END IF
    IF MID <> 0 THEN
       CALL FF1GetRec(MID): STAT1 = ASC(FF1REC.STATUS)
       RSET RID$ = FNX$(MID): RSTAT = MSTAT
       MotherDATA.NAME = FMTNAME$(0)
       CALL GetBDEvents(BD$, "", DD$, "", AGE$)
       MotherDATA.BDATE = LTRIM$(BD$)
       MotherDATA.DDATE = LTRIM$(DD$)
       IF (FF1REC.BY > 0) AND (FF1REC.DY > 0) AND (FF1REC.DY < 9999) THEN
          MotherDATA.AGE = LTRIM$(AGE$)
       END IF
    END IF
    CALL FF1GetRec(XID)
    RETURN

630 ' Marriage Information
    MarriageDATA.MDATE = XRPTDate$(FF3SPOUSE.D1MD, FF3SPOUSE.D1Y, DS)
    MarriageDATA.DVDATE = XRPTDate$(FF3SPOUSE.D2MD, FF3SPOUSE.D2Y, DS)
    IF MLOC <> 0 THEN
       CALL FF3GetRec(MLOC)
       MarriageDATA.MPLACE = FF3PLACE.PLACE1
       MarriageDATA.DVPLACE = FF3PLACE.PLACE2
    END IF
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
    IF GCH > 0 THEN NameDATA.NUMCH = LTRIM$(STR$(GCH))
    RETURN

2300 ' Print Subject Information
     PAGE.LOFST = 1: PAGE.COFST = 3: Z = 1
     IF AncSECTION THEN
        CALL RwrkGET("RELREC", RC, RELPTR&(RID))
        LSET CL$ = LTRIM$(STR$(WFANCREC.LINEAGE))
        ELSE
        LSET CL$ = FNX$(ABS(IDLIST(RID)))
     END IF
     Z = INSTR(Z, CL$, "  ")
     MID$(CL$, Z, 1) = ".": Z = Z + 2
     PAGE.COFST = Z
     IF SHOWLIN AND (NOT AncSECTION) THEN
        CALL LINEAGE("BUILD")
        IF ShowREL THEN CALL LINEAGE("PRINT")
     END IF
     NMOFST = PAGE.COFST: SPW = PAGE.BWIDTH - NMOFST: Z = NMOFST
     IF AncSECTION AND NoDetail THEN
        MID$(CL$, Z, 255) = "(" + FNX$(ABS(IDLIST(RID))) + ")"
        Z = INSTR(Z, CL$, " ") + 1
     END IF
     MID$(CL$, Z, 255) = RFLBL$(5): Z = Z + RFLTH(5) + (RFLTH(5) = 0)
     IF RptOPTION.PRID THEN
        MID$(CL$, Z + 1, 255) = "(" + RFLBL$(2): Z = Z + RFLTH(2) + 2
        MID$(CL$, Z, 6) = FNX$(RID) + ")": Z = INSTR(Z, CL$, "  ")
     END IF
     MID$(CL$, Z + 1, 52) = NameDATA.NAME: Z = INSTR(Z, CL$, "   ")
     PAGE.XID = RID: PAGE.XTYPE = 2
     IF SHOWLIN AND (NOT ShowREL) AND (NOT AncSECTION) THEN
        IF SID > 1 THEN CALL LINEAGE("PRINT"): Z = 1
     END IF
     IF (RTRIM$(NameDATA.BDATE) <> "") OR (RTRIM$(NameDATA.BPLACE) <> "") THEN
        MID$(CL$, Z, 255) = RFLBL$(8): Z = Z + RFLTH(8)
        IF RTRIM$(NameDATA.BDATE) <> "" THEN
           MID$(CL$, Z + 1, 12) = NameDATA.BDATE: Z = INSTR(Z, CL$, "   ")
        END IF
        IF RTRIM$(NameDATA.BPLACE) <> "" THEN
           MID$(CL$, Z, 255) = RFLBL$(9): Z = Z + RFLTH(9)
           MID$(CL$, Z + 1, 52) = NameDATA.BPLACE: Z = INSTR(Z, CL$, "   ")
        END IF
     END IF
     IF (RTRIM$(NameDATA.DDATE) <> "") OR (RTRIM$(NameDATA.DPLACE) <> "") THEN
        MID$(CL$, Z, 255) = RFLBL$(10): Z = Z + RFLTH(10)
        IF RTRIM$(NameDATA.DDATE) <> "" THEN
           MID$(CL$, Z + 1, 12) = NameDATA.DDATE: Z = INSTR(Z, CL$, "   ")
        END IF
        IF RTRIM$(NameDATA.DPLACE) <> "" THEN
           MID$(CL$, Z, 255) = RFLBL$(11): Z = Z + RFLTH(11)
           MID$(CL$, Z + 1, 41) = NameDATA.DPLACE
        END IF
     END IF
     CALL PutDATA
     IF FID <> 0 THEN
        MID$(CL$, 1, 255) = RFLBL$(6): Z = 1 + RFLTH(6)
        IF RptOPTION.PRID THEN
           MID$(CL$, Z + 1, 255) = "(" + RFLBL$(2): Z = Z + 2 + RFLTH(2)
           MID$(CL$, Z, 6) = FNX$(FID) + ")": Z = INSTR(Z, CL$, "  ")
        END IF
        MID$(CL$, Z + 1, 52) = FatherDATA.NAME: Z = INSTR(Z, CL$, "   ")
        IF RTRIM$(FatherDATA.BDATE) <> "" THEN
           MID$(CL$, Z, 255) = RFLBL$(8): Z = Z + RFLTH(8)
           MID$(CL$, Z + 1, 12) = FatherDATA.BDATE: Z = INSTR(Z + 1, CL$, "   ")
        END IF
        IF (RTRIM$(FatherDATA.DDATE) <> "") AND (MID$(FatherDATA.DDATE, 1, 3) <> "999") THEN
           MID$(CL$, Z, 255) = RFLBL$(10): Z = Z + RFLTH(10)
           MID$(CL$, Z + 1, 12) = FatherDATA.DDATE: Z = INSTR(Z, CL$, "   ")
           IF RTRIM$(FatherDATA.AGE) <> "" THEN
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
           MID$(CL$, Z, 6) = FNX$(MID) + ")": Z = INSTR(Z, CL$, "  ")
        END IF
        MID$(CL$, Z + 1, 52) = MotherDATA.NAME: Z = INSTR(Z, CL$, "   ")
        IF RTRIM$(MotherDATA.BDATE) <> "" THEN
           MID$(CL$, Z, 255) = RFLBL$(8): Z = Z + RFLTH(8)
           MID$(CL$, Z + 1, 12) = MotherDATA.BDATE: Z = INSTR(Z, CL$, "   ")
        END IF
        IF (RTRIM$(MotherDATA.DDATE) <> "") AND (MID$(MotherDATA.DDATE, 1, 3) <> "999") THEN
           MID$(CL$, Z, 255) = RFLBL$(10): Z = Z + RFLTH(10)
           MID$(CL$, Z + 1, 12) = MotherDATA.DDATE: Z = INSTR(Z, CL$, "   ")
           IF MotherDATA.AGE <> SPACE$(3) THEN
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
2325 RETURN

2350 ' Print Spouse Info
     CALL WriteLINE: IF A = 27 THEN 2380
     PAGE.LOFST = NMOFST: PAGE.COFST = NMOFST
     MID$(CL$, 1, 255) = RFLBL$(14): Z = RFLTH(14) + 1
     IF RTRIM$(MarriageDATA.MDATE) <> "" THEN
        MID$(CL$, Z + 1, 11) = MarriageDATA.MDATE: Z = INSTR(Z, CL$, "   ")
     END IF
     IF RTRIM$(MarriageDATA.MPLACE) <> "" THEN
        MID$(CL$, Z, 255) = RFLBL$(15): Z = Z + RFLTH(15)
        MID$(CL$, Z + 1, 41) = MarriageDATA.MPLACE: Z = INSTR(Z, CL$, "   ")
     END IF
     IF RTRIM$(MarriageDATA.DVDATE) <> "" THEN
        MID$(CL$, Z, 255) = RFLBL$(16): Z = Z + RFLTH(16)
        MID$(CL$, Z + 1, 11) = MarriageDATA.DVDATE: Z = INSTR(Z, CL$, "  ")
        IF (RTRIM$(MarriageDATA.DVPLACE) <> "") THEN
           MID$(CL$, Z, 255) = RFLBL$(17): Z = Z + RFLTH(17)
           MID$(CL$, Z + 1, 41) = MarriageDATA.DVPLACE: Z = INSTR(Z, CL$, "   ")
        END IF
     END IF
     MID$(CL$, Z, 255) = RFLBL$(13): Z = Z + RFLTH(13)
     IF CSPID <> 0 THEN
        IF (IDLIST(CSPID) <> 0) AND ((RELATE(CSPID) AND 127) = 0) THEN
           CALL RwrkGET("RELREC", RC, RELPTR&(CSPID))
           MID$(CL$, Z + 1, 255) = "[" + LTRIM$(STR$(WFANCREC.LINEAGE)) + "]"
           Z = INSTR(Z, CL$, "  ")
        END IF
        IF RptOPTION.PRID THEN
           MID$(CL$, Z + 1, 255) = "(" + RFLBL$(2): Z = Z + RFLTH(2) + 2
           MID$(CL$, Z, 6) = FNX$(CSPID) + ")": Z = INSTR(Z, CL$, "  ")
        END IF
     END IF
     MID$(CL$, Z + 1, 52) = NameDATA.NAME: Z = INSTR(Z, CL$, "   ")
     IF (RTRIM$(NameDATA.BDATE) <> "") OR (RTRIM$(NameDATA.BPLACE) <> "") THEN
        MID$(CL$, Z, 255) = RFLBL$(8): Z = Z + RFLTH(8)
        IF RTRIM$(NameDATA.BDATE) <> "" THEN
           MID$(CL$, Z + 1, 12) = NameDATA.BDATE: Z = INSTR(Z, CL$, "   ")
        END IF
        IF RTRIM$(NameDATA.BPLACE) <> "" THEN
           MID$(CL$, Z, 255) = RFLBL$(9): Z = Z + RFLTH(9)
           MID$(CL$, Z + 1, 41) = NameDATA.BPLACE: Z = INSTR(Z, CL$, "   ")
        END IF
     END IF
     IF (RTRIM$(NameDATA.DDATE) <> "") OR (RTRIM$(NameDATA.DPLACE) <> "") THEN
        MID$(CL$, Z, 255) = RFLBL$(10): Z = Z + RFLTH(10)
        IF RTRIM$(NameDATA.DDATE) <> "" THEN
           MID$(CL$, Z + 1, 12) = NameDATA.DDATE: Z = INSTR(Z, CL$, "   ")
        END IF
        IF RTRIM$(NameDATA.DPLACE) <> "" THEN
           MID$(CL$, Z, 255) = RFLBL$(11): Z = Z + RFLTH(11)
           MID$(CL$, Z + 1, 41) = NameDATA.DPLACE
        END IF
     END IF
     PAGE.XID = CSPID: PAGE.XTYPE = 3: CALL PutDATA
     IF FID <> 0 THEN
        MID$(CL$, 1, 255) = RFLBL$(6): Z = 1 + RFLTH(6)
        IF RptOPTION.PRID THEN
           MID$(CL$, Z + 1, 255) = "(" + RFLBL$(2): Z = Z + 2 + RFLTH(2)
           MID$(CL$, Z, 6) = FNX$(FID) + ")": Z = INSTR(Z, CL$, "  ")
        END IF
        MID$(CL$, Z + 1, 52) = FatherDATA.NAME: Z = INSTR(Z, CL$, "   ")
        IF RTRIM$(FatherDATA.BDATE) <> "" THEN
           MID$(CL$, Z, 255) = RFLBL$(8): Z = Z + RFLTH(8)
           MID$(CL$, Z + 1, 12) = FatherDATA.BDATE
           Z = INSTR(Z + 1, CL$, "   ")
        END IF
        IF (RTRIM$(FatherDATA.DDATE) <> "") AND (MID$(FatherDATA.DDATE, 1, 3) <> "999") THEN
           MID$(CL$, Z, 255) = RFLBL$(10): Z = Z + RFLTH(10)
           MID$(CL$, Z + 1, 12) = FatherDATA.DDATE: Z = INSTR(Z, CL$, "   ")
           IF RTRIM$(FatherDATA.AGE) <> "" THEN
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
           MID$(CL$, Z, 6) = FNX$(MID) + ")": Z = INSTR(Z, CL$, "  ")
        END IF
        MID$(CL$, Z + 1, 52) = MotherDATA.NAME: Z = INSTR(Z, CL$, "   ")
        IF RTRIM$(MotherDATA.BDATE) <> "" THEN
           MID$(CL$, Z, 255) = RFLBL$(8): Z = Z + RFLTH(8)
           MID$(CL$, Z + 1, 12) = MotherDATA.BDATE: Z = INSTR(Z, CL$, "   ")
        END IF
        IF (RTRIM$(MotherDATA.DDATE) <> "") AND (MID$(MotherDATA.DDATE, 1, 3) <> "999") THEN
           MID$(CL$, Z, 255) = RFLBL$(10): Z = Z + RFLTH(10)
           MID$(CL$, Z + 1, 12) = MotherDATA.DDATE: Z = INSTR(Z, CL$, "   ")
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
2380 MCNT = MCNT + 1
        RETURN

2400 ' Print Child Information Header
     CALL WriteLINE
     IF A <> 27 THEN
        MID$(PLINE$, NMOFST, 255) = ChildHDR$
        CALL PrintLINE
     END IF
     RETURN

2410 ' Print Child Information
     PAGE.LOFST = NMOFST: CH = CH + 1: Z = NMOFST: L = 7
     IF NOT AncSECTION THEN
        IF CountALL THEN
           IF IDLIST(CC) > 0 THEN MID$(PLINE$, Z - 1, 1) = "+"
        END IF
        IF IDLIST(CC) <> 0 THEN
           MID$(PLINE$, Z, 255) = FNX$(ABS(IDLIST(CC))) + "."
           Z = INSTR(Z, PLINE$, "  ")
        END IF
        Z = Z - NMOFST: IF Z < 4 THEN Z = 4
        L = 11 - Z: Z = NMOFST + Z
     END IF
     MID$(PLINE$, Z, L) = RIGHT$(ROMAN$(CH), L)
     Z = Z + L
     IF CSTAT > 0 THEN
        FF1REC.STATUS = CHR$(CSTAT)
        MID$(PLINE$, Z, 1) = FDSTAT$("FID"): CSTAT = 0
     END IF
     Z = Z + 1: PAGE.LOFST = Z: PAGE.COFST = Z: Z = 1

     'SKIP Gender for Now
     'MID$(CL$, 1, 1) = MID$(NameDATA$, 131, 1): Z = 3
     'PAGE.COFST = PAGE.LOFST + 2

     'IF AncSECTION AND (NOT NoDetail) THEN
     IF SHOWLIN THEN
        IF (IDLIST(CC) <> 0) AND ((RELATE(CC) AND 127) = 0) THEN
           CALL RwrkGET("RELREC", RC, RELPTR&(CC))
           MID$(CL$, Z, 255) = "[" + LTRIM$(STR$(WFANCREC.LINEAGE)) + "]"
           Z = INSTR(Z, CL$, "  ") + 1
        END IF
     END IF
     'END IF
     IF RptOPTION.PRID THEN
        MID$(CL$, Z, 255) = "(" + RFLBL$(2) + LTRIM$(NameDATA.RID)
        Z = INSTR(Z, CL$, "  ")
        MID$(CL$, Z, 255) = RFLBL$(3) + LTRIM$(NameDATA.PID)
        Z = INSTR(Z, CL$, "  ")
        IF PSTAT <> 0 THEN
           FF1REC.STATUS = CHR$(PSTAT)
           MID$(CL$, Z, 1) = FDSTAT$("FID"): PSTAT = 0
           Z = INSTR(Z, CL$, " ")
        END IF
        MID$(CL$, Z, 1) = ")": Z = Z + 2
     END IF
     MID$(CL$, Z, 52) = NameDATA.NAME: Z = INSTR(Z, CL$, "   ")
     IF (RTRIM$(NameDATA.BDATE) <> "") OR (RTRIM$(NameDATA.BPLACE) <> "") THEN
        MID$(CL$, Z, 255) = RFLBL$(8): Z = Z + RFLTH(8)
        IF RTRIM$(NameDATA.BDATE) <> "" THEN
           MID$(CL$, Z + 1, 12) = NameDATA.BDATE: Z = INSTR(Z, CL$, "   ")
        END IF
        IF RTRIM$(NameDATA.BPLACE) <> "" THEN
           MID$(CL$, Z, 255) = RFLBL$(9): Z = Z + RFLTH(9)
           MID$(CL$, Z + 1, 52) = NameDATA.BPLACE
           Z = INSTR(Z, CL$, "   ")
        END IF
     END IF
     IF (RTRIM$(NameDATA.DDATE) <> "") OR (RTRIM$(NameDATA.DPLACE) <> "") THEN
        MID$(CL$, Z, 255) = RFLBL$(10): Z = Z + RFLTH(10)
        IF RTRIM$(NameDATA.DDATE) <> "" THEN
           MID$(CL$, Z + 1, 12) = NameDATA.DDATE: Z = INSTR(Z, CL$, "   ")
        END IF
        IF RTRIM$(NameDATA.DPLACE) <> "" THEN
           MID$(CL$, Z, 255) = RFLBL$(11): Z = Z + RFLTH(11)
           MID$(CL$, Z + 1, 41) = NameDATA.DPLACE: Z = INSTR(Z, CL$, "   ")
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
        Z = INSTR(Z, CL$, "  ")
     END IF
     PAGE.XID = CC: PAGE.XTYPE = 4: CALL PutDATA
     IF RTRIM$(MarriageDATA.SPNAME) <> "" THEN
        MID$(CL$, 1, 255) = RFLBL$(25): Z = 1 + RFLTH(25)
        IF RptOPTION.PRID THEN
           MID$(CL$, Z + 1, 255) = "(" + RFLBL$(2): Z = Z + RFLTH(2) + 2
           MID$(CL$, Z, 6) = LTRIM$(MarriageDATA.SPID) + ")"
           Z = INSTR(Z, CL$, "  ")
        END IF
        MID$(CL$, Z + 1, 52) = MarriageDATA.SPNAME: Z = INSTR(Z, CL$, "   ")
        IF RTRIM$(MarriageDATA.MDATE) <> "" THEN
           MID$(CL$, Z, 255) = RFLBL$(30): Z = Z + RFLTH(30)
           MID$(CL$, Z + 1, 11) = MarriageDATA.MDATE
           Z = INSTR(Z, CL$, "   ")
           IF RTRIM$(MarriageDATA.MPLACE) <> "" THEN
              MID$(CL$, Z, 255) = RFLBL$(15): Z = Z + RFLTH(15)
              MID$(CL$, Z + 1, 41) = MarriageDATA.MPLACE: Z = INSTR(Z, CL$, "   ")
           END IF
        END IF
        PAGE.XID = CSPID: PAGE.XTYPE = 9: CALL PutDATA
     END IF
     IF MID$(IA$, 12, 1) <> "2" THEN GOSUB PrintCOM
     CALL PrintLINE
     'IF CHILDSEP AND (NCH > 0) THEN CALL WriteLINE
     RETURN

2500 ' Print Events
PrintEventH:
     'MID$(PLINE$, NMOFST, SPW) = LINESEP2$: CALL PrintLINE
     CALL WriteLINE
     EVNTCNT = 0
     IF A <> 27 THEN
        MID$(PLINE$, NMOFST, 255) = RFLBL$(42)
        CALL PrintLINE
     END IF
     RETURN

PrintEVENT:
     IF EVNTCNT <> 0 THEN
        'MID$(PLINE$, NMOFST, SPW) = LINESEP3$: CALL PrintLINE
        CALL WriteLINE
     END IF
2510 EVNTCNT = EVNTCNT + 1: EventDATA = BEDATA: ComPTR = 0
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
     EventDATA.DSTAT = FF3EVENT.DSTAT
     IF FF3EVENT.LPLNM <> 0 THEN
        CALL FF3GetRec(FF3EVENT.LPLNM)
        LSET EventDATA.PLACE = FF3PLACE2.PLACE
        ELSE
        LSET EventDATA.PLACE = FF3EVENT.PLACE
     END IF
     X = ASC(FF3EVENT.IMPORT)
     IF X > 0 THEN EventDATA.IMPORT = LTRIM$(STR$(X))

2520 ' Print Information Formatted into EventDATA
     PAGE.LOFST = NMOFST + 2: PAGE.COFST = PAGE.LOFST: X = 1
     MID$(CL$, X, 255) = RFLBL$(43): X = X + RFLTH(43)
     MID$(CL$, X + 1, 15) = EventDATA.DESC
     X = INSTR(X, CL$, "  ")
2525 IF RTRIM$(EventDATA.DATE) <> "" THEN
        MID$(CL$, X, 255) = RFLBL$(44): X = X + RFLTH(44)
        MID$(CL$, X + 1, 12) = EventDATA.DATE
        X = INSTR(X, CL$, "  ")
        IF EventDATA.DSTAT <> " " THEN
           MID$(CL$, X, 3) = "(" + EventDATA.DSTAT + ")": X = X + 3
        END IF
     END IF
2530 IF RTRIM$(EventDATA.PLACE) <> "" THEN
        MID$(CL$, X, 255) = RFLBL$(45): X = X + RFLTH(45)
        MID$(CL$, X + 1, 255) = EventDATA.PLACE
        X = INSTR(X, CL$, "  ")
     END IF
     IF (RTRIM$(EventDATA.IMPORT) <> "") THEN
        MID$(CL$, X, 255) = RFLBL$(46): X = X + RFLTH(46)
        MID$(CL$, X, 3) = EventDATA.IMPORT
     END IF
     LSET PLINE$ = "": CALL PutDATA
     IF MID$(IA$, 10, 1) = "1" THEN
        GOSUB PrintCOM
     END IF
     RETURN

PrintRESH:
     'MID$(PLINE$, NMOFST, SPW) = LINESEP2$: CALL PrintLINE
     CALL WriteLINE
     IF A <> 27 THEN
        MID$(PLINE$, NMOFST, 255) = RFLBL$(31)
        'CALL PrintLINE
     END IF
     RETURN

PrintRESI: CALL WriteLINE: IF A = 27 THEN 2595
     LSET ADRDATA$ = " ": ComPTR = 0
     CALL FF2GetRec(APTR): ComPTR = FF2PFX.COM
     MID$(ADRDATA$, 1, 11) = RTRIM$(XRPTDate$(FF2PFX.D1MD, FF2PFX.D1Y, DS))
     MID$(ADRDATA$, 12, 11) = RTRIM$(XRPTDate$(FF2PFX.D2MD, FF2PFX.D2Y, DS))
     IF AFMT = 0 THEN
        IF (FF2SREC.FON1 <> 0) OR (FF2SREC.FON2 <> 0) OR (FF2SREC.FON3 <> 0) THEN
           MID$(ADRDATA$, 23, 14) = ZFILL$(FF2SREC.FON1, 3) + "-" + ZFILL$(FF2SREC.FON2, 3) + "-" + ZFILL$(FF2SREC.FON3, 4) + "  "
        END IF
        MID$(ADRDATA$, 37, 30) = LTRIM$(RTRIM$(FF2SREC.LINE1))
        MID$(ADRDATA$, 67, 30) = LTRIM$(RTRIM$(FF2SREC.LINE2))
        MID$(ADRDATA$, 97, 15) = LTRIM$(RTRIM$(FF2SREC.CITY))
        MID$(ADRDATA$, 117, 4) = LTRIM$(RTRIM$(FF2SREC.STATE))
        MID$(ADRDATA$, 129, 5) = LTRIM$(RTRIM$(FF2SREC.ZIP))
        ELSE
        IF (LTRIM$(FF2LREC.PHONE) <> "") AND (LEFT$(FF2LREC.PHONE, 13) <> "(000)000-0000") AND (LEFT$(FF2LREC.PHONE, 12) <> "000-000-0000") THEN
           MID$(ADRDATA$, 23, 14) = LTRIM$(RTRIM$(FF2LREC.PHONE))
        END IF
        MID$(ADRDATA$, 37, 30) = LTRIM$(RTRIM$(FF2LREC.LINE1))
        MID$(ADRDATA$, 67, 30) = LTRIM$(RTRIM$(FF2LREC.LINE2))
        MID$(ADRDATA$, 97, 20) = LTRIM$(RTRIM$(FF2LREC.CITY))
        MID$(ADRDATA$, 117, 12) = LTRIM$(RTRIM$(FF2LREC.STATE))
        MID$(ADRDATA$, 129, 12) = LTRIM$(RTRIM$(FF2LREC.ZIP))
        MID$(ADRDATA$, 141, 16) = LTRIM$(RTRIM$(FF2LREC.COUNTRY))
     END IF

     ' Print Information Formatted into ADRDATA$
     PAGE.LOFST = NMOFST + 2: PAGE.COFST = PAGE.LOFST
     X = 1: LSET CL$ = ""
     IF (MID$(ADRDATA$, 1, 1) <> CHR$(FC)) THEN
        MID$(CL$, X, 255) = RFLBL$(32): X = X + RFLTH(32)
        MID$(CL$, X + 1, 11) = MID$(ADRDATA$, 1, 11)
        X = INSTR(X, CL$, "     ")
     END IF
     IF (MID$(ADRDATA$, 12, 1) <> CHR$(FC)) THEN
        MID$(CL$, X, 255) = RFLBL$(33): X = X + RFLTH(33)
        MID$(CL$, X + 1, 11) = MID$(ADRDATA$, 12, 11)
        X = INSTR(X, CL$, "   ")
     END IF
     IF (MID$(ADRDATA$, 23, 1) <> CHR$(FC)) THEN
        MID$(CL$, X, 255) = RFLBL$(34): X = X + RFLTH(34)
        MID$(CL$, X + 1, 14) = MID$(ADRDATA$, 23, 14)
        X = INSTR(X, CL$, "   ")
     END IF
     MID$(CL$, X, 255) = RFLBL$(35): X = X + RFLTH(35)
     MID$(CL$, X + 1, 30) = MID$(ADRDATA$, 37, 30)
     X2 = X: X = INSTR(X, CL$, "  ")
     IF MID$(ADRDATA$, 67, 1) <> " " THEN
        IF (X > X2) THEN MID$(CL$, X, 1) = ",": X = X + 2
        MID$(CL$, X, 30) = MID$(ADRDATA$, 67, 30)
     END IF

     'Free Form Address for Merged Reports
     CALL PutDATA: X = 1
     X = INSTR(X, CL$, "  ")
     IF EuroAdr THEN
        X2 = 129: ZL = 5 - 7 * (AFMT = 1): EuroZIP = 0
        IF LTRIM$(MID$(ADRDATA$, X2, ZL)) <> "" THEN
           MID$(CL$, X, 1) = ",": X = X + 2: EuroZIP = -1
           MID$(CL$, X, ZL) = LTRIM$(MID$(ADRDATA$, X2, ZL))
           X = INSTR(X, CL$, "  ")
        END IF
     END IF
     X2 = 97
     ZL = 15 - 5 * (AFMT = 1)
     IF LTRIM$(MID$(ADRDATA$, X2, ZL)) <> "" THEN
        IF NOT EuroZIP THEN MID$(CL$, X, 1) = ",": X = X + 1
        X = X + 1
        MID$(CL$, X, ZL) = LTRIM$(MID$(ADRDATA$, 97, ZL))
        X = INSTR(X, CL$, "  ")
     END IF
     X2 = 117: ZL = 4 - 8 * (AFMT = 1)
     IF LTRIM$(MID$(ADRDATA$, X2, ZL)) <> "" THEN
        MID$(CL$, X, 1) = ","
        MID$(CL$, X + 2, ZL) = LTRIM$(MID$(ADRDATA$, X2, ZL))
        X = INSTR(X + 2, CL$, "  ")
     END IF
     IF NOT EuroAdr THEN
        X2 = 129: ZL = 5 - 7 * (AFMT = 1)
        IF LTRIM$(MID$(ADRDATA$, X2, ZL)) <> "" THEN
           MID$(CL$, X + 2, ZL) = LTRIM$(MID$(ADRDATA$, X2, ZL))
           X = INSTR(X + 2, CL$, "  ")
        END IF
     END IF
     IF AFMT = 1 THEN
        X2 = 141: ZL = 16
        IF LTRIM$(MID$(ADRDATA$, X2, ZL)) <> "" THEN
           MID$(CL$, X, 1) = ","
           MID$(CL$, X + 2, ZL) = MID$(ADRDATA$, X2, ZL)
        END IF
     END IF
2560 CALL PutDATA
     IF MID$(IA$, 10, 1) = "1" THEN
        GOSUB PrintCOM
     END IF
     CALL PrintLINE
2595 RETURN

PrintCOM: ' Process COMMENTS
     RI = 0: PAGE.CINDENT = 0
     COMBGN = 0: PAGE.LWIDTH = CWIDTH: CWIDTH = 0
     IF PAGE.LWIDTH = 0 THEN PAGE.LWIDTH = PAGE.BWIDTH - PAGE.CINDENT
     IF ((ComPTR <> 0) OR (COMPTR2 <> 0)) THEN
        IF ComPTR <> 0 THEN CALL ComPRINT(ComPTR, COMBGN): ComPTR = 0
        IF COMPTR2 <> 0 THEN CALL ComPRINT(COMPTR2, COMBGN): COMPTR2 = 0
     END IF
     CALL PrintLINE
2695 RETURN

2700 ' Print OTHER Info
     IF (OA(OT) = 0) OR (MID$(IA$, OT + 4, 1) <> "1") THEN 2795
     CALL WriteLINE: IF A = 27 THEN 2790
     'MID$(PLINE$, NMOFST, SPW) = LINESEP2$: CALL PrintLINE
     IF A = 27 THEN 2790
     SELECT CASE OT
       CASE 1: X = 54
       CASE 2: X = 49
       CASE 3: X = 62
       CASE 4: X = 59
     END SELECT
     MID$(PLINE$, NMOFST, 255) = RFLBL$(X)
     'CALL PrintLINE: IF A = 27 THEN 2790
     Y = 0
2720 'IF Y > 0 THEN MID$(PLINE$, NMOFST, SPW) = LINESEP3$
     CALL WriteLINE: IF A = 27 THEN 2790
     ComPTR = 0: APTR = 0: LSET OTHDATA$ = STRING$(LEN(OTHDATA$), FC)
     IF OA(OT) <> 0 THEN
        CALL FF3GetRec(OA(OT))
        ComPTR = FF3PNTRS.COM: APTR = FF3PNTRS.LOC
        MID$(OTHDATA$, 1, 11) = XRPTDate$(FF3PNTRS.D1MD, FF3PNTRS.D1Y, DS)
        MID$(OTHDATA$, 12, 11) = XRPTDate$(FF3PNTRS.D2MD, FF3PNTRS.D2Y, DS)
        SELECT CASE OT
           CASE 1: IF MID$(FF3HEALTH.HLTHDIAG, 1, 1) <> " " THEN MID$(OTHDATA$, 23, 22) = RTRIM$(FF3HEALTH.HLTHDIAG)
                   IF MID$(FF3HEALTH.HLTHSTAT, 1, 1) <> " " THEN MID$(OTHDATA$, 45, 8) = RTRIM$(FF3HEALTH.HLTHSTAT)
           CASE 2: IF MID$(FF3EDUC.EDLEV, 1, 1) <> " " THEN MID$(OTHDATA$, 23, 8) = RTRIM$(FF3EDUC.EDLEV)
                   IF MID$(FF3EDUC.EDDEG, 1, 1) <> " " THEN MID$(OTHDATA$, 31, 4) = RTRIM$(FF3EDUC.EDDEG)
                   IF MID$(FF3EDUC.EDSUB1, 1, 1) <> " " THEN MID$(OTHDATA$, 35, 9) = RTRIM$(FF3EDUC.EDSUB1)
                   IF MID$(FF3EDUC.EDSUB2, 1, 1) <> " " THEN MID$(OTHDATA$, 44, 9) = RTRIM$(FF3EDUC.EDSUB2)
           CASE 3: IF MID$(FF3WORK.WORKTP, 1, 1) <> " " THEN MID$(OTHDATA$, 23, 30) = RTRIM$(FF3WORK.WORKTP)
           CASE 4: IF MID$(FF3MIL.MILRANK, 1, 1) <> " " THEN MID$(OTHDATA$, 23, 22) = RTRIM$(FF3MIL.MILRANK)
                   IF MID$(FF3MIL.MILSTAT, 1, 1) <> " " THEN MID$(OTHDATA$, 45, 8) = RTRIM$(FF3MIL.MILSTAT)
        END SELECT
     END IF

     ' Print Information Formatted in OTHDATA$
     PAGE.LOFST = NMOFST + 2: PAGE.COFST = PAGE.LOFST: X = 1
     IF OT > 1 THEN
        IF (MID$(OTHDATA$, 1, 1) <> " ") THEN
           MID$(CL$, 2, 255) = RFLBL$(47): X = 3 + RFLTH(47)
           MID$(CL$, X, 11) = MID$(OTHDATA$, 1, 11)
           X = INSTR(X, CL$, "  ")
        END IF
        IF (MID$(OTHDATA$, 12, 1) <> " ") THEN
           MID$(CL$, X, 255) = RFLBL$(48): X = X + 1 + RFLTH(48)
           MID$(CL$, X, 11) = MID$(OTHDATA$, 12, 11)
           X = INSTR(X, CL$, "  ")
        END IF
     END IF
     SELECT CASE OT
        CASE 1: ' Medical Information
                IF (MID$(OTHDATA$, 23, 1) <> " ") THEN
                   MID$(CL$, X, 255) = RFLBL$(56): X = X + RFLTH(56)
                   MID$(CL$, X + 1, 22) = MID$(OTHDATA$, 23, 22)
                   X = INSTR(X, CL$, "  ")
                END IF
                IF (MID$(OTHDATA$, 45, 1) <> " ") THEN
                   MID$(CL$, X, 255) = RFLBL$(57): X = X + RFLTH(57)
                   MID$(CL$, X + 1, 8) = MID$(OTHDATA$, 45, 8)
                   X = INSTR(X, CL$, "  ")
                END IF
                IF (MID$(OTHDATA$, 12, 1) <> " ") THEN
                   MID$(CL$, X, 255) = RFLBL$(58): X = X + RFLTH(58)
                   MID$(CL$, X + 1, 11) = MID$(OTHDATA$, 12, 11)
                END IF
        CASE 2: ' Education Information
                MID$(CL$, X, 255) = RFLBL$(51): X = X + RFLTH(51)
                IF (MID$(OTHDATA$, 35, 1) <> " ") THEN
                   MID$(CL$, X + 1, 9) = MID$(OTHDATA$, 35, 9)
                   X = INSTR(X, CL$, "  ")
                END IF
                IF (MID$(OTHDATA$, 44, 1) <> " ") THEN
                   MID$(CL$, X, 255) = RFLBL$(52): X = X + RFLTH(52)
                   MID$(CL$, X + 1, 9) = MID$(OTHDATA$, 44, 9)
                   X = INSTR(X, CL$, "  ")
                END IF
                IF MID$(OTHDATA$, 23, 1) <> " " THEN
                   MID$(CL$, X, 255) = RFLBL$(50): X = X + RFLTH(50)
                   MID$(CL$, X + 1, 8) = MID$(OTHDATA$, 23, 8)
                   X = INSTR(X, CL$, "  ")
                   MID$(CL$, X, 255) = RFLBL$(66): X = X + RFLTH(66)
                END IF
                IF MID$(OTHDATA$, 31, 1) <> " " THEN
                   MID$(CL$, X, 255) = RFLBL$(53): X = X + RFLTH(53)
                   MID$(CL$, X + 1, 4) = MID$(OTHDATA$, 31, 4)
                END IF
        CASE 3: ' Occupation Information
                IF MID$(OTHDATA$, 23, 1) <> " " THEN
                   MID$(CL$, X, 255) = RFLBL$(63): X = X + RFLTH(63)
                   MID$(CL$, X + 1, 30) = MID$(OTHDATA$, 23, 30)
                END IF
        CASE 4: ' Military Information
                IF (MID$(OTHDATA$, 23, 22) <> " ") THEN
                   MID$(CL$, X, 255) = RFLBL$(60): X = X + RFLTH(60)
                   MID$(CL$, X + 1, 22) = MID$(OTHDATA$, 23, 22)
                   X = INSTR(X, CL$, "  ")
                END IF
                IF (MID$(OTHDATA$, 45, 1) <> " ") THEN
                   MID$(CL$, X, 255) = RFLBL$(61): X = X + RFLTH(61)
                   MID$(CL$, X + 1, 8) = MID$(OTHDATA$, 45, 8)
                END IF
     END SELECT
     CALL PutDATA
     Y = Y + 1: IF OA(OT) <> 0 THEN OA(OT) = FF3PNTRS.NXT
     IF MID$(IA$, 10, 1) = "1" THEN
        GOSUB PrintCOM: IF A = 27 THEN 2790
     END IF
     IF MID$(IA$, 9, 1) = "1" THEN
        IF (APTR <> 0) THEN
           AX = 0
           WHILE APTR <> 0
              AX = AX + 1: GOSUB PrintRESI
              IF A = 27 THEN APTR = 0 ELSE APTR = FF2PFX.NXT
           WEND
           IF A = 27 THEN 2790
        END IF
     END IF
     IF OA(OT) <> 0 THEN 2720
     IF Y < OD THEN 2720 ELSE 2795
2790 OT = 99
2795 RETURN

END SUB

REM $STATIC
SUB GetNextMARRIAGE
SHARED SPReSeq, MLIST(), SPMCNT, SC, RID
   IF SPReSeq THEN
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

SUB LINEAGE (OPT$)
SHARED RELATE(), RELPTR&(), IDLIST(), LINTAB$()
SHARED ShowREL, ShowChNUM, ShowRegNUM, ShowPNAME, ExtAdoptRule
SHARED LinCNT, NMOFST, RID, Z

     ON LOCAL ERROR GOTO 2900
     SELECT CASE OPT$
        CASE "BUILD": GOSUB 62300
        CASE "PRINT": GOSUB 62350
     END SELECT
     ON LOCAL ERROR GOTO 0
     EXIT SUB

62300 ' Build LINEAGE Table
      X = RELATE(RID): LinCNT = 0
      IF X = 0 THEN EXIT SUB
      AGL = -(X \ 256): DGL = X AND 127: RGL = AGL + DGL
      ExtAdopt = ((X AND 128) > 0)
      'CALL RwrkGET("RELREC", RC, RELPTR&(RID))
      IF ASC(WFHDR.AMAXLV) > 0 THEN
         LinCNT = LinCNT + 1: Y$ = ""
         CS = WFANCREC.PSTAT
         IF NOT ExtAdoptRule THEN CS = CS AND 3
         IF CS > 0 THEN Y$ = MID$("*? *", CS, 1)
         LINTAB$(DGL + 1) = "[" + LTRIM$(STR$(WFANCREC.LINEAGE)) + Y$ + "]"
      END IF
      IF DGL > 0 THEN
         IF ShowChNUM AND (ShowRegNUM OR ShowPNAME) THEN P1$ = "."
         IF ShowPNAME AND (ShowRegNUM) THEN P2$ = " "
         X$ = "("
         FOR X = 1 TO DGL
            IF ShowChNUM THEN CH$ = LTRIM$(ROMAN$(ASC(WFRELREC.CHNUM)))
            CS = (WFRELREC.CSTAT AND 3)
            IF NOT ExtAdoptRule THEN CS = CS AND 3
            GET #5, WFRELREC.PARPTR + 1, WFRELREC
            IF ShowRegNUM THEN RN$ = FNX$(ABS(IDLIST(ABS(WFRELREC.ID))))
            IF X > 1 THEN X$ = ","
            IF X = DGL THEN Z$ = ")"
            IF CS > 0 THEN Y$ = MID$("*? *", CS, 1) ELSE Y$ = ""
            IF ShowPNAME THEN
               CALL FF1GetRec(ABS(WFRELREC.ID))
               L = INSTR(FF1REC.GIVEN, " ") - 1
               IF L > 0 THEN N$ = LEFT$(FF1REC.GIVEN, L) ELSE N$ = ""
            END IF
            LINTAB$(X) = X$ + Y$ + CH$ + P1$ + RN$ + P2$ + N$ + Z$
         NEXT X
         LinCNT = LinCNT + DGL
      END IF
      IF FF1REC.RID <> RID THEN CALL FF1GetRec(RID)
      RETURN

62350 ' Format Lineage
      IF LinCNT > 0 THEN
         X = RELATE(RID)
         AGL = -(X \ 256): DGL = X AND 127: RGL = AGL + DGL
         IF ShowREL THEN
            Z = PAGE.COFST
            MID$(CL$, Z, 255) = RFLBL$(64)
            Z = Z + RFLTH(64) + 1
            ExtAdopt = ((X AND 128) > 0)
            IF ExtAdopt THEN MID$(CL$, Z, 1) = "*": Z = Z + 1
            IF FF1REC.RID <> RID THEN CALL FF1GetRec(RID)
            MID$(CL$, Z, 12) = Relation$("M", FF1REC.SEX, AGL, RGL, RRELCD$())
            Z = INSTR(Z, CL$, "  ") + 2
            MID$(CL$, Z, 255) = RFLBL$(65): Z = Z + RFLTH(65)
            ELSE
            CALL PutDATA: Z = 2
         END IF
         FOR X = 1 TO LinCNT
             MID$(CL$, Z, 255) = LINTAB$(X): CALL PutDATA: Z = 1
             IF A = 27 THEN X = LinCNT
         NEXT X
         IF ShowREL THEN CALL PrintLINE
      END IF
      RETURN

2900 ' Error Routine
     SYSVAR.NVAR1 = ERL
     CALL ERRMessage("XXXXM098", ERR, ERL)
     RESUME 2910
2910 ' END ON ERROR
     EXIT SUB

END SUB

FUNCTION ROMAN$ (A)
SHARED RomanCAPS
   U = A MOD 10
   IF RomanCAPS THEN
      SELECT CASE U
         CASE 1: U$ = "I"
         CASE 2: U$ = "II"
         CASE 3: U$ = "III"
         CASE 4: U$ = "IV"
         CASE 5: U$ = "V"
         CASE 6: U$ = "VI"
         CASE 7: U$ = "VII"
         CASE 8: U$ = "VIII"
         CASE 9: U$ = "IX"
      END SELECT
      ELSE
      SELECT CASE U
         CASE 1: U$ = "i"
         CASE 2: U$ = "ii"
         CASE 3: U$ = "iii"
         CASE 4: U$ = "iv"
         CASE 5: U$ = "v"
         CASE 6: U$ = "vi"
         CASE 7: U$ = "vii"
         CASE 8: U$ = "viii"
         CASE 9: U$ = "iX"
      END SELECT
   END IF
   T = ((A - U) MOD 100) / 10
   SELECT CASE T
      CASE 1: T$ = "X"
      CASE 2: T$ = "XX"
      CASE 3: T$ = "XXX"
      CASE 4: T$ = "XL"
      CASE 5: T$ = "L"
      CASE 6: T$ = "LX"
      CASE 7: T$ = "LXX"
      CASE 8: T$ = "XXC"
      CASE 9: T$ = "XC"
   END SELECT
   IF A > 100 THEN Y$ = "C"
   X$ = SPACE$(10): RSET X$ = Y$ + T$ + U$
   ROMAN$ = RIGHT$(X$, 7)
END FUNCTION

