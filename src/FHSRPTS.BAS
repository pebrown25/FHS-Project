DECLARE SUB BuildTABLES ()
DECLARE SUB ProcessFLDS ()
DECLARE SUB ProcessHTFS ()
DECLARE SUB ProcessOPTS (RNAME$)
DECLARE SUB ProcessTABS ()
DECLARE SUB ProcessVARS ()
DECLARE SUB RPTSFile ()
DECLARE SUB RPTSReport ()
REM $DYNAMIC
1  REM $INCLUDE: 'FHSCOMON.BAS'
   PN$ = " FAMILY HISTORY SYSTEM - FHSRPTS - Report Customization program"
   CY$ = " (C) Copyright 1995,1996 by Phillip E. Brown"
   IF CONFIG.BP = 0 THEN RUN "FHSINIT"
   FUN = 0: PGMRTRN$ = SPACE$(7)
   FMTNAME$ = SPACE$(8): TNAME$ = SPACE$(8)
   IF ENV.PGMRTRN <> "FHSCUST" THEN
      ENTRptIX = ENV.RPTIX
      IF (ENTRptIX = 0) OR (ENTRptIX > RPTMAX) THEN ENV.PGMRTRN = "FHSCUST"
   END IF
   OEDMAST = ENV.EDMAST
   DIM RDFINDX AS ReportINDEX
   DIM RTABTable AS TablePARMS, RTABDATA$(1), RTABNM$(1), RTABCHN(1) AS ChainPTRS
   DIM RFGTable AS TablePARMS, RFLDGRP$(1), RFGCHN(1) AS ChainPTRS
   DIM RFFTable AS TablePARMS, RFLDTAB$(1), RFFCHN(1) AS ChainPTRS, RFFBGN(1)
   DIM RMONTable AS TablePARMS, RMCHN(1) AS ChainPTRS, RMONTH$(1)
   DIM RGENTable AS TablePARMS, RGCHN(1) AS ChainPTRS, RGENDER$(1)
   DIM RRELTable AS TablePARMS, RRCHN(1) AS ChainPTRS, RRELCODE$(1)
   DIM RREFTable AS TablePARMS, RREFCHN(1) AS ChainPTRS, RREFSYM$(1)
   DIM REVTTable AS TablePARMS, REVTCHN(1) AS ChainPTRS, REVTCODE$(1)
   DIM RTable AS TablePARMS, REPORT$(1), RPTCHN(1) AS ChainPTRS
   DIM RFTable AS TablePARMS, RPTFLD$(1)
   DIM RVARTable AS TablePARMS, RVARTAB$(1)
   DIM HTFTable AS TablePARMS
   DIM RPTDef AS ReportDefREC, RO$, RBGN, REND
   DIM XOPTTable AS TablePARMS
   DIM NEWTABLE AS TablePARMS
   DIM NewReport AS ReportDefREC
   HTFTable.TABNAME = "HTFLINES"
   HTFTable.FMTNAME = "RPTSS022"
   HTFTable.KEYNAME = "KEY"
   HTFTable.XL = 131: HTFTable.ML = HTFTable.XL
   XOPTTable = TABLE(ENV.OPTTAB)
   XOPTTable.FMTNAME = "XXXXO003": XOPTTable.FMTNUM = 0
   NewReport.PRTMSG = "M015"
   NewReport.WAITMSG = "M016"
   NewReport.PAGEMSG = "M017"
   NewReport.INTMSG = "M018"
   NewReport.ENDMSG = "M019"
   IF RDFHdr.FTYPE = "RDF" THEN
      CALL RptLOAD("ALL"): IF NOT OKAY THEN GOTO 1000
      ELSE
      IF RDFHdr.FTYPE = "CHG" THEN
         FOR X = 1 TO RPTMAX
             IF NOT (RDFIndex(X).CHG < 0) THEN
                CALL RptLOAD(RDFIndex(X).CODE)
             END IF
         NEXT X
      END IF
   END IF
   HTFTable.FF = HTFFree
   CLSCRN = -1: CALL BuildTABLES
   X = TABLE(ENV.OPTTAB).FT: FOUND = 0
   WHILE (X > 0) AND (NOT FOUND)
      FOUND = (MID$(TABDATA$(X), 1, 4) = "-GO-")
      IF NOT FOUND THEN X = TABDATACHN(X).FWD
   WEND
   IF FOUND THEN GOPTHEAD$ = " " + MID$(TABDATA$(X), 10, 30)

25 ' Open Windows for Program
   IF CLSCRN THEN CALL WinCLR: CLSCRN = 0
   CALL WinINIT("UH2")
   CALL WinINIT("LL2")
   CALL WinINIT("LR2")
   CALL WinINIT("LH2")
   CALL WinSWITCH(2)

40 ' Get FORMAT Numbers
   CALL FmtFIND("RPTSS001"): FMT001 = CURFMT
   CALL FmtFIND("RPTSS002"): FMT002 = CURFMT
   CALL FmtFIND("RPTSS003"): FMT003 = CURFMT
   CALL FmtFIND("RPTSS004"): FMT004 = CURFMT
   CALL FmtFIND("RPTSS005"): FMT005 = CURFMT
   CALL FmtFIND("RPTSS006"): FMT006 = CURFMT
   CALL FmtFIND("RPTSS007"): FMT007 = CURFMT
   CALL FmtFIND("RPTSS008"): FMT008 = CURFMT
   CALL FmtFIND("RPTSS009"): FMT009 = CURFMT
   CALL FmtFIND("RPTSS010"): FMT010 = CURFMT
   CALL FmtFIND("RPTSS011"): FMT011 = CURFMT
   CALL FmtFIND("RPTSS012"): FMT012 = CURFMT
   CALL FmtFIND("RPTSS013"): FMT013 = CURFMT
   CALL FmtFIND("RPTSS014"): FMT014 = CURFMT
   CALL FmtFIND("RPTSS015"): FMT015 = CURFMT
   CALL FmtFIND("RPTSS016"): FMT016 = CURFMT
   CALL FmtFIND("RPTSS017"): FMT017 = CURFMT
   CALL FmtFIND("RPTSS018"): FMT018 = CURFMT
   CALL FmtFIND("RPTSS020"): FMT020 = CURFMT
   CALL FmtFIND("RPTSS021"): FMT021 = CURFMT
   CALL FmtFIND("RPTSS022"): FMT022 = CURFMT
   CALL FmtFIND("RPTSS024"): FMT024 = CURFMT
   CALL FmtFIND("RPTSS026"): FMT026 = CURFMT
   CALL FmtFIND("RPTSS050"): FMT050 = CURFMT
   CALL FmtFIND("RPTSS090"): FMT090 = CURFMT
   CALL FmtFIND("XXXXS090"): FMTX90 = CURFMT
   CALL FmtFIND("RPTSS091"): FMT091 = CURFMT
   CALL FmtFIND("RPTSS092"): FMT092 = CURFMT
   CALL FmtFIND("RPTSS093"): FMT093 = CURFMT
   CALL FmtFIND("XXXXS093"): FMTX93 = CURFMT
   CALL FmtFIND("RPTSS094"): FMT094 = CURFMT
   CALL FmtFIND("RPTSS095"): FMT095 = CURFMT
   CALL FmtFIND("XXXXS095"): FMTX95 = CURFMT
   CALL FmtFIND("XXXXS096"): FMTX96 = CURFMT

60 ' Format Display
   IF CLSCRN THEN CALL WinCLR: CLSCRN = 0
   LSET PGMRTRN$ = ENV.PGMRTRN
   IF PGMRTRN$ = "FHSCUST" THEN GOTO 200

   ' Setup to Process Report for Calling Program
   RTable.CT = ENTRptIX
   RN = RDFIndex(RTable.CT).DEFIX
   IF RN > 0 THEN
      RDFReport = RDFReports(RN)
      ELSE
      X = RDFHdr.REPORTS + 1
      REDIM PRESERVE RDFReports(X) AS ReportDefREC, RDFCHN(X) AS ChainPTRS
      RDFHdr.REPORTS = X
      RN = X: RDFIndex(RTable.CT).DEFIX = RN
      RDFReport = NewReport
      RDFReport.CODE = RDFIndex(RTable.CT).CODE
      RDFReport.NAME = MID$(REPORT$(RTable.CT), 21, 255)
      RDFReports(RN) = RDFReport
    END IF
    GOTO 305

' Routines for Processing RDF Header and REPORT Data in Format Fields
'   There are separate routines for File to Screen and Screen to File

100 '  Move Report Definition File HEADER DATA to Format
    CALL WinSWITCH(2)
    CURFMT = FMT001: FMT = FMTTAB(CURFMT)
    X = FMT.BGNFLD
    WHILE X > 0
       FFLD = FLDTAB(X)
       IF FFLD.UTXT <> 0 THEN
          X$ = UTXT$(FFLD.UTXT)
          SELECT CASE FFLD.NAME
             CASE "RDF ": LSET X$ = CONFIG.RDF
             CASE "SIZE": LSET X$ = LTRIM$(STR$(RDFHdr.EOFRBA))
             CASE "RCNT": LSET X$ = LTRIM$(STR$(RDFHdr.REPORTS))
             CASE "TCNT": LSET X$ = LTRIM$(STR$(RDFHdr.TABLES))
             CASE "FCNT": LSET X$ = LTRIM$(STR$(RDFHdr.FLDTOTL))
             CASE "VCNT": LSET X$ = LTRIM$(STR$(RDFHdr.VARTOTL))
             CASE "LCNT": LSET X$ = LTRIM$(STR$(RDFHdr.HTFTOTL))
             CASE "OCNT": LSET X$ = LTRIM$(STR$(RDFHdr.OPTTOTL))
             CASE "GVAR": LSET X$ = LTRIM$(STR$(RDFHdr.VARIABLE))
             CASE "GFLD": LSET X$ = LTRIM$(STR$(RDFHdr.FIELDS))
             CASE "GOPT": LSET X$ = LTRIM$(STR$(RDFHdr.OPTIONS))
          END SELECT
          LSET UTXT$(FFLD.UTXT) = X$
       END IF
       X = FLDCHN(X).FWD
    WEND
    CALL WinFORMAT(FMT001)
    RETURN

110 ' Move Report Parms to Format
    CURFMT = FMT020: FMT = FMTTAB(CURFMT)
    X = FMT.BGNFLD
    WHILE X > 0
       FFLD = FLDTAB(X)
       IF FFLD.UTXT <> 0 THEN
          X$ = UTXT$(FFLD.UTXT)
          SELECT CASE FFLD.NAME
             CASE "RDF ": LSET X$ = CONFIG.RDF
             CASE "CODE": LSET X$ = RDFReport.CODE
             CASE "DESC": LSET X$ = RDFReport.NAME
             CASE "RFLD": LSET X$ = LTRIM$(STR$(RDFReport.FLDDEFS))
             CASE "RVAR": LSET X$ = LTRIM$(STR$(RDFReport.VARDEFS))
             CASE "RHTF": LSET X$ = LTRIM$(STR$(RDFReport.HTFCNT))
             CASE "ROPT": LSET X$ = LTRIM$(STR$(RDFReport.OPTCNT))
             CASE "NAME": LSET X$ = RDFReport.NAME
             CASE "TYPE": LSET X$ = LTRIM$(STR$(RDFReport.TYPE))
             CASE "PWTH": RSET X$ = LTRIM$(STR$(RDFReport.PWIDTH))
             CASE "BLHD": RSET X$ = LTRIM$(STR$(RDFReport.BLHEAD))
             CASE "BLTI": RSET X$ = LTRIM$(STR$(RDFReport.BLTITLE))
             CASE "BLFT": RSET X$ = LTRIM$(STR$(RDFReport.BLFOOT))
             CASE "BLDL": RSET X$ = LTRIM$(STR$(RDFReport.BLBODY))
             CASE "PMSG": LSET X$ = RDFReport.PRTMSG
             CASE "WMSG": LSET X$ = RDFReport.WAITMSG
             CASE "PGMS": LSET X$ = RDFReport.PAGEMSG
             CASE "IMSG": LSET X$ = RDFReport.INTMSG
             CASE "EMSG": LSET X$ = RDFReport.ENDMSG
          END SELECT
          LSET UTXT$(FFLD.UTXT) = X$
       END IF
       X = FLDCHN(X).FWD
    WEND
    CALL WinDisplayDATA(FMT020)
    RETURN

115 ' Move Report Parms to RDFreport
    CURFMT = FMT020: FMT = FMTTAB(CURFMT)
    X = FMT.BGNFLD
    WHILE X > 0
       FFLD = FLDTAB(X)
       IF FFLD.UTXT <> 0 THEN
          X$ = UTXT$(FFLD.UTXT)
          SELECT CASE FFLD.NAME
             CASE "NAME": LSET RDFReport.NAME = X$
             CASE "TYPE": RDFReport.TYPE = VAL(X$)
             CASE "PWTH": RDFReport.PWIDTH = VAL(X$)
             CASE "BLHD": RDFReport.BLHEAD = VAL(X$)
             CASE "BLTI": RDFReport.BLTITLE = VAL(X$)
             CASE "BLFT": RDFReport.BLFOOT = VAL(X$)
             CASE "BLDL": RDFReport.BLBODY = VAL(X$)
             CASE "PMSG": LSET RDFReport.PRTMSG = X$
             CASE "WMSG": LSET RDFReport.WAITMSG = X$
             CASE "PGMS": LSET RDFReport.PAGEMSG = X$
             CASE "IMSG": LSET RDFReport.INTMSG = X$
             CASE "EMSG": LSET RDFReport.ENDMSG = X$
          END SELECT
       END IF
       X = FLDCHN(X).FWD
    WEND
    RDFReports(RN) = RDFReport
    GOSUB 110
    RETURN

200 ' Get Processing Option
    CALL WinSWITCH(2)
    IF WIN.WFMT <> FMT001 THEN
       CALL WinCLR
       CALL WinDisplayLABELS(FMT001)
    END IF
    GOSUB 100
210 ' Get Request
    CALL PutMSG(" ")
    CALL OptDISPLAY(FMT090)
    CALL GetKEY
    SELECT CASE A
       CASE 59: FUN$ = "F1"
       CASE 60: FUN$ = "F2"
       CASE 61: FUN$ = "F3"
       CASE 62: FUN$ = "F4"
       CASE 63: FUN$ = "F5"
       CASE 64: FUN$ = "F6"
       CASE 67: FUN$ = "F9"
       CASE ELSE: CALL ErrBEEP(0): GOTO 210
    END SELECT
    CALL OptHILITE(FUN$, "", 6)
    SELECT CASE FUN$
       CASE "F1": CALL RPTSFile     ' Process FILE Request
       CASE "F2": GOTO 300   ' Select REPORT
       CASE "F3": GOTO 250
       CASE "F6": RO$ = "": CALL RPTSReport' Process PRINT Request
       CASE "F9": GOTO 1000  ' Return
    END SELECT
    GOTO 200

250 ' Process GLOBAL Parms Requests
    CALL PutMSG("RPTSM011")
    IF A$ = CHR$(27) THEN GOTO 200
    SELECT CASE A$
       CASE "1": CALL ProcessTABS                   ' Process TABLEs
       CASE "2": RVARTable.FT = RDFIndex(0).VARBGN
                 CALL ProcessVARS                   ' Process VARIABLEs
                 RDFIndex(0).VARBGN = RVARTable.FT
                 RDFHdr.VARIABLE = RVARTable.SIZE
       CASE "3": RFTable.FT = RDFIndex(0).FLDBGN
                 CALL ProcessFLDS                   ' Process FIELDs
                 RDFIndex(0).FLDBGN = RFTable.FT
                 RDFHdr.FIELDS = RFTable.SIZE
       CASE "4": OPTTable.FT = RDFIndex(0).OPTBGN: ENV.RPTIX = 0
                 CALL ProcessOPTS(GOPTHEAD$)        ' Process OPTIONS
                 RDFIndex(0).OPTBGN = OPTTable.FT
                 RDFHdr.OPTIONS = OPTTable.SIZE
                 RPTOPTION = RPTOPTS(0)
                 Y = OPTTable.FT: X$ = SPACE$(6)
                 WHILE Y > 0
                    LSET X$ = RPTOPT$(Y)
                    CALL RptOptSET(X$, RPTOPTION)
                    Y = OPTCHN(Y).FWD
                 WEND
                 RPTOPTS(0) = RPTOPTION
       CASE ELSE: GOTO 250
    END SELECT
    GOTO 200

300 ' Process REPORT$() Table
    CALL WinSWITCH(5)
    IF WIN.WFMT <> FMT002 THEN
       CALL WinCLR
       CALL WinFORMAT(FMT002)
       CALL TableOPEN(RTable)
    END IF
    CALL TableSHOW(RTable, REPORT$(), RPTCHN())
301 ' Select REPORT from Table
    CALL OptDISPLAY(FMTX90)
    CALL TableSELECT(RTable, REPORT$(), RPTCHN())
    CALL WinCLR
    IF A = 27 THEN GOTO 200
    RN = RDFIndex(RTable.CT).DEFIX
    IF RN > 0 THEN
       RDFReport = RDFReports(RN)
       ELSE
       X = RDFHdr.REPORTS + 1
       REDIM PRESERVE RDFReports(X) AS ReportDefREC, RDFCHN(X) AS ChainPTRS
       RDFHdr.REPORTS = X
       RN = X: RDFIndex(RTable.CT).DEFIX = RN
       RDFReport = NewReport
       RDFReport.CODE = RDFIndex(RTable.CT).CODE
       RDFReport.NAME = MID$(REPORT$(RTable.CT), 21, 255)
       RDFReports(RN) = RDFReport
    END IF

305 ' Display Report Info
    CALL WinSWITCH(2)
    IF WIN.WFMT <> FMT020 THEN
       CALL WinCLR
       CALL WinDisplayLABELS(FMT020)
       CALL OptDISPLAY(FMT093)
    END IF
    GOSUB 110
310 ' Get Report Request
    CALL PutMSG("")
    CALL OptDISPLAY(FMT093)
    CALL GetKEY
    SELECT CASE A
       CASE 59: FUN$ = "F1"
       CASE 60: FUN$ = "F2"
       CASE 61: FUN$ = "F3"
       CASE 62: FUN$ = "F4"
       CASE 63: FUN$ = "F5"
       CASE 64: FUN$ = "F6"
       CASE 27: FUN$ = "ESC"
       CASE ELSE: CALL ErrBEEP(0): GOTO 310
    END SELECT
    CALL OptHILITE(FUN$, "", 6)
    SELECT CASE FUN$
       CASE "F1": GOTO 320   ' Change Report Description Info
       CASE "F2": OPTTable.FT = RDFIndex(RTable.CT).OPTBGN
                  CALL ProcessOPTS(RDFReport.NAME)  ' Update OPTIONS
                  RDFIndex(RTable.CT).OPTBGN = OPTTable.FT
                  RDFReport.OPTCNT = OPTTable.SIZE
                  GOTO 305
       CASE "F3": HTFTable.FT = RDFIndex(RTable.CT).HTFBGN
                  CALL ProcessHTFS  ' Update HTFLines
                  RDFIndex(RTable.CT).HTFBGN = HTFTable.FT
                  RDFReport.HTFCNT = HTFTable.SIZE
                  GOTO 305
       CASE "F4": RVARTable.FT = RDFIndex(RTable.CT).VARBGN
                  CALL ProcessVARS  ' Process VARIABLEs
                  RDFIndex(RTable.CT).VARBGN = RVARTable.FT
                  RDFReport.VARDEFS = RVARTable.SIZE
                  GOTO 305
       CASE "F5": RFTable.FT = RDFIndex(RTable.CT).FLDBGN
                  CALL ProcessFLDS  ' Update Field Defs
                  RDFIndex(RTable.CT).FLDBGN = RFTable.FT
                  RDFReport.FLDDEFS = RFTable.SIZE
                  GOTO 305
       CASE "F6": RO$ = "5" ' Process PRINT Request
                  RBGN = RTable.CT: REND = RBGN
                  CALL RPTSReport
                  RDFReport = RDFReports(RDFIndex(RTable.CT).DEFIX)
                  GOTO 305
       CASE "ESC": RDFReports(RN) = RDFReport
                   IF RDFReport.CHG = -1 THEN
                      RDFHdr.VER = CHR$(255)
                      C = RDFIndex(RTable.CT).CHG
                      IF C = 0 THEN C = -256 ELSE IF C > 0 THEN C = -C
                      RDFIndex(RTable.CT).CHG = C
                      X$ = REPORT$(RTable.CT)
                      MID$(X$, 8, 1) = "+"
                      RSET CNT$ = LTRIM$(STR$(RDFReport.OPTCNT))
                      MID$(X$, 9, 3) = CNT$
                      RSET CNT$ = LTRIM$(STR$(RDFReport.HTFCNT))
                      MID$(X$, 12, 3) = CNT$
                      RSET CNT$ = LTRIM$(STR$(RDFReport.VARDEFS))
                      MID$(X$, 15, 3) = CNT$
                      RSET CNT$ = LTRIM$(STR$(RDFReport.FLDDEFS))
                      MID$(X$, 18, 3) = CNT$
                      RSET REPORT$(RTable.CT) = X$
                   END IF
                   IF PGMRTRN$ <> "FHSCUST" THEN GOTO 1000
                   GOTO 300  ' Return
    END SELECT
    GOTO 305

320 ' Update Report Description
    CALL PutMSG("XXXXM012")
    FMT = FMTTAB(FMT020)
    CALL WinUPDATE
    IF A = 27 GOTO 305
    RDFReport.CHG = -1
    GOSUB 115
    GOTO 305

1000 ' Return to FHSMENU.BAS
     IF (RDFHdr.VER = CHR$(255)) THEN
        IF PGMRTRN$ = "FHSCUST" THEN
           CALL PutMSG("XXXXM020")
           IF A$ <> SNGLKEY$(1) THEN GOTO 200
        END IF
        IF RDFHdr.FTYPE = "RDF" THEN RDFHdr.FTYPE = "CHG"
     END IF
     CALL PutMSG("")
     CLOSE : A = 0
     CALL PgmPREP(PGMRTRN$)
     IF (NOT OKAY) OR (A = 27) THEN GOTO 200
     ENV.WINHILVL = 1
     IF PGMRTRN$ <> "FHSTABL" THEN
        CALL TableFREE("REPORTS")
        CALL TableFREE("RPTTABLE")
     END IF
     IF RDFHdr.FTYPE = "RDF" THEN CALL RptLOAD("INITIAL")
     CHAIN PGMRTRN$

SUB BuildTABLES
   SHARED RDFINDX AS ReportINDEX
   SHARED RTABTable AS TablePARMS, RTABDATA$(), RTABNM$(), RTABCHN() AS ChainPTRS
   SHARED RFGTable AS TablePARMS, RFLDGRP$(), RFGCHN() AS ChainPTRS
   SHARED RFFTable AS TablePARMS, RFLDTAB$(), RFFCHN() AS ChainPTRS, RFFBGN()
   SHARED RMONTable AS TablePARMS, RMCHN() AS ChainPTRS, RMONTH$()
   SHARED RGENTable AS TablePARMS, RGCHN() AS ChainPTRS, RGENDER$()
   SHARED RRELTable AS TablePARMS, RRCHN() AS ChainPTRS, RRELCODE$()
   SHARED RREFTable AS TablePARMS, RREFCHN() AS ChainPTRS, RREFSYM$()
   SHARED REVTTable AS TablePARMS, REVTCHN() AS ChainPTRS, REVTCODE$()
   SHARED RTable AS TablePARMS, REPORT$(), RPTCHN() AS ChainPTRS
   SHARED RFTable AS TablePARMS, RPTFLD$()
   SHARED RVARTable AS TablePARMS, RVARTAB$()
   SHARED HTFTable AS TablePARMS, HTFCHN() AS ChainPTRS
   SHARED RPTDef AS ReportDefREC, RO$, RBGN, REND
   SHARED XOPTTable AS TablePARMS
   SHARED NEWTABLE AS TablePARMS
   SHARED NewReport AS ReportDefREC
   ' Build Report TABLEs Tables:
   ' RMONTH$(), RGENDER$(), RRELCODE$(), RRELSYM$(), REVTCODE$()
   ON LOCAL ERROR GOTO 19
10 CALL TableLOAD("RPTTABLE", RTABTNUM, 0)
   RTABTable = TABLE(RTABTNUM)
   RTABTable.FMTNAME = "RPTSS005": RTABTable.FMTNUM = 0
   RTABTable.KEYNAME = "KEY"
   REDIM RTABDATA$(RTABTable.SIZE), RTABCHN(RTABTable.SIZE) AS ChainPTRS
   REDIM RTABNM$(RTABTable.SIZE)
   CALL ChnINIT(RTABCHN(), 1, RTABTable.SIZE)
   Y = RTABTable.FT: X = 0
   WHILE Y > 0
       X = X + 1
       RTABNM$(X) = MID$(TABDATA$(Y), 1, 8)
       RTABDATA$(X) = MID$(TABDATA$(Y), 9, 33)
       Y = TABDATACHN(Y).FWD
   WEND
   RTABTable.XL = 48: RTABTable.ML = RTABTable.XL
   RTABTable.FT = 1: RTABTable.LT = RTABTable.SIZE
11 CALL TableLOAD("EVENTS", REVTTNUM, 0)
   REVTTable = TABLE(REVTTNUM)
   REDIM REVTCODE$(REVTTable.SIZE), REVTCHN(REVTTable.SIZE) AS ChainPTRS
   EVTABSZ = VAL(REVENT$(0))
   NEWTAB = (EVTABSZ <> REVTTable.SIZE)
   'IF NEWTAB THEN REDIM REVENT$(REVTTable.SIZE)
   X = 0: Y = REVTTable.FT
   WHILE Y > 0
       X = X + 1: REVTCODE$(X) = SPACE$(48)
       Z = 0: FOUND = 0
       WHILE (Z < EVTABSZ) AND NOT FOUND
          Z = Z + 1
          FOUND = (RTRIM$(LEFT$(REVENT$(Z), 8))) = RTRIM$(LEFT$(TABDATA$(Y), 8))
       WEND
       IF FOUND THEN
          LSET REVTCODE$(X) = REVENT$(Z)
          ELSE
          MID$(REVTCODE$(X), 1, 8) = LEFT$(TABDATA$(Y), 8)
          MID$(REVTCODE$(X), 9, 8) = LEFT$(TABDATA$(Y), 8)
       END IF
       MID$(REVTCODE$(X), 24, 25) = MID$(TABDATA$(Y), 13, 25)
       Y = TABDATACHN(Y).FWD
   WEND
   IF NEWTAB THEN
      REDIM REVENT$(REVTTable.SIZE)
      FOR X = 1 TO REVTTable.SIZE
         REVENT$(X) = RTRIM$(LEFT$(REVTCODE$, 23))
      NEXT X
      REVENT$(0) = STR$(REVTTable.SIZE)
   END IF
   CALL ChnINIT(REVTCHN(), 1, REVTTable.SIZE)
   REVTTable.FT = 1: REVTTable.LT = REVTTable.SIZE
   REVTTable.FMTNAME = "RPTSS027": REVTTable.FMTNUM = 0
   CALL TableFREE("EVENTS")
12 CALL TableLOAD("MONTHS", RMONTNUM, 0)
   RMONTable = TABLE(RMONTNUM)
   REDIM RMONTH$(RMONTable.SIZE), RMCHN(RMONTable.SIZE) AS ChainPTRS
   NEWTAB = (VAL(RMONTB$(0)) <> RMONTable.SIZE)
   IF NEWTAB THEN REDIM RMONTB$(RMONTable.SIZE)
   X = 0: Y = RMONTable.FT
   WHILE Y > 0
       X = X + 1: RMONTH$(X) = LEFT$(TABDATA$(Y), 5) + MID$(TABDATA$(Y), 3, 26)
       IF NEWTAB THEN
          RMONTB$(X) = LEFT$(RMONTH$(X), 5)
          ELSE
          MID$(RMONTH$(X), 1, 5) = RMONTB$(X)
       END IF
       Y = TABDATACHN(Y).FWD
   WEND
   CALL ChnINIT(RMCHN(), 1, RMONTable.SIZE)
   RMONTable.FT = 1: RMONTable.LT = RMONTable.SIZE
   RMONTable.FMTNAME = "RPTSS007": RMONTable.FMTNUM = 0
   RMONTable.KEYNAME = "KEY"
   CALL TableFREE("MONTHS")
13 CALL TableLOAD("GENDER", RGENTNUM, 0)
   RGENTable = TABLE(RGENTNUM)
   REDIM RGENDER$(RGENTable.SIZE), RGCHN(RGENTable.SIZE) AS ChainPTRS
   NEWTAB = (VAL(RGENDR$(0)) <> RGENTable.SIZE)
   IF NEWTAB THEN REDIM RGENDR$(RGENTable.SIZE)
   X = 0: Y = RGENTable.FT
   WHILE Y > 0
       X = X + 1
       IF NEWTAB THEN RGENDR$(X) = MID$(TABDATA$(Y), 1, 11)
       RGENDER$(X) = RGENDR$(X)
       Y = TABDATACHN(Y).FWD
   WEND
   CALL ChnINIT(RGCHN(), 1, RGENTable.SIZE)
   RGENTable.FT = 1: RGENTable.LT = RGENTable.SIZE
   RGENTable.FMTNAME = "RPTSS009": RGENTable.FMTNUM = 0
   RGENTable.KEYNAME = "KEY"
   CALL TableFREE("GENDER")
14 CALL TableLOAD("RELCODES", RRELTNUM, 0)
   RRELTable = TABLE(RRELTNUM)
   REDIM RRELCODE$(RRELTable.SIZE), RRCHN(RRELTable.SIZE) AS ChainPTRS
   NEWTAB = (VAL(RRELCD$(0)) <> RRELTable.SIZE)
   IF NEWTAB THEN REDIM RRELCD$(RRELTable.SIZE)
   X = 0: Y = RRELTable.FT
   WHILE Y > 0
       X = X + 1: RRELCODE$(X) = MID$(TABDATA$(Y), 1, 17) + MID$(TABDATA$(Y), 6, 12)
       IF NEWTAB THEN
          RRELCD$(X) = LEFT$(RRELCODE$(X), 17)
          ELSE
          MID$(RRELCODE$(X), 1, 17) = RRELCD$(X)
       END IF
       Y = TABDATACHN(Y).FWD
   WEND
   CALL ChnINIT(RRCHN(), 1, RRELTable.SIZE)
   RRELTable.FT = 1: RRELTable.LT = RRELTable.SIZE
   RRELTable.FMTNAME = "RPTSS011": RRELTable.FMTNUM = 0
   RRELTable.KEYNAME = "KEY"
   CALL TableFREE("RELCODES")
15 CALL TableLOAD("REFERENC", RREFTNUM, 0)
   RREFTable = TABLE(RREFTNUM)
   REDIM RREFSYM$(RREFTable.SIZE), RREFCHN(RREFTable.SIZE) AS ChainPTRS
   NEWTAB = (VAL(RXSYM$(0)) <> RREFTable.SIZE)
   IF NEWTAB THEN REDIM RXSYM$(RREFTable.SIZE)
   X = 0: Y = RREFTable.FT
   WHILE Y > 0
       X = X + 1: Z = VAL(MID$(TABDATA$(Y), 1, 2))
       IF (Z > 0) AND (Z <= RREFTable.SIZE) THEN
          RREFSYM$(Z) = MID$(TABDATA$(Y), 1, 2) + "   " + MID$(TABDATA$(Y), 5, 30)
          IF NEWTAB THEN
             RXSYM$(Z) = MID$(RREFSYM$(X), 3, 3)
             ELSE
             MID$(RREFSYM$(Z), 3, 3) = RXSYM$(Z)
          END IF
       END IF
       Y = TABDATACHN(Y).FWD
   WEND
   CALL ChnINIT(RREFCHN(), 1, RREFTable.SIZE)
   RREFTable.FT = -(RREFTable.SIZE > 0): RREFTable.LT = RREFTable.SIZE
   RREFTable.FMTNAME = "RPTSS025": RREFTable.FMTNUM = 0
   RREFTable.KEYNAME = "KEY"
   CALL TableFREE("REFERENC")

16 ' Build RFLDGRP$(),RFLDTAB$() Tables from "RPTFIELD" System Table
   CALL TableLOAD("RPTFIELD", RFLDTNUM, 0)
   FGMAX = 0
   Y = TABLE(RFLDTNUM).FT
   WHILE Y > 0
      IF MID$(TABDATA$(Y), 6, 1) = "G" THEN FGMAX = FGMAX + 1
      Y = TABDATACHN(Y).FWD
   WEND
   FFMAX = TABLE(RFLDTNUM).SIZE - FGMAX
   REDIM RFLDGRP$(FGMAX), RFGCHN(FGMAX) AS ChainPTRS, RFFBGN(FGMAX + 1)
   REDIM RFLDTAB$(FFMAX), RFFCHN(FFMAX) AS ChainPTRS
   RFGTable = NEWTABLE
   RFGTable.TABNAME = "RFLDGRPS"
   RFGTable.FMTNAME = "RPTSS017": RFGTable.FMTNUM = 0
   RFGTable.KEYNAME = "KEY"
   RFGTable.SIZE = FGMAX
   RFGTable.XL = 40: RFGTable.ML = RFGTable.XL
   RFGTable.FT = 1: RFGTable.LT = FGMAX
   RFGTable.FF = 0
   CALL ChnINIT(RFGCHN(), 1, FGMAX)
   RFFTable = NEWTABLE
   RFFTable.TABNAME = "RPTFLDS"
   RFFTable.FMTNAME = "RPTSS017": RFFTable.FMTNUM = 0
   RFFTable.KEYNAME = "KEY"
   RFFTable.SIZE = FFMAX
   RFFTable.XL = 40: RFFTable.ML = RFFTable.XL
   RFFTable.FF = 0
   CALL ChnINIT(RFFCHN(), 1, FFMAX)
   X$ = SPACE$(RFFTable.XL)
   W = TABLE(RFLDTNUM).FT: Y = 0: Z = 0
   FOR X = 1 TO TABLE(RFLDTNUM).SIZE
      LSET X$ = MID$(TABDATA$(W), 1, 5)
      MID$(X$, 6, 33) = MID$(TABDATA$(W), 7, 33)
      IF MID$(TABDATA$(W), 6, 1) = "G" THEN
         RFFCHN(Z).FWD = 0
         Y = Y + 1: RFLDGRP$(Y) = X$
         RFFBGN(Y) = Z + 1
         RFFCHN(Z + 1).BWD = 0
         ELSE
         Z = Z + 1: RFLDTAB$(Z) = X$
      END IF
      W = TABDATACHN(W).FWD
   NEXT X
   RFFBGN(FGMAX + 1) = FFMAX + 1

   ' Build RPTFLD$() from RFLDS(), RFLD$() & RFLDTAB$()
   RFTable = NEWTABLE
   RFTable.TABNAME = "RFLDTAB"
   RFTable.FMTNAME = "RPTSS015": RFTable.FMTNUM = 0
   RFTable.KEYNAME = "KEY"
   RFTable.XL = 95: RFTable.ML = RFTable.XL
   RFTable.FF = RFFree
   REDIM RPTFLD$(RFMAX)
   R = 0: X$ = SPACE$(RFTable.XL): Y$ = SPACE$(54)
   WHILE R <= RPTMAX
      Y = RDFIndex(R).FLDBGN
      WHILE Y > 0
         RDFFld = RFLDS(Y)
         W = ASC(RDFFld.RLTH)
         A = ASC(RDFFld.ALTH): L = ASC(RDFFld.LLTH): H = ASC(RDFFld.HLTH)
         LSET X$ = RDFFld.NAME: X = 1: LSET Y$ = RFLD$(Y)
         MID$(X$, 14, 2) = RIGHT$(STR$(W), 2)
         IF A > 0 THEN MID$(X$, 16, 4) = MID$(Y$, X, A): X = X + A
         IF L > 0 THEN MID$(X$, 20, 25) = MID$(Y$, X, L): X = X + L
         IF H > 0 THEN MID$(X$, 45, 25) = MID$(Y$, X, H)
         Z = RFFBGN(VAL(MID$(RDFFld.NAME, 1, 2))): FOUND = 0
         WHILE (Z > 0) AND (NOT FOUND)
            FOUND = (MID$(RFLDTAB$(Z), 1, 5) = RDFFld.NAME)
            IF NOT FOUND THEN Z = RFFCHN(Z).FWD
         WEND
         IF FOUND THEN
            MID$(X$, 6, 8) = MID$(RFLDTAB$(Z), 8, 8)
            IF W = 0 THEN
               MID$(X$, 14, 2) = MID$(RFLDTAB$(Z), 6, 2)
               RFLDS(Y).RLTH = CHR$(VAL(MID$(X$, 14, 2)))
            END IF
            MID$(X$, 70, 25) = MID$(RFLDTAB$(Z), 16, 25)
         END IF
         RPTFLD$(Y) = X$
         Y = RFLDCHN(Y).FWD
      WEND
      R = R + 1
   WEND
   CALL TableFREE("RPTFIELD")

17 ' Build RVARTAB$() Tables from RVARS(),RVAR$() System Tables
   RVARTable = NEWTABLE
   RVARTable.TABNAME = "RPTVARS"
   RVARTable.FMTNAME = "RPTSS013"
   RVARTable.KEYNAME = "KEY"
   RVARTable.XL = 97: RVARTable.ML = RVARTable.XL
   RVARTable.FF = RVFree
   REDIM RVARTAB$(RVMAX)
   Y = RVMAX
   WHILE Y > 0
      RVARTAB$(Y) = RVARS(Y).NAME + RVAR$(Y)
      Y = Y - 1
   WEND

   ' Build REPORT$() Table from "REPORTS" System Table and RDFIndex()
   CALL TableLOAD("REPORTS", RPTTNUM, 0)
   REDIM REPORT$(RPTMAX), RPTCHN(RPTMAX) AS ChainPTRS
   RTable = NEWTABLE
   RTable.TABNAME = "REPORTS"
   RTable.FMTNAME = "RPTSS003": RTable.FMTNUM = 0
   RTable.KEYNAME = "KEY"
   RTable.SIZE = RPTMAX
   RTable.XL = 64: RTable.ML = RTable.XL
   RTable.FT = 1: RTable.LT = RPTMAX
   RTable.FF = 0
   CALL ChnINIT(RPTCHN(), 1, RPTMAX)
   CNT$ = "   ": X$ = SPACE$(RTable.XL): FT = TABLE(RPTTNUM).FT
   FOR X = 1 TO RPTMAX
      RDFINDX = RDFIndex(X)
      LSET X$ = RDFINDX.CODE
      RSET CNT$ = LTRIM$(STR$(ABS(RDFINDX.CHG) AND 255))
      MID$(X$, 5, 3) = CNT$
      IF RDFINDX.CHG < 0 THEN MID$(X$, 8, 1) = "+"
      IF RDFINDX.DEFIX > 0 THEN
         RDFReport = RDFReports(RDFINDX.DEFIX)
         RSET CNT$ = LTRIM$(STR$(RDFReport.OPTCNT))
         MID$(X$, 9, 3) = CNT$
         RSET CNT$ = LTRIM$(STR$(RDFReport.HTFCNT))
         MID$(X$, 12, 3) = CNT$
         RSET CNT$ = LTRIM$(STR$(RDFReport.VARDEFS))
         MID$(X$, 15, 3) = CNT$
         RSET CNT$ = LTRIM$(STR$(RDFReport.FLDDEFS))
         MID$(X$, 18, 3) = CNT$
      END IF
      Y = FT: FOUND = 0
      WHILE (Y > 0) AND (NOT FOUND)
         FOUND = (RDFINDX.CODE = MID$(TABDATA$(Y), 1, 4))
         IF NOT FOUND THEN Y = TABDATACHN(Y).FWD
      WEND
      IF FOUND THEN MID$(X$, 21, 45) = MID$(TABDATA$(Y), 5, 45)
      REPORT$(X) = X$
   NEXT X
   OKAY = -1
18 'Return
   ON LOCAL ERROR GOTO 0
   EXIT SUB

19 'Error Routine
   CALL ErrMESSAGE("Error Building Tables...[EMSG]...", ERR, ERL)
   CALL GetRESP
   OKAY = 0
   RESUME 18

END SUB

REM $STATIC
SUB ProcessFLDS
SHARED RFTable AS TablePARMS, RFFTable AS TablePARMS, RFGTable AS TablePARMS
SHARED RPTFLD$(), RFLDGRP$(), RFGCHN() AS ChainPTRS, RFFBGN(), RFLDTAB$()
SHARED RFFCHN() AS ChainPTRS
SHARED FMT014, FMT016, FMT018, FMT091, FMTX93
400 ' Process FIELD Definitions
    CALL WinSWITCH(5)
    CALL WinFORMAT(FMT014)
    CALL OptDISPLAY(FMTX93)
    RFTable.CL = 0: RFTable.SIZE = 0
    X = RFTable.FT: RFTable.LT = 0
    WHILE X > 0
        RFTable.LT = X
        RFTable.SIZE = RFTable.SIZE + 1
        X = RFLDCHN(X).FWD
    WEND
    CALL TableOPEN(RFTable)
    IF RFTable.FT > 0 THEN CALL TableSHOW(RFTable, RPTFLD$(), RFLDCHN())
401 ' Select Entry
    CALL OptDISPLAY(0)
    IF RFTable.FT = 0 THEN FUN$ = "ADD": GOTO 402
    CALL TableSELECT(RFTable, RPTFLD$(), RFLDCHN())
    SELECT CASE A
       CASE 13: FUN$ = "CHG"
       CASE 59: FUN$ = "CHG"
       CASE 60: FUN$ = "ADD"
       CASE 61: FUN$ = "DEL"
       CASE 62: FUN$ = "MOV"
       CASE 27: FUN$ = "ESC"
       CASE ELSE: CALL ErrBEEP(0): GOTO 401
    END SELECT
    IF FUN$ = "DEL" THEN
       CALL PutMSG("RPTSM004")
       IF A$ <> SNGLKEY$(1) THEN GOTO 401
    END IF
402 ' Process Request
    CALL OptHILITE(FUN$, "", 6)
    SELECT CASE FUN$
       CASE "ADD": GOTO 403  ' ADD New Field
       CASE "ESC": GOTO 409  ' Return
    END SELECT
    Y = RFTable.CT: RDFFld = RFLDS(Y)
    CALL TableUPDT(FUN$, RFTable, RPTFLD$(), RFLDCHN())
    IF (A <> 27) OR (RFTable.CT <> Y) THEN RDFReport.CHG = -1
    IF FUN$ = "DEL" THEN
       RDFHdr.FLDTOTL = RDFHdr.FLDTOTL - 1
       GOTO 401
    END IF
    IF A = 27 THEN GOTO 401
    X$ = RPTFLD$(RFTable.CT)
    RDFFld.RLTH = CHR$(VAL(MID$(X$, 14, 2)))
    W = LEN(RTRIM$(MID$(X$, 16, 4))): RDFFld.ALTH = CHR$(W)
    X = LEN(RTRIM$(MID$(X$, 20, 25))): RDFFld.LLTH = CHR$(X)
    Y = LEN(RTRIM$(MID$(X$, 45, 25))): RDFFld.HLTH = CHR$(Y)
    RFLDS(RFTable.CT) = RDFFld
    Y$ = SPACE$(W + X + Y)
    Z = 1
    IF W > 0 THEN MID$(Y$, Z, W) = MID$(X$, 16, 4): Z = Z + W
    IF X > 0 THEN MID$(Y$, Z, X) = MID$(X$, 20, 25): Z = Z + X
    IF Y > 0 THEN MID$(Y$, Z, Y) = MID$(X$, 45, 25)
    RFLD$(RFTable.CT) = Y$
    RDFHdr.VER = CHR$(255)
    GOTO 401
403 ' ADD New Field Entry
    CALL WinSWITCH(4)
404 ' Select Field Group
    IF WIN.WFMT <> FMT016 THEN
       CALL WinCLR
       CALL WinFORMAT(FMT016)
       CALL OptDISPLAY(FMT091)
       CALL TableOPEN(RFGTable)
    END IF
    CALL TableSHOW(RFGTable, RFLDGRP$(), RFGCHN())
    CALL TableSELECT(RFGTable, RFLDGRP$(), RFGCHN())
    IF A = 27 THEN GOTO 407
405 ' Select Field from Group
    CALL WinCLR
    CALL WinFORMAT(FMT018)
    CALL OptDISPLAY(FMT091)
    IF RFFTable.TABNAME <> MID$(RFLDGRP$(RFGTable.CT), 1, 8) THEN
       RFFTable.TABNAME = RFLDGRP$(RFGTable.CT)
       RFFTable.CL = 0
       RFFTable.FT = RFFBGN(RFGTable.CT)
       RFFTable.SIZE = RFFBGN(RFGTable.CT + 1) - RFFTable.FT
    END IF
    CALL TableOPEN(RFFTable)
    CALL TableSHOW(RFFTable, RFLDTAB$(), RFFCHN())
406 CALL TableSELECT(RFFTable, RFLDTAB$(), RFFCHN())
    IF A = 27 THEN GOTO 404
    LSET RDFFld.NAME = RFLDTAB$(RFFTable.CT)
    X = RFTable.FT: FOUND = 0
    WHILE (X > 0) AND (NOT FOUND)
       FOUND = (RDFFld.NAME = MID$(RPTFLD$(X), 1, 5))
       IF NOT FOUND THEN X = RFLDCHN(X).FWD
    WEND
    IF FOUND THEN
       CALL PutMSG("RPTSM005")
       GOTO 406
    END IF
    RDFFld.RLTH = CHR$(0)
    RDFFld.ALTH = CHR$(0)
    RDFFld.LLTH = CHR$(0)
    RDFFld.HLTH = CHR$(0)
    IF RFTable.FF = 0 THEN
       RFMAX = RFMAX + 1
       REDIM PRESERVE RPTFLD$(RFMAX), RFLDCHN(RFMAX) AS ChainPTRS
       REDIM PRESERVE RFLDS(RFMAX) AS ReportFieldREC, RFLD$(RFMAX)
       RFTable.FF = RFMAX
    END IF
    CALL TableUPDT("NEW", RFTable, RPTFLD$(), RFLDCHN())
    RFFree = RFTable.FF
    RFLDS(RFTable.CT) = RDFFld
    RFLD$(RFTable.CT) = ""
    RPTFLD$(RFTable.CT) = SPACE$(RFTable.XL)
    LSET RPTFLD$(RFTable.CT) = RDFFld.NAME
    MID$(RPTFLD$(RFTable.CT), 14, 2) = MID$(RFLDTAB$(RFFTable.CT), 6, 2)
    MID$(RPTFLD$(RFTable.CT), 6, 8) = MID$(RFLDTAB$(RFFTable.CT), 8, 8)
    MID$(RPTFLD$(RFTable.CT), 70, 25) = MID$(RFLDTAB$(RFFTable.CT), 16, 25)
    RDFHdr.FLDTOTL = RDFHdr.FLDTOTL + 1
    FUN$ = "CHG"
407 ' Return from ADD Field
    CALL WinCLR
    CALL WinERASE(0, 0)
    CALL WinSWITCH(5)
    CALL WinFORMAT(FMT014)
    CALL OptDISPLAY(FMTX93)
    CALL TableSHOW(RFTable, RPTFLD$(), RFLDCHN())
    RDFHdr.VER = CHR$(255)
    IF FUN$ = "CHG" THEN GOTO 402
    IF RFTable.FT > 0 THEN GOTO 401
409 ' Return
    RFFree = RFTable.FF
    CALL WinCLR

END SUB

SUB ProcessHTFS
SHARED FMT021, FMTX93
SHARED HTFTable AS TablePARMS
340 ' Process HTFLine Definitions
    CALL WinSWITCH(5): CALL WinCLR
    CALL WinFORMAT(FMT021)
    CALL OptDISPLAY(FMTX93)
    Y = HTFTable.FT: HTFTable.SIZE = 0
    WHILE Y > 0
       HTFTable.SIZE = HTFTable.SIZE + 1
       Y = HTFCHN(Y).FWD
    WEND
    HTFTable.LT = Y
    HTFTable.CL = 0
    CALL TableOPEN(HTFTable)
    IF HTFTable.SIZE = 0 THEN FUN$ = "ADD": GOTO 342
    CALL TableSHOW(HTFTable, HTFTXT$(), HTFCHN())
341 ' Select Entry
    CALL OptDISPLAY(0)
    CALL TableSELECT(HTFTable, HTFTXT$(), HTFCHN())
    SELECT CASE A
       CASE 13: FUN$ = "CHG"
       CASE 59: FUN$ = "CHG"
       CASE 60: FUN$ = "ADD"
       CASE 61: FUN$ = "DEL"
       CASE 62: FUN$ = "MOV"
       CASE 27: FUN$ = "ESC"
       CASE ELSE: CALL ErrBEEP(0): GOTO 341
    END SELECT
    IF FUN$ = "DEL" THEN
       CALL PutMSG("RPTSM003")
       IF A$ <> SNGLKEY$(1) THEN GOTO 341
    END IF
342 ' Process Request
    CALL OptHILITE(FUN$, "", 6)
    SELECT CASE FUN$
       CASE "ADD": GOTO 343  ' ADD New Variable
       CASE "ESC": GOTO 349  ' Return
    END SELECT
    CALL TableUPDT(FUN$, HTFTable, HTFTXT$(), HTFCHN())
    IF A <> 27 THEN RDFReport.CHG = -1
    IF FUN$ = "DEL" THEN
       RDFHdr.HTFTOTL = RDFHdr.HTFTOTL - 1
       IF HTFTable.SIZE = 0 THEN GOTO 349
    END IF
    RDFHdr.VER = CHR$(255)
    GOTO 341
343 ' ADD New HTFLine
    IF HTFTable.FF = 0 THEN
       HTFMAX = HTFMAX + 1
       REDIM PRESERVE HTFTXT$(HTFMAX), HTFCHN(HTFMAX) AS ChainPTRS
       HTFTable.FF = HTFMAX
    END IF
    CALL TableUPDT("ADD", HTFTable, HTFTXT$(), HTFCHN())
    IF A <> 27 THEN
       RDFReport.CHG = -1
       RDFHdr.HTFTOTL = RDFHdr.HTFTOTL + 1
    END IF
    RDFHdr.VER = CHR$(255)
    IF HTFTable.FT > 0 THEN GOTO 341
349 ' Return to Report Display
    HTFFree = HTFTable.FF
    CALL WinCLR

END SUB

SUB ProcessOPTS (RNAME$)
SHARED XOPTTable AS TablePARMS, FMT091, FMT092, FMT095
330 ' Process Report OPTIONS
    CALL WinSWITCH(3)
    FMTNAME$ = "XXXXO001"
    CALL FmtFIND(FMTNAME$): OLFMT = CURFMT
    CALL FmtFindFLD("RPTN"): UTXT$(FFLD.UTXT) = RTRIM$(RNAME$)
    CALL WinFORMAT(OLFMT)
    CALL FmtFIND("XXXXO004"): FMTO04 = CURFMT
    CALL FmtFIND("XXXXO005"): FMTO05 = CURFMT
    CALL FmtFIND("XXXXO006"): FMTO06 = CURFMT
    CALL FmtFIND("XXXXO007"): FMTO07 = CURFMT
    CALL FmtFIND("XXXXO008"): FMTO08 = CURFMT
    CALL FmtFIND("XXXXO009"): FMTO09 = CURFMT
    OPTTable.FMTNAME = "XXXXO004": OPTTable.FMTNUM = 0
    OPTTable.XL = 0: OPTTable.SIZE = 0
    CALL TableOPEN(OPTTable)
    X = OPTTable.FT: X$ = SPACE$(16): Y$ = SPACE$(4): ACHAR$ = "   "
    WHILE X > 0
       OPTTable.SIZE = OPTTable.SIZE + 1
       OPTTable.LT = X
       LSET X$ = RPTOPT$(X)
       CALL RptOptFIND(X$)
       CALL RptOptEXPAND(X$)
       OPTIX = CVI(MID$(X$, 7, 2))
       IF OPTIX > 0 THEN
          RPTOPT$(X) = X$ + MID$(TABDATA$(OPTIX), 10, 30)
          ELSE
          LSET RPTOPT$(X) = X$
       END IF
       X = OPTCHN(X).FWD
    WEND
    ENV.EDMAST = (OPTTable.SIZE = 0)
    IF OPTTable.SIZE > 0 THEN
       CALL TableSHOW(OPTTable, RPTOPT$(), OPTCHN())
    END IF
331 ' Select Entry
    IF ENV.EDMAST THEN OFMT = FMT092 ELSE OFMT = FMT095
    CALL OptDISPLAY(OFMT): FUN$ = ""
    IF OPTTable.SIZE = 0 THEN FUN$ = "ADD": GOTO 332
    OPTTable.FMTNUM = FMTO04
    CALL TableSELECT(OPTTable, RPTOPT$(), OPTCHN())
    SELECT CASE A
       CASE 13: FUN$ = "CHG"
       CASE 59: FUN$ = "CHG"
       CASE 60: IF ENV.EDMAST THEN FUN$ = "ADD"
       CASE 61: IF ENV.EDMAST THEN FUN$ = "DEL"
       CASE 62: FUN$ = "MOV"
       CASE 94: ENV.EDMAST = NOT ENV.EDMAST: GOTO 331
       CASE 27: FUN$ = "ESC"
    END SELECT
    IF FUN$ = "" THEN CALL ErrBEEP(0): GOTO 331
    IF FUN$ = "DEL" THEN
       CALL PutMSG("RPTSM001")
       IF A$ <> SNGLKEY$(1) THEN GOTO 331
    END IF
332 ' Process Request
    CALL OptHILITE(FUN$, "", 6)
    SELECT CASE FUN$
       CASE "ADD": GOTO 333  ' ADD New Option
       CASE "ESC": GOTO 339  ' Return
    END SELECT
    IF FUN$ = "CHG" THEN
       EDTIX = ASC(MID$(RPTOPT$(OPTTable.CT), 9, 1))
       IF EDTIX = 255 THEN CALL ErrBEEP(0): GOTO 331
       SELECT CASE EDTIX
          CASE 22: OPTTable.FMTNUM = FMTO06
          CASE 23: OPTTable.FMTNUM = FMTO07
          CASE 24: OPTTable.FMTNUM = FMTO08
          CASE 25: OPTTable.FMTNUM = FMTO09
          CASE ELSE: OPTTable.FMTNUM = FMTO05
       END SELECT
    END IF
    Y = OPTTable.CT
    CALL TableUPDT(FUN$, OPTTable, RPTOPT$(), OPTCHN())
    IF (A <> 27) OR (OPTTable.CT <> Y) THEN
       RDFReport.CHG = -1: RDFHdr.VER = CHR$(255)
    END IF
    IF A = 27 THEN GOTO 331
    IF FUN$ = "DEL" THEN
       RDFHdr.OPTTOTL = RDFHdr.OPTTOTL - 1: RDFHdr.VER = CHR$(255)
       GOTO 331
    END IF
    IF FUN$ = "MOV" THEN GOTO 331
    CALL RptOptEDIT(RPTOPT$(OPTTable.CT))
    IF NOT OKAY THEN GOTO 332 ELSE RDFHdr.VER = CHR$(255)
    GOTO 331
333 ' ADD New Option Entry
    CALL WinSWITCH(4)
    CALL WinFORMAT(OLFMT)
    CALL OptDISPLAY(FMT091)
    CALL TableOPEN(XOPTTable)
334 ' List Option Selections
    CALL TableSHOW(XOPTTable, TABDATA$(), TABDATACHN())
335 ' Select Option from List
    CALL TableSELECT(XOPTTable, TABDATA$(), TABDATACHN())
    IF A = 27 THEN GOTO 337
    LSET Y$ = TABDATA$(XOPTTable.CT)
    X = OPTTable.FT: FOUND = 0
    WHILE (X > 0) AND (NOT FOUND)
       FOUND = (Y$ = MID$(RPTOPT$(X), 1, 4))
       IF NOT FOUND THEN X = OPTCHN(X).FWD
    WEND
    IF FOUND THEN
       CALL PutMSG("RPTSM002")
       GOTO 335
    END IF
    LSET X$ = LEFT$(TABDATA$(XOPTTable.CT), 4) + STRING$(7, 0)
    CALL RptOptFIND(X$)
    MID$(X$, 5, 2) = MID$(X$, 10, 2)
    CALL RptOptEXPAND(X$)
    IF OPTTable.FF = 0 THEN
       OPTMAX = OPTMAX + 1
       REDIM PRESERVE RPTOPT$(OPTMAX), OPTCHN(OPTMAX) AS ChainPTRS
       OPTTable.FF = OPTMAX
    END IF
    CALL TableUPDT("NEW", OPTTable, RPTOPT$(), OPTCHN())
    OPTFree = OPTTable.FF
    RPTOPT$(OPTTable.CT) = X$ + MID$(TABDATA$(XOPTTable.CT), 10, 30)
    RDFHdr.OPTTOTL = RDFHdr.OPTTOTL + 1: RDFHdr.VER = CHR$(255)
    'FUN$ = "CHG"
337 ' Return from ADD Option
    CALL WinCLR
    CALL WinSWITCH(3)
    CALL TableSHOW(OPTTable, RPTOPT$(), OPTCHN())
    IF FUN$ = "CHG" THEN CALL OptDISPLAY(0): GOTO 332
    IF OPTTable.SIZE > 0 THEN GOTO 331
339 ' Return
    OPTFree = OPTTable.FF
    X = OPTTable.FT
    WHILE X > 0
       RPTOPT$(X) = LEFT$(RPTOPT$(X), 16)
       X = OPTCHN(X).FWD
    WEND
    CALL WinCLR

END SUB

SUB ProcessTABS
SHARED FMT004, FMT006, FMT008, FMT010, FMT024, FMT026, FMT091, FMTX95, TNAME$
SHARED RTABTable AS TablePARMS, RMONTable AS TablePARMS
SHARED RGENTable AS TablePARMS, RRELTable AS TablePARMS, RREFTable AS TablePARMS
SHARED REVTTable AS TablePARMS
SHARED RTABDATA$(), RTABCHN() AS ChainPTRS, RTABNM$()
SHARED RMONTH$(), RMCHN() AS ChainPTRS, RGENDER$(), RGCHN() AS ChainPTRS
SHARED RRELCODE$(), RRCHN() AS ChainPTRS, RREFSYM$(), RREFCHN() AS ChainPTRS
SHARED REVTCODE$(), REVTCHN() AS ChainPTRS
350 ' Process Report TABLEs
    CALL WinSWITCH(5): CALL WinCLR
    CALL WinSWITCH(3)
    IF WIN.WFMT <> FMT004 THEN
       CALL WinCLR
       CALL WinFORMAT(FMT004)
       CALL TableOPEN(RTABTable)
    END IF
    CALL TableSHOW(RTABTable, RTABDATA$(), RTABCHN())
351 ' Select TABLE from Table
    CALL OptDISPLAY(FMT091)
    CALL TableSELECT(RTABTable, RTABDATA$(), RTABCHN())
    SELECT CASE A
       CASE 27: FUN$ = "ESC"
       CASE ELSE: FUN$ = "SEL"
    END SELECT
    CALL OptHILITE(FUN$, "", 6)
    SELECT CASE FUN$
       CASE "ESC": CALL WinCLR: EXIT SUB ' Return
    END SELECT
    LSET TNAME$ = RTABNM$(RTABTable.CT): TFMT = 0
    SELECT CASE TNAME$
       CASE "MONTHS  ": TFMT = FMT006
       CASE "GENDER  ": TFMT = FMT008
       CASE "RELATION": TFMT = FMT010
       CASE "REFERENC": TFMT = FMT024
       CASE "EVENTS  ": TFMT = FMT026
    END SELECT
    IF TFMT = 0 THEN GOTO 351
    CALL WinSWITCH(4)
    CALL WinFORMAT(TFMT)
    CALL OptDISPLAY(FMTX95)
    SELECT CASE TNAME$
       CASE "MONTHS  ": GOTO 355
       CASE "GENDER  ": GOTO 360
       CASE "RELATION": GOTO 365
       CASE "REFERENC": GOTO 370
       CASE "EVENTS  ": GOTO 380
       CASE ELSE: GOTO 351
    END SELECT

353 ' EXIT From Table Processing
    CALL WinCLR
    CALL WinSWITCH(5): CALL WinCLR
    EXIT SUB

354 ' Return From Table Processing
    CALL WinCLR
    CALL WinSWITCH(3)
    GOTO 351

355 ' Process MONTHS() Table
    CALL TableOPEN(RMONTable)
    CALL TableSHOW(RMONTable, RMONTH$(), RMCHN())
356 ' Select Entry
    CALL TableSELECT(RMONTable, RMONTH$(), RMCHN())
    IF (A = 27) OR (A = 59) THEN
       FOR X = 1 TO RMONTable.SIZE: LSET RMONTB$(X) = RMONTH$(X): NEXT X
       GOTO 354
    END IF
    CALL TableUPDT("CHG", RMONTable, RMONTH$(), RMCHN())
    IF A <> 27 THEN RDFHdr.VER = CHR$(255)
    GOTO 356

360 ' Process GENDER() Table
    CALL TableOPEN(RGENTable)
    CALL TableSHOW(RGENTable, RGENDER$(), RGCHN())
361 ' Select Entry
    CALL TableSELECT(RGENTable, RGENDER$(), RGCHN())
    IF (A = 27) OR (A = 59) THEN
       FOR X = 1 TO RGENTable.SIZE: RGENDR$(X) = RGENDER$(X): NEXT X
       GOTO 354
    END IF
    CALL TableUPDT("CHG", RGENTable, RGENDER$(), RGCHN())
    IF A <> 27 THEN RDFHdr.VER = CHR$(255)
    GOTO 361

365 ' Process RELATION() Table
    CALL TableOPEN(RRELTable)
    CALL TableSHOW(RRELTable, RRELCODE$(), RRCHN())
366 ' Select Entry
    CALL TableSELECT(RRELTable, RRELCODE$(), RRCHN())
    IF (A = 27) OR (A = 59) THEN
       FOR X = 1 TO RRELTable.SIZE: LSET RRELCD$(X) = RRELCODE$(X): NEXT X
       GOTO 354
    END IF
    CALL TableUPDT("CHG", RRELTable, RRELCODE$(), RRCHN())
    IF A <> 27 THEN RDFHdr.VER = CHR$(255)
    GOTO 366

370 ' Process REFERENC Table
    CALL TableOPEN(RREFTable)
    CALL TableSHOW(RREFTable, RREFSYM$(), RREFCHN())
371 ' Select Entry
    CALL TableSELECT(RREFTable, RREFSYM$(), RREFCHN())
    IF (A = 27) OR (A = 59) THEN
       FOR X = 1 TO RREFTable.SIZE
           Z = VAL(MID$(RREFSYM$(X), 1, 2))
           IF (Z > 0) AND (Z <= RREFTable.SIZE) THEN
              RXSYM$(Z) = RTRIM$(MID$(RREFSYM$(X), 3, 3))
           END IF
       NEXT X
       GOTO 354
    END IF
    CALL TableUPDT("CHG", RREFTable, RREFSYM$(), RREFCHN())
    IF A <> 27 THEN RDFHdr.VER = CHR$(255)
    GOTO 371

380 ' Process EVENTS() Table
    CALL TableOPEN(REVTTable)
    CALL TableSHOW(REVTTable, REVTCODE$(), REVTCHN())
385 ' Select Entry
    CALL TableSELECT(REVTTable, REVTCODE$(), REVTCHN())
    IF (A = 27) OR (A = 59) THEN
       REDIM REVENT$(REVTTable.SIZE)
       FOR X = 1 TO REVTTable.SIZE
          REVENT$(X) = RTRIM$(LEFT$(REVTCODE$(X), 23))
       NEXT X
       REVENT$(0) = STR$(REVTTable.SIZE)
       GOTO 354
    END IF
    CALL TableUPDT("CHG", REVTTable, REVTCODE$(), REVTCHN())
    IF A <> 27 THEN RDFHdr.VER = CHR$(255)
    GOTO 385

END SUB

SUB ProcessVARS
SHARED RVARTable AS TablePARMS, RVARTAB$(), RVARCHN() AS ChainPTRS
SHARED FMT012, FMT092
450 ' Process VARIABLE Definitions
    CALL WinSWITCH(3)
    CALL WinFORMAT(FMT012)
    CALL OptDISPLAY(FMT092)
    RVARTable.CL = 0
    RVARTable.LT = RVARTable.FT
    RVARTable.SIZE = 0
    WHILE RVARTable.LT > 0
        RVARTable.LT = RVARCHN(RVARTable.LT).FWD
        RVARTable.SIZE = RVARTable.SIZE + 1
    WEND
    CALL TableOPEN(RVARTable)
    IF RVARTable.FT = 0 THEN FUN$ = "ADD": GOTO 452
    CALL TableSHOW(RVARTable, RVARTAB$(), RVARCHN())
451 ' Select Entry
    CALL PutMSG("")
    CALL OptDISPLAY(0)
    CALL TableSELECT(RVARTable, RVARTAB$(), RVARCHN())
    SELECT CASE A
       CASE 13: FUN$ = "CHG"
       CASE 59: FUN$ = "CHG"
       CASE 60: FUN$ = "ADD"
       CASE 61: FUN$ = "DEL"
       CASE 62: FUN$ = "MOV"
       CASE 27: FUN$ = "ESC"
       CASE ELSE: CALL ErrBEEP(0): GOTO 451
    END SELECT
    IF FUN$ = "DEL" THEN
       CALL PutMSG("RPTSM006")
       IF A$ <> SNGLKEY$(1) THEN GOTO 451
    END IF
452 ' Process Request
    CALL OptHILITE(FUN$, "", 6)
    SELECT CASE FUN$
       CASE "ADD": GOTO 453  ' ADD New Variable
       CASE "ESC": GOTO 459  ' Return
    END SELECT
    CALL TableUPDT(FUN$, RVARTable, RVARTAB$(), RVARCHN())
    IF A <> 27 THEN
       RDFReport.CHG = -1
       IF FUN$ = "DEL" THEN
          RDFHdr.VARTOTL = RDFHdr.VARTOTL - 1
          ELSE
          RDFVar.NAME = RVARTAB$(RVARTable.CT)
          L = LEN(RTRIM$(MID$(RVARTAB$(RVARTable.CT), 9, 255)))
          RDFVar.LTH = CHR$(L)
          RVARS(RVARTable.CT) = RDFVar
          RVAR$(RVARTable.CT) = MID$(RVARTAB$(RVARTable.CT), 9, L)
       END IF
    END IF
    RDFHdr.VER = CHR$(255)
    IF RVARTable.FT > 0 THEN GOTO 451 ELSE GOTO 459
453 ' ADD New Variable Entry
    IF RVARTable.FF = 0 THEN
       RVMAX = RVMAX + 1
       REDIM PRESERVE RVARTAB$(RVMAX), RVARCHN(RVMAX) AS ChainPTRS
       REDIM PRESERVE RVARS(RVMAX) AS ReportVariableREC, RVAR$(RVMAX)
       RVARTable.FF = RVMAX
    END IF
    CALL TableUPDT("ADD", RVARTable, RVARTAB$(), RVARCHN())
    IF A <> 27 THEN
       RDFReport.CHG = -1
       RDFHdr.VARTOTL = RDFHdr.VARTOTL + 1
       RDFVar.NAME = RVARTAB$(RVARTable.CT)
       L = LEN(RTRIM$(MID$(RVARTAB$(RVARTable.CT), 9, 255)))
       RDFVar.LTH = CHR$(L)
       RVARS(RVARTable.CT) = RDFVar
       RVAR$(RVARTable.CT) = MID$(RVARTAB$(RVARTable.CT), 9, L)
    END IF
    RDFHdr.VER = CHR$(255)
    IF RVARTable.FT > 0 THEN GOTO 451
459 ' Return
    RVFree = RVARTable.FF
    CALL WinCLR
END SUB

REM $DYNAMIC
SUB RPTSFile
   SHARED RDFINDX AS ReportINDEX
   SHARED RTABTable AS TablePARMS, RTABDATA$(), RTABNM$(), RTABCHN() AS ChainPTRS
   SHARED RFGTable AS TablePARMS, RFLDGRP$(), RFGCHN() AS ChainPTRS
   SHARED RFFTable AS TablePARMS, RFLDTAB$(), RFFCHN() AS ChainPTRS, RFFBGN()
   SHARED RMONTable AS TablePARMS, RMCHN() AS ChainPTRS, RMONTH$()
   SHARED RGENTable AS TablePARMS, RGCHN() AS ChainPTRS, RGENDER$()
   SHARED RRELTable AS TablePARMS, RRCHN() AS ChainPTRS, RRELCODE$()
   SHARED RREFTable AS TablePARMS, RREFCHN() AS ChainPTRS, RREFSYM$()
   SHARED REVTTable AS TablePARMS, REVTCHN() AS ChainPTRS, REVTCODE$()
   SHARED RTable AS TablePARMS, REPORT$(), RPTCHN() AS ChainPTRS
   SHARED RFTable AS TablePARMS, RPTFLD$()
   SHARED RVARTable AS TablePARMS, RVARTAB$()
   SHARED HTFTable AS TablePARMS, HTFCHN() AS ChainPTRS
   SHARED RPTDef AS ReportDefREC, RO$, RBGN, REND
   SHARED XOPTTable AS TablePARMS
   SHARED NEWTABLE AS TablePARMS
   SHARED NewReport AS ReportDefREC
500 ' Process FILE Requests
    CALL PutMSG("XXXXM005")
    IF ENV.RDF <> "   " THEN CONFIG.RDF = ENV.RDF
    IF A = 27 THEN OKAY = 0: GOTO 505 ELSE R$ = A$
    IF (R$ < "0") OR (R$ > "3") THEN SOUND BP, DUR: GOTO 500
    GOSUB 510: IF A = 27 THEN GOTO 500
    L$ = " ": ZCHG = (R$ = "0")
    SELECT CASE R$
       CASE "0": GOTO 520
       CASE "1": GOTO 520
       CASE "2": GOTO 530
       CASE "3": GOTO 540
    END SELECT
    GOTO 510

505 ' Early RETURN
    EXIT SUB

510 ' Get RDF File Prefix
    CALL FmtFIND("RPTSS050")
    CALL FmtFindFLD("RDF")
    LSET UTXT$(FFLD.UTXT) = CONFIG.RDF
    CALL MsgFORMAT: CALL WinSWITCH(2)
    IF A = 27 THEN RETURN
    CALL FmtFindFLD("RDF"): CONFIG.RDF = UTXT$(FFLD.UTXT)
    IF CONFIG.RDF <> ENV.RDF THEN CONFIG.VER = -1
    RETURN

520 ' Save Report Descriptions Into RDF File
    CALL RDFOpen("NEW"): IF NOT OKAY THEN GOTO 500
    CALL PutMSG("RPTSM007")
    RDFHdr.FTYPE = "RDF": RDFHdr.VER = CHR$(1): RDFHdr.EOFRBA = 0
    RDFHdr.REPORTS = 0
    RDFHdr.HTFTOTL = 0: RDFHdr.VARTOTL = 0: RDFHdr.FLDTOTL = 0
    RDFHdr.OPTTOTL = 0
    RDFHdr.VARIABLE = 0: RDFHdr.FIELDS = 0: RDFHdr.OPTIONS = 0
    PUT #1, 1, RDFHdr
521 FOR X = 1 TO RPTMAX
        IF RDFIndex(X).DEFIX <> 0 THEN
           RDFHdr.REPORTS = RDFHdr.REPORTS + 1
           PUT #1, , RDFDir
        END IF
    NEXT X
    RDFHdr.TABLES = 5
522 FOR X = 1 TO 5
        SELECT CASE X
           CASE 1: RDFTable.NAME = "MONTHS": RDFTable.ENTCNT = RMONTable.SIZE
           CASE 2: RDFTable.NAME = "GENDER": RDFTable.ENTCNT = RGENTable.SIZE
           CASE 3: RDFTable.NAME = "RELATION": RDFTable.ENTCNT = RRELTable.SIZE
           CASE 4: RDFTable.NAME = "REFERENC": RDFTable.ENTCNT = RREFTable.SIZE
           CASE 5: RDFTable.NAME = "EVENTS": RDFTable.ENTCNT = REVTTable.SIZE
        END SELECT
        RDFTable.VER = 0
        PUT #1, , RDFTable
523     FOR Y = 1 TO RDFTable.ENTCNT
            SELECT CASE X
               CASE 1: X$ = RMONTB$(Y)
               CASE 2: X$ = RGENDR$(Y)
               CASE 3: X$ = RRELCD$(Y)
               CASE 4: X$ = LEFT$(RREFSYM$(Y), 5)
               CASE 5: X$ = REVENT$(Y)
            END SELECT
            L$ = CHR$(LEN(X$))
            PUT #1, , L$
            PUT #1, , X$
        NEXT Y
    NEXT X
    X = RDFIndex(0).VARBGN
524 WHILE X > 0
       PUT #1, , RVARS(X)
       PUT #1, , RVAR$(X)
       RDFHdr.VARIABLE = RDFHdr.VARIABLE + 1
       X = RVARCHN(X).FWD
    WEND
    RDFHdr.VARTOTL = RDFHdr.VARIABLE
    X = RDFIndex(0).FLDBGN
525 WHILE X > 0
       PUT #1, , RFLDS(X)
       PUT #1, , RFLD$(X)
       RDFHdr.FIELDS = RDFHdr.FIELDS + 1
       X = RFLDCHN(X).FWD
    WEND
    RDFHdr.FLDTOTL = RDFHdr.FIELDS
    X = RDFIndex(0).OPTBGN: X$ = SPACE$(6)
526 WHILE X > 0
       LSET X$ = RPTOPT$(X)
       IF LEFT$(X$, 4) = "BGPP" THEN MID$(X$, 5, 2) = MKI$(1)
       PUT #1, , X$
       RDFHdr.OPTIONS = RDFHdr.OPTIONS + 1
       X = OPTCHN(X).FWD
    WEND
    RDFHdr.OPTTOTL = RDFHdr.OPTIONS
527 FOR X = 1 TO RPTMAX
       RDFINDX = RDFIndex(X)
       IF RDFINDX.DEFIX > 0 THEN
          R = RDFINDX.DEFIX
          RDFINDX.DEFRBA = SEEK(1)
          RDFReport = RDFReports(R)
          IF ZCHG THEN
             RDFReport.CHG = 0
             ELSE
             IF RDFINDX.CHG < 0 THEN
                RDFReport.CHG = (((ABS(RDFINDX.CHG) AND 255) + 1) AND 255)
             END IF
          END IF
          RDFINDX.CHG = RDFReport.CHG
          PUT #1, , RDFReport
          Y = RDFINDX.OPTBGN: X$ = SPACE$(6)
          WHILE Y > 0
             LSET X$ = RPTOPT$(Y)
             IF LEFT$(X$, 4) = "BGPP" THEN MID$(X$, 5, 2) = MKI$(1)
             PUT #1, , X$
             RDFHdr.OPTTOTL = RDFHdr.OPTTOTL + 1
             Y = OPTCHN(Y).FWD
          WEND
          Y = RDFINDX.HTFBGN
          WHILE Y > 0
             RDFHTFLine.TYPE = MID$(HTFTXT$(Y), 1, 1)
             RDFHTFLine.SECTION = MID$(HTFTXT$(Y), 2, 1)
             RDFHTFLine.USE = MID$(HTFTXT$(Y), 3, 1)
             L = LEN(RTRIM$(MID$(HTFTXT$(Y), 4, 255)))
             RDFHTFLine.LTH = CHR$(L)
             X$ = MID$(HTFTXT$(Y), 4, L)
             PUT #1, , RDFHTFLine
             PUT #1, , X$
             RDFHdr.HTFTOTL = RDFHdr.HTFTOTL + 1
             Y = HTFCHN(Y).FWD
          WEND
          Y = RDFINDX.VARBGN
          WHILE Y > 0
             PUT #1, , RVARS(Y)
             PUT #1, , RVAR$(Y)
             RDFHdr.VARTOTL = RDFHdr.VARTOTL + 1
             Y = RVARCHN(Y).FWD
          WEND
          Y = RDFINDX.FLDBGN
          WHILE Y > 0
             PUT #1, , RFLDS(Y)
             PUT #1, , RFLD$(Y)
             RDFHdr.FLDTOTL = RDFHdr.FLDTOTL + 1
             Y = RFLDCHN(Y).FWD
          WEND
          RDFReports(R) = RDFReport
       END IF
       RDFIndex(X) = RDFINDX
    NEXT X
    RDFHdr.EOFRBA = SEEK(1) - 1
    PUT #1, 1, RDFHdr
528 FOR X = 1 TO RPTMAX
        RDFINDX = RDFIndex(X)
        IF RDFINDX.DEFIX <> 0 THEN
           RDFReport = RDFReports(RDFINDX.DEFIX)
           RDFDir.CODE = RDFINDX.CODE
           RDFDir.CHG = RDFINDX.CHG
           RDFDir.DEFRBA = RDFINDX.DEFRBA
           PUT #1, , RDFDir
        END IF
    NEXT X
    CLOSE
    CALL RptLOAD("ALL"): IF NOT OKAY THEN GOTO 500
    CALL BuildTABLES: GOTO 505

530 ' Load Report Definitions from RDF File
    IF RDFHdr.VER = CHR$(255) THEN
       CALL PutMSG("XXXXM020")
       IF A$ <> SNGLKEY$(1) THEN GOTO 500
    END IF
    RDFHdr.FTYPE = "RDF"
    CALL PutMSG("RPTSM008")
    CALL RptLOAD("ALL"): IF NOT OKAY THEN GOTO 500
    CALL BuildTABLES
    GOTO 505

540 ' Merge Report Definitions from RDF File
    IF RDFHdr.VER = CHR$(255) THEN
       CALL PutMSG("XXXXM020")
       IF A$ <> SNGLKEY$(1) THEN GOTO 500
    END IF
    DIM MergRDFHdr AS ReportFileHDR
    MergRDFHdr = RDFHdr
    CALL RDFOpen("READ"): IF NOT OKAY THEN RDFHdr = MergRDFHdr: GOTO 500
    CALL PutMSG("RPTSM009")
    X = RDFHdr.REPORTS
    REDIM MRDEF(RPTMAX)
    WHILE X > 0
       GET #1, , RDFDir
       Y = RPTMAX: FOUND = 0
       WHILE (Y > 0) AND (NOT FOUND)
          FOUND = (RDFIndex(Y).CODE = RDFDir.CODE)
          IF NOT FOUND THEN Y = Y - 1
       WEND
       IF FOUND THEN MRDEF(Y) = RDFDir.DEFRBA
       X = X - 1
    WEND

    'NOTE: Table in Merging File must be a Table in MASTER File to be processed
    '      HOWEVER
    '      ALL TABLE Entries in MERGING File Replace ALL Table Entries in MASTER File
    X = RDFHdr.TABLES
    WHILE X > 0
       GET #1, , RDFTable
       IF RDFTable.ENTCNT > 0 THEN
          SELECT CASE RDFTable.NAME
             CASE "MONTHS  ": REDIM RMONTB$(RDFTable.ENTCNT)
                              RMONTB$(0) = STR$(RDFTable.ENTCNT)
             CASE "GENDER  ": REDIM RGENDR$(RDFTable.ENTCNT)
                              RGENDR$(0) = STR$(RDFTable.ENTCNT)
             CASE "RELATION": REDIM RRELCD$(RDFTable.ENTCNT)
                              RRELCD$(0) = STR$(RDFTable.ENTCNT)
             CASE "REFERENC": REDIM RXSYM$(RDFTable.ENTCNT)
                              RXSYM$(0) = STR$(RDFTable.ENTCNT)
          END SELECT
          FOR Y = 1 TO RDFTable.ENTCNT
             GET #1, , L$: L = ASC(L$)
             X$ = SPACE$(L)
             GET #1, , X$
             SELECT CASE RDFTable.NAME
                CASE "MONTHS  ": RMONTB$(Y) = X$
                CASE "GENDER  ": RGENDR$(Y) = X$
                CASE "RELATION": RRELCD$(Y) = X$
                CASE "REFERENC": RXSYM$(Y) = MID$(X$, 3, 3)
             END SELECT
          NEXT Y
       END IF
       X = X - 1
    WEND

    ' Process GLOBAL Variables & Field Definitions
    X = RDFHdr.VARIABLE: Y0 = RDFIndex(0).VARBGN: GOSUB 550
    X = RDFHdr.FIELDS: Y0 = RDFIndex(0).FLDBGN: GOSUB 555
    X = RDFHdr.OPTIONS: Y0 = RDFIndex(0).OPTBGN: GOSUB 560

    ' Process Individual Merging Reports
    FOR W = 1 TO RPTMAX
       RDFINDX = RDFIndex(W)
       RDFINDX.DEFRBA = MRDEF(W)
       IF RDFINDX.DEFRBA > 0 THEN
          GET #1, MRDEF(W), RDFReport
          IF RDFINDX.DEFIX = 0 THEN
             IF RDFFree = 0 THEN
                RDFMAX = RDFMAX + 1
                REDIM PRESERVE RDFReports(RDFMAX) AS ReportDefREC
                REDIM PRESERVE RDFCHN(RDFMAX) AS ChainPTRS
                RDFFree = RDFMAX
             END IF
             RDFINDX.DEFIX = RDFFree
             RDFFree = RDFCHN(RDFFree).FWD
          END IF
          RDFReports(RDFINDX.DEFIX) = RDFReport
          IF RDFReport.OPTCNT > 0 THEN
             X = RDFReport.OPTCNT: Y0 = RDFINDX.OPTBGN
             GOSUB 560
          END IF
          IF RDFReport.HTFCNT > 0 THEN GOSUB 565
          X = RDFReport.VARDEFS: Y0 = RDFINDX.VARBGN: GOSUB 550
          X = RDFReport.FLDDEFS: Y0 = RDFINDX.FLDBGN: GOSUB 555
       END IF
       RDFIndex(W) = RDFINDX
    NEXT W
    CLOSE
    RDFHdr = MergRDFHdr
    CALL BuildTABLES
    GOTO 505

550 ' Merge Report Variable Records
    WHILE X > 0
       GET #1, , RDFVar
       X$ = SPACE$(ASC(RDFVar.LTH))
       GET #1, , X$
       Y = Y0: Z = Y: FOUND = 0
       WHILE (Y > 0) AND (NOT FOUND)
          FOUND = (RVARS(Y).NAME = RDFVar.NAME)
          IF NOT FOUND THEN Z = Y: Y = RVARCHN(Y).FWD
       WEND
       IF FOUND THEN
          RVARS(Y) = RDFVar
          RVAR$(Y) = X$
       END IF
       X = X - 1
    WEND
    RETURN

555 ' Merge Report Field Records
    WHILE X > 0
       GET #1, , RDFFld
       X$ = SPACE$(ASC(RDFFld.ALTH) + ASC(RDFFld.LLTH) + ASC(RDFFld.HLTH))
       GET #1, , X$
       Y = Y0: Z = Y: FOUND = 0
       WHILE (Y > 0) AND (NOT FOUND)
          FOUND = (RFLDS(Y).NAME = RDFFld.NAME)
          IF NOT FOUND THEN Z = Y: Y = RFLDCHN(Y).FWD
       WEND
       IF FOUND THEN
          RFLDS(Y) = RDFFld
          RFLD$(Y) = X$
       END IF
       X = X - 1
    WEND
    RETURN

560 ' Merge Options
    X$ = SPACE$(6)
    WHILE X > 0
       GET #1, , X$
       Y = Y0: FOUND = 0
       WHILE (Y > 0) AND (NOT FOUND)
          FOUND = (O$ = LEFT$(RPTOPT$(Y), 4))
          IF NOT FOUND THEN Y = OPTCHN(Y).FWD
       WEND
       IF FOUND THEN MID$(RPTOPT$(Y), 5, 2) = MID$(X$, 5, 2)
       X = X - 1
    WEND
    RETURN

565 ' Merge HTFLines
    X0 = RDFINDX.HTFBGN
    IF X0 > 0 THEN
       Y = X0
       WHILE Y > 0
          HTFTXT$(Y) = "": Z = Y: Y = HTFCHN(Y).FWD
          MergRDFHdr.HTFTOTL = MergRDFHdr.HTFTOTL - 1
       WEND
       IF HTFFree > 0 THEN HTFCHN(HTFFree).BWD = Z
       HTFCHN(Z).FWD = HTFFree
       HTFFree = X0
    END IF
    X0 = HTFFree
    X = RDFReport.HTFCNT: Z = 0
    WHILE X > 0
       GET #1, , RDFHTFLine
       X$ = SPACE$(ASC(RDFHTFLine.LTH))
       GET #1, , X$
       IF HTFFree = 0 THEN
          HTFFree = HTFMAX + 1
          HTFMAX = HTFMAX + X
          REDIM PRESERVE HTFTXT$(HTFMAX), HTFCHN(HTFMAX) AS ChainPTRS
          CALL ChnINIT(HTFCHN(), HTFFree, HTFMAX)
          IF X0 = 0 THEN X0 = HTFFree
       END IF
       Y = HTFFree
       IF Z > 0 THEN HTFCHN(Z).FWD = Y
       HTFCHN(Y).BWD = Z
       HTFFree = HTFCHN(Y).FWD: IF HTFFree > 0 THEN HTFCHN(HTFFree).BWD = 0
       HTFTXT$(Y) = RDFHTFLine.TYPE + RDFHTFLine.SECTION + RDFHTFLine.USE + X$
       MergRDFHdr.HTFTOTL = MergRDFHdr.HTFTOTL + 1
       Z = Y
       X = X - 1
    WEND
    RDFINDX.HTFBGN = X0
    RETURN

END SUB

SUB RPTSReport
   SHARED RDFINDX AS ReportINDEX
   SHARED RTABTable AS TablePARMS, RTABDATA$(), RTABNM$(), RTABCHN() AS ChainPTRS
   SHARED RFGTable AS TablePARMS, RFLDGRP$(), RFGCHN() AS ChainPTRS
   SHARED RFFTable AS TablePARMS, RFLDTAB$(), RFFCHN() AS ChainPTRS, RFFBGN()
   SHARED RMONTable AS TablePARMS, RMCHN() AS ChainPTRS, RMONTH$()
   SHARED RGENTable AS TablePARMS, RGCHN() AS ChainPTRS, RGENDER$()
   SHARED RRELTable AS TablePARMS, RRCHN() AS ChainPTRS, RRELCODE$()
   SHARED RREFTable AS TablePARMS, RREFCHN() AS ChainPTRS, RREFSYM$()
   SHARED REVTTable AS TablePARMS, REVTCHN() AS ChainPTRS, REVTCODE$()
   SHARED RTable AS TablePARMS, REPORT$(), RPTCHN() AS ChainPTRS
   SHARED RFTable AS TablePARMS, RPTFLD$()
   SHARED RVARTable AS TablePARMS, RVARTAB$()
   SHARED HTFTable AS TablePARMS, HTFCHN() AS ChainPTRS
   SHARED RPTDef AS ReportDefREC, RO$, RBGN, REND
   SHARED XOPTTable AS TablePARMS
   SHARED NEWTABLE AS TablePARMS
   SHARED NewReport AS ReportDefREC
600 ' Process PRINT Request
    IF RO$ <> "" THEN GOTO 601
    HSETUP = ENV.PSETUP
    CALL PutMSG("RPTSM010")
    RO$ = A$: IF A = 27 THEN EXIT SUB
    IF (RO$ < "1") OR (RO$ > "4") THEN CALL ErrBEEP(0): GOTO 600
601 ' Print REPORTS
    CALL RptOPEN("RPTS"): IF (NOT OKAY) THEN 605
    CALL PrintOPEN
    IF LEN(PLINE$) < 100 THEN PLINE$ = SPACE$(100)
602 ' Begin Printing
    SYSVAR.FILENAME = CONFIG.RDF + "RPTS.RDF"
    CALL GetRVAR("TOTAL", X): TOTAL$ = RVAR$(X)
    PRT.SECTION = "0"
    CALL PrintHEAD
    IF RO$ = "2" THEN GOTO 603
    IF RO$ = "5" THEN GOTO 604
    RBGN = 1: REND = RPTMAX
    IF RO$ = "3" THEN GOTO 604
    CALL GetRVAR("RPTTOTAL", X): RPTTOTAL$ = RVAR$(X)
    CALL GetRVAR("HTFTOTAL", X): HTFTOTAL$ = RVAR$(X)
    CALL GetRVAR("OPTTOTAL", X): OPTTOTAL$ = RVAR$(X)
    CALL GetRVAR("VARTOTAL", X): VARTOTAL$ = RVAR$(X)
    CALL GetRVAR("FLDTOTAL", X): FLDTOTAL$ = RVAR$(X)
    MID$(PLINE$, 5, 30) = TOTAL$ + "  " + RPTTOTAL$ + STR$(RDFHdr.REPORTS)
    X = INSTR(5, PLINE$, "   ")
    MID$(PLINE$, X + 2, 30) = HTFTOTAL$ + STR$(RDFHdr.HTFTOTL)
    X = INSTR(X + 2, PLINE$, "   ")
    MID$(PLINE$, X + 2, 30) = OPTTOTAL$ + STR$(RDFHdr.OPTTOTL)
    X = INSTR(X + 2, PLINE$, "   ")
    MID$(PLINE$, X + 2, 30) = VARTOTAL$ + STR$(RDFHdr.VARTOTL)
    X = INSTR(X + 2, PLINE$, "   ")
    MID$(PLINE$, X + 2, 30) = FLDTOTAL$ + STR$(RDFHdr.FLDTOTL)
    CALL WriteLINE
    PRT.SECTION = "1"      ' MONTHS Table
    CALL WriteLINE: CALL PrintTITLE
    X = RMONTable.FT
    WHILE X > 0
       MID$(PLINE$, 21, 2) = MID$(RMONTH$(X), 1, 2)
       MID$(PLINE$, 27, 3) = MID$(RMONTH$(X), 3, 3)
       MID$(PLINE$, 34, 30) = MID$(RMONTH$(X), 6, 30)
       CALL WriteLINE
       X = RMCHN(X).FWD
    WEND
    PRT.SECTION = "2"      ' GENDER Table
    CALL WriteLINE: CALL PrintTITLE
    X = RGENTable.FT
    WHILE X > 0
       MID$(PLINE$, 21, 1) = MID$(RGENDER$(X), 1, 1)
       MID$(PLINE$, 28, 1) = MID$(RGENDER$(X), 2, 1)
       MID$(PLINE$, 34, 30) = MID$(RGENDER$(X), 3, 30)
       CALL WriteLINE
       X = RGCHN(X).FWD
    WEND
    PRT.SECTION = "3"      ' RELCODES Table
    CALL WriteLINE: CALL PrintTITLE
    X = RRELTable.FT
    WHILE X > 0
       MID$(PLINE$, 20, 2) = MID$(RRELCODE$(X), 1, 2)
       MID$(PLINE$, 27, 3) = MID$(RRELCODE$(X), 3, 3)
       MID$(PLINE$, 34, 12) = MID$(RRELCODE$(X), 6, 12)
       MID$(PLINE$, 48, 12) = MID$(RRELCODE$(X), 18, 12)
       CALL WriteLINE
       X = RRCHN(X).FWD
    WEND
    PRT.SECTION = "4"      ' REFSYMS Table
    CALL WriteLINE: CALL PrintTITLE
    X = RREFTable.FT
    WHILE X > 0
       MID$(PLINE$, 20, 2) = MID$(RREFSYM$(X), 1, 2)
       MID$(PLINE$, 27, 3) = MID$(RREFSYM$(X), 3, 3)
       MID$(PLINE$, 34, 30) = MID$(RREFSYM$(X), 6, 30)
       CALL WriteLINE
       X = RREFCHN(X).FWD
    WEND
    PRT.SECTION = "5"      ' EVENTS Table
    CALL WriteLINE: CALL PrintTITLE
    X = REVTTable.FT
    WHILE X > 0
       MID$(PLINE$, 20, 8) = MID$(REVTCODE$(X), 1, 8)
       MID$(PLINE$, 30, 15) = MID$(REVTCODE$(X), 9, 15)
       MID$(PLINE$, 48, 25) = MID$(REVTCODE$(X), 24, 25)
       CALL WriteLINE
       X = REVTCHN(X).FWD
    WEND
    Y = RDFIndex(0).VARBGN: GOSUB 610
    Y = RDFIndex(0).FLDBGN: GOSUB 620
    Y = RDFIndex(0).OPTBGN: GOSUB 630
    IF RO$ = "1" THEN GOTO 605
603 ' Report LIST
    PRT.SECTION = "R"
    X$ = SPACE$(RTable.XL)
    CALL WriteLINE: CALL PrintTITLE
    T1 = 0: T2 = 0: T3 = 0: T4 = 0: RNUM5$ = SPACE$(5)
    FOR X = 1 TO RPTMAX
        LSET X$ = REPORT$(X)
        MID$(PLINE$, 5, 4) = MID$(X$, 1, 4)
        MID$(PLINE$, 16, 4) = MID$(X$, 5, 4)
        MID$(PLINE$, 22, 3) = MID$(X$, 9, 3): T1 = T1 + VAL(MID$(X$, 9, 3))
        MID$(PLINE$, 27, 3) = MID$(X$, 12, 3): T2 = T2 + VAL(MID$(X$, 12, 3))
        MID$(PLINE$, 32, 3) = MID$(X$, 15, 3): T3 = T3 + VAL(MID$(X$, 15, 3))
        MID$(PLINE$, 37, 3) = MID$(X$, 18, 3): T4 = T4 + VAL(MID$(X$, 18, 3))
        MID$(PLINE$, 44, 45) = MID$(X$, 21, 45)
        CALL WriteLINE
    NEXT X
    MID$(PLINE$, 21, 4) = "----": MID$(PLINE$, 26, 4) = "----"
    MID$(PLINE$, 31, 4) = "----": MID$(PLINE$, 36, 4) = "----"
    CALL WriteLINE
    MID$(PLINE$, 10, 10) = TOTAL$
    RSET RNUM5$ = STR$(T1): MID$(PLINE$, 20, 5) = RNUM5$
    RSET RNUM5$ = STR$(T2): MID$(PLINE$, 25, 5) = RNUM5$
    RSET RNUM5$ = STR$(T3): MID$(PLINE$, 30, 5) = RNUM5$
    RSET RNUM5$ = STR$(T4): MID$(PLINE$, 35, 5) = RNUM5$
    CALL WriteLINE
    IF RO$ = "2" THEN GOTO 605
604 ' Report Definitions
    CALL GetRVAR("REPORT", X): REPORT$ = RVAR$(X)
    CALL GetRVAR("RPTTYPE", X): RPTTYPE$ = RVAR$(X)
    CALL GetRVAR("PAGEWDTH", X): PAGEWDTH$ = RVAR$(X)
    CALL GetRVAR("BLNKLINE", X): BLNKLINE$ = RVAR$(X)
    CALL GetRVAR("BLHEAD", X): BLHEAD$ = RVAR$(X)
    CALL GetRVAR("BLTITLE", X): BLTITLE$ = RVAR$(X)
    CALL GetRVAR("BLFOOT", X): BLFOOT$ = RVAR$(X)
    CALL GetRVAR("MESSAGES", X): MESSAGE$ = RVAR$(X)
    CALL GetRVAR("PRINTMSG", X): PRINTMSG$ = RVAR$(X)
    CALL GetRVAR("WAITMSG", X): WAITMSG$ = RVAR$(X)
    CALL GetRVAR("PAGEMSG", X): PAGEMSG$ = RVAR$(X)
    CALL GetRVAR("INTMSG", X): INTMSG$ = RVAR$(X)
    CALL GetRVAR("ENDMSG", X): ENDMSG$ = RVAR$(X)
    FOR X = RBGN TO REND
        IF RDFIndex(X).DEFIX > 0 THEN
           PRT.SECTION = "D"
           CALL WriteLINE
           Y = RDFIndex(X).DEFIX: GOSUB 650
           Y = RDFIndex(X).OPTBGN: GOSUB 630
           Y = RDFIndex(X).HTFBGN: GOSUB 640
           Y = RDFIndex(X).VARBGN: GOSUB 610
           Y = RDFIndex(X).FLDBGN: GOSUB 620
           CALL WriteLINE
           IF A = 27 THEN X = REND
        END IF
    NEXT X

605 ' End of Print
    CALL RptCLOSE: IF PRT.AGAIN THEN GOTO 602
    EXIT SUB

610 ' Print Variables
    IF Y = 0 THEN RETURN
    PRT.SECTION = "V"
    CALL WriteLINE: CALL PrintTITLE
    WHILE Y > 0
       MID$(PLINE$, 13, 8) = MID$(RVARTAB$(Y), 1, 8)
       MID$(PLINE$, 25, 255) = MID$(RVARTAB$(Y), 9, 255)
       CALL WriteLINE
       Y = RVARCHN(Y).FWD
    WEND
    RETURN

620 ' Print Fields
    IF Y = 0 THEN RETURN
    PRT.SECTION = "F"
    X$ = SPACE$(95)
    CALL WriteLINE: CALL PrintTITLE
    WHILE Y > 0
       LSET X$ = RPTFLD$(Y)
       MID$(PLINE$, 13, 5) = MID$(X$, 1, 5)
       MID$(PLINE$, 22, 8) = MID$(X$, 6, 8)
       MID$(PLINE$, 34, 2) = MID$(X$, 14, 2)
       MID$(PLINE$, 39, 4) = MID$(X$, 16, 4)
       MID$(PLINE$, 45, 25) = MID$(X$, 20, 25)
       MID$(PLINE$, 72, 25) = MID$(X$, 45, 25)
       MID$(PLINE$, 100, 25) = MID$(X$, 70, 25)
       CALL WriteLINE
       Y = RFLDCHN(Y).FWD
    WEND
    RETURN

630 ' Print OPTIONS
    IF Y = 0 THEN RETURN
    PRT.SECTION = "O"
    CALL WriteLINE: CALL PrintTITLE
    X$ = SPACE$(OPTTable.XL): Y$ = SPACE$(4)
    WHILE Y > 0
       LSET X$ = RPTOPT$(Y): LSET Y$ = X$
       CALL RptOptFIND(X$)
       CALL RptOptEXPAND(X$)
       MID$(PLINE$, 14, 4) = MID$(X$, 1, 4)
       MID$(PLINE$, 20, 5) = MID$(X$, 12, 5)
       OPTIX = CVI(MID$(X$, 7, 2))
       IF OPTIX > 0 THEN MID$(PLINE$, 31, 30) = MID$(TABDATA$(OPTIX), 10, 30)
       CALL WriteLINE
       Y = OPTCHN(Y).FWD
    WEND
    RETURN

640 ' Print HTF Lines
    IF Y = 0 THEN RETURN
    PRT.SECTION = "H"
    CALL WriteLINE: CALL PrintTITLE
    X$ = SPACE$(HTFTable.XL)
    WHILE Y > 0
       LSET X$ = HTFTXT$(Y)
       MID$(PLINE$, 16, 1) = MID$(X$, 1, 1)
       MID$(PLINE$, 23, 1) = MID$(X$, 2, 1)
       MID$(PLINE$, 29, 1) = MID$(X$, 3, 1)
       MID$(PLINE$, 35, 127) = MID$(X$, 4, 127)
       CALL WriteLINE
       Y = HTFCHN(Y).FWD
    WEND
    RETURN

650 ' Basic REPORT Information
    IF A = 27 THEN RETURN
    RPTDef = RDFReports(Y)
    MID$(PLINE$, 1, 10) = REPORT$: Y = INSTR(PLINE$, "   ")
    MID$(PLINE$, Y + 1, 4) = RPTDef.CODE
    MID$(PLINE$, Y + 8, 255) = RPTDef.NAME
    CALL WriteLINE
    MID$(PLINE$, 5, 255) = RPTTYPE$
    Z = INSTR(5, PLINE$, "  "): MID$(PLINE$, Z, 2) = STR$(RPTDef.TYPE)
    Z = LEN(RPTTYPE$): MID$(PLINE$, Z + 10, 255) = PAGEWDTH$ + STR$(RPTDef.PWIDTH)
    CALL WriteLINE
    MID$(PLINE$, 5, 255) = BLNKLINE$
    Z = LEN(BLNKLINE$) + 8
    MID$(PLINE$, Z, 255) = BLHEAD$ + STR$(RPTDef.BLHEAD)
    Z = INSTR(Z, PLINE$, "  ") + 3
    MID$(PLINE$, Z, 255) = BLTITLE$ + STR$(RPTDef.BLTITLE)
    Z = INSTR(Z, PLINE$, "  ") + 3
    MID$(PLINE$, Z, 255) = BLFOOT$ + STR$(RPTDef.BLFOOT)
    CALL WriteLINE
    MID$(PLINE$, 5, 255) = MESSAGE$
    Z = LEN(MESSAGE$) + 8
    MID$(PLINE$, Z, 255) = PRINTMSG$ + " " + RPTDef.PRTMSG
    Z = INSTR(Z, PLINE$, "  ") + 3
    MID$(PLINE$, Z, 255) = WAITMSG$ + " " + RPTDef.WAITMSG
    Z = INSTR(Z, PLINE$, "  ") + 3
    MID$(PLINE$, Z, 255) = PAGEMSG$ + " " + RPTDef.PAGEMSG
    Z = INSTR(Z, PLINE$, "  ") + 3
    MID$(PLINE$, Z, 255) = INTMSG$ + " " + RPTDef.INTMSG
    Z = INSTR(Z, PLINE$, "  ") + 3
    MID$(PLINE$, Z, 255) = ENDMSG$ + " " + RPTDef.ENDMSG
    CALL WriteLINE
    RETURN
END SUB

