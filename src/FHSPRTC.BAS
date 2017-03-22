DECLARE SUB Dec2Code (ICODE$, OCODE$)
DECLARE SUB Hex2Code (ICODE$, OCODE$)
DECLARE SUB Code2Hex (ICODE$, OCODE$)
DECLARE SUB Code2Dec (ICODE$, OCODE$)
1  REM $INCLUDE: 'FHSCOMON.BAS'
   PN$ = " FAMILY HISTORY SYSTEM - FHSPRTC - Printer Table Customization program"
   CY$ = " (C) Copyright 1995 by Phillip E. Brown"
   IF CONFIG.BP = 0 THEN RUN "FHSINIT"
   FUN = 0: CNT$ = "   ": GS$ = " ": L$ = " ": PGMRTRN$ = RTRIM$(ENV.PGMRTRN)
   FMTNAME$ = SPACE$(8): FILENAME$ = SPACE$(12): HSETUP$ = SPACE$(46)
   CODE$ = SPACE$(26): XCODE$ = SPACE$(124): XNAME$ = SPACE$(8)
   CALL RptLOAD("PRTC")
   CALL TableLOAD("PRINTERS", PRTTNUM, 0)
   CLSCRN = -1: CTYPE = 2: STYPE = 1
   GOSUB 5: GOTO 10
   DIM PRTSetup AS PrintSetupREC, PRTForms AS PrintFormsREC
   DIM PRTCodes AS PrintCodesREC, PRTSymbols AS PrintSymbolREC

5  ' Build XPSETUP$() Table from PSETUP$()
   PSUTable.FMTNAME = "PRTCS021": PSUTable.FMTNUM = 0
   PSUTable.ML = 46: PSUTable.XL = 46
   REDIM XPSETUP$(PDFHdr.SETUPS)
   FOR X = 1 TO PDFHdr.SETUPS
       XPSETUP$(X) = SPACE$(46)
       LSET XPSETUP$(X) = PSETUP$(X)
       MID$(XPSETUP$(X), 34, 8) = PSETUPS(X).PRINTER
       MID$(XPSETUP$(X), 42, 5) = PSETUPS(X).DEST
   NEXT X
   RETURN

10 ' Build PRINTER$() Table from "PRINTERS" System Table
   PRTRMAX = Table(PRTTNUM).SIZE
   REDIM PRINTER$(PRTRMAX), PRTCHN(PRTRMAX) AS ChainPTRS
   DIM PTABLE AS TablePARMS
   PTABLE.TABNAME = "PRINTERS"
   PTABLE.FMTNAME = "PRTCS031": PTABLE.FMTNUM = 0
   PTABLE.KEYNAME = "CODE": PTABLE.KL = 8: PTABLE.KO = 1
   PTABLE.SIZE = PRTRMAX
   PTABLE.XL = 33: PTABLE.ML = 33
   PTABLE.FT = 1: PTABLE.LT = PRTRMAX
   PTABLE.FF = 0: Y = Table(PRTTNUM).FT
   CALL ChnINIT(PRTCHN(), 1, PRTRMAX)
   FOR X = 1 TO PRTRMAX
      PRINTER$(X) = MID$(TABDATA$(Y), 1, 8) + MID$(TABDATA$(Y), 33, 25)
      Y = TABDATACHN(Y).FWD
   NEXT X

   ' Build CODESET Table
   CODEMAX = 16
   REDIM CODESET(CODEMAX) AS PrintCodesREC, CODE$(7)
   FOR X = 1 TO 7: CODE$(X) = SPACE$(25): NEXT X
   RESTORE 12
   FOR X = 1 TO CODEMAX
       READ N, PDFCodes.NAME
       FOR Y = 1 TO 7
           READ M: LSET CODE$(Y) = CHR$(M)
           Z = 1
           WHILE M > 0
              M = M - 1: Z = Z + 1
              READ N: MID$(CODE$(Y), Z, 1) = CHR$(N)
           WEND
       NEXT Y
       LSET PDFCodes.INIT = CODE$(1)
       LSET PDFCodes.FONT1 = CODE$(2)
       LSET PDFCodes.FONT2 = CODE$(3)
       LSET PDFCodes.FONT3 = CODE$(4)
       LSET PDFCodes.LINEFEED = CODE$(5)
       LSET PDFCodes.FormFeed = CODE$(6)
       LSET PDFCodes.RESET = CODE$(7)
       CODESET(X) = PDFCodes
   NEXT X
   REDIM CODE$(1)
   PDFCodes = PCODES(ENV.CODES)
12 DATA 1,"NOCODES",0,0,0,0,2,13,10,0,0
   DATA 2,"EPSON",9,27,64,27,67,0,11,27,51,33,1,15,0,0,1,13,1,12,2,27,64
   DATA 3,"IBMGRAPH",7,27,67,0,11,27,51,33,1,15,0,0,1,13,1,12,0
   DATA 4,"PROPRINT",9,27,67,0,11,27,51,33,27,54,5,18,27,73,0,15,0,0,1,13,1,12,0
   DATA 5,"LASERPLS",10,27,38,107,50,83,27,38,107,55,72,5,27,38,108,55,67,0,0,1,13,1,12,0
   DATA 6,"CITOH",2,27,64,2,27,81,0,0,1,13,1,12,2,27,64
   DATA 7,"EPSONLQ",5,27,64,27,116,1,4,27,80,27,15,0,0,1,13,1,12,2,27,64
   DATA 8,"STARNX",6,27,67,66,27,53,0,4,27,77,27,15,0,0,1,13,1,12,0
   DATA 9,"GEMINI",9,27,64,27,67,0,11,27,65,11,1,15,0,0,1,13,1,12,2,27,64
   DATA 10,"TANDYDMP",8,27,52,66,27,64,22,27,58,2,27,20,0,0,1,13,1,12,0
   DATA 11,"NEC5200",9,27,64,27,116,1,27,67,0,11,5,27,15,28,51,55,0,0,1,13,1,12,2,27,64
   DATA 12,"HPDSKJET",8,27,69,27,40,115,50,48,72,05,27,38,108,56,68,0,0,1,13,1,12,0
   DATA 13,"OKIDATA",7,27,70,66,27,37,57,20,1,29,0,0,1,13,1,12,0
   DATA 14,"HPLASER",10,27,38,107,50,83,27,40,49,48,85,5,27,38,108,56,68,0,0,1,13,1,12,0
'   DATA 15,"CANONBJC",5,27,54,27,65,11,2,27,15,0,0,1,13,1,12,0
   DATA 15,"CANONBJC",0,2,27,15,0,0,1,13,1,12,0
   DATA 16,"EPSONSTY",4,27,64,27,48,3,27,80,15,0,0,1,13,1,12,2,27,64

   ' Build SYMBSET Table
   SYMBMAX = 5: RESTORE 15
   REDIM SYMBSET(SYMBMAX) AS PrintSymbolREC
   FOR X = 1 TO SYMBMAX
       READ N, PDFSymbols.NAME, X$
       READ PDFSymbols.SET1
       READ PDFSymbols.SET2
       READ PDFSymbols.SET3
       SYMBSET(X) = PDFSymbols
   NEXT X
   PDFSymbols = PSYMBOLS(ENV.SYMBOLS)
15 DATA 1,"TYPEWRIT","Typewriter Symbol Set"
   DATA "+++!++-++++","+++!++-+++","+++!++-+++"
16 DATA 2,"CP437","Code Page 437 Symbol Set"
   DATA "ÚÁ¿³´ÃÄÀÂÙÅ","ÚÁ¿³´ÃÄÀÂÙ","ÉÏ»º¶ÇÍÈÑ¼"
17 DATA 3,"CP850","Code Page 850 Symbol Set"
   DATA "ÚÁ¿³´ÃÄÀÂÙÅ","ÚÁ¿³´ÃÄÀÂÙ","ÚÁ¿³´ÃÄÀÂÙ"
18 DATA 4,"GEMINI","Gemini Printer Symbol Set"
   DATA "ÚÁ¿³´ÃÄÀÂÙÅ","ÚÁ¿³´ÃÄÀÂÙ","ÉÏ»º¶ÇÍÈÑ¼"
19 DATA 5,"LASERPLS","HP LaserJet+ Symbol Set"
   DATA "+++|++-++++","+++|++-+++","+++|++-+++"

40 ' Get FORMAT Numbers
   CALL FmtFIND("PRTCS001"): FMT001 = CURFMT
   CALL FmtFIND("PRTCS005"): FMT005 = CURFMT
   CALL FmtFIND("PRTCS010"): FMT010 = CURFMT
   CALL FmtFIND("PRTCS011"): FMT011 = CURFMT
   CALL FmtFIND("PRTCS012"): FMT012 = CURFMT
   CALL FmtFIND("PRTCS020"): FMT020 = CURFMT
   CALL FmtFIND("PRTCS021"): FMT021 = CURFMT
   CALL FmtFIND("PRTCS030"): FMT030 = CURFMT
   CALL FmtFIND("PRTCS031"): FMT031 = CURFMT
   CALL FmtFIND("PRTCS040"): FMT040 = CURFMT
   CALL FmtFIND("PRTCS041"): FMT041 = CURFMT
   CALL FmtFIND("PRTCS042"): FMT042 = CURFMT
   CALL FmtFIND("PRTCS043"): FMT043 = CURFMT
   CALL FmtFIND("PRTCS046"): FMT046 = CURFMT
   CALL FmtFIND("PRTCS047"): FMT047 = CURFMT
   CALL FmtFIND("PRTCS050"): FMT050 = CURFMT
   CALL FmtFIND("PRTCS091"): FMT091 = CURFMT
   CALL FmtFIND("PRTCS092"): FMT092 = CURFMT
   CALL FmtFIND("PRTCS093"): FMT093 = CURFMT
   CALL FmtFIND("PRTCS094"): FMT094 = CURFMT
   CALL FmtFIND("PRTCS095"): FMT095 = CURFMT
   CALL WinINIT("UH2")
   CALL WinINIT("LL2"): CALL WinINIT("LR2")
   CALL WinINIT("LH2")
   CALL WinPREP(1)

50 ' Format Display for INITIAL Screen
   IF CLSCRN THEN CALL WinCLR: CLSCRN = 0
   CALL WinSWITCH(2)
   GOSUB 100: GOSUB 110
   CALL OptDISPLAY(FMT091)
   GOSUB 120: GOSUB 130
   GOTO 200

' Routines for Processing PDF Header and SETUP Data in Format Fields
'   There are separate routines for File to Screen and Screen to File

100 ' Move Print Definition File HEADER DATA to Format
    CALL WinPREP(2)
    FMT = FMTTAB(FMT001)
    CALL FmtFindFLD("PDF"): IF FOUND THEN LSET UTXT$(FFLD.UTXT) = CONFIG.PDF
    CALL FmtFindFLD("SCNT")
    IF FOUND THEN LSET UTXT$(FFLD.UTXT) = LTRIM$(STR$(PDFHdr.SETUPS))
    CALL FmtFindFLD("FSIZ")
    IF FOUND THEN LSET UTXT$(FFLD.UTXT) = LTRIM$(STR$(PDFHdr.EOFRBA))
101 ' Process DEFAULT Setup Fields
    FMT = FMTTAB(FMT001)
    CALL FmtFindFLD("DPSU")
    IF FOUND THEN LSET UTXT$(FFLD.UTXT) = PSETUPS(PDFHdr.DEFAULT).NAME
    CALL FmtFindFLD("DPSD")
    IF FOUND THEN LSET UTXT$(FFLD.UTXT) = PSETUPS(PDFHdr.DEFAULT).DESC
    CALL FmtFindFLD("DPBS")
    IF FOUND THEN LSET UTXT$(FFLD.UTXT) = PSETUPS(PDFHdr.DEFAULT).PRINTER
    CALL FmtFindFLD("DPDE")
    IF FOUND THEN LSET UTXT$(FFLD.UTXT) = PSETUPS(PDFHdr.DEFAULT).DEST
102 ' Move CURRENT Setup Descriptive DATA to Format
    FMT = FMTTAB(FMT001)
    CALL FmtFindFLD("CPSU"): IF FOUND THEN LSET UTXT$(FFLD.UTXT) = PDFSetup.NAME
    CALL FmtFindFLD("CPSD"): IF FOUND THEN LSET UTXT$(FFLD.UTXT) = PDFSetup.DESC
    CALL FmtFindFLD("CPBS"): IF FOUND THEN LSET UTXT$(FFLD.UTXT) = PDFSetup.PRINTER
    CALL WinFORMAT(FMT001)
    RETURN

110 ' Move PDFSetup PARMS and OPTIONS to Format
    FMT = FMTTAB(FMT005)
    CALL GetUFLD("FIRST"): FF = Win.CURFLD
111 ' Process Next UFLD
    SELECT CASE FFLD.NAME
       CASE "CPDE": LSET UTXT$(FFLD.UTXT) = PDFSetup.DEST
       CASE "FMWI": RSET UTXT$(FFLD.UTXT) = LTRIM$(STR$(PDFFORMS.FWIDTH))
       CASE "FMLT": RSET UTXT$(FFLD.UTXT) = LTRIM$(STR$(PDFFORMS.FLENGTH))
       CASE "TMGN": RSET UTXT$(FFLD.UTXT) = LTRIM$(STR$(PDFFORMS.TMARGIN))
       CASE "BMGN": RSET UTXT$(FFLD.UTXT) = LTRIM$(STR$(PDFFORMS.BMARGIN))
       CASE "IMGN": RSET UTXT$(FFLD.UTXT) = LTRIM$(STR$(PDFFORMS.IMARGIN))
       CASE "OMGN": RSET UTXT$(FFLD.UTXT) = LTRIM$(STR$(PDFFORMS.OMARGIN))
       CASE "OPSS": IF PDFFORMS.OPTIONS AND 1 THEN O$ = SNGLKEY$(1) ELSE O$ = SNGLKEY$(2)
                    LSET UTXT$(FFLD.UTXT) = O$
       CASE "OPFF": IF PDFFORMS.OPTIONS AND 2 THEN O$ = SNGLKEY$(1) ELSE O$ = SNGLKEY$(2)
                    LSET UTXT$(FFLD.UTXT) = O$
       CASE "OPPW": IF PDFFORMS.OPTIONS AND 4 THEN O$ = SNGLKEY$(1) ELSE O$ = SNGLKEY$(2)
                    LSET UTXT$(FFLD.UTXT) = O$
    END SELECT
    CALL GetUFLD("NEXT")
    IF Win.CURFLD <> FF THEN GOTO 111
    CALL WinFORMAT(FMT005)
    RETURN

115 ' Move PDFSetup DEST and FORMS Parms From Format
    FMT = FMTTAB(FMT005)
    CALL GetUFLD("FIRST"): FF = Win.CURFLD
116 ' Process Next UFLD
    SELECT CASE FFLD.NAME
       CASE "CPDE": LSET PDFSetup.DEST = UTXT$(FFLD.UTXT)
       CASE "FMWI": PDFFORMS.FWIDTH = VAL(UTXT$(FFLD.UTXT))
       CASE "FMLT": PDFFORMS.FLENGTH = VAL(UTXT$(FFLD.UTXT))
       CASE "TMGN": PDFFORMS.TMARGIN = VAL(UTXT$(FFLD.UTXT))
       CASE "BMGN": PDFFORMS.BMARGIN = VAL(UTXT$(FFLD.UTXT))
       CASE "IMGN": PDFFORMS.IMARGIN = VAL(UTXT$(FFLD.UTXT))
       CASE "OMGN": PDFFORMS.OMARGIN = VAL(UTXT$(FFLD.UTXT))
       CASE "OPSS"
           IF UTXT$(FFLD.UTXT) = SNGLKEY$(1) THEN
              PDFFORMS.OPTIONS = (PDFFORMS.OPTIONS OR 1)
              ELSE
              PDFFORMS.OPTIONS = (PDFFORMS.OPTIONS AND 254)
           END IF
       CASE "OPFF"
           IF UTXT$(FFLD.UTXT) = SNGLKEY$(1) THEN
              PDFFORMS.OPTIONS = (PDFFORMS.OPTIONS OR 2)
              ELSE
              PDFFORMS.OPTIONS = (PDFFORMS.OPTIONS AND 253)
           END IF
       CASE "OPPW"
           IF UTXT$(FFLD.UTXT) = SNGLKEY$(1) THEN
              PDFFORMS.OPTIONS = (PDFFORMS.OPTIONS OR 4)
              ELSE
              PDFFORMS.OPTIONS = (PDFFORMS.OPTIONS AND 251)
           END IF
    END SELECT
    CALL GetUFLD("NEXT")
    IF Win.CURFLD <> FF THEN GOTO 116
    RETURN

120 ' Move Printer CODES Data to Format
    CALL WinPREP(3)
    CALL WinFORMAT(FMT010)
    SELECT CASE CTYPE
       CASE 1: CURFMT = FMT011
       CASE 2: CURFMT = FMT012
    END SELECT
    FMT = FMTTAB(CURFMT)
    CALL GetUFLD("FIRST"): FF = Win.CURFLD
121 ' Process Next Printer CODE Entry
    SELECT CASE FFLD.NAME
       CASE "INIT": LSET CODE$ = PDFCodes.INIT
       CASE "FNT1": LSET CODE$ = PDFCodes.FONT1
       CASE "FNT2": LSET CODE$ = PDFCodes.FONT2
       CASE "FNT3": LSET CODE$ = PDFCodes.FONT3
       CASE "CRLF": LSET CODE$ = PDFCodes.LINEFEED
       CASE "FMFD": LSET CODE$ = PDFCodes.FormFeed
       CASE "REST": LSET CODE$ = PDFCodes.RESET
    END SELECT
    SELECT CASE CTYPE
       CASE 1: CALL Code2Hex(CODE$, XCODE$)
       CASE 2: CALL Code2Dec(CODE$, XCODE$)
       CASE ELSE: RETURN
    END SELECT
    LSET UTXT$(FFLD.UTXT) = XCODE$
    CALL GetUFLD("NEXT")
    IF Win.CURFLD <> FF THEN GOTO 121
    CALL WinFORMAT(0)
    RETURN

125 ' Move Printer CODES Data From Format
    SELECT CASE CTYPE
       CASE 1: FMT = FMTTAB(FMT011)
       CASE 2: FMT = FMTTAB(FMT012)
    END SELECT
    CALL GetUFLD("FIRST"): FF = Win.CURFLD
126 ' Process Next Printer CODE Entry
    LSET XCODE$ = UTXT$(FFLD.UTXT)
    SELECT CASE CTYPE
       CASE 1: CALL Hex2Code(XCODE$, CODE$)
       CASE 2: CALL Dec2Code(XCODE$, CODE$)
       CASE ELSE: RETURN
    END SELECT
    SELECT CASE FFLD.NAME
       CASE "INIT": LSET PDFCodes.INIT = CODE$
       CASE "FNT1": LSET PDFCodes.FONT1 = CODE$
       CASE "FNT2": LSET PDFCodes.FONT2 = CODE$
       CASE "FNT3": LSET PDFCodes.FONT3 = CODE$
       CASE "CRLF": LSET PDFCodes.LINEFEED = CODE$
       CASE "FMFD": LSET PDFCodes.FormFeed = CODE$
       CASE "REST": LSET PDFCodes.RESET = CODE$
    END SELECT
    CALL GetUFLD("NEXT")
    IF Win.CURFLD <> FF THEN GOTO 126
    RETURN

130 ' Move SYMBOL Set DATA to Format
    CALL WinPREP(4)
    CALL WinFORMAT(FMT040)
    SELECT CASE STYPE
       CASE 1: CURFMT = FMT041
       CASE 2: CURFMT = FMT042
       CASE 3: CURFMT = FMT043
    END SELECT
    CALL WinFORMAT(0)
    SELECT CASE STYPE
       CASE 1: CURFMT = FMT046: LSET CODE$ = PDFSymbols.SET1
       CASE 2: CURFMT = FMT047: LSET CODE$ = PDFSymbols.SET2
       CASE 3: CURFMT = FMT047: LSET CODE$ = PDFSymbols.SET3
    END SELECT
    FMT = FMTTAB(CURFMT)
    Z = 0
    CALL GetUFLD("FIRST"): FF = Win.CURFLD
131 ' Process Next SYMBOL
    Z = Z + 1
    RSET UTXT$(FFLD.UTXT) = LTRIM$(STR$(ASC(MID$(CODE$, Z, 1))))
    CALL GetUFLD("NEXT")
    IF Win.CURFLD <> FF THEN GOTO 131
    CALL WinFORMAT(0)
    RETURN

135 ' Move SYMBOL Set DATA From Format
    SELECT CASE STYPE
       CASE 1: CURFMT = FMT046: LSET CODE$ = PDFSymbols.SET1
       CASE 2: CURFMT = FMT047: LSET CODE$ = PDFSymbols.SET2
       CASE 3: CURFMT = FMT047: LSET CODE$ = PDFSymbols.SET3
    END SELECT
    FMT = FMTTAB(CURFMT)
    Z = 0
    CALL GetUFLD("FIRST"): FF = Win.CURFLD
136 ' Process Next SYMBOL
    Z = Z + 1
    C = VAL(UTXT$(FFLD.UTXT))
    IF (C > 0) AND (C < 256) THEN MID$(CODE$, Z, 1) = CHR$(C)
    CALL GetUFLD("NEXT")
    IF Win.CURFLD <> FF THEN GOTO 136
    SELECT CASE STYPE
       CASE 1: LSET PDFSymbols.SET1 = CODE$
       CASE 2: LSET PDFSymbols.SET2 = CODE$
       CASE 3: LSET PDFSymbols.SET3 = CODE$
    END SELECT
    GOSUB 130
    RETURN

200 ' Get Processing Option
    CALL WinSWITCH(2)
    CALL PutMSG(" ")
    CALL OptDISPLAY(FMT091)
210 ' Get Request
    CALL GetKEY: FUN$ = ""
    SELECT CASE A
       CASE 59: FUN$ = "F1"
       CASE 60: FUN$ = "F2"
       CASE 61: FUN$ = "F3"
       CASE 62: FUN$ = "F4"
       CASE 64: FUN$ = "F6"
       CASE 66: CALL DOSShell: GOTO 50
       CASE 67: FUN$ = "F9": ENV.RPTIX = 0
       CASE 109: IF PGMRTRN$ = "FHSMENU" THEN FUN$ = "F6": PGMRTRN$ = "FHSRPTS"
    END SELECT
    IF FUN$ = "" THEN CALL ErrBEEP(0): GOTO 210
    CALL OptHILITE(FUN$, "", 6)
    SELECT CASE FUN$
       CASE "F1": GOTO 500   ' Process FILE Request
       CASE "F2": GOTO 300   ' CHANGE Current Setup Information
       CASE "F3": GOTO 250   ' Process SETUPS Table
       CASE "F4": GOTO 220   ' Select DEFAULT Setup
       CASE "F6": IF A = 109 THEN GOTO 1000
                  GOTO 600   ' Process PRINT Request
       CASE "F9": GOTO 1000  ' Return
    END SELECT

220 ' Select DEFAULT Setup
    CALL WinSWITCH(5)
    CALL WinFORMAT(FMT020)
    CALL OptDISPLAY(FMT095)
    CALL TableOPEN(PSUTable)
    CALL TableSHOW(PSUTable, XPSETUP$(), PSUCHN())
    CALL TableSELECT(PSUTable, XPSETUP$(), PSUCHN())
    IF A <> 27 THEN
       LSET XNAME$ = XPSETUP$(PSUTable.CT)
       X = PDFHdr.SETUPS: FOUND = 0
       WHILE (X > 0) AND (NOT FOUND)
          FOUND = (PSETUPS(X).NAME = XNAME$)
          IF NOT FOUND THEN X = X - 1
       WEND
       PDFHdr.DEFAULT = X
    END IF
    CLSCRN = -1
    GOTO 50

250 ' Process SETUPS Table
    CALL WinOPEN("LH2")
    CALL WinFORMAT(FMT020)
    CALL TableOPEN(PSUTable)
252 ' Display List of SETUPS
    CALL TableSHOW(PSUTable, XPSETUP$(), PSUCHN())
255 ' Get Next SETUPS Request
    CALL PutMSG("")
    CALL OptDISPLAY(FMT094)
    CALL TableSELECT(PSUTable, XPSETUP$(), PSUCHN())
    CT = PSUTable.CT
    LSET XNAME$ = XPSETUP$(CT)
    X = PDFHdr.SETUPS: FOUND = 0
    WHILE (X > 0) AND (NOT FOUND)
       FOUND = (PSETUPS(X).NAME = XNAME$)
       IF NOT FOUND THEN X = X - 1
    WEND
    CTSETUP = X
    LSET HSETUP$ = XPSETUP$(CT)
    SELECT CASE A
       CASE 13: FUN$ = "SEL"
       CASE 27: FUN$ = "ESC"
       CASE 59: FUN$ = "CHG"
       CASE 60: FUN$ = "ADD"
       CASE 61: FUN$ = "DEL"
       CASE 62: FUN$ = "MOV"
       CASE 67: FUN$ = "ESC"
       CASE ELSE: CALL ErrBEEP(0): GOTO 255
    END SELECT
    CALL OptHILITE(FUN$, "", 6)
    SELECT CASE FUN$
       CASE "CHG": GOTO 260   ' Modify SETUP Name and Description
       CASE "ADD": GOTO 265   ' ADD New SETUP
       CASE "DEL": GOTO 270   ' Delete This SETUP
       CASE "MOV": GOTO 280   ' Move This SETUP
    END SELECT
    GOTO 299

260 ' Modify SETUP Entry
    CALL TableUPDT("CHG", PSUTable, XPSETUP$(), PSUCHN())
    IF A = 27 THEN
       XPSETUP$(CT) = HSETUP$
       GOTO 252
    END IF
    LSET XNAME$ = XPSETUP$(CT)
    IF XNAME$ = SPACE$(8) THEN CALL PutMSG("PRTCM001"): GOTO 260
    IF XNAME$ <> PSETUPS(CTSETUP).NAME THEN
       X = PDFHdr.SETUPS: FOUND = 0
       WHILE (X > 0) AND (NOT FOUND)
          FOUND = (PSETUPS(X).NAME = XNAME$)
          IF NOT FOUND THEN X = X - 1
       WEND
       IF FOUND THEN CALL PutMSG("PRTCM002"): GOTO 260
       PSETUPS(CTSETUP).NAME = XNAME$
    END IF
    PSETUPS(CTSETUP).DESC = MID$(XPSETUP$(CT), 9, 25)
    LSET HSETUP$ = XPSETUP$(CT): XPSETUP$(CT) = HSETUP$
    IF (PDFHdr.DEFAULT = CTSETUP) OR (ENV.PSETUP = CTSETUP) THEN
       CALL WinPREP(2)
       IF PDFHdr.DEFAULT = CTSETUP THEN GOSUB 100
       IF ENV.PSETUP = CTSETUP THEN
          PDFSetup = PSETUPS(CTSETUP)
          GOSUB 101: GOSUB 110
       END IF
       CALL WinPREP(5)
    END IF
    GOTO 255

265 ' Add New SETUP
    IF PSUTable.FF = 0 THEN
       X = PDFHdr.SETUPS + 1
       REDIM PRESERVE PSETUPS(X) AS PrintSetupREC, PSUCHN(X) AS ChainPTRS
       REDIM PRESERVE XPSETUP$(X)
       XPSETUP$(X) = SPACE$(46)
       PSUTable.FF = X
    END IF
    CALL TableUPDT("ADD", PSUTable, XPSETUP$(), PSUCHN())
    IF A = 27 THEN GOTO 255
    PDFHdr.SETUPS = PDFHdr.SETUPS + 1
    CTSETUP = PDFHdr.SETUPS
    CT = PSUTable.CT
    PDFSetup = PSETUPS(ENV.PSETUP)
    PDFSetup.NAME = MID$(XPSETUP$(CT), 1, 8)
    PDFSetup.DESC = MID$(XPSETUP$(CT), 9, 25)
    PDFSetup.PRINTER = PRINTERS(ENV.PRINT).NAME
    PDFSetup.FORMS = PFORMS(ENV.FORMS).NAME
    PDFSetup.CODES = PCODES(ENV.CODES).NAME
    PDFSetup.SYMBOLS = PSYMBOLS(ENV.SYMBOLS).NAME
    PSETUPS(CTSETUP) = PDFSetup
    LSET XNAME$ = PDFSetup.NAME
    CALL PDFGetSet(XNAME$)
    PRINTERS(ENV.PRINT).REF = PRINTERS(ENV.PRINT).REF + 1
    PFORMS(ENV.FORMS).REF = PFORMS(ENV.FORMS).REF + 1
    PCODES(ENV.CODES).REF = PCODES(ENV.CODES).REF + 1
    PSYMBOLS(ENV.SYMBOLS).REF = PSYMBOLS(ENV.SYMBOLS).REF + 1
    LSET HSETUP$ = XPSETUP$(CT): XPSETUP$(CT) = HSETUP$
    MID$(XPSETUP$(CT), 34, 8) = PDFSetup.PRINTER
    MID$(XPSETUP$(CT), 42, 5) = PDFSetup.DEST
    GOTO 299

270 ' Delete This SETUP
    IF CT = PDFHdr.DEFAULT THEN CALL PutMSG("PRTCM004"): GOTO 255
    IF CT = ENV.PSETUP THEN CALL PutMSG("PRTCM005"): GOTO 255
    LSET SYSVAR.MVAR = LEFT$(XPSETUP$(CT), 8)
    CALL PutMSG("PRTCM006")
    IF A$ <> SNGLKEY$(1) THEN GOTO 255
    PDFHdr.VER = CHR$(255)
    HSETUP = ENV.PSETUP
    LSET XNAME$ = XPSETUP$(CT)
    CALL PDFGetSet(XNAME$)
    PRINTERS(ENV.PRINT).REF = PRINTERS(ENV.PRINT).REF - 1
    PFORMS(ENV.FORMS).REF = PFORMS(ENV.FORMS).REF - 1
    PCODES(ENV.CODES).REF = PCODES(ENV.CODES).REF - 1
    PSYMBOLS(ENV.SYMBOLS).REF = PSYMBOLS(ENV.SYMBOLS).REF - 1
    LSET XNAME$ = PSETUPS(HSETUP).NAME
    CALL PDFGetSet(XNAME$)
    X = CTSETUP
    WHILE X < PDFHdr.SETUPS
       PSETUPS(X) = PSETUPS(X + 1)
       X = X + 1
    WEND
    IF PDFHdr.DEFAULT > CTSETUP THEN PDFHdr.DEFAULT = PDFHdr.DEFAULT - 1
    CALL TableUPDT("DEL", PSUTable, XPSETUP$(), PSUCHN())
    PDFHdr.SETUPS = PDFHdr.SETUPS - 1
    GOTO 255

280 ' Move This SETUP
    CALL TableUPDT("MOV", PSUTable, XPSETUP$(), PSUCHN())
    CT = PSUTable.CT
    GOTO 255

299 ' Return from SETUPS
    IF (FUN$ = "SEL") THEN
       LSET XNAME$ = PSETUPS(CTSETUP).NAME
       CALL PDFGetSet(XNAME$)
    END IF
    CALL WinCLOSE(0, 2)
    GOSUB 130: GOSUB 120
    CALL WinSWITCH(2)
    GOSUB 100: GOSUB 110
    IF (FUN$ = "ADD") THEN GOTO 300
    GOTO 200

300 ' Update Current SETUP Parms
    CALL WinSWITCH(2)
    CALL PutMSG("")
    CALL OptDISPLAY(FMT093)
301 CALL GetKEY
    IF A = 27 THEN GOTO 200
    SELECT CASE A
       CASE 59: FUN$ = "PRT"
       CASE 60: FUN$ = "FORM"
       CASE 61: FUN$ = "CODE"
       CASE 62: FUN$ = "SYMB"
       CASE ELSE: CALL ErrBEEP(0): GOTO 301
    END SELECT
    CALL OptHILITE(FUN$, "", 6)
    SELECT CASE FUN$
       CASE "PRT": GOTO 330
       CASE "FORM": GOTO 305
       CASE "CODE": GOTO 310
       CASE "SYMB": GOTO 320
    END SELECT

305 ' Update Current DEST and FORMS
    CALL PutMSG("XXXXM012")
    FMT = FMTTAB(FMT005): CALL FmtUHOLD
    Win.CLIN = 0
    CALL WinUPDATE
    IF A <> 27 THEN
       GOSUB 115
       PDFHdr.VER = CHR$(255)
       PDFSetup.FORMS = PDFSetup.NAME
       PSETUPS(ENV.PSETUP) = PDFSetup
       IF PDFFORMS.NAME <> PDFSetup.NAME THEN
          PDFFORMS.NAME = PDFSetup.NAME
          PFORMS(ENV.FORMS).REF = PFORMS(ENV.FORMS).REF - 1
          X = 0: FOUND = 0
          WHILE (X < PDFHdr.FORMS) AND (NOT FOUND)
             X = X + 1: FOUND = (PFORMS(X).NAME = PDFFORMS.NAME)
          WEND
          IF FOUND THEN
             PDFFORMS.REF = PFORMS(X).REF + 1
             ELSE
             X = X + 1
             PDFHdr.FORMS = X
             PDFFORMS.REF = 1
             REDIM PRESERVE PFORMS(X) AS PrintFormsREC
          END IF
          ENV.FORMS = X
       END IF
       PFORMS(ENV.FORMS) = PDFFORMS
       X = PSUTable.FT: FOUND = 0
       WHILE (X > 0) AND (NOT FOUND)
          FOUND = (MID$(XPSETUP$(X), 1, 8) = PDFSetup.NAME)
          IF NOT FOUND THEN X = PSUCHN(X).FWD
       WEND
       MID$(XPSETUP$(X), 42, 5) = PDFSetup.DEST
    END IF
    GOSUB 110
    GOSUB 100
    GOTO 300

310 ' Update Current Printer CODES
    CALL WinSWITCH(3)
    SELECT CASE CTYPE
       CASE 1: CURFMT = FMT011
       CASE 2: CURFMT = FMT012
    END SELECT
    FMT = FMTTAB(CURFMT)
    CALL FmtUHOLD
    CALL WinUPDATE
    IF A = 60 THEN FUN$ = "CHG" ELSE FUN$ = ""
    GOSUB 125
    PDFHdr.VER = CHR$(255)
    PDFSetup.CODES = PDFSetup.NAME
    PSETUPS(ENV.PSETUP) = PDFSetup
    IF PDFCodes.NAME <> PDFSetup.NAME THEN
       PDFCodes.NAME = PDFSetup.CODES
       PCODES(ENV.CODES).REF = PCODES(ENV.CODES).REF - 1
       X = 0: FOUND = 0
       WHILE (X < PDFHdr.CODES) AND (NOT FOUND)
          X = X + 1: FOUND = (PCODES(X).NAME = PDFCodes.NAME)
       WEND
       IF FOUND THEN
          PDFCodes.REF = PCODES(X).REF + 1
          ELSE
          X = X + 1
          PDFHdr.CODES = X
          PDFCodes.REF = 1
          REDIM PRESERVE PCODES(X) AS PrintCodesREC
       END IF
       ENV.CODES = X
    END IF
    PCODES(ENV.CODES) = PDFCodes
    IF FUN$ = "CHG" THEN
       CTYPE = CTYPE + 1: IF CTYPE > 2 THEN CTYPE = 1
       GOSUB 120
       GOTO 310
    END IF
    GOSUB 120
    GOTO 300

320 ' Update Current Printer SYMBOLS
    CALL WinSWITCH(4)
    SELECT CASE STYPE
       CASE 1: CURFMT = FMT046
       CASE 2: CURFMT = FMT047
       CASE 3: CURFMT = FMT047
    END SELECT
    FMT = FMTTAB(CURFMT)
    CALL FmtUHOLD
    CALL WinUPDATE
    SELECT CASE A
       CASE 60: FUN$ = "F2"
       CASE 73: FUN$ = "PGUP"
       CASE 81: FUN$ = "PGDN"
       CASE ELSE: FUN$ = "END"
    END SELECT
    GOSUB 135
    SELECT CASE FUN$
       CASE "F2": GOTO 321
       CASE "END": GOTO 322
       CASE "PGUP": STYPE = STYPE - 1: IF STYPE < 1 THEN STYPE = 3
       CASE "PGDN": STYPE = STYPE + 1: IF STYPE > 3 THEN STYPE = 1
    END SELECT
    GOSUB 130
    GOTO 320

321 ' Load Standard Symbol Set
    Win.OFMT = FMT040: CALL OptHILITE(FUN$, "", 6)
    CALL PutMSG("PRTCM012")
    SELECT CASE A$
       CASE "1": PDFSymbols = SYMBSET(1)
       CASE "2": PDFSymbols = SYMBSET(2)
       CASE "3": PDFSymbols = SYMBSET(3)
       CASE ELSE: CALL ErrBEEP(0): GOTO 320
    END SELECT
    GOSUB 130
    GOTO 320


322 ' Save and Return
    PDFSetup.SYMBOLS = PDFSetup.NAME
    PSETUPS(ENV.PSETUP) = PDFSetup
    PDFHdr.VER = CHR$(255)
    IF PDFSymbols.NAME <> PDFSetup.NAME THEN
       PDFSymbols.NAME = PDFSetup.NAME
       PSYMBOLS(ENV.SYMBOLS).REF = PSYMBOLS(ENV.SYMBOLS).REF - 1
       X = 0: FOUND = 0
       WHILE (X < PDFHdr.SYMBOLS) AND (NOT FOUND)
          X = X + 1: FOUND = (PSYMBOLS(X).NAME = PDFSymbols.NAME)
       WEND
       IF FOUND THEN
          PDFSymbols.REF = PSYMBOLS(X).REF + 1
          ELSE
          X = X + 1
          PDFHdr.SYMBOLS = X
          PDFSymbols.REF = 1
          REDIM PRESERVE PSYMBOLS(X) AS PrintSymbolREC
       END IF
       ENV.SYMBOLS = X
    END IF
    PSYMBOLS(ENV.SYMBOLS) = PDFSymbols
    GOTO 300

330 ' Select PRINTER
    CALL WinOPEN("LR2")
    CALL WinFORMAT(FMT030)
    CALL OptDISPLAY(FMT092)
    CALL TableOPEN(PTABLE)
    CALL TableSHOW(PTABLE, PRINTER$(), PRTCHN())
331 CALL TableSELECT(PTABLE, PRINTER$(), PRTCHN())
    IF A <> 13 THEN
       CALL WinCLOSE(0, 2): GOSUB 130: GOTO 300
    END IF
    LSET XNAME$ = LEFT$(PRINTER$(PTABLE.CT), 8)
    Y = Table(PRTTNUM).FT: FOUND = 0
    WHILE (Y > 0) AND (NOT FOUND)
       FOUND = (XNAME$ = MID$(TABDATA$(Y), 1, 8))
       IF NOT FOUND THEN Y = TABDATACHN(Y).FWD
    WEND
    IF NOT FOUND THEN CALL ErrBEEP(0): GOTO 331
' Process PRINTERS()
    PDFSetup.PRINTER = XNAME$
    PDFPrinter.NAME = XNAME$
    PDFPrinter.CODES = MID$(TABDATA$(Y), 9, 8)
    PDFPrinter.SYMBOLS = MID$(TABDATA$(Y), 17, 8)
    PDFPrinter.FORMS = MID$(TABDATA$(Y), 25, 8)
    PRINTERS(ENV.PRINT).REF = PRINTERS(ENV.PRINT).REF - 1
    W = 0: FOUND = 0
    WHILE (W < PDFHdr.PRINTERS) AND (NOT FOUND)
       W = W + 1: FOUND = (PDFPrinter.NAME = PRINTERS(W).NAME)
    WEND
    IF FOUND THEN
       PDFPrinter.REF = PRINTERS(W).REF + 1
       ELSE
       W = W + 1
       PDFHdr.PRINTERS = W
       REDIM PRESERVE PRINTERS(W) AS PrintDefREC
       PDFPrinter.REF = 1
    END IF
    ENV.PRINT = W
    PRINTERS(ENV.PRINT) = PDFPrinter
' Process PCODES()
    PCODES(ENV.CODES).REF = PCODES(ENV.CODES).REF - 1
    W = CODEMAX: FOUND = 0
    WHILE (W > 0) AND (NOT FOUND)
       FOUND = (PDFPrinter.CODES = CODESET(W).NAME)
       IF NOT FOUND THEN W = W - 1
    WEND
    PDFCodes = CODESET(W)
    LSET PDFSetup.CODES = " "
    W = 0: FOUND = 0
    WHILE (W < PDFHdr.CODES) AND (NOT FOUND)
       W = W + 1: FOUND = (PCODES(W).NAME = PDFCodes.NAME)
    WEND
    IF FOUND THEN
       PDFCodes.REF = PCODES(W).REF + 1
       ELSE
       W = W + 1
       PDFHdr.CODES = W
       REDIM PRESERVE PCODES(W) AS PrintCodesREC
       PDFCodes.REF = 1
    END IF
    ENV.CODES = W
    PCODES(ENV.CODES) = PDFCodes
' Process PSYMBOLS()
    PSYMBOLS(ENV.SYMBOLS).REF = PSYMBOLS(ENV.SYMBOLS).REF - 1
    W = SYMBMAX: FOUND = 0
    WHILE (W > 0) AND (NOT FOUND)
       FOUND = (PDFPrinter.SYMBOLS = SYMBSET(W).NAME)
       IF NOT FOUND THEN W = W - 1
    WEND
    PDFSymbols = SYMBSET(W)
    LSET PDFSetup.SYMBOLS = " "
    W = 0: FOUND = 0
    WHILE (W < PDFHdr.SYMBOLS) AND (NOT FOUND)
       W = W + 1: FOUND = (PSYMBOLS(W).NAME = PDFSymbols.NAME)
    WEND
    IF FOUND THEN
       PDFSymbols.REF = PSYMBOLS(W).REF + 1
       ELSE
       W = W + 1
       PDFHdr.SYMBOLS = W
       REDIM PRESERVE PSYMBOLS(W) AS PrintSymbolREC
       PDFSymbols.REF = 1
    END IF
    ENV.SYMBOLS = W
    PSYMBOLS(ENV.SYMBOLS) = PDFSymbols
' Process PFORMS()
    PFORMS(ENV.FORMS).REF = PFORMS(ENV.FORMS).REF - 1
    LSET PDFFORMS.NAME = PDFPrinter.FORMS
    LSET PDFSetup.FORMS = " "
    W = 0: FOUND = 0
    WHILE (W < PDFHdr.FORMS) AND (NOT FOUND)
       W = W + 1: FOUND = (PDFFORMS.NAME = PFORMS(W).NAME)
    WEND
    IF FOUND THEN
       PDFFORMS = PFORMS(W)
       ELSE
       W = W + 1
       PDFHdr.FORMS = W
       REDIM PRESERVE PFORMS(W) AS PrintFormsREC
       PDFFORMS.REF = 0
       PDFFORMS.FWIDTH = VAL(MID$(PDFFORMS.NAME, 2, 3))
       PDFFORMS.FLENGTH = VAL(MID$(PDFFORMS.NAME, 6, 3))
       X = (PDFFORMS.FWIDTH - 133)
       IF X < 0 THEN X = 0
       IF X < 5 THEN
          X1 = X: X2 = 0
          ELSE
          X1 = (X \ 2): IF X1 < 4 THEN X1 = 4 ELSE IF X1 > 15 THEN X1 = 15
          X2 = X - X1: IF X2 > 15 THEN X2 = 15
       END IF
       PDFFORMS.IMARGIN = X1: PDFFORMS.OMARGIN = X2
       X = (PDFFORMS.FLENGTH - 65)
       IF X < 0 THEN
          X = 0: Y = 0
          ELSE
          Y = (X \ 2): X = X - Y
          IF X > 3 THEN X = 3: Y = 3
       END IF
       PDFFORMS.TMARGIN = Y: PDFFORMS.BMARGIN = X
       PDFFORMS.OPTIONS = 3
    END IF
    ENV.FORMS = W
    PDFFORMS.REF = PDFFORMS.REF + 1
    PFORMS(W) = PDFFORMS
' Update PSETUPS()
    PSETUPS(ENV.PSETUP) = PDFSetup
    PDFHdr.VER = CHR$(255)
    X = PSUTable.FT: FOUND = 0
    WHILE (X > 0) AND (NOT FOUND)
       FOUND = (MID$(XPSETUP$(X), 1, 8) = PDFSetup.NAME)
       IF NOT FOUND THEN X = PSUCHN(X).FWD
    WEND
    MID$(XPSETUP$(X), 34, 8) = PDFSetup.PRINTER
    CALL WinCLOSE(0, 2)
    GOSUB 130: GOSUB 120
    GOSUB 100: GOSUB 110
    GOTO 300

500 ' Process FILE Requests
    CALL PutMSG("XXXXM006")
    IF ENV.PDF <> SPACE$(8) THEN CONFIG.PDF = ENV.PDF
    IF A = 27 THEN GOTO 200 ELSE R$ = A$
    IF (R$ < "1") OR (R$ > "2") THEN SOUND BP, DUR: GOTO 500
    GOSUB 510: IF A = 27 THEN GOTO 500
    SELECT CASE R$
       CASE "1": GOTO 520
       CASE "2": GOTO 530
    END SELECT

510 ' Get PDF File Prefix
    CALL FmtFIND("PRTCS050")
    CALL FmtFindFLD("PDF")
    LSET UTXT$(FFLD.UTXT) = CONFIG.PDF
    CALL MsgFORMAT: CALL WinSWITCH(2)
    IF A = 27 THEN RETURN
    CALL FmtFindFLD("PDF"): CONFIG.PDF = UTXT$(FFLD.UTXT)
    IF CONFIG.PDF <> ENV.PDF THEN CONFIG.VER = -1
    RETURN

520 ' Save TABLE Changes Into PDF File
    CALL PDFOpen("NEW"): IF NOT OKAY THEN GOTO 500
    CALL PutMSG("PRTCM010")
    PDFHdr.FTYPE = "PDF": PDFHdr.VER = CHR$(0): PDFHdr.EOFRBA = 0
    PUT #1, 1, PDFHdr
    X = PSUTable.FT
    WHILE X > 0
        LSET XNAME$ = XPSETUP$(X)
        Y = PDFHdr.SETUPS: FOUND = 0
        WHILE (Y > 0) AND (NOT FOUND)
           FOUND = (PSETUPS(Y).NAME = XNAME$)
           IF NOT FOUND THEN Y = Y - 1
        WEND
        PUT #1, , PSETUPS(Y)
        X = PSUCHN(X).FWD
    WEND
    X = 0: Y = 0
    WHILE X < PDFHdr.PRINTERS
        X = X + 1
        IF PRINTERS(X).REF > 0 THEN Y = Y + 1: PUT #1, , PRINTERS(X)
    WEND
    PDFHdr.PRINTERS = Y
    X = 0: Y = 0
    WHILE X < PDFHdr.FORMS
        X = X + 1
        IF PFORMS(X).REF > 0 THEN Y = Y + 1: PUT #1, , PFORMS(X)
    WEND
    PDFHdr.FORMS = Y
    X = 0: Y = 0
    WHILE X < PDFHdr.CODES
        X = X + 1
        IF PCODES(X).REF > 0 THEN Y = Y + 1: PUT #1, , PCODES(X)
    WEND
    PDFHdr.CODES = Y
    X = 0: Y = 0
    WHILE X < PDFHdr.SYMBOLS
        X = X + 1
        IF PSYMBOLS(X).REF > 0 THEN Y = Y + 1: PUT #1, , PSYMBOLS(X)
    WEND
    PDFHdr.SYMBOLS = Y
    PDFHdr.EOFRBA = SEEK(1) - 1
    PUT #1, 1, PDFHdr
    CLOSE
    LSET XNAME$ = PDFSetup.NAME
    CALL PDFOpen("INIT")
    CALL PDFGetSet(XNAME$)
    GOSUB 5: GOTO 50

530 ' Load Printer SETUP Data from PDF File
    IF PDFHdr.VER = CHR$(255) THEN
       CALL PutMSG("XXXXM020")
       IF A$ <> SNGLKEY$(1) THEN GOTO 500
    END IF
    CALL PutMSG("PRTCM011")
    CALL PDFOpen("INIT"): IF NOT OKAY THEN GOTO 500
    GOSUB 5: GOTO 50

600 ' Process PRINT Request
    HSETUP = ENV.PSETUP
    CALL PutMSG("PRTCM015")
    IF A = 27 THEN GOTO 200
    IF (A$ < "1") OR (A$ > "2") THEN CALL ErrBEEP(0): GOTO 600
    PRINTALL = (A$ = "2")
    'RDFReport.NAME = "Printer SETUP Listing"
    'HTFMAX = 8: REDIM HTFTXT$(HTFMAX)
    'HTFTXT$(1) = "H01<Family History System>"
    'HTFTXT$(2) = "H01<[RPTNAME]>"
    'HTFTXT$(3) = "H01"
    'HTFTXT$(4) = "H01Printer Definition File: [FILENAME]"
    'HTFTXT$(5) = "TF2   Forms Parameters and Options:"
    'HTFTXT$(6) = "TC2   Printer CODEs"
    'HTFTXT$(7) = "TS2   Chart Drawing SYMBOLs"
    'HTFTXT$(8) = "F02<Page [PAGE]>"
    CALL RptOPEN("PRTC"): IF (NOT OKAY) THEN 605
    CALL PrintOPEN: IF NOT OKAY THEN 605
601 PRT.SECTION = "0"
    REDIM RPTLIT$(24)
    CALL GetRVAR("DEFAULT", X): RPTLIT$(0) = RVAR$(X) + " "
    CALL GetRVAR("SETUP", X): RPTLIT$(1) = RVAR$(X) + " "
    CALL GetRVAR("DESCRIPT", X): RPTLIT$(2) = "  " + RVAR$(X) + " "
    CALL GetRVAR("PRINTER", X): RPTLIT$(3) = "   " + RVAR$(X) + " "
    CALL GetRVAR("DEST", X): RPTLIT$(4) = "  " + RVAR$(X) + " "
    CALL GetRVAR("WIDTH", X): RPTLIT$(5) = RVAR$(X)
    CALL GetRVAR("LENGTH", X): RPTLIT$(6) = RVAR$(X)
    CALL GetRVAR("TOPMGN", X): RPTLIT$(7) = RVAR$(X)
    CALL GetRVAR("BOTMGN", X): RPTLIT$(8) = RVAR$(X)
    CALL GetRVAR("INSMGN", X): RPTLIT$(9) = RVAR$(X)
    CALL GetRVAR("OUTMGN", X): RPTLIT$(10) = RVAR$(X)
    CALL GetRVAR("LINES", X): RPTLIT$(11) = " " + RVAR$(X) + "   "
    CALL GetRVAR("CHARS", X): RPTLIT$(12) = " " + RVAR$(X) + "   "
    CALL GetRVAR("OPTIONS", X): RPTLIT$(13) = RVAR$(X) + "   "
    CALL GetRVAR("OPTSSIDE", X): RPTLIT$(14) = RVAR$(X) + " "
    CALL GetRVAR("OPTFRMFD", X): RPTLIT$(15) = "  " + RVAR$(X) + " "
    CALL GetRVAR("OPTPGWT", X): RPTLIT$(16) = "  " + RVAR$(X) + " "
    CALL GetRVAR("INITIAL", X): RPTLIT$(17) = RVAR$(X)
    CALL GetRVAR("FONT", X): RPTLIT$(18) = RVAR$(X)
    CALL GetRVAR("NEWLINE", X): RPTLIT$(19) = RVAR$(X)
    CALL GetRVAR("FORMFEED", X): RPTLIT$(20) = RVAR$(X)
    CALL GetRVAR("RESET", X): RPTLIT$(21) = RVAR$(X)
    CALL GetRVAR("LINESYMB", X): RPTLIT$(22) = RVAR$(X)
    CALL GetRVAR("REGBOX", X): RPTLIT$(23) = RVAR$(X)
    CALL GetRVAR("HIBOX", X): RPTLIT$(24) = RVAR$(X)
    SYSVAR.FILENAME = CONFIG.PDF + ".PDF"
    LSET PLINE$ = RPTLIT$(0) + PSETUPS(PDFHdr.DEFAULT).NAME + RPTLIT$(2) + PSETUPS(PDFHdr.DEFAULT).DESC
    CALL PrintLINE: CALL WriteLINE
    Y = PSUTable.FT
    WHILE Y > 0
       LSET XNAME$ = XPSETUP$(Y)
       IF PRINTALL OR (PSETUPS(ENV.PSETUP).NAME = XNAME$) THEN GOSUB 610
       Y = PSUCHN(Y).FWD
    WEND
    CALL RptCLOSE: IF PRT.AGAIN THEN 601

605 ' End of Print
    CLOSE
    LSET XNAME$ = PSETUPS(HSETUP).NAME
    CALL PDFGetSet(XNAME$)
    GOTO 50

610 ' Print SETUP Information
    PRT.SECTION = "0"
    CALL PrintLINE
    CALL PDFGetSet(XNAME$)
    PRTSetup = PDFSetup: PRTForms = PDFFORMS
    PRTCodes = PDFCodes: PRTSymbols = PDFSymbols
    LSET XNAME$ = PSETUPS(HSETUP).NAME
    CALL PDFGetSet(XNAME$)
    LSET PLINE$ = RPTLIT$(1) + PRTSetup.NAME + RPTLIT$(2) + PRTSetup.DESC + RPTLIT$(3) + MID$(XPSETUP$(Y), 34, 8) + RPTLIT$(4) + PRTSetup.DEST
    CALL PrintLINE
    PRT.SECTION = "F"
    CALL PrintTitle
    MID$(PLINE$, 8, 255) = RPTLIT$(5) + STR$(PRTForms.FWIDTH) + RPTLIT$(12)
    MID$(PLINE$, 40, 255) = RPTLIT$(6) + STR$(PRTForms.FLENGTH) + RPTLIT$(11)
    CALL PrintLINE
    MID$(PLINE$, 8, 255) = RPTLIT$(7) + STR$(PRTForms.TMARGIN) + RPTLIT$(11)
    MID$(PLINE$, 40, 255) = RPTLIT$(8) + STR$(PRTForms.BMARGIN) + RPTLIT$(11)
    CALL PrintLINE
    MID$(PLINE$, 8, 255) = RPTLIT$(9) + STR$(PRTForms.IMARGIN) + RPTLIT$(12)
    MID$(PLINE$, 40, 255) = RPTLIT$(10) + STR$(PRTForms.OMARGIN) + RPTLIT$(12)
    CALL PrintLINE
    IF (PRTForms.OPTIONS AND 1) THEN S$ = SNGLKEY$(1) ELSE S$ = SNGLKEY$(2)
    IF (PRTForms.OPTIONS AND 2) THEN F$ = SNGLKEY$(1) ELSE F$ = SNGLKEY$(2)
    IF (PRTForms.OPTIONS AND 8) THEN P$ = SNGLKEY$(1) ELSE P$ = SNGLKEY$(2)
    MID$(PLINE$, 8, 255) = RPTLIT$(13) + RPTLIT$(14) + S$ + RPTLIT$(15) + F$ + RPTLIT$(16) + P$
    CALL PrintLINE
    PRT.SECTION = "C"
    CALL PrintTitle
    MID$(PLINE$, 8, 16) = RPTLIT$(17)
    LSET CODE$ = PRTCodes.INIT
    GOSUB 611
    MID$(PLINE$, 8, 16) = RPTLIT$(18)
    LSET CODE$ = PRTCodes.FONT1
    GOSUB 611
    MID$(PLINE$, 8, 16) = RPTLIT$(19)
    LSET CODE$ = PRTCodes.LINEFEED
    GOSUB 611
    MID$(PLINE$, 8, 16) = RPTLIT$(20)
    LSET CODE$ = PRTCodes.FormFeed
    GOSUB 611
    MID$(PLINE$, 8, 12) = RPTLIT$(21)
    LSET CODE$ = PRTCodes.RESET
    GOSUB 611
    PRT.SECTION = "S"
    CALL PrintTitle
    MID$(PLINE$, 8, 12) = RPTLIT$(22)
    LSET CODE$ = CHR$(11) + PRTSymbols.SET1
    CALL Code2Dec(CODE$, XCODE$)
    MID$(PLINE$, 25, 255) = XCODE$
    CALL PrintLINE
    MID$(PLINE$, 8, 11) = RPTLIT$(23)
    LSET CODE$ = CHR$(10) + PRTSymbols.SET2
    CALL Code2Dec(CODE$, XCODE$)
    MID$(PLINE$, 25, 255) = XCODE$
    CALL PrintLINE
    MID$(PLINE$, 8, 11) = RPTLIT$(24)
    LSET CODE$ = CHR$(10) + PRTSymbols.SET3
    CALL Code2Dec(CODE$, XCODE$)
    MID$(PLINE$, 25, 255) = XCODE$
    CALL PrintLINE
    CALL WriteLINE: CALL WriteLINE
    RETURN

611 ' Print CODE Line
    SELECT CASE CTYPE
       CASE 1: CALL Code2Hex(CODE$, XCODE$)
       CASE 2: CALL Code2Dec(CODE$, XCODE$)
    END SELECT
    MID$(PLINE$, 25, 255) = XCODE$
    CALL PrintLINE
    RETURN

1000 ' Return to FHSMENU.BAS
     IF (PDFHdr.VER = CHR$(255)) THEN
        CALL PutMSG("XXXXM020")
        IF A$ <> SNGLKEY$(1) THEN GOTO 200
     END IF
     CALL PutMSG("")
     CLOSE : A = 0
     CALL PgmPREP(PGMRTRN$)
     IF (NOT OKAY) OR (A = 27) THEN GOTO 200
     ENV.WINHILVL = 1
     REDIM PSETUP$(PDFHdr.SETUPS)
     X = PSUTable.FT: Y = 0
     WHILE X > 0
        Y = Y + 1
        PSETUP$(Y) = LEFT$(XPSETUP$(X), 33)
        X = PSUCHN(X).FWD
     WEND
     REDIM PSUCHN(PDFHdr.SETUPS) AS ChainPTRS
     CALL ChnINIT(PSUCHN(), 1, PDFHdr.SETUPS)
     PSUMAX = PDFHdr.SETUPS: PSUTable.SIZE = PSUMAX
     PSUTable.FF = 0: PSUTable.LT = PDFHdr.SETUPS
     PSUTable.ML = 33: PSUTable.XL = 33
     IF PGMRTRN$ <> "FHSRPTS" THEN CALL RptFREE("FILE")
     IF PGMRTRN$ <> "FHSTABL" THEN CALL TableFREE("PRINTERS")
     CHAIN PGMRTRN$

REM $DYNAMIC
SUB Code2Dec (ICODE$, OCODE$)
    N = ASC(ICODE$): X = 2: Y = 1: LSET OCODE$ = " "
    WHILE N > 0
       IF Y > 1 THEN MID$(OCODE$, Y, 1) = ",": Y = Y + 1
       MID$(OCODE$, Y, 3) = LTRIM$(STR$(ASC(MID$(ICODE$, X, 1))))
       Y = INSTR(Y, OCODE$, " ")
       X = X + 1: N = N - 1
    WEND
END SUB

SUB Code2Hex (ICODE$, OCODE$)
   N = ASC(ICODE$): X = 2: Y = 1: LSET OCODE$ = " "
   WHILE N > 0
      H1 = ASC(MID$(ICODE$, X, 1)): H2 = (H1 AND 15): H1 = (H1 \ 16)
      MID$(OCODE$, Y, 1) = MID$("0123456789ABCDEF", H1 + 1, 1)
      MID$(OCODE$, Y + 1, 1) = MID$("0123456789ABCDEF", H2 + 1, 1)
      X = X + 1: Y = Y + 3
      N = N - 1
   WEND
END SUB

SUB Dec2Code (ICODE$, OCODE$)
    LSET OCODE$ = ""
    X = 1: Z = 1
    L = LEN(ICODE$) + 1
    WHILE X < L
       WHILE MID$(ICODE$, X, 1) = " ": X = X + 1: WEND
       WHILE MID$(ICODE$, X, 1) = ",": X = X + 1: WEND
       Y1 = INSTR(X, ICODE$, ","): Y2 = INSTR(X, ICODE$, " ")
       IF Y1 = 0 THEN Y1 = Y2 ELSE IF Y2 = 0 THEN Y2 = Y1
       IF Y1 < Y2 THEN Y = Y1 ELSE Y = Y2
       IF Y = 0 THEN Y = L
       IF Y > X THEN
          N# = VAL(MID$(ICODE$, X, Y - X)): IF N# > 255 THEN N# = 0
          N = N#: Z = Z + 1: MID$(OCODE$, Z, 1) = CHR$(N)
       END IF
       X = Y + 1
    WEND
    MID$(OCODE$, 1, 1) = CHR$(Z - 1)
END SUB

SUB Hex2Code (ICODE$, OCODE$)
    LSET OCODE$ = " ": L = LEN(ICODE$) + 1
    X = 1: Z = 1
    WHILE X < L
       WHILE MID$(ICODE$, X, 1) = " ": X = X + 1: WEND
       WHILE MID$(ICODE$, X, 1) = ",": X = X + 1: WEND
       Y1 = INSTR(X, ICODE$, ","): Y2 = INSTR(X, ICODE$, " ")
       IF Y1 = 0 THEN Y1 = Y2 ELSE IF Y2 = 0 THEN Y2 = Y1
       IF Y1 < Y2 THEN Y = Y1 ELSE Y = Y2
       IF Y = 0 THEN Y = L
       IF (Y - X) = 2 THEN
          X$ = UCASE$(MID$(ICODE$, X, 1))
          H1 = INSTR("0123456789ABCDEF", X$): IF H1 > 0 THEN H1 = H1 - 1
          X$ = UCASE$(MID$(ICODE$, X + 1, 1))
          H2 = INSTR("0123456789ABCDEF", X$): IF H2 > 0 THEN H2 = H2 - 1
          Z = Z + 1: MID$(OCODE$, Z, 1) = CHR$(16 * H1 + H2)
          ELSE
          IF (Y - X) = 1 THEN
             N = ASC(MID$(ICODE$, X, 1))
             ELSE
             IF Y > X THEN
                N# = VAL(MID$(ICODE$, X, Y - X))
                IF N# > 255 THEN N = 0 ELSE N = N#
             END IF
          END IF
          IF N > 0 THEN Z = Z + 1: MID$(OCODE$, Z, 1) = CHR$(N): N = 0
       END IF
       X = Y + 1
    WEND
    MID$(OCODE$, 1, 1) = CHR$(Z - 1)
END SUB

