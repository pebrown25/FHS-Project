Attribute VB_Name = "defCommon"
 DefInt A-Z
 
 Type ConfigPARMS ' Configuration Parameters   - 69+ALx Bytes
      CFG     As String * 8 ' Prefix for Name of Configuration File
      VER     As Integer    ' Version Number = 1 (03/18/96)
      EDA     As String * 1 ' Edit Assist - "Y" or "N"
      DFM     As String * 1 ' Date Format for Data Entry
      FDF     As String * 8 ' Prefix for Name of File Definition File
      PDF     As String * 8 ' Prefix for Name of Printer Definition File
      TDF     As String * 3 ' Language Prefix for Table Definition File
      MDF     As String * 3 ' Language Prefix for Message Definition File
      SDF     As String * 3 ' Language Prefix for Screen Definition File
      RDF     As String * 3 ' Language Prefix for Report Definition File
      SCOLORS As String * 12 ' Screen Colors
      BP      As Integer     ' Error Tone Frequency
      DUR     As Integer     ' Error Tone Duration
      VCMD    As String * 8  ' Name of Utility used to view Report Files
      AL1L    As String * 1  ' Length of AL1
      AL2L    As String * 1  ' Length of AL2
      AL3L    As String * 1  ' Length of AL3
      AL4L    As String * 1  ' Length of AL4
      AL5L    As String * 1  ' Length of AL5
      'AL1     AS STRING * 30 ' Name and Address
      'AL2     AS STRING * 30
      'AL3     AS STRING * 30
      'AL4     AS STRING * 30
      'AL5     AS STRING * 30
 End Type

 Type SystemVARS  ' System Variables  - 264 Bytes
      PGMNAME  As String * 8
      TABNAME  As String * 8
      MSGNAME  As String * 8
      FILENAME As String * 100
      ERR      As String * 5
      ERL      As String * 5
      SETUP    As String * 8
      MVAR     As String * 50    ' Miscellaneous Message Variable
      NVAR1    As Long
      NVAR2    As Long
      NVAR3    As Long
      NVAR4    As Long
      NVAR5    As Long
      NVAR6    As Long
      CHARTID  As String * 35
      SUBJECT  As String * 52
      DATE     As String * 11
 End Type

 Type TablePARMS  ' TABLE Parameters     - 60 Bytes
      TABNAME As String * 8  ' Name of Table
      OPEN    As Integer     ' =-1 If Table is OPEN
      SIZE    As Integer     ' Number of Entries in Table
      XL      As Integer     ' Max Length of Table Data Entry in File
      KO      As Integer     ' Offset of Key Field in Table Entry
      KL      As Integer     ' Length of Key Field
      ' Fields Describing Table Format
      FMTNAME As String * 8  ' Format for Table Entry Display
      FMTNUM  As Integer     ' Seq Number of Table Format in FMTTAB()
      KEYNAME As String * 4  ' Name of Key Field in FLDTAB()
      KEYFLD  As Integer     ' Seq Number of Key Field in FLDTAB()
      KCOLOR  As Integer     ' Color of Key Field
      KC      As Integer     ' Column of Key Field on Screen
      ' Fields Describing Table Display
      TL      As Integer     ' Line Number of Top Line
      BL      As Integer     ' Line Number of Bottom Line
      CL      As Integer     ' Line Number of Current Entry
      LL      As Integer     ' Last Line of Data on Current Screen
      ' Fields Describing Entries in Table Data array
      ML      As Integer     ' Max Length of Table Data Entry in TABDATA$()
      FT      As Integer     ' First Table Entry
      LT      As Integer     ' Last Table Entry
      CT      As Integer     ' Current Table Entry
      FS      As Integer     ' First Entry on Current Screen Display
      LS      As Integer     ' Last Entry on Current Screen
      FF      As Integer     ' First Free Entry in Table
 End Type

 Type GroupPARMS             ' 22 Bytes - MDFIndex(), SDFIndex()
      NAME    As String * 4
      CHG     As Integer     ' Change Level of Group (IF Changed then >127)
      OPEN    As Integer     ' = -1 if Group is OPEN, -2 if changed?
      FDIR    As Integer     ' RBA of File Dir Entry for Group
      BGNRBA  As Long        ' RBA of First Group Entry in File
      SIZE    As Integer     ' Number of Entries in Group
      FIRST   As Integer     ' First Group Entry in Table
      LAST    As Integer     ' Last Group Entry in Table
      DESC    As Integer     ' Descriptive Entry in GROUPS or REPORTS Table
 End Type

 Type ReportINDEX            ' 20 Bytes - RDFIndex()
      CODE    As String * 4
      CHG     As Integer     ' If <0 then Change has been made
      OPEN    As Integer     ' =-1 IF Report is OPEN, -2 IF Changed
      DEFIX   As Integer     ' Report Entry in RDFReports()
      DEFRBA  As Integer     ' RBA of RDF File Report Definition Record
      HTFBGN  As Integer     ' First HTF Entry in HTFTXT$()
      OPTBGN  As Integer     ' First OPT Entry in ROPT$()
      VARBGN  As Integer     ' First Variable Entry in RVARS()
      FLDBGN  As Integer     ' First Report Field Entry in RFLDS()
 End Type

 Type ChainPTRS  ' Pointers Connecting Chained Table Entries
      FWD     As Integer
      BWD     As Integer
 End Type

 Type ENVParms   ' ENV, ENVTAB()    - 168 Bytes
      SYSMODE  As Integer    ' 0=Basic Programs, 1=Registered Version
      DSNCNT   As Integer    ' Number of DATASET Name Entries in FT$()
      PGMRTRN  As String * 8 ' Name of Program to Return to
      PGMNAME  As String * 8 ' Name of Processing Program
      PGMSFX   As String * 4 ' Suffix of Program = ".EXE" or ".BAS"
      PGRP     As Integer    ' Number of Program Group in GROUP Table
      EDMAST   As Integer    ' Extended EDIT Options if = -1
      EDINSRT  As Integer    ' Indicate INSERT Mode in Win Update
      SortOPTS As Integer    ' =NDXHdr.OPT\64
      SortSTAT As Integer    ' Gather,Display Sort Statistics
      HUSBSURNM As Integer   ' Use Husband's Surname for Wife's
      SHOWMSGID As Integer   ' Indicate Mode for Showing MSGID
      LARGEMEM As Integer    ' Load ALL Messages and Formats in memory
      RELRULES As Integer    ' Relationship Rules
                     '   0:COMMON,MIXED 1:CIVIL 2:LATIN 4:PATERNAL 8:MATERNAL
      FFOPEN   As Integer    ' Indicate FAMILY FILE is Open
      LPNSupport As Integer    ' Long Place Name Support
      FDF      As String * 8 ' Prefix for Name of File Definition File
      FSETUP   As Integer    ' Current FDF Setup Entry
      PDF      As String * 8 ' Prefix for Name of Printer Definition File
      PSETUP   As Integer    ' Current PDF Setup Entry
      RPTIX    As Integer    ' Current Report RDFIndex() Entry
      OPTTAB   As Integer    ' Table() Entry for RPTOPTS System Table
      ScrnCLR  As Integer    ' Indicator that Screen needs to be reformatted
      TDF      As String * 3 ' Active TDF Prefix
      MDF      As String * 3 ' Active MDF Prefix
      SDF      As String * 3 ' Active SDF Prefix
      RDF      As String * 3 ' Active RDF Prefix
      RTITLE   As String * 72 ' User Title for Report
 End Type

 Type RptOPTIONS  ' RPTOPT, RPTOPTS() - 50 BYTES
      INITCHK  As Integer   ' =255 IF Option HAS BEEN Initialised
      FREEFORM As Integer
      TIMELINE As Integer
      OUTLINE  As Integer
      TWOLINES As Integer
      PRGL     As Integer
      PRID     As Integer   ' 1=PRID,2=PRID in Heading Only
      Relation As Integer
      NAMEFMT  As Integer   ' 1=SURNAME FIRST, 2=SURNAME UPPER CASE
      PRSEX    As Integer
      PRAGE    As Integer   ' 1=PRAGE,2=PRDAGE
      DateS    As Integer
      PLACES   As Integer
      SPOUSE   As Integer
      COMMENTS As Integer
      IGNBKT   As Integer
      RMVBKT   As Integer
      BLNKSEP  As Integer
      BLDREF   As Integer
      SNGLREF  As Integer      ' Only Save "Most Significant" Reference
      LINEREF  As Integer
      EVENPAGE As Integer
      PRTHEAD  As Integer
      PRTFOOT  As Integer
      COPYRITE As Integer
 End Type

 Type PagePARMS   ' PAGE - 84 Bytes
      ' Initialised from PDFForms
      BLTOP    As Integer      ' Number Blank Lines at TOP of Page
      BLBOT    As Integer      ' Number Blank Lines BETWEEN Foot and Bottom
      IMARGIN  As Integer      ' Inside Margin
      OMARGIN  As Integer      ' Outside Margin
      ' Initialised from RDFReport
      BLHEAD   As Integer      ' Number Blank Lines BETWEEN Head and Title
      BLTITLE  As Integer      ' Number Blank Lines BETWEEN Title and Body
      BLBODY   As Integer      ' Number Blank Lines BETWEEN Body Lines
      BLFOOT   As Integer      ' Number Blank Lines BETWEEN Body and Foot
      ' Initialised from Report Options
      CINDENT  As Integer      ' Indent for Continued Comment Lines
      ' Following Fields are Dynamic
      BGNPAGE  As Integer      ' Number of First Page of Report
      MARGIN   As Integer      ' Left Margin for Current Page
      RMARGIN  As Integer      ' Right Margin for Current Page
      WIDTH    As Integer      ' Page Width (including Margins)
      FULLWIDTH As Integer     ' Print Report Full Width instead of Strips
      BWIDTH   As Integer      ' Body Width (Text Characters between margins)
      BLENGTH  As Integer      ' Number Lines between Heading and Footing
      LWIDTH   As Integer      ' Line Width for Free Form Lines
      LINE     As Long         ' Number of Last Line Printed on Page
      TotLINES As Long         ' Total Number of Lines Printed
      COUNT    As Integer      ' Number of Current Page
      XREF     As Integer      ' Number of Waiting Entries in RREFXID()
      XID      As Integer      ' RID to Add to RREFXID() in next PutDATA
      xType    As Integer      ' Type Entry to Add to RREFTYPE()
      RLINE    As Double       ' REF LINE for this DATA Block
      RPAGE    As Integer      ' REF PAGE for this DATA Block
      LOFST    As Integer      ' Offset to Next Char on Data Line
      COFST    As Integer      ' Offset for Continued Free Form Text
      DFCNT    As Integer      ' Number of Entries in Report Field Def TABLE
      PFCNT    As Integer      ' Number of Prefix Fields in Report Field TABLE
      CSID     As Integer      ' Current Spouse ID
      WAIT     As Integer      ' Indicator for Waiting at Top of Next Page
      TOP      As Integer      ' Indicator for Printing Top Margin
      FOOT     As Integer      ' Indicator for Printing End of Page
      BREAK    As Integer      ' Indicator for Page Break in progress
      HLCNT    As Integer
      TLCNT    As Integer
      DLCNT    As Integer
      FLCNT    As Integer
 End Type

 Type PrtPARMS    ' PRT, PRTOPTS() - 40 Bytes
      OPT      As Integer      ' 1=SCREEN, 2=PRINTER, 3=FILE
      DEST     As String * 14  ' Built During Print OPEN
      SECTION  As String * 1   ' Section of Report being printed
      OPEN     As String * 1   ' IF =255 THEN Print File is OPEN
      AGAIN    As Integer
      ALL      As Integer
      ' Dynamic Parms
      BGNPAGE  As Integer      ' Page Number at which to START Printing
      EOF      As Integer      ' Indicator for End of Report
      FOFFSET  As Integer      ' Offset of First FORM Print Position on PLINE
      CFORMS   As Integer      ' Indicator for Continuous Forms (no Page Brk)
      ' From PDFForms Record
      FWIDTH   As Integer
      FLENGTH  As Integer
      SNGLSIDE As Integer
      FORMFEED As Integer
      PAGEWAIT As Integer      ' Default for Waiting at Top of Next Page
 End Type

 Type DateFields
      CC    As Integer
      yy    As Integer
      yyyy  As Integer
      mm    As Integer
      ww    As Integer
      DD    As Integer
      ddd   As Integer
      adat  As Double
      ly    As Integer
      wday  As Integer
 End Type

 Rem $INCLUDE: 'FAMFILES.INC'
 Rem $INCLUDE: 'WRKFILES.INC'
 Rem $INCLUDE: 'LNGFILES.INC'

 Rem $DYNAMIC

 Public MO, WB$, BP, DUR, CRLF$, FG(), BG(), PrintCntl$(), FT$(), FType$(), FGrp$()
 Public CONFIG As ConfigPARMS, SysVAR As SystemVARS, CFGAL$()
 Public ENV As ENVParms, CurENV, ENVTAB() As ENVParms, ENVMAX
 Public FDFHDr As FileDefHDR, FDFSURBA()
 Public FDFDSRec As FileDatasetREC, FDFSetup As FileSetupREC
 Public FSUTABLE As TablePARMS, FSETUP$(), FSUCHN() As ChainPTRS, FSUMAX
 Public FSETUPS() As FileSetupREC, FFNAME$()
 Public FAMFPFX$, WRKFPFX$, TRANPFX$, RPTFPFX$
 Public PDFHdr As PrintDefHDR, PDFSetup As PrintSetupREC
 Public PSUTABLE As TablePARMS, PSUCHN() As ChainPTRS, PSUMAX
 Public PSETUPS() As PrintSetupREC
 Public PRINTERS$(), PFONT$()
 Public TDFHdr As TableFileHDR, TDFDir As TableFileDIR
 Public TDFIndex() As TableFileDIR, TABMAX
 Public MDFHdr As MsgFileHDR, MDFGrpDIR As MsgGroupDIR
 Public MDFIndex() As GroupPARMS, GRPMAX
 Public GrpTABLE As TablePARMS
 Public TABLE() As TablePARMS, TABDATAMAX
 Public TABDATA$(), TABDATACHN() As ChainPTRS, TABDATAFree
 Public MSGTABLE As TablePARMS, MSGMAX
 Public MSGDATA$(), MSGDATACHN() As ChainPTRS, MSGDATAFree
 Public RDFHdr As ReportFileHDR
 Public RDFDir As ReportFileDIR, RDFReport As ReportDefREC
 Public RDFHTFLine As ReportHTFLine
 Public RDFVar As ReportVariableREC, RDFFld As ReportFieldREC
 Public RDFIndex() As ReportINDEX, RPTMAX, RDFMAX, RDFFree
 Public RDFReports() As ReportDefREC, RDFChn() As ChainPTRS
 Public RDFTable As ReportTableHDR
 Public RMONTB$(), RGENDR$(), RRELCD$()
 Public HTFTXT$(), HTFCHN() As ChainPTRS, HTFMAX, HTFFree
 Public HTFLINE$(), HTFLMAX
 Public OPTTable As TablePARMS, RPTOPTS() As RptOPTIONS
 Public RPTOPT$(), OPTCHN() As ChainPTRS, OPTMAX, OPTFREE
 Public RFLDS() As ReportFieldREC, RFLDCHN() As ChainPTRS
 Public RFLD$(), RFMAX, RFFree
 Public RVARS() As ReportVariableREC, RVARCHN() As ChainPTRS
 Public RVAR$(), RVMAX, RVFree
 Public RptOPTION As RptOPTIONS
 Public PRT As PrtPARMS, PAGE As PagePARMS
 Public a, CA$, AX$, C$, BCOL, ECOL, FATR, FLIN, FGND, BGND
 Public FUN, OKAY, FOUND, XCHANGED, FFCHANGED, EMCNT, EMSG$
 Public RELCODE$(), RELRULE$(), FGENDR$(), ERRMSG$()
 Public LF$, PLINE$, CL$, SURNM$, AFMT
 Public SNDXTB$, UCT1$, UCTRAN$(), SNGLKEY$(), SKEG()
 Public RFLBL$(), RFLTH()
 Public FF1Hdr As NameFileHDR
 Public FF1Rec As NameRecord, NameREC As NameRecord
 Public FF2Hdr As AddrFileHDR, FF2PFX As AddrPNTRS
 Public FF2SREC As AddrRECORD1, FF2LREC As AddrRECORD2
 Public FF3HDR As OthrFileHDR
 Public FF3PFX As OthrFileRecPFX, FF3PNTRS As OthrFileRecPNTRS
 Public FF3PLACE As OthrPlaceREC, FF3PLACE2 As PlaceNameREC
 Public BPLACE As String * 41, DPLACE As String * 41
 Public OBPLACE As String * 41, ODPLACE As String * 41
 Public MPLACE As String * 41, DVPLACE As String * 41
 Public OMPLACE As String * 41, ODVPLACE As String * 41
 Public FF3SPOUSE As OthrFamilyREC
 Public FF3COMMENT As OthrCommentREC, FF3EVENT As OthrEventREC
 Public FF3EDUC As OthrEducDATA, FF3WORK As OthrWorkDATA
 Public FF3MIL As OthrMilitaryDATA, FF3HEALTH As OthrHealthDATA
 Public SFCNT, SX(), SFLD(12), NDXID(), NDXRECS, NDXHDR As NDXFileHDR
 Public WFHDR As WorkFileHDR, WFBASE
 Public WFANCREC As WorkFileAncestorREC
 Public WFAGLREC As WorkFileGLREC, WFGLREC As WorkFileGLREC
 Public WFRELREC As WorkFileRelativeREC, GBL(), GLNUMCH(), PARPTR&()
 Public RREFHDR As RefFileHDR, RREFRPT As RefFileRptREC
 Public RREFREC As RefFileREC, RREFXData$
 Public RX&(), RXIMP(), RXSYM$(), RREFXID(), RREFTYPE()
 Public SLCTHDR As SlctFileHDR
 Public SortPHASE, SortFMT, SortOK, SDL, SS, SN1, SN2, SortSTAT!()
 Public MRGID(), MRGBX(), MRGIX(), MRGSS(), MRGSD$()
 Public ERID()
 
' New Variables for Windows Version
Public FamSetupOKAY As Integer
Public FamFileOKAY As Integer, FamUpdate As Integer
Public FormLoaded(9) As Integer
Public Const RelWForm = 1, RPgmForm = 2, FileForm = 4
Public Const ViewForm = 8, SlctForm = 16, NdWkForm = 32
Public Const UPgmForm = 64, PrntForm = 128
Public Const ListForm = 1, MSGForm = 1, MSGDForm = 2
Public Const FamDefForm = 4, RelCalcForm = 8, DateCalcForm = 16
Public Const UpdtForm = 32, SpouseForm = 64, NotesForm = 128
Public Const ParentForm = 1, ChildForm = 2, AddressForm = 4, EventForm = 8
Public Const MrecForm = 16, IndexForm = 32, TTFLForm = 64
Public RelFileOKAY As Integer, WRelType As Integer
Public RelWType As Integer, RelWDSN As String, RelwBASE As Integer
Public SelFileOKAY As Integer, SlecType As Integer, BRID As Integer
Public RptPGM As String, RptWType As Integer, RptCODE As String
Public UtilPGM As String, HLIT$()
Public TLBY, TLEY, ALIN#
Public REVENT$(), IDataNM$(), IDataBGN(), IDataCNT As Integer
Public sTYPE$, SBASE, SCNT, PSCOUNT, SSCOUNT, SelSET() As Integer
Public LOID, HIID, LRELATE(), LRELStyle, LRELOkay, RelChanged
Public NodeHDR As NodeFileHDR, NodeREC As NodeFileREC, NodeGLREC As NodeFileGLREC
Public NWrkOkay, NWrkType
Public StdWidth As String, StdHeight As String
' Fields for Date Calculations - moved from FHSWDCLC on 07/15/99
Public CalType, DFM
Public XDATE As String, DateW As DateFields
Public days_table(13), day_table() As String * 3
Public mon_table(12) As String * 3
Public minADAT, maxADAT As Double, WkLth
Public DateMask As String, DateMasks(9) As String, BMO
Public PGMPath As String

