Attribute VB_Name = "modFHSWRegr"
DefInt A-Z

Type NDATA      ' 174 Bytes
     RID      As String * 5
     PID      As String * 5
     NAME     As String * 53
     BDATE    As String * 12
     BPLACE   As String * 41
     DDATE    As String * 12
     DPLACE   As String * 41
     SEX      As String * 1
     NUMCH    As String * 3
End Type

Type PDATA       ' 79 Bytes - Father, Mother
    NAME      As String * 52
    BDATE     As String * 12
    DDATE     As String * 12
    AGE       As String * 3
End Type

Type MDATA       ' 164 Bytes
     MDATE    As String * 11
     MPLACE   As String * 41
     DVDATE   As String * 11
     DVPLACE  As String * 41
     ANNIV    As String * 3
     SPID     As String * 5
     SPNAME   As String * 52
End Type

Type EDATA       ' 79 Bytes
     CODE     As String * 8
     DESC     As String * 15
     DATE     As String * 11
     DSTAT    As String * 1
     PLACE    As String * 41
     IMPORT   As String * 3
End Type
Dim SPRESEQ, SC, SPMCNT, MAXMarry, MLIST()
Dim IA$, MA$, RRPT$, RSECHDR$, DSECHD1$, BgnSection
Dim NDXID(), RELATE(), RELPTR&()
Dim sTYPE, LINTAB$()
Dim ShowREL, ShowChNUM, ShowRegNUM, ShowPNAME, ExtAdoptRule, RomanCAPS
Dim LinCNT, NMOFST, RID, z


Sub RptREGR(OPT$)
1  Rem $INCLUDE: 'FHSCOMON.BAS'
   PN$ = " FAMILY HISTORY SYSTEM - FHSWREGR - Family Register Report Program"
   CY$ = " (C) Copyright 1997, 1998 by Phillip E. Brown"
   
   Select Case OPT$
      Case "INIT": GoTo 40
      Case "PRINT": GoTo 1000
   End Select
   
40 ' Format Display
   Call PgmPREP("FHSREGR")
   Load frmFHSWRPgm
   Call RptLOAD("REGR")
   frmFHSWRPgm.RPSetup.AddItem "REGR"
   frmFHSWRPgm.RPSetup_Click
   frmFHSWRPgm.Show
   Exit Sub

1000 ' Print Report
   MAXMarry = 10: ReDim MLIST(MAXMarry)
   NAME$ = Space$(52): IA$ = "": MA$ = "": READ$ = "READQ"
   ReDim RELATE(1), RELPTR&(1), LINTAB$(99)
   ReDim SelSET(1), NDXID(1)
   BRID = 0: ShowPNAME = 0
   Call RWRKOpen("READQ", WRelType, 0): Close
   If Not OKAY Then
       Call PutMSG("REGRM002")
       Else
       Call FREGReport
   End If
   Call ClearMSG
End Sub
Rem $DYNAMIC
Sub BldMLIST()
   SPMCNT = 0: SC = FF1Rec.SPOUSE
   If SPRESEQ Then
      While SC <> 0
         SPMCNT = SPMCNT + 1
         If SPMCNT > MAXMarry Then
            MAXMarry = MAXMarry + 5
            ReDim Preserve MLIST(MAXMarry)
         End If
         MLIST(SPMCNT) = SC
         Call FF3GetRec(SC)
         If FF1Rec.RID = FF3SPOUSE.SP1ID Then
            SC = FF3SPOUSE.SP1NXT
            Else
            If FF1Rec.RID = FF3SPOUSE.SP2ID Then
               SC = FF3SPOUSE.SP2NXT
               Else
               SC = 0
            End If
         End If
      Wend
      SC = MLIST(SPMCNT)
   End If

End Sub

Sub BLDREGRDef()
      PAGE.DFCNT = 67
      ReDim RFLTH(PAGE.DFCNT), RFLBL$(PAGE.DFCNT)

58500 ' Select Report Fields - Based Upon Report Options
      IA$ = Space$(12): MA$ = Space$(24): X = 0
      XI$ = "PRSP": GoSub 58505
      XI$ = "PRCH": GoSub 58505
      XI$ = "PRAD": GoSub 58505
      XI$ = "PRCM": GoSub 58505
      XI$ = "PRME": GoSub 58505
      XI$ = "PRED": GoSub 58505
      XI$ = "PRWK": GoSub 58505
      XI$ = "PRML": GoSub 58505
      XI$ = "PRXA": GoSub 58505
      XI$ = "PRXC": GoSub 58505
      XI$ = "PREV": GoSub 58505
      XI$ = "CHCM": GoSub 58505
      GoTo 58510

58505 ' Process Information Option
      X = X + 1: Call RptOptGET(XI$, FX)
      If FOUND Then Mid$(IA$, X, 1) = LTrim$(Str$(FX))
      Return

58510 ' Create Report Definition Tables
      X = 1: RFNAME$ = "06010": GoSub 58515       'Comments
      X = 2: RFNAME$ = "01010": GoSub 58515       'RID
      X = 3: RFNAME$ = "01013": GoSub 58515       'Parent ID
      X = 4: RFNAME$ = "02010": GoSub 58515       'Spouse's ID
      X = 5: RFNAME$ = "01020": GoSub 58515       'Fullname
      X = 6: RFNAME$ = "01023": GoSub 58515       'Father's Name
      X = 7: RFNAME$ = "01024": GoSub 58515       'Mother's Name
      X = 8: RFNAME$ = "01030": GoSub 58515       'Birth Date
      X = 9: RFNAME$ = "01032": GoSub 58515       'Birth Place
      X = 10: RFNAME$ = "01040": GoSub 58515      'Death Date
      X = 11: RFNAME$ = "01042": GoSub 58515      'Death Place
      X = 12: RFNAME$ = "01016": GoSub 58515      'Age
      X = 13: RFNAME$ = "02011": GoSub 58515      'Spouse's Name
      X = 14: RFNAME$ = "02020": GoSub 58515      'Marriage Date
      X = 15: RFNAME$ = "02021": GoSub 58515      'Marriage Place
      X = 16: RFNAME$ = "02030": GoSub 58515      'Divorce Date
      X = 17: RFNAME$ = "02031": GoSub 58515      'Divorce Place
      X = 18: RFNAME$ = "02015": GoSub 58515      'Anniversary (for...years)
      X = 19: RFLBL$(19) = MID$(RFLD$(FX), ALTH + LLTH + 1, HLTH)
      X = 20: RFNAME$ = "01015": GoSub 58515      'Child's Gender
      X = 21: RFNAME$ = "01026": GoSub 58515      'Child's Name
      X = 22: RFNAME$ = "01034": GoSub 58515      'Birth Date/Place
      X = 23: RFNAME$ = "01027": GoSub 58515      'Number Children
              If FOUND Then
                 RFLBL$(67) = MID$(RFLD$(FX), ALTH + 1, LLTH)
                 RFLTH(67) = LLTH
              End If
      X = 24: RFNAME$ = "01044": GoSub 58515      'Death Date/Place
      X = 25: RFNAME$ = "01028": GoSub 58515      'Married Date/Spouse
      X = 26: RFNAME$ = "01030": GoSub 58515      'Birth Date Abbrev.
      X = 27: RFNAME$ = "01032": GoSub 58515      'Birth Place Abbrev.
      X = 28: RFNAME$ = "01040": GoSub 58515      'Death Date Abbrev.
      X = 29: RFNAME$ = "01042": GoSub 58515      'Death Place Abbrev.
      X = 30: RFNAME$ = "02020": GoSub 58515      'Marriage Date Abbrev.
      X = 31: RFNAME$ = "03001": GoSub 58515      'Address Section Label
      X = 32: RFNAME$ = "03010": GoSub 58515      'Address Begin Date
      X = 33: RFNAME$ = "03020": GoSub 58515      'Address End Date
      X = 34: RFNAME$ = "03030": GoSub 58515      'Phone
      X = 35: RFNAME$ = "03031": GoSub 58515      'Address Lines
      X = 36: RFNAME$ = "03032": GoSub 58515      'City
      X = 37: RFNAME$ = "03033": GoSub 58515      'State
      X = 38: RFNAME$ = "03034": GoSub 58515      'Zip Code
      X = 39: RFNAME$ = "03035": GoSub 58515      'Province
      X = 40: RFNAME$ = "03037": GoSub 58515      'Postal Code
      X = 41: RFNAME$ = "03036": GoSub 58515      'Country
      X = 42: RFNAME$ = "04001": GoSub 58515      'Event Information
      X = 43: RFNAME$ = "04010": GoSub 58515      'Event Code
      X = 44: RFNAME$ = "04011": GoSub 58515      'Event Date
      X = 45: RFNAME$ = "04012": GoSub 58515      'Event Place
      X = 46: RFNAME$ = "04013": GoSub 58515      'Event Import Level
      X = 47: RFNAME$ = "05010": GoSub 58515      'Period Begin Date
      X = 48: RFNAME$ = "05011": GoSub 58515      'Period End Date
      X = 49: RFNAME$ = "05030": GoSub 58515      'Education Information
      X = 50: RFNAME$ = "05031": GoSub 58515      'Level
      X = 66: RFLBL$(66) = MID$(RFLD$(FX), ALTH + LLTH + 1, HLTH)
              RFLTH(66) = HLTH
      X = 51: RFNAME$ = "05032": GoSub 58515      'Subject1
      X = 52: RFNAME$ = "05033": GoSub 58515      'Subject2
      X = 53: RFNAME$ = "05034": GoSub 58515      'Degree
      X = 54: RFNAME$ = "05040": GoSub 58515      'Medical Information
      X = 55: RFNAME$ = "05041": GoSub 58515      'Diagnosis Date
      X = 56: RFNAME$ = "05042": GoSub 58515      'Illness
      X = 57: RFNAME$ = "05043": GoSub 58515      'Latest Status
      X = 58: RFNAME$ = "05044": GoSub 58515      'Status Date
      X = 59: RFNAME$ = "05050": GoSub 58515      'Military Information
      X = 60: RFNAME$ = "05051": GoSub 58515      'Rank
      X = 61: RFNAME$ = "05052": GoSub 58515      'Status
      X = 62: RFNAME$ = "05060": GoSub 58515      'Occupation Information
      X = 63: RFNAME$ = "05061": GoSub 58515      'Type Work
      X = 64: RFNAME$ = "07004": GoSub 58515      'Relationship
      X = 65: RFNAME$ = "07010": GoSub 58515      'Lineage
      GoTo 58530

58515 ' Process Report Field
      Call GetRFLD(RFNAME$, FX)
      If FOUND Then
         RDFFld = RFLDS(FX)
         ALTH = Asc(RDFFld.ALTH): LLTH = Asc(RDFFld.LLTH): HLTH = Asc(RDFFld.HLTH)
         If (X < 20) Or (X > 30) Then
            RFLBL$(X) = MID$(RFLD$(FX), ALTH + 1, LLTH)
            Else
            If (X < 26) Then
               RFLBL$(X) = MID$(RFLD$(FX), ALTH + LLTH + 1, HLTH)
               Else
               RFLBL$(X) = MID$(RFLD$(FX), 1, ALTH)
            End If
         End If
      End If
      RFLTH(X) = Len(RFLBL$(X))
      Return

58530 ' Check Report Width
      OKAY = -1: RWIDTH = RDFReport.PWIDTH
      If RWIDTH = 0 Then RWIDTH = 125
      If PAGE.BWIDTH = 0 Then
         BWIDTH = PRT.FWIDTH - PAGE.IMARGIN - PAGE.OMARGIN
         Else
         BWIDTH = PAGE.BWIDTH
      End If

58560 ' Return to Calling Program
      PAGE.BWIDTH = BWIDTH
End Sub

Sub FREGReport()
    On Local Error GoTo 900
    'Process PRINT Requests
    Call RptOPEN(RRPT$): If Not OKAY Then GoTo 399
    Call RptOptGET("BLKL", X): BLNKBTWN = (X = 1): CHILDSEP = BLNKBTWN
    Call RptOptGET("EURA", X): EuroAdr = (X = 1)
    Call RptOptGET("URMN", X): RomanCAPS = (X = 1)
    Call BLDREGRDef
    ReDim OA(4)
    zRID$ = Space$(5)
    xRID$ = "(" + RFLBL$(2) + zRID$ + ")"
    XOPID$ = RFLBL$(3) + zRID$ + ")"
    XSPID$ = "(" + RFLBL$(4) + zRID$ + ")"
    zCDATE$ = Space$(11)
    Dim BNDATA As NDATA, BPDATA As PDATA, BMDATA As MDATA, BEDATA As EDATA
    Dim NameDATA As NDATA, FatherDATA As PDATA, MotherDATA As PDATA
    Dim MarriageDATA As MDATA, EventDATA As EDATA
    BNDATA.RID = Space$(Len(BNDATA.RID))
    BNDATA.PID = Space$(Len(BNDATA.PID))
    BNDATA.NAME = Space$(Len(BNDATA.NAME))
    BNDATA.BDATE = Space$(Len(BNDATA.BDATE))
    BNDATA.BPLACE = Space$(Len(BNDATA.BPLACE))
    BNDATA.DDATE = Space$(Len(BNDATA.DDATE))
    BNDATA.DPLACE = Space$(Len(BNDATA.DPLACE))
    BNDATA.SEX = Space$(Len(BNDATA.SEX))
    BNDATA.NUMCH = Space$(Len(BNDATA.NUMCH))
    BPDATA.NAME = Space$(Len(BPDATA.NAME))
    BPDATA.BDATE = Space$(Len(BPDATA.BDATE))
    BPDATA.DDATE = Space$(Len(BPDATA.DDATE))
    BPDATA.AGE = Space$(Len(BPDATA.AGE))
    BMDATA.MDATE = Space$(Len(BMDATA.MDATE))
    BMDATA.MPLACE = Space$(Len(BMDATA.MPLACE))
    BMDATA.DVDATE = Space$(Len(BMDATA.DVDATE))
    BMDATA.DVPLACE = Space$(Len(BMDATA.DVPLACE))
    BMDATA.ANNIV = Space$(Len(BMDATA.ANNIV))
    BMDATA.SPID = Space$(Len(BMDATA.SPID))
    BMDATA.SPNAME = Space$(Len(BMDATA.SPNAME))
    BEDATA.CODE = Space$(Len(BEDATA.CODE))
    BEDATA.DESC = Space$(Len(BEDATA.DESC))
    BEDATA.DATE = Space$(Len(BEDATA.DATE))
    BEDATA.DSTAT = Space$(Len(BEDATA.DSTAT))
    BEDATA.PLACE = Space$(Len(BEDATA.PLACE))
    BEDATA.IMPORT = Space$(Len(BEDATA.IMPORT))

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
    ADRDATA$ = Space$(157): MAXADR = 30
    OTHDATA$ = Space$(52)
    ' Build Separator Lines from Report Variables
    LINESEP1$ = Space$(PAGE.BWIDTH)
    Call GetRVAR("LINESEP1", X)
    If FOUND Then
       LSet LINESEP1$ = RVAR$(X): Call HTFBuild(LINESEP1$, PAGE.BWIDTH)
    End If
    LINESEP2$ = Space$(PAGE.BWIDTH)
    Call GetRVAR("LINESEP2", X)
    If FOUND Then
       LSet LINESEP2$ = RVAR$(X): Call HTFBuild(LINESEP2$, PAGE.BWIDTH)
    End If
    LINESEP3$ = Space$(PAGE.BWIDTH)
    Call GetRVAR("LINESEP3", X)
    If FOUND Then
       LSet LINESEP3$ = RVAR$(X): Call HTFBuild(LINESEP3$, PAGE.BWIDTH)
    End If
    ' Build Child Information Header
    ChildHDR$ = RFLBL$(21)
    ' Get Misc Options
    ShowREL = RptOPTION.Relation
    Call RptOptGET("PLIN", X): ShowLIN = (X = 1)
    Call RptOptGET("BYPS", X)
    BYPASS = ((X = 10) Or (X = 11)): BYPASS2 = ((X = 10) Or (X = 12))
    Call RptOptGET("SP12", X): SPRESEQ = (X = 1)
    Call RptOptGET("NGSA", X): CountALL = (X = 1)
    Call RptOptGET("NGSC", X): ShowChNUM = (X = 1)
    Call RptOptGET("NGSR", X): ShowRegNUM = (X = 1)
    Call RptOptGET("NGSN", X): ShowPNAME = (X = 1)
    Call RptOptGET("NGHL", X): RelInHEAD = (X = 1)
    If ShowLIN And Not (ShowRegNUM Or ShowPNAME) Then ShowChNUM = -1
    Call GetRVAR("UNKNOWN", X): If FOUND Then Unknown$ = RVAR$(X)

    ' Prepare Register Tables
    Call FAMOpen("READ"): If Not OKAY Then GoTo 399
    A = 0
    BD$ = Space$(12): BPL$ = Space$(41)
    DD$ = BD$: DPL$ = BPL$: AGE$ = "   "
    DS = 0: FC = Asc(" ")
    sTYPE = 3: SCNT = 0: HIID = 0: LOID = FF1Hdr.MAXID
    Call RWRKOpen("READQ", 0, 0): If Not OKAY Then GoTo 395
    NDXRECS = 0: ReDim NDXID(FF1Hdr.MAXID), SelSET(FF1Hdr.MAXID)
    'IF SHOWLIN THEN
       ReDim RELATE(FF1Hdr.MAXID), RELPTR&(FF1Hdr.MAXID)
       Call RptOptGET("EXAD", X): ExtAdoptRule = (X = 1)
    'END IF
    Call PutMSG("XXXXM064")
    ' The Next Statement Forces to Read Workfile by Generations
    WFHDR.WFRULES = Chr$((Asc(WFHDR.WFRULES) And 252) + 1)
    Call RWRKGet("BASE", RC, PTR&)
    While (0 < RC) And (RC < 9)
       If WFRELREC.id > 0 Then
          RID = WFRELREC.id
          If SelSET(RID) = 0 Then
             SCNT = SCNT + 1: SelSET(RID) = 1
             If RID < LOID Then LOID = RID
             If HIID < RID Then HIID = RID
          End If
          If SelSET(RID) > 0 Then
             NDXRECS = NDXRECS + 1: NDXID(NDXRECS) = RID
             SelSET(RID) = -1
          End If
          'IF SHOWLIN THEN
             If WFGLREC.DGL = 0 Then
                ExtAdopt = ((WFANCREC.PSTAT And 5) > 0)
                Else
                ExtAdopt = ((WFRELREC.CSTAT And 5) > 0)
             End If
             X = RELATE(RID)
             If (X = 0) Or (ExtAdoptRule And ((X And 128) > 0) And (Not ExtAdopt)) Then
                X = 256 * (-WFGLREC.AGL) + WFGLREC.DGL
                If ExtAdoptRule Then X = X - 128 * ExtAdopt
                RELATE(RID) = X
                RELPTR&(RID) = PTR&
             End If
          'END IF
       End If
       Call RWRKGet("WRKREC", RC, PTR&)
    Wend
    SelSET(0) = LOID
    RID = LOID
    If SCNT = 0 Then
       Call PutMSG("REGRM003"): GoTo 399
    End If
    ReDim Preserve NDXID(NDXRECS)
    ' Assign Sequence Numbers to Individuals
    ' (applying BYPASS Rules if in effect)
310 SID = 0: XID = 0
312 SID = SID + 1: If SID > NDXRECS Then id = 0: GoTo 319
    id = NDXID(SID)
315 If (SelSET(id) = 0) Or (id = 0) Then GoTo 312
    Call FF1GetRec(id)
    If ((MID$(IA$, 5, 1) = "1") And (FF1Rec.EDUC <> 0)) Then GoTo 319
    If ((MID$(IA$, 6, 1) = "1") And (FF1Rec.WORK <> 0)) Then GoTo 319
    If ((MID$(IA$, 7, 1) = "1") And (FF1Rec.MIL <> 0)) Then GoTo 319
    If ((MID$(IA$, 8, 1) = "1") And (FF1Rec.HLTH <> 0)) Then GoTo 319
    If ((MID$(IA$, 11, 1) = "1") And (FF1Rec.EVENT <> 0)) Then GoTo 319
    If (BYPASS) And (MID$(IA$, 2, 1) <> "2") Then
       FID = FF1Rec.FID: xMID = FF1Rec.MID
       If ((FID > 0) * (SelSET(FID) <> 0)) + ((xMID > 0) * (SelSET(xMID) <> 0)) = 0 Then GoTo 317
       If (FF1Rec.OLDCH <> 0) Then GoTo 317
       If ((MID$(IA$, 3, 1) = "1") And (FF1Rec.ADRS <> 0)) Then GoTo 317
       If ((MID$(IA$, 4, 1) <> "2") And (FF1Rec.COM <> 0) And (MID$(IA$, 12, 1) = "2")) Then GoTo 317
       If ((MID$(IA$, 1, 1) <> "1") Or (FF1Rec.SPOUSE = 0)) Then GoTo 318
       Call FF3GetRec(FF1Rec.SPOUSE)
       If (FF3SPOUSE.MLOC <> 0) Then GoTo 317
       If ((MID$(IA$, 3, 1) = "1") And (FF3SPOUSE.LOC <> 0)) Then GoTo 317
       If ((MID$(IA$, 4, 1) <> "2") And (FF3SPOUSE.COM <> 0)) Then GoTo 317
       If id = FF3SPOUSE.SP1ID Then
          CSPID = FF3SPOUSE.SP2ID: SC = FF3SPOUSE.SP2NXT
          Else
          CSPID = FF3SPOUSE.SP1ID: SC = FF3SPOUSE.SP1NXT
       End If
       If SC <> 0 Then GoTo 317
       If CSPID = 0 Then GoTo 318 Else Call FF1GetRec(CSPID)
       If (MID$(IA$, 4, 1) = "4") And (FF1Rec.COM <> 0) Then GoTo 317
       FID = FF1Rec.FID: xMID = FF1Rec.MID: If (FID > 0) Or (xMID > 0) Then GoTo 317
       If (FF1Rec.BLOC = 0) And (FF1Rec.BY = 0) And (FF1Rec.BMD = 0) And ((FF1Rec.DY = 0) Or (FF1Rec.DY = 9999)) And (FF1Rec.DMD = 0) Then GoTo 318
    End If
317 If (BYPASS2) And (MID$(IA$, 1, 1) = "1") Then
       Call FF1GetRec(id)
       SC = FF1Rec.SPOUSE: If (FF1Rec.SEX <> FGENDR$(2)) Or (SC = 0) Then GoTo 319
       If (MID$(IA$, 4, 1) <> "1") Or (FF1Rec.COM = 0) Then
          Call FF3GetRec(SC)
          If id = FF3SPOUSE.SP1ID Then
             CSPID = FF3SPOUSE.SP2ID: SC = FF3SPOUSE.SP1NXT
             Else
             CSPID = FF3SPOUSE.SP1ID: SC = FF3SPOUSE.SP2NXT
          End If
          If SC <> 0 Then GoTo 319
          If (CSPID = 0) Or (SelSET(CSPID) = 0) Then GoTo 319
          If (MID$(IA$, 2, 1) = "2") Or (FF1Rec.OLDCH = 0) Then GoTo 318
          CP = FF1Rec.OLDCH
          While CP <> 0
             Call FF1GetRec(CP)
             If CSPID = FF1Rec.FID Then CP = FF1Rec.MCH Else CP = 0
          Wend
          If CSPID = FF1Rec.FID Then GoTo 318
       End If
    End If
    GoTo 319
318 ' Record Bypassed (Assign Register Number if CountALL Option)
    If CountALL Then
       XID = XID + 1: SelSET(id) = -XID
       Else
       SelSET(id) = 0
    End If
    GoTo 312
319 ' Assign Register Number for Records to be Printed
    If id > 0 Then
       XID = XID + 1: SelSET(id) = XID
       GoTo 312
    End If

    ' Begin Printing Register Report
    D1Y = Val(Right$(DATE$, 4))
    D1MD = (100 * Val(Left$(DATE$, 2)) + Val(MID$(DATE$, 4, 2)))
    zCDATE$ = XRPTDate$(D1Y, D1MD, 0)
    MA$ = " 0 0 0 0 0 0 0 0 0 0 0 0"
    AMAXLV = Asc(WFHDR.AMAXLV): DMAXLV = Asc(WFHDR.DMAXLV)
    If DMAXLV = 0 Then AncSECTION = -1
    If (AMAXLV > 0) And (DMAXLV > 0) Then
       Call PutMSG("REGRM004")
       If A = 27 Then GoTo 399
       AncONLY = (CA$ = "1"): DescONLY = (CA$ = "2")
       If (Not DescONLY) And (AMAXLV > 0) Then
          Call GetRVAR("ASECHDR", X): If FOUND Then RSECHDR$ = RVAR$(X)
          Else
          Call GetRVAR("DSECHD1", X): If FOUND Then RSECHDR$ = RVAR$(X)
       End If
       If (Not AncONLY) And (DMAXLV > 0) Then
          Call GetRVAR("DSECHDR", X): If FOUND Then DSECHD1$ = RVAR$(X)
       End If
    End If
    RptOPTION.SNGLREF = -1  ' Allow to Ask About Limiting References Recorded
    POPT$ = SNGLKEY$(4): PRTHDRID = 0
    Call PrintOPEN: If Not OKAY Then GoTo 399
    If Len(PLINE$) < 255 Then PLINE$ = Space$(255)
    CL$ = PLINE$: xLINEAGE$ = CL$
    AFMT = FF2Hdr.REVLVL
    AMAXLV = Asc(WFHDR.AMAXLV): DMAXLV = Asc(WFHDR.DMAXLV)
    If DMAXLV = 0 Then AncSECTION = -1
    Call ClearMSG
    Call PutMSG("XXXXM015")
    PRT.SECTION = "R"
    If (DMAXLV = 0) Then PRT.SECTION = "A"
    If (AMAXLV = 0) Then PRT.SECTION = "D"
    'Select Case RelWType
    '   Case 1: PRT.SECTION = "A"
    '   Case 2: PRT.SECTION = "D"
    '   Case 3: PRT.SECTION = "R"
    'End Select
320 If (AMAXLV > 0) And (DMAXLV > 0) And (Not DescONLY) Then
       AncSECTION = -1: NoDetail = -1
    End If
    
325 SID = 0: RID = 0: NMOFST = 4: SPW = PAGE.BWIDTH - NMOFST
    AGL = 0: DGL = 0:
    If PRT.SECTION = "R" Then BgnSection = -1

330 ' Print Next Register Entry
    SID = SID + 1: If SID > NDXRECS Then id = 0: GoTo 370
    id = NDXID(SID): If id = 0 Then GoTo 330
    If (Not AncSECTION) And (SelSET(id) < 1) Then GoTo 330
    Call FF1GetRec(id)
    LAGL = AGL: LDGL = DGL
    X = RELATE(id)
    AGL = -(X \ 256): DGL = X And 127: RGL = AGL + DGL
    If AncSECTION And (DGL > 0) Then GoTo 330
    ExtAdopt = ((X And 128) > 0)
    Call RWRKGet("RELREC", RC, RELPTR&(id))
    If RID = 0 Then
       LSet SysVAR.SUBJECT = FMTNAME$(1)
       If PRTHDRID Then
          X = Len(RTrim$(SysVAR.SUBJECT))
          If X > 2 Then
             Mid$(SysVAR.SUBJECT, X + 2, 255) = "(" + RIDLbl$ + fnx$(RID) + ")"
          End If
       End If
    End If

350 ' Print Appropriate Section Headers
    ' If AncSECTION And NoDetail And (RID = 0) Then sTYPE = 1: GoTo 355
    If BgnSection Then sTYPE = 1: GoTo 355
    If (AMAXLV > 0) And (DGL = 0) And (Not AncSECTION) Then
       sTYPE = 2: GoTo 355
    End If
    If (AncSECTION And (AGL <> LAGL)) Then
       sTYPE = 3: GoTo 355
    End If
    If ((Not AncSECTION) And (DGL <> LDGL)) Then
       sTYPE = 4: GoTo 355
    End If
    If (RID <> 0) Then
       Call WriteLINE
    End If
    GoTo 360

355 ' Print Section Header
    If (PAGE.BLENGTH - PAGE.LINE) < 5 Then
       Call PageBREAK
       Else
       Call WriteLINE
    End If
    Select Case sTYPE
       Case 1: Mid$(PLINE$, NMOFST, SPW) = LINESEP1$
       Case 2: Mid$(PLINE$, NMOFST, SPW) = LINESEP1$
       Case Else: Mid$(PLINE$, NMOFST, SPW) = LINESEP2$
    End Select
    Call PrintLINE
    Select Case sTYPE
       Case 1: LSet PLINE$ = RSECHDR$
       Case 2: LSet SysVAR.SUBJECT = FMTNAME$(1)
               LSet SysVAR.MVAR = "[" + LTrim$(Str$(WFANCREC.LINEAGE)) + "]"
               LSet PLINE$ = DSECHD1$
       Case 3: LSet PLINE$ = "<<" + Relation$("L", "X", AGL, AGL, RRELCD$()) + ">>"
               LSet PLINE$ = UCX$(PLINE$)
       Case 4: If RelInHEAD Then xAGL = AGL: xRGL = RGL Else xAGL = 0: xRGL = DGL
               LSet PLINE$ = "<<" + Relation$("L", "X", xAGL, xRGL, RRELCD$()) + ">>"
               LSet PLINE$ = UCX$(PLINE$)
    End Select
    Call HTFBuild(PLINE$, PAGE.BWIDTH)
    Call PrintLINE
    Select Case sTYPE
       Case 1: Mid$(PLINE$, NMOFST, SPW) = LINESEP1$
       Case 2: Mid$(PLINE$, NMOFST, SPW) = LINESEP1$
       Case Else: Mid$(PLINE$, NMOFST, SPW) = LINESEP2$
    End Select
    Call PrintLINE
    If sTYPE = 1 Then
       BgnSection = 0
       If AGL < 0 Or DGL > 0 Then GoTo 350
    End If
    Call WriteLINE
    
360 ' Print Register Entry
    RID = id: GoSub PrtReport
    If (A = 27) Then GoTo 395
    GoTo 330

370 ' Check For Additional Sections
    If (AMAXLV = 0) Or (DMAXLV = 0) Then GoTo 395
    If AncONLY Or (Not AncSECTION) Then GoTo 395
    Call EvenBREAK("noEOF"): If A = 27 Then GoTo 395
    AncSECTION = 0: NoDetail = 0: RSECHDR$ = ""
    Call GetRVAR("DSECHD1", X): If FOUND Then RSECHDR$ = RVAR$(X)
    GoTo 325

395 ' Finished Printing Reports
    Call RptCLOSE
    If Not OKAY Then GoTo 399
    If PRT.AGAIN Then
       PAGE.COUNT = PAGE.BGNPAGE - 1
       GoTo 320
    End If
399 ' Return to Main Program
    On Local Error GoTo 0
    Call FamCLOSE
    Exit Sub

900 ' Error Routine
    SysVAR.NVAR1 = ERL
    Call ERRMessage("XXXXM098", ERR, ERL)
    Resume 399

PrtReport: ' Print Family Group Report
    If FF1Rec.RID <> RID Then Call FF1GetRec(RID)
    ReDim APTR(MAXADR), AIX(MAXADR), ADRDATE$(MAXADR)
    AX = -1: APTR(0) = FF1Rec.ADRS: If APTR(0) <> 0 Then AX = 0
    SC = FF1Rec.SPOUSE: CC = FF1Rec.OLDCH: ComPTR = FF1Rec.COM
    Call BldMLIST
    EVENTPTR = FF1Rec.EVENT
    OA(1) = FF1Rec.HLTH: OA(2) = FF1Rec.EDUC
    OA(3) = FF1Rec.WORK: OA(4) = FF1Rec.MIL
    GoSub 600: GoSub 610: MCNT = 0
    GoSub 2300: If A = 27 Then GoTo 590
    If NoDetail Then GoTo 590
    If MID$(IA$, 1, 1) <> "2" Then
       GoSub 600
       While SC <> 0
          Call FF3GetRec(SC): Call GetNextMARRIAGE
          MLOC = FF3SPOUSE.MLOC
          If RID = FF3SPOUSE.SP1ID Then
             CSPID = FF3SPOUSE.SP2ID
             Else
             CSPID = FF3SPOUSE.SP1ID
          End If
          ComPTR = FF3SPOUSE.COM
          APTR(AX + 1) = FF3SPOUSE.LOC: If APTR(AX + 1) > 0 Then AX = AX + 1
          If CSPID > 0 Then
             Call FF1GetRec(CSPID): COMPTR2 = FF1Rec.COM
             GoSub 610
             Else
             COMPTR2 = 0: NameDATA.NAME = "(" + Unknown$ + ")"
             FID = 0: xMID = 0
          End If
          GoSub 630: GoSub 2350: If A = 27 Then GoTo 590
          GoSub 600
       Wend
       CSPID = 0: xMID = 0: FID = 0
       MLOC = 0
    End If
    If (MID$(IA$, 2, 1) <> "2") And ((CC <> 0)) Then
       GoSub 2400: If A = 27 Then GoTo 590
       CH = 0
       While CC <> 0
          Call FF1GetRec(CC): ComPTR = FF1Rec.COM
          STAT1 = Asc(FF1Rec.STATUS)
          MSTAT = ((STAT1 \ 4) And 3): FSTAT = (STAT1 And 3)
          If RID = FF1Rec.MID Then
             NCH = FF1Rec.MCH: PID = FF1Rec.FID: CSTAT = MSTAT: PSTAT = FSTAT
             Else
             NCH = FF1Rec.FCH: PID = FF1Rec.MID: CSTAT = FSTAT: PSTAT = MSTAT
          End If
          NameDATA = BNDATA
          NameDATA.RID = SFILL$(CC, 5)
          NameDATA.PID = SFILL$(PID, 5)
          NameDATA.NAME = FMTNAME$(0): GoSub 615: GoSub 650
          MarriageDATA = BMDATA
          If FF1Rec.SPOUSE <> 0 Then
             Call FF3GetRec(FF1Rec.SPOUSE)
             MLOC = FF3SPOUSE.MLOC
             GoSub 630
             If CC = FF3SPOUSE.SP1ID Then CSPID = FF3SPOUSE.SP2ID Else CSPID = FF3SPOUSE.SP1ID
             If CSPID = 0 Then
                MarriageDATA.SPNAME = "(" + Unknown$ + ")"
                Else
                Call FF1GetRec(CSPID)
                MarriageDATA.SPID = SFILL$(CSPID, 5)
                MarriageDATA.SPNAME = FMTNAME$(0)
             End If
          End If
          GoSub 2410: If A = 27 Then GoTo 590
          CC = NCH
       Wend
    End If

460 ' Process RESIDENCEes
    If (MID$(IA$, 3, 1) <> "2") And ((APTR(0) <> 0)) Then
       GoSub PrintRESH: If A = 27 Then GoTo 590
       Y = 0
       While APTR(Y) <> 0
          If Not (AX < MAXADR) Then
             MAXADR = MAXADR + 10
             ReDim Preserve APTR(MAXADR), AIX(MAXADR), ADRDATE$(MAXADR)
          End If
          ADRDATE$(Y) = Space$(8)
          Call FF2GetRec(APTR(Y))
          APTR(AX + 1) = FF2PFX.NXT
          If (APTR(AX + 1) <> 0) And (AX < MAXADR) Then AX = AX + 1
          Mid$(ADRDATE$(Y), 1, 4) = ZFILL$(FF2PFX.D1Y, 4)
          Mid$(ADRDATE$(Y), 5, 4) = ZFILL$(FF2PFX.D1MD, 4)
          AIX(Y) = Y: z = Y - 1
          While (z >= 0)
             If Left$(ADRDATE$(AIX(z)), 4) < Left$(ADRDATE$(AIX(z + 1)), 4) Then
                ZZ = AIX(z): AIX(z) = AIX(z + 1): AIX(z + 1) = ZZ
                z = z - 1
                Else
                z = -1
             End If
          Wend
          If Y < MAXADR Then Y = Y + 1 Else APTR(Y) = 0
       Wend
       Y = 0
       While Y <= AX
          z = AIX(Y)
          xAPTR = APTR(z)
          GoSub PrintRESI
          If A = 27 Then GoTo 590 Else Y = Y + 1
       Wend
    End If

470 ' Process EVENTS
    If (MID$(IA$, 11, 1) <> "2") And ((EVENTPTR <> 0)) Then
       GoSub PrintEventH: If A = 27 Then GoTo 590
       While EVENTPTR <> 0
          GoSub PrintEVENT: EVENTPTR = FF3EVENT.NXT
       Wend
    End If

500 ' Process OTHER Info
    For OT = 1 To 4: GoSub 2700: Next OT
590 ' End of PrtReport
    If A = 27 Then LSet PLINE$ = " ": LOFFSET = TWIDTH
    Return

600 ' Initialise Data Areas - Individual
    NameDATA = BNDATA
    FatherDATA = BPDATA
    MotherDATA = BPDATA
    MarriageDATA = BMDATA
    Return

610 ' Place Information in Data Areas
    NameDATA.NAME = FMTNAME$(0)
    FID = FF1Rec.FID: xMID = FF1Rec.MID
    STAT1 = Asc(FF1Rec.STATUS)
    MSTAT = ((STAT1 \ 4) And 3): FSTAT = (STAT1 And 3)
    GoSub 620

615 ' Name Record Information Common to Subject, Spouse, Child
    Call GetBDEvents(BD$, BPL$, DD$, DPL$, "")
    NameDATA.BDATE = LTrim$(BD$)
    NameDATA.DDATE = LTrim$(DD$)
    NameDATA.BPLACE = BPL$
    NameDATA.DPLACE = DPL$
    NameDATA.SEX = FF1Rec.SEX
    Return

620 ' Father, Mother Information
    XID = FF1Rec.RID
    If FID <> 0 Then
       Call FF1GetRec(FID): STAT1 = Asc(FF1Rec.STATUS)
       FatherDATA.NAME = FMTNAME$(0)
       Call GetBDEvents(BD$, "", DD$, "", AGE$)
       FatherDATA.BDATE = LTrim$(BD$)
       FatherDATA.DDATE = LTrim$(DD$)
       If (FF1Rec.BY > 0) And (FF1Rec.DY > 0) And (FF1Rec.DY < 9999) Then
          FatherDATA.AGE = LTrim$(AGE$)
       End If
    End If
    If xMID <> 0 Then
       Call FF1GetRec(xMID): STAT1 = Asc(FF1Rec.STATUS)
       RSet zRID$ = fnx$(xMID): RSTAT = MSTAT
       MotherDATA.NAME = FMTNAME$(0)
       Call GetBDEvents(BD$, "", DD$, "", AGE$)
       MotherDATA.BDATE = LTrim$(BD$)
       MotherDATA.DDATE = LTrim$(DD$)
       If (FF1Rec.BY > 0) And (FF1Rec.DY > 0) And (FF1Rec.DY < 9999) Then
          MotherDATA.AGE = LTrim$(AGE$)
       End If
    End If
    Call FF1GetRec(XID)
    Return

630 ' Marriage Information
    MarriageDATA.MDATE = XRPTDate$(FF3SPOUSE.D1MD, FF3SPOUSE.D1Y, DS)
    MarriageDATA.DVDATE = XRPTDate$(FF3SPOUSE.D2MD, FF3SPOUSE.D2Y, DS)
    If MLOC <> 0 Then
       Call FF3GetRec(MLOC)
       MarriageDATA.MPLACE = FF3PLACE.PLACE1
       MarriageDATA.DVPLACE = FF3PLACE.PLACE2
    End If
    Return

650 ' Count Children of Child
    GCH = 0
    If CC <> 0 Then
       NGCH = FF1Rec.OLDCH
       While NGCH <> 0
          GCH = GCH + 1: Call FF1GetRec(NGCH): NGCH = 0
          If CC = FF1Rec.FID Then
             NGCH = FF1Rec.FCH
             Else
             If CC = FF1Rec.MID Then NGCH = FF1Rec.MCH
          End If
       Wend
       Call FF1GetRec(CC)
    End If
    If GCH > 0 Then NameDATA.NUMCH = LTrim$(Str$(GCH))
    Return

2300 ' Print Subject Information
     PAGE.LOFST = 1: PAGE.COFST = 3: z = 1
     If AncSECTION Then
        Call RWRKGet("RELREC", RC, RELPTR&(RID))
        LSet CL$ = LTrim$(Str$(WFANCREC.LINEAGE))
        Else
        LSet CL$ = fnx$(Abs(SelSET(RID)))
     End If
     z = InStr(z, CL$, "  ")
     Mid$(CL$, z, 1) = ".": z = z + 2
     PAGE.COFST = z
     If ShowLIN And (Not AncSECTION) Then
        Call LINEAGE("BUILD")
        If ShowREL Then Call LINEAGE("PRINT")
     End If
     NMOFST = PAGE.COFST: SPW = PAGE.BWIDTH - NMOFST: z = NMOFST
     If AncSECTION And NoDetail Then
        Mid$(CL$, z, 255) = "(" + fnx$(Abs(SelSET(RID))) + ")"
        z = InStr(z, CL$, " ") + 1
     End If
     Mid$(CL$, z, 255) = RFLBL$(5): z = z + RFLTH(5) + (RFLTH(5) = 0)
     If RptOPTION.PRID Then
        Mid$(CL$, z + 1, 255) = "(" + RFLBL$(2): z = z + RFLTH(2) + 2
        Mid$(CL$, z, 6) = fnx$(RID) + ")": z = InStr(z, CL$, "  ")
     End If
     Mid$(CL$, z + 1, 52) = NameDATA.NAME: z = InStr(z, CL$, "   ")
     PAGE.XID = RID: PAGE.xType = 2
     If ShowLIN And (Not ShowREL) And (Not AncSECTION) Then
        If SID > 1 Then Call LINEAGE("PRINT"): z = 1
     End If
     If (RTrim$(NameDATA.BDATE) <> "") Or (RTrim$(NameDATA.BPLACE) <> "") Then
        Mid$(CL$, z, 255) = RFLBL$(8): z = z + RFLTH(8)
        If RTrim$(NameDATA.BDATE) <> "" Then
           Mid$(CL$, z + 1, 12) = NameDATA.BDATE: z = InStr(z, CL$, "   ")
        End If
        If RTrim$(NameDATA.BPLACE) <> "" Then
           Mid$(CL$, z, 255) = RFLBL$(9): z = z + RFLTH(9)
           Mid$(CL$, z + 1, 52) = NameDATA.BPLACE: z = InStr(z, CL$, "   ")
        End If
     End If
     If (RTrim$(NameDATA.DDATE) <> "") Or (RTrim$(NameDATA.DPLACE) <> "") Then
        Mid$(CL$, z, 255) = RFLBL$(10): z = z + RFLTH(10)
        If RTrim$(NameDATA.DDATE) <> "" Then
           Mid$(CL$, z + 1, 12) = NameDATA.DDATE: z = InStr(z, CL$, "   ")
        End If
        If RTrim$(NameDATA.DPLACE) <> "" Then
           Mid$(CL$, z, 255) = RFLBL$(11): z = z + RFLTH(11)
           Mid$(CL$, z + 1, 41) = NameDATA.DPLACE
        End If
     End If
     Call PutDATA
2310 If FID <> 0 Then
        Mid$(CL$, 1, 255) = RFLBL$(6): z = 1 + RFLTH(6)
        If RptOPTION.PRID Then
           Mid$(CL$, z + 1, 255) = "(" + RFLBL$(2): z = z + 2 + RFLTH(2)
           Mid$(CL$, z, 6) = fnx$(FID) + ")": z = InStr(z, CL$, "  ")
        End If
        Mid$(CL$, z + 1, 52) = FatherDATA.NAME: z = InStr(z, CL$, "   ")
        If RTrim$(FatherDATA.BDATE) <> "" Then
           Mid$(CL$, z, 255) = RFLBL$(8): z = z + RFLTH(8)
           Mid$(CL$, z + 1, 12) = FatherDATA.BDATE: z = InStr(z + 1, CL$, "   ")
        End If
        If (RTrim$(FatherDATA.DDATE) <> "") And (MID$(FatherDATA.DDATE, 1, 3) <> "999") Then
           Mid$(CL$, z, 255) = RFLBL$(10): z = z + RFLTH(10)
           Mid$(CL$, z + 1, 12) = FatherDATA.DDATE: z = InStr(z, CL$, "   ")
           If RTrim$(FatherDATA.AGE) <> "" Then
              Mid$(CL$, z, 255) = RFLBL$(12): z = z + RFLTH(12)
              Mid$(CL$, z + 1, 3) = FatherDATA.AGE
           End If
        End If
        PAGE.XID = FID: PAGE.xType = 5
        Call PutDATA
     End If
2311 If xMID <> 0 Then
        Mid$(CL$, 1, 255) = RFLBL$(7): z = 1 + RFLTH(7)
        If RptOPTION.PRID Then
           Mid$(CL$, z + 1, 255) = "(" + RFLBL$(2): z = z + 2 + RFLTH(2)
           Mid$(CL$, z, 6) = fnx$(xMID) + ")": z = InStr(z, CL$, "  ")
        End If
        Mid$(CL$, z + 1, 52) = MotherDATA.NAME: z = InStr(z, CL$, "   ")
        If RTrim$(MotherDATA.BDATE) <> "" Then
           Mid$(CL$, z, 255) = RFLBL$(8): z = z + RFLTH(8)
           Mid$(CL$, z + 1, 12) = MotherDATA.BDATE: z = InStr(z, CL$, "   ")
        End If
        If (RTrim$(MotherDATA.DDATE) <> "") And (MID$(MotherDATA.DDATE, 1, 3) <> "999") Then
           Mid$(CL$, z, 255) = RFLBL$(10): z = z + RFLTH(10)
           Mid$(CL$, z + 1, 12) = MotherDATA.DDATE: z = InStr(z, CL$, "   ")
           If MotherDATA.AGE <> Space$(3) Then
              Mid$(CL$, z, 255) = RFLBL$(12): z = z + RFLTH(12)
              Mid$(CL$, z + 1, 3) = MotherDATA.AGE
           End If
        End If
2312    PAGE.XID = xMID: PAGE.xType = 6
        Call PutDATA
     End If
2315 If MID$(IA$, 4, 1) = "2" Then Call PrintLINE: GoTo 2325
     GoSub PrintCOM
     Call PrintLINE
2325 Return

2350 ' Print Spouse Info
     Call WriteLINE: If A = 27 Then GoTo 2380
     PAGE.LOFST = NMOFST: PAGE.COFST = NMOFST
     Mid$(CL$, 1, 255) = RFLBL$(14): z = RFLTH(14) + 1
     If RTrim$(MarriageDATA.MDATE) <> "" Then
        Mid$(CL$, z + 1, 11) = MarriageDATA.MDATE: z = InStr(z, CL$, "   ")
     End If
     If RTrim$(MarriageDATA.MPLACE) <> "" Then
        Mid$(CL$, z, 255) = RFLBL$(15): z = z + RFLTH(15)
        Mid$(CL$, z + 1, 41) = MarriageDATA.MPLACE: z = InStr(z, CL$, "   ")
     End If
     If RTrim$(MarriageDATA.DVDATE) <> "" Then
        Mid$(CL$, z, 255) = RFLBL$(16): z = z + RFLTH(16)
        Mid$(CL$, z + 1, 11) = MarriageDATA.DVDATE: z = InStr(z, CL$, "  ")
        If (RTrim$(MarriageDATA.DVPLACE) <> "") Then
           Mid$(CL$, z, 255) = RFLBL$(17): z = z + RFLTH(17)
           Mid$(CL$, z + 1, 41) = MarriageDATA.DVPLACE: z = InStr(z, CL$, "   ")
        End If
     End If
     Mid$(CL$, z, 255) = RFLBL$(13): z = z + RFLTH(13)
     If CSPID <> 0 Then
        If (SelSET(CSPID) <> 0) And ((RELATE(CSPID) And 127) = 0) Then
           Call RWRKGet("RELREC", RC, RELPTR&(CSPID))
           Mid$(CL$, z + 1, 255) = "[" + LTrim$(Str$(WFANCREC.LINEAGE)) + "]"
           z = InStr(z, CL$, "  ")
        End If
        If RptOPTION.PRID Then
           Mid$(CL$, z + 1, 255) = "(" + RFLBL$(2): z = z + RFLTH(2) + 2
           Mid$(CL$, z, 6) = fnx$(CSPID) + ")": z = InStr(z, CL$, "  ")
        End If
     End If
     Mid$(CL$, z + 1, 52) = NameDATA.NAME: z = InStr(z, CL$, "   ")
     If (RTrim$(NameDATA.BDATE) <> "") Or (RTrim$(NameDATA.BPLACE) <> "") Then
        Mid$(CL$, z, 255) = RFLBL$(8): z = z + RFLTH(8)
        If RTrim$(NameDATA.BDATE) <> "" Then
           Mid$(CL$, z + 1, 12) = NameDATA.BDATE: z = InStr(z, CL$, "   ")
        End If
        If RTrim$(NameDATA.BPLACE) <> "" Then
           Mid$(CL$, z, 255) = RFLBL$(9): z = z + RFLTH(9)
           Mid$(CL$, z + 1, 41) = NameDATA.BPLACE: z = InStr(z, CL$, "   ")
        End If
     End If
     If (RTrim$(NameDATA.DDATE) <> "") Or (RTrim$(NameDATA.DPLACE) <> "") Then
        Mid$(CL$, z, 255) = RFLBL$(10): z = z + RFLTH(10)
        If RTrim$(NameDATA.DDATE) <> "" Then
           Mid$(CL$, z + 1, 12) = NameDATA.DDATE: z = InStr(z, CL$, "   ")
        End If
        If RTrim$(NameDATA.DPLACE) <> "" Then
           Mid$(CL$, z, 255) = RFLBL$(11): z = z + RFLTH(11)
           Mid$(CL$, z + 1, 41) = NameDATA.DPLACE
        End If
     End If
     PAGE.XID = CSPID: PAGE.xType = 3: Call PutDATA
     If FID <> 0 Then
        Mid$(CL$, 1, 255) = RFLBL$(6): z = 1 + RFLTH(6)
        If RptOPTION.PRID Then
           Mid$(CL$, z + 1, 255) = "(" + RFLBL$(2): z = z + 2 + RFLTH(2)
           Mid$(CL$, z, 6) = fnx$(FID) + ")": z = InStr(z, CL$, "  ")
        End If
        Mid$(CL$, z + 1, 52) = FatherDATA.NAME: z = InStr(z, CL$, "   ")
        If RTrim$(FatherDATA.BDATE) <> "" Then
           Mid$(CL$, z, 255) = RFLBL$(8): z = z + RFLTH(8)
           Mid$(CL$, z + 1, 12) = FatherDATA.BDATE
           z = InStr(z + 1, CL$, "   ")
        End If
        If (RTrim$(FatherDATA.DDATE) <> "") And (MID$(FatherDATA.DDATE, 1, 3) <> "999") Then
           Mid$(CL$, z, 255) = RFLBL$(10): z = z + RFLTH(10)
           Mid$(CL$, z + 1, 12) = FatherDATA.DDATE: z = InStr(z, CL$, "   ")
           If RTrim$(FatherDATA.AGE) <> "" Then
              Mid$(CL$, z, 255) = RFLBL$(12): z = z + RFLTH(12)
              Mid$(CL$, z + 1, 3) = FatherDATA.AGE
           End If
        End If
        PAGE.XID = FID: PAGE.xType = 7: Call PutDATA
     End If
     If xMID <> 0 Then
        Mid$(CL$, 1, 255) = RFLBL$(7): z = 1 + RFLTH(7)
        If RptOPTION.PRID Then
           Mid$(CL$, z + 1, 255) = "(" + RFLBL$(2): z = z + 2 + RFLTH(2)
           Mid$(CL$, z, 6) = fnx$(xMID) + ")": z = InStr(z, CL$, "  ")
        End If
        Mid$(CL$, z + 1, 52) = MotherDATA.NAME: z = InStr(z, CL$, "   ")
        If RTrim$(MotherDATA.BDATE) <> "" Then
           Mid$(CL$, z, 255) = RFLBL$(8): z = z + RFLTH(8)
           Mid$(CL$, z + 1, 12) = MotherDATA.BDATE: z = InStr(z, CL$, "   ")
        End If
        If (RTrim$(MotherDATA.DDATE) <> "") And (MID$(MotherDATA.DDATE, 1, 3) <> "999") Then
           Mid$(CL$, z, 255) = RFLBL$(10): z = z + RFLTH(10)
           Mid$(CL$, z + 1, 12) = MotherDATA.DDATE: z = InStr(z, CL$, "   ")
           If MotherDATA.AGE <> Space$(3) Then
              Mid$(CL$, z, 255) = RFLBL$(12): z = z + RFLTH(12)
              Mid$(CL$, z + 1, 3) = MotherDATA.AGE
           End If
        End If
        PAGE.XID = xMID: PAGE.xType = 8: Call PutDATA
     End If
     If MID$(IA$, 4, 1) <> "2" Then
        If (MID$(IA$, 4, 1) <> "4") Then COMPTR2 = 0
        GoSub PrintCOM
     End If
2380 MCNT = MCNT + 1
        Return

2400 ' Print Child Information Header
     Call WriteLINE
     If A <> 27 Then
        Mid$(PLINE$, NMOFST, 255) = ChildHDR$
        Call PrintLINE
     End If
     Return

2410 ' Print Child Information
     PAGE.LOFST = NMOFST: CH = CH + 1: z = NMOFST: L = 7
     If Not AncSECTION Then
        If CountALL Then
           If SelSET(CC) > 0 Then Mid$(PLINE$, z - 1, 1) = "+"
        End If
        If SelSET(CC) <> 0 Then
           Mid$(PLINE$, z, 255) = fnx$(Abs(SelSET(CC))) + "."
           z = InStr(z, PLINE$, "  ")
        End If
        z = z - NMOFST: If z < 4 Then z = 4
        L = 11 - z: z = NMOFST + z
     End If
     Mid$(PLINE$, z, L) = Right$(ROMAN$(CH), L)
     z = z + L
     If CSTAT > 0 Then
        FF1Rec.STATUS = Chr$(CSTAT)
        Mid$(PLINE$, z, 1) = FDSTAT$("FID"): CSTAT = 0
     End If
     z = z + 1: PAGE.LOFST = z: PAGE.COFST = z: z = 1

     'SKIP Gender for Now
     'MID$(CL$, 1, 1) = MID$(NameDATA$, 131, 1): Z = 3
     'PAGE.COFST = PAGE.LOFST + 2

     'IF AncSECTION AND (NOT NoDetail) THEN
     If ShowLIN Then
        If (SelSET(CC) <> 0) And ((RELATE(CC) And 127) = 0) Then
           Call RWRKGet("RELREC", RC, RELPTR&(CC))
           Mid$(CL$, z, 255) = "[" + LTrim$(Str$(WFANCREC.LINEAGE)) + "]"
           z = InStr(z, CL$, "  ") + 1
        End If
     End If
     'END IF
     If RptOPTION.PRID Then
        Mid$(CL$, z, 255) = "(" + RFLBL$(2) + LTrim$(NameDATA.RID)
        z = InStr(z, CL$, "  ")
        Mid$(CL$, z, 255) = RFLBL$(3) + LTrim$(NameDATA.PID)
        z = InStr(z, CL$, "  ")
        If PSTAT <> 0 Then
           FF1Rec.STATUS = Chr$(PSTAT)
           Mid$(CL$, z, 1) = FDSTAT$("FID"): PSTAT = 0
           z = InStr(z, CL$, " ")
        End If
        Mid$(CL$, z, 1) = ")": z = z + 2
     End If
     Mid$(CL$, z, 52) = NameDATA.NAME: z = InStr(z, CL$, "   ")
     If (RTrim$(NameDATA.BDATE) <> "") Or (RTrim$(NameDATA.BPLACE) <> "") Then
        Mid$(CL$, z, 255) = RFLBL$(8): z = z + RFLTH(8)
        If RTrim$(NameDATA.BDATE) <> "" Then
           Mid$(CL$, z + 1, 12) = NameDATA.BDATE: z = InStr(z, CL$, "   ")
        End If
        If RTrim$(NameDATA.BPLACE) <> "" Then
           Mid$(CL$, z, 255) = RFLBL$(9): z = z + RFLTH(9)
           Mid$(CL$, z + 1, 52) = NameDATA.BPLACE
           z = InStr(z, CL$, "   ")
        End If
     End If
     If (RTrim$(NameDATA.DDATE) <> "") Or (RTrim$(NameDATA.DPLACE) <> "") Then
        Mid$(CL$, z, 255) = RFLBL$(10): z = z + RFLTH(10)
        If RTrim$(NameDATA.DDATE) <> "" Then
           Mid$(CL$, z + 1, 12) = NameDATA.DDATE: z = InStr(z, CL$, "   ")
        End If
        If RTrim$(NameDATA.DPLACE) <> "" Then
           Mid$(CL$, z, 255) = RFLBL$(11): z = z + RFLTH(11)
           Mid$(CL$, z + 1, 41) = NameDATA.DPLACE: z = InStr(z, CL$, "   ")
        End If
     End If
     If NameDATA.NUMCH <> Space$(3) Then
        Mid$(CL$, z, 1) = ","
        xNCH = Val(NameDATA.NUMCH)
        Mid$(CL$, z + 1, 3) = Str$(xNCH)
        z = InStr(z + 2, CL$, "   ") + 1
        If xNCH > 1 Then
           Mid$(CL$, z, 255) = RFLBL$(23)
           Else
           Mid$(CL$, z, 255) = RFLBL$(67)
        End If
        z = InStr(z, CL$, "  ")
     End If
     PAGE.XID = CC: PAGE.xType = 4: Call PutDATA
     If RTrim$(MarriageDATA.SPNAME) <> "" Then
        Mid$(CL$, 1, 255) = RFLBL$(25): z = 1 + RFLTH(25)
        If RptOPTION.PRID Then
           Mid$(CL$, z + 1, 255) = "(" + RFLBL$(2): z = z + RFLTH(2) + 2
           Mid$(CL$, z, 6) = LTrim$(MarriageDATA.SPID) + ")"
           z = InStr(z, CL$, "  ")
        End If
        Mid$(CL$, z + 1, 52) = MarriageDATA.SPNAME: z = InStr(z, CL$, "   ")
        If RTrim$(MarriageDATA.MDATE) <> "" Then
           Mid$(CL$, z, 255) = RFLBL$(30): z = z + RFLTH(30)
           Mid$(CL$, z + 1, 11) = MarriageDATA.MDATE
           z = InStr(z, CL$, "   ")
           If RTrim$(MarriageDATA.MPLACE) <> "" Then
              Mid$(CL$, z, 255) = RFLBL$(15): z = z + RFLTH(15)
              Mid$(CL$, z + 1, 41) = MarriageDATA.MPLACE: z = InStr(z, CL$, "   ")
           End If
        End If
        PAGE.XID = CSPID: PAGE.xType = 9: Call PutDATA
     End If
     If MID$(IA$, 12, 1) <> "2" Then GoSub PrintCOM
     Call PrintLINE
     'IF CHILDSEP AND (NCH > 0) THEN CALL WriteLINE
     Return

2500 ' Print Events
PrintEventH:
     'MID$(PLINE$, NMOFST, SPW) = LINESEP2$: CALL PrintLINE
     Call WriteLINE
     EVNTCNT = 0
     If A <> 27 Then
        Mid$(PLINE$, NMOFST, 255) = RFLBL$(42)
        Call PrintLINE
     End If
     Return

PrintEVENT:
     If EVNTCNT <> 0 Then
        'MID$(PLINE$, NMOFST, SPW) = LINESEP3$: CALL PrintLINE
        Call WriteLINE
     End If
2510 EVNTCNT = EVNTCNT + 1: EventDATA = BEDATA: ComPTR = 0
     Call FF3GetRec(EVENTPTR): ComPTR = FF3EVENT.COM
     EventDATA.CODE = FF3EVENT.CODE
     M = Val(REVENT$(0)): X = 0: FOUND = 0
     While (X < M) And (Not FOUND)
        X = X + 1
        FOUND = (RTrim$(EventDATA.CODE) = RTrim$(MID$(REVENT$(X), 1, 8)))
     Wend
     If FOUND Then
        EventDATA.DESC = MID$(REVENT$(X), 9, 15)
        Else
        EventDATA.DESC = EventDATA.CODE
     End If
     EventDATA.DATE = XRPTDate$(FF3EVENT.D1MD, FF3EVENT.D1Y, DS)
     EventDATA.DSTAT = FF3EVENT.DSTAT
     If FF3EVENT.LPLNM <> 0 Then
        Call FF3GetRec(FF3EVENT.LPLNM)
        LSet EventDATA.PLACE = FF3PLACE2.PLACE
        Else
        LSet EventDATA.PLACE = FF3EVENT.PLACE
     End If
     X = Asc(FF3EVENT.IMPORT)
     If X > 0 Then EventDATA.IMPORT = LTrim$(Str$(X))

2520 ' Print Information Formatted into EventDATA
     PAGE.LOFST = NMOFST + 2: PAGE.COFST = PAGE.LOFST: X = 1
     Mid$(CL$, X, 255) = RFLBL$(43): X = X + RFLTH(43)
     Mid$(CL$, X + 1, 15) = EventDATA.DESC
     X = InStr(X, CL$, "  ")
2525 If RTrim$(EventDATA.DATE) <> "" Then
        Mid$(CL$, X, 255) = RFLBL$(44): X = X + RFLTH(44)
        Mid$(CL$, X + 1, 12) = EventDATA.DATE
        X = InStr(X, CL$, "  ")
        If EventDATA.DSTAT <> " " Then
           Mid$(CL$, X, 3) = "(" + EventDATA.DSTAT + ")": X = X + 3
        End If
     End If
2530 If RTrim$(EventDATA.PLACE) <> "" Then
        Mid$(CL$, X, 255) = RFLBL$(45): X = X + RFLTH(45)
        Mid$(CL$, X + 1, 255) = EventDATA.PLACE
        X = InStr(X, CL$, "  ")
     End If
     If (RTrim$(EventDATA.IMPORT) <> "") Then
        Mid$(CL$, X, 255) = RFLBL$(46): X = X + RFLTH(46)
        Mid$(CL$, X, 3) = EventDATA.IMPORT
     End If
     LSet PLINE$ = "": Call PutDATA
     If MID$(IA$, 10, 1) = "1" Then
        GoSub PrintCOM
     End If
     Return

PrintRESH:
     'MID$(PLINE$, NMOFST, SPW) = LINESEP2$: CALL PrintLINE
     Call WriteLINE
     If A <> 27 Then
        Mid$(PLINE$, NMOFST, 255) = RFLBL$(31)
        'CALL PrintLINE
     End If
     Return

PrintRESI: Call WriteLINE: If A = 27 Then GoTo 2595
     LSet ADRDATA$ = " ": ComPTR = 0
     Call FF2GetRec(xAPTR): ComPTR = FF2PFX.COM
     Mid$(ADRDATA$, 1, 11) = RTrim$(XRPTDate$(FF2PFX.D1MD, FF2PFX.D1Y, DS))
     Mid$(ADRDATA$, 12, 11) = RTrim$(XRPTDate$(FF2PFX.D2MD, FF2PFX.D2Y, DS))
     If AFMT = 0 Then
        If (FF2SREC.FON1 <> 0) Or (FF2SREC.FON2 <> 0) Or (FF2SREC.FON3 <> 0) Then
           Mid$(ADRDATA$, 23, 14) = ZFILL$(FF2SREC.FON1, 3) + "-" + ZFILL$(FF2SREC.FON2, 3) + "-" + ZFILL$(FF2SREC.FON3, 4) + "  "
        End If
        Mid$(ADRDATA$, 37, 30) = LTrim$(RTrim$(FF2SREC.LINE1))
        Mid$(ADRDATA$, 67, 30) = LTrim$(RTrim$(FF2SREC.LINE2))
        Mid$(ADRDATA$, 97, 15) = LTrim$(RTrim$(FF2SREC.CITY))
        Mid$(ADRDATA$, 117, 4) = LTrim$(RTrim$(FF2SREC.STATE))
        Mid$(ADRDATA$, 129, 5) = LTrim$(RTrim$(FF2SREC.ZIP))
        Else
        If (LTrim$(FF2LREC.PHONE) <> "") And (Left$(FF2LREC.PHONE, 13) <> "(000)000-0000") And (Left$(FF2LREC.PHONE, 12) <> "000-000-0000") Then
           Mid$(ADRDATA$, 23, 14) = LTrim$(RTrim$(FF2LREC.PHONE))
        End If
        Mid$(ADRDATA$, 37, 30) = LTrim$(RTrim$(FF2LREC.LINE1))
        Mid$(ADRDATA$, 67, 30) = LTrim$(RTrim$(FF2LREC.LINE2))
        Mid$(ADRDATA$, 97, 20) = LTrim$(RTrim$(FF2LREC.CITY))
        Mid$(ADRDATA$, 117, 12) = LTrim$(RTrim$(FF2LREC.STATE))
        Mid$(ADRDATA$, 129, 12) = LTrim$(RTrim$(FF2LREC.ZIP))
        Mid$(ADRDATA$, 141, 16) = LTrim$(RTrim$(FF2LREC.COUNTRY))
     End If

     ' Print Information Formatted into ADRDATA$
     PAGE.LOFST = NMOFST + 2: PAGE.COFST = PAGE.LOFST
     X = 1: LSet CL$ = ""
     If (MID$(ADRDATA$, 1, 1) <> Chr$(FC)) Then
        Mid$(CL$, X, 255) = RFLBL$(32): X = X + RFLTH(32)
        Mid$(CL$, X + 1, 11) = MID$(ADRDATA$, 1, 11)
        X = InStr(X, CL$, "     ")
     End If
     If (MID$(ADRDATA$, 12, 1) <> Chr$(FC)) Then
        Mid$(CL$, X, 255) = RFLBL$(33): X = X + RFLTH(33)
        Mid$(CL$, X + 1, 11) = MID$(ADRDATA$, 12, 11)
        X = InStr(X, CL$, "   ")
     End If
     If (MID$(ADRDATA$, 23, 1) <> Chr$(FC)) Then
        Mid$(CL$, X, 255) = RFLBL$(34): X = X + RFLTH(34)
        Mid$(CL$, X + 1, 14) = MID$(ADRDATA$, 23, 14)
        X = InStr(X, CL$, "   ")
     End If
     Mid$(CL$, X, 255) = RFLBL$(35): X = X + RFLTH(35)
     Mid$(CL$, X + 1, 30) = MID$(ADRDATA$, 37, 30)
     X2 = X: X = InStr(X, CL$, "  ")
     If MID$(ADRDATA$, 67, 1) <> " " Then
        If (X > X2) Then Mid$(CL$, X, 1) = ",": X = X + 2
        Mid$(CL$, X, 30) = MID$(ADRDATA$, 67, 30)
     End If

     'Free Form Address for Merged Reports
     Call PutDATA: X = 1
     X = InStr(X, CL$, "  ")
     If EuroAdr Then
        X2 = 129: ZL = 5 - 7 * (AFMT = 1): EuroZIP = 0
        If LTrim$(MID$(ADRDATA$, X2, ZL)) <> "" Then
           Mid$(CL$, X, 1) = ",": X = X + 2: EuroZIP = -1
           Mid$(CL$, X, ZL) = LTrim$(MID$(ADRDATA$, X2, ZL))
           X = InStr(X, CL$, "  ")
        End If
     End If
     X2 = 97
     ZL = 15 - 5 * (AFMT = 1)
     If LTrim$(MID$(ADRDATA$, X2, ZL)) <> "" Then
        If Not EuroZIP Then Mid$(CL$, X, 1) = ",": X = X + 1
        X = X + 1
        Mid$(CL$, X, ZL) = LTrim$(MID$(ADRDATA$, 97, ZL))
        X = InStr(X, CL$, "  ")
     End If
     X2 = 117: ZL = 4 - 8 * (AFMT = 1)
     If LTrim$(MID$(ADRDATA$, X2, ZL)) <> "" Then
        Mid$(CL$, X, 1) = ","
        Mid$(CL$, X + 2, ZL) = LTrim$(MID$(ADRDATA$, X2, ZL))
        X = InStr(X + 2, CL$, "  ")
     End If
     If Not EuroAdr Then
        X2 = 129: ZL = 5 - 7 * (AFMT = 1)
        If LTrim$(MID$(ADRDATA$, X2, ZL)) <> "" Then
           Mid$(CL$, X + 2, ZL) = LTrim$(MID$(ADRDATA$, X2, ZL))
           X = InStr(X + 2, CL$, "  ")
        End If
     End If
     If AFMT = 1 Then
        X2 = 141: ZL = 16
        If LTrim$(MID$(ADRDATA$, X2, ZL)) <> "" Then
           Mid$(CL$, X, 1) = ","
           Mid$(CL$, X + 2, ZL) = MID$(ADRDATA$, X2, ZL)
        End If
     End If
2560 Call PutDATA
     If MID$(IA$, 10, 1) = "1" Then
        GoSub PrintCOM
     End If
     Call PrintLINE
2595 Return

PrintCOM: ' Process COMMENTS
     RI = 0: PAGE.CINDENT = 0
     COMBGN = 0: PAGE.LWIDTH = CWIDTH: CWIDTH = 0
     If PAGE.LWIDTH = 0 Then PAGE.LWIDTH = PAGE.BWIDTH - PAGE.CINDENT
     If ((ComPTR <> 0) Or (COMPTR2 <> 0)) Then
        If ComPTR <> 0 Then Call ComPRINT(ComPTR, COMBGN): ComPTR = 0
        If COMPTR2 <> 0 Then Call ComPRINT(COMPTR2, COMBGN): COMPTR2 = 0
     End If
     Call PrintLINE
2695 Return

2700 ' Print OTHER Info
     If (OA(OT) = 0) Or (MID$(IA$, OT + 4, 1) <> "1") Then GoTo 2795
     Call WriteLINE: If A = 27 Then GoTo 2790
     'MID$(PLINE$, NMOFST, SPW) = LINESEP2$: CALL PrintLINE
     If A = 27 Then GoTo 2790
     Select Case OT
       Case 1: X = 54
       Case 2: X = 49
       Case 3: X = 62
       Case 4: X = 59
     End Select
     Mid$(PLINE$, NMOFST, 255) = RFLBL$(X)
     'CALL PrintLINE: IF A = 27 THEN 2790
     Y = 0
2720 'IF Y > 0 THEN MID$(PLINE$, NMOFST, SPW) = LINESEP3$
     Call WriteLINE: If A = 27 Then GoTo 2790
     ComPTR = 0: xAPTR = 0: LSet OTHDATA$ = String$(Len(OTHDATA$), FC)
     If OA(OT) <> 0 Then
        Call FF3GetRec(OA(OT))
        ComPTR = FF3PNTRS.COM: xAPTR = FF3PNTRS.LOC
        Mid$(OTHDATA$, 1, 11) = XRPTDate$(FF3PNTRS.D1MD, FF3PNTRS.D1Y, DS)
        Mid$(OTHDATA$, 12, 11) = XRPTDate$(FF3PNTRS.D2MD, FF3PNTRS.D2Y, DS)
        Select Case OT
           Case 1: If MID$(FF3HEALTH.HLTHDIAG, 1, 1) <> " " Then Mid$(OTHDATA$, 23, 22) = RTrim$(FF3HEALTH.HLTHDIAG)
                   If MID$(FF3HEALTH.HLTHSTAT, 1, 1) <> " " Then Mid$(OTHDATA$, 45, 8) = RTrim$(FF3HEALTH.HLTHSTAT)
           Case 2: If MID$(FF3EDUC.EDLEV, 1, 1) <> " " Then Mid$(OTHDATA$, 23, 8) = RTrim$(FF3EDUC.EDLEV)
                   If MID$(FF3EDUC.EDDEG, 1, 1) <> " " Then Mid$(OTHDATA$, 31, 4) = RTrim$(FF3EDUC.EDDEG)
                   If MID$(FF3EDUC.EDSUB1, 1, 1) <> " " Then Mid$(OTHDATA$, 35, 9) = RTrim$(FF3EDUC.EDSUB1)
                   If MID$(FF3EDUC.EDSUB2, 1, 1) <> " " Then Mid$(OTHDATA$, 44, 9) = RTrim$(FF3EDUC.EDSUB2)
           Case 3: If MID$(FF3WORK.WORKTP, 1, 1) <> " " Then Mid$(OTHDATA$, 23, 30) = RTrim$(FF3WORK.WORKTP)
           Case 4: If MID$(FF3MIL.MILRANK, 1, 1) <> " " Then Mid$(OTHDATA$, 23, 22) = RTrim$(FF3MIL.MILRANK)
                   If MID$(FF3MIL.MILSTAT, 1, 1) <> " " Then Mid$(OTHDATA$, 45, 8) = RTrim$(FF3MIL.MILSTAT)
        End Select
     End If

     ' Print Information Formatted in OTHDATA$
     PAGE.LOFST = NMOFST + 2: PAGE.COFST = PAGE.LOFST: X = 1
     If OT > 1 Then
        If (MID$(OTHDATA$, 1, 1) <> " ") Then
           Mid$(CL$, 2, 255) = RFLBL$(47): X = 3 + RFLTH(47)
           Mid$(CL$, X, 11) = MID$(OTHDATA$, 1, 11)
           X = InStr(X, CL$, "  ")
        End If
        If (MID$(OTHDATA$, 12, 1) <> " ") Then
           Mid$(CL$, X, 255) = RFLBL$(48): X = X + 1 + RFLTH(48)
           Mid$(CL$, X, 11) = MID$(OTHDATA$, 12, 11)
           X = InStr(X, CL$, "  ")
        End If
     End If
     Select Case OT
        Case 1: ' Medical Information
                If (MID$(OTHDATA$, 23, 1) <> " ") Then
                   Mid$(CL$, X, 255) = RFLBL$(56): X = X + RFLTH(56)
                   Mid$(CL$, X + 1, 22) = MID$(OTHDATA$, 23, 22)
                   X = InStr(X, CL$, "  ")
                End If
                If (MID$(OTHDATA$, 45, 1) <> " ") Then
                   Mid$(CL$, X, 255) = RFLBL$(57): X = X + RFLTH(57)
                   Mid$(CL$, X + 1, 8) = MID$(OTHDATA$, 45, 8)
                   X = InStr(X, CL$, "  ")
                End If
                If (MID$(OTHDATA$, 12, 1) <> " ") Then
                   Mid$(CL$, X, 255) = RFLBL$(58): X = X + RFLTH(58)
                   Mid$(CL$, X + 1, 11) = MID$(OTHDATA$, 12, 11)
                End If
        Case 2: ' Education Information
                Mid$(CL$, X, 255) = RFLBL$(51): X = X + RFLTH(51)
                If (MID$(OTHDATA$, 35, 1) <> " ") Then
                   Mid$(CL$, X + 1, 9) = MID$(OTHDATA$, 35, 9)
                   X = InStr(X, CL$, "  ")
                End If
                If (MID$(OTHDATA$, 44, 1) <> " ") Then
                   Mid$(CL$, X, 255) = RFLBL$(52): X = X + RFLTH(52)
                   Mid$(CL$, X + 1, 9) = MID$(OTHDATA$, 44, 9)
                   X = InStr(X, CL$, "  ")
                End If
                If MID$(OTHDATA$, 23, 1) <> " " Then
                   Mid$(CL$, X, 255) = RFLBL$(50): X = X + RFLTH(50)
                   Mid$(CL$, X + 1, 8) = MID$(OTHDATA$, 23, 8)
                   X = InStr(X, CL$, "  ")
                   Mid$(CL$, X, 255) = RFLBL$(66): X = X + RFLTH(66)
                End If
                If MID$(OTHDATA$, 31, 1) <> " " Then
                   Mid$(CL$, X, 255) = RFLBL$(53): X = X + RFLTH(53)
                   Mid$(CL$, X + 1, 4) = MID$(OTHDATA$, 31, 4)
                End If
        Case 3: ' Occupation Information
                If MID$(OTHDATA$, 23, 1) <> " " Then
                   Mid$(CL$, X, 255) = RFLBL$(63): X = X + RFLTH(63)
                   Mid$(CL$, X + 1, 30) = MID$(OTHDATA$, 23, 30)
                End If
        Case 4: ' Military Information
                If (MID$(OTHDATA$, 23, 22) <> " ") Then
                   Mid$(CL$, X, 255) = RFLBL$(60): X = X + RFLTH(60)
                   Mid$(CL$, X + 1, 22) = MID$(OTHDATA$, 23, 22)
                   X = InStr(X, CL$, "  ")
                End If
                If (MID$(OTHDATA$, 45, 1) <> " ") Then
                   Mid$(CL$, X, 255) = RFLBL$(61): X = X + RFLTH(61)
                   Mid$(CL$, X + 1, 8) = MID$(OTHDATA$, 45, 8)
                End If
     End Select
     Call PutDATA
     Y = Y + 1: If OA(OT) <> 0 Then OA(OT) = FF3PNTRS.NXT
     If MID$(IA$, 10, 1) = "1" Then
        GoSub PrintCOM: If A = 27 Then GoTo 2790
     End If
     If MID$(IA$, 9, 1) = "1" Then
        If (xAPTR <> 0) Then
           AX = 0
           While xAPTR <> 0
              AX = AX + 1: GoSub PrintRESI
              If A = 27 Then xAPTR = 0 Else xAPTR = FF2PFX.NXT
           Wend
           If A = 27 Then GoTo 2790
        End If
     End If
     If OA(OT) <> 0 Then GoTo 2720
     If Y < OD Then GoTo 2720 Else GoTo 2795
2790 OT = 99
2795 Return

End Sub

Rem $STATIC
Sub GetNextMARRIAGE()
   If SPRESEQ Then
      If SPMCNT > 0 Then SPMCNT = SPMCNT - 1
      SC = MLIST(SPMCNT)
      Else
      If RID = FF3SPOUSE.SP1ID Then
         SC = FF3SPOUSE.SP1NXT
         Else
         If RID = FF3SPOUSE.SP2ID Then
            SC = FF3SPOUSE.SP2NXT
            Else
            SC = 0
         End If
      End If
   End If

End Sub

Sub LINEAGE(OPT$)
     On Local Error GoTo 2900
     Select Case OPT$
        Case "BUILD": GoSub 62300
        Case "PRINT": GoSub 62350
     End Select
     On Local Error GoTo 0
     Exit Sub

62300 ' Build LINEAGE Table
      X = RELATE(RID): LinCNT = 0
      If X = 0 Then Exit Sub
      AGL = -(X \ 256): DGL = X And 127: RGL = AGL + DGL
      ExtAdopt = ((X And 128) > 0)
      'CALL RwrkGET("RELREC", RC, RELPTR&(RID))
      If Asc(WFHDR.AMAXLV) > 0 Then
         LinCNT = LinCNT + 1: Y$ = ""
         CS = WFANCREC.PSTAT
         If Not ExtAdoptRule Then CS = CS And 3
         If CS > 0 Then Y$ = MID$("*? *", CS, 1)
         LINTAB$(DGL + 1) = "[" + LTrim$(Str$(WFANCREC.LINEAGE)) + Y$ + "]"
      End If
      If DGL > 0 Then
         If ShowChNUM And (ShowRegNUM Or ShowPNAME) Then P1$ = "."
         If ShowPNAME And (ShowRegNUM) Then P2$ = " "
         xx$ = "("
         For X = 1 To DGL
            If ShowChNUM Then CH$ = LTrim$(ROMAN$(Asc(WFRELREC.CHNUM)))
            CS = (WFRELREC.CSTAT And 3)
            If Not ExtAdoptRule Then CS = CS And 3
            Get #5, WFRELREC.PARPTR + 1, WFRELREC
            If ShowRegNUM Then RN$ = fnx$(Abs(SelSET(Abs(WFRELREC.id))))
            If X > 1 Then xx$ = ","
            If X = DGL Then xz$ = ")"
            If CS > 0 Then Y$ = MID$("*? *", CS, 1) Else Y$ = ""
            If ShowPNAME Then
               Call FF1GetRec(Abs(WFRELREC.id))
               L = InStr(FF1Rec.GIVEN, " ") - 1
               If L > 0 Then N$ = Left$(FF1Rec.GIVEN, L) Else N$ = ""
            End If
            LINTAB$(X) = xx$ + Y$ + CH$ + P1$ + RN$ + P2$ + N$ + xz$
         Next X
         LinCNT = LinCNT + DGL
      End If
      If FF1Rec.RID <> RID Then Call FF1GetRec(RID)
      Return

62350 ' Format Lineage
      If LinCNT > 0 Then
         X = RELATE(RID)
         AGL = -(X \ 256): DGL = X And 127: RGL = AGL + DGL
         If ShowREL Then
            z = PAGE.COFST
            Mid$(CL$, z, 255) = RFLBL$(64)
            z = z + RFLTH(64) + 1
            ExtAdopt = ((X And 128) > 0)
            If ExtAdopt Then Mid$(CL$, z, 1) = "*": z = z + 1
            If FF1Rec.RID <> RID Then Call FF1GetRec(RID)
            Mid$(CL$, z, 12) = Relation$("M", FF1Rec.SEX, AGL, RGL, RRELCD$())
            z = InStr(z, CL$, "  ") + 2
            Mid$(CL$, z, 255) = RFLBL$(65): z = z + RFLTH(65)
            Else
            Call PutDATA: z = 2
         End If
         For X = 1 To LinCNT
             Mid$(CL$, z, 255) = LINTAB$(X): Call PutDATA: z = 1
             If A = 27 Then X = LinCNT
         Next X
         If ShowREL Then Call PrintLINE
      End If
      Return

2900 ' Error Routine
     SysVAR.NVAR1 = ERL
     Call ERRMessage("XXXXM098", ERR, ERL)
     Resume 2910
2910 ' END ON ERROR
     Exit Sub

End Sub

Function ROMAN$(A)
   XU = A Mod 10
   If RomanCAPS Then
      Select Case XU
         Case 1: U$ = "I"
         Case 2: U$ = "II"
         Case 3: U$ = "III"
         Case 4: U$ = "IV"
         Case 5: U$ = "V"
         Case 6: U$ = "VI"
         Case 7: U$ = "VII"
         Case 8: U$ = "VIII"
         Case 9: U$ = "IX"
      End Select
      Else
      Select Case XU
         Case 1: U$ = "i"
         Case 2: U$ = "ii"
         Case 3: U$ = "iii"
         Case 4: U$ = "iv"
         Case 5: U$ = "v"
         Case 6: U$ = "vi"
         Case 7: U$ = "vii"
         Case 8: U$ = "viii"
         Case 9: U$ = "iX"
      End Select
   End If
   XT = ((A - XU) Mod 100) / 10
   Select Case XT
      Case 1: T$ = "X"
      Case 2: T$ = "XX"
      Case 3: T$ = "XXX"
      Case 4: T$ = "XL"
      Case 5: T$ = "L"
      Case 6: T$ = "LX"
      Case 7: T$ = "LXX"
      Case 8: T$ = "XXC"
      Case 9: T$ = "XC"
   End Select
   If A > 100 Then Y$ = "C"
   xx$ = Space$(10): RSet xx$ = Y$ + T$ + U$
   ROMAN$ = Right$(xx$, 7)
End Function

