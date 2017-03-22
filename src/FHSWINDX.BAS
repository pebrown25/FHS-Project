Attribute VB_Name = "modFHSWIndx"
DefInt A-Z
Dim SST1$

  Private Function FNBH$(a)
  FNBH$ = Chr$((Abs(a) \ 256) And 255)
  End Function
  Private Function FNBL$(a)
  FNBL$ = Chr$(Abs(a) And 255)
  End Function

Public Sub UtilINDX(OPT$)
1  Rem $INCLUDE: 'FHSCOMON.BAS'
   PN$ = " FAMILY HISTORY SYSTEM - FHSWINDX - Index File Generation Program"
   CY$ = " (C) Copyright 1985-2008 by Phillip E. Brown"
   SST1$ = ""
   
40 ' Format Display
   Call PgmPREP("FHSINDX")
   Load frmFHSWUPgm
   RptCODE = "INDX"
   frmFHSWUPgm.UPSetup.AddItem "INDX"
   frmFHSWUPgm.UPSetup.TEXT = RptCODE
   frmFHSWUPgm.UPSetup_Click
   frmFHSWUPgm.Show
   Exit Sub


End Sub



Private Function DateDATA$(MD, CY, ASCEND)
         MD = Abs(MD): CY = Abs(CY)
         If (MD = 0) And (CY = 0) And ASCEND Then
            mm = 255: DD = 255: CC = 255: yy = 255
            Else
            mm = (MD \ 100): If mm > 255 Then mm = 255
            DD = MD - mm * 100: If DD > 255 Then DD = 255
            CC = (CY \ 100): If CC > 255 Then CC = 255
            yy = CY - 100 * CC: If yy > 255 Then yy = 255
         End If
         X$ = Space$(4): X = 1
         Select Case ((ENV.SortOPTS \ 4) And 7)
            Case 0: Mid$(X$, X, 1) = Chr$(CC)
                    Mid$(X$, X + 1, 1) = Chr$(yy)
                    Mid$(X$, X + 2, 1) = Chr$(mm)
                    Mid$(X$, X + 3, 1) = Chr$(DD)
            Case 1: Mid$(X$, X, 1) = Chr$(mm)
                    Mid$(X$, X + 1, 1) = Chr$(DD)
                    Mid$(X$, X + 2, 1) = Chr$(CC)
                    Mid$(X$, X + 3, 1) = Chr$(yy)
            Case 2: Mid$(X$, X, 1) = Chr$(DD)
                    Mid$(X$, X + 1, 1) = Chr$(mm)
                    Mid$(X$, X + 2, 1) = Chr$(CC)
                    Mid$(X$, X + 3, 1) = Chr$(yy)
            Case 3: Mid$(X$, X, 1) = Chr$(CC)
                    Mid$(X$, X + 1, 1) = Chr$(yy)
                    Mid$(X$, X + 2, 1) = Chr$(mm)
            Case 4: Mid$(X$, X, 1) = Chr$(mm)
                    Mid$(X$, X + 1, 1) = Chr$(DD)
            Case 5: Mid$(X$, X, 1) = Chr$(CC)
                    Mid$(X$, X + 1, 1) = Chr$(yy)
            Case 6: Mid$(X$, X, 1) = Chr$(mm)
            Case 7: Mid$(X$, X, 1) = Chr$(DD)
         End Select
         DateDATA$ = X$
End Function

Private Sub INDEXBuild()
'     SHARED FMT001, FMT030, FMT031, ShowSTATS, SST1$, SFLD()
1000 ' Create Index File
     Call FAMOpen("READ"): If Not OKAY Then GoTo 1090 ' Open Name file
     If FF1Hdr.MAXID = 0 Then
        Call PutMSG("INDXM001")
        R$ = Chr$(27): GoTo 1090
     End If
     TextSORT = ((SFLD(1) \ 256) > 0) Or ((SFLD(2) \ 256) > 0) Or ((SFLD(6) \ 256) > 0) Or ((SFLD(7) \ 256) > 0)
     If TextSORT Then
        Call PutMSG("INDXM002")
        If a = 27 Then R$ = a$: GoTo 1090
        ENV.HUSBSURNM = (a$ = SNGLKEY$(1))
        Call PutMSG("INDXM010")
        If a = 27 Then R$ = a$: GoTo 1090
        TextSORT = (Val(a$) - 1)
        If (TextSORT < 0) Or (TextSORT > 3) Then
           TextSORT = ENV.SortOPTS And 3
        End If
        Call PutMSG("INDXM012")
        If a = 27 Then R$ = a$: GoTo 1090
        UseSEQTable = (a$ = SNGLKEY$(1))
        ENV.SortOPTS = (ENV.SortOPTS And 220) + TextSORT - 32 * UseSEQTable
        Call ShowSortOPTS
     End If
     DateSORT = ((SFLD(3) \ 256) > 0) Or ((SFLD(4) \ 256) > 0) Or ((SFLD(5) \ 256) > 0)
     If DateSORT Then
        Call PutMSG("INDXM011")
        If a = 27 Then R$ = a$: GoTo 1090
        DateSORT = (Val(a$) - 1)
        If (DateSORT < 0) Or (DateSORT > 7) Then
           DateSORT = (ENV.SortOPTS \ 4) And 7
        End If
        ENV.SortOPTS = (ENV.SortOPTS And 35) + 4 * DateSORT
        Call ShowSortOPTS
     End If
     ENV.SortSTAT = -1: ShowStats = -1
     GoSub 1800                      ' Allocate NDXID(),SEQDATA$()
     If Not OKAY Then GoTo 1090
     'CURFMT = FMT030: FMT = FMTTAB(CURFMT)
     'Call WinSWITCH(2): Call WinCLR
     'Call FmtDREST
     'Call WinFORMAT(FMT030)
     'X = FMT.BGNFLD
     'While X > 0
     '   FFLD = FLDTAB(X)
     '   If FFLD.UTXT > 0 Then
     '      X$ = UTXT$(FFLD.UTXT)
     '      Select Case FFLD.NAME
     '         Case "IDCT": LSet X$ = Str$(FF1Hdr.MAXID)
     '         Case "TSZB": LSet X$ = Str$(SN1 * DL#)
     '         Case "MSZB": LSet X$ = Str$(SN2 * DL#)
     '         Case "WSZB": LSet X$ = Str$(SWSZ#)
     '         Case "SSEG": RSet X$ = LTrim$(Str$(N4))
     '         Case "SSCT": LSet X$ = Str$(SN2)
     '         Case "LSCT": LSet X$ = Str$(N3)
     '      End Select
     '      LSet UTXT$(FFLD.UTXT) = X$
     '   End If
     '   X = FLDCHN(X).FWD
     'Wend
     'Call WinDisplayDATA(0)
     'SortFMT = FMT031
     'CURFMT = SortFMT: FMT = FMTTAB(CURFMT)
     'Call FmtDREST
     'Call FmtFindFLD("TIM1")
     'If FOUND Then LSet UTXT$(FFLD.UTXT) = Time$
     frmFHSWUPgm.spstat(0) = Str$(FF1Hdr.MAXID)
     frmFHSWUPgm.spstat(1) = Str$(SN1 * dl#)
     frmFHSWUPgm.spstat(2) = Str$(SN2 * dl#)
     frmFHSWUPgm.spstat(3) = Str$(swsz#)
     frmFHSWUPgm.spstat(4) = LTrim$(Str$(n4))
     frmFHSWUPgm.spstat(5) = Str$(SN2)
     frmFHSWUPgm.spstat(6) = Str$(n3)
     frmFHSWUPgm.spstat(20) = Time$
     'Call WinFORMAT(0)
     GoSub 1850             ' Get DATA, Sort, and Create .NDX File
     If Not OKAY Then GoTo 1090
     Call PutMSG("INDXM003")
     R$ = ""

1090 ' Return
     On Local Error GoTo 0
     If swsz# > 0 Then GoSub 5630
     Close: ENV.FFOPEN = 0: OKAY = (R$ <> Chr$(27))

     Exit Sub

1800 ' Initialize Arrays to be used in SORT
     SN1 = FF1Hdr.MAXID: SN2 = SN1
     SDL = 0: BDSORT = 0: BPSORT = 0
     ReDim SFL(SFCNT)
     Call SFLD2SX
     For NX = 1 To SFCNT
         X = SX(NX) \ 256: L = SX(NX) And 127
         If (X = 3) Or (X = 5) Then BDSORT = -1
         If (X = 6) Or (X = 7) Then BPSORT = -1
         If (X < 3) Or ((X > 5) And (X < 8)) Then
            Select Case TextSORT
               Case 2: L = 4
               Case 3: L = 4 + L
            End Select
            Else
            If (X > 2) And (X < 6) Then
               If (DateSORT > 5) Then
                  L = 1
                  Else
                  If (DateSORT = 5) Then
                     L = 2
                     Else
                     If (DateSORT > 2) Then L = 3
                  End If
               End If
            End If
         End If
         SFL(NX) = L: SDL = SDL + L
     Next NX
     SDATA$ = String$(SDL, 0)

1810 ' Build Sort Arrays
     On Local Error GoTo 1849
     ReDim NDXID(SN1), SEQDATA$(1)
     Z# = FRE(SEQDATA$(1)) - 100
     dl# = SDL: SN2# = Z# / (dl# + 10)
     If SN2# > 32767 Then SN2 = SN1 Else SN2 = SN2#: If SN2 > SN1 Then SN2 = SN1
     'X# = Z# - (DL# + 4) * SN2
     'IF X# < 5000 THEN X# = (Z# - 5000) / (DL# + 4): SN2 = X#
1820 ReDim SEQDATA$(SN2)
     Call INDEXSort("SETUP", NDXID(), SEQDATA$())
     If Not SortOK Then Return
     If SN1 = SN2 Then GoTo 1825
     swsz# = dl# * (SN1 - SN2): SysVAR.MVAR = LTrim$(Str$(swsz#))
     If ENV.SortSTAT Then Call SortStatACCUM
     SWD$ = WRKFPFX$: GoTo 1822
1821 Call PutMSG("INDXM004")
     R$ = a$: If Asc(R$) = 27 Then GoTo 1840
     If a$ < "A" Or a$ > "Z" Then SOUND BP, DUR: GoTo 1821
     SWD$ = a$ + ":"
1822 GoSub 5600: If Not OKAY Then GoTo 1821
     On Local Error GoTo 1845
     If ENV.SortSTAT Then SortSTAT!(0) = Timer
1825 For ix = 1 To SN2
        SEQDATA$(ix) = SDATA$: HIINDEX = ix
        R$ = INKEY$
        If R$ <> "" Then
           If R$ = Chr$(27) Then
              ix = SN2
              Else
              If ENV.SortSTAT Then
                 ShowStats = (Asc(Right$(R$, 1)) = 81)
                 Call SortStatSHOW
              End If
           End If
        End If
1829 Next ix
     If (R$ = Chr$(27)) Or (HIINDEX = SN2) Then GoTo 1840
     On Local Error GoTo 0
     GoTo 1820

1840 On Local Error GoTo 0
     MSZ# = SN2 * dl#: swsz# = (SN1 - SN2) * dl#
     n3 = SN1 Mod SN2: n4 = (SN1 \ SN2)
     OKAY = OKAY And (R$ <> Chr$(27))
     Return

1845 ' Return After Error
     Erase SEQDATA$: SysVAR.NVAR1 = HIINDEX
     ix = SN2: SN2 = HIINDEX - 30: OKAY = 0
     If SN2 < 0 Then
        Call ERRMessage("INDXM007", ERR, ERL)
        Resume 1849
     End If
     Resume 1829

1849 ' Error Allocating Arrays
     Call ERRMessage("INDXM006", ERR, ERL)
     OKAY = 0
     Resume 1840

1850 ' Put Name file sequence data in SEQDATA$() and $ORTWORK (if necessary)
     If BDSORT Or BPSORT Then
        BD$ = Space$(11): DD$ = BD$
        If BPSORT Then BPL$ = Space$(41) Else BPL$ = ""
        DPL$ = BPL$: AGE$ = ""
        DFM = Val(CONFIG.DFM)
     End If
     JX = 0: SortOK = -1: OKAY = -1
     SNDX$ = Space$(4): xDATA$ = Space$(34)
     For XIN = 1 To SN1
         NDXID(XIN) = XIN
         Call FF1GetRec(XIN): OFST = 0: SPREC = FF1Rec.SPOUSE
         If ENV.SortSTAT Then SortSTAT!(6) = SortSTAT!(6) + 1
         If BDSORT Or BPSORT Then
            Call GetBDEvents(BD$, BPL$, DD$, DPL$, AGE$)
            BMD = DATEMD(BD$, DFM): byr = DATEY(BD$, DFM)
            DMD = DATEMD(DD$, DFM): DYR = DATEY(DD$, DFM)
         End If
         For NX = 1 To SFCNT
             L = SFL(NX): ASCEND = ((SX(NX) And 128) = 0)
             Select Case SX(NX) \ 256
                Case 1: LSet xDATA$ = LTrim$(RTrim$(FF1Rec.SURNM))
                        If ENV.HUSBSURNM And (FF1Rec.SEX = FGENDR$(2)) And (FF1Rec.SPOUSE <> 0) And ((Asc(FF1Rec.STATUS2) And 128) = 0) Then
                           LSet xDATA$ = FMTNAME$(1): LSet xDATA$ = SURNM$
                        End If
                        Mid$(SDATA$, OFST + 1, L) = TextDATA$(xDATA$, L, ASCEND)
                Case 2: Mid$(SDATA$, OFST + 1, L) = TextDATA$(FF1Rec.GIVEN, L, ASCEND)
                Case 3: Mid$(SDATA$, OFST + 1, L) = DateDATA$(BMD, byr, ASCEND)
                Case 4: MMD = 0: MYR = 0
                        If FF1Rec.SPOUSE <> 0 Then
                           Call FF3GetRec(FF1Rec.SPOUSE)
                           If (F3BY$ <> MKI$(0)) Or (F3BMD$ <> MKI$(0)) Then
                              MMD = FF3SPOUSE.D1MD: MYR = FF3SPOUSE.D1Y
                           End If
                        End If
                        Mid$(SDATA$, OFST + 1, L) = DateDATA$(MMD, MYR, ASCEND)
                Case 5: Mid$(SDATA$, OFST + 1, L) = DateDATA$(DMD, DYR, ASCEND)
                Case 6: Mid$(SDATA$, OFST + 1, L) = TextDATA$(BPL$, L, ASCEND)
                Case 7: Mid$(SDATA$, OFST + 1, L) = TextDATA$(BPL$, L, ASCEND)
                Case 8: Mid$(SDATA$, OFST + 1, 1) = FF1Rec.SEX
                Case 9: Mid$(SDATA$, OFST + 1, 1) = FNBH$(XIN)
                        Mid$(SDATA$, OFST + 2, 1) = FNBL$(XIN)
             End Select
             If ((SX(NX) And 128) <> 0) Then
                For N = 1 To L
                    Mid$(SDATA$, OFST + N, 1) = Chr$(Asc(MID$(SDATA$, OFST + N, 1)) Xor 255)
                Next N
             End If
             OFST = OFST + L
         Next NX
         JX = JX + 1: LSet SEQDATA$(JX) = SDATA$: LSet SDATA$ = ""
         If (XIN = n3) Or (JX = SN2) Then
            SS = (XIN - n3) \ SN2
            If ENV.SortSTAT And ShowStats Then Call SortStatSHOW
            Call INDEXSort("SEGMENT", NDXID(), SEQDATA$())
            If SortOK Then
               If (SN1 <> SN2) And (SS < n4) Then
                  If ENV.SortSTAT Then
                     Call SortStatACCUM: HP3 = SortPHASE: SortPHASE = 3
                  End If
                  For X = 1 To JX: Put #5, , SEQDATA$(X): Next X
                  If ENV.SortSTAT Then
                     SortSTAT!(7) = SortSTAT!(7) + JX - 1
                     Call SortStatACCUM: SortPHASE = HP3
                  End If
               End If
               JX = 0: If SS < n4 Then SS = SS + 1
               Else
               XIN = SN1
            End If
         End If
         R$ = INKEY$
         If R$ <> "" Then
            SortOK = (R$ <> Chr$(27))
            If Not SortOK Then
               XIN = SN1
               Else
               If ENV.SortSTAT Then
                  ShowStats = (Asc(Right$(R$, 1)) = 81)
                  Call SortStatSHOW
               End If
            End If
         End If
         If ShowStats And (XIN Mod 100 = 0) Then Call SortStatSHOW
1876 Next XIN
     If ENV.SortSTAT Then Call SortStatSHOW: SortPHASE = 0
     OKAY = OKAY And SortOK
     If OKAY Then
        If (SN1 <> SN2) Then
           Call INDEXSort("MERGE", NDXID(), SEQDATA$())
           OKAY = SortOK
        End If
        If OKAY Then Call INDEXFile("CREATE")   ' Write Index Array to File
     End If
     OKAY = OKAY And SortOK
     Call INDEXSort("CLEANUP", NDXID(), SEQDATA$())
     Erase NDXID, SEQDATA$
     Return

5600 ' Open $ORTWORK File
     OKAY = 0: SWN$ = SWD$ + "$ORTWORK"
     GoSub 5630
     On Local Error GoTo 5610
     Open SWN$ For Binary Access Read Write As #5
     GoTo 5611
5610 SysVAR.FILENAME = SWN$
     Call PutMSG("XXXXM004")
     R$ = Chr$(27)
     Resume 5620
5611 On Local Error GoTo 5615
     L# = (SN1 - SN2 - 1) * dl# + 1
     Put #5, L#, SDATA$
     Seek #5, 1
     OKAY = -1: GoTo 5620
5615 Call PutMSG("INDXM008")
     R$ = Chr$(27): Resume 5616
5616 GoSub 5630
5620 On Local Error GoTo 0
     Return

5630 ' Kill $ORTWORK File
     On Local Error GoTo 5634
     Close #5: GoTo 5635
5634 Resume 5635
5635 On Local Error GoTo 5637
     Kill SWN$: GoTo 5638
5637 Resume 5638
5638 On Error GoTo 0
     Return
End Sub

Private Sub INDEXFields(OPT$)
'    SHARED FMT020, FMT021, FMT029, SFLD()
    Select Case OPT$
       Case "SFLD2SX": GoSub 550
       Case "SX2SFLD": GoSub 580
    End Select
    Exit Sub

550 ' Transfer SFLD() to SX()
    For NX = 0 To 9: SX(NX) = 0: Next NX
    For ix = 1 To 8: NX = SFLD(ix) \ 256
        If SX(NX) = 0 Then SX(NX) = ix * 256 + (SFLD(ix) And 255)
    Next ix: SFCNT = 0
    For NX = 1 To 8
        If ((SX(NX) > 0) And (SFCNT < 9)) Then
           SFCNT = SFCNT + 1
           If NX > SFCNT Then SX(SFCNT) = SX(NX): SX(NX) = 0
        End If
    Next NX
    SFCNT = SFCNT + 1
    SX(SFCNT) = 9 * 256 + 2
    Return

580 ' Transfer SX() data to SFLD() array
    For ix = 1 To 9: SFLD(ix) = SFLD(ix) And 255: Next ix
    For NX = 1 To SFCNT
       ix = (SX(NX) \ 256)
       SFLD(ix) = (256 * NX) + (SX(NX) And 128) + (SFLD(ix) And 127)
       If (ix < 3) Or ((ix > 5) And (ix < 8)) Then
          SFLD(ix) = SFLD(ix) - (SFLD(ix) And 127) + (SX(NX) And 127)
       End If
    Next NX
    Return

End Sub


Private Sub INDEXSort(OPT$, AID(), SEQDATA$())
'   SHARED ShowSTATS
    ' OPT$ = "ALL"      Complete Sort of AID()
    '      = "SETUP"    Allocate Required Arrays
    '      = "SEGMENT"  Sort SEQDATA$() <--> "SS" Segment of AID()
    '      = "MERGE"    Merge Sort Segments and Create Index File
    '      = "CLEANUP"  Erase Arrays
    ' AID() is the Array of ID's that is to be placed in Sorted Sequence
    ' SEQDATA$() is Array where DATA is kept in Memory
    ' SN1 = Number of Entries in AID()
    ' SN2 = Number of Entries in SEQDATA$()
    ' SDL = Length of Each Entry in SEQDATA$()
    ' If SN1<>SN2 then using $ORTWORK file as File #5
    ' SS = # Sort Segment (OPT$="SEGMENT")

1200 ' Use Segmented SHELL Sort/Merge for Sequencing AID()
     On Local Error GoTo 1207
     R$ = "": SortOK = -1
     SORTDATA$ = Space$(SDL)
     dl& = SDL

     ' Compute N3=Size of XTRA Segment, N4=#Sort Segments
     SB = 0: n3 = SN1 Mod SN2: n4 = SN1 \ SN2

     Select Case OPT$
        Case "SEGMENT": If SS = 0 Then NS = n3 Else NS = SN2: SB = n3 + (SS - 1) * SN2
                        GoSub 1250: GoTo 1209
        Case "MERGE": GoSub 1210: GoTo 1209
        Case "CLEANUP": GoTo 1208
     End Select

     If ENV.SortSTAT Then
        ReDim SortSTAT!(10)
        SortPHASE = 1: SortSTAT!(0) = Timer
     End If
     If SN1 <> SN2 Then
        ReDim MRGID(SN1), MRGBX(n4 + 1), MRGIX(n4 + 1), MRGSS(n4 + 1), MRGSD$(n4 + 1)
        For X = 0 To n4: MRGSD$(X) = SORTDATA$: Next X
     End If
     If OPT$ = "SETUP" Then GoTo 1209

     If n3 > 0 Then
        NS = n3: GoSub 1250: If Not SortOK Then GoTo 1209
        If OPT$ = "ALL" Then GoSub 1240: If Not SortOK Then GoTo 1209
     End If
     SS = 1: NS = SN2

1205 ' Perform SORT
     If SS <= n4 Then
        GoSub 1250: If Not SortOK Then SS = n4
     End If
     If (Not SortOK) Or (SN1 = SN2) Then GoTo 1209
     If SS < n4 Then
        GoSub 1240: If Not SortOK Then GoTo 1209
        SS = SS + 1: GoTo 1205
     End If
     GoSub 1210: GoTo 1208

1207 ' Error Routine
     XERR = ERR: XERL = ERL
     ReDim MRGSD$(1)
     Call ERRMessage("XXXXM071", XERR, XERL)
     SortOK = 0
     Resume 1208

1208 ' Cleanup at END of Sort
     If SN1 <> SN2 Then Erase MRGID, MRGBX, MRGIX, MRGSS, MRGSD$
     SortPHASE = 0: ENV.SortSTAT = 0: SortFMT = 0

1209 ' Exit from SORT Routine
     On Local Error GoTo 0
     Exit Sub

1210 ' Merge SORTed Segments
     If ENV.SortSTAT Then
        Call SortStatACCUM: HP4 = SortPHASE: SortPHASE = 4
        SortSTAT!(10) = 0: Call SortStatSHOW
     End If
     ' MRGID()=ID'S in Sorted Segment Sequence
     ' MRGSD$(SS)=SORTDATA Entry for Segment "SS"
     '    These are Sorted by rearranging MRGSS()
     ' MRGSS(..)=Segment Containing MRGSD$(..)
     ' MRGIX(SS)=Sequence# of Next Entry in Segment "SS"
     ' MRGBX(SS)=Sequence # of Last MRGID() Entry in Segment "SS"
     MRGSS(0) = 0:  MRGIX(0) = 1: MRGBX(0) = n3: SSL = n4
     For ix = 1 To SN1: MRGID(ix) = AID(ix): Next ix
     For Z = 1 To n4
         MRGSS(Z) = Z
         MRGIX(Z) = MRGBX(Z - 1) + 1
         MRGBX(Z) = MRGBX(Z - 1) + SN2
     Next Z
     HiDATA$ = String$(SDL, 255): YDATA$ = HiDATA$
     For MS = 0 To n4
         MRGSD$(MS) = HiDATA$: GoSub 1235: TEOS = TEOS + EOS
         If Not SortOK Then MS = n4
     Next MS
     If Not SortOK Then Return
     GoSub 1230: SSL = SSL + TEOS
     For ix = 1 To SN1
         GoSub 1225
         If Not SortOK Then
            ix = SN1
            Else
            AID(ix) = SX
            If ENV.SortSTAT Then
               SortSTAT!(10) = SortSTAT!(10) + 1
               If ShowStats And (ix Mod 100 = 1) Then Call SortStatSHOW
            End If
         End If
     Next ix
     Erase MRGID, MRGBX, MRGIX, MRGSS, MRGSD$
     If ENV.SortSTAT Then Call SortStatSHOW: SortPHASE = HP4
     Return

1225 ' Find Next Index
     GoSub 1280
     If Not SortOK Then
        ix = SN1
        Else
        MS = MRGSS(0): X = MRGIX(MS): SX = MRGID(X)
        MRGIX(MS) = MRGIX(MS) + 1
        GoSub 1235: GoSub 1230: SSL = SSL + EOS
     End If
     Return

1230 ' Bubble Sort MRGSD$() into order by MRGSS()
     JX = 0
     While JX < SSL
        KX = JX + 1: Y = MRGSS(KX): LSet YDATA$ = MRGSD$(Y)
        While KX > 0
           KX = KX - 1: X = MRGSS(KX)
           If MRGSD$(X) > YDATA$ Then
              SWAP MRGSS(KX), MRGSS(KX + 1)
              Else
              KX = 0
           End If
        Wend
        JX = JX + 1
     Wend
     Return

1235 ' Get Next Entry in Segment MS
     X = MRGIX(MS): EOS = (X > MRGBX(MS))
     If EOS Then
        LSet MRGSD$(MS) = HiDATA$
        Else
        If MS < n4 Then
           If ENV.SortSTAT Then
              Call SortStatACCUM: HP3 = SortPHASE: SortPHASE = 3
           End If
           RBA& = (MRGID(X) - 1) * dl& + 1
           Get #5, RBA&, SORTDATA$
           LSet MRGSD$(MS) = SORTDATA$
           If ENV.SortSTAT Then
              SortSTAT!(9) = SortSTAT!(9) + 1
              Call SortStatACCUM: SortPHASE = HP3
           End If
           Else
           X = MRGID(X) - MRGBX(MS - 1)
           LSet MRGSD$(MS) = SEQDATA$(X)
        End If
     End If
     Return

1240 ' Exchange SEQDATA$() for $ORTWORK data
     If ENV.SortSTAT Then
        Call SortStatACCUM: HP3 = SortPHASE: SortPHASE = 3
     End If
     For ix = SB + 1 To SB + NS
         RBA& = (ix - 1) * dl& + 1
         Get #5, RBA&, SORTDATA$
         SWAP SORTDATA$, SEQDATA$(ix - SB)
         Put #5, RBA&, SORTDATA$
         If ENV.SortSTAT Then SortSTAT!(8) = SortSTAT!(8) + 2
         GoSub 1280: If Not SortOK Then ix = SB + NS
     Next ix
     SB = SB + NS
     If ENV.SortSTAT Then Call SortStatACCUM: SortPHASE = HP3
     Return

1250 ' Sort Segment of File
     If ENV.SortSTAT Then
        Call SortStatACCUM: HP2 = SortPHASE: SortPHASE = 2
     End If
     For ix = SB + 1 To SB + NS: AID(ix) = AID(ix) - SB: Next ix
     D = 2 ^ (Int(Log(NS + 1) / Log(2))) - 1
1254 For ix = SB + 1 To SB + D
         For JX = ix To SB + NS - D Step D
             X = AID(JX): Y = AID(JX + D)
             GoSub 1280
             If Not SortOK Then
                ix = SB + D: JX = SB + NS - D: D = 0: GoTo 1270
             End If
             If Not (SEQDATA$(X) > SEQDATA$(Y)) Then GoTo 1270
             Z = Y: KX = JX: AID(KX + D) = X
1261         KX = KX - D
             If KX >= ix Then
                Y = AID(KX)
                GoSub 1280
                If Not SortOK Then
                   ix = SB + D: JX = SB + NS - D: D = 0: GoTo 1270
                End If
                If SEQDATA$(Z) < SEQDATA$(Y) Then AID(KX + D) = Y: GoTo 1261
             End If
             AID(KX + D) = Z
1270     Next JX
     Next ix
     If ENV.SortSTAT And ShowStats Then Call SortStatSHOW
     If SortOK Then
        D = D \ 2: If D > 0 Then GoTo 1254
        For ix = SB + 1 To SB + NS: AID(ix) = AID(ix) + SB: Next ix
        If ENV.SortSTAT Then Call SortStatSHOW: SortPHASE = HP2
     End If
     Return

1280 ' Check for User Request
     R$ = INKEY$
     If R$ <> "" Then
        SortOK = (R$ <> Chr$(27))
        If SortOK And ENV.SortSTAT Then
           ShowStats = (Asc(Right$(R$, 1)) = 81)
           Call SortStatSHOW
        End If
     End If
     Return

End Sub
Sub SortStatACCUM()
    If SortPHASE <> 0 Then
       SortSTAT!(SortPHASE) = SortSTAT!(SortPHASE) + Timer - SortSTAT!(0)
    End If
    SortSTAT!(0) = Timer
End Sub


Private Sub SortSEQ(OPT$, SST1$, X$, L%)
    Select Case OPT$
       Case "BUILD": 'Create Sort Sequence Table
                      Call TableLOAD("SORTSEQ", T, 0)
                      SST1$ = String$(255, 0)
                      X = TABLE(T).FT
                      While X > 0
                         X$ = TABDATA$(X)
                         If X$ <> "" Then
                            Y = Asc(X$): Z = Val(MID$(X$, 2, 3))
                            If (-1 < Z) And (Z < 256) Then
                               Mid$(SST1$, Y, 1) = Chr$(Z)
                            End If
                         End If
                         X = TABDATACHN(X).FWD
                      Wend
                      Call TableFREE("SORTSEQ")
       Case "TRAN": 'Translate X$ using SST1$
                     X = Len(X$): If L < X Then X = L
                     Y = 0: w = 0
                     While Y < X
                        Y = Y + 1: Z = Asc(MID$(X$, Y, 1))
                        If Z > 0 Then
                           Z$ = MID$(SST1$, Z, 1)
                           If Asc(Z$) > 0 Then
                              w = w + 1: Mid$(X$, w, 1) = Z$
                           End If
                        End If
                     Wend
                     While w < X
                        w = w + 1: Mid$(X$, w, 1) = Chr$(0)
                     Wend
    End Select

End Sub

Sub SortStatSHOW()
    Call SortStatACCUM
    'If SortFMT <> 0 Then
    '   If (CurFMT <> SortFMT) Then CurFMT = SortFMT: Fmt = FMTTAB(SortFMT)
    '   X = Fmt.BGNFLD
    '   While X > 0
    '      FFLD = FLDTAB(X)
    '      If FFLD.UTXT > 0 Then
    '         xx$ = UTXT$(FFLD.UTXT): XL = Len(xx$)
    '         Select Case FFLD.NAME
    '            Case "SS  ": RSet xx$ = LTrim$(Str$(SS))
    '            Case "SEC1": RSet xx$ = MID$(Str$(SortSTAT!(1)), 2, XL)
    '            Case "SEC2": RSet xx$ = MID$(Str$(SortSTAT!(2)), 2, XL)
    '            Case "SEC3": RSet xx$ = MID$(Str$(SortSTAT!(3)), 2, XL)
    '            Case "SEC4": RSet xx$ = MID$(Str$(SortSTAT!(4)), 2, XL)
    '            Case "SEC5": RSet xx$ = MID$(Str$(SortSTAT!(5)), 2, XL)
    '            Case "CNT1": RSet xx$ = LTrim$(Str$(SortSTAT!(6)))
    '            Case "CNT2": RSet xx$ = LTrim$(Str$(SortSTAT!(7)))
    '            Case "CNT3": RSet xx$ = LTrim$(Str$(SortSTAT!(8)))
    '            Case "CNT4": RSet xx$ = LTrim$(Str$(SortSTAT!(9)))
    '            Case "CNTW": RSet xx$ = LTrim$(Str$(SortSTAT!(7) + SortSTAT!(8) + SortSTAT!(9)))
    '            Case "CNT5": RSet xx$ = LTrim$(Str$(SortSTAT!(10)))
    '            Case "TIM2": LSet xx$ = Time$
    '         End Select
    '         LSet UTXT$(FFLD.UTXT) = xx$
    '      End If
    '      X = FLDCHN(X).FWD
    '   Wend
       'Call WinDisplayDATA(0)
       SortSTAT!(0) = Timer
    'End If
End Sub

Private Function TextDATA$(X$, L%, ASCEND)
' SHARED SST1$
999   If ((LTrim$(X$) = "") And ASCEND) Then
      TextDATA$ = Chr$(255)
      Else
      Select Case (ENV.SortOPTS And 3)
         Case 0: Y$ = X$
         Case 1: Y$ = UCX$(X$)
         Case 2: TD$ = Space$(4): Call ComputeSNDX(X$, TD$)
                 Y$ = TD$
         Case 3: TD$ = Space$(4): Call ComputeSNDX(X$, TD$)
                 Y$ = TD$ + X$
      End Select
      If (ENV.SortOPTS And 32) > 0 Then
         Call SortSEQ("TRAN", SST1$, Y$, L)
      End If
      TextDATA$ = Y$
   End If
End Function

