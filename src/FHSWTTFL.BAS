Attribute VB_Name = "modFHSWTTfl"
DefInt A-Z

Dim NDX()
Dim ARID(), DRID(), RCNT(), LVL(), AID(), SNDX()
Dim SLIST$(), SNIX(), HLIT$()
Dim ICHG, NMCNT, SURNAME$, CURSURNM$, SNL, TSK, ILVL$
Public Sub UtilTTFL(OPT$)
1  'Rem $INCLUDE: 'FHSCOMON.BAS'
   PN$ = " FAMILY HISTORY SYSTEM - FHSTTFL - Tiny Tafel EXPORT Program"
   CY$ = " (C) Copyright 1991,1992,1993,1996,2008 by Phillip E. Brown"
   'DEF fnx$(a) = Right$(Str$(a), Len(Str$(a)) - 1)
   'DEF FNU$(a$) = Chr$(Asc(a$) And (255 + (160 * ((Asc(a$) > 96) And (Asc(a$) < 123)))))

   Select Case OPT$
      Case "INIT": GoTo 40
      Case "BUILD": Call BuildTABLEs
                    If NMCNT > 0 Then
                       Call ShowTable
                       If ICHG > 1 Then Call SaveCHANGEs: ICHG = ICHG And 1
                       frmFHSWUPgm.CMDButton(2).Caption = "Show Table"
                    End If
      Case "SHOW":  Call frmFHSWTTFL.ShowTable(SLIST$(), NMCNT, ICHG)
                    If ICHG > 0 Then Call SaveCHANGEs
      Case "PRINT": Call TTFLReport
   End Select
   Call FamCLOSE
   Exit Sub
   
40 ' Format Display
   Call PgmPREP("FHSTTFL")
   RptCODE = "TTFL"
   Load frmFHSWUPgm
   frmFHSWUPgm.UPSetup.AddItem RptCODE
   frmFHSWUPgm.UPSetup_Click
   frmFHSWUPgm.CMDButton(2).Caption = "Build Table"
   frmFHSWUPgm.CMDButton(2).Enabled = False
   frmFHSWUPgm.CMDButton(3).Caption = "Print"
   frmFHSWUPgm.CMDButton(3).Enabled = False
   frmFHSWUPgm.Show
   ReDim HLIT$(10)
   HLIT$(0) = Space$(29): HLIT$(1) = Space$(26)
   HLIT$(2) = Space$(26): HLIT$(3) = Space$(26)
   HLIT$(4) = "(000)000-0000": HLIT$(5) = Space$(10)
   HLIT$(6) = Space$(10): HLIT$(7) = Space$(26)
   HLIT$(8) = Space$(31): HLIT$(9) = Space$(31)
   Exit Sub

End Sub

Rem $DYNAMIC
Sub BuildTABLEs()
    READ$ = "READQ"
    SlctCNT = 0
    SURNAME$ = Space$(20): CURSURNM$ = SURNAME$
    ILVL$ = " .:*": ICHG = 0
    SXTB$ = "01230120022455012623010202"
    MINFAM = 1: MAXSUR = 1000
    KT$ = "QWERTYUIOPASDFGHJKLZXCVBNM"

600 ' Determine Records to be Processed
    If FF1Hdr.MAXID = 0 Then
       Call PutMSG("TTFLM002")
       GoTo 990
    End If

800 ' Gather SURNAME Information
    SlctCNT = SCNT
    If SlctCNT = 0 Then
       Call PutMSG("TTFLM004")
       GoTo 990
    End If
    Call FAMOpen("READ"): If Not OKAY Then GoTo 990
    xSNDX$ = "    ": CSNDX$ = xSNDX$
    LSet SURNAME$ = "": CURSURNM$ = SURNAME$
    ' Check Summary Options
    Call INDEXFile("CHECK"): If Not OKAY Then GoTo 990
    If (SX(1) \ 256) <> 1 Then
       Call PutMSG("TTFLM005")
       GoTo 990
    End If
    SNL = SX(1) And 127
    TSK = (ENV.SortOPTS And 3)
    Call INDEXFile("READQ")
    MAXSUR = 250
    ReDim ARID(MAXSUR + 2), DRID(MAXSUR + 2), RCNT(MAXSUR + 2), LVL(MAXSUR + 2), AID(MAXSUR + 2), SNDX(MAXSUR + 2)
    Call RptOptGET("MNLV", MINLVL)
    If MINLVL > 3 Then MINLVL = 3
    If MINLVL > 0 Then
       MINFAM = 1
       Else
       Call RptOptGET("MNSG", MINFAM)
    End If
    If MINFAM < 1 Then MINFAM = 1
    Call PutMSG("TTFLM007")

    ' Process Indexed Records
    SSEQ = 1: NMCNT = 0: SCNT = 0: ICHG = 0
    For ix = 1 To NDXHDR.NDXRECS
        xNDXID = NDXID(ix)
        If SelSET(xNDXID) <> 0 Then GoSub 860
        CA$ = INKEY$
        If CA$ <> "" Then
           If CA$ = Chr$(27) Then ix = NDXRECS
        End If
        If NMCNT > MAXSUR - 1 Then
           MAXSUR = MAXSUR + 50
           ReDim Preserve ARID(MAXSUR + 2), DRID(MAXSUR + 2), RCNT(MAXSUR + 2), LVL(MAXSUR + 2), AID(MAXSUR + 2), SNDX(MAXSUR + 2)
        End If
    Next ix
    If nRCNT > 0 Then GoSub 920
    Close: ZDEL = -1
    GoTo 990

860 ' Process Records for Surname
    Call FF1GetRec(xNDXID)
    RID = xNDXID: Call GetSURNM
    Call ComputeSNDX(SURNAME$, xSNDX$)
    If TSK > 1 Then        ' Index by Surname SNDX Only
       If (SCNT > 0) Then
          If (xSNDX$ <> CSNDX$) Then
             GoSub 920
             Else
             If TSK = 3 And Left$(SURNAME$, SNL) <> Left$(CURSURNM$, SNL) Then GoSub 920
          End If
       End If
       Else
       If (TSK = 1) Then LSet SURNAME$ = UCX$(SURNAME$)
       If SCNT > 0 Then
          If Left$(SURNAME$, SNL) <> Left$(CURSURNM$, SNL) Then GoSub 920
       End If
    End If
    LSet CURSURNM$ = SURNAME$: LSet CSNDX$ = xSNDX$
    SCNT = SCNT + 1
    If nRCNT = 0 Then xARID = 0: xDRID = 0: ABY = 9999: DBY = 0: xLVL = 0

    ' Check for Root (Earliest) Ancestor
    If FF1Rec.BY < ABY Then
       If (FF1Rec.BY > 0) Or (nRCNT = 0) Then
          ABY = FF1Rec.BY: xARID = xNDXID
          xLVL = ((xLVL And 243) Or (Asc(FF1Rec.STATUS2) And 12))
          If ABY = 0 Then ABY = 9999
       End If
    End If

    ' Check for End of Line (Latest) Descendant
    If FF1Rec.BY >= DBY Then
       DBY = FF1Rec.BY: xDRID = xNDXID
       xLVL = ((xLVL And 252) Or (Asc(FF1Rec.STATUS2) And (3)))
    End If
    nRCNT = nRCNT + 1
    Return

920 ' Put Info in Table
    If MINLVL > 0 Then
       AIL = ((xLVL And 12) / 4): DIL = (xLVL And 3)
       If (AIL < MINLVL) And (DIL < MINLVL) Then GoTo 929
    End If
    If nRCNT < MINFAM Then GoTo 929
    If CSNDX$ = "0000" Then GoTo 929
    If xARID = 0 Then xARID = xDRID: xLVL = (xLVL + ((xLVL And 3) * 4))
    xLVL = (xLVL - 64 * (DBY = 0)): NMCNT = NMCNT + 1
    ARID(NMCNT) = xARID: DRID(NMCNT) = xDRID
    RCNT(NMCNT) = nRCNT: LVL(NMCNT) = xLVL
    AID(NMCNT) = NMCNT: X = (Asc(CSNDX$) - 64)
    SNDX(NMCNT) = (1000 * X + Val(Right$(CSNDX$, 3)))
    TRPTCNT = TRPTCNT + nRCNT
929 nRCNT = 0
    Return

990 ' Finished
    ReDim Preserve ARID(NMCNT), DRID(NMCNT), RCNT(NMCNT), LVL(NMCNT), AID(NMCNT), SNDX(NMCNT)
    Call PutMSG("")
    
    Call FamCLOSE
    Exit Sub

End Sub


Rem $STATIC
Sub GetSURNM()
1670 ' Build Surname
     LSet SURNAME$ = FF1Rec.SURNM
     SPREC = FF1Rec.SPOUSE
     If SPREC <> 0 Then
        SPID = 0
        If (ENV.HUSBSURNM And (FF1Rec.SEX = "F") And ((Asc(FF1Rec.STATUS2) And 128) = 0)) Then
           Call FF3GetRec(SPREC)
           If RID = FF3SPOUSE.SP1ID Then
              SPID = FF3SPOUSE.SP2ID
              Else
              SPID = FF3SPOUSE.SP1ID
           End If
        End If
        If SPID > 0 Then
           Call FF1GetRec(SPID): LSet SURNAME$ = FF1Rec.SURNM
           Call FF1GetRec(RID)
        End If
     End If
End Sub

Rem $DYNAMIC
Sub ShowTable()
'SHARED ARID(), DRID(), RCNT(), LVL(), AID(), SNDX()
'SHARED FMT005, ICHG, SURNAME$, ILVL$, KT$, ZDEL
1700
   If NMCNT = 0 Then
      Call PutMSG("TTFLM009"): Exit Sub
   End If
   Call FAMOpen("READ"): If Not OKAY Then Exit Sub
   ReDim SNIX(NMCNT), SLIST$(NMCNT)
   SNIX(0) = 1: For X = 1 To NMCNT: SNIX(X) = X: Next X
   xX$ = Space$(37): For X = 1 To NMCNT: SLIST$(X) = xX$: Next X
   X = SNIX(0)
   While (X <= NMCNT)
      GoSub 1725: X = X + 1
   Wend
   Call frmFHSWTTFL.ShowTable(SLIST$(), NMCNT, ICHG)
   If ICHG > 0 Then Call SaveCHANGEs
1710
   Call FamCLOSE
   Exit Sub

1725
   Call FF1GetRec(ARID(X))
      ABYR = FF1Rec.BY: Call GetSURNM
   Call FF1GetRec(DRID(X))
      DBYR = FF1Rec.BY
   xLVL = LVL(X)
   Mid$(xX$, 2, 1) = Chr$(32 - (Asc("X") - 32) * ((xLVL And 128) = 128))
   Mid$(xX$, 3, 4) = SFILL$(RCNT(X), 4)
   Mid$(xX$, 7, 4) = ZFILL$(ABYR, 4)
   Mid$(xX$, 11, 1) = MID$(ILVL$, ((xLVL And 12) / 4) + 1, 1)
   Mid$(xX$, 12, 4) = ZFILL$(DBYR, 4)
   Mid$(xX$, 16, 1) = MID$(ILVL$, (xLVL And 3) + 1, 1)
   xSNDX$ = Chr$(SNDX(X) / 1000 + 64) + LTrim$(Str$(SNDX(X) Mod 1000))
   If TSK > 1 Then
      Mid$(xX$, 17, 20) = xSNDX$ + "  " + Left$(SURNAME$, SNL)
      Else
      Mid$(xX$, 17, 20) = Left$(SURNAME$, SNL)
   End If
   LSet SLIST$(X) = xX$
   Return


End Sub

Sub TTFLReport()
'SHARED FMT002, TSK, SSEQ, NMCNT, RID, SURNAME$, CDMY$, RRPT$, ILVL$
'SHARED ARID(), DRID(), RCNT(), LVL(), AID(), SNDX()
2000 ' Produce TINY TAFEL File/Report
     D1Y = Val(Right$(DATE$, 4)): D1MD = (100 * Val(Left$(DATE$, 2))) + Val(MID$(DATE$, 4, 2))
     CDMY$ = XRPTDate$(D1MD, D1Y, 0)
     If NMCNT = 0 Then
        Call PutMSG("TTFLM009")
        GoTo 2020
     End If
     SysVAR.SUBJECT = FDFSetup.NAME
     Call RptOPEN(RptCODE): If Not OKAY Then GoTo 2020
     Call FAMOpen("READ"): If Not OKAY Then GoTo 2020
     POFFSET = FOFFSET: 'LineCount = 0:
     GoTo 2040
2020 ' Early Return
     Call FamCLOSE
     Exit Sub

2040 ' Check Summary Options
     If TSK < 2 Then
        Call PutMSG("TTFLM010")
        If CA$ = Chr$(27) Then GoTo 2020
        X = Val(CA$): If (X > 2) Or (X < 1) Then X = 1
        If X <> SSEQ Then
           SSEQ = X
           If SSEQ = 2 Then
              GoSub 1200
              If a = 27 Then GoTo 2020
              Else
              For ix = 1 To NMCNT: AID(ix) = ix: Next ix
           End If
        End If
     End If
     Call RptOptGET("MNLV", MINLVL)
     If MINLVL > 3 Then MINLVL = 3
     If MINLVL > 0 Then
        MINRPT = MINFAM
        Else
        Call RptOptGET("MNSG", MINRPT)
        If MINFAM > MINRPT Then MINRPT = MINFAM
     End If
     Call RptOptGET("TTCT", X): ShowCOUNT = (X = 1)

     ' Enter TINY Tafel Header Info
     Call frmFHSWMsgD.GetTTHdrInfo(HLIT$(), "GET")
     HDRNAME$ = HLIT$(0): HDRADR1$ = HLIT$(1)
     HDRADR2$ = HLIT$(2): HDRADR3$ = HLIT$(3)
     HDRPHON$ = HLIT$(4): HDRSVC$ = HLIT$(5)
     HDRBBS$ = HLIT$(6): HDRREM1$ = HLIT$(7)
     HDRREM2$ = HLIT$(8): HDRREM3$ = HLIT$(9)

     Call PrintOPEN: If Not OKAY Then GoTo 2020
     PRT.AGAIN = 0

2105 ' Print TINY TAFEL Header Info
     If RTrim$(HDRNAME$) <> "" Then
        Mid$(PLINE$, 1, 1) = "N"
        Mid$(PLINE$, 3, 255) = HDRNAME$
        Call PrintLINE: If a = 27 Then GoTo 2175
     End If
     If RTrim$(HDRADR1$) <> "" Then
        Mid$(PLINE$, 1, 1) = "A"
        Mid$(PLINE$, 3, 255) = HDRADR1$
        Call PrintLINE: If a = 27 Then GoTo 2175
     End If
     If RTrim$(HDRADR2$) <> "" Then
        Mid$(PLINE$, 1, 1) = "A"
        Mid$(PLINE$, 3, 255) = HDRADR2$
        Call PrintLINE: If a = 27 Then GoTo 2175
     End If
     If RTrim$(HDRADR3$) <> "" Then
        Mid$(PLINE$, 1, 1) = "A"
        Mid$(PLINE$, 3, 255) = HDRADR3$
        Call PrintLINE: If a = 27 Then GoTo 2175
     End If
     If (RTrim$(HDRPHON$) <> "") And (RTrim$(HDRPHON$) <> "(000)000-0000") Then
        Mid$(PLINE$, 1, 1) = "T"
        Mid$(PLINE$, 3, 255) = HDRPHON$
        Call PrintLINE: If a = 27 Then GoTo 2175
     End If
     If RTrim$(HDRSVC$) <> "" Then
        Mid$(PLINE$, 1, 1) = "S"
        Mid$(PLINE$, 3, 255) = HDRSVC$
        Call PrintLINE: If a = 27 Then GoTo 2175
     End If
     If RTrim$(HDRBBS$) <> "" Then
        Mid$(PLINE$, 1, 1) = "B"
        Mid$(PLINE$, 3, 255) = HDRBBS$
        Call PrintLINE: If a = 27 Then GoTo 2175
     End If
     If RTrim$(HDRREM1$) <> "" Then
        Mid$(PLINE$, 1, 1) = "R"
        Mid$(PLINE$, 3, 255) = HDRREM1$
        Call PrintLINE: If a = 27 Then GoTo 2175
     End If
     If RTrim$(HDRREM2$) <> "" Then
        Mid$(PLINE$, 1, 1) = "R"
        Mid$(PLINE$, 3, 255) = HDRREM2$
        Call PrintLINE: If a = 27 Then GoTo 2175
     End If
     If RTrim$(HDRREM3$) <> "" Then
        Mid$(PLINE$, 1, 1) = "R"
        Mid$(PLINE$, 3, 255) = HDRREM3$
        Call PrintLINE: If a = 27 Then GoTo 2175
     End If

     ' Family File Description
     Mid$(PLINE$, 1, 1) = "F"
     Mid$(PLINE$, 3, 39) = "Family History System (             )"
     Mid$(PLINE$, 26, 14) = FDFSetup.NAME + " File"
     Call PrintLINE: If a = 27 Then GoTo 2175
     SCNT = 0: nRCNT = 0
     For ix = 1 To NMCNT
         If (LVL(ix) And 128) = 0 Then
            AIL = ((LVL(ix) And 12) / 4): DIL = (LVL(ix) And 3)
            If (AIL >= MINLVL) And (DIL >= MINLVL) Then
               If (RCNT(ix) >= MINRPT) Then
                  SCNT = SCNT + 1: nRCNT = nRCNT + RCNT(ix)
               End If
            End If
         End If
     Next ix
     Mid$(PLINE$, 1, 1) = "Z": Mid$(PLINE$, 3, 5) = fnx$(SCNT): O = 9
     If MINLVL > 0 Then
        Mid$(PLINE$, O, 2) = "I=": Mid$(PLINE$, O + 2, 5) = fnx$(MINLVL)
        O = InStr(O, PLINE$, "  ") + 2
        Else
        If MINRPT > 1 Then
           Mid$(PLINE$, O, 2) = "M=": Mid$(PLINE$, O + 2, 5) = fnx$(MINRPT)
           O = InStr(O, PLINE$, "  ") + 2
        End If
     End If
     Mid$(PLINE$, O, 2) = "T=": Mid$(PLINE$, O + 2, 5) = fnx$(nRCNT)
     O = InStr(O, PLINE$, "  ") + 2
     Mid$(PLINE$, O, 2) = "F=": Mid$(PLINE$, O + 2, 5) = fnx$(FF1Hdr.MAXID)
     Call PrintLINE: If a = 27 Then GoTo 2175

2150 ' Process Surname Table
     xSNDX$ = "    "
     For JX = 1 To NMCNT: ix = AID(JX)
         If (RCNT(ix) >= MINRPT) And ((LVL(ix) And 128) = 0) Then
            GoSub 2400: CA$ = INKEY$: If CA$ = Chr$(27) Then JX = NMCNT: a = 27
         End If
     Next JX: If a = 27 Then GoTo 2175
     Mid$(PLINE$, 1, 1) = "W": Mid$(PLINE$, 3, 13) = CDMY$
     Call PrintLINE

2175 ' Finished
     Call RptCLOSE
     'If PRT.OPT = 1 Then Call ScrnCLR: Call WinSWITCH(1)
     PRT.AGAIN = 0
     Call FamCLOSE
     Exit Sub

2400 ' Process Records for Surname
     xARID = ARID(ix): xDRID = DRID(ix)
     AIL = ((LVL(ix) And 12) / 4): DIL = (LVL(ix) And 3)
     If ((AIL >= MINLVL) Or (DIL >= MINLVL)) Then
        Call FF1GetRec(xARID): SPREC = FF1Rec.SPOUSE
        RID = xARID: Call GetSURNM
        ABLOC = FF1Rec.BLOC: ABY = FF1Rec.BY
        Call FF1GetRec(xDRID)
        DBLOC = FF1Rec.BLOC: DBY = FF1Rec.BY
        GoSub 2550
     End If
     Return

2550 ' Print Summary Line
     If a = 27 Then Return
     Call ComputeSNDX(SURNAME$, xSNDX$)
     Mid$(PLINE$, 1, 4) = xSNDX$
     If ABY > 0 Then
        Mid$(PLINE$, 6, 4) = ZFILL$(ABY, 4)
     End If
     Mid$(PLINE$, 10, 1) = MID$(ILVL$, AIL + 1, 1)
     If DBY > 0 Then
        Mid$(PLINE$, 11, 4) = ZFILL$(DBY, 4)
     End If
     Mid$(PLINE$, 15, 1) = MID$(ILVL$, DIL + 1, 1)
     If ShowCOUNT Then
        Mid$(PLINE$, 16, 6) = "(" + SFILL$(RCNT(ix), 4) + ")"
        O = 23
        Else
        O = 16
     End If
     Mid$(PLINE$, O, 25) = SURNAME$
     O = InStr(O, PLINE$, "   ")
     If ABLOC <> 0 Then
        LSet BPLACE = ""
        Call FF3GetRec(ABLOC)
        If Asc(BPLACE) <> 32 Then
           Mid$(PLINE$, O, 1) = "\"
           Mid$(PLINE$, O + 1, 255) = BPLACE
           O = InStr(O, PLINE$, "   ")
        End If
     End If
     If DBLOC <> 0 Then
        LSet BPLACE = ""
        Call FF3GetRec(DBLOC)
        If Asc(BPLACE) <> 32 Then
           Mid$(PLINE$, O, 1) = "/"
           Mid$(PLINE$, O + 1, 255) = BPLACE
           O = InStr(O, PLINE$, "   ")
        End If
     End If
     Call PrintLINE
     Return

1200 ' Use Segmented SHELL Sort/Merge for sorting SNDX()
     NS = NMCNT: a = 0: CA$ = "": Call PutMSG("TTFLM012")
1250 ' Sort Segments of File
     D = 2 ^ (Int(Log(NS + 1) / Log(2))) - 1
1254 For ix = 1 To D: For JX = ix To NS - D Step D: X = AID(JX): Y = AID(JX + D)
     CA$ = INKEY$
     If CA$ <> "" Then
        'LSet MDY$ = Time$: LOCATE 25, 1, 0
        'Print Tab(27); "NS="; Str$(NS); ",D="; Str$(D); ",IX="; Str$(ix); ",JX="; Str$(JX); Tab(78);
        If Asc(CA$) = 27 Then ix = D: JX = NS - D: D = 0: GoTo 1270
     End If
     If Not (SNDX(X) > SNDX(Y)) Then GoTo 1270
1260 Z = Y: kx = JX: AID(kx + D) = X
1261 kx = kx - D
     If kx < ix Then GoTo 1268 Else Y = AID(kx)
     CA$ = INKEY$
     If CA$ <> "" Then
        'LSet MDY$ = Time$: LOCATE 25, 1, 0
        'Print Tab(27); "NS="; Str$(NS); ",D="; Str$(D); ",IX="; Str$(IXB); ",JX="; Str$(JX); ",KX="; Str$(kx); Tab(78);
        If Asc(CA$) = 27 Then ix = D: JX = NS - D: D = 0: GoTo 1270
     End If
1266 If SNDX(Z) < SNDX(Y) Then AID(kx + D) = Y: GoTo 1261
1268 AID(kx + D) = Z
1270 Next JX: Next ix: D = D \ 2: If D > 0 Then GoTo 1254
     If CA$ <> "" Then a = Asc(CA$)
     Return
End Sub

Sub SaveCHANGEs()
1705 ' Transfer Level Table Changes to LVL()
     X = SNIX(0): xX$ = Space$(Len(SLIST$(1)))
     While (X <= NMCNT)
        If Asc(SLIST$(X)) > 0 Then GoSub 1740
        X = X + 1
     Wend
     GoTo 1750
   
1740 ' Put Level Changes Into Level Table
     xLVL = LVL(X): LSet xX$ = SLIST$(X)
     DFLG = -(128 * (MID$(xX$, 2, 1) <> " "))
     If (xLVL And 128) <> DFLG Then
        xLVL = (xLVL And 127) Or DFLG
     End If
     ALVL$ = MID$(xX$, 11, 1): DLVL$ = MID$(xX$, 16, 1)
     l1$ = MID$(ILVL$, 2, 1): L2$ = MID$(ILVL$, 3, 1): L3$ = MID$(ILVL$, 4, 1)
     DIL = -1 * (DLVL$ = l1$) - 2 * (DLVL$ = L2$) - 3 * (DLVL$ = L3$)
     AIL = -4 * (ALVL$ = l1$) - 8 * (ALVL$ = L2$) - 12 * (ALVL$ = L3$)
     If (xLVL And 3) <> DIL Then
        xLVL = ((xLVL And 252) Or DIL) Or 16: ICHG = -1
     End If
     If (xLVL And 12) <> AIL Then
        xLVL = ((xLVL And 243) Or AIL) Or 32: ICHG = -1
     End If
     LVL(X) = xLVL
     Return
   
1750 ' Save Level Changes in Family File
     If ICHG = 0 Then GoTo 1799
     Call PutMSG("TTFLM013")
     If CA$ <> SNGLKEY$(1) Then GoTo 1799
     Call FamCLOSE: Call FAMOpen("WRITE"): If Not OKAY Then GoTo 1799
     For ix = 1 To NMCNT:
         If Asc(SLIST$(ix)) > 0 Then
            xLVL = LVL(ix)
            If (xLVL And 16) <> 0 Then
               Call FF1GetRec(DRID(ix))
               If OKAY Then
                  FF1Rec.STATUS2 = Chr$((Asc(FF1Rec.STATUS2) And 252) Or (xLVL And 3))
                  Call FF1PutREC
               End If
            End If
            If (xLVL And 32) <> 0 Then
               Call FF1GetRec(ARID(ix))
               If OKAY Then
                  FF1Rec.STATUS2 = Chr$((Asc(FF1Rec.STATUS2) And 243) Or (xLVL And 12))
                  Call FF1PutREC
               End If
            End If
            LVL(ix) = (xLVL And 207)
            CA$ = INKEY$
            If CA$ <> "" Then
               If CA$ = Chr$(27) Then ix = NMCNT
               Else
               SysVAR.NVAR1 = ix
               Call PutMSG("TTFLM014")
            End If
            Mid$(SLIST$(ix), 1, 1) = Chr$(0)
         End If
     Next ix
     ICHG = 0
     
1799 ' Return from Saving Level Changes
     Call FamCLOSE
     Call PutMSG("")
End Sub

