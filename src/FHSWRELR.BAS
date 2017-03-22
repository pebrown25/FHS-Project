Attribute VB_Name = "modFHSWRelr"
DefInt A-Z

Type SUMREC
   TOTAL    As Integer
   MALE     As Integer
   FEMALE   As Integer
   CLINE    As Integer
   DUPID    As Integer
   NOP      As Integer
   BDT      As Integer
   BPL      As Integer
   LOBY     As Integer
   HIBY     As Integer
   MARRIED  As Integer
   MARCNT   As Integer
   MDT      As Integer
   MPL      As Integer
   LOMY     As Integer
   HIMY     As Integer
   DDT      As Integer
   DDTV     As Integer
   DPL      As Integer
   LODY     As Integer
   HIDY     As Integer
End Type

Dim DFNUM(), DFPP()
Dim SPRESEQ, SPRNO, SPCNT, MAXMarry, MLIST()
Dim DATECNT, DATERNG, PLACCNT, MARRCNT, RWIDTH, VRTCL
Dim RELID, PRPAGE(), PRLINE(), RX()
Dim SUMBASE As SUMREC, SUM As SUMREC, TOT As SUMREC

Sub RptRELR(OPT$)
1  Rem $INCLUDE: 'FHSCOMON.BAS'
   PN$ = " FAMILY HISTORY SYSTEM - FHSWRELR - Ancestor/Descendant/Relative Report Program"
   CY$ = " (C) Copyright 1986-1998 by Phillip E. Brown"
   Select Case OPT$
      Case "INIT": GoTo 40
      Case "PRINT": GoTo 1000
   End Select
   
40 ' Format Display
   Call PgmPREP("FHSRELR")
   Load frmFHSWRPgm
   Select Case RptCODE
      Case "ARPT": SRPT$ = "ASUM": RRPT$ = "ARPT": DRPT$ = "ADUP"
      Case "ASUM": SRPT$ = "ASUM": RRPT$ = "ARPT": DRPT$ = "ADUP"
      Case "DRPT": SRPT$ = "DSUM": RRPT$ = "DRPT": DRPT$ = "DDUP"
      Case "DSUM": SRPT$ = "DSUM": RRPT$ = "DRPT": DRPT$ = "DDUP"
      Case "RRPT": SRPT$ = "RSUM": RRPT$ = "RRPT": DRPT$ = ""
      Case "RSUM": SRPT$ = "RSUM": RRPT$ = "RRPT": DRPT$ = ""
   End Select
   Call RptLOAD(SRPT$): Call RptLOAD(RRPT$)
   If DRPT$ <> "" Then Call RptLOAD(DRPT$)
   frmFHSWRPgm.RPSetup.AddItem RRPT$
   Y = frmFHSWRPgm.RPSetup.NewIndex
   frmFHSWRPgm.RPSetup.ItemData(Y) = 1
   frmFHSWRPgm.RPSetup.AddItem SRPT$
   Y = frmFHSWRPgm.RPSetup.NewIndex
   frmFHSWRPgm.RPSetup.ItemData(Y) = 2
   If DRPT$ <> "" Then
      frmFHSWRPgm.RPSetup.AddItem DRPT$
      Y = frmFHSWRPgm.RPSetup.NewIndex
      frmFHSWRPgm.RPSetup.ItemData(Y) = 3
   End If
   frmFHSWRPgm.RPSetup_Click
   frmFHSWRPgm.Show
   Exit Sub

1000 'PROCESS REPORT OPTIONS
    Dim SUMBASE As SUMREC, SUM As SUMREC, TOT As SUMREC
    ReDim PTRTAB&(10), RX(1), PRPAGE(1), PRLINE(1)
    ReDim DFNUM(1), DFPP(1), RFLTH(1), RFLBL$(1)
    MAXMarry = 10: ReDim MLIST(MAXMarry)
    D1Y = Val(Right$(DATE$, 4))
    D1MD = (100 * Val(Left$(DATE$, 2)) + Val(MID$(DATE$, 4, 2)))
    CurDATE$ = XSCRNDate$(D1MD, D1Y)
    ROPT$ = Space$(31)
    XRELATION$ = Space$(12): WTYPE = MO: BRID = 0: READ$ = "READQ"
    Select Case RptCODE
       Case "ARPT": WTYPE = 1: X = 1
       Case "ASUM": WTYPE = 1: X = 2
       Case "ADUP": WTYPE = 1: X = 3
       Case "DRPT": WTYPE = 2: X = 1
       Case "DSUM": WTYPE = 2: X = 2
       Case "DDUP": WTYPE = 2: X = 3
       Case "RRPT": WTYPE = 3: X = 1
       Case "RSUM": WTYPE = 3: X = 2
    End Select
    Select Case X
       Case 1: Call XDETReport(WTYPE)
       Case 2: Call XSUMReport(WTYPE)
       Case 3: Call XDUPReport(WTYPE)
    End Select
    Call ClearMSG
End Sub

Rem $DYNAMIC
Sub BldDetDEF()
'SHARED DFNUM(), DFPP()
      RFCNT = 24
      ReDim DFNUM(RFCNT), DFPP(RFCNT)
      ReDim RFNUM(RFCNT), RFSEP(RFCNT), RFLTH(RFCNT), RFLBL$(RFCNT)
      Call RptOptGET("MNGP", MINGAP)
      If (Not OKAY) Or (MINGAP = 0) Then MINGAP = 2
      Call RptOptGET("MXGP", MAXGAP)
      If (Not OKAY) Or (MAXGAP = 0) Then MAXGAP = 5
      Call RptOptGET("MINM", MINMLTH)
      Call RptOptGET("PNML", PLNMLTH)


59500 ' Select Report Fields - Based Upon Report Options
      ' First Select Prefix Fields
      OKAY = 0
      PAGE.DFCNT = 0: IDFX = 0: BFX = 0: MFX = 0: DFX = 0
      RptOPTION.TWOLINES = 0: PRID2 = 0
      AMAXLV = Asc(WFHDR.AMAXLV): DMAXLV = Asc(WFHDR.DMAXLV)
      If (PAGE.BWIDTH = 0) Then
         BWIDTH = PRT.FWIDTH - PAGE.IMARGIN - PAGE.OMARGIN
         Else
         BWIDTH = PAGE.BWIDTH
      End If
      If AMAXLV > 0 Then
         PAGE.DFCNT = PAGE.DFCNT + 1
         dLINEAGE# = -(2 ^ (AMAXLV + 1) - 1#)
         DFNUM(PAGE.DFCNT) = 2
         RFLTH(2) = Len(Str$(Val(Str$(dLINEAGE#))))
      End If
      If RptOPTION.PRGL Then
         If (AMAXLV > 0) Or (DMAXLV > 0) Then
            X = AMAXLV: If X < DMAXLV Then X = DMAXLV
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 24
            RFLTH(24) = Len(RTrim$(Str$(X)))
         End If
      End If
      If RptOPTION.Relation Then
         PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 3
      End If
      If (DMAXLV > 0) And (Not RptOPTION.OUTLINE) Then
         PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 4
         RFLTH(4) = 3 * DMAXLV - 1
      End If
      If RptOPTION.PRID And (RptOPTION.FREEFORM Or (RptOPTION.COMMENTS = 0)) Then
         IDFX = PAGE.DFCNT + 1: RFSEP(6) = MINGAP + 1
         DFNUM(IDFX) = 6: DFNUM(IDFX + 1) = 7: DFNUM(IDFX + 2) = 8
         PAGE.DFCNT = PAGE.DFCNT + 3
         If ((RptOPTION.SPOUSE And 1) > 0) Then
            DFNUM(IDFX + 3) = 9: PAGE.DFCNT = PAGE.DFCNT + 1
            Else
            RFLTH(9) = 0
         End If
      End If

      ' Select Data Fields
      RFSEP(DFNUM(1)) = 1: PAGE.PFCNT = PAGE.DFCNT
      PAGE.DFCNT = PAGE.DFCNT + 1
      DFNUM(PAGE.DFCNT) = 5: RFSEP(5) = MINGAP + 1
      RFLTH(5) = MINMLTH - RptOPTION.OUTLINE * (3 * DMAXLV) - 20 * RptOPTION.FREEFORM
      If RptOPTION.FREEFORM Then GoTo 59510
      If RptOPTION.PRID And RptOPTION.COMMENTS Then
         IDFX = PAGE.DFCNT + 1
         DFNUM(IDFX) = 6: DFNUM(IDFX + 1) = 7: DFNUM(IDFX + 2) = 8
         PAGE.DFCNT = PAGE.DFCNT + 3
         If ((RptOPTION.SPOUSE And 1) > 0) Then
            DFNUM(IDFX + 3) = 9: PAGE.DFCNT = PAGE.DFCNT + 1
            Else
            RFLTH(9) = 0
         End If
      End If
      If RptOPTION.TIMELINE Then
         DFNUM(PAGE.DFCNT + 1) = 22: DFNUM(PAGE.DFCNT + 2) = 23
         PAGE.DFCNT = PAGE.DFCNT + 2
         GoTo 59510
      End If
      If RptOPTION.PRSEX Then PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 10
      If RptOPTION.PRAGE Then PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 11
      If RptOPTION.DateS Then
         PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 12
         BFX = PAGE.DFCNT
      End If
      If RptOPTION.PLACES Then
         PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 13
      End If
      If RptOPTION.SPOUSE Then
         If RptOPTION.DateS Then
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 15: MFX = PAGE.DFCNT
         End If
         If RptOPTION.PRAGE Then
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 18
         End If
         If RptOPTION.PLACES Then
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 16
         End If
      End If
      If RptOPTION.DateS Then
         PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 19: DFX = PAGE.DFCNT
      End If
      If RptOPTION.PLACES Then PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 20

59510 ' Create Report Definition Tables
      X = 1: RFNAME$ = "06010": GoSub 59505     'COMMENTS
      X = 2: RFNAME$ = "07011": GoSub 59505     'LNUM
      X = 3: RFNAME$ = "07004": GoSub 59505     'RELATION
      X = 4: RFNAME$ = "07012": GoSub 59505     'BLOODLINE
      X = 5: RFNAME$ = "01020": GoSub 59505     'FULL NAME
      X = 6: RFNAME$ = "01010": GoSub 59505     'RID
      X = 7: RFNAME$ = "01011": GoSub 59505     'FID
      X = 8: RFNAME$ = "01012": GoSub 59505     'MID
      X = 9: RFNAME$ = "02010": GoSub 59505     'SPID
      X = 10: RFNAME$ = "01015": GoSub 59505    'SEX CODE
      X = 11: RFNAME$ = "01016": GoSub 59505    'AGE
      X = 12: RFNAME$ = "01030": GoSub 59505    'BIRTH DATE
      X = 13: RFNAME$ = "01032": GoSub 59505    'BIRTH PLACE
              RFLTH(13) = PLNMLTH
      X = 14: RFNAME$ = "01034": GoSub 59505    'BIRTH DATE & PLACE
              RFLTH(14) = PLNMLTH
      X = 15: RFNAME$ = "02020": GoSub 59505    'MARRIAGE DATE
      X = 16: RFNAME$ = "02021": GoSub 59505    'MARRIAGE PLACE
              RFLTH(16) = PLNMLTH
      X = 17: RFNAME$ = "02024": GoSub 59505    'MARRIAGE DATE & PLACE
              RFLTH(17) = PLNMLTH
      X = 18: RFNAME$ = "02015": GoSub 59505    'Last ANNIVERSARY
      X = 19: RFNAME$ = "01040": GoSub 59505    'DEATH DATE
      X = 20: RFNAME$ = "01042": GoSub 59505    'DEATH PLACE
              RFLTH(20) = PLNMLTH
      X = 21: RFNAME$ = "01044": GoSub 59505    'DEATH DATE & PLACE
              RFLTH(21) = PLNMLTH
      X = 22: RFNAME$ = "07020": GoSub 59505    'LIFE SPAN
      X = 23: RFNAME$ = "07021": GoSub 59505    'LIFE LINE
      X = 24                                    'AGL, DGL or RGL
      If AMAXLV > 0 Then
         If DMAXLV = 0 Then
            RFNAME$ = "07001"
            Else
            RFNAME$ = "07003"
         End If
         Else
         RFNAME$ = "07002"
      End If
      GoSub 59505
      GoTo 59520

59505 ' Process Report Field
      Call GetRFLD(RFNAME$, FX)
      If RFSEP(X) = 0 Then RFSEP(X) = MINGAP
      If FOUND Then
         RDFFld = RFLDS(FX)
         If RFLTH(X) = 0 Then RFLTH(X) = Asc(RDFFld.RLTH)
         ALTH = Asc(RDFFld.ALTH): LLTH = Asc(RDFFld.LLTH)
         HLTH = Asc(RDFFld.HLTH)
         If (X = 1) Or (RptOPTION.FREEFORM And ((X > 9) And (X < 24))) Then
            RFLBL$(X) = MID$(RFLD$(FX), ALTH + 1, LLTH)
            Else
            RFLBL$(X) = MID$(RFLD$(FX), ALTH + LLTH + 1, HLTH)
         End If
      End If
      Return

59520 ' Check Report Width
      RWIDTH = 0: FX = 0
      While FX < PAGE.DFCNT
          FX = FX + 1: X = DFNUM(FX): FSEP = RFSEP(X)
          If (X > 0) And (FSEP > 0) Then
             Y = RFLTH(X): z = Len(RFLBL$(X)): If Y < z Then Y = z
             RWIDTH = RWIDTH + FSEP + Y
          End If
      Wend
      RWIDTH = RWIDTH - 1
      If RWIDTH <= BWIDTH Then GoTo 59540
      If RptOPTION.TIMELINE Or RptOPTION.FREEFORM Then GoTo 59530
      If Not (RptOPTION.DateS And RptOPTION.PLACES) Then GoTo 59530
      If Not RptOPTION.TWOLINES Then
         RptOPTION.TWOLINES = -1
         DFNUM(BFX) = 14: DFNUM(BFX + 1) = 0
         DFNUM(DFX) = 21: DFNUM(DFX + 1) = 0
         If RptOPTION.SPOUSE Then
            DFNUM(MFX) = 17: DFNUM(MFX + 1) = 0
            If RptOPTION.PRAGE Then DFNUM(MFX + 2) = 0
         End If
         GoTo 59520
      End If
      If RptOPTION.PRID And (Not PRID2) Then
         PRID2 = -1
         DFNUM(IDFX) = 0: DFNUM(IDFX + 1) = 0: DFNUM(IDFX + 2) = 0
         If ((RptOPTION.SPOUSE And 1) > 0) Then DFNUM(IDFX + 3) = 0
         X = RFLTH(6) + RFLTH(7) + RFLTH(8) + RFLTH(9) + 18
         If X > RFLTH(5) Then RFLTH(5) = X
         GoTo 59520
      End If

59530 ' Final Check After Trying TWOLINES and PRID2 Options
      PWIDTH = BWIDTH + PAGE.IMARGIN + PAGE.OMARGIN
      RPWIDTH = RWIDTH + PAGE.IMARGIN + PAGE.OMARGIN
      If PWIDTH < RPWIDTH Then
         If (PRT.OPT = 2) Then
            SysVAR.NVAR1 = PWIDTH: SysVAR.NVAR2 = RPWIDTH
            Call PutMSG("XXXXM009")
            If CA$ <> SNGLKEY$(1) Then OKAY = 0: Exit Sub
         End If
         BWIDTH = RWIDTH: PAGE.WIDTH = RPWIDTH
         If Len(PLINE$) < RPWIDTH Then PLINE$ = Space$(RPWIDTH)
      End If

59540 ' Adjust to Fit
      If RptOPTION.FREEFORM Then GoTo 59550
      If Not (RWIDTH < BWIDTH) Then GoTo 59550
      If RptOPTION.TIMELINE Then
         RFLTH(23) = RFLTH(23) + (BWIDTH - RWIDTH): GoTo 59550
      End If
      GAPS = 0: FX = PAGE.PFCNT
      While FX <= PAGE.DFCNT
          X = DFNUM(FX)
          If (X > 0) And (RFSEP(X) > 0) And (X <> 7) And (X <> 8) Then
             GAPS = GAPS + 1
          End If
          FX = FX + 1
      Wend
      GAP = (BWIDTH - RWIDTH) \ GAPS
      GAP1 = (BWIDTH - RWIDTH) - GAPS * GAP
      If GAP >= (MAXGAP - MINGAP) Then GAP = MAXGAP - MINGAP: GAP1 = 0
      FX = PAGE.PFCNT
      While FX <= PAGE.DFCNT
          X = DFNUM(FX): FSEP = RFSEP(X)
          If (X > 0) And (FSEP > 0) And (X <> 7) And (X <> 8) Then
             RFSEP(X) = FSEP + GAP - (GAP1 > 0): GAP1 = GAP1 - 1
          End If
          FX = FX + 1
      Wend

59550 ' Determine Field Print Positions
      PP = 0: FX = 0
      While FX < PAGE.DFCNT
          FX = FX + 1: X = DFNUM(FX): FSEP = RFSEP(X)
          If (X > 0) And (FSEP > 0) Then
             FL = RFLTH(X): HL = Len(RFLBL$(X))
             z = FL: If z < HL Then z = HL
             PP = PP + FSEP + (z + 1) \ 2 - (FL + 1) \ 2
             DFPP(FX) = PP
             PP = PP + z
          End If
      Wend

59560 ' Return to Calling Program
      PAGE.BWIDTH = BWIDTH: PAGE.LWIDTH = BWIDTH
      If RptOPTION.PRID Then
         If PRID2 Then RptOPTION.PRID = 2 Else RptOPTION.PRID = 1
      End If
      OKAY = -1
End Sub

Rem $STATIC
Sub BldDetTITLE()
'SHARED DFNUM(), DFPP()
'SHARED TLBY, TLEY
     ' Build Title Line
      TL$ = Space$(PAGE.BWIDTH)
      For FX = 1 To PAGE.DFCNT: X = DFNUM(FX): PP = DFPP(FX)
         If (X > 0) And (PP > 0) Then
            FL = RFLTH(X): HL = Len(RFLBL$(X))
            If X = 23 Then
               Mid$(TL$, PP, 4) = FNX$(TLBY)
               Mid$(TL$, PP + FL - 5, 4) = SFILL$(TLEY, 4)
            End If
            If HL > 0 Then
               If (RptOPTION.FREEFORM = 0) Or (FX <= PAGE.PFCNT) Then
                  PP = PP + (FL + 1) \ 2 - (HL + 1) \ 2
                  If PP > 0 Then Mid$(TL$, PP, HL) = RFLBL$(X)
               End If
            End If
         End If
      Next FX
      X = Len(RTrim$(TL$))
      If X > 0 Then
         HTFLMAX = HTFLMAX + 1
         ReDim Preserve HTFLINE$(HTFLMAX)
         HTFLINE$(HTFLMAX) = "T00" + RTrim$(TL$)
      End If
      OKAY = -1
End Sub

Sub BldDupDEF()
      PAGE.DFCNT = 4
      ReDim DFNUM(PAGE.DFCNT), DFPP(PAGE.DFCNT)
      ReDim RFNUM(PAGE.DFCNT), RFSEP(PAGE.DFCNT), RFLTH(PAGE.DFCNT), RFLBL$(PAGE.DFCNT)
      X = 1: RFNAME$ = "01010": GoSub DupFLD     'ID
      X = 2: RFNAME$ = "01020": GoSub DupFLD     'NAME
      X = 3: RFNAME$ = "07004": GoSub DupFLD     'Relation
      X = 4: RFNAME$ = "07010": GoSub DupFLD     'LINEAGE
      Exit Sub

DupFLD: ' Process Report Field
      Call GetRFLD(RFNAME$, FX)
      If FOUND Then
         RDFFld = RFLDS(FX)
         If RFLTH(X) = 0 Then RFLTH(X) = Asc(RDFFld.RLTH)
         ALTH = Asc(RDFFld.ALTH): LLTH = Asc(RDFFld.LLTH)
         RFLBL$(X) = MID$(RFLD$(FX), ALTH + 1, LLTH)
      End If
      Return

End Sub

Sub BldMLIST()
'SHARED SPRESEQ, SPRNO, SPCNT, MAXMarry, MLIST()
   SPCNT = 0
   If SPRESEQ Then
      While SPRNO <> 0
         SPCNT = SPCNT + 1
         If SPCNT > MAXMarry Then
            MAXMarry = MAXMarry + 5
            ReDim Preserve MLIST(MAXMarry)
         End If
         MLIST(SPCNT) = SPRNO
         Call FF3GetRec(SPRNO)
         If FF1Rec.RID = FF3SPOUSE.SP1ID Then
            SPRNO = FF3SPOUSE.SP1NXT
            Else
            If FF1Rec.RID = FF3SPOUSE.SP2ID Then
               SPRNO = FF3SPOUSE.SP2NXT
               Else
               SPRNO = 0
            End If
         End If
      Wend
      SPRNO = MLIST(SPCNT)
   End If

End Sub

Sub BldSumDEF()
'SHARED DFNUM(), DFPP()
'SHARED DATECNT, DATERNG, PLACCNT, MARRCNT, RWIDTH, VRTCL
      AMAXLV = Asc(WFHDR.AMAXLV): DMAXLV = Asc(WFHDR.DMAXLV)
      RELSUMOPT = (AMAXLV > 0) And (DMAXLV > 0)
      If RELSUMOPT Then RFCNT = AMAXLV + 3 Else RFCNT = 23
      PAGE.DFCNT = 0
      ReDim DFNUM(RFCNT), DFPP(RFCNT)
      ReDim RFNUM(RFCNT), RFSEP(RFCNT), RFLTH(RFCNT), RFLBL$(RFCNT)
      Call RptOptGET("MNGP", MINGAP)
      If (Not OKAY) Or (MINGAP = 0) Then MINGAP = 2
      DFPP(0) = MINGAP
      Call RptOptGET("MXGP", MAXGAP)
      If (Not OKAY) Or (MAXGAP = 0) Then MAXGAP = 5
      If RELSUMOPT Then GoTo 58500
      If DMAXLV = 0 Then
         Call RptOptGET("CNOP", X): CNTNOP = (X = 1)
      End If
      Call RptOptGET("CDUP", X): CNTDUP = (X = 1)
      Call RptOptGET("CLIN", X): CNTCLN = (X = 1)
      Call RptOptGET("DTCT", X): DATECNT = (X = 1)
      Call RptOptGET("DTRG", X): DATERNG = (X = 1)
      Call RptOptGET("PLCT", X): PLACCNT = (X = 1)
      Call RptOptGET("MCNT", X): MARRCNT = (X = 1)

58500 ' Select Report Fields - Based Upon Report Options
      OKAY = 0
      If (PAGE.BWIDTH = 0) Then
         BWIDTH = PRT.FWIDTH - PAGE.IMARGIN - PAGE.OMARGIN
         Else
         BWIDTH = PAGE.BWIDTH
      End If
      If RELSUMOPT Then GoTo 58520
      PAGE.DFCNT = 1: DFNUM(1) = 1
      If RptOPTION.Relation Or VRTCL Then
         PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 2
      End If
      If CNTCLN Then PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 4
      If CNTNOP Then PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 6
      If CNTDUP Then PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 5
      PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 7
      PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 8
      PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 3
      If DATECNT Then PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 9
      If DATERNG Then
         DFNUM(PAGE.DFCNT + 1) = 11: DFNUM(PAGE.DFCNT + 2) = 12
         PAGE.DFCNT = PAGE.DFCNT + 2
      End If
      If PLACCNT Then PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 10
      If MARRCNT Then
         DFNUM(PAGE.DFCNT + 1) = 13: DFNUM(PAGE.DFCNT + 2) = 14
         PAGE.DFCNT = PAGE.DFCNT + 2
         If DATECNT Then PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 15
         If DATERNG Then
            DFNUM(PAGE.DFCNT + 1) = 17: DFNUM(PAGE.DFCNT + 2) = 18
            PAGE.DFCNT = PAGE.DFCNT + 2
         End If
         If PLACCNT Then PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 16
      End If
      If DATECNT Then
         DFNUM(PAGE.DFCNT + 1) = 19: DFNUM(PAGE.DFCNT + 2) = 20
         PAGE.DFCNT = PAGE.DFCNT + 2
      End If
      If DATERNG Then
         DFNUM(PAGE.DFCNT + 1) = 22: DFNUM(PAGE.DFCNT + 2) = 23
         PAGE.DFCNT = PAGE.DFCNT + 2
      End If
      If PLACCNT Then PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 21

58510 ' Create Report Definition Tables
      X = 1                                     'AGL or DGL
      If (DMAXLV > 0) And (AMAXLV = 0) Then
         RFNAME$ = "07002"
         Else
         RFNAME$ = "07001"
      End If
      GoSub 58515
      X = 2: RFNAME$ = "07004": GoSub 58515     'Relation
      X = 3: RFNAME$ = "08001": GoSub 58515     'Total
      X = 4: RFNAME$ = "08002": GoSub 58515     'Lines
      X = 5: RFNAME$ = "08003": GoSub 58515     'Duplicate ID's
      X = 6: RFNAME$ = "08004": GoSub 58515     'No Parents
      X = 7: RFNAME$ = "08006": GoSub 58515     'Males
      X = 8: RFNAME$ = "08007": GoSub 58515     'Females
      X = 9: RFNAME$ = "08015": GoSub 58515     'BY>0
      X = 10: RFNAME$ = "08016": GoSub 58515    '#BPL
      X = 11: RFNAME$ = "08017": GoSub 58515    'LOBY
      X = 12: RFNAME$ = "08018": GoSub 58515    'HIBY
      X = 13: RFNAME$ = "08020": GoSub 58515    'MAR
      X = 14: RFNAME$ = "08021": GoSub 58515    '#MAR
      X = 15: RFNAME$ = "08025": GoSub 58515    'MY>0
      X = 16: RFNAME$ = "08026": GoSub 58515    '#MPL
      X = 17: RFNAME$ = "08027": GoSub 58515    'LOMY
      X = 18: RFNAME$ = "08028": GoSub 58515    'HIMY
      X = 19: RFNAME$ = "08034": GoSub 58515    'DY>0
      X = 20: RFNAME$ = "08035": GoSub 58515    'DY<9
      X = 21: RFNAME$ = "08036": GoSub 58515    '#DPL
      X = 22: RFNAME$ = "08037": GoSub 58515    'LODY
      X = 23: RFNAME$ = "08038": GoSub 58515    'HIDY
      GoTo 58525

58515 ' Process Report Field
      Call GetRFLD(RFNAME$, FX)
      If RFSEP(X) = 0 Then RFSEP(X) = MINGAP
      If FOUND Then
         RDFFld = RFLDS(FX)
         If RFLTH(X) = 0 Then RFLTH(X) = Asc(RDFFld.RLTH)
         ALTH = Asc(RDFFld.ALTH): LLTH = Asc(RDFFld.LLTH)
         HLTH = Asc(RDFFld.HLTH)
         If VRTCL Then
            RFLBL$(X) = MID$(RFLD$(FX), ALTH + 1, LLTH)
            Else
            RFLBL$(X) = MID$(RFLD$(FX), ALTH + LLTH + 1, HLTH)
         End If
      End If
      Return

58520 ' Relationship Summary Table
      PAGE.DFCNT = RFCNT
      DFNUM(1) = 1: RFSEP(1) = 1: RFLTH(1) = 8
      Call GetRFLD("07003", FX)
      If FOUND Then
         RDFFld = RFLDS(FX)
         O = Asc(RDFFld.ALTH) + Asc(RDFFld.LLTH): L = Asc(RDFFld.HLTH)
         RFLBL$(1) = MID$(RFLD$(FX), O + 1, L)
      End If
      For X = 2 To RFCNT
          DFNUM(X) = X
          RFSEP(X) = MINGAP: RFLTH(X) = 6
          RFLBL$(X) = Str$(2 - X)
      Next X
      Call GetRFLD("08001", FX)
      If FOUND Then
         RDFFld = RFLDS(FX)
         O = Asc(RDFFld.ALTH) + Asc(RDFFld.LLTH): L = Asc(RDFFld.HLTH)
         RFLBL$(RFCNT) = MID$(RFLD$(FX), O + 1, L)
      End If

58525 ' Check Report Width
      GoSub 58530
      If RWIDTH <= BWIDTH Then GoTo 58540
      GoTo 58535

58530 ' Determine Required Width
      RWIDTH = 1: FX = 0
      While FX < PAGE.DFCNT
          FX = FX + 1: X = DFNUM(FX): FSEP = RFSEP(X)
          If (X > 0) And (FSEP > 0) Then
             Y = RFLTH(X): z = Len(RFLBL$(X))
             If VRTCL Then
                If RWIDTH < z Then RWIDTH = z
                Else
                If Y < z Then Y = z
                RWIDTH = RWIDTH + FSEP + Y
             End If
          End If
      Wend
      If VRTCL Then
         RFLTH(0) = RWIDTH
         RWIDTH = RWIDTH + (AMAXLV + DMAXLV + 2) * (6 + DFPP(0))
      End If
      Return

58535 ' Check Report Width
      RPWIDTH = RWIDTH + PAGE.IMARGIN + PAGE.OMARGIN
      PWIDTH = BWIDTH + PAGE.IMARGIN + PAGE.OMARGIN
      If (PWIDTH < RPWIDTH) Then
         If (PRT.OPT = 2) Then
            SysVAR.NVAR1 = PWIDTH: SysVAR.NVAR2 = RPWIDTH
            Call PutMSG("XXXXM009")
            If CA$ <> SNGLKEY$(1) Then OKAY = 0: GoTo 58560
         End If
         BWIDTH = RWIDTH: PAGE.WIDTH = RPWIDTH
         If Len(PLINE$) < RPWIDTH Then PLINE$ = Space$(RPWIDTH)
      End If
      If VRTCL Then OKAY = -1: GoTo 58560

58540 ' Adjust to Fit
      If Not (RWIDTH < BWIDTH) Then GoTo 58550
      GAPS = 0: FX = 2
      While FX <= PAGE.DFCNT
          X = DFNUM(FX)
          If (X > 0) And (RFSEP(X) > 0) Then
             GAPS = GAPS + 1
          End If
          FX = FX + 1
      Wend
      GAP = (BWIDTH - RWIDTH) \ GAPS
      GAP1 = (BWIDTH - RWIDTH) - GAPS * GAP
      If GAP >= (MAXGAP - MINGAP) Then GAP = MAXGAP - MINGAP: GAP1 = 0
      FX = 2
      While FX <= PAGE.DFCNT
          X = DFNUM(FX): FSEP = RFSEP(X)
          If (X > 0) And (FSEP > 0) Then
             RFSEP(X) = FSEP + GAP - (GAP1 > 0): GAP1 = GAP1 - 1
          End If
          FX = FX + 1
      Wend
      GoSub 58530

58550 ' Determine Field Print Positions
      FX = 0
      If RWIDTH < BWIDTH Then PP = (BWIDTH - RWIDTH) \ 2 Else PP = 0
      While FX < PAGE.DFCNT
          FX = FX + 1: X = DFNUM(FX): FSEP = RFSEP(X)
          If (X > 0) And (FSEP > 0) Then
             FL = RFLTH(X): HL = Len(RFLBL$(X))
             z = FL: If z < HL Then z = HL
             PP = PP + FSEP + (z + 1) \ 2 - (FL + 1) \ 2
             DFPP(FX) = PP
             PP = PP + z
          End If
      Wend
      OKAY = -1

58560 ' Return to Calling Program
      PAGE.BWIDTH = BWIDTH
End Sub

Sub BldSumTITLE()
'SHARED DFNUM(), DFPP(), VRTCL
      AMAXLV = Asc(WFHDR.AMAXLV): DMAXLV = Asc(WFHDR.DMAXLV)
      RELSUMOPT = (AMAXLV > 0) And (DMAXLV > 0)
     ' Build Summary Report Title Line
      TL$ = Space$(Len(PLINE$) + 20): Y = 0
      If Not VRTCL Then
         For FX = 1 To PAGE.DFCNT: X = DFNUM(FX): PP = DFPP(FX)
             If (X > 0) And (PP > 0) Then
                FL = RFLTH(X): HL = Len(RFLBL$(X))
                If HL > 0 Then
                   PP = PP + (FL + 1) \ 2 - (HL + 1) \ 2
                   If PP > 0 Then Mid$(TL$, PP, HL) = RFLBL$(X)
                End If
             End If
         Next FX
         Else
         LSet TL$ = RFLBL$(1)
         TL2$ = Space$(PAGE.BWIDTH)
         PP = RFLTH(0) + DFPP(0)
         X = AMAXLV + DMAXLV
         Y = 0: xY$ = Space$(6): If AMAXLV > 0 Then z = -1 Else z = 1
         While X > -1
            RSet xY$ = Str$(Y * z)
            Mid$(TL$, PP, 6) = xY$
            Mid$(TL2$, PP, 6) = String$(6, "-")
            PP = PP + 6 + DFPP(0): Y = Y + 1: X = X - 1
         Wend
         Call GetRFLD("08001", FX)
         If FOUND Then
            RDFFld = RFLDS(FX)
            O = Asc(RDFFld.ALTH) + Asc(RDFFld.LLTH): L = Asc(RDFFld.HLTH)
            RSet xY$ = MID$(RFLD$(FX), O + 1, L)
            Mid$(TL$, PP, 6) = xY$
            Mid$(TL2$, PP, 6) = String$(6, "-")
         End If
      End If
      X = Len(RTrim$(TL$))
      If X > 0 Then
         HTFLMAX = HTFLMAX + 1
         ReDim Preserve HTFLINE$(HTFLMAX)
         HTFLINE$(HTFLMAX) = "T00" + RTrim$(TL$)
         If VRTCL Then
            HTFLMAX = HTFLMAX + 1
            ReDim Preserve HTFLINE$(HTFLMAX)
            HTFLINE$(HTFLMAX) = "T00" + RTrim$(TL2$)
         End If
      End If
      OKAY = -1
End Sub

Function BLOODLINE$(OPT)
   DGL = WFGLREC.DGL
   X = DGL
   If OPT = 1 Then ReDim GBL(DGL), PARPTR&(DGL)
   While X > 0
      GBL(X) = Asc(WFRELREC.CHNUM)
      If WFRELREC.PARPTR <> PARPTR&(X) Then
         Call RWRKGet("PTRREC", RC, WFRELREC.PARPTR)
         PARPTR&(X) = WFRELREC.PARPTR: X = X - 1
         Else
         X = 0
      End If
   Wend
   If DGL > 0 Then
      xx$ = Space$(4 * DGL + 2)
      If OPT = 1 Then Mid$(xx$, 1, 1) = "(": w = 2 Else w = 1
      While X < DGL
         X = X + 1: z = GBL(X) And 255
         If OPT = 1 Then
            Mid$(xx$, w, 3) = LTrim$(Str$(z))
            w = InStr(w, xx$, " ")
            Else
            Mid$(xx$, w, 2) = SFILL$(z, 2)
            w = w + 2
         End If
         z = GBL(X) \ 256
         If z > 0 Then
            Mid$(xx$, w, 1) = MID$("*?", z + 1, 1)
            If OPT = 1 Then w = w + 1
         End If
         If OPT = 1 Then Mid$(xx$, w, 1) = ","
         w = w + 1
      Wend
      If OPT = 1 Then Mid$(xx$, w - 1, 1) = ")"
   End If
   BLOODLINE$ = RTrim$(xx$)

End Function

Sub GetNextMARRIAGE()
'SHARED SPRESEQ, MLIST(), SPCNT, SPRNO, RELID
   If SPRESEQ Then
      If SPCNT > 0 Then SPCNT = SPCNT - 1
      SPRNO = MLIST(SPCNT)
      Else
      If RELID = FF3SPOUSE.SP1ID Then
         SPRNO = FF3SPOUSE.SP1NXT
         Else
         If RELID = FF3SPOUSE.SP2ID Then
            SPRNO = FF3SPOUSE.SP2NXT
            Else
            SPRNO = 0
         End If
      End If
   End If

End Sub

Sub XDETReport(WTYPE)
'SHARED DFNUM(), DFPP()
'SHARED RELID, PRPAGE(), PRLINE(), RX()
'SHARED TLBY, TLEY
'SHARED SPRESEQ, SPCNT, SPRNO, MAXMarry, MLIST()
     BD$ = Space$(12): DD$ = Space$(12): AGE$ = "   ": ANVSRC$ = " "
     BPL$ = Space$(41): MPL$ = BPL$: DPL$ = BPL$
     D1D$ = BD$: D2D$ = BD$: D1PL$ = BPL$: D2PL$ = BPL$: AGE2$ = AGE$
     NUMED$ = Space$(6)
     Select Case WTYPE
        Case 1: XRPT$ = "ARPT"
        Case 2: XRPT$ = "DRPT"
        Case 3: XRPT$ = "RRPT"
     End Select
     Call RptOPEN(XRPT$): If Not OKAY Then GoTo 1040
     WOPEN$ = "READQ"
1010 ' Verify Files by Opening
     Call FAMOpen("READ"): If Not OKAY Then GoTo 1040
     Call RWRKOpen(WOPEN$, WTYPE, 0): If OKAY Then GoTo 1050
     If WOPEN$ = "READQ" Then
        Call FamCLOSE
        LSet WOPEN$ = "READ": GoTo 1010
     End If

1040 ' Bad Return
     Close: ENV.FFOPEN = 0: Exit Sub

1050 ' Get Print Time Options
     DMAXLV = Asc(WFHDR.DMAXLV)
     Call RWRKGet("BASE", RC, PTR&): If RC <> 1 Then GoTo 1040
     SysVAR.SUBJECT = FMTNAME$(1)
     RptOPTION.OUTLINE = 0
     If DMAXLV = 0 Then
        GENLIST = -1
        Else
        Call PutMSG("RELRM011"): If A = 27 Then GoTo 1040
        GENLIST = (CA$ = "1")
        If Not GENLIST Then
           Call PutMSG("RELRM012"): If A = 27 Then GoTo 1040
           RptOPTION.OUTLINE = (CA$ <> "2")
        End If
     End If

     Call PutMSG("RELRM013"): If A = 27 Then GoTo 1040
     RptOPTION.FREEFORM = (CA$ = "2"): RptOPTION.TIMELINE = (CA$ = "3")

1060 If RptOPTION.TIMELINE Then
        'Call FmtFIND("XXXXS030")
        'Call MsgFORMAT: If A = 27 Then GoTo 1040
        'Call FmtFindFLD("TLBY"): If FOUND Then TLBY = Val(UTXT$(FFLD.UTXT))
        'Call FmtFindFLD("TLEY"): If FOUND Then TLEY = Val(UTXT$(FFLD.UTXT))
        Call frmFHSWMsgD.GetTLYears(TLBY, TLEY)
        If A = 27 Then GoTo 1040
        If (TLEY - TLBY) < 1 Then Call ErrBEEP(0): GoTo 1060
     End If

1200 ' Setup Report Tables
     Call BldDetDEF: If Not OKAY Then GoTo 1040
     PRDAGE = (RptOPTION.PRAGE = 2)
     PRID2 = ((RptOPTION.PRID And 2) > 0)
     RptOPTION.PRID = (RptOPTION.PRID > 0)
     'X = Len(PLINE$)
     Call PrintOPEN: If Not OKAY Then GoTo 1040
     'PLINE$ = Space$(X)
     PL2$ = PLINE$
     WFHDR.WFRULES = Chr$((Asc(WFHDR.WFRULES) And 252) - (GENLIST <> 0) - 2 * (GENLIST = 0))
     xxRGL$ = Space$(RFLTH(24)): xLINEAGE$ = Space$(RFLTH(2))
     XRELATION$ = Space$(RFLTH(3)): NFMT$ = Space$(RFLTH(5))
     AGE$ = Space$(RFLTH(11)): xAGE$ = AGE$: ANNV$ = Space$(RFLTH(18))
     Call BldDetTITLE
     Call GetRVAR("SPOUSE", X): If FOUND Then SPX$ = RVAR$(X)
     Call GetRVAR("YEARS", X): If FOUND Then YEARS$ = RVAR$(X)
     Call GetRVAR("ATAGE", X): If FOUND Then ATAGE$ = RVAR$(X)
     Call GetRVAR("UNKNOWN", X): If FOUND Then Unknown$ = RVAR$(X)
     Call GetRVAR("SEEPRIOR", X): If FOUND Then SeePrior$ = RVAR$(X)
     Call GetRVAR("LINE", X): If FOUND Then LINEX$ = RVAR$(X)
     ASPOUSE = ((RptOPTION.SPOUSE And 1) > 0)
     ALLSPOUSE = ((RptOPTION.SPOUSE And 2) > 0)
     If ALLSPOUSE Then Call RptOptGET("SP12", X): SPRESEQ = (X = 1)
     Call RptOptGET("EXAD", X): ExtAdoptRule = (X = 1)
     If Not RptOPTION.FREEFORM Then
        Call RptOptGET("PNML", PLNMLTH)
     End If

     If Not RptOPTION.TIMELINE Then GoTo 1400
     TMCHR$ = "<    >"
     Call RptOptGET("TLBL", X): TLBLCHR$ = Chr$(X)
     Call RptOptGET("TLLL", X): TLLLCHR$ = Chr$(X)
     Call RptOptGET("TLBD", X): TLBDCHR$ = Chr$(X)
     Call RptOptGET("TLMD", X): TLMDCHR$ = Chr$(X)
     Call RptOptGET("TLBC", X): TLBCCHR$ = Chr$(X)
     Call RptOptGET("TLDD", X): TLDDCHR$ = Chr$(X)
     Mid$(TMCHR$, 2, 2) = TLBDCHR$ + TLMDCHR$
     Mid$(TMCHR$, 4, 2) = TLBCCHR$ + TLDDCHR$
     TMLTH = RFLTH(23): TMLINE$ = Space$(TMLTH)
     TMFILL$ = String$(TMLTH, Asc(TLBLCHR$))
     TMLIFE$ = String$(TMLTH, Asc(TLLLCHR$))
     TMDELT# = (TLEY - TLBY) / TMLTH

1400 ' Begin with Base Record
     Call PutMSG("XXXXM015")
     ReDim PRPAGE(FF1Hdr.MAXID), PRLINE(FF1Hdr.MAXID)
     A = 0: LGL = -1: DUPID = 0
     PAGE.CSID = 0
     PLINEAGE# = 0: ANCID = 0
     Call RWRKGet("BASE", RC, PTR&): If RC <> 1 Then GoTo 1040
     GoTo 1450

1415 ' Process next entry on chart
     Call RWRKGet("NEXT", RC, PTR&): If (RC < 1) Or (RC > 8) Then GoTo 1490
     DUPID = (WFRELREC.id < 0)

1450 ' Process Workfile record
     RELID = Abs(WFRELREC.id)
     REPID = (PRPAGE(RELID) <> 0)
     AGL = -WFAGLREC.AGL: RGL = AGL + WFGLREC.DGL
     If RC < 3 Then
        PLINEAGE# = dLINEAGE#: PANCID = ANCID
        dLINEAGE# = WFANCREC.LINEAGE: ANCID = WFANCREC.ANCID
        ASTAT = WFANCREC.PSTAT And 3
        If ((DMAXLV > 0) Or (RptOPTION.BLNKSEP)) And (PAGE.LINE > 0) Then Call WriteLINE
        Else
        If ((RptOPTION.BLNKSEP) And (PAGE.LINE > 0)) Then Call WriteLINE
     End If
     ExtAdopt = ((WFRELREC.CSTAT And 5) > 0)
     PAGE.RPAGE = 0: PAGE.RLINE = 0
     PAGE.XID = RELID: PAGE.xType = 1
     RREFREC.AGL = Chr$(-AGL): RREFREC.DGL = Chr$(RGL - AGL)
     Call RREFAdd
     ComPTR = FF1Rec.COM
     Call GetBDEvents(BD$, BPL$, DD$, DPL$, AGE$)
     BDY = FF1Rec.BY: DDY = FF1Rec.DY
     If RptOPTION.SPOUSE Then
        PAGE.CSID = 0: MLOC = 0: MCOMPTR = 0
        SDY = 0: SDMD = 0: M2Y = 0: M2MD = 0
        S1DY = FF1Rec.DY: S1DMD = FF1Rec.DMD
        SPRNO = FF1Rec.SPOUSE
        If (ASPOUSE Or (DMAXLV = 0)) And (FF1Rec.SPOUSE <> 0) Then
           GoSub 1530
           Else
           If ALLSPOUSE Then Call BldMLIST
        End If
     End If
     If RptOPTION.Relation Then
        If (REPID) And (WFGLREC.DGL > 0) Then
           XID = RELID: GoSub 1550
           Else
           LSet XRELATION$ = Relation$("M", FF1Rec.SEX, AGL, RGL, RRELCD$())
           If ExtAdoptRule And ExtAdopt Then
              LSet XRELATION$ = "*" + XRELATION$
           End If
        End If
     End If
     GoSub 1800   ' Process DATA Fields
     If A = 27 Then GoTo 1490

     PRPAGE(RELID) = PAGE.RPAGE
     PRLINE(RELID) = PAGE.RLINE

     If Not ALLSPOUSE Then GoTo 1470  ' Didn't Request ALL Marriages
     If (DUPID) And (WFGLREC.DGL > 0) Then GoTo 1470
     If DMAXLV = 0 Then GoTo 1470     'no spouse info for ancestor rpt

     ' Process Marriage Records
     PAGE.xType = 3
     While SPRNO <> 0
        DUPID = 0: REPID = 0: PAGE.RPAGE = 0: PAGE.RLINE = 0
        GoSub 1530
        PAGE.XID = PAGE.CSID: PAGE.xType = 3: Call RREFAdd
        BDY = FF1Rec.BY: DDY = FF1Rec.DY
        LSet BD$ = D1D$: LSet BPL$ = D1PL$
        LSet DD$ = D2D$: LSet DPL$ = D2PL$: LSet AGE$ = AGE2$
        GoSub 1800
        If (PAGE.RPAGE > 0) And (PAGE.CSID > 0) Then
           PRPAGE(PAGE.CSID) = -(PAGE.RPAGE) * (1 + 2 * REPID)
           PRLINE(PAGE.CSID) = PAGE.RLINE
        End If
        If A = 27 Then SPRNO = 0
        SDY = 0: SDMD = 0
     Wend

1470 ' Finish Line and Go Get Another WF Record
     Call PrintLINE
     If A = 27 Then GoTo 1490
     GoTo 1415

1490 ' END OF REPORT
     Call ClearMSG
     Call RptCLOSE: If PRT.AGAIN Then GoTo 1400
     Call FamCLOSE
     ReDim PRPAGE(1), PRLINE(1)
     Exit Sub

1530 ' Get Information for SPRNO
     If DMAXLV = 0 Then GoSub 1540: If PAGE.CSID = 0 Then Return
     If (SPRNO <> 0) Or (PAGE.CSID > 0) Then
        If SPRNO <> 0 Then Call FF3GetRec(SPRNO)
        MCOMPTR = FF3SPOUSE.COM: MLOC = FF3SPOUSE.MLOC
        M1Y = FF3SPOUSE.D1Y: M1MD = FF3SPOUSE.D1MD
        M2Y = FF3SPOUSE.D2Y: M2MD = FF3SPOUSE.D2MD
        If FF3SPOUSE.SP1ID = RELID Then
           PAGE.CSID = FF3SPOUSE.SP2ID
           Else
           PAGE.CSID = FF3SPOUSE.SP1ID
        End If
        If ALLSPOUSE Then Call GetNextMARRIAGE Else SPRNO = 0
        If PAGE.CSID > 0 Then
           Call FF1GetRec(PAGE.CSID)
           If PAGE.xType = 3 Then ComPTR = FF1Rec.COM
           Call GetBDEvents(D1D$, D1PL$, D2D$, D2PL$, AGE2$)
           SDY = FF1Rec.DY: SDMD = FF1Rec.DMD
           If PAGE.xType = 1 Then
              Call FF1GetRec(RELID)
              Else
              REPID = (PRPAGE(PAGE.CSID) <> 0)
              If RptOPTION.Relation Then
                 If REPID Then
                    XID = PAGE.CSID: GoSub 1550
                    Else
                    LSet XRELATION$ = " "
                 End If
              End If
           End If
        End If
        If RptOPTION.PRAGE Then GoSub 1545
        If RptOPTION.PLACES Then
           If MLOC Then
              Call FF3GetRec(MLOC): LSet MPL$ = MPLACE
              Else
              LSet MPL$ = ""
           End If
        End If
     End If
     Return

1540 ' Locate Marriage of Ancestors
     PAGE.CSID = 0: MLOC = 0
     If (FF1Rec.SEX <> FGENDR$(2)) Or (PLINEAGE# <> (dLINEAGE# - 1#)) Then Return
     While SPRNO <> 0: Call FF3GetRec(SPRNO)
        If (ANCID = FF3SPOUSE.SP1ID) Then
           SPRNO = FF3SPOUSE.SP1NXT: PAGE.CSID = FF3SPOUSE.SP2ID
           Else
           SPRNO = FF3SPOUSE.SP2NXT: PAGE.CSID = FF3SPOUSE.SP1ID
        End If
        If PAGE.CSID = PANCID Then SPRNO = 0 Else PAGE.CSID = 0
     Wend
     Return

1545 ' Determine Last Anniversary
     If ((M2Y > 0 Or M2MD > 0) Or (S1DY = 0 And SDY = 0)) Then
        LSet ANVSRC$ = " "
        Else
        M2Y = SDY: M2MD = SDMD: LSet ANVSRC$ = "*"
        If ((S1DY > 0) And ((SDY = 0) Or ((S1DY < SDY) Or (S1DY = SDY And S1DMD < SDMD)))) Then M2Y = S1DY: M2MD = S1DMD
     End If
     RSet ANNV$ = DATEDif$(M1MD, M1Y, M2MD, M2Y)
     Return

1550      ' Format BackREF in XRELATION$
     PREF = Abs(PRPAGE(XID)): LREF = PRLINE(XID)
     LSet XRELATION$ = ">"
     Mid$(XRELATION$, 2, 4) = FNX$(PREF): L1 = InStr(XRELATION$, " ")
     Mid$(XRELATION$, L1 + 1, 4) = FNX$(LREF)
     Return

1800 ' PRINT Prefix Fields (Those Before NAME on Data Line)
     COMBGN = 0: PAGE.COFST = DFPP(PAGE.PFCNT + 1) ' Set Continuation Offset
     If PAGE.LOFST = 0 Then PAGE.LOFST = DFPP(1)
     FX = 0
     While FX < PAGE.PFCNT
        FX = FX + 1: DFN = DFNUM(FX): X = DFPP(FX)
        If (DFN > 0) And (X > 0) Then
           FLTH = RFLTH(DFN): PAGE.LOFST = X
           If DFN = 24 Then
              GoSub 1821
              Else
              On DFN GoSub 1821, 1822, 1823, 1824, 1825, 1826, 1826, 1826, 1826
           End If
        End If
     Wend

1805 ' Process DATA Items After NAME Field on DATA Line
     If ((Not REPID) Or (RptOPTION.Relation And ((WFGLREC.DGL > 0) Or (PAGE.CSID > 0)))) Then
        If RptOPTION.FREEFORM Then GoSub 2800 Else GoSub 2500
        Else
        PAGE.LOFST = DFPP(PAGE.PFCNT + 1): FLTH = RFLTH(5): GoSub 1825
        If RptOPTION.FREEFORM Then
           PAGE.LOFST = InStr(PAGE.LOFST, PLINE$, "   ") + 1
           Else
           PAGE.LOFST = DFPP(PAGE.PFCNT + 2)
        End If
        LSet CL$ = "(" + SeePrior$: X = InStr(1, CL$, "   ")
        If X > 0 Then
           PREF = Abs(PRPAGE(FF1Rec.RID)): LREF = PRLINE(FF1Rec.RID)
           Mid$(CL$, X + 1, 3) = FNX$(PREF): X = InStr(X, CL$, "   ")
           Mid$(CL$, X + 1, Len(LINEX$)) = LINEX$: X = InStr(X, CL$, "  ")
           Mid$(CL$, X + 1, 3) = FNX$(LREF): X = InStr(X, CL$, "   ")
           If X > 0 Then Mid$(CL$, X, 1) = ")"
        End If
        If RptOPTION.FREEFORM Then
           Call PutDATA
           Else
           If PAGE.LOFST > 0 Then Mid$(PLINE$, PAGE.LOFST, 255) = CL$
        End If
     End If
     If Not RptOPTION.FREEFORM Then
        Call PrintLINE
        If RptOPTION.TWOLINES Then
           LSet PLINE$ = PL2$
           Call PrintLINE: LSet PL2$ = " "
        End If
     End If

     ' Print Comments - Individual and Marriage as appropriate
     If Not (REPID Or DUPID) Then
        If RptOPTION.COMMENTS Then
           COMBGN = 0: PAGE.LWIDTH = PAGE.BWIDTH - PAGE.CINDENT
           If (PAGE.xType = 3) And MCOMPTR Then Call ComPRINT(MCOMPTR, COMBGN)
           If ComPTR Then Call ComPRINT(ComPTR, COMBGN)
           If (PAGE.xType = 1) And MCOMPTR Then Call ComPRINT(MCOMPTR, COMBGN)
        End If
     End If
     Call PrintLINE
     Return

1820 ' Routines for Processing PREFIX Data Fields
1821 ' AGL, DGL or RGL
     If (PAGE.xType = 1) And (RGL <> LGL) Then
        RSet xxRGL$ = Str$(RGL)
        Mid$(PLINE$, PAGE.LOFST, FLTH) = xxRGL$
        LGL = RGL
     End If
     GoTo 1829
1822 ' Lineage Number
     If (PAGE.xType = 1) And (RC < 3) Then
        RSet xLINEAGE$ = Str$(dLINEAGE#)
        Mid$(PLINE$, PAGE.LOFST, FLTH) = xLINEAGE$
        Mid$(PLINE$, PAGE.LOFST + FLTH, 1) = Chr$(32 + (32 - Asc("*")) * ((ASTAT And 1) > 0) + (32 - Asc("?")) * ((ASTAT And 2) > 0))
     End If
     GoTo 1829
1823 ' Relationship
     Mid$(PLINE$, PAGE.LOFST, FLTH) = XRELATION$
     LSet XRELATION$ = " "
     GoTo 1829
1824 ' Format BLOODLINE on PLINE$
     If (PAGE.xType = 1) Or RptOPTION.OUTLINE Then
        L = 1: X = PAGE.LOFST
        While (GBL(L) <> 0) And (L <= WFGLREC.DGL)
           If PAGE.xType = 1 Then
              C = GBL(L) And 255: S = GBL(L) \ 256
              If (RptOPTION.OUTLINE And L < WFGLREC.DGL) Then
                 Mid$(PLINE$, X + 1, 1) = "."
                 Else
                 RSet NUMED$ = Str$(C)
                 Mid$(PLINE$, X, 2) = Right$(NUMED$, 2)
                 Mid$(PLINE$, X + 2, 1) = Chr$(32 + (32 - Asc("*")) * ((S And 1) > 0) + (32 - Asc("?")) * ((S And 2) > 0))
              End If
           End If
           X = X + 3: L = L + 1
        Wend
     End If
     GoTo 1829
1825 ' Name - and Outline Child Numbers
     NMLTH = FLTH
     If RptOPTION.OUTLINE Then
        GoSub 1824
        If Not RptOPTION.FREEFORM Then NMLTH = FLTH - (X - PAGE.LOFST)
        PAGE.LOFST = X
     End If
     If PAGE.xType = 3 Then
        Mid$(PLINE$, PAGE.LOFST, 255) = SPX$
        X = InStr(PAGE.LOFST, PLINE$, "  ") + 1
        If Not RptOPTION.FREEFORM Then NMLTH = FLTH - (X - PAGE.LOFST)
        PAGE.LOFST = X
     End If
     If (PAGE.xType = 1) Or (PAGE.CSID > 0) Then
        LSet NFMT$ = FMTNAME$(0)
        Else
        LSet NFMT$ = "(" + Unknown$ + ")"
     End If
     Mid$(PLINE$, PAGE.LOFST, NMLTH) = NFMT$
     PAGE.COFST = PAGE.LOFST
     If PRID2 Then GoSub 1828
     If RptOPTION.FREEFORM Then
        PAGE.LOFST = InStr(PAGE.LOFST, PLINE$, "   ")
     End If
     GoTo 1829
1826 ' ID Numbers - First Data Line
     Select Case DFN
        Case 6: Mid$(PLINE$, PAGE.LOFST, 5) = SFILL$(FF1Rec.RID, 5)
        Case 7: Mid$(PLINE$, PAGE.LOFST, 5) = SFILL$(FF1Rec.FID, 5)
                Mid$(PLINE$, PAGE.LOFST + 5, 1) = FDSTAT$("FID")
        Case 8: Mid$(PLINE$, PAGE.LOFST, 5) = SFILL$(FF1Rec.MID, 5)
                Mid$(PLINE$, PAGE.LOFST + 5, 1) = FDSTAT$("MID")
        Case 9: If PAGE.CSID > 0 Then Mid$(PLINE$, PAGE.LOFST, 5) = SFILL$(PAGE.CSID, 5)
     End Select
     GoTo 1829
1828 ' ID Numbers - Second Data Line
     X = PAGE.LOFST
     Mid$(PL2$, X, 255) = RFLBL$(6) + "="
     X = InStr(X, PL2$, "   ")
     Mid$(PL2$, X, 5) = SFILL$(FF1Rec.RID, 5)
     Mid$(PL2$, X + 7, 255) = RFLBL$(7) + "="
     X = InStr(X + 6, PL2$, "   ")
     Mid$(PL2$, X, 5) = SFILL$(FF1Rec.FID, 5)
     Mid$(PL2$, X + 5, 1) = FDSTAT$("FID")
     X = InStr(X + 5, PL2$, "   ")
     Mid$(PL2$, X + 2, 255) = RFLBL$(8) + "="
     X = InStr(X, PL2$, "   ")
     Mid$(PL2$, X, 5) = SFILL$(FF1Rec.MID, 5)
     Mid$(PL2$, X + 5, 1) = FDSTAT$("MID")
     GoTo 1829
1829 ' Return from DATA Item
     Return

2500 ' Print Information for Fixed Format Report - Fields After NAME
     FX = PAGE.PFCNT
     While FX < PAGE.DFCNT
        FX = FX + 1: DFN = DFNUM(FX): Y = DFPP(FX)
        If (DFN > 0) And (Y > 0) Then
           FLTH = RFLTH(DFN): PAGE.LOFST = Y
           On DFN - 4 GoSub 1825, 1826, 1826, 1826, 1826, 2501, 2502, 2510, 2515, 2510, 2520, 2525, 2520, 2521, 2530, 2535, 2530, 2600, 2610
           PAGE.LOFST = PAGE.LOFST + RFLTH(DFN)
        End If
     Wend
     Return

2501 ' Gender
     If (PAGE.xType = 1) Or (PAGE.CSID > 0) Then
        If FF1Rec.SEX = FGENDR$(1) Then
           X = 1
           Else
           If FF1Rec.SEX = FGENDR$(2) Then
              X = 2
              Else
              X = 3
           End If
        End If
        Mid$(PLINE$, PAGE.LOFST, 1) = MID$(RGENDR$(X), 2, 1)
     End If
     Return
2502 ' Age
     If (PAGE.xType = 1) Or (PAGE.CSID > 0) Then
        If (Not (PRDAGE And (DDY = 0))) Then
           Mid$(PLINE$, PAGE.LOFST, 10) = AGE$
        End If
     End If
     Return
2510 ' Birth Date
     If (PAGE.xType = 1) Or (PAGE.CSID > 0) Then
        If FF1Rec.BY > 0 Then
           Mid$(PLINE$, PAGE.LOFST, 12) = BD$
        End If
        If RptOPTION.TWOLINES Then GoTo 2515
     End If
     Return
2515 ' Birth Place
     If (PAGE.xType = 1) Or (PAGE.CSID > 0) Then
        If RptOPTION.TWOLINES Then
           Mid$(PL2$, PAGE.LOFST, PLNMLTH) = BPL$
           Else
           Mid$(PLINE$, PAGE.LOFST, PLNMLTH) = BPL$
        End If
     End If
     Return
2520 ' Marriage Date
     If (PAGE.xType = 3) Or (PAGE.CSID > 0) Then
        If M1Y > 0 Then
           Mid$(PLINE$, PAGE.LOFST, 11) = XRPTDate$(M1MD, M1Y, 1)
           If RptOPTION.PRAGE And RptOPTION.TWOLINES Then
              If (Not PRDAGE) Or (M2Y > 0) Then
                 Mid$(PLINE$, PAGE.LOFST + 12, 3) = ANNV$
                 Mid$(PLINE$, PAGE.LOFST + 15, 1) = ANVSRC$
                 Mid$(PLINE$, PAGE.LOFST + 17, 5) = YEARS$
              End If
           End If
        End If
        If RptOPTION.TWOLINES Then GoTo 2525
     End If
     Return
2521 ' Last Anniversary
     If (PAGE.xType = 3) Or (PAGE.CSID > 0) Then
        If (M1Y > 0) And (Not (PRDAGE And (M2Y = 0))) Then
           Mid$(PLINE$, PAGE.LOFST, 10) = ANNV$ + ANVSRC$
        End If
     End If
     Return
2525 ' Marriage PLACE
     If MLOC <> 0 Then
        If RptOPTION.TWOLINES Then
           Mid$(PL2$, PAGE.LOFST, PLNMLTH) = MPL$
           Else
           Mid$(PLINE$, PAGE.LOFST, PLNMLTH) = MPL$
        End If
     End If
     Return
2530 ' Death Date
     If (PAGE.xType = 1) Or (PAGE.CSID > 0) Then
        If (FF1Rec.DY > 0) And (FF1Rec.DY < 9999) Then
           Mid$(PLINE$, PAGE.LOFST, 12) = DD$
        End If
        If RptOPTION.TWOLINES Then GoTo 2535
     End If
     Return
2535 ' Death Place
     If (PAGE.xType = 1) Or (PAGE.CSID > 0) Then
        If RptOPTION.TWOLINES Then
           Mid$(PL2$, PAGE.LOFST, PLNMLTH) = DPL$
           Else
           Mid$(PLINE$, PAGE.LOFST, PLNMLTH) = DPL$
        End If
     End If
     Return

2600 ' Build Timeline
     If BDY > 0 Then Mid$(PLINE$, PAGE.LOFST, 4) = SFILL$(BDY, 4)
     Mid$(PLINE$, PAGE.LOFST + 4, 1) = "-"
     If (DDY = 9999) Then
        Mid$(PLINE$, PAGE.LOFST + 5, 4) = " ???"
        Else
        If (DDY > 0) Then Mid$(PLINE$, PAGE.LOFST + 5, 4) = SFILL$(DDY, 4)
     End If
     Return
2610 ' Build Timeline
     LSet TMLINE$ = TMFILL$
     TMRID = FF1Rec.RID: TMDYR = DDY
     TMYEAR = BDY: If TMYEAR > 0 Then TMTYPE = 1: GoSub 2650
     If MID$(TMCHR$, 4, 1) = " " Then GoTo 2630
     TMTYPE = 3: TMCHID = FF1Rec.OLDCH
     While TMCHID <> 0: Call FF1GetRec(TMCHID)
        Call GetBDEvents("", "", "", "", "")
        TMYEAR = FF1Rec.BY: If TMYEAR > 0 Then GoSub 2650
        If TMRID = FF1Rec.FID Then TMCHID = FF1Rec.FCH: Else TMCHID = FF1Rec.MCH
     Wend
     If TMRID <> FF1Rec.RID Then Call FF1GetRec(TMRID)
2630 If MID$(TMCHR$, 3, 1) = " " Then GoTo 2640
     TMTYPE = 2: TMSPRNO = FF1Rec.SPOUSE
     While TMSPRNO <> 0: Call FF3GetRec(TMSPRNO)
        TMYEAR = FF3SPOUSE.D1Y: If TMYEAR > 0 Then GoSub 2650
        If TMRID = FF3SPOUSE.SP1ID Then TMSPRNO = FF3SPOUSE.SP1NXT: Else TMSPRNO = FF3SPOUSE.SP2NXT
     Wend
2640 TMYEAR = TMDYR
     If (TMYEAR > 0) And (TMYEAR < 9999) Then TMTYPE = 4: GoSub 2650
     Mid$(PLINE$, PAGE.LOFST, TMLTH) = TMLINE$
     Return

2650 ' Process Event
     If TMYEAR > TLEY Then
        TMSYM = 5: TMX = TMLTH
        Else
        TMSYM = TMTYPE
        TMX = (TMYEAR - TLBY) / TMDELT#: If TMX < 0 Then TMX = 1: TMSYM = 0
        If TMX < 1 Then
           TMX = 1
           Else
           If TMX > TMLTH Then TMSYM = 5: TMX = TMLTH
        End If
     End If
     Mid$(TMLINE$, TMX, 1) = MID$(TMCHR$, TMSYM + 1, 1)
     If (TMTYPE = 1) And (TMX < TMLTH) And (TMDYR < 9999) Then
        Mid$(TMLINE$, TMX + 1, TMLTH) = TMLIFE$
     End If
     If (TMTYPE = 4) And (TMX < TMLTH) Then
        Mid$(TMLINE$, TMX + 1, TMLTH) = TMFILL$
     End If
     Return

2800 ' Print Information for Free Format Report
     'IF (RPTOPTION.PRID AND RPTOPTION.COMMENTS THEN
     '   PAGE.LOFST = PAGE.LOFST + 1
     '   LSET CL$ = "(" + RFLBL$(6): X = INSTR(1, CL$, "   ")
     '   MID$(CL$, X, 5) = FNX$(FF1REC.RID): X = INSTR(X, CL$, "   ")
     '   MID$(CL$, X, 255) = "," + RFLBL$(7): X = INSTR(X, CL$, "   ")
     '   MID$(CL$, X, 5) = FNX$(FF1REC.FID): X = INSTR(X, CL$, "   ")
     '   MID$(CL$, X, 1) = FDSTAT$("FID"): X = INSTR(X, CL$, "   ")
     '   MID$(CL$, X, 255) = "," + RFLBL$(8): X = INSTR(X, CL$, "   ")
     '   MID$(CL$, X, 5) = FNX$(FF1REC.MID): X = INSTR(X, CL$, "   ")
     '   MID$(CL$, X, 1) = FDSTAT$("MID"): X = INSTR(X, CL$, "   ")
     '   MID$(CL$, X, 1) = ")"
     '   CALL PutDATA
     'END IF
     PAGE.LOFST = DFPP(PAGE.PFCNT + 1): FLTH = RFLTH(5): GoSub 1825
     PAGE.LOFST = InStr(PAGE.LOFST, PLINE$, "   ")
     If (PAGE.xType = 1) Or (PAGE.CSID > 0) Then
        If RptOPTION.PRSEX Then
           If FF1Rec.SEX = FGENDR$(1) Then
              X = 1
              Else
              If FF1Rec.SEX = FGENDR$(2) Then
                 X = 2
                 Else
                 X = 3
              End If
           End If
           LSet CL$ = ", " + MID$(RGENDR$(X), 3, 10)
           Call PutDATA
        End If
     End If

     ' AGE - if Still Living
     If (PAGE.xType = 1) Or (PAGE.CSID > 0) Then
        If (RptOPTION.PRAGE) And (Not PRDAGE) And (DDY = 0) Then
           LSet CL$ = ", " + RFLBL$(11) + " " + LTrim$(AGE$)
           Call PutDATA
        End If
     End If
  
     ' Birth Date/Place
     If (PAGE.xType = 1) Or (PAGE.CSID > 0) Then
        If RptOPTION.PLACES Then PLTH = Len(RTrim$(BPL$))
        If ((BD$ <> Space$(12)) And RptOPTION.DateS) Or (PLTH > 0) Then
           LSet CL$ = ", " + RFLBL$(12): X = 1
           If RptOPTION.DateS And (BD$ <> Space$(12)) Then
              X = InStr(X, CL$, "   ")
              Mid$(CL$, X + 1, 12) = LTrim$(BD$)
           End If
           If PLTH > 0 Then
              X = InStr(X, CL$, "   ")
              Mid$(CL$, X + 1, 255) = RFLBL$(13) + " " + BPL$
           End If
           Call PutDATA
        End If
     End If

     ' Marriage DATE/Place/Anniv
     If (PAGE.xType = 3) Or (PAGE.CSID > 0) Then
        If RptOPTION.PLACES Then PLTH = Len(RTrim$(MPL$))
        If (((M1Y <> 0) Or (M1MD <> 0)) And RptOPTION.DateS) Or (PLTH > 0) Then
           LSet CL$ = ", " + RFLBL$(15): X = 1
           If RptOPTION.DateS And ((M1Y > 0) Or (M1MD > 0)) Then
              X = InStr(X, CL$, "   ")
              Mid$(CL$, X + 1, 11) = XRPTDate$(M1MD, M1Y, 0)
           End If
           If PLTH > 0 Then
              X = InStr(X, CL$, "    ")
              Mid$(CL$, X + 1, 255) = RFLBL$(16) + " " + MPL$
           End If
           If (M1Y > 0) And (RptOPTION.PRAGE) Then
              If Not (PRDAGE And (M2Y = 0)) Then
                 X = InStr(X, CL$, "  ")
                 Mid$(CL$, X + 1, 255) = RFLBL$(18) + " " + LTrim$(ANNV$)
                 X = InStr(X, CL$, "  ")
                 Mid$(CL$, X, 1) = ANVSRC$
                 X = InStr(X, CL$, "  ")
                 Mid$(CL$, X + 1, 255) = YEARS$
              End If
           End If
           Call PutDATA
        End If
     End If

     ' Death DATE/Place/Age
     If (PAGE.xType = 1) Or (PAGE.CSID > 0) Then
        If RptOPTION.PLACES Then PLTH = Len(RTrim$(DPL$))
        If (RptOPTION.DateS And (DD$ <> Space$(12))) Or (PLTH > 0) Then
           LSet CL$ = ", " + RFLBL$(19): X = 1
           If RptOPTION.DateS And (DD$ <> Space$(12)) Then
              X = InStr(X, CL$, "  ")
              Mid$(CL$, X + 1, 12) = LTrim$(DD$)
           End If
           If PLTH > 0 Then
              X = InStr(X, CL$, "  ")
              Mid$(CL$, X + 1, 255) = RFLBL$(20) + " " + DPL$
           End If
           If (RptOPTION.PRAGE) And (AGE$ <> "   ") Then
              X = InStr(X, CL$, "  ")
              Mid$(CL$, X + 1, 255) = ATAGE$ + " " + LTrim$(AGE$)
           End If
           Call PutDATA
        End If
     End If
     Return

End Sub

Sub XDUPReport(WTYPE)
     Select Case WTYPE
        Case 1: DRPT$ = "ADUP"
        Case 2: DRPT$ = "DDUP"
     End Select
     Call RptOPEN(DRPT$): If Not OKAY Then GoTo S2690
     WOPEN$ = "READQ"
S2610: ' Verify Files by Opening
     Call FAMOpen("READ"): If Not OKAY Then GoTo S2690
     Call RWRKOpen(WOPEN$, WTYPE, 0): If OKAY Then GoTo S2615
     If WOPEN$ = "READQ" Then
        Call FamCLOSE
        LSet WOPEN$ = "READ": GoTo S2610
        Else
        GoTo S2690
     End If
S2615: ' Get BASE Record
     Call PutMSG("RELRM016")
     Call RWRKGet("BASE", RC, PTR&): If RC <> 1 Then GoTo S2690
     DupCNT1 = 0: DupCNT2 = 0: DUPMAX = 100
     ReDim DUPID(FF1Hdr.MAXID), DUPPTR&(DUPMAX)
S2620: ' Check for Duplicate Record
     If (WFRELREC.id > 0) And (WFRELREC.NREF <> 0) Then
        DupCNT1 = DupCNT1 + 1
        If DupCNT1 > DUPMAX Then
           DUPMAX = DUPMAX + 20
           ReDim Preserve DUPPTR&(DUPMAX)
        End If
        DUPID(WFRELREC.id) = DupCNT1
        DUPPTR&(DupCNT1) = PTR&
     End If
     If (WFRELREC.id < 0) Or (WFRELREC.NREF <> 0) And (DupCNT2 < 32756) Then DupCNT2 = DupCNT2 + 1
     Call RWRKGet("WRKREC", RC, PTR&)
     If (0 < RC) And (RC < 9) Then GoTo S2620
     SysVAR.NVAR1 = DupCNT1: SysVAR.NVAR2 = DupCNT2
     If SysVAR.NVAR1 = 0 Then
        Call PutMSG("RELRM017"): GoTo S2690
     End If
     Call PrintOPEN: If Not OKAY Then GoTo S2690

S2650: ' Print Report of Duplicate Ancestors/Descendants
     Call BldDupDEF
     SysVAR.SUBJECT = FMTNAME$(1)
     Call PutMSG("XXXXM015")
     x1 = Len(RFLBL$(1)): X2 = Len(RFLBL$(2))
     X3 = Len(RFLBL$(3)): L3 = RFLTH(3): X4 = Len(RFLBL$(4))
     LSet CL$ = Space$(4 * Asc(WFHDR.DMAXLV) + 2)
     PAGE.COFST = x1 + X3 + X4 + L3 + 13
     Call RWRKGet("BASE", RC, PTR&): If RC <> 1 Then GoTo S2690
SysVAR.NVAR1 = DupCNT1
     For id = 1 To FF1Hdr.MAXID
         DUPIX = DUPID(id)
         If DUPIX > 0 Then
            RELPTR& = DUPPTR&(DUPIX)
            Call FF1GetRec(id)
            PAGE.RPAGE = 0: PAGE.RLINE = 0
            PAGE.XID = id: PAGE.xType = 1: Call RREFAdd
            LSet PLINE$ = RFLBL$(1)
            Mid$(PLINE$, x1 + 1, 6) = Str$(id)
            Mid$(PLINE$, x1 + 8, X2) = RFLBL$(2)
            Mid$(PLINE$, x1 + X2 + 9, 52) = FMTNAME$(0)
            Call PrintLINE: If A = 27 Then RELPTR& = 0
            While RELPTR& <> 0:
               Call RWRKGet("RELREC", RC, RELPTR&)
               RELPTR& = WFRELREC.NREF
               AGL = WFGLREC.AGL: DGL = WFGLREC.DGL: RGL = AGL + DGL
               Mid$(PLINE$, x1 + 10, X3) = RFLBL$(3)
               Mid$(PLINE$, x1 + X3 + 11, L3) = Relation$("M", FF1Rec.SEX, AGL, RGL, RRELCD$())
               Mid$(PLINE$, x1 + X3 + L3 + 12, X4) = RFLBL$(4)
               X = PAGE.COFST
               If AGL <> 0 Then
                  Mid$(PLINE$, X, 255) = LTrim$(Str$(WFANCREC.LINEAGE))
                  X = InStr(X, PLINE$, " ")
               End If
               If DGL <> 0 Then
                  LSet CL$ = BLOODLINE$(1): Call PutDATA
               End If
               Call PrintLINE: If A = 27 Then RELPTR& = 0
            Wend: If A = 27 Then id = FF1Hdr.MAXID
         End If
     Next id
     Call RptCLOSE: If PRT.AGAIN Then GoTo S2650
S2690: ' Return to Main Program
     Close: ENV.FFOPEN = 0

End Sub

Sub XSUMReport(WTYPE)
'SHARED DFNUM(), DFPP()
'SHARED DATECNT, DATERNG, PLACCNT, MARRCNT, RWIDTH, VRTCL
'SHARED SUMBASE AS SUMREC, SUM AS SUMREC, TOT AS SUMREC
     ' Produce Generation Summary Report
     RELSUMOPT = (WTYPE = 3): GENSUMOPT = (WTYPE < 3)
     A = 0
     If GENSUMOPT Then
        Call PutMSG("XXXXM044")
        If A = 27 Then GoTo S1040
        VRTCL = (CA$ = "2")
     End If
     Select Case WTYPE
        Case 1: SRPT$ = "ASUM"
        Case 2: SRPT$ = "DSUM"
        Case 3: SRPT$ = "RSUM"
     End Select
     Call RptOPEN(SRPT$): If Not OKAY Then GoTo S1040
     WOPEN$ = "READQ"
S1010: ' Verify Files by Opening
     Call FAMOpen("READ"): If Not OKAY Then GoTo S1040
     Call RWRKOpen(WOPEN$, WTYPE, 0): If OKAY Then GoTo S1050
     If WOPEN$ = "READQ" Then
        Call FamCLOSE
        LSet WOPEN$ = "READ": GoTo S1010
     End If

S1040: ' Bad Return
     Close: ENV.FFOPEN = 0: Exit Sub

S1050: ' Prepare to Print
     On Local Error GoTo S1500
     Call RWRKGet("BASE", RC, PTR&): If RC <> 1 Then GoTo S1040
     SysVAR.SUBJECT = FMTNAME$(1)
     If DMAXLV = 0 Then
        Call RptOptGET("CNOP", X): COUNTNOP = (X = 1)
        If Not VRTCL Then Call RptOptGET("LNOP", X): LISTNODES = (X = 1)
        LALIN = Len(Str$(2 ^ Asc(WFHDR.AMAXLV))): zALIN$ = Space$(LALIN)
     End If
     Call RptOptGET("CDUP", X): COUNTDUPS = (X = 1)
     Call RptOptGET("CLIN", X): COUNTCLINE = (X = 1)
     Call RptOptGET("DTCT", X): COUNTDATES = (X = 1)
     Call RptOptGET("DTRG", X): SHOWRANGE = (X = 1)
     Call RptOptGET("PLCT", X): COUNTPLACES = (X = 1)
     If COUNTPLACES Then BPL$ = " ": DPL$ = BPL$: Else BPL$ = "": DPL$ = ""
     Call RptOptGET("MCNT", X): COUNTMARRY = (X = 1)
     ShowREL = RptOPTION.Relation
     If RELSUMOPT Then
        LORGL = -Asc(WFHDR.AMAXLV): HIRGL = Asc(WFHDR.HIRGL)
        ReDim RELCNT(LORGL To 0, LORGL To HIRGL)
        Else
        HIRGL = Asc(WFHDR.AMAXLV) + Asc(WFHDR.DMAXLV)
        If VRTCL Then ReDim SUMTAB(HIRGL) As SUMREC
     End If
     If Not VRTCL Then
        Call GetRFLD("08001", FX)
        If FOUND Then
           RDFFld = RFLDS(FX)
           O = Asc(RDFFld.ALTH): L = Asc(RDFFld.LLTH)
           TOTAL$ = MID$(RFLD$(FX), O + 1, L)
        End If
        If RELSUMOPT Then
           Call GetRFLD("08008", FX)
           If FOUND Then
              RDFFld = RFLDS(FX)
              O = Asc(RDFFld.ALTH): L = Asc(RDFFld.LLTH)
              ACCUM$ = MID$(RFLD$(FX), O + 1, L)
           End If
        End If
     End If

S1200: ' Setup Report Tables & Open Files
     Call BldSumDEF: If Not OKAY Then GoTo S1040
     SPL = Len(PLINE$)
     If Not RELSUMOPT Then
        Call PrintOPEN: If Not OKAY Then GoTo S1040
        PLINE$ = Space$(SPL)
        Call BldSumTITLE
        ' IF Ancestor or Descendant Summary, Retrieve Records by Generation
        WFHDR.WFRULES = Chr$((Asc(WFHDR.WFRULES) And 252) + 1)
     End If
     xxRGL$ = Space$(RFLTH(1))

S1400: ' Begin with Base Record
     SUMBASE.LOBY = 10000: SUMBASE.LOMY = 10000: SUMBASE.LODY = 10000
     SUM = SUMBASE: TOT = SUMBASE
     Call PutMSG("XXXXM015")
     L = 20: LAGL = 0: LRGL = 0: PAGE.COFST = 1
     A = 0: LGL = -1: DUPID = 0: DupCNT1 = 0: DupCNT2 = 0
     If RELSUMOPT Then Call PutMSG("XXXXM064")
     Call RWRKGet("BASE", RC, PTR&): If RC <> 1 Then GoTo S1040
     FIRSTONE = -1: GoTo S1450

S1415: ' Process next entry on chart
     If RELSUMOPT Then
        Call RWRKGet("WRKREC", RC, PTR&)
        Else
        Call RWRKGet("NEXT", RC, PTR&)
     End If
     If (RC < 1) Or (RC > 8) Then GoTo S1455
     DUPID = ((WFRELREC.id < 0) Or (WFRELREC.NREF <> 0))

S1450: ' Process Workfile record
     RELID = Abs(WFRELREC.id)
     If (WFRELREC.id > 0) And (WFRELREC.NREF <> 0) Then DupCNT1 = DupCNT1 + 1
     If ((WFRELREC.id < 0) Or (WFRELREC.NREF <> 0)) And (DupCNT2 < 32756) Then DupCNT2 = DupCNT2 + 1
     AGL = -WFAGLREC.AGL: RGL = AGL + WFGLREC.DGL
     If RELSUMOPT Then GoSub S2200: Else GoSub S2400
     If A <> 27 Then GoTo S1415

S1455: ' Process Final Totals
     If RELSUMOPT Then
        GoSub S2210
        Else
        GoSub S2460
        If VRTCL Then GoSub S2700
        Call RptCLOSE
        If PRT.AGAIN Then
           If VRTCL Then GoTo S1455 Else GoTo S1400
        End If
     End If
     Call FamCLOSE
     Exit Sub
     
S1500: ' Error Routine
    SysVAR.NVAR1 = ERL
    Call ERRMessage("XXXXM098", ERR, ERL)
    PRT.AGAIN = 0
    Resume S1455

S2200: ' Process Record for RELATIVE Table
     If WFRELREC.id > 0 Then
        RELCNT(AGL, RGL) = RELCNT(AGL, RGL) + 1
     End If
     Return

S2210: ' Print Relationship Summary Table
      Call PrintOPEN
      If Not OKAY Then GoTo S1040
      FULLWIDTH = 0: PLINE$ = Space$(SPL)
      If (PRT.OPT <> 2) And (PAGE.BWIDTH > PRT.FWIDTH - PAGE.IMARGIN - PAGE.OMARGIN) Then
         Call PutMSG("XXXXM070")
         If Asc(CA$) = 27 Then GoTo S1040 Else FULLWIDTH = (CA$ <> "1")
         If FULLWIDTH Then
            PRT.CFORMS = -1: NOSPAN = -1
            PRT.FWIDTH = PAGE.BWIDTH + PAGE.IMARGIN + PAGE.OMARGIN
            PLINE$ = Space$(PRT.FWIDTH)
         End If
      End If
      RELABEL$ = PLINE$: RELCONT$ = PLINE$
      Call BldSumTITLE

S2215: ' Print Table (AGAIN?)
      ReDim AGLTOT(LORGL To 0)
      RGLTOT = 0: RGLTOTAL = 0
      xRGL = LORGL: SREL$ = Space$(RFLTH(2))
      While xRGL <= HIRGL: xAGL = 0: FX = 2
         LSet PLINE$ = "": Call WriteLINE: RGLTOT = 0
         RSet xxRGL$ = Str$(xRGL): Mid$(RELCONT$, DFPP(1), RFLTH(1)) = xxRGL$
         While xAGL >= LORGL
            PP = DFPP(FX): FLTH = RFLTH(DFNUM(FX))
            RCNT = RELCNT(xAGL, xRGL)
            If (xAGL > xRGL) Or (RCNT = 0) Then
               Mid$(RELCONT$, PP + FLTH - 1, 1) = "."
               Else
               If ShowREL Then
                  RSet SREL$ = Relation$("S", "*", xAGL, xRGL, RRELCD$())
                  Mid$(RELABEL$, PP, FLTH) = SREL$
               End If
               If RCNT > 0 Then Mid$(RELCONT$, PP, FLTH) = SFILL$(RCNT, FLTH)
               AGLTOT(xAGL) = AGLTOT(xAGL) + RCNT
               RGLTOT = RGLTOT + RCNT
            End If
            FX = FX + 1: xAGL = xAGL - 1
         Wend
         Mid$(RELCONT$, DFPP(FX), RFLTH(FX)) = SFILL$(RGLTOT, RFLTH(FX))
         If ShowREL Then LSet PLINE$ = RELABEL$: Call PrintLINE: LSet RELABEL$ = ""
         LSet PLINE$ = RELCONT$: Call PrintLINE: LSet RELCONT$ = ""
         xRGL = xRGL + 1: RGLTOTAL = RGLTOTAL + RGLTOT
      Wend
      Call WriteLINE
      Mid$(PLINE$, DFPP(1), 255) = TOTAL$
      xAGL = 0: FX = 2
      While xAGL >= LORGL
         PP = DFPP(FX): FLTH = RFLTH(DFNUM(FX))
         Mid$(PLINE$, PP, FLTH) = SFILL$(AGLTOT(xAGL), FLTH)
         FX = FX + 1: xAGL = xAGL - 1
      Wend
      Mid$(PLINE$, DFPP(FX), RFLTH(FX)) = SFILL$(RGLTOTAL, RFLTH(FX))
      Call PrintLINE
      Mid$(PLINE$, DFPP(1), 255) = ACCUM$
      xAGL = 0: FX = 2: AGLACCUM = 0
      While xAGL >= LORGL
         AGLACCUM = AGLACCUM + AGLTOT(xAGL)
         PP = DFPP(FX): FLTH = RFLTH(DFNUM(FX))
         Mid$(PLINE$, PP, FLTH) = SFILL$(AGLACCUM, FLTH)
         FX = FX + 1: xAGL = xAGL - 1
      Wend
      Call PrintLINE
      Call RptCLOSE
      If PRT.AGAIN Then GoTo S2215
      Return

S2400: ' Process Records for ANC/DESC Genlevel Summaries
     Call FF1GetRec(RELID)
     If (SUM.TOTAL = 0) Or (RGL = LRGL) Then GoTo S2420
     If Not VRTCL Then GoSub S2550
     GoSub S2570
     FIRSTONE = -1: LAGL = AGL: LRGL = RGL
     SUM = SUMBASE

S2420: ' Increment Summary Counters
     Call GetBDEvents("", BPL$, "", DPL$, "")
     C = -((WFRELREC.id > 0) Or (DUPID And COUNTDUPS))
     If (COUNTPLACES And C) Then
        If RTrim$(BPLACE) <> "" Then SUM.BPL = SUM.BPL + 1
        If RTrim$(DPLACE) <> "" Then SUM.DPL = SUM.DPL + 1
     End If
     If ((FF1Rec.BY > 0) And C) Then
        SUM.BDT = SUM.BDT + 1
        If (FF1Rec.BY < SUM.LOBY) Then SUM.LOBY = FF1Rec.BY
        If (FF1Rec.BY > SUM.HIBY) Then SUM.HIBY = FF1Rec.BY
     End If
     If ((FF1Rec.DY > 0) And C) Then
        SUM.DDT = SUM.DDT + 1
        If FF1Rec.DY < 9999 Then
           SUM.DDTV = SUM.DDTV + 1
           If (FF1Rec.DY < SUM.LODY) Then SUM.LODY = FF1Rec.DY
           If (FF1Rec.DY > SUM.HIDY) Then SUM.HIDY = FF1Rec.DY
        End If
     End If
     If (COUNTMARRY And C) Then
        SPREC = FF1Rec.SPOUSE
        If SPREC <> 0 Then SUM.MARRIED = SUM.MARRIED + 1
        While (SPREC <> 0)
           SUM.MARCNT = SUM.MARCNT + 1
           Call FF3GetRec(SPREC)
           If (FF3SPOUSE.D1Y > 0) Then
              SUM.MDT = SUM.MDT + 1
              If (FF3SPOUSE.D1Y < SUM.LOMY) Then SUM.LOMY = FF3SPOUSE.D1Y
              If (FF3SPOUSE.D1Y > SUM.HIMY) Then SUM.HIMY = FF3SPOUSE.D1Y
           End If
           If (FF3SPOUSE.MLOC <> 0) Then SUM.MPL = SUM.MPL + 1
           If (FF3SPOUSE.SP1ID = RELID) Then
              SPREC = FF3SPOUSE.SP1NXT
              Else
              SPREC = FF3SPOUSE.SP2NXT
           End If
        Wend
     End If
     If ((COUNTNOP Or LISTNODES) And C) Then
        If (FF1Rec.FID = 0) And (FF1Rec.MID = 0) Then
           SUM.NOP = SUM.NOP + 1
           If LISTNODES Then GoSub S2500
        End If
     End If
     If COUNTDUPS Then SUM.DUPID = SUM.DUPID - DUPID
     If COUNTCLINE Then
        If RC = 1 Then SUM.CLINE = SUM.CLINE + WFAGLREC.CLINES
        If RC = 3 Then SUM.CLINE = SUM.CLINE + WFGLREC.CLINES
     End If
     SUM.TOTAL = SUM.TOTAL + C
     If FF1Rec.SEX = FGENDR$(1) Then
        SUM.MALE = SUM.MALE + C
        Else
        If FF1Rec.SEX = FGENDR$(2) Then SUM.FEMALE = SUM.FEMALE + C
     End If
     Return

S2460: ' Process last summary Line & Totals
     If Not VRTCL Then GoSub S2550
     If FIRSTONE Then GoSub S2570
     FIRSTONE = 0
     If Not VRTCL Then
        Call WriteLINE
        Mid$(PLINE$, DFPP(1), 5) = TOTAL$
        SUM = TOT: FINALTOT = -1
        GoSub S2550
     End If
     Return

S2500: ' Print Nodal Person
     If FIRSTONE Then
        RSet xxRGL$ = Str$(LRGL): Mid$(PLINE$, DFPP(1), RFLTH(1)) = xxRGL$
        If RptOPTION.Relation Then Mid$(PLINE$, DFPP(2), RFLTH(2)) = Relation$("M", "*", LAGL, LRGL, RRELCD$())
        Call PrintLINE: FIRSTONE = 0
     End If
     FX = 2 - RptOPTION.Relation: PP = DFPP(FX)
     RSet zALIN$ = Str$(WFANCREC.LINEAGE)
     Mid$(PLINE$, PP, LALIN) = zALIN$: PP = PP + LALIN + 2
     If FF1Rec.BY > 0 Then Mid$(PLINE$, PP, 4) = ZFILL$(FF1Rec.BY, 4)
     Mid$(PLINE$, PP + 4, 1) = "-"
     If FF1Rec.DY > 0 Then Mid$(PLINE$, PP + 5, 4) = ZFILL$(FF1Rec.DY, 4)
     Mid$(PLINE$, PP + 11, 52) = FMTNAME$(0)
     Call PrintLINE
     Return

S2550: ' Print Ancestor/Descendant Summary Line - Horizontal Style
     FX = 1
     While FX <= PAGE.DFCNT
        X = DFNUM(FX)
        PP = DFPP(FX): FLTH = RFLTH(X)
        Select Case X
           Case 1:  If FIRSTONE Then RSet xxRGL$ = Str$(LRGL): Mid$(PLINE$, PP, FLTH) = xxRGL$
           Case 2:  If FIRSTONE Then Mid$(PLINE$, PP, FLTH) = Relation$("M", "*", LAGL, LRGL, RRELCD$())
           Case 3:  Mid$(PLINE$, PP, FLTH) = SFILL$(SUM.TOTAL, FLTH)
                    If SUM.TOTAL > (SUM.MALE + SUM.FEMALE) Then Mid$(PLINE$, PP + FLTH, 1) = "*"
           Case 4:  If Not FINALTOT Then Mid$(PLINE$, PP, FLTH) = SFILL$(SUM.CLINE, FLTH)
           Case 5:  Mid$(PLINE$, PP, FLTH) = SFILL$(SUM.DUPID, FLTH)
           Case 6:  Mid$(PLINE$, PP, FLTH) = SFILL$(SUM.NOP, FLTH)
           Case 7:  Mid$(PLINE$, PP, FLTH) = SFILL$(SUM.MALE, FLTH)
           Case 8:  Mid$(PLINE$, PP, FLTH) = SFILL$(SUM.FEMALE, FLTH)
           Case 9:  Mid$(PLINE$, PP, FLTH) = SFILL$(SUM.BDT, FLTH)
           Case 10: Mid$(PLINE$, PP, FLTH) = SFILL$(SUM.BPL, FLTH)
           Case 11: If SUM.BDT > 0 Then Mid$(PLINE$, PP, FLTH) = ZFILL$(SUM.LOBY, FLTH)
           Case 12: If SUM.BDT > 0 Then Mid$(PLINE$, PP, FLTH) = ZFILL$(SUM.HIBY, FLTH)
           Case 13: Mid$(PLINE$, PP, FLTH) = SFILL$(SUM.MARRIED, FLTH)
           Case 14: Mid$(PLINE$, PP, FLTH) = SFILL$(SUM.MARCNT, FLTH)
           Case 15: Mid$(PLINE$, PP, FLTH) = SFILL$(SUM.MDT, FLTH)
           Case 16: Mid$(PLINE$, PP, FLTH) = SFILL$(SUM.MPL, FLTH)
           Case 17: If SUM.MDT > 0 Then Mid$(PLINE$, PP, FLTH) = ZFILL$(SUM.LOMY, FLTH)
           Case 18: If SUM.MDT > 0 Then Mid$(PLINE$, PP, FLTH) = ZFILL$(SUM.HIMY, FLTH)
           Case 19: Mid$(PLINE$, PP, FLTH) = SFILL$(SUM.DDT, FLTH)
           Case 20: Mid$(PLINE$, PP, FLTH) = SFILL$(SUM.DDTV, FLTH)
           Case 21: Mid$(PLINE$, PP, FLTH) = SFILL$(SUM.DPL, FLTH)
           Case 22: If SUM.DDTV > 0 Then Mid$(PLINE$, PP, FLTH) = ZFILL$(SUM.LODY, FLTH)
           Case 23: If SUM.DDTV > 0 Then Mid$(PLINE$, PP, FLTH) = ZFILL$(SUM.HIDY, FLTH)
        End Select
        FX = FX + 1
     Wend
     Call PrintLINE
     Return

S2570: ' Update Overall Counters
     TOT.TOTAL = TOT.TOTAL + SUM.TOTAL
     TOT.NOP = TOT.NOP + SUM.NOP
     TOT.DUPID = TOT.DUPID + SUM.DUPID
     TOT.CLINE = TOT.CLINE + SUM.CLINE
     TOT.MALE = TOT.MALE + SUM.MALE
     TOT.FEMALE = TOT.FEMALE + SUM.FEMALE
     TOT.MARRIED = TOT.MARRIED + SUM.MARRIED
     TOT.MARCNT = TOT.MARCNT + SUM.MARCNT
     TOT.BPL = TOT.BPL + SUM.BPL: TOT.MPL = TOT.MPL + SUM.MPL
     TOT.DPL = TOT.DPL + SUM.DPL
     TOT.BDT = TOT.BDT + SUM.BDT: TOT.DDT = TOT.DDT + SUM.DDT
     TOT.DDTV = TOT.DDTV + SUM.DDTV: TOT.MDT = TOT.MDT + SUM.MDT
     If (SUM.LOBY > 0) And (SUM.LOBY < TOT.LOBY) Then TOT.LOBY = SUM.LOBY
     If SUM.HIBY > TOT.HIBY Then TOT.HIBY = SUM.HIBY
     If (SUM.LODY > 0) And (SUM.LODY < TOT.LODY) Then TOT.LODY = SUM.LODY
     If SUM.HIDY > TOT.HIDY Then TOT.HIDY = SUM.HIDY
     If (SUM.LOMY > 0) And (SUM.LOMY < TOT.LOMY) Then TOT.LOMY = SUM.LOMY
     If SUM.HIMY > TOT.HIMY Then TOT.HIMY = SUM.HIMY
     If VRTCL Then SUMTAB(Abs(LRGL)) = SUM
     Return

S2700: ' Print Ancestor/Descendant Summary Line - Vertical Style
     FX = 3
     While FX <= PAGE.DFCNT
        X = DFNUM(FX): Mid$(PLINE$, 1, 255) = RFLBL$(X)
        FLTH = RFLTH(X)
        PP = RFLTH(0) + DFPP(0)
        For Y = 0 To HIRGL + 1
            PP = PP + 6 - FLTH
            If (Y <= HIRGL) Then SUM = SUMTAB(Y) Else SUM = TOT
            Select Case X
               Case 3:  Mid$(PLINE$, PP, FLTH) = SFILL$(SUM.TOTAL, FLTH)
                        If SUM.TOTAL > (SUM.MALE + SUM.FEMALE) Then Mid$(PLINE$, PP + FLTH, 1) = "*"
               Case 4:  Mid$(PLINE$, PP, FLTH) = SFILL$(SUM.CLINE, FLTH)
               Case 5:  Mid$(PLINE$, PP, FLTH) = SFILL$(SUM.DUPID, FLTH)
               Case 6:  Mid$(PLINE$, PP, FLTH) = SFILL$(SUM.NOP, FLTH)
               Case 7:  Mid$(PLINE$, PP, FLTH) = SFILL$(SUM.MALE, FLTH)
               Case 8:  Mid$(PLINE$, PP, FLTH) = SFILL$(SUM.FEMALE, FLTH)
               Case 9:  Mid$(PLINE$, PP, FLTH) = SFILL$(SUM.BDT, FLTH)
               Case 10: Mid$(PLINE$, PP, FLTH) = SFILL$(SUM.BPL, FLTH)
               Case 11: If SUM.BDT > 0 Then Mid$(PLINE$, PP, FLTH) = ZFILL$(SUM.LOBY, FLTH)
               Case 12: If SUM.BDT > 0 Then Mid$(PLINE$, PP, FLTH) = ZFILL$(SUM.HIBY, FLTH)
               Case 13: Mid$(PLINE$, PP, FLTH) = SFILL$(SUM.MARRIED, FLTH)
               Case 14: Mid$(PLINE$, PP, FLTH) = SFILL$(SUM.MARCNT, FLTH)
               Case 15: Mid$(PLINE$, PP, FLTH) = SFILL$(SUM.MDT, FLTH)
               Case 16: Mid$(PLINE$, PP, FLTH) = SFILL$(SUM.MPL, FLTH)
               Case 17: If SUM.MDT > 0 Then Mid$(PLINE$, PP, FLTH) = ZFILL$(SUM.LOMY, FLTH)
               Case 18: If SUM.MDT > 0 Then Mid$(PLINE$, PP, FLTH) = ZFILL$(SUM.HIMY, FLTH)
               Case 19: Mid$(PLINE$, PP, FLTH) = SFILL$(SUM.DDT, FLTH)
               Case 20: Mid$(PLINE$, PP, FLTH) = SFILL$(SUM.DDTV, FLTH)
               Case 21: Mid$(PLINE$, PP, FLTH) = SFILL$(SUM.DPL, FLTH)
               Case 22: If SUM.DDTV > 0 Then Mid$(PLINE$, PP, FLTH) = ZFILL$(SUM.LODY, FLTH)
               Case 23: If SUM.DDTV > 0 Then Mid$(PLINE$, PP, FLTH) = ZFILL$(SUM.HIDY, FLTH)
            End Select
            PP = PP + FLTH + DFPP(0)
        Next Y
        Call PrintLINE
        FX = FX + 1
     Wend
     Return

End Sub

