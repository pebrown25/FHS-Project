Attribute VB_Name = "modFHSWMLMG"
DefInt A-Z
Dim HIID, HIMRN, S, FMT001, MLMGFile$
Dim INFOSW, RECCNT()
Dim TLINE$

Public Sub UtilMLMG(OPT$)
1  Rem $INCLUDE: 'FHSCOMON.BAS'
   PN$ = " FAMILY HISTORY SYSTEM - FHSWMLMG - MAILMERGE Export Program"
   CY$ = " (C) Copyright 1992-1998 by Phillip E. Brown"
   'DEF FNCHR(CA$) = (((CA$ > "@") And (CA$ < "[")) Or ((CA$ > "`") And (CA$ < "{")))
   'DEF FNNUM(CA$) = ((CA$ > "/") And (CA$ < ":"))
   ReDim RECCNT(9), TAG$(140)
   MLMGFile$ = FT$(12): READ$ = "READQ"
   
   Select Case OPT$
      Case "INIT": GoTo 40
      Case "PROCESS": GoTo 1000
      Case "PRINT": GoTo 1000
   End Select
   
40 ' Format Display
   Call PgmPREP("FHSMLMG")
   Load frmFHSWUPgm
   Call RptLOAD("MMEX"): Call RptLOAD("MMLS")
   frmFHSWUPgm.UPSetup.AddItem "MMEX"
   frmFHSWUPgm.UPSetup.AddItem "MMLS"
   'frmFHSWUPgm.UPSetup.TEXT = RptCODE
   frmFHSWUPgm.UPSetup_Click
   frmFHSWUPgm.Show
   Exit Sub

1000: ' Process EXPORT Request
   Select Case RptCODE
      Case "MMEX": Call MLMGExport
      Case "MMLS": Call MLMGList
   End Select
End Sub
Private Sub MLMGExport()

If SCNT = 0 Then
   Call PutMSG("No Records have been selected...")
   Exit Sub
End If
Call frmFHSWUPgm.ShowStats("MAILMERG", -1, "")
600 ' Export Information to GEDCOM formatted dataset
    On Local Error GoTo 601
    ReDim XMONTB$(14): For X = 1 To 14: XMONTB$(X) = RMONTB$(X): Next X
    Call RptLOAD("MMEX"): If Not OKAY Then GoTo 750
    'Call WinSWITCH(2): Call WinCLR
    'Call RptOptDisplay("MMEX")
    Call GetRVAR("DIVORCED", X)
    If FOUND Then xDIVORCED$ = RTrim$(RVAR$(X)) Else xDIVORCED$ = "Divorced"
    INFOSW = 0
    For X = 0 To 11
        xX$ = "MM" + ZFILL$(X, 2)
        Call RptOptGET(xX$, v)
        If FOUND Then If v <> 2 Then INFOSW = INFOSW Or (2 ^ X)
    Next X
    Labels = ((INFOSW And 1) > 0)
    EXPORTID = ((INFOSW And 2) > 0)
    MARRIAGE = ((INFOSW And 4) > 0)
    DateS = ((INFOSW And 8) > 0)
    PLACES = ((INFOSW And 16) > 0)
    ADDRESSES = ((INFOSW And 32) > 0)
    COMBINED = ((INFOSW And 128) > 0)
    GENDER = ((INFOSW And 256) > 0)
    ShowREL = ((INFOSW And 512) > 0): RELTYPE$ = "M"
    BQUOTES = ((INFOSW And 1024) > 0)
    TITLE = ((INFOSW And 2048) > 0)
    MMTEXT$ = Space$(75)

    If Labels Then
       Call TableLOAD("MAILMERG", MMTAB, 0)
       ReDim MMLABEL$(56)
       Y = TABLE(MMTAB).FT
       While Y > 0
          Z = Val(MID$(TABDATA$(Y), 3, 2))
          If (Z < 57) Then
             MMLABEL$(Z) = RTrim$(MID$(TABDATA$(Y), 9, 255))
          End If
          Y = TABDATACHN(Y).FWD
       Wend
       Call TableFREE("MAILMERG")
    End If

    If TITLE Then
       ReDim xTITLE$(6)
       For X = 1 To 6
           xX$ = "TITL" + ZFILL$(X, 2)
           Call GetRVAR(xX$, Y)
           If FOUND Then xTITLE$(X) = RVAR$(Y)
       Next X
    End If
    GoTo 605

601 'Error Setting up for GEDCOM Export
    XERR = ERR: XERL = ERL
    ReDim NDXID(1), SelSET(1), RECCNT(1)
    Call ERRMessage("MLMGM008", XERR, XERL)
    Resume 750

605 ' Determine Records to Process
    'Call PutMSG("MLMGM002")
    'If A = 27 Then GoTo 750
    'ALLRECS = (CA$ = "1"): SELECTREL = (CA$ = "3")
    'ReDim SelSET(FF1Hdr.MAXID)
    'SlctCNT = 0: SLCTMAX = 0
    'If ALLRECS Then
    '   SLCTMAX = FF1Hdr.MAXID
    '   For x = 1 To SLCTMAX: SelSET(x) = -1: Next x
    '   Else
    '   If SELECTREL Then
    '      GoSub 670: If Not OKAY Then GoTo 605
    '      Else
    '      Call SWRKBldList(SelSET(), SLCTMAX, 0)
    '      If Not OKAY Then GoTo 605
    '      For x = 1 To SLCTMAX
    '          If SelSET(x) <> 0 Then SelSET(x) = -1: SlctCNT = SlctCNT + 1
    '      Next x
    '   End If
    '   SysVAR.NVAR1 = SlctCNT: SysVAR.NVAR2 = FF1Hdr.MAXID
    '   Call PutMSG("MLMGM003"): If A = 27 Then GoTo 690
    If SCNT < FF1Hdr.MAXID Then
       If COMBINED Then
          Call PutMSG("MLMGM011"): If a = 27 Then GoTo 690
          ADDSPOUSE = (CA$ = SNGLKEY$(1))
          If ADDSPOUSE Then
             Call PutMSG("MLMGM012"): If a = 27 Then GoTo 690
             ADDLIVING = (CA$ = SNGLKEY$(1))
             GoSub 680: If CA$ = Chr$(27) Then GoTo 690
          End If
       End If
    End If

660 ' Determine Sequence of Exported Records
    Call PutMSG("MLMGM004")
    If a = 27 Then GoTo 690
    If CA$ < "1" Or CA$ > "2" Then Call ErrBEEP(0): GoTo 660
    INDEXED = (CA$ = "2")
    ReDim NDXID(FF1Hdr.MAXID)
    If INDEXED Then
       Call INDEXFile("READ")
       If Asc(CA$) = 27 Then GoTo 690 Else If Not OKAY Then GoTo 660
       NDXRECS = NDXHDR.NDXRECS: HUSBSURNM = ENV.HUSBSURNM
       Else
       For I = 1 To FF1Hdr.MAXID: NDXID(I) = I: Next I
       NDXRECS = FF1Hdr.MAXID
    End If
    Call PutMSG("")
    HIID = FF1Hdr.MAXID
    If ShowREL And (Not SELECTREL) Then
       GoSub 670: If a = 27 Then GoTo 690
    End If
    GoTo 700

670 ' Build Relationship Table
    If ShowREL Then
       Call RptOptGET("EXAD", X): ExtAdoptRule = (X = 1)
    End If
    OKAY = 0
    While (Not OKAY) And (a <> 27)
       Close: Call RWRKOpen("READ", 0, 0)
    Wend
    If Not OKAY Then
       If Not SELECTREL Then
          Call PutMSG("MLMGM010")
          If (CA$ <> SNGLKEY$(1)) Then a = 27
          ShowREL = 0
       End If
    End If
    If OKAY Then
       If ShowREL Then ReDim RELATE(FF1Hdr.MAXID)
       Call PutMSG("XXXXM064")
       Call RWRKGet("BASE", RC, PTR&)
       While (0 < RC) And (RC < 9)
          If WFRELREC.id > 0 Then
             If ShowREL Then
                X = RELATE(WFRELREC.id)
                If (X = 0) Or (WFGLREC.DGL = 0) Then
                   X = 256 * (-WFGLREC.AGL) + WFGLREC.DGL
                   If ExtAdoptRule Then X = X - 128 * ((WFRELREC.CSTAT And 5) > 0)
                   RELATE(WFRELREC.id) = X
                End If
             End If
             If SELECTREL Then
                If SelSET(WFRELREC.id) = 0 Then
                   SlctCNT = SlctCNT + 1: SelSET(WFRELREC.id) = -1
                   If WFRELREC.id > SLCTMax Then SLCTMax = WFRELREC.id
                End If
             End If
          End If
          Call RWRKGet("WRKREC", RC, PTR&)
       Wend
    End If
    Return

680 ' ADD Spouse to Selection table
    Call PutMSG("MLMGM013")
    U = SLCTMax: v = SlctCNT: CA$ = "": a = 0: C = 0
    Call FAMOpen("READ"): If Not OKAY Then Return
    For X = 1 To SLCTMax
       If SelSET(X) < 0 Then
          Call FF1GetRec(X)
          If OKAY Then
             Z = FF1Rec.SPOUSE
             If Z <> 0 Then
                USEOWN = ((Asc(FF1Rec.STATUS2) And 128) = 128)
                MALE = (FF1Rec.SEX = FGENDR$(1))
                FEMALE = (FF1Rec.SEX = FGENDR$(2))
                If (MALE) Or (FEMALE And (Not USEOWN)) Then
                   Call FF3GetRec(Z)
                   w = FF3SPOUSE.SP1ID: If w = X Then w = FF3SPOUSE.SP2ID
                   DIVORCED = (RTrim$(FF3SPOUSE.STATUS2) = xDIVORCED$)
                   ' Check to see if this is Most Recent Marriage of Spouse
                   If (w <> 0) And (Not DIVORCED) Then
                      If SelSET(w) = 0 Then
                         Call FF1GetRec(w)
                         If (Not ADDLIVING) Or (FF1Rec.DY = 0) Then
                            If Z = FF1Rec.SPOUSE Then
                               SPUSEOWN = ((Asc(FF1Rec.STATUS2) And 128) = 128)
                               If (FEMALE) Or (MALE And (Not SPUSEOWN)) Then
                                  SelSET(w) = 1: v = v + 1
                                  If w > U Then U = w
                               End If
                            End If
                         End If
                      End If
                   End If
                End If
             End If
          End If
       End If
       If CA$ <> Chr$(27) Then CA$ = INKEY$
       If CA$ = Chr$(27) Then X = SLCTMax
    Next X
    If (CA$ <> Chr$(27)) And (v > SlctCNT) Then
       SysVAR.NVAR1 = v - SlctCNT
       Call PutMSG("MLMGM014")
       SlctCNT = v: SLCTMax = U
    End If
    Return

690 ' Early Return from EXPORT operation
    Call FamCLOSE
    GoTo 735

700 ' Export Data for Selected Name Records
    'Call WinCLR
    'Call FmtFIND("MLMGS010"): Call WinFORMAT(0)
    If Not INDEXED Then HUSBSURNM = ((INFOSW And 64) > 0)
    Call FAMOpen("READ"): If Not OKAY Then GoTo 690
    Call MLMGOpen("WRITE"): If Not OKAY Then GoTo 750
    Call PutMSG("MLMGM005")
    On Local Error GoTo 705
    GoTo 710

705 ' Error During Export
    Call ERRMessage("MLMGM009", ERR, ERL)
    OKAY = 0: SID = HIID: I = HIMRN: a = 27
    Select Case ERRReturn
       Case 1: Resume 721
       Case 2: Resume 750
    End Select
    Resume 730

710 ' Create LABELs Record
    ReDim RECCNT(10)
    If Labels Then
       RT = 0: GoSub 900: Print #5, Chr$(34); MMLABEL$(0); Chr$(34);
       If EXPORTID Then
          LSet MMTEXT$ = MMLABEL$(1): GoSub 890
          LSet MMTEXT$ = MMLABEL$(2): GoSub 890
          If COMBINED Or MARRIAGE Then
             LSet MMTEXT$ = MMLABEL$(3): GoSub 890
          End If
       End If
       If TITLE Then LSet MMTEXT$ = MMLABEL$(4): GoSub 890
       LSet MMTEXT$ = MMLABEL$(5): GoSub 890
       LSet MMTEXT$ = MMLABEL$(11): GoSub 890
       If ShowREL Then LSet MMTEXT$ = MMLABEL$(12): GoSub 890
       If GENDER Then LSet MMTEXT$ = MMLABEL$(13): GoSub 890
       If DateS Then LSet MMTEXT$ = MMLABEL$(14): GoSub 890
       If PLACES Then LSet MMTEXT$ = MMLABEL$(15): GoSub 890
       If DateS Then LSet MMTEXT$ = MMLABEL$(16): GoSub 890
       If PLACES Then LSet MMTEXT$ = MMLABEL$(17): GoSub 890
       If MARRIAGE Then
          If DateS Then LSet MMTEXT$ = MMLABEL$(20): GoSub 890
          If PLACES Then LSet MMTEXT$ = MMLABEL$(21): GoSub 890
       End If
       If COMBINED Then
          LSet MMTEXT$ = MMLABEL$(31): GoSub 890
          If ShowREL Then LSet MMTEXT$ = MMLABEL$(32): GoSub 890
          If GENDER Then LSet MMTEXT$ = MMLABEL$(33): GoSub 890
          If DateS Then LSet MMTEXT$ = MMLABEL$(34): GoSub 890
          If PLACES Then LSet MMTEXT$ = MMLABEL$(35): GoSub 890
          If DateS Then LSet MMTEXT$ = MMLABEL$(36): GoSub 890
          If PLACES Then LSet MMTEXT$ = MMLABEL$(37): GoSub 890
       End If
       If ADDRESSES Then
          For Y = 51 To 54: LSet MMTEXT$ = MMLABEL$(Y): GoSub 890: Next Y
          If AFMT = 1 Then LSet MMTEXT$ = MMLABEL$(55): GoSub 890
          LSet MMTEXT$ = MMLABEL$(56): GoSub 890
       End If
       GoSub 899
    End If

720 ' Process Selected Records
    SelSET(0) = 0
    For SID = 1 To HIID
        ERRReturn = 1
        id = NDXID(SID)
        If SelSET(id) <> 0 Then
           Call FF1GetRec(id): GoSub 800
           If (CA$ <> Chr$(27)) Then CA$ = INKEY$
           If CA$ = Chr$(27) Then SID = HIID
        End If
721 Next SID

730 ' End of EXPORT
    ERRReturn = 2
    Call PutMSG("MLMGM006")
735 Call PutMSG("")
    'Call WinPREP(2): Call WinCLR
750 Call FamCLOSE
    Call frmFHSWUPgm.ShowStats("MAILMERG", 99, "")
    For X = 1 To 14: RMONTB$(X) = XMONTB$(X): Next X
    Exit Sub

800 ' Produce MAILMERGE File Record Line
    SP1 = 0: MLOCPTR = 0: MADRSPTR = 0: MDY = 0: MDMD = 0: DIVORCED = 0
    USEOWN = ((Asc(FF1Rec.STATUS2) And 128) = 128)
    SEX$ = FF1Rec.SEX
    MALE = (SEX$ = FGENDR$(1)): FEMALE = (SEX$ = FGENDR$(2))
    SPPTR = FF1Rec.SPOUSE: SP1 = 0: SPMATCH = 0: SPUSEOWN = 0
    If SPPTR <> 0 Then
       Call FF3GetRec(SPPTR)
       DIVORCED = (RTrim$(FF3SPOUSE.STATUS2) = xDIVORCED$)
       SP1 = FF3SPOUSE.SP1ID: If SP1 = id Then SP1 = FF3SPOUSE.SP2ID
       ' Check to see if this is Most Recent Marriage of Spouse
       If COMBINED And (SP1 <> 0) And (Not DIVORCED) Then
          Call FF1GetRec(SP1)
          If (SelSET(SP1) <> 0) And (FF1Rec.SEX <> SEX$) Then
             SPMATCH = (FF1Rec.SPOUSE = SPPTR)
             If MALE Then
                SPUSEOWN = ((Asc(FF1Rec.STATUS2) And 128) = 128)
                Else
                SPUSEOWN = USEOWN
             End If
          End If
          Call FF1GetRec(id)
       End If
    End If
    MULTREC = (COMBINED And SPMATCH And (Not SPUSEOWN))
    If MULTREC And FEMALE Then Return
    ADRSPTR = FF1Rec.ADRS
    If SPPTR <> 0 Then
       MADRSPTR = FF3SPOUSE.LOC
       MDY = FF3SPOUSE.D1Y: MDMD = FF3SPOUSE.D1MD: MLOCPTR = FF3SPOUSE.MLOC
    End If
    If MULTREC And FEMALE Then Return
    RT = 0: GoSub 900: Print #5, SNX$(id);
    If EXPORTID Then
       LSet MMTEXT$ = SNX$(FF1Rec.FID): GoSub 895
       LSet MMTEXT$ = SNX$(FF1Rec.MID): GoSub 895
       If COMBINED Or MARRIAGE Then
          If (Not MARRIAGE) And (Not MULTREC) Then
             LSet MMTEXT$ = "0": GoSub 895
             Else
             LSet MMTEXT$ = SNX$(SP1): GoSub 895
          End If
       End If
    End If
    If TITLE Then
       If MULTREC Then
          X = 6
          Else
          If MALE Then
             If SPPTR = 0 Then X = 1 Else X = 3
             Else
             If FEMALE Then
                If SPPTR = 0 Then
                   X = 2
                   Else
                   If (HUSBSURNM And (Not USEOWN)) Then
                      X = 4
                      Else
                      X = 5
                   End If
                End If
                Else
                X = 0
             End If
          End If
       End If
       LSet MMTEXT$ = xTITLE$(X): GoSub 890
    End If
    LSet MMTEXT$ = FF1Rec.SURNM
    If FEMALE And HUSBSURNM And (SP1 > 0) And (Not USEOWN) Then
       Call FF1GetRec(SP1): LSet MMTEXT$ = FF1Rec.SURNM
       Call FF1GetRec(id)
    End If
    GoSub 890
    GoSub 820
    If MARRIAGE Then
       If SPPTR <> 0 Then RT = 2: GoSub 900
       If DateS Then
          If SPPTR = 0 Then
             GoSub 890
             Else
             D1Y = MDY: D1MD = MDMD
             If D1Y > 0 Or D1MD > 0 Then GoSub 1000
             RT = 3: GoSub 895
          End If
       End If
       If PLACES Then
          LSet MPLACE = ""
          If MLOCPTR <> 0 Then
             Call FF3GetRec(MLOCPTR): LSet MMTEXT$ = MPLACE
          End If
          RT = 4: GoSub 890
       End If
    End If
    If COMBINED Then
       If MULTREC Then
          Call FF1GetRec(SP1): GoSub 820
          Call FF1GetRec(id)
          Else
          GoSub 825
       End If
    End If
    GoTo 850

820 ' Process Basic Info
    RT = 1: GoSub 900
    LSet MMTEXT$ = FF1Rec.GIVEN: GoSub 890
    If ShowREL Then GoSub 1200: GoSub 890
    If GENDER Then
       X = 1: FOUND = 0
       While (X < 3) And (Not FOUND)
          FOUND = (FGENDR$(X) = FF1Rec.SEX)
          If Not FOUND Then X = X + 1
       Wend
       LSet MMTEXT$ = MID$(RGENDR$(X), 2, 1): GoSub 890
    End If
    If PLACES Then
       BLOCPTR = FF1Rec.BLOC: LSet BPLACE = "": LSet DPLACE = ""
       If (BLOCPTR <> 0) Then
          Call FF3GetRec(BLOCPTR):
          If (FF3PFX.RTYPE = Chr$(12)) And (FF3PLACE2.PNEXT <> 0) Then
             Call FF3GetRec(FF3PLACE2.PNEXT)
          End If
       End If
    End If
    If DateS Then
       D1Y = FF1Rec.BY: D1MD = FF1Rec.BMD
       If D1Y > 0 Or D1MD > 0 Then GoSub 1000
       RT = 3: GoSub 895
    End If
    If PLACES Then
       RT = 4: LSet MMTEXT$ = BPLACE: GoSub 890
    End If
    If DateS Then
       D1Y = FF1Rec.DY: D1MD = FF1Rec.DMD
       If D1Y > 0 Or D1MD > 0 Then GoSub 1000
       RT = 3: GoSub 895
    End If
    If PLACES Then
       RT = 4: LSet MMTEXT$ = DPLACE: GoSub 890
    End If
    Return

825 ' Blank Info
    GoSub 890
    If ShowREL Then GoSub 890
    If GENDER Then GoSub 890
    If DateS Then GoSub 895
    If PLACES Then GoSub 890
    If DateS Then GoSub 895
    If PLACES Then GoSub 890
    Return

850 ' Process Residence
    If (Not ADDRESSES) Then GoTo 899
    D1Y = 0: D1MD = 0: D2Y = 0: D2MD = 0
    If ADRSPTR <> 0 Then
       Call FF2GetRec(ADRSPTR): D1Y = FF2PFX.D1Y: D1MD = FF2PFX.D1MD
    End If
    If MADRSPTR <> 0 Then
       Call FF2GetRec(MADRSPTR): D2Y = FF2PFX.D1Y
       If (D1Y < D2Y) Or ((D1Y = D2Y) And (D1MD <= FF2PFX.D1MD)) Then
          ADRSPTR = MADRSPTR
       End If
    End If
    If ADRSPTR = 0 Then
       For X = 1 To 6: GoSub 890: Next X
       If AFMT = 1 Then GoSub 890
       GoTo 899
    End If
    RT = 5: GoSub 900: Call FF2GetRec(ADRSPTR)
    Select Case AFMT
        Case 0: GoSub 1100: GoSub 890
                LSet MMTEXT$ = FF2SREC.LINE1: GoSub 890
                LSet MMTEXT$ = FF2SREC.LINE2: GoSub 890
                LSet MMTEXT$ = FF2SREC.CITY: GoSub 890
                LSet MMTEXT$ = FF2SREC.STATE: GoSub 890
                LSet MMTEXT$ = FF2SREC.ZIP: GoSub 890
        Case 1: LSet MMTEXT$ = FF2LREC.PHONE: GoSub 890
                LSet MMTEXT$ = FF2LREC.LINE1: GoSub 890
                LSet MMTEXT$ = FF2LREC.LINE2: GoSub 890
                LSet MMTEXT$ = FF2LREC.CITY: GoSub 890
                LSet MMTEXT$ = FF2LREC.STATE: GoSub 890
                LSet MMTEXT$ = FF2LREC.ZIP: GoSub 890
                LSet MMTEXT$ = FF2LREC.COUNTRY: GoSub 890
    End Select
    GoTo 899

890 ' Process Text Field Data in MMTEXT$
    If (RTrim$(MMTEXT$) = "") Then
       If BQUOTES Then
          Print #5, ","; Chr$(34); Chr$(34);
          Else
          Print #5, ",";
       End If
       Else
       Print #5, ","; Chr$(34); RTrim$(MMTEXT$); Chr$(34);
       If RT > 0 Then GoSub 900
    End If
    LSet MMTEXT$ = "": RT = 0
    Return

895 ' Process Text Field Data in MMTEXT$
    If RTrim$(MMTEXT$) = "" Then
       If BQUOTES Then
          Print #5, ","; Chr$(34); Chr$(34);
          Else
          Print #5, ",";
       End If
       Else
       Print #5, ","; RTrim$(MMTEXT$);
       If RT > 0 Then GoSub 900
    End If
    LSet MMTEXT$ = "": RT = 0
    Return

899 ' Finished with Record
    Print #5, ""
    Return

900 ' Process Record Counts
    RECCNT(RT) = RECCNT(RT) + 1
    'x = 15 + RT - 2 * (RT > 0)
    'COLOR FG(4), BG(4): LOCATE x, 43, 0
    'Print USING; "#####"; RECCNT(RT);
    S = RT: If S > 0 Then S = S + 1
    SV$ = SFILL$(RECCNT(RT), 5)
    Call frmFHSWUPgm.ShowStats("MAILMERG", S, SV$)
    If (RT = 0) And MULTREC Then
       RECCNT(10) = RECCNT(10) + 1
       SV$ = SFILL$(RECCNT(10), 5)
       'LOCATE x + 1, 43, 0: Print USING; "#####"; RECCNT(10);
       Call frmFHSWUPgm.ShowStats("MAILMERG", 1, SV$)
    End If
    RT = 0
    Return

1000 'd1y,d1md-->YYYYMMDD
     L = 8: LSet MMTEXT$ = "00000000"
     If (D1Y > 0) Or (D1MD > 0) Then
        DM = Int(D1MD / 100): DD = D1MD Mod 100
        If D1Y > 0 Then Mid$(MMTEXT$, 1, 4) = ZFILL$(D1Y, 4)
        If DM > 0 Then Mid$(MMTEXT$, 5, 2) = ZFILL$(DM, 2)
        If DD > 0 Then Mid$(MMTEXT$, 7, 2) = ZFILL$(DD, 2)
     End If
     Return

1100 ' Format Phone Number
     F1 = FF2SREC.FON1: F2 = FF2SREC.FON2: F3 = FF2SREC.FON3
     If (F1 > 0) Or (F2 > 0) Or (F3 > 0) Then
        LSet MMTEXT$ = "(000) 000-0000"
        Mid$(MMTEXT$, 2, 3) = ZFILL$(F1, 3)
        Mid$(MMTEXT$, 7, 3) = ZFILL$(F2, 3)
        Mid$(MMTEXT$, 11, 4) = ZFILL$(F3, 4)
     End If
     Return

1200 ' Get RELATIONSHIP
     X = RELATE(FF1Rec.RID)
     If X <> 0 Then
        Y = X And 128
        AGL = -(X \ 256): DGL = X And 127
        RGL = AGL + DGL
        LSet MMTEXT$ = Relation$(RELTYPE$, FF1Rec.SEX, AGL, RGL, RRELCD$())
        If ExtAdoptRule Then
           If (X And 128) > 0 Then LSet MMTEXT$ = "*" + MMTEXT$
        End If
        Else
        AGL = 0: DGL = 0: RGL = 0: LSet MMTEXT$ = ""
     End If
     Return

End Sub

Sub MLMGList()

2000 ' Print Exported Data File Listing
     Call RptOPEN("MMLS"): If (Not OKAY) Then GoTo 2069
     Call MLMGOpen("READ"): If (Not OKAY) Or (a = 27) Then GoTo 2069
     Call PrintOPEN: If Not OKAY Then GoTo 2069
2060 ' Print Report
     While Not EOF(5)
        Line Input #5, TLINE$
        TLTH = Len(RTrim$(TLINE$))
        TOFFSET = 1: POFFSET = 1
        PLTH = PAGE.BWIDTH
        While (TOFFSET <= TLTH)
           Mid$(PLINE$, POFFSET, PLTH) = MID$(TLINE$, TOFFSET, PLTH)
           Call PrintLINE: If a = 27 Then GoTo 2065
           TOFFSET = TOFFSET + PLTH
           POFFSET = 5: PLTH = PAGE.BWIDTH - POFFSET + 1
           While (TOFFSET <= TLTH) And (MID$(TLINE$, TOFFSET, 1) = " ")
              TOFFSET = TOFFSET + 1
           Wend
        Wend
     Wend
2065 Call RptCLOSE
     If PRT.AGAIN Then
        Call MLMGOpen("READQ"): If OKAY Then GoTo 2060
     End If
2069 Close
     Exit Sub
End Sub

Sub MLMGOpen(OPT$)

18000 ' Open MailMerge File
      ' Make Sure File is Closed
      MMDSN$ = TRANPFX$ + MLMGFile$
      On Local Error GoTo 18001
      Close #5
      GoTo 18002
18001 Resume 18002
18002 OKAY = 0: XOPT$ = OPT$: L = Len(XOPT$)
      QUIET = (MID$(XOPT$, L, 1) = "Q")
      If QUIET Then L = L - 1: XOPT$ = Left$(XOPT$, L)
      If Not QUIET Then
         Call frmFHSWFile.LOCATE("MAILMERG", "FULL")
         'Call FmtFIND("MLMGS004")
         'Call FmtFindFLD("DSNM"): y = FFLD.UTXT: LSet UTXT$(y) = MLMGFile$
         'Call MsgFORMAT
         If a = 27 Then OKAY = 0: GoTo 18300
         'MLMGFile$ = RTrim$(UTXT$(y))
         FDISK = (FDFSetup.TRANDRV = "A") Or (FDFSetup.TRANDRV = "B")
         MMDSN$ = TRANPFX$ + FT$(12)
         'If Not FDISK Then
         '   MMDSN$ = TRANPFX$ + FT$(12)
         '   Else
         '   MMDSN$ = MLMGFile$
         'End If
      End If
      SysVAR.FILENAME = MMDSN$
      Select Case XOPT$
         Case "WRITE": GoTo 18100
         Case "READ": GoTo 18200
         Case Else: GoTo 18300
      End Select

18100 ' Open TRANSFER File for output
      On Local Error GoTo 18140
      Open MMDSN$ For Input As #5: Close #5: Kill MMDSN$: GoTo 18150
18140 Resume 18150
18150 On Local Error GoTo 18240
      Open MMDSN$ For Output As #5: OKAY = -1: GoTo 18300

18200 ' Open Exported DATA File for Input
      On Local Error GoTo 18240
      Open MMDSN$ For Input As #5
      If EOF(5) Then
        Call ErrBEEP(0): Call PutMSG("MLMGM007")
        GoTo 18000
      End If
      OKAY = -1: GoTo 18300
18240 If (Not QUIET) Then
         XERR = ERR
         Call ERRMessage("XXXXM030", XERR, ERL)
      End If
      OKAY = 0: Resume 18300

18300 ' Return to Caller
      If OKAY Then
         If Not FDISK Then LSet FT$(12) = MLMGFile$
      End If

End Sub

