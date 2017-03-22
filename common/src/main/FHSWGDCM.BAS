Attribute VB_Name = "modFHSWGdcm"
DefInt A-Z
   Dim INFOSW, GDCMTYPE, GDCMRULE, TAG$()
   Dim NOPAUSE, PrintERR, FIREF, IIREF, TLINE$, RT, RECCNT()
   Dim HIID, HIMRN, CNVTCASE, S, GEDLine&, GEDFile$
   Dim LDiskNUM, DiskNUM, DISKEOF
   Dim INDI, FAMI, MRN, GL
   Dim INGL, INTAG$, XREF, xXREF$, XID, ST$, XVALUE$

Type PtrWrkREC1    ' 20 Bytes
   WFID     As Integer
   WFFID    As Integer
   WFMID    As Integer
   WFCH     As Integer
   WFFSIB   As Integer
   WFMSIB   As Integer
   WFMREC   As Integer
   WFPMREC  As Integer
   WFP      As Integer
   WFSEX    As String * 1
   WFS1     As String * 1
End Type
Type PtrWrkREC2    ' 14 Bytes
   WFMRN    As Integer
   WFHID    As Integer
   WFWID    As Integer
   WFHOF    As Integer
   WFWOF    As Integer
   WF1CH    As Integer
   WFSP     As Integer
End Type
Private Function FNCHR(xX$)
   FNCHR = (((xX$ > "@") And (xX$ < "[")) Or ((xX$ > "`") And (xX$ < "{")))
End Function
Private Function FNNUM(xX$)
   FNNUM = ((xX$ > "/") And (xX$ < ":"))
End Function
Public Sub UtilGDCM(OPT$)
1  Rem $INCLUDE: 'FHSCOMON.BAS'
   PN$ = " FAMILY HISTORY SYSTEM - FHSWGEDCM - GEDCOM Export/Import Program"
   CY$ = " (C) Copyright 1989-1998 by Phillip E. Brown"
   ReDim RECCNT(11), TAG$(140)
   GEDFile$ = TRANPFX$ + FT$(11)
   FUN = 0: READ$ = "READQ"

   Select Case OPT$
      Case "INIT": GoTo 40
      Case "PROCESS": GoTo 1000
      Case "PRINT": GoTo 1000
   End Select
   
40 ' Format Display
   Call PgmPREP("FHSGDCM")
   Load frmFHSWUPgm
   Call RptLOAD(RptCODE): Call RptLOAD("GDLS")
   frmFHSWUPgm.UPSetup.AddItem RptCODE
   frmFHSWUPgm.UPSetup.AddItem "GDLS"
   frmFHSWUPgm.UPSetup_Click
   frmFHSWUPgm.Show
   Exit Sub

1000: ' Process EXPORT Request
   Select Case RptCODE
      Case "GDEX": Call GEDCOMExport
      Case "GDIM": Call GEDCOMImport
      Case "GDLS": Call GEDCOMList
   End Select
End Sub

Rem $DYNAMIC
Sub ConvertADRS()
    Call PgmPREP("FHSGDCM"): If Not OKAY Then GoTo CADRSExit
    Call FF1Open("READ"): If Not OKAY Then GoTo CADRSExit
    Call FF2Open("READ"): If Not OKAY Then GoTo CADRSExit
    Close #2: AFN$ = RTrim$(SysVAR.FILENAME)
    OAF$ = AFN$
    OAFMT = FF2Hdr.REVLVL And 255
    If OAFMT = 1 Then
       Call PutMSG("GDCMM050")
       GoTo CADRSExit
    End If
    Call PutMSG("GDCMM051")
    If CA$ <> SNGLKEY$(1) Then GoTo CADRSExit
    X = InStr(OAF$, ".")
    If X = 0 Then OAF$ = OAF$ + ".OAF" Else OAF$ = Left$(OAF$, X) + "OAF"
    On Local Error GoTo CADRSSkp1
    Kill OAF$
    GoTo CADRSSkp2
CADRSSkp1: Resume CADRSSkp2
CADRSSkp2: On Local Error GoTo CADRSSkp3
    Name AFN$ As OAF$
    SysVAR.FILENAME = OAF$
    Call PutMSG("Old Address File has been renamed: [FILENAME]...")
    If CA$ = Chr$(27) Then
       Name OAF$ As AFN$: GoTo CADRSExit
    End If
    GoTo CADRSSkp4
CADRSSkp3: Call ERRMessage("GDCMM052", ERR, ERL)
    OKAY = 0
    Resume CADRSExit
CADRSSkp4:
    On Local Error GoTo CADRSExit
    Open OAF$ For Binary Access Read As #3
    FF2Hdr.REVLVL = (FF2Hdr.REVLVL \ 256) * 256 + 1
    Open AFN$ For Binary Access Read Write As #2
    Put #2, 1, FF2Hdr
    xX$ = String$(133, 0)
    Put #2, , xX$
    xX$ = Space$(108): Get #3, 1, xX$
    Call PutMSG("GDCMM053")
    For X = 1 To FF2Hdr.HIREC
        Get #3, , FF2PFX: Put #2, , FF2PFX
        Get #3, , FF2SREC
        FF2LREC.PHONE = ZFILL$(FF2SREC.FON1, 3) + "-" + ZFILL$(FF2SREC.FON2, 3) + "-" + ZFILL$(FF2SREC.FON3, 4)
        FF2LREC.LINE1 = FF2SREC.LINE1
        FF2LREC.LINE2 = FF2SREC.LINE2
        LSet FF2LREC.CITY = FF2SREC.CITY
        LSet FF2LREC.STATE = FF2SREC.STATE
        LSet FF2LREC.ZIP = FF2SREC.ZIP
        LSet FF2LREC.COUNTRY = ""
        Put #2, , FF2LREC
    Next
    Call PutMSG("GDCMM054")
    OKAY = -1
CADRSExit:
    Close
End Sub

Sub ConvertCASE(OPT$, TEXT$, S)
    ' OPT$: NAME, TEXT   C: 0=Begin Text, 1=Continued Text
    TextFLD = (OPT$ = "TEXT"): w = 0: If Not TextFLD Then S = 0
    X = Len(RTrim$(TEXT$))
    If X > 0 Then
       For Z = 1 To X
           a = Asc(MID$(TEXT$, Z, 1))
           ' End Sentence at .,!,?
           If (a = 46) Or (a = 21) Or (a = 63) Then
              S = 0: w = 0
              Else
              'Count characters since Beginning of Sentence
              S = S - (a <> 32)
              ' Non-Alpha Character: End of Word
              If (a < 65) Or ((a > 90) And (a < 97)) Or (a > 122) Then
                 w = 0
                 Else
                 ' Lower-Case Character: Stop Conversion
                 If (a > 96) Then
                    Z = X
                    Else
                    ' Upper-Case Character
                    w = w + 1
                    If (w > 1) Or (TextFLD And (S > 1)) Then
                       Mid$(TEXT$, Z, 1) = Chr$(a + 32)
                    End If
                 End If
              End If
           End If
       Next Z
       a = 0
    End If
End Sub

Rem $STATIC
Sub ConvertDATE(XDATE$, XDY, XDMD, XST$)
9010 ' yyyymmdd-->d1y,d1md
     If RTrim$(XDATE$) = "" Then
        XDMD = 0
        Else
        X = 1
        If (GDCMRULE < 3) Or (GDCMRULE = 5) Then GoTo 9030
        If MID$(XDATE$, 1, 1) = "D" Then X = X + 6
        dfmt$ = MID$(XDATE$, X, 8)
        XST$ = MID$(XDATE$, X + 8, 1)
        XDY = Val(MID$(dfmt$, 1, 4))
        XDMD = 100 * Val(MID$(dfmt$, 5, 2)) + Val(MID$(dfmt$, 7, 2))
     End If
     Exit Sub

9030 ' dd mmm yyyy-->d1y,d1md
     dl = Len(XDATE$)
     V1 = 0: V2 = 0: V3 = 0: XST$ = " "
     Y = X: xX$ = MID$(XDATE$, X, 2)
     If (xX$ = "AB") Or (xX$ = "AF") Or (xX$ = "BE") Then
        LSet ST$ = "A": GoSub 9050: GoSub 9052: Y = X
     End If
     GoSub 9051: L = X - Y
     If L > 0 Then
        V1 = Val(MID$(XDATE$, Y, L)): X = X + 1
        If ST$ = " " Then LSet ST$ = MID$(XDATE$, X - 1, 1)
     End If
     If X > dl Then GoTo 9045
     GoSub 9052: Y = X
     GoSub 9050: L = X - Y: If L > 0 Then GoSub 9070: V2 = mm
     If X > dl Then GoTo 9045
     GoSub 9052: Y = X
     GoSub 9051: L = X - Y
     If L > 0 Then
        V3 = Val(MID$(XDATE$, Y, L))
        LSet ST$ = MID$(XDATE$, X, 1)
     End If
9045 If (V2 = 0 And V3 = 0) Then V3 = V1: V1 = 0
     If V3 > 0 Then XDY = V3
     XDMD = 100 * V2 + V1
     Exit Sub

9050 While ((X < (dl + 1)) And FNCHR(MID$(XDATE$, X, 1)))
        X = X + 1
     Wend: Return
9051 While ((X < (dl + 1)) And FNNUM(MID$(XDATE$, X, 1)))
        X = X + 1
     Wend: Return
9052 While (X < (dl + 1)) And (MID$(XDATE$, X, 1) = " ")
        X = X + 1
     Wend: Return


9070 ' mmm-->mm
     MON$ = UCX$(MID$(XDATE$, Y, 3))
     mm = 13
     While (mm > 1) And (MON$ <> MID$(RMONTB$(mm), 3, 3)): mm = mm - 1: Wend
     mm = mm - 1
     Return

End Sub

Sub ConvertNAME(XVALUE$, XSURNM$, XGIVEN$)

    'XVALUE$-->FF1REC.Surnm & FF1REC.Given
     L = Len(RTrim$(XVALUE$)): LSet XSURNM$ = "": LSet XGIVEN$ = ""
     If L > 0 Then
        X = InStr(1, XVALUE$, "/")
        If X > 0 Then
           If X > 1 Then LSet XGIVEN$ = MID$(XVALUE$, 1, X - 1)
           Y = InStr(X + 1, XVALUE$, "/")
           If Y = 0 Then Y = 255
           LSet XSURNM$ = MID$(XVALUE$, X + 1, Y - X - 1)
           Else
           X = InStr(1, XVALUE$, ",")
           If X > 1 Then LSet XSURNM$ = MID$(XVALUE$, 1, X - 1)
           X = X + 1
           If X < L Then LSet XGIVEN$ = LTrim$(MID$(XVALUE$, X, 255))
        End If
        X = InStr(1, XSURNM$, "_")
        While X > 0
           Mid$(XSURNM$, X, 1) = " ": X = InStr(X + 1, XSURNM$, "_")
        Wend
     End If

End Sub

Private Sub ConvertTIME(xXTIME$, XTIME)
    X = InStr(xXTIME$, ":"): Y = Len(RTrim$(xXTIME$))
    If X = 0 Then
       XTIME = Val(Left$(xXTIME$, Y))
       Else
       XTIME = 100 * (Val(Left$(xXTIME$, X - 1)))
       If Y > X Then XTIME = XTIME + Val(MID$(xXTIME$, X + 1, Y - X))
    End If
End Sub

Rem $DYNAMIC
Sub CountRECS()

990 ' Process Record Counts
    If CA$ = Chr$(27) Then Exit Sub
    If RT > 11 Then RT = 11
    RECCNT(RT) = RECCNT(RT) + 1
    'If (Not PrintERR) Or (PRT.OPT <> 1) Then
       If RT < 11 Then
          'COLOR FG(4), BG(4): LOCATE 13 + RT, 43, 0
          'Print USING; "#####"; RECCNT(RT);
          Call frmFHSWUPgm.ShowStats("GEDCOM", RT, Str$(RECCNT(RT)))
          Else
          'LOCATE 25, 1, 0: COLOR FG(1), BG(1)
          'Print Left$(TLINE$, 73); Tab(74);
          'COLOR FG(5), BG(5)
          'If Not NOPAUSE Then
          '   LOCATE 25, 74, 0: Print " ... ";
          '   Call GetKEY
          '   If (Len(CA$) = 2) And (A = 81) Then NOPAUSE = -1
          '   If CA$ <> Chr$(27) Then CA$ = ""
          'End If
          'LOCATE 25, 74, 0: Print USING; "#####"; RECCNT(11);
       End If
    'End If
    If (RT > 10) Then
       If PrintERR And (CA$ <> Chr$(27)) Then
          Mid$(PLINE$, 1, 5) = SFILL$(IIREF, 5)
          Mid$(PLINE$, 7, 5) = SFILL$(FIREF, 5)
          Mid$(PLINE$, 15, 255) = TLINE$
          Call PrintLINE
          If (Len(CA$) = 2) And (a = 81) Then NOPAUSE = -1
          If CA$ <> Chr$(27) Then CA$ = ""
       End If
    End If
End Sub

Sub GEDCOMExport()

600 ' Export Information to GEDCOM formatted dataset
    On Local Error GoTo 601
    ReDim XMONTB$(14): For X = 1 To 14: XMONTB$(X) = RMONTB$(X): Next X
    Call RptLOAD("GDEX"): If Not OKAY Then GoTo 750
    INFOSW = 0
    For X = 0 To 9
        xX$ = "GD" + ZFILL$(X, 2)
        Call RptOptGET(xX$, v): If v = 1 Then INFOSW = INFOSW Or (2 ^ X)
    Next X
    NFMT$ = Space$(55): dfmt$ = Space$(12)
    CLINE$ = Space$(150): XLINE$ = Space$(100)
    xxBLOC$ = Space$(45): xDLOC$ = xxBLOC$: XTIM$ = Space$(5): xxXREF$ = Space$(10)
    AL1$ = Space$(30): AL2$ = AL1$: CITY$ = Space$(20): STATE$ = Space$(12)
    ZIP$ = Space$(12): PHONE$ = Space$(14): CNTRY$ = Space$(Len(FF2LREC.COUNTRY))
    HIREC& = FFRBA&(FF3HDR.HIREC, 1)
    If HIREC& > 32766 Then
       MC1 = 32766: zx& = HIREC& - 32766: If zx& > 32766 Then zx& = 32766
       MC2 = zx&
       Else
       MC1 = HIREC&: MC2 = 0
    End If
    ReDim LoMISCSEQ(MC1), HiMISCSEQ(MC2)
    GoTo 605

601 'Error Setting up for GEDCOM Export
    Call ERRMessage("GDCMM016", ERR, ERL)
    ReDim LoMISCSEQ(1), HiMISCSEQ(1), NDXID(1), RECCNT(1)
    Resume 750

605 ' Determine Records to Process
    'Call PutMSG("GDCMM003")
    'If Asc(CA$) = 27 Then GoTo 750 Else ALLRECS = (CA$ = "1")
    'ReDim SelSET(FF1Hdr.MAXID)
    'If ALLRECS Then
    '   SLCTMAX = FF1Hdr.MAXID
    '   For x = 1 To SLCTMAX: SelSET(x) = -1: Next x
    '   Else
    '   Call SWRKBldList(SelSET(), SLCTMAX, 0)
    '   If Not OKAY Then GoTo 605
    '   For x = 1 To SLCTMAX
    '       If SelSET(x) <> 0 Then SelSET(x) = -1
    '   Next x
    'End If
    If SCNT = 0 Then
       Call PutMSG("No records have been selected...")
       GoTo 750
    End If
    SelSET(0) = 0: SLCTMax = 0
    For X = 1 To FF1Hdr.MAXID
       If SelSET(X) <> 0 Then SelSET(X) = -1: SLCTMax = X
    Next X
    Call frmFHSWUPgm.ShowStats("GEDCOM", -1, 0)

660 ' Determine Sequence of Exported Records
    Call PutMSG("GDCMM004")
    If Asc(CA$) = 27 Then GoTo 690
    If CA$ < "1" Or CA$ > "3" Then Call ErrBEEP(0): GoTo 660
    RELSEQ = (CA$ = "2"): INDEXED = (CA$ = "3")
    ReDim NDXID(FF1Hdr.MAXID)
    If RELSEQ Then
       Call RWRKOpen("READ", 0, 0): If Not OKAY Then GoTo 660
       NDXRECS = 0
       Call PutMSG("XXXXM064")
       Call RWRKGet("BASE", RC, PTR&)
       While (0 < RC) And (RC < 9)
          If WFRELREC.id > 0 Then
             If SelSET(WFRELREC.id) < 0 Then
                NDXRECS = NDXRECS + 1: NDXID(NDXRECS) = WFRELREC.id
                SelSET(WFRELREC.id) = 1
             End If
          End If
          Call RWRKGet("WRKREC", RC, PTR&)
       Wend
       For X = 1 To SLCTMax
           Y = SelSET(X)
           If Y < 0 Then
              NDXRECS = NDXRECS + 1: NDXID(NDXRECS) = X
              Else
              If Y > 0 Then SelSET(X) = -1
           End If
       Next X
       Call PutMSG("")
       Else
       If INDEXED Then
          Call INDEXFile("READ")
          If Not OKAY Then GoTo 660
          NDXRECS = NDXHDR.NDXRECS
          Else
          For I = 1 To FF1Hdr.MAXID: NDXID(I) = I: Next I
          NDXRECS = FF1Hdr.MAXID
       End If
    End If
    HIID = 0
    For X = 1 To NDXRECS
        Y = NDXID(X)
        If SelSET(Y) < 0 Then HIID = HIID + 1: SelSET(Y) = HIID
    Next X
    ReDim Preserve NDXID(NDXRECS)
    For X = 1 To SLCTMax
        If SelSET(X) < 0 Then SelSET(X) = 0
    Next X
    'Call GEDCOMStats

665 ' Get GEDCOM Tags - TAG$()
    Call GetTAGS("EXPORT"): If Not OKAY Then GoTo 750
    If GDCMTYPE = 5 Then
       Call PutMSG("GDCMM006"): If a = 27 Then GoTo 750
       GDCMRULE = Val(CA$)
       If GDCMRULE = 4 Then GDCMRULE = 5
       If (GDCMRULE < 2) Or (GDCMRULE > 5) Then GDCMRULE = 1
       Else
       GDCMRULE = GDCMTYPE
    End If
    Select Case GDCMRULE
       Case 1: LDSv55 = -1: GDRULE$ = "LDSV55"
       Case 2: PAFV21 = -1: GDRULE$ = "PAFV21"
       Case 3: PAFV20 = -1: GDRULE$ = "PAFV20"
       Case 4: FHSv1 = -1: GDRULE$ = "FHSV1"
       Case 5: FHSv2 = -1: GDRULE$ = "FHSV2"
    End Select
    CUST = (GDCMTYPE = 5)
    If (GDCMTYPE < 4) Or ((INFOSW And 1) = 0) Then
       If (GDCMRULE = 2) Then
          xX$ = " 0    1JAN 2FEB 3MAR 4APR 5MAY 6JUN 7JUL 8AUG 9SEP10OCT11NOV12DEC13???"
          Else
          xX$ = " 0    1Jan 2Feb 3Mar 4Apr 5May 6Jun 7Jul 8Aug 9Sep10Oct11Nov12Dec13???"
       End If
       For X = 1 To 14
           RMONTB$(X) = MID$(xX$, 5 * (X - 1) + 1, 5)
       Next X
    End If

    Call PutMSG("GDCMM019"): If a = 27 Then GoTo 750
    CNVTCASE = (CA$ = SNGLKEY$(1))

    ' Create Pointer Workfile
    Dim PtrREC1 As PtrWrkREC1, PtrREC2 As PtrWrkREC2
    Call PointerFILES("OPEN"): If Not OKAY Then GoTo 690

670 ' Enter Header Information
    'Call WinPREP(1): Call WinSWITCH(2): Call WinCLR
    'Call FmtFIND("GDCMS005"): Call WinFORMAT(0)
    'Call PutMSG("GDCMM007")
    'Call WinUPDATE: If A = 27 Then GoTo 690
    'Call PutMSG("")
    'ReDim HLIT$(10)
    Call frmFHSWMsgD.GetGCHdrInfo(HLIT$(), "GET")
    If a = 27 Then GoTo 690
    'x = Fmt.BGNFLD
    'While x > 0
    '   FFLD = FLDTAB(x)
    '   If FFLD.UTXT <> 0 Then
    '      xx$ = UTXT$(FFLD.UTXT)
    '      Select Case FFLD.NAME
    '         Case "NAME": HLIT$(1) = xx$
    '         Case "ADR1": HLIT$(2) = xx$
    '         Case "ADR2": HLIT$(3) = xx$
    '         Case "ADR3": HLIT$(4) = xx$
    '         Case "PHON": HLIT$(5) = xx$
    '         Case "SOUR": HLIT$(6) = xx$
    '         Case "DEST": HLIT$(7) = xx$
    '         Case "COM1": HLIT$(8) = xx$
    '         Case "COM2": HLIT$(9) = xx$
    '         Case "COM3": HLIT$(10) = xx$
    '      End Select
    '   End If
    '   x = FLDCHN(x).FWD
    'Wend
    GoTo 700

690 ' Early Return from EXPORT operation
    Call FamCLOSE
    GoTo 735

700 ' Export Data for Selected Name Records
    Call FAMOpen("READ"): If Not OKAY Then GoTo 690
    Call GEDCOMOpen("WRITE"): If Not OKAY Then GoTo 750
    GoSub 19100                                     ' Build POINTER Workfile
    If (Not OKAY) Or (a = 27) Then GoTo 690
    GL = 0: MISCSEQ = 0: GoSub 800                  ' Create HEADER Record
    'Call WinCLR: Call FmtFIND("GDCMS010"): Call WinFORMAT(0)
    Call PutMSG("GDCMM008")
    On Local Error GoTo 705
    'ON LOCAL ERROR GOTO 0
    GoTo 720

705 ' Error During Export
    XERR = ERR: XERL = ERL: Call GEDCOMStats
    Call ERRMessage("GDCMM017", XERR, XERL)
    OKAY = 0: SID = HIID: I = HIMRN: a = 27: CA$ = Chr$(27)
    Select Case ERRReturn
       Case 1: Resume 721
       Case 2: Resume 979
       Case 3: ComPTR = 0: Resume 981
       Case 4: Resume 750
    End Select
    Resume 730

720 ' Process Records in Pointer Workfiles
    ReDim RECCNT(13): SelSET(0) = 0
    If FHSv1 Or PAFV20 Then IIPFX$ = "" Else IIPFX$ = "I"
    If PAFV20 Then FIPFX$ = "" Else If FHSv1 Then FIPFX$ = "M" Else FIPFX$ = "F"
    For SID = 1 To HIID
        Get #6, SID, PtrREC1: id = PtrREC1.WFID
        Call FF1GetRec(id): GoSub 830
        If (CA$ = Chr$(27)) Or DISKEOF Then
           SID = HIID
           Else
           CA$ = INKEY$
           If CA$ <> "" Then
              K$ = CA$: a = Asc(K$): If a = 27 Then SID = HIID
           End If
        End If
        If (K$ <> "") Or (GEDLine& - HGEDLine& > 200) Then
           Call GEDCOMStats: K$ = "": HGEDLine& = GEDLine&
        End If
721 Next SID
    If (CA$ <> Chr$(27)) And (Not DISKEOF) Then
       GoSub 900    ' Create Family records
       GL = 0
       If Not DISKEOF Then
          Call GEDCOMLine(140, "", "")    ' Create Trailer Record
       End If
    End If

730 ' End of EXPORT
    ERRReturn = 4
    Call GEDCOMStats
    Call PutMSG("GDCMM009")
735 Call PutMSG("")
    'Call WinPREP(2): Call WinCLR
750 Call FamCLOSE
    Call frmFHSWUPgm.ShowStats("GEDCOM", 99, "")
    Call PointerFILES("DELETE")
    For X = 1 To 14: RMONTB$(X) = XMONTB$(X): Next X
    Exit Sub

800 ' Create GEDCOM File Header Record
    GL = 0
    Call GEDCOMLine(1, "", "")
    GL = GL + 1
    If Not (PAFV20 Or FHSv1) Then
       HL = 5: CTAG = 10: GoSub 810
       HL = 6: CTAG = 13: GoSub 810
       LSet dfmt$ = DATE$: GoSub 9090
       Call GEDCOMLine(14, "", dfmt$)
       Call GEDCOMLine(15, "", GEDFile$)
    End If
    If TAG$(16) <> TAG$(0) Then Call GEDCOMLine(16, "", LTrim$(Str$(HIID)))
    If TAG$(17) <> TAG$(0) Then Call GEDCOMLine(17, "", LTrim$(Str$(HIMRN)))
    If TAG$(20) <> TAG$(0) Then Call GEDCOMLine(20, "", GDRULE$)
    If LDSv55 Then
       TAG$(21) = "GEDC": TAG$(22) = "VERS"
       Call GEDCOMLine(21, "", "")
       GL = GL + 1: Call GEDCOMLine(22, "", "5.5"): GL = GL - 1
    End If
    If PAFV21 Or LDSv55 Then
       GL = GL - 1: Call GEDCOMLine(4, "@S1@", "")
       GL = GL + 1
    End If
    If PAFV20 Then Call GEDCOMLine(4, "", ""): GL = GL + 1
    HL = 1: CTAG = 5: GoSub 810
    CTAG = 6: GoSub 810
    If Not (FHSv1 Or PAFV20) Then GL = GL + 1
    CTAG = 7: GoSub 810
    If Not (FHSv1 Or PAFV20) Then GL = GL - 1
    CTAG = 8: GoSub 810
    HL = 7
    CTAG = 18: GoSub 810
    If Not (FHSv1 Or PAFV20) Then GL = GL + 1
    CTAG = 19: GoSub 810
    If (FHSv1 Or PAFV20) Then
       If Not FHSv1 Then GL = GL - 1: Call GEDCOMLine(9, "", ""): GL = GL + 1
       HL = 5: CTAG = 10: GoSub 810: CTAG = 13: GoSub 810
    End If
    GL = GL - 1
    Return

810 ' Create Heading Lines
    If (RTrim$(HLIT$(HL)) <> "") And (HLIT$(HL) <> "(000) 000-0000") Then
       Call GEDCOMLine(CTAG, "", HLIT$(HL))
    End If
    HL = HL + 1
    Return

830 ' Create NAME Record Lines
    GL = 0: RT = 1: Call CountRECS
    LSet xxXREF$ = "@" + IIPFX$ + fnx$(SID) + "@"
    Call GEDCOMLine(30, xxXREF$, "")
    GL = GL + 1
    Call GEDCOMLine(31, "", fnx$(PtrREC1.WFID))
    GoSub 9110: Call GEDCOMLine(32, "", NFMT$)
    S1 = Asc(FF1Rec.STATUS): S2 = Asc(FF1Rec.STATUS2)
    If ((S2 And 128) > 0) And (TAG$(33) <> TAG$(0)) Then
       Call GEDCOMLine(33, "", "Y")
    End If
    If FF1Rec.SEX <> " " Then Call GEDCOMLine(34, "", FF1Rec.SEX)

840 ' Birth Information
    BLOCPTR = FF1Rec.BLOC: BPL = 0: DPL = 0
    If (((INFOSW And 2) > 0) And (BLOCPTR <> 0)) Then
       Call FF3GetRec(BLOCPTR)
       If (FF3PFX.RTYPE = Chr$(12)) And (FF3PLACE2.PNEXT <> 0) Then
          Call FF3GetRec(FF3PLACE2.PNEXT)
       End If
       LSet xxBLOC$ = BPLACE
       LSet xDLOC$ = DPLACE
       Else
       LSet xxBLOC$ = " ": LSet xDLOC$ = " "
    End If
    D1Y = FF1Rec.BY: D1MD = FF1Rec.BMD: TIM = FF1Rec.BTM
    If FHSv1 Then
       TD = 36: TT = 38: TP = 39
       Else
       TD = 75: TT = 77: TP = 78
    End If
    BPL = Len(RTrim$(xxBLOC$))
    If (D1Y = 0) And (D1MD = 0) And ((TIM = 0) Or (TAG$(TT) = TAG$(0))) And (BPL = 0) Then
       GoTo 850
    End If
    If Not FHSv1 Then Call GEDCOMLine(35, "", ""): GL = GL + 1
    If (D1Y > 0) Or (D1MD > 0) Or ((S1 And 48) > 0) Then
       GoSub 9000: X = Len(RTrim$(dfmt$)) + 1
       If PAFV20 Then
          Mid$(dfmt$, X, 1) = Chr$(32 + (32 - Asc("A")) * ((S1 And 16) = 16))
          Else
          If FHSv1 Then
             Mid$(dfmt$, X, 1) = FDSTAT$("BD")
          End If
       End If
       Call GEDCOMLine(TD, "", dfmt$)
       If Not (PAFV20 Or FHSv1) Then
          If ((S1 And 48) <> 0) And (TAG$(76) <> TAG$(0)) Then
             'IF (S1 AND 32) > 0 THEN xx$ = "3" ELSE xx$ = "0"
             xX$ = FDSTAT$("BD")
             Call GEDCOMLine(76, "", xX$)
          End If
       End If
    End If
    If (TIM > 0) And (TAG$(TT) <> TAG$(0)) Then
       If FHSv1 Then
          LSet XTIM$ = ZFILL$(TIM, 4)
          Else
          LSet XTIM$ = ZFILL$(TIM \ 100, 2) + ":" + ZFILL$(TIM Mod 100, 2)
       End If
       Call GEDCOMLine(TT, "", XTIM$)
    End If
    If BPL > 0 Then
       RT = 2: Call CountRECS
       Call GEDCOMLine(TP, "", xxBLOC$)
    End If
    If Not FHSv1 Then GL = GL - 1

850 ' Death Information
    If FHSv1 Then TD = 41: TT = 43: TP = 45
    D1Y = FF1Rec.DY: D1MD = FF1Rec.DMD: TIM = FF1Rec.DTM
    DPL = Len(RTrim$(xDLOC$))
    If (D1Y = 0) And (D1MD = 0) And ((TIM = 0) Or (TAG$(TT) = TAG$(0))) And (DPL = 0) And ((S1 And 192) = 0) Then
       GoTo 860
    End If
    If Not FHSv1 Then Call GEDCOMLine(40, "", ""): GL = GL + 1
    If (((D1Y > 0) And ((D1Y < 9999) Or (FHSv1 And (D1Y = 9999)))) Or D1MD > 0 Or ((S1 And 192) > 0)) Then
       GoSub 9000: X = Len(RTrim$(dfmt$)) + 1
       If PAFV20 Then
          Mid$(dfmt$, X, 1) = Chr$(32 + (32 - Asc("A")) * ((S1 And 64) = 64))
          Else
          If FHSv1 Then
             Mid$(dfmt$, X, 1) = FDSTAT$("DD")
          End If
       End If
       Call GEDCOMLine(TD, "", dfmt$)
       If Not (PAFV20 Or FHSv1) Then
          If ((S1 And 192) <> 0) And (TAG$(76) <> TAG$(0)) Then
             'IF (S1 AND 128) > 0 THEN xx$ = "3" ELSE xx$ = "0"
             xX$ = FDSTAT$("DD")
             Call GEDCOMLine(76, "", xX$)
          End If
       End If
    End If
    If (TIM > 0) And (TAG$(TT) <> TAG$(0)) Then
       If FHSv1 Then
          LSet XTIM$ = ZFILL$(TIM, 4)
          Else
          LSet XTIM$ = ZFILL$(TIM \ 100, 2) + ":" + ZFILL$(TIM Mod 100, 2)
       End If
       Call GEDCOMLine(TT, "", XTIM$)
    End If
    If DPL > 0 Then
       RT = 2: Call CountRECS
       Call GEDCOMLine(TP, "", xDLOC$)
    End If
    If Not FHSv1 Then GL = GL - 1

860 ' Name Record Relationships
    If Not (FHSv1 Or FHSv2) Then GoTo 865
    ST$ = " "
    X = PtrREC1.WFFID: T = 61: LSet ST$ = FDSTAT$("FID"): GoSub 861
    X = PtrREC1.WFMID: T = 63: LSet ST$ = FDSTAT$("MID"): GoSub 861
    X = PtrREC1.WFFSIB: T = 62: GoSub 861
    X = PtrREC1.WFMSIB: T = 64: GoSub 861
    X = PtrREC1.WFCH: T = 60: GoSub 861
    If ((INFOSW And 4) > 0) And (FF1Rec.SPOUSE <> 0) Then
       zx& = FFRBA&(FF1Rec.SPOUSE, 1)
       If zx& < 32767 Then
          MRN = LoMISCSEQ(zx&)
          Else
          MRN = HiMISCSEQ(zx& - 32766)
       End If
       LSet xxXREF$ = "@M" + fnx$(MRN) + "@"
       Call GEDCOMLine(51, "", xxXREF$)
       Else
       MRN = 0
    End If
    GoTo 880

861 ' Process FHS Relationship Field
    If X > 0 Then
       LSet xxXREF$ = "@" + IIPFX$ + fnx$(X) + "@" + ST$
       Call GEDCOMLine(T, "", xxXREF$)
       LSet ST$ = ""
    End If
    Return

865 ' Parent Marriage Record
    MRN = PtrREC1.WFPMREC: If MRN = 0 Then GoTo 870
    LSet xxXREF$ = "@" + FIPFX$ + fnx$(MRN) + "@"
    If Not PAFV20 Then
       Call GEDCOMLine(54, "", xxXREF$)
       If (LDSv55 And ((S1 And 5) > 0) And (TAG$(55) <> TAG$(0))) Then
          GL = GL + 1
          If (S1 And 5) = 1 Then
             LSet xxXREF$ = "HUSB"
             Else
             If (S1 And 5) = 4 Then
                LSet xxXREF$ = "WIFE"
                Else
                LSet xxXREF$ = "BOTH"
             End If
          End If
          Call GEDCOMLine(55, "", xxXREF$)
          GL = GL - 1
       End If
       GoTo 870
    End If
    Call GEDCOMLine(54, "", "")
    GL = GL + 1
    Call GEDCOMLine(31, "", xxXREF$)
    If ((S1 And 1) > 0) And (TAG$(57) <> TAG$(0)) Then Call GEDCOMLine(57, "", "")
    If ((S1 And 4) > 0) And (TAG$(56) <> TAG$(0)) Then Call GEDCOMLine(56, "", "")
    GL = GL - 1

870 ' Own Marriage Record(s)
    MRN = Abs(PtrREC1.WFMREC)
871 If MRN = 0 Then GoTo 875
    LSet xxXREF$ = "@" + FIPFX$ + fnx$(MRN) + "@"
    If PAFV20 Then
       Call GEDCOMLine(50, "", "")
       GL = GL + 1: Call GEDCOMLine(31, "", xxXREF$)
       GL = GL - 1: GoTo 875
    End If
    Call GEDCOMLine(50, "", xxXREF$)
    Get #7, MRN, PtrREC2
    If SID = PtrREC2.WFHID Then
       MRN = PtrREC2.WFHOF
       Else
       MRN = PtrREC2.WFWOF
    End If
    MRN = Abs(MRN)
    GoTo 871

875 ' Find Older Sibling with Same Parents (PAF 2.0)
    If Not PAFV20 Then GoTo 880
    FID = Abs(PtrREC1.WFFID): xxMID = Abs(PtrREC1.WFMID): PID = FID
    If PID > 0 Then
       SIBPTR = PtrREC1.WFFSIB
       Else
       PID = xxMID
       If PID > 0 Then SIBPTR = PtrREC1.WFMSIB Else GoTo 880
    End If
    XID = 0
    While SIBPTR > 0: Get #6, SIBPTR, PtrREC1
       xMID = Abs(PtrREC1.WFMID): xFID = (PtrREC1.WFFID)
       If (FID = xFID) And (xxMID = xMID) Then
          XID = SIBPTR: SIBPTR = 0
          Else
          If PID = FID Then SIBPTR = PtrREC1.WFFSIB Else SIBPTR = PtrREC1.WFMSIB
       End If
    Wend
    If XID = 0 Then GoTo 880
    Call GEDCOMLine(58, "", "")
    GL = GL + 1: LSet xxXREF$ = "@" + fnx$(XID) + "@"
    Call GEDCOMLine(59, "", xxXREF$)
    GL = GL - 1

880 ' Process NAME Record Comments
    ComPTR = FF1Rec.COM: CTAG = 79: GoSub 980

    ' Process Address Records
    ADRPTR = FF1Rec.ADRS: GoSub 1000

    ' Process Events & Miscellaneous Records
    XPTR = FF1Rec.EVENT: GoSub 2010
    XPTR = FF1Rec.EDUC: xType = 1: GoSub 2020
    XPTR = FF1Rec.WORK: xType = 2: GoSub 2020
    XPTR = FF1Rec.MIL: xType = 3: GoSub 2020
    XPTR = FF1Rec.HLTH: xType = 4: GoSub 2020

899 ' Return from Name Record Routines
    GL = GL - 1
    Return

900 ' Produce Family Record Lines
    If (HIMRN = 0) Then Return
    ERRReturn = 2
    GL = 0
    For I = 1 To HIMRN
    RT = 3: Call CountRECS
    Get #7, I, PtrREC2: ComPTR = 0
    HID = PtrREC2.WFHID: WID = PtrREC2.WFWID
    HOF = Abs(PtrREC2.WFHOF): WOF = Abs(PtrREC2.WFWOF)
    MRN = PtrREC2.WFMRN
    If MRN <> 0 Then
       Call FF3GetRec(MRN): ComPTR = FF3SPOUSE.COM
    End If
    LSet xxXREF$ = "@" + FIPFX$ + fnx$(I) + "@"
    Call GEDCOMLine(110, xxXREF$, "")
    GL = GL + 1
    If (MRN = 0) And (TAG$(111) <> TAG$(0)) Then Call GEDCOMLine(111, "", "")
    If HID <> 0 Then
       LSet xxXREF$ = "@" + IIPFX$ + fnx$(HID) + "@"
       If PAFV20 Then
          Call GEDCOMLine(112, "", "")
          GL = GL + 1: Call GEDCOMLine(113, "", xxXREF$)
          If HOF > 0 Then
             LSet xxXREF$ = "@" + FIPFX$ + fnx$(HOF)
             Call GEDCOMLine(114, "", xxXREF$)
          End If
          GL = GL - 1
          Else
          Call GEDCOMLine(112, "", xxXREF$)
       End If
    End If
    If WID <> 0 Then
       LSet xxXREF$ = "@" + IIPFX$ + fnx$(WID) + "@"
       If PAFV20 Then
          Call GEDCOMLine(115, "", "")
          GL = GL + 1: Call GEDCOMLine(113, "", xxXREF$)
          If WOF < 0 Then WOF = -WOF
          If WOF > 0 Then
             LSet xxXREF$ = "@" + FIPFX$ + fnx$(WOF)
             Call GEDCOMLine(114, "", xxXREF$)
          End If
          GL = GL - 1
          Else
          Call GEDCOMLine(115, "", xxXREF$)
       End If
    End If
    If FHSv1 Or FHSv2 Then
       If HOF > 0 Then
          LSet xxXREF$ = "@" + FIPFX$ + fnx$(HOF) + "@"
          Call GEDCOMLine(130, "", xxXREF$)
       End If
       If WOF > 0 Then
          LSet xxXREF$ = "@" + FIPFX$ + fnx$(WOF) + "@"
          Call GEDCOMLine(131, "", xxXREF$)
       End If
    End If
    NXCH = PtrREC2.WF1CH
    If NXCH = 0 Then GoTo 920 ' FHS Rules will always have NXCH=0
    If PAFV20 Then
       Call GEDCOMLine(132, "", "")
       LSet xxXREF$ = "@" + IIPFX$ + fnx$(NXCH) + "@"
       GL = GL + 1: Call GEDCOMLine(133, "", xxXREF$)
       GL = GL - 1: GoTo 920
    End If
    PID = HID: If PID = 0 Then PID = WID: If PID = 0 Then GoTo 920
    Get #6, NXCH, PtrREC1

910 ' Print Line for Each Child for PAF v2.1, LDS v5.5
    LSet xxXREF$ = "@" + IIPFX$ + fnx$(NXCH) + "@"
    Call GEDCOMLine(132, "", xxXREF$)
    S1 = Asc(PtrREC1.WFS1)
    If (PAFV21 And ((S1 And 5) > 0) And (TAG$(134) <> TAG$(0))) Then
       If (S1 And 5) = 1 Then
          LSet xXREF$ = "HUSB"
          Else
          If (S1 And 5) = 4 Then
             LSet xXREF$ = "WIFE"
             Else
             LSet xXREF$ = "BOTH"
          End If
       End If
       GL = GL + 1: Call GEDCOMLine(134, "", xXREF$)
       GL = GL - 1
    End If

915 ' Find Next Child
    If PID = PtrREC1.WFFID Then
       NXCH = PtrREC1.WFFSIB
       Else
       If PID = PtrREC1.WFMID Then NXCH = PtrREC1.WFMSIB Else NXCH = 0
    End If
    If NXCH > 0 Then
       Get #6, NXCH, PtrREC1
       ' The Folowing Line assumes Husband<-->Father and Wife<-->Mother
       If HID = (PtrREC1.WFFID) And (WID = PtrREC1.WFMID) Then GoTo 910
       GoTo 915
    End If

920 ' Process Marriage Date & Place Fields
    If MRN = 0 Then GoTo 975
    MLOC = 0: LSet MPLACE = " ": LSet DVPLACE = " "
    D1Y = FF3SPOUSE.D1Y: D1MD = FF3SPOUSE.D1MD
    D2Y = FF3SPOUSE.D2Y: D2MD = FF3SPOUSE.D2MD
    If (INFOSW And 2) > 0 Then
       MLOC = FF3SPOUSE.MLOC
       If MLOC <> 0 Then
          Call FF3GetRec(MLOC)
          If (FF3PFX.RTYPE = Chr$(12)) And (FF3PLACE2.PNEXT <> 0) Then
             Call FF3GetRec(FF3PLACE2.PNEXT)
          End If
       End If
    End If
    MPL = Len(RTrim$(MPLACE))
    DVPL = Len(RTrim$(DVPLACE))
    If (FHSv1 Or FHSv2) Then
       If RTrim$(FF3SPOUSE.STATUS1) <> "" Then
          Call GEDCOMLine(121, "", FF3SPOUSE.STATUS1)
       End If
       If (D1Y > 0) Or (D1MD > 0) Then
          GoSub 9000: Call GEDCOMLine(122, "", dfmt$)
       End If
       If MPL > 0 Then
          RT = 2: Call CountRECS
          Call GEDCOMLine(124, "", MPLACE)
       End If
       If RTrim$(FF3SPOUSE.STATUS2) <> "" Then
          Call GEDCOMLine(127, "", FF3SPOUSE.STATUS2)
       End If
       If (D2Y > 0) Or (D2MD > 0) Then
          D1Y = D2Y: D1MD = D2MD: GoSub 9000
          Call GEDCOMLine(128, "", dfmt$)
       End If
       If DVPL > 0 Then
          RT = 2: Call CountRECS
          Call GEDCOMLine(129, "", DVPLACE)
       End If
       GoTo 970
    End If

    If (D1Y > 0) Or (D1MD > 0) Or (MPL > 0) Then
       If LDSv55 Then
          Call GEDCOMLine(70, "", ""): GL = GL + 1
          If RTrim$(FF3SPOUSE.STATUS1) <> "" Then
             Call GEDCOMLine(71, "", FF3SPOUSE.STATUS1)
          End If
          Else
          Call GEDCOMLine(120, "", ""): GL = GL + 1
       End If
       If (D1Y > 0) Or (D1MD > 0) Then
          GoSub 9000: Call GEDCOMLine(75, "", dfmt$)
       End If
       If (MPL > 0) Then
          RT = 2: Call CountRECS
          Call GEDCOMLine(78, "", MPLACE)
       End If
       GL = GL - 1
    End If

    If LDSv55 Then
       If (D2Y > 0) Or (D2MD > 0) Or (RTrim$(FF3SPOUSE.STATUS2) <> "") Or (DVPL > 0) Then
          Call GEDCOMLine(70, "", ""): GL = GL + 1
          Call GEDCOMLine(71, "", "TERM")
          If RTrim$(FF3SPOUSE.STATUS2) <> "" Then
             Call GEDCOMLine(71, "", FF3SPOUSE.STATUS2)
          End If
          If ((D2Y > 0) And ((D2Y < 9999) Or (FHSv1 And (D2Y = 9999)))) Or (D2MD > 0) Then
             D1Y = D2Y: D1MD = D2MD
             GoSub 9000: Call GEDCOMLine(75, "", dfmt$)
          End If
          If (DVPL > 0) Then
             RT = 2: Call CountRECS
             Call GEDCOMLine(78, "", DVPLACE)
          End If
          GL = GL - 1
       End If
       GoTo 970
    End If

    ' Check Divorced Status v2.0-2.1
    xX$ = UCX$(Left$(FF3SPOUSE.STATUS2, 7))
    DIVORCED = (xX$ = "DIVORCE")
    If Not DIVORCED Then GoTo 970
    If (TAG$(125) = TAG$(0)) Then
       If TAG$(126) <> TAG$(0) Then
          Call GEDCOMLine(126, "", "Y")
       End If
       GoTo 970
    End If
    Call GEDCOMLine(125, "", "")
    GL = GL + 1: D1Y = FF3SPOUSE.D2Y: D1MD = FF3SPOUSE.D2MD
    If ((D2Y > 0) And ((D2Y < 9999) Or (FHSv1 And (D2Y = 9999)))) Or (D2MD > 0) Then
       D1Y = D2Y: D1MD = D2MD: GoSub 9000: Call GEDCOMLine(75, "", dfmt$)
    End If
    If DVPL > 0 Then
       RT = 2: Call CountRECS
       Call GEDCOMLine(78, "", DVPLACE)
    End If
    GL = GL - 1

970 ' Marriage Record Comments
    CTAG = 135: GoSub 980

    ' Marriage Record Addresses
    ADRPTR = FF3SPOUSE.LOC: GoSub 1000

975 GL = GL - 1
    CA$ = INKEY$
    If CA$ <> "" Then
       K$ = CA$
       If CA$ = Chr$(27) Then
          I = HIMRN
          Else
          SysVAR.NVAR1 = I: SysVAR.NVAR2 = HIMRN
          Call PutMSG("GDCMM010")
       End If
    End If
    If (K$ <> "") Or (GEDLine& - HGEDLine& > 200) Then
        Call GEDCOMStats: K$ = "": HGEDLine& = GEDLine&
    End If

979 Next I
    Return

980 ' Export Comment Lines
    If ((INFOSW And 2 ^ 9) = 0) Or (TAG$(CTAG) = TAG$(0)) Or (ComPTR = 0) Then Return
    HGL = GL: NGL = GL + 1: NTAG = CTAG + 1: S = 0
    OFST = 1: LSet CLINE$ = " "
    ERRReturn = 3
    While ComPTR <> 0
       RT = 10: Call CountRECS
       Call FF3GetRec(ComPTR): ComPTR = FF3COMMENT.FPTR
       Mid$(CLINE$, OFST, 38) = FF3COMMENT.TEXT
       OFST = InStr(OFST, CLINE$, "   ") - (MID$(FF3COMMENT.TEXT, 38, 1) = " ")
       If OFST > 69 Then
          OFST = 70: FOUND = 0
          While (OFST > 0) And (Not FOUND)
             FOUND = (MID$(CLINE$, OFST, 1) = " ")
             If Not FOUND Then OFST = OFST - 1
          Wend
          If OFST = 0 Then OFST = 70
          LSet XLINE$ = Left$(CLINE$, OFST)
          Call GEDCOMLine(CTAG, "", XLINE$): S = 1
          LSet CLINE$ = MID$(CLINE$, OFST + 1, 255)
          OFST = InStr(1, CLINE$, "   ") - (MID$(FF3COMMENT.TEXT, 38, 1) = " ")
          If Not (PAFV20 Or FHSv1) Then GL = NGL
          CTAG = NTAG
       End If
981 Wend
    If a <> 27 Then
       If OFST > 0 Then
          LSet XLINE$ = LTrim$(Left$(CLINE$, OFST - 1))
          Call GEDCOMLine(CTAG, "", XLINE$)
       End If
    End If
    GL = HGL
    Return

1000 ' Export Address Records
     If ((INFOSW And 2 ^ 8) = 0) Or (TAG$(100) = TAG$(0)) Then Return
     AFMT = FF2Hdr.REVLVL And 1
1020 While ADRPTR <> 0
        RT = 9: Call CountRECS
        Call GEDCOMLine(100, "", "")
        GL = GL + 1: XDATE$ = ""
        Call FF2GetRec(ADRPTR): ADRPTR = FF2PFX.NXT
        D1Y = FF2PFX.D1Y: D1MD = FF2PFX.D1MD
1021    If (D1Y > 0) Or (D1MD > 0) Then
           GoSub 9000
           If FHSv1 Or FHSv2 Then
              Call GEDCOMLine(81, "", dfmt$)
              Else
              XDATE$ = "from " + LTrim$(RTrim$(dfmt$)) + " "
           End If
        End If
        D1Y = FF2PFX.D2Y: D1MD = FF2PFX.D2MD
1022    If (D1Y > 0) Or (D1MD > 0) Then
           GoSub 9000
           If FHSv1 Or FHSv2 Then
              Call GEDCOMLine(82, "", dfmt$)
              Else
              XDATE$ = XDATE$ + "until " + LTrim$(RTrim$(dfmt$))
           End If
        End If
        If XDATE$ <> "" Then Call GEDCOMLine(75, "", XDATE$)
1023    Select Case AFMT
           Case 0: LSet AL1$ = FF2SREC.LINE1: LSet AL2$ = FF2SREC.LINE2
                   LSet CITY$ = FF2SREC.CITY: LSet STATE$ = FF2SREC.STATE
                   LSet ZIP$ = FF2SREC.ZIP: LSet CNTRY$ = "": GoSub 9200
           Case 1: LSet AL1$ = FF2LREC.LINE1: LSet AL2$ = FF2LREC.LINE2
                   LSet CITY$ = FF2LREC.CITY: LSet STATE$ = FF2LREC.STATE
                   LSet ZIP$ = FF2LREC.ZIP: LSet CNTRY$ = FF2LREC.COUNTRY
                   LSet PHONE$ = FF2LREC.PHONE
        End Select
1024    If Len(RTrim$(AL1$)) > 0 Then
           Call GEDCOMLine(101, "", AL1$)
           If (Len(RTrim$(AL2$)) > 0) Then
              Call GEDCOMLine(102, "", AL2$)
           End If
        End If
1025    If Len(RTrim$(CITY$)) > 0 Then
           Call GEDCOMLine(103, "", CITY$)
        End If
1026    If Len(RTrim$(STATE$)) > 0 Then
           Call GEDCOMLine(104, "", STATE$)
        End If
1027    If Len(RTrim$(ZIP$)) > 0 Then
           Call GEDCOMLine(105, "", ZIP$)
        End If
1028    If Len(RTrim$(CNTRY$)) > 0 Then
           Call GEDCOMLine(106, "", CNTRY$)
        End If
1029    If Len(RTrim$(PHONE$)) > 0 Then Call GEDCOMLine(107, "", PHONE$)
        ComPTR = FF2PFX.COM: CTAG = 108: GoSub 980
        GL = GL - 1
     Wend
     Return

2010 ' Export EVENT Records
     If ((INFOSW And 2 ^ 7) > 0) And (TAG$(70) <> TAG$(0)) Then
        While XPTR <> 0
           RT = 8: Call CountRECS
           Call FF3GetRec(XPTR): XPTR = FF3EVENT.NXT
           Call GEDCOMLine(70, "", ""): GL = GL + 1
           Call GEDCOMLine(71, "", FF3EVENT.CODE)
           Call GEDCOMLine(72, "", Str$(Asc(FF3EVENT.IMPORT)))
           If (FF3EVENT.D1Y > 0) Or (FF3EVENT.D1MD > 0) Then
              D1Y = FF3EVENT.D1Y: D1MD = FF3EVENT.D1MD: GoSub 9000
              Call GEDCOMLine(75, "", dfmt$)
              If (FF3EVENT.DSTAT <> " ") And (TAG$(76) <> TAG$(0)) Then
                 Call GEDCOMLine(76, "", FF3EVENT.DSTAT)
              End If
           End If
           If RTrim$(FF3EVENT.PLACE) <> "" Then
              Call GEDCOMLine(78, "", FF3EVENT.PLACE)
           End If
           ComPTR = FF3EVENT.COM: CTAG = 79: GoSub 980
           CA$ = INKEY$
           If CA$ <> "" Then K$ = CA$: If CA$ = Chr$(27) Then XPTR = 0
           GL = GL - 1
        Wend
     End If
     Return

2020 ' Export EDUC,WORK,MIL,HLTH Records
     Select Case xType
        Case 1: XTAG = 86
        Case 2: XTAG = 90
        Case 3: XTAG = 92
        Case 4: XTAG = 95
     End Select
     If ((INFOSW And 2 ^ (xType + 2)) > 0) And (TAG$(XTAG) <> TAG$(0)) Then
        While XPTR <> 0
           RT = xType + 3: Call CountRECS
           Call FF3GetRec(XPTR): XPTR = FF3PNTRS.NXT
           Call GEDCOMLine(XTAG, "", ""): GL = GL + 1: XDATE$ = ""
           D1Y = FF3PNTRS.D1Y: D1MD = FF3PNTRS.D1MD
           If (D1Y > 0) Or (D1MD > 0) Then
              GoSub 9000
              If FHSv1 Or FHSv2 Then
                 Call GEDCOMLine(81, "", dfmt$)
                 Else
                 XDATE$ = "from " + LTrim$(RTrim$(dfmt$)) + " "
              End If
           End If
           D1Y = FF3PNTRS.D2Y: D1MD = FF3PNTRS.D2MD
           If (D1Y > 0) Or (D1MD > 0) Then
              GoSub 9000
              If FHSv1 Or FHSv2 Then
                 Call GEDCOMLine(82, "", dfmt$)
                 Else
                 XDATE$ = XDATE$ + "until " + LTrim$(RTrim$(dfmt$))
              End If
           End If
           If XDATE$ <> "" Then Call GEDCOMLine(75, "", XDATE$)
           Select Case xType
              Case 1: If RTrim$(FF3EDUC.EDLEV) <> "" Then Call GEDCOMLine(88, "", FF3EDUC.EDLEV)
                      If RTrim$(FF3EDUC.EDSUB1) <> "" Then Call GEDCOMLine(87, "", FF3EDUC.EDSUB1)
                      If RTrim$(FF3EDUC.EDSUB2) <> "" Then Call GEDCOMLine(87, "", FF3EDUC.EDSUB2)
                      If RTrim$(FF3EDUC.EDDEG) <> "" Then Call GEDCOMLine(89, "", FF3EDUC.EDDEG)
              Case 2: If RTrim$(FF3WORK.WORKTP) <> "" Then Call GEDCOMLine(91, "", FF3WORK.WORKTP)
              Case 3: If RTrim$(FF3MIL.MILRANK) <> "" Then Call GEDCOMLine(93, "", FF3MIL.MILRANK)
                      If RTrim$(FF3MIL.MILSTAT) <> "" Then Call GEDCOMLine(83, "", FF3MIL.MILSTAT)
              Case 4: If RTrim$(FF3HEALTH.HLTHDIAG) <> "" Then Call GEDCOMLine(96, "", FF3HEALTH.HLTHDIAG)
                      If RTrim$(FF3HEALTH.HLTHSTAT) <> "" Then Call GEDCOMLine(83, "", FF3HEALTH.HLTHSTAT)
           End Select
           ComPTR = FF3PNTRS.COM: CTAG = 84: GoSub 980
           ADRPTR = FF3PNTRS.LOC: GoSub 1000
           CA$ = INKEY$:
           If CA$ <> "" Then K$ = CA$: If CA$ = Chr$(27) Then XPTR = 0
           GL = GL - 1
        Wend
     End If
     Return

9000 'd1y,d1md-->YYYYMMDD
     If Not (PAFV20 Or FHSv1) Then
        dfmt$ = XRPTDate$(D1MD, D1Y, 0)
        Else
        L = 8: LSet dfmt$ = "00000000"
        If (D1Y > 0) Or (D1MD > 0) Then
           DM = Int(D1MD / 100): DD = D1MD Mod 100
           If D1Y > 0 Then Mid$(dfmt$, 1, 4) = Right$("000" + fnx$(D1Y), 4)
           If DM > 0 Then Mid$(dfmt$, 5, 2) = Right$("0" + fnx$(DM), 2)
           If DD > 0 Then Mid$(dfmt$, 7, 2) = Right$("0" + fnx$(DD), 2)
        End If
     End If
     Return

9090 ' MM-DD-YYYY-->dd mmm yyyy
     D1Y = Val(MID$(dfmt$, 7, 4))
     D1MD = 100 * (Val(MID$(dfmt$, 1, 2))) + Val(MID$(dfmt$, 4, 2))
     dfmt$ = XRPTDate$(D1MD, D1Y, 0)
     Return

9110 'nfmt$=f1given$/f1surnm$/
     If CNVTCASE Then
        Call ConvertCASE("NAME", FF1Rec.GIVEN, S)
        Call ConvertCASE("NAME", FF1Rec.SURNM, S)
     End If
     If Not FHSv1 Then
        LSet NFMT$ = FF1Rec.GIVEN: Z = 1
        Z = InStr(Z, NFMT$, "   ")
        Mid$(NFMT$, Z, 1) = "/": Mid$(NFMT$, Z + 1, 30) = FF1Rec.SURNM
        NLTH = InStr(Z, NFMT$, "   ")
        Mid$(NFMT$, NLTH, 1) = "/"
        Else
        LSet NFMT$ = FF1Rec.SURNM
        X = Len(RTrim$(NFMT$))
        Mid$(NFMT$, X + 1, 255) = ", " + FF1Rec.GIVEN
     End If
     Return

9200 ' Format Phone Number (Short Address Format)
     F1 = FF2SREC.FON1: F2 = FF2SREC.FON2: F3 = FF2SREC.FON3
     If (F1 > 0) Or (F2 > 0) Or (F3 > 0) Then
        If FHSv1 Then
           LSet PHONE$ = "   -   -    "
           Mid$(PHONE$, 1, 3) = ZFILL$(F1, 3)
           Mid$(PHONE$, 5, 3) = ZFILL$(F2, 3)
           Mid$(PHONE$, 9, 4) = ZFILL$(F3, 4)
           Else
           LSet PHONE$ = "(000) 000-0000"
           Mid$(PHONE$, 2, 3) = ZFILL$(F1, 3)
           Mid$(PHONE$, 7, 3) = ZFILL$(F2, 3)
           Mid$(PHONE$, 11, 4) = ZFILL$(F3, 4)
        End If
        Else
        LSet PHONE$ = ""
     End If
     Return

19100 ' Build Pointer Workfile
      OKAY = -1: Call PutMSG("GDCMM011")
      On Local Error GoTo 19105
      SelSET(0) = 0: HIMRN = 0: HIID = 0
      ERRReturn = 1
      GoTo 19110

19105 ' Error Routine
      Call ERRMessage("GDCMM018", ERR, ERL)
      OKAY = 0
      Select Case ERRReturn
         Case 1: ix = NDXRECS: Resume 19111
         Case 2: SID = HIID: Resume 19141
         Case 3: SID = HIID: Resume 19195
      End Select
      Resume 19199

19110 For ix = 1 To NDXRECS: id = NDXID(ix)
          SID = SelSET(id): If SID = 0 Then GoTo 19111
          Call FF1GetRec(id)
          CA$ = INKEY$
          If CA$ <> "" Then
             If Asc(CA$) = 27 Then OKAY = 0: ix = NDXRECS: GoTo 19111
             SysVAR.NVAR1 = ix: SysVAR.NVAR2 = NDXRECS
             Call PutMSG("GDCMM012")
          End If
          FID = SelSET(FF1Rec.FID): xxMID = SelSET(FF1Rec.MID)
          SPPTR = FF1Rec.SPOUSE
          SEX$ = FF1Rec.SEX
          PtrREC1.WFID = id
          PtrREC1.WFFID = FID: PtrREC1.WFMID = xxMID: PtrREC1.WFSEX = SEX$
          PtrREC1.WFS1 = FF1Rec.STATUS
          FCH = FF1Rec.FCH: MCH = FF1Rec.MCH
          XID = FF1Rec.OLDCH: PID = id: GoSub 19200: PtrREC1.WFCH = XID
          If FID > 0 Then
             XID = FCH: PID = FID: GoSub 19200: FCH = XID
             Else
             FCH = 0
          End If
          PtrREC1.WFFSIB = FCH
          If xxMID > 0 Then
             XID = MCH: PID = xxMID: GoSub 19200: MCH = XID
             Else
             MCH = 0
          End If
          PtrREC1.WFMSIB = MCH
          Put #6, SID, PtrREC1: GoSub 19230
          Get #6, SID, PtrREC1: PtrREC1.WFMREC = SPPTR
          Put #6, SID, PtrREC1
          HIID = SID
19111 Next ix
      If Not OKAY Then GoTo 19199
      If FHSv1 Or FHSv2 Then GoTo 19199

19140 'Reverse Directions of Sibling Pointers for PAF 2.0
      If Not PAFV20 Then GoTo 19150
      Call PutMSG("GDCMM013")
      ERRReturn = 2
      For SID = 1 To HIID
          Get #6, SID, PtrREC1
          CHPTR = PtrREC1.WFCH: SIBPTR = 0
          While CHPTR <> 0: XID = CHPTR: Get #6, XID, PtrREC1
             If SID = PtrREC1.WFFID Then
                CHPTR = PtrREC1.WFFSIB: PtrREC1.WFFSIB = SIBPTR
                Else
                CHPTR = PtrREC1.WFMSIB: PtrREC1.WFMSIB = SIBPTR
             End If
             Put #6, XID, PtrREC1: SIBPTR = XID
          Wend
          Get #6, SID, PtrREC1: PtrREC1.WFCH = SIBPTR
          Put #6, SID, PtrREC1
          CA$ = INKEY$
          If CA$ <> "" Then
             If Asc(CA$) = 27 Then OKAY = 0: SID = HIID
             SysVAR.NVAR1 = SID: SysVAR.NVAR2 = HIID
             Call PutMSG("GDCMM012")
          End If
19141 Next SID: If Not OKAY Then Return

19150 ' Find Parent Family Record
      CFREC = 0: ERRReturn = 3
      If (INFOSW And 4) = 0 Then GoTo 19199
      Call PutMSG("GDCMM014"): SysVAR.NVAR2 = HIID
      For SID = 1 To HIID
          SysVAR.NVAR1 = SID
          Get #6, SID, PtrREC1
          CA$ = INKEY$
          If CA$ <> "" Then
             If Asc(CA$) = 27 Then OKAY = 0: SID = HIID: GoTo 19195
             Call PutMSG("GDCMM012")
          End If
          If PtrREC1.WFPMREC <> 0 Then GoTo 19195
          FID = PtrREC1.WFFID: xxMID = PtrREC1.WFMID
          PID = FID: If PID = 0 Then PID = xxMID: If PID = 0 Then GoTo 19195
          Get #6, PID, PtrREC1
          PFPTR = 0: xMRN = PtrREC1.WFMREC
          While xMRN > 0
             Get #7, xMRN, PtrREC2
             If ((FID = PtrREC2.WFHID) And (xxMID = PtrREC2.WFWID)) Or ((FID = PtrREC2.WFWID) And (xxMID = PtrREC2.WFHID)) Then
                PFPTR = xMRN: xMRN = 0
                Else
                If PID = PtrREC2.WFHID Then xMRN = PtrREC2.WFHOF Else xMRN = PtrREC2.WFWOF
             End If
             CA$ = INKEY$: If CA$ = Chr$(27) Then xMRN = 0
          Wend
          If CA$ = Chr$(27) Then OKAY = 0: SID = HIID: GoTo 19195

          ' Create NEW Family Record for Unmarried Parents
          If PFPTR = 0 Then
             CFREC = CFREC + 1: HIMRN = HIMRN + 1
             PFPTR = HIMRN: HOF = 0: WOF = 0
             If FID > 0 Then
                Get #6, FID, PtrREC1: HOF = PtrREC1.WFMREC
                PtrREC1.WFMREC = PFPTR
                Put #6, FID, PtrREC1
             End If
             If xxMID > 0 Then
                Get #6, xxMID, PtrREC1: WOF = PtrREC1.WFMREC
                PtrREC1.WFMREC = PFPTR
                Put #6, xxMID, PtrREC1
             End If
             PtrREC2.WFMRN = 0
             PtrREC2.WFHID = FID: PtrREC2.WFWID = xxMID
             PtrREC2.WFHOF = HOF: PtrREC2.WFWOF = WOF
             PtrREC2.WF1CH = 0
             Put #7, HIMRN, PtrREC2
          End If

          ' Update Parent family record for all siblings
          Get #6, PID, PtrREC1: N1CH = 0: SIBPTR = PtrREC1.WFCH
          While SIBPTR > 0
             Get #6, SIBPTR, PtrREC1
             If ((FID = PtrREC1.WFFID) And (xxMID = PtrREC1.WFMID)) Or ((xxMID = PtrREC1.WFFID) And (FID = PtrREC1.WFMID)) Then
                PtrREC1.WFPMREC = PFPTR: Put #6, SIBPTR, PtrREC1
                If N1CH = 0 Then N1CH = SIBPTR
             End If
             If PID = PtrREC1.WFFID Then
                SIBPTR = PtrREC1.WFFSIB
                Else
                SIBPTR = PtrREC1.WFMSIB
             End If
             CA$ = INKEY$: If CA$ = Chr$(27) Then SIBPTR = 0
             If CA$ <> "" Then
                SysVAR.NVAR3 = SIBPTR
                Call PutMSG("Processing [NVAR1] of [NVAR2], Sibling=[NVAR3]")
             End If
          Wend
          If CA$ = Chr$(27) Then OKAY = 0: SID = HIID: GoTo 19195

          ' Update Parent family record with first sibling (same FID, xxMID)
          If N1CH > 0 Then
             Get #7, PFPTR, PtrREC2: PtrREC2.WF1CH = N1CH
             Put #7, PFPTR, PtrREC2
          End If

19195 ' Next ENTRY
      Next SID: If Not OKAY Then GoTo 19199
      If CFREC > 0 Then
         SysVAR.NVAR1 = CFREC
         Call PutMSG("GDCMM015")
      End If

19199 ' Return from Building Pointer WorkFiles
      On Local Error GoTo 0
      Return

19200 ' Locate First Selected Sibling Parented by PID (Starting with XID)
      While (XID <> 0) And (SelSET(XID) = 0)
         Call FF1GetRec(XID)
         If PID = FF1Rec.FID Then
            XID = FF1Rec.FCH
            Else
            If PID = FF1Rec.MID Then
               XID = FF1Rec.MCH
               Else
               XID = 0
            End If
         End If
      Wend
      XID = SelSET(XID)
      Return

19230 ' Create Pointer File Records for Existing Marriages
      ' - In Most Recent to Least Recent order for FHS
      ' - In Chronological sequence for PAF or LDS
      FFPTR = 0: LFPTR = 0: If ((INFOSW And 4) = 0) Then GoTo 19245
19232 If SPPTR = 0 Then GoTo 19245
      zx& = FFRBA&(SPPTR, 1)
      If zx& < 32767 Then
         xMRN = LoMISCSEQ(zx&)
         Else
         xMRN = HiMISCSEQ(zx& - 32766)
      End If
      If xMRN > 0 Then GoTo 19240
      HIMRN = HIMRN + 1
      If zx& < 32767 Then
         LoMISCSEQ(zx&) = HIMRN
         Else
         HiMISCSEQ(zx& - 32766) = HIMRN
      End If
      Call FF3GetRec(SPPTR)
      OFPTR = 0: OSPID = 0: OSPOF = 0: HOF = 0: WOF = 0
      If id = FF3SPOUSE.SP1ID Then
         OFPTR = FF3SPOUSE.SP1NXT: OSPID = FF3SPOUSE.SP2ID
         OSPOF = FF3SPOUSE.SP2NXT
         Else
         If id = FF3SPOUSE.SP2ID Then
            OFPTR = FF3SPOUSE.SP2NXT: OSPID = FF3SPOUSE.SP1ID
            OSPOF = FF3SPOUSE.SP1NXT
         End If
      End If
      PtrREC2.WFMRN = SPPTR
      If SEX$ = FGENDR$(1) Then
         HID = SID: SPPTR = OFPTR
         If (FHSv1 Or FHSv2) Then HOF = OFPTR Else HOF = LFPTR
         WID = SelSET(OSPID): If WID > 0 Then WOF = OSPOF
         Else
         WID = SID: SPPTR = OFPTR
         If (FHSv1 Or FHSv2) Then WOF = OFPTR Else WOF = LFPTR
         HID = SelSET(OSPID): If HID > 0 Then HOF = OSPOF
      End If
      PtrREC2.WFHID = HID: PtrREC2.WFWID = WID
      PtrREC2.WFHOF = HOF: PtrREC2.WFWOF = WOF
      PtrREC2.WF1CH = 0
      xMRN = HIMRN
      GoTo 19241

19240 ' Process Other Marriage Records
      Get #7, xMRN, PtrREC2
      If SID = PtrREC2.WFHID Then
         SPPTR = PtrREC2.WFHOF
         If Not (FHSv1 Or FHSv2) Then PtrREC2.WFHOF = LFPTR
         Else
         If SID = PtrREC2.WFWID Then
            SPPTR = PtrREC2.WFWOF
            If Not (FHSv1 Or FHSv2) Then PtrREC2.WFWOF = LFPTR
            Else
            SPPTR = 0
         End If
      End If
19241 Put #7, xMRN, PtrREC2
      If FFPTR = 0 Then FFPTR = xMRN
      If FHSv1 Or FHSv2 Then
         If LFPTR > 0 Then
            Get #7, LFPTR, PtrREC2
            If SID = PtrREC2.WFHID Then
               PtrREC2.WFHOF = xMRN
               Else
               PtrREC2.WFWOF = xMRN
            End If
            Put #7, LFPTR, PtrREC2
         End If
      End If
      LFPTR = xMRN
      GoTo 19232

19245 ' Finished with Marriage Records for SID
      If (FHSv1 Or FHSv2) Then SPPTR = FFPTR Else SPPTR = LFPTR
      Return

End Sub

Sub GEDCOMImport()

     ' Import Information from GEDCOM format dataset
     ReDim XMONTB$(14)
     For X = 1 To 14
         XMONTB$(X) = RMONTB$(X)
     Next X
     Call GEDCOMOpen("READ"): If OKAY Then GoTo 1010

1009 ' Exit Before Beginning Import
     Call ClearMSG: Close
     Call frmFHSWUPgm.ShowStats("GEDCOM", 99, "")
     For X = 1 To 14: RMONTB$(X) = XMONTB$(X): Next X
     Exit Sub

1010 ' Prepare for Import
     F1RID = 0: F1MAXID = FF1Hdr.MAXID: GL = 0: NameCNT = 0: PrintERR = 0
     xBLOC$ = Space$(50): xDLOC$ = xBLOC$: xMLOC$ = xBLOC$: DVLOC$ = xBLOC$
     dfmt$ = Space$(15): CLINE$ = Space$(150)
     ReDim RECCNT(11), HLIT$(10)
     Call RptOPEN("GDIM"): If Not OKAY Then GoTo 1009
     Call GetRVAR("SEENOTES", X)
     If FOUND Then SEENOTE$ = RVAR$(X) Else SEENOTE$ = "see notes"
     Call GetRVAR("BORN", X)
     If FOUND Then BORN$ = RVAR$(X) + ":" Else BORN$ = "Born:"
     Call GetRVAR("MARRIED", X)
     If FOUND Then MARRIED$ = RVAR$(X) Else MARRIED$ = "Married"
     Call GetRVAR("DIED", X)
     If FOUND Then DIED$ = RVAR$(X) + ":" Else DIED$ = "Died:"
     Call GetRVAR("PARNTREC", X)
     If FOUND Then PARNTREC$ = RVAR$(X) Else PARNTREC$ = "Parent REC"
     INFOSW = 0
     For X = 0 To 9
         xX$ = "GD" + ZFILL$(X, 2)
         Call RptOptGET(xX$, v)
         If v = 1 Then INFOSW = INFOSW Or (2 ^ X)
     Next X
     If FF1Hdr.MAXID > 0 Then
        SysVAR.NVAR1 = FF1Hdr.MAXID + 1
        Call PutMSG("GDCMM020")
        If Asc(CA$) = 27 Then GoTo 1009
     End If
     Call GetTAGS("IMPORT"): If Not OKAY Then GoTo 1009
     CUST = (GDCMTYPE = 5)

BeginIMPORT:
     On Local Error GoTo 1999
     INTAG$ = Space$(Len(TAG$(0))): ST$ = " ": XVALUE$ = Space$(100)
     Call FAMOpen("WRITE"): If Not OKAY Then GoTo 1009
     AFMT = FF2Hdr.REVLVL And 1: XGEDLine& = GEDLine&
     Call GEDCOMOpen("READQ"): If (Not OKAY) Or (a = 27) Then GoTo 1009
     Call PointerFILES("OPEN"): If Not OKAY Then GoTo 1009
     GL = 0: Call GetLINE: If Not OKAY Then GoTo 1009
     If (INGL > 0) Or (INTAG$ <> TAG$(1)) Then
        Call PutMSG("GDCMM022")
        If (CA$ <> SNGLKEY$(1)) Then GoTo 1009
        GoTo 1100
     End If
     FIREF = 0: IIREF = 0

1050 ' Get Header Lines
     On Local Error GoTo 0
     If CA$ = Chr$(27) Then GoTo 1009
     Call GetLINE: If (Not OKAY) Or EOF(5) Then GoTo 1009
     If (INGL < (GL + 1)) And (INTAG$ <> TAG$(4)) Then GoTo 1060
     EndOfHDR& = Seek(5): HDRLine& = GEDLine&
     If (INTAG$ = TAG$(4)) Or (INTAG$ = TAG$(9)) Or (INTAG$ = "CHEC") Then GoTo 1050
     If (INTAG$ = TAG$(14)) Or (INTAG$ = TAG$(15)) Then GoTo 1050
     Select Case INTAG$
        Case TAG$(5): HT = 1
        Case TAG$(6): HT = 2
        Case TAG$(8): HT = 4
        Case TAG$(10): HT = 5
        Case TAG$(13): HT = 6
        Case TAG$(18): HT = 7
        Case TAG$(16): MAXRID = Val(RTrim$(XVALUE$)): GoTo 1050
        Case TAG$(17): MAXMRID = Val(RTrim$(XVALUE$)): GoTo 1050
        Case TAG$(20):
             Select Case RTrim$(XVALUE$)
               Case "LDSV55": GDCMRULE = 1
               Case "PAFV21": GDCMRULE = 2
               Case "PAFV20": GDCMRULE = 3
               Case "FHSV1": GDCMRULE = 4
               Case "FHSV2": GDCMRULE = 5
             End Select
             GoTo 1050
        Case "GEDC":
             Call GetLINE: If INTAG$ = "VERS" Then GDCMRULE = 1
             GoTo 1050
        Case Else: RT = 99: Call CountRECS: GoTo 1050
     End Select
     If (INTAG$ = TAG$(7)) Or (INTAG$ = TAG$(19)) Then
        If (HT < 4) Or ((HT > 4) And (HT < 8)) Or (HT > 8) Then
           RT = 99: Call CountRECS: GoTo 1050
        End If
     End If
1055 ' Put Value in HLIT$()
     HLIT$(HT) = RTrim$(XVALUE$)
     HT = HT + 1
     GoTo 1050

1060 'Call WinSWITCH(2)
     Call frmFHSWUPgm.ShowStats("GEDCOM", -1, "")
     If MAXRID > 0 Then
        INDI = MAXRID: FAMI = MAXMRID: Call GEDCOMStats
        INDI = 0: FAMI = 0
     End If
    'Determine GEDCOM Rules
     If GDCMRULE = 0 Then
        If GDCMTYPE = 5 Then
           Call PutMSG("GDCMM006"): If a = 27 Then GoTo 1009
           GDCMRULE = Val(CA$)
           If GDCMRULE = 4 Then GDCMRULE = 5
           If (GDCMRULE < 1) Or (GDCMRULE > 4) Then GDCMRULE = 1
           Else
           GDCMRULE = GDCMTYPE
        End If
     End If
     Select Case GDCMRULE
        Case 1: LDSv55 = -1
        Case 2: PAFV21 = -1
        Case 3: PAFV20 = -1
        Case 4: FHSv1 = -1
        Case 5: FHSv2 = -1
     End Select
     If (GDCMTYPE < 4) Or ((INFOSW And 1) = 0) Then
        xX$ = " 0    1JAN 2FEB 3MAR 4APR 5MAY 6JUN 7JUL 8AUG 9SEP10OCT11NOV12DEC13???"
        For X = 1 To 14
            RMONTB$(X) = MID$(xX$, 5 * (X - 1) + 1, 5)
        Next X
     End If
     For X = 1 To 14: RMONTB$(X) = UCX$(RMONTB$(X)): Next X

1065 ' Show Header Information
     'Call WinCLR
     'Call FmtFIND("GDCMS005")
     'x = Fmt.BGNFLD
     'While x > 0
     '   FFLD = FLDTAB(x)
     '   If FFLD.UTXT <> 0 Then
     '      xx$ = UTXT$(FFLD.UTXT)
     '      Select Case FFLD.NAME
     '         Case "NAME": xx$ = HLIT$(1)
     '         Case "ADR1": xx$ = HLIT$(2)
     '         Case "ADR2": xx$ = HLIT$(3)
     '         Case "ADR3": xx$ = HLIT$(4)
     '         Case "PHON": xx$ = HLIT$(5)
     '         Case "SOUR": xx$ = HLIT$(6)
     '         Case "DEST": xx$ = HLIT$(7)
     '         Case "COM1": xx$ = HLIT$(8)
     '         Case "COM2": xx$ = HLIT$(9)
     '         Case "COM3": xx$ = HLIT$(10)
     '      End Select
     '      LSet UTXT$(FFLD.UTXT) = xx$
     '   End If
     '   x = FLDCHN(x).FWD
     'Wend
     'Call WinFORMAT(0)
     Call frmFHSWMsgD.GetGCHdrInfo(HLIT$(), "SHOW")
     
     Call PutMSG("GDCMM024")
     Unload frmFHSWMsgD
     If ((CA$ = SNGLKEY$(2)) Or (Asc(CA$) = 27)) Then GoTo 1900
     If Not ENV.LPNSupport Then
        If (Not (FHSv1 Or FHSv2)) And ((INFOSW And 2 ^ 8) > 0) Then
           Call PutMSG("GDCMM023")
           If Asc(CA$) = 27 Then GoTo 1009
           NotePLACES = (CA$ = SNGLKEY$(1))
        End If
     End If
     'Call PutMSG("GDCMM021")
     'If Asc(CA$) = 27 Then GoTo 1009
     'PrintERR = (CA$ = SNGLKEY$(1))
     PrintERR = -1
     If PrintERR Then Call PrintOPEN: If Not OKAY Then GoTo 1009

     Call PutMSG("GDCMM025")
     Dim PtrREC1 As PtrWrkREC1, PtrREC2 As PtrWrkREC2
     Dim NewPtrREC1 As PtrWrkREC1, NewPtrREC2 As PtrWrkREC2
     ReDim IIRNO(MAXRID), FIRNO(MAXMRID)
     HIID = 0: HIMRN = 0: CA$ = ""

1070 ' Scan GEDCOM File and Build Pointer Work Files
     If (INGL > 0) Then If SUBSW = 0 Then GoTo 1071 Else GoTo 1075
     If XREF < 0 Then GoTo 1071
     ' Record Level Entries
     If INTAG$ = TAG$(30) Then GoTo 1072 ' Individual Records
     If INTAG$ = TAG$(110) Then GoTo 1074 ' Family Records
     If INTAG$ = TAG$(140) Then GoTo 1080 ' Trailer Record
1071 ' Unprocessed Record Level Lines
     RT = 99: Call CountRECS: If (CA$ = Chr$(27)) Then GoTo 1995
     GoTo 1079

1072 ' New Individual Record
     If XREF = 0 Then GoTo 1071
     If SUBSW > 0 Then
        If (HIID > 0) And (SUBSW < 4) Then Put #6, HIID, PtrREC1
        If (HIMRN > 0) And (SUBSW > 4) Then Put #7, HIMRN, PtrREC2
        SUBSW = 0
     End If
     If XREF > MAXRID Then
        zx& = XREF + 100: If zx& > 32766 Then zx& = 32766
        MAXRID = zx&: ReDim Preserve IIRNO(MAXRID)
     End If
     If IIRNO(XREF) <> 0 Then GoTo 1071
     PtrREC1 = NewPtrREC1: PtrREC1.WFID = XREF
     HIID = HIID + 1: IIRNO(XREF) = HIID
     SUBSW = 1
     GoTo 1079

1074 ' New Family Record
     If XREF = 0 Then GoTo 1071
     If SUBSW > 0 Then
        If (HIID > 0) And (SUBSW < 4) Then Put #6, HIID, PtrREC1
        If (HIMRN > 0) And (SUBSW > 4) Then Put #7, HIMRN, PtrREC2
        SUBSW = 0
     End If
     If XREF > MAXMRID Then
        zx& = XREF + 100: If zx& > 32766 Then zx& = 32766
        MAXMRID = zx&: ReDim Preserve FIRNO(MAXMRID)
     End If
     If FIRNO(XREF) <> 0 Then GoTo 1071
     PtrREC2 = NewPtrREC2
     HIMRN = HIMRN + 1: FIRNO(XREF) = HIMRN
     SUBSW = 5
     GoTo 1079

1075 ' INDI Record Relationship TAGs
     If SUBSW > 4 Then GoTo 1077
     Select Case INTAG$
        Case TAG$(31):
             Select Case SUBSW   ' v2.0 family pointers
                Case 2: PtrREC1.WFMREC = XID
                Case 3: PtrREC1.WFPMREC = XID
             End Select
        Case TAG$(34): PtrREC1.WFSEX = Left$(XVALUE$, 1)
        Case TAG$(50): If PAFV20 Then SUBSW = 2
        Case TAG$(51): PtrREC1.WFMREC = XID
        Case TAG$(54): If (PAFV20 Or LDSv55) Then SUBSW = 3
        Case TAG$(55): ' Adoption Status (LDSv55)
                       S1 = Asc(PtrREC1.WFS1)
                       Select Case Left$(XVALUE$, 4)
                          Case "HUSB": S1 = S1 Or 1
                          Case "WIFE": S1 = S1 Or 4
                          Case "FATH": S1 = S1 Or 1
                          Case "MOTH": S1 = S1 Or 4
                          Case "BOTH": S1 = S1 Or 5
                       End Select
                       PtrREC1.WFS1 = Chr$(S1)
        Case TAG$(56): PtrREC1.WFS1 = Chr$(Asc(PtrREC1.WFS1) Or 1)
        Case TAG$(57): PtrREC1.WFS1 = Chr$(Asc(PtrREC1.WFS1) Or 4)
        Case TAG$(59): ' v2.0 older sibling
                       PtrREC1.WFFSIB = XID: PtrREC1.WFMSIB = XID
        Case TAG$(60): PtrREC1.WFCH = XID
        Case TAG$(61): PtrREC1.WFFID = XID
                       S1 = Asc(PtrREC1.WFS1)
                       Select Case ST$
                         Case "*": S1 = S1 Or 1
                         Case "?": S1 = S1 Or 2
                       End Select
                       PtrREC1.WFS1 = Chr$(S1)
        Case TAG$(62): PtrREC1.WFFSIB = XID
        Case TAG$(63): PtrREC1.WFMID = XID
                       S1 = Asc(PtrREC1.WFS1)
                       Select Case ST$
                         Case "*": S1 = S1 Or 4
                         Case "?": S1 = S1 Or 8
                       End Select
                       PtrREC1.WFS1 = Chr$(S1)
        Case TAG$(64): PtrREC1.WFMSIB = XID
     End Select
     GoTo 1079

1077 ' FAMI Record Relationship TAGs
     Select Case INTAG$
        Case TAG$(112): If PAFV20 Then SUBSW = 6 Else PtrREC2.WFHID = XID
        Case TAG$(113): If SUBSW = 6 Then PtrREC2.WFHID = XID Else PtrREC2.WFWID = XID
        Case TAG$(114): If SUBSW = 6 Then PtrREC2.WFHOF = XID Else PtrREC2.WFWOF = XID
        Case TAG$(115): If PAFV20 Then SUBSW = 7 Else PtrREC2.WFWID = XID
        Case TAG$(130): PtrREC2.WFHOF = XID
        Case TAG$(131): PtrREC2.WFWOF = XID
        Case TAG$(132): If PAFV20 Then GoTo 1079 ' CHIL tag
                        ' IF NOT PAFV20 then 1 Line for each CHILd
                        ' Link together Last child in Marriage-->First (as with 2.0?)
                        ' NOTE: These lines assume that the II records for both parents
                        ' and children occur in GED file prior to the FI records
                        CRNO = IIRNO(XID)
                        If CRNO <> 0 Then
                           Get #6, CRNO, PtrREC1
                          'PtrREC1.WFFID = PtrREC2.WFHID
                          'PtrREC1.WFMID = PtrREC2.WFWID
                           'If PtrREC2.WFHID = 243 Then
                           '   Call PutMSG("Waiting at 1077 for HID=" + Str$(PtrREC2.WFHID) + "...")
                           '   Call GetRESP
                           'End If
                           'NOTE: This will not link children of different marriages
                           '      That must be done in paragraph 1950
                           If (PtrREC2.WFHID <> 0) Then PtrREC1.WFFSIB = PtrREC2.WF1CH
                           If (PtrREC2.WFWID <> 0) Then PtrREC1.WFMSIB = PtrREC2.WF1CH
                           Put #6, CRNO, PtrREC1
                        End If
                        PtrREC2.WF1CH = XID
        Case TAG$(133): ' Youngest Child (PAFV20)
                        PtrREC2.WF1CH = XID
        Case TAG$(134): ' Adoption Status (PAFV21)
                        If CRNO <> 0 Then
                           S1 = Asc(PtrREC1.WFS1)
                           Select Case Left$(XVALUE$, 4)
                              Case "HUSB": S1 = S1 Or 1
                              Case "WIFE": S1 = S1 Or 4
                              Case "FATH": S1 = S1 Or 1
                              Case "MOTH": S1 = S1 Or 4
                              Case "BOTH": S1 = S1 Or 5
                           End Select
                           PtrREC1.WFS1 = Chr$(S1)
                           Put #6, CRNO, PtrREC1
                        End If
     End Select

1079 ' Next Line
     Call GetLINE: If (CA$ = Chr$(27)) Then GoTo 1995
     If DISKEOF Or (INTAG$ = TAG$(140)) Then GoTo 1080
     If CA$ = "" Then CA$ = INKEY$
     If CA$ <> "" Then
        Call PutMSG("GDCMM025")
        NOPAUSE = 0: K$ = CA$
        If CA$ = Chr$(27) Then Call GetRESP: If CA$ = Chr$(27) Then GoTo 1995
        CA$ = ""
     End If
     If (K$ <> "") Or (GEDLine& - HGEDLine& > 200) Then
        'If (Not PrintERR) Or (PRT.OPT <> 1) Then
           Call GEDCOMStats
        'End If
        K$ = "": HGEDLine& = GEDLine&
     End If
     If Not DISKEOF Then GoTo 1070

1080 ' Finish Work Files
     If SUBSW > 0 Then
        If (SUBSW < 5) Then
           Put #6, HIID, PtrREC1
           Else
           Put #7, HIMRN, PtrREC2
        End If
        SUBSW = 0
     End If
     ReDim MISCSEQ(HIMRN)
     Call ClearMSG
     XGEDLine& = GEDLine&
     Call GEDCOMStats

1090 ' Return to End Of Header Section
     Call GEDCOMOpen("READQ"): If Not OKAY Then GoTo 1995
     Seek #5, EndOfHDR&
     GEDLine& = HDRLine&
     Call GetLINE

     ' Display Format for Record Counts
     'If (Not PrintERR) Or (PRT.opt <> 1) Then
     'Call WinCLR: Call FmtFIND("GDCMS010"): Call WinFORMAT(0)
     'End If
     HGEDLine& = 0: CA$ = ""
     'Call PutMSG("GDCMM030")

1100 ' Process Next GEDCOM Record
     If (CA$ = Chr$(27)) Then GoTo 1995
     If CA$ = "" Then CA$ = INKEY$
     If CA$ <> "" Then
        K$ = CA$
        If CA$ = Chr$(27) Then GoTo 1995
        If (Len(CA$) = 2) And (a = 81) Then NOPAUSE = -1 Else NOPAUSE = 0
        If PrintERR And (PRT.OPT = 1) Then
           SysVAR.NVAR1 = II: SysVAR.NVAR2 = MRID
           Call PutMSG("GDCMM026")
        End If
        CA$ = ""
     End If
     If (K$ <> "") Or (GEDLine& - HGEDLine& > 200) Then
       If Not (PrintERR And (PRT.OPT = 1)) Then Call GEDCOMStats
       K$ = "": HGEDLine& = GEDLine&
     End If
     If DISKEOF Then GoTo 1900
     If INTAG$ = TAG$(140) Then GoTo 1900 ' Trailer Record
     If (INGL > 0) Or (XREF = 0) Then GoTo 1190
     SUBSW = 0: OLDCH = 0: HID = 0: WID = 0: GL = 0
     If INTAG$ = TAG$(110) Then GoTo 1200 ' Family Records
     If INTAG$ = TAG$(30) Then GoTo 1300 ' Individual Records
1190 RT = 99: Call CountRECS
     Call GetLINE
     GoTo 1100

1200 ' Process Family Information Record
     If ((XREF = 0) Or (XREF > MAXMRID) Or ((INFOSW And 4) = 0)) Then GoTo 1190
     GL = GL + 1
     FIREF = XREF: MRID = FIRNO(XREF): If MRID = 0 Then GoTo 1190
     Get #7, MRID, PtrREC2: MRN = PtrREC2.WFMRN
     FF3PFX.RTYPE = Chr$(3): Call FF3GetRec(MRN): If Not OKAY Then GoTo 1190
     If MRN = 0 Then
        RT = 3: Call CountRECS
        MRN = FF3PFX.RNO: MISCSEQ(MRID) = MRN
        PtrREC2.WFMRN = MRN: Put #7, MRID, PtrREC2
        FF3SPOUSE.SRTYPE = Chr$(1): LSet FF3SPOUSE.STATUS1 = MARRIED$
        LSet FF3SPOUSE.STATUS2 = " "
        Call FF3PutRec
     End If
     LSet xMLOC$ = " ": LSet DVLOC$ = " ": MPL = 0: DVPL = 0
     LAREC = 0: LCOM = 0
     SRNO = MRN: SRTYPE = 3: LAREC = 0

1210 ' Process FAMI Record Lines
     Call GetLINE
1215 If INGL < GL Then GoTo 1290
     If INGL < GL + 1 Then SUBSW = 0
     If INGL > (GL - (SUBSW > 0)) Then GoTo 1289
     If INTAG$ = "CHEC" Then GoTo 1210
     If INTAG$ = TAG$(111) Then
        LSet FF3SPOUSE.STATUS1 = PARNTREC$: GoTo 1210
     End If
     If (INTAG$ = TAG$(112)) Or (INTAG$ = TAG$(115)) Then
        If PAFV20 Then SUBSW = 99
        GoTo 1210
     End If
     If (INTAG$ = TAG$(113)) Or (INTAG$ = TAG$(114)) Then GoTo 1210
     If FHSv1 Or FHSv2 Then
        If INTAG$ = TAG$(121) Then
           LSet FF3SPOUSE.STATUS1 = XVALUE$: GoTo 1210
        End If
        If INTAG$ = TAG$(122) Then
           FF3SPOUSE.D1Y = 0
           Call ConvertDATE(XVALUE$, FF3SPOUSE.D1Y, FF3SPOUSE.D1MD, ST$)
           GoTo 1210
        End If
        If INTAG$ = TAG$(124) Then
           MPL = Len(RTrim$(XVALUE$))
           If (MPL = 0) Or ((INFOSW And 2) = 0) Then GoTo 1289
           LSet xMLOC$ = XVALUE$
           GoTo 1210
        End If
        If INTAG$ = TAG$(127) Then
           LSet FF3SPOUSE.STATUS2 = XVALUE$: GoTo 1210
        End If
        If INTAG$ = TAG$(128) Then
           FF3SPOUSE.D2Y = 9999
           Call ConvertDATE(XVALUE$, FF3SPOUSE.D2Y, FF3SPOUSE.D2MD, ST$)
           GoTo 1210
        End If
        If INTAG$ = TAG$(129) Then
           DVPL = Len(RTrim$(XVALUE$))
           If (DVPL = 0) Or ((INFOSW And 2) = 0) Then GoTo 1289
           LSet DVLOC$ = XVALUE$
           GoTo 1210
        End If
     End If

     ' Marriage Status, Date, Place
     If INTAG$ = TAG$(70) Then SUBSW = 13: GoTo 1210
     If INTAG$ = TAG$(120) Then SUBSW = 13: GoTo 1210
     If SUBSW = 13 Then
        If INTAG$ = TAG$(71) Then
           If RTrim$(XVALUE$) = "TERM" Then SUBSW = 14: GoTo 1210
           LSet FF3SPOUSE.STATUS1 = XVALUE$: GoTo 1210
        End If
        If (INTAG$ = TAG$(75)) Then
           FF3SPOUSE.D1Y = 0
           Call ConvertDATE(XVALUE$, FF3SPOUSE.D1Y, FF3SPOUSE.D1MD, ST$)
           GoTo 1210
        End If
        If (INTAG$ = TAG$(78)) Then
           MPL = Len(RTrim$(XVALUE$))
           If (MPL = 0) Or ((INFOSW And 2) = 0) Then GoTo 1289
           LSet xMLOC$ = XVALUE$
           If (MPL < (23 - 19 * ENV.LPNSupport)) Then GoTo 1210
           If NotePLACES Then GoSub 1285: GoTo 1210
           GoTo 1289
        End If
     End If

     ' "Divorce" Status, Date, Place
     If (INTAG$ = TAG$(125)) Or (INTAG$ = TAG$(126)) Then
        SUBSW = 14: LSet FF3SPOUSE.STATUS2 = "Divorced"
        FF3SPOUSE.D2Y = 9999: GoTo 1210
     End If
     If SUBSW = 14 Then
        If (INTAG$ = TAG$(71)) Then
           LSet FF3SPOUSE.STATUS2 = XVALUE$: GoTo 1210
        End If
        If (INTAG$ = TAG$(75)) Then
           FF3SPOUSE.D1Y = 9999
           Call ConvertDATE(XVALUE$, FF3SPOUSE.D1Y, FF3SPOUSE.D2MD, ST$)
           GoTo 1210
        End If
        If (INTAG$ = TAG$(78)) Then
           DVPL = Len(RTrim$(XVALUE$))
           If (DVPL = 0) Or ((INFOSW And 2) = 0) Then GoTo 1289
           LSet DVLOC$ = XVALUE$
           If (DVPL < (23 - 19 * ENV.LPNSupport)) Then GoTo 1210
           If PLACES Then GoSub 1285: GoTo 1210
           GoTo 1289
        End If
     End If

     ' Next Marriage & Child tags Already Processed
     If (INTAG$ = TAG$(130)) Or (INTAG$ = TAG$(131)) Then GoTo 1210
     If (INTAG$ = TAG$(132)) Then SUBSW = 99: GoTo 1210
     If (INTAG$ = TAG$(133)) Or (INTAG$ = TAG$(134)) Then GoTo 1210

     ' Address Records
     If INTAG$ = TAG$(100) Then
        Call FF3PutRec
        SUBSW = 15: GoSub 1600
        Call FF3GetRec(MRN)
        If FF3SPOUSE.LOC = 0 Then FF3SPOUSE.LOC = LAREC
        GoTo 1215
     End If

1280 ' Family Record Comments
     If INTAG$ <> TAG$(135) Then GoTo 1289
     SUBSW = 0: If ((INFOSW And 2 ^ 8) = 0) Then GoTo 1289
     GoSub 1285: GoTo 1215

1285 ' Process Comment Line or Long Place Name
     Call FF3PutRec: CBGN = FF3SPOUSE.COM: GoSub 1750: LCOM = BPTR
     Call FF3GetRec(MRN)
     If CBGN <> 0 Then FF3SPOUSE.COM = CBGN: Call FF3PutRec
     Return

1289 ' Bad Family Record Line
     RT = 99: Call CountRECS: If CA$ = Chr$(27) Then GoTo 1995
     GoTo 1210

1290 ' Write Family Record & Create Place Record
     Call FF3PutRec: MLOC = 0
     If ((INFOSW And 2) <> 0) Then
        If (MPL > 0) Or (DVPL > 0) Then
           If ((MPL > 22) Or (DVPL > 22)) And ENV.LPNSupport Then
              FF3PFX.RTYPE = Chr$(12)
              If DVPL > 0 Then
                 Call FF3GetRec(0): If Not OKAY Then GoTo 1399
                 MLOC = FF3PFX.RNO
                 FF3PLACE2.SRTYPE = Chr$(3): FF3PLACE2.SRNO = MRN
                 FF3PLACE2.PTYPE = Chr$(4)
                 LSet FF3PLACE2.PLACE = xMLOC$
                 Call FF3PutRec
              End If
              If MPL > 0 Then
                 Call FF3GetRec(0): If Not OKAY Then GoTo 1399
                 FF3PLACE2.PNEXT = MLOC: MLOC = FF3PFX.RNO
                 FF3PLACE2.SRTYPE = Chr$(3): FF3PLACE2.SRNO = MRN
                 FF3PLACE2.PTYPE = Chr$(3)
                 LSet FF3PLACE2.PLACE = DVLOC$
                 Call FF3PutRec
              End If
              Else
              If (MPL > 22) And (NotePLACES) Then LSet xMLOC$ = SEENOTE$
              If (DVPL > 22) And (NotePLACES) Then LSet DVLOC$ = SEENOTE$
              FF3PFX.RTYPE = Chr$(4): Call FF3GetRec(0): If Not OKAY Then GoTo 1399
              MLOC = FF3PFX.RNO
              FF3PLACE.SRTYPE = Chr$(3): FF3PLACE.SRNO = MRN
              LSet FF3PLACE.PLACE1 = xMLOC$: LSet FF3PLACE.PLACE2 = DVLOC$
              Call FF3PutRec
           End If
           Call FF3GetRec(MRN): FF3SPOUSE.MLOC = MLOC
           Call FF3PutRec
           If MPL > 0 Then RT = 2: Call CountRECS
           If DVPL > 0 Then RT = 2: Call CountRECS
        End If
     End If
     GL = 0: GoTo 1100

1300 ' Process Individual Information Record
     If (XREF > MAXRID) Or (FF1Hdr.MAXID > 32765) Then GoTo 1190
     Call FF1GetNew: If Not OKAY Then GoTo 1190
     RT = 1: FIREF = 0: IIREF = XREF: Call CountRECS
     F1RID = FF1Rec.RID: II = IIRNO(XREF)
     GL = GL + 1: BPL = 0: DPL = 0: LFAM = 0: LCOM = 0
     LADRS = 0: LEVNT = 0: LMIL = 0: LEDU = 0: LMED = 0: LOCC = 0
     LSet xBLOC$ = " ": LSet xDLOC$ = " "
     SRNO = FF1Rec.RID: SRTYPE = 1: SUBSW = 0

1310 ' Process INDI Record Lines
     Call GetLINE
1315 If (INGL = 0) Then GoTo 1400   ' Write Name Record to File
     If INGL > (GL - (SUBSW > 0)) Then GoTo 1399
     If INGL < (GL - (SUBSW > 0)) Then SUBSW = 0: GL = INGL
     If INTAG$ = "CHEC" Then GoTo 1310
     If (INTAG$ = TAG$(31)) And (INGL = 1) Then GoTo 1310
     If INTAG$ = TAG$(32) Then
        Call ConvertNAME(XVALUE$, FF1Rec.SURNM, FF1Rec.GIVEN): GoTo 1310
     End If
     If INTAG$ = TAG$(33) Then
        FF1Rec.STATUS2 = Chr$((Asc(FF1Rec.STATUS2) Or 128)): GoTo 1310
     End If
     If INTAG$ = TAG$(34) Then FF1Rec.SEX = Left$(XVALUE$, 1): GoTo 1310

     If FHSv1 Then
        ' Birth Date, Time, Place
        If INTAG$ = TAG$(36) Then
           FF1Rec.BY = 0
           Call ConvertDATE(XVALUE$, FF1Rec.BY, FF1Rec.BMD, ST$)
           FF1Rec.STATUS = Chr$(Asc(FF1Rec.STATUS) Or (-16 * (ST$ = "?")) Or (-32 * (ST$ = "!")))
           GoTo 1310
        End If
        If INTAG$ = TAG$(38) Then
           Call ConvertTIME(XVALUE$, FF1Rec.BTM)
           GoTo 1310
        End If
        If INTAG$ = TAG$(39) Then
           BPL = Len(RTrim$(XVALUE$))
           If (BPL = 0) Or ((INFOSW And 2) = 0) Then GoTo 1399
           LSet xBLOC$ = XVALUE$
           GoTo 1310
        End If
        ' Death Date, Time, Place
        If INTAG$ = TAG$(41) Then
           FF1Rec.DY = 9999
           Call ConvertDATE(XVALUE$, FF1Rec.DY, FF1Rec.DMD, ST$)
           FF1Rec.STATUS = Chr$(Asc(FF1Rec.STATUS) Or (-64 * (ST$ = "?")) Or (-128 * (ST$ = "!")))
           GoTo 1310
        End If
        If INTAG$ = TAG$(43) Then
           Call ConvertTIME(XVALUE$, FF1Rec.DTM)
           GoTo 1310
        End If
        If INTAG$ = TAG$(45) Then
           DPL = Len(RTrim$(XVALUE$))
           If (DPL = 0) Or ((INFOSW And 2) = 0) Then GoTo 1399
           LSet xDLOC$ = XVALUE$
           GoTo 1310
        End If
     End If
        
     ' Process BIRTh Date, Time, Place
     If INTAG$ = TAG$(35) Then SUBSW = 1: GoTo 1310
     If SUBSW = 1 Then
        If INTAG$ = TAG$(75) Then
           FF1Rec.BY = 0
           Call ConvertDATE(XVALUE$, FF1Rec.BY, FF1Rec.BMD, ST$)
           FF1Rec.STATUS = Chr$(Asc(FF1Rec.STATUS) Or (-16 * ((ST$ = "A") Or (ST$ = "B") Or (ST$ = "F"))))
           GoTo 1310
        End If
        If INTAG$ = TAG$(76) Then
           ST = Val(RTrim$(XVALUE$))
           FF1Rec.STATUS = Chr$(Asc(FF1Rec.STATUS) Or (-32 * (ST = 3)) Or (-16 * (ST = 0)))
           GoTo 1310
        End If
        If INTAG$ = TAG$(77) Then
           Call ConvertTIME(XVALUE$, FF1Rec.BTM)
           GoTo 1310
        End If
        If INTAG$ = TAG$(78) Then
           BPL = Len(RTrim$(XVALUE$))
           If (BPL = 0) Or ((INFOSW And 2) = 0) Then GoTo 1399
           LSet xBLOC$ = XVALUE$
           If (BPL < (23 - 19 * ENV.LPNSupport)) Then GoTo 1310
           If NotePLACES Then
              CBGN = FF1Rec.COM: GoSub 1750: FF1Rec.COM = CBGN
              GoTo 1310
           End If
           GoTo 1399
        End If
     End If

     ' Process DEATh  Date, Time, Place TAGs
     If INTAG$ = TAG$(40) Then SUBSW = 2: FF1Rec.DY = 9999: GoTo 1310
     If SUBSW = 2 Then
        If INTAG$ = TAG$(75) Then
           FF1Rec.DY = 9999
           Call ConvertDATE(XVALUE$, FF1Rec.DY, FF1Rec.DMD, ST$)
           FF1Rec.STATUS = Chr$(Asc(FF1Rec.STATUS) Or (-64 * ((ST$ = "A") Or (ST$ = "B") Or (ST$ = "F"))))
           GoTo 1310
        End If
        If INTAG$ = TAG$(76) Then
           ST = Val(RTrim$(XVALUE$))
           FF1Rec.STATUS = Chr$(Asc(FF1Rec.STATUS) Or (-128 * (ST = 3)) Or (-64 * (ST = 0)))
           GoTo 1310
        End If
        If INTAG$ = TAG$(77) Then
           Call ConvertTIME(XVALUE$, FF1Rec.DTM)
           GoTo 1310
        End If
        If INTAG$ = TAG$(78) Then
           DPL = Len(RTrim$(XVALUE$))
           If (DPL = 0) Or ((INFOSW And 2) = 0) Then GoTo 1399
           LSet xDLOC$ = XVALUE$
           If (DPL < (23 - 19 * ENV.LPNSupport)) Then GoTo 1310
           If NotePLACES Then
              CBGN = FF1Rec.COM: GoSub 1750: FF1Rec.COM = CBGN
              GoTo 1310
           End If
           GoTo 1399
        End If
     End If

     ' Marriage XREF
     If INTAG$ = TAG$(50) Then
        If PAFV20 Then SUBSW = 3: GoTo 1310
        If (PAFV21 Or LDSv55) Then
           Get #6, II, PtrREC1
           X = FIRNO(XID)
           If X <> 0 Then
              Get #7, X, PtrREC2
              If (PtrREC2.WFHID = PtrREC1.WFID) And ((PtrREC2.WFSP And 1) = 0) Then
                 If PtrREC2.WFHOF = 0 Then
                    PtrREC2.WFHOF = PtrREC1.WFMREC
                    PtrREC1.WFMREC = XID
                    PtrREC2.WFSP = PtrREC2.WFSP Or 1
                 End If
                 Else
                 If (PtrREC2.WFWID = PtrREC1.WFID) And ((PtrREC2.WFSP And 2) = 0) Then
                    If PtrREC2.WFWOF = 0 Then
                       PtrREC2.WFWOF = PtrREC1.WFMREC
                       PtrREC1.WFMREC = XID
                       PtrREC2.WFSP = PtrREC2.WFSP Or 2
                    End If
                 End If
              End If
              Put #7, X, PtrREC2
           End If
           Put #6, II, PtrREC1
           GoTo 1310
        End If
        GoTo 1399
     End If
     If (SUBSW = 3) And (INTAG$ = TAG$(31)) Then GoTo 1310

     ' Bypass Parent's Marriage Line(s)
     If INTAG$ = TAG$(54) Then
        If (PAFV20 Or PAFV21 Or LDSv55) Then SUBSW = 4: GoTo 1310
     End If
     If (SUBSW = 4) And (INTAG$ = TAG$(31)) Then GoTo 1310

     ' Check for Christening, Baptismal or Burial information
     If Left$(INTAG$, 3) = "BAP" Then SUBSW = 8: ECODE$ = "BAPTISM"
     If Left$(INTAG$, 3) = "CHR" Then SUBSW = 8: ECODE$ = "CHRISTEN"
     If Left$(INTAG$, 3) = "BUR" Then SUBSW = 8: ECODE$ = "BURIAL"
     If Left$(INTAG$, 3) = "ENG" Then SUBSW = 8: ECODE$ = "ENGAGED"
     If SUBSW = 8 Then
        If ((INFOSW And 2 ^ 7) <> 0) Then
           FF3PFX.RTYPE = Chr$(6): Call FF3GetRec(0)
           Else
           OKAY = 0
        End If
        If Not OKAY Then GoTo 1399
        LSet FF3EVENT.SRTYPE = Chr$(SRTYPE): FF3EVENT.SRNO = SRNO
        LSet FF3EVENT.CODE = ECODE$
        RT = 8: GoSub 1555: GoTo 1315
     End If

     ' Adoption TAGs
     If INTAG$ = TAG$(55) Then GoTo 1310
     If PAFV20 Then
        If (INTAG$ = TAG$(56)) Or (INTAG$ = TAG$(57)) Then GoTo 1310
     End If
       

     ' Bypass PAFV20,FHSV1,FHSV2 Sibling Tags
     If PAFV20 Then
        If (INTAG$ = TAG$(58)) Then SUBSW = 99: GoTo 1310
        If (INTAG$ = TAG$(59)) Then GoTo 1310
     End If
     If FHSv1 Or FHSv2 Then
        If (INTAG$ = TAG$(60)) Or (INTAG$ = TAG$(62)) Or (INTAG$ = TAG$(64)) Then GoTo 1310
        If (INTAG$ = TAG$(51)) Or (INTAG$ = TAG$(61)) Or (INTAG$ = TAG$(63)) Then GoTo 1310
     End If

     ' Miscellaneous Records
     If (INTAG$ = TAG$(86)) And ((INFOSW And 2 ^ 3) <> 0) Then RT = 4
     If (INTAG$ = TAG$(90)) And ((INFOSW And 2 ^ 4) <> 0) Then RT = 5
     If (INTAG$ = TAG$(92)) And ((INFOSW And 2 ^ 5) <> 0) Then RT = 6
     If (INTAG$ = TAG$(95)) And ((INFOSW And 2 ^ 6) <> 0) Then RT = 7
     If (RT > 3) And (RT < 8) Then GoSub 1500: GoTo 1315

     ' Event Records
     If (INTAG$ = TAG$(70)) And ((INFOSW And 2 ^ 7) <> 0) Then
        RT = 8: GoSub 1550: GoTo 1315
     End If

     ' Address Records
     If INTAG$ = TAG$(100) Then
        SUBSW = 15: LAREC = LADRS: GoSub 1600: LADRS = LAREC
        If FF1Rec.ADRS = 0 Then FF1Rec.ADRS = LADRS
        GoTo 1315
     End If

1380 ' Process Name Record Comments
     If INTAG$ <> TAG$(79) Then GoTo 1399
     SUBSW = 0: If ((INFOSW And 2 ^ 8) = 0) Then GoTo 1399
     CBGN = FF1Rec.COM: GoSub 1750: FF1Rec.COM = CBGN
     GoTo 1315

1399 ' Error Tag in Name Record
     RT = 99: Call CountRECS: If CA$ = Chr$(27) Then GoTo 1995
     GoTo 1310

1400 ' Write Name Record to File & Create Place Name Record
     If (SUBSW = 6) Then Call FF3PutRec: SUBSW = 0: OKAY = -1
     If ((INFOSW And 2) <> 0) Then
        BLOC = 0
        If (BPL > 0) Or (DPL > 0) Then
           If ((BPL > 22) Or (DPL > 22)) And ENV.LPNSupport Then
              FF3PFX.RTYPE = Chr$(12)
              If DPL > 0 Then
                 Call FF3GetRec(0)
                 If OKAY Then
                    BLOC = FF3PFX.RNO
                    FF3PLACE2.SRTYPE = Chr$(1): FF3PLACE2.SRNO = F1RID
                    FF3PLACE2.PTYPE = Chr$(2)
                    LSet FF3PLACE2.PLACE = xDLOC$
                    Call FF3PutRec
                 End If
              End If
              If BPL > 0 Then
                 Call FF3GetRec(0)
                 If OKAY Then
                    FF3PLACE2.PNEXT = BLOC: BLOC = FF3PFX.RNO
                    FF3PLACE2.SRTYPE = Chr$(1): FF3PLACE2.SRNO = F1RID
                    FF3PLACE2.PTYPE = Chr$(1)
                    LSet FF3PLACE2.PLACE = xBLOC$
                    Call FF3PutRec
                 End If
              End If
              Else
              If (BPL > 22) And (NotePLACES) Then LSet xxBLOC$ = SEENOTE$
              If (DPL > 22) And (NotePLACES) Then LSet xDLOC$ = SEENOTE$
              FF3PFX.RTYPE = Chr$(4): Call FF3GetRec(0)
              If OKAY Then
                 BLOC = FF3PFX.RNO
                 FF3PLACE.SRTYPE = Chr$(1): FF3PLACE.SRNO = F1RID
                 LSet FF3PLACE.PLACE1 = xxBLOC$: LSet FF3PLACE.PLACE2 = xDLOC$
                 Call FF3PutRec
              End If
           End If
           If OKAY Then
              If BPL > 0 Then RT = 2: Call CountRECS
              If DPL > 0 Then RT = 2: Call CountRECS
           End If
        End If
     End If
     If Not OKAY Then RT = 99: Call CountRECS: BLOC = 0
     FF1Rec.BLOC = BLOC
     Call FF1PutREC
     GL = 0: GoTo 1100

1500 ' Process Miscellaneous Record Types
     FF3PFX.RTYPE = Chr$(RT + 3): Call FF3GetRec(0)
     If Not OKAY Then
        RT = 99: Call CountRECS: Call GetLINE: GoTo 1549
     End If
     Call CountRECS: xGL = INGL + 1: XSUB = 0: XRT = RT
     FF3PNTRS.SRTYPE = Chr$(SRTYPE): FF3PNTRS.SRNO = SRNO
     SRTYPE = Asc(FF3PFX.RTYPE): SRNO = FF3PFX.RNO
1510 Call GetLINE
1511 If INGL < xGL Then GoTo 1520
     If INTAG$ = TAG$(75) Then
        X = InStr(1, XVALUE$, "from")
        If X > 0 Then
           XDATE$ = MID$(XVALUE$, X + 5, 11)
           Call ConvertDATE(XDATE$, FF3PNTRS.D1Y, FF3PNTRS.D1MD, ST$)
        End If
        X = InStr(1, XVALUE$, "until")
        If X > 0 Then
           XDATE$ = MID$(XVALUE$, X + 6, 11)
           Call ConvertDATE(XDATE$, FF3PNTRS.D2Y, FF3PNTRS.D2MD, ST$)
        End If
        GoTo 1510
     End If
     If INTAG$ = TAG$(81) Then
        Call ConvertDATE(XVALUE$, FF3PNTRS.D1Y, FF3PNTRS.D1MD, ST$): GoTo 1510
     End If
     If INTAG$ = TAG$(82) Then
        Call ConvertDATE(XVALUE$, FF3PNTRS.D2Y, FF3PNTRS.D2MD, ST$): GoTo 1510
     End If
     If XRT = 4 Then
        If INTAG$ = TAG$(87) Then
           XSUB = XSUB + 1
           Select Case XSUB
              Case 1: LSet FF3EDUC.EDSUB1 = XVALUE$
              Case 2: LSet FF3EDUC.EDSUB2 = XVALUE$
              Case Else: GoTo 1515
           End Select
           GoTo 1510
        End If
        If INTAG$ = TAG$(88) Then LSet FF3EDUC.EDLEV = XVALUE$: GoTo 1510
        If INTAG$ = TAG$(89) Then LSet FF3EDUC.EDDEG = XVALUE$: GoTo 1510
     End If
     If XRT = 5 Then
        If INTAG$ = TAG$(91) Then LSet FF3WORK.WORKTP = XVALUE$: GoTo 1510
     End If
     If XRT = 6 Then
        If INTAG$ = TAG$(93) Then LSet FF3MIL.MILRANK = XVALUE$: GoTo 1510
        If INTAG$ = TAG$(83) Then LSet FF3MIL.MILSTAT = XVALUE$: GoTo 1510
     End If
     If XRT = 7 Then
        If INTAG$ = TAG$(96) Then LSet FF3HEALTH.HLTHDIAG = XVALUE$: GoTo 1510
        If INTAG$ = TAG$(83) Then LSet FF3HEALTH.HLTHSTAT = XVALUE$: GoTo 1510
     End If

     ' Address Records
     If INTAG$ = TAG$(100) Then
        Call FF3PutRec: LAREC = FF3PNTRS.LOC
        SUBSW = 15: GoSub 1600
        Call FF3GetRec(SRNO)
        If FF3PNTRS.LOC = 0 Then FF3PNTRS.LOC = LAREC
        GoTo 1511
     End If

     'Comments
     If (INTAG$ = TAG$(79)) And ((INFOSW And 2 ^ 8) <> 0) Then
        Call FF3PutRec
        CBGN = FF3PNTRS.COM: GoSub 1750
        Call FF3GetRec(SRNO): FF3PNTRS.COM = CBGN
        GoTo 1511
     End If

1515 RT = 99: Call CountRECS: GoTo 1510

1520 ' Write Miscellaneous Record
     Call FF3PutRec
     SRTYPE = Asc(FF3PNTRS.SRTYPE): SRNO = FF3PNTRS.SRNO
     Select Case XRT
        Case 4: If FF1Rec.EDUC = 0 Then FF1Rec.EDUC = FF3PFX.RNO
                X = LEDU: LEDU = FF3PFX.RNO
        Case 5: If FF1Rec.WORK = 0 Then FF1Rec.WORK = FF3PFX.RNO
                X = LOCC: LOCC = FF3PFX.RNO
        Case 6: If FF1Rec.MIL = 0 Then FF1Rec.MIL = FF3PFX.RNO
                X = LMIL: LMIL = FF3PFX.RNO
        Case 7: If FF1Rec.HLTH = 0 Then FF1Rec.HLTH = FF3PFX.RNO
                X = LMED: LMED = FF3PFX.RNO
     End Select
     If X <> 0 Then
        Y = FF3PFX.RNO
        Call FF3GetRec(X): FF3PNTRS.NXT = Y: Call FF3PutRec
     End If
1549 ' Return to Main Procedure
     Return

1550 ' Process Event Record
     FF3PFX.RTYPE = Chr$(6): Call FF3GetRec(0)
     If Not OKAY Then
        RT = 99: Call CountRECS: Call GetLINE: GoTo 1599
     End If
1555 Call CountRECS: xGL = INGL + 1
     FF3EVENT.SRTYPE = Chr$(SRTYPE): FF3EVENT.SRNO = SRNO
     SRTYPE = Asc(FF3PFX.RTYPE): SRNO = FF3PFX.RNO
1560 Call GetLINE
1561 If INGL < xGL Then GoTo 1570
     If INTAG$ = TAG$(71) Then LSet FF3EVENT.CODE = XVALUE$: GoTo 1560
     If INTAG$ = TAG$(72) Then
        X = Val(XVALUE$)
        If (X < 0) Then X = 0 Else If (X > 255) Then X = 255
        LSet FF3EVENT.IMPORT = Chr$(X): GoTo 1560
     End If
     If INTAG$ = TAG$(75) Then
        Call ConvertDATE(XVALUE$, FF3EVENT.D1Y, FF3EVENT.D1MD, FF3EVENT.DSTAT)
        GoTo 1560
     End If
     If INTAG$ = TAG$(76) Then LSet FF3EVENT.DSTAT = XVALUE$: GoTo 1560
     If INTAG$ = TAG$(78) Then
        EPL = Len(RTrim$(XVALUE$))
        If (EPL < (23 - 19 * ENV.LPNSupport)) Or (Not NotePLACES) Then
           LSet FF3EVENT.PLACE = XVALUE$
           If (EPL > 22) And ENV.LPNSupport Then
              Call FF3PutRec
              FF3PFX.RTYPE = Chr$(12): Call FF3GetRec(0)
              FF3PLACE2.SRTYPE = Chr$(SRTYPE): FF3PLACE2.SRNO = SRNO
              LSet FF3PLACE2.PLACE = XVALUE$
              Call FF3PutRec: LPLNM = FF3PFX.RNO
              Call FF3GetRec(SRNO)
              FF3EVENT.LPLNM = LPLNM
           End If
           GoTo 1560
        End If
        LSet FF3EVENT.PLACE = SEENOTE$
        LSet INTAG$ = TAG$(79)
     End If
     If (INTAG$ = TAG$(79)) And ((INFOSW And 2 ^ 8) <> 0) Then
        Call FF3PutRec
        CBGN = FF3EVENT.COM: GoSub 1750
        Call FF3GetRec(SRNO): FF3EVENT.COM = CBGN
        GoTo 1561
     End If
     RT = 99: Call CountRECS: GoTo 1560
1570 ' Write Event Record
     Call FF3PutRec
     If FF1Rec.EVENT = 0 Then FF1Rec.EVENT = SRNO
     SRTYPE = Asc(FF3EVENT.SRTYPE): SRNO = FF3EVENT.SRNO
     X = FF3PFX.RNO
     If LEVNT <> 0 Then
        Call FF3GetRec(LEVNT): FF3EVENT.NXT = X: Call FF3PutRec
     End If
     LEVNT = X
1599 ' Return to Main Procedure
     Return


1600 ' Process Address Record
     If ((INFOSW And 2 ^ 7) = 0) Then
        RT = 99: Call CountRECS: Call GetLINE
        Return
     End If
     RT = 9: Call CountRECS
     Call FF2GetRec(0)
     FF2PFX.SRTYPE = Chr$(SRTYPE): FF2PFX.SRNO = SRNO
     SRTYPE = Asc(FF2PFX.RTYPE): SRNO = FF2PFX.RNO
     xGL = INGL + 1
1610 Call GetLINE
1611 If INGL < xGL Then GoTo 1620
     If INTAG$ = TAG$(75) Then
        X = InStr(1, XVALUE$, "from")
        If X > 0 Then
           XDATE$ = MID$(XVALUE$, X + 5, 11)
           Call ConvertDATE(XDATE$, FF2PFX.D1Y, FF2PFX.D1MD, ST$)
        End If
        X = InStr(1, XVALUE$, "until")
        If X > 0 Then
           XDATE$ = MID$(XVALUE$, X + 6, 11)
           Call ConvertDATE(XDATE$, FF2PFX.D2Y, FF2PFX.D2MD, ST$)
        End If
        GoTo 1610
     End If
     If INTAG$ = TAG$(81) Then
        Call ConvertDATE(XVALUE$, FF2PFX.D1Y, FF2PFX.D1MD, ST$): GoTo 1610
     End If
     If INTAG$ = TAG$(82) Then
        Call ConvertDATE(XVALUE$, FF2PFX.D2Y, FF2PFX.D2MD, ST$): GoTo 1610
     End If
     If INTAG$ = TAG$(101) Then
        If AFMT = 0 Then
           LSet FF2SREC.LINE1 = XVALUE$
           Else
           LSet FF2LREC.LINE1 = XVALUE$
        End If
        GoTo 1610
     End If
     If INTAG$ = TAG$(102) Then
        If AFMT = 0 Then
           LSet FF2SREC.LINE2 = XVALUE$
           Else
           LSet FF2LREC.LINE2 = XVALUE$
        End If
        GoTo 1610
     End If
     If INTAG$ = TAG$(103) Then
        If AFMT = 0 Then
           LSet FF2SREC.CITY = XVALUE$
           Else
           LSet FF2LREC.CITY = XVALUE$
        End If
        GoTo 1610
     End If
     If INTAG$ = TAG$(104) Then
        If AFMT = 0 Then
           LSet FF2SREC.STATE = XVALUE$
           Else
           LSet FF2LREC.STATE = XVALUE$
        End If
        GoTo 1610
     End If
     If INTAG$ = TAG$(105) Then
        If AFMT = 0 Then
           LSet FF2SREC.ZIP = XVALUE$
           Else
           LSet FF2LREC.ZIP = XVALUE$
        End If
        GoTo 1610
     End If
     If INTAG$ = TAG$(106) Then
        If AFMT = 0 Then GoTo 1615
        LSet FF2LREC.COUNTRY = XVALUE$
        GoTo 1610
     End If
     If INTAG$ = TAG$(107) Then
        If AFMT = 0 Then
           Call GetSPHONE(XVALUE$)
           Else
           LSet FF2LREC.PHONE = XVALUE$
        End If
        GoTo 1610
     End If
     If (INTAG$ = TAG$(108)) And ((INFOSW And 2 ^ 8) <> 0) Then
        CBGN = FF2PFX.COM: GoSub 1750: FF2PFX.COM = CBGN
        GoTo 1611
     End If
1615 RT = 99: Call CountRECS
     GoTo 1610
1620 ' Write Address Record
     Call FF2PutRec
     If LAREC <> 0 Then
        X = FF2PFX.RNO
        Call FF2GetRec(LAREC): FF2PFX.NXT = X: Call FF2PutRec
        Call FF2GetRec(X)
     End If
     SRTYPE = Asc(FF2PFX.SRTYPE): SRNO = FF2PFX.SRNO
     LAREC = FF2PFX.RNO
     Return
     

1750 ' Process Comment Lines
     GoSub 1780
     If RTrim$(XVALUE$) = "" Then GoTo 1770
1755 ' Next Comment Line
     L = Len(RTrim$(XVALUE$))
     If L > 0 Then Mid$(CLINE$, O, L) = XVALUE$
1760 O = Len(RTrim$(CLINE$)) + 1
     If (O > 38) Then
        GoSub 1790: LSet CLINE$ = MID$(CLINE$, 39, 255)
        GoTo 1760
     End If
1770 ' Next Comment Line
     ' Skip Ahead if processing Long Place Names
     If ((SUBSW > 0) And (SUBSW < 3)) Or ((SUBSW > 12) And (SUBSW < 15)) Then GoTo 1775
     Call GetLINE
     If (INTAG$ = TAG$(79)) Or (INTAG$ = TAG$(84)) Or (INTAG$ = TAG$(108)) Or (INTAG$ = TAG$(135)) Then
        Mid$(CLINE$, O, 1) = "\": O = O + 1
        Else
        If ((INTAG$ = TAG$(80)) Or (INTAG$ = TAG$(85)) Or (INTAG$ = TAG$(109)) Or (INTAG$ = TAG$(136))) Then
           If O > 1 Then O = O + 1
           Else
           GoTo 1775
        End If
     End If
     GoTo 1755

1775 ' Final Line of Comments
     GoSub 1790: LSet CLINE$ = ""
     If (CBGN <> 0) And (CCNT > 1) Then
        Call FF3GetRec(CBGN)
        FF3COMMENT.CCNT = CCNT
        Call FF3PutRec
     End If
     Return

1780 ' Prepare to Create Comment Lines
     CCNT = 0: If CBGN <> 0 Then Call FF3GetRec(CBGN): CCNT = FF3COMMENT.CCNT
     FPTR = CBGN: BPTR = 0: O = 1
     While FPTR <> 0
        BPTR = FPTR: Call FF3GetRec(FPTR): FPTR = FF3COMMENT.FPTR
     Wend
     If INTAG$ = "PLAC" Then
        If SUBSW = 1 Then Mid$(CLINE$, O, 20) = BORN$: O = O + Len(BORN$)
        If SUBSW = 2 Then Mid$(CLINE$, O, 20) = DIED$: O = O + Len(DIED$)
        If SUBSW = 13 Then Mid$(CLINE$, O, 20) = MARRIED$ + ":": O = O + Len(MARRIED$) + 1
     End If
     Return

1790 ' Create Comment Record
     If (Left$(CLINE$, 38) <> Space$(38)) Then
        FF3PFX.RTYPE = Chr$(5): Call FF3GetRec(0)
        If OKAY Then
           RT = 10: Call CountRECS
           LSet FF3COMMENT.TEXT = CLINE$
           FF3COMMENT.SRTYPE = Chr$(SRTYPE): FF3COMMENT.SRNO = SRNO
           CCNT = CCNT + 1: FF3COMMENT.CCNT = CCNT
           If CBGN = 0 Then CBGN = FF3PFX.RNO
           FF3COMMENT.BPTR = BPTR: FF3COMMENT.FPTR = 0
           Call FF3PutRec
           If BPTR <> 0 Then
              FPTR = FF3PFX.RNO
              Call FF3GetRec(BPTR): FF3COMMENT.FPTR = FPTR
              Call FF3PutRec: BPTR = FPTR
              Else
              BPTR = FF3PFX.RNO
           End If
        End If
     End If
     Return

1900 ' Final Steps of Import Operation
     Call GEDCOMStats
     If (RECCNT(1) = 0) Then GoSub 1995
     If INTAG$ <> TAG$(140) Then
        Call PutMSG("No Trailer Record...Continue? (Y/N)")
        Call GetRESP: Call ClearMSG
     End If

1950 ' Update Relationship & Family Record Pointers
     On Local Error GoTo 1996
     If PAFV20 Then   ' Reverse sequence of marriage records
        Call PutMSG("GDCMM032")
        For I = 1 To HIID
            Get #6, I, PtrREC1
            X = PtrREC1.WFMREC: Y = FIRNO(X): Z = 0
            While Y <> 0
               Get #7, Y, PtrREC2
               If PtrREC1.WFID = PtrREC2.WFHID Then
                  'SWAP Z, PtrREC2.WFHOF
                  hz = Z: Z = PtrREC2.WFHOF: PtrREC2.WFHOF = hz
                  Else
                  If PtrREC1.WFID = PtrREC2.WFWID Then
                     'SWAP Z, PtrREC2.WFWOF
                     hz = Z: Z = PtrREC2.WFWOF: PtrREC2.WFWOF = hz
                     Else
                     Z = 0
                  End If
               End If
               Put #7, Y, PtrREC2
               Y = FIRNO(Z): 'SWAP Z, x
               hz = Z: Z = X: X = hz
            Wend
            PtrREC1.WFMREC = Z
            Put #6, I, PtrREC1
            CA$ = INKEY$
            If CA$ <> "" Then
               If CA$ = Chr$(27) Then
                  I = HIID
                  Else
                  SysVAR.NVAR1 = I: SysVAR.NVAR2 = HIID
                  Call PutMSG("GDCMM027")
               End If
            End If
        Next I
     End If
     If CA$ = Chr$(27) Then GoTo 1995
     If Not (FHSv1 Or FHSv2) Then   ' Reverse sequence of children
        Call PutMSG("GDCMM033")
        For I = 1 To HIID
            Get #6, I, PtrREC1: id = PtrREC1.WFID
            X = PtrREC1.WFMREC: Y = FIRNO(X)
            Z0 = PtrREC1.WFCH: WFP = PtrREC1.WFP
            While Y <> 0
               Get #7, Y, PtrREC2
               Z1 = PtrREC2.WF1CH: Z2 = IIRNO(Z1)
               While Z2 <> 0
                  Get #6, Z2, PtrREC1
                  If (id = PtrREC2.WFHID) And ((PtrREC1.WFP And 1) = 0) Then
                     PtrREC1.WFFID = id
                     hz = Z0: Z0 = Z1
                     Z1 = PtrREC1.WFFSIB
                     PtrREC1.WFFSIB = hz
                     PtrREC1.WFP = PtrREC1.WFP Or 1
                     Else
                     If (id = PtrREC2.WFWID) And ((PtrREC1.WFP And 2) = 0) Then
                        PtrREC1.WFMID = id
                        hz = Z0: Z0 = Z1
                        Z1 = PtrREC1.WFMSIB
                        PtrREC1.WFMSIB = hz
                        PtrREC1.WFP = PtrREC1.WFP Or 2
                        Else
                        Z1 = 0
                     End If
                  End If
                  Put #6, Z2, PtrREC1
                  Z2 = IIRNO(Z1)
                  'If (I = 0) Then
                  '   Call PutMSG("Waiting at I=" + Str$(I) + " Z2=" + Str$(Z2) + "...")
                  '   Call GetRESP
                  '   If CA$ = Chr$(27) Then I = HIID: Z2 = 0: Y = 0
                  'End If
               Wend
               If (id = PtrREC2.WFHID) And ((WFP And 4) = 0) Then
                  X = PtrREC2.WFHOF
                  WFP = WFP Or 4
                  Else
                  If (id = PtrREC2.WFWID) And ((WFP And 8) = 0) Then
                     X = PtrREC2.WFWOF
                     WFP = WFP Or 8
                     Else
                     X = 0
                  End If
               End If
               Y = FIRNO(X)
            Wend
            If Z0 <> 0 Then
               Get #6, I, PtrREC1
               PtrREC1.WFCH = Z0
               PtrREC1.WFP = WFP
               Put #6, I, PtrREC1
            End If
            CA$ = INKEY$
            If CA$ <> "" Then
               If CA$ = Chr$(27) Then
                  I = HIID
                  Else
                  K$ = CA$
                  SysVAR.NVAR1 = I: SysVAR.NVAR2 = HIID
                  Call PutMSG("GDCMM027")
               End If
            End If
        Next I
     End If
     If CA$ = Chr$(27) Then GoTo 1995
     ' Update Pointers in Family File records
     Call PutMSG("GDCMM034")
     For I = 1 To HIID
         Get #6, I, PtrREC1
         id = F1MAXID + I: Call FF1GetRec(id)
         X = IIRNO(PtrREC1.WFFID)
         If X <> 0 Then
            FF1Rec.FID = X + F1MAXID
            X = IIRNO(PtrREC1.WFFSIB)
            If X <> 0 Then FF1Rec.FCH = X + F1MAXID
         End If
         X = IIRNO(PtrREC1.WFMID)
         If X <> 0 Then
            FF1Rec.MID = X + F1MAXID
            X = IIRNO(PtrREC1.WFMSIB)
            If X <> 0 Then FF1Rec.MCH = X + F1MAXID
         End If
         X = IIRNO(PtrREC1.WFCH)
         If X <> 0 Then FF1Rec.OLDCH = X + F1MAXID
         X = MISCSEQ(FIRNO(PtrREC1.WFMREC))
         If X <> 0 Then FF1Rec.SPOUSE = X
         FF1Rec.STATUS = Chr$(Asc(FF1Rec.STATUS) Or Asc(PtrREC1.WFS1))
         Call FF1PutREC
         CA$ = INKEY$
         If CA$ <> "" Then
            If CA$ = Chr$(27) Then
               I = HIID
               Else
               K$ = CA$
               SysVAR.NVAR1 = I: SysVAR.NVAR2 = HIID
               Call PutMSG("GDCMM027")
            End If
         End If
     Next I
     If CA$ = Chr$(27) Then GoTo 1995
     Call PutMSG("GDCMM035")
     For I = 1 To HIMRN
         Get #7, I, PtrREC2
         MRN = PtrREC2.WFMRN
         If MRN <> 0 Then
            Call FF3GetRec(MRN)
            X = IIRNO(PtrREC2.WFHID)
            If X <> 0 Then
               FF3SPOUSE.SP1ID = X + F1MAXID
               X = MISCSEQ(FIRNO(PtrREC2.WFHOF))
               If X <> 0 Then FF3SPOUSE.SP1NXT = X
            End If
            X = IIRNO(PtrREC2.WFWID)
            If X <> 0 Then
               FF3SPOUSE.SP2ID = X + F1MAXID
               X = MISCSEQ(FIRNO(PtrREC2.WFWOF))
               If X <> 0 Then FF3SPOUSE.SP2NXT = X
            End If
            Call FF3PutRec
            CA$ = INKEY$
            If CA$ <> "" Then
               If CA$ = Chr$(27) Then
                  I = HIMRN
                  Else
                  SysVAR.NVAR1 = I: SysVAR.NVAR2 = HIMRN
                  Call PutMSG("GDCMM027")
                  Call GetRESP: Call ClearMSG
               End If
            End If
         End If
     Next I

1995 ' Return
     If (PrintERR And (RECCNT(11) > 0)) Then Call RptCLOSE
     Call FamCLOSE
     If RECCNT(1) > 0 Then ReDim Preserve SelSET(FF1Hdr.MAXID)
     On Local Error GoTo 0
     For X = 1 To 14: RMONTB$(X) = XMONTB$(X): Next X
     If XGEDLine& > GEDLine& Then
        GEDLine& = XGEDLine&: Call GEDCOMStats
     End If
     Call PutMSG("GDCMM036")
     Call ClearMSG
     Call frmFHSWUPgm.ShowStats("GEDCOM", 99, "")
     Exit Sub

1996 ' Error Routines
     XERR = ERR: XERL = ERL: Call GEDCOMStats
     Call ERRMessage("GDCMM037", XERR, XERL)
     Resume 1995
1999 XERR = ERR: XERL = ERL: Call GEDCOMStats
     Call ERRMessage("GDCMM038", XERR, XERL)
     Resume 1995

End Sub

Rem $STATIC
Sub GEDCOMLine(GDTAG, xXREF$, VALUE$)

5000 ' Write Line to GEDCOM File
   If GDCMRULE = 4 Then xX$ = "" Else xX$ = " "
   If (GDCMRULE < 3) And (xXREF$ <> "") Then Y$ = " " Else Y$ = ""
   If CNVTCASE Then
      If ((GDTAG > 38) And ((GDTAG < 46) Or (GDTAG > 70)) And ((GDTAG < 110) Or (GDTAG > 120)) And (GDTAG < 136)) Then
         Select Case GDTAG
            Case 39: X = 1
            Case 45: X = 1
            Case 71: X = 1
            Case 78: X = 1
            Case 79: X = 2
            Case 80: X = 3
            Case 83: X = 1
            Case 84: X = 2
            Case 85: X = 3
            Case 87: X = 1
            Case 88: X = 1
            Case 89: X = 1
            Case 91: X = 1
            Case 93: X = 1
            Case 94: X = 1
            Case 96: X = 1
            Case 97: X = 1
            Case 101: X = 1
            Case 102: X = 1
            Case 103: X = 1
            Case 104: X = 1
            Case 105: X = 1
            Case 106: X = 1
            Case 108: X = 2
            Case 109: X = 3
            Case 121: X = 1
            Case 124: X = 1
            Case 129: X = 1
            Case 135: X = 2
            Case 136: X = 3
         End Select
         If X > 0 Then
            If X = 1 Then
               Call ConvertCASE("NAME", VALUE$, S)
               Else
               If X = 3 Then S = 0
               Call ConvertCASE("TEXT", VALUE$, S)
            End If
         End If
      End If
   End If
   On Local Error GoTo 5090
5010 ' Write Line
   Print #5, fnx$(GL); xX$; RTrim$(xXREF$); Y$; TAG$(GDTAG); " "; RTrim$(VALUE$)
   If GDTAG = 30 Then
      INDI = INDI + 1
      Else
      If GDTAG = 110 Then FAMI = FAMI + 1
   End If
   GEDLine& = GEDLine& + 1: OKAY = -1
   GoTo 5099

5090 ' Error Writing to GEDCOM File
     XERR = ERR: DISKEOF = -1
     Call ERRMessage("GDCMM042", ERR, ERL)
     FDISK = (MID$(GEDFile$, 2, 1) = ":")
     If (XERR <> 61) Or (Not FDISK) Then Resume 5099
     Call GEDCOMStats
     Call GEDCOMOpen("WRITENQ")
     If Not OKAY Then Resume 5099
     Resume 5010

5099 ' Exit

End Sub

Rem $DYNAMIC
Sub GEDCOMList()

2000 ' Print Exported Data File Listing
     TAG$(0) = ""
     Call RptOPEN("GDLS"): If (Not OKAY) Then GoTo 2069
     Call GEDCOMOpen("READ"): If (Not OKAY) Or (a = 27) Then Exit Sub
     Call PutMSG("GDCMM040")
     If CA$ = Chr$(27) Then GoTo 2069
     OUTLINE = (CA$ <> SNGLKEY$(2))
     Call PrintOPEN: If Not OKAY Then GoTo 2069
     C$ = " "
2060 If Not DISKEOF Then
        Call GetLINE
        If OUTLINE Then
           LSet C$ = TLINE$: O = 2 * Val(C$)
           Mid$(PLINE$, O + 1, 255) = TLINE$
           Else
           LSet PLINE$ = TLINE$
        End If
        Call PrintLINE: If a <> 27 Then GoTo 2060
     End If
     Call RptCLOSE
     If PRT.AGAIN Then
        Call GEDCOMOpen("READQ"): If OKAY Then GoTo 2060
     End If
2069 Close
     Exit Sub
End Sub

Sub GEDCOMOpen(OPT$)

18000 ' Get GEDFile Prefix
      GED$ = GEDFile$: X = InStr(GED$, ".")
      If X > 0 Then GED$ = Left$(GED$, X - 1)
      ' Make Sure File is Closed
      On Local Error GoTo 18001
      Close #5
      GoTo 18002
18001 Resume 18002
18002 OKAY = 0: XOPT$ = OPT$: L = Len(XOPT$)
      QUIET = (MID$(XOPT$, L, 1) = "Q")
      If QUIET Then L = L - 1: XOPT$ = Left$(XOPT$, L)
      NxtDISK = (MID$(XOPT$, L, 1) = "N")
      If NxtDISK Then
         XOPT$ = Left$(XOPT$, L - 1)
         NDISK = DiskNUM + 1
         GED$ = GED$ + ".G" + ZFILL$(DiskNUM, 2)
         Else
         NDISK = 0: GED$ = GED$ + ".GED"
      End If
      If Not QUIET Then
         Call frmFHSWFile.LOCATE("GEDCOM", "FULL")
         If a = 27 Then OKAY = 0: GoTo 18300
         GED$ = TRANPFX$ + FT$(11)
      End If
      DSNM$ = GED$
      FDISK = (MID$(DSNM$, 2, 1) = ":")
      If FDISK Then
         FDRV$ = FDFSetup.TRANDRV
         FDISK = ((FDRV$ = "A") Or (FDRV$ = "B"))
      End If
18010 ' Check for More Input Diskettes
      If FDISK And NxtDISK And (XOPT$ = "READ") Then
         Call PutMSG("GDCMM044")
         OKAY = (CA$ = SNGLKEY$(1))
         If Not OKAY Then
            If (CA$ = Chr$(27)) Or (CA$ = SNGLKEY$(2)) Then GoTo 18300
            GoTo 18010
         End If
      End If
      SysVAR.FILENAME = DSNM$
      Select Case XOPT$
         Case "WRITE": GoTo 18100
         Case "READ": GoTo 18200
         Case Else: GoTo 18300
      End Select

18100 ' Open TRANSFER File for output
      On Local Error GoTo 18140
18110 'IF Floppy Disk, Prompt for New Diskette
      If FDISK Then
         LSet SysVAR.MVAR = FDRV$: SysVAR.NVAR1 = NDISK
         CA$ = INKEY$: While CA$ <> "": CA$ = INKEY$: Wend
         Call PutMSG("GDCMM043")
         If CA$ = Chr$(27) Then OKAY = 0: GoTo 18300
      End If
      Open DSNM$ For Input As #5: Close #5
      Call PutMSG("GDCMM055")
      If CA$ <> SNGLKEY$(1) Then OKAY = 0: GoTo 18300
      Kill DSNM$: GoTo 18150
18140 Resume 18150
18150 On Local Error GoTo 18240
      Open DSNM$ For Output As #5: OKAY = -1: GoTo 18300

18200 ' Open Exported DATA File for Input
      On Local Error GoTo 18240
      If FDISK Then
         If ((Not QUIET) Or (NDISK > 0) Or (LDiskNUM > 0)) Then
            If NDISK = 0 Then xX$ = "GDCMM045" Else xX$ = "GDCMM046"
            LSet SysVAR.MVAR = FDRV$: SysVAR.NVAR1 = NDISK
            Call PutMSG(xX$)
            If CA$ = Chr$(27) Then OKAY = 0: GoTo 18300
         End If
      End If
      Open DSNM$ For Input As #5
      If EOF(5) Then
        Call ErrBEEP(0): Call PutMSG("GDCMM041")
        GoTo 18000
      End If
      OKAY = -1: GoTo 18300
18240 If (Not QUIET) Or FDISK Then
         XERR = ERR
         Call ERRMessage("XXXXM030", XERR, ERL)
         If FDISK Then Resume 18000
      End If
      OKAY = 0: Resume 18300

18300 ' Return to Caller
      If OKAY Then
         If (NDISK = 0) Then
            GEDLine& = 0: If TAG$(0) <> "" Then INDI = 0: FAMI = 0
         End If
         DiskNUM = NDISK: LDiskNUM = DiskNUM
         GEDFile$ = GED$
      End If
      DISKEOF = Not OKAY

End Sub

Rem $STATIC
Sub GEDCOMStats()

    ' Show TRANSFER Dataset Stats
    'CurFMT = FMT001: Fmt = FMTTAB(CurFMT)
    'FF = Fmt.BGNFLD
    'While FF > 0
    '   FFLD = FLDTAB(FF)
    '   If FFLD.UTXT > 0 Then
    '      xx$ = UTXT$(FFLD.UTXT)
    '      Select Case FFLD.NAME
    '         Case "GDDS": LSet xx$ = GEDFile$
    '         Case "LINS": LSet xx$ = Str$(GEDLine&)
    '         Case "INDI": LSet xx$ = Str$(INDI)
    '         Case "FAMI": LSet xx$ = Str$(FAMI)
    '      End Select
    '      LSet UTXT$(FFLD.UTXT) = xx$
    '   End If
    '   FF = FLDCHN(FF).FWD
    'Wend
    'If Not (PrintERR And (PRT.opt = 1)) Then
    '   Call WinPREP(1): Call WinDisplayDATA(FMT001)
    '   Call WinPREP(2)
    'End If
    Call frmFHSWUPgm.ShowStats("GEDCOM", 11, Str$(INDI))
    Call frmFHSWUPgm.ShowStats("GEDCOM", 12, Str$(FAMI))
    Call frmFHSWUPgm.ShowStats("GEDCOM", 13, Str$(GEDLine&))
    frmFHSWUPgm.Refresh

End Sub

Sub GetLINE()

1800 ' Get Next TRANSFER File Line
     If (CA$ = Chr$(27)) Or (a = 27) Then INGL = 0: Exit Sub
     On Local Error GoTo 1899
     OKAY = -1: INGL = 0: XREF = 0: XID = 0: Z = 0: xX$ = " "
     LSet INTAG$ = " ": LSet ST$ = " ": LSet XVALUE$ = ""
     If DISKEOF Then OKAY = 0: GoTo 1809
     Line Input #5, TLINE$
     GEDLine& = GEDLine& + 1
     TL = Len(TLINE$)
     X = 1: While (MID$(TLINE$, X, 1) = " ") And (X < TL): X = X + 1: Wend
     INGL = Val(MID$(TLINE$, X, 1))
     X = X + 1: GoSub 1810: If X > TL Then Z = 0: GoTo 1809
     If MID$(TLINE$, X, 1) = "@" Then
        Y = InStr(X + 1, TLINE$, "@"): If Y = 0 Then Y = TL + 1
        X = X + 1: a = Asc(MID$(TLINE$, X, 1))
        If (a < 48) Or (a > 57) Then X = X + 1
        xY! = Val(MID$(TLINE$, X, Y - X)): X = Y + 1
        If xY! > 32766 Then xY! = 0
        XREF = xY!
        If Not (X > TL) Then GoSub 1810
     End If
     If Not (X > TL) Then
        Y = X: GoSub 1820
        LSet INTAG$ = MID$(TLINE$, Y, X - Y): GoSub 1810
     End If
     If Not (X > TL) Then
        If MID$(TLINE$, X, 1) = "@" Then
           Y = InStr(X + 1, TLINE$, "@"): If Y = 0 Then Y = TL + 1
           X = X + 1: a = Asc(MID$(TLINE$, X, 1))
           If (a < 48) Or (a > 57) Then X = X + 1
           xY! = Val(MID$(TLINE$, X, Y - X))
           If xY! > 32766 Then xY! = 0
           XID = xY!
           If Y < TL Then LSet ST$ = MID$(TLINE$, Y + 1, 1)
           X = Y: GoSub 1810
        End If
     End If
     If Not (X > TL) Then
        Z = X
        LSet XVALUE$ = MID$(TLINE$, X, TL - X + 1)
        Else
        Z = 0
     End If
     If EOF(5) Then
        DISKEOF = -1
        FDISK = (MID$(GEDFile$, 2, 1) = ":")
        If FDISK Then
           FDRV$ = Left$(GEDFile$, 1)
           FDISK = ((FDRV$ = "A") Or (FDRV$ = "B"))
           If FDISK Then Call GEDCOMOpen("READNQ")
        End If
     End If
     If TAG$(0) <> "" Then
        If INTAG$ = TAG$(30) Then
           INDI = INDI + 1
           Else
           If INTAG$ = TAG$(110) Then FAMI = FAMI + 1
        End If
     End If
1809 Exit Sub

1810 While (X < (TL + 1)) And (MID$(TLINE$, X, 1) = " ")
        X = X + 1
     Wend: Return

1820 While (X < (TL + 1)) And ((MID$(TLINE$, X, 1) <> " ") And (MID$(TLINE$, X, 1) <> "@"))
        X = X + 1
     Wend: Return

1899 ' Error During Input
     Call ERRMessage("GDCMM039", ERR, ERL)
     OKAY = 0: Resume 1809

End Sub

Sub GetSPHONE(xX$)
    X = 1: L = Len(xX$) + 1
    FF2SREC.FON1 = 0: FF2SREC.FON2 = 0: FF2SREC.FON3 = 0
    If X < L Then
       Y = Asc(MID$(xX$, X, 1))
       GoSub GetSPH1: B = X: GoSub GetSPH2
       If X > B Then
          xY! = Val(MID$(xX$, B, X - B))
          If xY! < 32767 Then FF2SREC.FON1 = xY!
          GoSub GetSPH1: B = X: GoSub GetSPH2
          If X > B Then
             xY! = Val(MID$(xX$, B, X - B))
             If xY! < 32767 Then FF2SREC.FON2 = xY!
             GoSub GetSPH1: B = X: GoSub GetSPH2
             If X > B Then
                xY! = Val(MID$(xX$, B, X - B))
                If xY! < 32767 Then FF2SREC.FON3 = xY!
             End If
          End If
       End If
    End If
    Exit Sub

GetSPH1:
   While (X < L) And ((Y < 48) Or (Y > 57))
      X = X + 1: If X < L Then Y = Asc(MID$(xX$, X, 1))
   Wend
   Return

GetSPH2:
   While (X < L) And ((Y > 47) And (Y < 58))
      X = X + 1: If X < L Then Y = Asc(MID$(xX$, X, 1))
   Wend
   Return

End Sub

Sub GetTAGS(OPT$)
   Call PutMSG("GDCMM005")
   OKAY = (a <> 27)
   If OKAY Then
      GDCMTYPE = Val(CA$)
      GoSub BuildTAGS
   End If
   Exit Sub

BuildTAGS:
   Dim TagTABLE As TablePARMS
   Call TableLOAD("GEDCOM", TNUM, 22): If Not OKAY Then Return
   TagTABLE = TABLE(TNUM)
   ReDim TAG$(140)
   Select Case GDCMTYPE
      Case 1: BT = 15: TL = 4
      Case 2: BT = 11: TL = 4
      Case 3: BT = 7: TL = 4
      Case 4: BT = 5: TL = 2
      Case 5: BT = 19: TL = 4
   End Select
   TAG$(0) = String$(TL, ".")
   Y = TagTABLE.FT
   While Y > 0
       TAGNM$ = MID$(TABDATA$(Y), 1, 4)
       TAGVALUE$ = MID$(TABDATA$(Y), BT, TL)
       If RTrim$(TAGVALUE$) = "" Then TAGVALUE$ = TAG$(0)
       If OPT$ = "EXPORT" Then TAGVALUE$ = RTrim$(TAGVALUE$)
       Select Case TAGNM$
          Case "H000": TAG$(1) = TAGVALUE$   ' Header Record
          Case "H020": TAG$(4) = TAGVALUE$   ' Submitter's Section
          Case "H021": TAG$(5) = TAGVALUE$   ' Submitter's Name
          Case "H030": TAG$(6) = TAGVALUE$   ' Submitter's Address Line 1
          Case "H031": TAG$(7) = TAGVALUE$   ' Submitter's Address Line 2,3
          Case "H032": TAG$(8) = TAGVALUE$   ' Submitter's Phone Number
          Case "H040": TAG$(9) = TAGVALUE$   ' Source System Section (2.1)
          Case "H041": TAG$(10) = TAGVALUE$  ' Source System Identifier (FHS)
          Case "H042": TAG$(11) = TAGVALUE$  ' Source Version
          Case "H043": TAG$(12) = TAGVALUE$  ' Source Full Name
          Case "H045": TAG$(13) = TAGVALUE$  ' Dest System Identifier
          Case "H050": TAG$(14) = TAGVALUE$  ' DATE GEDCOM File Built
          Case "H060": TAG$(15) = TAGVALUE$  ' File Name
          Case "H061": TAG$(16) = TAGVALUE$  ' Number of II records
          Case "H062": TAG$(17) = TAGVALUE$  ' Number of II records"
          Case "H063": TAG$(20) = TAGVALUE$  ' GEDCOM Rules Used
          Case "H090": TAG$(18) = TAGVALUE$  ' Submitter's Comments
          Case "H091": TAG$(19) = TAGVALUE$  ' Submitter's Comments (Cont)
          Case "I000": TAG$(30) = TAGVALUE$  ' Individual Information Record
          Case "I001": TAG$(31) = TAGVALUE$  ' ID# of Source Record in FHS
          Case "I010": TAG$(32) = TAGVALUE$  ' Full Name (Given/Surname/)
          Case "I011": TAG$(33) = TAGVALUE$  ' "Surname Use" Field
          Case "I015": TAG$(34) = TAGVALUE$  ' Sex Code
          Case "I020": TAG$(35) = TAGVALUE$  ' Birth Information
          Case "I021": TAG$(36) = TAGVALUE$  ' Birth Date
          Case "I023": TAG$(38) = TAGVALUE$  ' Birth Time
          Case "I025": TAG$(39) = TAGVALUE$  ' Birth Place
          Case "I030": TAG$(40) = TAGVALUE$  ' Death Information
          Case "I031": TAG$(41) = TAGVALUE$  ' Death Date
          Case "I033": TAG$(43) = TAGVALUE$  ' Death Time
          Case "I035": TAG$(45) = TAGVALUE$  ' Death Place
          Case "I040": TAG$(50) = TAGVALUE$  ' Marriage Information
          Case "I041": TAG$(51) = TAGVALUE$  ' "Pointer" to most recent
          Case "I044": TAG$(54) = TAGVALUE$  ' Parent's Family Record
          Case "I045": TAG$(55) = TAGVALUE$  ' Adoption: HUSB, WIFE, BOTH
          Case "I046": TAG$(56) = TAGVALUE$  ' Adoption by Mother (2.0 n.s.)
          Case "I047": TAG$(57) = TAGVALUE$  ' Adoption by Father (2.0 n.s.)
          Case "I048": TAG$(58) = TAGVALUE$  ' Sibling Section (2.0)
          Case "I049": TAG$(59) = TAGVALUE$  ' Next Oldest Sibling
          Case "I050": TAG$(60) = TAGVALUE$  ' Oldest Child
          Case "I051": TAG$(61) = TAGVALUE$  ' Father Identifier
          Case "I052": TAG$(62) = TAGVALUE$  ' Younger Sibling, Same Father
          Case "I053": TAG$(63) = TAGVALUE$  ' Mother Identifier
          Case "I054": TAG$(64) = TAGVALUE$  ' Younger Sibling, Same Mother
          Case "E000": TAG$(70) = TAGVALUE$  ' Event Record
          Case "E001": TAG$(71) = TAGVALUE$  ' Event Type
          Case "E002": TAG$(72) = TAGVALUE$  ' Event Importance Level
          Case "E003": TAG$(75) = TAGVALUE$  ' Event Date
          Case "E004": TAG$(76) = TAGVALUE$  ' Date Status Field
          Case "E005": TAG$(77) = TAGVALUE$  ' Event Time
          Case "E006": TAG$(78) = TAGVALUE$  ' Event Place
          Case "N000": TAG$(79) = TAGVALUE$  ' First Comment Line
          Case "N001": TAG$(80) = TAGVALUE$  ' Continuation of Comments
                                             ' Fields for Extra Records
          Case "X101": TAG$(81) = TAGVALUE$  ' Beginning Effective Date
          Case "X103": TAG$(82) = TAGVALUE$  ' Ending Effective Date
          Case "X104": TAG$(83) = TAGVALUE$  ' Ending Status
          Case "X105": TAG$(84) = TAGVALUE$  ' First Comment Line (Misc Info)
          Case "X106": TAG$(85) = TAGVALUE$  ' Continued Comments (Misc Info)
          Case "X120": TAG$(86) = TAGVALUE$  ' Education Record
          Case "X121": TAG$(87) = TAGVALUE$  ' Subject (may be two)
          Case "X122": TAG$(88) = TAGVALUE$  ' Education Level
          Case "X123": TAG$(89) = TAGVALUE$  ' Degree or Certificate Earned
          Case "X130": TAG$(90) = TAGVALUE$  ' Work Record
          Case "X131": TAG$(91) = TAGVALUE$  ' Work Type
          Case "X140": TAG$(92) = TAGVALUE$  ' Military Record
          Case "X141": TAG$(93) = TAGVALUE$  ' Military Rank
          Case "X150": TAG$(95) = TAGVALUE$  ' Medical Record
          Case "X151": TAG$(96) = TAGVALUE$  ' Medical Diagnosis
          Case "A010": TAG$(100) = TAGVALUE$ ' Address Record
          Case "A011": TAG$(101) = TAGVALUE$ ' Address Line 1
          Case "A012": TAG$(102) = TAGVALUE$ ' Address Line(s) continued
          Case "A013": TAG$(103) = TAGVALUE$ ' City
          Case "A014": TAG$(104) = TAGVALUE$ ' State/Province
          Case "A015": TAG$(105) = TAGVALUE$ ' Zip/Postal Code
          Case "A016": TAG$(106) = TAGVALUE$ ' Country
          Case "A017": TAG$(107) = TAGVALUE$ ' Phone Number
          Case "A018": TAG$(108) = TAGVALUE$ ' First Comment Line (ADDRESS)
          Case "A019": TAG$(109) = TAGVALUE$ ' Next Comment Line  (ADDRESS)
          Case "F000": TAG$(110) = TAGVALUE$ ' Family/Marriage Record
          Case "F001": TAG$(111) = TAGVALUE$ ' Created for Unmarried Parents
          Case "F010": TAG$(112) = TAGVALUE$ ' Husband Identifier
          Case "F011": TAG$(113) = TAGVALUE$ ' Husb/Wife II Number (2.0)
          Case "F012": TAG$(114) = TAGVALUE$ ' Husb/Wife Other Marriages (2.0)
          Case "F015": TAG$(115) = TAGVALUE$ ' Wife's Identifier
          Case "F020": TAG$(120) = TAGVALUE$ ' Marriage Beginning Event
          Case "F021": TAG$(121) = TAGVALUE$ ' Beginning Marriage Status
          Case "F022": TAG$(122) = TAGVALUE$ ' Marriage Date
          Case "F024": TAG$(124) = TAGVALUE$ ' Place of Marriage
          Case "F025": TAG$(125) = TAGVALUE$ ' Divorce Event (2.0)
          Case "F026": TAG$(126) = TAGVALUE$ ' Divorce Flag (2.1)
          Case "F027": TAG$(127) = TAGVALUE$ ' Ending Status
          Case "F028": TAG$(128) = TAGVALUE$ ' Marriage Termination Date
          Case "F029": TAG$(129) = TAGVALUE$ ' Place of Marriage Termination
          Case "F030": TAG$(130) = TAGVALUE$ ' Husband's next earlier marriage
          Case "F031": TAG$(131) = TAGVALUE$ ' Wife's next earlier marriage
          Case "F040": TAG$(132) = TAGVALUE$ ' Child Identifier
          Case "F041": TAG$(133) = TAGVALUE$ ' Youngest Child's Identifier
          Case "F042": TAG$(134) = TAGVALUE$ ' Adoption (2.1)
          Case "F090": TAG$(135) = TAGVALUE$ ' First Comment Line   (FAMILY)
          Case "F091": TAG$(136) = TAGVALUE$ ' Comment Continuation (FAMILY)
          Case "T000": TAG$(140) = TAGVALUE$ ' Trailer (END) Record
       End Select
       Y = TABDATACHN(Y).FWD
   Wend
   Call TableFREE("GEDCOM")
   Return

End Sub

Sub PointerFILES(OPT$)
   Select Case OPT$
      Case "OPEN": GoSub 19000
      Case "DELETE": GoSub 19050
   End Select
   Exit Sub

19000 ' Open Pointer Workfiles
      OKAY = 0
      WF1$ = WRKFPFX$ + "FH$PTR1": WF2$ = WRKFPFX$ + "FH$PTR2"
      GoSub 19050
      On Local Error GoTo 19010
      SysVAR.MVAR = WF1$
      Open WF1$ For Random Access Read Write As #6 Len = 20
      SysVAR.MVAR = WF2$
      Open WF2$ For Random Access Read Write As #7 Len = 14
      OKAY = -1
      GoTo 19019
19010 ' Error Opening Work Files
      Call ERRMessage("XXXXM030", ERR, ERL)
      If a = 27 Then Resume 19019 Else Resume 19000
19019 Return

19050 ' KILL Pointer Workfiles
19055 On Local Error GoTo 19056
      Kill WF1$: GoTo 19060
19056 Resume 19060
19060 On Local Error GoTo 19061
      Kill WF2$: GoTo 19070
19061 Resume 19070
19070 Return

End Sub

