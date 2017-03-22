Attribute VB_Name = "modFHSWAcht"
DefInt A-Z
Dim LTAB$(), NODEPT(), DCHAR$, PCHAR$, SCHAR$
Dim DFNUM(), DFPP()

Private Function FNODD(DA#)
dsx$ = FNDS$(DA#): dsl = Len(dsx$)
If dsl > 16 Then
   If MID$(dsx$, dsl - 2, 1) = "+" Then dsl = dsl - 4
End If
FNODD = (Val(MID$(dsx$, dsl, 1)) Mod 2)
'    FNODD = (Val(Right$(FNDS$(DA#), 1)) Mod 2)

End Function

Public Sub RptACHT(OPT$)
1  Rem $INCLUDE: 'FHSCOMON.BAS'
   PN$ = " FAMILY HISTORY SYSTEM - FHSWACHT - Ancestor Chart/MAP Program"
   CY$ = " (C) Copyright 1985-1998 by Phillip E. Brown"
   
   Select Case OPT$
      Case "INIT": GoTo 40
      Case "PRINT": GoTo 1000
   End Select
   
40 ' Format Display
   Call PgmPREP("FHSACHT")
   Load frmFHSWRPgm
   Call RptLOAD("ACHT"): Call RptLOAD("AMAP")
   frmFHSWRPgm.RPSetup.AddItem "ACHT"
   frmFHSWRPgm.RPSetup.AddItem "AMAP"
   frmFHSWRPgm.RPSetup_Click
   frmFHSWRPgm.Show
   Exit Sub

1000 'PROCESS REPORT OPTIONS
    ReDim DFNUM(1), DFPP(1), RFLTH(1), RFLBL$(1)
    ReDim LTAB$(62), NODEPT(33, 2, 2)
    Call DataINIT("ACHTNODE")
    ' Create NODEPT() Array for 4GEN Chart
    For I = 0 To 62: LTAB$(I) = Space$(160): Next
    For X = 1 To 2
        For I = 1 To 33
          Input #5, NODEPT(I, 1, X), NODEPT(I, 2, X)
        Next I
    Next X
    Close #5
    ' 4 GEN Chart Node Points
    'Data 0, 29, 4, 13, 4, 45, 44, 5, 44, 21, 44, 37, 44, 53
    'Data 85, 1, 85, 10, 85, 17, 85, 26, 85, 33, 85, 42, 85, 49, 85, 58
    'Data 126, 0, 126, 2, 126, 9, 126, 11, 126, 16, 126, 18, 126, 25, 126, 27
    'Data 126, 32, 126, 34, 126, 41, 126, 43, 126, 48, 126, 50, 126, 57, 126, 59, 0, 0, 4, 56
    ' 5 GEN Chart Node Points
    'Data 0, 32, 4, 16, 4, 48, 6, 7, 6, 23, 6, 39, 6, 55
    'Data 48, 3, 48, 11, 48, 19, 48, 27, 48, 35, 48, 43, 48, 51, 48, 59
    'Data 89, 1, 89, 5, 89, 9, 89, 13, 89, 17, 89, 21, 89, 25, 89, 29
    'Data 89, 33, 89, 37, 89, 41, 89, 45, 89, 49, 89, 53, 89, 57, 89, 61, 0, 0, 4, 60
   ' Create CHART Symbol Set Variables - SCHAR$=Screen, PCHAR$=Printer
    SCHAR$ = Chr$(196) + Chr$(180) + Chr$(179) + Chr$(218) + Chr$(192) + Chr$(197)
    PCHAR$ = MID$(PDFSetup.SET1, 7, 1) + MID$(PDFSetup.SET1, 5, 1) + MID$(PDFSetup.SET1, 4, 1) + MID$(PDFSetup.SET1, 1, 1) + MID$(PDFSetup.SET1, 8, 1) + MID$(PDFSetup.SET1, 11, 1)
    SCHAR$ = PCHAR$  ' Added 9/4/99
    DCHAR$ = Space$(6)
    Select Case RptCODE
       Case "ACHT": Call PutMSG("Select:  1) Blank Chart  2) Filled in Chart")
                    Call GetRESP
                    Call ACHTReport(CA$)
       Case "AMAP": Call AMAPReport
     End Select
End Sub

Private Function FNDS$(DA#)
   FNDS$ = Str$(Val(Str$(DA#)))
End Function

Rem $DYNAMIC
Sub ACHTReport(XOPT$)
   COPT$ = XOPT$: NFMT$ = Space$(80): xDTXT$ = Space$(80): ANS$ = ""
   Call RptOPEN("ACHT"): If Not OKAY Then GoTo 60
   Call GetRVAR("ATAGE", X): xATAGE$ = RVAR$(X)
   Call GetRVAR("FOR", X): xFOR$ = RVAR$(X)
   Call GetRVAR("CONTFROM", X): xCONTFROM$ = RVAR$(X)
   Call GetRVAR("CHART", X): xFROMCHRT$ = RVAR$(X)
   Call GetRVAR("PAGE", X): xFROMPAGE$ = RVAR$(X)
   Call GetRVAR("IND1LIN", X): xIND1LIN$ = RVAR$(X)
   Call GetRVAR("DATEPREP", X): xDATEPREP$ = RVAR$(X)
   Call GetRVAR("SPOUSE1", X): xSPOUSE1$ = RVAR$(X)
   Call PutMSG("ACHTM003")
   If CA$ = Chr$(27) Then GoTo 60
   CTYPE = Val(CA$): If (CTYPE < 1) Or (CTYPE > 2) Then CTYPE = 1
   Call BldChartDEF(CTYPE)
   Select Case COPT$
      Case "1": GoTo X150
      Case Else: GoTo X400
   End Select

60 Close: Exit Sub

X150: ' PRODUCE BLANK CHART
    RptOPTION.BLDREF = 0
    ShowID = 0: NOSHOW = 0: ShowLIN = 0
    Call PrintOPEN: If A = 27 Then GoTo 60
    If PRT.OPT = 1 Then
       Copies = 1
       Else
       Call PutMSG("ACHTM004")
       Copies = Val(CA$)
    End If
    If Copies > 0 Then
       SysVAR.CHARTID = ""
       SysVAR.NVAR2 = Copies
       If OKAY Then
          GoSub 900
          Mid$(LTAB$(0), Z0 + 20, 255) = xFOR$
          Mid$(LTAB$(1), 1, 255) = xCONTFROM$: Z = InStr(LTAB$(1), "  ") + 1
          Mid$(LTAB$(1), Z, 255) = xFROMCHART$: Z1 = InStr(LTAB$(1), "  ") + 1
          Mid$(LTAB$(3), 1, 255) = xDATEPREP$: Z3 = InStr(LTAB$(3), "  ") + 1
          For CC = 1 To Copies
              SysVAR.NVAR1 = CC
              Call PutMSG("ACHTM005")
              GoSub X1000
              If A = 27 Then CC = Copies
          Next CC
          PRT.AGAIN = 0: Call RptCLOSE
       End If
    End If
    Call PutMSG("")
    Exit Sub

X400: 'PRODUCE ANCESTOR CHARTS
    BD$ = Space$(12): DD$ = BD$: BPL$ = Space$(41): DPL$ = BPL$: AGE$ = "   "
    Call RptOptGET("ALIN", X): ShowLIN = (X = 1): zALIN$ = ""
    Call RptOptGET("LKPG", X): PAGELNK = (X = 1)
    If RptOPTION.PRID = 2 Then PRTHDRID = -1: RptOPTION.PRID = 0
    Call FAMOpen("READ"): If Not OKAY Then GoTo 60
    Call RWRKOpen("READ", 1, 0): If Not OKAY Then GoTo 60
    Call GetRVAR("SAMEAS", X): xSAMEAS$ = RVAR$(X)
    Call GetRVAR("ONCHART", X): xONCHART$ = RVAR$(X)
    Call GetRVAR("ONPAGE", X): xONPAGE$ = RVAR$(X)
    Call GetRVAR("UNKNOWN", X): xUNKNOWN$ = RVAR$(X)
    MAXGL = Asc(WFHDR.AMAXLV): WFHDR.DMAXLV = Chr$(0)
    AMAXLV = MAXGL: HIGBL = MAXGL \ 4: RREFXData$ = ""
    Call PrintOPEN: If A = 27 Then GoTo 60
    Call RWRKGet("BASE", RC, PTR&): If RC = 9 Then GoTo 60
    SysVAR.SUBJECT = RTrim$(FMTNAME$(1))
    GoSub 900
    ReDim GENDPG(HIGBL + 1), GBGN&(HIGBL + 1), L(3)
    ReDim BAPGNUM(FF1Hdr.MAXID)
    GBGN&(0) = PTR&: GENDPG(0) = PAGE.BGNPAGE
    I = 0: CHTOT = 1: CHTOTPG = PAGE.BGNPAGE
    BAPGNUM(WFANCREC.ANCID) = PAGE.BGNPAGE
    While (I < MAXGL) And (WFAGLREC.NEXT <> 0)
       Get #5, WFAGLREC.NEXT + 1, WFAGLREC
       I = I + 1
       If (I Mod 4) = 0 Then
          BGL = I \ 4
          GBGN&(BGL) = WFAGLREC.ANCPTR
          PTR& = WFAGLREC.ANCPTR: J = WFAGLREC.TOTREC
          While J > 0
             Get #5, PTR& + 1, WFANCREC: PTR& = PTR& + Len(WFANCREC)
             If WFANCREC.ANCID > 0 Then
                CHTOT = CHTOT + 1: CHTOTPG = CHTOTPG + 1
                BAPGNUM(WFANCREC.ANCID) = CHTOTPG
             End If
             J = J - 1
          Wend
          GENDPG(BGL) = CHTOTPG
       End If
    Wend
    On Local Error GoTo 0
    DUPID = 0: xGBL = 0: BGL = 0: GLD = 0: ANS$ = SNGLKEY$(2)
    CHNUMPG = PAGE.BGNPAGE: CHNUM = 1
    BPTR& = GBGN&(0)
    GoTo X410

X405: ' Get Next ChartID
    CHNUMPG = CHNUMPG + 1: CHNUM = CHNUM + 1
    If CHNUMPG > GENDPG(xGBL) Then
       xGBL = xGBL + 1: BGL = 4 * xGBL
       If xGBL > HIGBL Then GoTo 1100
       BPTR& = GBGN&(xGBL)
       Else
       BPTR& = BPTR& + Len(WFANCREC)
    End If

X410: ' Process Next Chart
    Call RWRKGet("PTRREC", RC, BPTR&): If Not OKAY Then GoTo 60
    If (CTYPE = 2) And (WFANCREC.ANCID < 0) Then
       BPTR& = BPTR& + Len(WFANCREC): GoTo X410
    End If
    BCHARTID$ = XChartID$(BGL, WFANCREC.LINEAGE)
    SysVAR.CHARTID = BCHARTID$
    SysVAR.NVAR1 = CHNUM: SysVAR.NVAR2 = CHTOT
    If ANS$ = SNGLKEY$(4) Then GoTo X430

X420: ' Get Approval to Print Chart
    Call PutMSG("ACHTM006")
    ANS$ = CA$: If Asc(ANS$) = 27 Then GoTo 1100
    If (ANS$ = SNGLKEY$(4)) Or (ANS$ = SNGLKEY$(1)) Or (ANS$ = SNGLKEY$(13)) Then GoTo X430
    If (ANS$ <> SNGLKEY$(2)) Then Call ErrBEEP(0): GoTo X420
    GoTo X405

X430: ' Print Chart with Base Record BPTR&
    Call PutMSG("XXXXM045")
    If RptOPTION.BLDREF Then
       RREFREC.DGL = Chr$(0): RREFREC.AGL = Chr$(BGL + GLD)
       If Not PAGELNK Then RREFXData$ = BCHARTID$
       PAGE.RPAGE = CHNUMPG
    End If
    BLNUM# = WFANCREC.LINEAGE
    DUPID = (WFANCREC.ANCID < 0)
    F1RID = Abs(WFANCREC.ANCID)
    RSTAT = WFANCREC.PSTAT And 3
    Call FF1GetRec(F1RID)
    TL = 0
    If L(0) > 0 Then Mid$(LTAB$(0), Z0, L(0)) = Space$(L(0))
    Mid$(LTAB$(0), Z0, 255) = RTrim$(BCHARTID$): Z = InStr(Z0, LTAB$(0), "   ") + 1
    Mid$(LTAB$(0), Z, 255) = xFOR$: Z = InStr(Z, LTAB$(0), "  ") + 1
    Mid$(LTAB$(0), Z, 52) = SysVAR.SUBJECT: L(0) = InStr(Z, LTAB$(0), "  ") - Z0
    If Not (ShowLIN Or NOSHOW) Then
       TL = TL + 1
       If L(TL) > 0 Then Mid$(LTAB$(TL), 1, 255) = Space$(L(TL))
       Mid$(LTAB$(TL), 1, 255) = xIND1LIN$: Z = InStr(LTAB$(TL), "  ")
       Mid$(LTAB$(TL), Z, 255) = FNDS$(BLNUM#)
       If PRTHDRID Or RptOPTION.PRID Then
          Z = InStr(Z, LTAB$(TL), "   ")
          Mid$(LTAB$(TL), Z + 2, 255) = "(" + RFLBL$(3) + fnx$(F1RID) + ")"
       End If
       L(TL) = InStr(Z, LTAB$(TL), "   ") - 1
    End If
    If BGL > 0 Then
       TL = TL + 1
       If L(TL) > 0 Then Mid$(LTAB$(TL), 1, 255) = Space$(L(TL))
       Mid$(LTAB$(TL), 1, 255) = xCONTFROM$: Z = InStr(LTAB$(TL), "  ") + 1
       J = 4
       While (J > 0) And (WFANCREC.ANCCH <> 0)
          Get #5, WFANCREC.ANCCH + 1, WFANCREC
          J = J - 1
       Wend
       If PAGELNK Then
          PPAGE = BAPGNUM(WFANCREC.ANCID)
          Mid$(LTAB$(TL), Z, 255) = xFROMPAGE$: Z = InStr(LTAB$(TL), "  ")
          Mid$(LTAB$(TL), Z, 255) = Str$(PPAGE)
          Else
          Mid$(LTAB$(TL), Z, 255) = xFROMCHRT$: Z = InStr(LTAB$(TL), "  ") + 1
          CHAINID$ = XChartID$(BGL - 4, WFANCREC.LINEAGE)
          Mid$(LTAB$(TL), Z, 255) = CHAINID$
       End If
       Get #5, BPTR& + 1, WFANCREC
       L(TL) = InStr(Z, LTAB$(TL), "  ") - 1
    End If
    TL = TL + 1
    If L(TL) > 0 Then Mid$(LTAB$(TL), 1, 255) = Space$(L(TL))
    Mid$(LTAB$(TL), 1, 255) = xDATEPREP$
    Z = InStr(LTAB$(TL), "  ") + 1: Mid$(LTAB$(TL), Z, 255) = SysVAR.DATE
    L(TL) = InStr(Z, LTAB$(TL), "  ") - 1

    ' Process NODEs 1,33 (GEN 0)
    ReDim APTR&(4), NDNUM(4), xMID(4)
    xAPTR& = BPTR&: APTR&(0) = xAPTR&:
    NDNUM(0) = 1: xMID(0) = FF1Rec.MID
    GoSub 920: APARENT& = WFANCREC.FIRSTANC
    GLD = 0: xNDNUM = 1: GoSub 940
    If Not (SPID < 0) Then
       If SPID = 0 Then Unknown = -1 Else Call FF1GetRec(SPID)
       xNDNUM = 33: RSTAT = 0: S = 0: DUPID = 0: GoSub 945
       PAGE.xType = 3: PAGE.XID = SPID
       Call RREFAdd: Call RREFBld
    End If
    S = 0: Unknown = 0: GLD = 1: xAPTR& = APARENT&

X440: ' Process NODEs 2-31 (GEN 1-4)
    While xAPTR& <> 0: APTR&(GLD) = xAPTR&
       'X = 1 / 0
       Get #5, xAPTR& + 1, WFANCREC
       DUPID = (WFANCREC.ANCID < 0): F1RID = Abs(WFANCREC.ANCID)
       APARENT& = WFANCREC.FIRSTANC
       Call FF1GetRec(F1RID)
       xMID(GLD) = FF1Rec.MID
       xNDNUM = 2 * NDNUM(GLD - 1)
       If F1RID = xMID(GLD - 1) Then
          xMID(GLD - 1) = -xMID(GLD - 1)
          xNDNUM = xNDNUM + 1
       End If
       NDNUM(GLD) = xNDNUM: GoSub 940
       If DUPID Or (GLD = 4) Then xMID(GLD) = 0
       If (GLD < 4) And (APARENT& <> 0) Then
          GLD = GLD + 1: xAPTR& = APARENT&
          Else
          xAPTR& = 0
          While (GLD > -1) And (xAPTR& = 0)
             If ((Asc(WFHDR.WFRULES) And 8) = 0) Then
                If (xMID(GLD) > 0) Then
                   GLD = GLD + 1
                   xAPTR& = APTR&(GLD) + Len(WFANCREC)
                   Else
                   GLD = GLD - 1
                End If
             End If
          Wend
       End If
    Wend

X450: ' Print Lines of CHART
    GoSub X1000: If A = 27 Then ANS$ = ""
    GoTo X405

900 ' INITIALIZE CHART
    For I = 0 To 62: LSet LTAB$(I) = " ": Next I
    Mid$(LTAB$(0), 1, 255) = RDFReport.NAME: Z0 = InStr(LTAB$(0), "  ") + 1
    'If PRT.OPT = 1 Then
    '   LSet DCHAR$ = SCHAR$
    '   Else
       LSet DCHAR$ = PCHAR$
    'End If
    LSet PLINE$ = String$(52, Asc(DCHAR$))
    For I = 1 To 15
       x1 = NODEPT(I, 1, CTYPE): Y1 = NODEPT(I, 2, CTYPE)
       X2 = NODEPT(2 * I, 1, CTYPE): Y2 = NODEPT(2 * I, 2, CTYPE)
       X3 = NODEPT(2 * I + 1, 1, CTYPE): Y3 = NODEPT(2 * I + 1, 2, CTYPE)
       ' Draw Horizontal Line for Name
       Mid$(LTAB$(Y1), x1 + 1, X2 - x1) = PLINE$
       ' Insert "T" at End of Line
       Mid$(LTAB$(Y1), X2, 1) = MID$(DCHAR$, 2, 1)
       ' Draw Vertical Line to Father's Node
       For Y = Y1 - 1 To Y2 + 1 Step -1
          Mid$(LTAB$(Y), X2, 1) = MID$(DCHAR$, 3, 1)
          Next Y
       ' Insert Upper Left Corner
       Mid$(LTAB$(Y2), X2, 1) = MID$(DCHAR$, 4, 1)
       ' Draw Vertical Line to Mother's Node
       For Y = Y1 + 1 To Y3 - 1 Step 1
          Mid$(LTAB$(Y), X3, 1) = MID$(DCHAR$, 3, 1)
          Next Y
       ' Insert Lower Left Corner
       Mid$(LTAB$(Y3), X3, 1) = MID$(DCHAR$, 5, 1)
       ' Draw Name Line for Link Nodes on 5 Gen Chart
       If CTYPE = 2 Then
          If I > 7 Then
             Mid$(LTAB$(Y2), X2 + 1, 125 - X2) = PLINE$
             Mid$(LTAB$(Y3), X3 + 1, 125 - X3) = PLINE$
          End If
       End If
    Next I
    I = 1: GoSub 905
    If (Not ShowLIN) Then
       Mid$(LTAB$(Y - 1), NODEPT(1, 1, CTYPE) + 2, 2) = "1."
    End If
    For I = 2 To 15: GoSub 905: Next I
    I = 33: GoSub 905
    LSet xDTXT$ = xSPOUSE1$
    Mid$(LTAB$(Y - 1), X - 2, NMLTH + 2) = xDTXT$
    GoTo 910

905 ' Insert Entry Numbers, Field Labels and Blank Data Fields
    X = NODEPT(I, 1, CTYPE) + 1: Y = NODEPT(I, 2, CTYPE)
    If I < 16 Then X2 = NODEPT(2 * I, 1, CTYPE) Else X2 = 124
    If I < (2 - 2 * (CTYPE = 2)) Then
       If I = 1 Then X = X2 + 2 Else X = X2 + 1
       NMLTH = NODEPT(16, 1, CTYPE) - X
       DLTH = NODEPT(4 * CTYPE, 1, CTYPE) - X - 1
       Else
       If I < 16 Then
          NMLTH = X2 - X: DLTH = NMLTH - 1
          Else
          NMLTH = NODEPT(7 * CTYPE, 1, CTYPE) - X - 2
          DLTH = NMLTH
       End If
    End If
    S = ((I + 1) Mod 2) - (I = 1)
    LSet NFMT$ = ""
    If (Not ShowLIN) Then
       If (I > 1) And (I < 33) Then
          LSet NFMT$ = Right$(Str$(I), 2) + "."
       End If
    End If
    Mid$(LTAB$(Y + (I < 33)), X, NMLTH) = NFMT$
    If (I > 1) And (I < 33) Then X = X + 1
    If (CTYPE = 1) Then
       LSet xDTXT$ = RFLBL$(5): Mid$(LTAB$(Y + 1), X, DLTH) = xDTXT$
       LSet xDTXT$ = RFLBL$(6): Mid$(LTAB$(Y + 2), X, DLTH) = xDTXT$
       If (S = 1) Then
          LSet xDTXT$ = RFLBL$(8): Mid$(LTAB$(Y + 3), X, DLTH) = xDTXT$
          LSet xDTXT$ = RFLBL$(9): Mid$(LTAB$(Y + 4), X, DLTH) = xDTXT$
       End If
       LSet xDTXT$ = RFLBL$(11): Mid$(LTAB$(Y + 3 + 2 * S), X, DLTH) = xDTXT$
       LSet xDTXT$ = RFLBL$(12): Mid$(LTAB$(Y + 4 + 2 * S), X, DLTH) = xDTXT$
       Else
       LSet xDTXT$ = RFLBL$(5): Z = InStr(xDTXT$, "  ")
       Mid$(xDTXT$, Z + 13, 255) = RFLBL$(6)
       Mid$(LTAB$(Y + 1), X, DLTH) = xDTXT$
       If (S = 1) Then
          LSet xDTXT$ = RFLBL$(8): Z = InStr(xDTXT$, "  ")
          Mid$(xDTXT$, Z + 13, 255) = RFLBL$(9)
          Mid$(LTAB$(Y + 2), X, DLTH) = xDTXT$
       End If
       LSet xDTXT$ = RFLBL$(11): Z = InStr(xDTXT$, "  ")
       Mid$(xDTXT$, Z + 13, 255) = RFLBL$(12)
       Mid$(LTAB$(Y + 2 + S), X, DLTH) = xDTXT$
    End If
    If (I > 7) And (I < 16) Then
       Y2 = NODEPT(2 * I, 2, CTYPE): Y3 = NODEPT(2 * I + 1, 2, CTYPE)
       If CTYPE = 2 Then
          DLTH = 123 - X2: NMLTH = DLTH + 1
          If Not ShowLIN Then Mid$(NFMT$, 1, 3) = LTrim$(Str$(2 * I)) + "."
          Mid$(LTAB$(Y2 - 1), X2 + 1, NMLTH) = NFMT$
          If Not ShowLIN Then Mid$(NFMT$, 1, 2) = LTrim$(Str$(2 * I + 1))
          Mid$(LTAB$(Y3 - 1), X2 + 1, NMLTH) = NFMT$
          LSet xDTXT$ = RFLBL$(5): Z = InStr(xDTXT$, "  ")
          Mid$(xDTXT$, Z + 16, 255) = RFLBL$(11)
          Mid$(LTAB$(Y2 + 1), X2 + 2, DLTH) = xDTXT$
          Mid$(LTAB$(Y3 + 1), X2 + 2, DLTH) = xDTXT$
          X2 = 124
       End If
       LSet xDTXT$ = ""
       Mid$(LTAB$(Y2), X2 + 1, 50) = xDTXT$
       Mid$(LTAB$(Y3), X2 + 1, 50) = xDTXT$
    End If
    Return

910 ' Draw Name Lines for Entry 1 (and Entries 2,3 for 5 Gen Chart)
    X = NODEPT(2, 1, CTYPE): Y = NODEPT(1, 2, CTYPE)
    Mid$(LTAB$(Y), X, 1) = MID$(DCHAR$, 6, 1)
    Mid$(LTAB$(Y), X + 1, 52) = PLINE$
    If (CTYPE = 2) Then
       For I = 2 To 3
          X = NODEPT(2 * I, 1, 2): Y = NODEPT(I, 2, 2)
          Mid$(LTAB$(Y), X, 1) = MID$(DCHAR$, 6, 1)
          Mid$(LTAB$(Y), X + 1, 52) = PLINE$
       Next I
    End If
    Return

920 ' Space Fill Entries on Chart
    For I = 1 To 15: GoSub 905: Next I
    I = 33: GoSub 905
    Return

940  ' Process NDNUM Entry on Chart
    SPID = -1
    If (CTYPE = 2) And (xNDNUM > 15) Then S = 0 Else S = (((xNDNUM + 1) Mod 2) - (xNDNUM = 1))
    LNUM# = WFANCREC.LINEAGE
    If RptOPTION.BLDREF Then
       If (CTYPE = 2) Or (xNDNUM < 16) Then
          RREFREC.AGL = Chr$(BGL + GLD)
          PAGE.xType = 1: PAGE.XID = Abs(WFANCREC.ANCID)
          PAGE.RLINE = xNDNUM
          Call RREFAdd: Call RREFBld
       End If
    End If
945 X = NODEPT(xNDNUM, 1, CTYPE) + 1: Y = NODEPT(xNDNUM, 2, CTYPE)
    If (CTYPE = 1) And (xNDNUM > 15) And (xNDNUM < 33) Then X = X - 1: GoTo 960
    If xNDNUM < 16 Then X2 = NODEPT(2 * xNDNUM, 1, CTYPE) Else X2 = 124
    If xNDNUM < (2 - 2 * (CTYPE = 2)) Then
       If xNDNUM = 1 Then X = X2 + 2 Else X = X2 + 1
       NMLTH = NODEPT(16, 1, CTYPE) - X
       DLTH = NODEPT(4 * CTYPE, 1, CTYPE) - X - 1
       Else
       If xNDNUM < 16 Then
          NMLTH = X2 - X: DLTH = NMLTH - 1
          Else
          If (xNDNUM < 33) Then
             DLTH = 124 - X: NMLTH = DLTH + 1
             Else
             NMLTH = NODEPT(7 * CTYPE, 1, CTYPE) - X - 2
             DLTH = NMLTH
          End If
       End If
    End If
    LSet NFMT$ = "": Z = 1
    If (xNDNUM < 33) Then
       If ShowLIN Then
          Mid$(NFMT$, 1, 255) = FNDS$(LNUM#)
          Z = InStr(NFMT$, "  ") + 1
          Else
          If (xNDNUM > 1) Then
             Mid$(NFMT$, 1, 3) = Right$(Str$(xNDNUM), 2) + ".": Z = 4
          End If
       End If
    End If
    If Unknown Then
       Mid$(NFMT$, Z, 255) = "(" + xUNKNOWN$ + ")"
       Else
       If RptOPTION.PRID Then
          Mid$(NFMT$, Z, 255) = "(" + LTrim$(Str$(FF1Rec.RID)) + ")"
          Z = InStr(Z, NFMT$, " ")
       End If
       Mid$(NFMT$, Z, 255) = FMTNAME(0)
    End If
    Mid$(LTAB$(Y - 1 - (xNDNUM = 33)), X, NMLTH) = NFMT$
    If (xNDNUM > 1) And (xNDNUM < 33) Then X = X + 1
    w = Y + 3 + S * (2 + (CTYPE = 2)) + (CTYPE = 2)
    If Unknown Then GoTo 950
    Call GetBDEvents(BD$, BPL$, DD$, DPL$, AGE$)
    LSet xDTXT$ = RFLBL$(5): Z = InStr(xDTXT$, "  ")
    Mid$(xDTXT$, Z, 12) = BD$
    If (CTYPE = 2) Then
       If (xNDNUM < 16) Or (xNDNUM > 32) Then
          Mid$(xDTXT$, Z + 13, 255) = RFLBL$(6): Z = InStr(Z + 13, xDTXT$, "  ")
          Mid$(xDTXT$, Z, 255) = BPL$
          Else
          Mid$(xDTXT$, Z + 15, 255) = RFLBL$(11): Z = InStr(Z + 15, xDTXT$, "  ")
          Mid$(xDTXT$, Z, 255) = DD$
       End If
    End If
    Mid$(LTAB$(Y + 1), X, DLTH) = xDTXT$
    If CTYPE = 1 Then
       LSet xDTXT$ = RFLBL$(6): Z = InStr(xDTXT$, "  ")
       Mid$(xDTXT$, Z, 255) = BPL$
       Mid$(LTAB$(Y + 2), X, DLTH) = xDTXT$
    End If
    If DUPID Then GoSub 980: GoTo 960
    If (xNDNUM > 15) And (xNDNUM < 32) Then GoTo 960
    LSet xDTXT$ = RFLBL$(11): Z = InStr(xDTXT$, "  ")
    Mid$(xDTXT$, Z, 255) = DD$
    If (CTYPE = 1) Then
       If (DD$ <> Space$(12)) And (AGE$ <> " ??") Then
          Z = InStr(Z + 11, xDTXT$, "  ")
          Mid$(xDTXT$, Z + 1, 255) = xATAGE$ + " " + LTrim$(AGE$)
       End If
       Else
       Mid$(xDTXT$, Z + 13, 255) = RFLBL$(12): Z = InStr(Z + 13, xDTXT$, "  ")
       Mid$(xDTXT$, Z, 255) = DPL$
    End If
    Mid$(LTAB$(w), X, DLTH) = xDTXT$
    If CTYPE = 1 Then
       LSet xDTXT$ = RFLBL$(12): Z = InStr(xDTXT$, "  ")
       Mid$(xDTXT$, Z, 255) = DPL$
       Mid$(LTAB$(w + 1), X, DLTH) = xDTXT$
    End If

950 ' Determine Marriage Info
    If (S = 0) Then GoTo 960
    SPRID = FF1Rec.SPOUSE
    If xNDNUM > 1 Then
       ASPID = Abs(xMID(GLD - 1))
       If ASPID = 0 Then SPID = -1: GoTo 975
    End If
955 If SPRID = 0 Then SPID = -1: GoTo 975
    Call FF3GetRec(SPRID)
    If F1RID = FF3SPOUSE.SP1ID Then
       SPID = FF3SPOUSE.SP2ID: SPRID = FF3SPOUSE.SP1NXT
       Else
       SPID = FF3SPOUSE.SP1ID: SPRID = FF3SPOUSE.SP2NXT
    End If
    If (xNDNUM > 1) Then If (SPID <> ASPID) Then GoTo 955
    MLOC = FF3SPOUSE.MLOC: LSet MPLACE = ""
    If MLOC Then Call FF3GetRec(MLOC)
    LSet xDTXT$ = RFLBL$(8): Z = InStr(xDTXT$, "  ")
    Mid$(xDTXT$, Z, 255) = XRPTDate$(FF3SPOUSE.D1MD, FF3SPOUSE.D1Y, 1)
    If (CTYPE = 2) And (xNDNUM < 16) Then
       Mid$(xDTXT$, Z + 13, 255) = RFLBL$(9): Z = InStr(Z + 13, xDTXT$, "  ")
       Mid$(xDTXT$, Z, 255) = MPLACE
    End If
    Mid$(LTAB$(Y + 3 + (CTYPE = 2)), X, DLTH) = xDTXT$
    If (CTYPE = 1) Then
       LSet xDTXT$ = RFLBL$(9): Z = InStr(xDTXT$, "  ")
       Mid$(xDTXT$, Z, 255) = MPLACE
       Mid$(LTAB$(Y + 4), X, DLTH) = xDTXT$
    End If
    GoTo 975

960 'Create Link Entries on Right Side of Chart
    If (xNDNUM < 16) Or (xNDNUM > 31) Or ((CTYPE = 2) And DUPID) Then GoTo 975
    If CTYPE = 2 Then X = 124
    Mid$(LTAB$(Y), X + 1, 1) = ">": X = X + 2
    If (CTYPE = 1) And DUPID Then GoSub 980
    If PAGELNK Then
       If (CTYPE = 1) And DUPID Then
          Mid$(LTAB$(Y), X, 255) = Str$(PPAGE)
          Else
          LPGNUM = BAPGNUM(WFANCREC.ANCID)
          Mid$(LTAB$(Y), X, 255) = Str$(LPGNUM)
       End If
       Else
       If Not ((CTYPE = 1) And DUPID) Then
          CHAINID$ = XChartID$(BGL + 4, LNUM#)
       End If
       Mid$(LTAB$(Y), X, 255) = CHAINID$
    End If
    If (CTYPE = 1) And DUPID Then
       X = InStr(X, LTAB$(Y), "  "): Mid$(LTAB$(Y), X, 1) = "("
       If ShowLIN Then
          Mid$(LTAB$(Y), X + 1, 255) = LTrim$(Str$(PLNUM#))
          Else
          Mid$(LTAB$(Y), X + 1, 255) = LTrim$(Str$(PNDNUM))
       End If
       X = InStr(X, LTAB$(Y), "  "): Mid$(LTAB$(Y), X, 1) = ")"
    End If

975 Return

980 ' Print notice of duplicate ID
    While WFANCREC.PRVDUP <> 0
       Call RWRKGet("PTRREC", RC, WFANCREC.PRVDUP)
    Wend
    PGL = Abs(WFAGLREC.AGL)
    PLNUM# = WFANCREC.LINEAGE
    Z = PGL Mod 4: If (CTYPE = 2) And (Z = 0) Then Z = 4
    PBGL = PGL - Z
    While (Z > 0) And (WFANCREC.ANCCH <> 0)
       Call RWRKGet("PTRREC", RC, WFANCREC.ANCCH): Z = Z - 1
    Wend
    PBLNUM# = WFANCREC.LINEAGE
    PNDNUM = PLNUM# - (PBLNUM# - 1#) * (2 ^ (PGL - PBGL))
    If PAGELNK Then
       PPAGE = BAPGNUM(WFANCREC.ANCID)
       Else
       CHAINID$ = XChartID$(PBGL, PBLNUM#)
    End If
    If (xNDNUM < 16) Or ((CTYPE = 2) And (xNDNUM < 32)) Then
       w = Y + 3 + (CTYPE = 2) + (xNDNUM > 15)
       LSet xDTXT$ = "(" + xSAMEAS$: Z = InStr(xDTXT$, "  ")
       If ShowLIN Then
          Mid$(xDTXT$, Z, 255) = Str$(PLNUM#)
          Else
          Mid$(xDTXT$, Z, 255) = Str$(PNDNUM)
       End If
       Z = InStr(Z, xDTXT$, "  ") + 1
       If PAGELNK Then
          Mid$(xDTXT$, Z, 255) = xONPAGE$: Z = InStr(Z, xDTXT$, "  ")
          Mid$(xDTXT$, Z, 255) = Str$(PPAGE): Z = InStr(Z, xDTXT$, "  ")
          Else
          Mid$(xDTXT$, Z, 255) = xONCHART$: Z = InStr(Z, xDTXT$, "  ") + 1
          Mid$(xDTXT$, Z, 255) = CHAINID$: Z = InStr(Z, xDTXT$, "  ")
       End If
       Mid$(xDTXT$, Z, 1) = ")"
       Mid$(LTAB$(w), X, DLTH) = xDTXT$: LSet xDTXT$ = ""
       If (xNDNUM < 16) Or (xNDNUM > 32) Then
          K = (1 - (CTYPE = 1)) * (S + 1) - 1
          While K > 0
             Mid$(LTAB$(w + K), X, DLTH) = xDTXT$: K = K - 1
          Wend
       End If
    End If
    Call RWRKGet("PTRREC", RC, xAPTR&)
    Return

X1000: ' PRINT CHART
     If ANS$ = SNGLKEY$(13) Then
        If RREFRPT.BGNPAGE = 0 Then RREFRPT.BGNPAGE = PAGE.RPAGE
        PAGE.COUNT = PAGE.RPAGE
        GoTo 1049
     End If
     A = 0: DTIMEOUT = 0:
     HPAGE = 0: HOFFSET = PRT.FOFFSET: PAGE.BGNPAGE = PAGE.COUNT
     LSet SysVAR.MVAR = SysVAR.CHARTID
X1010:
     For I = 0 To 62
        LSet PLINE$ = LTAB$(I): Call PrintLINE
        If A = 27 Then I = 62
     Next I
     Call PageBREAK
     PRTAGAIN = (A <> 27) And ((PRT.FOFFSET + PRT.FWIDTH) < PAGE.WIDTH)
     H1 = SysVAR.NVAR1
     If PRTAGAIN And (Not (ANS$ = SNGLKEY$(4))) Then
        SysVAR.NVAR1 = PRT.FOFFSET + PRT.FWIDTH
        Call PutMSG("XXXXM007")
        PRTAGAIN = Not (CA$ = SNGLKEY$(2))
     End If
     If PRTAGAIN Then
        SysVAR.NVAR1 = H1
        HPAGE = PAGE.COUNT
        PAGE.COUNT = PAGE.BGNPAGE
        PRT.FOFFSET = PRT.FOFFSET + PRT.FWIDTH
        GoTo X1010
     End If
     If HPAGE Then PAGE.COUNT = HPAGE: PRT.FOFFSET = HOFFSET
1049 Return

1100 ' END OF REPORT
     PRT.AGAIN = 0
     Call RptCLOSE
     'If CLSCRN Then Cls
     GoTo 60

End Sub

Sub AMAPReport()
    GoTo 2000

X60: ' Error Exit
     Call FamCLOSE
     Exit Sub

2000 'Print ALLGEN Report
     If RptOPTION.BLDREF Then
        PAGE.RPAGE = 0: PAGE.RLINE = 0
        RREFXData$ = "": PAGE.xType = 1
     End If
     NFMT$ = Space$(80): xDTXT$ = Space$(50): mline$ = Space$(255)
     BD$ = Space$(12): DD$ = BD$: BPL$ = Space$(41): DPL$ = BPL$: AGE$ = "   "
     Call RptOPEN("AMAP"): If Not OKAY Then GoTo X60
     Call BldMapDEF
     Call RptOptGET("ALIN", X): ShowLIN = (X = 1): zALIN$ = ""
     Call RptOptGET("PRID", X): ShowID = (X = 1)
     NOSHOW = Not (ShowLIN Or ShowID)
     Call RptOptGET("PDTE", X): SHOWDATES = (X = 1)
     Call RptOptGET("PPLC", X): SHOWPLACE = (X = 1)
     Call RptOptGET("NDLN", X): NUMDATALN = X
     If NUMDATALN < 1 Then NUMDATALN = 1 Else If NUMDATALN > 3 Then NUMDATALN = 3
     Call RptOptGET("BLKL", X): DUBLSPACE = (X = 1)
     Call GetRVAR("SAMEAS", X): xSAMEAS$ = RVAR$(X)
     Call GetRVAR("ATAGE", X): xATAGE$ = RVAR$(X)
     Call FAMOpen("READ"): If Not OKAY Then GoTo X60
     Call RWRKOpen("READ", 1, 0): If Not OKAY Then GoTo X60
     PRDAGE = (RptOPTION.PRAGE > 0)
     MAXGL = Asc(WFHDR.AMAXLV): WFHDR.DMAXLV = Chr$(0)
     If MAXGL > 53 Then
        Call PutMSG("ACHTM007")
        If (A = 27) Or (CA$ = SNGLKEY$(2)) Then GoTo X60 Else MAXGL = 53
     End If
     Call PrintOPEN
     AMAXLV = MAXGL

2020 ' Print (Strip of) MAP
     Call RWRKGet("BASE", RC, PTR&): If RC = 9 Then GoTo X60
     SysVAR.SUBJECT = RTrim$(FMTNAME$(1))
     ONEMORE = 0: A = 0
     'If PRT.OPT = 1 Then
     '   LSet DCHAR$ = SCHAR$
     '   Else
        LSet DCHAR$ = PCHAR$
     'End If
     N$ = " ": v$ = MID$(DCHAR$, 3, 1): h$ = MID$(DCHAR$, 1, 1): xx$ = MID$(DCHAR$, 2, 1)
     M$ = MID$(DCHAR$, 4, 1): F$ = MID$(DCHAR$, 5, 1)
     ReDim LNUM#(AMAXLV + 1)
     ReDim glaptr&(AMAXLV + 1), GLAEND&(AMAXLV + 1), GLACH&(AMAXLV + 1)
     GL = 0: ANUM = 1: xLNUM# = 1
     glaptr&(0) = WFAGLREC.ANCPTR: LNUM#(0) = 1: LNUM#(AMAXLV + 1) = 0
     While WFAGLREC.NEXT <> 0
       MAXPTR& = WFAGLREC.NEXT - 1
       Get #5, WFAGLREC.NEXT + 1, WFAGLREC
       Get #5, WFAGLREC.ANCPTR + 1, WFANCREC
       GL = GL + 1: glaptr&(GL) = WFAGLREC.ANCPTR
       GLACH&(GL) = WFANCREC.ANCCH: GLAEND&(GL) = MAXPTR&
       LNUM#(GL) = WFANCREC.LINEAGE
     Wend
     If AMAXLV > 53 Then LNUM#(54) = 0: MAXPTR& = glaptr&(54) - 1
     GL = 0: DTIMEOUT = 0: LSet mline$ = " "

2050 ' Print Ancestor Line
     While (glaptr&(GL + 1) > 0) And (LNUM#(GL + 1) = 2 * LNUM#(GL))
        GL = GL + 1
     Wend
     If glaptr&(GL) = 0 Then GoTo X1100
     glofst = 1 + 3 * GL: xLNUM# = LNUM#(GL)
     Get #5, glaptr&(GL) + 1, WFANCREC
     AID = WFANCREC.ANCID: DUPID = (AID < 0): If DUPID Then AID = -AID
     Call FF1GetRec(AID)
     Mid$(mline$, glofst, 255) = Space$(255)
     RSTAT = WFANCREC.PSTAT And 3
     If (WFANCREC.PSTAT And 48) > 0 Then
       If (WFANCREC.PSTAT And 32) > 0 Then LSet N$ = F$ Else LSet N$ = M$
       Else
       If FNODD(xLNUM#) = 0 Then LSet N$ = M$ Else LSet N$ = F$
     End If
     If GL = 0 Then LSet N$ = " "
     Mid$(mline$, glofst, 1) = N$
     Mid$(mline$, glofst + 1, 1) = h$
     Mid$(mline$, glofst + 2, 1) = h$
     Mid$(mline$, glofst + 3, 1) = xx$
     If (FF1Rec.FID = 0) And (FF1Rec.MID = 0) Then
        Mid$(mline$, glofst + 3, 1) = h$
     End If
     If ShowLIN Then Mid$(mline$, glofst + 4, 255) = FNDS$(xLNUM#)
     DOFST = InStr(glofst, mline$, "  ")
     If (RSTAT And 1) Then
        Mid$(mline$, DOFST, 1) = "*"
        Else
        If (RSTAT And 2) Then Mid$(mline$, DOFST, 1) = "?"
     End If
     OFST = DOFST + 1:
     If ShowID Then
        Mid$(mline$, OFST, 255) = "(" + LTrim$(Str$(AID)) + ")"
        OFST = InStr(OFST, mline$, "  ")
     End If
     Mid$(mline$, OFST, 255) = FMTNAME$(0): OFST = InStr(OFST, mline$, "   ")
     If DUPID Then
        Call RWRKGet("PTRREC", RC, WFANCREC.PRVDUP)
        Mid$(mline$, OFST + 1, 255) = "(" + xSAMEAS$
        OFST = InStr(OFST + 1, mline$, "  ")
        Mid$(mline$, OFST, 255) = FNDS$(WFANCREC.LINEAGE)
        OFST = InStr(OFST + 1, mline$, "  ")
        Mid$(mline$, OFST, 1) = ")": OFST = OFST + 1
     End If
     If RptOPTION.BLDREF Then
        RREFREC.AGL = Chr$(GL)
        PAGE.XID = AID: Call RREFAdd
     End If
     If DUPID Then GoTo 2130
     If NUMDATALN > 1 Then GoSub 2160: If A = 27 Then GoTo 2145 Else OFST = DOFST
2110 ' put data on line
     If Not (SHOWDATES Or SHOWPLACE) Then GoTo 2130
     Call GetBDEvents(BD$, BPL$, DD$, DPL$, AGE$)
     If SHOWDATES Then
        If BD$ <> Space$(12) Then
           Mid$(mline$, OFST, 255) = "," + RFLBL$(5)
           OFST = InStr(OFST, mline$, "  ")
           If MID$(mline$, OFST - 1, 1) <> "." Then OFST = OFST + 1
           Mid$(mline$, OFST, 255) = LTrim$(BD$)
           OFST = InStr(OFST + 1, mline$, "   ")
        End If
     End If
     If SHOWPLACE Then
        If Asc(BPL$) <> 32 Then
           If (BD$ = Space$(12)) Or (Not SHOWDATES) Then
              Mid$(mline$, OFST, 255) = "," + RFLBL$(5)
              Else
              Mid$(mline$, OFST, 1) = ","
           End If
           OFST = InStr(OFST, mline$, "  ")
           If MID$(mline$, OFST - 1, 1) <> "." Then OFST = OFST + 1
           Mid$(mline$, OFST, 255) = BPL$
           OFST = InStr(OFST, mline$, "   ")
        End If
     End If
     If (NUMDATALN > 2) And (OFST > DOFST) Then
        Mid$(mline$, DOFST, 1) = " ": GoSub 2160: OFST = DOFST
        If A = 27 Then GoTo X1100
     End If
     If SHOWDATES Then
        If DD$ <> Space$(12) Then
           Mid$(mline$, OFST, 255) = "," + RFLBL$(11)
           OFST = InStr(OFST, mline$, "  ")
           If MID$(mline$, OFST - 1, 1) <> "." Then OFST = OFST + 1
           Mid$(mline$, OFST, 255) = LTrim$(DD$)
           OFST = InStr(OFST, mline$, "   ")
        End If
     End If
     If SHOWPLACE Then
        If Asc(DPL$) <> 32 Then
           If (DD$ = Space$(12)) Or (Not SHOWDATES) Then
              Mid$(mline$, OFST, 255) = "," + RFLBL$(11)
              Else
              Mid$(mline$, OFST, 1) = ","
           End If
           OFST = InStr(OFST, mline$, "  ")
           If MID$(mline$, OFST - 1, 1) <> "." Then OFST = OFST + 1
           Mid$(mline$, OFST, 255) = DPL$
           OFST = InStr(OFST, mline$, "   ")
        End If
     End If
     If PRDAGE Then
        If (SHOWDATES And (DD$ <> Space$(12))) Or (SHOWPLACE And (Asc(DPL$) <> 32)) Then
           If (AGE$ <> "   ") And (AGE$ <> " ??") Then
              Mid$(mline$, OFST + 1, 255) = xATAGE$
              OFST = InStr(OFST, mline$, "  ")
              Mid$(mline$, OFST + 1, 255) = LTrim$(AGE$)
              OFST = InStr(OFST, mline$, "   ")
           End If
        End If
     End If
2130 If (NUMDATALN = 1) Then
        GoSub 2160
        Else
        If (OFST > DOFST) Then Mid$(mline$, DOFST, 1) = " ": GoSub 2160
     End If
     If A = 27 Then GoTo X1100
     If DUBLSPACE Then GoSub 2160: If A = 27 Then GoTo X1100
     glaptr&(GL) = glaptr&(GL) + Len(WFANCREC)
     If glaptr&(GL) > GLAEND&(GL) Then
        glaptr&(GL) = 0
        If glaptr&(GL + 1) > 0 Then
           Mid$(mline$, glofst + 3, 1) = v$
           GL = GL + 1: GoTo 2145
           Else
           While (GL > 0) And (glaptr&(GL) = 0)
              GL = GL - 1: xLNUM# = (xLNUM# - FNODD(xLNUM#)) / 2
           Wend
        End If
        Else
        Get #5, glaptr&(GL) + 1, WFANCREC
        LNUM#(GL) = WFANCREC.LINEAGE
        GLACH&(GL) = WFANCREC.ANCCH
        If LNUM#(GL + 1) = ((2 * xLNUM#) + 1#) Then
           GL = GL + 1: Mid$(mline$, glofst + 3, 1) = v$: GoTo 2145
        End If
     End If
     While (GL > 0) And (xLNUM# < LNUM#(GL))
        GL = GL - 1: xLNUM# = (xLNUM# - FNODD(xLNUM#)) / 2
     Wend
     If GL < 0 Then
        SysVAR.MVAR = Str$(glaptr&(GL))
        Call PutMSG("ERROR...GL<0 at GLAPTR=[MVAR]...")
        Call GetRESP: Call PutMSG("")
        GoTo X1100
     End If
2145 If (glaptr&(GL) = 0) Or (A = 27) Then GoTo X1100
     GoTo 2050

2160 ' Print Line of Ancestor Map
     LSet PLINE$ = mline$: Call PrintLINE
     If A = 27 Then GoTo 2167
     Mid$(mline$, glofst, 255) = Space$(255)
     If (WFANCREC.PSTAT And 48) > 0 Then
       If (WFANCREC.PSTAT And 32) = 0 Then Mid$(mline$, glofst, 1) = v$
       Else
       If FNODD(xLNUM#) = 0 Then Mid$(mline$, glofst, 1) = v$
     End If
     If Not (DUBLSPACE Or (NUMDATALN > 1)) Then GoTo 2167
     If Not (DUPID Or (FF1Rec.MID = 0)) Then Mid$(mline$, glofst + 3, 1) = v$
2167 Return

X1100: ' END OF REPORT
     Call RptCLOSE
     If PRT.AGAIN Then RptOPTION.BLDREF = 0: GoTo 2020
     GoTo X60

End Sub

Sub BldChartDEF(CTYPE)
      RFCNT = 12
      ReDim DFNUM(RFCNT), DFPP(RFCNT)
      ReDim RFNUM(RFCNT), RFSEP(RFCNT), RFLTH(RFCNT), RFLBL$(RFCNT)
      Call RptOptGET("MINM", MINMLTH)
      If (Not OKAY) Or (MINMLTH = 0) Then MINMLTH = 35
      If (PAGE.BWIDTH = 0) Then
         BWIDTH = PRT.FWIDTH - PAGE.IMARGIN - PAGE.OMARGIN
         Else
         BWIDTH = PAGE.BWIDTH
      End If

59510 ' Create Report Definition Tables
      X = 1: RFNAME$ = "07011": GoSub 59505    'LNUM
      X = 2: RFNAME$ = "01020": GoSub 59505    'FULL NAME
      X = 3: RFNAME$ = "01010": GoSub 59505    'RID
      X = 4: RFNAME$ = "01016": GoSub 59505    'AGE
      X = 5: RFNAME$ = "01030": GoSub 59505    'BIRTH DATE
      X = 6: RFNAME$ = "01032": GoSub 59505    'BIRTH PLACE
      X = 7: RFNAME$ = "02011": GoSub 59505    'SPOUSE
      X = 8: RFNAME$ = "02020": GoSub 59505    'MARRIAGE DATE
      X = 9: RFNAME$ = "02021": GoSub 59505    'MARRIAGE PLACE
      X = 10: RFNAME$ = "02015": GoSub 59505   'Last ANNIVERSARY
      X = 11: RFNAME$ = "01040": GoSub 59505   'DEATH DATE
      X = 12: RFNAME$ = "01042": GoSub 59505   'DEATH PLACE
      GoTo 59520

59505 ' Process Report Field
      Call GetRFLD(RFNAME$, FX)
      If RFSEP(X) = 0 Then RFSEP(X) = MINGAP
      If FOUND Then
         RDFFld = RFLDS(FX)
         If RFLTH(X) = 0 Then RFLTH(X) = Asc(RDFFld.RLTH)
         ALTH = Asc(RDFFld.ALTH): LLTH = Asc(RDFFld.LLTH)
         HLTH = Asc(RDFFld.HLTH)
         If (CTYPE = 1) Or (X = 3) Then
            RFLBL$(X) = MID$(RFLD$(FX), ALTH + 1, LLTH)
            Else
            RFLBL$(X) = MID$(RFLD$(FX), 1, ALTH)
         End If
      End If
      Return

59520 ' Check Report Width
      PWIDTH = PRT.FWIDTH
      RPWIDTH = 132 + PAGE.IMARGIN + PAGE.OMARGIN
      If PWIDTH < RPWIDTH Then
         SysVAR.NVAR1 = PWIDTH: SysVAR.NVAR2 = RPWIDTH
         Call PutMSG("XXXXM009")
         If CA$ <> SNGLKEY$(1) Then BWIDTH = RWIDTH
      End If

59560 ' Return to Calling Program
      PAGE.BWIDTH = BWIDTH
      OKAY = -1
End Sub

Sub BldMapDEF()
      RFCNT = 12: CTYPE = 2
      ReDim DFNUM(RFCNT), DFPP(RFCNT)
      ReDim RFNUM(RFCNT), RFSEP(RFCNT), RFLTH(RFCNT), RFLBL$(RFCNT)
      Call RptOptGET("ALIN", X): ShowLIN = (X = 1)
      Call RptOptGET("MINM", MINMLTH)
      If (Not OKAY) Or (MINMLTH = 0) Then MINMLTH = 35
      Call RptOptGET("NDLN", NUMDATALN)
      If NUMDATALN < 1 Then NUMDATALN = 1 Else If NUMDATALN > 3 Then NUMDATALN = 3

      OKAY = 0
      PAGE.DFCNT = 0: IDFX = 0: BFX = 0: MFX = 0: DFX = 0
      AMAXLV = Asc(WFHDR.AMAXLV): If AMAXLV > 53 Then AMAXLV = 53
      If (PAGE.BWIDTH = 0) Then
         BWIDTH = PRT.FWIDTH - PAGE.IMARGIN - PAGE.OMARGIN
         Else
         BWIDTH = PAGE.BWIDTH
      End If

      ' Create Report Definition Tables
      X = 1: RFNAME$ = "07011": GoSub M59505    'LNUM
      X = 2: RFNAME$ = "01020": GoSub M59505    'FULL NAME
      X = 3: RFNAME$ = "01010": GoSub M59505    'RID
      X = 4: RFNAME$ = "01016": GoSub M59505    'AGE
      X = 5: RFNAME$ = "01030": GoSub M59505    'BIRTH DATE
      X = 6: RFNAME$ = "01032": GoSub M59505    'BIRTH PLACE
      X = 7: RFNAME$ = "02011": GoSub M59505    'SPOUSE
      X = 8: RFNAME$ = "02020": GoSub M59505    'MARRIAGE DATE
      X = 9: RFNAME$ = "02021": GoSub M59505    'MARRIAGE PLACE
      X = 10: RFNAME$ = "02015": GoSub M59505   'Last ANNIVERSARY
      X = 11: RFNAME$ = "01040": GoSub M59505   'DEATH DATE
      X = 12: RFNAME$ = "01042": GoSub M59505   'DEATH PLACE
      GoTo M59520

M59505: ' Process Report Field
      Call GetRFLD(RFNAME$, FX)
      If RFSEP(X) = 0 Then RFSEP(X) = MINGAP
      If FOUND Then
         RDFFld = RFLDS(FX)
         If RFLTH(X) = 0 Then RFLTH(X) = Asc(RDFFld.RLTH)
         ALTH = Asc(RDFFld.ALTH): LLTH = Asc(RDFFld.LLTH)
         HLTH = Asc(RDFFld.HLTH)
         If CTYPE = 1 Then
            RFLBL$(X) = MID$(RFLD$(FX), ALTH + 1, LLTH)
            Else
            RFLBL$(X) = MID$(RFLD$(FX), 1, ALTH)
         End If
      End If
      Return

M59520: ' Check Report Width
      RWIDTH = 0: FX = 0
      RWIDTH = 3 * (AMAXLV + 1)
      If ShowLIN Then RWIDTH = RWIDTH + Len(Str$(2 ^ AMAXLV))
      Select Case NUMDATALN
         Case 3: RWIDTH = RWIDTH + 50
         Case 2: RWIDTH = RWIDTH + 75
         Case 1: RWIDTH = RWIDTH + MINMLTH + 50 + Len(RFLBL$(5)) + Len(RFLBL$(11))
      End Select
      RWIDTH = RWIDTH - 1
      If RWIDTH <= BWIDTH Then GoTo M59560

      ' Final Check
      PWIDTH = BWIDTH + PAGE.IMARGIN + PAGE.OMARGIN
      RPWIDTH = RWIDTH + PAGE.IMARGIN + PAGE.OMARGIN
      If PWIDTH < RPWIDTH Then
         SysVAR.NVAR1 = PWIDTH: SysVAR.NVAR2 = RPWIDTH
         Call PutMSG("XXXXM009")
         If CA$ = SNGLKEY$(1) Then BWIDTH = RWIDTH
      End If

M59560: ' Return to Calling Program
      PAGE.BWIDTH = BWIDTH
      OKAY = -1
End Sub

Rem $STATIC
Function XChartID$(BGL, BLNUM#)
    CH# = BLNUM# - (2 ^ BGL) + 1#
    XChartID$ = LTrim$(Str$(BGL)) + "-" + LTrim$(FNDS$(CH#))
End Function

