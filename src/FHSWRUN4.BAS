Attribute VB_Name = "dllFHSWRun4"
DefInt A-Z
Sub FldUpdate(xKeyCode, xShift, xFT, xSS, xSL, XVALUE As String)
' xSS = Field SelStart value
' xSL = Field SelLength value
' xFT = Field Type: 1=Alpha, 2=Numeric, 3=Date, 4=Time, 5=upcase, 6=Phone
' xVALUE$ = Current Field Value
If (xKeyCode = 37) Or (xKeyCode = 39) Then Exit Sub
If (xKeyCode = 38) Or (xKeyCode = 40) Then GoTo NullKey
If (xKeyCode = 46) Then
   If ((xFT > 2) And (xFT < 5)) Or (xFT = 6) Then GoTo NullKey
   Exit Sub
End If
If (xKeyCode = 8) Then
   If xSS > 0 Then xSS = xSS - 1
   xKeyCode = 0
   Exit Sub
End If
If xKeyCode < 32 Then Exit Sub
' If ((xKeyCode > 96) And (xKeyCode < 106)) Then xKeyCode = xKeyCode - 48
If ((xKeyCode > 64) And (xKeyCode < 91)) And ((xShift And 1) = 0) And (xFT <> 5) Then
   xKeyCode = xKeyCode + 32
End If
If (xFT = 5) And ((xKeyCode < 65) Or (xKeyCode > 90)) And (xKeyCode <> 32) Then
   xKeyCode = 0: Exit Sub
End If
Select Case xFT
   Case 3: FldMask$ = DateMask
   Case 4: FldMask$ = "99:99"
   Case 6: FldMask$ = "(999) 999-9999"
   Case Else: FldMask$ = ""
End Select
If ((xFT = 2) Or ((xFT = 3) And (DFM = 6)) Or (xFT = 6)) Then
   If (xKeyCode < Asc("0")) Or (xKeyCode > Asc("9")) Then
      GoTo NullKey
   End If
End If
If xKeyCode <> 32 Then
   If (xKeyCode < 48) Then GoTo NullKey
   If (xKeyCode > 57) And (xKeyCode < 65) Then GoTo NullKey
   If (xKeyCode > 90) And (xKeyCode < 97) Then GoTo NullKey
   If (xKeyCode > 122) And (xKeyCode < 160) Then GoTo NullKey
End If
If ((xFT > 2) And (xFT < 5)) Or (xFT = 6) Then
   'If Len(XVALUE) < Len(DateMask) Then
   '   XVALUE = XDATE
   'End If
   If xSS = Len(FldMask$) Then GoTo NullKey
   If (xKeyCode = 32) Then
      xKeyCode = 0
      Else
      If (xKeyCode > 96) And (xKeyCode < 123) Then xKeyCode = xKeyCode - 32
      If (xSS < Len(FldMask$)) Then
         dma = Asc(MID$(FldMask$, xSS + 1, 1))
         If (dma = Asc("M")) And (BMO > 0) Then
            If (xKeyCode < 65) Or (xKeyCode > 90) Then GoTo NullKey
            Else
            If ((dma < 48) Or ((dma > 57) And (dma < 65)) Or ((dma > 90) And (dma < 160))) Then GoTo NullKey
            If (xKeyCode < Asc("0")) Or (xKeyCode > Asc("9")) Then GoTo NullKey
         End If
         Mid$(XVALUE, xSS + 1, 1) = Chr$(xKeyCode)
         If (xFT = 3) And (BMO > 0) And (((xSS - BMO) > -2) And ((xSS - BMO) < 2)) Then
            MCHK$ = MID$(XVALUE, BMO, xSS - BMO + 2)
            FOUND = 0: X = 1
            While (X < 13) And (Not FOUND)
               FOUND = (MCHK$ = MID$(mon_table(X), 1, xSS - BMO + 2))
               If Not FOUND Then X = X + 1
            Wend
            If FOUND Then
               Mid$(XVALUE, BMO, 3) = mon_table(X)
               Else
               GoTo NullKey
            End If
         End If
      End If
   End If
   xSS = xSS + 1
   If (xSS < Len(FldMask$)) Then
      dma = Asc(MID$(FldMask$, xSS + 1, 1))
      While ((dma < 48) Or ((dma > 57) And (dma < 65)) Or ((dma > 90) And (dma < 160)))
         xSS = xSS + 1
         If xSS < Len(FldMask$) Then
            dma = Asc(MID$(FldMask$, xSS + 1, 1))
            Else
            dma = 48
         End If
      Wend
   End If
   GoTo NullKey
End If
XVALUE = Left$(XVALUE, xSS) + Chr$(xKeyCode) + MID$(XVALUE, xSS + xSL + 1, 255)
Mid$(XVALUE, xSS + 1, 1) = Chr$(xKeyCode)
xSS = xSS + 1: xSL = 0
Exit Sub
NullKey: xKeyCode = 0
End Sub

Sub SetCalendarType(CTYPE, cDFM)
   CalType = CTYPE: DFM = cDFM
   If DFM < 0 Then DFM = Val(CONFIG.DFM) - 1
   minADAT = 3 + 2 * (CTYPE > 0)
   maxADAT = 3652061 - 73 * (CTYPE > 0)
   WkLth = 7
   ' Create Tables for Date Calculations
   '    (moved from FHSWDCLC on 07/15/99)
   For X = 0 To 12
       mon_table(X) = UCX$(MID$(RMONTB$(X + 1), 3, 3))
   Next X
   ReDim day_table(WkLth)
   day_table(1) = "Sunday": day_table(2) = "Monday": day_table(3) = "Tuesday"
   day_table(4) = "Wednesday": day_table(5) = "Thursday": day_table(6) = "Friday"
   day_table(7) = "Saturday"
   days_table(1) = 31: days_table(2) = 28: days_table(3) = 31
   days_table(4) = 30: days_table(5) = 31: days_table(6) = 30
   days_table(7) = 31: days_table(8) = 31: days_table(9) = 30
   days_table(10) = 31: days_table(11) = 30: days_table(12) = 31
   DateMasks(0) = "MM-DD-YYYY": DateMasks(1) = "DD.MM.YYYY"
   DateMasks(2) = "YYYY.MM.DD": DateMasks(3) = "YYYY DDD"
   DateMasks(4) = "MMM DD,YYYY": DateMasks(5) = "DD MMM YYYY"
   DateMasks(6) = "DDDDDDD"
   DateMask = DateMasks(DFM)
   Select Case DFM
      Case 4: BMO = 1
      Case 5: BMO = 4
      Case Else: BMO = 0
   End Select

End Sub

Sub adat2yyyyddd()
   If DateW.adat < minADAT Then
      DateW.adat = minADAT
      Else
      If DateW.adat > maxADAT Then
         DateW.adat = maxADAT
      End If
   End If
   DateW.mm = 1: DateW.DD = 1
   DateW.ddd = 1: adat# = DateW.adat
   DateW.yyyy = (DateW.adat \ 366) + 1
   Call yyyyddd2adat
   DateW.ddd = adat# - DateW.adat + 1
   Call ddd2mmdd
   Call yyyyddd2adat
   If DateW.adat <> adat# Then
      SysVAR.MVAR = DateW.adat - adat#
      Call PutMSG("DateW.adat - adat=[MVAR]...")
      Call GetRESP
      Call ClearMSG
   End If
End Sub
Sub mmdd2ddd()
' First Normalize Date
LYCHK:
    Call ChkLY
    While DateW.mm > 12
       DateW.yyyy = DateW.yyyy + 1
       DateW.mm = DateW.mm - 12
       GoTo LYCHK
    Wend
    While DateW.DD > days_table(DateW.mm)
       DateW.DD = DateW.DD - days_table(DateW.mm)
       DateW.mm = DateW.mm + 1
       If DateW.mm > 12 Then GoTo LYCHK
    Wend
' Compute Julian Day
    DateW.ddd = DateW.DD
    For X = 0 To DateW.mm - 1
       DateW.ddd = DateW.ddd + days_table(X)
    Next X
End Sub

Sub yyyyddd2adat()
    Call ChkLY
    ' n1 = number of leap years in prior centuries
    If CalType = 0 Then
       n1 = (24 * DateW.CC + DateW.CC \ 4) + 2
       Else
       n1 = 25 * DateW.CC
    End If
    ' n2 = number of leap years so far this century
    n2 = DateW.yy \ 4 + DateW.ly
    ' compute number of non-leap days before today
    DateW.adat = DateW.yyyy - 1
    DateW.adat = DateW.adat * 365 + DateW.ddd
    DateW.adat = n1 + n2 + DateW.adat
    If (DateW.adat < 0) Then
       DateW.yyyy = DateW.yyyy - 1: DateW.adat = DateW.adat + 365
    End If
    DateW.wday = ((DateW.adat - 2) Mod WkLth) + 1
    While DateW.wday < 0
       DateW.wday = DateW.wday + WkLth
    Wend
End Sub
Function FormattedDate() As String
   XDATE = DateMasks(DFM)
   Select Case DFM
      Case 0: ' MM-DD-YYYY
              Mid$(XDATE, 1, 2) = ZFILL$(DateW.mm, 2)
              Mid$(XDATE, 4, 2) = ZFILL$(DateW.DD, 2)
              Mid$(XDATE, 7, 4) = ZFILL$(DateW.yyyy, 4)
      Case 1: ' DD.MM.YYYY
              Mid$(XDATE, 1, 2) = ZFILL$(DateW.DD, 2)
              Mid$(XDATE, 4, 2) = ZFILL$(DateW.mm, 2)
              Mid$(XDATE, 7, 4) = ZFILL$(DateW.yyyy, 4)
      Case 2: ' YYYY.MM.DD
              Mid$(XDATE, 1, 4) = ZFILL$(DateW.yyyy, 4)
              Mid$(XDATE, 6, 2) = ZFILL$(DateW.mm, 2)
              Mid$(XDATE, 9, 2) = ZFILL$(DateW.DD, 2)
      Case 3: ' YYYY.DDD
              Mid$(XDATE, 1, 4) = ZFILL$(DateW.yyyy, 4)
              Mid$(XDATE, 6, 3) = ZFILL$(DateW.ddd, 3)
      Case 4: ' MMM DD,YYYY
              Mid$(XDATE, 1, 3) = mon_table(DateW.mm)
              Mid$(XDATE, 5, 2) = ZFILL$(DateW.DD, 2)
              Mid$(XDATE, 8, 4) = ZFILL$(DateW.yyyy, 4)
      Case 5: ' DD MMM YYYY
              Mid$(XDATE, 1, 2) = ZFILL$(DateW.DD, 2)
              Mid$(XDATE, 4, 3) = mon_table(DateW.mm)
              Mid$(XDATE, 8, 4) = ZFILL$(DateW.yyyy, 4)
      Case 6: ' DDDDDD - Absolute Date
              LSet XDATE = DateW.adat
    End Select
    FormattedDate = RTrim$(XDATE)
End Function


Sub UnformatDate(xd As String, dfmt As Integer)
  X = dfmt
  If X < 0 Then X = DFM
  Select Case X
      Case 0: ' MM-DD-YYYY
              DateW.mm = Val(MID$(xd, 1, 2))
              DateW.DD = Val(MID$(xd, 4, 2))
              DateW.yyyy = Val(MID$(xd, 7, 4))
      Case 1: ' DD.MM.YYYY
              DateW.DD = Val(MID$(xd, 1, 2))
              DateW.mm = Val(MID$(xd, 4, 2))
              DateW.yyyy = Val(MID$(xd, 7, 4))
      Case 2: ' YYYY.MM.DD
              DateW.yyyy = Val(MID$(xd, 1, 4))
              DateW.mm = Val(MID$(xd, 6, 2))
              DateW.DD = Val(MID$(xd, 9, 2))
      Case 3: ' YYYY.DDD
              DateW.yyyy = Val(MID$(xd, 1, 4))
              DateW.ddd = Val(MID$(xd, 6, 3))
              Call ddd2mmdd
      Case 4: ' MMM DD,YYYY
              xxx$ = MID$(xd, 1, 3)
              N = 0: FOUND = 0
              While (N < 12) And (Not FOUND)
                 N = N + 1
                 FOUND = (UCase$(xxx$)) = mon_table(N)
              Wend
              If FOUND Then DateW.mm = N Else N = 0
              DateW.DD = Val(MID$(xd, 5, 2))
              DateW.yyyy = Val(MID$(xd, 8, 4))
      Case 5: ' DD MMM YYYY
              DateW.DD = Val(MID$(xd, 1, 2))
              xxx$ = MID$(xd, 4, 3)
              N = 1: FOUND = 0
              While (N < 13) And (Not FOUND)
                 N = N + 1
                 FOUND = (UCase$(xxx$)) = UCase$(MID$(RMONTB$(N), 3, 3))
              Wend
              If FOUND Then DateW.mm = N - 1 Else N = 1
              DateW.yyyy = Val(MID$(xd, 8, 4))
      Case 6: ' DDDDDD - Absolute Date
              DateW.adat = Val(xd)
              Call adat2yyyyddd
              Call ddd2mmdd
    End Select
    Call mmdd2ddd
    Call yyyyddd2adat
End Sub

Sub ChkLY()
    DateW.CC = DateW.yyyy \ 100
    DateW.yy = DateW.yyyy Mod 100
    If (DateW.yy <> 0) Then
       If ((DateW.yy Mod 4) = 0) Then
          DateW.ly = -1
          Else
          DateW.ly = 0
       End If
       Else
       If ((DateW.CC Mod 4) = 0) Then
          DateW.ly = -1
          Else
          If (CalType = 0) Then
             DateW.ly = 0
             Else
             DateW.ly = -1
          End If
      End If
    End If
    If DateW.ly Then days_table(2) = 29 Else days_table(2) = 28
End Sub

Sub ddd2mmdd()
    'While DateW.mm > 12
    '     DateW.yyyy = DateW.yyyy + 1
    '     DateW.mm = DateW.mm - 12
    'Wend
    DateW.mm = 0
    Call ChkLY
    DateW.DD = DateW.ddd
    While DateW.DD > days_table(DateW.mm)
         DateW.DD = DateW.DD - days_table(DateW.mm)
         DateW.mm = DateW.mm + 1
         If DateW.mm > 12 Then
            DateW.yyyy = DateW.yyyy + 1
            DateW.ddd = DateW.ddd - 365 + DateW.ly
            DateW.mm = 1
            Call ChkLY
         End If
    Wend
End Sub
Sub UpdtPLACE(XLOC, SRTYPE, SRNO)

670 ' Update PLACE Information
    Select Case SRTYPE
       Case 1: X1PL$ = RTrim$(BPLACE): Y1PL$ = RTrim$(DPLACE)
               X2PL$ = RTrim$(OBPLACE): Y2PL$ = RTrim$(ODPLACE)
               LSet BPLACE = OBPLACE: LSet DPLACE = ODPLACE
       Case 3: X1PL$ = RTrim$(MPLACE): Y1PL$ = RTrim$(DVPLACE)
               X2PL$ = RTrim$(OMPLACE): Y2PL$ = RTrim$(ODVPLACE)
               LSet MPLACE = OMPLACE: LSet DVPLACE = ODVPLACE
    End Select
    If (X1PL$ = X2PL$) And (Y1PL$ = Y2PL$) Then Exit Sub
    If XLOC <> 0 Then
       Call FF3GetRec(XLOC)
       If FF3PFX.RTYPE = Chr$(12) Then
          If FF3PLACE2.PNEXT <> 0 Then Call FF3DelRec(FF3PLACE2.PNEXT)
       End If
       Call FF3DelRec(XLOC)
    End If
    XLOC = 0: FF3PFX.RNO = 0
    PL1 = Len(X1PL$): PL2 = Len(Y1PL$)
    If (PL1 > 0) Or (PL2 > 0) Then
       If ((PL1 > 22) Or (PL2 > 22)) And ENV.LPNSupport Then
          FF3PFX.RTYPE = Chr$(12)
          If PL2 > 0 Then
             Call FF3GetRec(0)
             If OKAY Then
                FF3PLACE2.SRTYPE = Chr$(SRTYPE): FF3PLACE2.SRNO = SRNO
                FF3PLACE2.PTYPE = Chr$(SRTYPE + 1)
                LSet FF3PLACE2.PLACE = Y1PL$
                Call FF3PutRec
                If OKAY Then XLOC = FF3PFX.RNO: FF3PFX.RNO = 0
             End If
          End If
          If PL1 > 0 Then
             Call FF3GetRec(0)
             If OKAY Then
                FF3PLACE2.SRTYPE = Chr$(SRTYPE): FF3PLACE2.SRNO = SRNO
                FF3PLACE2.PTYPE = Chr$(SRTYPE)
                FF3PLACE2.PNEXT = XLOC
                LSet FF3PLACE2.PLACE = X1PL$
                Call FF3PutRec
             End If
          End If
          Else
          FF3PFX.RTYPE = Chr$(4)
          Call FF3GetRec(0)
          If OKAY Then
             FF3PLACE.SRTYPE = Chr$(SRTYPE): FF3PLACE.SRNO = SRNO
             LSet FF3PLACE.PLACE1 = X1PL$: LSet FF3PLACE.PLACE2 = Y1PL$
             LSet X1PL$ = FF3PLACE.PLACE1: LSet Y1PL$ = FF3PLACE.PLACE2
             Call FF3PutRec
          End If
       End If
       If OKAY Then
          XLOC = FF3PFX.RNO
          Select Case SRTYPE
             Case 1: LSet BPLACE = X1PL$: LSet DPLACE = Y1PL$
             Case 3: LSet MPLACE = X1PL$: LSet DVPLACE = Y1PL$
          End Select
       End If
    End If
End Sub


Sub NodeWorkPUT(RTP$, RNUM)
3200 ' Write record to Node Work file
    ANCSEARCH = ((Asc(NodeHDR.WFOPT) And 16) > 0)
    Select Case RTP$
       Case "HDR": Put #5, 1, NodeHDR
       Case "GLREC": RBA& = Len(NodeHDR)
                     RBA& = RBA& + RNUM * Len(NodeGLREC) + 1
                     Put #5, RBA&, NodeGLREC
       Case "NODEREC": RBA& = Len(NodeHDR) + (NodeHDR.WFMAXGL + 1) * Len(NodeGLREC)
                       X& = RNUM - 1: X& = X& * (Len(NodeREC) - 8 * ANCSEARCH)
                       RBA& = RBA& + X& + 1
                       Put #5, RBA&, NodeREC
                       If ANCSEARCH Then Put #5, , ALIN#
    End Select
    NodeHDR.WFUPDTY = DATEY(DATE$, 0)
    NodeHDR.WFUPDTMD = DATEMD(DATE$, 0)

End Sub



Sub NodeWorkGET(RTP$, RNUM)
3100 ANCSEARCH = ((Asc(NodeHDR.WFOPT) And 16) > 0)
    Select Case RTP$
       Case "HDR": Get #5, 1, NodeHDR
       Case "GLREC": RBA& = Len(NodeHDR)
                     RBA& = RBA& + RNUM * Len(NodeGLREC) + 1
                     Get #5, RBA&, NodeGLREC
       Case "NODEREC": RBA& = Len(NodeHDR) + (NodeHDR.WFMAXGL + 1) * Len(NodeGLREC)
                       X& = RNUM - 1: X& = X& * (Len(NodeREC) - 8 * ANCSEARCH)
                       RBA& = RBA& + X& + 1
                       Get #5, RBA&, NodeREC
                       If ANCSEARCH Then Get #5, , ALIN#
    End Select
End Sub

Sub NodeWrkOPEN(OPT$)
    ' Open Index REFERENCE File for Creation
    QUIET = (Right$(OPT$, 1) = "Q"): X = Len(RTrim$(OPT$)) + QUIET
    XOPT$ = MID$(OPT$, 1, X): DSN$ = FT$(8): NODEXData$ = ""
NODEDSN:
    If Not QUIET Then
       'Call FmtFIND("XXXXS017")
       'Call FmtFindFLD("DSNM"): LSet UTXT$(FFLD.UTXT) = DSN$
       'Call MsgFORMAT: If A = 27 Then GoTo NODEOpenBAD
       'Call FmtFindFLD("DSNM"): LSet DSN$ = UTXT$(FFLD.UTXT)
    End If
    DSNAME$ = RTrim$(WRKFPFX$ + DSN$)
    LSet SysVAR.FILENAME = DSNAME$
    On Local Error GoTo NONODE
    Open DSNAME$ For Input As #5
NODECHK:
    Close #5
    On Local Error GoTo NODEOpenERR
    If XOPT$ = "NEW" Then
       Call PutMSG("XXXXM002")
       If CA$ <> SNGLKEY$(1) Then GoTo NODEOpenBAD
    End If
    If (XOPT$ = "NEW") Then
       Kill SysVAR.FILENAME
       GoTo NODENEW
    End If
    GoTo NODEOLD
NONODE:
    If (XOPT$ = "NEW") Then Resume NODENEW
    If QUIET Then Resume NODEOpenBAD
    Resume NONODEMSG
NONODEMSG:
    Call PutMSG("XXXXM029"): If a = 27 Then GoTo NODEOpenBAD
    GoTo NODEDSN

NODEOLD:
    On Local Error GoTo NODEOpenERR
    Select Case XOPT$
       Case "INIT": Open SysVAR.FILENAME For Binary Access Read As #5: QUIET = 0
       Case "UPDT": Open SysVAR.FILENAME For Binary Access Read Write As #5
       Case "READ": Open SysVAR.FILENAME For Binary Access Read As #5
       Case "WRITE": Open SysVAR.FILENAME For Binary Access Read Write As #5
       Case Else: GoTo NODEOpenBAD
    End Select
    SysVAR.MVAR = ".WRK"
    Get #5, 1, NodeHDR
    If (NodeHDR.FType <> "W") Or (NodeHDR.WFTYPE <> Chr$(6)) Or (NodeHDR.WFVER <> Chr$(0)) Then
       Close #5: If QUIET Then GoTo NODEOpenBAD
       Call PutMSG("XXXXM033")
       If XOPT$ = "INIT" Then GoTo NODEDSN
       GoTo NODEOpenBAD
    End If
    If (NodeHDR.REORGY <> FF1Hdr.REORGY) Or (NodeHDR.REORGMD <> FF1Hdr.REORGMD) Or (NodeHDR.REORGTM <> FF1Hdr.REORGTM) Then
       Close #5: If QUIET Then GoTo NODEOpenBAD
       Call PutMSG("XXXXM034")
       If XOPT$ = "INIT" Then GoTo NODEDSN
       GoTo NODEOpenBAD
    End If
NODEOLDEnd:
    OKAY = -1: GoTo NODEOpenEnd

NODENEW:
    On Local Error GoTo NODEOpenERR
    If XOPT$ = "WRITE" Then Close #5: Kill SysVAR.FILENAME
    Open SysVAR.FILENAME For Binary Access Read Write As #5
    Dim NEWHdr As NodeFileHDR: NodeHDR = NEWHdr
    NodeHDR.FType = "W": NodeHDR.WFTYPE = Chr$(6): NodeHDR.WFVER = Chr$(255)
    NodeHDR.WFOPT = Chr$(0)
    NodeHDR.REORGY = FF1Hdr.REORGY
    NodeHDR.REORGMD = FF1Hdr.REORGMD
    NodeHDR.REORGTM = FF1Hdr.REORGTM
    NodeHDR.WFUPDTY = DATEY(DATE$, 0)
    NodeHDR.WFUPDTMD = DATEMD(DATE$, 0)
    Put #5, 1, NodeHDR
    OKAY = -1: GoTo NODEOpenEnd

NODEOpenERR:
    'CALL PutMSG("XXXXM004")   ' "Unable to OPEN File [FILENAME]..."
    SysVAR.NVAR1 = ERR
    Call PutMSG("XXXXM030")    ' "Error Opening [FILENAME]...ERR=[NVAR1]..."
    Resume NODEOpenBAD
NODEOpenBAD: OKAY = 0

NODEOpenEnd:
    If OKAY Then LSet FT$(8) = DSN$
End Sub


Sub BuildBASE(xType)
SBASE = xType
SCNT = 0: PSCOUNT = 0: SSCOUNT = 0
SlecType = SBASE
Select Case SBASE
   Case 0: GoSub 100
   Case 1: GoSub 150
   Case 2: GoSub 150
   Case 3: GoSub 150
   Case 4: GoSub 160
End Select
If SBASE <> 4 Then PSCOUNT = SCNT
Exit Sub

100 ' All Records
    SCNT = FF1Hdr.MAXID
    LOID = 1: HIID = SCNT
    ReDim SelSET(SCNT)
    For X = 1 To SCNT: SelSET(X) = 1: Next X
    Return


150 ' Load selset from Relative WRK File
    Call frmFHSWRelW.BUILD(SBASE)
    If OKAY Then
       Call RWRKOpen("READ", SBASE, 0)
       ReDim SelSET(FF1Hdr.MAXID)
       If OKAY Then
          Call RWRKGet("BASE", RC, PTR&)
          'Call PutMSG("XXXXM064")
          While (0 < RC) And (RC < 9)
             If WFRELREC.id > 0 Then
                RID = WFRELREC.id
                If SelSET(RID) = 0 Then
                   SCNT = SCNT + 1: SelSET(RID) = 1
                   If RID < LOID Then LOID = RID
                   If HIID < RID Then HIID = RID
                End If
             End If
             Call RWRKGet("WRKREC", RC, PTR&)
          Wend
          SelSET(0) = LOID
          RID = LOID
       End If
       'Call ClearMSG
       Call FamCLOSE
    End If
    Return
    
160 ' Select File
    ReDim SelSET(FF1Hdr.MAXID)
    HIID = 0
    Call SWRKBldList(SelSET(), HIID, 0)
    If HIID > 0 Then
       PSCOUNT = SLCTHDR.SlctCNT
       SSCOUNT = SLCTHDR.SECSCNT
       SCNT = SLCTHDR.SLCTTOT
       LOID = SLCTHDR.LOID
       HIID = SLCTHDR.HIID
    End If
    Return
End Sub

Function FNXL$(a&)
  FNXL$ = LTrim$(Str$(a&))
End Function
Function fnx$(a)
  fnx$ = LTrim$(Str$(a))
End Function
Rem $STATIC
Sub UpdtFamPTRS(OPT$, xRID, MRNO)
    OKAY = 0
    Select Case OPT$
       Case "REM": GoSub 910
       Case "ADD": GoSub 920
    End Select
    Exit Sub

910 ' Remove Family Record from Chain
    ' MRNO=Family Record Number, XRID = Name Record RID of Spouse Processed
    If (xRID = 0) Or (MRNO = 0) Then Return
    Call FF3GetRec(MRNO)
    If FF3SPOUSE.SP1ID = xRID Then
       N = FF3SPOUSE.SP1NXT
       Else
       If FF3SPOUSE.SP2ID = xRID Then
          N = FF3SPOUSE.SP2NXT
          Else
          N = 0
       End If
    End If
    Call FF1GetRec(xRID): If Not OKAY Then Return
    If FF1Rec.SPOUSE = MRNO Then
       FF1Rec.SPOUSE = N: Call FF1PutREC
       Else
       X = FF1Rec.SPOUSE: FOUND = 0
       While (X <> 0) And (Not FOUND)
          FOUND = (X = MRNO)
          If Not FOUND Then
             Call FF3GetRec(X)
             If Not OKAY Then
                X = 0
                Else
                If xRID = FF3SPOUSE.SP1ID Then
                   X = FF3SPOUSE.SP1NXT
                   Else
                   If xRID = FF3SPOUSE.SP2ID Then
                      X = FF3SPOUSE.SP2NXT
                      Else
                      X = 0
                   End If
                End If
             End If
          End If
       Wend
       If FOUND Then
          If FF3SPOUSE.SP1ID = xRID Then
             FF3SPOUSE.SP1NXT = N
             Else
             FF3SPOUSE.SP2NXT = N
          End If
          Call FF3PutRec
       End If
    End If
    Return

920 ' Add Family Record to Chain
    ' MRNO=Family Record Number, XRID = Name Record RID of Spouse Processed
    If (xRID = 0) Or (MRNO = 0) Then Return
    Call FF1GetRec(xRID): If Not OKAY Then Return
    If FF1Rec.SPOUSE = 0 Then
       FF1Rec.SPOUSE = MRNO
       Call FF1PutREC
       Else
       Call FF3GetRec(MRNO): If Not OKAY Then Return
       D1MD = FF3SPOUSE.D1MD: D1Y = FF3SPOUSE.D1Y
       X = FF1Rec.SPOUSE: Y = 0: FOUND = 0
       While (X <> 0) And (Not FOUND)
          Call FF3GetRec(X)
          If Not OKAY Then
             X = 0
             Else
             If (FF3SPOUSE.D1Y > D1Y) Or ((FF3SPOUSE.D1Y = D1Y) And (FF3SPOUSE.D1MD > D1MD)) Then
                Y = X
                If xRID = FF3SPOUSE.SP1ID Then
                   X = FF3SPOUSE.SP1NXT
                   Else
                   If xRID = FF3SPOUSE.SP2ID Then
                      X = FF3SPOUSE.SP2NXT
                      Else
                      X = 0
                   End If
                End If
                Else
                FOUND = -1
             End If
          End If
       Wend
       If OKAY Then
          If Y = 0 Then
             N = FF1Rec.SPOUSE
             FF1Rec.SPOUSE = MRNO
             Call FF1PutREC
             Else
             Call FF3GetRec(Y)
             If X = 0 Then
                If FF3SPOUSE.SP1ID = xRID Then
                   FF3SPOUSE.SP1NXT = MRNO
                   Else
                   FF3SPOUSE.SP2NXT = MRNO
                End If
                N = 0
                Else
                If FF3SPOUSE.SP1ID = xRID Then
                   N = FF3SPOUSE.SP1NXT: FF3SPOUSE.SP1NXT = MRNO
                   Else
                   N = FF3SPOUSE.SP2NXT: FF3SPOUSE.SP2NXT = MRNO
                End If
             End If
             Call FF3PutRec
          End If
       End If
       Call FF3GetRec(MRNO)
       If FF3SPOUSE.SP1ID = xRID Then
          FF3SPOUSE.SP1NXT = N
          Else
          FF3SPOUSE.SP2NXT = N
       End If
       Call FF3PutRec
    End If
    Return

End Sub

Sub UpdtRelPTRS(OPT$, XID)
    PID = XID
    Select Case OPT$
       Case "REM": GoSub 665
       Case "ADD": GoSub 667
    End Select
    Exit Sub

' UPDATE NAME Record Relationship Chains
665 ' Remove from PID=CFID/CMID child/sibling chain
    CFID = FF1Rec.FID: CMID = FF1Rec.MID
    Call FF1GetRec(PID)
    If OKAY Then
       If FF1Rec.OLDCH = NameREC.RID Then
          If PID = CFID Then
             FF1Rec.OLDCH = NameREC.FCH
             Else
             If PID = CMID Then
                FF1Rec.OLDCH = NameREC.MCH
                Else
                FF1Rec.OLDCH = 0
             End If
          End If
          Else
          X = FF1Rec.OLDCH
          While (X > 0) And OKAY
             Call FF1GetRec(X)
             If OKAY Then
                If PID = FF1Rec.FID Then
                   If FF1Rec.FCH = NameREC.RID Then
                      If PID = CFID Then
                         FF1Rec.FCH = NameREC.FCH
                         Else
                         If PID = CMID Then
                            FF1Rec.FCH = NameREC.MCH
                            Else
                            FF1Rec.FCH = 0
                         End If
                      End If
                      X = 0
                      Else
                      X = FF1Rec.FCH
                   End If
                   Else
                   If FF1Rec.MCH = NameREC.RID Then
                      If PID = CFID Then
                         FF1Rec.MCH = NameREC.FCH
                         Else
                         If PID = CMID Then
                            FF1Rec.MCH = NameREC.MCH
                            Else
                            FF1Rec.MCH = 0
                         End If
                      End If
                      X = 0
                      Else
                      X = FF1Rec.MCH
                   End If
                End If
             End If
          Wend
       End If
       If OKAY Then Call FF1PutREC
       Call FF1GetRec(NameREC.RID)
       If PID = CFID Then NameREC.FCH = 0: FF1Rec.FID = 0: FF1Rec.FCH = 0
       If PID = CMID Then NameREC.MCH = 0: FF1Rec.MID = 0: FF1Rec.MCH = 0
       Call FF1PutREC
    End If
    Return

667 ' ADD to PID=UFID/UMID child/sibling chain
    UFID = NameREC.FID: UMID = NameREC.MID
    Call FF1GetRec(PID)
    If OKAY Then
       If FF1Rec.OLDCH = 0 Then
          FF1Rec.OLDCH = NameREC.RID
          If PID = UFID Then NameREC.FCH = 0 Else NameREC.MCH = 0
          Else
          X = FF1Rec.OLDCH: Y = 0
          While (X > 0) And OKAY
             Call FF1GetRec(X)
             If OKAY Then
                If (FF1Rec.BY < NameREC.BY) Or ((FF1Rec.BY = NameREC.BY) And (FF1Rec.BMD <= NameREC.BMD)) Then
                   Y = X
                   If PID = FF1Rec.FID Then X = FF1Rec.FCH Else X = FF1Rec.MCH
                   Else
                   X = 0
                End If
             End If
          Wend
          If OKAY Then
             If Y = 0 Then
                Call FF1GetRec(PID)
                If PID = NameREC.FID Then
                   NameREC.FCH = FF1Rec.OLDCH
                   Else
                   NameREC.MCH = FF1Rec.OLDCH
                End If
                FF1Rec.OLDCH = NameREC.RID
                Else
                Call FF1GetRec(Y)
                If PID = NameREC.FID Then
                   If FF1Rec.FID = PID Then
                      NameREC.FCH = FF1Rec.FCH: FF1Rec.FCH = NameREC.RID
                      Else
                      If FF1Rec.MID = PID Then
                         NameREC.FCH = FF1Rec.MCH: FF1Rec.MCH = NameREC.RID
                         Else
                         NameREC.FCH = 0
                      End If
                   End If
                   Else    ' i.e. PID=NameREC.MID
                   If FF1Rec.FID = PID Then
                      NameREC.MCH = FF1Rec.FCH: FF1Rec.FCH = NameREC.RID
                      Else
                      If FF1Rec.MID = PID Then
                         NameREC.MCH = FF1Rec.MCH: FF1Rec.MCH = NameREC.RID
                         Else
                         NameREC.MCH = 0
                      End If
                   End If
                End If
             End If
          End If
       End If
       If OKAY Then Call FF1PutREC
    End If
    Return

End Sub

