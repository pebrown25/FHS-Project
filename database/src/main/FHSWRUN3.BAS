Attribute VB_Name = "dllFHSWRun3"
DefInt A-Z
Function CharCount(X)
' Convert inches*100 into "inches" of characters
Printer.ScaleMode = 5
xx# = Printer.TextWidth(StdWidth) / Len(StdWidth)
If xx# = 0 Then xx# = 1
CharCount = Int(X / (100 * xx#))
End Function
Sub FF3Open(OPT$)
    OKAY = 0: D$ = FDFSetup.NFDRV: CA$ = ""
    QUIET = (Right$(OPT$, 1) = "Q"): X = Len(RTrim$(OPT$)) + QUIET
    XOPT$ = MID$(OPT$, 1, X)
    If (("A" <= D$) And (D$ <= "Z")) And (D$ <> FDFSetup.FAMDRV) Then
       DSNAME$ = D$ + ":" + FT$(3)
       Else
       DSNAME$ = FAMFPFX$ + FT$(3)
    End If
    LSet SysVAR.FILENAME = DSNAME$
    On Local Error GoTo NOFF3
CHKFF3:
    Open DSNAME$ For Input As #3
    Close #3
    If XOPT$ = "INIT" Then
       If Not QUIET Then Call PutMSG("XXXXM051")
       If CA$ <> SNGLKEY$(1) Then GoTo FF3OpenEND
       GoTo OLDFF3
    End If
    If XOPT$ = "NEW" Then
       If Not QUIET Then Call PutMSG("XXXXM002")
       If CA$ <> SNGLKEY$(1) Then GoTo FF3OpenEND
       Kill SysVAR.FILENAME
       GoTo NEWFF3
    End If
    GoTo OLDFF3
NOFF3:
    X = ERR
    If XOPT$ = "NEW" Then Resume NEWFF3
    If XOPT$ = "INIT" Then OKAY = -1: Resume FF3OpenEND
    If Not QUIET Then Call ERRMessage("XXXXM030", X, 0)
    If (CA$ = Chr$(27)) Or (X <> 71) Then Resume FF3OpenEND
    Resume CHKFF3
OLDFF3:
    On Local Error GoTo FF3OpenERR
    Select Case XOPT$
       Case "INIT": Open SysVAR.FILENAME For Binary Access Read As #3
       Case "READ": Open SysVAR.FILENAME For Binary Access Read As #3
       Case "WRITE": Open SysVAR.FILENAME For Binary Access Read Write As #3
       Case Else: GoTo FF3OpenEND
    End Select
    Get #3, 1, FF3HDR
    LSet SysVAR.MVAR = "MISCINFO"
    If FF3HDR.FType <> "M" Then
       If Not QUIET Then Call PutMSG("XXXXM033")
       If CA$ <> SNGLKEY$(1) Then GoTo FF3OpenEND
       FF3HDR.FType = "M"
       If XOPT$ = "WRITE" Then FF3HDR.REVLVL = FF3HDR.REVLVL Or 256
    End If
    If ((FF3HDR.REVLVL And 255) <> 0) Then
       If Not QUIET Then Call PutMSG("XXXXM036")
       If CA$ <> SNGLKEY$(1) Then GoTo FF3OpenEND
       FF3HDR.REVLVL = 0
       If XOPT$ = "WRITE" Then FF3HDR.REVLVL = FF3HDR.REVLVL Or 256
    End If
    If (FF3HDR.REORGY <> FF1Hdr.REORGY) Or (FF3HDR.REORGMD <> FF1Hdr.REORGMD) Or (FF3HDR.REORGTM <> FF1Hdr.REORGTM) Then
       If Not QUIET Then Call PutMSG("XXXXM034")
       If CA$ <> SNGLKEY$(1) Then GoTo FF3OpenEND
       FF3HDR.REORGY = FF1Hdr.REORGY: FF3HDR.REORGMD = FF1Hdr.REORGMD
       FF3HDR.REORGTM = FF1Hdr.REORGTM
       If XOPT$ = "WRITE" Then FF3HDR.REVLVL = FF3HDR.REVLVL Or 256
    End If
    OKAY = -1
    GoTo FF3OpenEND
NEWFF3:
    On Local Error GoTo FF3OpenERR
    Open SysVAR.FILENAME For Binary Access Read Write As #3
    FF3HDR.FType = "M"
    FF3HDR.REVLVL = 0
    FF3HDR.HIREC = 0
    FF3HDR.REORGY = FF1Hdr.REORGY
    FF3HDR.REORGMD = FF1Hdr.REORGMD
    FF3HDR.REORGTM = FF1Hdr.REORGTM
    FF3HDR.UPDTY = FF1Hdr.UPDTY
    FF3HDR.UPDTMD = FF1Hdr.UPDTMD
    FF3HDR.FREEBGN = 0: FF3HDR.FREECNT = 0
    'FF3HDR.HDRFILL = STRING$(LEN(FF3HDR.HDRFILL), 0)
    Put #3, 1, FF3HDR
    OKAY = -1
    GoTo FF3OpenEND
FF3OpenERR:
    If Not QUIET Then Call ERRMessage("XXXXM030", ERR, ERL)
    Resume FF3OpenEND
FF3OpenEND:
End Sub

Rem $DYNAMIC
Sub ChkID(X&)
    If X& > FF1Hdr.MAXID Then Call ErrBEEP(0): OKAY = 0 Else OKAY = -1
End Sub

Sub ComPRINT(ComPTR, COMBGN)
2920 ' Print COMMENT Lines
     BGNBKT = 0
     COFST = PAGE.COFST: PAGE.COFST = COFST + PAGE.CINDENT
     BWIDTH = PAGE.BWIDTH: PAGE.BWIDTH = BWIDTH - PAGE.CINDENT
     If PAGE.LOFST = 0 Then PAGE.LOFST = COFST
     If RptOPTION.FREEFORM Or COMBGN Then
        LSet CL$ = ";": OFST = 3
        Else
        LSet CL$ = RFLBL$(1)
        OFST = InStr(CL$, " ") + 1
     End If
     ' Print Comment Information
     While ComPTR <> 0: Call FF3GetRec(ComPTR): ComPTR = FF3COMMENT.FPTR
        Mid$(CL$, OFST, 38) = FF3COMMENT.TEXT
        If RptOPTION.IGNBKT Then
           ENDBKT = 1
           While ENDBKT <> 0
              If BGNBKT = 0 Then BGNBKT = InStr(OFST, CL$, "{")
              If BGNBKT = 0 Then
                 ENDBKT = 0
                 Else
                 OFST = BGNBKT
                 ENDBKT = InStr(BGNBKT, CL$, "}")
                 If ENDBKT = 0 Then
                    Mid$(CL$, BGNBKT, 38) = Space$(38)
                    Else
                    Mid$(CL$, BGNBKT, 1) = " "
                    Mid$(CL$, BGNBKT + 1, 38) = MID$(CL$, ENDBKT + 1, 38)
                    BGNBKT = 0
                 End If
              End If
           Wend
           Else
           If RptOPTION.RMVBKT Then
              X = 1
              While X > 0
                 X = InStr(X, CL$, "{")
                 If X > 0 Then Mid$(CL$, X, 1) = " "
              Wend
              X = 1
              While X > 0
                 X = InStr(X, CL$, "}")
                 If X > 0 Then Mid$(CL$, X, 1) = " "
              Wend
           End If
        End If

        ' Process NEWLINE requests
        LCTRL = 1
        While LCTRL > 0
           LCTRL = InStr(1, CL$, "\")
           If LCTRL > 0 Then
              HLCNT = PAGE.LINE
              xx$ = RTrim$(MID$(CL$, LCTRL + 1, 255))
              LSet CL$ = Left$(CL$, LCTRL - 1): Call PutDATA: COMBGN = 1
              If (HLCNT = PAGE.LINE) Then Call WriteLINE: PAGE.LOFST = PAGE.COFST
              LSet CL$ = xx$
           End If
        Wend

        ' Finish Processing COMMENT Record
        xx$ = RTrim$(CL$)
        If Not ((xx$ = "") Or (xx$ = ";") Or (xx$ = RFLBL$(1))) Then
           Call PutDATA: xx$ = "": COMBGN = 1: If BGNBKT > 0 Then BGNBKT = 1
        End If
        OFST = Len(xx$) + 1
        If BGNBKT > 0 Then
           BGNBKT = OFST
           Else
           If MID$(FF3COMMENT.TEXT, 38, 1) = " " Then OFST = OFST + 1
        End If
        If a = 27 Then ComPTR = 0
     Wend

    ' Return to Caller
    PAGE.COFST = COFST: PAGE.BWIDTH = BWIDTH: PAGE.LWIDTH = BWIDTH
End Sub

Sub ComputeSNDX(T$, S$)

     ' Compute Soundex
     X = 1: TL = Len(T$): Y = 1: SL = Len(S$)
     SXA = 0: SXH = 0
     LSet S$ = String$(SL, Asc("0"))
SNDX1:
     If (X > TL) Or (Y > SL) Then GoTo SNDXE
     SXA = Asc(MID$(T$, X, 1)): X = X + 1
     SXT = Asc(MID$(SNDXTB$, SXA, 1))
     If (SXT > 254) Then If Y = 1 Then GoTo SNDX1 Else GoTo SNDXE
     If Y = 1 Then
        Mid$(S$, 1, 1) = UCX$(Chr$(SXA))
        Else
        If (SXT = 48) Or (SXT = SXH) Then GoTo SNDX1
        Mid$(S$, Y, 1) = Chr$(SXT)
        SXH = SXT
     End If
     Y = Y + 1
     GoTo SNDX1
SNDXE: 'RETURN

End Sub

Function DATEMD(xx$, Y)
   M = Val(MID$(xx$, 1 - 3 * (Y = 1) - 5 * (Y = 2), 2))
   D = Val(MID$(xx$, 4 + 3 * (Y = 1) - 5 * (Y = 2), 2))
   DATEMD = 100 * M + D
End Function

Function DATEDif$(D1MD, D1Y, D2MD, D2Y)
   MD1 = D1MD: Y1 = D1Y
   MD2 = D2MD: Y2 = D2Y
   If (Y2 = 0) And (MD2 = 0) Then Y2 = DATEY(DATE$, 0): MD2 = DATEMD(DATE$, 0)
   If (Y1 = 0) Or (Y2 = 0) Or (Y2 = 9999) Then
      DATEDif$ = " ??"
      Else
      If MD2 < MD1 Then DIF = Y2 - Y1 - 1 Else DIF = Y2 - Y1
      If DIF < 0 Then
         DATEDif$ = " --"
         Else
         DATEDif$ = LTrim$(Str$(DIF))
      End If
   End If
End Function


Function DATEY(xx$, Y)
    DATEY = Val(MID$(xx$, 7 + 6 * (Y = 2), 4))
End Function

Rem $STATIC
Sub EvenBREAK(xx$)
    If (RptOPTION.EVENPAGE) And ((PAGE.COUNT Mod 2) > 0) Then
       Call PageBREAK: If xx$ = "EOF" Then PRT.EOF = -1
       Call GetRVAR("BLANKPG", X)
       BL$ = PLINE$
       If FOUND Then
          LSet BL$ = RVAR$(X)
          Call HTFBuild(BL$, PAGE.BWIDTH)
          Else
          BL$ = ""
       End If
       X = PAGE.BLENGTH / 3
       While X > 0: Call WriteLINE: X = X - 1: Wend
       LSet PLINE$ = BL$
       Call WriteLINE
       Else
       If xx$ = "EOF" Then PRT.EOF = -1
    End If
    Call PageBREAK
End Sub

Rem $DYNAMIC
Sub FamCLOSE()
   If ENV.FFOPEN And FFCHANGED Then
      Y = DATEY(DATE$, 0): MD = DATEMD(DATE$, 0)
      If FF1Hdr.REVLVL > 255 Then
         FF1Hdr.UPDTY = Y: FF1Hdr.UPDTMD = MD
         FF1Hdr.REVLVL = FF1Hdr.REVLVL And 255: Put #1, 1, FF1Hdr
      End If
      If FF2Hdr.REVLVL > 255 Then
         FF2Hdr.UPDTY = Y: FF2Hdr.UPDTMD = MD
         FF2Hdr.REVLVL = FF2Hdr.REVLVL And 255: Put #2, 1, FF2Hdr
      End If
      If FF3HDR.REVLVL > 255 Then
         FF3HDR.UPDTY = Y: FF3HDR.UPDTMD = MD
         FF3HDR.REVLVL = FF3HDR.REVLVL And 255: Put #3, 1, FF3HDR
      End If
   End If
   Close: ENV.FFOPEN = 0: FFCHANGED = 0: PRT.OPEN = Chr$(0)
End Sub

Rem $STATIC
Sub FAMOpen(OPT$)
    If ENV.FFOPEN Then Call FamCLOSE
    xx$ = OPT$
    If (xx$ = "VLDT") Or (xx$ = "VLDTQ") Then
       VOPT = -1
       If xx$ = "VLDTQ" Then xx$ = "WRITEQ" Else xx$ = "WRITE"
    End If
    Call FF1Open(xx$): If Not OKAY Then GoTo FAMOpenEND
    Call FF2Open(xx$): If Not OKAY Then GoTo FAMOpenEND
    Call FF3Open(xx$): If Not OKAY Then GoTo FAMOpenEND
    If (xx$ <> "INIT") And (xx$ <> "INITQ") Then GoTo FAMOpenEND
    Close
    Call FF1Open("NEW"): If Not OKAY Then GoTo FAMOpenEND
    Call FF2Open("NEW"): If Not OKAY Then GoTo FAMOpenEND
    Call FF3Open("NEW")
FAMOpenEND:
    ENV.FFOPEN = OKAY: FFCHANGED = 0
    If Not OKAY Then
       Close
       Else
       If Not VOPT Then
          If ((FF1Hdr.REVLVL > 255) Or (FF2Hdr.REVLVL > 255) Or (FF3HDR.REVLVL > 255)) Then
             Call PutMSG("XXXXM073")
          End If
       End If
    End If
End Sub

Sub FDFGetSet(XSETUP$)
    OKAY = 0: SysVAR.MVAR = XSETUP$
    On Local Error GoTo FDFGetSetERR
    If XSETUP$ = "DEFAULT" Then
       X = FDFHDr.DEFAULT
       If (X < 1) Or (X > FDFHDr.SETUPS) Then X = 1: FDFHDr.DEFAULT = 1
    Else
       ZSETUP$ = Space$(8): LSet ZSETUP$ = XSETUP$
       X = FSUTABLE.FT: FOUND = 0
       While (X > 0) And (Not FOUND)
          FOUND = (MID$(FSETUP$(X), 1, 8) = ZSETUP$)
          If Not FOUND Then X = FSUCHN(X).FWD
       Wend
       If Not FOUND Then
          Call PutMSG("XXXXM027")
          Exit Sub
       End If
       If X = ENV.FSETUP Then OKAY = -1: Exit Sub
    End If
    'If FDFSURBA(x) = 0 Then
    '   Call PutMSG("XXXXM028")
    '   Exit Sub
    'End If
    'Call FDFOpen("READ")
    'If Not OKAY Then Exit Sub
    ENV.FSETUP = X
    'Get #1, FDFSURBA(x), FDFSetup
    ' FDFSETUP.NAME
    FDFSetup = FSETUPS(X)
    'For x = 1 To FDFHDr.FAMDSN
    '    Get #1, , FT$(x)
    'Next x
    For Z = 1 To FDFHDr.FAMDSN
       LSet FT$(Z) = FFNAME$(FDFHDr.FAMDSN * (X - 1) + Z)
    Next Z
    FAMFPFX$ = "": WRKFPFX$ = "": TRANPFX$ = "": RPTFPFX$ = ""
    D$ = FDFSetup.FAMDRV: If ("A" <= D$) And (D$ <= "Z") Then FAMFPFX$ = D$ + ":"
    If MID$(FDFSetup.FAMDIR, 1, 1) <> "\" Then
       SysVAR.FILENAME = FAMFPFX$
       FAMFPFX$ = CurDir$(FAMFPFX$)
       If MID$(FAMFPFX$, Len(FAMFPFX$), 1) <> "\" Then FAMFPFX$ = FAMFPFX$ + "\"
    End If
    FAMFPFX$ = RTrim$(FAMFPFX$ + FDFSetup.FAMDIR)
    If MID$(FAMFPFX$, Len(FAMFPFX$), 1) <> "\" Then FAMFPFX$ = FAMFPFX$ + "\"
    D$ = FDFSetup.WRKDRV: If ("A" <= D$) And (D$ <= "Z") Then WRKFPFX$ = D$ + ":"
    If MID$(FDFSetup.WRKDIR, 1, 1) <> "\" Then
       SysVAR.FILENAME = WRKFPFX$
       WRKFPFX$ = CurDir$(WRKFPFX$)
       If MID$(WRKFPFX$, Len(WRKFPFX$), 1) <> "\" Then WRKFPFX$ = WRKFPFX$ + "\"
    End If
    WRKFPFX$ = RTrim$(WRKFPFX$ + FDFSetup.WRKDIR)
    If MID$(WRKFPFX$, Len(WRKFPFX$), 1) <> "\" Then WRKFPFX$ = WRKFPFX$ + "\"
    D$ = FDFSetup.TRANDRV: If ("A" <= D$) And (D$ <= "Z") Then TRANPFX$ = D$ + ":"
    If MID$(FDFSetup.TRANDIR, 1, 1) <> "\" Then
       SysVAR.FILENAME = TRANPFX$
       TRANPFX$ = CurDir$(TRANPFX$)
       If MID$(TRANPFX$, Len(TRANPFX$), 1) <> "\" Then TRANPFX$ = TRANPFX$ + "\"
    End If
    TRANPFX$ = RTrim$(TRANPFX$ + FDFSetup.TRANDIR)
    If MID$(TRANPFX$, Len(TRANPFX$), 1) <> "\" Then TRANPFX$ = TRANPFX$ + "\"
    D$ = FDFSetup.RPTDRV: If ("A" <= D$) And (D$ <= "Z") Then RPTFPFX$ = D$ + ":"
    If MID$(FDFSetup.RPTDIR, 1, 1) <> "\" Then
       SysVAR.FILENAME = RPTFPFX$
       RPTFPFX$ = CurDir$(RPTFPFX$)
       If MID$(RPTFPFX$, Len(RPTFPFX$), 1) <> "\" Then RPTFPFX$ = RPTFPFX$ + "\"
    End If
    RPTFPFX$ = RTrim$(RPTFPFX$ + FDFSetup.RPTDIR)
    If MID$(RPTFPFX$, Len(RPTFPFX$), 1) <> "\" Then RPTFPFX$ = RPTFPFX$ + "\"
    OKAY = -1: GoTo FDFGetSetExit
    
FDFGetSetERR:
    Call ERRMessage("XXXXM097", ERR, ERL)
    OKAY = 0: Resume Next
    
FDFGetSetExit:
    Call FamCLOSE
End Sub

Sub FDFOpen(OPT$)
    OKAY = 0: XOPT$ = OPT$
    QUIET = (MID$(OPT$, Len(OPT$), 1) = "Q")
    If QUIET Then XOPT$ = Left$(OPT$, Len(OPT$) - 1)
LocateFDF:
    If Not QUIET Then
       Call frmFHSWFile.LOCATE("FDF", "FILE")
       If Not OKAY Then GoTo FDFOpenEND
    End If
    LSet SysVAR.FILENAME = CONFIG.FDF + ".FDF"
    On Local Error GoTo NOFDF
CHKFDF:
    Open SysVAR.FILENAME For Input As #1
    Close #1
    If XOPT$ = "NEW" Then
       Call PutMSG("XXXXM002")
       If CA$ <> SNGLKEY$(1) Then OKAY = False: GoTo FDFOpenEND
       Kill SysVAR.FILENAME
       GoTo NEWFDF
    End If
    GoTo OLDFDF
NOFDF:
    X = ERR: Call FamCLOSE
    If XOPT$ = "NEW" Then Resume NEWFDF
    Call ERRMessage("XXXXM030", X, 0)
    If CA$ = Chr$(27) Then Resume FDFOpenEND
    Resume CHKFDF
OLDFDF:
    On Local Error GoTo FDFOpenERR
    Select Case XOPT$
       Case "INIT": Open SysVAR.FILENAME For Binary Access Read As #1
       Case "READ": Open SysVAR.FILENAME For Binary Access Read As #1
       Case "WRITE": Open SysVAR.FILENAME For Binary Access Read Write As #1
       Case Else: GoTo FDFOpenEND
    End Select
    Get #1, , FDFHDr
    LSet SysVAR.MVAR = "FDF"
    If FDFHDr.FType <> "FDF" Then
       Call PutMSG("XXXXM033")
       GoTo FDFOpenEND
    End If
    If FDFHDr.VER <> Chr$(0) Then
       Call PutMSG("XXXXM036")
       GoTo FDFOpenEND
    End If
    OKAY = -1
    If XOPT$ <> "INIT" Then GoTo FDFOpenEND
    For X = 1 To (FDFHDr.FAMDSN + FDFHDr.OTHERDSN)
        Get #1, , FDFDSRec
        Y = 0: FOUND = 0
        While (Y < ENV.DSNCNT) And (Not FOUND)
           Y = Y + 1: FOUND = (FType$(Y) = FDFDSRec.TYPE)
        Wend
        If FOUND Then FT$(Y) = FDFDSRec.DSNNAME
    Next X
    FSUMAX = FDFHDr.SETUPS
    ReDim FDFSURBA(FSUMAX), FSETUP$(FSUMAX), FSUCHN(FSUMAX) As ChainPTRS
    ReDim FSETUPS(FSUMAX) As FileSetupREC, FFNAME$(FDFHDr.FAMDSN * FSUMAX)
    Call ChnINIT(FSUCHN(), 1, FSUMAX)
    FSUTABLE.FT = 1: FSUTABLE.LT = FSUMAX: FSUTABLE.CL = 0
    FSUTABLE.SIZE = FSUMAX: FSUTABLE.FF = 0
    X = 0: RBA = Seek(1): XFN$ = Space$(12)
    While X < FDFHDr.SETUPS
       X = X + 1
       FDFSURBA(X) = RBA
       Get #1, RBA, FDFSetup
       FSETUP$(X) = Space$(FSUTABLE.XL)
       Mid$(FSETUP$(X), 1, 8) = FDFSetup.NAME
       Mid$(FSETUP$(X), 9, 255) = FDFSetup.DESC
       FSETUPS(X) = FDFSetup
       For Z = 1 To FDFHDr.FAMDSN
          Get #1, , XFN$
          FFNAME$(FDFHDr.FAMDSN * (X - 1) + Z) = RTrim$(XFN$)
       Next Z
       RBA = RBA + Len(FDFSetup) + FDFHDr.FAMDSN * 12
    Wend
    Close #1
    Call FDFGetSet("DEFAULT")
    GoTo FDFOpenEND
NEWFDF:
    On Local Error GoTo FDFOpenERR
    Open SysVAR.FILENAME For Binary Access Read Write As #1
    OKAY = -1
    GoTo FDFOpenEND
FDFOpenERR:
    Call ERRMessage("XXXXM030", ERR, ERL)
    Resume FDFOpenEND
FDFOpenEND:
    If Not OKAY Then Call FamCLOSE Else ENV.FDF = CONFIG.FDF
    On Local Error GoTo 0
End Sub


Function FDSTAT$(ST$)
   xx$ = " ": X = Asc(FF1Rec.STATUS)
   Select Case ST$
      Case "FID": Select Case (X And 3)
                    Case 1: LSet xx$ = "*"
                    Case 2: LSet xx$ = "?"
                  End Select
      Case "MID": Select Case (X And 12)
                    Case 4: LSet xx$ = "*"
                    Case 8: LSet xx$ = "?"
                  End Select
      Case "BD": Select Case (X And 48)
                   Case 16: LSet xx$ = "?"
                   Case 32: LSet xx$ = "!"
                 End Select
      Case "DD": Select Case (X And 192)
                   Case 64: LSet xx$ = "?"
                   Case 128: LSet xx$ = "!"
                 End Select
      Case "SNU": If ((Asc(FF1Rec.STATUS2) And 128) = 128) Then LSet xx$ = SNGLKEY$(1)
   End Select
   FDSTAT$ = xx$
End Function

Sub FF1GetNew()
   MAXID& = FFRBA&(FF1Hdr.MAXID, 1) + 1
   xMAXID = FFRNUM(MAXID&, 1)
   If xMAXID > 0 Then
      FF1Hdr.MAXID = xMAXID
      RBA& = 100 * MAXID& + 1
      NF$ = String$(50, 0) + Space$(50)
      Put #1, RBA&, NF$
      Get #1, RBA&, FF1Rec
      FF1Rec.RTYPE = Chr$(1)
      FF1Rec.RID = xMAXID
      FF1Rec.SEX = " "
      FF1Hdr.REVLVL = FF1Hdr.REVLVL Or 256
      Put #1, 1, FF1Hdr
      Put #1, RBA&, FF1Rec
      OKAY = -1
   End If
End Sub

Rem $DYNAMIC
Sub FF1GetRec(RNUM)
   OKAY = 0
   If ENV.FFOPEN Then
      MAXID& = FFRBA&(FF1Hdr.MAXID, 1)
      RBA& = FFRBA&(RNUM, 1)
      If (RBA& < 1) Or (RBA& > MAXID&) Then
         SysVAR.MVAR = LTrim$(Str$(RBA&))
         Call PutMSG("XXXXM050")
         Else
         RBA& = RBA& * 100 + 1
         Get #1, RBA&, FF1Rec
         OKAY = -1
      End If
   End If
End Sub


Sub FF3GetRec(RNUM)
   OKAY = 0
   If Not ENV.FFOPEN Then GoTo FF3GetEND
   HIREC& = FFRBA&(FF3HDR.HIREC, 1)
   If RNUM = 0 Then GoTo FF3GetNew
   RBA& = FFRBA&(RNUM, 1)
   LSet SysVAR.MVAR = "MISCINFO"
   If (RBA& < 1) Or (RBA& > HIREC&) Then
      SysVAR.MVAR = LTrim$(Str$(RBA&))
      Call PutMSG("XXXXM054")
      GoTo FF3GetEND
   End If
   RBA& = RBA& * 50 + 1
FF3GetIT:
   Get #3, RBA&, FF3PFX
   Select Case FF3PFX.RTYPE
      Case Chr$(3): Get #3, , FF3SPOUSE
      Case Chr$(4): Get #3, , FF3PLACE
                Select Case FF3PLACE.SRTYPE
                   Case Chr$(1): LSet BPLACE = FF3PLACE.PLACE1
                                 LSet DPLACE = FF3PLACE.PLACE2
                   Case Chr$(3): LSet MPLACE = FF3PLACE.PLACE1
                                 LSet DVPLACE = FF3PLACE.PLACE2
                End Select
      Case Chr$(5): Get #3, , FF3COMMENT
      Case "3": Get #3, , FF3SPOUSE              ' Old SPOUSE Records
      Case "4": Get #3, , FF3PLACE               ' Old PLACE Records
                Select Case FF3PLACE.SRTYPE
                   Case Chr$(1): LSet BPLACE = FF3PLACE.PLACE1
                                 LSet DPLACE = FF3PLACE.PLACE2
                   Case Chr$(3): LSet MPLACE = FF3PLACE.PLACE1
                                 LSet DVPLACE = FF3PLACE.PLACE2
                End Select
      Case "F": Get #3, , FF3COMMENT             ' Old COMMENT Records
      Case Chr$(6): Get #3, , FF3EVENT           ' New EVENT Records
      Case Chr$(11): Get #3, , FF3EVENT          ' Old EVENT Records
      Case Chr$(12): Get #3, , FF3PLACE2
                Select Case FF3PLACE2.PTYPE
                   Case Chr$(1): LSet BPLACE = FF3PLACE2.PLACE
                   Case Chr$(2): LSet DPLACE = FF3PLACE2.PLACE
                   Case Chr$(3): LSet MPLACE = FF3PLACE2.PLACE
                   Case Chr$(4): LSet DVPLACE = FF3PLACE2.PLACE
                End Select
      Case Else: Get #3, , FF3PNTRS
                 Select Case FF3PFX.RTYPE
                    Case Chr$(7): Get #3, , FF3EDUC
                    Case Chr$(8): Get #3, , FF3WORK
                    Case Chr$(9): Get #3, , FF3MIL
                    Case Chr$(10): Get #3, , FF3HEALTH
                    Case "7": Get #3, , FF3EDUC    ' Old EDUC Records
                    Case "8": Get #3, , FF3WORK    ' Old WORK Records
                    Case "9": Get #3, , FF3MIL     ' Old MIL Records
                    Case "A": Get #3, , FF3HEALTH  ' Old HLTH Records
                    Case Chr$(255):
                    Case Else: SysVAR.MVAR = LTrim$(Str$(Asc(FF3PFX.RTYPE)))
                               'CALL PutMSG("XXXXM055")
                               GoTo FF3GetEND
                 End Select
   End Select
   OKAY = -1
   GoTo FF3GetEND
FF3GetNew:
   If FF3HDR.FREECNT <> 0 Then
      RNO = FF3HDR.FREEBGN
      RBA& = FFRBA&(RNO, 50) + 1
      If RBA& > 1 Then
         Get #3, RBA& + 3, FF3PNTRS
         FF3HDR.FREEBGN = FF3PNTRS.NXT
         FF3HDR.FREECNT = FFRNUM(FFRBA&(FF3HDR.FREECNT, 1) - 1, 1)
      End If
      Else
      RBA& = 0
   End If
   If RBA& = 0 Then
      HIREC& = HIREC& + 1: HIREC3 = FFRNUM(HIREC&, 1)
      If HIREC3 = 0 Then GoTo FF3GetEND
      FF3HDR.HIREC = HIREC3
      RNO = HIREC3
      RBA& = HIREC& * 50 + 1
   End If
   If RBA& > 1 Then
      FF3HDR.REVLVL = FF3HDR.REVLVL Or 256
      Put #3, 1, FF3HDR
      Select Case FF3PFX.RTYPE
         Case Chr$(3): ZL = 23: XL = 24
         Case Chr$(4): ZL = 3: XL = 44
         Case Chr$(5): ZL = 9: XL = 38
         Case Chr$(6): ZL = 16: XL = 31
         Case Chr$(11): ZL = 16: XL = 31  ' Old Event Records
         Case Chr$(12): ZL = 6: XL = 41   ' New PLACE Name Record
         Case Else: ZL = 17: XL = 30
      End Select
      AF$ = FF3PFX.RTYPE + MKI$(RNO) + String$(ZL, 0) + Space$(XL)
      Put #3, RBA&, AF$
      GoTo FF3GetIT
   End If
FF3GetEND:
End Sub

Sub FF1Open(OPT$)
    OKAY = 0: D$ = FDFSetup.NFDRV: CA$ = ""
    QUIET = (Right$(OPT$, 1) = "Q")
    X = Len(RTrim$(OPT$)) + QUIET
    XOPT$ = MID$(OPT$, 1, X)
    If (("A" <= D$) And (D$ <= "Z")) And (D$ <> FDFSetup.FAMDRV) Then
       DSNAME$ = RTrim$(D$ + ":" + FT$(1))
       Else
       DSNAME$ = RTrim$(RTrim$(FAMFPFX$) + FT$(1))
    End If
    LSet SysVAR.FILENAME = DSNAME$
    On Local Error GoTo NOFF1
CHKFF1:
    Open DSNAME$ For Input As #1
    Close #1
    If XOPT$ = "INIT" Then
       If Not QUIET Then Call PutMSG("XXXXM051")
       If CA$ <> SNGLKEY$(1) Then GoTo FF1OpenEND
       GoTo OLDFF1
    End If
    If XOPT$ = "NEW" Then
       If Not QUIET Then Call PutMSG("XXXXM002")
       If CA$ <> SNGLKEY$(1) Then GoTo FF1OpenEND
       Kill SysVAR.FILENAME
       GoTo NEWFF1
    End If
    GoTo OLDFF1
NOFF1:
    X = ERR
    If (XOPT$ = "NEW") Then Resume NEWFF1
    If (XOPT$ = "INIT") Then OKAY = -1: Resume FF1OpenEND
    If QUIET Then Resume FF1OpenEND
    Call ERRMessage("XXXXM030", X, 0)
    If (CA$ = Chr$(27)) Or (X <> 71) Then Resume FF1OpenEND
    Resume CHKFF1
OLDFF1:
    On Local Error GoTo FF1OpenERR
    Select Case XOPT$
       Case "INIT": Open SysVAR.FILENAME For Binary Access Read As #1
       Case "READ": Open SysVAR.FILENAME For Binary Access Read As #1
       Case "WRITE": Open SysVAR.FILENAME For Binary Access Read Write As #1
       Case Else: GoTo FF1OpenEND
    End Select
    Get #1, , FF1Hdr
    LSet SysVAR.MVAR = "NAME"
    If FF1Hdr.FType <> "N" Then
       If Not QUIET Then Call PutMSG("XXXXM033")
       OKAY = 0: GoTo FF1OpenEND
    End If
    If ((FF1Hdr.REVLVL And 254) <> 0) Then
       Call PutMSG("XXXXM036")
       If CA$ <> SNGLKEY$(1) Then GoTo FF1OpenEND
       If XOPT$ = "WRITE" Then
          FF1Hdr.REVLVL = (FF1Hdr.REVLVL And 3) Or 256
          Else
          FF1Hdr.REVLVL = (FF1Hdr.REVLVL And 3)
       End If
    End If
    OKAY = -1
    GoTo FF1OpenEND
NEWFF1:
    Call PutMSG("XXXXM056")
    If CA$ = Chr$(27) Then GoTo FF1OpenEND
    ENV.LPNSupport = (CA$ = SNGLKEY$(1))
    On Local Error GoTo FF1OpenERR
    Open SysVAR.FILENAME For Binary Access Read Write As #1
    xx$ = String$(100, 0): Put #1, 1, xx$: Get #1, 1, FF1Hdr
    FF1Hdr.FType = "N"
    FF1Hdr.REVLVL = -ENV.LPNSupport
    FF1Hdr.REORGY = DATEY(DATE$, 0)
    FF1Hdr.REORGMD = DATEMD(DATE$, 0)
    FF1Hdr.REORGTM = TIMEHHMM(Time$)
    FF1Hdr.UPDTY = FF1Hdr.REORGY
    FF1Hdr.UPDTMD = FF1Hdr.REORGMD
    FF1Hdr.CHGDTY = FF1Hdr.REORGY
    FF1Hdr.CHGDTMD = FF1Hdr.REORGMD
    Put #1, 1, FF1Hdr
    OKAY = -1
    GoTo FF1OpenEND
FF1OpenERR:
    If Not QUIET Then Call ERRMessage("XXXXM030", ERR, ERL)
    Resume FF1OpenEND
FF1OpenEND:
    ENV.LPNSupport = (FF1Hdr.REVLVL > 0)
End Sub

Sub FF1PutREC()
    HIRBA& = FFRBA&(FF1Hdr.MAXID, 1)
    RBA& = FFRBA&(FF1Rec.RID, 1)
    If (RBA& > 0) And (RBA& <= HIRBA&) Then
       RBA& = RBA& * 100 + 1
       X = Asc(FF1Rec.STATUS2)
       If ((X And 96) = 0) Then FF1Hdr.CHGCNT2 = FF1Hdr.CHGCNT2 + 1
       If XCHANGED Then               ' Turn on record Change Bit
          X = X Or 32: XCHANGED = 0
          Else
          If ((X And 64) = 0) Then FF1Hdr.CHGCNT1 = FF1Hdr.CHGCNT1 + 1
          X = X Or 64
       End If
       FF1Rec.STATUS2 = Chr$(X)
       Put #1, RBA&, FF1Rec
       If FF1Hdr.REVLVL < 256 Then
          FF1Hdr.REVLVL = FF1Hdr.REVLVL Or 256
          Put #1, 1, FF1Hdr
       End If
       OKAY = -1: FFCHANGED = -1
       Else
       Call ErrBEEP(0): OKAY = 0
    End If
End Sub

Sub FF2DelRec(RNUM)
    HIRBA& = FFRBA&(FF2Hdr.HIREC, 1)
    RNO& = FFRBA&(RNUM, 1)
    If (RNO& > 0) And (RNO& <= HIRBA&) Then
       AFMT = FF2Hdr.REVLVL And 255
       If AFMT = 0 Then RL = 108 Else RL = 152
       RBA& = RNO& * RL + 1
       Get #2, RBA&, FF2PFX
       If FF2PFX.RTYPE <> Chr$(255) Then
          FF2PFX.RTYPE = Chr$(255)
          FF2PFX.NXT = FF2Hdr.FREEBGN
          Put #2, RBA&, FF2PFX
          FF2Hdr.FREEBGN = RNUM
          FF2Hdr.FREECNT = FF2Hdr.FREECNT + 1
          If FF2Hdr.REVLVL < 256 Then
             FF2Hdr.REVLVL = 256 + AFMT
             Put #2, 1, FF2Hdr
          End If
          OKAY = -1: XCHANGED = -1: FFCHANGED = -1
       End If
       Else
       OKAY = 0
    End If
End Sub

Sub FF2GetRec(RNUM)
   OKAY = 0
   If Not ENV.FFOPEN Then GoTo FF2GetEND
   AFMT = FF2Hdr.REVLVL And 255
   If AFMT = 0 Then RL = 108 Else RL = 152
   HIREC& = FFRBA&(FF2Hdr.HIREC, 1)
   If RNUM = 0 Then GoTo FF2GetNew
   RNO& = FFRBA&(RNUM, 1)
   If (RNO& < 1) Or (RNO& > HIREC&) Then
      SysVAR.MVAR = LTrim$(Str$(RNO&))
      Call PutMSG("XXXXM052")
      GoTo FF2GetEND
   End If
   RBA& = RNO& * RL + 1
   Get #2, RBA&, FF2PFX
   Select Case AFMT
      Case 0: Get #2, , FF2SREC
      Case Else: Get #2, , FF2LREC
   End Select
   OKAY = -1
   GoTo FF2GetEND
FF2GetNew:
   If FF2Hdr.FREECNT > 0 Then
      RBA& = FFRBA&(FF2Hdr.FREEBGN, RL) + 1
      If RBA& > 1 Then
         Get #2, RBA&, FF2PFX
         FF2Hdr.FREEBGN = FF2PFX.NXT
         FF2Hdr.FREECNT = FF2Hdr.FREECNT - 1
         Else
         RBA& = 0
      End If
      Else
      RBA& = 0
   End If
   If RBA& = 0 Then
      HIREC& = HIREC& + 1
      xHIREC = FFRNUM(HIREC&, 1)
      If xHIREC = 0 Then GoTo FF2GetEND
      FF2Hdr.HIREC = xHIREC
      RBA& = HIREC& * RL + 1
   End If
   xRNO = FFRNUM(RBA& - 1, RL)
   FF2Hdr.UPDTY = DATEY(DATE$, 0): FF2Hdr.UPDTMD = DATEMD(DATE$, 0)
   FF2Hdr.REVLVL = 256 + AFMT: Z = 18 - 6 * (AFMT = 0)
   AF$ = String$(Z, 0) + Space$(RL - Z)
   Mid$(AF$, 1, 3) = Chr$(2) + MKI$(xRNO)
   Put #2, RBA&, AF$
   Get #2, RBA&, FF2PFX
   FF2PFX.RNO = xRNO
   Select Case AFMT
      Case 0: Get #2, , FF2SREC
      Case Else: Get #2, , FF2LREC
   End Select
   Put #2, 1, FF2Hdr
   OKAY = -1
FF2GetEND:
End Sub

Rem $DYNAMIC
Sub FF2Open(OPT$)
    OKAY = 0: D$ = FDFSetup.NFDRV: CA$ = ""
    QUIET = (Right$(OPT$, 1) = "Q"): X = Len(RTrim$(OPT$)) + QUIET
    XOPT$ = MID$(OPT$, 1, X)
    If (("A" <= D$) And (D$ <= "Z")) And (D$ <> FDFSetup.FAMDRV) Then
       DSNAME$ = D$ + ":" + FT$(2)
       Else
       DSNAME$ = FAMFPFX$ + FT$(2)
    End If
    LSet SysVAR.FILENAME = DSNAME$
    On Local Error GoTo NOFF2
CHKFF2:
    Open DSNAME$ For Input As #2
    Close #2
    If XOPT$ = "INIT" Then
       If Not QUIET Then Call PutMSG("XXXXM051")
       If CA$ <> SNGLKEY$(1) Then GoTo FF2OpenEND
       GoTo OLDFF2
    End If
    If XOPT$ = "NEW" Then
       If Not QUIET Then Call PutMSG("XXXXM002")
       If CA$ <> SNGLKEY$(1) Then GoTo FF2OpenEND
       Kill SysVAR.FILENAME
       GoTo NEWFF2
    End If
    GoTo OLDFF2
NOFF2:
    X = ERR
    If (XOPT$ = "NEW") Then Resume NEWFF2
    If (XOPT$ = "INIT") Then OKAY = -1: Resume FF2OpenEND
    If Not QUIET Then Call ERRMessage("XXXXM030", X, 0)
    If (CA$ = Chr$(27)) Or (X <> 71) Then Resume FF2OpenEND
    Resume CHKFF2
OLDFF2:
    On Local Error GoTo FF2OpenERR
    Select Case XOPT$
       Case "INIT": Open SysVAR.FILENAME For Binary Access Read As #2
       Case "READ": Open SysVAR.FILENAME For Binary Access Read As #2
       Case "WRITE": Open SysVAR.FILENAME For Binary Access Read Write As #2
       Case Else: GoTo FF2OpenEND
    End Select
    Get #2, , FF2Hdr
    LSet SysVAR.MVAR = "ADDRESS"
    If FF2Hdr.FType <> "A" Then
       Call PutMSG("XXXXM033")
       If CA$ <> SNGLKEY$(1) Then GoTo FF2OpenEND
       FF2Hdr.FType = "A"
    End If
    AFMT = FF2Hdr.REVLVL And 255
    If (AFMT > 1) Then
       Call PutMSG("XXXXM036")
       If CA$ <> SNGLKEY$(1) Then GoTo FF2OpenEND
       If XOPT$ = "WRITE" Then
          FF2Hdr.REVLVL = AFMT + 256
          Else
          FF2Hdr.REVLVL = AFMT
       End If
    End If
    If (FF2Hdr.REORGY <> FF1Hdr.REORGY) Or (FF2Hdr.REORGMD <> FF1Hdr.REORGMD) Or (FF2Hdr.REORGTM <> FF1Hdr.REORGTM) Then
       Call PutMSG("XXXXM034")
       If CA$ <> SNGLKEY$(1) Then GoTo FF2OpenEND
       FF2Hdr.REORGY = FF1Hdr.REORGY: FF2Hdr.REORGMD = FF1Hdr.REORGMD
       FF2Hdr.REORGTM = FF1Hdr.REORGTM
       If XOPT$ = "WRITE" Then FF2Hdr.REVLVL = FF2Hdr.REVLVL Or 256
    End If
    OKAY = -1
    GoTo FF2OpenEND
NEWFF2:
    Call PutMSG("XXXXM053")
    If CA$ = Chr$(27) Then GoTo FF2OpenEND
    Select Case CA$
       Case "1": AFMT = 0
       Case "2": AFMT = 1
       Case Else: Call ErrBEEP(0): GoTo NEWFF2
    End Select
    On Local Error GoTo FF2OpenERR
    Open SysVAR.FILENAME For Binary Access Read Write As #2
    FF2Hdr.FType = "A"
    FF2Hdr.REVLVL = AFMT
    FF2Hdr.HIREC = 0
    FF2Hdr.REORGY = FF1Hdr.REORGY
    FF2Hdr.REORGMD = FF1Hdr.REORGMD
    FF2Hdr.REORGTM = FF1Hdr.REORGTM
    FF2Hdr.UPDTY = FF1Hdr.UPDTY
    FF2Hdr.UPDTMD = FF1Hdr.UPDTMD
    FF2Hdr.FREEBGN = 0: FF2Hdr.FREECNT = 0
    Put #2, 1, FF2Hdr
    xx$ = String$(89 - 44 * (AFMT = 1), 0)
    Put #2, , xx$
    OKAY = -1
    GoTo FF2OpenEND
FF2OpenERR:
    Call ERRMessage("XXXXM030", ERR, ERL)
    Resume FF2OpenEND
FF2OpenEND:
End Sub

Rem $STATIC
Sub FF2PutRec()
    HIRBA& = FFRBA&(FF2Hdr.HIREC, 1)
    RBA& = FFRBA&(FF2PFX.RNO, 1)
    If (RBA& > 0) And (RBA& <= HIRBA&) Then
       AFMT = FF2Hdr.REVLVL And 255
       If AFMT = 0 Then RL = 108 Else RL = 152
       RBA& = RBA& * RL + 1
       Put #2, RBA&, FF2PFX
       XCHANGED = -1
       Select Case AFMT
          Case 0: Put #2, , FF2SREC
          Case 1: Put #2, , FF2LREC
       End Select
       If FF2Hdr.REVLVL < 256 Then
          FF2Hdr.REVLVL = 256 + AFMT
          Put #2, 1, FF2Hdr
       End If
       OKAY = -1: FFCHANGED = -1: XCHANGED = -1
       Else
       Call ErrBEEP(0): OKAY = 0
    End If
End Sub

Sub FF3DelRec(RNUM)
    HIRBA& = FFRBA&(FF3HDR.HIREC, 1)
    RBA& = FFRBA&(RNUM, 1)
    If (RBA& > 0) And (RBA& <= HIRBA&) Then
       RBA& = RBA& * 50 + 1
       Get #3, RBA&, FF3PFX
       If FF3PFX.RTYPE <> Chr$(255) Then
          FF3PFX.RTYPE = Chr$(255)
          Put #3, RBA&, FF3PFX
          Get #3, RBA& + 3, FF3PNTRS
          FF3PNTRS.NXT = FF3HDR.FREEBGN
          Put #3, RBA& + 3, FF3PNTRS
          FF3HDR.FREEBGN = RNUM
          FF3HDR.FREECNT = FFRNUM(FFRBA&(FF3HDR.FREECNT, 1) + 1, 1)
          If FF3HDR.REVLVL < 256 Then
             FF3HDR.REVLVL = FF3HDR.REVLVL Or 256
             Put #3, 1, FF3HDR
          End If
          OKAY = -1: XCHANGED = -1: FFCHANGED = -1
       End If
       Else
       OKAY = 0
    End If
End Sub

Rem $DYNAMIC
Function FFRBA&(RNUM, RLTH)
   X& = RNUM
   ' New Routine - After May 20, 1997
     If X& < 0 Then X& = 65535 + X&: If X& < 0 Then X& = 0
   ' Old Routine - Prior to May 21, 1997
   ' IF X& < 0 THEN X& = 32766 - X&
   FFRBA& = X& * RLTH
End Function

Sub FF3PutRec()
    OKAY = 0
    HIRBA& = FFRBA&(FF3HDR.HIREC, 1)
    RBA& = FFRBA&(FF3PFX.RNO, 1)
    If (RBA& > 0) And (RBA& <= HIRBA&) Then
       RBA& = RBA& * 50 + 1
       Put #3, RBA&, FF3PFX
       Select Case FF3PFX.RTYPE
          Case Chr$(3): Put #3, , FF3SPOUSE
          Case Chr$(4): Put #3, , FF3PLACE
          Case Chr$(5): Put #3, , FF3COMMENT
          Case Chr$(6): Put #3, , FF3EVENT
          Case Chr$(11): Put #3, , FF3EVENT     ' Old EVENT Records
          Case Chr$(12): Put #3, , FF3PLACE2
          Case Else: Put #3, , FF3PNTRS
                     Select Case FF3PFX.RTYPE
                        Case Chr$(7): Put #3, , FF3EDUC
                        Case Chr$(8): Put #3, , FF3WORK
                        Case Chr$(9): Put #3, , FF3MIL
                        Case Chr$(10): Put #3, , FF3HEALTH
                        Case Chr$(255):
                        Case Else: GoTo FF3PutEND
                     End Select
       End Select
       If FF3HDR.REVLVL < 256 Then
          FF3HDR.REVLVL = FF3HDR.REVLVL Or 256
          Put #3, 1, FF3HDR
       End If
       OKAY = -1: XCHANGED = -1: FFCHANGED = -1
    End If
FF3PutEND:
   If Not OKAY Then Call ErrBEEP(0)
End Sub


Function FFRNUM(RBA&, RLTH)
   If RLTH > 0 Then X& = RBA& / RLTH Else X& = RBA&
   ' New Routine - After May 20,1997
     If X& > 32767 Then X& = X& - 65535: If X& > 0 Then X& = 0
     If X& < -32767 Then X& = 0
   ' Old Routine - prior to May 21, 1997
   ' IF X& > 32766 THEN X& = 32766 - X&
   ' IF X& < -32766 THEN X& = 0
   FFRNUM = X&
End Function

Function FFSTAT$(ST$, S$)
   X = Asc(FF1Rec.STATUS)
   Select Case ST$
      Case "FID": X = (X And (255 - 3))
                  Select Case S$
                    Case "*": X = X Or 1
                    Case "?": X = X Or 2
                  End Select
      Case "MID": X = (X And (255 - 12))
                  Select Case S$
                    Case "*": X = X Or 4
                    Case "?": X = X Or 8
                  End Select
      Case "BD": X = (X And (255 - 48))
                 Select Case S$
                   Case "?": X = X Or 16
                   Case "!": X = X Or 32
                 End Select
      Case "DD": X = (X And (255 - 192))
                 Select Case S$
                   Case "?": X = X Or 64
                   Case "!": X = X Or 128
                 End Select
      Case "SNU": X = Asc(FF1Rec.STATUS2) And 127: If S$ = "Y" Then X = X Or 128
   End Select
   FFSTAT$ = Chr$(X)
End Function

Rem $STATIC
Function FMTNAME$(XNFMT)
   NFMT = XNFMT: If NFMT = 0 Then NFMT = RptOPTION.NAMEFMT + 1
   LASTFIRST = ((NFMT = 2) Or (NFMT = 4))
   UPCASE = ((NFMT = 3) Or (NFMT = 4))
   GIVEN$ = LTrim$(RTrim$(FF1Rec.GIVEN))
   SURNM$ = LTrim$(RTrim$(FF1Rec.SURNM))
   If ENV.HUSBSURNM And (FF1Rec.SEX = FGENDR$(2)) Then
      If (FF1Rec.SPOUSE <> 0) And ((Asc(FF1Rec.STATUS2) And 128) = 0) Then
         SPID = 0
         Call FF3GetRec(FF1Rec.SPOUSE)
         If FF1Rec.RID = FF3SPOUSE.SP1ID Then
            SPID = FF3SPOUSE.SP2ID
            Else
            SPID = FF3SPOUSE.SP1ID
         End If
         If SPID > 0 Then
            RID = FF1Rec.RID: Call FF1GetRec(SPID)
            SURNM$ = LTrim$(RTrim$(FF1Rec.SURNM))
            Call FF1GetRec(RID)
            GIVEN$ = GIVEN$ + " " + LTrim$(RTrim$(FF1Rec.SURNM))
         End If
      End If
   End If
   If UPCASE Then SURNM$ = UCX$(SURNM$)
   If LASTFIRST Then
     FMTNAME$ = SURNM$ + ", " + GIVEN$
     Else
     FMTNAME$ = GIVEN$ + " " + SURNM$
   End If
End Function

Rem $DYNAMIC
Sub INDEXFile(OPT$)
    QUIET = (Right$(OPT$, 1) = "Q"): X = Len(RTrim$(OPT$)) + QUIET
    XOPT$ = MID$(OPT$, 1, X)
    OKAY = 0: D$ = FDFSetup.NFDRV
    If (("A" <= D$) And (D$ <= "Z")) And (D$ <> FDFSetup.FAMDRV) Then
       NDXPFX$ = D$ + ":"
       Else
       NDXPFX$ = FAMFPFX$
    End If
GETNDXNM:
    If Not QUIET Then
       'Call FmtFIND("XXXXS012")
       'Call FmtFindFLD("DSNM"): LSet UTXT$(FFLD.UTXT) = FT$(4)
       'Call MsgFORMAT: If A = 27 Then GoTo NDXOpenBAD
       'Call FmtFindFLD("DSNM"): LSet FT$(4) = UTXT$(FFLD.UTXT)
    End If
    DSNAME$ = RTrim$(NDXPFX$ + FT$(4))
    LSet SysVAR.FILENAME = DSNAME$
    On Local Error GoTo NONDX
CHKNDX:
    Open DSNAME$ For Input As #4
    Close #4
    If XOPT$ = "CREATE" Then
       CA$ = ""
       While CA$ <> SNGLKEY$(1)
          Call PutMSG("XXXXM002")
          If (CA$ = SNGLKEY$(2)) Or (CA$ = Chr$(27)) Then GoTo NDXOpenBAD
       Wend
       Kill SysVAR.FILENAME
       GoTo NEWNDX
    End If
    GoTo OLDNDX
NONDX:
    X = ERR
    If (XOPT$ = "CREATE") Then Resume NEWNDX
    If QUIET Then Resume NDXOpenBAD
    Call ERRMessage("XXXXM030", X, 0)
    'IF CA$ = CHR$(27) THEN RESUME GETNDXNM
    'RESUME CHKNDX
    Resume NDXOpenBAD
OLDNDX:
    On Local Error GoTo NDXOpenERR
    Select Case XOPT$
       Case "CHECK": Open SysVAR.FILENAME For Binary Access Read As #4
       Case "READ": Open SysVAR.FILENAME For Binary Access Read As #4
       Case "WRITE": Open SysVAR.FILENAME For Binary Access Read Write As #4
       Case Else: GoTo NDXOpenBAD
    End Select
    Get #4, , NDXHDR
    LSet SysVAR.MVAR = "INDEX"
    If NDXHDR.FType <> "X" Then
       If Not QUIET Then Call PutMSG("XXXXM033")
       GoTo NDXOpenBAD
    End If
    If NDXHDR.VER <> Chr$(1) Then
       If Not QUIET Then Call PutMSG("XXXXM037")
       GoTo NDXOpenBAD
    End If
    If (NDXHDR.REORGY <> FF1Hdr.REORGY) Or (NDXHDR.REORGMD <> FF1Hdr.REORGMD) Or (NDXHDR.REORGTM <> FF1Hdr.REORGTM) Then
       If Not QUIET Then Call PutMSG("XXXXM034")
       GoTo NDXOpenBAD
    End If
    If (XOPT$ = "READ") And (FF1Hdr.MAXID > NDXHDR.NDXRECS) Then
       Call PutMSG("XXXXM074")
       If CA$ <> SNGLKEY$(1) Then GoTo NDXOpenBAD
    End If
    OKAY = -1
    SFCNT = NDXHDR.OPT And 31
    ReDim SX(12)
    X = 0: While X < SFCNT: X = X + 1: Get #4, , SX(X): Wend
    ENV.HUSBSURNM = ((NDXHDR.OPT And 32) > 0)
    ENV.SortOPTS = (NDXHDR.OPT \ 64)
    If XOPT$ = "CHECK" Then GoTo NDXOpenEND
    If NDXHDR.NDXRECS > FF1Hdr.MAXID Then NDXHDR.NDXRECS = FF1Hdr.MAXID
    ReDim NDXID(FF1Hdr.MAXID)
    C = 0: X = 0
    While X < NDXHDR.NDXRECS
       X = X + 1: Get #4, , XID
       If XID <= FF1Hdr.MAXID Then C = C + 1: NDXID(C) = XID
    Wend
    GoTo NDXOpenEND
NEWNDX:
    On Local Error GoTo NDXOpenERR
    If ENV.SortSTAT Then
       SortSTAT!(10) = 0: Call SortStatSHOW: HP5 = SortPHASE
       SortPHASE = 5: Call SortStatACCUM
    End If
    Open SysVAR.FILENAME For Binary Access Read Write As #4
    NDXHDR.FType = "X"
    NDXHDR.VER = Chr$(1)
    NDXHDR.OPT = (SFCNT - 32 * ENV.HUSBSURNM + 64 * ENV.SortOPTS)
    NDXHDR.REORGY = FF1Hdr.REORGY
    NDXHDR.REORGMD = FF1Hdr.REORGMD
    NDXHDR.REORGTM = FF1Hdr.REORGTM
    NDXHDR.UPDTY = Val(Right$(DATE$, 4))
    NDXHDR.UPDTMD = 100 * (Val(Left$(DATE$, 2))) + Val(MID$(DATE$, 4, 2))
    NDXHDR.NDXRECS = FF1Hdr.MAXID
    Put #4, 1, NDXHDR
    xx$ = " ": X = 0
    While X < SFCNT: X = X + 1: Put #4, , SX(X): Wend
    X = 0
    While X < FF1Hdr.MAXID
       X = X + 1: Put #4, , NDXID(X)
       If ENV.SortSTAT And (X Mod 100 = 0) Then
          SortSTAT!(10) = X: Call SortStatSHOW
       End If
    Wend
    If ENV.SortSTAT Then
       SortSTAT!(10) = X: Call SortStatSHOW: SortPHASE = HP5
    End If
    OKAY = -1
    GoTo NDXOpenEND
NDXOpenERR:
    Call ERRMessage("XXXXM030", ERR, ERL)
    Resume NDXOpenBAD
NDXOpenBAD:
    OKAY = 0
NDXOpenEND:
    If Not OKAY Then Call FamCLOSE Else Close #4
End Sub
Public Sub SortStatSHOW()
    Call SortStatACCUM
    frmFHSWINDX.SPValue(8) = LTrim$(Str$(SS))
    frmFHSWINDX.SPValue(9) = MID$(Str$(SortSTAT!(1)), 2, XL)
    frmFHSWINDX.SPValue(10) = MID$(Str$(SortSTAT!(2)), 2, XL)
    frmFHSWINDX.SPValue(11) = MID$(Str$(SortSTAT!(3)), 2, XL)
    frmFHSWINDX.SPValue(12) = MID$(Str$(SortSTAT!(4)), 2, XL)
    frmFHSWINDX.SPValue(10) = MID$(Str$(SortSTAT!(5)), 2, XL)
    frmFHSWINDX.SPValue(11) = LTrim$(Str$(SortSTAT!(6)))
    frmFHSWINDX.SPValue(12) = LTrim$(Str$(SortSTAT!(7)))
    frmFHSWINDX.SPValue(13) = LTrim$(Str$(SortSTAT!(8)))
    frmFHSWINDX.SPValue(14) = LTrim$(Str$(SortSTAT!(9)))
    'frmFHSWINDX.SPValue(15) = LTrim$(Str$(SortSTAT!(7) + SortSTAT!(8) + SortSTAT!(9)))
    frmFHSWINDX.SPValue(15) = LTrim$(Str$(SortSTAT!(10)))
    frmFHSWINDX.SPValue(17) = Time$
    SortSTAT!(0) = Timer
End Sub
Public Sub SortStatACCUM()
    If SortPHASE <> 0 Then
       SortSTAT!(SortPHASE) = SortSTAT!(SortPHASE) + Timer - SortSTAT!(0)
    End If
    SortSTAT!(0) = Timer
End Sub





Public Sub SX2SFLD()
580 ' Transfer SX() data to SFLD() array
    For ix = 1 To 9: SFLD(ix) = SFLD(ix) And 255: Next ix
    For NX = 1 To SFCNT
       ix = (SX(NX) \ 256)
       SFLD(ix) = (256 * NX) + (SX(NX) And 128) + (SFLD(ix) And 127)
       If (ix < 3) Or ((ix > 5) And (ix < 8)) Then
          SFLD(ix) = SFLD(ix) - (SFLD(ix) And 127) + (SX(NX) And 127)
       End If
    Next NX
End Sub

Public Sub SFLD2SX()
550 ' Transfer SFLD() to SX()
    ReDim SX(12)
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
End Sub

Sub GetBDEvents(BD$, BPL$, DD$, DPL$, AGE$)
     ' Get Birth(Baptism)/Death(Burial)/Age Fields
     D1ST$ = FDSTAT$("BD"): D2ST$ = FDSTAT$("DD")
     LSet BPLACE = "": LSet DPLACE = ""
     If BPL$ <> "" Then
        If FF1Rec.BLOC <> 0 Then
           Call FF3GetRec(FF1Rec.BLOC)
           If FF3PFX.RTYPE = Chr$(12) Then
              If FF3PLACE2.PNEXT <> 0 Then
                 Call FF3GetRec(FF3PLACE2.PNEXT)
              End If
           End If
        End If
     End If
     EVNTPTR = FF1Rec.EVENT
     While ((FF1Rec.BY = 0) Or (FF1Rec.DY = 0)) And (EVNTPTR <> 0)
        Call FF3GetRec(EVNTPTR)
        If FF1Rec.BY = 0 Then
           Select Case FF3EVENT.CODE
              Case "BIRTH   ": FF1Rec.BY = FF3EVENT.D1Y
                               FF1Rec.BMD = FF3EVENT.D1MD
                               LSet BPLACE = FF3EVENT.PLACE
                               D1ST$ = FF3EVENT.DSTAT
              Case "BAPTISM ": FF1Rec.BY = FF3EVENT.D1Y
                               FF1Rec.BMD = FF3EVENT.D1MD
                               LSet BPLACE = FF3EVENT.PLACE
                               D1ST$ = "+"
           End Select
        End If
        If FF1Rec.DY = 0 Then
           Select Case FF3EVENT.CODE
              Case "DEATH   ": FF1Rec.DY = FF3EVENT.D1Y
                               FF1Rec.DMD = FF3EVENT.D1MD
                               DPLACE = FF3EVENT.PLACE
                               D2ST$ = FF3EVENT.DSTAT
              Case "BURIAL  ": FF1Rec.DY = FF3EVENT.D1Y
                               FF1Rec.DMD = FF3EVENT.D1MD
                               DPLACE = FF3EVENT.PLACE
                               D2ST$ = "+"
           End Select
        End If
        EVNTPTR = FF3EVENT.NXT
     Wend
     If BD$ <> "" Then
        Select Case Len(BD$)
           Case 11: LSet BD$ = XSCRNDate$(FF1Rec.BMD, FF1Rec.BY): Mid$(BD$, 11, 1) = D1ST$
                    LSet DD$ = XSCRNDate$(FF1Rec.DMD, FF1Rec.DY): Mid$(DD$, 11, 1) = D2ST$
           Case 12: LSet BD$ = XRPTDate$(FF1Rec.BMD, FF1Rec.BY, 1): Mid$(BD$, 12, 1) = D1ST$
                    LSet DD$ = XRPTDate$(FF1Rec.DMD, FF1Rec.DY, 1): Mid$(DD$, 12, 1) = D2ST$
        End Select
     End If
     If BPL$ <> "" Then
        LSet BPL$ = BPLACE: LSet DPL$ = DPLACE
     End If
     If AGE$ <> "" Then
        RSet AGE$ = DATEDif$(FF1Rec.BMD, FF1Rec.BY, FF1Rec.DMD, FF1Rec.DY)
     End If
End Sub


Function LineCount(X)
' Convert x=inches*100 to "Inches" of lines
Printer.ScaleMode = 5
xx# = Printer.TextHeight(StdHeight) / 10
LineCount = Int((X / xx#) / 100)
End Function


Sub PrintINIT()
X = (PDFSetup.OPTIONS And 48) / 16
If (X < 1) Or (X > 2) Then X = 1
On Local Error GoTo errOrient
Printer.Orientation = X
GoTo setFONT
errOrient: Resume Next
setFONT:
On Local Error GoTo 0
Printer.Font = PFONT$(ENV.PSETUP)
Printer.Font.SIZE = PDFSetup.FONTSIZE
Printer.Font.Bold = ((PDFSetup.FONTSTYL And 1) > 0)
Printer.Font.Italic = ((PDFSetup.FONTSTYL And 2) > 0)
End Sub

Rem $STATIC
Sub PrintLINE()
2100 ' Print Line of DATA
     If a <> 27 Then
        PAGE.LOFST = Len(RTrim$(PLINE$))
        If PAGE.LOFST > 0 Then
           PAGE.RPAGE = 0: PAGE.RLINE = 0: Call WriteLINE
           If RptOPTION.BLDREF Then
              If RREFRPT.BGNPAGE = 0 Then RREFRPT.BGNPAGE = PAGE.COUNT
              If (PAGE.XREF > 0) Then Call RREFBld
           End If
        End If
     End If
     PAGE.LOFST = PAGE.COFST: PAGE.LWIDTH = PAGE.BWIDTH
End Sub



Sub PrintOPEN()
    MD = DATEMD(DATE$, 0): Y = DATEY(DATE$, 0)
    SysVAR.DATE = XRPTDate$(MD, Y, 0)
    If (ENV.SYSMODE = 0) Or (PRT.OPT = 1) Then
       RptOPTION.BLDREF = 0
       Else
       If RptOPTION.BLDREF Then
          Call PutMSG("XXXXM061"): If a = 27 Then GoTo 1390
          RptOPTION.BLDREF = (CA$ = SNGLKEY$(1))
          If RptOPTION.BLDREF Then
             If RptOPTION.SNGLREF Then
                Call PutMSG("XXXXM063"): If a = 27 Then GoTo 1390
                RptOPTION.SNGLREF = Val(CA$) - 1
             End If
             Call RREFOpen("WRITE"): If (Not OKAY) Or (a = 27) Then GoTo 1390
          End If
       End If
    End If
    Call PageSetup
    Call RptOptGET("CMIN", X)
    Call RptOptGET("FOFS", X): PRT.FOFFSET = X
    Call RptOptGET("CFRM", X): PRT.CFORMS = (X = 1)
    'PRT.DEST = PDFSetup.DEST
    Call RptOptGET("APST", X)
    PRT.ALL = (X = 4): PRT.AGAIN = ((X = 1) Or (PRT.ALL))
    PAGE.LOFST = 1: PAGE.COFST = 1
    PAGE.RLINE = 0: PAGE.RPAGE = 0: PAGE.WAIT = PRT.PAGEWAIT
    PRT.EOF = 0
    LF$ = ""
    If PAGE.BWIDTH > 0 Then
       PAGE.WIDTH = PAGE.BWIDTH + PAGE.IMARGIN + PAGE.OMARGIN
       Else
       PAGE.WIDTH = PRT.FWIDTH
       PAGE.BWIDTH = PAGE.WIDTH - PAGE.IMARGIN - PAGE.OMARGIN
    End If
    PAGE.LWIDTH = PAGE.BWIDTH
    If PRT.OPT <> 2 Then PRT.FWIDTH = PAGE.WIDTH
    If PAGE.WIDTH < 32712 Then
       PLINE$ = Space$(PAGE.WIDTH + 1)
       Else
       PLINE$ = Space$(255)
    End If
    CL$ = PLINE$ + Space$(55)
    Call RptOptGET("LREF", X)
    Call RptOptGET("REFL", X)
    If Not PRT.SNGLSIDE Then
       Call PutMSG("XXXXM072")
       If Asc(CA$) = 27 Then GoTo 1390
       PRT.SNGLSIDE = Val(CA$)
       If (PRT.SNGLSIDE < 1) Or (PRT.SNGLSIDE > 2) Then
          PRT.SNGLSIDE = 3
          Else
          PRT.SNGLSIDE = PRT.SNGLSIDE Xor 3
       End If
       Else
       If ((RptOPTION.LINEREF And 1) > 0) Then
          If ((RptOPTION.LINEREF And 24) > 0) Then
             RptOPTION.LINEREF = 1 + (RptOPTION.LINEREF \ 4)
          End If
       End If
    End If
    'If PDFSetup.DEST = "SCRN:" Then Opt = 1: GoTo 1265
    'If PDFSetup.DEST = "FILE " Then PRT.Opt = 3: GoTo 1275
1260 'Call PutMSG("XXXXM008")
     'If Asc(CA$) = 27 Then GoTo 1390
     'Opt = Val(CA$): If Opt < 1 Or Opt > 3 Then Call ErrBEEP(0): GoTo 1260
1265 'PRT.Opt = Opt
     'If Opt = 1 Then PRT.DEST = "SCRN:" Else PRT.DEST = PDFSetup.DEST
     'If PRT.Opt <> 3 Then DEST$ = PRT.DEST: GoTo 1300
1275 ' File Output
     'Call FmtFIND("XXXXS011")
     'Call FmtFindFLD("FNAME"): NMFLD = WIN.CurFLD
     'LSet UTXT$(FLDTAB(NMFLD).UTXT) = FT$(14)
     'Call MsgFORMAT
     'Call frmFHSWFile.LOCATE("REPORT", "FILE")
     'If A = 27 Then GoTo 1260
     'LSet PRT.DEST = UTXT$(FLDTAB(NMFLD).UTXT)
     Select Case PRT.OPT
        Case 2: SysVAR.FILENAME = PRT.DEST
        Case Else: xx$ = ""
                xx$ = RPTFPFX$ + RTrim$(PRT.DEST)
                SysVAR.FILENAME = xx$
     End Select
     CRLF$ = Chr$(13) + Chr$(10)
1300 DEST$ = RTrim$(SysVAR.FILENAME)
     On Local Error GoTo 1330
     GoTo 1350
1330 ' Error Opening Print
     Call ERRMessage("XXXXM030", ERR, ERL)
     Resume 1390
1350 If PRT.OPT = 2 Then
        Call PrintINIT
        Else
        Open DEST$ For Output As #4
        If PRT.OPT = 3 Then
           LSet FT$(14) = PRT.DEST
        End If
     End If
     'If PRT.Opt = 2 Then
     '   Print #4, PrintCntl$(1); PrintCntl$(2);
     '   Else
     '   Print #4, Chr$(13);
     'End If
     On Local Error GoTo 0
     'If PRT.Opt = 1 Then
     '   If Not ENV.ScrnCLR Then Call ScrnCLR
     '   PRT.FWIDTH = 79: PRT.FLENGTH = 23
     '   PRT.PAGEWAIT = -1: PAGE.WAIT = -1
     '   PRT.FORMFEED = 0: PAGE.BLTOP = 0: PAGE.BLBOT = 0
     'End If
     'Call PutMSG("XXXX" + RDFReport.PRTMSG)
     OKAY = -1: PRT.OPEN = Chr$(255): Exit Sub
1390 ' Bad EXIT From Routine
     On Local Error GoTo 0
     a = 27: CA$ = Chr$(27): OKAY = 0
End Sub
Sub PutDATA()
2000 ' Place Data on Line
     If (PAGE.LOFST > (PAGE.LWIDTH + 1)) Or ((PAGE.LOFST = PAGE.LWIDTH) And (MID$(CL$, 1, 1) = " ")) Then
        Call PrintLINE: If a = 27 Then GoTo 2009
     End If
     If (RptOPTION.BLDREF) And (PAGE.XID > 0) Then Call RREFAdd
     If PAGE.LOFST = 0 Then PAGE.LOFST = PAGE.COFST
2005 ET = Len(RTrim$(CL$)): If ET < 1 Then GoTo 2009
     BT = 1
     If (PAGE.LOFST <= PAGE.COFST) Then
        C$ = MID$(CL$, 1, 1)
        While (BT < ET + 1) And ((C$ = " ") Or (C$ = ".") Or (C$ = ",") Or (C$ = ";"))
           BT = BT + 1: LSet C$ = MID$(CL$, BT, 1)
        Wend: If ET < BT Then LSet CL$ = "": GoTo 2009
     End If
2006 X = BT + PAGE.LWIDTH - PAGE.LOFST + 1
     If X <= ET Then XL$ = MID$(CL$, X, ET - X + 1): ET = X - 1
     Mid$(PLINE$, PAGE.LOFST, ET - BT + 1) = MID$(CL$, BT, ET - BT + 1)
     PAGE.LOFST = PAGE.LOFST + ET - BT + 1
2007 If (XL$ <> "") Then
        X = PAGE.LWIDTH: C$ = MID$(PLINE$, X, 1)
        While (X > PAGE.COFST) And (C$ <> " ") And (C$ <> ",") And (C$ <> ".") And (C$ <> ";")
           X = X - 1: LSet C$ = MID$(PLINE$, X, 1)
        Wend: If X = PAGE.COFST Then X = PAGE.LWIDTH
2008    CW = PAGE.LWIDTH - X: LSet CL$ = MID$(PLINE$, X + 1, CW)
        Mid$(CL$, CW + 1, 255) = XL$: XL$ = ""
        If CW Then Mid$(PLINE$, X + 1, CW) = Space$(CW)
        Call PrintLINE: If a = 27 Then GoTo 2009
        GoTo 2005
     End If
2009 ' Finished
     LSet CL$ = ""
End Sub

Function Relation$(RT$, RSEX$, AGL, RGL, RCODE$())
     RELWORK$ = Space$(50): L1 = 1
     LATREL = ((ENV.RELRULES And 2) = 2)
     Select Case RT$
        Case "S": x1 = 3: Y1 = 3: X2 = 3: Y2 = 3
        Case "L": x1 = 6: Y1 = 12: X2 = 6: Y2 = 12: LongREL = -1
        Case Else: x1 = 3: Y1 = 3: X2 = 6: Y2 = 12
     End Select
     If (RGL < AGL) Then LSet RELWORK$ = MID$(RCODE$(1), X2, Y2): GoTo 9549
     If (RGL = 0) And (AGL = 0) Then LSet RELWORK$ = MID$(RCODE$(2), X2, Y2): GoTo 9549
     If RSEX$ = FGENDR$(1) Then XGEN = 0 Else If RSEX$ = FGENDR$(2) Then XGEN = 1 Else XGEN = 2
     OGEN = XGEN + 2 * (XGEN > 1)
9505 ' Direct DESCENDANTS
     If AGL <> 0 Then GoTo 9510
     xGL = -RGL: xC = 6
     If LongREL And (RSEX$ = "X") Then xC = 33: XGEN = 0
     If LATREL And (xGL < -1) Then
        If xC = 6 Then xC = 28 Else xC = 34
     End If
     GoTo 9535
9510 ' Direct ANCESTORS
     If AGL <> RGL Then GoTo 9520
     xGL = AGL: xC = 3
     If LongREL And (RSEX$ = "X") Then xC = 31: XGEN = 0
     If LATREL And (xGL < -1) Then
        If xC = 3 Then xC = 25 Else xC = 32
     End If
     GoTo 9535
9520 ' Brothers and Sisters
     If AGL <> -1 Then GoTo 9530
     If RGL > 0 Then GoTo 9525
     xGL = AGL: xC = 9: GoTo 9535
9525 ' Nieces and Nephews
     xGL = -RGL: xC = 15: XG = 28: GoTo 9535
9530 ' Aunts and Uncles
     If (RGL - AGL) > 1 Then GoTo 9540
     xGL = AGL + 1: xC = 12: XG = 25
9535 ' Build Relation
     If LATREL Then
        If xGL < -2 Then Mid$(RELWORK$, L1, 3) = LTrim$(Str$(-(xGL + 1))) + "   "
        L1 = InStr(L1, RELWORK$, "  ") - LongREL
        Else
        If LongREL And (xGL < -3) Then
           Mid$(RELWORK$, L1, 3) = LTrim$(Str$(-(xGL + 2))) + "   "
           L1 = InStr(L1, RELWORK$, "  ") + 1
        End If
        If xGL < -2 Then
           Mid$(RELWORK$, L1, Y1) = MID$(RCODE$(22 + OGEN), x1, Y1) + "   "
           L1 = InStr(L1, RELWORK$, "  ") - LongREL
           If (xGL < -3) And (Not LongREL) Then
              Mid$(RELWORK$, L1, 3) = LTrim$(Str$(-(xGL + 2))) + "   "
              L1 = InStr(L1, RELWORK$, "  ")
           End If
        End If
        If (xGL < -1) Then
           Mid$(RELWORK$, L1, Y1) = MID$(RCODE$(20 + OGEN), x1, Y1)
           L1 = InStr(RELWORK$, "  ") - LongREL
        End If
     End If
9536 xZ$ = RCODE$(xC + XGEN)
     Mid$(RELWORK$, L1, Y2) = MID$(xZ$, X2, Y2) + "   "
     L1 = InStr(L1, RELWORK$, " ") - LongREL
     If LATREL And (XG > 0) Then
        If xGL < -1 Then
           Mid$(RELWORK$, L1, Y2) = MID$(RCODE$(XG + XGEN), X2, Y2) + "   "
        End If
     End If
     GoTo 9549
9540 ' Cousins
     CIVREL = ((ENV.RELRULES And 3) > 0)
     CGL = Abs(AGL + (RGL) * (RGL < 0) + 1)
     If CIVREL Then CGL = (2 * CGL + Abs(RGL) - 1)
     LSet RELWORK$ = LTrim$(Str$(CGL))
     L1 = InStr(RELWORK$, "  ") - LongREL
     If (RGL = 0) Or (CIVREL) Then
        Mid$(RELWORK$, L1, Y2) = MID$(RCODE$(18 + OGEN), X2, Y2) + "   "
        GoTo 9549
     End If
     Mid$(RELWORK$, L1, Y1) = MID$(RCODE$(18 + OGEN), x1, Y1) + "   "
     L1 = InStr(L1, RELWORK$, " ") - LongREL
     Mid$(RELWORK$, L1, 4) = LTrim$(Str$(Abs(RGL))) + "   "
     L1 = InStr(L1, RELWORK$, " ") - LongREL
     Mid$(RELWORK$, L1, Y1) = MID$(RCODE$(24), x1, Y1) + "   "
9549 L1 = InStr(RELWORK$, "  ")
     Relation$ = MID$(RELWORK$, 1, L1)
End Function


Rem $STATIC
Sub RptCLOSE()
    Call ClearMSG
    If a <> 27 Then
       Call EvenBREAK("EOF")
       If Len(RTrim$(ENV.RTITLE)) > 0 Then
          Call HTFExpand(ENV.RTITLE, 1)
       End If
       If RptOPTION.BLDREF Then Call RREFClose: RptOPTION.BLDREF = 0
       Else
       Call PutMSG("XXXX" + RDFReport.INTMSG)
    End If
    If (a <> 27) And (PRT.AGAIN) Then
       If PRT.OPT = 2 Then
          PRT.AGAIN = ((PRT.FOFFSET + PRT.FWIDTH) < PAGE.WIDTH)
          Else
          PRT.AGAIN = 0
       End If
       PRT.ALL = (PRT.AGAIN And PRT.ALL)
       While INKEY$ <> "": Wend      ' Clear Key Buffer
       If (PRT.AGAIN) And (Not PRT.ALL) Then
          PRT.FOFFSET = PRT.FOFFSET + PRT.FWIDTH
          SysVAR.NVAR1 = PRT.FOFFSET
          Call PutMSG("XXXXM007")
          PRT.ALL = (CA$ = SNGLKEY$(4))
          PRT.AGAIN = ((CA$ = SNGLKEY$(1)) Or PRT.ALL)
       End If
       If (Not PRT.AGAIN) And (PRT.SNGLSIDE = 2) Then
          Call PutMSG("XXXXM081")
          PRT.AGAIN = (CA$ = SNGLKEY$(1))
          If PRT.AGAIN Then
             Call RptOptGET("FOFS", X): PRT.FOFFSET = X
             PRT.SNGLSIDE = 1
          End If
       End If
       Else
       PRT.AGAIN = 0: PRT.ALL = 0
    End If
    If PRT.AGAIN Then
       PAGE.COUNT = PAGE.BGNPAGE - 1
       Call PutMSG("XXXX" + RDFReport.PRTMSG)
       Else
       Call PrintCLOSE
    End If
    PRT.EOF = 0
End Sub

Sub RREFAdd()
    If RptOPTION.BLDREF Then
       If (PAGE.XID > 0) And (PAGE.XREF < 100) Then
          PAGE.XREF = PAGE.XREF + 1
          RREFXID(PAGE.XREF) = PAGE.XID: RREFTYPE(PAGE.XREF) = PAGE.xType
       End If
       PAGE.XID = 0
    End If
End Sub

Sub RREFBld()
     If PAGE.XREF > 0 Then
        Z = Len(RREFXData$): xZ$ = Chr$(Z): Z = Z - (Z > 0)
        For X = 1 To PAGE.XREF: GoSub 1920: Next X
        PAGE.XREF = 0
     End If
     Exit Sub

1920 ' Create Reference File Entry
     XID = RREFXID(X): xType = RREFTYPE(X): XPTR& = RX&(XID)
     AGL = Asc(RREFREC.AGL): DGL = Asc(RREFREC.DGL)
     RREFRPT.NUMREFS = RREFRPT.NUMREFS + 1: OK = -1
     If XPTR& <> 0 Then
        Get #6, XPTR& + 1, RREFREC
        If RptOPTION.SNGLREF And (XPTR& > RREFHDR.EOFRBA) Then
           OK = 0
           X = RREFREC.REFTYPE: Y = xType
           If RptOPTION.SNGLREF = 1 Then
              P1 = 1: P2 = 0
              Else
              P1 = RXIMP(X): P2 = RXIMP(Y)
           End If
           If (P2 <= P1) Then
              RREFREC.REFTYPE = xType + (RREFREC.REFTYPE And 1024)
              RREFREC.Pgnum = PAGE.RPAGE: RREFREC.LINENUM = PAGE.RLINE
           End If
           Else
           RREFREC.NEXT = RREFRPT.EOFRBA
        End If
        Put #6, XPTR& + 1, RREFREC
        Else
        RREFRPT.NEWID = RREFRPT.NEWID + 1
     End If
     If OK Then
        If XPTR& < RREFHDR.EOFRBA Then RREFRPT.NUMID = RREFRPT.NUMID + 1
        RX&(XID) = RREFRPT.EOFRBA
        RREFREC.RID = XID: RREFREC.AGL = Chr$(AGL): RREFREC.DGL = Chr$(DGL)
        RREFREC.REFTYPE = xType - 1024 * (Z > 0)
        RREFREC.Pgnum = PAGE.RPAGE: RREFREC.LINENUM = PAGE.RLINE
        RREFREC.NEXT = 0
        Put #6, RREFRPT.EOFRBA + 1, RREFREC
        If Z > 0 Then Put #6, , xZ$: Put #6, , RREFXData$
        RREFRPT.EOFRBA = RREFRPT.EOFRBA + Len(RREFREC) + Z
     End If
     RREFREC.AGL = Chr$(0): RREFREC.DGL = Chr$(0)
     Return

End Sub

Sub RREFClose()
    If RREFRPT.BGNPAGE > 0 Then
       RREFRPT.LASTPAGE = PAGE.COUNT
       RREFRPT.TITLE = ENV.RTITLE
       Put #6, RREFHDR.EOFRBA + 1, RREFRPT
       RREFHDR.WFVER = Chr$(0)
       RREFHDR.RPTS = Chr$(Asc(RREFHDR.RPTS) + 1)
       RREFHDR.NUMID = RREFHDR.NUMID + RREFRPT.NEWID
       RREFHDR.NUMREFS = RREFHDR.NUMREFS + RREFRPT.NUMREFS
       If RREFHDR.LASTPAGE < RREFRPT.LASTPAGE Then RREFHDR.LASTPAGE = RREFRPT.LASTPAGE
       RREFHDR.EOFRBA = RREFRPT.EOFRBA
       Put #6, 1, RREFHDR
       ReDim RX&(1): RREFXData$ = ""
    End If
    Close #6
End Sub

Rem $DYNAMIC
Sub RREFOpen(OPT$)
    ' Open Index REFERENCE File for Creation
    QUIET = (Right$(OPT$, 1) = "Q"): X = Len(RTrim$(OPT$)) + QUIET
    XOPT$ = MID$(OPT$, 1, X): DSN$ = FT$(9)
REFDSN:
    If Not QUIET Then
       'Call FmtFIND("XXXXS017")
       'Call FmtFindFLD("DSNM"): LSet UTXT$(FFLD.UTXT) = DSN$
       'Call MsgFORMAT: If A = 27 Then GoTo REFOpenBAD
       'Call FmtFindFLD("DSNM"): LSet DSN$ = UTXT$(FFLD.UTXT)
    End If
    DSNAME$ = RTrim$(WRKFPFX$ + DSN$)
    LSet SysVAR.FILENAME = DSNAME$
    On Local Error GoTo NOREF
    Open DSNAME$ For Input As #6
REFCHK:
    Close #6
    On Local Error GoTo REFOpenERR
    If XOPT$ = "NEW" Then
       Call PutMSG("XXXXM002")
       If CA$ <> SNGLKEY$(1) Then GoTo REFOpenBAD
    End If
    If (XOPT$ = "NEW") Then
       Kill SysVAR.FILENAME
       GoTo REFNEW
    End If
    GoTo REFOLD
NOREF:
    X = ERR
    If XOPT$ = "WRITE" Then XOPT$ = "NEW"
    If (XOPT$ = "NEW") Then Resume REFNEW
    If QUIET Then Resume REFOpenBAD
    Call ERRMessage("XXXXM030", X, 0)
    If a = 27 Then Resume REFOpenBAD
    Resume REFDSN
REFOLD:
    On Local Error GoTo REFOpenERR
    'ON LOCAL ERROR GOTO 0
    Select Case XOPT$
       Case "INIT": Open SysVAR.FILENAME For Binary Access Read As #6
       Case "UPDT": Open SysVAR.FILENAME For Binary Access Read Write As #6
       Case "READ": Open SysVAR.FILENAME For Binary Access Read As #6
       Case "WRITE": Open SysVAR.FILENAME For Binary Access Read Write As #6
       Case Else: GoTo REFOpenBAD
    End Select
    SysVAR.MVAR = ".REF"
    Get #6, , RREFHDR
    If (RREFHDR.FType <> "W") Or (RREFHDR.WTYPE <> Chr$(4)) Or (RREFHDR.WFVER <> Chr$(0)) Then
       If XOPT$ = "WRITE" Then XOPT$ = "NEW": GoTo REFCHK
       Close #6: If QUIET Then GoTo REFOpenBAD
       Call PutMSG("XXXXM033")
       If XOPT$ = "INIT" Then GoTo REFDSN
       GoTo REFOpenBAD
    End If
    If (RREFHDR.REORGY <> FF1Hdr.REORGY) Or (RREFHDR.REORGMD <> FF1Hdr.REORGMD) Or (RREFHDR.REORGTM <> FF1Hdr.REORGTM) Then
       If XOPT$ = "WRITE" Then XOPT$ = "NEW": GoTo REFCHK
       Close #6: If QUIET Then GoTo REFOpenBAD
       Call PutMSG("XXXXM034")
       If XOPT$ = "INIT" Then GoTo REFDSN
       GoTo REFOpenBAD
    End If
    R = Asc(RREFHDR.RPTS)
    If (XOPT$ = "WRITE") Then
       If R > 0 Then
          SysVAR.NVAR1 = R + 1
          Call PutMSG("XXXXM062")
          If a = 27 Then GoTo REFOpenBAD
       End If
       If (R = 0) Or (CA$ = "1") Then GoTo REFNEW
    End If
    xx& = Len(RREFHDR): Z$ = " "
    ReDim RX&(FF1Hdr.MAXID)
    If XOPT$ = "INIT" Then GoTo REFOLDEnd
    While xx& < RREFHDR.EOFRBA
       Get #6, xx& + 1, RREFRPT
       If Asc(RREFRPT.RTYPE) And 128 Then
          Get #6, , Z$: xZ = Asc(Z$)
          Else
          xZ = 0
       End If
       xx& = xx& + Len(RREFRPT) + xZ
       While xx& < RREFRPT.EOFRBA
          Get #6, xx& + 1, RREFREC
          ' IF "READ" then RX&()=First REF, else RX&()=Last REF
          If (RREFREC.RID > 0) And (RREFREC.RID <= FF1Hdr.MAXID) Then
             If (XOPT$ = "READ") Then
                If RX&(RREFREC.RID) = 0 Then RX&(RREFREC.RID) = xx&
                Else
                RX&(RREFREC.RID) = xx&
             End If
          End If
          If RREFREC.REFTYPE And 1024 Then
             Get #6, , Z$: xZ = Asc(Z$) + 1
             Else
             xZ = 0
          End If
          xx& = xx& + Len(RREFREC) + xZ
       Wend
    Wend
    If XOPT$ = "WRITE" Then GoTo RPTNEW
REFOLDEnd:
    If R > 0 Then Get #6, Len(RREFHDR) + 1, RREFRPT
    OKAY = -1: GoTo REFOpenEnd
REFNEW:
    On Local Error GoTo REFOpenERR
    If XOPT$ = "WRITE" Then Close #6: Kill SysVAR.FILENAME
    Open SysVAR.FILENAME For Binary Access Read Write As #6
    Dim NEWHdr As RefFileHDR: RREFHDR = NEWHdr
    RREFHDR.FType = "W": RREFHDR.WTYPE = Chr$(4): RREFHDR.WFVER = Chr$(255)
    RREFHDR.REORGY = FF1Hdr.REORGY
    RREFHDR.REORGMD = FF1Hdr.REORGMD
    RREFHDR.REORGTM = FF1Hdr.REORGTM
    RREFHDR.UPDTY = DATEY(DATE$, 0)
    RREFHDR.UPDTMD = DATEMD(DATE$, 0)
    RREFHDR.EOFRBA = Len(RREFHDR)
    Put #6, 1, RREFHDR
    ReDim RX&(FF1Hdr.MAXID)
RPTNEW:
    Dim NEWRpt As RefFileRptREC: RREFRPT = NEWRpt
    xZ = Len(RREFXData$)
    RREFRPT.RTYPE = Chr$(RDFReport.TYPE - 128 * (xZ > 0))
    RREFRPT.RCODE = RDFReport.CODE
    RREFRPT.TITLE = RDFReport.NAME
    RREFRPT.UPDTY = DATEY(DATE$, 0)
    RREFRPT.UPDTMD = DATEMD(DATE$, 0)
    RREFRPT.UPDTTM = TIMEHHMM(Time$)
    RREFRPT.EOFRBA = RREFHDR.EOFRBA + Len(RREFRPT)
    Put #6, RREFHDR.EOFRBA + 1, RREFRPT
    If xZ > 0 Then
       Z$ = Chr$(xZ): Put #6, , Z$
       Put #6, , RREFXData$: RREFXData$ = ""
    End If
    ReDim RREFXID(100), RREFTYPE(100): PAGE.XREF = 0
    If PAGE.BGNPAGE = 0 Then PAGE.BGNPAGE = RREFHDR.LASTPAGE + 1
    OKAY = -1: GoTo REFOpenEnd
REFOpenERR:
    Call ERRMessage("XXXXM030", ERR, ERL)
    Resume REFOpenBAD
REFOpenBAD: OKAY = 0
REFOpenEnd:
    RREFREC.AGL = Chr$(0): RREFREC.DGL = Chr$(0)
    If OKAY Then LSet FT$(9) = DSN$
End Sub

Rem $STATIC
Sub RWRKGet(OPT$, RC, PTR&)
    ' OPT$="NEXT" gets next Work Rec in standard sequence
    '     ="WRKREC" same as "NEXT" without getting NAME rec
    '     ="RELREC" gets Work Rec pointed to by PTR&
    '     ="PTRREC" same as "RELREC" without getting NAME rec
    OKAY = -1: If (OPT$ <> "RELREC") And (OPT$ <> "PTRREC") Then PTR& = 0
    ' On Local Error GoTo RWRKGNErr
    Select Case OPT$
       Case "BASE": WFAGLREC.NEXT = Len(WFHDR) + Len(WFANCREC)
                    WFAGLREC.PRIOR = 0: WFAGLREC.TOTREC = 0
                    GoTo RWRKGNAnc
       Case "NEXT":
       Case "WRKREC":
       Case "RELREC": GoTo RWRKGNPtr
       Case "PTRREC": GoTo RWRKGNPtr
       Case Else: GoTo RWRKGNBad
    End Select
    If Asc(WFHDR.DMAXLV) > 0 Then GoTo RWRKGNDesc Else GoTo RWRKGNAnc

RWRKGNPtr:
    X& = PTR&
    If Asc(WFHDR.DMAXLV) = 0 Then
       Get #5, X& + 1, WFANCREC
       w& = WFANCREC.GACNT - 1
       Y& = X& - w& * Len(WFANCREC) - Len(WFAGLREC)
       Get #5, Y& + 1, WFAGLREC
       Get #5, WFAGLREC.NEXT + 1, WFAGLREC
       WFGLREC.AGL = -WFAGLREC.AGL: WFGLREC.DGL = 0
       WFRELREC.id = WFANCREC.ANCID: WFRELREC.NREF = WFANCREC.NXTDUP
       WFRELREC.CSTAT = (WFANCREC.PSTAT And 15)
       Else
       Get #5, X& + 1, WFRELREC: X& = WFRELREC.GLPTR
       If OPT$ <> "PTRREC" Then
          Get #5, X& + 1, WFGLREC: X& = WFGLREC.ANCPTR
          Get #5, X& + 1, WFANCREC
       End If
    End If
    RC = 0: GoTo RWRKGNEnd

RWRKGNAnc:
    Z = Asc(WFHDR.DMAXLV) + 1
    ReDim GBL(Z), GLNUMCH(Z), PARPTR&(Z)
    If WFAGLREC.PRIOR < WFAGLREC.TOTREC Then
       PTR& = WFAGLREC.ANCPTR + WFAGLREC.PRIOR * Len(WFANCREC)
       Get #5, PTR& + 1, WFANCREC
       WFAGLREC.PRIOR = WFAGLREC.PRIOR + 1
       RC = 2
       Else
       If WFAGLREC.NEXT = 0 Then
          RC = 0: GoTo RWRKGNEnd
          Else
          Get #5, WFAGLREC.NEXT + 1, WFAGLREC
          Get #5, WFAGLREC.ANCPTR + 1, WFANCREC
          PTR& = WFAGLREC.ANCPTR
          WFAGLREC.PRIOR = 1
          RC = 1
       End If
    End If
    If Asc(WFHDR.DMAXLV) = 0 Then
       WFGLREC.AGL = -WFAGLREC.AGL: WFGLREC.DGL = 0
       WFRELREC.id = WFANCREC.ANCID: WFRELREC.NREF = WFANCREC.NXTDUP
       WFRELREC.CSTAT = (WFANCREC.PSTAT And 15)
       Else
       Get #5, WFANCREC.FIRSTGL + 1, WFGLREC
       Get #5, , WFRELREC
       PTR& = WFANCREC.FIRSTGL + Len(WFGLREC)
       GLNUMCH(0) = Asc(WFRELREC.NUMCH)
    End If
    GoTo RWRKGNEnd

RWRKGNDesc:
    If (Asc(WFHDR.WFRULES) And 1) Then GoTo RWRKGNGRel
    If (WFRELREC.id < 0) Or (WFRELREC.OLDCH = 0) Then GoTo RWRKGNSib
    PTR& = WFRELREC.OLDCH
    Get #5, WFGLREC.NEXT + 1, WFGLREC
    Get #5, WFRELREC.OLDCH + 1, WFRELREC
    If GBL(WFGLREC.DGL) > 0 Then RC = 4 Else RC = 3
    GBL(WFGLREC.DGL) = Asc(WFRELREC.CHNUM) + WFRELREC.CSTAT * 256
    GLNUMCH(WFGLREC.DGL) = Asc(WFRELREC.NUMCH)
    Z = WFGLREC.DGL
    If (Z > 0) And (PARPTR&(Z) <> WFRELREC.PARPTR) Then
       While (Z > 0) And (PARPTR&(Z) <> WFRELREC.PARPTR)
          PARPTR&(Z) = WFRELREC.PARPTR
          Get #5, WFRELREC.PARPTR + 1, WFRELREC
          Z = Z - 1
          GBL(Z) = Asc(WFRELREC.CHNUM) + WFRELREC.CSTAT * 256
          GLNUMCH(Z) = Asc(WFRELREC.NUMCH)
       Wend
       Get #5, PTR& + 1, WFRELREC
       Else
       PARPTR&(Z) = WFRELREC.PARPTR
    End If
    GoTo RWRKGNEnd

RWRKGNSib:
    If WFGLREC.DGL < 1 Then GoTo RWRKGNAnc
    If (GBL(WFGLREC.DGL) And 255) < GLNUMCH(WFGLREC.DGL - 1) Then GoTo RWRKGNGRel
    GBL(WFGLREC.DGL) = 0: GLNUMCH(WFGLREC.DGL) = 0: PARPTR&(WFGLREC.DGL) = 0
    Get #5, WFRELREC.PARPTR + 1, WFRELREC
    Get #5, WFRELREC.GLPTR + 1, WFGLREC
    GoTo RWRKGNSib

RWRKGNGRel:
    If (WFRELREC.GLNEXT <> 0) Then
       PTR& = WFRELREC.GLNEXT
       Get #5, PTR& + 1, WFRELREC
       'IF WFRELREC.ID < 0 THEN GOTO RWRKGNGRel
       If GBL(WFGLREC.DGL) = 0 Then RC = 3 Else RC = 4
       GBL(WFGLREC.DGL) = Asc(WFRELREC.CHNUM) + 256 * WFRELREC.CSTAT
       GLNUMCH(WFGLREC.DGL) = Asc(WFRELREC.NUMCH)
       If (Asc(WFRELREC.CHNUM) > 1) Then GoTo RWRKGNEnd
       Z = WFGLREC.DGL
       If (Z > 0) And (PARPTR&(Z) <> WFRELREC.PARPTR) Then
          While (Z > 0) And (PARPTR&(Z) <> WFRELREC.PARPTR)
             PARPTR&(Z) = WFRELREC.PARPTR
             Get #5, WFRELREC.PARPTR + 1, WFRELREC
             Z = Z - 1
             GBL(Z) = Asc(WFRELREC.CHNUM) + WFRELREC.CSTAT * 256
             GLNUMCH(Z) = Asc(WFRELREC.NUMCH)
          Wend
          Get #5, PTR& + 1, WFRELREC
          Else
          PARPTR&(Z) = WFRELREC.PARPTR
       End If
       GoTo RWRKGNEnd
    End If
    If WFGLREC.NEXT = 0 Then GoTo RWRKGNAnc
    WFRELREC.GLNEXT = WFGLREC.NEXT + Len(WFGLREC)
    Get #5, WFGLREC.NEXT + 1, WFGLREC
    GoTo RWRKGNGRel

RWRKGNErr:
    Call ERRMessage("XXXXM031", ERR, ERL)
    Resume RWRKGNBad
RWRKGNBad:
    RC = 9: OKAY = 0
RWRKGNEnd:
    If (RC > 0) And (RC < 9) And (OPT$ <> "WRKREC") And (OPT$ <> "PTRREC") And ENV.FFOPEN Then
       If WFRELREC.id <> FF1Rec.RID Then Call FF1GetRec(Abs(WFRELREC.id))
    End If
End Sub

Sub RWRKOpen(OPT$, xType, DSNFMT)
    QUIET = (Right$(OPT$, 1) = "Q"): X = Len(RTrim$(OPT$)) + QUIET
    XOPT$ = MID$(OPT$, 1, X)
    xFT = xType
    If xFT = 0 Then xFT = -(Asc(WFHDR.AMAXLV) > 0) - 2 * (Asc(WFHDR.DMAXLV) > 0)
    If ((xFT < 1) Or (xFT > 3)) Then xFT = 3
    X = xFT + 4
    Select Case xFT           ' Set File Name Table Index
       Case 1: X = 5
       Case 2: X = 6
       Case 3: X = 7
       Case Else: GoTo WFOpenBAD
    End Select
    DSNAME$ = RTrim$(WRKFPFX$ + FT$(X))
    ' Allow User to Locate File
    If Not QUIET Then
    End If
    LSet SysVAR.FILENAME = DSNAME$
    On Local Error GoTo NOWF
CHKWF:
    Open SysVAR.FILENAME For Input As #5
    Close #5
    If XOPT$ = "NEW" Then
       Call PutMSG("XXXXM002")
       If CA$ <> SNGLKEY$(1) Then GoTo WFOpenBAD
       Kill SysVAR.FILENAME
       GoTo NEWWF
    End If
    GoTo OLDWF
NOWF: X = ERR
    If XOPT$ = "NEW" Then Resume NEWWF
    If Not QUIET Then Call ERRMessage("XXXXM030", X, 0)
    Resume WFOpenBAD
OLDWF:
    On Local Error GoTo WFOpenERR
    Select Case XOPT$
       Case "READ": Open SysVAR.FILENAME For Binary Access Read As #5
       Case "WRITE": Open SysVAR.FILENAME For Binary Access Read Write As #5
       Case Else: GoTo WFOpenBAD
    End Select
    SysVAR.MVAR = ".WRK"
    Get #5, , WFHDR
    If (WFHDR.FType <> "W") Or (WFHDR.WFTYPE <> Chr$(3)) Or (WFHDR.WFVER <> Chr$(2)) Then
       If Not QUIET Then Call PutMSG("XXXXM033")
       GoTo WFOpenBAD
    End If
    If (WFHDR.REORGY <> FF1Hdr.REORGY) Or (WFHDR.REORGMD <> FF1Hdr.REORGMD) Or (WFHDR.REORGTM <> FF1Hdr.REORGTM) Then
       If Not QUIET Then Call PutMSG("XXXXM034")
       GoTo WFOpenBAD
    End If
    Get #5, 1, WFHDR
    Get #5, , WFANCREC
    OKAY = -1
    GoTo WFOpenEND
NEWWF:
    On Local Error GoTo WFOpenERR
    Close #5
    Open SysVAR.FILENAME For Binary Access Read Write As #5
    WFHDR.FType = "W": WFHDR.WFTYPE = Chr$(3): WFHDR.WFVER = Chr$(255)
    WFHDR.REORGY = FF1Hdr.REORGY
    WFHDR.REORGMD = FF1Hdr.REORGMD
    WFHDR.REORGTM = FF1Hdr.REORGTM
    WFHDR.UPDTY = Val(Right$(DATE$, 4))
    WFHDR.UPDTMD = 100 * (Val(Left$(DATE$, 2))) + Val(MID$(DATE$, 4, 2))
    Put #5, 1, WFHDR
    OKAY = -1
    GoTo WFOpenEND
WFOpenERR:
    Call ERRMessage("XXXXM030", ERR, ERL)
    Resume WFOpenBAD
WFOpenBAD:
    OKAY = 0
WFOpenEND:
End Sub


Function SFILL$(X, L)
   S$ = LTrim$(Str$(FFRBA&(X, 1))): SL = Len(S$)
   If SL <= L Then
      SFILL$ = String$(L, Asc(" "))
      Mid$(SFILL$, L - SL + 1, SL) = S$
      Else
      SFILL$ = S$
   End If
End Function



Rem $DYNAMIC
Sub SWRKBldList(IDX(), SMAX, LFMT)
   ' Select File = 1 Byte for Each ID in .NAM, <>CHR$(0) if selected
   ' IDX() = Array of Size FF1HDR.MAXID provided by Calling Progam
   ' SMAX=High Non Zero Element of IDX() generated by SWRKBldList
   ' LFMT=0 --> IDX() = Full ID list, <>0 if selected (value of Byte)
   ' LFMT=1 --> IDX() = Min ID List, = ID's Selected
   Call SWRKOpen("READ")
   If OKAY Then
      Dim XRULE As SlctFileRULE
      SRCNT = Asc(SLCTHDR.SRULCNT)
      Seek #6, Len(SLCTHDR) + SRCNT * Len(XRULE) + 1
      Z$ = " ": X = 0: Y = 0: HI = 0: XMAX = SLCTHDR.HIID
      If XMAX > FF1Hdr.MAXID Then XMAX = FF1Hdr.MAXID
      While X < XMAX
         X = X + 1
         Get #6, , Z$: xZ = Asc(Z$)
         Select Case LFMT
            Case 0: Y = Y + 1: IDX(Y) = xZ: If xZ <> 0 Then HI = Y
            Case 1: If xZ <> 0 Then Y = Y + 1: IDX(Y) = X
         End Select
      Wend
      IDX(0) = SLCTHDR.LOID
      If LFMT = 1 Then SMAX = Y Else SMAX = HI
   End If
   Call FamCLOSE: ClearMSG
End Sub

Sub SWRKCreate(IDX(), SMAX, LFMT)
   ' Used to Create SELECT File with NO Rules
   ' LFMT 0:IDX() Full Check List  1:IDX() Short ID List
   Call SWRKOpen("NEW")
   If OKAY Then
      Call PutMSG("XXXXM060")
      SLCTHDR.SPGM = MID$(ENV.PGMNAME, 4, 4)
      SLCTHDR.LOID = 32767: SLCTHDR.HIID = 0
      Seek #6, Len(SLCTHDR) + 1
      X = 0: Y = 0: Z0$ = Chr$(0): Z$ = " "
      While X < SMAX
         X = X + 1: xZ = IDX(X)
         Select Case LFMT
            Case 0: xZ = xZ And 255: If xZ <> 0 Then Y = X
            Case 1: Y = Y + 1
                    While Y < xZ: Y = Y + 1: Put #6, , Z0$: Wend
                    xZ = 1
         End Select
         LSet Z$ = Chr$(xZ): Put #6, , Z$
         If xZ > 0 Then
            SLCTHDR.SlctCNT = SLCTHDR.SlctCNT - ((xZ And 1) = 1)
            SLCTHDR.SECSCNT = SLCTHDR.SECSCNT - ((xZ And 3) = 2)
            If Y < SLCTHDR.LOID Then SLCTHDR.LOID = Y
            SLCTHDR.HIID = Y
         End If
      Wend
      SLCTHDR.SLCTTOT = SLCTHDR.SlctCNT + SLCTHDR.SECSCNT
      SLCTHDR.WFVER = Chr$(2)
      Put #6, 1, SLCTHDR
   End If
   Call FamCLOSE: ClearMSG
End Sub

Sub SWRKOpen(XOPT$)
    QUIET = (Right$(XOPT$, 1) = "Q"): X = Len(RTrim$(XOPT$)) + QUIET
    OPT$ = MID$(XOPT$, 1, X)
    If Not QUIET Then
       Call frmFHSWFile.LOCATE("SELECT", "FULL")
       If a = 27 Then GoTo SFOpenBAD
       'Call FmtFIND("XXXXS013")
       'Call FmtFindFLD("DSNM"): LSet UTXT$(FFLD.UTXT) = FT$(10)
       'Call MsgFORMAT: If A = 27 Then GoTo SFOpenBAD
       'Call FmtFindFLD("DSNM"): LSet FT$(10) = UTXT$(FFLD.UTXT)
    End If
    DSNAME$ = RTrim$(WRKFPFX$ + FT$(10))
    LSet SysVAR.FILENAME = DSNAME$
    On Local Error GoTo NOSF
CHKSF:
    Open DSNAME$ For Input As #6
    Close #6
    If OPT$ = "NEW" Then
       If Not QUIET Then
          Call PutMSG("XXXXM002")
          If CA$ <> SNGLKEY$(1) Then GoTo SFOpenBAD
       End If
       Kill SysVAR.FILENAME
       GoTo NEWSF
    End If
    GoTo OLDSF
NOSF:
    X = ERR
    If OPT$ = "NEW" Then Resume NEWSF
    If Not QUIET Then Call ERRMessage("XXXXM030", X, 0)
    Resume SFOpenBAD
OLDSF:
    On Local Error GoTo SFOpenERR
    Select Case OPT$
       Case "READ": Open SysVAR.FILENAME For Binary Access Read As #6
       Case "WRITE": Open SysVAR.FILENAME For Binary Access Read Write As #6
       Case Else: GoTo SFOpenBAD
    End Select
    SysVAR.MVAR = "SELECT"
    Get #6, 1, SLCTHDR
    If (SLCTHDR.FType <> "W") Or (SLCTHDR.WFTYPE <> Chr$(5)) Or (SLCTHDR.WFVER <> Chr$(2)) Then
       If Not QUIET Then Call PutMSG("XXXXM033")
       GoTo SFOpenBAD
    End If
    If (SLCTHDR.REORGY <> FF1Hdr.REORGY) Or (SLCTHDR.REORGMD <> FF1Hdr.REORGMD) Or (SLCTHDR.REORGTM <> FF1Hdr.REORGTM) Then
       If Not QUIET Then Call PutMSG("XXXXM034")
       GoTo SFOpenBAD
    End If
    OKAY = -1
    GoTo SFOpenEND
NEWSF:
    On Local Error GoTo SFOpenERR
    Dim NEWSHDR As SlctFileHDR
    Close #6
    Open SysVAR.FILENAME For Binary Access Read Write As #6
    NEWSHDR.FType = "W": NEWSHDR.WFTYPE = Chr$(5): NEWSHDR.WFVER = Chr$(255)
    NEWSHDR.REORGY = FF1Hdr.REORGY
    NEWSHDR.REORGMD = FF1Hdr.REORGMD
    NEWSHDR.REORGTM = FF1Hdr.REORGTM
    NEWSHDR.UPDTY = Val(Right$(DATE$, 4))
    NEWSHDR.UPDTMD = 100 * (Val(Left$(DATE$, 2))) + Val(MID$(DATE$, 4, 2))
    NEWSHDR.MAXID = FF1Hdr.MAXID
    SLCTHDR = NEWSHDR
    Put #6, 1, SLCTHDR
    OKAY = -1
    GoTo SFOpenEND
SFOpenERR:
    Call ERRMessage("XXXXM030", ERR, ERL)
    Resume SFOpenBAD
SFOpenBAD:
    OKAY = 0
SFOpenEND:
End Sub

Rem $STATIC
Function TIMEHHMM(xx$)
   TIMEHHMM = Val(MID$(xx$, 1, 2)) * 100 + Val(MID$(xx$, 4, 2))
End Function

Function UCX$(xx$)
   X = Len(RTrim$(xx$)): Y = 0: Z = 0
   xY$ = Space$(100)
   While Y < X
      Y = Y + 1: C = Asc(MID$(xx$, Y, 1)): If C = 0 Then C = 32
      D = Asc(MID$(UCT1$, C, 1))
      If D > 0 Then
         Mid$(xY$, Z + 1, 255) = UCTRAN$(D): Z = Z + Len(UCTRAN$(D))
         Else
         Z = Z + 1: Mid$(xY$, Z, 1) = UCase(Chr$(C))
      End If
   Wend
   UCX$ = RTrim$(xY$)
End Function

Function XRPTDate$(MD, Y, O)
   xx$ = Space$(11)
   If (MD > 0) Or (Y > 0) Then
      M = MD \ 100: D = MD Mod 100
      If (MD = 0) And (Y = 9999) Then
         M = 13
         Else
         If (M > 13) Then M = 0
      End If
      If (0 < Y) And (Y < 9999) Then RSet xx$ = Str$(Y)
      If (MD > 0) Or (Y = 9999) Then
         Mid$(xx$, 4, 3) = MID$(RMONTB$(M + 1), 3, 3)
      End If
      If D > 0 Then Mid$(xx$, 1, 2) = SFILL$(D, 2)
   End If
   If O = 0 Then
      XRPTDate$ = LTrim$(RTrim$(xx$))
      Else
      XRPTDate$ = xx$
   End If
End Function

Function XSCRNDate$(MD, Y)
   M = MD \ 100: D = MD Mod 100
   xx$ = Space$(10):
   xY$ = Right$("000" + LTrim$(Str$(Y)), 4)
   XM$ = Right$("0" + LTrim$(Str$(M)), 2)
   xd$ = Right$("0" + LTrim$(Str$(D)), 2)
   'X = DFM: If X > 2 Then X = Val(CONFIG.DFM) - 1
   X = Val(CONFIG.DFM) - 1
   Select Case X
      Case 0: LSet xx$ = "  -  -": Mid$(xx$, 1, 2) = XM$: Mid$(xx$, 4, 2) = xd$
      Case 1: LSet xx$ = "  .  .": Mid$(xx$, 1, 2) = xd$: Mid$(xx$, 4, 2) = XM$
      Case 2: LSet xx$ = "    .  .": Mid$(xx$, 6, 2) = XM$: Mid$(xx$, 9, 2) = xd$
   End Select
   If X = 2 Then Mid$(xx$, 1, 4) = xY$ Else Mid$(xx$, 7, 4) = xY$
   XSCRNDate$ = xx$
End Function

Function XSCRNTime$(HHMM)
   HH$ = Right$("0" + LTrim$(Str$(HHMM \ 100)), 2)
   mm$ = Right$("0" + LTrim$(Str$(HHMM Mod 100)), 2)
   XSCRNTime$ = HH$ + ":" + mm$
End Function

Function ZFILL$(X, L)
   Z$ = LTrim$(Str$(FFRBA&(X, 1))): ZL = Len(Z$)
   If ZL <= L Then
      ZFILL$ = String$(L, Asc("0"))
      Mid$(ZFILL$, L - ZL + 1, ZL) = Z$
      Else
      ZFILL$ = Z$
   End If
End Function

