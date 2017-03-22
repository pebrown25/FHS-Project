Attribute VB_Name = "dllFHSWRun1"
DefInt A-Z
Sub RefreshForms()
mnuFHSWMain.Refresh
If FormLoaded(9) And MSGDForm Then frmFHSWMsgD.Refresh
If FormLoaded(9) And MSGForm Then frmFHSWMsg.Refresh
If FormLoaded(2) And ListForm Then frmFHSWList.Refresh
If FormLoaded(1) And FileForm Then frmFHSWFile.Refresh
If FormLoaded(1) And ViewForm Then frmFHSWView.Refresh
If FormLoaded(1) And RelWForm Then frmFHSWRelW.Refresh
If FormLoaded(1) And SlctForm Then frmFHSWSlct.Refresh
If FormLoaded(1) And NdWkForm Then frmFHSWNdWk.Refresh
If FormLoaded(1) And UPgmForm Then frmFHSWUPgm.Refresh
If FormLoaded(1) And RPgmForm Then frmFHSWRPgm.Refresh
If FormLoaded(1) And PrntForm Then frmFHSWPrinter.Refresh
If FormLoaded(2) And FamDefForm Then frmFHSWFFDF.Refresh
If FormLoaded(2) And RelCalcForm Then frmFHSWRCLC.Refresh
End Sub

Sub TableUPDT(OPT$, TP As TablePARMS, TCHN() As ChainPTRS)
    Select Case OPT$
       Case "ADD": GoSub TableADD
       Case "DEL": GoSub TableDEL
    End Select
    Call PutMSG("")
TUExit:
    Exit Sub

TableADD:
    NT = TP.FF
    If NT = 0 Then
       LSet SysVAR.MVAR = TP.TABNAME
       Call PutMSG("XXXXM021")
       OKAY = 0: A = 27: GoTo TADDEnd
       End If
    TP.FF = TCHN(NT).FWD: If TP.FF > 0 Then TCHN(TP.FF).BWD = 0
    If TP.FT = 0 Then
       TCHN(NT).BWD = 0: TCHN(NT).FWD = 0
       TP.FT = NT: TP.FS = NT: TP.CT = NT: TP.LT = NT
       TP.CL = TP.TL
    Else
       TCHN(NT).BWD = TP.CT: TCHN(NT).FWD = TCHN(TP.CT).FWD
       If TCHN(NT).FWD > 0 Then TCHN(TCHN(NT).FWD).BWD = NT
       TCHN(TP.CT).FWD = NT
       If TP.CL = TP.BL Then TP.FS = TCHN(TP.FS).FWD: Else TP.CL = TP.CL + 1
       If TP.LT = TP.CT Then TP.LT = NT
    End If
    TP.CT = NT
    TP.SIZE = TP.SIZE + 1
TADDEnd: Return

TableDEL:
    TF = TP.FF
    TP.FF = TP.CT
    BW = TCHN(TP.CT).BWD: FW = TCHN(TP.CT).FWD
    TCHN(TP.FF).FWD = TF: TCHN(TP.FF).BWD = 0
    If TF > 0 Then TCHN(TF).BWD = TP.FF
    If BW > 0 Then TCHN(BW).FWD = FW
    If FW > 0 Then TCHN(FW).BWD = BW
    If TP.CL = TP.TL Then
       TP.FS = FW: If TP.FS < 1 Then TP.FS = BW
    Else
       If FW = 0 Then TP.CL = TP.CL - 1
    End If
    If TP.FT = TP.CT Then TP.FT = FW
    If TP.LT = TP.CT Then TP.LT = BW
    TP.SIZE = TP.SIZE - 1
TDELEnd: Return

End Sub

Rem $DYNAMIC
Sub ChnINIT(CHN() As ChainPTRS, CHNBGN, CHNMAX)
    If CHNMAX > CHNBGN Then
       For X = CHNBGN To CHNMAX
           CHN(X).FWD = X + 1: CHN(X).BWD = X - 1
       Next X
       CHN(CHNBGN).BWD = 0: CHN(CHNMAX).FWD = 0
    End If
End Sub

Public Function ClearMSG()
If (FormLoaded(9) And MSGForm) Then
   Unload frmFHSWMsg
   RefreshForms
End If

End Function

Sub DataINIT(OPT$)
 ' Prepare to get DATA from FHSWINIT.DAT
   Close
   On Local Error GoTo INITErr
   Open "FHSWINIT.DAT" For Input As #5
   If OPT$ = "INIT" Then
      Input #5, IDataCNT: ReDim IDataNM$(IDataCNT), IDataBGN(IDataCNT)
      For X = 1 To IDataCNT
          Input #5, xx$, z$
          IDataNM$(X) = RTrim$(xx$): IDataBGN(X) = Val(z$)
      Next X
      Close #5
      OKAY = -1: GoTo INITEnd
   End If
   X = IDataCNT: OKAY = 0
   While (X > 0) And (Not OKAY)
      OKAY = (IDataNM$(X) = OPT$)
      If Not OKAY Then X = X - 1
   Wend
   If OKAY Then
      Seek #5, IDataBGN(X)
      Else
      SysVAR.MVAR = OPT$
      Call PutMSG("Unable to Locate DATA for [MVAR]...")
      Call GetRESP
   End If
   GoTo INITEnd

INITErr:
   MsgBox "Error Processing File FHSWINIT.DAT..."
   OKAY = 0: Resume INITEnd
INITEnd:

End Sub


Rem $STATIC
Sub ErrBEEP(XSND)
    'If DUR > 0 Then Beep BP, DUR
End Sub

Rem $DYNAMIC
Sub ERRMessage(XEM$, XET, XEL)
    X = EMCNT: FOUND = 0
    While (X > 0) And (Not FOUND)
       FOUND = (XET = Val(MID$(ERRMSG$(X), 1, 3)))
       If Not FOUND Then X = X - 1
    Wend
    EMSG$ = MID$(ERRMSG$(X), 4, 255)
    SysVAR.NVAR1 = XET: SysVAR.NVAR2 = XEL
    Call PutMSG(XEM$)
End Sub

Function FNBG(X)
  FNBG = (X Mod 128) \ 16
End Function

Function FNFG(X)
  FNFG = (X Mod 16) - 16 * (X > 127)
End Function

Rem $STATIC
Sub GetKEY()
GNKEY:
    CA$ = ""
    While CA$ = "": CA$ = INKEY$: Wend
    A = Asc(Right$(CA$, 1))
    'IF (LEN(A$) = 2 AND A = 50) THEN
    '   ENV.SHOWMSGID = NOT ENV.SHOWMSGID
    '   IF ENV.SHOWMSGID AND (SYSVAR.MSGNAME <> SPACE$(8)) THEN
    '      LOCATE 25, 1, 0: COLOR FG(5), BG(5): PRINT SYSVAR.MSGNAME;
    '   END IF
    '   GOTO GNKEY
    'END IF
End Sub

Sub GetRESP()
    'CA$ = UCASE$(INPUT$(1)): A = ASC(CA$)
End Sub


Rem $STATIC
Sub MDFOpen(OPT$)
    OKAY = 0
    LSet SysVAR.FILENAME = PGMPath + CONFIG.MDF + "MSGS.MDF"
    On Local Error GoTo NOMDF
CHKMDF:
    Open SysVAR.FILENAME For Input As #1
    Close #1
    If OPT$ = "NEW" Then
       Call PutMSG("XXXXM002")
       If CA$ <> SNGLKEY$(1) Then GoTo MDFOpenEND
       Kill SysVAR.FILENAME
       GoTo NEWMDF
    End If
    GoTo OLDMDF
NOMDF:
    X = ERR
    Close
    If OPT$ = "NEW" Then Resume NEWMDF
    Call ERRMessage("XXXXM030", X, 0)
    If A = 27 Then Resume MDFOpenEND
    Resume CHKMDF
OLDMDF:
    On Local Error GoTo MDFOpenERR
    Select Case OPT$
       Case "READ": Open SysVAR.FILENAME For Binary Access Read As #1
       Case "WRITE": Open SysVAR.FILENAME For Binary Access Read Write As #1
       Case Else: GoTo MDFOpenEND
    End Select
    Get #1, , MDFHdr
    LSet SysVAR.MVAR = "MDF"
    If MDFHdr.FType <> "MDF" Then
       Call PutMSG("XXXXM033")
       GoTo MDFOpenEND
    End If
    If MDFHdr.VER <> Chr$(0) Then
       Call PutMSG("XXXXM036")
       If CA$ <> SNGLKEY$(1) Then GoTo MDFOpenEND
    End If
    OKAY = -1
    GoTo MDFOpenEND
NEWMDF:
    On Local Error GoTo MDFOpenERR
    Open SysVAR.FILENAME For Binary Access Read Write As #1
    OKAY = -1
    GoTo MDFOpenEND
MDFOpenERR:
    Call ERRMessage("XXXXM030", ERR, ERL)
    Resume MDFOpenEND
MDFOpenEND:
    If Not OKAY Then Close Else ENV.MDF = CONFIG.MDF
    On Local Error GoTo 0
End Sub

Sub MSGExpand(ITXT$)
   OTXT$ = Space$(255): VDEPTH = 0
EXPTEXT:
   VCNT = 0: LSet OTXT$ = " "
   x1 = 1: Z1 = 1: ML = Len(RTrim$(ITXT$))
NXTMVAR:
   Y1 = InStr(x1, ITXT$, "["): If Y1 = 0 Then GoTo LASTMSEG
   Y2 = InStr(Y1, ITXT$, "]"): If Y2 = 0 Then GoTo LASTMSEG
   Mid$(OTXT$, Z1, Y1 - x1) = MID$(ITXT$, x1, Y1 - x1)
   Z1 = Z1 + Y1 - x1: x1 = Y2 + 1: VCNT = VCNT + 1
   MVAR$ = MID$(ITXT$, Y1 + 1, Y2 - Y1 - 1)
   Select Case MVAR$
      Case "MVAR": MVAL$ = RTrim$(SysVAR.MVAR)
      Case "NVAR1": MVAL$ = LTrim$(Str$(SysVAR.NVAR1))
      Case "NVAR2": MVAL$ = LTrim$(Str$(SysVAR.NVAR2))
      Case "NVAR3": MVAL$ = LTrim$(Str$(SysVAR.NVAR3))
      Case "PAGE": MVAL$ = LTrim$(Str$(PAGE.COUNT))
      Case "LINE": MVAL$ = LTrim$(Str$(PAGE.LINE))
      Case "TIME": MVAL$ = Time$
      Case "FAMFILE": MVAL$ = RTrim$(Left$(FSETUP$(ENV.FSETUP), 8))
      Case "FILENAME": MVAL$ = RTrim$(SysVAR.FILENAME)
      Case "PGMNAME": MVAL$ = RTrim$(SysVAR.PGMNAME)
      Case "RPTNAME": MVAL$ = RTrim$(RDFReport.NAME)
      Case "TABNAME": MVAL$ = RTrim$(SysVAR.TABNAME)
      Case "CHARTID": MVAL$ = RTrim$(SysVAR.CHARTID)
      Case "SUBJECT": MVAL$ = RTrim$(SysVAR.SUBJECT)
      Case "DATE": MVAL$ = SysVAR.DATE
      Case "PSFX": MVAL$ = ENV.PGMSFX
      Case "EMSG": MVAL$ = EMSG$
      Case Else: GoTo NOVAR
   End Select
   GoTo MOVESEG:
NOVAR:
   MVAL$ = MID$(ITXT$, Y1, Y2 - Y1 + 1): VCNT = VCNT - 1
MOVESEG:
   Mid$(OTXT$, Z1, Len(MVAL$)) = MVAL$
   Z1 = Z1 + Len(MVAL$)
   GoTo NXTMVAR
LASTMSEG:
   LS = ML - x1 + 1
   If LS > 0 Then Mid$(OTXT$, Z1, LS) = MID$(ITXT$, x1, LS)
   LSet ITXT$ = OTXT$: VDEPTH = VDEPTH + 1
   If (VCNT > 0) And (VDEPTH < 5) Then GoTo EXPTEXT
End Sub

Sub MSGFind(MSGID$, MNUM)
   ZPFX$ = Left$(MSGID$, 4)
   MNUM = 0: FOUND = 0: ZGRP = 0
   While (ZGRP < GRPMAX) And (Not FOUND)
      ZGRP = ZGRP + 1
      FOUND = (ZPFX$ = MDFIndex(ZGRP).NAME)
   Wend
   If Not FOUND Then Exit Sub
   ZSFX$ = "    ": LSet ZSFX$ = MID$(MSGID$, 5, 4)
   MNUM = MDFIndex(ZGRP).FIRST: FOUND = 0
   While (MNUM > 0) And (Not FOUND)
      FOUND = (ZSFX$ = Left$(MSGDATA$(MNUM), 4))
      If Not FOUND Then MNUM = MSGDATACHN(MNUM).FWD
   Wend
   If FOUND Then SysVAR.MSGNAME = MSGID$
End Sub


Sub MsgGrpFREE(GNAME$)
    GNUM = 1: OKAY = -1
    If GNAME$ = "ALL" Then GoTo MGRPFREE
    GN$ = Space$(4): LSet GN$ = GNAME$
    FOUND = 0
    While (GNUM < (GRPMAX + 1)) And (Not FOUND)
       FOUND = (MDFIndex(GNUM).NAME = GN$)
       If Not FOUND Then GNUM = GNUM + 1
    Wend
MGRPFREE:
    If GNUM > GRPMAX Then Exit Sub
    If (MDFIndex(GNUM).OPEN) Then
       If (MDFIndex(GNUM).FIRST > 0) Then
          X = MSGDATAFree
          Y = MDFIndex(GNUM).LAST: MDFIndex(GNUM).LAST = 0
          If X > 0 Then MSGDATACHN(X).BWD = Y
          If Y > 0 Then MSGDATACHN(Y).FWD = X
          While Y > 0
             MSGDATA$(Y) = ""
             Y = MSGDATACHN(Y).BWD
          Wend
          X = MDFIndex(GNUM).FIRST
          MDFIndex(GNUM).FIRST = 0
          MSGDATACHN(X).BWD = 0
          MSGDATAFree = X
       End If
       MDFIndex(GNUM).OPEN = 0
    End If
    If GNAME$ = "ALL" Then GNUM = GNUM + 1: GoTo MGRPFREE
End Sub

Sub MsgGrpLOAD(GNAME$, GNUM)
    On Local Error GoTo MGLERR
    GN$ = Space$(4): LSet GN$ = GNAME$
    SysVAR.FILENAME = CONFIG.MDF + "MSGS.MDF"
    If (GNAME$ = "ALL") Or (GNAME$ = "INITIAL") Then GoTo MGRPINIT
    If (CONFIG.MDF = ENV.MDF) Then GoTo MGRPLKUP
    MDFCHANGE = -1
MGRPINIT:
    LOADALL = ((GNAME$ = "ALL") Or ((GNAME$ = "INITIAL") And ENV.LARGEMEM))
    Call MDFOpen("READ"): If Not OKAY Then GoTo MGLEND
    FILEOPEN = -1
    For X = 1 To GRPMAX
       MDFIndex(X).CHG = 0: MDFIndex(X).OPEN = 0: MDFIndex(X).FDIR = 0
       MDFIndex(X).BGNRBA = 0: MDFIndex(X).SIZE = 0
       MDFIndex(X).FIRST = 0: MDFIndex(X).LAST = 0
    Next X
    If LOADALL Then MSGMAX = MDFHdr.MsgCNT Else MSGMAX = 0
    ReDim MSGDATA$(MSGMAX), MSGDATACHN(MSGMAX) As ChainPTRS
    Call ChnINIT(MSGDATACHN(), 1, MSGMAX)
    MSGDATAFree = -(MSGMAX > 0)
    Seek #1, Len(MDFHdr) + 1
    For X = 1 To MDFHdr.GrpCnt
        FDIR = Seek(1)
        Get #1, , MDFGrpDIR
        Y = 0: FOUND = 0
        While (Y < GRPMAX) And (Not FOUND)
           Y = Y + 1
           FOUND = (MDFGrpDIR.NAME = MDFIndex(Y).NAME)
        Wend
        If FOUND Then
           MDFIndex(Y).FDIR = FDIR
           MDFIndex(Y).CHG = MDFGrpDIR.CHG
           MDFIndex(Y).SIZE = MDFGrpDIR.MsgCNT
           MDFIndex(Y).BGNRBA = MDFGrpDIR.FIRST
           MDFIndex(Y).FIRST = 0
           MDFIndex(Y).LAST = 0
        End If
    Next X
    ENV.MDF = CONFIG.MDF
    If LOADALL Then XGN = 1: GoTo MGRPLOAD
CHGMGRPS:
    CN = 0
CHGMGNXT:
    CN = CN + 1
    Select Case CN
       Case 1: LSet GN$ = "XXXX": GoTo MGRPLKUP
       Case 2: LSet GN$ = MID$(ENV.PGMNAME, 4, 4): GoTo MGRPLKUP
    End Select
    If GNAME$ = "INITIAL" Then GoTo MGLEND
    LSet GN$ = GNAME$: MDFCHANGE = 0
MGRPLKUP:
    XGN = 0: OKAY = 0
    FOUND = 0
    While (XGN < GRPMAX) And (Not FOUND)
       XGN = XGN + 1: FOUND = (MDFIndex(XGN).NAME = GN$)
    Wend
    If Not FOUND Then
       SysVAR.MVAR = GN$
       Call PutMSG("XXXXM024")
       GoTo MGRPNEXT
    End If
    If GNAME$ <> "INITIAL" Then GNUM = XGN
MGRPLOAD:
    OKAY = -1
    If XGN > GRPMAX Then GoTo MGLEND
    If MDFIndex(XGN).OPEN Or (MDFIndex(XGN).BGNRBA = 0) Then GoTo MGRPNEXT
    If Not FILEOPEN Then
       Call MDFOpen("READ")
       If Not OKAY Then GoTo MGLEND
       FILEOPEN = -1
    End If
    Get #1, MDFIndex(XGN).FDIR, MDFGrpDIR
    Seek #1, MDFGrpDIR.FIRST
    X = MDFGrpDIR.MsgCNT: x1 = 0
    L$ = " "
    While (X > 0)
       If MSGDATAFree = 0 Then
          MSGDATAFree = MSGMAX + 1
          MSGMAX = MSGMAX + X
          ReDim Preserve MSGDATA$(MSGMAX), MSGDATACHN(MSGMAX) As ChainPTRS
          Call ChnINIT(MSGDATACHN(), MSGDATAFree, MSGMAX)
          If x1 > 0 Then MSGDATACHN(x1).FWD = MSGDATAFree
       End If
       If x1 = 0 Then
          MDFIndex(XGN).FIRST = MSGDATAFree
          MSGDATACHN(MSGDATAFree).BWD = -XGN
          Else
          MSGDATACHN(MSGDATAFree).BWD = x1
       End If
       x1 = MSGDATAFree: MSGDATAFree = MSGDATACHN(MSGDATAFree).FWD
       Get #1, , L$
       MSGDATA$(x1) = Space$(Asc(L$))
       Get #1, , MSGDATA$(x1)
       X = X - 1
    Wend
    MDFIndex(XGN).LAST = x1
    If MSGDATAFree > 0 Then MSGDATACHN(MSGDATAFree).BWD = 0
    If x1 > 0 Then MSGDATACHN(x1).FWD = 0
    MDFIndex(XGN).OPEN = -1
MGRPNEXT:
    If LOADALL Then XGN = XGN + 1: GoTo MGRPLOAD
    If GNAME$ = "INITIAL" Then GoTo CHGMGNXT
    If MDFCHANGE Then GoTo CHGMGNXT
    GoTo MGLEND
MGLERR:
    SysVAR.MVAR = GN$: SysVAR.NVAR1 = ERR: SysVAR.NVAR2 = ERL
    Call PutMSG("XXXXM078")
    OKAY = 0: Resume MGLFREE
MGLFREE:
    Call MsgGrpFREE(GN$)
MGLEND:
    Close
End Sub



Sub PgmPREP(PNAME$)
   OKAY = 0: LSet SysVAR.PGMNAME = PNAME$
   xOGRP$ = MID$(ENV.PGMNAME, 4, 4)
   PGRP$ = MID$(SysVAR.PGMNAME, 4, 4)
   XGRP = 0: FOUND = 0
   While (XGRP < GRPMAX) And (Not FOUND)
      XGRP = XGRP + 1
      FOUND = (PGRP$ = MDFIndex(XGRP).NAME)
   Wend
   If Not FOUND Then
      Call PutMSG("XXXXM025")
      Exit Sub
   End If
   Call MsgGrpLOAD(PGRP$, X)
   If OKAY And (xOGRP$ <> "    ") And (xOGRP$ <> PGRP$) Then
      Call MsgGrpFREE(xOGRP$)
   End If
   If OKAY Then
      ENV.PGMRTRN = ENV.PGMNAME
      ENV.PGMNAME = SysVAR.PGMNAME
      ENV.PGRP = XGRP
   End If
   On Local Error GoTo 0
   Close
End Sub

Rem $DYNAMIC
Sub PutMSG(MSGTXT$)
   msgresp = 0: msg$ = Space$(255): LSet SysVAR.MSGNAME = ""
   If Len(MSGTXT$) = 0 Then
      Unload frmFHSWMsg
      Exit Sub
   End If
   If (Len(MSGTXT$) <> 8) Then
      LSet msg$ = MSGTXT$: msgresp = -1
      RD$ = Chr$(13)
      GoTo EXPVARS
   End If
   Call MSGFind(MSGTXT$, zmsg)
   If FOUND Then GoTo BUILDMSG
   Call MSGFind("XXXXM000", zmsg)
   If FOUND Then
      SysVAR.MVAR = MSGTXT$
      GoTo BUILDMSG
   End If
   LSet msg$ = "Message " + RTrim$(MSGTXT$) + " Not Found"
   msgresp = -1: RD$ = Chr$(13): GoTo MSGPUT
BUILDMSG:
   msgresp = (MID$(MSGDATA$(zmsg), 5, 1) <> " ")
   If msgresp Then RD$ = MID$(MSGDATA$(zmsg), 6, 1)
   ERRTone = Val((MID$(MSGDATA$(zmsg), 7, 1)))
   LSet msg$ = MID$(MSGDATA$(zmsg), 8, 255)
EXPVARS:
   Call MSGExpand(msg$)
MSGPUT:
   If ENV.SHOWMSGID And SysVAR.MSGNAME <> Space$(8) Then
      msg$ = RTrim$(SysVAR.MSGNAME) + ": " + msg$
   End If
   If (FormLoaded(9) And MSGForm) Then
      frmFHSWMsg.Visible = False
      frmFHSWMsg.Refresh
      If msgresp Then ClearMSG
   End If
   If Not ((FormLoaded(9) And MSGForm) > 0) Then
      Load frmFHSWMsg
   End If
   frmFHSWMsg.FHSWMsg.TEXT = RTrim$(msg$)
   frmFHSWMsg.Show -msgresp
   If ERRTone <> 0 Then Call ErrBEEP(ERRTone)
   If msgresp Then
      If Len(CA$) = 1 Then CA$ = UCase$(CA$)
      If (CA$ = " ") Or (CA$ = Chr$(13)) Then
         CA$ = RD$: A = Asc(CA$)
      End If
      Else
      frmFHSWMsg.FHSWMsg.Refresh
      CA$ = Chr$(0): A = 0
   End If
End Sub




Rem $STATIC
Function SNX$(A)
   SNX$ = RTrim$(LTrim$(Str$(Abs(A))))
End Function

Rem $DYNAMIC
Sub TabINIT()
    Call DataINIT("TABLES"): If Not OKAY Then Exit Sub
    Input #5, TABMAX
    ReDim XTDF(TABMAX) As TableFileDIR, TABLE(TABMAX) As TablePARMS
    ReDim SNGLKEY$(0)
    For X = 1 To TABMAX
        Input #5, TDFDir.NAME, TDFDir.EntCNT, TDFDir.XL
        XTDF(X) = TDFDir
    Next X
    ReDim TDFIndex(TABMAX) As TableFileDIR
    For X = 1 To TABMAX: TDFIndex(X) = XTDF(X): Next X
    Call TableLOAD("INITIAL", T, 0)
    If Not OKAY Then Exit Sub
    For X = 1 To TABMAX
        If (TDFIndex(X).NAME = XTDF(X).NAME) Or (TDFIndex(X).NAME = "") Then
           TDFIndex(X).NAME = XTDF(X).NAME
           TDFIndex(X).XL = XTDF(X).XL
        End If
    Next X

    ' Load ERRMSGS Table
    Call TableLOAD("ERRMSGS", T, 0)
    EMCNT = TABLE(T).SIZE: ReDim ERRMSG$(EMCNT)
    X = TABLE(T).FT: z = 0
    While X > 0
       ERRMSG$(z) = RTrim$(TABDATA$(X))
       z = z + 1: X = TABDATACHN(X).FWD
    Wend
    Call TableFREE("ERRMSGS")

    'Load RELRULES Table
    Call TableLOAD("RELRULES", T, 0): ts = TABLE(T).SIZE
    ReDim RELRULE$(ts)
    X = TABLE(T).FT: z = 0
    While X > 0
       z = Val(MID$(TABDATA$(X), 1, 2))
       If z <= ts Then
          RELRULE$(z) = RTrim$(MID$(TABDATA$(X), 3, 12))
       End If
       X = TABDATACHN(X).FWD
    Wend
    Call TableFREE("RELRULES")

    ' Load RELCODES Table
    Call TableLOAD("RELCODES", T, 0): ts = TABLE(T).SIZE
    ReDim RELCODE$(ts)
    X = TABLE(T).FT: z = 0
    While X > 0
       z = Val(MID$(TABDATA$(X), 1, 2))
       If z <= ts Then
          RELCODE$(z) = RTrim$(MID$(TABDATA$(X), 1, 17))
       End If
       X = TABDATACHN(X).FWD
    Wend
    Call TableFREE("RELCODES")

    ' Load GENDER Table
    Call TableLOAD("GENDER", T, 0): ts = TABLE(T).SIZE
    ReDim FGENDR$(ts)
    X = TABLE(T).FT: z = 0
    While X > 0
      z = Val(MID$(TABDATA$(X), 1, 1))
      If z <= ts Then
         FGENDR$(z) = MID$(TABDATA$(X), 2, 1)
      End If
      X = TABDATACHN(X).FWD
    Wend
    Call TableFREE("GENDER")

    'Load REFERENC Table
    Call TableLOAD("REFERENC", T, 0): ts = TABLE(T).SIZE
    ReDim RXIMP(ts), RXSYM$(ts)
    X = TABLE(T).FT: z = 0
    While X > 0
      z = Val(MID$(TABDATA$(X), 1, 2))
      If z <= ts Then
         RXIMP(z) = Val(MID$(TABDATA$(X), 3, 2))
      End If
      X = TABDATACHN(X).FWD
    Wend
    Call TableFREE("REFERENC")

    'Load UCTRAN Table
    Call TableLOAD("UCTRAN", T, 0): ts = TABLE(T).SIZE
    UCT1$ = String$(255, 0)
    ReDim UCTRAN$(ts)
    X = TABLE(T).FT: Y = 0: z = 0
    While X > 0
      If Y < ts Then
         Y = Y + 1
         z = Asc(MID$(TABDATA$(X), 1, 1))
         Mid$(UCT1$, z, 1) = Chr$(Y)
         UCTRAN$(Y) = RTrim$(MID$(TABDATA$(X), 2, 3))
      End If
      X = TABDATACHN(X).FWD
    Wend
    Call TableFREE("UCTRAN")

    ' Load SOUNDEX Table
    Call TableLOAD("SOUNDEX", T, 0)
    SNDXTB$ = String$(255, 255)
    SXTB$ = "01230120022455012623010202"
    xx$ = " "
    For X = 1 To 26
        LSet xx$ = MID$(SXTB$, X, 1)
        Mid$(SNDXTB$, 64 + X, 1) = xx$
        Mid$(SNDXTB$, 96 + X, 1) = xx$
    Next X
    X = TABLE(T).FT
    While X > 0
       If Len(TABDATA$(X)) = 2 Then
          Z1 = Asc(MID$(TABDATA$(X), 1, 1))
          Mid$(SNDXTB$, Z1, 1) = MID$(TABDATA$(X), 2, 1)
       End If
       X = TABDATACHN(X).FWD
    Wend
    Call TableFREE("SOUNDEX")

250 ' Build SNGLKEY$() AND SKEG() - the Single Key Edit Group Table
    Call TableLOAD("SINGLKEY", T, 0)
    z = TABLE(T).SIZE
    ReDim SNGLKEY$(z)
    X = TABLE(T).FT: SNGLKEY$(0) = Chr$(z)
    While X > 0
       Y = Val(MID$(TABDATA$(X), 1, 2))
       SNGLKEY$(Y) = MID$(TABDATA$(X), 3, 1)
       X = TABDATACHN(X).FWD
    Wend
    Call TableFREE("SINGLKEY")
    Call DataINIT("SKEGTAB")
    If OKAY Then
       Input #5, z
       ReDim SKEG(z)
       For X = 0 To z: Input #5, SKEG(X): Next X
    End If

End Sub

Rem $STATIC
Sub TableFREE(TNAME$)
    TNUM = TABMAX: OKAY = -1
    If TNAME$ = "ALL" Then GoTo TABFREE
    TN$ = Space$(8): LSet TN$ = TNAME$
    FOUND = 0
    While (TNUM > 0) And (Not FOUND)
       FOUND = (TABLE(TNUM).TABNAME = TN$)
       If Not FOUND Then TNUM = TNUM - 1
    Wend
TABFREE:
    If TNUM < 1 Then Exit Sub
    If (TABLE(TNUM).OPEN) And (TDFIndex(TNUM).FIRST > 0) Then
       X = TABDATAFree
       Y = TABLE(TNUM).LT: TABLE(TNUM).LT = 0
       If X > 0 Then TABDATACHN(X).BWD = Y
       If Y > 0 Then TABDATACHN(Y).FWD = X
       While Y > 0
          TABDATA$(Y) = ""
          Y = TABDATACHN(Y).BWD
       Wend
       X = TABLE(TNUM).FT: TABLE(TNUM).FT = 0
       If X > 0 Then
          TABDATACHN(X).BWD = 0
          TABDATAFree = X
       End If
       TABLE(TNUM).OPEN = 0
    End If
    If TNAME$ = "ALL" Then TNUM = TNUM - 1: GoTo TABFREE
End Sub

Sub TableLOAD(TNAME$, TNUM As Integer, XL As Integer)
    On Local Error GoTo TLERR
    TL = XL
    TN$ = Space$(8): LSet TN$ = TNAME$
    If (TNAME$ = "ALL") Or (TNAME$ = "INITIAL") Then GoTo TabINIT
    If (CONFIG.TDF = ENV.TDF) Then GoTo TABLKUP
    TDFChange = -1
TabINIT:
    Call TDFOpen("READ"): If Not OKAY Then GoTo TLEND
    FILEOPEN = -1
    TDFMAX = TDFHdr.TabCNT
    If TNAME$ = "ALL" Then TABDATAMAX = TDFHdr.EntCNT Else TABDATAMAX = 0
    ReDim TABDATA$(TABDATAMAX), TABDATACHN(TABDATAMAX) As ChainPTRS
    Call ChnINIT(TABDATACHN(), 1, TABDATAMAX)
    TABDATAFree = -(TABDATAMAX > 0)
    Seek #1, Len(TDFHdr) + 1
    For XTN = 1 To TABMAX
       TDFIndex(XTN).CHG = 0
       TDFIndex(XTN).EntCNT = 0: TDFIndex(XTN).FIRST = 0
    Next XTN
    For XTN = 1 To TDFMAX
       Get #1, , TDFDir
       xn = 0: FOUND = 0
       While (xn < TABMAX) And (Not FOUND)
          xn = xn + 1
          FOUND = (TDFDir.NAME = TDFIndex(xn).NAME)
       Wend
       If FOUND Then
          XL = TDFIndex(xn).XL
          TDFIndex(xn) = TDFDir
          TDFIndex(xn).XL = XL
       End If
    Next XTN
    For XTN = 1 To TABMAX
        TABLE(XTN).TABNAME = TDFIndex(XTN).NAME
        TABLE(XTN).OPEN = 0
        TABLE(XTN).SIZE = TDFIndex(XTN).EntCNT
        TABLE(XTN).XL = TDFIndex(XTN).XL
        TABLE(XTN).KO = TDFIndex(XTN).KO
        TABLE(XTN).KL = TDFIndex(XTN).KL
        TABLE(XTN).FT = 0: TABLE(XTN).LT = 0
    Next XTN
    OKAY = -1
    ENV.TDF = CONFIG.TDF
    If TNAME$ = "ALL" Then GoTo TABLKUP
CHGGRPS:
    CN = 0
CHGNEXT:
    CN = CN + 1
    Select Case CN
       Case 1: LSet TN$ = "SINGLKEY": TL = 3: GoTo TABLKUP
       Case 2: LSet TN$ = "RPTOPTS": TL = 45: GoTo TABLKUP
    End Select
    If TNAME$ = "INITIAL" Then GoTo TLEND
    TL = XL: TDFChange = 0
TABLKUP:
    XTN = 1: OKAY = 0
    If TNAME$ = "ALL" Then GoTo TABLOAD
    FOUND = 0: TNUM = 0
    While (XTN < TABMAX + 1) And (Not FOUND)
       FOUND = (TDFIndex(XTN).NAME = TN$)
       If Not FOUND Then XTN = XTN + 1
    Wend
    If Not FOUND Then
       LSet SysVAR.MVAR = TN$
       Call PutMSG("XXXXM013")
       Exit Sub
    End If
    If TNAME$ <> "INITIAL" Then TNUM = XTN
TABLOAD:
    If XTN > TABMAX Then GoTo TLEND
    If TL > 0 Then ML = TL Else ML = TABLE(XTN).XL
    If (TABLE(XTN).OPEN) Then
       If ((Not (ML > TABLE(XTN).ML)) Or (TDFIndex(XTN).FIRST = 0)) Then
          OKAY = -1: GoTo TABNEXT
       End If
    End If:
    If Not FILEOPEN Then
       Call TDFOpen("READ")
       If Not OKAY Then GoTo TLEND
       FILEOPEN = -1
    End If
    If TDFIndex(XTN).FIRST = 0 Then
       If (TNAME$ <> "ALL") And (TNAME$ <> "UCTRAN") And (TNAME$ <> "SOUNDEX") Then
          SysVAR.MVAR = TABLE(XTN).TABNAME
          Call PutMSG("XXXXM014")
       End If
       GoTo TABNEXT
    End If
    L$ = " "
    Seek #1, TDFIndex(XTN).FIRST
    X = TABLE(XTN).SIZE: x1 = 0
    While (X > 0)
       If TABDATAFree = 0 Then
          TABDATAFree = TABDATAMAX + 1
          TABDATAMAX = TABDATAMAX + X
          ReDim Preserve TABDATA$(TABDATAMAX), TABDATACHN(TABDATAMAX) As ChainPTRS
          Call ChnINIT(TABDATACHN(), TABDATAFree, TABDATAMAX)
          If x1 > 0 Then TABDATACHN(x1).FWD = TABDATAFree
       End If
       If x1 = 0 Then
          TABLE(XTN).FT = TABDATAFree
          TABDATACHN(TABDATAFree).BWD = -XTN
          Else
          TABDATACHN(TABDATAFree).BWD = x1
       End If
       x1 = TABDATAFree: TABDATAFree = TABDATACHN(TABDATAFree).FWD
       Get #1, , L$
       ENTRY$ = Space$(Asc(L$))
       Get #1, , ENTRY$
       TABDATA$(x1) = Left$(ENTRY$, ML)
       X = X - 1
    Wend
    TABLE(XTN).LT = x1
    If TABDATAFree > 0 Then TABDATACHN(TABDATAFree).BWD = 0
    If x1 > 0 Then TABDATACHN(x1).FWD = 0
    TABLE(XTN).ML = ML
    TABLE(XTN).OPEN = -1
    OKAY = -1
    If TN$ = "RPTOPTS " Then ENV.OPTTAB = XTN
TABNEXT:
    If TNAME$ = "ALL" Then XTN = XTN + 1: GoTo TABLOAD
    If TNAME$ = "INITIAL" Then GoTo CHGNEXT
    If TDFChange Then GoTo CHGGRPS
    GoTo TLEND
TLERR:
    SysVAR.MVAR = TN$
    Call PutMSG("XXXXM079")
    OKAY = 0: Resume TLEND
TLEND:
    Close
    If (Not OKAY) And (CONFIG.TDF <> ENV.TDF) And (A <> 27) Then
       CONFIG.TDF = ENV.TDF: GoTo TABLKUP
    End If
    On Local Error GoTo 0
End Sub



Sub TDFOpen(OPT$)
    OKAY = 0
    SysVAR.FILENAME = PGMPath + CONFIG.TDF + "TABLE.TDF"
    On Local Error GoTo NOTDF
CHKTDF:
    Open SysVAR.FILENAME For Input As #1
    Close #1
    If OPT$ = "NEW" Then
       Call PutMSG("XXXXM002")
       If CA$ <> SNGLKEY$(1) Then GoTo TDFOpenEND
       Kill SysVAR.FILENAME
       GoTo NEWTDF
    End If
    GoTo OLDTDF
NOTDF:
    X = ERR: Close
    If OPT$ = "NEW" Then Resume NEWTDF
    Call ERRMessage("XXXXM030", X, 0)
    If A = 27 Then Resume TDFOpenEND
    Resume CHKTDF
OLDTDF:
    On Local Error GoTo TDFOpenERR
    Select Case OPT$
       Case "READ": Open SysVAR.FILENAME For Binary Access Read As #1
       Case "WRITE": Open SysVAR.FILENAME For Binary Access Write As #1
       Case Else: GoTo TDFOpenEND
    End Select
    Get #1, , TDFHdr
    If TDFHdr.FType <> "TDF" Then
       Call PutMSG("TABLM010")
       GoTo TDFOpenEND
    End If
    If TDFHdr.VER <> Chr$(0) Then
       Call PutMSG("TABLM011")
       If CA$ <> SNGLKEY$(1) Then GoTo TDFOpenEND
    End If
    OKAY = -1
    GoTo TDFOpenEND
NEWTDF:
    On Local Error GoTo TDFOpenERR
    Open SysVAR.FILENAME For Binary Access Read Write As #1
    OKAY = -1
    GoTo TDFOpenEND
TDFOpenERR:
    Call ERRMessage("XXXXM030", ERR, ERL)
    Resume TDFOpenEND
TDFOpenEND:
    If Not OKAY Then Close Else ENV.TDF = CONFIG.TDF
    On Local Error GoTo 0
End Sub














