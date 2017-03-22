Attribute VB_Name = "modFHSWPrtc"
DefInt A-Z
   Dim PRTSetup As PrintSetupREC
   Dim CNT$, GS$, L$, PGMRTRN$
   Dim FILENAME$, HSETUP$
   Dim CODE$, XCODE$, xname$

Public Sub RptPRTC(OPT$)
1  Rem $INCLUDE: 'FHSCOMON.BAS'
   PN$ = " FAMILY HISTORY SYSTEM - FHSWPRNT - Printer Setup Table program"
   CY$ = " (C) Copyright 1995-1998 by Phillip E. Brown"
   Call RptLOAD("PRTC")
   Call TableLOAD("PRINTERS", PRTTNUM, 0)
   CNT$ = "   ": GS$ = " ": L$ = " ": PGMRTRN$ = RTrim$(ENV.PGMRTRN)
   FILENAME$ = Space$(12): HSETUP$ = Space$(46)
   CODE$ = Space$(26): XCODE$ = Space$(124): xname$ = Space$(8)

   Select Case OPT$
      Case "INIT": GoTo 40
      Case "PRINT": GoTo 1000
   End Select
   
40 ' INIT For Report
   Call PgmPREP("FHSPRTC")
   Load frmFHSWRPgm
   Call RptLOAD("PRTC")
   frmFHSWRPgm.RPSetup.AddItem "PRTC"
   frmFHSWRPgm.RPSetup_Click
   frmFHSWRPgm.Show
   Exit Sub

1000 'Print Reports
   Call PrintREPORT
End Sub
Private Sub PrintREPORT()
600 ' Process PRINT Request
    CTYPE = 2: xname$ = Space$(15)
    HSETUP = ENV.PSETUP
    Call PutMSG("PRTCM015")
    If A = 27 Then Exit Sub
    If (CA$ < "1") Or (CA$ > "2") Then Call ErrBEEP(0): GoTo 600
    PRINTALL = (CA$ = "2")
    'RDFReport.NAME = "Printer SETUP Listing"
    'HTFMAX = 8: REDIM HTFTXT$(HTFMAX)
    'HTFTXT$(1) = "H01<Family History System>"
    'HTFTXT$(2) = "H01<[RPTNAME]>"
    'HTFTXT$(3) = "H01"
    'HTFTXT$(4) = "H01Printer Definition File: [FILENAME]"
    'HTFTXT$(5) = "TF2   Forms Parameters and Options:"
    'HTFTXT$(6) = "TC2   Printer CODEs"
    'HTFTXT$(7) = "TS2   Chart Drawing SYMBOLs"
    'HTFTXT$(8) = "F02<Page [PAGE]>"
    Call RptOPEN("PRTC"): If (Not OKAY) Then GoTo 605
    Call PrintOPEN: If Not OKAY Then GoTo 605
601 PRT.SECTION = "0"
    ReDim RPTLIT$(24)
    Call GetRVAR("DEFAULT", x): RPTLIT$(0) = RVAR$(x) + " "
    Call GetRVAR("SETUP", x): RPTLIT$(1) = RVAR$(x) + " "
    Call GetRVAR("DESCRIPT", x): RPTLIT$(2) = "  " + RVAR$(x) + " "
    Call GetRVAR("PRINTER", x): RPTLIT$(3) = "   " + RVAR$(x) + " "
    Call GetRVAR("DEST", x): RPTLIT$(4) = "  " + RVAR$(x) + " "
    Call GetRVAR("WIDTH", x): RPTLIT$(5) = RVAR$(x)
    Call GetRVAR("LENGTH", x): RPTLIT$(6) = RVAR$(x)
    Call GetRVAR("TOPMGN", x): RPTLIT$(7) = RVAR$(x)
    Call GetRVAR("BOTMGN", x): RPTLIT$(8) = RVAR$(x)
    Call GetRVAR("INSMGN", x): RPTLIT$(9) = RVAR$(x)
    Call GetRVAR("OUTMGN", x): RPTLIT$(10) = RVAR$(x)
    Call GetRVAR("LINES", x): RPTLIT$(11) = " " + RVAR$(x) + "   "
    RPTLIT$(11) = " inches"
    Call GetRVAR("CHARS", x): RPTLIT$(12) = " " + RVAR$(x) + "   "
    RPTLIT$(12) = " inches"
    Call GetRVAR("OPTIONS", x): RPTLIT$(13) = RVAR$(x) + "   "
    Call GetRVAR("OPTSSIDE", x): RPTLIT$(14) = RVAR$(x) + " "
    Call GetRVAR("OPTFRMFD", x): RPTLIT$(15) = "  " + RVAR$(x) + " "
    Call GetRVAR("OPTPGWT", x): RPTLIT$(16) = "  " + RVAR$(x) + " "
    Call GetRVAR("INITIAL", x): RPTLIT$(17) = RVAR$(x)
    Call GetRVAR("FONT", x): RPTLIT$(18) = RVAR$(x)
    Call GetRVAR("NEWLINE", x): RPTLIT$(19) = RVAR$(x)
    Call GetRVAR("FORMFEED", x): RPTLIT$(20) = RVAR$(x)
    Call GetRVAR("RESET", x): RPTLIT$(21) = RVAR$(x)
    Call GetRVAR("LINESYMB", x): RPTLIT$(22) = RVAR$(x)
    Call GetRVAR("REGBOX", x): RPTLIT$(23) = RVAR$(x)
    Call GetRVAR("HIBOX", x): RPTLIT$(24) = RVAR$(x)
    SysVAR.FILENAME = CONFIG.PDF + ".PDF"
    LSet PLINE$ = RPTLIT$(0) + PSETUPS(PDFHdr.DEFAULT).NAME + RPTLIT$(2) + PSETUPS(PDFHdr.DEFAULT).DESC
    Call PrintLINE: Call WriteLINE
    Y = PSUTABLE.FT
    While Y > 0
       LSet xname$ = PSETUPS(Y).NAME
       If PRINTALL Or (PSETUPS(HSETUP).NAME = xname$) Then
          PDFSetup = PSETUPS(Y)
          ENV.PSETUP = Y
          GoSub 610
       End If
       Y = PSUCHN(Y).FWD
    Wend
    Call RptCLOSE: If PRT.AGAIN Then GoTo 601

605 ' End of Print
    Call PrintCLOSE
    LSet xname$ = PSETUPS(HSETUP).NAME
    Call PDFGetSet(xname$)
    Exit Sub

610 ' Print SETUP Information
    PRT.SECTION = "0"
    Call PrintLINE
    LSet PLINE$ = RPTLIT$(1) + PDFSetup.NAME + RPTLIT$(2) + RTrim$(PDFSetup.DESC)
    Call PrintLINE
    LSet PLINE$ = LTrim$(RPTLIT$(4)) + PDFSetup.DEST + "   " + "Printer: " + PRINTERS$(ENV.PSETUP)
    Call PrintLINE
    Call PDFGetSet(xname$)
    Style$ = "Normal "
    If (PDFSetup.FONTSTYL And 1) > 0 Then
       Style$ = "Bold "
    End If
    If (PDFSetup.FONTSTYL And 2) > 0 Then
       Style$ = Style$ + "Italic"
    End If
    LSet PLINE$ = "Font: " + PFONT$(ENV.PSETUP) + "     Font Size: " + Str$(PDFSetup.FONTSIZE) + "     Font Style: " + Style$
    Call PrintLINE
    PRT.SECTION = "F"
    Call PrintTITLE
    If (PDFSetup.OPTIONS And 32) > 0 Then
       O$ = "Landscape"
       Else
       O$ = "Portrait"
    End If
    Mid$(PLINE$, 8, 255) = "Orientation: " + O$
    Call PrintLINE
    Mid$(PLINE$, 8, 255) = RPTLIT$(5) + Str$(PDFSetup.FWIDTH / 100) + RPTLIT$(12)
    Mid$(PLINE$, 40, 255) = RPTLIT$(6) + Str$(PDFSetup.FLENGTH / 100) + RPTLIT$(11)
    Call PrintLINE
    Mid$(PLINE$, 8, 255) = RPTLIT$(7) + Str$(PDFSetup.TMARGIN / 100) + RPTLIT$(11)
    Mid$(PLINE$, 40, 255) = RPTLIT$(8) + Str$(PDFSetup.BMARGIN / 100) + RPTLIT$(11)
    Call PrintLINE
    Mid$(PLINE$, 8, 255) = RPTLIT$(9) + Str$(PDFSetup.IMARGIN / 100) + RPTLIT$(12)
    Mid$(PLINE$, 40, 255) = RPTLIT$(10) + Str$(PDFSetup.OMARGIN / 100) + RPTLIT$(12)
    Call PrintLINE
    If (PDFSetup.OPTIONS And 1) Then S$ = SNGLKEY$(1) Else S$ = SNGLKEY$(2)
    If (PDFSetup.OPTIONS And 2) Then F$ = SNGLKEY$(1) Else F$ = SNGLKEY$(2)
    If (PDFSetup.OPTIONS And 8) Then P$ = SNGLKEY$(1) Else P$ = SNGLKEY$(2)
    Mid$(PLINE$, 8, 255) = RPTLIT$(13) + RPTLIT$(14) + S$ + RPTLIT$(15) + F$ + RPTLIT$(16) + P$
    Call PrintLINE
    PRT.SECTION = "S"
    Call PrintTITLE
    Mid$(PLINE$, 8, 12) = RPTLIT$(22)
    LSet CODE$ = Chr$(11) + PDFSetup.SET1
    Call Code2Dec(CODE$, XCODE$)
    Mid$(PLINE$, 25, 255) = XCODE$
    Call PrintLINE
    Mid$(PLINE$, 8, 11) = RPTLIT$(23)
    LSet CODE$ = Chr$(10) + PDFSetup.SET2
    Call Code2Dec(CODE$, XCODE$)
    Mid$(PLINE$, 25, 255) = XCODE$
    Call PrintLINE
    Mid$(PLINE$, 8, 11) = RPTLIT$(24)
    LSet CODE$ = Chr$(10) + PDFSetup.SET3
    Call Code2Dec(CODE$, XCODE$)
    Mid$(PLINE$, 25, 255) = XCODE$
    Call PrintLINE
    Call WriteLINE: Call WriteLINE
    Return

611 ' Print CODE Line
    Select Case CTYPE
       Case 1: Call Code2Hex(CODE$, XCODE$)
       Case 2: Call Code2Dec(CODE$, XCODE$)
    End Select
    Mid$(PLINE$, 25, 255) = XCODE$
    Call PrintLINE
    Return

End Sub

Sub Code2Hex(ICODE$, OCODE$)
   N = Asc(ICODE$): x = 2: Y = 1: LSet OCODE$ = " "
   While N > 0
      H1 = Asc(MID$(ICODE$, x, 1)): H2 = (H1 And 15): H1 = (H1 \ 16)
      Mid$(OCODE$, Y, 1) = MID$("0123456789ABCDEF", H1 + 1, 1)
      Mid$(OCODE$, Y + 1, 1) = MID$("0123456789ABCDEF", H2 + 1, 1)
      x = x + 1: Y = Y + 3
      N = N - 1
   Wend
End Sub

Rem $DYNAMIC
Sub Code2Dec(ICODE$, OCODE$)
    N = Asc(ICODE$): x = 2: Y = 1: LSet OCODE$ = " "
    While N > 0
       If Y > 1 Then Mid$(OCODE$, Y, 1) = ",": Y = Y + 1
       Mid$(OCODE$, Y, 3) = LTrim$(Str$(Asc(MID$(ICODE$, x, 1))))
       Y = InStr(Y, OCODE$, " ")
       x = x + 1: N = N - 1
    Wend
End Sub


Sub Dec2Code(ICODE$, OCODE$)
    LSet OCODE$ = ""
    x = 1: Z = 1
    L = Len(ICODE$) + 1
    While x < L
       While MID$(ICODE$, x, 1) = " ": x = x + 1: Wend
       While MID$(ICODE$, x, 1) = ",": x = x + 1: Wend
       Y1 = InStr(x, ICODE$, ","): Y2 = InStr(x, ICODE$, " ")
       If Y1 = 0 Then Y1 = Y2 Else If Y2 = 0 Then Y2 = Y1
       If Y1 < Y2 Then Y = Y1 Else Y = Y2
       If Y = 0 Then Y = L
       If Y > x Then
          N# = Val(MID$(ICODE$, x, Y - x)): If N# > 255 Then N# = 0
          ZN = N#: Z = Z + 1: Mid$(OCODE$, Z, 1) = Chr$(ZN)
       End If
       x = Y + 1
    Wend
    Mid$(OCODE$, 1, 1) = Chr$(Z - 1)
End Sub

Sub Hex2Code(ICODE$, OCODE$)
    LSet OCODE$ = " ": L = Len(ICODE$) + 1
    x = 1: Z = 1
    While x < L
       While MID$(ICODE$, x, 1) = " ": x = x + 1: Wend
       While MID$(ICODE$, x, 1) = ",": x = x + 1: Wend
       Y1 = InStr(x, ICODE$, ","): Y2 = InStr(x, ICODE$, " ")
       If Y1 = 0 Then Y1 = Y2 Else If Y2 = 0 Then Y2 = Y1
       If Y1 < Y2 Then Y = Y1 Else Y = Y2
       If Y = 0 Then Y = L
       If (Y - x) = 2 Then
          xx$ = UCase$(MID$(ICODE$, x, 1))
          H1 = InStr("0123456789ABCDEF", xx$): If H1 > 0 Then H1 = H1 - 1
          xx$ = UCase$(MID$(ICODE$, x + 1, 1))
          H2 = InStr("0123456789ABCDEF", xx$): If H2 > 0 Then H2 = H2 - 1
          Z = Z + 1: Mid$(OCODE$, Z, 1) = Chr$(16 * H1 + H2)
          Else
          If (Y - x) = 1 Then
             ZN = Asc(MID$(ICODE$, x, 1))
             Else
             If Y > x Then
                N# = Val(MID$(ICODE$, x, Y - x))
                If N# > 255 Then ZN = 0 Else ZN = N#
             End If
          End If
          If ZN > 0 Then Z = Z + 1: Mid$(OCODE$, Z, 1) = Chr$(ZN): ZN = 0
       End If
       x = Y + 1
    Wend
    Mid$(OCODE$, 1, 1) = Chr$(Z - 1)
End Sub

