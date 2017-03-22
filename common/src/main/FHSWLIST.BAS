Attribute VB_Name = "modFHSWList"
DefInt A-Z
Type SUMREC
   TOTAL    As Integer
   MALE     As Integer
   FEMALE   As Integer
   GROUPS   As Integer
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

Dim SlctCNT, SecSlctCNT
Dim SBASE, SECSLCT, SAVEWORK
Dim SRuleCNT, UFLD(), RULES() As SlctFileRULE
Dim DFNUM(), DFPP(), DFSEP(), INDEXED
Dim RID, RELID, xSEL
Dim SPRESEQ, SPRNO, SPCNT, MAXMarry, MLIST()
Dim DATECNT, DATERNG, PLACCNT, MARRCNT, RWIDTH, VRTCL
Dim INDID, ADRSPTR, ADRSSPID
Dim XBASE
Dim SNDX$, KeyWORD$, SrchWORD$
Dim SUMBASE As SUMREC, sum As SUMREC, TOT As SUMREC
Dim SST1$

Rem $DYNAMIC
Sub FSUMReport()
   Dim STOT(12, 12), RTYPE$(12)
   Call FAMOpen("READ"): If Not OKAY Then Exit Sub
   'Call PutMSG("UPDTM012")
   frmFHSWRPgm.MousePointer = 11
   X = 0
   While X < FF1Hdr.MAXID
      X = X + 1
      Call FF1GetRec(X)
      STOT(1, 0) = STOT(1, 0) + 1
      STOT(0, 1) = STOT(0, 1) + 1
      If FF1Rec.FID <> 0 Then STOT(0, 2) = STOT(0, 2) + 1
      If FF1Rec.MID <> 0 Then STOT(0, 3) = STOT(0, 3) + 1
      If (FF1Rec.MID <> 0) And (FF1Rec.FID <> 0) Then STOT(0, 4) = STOT(0, 4) + 1
      If FF1Rec.OLDCH <> 0 Then STOT(0, 5) = STOT(0, 5) + 1
      If FF1Rec.ADRS <> 0 Then STOT(1, 2) = STOT(1, 2) + 1
      If FF1Rec.SPOUSE <> 0 Then STOT(1, 3) = STOT(1, 3) + 1
      If FF1Rec.BLOC <> 0 Then STOT(1, 4) = STOT(1, 4) + 1
      If FF1Rec.COM <> 0 Then STOT(1, 5) = STOT(1, 5) + 1
      If FF1Rec.EVENT <> 0 Then STOT(1, 6) = STOT(1, 6) + 1
      If FF1Rec.EDUC <> 0 Then STOT(1, 7) = STOT(1, 7) + 1
      If FF1Rec.WORK <> 0 Then STOT(1, 8) = STOT(1, 8) + 1
      If FF1Rec.MIL <> 0 Then STOT(1, 9) = STOT(1, 9) + 1
      If FF1Rec.HLTH <> 0 Then STOT(1, 10) = STOT(1, 10) + 1
      CA$ = INKEY$: If CA$ = Chr$(27) Then X = FF1Hdr.MAXID
   Wend
   'Call PutMSG("UPDTM013")
   xX& = 0: H2& = FFRBA&(FF2Hdr.HIREC, 1): D = 1
   While xX& < H2&
      xX& = xX& + 1: X = FFRNUM%(xX&, 1)
      Call FF2GetRec(X): If Not OKAY Then xX& = H2&
      If OKAY And (FF2PFX.RTYPE <> Chr$(255)) Then GoSub ADRTOT
   Wend
   X = FF2Hdr.FREEBGN: D = -1
   While X <> 0
      STOT(0, 9) = STOT(0, 9) + 1
      Call FF2GetRec(X)
      If OKAY Then X = FF2PFX.NXT: Else X = 0
   Wend
   'Call PutMSG("UPDTM014")
   xX& = 0: H3& = FFRBA&(FF3HDR.HIREC, 1): D = 1
   While xX& < H3&
      xX& = xX& + 1: X = FFRNUM%(xX&, 1)
      Call FF3GetRec(X)
      If Not OKAY Then xX& = H3&
      If OKAY And (FF3PFX.RTYPE <> Chr$(255)) Then GoSub OTHTOT
      CA$ = INKEY$: If CA$ = Chr$(27) Then xX& = H3&
   Wend
   X = FF3HDR.FREEBGN: D = -1
   While X <> 0
      STOT(0, 10) = STOT(0, 10) + 1
      Call FF3GetRec(X): If Not OKAY Then X = 0
      If OKAY And (FF3PFX.RTYPE = Chr$(255)) Then X = FF3PNTRS.NXT Else X = 0
      CA$ = INKEY$: If CA$ = Chr$(27) Then X = 0
   Wend
   Call FamCLOSE
   GoTo FSUMPrint

ADRTOT:
   If D > 0 Then STOT(2, 0) = STOT(2, 0) + D
   Select Case Asc(FF2PFX.SRTYPE)
      Case 1: STOT(2, 1) = STOT(2, 1) + D
              If FF2PFX.COM <> 0 Then STOT(4, 5) = STOT(4, 5) + D
      Case 3: STOT(2, 3) = STOT(2, 3) + D
              If FF2PFX.COM <> 0 Then STOT(4, 6) = STOT(4, 6) + D
      Case 7: STOT(2, 7) = STOT(2, 7) + D
              If FF2PFX.COM <> 0 Then STOT(4, 7) = STOT(4, 7) + D
      Case 8: STOT(2, 8) = STOT(2, 8) + D
              If FF2PFX.COM <> 0 Then STOT(4, 8) = STOT(4, 8) + D
      Case 9: STOT(2, 9) = STOT(2, 9) + D
              If FF2PFX.COM <> 0 Then STOT(4, 9) = STOT(4, 9) + D
      Case 10: STOT(2, 10) = STOT(2, 10) + D
              If FF2PFX.COM <> 0 Then STOT(4, 10) = STOT(4, 10) + D
   End Select
   Return

OTHTOT:
   If D > 0 Then STOT(3, 0) = FFRNUM(FFRBA&(STOT(3, 0), 1) + D, 1)
   Select Case Asc(FF3PFX.RTYPE)
      Case 3: STOT(3, 1) = STOT(3, 1) + D
              If FF3SPOUSE.LOC <> 0 Then STOT(3, 2) = STOT(3, 2) + D
              If FF3SPOUSE.MLOC <> 0 Then STOT(3, 4) = STOT(3, 4) + D
              If FF3SPOUSE.COM <> 0 Then STOT(3, 5) = STOT(3, 5) + D
      Case 4: STOT(4, 0) = STOT(4, 0) + D:
              Select Case Asc(FF3PLACE.SRTYPE)
                 Case 1: STOT(4, 1) = STOT(4, 1) + D
                 Case 3: STOT(4, 3) = STOT(4, 3) + D
              End Select
      Case 5: STOT(5, 0) = FFRNUM(FFRBA&(STOT(5, 0), 1) + D, 1)
              Select Case Asc(FF3COMMENT.SRTYPE)
                 Case 1: STOT(5, 1) = FFRNUM(FFRBA&(STOT(5, 1), 1) + D, 1)
                 Case 2: Z& = FFRBA&(FF3COMMENT.SRNO, 1)
                         If Z& <= H2& Then
                            Call FF2GetRec(FF3COMMENT.SRNO)
                            If FF2PFX.RTYPE = Chr$(2) Then
                               Select Case Asc(FF2PFX.SRTYPE)
                                  Case 1: STOT(5, 4) = FFRNUM(FFRBA&(STOT(5, 4), 1) + D, 1)
                                  Case 3: STOT(6, 4) = STOT(6, 4) + D
                                  Case 7: STOT(7, 4) = STOT(7, 4) + D
                                  Case 8: STOT(8, 4) = STOT(8, 4) + D
                                  Case 9: STOT(9, 4) = STOT(9, 4) + D
                                  Case 10: STOT(10, 4) = STOT(10, 4) + D
                               End Select
                            End If
                         End If
                 Case 3: STOT(5, 3) = STOT(5, 3) + D
                 Case 6: STOT(5, 6) = STOT(5, 6) + D
                 Case 7: STOT(5, 7) = STOT(5, 7) + D
                 Case 8: STOT(5, 8) = STOT(5, 8) + D
                 Case 9: STOT(5, 9) = STOT(5, 9) + D
                 Case 10: STOT(5, 10) = STOT(5, 10) + D
              End Select
      Case 6: STOT(6, 1) = STOT(6, 1) + D
              If FF3EVENT.COM <> 0 Then STOT(6, 5) = STOT(6, 5) + D
      Case 7: STOT(7, 1) = STOT(7, 1) + D
              If FF3PNTRS.LOC <> 0 Then STOT(7, 2) = STOT(7, 2) + D
              If FF3PNTRS.COM <> 0 Then STOT(7, 5) = STOT(7, 5) + D
      Case 8: STOT(8, 1) = STOT(8, 1) + D
              If FF3PNTRS.LOC <> 0 Then STOT(8, 2) = STOT(8, 2) + D
              If FF3PNTRS.COM <> 0 Then STOT(8, 5) = STOT(8, 5) + D
      Case 9: STOT(9, 1) = STOT(9, 1) + D
              If FF3PNTRS.LOC <> 0 Then STOT(9, 2) = STOT(9, 2) + D
              If FF3PNTRS.COM <> 0 Then STOT(9, 5) = STOT(9, 5) + D
      Case 10: STOT(10, 1) = STOT(10, 1) + D
              If FF3PNTRS.LOC <> 0 Then STOT(10, 2) = STOT(10, 2) + D
              If FF3PNTRS.COM <> 0 Then STOT(10, 5) = STOT(10, 5) + D
      Case 11: STOT(11, 1) = STOT(11, 1) + D
              If FF3EVENT.COM <> 0 Then STOT(11, 5) = STOT(11, 5) + D
      Case 12: STOT(12, 0) = STOT(12, 0) + D
              Select Case Asc(FF3PLACE2.SRTYPE)
                 Case 1: STOT(12, 1) = STOT(12, 1) + D
                 Case 3: STOT(12, 3) = STOT(12, 3) + D
                 Case 6: STOT(12, 6) = STOT(12, 6) + D
                 Case 11: STOT(12, 11) = STOT(12, 11) + D
              End Select
   End Select
   Return

FSUMPrint:
   Call RptOPEN("FSUM")
   If OKAY Then Call PrintOPEN: If Not OKAY Then Exit Sub
   PRT.SECTION = "0"
   Call WriteLINE
   Call GetRVAR("FFSETUP", X)
   Mid$(PLINE$, 1, 25) = RVAR$(X): X = InStr(X, PLINE$, "  ")
   Mid$(PLINE$, X + 1, 50) = RTrim$(FDFSetup.NAME) + ", " + RTrim$(FDFSetup.DESC)
   Call WriteLINE
   PRT.SECTION = "F": FS$ = Space$(12): S5$ = Space$(5)
   Call PrintTITLE
   Mid$(PLINE$, 3, 8) = FType$(1): Mid$(PLINE$, 14, 3) = SFILL$(FF1Hdr.REVLVL, 3)
   Mid$(PLINE$, 21, 10) = XSCRNDate$(FF1Hdr.UPDTMD, FF1Hdr.UPDTY)
   Mid$(PLINE$, 36, 5) = SFILL$(FF1Hdr.MAXID, 5)
   RSet FS$ = Str$(FFRBA&(FF1Hdr.MAXID + 1, 100))
   Mid$(PLINE$, 55, 12) = FS$
   Call WriteLINE
   If (FF2Hdr.REVLVL And 255) = 0 Then RL = 108 Else RL = 152
   Mid$(PLINE$, 3, 8) = FType$(2): Mid$(PLINE$, 14, 3) = SFILL$(FF2Hdr.REVLVL, 3)
   Mid$(PLINE$, 21, 10) = XSCRNDate$(FF2Hdr.UPDTMD, FF2Hdr.UPDTY)
   Mid$(PLINE$, 36, 5) = SFILL$(FF2Hdr.HIREC, 5)
   Mid$(PLINE$, 45, 5) = SFILL$(FF2Hdr.FREECNT, 5)
   RSet FS$ = Str$(FFRBA&(FF2Hdr.HIREC + 1, RL)): Mid$(PLINE$, 55, 12) = FS$
   Call WriteLINE
   Mid$(PLINE$, 3, 8) = FType$(3): Mid$(PLINE$, 14, 3) = SFILL$(FF3HDR.REVLVL, 3)
   Mid$(PLINE$, 21, 10) = XSCRNDate$(FF3HDR.UPDTMD, FF3HDR.UPDTY)
   Mid$(PLINE$, 36, 5) = SFILL$(FF3HDR.HIREC, 5)
   Mid$(PLINE$, 45, 5) = SFILL$(FF3HDR.FREECNT, 5)
   RSet FS$ = Str$(FFRBA&(FFRNUM(FFRBA&(FF3HDR.HIREC, 1) + 1, 1), 50))
   Mid$(PLINE$, 55, 12) = FS$
   Call WriteLINE
   Call WriteLINE
   Call GetRVAR("CHANGES", X)
   Mid$(PLINE$, 3, 255) = RVAR$(X): Y = InStr(3, PLINE$, "  ") + 3
   Call GetRVAR("CHGDATE", X)
   Mid$(PLINE$, Y, 255) = RVAR$(X): xZ = InStr(Y, PLINE$, "  ")
   Mid$(PLINE$, xZ + 1, 10) = XSCRNDate$(FF1Hdr.CHGDTMD, FF1Hdr.CHGDTY)
   Call WriteLINE
   Call GetRVAR("CHGNAME", X)
   Mid$(PLINE$, Y, 255) = RVAR$(X): xZ = InStr(Y, PLINE$, "  ")
   Mid$(PLINE$, xZ, 6) = Str$(FF1Hdr.CHGCNT1)
   Call WriteLINE
   Call GetRVAR("CHGXNAME", X)
   Mid$(PLINE$, Y, 255) = RVAR$(X): xZ = InStr(Y, PLINE$, "  ")
   Mid$(PLINE$, xZ, 6) = Str$(FF1Hdr.CHGCNT2)
   Call WriteLINE: Call WriteLINE
   PRT.SECTION = "D"
   Call PrintTITLE
   For ZZ = 1 To 12
       xZZ$ = "RTYPE" + LTrim$(RTrim$(Str$(ZZ)))
       Call GetRVAR(xZZ$, X): RTYPE$(ZZ) = RVAR$(X)
   Next ZZ
   Mid$(PLINE$, 2, 32) = "1 " + RTYPE$(1)
   Mid$(PLINE$, 35, 5) = SFILL$(STOT(0, 1), 5)
   Mid$(PLINE$, 45, 5) = SFILL$(STOT(1, 0), 5)
   Call WriteLINE
   Call GetRVAR("FATHER", X)
   Mid$(PLINE$, 6, 30) = RVAR$(X)
   Mid$(PLINE$, 35, 5) = SFILL$(STOT(0, 2), 5)
   Call WriteLINE
   Call GetRVAR("MOTHER", X)
   Mid$(PLINE$, 6, 30) = RVAR$(X)
   Mid$(PLINE$, 35, 5) = SFILL$(STOT(0, 3), 5)
   Call WriteLINE
   Call GetRVAR("BOTHPAR", X)
   Mid$(PLINE$, 6, 30) = RVAR$(X)
   Mid$(PLINE$, 35, 5) = SFILL$(STOT(0, 4), 5)
   Call WriteLINE
   Call GetRVAR("CHILDREN", X)
   Mid$(PLINE$, 6, 30) = RVAR$(X)
   Mid$(PLINE$, 35, 5) = SFILL$(STOT(0, 5), 5)
   Call WriteLINE
   Mid$(PLINE$, 6, 30) = RTYPE$(4)
   Mid$(PLINE$, 35, 5) = SFILL$(STOT(1, 4), 5)
   Mid$(PLINE$, 65, 5) = SFILL$(STOT(4, 1) + STOT(12, 1), 5)
   Call WriteLINE
   Mid$(PLINE$, 6, 30) = RTYPE$(5)
   Mid$(PLINE$, 35, 5) = SFILL$(STOT(1, 5), 5)
   Mid$(PLINE$, 65, 5) = SFILL$(STOT(5, 1), 5)
   Call WriteLINE
   Mid$(PLINE$, 2, 32) = "2 " + RTYPE$(2)
   Mid$(PLINE$, 35, 5) = SFILL$(STOT(1, 2), 5)
   Mid$(PLINE$, 55, 5) = SFILL$(STOT(2, 1), 5)
   Call WriteLINE
   Mid$(PLINE$, 6, 30) = RTYPE$(5)
   Mid$(PLINE$, 35, 5) = SFILL$(STOT(4, 5), 5)
   RSet S5$ = LTrim$(Str$(FFRBA&(STOT(5, 4), 1)))
   Mid$(PLINE$, 65, 5) = S5$
   Call WriteLINE
   Mid$(PLINE$, 2, 32) = "3 " + RTYPE$(3)
   Mid$(PLINE$, 35, 5) = SFILL$(STOT(1, 3), 5)
   Mid$(PLINE$, 65, 5) = SFILL$(STOT(3, 1), 5)
   Call WriteLINE
   Mid$(PLINE$, 6, 30) = RTYPE$(4)
   Mid$(PLINE$, 35, 5) = SFILL$(STOT(3, 4), 5)
   Mid$(PLINE$, 65, 5) = SFILL$(STOT(4, 3) + STOT(12, 3), 5)
   Call WriteLINE
   Mid$(PLINE$, 6, 30) = RTYPE$(5)
   Mid$(PLINE$, 35, 5) = SFILL$(STOT(3, 5), 5)
   Mid$(PLINE$, 65, 5) = SFILL$(STOT(5, 3), 5)
   Call WriteLINE
   Mid$(PLINE$, 6, 30) = RTYPE$(2)
   Mid$(PLINE$, 35, 5) = SFILL$(STOT(3, 2), 5)
   Mid$(PLINE$, 55, 5) = SFILL$(STOT(2, 3), 5)
   Call WriteLINE
   Mid$(PLINE$, 8, 30) = RTYPE$(5)
   Mid$(PLINE$, 35, 5) = SFILL$(STOT(4, 6), 5)
   Mid$(PLINE$, 65, 5) = SFILL$(STOT(6, 4), 5)
   Call WriteLINE
   Mid$(PLINE$, 2, 32) = "4 " + RTYPE$(4)
   Mid$(PLINE$, 35, 5) = SFILL$(STOT(4, 0), 5)
   Call WriteLINE
   Mid$(PLINE$, 2, 32) = "5 " + RTYPE$(5)
   RSet S5$ = LTrim$(Str$(FFRBA&(STOT(5, 0), 1)))
   Mid$(PLINE$, 35, 5) = S5$
   Call WriteLINE
   For RT = 6 To 11: GoSub PrintTYPE: Next RT
   Mid$(PLINE$, 1, 33) = "12 " + RTYPE$(12)
   RSet S5$ = LTrim$(Str$(FFRBA&(STOT(12, 0), 1)))
   Mid$(PLINE$, 35, 5) = S5$
   Call WriteLINE
   Call WriteLINE
   Call GetRVAR("TOTAL1", X)
   Mid$(PLINE$, 4, 30) = RVAR$(X)
   Mid$(PLINE$, 45, 5) = SFILL$(STOT(1, 0), 5)
   Mid$(PLINE$, 55, 5) = SFILL$(STOT(2, 0), 5)
   RSet S5$ = LTrim$(Str$(FFRBA&(STOT(3, 0), 1)))
   Mid$(PLINE$, 65, 5) = S5$
   Call WriteLINE
   Call GetRVAR("TOTAL2", X)
   Mid$(PLINE$, 4, 30) = RVAR$(X)
   Mid$(PLINE$, 55, 5) = SFILL$(STOT(0, 9), 5)
   Mid$(PLINE$, 65, 5) = SFILL$(STOT(0, 10), 5)
   Call WriteLINE
   Call RptCLOSE
   Call FamCLOSE
   frmFHSWRPgm.MousePointer = 0
   Exit Sub

PrintTYPE:
   Mid$(PLINE$, 1, 2) = SFILL$(RT, 2)
   Mid$(PLINE$, 4, 30) = RTYPE$(RT)
   Mid$(PLINE$, 35, 5) = SFILL$(STOT(1, RT), 5)
   Mid$(PLINE$, 65, 5) = SFILL$(STOT(RT, 1), 5)
   Call WriteLINE
   If (RT = 6) Or (RT = 11) Then
      Mid$(PLINE$, 6, 30) = RTYPE$(12)
      Mid$(PLINE$, 65, 5) = SFILL$(STOT(12, RT), 5)
      Call WriteLINE
   End If
   Mid$(PLINE$, 6, 30) = RTYPE$(5)
   If (RT = 6) Or (RT = 11) Then Return
   Mid$(PLINE$, 35, 5) = SFILL$(STOT(RT, 5), 5)
   Mid$(PLINE$, 65, 5) = SFILL$(STOT(5, RT), 5)
   Call WriteLINE
   Mid$(PLINE$, 6, 30) = RTYPE$(2)
   Mid$(PLINE$, 35, 5) = SFILL$(STOT(RT, 2), 5)
   Mid$(PLINE$, 55, 5) = SFILL$(STOT(2, RT), 5)
   Call WriteLINE
   Mid$(PLINE$, 8, 30) = RTYPE$(5)
   Mid$(PLINE$, 35, 5) = SFILL$(STOT(4, RT), 5)
   Mid$(PLINE$, 65, 5) = SFILL$(STOT(RT, 4), 5)
   Call WriteLINE
   Return
End Sub





Sub RptLIST(OPT$)
   PN$ = " FAMILY HISTORY SYSTEM - FHSWLIST - Search/Select/LIST Program"
   CY$ = " (C) Copyright 1986-1998 by Phillip E. Brown"
   
   Select Case OPT$
      Case "INIT": GoTo 40
      Case "PRINT": GoTo 1000
   End Select
   
40 ' Format Display
   Call PgmPREP("FHSLIST")
   Load frmFHSWRPgm
   If Left$(RptCODE, 1) = "I" Then
      RptLOAD ("INDX"): RptLOAD ("ISUM")
      frmFHSWRPgm.RPSetup.AddItem "INDX"
      frmFHSWRPgm.RPSetup.AddItem "ISUM"
      Else
      If RptCODE = "FSUM" Then
         RptLOAD ("FSUM")
         frmFHSWRPgm.RPSetup.AddItem "FSUM"
         Else
         If Right$(RptCODE, 3) = "SUM" Then
            RptLOAD ("ISUM")
            RptLOAD ("ASUM"): RptLOAD ("DSUM"): RptLOAD ("RSUM")
            frmFHSWRPgm.RPSetup.AddItem "ISUM"
            frmFHSWRPgm.RPSetup.AddItem "ASUM"
            frmFHSWRPgm.RPSetup.AddItem "DSUM"
            frmFHSWRPgm.RPSetup.AddItem "RSUM"
            Else
            RptLOAD ("LIST"): RptLOAD ("INDX")
            frmFHSWRPgm.RPSetup.AddItem "LIST"
            frmFHSWRPgm.RPSetup.AddItem "INDX"
        End If
      End If
   End If
   frmFHSWRPgm.RPSetup_Click
   frmFHSWRPgm.Show
   Exit Sub

1000 'Print Reports
   If (SCNT = 0) And (RptCODE <> "FSUM") Then
      Call PutMSG("LISTM005")
      Exit Sub
   End If
   ReDim DFNUM(1), DFPP(1)
   ReDim RELATE(1), RELPTR&(1)
   ReDim NDXID(1)
   MAXMarry = 10: ReDim MLIST(MAXMarry)
   DetailRptTITLE$ = RTrim$(RDFReport.NAME)

   SXK$ = " ": SNDX$ = "    ": SNDXCHK$ = SNDX$
   KeyWORD$ = Space$(30): SrchWORD$ = Space$(30)
   SXTB$ = "01230120022455012623010202"
   'Call SelectSET("DEFAULT")

   ' Print Reports
    If (SlctCNT > 0) And (SecSlctCNT > 0) Then
       Call PutMSG("LISTM006")
       If a = 27 Then Exit Sub
       xSEL = Val(CA$)
       If (xSEL < 1) Or (xSEL > 3) Then xSEL = 3
       Else
       xSEL = 3
    End If
    Select Case RptCODE
       Case "LIST": Call DETAILReport
       Case "INDX": Call INDEXReport
       Case "FSUM": Call FSUMReport
       Case Else: Call XSUMReport
    End Select
End Sub

Rem $DYNAMIC
Sub BldDetDEF()
      RFCNT = 29
      ReDim DFNUM(RFCNT), DFPP(RFCNT)
      ReDim RFNUM(RFCNT), RFSEP(RFCNT), RFLTH(RFCNT), RFLBL$(RFCNT)
      Call RptOptGET("MNGP", MINGAP)
      If (Not OKAY) Or (MINGAP = 0) Then MINGAP = 2
      Call RptOptGET("MXGP", MAXGAP)
      If (Not OKAY) Or (MAXGAP = 0) Then MAXGAP = 5
      Call RptOptGET("MINM", MINMLTH)
      If (Not OKAY) Then MINMLTH = 35
      If MINMLTH > 52 Then MINMLTH = 52
      Call RptOptGET("PNML", PLNMLTH)
      If Not OKAY Then PLNMLTH = 22 - 8 * ENV.LPNSupport
      If RptOPTION.PRID Then
         Call RptOptGET("PPID", X): PrintPID = (X = 1)
      End If
      Call RptOptGET("SNDX", X): PrintSNDX = (X = 1)
      If INDEXED Then
         ts = ENV.SortOPTS And 3: DS = (ENV.SortOPTS \ 4) And 7
      '   FOR X = 1 TO SFCNT
      '       Y = SX(X) \ 256
      '       SELECT CASE Y
      '          CASE 1: X1 = X
      '          CASE 2: X2 = X
      '       END SELECT
      '   NEXT X
      '   IF X1 > 0 THEN
      '      RPTOPTION.NAMEFMT = (RPTOPTION.NAMEFMT AND 1) - 2 * (TS = 1)
      '      IF X2 > 0 THEN
      '         RPTOPTION.NAMEFMT = (RPTOPTION.NAMEFMT AND 2) - (X1 < X2)
      '      END IF
      '   END IF
      End If

59500 ' Select Report Fields - Based Upon Report Options
      ' First Select Prefix Fields
      OKAY = 0
      PAGE.DFCNT = 0: RptOPTION.TWOLINES = 0: PRID2 = 0
      If (PAGE.BWIDTH = 0) Then
         BWIDTH = PRT.FWIDTH - PAGE.IMARGIN - PAGE.OMARGIN
         Else
         BWIDTH = PAGE.BWIDTH
      End If
      If RptOPTION.Relation Then
         PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 3
      End If
      If Not INDEXED Then
         PSK = 9
         Else
         PSK = SX(1) \ 256
         PAGE.DFCNT = PAGE.DFCNT - ((PSK > 2) And (PSK < 8))
         If (ts > 1) And ((PSK > 5) And (PSK < 8)) Then
            DFNUM(PAGE.DFCNT) = 24
            PAGE.DFCNT = PAGE.DFCNT + 1: PSNDX = PAGE.DFCNT
         End If
         Select Case PSK
            Case 3: DFNUM(PAGE.DFCNT) = 12: BFX = PAGE.DFCNT
            Case 4: DFNUM(PAGE.DFCNT) = 15: MFX = PAGE.DFCNT
            Case 5: DFNUM(PAGE.DFCNT) = 19: DFX = PAGE.DFCNT
            Case 6: DFNUM(PAGE.DFCNT) = 13: BFP = PAGE.DFCNT
            Case 7: DFNUM(PAGE.DFCNT) = 20: DFP = PAGE.DFCNT
            Case 8: DFNUM(PAGE.DFCNT) = 10: PSEX = PAGE.DFCNT
            Case 9: INDEXED = 0
         End Select
      End If
      If RptOPTION.PRID And (RptOPTION.FREEFORM Or (Not INDEXED)) Then
         PAGE.DFCNT = PAGE.DFCNT + 1: RFSEP(6) = MINGAP + 1
         DFNUM(PAGE.DFCNT) = 6: IDFX = PAGE.DFCNT
         If ((RptOPTION.SPOUSE And 3) > 0) Then
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 9
            Else
            RFLTH(9) = 0
         End If
         If PrintPID Then
            DFNUM(PAGE.DFCNT + 1) = 7: DFNUM(PAGE.DFCNT + 2) = 8
            PAGE.DFCNT = PAGE.DFCNT + 2
         End If
      End If
      If INDEXED And (ts > 1) And (PSK = 1) Then
         PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 24
         PSNDX = PAGE.DFCNT
      End If

      ' Select Data Fields
      RFSEP(DFNUM(1)) = 1: PAGE.PFCNT = PAGE.DFCNT
      If (MINMLTH > 0) Or RptOPTION.FREEFORM Then
         PAGE.DFCNT = PAGE.DFCNT + 1
         DFNUM(PAGE.DFCNT) = 5: RFSEP(5) = MINGAP + 1
         RFLTH(5) = MINMLTH - (52 - MINMLTH) * RptOPTION.FREEFORM
      End If
      If RptOPTION.FREEFORM Then GoTo 59510
      If PrintSNDX And (PSNDX = 0) Then
         PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 24
      End If
      If RptOPTION.PRID And Not ((PSK = 9)) Then
         PAGE.DFCNT = PAGE.DFCNT + 1: IDFX = PAGE.DFCNT
         DFNUM(PAGE.DFCNT) = 6
         If PrintPID Then
            DFNUM(PAGE.DFCNT + 1) = 7
            DFNUM(PAGE.DFCNT + 2) = 8
            PAGE.DFCNT = PAGE.DFCNT + 2
         End If
         If ((RptOPTION.SPOUSE And 1) > 0) Then
            PAGE.DFCNT = PAGE.DFCNT + 1
            DFNUM(PAGE.DFCNT) = 9
            Else
            RFLTH(9) = 0
         End If
      End If
      If RptOPTION.TIMELINE Then
         DFNUM(PAGE.DFCNT + 1) = 22: DFNUM(PAGE.DFCNT + 2) = 23
         PAGE.DFCNT = PAGE.DFCNT + 2
         GoTo 59510
      End If
      If RptOPTION.PRSEX And (PSEX = 0) Then
         If SFX = 0 Then
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 10
         End If
      End If
      Call RptOptGET("#CH", X): PNUMChild = (X = 1)
      If PNUMChild Then
         PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 26
      End If
      If RptOPTION.PRAGE Then
         PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 11
      End If
      If RptOPTION.DateS Then
         If BFX = 0 Then
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 12: BFX = PAGE.DFCNT
         End If
         If RptOPTION.SPOUSE Then
            If MFX = 0 Then
               PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 15: MFX = PAGE.DFCNT
               If RptOPTION.PRAGE Then
                  PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 18: AFX = PAGE.DFCNT
               End If
            End If
         End If
         If DFX = 0 Then
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 19: DFX = PAGE.DFCNT
         End If
      End If
      If RptOPTION.PLACES Then
         If BFP = 0 Then
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 13
            BFP = PAGE.DFCNT
         End If
         If RptOPTION.SPOUSE Then
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 16: MFP = PAGE.DFCNT
         End If
         If DFP = 0 Then
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 20
            DFP = PAGE.DFCNT
         End If
      End If

59510 ' Create Report Definition Tables
      X = 1: RFNAME$ = "06010": GoSub 59505     'COMMENTS
      X = 2: RFNAME$ = "07015": GoSub 59505     'SOUNDEX
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
      X = 13: RFNAME$ = "01032": RFLTH(13) = PLNMLTH
              GoSub 59505                       'BIRTH PLACE
      X = 14: RFNAME$ = "01034": RFLTH(14) = PLNMLTH
              GoSub 59505    'BIRTH DATE & PLACE
      X = 15: RFNAME$ = "02020": GoSub 59505    'MARRIAGE DATE
      X = 16: RFNAME$ = "02021": RFLTH(16) = PLNMLTH
              GoSub 59505                       'MARRIAGE PLACE
      X = 17: RFNAME$ = "02024": RFLTH(17) = PLNMLTH
              GoSub 59505                       'MARRIAGE DATE & PLACE
      X = 18: RFNAME$ = "02015": GoSub 59505    'Last ANNIVERSARY
      X = 19: RFNAME$ = "01040": GoSub 59505    'DEATH DATE
      X = 20: RFNAME$ = "01042": RFLTH(20) = PLNMLTH
              GoSub 59505                       'DEATH PLACE
      X = 21: RFNAME$ = "01044": RFLTH(21) = PLNMLTH
              GoSub 59505                       'DEATH DATE & PLACE
      X = 22: RFNAME$ = "07020": GoSub 59505    'LIFE SPAN
      X = 23: RFNAME$ = "07021": GoSub 59505    'LIFE LINE
      X = 24: RFNAME$ = "07015": GoSub 59505    'SOUNDEX
      X = 25: RFNAME$ = "07003": GoSub 59505    'RGL
      X = 26: RFNAME$ = "01027": GoSub 59505    'Child COUNT
              If RptOPTION.FREEFORM And FOUND Then
                 RFLBL$(29) = MID$(RFLD$(FX), ALTH + LLTH + 1, HLTH)
              End If
      X = 27: RFNAME$ = "03030": GoSub 59505    'Phone
      X = 28: RFNAME$ = "03031": GoSub 59505    'Address Line
      GoTo 59520

59505 ' Process Report Field
      Call GetRFLD(RFNAME$, FX)
      If RFSEP(X) = 0 Then RFSEP(X) = MINGAP
      If FOUND Then
         RDFFld = RFLDS(FX)
         If RFLTH(X) = 0 Then RFLTH(X) = Asc(RDFFld.RLTH)
         ALTH = Asc(RDFFld.ALTH): LLTH = Asc(RDFFld.LLTH)
         HLTH = Asc(RDFFld.HLTH)
         If (X = 1) Or (X > 26) Or (RptOPTION.FREEFORM And (((X > 9) And (X < 25)) Or (X = 26))) Then
            RFLBL$(X) = MID$(RFLD$(FX), ALTH + 1, LLTH)
            Else
            If X = 26 Then
               RFLBL$(26) = MID$(RFLD$(FX), 1, ALTH)
               Else
               RFLBL$(X) = MID$(RFLD$(FX), ALTH + LLTH + 1, HLTH)
            End If
         End If
      End If
      Return

59520 ' Check Report Width
      RWIDTH = 0: FX = 0
      While FX < PAGE.DFCNT
          FX = FX + 1: X = DFNUM(FX): FSEP = RFSEP(X)
          If (X > 0) And (FSEP > 0) Then
             Y = RFLTH(X): Z = Len(RFLBL$(X)): If Y < Z Then Y = Z
             RWIDTH = RWIDTH + FSEP + Y
          End If
      Wend
      RWIDTH = RWIDTH - 1
      If RWIDTH <= BWIDTH Then GoTo 59540
      If RptOPTION.TIMELINE Or RptOPTION.FREEFORM Then GoTo 59530
      If (Not RptOPTION.PLACES) Or (Not RptOPTION.DateS) Then GoTo 59530
      If (Not RptOPTION.TWOLINES) Then
         If (BFX > 2) And (DFX > 2) And (MFX > 2) Then
            RptOPTION.TWOLINES = -1
            DFNUM(BFX) = 14: DFNUM(BFP) = 0
            DFNUM(DFX) = 21: DFNUM(DFP) = 0
            If RptOPTION.SPOUSE Then
               DFNUM(MFX) = 17: DFNUM(MFP) = 0
               If RptOPTION.PRAGE Then DFNUM(AFX) = 0
            End If
            GoTo 59520
         End If
         GoTo 59530
      End If
      If RptOPTION.PRID And (Not PRID2) And (INDEXED) Then
         PRID2 = -1
         DFNUM(IDFX) = 0: DFNUM(IDFX + 1) = 0: DFNUM(IDFX + 2) = 0
         If ((RptOPTION.SPOUSE And 1) > 0) Then DFNUM(IFDFX + 3) = 0
         X = RFLTH(6) + RFLTH(7) + RFLTH(8) + RFLTH(9) + 18
         If X > RFLTH(5) Then RFLTH(5) = X
         GoTo 59520
      End If

59530 ' Final Check After Trying TWOLINES and PRID2 Options
      PWIDTH = BWIDTH + PAGE.IMARGIN + PAGE.OMARGIN
      RPWIDTH = RWIDTH + PAGE.IMARGIN + PAGE.OMARGIN
      If PWIDTH < RPWIDTH Then
         SysVAR.NVAR1 = PWIDTH: SysVAR.NVAR2 = RPWIDTH
         Call PutMSG("XXXXM009")
         If CA$ <> SNGLKEY$(1) Then GoTo 59560
         BWIDTH = RWIDTH
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
             Z = FL: If Z < HL Then Z = HL
             PP = PP + FSEP + (Z + 1) \ 2 - (FL + 1) \ 2
             DFPP(FX) = PP
             PP = PP + Z
          End If
      Wend

59560 ' Return to Calling Program
      PAGE.BWIDTH = BWIDTH
      If RptOPTION.PRID Then
         If PRID2 Then RptOPTION.PRID = 2 Else RptOPTION.PRID = 1
      End If
      OKAY = -1
End Sub

Rem $STATIC
Sub BldDetTITLE()
     ' Build Title Line
      TL$ = Space$(PAGE.BWIDTH)
      For FX = 1 To PAGE.DFCNT: X = DFNUM(FX): PP = DFPP(FX)
         If (X > 0) And (PP > 0) Then
            FL = RFLTH(X): HL = Len(RFLBL$(X))
            If X = 23 Then
               Mid$(TL$, PP, 4) = fnx$(TLBY)
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

Sub BldMLIST()
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

Sub BldSumDEF(WTYPE)
      If WTYPE < 4 Then
         AMAXLV = Asc(WFHDR.AMAXLV): DMAXLV = Asc(WFHDR.DMAXLV)
         RELSUMOPT = (AMAXLV > 0) And (DMAXLV > 0)
      End If
      If RELSUMOPT Then RFCNT = AMAXLV + 3 Else RFCNT = 33
      PAGE.DFCNT = 0
      ReDim DFNUM(RFCNT), DFPP(RFCNT)
      ReDim RFNUM(RFCNT), RFSEP(RFCNT), RFLTH(RFCNT), RFLBL$(RFCNT)
      Call RptOptGET("MNGP", MINGAP)
      If (Not OKAY) Or (MINGAP = 0) Then MINGAP = 2
      DFPP(0) = MINGAP
      Call RptOptGET("MXGP", MAXGAP)
      If (Not OKAY) Or (MAXGAP = 0) Then MAXGAP = 5
      If RELSUMOPT Then GoTo 58500
      If WTYPE < 4 Then
         If DMAXLV = 0 Then
            Call RptOptGET("CNOP", X): CNTNOP = (X = 1)
         End If
         Call RptOptGET("CDUP", X): CNTDUP = (X = 1)
         Call RptOptGET("CLIN", X): CNTCLN = (X = 1)
         Else
         Call RptOptGET("SNDX", X): PRTSNDX = (X = 1)
      End If
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
      If WTYPE < 4 Then             ' Relationship Summary Reports
         PAGE.DFCNT = 1: DFNUM(1) = 1
         If RptOPTION.Relation Or VRTCL Then
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 2
         End If
         If CNTCLN Then PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 4
         If CNTNOP Then PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 6
         If CNTDUP Then PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 5
         GoTo 58505
      End If

      ' Select Prefix Fields for INDEXED Summary
      ts = ENV.SortOPTS And 3: DS = (ENV.SortOPTS \ 4) And 7
      PSK = SX(1) \ 256
      If (2 < PSK) And (PSK < 6) Then
         Select Case DS
            Case 0: PFK1 = 26: PFK2 = 27: PFK3 = 28
            Case 1: PFK1 = 27: PFK2 = 28: PFK3 = 26
            Case 2: PFK1 = 28: PFK2 = 27: PFK3 = 26
            Case 3: PFK1 = 26: PFK2 = 27
            Case 4: PFK1 = 27: PFK2 = 28
            Case 5: PFK1 = 26
            Case 6: PFK1 = 27
            Case 7: PFK1 = 28
         End Select
         Else
         If PSK > 7 Then
            PFK1 = 33
            Else
            If ts > 1 Then PFK1 = 24 Else If PRTSNDX Then PFK3 = 24
            Select Case PSK
               Case 1: PFK2 = 29
               Case 2: PFK2 = 30
               Case 6: PFK2 = 31
               Case 7: PFK2 = 32
            End Select
            If PFK2 > 0 Then RFLTH(PFK2) = SX(1) And 127
         End If
      End If
      If PFK1 > 0 Then
         PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = PFK1
      End If
      If PFK2 > 0 Then
         PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = PFK2
      End If
      If PFK3 > 0 Then
         PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = PFK3
      End If

58505 ' Select Counter Fields
      If Not (PSK = 8) Then
         PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 7
         PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 8
      End If
      PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 3
      If Not (PSK = 3) Then
         If DATECNT Then PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 9
         If DATERNG Then
            DFNUM(PAGE.DFCNT + 1) = 11: DFNUM(PAGE.DFCNT + 2) = 12
            PAGE.DFCNT = PAGE.DFCNT + 2
         End If
      End If
      If Not (PSK = 6) Then
         If PLACCNT Then
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 10
         End If
      End If
      If MARRCNT Then
         If Not (PSK = 4) Then
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 13
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 14
            If DATECNT Then
               PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 15
            End If
            If DATERNG Then
               PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 17
               PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 18
            End If
         End If
         If PLACCNT Then
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 16
         End If
      End If
      If Not (PSK = 5) Then
         If DATECNT Then
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 19
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 20
         End If
         If DATERNG Then
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 22
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 23
         End If
      End If
      If Not (PSK = 7) Then
         If PLACCNT Then
            PAGE.DFCNT = PAGE.DFCNT + 1: DFNUM(PAGE.DFCNT) = 21
         End If
      End If

58510 ' Create Report Definition Tables
      X = 1                                     'AGL or DGL
      If WTYPE < 4 Then
         If (DMAXLV > 0) And (AMAXLV = 0) Then
            RFNAME$ = "07002"
            Else
            RFNAME$ = "07001"
         End If
         GoSub 58515
      End If
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
      X = 24: RFNAME$ = "07015": GoSub 58515    'SOUNDEX
      X = 25: RFNAME$ = "08005": GoSub 58515    'SURNMCNT
      X = 26: RFNAME$ = "08009": GoSub 58515    'YEAR
      X = 27: RFNAME$ = "08010": GoSub 58515    'MONTH
      X = 28: RFNAME$ = "08011": GoSub 58515    'DAY
      X = 29: RFNAME$ = "01022": GoSub 58515    'Surname
      X = 30: RFNAME$ = "01021": GoSub 58515    'GivenName
      X = 31: RFNAME$ = "01032": GoSub 58515    'BirthPlace
      X = 32: RFNAME$ = "01042": GoSub 58515    'DeathPlace
      X = 33: RFNAME$ = "01015": GoSub 58515    'Sex Code
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
      DFNUM(1) = 1: RFSEP(1) = 1: RFLTH(1) = 8: DFPP(1) = 1
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
             Y = RFLTH(X): Z = Len(RFLBL$(X))
             If VRTCL Then
                If RWIDTH < Z Then RWIDTH = Z
                Else
                If Y < Z Then Y = Z
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
      If PWIDTH < RPWIDTH Then
         SysVAR.NVAR1 = PWIDTH: SysVAR.NVAR2 = RPWIDTH
         Call PutMSG("XXXXM009")
         If CA$ <> SNGLKEY$(1) Then OKAY = 0: GoTo 58560
         BWIDTH = RWIDTH
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
             Z = FL: If Z < HL Then Z = HL
             PP = PP + FSEP + (Z + 1) \ 2 - (FL + 1) \ 2
             DFPP(FX) = PP
             PP = PP + Z
          End If
      Wend
      OKAY = -1

58560 ' Return to Calling Program
      PAGE.BWIDTH = BWIDTH
End Sub

Sub BldSumTITLE()
      AMAXLV = Asc(WFHDR.AMAXLV): DMAXLV = Asc(WFHDR.DMAXLV)
      RELSUMOPT = (AMAXLV > 0) And (DMAXLV > 0)
     ' Build Summary Report Title Line
      TL$ = Space$(PAGE.BWIDTH): Y = 0
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
         Y = 0: xY$ = Space$(6): If AMAXLV > 0 Then Z = -1 Else Z = 1
         While X > -1
            RSet xY$ = Str$(Y * Z)
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
         HTFLINE$(HTFLMAX) = "T" + PRT.SECTION + "0" + RTrim$(TL$)
         If VRTCL Then
            HTFLMAX = HTFLMAX + 1
            ReDim Preserve HTFLINE$(HTFLMAX)
            HTFLINE$(HTFLMAX) = "T00" + RTrim$(TL2$)
         End If
      End If
      OKAY = -1
End Sub

Sub DETAILReport()
     ' Produce LIST of Selected Individuals
     Call RptOPEN("LIST"): If Not OKAY Then GoTo 1040
1010 Call PutMSG("LISTM007")
     If a = 27 Then GoTo 1040
     RptOPTION.FREEFORM = (CA$ = "2"): RptOPTION.TIMELINE = (CA$ = "3")
1020 If RptOPTION.TIMELINE Then
        'Call FmtFIND("XXXXS030")
        'Call MsgFORMAT: If A = 27 Then GoTo 1010
        'Call FmtFindFLD("TLBY"): If FOUND Then TLBY = Val(UTXT$(FFLD.UTXT))
        'Call FmtFindFLD("TLEY"): If FOUND Then TLEY = Val(UTXT$(FFLD.UTXT))
        Call frmFHSWMsgD.GetTLYears(TLBY, TLEY)
        If a = 27 Then GoTo 1010
        If (TLEY - TLBY) < 1 Then Call ErrBEEP(0): GoTo 1020
     End If

     Call PutMSG("LISTM008")
     If Asc(CA$) = 27 Then GoTo 1040
     a = Val(CA$): If a < 1 Or a > 2 Then a = 2
     INDEXED = (a = 1)
     If Not INDEXED Then
        SFCNT = 1: ReDim SX(1):
        SX(1) = 9 * 256: NDXRECS = FF1Hdr.MAXID
        RptOPTION.PRID = -1: SXK = 0
        Else
        Call INDEXFile("READ")
        If Not OKAY Then GoTo 1040
        NDXRECS = NDXHDR.NDXRECS
        ts = ENV.SortOPTS And 3: DS = (ENV.SortOPTS \ 4) And 7
        SXK = SX(1) \ 256
        If ((SXK > 2) And (SXK < 6)) Or (SXK > 7) Then SXK = 1
     End If

     If RptOPTION.Relation Then
        Call RptOptGET("EXAD", X): ExtAdoptRule = (X = 1)
        Call RWRKOpen("READ", 0, 0): RptOPTION.Relation = OKAY
        If OKAY Then
           ReDim RELATE(FF1Hdr.MAXID)
           Call PutMSG("XXXXM064")
           Call RWRKGet("BASE", RC, PTR&)
           While (0 < RC) And (RC < 9)
              If WFRELREC.id > 0 Then
                 X = RELATE(WFRELREC.id)
                 If (X = 0) Or (WFGLREC.DGL = 0) Then
                    X = 256 * (-WFGLREC.AGL) + WFGLREC.DGL
                    If ExtAdoptRule Then X = X - 128 * ((WFRELREC.CSTAT And 5) > 0)
                    RELATE(WFRELREC.id) = X
                 End If
              End If
              Call RWRKGet("WRKREC", RC, PTR&)
           Wend
           Call FamCLOSE
        End If
     End If
     'Call FmtFIND("LISTS009")
     'Call FmtFindFLD("TITL")
     'If FOUND Then
     '   If Len(STITLE$) > 0 Then
     '      LSet UTXT$(FFLD.UTXT) = STITLE$
     '      Else
     '      LSet UTXT$(FFLD.UTXT) = ENV.RTITLE
     '   End If
     '   'Call MsgFORMAT
     '   STITLE$ = RTrim$(UTXT$(FFLD.UTXT))
     '   If (Len(STITLE$) > 0) Then
     '      LSet ENV.RTITLE = STITLE$
     '   End If
     'End If
     STITLE$ = ENV.RTITLE
     BD$ = Space$(12): DD$ = Space$(12): AGE$ = "   "
     BPL$ = Space$(41): MPL$ = BPL$: DPL$ = BPL$: CL$ = Space$(255)
     D1D$ = BD$: D2D$ = BD$: D1PL$ = BPL$: D2PL$ = BPL$: AGE2$ = AGE$
     NUMED$ = Space$(6): SNDX$ = "    ": SNDXText$ = Space$(30)
     SURNM$ = Space$(20): ADRDATA$ = Space$(135)
     Call FAMOpen("READ")
     AFMT = FF2Hdr.REVLVL
     If OKAY Then GoTo 1200

1040 ' Bad Return
     Close: ENV.FFOPEN = 0: Exit Sub

1200 ' Setup Report Tables
     Call BldDetDEF: If Not OKAY Then GoTo 1040
     PRDAGE = (RptOPTION.PRAGE = 2)
     PRID2 = ((RptOPTION.PRID And 2) > 0)
     RptOPTION.PRID = (RptOPTION.PRID <> 0)
     Call RptOptGET("#CH", X): PNUMChild = (X = 1)
     Call RptOptGET("LRES", X): PLastRES = (X = 1)
     Call RptOptGET("EURA", X): EuroAdr = (X = 1)
     Call RptOptGET("PNML", PLNMLTH)
     If PLNMLTH = 0 Then PLNMLTH = 22 - 8 * ENV.LPNSupport
     Call PrintOPEN: If Not OKAY Then GoTo 1040
     PLINE$ = Space$(255): PL2$ = PLINE$
     xRGL$ = Space$(RFLTH(25)): xLINEAGE$ = Space$(RFLTH(2))
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
     CINDENT = PAGE.CINDENT: If RptOPTION.FREEFORM Then PAGE.CINDENT = 0

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

1400 ' Begin with First Record
     a = 0: PAGE.CSID = 0: xNDX = 0

1415 ' Process next entry on chart
     If Not (xNDX < NDXRECS) Then GoTo 1490
     xNDX = xNDX + 1
     If INDEXED Then INDID = NDXID(xNDX) Else INDID = xNDX
     If (SelSET(INDID) And xSEL) = 0 Then GoTo 1415

1450 ' Process Individual record
     Call FF1GetRec(INDID)
     If PLastRES Then Call GetLastRES
     If RptOPTION.Relation Then
        If RELATE(INDID) <> 0 Then
           X = RELATE(INDID): Y = X And 128
           AGL = -(X \ 256): DGL = X And 127
           RGL = AGL + DGL
           LSet XRELATION$ = Relation$("M", FF1Rec.SEX, AGL, RGL, RRELCD$())
           If ExtAdoptRule Then
              If (X And 128) > 0 Then LSet XRELATION$ = "*" + XRELATION$
           End If
           Else
           AGL = 0: DGL = 0: RGL = 0: LSet XRELATION$ = ""
        End If
     End If
     If ((RptOPTION.BLNKSEP) And (PAGE.LINE > 0)) Then Call WriteLINE
     PAGE.RPAGE = 0: PAGE.RLINE = 0
     PAGE.XID = INDID: PAGE.xType = 1
     RREFREC.AGL = Chr$(Abs(AGL)): RREFREC.DGL = Chr$(DGL): Call RREFAdd
     ComPTR = FF1Rec.COM
     Call GetBDEvents(BD$, BPL$, DD$, DPL$, AGE$)
     BDY = FF1Rec.BY: DDY = FF1Rec.DY
     If PNUMChild Then GoSub 1510
     If RptOPTION.SPOUSE Then
        PAGE.CSID = 0: MLOC = 0: MCOMPTR = 0
        SDY = 0: SDMD = 0: M2Y = 0: M2MD = 0
        SPRNO = FF1Rec.SPOUSE: S1DY = FF1Rec.DY: S1DMD = FF1Rec.DMD
        If ASPOUSE And (SPRNO <> 0) Then
           GoSub 1530
           Else
           If ALLSPOUSE Then Call BldMLIST
        End If
     End If
     GoSub 1800   ' Process DATA Fields
     If a = 27 Then GoTo 1490

     If Not ALLSPOUSE Then GoTo 1470  ' Didn't Request ALL Marriages

     ' Process Marriage Records
     While SPRNO <> 0
        PAGE.RPAGE = 0: PAGE.RLINE = 0
        PAGE.xType = 3: GoSub 1530
        PAGE.XID = PAGE.CSID: Call RREFAdd
        BDY = FF1Rec.BY: DDY = FF1Rec.DY
        LSet BD$ = D1D$: LSet BPL$ = D1PL$
        LSet DD$ = D2D$: LSet DPL$ = D2PL$: LSet AGE$ = AGE2$
        GoSub 1800
        If a = 27 Then SPRNO = 0
        SDY = 0: SDMD = 0
     Wend

1470 ' Finish Line and Go Get Another WF Record
     Call PrintLINE
     If a = 27 Then GoTo 1490
     GoTo 1415

1490 ' END OF REPORT
     Call RptCLOSE: If PRT.AGAIN Then GoTo 1400
     Call FamCLOSE
     PAGE.CINDENT = CINDENT
     Exit Sub

1510 ' Count Children
     X = FF1Rec.OLDCH: NUMCH = 0
     While X <> 0
        NUMCH = NUMCH + 1
        Call FF1GetRec(X)
        If FF1Rec.FID = INDID Then
           X = FF1Rec.FCH
           Else
           If FF1Rec.MID = INDID Then
              X = FF1Rec.MCH
              Else
              X = 0
           End If
        End If
     Wend
     Call FF1GetRec(INDID)
     Return

1530 ' Get Information for SPRNO
     If (SPRNO <> 0) Or (PAGE.CSID > 0) Then
        If SPRNO <> 0 Then Call FF3GetRec(SPRNO)
        MCOMPTR = FF3SPOUSE.COM: MLOC = FF3SPOUSE.MLOC
        M1Y = FF3SPOUSE.D1Y: M1MD = FF3SPOUSE.D1MD
        M2Y = FF3SPOUSE.D2Y: M2MD = FF3SPOUSE.D2MD
        If FF3SPOUSE.SP1ID = INDID Then
           PAGE.CSID = FF3SPOUSE.SP2ID
           Else
           PAGE.CSID = FF3SPOUSE.SP1ID
        End If
        If ALLSPOUSE Then Call GetNextMARRIAGE Else SPRNO = 0
        If PAGE.CSID > 0 Then
           Call FF1GetRec(PAGE.CSID)
           Call GetBDEvents(D1D$, D1PL$, D2D$, D2PL$, AGE2$)
           If ALLSPOUSE Then ComPTR = FF1Rec.COM
           SDY = FF1Rec.DY: SDMD = FF1Rec.DMD
           If PAGE.xType = 1 Then
              Call FF1GetRec(INDID)
              Else
              If RptOPTION.Relation Then LSet XRELATION$ = " "
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

1545 ' Determine Last Anniversary
     If ((M2Y > 0 Or M2MD > 0) Or (S1DY = 0 And SDY = 0)) Then
        LSet ANVSRC$ = " "
        Else
        M2Y = SDY: M2MD = SDMD: LSet ANVSRC$ = "*"
        If ((S1DY > 0) And ((SDY = 0) Or ((S1DY < SDY) Or (S1DY = SDY And S1DMD < SDMD)))) Then M2Y = S1DY: M2MD = S1DMD
     End If
     RSet ANNV$ = DATEDif$(M1MD, M1Y, M2MD, M2Y)
     Return

1800 ' PRINT Prefix Fields (Those Before NAME on Data Line)
     If (PAGE.xType = 1) Or (PAGE.CSID > 0) Then
        LSet NFMT$ = FMTNAME$(0)
        Else
        LSet NFMT$ = "(" + Unknown$ + ")"
     End If
     COMBGN = 0
     FX = 0
     While FX < PAGE.PFCNT
        FX = FX + 1: DFN = DFNUM(FX): X = DFPP(FX)
        If (DFN > 0) And (X > 0) Then
           FLTH = RFLTH(DFN)
           PAGE.LOFST = X
           GoSub 1810
        End If
     Wend
                              
1805 ' Process DATA Items After NAME Field on DATA Line
     If RptOPTION.FREEFORM Then GoSub 2800 Else GoSub 2500
     If Not RptOPTION.FREEFORM Then
        Call PrintLINE
        If RptOPTION.TWOLINES Then
           LSet PLINE$ = PL2$
           Call PrintLINE: LSet PL2$ = " "
        End If
     End If

     ' Print Comments - Individual and Marriage as appropriate
     If RptOPTION.COMMENTS Then
        COMBGN = 0: PAGE.LWIDTH = PAGE.BWIDTH - PAGE.CINDENT
        If (PAGE.xType = 3) And MCOMPTR Then Call ComPRINT(MCOMPTR, COMBGN)
        If ComPTR Then Call ComPRINT(ComPTR, COMBGN)
        If (PAGE.xType = 1) And MCOMPTR Then Call ComPRINT(MCOMPTR, COMBGN)
     End If
     Call PrintLINE
     If (PLastRES And (ADRSPTR <> 0)) Then
        If (ALLSPOUSE And (ADRSSPID = PAGE.CSID)) Or (Not ALLSPOUSE) Then
           GoSub 1900
        End If
     End If
     Return

1810 ' Process Data Fields by DFN
     If DFN < 10 Then
        Select Case DFN
           Case 3: GoTo 1823
           Case 5: GoTo 1825
           Case 6: GoTo 1826
           Case 7: GoTo 1826
           Case 8: GoTo 1826
           Case 9: GoTo 1826
        End Select
        Else
        If DFN < 21 Then
           Select Case DFN
              Case 10: GoTo 2501
              Case 11: GoTo 2502
              Case 12: GoTo 2510
              Case 13: GoTo 2515
              Case 14: GoTo 2510
              Case 15: GoTo 2520
              Case 16: GoTo 2525
              Case 17: GoTo 2520
              Case 18: GoTo 2521
              Case 19: GoTo 2530
              Case 20: GoTo 2535
           End Select
           Else
           Select Case DFN
              Case 21: GoTo 2530
              Case 22: GoTo 2600
              Case 23: GoTo 2610
              Case 24: GoTo 1822
              Case 25: GoTo 1821
              Case 26: GoTo 1824
           End Select
        End If
     End If
     Return

1820 ' Routines for Processing PREFIX Data Fields
1821 ' AGL, DGL or RGL
     RSet xRGL$ = Str$(RGL)
     Mid$(PLINE$, PAGE.LOFST, FLTH) = xRGL$
     GoTo 1829
1822 ' Soundex
     If (SXK = 0) Or (PAGE.xType = 1) Then
        Select Case SXK
           Case 0: LSet SNDXText$ = SURNM$
           Case 1: LSet SNDXText$ = SURNM$
           Case 2: LSet SNDXText$ = FF1Rec.GIVEN
           Case 6: LSet SNDXText$ = BPL$
           Case 7: LSet SNDXText$ = DPL$
        End Select
        Call ComputeSNDX(SNDXText$, SNDX$)
        Mid$(PLINE$, PAGE.LOFST, FLTH) = SNDX$
     End If
     GoTo 1829
1824 ' Count Children
     If NUMCH > 0 Then Mid$(PLINE$, PAGE.LOFST, FLTH) = SFILL$(NUMCH, FLTH)
     GoTo 1829
1823 ' Relationship
     Mid$(PLINE$, PAGE.LOFST, FLTH) = XRELATION$
     LSet XRELATION$ = " "
     GoTo 1829
1825 ' Name
     NMLTH = FLTH
     If PAGE.xType = 3 Then
        If Not RptOPTION.BLNKSEP Then PAGE.LOFST = PAGE.LOFST + CINDENT
        If Len(SPX$) > 0 Then
           Mid$(PLINE$, PAGE.LOFST, 255) = SPX$
           X = InStr(PAGE.LOFST, PLINE$, "  ") + 1
        End If
        If Not RptOPTION.FREEFORM Then NMLTH = FLTH - (X - PAGE.LOFST)
        PAGE.LOFST = X
     End If
     Mid$(PLINE$, PAGE.LOFST, NMLTH) = NFMT$
     PAGE.COFST = PAGE.LOFST
     If PAGE.xType = 1 Then
        PAGE.COFST = PAGE.COFST - CINDENT * RptOPTION.FREEFORM
     End If
     If PRID2 Then GoSub 1828
     If RptOPTION.FREEFORM Then
        PAGE.LOFST = InStr(PAGE.LOFST, PLINE$, "   ")
     End If
     GoTo 1829
1826 ' ID Numbers - First Data Line
     Select Case DFN
        Case 6: If (ASPOUSE Or (PAGE.CSID = 0) Or (((SX(1) \ 256) <> 9) And (Not RptOPTION.FREEFORM))) Then Mid$(PLINE$, PAGE.LOFST, 5) = SFILL$(FF1Rec.RID, 5)
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

1900 ' Format ADDRESS Line
     If (Not RptOPTION.BLNKSEP) And (Not RptOPTION.FREEFORM) Then PAGE.LOFST = PAGE.LOFST + CINDENT
     LSet ADRDATA$ = "": X = 1
     If AFMT = 0 Then
        If (FF2SREC.FON1 <> 0) Or (FF2SREC.FON2 <> 0) Or (FF2SREC.FON3 <> 0) Then
           Mid$(ADRDATA$, 1, 14) = ZFILL$(FF2SREC.FON1, 3) + "-" + ZFILL$(FF2SREC.FON2, 3) + "-" + ZFILL$(FF2SREC.FON3, 4) + "  "
        End If
        Mid$(ADRDATA$, 15, 30) = LTrim$(RTrim$(FF2SREC.LINE1))
        Mid$(ADRDATA$, 45, 30) = LTrim$(RTrim$(FF2SREC.LINE2))
        Mid$(ADRDATA$, 75, 15) = LTrim$(RTrim$(FF2SREC.CITY))
        Mid$(ADRDATA$, 95, 4) = LTrim$(RTrim$(FF2SREC.STATE))
        Mid$(ADRDATA$, 107, 5) = LTrim$(RTrim$(FF2SREC.ZIP))
        Else
        If (Asc(FF2LREC.PHONE) <> 0) And (Left$(FF2LREC.PHONE, 13) <> "(000)000-0000") Then
           Mid$(ADRDATA$, 1, 14) = LTrim$(RTrim$(FF2LREC.PHONE))
        End If
        Mid$(ADRDATA$, 15, 30) = LTrim$(RTrim$(FF2LREC.LINE1))
        Mid$(ADRDATA$, 45, 30) = LTrim$(RTrim$(FF2LREC.LINE2))
        Mid$(ADRDATA$, 75, 20) = LTrim$(RTrim$(FF2LREC.CITY))
        Mid$(ADRDATA$, 95, 12) = LTrim$(RTrim$(FF2LREC.STATE))
        Mid$(ADRDATA$, 107, 12) = LTrim$(RTrim$(FF2LREC.ZIP))
        Mid$(ADRDATA$, 119, 16) = LTrim$(RTrim$(FF2LREC.COUNTRY))
     End If
     If (FF2PFX.D2Y > 0) Or (FF2PFX.D2MD > 0) Then
        Mid$(CL$, X, 1) = "*": X = X + 1
     End If
     If Left$(ADRDATA$, 1) <> " " Then
        Mid$(CL$, X, 255) = RFLBL$(27): X = InStr(CL$, "  ")
        Mid$(CL$, X, 1) = ":": X = X + 1
        Mid$(CL$, X, 255) = MID$(ADRDATA$, 1, 14): X = InStr(X, CL$, "  ") + 2
     End If
     Mid$(CL$, X, 255) = RFLBL$(28): X = InStr(X, CL$, "  ")
     Mid$(CL$, X, 1) = ":": X = X + 1
     If MID$(ADRDATA$, 15, 1) <> " " Then
        Mid$(CL$, X, 255) = MID$(ADRDATA$, 15, 30): X = InStr(X, CL$, "  ")
        Mid$(CL$, X, 1) = ",": X = X + 2
     End If
     If MID$(ADRDATA$, 45, 1) <> " " Then
        Mid$(CL$, X, 255) = MID$(ADRDATA$, 45, 30): X = InStr(X, CL$, "  ")
        Mid$(CL$, X, 1) = ",": X = X + 2
     End If
     If EuroAdr Then
        If MID$(ADRDATA$, 107, 1) <> " " Then
           Mid$(CL$, X, 255) = MID$(ADRDATA$, 107, 12): X = InStr(X, CL$, "  ")
           X = X + 1
        End If
     End If
     If MID$(ADRDATA$, 75, 1) <> " " Then
        Mid$(CL$, X, 255) = MID$(ADRDATA$, 75, 20): X = InStr(X, CL$, "  ")
        Mid$(CL$, X, 1) = ",": X = X + 2
     End If
     If MID$(ADRDATA$, 95, 1) <> " " Then
        Mid$(CL$, X, 255) = MID$(ADRDATA$, 95, 12): X = InStr(X, CL$, "  ") + 1
        If EuroAdr Then Mid$(CL$, X - 1, 1) = ",": X = X + 1
     End If
     If Not EuroAdr Then
        If MID$(ADRDATA$, 107, 1) <> " " Then
           Mid$(CL$, X, 255) = MID$(ADRDATA$, 107, 12): X = InStr(X, CL$, "  ")
           Mid$(CL$, X, 1) = ",": X = X + 2
        End If
     End If
     If MID$(ADRDATA$, 119, 1) <> " " Then
        Mid$(CL$, X, 255) = MID$(ADRDATA$, 119, 16): X = InStr(X, CL$, "  ")
        Mid$(CL$, X, 1) = ",": X = X + 2
     End If
     If X > 2 Then If MID$(CL$, X - 2, 1) = "," Then Mid$(CL$, X - 2, 1) = " "
     Call PutDATA: Call PrintLINE
     If RptOPTION.COMMENTS Then
        COMBGN = 0: PAGE.LWIDTH = PAGE.BWIDTH - PAGE.CINDENT
        If (FF2PFX.COM <> 0) Then Call ComPRINT(FF2PFX.COM, COMBGN)
     End If
     Call PrintLINE
     Return

2500 ' Print Information for Fixed Format Report - Fields After Prefix
     If (PAGE.xType = 1) Or (PAGE.CSID > 0) Then
        LSet NFMT$ = FMTNAME$(0)
        Else
        LSet NFMT$ = "(" + Unknown$ + ")"
     End If
     FX = PAGE.PFCNT
     While FX < PAGE.DFCNT
        FX = FX + 1: DFN = DFNUM(FX): Y = DFPP(FX)
        If (DFN > 0) And (Y > 0) Then
           FLTH = RFLTH(DFN): PAGE.LOFST = Y
           GoSub 1810
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
           Mid$(PLINE$, PAGE.LOFST, 10) = ANNV$
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
     If (TMTYPE = 1) And (TMX < TMLTH) And (TMDYR < 9999) Then Mid$(TMLINE$, TMX + 1, TMLTH) = TMLIFE$
     If (TMTYPE = 4) And (TMX < TMLTH) Then Mid$(TMLINE$, TMX + 1, TMLTH) = TMFILL$
     Return

2800 ' Print Information for Free Format Report
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
        If (BD$ <> Space$(12)) Or (PLTH > 1) Then
           LSet CL$ = ", " + RFLBL$(12): X = 1
           If (BD$ <> Space$(12)) Then
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
        If (M1Y <> 0) Or (M1MD <> 0) Or (PLTH > 1) Then
           LSet CL$ = ", " + RFLBL$(15): X = 1
           If (M1Y > 0) Or (M1MD > 0) Then
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
        If (DD$ <> Space$(12)) Or (PLTH > 1) Then
           LSet CL$ = ", " + RFLBL$(19): X = 1
           If (DD$ <> Space$(12)) Then
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

     ' Number of Children
     If PNUMChild And (NUMCH > 0) Then
        LSet CL$ = ";" + Str$(NUMCH)
        Z = InStr(CL$, "  ") + 1
        If NUMCH = 1 Then
           Mid$(CL$, Z, 255) = RFLBL$(26)
           Else
           Mid$(CL$, Z, 255) = RFLBL$(29)
        End If
        Call PutDATA
     End If
     Return

End Sub

Private Sub INDEXReport()
' SHARED DFNUM(), DFPP(), SFLD(), ShowSTATS, FMT001, SST1$

410 ' Print INDEX Detail Report
    cDFM$ = CONFIG.DFM
    If FF1Hdr.MAXID < 1 Then
       Call PutMSG("INDXM001")
       GoTo 420
    End If
    ReDim NDXID(FF1Hdr.MAXID)
    Call INDEXFile("READ"): If Not OKAY Then GoTo 420
    Call SX2SFLD
    Call RptOPEN("INDX"): If Not OKAY Then GoTo 420
    Call FAMOpen("READ"): If Not OKAY Then GoTo 420
    Call PrintOPEN: If Not OKAY Then GoTo 420
    Call BldINDXDEF: If Not OKAY Then GoTo 420
    Call BldINDXTitle
    BD$ = Space$(11): DD$ = BD$
    BPL$ = Space$(41): DPL$ = BPL$: AGE$ = ""
    X$ = Space$(52): Y$ = Space$(57): SNDX$ = Space$(4)
    ts = ENV.SortOPTS And 3: DS = ((ENV.SortOPTS \ 4) And 7)
    UseSEQTable = ((ENV.SortOPTS And 32) > 0)
    NMSORT = ((SFLD(1) \ 256) > 0) Or ((SFLD(2) \ 256) > 0)
    If NMSORT Then
       UPCASE = (ts = 1)
       LASTFIRST = ((SX(1) \ 256) = 1)
       If LASTFIRST Then
          RptOPTION.NAMEFMT = 1 - 2 * UPCASE
          Else
          RptOPTION.NAMEFMT = -2 * UPCASE
       End If
    End If
    BPSORT = (SFLD(6) \ 256) > 0: DPSORT = (SFLD(7) \ 256) > 0
    Select Case DS    ' Determine Date Format for Display & Date Sort Key Lth
       Case 0: PDFM$ = "3"
       Case 1: PDFM$ = "1"
       Case 2: PDFM$ = "2"
       Case 3: PDFM$ = "3"
       Case 4: PDFM$ = "1"
       Case 5: PDFM$ = "3"
       Case 6: PDFM$ = "1"
       Case 7: PDFM$ = "2"
    End Select
    CONFIG.DFM = PDFM$
    PSK = SX(1) \ 256      ' Get Primary Sort Key Field & Set Sort Key Lth
    If (2 < PSK) And (PSK < 6) Then
       Select Case DS    ' Determine Date Sort Key Lth
          Case 0: DSKL = 11: DSK1 = 4: DSK2 = 7
          Case 1: DSKL = 11: DSK1 = 2: DSK2 = 5
          Case 2: DSKL = 11: DSK1 = 2: DSK2 = 5
          Case 3: DSKL = 7: DSK1 = 4
          Case 4: DSKL = 5: DSK1 = 2
          Case 5: DSKL = 4
          Case 6: DSKL = 2
          Case 7: DSKL = 2
       End Select
    End If
    Select Case PSK
       Case 1: If ts > 1 Then PSKL = 4 Else PSKL = 57
       Case 2: If ts > 1 Then PSKL = 4 Else PSKL = 57
       Case 3: PSKL = DSKL
       Case 4: PSKL = DSKL
       Case 5: PSKL = DSKL
       Case 6: If ts > 1 Then PSKL = 4 Else PSKL = 22
       Case 7: If ts > 1 Then PSKL = 4 Else PSKL = 22
       Case 8: PSKL = 1
       Case 9: PSKL = 5
    End Select
    LSK$ = Space$(PSKL): CSK$ = LSK$

415 Z = 0
    While (Z < FF1Hdr.MAXID) And OKAY
       Z = Z + 1: RID = NDXID(Z)
       If RID > 0 Then Call FF1GetRec(RID) Else OKAY = 0
       If OKAY Then
          Call GetBDEvents(BD$, BPL$, DD$, DPL$, AGE$)
          MARRIED = (FF1Rec.SPOUSE <> 0)
          If MARRIED Then Call FF3GetRec(FF1Rec.SPOUSE)
          FX = 0
          While FX < PAGE.DFCNT
             FX = FX + 1
             DFN = DFNUM(FX)
             PP = DFPP(FX): FLTH = RFLTH(DFN)
             Select Case DFN
                Case 1: LSet X$ = SFILL$(RID, FLTH)
                Case 2: LSet X$ = FMTNAME$(0): If NMSORT Then GoSub 416
                Case 3: LSet X$ = FF1Rec.SEX
                Case 4: LSet X$ = BD$
                Case 5: If MARRIED Then LSet X$ = XSCRNDate$(FF3SPOUSE.D1MD, FF3SPOUSE.D1Y)
                Case 6: LSet X$ = DD$
                Case 7: LSet X$ = BPL$: If BPSORT Then GoSub 416
                Case 8: LSet X$ = DPL$: If DPSORT Then GoSub 416
             End Select
             If FX = 1 Then
                LSet CSK$ = X$: ML = PSKL
                MATCHED = (MID$(LSK$, 1, ML) = MID$(CSK$, 1, ML))
                If Not MATCHED Then
                   If (DFN = 2) And (ts < 2) Then
                      DSK1 = 0
                      If LASTFIRST Then
                         DSK2 = Len(RTrim$(SURNM$))
                         Else
                         DSK2 = Len(RTrim$(FF1Rec.GIVEN))
                      End If
                   End If
                   If DSK2 > 0 Then
                      ML = DSK2
                      MATCHED = (MID$(LSK$, 1, ML) = MID$(CSK$, 1, ML))
                   End If
                End If
                If Not MATCHED Then
                   If DSK1 > 0 Then
                      ML = DSK1
                      MATCHED = (MID$(LSK$, 1, ML) = MID$(CSK$, 1, ML))
                   End If
                End If
                If MATCHED Then
                   Mid$(X$, 1, ML) = Space$(ML)
                End If
                LSK$ = CSK$
             End If
             Mid$(PLINE$, PP, FLTH) = X$: LSet X$ = ""
          Wend
          Call PrintLINE
          If a = 27 Then Z = FF1Hdr.MAXID + 1
       End If
    Wend
    Call RptCLOSE: If PRT.AGAIN Then GoTo 415
    GoTo 420

416 ' Build Text Field
    If ts > 1 Then
       Call ComputeSNDX(X$, SNDX$): LSet Y$ = SNDX$ + " " + X$: LSet X$ = Y$
    End If
    Return

420 ' Return to Menu
    CONFIG.DFM = cDFM$
    ReDim NDXID(1)
    Close: ENV.FFOPEN = 0

End Sub


Private Sub BldINDXDEF()
      'SHARED DFNUM(), DFPP(), SFLD()
      PAGE.DFCNT = 0: RFCNT = 8
      ReDim DFNUM(10), DFPP(10), DFCHK(10), DFCHK2(10)
      ReDim RFNUM(RFCNT), RFSEP(RFCNT), RFLTH(RFCNT), RFLBL$(RFCNT)
      Call RptOptGET("MNGP", MINGAP)
      If (Not OKAY) Or (MINGAP = 0) Then MINGAP = 2
      Call RptOptGET("MXGP", MAXGAP)
      If (Not OKAY) Or (MAXGAP = 0) Then MAXGAP = 5
      Call RptOptGET("PNML", PLNMLTH)
      If (Not OKAY) Or (PLNMLTH = 0) Then PLNMLTH = 22 - 8 * ENV.LPNSupport

58500 ' Select Report Fields - Based Upon Report Options
      OKAY = 0
      If (PAGE.BWIDTH = 0) Then
         BWIDTH = PRT.FWIDTH - PAGE.IMARGIN - PAGE.OMARGIN
         Else
         BWIDTH = PAGE.BWIDTH
      End If
      DFCHK(1) = 2: DFCHK(2) = 2: DFCHK(3) = 4: DFCHK(4) = 5
      DFCHK(5) = 6: DFCHK(6) = 7: DFCHK(7) = 8: DFCHK(8) = 3
      DFCHK(9) = 1
      X = SFCNT: Y = 0: Z = 0
      While X > 0
         Y = Y + 1: B = SX(Y) \ 256
         C = DFCHK(B)
         If DFCHK2(C) = 0 Then
            DFCHK2(C) = B
            Z = Z + 1: DFNUM(Z) = C
         End If
         X = X - 1
      Wend
      For X = 1 To 6
         If DFCHK2(X) = 0 Then
            Z = Z + 1: DFNUM(Z) = X
         End If
      Next X
      If SX(1) \ 256 = 1 Then
         RptOPTION.NAMEFMT = 1 - 2 * (RptOPTION.NAMEFMT > 1)
         Else
         If SX(1) \ 256 = 2 Then
            RptOPTION.NAMEFMT = -2 * (RptOPTION.NAMEFMT > 1)
         End If
      End If
      PAGE.DFCNT = Z: PAGE.PFCNT = 0
      ts = (ENV.SortOPTS And 3): DS = ((ENV.SortOPTS \ 4) And 7)
      NMSORT = ((SFLD(1) \ 256) > 0) Or ((SFLD(2) \ 256) > 0)
      BPSORT = ((SFLD(6) \ 256) > 0): DPSORT = ((SFLD(7) \ 256) > 0)

58510 ' Create Report Definition Tables
      X = 1: RFNAME$ = "01010": GoSub 58515     'RID
      X = 2: RFNAME$ = "01020": GoSub 58515     'Full Name
      X = 3: RFNAME$ = "01015": GoSub 58515     'Gender
      X = 4: RFNAME$ = "01030": GoSub 58515     'Birth Date
      X = 5: RFNAME$ = "02020": GoSub 58515     'Marriage Date
      X = 6: RFNAME$ = "01040": GoSub 58515     'Death Date
      X = 7: RFNAME$ = "01032": RFLTH(X) = PLNMLTH
             GoSub 58515                        'Birth Place
      X = 8: RFNAME$ = "01042": RFLTH(X) = PLNMLTH
             GoSub 58515     'Death Place
      GoTo 58525

58515 ' Process Report Field
      Call GetRFLD(RFNAME$, FX)
      If RFSEP(X) = 0 Then RFSEP(X) = MINGAP
      If FOUND Then
         RDFFld = RFLDS(FX)
         If RFLTH(X) = 0 Then L = Asc(RDFFld.RLTH) Else L = RFLTH(X)
         ALTH = Asc(RDFFld.ALTH): LLTH = Asc(RDFFld.LLTH)
         HLTH = Asc(RDFFld.HLTH)
         RFLBL$(X) = MID$(RFLD$(FX), ALTH + LLTH + 1, HLTH)
         If (X = 2) Or (X = 7) Or (X = 8) Then
            If ts > 1 Then L = 5 + L
         End If
         RFLTH(X) = L
      End If
      Return

58525 ' Check Report Width
      GoSub 58530
      If RWIDTH <= BWIDTH Then GoTo 59540
      GoTo 58535

58530 ' Determine Required Width
      RWIDTH = 0: FX = 0
      While FX < PAGE.DFCNT
          FX = FX + 1: X = DFNUM(FX): FSEP = RFSEP(X)
          If (X > 0) And (FSEP > 0) Then
             Y = RFLTH(X): Z = Len(RFLBL$(X)): If Y < Z Then Y = Z
             RWIDTH = RWIDTH + FSEP + Y
          End If
      Wend
      Return

58535 ' Check Report Width
      RPWIDTH = RWIDTH + PAGE.IMARGIN + PAGE.OMARGIN
      PWIDTH = BWIDTH + PAGE.IMARGIN + PAGE.OMARGIN
      If PWIDTH < RPWIDTH Then
         SysVAR.NVAR1 = PWIDTH: SysVAR.NVAR2 = RPWIDTH
         Call PutMSG("XXXXM009")
         If CA$ <> SNGLKEY$(1) Then OKAY = 0: GoTo 58590
         BWIDTH = RWIDTH
      End If

59540 ' Adjust to Fit
      If Not (RWIDTH < BWIDTH) Then GoTo 59550
      GAPS = 0: FX = PAGE.PFCNT
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
      FX = PAGE.PFCNT
      While FX <= PAGE.DFCNT
          X = DFNUM(FX): FSEP = RFSEP(X)
          If (X > 0) And (FSEP > 0) Then
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
             Z = FL: If Z < HL Then Z = HL
             PP = PP + FSEP + (Z + 1) \ 2 - (FL + 1) \ 2
             DFPP(FX) = PP
             PP = PP + Z
          End If
      Wend

58560 ' Return to Calling Program
      PAGE.BWIDTH = BWIDTH
      PLINE$ = Space$(PAGE.IMARGIN + PAGE.BWIDTH + PAGE.OMARGIN)
      OKAY = -1
      
58590 ' Skip OKAY
End Sub

Private Sub BldINDXTitle()
     'SHARED DFNUM(), DFPP()
     ' Build Title Line
      TL$ = Space$(PAGE.BWIDTH)
      For FX = 1 To PAGE.DFCNT: X = DFNUM(FX): PP = DFPP(FX)
         If (X > 0) And (PP > 0) Then
            FL = RFLTH(X): HL = Len(RFLBL$(X))
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


Sub GetLastRES()
    ' Get Residence
    D1Y = 0: D1MD = 0: ADRSPTR = FF1Rec.ADRS: ADRSSPID = 0
    If ADRSPTR <> 0 Then
       Call FF2GetRec(ADRSPTR)
       D1Y = FF2PFX.D1Y: D1MD = FF2PFX.D1MD
    End If
    MPTR = FF1Rec.SPOUSE: MADRS = 0
    While (MPTR <> 0) And (MADRS = 0)
       Call FF3GetRec(MPTR): MADRS = FF3SPOUSE.LOC
       If FF1Rec.RID = FF3SPOUSE.SP1ID Then
          SPID = FF3SPOUSE.SP2ID: MPTR = FF3SPOUSE.SP1NXT
          Else
          SPID = FF3SPOUSE.SP1ID: MPTR = FF3SPOUSE.SP2NXT
       End If
       If MADRS <> 0 Then
          Call FF2GetRec(MADRS)
          If ADRSPTR <> 0 Then
             D2Y = FF2PFX.D1Y: D2MD = FF2PFX.D1MD
             If ((D1Y < D2Y) Or ((D1Y = D2Y) And (D1MD <= D2MD))) Then
                ADRSPTR = MADRS: ADRSSPID = SPID
                Else
                Call FF2GetRec(ADRSPTR)
             End If
             Else
             ADRSPTR = MADRS: ADRSSPID = SPID
          End If
          Else
       End If
    Wend
End Sub

Sub GetNextMARRIAGE()
   If SPRESEQ Then
      If SPCNT > 0 Then SPCNT = SPCNT - 1
      SPRNO = MLIST(SPCNT)
      Else
      If INDID = FF3SPOUSE.SP1ID Then
         SPRNO = FF3SPOUSE.SP1NXT
         Else
         If INDID = FF3SPOUSE.SP2ID Then
            SPRNO = FF3SPOUSE.SP2NXT
            Else
            SPRNO = 0
         End If
      End If
   End If

End Sub

Sub XSUMReport()
     
     HDFM$ = CONFIG.DFM

     Select Case RptCODE
        Case "ASUM": WTYPE = 1
        Case "DSUM": WTYPE = 2
        Case "RSUM": WTYPE = 3
        Case "ISUM": WTYPE = 4
        Case Else: WTYPE = 4
     End Select

     ' Produce Summary Report
     RELSUMOPT = (WTYPE = 3): GENSUMOPT = (WTYPE < 3)
     INDEXOPT = (WTYPE = 4)
     If GENSUMOPT Then
        Call PutMSG("XXXXM044")
        If a = 27 Then GoTo S1040
        VRTCL = (CA$ = "2")
        Else
        VRTCL = 0
     End If
     SRPT$ = RptCODE
     Call RptOPEN(SRPT$): If Not OKAY Then GoTo S1040
     WOPEN$ = "READQ"
S1010: ' Verify Files by Opening
     Call FAMOpen("READ"): If Not OKAY Then GoTo S1040
     If INDEXOPT Then
        Call INDEXFile("READ")
        If (a = 27) Or (Not OKAY) Then GoTo S1040
        PSK = SX(1) \ 256
        If (PSK) > 8 Then
           Call PutMSG("LISTM025")
           Call FamCLOSE: GoTo S1010
        End If
        If (2 < PSK) And (PSK < 6) Then
           DS = (ENV.SortOPTS \ 4) And 7
           Else
           If PSK <> 8 Then
              ts = ENV.SortOPTS And 3
              If ts < 2 Then Call RptOptGET("SNDX", X): PrntSNDX = (X = 1)
           End If
        End If
        xX$ = "SFLD" + LTrim$(Str$(PSK))
        Call GetRVAR(xX$, X): If FOUND Then SysVAR.SUBJECT = RVAR$(X)
        GoTo S1050
        Else
        Call RWRKOpen(WOPEN$, WTYPE, 0): If OKAY Then GoTo S1050
        If WOPEN$ = "READQ" Then
           Call FamCLOSE
           LSet WOPEN$ = "READ": GoTo S1010
        End If
     End If

S1040: ' Bad Return
     CONFIG.DFM = HDFM$
     Close: ENV.FFOPEN = 0: Exit Sub

S1050: ' Prepare to Print
     If INDEXOPT Then
        RptOPTION.NAMEFMT = (RptOPTION.NAMEFMT And 1) - 2 * (ts = 2)
        Else
        Call RWRKGet("BASE", RC, PTR&): If RC <> 1 Then GoTo S1040
        SysVAR.SUBJECT = FMTNAME$(1)
        If DMAXLV = 0 Then
           Call RptOptGET("CNOP", X): COUNTNOP = (X = 1)
           If Not VRTCL Then Call RptOptGET("LNOP", X): LISTNODES = (X = 1)
           LALIN = Len(Str$(2 ^ Asc(WFHDR.AMAXLV))): zALIN$ = Space$(LALIN)
        End If
        Call RptOptGET("CDUP", X): COUNTDUPS = (X = 1)
        Call RptOptGET("CLIN", X): COUNTCLINE = (X = 1)
     End If
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
        If GENSUMOPT Then
           HIRGL = Asc(WFHDR.AMAXLV) + Asc(WFHDR.DMAXLV)
           If VRTCL Then ReDim SUMTAB(HIRGL) As SUMREC
        End If
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
     If INDEXOPT Then
        If (2 < PSK) And (PSK < 6) Then
           Select Case DS
              Case 0: cDFM$ = "3": DSKL = 11
              Case 1: cDFM$ = "1": DSKL = 11
              Case 2: cDFM$ = "2": DSKL = 11
              Case 3: cDFM$ = "3": DSKL = 7
              Case 4: cDFM$ = "1": DSKL = 5
              Case 5: cDFM$ = "3": DSKL = 4
              Case 6: cDFM$ = "1": DSKL = 2
              Case 7: cDFM$ = "2": DSKL = 2
           End Select
           Select Case cDFM$
              Case "1": MMX = 1: DDX = 4: YYX = 7
              Case "2": DDX = 1: MMX = 4: YYX = 7
              Case "3": YYX = 1: MMX = 6: DDX = 9
           End Select
           CONFIG.DFM = cDFM$
        End If
        Select Case PSK
           Case 1: PSKL = SX(1) And 127
           Case 2: PSKL = SX(1) And 127
           Case 3: PSKL = DSKL: CONFIG.DFM = cDFM$
           Case 4: PSKL = DSKL: CONFIG.DFM = cDFM$
           Case 5: PSKL = DSKL: CONFIG.DFM = cDFM$
           Case 6: PSKL = SX(1) And 127: BPL$ = Space$(PSKL)
           Case 7: PSKL = SX(1) And 127: DPL$ = Space$(PSKL)
           Case 8: PSKL = 1
        End Select
     End If

S1200: ' Setup Report Tables & Open Files
     Call BldSumDEF(WTYPE): If Not OKAY Then GoTo S1040
     If Not RELSUMOPT Then
        Call PrintOPEN: If Not OKAY Then GoTo S1040
        PLINE$ = Space$(PAGE.BWIDTH + PAGE.IMARGIN + PAGE.OMARGIN)
        If INDEXOPT Then PRT.SECTION = "S"
        Call BldSumTITLE
        If GENSUMOPT Then
        ' IF Ancestor or Descendant Summary, Retrieve Records by Generation
           WFHDR.WFRULES = Chr$((Asc(WFHDR.WFRULES) And 252) + 1)
        End If
     End If

S1400: ' Begin with Base Record
     SUMBASE.LOBY = 10000: SUMBASE.LOMY = 10000: SUMBASE.LODY = 10000
     sum = SUMBASE: TOT = SUMBASE
     L = 20: LAGL = 0: LRGL = 0
     a = 0: LGL = -1: FINALTOT = 0
     xRGL$ = Space$(RFLTH(1))
     If RELSUMOPT Then Call PutMSG("XXXXM064")
     If INDEXOPT Then
        LSK$ = Space$(PSKL): CSK$ = LSK$
        LSNDX$ = Space$(4): CSNDX$ = LSNDX$
        xNDX = 1: RC = 1: PRT.SECTION = "S"
        Else
        Call RWRKGet("BASE", RC, PTR&): If RC <> 1 Then GoTo S1040
     End If
     FIRSTONE = -1: GoTo S1450

S1415: ' Process next entry on chart
     If INDEXOPT Then
        xNDX = xNDX + 1: If xNDX > NDXHDR.NDXRECS Then RC = 9: GoTo S1455
        Else
        If RELSUMOPT Then
           Call RWRKGet("WRKREC", RC, PTR&)
           Else
           Call RWRKGet("NEXT", RC, PTR&)
        End If
        If (RC < 1) Or (RC > 8) Then GoTo S1455
     End If

S1450: ' Process Workfile record
     If INDEXOPT Then
        RELID = NDXID(xNDX)
        If (SelSET(RELID) And xSEL) = 0 Then GoTo S1415
        Else
        RELID = Abs(WFRELREC.id)
        If (SelSET(RELID) And xSEL) = 0 Then GoTo S1415
        AGL = -WFAGLREC.AGL: RGL = AGL + WFGLREC.DGL
     End If
     If RELSUMOPT Then GoSub S2200: Else GoSub S2400
     If a <> 27 Then GoTo S1415

S1455: ' Process Final Totals
     If RELSUMOPT Then
        GoSub S2210
        Else
        GoSub S2460
        If VRTCL Then GoSub S2700
        If INDEXOPT Then
           Call WriteLINE
           PRT.SECTION = "T": SysVAR.NVAR1 = TOT.GROUPS
           Call PrintTITLE
        End If
        Call RptCLOSE
        If PRT.AGAIN Then
           If VRTCL Then GoTo S1455 Else GoTo S1400
        End If
     End If
     CONFIG.DFM = HDFM$
     Call FamCLOSE
     Exit Sub

S2200: ' Process Record for RELATIVE Table
     If WFRELREC.id > 0 Then
        RELCNT(AGL, RGL) = RELCNT(AGL, RGL) + 1
     End If
     Return

S2210: ' Print Relationship Summary Table
      Call PrintOPEN
      If Not OKAY Then GoTo S1040
      FULLWIDTH = 0: PLINE$ = Space$(PAGE.BWIDTH + PAGE.IMARGIN + PAGE.OMARGIN)
      If (PRT.OPT > 2) And (PAGE.BWIDTH > PRT.FWIDTH - PAGE.IMARGIN - PAGE.OMARGIN) Then
         Call PutMSG("XXXXM070")
         If Asc(CA$) = 27 Then GoTo S1040 Else FULLWIDTH = (CA$ <> "1")
         If FULLWIDTH Then
            PRT.CFORMS = -1: NOSPAN = -1
            PRT.FWIDTH = PAGE.BWIDTH + PAGE.IMARGIN + PAGE.OMARGIN
         End If
      End If
      RELABEL$ = PLINE$: xRELCNT$ = PLINE$
      Call BldSumTITLE

S2215: ' Print Table (AGAIN?)
      ReDim AGLTOT(LORGL To 0)
      RGLTOT = 0: RGLTOTAL = 0
      RGL = LORGL: SREL$ = Space$(RFLTH(2))
      While RGL <= HIRGL: xAGL = 0: FX = 2
         LSet PLINE$ = "": Call WriteLINE: RGLTOT = 0
         RSet xRGL$ = Str$(RGL): Mid$(xRELCNT$, DFPP(1), RFLTH(1)) = xRGL$
         While xAGL >= LORGL
            PP = DFPP(FX): FLTH = RFLTH(DFNUM(FX))
            RCNT = RELCNT(xAGL, RGL)
            If (xAGL > RGL) Or (RCNT = 0) Then
               Mid$(xRELCNT$, PP + FLTH - 1, 1) = "."
               Else
               If ShowREL Then
                  RSet SREL$ = Relation$("S", "*", xAGL, RGL, RRELCD$())
                  Mid$(RELABEL$, PP, FLTH) = SREL$
               End If
               If RCNT > 0 Then Mid$(xRELCNT$, PP, FLTH) = SFILL$(RCNT, FLTH)
               AGLTOT(xAGL) = AGLTOT(xAGL) + RCNT
               RGLTOT = RGLTOT + RCNT
            End If
            FX = FX + 1: xAGL = xAGL - 1
         Wend
         Mid$(xRELCNT$, DFPP(FX), RFLTH(FX)) = SFILL$(RGLTOT, RFLTH(FX))
         If ShowREL Then LSet PLINE$ = RELABEL$: Call PrintLINE: LSet RELABEL$ = ""
         LSet PLINE$ = xRELCNT$: Call PrintLINE: LSet xRELCNT$ = ""
         RGL = RGL + 1: RGLTOTAL = RGLTOTAL + RGLTOT
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

S2400: ' Process Records for Indexed & ANC/DESC Genlevel Summaries
     Call FF1GetRec(RELID)
     If (PSK = 1) Then
        If ENV.HUSBSURNM Then xname$ = FMTNAME$(0) Else SURNM$ = FF1Rec.SURNM
     End If
     Call GetBDEvents("", BPL$, "", DPL$, "")
     If INDEXOPT Then
        Select Case PSK
           Case 1: LSet CSK$ = SURNM$
           Case 2: LSet CSK$ = FF1Rec.GIVEN
           Case 3: LSet CSK$ = XSCRNDate$(FF1Rec.BMD, FF1Rec.BY)
           Case 4: MRID = FF1Rec.SPOUSE: LSet CSK$ = ""
                   If MRID <> 0 Then
                      Call FF3GetRec(MRID)
                      If OKAY Then
                         LSet CSK$ = XSCRNDate$(FF3SPOUSE.D1MD, FF3SPOUSE.D1Y)
                      End If
                   End If
           Case 5: LSet CSK$ = XSCRNDate$(FF1Rec.DMD, FF1Rec.DY)
           Case 6: LSet CSK$ = BPL$
           Case 7: LSet CSK$ = DPL$
           Case 8: LSet CSK$ = FF1Rec.SEX
        End Select
        If ts = 1 Then LSet CSK$ = UCX$(CSK$)
        If PrntSNDX Or (ts > 1) Then Call ComputeSNDX(CSK$, CSNDX$)
     End If
     If (sum.TOTAL = 0) Then
        If INDEXOPT Then LSet LSK$ = CSK$: LSet LSNDX$ = CSNDX$
        GoTo S2420
     End If
     If INDEXOPT Then
        MATCHED = -1
        If ts > 1 Then MATCHED = (CSNDX$ = LSNDX$)
        If MATCHED And (ts <> 2) Then
           MATCHED = (MID$(CSK$, 1, PSKL) = MID$(LSK$, 1, PSKL))
        End If
        Else
        MATCHED = (RGL = LRGL)
     End If
     If MATCHED Then GoTo S2420
     If Not VRTCL Then GoSub S2550
     GoSub S2570
     FIRSTONE = -1: LAGL = AGL: LRGL = RGL
     LSet LSK$ = CSK$: LSet LSNDX$ = CSNDX$
     sum = SUMBASE

S2420: ' Increment Summary Counters
     If RELSUMOPT Then
        C = -(WFRELREC.id > 0)
        Else
        C = 1
     End If
     If (COUNTPLACES And C) Then
        If RTrim$(BPLACE) <> "" Then sum.BPL = sum.BPL + 1
        If RTrim$(DPLACE) <> "" Then sum.DPL = sum.DPL + 1
     End If
     If ((FF1Rec.BY > 0) And C) Then
        sum.BDT = sum.BDT + 1
        If (FF1Rec.BY < sum.LOBY) Then sum.LOBY = FF1Rec.BY
        If (FF1Rec.BY > sum.HIBY) Then sum.HIBY = FF1Rec.BY
     End If
     If ((FF1Rec.DY > 0) And C) Then
        sum.DDT = sum.DDT + 1
        If FF1Rec.DY < 9999 Then
           sum.DDTV = sum.DDTV + 1
           If (FF1Rec.DY < sum.LODY) Then sum.LODY = FF1Rec.DY
           If (FF1Rec.DY > sum.HIDY) Then sum.HIDY = FF1Rec.DY
        End If
     End If
     If (COUNTMARRY And C) Then
        SPREC = FF1Rec.SPOUSE
        If SPREC <> 0 Then sum.MARRIED = sum.MARRIED + 1
        While (SPREC <> 0)
           sum.MARCNT = sum.MARCNT + 1
           Call FF3GetRec(SPREC)
           If (FF3SPOUSE.D1Y > 0) Then
              sum.MDT = sum.MDT + 1
              If (FF3SPOUSE.D1Y < sum.LOMY) Then sum.LOMY = FF3SPOUSE.D1Y
              If (FF3SPOUSE.D1Y > sum.HIMY) Then sum.HIMY = FF3SPOUSE.D1Y
           End If
           If (FF3SPOUSE.MLOC <> 0) Then sum.MPL = sum.MPL + 1
           If (FF3SPOUSE.SP1ID = RELID) Then
              SPREC = FF3SPOUSE.SP1NXT
              Else
              SPREC = FF3SPOUSE.SP2NXT
           End If
        Wend
     End If
     If ((COUNTNOP Or LISTNODES) And C) Then
        If (FF1Rec.FID = 0) And (FF1Rec.MID = 0) Then
           sum.NOP = sum.NOP + 1
           If LISTNODES Then GoSub S2500
        End If
     End If
     If COUNTCLINE Then
        If RC = 1 Then sum.CLINE = sum.CLINE + WFAGLREC.CLINES
        If RC = 3 Then sum.CLINE = sum.CLINE + WFGLREC.CLINES
     End If
     sum.TOTAL = sum.TOTAL + C
     If FF1Rec.SEX = FGENDR$(1) Then
        sum.MALE = sum.MALE + C
        Else
        If FF1Rec.SEX = FGENDR$(2) Then sum.FEMALE = sum.FEMALE + C
     End If
     Return

S2460: ' Process last summary Line & Totals
     If Not VRTCL Then GoSub S2550
     If FIRSTONE Then GoSub S2570
     FIRSTONE = 0
     If Not VRTCL Then
        Call WriteLINE
        Mid$(PLINE$, DFPP(1), 5) = TOTAL$
        sum = TOT: FINALTOT = -1
        GoSub S2550
     End If
     Return

S2500: ' Print Nodal Person
     If FIRSTONE Then
        RSet xRGL$ = Str$(LRGL): Mid$(PLINE$, DFPP(1), RFLTH(1)) = xRGL$
        If RptOPTION.Relation Then Mid$(PLINE$, DFPP(2), RFLTH(2)) = Relation$("M", "*", LAGL, LRGL, RRELCD$())
        Call WriteLINE: FIRSTONE = 0
     End If
     FX = 2 - RptOPTION.Relation: PP = DFPP(FX)
     RSet zALIN$ = Str$(WFANCREC.LINEAGE)
     Mid$(PLINE$, PP, LALIN) = zALIN$: PP = PP + LALIN + 2
     If FF1Rec.BY > 0 Then Mid$(PLINE$, PP, 4) = ZFILL$(FF1Rec.BY, 4)
     Mid$(PLINE$, PP + 4, 1) = "-"
     If FF1Rec.DY > 0 Then Mid$(PLINE$, PP + 5, 4) = ZFILL$(FF1Rec.DY, 4)
     Mid$(PLINE$, PP + 11, 52) = FMTNAME$(0)
     Call WriteLINE
     Return

S2550: ' Print Ancestor/Descendant Summary Line - Horizontal Style
     FX = 1
     While FX <= PAGE.DFCNT
        X = DFNUM(FX)
        If (Not FINALTOT) Or (X < 24) Then
           PP = DFPP(FX): FLTH = RFLTH(X)
           Select Case X
              Case 1: If FIRSTONE Then RSet xRGL$ = Str$(LRGL): Mid$(PLINE$, PP, FLTH) = xRGL$
              Case 2: If FIRSTONE Then Mid$(PLINE$, PP, FLTH) = Relation$("M", "*", LAGL, LRGL, RRELCD$())
              Case 3: Mid$(PLINE$, PP, FLTH) = SFILL$(sum.TOTAL, FLTH)
                      If sum.TOTAL > (sum.MALE + sum.FEMALE) Then Mid$(PLINE$, PP + FLTH, 1) = "*"
              Case 4: If Not FINALTOT Then Mid$(PLINE$, PP, FLTH) = SFILL$(sum.CLINE, FLTH)
              Case 6: Mid$(PLINE$, PP, FLTH) = SFILL$(sum.NOP, FLTH)
              Case 7: Mid$(PLINE$, PP, FLTH) = SFILL$(sum.MALE, FLTH)
              Case 8: Mid$(PLINE$, PP, FLTH) = SFILL$(sum.FEMALE, FLTH)
              Case 9: Mid$(PLINE$, PP, FLTH) = SFILL$(sum.BDT, FLTH)
              Case 10: Mid$(PLINE$, PP, FLTH) = SFILL$(sum.BPL, FLTH)
              Case 11: If sum.BDT > 0 Then Mid$(PLINE$, PP, FLTH) = ZFILL$(sum.LOBY, FLTH)
              Case 12: If sum.BDT > 0 Then Mid$(PLINE$, PP, FLTH) = ZFILL$(sum.HIBY, FLTH)
              Case 13: Mid$(PLINE$, PP, FLTH) = SFILL$(sum.MARRIED, FLTH)
              Case 14: Mid$(PLINE$, PP, FLTH) = SFILL$(sum.MARCNT, FLTH)
              Case 15: Mid$(PLINE$, PP, FLTH) = SFILL$(sum.MDT, FLTH)
              Case 16: Mid$(PLINE$, PP, FLTH) = SFILL$(sum.MPL, FLTH)
              Case 17: If sum.MDT > 0 Then Mid$(PLINE$, PP, FLTH) = ZFILL$(sum.LOMY, FLTH)
              Case 18: If sum.MDT > 0 Then Mid$(PLINE$, PP, FLTH) = ZFILL$(sum.HIMY, FLTH)
              Case 19: Mid$(PLINE$, PP, FLTH) = SFILL$(sum.DDT, FLTH)
              Case 20: Mid$(PLINE$, PP, FLTH) = SFILL$(sum.DDTV, FLTH)
              Case 21: Mid$(PLINE$, PP, FLTH) = SFILL$(sum.DPL, FLTH)
              Case 22: If sum.DDTV > 0 Then Mid$(PLINE$, PP, FLTH) = ZFILL$(sum.LODY, FLTH)
              Case 23: If sum.DDTV > 0 Then Mid$(PLINE$, PP, FLTH) = ZFILL$(sum.HIDY, FLTH)
              Case 24: Mid$(PLINE$, PP, FLTH) = LSNDX$
              Case 25:
              Case 26: Mid$(PLINE$, PP, 4) = MID$(LSK$, YYX, 4)
              Case 27: mm = Val(MID$(LSK$, MMX, 2))
                       If (0 < mm) And (mm < 13) Then
                          Mid$(PLINE$, PP, 3) = MID$(RMONTB$(mm + 1), 3, 3)
                       End If
              Case 28: DD = Val(MID$(LSK$, DDX, 2))
                       If DD > 0 Then Mid$(PLINE$, PP, 2) = MID$(LSK$, DDX, 2)
              Case Else: Mid$(PLINE$, PP, FLTH) = LSK$
           End Select
        End If
        FX = FX + 1
     Wend
     Call PrintLINE
     Return

S2570: ' Update Overall Counters
     TOT.GROUPS = TOT.GROUPS + 1
     TOT.TOTAL = TOT.TOTAL + sum.TOTAL
     TOT.NOP = TOT.NOP + sum.NOP
     TOT.CLINE = TOT.CLINE + sum.CLINE
     TOT.MALE = TOT.MALE + sum.MALE
     TOT.FEMALE = TOT.FEMALE + sum.FEMALE
     TOT.MARRIED = TOT.MARRIED + sum.MARRIED
     TOT.MARCNT = TOT.MARCNT + sum.MARCNT
     TOT.BPL = TOT.BPL + sum.BPL: TOT.MPL = TOT.MPL + sum.MPL
     TOT.DPL = TOT.DPL + sum.DPL
     TOT.BDT = TOT.BDT + sum.BDT: TOT.DDT = TOT.DDT + sum.DDT
     TOT.DDTV = TOT.DDTV + sum.DDTV: TOT.MDT = TOT.MDT + sum.MDT
     If (sum.LOBY > 0) And (sum.LOBY < TOT.LOBY) Then TOT.LOBY = sum.LOBY
     If sum.HIBY > TOT.HIBY Then TOT.HIBY = sum.HIBY
     If (sum.LODY > 0) And (sum.LODY < TOT.LODY) Then TOT.LODY = sum.LODY
     If sum.HIDY > TOT.HIDY Then TOT.HIDY = sum.HIDY
     If (sum.LOMY > 0) And (sum.LOMY < TOT.LOMY) Then TOT.LOMY = sum.LOMY
     If sum.HIMY > TOT.HIMY Then TOT.HIMY = sum.HIMY
     If VRTCL Then SUMTAB(Abs(LRGL)) = sum
     Return

S2700: ' Print Ancestor/Descendant Summary Line - Vertical Style
     FX = 3
     While FX <= PAGE.DFCNT
        X = DFNUM(FX): Mid$(PLINE$, 1, 255) = RFLBL$(X)
        FLTH = RFLTH(X)
        PP = RFLTH(0) + DFPP(0)
        For Y = 0 To HIRGL + 1
            PP = PP + 6 - FLTH
            If (Y <= HIRGL) Then sum = SUMTAB(Y) Else sum = TOT
            Select Case X
               Case 3:  Mid$(PLINE$, PP, FLTH) = SFILL$(sum.TOTAL, FLTH)
                        If sum.TOTAL > (sum.MALE + sum.FEMALE) Then Mid$(PLINE$, PP + FLTH, 1) = "*"
               Case 4:  Mid$(PLINE$, PP, FLTH) = SFILL$(sum.CLINE, FLTH)
               Case 6:  Mid$(PLINE$, PP, FLTH) = SFILL$(sum.NOP, FLTH)
               Case 7:  Mid$(PLINE$, PP, FLTH) = SFILL$(sum.MALE, FLTH)
               Case 8:  Mid$(PLINE$, PP, FLTH) = SFILL$(sum.FEMALE, FLTH)
               Case 9:  Mid$(PLINE$, PP, FLTH) = SFILL$(sum.BDT, FLTH)
               Case 10: Mid$(PLINE$, PP, FLTH) = SFILL$(sum.BPL, FLTH)
               Case 11: If sum.BDT > 0 Then Mid$(PLINE$, PP, FLTH) = ZFILL$(sum.LOBY, FLTH)
               Case 12: If sum.BDT > 0 Then Mid$(PLINE$, PP, FLTH) = ZFILL$(sum.HIBY, FLTH)
               Case 13: Mid$(PLINE$, PP, FLTH) = SFILL$(sum.MARRIED, FLTH)
               Case 14: Mid$(PLINE$, PP, FLTH) = SFILL$(sum.MARCNT, FLTH)
               Case 15: Mid$(PLINE$, PP, FLTH) = SFILL$(sum.MDT, FLTH)
               Case 16: Mid$(PLINE$, PP, FLTH) = SFILL$(sum.MPL, FLTH)
               Case 17: If sum.MDT > 0 Then Mid$(PLINE$, PP, FLTH) = ZFILL$(sum.LOMY, FLTH)
               Case 18: If sum.MDT > 0 Then Mid$(PLINE$, PP, FLTH) = ZFILL$(sum.HIMY, FLTH)
               Case 19: Mid$(PLINE$, PP, FLTH) = SFILL$(sum.DDT, FLTH)
               Case 20: Mid$(PLINE$, PP, FLTH) = SFILL$(sum.DDTV, FLTH)
               Case 21: Mid$(PLINE$, PP, FLTH) = SFILL$(sum.DPL, FLTH)
               Case 22: If sum.DDTV > 0 Then Mid$(PLINE$, PP, FLTH) = ZFILL$(sum.LODY, FLTH)
               Case 23: If sum.DDTV > 0 Then Mid$(PLINE$, PP, FLTH) = ZFILL$(sum.HIDY, FLTH)
            End Select
            PP = PP + FLTH + DFPP(0)
        Next Y
        Call PrintLINE
        FX = FX + 1
     Wend
     Return

End Sub

