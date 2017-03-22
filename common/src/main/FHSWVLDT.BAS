Attribute VB_Name = "modFHSWVLDT"
DefInt A-Z
Dim DFNUM(), DFPP(), RFABR$(), MFCNT, SLINE$
Dim xRID$, UPDTFile, NOPAUSE
Dim LastRID, ALLREC
Dim ERRCNT(), ScrnERRCNT()
Dim CRID

Sub UtilVLDT(OPT$)
1  Rem $INCLUDE: 'FHSCOMON.BAS'
   PN$ = " FAMILY HISTORY SYSTEM - FHSWVLDT - File Validation Program"
   CY$ = " (C) Copyright 1985-1998 by Phillip E. Brown"
   ReDim ERID(1), DFNUM(1), DFPP(1), RFABR$(1)
   READ$ = "READQ"
   WorkSAVED = -1
   
   Select Case OPT$
      Case "INIT": GoTo 40
      Case "PROCESS": GoTo 1000
   End Select
   
40 ' Format Display
   Call PgmPREP("FHSVLDT")
   Call RptLOAD(RptCODE)
   Load frmFHSWUPgm
   frmFHSWUPgm.UPSetup.AddItem RptCODE
   frmFHSWUPgm.UPSetup_Click
   frmFHSWUPgm.Show
   Exit Sub

1000: 'PROCESS Validation Request
     'If Not FAMOKAY Then Call PutMSG("VLDTM001"): GoTo 200
     If FF1Hdr.MAXID = 0 Then Call PutMSG("VLDTM015"): Exit Sub
     Select Case RptCODE
        Case "PVLD": Call PtrCHECK
        Case "DVLD": Call DataCHECK
        Case "MDAT": Call MatchCHECK
     End Select
     WorkSAVED = (ERID(0) = 0)
End Sub

Sub BldMatchDEF()
      PAGE.DFCNT = 0: RFCNT = 9
      ReDim DFNUM(10), DFPP(10)
      ReDim RFSEP(RFCNT), RFLTH(RFCNT), RFLBL$(RFCNT), RFABR$(RFCNT)
      Call RptOptGET("MNGP", MINGAP)
      If (Not OKAY) Or (MINGAP = 0) Then MINGAP = 2
      Call RptOptGET("MXGP", MAXGAP)
      If (Not OKAY) Or (MAXGAP = 0) Then MAXGAP = 5

58500 ' Select Report Fields - Based Upon Report Options
      OKAY = 0: Z = 0
      If (PAGE.BWIDTH = 0) Then
         BWIDTH = PRT.FWIDTH - PAGE.IMARGIN - PAGE.OMARGIN
         Else
         BWIDTH = PAGE.BWIDTH
      End If
      DFNUM(1) = 1: Z = 1
      For X = 1 To MFCNT
         Z = Z + 1: DFNUM(Z) = (SX(X) \ 256) + 1
      Next X
      PAGE.DFCNT = Z: PAGE.PFCNT = 0

58510 ' Create Report Definition Tables
      X = 1: RFNAME$ = "08001": GoSub 58515     'COUNT
      X = 2: RFNAME$ = "01022": GoSub 58515     'Surname
      X = 3: RFNAME$ = "01021": GoSub 58515     'Given Name
      X = 4: RFNAME$ = "01030": GoSub 58515     'Birth Date
      X = 5: RFNAME$ = "02020": GoSub 58515     'Marriage Date
      X = 6: RFNAME$ = "01040": GoSub 58515     'Death Date
      X = 7: RFNAME$ = "01032": GoSub 58515     'Birth Place
      X = 8: RFNAME$ = "01042": GoSub 58515     'Death Place
      X = 9: RFNAME$ = "01015": GoSub 58515     'Gender
      GoTo 58525

58515 ' Process Report Field
      Call GetRFLD(RFNAME$, FX)
      If RFSEP(X) = 0 Then RFSEP(X) = MINGAP
      If FOUND Then
         RDFFld = RFLDS(FX)
         If RFLTH(X) = 0 Then RFLTH(X) = Asc(RDFFld.RLTH)
         ALTH = Asc(RDFFld.ALTH): LLTH = Asc(RDFFld.LLTH)
         HLTH = Asc(RDFFld.HLTH)
         RFABR$(X) = MID$(RFLD$(FX), 1, ALTH)
         RFLBL$(X) = MID$(RFLD$(FX), ALTH + LLTH + 1, HLTH)
      End If
      Return

58525 ' Check Report Width
      GoSub 58530
      If RWIDTH <= BWIDTH Then GoTo 59540
      GoTo 58535

58530 ' Determine Required Width
      RWIDTH = 0: SWIDTH = 0: FX = 0
      While FX < PAGE.DFCNT
          FX = FX + 1: X = DFNUM(FX): FSEP = RFSEP(X)
          If (X > 0) And (FSEP > 0) Then
             Y = RFLTH(X)
             If FX > 1 Then SWIDTH = SWIDTH + Len(RFABR$(X)) + FSEP + Y
             Z = Len(RFLBL$(X)): If Y < Z Then Y = Z
             If PRT.OPT > 1 Then RWIDTH = RWIDTH + 2 + Y
          End If
      Wend
      If SWIDTH < 78 Then SWIDTH = 78
      Return

58535 ' Check Report Width
      'If PRT.OPT > 1 Then
         RPWIDTH = RWIDTH + PAGE.IMARGIN + PAGE.OMARGIN
         PWIDTH = BWIDTH + PAGE.IMARGIN + PAGE.OMARGIN
         If PWIDTH < RPWIDTH Then
            SysVAR.NVAR1 = PWIDTH: SysVAR.NVAR2 = RPWIDTH
            Call PutMSG("XXXXM009")
            If CA$ <> SNGLKEY$(1) Then OKAY = 0: GoTo 58560
            BWIDTH = RWIDTH
         End If
      '   Else
      '   BWIDTH = 0: GoTo 58560
      'End If

59540 ' Adjust to Fit
      'If PRT.OPT = 1 Then GoTo 58560
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
      PLINE$ = Space$(BWIDTH + PAGE.IMARGIN + PAGE.OMARGIN)
      SLINE$ = Space$(SWIDTH)
      PAGE.BWIDTH = BWIDTH
      OKAY = -1
End Sub

Rem $STATIC
Sub BldMatchTITLE()
995   ' Build Title Line
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
         HTFLINE$(HTFLMAX) = "T10" + RTrim$(TL$)
      End If
      OKAY = -1
End Sub

Rem $DYNAMIC
Sub ChkForPAUSE()
    
    If a = 27 Then Exit Sub
    CA$ = INKEY$
    If CA$ <> "" Then
       Call PutMSG("VLDTM012")
       'COLOR FG(5), BG(5): LOCATE 25, 74, 0: Print XRID$;
       Call GetKEY
       NOPAUSE = ((Len(CA$) = 2) And (a = 81))
       If a <> 27 Then
          Call PutMSG("VLDTM013")
          Else
          UPDTFile = 0
       End If
    End If
End Sub

Sub DataCHECK()

1100 ' Prepare for Data Validation
     ReDim SPID(50), ERRCNT(12), ScrnERRCNT(12)
     ReDim ERID(FF1Hdr.MAXID), byr(FF1Hdr.MAXID)

1105 ' Open files for processing
     Call frmFHSWUPgm.ShowStats("DATVLDT", -1, "")
     SysVAR.SUBJECT = FDFSetup.NAME
     Call RptOPEN("DVLD"): If Not OKAY Then GoTo 1190
     Call FAMOpen("READ"): If Not OKAY Then GoTo 1190
     LastRID = 0: xRID$ = Space$(5): NOPAUSE = 0
     FSNDX$ = Space$(4): SNDX$ = FSNDX$: SURNM$ = Space$(Len(FF1Rec.SURNM))
     ENV.ScrnCLR = -1
     Call PrintOPEN: If Not OKAY Then GoTo 1190
     PRT.FWIDTH = PAGE.WIDTH
     Dim RTYPE$(7)
     For RT = 1 To 7
         xX$ = "RT" + fnx$(RT): Call GetRVAR(xX$, X)
         If FOUND Then RTYPE$(RT) = RVAR$(X) + " "
     Next RT
     Dim OTYPE$(1)
     Call GetRVAR("ADDRESS", X): If FOUND Then OTYPE$(1) = RVAR$(X) + " "
     Dim DTYPE$(4)
     For DT = 1 To 4
         xX$ = "DT" + fnx$(DT): Call GetRVAR(xX$, X)
         If FOUND Then DTYPE$(DT) = RVAR$(X) + " "
     Next DT
     Call RptOptGET("CKND", X): VALNMDT = (X = 1)
     Call RptOptGET("CKMD", X): VALMYDT = (X = 1)
     Call RptOptGET("CKAD", X): VALADDT = (X = 1)
     Call RptOptGET("CKOD", X): VALMSDT = (X = 1)
     CHKMREC = (VALADDT + VALMYDT)
     VALIDATES = (VALNMDT Or VALMYDT Or VALADDT Or VALMSDT)
     Call RptOptGET("CKAG", X): CHKAGES = (X = 1)
     If CHKAGES Then
        Call RptOptGET("CKMA", X): CHKMAGE = (X = 1)
        Call RptOptGET("CKSA", X): CHKSPAGE = (X = 1)
        Call RptOptGET("CKAB", X): CHKPAGE = (X = 1)
     End If
     Call RptOptGET("ESBY", X): ESTIMBYR = (X = 1)
     If ESTIMBYR Then
        Call RptOptGET("ESMA", ESTMAGE)
        Call RptOptGET("ESPA", ESTPAGE)
     End If
     Call RptOptGET("CKUP", X):
     CHKSPAR = (X = 4): CHKPMAR = (X = 1) Or CHKSPAR
     Call RptOptGET("CKBM", X): CHKBLKMAR = (X = 1)
     Call RptOptGET("CKFS", X): CHKSURNM = (X = 1)
     If CHKSURNM Then
        Call RptOptGET("CKSN", X): CHKSNDX = (X = 1)
        CHKSMNM = Not CHKSNDX
     End If
     Call RptOptGET("CKSX", X): CHKSEX = (X = 1)
     If CHKSEX Then
        Call RptOptGET("CKPS", X): CHKPSEX = (X = 1)
        Call RptOptGET("CKSS", X): CHKSPSEX = (X = 1)
     End If
     Call RptOptGET("HIAG", MAXAGE)
     Call RptOptGET("MDIF", MDIF)
     Call RptOptGET("LOMA", AGEYMAR)
     Call RptOptGET("HIMA", AGEOMAR)
     Call RptOptGET("LOAB", AGEYPAR)
     Call RptOptGET("HIAB", AGEOPAR)
     CHKMREC = CHKMREC Or CHKMAGE Or CHKSPAGE Or CHKPMAR Or CHKBLKMAR Or CHKSPSEX
     BLDBYTBL = CHKAGES Or CHKMAGE Or CHKSPAGE Or CHKPAGE Or CHKPSEX Or CHKSPSEX
     XSEX$ = " ": MALE$ = FGENDR$(1): FEMALE$ = FGENDR$(2): Unknown$ = FGENDR$(3)
     FATHER$ = Relation$("L", MALE$, -1, -1, RRELCD$())
     MOTHER$ = Relation$("L", FEMALE$, -1, -1, RRELCD$())
     CYR = Val(Right$(DATE$, 4)): NUMED$ = Space$(6)

     ' Print Validation Criteria in Report
     'If PRT.opt > 1 Then
        PRT.SECTION = "0": Call PrintHEAD
        PRT.SECTION = "W": Call PrintTITLE
        If VALIDATES <> 0 Then PRT.SECTION = "1": Call PrintTITLE
        If CHKAGES Then
           SysVAR.NVAR1 = MAXAGE
           PRT.SECTION = "2": Call PrintTITLE
        End If
        If CHKMAGE Then
           SysVAR.NVAR1 = AGEYMAR: SysVAR.NVAR2 = AGEOMAR
           PRT.SECTION = "3": Call PrintTITLE
        End If
        If CHKSPAGE Then
           SysVAR.NVAR1 = MDIF
           PRT.SECTION = "4": Call PrintTITLE
        End If
        If CHKPAGE Then
           SysVAR.NVAR1 = AGEYPAR: SysVAR.NVAR2 = AGEOPAR
           PRT.SECTION = "5": Call PrintTITLE
        End If
        If CHKPMAR Then PRT.SECTION = "6": Call PrintTITLE
        If CHKBLKMAR Then PRT.SECTION = "7": Call PrintTITLE
        If CHKSURNM Then PRT.SECTION = "8": Call PrintTITLE
        If CHKSEX Then PRT.SECTION = "9": Call PrintTITLE
        If CHKPSEX Then PRT.SECTION = "A": Call PrintTITLE
        If CHKSPSEX Then PRT.SECTION = "B": Call PrintTITLE
        Call WriteLINE
        PRT.SECTION = "X": Call PrintTITLE
     'End If
     If a = 27 Then GoTo 1190

1110 ' Build BYR() Tables
     If Not BLDBYTBL Then GoTo 1200
     If CHKPSEX Or CHKSPSEX Then
        On Local Error GoTo 1189
        SysVAR.FILENAME = WRKFPFX$ + "FH$VDTMP"
        Open SysVAR.FILENAME For Binary Access Read Write As #5
        On Local Error GoTo 0
     End If
     Call PutMSG("VLDTM018")
     For X = 1 To FF1Hdr.MAXID: Call FF1GetRec(X)
         RSet xRID$ = fnx$(X)
         'COLOR FG(5), BG(5): LOCATE 25, 74, 0: Print XRID$;
         CA$ = INKEY$
         If CA$ <> "" Then
            a = Asc(CA$): If a = 27 Then id = 0
         End If
         zBYR = FF1Rec.BY: DYR = FF1Rec.DY
         If (zBYR = 0) And (ESTIMBYR > 0) And (DYR > ESTAGE) And (DYR < 9999) Then
            zBYR = -(DYR - ESTAGE)
         End If
         byr(X) = zBYR
         If CHKPSEX Or CHKSPSEX Then Put #5, X, FF1Rec.SEX
     Next X: If CA$ = Chr$(27) Then GoTo 1190
     Call ClearMSG
     GoTo 1200

1189 ' Error Opening FH$VLDT
     Call ERRMessage("XXXXM097", ERR, ERL)
     Resume 1190

1190 ' Early EXIT
     Call frmFHSWUPgm.ShowStats("DATVLDT", 99, "")
     Call FamCLOSE: Exit Sub

1200 ' Begin Data Validation
     IDX = 0: IDCNT = 0
1210 ' Select Next Name Record
     IDX = IDX + 1: If IDX > FF1Hdr.MAXID Then GoTo 1350
     If SelSET(IDX) = 0 Then GoTo 1210
     IDCNT = IDCNT + 1
     Call frmFHSWUPgm.ShowStats("DATVLDT", 11, Str$(IDCNT))
     id = IDX: RSet xRID$ = fnx$(id)
     Call ChkForPAUSE: If a = 27 Then GoTo 1350
     'COLOR FG(5), BG(5): LOCATE 25, 74, 0: Print XRID$;
     Call FF1GetRec(id)
     zBYR = FF1Rec.BY: DYR = FF1Rec.DY
     FID = FF1Rec.FID: zMID = FF1Rec.MID
     If (zBYR = 0) And BLDBYTBL Then zBYR = Abs(byr(id))
     SEX$ = FF1Rec.SEX
     LSet SURNM$ = FF1Rec.SURNM: CHILD = FF1Rec.OLDCH
     MREC = FF1Rec.SPOUSE
     If (CHKSURNM And CHKSNDX) And (SEX$ = MALE$) And (CHILD <> 0) Then
        Call ComputeSNDX(SURNM$, FSNDX$)
     End If
     If VALNMDT Then
        Call ValiDATE(FF1Rec.BMD, FF1Rec.BY)
        If (Not OKAY) Then SysVAR.SUBJECT = DTYPE$(3): Call DataERR(1, 1)
        Call ValiDATE(FF1Rec.DMD, FF1Rec.DY)
        If (Not OKAY) Or ((FF1Rec.DY > 0) And (FF1Rec.DY < FF1Rec.BY)) Then
           SysVAR.MVAR = XSCRNDate$(FF1Rec.DMD, FF1Rec.DY)
           SysVAR.SUBJECT = DTYPE$(4): Call DataERR(1, 1)
        End If
     End If
     If CHKSEX Then
        If (SEX$ <> MALE$) And (SEX$ <> FEMALE$) And (SEX$ <> Unknown$) Then
           SysVAR.MVAR = FF1Rec.SEX: Call DataERR(8, 9)
        End If
     End If
     If CHKPSEX Then
        If FID > 0 Then
           Get #5, FID, XSEX$
           If XSEX$ <> MALE$ Then
              SysVAR.SUBJECT = FATHER$: SysVAR.MVAR = XSEX$
              Call DataERR(9, 10)
           End If
        End If
        If zMID > 0 Then
           Get #5, zMID, XSEX$
           If XSEX$ <> FEMALE$ Then
              SysVAR.SUBJECT = MOTHER$: SysVAR.MVAR = XSEX$
              Call DataERR(9, 10)
           End If
        End If
     End If
     If CHKAGES And (DYR < 9999) And (zBYR > 0) Then
        If DYR > 0 Then AGE = DYR - zBYR Else AGE = CYR - zBYR
        If (AGE < 0) Or (AGE > MAXAGE) Then
           SysVAR.NVAR1 = AGE: Call DataERR(2, 2)
        End If
     End If
     If CHKPAGE And (zBYR > 0) Then
        PBYR = byr(FID)
        If (PBYR <> 0) And (zBYR > 0) Then
           AGE = zBYR - Abs(PBYR)
           If ((AGE < AGEYPAR) Or (AGE > AGEOPAR)) Then
              SysVAR.SUBJECT = FATHER$: SysVAR.NVAR1 = AGE
              Call DataERR(4, 5)
           End If
        End If
        PBYR = byr(zMID)
        If (PBYR <> 0) And (zBYR > 0) Then
           AGE = zBYR - Abs(PBYR)
           If ((AGE < AGEYPAR) Or (AGE > AGEOPAR)) Then
              SysVAR.SUBJECT = MOTHER$: SysVAR.NVAR1 = AGE
              Call DataERR(4, 5)
           End If
        End If
     End If
     If (Not CHKMREC) Or (MREC = 0) Then GoTo 1300
     RT = 2: MCNT = 0
     While (MREC <> 0) And (MCNT < 50)
        MCNT = MCNT + 1: Call FF3GetRec(MREC)
        If FF3SPOUSE.SP1ID = id Then
           zSPID = FF3SPOUSE.SP2ID: MREC = FF3SPOUSE.SP1NXT
           Else
           zSPID = FF3SPOUSE.SP1ID: MREC = FF3SPOUSE.SP2NXT
        End If
        GoSub 1250
     Wend
     If (Not CHKPMAR) And ((Not CHKSURNM) Or (SEX$ <> MALE$)) Then GoTo 1300
     While CHILD > 0
        Call FF1GetRec(CHILD)
        FID = FF1Rec.FID: zMID = FF1Rec.MID: SysVAR.NVAR1 = CHILD
        If id = FID Then
           PID = zMID: CHILD = FF1Rec.FCH: SysVAR.MVAR = MOTHER$
           Else
           PID = FID: CHILD = FF1Rec.MCH: SysVAR.MVAR = FATHER$
        End If
        If CHKPMAR And (CHKSPAR Or (PID <> 0)) Then
           FOUND = 0: X = 1
           While (X < (MCNT + 1)) And (Not FOUND)
              FOUND = (SPID(X) = PID)
              If Not FOUND Then X = X + 1
           Wend
           If Not FOUND Then Call DataERR(5, 6)
        End If
        If CHKSURNM And (SEX$ = MALE$) And ((Asc(FF1Rec.STATUS) And 1) = 0) Then
           If CHKSNDX Then Call ComputeSNDX(FF1Rec.SURNM, SNDX$)
           If ((CHKSMNM And (SURNM$ <> FF1Rec.SURNM)) Or (CHKSNDX And (FSNDX$ <> SNDX$))) Then
              Call DataERR(7, 8)
           End If
        End If
     Wend
     GoTo 1300

1250 ' Check Marriage Record Dates, Age at Marriage & Spouse Sex Code
     SPID(MCNT) = zSPID: SPBYR = byr(zSPID): MADR = FF3SPOUSE.LOC
     D1MD = FF3SPOUSE.D1MD: D1Y = FF3SPOUSE.D1Y
     D2MD = FF3SPOUSE.D2MD: D2Y = FF3SPOUSE.D2Y
     If VALMYDT Then OT = 0: GoSub 1380
     If CHKBLKMAR Then
        If (zSPID = 0) And (D1MD = 0) And (D1Y = 0) And (D2MD = 0) And (D2Y = 0) And (FF3SPOUSE.MLOC = 0) And (FF3SPOUSE.COM = 0) And (FF3SPOUSE.LOC = 0) Then
           SysVAR.NVAR1 = MCNT: Call DataERR(6, 7)
        End If
     End If
     If CHKMAGE And (zBYR > 0) And (D1Y > 0) Then
        MAGE = D1Y - zBYR
        If (MAGE < AGEYMAR) Or (MAGE > AGEOMAR) Then
           SysVAR.NVAR1 = zSPID: SysVAR.NVAR2 = MAGE
           Call DataERR(3, 3)
        End If
     End If
     If ((id <= zSPID) Or (SelSET(zSPID) = 0)) Then
        If CHKSPAGE And (zBYR > 0) And (SPBYR > 0) Then
           MAGE = Abs(zBYR - SPBYR)
           If (MAGE > MDIF) Then
              SysVAR.NVAR1 = zSPID: SysVAR.NVAR2 = MAGE
              Call DataERR(3, 4)
           End If
        End If
        If CHKSPSEX And (zSPID > 0) Then
           Get #5, zSPID, XSEX$
           If SEX$ = XSEX$ Then SysVAR.NVAR1 = zSPID: Call DataERR(8, 11)
        End If
     End If
     If VALADDT And (MADR <> 0) Then
        ADR = MADR: GoSub 1370
     End If
     Return
    
1300 ' Validate Other Dates
     Call FF1GetRec(id): ADR = FF1Rec.ADRS
     If VALADDT And (ADR <> 0) Then RT = 1: GoSub 1370
     If VALMSDT Then
        OREC = FF1Rec.EDUC: If OREC <> 0 Then RT = 3: GoSub 1360
        OREC = FF1Rec.WORK: If OREC <> 0 Then RT = 4: GoSub 1360
        OREC = FF1Rec.MIL: If OREC <> 0 Then RT = 5: GoSub 1360
        OREC = FF1Rec.HLTH: If OREC <> 0 Then RT = 6: GoSub 1360
        OREC = FF1Rec.EVENT: If OREC <> 0 Then RT = 7: GoSub 1360
     End If

1350 ' Go Get  Next ID
     If a = 27 Then GoTo 2900
     If (IDX < FF1Hdr.MAXID) Then GoTo 1210
     GoTo 1900

1360 ' Validate Miscellaneous Record Dates
     While OREC <> 0: Call FF3GetRec(OREC)
        If RT = 7 Then
           D1MD = FF3EVENT.D1MD: D1Y = FF3EVENT.D1Y
           OREC = FF3EVENT.NXT
           Else
           D1MD = FF3PNTRS.D1MD: D1Y = FF3PNTRS.D1Y
           D2MD = FF3PNTRS.D2MD: D2Y = FF3PNTRS.D2Y
           OREC = FF3PNTRS.NXT
        End If
        OT = 0: GoSub 1380
        If VALADDT And (RT < 7) Then
           ADR = FF3PNTRS.LOC: If ADR <> 0 Then GoSub 1370
        End If
     Wend
     Return

1370 ' Validate Address Dates
     OT = 1
     While ADR > 0: Call FF2GetRec(ADR)
        D1MD = FF2PFX.D1MD: D1Y = FF2PFX.D1Y
        D2MD = FF2PFX.D2MD: D2Y = FF2PFX.D2Y
        GoSub 1380
        ADR = FF2PFX.NXT
     Wend
     Return

1380 ' Validate Begin, End Dates
     Call ValiDATE(D1MD, D1YR)
     If (Not OKAY) Or (D1YR = 9999) Then
        SysVAR.SUBJECT = RTYPE$(RT) + OTYPE$(OT) + DTYPE$(1)
        Call DataERR(1, 1)
     End If
     If RT <> 7 Then
        Call ValiDATE(D2MD, D2YR)
        If (Not OKAY) Or (D2Y > 0) And (D2Y < D1Y) Then
           SysVAR.SUBJECT = RTYPE$(RT) + OTYPE$(OT) + DTYPE$(2)
           Call DataERR(1, 1)
        End If
     End If
     Return

1900 ' Return from VALIDATION operation
     If ERRCNT(12) > 0 Then
        SysVAR.NVAR1 = ERRCNT(12): SysVAR.NVAR2 = ERRCNT(0)
        Call PutMSG("VLDTM019")
     End If

2500 ' END OF REPORT
     If a = 27 Then GoTo 2900
     'If PRT.opt = 1 Then GoTo 2950
     LSet PLINE$ = " ": Call WriteLINE
     SysVAR.NVAR1 = ERRCNT(12): SysVAR.NVAR2 = ERRCNT(0)
     PRT.SECTION = "Y": Call PrintTITLE
     If ERRCNT(12) > 0 Then
        Call WriteLINE
        PRT.SECTION = "Z": Call PrintTITLE
        For ET = 1 To 11
            If ERRCNT(ET) > 0 Then
               EMSG$ = "ET" + ZFILL$(ET, 2)
               Call GetRVAR(EMSG$, X)
               If FOUND Then
                  RSet NUMED$ = fnx$(ERRCNT(ET))
                  LSet PLINE$ = NUMED$ + "  " + RVAR$(X)
                  Call PrintLINE: If a = 27 Then ET = 11
               End If
            End If
        Next ET
     End If
     GoTo 2950

2900 ' Early Termination
     'If PRT.opt = 1 Then
     '   If id > 0 Then
     '      SysVAR.NVAR1 = id
     '      Call PutMSG("VLDTM010")
     '      Else
     '      Call PutMSG("VLDTM011")
     '   End If
     '   Else
        Call RptCLOSE
     'End If
     GoTo 2990

2950 ' Validation Procedure Complete
     'If PRT.opt = 1 Then
     '   Call PutMSG("VLDTM020")
     '   Else
        Call RptCLOSE
     'End If

2990 ' Close Files
     Call FamCLOSE
     Call PutMSG("")
     If CHKPSEX Or CHKSPSEX Then
        On Local Error GoTo 2995
        SysVAR.FILENAME = WRKFPFX$ + "FH$VDTMP"
        Kill SysVAR.FILENAME
     End If
     GoTo 2999
2995 Resume 2999
2999 ' END OF ROUTINE
     Call frmFHSWUPgm.ShowStats("DATVLDT", 99, "")

End Sub

Sub DataERR(ERRGrp, EMNUM)
   If a <> 27 Then Call ChkForPAUSE
   If a = 27 Then Exit Sub
   If EMNUM > 0 Then ERRCNT(EMNUM) = ERRCNT(EMNUM) + 1
   If ERRGrp <> 0 Then
      ScrnERRCNT(ERRGrp) = ScrnERRCNT(ERRGrp) + 1
      Call frmFHSWUPgm.ShowStats("DATVLDT", ERRGrp, Str$(ScrnERRCNT(ERRGrp)))
      ERRCNT(12) = ERRCNT(12) + 1
      Call frmFHSWUPgm.ShowStats("DATVLDT", 10, Str$(ERRCNT(12)))
   End If
   'EFLD$ = "EC" + ZFILL$(ERRGrp, 2)
   'If CurFMT <> FMT019 Then CurFMT = FMT019: Fmt = FMTTAB(CurFMT)
   'Call FmtFindFLD(EFLD$)
   'If FOUND Then RSet UTXT$(FFLD.UTXT) = FNX$(ScrnERRCNT(ERRGrp))
   'Call FmtFindFLD("ETOT")
   'If FOUND Then
   '   RSet UTXT$(FFLD.UTXT) = FNX$(ERRCNT(12))
   '   Call WinDisplayDATA(0)
   'End If
   CRID = Val(xRID$)
   If CRID > 0 Then
      If ERID(CRID) = 0 Then ERID(0) = ERID(0) + 1
      ERID(CRID) = ERID(CRID) + 1
   End If
   If CRID <> LastRID Then
      ERRCNT(0) = ERRCNT(0) + 1: LastRID = CRID
      Call frmFHSWUPgm.ShowStats("DATVLDT", 12, Str$(ERRCNT(0)))
      LSet PLINE$ = xRID$
      Else
      LSet PLINE$ = ""
   End If
   EMNM$ = "EM" + ZFILL$(EMNUM, 2)
   Call GetRVAR(EMNM$, X)
   If FOUND Then
      xERRMSG$ = RVAR$(X)
      Else
      xERRMSG$ = "Error Message " + EMNM$ + " Not Found"
   End If
   Mid$(PLINE$, 8, 255) = xERRMSG$
   Call HTFExpand(PLINE$, 1)
   'COLOR FG(5), BG(5): LOCATE 25, 1, 0: Print XRID$;
   'COLOR FG(1), BG(1): LOCATE 25, 7, 0: Print MID$(PLINE$, 8, 70); Tab(79);
   'If PRT.opt <> 1 Then
      Call PrintLINE
   '   Else
   '   If Not NOPAUSE Then
   '      Call ErrBEEP(0): Call GetKEY
   '      NOPAUSE = ((Len(CA$) = 2) And (A = 81))
   '      If Not PAGE.WAIT Then PAGE.WAIT = NOPAUSE
   '   End If
   'End If
   Exit Sub
End Sub

Rem $STATIC
Sub MatchCHECK()

    cDFM$ = CONFIG.DFM
    STATFMT = CURFMT: 'Call FmtDREST
    On Local Error GoTo MCError
MCGetIndex:
    'Call FMTFind("VLDTS005"): If Not FOUND Then GoTo MC190
    'Call FmtFindFLD("DSNM"): y = FFLD.UTXT: LSet UTXT$(y) = FT$(4)
    'Call MsgFORMAT: If A = 27 Then GoTo MC190
    'LSet FT$(4) = UTXT$(y)
    Call INDEXFile("READ")
    If Not OKAY Then GoTo MC190
    MFCNT = SFCNT - 1
    If MFCNT = 0 Then
       Call PutMSG("VLDTM024")
       GoTo MC190
    End If
    ReDim LMDATA$(MFCNT), MDATA$(MFCNT)
    'Fmt = FMTTAB(STATFMT)
    Call frmFHSWUPgm.ShowStats("DUPVLDT", -1, "")
    For MF = 1 To MFCNT
        X = SX(MF) \ 256
        Select Case X:
           Case 1: SNCHK = -1: Call frmFHSWUPgm.ShowStats("DUPVLDT", 1, "")
           Case 2: Call frmFHSWUPgm.ShowStats("DUPVLDT", 2, "")
           Case 3: BDCHK = -1: Call frmFHSWUPgm.ShowStats("DUPVLDT", 4, "")
           Case 4: MDCHK = -1: Call frmFHSWUPgm.ShowStats("DUPVLDT", 5, "")
           Case 5: BDCHK = -1: Call frmFHSWUPgm.ShowStats("DUPVLDT", 6, "")
           Case 6: BPCHK = -1: Call frmFHSWUPgm.ShowStats("DUPVLDT", 7, "")
           Case 7: BPCHK = -1: Call frmFHSWUPgm.ShowStats("DUPVLDT", 8, "")
           Case 8: Call frmFHSWUPgm.ShowStats("DUPVLDT", 3, "")
        End Select
    Next MF
    ReDim SELSET1(FF1Hdr.MAXID), SELSET2(FF1Hdr.MAXID)
    Call MatchGrp(1, SELSET1(), G1): If Not OKAY Then GoTo MC190
    Call frmFHSWUPgm.ShowStats("DUPVLDT", 10, Str$(SELSET1(0)))
    Call MatchGrp(2, SELSET2(), G2): If Not OKAY Then GoTo MC190
    Call frmFHSWUPgm.ShowStats("DUPVLDT", 11, Str$(SELSET2(0)))
    frmFHSWUPgm.Refresh
    FullFILE = (G1 = 1) Or (G2 = 1)
    ts = ENV.SortOPTS And 3: DS = (ENV.SortOPTS \ 4) And 7
    Select Case DS          ' Determine Date Format and Date Sort Key Length
       Case 0: CONFIG.DFM = "3": DSKL = 10
       Case 1: CONFIG.DFM = "1": DSKL = 10
       Case 2: CONFIG.DFM = "2": DSKL = 10
       Case 3: CONFIG.DFM = "3": DSKL = 7
       Case 4: CONFIG.DFM = "1": DSKL = 5
       Case 5: CONFIG.DFM = "3": DSKL = 4
       Case 6: CONFIG.DFM = "1": DSKL = 2
       Case 7: CONFIG.DFM = "2": DSKL = 2
    End Select

    ' Open files for processing
    SysVAR.SUBJECT = FDFSetup.NAME
    Call RptOPEN("MDAT"): If Not OKAY Then GoTo MC190
    Call RptOptGET("MNMC", MinMatch)
    If (Not OKAY) Or (MinMatch = 0) Then MinMatch = 2
    MaxMATCH = MinMatch: MaxMATCH2 = MinMatch
    Call frmFHSWUPgm.ShowStats("DUPVLDT", 12, Str$(MinMatch))
    frmFHSWUPgm.Refresh
    Call FAMOpen("READ"): If Not OKAY Then GoTo MC190
    'Call FmtFindFLD("DCNT"): If FOUND Then DCFX = FFLD.UTXT
    'Call FmtFindFLD("GCNT"): If FOUND Then GCFX = FFLD.UTXT
    ENV.ScrnCLR = -1: NOPAUSE = 0
    Call PrintOPEN: If Not OKAY Then GoTo MC190
    If BDCHK Then BDT$ = Space$(11): DDT$ = BDT$
    If MDCHK Then MDT$ = Space$(11)
    If BPCHK Then BPL$ = Space$(22): DPL$ = BPL$
    BDPCHK = (BDCHK Or BPCHK)
    Call BldMatchDEF: Call BldMatchTITLE: NOPAUSE = -1
    GoTo MCBeginReport

MC190: ' Early Exit
    CONFIG.DFM = cDFM$
    Call FamCLOSE: Exit Sub

MCBeginReport:
    SlctCNT = 0: SecSlctCNT = 0
    SX1 = 0: M1CNT = 0: SX2 = 0: M2CNT = 0: S2 = 0: GCNT = 0
    IDX = 0: XTEXT$ = Space$(57)
    ReDim ERID(FF1Hdr.MAXID), MatchRID(MaxMATCH), MatchRID2(MaxMATCH2)
    'If (DCFX <> 0) Then LSet UTXT$(DCFX) = LTrim$(Str$(SlctCNT))
    'If (GCFX <> 0) Then LSet UTXT$(GCFX) = LTrim$(Str$(GCNT))
    'Call WinDisplayDATA(STATFMT)
    Call frmFHSWUPgm.ShowStats("DUPVLDT", 13, Str$(SlctCNT))
    Call frmFHSWUPgm.ShowStats("DUPVLDT", 14, Str$(GCNT))
    PRT.SECTION = "1"

MCGetNextRID:
    IDX = IDX + 1: If IDX > FF1Hdr.MAXID Then GoTo MCChkLastGroup
    'COLOR FG(5), BG(5): LOCATE 25, 74, 0: Print SFILL$(IDX, 5);
    CA$ = INKEY$
    If CA$ <> "" Then
       a = Asc(Right$(CA$, 1)): OKAY = (a <> 27)
       If Not OKAY Then GoTo MCFinish
       NOPAUSE = 0
       SysVAR.NVAR1 = IDX: SysVAR.NVAR2 = FF1Hdr.MAXID
       Call PutMSG("VLDTM025")
    End If
    id = NDXID(IDX)
    If (id = 0) Or (SELSET1(id) = 0) And (SELSET2(id) = 0) Then GoTo MCGetNextRID
    Call FF1GetRec(id)
    If SNCHK Then
       If ENV.HUSBSURNM Then
          LSet xname$ = FMTNAME$(0)    ' Have Function Build SURNM$
          Else
          SURNM$ = FF1Rec.SURNM
       End If
    End If
    If BDPCHK Then Call GetBDEvents(BDT$, BPL$, DDT$, DPL$, AGE$)
    If MDCHK Then
       If FF1Rec.SPOUSE <> 0 Then
          Call FF3GetRec(FF1Rec.SPOUSE)
          LSet MDT$ = LTrim$(XSCRNDate$(FF3SPOUSE.D1MD, FF3SPOUSE.D1Y))
          Else
          LSet MDT$ = ""
       End If
    End If
    MATCH = -1
    For MF = 1 To MFCNT
        X = SX(MF) \ 256
        Select Case X
           Case 1: LSet XTEXT$ = TextDATA$(SURNM$): MKL = 20
           Case 2: LSet XTEXT$ = TextDATA$(FF1Rec.GIVEN): MKL = 30
           Case 3: LSet XTEXT$ = MID$(BDT$, 1, DSKL)
           Case 4: LSet XTEXT$ = MID$(MDT$, 1, DSKL)
           Case 5: LSet XTEXT$ = MID$(DDT$, 1, DSKL)
           Case 6: LSet XTEXT$ = TextDATA$(BPL$): MKL = 22
           Case 7: LSet XTEXT$ = TextDATA$(DPL$): MKL = 22
           Case 8: LSet XTEXT$ = FF1Rec.SEX
        End Select
        If (X > 2) And (X < 6) Then
           SKL = DSKL
           Else
           If X < 8 Then
              If ts = 2 Then
                 SKL = 4
                 Else
                 SKL = (SX(MF) And 127)
                 MKL = MKL - 4 * (ts > 2)
                 If SKL > MKL Then SKL = MKL
              End If
              Else
              SKL = 1
           End If
        End If
        MDATA$(MF) = RTrim$(XTEXT$)
        If MATCH Then
           MATCH = (MID$(MDATA$(MF), 1, SKL) = MID$(LMDATA$(MF), 1, SKL))
        End If
    Next MF
    If MATCH Then
       M1CNT = M1CNT - (SELSET1(id) > 0)
       M2CNT = M2CNT - (SELSET2(id) > 0)
       If SELSET1(id) > 0 Then
          If ((M1CNT + S2) < MinMatch) Or (M2CNT = 0) Then
             SX1 = SX1 + 1
             If SX1 > MaxMATCH Then
                MaxMATCH = MaxMATCH + 10
                ReDim Preserve MatchRID(MaxMATCH)
             End If
             MatchRID(SX1) = id
             Else
             ERID(id) = 1
          End If
       End If
       If SELSET2(id) > 0 Then
          If ((M1CNT + S2) < MinMatch) Or (M1CNT = 0) Then
             If SELSET1(id) = 0 Then
                S2 = S2 + 1: SX2 = SX2 + 1
                If SX2 > MaxMATCH2 Then
                   MaxMATCH2 = MaxMATCH2 + 10
                   ReDim Preserve MatchRID2(MaxMATCH2)
                End If
                MatchRID2(SX2) = id
             End If
             Else
             If SELSET1(id) = 0 Then
                ERID(id) = 2: S2 = S2 + 1
             End If
          End If
       End If
       GoTo MCGetNextRID
    End If
MCChkLastGroup:
    If ((M1CNT + S2) >= MinMatch) And (M1CNT > 0) And (M2CNT > 0) Then
       X = 0: While X < SX1: X = X + 1: ERID(MatchRID(X)) = 1: Wend
       X = 0: While X < SX2: X = X + 1: ERID(MatchRID2(X)) = 2: Wend
       SlctCNT = SlctCNT + M1CNT
       SecSlctCNT = SecSlctCNT + S2
       GCNT = GCNT + 1: LP = 2
       For X = 1 To PAGE.DFCNT
           DFN = DFNUM(X)
           'If PRT.OPT > 1 Then
              PP = DFPP(X): If PP = 0 Then PP = 1
              FLTH = RFLTH(DFN)
              If DFN > 1 Then
                 Mid$(PLINE$, PP, FLTH) = LMDATA$(X - 1)
                 Else
                 If FullFILE Then
                    L = (FLTH \ 2) + 3
                    Mid$(PLINE$, PP, FLTH) = SFILL$(M1CNT + S2, L)
                    Else
                    L = (FLTH - 1) \ 2
                    Mid$(PLINE$, PP, FLTH) = SFILL$(M1CNT, L) + "," + LTrim$(Str$(S2))
                 End If
              End If
           'End If
           If X > 1 Then
              If X > 2 Then Mid$(SLINE$, LP, 1) = ";": LP = LP + 2
              Mid$(SLINE$, LP, 255) = RFABR$(DFN) + LMDATA$(X - 1)
              LP = Len(RTrim$(SLINE$)) + 1
           End If
       Next X
       'If DCFX <> 0 Then
          If FullFILE Then
             'LSet UTXT$(DCFX) = LTrim$(Str$(SlctCNT + SecSlctCNT))
             Call frmFHSWUPgm.ShowStats("DUPVLDT", 13, Str$(SlctCNT + SecSlctCNT))
             Else
             'LSet UTXT$(DCFX) = LTrim$(Str$(SlctCNT)) + "," + LTrim$(Str$(SecSlctCNT))
             Call frmFHSWUPgm.ShowStats("DUPVLDT", 13, LTrim$(Str$(SlctCNT)) + "," + LTrim$(Str$(SecSlctCNT)))
          End If
       'End If
       'If (GCFX <> 0) Then LSet UTXT$(GCFX) = LTrim$(Str$(GCNT))
       Call frmFHSWUPgm.ShowStats("DUPVLDT", 14, Str$(GCNT))
       frmFHSWUPgm.Refresh
       'Call WinDisplayDATA(STATFMT)
       Call PrintLINE: If a = 27 Then GoTo MCFinish
       'COLOR FG(5), BG(5): LOCATE 25, 1, 0: Print RFABR$(1);
       'If FullFILE Then
       '   Print LTrim$(Str$(M1CNT + S2));
       '   Else
       '   Print LTrim$(Str$(M1CNT)) + "," + LTrim$(Str$(S2));
       'End If
       'COLOR FG(1), BG(1): X25 = POS(0)
       'Print (Left$(SLINE$, 79 - X25));
       'If Not NOPAUSE Then
          'COLOR FG(5), BG(5): LOCATE 25, 76, 0: Print "...";
       '   Call ErrBEEP(0): Call GetKEY
       '   NOPAUSE = ((Len(CA$) = 2) And (a = 81))
       '   If Not PAGE.WAIT Then PAGE.WAIT = NOPAUSE
          'COLOR FG(1), BG(1): LOCATE 25, 76, 0
          'Print (MID$(SLINE$, 77 - X25, 3));
       'End If
       LSet SLINE$ = ""
       If a = 27 Then GoTo MCFinish
    End If
    If IDX > FF1Hdr.MAXID Then GoTo MCFinish
    For MF = 1 To MFCNT: LMDATA$(MF) = MDATA$(MF): Next MF
    M1CNT = -(SELSET1(id) > 0): SX1 = 0
    M2CNT = -(SELSET2(id) > 0): SX2 = 0
    If M1CNT > 0 Then
       MatchRID(1) = id: SX1 = 1
       Else
       MatchRID2(1) = id: SX2 = 1
    End If
    S2 = SX2
    GoTo MCGetNextRID

MCError:
    ' Error During Processing
    Call ERRMessage("VLDTM027", ERR, ERL)
    Resume MCFinish

MCFinish:
    ERID(0) = SlctCNT
    Call PutMSG("VLDTM026")
    'If PRT.OPT > 1 Then
       PRT.SECTION = "2"
       Call WriteLINE
       If FullFILE Then
          SysVAR.MVAR = LTrim$(Str$(SlctCNT + SecSlctCNT))
          Else
          SysVAR.MVAR = "(" + LTrim$(Str$(SlctCNT)) + "," + LTrim$(Str$(SecSlctCNT)) + ")=" + LTrim$(Str$(SlctCNT + SecSlctCNT))
       End If
       SysVAR.NVAR1 = GCNT
       Call PrintTITLE
       Call RptCLOSE: If PRT.AGAIN Then GoTo MCBeginReport
    'End If
    CONFIG.DFM = cDFM$
    Call RptCLOSE
    Call FamCLOSE
    Call frmFHSWUPgm.ShowStats("DUPVLDT", 99, "")
    Call PutMSG("")

End Sub

Sub MatchGrp(N, id(), G)

GetMatchGrp:
    SysVAR.NVAR1 = N: OKAY = -1
    Call PutMSG("VLDTM030")
    If a = 27 Then OKAY = 0: Exit Sub
    G = Val(CA$): If (G < 1) Or (G > 4) Then G = 1
    Select Case G
       Case 1: For X = 1 To FF1Hdr.MAXID: id(X) = 1: Next X
               XCNT = FF1Hdr.MAXID
       Case 2: Call RWRKOpen("READ", 0, 0)
               If OKAY Then
                  Call RWRKGet("BASE", RC, PTR&)
                  Call PutMSG("XXXXM064")
                  While (0 < RC) And (RC < 9)
                     If WFRELREC.id > 0 Then id(WFRELREC.id) = 1
                     Call RWRKGet("WRKREC", RC, PTR&)
                  Wend
                  Call PutMSG("")
                  Close #5
               End If
       Case 3: Call SWRKBldList(id(), X, 0)
       Case 4: Call frmFHSWMsgD.GetIDRange(ID1, ID2)
               If OKAY Then
                  If (ID2 < 1) Or (ID2 > FF1Hdr.MAXID) Then ID2 = FF1Hdr.MAXID
                  If (ID1 < 1) Or (ID1 > ID2) Then ID1 = 1
                  If (ID1 > 1) Then
                     XID1 = 1: XID2 = ID1 - 1
                     Else
                     If ID2 < FF1Hdr.MAXID Then
                        XID1 = ID2 + 1: XID2 = FF1Hdr.MAXID
                     End If
                  End If
                  For X = ID1 To ID2: id(X) = 1: Next X
                  XCNT = ID2 - ID1 + 1
               End If
    End Select
    If Not OKAY Then GoTo GetMatchGrp
    If (1 < G) And (G < 4) Then
       XCNT = 0
       For X = 1 To FF1Hdr.MAXID
           If id(X) <> 0 Then id(X) = 1: XCNT = XCNT + 1
       Next X
    End If
    'If CurFMT <> WIN.WFMT Then
    '   CurFMT = WIN.WFMT: Fmt = FMTTAB(CurFMT)
    'End If
    'FLDNM$ = "CNT" + LTrim$(Str$(N))
    'Call FmtFindFLD(FLDNM$)
    'If FOUND Then
    '   If FFLD.UTXT > 0 Then
    '      RSet UTXT$(FFLD.UTXT) = LTrim$(Str$(XCNT))
    '      Call WinDisplayDATA(0)
    '   End If
    'End If
    id(0) = XCNT: OKAY = -1

End Sub

Rem $DYNAMIC
Sub PtrCHECK()

P1000: ' Validate File Headers and Pointers
     ALLREC = (SCNT = FF1Hdr.MAXID)
     If Not ALLREC Then
        SysVAR.NVAR1 = SCNT: SysVAR.NVAR2 = FF1Hdr.MAXID
        Call PutMSG("Checking [NVAR1] of [NVAR2] records...")
        If Asc(CA$) = 27 Then GoTo 1090
     End If
     Call PutMSG("VLDTM002")
     If Asc(CA$) = 27 Then GoTo 1090
     UPDTFile = (CA$ = SNGLKEY$(1))
     On Local Error GoTo 1050
     GoTo P1100

1050 ' Error During Processing
     ' SYSVAR.NVAR1 = CRID: CALL PutMSG("CRID=[NVAR1]..."): CALL GetRESP
     Call ERRMessage("VLDTM027", ERR, ERL)
     Resume 1090

1090 ' EARLY EXIT
     Call FamCLOSE
     Call frmFHSWUPgm.ShowStats("PTRVLDT", 99, "")
     Exit Sub

P1100: ' Open files for processing
     Call frmFHSWUPgm.ShowStats("PTRVLDT", -1, "")
     IDCNT = 0:
     SysVAR.SUBJECT = FDFSetup.NAME
     Call RptOPEN("PVLD"): If Not OKAY Then GoTo 1090
     If UPDTFile Then Call FAMOpen("VLDT") Else Call FAMOpen("READ")
     If Not OKAY Then GoTo 1090
     MAXID = FF1Hdr.MAXID: HIREC2 = FF2Hdr.HIREC
     HIREC3& = FFRBA&(FF3HDR.HIREC, 1)
     If (CRID < 1) Or (CRID > MAXID) Then CRID = 1
     ReDim PTRCK1(MAXID), PTRCK2(HIREC2), ERRCNT(20)
     Call PtrCHK3("OPEN", HIREC3&, 0): If Not OKAY Then GoTo 1090
     ReDim ERID(MAXID)
     MAXAGE = 9999: CYR = Val(Right$(DATE$, 4))
     Call RptOptGET("CKDD", X): If FOUND Then FixMissingDD = (X = 1)
     If FixMissingDD Then
        Call RptOptGET("HIAG", X): If FOUND Then MAXAGE = X
     End If
     LastRID = 0: xRID$ = Space$(5): NOPAUSE = 0
     ENV.ScrnCLR = -1
     Call PrintOPEN: If Not OKAY Then GoTo 1090
     PRT.FWIDTH = PAGE.WIDTH: If Len(PLINE$) < 100 Then PLINE$ = Space$(100)
     Call PutMSG("VLDTM003")
     Dim RTYPE$(11)
     For RT = 1 To 10
         xX$ = "RT" + ZFILL$(RT, 2): Call GetRVAR(xX$, X)
         If FOUND Then RTYPE$(RT) = RVAR$(X) + " "
     Next RT
     OLDTYPES$ = "1234F6789ABCDE": Mid$(OLDTYPES$, 6, 1) = Chr$(11)
     Dim OLDTYPE$(14)
     For N = 1 To 14: OLDTYPE$(N) = MID$(OLDTYPES$, N, 1): Next N
     'If PRT.opt > 1 Then
        PRT.SECTION = " ": Call PrintHEAD
        PRT.SECTION = "E": Call PrintTITLE
     'End If
     IDX = 0

1150 ' Check for User Input
     If a <> 27 Then Call ChkForPAUSE
     If a = 27 Then GoTo P2900

1155 ' Process Next Name Record
     IDX = IDX + 1: If IDX > SCNT Then GoTo 2000
     If SelSET(IDX) = 0 Then GoTo 1155
     IDCNT = IDCNT + 1
     Call frmFHSWUPgm.ShowStats("PTRVLDT", 11, Str$(IDCNT))
     CRID = IDX
     RSet xRID$ = fnx$(CRID)
     'COLOR FG(5), BG(5): LOCATE 25, 74, 0: Print XRID$;
     CHG1 = 0: Call FF1GetRec(CRID)

1160 ' Validate Name record Header Info
     If FF1Rec.RTYPE = "1" Then LSet FF1Rec.RTYPE = Chr$(1): CHG1 = -1
     If FF1Rec.RTYPE <> Chr$(1) Then
        SysVAR.NVAR1 = Asc(FF1Rec.RTYPE): Call PtrERR(1, 1)
        LSet FF1Rec.RTYPE = Chr$(1): CHG1 = -1
     End If
     If FF1Rec.RID <> CRID Then
        SysVAR.NVAR1 = FF1Rec.RID: Call PtrERR(1, 2)
        FF1Rec.RID = CRID: CHG1 = -1
     End If
     If FixMissingDD Then
        If (FF1Rec.BY <> 0) And (FF1Rec.DY = 0) And (FF1Rec.DMD = 0) Then
           SysVAR.NVAR1 = CYR - FF1Rec.BY
           If SysVAR.NVAR1 > MAXAGE Then
              Call PtrERR(1, 12)
              FF1Rec.DY = 9999: CHG1 = -1
           End If
        End If
     End If
     MRID = FF1Rec.MID: FRID = FF1Rec.FID
     FSRID = FF1Rec.FCH: MSRID = FF1Rec.MCH
     CHRID = FF1Rec.OLDCH: CHCNT = 0: LASTCH = 0
     SPREC = FF1Rec.SPOUSE: SPCNT = 0: LASTSP = 0
     COMREC = FF1Rec.COM: ADREC = FF1Rec.ADRS
     LSet MISCPTR$ = F1MISCPTR$: BLOC = FF1Rec.BLOC
     If (MRID < 0) Or (MRID > MAXID) Then
        SysVAR.NVAR1 = MRID: Call PtrERR(1, 3)
        FF1Rec.MID = 0: MRID = 0: CHG1 = -1
     End If
     If (FRID < 0) Or (FRID > MAXID) Then
        SysVAR.NVAR1 = FRID: Call PtrERR(1, 4)
        FF1Rec.FID = 0: FRID = 0: CHG1 = -1
     End If
     If (CHRID < 0) Or (CHRID > MAXID) Then
        SysVAR.NVAR1 = CHRID: Call PtrERR(1, 5)
        FF1Rec.OLDCH = 0: CHRID = 0: CHG1 = -1
     End If
     If (FSRID < 0) Or (FSRID > MAXID) Then
        SysVAR.NVAR1 = FSRID: Call PtrERR(1, 6)
        FF1Rec.FCH = 0: FSRID = 0: CHG1 = -1
     End If
     If (MSRID < 0) Or (MSRID > MAXID) Then
        SysVAR.NVAR1 = MSRID: Call PtrERR(1, 7)
        FF1Rec.MCH = 0: MSRID = 0: CHG1 = -1
     End If
     If (FRID = 0) And (FSRID > 0) Then
        SysVAR.NVAR1 = FSRID: Call PtrERR(1, 8)
        FF1Rec.FCH = 0: FSRID = 0: CHG1 = -1
     End If
     If (MRID = 0) And (MSRID > 0) Then
        SysVAR.NVAR1 = MSRID: Call PtrERR(1, 9)
        FF1Rec.MCH = 0: MSRID = 0: CHG1 = -1
     End If
     PTR& = FFRBA&(SPREC, 1)
     If (PTR& > HIREC3&) Then
        SysVAR.NVAR1 = PTR&: Call PtrERR(1, 10)
        FF1Rec.SPOUSE = 0: CHG1 = -1
     End If
     X = InStr(FF1Rec.GIVEN, Chr$(0)): Y = X
     If X > 0 Then
        Call PtrERR(1, 70)
        If UPDTFile Then
           While (X > 0) And (X <= Len(FF1Rec.GIVEN))
                 Mid$(FF1Rec.GIVEN, X, 1) = " "
                 X = InStr(X + 1, FF1Rec.GIVEN, Chr$(0))
           Wend
        End If
        CHG1 = -1
     End If
     X = InStr(FF1Rec.SURNM, Chr$(0))
     If X > 0 Then
        If Y = 0 Then Call PtrERR(1, 70)
        If UPDTFile Then
           While (X > 0) And (X <= Len(FF1Rec.SURNM))
                 Mid$(FF1Rec.SURNM, X, 1) = " "
                 X = X + 1
           Wend
        End If
        CHG1 = -1
     End If
     If (UPDTFile And CHG1) Then
        Call FF1PutREC: CHG1 = 0
     End If

1180 ' Check FID
     If FRID = 0 Then GoTo P1190
     If FRID = MRID Then
        SysVAR.NVAR1 = FF1Rec.FID: Call PtrERR(1, 11)
     End If
     Call FF1GetRec(FRID)
     PTRCK1(CRID) = PTRCK1(CRID) Or 1

P1190: ' Check MID
     If MRID = 0 Then GoTo P1200
     Call FF1GetRec(MRID)
     PTRCK1(CRID) = PTRCK1(CRID) Or 4

P1200: ' Check Child Chain
     If a = 27 Then GoTo P2900
     If CHRID = 0 Then GoTo P1300
     If (CHRID < 0) Or (CHRID > MAXID) Then
        SysVAR.NVAR1 = CHCNT: SysVAR.NVAR2 = LASTCH: Call PtrERR(1, 14)
        GoTo 1240
     End If
     Call FF1GetRec(CHRID): CHCNT = CHCNT + 1
     If CRID = FF1Rec.FID Then
        LASTCH = CHRID: CHRID = FF1Rec.FCH
        PTRCK1(LASTCH) = PTRCK1(LASTCH) Or 2
        If (CHRID > 0) And (CHRID <= MAXID) Then
           If ((PTRCK1(CHRID) And 2) = 2) Then GoTo 1229
        End If
        GoTo P1200
     End If
     If CRID = FF1Rec.MID Then
        LASTCH = CHRID: CHRID = FF1Rec.MCH
        PTRCK1(LASTCH) = PTRCK1(LASTCH) Or 8
        If (CHRID > 0) And (CHRID <= MAXID) Then
           If ((PTRCK1(CHRID) And 8) = 8) Then GoTo 1229
        End If
        GoTo P1200
     End If
     SysVAR.NVAR1 = CHCNT: SysVAR.NVAR2 = CHRID
     SysVAR.NVAR3 = FF1Rec.MID: SysVAR.NVAR4 = FF1Rec.FID
     SysVAR.NVAR5 = LASTCH: Call PtrERR(1, 15)
     GoTo 1240
1229 SysVAR.NVAR1 = CHCNT: SysVAR.NVAR2 = LASTCH
     Call PtrERR(1, 16)

1240 ' Update OLDCH/LASTCH PTRs
     If UPDTFile Then
        If LASTCH = 0 Then
           Call FF1GetRec(CRID): FF1Rec.OLDCH = 0: Call FF1PutREC
           Else
           Call FF1GetRec(LASTCH)
           If FF1Rec.MID = CRID Then FF1Rec.MCH = 0
           If FF1Rec.FID = CRID Then FF1Rec.FCH = 0
           Call FF1PutREC
        End If
     End If

P1300: ' Check Birth/Death dates and Location Record
     If a = 27 Then GoTo P2900
     Call FF1GetRec(CRID)
     If FF1Rec.BY < 0 Then
        SysVAR.NVAR1 = FF1Rec.BY: Call PtrERR(1, 17)
        FF1Rec.BY = 0: CHG1 = -1
     End If
     If FF1Rec.DY < 0 Then
        SysVAR.NVAR2 = FF1Rec.DY: Call PtrERR(1, 18)
        FF1Rec.DY = 0: CHG1 = -1
     End If
     If (UPDTFile And CHG1) Then Call FF1PutREC: CHG1 = 0
P1310: ' Check Birth/Death Place Name Record(s)
     If BLOC = 0 Then GoTo 1400
     PTR& = FFRBA&(BLOC, 1): SysVAR.NVAR1 = PTR&
     If (PTR& > HIREC3&) Then
        Call PtrERR(1, 19): GoTo P1350
     End If
     Call PtrCHK3("GET", PTR&, PTRCK3)
     If PTRCK3 <> 0 Then
        SysVAR.NVAR2 = PTRCK3: Call PtrERR(7, 20)
        GoTo P1350
     End If
     CHG3 = 0: Call FF3GetRec(BLOC)
     If FF3PFX.RTYPE = OLDTYPE$(4) Then
        FF3PFX.RTYPE = Chr$(4): CHG3 = -1
     End If
     If (FF3PFX.RTYPE <> Chr$(4)) And (FF3PFX.RTYPE <> Chr$(12)) Then
        SysVAR.NVAR2 = Asc(FF3PFX.RTYPE): Call PtrERR(6, 21)
        GoTo P1350
     End If
     If BLOC <> FF3PFX.RNO Then
        SysVAR.NVAR2 = FFRBA&(FF3PFX.RNO, 1): Call PtrERR(6, 23)
        FF3PFX.RNO = BLOC: CHG3 = -1
     End If
     If FF3PFX.RTYPE = Chr$(4) Then
        If FF3PLACE.SRNO <> CRID Then
           SysVAR.NVAR2 = FFRBA&(FF3PLACE.SRNO, 1): Call PtrERR(6, 22)
           GoTo P1350
        End If
        If FF3PLACE.SRTYPE = OLDTYPE$(1) Then
           FF3PLACE.SRTYPE = Chr$(1): CHG3 = -1
        End If
        If FF3PLACE.SRTYPE <> Chr$(1) Then
           SysVAR.NVAR2 = Asc(FF3PLACE.SRTYPE): Call PtrERR(6, 24)
           FF3PLACE.SRTYPE = Chr$(1): CHG3 = -1
        End If
        If (UPDTFile And CHG3) Then Call FF3PutRec: CHG3 = 0
        PTRCK3 = 4 * 256 + 64 + 1
        Call PtrCHK3("PUT", PTR&, PTRCK3)
        GoTo 1400
     End If
     ' Check Long Place Name Record(S)
     If FF3PFX.RTYPE = Chr$(12) Then
        If FF3PLACE2.SRNO <> CRID Then
           SysVAR.NVAR2 = FFRBA&(FF3PLACE2.SRNO, 1): Call PtrERR(6, 22)
           GoTo P1350
        End If
        If FF3PLACE2.SRTYPE <> Chr$(1) Then
           SysVAR.NVAR2 = Asc(FF3PLACE2.SRTYPE): Call PtrERR(6, 24)
           FF3PLACE2.SRTYPE = Chr$(1): CHG3 = -1
        End If
        If (UPDTFile And CHG3) Then Call FF3PutRec: CHG3 = 0
        PTRCK3 = 12 * 256 + 64 + 1
        Call PtrCHK3("PUT", PTR&, PTRCK3)
        BLOC = FF3PLACE2.PNEXT
        GoTo P1310
     End If

P1350: ' ZERO Out Birth/Death PLACE Pointer
     If UPDTFile Then
        Call FF1GetRec(CRID): FF1Rec.BLOC = 0: Call FF1PutREC
     End If

1400 ' Check Individual's Comments
     If a = 27 Then GoTo P2900
     If COMREC = 0 Then GoTo 1450
     SRTYPE = 1: SRNO = CRID: COMCNT = 0: GoSub P1900

1450 ' Check Individual's Residence Records
     If a = 27 Then GoTo P2900
     If ADREC = 0 Then GoTo 1500
     ASRTYPE = 1: ASRNO = CRID: ADRCNT = 0: LastADR = 0: GoSub 1800

1500 ' Check Spouse Records
     If a = 27 Then GoTo P2900
     If SPREC = 0 Then GoTo 1600 Else SPNUM = 0
     SysVAR.NVAR1 = SPCNT + 1
     PTR& = FFRBA&(SPREC, 1)
     If (PTR& > HIREC3&) Then
        SysVAR.NVAR2 = PTR&: Call PtrERR(2, 25)
        GoTo 1570
     End If
     CHG3 = 0: SPCNT = SPCNT + 1
     Call FF3GetRec(SPREC)
     If FF3PFX.RTYPE = OLDTYPE$(3) Then
        LSet FF3PFX.RTYPE = Chr$(3): CHG3 = -1
     End If
     If FF3PFX.RTYPE <> Chr$(3) Then
        SysVAR.NVAR2 = Asc(FF3PFX.RTYPE): Call PtrERR(2, 26)
        GoTo 1570
     End If
     If SPREC <> FF3PFX.RNO Then
        SysVAR.NVAR2 = FFRBA&(FF3PFX.RNO, 1): Call PtrERR(2, 27)
        FF3PFX.RNO = SPREC: CHG3 = -1
     End If
     If (UPDTFile And CHG3) Then Call FF3PutRec: CHG3 = 0

1520 ' Check Marriage/Divorce dates and Location Record
     If a = 27 Then GoTo P2900
     Call FF3GetRec(SPREC)
     If FF3SPOUSE.D1Y < 0 Then
        SysVAR.NVAR2 = FF3SPOUSE.D1Y: Call PtrERR(2, 28)
        FF3SPOUSE.D1Y = 0: CHG3 = -1
     End If
     If FF3SPOUSE.D2Y < 0 Then
        SysVAR.NVAR2 = FF3SPOUSE.D2Y: Call PtrERR(2, 29)
        FF3SPOUSE.D2Y = 0: CHG3 = -1
     End If
     If (UPDTFile And CHG3) Then Call FF3PutRec
     CHG3 = 0: MLOC = FF3SPOUSE.MLOC
P1525: ' Check Place Name Record
     If MLOC = 0 Then GoTo 1550
     PTR& = FFRBA&(MLOC, 1)
     If (PTR& > HIREC3&) Then
        SysVAR.NVAR2 = PTR&: Call PtrERR(2, 30)
        GoTo 1535
     End If
     CHG3 = 0: Call FF3GetRec(MLOC)
     If (FF3PFX.RTYPE <> Chr$(4)) And (FF3PFX.RTYPE <> Chr$(12)) Then
        SysVAR.NVAR2 = Asc(FF3PFX.RTYPE): Call PtrERR(6, 31)
        GoTo 1535
     End If
     If MLOC <> FF3PFX.RNO Then
        SysVAR.NVAR2 = PTR&: SysVAR.NVAR3 = FFRBA&(FF3PFX.RNO, 1)
        Call PtrERR(6, 32)
        FF3PFX.RNO = MLOC: CHG3 = -1
     End If
     If FF3PFX.RTYPE = OLDTYPE$(4) Then LSet FF3PFX.RTYPE = Chr$(4): CHG3 = -1
     If FF3PFX.RTYPE = Chr$(4) Then
        If FF3PLACE.SRTYPE = OLDTYPE$(3) Then FF3PLACE.SRTYPE = Chr$(3): CHG3 = -1
        If FF3PLACE.SRTYPE <> Chr$(3) Then
           SysVAR.NVAR2 = Asc(FF3PLACE.SRTYPE): Call PtrERR(6, 33)
           FF3PLACE.SRTYPE = Chr$(3): CHG3 = -1
        End If
        If FF3PLACE.SRNO <> SPREC Then
           SysVAR.NVAR2 = FFRBA&(SPREC, 1): SysVAR.NVAR3 = FFRBA&(FF3PLACE.SRNO, 1)
           Call PtrERR(6, 34)
           GoTo 1535
        End If
        If (UPDTFile And CHG3) Then Call FF3PutRec: CHG3 = 0
        PTRCK3 = 4 * 256 + 32 + 1
        Call PtrCHK3("PUT", PTR&, PTRCK3)
        GoTo 1550
     End If
     ' Check Long Place Name Record
     If FF3PFX.RTYPE = Chr$(12) Then
        If FF3PLACE2.SRTYPE <> Chr$(3) Then
           SysVAR.NVAR2 = Asc(FF3PLACE2.SRTYPE): Call PtrERR(6, 33)
           FF3PLACE2.SRTYPE = Chr$(3): CHG3 = -1
        End If
        If FF3PLACE2.SRNO <> SPREC Then
           SysVAR.NVAR2 = FFRBA&(SPREC, 1): SysVAR.NVAR3 = FFRBA&(FF3PLACE2.SRNO, 1)
           Call PtrERR(6, 34)
           GoTo 1535
        End If
        If (UPDTFile And CHG3) Then Call FF3PutRec: CHG3 = 0
        PTRCK3 = 4 * 256 + 32 + 1
        Call PtrCHK3("PUT", PTR&, PTRCK3)
        If FF3PLACE2.PNEXT <> 0 Then
           MLOC = FF3PLACE2.PNEXT: GoTo P1525
        End If
        GoTo 1550
     End If

1535 ' ZERO Out Spouse PLACE Pointer
     If UPDTFile Then
        Call FF3GetRec(SPREC): FF3SPOUSE.MLOC = 0: Call FF3PutRec
     End If

1550 ' Get Next Spouse Record Number
     If a = 27 Then GoTo P2900
     Call FF3GetRec(SPREC): SPNUM = 0
     SysVAR.NVAR2 = FF3SPOUSE.SP1ID: SysVAR.NVAR3 = FF3SPOUSE.SP2ID
     If CRID = FF3SPOUSE.SP1ID Then
        LASTSP = SPREC: SPREC = FF3SPOUSE.SP1NXT: SPNUM = 1
        GoTo 1560
     End If
     If CRID = FF3SPOUSE.SP2ID Then
        LASTSP = SPREC: SPREC = FF3SPOUSE.SP2NXT: SPNUM = 2
        GoTo 1560
     End If
     Call PtrERR(2, 35)
     GoTo 1570
1560 LSP& = FFRBA&(LASTSP, 1)
     Call PtrCHK3("GET", LSP&, PTRCK3)
     If (PTRCK3 And SPNUM) > 0 Then
        Call PtrERR(2, 36)
        GoTo 1570
     End If
     PTRCK3 = (((PTRCK3 Or SPNUM) Or 64) Or 3 * 256)
     Call PtrCHK3("PUT", LSP&, PTRCK3)
     If (PTRCK3 And 15) <> SPNUM Then GoTo 1500 ' Skip COM & ADR chk for 2nd spouse
     COMREC = FF3SPOUSE.COM: ADREC = FF3SPOUSE.LOC
     If COMREC <> 0 Then SRTYPE = 3: SRNO = LASTSP: COMCNT = 0: GoSub P1900
     If ADREC > 0 Then ASRTYPE = 3: ASRNO = LASTSP: ADRCNT = 0: GoSub 1800
     GoTo 1500

1570 ' Terminate Spouse Records with Last one
     If UPDTFile Then
        If LASTSP = 0 Then
           Call FF1GetRec(CRID)
           FF1Rec.SPOUSE = 0: Call FF1PutREC
           Else
           Call FF3GetRec(LASTSP)
           If CRID = FF3SPOUSE.SP1ID Then FF3SPOUSE.SP1NXT = 0
           If CRID = FF3SPOUSE.SP2ID Then FF3SPOUSE.SP2NXT = 0
           Call FF3PutRec
        End If
     End If

1600 ' Check Miscellaneous Info
     For MRTYPE = 6 To 10
         ET = 6
         Select Case MRTYPE
            Case 6: MSCREC = FF1Rec.EVENT: ET = 5
            Case 7: MSCREC = FF1Rec.EDUC
            Case 8: MSCREC = FF1Rec.WORK
            Case 9: MSCREC = FF1Rec.MIL
            Case 10: MSCREC = FF1Rec.HLTH
         End Select
         If MSCREC <> 0 Then
            MSCCNT = 0: LASTMSC = 0: LastADR = 0
            GoSub 1700
         End If
     Next MRTYPE
     If a = 27 Then GoTo P2900
     If UPDTFile And XCHANGED Then
        Call FF1GetRec(CRID): Call FF1PutREC
     End If
     GoTo 1150

1700 ' Check Miscellaneous Information records
     If a = 27 Then Return
     If MSCREC = 0 Then GoTo 1799
     SysVAR.MVAR = RTYPE$(MRTYPE)
     MSCCNT = MSCCNT + 1
     SysVAR.NVAR2 = MSCCNT
     PTR& = FFRBA&(MSCREC, 1)
     If (PTR& > HIREC3&) Then
        SysVAR.NVAR3 = PTR&: Call PtrERR(ET, 37)
        GoTo 1730
     End If
     CHG3 = 0: Call FF3GetRec(MSCREC)
     SysVAR.MVAR = RTYPE$(MRTYPE)
     Call PtrCHK3("GET", PTR&, PTRCK3)
     If PTRCK3 <> 0 Then
        SysVAR.NVAR3 = PTRCK3: Call PtrERR(7, 38)
        GoTo 1730
     End If
     If FF3PFX.RTYPE = OLDTYPE$(MRTYPE) Then
        FF3PFX.RTYPE = Chr$(MRTYPE): CHG3 = -1
     End If
     If FF3PFX.RTYPE <> Chr$(MRTYPE) Then
        SysVAR.NVAR3 = Asc(FF3PFX.RTYPE): Call PtrERR(ET, 39)
        GoTo 1730
     End If
     If MSCREC <> FF3PFX.RNO Then
        SysVAR.NVAR3 = FFRBA&(MSCREC, 1): SysVAR.NVAR4 = FFRBA&(FF3PFX.RNO, 1)
        Call PtrERR(ET, 40)
        GoTo 1730
     End If
     Select Case MRTYPE
        Case 6: xSRTYPE$ = FF3EVENT.SRTYPE: MRSRNO = FF3EVENT.SRNO
        Case Else: xSRTYPE$ = FF3PNTRS.SRTYPE: MRSRNO = FF3PNTRS.SRNO
     End Select
     If xSRTYPE$ = OLDTYPE$(1) Then LSet xSRTYPE$ = Chr$(1): CHG3 = -1
     If xSRTYPE$ <> Chr$(1) Then
        SysVAR.NVAR3 = Asc(xSRTYPE$): Call PtrERR(ET, 41)
        LSet xSRTYPE$ = Chr$(1): CHG3 = -1
     End If
     If MRSRNO <> CRID Then
        SysVAR.NVAR3 = MRSRNO: Call PtrERR(ET, 42)
        GoTo 1730
     End If
     If (UPDTFile And CHG3) Then
        Select Case MRTYPE
           Case 6: FF3EVENT.SRTYPE = xSRTYPE$
           Case Else: FF3PNTRS.SRTYPE = xSRTYPE$
        End Select
        Call FF3PutRec: CHG3 = 0
     End If
     PTRCK3 = MRTYPE * 256 + 64 + 1
     Call PtrCHK3("PUT", PTR&, PTRCK3)
     LASTMSC = MSCREC
     Select Case MRTYPE
        Case 6: MSCREC = FF3EVENT.NXT: COMREC = FF3EVENT.COM: ADREC = 0
                If FF3EVENT.LPLNM <> 0 Then
                   PTRCK3 = 12 * 256 + 64 + 1
                   PTR& = FFRBA&(FF3EVENT.LPLNM, 1)
                   Call PtrCHK3("PUT", PTR&, PTRCK3)
                End If
        Case Else: MSCREC = FF3PNTRS.NXT: COMREC = FF3PNTRS.COM
                   ADREC = FF3PNTRS.LOC
     End Select
     If COMREC <> 0 Then
        SRTYPE = MRTYPE: SRNO = LASTMSC: COMCNT = 0: GoSub P1900
     End If
     If ADREC > 0 Then
        ASRTYPE = MRTYPE: ASRNO = LASTMSC: ADRCNT = 0: GoSub 1800
     End If
     GoTo 1700

1730 ' Stop Miscellaneous Records with Lastmsc
     If UPDTFile Then
        MSCCNT = MSCCNT - 1
        If MSCCNT > 0 Then
           Call FF3GetRec(LASTMSC)
           Select Case MRTYPE
              Case 6: FF3EVENT.NXT = 0
              Case Else: FF3PNTRS.NXT = 0
           End Select
           Call FF3PutRec
           Else
           Call FF1GetRec(CRID)
           Select Case MRTYPE
              Case 6: FF1Rec.EVENT = 0
              Case 7: FF1Rec.EDUC = 0
              Case 8: FF1Rec.WORK = 0
              Case 9: FF1Rec.MIL = 0
              Case 10: FF1Rec.HLTH = 0
           End Select
           Call FF1PutREC
        End If
     End If
1799 Return

1800 ' Check Address Records
     If a = 27 Then Return
     If ADREC = 0 Then GoTo 1899
     ADRCNT = ADRCNT + 1: CHG2 = 0
     SysVAR.MVAR = RTYPE$(ASRTYPE)
     SysVAR.NVAR1 = ADRCNT: SysVAR.NVAR2 = ASRTYPE
     If (ADREC < 0) Or (ADREC > HIREC2) Then
        SysVAR.NVAR3 = ADREC: Call PtrERR(3, 43)
        GoTo 1830
     End If
     Call FF2GetRec(ADREC)
     If PTRCK2(ADREC) <> 0 Then
        SysVAR.NVAR3 = PTRCK2(ADREC): Call PtrERR(7, 44)
        GoTo 1830
     End If
     If FF2PFX.RTYPE = OLDTYPE$(2) Then FF2PFX.RTYPE = Chr$(2): CHG2 = -1
     If FF2PFX.RTYPE <> Chr$(2) Then
        SysVAR.NVAR3 = Asc(FF2PFX.RTYPE): Call PtrERR(3, 45)
        FF2PFX.RTYPE = Chr$(2): CHG2 = -1
     End If
     If ADREC <> FF2PFX.RNO Then
        SysVAR.NVAR3 = ADREC: SysVAR.NVAR4 = FF2PFX.RNO
        Call PtrERR(3, 46)
        FF2PFX.RNO = ADREC: CHG2 = -1
     End If
     If FF2PFX.SRTYPE = OLDTYPE$(ASRTYPE) Then LSet FF2PFX.SRTYPE = Chr$(ASRTYPE): CHG2 = -1
     If FF2PFX.SRTYPE <> Chr$(ASRTYPE) Then
        SysVAR.NVAR3 = Asc(FF2PFX.SRTYPE): Call PtrERR(3, 47)
        FF2PFX.SRTYPE = Chr$(ASRTYPE): CHG2 = -1
     End If
     If FF2PFX.SRNO <> ASRNO Then
        SysVAR.NVAR3 = FFRBA&(ASRNO, 1): SysVAR.NVAR4 = FFRBA&(FF2PFX.SRNO, 1)
        Call PtrERR(3, 48)
        GoTo 1830
     End If
     If (UPDTFile And CHG2) Then Call FF2PutRec: CHG2 = 0
     If ASRTYPE = 1 Then PTRCK = 64 Else PTRCK = 32
     PTRCK2(ADREC) = 2 * 256 + PTRCK + 1
     LastADR = ADREC: ADREC = FF2PFX.NXT
     COMREC = FF2PFX.COM
     If COMREC <> 0 Then
        SRTYPE = 2: SRNO = LastADR
        COMCNT = 0: GoSub P1900
     End If
     GoTo 1800

1830 ' Stop Address Records with Lastadr
     If UPDTFile Then
        ADRCNT = ADRCNT - 1
        If LastADR <> 0 Then
           Call FF2GetRec(LastADR): FF2PFX.NXT = 0: Call FF2PutRec
           Else
           If ASRTYPE = 1 Then
              Call FF1GetRec(CRID): FF1Rec.ADRS = 0: Call FF1PutREC
              Else
              Call FF3GetRec(ASRNO)
              Select Case ASRTYPE
                 Case 3: FF3SPOUSE.LOC = 0
                 Case Else: FF3PNTRS.LOC = 0
              End Select
              Call FF3PutRec
           End If
        End If
     End If
1899 Return

P1900: ' Check Comments
     If a = 27 Then Return
     SysVAR.MVAR = RTYPE$(SRTYPE)
     If COMREC = 0 Then GoTo 1940
     COMCNT = COMCNT + 1
     SysVAR.NVAR1 = COMCNT: SysVAR.NVAR2 = SRTYPE
     PTR& = FFRBA&(COMREC, 1)
     If (PTR& > HIREC3&) Then
        SysVAR.NVAR3 = PTR&: Call PtrERR(4, 49)
        GoTo 1930
     End If
     CHG3 = 0: Call FF3GetRec(COMREC)
     SysVAR.MVAR = RTYPE$(SRTYPE)
     If COMCNT = 1 Then
        COMTOT = FF3COMMENT.CCNT: LASTCOM = 0: FIRSTCOM = COMREC
     End If
     Call PtrCHK3("GET", PTR&, PTRCK3)
     If PTRCK3 <> 0 Then
        SysVAR.NVAR3 = PTRCK3: Call PtrERR(7, 50)
        GoTo 1930
     End If
     If FF3PFX.RTYPE = OLDTYPE$(5) Then FF3PFX.RTYPE = Chr$(5): CHG3 = -1
     If FF3PFX.RTYPE <> Chr$(5) Then
        SysVAR.NVAR3 = Asc(FF3PFX.RTYPE): Call PtrERR(4, 51)
        GoTo 1930
     End If
     If COMREC <> FF3PFX.RNO Then
        SysVAR.NVAR3 = FFRBA&(COMREC, 1): SysVAR.NVAR4 = FFRBA&(FF3PFX.RNO, 1)
        Call PtrERR(4, 52)
        FF3PFX.RNO = COMREC: CHG3 = -1
     End If
     If FF3COMMENT.SRTYPE = OLDTYPE$(SRTYPE) Then
        FF3COMMENT.SRTYPE = Chr$(SRTYPE): CHG3 = -1
     End If
     If FF3COMMENT.SRTYPE <> Chr$(SRTYPE) Then
        SysVAR.NVAR3 = Asc(FF3COMMENT.SRTYPE)
        Call PtrERR(4, 53)
        FF3COMMENT.SRTYPE = Chr$(SRTYPE): CHG3 = -1
     End If
     If FF3COMMENT.SRNO <> SRNO Then
        SysVAR.NVAR3 = FFRBA&(SRNO, 1): SysVAR.NVAR4 = FFRBA&(FF3COMMENT.SRNO, 1)
        Call PtrERR(4, 54)
        GoTo 1930
     End If
     If FF3COMMENT.BPTR <> LASTCOM Then
        SysVAR.NVAR3 = FFRBA&(FF3COMMENT.BPTR, 1)
        SysVAR.NVAR4 = FFRBA&(LASTCOM, 1)
        Call PtrERR(4, 55)
        FF3COMMENT.BPTR = LASTCOM: CHG3 = -1
     End If
     If (UPDTFile And CHG3) Then Call FF3PutRec: CHG3 = 0
     If SRTYPE = 1 Then PTRCK = 64 Else If SRTYPE = 2 Then PTRCK = 16 Else PTRCK = 32
     PTRCK3 = 5 * 256 + PTRCK + 1
     Call PtrCHK3("PUT", PTR&, PTRCK3)
     LASTCOM = COMREC: COMREC = FF3COMMENT.FPTR
     GoTo P1900

1930 ' Stop comments with Lastcom
     COMCNT = COMCNT - 1
     If UPDTFile Then
        If COMCNT >= 1 Then
           Call FF3GetRec(LASTCOM): FF3COMMENT.FPTR = 0: Call FF3PutRec
           Else
           Select Case SRTYPE
              Case 1: Call FF1GetRec(CRID): FF1Rec.COM = 0: Call FF1PutREC
              Case 2: Call FF2GetRec(SRNO): FF2PFX.COM = 0: Call FF2PutRec
              Case 3: Call FF3GetRec(SRNO): FF3SPOUSE.COM = 0: Call FF3PutRec
              Case 6: Call FF3GetRec(SRNO): FF3EVENT.COM = 0: Call FF3PutRec
              Case 11: Call FF3GetRec(SRNO): FF3EVENT.COM = 0: Call FF3PutRec
              Case Else: Call FF3GetRec(SRNO): FF3PNTRS.COM = 0: Call FF3PutRec
           End Select
        End If
     End If

1940 ' Correct Comment Record Count if necessary
     If ((COMCNT < 1) Or (COMCNT = COMTOT)) Then GoTo 1949
     SysVAR.NVAR3 = COMTOT: SysVAR.NVAR4 = COMCNT
     Call PtrERR(4, 56)
     If UPDTFile Then
        Call FF3GetRec(FIRSTCOM): FF3COMMENT.CCNT = COMCNT
        Call FF3PutRec
     End If
1949 Return

2000 ' Check for broken CHILD chains
     If a = 27 Then GoTo P2900
     If Not ALLREC Then GoTo P2500
     Call PutMSG("VLDTM004")
     For CRID = 1 To MAXID
         If PTRCK1(CRID) <> 0 Then
            X = PTRCK1(CRID) And 3: Y = PTRCK1(CRID) \ 4
            If (X = 1) Or (Y = 1) Then
               Call FF1GetRec(CRID)
               If X = 1 Then
                  SysVAR.NVAR1 = FF1Rec.FID: Call PtrERR(1, 57)
               End If
               If Y = 1 Then
                  SysVAR.NVAR1 = FF1Rec.MID: Call PtrERR(1, 58)
               End If
               If UPDTFile Then
                  If X = 1 Then FF1Rec.FID = 0: FF1Rec.FCH = 0
                  If Y = 1 Then FF1Rec.MID = 0: FF1Rec.MCH = 0
                  Call FF1PutREC
               End If
            End If
         End If
     Next CRID

2100 ' Check ADDRESS File FREE Chain
     If a = 27 Then GoTo P2900
     Call PutMSG("VLDTM005")
     FREEBGN = FF2Hdr.FREEBGN: FREECNT = FF2Hdr.FREECNT
     FCNT = 0: FREC = FREEBGN: LASTFREE = 0: CRID = 0
2120 If FREC = 0 Then GoTo 2150
     FCNT = FCNT + 1: SysVAR.NVAR1 = FCNT
     If (FREC < 0) Or (FREC > HIREC2) Then
        SysVAR.NVAR2 = FREC: Call PtrERR(9, 59)
        GoTo 2140
     End If
     If PTRCK2(FREC) > 0 Then
        SysVAR.NVAR2 = PTRCK2(FREC): Call PtrERR(7, 60)
        GoTo 2140
     End If
     PTRCK2(FREC) = 129
     LASTFREE = FREC
     Call FF2GetRec(FREC)
     If UPDTFile Then
        If (FF2PFX.RTYPE <> Chr$(255)) Or (FF2PFX.RNO <> FREC) Then
           FF2PFX.RNO = FREC
           FF2PFX.RTYPE = Chr$(255): Call FF2PutRec
        End If
     End If
     FREC = FF2PFX.NXT
     GoTo 2120

2140 ' Terminate FREE Chain with LastFree
     FCNT = FCNT - 1
     If LASTFREE = 0 Then
        FF2Hdr.FREEBGN = 0
        Else
        If UPDTFile Then
           Call FF2GetRec(LASTFREE): FF2PFX.NXT = 0: Call FF2PutRec
        End If
     End If

2150 ' Check ADDRESS File FREE REC Count
     If FREECNT = FCNT Then GoTo 2160
     SysVAR.NVAR2 = FREECNT: Call PtrERR(9, 61)
     FF2Hdr.FREECNT = FCNT
     If UPDTFile Then
        FFCHANGED = -1
        FF2Hdr.REVLVL = FF2Hdr.REVLVL Or 256
     End If

2160 ' Check for Isolated ADDRESS Records
     If a = 27 Then GoTo P2900
     Call PutMSG("VLDTM006")
     ISOCNT = 0
     For id = 1 To HIREC2
         If PTRCK2(id) = 0 Then
            Call FF2GetRec(id)
            SysVAR.NVAR1 = id: SysVAR.NVAR2 = Asc(FF2PFX.RTYPE)
            SysVAR.NVAR3 = FF2PFX.SRNO: SysVAR.NVAR4 = Asc(FF2PFX.SRTYPE)
            Call PtrERR(8, 62)
            ISOCNT = ISOCNT + 1
         End If
     Next id
     If ISOCNT = 0 Then GoTo 2200

2170 ' Rebuild ADDRESS File FREE Chain
     If a = 27 Then GoTo P2900
     SysVAR.NVAR1 = ISOCNT: Call PtrERR(0, 63)
     If UPDTFile Then
        FREEBGN = 0: FREECNT = 0: LASTFREE = 0
        For id = 1 To HIREC2
            If ((PTRCK2(id) < 1) Or (PTRCK2(id) = 129)) Then
               If FREEBGN = 0 Then FREEBGN = id: LASTFREE = id
               Call FF2GetRec(LASTFREE)
               FF2PFX.RTYPE = Chr$(255): FF2PFX.RNO = LASTFREE
               If LASTFREE <> id Then FF2PFX.NXT = id
               Call FF2PutRec
               FREECNT = FREECNT + 1: LASTFREE = id
            End If
        Next id
        If (FF2Hdr.FREEBGN <> FREEBGN) Or (FF2Hdr.FREECNT <> FREECNT) Then
           FF2Hdr.FREEBGN = FREEBGN: FF2Hdr.FREECNT = FREECNT
           FFCHANGED = -1: FF2Hdr.REVLVL = FF2Hdr.REVLVL Or 256
        End If
        If LASTFREE <> 0 Then
           Call FF2GetRec(LASTFREE)
           FF2PFX.RTYPE = Chr$(255)
           FF2PFX.NXT = 0: Call FF2PutRec
        End If
     End If

2200 ' Check MISCELLANEOUS File FREE Chain
     If a = 27 Then GoTo P2900
     Call PutMSG("VLDTM007")
     FREEBGN = FF3HDR.FREEBGN
     PTR& = FFRBA&(FF3HDR.FREEBGN, 1)
     If PTR& > HIREC3& Then FREEBGN = 0
     zFREECNT& = FFRBA&(FF3HDR.FREECNT, 1)
     zFCNT& = 0: FREC = FREEBGN: LASTFREE = 0
2210 If FREC = 0 Then GoTo 2250
     zFCNT& = zFCNT& + 1: SysVAR.NVAR1 = zFCNT&
     PTR& = FFRBA&(FREC, 1)
     If (PTR& > HIREC3&) Then
        SysVAR.NVAR2 = PTR&: Call PtrERR(9, 64)
        GoTo 2240
     End If
     Call PtrCHK3("GET", PTR&, PTRCK3)
     If PTRCK3 <> 0 Then
        SysVAR.NVAR2 = PTRCK3: Call PtrERR(7, 65)
        GoTo 2240
     End If
     Call PtrCHK3("PUT", PTR&, 129)
     LASTFREE = FREC: Call FF3GetRec(FREC)
     If UPDTFile Then
        If (FF3PFX.RTYPE <> Chr$(255)) Or (FF3PFX.RNO <> FREC) Then
           FF3PFX.RNO = FREC
           FF3PFX.RTYPE = Chr$(255): Call FF3PutRec
        End If
     End If
     FREC = FF3PNTRS.NXT
     GoTo 2210

2240 ' Terminate FREE Chain with LastFree
     zFCNT& = zFCNT& - 1
     If zFCNT& = 0 Then
        FREEBGN = 0
        Else
        If (LASTFREE <> 0) And UPDTFile Then
           Call FF3GetRec(LASTFREE)
           FF3PFX.RTYPE = Chr$(255): FF3PNTRS.NXT = 0
           Call FF3PutRec
        End If
     End If

2250 ' Check MISCELLANEOUS File FREE REC Count
     If a = 27 Then GoTo P2900
     If zFREECNT& <> zFCNT& Then
        SysVAR.NVAR2 = zFREECNT&: Call PtrERR(9, 66)
     End If
     If UPDTFile Then
        If (FREEBGN <> FF3HDR.FREEBGN) Or (zFREECNT& <> zFCNT&) Then
           FF3HDR.FREEBGN = FREEBGN: FF3HDR.FREECNT = FFRNUM(zFCNT&, 1)
           FFCHANGED = -1: FF3HDR.REVLVL = FF3HDR.REVLVL Or 256
           zFREECNT& = zFCNT&
        End If
     End If

2270 ' Check for Isolated MISC Records
     If a = 27 Then GoTo P2900
     Call PutMSG("VLDTM008")
     zISOCNT& = 0: zID& = 0
     While zID& < HIREC3&: zID& = zID& + 1
         XID = FFRNUM(zID&, 1): Call FF3GetRec(XID)
         Call PtrCHK3("GET", zID&, PTRCK3)
         If (PTRCK3 <> 0) Then
            If ((PTRCK3 \ 256) = 3) Then
               If ((PTRCK3 And 15) <> 3) Then
                  If (FF3SPOUSE.SP1ID <> 0) And (FF3SPOUSE.SP2ID <> 0) Then
                     SysVAR.NVAR1 = FF3SPOUSE.SP1ID
                     SysVAR.NVAR2 = FF3SPOUSE.SP2ID
                     SysVAR.NVAR3 = (PTRCK3 And 15)
                     Call PtrERR(2, 67)
                  End If
               End If
            End If
            Else
            RT = Asc(FF3PFX.RTYPE)
            Select Case RT
               Case 3: SRT = Asc(FF3SPOUSE.SRTYPE): SRN = FF3SPOUSE.SP1ID
                       If SRN = 0 Then SRN = FF3SPOUSE.SP2ID
               Case 4: SRT = Asc(FF3PLACE.SRTYPE): SRN = FF3PLACE.SRNO
               Case 5: SRT = Asc(FF3COMMENT.SRTYPE): SRN = FF3COMMENT.SRNO
               Case 6: SRT = Asc(FF3EVENT.SRTYPE): SRN = FF3EVENT.SRNO
               Case 11: SRT = Asc(FF3EVENT.SRTYPE): SRN = FF3EVENT.SRNO
               Case 12: SRT = Asc(FF3PLACE2.SRTYPE): SRN = FF3PLACE2.SRNO
               Case Else: SRT = Asc(FF3PNTRS.SRTYPE): SRN = FF3PNTRS.SRNO
            End Select
            SysVAR.NVAR1 = zID&: SysVAR.NVAR2 = RT
            SysVAR.NVAR3 = SRN: SysVAR.NVAR4 = SRT
            Call PtrERR(8, 68)
            zISOCNT& = zISOCNT& + 1
         End If
     Wend
     If a = 27 Then GoTo P2900
     If zISOCNT& = 0 Then GoTo P2500

2280 ' Rebuild MISC File FREE Chain
     SysVAR.NVAR1 = zISOCNT&: Call PtrERR(0, 69)
     If UPDTFile Then
        FREEBGN = 0: zFCNT& = 0: LASTFREE = 0: zID& = 0
        While zID& < HIREC3&: zID& = zID& + 1
            XID = FFRNUM(zID&, 1)
            Call PtrCHK3("GET", zID&, PTRCK3)
            If ((PTRCK3 = 0) Or (PTRCK3 = 129)) Then
               If FREEBGN = 0 Then FREEBGN = XID
               If LASTFREE <> 0 Then
                  Call FF3GetRec(LASTFREE)
                  FF3PFX.RTYPE = Chr$(255): FF3PFX.RNO = LASTFREE
                  FF3PNTRS.NXT = XID: Call FF3PutRec
               End If
               zFCNT& = zFCNT& + 1: LASTFREE = XID
            End If
        Wend
        If LASTFREE <> 0 Then
           Call FF3GetRec(LASTFREE)
           FF3PFX.RTYPE = Chr$(255): FF3PFX.RNO = LASTFREE
           FF3PNTRS.NXT = 0: Call FF3PutRec
        End If
        FF3HDR.FREEBGN = FREEBGN: FF3HDR.FREECNT = FFRNUM(zFCNT&, 1)
        FFCHANGED = -1: FF3HDR.REVLVL = FF3HDR.REVLVL Or 256
     End If

P2500: ' END OF REPORT
     On Local Error GoTo 0
     If a = 27 Then GoTo P2900
     'If (PRT.opt = 1) Or (ERRCNT(0) = 0) Then GoTo P2950
     LSet PLINE$ = " ": Call WriteLINE
     PRT.SECTION = "T": Call PrintTITLE
     NUMED$ = Space$(5)
     For ET = 1 To 9
        If ERRCNT(ET) > 0 Then
           EMSG$ = "ET" + ZFILL$(ET, 2)
           Call GetRVAR(EMSG$, X)
           If FOUND Then
              RSet NUMED$ = fnx$(ERRCNT(ET))
              LSet PLINE$ = NUMED$ + "  " + RVAR$(X)
              Call PrintLINE: If a = 27 Then ET = 11
           End If
        End If
     Next ET
     Call GetRVAR("ETOT", X)
     If FOUND Then
        RSet NUMED$ = fnx$(ERRCNT(0))
        LSet PLINE$ = NUMED$ + "  " + RVAR$(X)
        Call PrintLINE
     End If
     GoTo P2950

P2900: ' Early Termination
     'If PRT.opt = 1 Then
     '   If CRID > 0 Then
     '      SysVAR.NVAR1 = CRID + 1
     '      Call PutMSG("VLDTM009")
     '      Else
     '      Call PutMSG("VLDTM010")
     '   End If
     '   Else
        Call RptCLOSE
     'End If
     GoTo P2990

P2950: ' Validation Procedure Complete
     If ERID(0) <> 0 Then
        SysVAR.NVAR1 = 0
        For ET = 1 To 6: SysVAR.NVAR1 = SysVAR.NVAR1 + ERRCNT(ET): Next ET
        SysVAR.NVAR2 = FFRBA&(ERID(0), 1)
        Call PutMSG("VLDTM019")
        Call RptCLOSE
        Else
        Call PutMSG("No Errors Found...")
     End If
     'If PRT.opt = 1 Then
     '   Call PutMSG("VLDTM011")
     '   Else
     '   Call RptCLOSE
     'End If
     If UPDTFile Then FFCHANGED = -1

P2990: ' Return
     Call FamCLOSE: Call PtrCHK3("DLET", 0, 0)
     Call frmFHSWUPgm.ShowStats("PTRVLDT", 99, "")
     Call PutMSG("")

End Sub

Rem $STATIC
Sub PtrCHK3(OPT$, X&, B)
   OKAY = -1
   Select Case OPT$
      Case "OPEN": GoSub ptCK3Open
      Case "PUT": GoSub ptCK3Put
      Case "GET": GoSub ptCK3Get
      Case "DLET": GoSub ptCK3Dlet
   End Select
   Exit Sub

ptCK3Dlet: ' Delete Work File
   xX$ = WRKFPFX$ + "FH$PTCK3"
   On Local Error GoTo ptCK3DletX
   Kill xX$
   Return
ptCK3DletX: Resume ptCK3Rtrn

ptCK3Open: ' Open Work File
   GoSub ptCK3Dlet
   On Local Error GoTo ptCK3ERR
   SysVAR.FILENAME = xX$
   Open xX$ For Binary Access Read Write As #6
   Y& = (X& \ 100) + 1: xY$ = String$(200, 0)
   While Y& > 0: Put #6, (200 * (Y& - 1) + 1), xY$: Y& = Y& - 1: Wend
   Return

ptCK3Put: ' Put Number in Work File
   On Local Error GoTo ptCK3ERR
   Put #6, (2 * X& - 1), B
   Return

ptCK3Get: ' Get Number from Work File
   On Local Error GoTo ptCK3ERR
   Get #6, (2 * X& - 1), B
   Return

ptCK3ERR: ' Error Routine
   Call ERRMessage("XXXXM097", ERR, ERL)
   OKAY = 0
   Resume ptCK3Rtrn

ptCK3Rtrn: Return

End Sub

Rem $DYNAMIC
Sub PtrERR(ERRGrp, EMNUM)
   If a <> 27 Then Call ChkForPAUSE
   If a = 27 Then Exit Sub
   ERRCNT(ERRGrp) = ERRCNT(ERRGrp) + 1
   If ERRGrp <> 0 Then
      ERRCNT(12) = ERRCNT(12) + 1
      Call frmFHSWUPgm.ShowStats("PTRVLDT", ERRGrp, Str$(ERRCNT(ERRGrp)))
      Call frmFHSWUPgm.ShowStats("PTRVLDT", 10, Str$(ERRCNT(12)))
   End If
   'If CurFMT <> FMT019 Then CurFMT = FMT019: Fmt = FMTTAB(CurFMT)
   'EFLD$ = "EC" + ZFILL$(ERRGrp, 2)
   'Call FmtFindFLD(EFLD$)
   'If FOUND Then RSet UTXT$(FFLD.UTXT) = FNX$(ERRCNT(ERRGrp))
   'Call FmtFindFLD("ETOT")
   'If FOUND Then
   '   RSet UTXT$(FFLD.UTXT) = FNX$(ERRCNT(12))
   '   Call WinDisplayDATA(0)
   'End If
   If CRID > 0 Then
      ERID(CRID) = ERID(CRID) + 1
      RSet xRID$ = fnx$(CRID)
      Else
      LSet xRID$ = "-----"
   End If
   If CRID <> LastRID Then
      LastRID = CRID
      If CRID <> 0 Then
         ERID(0) = ERID(0) + 1
         Call frmFHSWUPgm.ShowStats("PTRVLDT", 12, Str$(ERID(0)))
      End If
      LSet PLINE$ = xRID$
      Else
      LSet PLINE$ = ""
   End If
   EMNM$ = "EM" + ZFILL$(EMNUM, 2)
   Call GetRVAR(EMNM$, X)
   If FOUND Then
      xERRMSG$ = RVAR$(X)
      Else
      xERRMSG$ = "Error Message " + EMNM$ + " Not Found"
   End If
   Mid$(PLINE$, 8, 255) = xERRMSG$
   Call HTFExpand(PLINE$, 1)
   'COLOR FG(5), BG(5): LOCATE 25, 1, 0: Print XRID$;
   'COLOR FG(1), BG(1): LOCATE 25, 7, 0: Print MID$(PLINE$, 8, 70); Tab(79);
   'If PRT.opt <> 1 Then
      Call PrintLINE
   '   Else
   '   LSet PLINE$ = ""
   '   If Not NOPAUSE Then
   '      Call ErrBEEP(0): Call GetKEY
   '      NOPAUSE = ((Len(CA$) = 2) And (A = 81))
   '   End If
   'End If
End Sub

Rem $STATIC
Function TextDATA$(xX$)
    Select Case (ENV.SortOPTS And 3)
       Case 0: TextDATA$ = xX$
       Case 1: TextDATA$ = UCX$(xX$)
       Case Else: TD$ = Space$(4): Call ComputeSNDX(xX$, TD$)
                  TextDATA$ = TD$ + xX$
    End Select
End Function

Sub ValiDATE(DTMD, DTYR)

9200 ' Validate Dates
     If (DTMD <> 0) Or (DTYR <> 0) Then
        DTMM = DTMD \ 100: DTDD = DTMD Mod 100
        CYR = Val(Right$(DATE$, 4))
        If ((DTYR < 9999) And (DTYR > CYR)) Or (DTMM > 12) Or (DTDD > 31) Then
           OKAY = 0
           Else
           If DTMM = 1 Or DTMM = 3 Or DTMM = 5 Or DTMM = 7 Or DTMM = 8 Or DTMM = 10 Or DTMM = 12 Then
              OKAY = -1
              Else
              If DTDD > 30 Then
                 OKAY = 0
                 Else
                 If DTMM <> 2 Then
                    OKAY = -1
                    Else
                    LPYR = ((DTYR Mod 4) = 0) And (((DTYR Mod 400) = 0) Or ((DTYR Mod 100) > 0))
                    OKAY = (DTDD < (29 - LPYR))
                 End If
              End If
           End If
        End If
     End If
     If Not OKAY Then
        LSet SysVAR.MVAR = XSCRNDate$(DTMD, DTYR)
     End If
End Sub

