Attribute VB_Name = "modFHSWFile"
DefInt A-Z
Private Sub PrintReport()

600 ' Process PRINT Request
    'RDFReport.NAME = "FHS File SETUP Listing"
    'HTFMAX = 9: REDIM HTFTXT$(HTFMAX)
    'HTFTXT$(1) = "H01<Family History System>"
    'HTFTXT$(2) = "H01<[RPTNAME]>"
    'HTFTXT$(3) = "H01"
    'HTFTXT$(4) = "H01File Definition File: [FILENAME]"
    'HTFTXT$(5) = "TD2System File Dataset Names:"
    'HTFTXT$(6) = "TD2 FileCode   GRP    Dataset         Description"
    'HTFTXT$(7) = "TF2Family File Dataset Names:"
    'HTFTXT$(8) = "TF2 FileCode   Drive    Dataset         Description"
    'HTFTXT$(9) = "F02<Page [PAGE]>"
    Call RptOPEN("FILE"): If (Not OKAY) Then GoTo 625
    Call PrintOPEN
610 PRT.SECTION = "0"
    ReDim RPTLIT$(4)
    Call GetRVAR("FAMFILE", x): RPTLIT$(1) = RVAR$(x)
    Call GetRVAR("WORKFILE", x): RPTLIT$(2) = RVAR$(x)
    Call GetRVAR("TRANSFER", x): RPTLIT$(3) = RVAR$(x)
    Call GetRVAR("REPORT", x): RPTLIT$(4) = RVAR$(x)
    PRT.SECTION = "M"
    SysVAR.FILENAME = CONFIG.FDF + ".FDF"
    SysVAR.SETUP = FSETUPS(FDFHDr.DEFAULT).NAME
    SysVAR.SUBJECT = FSETUPS(FDFHDr.DEFAULT).DESC
    Call WriteLINE
    PRT.SECTION = "D"
    Call PrintTITLE
    For x = (FDFHDr.FAMDSN + 1) To ENV.DSNCNT
        Mid$(PLINE$, 2, 8) = FType$(x)
        Mid$(PLINE$, 14, 1) = FGrp$(x)
        Mid$(PLINE$, 18, 12) = FT$(x)
        'Mid$(PLINE$, 34, 255) = MID$(DSDATA$(x), 22, 255)
        Call PrintLINE
    Next x
    Y = FSUTABLE.FT
    While Y <> 0
        PRT.SECTION = "S"
        SysVAR.SETUP = MID$(FSETUP$(Y), 1, 8)
        SysVAR.SUBJECT = MID$(FSETUP$(Y), 9, 255)
        Call WriteLINE: Call PrintTITLE
        Mid$(PLINE$, 8, 255) = RPTLIT$(1)
        Mid$(PLINE$, 26, 255) = FSETUPS(Y).FAMDRV
        Mid$(PLINE$, 34, 255) = FSETUPS(Y).FAMDIR
        Call PrintLINE
        Mid$(PLINE$, 8, 255) = RPTLIT$(2)
        Mid$(PLINE$, 26, 255) = FSETUPS(Y).WRKDRV
        Mid$(PLINE$, 34, 255) = FSETUPS(Y).WRKDIR
        Call PrintLINE
        Mid$(PLINE$, 8, 255) = RPTLIT$(3)
        Mid$(PLINE$, 26, 255) = FSETUPS(Y).TRANDRV
        Mid$(PLINE$, 34, 255) = FSETUPS(Y).TRANDIR
        Call PrintLINE
        Mid$(PLINE$, 8, 255) = RPTLIT$(4)
        Mid$(PLINE$, 26, 255) = FSETUPS(Y).RPTDRV
        Mid$(PLINE$, 34, 255) = FSETUPS(Y).RPTDIR
        Call PrintLINE
        PRT.SECTION = "F"
        Call WriteLINE: Call PrintTITLE
        W = FDFHDr.FAMDSN * (Y - 1)
        For Z = 1 To FDFHDr.FAMDSN
            Mid$(PLINE$, 10, 8) = FType$(Z)
            Select Case Z
               Case 1: Mid$(PLINE$, 26, 1) = FSETUPS(Y).NFDRV
               Case 2: Mid$(PLINE$, 26, 1) = FSETUPS(Y).AFDRV
               Case 3: Mid$(PLINE$, 26, 1) = FSETUPS(Y).MFDRV
               Case 4: Mid$(PLINE$, 26, 1) = FSETUPS(Y).IFDRV
            End Select
            Mid$(PLINE$, 34, 12) = FFNAME$(W + Z)
            'MID$(PLINE$, 50, 255) = MID$(DSDATA$(Z), 22, 255)
            Call PrintLINE
        Next Z
        Y = FSUCHN(Y).FWD
    Wend
    Call PageBREAK
625 ' End of Print
    Call RptCLOSE: If PRT.AGAIN Then GoTo 610
   
End Sub

Public Sub RptFile(OPT$)
   Rem $INCLUDE: 'FHSCOMON.BAS'
   PN$ = " FAMILY HISTORY SYSTEM - FHSWFILE - File Table Report program"
   CY$ = " (C) Copyright 1995-1998 by Phillip E. Brown"
   
   FILENAME$ = Space$(12): HSETUP$ = Space$(46)
   CODE$ = Space$(26): XCODE$ = Space$(124): xname$ = Space$(8)

   Select Case OPT$
      Case "INIT": GoTo 40
      Case "PRINT": GoTo 1000
   End Select
   
40 ' INIT For Report
   Call PgmPREP("FHSFILE")
   Load frmFHSWRPgm
   Call RptLOAD("FILE")
   frmFHSWRPgm.RPSetup.AddItem "FILE"
   frmFHSWRPgm.RPSetup_Click
   frmFHSWRPgm.Show
   Exit Sub

1000 'Print Reports
   Call PrintReport

End Sub
