VERSION 4.00
Begin VB.Form frmFHSWMisc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FHSWMisc"
   ClientHeight    =   7410
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   6810
   Height          =   7815
   Left            =   1080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   Top             =   1170
   Width           =   6930
   Begin VB.Frame MRData 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   2
      Left            =   120
      TabIndex        =   36
      Top             =   4200
      Visible         =   0   'False
      Width           =   4695
      Begin VB.TextBox MRValue 
         Height          =   285
         Index           =   5
         Left            =   960
         TabIndex        =   6
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label MRLabel 
         Caption         =   "Title:"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   37
         Top             =   270
         Width           =   615
      End
   End
   Begin VB.Frame MRData 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1095
      Index           =   3
      Left            =   120
      TabIndex        =   33
      Top             =   5040
      Visible         =   0   'False
      Width           =   4695
      Begin VB.TextBox MRValue 
         Height          =   285
         Index           =   7
         Left            =   960
         TabIndex        =   8
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox MRValue 
         Height          =   285
         Index           =   6
         Left            =   960
         TabIndex        =   7
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label MRLabel 
         Caption         =   "Status:"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   35
         Top             =   600
         Width           =   615
      End
      Begin VB.Label MRLabel 
         Caption         =   "Service:"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   34
         Top             =   270
         Width           =   615
      End
   End
   Begin VB.Frame MRData 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   1
      Left            =   120
      TabIndex        =   30
      Top             =   2640
      Visible         =   0   'False
      Width           =   4695
      Begin VB.TextBox MRValue 
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox MRValue 
         Height          =   285
         Index           =   2
         Left            =   3000
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox MRValue 
         Height          =   285
         Index           =   3
         Left            =   1560
         TabIndex        =   4
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox MRValue 
         Height          =   285
         Index           =   4
         Left            =   1560
         TabIndex        =   5
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label MRLabel 
         Caption         =   "ED Level:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   39
         Top             =   270
         Width           =   855
      End
      Begin VB.Label MRLabel 
         Caption         =   "Degree:"
         Height          =   255
         Index           =   2
         Left            =   2280
         TabIndex        =   38
         Top             =   270
         Width           =   615
      End
      Begin VB.Label MRLabel 
         Caption         =   "Major Subject 1:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   32
         Top             =   630
         Width           =   1215
      End
      Begin VB.Label MRLabel 
         Caption         =   "Major Subject 2:"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   31
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.Frame MRData 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1095
      Index           =   4
      Left            =   120
      TabIndex        =   27
      Top             =   6240
      Visible         =   0   'False
      Width           =   4695
      Begin VB.TextBox MRValue 
         Height          =   285
         Index           =   8
         Left            =   1200
         TabIndex        =   9
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox MRValue 
         Height          =   285
         Index           =   9
         Left            =   1200
         TabIndex        =   10
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label MRLabel 
         Caption         =   "Diagnosis:"
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   29
         Top             =   270
         Width           =   855
      End
      Begin VB.Label MRLabel 
         Caption         =   "Status:"
         Height          =   255
         Index           =   9
         Left            =   600
         TabIndex        =   28
         Top             =   630
         Width           =   615
      End
   End
   Begin VB.Frame DATAFrame 
      Height          =   2535
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   6735
      Begin VB.TextBox DTValue 
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   0
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox DTValue 
         Height          =   285
         Index           =   1
         Left            =   3120
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin VB.Frame NAVFrame 
         Height          =   975
         Left            =   4920
         TabIndex        =   21
         Top             =   120
         Width           =   1695
         Begin VB.CommandButton NAVButton 
            Caption         =   ">>"
            Height          =   315
            Index           =   3
            Left            =   1275
            TabIndex        =   18
            Top             =   240
            Width           =   315
         End
         Begin VB.CommandButton NAVButton 
            Caption         =   "<<"
            Height          =   315
            Index           =   0
            Left            =   100
            TabIndex        =   15
            Top             =   240
            Width           =   350
         End
         Begin VB.CommandButton NAVButton 
            Caption         =   "<-"
            Height          =   315
            Index           =   1
            Left            =   490
            TabIndex        =   16
            Top             =   240
            Width           =   350
         End
         Begin VB.CommandButton NAVButton 
            Caption         =   "->"
            Height          =   315
            Index           =   2
            Left            =   880
            TabIndex        =   17
            Top             =   240
            Width           =   350
         End
         Begin VB.TextBox NAValue 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   320
            MultiLine       =   -1  'True
            TabIndex        =   23
            Top             =   645
            Width           =   375
         End
         Begin VB.TextBox NAValue 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   1000
            MultiLine       =   -1  'True
            TabIndex        =   22
            Text            =   "FHSWMISC.frx":0000
            Top             =   645
            Width           =   375
         End
         Begin VB.Label NAVLabel 
            Alignment       =   2  'Center
            Caption         =   "of"
            Height          =   255
            Left            =   650
            TabIndex        =   24
            Top             =   645
            Width           =   375
         End
      End
      Begin VB.Frame CMDFrame 
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   1920
         Width           =   6375
         Begin VB.CommandButton CMDButton 
            Caption         =   "Notes"
            Height          =   280
            Index           =   2
            Left            =   2640
            TabIndex        =   13
            Top             =   155
            Width           =   1095
         End
         Begin VB.CommandButton CMDButton 
            Cancel          =   -1  'True
            Caption         =   "Return"
            Default         =   -1  'True
            Height          =   280
            Index           =   3
            Left            =   5280
            TabIndex        =   14
            Top             =   150
            Width           =   975
         End
         Begin VB.CommandButton CMDButton 
            Caption         =   "Address"
            Height          =   280
            Index           =   1
            Left            =   1320
            TabIndex        =   12
            Top             =   155
            Width           =   1215
         End
         Begin VB.CommandButton CMDButton 
            Caption         =   "Update"
            Height          =   280
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   155
            Width           =   1095
         End
      End
      Begin VB.Label DTLabel 
         Caption         =   "Date1:"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   26
         Top             =   270
         Width           =   495
      End
      Begin VB.Label DTLabel 
         Caption         =   "Date2:"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   25
         Top             =   270
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmFHSWMisc"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
DefInt A-Z
Dim MDRTYPE, MSRNO, MDRRNO, MDRBGN
Dim MDRHIX, MDRMIX, MDRIX, MD1MD, MD1Y, NEWMDR
Dim MDRCHN()
Dim LastFLD, ZERODATE$
Dim HOLDITEM As Variant, xFT, HRValue() As String

Private Sub AddMDRCHN(X)
    MDRRNO = X
    ReDim MDRCHN(MDRMIX)
    If MDRRNO = 0 Then Exit Sub
    Call FF3GetRec(MDRRNO)
    ED1Y = FF3PNTRS.D1Y: ED1MD = FF3PNTRS.D1MD
    If MDRBGN = 0 Then
       MDRBGN = MDRRNO: MDRIX = 1: MDRCHN(1) = MDRRNO: MDRHIX = 1
       Exit Sub
    End If
    MDRHIX = 0: Z = MDRBGN: MDRIX = 0
    While Z <> 0  ' Build chain, inserting MDRRNO where it belongs
       Call FF3GetRec(Z)
       If OKAY Then
          MDRHIX = MDRHIX + 1
          If Not (MDRHIX < MDRMIX) Then
             MDRMIX = MDRMIX + 10
             ReDim Preserve MDRCHN(MDRMIX)
          End If
          If (MDRIX > 0) Or ((FF3PNTRS.D1Y > ED1Y) Or ((FF3PNTRS.D1Y = ED1Y) And (FF3PNTRS.D1MD > ED1MD))) Then
             MDRCHN(MDRHIX) = Z
             Else
             MDRCHN(MDRHIX) = MDRRNO: MDRIX = MDRHIX
             MDRHIX = MDRHIX + 1
             MDRCHN(MDRHIX) = Z
          End If
          Z = FF3PNTRS.NXT
          Else
          Z = 0
       End If
    Wend
    If MDRIX = 0 Then
       If MDRCHN(MDRHIX) = MDRRNO Then
          MDRIX = MDRHIX
          Else
          MDRHIX = MDRHIX + 1: MDRIX = MDRHIX
          MDRCHN(MDRHIX) = MDRRNO
       End If
    End If
    If MDRIX > 1 Then
       Call FF3GetRec(MDRCHN(MDRIX - 1))
       If OKAY Then
          FF3PNTRS.NXT = MDRRNO
          Call FF3PutRec
       End If
    End If
    Call FF3GetRec(MDRRNO)
    If OKAY Then
       If FF3PNTRS.NXT <> MDRCHN(MDRIX + 1) Then
          FF3PNTRS.NXT = MDRCHN(MDRIX + 1)
          Call FF3PutRec
       End If
    End If
    MRValue(4) = MDRHIX - MDRIX + 1
    MRValue(5) = MDRHIX
    MDRBGN = MDRCHN(1)
End Sub

Private Sub BldMDRCHN(X)
    MDRRNO = X: MDRMIX = 10: MDRHIX = 0: MDRIX = 0: MDRBGN = 0
    ReDim MDRCHN(MDRMIX)
    If MDRRNO = 0 Then Exit Sub
    Z = MDRRNO
    While Z <> 0  ' Build chain, inserting MDRrno where it belongs
       Call FF3GetRec(Z)
       If OKAY Then
          MDRHIX = MDRHIX + 1
          If Not (MDRHIX < MDRMIX) Then
             MDRMIX = MDRMIX + 10
             ReDim Preserve MDRCHN(MDRMIX)
          End If
          MDRCHN(MDRHIX) = Z
          Z = FF3PNTRS.NXT
          Else
          Z = 0
       End If
    Wend
    If MDRHIX > 0 Then
       MDRIX = 1
       MRValue(4) = MDRHIX - MDRIX + 1
       MRValue(5) = MDRHIX
       MDRBGN = MDRCHN(1)
    End If
End Sub

Private Sub ClearSCRN()
   DTValue(0) = ZERODATE$
   DTValue(1) = ZERODATE$
   For X = 1 To 9: MRValue(X) = "": Next X
   NAVFrame.Enabled = False
   NAValue(0) = ""
   NAValue(1) = ""
   frmFHSWMisc.Show
   NEWMDR = -1
End Sub

Private Sub cmdSetup(cmdOPT)
For N = 1 To 3
    CMDButton(N).Visible = True
    CMDButton(N).Font.Bold = False
    CMDButton(N).Enabled = True
    CMDButton(N).DEFAULT = False
    CMDButton(N).Cancel = False
Next N
For N = 1 To 4
    MRData(N).Enabled = False
Next N
DTValue(0).Enabled = False
DTValue(1).Enabled = False
NAVFrame.Enabled = False
Select Case cmdOPT
   Case 1: ' Standard Misc Record Commands
         NAVFrame.Enabled = True
         NAVButton(1).Enabled = (MDRIX < MDRHIX)
         NAVButton(2).Enabled = (MDRIX > 1)
         CMDButton(0).Caption = "Update"
         CMDButton(1).Caption = "Address"
         CMDButton(1).Font.Bold = (FF3PNTRS.LOC <> 0)
         CMDButton(2).Caption = "Notes"
         CMDButton(2).Font.Bold = (FF3PNTRS.COM <> 0)
         CMDButton(2).Visible = True
         CMDButton(3).Caption = "Return"
         CMDButton(3).Cancel = True
         CMDButton(3).DEFAULT = True
         NEWADR = 0
   Case 2: ' Change Choices
         CMDButton(0).Caption = "Change"
         CMDButton(1).Caption = "Add"
         CMDButton(2).Visible = True
         CMDButton(2).Caption = "Delete"
         CMDButton(3).Caption = "Cancel"
         CMDButton(3).Cancel = True
         CMDButton(3).DEFAULT = True
   Case 3: ' Add/Update Choices
         DTValue(0).Enabled = True
         DTValue(1).Enabled = True
         MRData(MDRTYPE - 6).Enabled = True
         CMDButton(0).Caption = "Save"
         CMDButton(1).Caption = "Cancel"
         CMDButton(1).Cancel = True
         CMDButton(2).Caption = "Restore"
         CMDButton(3).Visible = False
         DTValue(0).SetFocus
End Select
End Sub

Private Sub DeleteMDRRec()
3025 ' Delete Current Miscellaneous Record
   Call PutMSG("UPDTM007")
   If CA$ <> SNGLKEY$(1) Then Exit Sub
   Call FF3GetRec(MDRRNO)
   ecom = FF3PNTRS.COM: eloc = FF3PNTRS.LOC
   If eloc <> 0 Then
      Call PutMSG("UPDTM008")
      If CA$ <> SNGLKEY$(1) Then Exit Sub
   End If
   If mcom <> 0 Then
      Call PutMSG("UPDTM009")
      If CA$ <> SNGLKEY$(1) Then Exit Sub
   End If
   While eloc <> 0
       Call FF2GetRec(eloc)
       If Not OKAY Then
          eloc = 0
          Else
          nloc = FF2PFX.NXT
          acom = FF2PFX.COM
          While acom <> 0
             Call FF3GetRec(acom)
             If Not OKAY Then
                ncom = 0: nloc = 0
                Else
                ncom = FF3COMMENT.FPTR
                Call FF3DelRec(acom)
                If Not OKAY Then
                   ncom = 0: nloc = 0
                End If
             End If
             acom = ncom
          Wend
          If OKAY Then Call FF2DelRec(eloc)
          If Not OKAY Then nloc = 0
          eloc = nloc
       End If
   Wend
   If Not OKAY Then Exit Sub
   While ecom <> 0
       Call FF3GetRec(ecom)
       If Not OKAY Then
          ncom = 0
          Else
          ncom = FF3COMMENT.FPTR
          Call FF3DelRec(ecom)
          If Not OKAY Then ncom = 0
       End If
       ecom = ncom
   Wend
   If Not OKAY Then Exit Sub
   Call FF3DelRec(MDRRNO)
   OBGN = MDRBGN: Call RemMDRCHN
   If OBGN <> MDRBGN Then
      Call FF1GetRec(MSRNO)
      If OKAY Then
         Select Case MDRTYPE
            Case 7: FF1Rec.EDUC = MDRBGN
            Case 8: FF1Rec.WORK = MDRBGN
            Case 9: FF1Rec.MIL = MDRBGN
            Case 10: FF1Rec.HLTH = MDRBGN
         End Select
         Call FF1PutREC
      End If
   End If
End Sub

Private Sub RemMDRCHN()
3017 ' Remove MDRRNO From Miscellaneous Record Chain
   If MDRIX > 1 Then
      Call FF3GetRec(MDRCHN(MDRIX - 1))
      FF3PNTRS.NXT = MDRCHN(MDRIX + 1)
      Call FF3PutRec
      Call FF3GetRec(MDRRNO)
   End If
   For X = MDRIX To MDRHIX - 1
      MDRCHN(X) = MDRCHN(X + 1)
   Next X
   MDRCHN(MDRHIX) = 0
   MDRHIX = MDRHIX - 1
   If MDRIX > MDRHIX Then MDRIX = MDRHIX
   MDRRNO = MDRCHN(MDRIX)
   MDRBGN = MDRCHN(1)
   If MDRHIX > 0 Then
      NAValue(0) = MDRHIX - MDRIX + 1
      NAValue(1) = MDRHIX
      Else
      NAValue(0) = ""
      NAValue(1) = ""
   End If
End Sub

Private Sub SaveMDRInfo()
 ' Write Miscellaneous Record to FILE
   If NEWMDR Then
      FF3PFX.RTYPE = Chr$(MDRTYPE)
      Call FF3GetRec(0)
      FF3PNTRS.SRTYPE = Chr$(1)
      FF3PNTRS.SRNO = MSRNO
   End If
 ' Move Miscellaneous Record Info from Screen
   xX$ = Space$(20)
   hD1Y = FF3PNTRS.D1Y: hD1MD = FF3PNTRS.D1MD
   LSet xX$ = DTValue(0).TEXT
   FF3PNTRS.D1MD = DATEMD(xX$, DFM)
   FF3PNTRS.D1Y = DATEY(xX$, DFM)
   LSet xX$ = DTValue(1).TEXT
   FF3PNTRS.D2MD = DATEMD(xX$, DFM)
   FF3PNTRS.D2Y = DATEY(xX$, DFM)
   Select Case MDRTYPE
      Case 7: LSet FF3EDUC.EDLEV = MRValue(1).TEXT
              LSet FF3EDUC.EDDEG = MRValue(2).TEXT
              LSet FF3EDUC.EDSUB1 = MRValue(3).TEXT
              LSet FF3EDUC.EDSUB2 = MRValue(4).TEXT
      Case 8: LSet FF3WORK.WORKTP = MRValue(5).TEXT
      Case 9: LSet FF3MIL.MILRANK = MRValue(6).TEXT
              LSet FF3MIL.MILSTAT = MRValue(7).TEXT
      Case 10: LSet FF3HEALTH.HLTHDIAG = MRValue(8).TEXT
               LSet FF3HEALTH.HLTHSTAT = MRValue(9).TEXT
   End Select
   Call FF3PutRec
   If OKAY Then
      hBGN = MDRBGN
      MDRDtChg = ((hD1Y <> FF3PNTRS.D1Y) Or (hD1MD <> FF3PNTRS.D1MD))
      If (Not NEWMDR) Then
         If MDRDtChg Then Call RemMDRCHN
      End If
      If NEWMDR Or MDRDtChg Then Call AddMDRCHN(FF3PFX.RNO)
      If hBGN <> MDRBGN Then
         Call FF1GetRec(MSRNO)
         If OKAY Then
            Select Case MDRTYPE
               Case 7: FF1Rec.EDUC = MDRBGN
               Case 8: FF1Rec.WORK = MDRBGN
               Case 9: FF1Rec.MIL = MDRBGN
               Case 10: FF1Rec.HLTH = MDRBGN
            End Select
            Call FF1PutREC
         End If
      End If
   End If
   NEWMDR = 0
End Sub

Private Sub ShowMDRFld(X)
    xX$ = Space$(30)
    Select Case X
       Case -1: LSet xX$ = XSCRNDate$(FF3PNTRS.D2MD, FF3PNTRS.D2Y)
       Case 0: LSet xX$ = XSCRNDate$(FF3PNTRS.D1MD, FF3PNTRS.D1Y)
       Case 1: LSet xX$ = FF3EDUC.EDLEV
       Case 2: LSet xX$ = FF3EDUC.EDDEG
       Case 3: LSet xX$ = FF3EDUC.EDSUB1
       Case 4: LSet xX$ = FF3EDUC.EDSUB2
       Case 5: LSet xX$ = FF3WORK.WORKTP
       Case 6: LSet xX$ = FF3MIL.MILRANK
       Case 7: LSet xX$ = FF3MIL.MILSTAT
       Case 8: LSet xX$ = FF3HEALTH.HLTHDIAG
       Case 9: LSet xX$ = FF3HEALTH.HLTHSTAT
    End Select
    If X < 1 Then
       DTValue(-X).TEXT = Left$(xX$, DTValue(-X).WIDTH)
       Else
       MRValue(X).TEXT = Left$(xX$, MRValue(X).WIDTH)
    End If
End Sub
Private Sub ShowMDRInfo()
3002 ' Move Miscellaneous INFO to Screen
    Call FF3GetRec(MDRRNO)
    xX$ = Space$(30)
    For X = -1 To 0: Call ShowMDRFld(X): Next X
    Select Case MDRTYPE
        Case 7: n1 = 1: n2 = 4
        Case 8: n1 = 5: n2 = 5
        Case 9: n1 = 6: n2 = 7
        Case 10: n1 = 8: n2 = 9
    End Select
    For X = n1 To n2: Call ShowMDRFld(X): Next X
    NAValue(0) = MDRHIX - MDRIX + 1
    NAValue(1) = MDRHIX
    NAVFrame.Enabled = True
    cmdSetup (1)
    frmFHSWMisc.Show
End Sub

Sub ViewMiscRec(X, Y, Z)
    MDRNO = X: MSRNO = Y: MDRTYPE = Z
3000 ' Prepare Form for Miscellaneous Information Record
    If Not ENV.FFOPEN Then Call FAMOpen("WRITE")
    ZERODATE$ = XSCRNDate$(0, 0)
    DTValue(0).TEXT = ZERODATE$
    DTValue(1).TEXT = ZERODATE$
    Select Case MDRTYPE:
       Case 7: frmFHSWMisc.Caption = "FHSWEDUC - Education Record"
               CMDFrame.TOP = 1920
               frmFHSWMisc.Height = 3015
               DATAFrame.Height = 2535
               MRValue(1).MaxLength = 8
               MRValue(2).MaxLength = 4
               MRValue(3).MaxLength = 9
               MRValue(4).MaxLength = 9
       Case 8: frmFHSWMisc.Caption = "FHSWOCCU - Occupation Record"
               CMDFrame.TOP = 1200
               frmFHSWMisc.Height = 2250
               DATAFrame.Height = 1770
               MRValue(5).MaxLength = 30
       Case 9: frmFHSWMisc.Caption = "FHSWMILI - Military Record"
               CMDFrame.TOP = 1560
               frmFHSWMisc.Height = 2625
               DATAFrame.Height = 2145
               MRValue(6).MaxLength = 22
               MRValue(7).MaxLength = 8
       Case 10: frmFHSWMisc.Caption = "FHSWMEDI - Medical Record"
                CMDFrame.TOP = 1560
                frmFHSWMisc.Height = 2625
                DATAFrame.Height = 2145
                MRValue(8).MaxLength = 22
                MRValue(9).MaxLength = 8
    End Select
    xf = MDRTYPE - 6
    MRData(xf).TOP = 480
    MRData(xf).Enabled = True
    MRData(xf).Visible = True
    If MDRNO = 0 Then
       Call ClearSCRN
       Call cmdSetup(3)
       MDRRNO = 0: MDRIX = 0: MDRHIX = 0: MDRBGN = 0
       Exit Sub
    End If

3005 ' Process Existing Record
    Call BldMDRCHN(MDRNO)
    Call ShowMDRInfo
End Sub




Private Sub CMDButton_Click(Index As Integer)
Select Case CMDButton(Index).Caption
  Case "Add"
       Call ClearSCRN
       Call cmdSetup(3)
  Case "Address":
       X = FF3PNTRS.LOC: Y = 10: Z = MDRRNO
       Call frmFHSWAdrs.ViewADDRESS(X, Y, Z)
  Case "Cancel":
       If MDRIX > 0 Then
          Call ShowMDRInfo
          Else
          Unload frmFHSWMisc
       End If
  Case "Change": Call cmdSetup(3)
  Case "Delete"
       Call DeleteMDRRec
       If MDRRNO <> 0 Then
          Call ShowMDRInfo
          Else
          Call ClearSCRN
          Call cmdSetup(3)
       End If
  Case "Update": Call cmdSetup(2)
  Case "Notes":
       X = FF3PNTRS.COM: Y = MDRTYPE: Z = MDRRNO
       Call frmFHSWNote.ViewCOMMENTS(X, Y, Z)
  Case "Restore":
       If LastFLD < 1 Then
          If RTrim$(DTValue(-LastFLD).TEXT) <> HOLDITEM Then
             DTValue(-LastFLD).TEXT = HOLDITEM
             Else
             Call ShowMDRFld(LastFLD)
          End If
          DTValue(-LastFLD).SetFocus
          Else:
          If RTrim$(MRValue(LastFLD).TEXT) <> HOLDITEM Then
             MRValue(LastFLD).TEXT = HOLDITEM
             Else
             Call ShowMDRFld(LastFLD)
          End If
          MRValue(LastFLD).SetFocus
       End If
  Case "Return": Unload frmFHSWMisc
  Case "Save"
       Call SaveMDRInfo
       Call ShowMDRInfo
End Select
End Sub


Private Sub DTValue_GotFocus(Index As Integer)
HOLDITEM = RTrim$(DTValue(Index).TEXT)
DTValue(Index).SelStart = 0
DTValue(Index).TEXT = RTrim$(DTValue(Index).TEXT)
DTValue(Index).SelLength = 0
xFT = 3: LastFLD = -Index
End Sub


Private Sub DTValue_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   XVALUE$ = DTValue(Index).TEXT
   xSS = DTValue(Index).SelStart
   xSL = DTValue(Index).SelLength
   Call FldUpdate(KeyCode, Shift, xFT, xSS, xSL, XVALUE$)
   DTValue(Index).TEXT = XVALUE$
   DTValue(Index).SelStart = xSS
   DTValue(Index).SelLength = xSL
End Sub


Private Sub DTValue_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = 0
End Sub




Private Sub Form_Activate()
If Not ENV.FFOPEN Then Call FAMOpen("WRITE")
If MDRRNO <> 0 Then Call cmdSetup(1)
End Sub

Private Sub Form_Load()
frmFHSWUPDT.Enabled = False
FormLoaded(3) = FormLoaded(3) Or MrecForm
' Set Default Form Location on Screen
frmHEIGHT = frmFHSWMisc.Height
frmWIDTH = frmFHSWMisc.WIDTH
t1 = GetSetting("FHSWIN", "frmMISC", "TOP", 0)
If t1 = 0 Then
   t00 = frmFHSWUPDT.TOP: w00 = frmFHSWUPDT.WIDTH
   l00 = frmFHSWUPDT.Left
   t1 = t00 + 2715
   l1 = l00 + (w00 - frmWIDTH) / 2
   If l1 < 0 Then
      l1 = l00 - frmWIDTH / 2
   End If
   Else
   t1 = 1440: l1 = 1440
End If
frmFHSWMisc.TOP = GetSetting("FHSWIN", "frmMREC", "Top", t1)
frmFHSWMisc.Left = GetSetting("FHSWIN", "frmMREC", "Left", l1)
MDRIX = 0: MDRMIX = 10: MDRHIX = 0
ReDim MDRCHN(MDRMIX)
End Sub


Private Sub Form_Unload(Cancel As Integer)
FormLoaded(3) = FormLoaded(3) And (255 - MrecForm)
SaveSetting "FHSWIN", "frmMRec", "Top", frmFHSWMisc.TOP
SaveSetting "FHSWIN", "frmMRec", "Left", frmFHSWMisc.Left
frmFHSWUPDT.Enabled = True
frmFHSWUPDT.SetFocus
End Sub


Private Sub MRValue_GotFocus(Index As Integer)
HOLDITEM = RTrim$(MRValue(Index).TEXT)
MRValue(Index).SelStart = 0
MRValue(Index).TEXT = RTrim$(MRValue(Index).TEXT)
MRValue(Index).SelLength = Len(MRValue(Index).TEXT)
LastFLD = Index
End Sub


Private Sub NAVButton_Click(Index As Integer)
Select Case Index
   Case 0: ' First
      MDRIX = MDRHIX: MDRRNO = MDRCHN(MDRIX)
      Call ShowMDRInfo
   Case 1: ' Prev
      If MDRIX < MDRHIX Then
         MDRIX = MDRIX + 1
         MDRRNO = MDRCHN(MDRIX)
         Call ShowMDRInfo
      End If
   Case 2: ' Next
      If MDRIX > 0 Then
         MDRIX = MDRIX - 1
         MDRRNO = MDRCHN(MDRIX)
         Call ShowMDRInfo
      End If
   Case 3: ' Last
      MDRIX = 1: MDRRNO = MDRCHN(1)
      Call ShowMDRInfo
End Select
End Sub


