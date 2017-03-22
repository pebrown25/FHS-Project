VERSION 4.00
Begin VB.Form frmFHSWNdWk 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FHSWNDWK - Node Work File"
   ClientHeight    =   2070
   ClientLeft      =   2115
   ClientTop       =   5745
   ClientWidth     =   7425
   Height          =   2475
   Left            =   2055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   Top             =   5400
   Width           =   7545
   Begin VB.ComboBox cboSIBLING 
      Height          =   315
      ItemData        =   "FHSWNDWK.frx":0000
      Left            =   6720
      List            =   "FHSWNDWK.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   1755
      Width           =   615
   End
   Begin VB.TextBox NFValue 
      Height          =   285
      Index           =   8
      Left            =   6240
      MaxLength       =   2
      TabIndex        =   26
      Text            =   "99"
      Top             =   1395
      Width           =   495
   End
   Begin VB.CommandButton cmdCREATE 
      Caption         =   "Create File"
      Height          =   375
      Left            =   5640
      TabIndex        =   24
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdCANCEL 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   6480
      TabIndex        =   23
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdOKAY 
      Caption         =   "OKAY"
      Default         =   -1  'True
      Height          =   315
      Left            =   5400
      TabIndex        =   22
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox cboAlign 
      Height          =   315
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdBROWSE 
      Caption         =   "Browse"
      Height          =   315
      Left            =   4440
      TabIndex        =   17
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox NFValue 
      Height          =   315
      Index           =   0
      Left            =   2760
      TabIndex        =   16
      Top             =   120
      Width           =   1575
   End
   Begin VB.ComboBox NFType 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "FHSWNDWK.frx":0014
      Left            =   840
      List            =   "FHSWNDWK.frx":0021
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox NFValue 
      Height          =   285
      Index           =   1
      Left            =   960
      TabIndex        =   12
      Top             =   585
      Width           =   615
   End
   Begin VB.TextBox NFValue 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS LineDraw"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   1920
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   600
      Width           =   3255
   End
   Begin VB.CommandButton cmdLIST 
      Caption         =   "<"
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   10
      Top             =   585
      Width           =   255
   End
   Begin VB.TextBox NFValue 
      Height          =   285
      Index           =   3
      Left            =   960
      TabIndex        =   9
      Top             =   945
      Width           =   615
   End
   Begin VB.TextBox NFValue 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS LineDraw"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   4
      Left            =   1920
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton cmdLIST 
      Caption         =   "<"
      Height          =   285
      Index           =   1
      Left            =   1560
      TabIndex        =   7
      Top             =   945
      Width           =   255
   End
   Begin VB.TextBox NFValue 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   7
      Left            =   4560
      TabIndex        =   2
      Top             =   1720
      Width           =   615
   End
   Begin VB.TextBox NFValue 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   6
      Left            =   2880
      TabIndex        =   1
      Top             =   1720
      Width           =   495
   End
   Begin VB.TextBox NFValue 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   5
      Left            =   840
      TabIndex        =   0
      Top             =   1720
      Width           =   615
   End
   Begin VB.Label NFLabel 
      Caption         =   "Show Siblings:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   5400
      TabIndex        =   27
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label NFLabel 
      Caption         =   "Max GL:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   5400
      TabIndex        =   25
      Top             =   1440
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   5280
      X2              =   7560
      Y1              =   915
      Y2              =   915
   End
   Begin VB.Label NFLabel 
      Caption         =   "Align:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   5400
      TabIndex        =   21
      Top             =   1005
      Width           =   735
   End
   Begin VB.Label NFLabel 
      Caption         =   "File:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   19
      Top             =   150
      Width           =   375
   End
   Begin VB.Label NFLabel 
      Caption         =   "Type:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   18
      Top             =   150
      Width           =   735
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   5280
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label NFLabel 
      Caption         =   "Base ID:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   600
      Width           =   855
   End
   Begin VB.Label NFLabel 
      Caption         =   "ID2:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   480
      TabIndex        =   13
      Top             =   945
      Width           =   495
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   5280
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label NFLabel 
      Caption         =   "Hi Column:"
      Height          =   255
      Index           =   7
      Left            =   3600
      TabIndex        =   6
      Top             =   1725
      Width           =   855
   End
   Begin VB.Label NFLabel 
      Caption         =   "Number Nodes:"
      Height          =   255
      Index           =   6
      Left            =   1680
      TabIndex        =   5
      Top             =   1725
      Width           =   1215
   End
   Begin VB.Label NFLabel 
      Caption         =   "HiGL:"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   4
      Top             =   1725
      Width           =   495
   End
   Begin VB.Label NFLabel 
      Caption         =   "Statistics:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   3
      Top             =   1450
      Width           =   975
   End
   Begin VB.Line Line2 
      X1              =   5280
      X2              =   5280
      Y1              =   0
      Y2              =   2040
   End
End
Attribute VB_Name = "frmFHSWNdWk"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
DefInt A-Z
   Dim ANCSEARCH, BNAME$, RTYPE$, WALIGN, ALIGNTAB$()
   Dim SCHAR$, BOXWIDTH, ORIENT, ORIENT2
   Dim BRID, BRID2, LastNF, DataChanged
   Dim HOLDITEM As Variant, NFDSN As String
   Dim NFAlign, NFSibling, NFMaxGL
Private Sub GetRULES()
NFAlign = GetSetting("FHSWIN", "frmNDWK", "Align", 0)
NFMaxGL = GetSetting("FHSWIN", "frmNDWK", "MaxGL", 99)
NFSibling = GetSetting("FHSWIN", "frmNDWK", "Siblings", 0)
End Sub

Private Sub ChangeEnable()
If NWrkType <> 3 Then
   NFLabel(8).Enabled = False
   NFValue(3).Enabled = False
   NFValue(4).Enabled = False
   cmdLIST(1).Enabled = False
   NFLabel(10).Enabled = False
   cboSIBLING.Enabled = False
End If
End Sub


Private Sub CheckNodeFile()
FAMPATH = (NWrkType = 3)
ANCSEARCH = (NWrkType = 1)
If NFDSN = "" Then
   NFDSN = FT$(8)
   NFValue(0).TEXT = NFDSN
End If
FT$(8) = NFValue(0).TEXT
Call NodeWrkOPEN("READQ")
NWrkOkay = OKAY
If OKAY Then
   NFDSN = FT$(8)
   Else
   FT$(8) = NFDSN
   GoTo NFChkExit
End If
WALIGN = Asc(NodeHDR.WFOPT) And 15
If WALIGN = 0 Then GoTo NFChkExit
If FAMPATH Then
   If ((Asc(NodeHDR.WFOPT) And 32) = 0) Then
      GoTo NFChkExit
      Else
      BRID = NodeHDR.WFID1: BRID2 = NodeHDR.WFID2
   End If
   Else
   If (ANCSEARCH And ((Asc(NodeHDR.WFOPT) And 16) = 0)) Then
      GoTo NFChkExit
      Else
      If ((Not ANCSEARCH) And ((Asc(NodeHDR.WFOPT) And 16) > 0)) Then
         GoTo NFChkExit
         Else
         BRID = NodeHDR.WFBRID
      End If
   End If
End If
Call ShowBASE(0, 0)
If WALIGN > 5 Then WALIGN = WALIGN - 4
Select Case (WALIGN And 7)
   Case 1: NFAlign = 0
   Case 2: NFAlign = 1
   Case 3: NFAlign = 3
   Case 4: NFAlign = 2
   Case Else: NFAlign = 0
End Select
NFMaxGL = NodeHDR.WFMAXGL
Call ShowRULES: Call ShowStats
NFChkExit:
Call FamCLOSE
DataChanged = 0
End Sub
Private Sub ClearBASE()
For X = 1 To 4: NFValue(X) = "": Next X
End Sub

Private Sub ShowBASE(X, Y)
If (X <> 0) Or (NodeHDR.WFID1 <> 0) Or (NodeHDR.WFID2 <> 0) Or (NodeHDR.WFBRID <> 0) Then
   Call FAMOpen("READ"): If Not OKAY Then Exit Sub
End If
If X <> 0 Then
   Call FF1GetRec(X)
   If OKAY Then
      NFValue(Y + 1).TEXT = FMTNAME(1)
   End If
   Else
   ClearBASE
   If Not (NodeHDR.WFID1 = 0) Then
      NFValue(1).TEXT = NodeHDR.WFID1
      Call FF1GetRec(NodeHDR.WFID1)
      If OKAY Then
         NFValue(2).TEXT = FMTNAME(1)
      End If
   End If
   If Not (NodeHDR.WFBRID = 0) Then
      NFValue(1).TEXT = NodeHDR.WFBRID
      Call FF1GetRec(NodeHDR.WFBRID)
      If OKAY Then
         NFValue(2).TEXT = FMTNAME(1)
      End If
   End If
   If Not (NodeHDR.WFID2 = 0) Then
      NFValue(3).TEXT = NodeHDR.WFID2
      Call FF1GetRec(NodeHDR.WFID2)
      If OKAY Then
         NFValue(4).TEXT = FMTNAME(1)
      End If
   End If
End If
Call FamCLOSE
End Sub
Private Sub ShowStats()
NFValue(5).TEXT = NodeHDR.WFHIGL
NFValue(6).TEXT = NodeHDR.WFNUMREL
NFValue(7).TEXT = (NodeHDR.WFMAXCOL + 1) \ 2
End Sub
Private Sub ClearSTATS()
For X = 5 To 7: NFValue(X).TEXT = "": Next X
End Sub

Private Sub ClearRULES()

End Sub
Private Sub ShowRULES()
cboAlign.ListIndex = NFAlign
NFValue(8).TEXT = NFMaxGL
cboSIBLING.ListIndex = NFSibling
End Sub

'Rem $STATIC
Sub BUILD()
    CTYPE = NWrkType
    ANCSEARCH = (CTYPE = 1): FAMPATH = (CTYPE = 3)
    If DataChanged Then GoTo 605
600 ' Create NODE Work File
    Call NodeWrkOPEN("WRITEQ"): If Not OKAY Then GoTo 605
    Call FAMOpen("READ"): If Not OKAY Then GoTo 695
    WALIGN = Asc(NodeHDR.WFOPT) And 15
    If WALIGN = 0 Then GoTo 605
    If (FAMPATH And ((Asc(NodeHDR.WFOPT) And 32) = 0)) Then GoTo 605
    If (ANCSEARCH And ((Asc(NodeHDR.WFOPT) And 16) = 0)) Then GoTo 605
    If ((Not ANCSEARCH) And ((Asc(NodeHDR.WFOPT) And 16) > 0)) Then GoTo 605
    Call PutMSG("HBOXM004")
    If CA$ = Chr$(27) Then GoTo 695 Else a = Val(CA$)
    If a = 2 Then GoTo 700

605 ' Create NEW NodeWork File
    Call FamCLOSE: AGLMAX = 0: DGLMAX = 0
    WALIGN = 0: xSTAT$ = " "
    NOSIBLING = (NFSibling = 1)
    If Not FAMPATH Then
       Select Case CTYPE
          Case 1: Call RptLOAD("AHBX")
                  'Call RptOptGET("MXAL", AGLMAX)
                  Call RptOptGET("ALNS", XLNS)
          Case 2: Call RptLOAD("DHBX")
                  'Call RptOptGET("MXDL", DGLMAX)
                  Call RptOptGET("DLNS", XLNS)
       End Select
       If (XLNS < 1) Or (XLNS > 3) Then XLNS = 1
       Call RptOptGET("ADOP", X): ADOPTOK = (X = 1)
    End If
    Call NodeWrkOPEN("NEW"): If Not OKAY Then GoTo 695
    Call FAMOpen("READ"): If Not OKAY Then GoTo 695
    If FAMPATH Then
       GoSub 1900: 'DGLMAX = NodeHDR.WFMAXGL
       If CA$ = Chr$(27) Then GoTo 695
       DGLMAX = NFMaxGL
       GoTo 615
    End If
    'Call FmtFIND("XXXXS010")
    'Call FmtFindFLD("BRID")
    'If FOUND Then RSet UTXT$(FFLD.UTXT) = FNX$(NodeHDR.WFBRID)
610 'Call MsgFORMAT: If A = 27 Then GoTo 695 Else x& = 0
    'Call FmtFindFLD("BRID"): If FOUND Then x& = Val(UTXT$(FFLD.UTXT))
    'If XX& = 0 Then GoTo 695
    'If 1 > XX& Or XX& > FF1Hdr.MAXID Then Call ErrBEEP(0): GoTo 610
    'x = XX&
    X = BRID
    Call FF1GetRec(X): If Not OKAY Then GoTo 610
    NodeHDR.WFBRID = X
    NodeHDR.WFOPT = Chr$(-64 * NOSIBLING - 32 * FAMPATH - 16 * ANCSEARCH + 5)
    NOSIBLING = 0: ID1 = 0: ID2 = 0
    NodeHDR.WFOPT2 = Chr$(-16 * ADOPTOK + XLNS - 1)
    If ANCSEARCH Then
       AGLMAX = NFMaxGL
       NodeHDR.WFMAXGL = AGLMAX
       Else
       DGLMAX = NFMaxGL
       NodeHDR.WFMAXGL = DGLMAX
    End If
    'BRID = x: BNAME$ = FMTNAME$(1): Call ShowINFO
615 ' Initialize GL Records
    NodeGLREC.WFGLBGN = 0: NodeGLREC.WFGLEND = 0
    NodeGLREC.WFLOCOL = 0: NodeGLREC.WFHICOL = 0
    For GL = 0 To NodeHDR.WFMAXGL
        NodeGLREC.WFGLNUM = GL
        Call NodeWorkPUT("GLREC", GL)
    Next GL
    ReDim GLFCH(NodeHDR.WFMAXGL + 1), GLLCH(NodeHDR.WFMAXGL + 1)
    ReDim xGLCOL(NodeHDR.WFMAXGL + 1), NodePTR(FF1Hdr.MAXID)
    GoSub 670: 'If CA$ = Chr$(27) Then GoTo 695
    WALIGN = Align: 'Call ShowINFO:
    WALIGN = 5

625 ' Create NODE Records for Relatives
    PREC = 0: id = NodeHDR.WFBRID
    NUMCH = 0: GLCOL = 0: a = 0: GL = 0
    Call FF1GetRec(id)
    If ANCSEARCH Then ALIN# = 1: CHBGN = 0 Else GoSub 1970
    GoSub 660
    If CA$ = Chr$(27) Then GoTo 695
    GLFCH(0) = 1
630 ' Process next GL of Relatives
    BGNREC = GLFCH(GL): If BGNREC = 0 Then GoTo 680
    xGLCOL(GL) = GLCOL: ENDREC = NodeHDR.WFNUMREL + 1: GLLCH(GL) = ENDREC
    GLCOL = 0: GL = GL + 1: If GL > NodeHDR.WFMAXGL Then GoTo 680
    SysVAR.MVAR = RTYPE$: SysVAR.NVAR1 = GL
    Call PutMSG("HBOXM005")
    For PREC = BGNREC To ENDREC
       Call NodeWorkGET("NODEREC", PREC)
       id = NodeREC.WFCHBGN: If id < 0 Then GoTo 635
       NUMCH = 0: FCH = 0: PID = NodeREC.WFRID
       If ANCSEARCH Then
          XALIN# = ALIN#: GoSub 650
          Else
          While id > 0: Call FF1GetRec(id)
             OKAY = (FAMPATH Or ADOPTOK)
             If Not OKAY Then
                If FF1Rec.FID = PID Then LSet xSTAT$ = FDSTAT$("FID") Else LSet xSTAT$ = FDSTAT$("MID")
                OKAY = (xSTAT$ <> "*")
             End If
             If OKAY Then
                NUMCH = NUMCH + 1
                NodeHDR.WFNUMREL = NodeHDR.WFNUMREL + 1
                GoSub 1970: GoSub 660
                If a = 27 Then
                   id = 0: PREC = ENDREC
                   Else
                   If FCH = 0 Then FCH = NodeHDR.WFNUMREL + 1
                   If GLFCH(GL) = 0 Then GLFCH(GL) = FCH
                   GoSub 1980
                End If
                Else
                GoSub 1980
             End If
          Wend
       End If
       ' Update FCH and NUMCH in "Parent" Record
       Call NodeWorkGET("NODEREC", PREC)
       NodeREC.WFNUMCH = NUMCH: NodeREC.WFCHBGN = FCH
       Call NodeWorkPUT("NODEREC", PREC)
635 Next PREC: If a = 27 Then GoTo 695
    GoTo 630

650 ' Create PARENT Records for Ancestor Search
    CNT = 0: CHBGN = 0: LSet xSTAT$ = " "
    Call FF1GetRec(PID)
    If (GL = 1) Or (XLNS = 1) Or ((XLNS = 2) And (FF1Rec.SEX = FGENDR$(1))) Or ((XLNS = 3) And (FF1Rec.SEX = FGENDR$(2))) Then
       id = FF1Rec.FID
       If id > 0 Then
          NUMCH = 1 + 1024 * (Asc(FF1Rec.STATUS) And 3)
          ALIN# = 2 * XALIN#: GoSub 655
       End If
       id = FF1Rec.MID
       If id > 0 Then
          NUMCH = 2 + 1024 * ((Asc(FF1Rec.STATUS) \ 4) And 3)
          ALIN# = 2 * XALIN# + 1: GoSub 655
       End If
    End If
    NUMCH = CNT
    Return

655 ' Create Initial NODE Record for Ancestor
    OKAY = (ADOPTOK Or ((NUMCH And 1024) = 0))
    If OKAY Then
       CNT = CNT + 1: NodeHDR.WFNUMREL = NodeHDR.WFNUMREL + 1
       If FCH = 0 Then FCH = NodeHDR.WFNUMREL + 1
       If GLFCH(GL) = 0 Then GLFCH(GL) = FCH
       GoSub 660
    End If
    Return

660 ' Create Initial NODE Record for Relative
    GLCOL = GLCOL + 1
    If GLCOL > NodeHDR.WFMAXCOL Then NodeHDR.WFMAXCOL = GLCOL
    NodeREC.WFPREC = PREC: NodeREC.WFCHNUM = NUMCH
    NodeREC.WFRID = id
    NodeREC.WFCOL = GLCOL: NodeREC.WFNUMCH = 0
    If GL = NodeHDR.WFMAXGL Then
       NodeREC.WFCHBGN = 0
       Else
       NodeREC.WFCHBGN = CHBGN
    End If
    Call NodeWorkPUT("NODEREC", NodeHDR.WFNUMREL + 1)
    If NodePTR(id) > 0 Then
       Call NodeWorkGET("NODEREC", NodePTR(id) + 1)
       hptr = NodeREC.WFCHBGN: HCNT = NodeREC.WFNUMCH
       Call NodeWorkGET("NODEREC", NodeHDR.WFNUMREL + 1)
       NodeREC.WFCHBGN = -NodePTR(id)
       NodeREC.WFNUMCH = -(HCNT + 1) * (hptr < 0)
       Call NodeWorkPUT("NODEREC", NodeHDR.WFNUMREL + 1)
    End If
    NodePTR(id) = NodeHDR.WFNUMREL + 1: NodeHDR.WFHIGL = GL
    CA$ = INKEY$: If CA$ <> "" Then a = Asc(CA$): Call ShowStats
    Return

670 ' Determine Alignment Type
    If FAMPATH And NOSIBLING And ((RGL1 = 0) Or (RGL2 = 0)) Then
       Align = 1
       Else
       'If ANCSEARCH Then
       '   Call PutMSG("HBOXM006")
       '   Else
       '   Call PutMSG("HBOXM007")
       'End If
       'If CA$ <> Chr$(27) Then
       '   Align = Val(CA$) + 1
       Select Case NFAlign
          Case 0: Align = 1
          Case 1: Align = 2
          Case 2: Align = 4
          Case 3: Align = 3
       End Select
       Align = Align + (Align > 1) * ANCSEARCH * 4
       If (Align < 1) Or (Align > 8) Then Align = 1
       'End If
    End If
    Return

680 ' Update Header Records
    NodeHDR.WFVER = Chr$(0)
    Call NodeWorkPUT("HDR", 1)
    X = 1
    For GL = 0 To NodeHDR.WFHIGL
        NodeGLREC.WFGLNUM = GL
        NodeGLREC.WFGLBGN = X: NodeGLREC.WFGLEND = GLLCH(GL)
        NodeGLREC.WFLOCOL = 1: NodeGLREC.WFHICOL = xGLCOL(GL)
        X = X + xGLCOL(GL)
        Call NodeWorkPUT("GLREC", GL)
    Next GL: Call ShowStats
    If a = 27 Then GoTo 695
    Erase GLFCH, GLLCH, xGLCOL, NodePTR
    GoTo 700

695 ' Premature End to NODE Work Create
    Call FamCLOSE
    Exit Sub

700 ' Determine Placement of BOXes
    If Align = 0 Then GoSub 670: If CA$ = Chr$(27) Then GoTo 799
    NodeHDR.WFOPT = Chr$((Asc(NodeHDR.WFOPT) And 240) + 5)
    Call NodeWorkPUT("HDR", 1)
    WALIGN = 5: Call ShowStats
    HIGL = NodeHDR.WFHIGL
    Call PutMSG("HBOXM008")
    ReDim GLFCH(HIGL + 1), GLLCH(HIGL + 1)
    ReDim GLLOCOL(HIGL + 1), GLHICOL(HIGL + 1)
    ReDim xGLREC(HIGL + 1), GLNUMCH(HIGL + 1)
    GL = 0: xGLREC(0) = 1: GLFCH(0) = 1: GLLCH(0) = 1: NodeHDR.WFMAXCOL = 0
    For X = 0 To HIGL: GLHICOL(X) = -1: Next X

710 ' First Placement for NODE Rec at GL Level
    GLREC = xGLREC(GL)
    Call NodeWorkGET("NODEREC", GLREC)
    GLCOL = GLHICOL(GL) + 2
    If GLCOL > NodeHDR.WFMAXCOL Then NodeHDR.WFMAXCOL = GLCOL
    GLHICOL(GL) = GLCOL
    FCH = NodeREC.WFCHBGN: NUMCH = NodeREC.WFNUMCH
    If GLNUMCH(GL) = 1 Then NodeREC.WFCHNUM = NodeREC.WFCHNUM Or 256
    NodeREC.WFCOL = GLCOL
    Call NodeWorkPUT("NODEREC", GLREC)
    CA$ = INKEY$: If CA$ <> "" Then If CA$ = Chr$(27) Then GoTo 799 Else Call ShowStats

720 ' Move from GL to GL+1 (Process Children)
    If FCH < 1 Then GoTo 730
    GL = GL + 1: xGLREC(GL) = FCH
    GLFCH(GL) = GLHICOL(GL) + 2: GLLCH(GL) = FCH + NUMCH - 1
    GLNUMCH(GL) = 1
    CA$ = INKEY$: If CA$ <> "" Then Call ShowStats: If CA$ = Chr$(27) Then GoTo 799
    GoTo 710

730 ' Process next Sibling at Level GL
    GLREC = xGLREC(GL): If GLREC = GLLCH(GL) Then GoTo 740
    xGLREC(GL) = GLREC + 1: GLNUMCH(GL) = GLNUMCH(GL) + 1
    CA$ = INKEY$: If CA$ <> "" Then Call ShowStats: If CA$ = Chr$(27) Then GoTo 799
    GoTo 710

740 ' Move Back to Prior GL (After Processing Children)
    If GL = 0 Then GoTo 790
    FCH = GLFCH(GL): LCH = GLHICOL(GL)
    Call NodeWorkGET("NODEREC", GLREC)
    NodeREC.WFCHNUM = NodeREC.WFCHNUM Or 512
    Call NodeWorkPUT("NODEREC", GLREC)
    GL = GL - 1: GLCOL = GLHICOL(GL): GLREC = xGLREC(GL)
745 On Align GoTo 746, 747, 748, 749, 746, 747, 748, 749
746 Shift = ((GLCOL - FCH) * (GLCOL < FCH) + (GLCOL - LCH) * (GLCOL > LCH)): GoTo 750
747 Shift = (FCH - GLCOL): GoTo 750
748 Shift = (((LCH + FCH) / 2) - GLCOL): GoTo 750
749 Shift = (LCH - GLCOL)

750 ' Check for Required Adjustment
    If Shift > 0 Then
       GoSub 760
       Else
       If Shift < 0 Then Shift = -Shift: BGL = GL: GoSub 770
    End If
    CA$ = INKEY$: If CA$ <> "" Then Call ShowStats: If CA$ = Chr$(27) Then GoTo 799
    GoTo 730

760 ' Shift Parent to Right
    GLCOL = GLCOL + Shift
    GLHICOL(GL) = GLCOL
    If GLCOL > NodeHDR.WFMAXCOL Then NodeHDR.WFMAXCOL = GLCOL
    If GLNUMCH(GL) = 1 Then GLFCH(GL) = GLCOL
    Call NodeWorkGET("NODEREC", GLREC)
    NodeREC.WFCOL = GLCOL
    Call NodeWorkPUT("NODEREC", GLREC)
    Return

770 ' Shift Children to Right
    Call NodeWorkGET("NODEREC", GLREC)
    FCH = NodeREC.WFCHBGN: NUMCH = NodeREC.WFNUMCH: If FCH < 1 Then GoTo 775
    GL = GL + 1: GLREC = FCH: xGLREC(GL) = FCH: GLLCH(GL) = FCH + NUMCH - 1
    Call NodeWorkGET("NODEREC", GLREC)
    GLCOL = NodeREC.WFCOL + Shift: NodeREC.WFCOL = GLCOL
    Call NodeWorkPUT("NODEREC", GLREC)
    GLFCH(GL) = GLCOL: GLHICOL(GL) = GLCOL: GoTo 770

775 ' Process Next Child This Level
    If xGLREC(GL) = GLLCH(GL) Then GoTo 780
    GLREC = xGLREC(GL) + 1: xGLREC(GL) = GLREC
    Call NodeWorkGET("NODEREC", GLREC)
    GLCOL = NodeREC.WFCOL + Shift: NodeREC.WFCOL = GLCOL
    Call NodeWorkPUT("NODEREC", GLREC)
    GLHICOL(GL) = GLCOL: GoTo 770

780 ' No More Children at This Level
    If GLHICOL(GL) > NodeHDR.WFMAXCOL Then NodeHDR.WFMAXCOL = GLHICOL(GL)
    GL = GL - 1: If GL > BGL Then GoTo 775
    Return

790 ' Finished Processing Placement of NODEs
    NodeHDR.WFOPT = Chr$((Asc(NodeHDR.WFOPT) And 240) + Align)
    Call NodeWorkPUT("HDR", 1)
    Call PutMSG("HBOXM009")
    For GL = 0 To HIGL
        Call NodeWorkGET("GLREC", GL)
        GLLOCOL(GL) = NodeGLREC.WFGLBGN
    Next GL
    For GL = 0 To HIGL
        Call NodeWorkGET("NODEREC", GLLOCOL(GL))
        GLLOCOL(GL) = NodeREC.WFCOL
    Next GL
    For GL = 0 To HIGL
        Call NodeWorkGET("GLREC", GL)
        NodeGLREC.WFLOCOL = GLLOCOL(GL)
        NodeGLREC.WFHICOL = GLHICOL(GL)
        Call NodeWorkPUT("GLREC", GL)
    Next GL
799 Call FamCLOSE
    DataChanged = 0
    Exit Sub

1900 ' Prepare for FAMILY PATH Search
     ReDim ANCESTOR(FF1Hdr.MAXID), CHILD(FF1Hdr.MAXID)
     ReDim BL1(101), BL2(101), CHCNT(101)
     'BRID = 0: BNAME$ = "": Call ShowINFO
1904 'Call FmtFIND("HBOXS011"): Call FmtDREST
1905 'Call MsgFORMAT: If A = 27 Then GoTo 1959
     'Call FmtFindFLD("ID1"): ID1# = Val(UTXT$(FFLD.UTXT))
     'If (xID1# < 1) Or (FF1Hdr.MAXID < xID1#) Then Call ErrBEEP(0): GoTo 1905
     ID1 = BRID
     'id1 = xID1#
     Call FF1GetRec(ID1)
     'BRID = id1: BNAME$ = FMTNAME$(1): Call ShowINFO
1909 'Call FmtFIND("HBOXS012"): Call FmtDREST
1910 'Call MsgFORMAT: If A = 27 Then GoTo 1904
     'Call FmtFindFLD("ID2"): ID2# = Val(UTXT$(FFLD.UTXT))
     'If (xID2# < 1) Or (FF1Hdr.MAXID < xID2#) Then Call ErrBEEP(0): GoTo 1910
     'id2 = xID2#
     ID2 = BRID2
     Call FF1GetRec(ID2)
     'BRID = id2: BNAME$ = FMTNAME$(1): Call ShowINFO
     ANCESTOR(1) = ID1: CHILD(ID1) = ID1: CHILD(0) = -1
     AX = 1: ix = 1: 'COLOR FG(1), BG(1)
     Call PutMSG("HBOXM010")
1925 id = ANCESTOR(ix): Call FF1GetRec(id)
     FID = FF1Rec.FID: xMID = FF1Rec.MID
     If CHILD(FID) = 0 Then AX = AX + 1: ANCESTOR(AX) = FID: CHILD(FID) = id
     If CHILD(xMID) = 0 Then AX = AX + 1: ANCESTOR(AX) = xMID: CHILD(xMID) = id
     If ix < AX Then ix = ix + 1: GoTo 1925
1930 RGL1 = 0: RGL2 = 0
     If CHILD(ID2) > 0 Then BL2(0) = ID2: AID = ID2: GoTo 1955
     ANCESTOR(1) = ID2: CHILD(ID2) = -ID2
     AX = 1: ix = 1
     Call PutMSG("HBOXM011")
1935 id = ANCESTOR(ix): Call FF1GetRec(id)
     FID = FF1Rec.FID: xMID = FF1Rec.MID
     If CHILD(FID) > 0 Then BL2(0) = FID: AID = id: GoTo 1945
     If CHILD(FID) = 0 Then AX = AX + 1: ANCESTOR(AX) = FID: CHILD(FID) = -id
     If CHILD(xMID) > 0 Then BL2(0) = xMID: AID = id: GoTo 1945
     If CHILD(xMID) = 0 Then AX = AX + 1: ANCESTOR(AX) = xMID: CHILD(xMID) = -id
     If ix < AX Then ix = ix + 1: GoTo 1935
     'CA$ = INKEY$: While CA$ <> "": CA$ = INKEY$: Wend
     Call PutMSG("HBOXM012")
     'If CA$ <> Chr$(27) Then GoTo 1909
     'BRID = 0: BNAME$ = "": Call ShowINFO
     CA$ = Chr$(27): GoTo 1959
1945 RGL2 = RGL2 + 1: BL2(RGL2) = AID
     If RGL2 < 100 Then
        If CHILD(AID) <> -AID Then AID = -CHILD(AID): GoTo 1945
     End If
     AID = BL2(0)
1955 BL1(RGL1) = AID
     If RGL1 < 100 Then
        If CHILD(AID) <> AID Then AID = CHILD(AID): RGL1 = RGL1 + 1: GoTo 1955
     End If

     ' Found Common Ancestor
     NodeHDR.WFBRID = BL1(0): Call FF1GetRec(NodeHDR.WFBRID)
     BRID = NodeHDR.WFBRID: BNAME$ = FMTNAME$(1): Call ShowStats
     Call PutMSG("HBOXM013")
     If CA$ = Chr$(27) Then GoTo 1959
     For GL = RGL1 + 1 To 101: BL1(GL) = 0: Next GL
     For GL = RGL2 + 1 To 101: BL2(GL) = 0: Next GL
     NodeHDR.WFMAXGL = RGL1 - (RGL2 - RGL1) * (RGL2 > RGL1)
     'CA$ = INKEY$: While CA$ <> "": CA$ = INKEY$: Wend
     'Call PutMSG("HBOXM014")
     'NOSIBLING = (CA$ = SNGLKEY$(2))
     NOSIBLING = (NFSibling = 1)
     'If CA$ <> Chr$(27) Then
        NodeHDR.WFOPT = Chr$(-64 * NOSIBLING + 32 + 5)
        NodeHDR.WFID1 = ID1: NodeHDR.WFID2 = ID2
        NodeHDR.WFRGL1 = Chr$(RGL1): NodeHDR.WFRGL2 = Chr$(RGL2)
     'End If
1959 Erase ANCESTOR, CHILD
     Return

1970 ' Determine First Child ID
     CHBGN = FF1Rec.OLDCH
     If FAMPATH Then
        If (BL1(GL) = id) And (BL1(GL + 1) <> 0) Then
           If NOSIBLING Then CHBGN = BL1(GL + 1)
           Else
           If (BL2(GL) = id) And (BL2(GL + 1) <> 0) Then
              If NOSIBLING Then CHBGN = BL2(GL + 1)
              Else
              CHBGN = 0
           End If
        End If
        Else
        If (GL > 0) And (XLNS > 1) Then
           If FF1Rec.SEX <> FGENDR$(XLNS - 1) Then CHBGN = 0
        End If
     End If
     Return

1980 ' Determine Next Child ID
     If (FAMPATH And NOSIBLING) Then
        If (BL1(GL) = id) And (BL1(GL - 1) = BL2(GL - 1)) Then
           id = BL2(GL)
           Else
           id = 0
        End If
        Else
        If FF1Rec.FID = PID Then id = FF1Rec.FCH Else id = FF1Rec.MCH
     End If
     Return

End Sub




Private Sub cboAlign_Click()
NFAlign = cboAlign.ListIndex
End Sub

Private Sub cboSIBLING_Click()
NFSibling = cboSIBLING.ListIndex
End Sub


Private Sub cmdBROWSE_Click()
HOLDITEM = NFValue(0).TEXT
Call frmFHSWFile.LOCATE("NODEWORK", "FULL")
If OKAY Then
   NFValue(0).TEXT = FT$(8)
   If NFValue(0).TEXT <> HOLDITEM Then
      Call NFValue_LostFocus(0)
   End If
End If
End Sub

Private Sub cmdCANCEL_Click()
a = 27: OKAY = 0: Unload frmFHSWNdWk
End Sub

Private Sub cmdCREATE_Click()
If (BRID > 0) And ((NWrkType < 3) Or (BRID2 > 0)) Then
   Call BUILD: Call PutMSG("")
   Else
   Call PutMSG("Base Record(s) have not been selected...")
End If
End Sub

Private Sub cmdLIST_Click(Index As Integer)
Z = 2 * Index + 1
HOLDITEM = Val(NFValue(Z).TEXT)
Call frmFHSWList.SelectName(X)
If a <> 27 Then
   NFValue(Z).TEXT = fnx$(X)
   NFValue_LostFocus (Z)
End If
End Sub


Private Sub cmdOKAY_Click()
If LastNF > 0 Then
   NFValue_LostFocus (LastNF - 1)
   Else
   Unload frmFHSWNdWk
End If
End Sub

Private Sub Form_Load()
frmFHSWRPgm.Enabled = False
FormLoaded(1) = FormLoaded(1) Or NdWkForm
' Set Default Form Location on Screen
frmHEIGHT = frmFHSWNdWk.Height
frmWIDTH = frmFHSWNdWk.WIDTH
t1 = GetSetting("FHSWIN", "frmNDWK", "TOP", 0)
If t1 = 0 Then
   t00 = frmFHSWRPgm.TOP: w00 = frmFHSWRPgm.WIDTH
   l00 = frmFHSWRPgm.Left
   t1 = (t00 + 2770)
   l1 = l00 + (w00 - frmWIDTH) / 2: If l1 < 0 Then l1 = 1
   Else
   t1 = 1440: l1 = 1440
End If
frmFHSWNdWk.TOP = GetSetting("FHSWIN", "frmNDWK", "Top", t1)
frmFHSWNdWk.Left = GetSetting("FHSWIN", "frmNDWK", "Left", l1)
cboAlign.AddItem "None"
Select Case NWrkType
    Case 1: ' Ancestor
            cboAlign.AddItem "Father"
            cboAlign.AddItem "Mother"
    Case 2: ' Descendant
            cboAlign.AddItem "First Child"
            cboAlign.AddItem "Last Child"
    Case 3: ' Family Path
            cboAlign.AddItem "First Child"
            cboAlign.AddItem "Last Child"
End Select
cboAlign.AddItem "Between"
NFType.ListIndex = NWrkType - 1
NFValue(0).TEXT = FT$(8)
Call GetRULES: Call ShowRULES
Call CheckNodeFile
Call ChangeEnable
End Sub

Private Sub Form_Unload(Cancel As Integer)
FormLoaded(1) = FormLoaded(1) And (255 - NdWkForm)
SaveSetting "FHSWIN", "frmNDWK", "Top", frmFHSWNdWk.TOP
SaveSetting "FHSWIN", "frmNDWK", "Left", frmFHSWNdWk.Left
SaveSetting "FHSWIN", "frmNDWK", "Align", NFAlign
SaveSetting "FHSWIN", "frmNDWK", "MaxGL", NFMaxGL
SaveSetting "FHSWIN", "frmNDWK", "Siblings", NFSibling
frmFHSWRPgm.Enabled = True
frmFHSWRPgm.SetFocus
End Sub

Private Sub NFValue_GotFocus(Index As Integer)
HOLDITEM = NFValue(Index).TEXT: LastNF = Index + 1
NFValue(Index).SelStart = 0
NFValue(Index).SelLength = Len(NFValue(Index).TEXT)
End Sub


Private Sub NFValue_LostFocus(Index As Integer)
If NFValue(Index).TEXT = "" Then
   NFValue(Index).TEXT = HOLDITEM
   Else
   If HOLDITEM <> NFValue(Index).TEXT Then
      Select Case Index
        Case 0: ' Change File Name
                DataChanged = DataChanged Or 1
                If Not Canceled Then
                   NFDSN = NFValue(Index).TEXT
                   CheckNodeFile
                End If
        Case 1, 3: ' Change Base ID
                HOLDITEM = NFValue(Index).TEXT
                If HOLDITEM = 0 Then
                   If Index = 1 Then XBASE = BRID Else XBASE = BRID2
                   NFValue(Index).TEXT = XBASE
                   Else
                   If HOLDITEM > FF1Hdr.MAXID Then
                      NFValue(Index).TEXT = FF1Hdr.MAXID
                   End If
                End If
                DataChanged = DataChanged Or 2
                If Not Canceled Then
                   XBASE = NFValue(Index).TEXT
                   If Index = 1 Then BRID = XBASE Else BRID2 = XBASE
                   Call ShowBASE(XBASE, Index): ClearSTATS
                End If
                HOLDITEM = XBASE
        Case 8: ' MaxGL
                NFMaxGL = Val(NFValue(Index).TEXT)
                DataChanged = DataChanged Or 4: ClearSTATS
      End Select
   End If
End If
LastNF = 0
End Sub


