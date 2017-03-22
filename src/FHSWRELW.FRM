VERSION 4.00
Begin VB.Form frmFHSWRelW 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FHSWRELW - Relationship Work File"
   ClientHeight    =   3120
   ClientLeft      =   8325
   ClientTop       =   7965
   ClientWidth     =   7065
   Height          =   3525
   Left            =   8265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   Top             =   7620
   Width           =   7185
   Begin VB.CommandButton cmdLIST 
      Caption         =   "<"
      Height          =   285
      Left            =   1560
      TabIndex        =   37
      Top             =   600
      Width           =   255
   End
   Begin VB.ComboBox cboWFType 
      Height          =   315
      ItemData        =   "FHSWRELW.frx":0000
      Left            =   840
      List            =   "FHSWRELW.frx":000D
      TabIndex        =   8
      Text            =   "WFType"
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdBROWSE 
      Caption         =   "Browse"
      Height          =   315
      Left            =   4440
      TabIndex        =   11
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdCREATE 
      Caption         =   "Create File"
      Height          =   375
      Left            =   5400
      TabIndex        =   15
      Top             =   480
      Width           =   1575
   End
   Begin VB.ComboBox cboValue 
      Enabled         =   0   'False
      Height          =   315
      Index           =   3
      ItemData        =   "FHSWRELW.frx":0031
      Left            =   5520
      List            =   "FHSWRELW.frx":003B
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton cmdCANCEL 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   6120
      TabIndex        =   16
      Top             =   75
      Width           =   855
   End
   Begin VB.CommandButton cmdOKAY 
      Caption         =   "OKAY"
      Default         =   -1  'True
      Height          =   315
      Left            =   5400
      TabIndex        =   14
      Top             =   75
      Width           =   615
   End
   Begin VB.TextBox WFValue 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   10
      Left            =   5400
      TabIndex        =   36
      Top             =   2760
      Width           =   600
   End
   Begin VB.TextBox WFValue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   9
      Left            =   3720
      TabIndex        =   34
      Top             =   2760
      Width           =   600
   End
   Begin VB.TextBox WFValue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   7
      Left            =   3720
      TabIndex        =   29
      Top             =   2400
      Width           =   600
   End
   Begin VB.TextBox WFValue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   8
      Left            =   2280
      TabIndex        =   32
      Top             =   2760
      Width           =   450
   End
   Begin VB.TextBox WFValue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   6
      Left            =   2280
      TabIndex        =   27
      Top             =   2370
      Width           =   450
   End
   Begin VB.TextBox WFValue 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   315
      Index           =   5
      Left            =   2280
      MaxLength       =   5
      TabIndex        =   4
      Top             =   1680
      Width           =   450
   End
   Begin VB.TextBox WFValue 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   315
      Index           =   4
      Left            =   2280
      MaxLength       =   5
      TabIndex        =   1
      Top             =   1320
      Width           =   450
   End
   Begin VB.ComboBox cboValue 
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      ItemData        =   "FHSWRELW.frx":0048
      Left            =   5160
      List            =   "FHSWRELW.frx":0055
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1680
      Width           =   1212
   End
   Begin VB.ComboBox cboValue 
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      ItemData        =   "FHSWRELW.frx":0070
      Left            =   2880
      List            =   "FHSWRELW.frx":007D
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1680
      Width           =   1335
   End
   Begin VB.ComboBox cboValue 
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      ItemData        =   "FHSWRELW.frx":00A4
      Left            =   2880
      List            =   "FHSWRELW.frx":00B1
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1305
      Width           =   1335
   End
   Begin VB.TextBox WFValue 
      Appearance      =   0  'Flat
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
      Index           =   3
      Left            =   1920
      MaxLength       =   60
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   600
      Width           =   3255
   End
   Begin VB.TextBox WFValue 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   960
      MaxLength       =   5
      TabIndex        =   0
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox WFValue 
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
      Height          =   315
      Index           =   1
      Left            =   2760
      MaxLength       =   12
      TabIndex        =   10
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Adoptions?"
      Height          =   255
      Left            =   4560
      TabIndex        =   22
      Top             =   1320
      Width           =   855
   End
   Begin VB.Line Line4 
      X1              =   5280
      X2              =   5280
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   7080
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   7080
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label WFLabel 
      Caption         =   "Total:"
      Height          =   255
      Index           =   16
      Left            =   4800
      TabIndex        =   35
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label WFLabel 
      Caption         =   "Number:"
      Height          =   255
      Index           =   15
      Left            =   3000
      TabIndex        =   33
      Top             =   2790
      Width           =   735
   End
   Begin VB.Label WFLabel 
      Caption         =   "Number:"
      Height          =   255
      Index           =   12
      Left            =   3000
      TabIndex        =   28
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label WFLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Hi GL:"
      Height          =   255
      Index           =   14
      Left            =   1560
      TabIndex        =   31
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label WFLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Hi GL:"
      Height          =   255
      Index           =   11
      Left            =   1560
      TabIndex        =   26
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label WFLabel 
      Caption         =   "Descendants:"
      Height          =   255
      Index           =   13
      Left            =   360
      TabIndex        =   30
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label WFLabel 
      Caption         =   "Ancestors:"
      Height          =   255
      Index           =   10
      Left            =   360
      TabIndex        =   25
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label WFLabel 
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
      Index           =   9
      Left            =   240
      TabIndex        =   24
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label WFLabel 
      Caption         =   "Max GL:"
      Height          =   255
      Index           =   7
      Left            =   1560
      TabIndex        =   21
      Top             =   1695
      Width           =   735
   End
   Begin VB.Label WFLabel 
      Caption         =   "Max GL:"
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   19
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label WFLabel 
      Caption         =   "Order:"
      Height          =   375
      Index           =   8
      Left            =   4560
      TabIndex        =   23
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label WFLabel 
      Caption         =   "Descendants:"
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   20
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label WFLabel 
      Caption         =   "Ancestors:"
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   18
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label WFLabel 
      Caption         =   "Search Rules:"
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
      Left            =   240
      TabIndex        =   17
      Top             =   1050
      Width           =   1215
   End
   Begin VB.Label WFLabel 
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   750
   End
   Begin VB.Label WFLabel 
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
      Height          =   315
      Index           =   1
      Left            =   2280
      TabIndex        =   9
      Top             =   120
      Width           =   375
   End
   Begin VB.Label WFLabel 
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
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   150
      Width           =   555
   End
End
Attribute VB_Name = "frmFHSWRelW"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
DefInt A-Z
Dim FormProcess As Integer, WFTYPE As Integer
Dim WFOKAY, WFDSN As String
Dim AMAXGL, DMAXGL, ComboVALUE(3) As Integer
Dim HOLDITEM As Variant, DataChanged As Integer, Canceled As Integer
Dim LastWF As Integer
Sub RWRKBuild(WTYPE, BRID, FMTA, FMTSTAT)
500 ' Create work file
    'On Local Error GoTo 595
    Call FAMOpen("READ"): If Not OKAY Then GoTo 590
    'FMTBRID = FMTA
    'If FMTBRID = 0 Then
    '   Call FmtFIND("XXXXS010")
    '   If FOUND Then FMTBRID = CurFMT Else GoTo 590
    'End If
    'CurFMT = FMTBRID: Fmt = FMTTAB(CurFMT)
    'Call FmtFindFLD("BRID"): If Not FOUND Then GoTo 590
    'BFLD = FFLD.UTXT: RSet UTXT$(BFLD) = SNX$(BRID)
501 'WIN.MFMT = FMTBRID: Call MsgFORMAT: If A = 27 Then GoTo 590
    'X# = Val(UTXT$(BFLD)): If X# = 0 Then GoTo 590
    'If (1 > X#) Or (X# > FF1Hdr.MAXID) Then Call ErrBEEP(0): GoTo 501
    'X = X#
    Call FF1GetRec(BRID): If Not OKAY Then GoTo 590
    'BRID = X
    'Call RptOptGET("ADOP", X): AOK = (X = 1)
    'AGLMAX = 0: DGLMAX = 0: DSRCH = 0
    'If WTYPE <> 2 Then Call RptOptGET("MXAL", AGLMAX)
    'If WTYPE > 1 Then Call RptOptGET("MXDL", DGLMAX)
    'If DGLMAX > 0 Then
    '   Call RptOptGET("ORDR", DSRCH)
    '   If (DSRCH < 1) Or (DSRCH > 2) Then DSRCH = 1
    '   Call RptOptGET("DLNS", DSEXCHK)
    '   If (DSEXCHK < 1) Or (DSEXCHK > 3) Then DSEXCHK = 1
    '   DSEXCHK = DSEXCHK - 1
    'End If
    'If AGLMAX > 0 Then
    '   Call RptOptGET("ALNS", ASEXCHK)
    '   If (ASEXCHK < 1) Or (ASEXCHK > 3) Then ASEXCHK = 1
    '   ASEXCHK = ASEXCHK - 1
    'End If
    'WFHDR.AMAXGL = AGLMAX: WFHDR.DMAXGL = DGLMAX
    AGLMAX = WFHDR.AMAXGL: DGLMAX = WFHDR.DMAXGL
    WFHDR.AMAXLV = Chr$(0): WFHDR.DMAXLV = Chr$(0)
    WFHDR.HIRGL = Chr$(0): HIRGL = 0
    WFHDR.NUMANC = 0: WFHDR.NUMREL = 0
    WFRULES = Asc(WFHDR.WFRULES)
    AOK = ((WFRULES And 64) > 0)
    DSRCH = (WFRULES And 3)
    ASEXCHK = ((WFRULES And 12) / 4): DSEXCHK = ((WFRULES And 48) / 16)
    GoSub X80
    Call RWRKOpen("NEW", WTYPE, 0): If Not OKAY Then GoTo 590
    'WFHDR.WFRULES = Chr$(DSRCH + 4 * ASEXCHK + 16 * DSEXCHK + 64 * Abs(AOK))
    WFHDR.WFRULES = Chr$(WFRULES)
    WFHDR.TOTREC = 2
    Put #5, 1, WFHDR
    WFANCREC.GACNT = 1
    WFANCREC.LINEAGE = 1: WFANCREC.ANCID = BRID
    WFANCREC.FIRSTGL = 0: WFANCREC.ANCCH = 0: WFANCREC.FIRSTANC = 0
    WFANCREC.PRVDUP = 0: WFANCREC.NXTDUP = 0
    Put #5, , WFANCREC
    ReDim PTRTAB&(FF1Hdr.MAXID)
    WFGLREC.ANCPTR = Len(WFHDR): WFGLREC.AGL = 0: WFGLREC.DGL = 0
    WFGLREC.PRIOR = 0: WFGLREC.NEXT = 0
    WFGLREC.CLINES = 0: WFGLREC.TOTREC = 1
    HIWRKRBA& = Len(WFHDR) + Len(WFANCREC)

505 ' Create WFAGLREC
    If WFGLREC.TOTREC = 0 Then GoTo 540
    WFHDR.AMAXLV = Chr$(WFGLREC.AGL)
    ANCEND& = HIWRKRBA&
    HIWRKRBA& = HIWRKRBA& + Len(WFGLREC)
    Put #5, ANCEND& + 1, WFGLREC: WFHDR.TOTREC = WFHDR.TOTREC + 1
    If WFGLREC.PRIOR <> 0 Then
       X& = WFGLREC.PRIOR
       Get #5, X& + 1, WFGLREC: WFGLREC.NEXT = ANCEND&
       Put #5, X& + 1, WFGLREC
       Get #5, ANCEND& + 1, WFGLREC
    End If
    WFGLREC.PRIOR = ANCEND&
    ANCRBA& = WFGLREC.ANCPTR: WFGLREC.ANCPTR = 0
    WFGLREC.CLINES = 0: WFGLREC.TOTREC = 0: WFGLREC.NEXT = 0

510 ' Locate next generation of Ancestors
    WFGLREC.AGL = WFGLREC.AGL + 1
    If WFGLREC.AGL > WFHDR.AMAXGL Then GoTo 540
    SysVAR.NVAR1 = WFGLREC.AGL
    'Call PutMSG("XXXXM065")
    'CA$ = INKEY$: If CA$ <> "" Then If Asc(CA$) = 27 Then GoTo 585
    GoSub X80

520 ' Record parents of next Ancestor in this Generation
    Get #5, ANCRBA& + 1, WFANCREC
    dLINEAGE# = WFANCREC.LINEAGE: FIRSTANC& = 0
    If WFANCREC.ANCID < 0 Then GoTo 530
    WFANCREC.ANCCH = ANCRBA&: ExtAdopt = ((WFANCREC.PSTAT And 5) > 0)
    Call FF1GetRec(WFANCREC.ANCID): C = 1
    If (WFGLREC.AGL > 1) And (ASEXCHK <> 0) And (FF1Rec.SEX <> FGENDR$(ASEXCHK)) Then GoTo 530
    ANCID = FF1Rec.FID: I = 0: S = Asc(FF1Rec.STATUS) And 15
    If ((ANCID <> 0) And (AOK Or ((S And 1) = 0))) Then GoSub 535
    ANCID = FF1Rec.MID: I = 1: S = S \ 4
    If ((ANCID <> 0) And (AOK Or ((S And 1) = 0))) Then GoSub 535
530 ' GO TO Next Ancestor Record
    If FIRSTANC& <> 0 Then
       Get #5, ANCRBA& + 1, WFANCREC
       WFANCREC.FIRSTANC = FIRSTANC&
       Put #5, ANCRBA& + 1, WFANCREC
    End If
    ANCRBA& = ANCRBA& + Len(WFANCREC)
    If ANCRBA& >= ANCEND& Then GoTo 505 ' finished this GEN, process Next GEN
    GoTo 520                        ' process Next ancestor, this GEN

535 ' Create Ancestor Record
    If PTRTAB&(ANCID) = 0 Then
       WFANCREC.ANCID = ANCID: WFHDR.NUMREL = WFHDR.NUMREL + 1
       Else
       WFANCREC.ANCID = -ANCID
    End If
    WFANCREC.LINEAGE = 2 * dLINEAGE# + I
    WFANCREC.PSTAT = ((S And 3) - 4 * ExtAdopt + 16 * (I + 1))
    WFANCREC.PRVDUP = PTRTAB&(ANCID)
    WFHDR.NUMANC = WFHDR.NUMANC + 1
    If C > 0 Then WFGLREC.CLINES = WFGLREC.CLINES + 1: C = 0
    WFGLREC.TOTREC = WFGLREC.TOTREC + 1
    WFANCREC.GACNT = WFGLREC.TOTREC: WFANCREC.FIRSTANC = 0
    WFANCREC.NXTDUP = 0
    Put #5, HIWRKRBA& + 1, WFANCREC: WFHDR.TOTREC = WFHDR.TOTREC + 1
    X& = PTRTAB&(ANCID)
    If X& <> 0 Then
       Get #5, X& + 1, WFANCREC
       WFANCREC.NXTDUP = HIWRKRBA&
       Put #5, X& + 1, WFANCREC
       Get #5, HIWRKRBA& + 1, WFANCREC
    End If
    PTRTAB&(ANCID) = HIWRKRBA&
    If WFGLREC.ANCPTR = 0 Then WFGLREC.ANCPTR = HIWRKRBA&
    If (FIRSTANC& = 0) Then FIRSTANC& = HIWRKRBA&
    HIWRKRBA& = HIWRKRBA& + Len(WFANCREC)
    Return

540 ' Finished Locating Ancestors
    If DGLMAX = 0 Then GoTo 580
    WFAGLREC.NEXT = Len(WFHDR) + Len(WFANCREC): WFHDR.NUMREL = 0
    ReDim PTRTAB&(FF1Hdr.MAXID)
    Dim GLS As WRKGLStats
    'If WFHDR.AMAXGL > 0 Then Call PutMSG("XXXXM066")

550 ' Look for Descendants of Next Level of Ancestors
    If WFAGLREC.NEXT = 0 Then GoTo 580
    GoSub X80
    Get #5, WFAGLREC.NEXT + 1, WFAGLREC
    AGL = -WFAGLREC.AGL: ANCNUM = 0
    ANCRBA& = WFAGLREC.ANCPTR - Len(WFANCREC)

555 ' Locate Descendants of Next Ancestor
    ANCNUM = ANCNUM + 1
    If ANCNUM > WFAGLREC.TOTREC Then GoTo 550
    ANCRBA& = ANCRBA& + Len(WFANCREC)
    Get #5, ANCRBA& + 1, WFANCREC
    S = (WFANCREC.PSTAT And 3): ExtAdopt = ((WFANCREC.PSTAT And 5) > 0)
    RELID = Abs(WFANCREC.ANCID)
    ReDim GLSTATS(WFHDR.DMAXGL + 1) As WRKGLStats
    DGL = 0: xPARPTR& = 0

560 ' Process Descendant RELID in Generation DGL
    GLS = GLSTATS(DGL)
    Call FF1GetRec(RELID)
    If (DGL > 0) Then
       S = Asc(FF1Rec.STATUS) And 15
       If (Parent = FF1Rec.MID) Then S = S \ 4
       If (Not AOK) And ((S And 1) > 0) Then FF1Rec.OLDCH = 0: GoTo 566
    End If

    ' First Create GLREC if this is First RELID in Generation DGL
    If GLS.GLPTR = 0 Then
       GLS.GLPTR = HIWRKRBA&
       If DGL > 0 Then
          X& = GLSTATS(DGL - 1).GLPTR
          Get #5, X& + 1, WFGLREC: WFGLREC.NEXT = GLS.GLPTR
          Put #5, X& + 1, WFGLREC
          WFGLREC.PRIOR = GLSTATS(DGL - 1).GLPTR
          Else
          WFANCREC.FIRSTGL = GLS.GLPTR
          Put #5, ANCRBA& + 1, WFANCREC
          WFGLREC.PRIOR = 0
       End If
       WFGLREC.ANCPTR = ANCRBA&
       WFGLREC.AGL = AGL: WFGLREC.DGL = DGL
       WFGLREC.NEXT = 0
       WFGLREC.CLINES = 0: WFGLREC.TOTREC = 0
       Put #5, HIWRKRBA& + 1, WFGLREC: WFHDR.TOTREC = WFHDR.TOTREC + 1
       HIWRKRBA& = HIWRKRBA& + Len(WFGLREC)
    End If

    ' Create Relative Record for RELID in DGL
    If (DGL > 0) And (GLS.CHNUM = 0) Then
       GLS.CLINES = GLS.CLINES + 1
       Get #5, xPARPTR& + 1, WFRELREC
       WFRELREC.OLDCH = HIWRKRBA&: ExtAdopt = ((WFRELREC.CSTAT And 5) > 0)
       Put #5, xPARPTR& + 1, WFRELREC
    End If
    If GLS.RELREC <> 0 Then
       Get #5, GLS.RELREC + 1, WFRELREC: WFRELREC.GLNEXT = HIWRKRBA&
       Put #5, GLS.RELREC + 1, WFRELREC
    End If
    GLS.RELID = RELID
    GLS.RELREC = HIWRKRBA&
    GLS.CHNUM = GLS.CHNUM + 1
    GLS.COUNT = GLS.COUNT + 1
    GLSTATS(DGL) = GLS
    WFRELREC.GLPTR = GLS.GLPTR
    WFRELREC.PARPTR = xPARPTR&
    WFRELREC.id = RELID
    WFRELREC.NREF = 0: WFRELREC.GLNEXT = 0
    WFRELREC.CHNUM = Chr$(GLS.CHNUM)
    WFRELREC.CSTAT = S - 4 * ExtAdopt
    WFRELREC.NUMCH = Chr$(0)
    If PTRTAB&(RELID) = 0 Then
       If (DGL < WFHDR.DMAXGL) And ((((DSEXCHK = 0) Or (FF1Rec.SEX = FGENDR$(DSEXCHK))))) Then
          WFRELREC.OLDCH = FF1Rec.OLDCH
          Else
          WFRELREC.OLDCH = 0
       End If
       Else
       WFRELREC.id = -WFRELREC.id
       WFRELREC.OLDCH = PTRTAB&(RELID)
       WFRELREC.NUMCH = Chr$(1)
    End If
    Put #5, HIWRKRBA& + 1, WFRELREC: WFHDR.TOTREC = WFHDR.TOTREC + 1
    If PTRTAB&(RELID) = 0 Then
       If (AGL <> 0) Or (DGL <> 0) Then WFHDR.NUMREL = WFHDR.NUMREL + 1
       Else
       X& = PTRTAB&(RELID)
       Get #5, X& + 1, WFRELREC: WFRELREC.NREF = HIWRKRBA&
       Put #5, X& + 1, WFRELREC
       If WFRELREC.id < 0 Then
          xX = Asc(WFRELREC.NUMCH)
          Get #5, HIWRKRBA& + 1, WFRELREC
          WFRELREC.NUMCH = Chr$(xX + 1)
          Put #5, HIWRKRBA& + 1, WFRELREC
          Else
          Get #5, HIWRKRBA& + 1, WFRELREC
       End If
    End If
    PTRTAB&(RELID) = HIWRKRBA&
    HIWRKRBA& = HIWRKRBA& + Len(WFRELREC)
    If DGL > DMAXLV Then DMAXLV = DGL: WFHDR.DMAXLV = Chr$(DGL)
    If (AGL + DGL) > HIRGL Then HIRGL = AGL + DGL: WFHDR.HIRGL = Chr$(HIRGL)

565 ' Select Next RELID to Process
    CA$ = INKEY$
    If CA$ <> "" Then If Asc(CA$) = 27 Then GoTo 585 Else GoSub X80
    If (DSRCH = 2) Or (DGL = 0) Then NEXTPAR& = GLS.RELREC: GoTo 567

566 ' Look for Younger Sibling of Child
    If Parent = FF1Rec.FID Then RELID = FF1Rec.FCH Else RELID = FF1Rec.MCH
    If Parent <> WFRELREC.id Then
       Get #5, xPARPTR& + 1, WFRELREC
       ExtAdopt = ((WFRELREC.CSTAT And 5) > 0)
    End If
    If RELID > 0 Then GoTo 560
    'IF GLS.CHNUM > 0 THEN
       WFRELREC.NUMCH = Chr$(GLS.CHNUM)
       If GLS.CHNUM = 0 Then WFRELREC.OLDCH = 0
       Put #5, xPARPTR& + 1, WFRELREC
       GLS.CHNUM = 0
    'END IF
    GLSTATS(DGL) = GLS
    DGL = DGL - 1
    GLS = GLSTATS(DGL)
    If (DSRCH = 2) Then GoTo 569
    Get #5, GLS.RELREC + 1, WFRELREC
    NEXTPAR& = WFRELREC.GLNEXT

567 ' Look for a parent in DGL - FF1REC should be GLS.RELID Record
    While NEXTPAR& <> 0
       If GLS.RELREC <> NEXTPAR& Then
          GLS.RELREC = NEXTPAR&
          Get #5, NEXTPAR& + 1, WFRELREC
       End If
       If WFRELREC.id > 0 Then
          RELID = WFRELREC.OLDCH
          If RELID > 0 Then
             Parent = WFRELREC.id: xPARPTR& = GLS.RELREC
             GLS.RELID = Parent
             GLSTATS(DGL) = GLS
             DGL = DGL + 1
             GLSTATS(DGL).CHNUM = 0
             GoTo 560
          End If
       End If
       NEXTPAR& = WFRELREC.GLNEXT
    Wend

    GLSTATS(DGL) = GLS
    If (DSRCH = 2) Then      ' Family Search
       If DGL = 0 Then GoTo 570 ' STOP - if we've returned to GL=0
       GoTo 566              ' ELSE Look for a Sibling of Childless Relative
    End If

    ' Gen Search - Process Next Gen Level
    If GLSTATS(DGL + 1).COUNT = 0 Then GoTo 570
    DGL = DGL + 1: GLS = GLSTATS(DGL)
    NEXTPAR& = GLS.GLPTR + Len(WFGLREC)  ' First RELREC in GL follows GLPTR
    GLS.RELREC = 0
    If AGLMAX = 0 Then
       GoSub X80
       SysVAR.NVAR1 = DGL
       'Call PutMSG("XXXXM067")
    End If
    GoTo 567

569 ' FAM Search - Look for Sibling of Parent
    If DGL < 1 Then GoTo 570
    If DGL = 1 Then GoSub X80
    xPARPTR& = GLSTATS(DGL - 1).RELREC
    Parent = GLSTATS(DGL - 1).RELID
    Call FF1GetRec(GLS.RELID)
    GoTo 566

570 ' GEN Search - Update WFGLRECs for This Ancestor's Descendants
    DGL = 0: GLS = GLSTATS(0)
    While GLS.GLPTR <> 0
       Get #5, GLS.GLPTR + 1, WFGLREC
       WFGLREC.CLINES = GLS.CLINES
       WFGLREC.TOTREC = GLS.COUNT
       Put #5, GLS.GLPTR + 1, WFGLREC
       DGL = DGL + 1: GLS = GLSTATS(DGL)
       If DGL > WFHDR.DMAXGL Then GLS.GLPTR = 0
    Wend
    GoTo 555

580 ' WorkFile is Complete
    WFHDR.WFVER = Chr$(2): GoSub X80
    'Call PutMSG("XXXXM068")

585 Put #5, 1, WFHDR

590 ' Return to Calling Program
    On Error GoTo 0
    Call FamCLOSE: ClearMSG
    Exit Sub

595 ' Error Writing to Work File
    Call ERRMessage("XXXXM032", ERR, ERL)
    Resume 590

X80: ' Display Work File Information
    Call frmFHSWRelW.ShowStats
    'CurFMT = FMTSTAT: Fmt = FMTTAB(CurFMT): X = Fmt.BGNFLD
    'While X > 0
    '   FFLD = FLDTAB(X)
    '   If FFLD.UTXT > 0 Then
    '      XX$ = UTXT$(FFLD.UTXT)
    '      If WFHDR.NUMREL < 1 Then
    '         Select Case FFLD.NAME
    '            Case "BRID": LSet XX$ = Str$(BRID)
    '            Case "NAME": LSet XX$ = FMTNAME$(1)
    '            Case "WFDT": LSet XX$ = ""
    '            Case "RULA": LSet XX$ = SNX$(((Asc(WFHDR.WFRULES) And 12) \ 4) + 1)
    '            Case "RULD": LSet XX$ = SNX$(((Asc(WFHDR.WFRULES) And 48) \ 16) + 1)
    '            Case "RULO": LSet XX$ = SNX$(Asc(WFHDR.WFRULES) And 3)
    '         End Select
    '      End If
    '      Select Case FFLD.NAME
    '         Case "WFSZ": LSet XX$ = Str$(WFHDR.TOTREC)
    '         Case "HIAL": LSet XX$ = Str$(Asc(WFHDR.AMAXLV))
    '         Case "HIDL": LSet XX$ = Str$(Asc(WFHDR.DMAXLV))
    '         Case "NUMA": LSet XX$ = Str$(WFHDR.NUMANC)
    '         Case "NUMR": LSet XX$ = Str$(WFHDR.NUMREL)
    '      End Select
    '      LSet UTXT$(FFLD.UTXT) = XX$
    '   End If
    '   X = FLDCHN(X).FWD
    'Wend
    'Call WinDisplayDATA(0)
    Return

End Sub


Public Sub BUILD(X)
  If (X = 0) Or (X > 3) Then
     X = 3
     cboWFType.Enabled = True
     Else
     cboWFType.Enabled = False
  End If
  WFTYPE = X
  cboWFType.ListIndex = WFTYPE - 1
  ChangeEnable (True)
  WFDSN = ""
  Call DefaultRules
  Call CheckWrkFile
  mnuFHSWMain.ChangeEnable
  Show 1
End Sub
Private Sub ChangeEnable(STATUS)
WFValue(1).Enabled = STATUS
WFValue(2).Enabled = STATUS
WFValue(4).Enabled = STATUS
WFValue(5).Enabled = STATUS
For X = 0 To 3: cboValue(X).Enabled = STATUS: Next X
Select Case WFTYPE
   Case 1: WFValue(5).Enabled = False
           cboValue(1).Enabled = False
           cboValue(2).Enabled = False
   Case 2: WFValue(4).Enabled = False
           cboValue(0).Enabled = False
End Select
End Sub


Private Sub SetupRules()
   WFHDR.AMAXGL = Val(WFValue(4))
   WFHDR.DMAXGL = Val(WFValue(5))
   X = cboValue(2).ListIndex
   X = X + 4 * cboValue(0).ListIndex
   X = X + 16 * cboValue(1).ListIndex
   X = X + 64 * cboValue(3).ListIndex
   WFHDR.WFRULES = Chr$(Abs(X))
End Sub

Private Sub ShowBASE()
WFValue(3) = ""
If Not (WFBASE = 0) Then
   Call FAMOpen("READ")
   If OKAY Then
      Call FF1GetRec(WFBASE)
      If OKAY Then WFValue(3) = FMTNAME(1)
   End If
End If
Call FamCLOSE
End Sub

Private Sub CheckWrkFile()
X = WFTYPE + 4
If WFDSN = "" Then
   WFDSN = FT$(X)
   WFValue(1).TEXT = WFDSN
End If
FT$(X) = WFDSN
Call RWRKOpen("READQ", WFTYPE, 0)
WFOKAY = OKAY
If OKAY Then
   WFDSN = FT$(X)
   WFBASE = WFANCREC.ANCID: WFValue(2).TEXT = WFBASE
   Call ShowBASE: ShowRULES: ShowStats
   Else
   FT$(X) = WFDSN
   Call ClearRULES: ClearSTATS
End If
Call FamCLOSE
DataChanged = 0
End Sub



Private Sub ClearRULES()
   WFValue(4) = Str$(AMAXGL)
   WFValue(5) = Str$(DMAXGL)
   Select Case WFTYPE
      Case 1: WFValue(5) = Str$(0)
      Case 2: WFValue(4) = Str$(0)
   End Select
   For X = 0 To 3
      cboValue(X).ListIndex = ComboVALUE(X)
   Next X
End Sub


Private Sub ClearSTATS()
   RelFileOKAY = 0
   For X = 6 To 10: WFValue(X) = "": Next X

End Sub


Private Sub ShowRULES()
   WFValue(4) = Abs(Str$(WFHDR.AMAXGL))
   WFValue(5) = Str$(WFHDR.DMAXGL)
   X = Asc(WFHDR.WFRULES)
   cboValue(0).ListIndex = ((X And 12) / 4)
   'Select Case X And 12
   '   Case 0: cboValue(0) = "All Lines"
   '   Case 4: cboValue(0) = "Male Only"
   '   Case 8: cboValue(0) = "Female Only"
   'End Select
   cboValue(1).ListIndex = ((X And 48) / 16)
   'Select Case X And 48
   '   Case 0: cboValue(1) = "All Lines"
   '   Case 16: cboValue(1) = "Male Only"
   '   Case 32: cboValue(1) = "Female Only"
   'End Select
   cboValue(2).ListIndex = (X And 3)
   'Select Case X And 3
   '   Case 0: cboValue(2) = ""
   '   Case 1: cboValue(2) = "Generations"
   '   Case 2: cboValue(2) = "Family"
   'End Select
   cboValue(3).ListIndex = ((X And 64) / 64)
   'Select Case X And 64
   '   Case 0: cboValue(3) = "No"
   '   Case 64: cboValue(3) = "Yes"
   'End Select
   For X = 0 To 3
       ComboVALUE(X) = cboValue(X).ListIndex
   Next X
End Sub

Public Sub ShowStats()
   WFValue(6) = Str$(Asc(WFHDR.AMAXLV))
   WFValue(7) = Str$(WFHDR.NUMANC)
   WFValue(8) = Str$(Asc(WFHDR.DMAXLV))
   If WFValue(8) > 0 Then
      WFValue(9) = Str$(WFHDR.NUMREL - WFHDR.NUMANC)
      Else
      WFValue(9) = 0
   End If
   WFValue(10) = Str$(WFHDR.NUMREL)
End Sub


Private Sub cboValue_GotFocus(Index As Integer)
HOLDITEM = cboValue(Index).ListIndex
End Sub


Private Sub cboValue_LostFocus(Index As Integer)
If HOLDITEM <> cboValue(Index).ListIndex Then
   DataChanged = DataChanged Or 4
End If
End Sub


Private Sub cmdBROWSE_Click()
HOLDITEM = WFValue(1).TEXT
Select Case WFTYPE
   Case 1: xX$ = "ANCESTOR"
   Case 2: xX$ = "DESCENDANT"
   Case 3: xX$ = "RELATIVE"
End Select
Call frmFHSWFile.LOCATE(xX$, "FULL")
If OKAY Then
   WFValue(1).TEXT = FT$(WFTYPE + 4)
   If WFValue(1).TEXT <> HOLDITEM Then
      Call WFValue_LostFocus(1)
   End If
End If
End Sub

Private Sub cmdCANCEL_Click()
OKAY = 0: Canceled = 0
If DataChanged = 0 Then
   a = 27: Unload frmFHSWRelW
   Else
   WFValue(1) = WFDSN
   WFValue(2) = WFBASE
   If Not (WFTYPE = 2) Then
      WFValue(4) = Str$(AMAXGL)
      cboValue(0).ListIndex = ComboVALUE(0)
   End If
   If Not (WFTYPE = 1) Then
      WFValue(5) = Str$(DMAXGL)
      cboValue(1).ListIndex = ComboVALUE(1)
      cboValue(2).ListIndex = ComboVALUE(2)
      End If
   cboValue(3).ListIndex = ComboVALUE(3)
   DataChanged = 0
End If
End Sub



Private Sub cmdCREATE_Click()
If LastWF <> 0 Then
   Call WFValue_LostFocus(LastWF)
End If
If WFBASE = 0 Then
   Call ErrBEEP(0)
   Else
   If DataChanged > 0 Then cmdOKAY_Click
   SetupRules
   frmFHSWRelW.MousePointer = 11
   Call RWRKBuild(WFTYPE, WFBASE, 0, 0)
   frmFHSWRelW.MousePointer = 0
   WFOKAY = OKAY: WFANCREC.ANCID = WFBASE
End If
LastWF = 0: DataChanged = 0
End Sub

Private Sub cmdLIST_Click()
HOLDITEM = Val(WFValue(2).TEXT)
Call frmFHSWList.SelectName(X)
If a <> 27 Then
   WFValue(2).TEXT = fnx$(X)
   WFValue_LostFocus (2)
End If
End Sub
 
Private Sub cmdOKAY_Click()
If LastWF <> 0 Then WFValue_LostFocus (LastWF)
OKAY = -1: a = 0
If DataChanged = 0 Then
   Unload frmFHSWRelW
   Else
   If (DataChanged And 4) Then
      'Change RULES
      AMAXGL = Val(WFValue(4))
      DMAXGL = Val(WFValue(5))
      For X = 0 To 3
          ComboVALUE(X) = cboValue(X).ListIndex
      Next X
      Call ClearSTATS
   End If
   DataChanged = 0: LastWF = 0
End If
End Sub






Private Sub Form_Load()
FormLoaded(1) = FormLoaded(1) Or RelWForm
' Set Default Form Location on Screen
frmHEIGHT = frmFHSWRelW.Height
frmWIDTH = frmFHSWRelW.WIDTH
t1 = GetSetting("FHSWIN", "frmRELW", "TOP", 0)
If t1 = 0 Then
   t00 = mnuFHSWMain.TOP: w00 = mnuFHSWMain.WIDTH
   l00 = mnuFHSWMain.Left
   t1 = (t00 + 1885)
   l1 = l00 + (w00 - frmWIDTH) / 2: If l1 < 0 Then l1 = 1
   Else
   t1 = 1440: l1 = 1440
End If
frmFHSWRelW.TOP = GetSetting("FHSWIN", "frmRELW", "Top", t1)
frmFHSWRelW.Left = GetSetting("FHSWIN", "frmRELW", "Left", l1)
End Sub
Private Sub DefaultRules()
AMAXGL = 53: DMAXGL = 99
cboValue(0).ListIndex = 0
cboValue(1).ListIndex = 0
cboValue(2).ListIndex = 2
cboValue(3).ListIndex = 0
For X = 0 To 3
   ComboVALUE(X) = cboValue(X).ListIndex
Next X
End Sub


Private Sub Form_Unload(Cancel As Integer)
FormLoaded(1) = FormLoaded(1) And (255 - RelWForm)
SaveSetting "FHSWIN", "frmRELW", "Top", frmFHSWRelW.TOP
SaveSetting "FHSWIN", "frmRELW", "Left", frmFHSWRelW.Left
End Sub



Private Sub WFValue_Change(Index As Integer)
Select Case Index
   Case Is < 3: DataChanged = DataChanged Or Index
   Case Is > 3: DataChanged = DataChanged Or 4
End Select
LastWF = Index
End Sub

Private Sub WFValue_GotFocus(Index As Integer)
HOLDITEM = WFValue(Index).TEXT
WFValue(Index).SelStart = 0
WFValue(Index).SelLength = Len(WFValue(Index).TEXT)
End Sub


Private Sub WFValue_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub WFValue_LostFocus(Index As Integer)
If WFValue(Index).TEXT = "" Then
   WFValue(Index).TEXT = HOLDITEM
   Else
   If HOLDITEM <> WFValue(Index).TEXT Then
      Select Case Index
        Case 1: DataChanged = DataChanged Or 1
                If Not Canceled Then
                   WFDSN = WFValue(Index).TEXT
                   CheckWrkFile
                End If
        Case 2: ' Change Base ID
                HOLDITEM = WFValue(Index).TEXT
                If HOLDITEM = 0 Then
                   WFValue(Index).TEXT = WFBASE
                   Else
                   If HOLDITEM > FF1Hdr.MAXID Then
                      WFValue(Index).TEXT = FF1Hdr.MAXID
                   End If
                End If
                DataChanged = DataChanged Or 2
                If Not Canceled Then
                   WFBASE = WFValue(Index).TEXT
                   Call ShowBASE: ClearSTATS
                End If
                HOLDITEM = WFBASE
        Case 4: DataChanged = DataChanged Or 4: ClearSTATS
        Case 5: DataChanged = DataChanged Or 4: ClearSTATS
      End Select
   End If
End If
LastWF = 0
End Sub


