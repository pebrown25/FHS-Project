VERSION 4.00
Begin VB.Form frmFHSWRCLC 
   Caption         =   "FHSWRCLC - Relationship Calculator"
   ClientHeight    =   4200
   ClientLeft      =   1965
   ClientTop       =   2595
   ClientWidth     =   5430
   Height          =   4605
   Left            =   1905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   Top             =   2250
   Width           =   5550
   Begin VB.CommandButton cmdLAST 
      Caption         =   ">>"
      Height          =   255
      Left            =   3480
      TabIndex        =   8
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cmdFIRST 
      Caption         =   "<<"
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cmdRETURN 
      Caption         =   "Return"
      Height          =   255
      Left            =   4440
      TabIndex        =   9
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdNEXT 
      Caption         =   ">"
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cmdPREV 
      Caption         =   "<"
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cmdSEARCH 
      Caption         =   "Search"
      Default         =   -1  'True
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtValue 
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
      Index           =   12
      Left            =   2880
      MultiLine       =   -1  'True
      TabIndex        =   37
      Top             =   3840
      Width           =   2415
   End
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Index           =   11
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   35
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox txtValue 
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
      Index           =   10
      Left            =   2880
      MultiLine       =   -1  'True
      TabIndex        =   33
      Top             =   3480
      Width           =   2415
   End
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Index           =   9
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   31
      Top             =   3480
      Width           =   615
   End
   Begin VB.TextBox txtValue 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   315
      Index           =   8
      Left            =   3600
      MultiLine       =   -1  'True
      TabIndex        =   29
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox txtValue 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   315
      Index           =   7
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   27
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox txtValue 
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
      Index           =   6
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   25
      Top             =   2760
      Width           =   3975
   End
   Begin VB.TextBox txtValue 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   315
      Index           =   5
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   24
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox txtValue 
      Enabled         =   0   'False
      Height          =   315
      Index           =   4
      Left            =   4800
      MultiLine       =   -1  'True
      TabIndex        =   22
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox txtValue 
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
      Index           =   3
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   20
      Top             =   1920
      Width           =   2415
   End
   Begin VB.TextBox txtValue 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   1920
      Width           =   375
   End
   Begin VB.ComboBox cboTYPE 
      Height          =   315
      ItemData        =   "FHSWRCLC.frx":0000
      Left            =   4080
      List            =   "FHSWRCLC.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox cboRULE 
      Height          =   315
      ItemData        =   "FHSWRCLC.frx":0036
      Left            =   1800
      List            =   "FHSWRCLC.frx":0043
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtValue 
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
      Left            =   1920
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   480
      Width           =   3375
   End
   Begin VB.CommandButton cmdLIST 
      Caption         =   "<"
      Height          =   315
      Index           =   1
      Left            =   1440
      TabIndex        =   3
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox txtRID 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   720
      MaxLength       =   5
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtValue 
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
      Index           =   0
      Left            =   1920
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   120
      Width           =   3375
   End
   Begin VB.CommandButton cmdLIST 
      Caption         =   "<"
      Height          =   315
      Index           =   0
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtRID 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   720
      MaxLength       =   5
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   5400
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label 
      Caption         =   "Lineage:"
      Height          =   255
      Index           =   12
      Left            =   2040
      TabIndex        =   36
      Top             =   3885
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "AGL2:"
      Height          =   255
      Index           =   11
      Left            =   840
      TabIndex        =   34
      Top             =   3885
      Width           =   495
   End
   Begin VB.Label Label 
      Caption         =   "Lineage:"
      Height          =   255
      Index           =   10
      Left            =   2040
      TabIndex        =   32
      Top             =   3525
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "AGL1:"
      Height          =   255
      Index           =   9
      Left            =   840
      TabIndex        =   30
      Top             =   3525
      Width           =   495
   End
   Begin VB.Label Label 
      Caption         =   "Died:"
      Height          =   255
      Index           =   8
      Left            =   3120
      TabIndex        =   28
      Top             =   3165
      Width           =   495
   End
   Begin VB.Label Label 
      Caption         =   "Born: "
      Height          =   255
      Index           =   7
      Left            =   840
      TabIndex        =   26
      Top             =   3165
      Width           =   495
   End
   Begin VB.Label Label 
      Caption         =   "Common Ancestor:"
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
      Index           =   6
      Left            =   120
      TabIndex        =   23
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label 
      Caption         =   "RGL="
      Height          =   255
      Index           =   5
      Left            =   4320
      TabIndex        =   21
      Top             =   1965
      Width           =   495
   End
   Begin VB.Label Label 
      Caption         =   "Relationship: "
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
      TabIndex        =   18
      Top             =   1965
      Width           =   1095
   End
   Begin VB.Label Label 
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
      Index           =   1
      Left            =   3480
      TabIndex        =   17
      Top             =   1470
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "Relationship Rule:"
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
      TabIndex        =   16
      Top             =   1470
      Width           =   1575
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   5400
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5400
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label 
      Caption         =   "ID #2:"
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
      Left            =   120
      TabIndex        =   14
      Top             =   510
      Width           =   615
   End
   Begin VB.Label Label 
      Caption         =   "ID #1:"
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
      TabIndex        =   12
      Top             =   150
      Width           =   615
   End
End
Attribute VB_Name = "frmFHSWRCLC"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
DefInt A-Z
Dim RID1, LastRID1, RID2, LastRID2, RLCT
Dim AGL, AGL1, RGL
Dim P2, AMAX1, AMAX2, X2, Y2, Z2, SEX2$
Dim CACNT, CAMAX, CATAB()
Dim RELATE(), ATAB(), ARGL1(), CARGL2()
Dim ALIN1() As Double, ALIN2() As Double, CALIN2() As Double
Dim RELRULE, RELTYPE, PATERNAL, MATERNAL
Dim RELTABOK, RELWORK$, SearchComplete
Dim frmHEIGHT, frmWIDTH
Private Sub GetRULES()
1510 ' GET RELRULES
     RELRULE = (ENV.RELRULES And 3)
     RELTYPE = ((ENV.RELRULES And 12) / 4)
     MATERNAL = ((RELTYPE And 1) > 0)
     PATERNAL = ((RELTYPE And 2) > 0)
     cboRULE.ListIndex = RELRULE
     cboTYPE.ListIndex = -MATERNAL - 2 * PATERNAL
End Sub

Public Sub RelCALC(X)
RID1 = X: If RID1 > 0 Then txtRID(0) = RID1
RID2 = 0: RLCT = 0
RELRULE = ENV.RELRULES And 3: RELTYPE = (ENV.RELRULES And 12) / 4
cboRULE.ListIndex = RELRULE
cboTYPE.ListIndex = RELTYPE
If Not ENV.FFOPEN Then Call FAMOpen("READ")
Show
End Sub


Private Sub cboRULE_Click()
RELRULE = cboRULE.ListIndex
ENV.RELRULES = ((ENV.RELRULES And 252) + RELRULE)
If (RLCT > 0) Then
   txtValue(3) = Relation$("L", SEX2$, AGL1, RGL, RELCODE$())
End If
End Sub


Private Sub cboTYPE_Click()
RELTYPE = cboTYPE.ListIndex
MATERNAL = ((RELTYPE And 1) > 0)
PATERNAL = ((RELTYPE And 2) > 0)
ENV.RELRULES = ((ENV.RELRULES And 243) + 4 * RELTYPE)
RELTABOK = 0: Call CalcREL("FIRST")
End Sub


Private Sub cmdFIRST_Click()
   Call CalcREL("FIRST")
End Sub

Private Sub cmdLAST_Click()
   Call CalcREL("LAST")
End Sub


Private Sub cmdLIST_Click(Index As Integer)
HOLDITEM = Val(txtRID(Index).TEXT)
Call frmFHSWList.SelectName(X)
If a <> 27 Then
   txtRID(Index).TEXT = fnx$(X)
   'txtValue_LostFocus (Z)
End If
End Sub

Sub CalcREL(OPT$)
If (RID1 = 0) Or (RID2 = 0) Then Exit Sub
If Not RELTABOK Then
   GoSub 1520
   If Not RELTABOK Then Exit Sub
End If
If CACNT = 0 Then GoSub 1550
Select Case OPT$
   Case "FIRST": If CACNT > 0 Then RLCT = 1
   Case "PREV": If RLCT > 1 Then RLCT = RLCT - 1
   Case "NEXT": If (RLCT = CACNT) Then GoSub 1555
                If RLCT < CACNT Then RLCT = RLCT + 1
   Case "LAST": While Not SearchComplete: GoSub 1555: Wend
                RLCT = CACNT
End Select
GoSub 1580
Exit Sub

1520 'Build Ancestor Table for ID #1
     'Call PgmPREP("FHSUPDT")
     'Call PutMSG("UPDTM011")
     'Call PutMSG("Building Ancestor table for RID1...")
     AMAX1 = 100: Y2 = 0
     ReDim RELATE(FF1Hdr.MAXID), ATAB(AMAX1), ARGL1(AMAX1), ALIN1(AMAX1) As Double
     AGL = 0: X = 1: Y = 1: ATAB(1) = RID1: ALIN1(1) = 1
     RELATE(RID1) = 1
     SearchComplete = 0: CACNT = 0: RLCT = 0
1525 ' Process next generation of ancestors
     AGL = AGL + 1
     Z = Y
     While (X <= Z)
        id = ATAB(X): Call FF1GetRec(id)
        id = FF1Rec.FID
        If (id > 0) And (Not MATERNAL) Then P2 = 1: GoSub 1540
        id = FF1Rec.MID
        If (id > 0) And (Not PATERNAL) Then P2 = 2: GoSub 1540
        X = X + 1
     Wend
     If X <= Y Then GoTo 1525     ' If any found, go process them
     RELTABOK = -1
     'Call ClearMSG
     Return

1540 If RELATE(id) = 0 Then
        Y = Y + 1: RELATE(id) = Y
        If Y > AMAX1 Then
           AMAX1 = AMAX1 + 100
           ReDim Preserve ATAB(AMAX1), ARGL1(AMAX1), ALIN1(AMAX1) As Double
        End If
        ATAB(Y) = id: ARGL1(Y) = AGL
        ALIN1(Y) = 2 * ALIN1(X) + P2 - 1
     End If
     Return

1550 ' Find Common Ancestor of RID2
     If SearchComplete Then Return
     CAMAX = 50: AMAX2 = AMAX1
     ReDim ATAB(AMAX2), ALIN2(AMAX2) As Double
     ReDim CATAB(CAMAX), CARGL2(CAMAX), CALIN2(CAMAX) As Double
     For X = 1 To FF1Hdr.MAXID
         If RELATE(X) < 0 Then RELATE(X) = 0
     Next X
     AGL = 0: id = 0: RLCT = 0
     X2 = 1: Y2 = 0: Z2 = 0: CACNT = 0
     ATAB(1) = RID2: ALIN2(1) = 1
     id = RID2
     SearchComplete = (RID1 = RID2)
     FOUND = (RELATE(id) > 0): If FOUND Then LIN2# = 1: GoTo 1575
     ' Look for Common Ancestor
     RLCT = 0: X2 = 1: Y2 = 1: RELATE(RID2) = -1
1552 ' Process next generation of ancestors
     AGL = AGL + 1: Z2 = Y2: P2 = 0
1555 CA$ = "": FOUND = 0
1556 While (X2 <= Z2) And (Not FOUND)
        id = ATAB(X2): Call FF1GetRec(id)
        If ((P2 And 1) = 0) Then
           P2 = P2 Or 1
           If (Not MATERNAL) Then
              id = FF1Rec.FID: If id > 0 Then GoSub 1570
           End If
        End If
        If (Not FOUND) Then
           P2 = 0
           If (Not PATERNAL) Then
              id = FF1Rec.MID: If id > 0 Then GoSub 1570
           End If
        End If
        If (P2 = 0) Then
           X2 = X2 + 1
        End If
     Wend
     If FOUND Then GoTo 1575
     If (X2 <= Z2) Then GoTo 1555
     If X2 <= Y2 Then GoTo 1552
     SearchComplete = -1
     Return

1570 ' Check Ancestor of RID2
     LIN2# = 2 * ALIN2(X2) - (P2 = 0)
     If (RELATE(id) = 0) Then
        Y2 = Y2 + 1: RELATE(id) = -Y2
        If Y2 > AMAX2 Then
           AMAX2 = AMAX2 + 100
           ReDim Preserve ATAB(AMAX2), ALIN2(AMAX2) As Double
        End If
        ATAB(Y2) = id
        ALIN2(Y2) = LIN2#
        Else
        FOUND = (RELATE(id) > 0)
     End If
     Return
     
1575 ' Save Information about Common Ancestor
     CACNT = CACNT + 1
     If CACNT > CAMAX Then
        CAMAX = CAMAX + 20
        ReDim Preserve CATAB(CAMAX), CARGL2(CAMAX), CALIN2(CAMAX) As Double
     End If
     CATAB(CACNT) = id
     CALIN2(CACNT) = LIN2#
     CARGL2(CACNT) = AGL
     Return

1580 ' Show Relationship Information
     If RLCT = 0 Then
        For X = 2 To 12: txtValue(X) = "": Next X
        txtValue(3) = Relation$("L", "", 0, -1, RELCODE$())
        Return
     End If
     id = CATAB(RLCT)
     LIN1# = ALIN1(RELATE(id)): AGL1 = -ARGL1(RELATE(id))
     LIN2# = CALIN2(RLCT): AGL2 = -CARGL2(RLCT)
     RGL = AGL1 - AGL2
     Call FF1GetRec(id)
     RELWORK$ = Relation$("L", SEX2$, AGL1, RGL, RELCODE$())
     For X = 2 To 12
           Select Case X
              Case 2: xX$ = LTrim$(Str$(RLCT))
              Case 3: xX$ = RELWORK$
              Case 4: xX$ = Str$(RGL)
              Case 5: xX$ = LTrim$(Str$(id))
              Case 6: xX$ = FMTNAME$(1)
              Case 7: xX$ = XRPTDate$(FF1Rec.BMD, FF1Rec.BY, 0) + FDSTAT$("BD")
              Case 8: xX$ = XRPTDate$(FF1Rec.DMD, FF1Rec.DY, 0) + FDSTAT$("DD")
              Case 9: xX$ = Str$(AGL1)
              Case 10: xX$ = Str$(LIN1#)
              Case 11: xX$ = Str$(AGL2)
              Case 12: xX$ = Str$(LIN2#)
           End Select
           txtValue(X) = xX$
    Next X
    frmFHSWRCLC.Refresh
    Return
End Sub



Private Sub cmdNEXT_Click()
   Call CalcREL("NEXT")
End Sub

Private Sub cmdPREV_Click()
   Call CalcREL("PREV")
End Sub

Private Sub cmdRETURN_Click()
Unload frmFHSWRCLC
End Sub


Private Sub cmdSEARCH_Click()
   Call CalcREL("FIRST")
End Sub

Private Sub Form_Load()
If FormLoaded(2) And SPSEFORM Then
   frmFHSWSpse.Enabled = False
   Else
   If FormLoaded(2) And UpdtForm Then
      frmFHSWUPDT.Enabled = False
   End If
End If
' Set Default Form Location on Screen
frmHEIGHT = frmFHSWRCLC.Height
frmWIDTH = frmFHSWRCLC.WIDTH
t1 = GetSetting("FHSWIN", "frmRCLC", "TOP", 0)
If t1 = 0 Then
   t00 = mnuFHSWMain.TOP: w00 = mnuFHSWMain.WIDTH
   l00 = mnuFHSWMain.Left
   t1 = (t00 + 1435)
   l1 = l00 + (w00 - frmWIDTH) / 2: If l1 < 0 Then l1 = 1
   Else
   t1 = 1440: l1 = 1440
End If
FormLoaded(2) = FormLoaded(2) Or RelCalcForm
frmFHSWRCLC.TOP = GetSetting("FHSWIN", "frmRCLC", "Top", t1)
frmFHSWRCLC.Left = GetSetting("FHSWIN", "frmRCLC", "Left", l1)
End Sub


Private Sub Form_Resize()
frmFHSWRCLC.Height = frmHEIGHT
frmFHSWRCLC.WIDTH = frmWIDTH
End Sub


Private Sub Form_Unload(Cancel As Integer)
FormLoaded(2) = FormLoaded(2) And (255 - RelCalcForm)
SaveSetting "FHSWIN", "frmRCLC", "Top", frmFHSWRCLC.TOP
SaveSetting "FHSWIN", "frmRCLC", "Left", frmFHSWRCLC.Left
If FormLoaded(2) And SpouseForm Then
   frmFHSWSpse.Enabled = True
   frmFHSWSpse.SetFocus
   Else
   If FormLoaded(2) And UpdtForm Then
      frmFHSWUPDT.Enabled = True
      frmFHSWUPDT.SetFocus
   Else
      Call FamCLOSE
      mnuFHSWMain.SetFocus
   End If
End If
End Sub


Private Sub txtRID_Change(Index As Integer)
xX& = Val(txtRID(Index))
If (xX& < 1) Or (xX& > FF1Hdr.MAXID) Then
   txtValue(Index) = ""
   Exit Sub
End If
X = xX&
txtRID(Index) = X
If Not ENV.FFOPEN Then Call FAMOpen("READ")
Call FF1GetRec(X)
txtValue(Index) = FMTNAME$(1)
Select Case Index
   Case 0: RID1 = X: RELTABOK = 0
   Case 1: RID2 = X: SEX2$ = FF1Rec.SEX
           RLCT = 0: SearchComplete = 0: CACNT = 0
End Select
For X = 2 To 12: txtValue(X) = "": Next X
End Sub


Private Sub txtRID_GotFocus(Index As Integer)
Select Case Index
   Case 0: LastRID1 = Val(txtRID(0).TEXT)
   Case 1: LastRID2 = Val(txtRID(1).TEXT)
End Select
txtRID(Index).SelStart = 0
txtRID(Index).SelLength = Len(txtRID(Index).TEXT)
End Sub


