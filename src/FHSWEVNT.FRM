VERSION 4.00
Begin VB.Form frmFHSWEvnt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FHSWEVNT - Event Information"
   ClientHeight    =   3285
   ClientLeft      =   5730
   ClientTop       =   6675
   ClientWidth     =   8535
   Height          =   3690
   Left            =   5670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   8535
   Top             =   6330
   Width           =   8655
   Begin VB.Frame frameEventLIST 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8500
      Begin VB.ComboBox EventLIST 
         Height          =   315
         Left            =   4440
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   600
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "Return"
         Height          =   255
         Index           =   5
         Left            =   6360
         TabIndex        =   8
         Top             =   2910
         Width           =   855
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "Comments"
         Height          =   255
         Index           =   4
         Left            =   5280
         TabIndex        =   7
         Top             =   2910
         Width           =   975
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "Move"
         Height          =   255
         Index           =   3
         Left            =   4320
         TabIndex        =   6
         Top             =   2910
         Width           =   855
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "Change"
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   3
         Top             =   2910
         Width           =   975
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "Add"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   2
         Top             =   2910
         Width           =   855
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "Delete"
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   1
         Top             =   2910
         Width           =   855
      End
      Begin VB.Label Label 
         Caption         =   $"FHSWEVNT.frx":0000
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   8295
      End
      Begin MSGrid.Grid grdEventLIST 
         Height          =   2415
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   8300
         _Version        =   65536
         _ExtentX        =   14631
         _ExtentY        =   4260
         _StockProps     =   77
         BackColor       =   16777215
         Rows            =   10
         Cols            =   7
         FixedRows       =   0
         FixedCols       =   0
         ScrollBars      =   2
      End
   End
End
Attribute VB_Name = "frmFHSWEvnt"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
DefInt A-Z
Dim EVTRNO, EVNTIX, EVNTData$(), EVNTChn() As ChainPTRS
Dim EVNTRno(), EVNTCom(), EVNTChg()
Dim EDataTable As TablePARMS, EventTable As TablePARMS
Dim ESRTYPE, ESRNO, UpdatesAllowed, ChangesMade, TNUM
Dim EDATA$, BData$, EvntRow, EvntCol, EvntMax, EvntCNT, NewEvent
Private Sub ADDEvent()
NewEvent = -1: TurnONUpdates
If EvntCNT = 0 Then addrow = 0 Else addrow = grdEventLIST.Row
If EvntCNT > 0 Then
   If Not (EvntCNT < EvntMax) Then
      EvntMax = EvntMax + 20
      ReDim Preserve EVNTChn(EvntMax) As ChainPTRS, EVNTData$(EvntMax)
      ReDim Preserve EVNTRno(EvntMax), EVNTCom(EvntMax), EVNTChg(EvntMax)
   End If
   If (addrow + 1) < EvntCNT Then
      For X = EvntCNT To EvntRow Step -1
         EVNTData$(X + 1) = EVNTData$(X)
         EVNTRno(X + 1) = EVNTRno(X)
         EVNTCom(X + 1) = EVNTCom(X)
         EVNTChg(X + 1) = EVNTChg(X)
      Next X
      Else
      addrow = addrow + 1
   End If
End If
EVNTData$(addrow) = BData$
EVNTRno(addrow) = 0
EVNTCom(addrow) = 0
EVNTChg(addrow) = 0
EvntCNT = EvntCNT + 1
Call ListEvents
EvntRow = addrow: EvntCol = 0
grdEventLIST.Row = EvntRow
grdEventLIST.col = EvntCol
For X = 6 To 0 Step -1: Call ShowEData(X): Next X
End Sub

Private Sub REMEvent()
remrow = grdEventLIST.Row
If (remrow + 1) < EvntCNT Then
   For X = (remrow + 1) To EvntCNT
      EVNTData$(X) = EVNTData$(X + 1)
      EVNTRno(X) = EVNTRno(X + 1)
      EVNTCom(X) = EVNTCom(X + 1)
      EVNTChg(X) = EVNTChg(X + 1)
   Next X
   Else
   remrow = remrow - 1
End If
EvntCNT = EvntCNT - 1
Call ListEvents
EvntRow = remrow: EvntCol = 0
grdEventLIST.Row = EvntRow
grdEventLIST.col = EvntCol
End Sub

Private Sub KeyDown(KC, S)
a = KC
If (a = 13) Or (a = vbKeyDown) Then ' Enter key or down arrow key
   If Not NewEvent Then
      If grdEventLIST.Row < grdEventLIST.Rows - 1 Then
         grdEventLIST.Row = grdEventLIST.Row + 1
      End If
   End If
   Else
   If a = 27 Then ' Escape key
      UpdatesAllowed = True: CMDButton_Click (3)
      Else
      If (a = vbKeyTab) Or (a = vbKeyRight) Then ' Tab Key or right arrow key
         If grdEventLIST.col < 5 Then
            grdEventLIST.col = grdEventLIST.col + 1 - (grdEventLIST.col = 1)
            Else
            grdEventLIST.col = 0
            If Not NewEvent Then
               If grdEventLIST.Row < grdEventLIST.Rows - 1 Then
                  grdEventLIST.Row = grdEventLIST.Row + 1
               End If
            End If
         End If
         Else
         If a = vbKeyLeft Then ' left arrow key
            If grdEventLIST.col > 0 Then
               grdEventLIST.col = grdEventLIST.col - 1 + (grdEventLIST.col = 3)
               Else
               grdEventLIST.col = 5
               If Not NewEvent Then
                  If grdEventLIST.Row > 0 Then
                     grdEventLIST.Row = grdEventLIST.Row - 1
                  End If
               End If
            End If
            Else
            If a = vbKeyUp Then ' Up Arrow key
               If Not NewEvent Then
                  If grdEventLIST.Row > 0 Then
                     grdEventLIST.Row = grdEventLIST.Row - 1
                  End If
               End If
               Else
               If a = vbKeyPageDown Then  ' PageDown key
                  If (grdEventLIST.Rows < (grdEventLIST.TopRow + 20)) Then
                     If grdEventLIST.TopRow > 10 Then
                        grdEventLIST.TopRow = grdEventLIST.Rows - 10
                        Else
                        grdEventLIST.TopRow = 0
                     End If
                     Else
                     grdEventLIST.TopRow = grdEventLIST.TopRow + 10
                  End If
                  If EvntRow < grdEventLIST.TopRow Then EvntRow = grdEventLIST.TopRow
                  grdEventLIST.Row = EvntRow
                  Else
                  If a = vbKeyPageUp Then  ' PageUp key
                     If Not NewEvent Then
                        If grdEventLIST.TopRow < 10 Then
                           grdEventLIST.TopRow = 0
                           Else
                           grdEventLIST.TopRow = grdEventLIST.TopRow - 10
                        End If
                        EvntRow = grdEventLIST.TopRow + 10
                        If EvntRow > grdEventLIST.Rows - 1 Then EvntRow = grdEventLIST.TopRow
                        grdEventLIST.Row = EvntRow
                     End If
                  End If
               End If
            End If
         End If
      End If
   End If
End If
End Sub
Private Sub AxxEvent()
   FF3PFX.RTYPE = Chr$(6)
   Call FF3GetRec(0): If Not OKAY Then Return
   FF3EVENT.SRTYPE = Chr$(ESRTYPE): FF3EVENT.SRNO = ESRNO
   Call FF3PutRec
   If EDataTable.FF = 0 Then
      EDataTable.FF = EvntMax + 1: EvntMax = EvntMax + 20
      ReDim Preserve EVNTChn(EvntMax) As ChainPTRS, EVNTData$(EvntMax)
      ReDim Preserve EVNTRno(EvntMax)
      Call ChnINIT(EVNTChn(), EDataTable.FF, EvntMax)
   End If
   Call TableUPDT("ADD", EDataTable, EVNTChn())
   LSet EDATA$ = "": For X = 0 To 6: ShowEData (X): Next X
   'Call TableSHOW(EDataTABLE, EVNTData$(), EVNTChn())
   Call ListEvents: NewItem = -1: TurnONUpdates
   'Call TableKEY("MARK", EDataTABLE, EVNTData$())
   EVNTIX = EDataTable.CT
   'GoSub GetEVENT: If a = 27 Then GoTo AddRETURN
   'GoSub GetEDATA
   'Call TableSHOW(EDataTABLE, EVNTData$(), EVNTChn())
   'Call TableUPDT("CHG", EDataTABLE, EVNTData$(), EVNTChn())
AddRETURN:
   If a = 27 Then
      Call FF3DelRec(FF3PFX.RNO)
      'Call TableUPDT("REM", EDataTABLE, EVNTData$(), EVNTChn())
      Else
      EVTRNO = FF3PFX.RNO: EVNTIX = EDataTable.CT
      EVNTRno(EVNTIX) = EVTRNO
      Call WriteEVENT
   End If
TurnOFFUpdates
End Sub
Private Sub DELEvent()
DeleteEvent:
   Call PutMSG("UPDTM007")
   If CA$ <> SNGLKEY$(1) Then GoTo DeleteRETURN
   EVTRNO = EVNTRno(EvntRow + 1)
   ECOM = EVNTCom(EvntRow + 1)
   Call FF3GetRec(EVTRNO)
   ECOM = FF3EVENT.COM: LPLNM = FF3EVENT.LPLNM
   If ECOM <> 0 Then
      Call PutMSG("UPDTM009")
      If CA$ <> SNGLKEY$(1) Then GoTo DeleteRETURN
   End If
   While ECOM <> 0
       Call FF3GetRec(ECOM)
       If Not OKAY Then
          ECOM = 0
          Else
          X = FF3COMMENT.FPTR
          Call FF3DelRec(ECOM)
          If Not OKAY Then
             ECOM = 0
             Else
             ECOM = X
          End If
       End If
   Wend
   EVNTCom(EvntRow + 1) = 0
   'If XBWD <> 0 Then
   '   Call FF3GetRec(EVNTRno(XBWD))
   '   FF3EVENT.NXT = EVNTRno(XFWD)
   '   Call FF3PutRec
   '   Else
   '   EVNTBGN = EVNTRno(XFWD)
   'End If
   If LPLNM <> 0 Then Call FF3DelRec(LPLNM)
   'Call FF3DelRec(EVTRNO)
   'EVTRNO = EVNTRno(XFWD)
   'If EVTRNO = 0 Then EVTRNO = EVNTRno(XBWD)
   'Call TableUPDT("REM", EDataTABLE, EVNTData$(), EVNTChn())
   Call REMEvent
DeleteRETURN:

End Sub

Private Sub GetEDATA()
GetEDATA:
    Mid$(EDATA$, 1, 9) = FF3EVENT.CODE + Chr$(9): FindEVENT
    Mid$(EDATA$, 10, 4) = SFILL$(Asc(FF3EVENT.IMPORT), 3) + Chr$(9)
    EVNTCom(EvntRow + 1) = FF3EVENT.COM
    If FF3EVENT.COM <> 0 Then
       Mid$(EDATA$, 14, 1) = SNGLKEY$(1)
       Else
       Mid$(EDATA$, 14, 1) = " "
    End If
    Mid$(EDATA$, 15, 1) = Chr$(9)
    Mid$(EDATA$, 16, 11) = XSCRNDate$(FF3EVENT.D1MD, FF3EVENT.D1Y) + Chr$(9)
    Mid$(EDATA$, 27, 2) = FF3EVENT.DSTAT + Chr$(9)
    If FF3EVENT.LPLNM = 0 Then
       Mid$(EDATA$, 29, 41) = FF3EVENT.PLACE + Space$(19)
       Else
       Call FF3GetRec(FF3EVENT.LPLNM)
       Mid$(EDATA$, 29, 41) = FF3PLACE2.PLACE
    End If
    Mid$(EDATA$, 70, 1) = Chr$(9)
End Sub
Private Sub FindEVENT()
    
    FOUND = 0
    Mid$(EDATA$, 71, 25) = Space$(25)
    If TNUM <> 0 Then
       X = EventTable.FT
       While (X > 0) And (Not FOUND)
          FOUND = (MID$(TABDATA$(X), 1, 8) = MID$(EDATA$, 1, 8))
          If Not FOUND Then X = TABDATACHN(X).FWD
       Wend
    End If
    If FOUND Then Mid$(EDATA$, 71, 25) = MID$(TABDATA$(X), 13, 25)

End Sub

Private Sub GetEvent()
   EventLIST.Left = grdEventLIST.Left
   EventLIST.TOP = grdEventLIST.TOP + (grdEventLIST.Row - grdEventLIST.TopRow) * (grdEventLIST.Height / 10) - 15
   EventLIST.ZOrder 0
   EventLIST.Visible = True
   Mid$(EDATA$, 1, 8) = MID$(grdEventLIST.TEXT, 1, 8)
   Call FindEVENT
   EventLIST.Visible = False
   If FOUND Then
      EventLIST.TEXT = MID$(EDATA$, 1, 8)
      Else
      EventLIST.TEXT = "UNKNOWN "
   End If
   EventLIST.Visible = True
   EventLIST.SetFocus
End Sub

Private Sub ListEvents()
    grdEventLIST.Rows = 1
    'X = EDataTable.FT: xrow = 0
    X = EvntCNT: xrow = 0
    While X > 0
       EDATA$ = EVNTData$(xrow + 1)
       grdEventLIST.AddItem EDATA$, xrow
       xrow = xrow + 1
       X = X - 1
    Wend
    grdEventLIST.Rows = xrow
    grdEventLIST.HighLight = False
End Sub

Private Sub ShowEData(col)
   grdEventLIST.col = col
   Select Case col
      Case 0: grdEventLIST.TEXT = RTrim$(MID$(EDATA$, 1, 8))
      Case 1: grdEventLIST.TEXT = RTrim$(MID$(EDATA$, 10, 3))
      Case 2: grdEventLIST.TEXT = RTrim$(MID$(EDATA$, 14, 1))
      Case 3: grdEventLIST.TEXT = RTrim$(MID$(EDATA$, 16, 10))
      Case 4: grdEventLIST.TEXT = RTrim$(MID$(EDATA$, 27, 1))
      Case 5: grdEventLIST.TEXT = RTrim$(MID$(EDATA$, 29, 41))
      Case 6: grdEventLIST.TEXT = RTrim$(MID$(EDATA$, 71, 41))
   End Select
   grdEventLIST.col = EvntCol
End Sub

Private Sub UpdateEvents()

UpdtEVENT:
   'Call OptDISPLAY(FMT914)
   'Call TableUPDT("CHG", EDataTABLE, EVNTData$(), EVNTChn())
   'Select Case a
   '   Case 13: FUN$ = "F1"
   '   Case 27: FUN$ = "ESC"
   '   Case 59: FUN$ = "F1"
   '   Case 60: If TNUM > 0 Then FUN$ = "F2"
   '   Case Else: Call ErrBEEP(0): GoTo UpdtEVENT
   'End Select
   'Call OptHILITE(FUN$, "", 6)
   'Select Case FUN$
   '   Case "ESC": Call FF3GetRec(EVTRNO): GoSub GetEDATA: GoTo ListSHOW
   '   Case "F1": GoSub WriteEVENT: GoTo ListSELECT
   '   Case "F2": LSet EDATA$ = EVNTData$(EDataTABLE.CT)
   '              GoSub GetEVENT
   '              LSet EVNTData$(EDataTABLE.CT) = EDATA$
   '              Call TableSHOW(EDataTABLE, EVNTData$(), EVNTChn())
   'End Select
   'GoTo UpdtEVENT
   'GoTo EUpdtRETURN
   
End Sub

Sub ViewEVENTS(EVNTBGN, SRTYPE, SRNO)
    ESRTYPE = SRTYPE: ESRNO = SRNO
    UpdatesAllowed = False: ChangesMade = False
    Load frmFHSWEvnt
    ' Process EVENT Information
    DFM = Val(CONFIG.DFM)
    EDataTable.XL = 100: EDATA$ = Space$(EDataTable.XL): BData$ = EDATA$
    Mid$(BData$, 1, 8) = "UNKNOWN": Mid$(BData$, 9, 1) = Chr$(9)
    Mid$(BData$, 13, 1) = Chr$(9)
    Mid$(BData$, 15, 1) = Chr$(9)
    Mid$(BData$, 16, 10) = XSCRNDate$(0, 0): Mid$(BData$, 26, 1) = Chr$(9)
    Mid$(BData$, 28, 1) = Chr$(9)
    Mid$(BData$, 70, 1) = Chr$(9)
    TNUM = 0: FOUND = 0
    While (TNUM < TABMAX) And (Not FOUND)
       TNUM = TNUM + 1: FOUND = TDFIndex(TNUM).NAME = "EVENTS  "
    Wend
    If TNUM <> 0 Then
       EventTable = TABLE(TNUM)
       X = EventTable.FT
       While (X > 0)
          EventLIST.AddItem MID$(TABDATA$(X), 1, 8)
          X = TABDATACHN(X).FWD
       Wend
    End If
    
PrepEVENTS:
    EvntMax = 10
    ReDim EVNTChn(EvntMax) As ChainPTRS, EVNTData$(EvntMax)
    ReDim EVNTRno(EvntMax), EVNTCom(EvntMax), EVNTChg(EvntMax)
    X = EVNTBGN: EvntCNT = 0
    While X <> 0
       EvntCNT = EvntCNT + 1
       Call FF3GetRec(X)
       If OKAY Then
          EVNTRno(EvntCNT) = X: EVNTCom(EvntCNT) = FF3EVENT.COM
          LSet EDATA$ = ""
          Call GetEDATA: EVNTData$(EvntCNT) = EDATA$
          If Not (EvntCNT < EvntMax) Then
             EvntMax = EvntMax + 20
             ReDim Preserve EVNTChn(EvntMax) As ChainPTRS, EVNTData$(EvntMax)
             ReDim Preserve EVNTRno(EvntMax), EVNTCom(EvntMax)
          End If
          X = FF3EVENT.NXT
          Else
          X = 0
       End If
    Wend
    Call ChnINIT(EVNTChn(), 1, EvntMax)
    EDataTable.FT = -(EvntCNT > 0)
    EDataTable.LT = EvntCNT
    EDataTable.SIZE = EDataTable.LT
    EDataTable.FF = EvntCNT + 1
    EVNTChn(EDataTable.SIZE).FWD = 0: EVNTChn(EDataTable.FF).BWD = 0
    If EvntCNT = 0 Then Call ADDEvent
    If EvntCNT <> 0 Then
       Call ListEvents
       frmFHSWEvnt.Show 1
    End If
    Unload frmFHSWEvnt

End Sub
Private Sub WriteEVENT()

WriteEVENT:
   EVTRNO = EVNTRno(EVNTIX)
   Call FF3GetRec(EVNTRno(EVTRNO))
   GoSub PutEDATA
   XBWD = EVNTChn(EDataTable.CT).BWD
   XFWD = EVNTChn(EDataTable.CT).FWD
   FF3EVENT.NXT = EVNTRno(XFWD)
   Call FF3PutRec
   If EPLLTH < 23 Then
      If FF3EVENT.LPLNM <> 0 Then
         Call FF3DelRec(FF3EVENT.LPLNM)
      End If
      Else
      FF3PFX.RTYPE = Chr$(12)
      Call FF3GetRec(FF3EVENT.LPLNM)
      FF3PLACE2.SRTYPE = Chr$(6)
      FF3PLACE2.SRNO = EVTRNO
      LSet FF3PLACE2.PLACE = MID$(EDATA$, 24, 41)
      Call FF3PutRec
      If FF3EVENT.LPLNM = 0 Then
         FF3EVENT.LPLNM = FF3PFX.RNO
         FF3PFX.RTYPE = Chr$(6): FF3PFX.RNO = EVTRNO
         Call FF3PutRec
      End If
   End If
   If XBWD <> 0 Then
      Call FF3GetRec(EVNTRno(XBWD))
      FF3EVENT.NXT = EVTRNO
      Call FF3PutRec
      Else
      EVNTBGN = EVTRNO
   End If
   Exit Sub

PutEDATA:
    LSet EDATA$ = EVNTData$(EVNTIX)
    FF3EVENT.CODE = MID$(EDATA$, 1, 8)
    X = Val(MID$(EDATA$, 9, 3))
    If X < 256 Then FF3EVENT.IMPORT = Chr$(X)
    FF3EVENT.D1MD = DATEMD(MID$(EDATA$, 13, 10), DFM)
    FF3EVENT.D1Y = DATEY(MID$(EDATA$, 13, 10), DFM)
    FF3EVENT.DSTAT = MID$(EDATA$, 23, 1)
    EPLLTH = Len(RTrim$(MID$(EDATA$, 24, 41)))
    If (EPLLTH > 22) And (Not ENV.LPNSupport) Then EPLLTH = 22
    LSet FF3EVENT.PLACE = MID$(EDATA$, 24, 22)
    Return
End Sub

Private Sub CMDButton_Click(Index As Integer)
Select Case CMDButton(Index).Caption
   Case "Add": Call ADDEvent
   Case "Change": TurnONUpdates
   Case "Delete": Call DELEvent
   Case "End": TurnOFFUpdates
   Case "Return": frmFHSWEvnt.Visible = False
End Select
End Sub
      

Private Sub EventLIST_Click()
If Not UpdatesAllowed Then Exit Sub
UpdatesAllowed = False
If grdEventLIST.TEXT <> EventLIST.TEXT Then
   grdEventLIST.TEXT = EventLIST.TEXT
   EDATA$ = EVNTData$(grdEventLIST.Row + 1)
   Mid$(EDATA$, 1, 8) = grdEventLIST.TEXT
   Call FindEVENT
   LSet EVNTData$(grdEventLIST.Row + 1) = EDATA$
   Call ShowEData(6)
End If
UpdatesAllowed = True
End Sub

Private Sub EventLIST_KeyDown(KeyCode As Integer, Shift As Integer)
If Not UpdatesAllowed Then KeyCode = 0: Exit Sub
UpdatesAllowed = False
Call KeyDown(KeyCode, Shift)
finish:
   UpdatesAllowed = True
   grdEventLIST.Refresh
   grdEventLIST_RowColChange
   KeyCode = 0: KeyAscii = 0
   'grdEventLIST.Enabled = True
End Sub


Private Sub Form_Load()
Call FormINIT
frmFHSWUPDT.Enabled = False
t0 = frmFHSWUPDT.TOP: l0 = frmFHSWUPDT.Left
'frmHEIGHT = frmFHSWEvnt.Height
'frmWIDTH = frmFHSWEvnt.WIDTH
frmHEIGHT = 3690: frmWIDTH = 8655
t1 = t0 + 2715
If (t1 + frmHEIGHT) > Screen.Height Then t1 = frmFHSWUPDT.TOP + 450
l1 = l0 - (frmWIDTH - frmFHSWUPDT.WIDTH) / 2
If l1 < 0 Then l1 = frmFHSWUPDT.Left
frmFHSWEvnt.TOP = GetSetting("FHSWIN", "frmEVNT", "Top", t1)
frmFHSWEvnt.Left = GetSetting("FHSWIN", "frmEVNT", "Left", l1)
'frmFHSWEvnt.WIDTH = GetSetting("FHSWIN", "frmEVNT", "Width", frmWIDTH)
'frmFHSWEvnt.Height = GetSetting("FHSWIN", "frmEVNT", "Height", frmHEIGHT)
FormLoaded(3) = FormLoaded(3) Or EventForm
End Sub
Private Sub FormINIT()
grdEventLIST.Visible = False
'grdEventLIST.RowHeight = (grdEventLIST.Height - 11) / 10
grdEventLIST.ColWidth(0) = 1000
grdEventLIST.ColWidth(1) = 400
grdEventLIST.ColWidth(2) = 400
grdEventLIST.ColWidth(3) = 1100
grdEventLIST.ColWidth(4) = 300
grdEventLIST.ColWidth(5) = 2300
grdEventLIST.ColWidth(6) = 2400
For X = 0 To 4: grdEventLIST.ColAlignment(X) = 2: Next X

'grdEventLIST.Row = 0
'For X = 0 To 6
'    grdEventLIST.Col = X
'    Select Case X
'       Case 0: grdEventLIST.ColAlignment = 2
'       Case 1: grdEventLIST.ColAlignment = 2
'       Case 2: grdEventLIST.ColAlignment = 2
'       Case 3: grdEventLIST.ColAlignment = 2
'       Case 4: grdEventLIST.ColAlignment = 2
'       Case 5: grdEventLIST.ColAlignment = 2
'       Case 6: grdEventLIST.ColAlignment = 2
'    End Select
'Next X
EventLIST.Left = grdEventLIST.Left
'EventLIST.Height = (grdEventLIST.Height) / 10
grdEventLIST.Visible = True
End Sub

Private Sub MOVEvent()

MoveEVENT:
   XBWD = EVNTChn(EDataTable.CT).BWD: XFWD = EVNTChn(EDataTable.CT).FWD
   'Call TableUPDT("MOV", EDataTABLE, EVNTData$(), EVNTChn())
   If XBWD = EVNTChn(EDataTable.CT).BWD Then GoTo MoveRETURN
   If XBWD > 0 Then
      Call FF3GetRec(EVNTRno(XBWD))
      FF3EVENT.NXT = EVNTRno(XFWD)
      Call FF3PutRec
   End If
   XBWD = EVNTChn(EDataTable.CT).BWD: XFWD = EVNTChn(EDataTable.CT).FWD
   If XBWD > 0 Then
      Call FF3GetRec(EVNTRno(XBWD))
      FF3EVENT.NXT = EVTRNO
      Call FF3PutRec
   End If
   Call FF3GetRec(EVTRNO)
   FF3EVENT.NXT = EVNTRno(XFWD)
   Call FF3PutRec
   EVNTBGN = EVNTRno(EDataTable.FT)
   EVNTIX = EDataTable.CT
MoveRETURN:
   Return

End Sub

Private Sub TurnOFFUpdates()
If UpdatesAllowed Then
   If NewEvent Then
      If RTrim$(grdEventLIST.TEXT) = "" Then
         Call REMEvent
      End If
      NewEvent = 0
   End If
   grdEventLIST.Enabled = False
   grdEventLIST.HighLight = False
   CMDButton(0).Visible = True
   CMDButton(1).Visible = True
   CMDButton(2).Caption = "Delete": CMDButton(2).TabStop = True
   CMDButton(3).Caption = "Move": CMDButton(3).TabStop = True
   CMDButton(4).Visible = True
   CMDButton(5).Visible = True
   EventLIST.ZOrder 1
   UpdatesAllowed = False
   grdEventLIST.Enabled = True
End If
End Sub
Private Sub TurnONUpdates()
If Not UpdatesAllowed Then
   grdEventLIST.Enabled = False
   NewLVLValu$ = "": OldLVLValu$ = ""
   CMDButton(0).Visible = False
   CMDButton(1).Visible = False
   CMDButton(2).Caption = "Save": CMDButton(2).TabStop = False
   CMDButton(3).Caption = "End": CMDButton(3).TabStop = False
   CMDButton(4).Visible = False
   CMDButton(5).Visible = False
   If grdEventLIST.col = 6 Then grdEventLIST.col = 0
   EvntRow = grdEventLIST.Row
   EvntCol = grdEventLIST.col
   grdEventLIST.SelStartRow = EvntRow
   grdEventLIST.SelStartCol = EvntCol
   grdEventLIST.SelEndRow = EvntRow
   grdEventLIST.SelEndCol = EvntCol
   grdEventLIST.HighLight = False
   UpdatesAllowed = True
   grdEventLIST.Enabled = True
   frmFHSWEvnt.Visible = True
   grdEventLIST.SetFocus
   If grdEventLIST.col = 0 Then Call GetEvent
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
FormLoaded(3) = FormLoaded(3) And (255 - EventForm)
SaveSetting "FHSWIN", "frmEVNT", "Top", frmFHSWEvnt.TOP
SaveSetting "FHSWIN", "frmEVNT", "Left", frmFHSWEvnt.Left
'SaveSetting "FHSWIN", "frmEVNT", "Width", frmFHSWEvnt.WIDTH
'SaveSetting "FHSWIN", "frmEVNT", "Height", frmFHSWEvnt.Height
frmFHSWUPDT.Enabled = True
'frmFHSWUPDT.SetFocus
End Sub


Private Sub grdEventLIST_KeyDown(KeyCode As Integer, Shift As Integer)
If Not UpdatesAllowed Then KeyCode = 0: Exit Sub
UpdatesAllowed = False
Call KeyDown(KeyCode, Shift)
finish:
   UpdatesAllowed = True
   grdEventLIST.Refresh
   grdEventLIST_RowColChange
   KeyCode = 0: KeyAscii = 0
   'grdEventLIST.Enabled = True
End Sub

Private Sub grdEventLIST_KeyPress(KeyAscii As Integer)
If Not UpdatesAllowed Then KeyAscii = 0: Exit Sub
UpdatesAllowed = False
'grdEventLIST.Enabled = False
KeyAscii = Asc(UCase(Chr(KeyAscii)))
a = KeyAscii: CA$ = Chr$(a)
If grdEventLIST.col = 1 Then
   If (CA$ = " ") Or ((Val(CA$) >= 0) And (Val(CA$) < 10)) Then
      grdEventLIST.TEXT = CA$: EVNTChg(EvntRow + 1) = -1
   End If
   Else
   If (grdEventLIST.col = 3) Then
      xFT = 3: xSS = 0: xSL = Len(grdEventLIST.TEXT): XVALUE$ = (grdEventLIST.TEXT)
      Call FldUpdate(KeyAscii, 0, xFT, xSS, xSL, XVALUE$)
      If XVALUE$ <> grdEventLIST.TEXT Then
         EVNTChg(EvntRow + 1) = -1: grdEventLIST.TEXT = XVALUE$
      End If
   End If
   If (grdEventLIST.col = 4) Then
      If (CA$ = "?") Or (CA$ = "*") Or (CA$ = "!") Or (CA$ = " ") Then
         grdEventLIST.TEXT = CA$: EVNTChg(EvntRow + 1) = -1
      End If
   End If
   If (grdEventLIST.col = 5) Then
      xFT = 1: xSS = 0: xSL = Len(grdEventLIST.TEXT): XVALUE$ = (grdEventLIST.TEXT)
      Call FldUpdate(KeyAscii, 0, xFT, xSS, xSL, XVALUE$)
      If XVALUE$ <> grdEventLIST.TEXT Then
         EVNTChg(EvntRow + 1) = -1: grdEventLIST.TEXT = XVALUE$
      End If
   End If
End If
finish:
UpdatesAllowed = True
grdEventLIST_SelChange
KeyAscii = 0
'grdEventLIST.Enabled = True
End Sub

Private Sub grdEventLIST_RowColChange()
  If Not UpdatesAllowed Then Exit Sub
  UpdatesAllowed = False
  'grdEventLIST.Enabled = False
  If (grdEventLIST.Row < 0) Then
     grdEventLIST.Row = 0
  End If
  xcol = grdEventLIST.col
  If (xcol = 6) Then
     grdEventLIST.Row = EvntRow
     grdEventLIST.col = EvntCol
  End If
  xrow = grdEventLIST.Row
  If xrow > grdEventLIST.Rows - 1 Then
     grdEventLIST.Row = grdEventLIST.Rows - 1
  End If
  EvntRow = grdEventLIST.Row
  EvntCol = grdEventLIST.col
  If Not (grdEventLIST.RowIsVisible(EvntRow)) Then
     TopRow = EvntRow - 9
     If TopRow < 0 Then TopRow = 0
     grdEventLIST.TopRow = TopRow
  End If
  grdEventLIST.SelStartRow = EvntRow
  grdEventLIST.SelStartCol = EvntCol
  grdEventLIST.SelEndRow = EvntRow
  grdEventLIST.SelEndCol = EvntCol
  'grdEventLIST.Enabled = True
  'grdEventLIST.SetFocus
  If grdEventLIST.col = 0 Then
     Call GetEvent
     Else
     EventLIST.Visible = False
  End If
  UpdatesAllowed = True
End Sub

Private Sub grdEventLIST_SelChange()
grdEventLIST.HighLight = False
End Sub


Private Sub Label1_Click()

End Sub


