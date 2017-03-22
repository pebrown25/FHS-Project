VERSION 4.00
Begin VB.Form frmFHSWTTFL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tiny Tafel Surname Table"
   ClientHeight    =   3675
   ClientLeft      =   5550
   ClientTop       =   6300
   ClientWidth     =   7095
   Height          =   4080
   Left            =   5490
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   Top             =   5955
   Width           =   7215
   Begin VB.CommandButton cmdButton 
      Caption         =   "Return"
      Height          =   375
      Index           =   1
      Left            =   3600
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Change Table"
      Height          =   375
      Index           =   0
      Left            =   1800
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Interest Levels:  0="" ""   1="".""   2="":""   3=""*"""
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
      Left            =   720
      TabIndex        =   1
      Top             =   2880
      Width           =   5415
   End
   Begin MSGrid.Grid grdTTFLTable 
      Height          =   2665
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6900
      _Version        =   65536
      _ExtentX        =   12171
      _ExtentY        =   4710
      _StockProps     =   77
      BackColor       =   16777215
      Cols            =   8
   End
End
Attribute VB_Name = "frmFHSWTTFL"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
DefInt A-Z
Dim xLVL(), cLVL, EDTIX
Dim LVLVLth, NewLVLValu$, OldLVLValu$
Dim LvlRow, LvlCol, NextRow, NextCol
Dim UpdatesAllowed, FormUnloaded
Sub ShowTable(TDATA$(), TMAX, TCHG)
grdTTFLTable.Enabled = False
UpdatesAllowed = False
scrollwid = 250
grdTTFLTable.Cols = 8
grdTTFLTable.Rows = TMAX + 1
grdTTFLTable.HighLight = False
grdTTFLTable.ColWidth(0) = 500
grdTTFLTable.ColWidth(1) = 400: grdTTFLTable.ColAlignment(1) = 2
grdTTFLTable.ColWidth(2) = 600: grdTTFLTable.ColAlignment(2) = 2
grdTTFLTable.ColWidth(3) = 600: grdTTFLTable.ColAlignment(3) = 2
grdTTFLTable.ColWidth(4) = 400: grdTTFLTable.ColAlignment(4) = 2
grdTTFLTable.ColWidth(5) = 600: grdTTFLTable.ColAlignment(5) = 2
grdTTFLTable.ColWidth(6) = 400: grdTTFLTable.ColAlignment(6) = 2
grdTTFLTable.ColWidth(7) = 3000
'h = (grdTTFLTable.Height - 100) / 10
grdTTFLTable.Row = 0
grdTTFLTable.Col = 1: grdTTFLTable.TEXT = " DEL"
grdTTFLTable.Col = 2: grdTTFLTable.TEXT = "      #"
grdTTFLTable.Col = 3: grdTTFLTable.TEXT = "   ABY"
grdTTFLTable.Col = 4: grdTTFLTable.TEXT = " AIL"
grdTTFLTable.Col = 5: grdTTFLTable.TEXT = "   DBY"
grdTTFLTable.Col = 6: grdTTFLTable.TEXT = " DIL"
For X = 1 To grdTTFLTable.Rows - 1
    grdTTFLTable.Row = X
    For Z = 0 To 7
        grdTTFLTable.Col = Z
        Select Case Z
           Case 0: grdTTFLTable.TEXT = Str$(X)
           Case 1: grdTTFLTable.TEXT = MID$(TDATA$(X), 2, 1)
           Case 2: grdTTFLTable.TEXT = MID$(TDATA$(X), 3, 4)
           Case 3: grdTTFLTable.TEXT = MID$(TDATA$(X), 7, 4)
           Case 4: grdTTFLTable.TEXT = MID$(TDATA$(X), 11, 1)
           Case 5: grdTTFLTable.TEXT = MID$(TDATA$(X), 12, 4)
           Case 6: grdTTFLTable.TEXT = MID$(TDATA$(X), 16, 1)
           Case 7: grdTTFLTable.TEXT = MID$(TDATA$(X), 17, 20)
        End Select
    Next Z
Next X
grdTTFLTable.Enabled = True
frmFHSWTTFL.Show 1
If FormUnloaded Then Exit Sub
grdTTFLTable.Enabled = False
UpdatesAllowed = False: TCHG = 0
For X = 1 To grdTTFLTable.Rows - 1
    grdTTFLTable.Row = X: rchg = 0
    grdTTFLTable.Col = 1
    If grdTTFLTable.TEXT <> MID$(TDATA$(X), 2, 1) Then
       Mid$(TDATA$(X), 2, 1) = grdTTFLTable.TEXT: rchg = rchg + 1
    End If
    grdTTFLTable.Col = 4
    If grdTTFLTable.TEXT <> MID$(TDATA$(X), 11, 1) Then
       Mid$(TDATA$(X), 11, 1) = grdTTFLTable.TEXT: rchg = rchg + 1
    End If
    grdTTFLTable.Col = 6
    If grdTTFLTable.TEXT <> MID$(TDATA$(X), 16, 1) Then
       Mid$(TDATA$(X), 16, 1) = grdTTFLTable.TEXT: rchg = rchg + 1
    End If
    Mid$(TDATA$(X), 1, 1) = Chr$(rchg)
    TCHG = TCHG + rchg
Next X
Unload frmFHSWTTFL
End Sub
Private Sub TurnONLvls()
If Not UpdatesAllowed Then
   grdTTFLTable.Enabled = False
   NewLVLValu$ = "": OldLVLValu$ = ""
   CMDButton(0).Caption = "Save Changes"
   CMDButton(1).Caption = "End"
   LvlRow = grdTTFLTable.TopRow
   LvlCol = 1
   grdTTFLTable.Row = LvlRow
   grdTTFLTable.Col = LvlCol
   grdTTFLTable.SelStartRow = LvlRow
   grdTTFLTable.SelStartCol = LvlCol
   grdTTFLTable.SelEndRow = LvlRow
   grdTTFLTable.SelEndCol = LvlCol
   grdTTFLTable.HighLight = False
   UpdatesAllowed = True
   grdTTFLTable.Enabled = True
   grdTTFLTable.SetFocus
End If
End Sub


Private Sub CMDButton_Click(Index As Integer)
Select Case CMDButton(Index).Caption
   Case "Change Table": TurnONLvls
   Case "Save Changes": TurnOFFLvls: ICHG = 9
        frmFHSWTTFL.Visible = False
   Case "End": TurnOFFLvls
   Case "Return": frmFHSWTTFL.Visible = False
End Select
End Sub

Private Sub TurnOFFLvls()
If UpdatesAllowed Then
   grdTTFLTable.Enabled = False
   grdTTFLTable.HighLight = False
   CMDButton(0).Caption = "Change Table"
   CMDButton(1).Caption = "Return"
   UpdatesAllowed = False
   grdTTFLTable.Enabled = True
End If
End Sub

Private Sub Form_Load()
   FormLoaded(3) = FormLoaded(3) Or TTFLForm
 ' Set Default Form Location on Screen
   frmHEIGHT = 4080
   frmWIDTH = frmFHSWTTFL.WIDTH
   t1 = GetSetting("FHSWIN", "frmTTFL", "TOP", 0)
   If t1 = 0 Then
      t00 = frmFHSWUPgm.TOP: w00 = frmFHSWUPgm.WIDTH
      l00 = frmFHSWUPgm.Left
      t1 = (t00 + 835)
      l1 = l00 + (w00 - frmWIDTH) / 2: If l1 < 0 Then l1 = 1
      Else
      t1 = 1440: l1 = 1440
   End If
   frmFHSWTTFL.TOP = GetSetting("FHSWIN", "frmTTFL", "Top", t1)
   frmFHSWTTFL.Left = GetSetting("FHSWIN", "frmTTFL", "Left", l1)
   frmFHSWTTFL.Height = frmHEIGHT
End Sub

Private Sub Form_Unload(Cancel As Integer)
If frmFHSWTTFL.Visible Then
   Call TurnOFFLvls
   frmFHSWTTFL.Visible = False
   FormUnloaded = True
   Exit Sub
End If
FormLoaded(3) = FormLoaded(3) And (255 - TTFLForm)
SaveSetting "FHSWIN", "frmTTFL", "Top", frmFHSWTTFL.TOP
SaveSetting "FHSWIN", "frmTTFL", "Left", frmFHSWTTFL.Left
frmFHSWUPgm.Enabled = True
frmFHSWUPgm.SetFocus
End Sub


Private Sub grdTTFLTable_KeyDown(KeyCode As Integer, Shift As Integer)
If Not UpdatesAllowed Then KeyCode = 0: Exit Sub
UpdatesAllowed = False
'grdTTFLTable.Enabled = False
KeyAscii = Asc(UCase(Chr(KeyCode)))
a = KeyAscii: CA$ = Chr$(a)
If (a = 13) Or (a = vbKeyDown) Then ' Enter key or down arrow key
   If grdTTFLTable.Row < grdTTFLTable.Rows - 1 Then
      grdTTFLTable.Row = grdTTFLTable.Row + 1
   End If
   Else
   If a = 27 Then ' Escape key
      UpdatesAllowed = True: CMDButton_Click (1)
      Else
      If (a = vbKeyTab) Or (a = vbKeyRight) Then ' Tab Key or right arrow key
         Select Case grdTTFLTable.Col
           Case 1: grdTTFLTable.Col = 4
           Case 4: grdTTFLTable.Col = 6
           Case 6: grdTTFLTable.Col = 1
                   If grdTTFLTable.Row < grdTTFLTable.Rows - 1 Then
                      grdTTFLTable.Row = grdTTFLTable.Row + 1
                   End If
         End Select
         Else
         If a = vbKeyLeft Then ' left arrow key
            Select Case grdTTFLTable.Col
              Case 1: grdTTFLTable.Col = 6
                      If grdTTFLTable.Row > 1 Then
                         grdTTFLTable.Row = grdTTFLTable.Row - 1
                      End If
              Case 4: grdTTFLTable.Col = 1
              Case 6: grdTTFLTable.Col = 4
            End Select
            Else
            If a = vbKeyUp Then ' Up Arrow key
               If grdTTFLTable.Row > 1 Then
                  grdTTFLTable.Row = grdTTFLTable.Row - 1
               End If
               Else
               If a = vbKeyPageDown Then  ' PageDown key
                  If (grdTTFLTable.Rows < (grdTTFLTable.TopRow + 20)) Then
                     If grdTTFLTable.TopRow > 10 Then
                        grdTTFLTable.TopRow = grdTTFLTable.Rows - 10
                        Else
                        grdTTFLTable.TopRow = 1
                     End If
                     Else
                     grdTTFLTable.TopRow = grdTTFLTable.TopRow + 10
                  End If
                  If LvlRow < grdTTFLTable.TopRow Then LvlRow = grdTTFLTable.TopRow
                  grdTTFLTable.Row = LvlRow
                  Else
                  If a = vbKeyPageUp Then  ' PageUp key
                     If grdTTFLTable.TopRow < 11 Then
                        grdTTFLTable.TopRow = 1
                        Else
                        grdTTFLTable.TopRow = grdTTFLTable.TopRow - 10
                     End If
                     LvlRow = grdTTFLTable.TopRow + 9
                     If LvlRow > grdTTFLTable.Rows - 1 Then LvlRow = grdTTFLTable.TopRow
                     grdTTFLTable.Row = LvlRow
                  End If
               End If
            End If
         End If
      End If
   End If
End If
finish:
   UpdatesAllowed = True
   grdTTFLTable.Refresh
   grdTTFLTable_RowColChange
   KeyCode = 0: KeyAscii = 0
   'grdTTFLTable.Enabled = True
End Sub

Private Sub grdTTFLTable_KeyPress(KeyAscii As Integer)
If Not UpdatesAllowed Then KeyAscii = 0: Exit Sub
UpdatesAllowed = False
'grdTTFLTable.Enabled = False
KeyAscii = Asc(UCase(Chr(KeyAscii)))
a = KeyAscii: CA$ = Chr$(a)
If grdTTFLTable.Col = 1 Then
   If (CA$ = " ") Or (CA$ = "X") Then grdTTFLTable.TEXT = CA$
End If
If (grdTTFLTable.Col = 4) Or (grdTTFLTable.Col = 6) Then
   If (CA$ = " ") Or (CA$ = ".") Or (CA$ = ":") Or (CA$ = "*") Then
      grdTTFLTable.TEXT = CA$
      Else
      If (47 < a) And (a < 52) Then
         a = a - 47: CA$ = MID$(" .:*", a, 1)
         grdTTFLTable.TEXT = CA$
      End If
   End If
End If
finish:
UpdatesAllowed = True
grdTTFLTable_SelChange
KeyAscii = 0
'grdTTFLTable.Enabled = True
End Sub

Private Sub grdTTFLTable_RowColChange()
  If Not UpdatesAllowed Then Exit Sub
  UpdatesAllowed = False
  'grdTTFLTable.Enabled = False
  If (grdTTFLTable.Row < 1) Then
     grdTTFLTable.Row = 1
  End If
  xcol = grdTTFLTable.Col
  If (xcol <> 1) And (xcol <> 4) And (xcol <> 6) Then
     grdTTFLTable.Row = LvlRow
     grdTTFLTable.Col = LvlCol
  End If
  LvlRow = grdTTFLTable.Row
  LvlCol = grdTTFLTable.Col
  If Not (grdTTFLTable.RowIsVisible(LvlRow)) Then
     TopRow = LvlRow - 9
     If TopRow < 1 Then TopRow = 1
     grdTTFLTable.TopRow = TopRow
  End If
  grdTTFLTable.SelStartRow = LvlRow
  grdTTFLTable.SelStartCol = LvlCol
  grdTTFLTable.SelEndRow = LvlRow
  grdTTFLTable.SelEndCol = LvlCol
  'grdTTFLTable.Enabled = True
  grdTTFLTable.SetFocus
  UpdatesAllowed = True
End Sub


Private Sub grdTTFLTable_SelChange()
grdTTFLTable.HighLight = False
End Sub


