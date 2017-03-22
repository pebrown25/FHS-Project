VERSION 4.00
Begin VB.Form frmFHSWChld 
   Caption         =   "FHSWCHLD -  Children"
   ClientHeight    =   3000
   ClientLeft      =   1245
   ClientTop       =   2145
   ClientWidth     =   9330
   Height          =   3405
   KeyPreview      =   -1  'True
   Left            =   1185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   Top             =   1800
   Width           =   9450
   Begin VB.Frame CMDFrame 
      Height          =   525
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   9135
      Begin VB.TextBox MsgLine 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   3
         Text            =   "FHSWCHLD.frx":0000
         Top             =   210
         Visible         =   0   'False
         Width           =   5895
      End
      Begin VB.CommandButton CMDButton 
         Cancel          =   -1  'True
         Caption         =   "Return"
         Default         =   -1  'True
         Height          =   315
         Index           =   2
         Left            =   7680
         TabIndex        =   1
         Top             =   150
         Width           =   1335
      End
      Begin VB.CommandButton CMDButton 
         Caption         =   "Move Child"
         Height          =   315
         Index           =   1
         Left            =   3840
         TabIndex        =   5
         Top             =   150
         Width           =   1815
      End
      Begin VB.CommandButton CMDButton 
         Caption         =   "Select Child"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   150
         Width           =   1455
      End
   End
   Begin VB.Frame INFOFrame 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      Begin VB.ListBox lstBOX 
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS LineDraw"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   8895
      End
      Begin VB.Label INFOLbl 
         Caption         =   "#    ID     Parents  Sex    Birth       Death    Age     Given Name"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS LineDraw"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   8775
      End
   End
End
Attribute VB_Name = "frmFHSWChld"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
DefInt A-Z
Dim CCNT, CTSIZE, CHLDTable As TablePARMS
Dim CHLDLIST$(), CHLDCHN() As ChainPTRS
Dim CGL, PRID, OPID, xCHID, MoveMode, HoldIX

Sub BuildChildLIST(PID1, PID2)
   If Not ENV.FFOPEN Then Call FAMOpen("WRITE")
   lstBOX.MousePointer = 11
   If CTSIZE = 0 Then CTSIZE = 10
   ReDim CHLDLIST$(CTSIZE), CHLDCHN(CTSIZE) As ChainPTRS
   Call ChnINIT(CHLDCHN(), 1, CTSIZE): CF = 1
   CHLDTable.FT = 0: CHLDTable.LT = 0: CHLDTable.FF = 0: CHLDTable.SIZE = 0
   CHLDTable.XL = 85
   'Call TableOPEN(CHLDTable)
   If PID1 > 0 Then PRID = PID1: OPID = PID2 Else PRID = PID2: OPID = PID1
   Call FF1GetRec(PRID)
   CZ = FF1Rec.OLDCH: X$ = Space$(CHLDTable.XL)
   CCNT = 0
   While CZ > 0
      Call FF1GetRec(CZ)
      If Not OKAY Then
         CZ = 0
         Else
         If (OPID = 0) Or (((FF1Rec.FID = PID1) And (FF1Rec.MID = PID2)) Or ((FF1Rec.MID = PID1) And (FF1Rec.FID = PID2))) Then
            CCNT = CCNT + 1
            Mid$(X$, 1, 2) = SFILL$(CCNT, 2)
            Mid$(X$, 4, 5) = SFILL$(CZ, 5)
            Mid$(X$, 10, 5) = SFILL$(FF1Rec.FID, 5)
            Mid$(X$, 15, 1) = FDSTAT$("FID")
            Mid$(X$, 16, 5) = SFILL$(FF1Rec.MID, 5)
            Mid$(X$, 21, 1) = FDSTAT$("MID")
            Mid$(X$, 24, 1) = FF1Rec.SEX
            Mid$(X$, 27, 10) = XSCRNDate$(FF1Rec.BMD, FF1Rec.BY)
            Mid$(X$, 37, 1) = FDSTAT$("BD")
            Mid$(X$, 39, 10) = XSCRNDate$(FF1Rec.DMD, FF1Rec.DY)
            Mid$(X$, 49, 1) = FDSTAT$("DD")
            If (FF1Rec.DY = 9999) Or (FF1Rec.BY = 0) Then
               Mid$(X$, 50, 4) = "  ??"
               Else
               xZ = Val(DATEDif$(FF1Rec.BMD, FF1Rec.BY, FF1Rec.DMD, FF1Rec.DY))
               Mid$(X$, 50, 4) = SFILL$(xZ, 4)
            End If
            Mid$(X$, 56, 29) = FF1Rec.GIVEN
            If CF = 0 Then
               CF = CTSIZE + 1: CTSIZE = CTSIZE + 5
               ReDim Preserve CHLDLIST$(CTSIZE), CHLDCHN(CTSIZE) As ChainPTRS
               Call ChnINIT(CHLDCHN(), CF, CTSIZE)
               CHLDCHN(CF).BWD = CF - 1: CHLDCHN(CF - 1).FWD = CF
            End If
            CHLDLIST$(CF) = X$: CHLDTable.LT = CF
            CF = CHLDCHN(CF).FWD
         End If
      End If
      If FF1Rec.FID = PRID Then CZ = FF1Rec.FCH Else CZ = FF1Rec.MCH
   Wend
   CMAX = CHLDTable.LT: CHLDCHN(CMAX).FWD = 0
   ReDim Preserve CHLDLIST$(CMAX), CHLDCHN(CMAX) As ChainPTRS
   CHLDTable.FT = 1: CHLDTable.SIZE = CMAX
   CHLDTable.CL = 0
   Call BuildListBOX
   lstBOX.MousePointer = 0
   If CMAX = 0 Then
      Call PutMSG("UPDTM018")
      Exit Sub
   End If
End Sub
Private Sub BuildListBOX()
   lstBOX.Clear
   CF = CHLDTable.FT
   X$ = Space$(CHLDTable.XL)
   While CF <> 0
      LSet X$ = CHLDLIST$(CF)
      CZ = Val(MID$(X$, 4, 5))
      lstBOX.AddItem RTrim$(X$)
      lstBOX.ItemData(lstBOX.NewIndex) = CZ
      CF = CHLDCHN(CF).FWD
   Wend
   lstBOX.Refresh
End Sub


Sub MoveCHILD()
805 ' Move Child in List
   If HoldIX = lstBOX.ListIndex Then Exit Sub
   XCB = CHLDCHN(CHLDTable.CT).BWD
   If XCB > 0 Then XCB = Val(MID$(CHLDLIST$(XCB), 6, 5))
   xRID = Val(MID$(CHLDLIST$(CHLDTable.CT), 6, 5))
   Call FF1GetRec(xRID)
   NameREC = FF1Rec
   CFID = NameREC.FID: CMID = NameREC.MID
   If (CFID <> PRID) And (CFID <> OPID) Then CFID = 0
   If (CMID <> PRID) And (CMID <> OPID) Then CMID = 0
   If (CFID > 0) Then Call UpdtRelPTRS("REM", CFID)
   If (CMID > 0) Then Call UpdtRelPTRS("REM", CMID)
   If (CFID > 0) Then UXID = CFID: GoSub 810
   If (CMID > 0) Then UXID = CMID: GoSub 810
   FF1Rec = NameREC
   Call FF1PutREC
   Exit Sub

810 ' Place NameREC after XCB in child chain of UXID
   Call FF1GetRec(UXID)
   If XCB = 0 Then
      If UXID = CFID Then
         NameREC.FCH = FF1Rec.OLDCH
         Else
         NameREC.MCH = FF1Rec.OLDCH
      End If
      FF1Rec.OLDCH = NameREC.RID: OKAY = -1
      Else
      X = FF1Rec.OLDCH: FOUND = 0
      While (X > 0) And (Not FOUND)
         Call FF1GetRec(X)
         If Not OKAY Then
            X = 0
            Else
            FOUND = (X = XCB)
            If Not FOUND Then
               If FF1Rec.FID = UXID Then
                  X = FF1Rec.FCH
                  Else
                  X = FF1Rec.MCH
               End If
            End If
         End If
      Wend
      OKAY = FOUND
      If FOUND Then
         If UXID = CFID Then
            If UXID = FF1Rec.FID Then
               NameREC.FCH = FF1Rec.FCH
               FF1Rec.FCH = NameREC.RID
               Else
               NameREC.FCH = FF1Rec.MCH
               FF1Rec.MCH = NameREC.RID
            End If
            Else
            If UXID = FF1Rec.FID Then
               NameREC.MCH = FF1Rec.FCH
               FF1Rec.FCH = NameREC.RID
               Else
               NameREC.MCH = FF1Rec.MCH
               FF1Rec.MCH = NameREC.RID
            End If
         End If
      End If
   End If
   If OKAY Then Call FF1PutREC
   Return

End Sub
Public Sub ShowCHILDREN(PID1, PID2, xGL)
a = 27: OnlySHOW = True
CGL = xGL + 1
If (lstBOX.ListCount = 0) Then
   Show
   frmFHSWChld.Refresh
   Call BuildChildLIST(PID1, PID2)
   Else
   Show
End If
End Sub
Private Sub CMDButton_Click(Index As Integer)
CheckButton:
Select Case Index
   Case 0:
      If MoveMode Then
         Call MoveCHILD
         HoldIX = lstBOX.ListIndex
         Call BuildChildLIST(PRID, OPID)
         MsgLine.Visible = False
         CMDButton(0).Caption = "Select Child"
         CMDButton(1).Visible = True
         CMDButton(2).Caption = "Return"
         MoveMode = 0
         lstBOX.ListIndex = HoldIX
         Exit Sub
      End If
      If lstBOX.ListIndex < 0 Then
         Call PutMSG("First Click on Child to Select...")
         Exit Sub
      End If
      xCHID = lstBOX.ItemData(lstBOX.ListIndex)
      If xCHID = 0 Then Exit Sub
      If FormLoaded(3) And ParentForm Then
         Unload frmFHSWPare
      End If
      If FormLoaded(2) And SpouseForm Then
         Unload frmFHSWSpse
      End If
      Call frmFHSWUPDT.NRDisplay(xCHID, CGL)
      Unload frmFHSWChld
   Case 1:
      If lstBOX.ListIndex < 0 Then
         Call PutMSG("First Click on Child to be Moved...")
         Exit Sub
      End If
      MoveMode = -1: FOUND = 0
      HoldIX = lstBOX.ListIndex
      XT = CHLDTable.FT
      While (XT <> 0) And Not FOUND
         cx = Val(MID$(CHLDLIST$(XT), 6, 5))
         FOUND = (cx = lstBOX.ItemData(HoldIX))
         If Not FOUND Then
            XT = CHLDCHN(XT).FWD
         End If
      Wend
      CHLDTable.CT = XT
      CMDButton(0).Caption = "Save Move"
      CMDButton(1).Visible = False
      MsgLine.TEXT = "Use Up and Down Cursor Keys to Move Child"
      MsgLine.Visible = True
      CMDButton(2).Caption = "Cancel"
      lstBOX.SetFocus
   Case 2:
      If MoveMode Then
         MsgLine.Visible = False
         CMDButton(0).Caption = "Select Child"
         CMDButton(1).Visible = True
         CMDButton(2).Caption = "Return"
         Call BuildChildLIST(PRID, OPID)
         lstBOX.ListIndex = HoldIX
         MoveMode = 0: HoldIX = -1
         Else
         If lstBOX.ListIndex < 0 Then
            Unload frmFHSWChld
            Else
            Index = 0: GoTo CheckButton
         End If
      End If
End Select
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Not MoveMode Then Exit Sub
xZ = lstBOX.ListIndex
If KeyCode = vbKeyUp Then
   If xZ > 0 Then
      CT = CHLDTable.CT:
      BWD = CHLDCHN(CT).BWD
      FWD = CHLDCHN(CT).FWD
      CHLDCHN(CT).FWD = BWD
      CHLDCHN(CT).BWD = CHLDCHN(BWD).BWD
      CHLDCHN(BWD).FWD = FWD
      CHLDCHN(BWD).BWD = CT
      CHLDCHN(FWD).BWD = BWD
      If CHLDCHN(CT).BWD = 0 Then
         CHLDTable.FT = CT
         Else
         CHLDCHN(CHLDCHN(CT).BWD).FWD = CT
      End If
      If CHLDCHN(BWD).FWD = 0 Then CHLDTable.LT = BWD
      Call BuildListBOX
      lstBOX.ListIndex = xZ
   End If
End If
If KeyCode = vbKeyDown Then
   If xZ < (lstBOX.ListCount - 1) Then
      CT = CHLDTable.CT:
      BWD = CHLDCHN(CT).BWD
      FWD = CHLDCHN(CT).FWD
      CHLDCHN(CT).BWD = FWD
      CHLDCHN(CT).FWD = CHLDCHN(FWD).FWD
      CHLDCHN(BWD).FWD = FWD
      CHLDCHN(FWD).FWD = CT
      CHLDCHN(FWD).BWD = BWD
      If CHLDCHN(CT).FWD = 0 Then
         CHLDTable.LT = CT
         Else
         CHLDCHN(CHLDCHN(CT).FWD).BWD = CT
      End If
      If CHLDCHN(FWD).BWD = 0 Then CHLDTable.FT = FWD
      Call BuildListBOX
      lstBOX.ListIndex = xZ
   End If
End If
End Sub

Private Sub Form_Load()
If FormLoaded(2) And SpouseForm Then
   frmFHSWSpse.Enabled = False
   Else
   If FormLoaded(3) And ParentForm Then
      frmFHSWPare.Enabled = False
      Else
      frmFHSWUPDT.Enabled = False
   End If
End If
t0 = frmFHSWUPDT.TOP: l0 = frmFHSWUPDT.Left
frmHEIGHT = frmFHSWChld.Height
frmWIDTH = frmFHSWChld.WIDTH
t1 = GetSetting("FHSWIN", "frmCHLD", "Top", 0)
If t1 = 0 Then
   frmHEIGHT = 3400: frmWIDTH = 8650
   t1 = t0 + 2715
   If (t1 + frmHEIGHT) > Screen.Height Then t1 = frmFHSWUPDT.TOP + 450
   l1 = l0 - (frmWIDTH - frmFHSWUPDT.WIDTH) / 2
   If l1 < 0 Then l1 = frmFHSWUPDT.Left: frmWIDTH = frmFHSWUPDT.WIDTH
   Else
   l1 = 1
End If
FormLoaded(3) = FormLoaded(3) Or ChildForm
frmFHSWChld.TOP = GetSetting("FHSWIN", "frmCHLD", "Top", t1)
frmFHSWChld.Left = GetSetting("FHSWIN", "frmCHLD", "Left", l1)
frmFHSWChld.WIDTH = GetSetting("FHSWIN", "frmCHLD", "Width", frmWIDTH)
frmFHSWChld.Height = GetSetting("FHSWIN", "frmCHLD", "Height", frmHEIGHT)
MsgLine.WIDTH = CMDButton(2).Left - MsgLine.Left - 50
End Sub


Private Sub Form_Resize()
If frmFHSWChld.WIDTH < 6600 Then frmFHSWChld.WIDTH = 6600
DW = frmFHSWChld.WIDTH - 9450
If frmFHSWChld.Height < 2000 Then frmFHSWChld.Height = 2000
dH = frmFHSWChld.Height - 3405
INFOLbl.WIDTH = 8775 + DW
INFOFrame.WIDTH = 9135 + DW
INFOFrame.Height = 2415 + dH
lstBOX.WIDTH = 8895 + DW
lstBOX.Height = 1815 + dH
CMDFrame.TOP = 2400 + dH
CMDFrame.WIDTH = INFOFrame.WIDTH
CMDButton(1).Left = 3840 + DW / 2
CMDButton(2).Left = 7680 + DW
MsgLine.WIDTH = CMDButton(2).Left - MsgLine.Left - 50
End Sub

Private Sub Form_Unload(Cancel As Integer)
FormLoaded(3) = FormLoaded(3) And (255 - ChildForm)
SaveSetting "FHSWIN", "frmCHLD", "Top", frmFHSWChld.TOP
SaveSetting "FHSWIN", "frmCHLD", "Left", frmFHSWChld.Left
SaveSetting "FHSWIN", "frmCHLD", "Width", frmFHSWChld.WIDTH
SaveSetting "FHSWIN", "frmCHLD", "Height", frmFHSWChld.Height
If FormLoaded(3) And ParentForm Then
   frmFHSWPare.Enabled = True
   frmFHSWPare.SetFocus
   Else
   If FormLoaded(2) And SpouseForm Then
      frmFHSWSpse.Enabled = True
      frmFHSWSpse.SetFocus
      Else
      frmFHSWUPDT.Enabled = True
      frmFHSWUPDT.SetFocus
   End If
End If
End Sub


Private Sub lstBOX_DblClick()
If Not MoveMode Then
   CMDButton(0).VALUE = True
End If
End Sub


