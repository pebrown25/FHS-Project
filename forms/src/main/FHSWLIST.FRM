VERSION 4.00
Begin VB.Form frmFHSWList 
   Caption         =   "FHSWLIST - Name List"
   ClientHeight    =   4050
   ClientLeft      =   7755
   ClientTop       =   5235
   ClientWidth     =   4725
   Height          =   4455
   Left            =   7695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   Top             =   4890
   Width           =   4845
   Begin VB.CommandButton cmdSWITCH 
      Caption         =   "><"
      Height          =   315
      Left            =   3400
      TabIndex        =   6
      Top             =   0
      Width           =   400
   End
   Begin VB.ComboBox SortSEQ 
      Height          =   315
      ItemData        =   "FHSWLIST.frx":0000
      Left            =   2200
      List            =   "FHSWLIST.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   0
      Width           =   1150
   End
   Begin VB.TextBox NameCNT 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton cmdCANCEL 
      Cancel          =   -1  'True
      Caption         =   "CANCEL"
      Height          =   315
      Left            =   3840
      TabIndex        =   1
      Top             =   0
      Width           =   855
   End
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
      Height          =   3630
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   4575
   End
   Begin VB.Label Labels 
      Caption         =   "Sequence"
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   4
      Top             =   30
      Width           =   735
   End
   Begin VB.Label Labels 
      Caption         =   "Names"
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   3
      Top             =   30
      Width           =   735
   End
End
Attribute VB_Name = "frmFHSWList"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
DefInt A-Z
Dim SortSequence As Integer, OnlySHOW As Integer
Dim ShowSelected
Dim ListID As Integer, ReShowFORM
Private Sub BuildNameLIST()
SortSEQ.Enabled = False
lstBOX.Clear
lstBOX.MousePointer = 11
X$ = Space$(100): a = 0
AddREL = (LRELOkay And (0 < LRELStyle) And (LRELStyle < 4))
If AddREL Then
   xrel$ = Space$(15): xr$ = Space$(20)
   ENV.RELRULES = (ENV.RELRULES And 252) Or (LRELStyle - 1)
   Else
   xr$ = ""
End If
iof = Not ENV.FFOPEN
If iof Then Call FAMOpen("READ")
MousePointer = 11
ReDim Preserve SelSET(FF1Hdr.MAXID)
For xX = 1 To FF1Hdr.MAXID
   If ShowSelected Then
      getrec = (SelSET(xX) <> 0)
      Else
      getrec = -1
   End If
   If getrec Then
      FF1GetRec (xX)
      If AddREL Then
         Call GetRelation(xX, xrel$)
         If RTrim$(xrel$) <> "" Then
            LSet xr$ = " [" + RTrim$(xrel$) + "]"
            Else
            LSet xr$ = ""
        End If
      End If
      If (FF1Rec.BY > 0) Or (FF1Rec.DY > 0) Then
         xB$ = "": xd$ = ""
         If FF1Rec.BY > 0 Then xB$ = fnx$(FF1Rec.BY)
         If FF1Rec.DY > 0 Then xd$ = fnx$(FF1Rec.DY)
         xBD$ = " (" + xB$ + "-" + xd$ + ")"
         Else
         xBD$ = ""
      End If
      Select Case SortSequence
         Case 0: LSet X$ = SFILL$(xX, 5) + " " + FMTNAME$(1) + xBD$ + xr$
         Case 1: LSet X$ = FMTNAME$(2) + xBD$ + xr$
         Case 2: LSet X$ = FMTNAME$(1) + xBD$ + xr$
      End Select
      lstBOX.AddItem RTrim$(X$)
      lstBOX.ItemData(lstBOX.NewIndex) = xX
      If xX Mod 10 = 0 Then
         NameCNT.TEXT = xX: NameCNT.Refresh
      End If
      If a = 27 Then xX = FF1Hdr.MAXID
   End If
Next xX
MousePointer = 0
If iof Then Call FamCLOSE
NameCNT.TEXT = SFILL$(lstBOX.ListCount, 5)
lstBOX.MousePointer = 0
SortSEQ.Enabled = True
If (ListID > 0) Then
   FOUND = 0: Z = 0
   While (Not FOUND) And (Z < lstBOX.ListCount)
      FOUND = (ListID = lstBOX.ItemData(Z))
      If Not FOUND Then Z = Z + 1
   Wend
   If FOUND Then lstBOX.ListIndex = Z
End If
ReShowFORM = -1
End Sub


Private Sub GetRelation(id, xrel$)
If LRELATE(id) <> 0 Then
   X = LRELATE(id): Y = X And 128
   AGL = -(X \ 256): DGL = X And 127
   RGL = AGL + DGL
   LSet xrel$ = Relation$("M", FF1Rec.SEX, AGL, RGL, RRELCD$())
   If ExtAdoptRule Then
      If (X And 128) > 0 Then
         LSet xrel$ = "*" + xrel$
      End If
   End If
   Else
   AGL = 0: DGL = 0: RGL = 0: LSet xrel$ = ""
End If
End Sub


Public Sub SelectName(IDNUM)
a = 27: OnlySHOW = False
If (lstBOX.ListCount = 0) Then
   X$ = Space$(57): ENV.HUSBSURNM = 0
   SortSEQ.ListIndex = SortSequence
   Call BuildNameLIST
End If
If (IDNUM > 0) Then
   FOUND = 0: Z = 0
   While (Not FOUND) And (Z < lstBOX.ListCount)
      FOUND = (IDNUM = lstBOX.ItemData(Z))
      If Not FOUND Then Z = Z + 1
   Wend
   If FOUND Then lstBOX.ListIndex = Z: ListID = IDNUM
End If
ShowForm:
ReShowFORM = 0
Show 1
If ReShowFORM Then GoTo ShowForm
If lstBOX.ListIndex < 0 Then a = 27
If (a <> 27) Then
   IDNUM = lstBOX.ItemData(lstBOX.ListIndex)
End If
ListID = IDNUM
End Sub

Public Sub ShowNames(OPT)
a = 27: OnlySHOW = True
If OPT < 0 Then ShowSelected = -1: lstBOX.Clear
If (lstBOX.ListCount <> FF1Hdr.MAXID) Or (RelChanged) Then
   ENV.HUSBSURNM = 0
   If SortSEQ.ListIndex < 0 Then
      SortSequence = 0
      SortSEQ.ListIndex = 0
   End If
   'Show 1
   frmFHSWList.Refresh
   Call BuildNameLIST
   RelChanged = 0
   Show 1
   Else
   Show 1
End If
End Sub


Private Sub cmdCANCEL_Click()
a = 27: frmFHSWList.Hide
End Sub



Private Sub cmdSWITCH_Click()
ShowSelected = Not ShowSelected
'If (SCNT <> FF1Hdr.MAXID) Or (lstBOX.ListCount <> (FF1Hdr.MAXID)) Then
   Call BuildNameLIST
'End If
End Sub


Private Sub Form_Load()
FormLoaded(2) = FormLoaded(2) Or ListForm
' Set Default Form Location on Screen
frmHEIGHT = 3990
frmWIDTH = 6870
t1 = GetSetting("FHSWIN", "frmLIST", "TOP", 0)
If t1 = 0 Then
   t00 = mnuFHSWMain.TOP: w00 = mnuFHSWMain.WIDTH
   l00 = mnuFHSWMain.Left
   t1 = (t00 + 2515)
   l1 = l00 + (w00 - frmWIDTH) / 2: If l1 < 0 Then l1 = 1
   Else
   t1 = 1440: l1 = 1440
End If
frmFHSWList.TOP = GetSetting("FHSWIN", "frmLIST", "Top", t1)
frmFHSWList.Left = GetSetting("FHSWIN", "frmLIST", "Left", l1)
frmFHSWList.WIDTH = GetSetting("FHSWIN", "frmLIST", "Width", frmWIDTH)
frmFHSWList.Height = GetSetting("FHSWIN", "frmLIST", "Height", frmHEIGHT)
End Sub

Private Sub Form_Resize()
lstBOX.WIDTH = frmFHSWList.WIDTH - 340
X = lstBOX.Left + lstBOX.WIDTH
Labels(1).Left = X - 3255
SortSEQ.Left = X - 2495
cmdSWITCH.Left = X - 1295
cmdCANCEL.Left = X - 855
lstBOX.Height = frmFHSWList.Height - 780
End Sub


Private Sub Form_Unload(Cancel As Integer)
MousePointer = 11
FormLoaded(2) = FormLoaded(2) And (255 - ListForm)
SaveSetting "FHSWIN", "frmLIST", "Top", frmFHSWList.TOP
SaveSetting "FHSWIN", "frmLIST", "Left", frmFHSWList.Left
SaveSetting "FHSWIN", "frmLIST", "Width", frmFHSWList.WIDTH
SaveSetting "FHSWIN", "frmLIST", "Height", frmFHSWList.Height
End Sub



Private Sub lstBOX_Click()
If Not OnlySHOW Then
   frmFHSWList.Hide: a = 13
End If
End Sub


Private Sub lstBOX_KeyDown(KeyCode As Integer, Shift As Integer)
'Select Case KeyCode
   'Case vbKeyDelete: ' Delete Key
   '         x = lstBOX.ItemData(lstBOX.ListIndex)
   '         xp = SelSET(x) And 1
   '         xs = (SelSET(x) And 2) / 2
   '         SelSET(x) = 0
   '         PSCOUNT = PSCOUNT - xp
   '         SSCOUNT = SSCOUNT - xs
   '         SCNT = SCNT - 1
   '         ShowSelected = True
   '         BuildNameList
'End Select
            
End Sub

Private Sub NameCNT_Change()
NameCNT.TEXT = SFILL$(lstBOX.ListCount, 5)
NameCNT.SelStart = 0
NameCNT.SelLength = 0
End Sub

Private Sub NameCNT_GotFocus()
NameCNT.SelStart = 0
NameCNT.SelLength = 0
End Sub


Private Sub SortSEQ_Click()
If SortSequence <> SortSEQ.ListIndex Then
   SortSequence = SortSEQ.ListIndex
   BuildNameLIST
End If
End Sub


