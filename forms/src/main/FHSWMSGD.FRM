VERSION 4.00
Begin VB.Form frmFHSWMsgD 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FHSWMSGD - Get Data"
   ClientHeight    =   7800
   ClientLeft      =   8265
   ClientTop       =   2160
   ClientWidth     =   5100
   Height          =   8205
   KeyPreview      =   -1  'True
   Left            =   8205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   Top             =   1815
   Width           =   5220
   Begin VB.Frame TTHdrFrame 
      Caption         =   "TTAFEL Header"
      Height          =   2775
      Left            =   360
      TabIndex        =   45
      Top             =   4680
      Visible         =   0   'False
      Width           =   4455
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   27
         Top             =   2400
         Width           =   375
      End
      Begin VB.TextBox TTHdrTxt 
         Height          =   285
         Index           =   0
         Left            =   960
         TabIndex        =   15
         Top             =   195
         Width           =   3375
      End
      Begin VB.TextBox TTHdrTxt 
         Height          =   285
         Index           =   1
         Left            =   960
         TabIndex        =   16
         Top             =   435
         Width           =   3375
      End
      Begin VB.TextBox TTHdrTxt 
         Height          =   285
         Index           =   2
         Left            =   960
         TabIndex        =   17
         Top             =   675
         Width           =   3375
      End
      Begin VB.TextBox TTHdrTxt 
         Height          =   285
         Index           =   4
         Left            =   960
         TabIndex        =   19
         Text            =   "000 000-0000"
         Top             =   1275
         Width           =   1815
      End
      Begin VB.TextBox TTHdrTxt 
         Height          =   285
         Index           =   5
         Left            =   960
         TabIndex        =   20
         Top             =   1635
         Width           =   1815
      End
      Begin VB.TextBox TTHdrTxt 
         Height          =   285
         Index           =   6
         Left            =   3360
         TabIndex        =   21
         Top             =   1635
         Width           =   975
      End
      Begin VB.TextBox TTHdrTxt 
         Height          =   285
         Index           =   7
         Left            =   960
         TabIndex        =   22
         Top             =   1920
         Width           =   3375
      End
      Begin VB.TextBox TTHdrTxt 
         Height          =   285
         Index           =   8
         Left            =   960
         TabIndex        =   23
         Top             =   2160
         Width           =   3375
      End
      Begin VB.TextBox TTHdrTxt 
         Height          =   285
         Index           =   9
         Left            =   960
         TabIndex        =   24
         Top             =   2400
         Width           =   3375
      End
      Begin VB.TextBox TTHdrTxt 
         Height          =   285
         Index           =   3
         Left            =   960
         TabIndex        =   18
         Top             =   920
         Width           =   3375
      End
      Begin VB.Label TTHdrLbl 
         Caption         =   "Name:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   51
         Top             =   240
         Width           =   615
      End
      Begin VB.Label TTHdrLbl 
         Caption         =   "Address:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   50
         Top             =   480
         Width           =   735
      End
      Begin VB.Label TTHdrLbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Phone:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   49
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label TTHdrLbl 
         Alignment       =   2  'Center
         Caption         =   "Service:"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   48
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label TTHdrLbl 
         Alignment       =   1  'Right Justify
         Caption         =   "BBS:"
         Height          =   255
         Index           =   5
         Left            =   2640
         TabIndex        =   47
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label TTHdrLbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Remarks:"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   46
         Top             =   1920
         Width           =   735
      End
   End
   Begin VB.Frame FrameGetRANGE 
      BackColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Visible         =   0   'False
      Width           =   4335
      Begin VB.TextBox txtID1 
         Height          =   285
         Left            =   2520
         MaxLength       =   4
         TabIndex        =   0
         Top             =   150
         Width           =   495
      End
      Begin VB.TextBox txtID2 
         Height          =   285
         Left            =   3720
         MaxLength       =   4
         TabIndex        =   1
         Top             =   150
         Width           =   495
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Last ="
         Height          =   270
         Index           =   7
         Left            =   3240
         TabIndex        =   44
         Top             =   180
         Width           =   495
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "First ="
         Height          =   270
         Index           =   6
         Left            =   2040
         TabIndex        =   43
         Top             =   180
         Width           =   495
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Enter Range of ID's:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   5
         Left            =   120
         TabIndex        =   42
         Top             =   180
         Width           =   1815
      End
   End
   Begin VB.Frame frameGetSETUP 
      BackColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   0
      TabIndex        =   39
      Top             =   1200
      Visible         =   0   'False
      Width           =   3255
      Begin VB.TextBox txtSETUP 
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
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   4
         Top             =   150
         Width           =   1335
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Enter Setup Name:"
         Height          =   270
         Index           =   3
         Left            =   240
         TabIndex        =   40
         Top             =   180
         Width           =   1455
      End
   End
   Begin VB.Frame GEDCOMHeader 
      Caption         =   "GEDCOM Header"
      Height          =   2055
      Left            =   360
      TabIndex        =   32
      Top             =   2520
      Visible         =   0   'False
      Width           =   4455
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox GCHText 
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
         Left            =   1200
         TabIndex        =   6
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox GCHText 
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
         Left            =   1200
         TabIndex        =   7
         Top             =   480
         Width           =   3135
      End
      Begin VB.TextBox GCHText 
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
         Index           =   2
         Left            =   1200
         TabIndex        =   8
         Top             =   720
         Width           =   3135
      End
      Begin VB.TextBox GCHText 
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
         Left            =   1200
         TabIndex        =   9
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox GCHText 
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
         Index           =   4
         Left            =   1200
         TabIndex        =   10
         Text            =   "FHSWIN"
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox GCHText 
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
         Index           =   5
         Left            =   3240
         TabIndex        =   11
         Top             =   1215
         Width           =   1095
      End
      Begin VB.TextBox GCHText 
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
         Left            =   1200
         TabIndex        =   12
         Top             =   1425
         Width           =   3135
      End
      Begin VB.TextBox GCHText 
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
         Index           =   7
         Left            =   600
         TabIndex        =   14
         Top             =   1680
         Width           =   3735
      End
      Begin VB.Label GCHLbl 
         Caption         =   "Submitted by:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   38
         Top             =   260
         Width           =   975
      End
      Begin VB.Label GCHLbl 
         Caption         =   "Address:"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   37
         Top             =   500
         Width           =   735
      End
      Begin VB.Label GCHLbl 
         Caption         =   "Phone:"
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   36
         Top             =   980
         Width           =   615
      End
      Begin VB.Label GCHLbl 
         Caption         =   "Source:"
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   35
         Top             =   1250
         Width           =   615
      End
      Begin VB.Label GCHLbl 
         Caption         =   "Destination:"
         Height          =   255
         Index           =   4
         Left            =   2280
         TabIndex        =   34
         Top             =   1245
         Width           =   975
      End
      Begin VB.Label GCHLbl 
         Caption         =   "Comments:"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   33
         Top             =   1440
         Width           =   855
      End
   End
   Begin VB.Frame frameGetYEARS 
      BackColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   0
      TabIndex        =   29
      Top             =   600
      Visible         =   0   'False
      Width           =   5055
      Begin VB.TextBox txtYEAR2 
         Height          =   285
         Left            =   4440
         MaxLength       =   4
         TabIndex        =   3
         Top             =   150
         Width           =   495
      End
      Begin VB.TextBox txtYEAR1 
         Height          =   285
         Left            =   3720
         MaxLength       =   4
         TabIndex        =   2
         Top             =   150
         Width           =   495
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "-"
         Height          =   255
         Index           =   2
         Left            =   4300
         TabIndex        =   31
         Top             =   170
         Width           =   135
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Enter Beginning and Ending Years for Time Line:"
         Height          =   270
         Index           =   1
         Left            =   120
         TabIndex        =   30
         Top             =   175
         Width           =   3495
      End
   End
   Begin VB.Frame frameGetBRID 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   13
      Top             =   1800
      Visible         =   0   'False
      Width           =   3925
      Begin VB.CommandButton cmdLIST 
         BackColor       =   &H00C0C0C0&
         Caption         =   "<"
         Height          =   285
         Left            =   3480
         TabIndex        =   25
         Top             =   150
         Width           =   375
      End
      Begin VB.TextBox txtBRID 
         Height          =   285
         Left            =   2880
         MaxLength       =   5
         TabIndex        =   5
         Top             =   150
         Width           =   615
      End
      Begin VB.Label Label 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Enter Base Record Id for the Chart:"
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   180
         Width           =   2610
      End
   End
End
Attribute VB_Name = "frmFHSWMsgD"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
DefInt A-Z
Dim Center
Public Sub GetGCHdrInfo(HLIT$(), OPT$)
GEDCOMHeader.TOP = 0: GEDCOMHeader.Left = 0
frmFHSWMsgD.Height = GEDCOMHeader.Height + 400
frmFHSWMsgD.WIDTH = GEDCOMHeader.WIDTH + 100
frmFHSWMsgD.Left = Center - (GEDCOMHeader.WIDTH / 2)
For X = 0 To 7: GCHText(X) = HLIT$(X + 1): Next X
GEDCOMHeader.Visible = True
If OPT$ = "SHOW" Then frmFHSWMsgD.Show: Exit Sub
a = 27: CA$ = Chr$(27): OKAY = 0
GEDCOMHeader.Visible = True
frmFHSWMsgD.Show 1
If ((FormLoaded(9) And MSGDForm) = 0) Then Exit Sub
For X = 0 To 7: HLIT$(X + 1) = GCHText(X): Next X
Unload frmFHSWMsgD
End Sub


Public Sub GetBRID(X)
frameGetBRID.TOP = 0
frmFHSWMsgD.Height = frameGetBRID.Height + 450
frmFHSWMsgD.WIDTH = frameGetBRID.WIDTH + 100
frmFHSWMsgD.Left = Center - (frmFHSWMsgD.WIDTH / 2)
txtbrid.TEXT = SFILL$(X, 5)
frameGetBRID.Visible = True
frmFHSWMsgD.Show 1
X = Val(txtbrid.TEXT)
If ((FormLoaded(9) And MSGDForm) = 0) Then Exit Sub
Unload frmFHSWMsgD
End Sub
Public Sub GetIDRange(ID1, ID2)
FrameGetRANGE.TOP = 0
frmFHSWMsgD.Height = FrameGetRANGE.Height + 450
frmFHSWMsgD.WIDTH = FrameGetRANGE.WIDTH + 100
frmFHSWMsgD.Left = Center - (frmFHSWMsgD.WIDTH / 2)
If ID1 = 0 Then ID1 = 1
If (ID2 = 0) Or (ID2 > FF1Hdr.MAXID) Then ID2 = FF1Hdr.MAXID
FrameGetRANGE.Visible = True
a = 27: OKAY = 0
frmFHSWMsgD.Show 1
ID1 = Val(txtID1.TEXT)
ID2 = Val(txtID2.TEXT)
If ((FormLoaded(9) And MSGDForm) = 0) Then Exit Sub
Unload frmFHSWMsgD
End Sub

Public Sub GetSETUP(xname$)
frameGetSETUP.TOP = 0
frmFHSWMsgD.Height = frameGetSETUP.Height + 450
frmFHSWMsgD.WIDTH = frameGetSETUP.WIDTH + 100
frmFHSWMsgD.Left = Center - (frmFHSWMsgD.WIDTH / 2)
frameGetSETUP.Visible = True
frmFHSWMsgD.Show 1
xname$ = txtSETUP.TEXT
If ((FormLoaded(9) And MSGDForm) = 0) Then Exit Sub
Unload frmFHSWMsgD
End Sub


Public Sub GetTLYears(byr, eyr)
frameGetYEARS.TOP = 0
frmFHSWMsgD.Height = frameGetYEARS.Height + 450
frmFHSWMsgD.WIDTH = frameGetYEARS.WIDTH + 100
frmFHSWMsgD.Left = Center - (frmFHSWMsgD.WIDTH / 2)
If eyr = 0 Then eyr = DATEY(DATE$, 0)
If (byr = 0) And (eyr > 200) Then byr = eyr - 200
txtyear1.TEXT = ZFILL$(byr, 4)
txtyear2.TEXT = ZFILL$(eyr, 4)
frameGetYEARS.Visible = True
a = 27: OKAY = 0
frmFHSWMsgD.Show 1
byr = Val(txtyear1.TEXT)
eyr = Val(txtyear2.TEXT)
If ((FormLoaded(9) And MSGDForm) = 0) Then Exit Sub
Unload frmFHSWMsgD
End Sub


Public Sub GetTTHdrInfo(HLIT$(), OPT$)
TTHdrFrame.TOP = 0: TTHdrFrame.Left = 0
frmFHSWMsgD.Height = TTHdrFrame.Height + 370
frmFHSWMsgD.WIDTH = TTHdrFrame.WIDTH
frmFHSWMsgD.Left = Center - (TTHdrFrame.WIDTH / 2)
For X = 0 To 7: TTHdrTxt(X) = HLIT$(X + 1): Next X
TTHdrFrame.Visible = True
If OPT$ = "SHOW" Then frmFHSWMsgD.Show: Exit Sub
a = 27: CA$ = Chr$(27): OKAY = 0
TTHdrFrame.Visible = True
frmFHSWMsgD.Show 1
If ((FormLoaded(9) And MSGDForm) = 0) Then Exit Sub
For X = 0 To 9: HLIT$(X + 1) = TTHdrTxt(X): Next X
Unload frmFHSWMsgD
End Sub

Private Sub cmdLIST_Click()
Call frmFHSWList.SelectName(X)
If a <> 27 Then
   txtbrid.TEXT = SFILL$(X, 5)
   frmFHSWMsgD.Visible = False
End If
End Sub


Public Sub cmdOK_Click(Index)
a = 13: OKAY = -1
frmFHSWMsgD.Visible = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
   Case 13: a = 13: OKAY = -1
            frmFHSWMsgD.Visible = False
   Case 27: a = 27: OKAY = 0
            frmFHSWMsgD.Visible = False
End Select
End Sub

Private Sub Form_Load()
FormLoaded(9) = FormLoaded(9) Or MSGDForm
' Set Default Form Location on Screen
t1 = GetSetting("FHSWIN", "frmMSG", "TOP", 0)
If t1 = 0 Then
   t1 = 0.75 * Screen.Height
   c1 = 0.5 * Screen.WIDTH
End If
frmFHSWMsgD.TOP = GetSetting("FHSWIN", "frmMSGD", "Top", t1)
Center = GetSetting("FHSWIN", "frmMSGD", "Center", c1)
frmFHSWMsgD.Left = Center - (frmFHSWMsgD.WIDTH \ 2)
a = 27: OKAY = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSetting "FHSWIN", "frmMSGD", "Top", frmFHSWMsgD.TOP
Center = frmFHSWMsgD.Left + (frmFHSWMsgD.WIDTH / 2)
SaveSetting "FHSWIN", "frmMSGD", "Center", Center
FormLoaded(9) = FormLoaded(9) And (255 - MSGDForm)
End Sub


Private Sub frameGetBRID_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   X = Val(txtbrid.TEXT)
   If X > 0 Then
      frmFHSWMsgD.Visible = False
   End If
End Sub

Private Sub Label_Click(Index As Integer)
Select Case Index
   Case 0: ' Click on GetBRID
         X = Val(txtbrid.TEXT)
         If X > 0 Then a = 13
         frmFHSWMsgD.Visible = False
   Case 1: ' click on GetYEARs
         a = 13: frmFHSWMsgD.Visible = False
   Case 3: ' Click on GetSETUP
         X = Len(RTrim$(txtSETUP))
         If X > 0 Then a = 13
         frmFHSWMsgD.Visible = False
End Select
End Sub


Private Sub txtbrid_GotFocus()
txtbrid.SelStart = 0
txtbrid.SelLength = Len(txtbrid.TEXT)
End Sub


Private Sub txtSETUP_GotFocus()
txtSETUP.SelStart = 0
txtSETUP.SelLength = Len(txtSETUP.TEXT)
End Sub

Private Sub txtSETUP_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
End Sub


Private Sub txtyear1_GotFocus()
txtyear1.SelStart = 0
txtyear1.SelLength = Len(txtyear1.TEXT)
End Sub


Private Sub txtyear2_GotFocus()
txtyear2.SelStart = 0
txtyear2.SelLength = Len(txtyear2.TEXT)
End Sub


