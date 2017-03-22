VERSION 4.00
Begin VB.Form mnuFHSWMain 
   AutoRedraw      =   -1  'True
   Caption         =   "   FHSW - Family History System - 2008"
   ClientHeight    =   1590
   ClientLeft      =   7260
   ClientTop       =   2895
   ClientWidth     =   5880
   FillColor       =   &H00FFFFFF&
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Height          =   2280
   Left            =   7200
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   5880
   Top             =   2265
   Width           =   6000
   Begin VB.CommandButton cmdLIST 
      Caption         =   "<"
      Height          =   300
      Left            =   2880
      TabIndex        =   15
      Top             =   125
      Width           =   255
   End
   Begin VB.TextBox MFCount 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   5160
      TabIndex        =   14
      Top             =   1290
      Width           =   615
   End
   Begin VB.TextBox MFCount 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   5160
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   900
      Width           =   615
   End
   Begin VB.ComboBox MFSetup 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS LineDraw"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      ItemData        =   "FHSWMAIN.frx":0000
      Left            =   1200
      List            =   "FHSWMAIN.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   120
      Width           =   1710
   End
   Begin VB.TextBox MFCount 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   5160
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   550
      Width           =   615
   End
   Begin VB.TextBox MFDesc 
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS LineDraw"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3200
      TabIndex        =   7
      Top             =   120
      Width           =   2655
   End
   Begin VB.TextBox MFName 
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS LineDraw"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   1200
      TabIndex        =   6
      Top             =   1200
      Width           =   2895
   End
   Begin VB.TextBox MFName 
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS LineDraw"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   1200
      TabIndex        =   4
      Top             =   480
      Width           =   2895
   End
   Begin VB.TextBox MFName 
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS LineDraw"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   1200
      TabIndex        =   2
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label MFLabel 
      Caption         =   "Relatives:"
      Height          =   255
      Index           =   6
      Left            =   4200
      TabIndex        =   13
      Top             =   1290
      Width           =   975
   End
   Begin VB.Label MFLabel 
      Caption         =   "Selected:"
      Enabled         =   0   'False
      Height          =   255
      Index           =   5
      Left            =   4200
      TabIndex        =   11
      Top             =   900
      Width           =   900
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   5800
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label MFLabel 
      Caption         =   "Names:"
      Height          =   255
      Index           =   4
      Left            =   4200
      TabIndex        =   8
      Top             =   550
      Width           =   900
   End
   Begin VB.Label MFLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Index File"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   3
      Left            =   0
      TabIndex        =   5
      Top             =   525
      Width           =   1200
   End
   Begin VB.Label MFLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Relation to"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   2
      Left            =   0
      TabIndex        =   3
      Top             =   1260
      Width           =   1200
   End
   Begin VB.Label MFLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Selections"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   900
      Width           =   1200
   End
   Begin VB.Label MFLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Family File"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   170
      Width           =   1200
   End
   Begin VB.Menu mnuFHSWMainO 
      Caption         =   "File"
      Index           =   10
      Begin VB.Menu mnuFHSWMainF 
         Caption         =   "Update Family Information"
         Index           =   10
      End
      Begin VB.Menu mnuFHSWMainF 
         Caption         =   "Family File Setup"
         Index           =   30
      End
      Begin VB.Menu mnuFHSWMainF 
         Caption         =   "Printer/Page Setup"
         Index           =   40
      End
      Begin VB.Menu mnuFHSWMainF 
         Caption         =   "Exit"
         Index           =   50
      End
   End
   Begin VB.Menu mnuFHSWMainO 
      Caption         =   "Relationships"
      Index           =   15
      Begin VB.Menu mnuFHSWMainRel 
         Caption         =   "Select Base"
         Index           =   10
      End
      Begin VB.Menu mnuFHSWMainRel 
         Caption         =   "Relationship Style"
         Index           =   20
         Begin VB.Menu mnuFHSWMainRelS 
            Caption         =   "Common"
            Index           =   10
         End
         Begin VB.Menu mnuFHSWMainRelS 
            Caption         =   "Civil"
            Index           =   20
         End
         Begin VB.Menu mnuFHSWMainRelS 
            Caption         =   "Latin"
            Index           =   30
         End
         Begin VB.Menu mnuFHSWMainRelS 
            Caption         =   "None"
            Index           =   40
         End
      End
      Begin VB.Menu mnuFHSWMainRel 
         Caption         =   "Calculator"
         Index           =   30
      End
   End
   Begin VB.Menu mnuFHSWMainO 
      Caption         =   "Select Records"
      Index           =   20
      Begin VB.Menu mnuFHSWMAINS 
         Caption         =   "All Records"
         Index           =   5
      End
      Begin VB.Menu mnuFHSWMAINS 
         Caption         =   "by Relationship"
         Index           =   10
         Begin VB.Menu mnuFHSWMainS10 
            Caption         =   "Ancestors"
            Index           =   10
         End
         Begin VB.Menu mnuFHSWMainS10 
            Caption         =   "Descendants"
            Index           =   20
         End
         Begin VB.Menu mnuFHSWMainS10 
            Caption         =   "Relatives"
            Index           =   30
         End
      End
      Begin VB.Menu mnuFHSWMAINS 
         Caption         =   "by Select File"
         Index           =   15
      End
      Begin VB.Menu mnuFHSWMAINS 
         Caption         =   "by Search"
         Index           =   20
      End
   End
   Begin VB.Menu mnuFHSWMainO 
      Caption         =   "Reports/Charts"
      Index           =   30
      Begin VB.Menu mnuFHSWMainR 
         Caption         =   "Relationship"
         Index           =   10
         Begin VB.Menu mnuFHSWMainR10 
            Caption         =   "Ancestor"
            Index           =   10
         End
         Begin VB.Menu mnuFHSWMainR10 
            Caption         =   "Descendant"
            Index           =   20
         End
         Begin VB.Menu mnuFHSWMainR10 
            Caption         =   "Relative"
            Index           =   30
         End
      End
      Begin VB.Menu mnuFHSWMainR 
         Caption         =   "Family Group"
         Index           =   20
         Begin VB.Menu mnuFHSWMainR20 
            Caption         =   "Blank Forms"
            Index           =   10
         End
         Begin VB.Menu mnuFHSWMainR20 
            Caption         =   "Fixed Format"
            Index           =   20
         End
         Begin VB.Menu mnuFHSWMainR20 
            Caption         =   "Free Format"
            Index           =   30
         End
      End
      Begin VB.Menu mnuFHSWMainR 
         Caption         =   "Register "
         Index           =   25
         Begin VB.Menu mnuFHSWMainR25 
            Caption         =   "Ancestor"
            Index           =   10
         End
         Begin VB.Menu mnuFHSWMainR25 
            Caption         =   "Descendant"
            Index           =   20
         End
         Begin VB.Menu mnuFHSWMainR25 
            Caption         =   "Relative"
            Index           =   30
         End
      End
      Begin VB.Menu mnuFHSWMainR 
         Caption         =   "Lists"
         Index           =   30
      End
      Begin VB.Menu mnuFHSWMainR 
         Caption         =   "Summary"
         Index           =   40
         Begin VB.Menu mnuFHSWMainR40 
            Caption         =   "Family File Summary"
            Index           =   10
         End
         Begin VB.Menu mnuFHSWMainR40 
            Caption         =   "Relationship"
            Index           =   20
            Begin VB.Menu mnuFHSWMainR41 
               Caption         =   "Ancestor"
               Index           =   10
            End
            Begin VB.Menu mnuFHSWMainR41 
               Caption         =   "Descendant"
               Index           =   20
            End
            Begin VB.Menu mnuFHSWMainR41 
               Caption         =   "Relative"
               Index           =   30
            End
         End
         Begin VB.Menu mnuFHSWMainR40 
            Caption         =   "Indexed"
            Index           =   40
         End
      End
      Begin VB.Menu mnuFHSWMainR 
         Caption         =   "Charts"
         Index           =   50
         Begin VB.Menu mnuFHSWMainR50 
            Caption         =   "Ancestor"
            Index           =   10
            Begin VB.Menu mnuFHSWMainR51 
               Caption         =   "4Gen/5Gen"
               Index           =   10
            End
            Begin VB.Menu mnuFHSWMainR51 
               Caption         =   "Map"
               Index           =   15
            End
            Begin VB.Menu mnuFHSWMainR51 
               Caption         =   "Vertical Box"
               Index           =   20
            End
            Begin VB.Menu mnuFHSWMainR51 
               Caption         =   "Horizontal Box"
               Index           =   30
            End
         End
         Begin VB.Menu mnuFHSWMainR50 
            Caption         =   "Descendant"
            Index           =   20
            Begin VB.Menu mnuFHSWMainR52 
               Caption         =   "Vertical Box"
               Index           =   10
            End
            Begin VB.Menu mnuFHSWMainR52 
               Caption         =   "Horizontal Box"
               Index           =   20
            End
         End
         Begin VB.Menu mnuFHSWMainR50 
            Caption         =   "Family Path"
            Index           =   30
         End
      End
      Begin VB.Menu mnuFHSWMainR 
         Caption         =   "Indexes"
         Index           =   60
         Begin VB.Menu mnuFHSWMainR60 
            Caption         =   "Table of Contents"
            Enabled         =   0   'False
            Index           =   10
         End
         Begin VB.Menu mnuFHSWMainR60 
            Caption         =   "Report/Chart Index"
            Enabled         =   0   'False
            Index           =   20
         End
      End
   End
   Begin VB.Menu mnuFHSWMainO 
      Caption         =   "Utilities"
      Index           =   40
      Begin VB.Menu mnuFHSWMainU 
         Caption         =   "Customization"
         Index           =   10
         Begin VB.Menu mnuFHSWMainU10 
            Caption         =   "Parameters"
            Enabled         =   0   'False
            Index           =   10
         End
         Begin VB.Menu mnuFHSWMainU10 
            Caption         =   "System Tables"
            Enabled         =   0   'False
            Index           =   20
         End
         Begin VB.Menu mnuFHSWMainU10 
            Caption         =   "Form Labels"
            Enabled         =   0   'False
            Index           =   30
         End
         Begin VB.Menu mnuFHSWMainU10 
            Caption         =   "Messages"
            Enabled         =   0   'False
            Index           =   40
         End
         Begin VB.Menu mnuFHSWMainU10 
            Caption         =   "Reports"
            Enabled         =   0   'False
            Index           =   50
         End
      End
      Begin VB.Menu mnuFHSWMainU 
         Caption         =   "Export/Import"
         Index           =   40
         Begin VB.Menu mnuFHSWMainU40 
            Caption         =   "GEDCOM"
            Index           =   10
            Begin VB.Menu mnuFHSWMainU42 
               Caption         =   "Import"
               Index           =   10
            End
            Begin VB.Menu mnuFHSWMainU42 
               Caption         =   "Export"
               Index           =   20
            End
         End
         Begin VB.Menu mnuFHSWMainU40 
            Caption         =   "MailMerge"
            Index           =   20
         End
      End
      Begin VB.Menu mnuFHSWMainU 
         Caption         =   "Family File Validation"
         Index           =   50
         Begin VB.Menu mnuFHSWMainU50 
            Caption         =   "Pointers"
            Index           =   10
         End
         Begin VB.Menu mnuFHSWMainU50 
            Caption         =   "Family Data"
            Index           =   20
         End
         Begin VB.Menu mnuFHSWMainU50 
            Caption         =   "Duplicate Data"
            Index           =   30
         End
      End
      Begin VB.Menu mnuFHSWMainU 
         Caption         =   "Conversion"
         Index           =   55
         Begin VB.Menu mnuFHSWMainU55 
            Caption         =   "to Long Address Format"
            Index           =   10
         End
         Begin VB.Menu mnuFHSWMainU55 
            Caption         =   "to Long Place Name Support"
            Index           =   20
         End
      End
      Begin VB.Menu mnuFHSWMainU 
         Caption         =   "Calculator"
         Index           =   60
         Begin VB.Menu mnuFHSWMainU60 
            Caption         =   "Date"
            Index           =   10
         End
         Begin VB.Menu mnuFHSWMainU60 
            Caption         =   "Relationship"
            Index           =   20
         End
      End
      Begin VB.Menu mnuFHSWMainU 
         Caption         =   "Index File"
         Index           =   70
      End
      Begin VB.Menu mnuFHSWMainU 
         Caption         =   "Tiny Tafel"
         Index           =   80
      End
   End
   Begin VB.Menu mnuFHSWMainAbout 
      Caption         =   "About"
      Begin VB.Menu mnuFHSWMainA 
         Caption         =   "Version"
         Enabled         =   0   'False
         Index           =   10
      End
      Begin VB.Menu mnuFHSWMainA 
         Caption         =   "Manual"
         Enabled         =   0   'False
         Index           =   20
      End
      Begin VB.Menu mnuFHSWMainA 
         Caption         =   "WWW"
         Enabled         =   0   'False
         Index           =   30
      End
   End
End
Attribute VB_Name = "mnuFHSWMain"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
DefInt A-Z

Dim RELCount As Integer, LastFAM As String
Dim frmHEIGHT, frmWIDTH

Sub UnloadFHSWForms()
If FormLoaded(9) And MSGDForm Then Unload frmFHSWMsgD
If FormLoaded(9) And MSGForm Then Unload frmFHSWMsg
If FormLoaded(2) And ListForm Then Unload frmFHSWList
If FormLoaded(1) And FileForm Then Unload frmFHSWFile
If FormLoaded(1) And ViewForm Then Unload frmFHSWView
If FormLoaded(1) And RelWForm Then Unload frmFHSWRelW
If FormLoaded(1) And SlctForm Then Unload frmFHSWSlct
If FormLoaded(1) And NdWkForm Then Unload frmFHSWNdWk
If FormLoaded(3) And IndexForm Then Unload frmFHSWINDX
If FormLoaded(1) And UPgmForm Then Unload frmFHSWUPgm
If FormLoaded(1) And RPgmForm Then Unload frmFHSWRPgm
If FormLoaded(1) And PrntForm Then Unload frmFHSWPrinter
If FormLoaded(2) And FamDefForm Then Unload frmFHSWFFDF
If FormLoaded(2) And RelCalcForm Then Unload frmFHSWRCLC
If FormLoaded(2) And DateCalcForm Then Unload frmFHSWDCLC
If FormLoaded(2) And UpdtForm Then Unload frmFHSWUPDT
If FormLoaded(2) And NotesForm Then
   Unload frmFHSWNote: If Not OKAY Then Exit Sub
End If
OKAY = -1
End Sub

Public Sub BuildFAMSetup()
MFSetup.Clear
X = FSUTABLE.FT
While X <> 0
   MFSetup.AddItem Left$(FSETUP$(X), 8)
   Y = MFSetup.NewIndex
   MFSetup.ItemData(Y) = X
   X = FSUCHN(X).FWD
Wend
X = 0: FOUND = 0
While (X < FDFHDr.SETUPS) And (Not FOUND)
   X = X + 1
   FOUND = MFSetup.List(X) = FDFSetup.NAME
Wend
If FOUND Then
   MFSetup.ListIndex = X
End If
SlecType = 0
If LastFAM <> FDFSetup.NAME Then
   ShowSetup
End If
ChangeEnable
End Sub

Private Sub BuildRelationTable()
'Call RptOptGET("EXAD", x):ExtAdoptRule = (x = 1)
Call RWRKOpen("READ", 3, 0)
If OKAY Then
   ExtAdoptRule = -1: LRELOkay = -1
   ReDim LRELATE(FF1Hdr.MAXID)
   Call PutMSG("XXXXM064")
   Call RWRKGet("BASE", RC, PTR&)
   While (0 < RC) And (RC < 9)
      If WFRELREC.id > 0 Then
         X = LRELATE(WFRELREC.id)
         If (X = 0) Or (WFGLREC.DGL = 0) Then
            X = 256 * (-WFGLREC.AGL) + WFGLREC.DGL
            If ExtAdoptRule Then X = X - 128 * ((WFRELREC.CSTAT And 5) > 0)
            LRELATE(WFRELREC.id) = X
         End If
      End If
      Call RWRKGet("WRKREC", RC, PTR&)
   Wend
   Call FamCLOSE
   Call ClearMSG
End If
End Sub

Public Sub ChangeEnable()
mnuFHSWMainF(10) = (FamFileOKAY And (Not FamUpdate))
mnuFHSWMainF(30) = (Not FamUpdate)
mnuFHSWMainF(40) = (Not FamUpdate)
mnuFHSWMainO(15) = (FamFileOKAY And (Not FamUpdate))
mnuFHSWMainO(20) = (FamFileOKAY And (Not FamUpdate))
mnuFHSWMainO(30) = (FamFileOKAY And (Not FamUpdate))
mnuFHSWMainO(40) = (Not FamUpdate)
mnuFHSWMainU(40) = (FamFileOKAY And (Not FamUpdate))
mnuFHSWMainU(50) = (FamFileOKAY And (Not FamUpdate))
mnuFHSWMainU(60) = (FamFileOKAY And (Not FamUpdate))
mnuFHSWMainU(70) = ((FormLoaded(3) And IndexForm) = 0)
MFLabel(0).Enabled = FamFileOKAY
MFLabel(1).Enabled = FamFileOKAY
MFLabel(2).Enabled = LRELOkay
MFLabel(3).Enabled = NdxFileOKAY
MFLabel(4).Enabled = FamFileOKAY
MFLabel(5).Enabled = FamFileOKAY
MFLabel(6).Enabled = LRELOkay
Select Case SlecType
   Case 0: sTYPE$ = "Full File"
   Case 1: sTYPE$ = "Ancestors"
   Case 2: sTYPE$ = "Descendants"
   Case 3: sTYPE$ = "Relatives"
   Case 4: sTYPE$ = FT$(10)
   Case 5: sTYPE$ = "Search"
End Select
MFName(1) = sTYPE$
If sTYPE$ = "" Then
   MFCount(1) = ""
   Else
   MFCount(1) = SCNT
End If
'MFName(2) = RelWDSN
If LRELOkay Then
   MFCount(2) = RELCount
   Else
   MFCount(2) = ""
End If
MFName(3) = FT$(4)
If (FormLoaded(1) > 0) Or (FamUpdate) Then
   MFSetup.Enabled = False
   cmdLIST.Enabled = False
   mnuFHSWMainO(15).Enabled = False
   mnuFHSWMainO(20).Enabled = False
   mnuFHSWMainO(30).Enabled = False
   mnuFHSWMainO(40).Enabled = False
   mnuFHSWMainU(40).Enabled = False
   mnuFHSWMainU(50).Enabled = False
   mnuFHSWMainU(60).Enabled = False
   mnuFHSWMainF(10).Enabled = False
   mnuFHSWMainF(30).Enabled = False
   If (FormLoaded(2) And ListForm) > 0 Then
      frmFHSWList.Hide
   End If
   Else
   MFSetup.Enabled = True
   cmdLIST.Enabled = FamFileOKAY
End If
mnuFHSWMain.Refresh
End Sub

Private Sub CheckRELWork()
Call RWRKOpen("READ", RelWType, 0)
RelFileOKAY = OKAY: Close
End Sub


Private Sub SelectIDList(xType)
  SlecType = xType
  Select Case SlecType
     Case 0: sTYPE$ = "Full File"
     Case 1: sTYPE$ = "Ancestors"
     Case 2: sTYPE$ = "Descendants"
     Case 3: sTYPE$ = "Relatives"
     Case 4: sTYPE$ = FT$(10)
     Case 5: sTYPE$ = "Search"
  End Select
  ChangeEnable
  Select Case SlecType
     Case 5: ' By Search
            frmFHSWSlct.Show 1
     Case Else: ' Relative Work
            Call BuildBASE(SlecType)
  End Select
  ChangeEnable
End Sub


Private Sub cmdLIST_Click()
Call frmFHSWList.ShowNames(0)
End Sub

Private Sub cmdLIST_LostFocus()
ChangeEnable
End Sub


Private Sub Form_GotFocus()
ChangeEnable
End Sub

Private Sub Form_Load()
frmHEIGHT = mnuFHSWMain.Height
frmWIDTH = mnuFHSWMain.WIDTH
t1 = GetSetting("FHSWIN", "mnuMAIN", "TOP", 0)
If t1 = 0 Then
   h00 = Screen.Height: w00 = Screen.WIDTH
   t1 = (h00 - 8200) / 2: If t1 < 0 Then t1 = 1000
   l1 = (w00 - frmWIDTH) / 2: If l1 < 0 Then l1 = 1
   Else
   t1 = 1440: l1 = 1440
End If
mnuFHSWMain.TOP = GetSetting("FHSWIN", "mnuMAIN", "Top", t1)
mnuFHSWMain.Left = GetSetting("FHSWIN", "mnuMAIN", "Left", l1)
LRELStyle = GetSetting("FHSWIN", "mnuMAIN", "RELStyle", 1)
Printer.Font = GetSetting("FHSWIN", "mnuMAIN", "PFont", "MS LineDraw")
If (0 < LRELStyle) And (LRELStyle < 5) Then
   mnuFHSWMainRelS(10 * LRELStyle).Checked = True
   Else
   LRELStyle = 0
End If
Call FDFGetSet("DEFAULT")
BuildFAMSetup
CurCMD = 0
Call ShowSetup
End Sub






Private Sub Form_LostFocus()
Call ChangeEnable
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call UnloadFHSWForms
If Not OKAY Then Cancel = -1
End Sub

Private Sub Form_Resize()
mnuFHSWMain.Height = frmHEIGHT
mnuFHSWMain.WIDTH = frmWIDTH

End Sub

Private Sub Form_Unload(Cancel As Integer)
Call FamCLOSE
SaveSetting "FHSWIN", "mnuMAIN", "Top", mnuFHSWMain.TOP
SaveSetting "FHSWIN", "mnuMAIN", "Left", mnuFHSWMain.Left
SaveSetting "FHSWIN", "mnuMAIN", "Height", mnuFHSWMain.Height
SaveSetting "FHSWIN", "mnuMAIN", "Width", mnuFHSWMain.WIDTH
SaveSetting "FHSWIN", "mnuMAIN", "RELStyle", LRELStyle
SaveSetting "FHSWIN", "mnuMAIN", "PFont", Printer.Font

End Sub


Private Sub MFSetup_Click()
If MFSetup.TEXT <> FDFSetup.NAME Then
   LSet LastFAM = ""
   Call ShowSetup
End If
End Sub
Private Sub ShowSetup()
   xX$ = MFSetup.TEXT
   If xX$ = "" Then xX$ = "DEFAULT"
   MFCount(0) = "": MFCount(1) = "": MFCount(2) = ""
   FamFileOKAY = 0: RelFileOKAY = 0: SelFileOKAY = 0
   NdxFileOKAY = 0: LRELOkay = 0
   Call FDFGetSet(xX$)
   FamSetupOKAY = OKAY
   If OKAY Then
      MFSetup.TEXT = FDFSetup.NAME
      MFDesc = FDFSetup.DESC
      MFName(2) = ""
      MFName(3) = FT$(4)
      Call FAMOpen("VLDTQ")
      FamFileOKAY = OKAY
      SlecType = 0
      If OKAY Then
         MFCount(0) = FF1Hdr.MAXID
         MFCount(1) = FF1Hdr.MAXID
         Call SelectIDList(SlecType)
      End If
      Else
      MFDesc = "":
      MFName(2) = "": MFName(3) = ""
   End If
   Call FamCLOSE
   If FormLoaded(2) And ListForm Then
      Unload frmFHSWList
   End If
   LastFAM = FDFSetup.NAME
End Sub


Private Sub MFSetup_LostFocus()
ChangeEnable
End Sub



Private Sub mnuFHSWMainA_Click(Index As Integer)
frmFHSWRPgm.Show
frmFHSWRPgm.Visible = True
End Sub

Private Sub mnuFHSWMainF_Click(Index As Integer)
Select Case Index
   Case 10: FamUpdate = -1: Load frmFHSWUPDT
   Case 20: ' Removed
   Case 30: frmFHSWFFDF.Show
   Case 40: frmFHSWPrinter.Show
   Case 50: Call FamCLOSE
            Unload mnuFHSWMain: End
End Select
End Sub








Private Sub mnuFHSWMainR_Click(Index As Integer)
If Index = 30 Then
   RptPGM = "LIST"
   RptCODE = "LIST"
   Call RptLIST("INIT")
End If
End Sub

Private Sub mnuFHSWMainR10_Click(Index As Integer)
RptPGM = "RELR": WRelType = Index / 10
'RelWDSN = FT$(RelWType + 4)
'Call CheckRELWork
Select Case Index
   Case 10: RptCODE = "ARPT"
   Case 20: RptCODE = "DRPT"
   Case 30: RptCODE = "RRPT"
End Select
Call RptRELR("INIT")
End Sub

Private Sub mnuFHSWMainR20_Click(Index As Integer)
RptPGM = "GRPR"
Select Case Index
   Case 10: RptCODE = "BLGR"
   Case 20: RptCODE = "FXGR"
   Case 30: RptCODE = "FFGR"
End Select
Call RptGRPR("INIT")
End Sub

Private Sub mnuFHSWMainR25_Click(Index As Integer)
RptPGM = "REGR": WRelType = Index / 10
'Call CheckRELWork
RptCODE = "REGR"
Call RptREGR("INIT")
End Sub

Private Sub mnuFHSWMainR40_Click(Index As Integer)
Select Case Index
   Case 10: RptPGM = "LIST"
            RptCODE = "FSUM"
            Call RptLIST("INIT")
   Case 40: RptPGM = "LIST"
            RptCODE = "ISUM"
            Call RptLIST("INIT")
End Select
End Sub

Private Sub mnuFHSWMainR41_Click(Index As Integer)
RptPGM = "RELR": WRelType = Index / 10
'Call CheckRELWork
Select Case Index
   Case 10: RptCODE = "ASUM"
   Case 20: RptCODE = "DSUM"
   Case 30: RptCODE = "RSUM"
End Select
Call RptRELR("INIT")
End Sub

Private Sub mnuFHSWMainR50_Click(Index As Integer)
If Index = 30 Then
   RptCODE = "FPTH": RptPGM = "HBOX"
   NWrkType = 3
   Call RptHBOX("INIT")
End If
End Sub

Private Sub mnuFHSWMainR51_Click(Index As Integer)
WRelType = 1
Select Case Index
   Case 10: RptCODE = "ACHT": RptPGM = "ACHT"
   Case 15: RptCODE = "AMAP": RptPGM = "ACHT"
   Case 20: RptCODE = "AVBX": RptPGM = "VBOX"
   Case 30: RptCODE = "AHBX": RptPGM = "HBOX"
            NWrkType = 1
End Select
Select Case RptPGM
   Case "ACHT": Call RptACHT("INIT")
   Case "VBOX": Call RptVBOX("INIT")
   Case "HBOX": Call RptHBOX("INIT")
End Select
End Sub

Private Sub mnuFHSWMainR52_Click(Index As Integer)
WRelType = 2
Select Case Index
   Case 10: RptCODE = "DVBX": RptPGM = "VBOX"
   Case 20: RptCODE = "DHBX": RptPGM = "HBOX"
            NWrkType = 2
End Select
Select Case RptPGM
   Case "VBOX": Call RptVBOX("INIT")
   Case "HBOX": Call RptHBOX("INIT")
End Select
End Sub

Private Sub mnuFHSWMainRel_Click(Index As Integer)
Select Case Index
   Case 10: ' Change Relationship Base
            RelWType = 3: WRelType = RelWType
            RelWDSN = FT$(RelWType + 4)
            RelFileOKAY = False: RELCount = 0
            Call frmFHSWRelW.BUILD(RelWType)
            If OKAY Then
               'RelWDSN = FT$(RelWType + 4)
               RELCount = WFHDR.NUMREL
               Call FAMOpen("READ")
               Call FF1GetRec(WFBASE)
               MFName(2) = FMTNAME$(1)
               BuildRelationTable
               RelChanged = -1
            End If
   Case 30: Call frmFHSWRCLC.RelCALC(0)
End Select
End Sub

Private Sub mnuFHSWMainRelS_Click(Index As Integer)
If (LRELStyle <> (Index / 10)) Then RelChanged = -1
If LRELStyle > 0 Then
   mnuFHSWMainRelS(10 * LRELStyle).Checked = False
End If
LRELStyle = Index / 10
mnuFHSWMainRelS(Index).Checked = True
End Sub

Private Sub mnuFHSWMainS_Click(Index As Integer)
  Select Case Index
     Case 5: X = 0
     Case 10: X = -1
     Case 15: X = 4
     Case 20: X = 5
  End Select
  If X >= 0 Then Call SelectIDList(X)
End Sub


Private Sub mnuFHSWMainS10_Click(Index As Integer)
WRelType = Index / 10
Call SelectIDList(WRelType)
End Sub


Private Sub mnuFHSWMainU_Click(Index As Integer)
Select Case Index:
       Case 70: frmFHSWINDX.Show
       Case 80: UtilPGM = "TTFL"
                Call UtilTTFL("INIT")
End Select
End Sub

Private Sub mnuFHSWMainU40_Click(Index As Integer)
Select Case Index
   Case 20: 'MailMerge Export
            UtilPGM = "MLMG": RptCODE = "MMEX"
            Call UtilMLMG("INIT")
End Select
End Sub


Private Sub mnuFHSWMainU42_Click(Index As Integer)
Select Case Index
   Case 10: 'GEDCOM Import
            UtilPGM = "GDCM": RptCODE = "GDIM"
            Call UtilGDCM("INIT")
   Case 20: 'GEDCOM Export
            UtilPGM = "GDCM": RptCODE = "GDEX"
            Call UtilGDCM("INIT")
End Select
End Sub


Private Sub mnuFHSWMainU50_Click(Index As Integer)
Select Case Index
   Case 10: 'Pointer Validation
            RptCODE = "PVLD"
   Case 20: 'DATA Validation
            RptCODE = "DVLD"
   Case 30: 'Pointer Validation
            RptCODE = "MDAT"
End Select
UtilPGM = "VLDT": Call UtilVLDT("INIT")

End Sub


Private Sub mnuFHSWMainU55_Click(Index As Integer)
Select Case Index
   Case 10: Call modFHSWGdcm.ConvertADRS
   Case 20: Call PgmPREP("FHSGDCM")
            If OKAY Then
               Call FF1Open("WRITE")
               If OKAY Then
                  If FF1Hdr.REVLVL = 1 Then
                     Call PutMSG("File already has support for Long Place Names...")
                     Else
                     Call PutMSG("Allow File to have Long Place Names?...(Y/N)")
                     If CA$ = SNGLKEY$(1) Then
                        FF1Hdr.REVLVL = 1: Put #1, 1, FF1Hdr
                        Call FamCLOSE
                        SysVAR.MVAR = FDFSetup.NAME
                        Call PutMSG("GDCMM057")
                     End If
                  End If
               End If
               Call FamCLOSE
            End If
End Select
End Sub


Private Sub mnuFHSWMainU60_Click(Index As Integer)
Select Case Index
   Case 10: frmFHSWDCLC.Show
   Case 20: Call frmFHSWRCLC.RelCALC(0)
End Select
End Sub



