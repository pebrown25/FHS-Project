VERSION 4.00
Begin VB.Form frmFHSWSpse 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FHSWSPSE - Spouse Information"
   ClientHeight    =   4830
   ClientLeft      =   3240
   ClientTop       =   1935
   ClientWidth     =   4950
   Height          =   5235
   Left            =   3180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   Top             =   1590
   Width           =   5070
   Begin VB.Frame FRMSGFrame 
      Height          =   615
      Left            =   0
      TabIndex        =   43
      Top             =   600
      Width           =   4935
      Begin VB.Label FRLabel 
         Alignment       =   2  'Center
         Caption         =   "No Family Records Present"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Frame FRMNUframe 
      Height          =   615
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   4935
      Begin VB.CommandButton cmdFRMNU 
         Caption         =   "Add Spouse"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdFRMNU 
         Caption         =   "Search"
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdFRMNU 
         Caption         =   "RELCalc"
         Height          =   255
         Index           =   3
         Left            =   2160
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdFRMNU 
         Cancel          =   -1  'True
         Caption         =   "Return"
         Height          =   255
         Index           =   5
         Left            =   4080
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdFRMNU 
         Caption         =   "Date Calc"
         Height          =   255
         Index           =   4
         Left            =   3120
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame FRMIFrame 
      Height          =   3015
      Left            =   0
      TabIndex        =   24
      Top             =   1200
      Width           =   4935
      Begin VB.CommandButton cmdSHOW 
         Caption         =   "Show"
         Default         =   -1  'True
         Height          =   255
         Left            =   2280
         TabIndex        =   10
         Top             =   165
         Width           =   615
      End
      Begin VB.TextBox FRValue 
         Height          =   285
         Index           =   2
         Left            =   1320
         TabIndex        =   8
         Top             =   160
         Width           =   495
      End
      Begin VB.TextBox FRValue 
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
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   34
         Top             =   550
         Width           =   3375
      End
      Begin VB.TextBox FRValue 
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
         Height          =   285
         Index           =   5
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   33
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox FRValue 
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
         Index           =   6
         Left            =   4320
         MultiLine       =   -1  'True
         TabIndex        =   32
         Top             =   1140
         Width           =   375
      End
      Begin VB.TextBox FRValue 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   31
         Top             =   1140
         Width           =   855
      End
      Begin VB.TextBox FRValue 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   2760
         MultiLine       =   -1  'True
         TabIndex        =   30
         Top             =   1140
         Width           =   855
      End
      Begin VB.CommandButton cmdSPID 
         Caption         =   "<"
         Height          =   255
         Left            =   1800
         TabIndex        =   9
         Top             =   160
         Width           =   255
      End
      Begin VB.TextBox FRValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   15
         Left            =   3720
         MultiLine       =   -1  'True
         TabIndex        =   29
         Top             =   1605
         Width           =   495
      End
      Begin VB.TextBox FRValue 
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
         Index           =   14
         Left            =   1440
         TabIndex        =   16
         Top             =   2640
         Width           =   3015
      End
      Begin VB.TextBox FRValue 
         Height          =   315
         Index           =   13
         Left            =   2160
         TabIndex        =   15
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox FRValue 
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
         Left            =   720
         TabIndex        =   14
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox FRValue 
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
         Index           =   11
         Left            =   1440
         TabIndex        =   13
         Top             =   1920
         Width           =   3015
      End
      Begin VB.TextBox FRValue 
         Height          =   315
         Index           =   10
         Left            =   2160
         TabIndex        =   12
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox FRValue 
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
         Index           =   9
         Left            =   720
         TabIndex        =   11
         Text            =   "Married"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   4920
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label FRLabel 
         Caption         =   "Spouse ID:"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   39
         Top             =   200
         Width           =   855
      End
      Begin VB.Label FRLabel 
         Caption         =   "Name:"
         Height          =   255
         Index           =   4
         Left            =   720
         TabIndex        =   38
         Top             =   550
         Width           =   495
      End
      Begin VB.Label FRLabel 
         Caption         =   "Age:"
         Height          =   255
         Index           =   5
         Left            =   3840
         TabIndex        =   37
         Top             =   1140
         Width           =   375
      End
      Begin VB.Label FRLabel 
         Caption         =   "Born:"
         Height          =   255
         Index           =   6
         Left            =   720
         TabIndex        =   36
         Top             =   1140
         Width           =   495
      End
      Begin VB.Label FRLabel 
         Caption         =   "Died:"
         Height          =   255
         Index           =   7
         Left            =   2280
         TabIndex        =   35
         Top             =   1140
         Width           =   495
      End
      Begin VB.Label FRLabel 
         Caption         =   "for              years"
         Height          =   255
         Index           =   3
         Left            =   3480
         TabIndex        =   28
         Top             =   1605
         Width           =   1335
      End
      Begin VB.Label FRLabel 
         Caption         =   "Place:"
         Height          =   255
         Index           =   9
         Left            =   960
         TabIndex        =   26
         Top             =   2640
         Width           =   495
      End
      Begin VB.Label FRLabel 
         Caption         =   "Place:"
         Height          =   255
         Index           =   8
         Left            =   960
         TabIndex        =   25
         Top             =   1920
         Width           =   495
      End
   End
   Begin VB.Frame FRIDFrame 
      Height          =   615
      Left            =   0
      TabIndex        =   22
      Top             =   600
      Width           =   4935
      Begin VB.CommandButton cmdRESHOW 
         Caption         =   "ReShow"
         Height          =   255
         Left            =   2760
         TabIndex        =   45
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox FRValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2040
         MultiLine       =   -1  'True
         TabIndex        =   42
         Top             =   250
         Width           =   375
      End
      Begin VB.TextBox FRValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   40
         Top             =   250
         Width           =   375
      End
      Begin VB.CommandButton cmdNEXT 
         Caption         =   "->"
         Height          =   255
         Left            =   4440
         TabIndex        =   7
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdPREV 
         Caption         =   "<-"
         Height          =   255
         Left            =   3960
         TabIndex        =   6
         Top             =   240
         Width           =   375
      End
      Begin VB.Label FRLabel 
         Caption         =   "of"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   41
         Top             =   240
         Width           =   255
      End
      Begin VB.Label FRLabel 
         Caption         =   "Family Record #"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame cmdFRAME 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   4200
      Width           =   4935
      Begin VB.CommandButton cmdBUTTON 
         Caption         =   "Notes"
         Height          =   255
         Index           =   5
         Left            =   3840
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdBUTTON 
         Caption         =   "Address"
         Height          =   255
         Index           =   4
         Left            =   2880
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdBUTTON 
         Caption         =   "Children"
         Height          =   255
         Index           =   3
         Left            =   1920
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdBUTTON 
         Caption         =   "Select"
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdBUTTON 
         Caption         =   "Update"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmFHSWSpse"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
DefInt A-Z
Dim OMRNO, OSPOUSE As OthrFamilyREC
Dim SPID1, SPID2, OSPID, SP1OLDCH, SP2OLDCH
Dim MRIX, MRHIX, MRMAX, MRCHN(), MRNO, NMRNO
Dim DD1Y, DD1MD, DD2Y, DD2MD, DDY, DDMD, RGL
Dim LastDataFLD As Integer, HOLDITEM As String
Dim xSS, xSL, xFT, ReShowDATA, NewSpouse
Dim HRValue() As String
Sub cmdSetup(cmdOPT)
For N = 1 To 5
    CMDButton(N).Visible = True
    CMDButton(N).Font.Bold = False
    CMDButton(N).Enabled = True
    CMDButton(N).DEFAULT = False
    CMDButton(N).Cancel = False
Next N
FRIDFrame.Enabled = True
FRMIFrame.Enabled = False
FRMNUframe.Enabled = False
Select Case cmdOPT
   Case 1: ' Standard SPOUSE Commands
         FRMNUframe.Enabled = True
         cmdFRMNU(5).Cancel = True
         CMDButton(1).Caption = "Update"
         CMDButton(2).Caption = "Select"
         CMDButton(3).Caption = "Children"
         CMDButton(3).Font.Bold = ((SP1OLDCH <> 0) And (SP2OLDCH <> 0))
         CMDButton(3).Enabled = CMDButton(3).Font.Bold
         CMDButton(4).Caption = "Res"
         CMDButton(4).Font.Bold = (FF3SPOUSE.LOC <> 0)
         CMDButton(5).Caption = "Notes"
         CMDButton(5).Font.Bold = (FF3SPOUSE.COM <> 0)
   Case 2: ' Change Choices
         CMDButton(1).Caption = "Change"
         CMDButton(2).Caption = "Add"
         CMDButton(3).Caption = "Delete"
         CMDButton(4).Caption = "Cancel"
         CMDButton(4).Cancel = True
         CMDButton(5).Visible = False
   Case 3: ' Add/Update Choices
         FRIDFrame.Enabled = False
         FRMIFrame.Enabled = True
         CMDButton(1).Caption = "Save"
         CMDButton(2).Caption = "Cancel"
         CMDButton(3).Caption = "Restore"
         CMDButton(2).Cancel = True
         CMDButton(4).Visible = False
         CMDButton(5).Visible = False
         FRValue(9).SetFocus
End Select
End Sub
Sub DeleteFamilyREC()
908 ' Delete Marriage Record
   If Not ENV.FFOPEN Then Call FAMOpen("WRITE")
   Call PutMSG("UPDTM007")
   If CA$ <> SNGLKEY$(1) Then Exit Sub
   If OSPID = 0 Then
      MADR = FF3SPOUSE.LOC
      If MADR <> 0 Then
         Call PutMSG("UPDTM008")
         If CA$ <> SNGLKEY$(1) Then Exit Sub
      End If
      mcom = FF3SPOUSE.COM
      If mcom <> 0 Then
         Call PutMSG("UPDTM009")
         If CA$ <> SNGLKEY$(1) Then Exit Sub
      End If
      If FF3SPOUSE.MLOC <> 0 Then
         Call FF3GetRec(FF3SPOUSE.MLOC)
         If FF3PFX.RTYPE = Chr$(12) Then
            If FF3PLACE2.PNEXT <> 0 Then Call FF3DelRec(FF3PLACE2.PNEXT)
         End If
         Call FF3DelRec(FF3SPOUSE.MLOC)
      End If
      X = MADR
      While X <> 0
         Call FF2GetRec(X)
         If Not OKAY Then
            X = 0
            Else
            Y = FF2PFX.COM
            While Y <> 0
               Call FF3GetRec(Y)
               If Not OKAY Then
                  Y = 0
                  Else
                  NY = FF3COMMENT.FPTR
                  Call FF3DelRec(Y)
                  If Not OKAY Then NY = 0
                  Y = NY
               End If
            Wend
            NX = FF2PFX.NXT
            Call FF2DelRec(X)
            If Not OKAY Then NX = 0
            X = NX
          End If
      Wend
      X = mcom
      While X <> 0
         Call FF3GetRec(X)
         If Not OKAY Then
            X = 0
            Else
            NX = FF3COMMENT.FPTR
            Call FF3DelRec(X)
            If Not OKAY Then NX = 0
            X = NX
         End If
      Wend
   End If
   Call UpdtFamPTRS("REM", SPID1, MRNO)
   If OSPID = 0 Then
      Call FF3DelRec(MRNO)
      Else
      Call FF3GetRec(MRNO)
      If FF3SPOUSE.SP1ID = SPID1 Then
         FF3SPOUSE.SP1ID = 0: FF3SPOUSE.SP1NXT = 0
         Else
         If FF3SPOUSE.SP2ID = SPID1 Then
            FF3SPOUSE.SP2ID = 0: FF3SPOUSE.SP2NXT = 0
         End If
      End If
      Call FF3PutRec
   End If
   If OKAY Then
      X = MRIX
      While X < MRHIX: MRCHN(X) = MRCHN(X + 1): X = X + 1: Wend
      MRHIX = MRHIX - 1
      If MRIX > MRHIX Then MRIX = MRHIX
      MRNO = MRCHN(MRIX)
   End If
   UpdateCHG
   FRDisplay (MRIX)
End Sub

Sub FRDisplay(MNUM)
   If MRHIX = 0 Then
      frmFHSWSpse.Height = 1650
      FRIDFrame.Visible = False
      FRLabel(10).Caption = "No Family Record Present"
      FRMSGFrame.Visible = True
      FRMNUframe.Enabled = True
      Exit Sub
   End If
   FRMSGFrame.Visible = False
   FRIDFrame.Visible = True
   frmFHSWSpse.Height = 5250
   MRIX = MNUM: If MRIX = 0 Then MRIX = 1
   MRNO = MRCHN(MRIX)
   FRValue(0) = MRHIX - MRIX + 1: FRValue(1) = MRHIX
   cmdPREV.Font.Bold = (MRIX < MRHIX)
   cmdNEXT.Font.Bold = (MRIX > 1)
   If Not ENV.FFOPEN Then Call FAMOpen("WRITE")
   Call FF3GetRec(MRNO): If Not OKAY Then Exit Sub
   If FF3SPOUSE.SP1ID = SPID1 Then
      SPID2 = FF3SPOUSE.SP2ID: NM = FF3SPOUSE.SP1NXT
      Else
      SPID2 = FF3SPOUSE.SP1ID: NM = FF3SPOUSE.SP2NXT
   End If
   MRCHN(MRIX + 1) = NM
   LSet MPLACE = "": LSet DVPLACE = ""
   If FF3SPOUSE.MLOC <> 0 Then
      Call FF3GetRec(FF3SPOUSE.MLOC)
      If FF3PFX.RTYPE = Chr$(12) Then
         If FF3PLACE2.PNEXT <> 0 Then Call FF3GetRec(FF3PLACE2.PNEXT)
      End If
   End If
   LSet OMPLACE = MPLACE: LSet ODVPLACE = DVPLACE
   OSPID = SPID2
   OSPOUSE = FF3SPOUSE
   D1MD = FF3SPOUSE.D1MD: D1Y = FF3SPOUSE.D1Y
   ShowSpouseDATA
   ShowMarriageDATA
   HRValue(2) = RTrim$(FRValue(X).TEXT)
   For X = 9 To 14
      HRValue(X) = RTrim$(FRValue(X).TEXT)
   Next X
   Call cmdSetup(1)
End Sub



Sub ReadFamilyDATA()
   If Not ENV.FFOPEN Then Call FAMOpen("WRITE")
   X& = FRValue(2): ChkID (X&)
   If OKAY Then
      SPID2 = X&
      If SPID2 = SPID1 Then
         Call PutMSG("UPDTM003"): Exit Sub
      End If
      If SPID2 > 0 Then
         Call FF1GetRec(SPID2)
         If Not OKAY Then Exit Sub
      End If
   End If
   FF3SPOUSE.STATUS1 = FRValue(9)
   FF3SPOUSE.D1MD = DATEMD(FRValue(10), DFM)
   FF3SPOUSE.D1Y = DATEY(FRValue(10), DFM)
   MPLACE = RTrim$(FRValue(11))
   FF3SPOUSE.STATUS2 = FRValue(12)
   FF3SPOUSE.D2MD = DATEMD(FRValue(13), DFM)
   FF3SPOUSE.D2Y = DATEY(FRValue(13), DFM)
   DVPLACE = RTrim$(FRValue(14))

End Sub

Sub ShowMarriageDATA()
   If Not ENV.FFOPEN Then Call FAMOpen("WRITE")
   FRValue(9) = FF3SPOUSE.STATUS1
   FRValue(10) = XSCRNDate$(FF3SPOUSE.D1MD, FF3SPOUSE.D1Y)
   FRValue(11) = MPLACE
   FRValue(12) = FF3SPOUSE.STATUS2
   FRValue(13) = XSCRNDate$(FF3SPOUSE.D2MD, FF3SPOUSE.D2Y)
   FRValue(14) = DVPLACE
   D2MD = FF3SPOUSE.D2MD: D2Y = FF3SPOUSE.D2Y
   If D2Y = 0 Then
      D2Y = DDY: If D2Y = 0 Then D2MD = 0
   End If
   FRValue(15) = DATEDif$(FF3SPOUSE.D1MD, FF3SPOUSE.D1Y, D2MD, D2Y)
End Sub

Sub ShowSpouseDATA()
   If Not ENV.FFOPEN Then Call FAMOpen("WRITE")
   If SPID2 > 0 Then
      Call FF1GetRec(SPID2): SP2OLDCH = FF1Rec.OLDCH: DDY = FF1Rec.DY
      If DDY > 0 Then If DDY < 9999 Then DDMD = FF1Rec.DMD Else DDY = 0
      Else
      DDY = 0: SP2OLDCH = 0
   End If
   If (0 < DD1Y) And (DD1Y < DDY) Then DDY = DD1Y: DDMD = DD1MD
   FRValue(2) = Str$(SPID2)
   If SPID2 > 0 Then
      FRValue(4) = FF1Rec.GIVEN
      FRValue(5) = FF1Rec.SURNM
      FRValue(6) = DATEDif$(FF1Rec.BMD, FF1Rec.BY, FF1Rec.DMD, FF1Rec.DY)
      FRValue(7) = XSCRNDate$(FF1Rec.BMD, FF1Rec.BY) + FDSTAT$("BD")
      FRValue(8) = XSCRNDate$(FF1Rec.DMD, FF1Rec.DY) + FDSTAT$("DD")
      Else
      For X = 4 To 8: FRValue(X) = "": Next X
   End If
End Sub
Sub RestoreField(FldNum)
If RTrim$(FRValue(FldNum).TEXT) <> HRValue(FldNum) Then
   FRValue(FldNum).TEXT = HRValue(FldNum)
   OKAY = -1
   Else
   OKAY = 0
End If
End Sub


Sub UpdateCHG()
   If Not ENV.FFOPEN Then Call FAMOpen("WRITE")
   If XCHANGED Then
      If (SPID1 <> 0) Then
         Call FF1GetRec(SPID1): XCHANGED = -1: Call FF1PutREC
      End If
      If (OSPID <> 0) And (OSPID <> SPID2) And (OSPID <> SPID1) Then
         Call FF1GetRec(OSPID): XCHANGED = -1: Call FF1PutREC
      End If
      If (SPID2 <> 0) Then
         Call FF1GetRec(SPID2): XCHANGED = -1: Call FF1PutREC
      End If
   End If
   XCHANGED = 0
End Sub

Sub UpdateFamilyREC()
A906: ' Save Changes
   ReadFamilyDATA
   If Not OKAY Then Exit Sub
   FF3PFX.RTYPE = Chr$(3): FF3PFX.RNO = NMRNO
   Call FF3PutRec
   XLOC = FF3SPOUSE.MLOC
   Call UpdtPLACE(XLOC, 3, NMRNO)
   Call FF3GetRec(NMRNO)
   FF3SPOUSE.MLOC = XLOC
   Call FF3PutRec
   If (NMRNO <> MRNO) Or (SPID2 <> OSPID) Or (FF3SPOUSE.D1MD <> D1MD) Or (FF3SPOUSE.D1Y <> D1Y) Then
      If (NMRNO = MRNO) Then
         Call UpdtFamPTRS("REM", SPID1, NMRNO)
         If OSPID <> 0 Then Call UpdtFamPTRS("REM", OSPID, NMRNO)
      End If
      If SPID2 <> OSPID Then
         Call FF3GetRec(NMRNO)
         If SPID1 = FF3SPOUSE.SP1ID Then
            FF3SPOUSE.SP2ID = SPID2
            Else
            If SPID1 = FF3SPOUSE.SP2ID Then FF3SPOUSE.SP1ID = SPID2
         End If
         Call FF3PutRec
      End If
      If SPID2 <> 0 Then
         Call UpdtFamPTRS("ADD", SPID2, NMRNO)
      End If
      Call UpdtFamPTRS("ADD", SPID1, NMRNO)
      Call FF1GetRec(SPID1)
      MRNO = NMRNO
   End If
   UpdateCHG
   Call FF1GetRec(SPID1)
   Call ViewMARRIAGES(SPID1, NMRNO, RID)
   NewSpouse = 0
End Sub
Sub ViewMARRIAGES(x1, Y1, xGL)
   If Not ENV.FFOPEN Then
      Call FAMOpen("WRITE")
      If Not OKAY Then Unload frmFHSWSpse: Exit Sub
   End If
   RGL = xGL
   frmFHSWUPDT.Enabled = False
   SPID1 = x1: SPID2 = 0
   Call FF1GetRec(SPID1): SP1OLDCH = FF1Rec.OLDCH
   DD1Y = FF1Rec.DY
   If DD1Y > 0 Then If DD1Y < 9999 Then DD1MD = FF1Rec.DMD Else DD1Y = 0
A900: ' Build List of Marriages
   MRIX = 0: MRHIX = 0: MRMAX = 10
   ReDim MRCHN(MRMAX)
   X = FF1Rec.SPOUSE
   While (X <> 0)
      MRHIX = MRHIX + 1
      If X = Y1 Then MRIX = MRHIX
      If Not (MRIX < MRMAX) Then
         MRMAX = MRMAX + 5: ReDim Preserve MRCHN(MRMAX)
      End If
      MRCHN(MRHIX) = X
      Call FF3GetRec(X)
      If Not OKAY Then
         X = 0
         Else
         If SPID1 = FF3SPOUSE.SP1ID Then
            X = FF3SPOUSE.SP1NXT
            Else
            If SPID1 = FF3SPOUSE.SP2ID Then
               X = FF3SPOUSE.SP2NXT
               Else
               X = 0
            End If
         End If
      End If
   Wend
   FRDisplay (MRIX)
   frmFHSWSpse.Show
End Sub

Private Sub CMDButton_Click(Index As Integer)
X = Index
CmdOption:
Select Case CMDButton(X).Caption
   Case "Add": Call NewFamilyRec
   Case "Cancel": ' Cancel Update
        If Index = 2 Then
           If Not ENV.FFOPEN Then Call FAMOpen("WRITE")
           If MRNO <> NMRNO Then
              Call FF3DelRec(NMRNO)
              NMRNO = 0
           End If
           FRDisplay (MRIX)
        End If
        cmdSetup (1)
   Case "Change": cmdSetup (3): NMRNO = MRNO
   Case "Children":
        frmFHSWSpse.Enabled = False
        Call frmFHSWChld.ShowCHILDREN(SPID1, SPID2, RGL)
   Case "Delete":
        Call DeleteFamilyREC
        cmdSetup (1)
   Case "Notes":
        COMBGN = FF3SPOUSE.COM: SRTYPE = 3: SRNO = MRNO
        Call frmFHSWNote.ViewCOMMENTS(FF3SPOUSE.COM, SRTYPE, SRNO)
   Case "Res":
        ADRBGN = FF3SPOUSE.LOC: SRTYPE = 3: SRNO = MRNO
        Call frmFHSWAdrs.ViewADDRESS(FF3SPOUSE.LOC, SRTYPE, SRNO)
        frmFHSWSpse.Enabled = False
   Case "Restore": ' Restore Last Value for Field
        If (RTrim$(FRValue(LastDataFLD).TEXT) <> HOLDITEM) Then
           FRValue(LastDataFLD).TEXT = HOLDITEM
           Else
           RestoreField (LastDataFLD)
           'If Not OKAY Then X = 2: GoTo CmdOption
        End If
        FRValue(LastDataFLD).SetFocus
   Case "Return": cmdSetup (1)
   Case "Save": ' Save Changes to Name Record
        Call UpdateFamilyREC
        If OKAY Then Call FRDisplay(MRIX)
   Case "Select": ' Display Spouse Name Record
        If SPID2 > 0 Then
           Call frmFHSWUPDT.NRDisplay(SPID2, RGL)
           Unload frmFHSWSpse
        End If
   Case "Update":
        'OBPLACE = BPLACE: ODPLACE = DPLACE
        'NameREC = FF1Rec
        cmdSetup (2)
End Select
End Sub

Private Sub cmdFRMNU_Click(Index As Integer)
Select Case Index
  Case 1: ' Add New Record
           Call NewFamilyRec
           If Not OKAY Then cmdSetup (1)
  Case 2: ' Search Family File
           frmFHSWSlct.Show
           CMDFrame.Enabled = False
  Case 3: ' Relationship Calculator
           Call frmFHSWRCLC.RelCALC(SPID2)
           CMDFrame.Enabled = False
  Case 4: ' Date Calculator
           frmFHSWDCLC.Show
           ReShowDATA = -1
           Exit Sub
  Case 5: Unload frmFHSWSpse: Exit Sub
End Select
FRMNUframe.Enabled = False
End Sub
Sub NewFamilyRec()
   If Not ENV.FFOPEN Then Call FAMOpen("WRITE")
907 ' Add Marriage Record
   FF3PFX.RTYPE = Chr$(3)
   Call FF3GetRec(0)
   If Not OKAY Then Exit Sub
   OMRNO = MRNO
   NMRNO = FF3PFX.RNO
   FF3SPOUSE.SRTYPE = Chr$(1)
   FF3SPOUSE.SP1ID = SPID1
   Call FF3PutRec
   If Not OKAY Then Exit Sub
   SPID2 = 0: OSPID = 0
   LSet OMPLACE = "": LSet ODVPLACE = ""
   LSet MPLACE = "": LSet DVPLACE = ""
   NewSpouse = True
   ShowSpouseDATA
   ShowMarriageDATA
   frmFHSWSpse.Height = 5250
   FRIDFrame.Visible = False
   FRLabel(10).Caption = "New Family Record"
   FRMSGFrame.Visible = True
   cmdSetup (3)
End Sub
Private Sub cmdNEXT_Click()
If MRIX > 1 Then FRDisplay (MRIX - 1)
End Sub

Private Sub cmdPREV_Click()
If MRIX < MRHIX Then FRDisplay (MRIX + 1)
End Sub


Private Sub cmdRESHOW_Click()
If MRIX <> 0 Then Call FRDisplay(MRIX)
End Sub

Private Sub cmdSHOW_Click()
SPID2 = Val(FRValue(2).TEXT)
ShowSpouseDATA
End Sub

Private Sub cmdSPID_Click()
X = Val(FRValue(2).TEXT)
Call frmFHSWList.SelectName(X)
If a <> 27 Then
   SPID2 = X: ShowSpouseDATA
End If
End Sub

Private Sub Form_Activate()
   frmFHSWUPDT.Enabled = False
   If Not ENV.FFOPEN Then Call FAMOpen("WRITE")
   CMDFrame.Enabled = True
   FRMNUframe.Enabled = True
   If ReShowDATA Then
      ShowSpouseDATA
      ShowMarriageDATA
      ReShowDATA = 0
   End If
   If Not NewSpouse Then Call cmdSetup(1)
End Sub


Private Sub Form_GotFocus()
If ReShowDATA Then
   CMDFrame.Enabled = True
   FRMNUframe.Enabled = True
   ShowSpouseDATA
   ShowMarriageDATA
   cmdSetup (1)
   ReShowDATA = 0
End If
End Sub

Private Sub Form_Load()
Call FAMOpen("WRITE"): If Not OKAY Then Exit Sub
FormLoaded(2) = FormLoaded(2) Or SpouseForm
' Set Default Form Location on Screen
frmHEIGHT = frmFHSWSpse.Height
frmWIDTH = frmFHSWSpse.WIDTH
t1 = GetSetting("FHSWIN", "frmSPSE", "TOP", 0)
If t1 = 0 Then
   t00 = frmFHSWUPDT.TOP: w00 = frmFHSWUPDT.WIDTH
   l00 = frmFHSWUPDT.Left
   t1 = t00
   l1 = l00 + frmWIDTH
   If (l1 + frmWIDTH) > Screen.WIDTH Then
      t1 = frmFHSWUPDT.TOP + 2715
      l1 = l00 + frmWIDTH / 2
   End If
   Else
   t1 = 1440: l1 = 1440
End If
frmFHSWSpse.TOP = GetSetting("FHSWIN", "frmSPSE", "Top", t1)
frmFHSWSpse.Left = GetSetting("FHSWIN", "frmSPSE", "Left", l1)
'Call SetCalendarType(0, 0)
ReDim HRValue(14)
End Sub


Private Sub Form_Unload(Cancel As Integer)
FormLoaded(2) = FormLoaded(2) And (255 - SpouseForm)
SaveSetting "FHSWIN", "frmSPSE", "Top", frmFHSWSpse.TOP
SaveSetting "FHSWIN", "frmSPSE", "Left", frmFHSWSpse.Left
frmFHSWUPDT.Enabled = True
frmFHSWUPDT.SetFocus
End Sub



Private Sub FRValue_GotFocus(Index As Integer)
HOLDITEM = RTrim$(FRValue(Index).TEXT)
Select Case Index
   Case 2: xFT = 2
   Case 10: xFT = 3
   Case 13: xFT = 3
   Case Else: xFT = 1
End Select
FRValue(Index).SelStart = 0
If (xFT > 2) Then
   FRValue(Index).SelLength = 0
   Else
   FRValue(Index).SelLength = Len(RTrim$(FRValue(Index).TEXT))
End If
LastDataFLD = Index
End Sub


Private Sub FRValue_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If xFT > 1 Then
   XVALUE$ = FRValue(Index).TEXT
   xSS = FRValue(Index).SelStart
   xSL = FRValue(Index).SelLength
   Call FldUpdate(KeyCode, Shift, xFT, xSS, xSL, XVALUE$)
   FRValue(Index).TEXT = XVALUE$
   FRValue(Index).SelStart = xSS
   FRValue(Index).SelLength = xSL
End If
End Sub

Private Sub FRValue_KeyPress(Index As Integer, KeyAscii As Integer)
If xFT > 1 Then KeyAscii = 0
End Sub


