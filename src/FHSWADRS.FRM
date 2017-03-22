VERSION 4.00
Begin VB.Form frmFHSWAdrs 
   Caption         =   "FHSWADRS - Address"
   ClientHeight    =   3105
   ClientLeft      =   1920
   ClientTop       =   2700
   ClientWidth     =   6300
   Height          =   3510
   Left            =   1860
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   Top             =   2355
   Width           =   6420
   Begin VB.Frame CMDFrame 
      Height          =   495
      Left            =   0
      TabIndex        =   30
      Top             =   2550
      Width           =   6375
      Begin VB.CommandButton CMDButton 
         Height          =   280
         Index           =   2
         Left            =   2640
         TabIndex        =   12
         Top             =   155
         Width           =   1095
      End
      Begin VB.CommandButton CMDButton 
         Cancel          =   -1  'True
         Caption         =   "Return"
         Default         =   -1  'True
         Height          =   280
         Index           =   3
         Left            =   5160
         TabIndex        =   13
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton CMDButton 
         Caption         =   "Notes"
         Height          =   280
         Index           =   1
         Left            =   1320
         TabIndex        =   11
         Top             =   155
         Width           =   1215
      End
      Begin VB.CommandButton CMDButton 
         Caption         =   "Update"
         Height          =   280
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   155
         Width           =   1095
      End
   End
   Begin VB.Frame DATAFrame 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      Begin VB.Frame NAVFrame 
         Height          =   975
         Left            =   4440
         TabIndex        =   31
         Top             =   120
         Width           =   1695
         Begin VB.CommandButton NAVButton 
            Caption         =   ">>"
            Height          =   315
            Index           =   3
            Left            =   1275
            TabIndex        =   17
            Top             =   240
            Width           =   315
         End
         Begin VB.CommandButton NAVButton 
            Caption         =   "<<"
            Height          =   315
            Index           =   0
            Left            =   100
            TabIndex        =   14
            Top             =   240
            Width           =   350
         End
         Begin VB.CommandButton NAVButton 
            Caption         =   "<-"
            Height          =   315
            Index           =   1
            Left            =   490
            TabIndex        =   15
            Top             =   240
            Width           =   350
         End
         Begin VB.CommandButton NAVButton 
            Caption         =   "->"
            Height          =   315
            Index           =   2
            Left            =   880
            TabIndex        =   16
            Top             =   240
            Width           =   350
         End
         Begin VB.TextBox ARValue 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   9
            Left            =   320
            MultiLine       =   -1  'True
            TabIndex        =   27
            Top             =   645
            Width           =   375
         End
         Begin VB.TextBox ARValue 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   10
            Left            =   1000
            MultiLine       =   -1  'True
            TabIndex        =   29
            Text            =   "FHSWADRS.frx":0000
            Top             =   645
            Width           =   375
         End
         Begin VB.Label ARLabel 
            Alignment       =   2  'Center
            Caption         =   "of"
            Height          =   255
            Index           =   9
            Left            =   650
            TabIndex        =   28
            Top             =   645
            Width           =   375
         End
      End
      Begin VB.TextBox ARValue 
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
         Index           =   8
         Left            =   4080
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   2040
         Width           =   2055
      End
      Begin VB.TextBox ARValue 
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
         Left            =   4680
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox ARValue 
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
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   8
         Text            =   "FHSWADRS.frx":0002
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox ARValue 
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
         Left            =   840
         MultiLine       =   -1  'True
         TabIndex        =   6
         Text            =   "FHSWADRS.frx":0004
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox ARValue 
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
         Left            =   840
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   1320
         Width           =   3495
      End
      Begin VB.TextBox ARValue 
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
         Left            =   840
         MultiLine       =   -1  'True
         TabIndex        =   4
         Text            =   "FHSWADRS.frx":0006
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox ARValue 
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
         Left            =   840
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox ARValue 
         Height          =   315
         Index           =   1
         Left            =   2880
         MultiLine       =   -1  'True
         TabIndex        =   2
         Text            =   "FHSWADRS.frx":0008
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox ARValue 
         Height          =   315
         Index           =   0
         Left            =   840
         MultiLine       =   -1  'True
         TabIndex        =   1
         Text            =   "FHSWADRS.frx":000A
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label ARLabel 
         Caption         =   "Country:"
         Height          =   255
         Index           =   8
         Left            =   3360
         TabIndex        =   26
         Top             =   2085
         Width           =   615
      End
      Begin VB.Label ARLabel 
         Caption         =   "Postal Code:"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   25
         Top             =   2085
         Width           =   1095
      End
      Begin VB.Label ARLabel 
         Caption         =   "State/Province:"
         Height          =   255
         Index           =   6
         Left            =   3360
         TabIndex        =   24
         Top             =   1725
         Width           =   1215
      End
      Begin VB.Label ARLabel 
         Caption         =   "City:"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   23
         Top             =   1725
         Width           =   495
      End
      Begin VB.Label ARLabel 
         Caption         =   "Line2:"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   22
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label ARLabel 
         Caption         =   "Line1:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   21
         Top             =   990
         Width           =   495
      End
      Begin VB.Label ARLabel 
         Caption         =   "Phone:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   20
         Top             =   630
         Width           =   615
      End
      Begin VB.Label ARLabel 
         Caption         =   "To:"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   19
         Top             =   275
         Width           =   495
      End
      Begin VB.Label ARLabel 
         Caption         =   "From:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   270
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmFHSWAdrs"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
DefInt A-Z
Dim AFMT, ASRTYPE, ASRNO, ADRRNO, ADRBGN
Dim ADRHIX, ADRMIX, ADRIX, AD1MD, AD1Y, NEWADR
Dim ADRCHN()
Dim LastFLD, ZERODATE$
Dim HOLDITEM As Variant, xFT, HRValue() As String
Dim frmHEIGHT, frmWIDTH
Sub BldADRCHN(X)
    ADRRNO = X
    ReDim ADRCHN(ADRMIX)
    If ADRRNO = 0 Then Exit Sub
    Call FF2GetRec(ADRRNO)
    AD1Y = FF2PFX.D1Y: AD1MD = FF2PFX.D1MD
    If ADRBGN = 0 Then ADRBGN = ADRRNO
    ADRHIX = 0: Z = ADRBGN: ADRIX = 0
    While Z <> 0  ' Build chain, inserting adrrno where it belongs
       Call FF2GetRec(Z)
       If OKAY Then
          ADRHIX = ADRHIX + 1
          If Not (ADRHIX < ADRMIX) Then
             ADRMIX = ADRMIX + 10
             ReDim Preserve ADRCHN(ADRMIX)
          End If
          If (ADRIX > 0) Or ((FF2PFX.D1Y > AD1Y) Or ((FF2PFX.D1Y = AD1Y) And (FF2PFX.D1MD > AD1MD))) Then
             ADRCHN(ADRHIX) = Z
             Else
             ADRCHN(ADRHIX) = ADRRNO: ADRIX = ADRHIX
             If Z <> ADRRNO Then
                ADRHIX = ADRHIX + 1
                ADRCHN(ADRHIX) = Z
             End If
          End If
          Z = FF2PFX.NXT
          Else
          Z = 0
       End If
    Wend
    If ADRIX = 0 Then
       ADRHIX = ADRHIX + 1: ADRIX = ADRHIX
       ADRCHN(ADRHIX) = ADRRNO
    End If
    If ADRIX > 1 Then
       Call FF2GetRec(ADRCHN(ADRIX - 1))
       If OKAY Then
          FF2PFX.NXT = ADRRNO
          Call FF2PutRec
       End If
    End If
    Call FF2GetRec(ADRRNO)
    If OKAY Then
       If FF2PFX.NXT <> ADRCHN(ADRIX + 1) Then
          FF2PFX.NXT = ADRCHN(ADRIX + 1)
          Call FF2PutRec
       End If
    End If
    ARValue(9) = ADRHIX - ADRIX + 1
    ARValue(10) = ADRHIX
    ADRBGN = ADRCHN(1)
End Sub
Sub ClearSCRN()
   ARValue(0) = ZERODATE$
   ARValue(1) = ZERODATE$
   ARValue(2) = "(000) 000-0000"
   For X = 3 To 8: ARValue(X) = "": Next X
   frmFHSWAdrs.Show
   NEWADR = -1
End Sub

Sub cmdSetup(cmdOPT)
For N = 1 To 3
    CMDButton(N).Visible = True
    CMDButton(N).Font.Bold = False
    CMDButton(N).Enabled = True
    CMDButton(N).DEFAULT = False
    CMDButton(N).Cancel = False
Next N
For N = 0 To 8
    ARValue(N).Enabled = False
Next N
NAVFrame.Enabled = False
Select Case cmdOPT
   Case 1: ' Standard ADDRESS Commands
         NAVFrame.Enabled = True
         NAVButton(1).Enabled = (ADRIX < ADRHIX)
         NAVButton(2).Enabled = (ADRIX > 1)
         CMDButton(0).Caption = "Update"
         CMDButton(1).Caption = "Notes"
         CMDButton(1).Font.Bold = (FF2PFX.COM <> 0)
         CMDButton(2).Visible = False
         CMDButton(3).Caption = "Return"
         CMDButton(3).Cancel = True
         CMDButton(3).DEFAULT = True
         NEWADR = 0
   Case 2: ' Change Choices
         CMDButton(0).Caption = "Change"
         CMDButton(1).Caption = "Add"
         CMDButton(2).Visible = True
         CMDButton(2).Caption = "Delete"
         CMDButton(3).Caption = "Cancel"
         CMDButton(3).Cancel = True
         CMDButton(3).DEFAULT = False
   Case 3: ' Add/Update Choices
         For N = 0 To 8
             ARValue(N).Enabled = True
         Next N
         CMDButton(0).Caption = "Save"
         CMDButton(1).Caption = "Cancel"
         CMDButton(1).Cancel = True
         CMDButton(2).Caption = "Restore"
         CMDButton(3).Visible = False
         ARValue(0).SetFocus
End Select
End Sub

Sub DeleteADRRec()
3025 ' Delete Current ADDRESS Record
   Call PutMSG("UPDTM007")
   If CA$ <> SNGLKEY$(1) Then Exit Sub
   acom = FF2PFX.COM
   If acom <> 0 Then
      Call PutMSG("UPDTM009")
      If CA$ <> SNGLKEY$(1) Then Exit Sub
   End If
   While acom <> 0
       Call FF3GetRec(acom)
       If Not OKAY Then
          acom = 0
          Else
          hptr = FF3COMMENT.FPTR
          Call FF3DelRec(acom)
          If Not OKAY Then
             acom = 0
             Else
             acom = hptr
          End If
       End If
   Wend
   Call FF2DelRec(ADRRNO)
   OBGN = ADRBGN: Call RemADRCHN
   If OBGN <> ADRBGN Then Call UpdateSRNO(ADRBGN)

End Sub
Sub GetADRInfo()
3004 ' Move ADDRESS Info from Screen
    xX$ = Space$(40)
    If AFMT = 0 Then
       For X = 0 To 7
          LSet xX$ = ARValue(X).TEXT
          Select Case X
             Case 0: FF2PFX.D1MD = DATEMD(xX$, DFM)
                     FF2PFX.D1Y = DATEY(xX$, DFM)
             Case 1: FF2PFX.D2MD = DATEMD(xX$, DFM)
                     FF2PFX.D2Y = DATEY(xX$, DFM)
             Case 2: FF2SREC.FON1 = Val(MID$(xX$, 2, 3))
                     FF2SREC.FON2 = Val(MID$(xX$, 7, 3))
                     FF2SREC.FON3 = Val(MID$(xX$, 11, 4))
             Case 3: LSet FF2SREC.LINE1 = xX$
             Case 4: LSet FF2SREC.LINE2 = xX$
             Case 5: LSet FF2SREC.CITY = xX$
             Case 6: LSet FF2SREC.STATE = xX$
             Case 7: LSet FF2SREC.ZIP = xX$
          End Select
          ARValue(X) = Left$(xX$, ARValue(X).WIDTH)
      Next X
      Else
      For X = 0 To 8
          LSet xX$ = ARValue(X).TEXT
          Select Case X
             Case 0: FF2PFX.D1MD = DATEMD(xX$, DFM)
                     FF2PFX.D1Y = DATEY(xX$, DFM)
             Case 1: FF2PFX.D2MD = DATEMD(xX$, DFM)
                     FF2PFX.D2Y = DATEY(xX$, DFM)
             Case 2: LSet FF2LREC.PHONE = xX$
             Case 3: LSet FF2LREC.LINE1 = xX$
             Case 4: LSet FF2LREC.LINE2 = xX$
             Case 5: LSet FF2LREC.CITY = xX$
             Case 6: LSet FF2LREC.STATE = xX$
             Case 7: LSet FF2LREC.ZIP = xX$
             Case 8: LSet FF2LREC.COUNTRY = xX$
          End Select
          ARValue(X) = Left$(xX$, ARValue(X).WIDTH)
      Next X
    End If
End Sub

Private Sub RemADRCHN()
3017 ' Remove ADRRNO From Address Chain
   If ADRIX > 1 Then
      Call FF2GetRec(ADRCHN(ADRIX - 1))
      FF2PFX.NXT = ADRCHN(ADRIX + 1)
      Call FF2PutRec
      Call FF2GetRec(ADRRNO)
   End If
   For X = ADRIX To ADRHIX - 1
      ADRCHN(X) = ADRCHN(X + 1)
   Next X
   ADRCHN(ADRHIX) = 0
   ADRHIX = ADRHIX - 1
   If ADRIX > ADRHIX Then ADRIX = ADRHIX
   ADRRNO = ADRCHN(ADRIX)
   ADRBGN = ADRCHN(1)
   ARValue(9) = ADRHIX - ADRIX + 1
   ARValue(10) = ADRHIX
End Sub

Sub ShowADRFld(X)
    xX$ = Space$(40)
    If AFMT = 0 Then
       Select Case X
          Case 0: LSet xX$ = XSCRNDate$(FF2PFX.D1MD, FF2PFX.D1Y)
          Case 1: LSet xX$ = XSCRNDate$(FF2PFX.D2MD, FF2PFX.D2Y)
          Case 2: LSet xX$ = "(   )    -"
                  Mid$(xX$, 2, 3) = ZFILL$(FF2SREC.FON1, 3)
                  Mid$(xX$, 7, 3) = ZFILL$(FF2SREC.FON2, 3)
                  Mid$(xX$, 11, 4) = ZFILL$(FF2SREC.FON3, 4)
          Case 3: LSet xX$ = FF2SREC.LINE1
          Case 4: LSet xX$ = FF2SREC.LINE2
          Case 5: LSet xX$ = FF2SREC.CITY
          Case 6: LSet xX$ = FF2SREC.STATE
          Case 7: LSet xX$ = FF2SREC.ZIP
       End Select
       Else
       Select Case X
          Case 0: LSet xX$ = XSCRNDate$(FF2PFX.D1MD, FF2PFX.D1Y)
          Case 1: LSet xX$ = XSCRNDate$(FF2PFX.D2MD, FF2PFX.D2Y)
          Case 2: LSet xX$ = FF2LREC.PHONE
          Case 3: LSet xX$ = FF2LREC.LINE1
          Case 4: LSet xX$ = FF2LREC.LINE2
          Case 5: LSet xX$ = FF2LREC.CITY
          Case 6: LSet xX$ = FF2LREC.STATE
          Case 7: LSet xX$ = FF2LREC.ZIP
          Case 8: LSet xX$ = FF2LREC.COUNTRY
       End Select
    End If
    ARValue(X).TEXT = Left$(xX$, ARValue(X).WIDTH)
End Sub

Sub ShowADRInfo()
3002 ' Move Address INFO to Screen
    If AFMT = 0 Then Z = 7 Else Z = 8
    Call FF2GetRec(ADRRNO)
    For X = 0 To Z
        Call ShowADRFld(X)
    Next X
    ARValue(9) = ADRHIX - ADRIX + 1
    frmFHSWAdrs.Show

End Sub
Private Sub UpdateSRNO(NBGN)
Select Case ASRTYPE
   Case 1: Call FF1GetRec(ASRNO)
           FF1Rec.ADRS = NBGN
           Call FF1PutREC
   Case 3: Call FF3GetRec(ASRNO)
           FF3SPOUSE.LOC = NBGN
           Call FF3PutRec
   Case Else: Call FF3GetRec(ASRNO)
           FF3PNTRS.LOC = NBGN
           Call FF3PutRec
End Select
End Sub

Sub ViewADDRESS(ADRNO, SRTYPE, SRNO)
3000 ' Process ADDRESS Information
    If Not ENV.FFOPEN Then Call FAMOpen("WRITE")
    ASRTYPE = Val(SRTYPE): ASRNO = SRNO
    AFMT = FF2Hdr.REVLVL And 255
    ARValue(2).MaxLength = 14
    ARValue(3).MaxLength = 30
    ARValue(4).MaxLength = 30
    If AFMT = 0 Then
       ARValue(5).MaxLength = 15
       ARValue(6).MaxLength = 4
       ARValue(7).MaxLength = 5
       Else
       ARValue(5).MaxLength = 20
       ARValue(6).MaxLength = 12
       ARValue(7).MaxLength = 12
    End If
    ARValue(8).MaxLength = 16
    ARLabel(8).Visible = (AFMT > 0)
    ARValue(8).Visible = (AFMT > 0)
    ZERODATE$ = XSCRNDate$(0, 0)
    ARValue(0).TEXT = ZERODATE$
    ARValue(1).TEXT = ZERODATE$
    ADRBGN = 0
    If ADRNO = 0 Then
       Call ClearSCRN
       Call cmdSetup(3)
       Exit Sub
    End If

3005 ' Process Existing ADDRESS Record
    Call BldADRCHN(ADRNO)
    Call ShowADRInfo
    Call cmdSetup(1)
End Sub
Sub SaveADRInfo()
3016 ' Write Address Record to FILE
   OBGN = ADRBGN
   If NEWADR Then
      Call FF2GetRec(0)
      FF2PFX.SRTYPE = Chr$(ASRTYPE)
      FF2PFX.SRNO = ASRNO
      Else
      Call RemADRCHN
   End If
   Call GetADRInfo
   Call FF2PutRec
   Call BldADRCHN(FF2PFX.RNO)
   If OBGN <> ADRBGN Then Call UpdateSRNO(ADRBGN)
End Sub

Private Sub ARValue_GotFocus(Index As Integer)
HOLDITEM = RTrim$(ARValue(Index).TEXT)
ARValue(Index).SelStart = 0
ARValue(Index).TEXT = RTrim$(ARValue(Index).TEXT)
Select Case Index
   Case 0, 1: xFT = 3
   Case 2: xFT = 6: ARValue(Index).SelStart = 1
   Case Else: xFT = 1
End Select
If (xFT > 2) Then
   ARValue(Index).SelLength = 0
   Else
   ARValue(Index).SelLength = Len(ARValue(Index).TEXT)
End If
LastFLD = Index
End Sub


Private Sub ARValue_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If xFT > 1 Then
   XVALUE$ = ARValue(Index).TEXT
   xSS = ARValue(Index).SelStart
   xSL = ARValue(Index).SelLength
   Call FldUpdate(KeyCode, Shift, xFT, xSS, xSL, XVALUE$)
   ARValue(Index).TEXT = XVALUE$
   ARValue(Index).SelStart = xSS
   ARValue(Index).SelLength = xSL
End If
End Sub


Private Sub ARValue_KeyPress(Index As Integer, KeyAscii As Integer)
If xFT > 1 Then KeyAscii = 0
End Sub


Private Sub CMDButton_Click(Index As Integer)
Select Case CMDButton(Index).Caption
  Case "Add"
       Call ClearSCRN
       Call cmdSetup(3)
  Case "Cancel":
       If ADRIX > 0 Then
          Call ShowADRInfo
          Call cmdSetup(1)
          Else
          Unload frmFHSWAdrs
       End If
  Case "Change": Call cmdSetup(3)
  Case "Delete"
       Call DeleteADRRec
       If ADRRNO <> 0 Then
          Call ShowADRInfo
          Call cmdSetup(1)
          Else
          Call ClearSCRN
          Call cmdSetup(3)
       End If
  Case "Update": Call cmdSetup(2)
  Case "Notes":
       X = FF2PFX.COM: Y = 2: Z = FF2PFX.RNO
       Call frmFHSWNote.ViewCOMMENTS(X, Y, Z)
  Case "Restore":
       If (RTrim$(ARValue(LastFLD).TEXT) <> HOLDITEM) Then
           ARValue(LastFLD).TEXT = HOLDITEM
           Else
           Call ShowADRFld(LastFLD)
        End If
        ARValue(LastFLD).SetFocus
  Case "Return": Unload frmFHSWAdrs
  Case "Save"
       Call SaveADRInfo
       Call ShowADRInfo
       Call cmdSetup(1)
End Select
End Sub







Private Sub Form_Activate()
If Not ENV.FFOPEN Then Call FAMOpen("WRITE")
If ADRIX <> 0 Then Call cmdSetup(1)
End Sub

Private Sub Form_Load()
frmHEIGHT = frmFHSWAdrs.Height
frmWIDTH = frmFHSWAdrs.WIDTH
t0 = frmFHSWUPDT.TOP: l0 = frmFHSWUPDT.Left
t1 = GetSetting("FHSWIN", "frmADRS", "Top", 0)
If t1 = 0 Then
   t1 = t0 + 3585
   If (t1 + frmHEIGHT) > Screen.Height Then t1 = frmFHSWUPDT.TOP + 450
   l1 = l0 - (frmWIDTH - frmFHSWUPDT.WIDTH) / 2
   If l1 < 0 Then l1 = frmFHSWUPDT.Left
End If
FormLoaded(3) = FormLoaded(3) Or AddressForm
frmFHSWAdrs.TOP = GetSetting("FHSWIN", "frmADRS", "Top", t1)
frmFHSWAdrs.Left = GetSetting("FHSWIN", "frmADRS", "Left", l1)
ADRIX = 0: ADRMIX = 10: ADRHIX = 0
ReDim ADRCHN(ADRMIX)
End Sub


Private Sub Form_Resize()
frmFHSWAdrs.Height = frmHEIGHT
frmFHSWAdrs.WIDTH = frmWIDTH
End Sub

Private Sub Form_Unload(Cancel As Integer)
FormLoaded(3) = FormLoaded(3) And (255 - AddressForm)
SaveSetting "FHSWIN", "frmADRS", "Top", frmFHSWAdrs.TOP
SaveSetting "FHSWIN", "frmADRS", "Left", frmFHSWAdrs.Left
If FormLoaded(2) And SpouseForm Then
   frmFHSWSpse.Enabled = True
   frmFHSWSpse.SetFocus
   Else
   If FormLoaded(3) And MrecForm Then
      frmFHSWMisc.Enabled = True
      frmFHSWMisc.SetFocus
      Else
      frmFHSWUPDT.Enabled = True
      frmFHSWUPDT.SetFocus
   End If
End If
End Sub


Private Sub NAVButton_Click(Index As Integer)
Select Case Index
   Case 0: ' First
      ADRIX = ADRHIX: ADRRNO = ADRCHN(ADRIX)
      Call ShowADRInfo
      Call cmdSetup(1)
   Case 1: ' Prev
      If ADRIX < ADRHIX Then
         ADRIX = ADRIX + 1
         ADRRNO = ADRCHN(ADRIX)
         Call ShowADRInfo
         Call cmdSetup(1)
      End If
   Case 2: ' Next
      If ADRIX > 0 Then
         ADRIX = ADRIX - 1
         ADRRNO = ADRCHN(ADRIX)
         Call ShowADRInfo
         Call cmdSetup(1)
      End If
   Case 3: ' Last
      ADRIX = 1: ADRRNO = ADRCHN(ADRIX)
      Call ShowADRInfo
      Call cmdSetup(1)
End Select
End Sub


