VERSION 4.00
Begin VB.Form frmFHSWPare 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FHSWPARE - Parent Information"
   ClientHeight    =   4005
   ClientLeft      =   3330
   ClientTop       =   2580
   ClientWidth     =   4845
   Height          =   4410
   Left            =   3270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   Top             =   2235
   Width           =   4965
   Begin VB.Frame CMDFrame 
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   3360
      Width           =   4815
      Begin VB.CommandButton CMDButton 
         Cancel          =   -1  'True
         Caption         =   "Return"
         Height          =   255
         Index           =   4
         Left            =   3960
         TabIndex        =   35
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CMDButton 
         Caption         =   "Children"
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CMDButton 
         Caption         =   "Select Mother"
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CMDButton 
         Caption         =   "Select Father"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame PARENTFrame 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.TextBox MIData 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Index           =   7
         Left            =   2280
         TabIndex        =   34
         Top             =   2880
         Width           =   2415
      End
      Begin VB.TextBox MIData 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   840
         TabIndex        =   32
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox MIData 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Index           =   5
         Left            =   2280
         TabIndex        =   30
         Top             =   2520
         Width           =   2415
      End
      Begin VB.TextBox MIData 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   840
         TabIndex        =   28
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox MIData 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   4200
         TabIndex        =   26
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox MIData 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Index           =   2
         Left            =   840
         TabIndex        =   24
         Top             =   2160
         Width           =   2535
      End
      Begin VB.TextBox MIData 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   4200
         TabIndex        =   23
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox MIData 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Index           =   1
         Left            =   840
         TabIndex        =   21
         Top             =   1800
         Width           =   2895
      End
      Begin VB.TextBox FIData 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   4200
         TabIndex        =   19
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox FIData 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   4200
         TabIndex        =   17
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox FIData 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Index           =   7
         Left            =   2280
         TabIndex        =   15
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox FIData 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   840
         TabIndex        =   13
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox FIData 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Index           =   5
         Left            =   2280
         TabIndex        =   11
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox FIData 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   840
         TabIndex        =   9
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox FIData 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Index           =   2
         Left            =   840
         TabIndex        =   7
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox FIData 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Index           =   1
         Left            =   840
         TabIndex        =   6
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label 
         Caption         =   "in"
         Height          =   255
         Index           =   13
         Left            =   2040
         TabIndex        =   33
         Top             =   2880
         Width           =   255
      End
      Begin VB.Label Label 
         Caption         =   "Died:"
         Height          =   255
         Index           =   12
         Left            =   360
         TabIndex        =   31
         Top             =   2880
         Width           =   495
      End
      Begin VB.Label Label 
         Caption         =   "in"
         Height          =   255
         Index           =   11
         Left            =   2040
         TabIndex        =   29
         Top             =   2520
         Width           =   255
      End
      Begin VB.Label Label 
         Caption         =   "Born:"
         Height          =   255
         Index           =   10
         Left            =   360
         TabIndex        =   27
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Label 
         Caption         =   "AGE"
         Height          =   255
         Index           =   9
         Left            =   3840
         TabIndex        =   25
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label 
         Caption         =   "MID="
         Height          =   255
         Index           =   8
         Left            =   3840
         TabIndex        =   22
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label 
         Caption         =   "MOTHER:"
         Height          =   255
         Index           =   7
         Left            =   75
         TabIndex        =   20
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label 
         Caption         =   "AGE"
         Height          =   255
         Index           =   2
         Left            =   3840
         TabIndex        =   18
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label 
         Caption         =   "FID="
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   16
         Top             =   240
         Width           =   375
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   4800
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label 
         Caption         =   "in"
         Height          =   255
         Index           =   6
         Left            =   2040
         TabIndex        =   14
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label 
         Caption         =   "Died:"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   12
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label 
         Caption         =   "in"
         Height          =   255
         Index           =   4
         Left            =   2040
         TabIndex        =   10
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label 
         Caption         =   "Born:"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   8
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label 
         Caption         =   "FATHER:"
         Height          =   255
         Index           =   0
         Left            =   75
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmFHSWPare"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
DefInt A-Z
Dim FID, MID, AGL
Sub ViewPARENTS(xFID, xMID, xAGL)
FID = xFID: MID = xMID: AGL = xAGL - 1
If Not ENV.FFOPEN Then Call FAMOpen("WRITE")
For X = 0 To 7: FIData(X) = "": MIData(X) = "": Next X
foldch = -1: moldch = -1
If FID <> 0 Then
   CMDButton(1).Font.Bold = True
   Call FF1GetRec(FID): foldch = FF1Rec.OLDCH
   FIData(0) = Str$(FID)
   FIData(1) = RTrim$(FF1Rec.GIVEN)
   FIData(2) = RTrim$(FF1Rec.SURNM)
   FIData(3) = DATEDif$(FF1Rec.BMD, FF1Rec.BY, FF1Rec.DMD, FF1Rec.DY)
   FIData(4) = XSCRNDate$(FF1Rec.BMD, FF1Rec.BY)
   FIData(6) = XSCRNDate$(FF1Rec.DMD, FF1Rec.DY)
   If FF1Rec.BLOC <> 0 Then
      Call FF3GetRec(FF1Rec.BLOC)
      FIData(5) = RTrim$(FF3PLACE.PLACE1)
      FIData(7) = RTrim$(FF3PLACE.PLACE2)
   End If
   Else
   CMDButton(1).Enabled = False
End If
If MID <> 0 Then
   CMDButton(2).Font.Bold = True
   Call FF1GetRec(MID): moldch = FF1Rec.OLDCH
   MIData(0) = Str$(MID)
   MIData(1) = RTrim$(FF1Rec.GIVEN)
   MIData(2) = RTrim$(FF1Rec.SURNM)
   MIData(3) = DATEDif$(FF1Rec.BMD, FF1Rec.BY, FF1Rec.DMD, FF1Rec.DY)
   MIData(4) = XSCRNDate$(FF1Rec.BMD, FF1Rec.BY)
   MIData(6) = XSCRNDate$(FF1Rec.DMD, FF1Rec.DY)
   If FF1Rec.BLOC <> 0 Then
      Call FF3GetRec(FF1Rec.BLOC)
      MIData(5) = RTrim$(FF3PLACE.PLACE1)
      MIData(7) = RTrim$(FF3PLACE.PLACE2)
   End If
   Else: CMDButton(2).Enabled = False
End If
CMDButton(3).Font.Bold = (foldch <> 0) And (moldch <> 0)
frmFHSWPare.Show
End Sub


Private Sub CMDButton_Click(Index As Integer)
Select Case Index
   Case 1:
      Call frmFHSWUPDT.NRDisplay(FID, AGL)
   Case 2:
      Call frmFHSWUPDT.NRDisplay(MID, AGL)
   Case 3: frmFHSWPare.Enabled = False
      Call frmFHSWChld.ShowCHILDREN(FID, MID, AGL)
      Exit Sub
   Case 4:
End Select
Unload frmFHSWPare
End Sub

Private Sub Form_Load()
frmFHSWUPDT.Enabled = False
FormLoaded(3) = FormLoaded(3) Or ParentForm
' Set Default Form Location on Screen
frmHEIGHT = frmFHSWPare.Height
frmWIDTH = frmFHSWPare.WIDTH
t1 = GetSetting("FHSWIN", "frmPARE", "TOP", 0)
If t1 = 0 Then
   t00 = frmFHSWUPDT.TOP: w00 = frmFHSWUPDT.WIDTH
   l00 = frmFHSWUPDT.Left
   t1 = t00
   l1 = l00 - frmWIDTH
   If l1 < 0 Then
      t1 = frmFHSWUPDT.TOP + 2715
      l1 = l00 - frmWIDTH / 2
   End If
   Else
   t1 = 1440: l1 = 1440
End If
frmFHSWPare.TOP = GetSetting("FHSWIN", "frmPARE", "Top", t1)
frmFHSWPare.Left = GetSetting("FHSWIN", "frmPARE", "Left", l1)
End Sub


Private Sub Form_Unload(Cancel As Integer)
FormLoaded(3) = FormLoaded(3) And (255 - ParentForm)
SaveSetting "FHSWIN", "frmPARE", "Top", frmFHSWPare.TOP
SaveSetting "FHSWIN", "frmPARE", "Left", frmFHSWPare.Left
frmFHSWUPDT.Enabled = True
frmFHSWUPDT.SetFocus
End Sub


