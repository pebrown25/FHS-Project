VERSION 4.00
Begin VB.Form frmFHSWUPDT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FHSWUPDT - Update Family File Information"
   ClientHeight    =   4845
   ClientLeft      =   7710
   ClientTop       =   3675
   ClientWidth     =   4950
   Height          =   5250
   Left            =   7650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   Top             =   3330
   Width           =   5070
   Begin VB.Frame NRMNUframe 
      Height          =   615
      Left            =   0
      TabIndex        =   50
      Top             =   0
      Width           =   4935
      Begin VB.CommandButton cmdNRMNU 
         Caption         =   "Date Calc"
         Height          =   255
         Index           =   4
         Left            =   3120
         TabIndex        =   55
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdNRMNU 
         Caption         =   "Return"
         Height          =   255
         Index           =   5
         Left            =   4080
         TabIndex        =   54
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdNRMNU 
         Caption         =   "RELCalc"
         Height          =   255
         Index           =   3
         Left            =   2160
         TabIndex        =   53
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdNRMNU 
         Caption         =   "Search"
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   52
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdNRMNU 
         Caption         =   "Add Name"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   51
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame BRIDFrame 
      Height          =   615
      Left            =   0
      TabIndex        =   43
      Top             =   600
      Width           =   4935
      Begin VB.TextBox newRID 
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
         Left            =   2040
         MaxLength       =   5
         TabIndex        =   48
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdRID 
         Caption         =   "<"
         Height          =   255
         Left            =   2640
         TabIndex        =   47
         Top             =   255
         Width           =   255
      End
      Begin VB.CommandButton cmdSHOW 
         Caption         =   "Show"
         Default         =   -1  'True
         Height          =   255
         Left            =   3120
         TabIndex        =   46
         Top             =   255
         Width           =   615
      End
      Begin VB.CommandButton cmdPREV 
         Caption         =   "<-"
         Height          =   255
         Left            =   3960
         TabIndex        =   45
         Top             =   255
         Width           =   375
      End
      Begin VB.CommandButton cmdNEXT 
         Caption         =   "->"
         Height          =   255
         Left            =   4440
         TabIndex        =   44
         Top             =   255
         Width           =   375
      End
      Begin VB.Label Label 
         Caption         =   "Display Record for ID:"
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
         TabIndex        =   49
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame cmdFRAME 
      Height          =   615
      Left            =   0
      TabIndex        =   42
      Top             =   4200
      Width           =   4935
      Begin VB.CommandButton cmdBUTTON 
         Caption         =   "Change"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdBUTTON 
         Caption         =   "Parents"
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   24
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdBUTTON 
         Caption         =   "Children"
         Height          =   255
         Index           =   3
         Left            =   2040
         TabIndex        =   25
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdBUTTON 
         Caption         =   "Spouse"
         Height          =   255
         Index           =   4
         Left            =   3120
         TabIndex        =   26
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdBUTTON 
         Caption         =   "Other"
         Height          =   255
         Index           =   5
         Left            =   4080
         TabIndex        =   27
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame NRFrame 
      Height          =   3015
      Left            =   0
      TabIndex        =   20
      Top             =   1200
      Width           =   4935
      Begin VB.TextBox NRValue 
         Alignment       =   2  'Center
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
         Index           =   26
         Left            =   4200
         MaxLength       =   1
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   840
         Width           =   255
      End
      Begin VB.TextBox NRValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   16
         Left            =   4200
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox NRValue 
         Alignment       =   2  'Center
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
         Index           =   25
         Left            =   4320
         MaxLength       =   1
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   2640
         Width           =   255
      End
      Begin VB.TextBox NRValue 
         Alignment       =   2  'Center
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
         Index           =   23
         Left            =   2040
         MaxLength       =   1
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   2640
         Width           =   255
      End
      Begin VB.TextBox NRValue 
         Alignment       =   2  'Center
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
         Index           =   19
         Left            =   2160
         MaxLength       =   1
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   1920
         Width           =   255
      End
      Begin VB.TextBox NRValue 
         Alignment       =   2  'Center
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
         Left            =   2160
         MaxLength       =   1
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1200
         Width           =   255
      End
      Begin VB.TextBox NRValue 
         Alignment       =   2  'Center
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
         Left            =   3480
         MaxLength       =   1
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   840
         Width           =   270
      End
      Begin VB.CommandButton cmdMID 
         Caption         =   "<"
         Height          =   255
         Left            =   4080
         TabIndex        =   17
         Top             =   2640
         Width           =   255
      End
      Begin VB.TextBox NRValue 
         Alignment       =   1  'Right Justify
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
         Index           =   24
         Left            =   3480
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   2640
         Width           =   615
      End
      Begin VB.CommandButton cmdFID 
         Caption         =   "<"
         Height          =   255
         Left            =   1800
         TabIndex        =   14
         Top             =   2640
         Width           =   255
      End
      Begin VB.TextBox NRValue 
         Alignment       =   1  'Right Justify
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
         Index           =   22
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox NRValue 
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
         Index           =   21
         Left            =   1200
         TabIndex        =   12
         Top             =   2280
         Width           =   3375
      End
      Begin VB.TextBox NRValue 
         Height          =   315
         Index           =   20
         Left            =   3120
         TabIndex        =   11
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox NRValue 
         Height          =   315
         Index           =   18
         Left            =   1200
         TabIndex        =   9
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox NRValue 
         Height          =   315
         Index           =   15
         Left            =   3120
         TabIndex        =   7
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox NRValue 
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
         Index           =   17
         Left            =   1200
         TabIndex        =   8
         Top             =   1560
         Width           =   3375
      End
      Begin VB.TextBox NRValue 
         Height          =   315
         Index           =   13
         Left            =   1200
         TabIndex        =   5
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox NRValue 
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
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   1
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox NRValue 
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
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   2
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox NRValue 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   480
         MultiLine       =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox NRValue 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   4560
         MultiLine       =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label 
         Caption         =   "Sex:"
         Height          =   255
         Index           =   15
         Left            =   3840
         TabIndex        =   19
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label 
         Caption         =   "Age:"
         Height          =   255
         Index           =   8
         Left            =   3840
         TabIndex        =   0
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label 
         Caption         =   "Mother ID:"
         Height          =   255
         Index           =   14
         Left            =   2640
         TabIndex        =   41
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label 
         Caption         =   "Father ID:"
         Height          =   255
         Index           =   13
         Left            =   360
         TabIndex        =   40
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label 
         Caption         =   "PLACE:"
         Height          =   255
         Index           =   12
         Left            =   480
         TabIndex        =   39
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label 
         Caption         =   "TIME:"
         Height          =   255
         Index           =   11
         Left            =   2520
         TabIndex        =   38
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label 
         Caption         =   "Death DATE:"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   37
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label 
         Caption         =   "TIME:"
         Height          =   255
         Index           =   7
         Left            =   2520
         TabIndex        =   36
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label 
         Caption         =   "PLACE:"
         Height          =   255
         Index           =   9
         Left            =   480
         TabIndex        =   35
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label 
         Caption         =   "Birth DATE:"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   34
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label 
         Caption         =   "Use:"
         Height          =   255
         Index           =   4
         Left            =   3120
         TabIndex        =   33
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label 
         Caption         =   "Given Name:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   32
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label 
         Caption         =   "Surname:"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   31
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label 
         Caption         =   "ID="
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label 
         Caption         =   "RGL="
         Height          =   255
         Index           =   1
         Left            =   4080
         TabIndex        =   28
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmFHSWUPDT"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
DefInt A-Z
Dim RGL, CFID, CMID, CRID, CBRID
Dim UFID!, UMID!, ORGL, OBRID
Dim SRCHTable As TablePARMS, CHLDTable As TablePARMS
Dim FUN$, CMAX, CTSIZE, LastFLD
Dim HOLDITEM As Variant, xFT
Dim HRValue() As String
Dim CurCMD, ReShowDATA

Sub UnloadFHSWForms()
If FormLoaded(9) And MSGDForm Then Unload frmFHSWMsgD
If FormLoaded(9) And MSGForm Then Unload frmFHSWMsg
If FormLoaded(2) And NotesForm Then
   Unload frmFHSWNote: If Not OKAY Then Exit Sub
End If
If FormLoaded(1) And RelWForm Then Unload frmFHSWRelW
If FormLoaded(1) And SlctForm Then Unload frmFHSWSlct
If FormLoaded(2) And RelCalcForm Then Unload frmFHSWRCLC
If FormLoaded(2) And DateCalcForm Then Unload frmFHSWDCLC
If FormLoaded(3) And AddressForm Then Unload frmFHSWAdrs
If FormLoaded(3) And MrecForm Then Unload frmFHSWMisc
If FormLoaded(2) And SpouseForm Then Unload frmFHSWSpse
If FormLoaded(3) And ParentForm Then Unload frmFHSWPare
If FormLoaded(3) And ChildForm Then Unload frmFHSWChld
OKAY = -1
End Sub
Sub FindListEntry(xRID)
If (frmFHSWList.lstBOX.ListCount > 0) Then
   FOUND = 0: X = 0
   While (Not FOUND) And (X < frmFHSWList.lstBOX.ListCount)
      FOUND = (xRID = frmFHSWList.lstBOX.ItemData(X))
      If Not FOUND Then X = X + 1
   Wend
   If FOUND Then frmFHSWList.lstBOX.ListIndex = X
End If

End Sub

Sub NewNameRec()
   If Not ENV.FFOPEN Then Call FAMOpen("WRITE")
500 ' ADD NAME Record
   Call FF1GetNew
   If Not OKAY Then Exit Sub
   OBRID = CBRID: CBRID = FF1Hdr.MAXID
   If OBRID = 0 Then OBRID = -1
   ORGL = RGL: RGL = 0
   NameREC = FF1Rec
   LSet OBPLACE = "": LSet ODPLACE = ""
   LSet BPLACE = "": LSet DPLACE = ""
   ShowNameDATA
   frmFHSWUPDT.Height = 5250
   cmdSetup (3)
End Sub


Sub cmdSetup(cmdOPT)
X = cmdOPT
If X = 0 Then X = CurCMD: If X = 0 Then X = 1
CurCMD = X
For N = 1 To 5
    CMDButton(N).Visible = True
    CMDButton(N).Font.Bold = False
    CMDButton(N).Enabled = True
    CMDButton(N).Cancel = False
Next N
cmdSHOW.DEFAULT = False
NRFrame.Enabled = False
BRIDFrame.Enabled = True
NRMNUframe.Enabled = True
If CMDFrame.Visible Then CMDFrame.Enabled = True
If (CBRID <> 0) And (FF1Rec.RID <> CBRID) Then
   Call FF1GetRec(CBRID)
End If
Select Case X
   Case 1: ' Standard NAMEREC Commands
         cmdNRMNU(5).Cancel = True
         cmdSHOW.DEFAULT = True
         CMDButton(1).Caption = "Update"
         CMDButton(2).Caption = "Parents"
         CMDButton(2).Font.Bold = ((FF1Rec.FID <> 0) Or (FF1Rec.MID <> 0))
         CMDButton(2).Enabled = CMDButton(2).Font.Bold
         CMDButton(3).Caption = "Children"
         CMDButton(3).Font.Bold = (FF1Rec.OLDCH <> 0)
         CMDButton(3).Enabled = CMDButton(3).Font.Bold
         CMDButton(4).Caption = "Spouse"
         CMDButton(4).Font.Bold = (FF1Rec.SPOUSE <> 0)
         CMDButton(5).Font.Bold = ((FF1Rec.ADRS <> 0) Or (FF1Rec.COM <> 0) Or (FF1Rec.EVENT <> 0) Or (FF1Rec.EDUC <> 0) Or (FF1Rec.WORK <> 0) Or (FF1Rec.MIL <> 0) Or (FF1Rec.HLTH <> 0))
         CMDButton(5).Caption = "Other"
         If (CMDFrame.Visible And CMDButton(5).Enabled) Then CMDButton(5).SetFocus
   Case 2: ' Change Choices
         CMDButton(1).Caption = "Change"
         CMDButton(2).Caption = "Add"
         CMDButton(3).Caption = "Cancel"
         CMDButton(3).Cancel = True
         CMDButton(4).Visible = False
         CMDButton(5).Visible = False
         BRIDFrame.Enabled = False
         NRMNUframe.Enabled = False
   Case 3: ' Add/Update Choices
         NRFrame.Enabled = True
         CMDButton(1).Caption = "Save"
         CMDButton(2).Caption = "Cancel"
         CMDButton(3).Caption = "Restore"
         CMDButton(3).Cancel = True
         CMDButton(4).Visible = False
         CMDButton(5).Visible = False
         BRIDFrame.Enabled = False
         NRMNUframe.Enabled = False
         NRValue(12).SetFocus
   Case 5: ' Other Choices #1
         CMDButton(1).Caption = "Res"
         CMDButton(1).Font.Bold = (FF1Rec.ADRS <> 0)
         CMDButton(2).Caption = "Notes"
         CMDButton(2).Font.Bold = (FF1Rec.COM <> 0)
         CMDButton(3).Caption = "Events"
         CMDButton(3).Font.Bold = (FF1Rec.EVENT <> 0)
         CMDButton(4).Caption = "Other"
         CMDButton(4).Font.Bold = ((FF1Rec.EDUC <> 0) Or (FF1Rec.WORK <> 0) Or (FF1Rec.MIL <> 0) Or (FF1Rec.HLTH <> 0))
         CMDButton(5).Caption = "Main"
         CMDButton(5).Cancel = True
         CMDButton(4).SetFocus
   Case 6: ' Other Choices #2
         CMDButton(1).Caption = "Educ"
         CMDButton(1).Font.Bold = (FF1Rec.EDUC <> 0)
         CMDButton(2).Caption = "Work"
         CMDButton(2).Font.Bold = (FF1Rec.WORK <> 0)
         CMDButton(3).Caption = "Military"
         CMDButton(3).Font.Bold = (FF1Rec.MIL <> 0)
         CMDButton(4).Caption = "Medical"
         CMDButton(4).Font.Bold = (FF1Rec.HLTH <> 0)
         CMDButton(5).Caption = "Main"
         CMDButton(5).Cancel = True
         CMDButton(5).SetFocus
End Select
End Sub



Sub RestoreField(FldNum)
If RTrim$(NRValue(FldNum).TEXT) <> HRValue(FldNum) Then
   NRValue(FldNum).TEXT = HRValue(FldNum)
   OKAY = -1
   Else
   OKAY = 0
End If
End Sub

Sub UpdtNameREC()
660 ' Get Name Record DATA from Screen and Update File
   If Not ENV.FFOPEN Then Call FAMOpen("WRITE")
   Call FF1GetRec(CBRID)
   BPLACE = OBPLACE: DPLACE = ODPLACE
   Call ReadNameDATA
   If (UFID! <> FF1Rec.FID) Or (UMID! <> FF1Rec.MID) Or ((UFID! > 0) And (UFID! = UMID!)) Then
      Call PutMSG("UPDTM003"): OKAY = 0
      Else
      NameREC = FF1Rec: xUFID = UFID!: xUMID = UMID!
      Call FF1GetRec(FF1Rec.RID): CFID = FF1Rec.FID: CMID = FF1Rec.MID
      BirthCHG = ((NameREC.BY <> FF1Rec.BY) Or (NameREC.BMD <> FF1Rec.BMD))
      If (BirthCHG Or (CFID <> xUFID)) And (CFID > 0) Then
         Call UpdtRelPTRS("REM", CFID)
      End If
      If (BirthCHG Or (CMID <> xUMID)) And (CMID > 0) Then
         Call UpdtRelPTRS("REM", CMID)
      End If
      If (BirthCHG Or (CFID <> xUFID)) And (xUFID > 0) Then
         Call UpdtRelPTRS("ADD", xUFID)
      End If
      If (BirthCHG Or (CMID <> xUMID)) And (xUMID > 0) Then
         Call UpdtRelPTRS("ADD", xUMID)
      End If
      Call UpdtPLACE(NameREC.BLOC, 1, NameREC.RID)
      FF1Rec = NameREC
      Call FF1PutREC
      Call ShowNameDATA
   End If

End Sub

Sub ReadNameDATA()
'SHARED FMT006, UFID!, UMID!
  ' Get Name Data From SCREEN
    OKAY = 0
    FF1Rec = NameREC
    FF1Rec.SURNM = NRValue(10)
    X$ = NRValue(11)
    If X$ = SNGLKEY$(1) Then X$ = "Y"
    FF1Rec.STATUS2 = FFSTAT$("SNU", X$)
    FF1Rec.GIVEN = NRValue(12)
    X$ = NRValue(13)
    FF1Rec.BMD = DATEMD(X$, DFM)
    FF1Rec.BY = DATEY(X$, DFM)
    X$ = NRValue(14)
    FF1Rec.STATUS = FFSTAT$("BD", X$)
    FF1Rec.BTM = TIMEHHMM(NRValue(15))
    BPLACE = NRValue(17)
    X$ = NRValue(18)
    FF1Rec.DMD = DATEMD(X$, DFM)
    FF1Rec.DY = DATEY(X$, DFM)
    X$ = NRValue(19)
    FF1Rec.STATUS = FFSTAT$("DD", X$)
    FF1Rec.DTM = TIMEHHMM(NRValue(20))
    DPLACE = NRValue(21)
    FF1Rec.SEX = NRValue(26)
    CFID = FF1Rec.FID
    CMID = FF1Rec.MID
    UFID! = Val(NRValue(22))
    FF1Rec.STATUS = FFSTAT$("FID", NRValue(23))
    UMID! = Val(NRValue(24))
    FF1Rec.STATUS = FFSTAT$("MID", NRValue(25))
    If Not ((UFID! > FF1Hdr.MAXID) Or (UFID! = FF1Rec.RID)) Then
       FF1Rec.FID = UFID!
    End If
    If Not ((UMID! > FF1Hdr.MAXID) Or (UMID! = FF1Rec.RID)) Then
       FF1Rec.MID = UMID!
    End If
    ShowNameDATA
End Sub


Sub NRDisplay(xRID, xGL)
600 ' Display NAME Record Number BRID
   If xRID < 1 Then Call ErrBEEP(0): Exit Sub
   If Not ENV.FFOPEN Then Call FAMOpen("WRITE")
   Call FF1GetRec(xRID): If Not OKAY Then Exit Sub
   CBRID = xRID: newRID = CBRID
   RGL = xGL
   CFID = FF1Rec.FID: CMID = FF1Rec.MID
   NameREC = FF1Rec
   LSet BPLACE = "": LSet DPLACE = ""
   If FF1Rec.BLOC <> 0 Then
      Call FF3GetRec(FF1Rec.BLOC)
      If FF3PFX.RTYPE = Chr$(12) Then
         If FF3PLACE2.PNEXT <> 0 Then Call FF3GetRec(FF3PLACE2.PNEXT)
      End If
   End If
   OBPLACE = BPLACE: ODPLACE = DPLACE
   Call ShowNameDATA: cmdSetup (CurCMD)
   HRValue(2) = RTrim$(NRValue(2))
   For X = 10 To 26
      HRValue(X) = RTrim$(NRValue(X).TEXT)
   Next X
   frmFHSWUPDT.Height = 5250
   NRFrame.Enabled = False
End Sub

Sub ShowNameDATA()
    NRValue(1) = Str$(CBRID)
    NRValue(2) = LTrim$(Str$(RGL))
    NRValue(10) = RTrim$(FF1Rec.SURNM)
    NRValue(11) = FDSTAT$("SNU")
    NRValue(12) = RTrim$(FF1Rec.GIVEN)
    NRValue(13) = XSCRNDate$(FF1Rec.BMD, FF1Rec.BY)
    NRValue(14) = Trim$(FDSTAT$("BD"))
    NRValue(15) = XSCRNTime$(FF1Rec.BTM)
    NRValue(17) = RTrim$(BPLACE)
    NRValue(18) = XSCRNDate$(FF1Rec.DMD, FF1Rec.DY)
    NRValue(19) = Trim$(FDSTAT$("DD"))
    NRValue(20) = XSCRNTime$(FF1Rec.DTM)
    NRValue(21) = RTrim$(DPLACE)
    NRValue(26) = RTrim$(FF1Rec.SEX)
    NRValue(16) = DATEDif$(FF1Rec.BMD, FF1Rec.BY, FF1Rec.DMD, FF1Rec.DY)
    CFID = FF1Rec.FID
    NRValue(22) = LTrim$(Str$(CFID))
    NRValue(23) = Trim$(FDSTAT$("FID"))
    CMID = FF1Rec.MID
    NRValue(24) = LTrim$(Str$(CMID))
    NRValue(25) = Trim$(FDSTAT$("MID"))
    ReShowDATA = 0
End Sub

Private Sub CMDButton_Click(Index As Integer)
X = Index
CmdOption:
Select Case CMDButton(X).Caption
   Case "Add": Call NewNameRec
        If Not OKAY Then cmdSetup (1)
   Case "Cancel": ' Cancel Update
        If X = 2 Then
           If Not ENV.FFOPEN Then Call FAMOpen("WRITE")
           If OBRID <> 0 Then
              FF1Hdr.MAXID = FF1Hdr.MAXID - 1
              FFCHANGED = -1
              CBRID = OBRID
              If CBRID < 0 Then CBRID = 0: OBRID = 0
           End If
           If CBRID > 0 Then
              Call NRDisplay(CBRID, 0)
              Else
              frmFHSWUPDT.Height = 1650
           End If
        End If
        cmdSetup (1)
   Case "Change": cmdSetup (3): OBRID = 0
   Case "Children": X = 0
        Call frmFHSWChld.ShowCHILDREN(CBRID, X, RGL)
   Case "Educ":
        Call frmFHSWMisc.ViewMiscRec(FF1Rec.EDUC, CBRID, 7)
   Case "Events":
        Call frmFHSWEvnt.ViewEVENTS(FF1Rec.EVENT, 1, CBRID)
   Case "Main": cmdSetup (1)
   Case "Medical"
        Call frmFHSWMisc.ViewMiscRec(FF1Rec.HLTH, CBRID, 10)
   Case "Military"
        Call frmFHSWMisc.ViewMiscRec(FF1Rec.MIL, CBRID, 9)
   Case "Notes":
        Call frmFHSWNote.ViewCOMMENTS(FF1Rec.COM, 1, CBRID)
   Case "Other": ' Other Options
        If Index = 5 Then cmdSetup (5) Else cmdSetup (6)
   Case "Parents":
        Call frmFHSWPare.ViewPARENTS(CFID, CMID, RGL)
        ReShowDATA = -1
   Case "Res":
        Call frmFHSWAdrs.ViewADDRESS(FF1Rec.ADRS, 1, CBRID)
        frmFHSWUPDT.Enabled = False
   Case "Restore": ' Restore Last Value for Field
        If (RTrim$(NRValue(LastFLD).TEXT) <> HOLDITEM) Then
           NRValue(LastFLD).TEXT = HOLDITEM
           Else
           RestoreField (LastFLD)
           'If Not OKAY Then X = 2: GoTo CmdOption
        End If
        NRValue(LastFLD).SetFocus
   Case "Save": ' Save Changes to Name Record
      Call UpdtNameREC
      Call cmdSetup(1)
   Case "Spouse": Call frmFHSWSpse.ViewMARRIAGES(CBRID, FF1Rec.SPOUSE, RGL)
      ReShowDATA = -1
   Case "Update":
        cmdSetup (2)
   Case "Work"
        Call frmFHSWMisc.ViewMiscRec(FF1Rec.WORK, CBRID, 8)
End Select

End Sub


Private Sub cmdFID_Click()
X = Val(NRValue(22).TEXT)
Call frmFHSWList.SelectName(X)
If Not ENV.FFOPEN Then Call FAMOpen("WRITE")
If a <> 27 Then
   NRValue(22).TEXT = fnx$(X)
End If
End Sub

Private Sub cmdMID_Click()
X = Val(NRValue(24).TEXT)
Call frmFHSWList.SelectName(X)
If a <> 27 Then
   NRValue(24).TEXT = fnx$(X)
End If
End Sub

Private Sub cmdNEXT_Click()
xRID = 0
If (frmFHSWList.lstBOX.ListCount > 0) Then
   FindListEntry (CBRID)
   If frmFHSWList.lstBOX.ListIndex < (frmFHSWList.lstBOX.ListCount - 1) Then
      frmFHSWList.lstBOX.ListIndex = frmFHSWList.lstBOX.ListIndex + 1
      xRID = frmFHSWList.lstBOX.ItemData(frmFHSWList.lstBOX.ListIndex)
   End If
   Else: If CBRID < FF1Hdr.MAXID Then xRID = CBRID + 1
End If
If xRID > 0 Then RGL = 0: Call NRDisplay(xRID, 0)
End Sub

Private Sub cmdNRMNU_Click(Index As Integer)
Select Case Index
  Case 1: ' Add New Record
           Call NewNameRec
           If Not OKAY Then cmdSetup (1)
  Case 2: ' Search Family File
           frmFHSWSlct.Show
           CMDFrame.Enabled = False
  Case 3: ' Relationship Calculator
           Call frmFHSWRCLC.RelCALC(CBRID)
           CMDFrame.Enabled = False
  Case 4: ' Date Calculator
           frmFHSWDCLC.Show
           ReShowDATA = -1
           Exit Sub
  Case 5: Unload frmFHSWUPDT: Exit Sub
End Select
NRMNUframe.Enabled = False
BRIDFrame.Enabled = False
End Sub

Private Sub cmdPREV_Click()
X = 0
If (frmFHSWList.lstBOX.ListCount > 0) Then
   FindListEntry (CBRID)
   If frmFHSWList.lstBOX.ListIndex > 0 Then
      frmFHSWList.lstBOX.ListIndex = frmFHSWList.lstBOX.ListIndex - 1
      X = frmFHSWList.lstBOX.ItemData(frmFHSWList.lstBOX.ListIndex)
   End If
   Else: If CBRID > 0 Then X = CBRID - 1
End If
If X > 0 Then RGL = 0: Call NRDisplay(X, 0)
End Sub

Private Sub cmdRID_Click()
xRID = Val(newRID.TEXT)
Call frmFHSWList.SelectName(xRID)
If a <> 27 Then
   RGL = 0
   If xRID > 0 Then Call NRDisplay(xRID, 0)
End If
End Sub









Private Sub cmdSHOW_Click()
If (newRID.TEXT) <> "WRITE" Then
   Z = Val(newRID.TEXT)
   Call NRDisplay(Z, 0)
End If
End Sub

Private Sub Form_Activate()
If Not ENV.FFOPEN Then Call FAMOpen("WRITE")
If ReShowDATA Then Call NRDisplay(CBRID, AGL)
cmdSetup (CurCMD)
End Sub

Private Sub Form_GotFocus()
If ReShowDATA Then Call NRDisplay(CBRID, AGL)
cmdSetup (CurCMD)
End Sub


Private Sub Form_Load()
Call PgmPREP("FHSUPDT")
Call FAMOpen("WRITE"): If Not OKAY Then Exit Sub
FormLoaded(2) = FormLoaded(2) Or UpdtForm
' Set Default Form Location on Screen
frmHEIGHT = 1650
frmWIDTH = frmFHSWUPDT.WIDTH
t1 = GetSetting("FHSWIN", "frmUPDT", "TOP", 0)
If t1 = 0 Then
   t00 = mnuFHSWMain.TOP: w00 = mnuFHSWMain.WIDTH
   l00 = mnuFHSWMain.Left
   t1 = (t00 + 1060)
   l1 = l00 + (w00 - frmWIDTH) / 2: If l1 < 0 Then l1 = 1
   Else
   t1 = 1440: l1 = 1440
End If
frmFHSWUPDT.TOP = GetSetting("FHSWIN", "frmUPDT", "Top", t1)
frmFHSWUPDT.Left = GetSetting("FHSWIN", "frmUPDT", "Left", l1)
ReDim HRValue(26)
CurCMD = 0: CBRID = 0
'If (CBRID > 0) And (CBRID <= FF1Hdr.MAXID) Then
'   Call NRDisplay(CBRID, 0)
'   Else
   frmFHSWUPDT.Height = 1650
'End If
frmFHSWUPDT.Show
'If FF1Hdr.MAXID = 0 Then
'   Call NewNameRec: cmdSetup (3)
'   Else
'   newRID.SetFocus
'End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call UnloadFHSWForms
If Not OKAY Then Cancel = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call FamCLOSE
FormLoaded(2) = FormLoaded(2) And (255 - UpdtForm)
SaveSetting "FHSWIN", "frmUPDT", "Top", frmFHSWUPDT.TOP
SaveSetting "FHSWIN", "frmUPDT", "Left", frmFHSWUPDT.Left
FamUpdate = 0
mnuFHSWMain.SetFocus
End Sub





Private Sub NRValue_GotFocus(Index As Integer)
HOLDITEM = RTrim$(NRValue(Index).TEXT)
Select Case Index
   Case 11, 26: xFT = 5
   Case 13, 18: xFT = 3
   Case 15, 20: xFT = 4
   Case 22, 24: xFT = 2
   Case Else: xFT = 1
End Select
NRValue(Index).SelStart = 0
If (xFT > 2) Then
   NRValue(Index).SelLength = 0
   Else
   NRValue(Index).SelLength = Len(RTrim$(NRValue(Index).TEXT))
End If
LastFLD = Index
End Sub


Private Sub NRValue_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If xFT > 1 Then
   XVALUE$ = NRValue(Index).TEXT
   xSS = NRValue(Index).SelStart
   xSL = NRValue(Index).SelLength
   Call FldUpdate(KeyCode, Shift, xFT, xSS, xSL, XVALUE$)
   NRValue(Index).TEXT = XVALUE$
   NRValue(Index).SelStart = xSS
   NRValue(Index).SelLength = xSL
End If
End Sub


Private Sub NRValue_KeyPress(Index As Integer, KeyAscii As Integer)
If xFT > 1 Then KeyAscii = 0
End Sub


