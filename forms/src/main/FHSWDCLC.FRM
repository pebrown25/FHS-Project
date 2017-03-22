VERSION 4.00
Begin VB.Form frmFHSWDCLC 
   Caption         =   "FHSWDCLC - Date Calculator"
   ClientHeight    =   5085
   ClientLeft      =   7545
   ClientTop       =   4605
   ClientWidth     =   5265
   Height          =   5490
   Left            =   7485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   Top             =   4260
   Width           =   5385
   Begin VB.TextBox DataField 
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
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   40
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox DataField 
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
      Index           =   18
      Left            =   2760
      MaxLength       =   2
      TabIndex        =   39
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox DataField 
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
      Left            =   4080
      MaxLength       =   3
      TabIndex        =   38
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton cmdCOMPUTE 
      Caption         =   "Compute"
      Default         =   -1  'True
      Height          =   315
      Left            =   3360
      TabIndex        =   11
      Top             =   2115
      Width           =   855
   End
   Begin VB.CommandButton cmdCLEAR 
      Caption         =   "Clear"
      Height          =   315
      Left            =   4320
      TabIndex        =   12
      Top             =   2115
      Width           =   735
   End
   Begin VB.TextBox DataField 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   2160
      TabIndex        =   32
      Top             =   645
      Width           =   495
   End
   Begin VB.TextBox DataField 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   4800
      TabIndex        =   31
      Top             =   645
      Width           =   495
   End
   Begin VB.TextBox DataField 
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
      Index           =   10
      Left            =   960
      MaxLength       =   1
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton cmdRETURN 
      Caption         =   "Return"
      Height          =   315
      Left            =   4320
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Frame frameCalendar1 
      Caption         =   "Calendar"
      Height          =   525
      Left            =   105
      TabIndex        =   28
      Top             =   2400
      Width           =   5070
      Begin VB.CommandButton cmdSHOW 
         Caption         =   "Date1"
         Height          =   315
         Index           =   0
         Left            =   3360
         TabIndex        =   14
         Top             =   150
         Width           =   735
      End
      Begin VB.ComboBox cboCalendarType 
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS LineDraw"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "FHSWDCLC.frx":0000
         Left            =   1440
         List            =   "FHSWDCLC.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   150
         Width           =   1815
      End
      Begin VB.CommandButton cmdSHOW 
         Caption         =   "Date 2"
         Height          =   315
         Index           =   1
         Left            =   4200
         TabIndex        =   15
         Top             =   150
         Width           =   735
      End
      Begin VB.Label Label 
         Caption         =   "Type Calendar:"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   29
         Top             =   195
         Width           =   1215
      End
   End
   Begin VB.Frame frameCalendar2 
      Height          =   2055
      Left            =   120
      TabIndex        =   26
      Top             =   2880
      Visible         =   0   'False
      Width           =   5070
      Begin VB.CommandButton cmdCHG 
         Caption         =   ">"
         Height          =   255
         Index           =   3
         Left            =   3840
         TabIndex        =   37
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmdCHG 
         Caption         =   "<"
         Height          =   255
         Index           =   2
         Left            =   3600
         TabIndex        =   36
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmdCHG 
         Caption         =   ">"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   35
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton cmdCHG 
         Caption         =   "<"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   34
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox DataField 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Index           =   20
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   33
         Top             =   220
         Width           =   1695
      End
      Begin VB.CommandButton cmdHide 
         Caption         =   "Hide"
         Height          =   315
         Left            =   4320
         TabIndex        =   30
         Top             =   200
         Width           =   615
      End
      Begin MSGrid.Grid gridCalendar 
         Height          =   1335
         Left            =   600
         TabIndex        =   27
         Top             =   600
         Width           =   3975
         _Version        =   65536
         _ExtentX        =   7011
         _ExtentY        =   2355
         _StockProps     =   77
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Rows            =   6
         Cols            =   7
         FixedCols       =   0
         ScrollBars      =   0
         GridLines       =   0   'False
         HighLight       =   0   'False
         MouseIcon       =   "FHSWDCLC.frx":0026
      End
   End
   Begin VB.TextBox DataField 
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
      Index           =   16
      Left            =   1440
      MaxLength       =   7
      TabIndex        =   10
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox DataField 
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
      Index           =   15
      Left            =   2760
      MaxLength       =   3
      TabIndex        =   9
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox DataField 
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
      MaxLength       =   4
      TabIndex        =   8
      Top             =   1800
      Width           =   615
   End
   Begin VB.ComboBox cboDateFormat 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS LineDraw"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "FHSWDCLC.frx":0042
      Left            =   1320
      List            =   "FHSWDCLC.frx":005B
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox DataField 
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
      Index           =   13
      Left            =   4080
      MaxLength       =   3
      TabIndex        =   7
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox DataField 
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
      Left            =   2760
      MaxLength       =   2
      TabIndex        =   6
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox DataField 
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
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox DataField 
      Height          =   315
      Index           =   1
      Left            =   3360
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox DataField 
      Height          =   315
      Index           =   0
      Left            =   720
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label 
      Caption         =   "years"
      Height          =   255
      Index           =   13
      Left            =   2160
      TabIndex        =   43
      Top             =   1485
      Width           =   495
   End
   Begin VB.Label Label 
      Caption         =   "weeks"
      Height          =   255
      Index           =   12
      Left            =   3360
      TabIndex        =   42
      Top             =   1485
      Width           =   615
   End
   Begin VB.Label Label 
      Caption         =   "days"
      Height          =   255
      Index           =   10
      Left            =   4680
      TabIndex        =   41
      Top             =   1485
      Width           =   375
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   5280
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label 
      Caption         =   "days"
      Height          =   255
      Index           =   9
      Left            =   2280
      TabIndex        =   25
      Top             =   2205
      Width           =   615
   End
   Begin VB.Label Label 
      Caption         =   "days"
      Height          =   255
      Index           =   8
      Left            =   3360
      TabIndex        =   24
      Top             =   1845
      Width           =   495
   End
   Begin VB.Label Label 
      Caption         =   "years"
      Height          =   255
      Index           =   7
      Left            =   2160
      TabIndex        =   23
      Top             =   1845
      Width           =   495
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5280
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label 
      Caption         =   "Date Format:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   22
      Top             =   150
      Width           =   1095
   End
   Begin VB.Label Label 
      Caption         =   "days"
      Height          =   255
      Index           =   6
      Left            =   4680
      TabIndex        =   21
      Top             =   1125
      Width           =   375
   End
   Begin VB.Label Label 
      Caption         =   "months"
      Height          =   255
      Index           =   5
      Left            =   3360
      TabIndex        =   20
      Top             =   1125
      Width           =   615
   End
   Begin VB.Label Label 
      Caption         =   "years"
      Height          =   255
      Index           =   4
      Left            =   2160
      TabIndex        =   19
      Top             =   1125
      Width           =   495
   End
   Begin VB.Label Label 
      Caption         =   "Date Dif:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   18
      Top             =   1125
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "Date 2:"
      Height          =   255
      Index           =   2
      Left            =   2760
      TabIndex        =   17
      Top             =   630
      Width           =   615
   End
   Begin VB.Label Label 
      Caption         =   "Date 1:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   630
      Width           =   735
   End
End
Attribute VB_Name = "frmFHSWDCLC"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
DefInt A-Z
Dim DateS(5) As DateFields
Dim DateL As DateFields, DateH As DateFields
Dim DateC As DateFields
Dim difyyyy, difmm, difww, difdd, difddd As Double, adat As Double
Dim difsign, difdays As Double, difwday As Double
Dim DateROW, DateCOL
Dim ldiftype, LastDataFLD, CalDate
Dim frmHEIGHT, frmWIDTH

Private Sub BuildCalendar()
   If frameCalendar2.Visible = False Then Exit Sub
   DateW = DateC
   DataField(20) = MID$(RMONTB$(DateW.mm + 1), 3, 255) + " " + ZFILL$(DateW.yyyy, 4)
   'gridCalendar.Enabled = False
   gridCalendar.Cols = 7
   Call ChkLY
   ' Find Row of last day of month
   adat = DateW.adat + days_table(DateW.mm) - DateW.DD
   ec = ((adat - 2) Mod 7)
   gridCalendar.Rows = ((days_table(DateW.mm) - ec + 5) \ 7) + 2
   gridCalendar.Height = (gridCalendar.RowHeight(0)) * gridCalendar.Rows
   ' Set Height of frame
   frameCalendar2.Height = 720 + gridCalendar.Height
   frmHEIGHT = frameCalendar2.TOP + frameCalendar2.Height + 450
   frmFHSWDCLC.Height = frmHEIGHT
   ' Set Alignment of day headings to center
   gridCalendar.Row = 0
   For X = 0 To 6
       gridCalendar.ColAlignment(X) = 2
       gridCalendar.FixedAlignment(X) = 2
       gridCalendar.Col = X
       gridCalendar.TEXT = Left$(day_table(X + 1), 1)
   Next X
   ' Set calendar entries to ""
   For X = 1 To gridCalendar.Rows - 1
       gridCalendar.Row = X
       For Y = 0 To gridCalendar.Cols - 1
           gridCalendar.Col = Y
           gridCalendar.TEXT = ""
       Next Y
   Next X
   ' Find location of Day 1 on calendar
   DateCOL = ((DateW.adat - DateW.DD - 1) Mod WkLth)
   While DateCOL < 0
      DateCOL = DateCOL + WkLth
   Wend
   DateROW = 1
   Call ChkLY: ML = days_table(DateW.mm)
   ' Fill in dates on calendar
   For X = 1 To ML
      gridCalendar.Row = DateROW
      gridCalendar.Col = DateCOL
      gridCalendar.TEXT = X
      DateCOL = DateCOL + 1
      If (DateCOL > (WkLth - 1)) Then
         DateCOL = 0
         DateROW = DateROW + 1
      End If
   Next X
   ' Find location of DateW on calendar
   DateCOL = DateW.wday - 1
   While DateCOL < 0
      DateCOL = DateCOL + WkLth
   Wend
   w = DateW.wday - WkLth * (DateW.wday = 0)
   DateROW = ((DateW.DD - w + WkLth - 1) \ WkLth) + 1
   gridCalendar.Row = DateROW
   gridCalendar.Col = DateCOL
   gridCalendar.SetFocus
   gridCalendar.Refresh
   DateC = DateW
End Sub


Private Sub ComputeDate(typedif)
  DateW = DateS(0)
  Select Case typedif
     Case 1: ' years, months, days
          If difsign > 0 Then
             DateW.yyyy = DateW.yyyy + difyyyy
             DateW.mm = DateW.mm + difmm
             While DateW.mm > 12
                DateW.mm = DateW.mm - 12
                DateW.yyyy = DateW.yyyy + 1
             Wend
             Call ChkLY
             If DateW.DD > days_table(DateW.mm) Then
                DateW.DD = days_table(DateW.mm)
             End If
             DateW.DD = DateW.DD + difdd
             While DateW.DD > days_table(DateW.mm)
                DateW.DD = DateW.DD - days_table(DateW.mm)
                DateW.mm = DateW.mm + 1
                If DateW.mm > 12 Then
                   DateW.yyyy = DateW.yyyy + 1: Call ChkLY
                   DateW.mm = 1
                End If
             Wend
          End If
          If difsign < 0 Then
             DateW.adat = DateW.adat - difdd
             Call adat2yyyyddd
             Call ddd2mmdd
             DateW.mm = DateW.mm - difmm
             While DateW.mm < 1
                DateW.mm = DateW.mm + 12
                DateW.yyyy = DateW.yyyy - 1
             Wend
             DateW.yyyy = DateW.yyyy - difyyyy
             If DateW.yyyy < 1 Then DateW.yyyy = 1
          End If
          Call mmdd2ddd
          Call yyyyddd2adat
     Case 2: ' years, days
          If difsign > 0 Then
             DateW.yyyy = DateW.yyyy + difyyyy
             Call mmdd2ddd
             Call yyyyddd2adat
             DateW.ddd = DateW.ddd + difddd
             Call ddd2mmdd
             Call yyyyddd2adat
          End If
          If difsign < 0 Then
             DateW.adat = DateW.adat - difddd
             Call adat2yyyyddd
             Call ddd2mmdd
             DateW.yyyy = DateW.yyyy - difyyyy
             If DateW.yyyy < 1 Then DateW.yyyy = 1
             Call mmdd2ddd
             Call yyyyddd2adat
         End If
     Case 3: ' days
             DateW.adat = DateW.adat + difsign * difdays
             Call adat2yyyyddd
             DateW.mm = 0
             Call ddd2mmdd
     Case 4: ' years, weeks and days
          difddd = WkLth * difww + difwday
          If difsign > 0 Then
             DateW.yyyy = DateW.yyyy + difyyyy
             Call mmdd2ddd
             Call yyyyddd2adat
             DateW.ddd = DateW.ddd + difddd
             Call ddd2mmdd
             Call yyyyddd2adat
          End If
          If difsign < 0 Then
             DateW.adat = DateW.adat - difddd
             Call adat2yyyyddd
             Call ddd2mmdd
             DateW.yyyy = DateW.yyyy - difyyyy
             If DateW.yyyy < 0 Then DateW.yyyy = 0
             Call mmdd2ddd
             Call yyyyddd2adat
         End If
  End Select
  If DateW.adat > maxADAT Then
     DateW.adat = maxADAT
     Call adat2yyyyddd
     Call ddd2mmdd
  End If
  DateS(1) = DateW
  Call ShowDate(2)
  Call ComputeDIF
End Sub

Private Sub ComputeDIF()
   difyyyy = 0: difmm = 0: difww = 0: difdd = 0
   difddd = 0: difdays = 0
   If (DateS(0).adat <= DateS(1).adat) Then
      DateL = DateS(0): DateH = DateS(1): difsign = 1
      Else
      DateL = DateS(1): DateH = DateS(0): difsign = -1
   End If
   difdays = DateH.adat - DateL.adat
   'If difsign > 0 Then
      difdd = DateH.DD - DateL.DD
      If difdd < 0 Then
         DateW = DateH: Call ChkLY
         If DateH.mm > 1 Then
            DateH.mm = DateH.mm - 1
            dlm = days_table(DateH.mm)
            Else
            DateH.mm = 12: dlm = 31
            DateH.yyyy = DateH.yyyy - 1
         End If
         DateH.DD = DateH.DD + dlm
         If DateW.DD > dlm Then
            difdd = DateH.DD
            Else
            difdd = difdd + dlm
         End If
      End If
      DateW = DateL
      DateW.yyyy = DateH.yyyy + difyyyy: Call ChkLY
      If DateW.mm > DateH.mm Then
         DateW.yyyy = DateW.yyyy - 1
         Call ChkLY
      End If
      difyyyy = DateW.yyyy - DateL.yyyy
      difmm = DateH.mm - DateW.mm
      If difmm < 0 Then
         difmm = difmm + 12
      End If
      Call mmdd2ddd
      Call yyyyddd2adat
      difddd = DateH.adat - DateW.adat
      difww = (DateH.adat - DateW.adat) \ WkLth
      difwday = DateH.adat - DateW.adat - WkLth * difww
      
   'End If
   Call ShowDif
End Sub


Private Sub ShowDate(dnum As Integer)
   DateW = DateS(dnum - 1)
   DataField(dnum - 1) = FormattedDate()
   If (DateW.wday > 0) And (DateW.wday < (WkLth + 1)) Then
      DataField(dnum + 1) = day_table(DateW.wday)
      Else
      DataField(dnum + 1) = DateW.wday
   End If
   If CalDate = dnum - 1 Then
      BuildCalendar
   End If
End Sub
Private Sub ShowDif()
   If difsign < 0 Then
      DataField(10) = "-"
      Else
      DataField(10) = "+"
   End If
   DataField(11) = difyyyy
   DataField(12) = difmm
   DataField(13) = difdd
   DataField(14) = difyyyy
   DataField(15) = difddd
   DataField(16) = difdays
   DataField(17) = difyyyy
   DataField(18) = difww
   DataField(19) = difwday
End Sub



Private Sub cboCalendarType_Click()
   Call SetCalendarType(cboCalendarType.ListIndex, cboDateFormat.ListIndex)
   DateW = DateS(0)
   Call adat2yyyyddd
   DateS(0) = DateW
   Call ShowDate(1)
   DateW = DateS(1)
   Call adat2yyyyddd
   DateS(1) = DateW
   Call ShowDate(2)
   Call ComputeDIF
   If frameCalendar2.Visible Then
      DateW = DateC
      Call adat2yyyyddd
      DateC = DateW
      Call BuildCalendar
   End If
   
End Sub

Private Sub cboDateFormat_Click()
Call SetCalendarType(cboCalendarType.ListIndex, cboDateFormat.ListIndex)
If cboDateFormat.ListIndex < 3 Then
   CONFIG.DFM = Trim$(Str$(cboDateFormat.ListIndex + 1))
End If
ShowDate (1): ShowDate (2)
End Sub


Private Sub cmdCHG_Click(Index As Integer)
DateW = DateC
Select Case Index
   Case 0: DateW.mm = DateW.mm - 1
           If DateW.mm < 1 Then
              If DateW.yyyy > 1 Then
                 DateW.mm = 12: DateW.yyyy = DateW.yyyy - 1
                 Else
                 DateW.mm = 1
              End If
           End If
   Case 1: DateW.mm = DateW.mm + 1
           If DateW.mm > 12 Then
              If DateW.yyyy < 9999 Then
                 DateW.mm = 1: DateW.yyyy = DateW.yyyy + 1
                 Else
                 DateW.mm = 12
              End If
           End If
   Case 2: DateW.yyyy = DateW.yyyy - 1
   Case 3: DateW.yyyy = DateW.yyyy + 1
End Select
If DateW.yyyy < 1 Then DateW.yyyy = 1
If DateW.yyyy > 9999 Then DateW.yyyy = 9999
If DateW.DD > days_table(DateW.mm) Then DateW.DD = days_table(DateW.mm)
Call mmdd2ddd: Call yyyyddd2adat
DateC = DateW
Call BuildCalendar
End Sub

Private Sub cmdCLEAR_Click()
difsign = 1: difdays = 0
difyyyy = 0: difmm = 0: difdd = 0: difddd = 0
difww = 0: difwday = 0
Call ShowDif
End Sub

Private Sub cmdCOMPUTE_Click()
   If LastDataFLD < 2 Then
      XDATE = DataField(LastDataFLD)
      If (DFM = 6) Then
         MA = 3 + 2 * (CalType > 0)
         If Val(XDATE) < MA Then XDATE = Str$(MA)
      End If
      Call UnformatDate(XDATE, -1)
      DateS(LastDataFLD) = DateW
      ShowDate (LastDataFLD + 1)
      Call ComputeDIF
   End If
   If LastDataFLD = 10 Then
      If ldiftype = 0 Then ldiftype = 1
      Call ComputeDate(ldiftype)
   End If
   If (LastDataFLD > 10) Then
      Select Case LastDataFLD
         Case 11: X = 1: difyyyy = Val(DataField(LastDataFLD))
         Case 12: X = 1: difmm = Val(DataField(LastDataFLD))
         Case 13: X = 1: difdd = Val(DataField(LastDataFLD))
         Case 14: X = 2: difyyyy = Val(DataField(LastDataFLD))
         Case 15: X = 2: difddd = Val(DataField(LastDataFLD))
         Case 16: X = 3: difdays = Val(DataField(LastDataFLD))
         Case 17: X = 4: difyyyy = Val(DataField(LastDataFLD))
         Case 18: X = 4: difww = Val(DataField(LastDataFLD))
         Case 19: X = 4: difwday = Val(DataField(LastDataFLD))
      End Select
      ldiftype = X
      Call ComputeDate(X)
   End If
   Call DataField_GotFocus(LastDataFLD)
End Sub

Private Sub cmdHide_Click()
frameCalendar1.Visible = True
frameCalendar2.Visible = False
frmHEIGHT = 3360
frmFHSWDCLC.Height = frmHEIGHT
End Sub

Private Sub cmdRETURN_Click()
Unload frmFHSWDCLC
End Sub

Private Sub cmdSHOW_Click(Index)
'frameCalendar2.TOP = frameCalendar1.TOP
DateW = DateS(Index): CalDate = Index
DateC = DateW
'frameCalendar1.Visible = False
frameCalendar2.Visible = True
Call BuildCalendar
'frmFHSWDCLC.Height = frameCalendar2.TOP + frameCalendar2.Height + 450
'frmFHSWDCLC.Height = 5460
End Sub



Private Sub DataField_GotFocus(Index As Integer)
DataField(Index).SelStart = 0
If (Index < 2) Then
   If (DFM < 6) Then
      XDATE = DataField(Index)
      DataField(Index).SelLength = 1
      Else
      DataField(Index).SelLength = Len(DataField(Index).TEXT)
   End If
   Else
   DataField(Index).SelLength = Len(DataField(Index).TEXT)
End If
LastDataFLD = Index
End Sub


Private Sub DataField_KeyPress(Index As Integer, KeyAscii As Integer)
SS = DataField(Index).SelStart
SL = DataField(Index).SelLength
If (KeyAscii = 8) And (SS > 0) Then SS = SS - 1: GoTo NullKey
If (Index < 2) Then
   If DFM = 6 Then
      If (KeyAscii < Asc("0")) Or (KeyAscii > Asc("9")) Then
         GoTo NullKey
         Else
         Exit Sub
      End If
   End If
   If Len(DataField(Index)) < Len(DateMask) Then
      DataField(Index) = XDATE
   End If
   If SS = Len(DateMask) Then SL = 0: GoTo NullKey
   SL = 1
   If (KeyAscii = 32) Then
      KeyAscii = 0
      Else
      If (KeyAscii < 48) Then GoTo NullKey
      If (KeyAscii > 57) And (KeyAscii < 65) Then GoTo NullKey
      If (KeyAscii > 90) And (KeyAscii < 97) Then GoTo NullKey
      If (KeyAscii > 122) And (KeyAscii < 160) Then GoTo NullKey
      If (KeyAscii > 96) And (KeyAscii < 123) Then KeyAscii = KeyAscii - 32
   End If
   If KeyAscii > 0 Then
      dma = Asc(MID$(DateMask, SS + 1, 1))
      If (dma = Asc("M")) And (BMO > 0) Then
         If (KeyAscii < 65) Or (KeyAscii > 90) Then GoTo NullKey
         Else
         If (KeyAscii < 48) Or (KeyAscii > 57) Then GoTo NullKey
      End If
      XDATE = DataField(Index)
      Mid$(XDATE, SS + 1, 1) = Chr$(KeyAscii)
      If (BMO > 0) And (((SS - BMO) > -2) And ((SS - BMO) < 2)) Then
         MCHK$ = MID$(XDATE, BMO, SS - BMO + 2)
         FOUND = 1: X = 1
         While (X < 13) And (Not FOUND)
            FOUND = (MCHK$ = MID$(mon_table(X), 1, SS - BMO + 2))
            If Not FOUND Then X = X + 1
         Wend
         If FOUND Then
            Mid$(XDATE, BMO, 3) = mon_table(X)
            Else
            GoTo NullKey
         End If
      End If
      DataField(Index) = XDATE
   End If
   SS = SS + 1
   If (SS < Len(DateMask)) And (DFM < 6) Then
      dma = Asc(MID$(DateMask, SS + 1, 1))
      While ((dma < 48) Or ((dma > 57) And (dma < 65)) Or ((dma > 90) And (dma < 160)))
         SS = SS + 1
         If SS < Len(DateMask) Then
            dma = Asc(MID$(DateMask, SS + 1, 1))
            Else
            dma = 48
         End If
      Wend
   End If
   GoTo NullKey
End If
If (Index = 10) Then
   If Chr$(KeyAscii) = "-" Then
      difsign = -1
      Else
      difsign = 1: KeyAscii = Asc("+")
   End If
   SS = 0: SL = 1
End If
If (Index > 10) And (Index < 20) Then
   If (KeyAscii < Asc("0")) Or (KeyAscii > Asc("9")) Then
      GoTo NullKey
   End If
End If
DataField(Index).SelStart = SS
DataField(Index).SelLength = SL
Exit Sub
NullKey: KeyAscii = 0
   DataField(Index).SelStart = SS
End Sub

Private Sub DataField_LostFocus(Index As Integer)
If Index < 2 Then
   XDATE = DataField(Index)
   If DFM = 6 Then
      If Val(XDATE) < minADAT Then
         XDATE = LTrim$(Str$(minADAT))
         Else
         If Val(XDATE) > maxADAT Then
            XDATE = LTrim$(Str$(maxADAT))
         End If
      End If
   End If
   Call UnformatDate(XDATE, -1)
   DateS(Index) = DateW
End If
If (Index > 10) Then
   Select Case Index
      Case 11: difyyyy = Val(DataField(Index))
      Case 12: difmm = Val(DataField(Index))
      Case 13: difdd = Val(DataField(Index))
      Case 14: difyyyy = Val(DataField(Index))
      Case 15: difddd = Val(DataField(Index))
      Case 16: difdays = Val(DataField(Index))
      Case 17: difyyyy = Val(DataField(Index))
      Case 18: difww = Val(DataField(Index))
      Case 19: difwday = Val(DataField(Index))
   End Select
End If
End Sub

Private Sub Form_Load()
If FormLoaded(2) And SPSEFORM Then
   frmFHSWSpse.Enabled = False
   Else
   If FormLoaded(2) And UpdtForm Then
      frmFHSWUPDT.Enabled = False
   End If
End If
FormLoaded(2) = FormLoaded(2) Or DateCalcForm
' Set Default Form Location on Screen
frmHEIGHT = 3360
frmWIDTH = frmFHSWDCLC.WIDTH
t1 = GetSetting("FHSWIN", "frmDCLC", "TOP", 0)
If t1 = 0 Then
   t00 = mnuFHSWMain.TOP: w00 = mnuFHSWMain.WIDTH
   l00 = mnuFHSWMain.Left
   t1 = (t00 + 1435)
   l1 = l00 + (w00 - frmWIDTH) / 2: If l1 < 0 Then l1 = 1
   Else
   t1 = 1440: l1 = 1440
End If
frmFHSWDCLC.TOP = GetSetting("FHSWIN", "frmDCLC", "Top", t1)
frmFHSWDCLC.Left = GetSetting("FHSWIN", "frmDCLC", "Left", l1)
frmFHSWDCLC.Height = frmHEIGHT
frameCalendar2.Visible = False
Call SetCalendarType(0, 0)
XDATE = DATE$: Call UnformatDate(XDATE, 0)
DateS(0) = DateW: Call BuildCalendar
DateS(1) = DateW: ComputeDIF
cboCalendarType.ListIndex = GetSetting("FHSWIN", "frmDCLC", "CalendarType", 0)
Call adat2yyyyddd: Call ddd2mmdd
cboDateFormat.ListIndex = GetSetting("FHSWIN", "frmDCLC", "DateFormat", 0)
Call SetCalendarType(cboCalendarType.ListIndex, cboDateFormat.ListIndex)
End Sub


Private Sub Form_Resize()
frmFHSWDCLC.Height = frmHEIGHT
frmFHSWDCLC.WIDTH = frmWIDTH
End Sub


Private Sub Form_Unload(Cancel As Integer)
FormLoaded(2) = FormLoaded(2) And (255 - DateCalcForm)
SaveSetting "FHSWIN", "frmDCLC", "Top", frmFHSWDCLC.TOP
SaveSetting "FHSWIN", "frmDCLC", "Left", frmFHSWDCLC.Left
SaveSetting "FHSWIN", "frmDCLC", "DateFormat", cboDateFormat.ListIndex
SaveSetting "FHSWIN", "frmDCLC", "CalendarType", cboCalendarType.ListIndex
If FormLoaded(2) And SpouseForm Then
   frmFHSWSpse.Enabled = True
   frmFHSWSpse.SetFocus
   Else
   If FormLoaded(2) And UpdtForm Then
      frmFHSWUPDT.Enabled = True
      frmFHSWUPDT.SetFocus
   Else
      mnuFHSWMain.SetFocus
   End If
End If
End Sub




Private Sub gridCalendar_Click()
If gridCalendar.TEXT <> "" Then
   DateW = DateC
   DateW.DD = gridCalendar.TEXT
   Call mmdd2ddd
   Call yyyyddd2adat
   DateC = DateW
   DateS(CalDate) = DateC
   Call ShowDate(CalDate + 1)
   Call ComputeDIF
End If
End Sub

Private Sub gridCalendar_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

