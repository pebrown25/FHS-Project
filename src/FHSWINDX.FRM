VERSION 4.00
Begin VB.Form frmFHSWINDX 
   Caption         =   "FHSWINDX - Index File Creation"
   ClientHeight    =   9645
   ClientLeft      =   14010
   ClientTop       =   900
   ClientWidth     =   4845
   Height          =   10050
   Left            =   13950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   Top             =   555
   Width           =   4965
   Begin VB.Frame SortProgressFrame 
      Height          =   3495
      Left            =   0
      TabIndex        =   65
      Top             =   6120
      Width           =   4815
      Begin VB.TextBox SPValue 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   17
         Left            =   3120
         TabIndex        =   96
         Top             =   3090
         Width           =   1095
      End
      Begin VB.TextBox SPValue 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   16
         Left            =   1080
         TabIndex        =   94
         Top             =   3090
         Width           =   1095
      End
      Begin VB.TextBox SPValue 
         Height          =   285
         Index           =   15
         Left            =   4080
         TabIndex        =   92
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox SPValue 
         Height          =   285
         Index           =   14
         Left            =   3240
         TabIndex        =   91
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox SPValue 
         Height          =   285
         Index           =   13
         Left            =   2400
         TabIndex        =   90
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox SPValue 
         Height          =   285
         Index           =   12
         Left            =   1560
         TabIndex        =   89
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox SPValue 
         Height          =   285
         Index           =   11
         Left            =   720
         TabIndex        =   88
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox SPValue 
         Height          =   285
         Index           =   10
         Left            =   3480
         TabIndex        =   86
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox SPValue 
         Height          =   285
         Index           =   9
         Left            =   2520
         TabIndex        =   85
         Top             =   1890
         Width           =   735
      End
      Begin VB.TextBox SPValue 
         Height          =   285
         Index           =   8
         Left            =   1560
         TabIndex        =   84
         Top             =   1890
         Width           =   735
      End
      Begin VB.TextBox SPValue 
         Height          =   285
         Index           =   7
         Left            =   840
         TabIndex        =   83
         Top             =   1890
         Width           =   495
      End
      Begin VB.TextBox SPValue 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   6
         Left            =   3840
         TabIndex        =   80
         Top             =   1170
         Width           =   615
      End
      Begin VB.TextBox SPValue 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   5
         Left            =   2400
         TabIndex        =   78
         Top             =   1170
         Width           =   615
      End
      Begin VB.TextBox SPValue 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   4
         Left            =   1080
         TabIndex        =   76
         Top             =   1170
         Width           =   495
      End
      Begin VB.TextBox SPValue 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   3
         Left            =   3120
         TabIndex        =   74
         Top             =   810
         Width           =   975
      End
      Begin VB.TextBox SPValue 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   73
         Top             =   810
         Width           =   975
      End
      Begin VB.TextBox SPValue 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   2880
         TabIndex        =   70
         Top             =   450
         Width           =   1095
      End
      Begin VB.TextBox SPValue 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   840
         TabIndex        =   68
         Top             =   450
         Width           =   735
      End
      Begin VB.Label SPLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "End"
         Height          =   255
         Index           =   12
         Left            =   2280
         TabIndex        =   95
         Top             =   3150
         Width           =   735
      End
      Begin VB.Label SPLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Begin:"
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   93
         Top             =   3120
         Width           =   735
      End
      Begin VB.Line Line2 
         X1              =   240
         X2              =   4680
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Label SPLabel 
         Caption         =   "Time:      IN              Sort            Work         Merge          OUT"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   87
         Top             =   2280
         Width           =   4455
      End
      Begin VB.Label SPLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "SS="
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   82
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label SPLabel 
         Caption         =   "I/O:                             IN                Work            Out"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   81
         Top             =   1680
         Width           =   4455
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   4680
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label SPLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "XSCT="
         Height          =   255
         Index           =   7
         Left            =   3000
         TabIndex        =   79
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label SPLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "SSCT="
         Height          =   255
         Index           =   6
         Left            =   1680
         TabIndex        =   77
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label SPLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "#SS="
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   75
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label SPLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "on Disk="
         Height          =   255
         Index           =   4
         Left            =   2160
         TabIndex        =   72
         Top             =   840
         Width           =   855
      End
      Begin VB.Label SPLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "in Memory="
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   71
         Top             =   840
         Width           =   975
      End
      Begin VB.Label SPLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Data (Bytes)="
         Height          =   255
         Index           =   2
         Left            =   1680
         TabIndex        =   69
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label SPLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "#ID's="
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   67
         Top             =   480
         Width           =   615
      End
      Begin VB.Label SPLabel 
         Alignment       =   2  'Center
         Caption         =   "Sort Progress Report"
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
         Left            =   1440
         TabIndex        =   66
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.Frame Q2Frame 
      Height          =   495
      Left            =   0
      TabIndex        =   63
      Top             =   5640
      Width           =   4815
      Begin VB.OptionButton Q2OptButton 
         Caption         =   "Yes"
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   31
         Top             =   150
         Width           =   735
      End
      Begin VB.OptionButton Q2OptButton 
         Caption         =   "No"
         Height          =   195
         Index           =   1
         Left            =   3000
         TabIndex        =   32
         Top             =   180
         Width           =   735
      End
      Begin VB.Label QFLabel 
         Caption         =   "Use SORTSEQ Table?"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   64
         Top             =   165
         Width           =   2055
      End
   End
   Begin VB.Frame Q1Frame 
      Height          =   495
      Left            =   0
      TabIndex        =   61
      Top             =   5160
      Width           =   4815
      Begin VB.OptionButton Q1OptButton 
         Caption         =   "No"
         Height          =   195
         Index           =   1
         Left            =   3000
         TabIndex        =   30
         Top             =   180
         Width           =   735
      End
      Begin VB.OptionButton Q1OptButton 
         Caption         =   "Yes"
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   29
         Top             =   120
         Width           =   735
      End
      Begin VB.Label QFLabel 
         Caption         =   "Use Husband's Surname?"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   62
         Top             =   160
         Width           =   2055
      End
   End
   Begin VB.Frame IndexFileFrame 
      Enabled         =   0   'False
      Height          =   615
      Left            =   0
      TabIndex        =   56
      Top             =   600
      Width           =   4815
      Begin VB.TextBox IFFlag 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000012&
         Height          =   285
         Left            =   4560
         TabIndex        =   60
         Top             =   240
         Width           =   135
      End
      Begin VB.TextBox IFDate 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   3360
         TabIndex        =   59
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton CMDBrowse 
         Caption         =   "<"
         Height          =   285
         Left            =   2160
         TabIndex        =   38
         Top             =   195
         Width           =   375
      End
      Begin VB.TextBox IFName 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   37
         Top             =   195
         Width           =   1095
      End
      Begin VB.Label IFLabel 
         Caption         =   "Created:"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   58
         Top             =   240
         Width           =   735
      End
      Begin VB.Label IFLabel 
         Caption         =   "Index Name:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   57
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame DateOptsFrame 
      Height          =   735
      Left            =   0
      TabIndex        =   54
      Top             =   4440
      Width           =   4815
      Begin VB.OptionButton DateOButton 
         Caption         =   "D"
         Height          =   195
         Index           =   7
         Left            =   2520
         TabIndex        =   28
         Top             =   400
         Width           =   735
      End
      Begin VB.OptionButton DateOButton 
         Caption         =   "YMD"
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   21
         Top             =   160
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton DateOButton 
         Caption         =   "MDY"
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   22
         Top             =   160
         Width           =   735
      End
      Begin VB.OptionButton DateOButton 
         Caption         =   "DMY"
         Height          =   195
         Index           =   2
         Left            =   2520
         TabIndex        =   23
         Top             =   160
         Width           =   735
      End
      Begin VB.OptionButton DateOButton 
         Caption         =   "YM"
         Height          =   195
         Index           =   3
         Left            =   3360
         TabIndex        =   24
         Top             =   160
         Width           =   735
      End
      Begin VB.OptionButton DateOButton 
         Caption         =   "MD"
         Height          =   195
         Index           =   4
         Left            =   4080
         TabIndex        =   25
         Top             =   165
         Width           =   615
      End
      Begin VB.OptionButton DateOButton 
         Caption         =   "Y"
         Height          =   195
         Index           =   5
         Left            =   720
         TabIndex        =   26
         Top             =   400
         Width           =   615
      End
      Begin VB.OptionButton DateOButton 
         Caption         =   "M"
         Height          =   195
         Index           =   6
         Left            =   1680
         TabIndex        =   27
         Top             =   400
         Width           =   615
      End
      Begin VB.Label OFLabel 
         Caption         =   "Dates:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   55
         Top             =   160
         Width           =   495
      End
   End
   Begin VB.Frame NameOptsFrame 
      Height          =   495
      Left            =   0
      TabIndex        =   52
      Top             =   3960
      Width           =   4815
      Begin VB.OptionButton NameOButton 
         Caption         =   "SNDX+Text"
         Height          =   195
         Index           =   3
         Left            =   3480
         TabIndex        =   20
         Top             =   160
         Width           =   1215
      End
      Begin VB.OptionButton NameOButton 
         Caption         =   "SNDX"
         Height          =   195
         Index           =   2
         Left            =   2640
         TabIndex        =   19
         Top             =   160
         Width           =   855
      End
      Begin VB.OptionButton NameOButton 
         Caption         =   "Caps"
         Height          =   195
         Index           =   1
         Left            =   1920
         TabIndex        =   18
         Top             =   160
         Width           =   735
      End
      Begin VB.OptionButton NameOButton 
         Caption         =   "Unchanged"
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   17
         Top             =   160
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.Label OFLabel 
         Caption         =   "Text:"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   53
         Top             =   160
         Width           =   615
      End
   End
   Begin VB.Frame SFLDFrame 
      Height          =   2775
      Left            =   0
      TabIndex        =   41
      Top             =   1200
      Width           =   4815
      Begin VB.CheckBox SFSeq 
         Caption         =   "1"
         Enabled         =   0   'False
         Height          =   255
         Index           =   8
         Left            =   1800
         TabIndex        =   105
         Top             =   2400
         Value           =   2  'Grayed
         Width           =   615
      End
      Begin VB.CheckBox SFSeq 
         Height          =   255
         Index           =   7
         Left            =   1800
         TabIndex        =   104
         Top             =   2160
         Width           =   615
      End
      Begin VB.CheckBox SFSeq 
         Height          =   255
         Index           =   6
         Left            =   1800
         TabIndex        =   103
         Top             =   1920
         Width           =   615
      End
      Begin VB.CheckBox SFSeq 
         Height          =   255
         Index           =   5
         Left            =   1800
         TabIndex        =   102
         Top             =   1680
         Width           =   615
      End
      Begin VB.CheckBox SFSeq 
         Height          =   255
         Index           =   4
         Left            =   1800
         TabIndex        =   101
         Top             =   1440
         Width           =   615
      End
      Begin VB.CheckBox SFSeq 
         Height          =   255
         Index           =   3
         Left            =   1800
         TabIndex        =   100
         Top             =   1200
         Width           =   615
      End
      Begin VB.CheckBox SFSeq 
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   99
         Top             =   960
         Width           =   615
      End
      Begin VB.CheckBox SFSeq 
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   98
         Top             =   720
         Width           =   615
      End
      Begin VB.CheckBox SFSeq 
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   97
         Top             =   450
         Width           =   615
      End
      Begin VB.TextBox SFLth 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000012&
         Height          =   285
         Index           =   8
         Left            =   3360
         TabIndex        =   39
         Text            =   "2"
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox SFLth 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000012&
         Height          =   285
         Index           =   7
         Left            =   3360
         TabIndex        =   15
         Text            =   "1"
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox SFLth 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000012&
         Height          =   285
         Index           =   6
         Left            =   3360
         TabIndex        =   13
         Text            =   "12"
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox SFLth 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000012&
         Height          =   285
         Index           =   5
         Left            =   3360
         TabIndex        =   11
         Text            =   "12"
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox SFLth 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000012&
         Height          =   285
         Index           =   4
         Left            =   3360
         TabIndex        =   9
         Text            =   "4"
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox SFLth 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000012&
         Height          =   285
         Index           =   3
         Left            =   3360
         TabIndex        =   7
         Text            =   "4"
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox SFLth 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000012&
         Height          =   285
         Index           =   2
         Left            =   3360
         TabIndex        =   5
         Text            =   "4"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox SFLth 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000012&
         Height          =   285
         Index           =   1
         Left            =   3360
         TabIndex        =   3
         Text            =   "15"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox SFLth 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000012&
         Height          =   285
         Index           =   0
         Left            =   3360
         TabIndex        =   1
         Text            =   "10"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox SFOrd 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000012&
         Height          =   285
         Index           =   8
         Left            =   2640
         TabIndex        =   16
         Text            =   "A"
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox SFOrd 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000012&
         Height          =   285
         Index           =   7
         Left            =   2640
         TabIndex        =   14
         Text            =   "A"
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox SFOrd 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000012&
         Height          =   285
         Index           =   6
         Left            =   2640
         TabIndex        =   12
         Text            =   "A"
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox SFOrd 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000012&
         Height          =   285
         Index           =   5
         Left            =   2640
         TabIndex        =   10
         Text            =   "A"
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox SFOrd 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000012&
         Height          =   285
         Index           =   4
         Left            =   2640
         TabIndex        =   8
         Text            =   "A"
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox SFOrd 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000012&
         Height          =   285
         Index           =   3
         Left            =   2640
         TabIndex        =   6
         Text            =   "A"
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox SFOrd 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000012&
         Height          =   285
         Index           =   2
         Left            =   2640
         TabIndex        =   4
         Text            =   "A"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox SFOrd 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000012&
         Height          =   285
         Index           =   0
         Left            =   2640
         TabIndex        =   0
         Text            =   "A"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox SFOrd 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000012&
         Height          =   285
         Index           =   1
         Left            =   2640
         TabIndex        =   2
         Text            =   "A"
         Top             =   720
         Width           =   495
      End
      Begin VB.Label SFLabel 
         Caption         =   "ID Number"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   51
         Top             =   2400
         Width           =   1395
      End
      Begin VB.Label SFLabel 
         Caption         =   "Sex Code"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   50
         Top             =   2160
         Width           =   1395
      End
      Begin VB.Label SFLabel 
         Caption         =   "Death Place"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   49
         Top             =   1920
         Width           =   1395
      End
      Begin VB.Label SFLabel 
         Caption         =   "Birth Place"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   48
         Top             =   1680
         Width           =   1395
      End
      Begin VB.Label SFLabel 
         Caption         =   "Death Date"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   47
         Top             =   1440
         Width           =   1395
      End
      Begin VB.Label SFLabel 
         Caption         =   "Marriage Date"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   46
         Top             =   1200
         Width           =   1395
      End
      Begin VB.Label SFLabel 
         Caption         =   "Birth Date"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   45
         Top             =   960
         Width           =   1395
      End
      Begin VB.Label SFLabel 
         Caption         =   "Given Name"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   44
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label SFLabel 
         Appearance      =   0  'Flat
         Caption         =   "Surname"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   43
         Top             =   480
         Width           =   1395
      End
      Begin VB.Label SFLabel 
         Caption         =   "  Sort Index Field               Seq          A/D          Lth"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   42
         Top             =   120
         Width           =   3855
      End
   End
   Begin VB.Frame CMDFrame 
      Height          =   615
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   4815
      Begin VB.CommandButton CMDButton 
         Caption         =   "Print"
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   35
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CMDButton 
         Caption         =   "Return"
         Height          =   255
         Index           =   3
         Left            =   3840
         TabIndex        =   36
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CMDButton 
         Caption         =   "Create"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   34
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CMDButton 
         Caption         =   "Set Options"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmFHSWINDX"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
DefInt A-Z
Dim SST1$, hNAME$, UseSEQTable, Return2Menu
Dim HOLDITEM As Variant, xFT


Private Sub SortOptsSET()
SFCNT = 1
For X = 0 To 8
    SFOrd(X) = SNGLKEY$(14)
    Select Case X
       Case 0: SFSeq(X).VALUE = 1
               SFLth(X) = 10
       Case 1: SFSeq(X).VALUE = 1
               SFLth(X) = 15
       Case 2: SFSeq(X).VALUE = 1
               SFLth(X) = 4
       Case 3: SFLth(X) = 4
       Case 4: SFLth(X) = 4
       Case 5: SFLth(X) = 12
       Case 6: SFLth(X) = 12
       Case 7: SFLth(X) = 1
       Case 8: SFSeq(X).Caption = SFCNT
               SFLth(X) = 2
    End Select
Next X
ENV.SortOPTS = 0
NameOButton(0).VALUE = True
DateOButton(0).VALUE = True
Q1OptButton(1) = True
Q2OptButton(1) = True
Call SortOptsGET
End Sub
Private Sub IFNameSET()
X = InStr(FT$(4), ".")
If X > 0 Then
    IFName.TEXT = Left$(FT$(4), X - 1)
    Else
    IFName.TEXT = FT$(4)
End If
End Sub

Private Sub SortOptsGET()
SFCNT = 0
For X = 0 To 8
    Y = Val(SFSeq(X).Caption)
    SFCNT = SFCNT - (Y > 0)
    SFLD(X + 1) = 256 * Y
    If SFOrd(X) = SNGLKEY$(15) Then
       SFLD(X + 1) = SFLD(X + 1) + 128
    End If
    Z = Val(SFLth(X))
    If Z > 35 Then Z = 35
    SFLD(X + 1) = SFLD(X + 1) + Z
Next X
ENV.SortOPTS = 0
For X = 0 To 3
    If NameOButton(X) Then ENV.SortOPTS = X
Next X
For X = 0 To 7
    If DateOButton(X) Then ENV.SortOPTS = ENV.SortOPTS + 4 * X
Next X
ENV.HUSBSURNM = Q1OptButton(0)
UseSEQTable = Q2OptButton(0)
ENV.SortOPTS = ENV.SortOPTS + 4 * xo - 32 * ENV.HUSBSURNM - 64 * UseSEQTable
SFLDFrame.Enabled = False
NameOptsFrame.Enabled = False
DateOptsFrame.Enabled = False
Q1Frame.Enabled = False
q2frame.Enabled = False
End Sub





Private Sub cmdBROWSE_Click()
hNAME$ = FT$(4)
Call frmFHSWFile.LOCATE("INDEX", "FILE")
If OKAY And (hNAME$ <> FT$(4)) Then Call IFNameSET
End Sub

Private Sub CMDButton_Click(Index As Integer)
Select Case CMDButton(Index).Caption
   Case "Set Options":
        IndexFileFrame.Enabled = False
        CMDButton(0).Visible = False
        CMDButton(1).Caption = "End"
        CMDButton(2).Caption = "Cancel"
        CMDButton(3).Visible = False
        SFLDFrame.Enabled = True
        NameOptsFrame.Enabled = True
        DateOptsFrame.Enabled = True
        Q1Frame.Enabled = True
        q2frame.Enabled = True
        SFSeq(0).SetFocus
   Case "Create": Call INDEXBuild
        Call SortOptsSHOW
   Case "Print":
        frmFHSWINDX.Enabled = False
        RptPGM = "LIST"
        RptCODE = "INDX"
        Call RptLIST("INIT")
   Case "End":
        Call SortOptsGET
        Call SortOptsSHOW
   Case "Cancel":
        Call SortOptsSHOW
   Case "Return": Unload frmFHSWINDX
End Select
End Sub

Private Sub Form_Load()
   FormLoaded(3) = FormLoaded(3) Or IndexForm
   PN$ = " FAMILY HISTORY SYSTEM - FHSWINDX - Index File Generation Program"
   CY$ = " (C) Copyright 1985-2008 by Phillip E. Brown"
   SST1$ = ""
 ' Set Default Form Location on Screen
   frmHEIGHT = 6540
   frmWIDTH = frmFHSWINDX.WIDTH
   t1 = GetSetting("FHSWIN", "frmINDX", "TOP", 0)
   If t1 = 0 Then
      t00 = mnuFHSWMain.TOP: w00 = mnuFHSWMain.WIDTH
      l00 = mnuFHSWMain.Left
      t1 = (t00 + 1075)
      l1 = l00 + (w00 - frmWIDTH) / 2: If l1 < 0 Then l1 = 1
      Else
      t1 = 1440: l1 = 1440
   End If
   frmFHSWINDX.TOP = GetSetting("FHSWIN", "frmINDX", "Top", t1)
   frmFHSWINDX.Left = GetSetting("FHSWIN", "frmINDX", "Left", l1)
   frmFHSWINDX.Height = frmHEIGHT
   SortProgressFrame.Visible = False
   Call SortOptsSET
   Call IFNameSET
End Sub

Private Sub INDEXBuild()
'     SHARED FMT001, FMT030, FMT031, ShowSTATS, SST1$, SFLD()
1000 ' Create Index File
     Call PgmPREP("FHSINDX")
     Call SortOptsGET
     textsort = ((SFLD(1) \ 256) > 0) Or ((SFLD(2) \ 256) > 0) Or ((SFLD(6) \ 256) > 0) Or ((SFLD(7) \ 256) > 0)
     If textsort Then
        textsort = ENV.SortOPTS And 3
        UseSEQTable = ENV.SortOPTS And 64
        If UseSEQTable Then Call SortSeqTable("BUILD", SST1$, "", 1)
     End If
     DateSORT = ((SFLD(3) \ 256) > 0) Or ((SFLD(4) \ 256) > 0) Or ((SFLD(5) \ 256) > 0)
     If DateSORT Then
        DateSORT = ((ENV.SortOPTS \ 4) And 7)
     End If
     ENV.SortSTAT = False
     If ENV.SortSTAT Then
        SFLDFrame.Visible = False
        NameOptsFrame.Visible = False
        DateOptsFrame.Visible = False
        Q1Frame.Visible = False
        q2frame.Visible = False
        SortProgressFrame.TOP = 1200
        SortProgressFrame.Visible = True
        frmFHSWINDX.Height = 5115
     End If
     Call FAMOpen("READ"): If Not OKAY Then GoTo 1090 ' Open Name file
     If FF1Hdr.MAXID = 0 Then
        Call PutMSG("INDXM001")
        R$ = Chr$(27): GoTo 1090
     End If
     GoSub 1800                      ' Allocate NDXID(),SEQDATA$()
     If Not OKAY Then GoTo 1090
     'CURFMT = FMT030: FMT = FMTTAB(CURFMT)
     'Call WinSWITCH(2): Call WinCLR
     'Call FmtDREST
     'Call WinFORMAT(FMT030)
     'X = FMT.BGNFLD
     'While X > 0
     '   FFLD = FLDTAB(X)
     '   If FFLD.UTXT > 0 Then
     '      X$ = UTXT$(FFLD.UTXT)
     '      Select Case FFLD.NAME
     '         Case "IDCT": LSet X$ = Str$(FF1Hdr.MAXID)
     '         Case "TSZB": LSet X$ = Str$(SN1 * DL#)
     '         Case "MSZB": LSet X$ = Str$(SN2 * DL#)
     '         Case "WSZB": LSet X$ = Str$(SWSZ#)
     '         Case "SSEG": RSet X$ = LTrim$(Str$(N4))
     '         Case "SSCT": LSet X$ = Str$(SN2)
     '         Case "LSCT": LSet X$ = Str$(N3)
     '      End Select
     '      LSet UTXT$(FFLD.UTXT) = X$
     '   End If
     '   X = FLDCHN(X).FWD
     'Wend
     'Call WinDisplayDATA(0)
     'SortFMT = FMT031
     'CURFMT = SortFMT: FMT = FMTTAB(CURFMT)
     'Call FmtDREST
     'Call FmtFindFLD("TIM1")
     'If FOUND Then LSet UTXT$(FFLD.UTXT) = Time$
     If ENV.SortSTAT Then
        frmFHSWINDX.SPValue(0) = Str$(FF1Hdr.MAXID)
        frmFHSWINDX.SPValue(1) = Str$(SN1 * dl#)
        frmFHSWINDX.SPValue(2) = Str$(SN2 * dl#)
        frmFHSWINDX.SPValue(3) = Str$(swsz#)
        frmFHSWINDX.SPValue(4) = LTrim$(Str$(n4))
        frmFHSWINDX.SPValue(5) = Str$(SN2)
        frmFHSWINDX.SPValue(6) = Str$(n3)
        frmFHSWINDX.SPValue(16) = Time$
     End If
     'Call WinFORMAT(0)
     GoSub 1850             ' Get DATA, Sort, and Create .NDX File
     If Not OKAY Then GoTo 1090
     Call IFName_Change
     Call PutMSG("INDXM003")
     R$ = ""

1090 ' Return
     On Local Error GoTo 0
     If swsz# > 0 Then GoSub 5630
     Close: ENV.FFOPEN = 0: OKAY = (R$ <> Chr$(27))
     If SortProgressFrame.Visible Then
        SortProgressFrame.Visible = False
        SFLDFrame.Visible = True
        NameOptsFrame.Visible = True
        DateOptsFrame.Visible = True
        Q1Frame.Visible = True
        q2frame.Visible = True
        frmFHSWINDX.Height = 6540
     End If
     Exit Sub

1800 ' Initialize Arrays to be used in SORT
     SN1 = FF1Hdr.MAXID: SN2 = SN1
     SDL = 0: BDSORT = 0: BPSORT = 0
     Call SFLD2SX
     ReDim SFL(SFCNT)
     For NX = 1 To SFCNT
         X = SX(NX) \ 256: L = SX(NX) And 127
         If (X = 3) Or (X = 5) Then BDSORT = -1
         If (X = 6) Or (X = 7) Then BPSORT = -1
         If (X < 3) Or ((X > 5) And (X < 8)) Then
            Select Case textsort
               Case 2: L = 4
               Case 3: L = 4 + L
            End Select
            Else
            If (X > 2) And (X < 6) Then
               If (DateSORT > 5) Then
                  L = 1
                  Else
                  If (DateSORT = 5) Then
                     L = 2
                     Else
                     If (DateSORT > 2) Then L = 3
                  End If
               End If
            End If
         End If
         SFL(NX) = L: SDL = SDL + L
     Next NX
     SDATA$ = String$(SDL, 0)

1810 ' Build Sort Arrays
     On Local Error GoTo 1849
     ReDim NDXID(SN1), SEQDATA$(1)
     'Z# = FRE(SEQDATA$(1)) - 100
     Z# = 500000
     dl# = SDL: xN2# = Z# / (dl# + 10)
     If xN2# > 32767 Then SN2 = SN1 Else SN2 = xN2#: If SN2 > SN1 Then SN2 = SN1
     'X# = Z# - (DL# + 4) * SN2
     'IF X# < 5000 THEN X# = (Z# - 5000) / (DL# + 4): SN2 = X#
1820 ReDim SEQDATA$(SN2)
     Call INDEXSort("SETUP", NDXID(), SEQDATA$())
     If Not SortOK Then Return
     If SN1 = SN2 Then GoTo 1825
     swsz# = dl# * (SN1 - SN2): SysVAR.MVAR = LTrim$(Str$(swsz#))
     If ENV.SortSTAT Then Call SortStatACCUM
     SWD$ = WRKFPFX$: GoTo 1822
1821 Call PutMSG("INDXM004")
     R$ = CA$: If Asc(R$) = 27 Then GoTo 1840
     If CA$ < "A" Or CA$ > "Z" Then Beep: GoTo 1821
     SWD$ = CA$ + ":"
1822 GoSub 5600: If Not OKAY Then GoTo 1821
     On Local Error GoTo 1845
     If ENV.SortSTAT Then SortSTAT!(0) = Timer
1825 For ix = 1 To SN2
        SEQDATA$(ix) = SDATA$: HIINDEX = ix
        R$ = INKEY$
        If R$ <> "" Then
           If R$ = Chr$(27) Then
              ix = SN2
              Else
              If ENV.SortSTAT Then
                 Call SortStatSHOW
              End If
           End If
        End If
1829 Next ix
     If (R$ = Chr$(27)) Or (HIINDEX = SN2) Then GoTo 1840
     On Local Error GoTo 0
     GoTo 1820

1840 On Local Error GoTo 0
     MSZ# = SN2 * dl#: swsz# = (SN1 - SN2) * dl#
     n3 = SN1 Mod SN2: n4 = (SN1 \ SN2)
     OKAY = OKAY And (R$ <> Chr$(27))
     Return

1845 ' Return After Error
     Erase SEQDATA$: SysVAR.NVAR1 = HIINDEX
     ix = SN2: SN2 = HIINDEX - 30: OKAY = 0
     If SN2 < 0 Then
        Call ERRMessage("INDXM007", ERR, ERL)
        Resume 1849
     End If
     Resume 1829

1849 ' Error Allocating Arrays
     Call ERRMessage("INDXM006", ERR, ERL)
     OKAY = 0
     Resume 1840

1850 ' Put Name file sequence data in SEQDATA$() and $ORTWORK (if necessary)
     If BDSORT Or BPSORT Then
        BD$ = Space$(11): DD$ = BD$
        If BPSORT Then BPL$ = Space$(41) Else BPL$ = ""
        DPL$ = BPL$: AGE$ = ""
        DFM = Val(CONFIG.DFM)
     End If
     JX = 0: SortOK = -1: OKAY = -1
     SNDX$ = Space$(4): xDATA$ = Space$(34)
     For XIN = 1 To SN1
         NDXID(XIN) = XIN
         Call FF1GetRec(XIN): OFST = 0: SPREC = FF1Rec.SPOUSE
         If ENV.SortSTAT Then SortSTAT!(6) = SortSTAT!(6) + 1
         If BDSORT Or BPSORT Then
            Call GetBDEvents(BD$, BPL$, DD$, DPL$, AGE$)
            BMD = DATEMD(BD$, DFM): byr = DATEY(BD$, DFM)
            DMD = DATEMD(DD$, DFM): DYR = DATEY(DD$, DFM)
         End If
         For NX = 1 To SFCNT
             L = SFL(NX): ASCEND = ((SX(NX) And 128) = 0)
             Select Case SX(NX) \ 256
                Case 1: LSet xDATA$ = LTrim$(RTrim$(FF1Rec.SURNM))
                        If ENV.HUSBSURNM And (FF1Rec.SEX = FGENDR$(2)) And (FF1Rec.SPOUSE <> 0) And ((Asc(FF1Rec.STATUS2) And 128) = 0) Then
                           LSet xDATA$ = FMTNAME$(1): LSet xDATA$ = SURNM$
                        End If
                        Mid$(SDATA$, OFST + 1, L) = TextDATA$(xDATA$, L, ASCEND)
                Case 2: Mid$(SDATA$, OFST + 1, L) = TextDATA$(FF1Rec.GIVEN, L, ASCEND)
                Case 3: Mid$(SDATA$, OFST + 1, L) = DateDATA$(BMD, byr, ASCEND)
                Case 4: MMD = 0: MYR = 0
                        If FF1Rec.SPOUSE <> 0 Then
                           Call FF3GetRec(FF1Rec.SPOUSE)
                           If (F3BY$ <> MKI$(0)) Or (F3BMD$ <> MKI$(0)) Then
                              MMD = FF3SPOUSE.D1MD: MYR = FF3SPOUSE.D1Y
                           End If
                        End If
                        Mid$(SDATA$, OFST + 1, L) = DateDATA$(MMD, MYR, ASCEND)
                Case 5: Mid$(SDATA$, OFST + 1, L) = DateDATA$(DMD, DYR, ASCEND)
                Case 6: Mid$(SDATA$, OFST + 1, L) = TextDATA$(BPL$, L, ASCEND)
                Case 7: Mid$(SDATA$, OFST + 1, L) = TextDATA$(DPL$, L, ASCEND)
                Case 8: Mid$(SDATA$, OFST + 1, 1) = FF1Rec.SEX
                Case 9: Mid$(SDATA$, OFST + 1, 1) = FNBH$(XIN)
                        Mid$(SDATA$, OFST + 2, 1) = FNBL$(XIN)
             End Select
             If ((SX(NX) And 128) <> 0) Then
                For N = 1 To L
                    Mid$(SDATA$, OFST + N, 1) = Chr$(Asc(MID$(SDATA$, OFST + N, 1)) Xor 255)
                Next N
             End If
             OFST = OFST + L
         Next NX
         JX = JX + 1: LSet SEQDATA$(JX) = SDATA$: LSet SDATA$ = ""
         If (XIN = n3) Or (JX = SN2) Then
            SS = (XIN - n3) \ SN2
            If ENV.SortSTAT Then Call SortStatSHOW
            Call INDEXSort("SEGMENT", NDXID(), SEQDATA$())
            If SortOK Then
               If (SN1 <> SN2) And (SS < n4) Then
                  If ENV.SortSTAT Then
                     Call SortStatACCUM: HP3 = SortPHASE: SortPHASE = 3
                  End If
                  For X = 1 To JX: Put #5, , SEQDATA$(X): Next X
                  If ENV.SortSTAT Then
                     SortSTAT!(7) = SortSTAT!(7) + JX - 1
                     Call SortStatACCUM: SortPHASE = HP3
                  End If
               End If
               JX = 0: If SS < n4 Then SS = SS + 1
               Else
               XIN = SN1
            End If
         End If
         R$ = INKEY$
         If R$ <> "" Then
            SortOK = (R$ <> Chr$(27))
            If Not SortOK Then
               XIN = SN1
               Else
               If ENV.SortSTAT Then
                  Call SortStatSHOW
               End If
            End If
         End If
         If ENV.SortSTAT And (XIN Mod 100 = 0) Then Call SortStatSHOW
1876 Next XIN
     If ENV.SortSTAT Then Call SortStatSHOW: SortPHASE = 0
     OKAY = OKAY And SortOK
     If OKAY Then
        If (SN1 <> SN2) Then
           Call INDEXSort("MERGE", NDXID(), SEQDATA$())
           OKAY = SortOK
        End If
        If OKAY Then Call INDEXFile("CREATE")   ' Write Index Array to File
     End If
     OKAY = OKAY And SortOK
     Call INDEXSort("CLEANUP", NDXID(), SEQDATA$())
     Erase NDXID, SEQDATA$
     Return

5600 ' Open $ORTWORK File
     OKAY = 0: SWN$ = SWD$ + "$ORTWORK"
     GoSub 5630
     On Local Error GoTo 5610
     Open SWN$ For Binary Access Read Write As #5
     GoTo 5611
5610 SysVAR.FILENAME = SWN$
     Call PutMSG("XXXXM004")
     R$ = Chr$(27)
     Resume 5620
5611 On Local Error GoTo 5615
     XL# = (SN1 - SN2 - 1) * dl# + 1
     Put #5, XL#, SDATA$
     Seek #5, 1
     OKAY = -1: GoTo 5620
5615 Call PutMSG("INDXM008")
     R$ = Chr$(27): Resume 5616
5616 GoSub 5630
5620 On Local Error GoTo 0
     Return

5630 ' Kill $ORTWORK File
     On Local Error GoTo 5634
     Close #5: GoTo 5635
5634 Resume 5635
5635 On Local Error GoTo 5637
     Kill SWN$: GoTo 5638
5637 Resume 5638
5638 On Error GoTo 0
     Return
End Sub

  Private Function FNBH$(a)
  FNBH$ = Chr$((Abs(a) \ 256) And 255)
  End Function

  Private Function FNBL$(a)
  FNBL$ = Chr$(Abs(a) And 255)
  End Function


Private Sub SortOptsSHOW()
For X = 0 To 8
    If X < 8 Then SFSeq(X).VALUE = 0
    SFSeq(X).Caption = ""
Next X
SFCNT = 0
For X = 0 To 8
    Y = SFLD(X + 1) \ 256
    If (Y > 0) Then
       If X < 8 Then SFSeq(X).VALUE = 1
       SFSeq(X).Caption = Y
       Else
       SFSeq(X).VALUE = 0: SFSeq(X).Caption = ""
    End If
    If SFLD(X + 1) And 128 Then
       SFOrd(X) = SNGLKEY$(15)
       Else
       SFOrd(X) = SNGLKEY$(14)
    End If
    SFLth(X) = SFLD(X + 1) And 127
Next X
SFCNT = SFCNT + 1
NameOButton(ENV.SortOPTS And 3) = True
DateOButton((ENV.SortOPTS And 28) / 4) = True
Q1OptButton(0) = ((ENV.SortOPTS And 32) > 0)
Q2OptButton(0) = ((ENV.SortOPTS And 64) > 0)
CMDButton(0).Visible = True
CMDButton(1).Caption = "Create"
CMDButton(2).Caption = "Print"
CMDButton(3).Visible = True
SFLDFrame.Enabled = False
NameOptsFrame.Enabled = False
DateOptsFrame.Enabled = False
Q1Frame.Enabled = False
q2frame.Enabled = False
IndexFileFrame.Enabled = True
'IFName.SetFocus
End Sub

Private Function TextDATA$(X$, L%, ASCEND)
' SHARED SST1$
999   If ((LTrim$(X$) = "") And ASCEND) Then
      TextDATA$ = Chr$(255)
      Else
      Select Case (ENV.SortOPTS And 3)
         Case 0: Y$ = X$
         Case 1: Y$ = UCX$(X$)
         Case 2: TD$ = Space$(4): Call ComputeSNDX(X$, TD$)
                 Y$ = TD$
         Case 3: TD$ = Space$(4): Call ComputeSNDX(X$, TD$)
                 Y$ = TD$ + X$
      End Select
      If (ENV.SortOPTS And 64) > 0 Then
         Call SortSeqTable("TRAN", SST1$, Y$, L)
      End If
      TextDATA$ = Y$
   End If
End Function


Private Sub INDEXSort(OPT$, xAID(), SEQDATA$())
'   SHARED ShowSTATS
    ' OPT$ = "ALL"      Complete Sort of xAID()
    '      = "SETUP"    Allocate Required Arrays
    '      = "SEGMENT"  Sort SEQDATA$() <--> "SS" Segment of xAID()
    '      = "MERGE"    Merge Sort Segments and Create Index File
    '      = "CLEANUP"  Erase Arrays
    ' xAID() is the Array of ID's that is to be placed in Sorted Sequence
    ' SEQDATA$() is Array where DATA is kept in Memory
    ' SN1 = Number of Entries in xAID()
    ' SN2 = Number of Entries in SEQDATA$()
    ' SDL = Length of Each Entry in SEQDATA$()
    ' If SN1<>SN2 then using $ORTWORK file as File #5
    ' SS = # Sort Segment (OPT$="SEGMENT")

1200 ' Use Segmented SHELL Sort/Merge for Sequencing xAID()
     On Local Error GoTo 1207
     R$ = "": SortOK = -1
     sortdata$ = Space$(SDL)
     dl& = SDL

     ' Compute N3=Size of XTRA Segment, N4=#Sort Segments
     sb = 0: n3 = SN1 Mod SN2: n4 = SN1 \ SN2

     Select Case OPT$
        Case "SEGMENT": If SS = 0 Then NS = n3 Else NS = SN2: sb = n3 + (SS - 1) * SN2
                        GoSub 1250: GoTo 1209
        Case "MERGE": GoSub 1210: GoTo 1209
        Case "CLEANUP": GoTo 1208
     End Select
     ' Allocate Arrays
     If ENV.SortSTAT Then
        ReDim SortSTAT!(10)
        SortPHASE = 1: SortSTAT!(0) = Timer
     End If
     If SN1 <> SN2 Then
        ReDim MRGID(SN1), MRGBX(n4 + 1), MRGIX(n4 + 1), MRGSS(n4 + 1), MRGSD$(n4 + 1)
        For X = 0 To n4: MRGSD$(X) = sortdata$: Next X
     End If
     If OPT$ = "SETUP" Then GoTo 1209

     If n3 > 0 Then
        NS = n3: GoSub 1250: If Not SortOK Then GoTo 1209
        If OPT$ = "ALL" Then GoSub 1240: If Not SortOK Then GoTo 1209
     End If
     SS = 1: NS = SN2

1205 ' Perform SORT
     If SS <= n4 Then
        GoSub 1250: If Not SortOK Then SS = n4
     End If
     If (Not SortOK) Or (SN1 = SN2) Then GoTo 1209
     If SS < n4 Then
        GoSub 1240: If Not SortOK Then GoTo 1209
        SS = SS + 1: GoTo 1205
     End If
     GoSub 1210: GoTo 1208

1207 ' Error Routine
     XERR = ERR: XERL = ERL
     ReDim MRGSD$(1)
     Call ERRMessage("XXXXM071", XERR, XERL)
     SortOK = 0
     Resume 1208

1208 ' Cleanup at END of Sort
     If SN1 <> SN2 Then Erase MRGID, MRGBX, MRGIX, MRGSS, MRGSD$
     SortPHASE = 0: ENV.SortSTAT = 0: SortFMT = 0

1209 ' Exit from SORT Routine
     On Local Error GoTo 0
     Exit Sub

1210 ' Merge SORTed Segments
     If ENV.SortSTAT Then
        Call SortStatACCUM: HP4 = SortPHASE: SortPHASE = 4
        SortSTAT!(10) = 0: Call SortStatSHOW
     End If
     ' MRGID()=ID'S in Sorted Segment Sequence
     ' MRGSD$(SS)=SORTDATA Entry for Segment "SS"
     '    These are Sorted by rearranging MRGSS()
     ' MRGSS(..)=Segment Containing MRGSD$(..)
     ' MRGIX(SS)=Sequence# of Next Entry in Segment "SS"
     ' MRGBX(SS)=Sequence # of Last MRGID() Entry in Segment "SS"
     MRGSS(0) = 0:  MRGIX(0) = 1: MRGBX(0) = n3: SSL = n4
     For ix = 1 To SN1: MRGID(ix) = xAID(ix): Next ix
     For Z = 1 To n4
         MRGSS(Z) = Z
         MRGIX(Z) = MRGBX(Z - 1) + 1
         MRGBX(Z) = MRGBX(Z - 1) + SN2
     Next Z
     HiDATA$ = String$(SDL, 255): YDATA$ = HiDATA$
     For MS = 0 To n4
         MRGSD$(MS) = HiDATA$: GoSub 1235: TEOS = TEOS + EOS
         If Not SortOK Then MS = n4
     Next MS
     If Not SortOK Then Return
     GoSub 1230: SSL = SSL + TEOS
     For ix = 1 To SN1
         GoSub 1225
         If Not SortOK Then
            ix = SN1
            Else
            xAID(ix) = xSX
            If ENV.SortSTAT Then
               SortSTAT!(10) = SortSTAT!(10) + 1
               If (ix Mod 100 = 1) Then Call SortStatSHOW
            End If
         End If
     Next ix
     Erase MRGID, MRGBX, MRGIX, MRGSS, MRGSD$
     If ENV.SortSTAT Then Call SortStatSHOW: SortPHASE = HP4
     Return

1225 ' Find Next Index
     GoSub 1280
     If Not SortOK Then
        ix = SN1
        Else
        MS = MRGSS(0): X = MRGIX(MS): xSX = MRGID(X)
        MRGIX(MS) = MRGIX(MS) + 1
        GoSub 1235: GoSub 1230: SSL = SSL + EOS
     End If
     Return

1230 ' Bubble Sort MRGSD$() into order by MRGSS()
     JX = 0
     While JX < SSL
        kx = JX + 1: Y = MRGSS(kx): LSet YDATA$ = MRGSD$(Y)
        While kx > 0
           kx = kx - 1: X = MRGSS(kx)
           If MRGSD$(X) > YDATA$ Then
              'SWAP MRGSS (KX), MRGSS(KX + 1)
              xyz = MRGSS(kx + 1)
              MRGSS(kx + 1) = MRGSS(kx)
              MRGSS(kx) = xyz
              Else
              kx = 0
           End If
        Wend
        JX = JX + 1
     Wend
     Return

1235 ' Get Next Entry in Segment MS
     X = MRGIX(MS): EOS = (X > MRGBX(MS))
     If EOS Then
        LSet MRGSD$(MS) = HiDATA$
        Else
        If MS < n4 Then
           If ENV.SortSTAT Then
              Call SortStatACCUM: HP3 = SortPHASE: SortPHASE = 3
           End If
           RBA& = (MRGID(X) - 1) * dl& + 1
           Get #5, RBA&, sortdata$
           LSet MRGSD$(MS) = sortdata$
           If ENV.SortSTAT Then
              SortSTAT!(9) = SortSTAT!(9) + 1
              Call SortStatACCUM: SortPHASE = HP3
           End If
           Else
           X = MRGID(X) - MRGBX(MS - 1)
           LSet MRGSD$(MS) = SEQDATA$(X)
        End If
     End If
     Return

1240 ' Exchange SEQDATA$() for $ORTWORK data
     If ENV.SortSTAT Then
        Call SortStatACCUM: HP3 = SortPHASE: SortPHASE = 3
     End If
     For ix = sb + 1 To sb + NS
         RBA& = (ix - 1) * dl& + 1
         Get #5, RBA&, sortdata$
         'SWAP SORTDATA$, SEQDATA$(ix - SB)
         xyz = SEQDATA$(ix - sb)
         SEQDATA$(ix - sb) = sortdata$
         sortdata$ = xyz
         Put #5, RBA&, sortdata$
         If ENV.SortSTAT Then SortSTAT!(8) = SortSTAT!(8) + 2
         GoSub 1280: If Not SortOK Then ix = sb + NS
     Next ix
     sb = sb + NS
     If ENV.SortSTAT Then Call SortStatACCUM: SortPHASE = HP3
     Return

1250 ' Sort Segment of File
     If ENV.SortSTAT Then
        Call SortStatACCUM: HP2 = SortPHASE: SortPHASE = 2
     End If
     For ix = sb + 1 To sb + NS: xAID(ix) = xAID(ix) - sb: Next ix
     D = 2 ^ (Int(Log(NS + 1) / Log(2))) - 1
1254 For ix = sb + 1 To sb + D
         For JX = ix To sb + NS - D Step D
             X = xAID(JX): Y = xAID(JX + D)
             GoSub 1280
             If Not SortOK Then
                ix = sb + D: JX = sb + NS - D: D = 0: GoTo 1270
             End If
             If Not (SEQDATA$(X) > SEQDATA$(Y)) Then GoTo 1270
             Z = Y: kx = JX: xAID(kx + D) = X
1261         kx = kx - D
             If kx >= ix Then
                Y = xAID(kx)
                GoSub 1280
                If Not SortOK Then
                   ix = sb + D: JX = sb + NS - D: D = 0: GoTo 1270
                End If
                If SEQDATA$(Z) < SEQDATA$(Y) Then xAID(kx + D) = Y: GoTo 1261
             End If
             xAID(kx + D) = Z
1270     Next JX
     Next ix
     If ENV.SortSTAT Then Call SortStatSHOW
     If SortOK Then
        D = D \ 2: If D > 0 Then GoTo 1254
        For ix = sb + 1 To sb + NS: xAID(ix) = xAID(ix) + sb: Next ix
        If ENV.SortSTAT Then Call SortStatSHOW: SortPHASE = HP2
     End If
     Return

1280 ' Check for User Request
     R$ = INKEY$
     If R$ <> "" Then
        SortOK = (R$ <> Chr$(27))
        If SortOK And ENV.SortSTAT Then
           Call SortStatSHOW
        End If
     End If
     Return

End Sub

Private Sub SortSeqTable(OPT$, SST1$, xX$, L%)
    Select Case OPT$
       Case "BUILD": 'Create Sort Sequence Table
                      Call TableLOAD("SORTSEQ", T, 0)
                      SST1$ = String$(255, 0)
                      X = TABLE(T).FT
                      While X > 0
                         xX$ = TABDATA$(X)
                         If xX$ <> "" Then
                            Y = Asc(xX$): Z = Val(MID$(xX$, 2, 3))
                            If (-1 < Z) And (Z < 256) Then
                               Mid$(SST1$, Y, 1) = Chr$(Z)
                            End If
                         End If
                         X = TABDATACHN(X).FWD
                      Wend
                      Call TableFREE("SORTSEQ")
       Case "TRAN": 'Translate X$ using SST1$
                     X = Len(xX$): If L < X Then X = L
                     Y = 0: w = 0
                     While Y < X
                        Y = Y + 1: Z = Asc(MID$(xX$, Y, 1))
                        If Z > 0 Then
                           xZ$ = MID$(SST1$, Z, 1)
                           If Asc(xZ$) > 0 Then
                              w = w + 1: Mid$(xX$, w, 1) = xZ$
                           End If
                        End If
                     Wend
                     While w < X
                        w = w + 1: Mid$(xX$, w, 1) = Chr$(0)
                     Wend
    End Select

End Sub


Private Function DateDATA$(MD, CY, ASCEND)
         MD = Abs(MD): CY = Abs(CY)
         If (MD = 0) And (CY = 0) And ASCEND Then
            mm = 255: DD = 255: CC = 255: yy = 255
            Else
            mm = (MD \ 100): If mm > 255 Then mm = 255
            DD = MD - mm * 100: If DD > 255 Then DD = 255
            CC = (CY \ 100): If CC > 255 Then CC = 255
            yy = CY - 100 * CC: If yy > 255 Then yy = 255
         End If
         xX$ = Space$(4): X = 1
         Select Case ((ENV.SortOPTS \ 4) And 7)
            Case 0: Mid$(xX$, X, 1) = Chr$(CC)
                    Mid$(xX$, X + 1, 1) = Chr$(yy)
                    Mid$(xX$, X + 2, 1) = Chr$(mm)
                    Mid$(xX$, X + 3, 1) = Chr$(DD)
            Case 1: Mid$(xX$, X, 1) = Chr$(mm)
                    Mid$(xX$, X + 1, 1) = Chr$(DD)
                    Mid$(xX$, X + 2, 1) = Chr$(CC)
                    Mid$(xX$, X + 3, 1) = Chr$(yy)
            Case 2: Mid$(xX$, X, 1) = Chr$(DD)
                    Mid$(xX$, X + 1, 1) = Chr$(mm)
                    Mid$(xX$, X + 2, 1) = Chr$(CC)
                    Mid$(xX$, X + 3, 1) = Chr$(yy)
            Case 3: Mid$(xX$, X, 1) = Chr$(CC)
                    Mid$(xX$, X + 1, 1) = Chr$(yy)
                    Mid$(xX$, X + 2, 1) = Chr$(mm)
            Case 4: Mid$(xX$, X, 1) = Chr$(mm)
                    Mid$(xX$, X + 1, 1) = Chr$(DD)
            Case 5: Mid$(xX$, X, 1) = Chr$(CC)
                    Mid$(xX$, X + 1, 1) = Chr$(yy)
            Case 6: Mid$(xX$, X, 1) = Chr$(mm)
            Case 7: Mid$(xX$, X, 1) = Chr$(DD)
         End Select
         DateDATA$ = xX$
End Function


Private Sub Form_Unload(Cancel As Integer)
FormLoaded(3) = FormLoaded(3) And (255 - IndexForm)
SaveSetting "FHSWIN", "frmINDX", "Top", frmFHSWINDX.TOP
SaveSetting "FHSWIN", "frmINDX", "Left", frmFHSWINDX.Left
mnuFHSWMain.ChangeEnable
mnuFHSWMain.SetFocus
End Sub


Private Sub IFName_Change()
FT$(4) = IFName.TEXT + ".NDX"
Call INDEXFile("CHECKQ")
If OKAY Then
   IFDate.TEXT = XSCRNDate$(NDXHDR.UPDTMD, NDXHDR.UPDTY)
   If NDXHDR.NDXRECS < FF1Hdr.MAXID Then
      IFFlag.TEXT = "<"
      Else
      If NDXHDR.NDXRECS > FF1Hdr.MAXID Then
         IFFlag.TEXT = ">"
         Else
         IFFlag.TEXT = "="
      End If
   End If
   Call SX2SFLD
   Call SortOptsSHOW
   Else
   IFDate.TEXT = "": IFFlag.TEXT = ""
End If
End Sub








Private Sub SFLth_Change(Index As Integer)
If Len(SFLth(Index).TEXT) > 2 Then
   SFLth(Index).TEXT = Left$(SFLth(Index).TEXT, 2)
End If
End Sub

Private Sub SFLth_GotFocus(Index As Integer)
HOLDITEM = RTrim$(SFLth(Index).TEXT)
SFLth(Index).SelStart = 0
SFLth(Index).SelLength = 2
End Sub


Private Sub SFOrd_Change(Index As Integer)
X$ = UCX$(SFOrd(Index).TEXT)
If X$ <> SNGLKEY(14) And X$ <> SNGLKEY(15) Then
   SFOrd(Index).TEXT = HOLDITEM
   Else
   HOLDITEM = X$
   SFOrd(Index).TEXT = HOLDITEM
End If
SFOrd(Index).SelStart = 0
SFOrd(Index).SelLength = 1
   
End Sub

Private Sub SFOrd_GotFocus(Index As Integer)
SFOrd(Index).SelStart = 0
SFOrd(Index).SelLength = 1
HOLDITEM = SFOrd(Index).TEXT
End Sub

Private Sub SFSeq_Click(Index As Integer)
If SFSeq(Index).VALUE = 1 Then
   SFCNT = SFCNT + 1
   SFSeq(Index).Caption = SFCNT - 1
   SFSeq(8).Caption = SFCNT
   Else
   SFCNT = SFCNT - 1
   Y = Val(SFSeq(Index).Caption)
   SFSeq(Index).Caption = ""
   For X = 0 To 8
       If SFSeq(X).VALUE > 0 Then
          Z = Val(SFSeq(X).Caption)
          If Z >= Y Then
             Z = Z - 1: SFSeq(X).Caption = Z
          End If
       End If
   Next X
End If
End Sub


