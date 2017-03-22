VERSION 4.00
Begin VB.Form frmFHSWUPgm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FHSWUTIL - Utility Program"
   ClientHeight    =   5940
   ClientLeft      =   1740
   ClientTop       =   2040
   ClientWidth     =   5925
   Height          =   6345
   KeyPreview      =   -1  'True
   Left            =   1680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   Top             =   1695
   Width           =   6045
   Begin VB.CommandButton CMDButton 
      Caption         =   "Save"
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   143
      Top             =   3960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame DDStats 
      Caption         =   "Duplicate Data Summary"
      Height          =   2775
      Left            =   1440
      TabIndex        =   111
      Top             =   3120
      Visible         =   0   'False
      Width           =   4455
      Begin VB.TextBox DDStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   275
         Index           =   2
         Left            =   2640
         MultiLine       =   -1  'True
         TabIndex        =   142
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox DDStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   275
         Index           =   0
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   141
         Top             =   1560
         Width           =   735
      End
      Begin VB.CheckBox DDCkBox 
         Caption         =   "Check3"
         Height          =   255
         Index           =   7
         Left            =   2880
         TabIndex        =   136
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox DDCkBox 
         Caption         =   "Check3"
         Height          =   255
         Index           =   6
         Left            =   840
         TabIndex        =   135
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox DDCkBox 
         Caption         =   "Check3"
         Height          =   255
         Index           =   5
         Left            =   2880
         TabIndex        =   134
         Top             =   720
         Width           =   255
      End
      Begin VB.CheckBox DDCkBox 
         Caption         =   "Check3"
         Height          =   255
         Index           =   4
         Left            =   1680
         TabIndex        =   133
         Top             =   720
         Width           =   255
      End
      Begin VB.CheckBox DDCkBox 
         Caption         =   "Check3"
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   132
         Top             =   720
         Width           =   255
      End
      Begin VB.CheckBox DDCkBox 
         Caption         =   "Check3"
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   131
         Top             =   480
         Width           =   255
      End
      Begin VB.CheckBox DDCkBox 
         Caption         =   "Check2"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   130
         Top             =   480
         Width           =   255
      End
      Begin VB.CheckBox DDCkBox 
         Caption         =   "Check1"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   129
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox DDStat 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1200
         TabIndex        =   114
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox DDStat 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   3240
         TabIndex        =   113
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox DDStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   275
         Index           =   1
         Left            =   2880
         MultiLine       =   -1  'True
         TabIndex        =   112
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label DDSLbl 
         Caption         =   "Minimum Matched Group Size:"
         Height          =   255
         Index           =   15
         Left            =   360
         TabIndex        =   140
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label DDSLbl 
         Caption         =   "Group #2:"
         Height          =   255
         Index           =   14
         Left            =   2040
         TabIndex        =   139
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label DDSLbl 
         Caption         =   "Group #1:"
         Height          =   255
         Index           =   13
         Left            =   360
         TabIndex        =   138
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label DDSLbl 
         Caption         =   "Comparison Group Size"
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
         Index           =   12
         Left            =   120
         TabIndex        =   137
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label DDSLbl 
         Caption         =   "Compare Fields"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   15
         Index           =   11
         Left            =   0
         TabIndex        =   128
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label DDSLbl 
         Caption         =   "Death"
         Height          =   255
         Index           =   10
         Left            =   3240
         TabIndex        =   127
         Top             =   960
         Width           =   735
      End
      Begin VB.Label DDSLbl 
         Caption         =   "Birth"
         Height          =   255
         Index           =   9
         Left            =   1200
         TabIndex        =   126
         Top             =   960
         Width           =   615
      End
      Begin VB.Label DDSLbl 
         Caption         =   "Place:"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   125
         Top             =   960
         Width           =   735
      End
      Begin VB.Label DDSLbl 
         Caption         =   "Death"
         Height          =   255
         Index           =   7
         Left            =   3240
         TabIndex        =   124
         Top             =   720
         Width           =   735
      End
      Begin VB.Label DDSLbl 
         Caption         =   "Marriage"
         Height          =   255
         Index           =   6
         Left            =   2040
         TabIndex        =   123
         Top             =   720
         Width           =   735
      End
      Begin VB.Label DDSLbl 
         Caption         =   "Birth"
         Height          =   255
         Index           =   5
         Left            =   1200
         TabIndex        =   122
         Top             =   720
         Width           =   615
      End
      Begin VB.Label DDSLbl 
         Caption         =   "Date:"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   121
         Top             =   720
         Width           =   735
      End
      Begin VB.Label DDSLbl 
         Caption         =   "Sex Code"
         Height          =   255
         Index           =   3
         Left            =   3240
         TabIndex        =   120
         Top             =   480
         Width           =   855
      End
      Begin VB.Label DDSLbl 
         Caption         =   "Given Name"
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   119
         Top             =   480
         Width           =   975
      End
      Begin VB.Label DDSLbl 
         Caption         =   "Surname"
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   118
         Top             =   480
         Width           =   735
      End
      Begin VB.Label DDSLbl 
         Caption         =   "Matches:"
         Height          =   255
         Index           =   16
         Left            =   360
         TabIndex        =   117
         Top             =   2310
         Width           =   855
      End
      Begin VB.Label DDSLbl 
         Caption         =   "Groups:"
         Height          =   255
         Index           =   17
         Left            =   2400
         TabIndex        =   116
         Top             =   2325
         Width           =   855
      End
      Begin VB.Label DDSLbl 
         Caption         =   "Compare Fields"
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
         Left            =   120
         TabIndex        =   115
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton CMDButton 
      Caption         =   "Counters"
      Enabled         =   0   'False
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   110
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Frame DVStats 
      Caption         =   "Counters"
      Height          =   2775
      Left            =   1440
      TabIndex        =   85
      Top             =   2640
      Visible         =   0   'False
      Width           =   4455
      Begin VB.TextBox DVStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   275
         Index           =   9
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   97
         Top             =   2110
         Width           =   735
      End
      Begin VB.TextBox DVStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   275
         Index           =   3
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   96
         Top             =   670
         Width           =   735
      End
      Begin VB.TextBox DVStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   275
         Index           =   10
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   95
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox DVStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   275
         Index           =   8
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   94
         Top             =   1870
         Width           =   735
      End
      Begin VB.TextBox DVStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   275
         Index           =   7
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   93
         Top             =   1630
         Width           =   735
      End
      Begin VB.TextBox DVStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   275
         Index           =   6
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   92
         Top             =   1390
         Width           =   735
      End
      Begin VB.TextBox DVStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   275
         Index           =   5
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   91
         Top             =   1150
         Width           =   735
      End
      Begin VB.TextBox DVStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   275
         Index           =   4
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   90
         Top             =   910
         Width           =   735
      End
      Begin VB.TextBox DVStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   275
         Index           =   2
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   89
         Top             =   430
         Width           =   735
      End
      Begin VB.TextBox DVStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   275
         Index           =   1
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   88
         Top             =   190
         Width           =   735
      End
      Begin VB.TextBox DVStat 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   12
         Left            =   3480
         TabIndex        =   87
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox DVStat 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   11
         Left            =   3480
         TabIndex        =   86
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label DVSLbl 
         Caption         =   "Sex Code of Parent"
         Height          =   255
         Index           =   9
         Left            =   1200
         TabIndex        =   109
         Top             =   2115
         Width           =   1695
      End
      Begin VB.Label DVSLbl 
         Caption         =   "TOTAL Errors"
         Height          =   255
         Index           =   10
         Left            =   1200
         TabIndex        =   108
         Top             =   2430
         Width           =   1335
      End
      Begin VB.Label DVSLbl 
         Caption         =   "Sex Code/Spouse Sex Code"
         Height          =   255
         Index           =   8
         Left            =   1200
         TabIndex        =   107
         Top             =   1875
         Width           =   2295
      End
      Begin VB.Label DVSLbl 
         Caption         =   "Father-Child Surname"
         Height          =   255
         Index           =   7
         Left            =   1200
         TabIndex        =   106
         Top             =   1635
         Width           =   2775
      End
      Begin VB.Label DVSLbl 
         Caption         =   "Blank Marriage Record"
         Height          =   255
         Index           =   6
         Left            =   1200
         TabIndex        =   105
         Top             =   1395
         Width           =   2895
      End
      Begin VB.Label DVSLbl 
         Caption         =   "Unmarried Parents"
         Height          =   255
         Index           =   5
         Left            =   1200
         TabIndex        =   104
         Top             =   1155
         Width           =   2895
      End
      Begin VB.Label DVSLbl 
         Caption         =   "Age at Birth of Child"
         Height          =   255
         Index           =   4
         Left            =   1200
         TabIndex        =   103
         Top             =   915
         Width           =   2895
      End
      Begin VB.Label DVSLbl 
         Caption         =   "Age at Marriage"
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   102
         Top             =   675
         Width           =   2895
      End
      Begin VB.Label DVSLbl 
         Caption         =   "Age/Spouse Age"
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   101
         Top             =   435
         Width           =   2895
      End
      Begin VB.Label DVSLbl 
         Caption         =   "Invalid Dates"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   100
         Top             =   195
         Width           =   2895
      End
      Begin VB.Label DVSLbl 
         Caption         =   "in Err"
         Height          =   255
         Index           =   12
         Left            =   3000
         TabIndex        =   99
         Top             =   2450
         Width           =   495
      End
      Begin VB.Label DVSLbl 
         Caption         =   "# IDs"
         Height          =   255
         Index           =   11
         Left            =   3000
         TabIndex        =   98
         Top             =   2190
         Width           =   495
      End
   End
   Begin VB.Frame PVStats 
      Caption         =   "Counters"
      Height          =   2775
      Left            =   1440
      TabIndex        =   60
      Top             =   2160
      Visible         =   0   'False
      Width           =   4455
      Begin VB.TextBox PVStat 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   11
         Left            =   3480
         TabIndex        =   82
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox PVStat 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   12
         Left            =   3480
         TabIndex        =   81
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox PVStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   275
         Index           =   1
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   70
         Top             =   190
         Width           =   735
      End
      Begin VB.TextBox PVStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   275
         Index           =   2
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   69
         Top             =   430
         Width           =   735
      End
      Begin VB.TextBox PVStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   275
         Index           =   4
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   68
         Top             =   910
         Width           =   735
      End
      Begin VB.TextBox PVStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   275
         Index           =   5
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   67
         Top             =   1150
         Width           =   735
      End
      Begin VB.TextBox PVStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   275
         Index           =   6
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   66
         Top             =   1390
         Width           =   735
      End
      Begin VB.TextBox PVStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   275
         Index           =   7
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   65
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox PVStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   275
         Index           =   8
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   64
         Top             =   1870
         Width           =   735
      End
      Begin VB.TextBox PVStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   275
         Index           =   10
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   63
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox PVStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   275
         Index           =   3
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   62
         Top             =   670
         Width           =   735
      End
      Begin VB.TextBox PVStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   275
         Index           =   9
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   61
         Top             =   2110
         Width           =   735
      End
      Begin VB.Label PVSLbl 
         Caption         =   "# IDs"
         Height          =   255
         Index           =   11
         Left            =   3000
         TabIndex        =   84
         Top             =   2190
         Width           =   495
      End
      Begin VB.Label PVSLbl 
         Caption         =   "in Err"
         Height          =   255
         Index           =   12
         Left            =   3000
         TabIndex        =   83
         Top             =   2450
         Width           =   495
      End
      Begin VB.Label PVSLbl 
         Caption         =   "Name Record Errors"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   80
         Top             =   195
         Width           =   2895
      End
      Begin VB.Label PVSLbl 
         Caption         =   "Spouse Record Errors"
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   79
         Top             =   435
         Width           =   2895
      End
      Begin VB.Label PVSLbl 
         Caption         =   "Address Record Errors"
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   78
         Top             =   675
         Width           =   2895
      End
      Begin VB.Label PVSLbl 
         Caption         =   "Comment Record Errors"
         Height          =   255
         Index           =   4
         Left            =   1200
         TabIndex        =   77
         Top             =   915
         Width           =   2895
      End
      Begin VB.Label PVSLbl 
         Caption         =   "Event Record Errors"
         Height          =   255
         Index           =   5
         Left            =   1200
         TabIndex        =   76
         Top             =   1155
         Width           =   2895
      End
      Begin VB.Label PVSLbl 
         Caption         =   "Miscellaneous Record Errors"
         Height          =   255
         Index           =   6
         Left            =   1200
         TabIndex        =   75
         Top             =   1395
         Width           =   2895
      End
      Begin VB.Label PVSLbl 
         Caption         =   "Duplicate References"
         Height          =   255
         Index           =   7
         Left            =   1200
         TabIndex        =   74
         Top             =   1635
         Width           =   2775
      End
      Begin VB.Label PVSLbl 
         Caption         =   "Unreferenced Records"
         Height          =   255
         Index           =   8
         Left            =   1200
         TabIndex        =   73
         Top             =   1875
         Width           =   2175
      End
      Begin VB.Label PVSLbl 
         Caption         =   "TOTAL Errors"
         Height          =   255
         Index           =   10
         Left            =   1200
         TabIndex        =   72
         Top             =   2430
         Width           =   1335
      End
      Begin VB.Label PVSLbl 
         Caption         =   "Other Errors"
         Height          =   255
         Index           =   9
         Left            =   1200
         TabIndex        =   71
         Top             =   2115
         Width           =   1695
      End
   End
   Begin VB.Frame GCStats 
      Caption         =   "Counters"
      Height          =   2775
      Left            =   6120
      TabIndex        =   33
      Top             =   120
      Visible         =   0   'False
      Width           =   4455
      Begin VB.TextBox GCStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   8
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   59
         Top             =   2110
         Width           =   735
      End
      Begin VB.TextBox GCStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   2
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   58
         Top             =   670
         Width           =   735
      End
      Begin VB.TextBox GCStat 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   12
         Left            =   3480
         TabIndex        =   56
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox GCStat 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   11
         Left            =   3480
         TabIndex        =   54
         Top             =   2175
         Width           =   855
      End
      Begin VB.TextBox GCStat 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   10
         Left            =   3480
         TabIndex        =   52
         Top             =   1935
         Width           =   855
      End
      Begin VB.TextBox GCStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   9
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   50
         Top             =   2365
         Width           =   735
      End
      Begin VB.TextBox GCStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   7
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   49
         Top             =   1870
         Width           =   735
      End
      Begin VB.TextBox GCStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   6
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   48
         Top             =   1630
         Width           =   735
      End
      Begin VB.TextBox GCStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   5
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   47
         Top             =   1390
         Width           =   735
      End
      Begin VB.TextBox GCStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   4
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   46
         Top             =   1150
         Width           =   735
      End
      Begin VB.TextBox GCStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   3
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   45
         Top             =   910
         Width           =   735
      End
      Begin VB.TextBox GCStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   1
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   44
         Top             =   430
         Width           =   735
      End
      Begin VB.TextBox GCStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   0
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   43
         Top             =   190
         Width           =   735
      End
      Begin VB.Label GCSLbl 
         Caption         =   "Address Records"
         Height          =   255
         Index           =   8
         Left            =   1200
         TabIndex        =   57
         Top             =   2115
         Width           =   1695
      End
      Begin VB.Label GCSLbl 
         Caption         =   "Lines"
         Height          =   255
         Index           =   32
         Left            =   3000
         TabIndex        =   55
         Top             =   2445
         Width           =   495
      End
      Begin VB.Label GCSLbl 
         Caption         =   "FAMI"
         Height          =   255
         Index           =   31
         Left            =   3000
         TabIndex        =   53
         Top             =   2205
         Width           =   495
      End
      Begin VB.Label GCSLbl 
         Caption         =   "INDI"
         Height          =   255
         Index           =   30
         Left            =   3000
         TabIndex        =   51
         Top             =   1965
         Width           =   495
      End
      Begin VB.Label GCSLbl 
         Caption         =   "Comment Records"
         Height          =   255
         Index           =   9
         Left            =   1200
         TabIndex        =   42
         Top             =   2355
         Width           =   1695
      End
      Begin VB.Label GCSLbl 
         Caption         =   "Event Records"
         Height          =   255
         Index           =   7
         Left            =   1200
         TabIndex        =   41
         Top             =   1875
         Width           =   1575
      End
      Begin VB.Label GCSLbl 
         Caption         =   "Medical Records"
         Height          =   255
         Index           =   6
         Left            =   1200
         TabIndex        =   40
         Top             =   1635
         Width           =   1575
      End
      Begin VB.Label GCSLbl 
         Caption         =   "Military Records"
         Height          =   255
         Index           =   5
         Left            =   1200
         TabIndex        =   39
         Top             =   1395
         Width           =   2895
      End
      Begin VB.Label GCSLbl 
         Caption         =   "Work Records"
         Height          =   255
         Index           =   4
         Left            =   1200
         TabIndex        =   38
         Top             =   1155
         Width           =   2895
      End
      Begin VB.Label GCSLbl 
         Caption         =   "Education Records"
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   37
         Top             =   915
         Width           =   2895
      End
      Begin VB.Label GCSLbl 
         Caption         =   "Spouse Records"
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   36
         Top             =   675
         Width           =   2895
      End
      Begin VB.Label GCSLbl 
         Caption         =   "Places  (Birth, Marriage, Death)"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   35
         Top             =   435
         Width           =   2895
      End
      Begin VB.Label GCSLbl 
         Caption         =   "Name Records"
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   34
         Top             =   195
         Width           =   2895
      End
   End
   Begin VB.Frame MMStats 
      Caption         =   "Counters"
      Height          =   2055
      Left            =   6120
      TabIndex        =   18
      Top             =   4440
      Visible         =   0   'False
      Width           =   4455
      Begin VB.TextBox MMStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   275
         Index           =   3
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   32
         Top             =   970
         Width           =   735
      End
      Begin VB.TextBox MMStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   275
         Index           =   6
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   28
         Top             =   1690
         Width           =   735
      End
      Begin VB.TextBox MMStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   275
         Index           =   5
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   27
         Top             =   1450
         Width           =   735
      End
      Begin VB.TextBox MMStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   275
         Index           =   4
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   1210
         Width           =   735
      End
      Begin VB.TextBox MMStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   275
         Index           =   2
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   23
         Top             =   730
         Width           =   735
      End
      Begin VB.TextBox MMStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   275
         Index           =   1
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   490
         Width           =   735
      End
      Begin VB.TextBox MMStat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   275
         Index           =   0
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   250
         Width           =   735
      End
      Begin VB.Label MMSLbl 
         Caption         =   "Residence Address"
         Height          =   255
         Index           =   6
         Left            =   1200
         TabIndex        =   31
         Top             =   1680
         Width           =   3135
      End
      Begin VB.Label MMSLbl 
         Caption         =   "Places    (Birth/Marriage/Death)"
         Height          =   255
         Index           =   5
         Left            =   1200
         TabIndex        =   30
         Top             =   1440
         Width           =   3135
      End
      Begin VB.Label MMSLbl 
         Caption         =   "Dates     (Birth/Marriage/Death)"
         Height          =   255
         Index           =   4
         Left            =   1200
         TabIndex        =   29
         Top             =   1200
         Width           =   3015
      End
      Begin VB.Label MMSLbl 
         Caption         =   "Marriages"
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   25
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label MMSLbl 
         Caption         =   "Name Records"
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   24
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label MMSLbl 
         Caption         =   "Combined Records"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   22
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label MMSLbl 
         Caption         =   "MailMerge File Records"
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   20
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.CommandButton cmdCHANGE 
      Caption         =   "Change"
      Height          =   315
      Left            =   5160
      TabIndex        =   17
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton CMDButton 
      Caption         =   "Return"
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   16
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox PRTDesc 
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
      Height          =   315
      Left            =   2400
      TabIndex        =   14
      Top             =   600
      Width           =   2535
   End
   Begin VB.ComboBox PRTSetup 
      Height          =   315
      Left            =   1080
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdRFView 
      Caption         =   "View"
      Height          =   285
      Left            =   5160
      TabIndex        =   12
      Top             =   1080
      Width           =   700
   End
   Begin VB.CommandButton cmdCANCEL 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   285
      Left            =   5160
      TabIndex        =   11
      Top             =   120
      Width           =   700
   End
   Begin VB.CommandButton cmdRFBrowse 
      Caption         =   "Browse"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4200
      TabIndex        =   6
      Top             =   1080
      Width           =   750
   End
   Begin VB.ComboBox PRTDest 
      Height          =   315
      ItemData        =   "FHSWUPGM.frx":0000
      Left            =   1080
      List            =   "FHSWUPGM.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox RFDest 
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
      Height          =   315
      Left            =   2400
      TabIndex        =   5
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton CMDButton 
      Caption         =   "Process"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton CMDButton 
      Caption         =   "Select"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton CMDButton 
      Caption         =   "Options"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox RPDesc 
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
      Height          =   315
      Left            =   2400
      TabIndex        =   3
      Top             =   120
      Width           =   2535
   End
   Begin VB.ComboBox UPSetup 
      Height          =   315
      ItemData        =   "FHSWUPGM.frx":0028
      Left            =   1080
      List            =   "FHSWUPGM.frx":002A
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Line Line5 
      X1              =   5040
      X2              =   5040
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Label lblPRTSetup 
      Caption         =   "Setup:"
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
      Left            =   240
      TabIndex        =   15
      Top             =   630
      Width           =   615
   End
   Begin VB.Line Line4 
      X1              =   50
      X2              =   5880
      Y1              =   1440
      Y2              =   1440
   End
   Begin MSGrid.Grid grdOPTIONS 
      Height          =   2310
      Left            =   1440
      TabIndex        =   10
      Top             =   1560
      Width           =   4395
      _Version        =   65536
      _ExtentX        =   7752
      _ExtentY        =   4075
      _StockProps     =   77
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS LineDraw"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Rows            =   10
      Cols            =   3
      ScrollBars      =   2
      HighLight       =   0   'False
      MouseIcon       =   "FHSWUPGM.frx":002C
   End
   Begin VB.Label lblPRTDest 
      Caption         =   "Print to:"
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
      Left            =   240
      TabIndex        =   1
      Tag             =   "p"
      Top             =   1110
      Width           =   855
   End
   Begin VB.Line Line3 
      X1              =   50
      X2              =   5880
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line2 
      X1              =   1320
      X2              =   1320
      Y1              =   1440
      Y2              =   4440
   End
   Begin VB.Line Line1 
      X1              =   50
      X2              =   5880
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label lblReport 
      Caption         =   "Report:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   150
      Width           =   735
   End
End
Attribute VB_Name = "frmFHSWUPgm"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
DefInt A-Z
Dim FileDEST$, xOPTIX(), cOPTIX, EDTIX
Dim OptVLth, OptRow, NewOptValu$, OldOptValu$
Dim xPRT As Variant, PRTCode As String
Dim CountersVisible, ProcessComplete
Private Sub BuildOptGrid()
grdOPTIONS.Enabled = False
scrollwid = 250
grdOPTIONS.Cols = 3
grdOPTIONS.Rows = RDFReport.OPTCNT + 1
grdOPTIONS.ColWidth(0) = 250
grdOPTIONS.ColWidth(1) = grdOPTIONS.WIDTH - 1180
grdOPTIONS.ColWidth(2) = 600
h = (grdOPTIONS.Height - 100) / 9
grdOPTIONS.Col = 0
For X = 0 To grdOPTIONS.Rows - 1
   grdOPTIONS.RowHeight(X) = h
   grdOPTIONS.Row = X
   If X > 0 Then
      grdOPTIONS.TEXT = X
   End If
Next X
grdOPTIONS.Row = 0
grdOPTIONS.Col = 1: grdOPTIONS.TEXT = "         Option"
grdOPTIONS.Col = 2: grdOPTIONS.TEXT = "Value"
Call UtilOptDisplay
End Sub

Private Sub BuildPRTSetup()
PRTSetup.Clear
For X = 1 To PDFHdr.SETUPS
    PRTSetup.AddItem PSETUPS(X).NAME
    PRTSetup.ItemData(PRTSetup.NewIndex) = X
Next X
PRTSetup.TEXT = PDFSetup.NAME
'Call PRTSetup_Click
End Sub
Private Sub BuildPRTDest()
Select Case RTrim$(PDFSetup.DEST)
   Case "SCRN:": PRTDest.ListIndex = 0
   Case "FILE": PRTDest.ListIndex = 2
   Case Else: PRTDest.ListIndex = 1
End Select
PRTDest_CLICK
End Sub

Private Sub OptCellINIT()
ChkOPT: ' Look for Real Option Entry
cOPTIX = xOPTIX(OptRow)
EDTIX = Asc(MID$(RPTOPT$(cOPTIX), 9, 1))
If EDTIX = 255 Then
   OptVLth = 0: OldOptValu$ = "": NewOptValu$ = ""
   Else
   If (21 < EDTIX) And (EDTIX < 26) Then
      OptVLth = EDTIX - 20
      Else
      OptVLth = 1
   End If
   OldOptValu$ = LTrim$(grdOPTIONS.TEXT)
   NewOptValu$ = ""
End If
grdOPTIONS.Row = OptRow
End Sub


Public Sub ShowStats(sTYPE$, X, Y$)
Select Case sTYPE$
   Case "SHOW": GoSub ShowCounters
   Case "HIDE": GoSub HideCounters
   Case "GEDCOM": GoSub ShowGCSTAT
   Case "MAILMERG": GoSub ShowMMSTAT
   Case "PTRVLDT": GoSub ShowPVSTAT
   Case "DATVLDT": GoSub ShowDVSTAT
   Case "DUPVLDT": GoSub ShowDDSTAT
End Select
Exit Sub

ShowCounters:
   Select Case RptCODE
      Case "GDEX": ' GEDCOM Export
                   GCStats.TOP = 1550
                   GCStats.Left = 1440
                   frmFHSWUPgm.Height = 4800
                   GCStats.Visible = True
                   GCStats.Refresh
      Case "GDIM": ' GEDCOM Import
                   GCStats.TOP = 1550
                   GCStats.Left = 1440
                   frmFHSWUPgm.Height = 4800
                   GCStats.Visible = True
                   GCStats.Refresh
      Case "MMEX": ' MailMerge Export
                   MMStats.TOP = 1550
                   MMStats.Left = 1440
                   MMStats.Visible = True
                   MMStats.Refresh
      Case "PVLD": ' Pointer Validation
                   PVStats.TOP = 1550
                   PVStats.Left = 1440
                   frmFHSWUPgm.Height = 4800
                   PVStats.Visible = True
                   PVStats.Refresh
                   CMDButton(5).Visible = True
      Case "DVLD": ' Data Validation
                   DVStats.TOP = 1550
                   DVStats.Left = 1440
                   frmFHSWUPgm.Height = 4800
                   DVStats.Visible = True
                   DVStats.Refresh
                   CMDButton(5).Visible = True
      Case "MDAT": ' Search for Duplicate Data
                   DDStats.TOP = 1550
                   DDStats.Left = 1440
                   frmFHSWUPgm.Height = 4800
                   DDStats.Visible = True
                   DDStats.Refresh
                   CMDButton(5).Visible = True
   End Select
   grdOPTIONS.Visible = False
   CountersVisible = True
   CMDButton(4).Enabled = True
   frmFHSWUPgm.Refresh
   Return
   
HideCounters:
   Select Case RptCODE
     Case "GDEX": GCStats.Visible = False
     Case "GDIM": GCStats.Visible = False
     Case "MMEX": MMStats.Visible = False
     Case "PVLD": PVStats.Visible = False
     Case "DVLD": DVStats.Visible = False
     Case "MDAT": DDStats.Visible = False
   End Select
   CMDButton(5).Visible = False
   frmFHSWUPgm.Height = 4350
   frmFHSWUPgm.Refresh
   CountersVisible = False
   grdOPTIONS.Visible = True
   ProcessComplete = True
   frmFHSWUPgm.Refresh
   Return
   
ShowGCSTAT:
  If X < 0 Then
     For Z = 0 To 12: GCStat(Z) = "": Next Z
     GoSub ShowCounters
     Else
     If X = 99 Then
        GoSub HideCounters
        Else
        GCStat(X - 1).TEXT = Y$
        GCStat(X - 1).Refresh
     End If
  End If
  Return
  
ShowMMSTAT:
  If X < 0 Then
     For Z = 0 To 6: MMStat(Z) = "": Next Z
     GoSub ShowCounters
     Else
     If X = 99 Then
        GoSub HideCounters
        Else
        MMStat(X).TEXT = Y$
        MMStat(X).Refresh
     End If
  End If
  Return
  
ShowPVSTAT:
  If X < 0 Then
     For Z = 1 To 12: PVStat(Z) = "": Next Z
     GoSub ShowCounters
     Else
     If X = 99 Then
        GoSub HideCounters
        Else
        PVStat(X).TEXT = Y$
        PVStat(X).Refresh
     End If
  End If
  Return

ShowDVSTAT:
  If X < 0 Then
     For Z = 1 To 12: DVStat(Z) = "": Next Z
     GoSub ShowCounters
     Else
     If X = 99 Then
        GoSub HideCounters
        Else
        DVStat(X).TEXT = Y$
        DVStat(X).Refresh
     End If
  End If
  Return

ShowDDSTAT:
  If X < 0 Then
     For Z = 0 To 7: DDCkBox(Z) = 0: Next Z
     For Z = 0 To 4: DDStat(Z) = "": Next Z
     GoSub ShowCounters
     Else
     If X = 99 Then
        GoSub HideCounters
        Else
        If X < 10 Then
           DDCkBox(X - 1) = 1
           DDCkBox(X - 1).Refresh
           Else
           DDStat(X - 10) = Y$
        End If
     End If
  End If
  Return

End Sub

Private Sub UtilOptUPDATE()
If (NewOptValu$ <> OldOptValu$) Then
   grdOPTIONS.Row = OptRow: grdOPTIONS.Col = 2
   Mid$(RPTOPT$(cOPTIX), 12, 5) = grdOPTIONS.TEXT
   Call RptOptEDIT(RPTOPT$(cOPTIX))
   If Not OKAY Then
      grdOPTIONS.TEXT = Right$("    " + OldOptValu$, 5)
      Mid$(RPTOPT$(cOPTIX), 12, 5) = grdOPTIONS.TEXT
      Else
      Call RptOptSET(RPTOPT$(cOPTIX), RptOPTION)
      RDFIndex(ENV.RPTIX).CHG = -1
      If RDFHdr.FType = "RDF" Then RDFHdr.FType = "CHG"
   End If
End If
End Sub


Private Sub UtilOptDisplay()
   OPTTable.FT = RDFIndex(ENV.RPTIX).OPTBGN
   OPTTable.XL = 0: OPTTable.SIZE = RDFReport.OPTCNT
   ReDim xOPTIX(RDFReport.OPTCNT)
   X = OPTTable.FT: xX$ = Space$(16)
   Y$ = Space$(4): ACHAR$ = "   "
   Z = 0
   While X > 0
      OPTTable.SIZE = OPTTable.SIZE + 1
      OPTTable.LT = X
      LSet xX$ = RPTOPT$(X)
      Call RptOptFIND(xX$)
      Call RptOptEXPAND(xX$)
      OPTIX = CVI(MID$(xX$, 7, 2))
      Z = Z + 1: grdOPTIONS.Row = Z
      xOPTIX(Z) = X
      grdOPTIONS.Col = 1
      If OPTIX > 0 Then
         grdOPTIONS.TEXT = MID$(TABDATA$(OPTIX), 10, 30)
         Else
         grdOPTIONS.TEXT = "Unknown Option"
      End If
      RPTOPT$(X) = xX$ + grdOPTIONS.TEXT
      grdOPTIONS.Col = 2
      grdOPTIONS.TEXT = MID$(xX$, 12, 5)
      X = OPTCHN(X).FWD
   Wend
   Call RptOptINIT
End Sub

Private Sub TurnOffOpts()
If grdOPTIONS.Enabled = True Then
   grdOPTIONS.Enabled = False
   grdOPTIONS.HighLight = False
   If NewOptValu$ <> "" Then UtilOptUPDATE
   RPTOPTS(ENV.RPTIX) = RptOPTION
   'cmdCANCEL.DEFAULT = True
   'cmdCANCEL.Cancel = True
   grdOPTIONS.TopRow = 1
   If CMDButton(2).Caption = "Show Table" Then
      CMDButton(2).Caption = "Build Table"
   End If
   CMDButton(4).Caption = "Return"
   CMDButton(4).DEFAULT = True
   CMDButton(0).Caption = "Options"
End If
End Sub

Private Sub TurnONOpts()
If grdOPTIONS.Enabled = False Then
   NewOptValu$ = "": OldOptValu$ = ""
   CMDButton(0).Caption = "Save Options"
   CMDButton(4).Caption = "End"
   CMDButton(4).DEFAULT = False
   Call RptOptINIT
   grdOPTIONS.TopRow = 1: grdOPTIONS.Row = 1
   grdOPTIONS.Col = 2: OptRow = grdOPTIONS.Row
   Call OptCellINIT
   grdOPTIONS.SelStartRow = OptRow: grdOPTIONS.SelStartCol = 2
   grdOPTIONS.SelEndRow = OptRow: grdOPTIONS.SelEndCol = 2
   grdOPTIONS.Enabled = True
   grdOPTIONS.HighLight = True
End If
End Sub


Private Sub CFCkBox_Click(Index As Integer)

End Sub

Private Sub CMDButton_Click(Index As Integer)
Select Case CMDButton(Index).Caption
   Case "Counters":
        TurnOffOpts
        If CountersVisible Then
           Call ShowStats("HIDE", 0, "")
           Else
           Call ShowStats("SHOW", 0, "")
        End If
   Case "Options":
        If CountersVisible Then Call ShowStats("HIDE", 0, "")
        TurnONOpts
        grdOPTIONS.Enabled = True
        grdOPTIONS.SetFocus
   Case "Save Options":
        TurnOffOpts
        Call RptOptSAVE
   Case "Process":
        TurnOffOpts
        Select Case UtilPGM
           Case "GDCM": Call UtilGDCM("PROCESS")
           Case "MLMG": Call UtilMLMG("PROCESS")
           Case "TTFL": Call UtilTTFL("PROCESS")
           Case "VLDT": Call UtilVLDT("PROCESS")
        End Select
   Case "Print":
        TurnOffOpts
        Select Case UtilPGM
           Case "GDCM": Call UtilGDCM("PRINT")
           Case "MLMG": Call UtilMLMG("PRINT")
           Case "TTFL": Call UtilTTFL("PRINT")
           Case "VLDT": Call UtilVLDT("PRINT")
        End Select
   Case "End":
        If grdOPTIONS.Enabled = True Then
           TurnOffOpts
           Else
           If CountersVisible Then
              Call ShowStats("HIDE", 0, "")
           End If
        End If
   Case "Save":
        If ERID(0) = 0 Then
           Call PutMSG("VLDTM022")
           Else
           Call SWRKCreate(ERID(), FF1Hdr.MAXID, 0)
           Call ShowStats("HIDE", 0, "")
        End If
   Case "WorkFile": TurnOffOpts
        Call frmFHSWRelW.BUILD(WRelType)
   Case "Select": TurnOffOpts
        frmFHSWSlct.Show 1
        If UtilPGM = "TTFL" Then
           CMDButton(2).Caption = "Build Table"
           CMDButton(2).Enabled = True
           CMDButton(3).Enabled = False
        End If
   Case "Build Table": TurnOffOpts
        Call UtilTTFL("BUILD")
        CMDButton(3).Enabled = True
   Case "Show Table": TurnOffOpts
        Call UtilTTFL("SHOW")
   Case "NodeFile": TurnOffOpts
        frmFHSWNdWk.Show
   Case "Return":
        Unload frmFHSWUPgm
   
End Select


End Sub


Private Sub cmdCANCEL_Click()
a = 27: CA$ = Chr$(27)
If grdOPTIONS.Enabled = True Then
   If (NewOptValu$ <> "") Then
      grdOPTIONS.TEXT = Right$("     " + OldOptValu$, 5)
      NewOptValu$ = ""
      Else
      TurnOffOpts
   End If
   Else
   If CountersVisible Then
      Call ShowStats("HIDE", 0, "")
      Else
      Unload frmFHSWUPgm
   End If
End If
End Sub

Private Sub cmdCHANGE_Click()
frmFHSWPrinter.Show
End Sub




Private Sub cmdRFBrowse_Click()
TurnOffOpts
xX$ = PRT.DEST: PRT.DEST = FileDEST$
Call frmFHSWFile.LOCATE("REPORT", "FULL")
If OKAY Then
   RFDest.TEXT = PRT.DEST
   FileDEST$ = PRT.DEST
   Else
   PRT.DEST = xX$
End If
If RTrim$(xX$) <> RTrim$(PRT.DEST) Then PAGE.COUNT = 0
RFDest.SetFocus
End Sub


Private Sub cmdRFView_Click()
Call frmFHSWView.ReportFIL
End Sub



Private Sub Form_Load()
FormLoaded(1) = FormLoaded(1) Or UPgmForm
' Set Default Form Location on Screen
frmHEIGHT = 4350
frmWIDTH = 6000
t1 = GetSetting("FHSWIN", "frmUPGM", "TOP", 0)
If t1 = 0 Then
   t00 = mnuFHSWMain.TOP: w00 = mnuFHSWMain.WIDTH
   l00 = mnuFHSWMain.Left
   t1 = (t00 + 1075)
   l1 = l00 + (w00 - frmWIDTH) / 2: If l1 < 0 Then l1 = 1
   Else
   t1 = 1440: l1 = 1440
End If
frmFHSWUPgm.TOP = GetSetting("FHSWIN", "frmUPGM", "Top", t1)
frmFHSWUPgm.Left = GetSetting("FHSWIN", "frmUPGM", "Left", l1)
If FormLoaded(1) And PrntForm Then
   cmdCHANGE.Enabled = False
End If
frmFHSWUPgm.Height = frmHEIGHT
frmFHSWUPgm.WIDTH = frmWIDTH
Select Case UtilPGM
   Case "GDCM": PgmLabel$ = "GEDCOM Import/Export Program"
   Case "INDX": PgmLabel$ = "INDEX File Creation"
   Case "MLMG": PgmLabel$ = "MailMerge Export Program"
   Case "TTFL": PgmLabel$ = "Tiny Tafel Program"
   Case "VLDT": PgmLabel$ = "File Validation Program"
End Select
If PgmLabel$ <> "" Then
   frmFHSWUPgm.Caption = "FHSWUTIL - " + PgmLabel$
End If
FileDEST$ = FT$(14): PRTCode = "FILE"
NewOptValu$ = ""
Call BuildPRTSetup
End Sub

Private Sub Form_Unload(Cancel As Integer)
For X = 0 To (UPSetup.ListCount - 1)
   Call RPTFree(UPSetup.List(X))
Next X
UtilPGM = "": RptCODE = "": ProcessComplete = 0
FormLoaded(1) = FormLoaded(1) And (255 - UPgmForm)
SaveSetting "FHSWIN", "frmUPGM", "Top", frmFHSWUPgm.TOP
SaveSetting "FHSWIN", "frmUPGM", "Left", frmFHSWUPgm.Left
mnuFHSWMain.SetFocus
End Sub



Private Sub grdOPTIONS_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
a = KeyAscii: CA$ = Chr$(a)
Select Case a
  Case 13: ' Enter key
   If (NewOptValu$ <> "") Then
      UtilOptUPDATE
      If Not OKAY Then
         grdOPTIONS.Row = OptRow
         grdOPTIONS_SelChange
         Exit Sub
      End If
   End If
   If grdOPTIONS.Row < RDFReport.OPTCNT Then
      OptRow = OptRow + 1
      If Not (grdOPTIONS.RowIsVisible(OptRow)) Then
         grdOPTIONS.TopRow = grdOPTIONS.TopRow + 1
      End If
      grdOPTIONS.Row = OptRow
   End If
  Case 27: ' Escape key
      cmdCANCEL_Click
  Case Else:
   If OptVLth = 0 Then Exit Sub
   If (EDTIX > 20) Then
      If ((KeyAscii < 48) Or (KeyAscii > 57)) And (CA$ <> "+") Then
         Call ErrBEEP(0): Exit Sub
      End If
   End If
   If Len(NewOptValu$) >= OptVLth Then
      Call ErrBEEP(0): Exit Sub
   End If
   NewOptValu$ = NewOptValu$ + CA$
   grdOPTIONS.TEXT = Right$("    " + NewOptValu$, 5)
   Exit Sub
End Select
grdOPTIONS_SelChange
End Sub


Private Sub grdOPTIONS_RowColChange()
  If (grdOPTIONS.Row < 1) And (grdOPTIONS.Enabled) Then
     grdOPTIONS.Row = 1
  End If
  grdOPTIONS_SelChange
End Sub


Private Sub grdOPTIONS_SelChange()
If grdOPTIONS.Enabled = False Then Exit Sub
grdOPTIONS.Enabled = False
If OptRow = 0 Then
   OptRow = 1
   grdOPTIONS.Row = 1
End If
NextRow = grdOPTIONS.Row
grdOPTIONS.Col = 2
If NewOptValu$ <> "" Then Call UtilOptUPDATE
OptRow = NextRow
grdOPTIONS.Row = OptRow
Call OptCellINIT
grdOPTIONS.SelStartRow = OptRow
grdOPTIONS.SelStartCol = 2
grdOPTIONS.SelEndRow = OptRow
grdOPTIONS.SelEndCol = 2
grdOPTIONS.Enabled = True
grdOPTIONS.SetFocus
End Sub


Private Sub PRTDest_CLICK()
TurnOffOpts
RFDest.Enabled = False: RFDest.TEXT = ""
cmdRFBrowse.Enabled = False
cmdRFView.Enabled = False
PRT.OPT = PRTDest.ListIndex + 1
Select Case PRTDest.ListIndex
   Case 0: PRT.DEST = "$PRTSCRN.RPT"
   Case 1: PRT.DEST = PDFSetup.DEST
   Case 2: PRT.DEST = FileDEST$
           cmdRFBrowse.Enabled = True
           cmdRFView.Enabled = True
           RFDest.Enabled = True
           RFDest.TEXT = PRT.DEST
End Select
End Sub

Private Sub PRTSetup_Click()
TurnOffOpts
If PRTSetup.TEXT <> PRTCode Then
   If PRTSetup.TEXT = "" Then PRTSetup.TEXT = PRTCode
   Call PDFGetSet(PRTSetup.TEXT)
   If OKAY Then
      PRTDesc.TEXT = PDFSetup.DESC
      PRTCode = PDFSetup.NAME
      Call BuildPRTDest
      Else
      PRTDesc.TEXT = ""
   End If
End If
End Sub


Private Sub RFDest_Change()
TurnOffOpts
If RFDest.Enabled Then
   X = RFDest.SelStart
   If InStr(RFDest.TEXT, ".") = 0 Then
      RFDest.TEXT = RTrim$(RFDest.TEXT) + ".RPT"
   End If
   RFDest.SelStart = X
End If
End Sub

Private Sub RFDest_GotFocus()
RFDest.SelStart = 0
RFDest.SelLength = Len(RFDest.TEXT)
End Sub


Private Sub RFDest_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub RFDest_LostFocus()
FileDEST$ = RFDest.TEXT
End Sub



Public Sub UPSetup_Click()
TurnOffOpts
If CountersVisible Then Call ShowStats("HIDE", 0, "")
If UPSetup.TEXT = "" Then
   X = 0: FOUND = 0
   While (X < UPSetup.ListCount) And (Not FOUND)
      FOUND = (RptCODE = UPSetup.List(X))
      If Not FOUND Then X = X + 1
   Wend
   If FOUND Then UPSetup.ListIndex = X
End If
RptCODE = UPSetup.TEXT
If UPSetup.ListIndex = 0 Then
   CMDButton(2).Caption = "Process"
   CMDButton(3).Enabled = ProcessComplete
   Else
   CMDButton(2).Caption = "Print"
   CMDButton(3).Enabled = False
End If
Call RptLOAD(RptCODE)
If OKAY Then
   RPDesc.TEXT = RDFReport.NAME
   Call BuildOptGrid
   Else
   RPDesc.TEXT = ""
End If
End Sub



