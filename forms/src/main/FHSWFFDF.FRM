VERSION 4.00
Begin VB.Form frmFHSWFFDF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FHSWFFDF - Family File Definition"
   ClientHeight    =   2535
   ClientLeft      =   2280
   ClientTop       =   1590
   ClientWidth     =   6825
   Height          =   2940
   Left            =   2220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   Top             =   1245
   Width           =   6945
   Begin VB.Frame frameSTATS 
      Height          =   1815
      Left            =   120
      TabIndex        =   61
      Top             =   6720
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CommandButton cmdRESET 
         Caption         =   "Reset"
         Height          =   285
         Left            =   5880
         TabIndex        =   87
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtSTAT 
         Alignment       =   2  'Center
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
         Index           =   13
         Left            =   5280
         MultiLine       =   -1  'True
         TabIndex        =   86
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox txtSTAT 
         Alignment       =   2  'Center
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
         Index           =   12
         Left            =   3840
         MultiLine       =   -1  'True
         TabIndex        =   84
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox txtSTAT 
         Alignment       =   2  'Center
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
         Index           =   11
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   82
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtSTAT 
         Alignment       =   1  'Right Justify
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
         Index           =   10
         Left            =   5520
         MultiLine       =   -1  'True
         TabIndex        =   76
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtSTAT 
         Alignment       =   1  'Right Justify
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
         Index           =   6
         Left            =   5520
         MultiLine       =   -1  'True
         TabIndex        =   75
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtSTAT 
         Alignment       =   1  'Right Justify
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
         Index           =   9
         Left            =   4320
         MultiLine       =   -1  'True
         TabIndex        =   74
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtSTAT 
         Alignment       =   1  'Right Justify
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
         Index           =   5
         Left            =   4320
         MultiLine       =   -1  'True
         TabIndex        =   73
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtSTAT 
         Alignment       =   1  'Right Justify
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
         Index           =   2
         Left            =   4320
         MultiLine       =   -1  'True
         TabIndex        =   72
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtSTAT 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Index           =   8
         Left            =   2760
         MultiLine       =   -1  'True
         TabIndex        =   71
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtSTAT 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   2760
         MultiLine       =   -1  'True
         TabIndex        =   70
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtSTAT 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   2760
         MultiLine       =   -1  'True
         TabIndex        =   69
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtSTAT 
         Alignment       =   2  'Center
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
         Index           =   7
         Left            =   1800
         MultiLine       =   -1  'True
         TabIndex        =   68
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtSTAT 
         Alignment       =   2  'Center
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
         Index           =   3
         Left            =   1800
         MultiLine       =   -1  'True
         TabIndex        =   67
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtSTAT 
         Alignment       =   2  'Center
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
         Index           =   0
         Left            =   1800
         MultiLine       =   -1  'True
         TabIndex        =   66
         Top             =   480
         Width           =   375
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   6600
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label 
         Caption         =   "Count 2="
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
         Index           =   27
         Left            =   4440
         TabIndex        =   85
         Top             =   1470
         Width           =   855
      End
      Begin VB.Label Label 
         Caption         =   "Count 1="
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
         Index           =   26
         Left            =   3000
         TabIndex        =   83
         Top             =   1470
         Width           =   855
      End
      Begin VB.Label Label 
         Caption         =   "Change  Date:"
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
         Index           =   25
         Left            =   120
         TabIndex        =   81
         Top             =   1470
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Free Recs"
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
         Left            =   5400
         TabIndex        =   80
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label 
         Caption         =   "Records"
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
         Index           =   24
         Left            =   4320
         TabIndex        =   79
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label 
         Caption         =   "Last Update"
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
         Index           =   23
         Left            =   2880
         TabIndex        =   78
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label 
         Caption         =   "VER"
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
         Index           =   22
         Left            =   1800
         TabIndex        =   77
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label 
         Caption         =   "MiscInfo"
         Height          =   255
         Index           =   21
         Left            =   360
         TabIndex        =   65
         Top             =   1000
         Width           =   855
      End
      Begin VB.Label Label 
         Caption         =   "Address"
         Height          =   255
         Index           =   20
         Left            =   360
         TabIndex        =   64
         Top             =   750
         Width           =   855
      End
      Begin VB.Label Label 
         Caption         =   "Name"
         Height          =   255
         Index           =   19
         Left            =   360
         TabIndex        =   63
         Top             =   500
         Width           =   735
      End
      Begin VB.Label Label 
         Caption         =   "Dataset   "
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
         Index           =   18
         Left            =   480
         TabIndex        =   62
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame frameFiles 
      Height          =   615
      Index           =   3
      Left            =   120
      TabIndex        =   58
      Top             =   4680
      Visible         =   0   'False
      Width           =   6615
      Begin VB.TextBox txtFILE 
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
         Left            =   1560
         MaxLength       =   12
         TabIndex        =   31
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtFILE 
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
         Left            =   4800
         MaxLength       =   12
         TabIndex        =   32
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label 
         Caption         =   "Report File:"
         Height          =   255
         Index           =   16
         Left            =   240
         TabIndex        =   60
         Top             =   315
         Width           =   1215
      End
      Begin VB.Label Label 
         Caption         =   "Tiny Tafel File:"
         Height          =   255
         Index           =   17
         Left            =   3480
         TabIndex        =   59
         Top             =   315
         Width           =   1335
      End
   End
   Begin VB.Frame frameFiles 
      Height          =   615
      Index           =   2
      Left            =   120
      TabIndex        =   55
      Top             =   3960
      Visible         =   0   'False
      Width           =   6615
      Begin VB.TextBox txtFILE 
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
         Left            =   1560
         MaxLength       =   12
         TabIndex        =   29
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtFILE 
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
         Left            =   4800
         MaxLength       =   12
         TabIndex        =   30
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label 
         Caption         =   "GEDCOM File:"
         Height          =   255
         Index           =   14
         Left            =   240
         TabIndex        =   57
         Top             =   315
         Width           =   1215
      End
      Begin VB.Label Label 
         Caption         =   "MailMerge File:"
         Height          =   255
         Index           =   15
         Left            =   3480
         TabIndex        =   56
         Top             =   315
         Width           =   1335
      End
   End
   Begin VB.Frame frameFiles 
      Height          =   1335
      Index           =   1
      Left            =   120
      TabIndex        =   48
      Top             =   5400
      Visible         =   0   'False
      Width           =   6615
      Begin VB.TextBox txtFILE 
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
         Left            =   4800
         MaxLength       =   12
         TabIndex        =   28
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtFILE 
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
         Left            =   4800
         MaxLength       =   12
         TabIndex        =   27
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtFILE 
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
         Left            =   4800
         MaxLength       =   12
         TabIndex        =   26
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtFILE 
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
         Left            =   1560
         MaxLength       =   12
         TabIndex        =   25
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtFILE 
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
         Left            =   1560
         MaxLength       =   12
         TabIndex        =   24
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtFILE 
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
         Left            =   1560
         MaxLength       =   12
         TabIndex        =   23
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label 
         Caption         =   "Select File:"
         Height          =   255
         Index           =   13
         Left            =   3480
         TabIndex        =   54
         Top             =   1035
         Width           =   1095
      End
      Begin VB.Label Label 
         Caption         =   "Reference File:"
         Height          =   255
         Index           =   12
         Left            =   3480
         TabIndex        =   53
         Top             =   675
         Width           =   1335
      End
      Begin VB.Label Label 
         Caption         =   "Node Reference:"
         Height          =   255
         Index           =   11
         Left            =   3480
         TabIndex        =   52
         Top             =   315
         Width           =   1335
      End
      Begin VB.Label Label 
         Caption         =   "Relative File:"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   51
         Top             =   1035
         Width           =   1095
      End
      Begin VB.Label Label 
         Caption         =   "Descendant  File:"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   50
         Top             =   675
         Width           =   1455
      End
      Begin VB.Label Label 
         Caption         =   "Ancestor File:"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   49
         Top             =   315
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdINIT 
      Caption         =   "Initialize"
      Height          =   315
      Left            =   4920
      TabIndex        =   6
      Top             =   570
      Width           =   855
   End
   Begin VB.CommandButton cmdSTATS 
      Caption         =   "Stats"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   570
      Width           =   735
   End
   Begin VB.CommandButton cmdFILES 
      Caption         =   "Files"
      Height          =   315
      Index           =   0
      Left            =   5280
      TabIndex        =   9
      Top             =   1095
      Width           =   495
   End
   Begin VB.Frame frameFiles 
      Height          =   1335
      Index           =   0
      Left            =   120
      TabIndex        =   43
      Top             =   2595
      Visible         =   0   'False
      Width           =   6615
      Begin VB.TextBox txtFILE 
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
         Left            =   4560
         MaxLength       =   14
         TabIndex        =   22
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtFILE 
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
         Left            =   1560
         MaxLength       =   14
         TabIndex        =   19
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtFILE 
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
         Left            =   1560
         MaxLength       =   14
         TabIndex        =   20
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtFILE 
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
         Left            =   1560
         MaxLength       =   14
         TabIndex        =   21
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label 
         Caption         =   "Index File:"
         Height          =   255
         Index           =   7
         Left            =   3600
         TabIndex        =   47
         Top             =   315
         Width           =   855
      End
      Begin VB.Label Label 
         Caption         =   "Name Dataset:"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   46
         Top             =   315
         Width           =   1215
      End
      Begin VB.Label Label 
         Caption         =   "Address Dataset:"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   45
         Top             =   675
         Width           =   1335
      End
      Begin VB.Label Label 
         Caption         =   "Misc. Dataset:"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   44
         Top             =   1035
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdFILES 
      Caption         =   "Files"
      Height          =   315
      Index           =   3
      Left            =   5280
      TabIndex        =   18
      Top             =   2175
      Width           =   495
   End
   Begin VB.CommandButton cmdFILES 
      Caption         =   "Files"
      Height          =   315
      Index           =   2
      Left            =   5280
      TabIndex        =   15
      Top             =   1815
      Width           =   495
   End
   Begin VB.CommandButton cmdFILES 
      Caption         =   "Files"
      Height          =   315
      Index           =   1
      Left            =   5280
      TabIndex        =   12
      Top             =   1455
      Width           =   495
   End
   Begin VB.TextBox txtPATH 
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
      Left            =   1440
      MaxLength       =   27
      TabIndex        =   16
      Top             =   2160
      Width           =   3495
   End
   Begin VB.CommandButton cmdBROWSE 
      Caption         =   "<"
      Height          =   315
      Index           =   3
      Left            =   4920
      TabIndex        =   17
      Top             =   2175
      Width           =   375
   End
   Begin VB.TextBox txtPATH 
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
      Left            =   1440
      MaxLength       =   27
      TabIndex        =   13
      Top             =   1800
      Width           =   3495
   End
   Begin VB.CommandButton cmdBROWSE 
      Caption         =   "<"
      Height          =   315
      Index           =   2
      Left            =   4920
      TabIndex        =   14
      Top             =   1815
      Width           =   375
   End
   Begin VB.TextBox txtPATH 
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
      Left            =   1440
      MaxLength       =   27
      TabIndex        =   10
      Top             =   1440
      Width           =   3495
   End
   Begin VB.CommandButton cmdBROWSE 
      Caption         =   "<"
      Height          =   315
      Index           =   1
      Left            =   4920
      TabIndex        =   11
      Top             =   1455
      Width           =   375
   End
   Begin VB.CommandButton cmdCANCEL 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   6000
      TabIndex        =   36
      Top             =   575
      Width           =   735
   End
   Begin VB.CommandButton cmdBROWSE 
      Caption         =   "<"
      Height          =   315
      Index           =   0
      Left            =   4920
      TabIndex        =   8
      Top             =   1095
      Width           =   375
   End
   Begin VB.TextBox txtPATH 
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
      Left            =   1440
      MaxLength       =   27
      TabIndex        =   7
      Top             =   1080
      Width           =   3495
   End
   Begin VB.CheckBox chkDefault 
      Caption         =   "Default Setup"
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdADD 
      Caption         =   "Add"
      Height          =   315
      Left            =   3120
      TabIndex        =   4
      Top             =   570
      Width           =   735
   End
   Begin VB.CommandButton cmdDEL 
      Caption         =   "Delete"
      Height          =   315
      Left            =   3960
      TabIndex        =   5
      Tag             =   " "
      Top             =   570
      Width           =   855
   End
   Begin VB.CommandButton cmdRETURN 
      Caption         =   "Return"
      Default         =   -1  'True
      Height          =   315
      Left            =   6000
      TabIndex        =   37
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdSAVE 
      Caption         =   "Save"
      Height          =   315
      Left            =   6000
      TabIndex        =   34
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton cmdPRINT 
      Caption         =   "Report"
      Height          =   315
      Left            =   6000
      TabIndex        =   33
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton cmdLOAD 
      Caption         =   "Load"
      Height          =   315
      Left            =   6000
      TabIndex        =   35
      Top             =   1440
      Width           =   735
   End
   Begin VB.ComboBox FAMSetup 
      Height          =   315
      Left            =   1200
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox FAMDesc 
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
      Left            =   3120
      MaxLength       =   30
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6840
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label Label 
      Caption         =   "Report File Path:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   42
      Top             =   2205
      Width           =   1455
   End
   Begin VB.Label Label 
      Caption         =   "Transfer File Path:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   41
      Top             =   1845
      Width           =   1455
   End
   Begin VB.Label Label 
      Caption         =   "Work File Path:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   40
      Top             =   1485
      Width           =   1455
   End
   Begin VB.Label Label 
      Caption         =   "Family File Path:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   39
      Top             =   1125
      Width           =   1335
   End
   Begin VB.Line Line6 
      X1              =   0
      X2              =   5880
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line4 
      X1              =   5880
      X2              =   5880
      Y1              =   0
      Y2              =   2520
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   5880
      X2              =   6840
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label1 
      Caption         =   "Family Setup:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   38
      Top             =   175
      Width           =   975
   End
End
Attribute VB_Name = "frmFHSWFFDF"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
DefInt A-Z
Dim DefSETUP
Private Sub BuildFAMSetup()
FAMSetup.Clear
X = FSUTABLE.FT: Y = 0
While X > 0
    FAMSetup.AddItem RTrim$(Left$(FSETUP$(X), 8))
    FAMSetup.ItemData(FAMSetup.NewIndex) = X
    X = FSUCHN(X).FWD
Wend
X = 0: FOUND = 0
While (X < FAMSetup.ListCount) And (Not FOUND)
   FOUND = (FAMSetup.ItemData(X) = ENV.FSETUP)
   If Not FOUND Then X = X + 1
Wend
If FOUND Then FAMSetup.ListIndex = X
End Sub

Private Sub ChkPath(PTYPE$, ptext$)
hpath$ = CurDir: hdrv$ = Left$(hpath$, 1)
xpath$ = ptext$: xdrv$ = ""
X = InStr(ptext$, ":"): OKAY = -1
If X > 0 Then
   xdrv$ = Left$(xpath$, 1)
   xpath$ = MID$(xpath$, X + 1, 255)
End If
On Local Error GoTo BadDrive
ChDrive xdrv$
On Local Error GoTo BadPath
If xpath$ <> "" Then ChDir xpath$
On Local Error GoTo 0
ChDrive hdrv$: ChDir hpath$
If Not OKAY Then a = 27: CA$ = Chr$(27)
Exit Sub

BadDrive:
   SysVAR.MVAR = xdrv$
   Call PutMSG("Drive [MVAR] Not Found...")
   OKAY = 0
   Resume Next
   
BadPath:
   SysVAR.MVAR = xpath$
   Call PutMSG("Path '[MVAR]' Does Not Exist...Create Path Now? (Y/N)")
   If (CA$ <> SNGLKEY$(1)) Then
      OKAY = 0
      Resume Next
   End If
   Call NewPath(xpath$)
   Resume Next
   
End Sub
Private Sub GetDrvPath(drv$, path$, drvpath$)
xX$ = drvpath$
X = InStr(drvpath$, ":")
If X > 0 Then
   drv$ = Left$(xX$, 1)
   xX$ = MID$(drvpath$, X + 1, 255)
   Else
   drv$ = ""
End If
XL = Len(xX$)
If XL > 0 Then
   If MID$(xX$, XL, 1) = "\" Then xX$ = Left$(xX$, XL - 1)
End If
path$ = RTrim$(xX$)
   
End Sub

Private Sub NewPath(xpath$)
X = InStr(xpath$, "\"): Y = 1
L = Len(xpath$): OKAY = -1
On Local Error GoTo BldPath
While (X > 0) And (X < L)
   Y = InStr(X + 1, xpath$, "\")
   If Y = 0 Then Y = L + 1
   ChDir Left$(xpath$, Y)
   If OKAY Then X = Y Else X = 0
Wend
Exit Sub

BldPath:
   If Not OKAY Then Resume Next
   OKAY = 0: MkDir Left$(xpath$, Y)
   OKAY = -1
   Resume Next
   
End Sub

Private Sub SaveSetup()
FDFSetup.DESC = FAMDesc
FDFHDr.DEFAULT = DefSETUP
FSETUP$(ENV.FSETUP) = FDFSetup.NAME + RTrim$(FDFSetup.DESC)
FAMFPFX$ = txtPATH(0)
Call GetDrvPath(FDFSetup.FAMDRV, FDFSetup.FAMDIR, FAMFPFX$)
WRKFPFX$ = txtPATH(1)
Call GetDrvPath(FDFSetup.WRKDRV, FDFSetup.WRKDIR, WRKFPFX$)
TRANPFX$ = txtPATH(2)
Call GetDrvPath(FDFSetup.TRANDRV, FDFSetup.TRANDIR, TRANPFX$)
RPTFPFX$ = txtPATH(3)
Call GetDrvPath(FDFSetup.RPTDRV, FDFSetup.RPTDIR, RPTFPFX$)
FDFSetup.NFDRV = " ": FDFSetup.AFDRV = " "
FDFSetup.MFDRV = " ": FDFSetup.IFDRV = " "
For X = 0 To 3
   xdrv$ = " ": xfile$ = Space$(12)
   Call GetDrvPath(xdrv$, xfile$, txtFILE(X))
   LSet FT$(X + 1) = xfile$
   Select Case X
      Case 0: FDFSetup.NFDRV = xdrv$
      Case 1: FDFSetup.AFDRV = xdrv$
      Case 2: FDFSetup.MFDRV = xdrv$
      Case 3: FDFSetup.IFDRV = xdrv$
   End Select
Next X
FSETUPS(ENV.FSETUP) = FDFSetup
For X = 4 To 13
   If X < 12 Then
      LSet FT$(X + 1) = RTrim$(txtFILE(X))
      Else
      If X = 12 Then
         LSet FT$(14) = RTrim$(txtFILE(12))
         Else
         LSet FT$(13) = RTrim$(txtFILE(13))
      End If
   End If
Next X
For X = 1 To FDFHDr.FAMDSN
    FFNAME$(FDFHDr.FAMDSN * (ENV.FSETUP - 1) + X) = FT$(X)
Next X
End Sub

Private Sub ShowSetup()
FAMDesc = RTrim$(FDFSetup.DESC)
chkDefault.VALUE = -(DefSETUP = ENV.FSETUP)
txtPATH(0) = RTrim$(FAMFPFX$)
txtPATH(1) = RTrim$(WRKFPFX$)
txtPATH(2) = RTrim$(TRANPFX$)
txtPATH(3) = RTrim$(RPTFPFX$)
For X = 0 To 13
   xdrv$ = ""
   If X < 4 Then
      Select Case X
         Case 0: xdrv$ = FDFSetup.NFDRV
         Case 1: xdrv$ = FDFSetup.AFDRV
         Case 2: xdrv$ = FDFSetup.MFDRV
         Case 3: xdrv$ = FDFSetup.IFDRV
      End Select
      If (xdrv$ = " ") Or (xdrv$ = Chr$(0)) Then
         xdrv$ = ""
         Else
         xdrv$ = xdrv$ + ":"
      End If
      Else
      xdrv$ = ""
   End If
   If X < 12 Then
      txtFILE(X) = xdrv$ + RTrim$(FT$(X + 1))
      Else
      If X = 12 Then
         txtFILE(12) = RTrim$(FT$(14))
         Else
         txtFILE(13) = RTrim$(FT$(13))
      End If
   End If
Next X
If frameSTATS.Visible Then ShowStats
End Sub
Private Sub ShowStats()
Call FAMOpen("READQ")
FamFileOKAY = OKAY
If OKAY Then
   txtSTAT(0) = FF1Hdr.REVLVL
   txtSTAT(1) = XRPTDate(FF1Hdr.UPDTMD, FF1Hdr.UPDTY, 0)
   txtSTAT(2) = FF1Hdr.MAXID
   txtSTAT(3) = FF2Hdr.REVLVL
   txtSTAT(4) = XRPTDate(FF2Hdr.UPDTMD, FF2Hdr.UPDTY, 0)
   txtSTAT(5) = FF2Hdr.HIREC
   txtSTAT(6) = FF2Hdr.FREECNT
   txtSTAT(7) = FF3HDR.REVLVL
   txtSTAT(8) = XRPTDate(FF3HDR.UPDTMD, FF3HDR.UPDTY, 0)
   txtSTAT(9) = FFRBA&(FF3HDR.HIREC, 1)
   txtSTAT(10) = FFRBA&(FF3HDR.FREECNT, 1)
   txtSTAT(11) = XRPTDate(FF1Hdr.CHGDTMD, FF1Hdr.CHGDTY, 0)
   txtSTAT(12) = FF1Hdr.CHGCNT1
   txtSTAT(13) = FF1Hdr.CHGCNT2
   Else
   For X = 0 To 13: txtSTAT(X) = "": Next X
End If
Call FamCLOSE
End Sub

Private Sub chkDefault_Click()
If chkDefault.VALUE = 1 Then
   DefSETUP = ENV.FSETUP
   Else
   If DefSETUP = ENV.FSETUP Then
      chkDefault.VALUE = 1
   End If
End If
End Sub

Private Sub cmdADD_Click()
xname$ = Space$(8)
SaveSetup
GetSetupName:
   Call frmFHSWMsgD.GetSETUP(xname$)
   If (a = 27) Then Exit Sub
   X = FSUTABLE.FT: FOUND = 0
   While (X > 0) And (Not FOUND)
      FOUND = (xname$ = Left$(FSETUP$(X), 8))
      If Not FOUND Then X = FSUCHN(X).FWD
   Wend
   If FOUND Then
      Call PutMSG("Setup Name Already Exists...")
      If a = 27 Then Exit Sub
      GoTo GetSetupName
   End If
' Add New SETUP
  If FSUTABLE.FF = 0 Then
     ts = FSUTABLE.SIZE + 1
     ReDim Preserve FSETUP$(ts)
     ReDim Preserve FSUCHN(ts) As ChainPTRS
     ReDim Preserve FSETUPS(ts) As FileSetupREC, FFNAME$(ts * FDFHDr.FAMDSN)
     FSUTABLE.SIZE = ts
     FSUTABLE.FF = ts
  End If
  FSUTABLE.CT = ENV.FSETUP
  Call TableUPDT("ADD", FSUTABLE, FSUCHN())
  If a = 27 Then Exit Sub
  FDFSetup.NAME = xname$
  FAMDesc.TEXT = ""
  FSETUPS(FSUTABLE.CT) = FDFSetup
  For Z = 1 To FDFHDr.FAMDSN
      FFNAME$(FDFHDr.FAMDSN * (FSUTABLE.CT - 1) + Z) = FFNAME$(FDFHDr.FAMDSN * ENV.FSETUP + Z)
  Next Z
  ENV.FSETUP = FSUTABLE.CT
  chkDefault.VALUE = 0
  FDFHDr.SETUPS = FDFHDr.SETUPS + 1
  FSETUP$(ENV.FSETUP) = xname$
  txtFILE(0).TEXT = xname$ + ".NAM"
  txtFILE(1).TEXT = xname$ + ".ADR"
  txtFILE(2).TEXT = xname$ + ".OTH"
  txtFILE(3).TEXT = xname$ + ".NDX"
  SaveSetup
  BuildFAMSetup
  ShowSetup
  If framefiles(0).Visible = False Then
     Call cmdFILES_Click(0)
  End If
  FAMDesc.SetFocus

End Sub

Private Sub cmdBROWSE_Click(Index As Integer)
Select Case Index
   Case 0: xY$ = "FAMFPATH": FAMFPFX$ = txtPATH(0)
   Case 1: xY$ = "WRKFPATH": WRKFPFX$ = txtPATH(1)
   Case 2: xY$ = "TRNFPATH": TRANPFX$ = txtPATH(2)
   Case 3: xY$ = "RPTFPATH": RPTFPFX$ = txtPATH(3)
End Select
xX$ = txtPATH(Index).TEXT
Call ChkPath(xY$, xX$)
If Not OKAY Then Exit Sub
Call frmFHSWFile.LOCATE(xY$, "PATH")
If OKAY Then
   Select Case Index
      Case 0: xX$ = UCase$(FAMFPFX$)
      Case 1: xX$ = UCase$(WRKFPFX$)
      Case 2: xX$ = UCase$(TRANPFX$)
      Case 3: xX$ = UCase$(RPTFPFX$)
   End Select
   If Len(RTrim$(xX$)) > 3 Then
      xX$ = RTrim$(xX$) + "\"
   End If
End If
txtPATH(Index) = UCase(xX$)
End Sub
Private Sub cmdDEL_Click()
   Call PgmPREP("FHSFILE")
' Delete This SETUP
    If ENV.FSETUP = FDFHDr.DEFAULT Then
       Call PutMSG("FILEM002")
       a = 27: OKAY = 0: Exit Sub
    End If
    If FDFHDr.SETUPS = 1 Then
       Call PutMSG("Cannot Delete the Only Setup...")
       Exit Sub
    End If
    LSet SysVAR.MVAR = FSETUPS(ENV.FSETUP).NAME
    Call PutMSG("FILEM004")
    If CA$ <> SNGLKEY$(1) Then
       OKAY = 0: a = 27
       Exit Sub
    End If
    FSUTABLE.CT = ENV.FSETUP
    Call TableUPDT("DEL", FSUTABLE, FSUCHN())
    If a = 27 Then Exit Sub
    FDFHDr.VER = Chr$(255)
    FDFHDr.SETUPS = FDFHDr.SETUPS - 1
    Call FDFGetSet("DEFAULT")
    BuildFAMSetup
    ShowSetup

End Sub

Private Sub cmdFILES_Click(Index As Integer)
FrameV = framefiles(Index).Visible
For X = 0 To 3: framefiles(X).Visible = False: Next X
frameSTATS.Visible = False
If FrameV = True Then
   frmFHSWFFDF.Height = 2950
   framefiles(Index).Visible = False
   Else
   framefiles(Index).Visible = True
   framefiles(Index).TOP = 2550
   framefiles(Index).Left = 120
   Select Case Index
      Case 0: ' Family Group of Files
              formh = 4360
      Case 1: ' Work Group of Files
              formh = 4360
      Case 2: ' Transfer Group of Files
              formh = 3625
      Case 3: ' Report Group of Files
              formh = 3625
   End Select
   frmFHSWFFDF.Height = formh
End If
End Sub

Private Sub cmdINIT_Click()
SaveSetup
Call ChkPath("FAMFPATH", FAMFPFX$): If Not OKAY Then Exit Sub
Call ChkPath("WRKFPATH", WRKFPFX$): If Not OKAY Then Exit Sub
Call ChkPath("TRANPATH", TRANPFX$): If Not OKAY Then Exit Sub
Call ChkPath("RPTFPATH", RPTFPFX$): If Not OKAY Then Exit Sub
Call FAMOpen("INIT"): FamFileOKAY = OKAY
frameSTATS.Visible = False
cmdSTATS_Click
End Sub

Private Sub cmdLOAD_Click()
Call FDFOpen("INIT"): If Not OKAY Then Exit Sub
BuildFAMSetup
ShowSetup
End Sub

Private Sub cmdPrint_Click()
Call PutMSG("Select  1) File Setup Report  2) Family File Summary Report")
If Asc(CA$) = 27 Then Exit Sub
xA$ = CA$
SaveSetup
Select Case xA$
   Case "2": RptPGM = "LIST"
             RptCODE = "FSUM"
             Call RptLIST("INIT")
   Case Else: RptPGM = "FILE"
              RptCODE = "FILE"
              Call RptFile("INIT")
End Select
End Sub

Private Sub cmdRESET_Click()
 Call FAMOpen("WRITE"): FamFileOKAY = OKAY
 If OKAY Then
    Call PutMSG("Reset Change Date and Counters?  (Y/N)")
    If CA$ = SNGLKEY$(1) Then
       id& = 0
       While id& < FF1Hdr.MAXID
          id& = id& + 1: X& = 100 * id& + 1
          Get #1, X&, FF1Rec: Y = Asc(FF1Rec.STATUS2)
          If (Y And 96) > 0 Then
             FF1Rec.STATUS2 = Chr$(Y And 159)
             Put #1, X&, FF1Rec
          End If
       Wend
       FF1Hdr.CHGDTY = Val(MID$(DATE$, 7, 4))
       FF1Hdr.CHGDTMD = 100 * Val(MID$(DATE$, 1, 2)) + Val(MID$(DATE$, 4, 2))
       FF1Hdr.CHGCNT1 = 0: FF1Hdr.CHGCNT2 = 0
       FF1Hdr.REVLVL = 256 + (FF1Hdr.REVLVL And 255)
       FFCHANGED = -1
    End If
 End If
 Call FamCLOSE
 Call ShowStats
End Sub

Private Sub cmdRETURN_Click()
SaveSetup
Call mnuFHSWMain.BuildFAMSetup
Unload frmFHSWFFDF
End Sub

Private Sub cmdCANCEL_Click()
   Call mnuFHSWMain.BuildFAMSetup
   Unload frmFHSWFFDF
End Sub

Private Sub cmdSAVE_Click()
Call SaveSetup
Call PgmPREP("FHSFILE")
HDef = FDFHDr.DEFAULT: FDFHDr.DEFAULT = 1
xn$ = Space$(12)
' Save TABLE Changes Into PDF File
    Call FDFOpen("NEW"): If Not OKAY Then Exit Sub
    'Call PutMSG("FILEM008")
    FDFHDr.FType = "FDF": FDFHDr.VER = Chr$(0): FDFHDr.EOFRBA = 0
    FDFHDr.FAMDSN = 4: FDFHDr.OTHERDSN = ENV.DSNCNT - FDFHDr.FAMDSN
    FDFHDr.SETUPS = 0
    Put #1, 1, FDFHDr
    For X = 1 To ENV.DSNCNT
        FDFDSRec.TYPE = FType$(X)
        FDFDSRec.DSNGRP = FGrp$(X)
        FDFDSRec.DSNNAME = FT$(X)
        Put #1, , FDFDSRec
    Next X
    For Z = 0 To FAMSetup.ListCount - 1
        X = FAMSetup.ItemData(Z)
        FDFHDr.SETUPS = FDFHDr.SETUPS + 1
        If X = HDef Then FDFHDr.DEFAULT = FDFHDr.SETUPS
        Put #1, , FSETUPS(X)
        For Y = 1 To FDFHDr.FAMDSN
            LSet xn$ = FFNAME$(FDFHDr.FAMDSN * (X - 1) + Y)
            Put #1, , xn$
        Next Y
    Next Z
    FDFHDr.EOFRBA = Seek(1) - 1
    Put #1, 1, FDFHDr
    Close
    xname$ = FDFSetup.NAME
    Call FDFOpen("INITQ")
    Call FDFGetSet(xname$)
    BuildFAMSetup
End Sub

Private Sub cmdSTATS_Click()
If frameSTATS.Visible = False Then
   SaveSetup
   For X = 0 To 3: framefiles(X).Visible = False: Next X
   frmFHSWFFDF.Height = 2950
   frameSTATS.TOP = 2550
   frameSTATS.Left = 120
   ShowStats
   frameSTATS.Visible = True
   frmFHSWFFDF.Height = 4850
   Else
   frmFHSWFFDF.Height = 2950
   frameSTATS.Visible = False
End If
End Sub

Private Sub FAMDesc_GotFocus()
FAMDesc.SelStart = 0
FAMDesc.SelLength = Len(RTrim$(FAMDesc.TEXT))
End Sub


Private Sub FAMSetup_Click()
   If RTrim$(FAMSetup.TEXT) <> RTrim$(FDFSetup.NAME) Then
      If RTrim$(FAMSetup.TEXT) <> "" Then
         SaveSetup
         Else
         FAMSetup.TEXT = FDFSetup.NAME
      End If
      Call FDFGetSet(FAMSetup.TEXT)
      If OKAY Then
         ShowSetup
         Else
         FAMDesc.TEXT = ""
      End If
   End If
End Sub


Private Sub Form_Load()
FormLoaded(2) = FormLoaded(2) Or FamDefForm
' Set Default Form Location on Screen
frmHEIGHT = frmFHSWFFDF.Height
frmWIDTH = frmFHSWFFDF.WIDTH
t1 = GetSetting("FHSWIN", "frmFFDF", "TOP", 0)
If t1 = 0 Then
   frmHEIGHT = 2940: frmWIDTH = 6945
   t00 = mnuFHSWMain.TOP: w00 = mnuFHSWMain.WIDTH
   l00 = mnuFHSWMain.Left
   t1 = (t00 + 400)
   l1 = l00 + (w00 - frmWIDTH) / 2: If l1 < 0 Then l1 = 1
   Else
   t1 = 1440: l1 = 1440
End If
frmFHSWFFDF.TOP = GetSetting("FHSWIN", "frmFFDF", "Top", t1)
frmFHSWFFDF.Left = GetSetting("FHSWIN", "frmFFDF", "Left", l1)
DefSETUP = FDFHDr.DEFAULT
BuildFAMSetup
ShowSetup
End Sub



Private Sub Form_Unload(Cancel As Integer)
FormLoaded(2) = FormLoaded(2) And (255 - FamDefForm)
SaveSetting "FHSWIN", "frmFFDF", "Top", frmFHSWFFDF.TOP
SaveSetting "FHSWIN", "frmFFDF", "Left", frmFHSWFFDF.Left
mnuFHSWMain.SetFocus
End Sub


Private Sub txtFILE_GotFocus(Index As Integer)
txtFILE(Index).SelStart = 0
txtFILE(Index).SelLength = Len(RTrim$(txtFILE(Index).TEXT))
End Sub

Private Sub txtFILE_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
End Sub


Private Sub txtPATH_GotFocus(Index As Integer)
txtPATH(Index).SelStart = 0
txtPATH(Index).SelLength = Len(RTrim$(txtPATH(Index).TEXT))
End Sub


Private Sub txtPATH_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
End Sub


