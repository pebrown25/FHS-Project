VERSION 4.00
Begin VB.Form frmFHSWSlct 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FHSWSLCT - Select Records to Process"
   ClientHeight    =   5910
   ClientLeft      =   1830
   ClientTop       =   2025
   ClientWidth     =   5730
   Height          =   6315
   Left            =   1770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   Top             =   1680
   Width           =   5850
   Begin VB.CommandButton cmdFILE 
      Caption         =   "AND"
      Height          =   255
      Index           =   2
      Left            =   2880
      TabIndex        =   6
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdFILE 
      Caption         =   "Exclude File"
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdFILE 
      Caption         =   "Merge File"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
   Begin VB.Frame SecondSelect 
      Height          =   1215
      Left            =   80
      TabIndex        =   62
      Top             =   4680
      Width           =   5580
      Begin VB.TextBox txt2ndSelect 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   615
         Left            =   360
         MultiLine       =   -1  'True
         TabIndex        =   65
         Text            =   "FHSWSLCT.frx":0000
         Top             =   450
         Width           =   1695
      End
      Begin VB.CheckBox chk2ndSelect 
         Caption         =   "(Latest) Spouse of Child"
         Height          =   255
         Index           =   5
         Left            =   3480
         TabIndex        =   40
         Top             =   840
         Width           =   2055
      End
      Begin VB.CheckBox chk2ndSelect 
         Caption         =   "Parent of Spouse"
         Height          =   195
         Index           =   4
         Left            =   3480
         TabIndex        =   38
         Top             =   540
         Width           =   1815
      End
      Begin VB.CheckBox chk2ndSelect 
         Caption         =   "Spouse "
         Height          =   255
         Index           =   3
         Left            =   3480
         TabIndex        =   36
         Top             =   240
         Width           =   1815
      End
      Begin VB.CheckBox chk2ndSelect 
         Caption         =   "Children "
         Height          =   210
         Index           =   2
         Left            =   2280
         TabIndex        =   39
         Top             =   840
         Width           =   1095
      End
      Begin VB.CheckBox chk2ndSelect 
         Caption         =   "Siblings"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   37
         Top             =   540
         Width           =   975
      End
      Begin VB.CheckBox chk2ndSelect 
         Caption         =   "Parents"
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   35
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lbl2ndSelect 
         Caption         =   "Secondary Selections:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   64
         Top             =   200
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdRule 
      Caption         =   "List"
      Height          =   255
      Index           =   2
      Left            =   2880
      TabIndex        =   9
      Top             =   940
      Width           =   855
   End
   Begin VB.CommandButton cmdFILE 
      Caption         =   "Save Results"
      Height          =   315
      Index           =   3
      Left            =   2880
      TabIndex        =   2
      Top             =   100
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Return"
      Default         =   -1  'True
      Height          =   315
      Left            =   4560
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdRule 
      Caption         =   "2nd Select"
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   8
      Top             =   940
      Width           =   1215
   End
   Begin VB.CommandButton cmdRule 
      Caption         =   "Search"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   1335
   End
   Begin VB.Frame Counters 
      Height          =   855
      Left            =   3820
      TabIndex        =   57
      Top             =   360
      Width           =   1830
      Begin VB.TextBox txtCount 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   960
         TabIndex        =   61
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtCount 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   960
         TabIndex        =   59
         Top             =   140
         Width           =   615
      End
      Begin VB.Label lblCount 
         Caption         =   "Secondary:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   60
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblCount 
         Caption         =   "Primary:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   58
         Top             =   170
         Width           =   615
      End
   End
   Begin VB.ComboBox cboBaseType 
      Height          =   315
      ItemData        =   "FHSWSLCT.frx":0042
      Left            =   1080
      List            =   "FHSWSLCT.frx":0055
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   100
      Width           =   1695
   End
   Begin VB.Frame SearchRule 
      Enabled         =   0   'False
      Height          =   3495
      Left            =   80
      TabIndex        =   55
      Top             =   1200
      Width           =   5580
      Begin VB.TextBox txtRule 
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
         Index           =   20
         Left            =   3720
         MaxLength       =   1
         MultiLine       =   -1  'True
         TabIndex        =   31
         Top             =   3120
         Width           =   255
      End
      Begin VB.CommandButton cmdRCHG 
         Caption         =   "C"
         Height          =   255
         Index           =   0
         Left            =   5040
         TabIndex        =   32
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton cmdRCHG 
         Caption         =   "H"
         Height          =   255
         Index           =   1
         Left            =   5040
         TabIndex        =   33
         Top             =   2040
         Width           =   255
      End
      Begin VB.CommandButton cmdRCHG 
         Caption         =   "R"
         Height          =   255
         Index           =   2
         Left            =   5040
         TabIndex        =   34
         Top             =   2400
         Width           =   255
      End
      Begin VB.TextBox txtRule 
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
         Left            =   3720
         MaxLength       =   1
         MultiLine       =   -1  'True
         TabIndex        =   28
         Top             =   2760
         Width           =   255
      End
      Begin VB.TextBox txtRule 
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
         Index           =   13
         Left            =   4080
         MaxLength       =   1
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   2040
         Width           =   255
      End
      Begin VB.TextBox txtRule 
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
         Index           =   9
         Left            =   4080
         MaxLength       =   1
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   1320
         Width           =   255
      End
      Begin VB.TextBox txtRulenum 
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
         Index           =   1
         Left            =   5000
         TabIndex        =   68
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtRulenum 
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
         Left            =   4320
         TabIndex        =   41
         Top             =   240
         Width           =   375
      End
      Begin VB.HScrollBar nxtRULE 
         Height          =   255
         Left            =   3600
         TabIndex        =   13
         Top             =   260
         Width           =   735
      End
      Begin VB.TextBox txtRule 
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
         Index           =   18
         Left            =   2040
         MaxLength       =   1
         MultiLine       =   -1  'True
         TabIndex        =   30
         Top             =   3120
         Width           =   255
      End
      Begin VB.TextBox txtRule 
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
         Index           =   17
         Left            =   120
         MaxLength       =   1
         MultiLine       =   -1  'True
         TabIndex        =   29
         Top             =   3120
         Width           =   255
      End
      Begin VB.TextBox txtRule 
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
         Index           =   16
         Left            =   2040
         MaxLength       =   1
         MultiLine       =   -1  'True
         TabIndex        =   27
         Top             =   2760
         Width           =   255
      End
      Begin VB.TextBox txtRule 
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
         Index           =   4
         Left            =   4920
         MaxLength       =   1
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   960
         Width           =   255
      End
      Begin VB.TextBox txtRule 
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
         Index           =   3
         Left            =   4230
         MaxLength       =   1
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   960
         Width           =   255
      End
      Begin VB.TextBox txtRule 
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
         Left            =   1200
         MaxLength       =   41
         TabIndex        =   25
         Top             =   2400
         Width           =   3615
      End
      Begin VB.TextBox txtRule 
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
         Index           =   15
         Left            =   120
         MaxLength       =   1
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   2760
         Width           =   255
      End
      Begin VB.TextBox txtRule 
         Height          =   315
         Index           =   12
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   23
         Text            =   "99-99-9999"
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtRule 
         Height          =   315
         Index           =   11
         Left            =   480
         MaxLength       =   10
         TabIndex        =   22
         Text            =   "00-00-0000"
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtRule 
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
         Left            =   1200
         MaxLength       =   41
         TabIndex        =   21
         Top             =   1680
         Width           =   3615
      End
      Begin VB.TextBox txtRule 
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
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   14
         Top             =   600
         Width           =   3255
      End
      Begin VB.TextBox txtRule 
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
         MaxLength       =   20
         TabIndex        =   15
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox txtRule 
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
         Index           =   2
         Left            =   3120
         MaxLength       =   1
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtRule 
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
         Index           =   1
         Left            =   1680
         MaxLength       =   5
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtRule 
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
         Index           =   0
         Left            =   120
         MaxLength       =   5
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtRule 
         Height          =   315
         Index           =   8
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   19
         Text            =   "99-99-9999"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtRule 
         Height          =   315
         Index           =   7
         Left            =   480
         MaxLength       =   10
         TabIndex        =   18
         Text            =   "00-00-0000"
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblRule 
         Caption         =   "Changed (Y/N/A)"
         Height          =   255
         Index           =   11
         Left            =   4050
         TabIndex        =   67
         Top             =   3180
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Events    (Y/N)"
         Height          =   255
         Left            =   4080
         TabIndex        =   66
         Top             =   2805
         Width           =   1095
      End
      Begin VB.Label lblRule 
         Caption         =   "  of"
         Height          =   255
         Index           =   10
         Left            =   4680
         TabIndex        =   63
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblCheck 
         Caption         =   "Comments  Y/N"
         Height          =   255
         Index           =   3
         Left            =   2400
         TabIndex        =   54
         Top             =   3150
         Width           =   1215
      End
      Begin VB.Label lblCheck 
         Caption         =   "Residence  Y/N/C"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   53
         Top             =   3150
         Width           =   1455
      End
      Begin VB.Label lblCheck 
         Caption         =   "Children  Y/N"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   52
         Top             =   2805
         Width           =   1215
      End
      Begin VB.Label lblRule 
         Caption         =   "M:"
         Height          =   255
         Index           =   3
         Left            =   4680
         TabIndex        =   46
         Top             =   990
         Width           =   255
      End
      Begin VB.Label lblRule 
         Caption         =   "F: "
         Height          =   255
         Index           =   2
         Left            =   4000
         TabIndex        =   45
         Top             =   990
         Width           =   135
      End
      Begin VB.Label lblRule 
         Caption         =   "Death Place:"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   50
         Top             =   2400
         Width           =   975
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   5400
         Y1              =   2715
         Y2              =   2715
      End
      Begin VB.Label lblCheck 
         Caption         =   "Married  Y/N/C"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   51
         Top             =   2805
         Width           =   1215
      End
      Begin VB.Label lblRule 
         Caption         =   "<= Death Date <="
         Height          =   255
         Index           =   8
         Left            =   1560
         TabIndex        =   49
         Top             =   2100
         Width           =   1335
      End
      Begin VB.Label lblRule 
         Caption         =   "Birth Place:"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   48
         Top             =   1725
         Width           =   975
      End
      Begin VB.Label lblRule 
         Caption         =   "Given Name:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   43
         Top             =   660
         Width           =   1095
      End
      Begin VB.Label lblRule 
         Caption         =   "Surname:"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   44
         Top             =   1005
         Width           =   855
      End
      Begin VB.Label lblRule 
         Caption         =   "Gender:"
         Height          =   255
         Index           =   1
         Left            =   2505
         TabIndex        =   42
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblRule 
         Caption         =   "<= RID <="
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   0
         Top             =   275
         Width           =   855
      End
      Begin VB.Label lblRule 
         Caption         =   " <= Birth Date <="
         Height          =   255
         Index           =   6
         Left            =   1560
         TabIndex        =   47
         Top             =   1350
         Width           =   1335
      End
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   3840
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblBase 
      Caption         =   "Base Set:"
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
      Left            =   120
      TabIndex        =   56
      Top             =   150
      Width           =   855
   End
End
Attribute VB_Name = "frmFHSWSlct"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
DefInt A-Z
Dim ChangeMODE

' Variables from FHSLIST
Dim SlctCNT, SecSlctCNT, SRuleCNT, RuleNUM
Dim SECSLCT, SAVEWORK
Dim UFLD(), RULES() As SlctFileRULE, XRULE As SlctFileRULE
Dim KeyWORD$, SrchWORD$, SNDX$
Dim HOLDITEM As Variant, xFT
Dim HTxtRule() As String



Private Sub ChkRULE(XRULE As SlctFileRULE, R)
565 ' Check for Modified SubRules
    XUFLD = ((XRULE.LOID <> RULES(12).LOID) Or (XRULE.HIID <> RULES(12).HIID))
    XUFLD = XUFLD + 2 * (XRULE.SEX <> RULES(12).SEX)
    XUFLD = XUFLD + 4 * (XRULE.SRNM <> RULES(12).SRNM)
    XUFLD = XUFLD + 8 * (XRULE.GVNM <> RULES(12).GVNM)
    XUFLD = XUFLD + 16 * ((XRULE.LOBY <> RULES(12).LOBY) Or (XRULE.LOBMD <> RULES(12).LOBMD) Or (XRULE.HIBY <> RULES(12).HIBY) Or (XRULE.HIBMD <> RULES(12).HIBMD))
    XUFLD = XUFLD + 32 * (XRULE.BPLC <> RULES(12).BPLC)
    XUFLD = XUFLD + 64 * ((XRULE.LODY <> RULES(12).LODY) Or (XRULE.LODMD <> RULES(12).LODMD) Or (XRULE.HIDY <> RULES(12).HIDY) Or (XRULE.HIDMD <> RULES(12).HIDMD))
    XUFLD = XUFLD + 128 * (XRULE.DPLC <> RULES(12).DPLC)
    XUFLD = XUFLD + 256 * (XRULE.YORN <> RULES(12).YORN)
    XUFLD = XUFLD + 512 * ((XRULE.BDST <> RULES(12).BDST) Or (XRULE.DDST <> RULES(12).DDST))
    XUFLD = XUFLD + 1024 * ((XRULE.FIDST <> RULES(12).FIDST) Or (XRULE.MIDST <> RULES(12).MIDST))
    UFLD(R) = -XUFLD
    If (UFLD(R) > 0) And (R > SRuleCNT) Then
       SRuleCNT = R
       'nxtRULE.Max = SRuleCNT + 1
    End If
End Sub


Private Sub SaveRule(R)
    xY = Asc(SNGLKEY$(1)): xn = Asc(SNGLKEY$(2))
    xA = Asc(SNGLKEY$(4)): xC = Asc(SNGLKEY$(18))
    YORN = 0: 'DFM = Val(CONFIG.DFM) - 1
    For X = 0 To 20
        xX$ = LTrim$(RTrim$(txtRule(X))): If xX$ = "" Then xX$ = " "
        Select Case X
           Case 0: GoSub 569: XRULE.LOID = v#
           Case 1: GoSub 569: XRULE.HIID = v#
           Case 2: XRULE.SEX = xX$
           Case 3: XRULE.FIDST = xX$
           Case 4: XRULE.MIDST = xX$
           Case 5: LSet XRULE.GVNM = xX$
           Case 6: LSet XRULE.SRNM = xX$
           Case 7: XRULE.LOBMD = DATEMD%(xX$, DFM)
                        XRULE.LOBY = DATEY%(xX$, DFM)
           Case 8: XRULE.HIBMD = DATEMD%(xX$, DFM)
                        XRULE.HIBY = DATEY%(xX$, DFM)
           Case 9: XRULE.BDST = xX$
           Case 10: LSet XRULE.BPLC = xX$
           Case 11: XRULE.LODMD = DATEMD%(xX$, DFM)
                        XRULE.LODY = DATEY%(xX$, DFM)
           Case 12: XRULE.HIDMD = DATEMD%(xX$, DFM)
                        XRULE.HIDY = DATEY%(xX$, DFM)
           Case 13: XRULE.DDST = xX$
           Case 14: LSet XRULE.DPLC = xX$
           Case 15: YORN = YORN - (1 * (Asc(xX$) = xY) + 2 * (Asc(xX$) = xn) + 4 * (Asc(xX$) = xC))
           Case 16: YORN = YORN - (8 * (Asc(xX$) = xY) + 16 * (Asc(xX$) = xn))
           Case 17: YORN = YORN - (32 * (Asc(xX$) = xY) + 64 * (Asc(xX$) = xn) + 128 * (Asc(xX$) = xC))
           Case 18: YORN = YORN - (256 * (Asc(xX$) = xY) + 512 * (Asc(xX$) = xn))
           Case 19: YORN = YORN - (1024 * (Asc(xX$) = xY) + 2048 * (Asc(xX$) = xn))
           Case 20: YORN = YORN - (4096 * (Asc(xX$) = xY) + 8192 * (Asc(xX$) = xn) + 16384 * (Asc(xX$) = xA))
        End Select
    Next X
    XRULE.YORN = YORN
    RULES(R) = XRULE
    If R < 11 Then Call ChkRULE(RULES(R), R)
    Exit Sub
    
569 ' CHECK FOR INTEGER TOO LARGE
    v# = Val(xX$): If v# > 32767 Then v# = 32767
    Return
End Sub

Private Sub ShowRULE(R)
    xY = Asc(SNGLKEY$(1)): xn = Asc(SNGLKEY$(2))
    xA = Asc(SNGLKEY$(4)): xC = Asc(SNGLKEY$(18))
    XRULE = RULES(R): YORN = XRULE.YORN
    If (0 < R) And (R < 11) Then
       'txtRulenum(0).TEXT = SFILL$(R, 3)
       'txtRulenum(1).TEXT = SFILL$(SRuleCNT, 3)
       txtRulenum(0).TEXT = R
       txtRulenum(1).TEXT = SRuleCNT
    End If
    For X = 0 To 20
        Select Case X
           Case 0: xX$ = (Str$(XRULE.LOID))
           Case 1: xX$ = (Str$(XRULE.HIID))
           Case 2: xX$ = XRULE.SEX
           Case 3: xX$ = XRULE.FIDST
           Case 4: xX$ = XRULE.MIDST
           Case 5: xX$ = (XRULE.GVNM)
           Case 6: xX$ = (XRULE.SRNM)
           Case 7: xX$ = XSCRNDate$(XRULE.LOBMD, XRULE.LOBY)
           Case 8: xX$ = XSCRNDate$(XRULE.HIBMD, XRULE.HIBY)
           Case 9: xX$ = XRULE.BDST
           Case 10: xX$ = (XRULE.BPLC)
           Case 11: xX$ = XSCRNDate$(XRULE.LODMD, XRULE.LODY)
           Case 12: xX$ = XSCRNDate$(XRULE.HIDMD, XRULE.HIDY)
           Case 13: xX$ = XRULE.DDST
           Case 14: xX$ = (XRULE.DPLC)
           Case 15: xX$ = Chr$(-(xY * ((YORN And 1) > 0) + xn * ((YORN And 2) > 0) + xC * ((YORN And 4) > 0) + 32 * ((YORN And 7) = 0)))
           Case 16: xX$ = Chr$(-(xY * ((YORN And 8) > 0) + xn * ((YORN And 16) > 0) + 32 * ((YORN And 24) = 0)))
           Case 17: xX$ = Chr$(-(xY * ((YORN And 32) > 0) + xn * ((YORN And 64) > 0) + xC * ((YORN And 128) > 0) + 32 * ((YORN And 224) = 0)))
           Case 18: xX$ = Chr$(-(xY * ((YORN And 256) > 0) + xn * ((YORN And 512) > 0) + 32 * ((YORN And 768) = 0)))
           Case 19: xX$ = Chr$(-(xY * ((YORN And 1024) > 0) + xn * ((YORN And 2048) > 0) + 32 * ((YORN And 3072) = 0)))
           Case 20: xX$ = Chr$(-(xY * ((YORN And 4096) > 0) + xn * ((YORN And 8192) > 0) + xA * ((YORN And 16384) > 0) + 32 * ((YORN And 28672) = 0)))
        End Select
        txtRule(X).TEXT = Trim$(xX$)
    Next X
End Sub
Sub TEXTSearch()

X1200: ' Compare KEYWORD$ and SRCHWORD$
     OKAY = 0
     KO = 1: KL = Len(RTrim$(KeyWORD$))
     L = Len(RTrim$(SrchWORD$)): If L = 0 Then L = 1
     MISMATCH = (MID$(KeyWORD$, 1, 1) = "#"): KO = KO - MISMATCH
     LOSEARCH = (MID$(KeyWORD$, KO, 1) = "^"): KO = KO - LOSEARCH
     EXSEARCH = (MID$(KeyWORD$, KO, 1) = "^"): KO = KO - EXSEARCH
     EXMATCH = (MID$(KeyWORD$, KO, 1) = "="): KO = KO - EXMATCH
     If MID$(KeyWORD$, KO, 1) = "~" Then GoTo X1250
     If (KL < KO) Then KL = KO
     If EXMATCH Then
        OKAY = (Left$(SrchWORD$, L) = MID$(KeyWORD$, KO, KL - KO + 1))
        GoTo 1290
     End If
1215 OFST = InStr(Left$(SrchWORD$, L), MID$(KeyWORD$, KO, KL - KO + 1))
     OKAY = (OFST > 0): If OKAY Then GoTo 1290
     If Not LOSEARCH Then GoTo 1290
     ASCA = Asc(MID$(KeyWORD$, KO, 1)): If (ASCA > 95) Then GoTo 1290
     Mid$(KeyWORD$, KO, 1) = Chr$(ASCA + 32)
     GoTo 1215

X1250: ' Perform Soundex Search
1260 Call ComputeSNDX(MID$(KeyWORD$, KO + 1, KL - KO), SNDX$)
     SNDXCHK$ = SNDX$: GoSub X1300

1290 ' Return to Caller
     If MISMATCH Then OKAY = (Not OKAY)
     Exit Sub

X1300: ' Soundex Search
     OFST = 0: SXK$ = Left$(SNDXCHK$, 1): OKAY = 0: L = 4
     If EXSEARCH Then
        While MID$(SNDXCHK$, L, 1) = "0": L = L - 1: Wend
     End If
1310 OFST = InStr(OFST + 1, SrchWORD$, SXK$)
     If OFST = 0 Then GoTo 1320
     LSet KeyWORD$ = MID$(SrchWORD$, OFST, KL)
     Call ComputeSNDX(KeyWORD$, SNDX$)
     If Left$(SNDX$, L) = Left$(SNDXCHK$, L) Then OKAY = -1: GoTo 1340
     If OFST < 30 Then GoTo 1310 Else OFST = 0
1320 If Not LOSEARCH Then GoTo 1340
     SXA = Asc(SXK$)
     If (SXA < 91) Then LSet SXK$ = Chr$(SXA + 32): GoTo 1310
1340 Return

End Sub


Rem $STATIC
Sub SelectSET(OPT$)
    SlctCNT = PSCOUNT: SecSlctCNT = SSCOUNT
    Select Case OPT$
       Case "App2ndS": If SECSLCT = 0 Then Exit Sub
                      Apply2ndSelect = True: GoSub 500
       Case "AppRule": If SRuleCNT = 0 Then Exit Sub
                      ApplyRule = True: GoSub 500
    End Select
    Exit Sub
    
250 ' Place Selection Rule Data on Screen
    ShowRULE (R)
    Return

500 ' Create Selection SET
    If ApplyRule Then
       R = 1: GoSub 250
       KeyWORD$ = Space$(41): SrchWORD$ = KeyWORD$
       SNDX$ = Space$(4)
    End If
    GoSub 600
    Return

560 ' Read Updated RULE INFO from screen
    'X = Fmt.BGNFLD
    R = Val(txtRulenum(0).TEXT)
    SaveRule (R)
    Return

570 ' Adjust Rule Count (if necessary)
    If (R >= SRuleCNT) And (UFLD(R) <> 0) Then
       SRuleCNT = R
       If (R < 9) Then
          RULES(R + 1) = RULES(12): UFLD(R + 1) = 0
       End If
    End If
575 If (R = SRuleCNT) And (UFLD(R) = 0) Then SRuleCNT = SRuleCNT - 1
    Return

595 ' Eliminate Blank Rules
    R = SRuleCNT: J = 0: I = 1
    While I <= SRuleCNT
       If (UFLD(I) = 0 And J = 0) Then
          J = I
          Else
          If (UFLD(I) > 0 And J > 0) Then
             RULES(J) = RULES(I): R = J: J = J + 1
          End If
       End If
       I = I + 1
    Wend: If UFLD(1) = 0 Then SRuleCNT = 0 Else SRuleCNT = R
599 GoTo 695

600 ' Apply RULES or Secondary Selections to Selection BASE
    If FF1Hdr.MAXID = 0 Then
       Call PutMSG("LISTM016")
       GoTo 695
    End If
    If (SlctCNT > 0) And (SecSlctCNT > 0) Then
       Call PutMSG("LISTM017")
       If a = 27 Then GoTo 695
       ST = Val(CA$): If (ST < 1) Or (ST > 3) Then ST = 3
       Else
       ST = 3
    End If
    PCNT = -(SlctCNT * (ST And 1) - SecSlctCNT * ((ST And 2) > 0))
    If PCNT = 0 Then
       Call PutMSG("LISTM018")
       GoTo 695
    End If

605 ' Determine LOID and HIID in Selection Rules
    If SRuleCNT = 0 Then GoTo 630
    LO = 32767: HI = 0
    For R = 1 To SRuleCNT
      If LO > RULES(R).LOID Then LO = RULES(R).LOID
      If HI < RULES(R).HIID Then HI = RULES(R).HIID
    Next R
    If LOID < LO Then LOID = LO
    If HIID > HI Then HIID = HI

630 ' Restrict to ID's between LOID and HIID
    If HIID > FF1Hdr.MAXID Then HIID = FF1Hdr.MAXID
    If LOID > HIID Then LOID = HIID
    SlctCNT = 0: SecSlctCNT = 0: CHKCNT = 0
    LO = 32767: HI = 0: I = 1
    While I < LOID: SelSET(I) = 0: I = I + 1: Wend
    While I < HIID + 1
       If (SelSET(I) > 0) Then
          If (SelSET(I) And ST) > 0 Then
             SelSET(I) = SelSET(I) Or 8
             CHKCNT = CHKCNT + 1
          End If
          SlctCNT = SlctCNT - ((SelSET(I) And 1) > 0)
          SecSlctCNT = SecSlctCNT - ((SelSET(I) And 2) > 0)
          If HI = 0 Then LO = I
          HI = I
       End If
       I = I + 1
    Wend
    While I < FF1Hdr.MAXID + 1: SelSET(I) = 0: I = I + 1: Wend
    LOID = LO: HIID = HI: 'GoSub 230

640 ' Process Selection Option
    If CHKCNT = 0 Then GoTo 695
    If Not ENV.FFOPEN Then Call FAMOpen("READ"): If Not OKAY Then GoTo 695
    LO = FF1Hdr.MAXID: HI = 0
    'If ApplyRule Then Call PutMSG("LISTM020") Else Call PutMSG("LISTM022")
    For I = LOID To HIID
       If (SelSET(I) And 8) > 0 Then
          If ApplyRule Then R = 1: GoSub 650 Else GoSub 700
       End If
    Next I
    GoTo 690

650 ' Process Selection Rule "R"
    XUFLD = UFLD(R): XRULE = RULES(R)
    If (XUFLD And 1) Then
       If (XRULE.LOID > I) Or (XRULE.HIID < I) Then
          OKAY = 0: GoTo 675
       End If
    End If
    Call FF1GetRec(I): If Not OKAY Then GoTo 679
    DY = FF1Rec.DY: SPREC = FF1Rec.SPOUSE
    COMREC = FF1Rec.COM: EVTREC = FF1Rec.EVENT
    OLDCH = FF1Rec.OLDCH: ADRSPTR = FF1Rec.ADRS
    STAT1 = Asc(FF1Rec.STATUS): STAT2 = Asc(FF1Rec.STATUS2)
    If XUFLD And 240 Then
       If XUFLD And 160 Then BPL$ = " " Else BPL$ = ""
       Call GetBDEvents(BD$, BPL$, DD$, DPL$, AGE$)
    End If
    If (XUFLD And 2) Then If XRULE.SEX <> FF1Rec.SEX Then GoTo 675
    If (XUFLD And 4) Then
       LSet KeyWORD$ = RTrim$(XRULE.SRNM)
       LSet SrchWORD$ = RTrim$(FF1Rec.SURNM)
       Call TEXTSearch
       If Not OKAY Then GoTo 675
    End If
    If (XUFLD And 8) Then
       LSet KeyWORD$ = RTrim$(XRULE.GVNM)
       LSet SrchWORD$ = RTrim$(FF1Rec.GIVEN)
       Call TEXTSearch
       If Not OKAY Then GoTo 675
    End If
    If (XUFLD And 16) Then
       If (FF1Rec.BY < XRULE.LOBY) Or (FF1Rec.BY > XRULE.HIBY) Then GoTo 675
       If (FF1Rec.BMD < XRULE.LOBMD) Or (FF1Rec.BMD > XRULE.HIBMD) Then GoTo 675
    End If
    If (XUFLD And 64) Then
       If (FF1Rec.DY < XRULE.LODY) Or (FF1Rec.DY > XRULE.HIDY) Then GoTo 675
       If (FF1Rec.DMD < XRULE.LODMD) Or (FF1Rec.DMD > XRULE.HIDMD) Then GoTo 675
    End If
655 If (XUFLD And 256) = 0 Then GoTo 660
    YORN = XRULE.YORN
    If ((YORN And 5) And (SPREC = 0)) Then GoTo 675
    If (YORN And 4) Then
       If DY > 0 Then GoTo 675
       Call FF3GetRec(SPREC)
       If (FF3SPOUSE.D2Y > 0) Or (FF3SPOUSE.D2MD > 0) Then GoTo 675
    End If
    If ((YORN And 2) And (SPREC <> 0)) Then GoTo 675
    If ((YORN And 8) And (OLDCH = 0)) Then GoTo 675
    If ((YORN And 16) And (OLDCH <> 0)) Then GoTo 675
    If ((YORN And 256) And (COMREC = 0)) Then GoTo 675
    If ((YORN And 512) And (COMREC <> 0)) Then GoTo 675
    If ((YORN And 1024) And (EVTREC = 0)) Then GoTo 675
    If ((YORN And 2048) And (EVTREC <> 0)) Then GoTo 675
    If ((YORN And 4096) And ((STAT1 And 64) = 0)) Then GoTo 675
    If ((YORN And 8192) And (((STAT1 And 64) > 0) Or ((STAT2 And 32) > 0))) Then GoTo 675
    If ((YORN And 16384) And (((STAT1 And 64) = 0) And ((STAT2 And 32) = 0))) Then GoTo 675
    If ((YORN And 224) = 0) Then GoTo 660
    If (ADRSPTR = 0) And (SPREC <> 0) Then
       Call FF3GetRec(SPREC): ADRSPTR = FF3SPOUSE.LOC
    End If
    If ((YORN And 160) And (ADRSPTR = 0)) Then GoTo 675
    If ((YORN And 64) And (ADRSPTR <> 0)) Then GoTo 675
    If ((YORN And 128) = 0) Then GoTo 660
    Call GetLastRES: If (FF2PFX.D2Y <> 0) Or (FF2PFX.D2MD <> 0) Then GoTo 675
660 If (XUFLD And 160) = 0 Then GoTo 665
    If (XUFLD And 32) Then
       LSet KeyWORD$ = RTrim$(XRULE.BPLC)
       LSet SrchWORD$ = RTrim$(BPLACE)
       Call TEXTSearch
       If Not OKAY Then GoTo 675
    End If
    If (XUFLD And 128) Then
       LSet KeyWORD$ = RTrim$(XRULE.DPLC)
       LSet SrchWORD$ = RTrim$(DPLACE)
       Call TEXTSearch
    End If
    If Not OKAY Then GoTo 675
665 If (XUFLD And 512) Then
       OKAY = (XRULE.BDST = RULES(12).BDST)
       If Not OKAY Then
          OKAY = (XRULE.BDST = FDSTAT$("BD"))
          If Not OKAY Then GoTo 675
       End If
       OKAY = (XRULE.DDST = RULES(12).DDST)
       If Not OKAY Then OKAY = (XRULE.DDST = FDSTAT$("DD"))
    End If
    If (XUFLD And 1024) Then
       OKAY = (XRULE.FIDST = RULES(12).FIDST)
       If Not OKAY Then
          OKAY = (XRULE.FIDST = FDSTAT$("FID"))
          If Not OKAY Then GoTo 675
       End If
       OKAY = (XRULE.MIDST = RULES(12).MIDST)
       If Not OKAY Then OKAY = (XRULE.MIDST = FDSTAT$("MID"))
       If Not OKAY Then GoTo 675
    End If
    If OKAY Then GoTo 679

675 ' increment rule count
    R = R + 1: If R <= SRuleCNT Then GoTo 650
    SlctCNT = SlctCNT - (SelSET(I) And 1)
    SecSlctCNT = SecSlctCNT + ((SelSET(I) And 2) > 0)
    SelSET(I) = 0
    OKAY = 0

679 ' Next ID
    Return

690 ' Compute SLCTCNT, SecSLCTCNT and LOID, HIID
    SlctCNT = 0: SecSlctCNT = 0: LOID = 32767: HIID = 0
    For I = 1 To FF1Hdr.MAXID
        X = SelSET(I) And 7
        If X > 0 Then
           If (SelSET(I) And 1) Then
              SlctCNT = SlctCNT + 1: SelSET(I) = 1
              Else
              If (SelSET(I) And 6) Then
                 SecSlctCNT = SecSlctCNT + 1: SelSET(I) = 2
              End If
           End If
           HIID = I: If I < LOID Then LOID = I
       End If
    Next I
    'GoSub 230
    PSCOUNT = SlctCNT
    SSCOUNT = SecSlctCNT
    SCNT = PSCOUNT + SSCOUNT
    If SlctCNT + SecSlctCNT = 0 Then GoTo 695
    SAVEWORK = 0

695 ' Finished Applying Rules
    Call ClearMSG
    Exit Sub

700 ' Mark Secondary Selections
    Call FF1GetRec(I): SPREC = FF1Rec.SPOUSE: OLDCH = FF1Rec.OLDCH
    If (SECSLCT And 2) Then
       id = FF1Rec.FID: GoSub 790
       id = FF1Rec.MID: GoSub 790
    End If
    If (SECSLCT And 4) Then
       FID = FF1Rec.FID: xMID = FF1Rec.MID
       If FID > 0 Then
          Call FF1GetRec(FID): id = FF1Rec.OLDCH
          While id > 0
             GoSub 790: Call FF1GetRec(id): id = FF1Rec.FCH
          Wend
       End If
       If xMID > 0 Then
          Call FF1GetRec(xMID): id = FF1Rec.OLDCH
          While id > 0
             GoSub 790: Call FF1GetRec(id): id = FF1Rec.MCH
          Wend
       End If
    End If
    If (SECSLCT And 72) Then
       id = OLDCH
       While id > 0
          Call FF1GetRec(id)
          If (SECSLCT And 8) Then GoSub 790
          If Not (((SECSLCT And 64) = 0) Or (FF1Rec.SPOUSE = 0)) Then
             Call FF3GetRec(FF1Rec.SPOUSE)
             If id = FF3SPOUSE.SP1ID Then
                id = FF3SPOUSE.SP2ID
                Else
                id = FF3SPOUSE.SP1ID
             End If
             GoSub 790
          End If
          If I = FF1Rec.FID Then id = FF1Rec.FCH Else id = FF1Rec.MCH
       Wend
    End If
    If (SECSLCT And 48) Then
       While SPREC <> 0
          Call FF3GetRec(SPREC)
          If I = FF3SPOUSE.SP1ID Then
             id = FF3SPOUSE.SP2ID: SPREC = FF3SPOUSE.SP1NXT
             Else
             id = FF3SPOUSE.SP1ID: SPREC = FF3SPOUSE.SP2NXT
          End If
          If id <> 0 Then
             If (SECSLCT And 16) Then GoSub 790
             If (SECSLCT And 32) Then
                Call FF1GetRec(id)
                id = FF1Rec.FID: GoSub 790
                id = FF1Rec.MID: GoSub 790
             End If
          End If
       Wend
    End If
    Return

790 ' Mark ID as Secondary Selection
    If (id <> 0) And (SelSET(id) = 0) Then
       SecSlctCNT = SecSlctCNT + 1
       SelSET(id) = SelSET(id) Or 4
       If id < LO Then LO = id
       If id > HI Then HI = id
    End If
    Return

End Sub


Sub SelectFILE(OPT$)
    Select Case OPT$
       Case "SAVE": FO = 1: GoTo 820
       Case "LOAD": FO = 2: GoTo 820
       Case "MERG": FO = 3: GoTo 820
       Case "AND": FO = 4: GoTo 820
       Case "EXCL": FO = 5: GoTo 820
    End Select

810 ' Get Select FILE Processing Option
    Call PutMSG("LISTM009")
    If a = 27 Then GoTo 899
    a = Val(CA$)
    If a < 1 Or a > 5 Then Call ErrBEEP(0): GoTo 810
    FO = a
820 ' Process Request
    Select Case FO
       Case 1: GoSub 830
       Case Else: GoSub 840
    End Select
    SAVEWORK = (FO < 3)
    ShowStats
    If ((OPT$ = "") And (a <> 27)) Then GoTo 810
    GoTo 899

830 ' Used to Create SELECT File from SELSET(), RULES()
   Call SWRKOpen("NEW")
   If OKAY Then
      frmFHSWSlct.MousePointer = 11
      'Call PutMSG("XXXXM060")
      Seek #6, Len(SLCTHDR) + 1
      SLCTHDR.SPGM = MID$(ENV.PGMNAME, 4, 4)
      SLCTHDR.SRULCNT = Chr$(SRuleCNT)
      SLCTHDR.SECSLCT = Chr$(SECSLCT)
      If SRuleCNT > 0 Then
         For X = 1 To SRuleCNT: Put #6, , RULES(X): Next X
      End If
      SLCTHDR.LOID = 32767: SLCTHDR.HIID = 0
      X = 0: Y = 0: Z0$ = Chr$(0): xZ$ = " "
      While X < HIID
         X = X + 1: Z = SelSET(X)
         LSet xZ$ = Chr$(Z): Put #6, , xZ$
         If Z > 0 Then
            SLCTHDR.SlctCNT = SLCTHDR.SlctCNT - ((Z And 1) = 1)
            SLCTHDR.SECSCNT = SLCTHDR.SECSCNT - ((Z And 3) = 2)
            If X < SLCTHDR.LOID Then SLCTHDR.LOID = X
            SLCTHDR.HIID = X
         End If
      Wend
      SLCTHDR.SLCTTOT = SLCTHDR.SlctCNT + SLCTHDR.SECSCNT
      SLCTHDR.WFVER = Chr$(2)
      Put #6, 1, SLCTHDR
      SCNT = SLCTHDR.SLCTTOT
      PSCOUNT = SLCTHDR.SlctCNT
      SSCOUNT = SLCTHDR.SECSCNT
   End If
   frmFHSWSlct.MousePointer = 0
   Return

840 ' Select File = 1 Byte for Each ID in .NAM, <>CHR$(0) if selected
   Call SWRKOpen("READ")
   If OKAY Then
      frmFHSWSlct.MousePointer = 11
      SECSLCT = Asc(SLCTHDR.SECSLCT)
      Seek #6, Len(SLCTHDR) + 1
      SRCNT = Asc(SLCTHDR.SRULCNT)
      If SRCNT > 0 Then
         X = SRuleCNT
         While X > 0: RULES(X) = RULES(12): UFLD(X) = 0: X = X - 1: Wend
         For X = 1 To SRCNT
             Get #6, , RULES(X)
             Call ChkRULE(RULES(X), X)
         Next X
         SRuleCNT = SRCNT
      End If
      xZ$ = " ": X = 0: Y = 0
      OHIID = HIID: LOID = 32767: HIID = 0
      HI = SLCTHDR.HIID: If HI > FF1Hdr.MAXID Then HI = FF1Hdr.MAXID
      If FO = 2 Then ReDim SelSET(FF1Hdr.MAXID)
      SlctCNT = 0: SecSlctCNT = 0
      While X < HI
         X = X + 1
         Get #6, , xZ$: Z = Asc(xZ$)
         Select Case FO
            Case 2:                                               'LOAD
            Case 3: Z = SelSET(X) Or Z: If (Z And 1) Then Z = 1   'MERG
            Case 4: Z = SelSET(X) And Z                           'AND
            Case 5: Z = SelSET(X) And (Z = 0)                     'EXCLUDE
         End Select
         SelSET(X) = Z
         If Z > 0 Then
            SlctCNT = SlctCNT - ((Z And 1) = 1)
            SecSlctCNT = SecSlctCNT - ((Z And 3) = 2)
            If X < LOID Then LOID = X
            If X > HIID Then HIID = X
         End If
      Wend
      If FO > 2 Then
         While X < OHIID
            X = X + 1
            Z = SelSET(X)
            If FA = 4 Then
               Z = 0
               Else
               If Z > 0 Then
                  SlctCNT = SlctCNT - ((Z And 1) = 1)
                  SecSlctCNT = SecSlctCNT - ((Z And 3) = 2)
                  If X < LOID Then LOID = X
                  If X > HIID Then HIID = X
               End If
            End If
         Wend
      End If
      SelSET(0) = LOID
      PSCOUNT = SlctCNT
      SSCOUNT = SecSlctCNT
      SCNT = PSCOUNT + SSCOUNT
      frmFHSWSlct.MousePointer = 0
   End If
   Return

899 ' Return to CALLER

End Sub


Sub ChangeEnable()
If ChangeMODE = 1 Then
   cmdCANCEL.Caption = "Cancel"
   cmdCANCEL.DEFAULT = False
   cboBaseType.Enabled = False
   cmdFILE(0).Enabled = False
   cmdFILE(1).Enabled = False
   cmdFILE(2).Enabled = False
   cmdFILE(3).Enabled = False
   cmdRule(0).Caption = "Apply Rules"
   cmdRule(0).DEFAULT = True
   cmdRule(1).Enabled = False
   cmdRule(2).Enabled = False
   SearchRule.Visible = True
   frmFHSWSlct.Height = 5150
   SearchRule.Enabled = True
   ChangeMODE = 3
   Else
   If ChangeMODE = 2 Then
   cmdCANCEL.Caption = "Cancel"
   cmdCANCEL.DEFAULT = False
   cboBaseType.Enabled = False
   cmdFILE(0).Enabled = False
   cmdFILE(1).Enabled = False
   cmdFILE(2).Enabled = False
   cmdFILE(3).Enabled = False
   cmdRule(0).Caption = "Add Selections"
   cmdRule(0).DEFAULT = True
   cmdRule(1).Enabled = False
   cmdRule(2).Enabled = False
   SearchRule.Visible = False
   SecondSelect.TOP = 1200
   SecondSelect.Visible = True
   frmFHSWSlct.Height = 2900
   SecondSelect.Enabled = True
   ChangeMODE = 4
   Else
   cmdRule(0).DEFAULT = False
   cmdCANCEL.Caption = "Return"
   cmdCANCEL.DEFAULT = True
   cboBaseType.Enabled = True
   cmdFILE(0).Enabled = True
   cmdFILE(1).Enabled = True
   cmdFILE(2).Enabled = True
   cmdFILE(3).Enabled = True
   cmdRule(0).Caption = "Search"
   cmdRule(1).Enabled = True
   cmdRule(2).Enabled = True
   SearchRule.Enabled = False
   SecondSelect.Visible = False
   frmFHSWSlct.Height = 1700
   ChangeMODE = 0
   End If
End If
End Sub






Sub ShowStats()
txtCount(0) = SFILL$(PSCOUNT, 6)
txtCount(1) = SFILL$(SSCOUNT, 6)
End Sub


Private Sub cboBaseType_Click()
If cboBaseType.Enabled Then
   If cboBaseType.ListIndex = 4 Then
      Call SelectFILE("LOAD")
      Else
      Call BuildBASE(cboBaseType.ListIndex)
   End If
   Call ShowStats
End If
End Sub



Private Sub cmdCANCEL_Click()
If ChangeMODE <> 0 Then
   ChangeEnable
   Else
   Unload frmFHSWSlct
End If
End Sub


Private Sub cmdFILE_Click(Index As Integer)
Select Case Index
   Case 0: Call SelectFILE("MERG")
   Case 1: Call SelectFILE("EXCL")
   Case 2: Call SelectFILE("AND")
   Case 3: Call SelectFILE("SAVE")
End Select
End Sub

Private Sub cmdRCHG_Click(Index As Integer)
Select Case Index
   Case 0: ShowRULE (12)
   Case 1: SaveRule (11)
   Case 2: ShowRULE (11)
End Select
End Sub

Private Sub cmdRule_Click(Index As Integer)
Select Case Index
   Case 0: ' Change Rules
           Select Case ChangeMODE
             Case 0: ' Change Rules
                     ChangeMODE = 1
                     ChangeEnable
                     RuleNUM = 1: Call ChkRULE(RULES(1), 1)
                     If SRuleCNT = 1 And UFLD(1) = 0 Then SRuleCNT = 0
                     ShowRULE (RuleNUM)
                     nxtRULE.VALUE = 1: nxtRULE.Max = 9
                     'nxtRULE.Max = SRuleCNT + 1
                     txtRule(0).SetFocus
             Case 3: ' Apply Rules
                     SlecType = 5
                     Call SaveRule(RuleNUM)
                     Call SelectSET("AppRule")
                     Call ShowStats
                     ChangeEnable
             Case 4: ' Apply Secondary Selections
                     SECSLCT = 0
                     For X = 0 To 5
                         If chk2ndSelect(X) = 1 Then
                            SECSLCT = SECSLCT + 2 ^ (X + 1)
                         End If
                     Next X
                     Call SelectSET("App2ndS")
                     SlecType = 5
                     ShowStats
                     ChangeEnable
           End Select
   Case 1: ' Change Secondary Selections
           ChangeMODE = 2
           ChangeEnable
           For X = 0 To 5
               chk2ndSelect(X) = -((SECSLCT And 2 ^ (X + 1)) > 0)
           Next X
           chk2ndSelect(0).SetFocus
   Case 2: ' Show List of Selected Records
           frmFHSWList.ShowNames (-1)
           ShowStats
End Select
End Sub


Private Sub Form_Load()
If FormLoaded(2) And SpouseForm Then
   frmFHSWSpse.Enabled = False
   Else
   If FormLoaded(2) And UpdtForm Then
      frmFHSWUPDT.Enabled = False
      Else
      If FormLoaded(1) And RelWForm Then
         frmFHSWRelW.Enabled = False
         Else
         If FormLoaded(1) And RPgmForm Then
            frmFHSWRPgm.Enabled = False
         End If
      End If
   End If
End If
FormLoaded(1) = FormLoaded(1) Or SlctForm
' Set Default Form Location on Screen
frmHEIGHT = 1700
frmWIDTH = frmFHSWSlct.WIDTH
t1 = GetSetting("FHSWIN", "frmSLCT", "TOP", 0)
If t1 = 0 Then
   t00 = mnuFHSWMain.TOP: w00 = mnuFHSWMain.WIDTH
   l00 = mnuFHSWMain.Left
   t1 = (t00 + 1960)
   l1 = l00 + 1815: If l1 < 0 Then l1 = 1
   Else
   t1 = 1440: l1 = 1440
End If
frmFHSWSlct.TOP = GetSetting("FHSWIN", "frmSLCT", "Top", t1)
frmFHSWSlct.Left = GetSetting("FHSWIN", "frmSLCT", "Left", l1)
frmFHSWSlct.Height = frmHEIGHT
cboBaseType.Enabled = False
cboBaseType.ListIndex = SBASE
cboBaseType.Enabled = True
SearchRule.Visible = False
SecondSelect.Visible = False
If ENV.FFOPEN Then Call FamCLOSE
If RptCODE <> "LIST" Then Call MsgGrpLOAD("LIST", X)
Call ShowStats
ChangeMODE = 0
ReDim RULES(12) As SlctFileRULE, UFLD(10)
XRULE.LOID = 1: XRULE.HIID = 32767
XRULE.HIBY = 9999: XRULE.HIBMD = 9999
XRULE.HIDY = 9999: XRULE.HIDMD = 9999
XRULE.BDST = " ": XRULE.DDST = " "
XRULE.FIDST = " ": XRULE.MIDST = " "
XRULE.SEX = " "
LSet XRULE.GVNM = " ": LSet XRULE.SRNM = " "
LSet XRULE.BPLC = " ": LSet XRULE.DPLC = " "
XRULE.YORN = 0
ReDim HTxtRule(20)
For I = 1 To 12: RULES(I) = XRULE: Next I
nxtRULE.Enabled = False
SRuleCNT = 1: nxtRULE.Min = 1
nxtRULE.Enabled = True
ShowRULE (1)
End Sub


Private Sub Form_Unload(Cancel As Integer)
If RptCODE <> "LIST" Then Call MsgGrpFREE("LIST")
FormLoaded(1) = FormLoaded(1) And (255 - SlctForm)
SaveSetting "FHSWIN", "frmSLCT", "Top", frmFHSWSlct.TOP
SaveSetting "FHSWIN", "frmSLCT", "Left", frmFHSWSlct.Left
If FormLoaded(2) And SpouseForm Then
   Call FAMOpen("WRITE")
   frmFHSWSpse.Enabled = True
   frmFHSWSpse.SetFocus
   Else
   If FormLoaded(2) And UpdtForm Then
      Call FAMOpen("WRITE")
      frmFHSWUPDT.Enabled = True
      frmFHSWUPDT.SetFocus
      Else
      Call FamCLOSE
      If FormLoaded(1) And RPgmForm Then
         frmFHSWRPgm.Enabled = True
         frmFHSWRPgm.SetFocus
         Else
         If FormLoaded(1) And RelWForm Then
            frmFHSWRelW.Enabled = True
            frmFHSWRelW.SetFocus
         End If
      End If
   End If
End If
End Sub




Private Sub nxtRULE_Change()
If nxtRULE.Enabled Then
   Call SaveRule(RuleNUM)
   If RuleNUM = SRuleCNT And UFLD(RuleNUM) = 0 Then SRuleCNT = SRuleCNT - 1
   If nxtRULE.VALUE <= SRuleCNT + 1 Then
      RuleNUM = nxtRULE.VALUE
      Call ShowRULE(RuleNUM)
      Else
      nxtRULE.VALUE = SRuleCNT + 1
   End If
End If
End Sub


Private Sub txtRule_GotFocus(Index As Integer)
HOLDITEM = RTrim$(txtRule(Index).TEXT)
Select Case Index
   Case 0, 1: xFT = 2
   Case 2, 15, 16, 17, 18, 19, 20: xFT = 5
   Case 7, 8, 11, 12: xFT = 3
   Case Else: xFT = 1
End Select
txtRule(Index).SelStart = 0
If (xFT > 2) Then
   txtRule(Index).SelLength = 0
   Else
   txtRule(Index).SelLength = Len(RTrim$(txtRule(Index).TEXT))
End If
LastFLD = Index
Exit Sub
txtRule(Index).SelStart = 0
txtRule(Index).SelLength = Len(txtRule(Index).TEXT)
End Sub

Private Sub txtRule_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If xFT > 1 Then
   XVALUE$ = txtRule(Index).TEXT
   xSS = txtRule(Index).SelStart
   xSL = txtRule(Index).SelLength
   Call FldUpdate(KeyCode, Shift, xFT, xSS, xSL, XVALUE$)
   txtRule(Index).TEXT = XVALUE$
   txtRule(Index).SelStart = xSS
   txtRule(Index).SelLength = xSL
End If
Exit Sub
End Sub


Private Sub txtRule_KeyPress(Index As Integer, KeyAscii As Integer)
If (xFT > 1) Then KeyAscii = 0: Exit Sub
Select Case Index
   Case 2, 15, 16, 17, 18, 19, 20:
       KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Select
End Sub


