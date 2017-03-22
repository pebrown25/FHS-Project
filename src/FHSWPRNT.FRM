VERSION 4.00
Begin VB.Form frmFHSWPrinter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FHSWPRNT - Printer/Page Setup Parameters"
   ClientHeight    =   6345
   ClientLeft      =   7500
   ClientTop       =   4425
   ClientWidth     =   7095
   Height          =   6750
   KeyPreview      =   -1  'True
   Left            =   7440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   Top             =   4080
   Width           =   7215
   Begin VB.ComboBox cboORIENT 
      Height          =   315
      ItemData        =   "FHSWPRNT.frx":0000
      Left            =   4680
      List            =   "FHSWPRNT.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   90
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Frame framePageParm 
      Caption         =   "Page Parameters"
      Height          =   1815
      Left            =   0
      TabIndex        =   62
      Top             =   2400
      Width           =   7095
      Begin VB.CommandButton cmdRESHOW 
         Caption         =   "Reshow"
         Height          =   255
         Left            =   6240
         TabIndex        =   91
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox PPValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   11
         Left            =   5880
         MultiLine       =   -1  'True
         TabIndex        =   88
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox PPValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   10
         Left            =   3120
         MultiLine       =   -1  'True
         TabIndex        =   84
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox PPValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   9
         Left            =   5880
         MultiLine       =   -1  'True
         TabIndex        =   80
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox PPValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   3120
         MultiLine       =   -1  'True
         TabIndex        =   77
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox PPValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   5880
         MultiLine       =   -1  'True
         TabIndex        =   73
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox PPValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   3120
         MultiLine       =   -1  'True
         TabIndex        =   71
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox cboUNIT 
         Height          =   315
         ItemData        =   "FHSWPRNT.frx":0033
         Left            =   2280
         List            =   "FHSWPRNT.frx":0040
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   195
         Width           =   735
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Single Side"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   23
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Form Feed"
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   24
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Page Wait"
         Height          =   255
         Index           =   3
         Left            =   4560
         TabIndex        =   25
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox PPValue 
         Height          =   285
         Index           =   0
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   16
         Top             =   195
         Width           =   495
      End
      Begin VB.TextBox PPValue 
         Height          =   285
         Index           =   2
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   19
         Top             =   560
         Width           =   495
      End
      Begin VB.TextBox PPValue 
         Height          =   285
         Index           =   1
         Left            =   5040
         MaxLength       =   4
         TabIndex        =   18
         Top             =   195
         Width           =   495
      End
      Begin VB.TextBox PPValue 
         Height          =   285
         Index           =   3
         Left            =   5040
         MaxLength       =   4
         TabIndex        =   20
         Top             =   555
         Width           =   495
      End
      Begin VB.TextBox PPValue 
         Height          =   285
         Index           =   4
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   21
         Top             =   920
         Width           =   495
      End
      Begin VB.TextBox PPValue 
         Height          =   285
         Index           =   5
         Left            =   5040
         MaxLength       =   4
         TabIndex        =   22
         Top             =   915
         Width           =   495
      End
      Begin VB.Label PPLbl 
         Caption         =   "Chars)"
         Height          =   255
         Index           =   19
         Left            =   6360
         TabIndex        =   89
         Top             =   960
         Width           =   495
      End
      Begin VB.Label PPLbl 
         Caption         =   "("
         Height          =   255
         Index           =   18
         Left            =   5760
         TabIndex        =   87
         Top             =   960
         Width           =   135
      End
      Begin VB.Label PPLbl 
         Caption         =   "in"
         Height          =   255
         Index           =   17
         Left            =   5520
         TabIndex        =   86
         Top             =   960
         Width           =   255
      End
      Begin VB.Label PPLbl 
         Caption         =   "Chars)        Outside="
         Height          =   255
         Index           =   16
         Left            =   3600
         TabIndex        =   85
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label PPLbl 
         Caption         =   "("
         Height          =   255
         Index           =   15
         Left            =   3000
         TabIndex        =   83
         Top             =   960
         Width           =   135
      End
      Begin VB.Label PPLbl 
         Caption         =   "in"
         Height          =   255
         Index           =   14
         Left            =   2760
         TabIndex        =   82
         Top             =   960
         Width           =   255
      End
      Begin VB.Label PPLbl 
         Caption         =   "Lines)"
         Height          =   255
         Index           =   12
         Left            =   6360
         TabIndex        =   81
         Top             =   600
         Width           =   495
      End
      Begin VB.Label PPLbl 
         Caption         =   "("
         Height          =   255
         Index           =   11
         Left            =   5760
         TabIndex        =   79
         Top             =   600
         Width           =   135
      End
      Begin VB.Label PPLbl 
         Caption         =   "Lines)         Bottom="
         Height          =   255
         Index           =   9
         Left            =   3600
         TabIndex        =   78
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label PPLbl 
         Caption         =   "("
         Height          =   255
         Index           =   8
         Left            =   3000
         TabIndex        =   76
         Top             =   600
         Width           =   135
      End
      Begin VB.Label PPLbl 
         Caption         =   "in"
         Height          =   255
         Index           =   7
         Left            =   2760
         TabIndex        =   75
         Top             =   600
         Width           =   255
      End
      Begin VB.Label PPLbl 
         Caption         =   "Lines)"
         Height          =   255
         Index           =   5
         Left            =   6360
         TabIndex        =   74
         Top             =   240
         Width           =   495
      End
      Begin VB.Label PPLbl 
         Caption         =   "("
         Height          =   255
         Index           =   4
         Left            =   5760
         TabIndex        =   72
         Top             =   240
         Width           =   135
      End
      Begin VB.Label PPLbl 
         Caption         =   "("
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   70
         Top             =   240
         Width           =   135
      End
      Begin VB.Label PPLbl 
         Caption         =   "Options:"
         Height          =   255
         Index           =   20
         Left            =   840
         TabIndex        =   69
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label PPLbl 
         Caption         =   "Form Size:    Width="
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   68
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label PPLbl 
         Caption         =   "Margins:       Top="
         Height          =   255
         Index           =   6
         Left            =   840
         TabIndex        =   67
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label PPLbl 
         Caption         =   "Chars)        Length="
         Height          =   255
         Index           =   2
         Left            =   3600
         TabIndex        =   66
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label PPLbl 
         Caption         =   "in"
         Height          =   255
         Index           =   3
         Left            =   5520
         TabIndex        =   65
         Top             =   240
         Width           =   255
      End
      Begin VB.Label PPLbl 
         Caption         =   "in"
         Height          =   255
         Index           =   10
         Left            =   5520
         TabIndex        =   64
         Top             =   600
         Width           =   255
      End
      Begin VB.Label PPLbl 
         Caption         =   "Inside="
         Height          =   255
         Index           =   13
         Left            =   1680
         TabIndex        =   63
         Top             =   960
         Width           =   615
      End
   End
   Begin VB.Frame frameSYMBOLS 
      Height          =   1935
      Left            =   0
      TabIndex        =   49
      Top             =   4320
      Visible         =   0   'False
      Width           =   7095
      Begin VB.CommandButton cmdHEX 
         Caption         =   "ASCII"
         Height          =   315
         Left            =   120
         TabIndex        =   26
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton cmdNEXT 
         Caption         =   ">"
         Height          =   315
         Left            =   6600
         TabIndex        =   39
         Top             =   1560
         Width           =   375
      End
      Begin VB.CommandButton cmdPREV 
         Caption         =   "<"
         Height          =   315
         Left            =   6240
         TabIndex        =   38
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox LSFValue 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   5
         Left            =   1560
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   31
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox LSFValue 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   2
         Left            =   2880
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   28
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox LSFValue 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   10
         Left            =   2880
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   35
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox LSFValue 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   12
         Left            =   5520
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   37
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox LSFValue 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   11
         Left            =   4200
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   36
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox LSFValue 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   9
         Left            =   1560
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   34
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox LSFValue 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   8
         Left            =   5520
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   33
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox LSFValue 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   7
         Left            =   4200
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   32
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox LSFValue 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   4
         Left            =   5520
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   30
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox LSFValue 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   3
         Left            =   4200
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   29
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox LSFValue 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   1560
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   27
         Top             =   480
         Width           =   495
      End
      Begin VB.Label LSFLbl 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "+"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS LineDraw"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   5160
         TabIndex        =   61
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label LSFLbl 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "+"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS LineDraw"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   11
         Left            =   3840
         TabIndex        =   60
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label LSFLbl 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "-"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS LineDraw"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   2520
         TabIndex        =   59
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label LSFLbl 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "+"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS LineDraw"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   1200
         TabIndex        =   58
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label LSFLbl 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "-"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS LineDraw"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   5160
         TabIndex        =   57
         Top             =   960
         Width           =   255
      End
      Begin VB.Label LSFLbl 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "+"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS LineDraw"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   3840
         TabIndex        =   56
         Top             =   960
         Width           =   255
      End
      Begin VB.Label LSFLbl 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "¦"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS LineDraw"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   1200
         TabIndex        =   55
         Top             =   960
         Width           =   255
      End
      Begin VB.Label LSFLbl 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "¦"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS LineDraw"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   5160
         TabIndex        =   54
         Top             =   480
         Width           =   255
      End
      Begin VB.Label LSFLbl 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "+"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS LineDraw"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   3840
         TabIndex        =   53
         Top             =   480
         Width           =   255
      End
      Begin VB.Label LSFLbl 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "-"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS LineDraw"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   52
         Top             =   480
         Width           =   255
      End
      Begin VB.Label LSFLbl 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "+"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS LineDraw"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   51
         Top             =   480
         Width           =   255
      End
      Begin VB.Label LSFLbl 
         Alignment       =   2  'Center
         Caption         =   "Line Drawing Symbols"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   50
         Top             =   120
         Width           =   2655
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdLOAD 
      Caption         =   "Load"
      Height          =   315
      Left            =   6240
      TabIndex        =   42
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton cmdCHGSYMB 
      Caption         =   "Show"
      Height          =   315
      Left            =   3600
      TabIndex        =   15
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton cmdCANCEL 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   6240
      TabIndex        =   1
      Top             =   500
      Width           =   735
   End
   Begin VB.ComboBox cboLineSymbols 
      Height          =   315
      ItemData        =   "FHSWPRNT.frx":0050
      Left            =   1440
      List            =   "FHSWPRNT.frx":0060
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton cmdDEL 
      Caption         =   "Del"
      Height          =   315
      Left            =   5520
      TabIndex        =   9
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton cmdADD 
      Caption         =   "Add"
      Height          =   315
      Left            =   4800
      TabIndex        =   8
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton cmdPRINT 
      Caption         =   "Report"
      Height          =   315
      Left            =   6240
      TabIndex        =   44
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton cmdSAVE 
      Caption         =   "Save"
      Height          =   315
      Left            =   6240
      TabIndex        =   40
      Top             =   1080
      Width           =   735
   End
   Begin VB.CheckBox chkDefault 
      Caption         =   "Default Setup"
      Height          =   315
      Left            =   3360
      TabIndex        =   7
      Top             =   960
      Width           =   1335
   End
   Begin VB.ComboBox cboPRTDest 
      Height          =   315
      ItemData        =   "FHSWPRNT.frx":00A5
      Left            =   1440
      List            =   "FHSWPRNT.frx":00B2
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox PRTDesc 
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
      Left            =   3360
      MaxLength       =   50
      TabIndex        =   5
      Top             =   600
      Width           =   2655
   End
   Begin VB.ComboBox PRTSetup 
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton cmdRETURN 
      Caption         =   "Return"
      Default         =   -1  'True
      Height          =   315
      Left            =   6240
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.ComboBox cboFontStyle 
      Height          =   315
      ItemData        =   "FHSWPRNT.frx":00CD
      Left            =   4800
      List            =   "FHSWPRNT.frx":00DD
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox cboFONT 
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1440
      Width           =   1815
   End
   Begin VB.ComboBox cboFontSize 
      Height          =   315
      ItemData        =   "FHSWPRNT.frx":0106
      Left            =   3720
      List            =   "FHSWPRNT.frx":0125
      Sorted          =   -1  'True
      TabIndex        =   12
      Text            =   "cboFontSize"
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdSYMBOLS 
      Caption         =   "@"
      Height          =   315
      Left            =   3240
      TabIndex        =   11
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton cmdCHGDRVR 
      Caption         =   "Change"
      Height          =   315
      Left            =   5280
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox PrinterName 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   3735
   End
   Begin VB.Line Line2 
      X1              =   6120
      X2              =   7080
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line6 
      X1              =   0
      X2              =   7080
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label 
      Caption         =   "Line Symbols:"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   48
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Line Line4 
      X1              =   6120
      X2              =   6120
      Y1              =   0
      Y2              =   2280
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   6120
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label 
      Caption         =   "Dest:"
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   47
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label 
      Caption         =   "Printer Setup:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   46
      Top             =   600
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   7440
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSGrid.Grid Grid1 
      Height          =   1455
      Left            =   7440
      TabIndex        =   45
      Top             =   1800
      Width           =   2295
      _Version        =   65536
      _ExtentX        =   4048
      _ExtentY        =   2566
      _StockProps     =   77
      BackColor       =   16777215
      Enabled         =   0   'False
      Rows            =   17
      Cols            =   17
      ScrollBars      =   0
      MouseIcon       =   "FHSWPRNT.frx":016E
   End
   Begin VB.Label Label 
      Caption         =   "Font:"
      Height          =   375
      Index           =   3
      Left            =   960
      TabIndex        =   43
      Top             =   1440
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   6120
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label 
      Caption         =   "Windows Driver:"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   41
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmFHSWPrinter"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
DefInt A-Z

Dim TypeSymbol, SymbolREP, SymbolBASE
Dim LSFValueFont As Variant
Dim HLSFValue As Variant, HLSFVIndex
Dim HPPValue As Variant, HPPIndex
Private Sub BuildFonTable()
cboFONT.Clear: xfont = 0
For X = 0 To Printer.FontCount - 1
   cboFONT.AddItem Printer.Fonts(X)
   cboFONT.ItemData(cboFONT.NewIndex) = X
   If Printer.Font = Printer.Fonts(X) Then
      xfont = X
   End If
Next X
X = 0: FOUND = 0
While (X < cboFONT.ListCount) And (Not FOUND)
   FOUND = (cboFONT.ItemData(X) = xfont)
   If Not FOUND Then X = X + 1
Wend
cboFONT.ListIndex = X
End Sub


Private Function StdValue(v As Variant)
Select Case cboUNIT.ListIndex
   Case 0: xX# = v
   Case 1: xX# = (v / 2.54)
   Case 2: xX# = (v / 25.4)
End Select
If xX# > 327 Then xX# = 327
StdValue = Int(xX# * 100)
End Function

Private Function UnitValue(X) As Variant
Select Case cboUNIT.ListIndex
   Case 0: xX# = X
   Case 1: xX# = (X * 2.54)
   Case 2: xX# = (X * 25.4)
End Select
UnitValue = Int(xX#) / 100
End Function

Private Function Dec2HEX$(X)
    x1 = X \ 16
    X2 = X - 16 * x1
    Dec2HEX$ = MID$("0123456789ABCDEF", x1 + 1, 1) + MID$("0123456789ABCDEF", X2 + 1, 1)
End Function

Private Function HEX2Dec(xX$)
If Len(xX$) <> 2 Then
   HEX2Dec = 0
   Else
   X = 0: xA$ = MID$(xX$, 1, 1): GoSub HC2D
   Y = 16 * X: xA$ = MID$(xX$, 2, 1): GoSub HC2D
   HEX2Dec = X + Y
End If
Exit Function

HC2D:
   X = Asc(xA$)
   If (X < 48) Then
      X = 0
      Else
      If X < 58 Then
         X = X - 48
         Else
         If X < 65 Then
            X = 0
            Else
            If X < 71 Then
               X = X - 55
               Else
               X = 0
            End If
         End If
      End If
    End If
    Return
   
   X = (Asc(xA$) - 48)
   
   
End Function


Private Sub SaveSetup()
If chkDefault.VALUE = 1 Then
   PDFHdr.DEFAULT = ENV.PSETUP
End If
PSETUPS(ENV.PSETUP) = PDFSetup
PRINTERS$(ENV.PSETUP) = PrinterName
hfont$ = Printer.Font
Printer.Font = PFONT$(ENV.PSETUP)
If hfont$ <> Printer.Font Then
   PFONT$(ENV.PSETUP) = hfont$
End If
Printer.Font = hfont$
End Sub


Private Sub ShowSetup()
On Local Error GoTo OrientERR
'cboORIENT.ListIndex = -((PDFSetup.OPTIONS And 32) > 0)
X = (PDFSetup.OPTIONS And 48) / 16
If (X < 1) Or (X > 2) Then X = 1
Printer.Orientation = X
cboORIENT.ListIndex = Printer.Orientation - 1
On Local Error GoTo 0
Printer.Font = PFONT$(ENV.PSETUP)
If Printer.Font <> cboFONT.TEXT Then
   cboFONT.TEXT = Printer.Font
End If
cboFontSize = PDFSetup.FONTSIZE
cboFontStyle.ListIndex = PDFSetup.FONTSTYL
PRTDesc.TEXT = PDFSetup.DESC
Select Case RTrim$(PDFSetup.DEST)
   Case "SCRN:": cboPRTDest.ListIndex = 0
   Case "FILE": cboPRTDest.ListIndex = 2
   Case Else: cboPRTDest.ListIndex = 1
End Select
Select Case RTrim$(PDFSetup.SYMBOLS)
   Case "TYPEWRIT": cboLineSymbols.ListIndex = 0
   Case "CP437": cboLineSymbols.ListIndex = 1
   Case "CP850": cboLineSymbols.ListIndex = 2
   Case Else: cboLineSymbols.ListIndex = 3
End Select
chkDefault.VALUE = -(ENV.PSETUP = PDFHdr.DEFAULT)
If frameSYMBOLS.Visible Then
   ShowSymbols
   Else
   ShowPageParms
End If
Exit Sub

OrientERR:
  cboORIENT.ListIndex = 2
  Resume Next

End Sub

Private Sub ShowSymbols()
Select Case TypeSymbol
   Case 0: LSFLbl(0) = "Line Drawing Symbols"
   Case 1: LSFLbl(0) = "Regular Box Drawing Symbols"
   Case 2: LSFLbl(0) = "Hilited Box Drawing Symbols"
End Select
GoSub ShowLbls
GoSub ShowValues
If Not frameSYMBOLS.Visible Then
   frameSYMBOLS.Left = 0
   frameSYMBOLS.TOP = 2280
   framePageParm.Visible = False
   framePageParm.Enabled = False
   frameSYMBOLS.Enabled = True
   frameSYMBOLS.Visible = True
   cmdCHGSYMB.Caption = "Hide"
   cmdCANCEL.Cancel = False
   cmdRETURN.DEFAULT = True
   cmdHEX.SetFocus
End If
Exit Sub

ShowLbls:
Select Case TypeSymbol
   Case 0: LSFLbl(1) = Chr$(218)
           LSFLbl(2) = Chr$(193)
           LSFLbl(3) = Chr$(191)
           LSFLbl(4) = Chr$(179)
           LSFLbl(5) = Chr$(180)
           LSFLbl(7) = Chr$(195)
           LSFLbl(8) = Chr$(196)
           LSFLbl(9) = Chr$(192)
           LSFLbl(10) = Chr$(194)
           LSFLbl(11) = Chr$(217)
           LSFLbl(12) = Chr$(197)
           LSFLbl(12).Visible = True
           LSFValue(12).Visible = True
   Case 1: LSFLbl(1) = Chr$(218)
           LSFLbl(2) = Chr$(193)
           LSFLbl(3) = Chr$(191)
           LSFLbl(4) = Chr$(179)
           LSFLbl(5) = Chr$(180)
           LSFLbl(7) = Chr$(195)
           LSFLbl(8) = Chr$(196)
           LSFLbl(9) = Chr$(192)
           LSFLbl(10) = Chr$(194)
           LSFLbl(11) = Chr$(217)
           LSFLbl(12).Visible = False
           LSFValue(12).Visible = False
   Case 2: LSFLbl(1) = Chr$(201)
           LSFLbl(2) = Chr$(207)
           LSFLbl(3) = Chr$(187)
           LSFLbl(4) = Chr$(186)
           LSFLbl(5) = Chr$(182)
           LSFLbl(7) = Chr$(199)
           LSFLbl(8) = Chr$(205)
           LSFLbl(9) = Chr$(200)
           LSFLbl(10) = Chr$(209)
           LSFLbl(11) = Chr$(188)
           LSFLbl(12).Visible = False
           LSFValue(12).Visible = False
End Select
Return

ShowValues:
Z = 0
For X = 1 To 12
   If X <> 6 Then
      Z = Z + 1: Y = 0
      If Not ((Z > 10) And (TypeSymbol > 0)) Then
         Select Case TypeSymbol
            Case 0: Y = Asc(MID$(PDFSetup.SET1, Z, 1))
            Case 1: Y = Asc(MID$(PDFSetup.SET2, Z, 1))
            Case 2: Y = Asc(MID$(PDFSetup.SET3, Z, 1))
         End Select
         Select Case SymbolREP
            Case 0: LSFValue(X) = Y
            Case 1: LSFValue(X) = Dec2HEX$(Y)
            Case 2: LSFValue(X) = Chr$(Y)
                    LSFValue(X).Font = Printer.Font
         End Select
      End If
   End If
Next X
frameSYMBOLS.Refresh
Return
End Sub
Private Sub BuildPRTSetup()
PRTSetup.Clear
X = PSUTABLE.FT: Y = 0
While X > 0
    PRTSetup.AddItem PSETUPS(X).NAME
    PRTSetup.ItemData(PRTSetup.NewIndex) = X
    X = PSUCHN(X).FWD
Wend
X = 0: FOUND = 0
While (X < PRTSetup.ListCount) And (Not FOUND)
   FOUND = (PRTSetup.ItemData(X) = ENV.PSETUP)
   If Not FOUND Then X = X + 1
Wend
If FOUND Then PRTSetup.ListIndex = X
End Sub




Private Sub ShowPageParms()
PPvalue(0) = UnitValue(PDFSetup.FWIDTH)
PPvalue(1) = UnitValue(PDFSetup.FLENGTH)
PPvalue(2) = UnitValue(PDFSetup.TMARGIN)
PPvalue(3) = UnitValue(PDFSetup.BMARGIN)
PPvalue(4) = UnitValue(PDFSetup.IMARGIN)
PPvalue(5) = UnitValue(PDFSetup.OMARGIN)
ChkOPT(1).VALUE = ((PDFSetup.OPTIONS And 1) / 1)
ChkOPT(2).VALUE = ((PDFSetup.OPTIONS And 2) / 2)
ChkOPT(3).VALUE = ((PDFSetup.OPTIONS And 4) / 4)
If frameSYMBOLS.Visible Then
   frameSYMBOLS.Visible = False
   frameSYMBOLS.Enabled = False
   framePageParm.Enabled = True
   framePageParm.Visible = True
   cmdCHGSYMB.Caption = "Show"
   cmdCANCEL.Cancel = True
   cmdRETURN.DEFAULT = True
   cmdRETURN.SetFocus
End If
End Sub

Private Sub cboFONT_Click()
Printer.Font = cboFONT.TEXT
Printer.Font.SIZE = PDFSetup.FONTSIZE
If Printer.Font.SIZE <> cboFontSize.TEXT Then
   cboFontSize.TEXT = Printer.Font.SIZE
   Else
   cboFontSize_Click
End If
cboFontSize.TEXT = Printer.Font.SIZE
Printer.Font.Bold = ((PDFSetup.FONTSTYL And 1) > 0)
Printer.Font.Italic = ((PDFSetup.FONTSTYL And 2) > 0)
cboFontSize = Printer.Font.SIZE
If Grid1.Visible = True Then
   Grid1.Font = Printer.Font
End If
If frameSYMBOLS.Visible And (SymbolREP = 2) Then
   For X = 1 To 12
       If X <> 6 Then LSFValue(X).Font = Printer.Font
   Next X
End If
End Sub

Private Sub cboFontSize_Change()
   X = Val(cboFontSize.TEXT)
   If X > 0 Then
      Printer.Font.SIZE = X
      cboFontSize.TEXT = Printer.Font.SIZE
      PPvalue(6) = CharCount(PDFSetup.FWIDTH)
      PPvalue(7) = LineCount(PDFSetup.FLENGTH)
      PPvalue(8) = LineCount(PDFSetup.TMARGIN)
      PPvalue(9) = LineCount(PDFSetup.BMARGIN)
      PPvalue(10) = CharCount(PDFSetup.IMARGIN)
      PPvalue(11) = CharCount(PDFSetup.OMARGIN)
   End If
End Sub

Private Sub cboFontSize_Click()
'If Printer.Font.SIZE <> cboFontSize Then
   If cboFontSize.ListIndex > 0 Then
      PDFSetup.FONTSIZE = Val(cboFontSize.List(cboFontSize.ListIndex))
   End If
   X = Val(cboFontSize.TEXT)
   If X > 0 Then
      Printer.Font.SIZE = X
      cboFontSize.TEXT = Printer.Font.SIZE
      PPvalue(6) = CharCount(PDFSetup.FWIDTH)
      PPvalue(7) = LineCount(PDFSetup.FLENGTH)
      PPvalue(8) = LineCount(PDFSetup.TMARGIN)
      PPvalue(9) = LineCount(PDFSetup.BMARGIN)
      PPvalue(10) = CharCount(PDFSetup.IMARGIN)
      PPvalue(11) = CharCount(PDFSetup.OMARGIN)
   End If

'End If
End Sub




Private Sub cboFontStyle_Click()
PDFSetup.FONTSTYL = cboFontStyle.ListIndex
Printer.Font.Bold = ((PDFSetup.FONTSTYL And 1) > 0)
Printer.Font.Italic = ((PDFSetup.FONTSTYL And 2) > 0)
End Sub

Private Sub cboLineSymbols_Click()
Select Case cboLineSymbols.ListIndex
   Case 0: PDFSetup.SYMBOLS = "TYPEWRIT"
           PDFSetup.SET1 = Chr$(43) + Chr$(43) + Chr$(43) + Chr$(33) + Chr$(43) + Chr$(43) + Chr$(45) + Chr$(43) + Chr$(43) + Chr$(43) + Chr$(43)
           PDFSetup.SET2 = Chr$(43) + Chr$(43) + Chr$(43) + Chr$(33) + Chr$(43) + Chr$(43) + Chr$(45) + Chr$(43) + Chr$(43) + Chr$(43)
           PDFSetup.SET3 = Chr$(43) + Chr$(43) + Chr$(43) + Chr$(33) + Chr$(43) + Chr$(43) + Chr$(45) + Chr$(43) + Chr$(43) + Chr$(43)
           SymbolBASE = 0
           If frameSYMBOLS.Visible Then ShowSymbols
   Case 1: PDFSetup.SYMBOLS = "CP437"
           PDFSetup.SET1 = Chr$(218) + Chr$(193) + Chr$(191) + Chr$(179) + Chr$(180) + Chr$(195) + Chr$(196) + Chr$(192) + Chr$(194) + Chr$(217) + Chr$(197)
           PDFSetup.SET2 = Chr$(218) + Chr$(193) + Chr$(191) + Chr$(179) + Chr$(180) + Chr$(195) + Chr$(196) + Chr$(192) + Chr$(194) + Chr$(217)
           PDFSetup.SET3 = Chr$(201) + Chr$(207) + Chr$(187) + Chr$(186) + Chr$(182) + Chr$(199) + Chr$(205) + Chr$(200) + Chr$(209) + Chr$(188)
           SymbolBASE = 1
           If frameSYMBOLS.Visible Then ShowSymbols
   Case 2: PDFSetup.SYMBOLS = "CP850"
           PDFSetup.SET1 = Chr$(218) + Chr$(193) + Chr$(191) + Chr$(179) + Chr$(180) + Chr$(195) + Chr$(196) + Chr$(192) + Chr$(194) + Chr$(217) + Chr$(197)
           PDFSetup.SET2 = Chr$(218) + Chr$(193) + Chr$(191) + Chr$(179) + Chr$(180) + Chr$(195) + Chr$(196) + Chr$(192) + Chr$(194) + Chr$(217)
           PDFSetup.SET3 = Chr$(218) + Chr$(193) + Chr$(191) + Chr$(179) + Chr$(180) + Chr$(195) + Chr$(196) + Chr$(192) + Chr$(194) + Chr$(217)
           SymbolBASE = 2
           If frameSYMBOLS.Visible Then ShowSymbols
   Case 3: cboLineSymbols.ListIndex = SymbolBASE
End Select
End Sub



Private Sub cboORIENT_Click()
X = PDFSetup.OPTIONS And (255 - 32)
Y = cboORIENT.ListIndex + 1
Printer.EndDoc
On Local Error GoTo OrientERR
If Y < 3 Then Printer.Orientation = Y
Y = Printer.Orientation
PDFSetup.OPTIONS = (X - 32 * (Y = 2))
Exit Sub

OrientERR:
  cboORIENT.ListIndex = 2
  Resume Next
End Sub

Private Sub cboPRTDest_Click()
Select Case cboPRTDest.ListIndex
   Case 0: PDFSetup.DEST = "SCRN:"
   Case 1: PDFSetup.DEST = "LPT1:"
   Case 2: PDFSetup.DEST = "FILE"
End Select
End Sub


Private Sub cboUNIT_Click()
'If hUNIT <> cboUNIT.ListIndex Then
   Select Case cboUNIT.ListIndex
      Case 0:  xX$ = "in"
      Case 1:  xX$ = "cm"
      Case 2:  xX$ = "mm"
   End Select
   PPLbl(3) = xX$: PPLbl(7) = xX$: PPLbl(10) = xX$
   PPLbl(14) = xX$: PPLbl(17) = xX$
   PPvalue(0) = UnitValue(PDFSetup.FWIDTH)
   PPvalue(1) = UnitValue(PDFSetup.FLENGTH)
   PPvalue(2) = UnitValue(PDFSetup.TMARGIN)
   PPvalue(3) = UnitValue(PDFSetup.BMARGIN)
   PPvalue(4) = UnitValue(PDFSetup.IMARGIN)
   PPvalue(5) = UnitValue(PDFSetup.OMARGIN)
'End If
End Sub


Private Sub ChkOPT_Click(Index As Integer)
X = PDFSetup.OPTIONS And (255 - 2 ^ (Index - 1))
PDFSetup.OPTIONS = (X + (2 ^ (Index - 1)) * ChkOPT(Index).VALUE)
End Sub

Private Sub cmdADD_Click()
SaveSetup
GetSetupName:
   Call frmFHSWMsgD.GetSETUP(xname$)
   If (a = 27) Then Exit Sub
   X = PSUTABLE.FT: FOUND = 0
   While (X > 0) And (Not FOUND)
      FOUND = (xname$ = PSETUPS(X).NAME)
      If Not FOUND Then X = PSUCHN(X).FWD
   Wend
   If FOUND Then
      Call PutMSG("Setup Name Already Exists...")
      If a = 27 Then Exit Sub
      GoTo GetSetupName
   End If
' Add New SETUP
  If PSUTABLE.FF = 0 Then
     ts = PSUTABLE.SIZE + 1
     ReDim Preserve PSETUPS(ts) As PrintSetupREC
     ReDim Preserve PSUCHN(ts) As ChainPTRS
     ReDim Preserve PRINTERS$(ts), PFONT$(ts)
     PSUTABLE.SIZE = ts
     PSUTABLE.FF = ts
  End If
  PSUTABLE.CT = ENV.PSETUP
  Call TableUPDT("ADD", PSUTABLE, PSUCHN())
  If a = 27 Then Exit Sub
  PDFSetup.NAME = xname$
  PRTDesc.TEXT = ""
  chkDefault.VALUE = 0
  PFONT$(PSUTABLE.CT) = PFONT$(ENV.PSETUP)
  ENV.PSETUP = PSUTABLE.CT
  PDFHdr.SETUPS = PDFHdr.SETUPS + 1
  SaveSetup
  BuildPRTSetup
  PRTDesc.SetFocus

End Sub

Private Sub cmdCANCEL_Click()
Dim ChkSetup As PrintSetupREC
ChkSetup = PSETUPS(ENV.PSETUP)
xchg = (ChkSetup.SYMBOLS <> PDFSetup.SYMBOLS)
xchg = symbchg Or (ChkSetup.SET1 <> PDFSetup.SET1)
xchg = symbchg Or (ChkSetup.SET2 <> PDFSetup.SET2)
xchg = symbchg Or (ChkSetup.SET3 <> PDFSetup.SET3)
If xchg Then
   PDFSetup.SYMBOLS = ChkSetup.SYMBOLS
   PDFSetup.SET1 = ChkSetup.SET1
   PDFSetup.SET2 = ChkSetup.SET2
   PDFSetup.SET3 = ChkSetup.SET3
   ShowSetup
   Exit Sub
End If
hfont$ = Printer.Font.NAME
Printer.Font = PFONT$(ENV.PSETUP)
xchg = (Printer.Font.NAME <> hfont$)
xchg = xchg Or (ChkSetup.DESC <> PDFSetup.DESC)
xchg = xchg Or (ChkSetup.FWIDTH <> PDFSetup.FWIDTH)
xchg = xchg Or (ChkSetup.FLENGTH <> PDFSetup.FLENGTH)
xchg = xchg Or (ChkSetup.TMARGIN <> PDFSetup.TMARGIN)
xchg = xchg Or (ChkSetup.BMARGIN <> PDFSetup.BMARGIN)
xchg = xchg Or (ChkSetup.IMARGIN <> PDFSetup.IMARGIN)
xchg = xchg Or (ChkSetup.OMARGIN <> PDFSetup.OMARGIN)
xchg = xchg Or (ChkSetup.OPTIONS <> PDFSetup.OPTIONS)
If xchg Then
   PDFSetup = ChkSetup
   ShowSetup
   Else
   Unload frmFHSWPrinter
End If
End Sub

Private Sub cmdCHGSYMB_Click()
If frameSYMBOLS.Visible Then
   ShowPageParms
   Else
   ShowSymbols
End If
End Sub

Private Sub cmdDEL_Click()
   Call PgmPREP("FHSPRTC")
' Delete This SETUP
    If ENV.PSETUP = PDFHdr.DEFAULT Then
       Call PutMSG("PRTCM004")
       a = 27: OKAY = 0: Exit Sub
    End If
    If PDFHdr.SETUPS = 1 Then
       Call PutMSG("Cannot Delete the Only Setup...")
       Exit Sub
    End If
    LSet SysVAR.MVAR = PSETUPS(ENV.PSETUP).NAME
    Call PutMSG("PRTCM006")
    If CA$ <> SNGLKEY$(1) Then
       OKAY = 0: a = 27
       Exit Sub
    End If
    PSUTABLE.CT = ENV.PSETUP
    Call TableUPDT("DEL", PSUTABLE, PSUCHN())
    If a = 27 Then Exit Sub
    PDFHdr.VER = Chr$(255)
    PDFHdr.SETUPS = PDFHdr.SETUPS - 1
    xname$ = PSETUPS(PDFHdr.DEFAULT).NAME
    Call PDFGetSet(xname$)
    chkDefault.VALUE = 1
    BuildPRTSetup

End Sub

Private Sub cmdHEX_Click()
SymbolREP = SymbolREP + 1
If SymbolREP > 2 Then SymbolREP = 0
For X = 1 To 12
    If (X <> 6) And (Not ((TypeSymbol > 0) And (X = 12))) Then
       Select Case SymbolREP
          Case 0: xX$ = LSFValue(X)
                  LSFValue(X).MaxLength = 3
                  LSFValue(X).Font = LSFValueFont
                  LSFValue(X) = Asc(xX$)
                  cmdHEX.Caption = "HEX"
          Case 1: Z = LSFValue(X)
                  LSFValue(X).MaxLength = 2
                  LSFValue(X) = Dec2HEX(Z)
                  cmdHEX.Caption = "CHAR"
          Case 2: xX$ = LSFValue(X)
                  LSFValue(X).MaxLength = 1
                  LSFValue(X).Font = Printer.Font
                  LSFValue(X) = Chr$(HEX2Dec(xX$))
                  cmdHEX.Caption = "ASCII"
       End Select
    End If
Next X
End Sub

Private Sub cmdLOAD_Click()
Call PDFOpen("INIT"): If Not OKAY Then Exit Sub
BuildPRTSetup
ShowSetup
End Sub

Private Sub cmdPrint_Click()
   SaveSetup
   RptPGM = "PRTC"
   RptCODE = "PRTC"
   Call RptPRTC("INIT")
End Sub

Private Sub cmdRESHOW_Click()
   If HPPValue <> PPvalue(HPPIndex) Then
      X = StdValue(PPvalue(HPPIndex))
      Select Case HPPIndex
        Case 0: PDFSetup.FWIDTH = X
                PPvalue(6) = CharCount(PDFSetup.FWIDTH)
        Case 1: PDFSetup.FLENGTH = X
                PPvalue(7) = LineCount(PDFSetup.FLENGTH)
        Case 2: PDFSetup.TMARGIN = X
                PPvalue(8) = LineCount(PDFSetup.TMARGIN)
        Case 3: PDFSetup.BMARGIN = X
                PPvalue(9) = LineCount(PDFSetup.BMARGIN)
        Case 4: PDFSetup.IMARGIN = X
                PPvalue(10) = CharCount(PDFSetup.IMARGIN)
        Case 5: PDFSetup.OMARGIN = X
                PPvalue(11) = CharCount(PDFSetup.OMARGIN)
      End Select
      HPPValue = PPvalue(HPPIndex)
      cmdRETURN.DEFAULT = yes
      cmdRETURN.SetFocus
   End If
End Sub

Private Sub cmdRETURN_Click()
SaveSetup
Unload frmFHSWPrinter
End Sub

Private Sub cmdSAVE_Click()
Call SaveSetup
HDef = PDFHdr.DEFAULT: PDFHdr.DEFAULT = 1
' Save TABLE Changes Into PDF File
    Call PDFOpen("NEW"): If Not OKAY Then Exit Sub
    'Call PutMSG("PRTCM010")
    PDFHdr.FType = "PDF"
    PDFHdr.VER = Chr$(1): PDFHdr.EOFRBA = 0
    PDFHdr.SETUPS = 0
    Put #1, 1, PDFHdr
    For Z = 0 To PRTSetup.ListCount - 1
        X = PRTSetup.ItemData(Z)
        PDFHdr.SETUPS = PDFHdr.SETUPS + 1
        If X = HDef Then PDFHdr.DEFAULT = PDFHdr.SETUPS
        Put #1, , PSETUPS(X)
        XL$ = Chr$(Len(PRINTERS$(X)))
        Put #1, , XL$
        Put #1, , PRINTERS$(X)
        XL$ = Chr$(Len(PFONT$(X)))
        Put #1, , XL$
        Put #1, , PFONT$(X)
    Next Z
    PDFHdr.EOFRBA = Seek(1) - 1
    Put #1, 1, PDFHdr
    Close
    xname$ = PDFSetup.NAME
    Call PDFOpen("INITQ")
    Call PDFGetSet(xname$)
    BuildPRTSetup

End Sub

Private Sub cmdPREV_Click()
TypeSymbol = TypeSymbol - 1
If TypeSymbol < 0 Then TypeSymbol = 2
ShowSymbols
End Sub

Private Sub cmdNEXT_Click()
TypeSymbol = TypeSymbol + 1
If TypeSymbol > 2 Then TypeSymbol = 0
ShowSymbols
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If (KeyAscii = 27) Then
   If (HLSFVIndex > 0) Then
      If LSFValue(HLSFVIndex) <> HLSFValue Then
         LSFValue(HLSFVIndex) = HLSFValue
         Else
         Call cmdCHGSYMB_Click
      End If
      Else
      Call cmdCHGSYMB_Click
   End If
   KeyAscii = 0
End If
If (KeyAscii = 13) Then
   Call cmdCHGSYMB_Click: KeyAscii = 0
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
FormLoaded(1) = FormLoaded(1) And (255 - PrntForm)
SaveSetting "FHSWIN", "frmPRNT", "Top", frmFHSWPrinter.TOP
SaveSetting "FHSWIN", "frmPRNT", "Left", frmFHSWPrinter.Left
SaveSetting "FHSWIN", "frmPRNT", "UnitMeas", cboUNIT.ListIndex
End Sub

Private Sub Grid1_Click()
Grid1.Visible = False
End Sub
Private Sub cmdCHGDRVR_Click()
'SaveSetup
Printer.EndDoc
CommonDialog.PrinterDefault = True
CommonDialog.Flags = &H40&
CommonDialog.ShowPrinter
PrinterName.TEXT = Printer.DeviceName
Printer.Font = PFONT$(ENV.PSETUP)
Printer.Font.SIZE = PDFSetup.FONTSIZE
Printer.Font.Bold = ((PDFSetup.FONTSTYL And 1) > 0)
Printer.Font.Italic = ((PDFSetup.FONTSTYL And 2) > 0)
On Local Error GoTo OrientERR
Printer.Orientation = 1 - ((PDFSetup.OPTIONS And 32) = 32)
cboORIENT.ListIndex = Printer.Orientation - 1
On Local Error GoTo 0
cboFontSize = Printer.Font.SIZE
Call BuildFonTable
Exit Sub

OrientERR:
   If (ERR = 380) Or (ERR = 483) Then
      cboORIENT.ListIndex = 2
   End If
   Resume Next
      
End Sub



Private Sub cmdSYMBOLS_Click()
If Grid1.Visible = False Then
   Grid1.Font = Printer.Font
   frmFHSWPrinter.Height = 6330
   Grid1.Visible = True
   Grid1.Refresh
   Else
   Grid1.Visible = False
   frmFHSWPrinter.Height = 4680
End If
End Sub


Private Sub Form_Load()
FormLoaded(1) = FormLoaded(1) Or PrntForm
frmFHSWPrinter.Height = 4680
' Set Default Screen Location
frmHEIGHT = frmFHSWPrinter.Height
frmWIDTH = frmFHSWPrinter.WIDTH
t1 = GetSetting("FHSWIN", "frmPRNT", "TOP", 0)
If t1 = 0 Then
   t00 = mnuFHSWMain.TOP: w00 = mnuFHSWMain.WIDTH
   l00 = mnuFHSWMain.Left
   t1 = (t00 + 400)
   l1 = l00 + (w00 - frmWIDTH) / 2: If l1 < 0 Then l1 = 1
   Else
   t1 = 1440: l1 = 1440
End If
frmFHSWPrinter.TOP = GetSetting("FHSWIN", "frmPRNT", "Top", t1)
frmFHSWPrinter.Left = GetSetting("FHSWIN", "frmPRNT", "Left", l1)
hUNIT = GetSetting("FHSWIN", "frmPRNT", "UnitMeas", 0)
If ((FormLoaded(1) And (RPgmForm + UPgmForm)) <> 0) Then
   cmdPrint.Enabled = False
End If
PrinterName.TEXT = Printer.DeviceName
Call BuildFonTable
Grid1.WIDTH = 3590: Grid1.Height = 4100
For X = 0 To 16: For Y = 0 To 16
   If Y = 0 Then
      Grid1.ColWidth(X) = 200
      If X > 0 Then
         Grid1.Row = 0: Grid1.Col = X
         Grid1.TEXT = MID$("0123456789ABCDEF", X, 1)
      End If
      Else
      If X = 0 Then
         Grid1.Row = Y: Grid1.Col = 0
         Grid1.TEXT = MID$("0123456789ABCDEF", Y, 1)
         Else
         Grid1.Row = Y: Grid1.Col = X
         Grid1.TEXT = Chr$((X - 1) * 16 + (Y - 1))
      End If
   End If
Next Y: Next X
Grid1.Visible = False: Grid1.Refresh
Grid1.TOP = cboFONT.TOP + cboFONT.Height
Grid1.Left = cboFONT.Left
LSFValueFont = LSFValue(1).Font
BuildPRTSetup
SymbolREP = 2
cboUNIT.ListIndex = hUNIT
ShowSetup
End Sub

Private Sub LSFValue_GotFocus(Index As Integer)
LSFValue(Index).SelStart = 0
LSFValue(Index).SelLength = Len(LSFValue(Index).TEXT)
HLSFValue = LSFValue(Index): HLSFVIndex = Index
End Sub


Private Sub LSFValue_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii < 32 Then GoTo KeyOK
Select Case SymbolREP
   Case 0: If (KeyAscii < 48) Or (KeyAscii > 57) Then GoTo KeyBAD
   Case 1: If (KeyAscii < 48) Or (KeyAscii > 102) Then GoTo KeyBAD
           KeyAscii = Asc(UCase(Chr(KeyAscii)))
           If (KeyAscii > 57) And (KeyAscii < 65) Then GoTo KeyBAD
           If (KeyAscii > 70) Then GoTo KeyBAD
           xX$ = LSFValue(Index)
           X = LSFValue(Index).SelStart + 1
           Mid$(xX$, X, 1) = Chr$(KeyAscii)
           KeyAscii = 0: LSFValue(Index) = xX$
           If X = 2 Then
              X = Index + 1
              If X = 6 Then X = 7
              If (X = 13) Or ((X = 12) And (TypeSymbol > 0)) Then X = 1
              LSFValue(X).SetFocus
              Else
              LSFValue(Index).SelStart = X
           End If
   Case 2: If KeyAscii < 32 Then GoTo KeyBAD
   Case Else: GoTo KeyOK
End Select
GoTo KeyOK
KeyBAD: KeyAscii = 0: Call ErrBEEP(0)
KeyOK: Exit Sub
End Sub


Private Sub LSFValue_LostFocus(Index As Integer)
If HLSFVIndex > 0 Then
   Select Case SymbolREP
      Case 0: Y = Val(LSFValue(Index))
      Case 1: Y = HEX2Dec(LSFValue(Index))
      Case 2: Y = Asc(LSFValue(Index))
   End Select
   If (0 < Y) And (Y < 256) Then
      Z = Index: If Index > 6 Then Z = Z - 1
      Select Case TypeSymbol
         Case 0: Mid$(PDFSetup.SET1, Z, 1) = Chr$(Y)
         Case 1: Mid$(PDFSetup.SET2, Z, 1) = Chr$(Y)
         Case 2: Mid$(PDFSetup.SET3, Z, 1) = Chr$(Y)
      End Select
   End If
   If LSFValue(HLSFVIndex) <> HLSFValue Then
      SymbolBASE = 3
      cboLineSymbols.ListIndex = 3
   End If
End If
HLSFVIndex = 0
End Sub


Private Sub PPvalue_Change(Index As Integer)
cmdRESHOW.DEFAULT = True
End Sub

Private Sub PPvalue_GotFocus(Index As Integer)
PPvalue(Index).SelStart = 0
PPvalue(Index).SelLength = Len(PPvalue(Index).TEXT)
HPPValue = PPvalue(Index): HPPIndex = Index
End Sub

Private Sub PPvalue_LostFocus(Index As Integer)
   If HPPValue <> PPvalue(Index) Then
   X = StdValue(PPvalue(Index))
   Select Case Index
     Case 0: PDFSetup.FWIDTH = X
             PPvalue(6) = CharCount(PDFSetup.FWIDTH)
     Case 1: PDFSetup.FLENGTH = X
             PPvalue(7) = LineCount(PDFSetup.FLENGTH)
     Case 2: PDFSetup.TMARGIN = X
             PPvalue(8) = LineCount(PDFSetup.TMARGIN)
     Case 3: PDFSetup.BMARGIN = X
             PPvalue(9) = LineCount(PDFSetup.BMARGIN)
     Case 4: PDFSetup.IMARGIN = X
             PPvalue(10) = CharCount(PDFSetup.IMARGIN)
     Case 5: PDFSetup.OMARGIN = X
             PPvalue(11) = CharCount(PDFSetup.OMARGIN)
   End Select
End If
cmdRETURN.DEFAULT = True
'If Index = 0 Then ComputeWidth
End Sub

Private Sub PRTDesc_GotFocus()
PRTDesc.SelStart = 0
PRTDesc.SelLength = Len(RTrim$(PRTDesc.TEXT))
End Sub


Private Sub PRTDesc_LostFocus()
PDFSetup.DESC = PRTDesc.TEXT
End Sub


Private Sub PRTSetup_Click()
   If RTrim$(PRTSetup.TEXT) <> RTrim$(PDFSetup.NAME) Then
      If PRTSetup.TEXT <> "" Then
         SaveSetup
         Else
         PRTSetup.TEXT = PDFSetup.NAME
      End If
      Call PDFGetSet(PRTSetup.TEXT)
      If OKAY Then
         ShowSetup
         Else
         PRTDesc.TEXT = ""
      End If
   End If

End Sub



