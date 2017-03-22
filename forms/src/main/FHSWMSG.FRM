VERSION 4.00
Begin VB.Form frmFHSWMsg 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FHSWMSG -  Message"
   ClientHeight    =   600
   ClientLeft      =   2820
   ClientTop       =   4245
   ClientWidth     =   6255
   FillColor       =   &H00C0C0C0&
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Height          =   1005
   KeyPreview      =   -1  'True
   Left            =   2760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2.5
   ScaleMode       =   4  'Character
   ScaleWidth      =   52.125
   ShowInTaskbar   =   0   'False
   Top             =   3900
   Width           =   6375
   Begin VB.TextBox FHSWMsg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS LineDraw"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "frmFHSWMsg"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
DefInt A-Z
Dim Center


Private Sub FHSWMsg_Change()
a = 27: CA$ = Chr$(27)
FHSWMsg.WIDTH = Len(RTrim$(FHSWMsg.TEXT))
frmFHSWMsg.ScaleMode = 1
frmFHSWMsg.WIDTH = FHSWMsg.WIDTH + 1200
FHSWMsg.Left = 600
frmFHSWMsg.Left = Center - (frmFHSWMsg.WIDTH / 2)
frmFHSWMsg.ScaleMode = 4
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
a = KeyCode: CA$ = Chr$(KeyCode): Call ClearMSG
End Sub

Private Sub Form_Load()
FormLoaded(9) = FormLoaded(9) Or MSGForm
' Set Default Form Location on Screen
t1 = GetSetting("FHSWIN", "frmMSG", "TOP", 0)
If t1 = 0 Then
   t1 = 0.75 * Screen.Height
   c1 = 0.5 * Screen.WIDTH
End If
frmFHSWMsg.TOP = GetSetting("FHSWIN", "frmMSG", "Top", t1)
Center = GetSetting("FHSWIN", "frmMSG", "Center", c1)
CA$ = ""
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X > FHSWMsg.Left Then
   FHSWMsg.Enabled = True
   FHSWMsg.SelStart = X - FHSWMsg.Left
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
xX = X - FHSWMsg.Left
xA = FHSWMsg.SelStart
XL = FHSWMsg.SelLength
If (xX > 0) And (xX < xA) Then
   FHSWMsg.SelStart = xX
   FHSWMsg.SelLength = XL + xA - xX
   Else
   If xX > xA Then
      FHSWMsg.SelLength = xX - xA
   End If
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Len(FHSWMsg.SelText) = 1 Then
   CA$ = FHSWMsg.SelText: a = Asc(CA$)
   Else
   If Len(FHSWMsg.SelText) = 0 Then
      a = 13: CA$ = Chr$(13)
   End If
End If
FHSWMsg.Enabled = False
If Len(FHSWMsg.SelText) < 2 Then Unload frmFHSWMsg
End Sub


Private Sub Form_Unload(Cancel As Integer)
SaveSetting "FHSWIN", "frmMSG", "Top", frmFHSWMsg.TOP
Center = frmFHSWMsg.Left + (frmFHSWMsg.WIDTH / 2)
SaveSetting "FHSWIN", "frmMSG", "Center", Center
FormLoaded(9) = FormLoaded(9) And (255 - MSGForm)
End Sub


