VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3000
   ClientLeft      =   2310
   ClientTop       =   1590
   ClientWidth     =   6000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2070.653
   ScaleMode       =   0  'User
   ScaleWidth      =   5634.31
   Begin VBShare.lvButtons_H cmdOK 
      Height          =   315
      Left            =   4920
      TabIndex        =   1
      Top             =   2580
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
      Caption         =   "&OK"
      CapAlign        =   2
      BackStyle       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Label lblHTTP 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3060
      MouseIcon       =   "frmAbout.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1860
      Width           =   1935
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Left = frmMain.Width * 0.327586206896552
    Me.Top = frmMain.Height * 0.29739336492891
    Me.Picture = frmSplash.imgSplash.Picture
    Me.Width = frmSplash.imgSplash.Width
    Me.Height = frmSplash.imgSplash.Height
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If FullExit = False Then
        Cancel = 1
        Me.Visible = False
    End If
End Sub

Private Sub lblHTTP_Click()
OpenSite "http://www.unsticky.net"
End Sub
