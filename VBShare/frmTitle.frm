VERSION 5.00
Begin VB.Form frmTitle 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Page Title"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3975
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTitle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   Begin VBShare.lvButtons_H cmdOK 
      Height          =   315
      Left            =   3240
      TabIndex        =   3
      Top             =   60
      Width           =   675
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "&OK"
      CapAlign        =   2
      BackStyle       =   4
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
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   1320
      Width           =   3855
   End
   Begin VBShare.lvButtons_H cmdCancel 
      Height          =   315
      Left            =   3240
      TabIndex        =   4
      Top             =   420
      Width           =   675
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "&Cancel"
      CapAlign        =   2
      BackStyle       =   4
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
   Begin VB.Label lbl 
      Caption         =   $"frmTitle.frx":030A
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   60
      Picture         =   "frmTitle.frx":03A0
      Top             =   780
      Width           =   3840
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Example:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   1
      Top             =   600
      Width           =   795
   End
End
Attribute VB_Name = "frmTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
    Unload Me
    frmHTML.txtHTML = toHTML(frmCode.txtCode, txtTitle, frmPref.txtBG, frmPref.txtCom, frmPref.txtKey, frmPref.txtNorm)
    frmHTML.Show
    frmHTML.SetFocus
End Sub

Private Sub Form_Activate()
    Me.Left = frmMain.Width * 0.327586206896552
    Me.Top = frmMain.Height * 0.29739336492891
    txtTitle = frmMain.GetFileTitle
End Sub

Private Sub Form_Deactivate()
If Me.Visible = True Then
    Beep
    txtTitle.SetFocus
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If FullExit = False Then
        Cancel = 1
        Me.Visible = False
    End If
End Sub

Private Sub txtTitle_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdOK_Click
End Sub
