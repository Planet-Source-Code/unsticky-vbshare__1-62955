VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   3030
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   6030
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
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSerial 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1860
      TabIndex        =   0
      Top             =   2700
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label lblStat 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1620
      TabIndex        =   2
      Top             =   2700
      Width           =   255
   End
   Begin VB.Image img2 
      Height          =   255
      Left            =   5100
      Picture         =   "frmSplash.frx":030A
      Top             =   600
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Image img1 
      Height          =   255
      Left            =   5100
      Picture         =   "frmSplash.frx":0C54
      Top             =   300
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Image cmdSubmit 
      Height          =   255
      Left            =   4680
      Top             =   2700
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Program Serial"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1980
      TabIndex        =   1
      Top             =   2520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image imgSplash 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   3030
      Left            =   0
      Picture         =   "frmSplash.frx":159E
      Top             =   0
      Width           =   6030
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSubmit_Click()
    lblStat.Caption = IIf(ValidSerial(txtSerial) <> False, Chr$(252), Chr$(251))
End Sub

Private Sub cmdSubmit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdSubmit.Picture = img2.Picture
End Sub

Private Sub cmdSubmit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdSubmit.Picture = img1.Picture
End Sub

Private Sub Form_Click()
    If txtSerial.Visible = False Then Unload Me
    If lblStat = Chr$(252) And ValidSerial(txtSerial) = True Then Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If txtSerial.Visible = False Then Unload Me
    If lblStat = Chr$(252) And ValidSerial(txtSerial) = True Then Unload Me
End Sub

Private Sub Form_Load()
    Dim Ser As String
    Ser$ = GetSetting$("VBShare", "General", "Serial", "")
    If ValidSerial(Ser$) = False Then
        cmdSubmit.Picture = img1.Picture
        lbl.Visible = True
        cmdSubmit.Visible = True
        txtSerial.Visible = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If lblStat = Chr$(252) And ValidSerial(txtSerial) = True Then SaveSetting "VBShare", "General", "Serial", txtSerial
    txtSerial = Empty
    frmMain.Show
End Sub

Private Sub imgSplash_Click()
    If txtSerial.Visible = False Then Unload Me
    If lblStat = Chr$(252) And ValidSerial(txtSerial) = True Then Unload Me
End Sub
