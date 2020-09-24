VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmHTML 
   Caption         =   "Formatted Code"
   ClientHeight    =   5955
   ClientLeft      =   5760
   ClientTop       =   3885
   ClientWidth     =   6735
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHTML.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5955
   ScaleWidth      =   6735
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Command3"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   6720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtHTML 
      Appearance      =   0  'Flat
      Height          =   5835
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   60
      Width           =   6615
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   2520
      Top             =   6720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmHTML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    txtHTML.Width = Me.Width - 240
    txtHTML.Height = Me.Height - 560
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If FullExit = False Then
        Cancel = 1
        Me.Visible = False
    End If
End Sub
