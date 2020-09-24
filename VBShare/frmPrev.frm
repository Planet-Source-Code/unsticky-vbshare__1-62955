VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form frmPrev 
   Caption         =   "Code Preview"
   ClientHeight    =   5805
   ClientLeft      =   3345
   ClientTop       =   2820
   ClientWidth     =   8400
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrev.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5805
   ScaleWidth      =   8400
   Begin SHDocVwCtl.WebBrowser Viewer 
      Height          =   5835
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   8415
      ExtentX         =   14843
      ExtentY         =   10292
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmPrev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    Viewer.Width = Me.Width - 100
    Viewer.Height = Me.Height - 400
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If FullExit = False Then
        Cancel = 1
        Me.Visible = False
    End If
End Sub

Private Sub Viewer_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    Me.Caption = "Code Preview" & IIf(Viewer.Document.Title <> "", " - " & Viewer.Document.Title, "")
End Sub
