VERSION 5.00
Begin VB.Form frmPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Code"
   ClientHeight    =   1395
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3735
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   "Print Quality"
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   1560
      TabIndex        =   7
      Top             =   300
      Width           =   1095
      Begin VB.OptionButton optQual 
         Appearance      =   0  'Flat
         Caption         =   "Low"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   10
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton optQual 
         Appearance      =   0  'Flat
         Caption         =   "Medium"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   9
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton optQual 
         Appearance      =   0  'Flat
         Caption         =   "High"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   60
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Print What"
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   60
      TabIndex        =   4
      Top             =   300
      Width           =   1395
      Begin VB.OptionButton optPrint 
         Appearance      =   0  'Flat
         Caption         =   "VB Code"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   6
         Top             =   240
         Width           =   1275
      End
      Begin VB.OptionButton optPrint 
         Appearance      =   0  'Flat
         Caption         =   "HTML Code"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   5
         Top             =   540
         Width           =   1275
      End
   End
   Begin VBShare.lvButtons_H cmdOK 
      Height          =   315
      Left            =   2820
      TabIndex        =   0
      Top             =   60
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      Caption         =   "OK"
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
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VBShare.lvButtons_H cmdCancel 
      Height          =   315
      Left            =   2820
      TabIndex        =   1
      Top             =   480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      Caption         =   "Cancel"
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
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Label lblPrinter 
      Caption         =   "Print Name"
      Height          =   195
      Left            =   660
      TabIndex        =   3
      Top             =   60
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "Printer:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   495
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If optQual(0).Value = True Then
        Printer.PrintQuality = 0
    ElseIf optQual(1).Value = True Then
        Printer.PrintQuality = 1
    ElseIf optQual(2).Value = True Then
        Printer.PrintQuality = 2
    End If
    If optPrint(0).Value = True Then
        Printer.Print frmCode.txtCode
    ElseIf optPrint(1).Value = True Then
        Printer.Print frmHTML.txtHTML
    End If
    Printer.EndDoc
    Unload Me
End Sub

Private Sub Form_Activate()
    lblPrinter.Caption = Printer.DeviceName
    If frmCode.txtCode <> "" And frmCode.Visible = True Then
        optPrint(0).Enabled = True
    Else
        optPrint(0).Enabled = False
    End If
    If frmHTML.txtHTML <> "" And frmHTML.Visible = True Then
        optPrint(1).Enabled = True
    Else
        optPrint(1).Enabled = False
    End If
    optPrint(0).Value = optPrint(0).Enabled
    optPrint(1).Value = optPrint(0).Enabled = False
    Me.Left = frmMain.Width * 0.327586206896552
    Me.Top = frmMain.Height * 0.29739336492891
End Sub

Private Sub Form_Deactivate()
    If Me.Visible = True Then
        Beep
        Me.SetFocus
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If FullExit = False Then
        Cancel = 1
        Me.Visible = False
    End If
End Sub

Public Sub LoadQuality(Quality As String)

End Sub
